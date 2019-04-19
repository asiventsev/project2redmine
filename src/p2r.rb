#=====================================================================
# Console Script for MS Project to Redmine synchronization
# (c) Siventsev Aleksei 2019
#=====================================================================
VER = '0.1 17/04/19'
HDR = "Console Script for MS Project to Redmine synchronization v#{VER}"

require 'yaml'
require 'win32ole'
require 'net/http'
require 'p2r_lib.rb'

puts '', HDR, ('=' * HDR.scan(/./mu).size), ''

#---------------------------------------------------------------------
# process command line arguments
#---------------------------------------------------------------------
# answer help request and exit
chk !(ARGV & %w(h H -h -H /h /H ? -? /? help -help --help)).empty?, HELP
# check execution request
dry_run=!(ARGV & %w(e E -e -E /e /E exec -exec --exec execute -execute --execute)).empty?

#---------------------------------------------------------------------
# connect to .msp
#---------------------------------------------------------------------
msg = 'Please open your MS Project file and leave it active with no dialogs open'
begin
  pserver=WIN32OLE.connect 'MSProject.Application'
rescue
  chk true, msg
end
msp=pserver.ActiveProject
chk !msp,msg

#---------------------------------------------------------------------
# find and process settings task
#---------------------------------------------------------------------
settings_task = nil
puts msp.Tasks.Count
(1..msp.Tasks.Count).each do |i|
  raw=msp.Tasks(i)
  if raw && raw.Name == 'Redmine Synchronization'
    settings_task = raw
    break
  end
end
chk !settings_task, 'ERROR: task with name \'Redmine Sysncronization\' was not found in the project.'

begin
  pak = YAML.load settings_task.Notes.to_s.gsub("\r", "\n")
rescue
  chk true, 'ERROR: could not extract settiings from Notes in \'Redmine Sysncronization\' task (YAML format expected)'
end

rpr_id = pak.delete 'redmine_project_id'
missed_pars = %w(redmine_host redmine_api_key redmine_project_uuid task_redmine_id_field resource_email_field) - pak.keys

chk !missed_pars.empty?, "ERROR: following settings not found in 'Redmine Sysncronization' task: #{missed_pars.sort.join ', '}"

puts "Redmine project #{rpr_id}"

def rm_request pak, path, data=nil
  host, port = pak['redmine_host'].split ':'
  reply = Net::HTTP.start(pak['redmine_host'], pak['redmine_port']) do |http|
    if data
      http.request_post(path,data.to_json,'Content-Type'=>'application/json')
    else
      http.request_get(path)
    end
  end
  return reply
end

#---------------------------------------------------------------------
# check Redmine project availability
#---------------------------------------------------------------------
puts "/projects/#{pak['redmine_project_uuid']}.json"
puts rm_request(pak,"/projects/#{pak['redmine_project_uuid']}.json").inspect

