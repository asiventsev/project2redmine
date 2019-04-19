#=====================================================================
# Console Script for MS Project to Redmine synchronization
# (c) Siventsev Aleksei 2019
#=====================================================================
VER = '0.1 17/04/19'
HDR = "Console Script for MS Project to Redmine synchronization v#{VER}"

require 'yaml'
require 'win32ole'
require 'net/http'
require 'json'
require './p2r_lib.rb'

puts '', HDR, ('=' * HDR.scan(/./mu).size), ''

#---------------------------------------------------------------------
# process command line arguments
#---------------------------------------------------------------------
# answer help request and exit
chk !(ARGV & %w(h H -h -H /h /H ? -? /? help -help --help)).empty?, HELP
# check execution request
DRY_RUN = (ARGV & %w(e E -e -E /e /E exec -exec --exec execute -execute --execute)).empty?
puts "DRY RUN\n\n" if DRY_RUN

#---------------------------------------------------------------------
# connect to .msp
#---------------------------------------------------------------------
msg = 'Please open your MS Project file and leave it active with no dialogs open'
begin
  pserver = WIN32OLE.connect 'MSProject.Application'
rescue
  chk true, msg
end
msp = pserver.ActiveProject
chk !msp,msg

#---------------------------------------------------------------------
# find and process settings task
#---------------------------------------------------------------------
settings_task = nil
(1..msp.Tasks.Count).each do |i|
  raw = msp.Tasks(i)
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

rmp_id = pak.delete 'redmine_project_id'
missed_pars = %w(redmine_host redmine_api_key redmine_project_uuid task_redmine_id_field resource_email_field) - pak.keys

chk !missed_pars.empty?, "ERROR: following settings not found in 'Redmine Sysncronization' task: #{missed_pars.sort.join ', '}"

def rm_request pak, path, data=nil
  host, port = pak['redmine_host'].split ':'
  reply = Net::HTTP.start(pak['redmine_host'], pak['redmine_port']) do |http|
    if data
      http.request_post(path,data.to_json, 'Content-Type'=>'application/json', 'X-Redmine-API-Key' => pak['redmine_api_key'])
    else
      http.request_get(path, 'X-Redmine-API-Key' => pak['redmine_api_key'])
    end
  end
  return reply
end

#---------------------------------------------------------------------
# check Redmine project availability
#---------------------------------------------------------------------
# 401 ERROR: not authorized bad key
# 404 not found
#   if rpr_id then ERROR: suppose project has been published already
#   else project is to be published
# 403 forbidden
# 200 ок
#   if rpr_id then
#     if prp_id == project id then OK to proceed
#     else ERROR: different ids in project and redmine
#   else ERROR: suppose project is to be published but found it is already published
#
project_path="/projects/#{pak['redmine_project_uuid']}.json"
re = rm_request(pak, project_path)

case re.code
  when '401'
    chk true, 'ERROR: not authorized by Redmine (maybe bad api key?)'
  when '404'
    if rmp_id # else proceed
      chk true, "ERROR: suppose project '#{pak['redmine_project_uuid']}' has been published already (because redmine_project_id is provided) but have not found it"
    end
  when '403'
    chk true, "ERROR: access to project '#{pak['redmine_project_uuid']}' in Redmine is forbidden, ask Redmine admin"
  when '200'
    begin
      rmp = JSON.parse(re.body)
    rescue
      chk true, "ERROR: wrong reply format to '/projects/#{pak['redmine_project_uuid']}.json' (JSON expected)"
    end
    rmp = rmp['project']
    chk !rmp, "ERROR: wrong reply format to '/projects/#{pak['redmine_project_uuid']}.json' ('project' key not found)"
    if rmp_id
      unless rmp_id == rmp['id'] # else proceed
        chk true, "ERROR: Redmine project id does not comply with redmine_project_id provided in settings"
      end
    else
      chk true, "ERROR: suppose have to create new project '#{pak['redmine_project_uuid']}' (because redmine_project_id is not provided) but found the project has been published already"
    end
  else
    chk true, "ERROR: #{re.code} #{re.message}"
end

if rmp_id
  #=====================================================================
  # existing Redmine project update
  #---------------------------------------------------------------------

  puts 'existing Redmine project update'
  rmt_id_field = "mst.#{pak['task_redmine_id_field']}"
  (1..msp.Tasks.Count).each do |i|
    next unless mst = msp.Tasks(i)
    rmt_id = eval(rmt_id_field)
    next if rmt_id.empty? # task not marked for sync
    rmt_id = rmt_id.to_i rescue 0
    if rmt_id == 0
      # create new task
      unless DRY_RUN
        rmt={
            project_id: rmp_id, subject: mst.Name, description: "-----\nSynchronized with task #{mst.ID} from MSP project #{msp.Name}\n-----\n",
            start_date: mst.Start.strftime('%Y-%m-%d'), due_date: mst.Finish.strftime('%Y-%m-%d')
        }
        puts rmt
        re = rm_request pak, '/issues.json', issue: rmt
        puts '- - - -',re.code,re.msg,re.body.inspect
      else
        # keep task to be updated
      end
    else
      # update existing task
    end
  end

else
  #=====================================================================
  # new Redmine project creation
  #---------------------------------------------------------------------
  if DRY_RUN
    # project creation requested - exit on dry run
    chk true, 'DRY RUN SUCCESS: new project to be created in Redmine'
  end

  #---------------------------------------------------------------------
  # new Redmine project create
  #---------------------------------------------------------------------
  rmp = {name: msp.Title, identifier: pak['redmine_project_uuid'], is_public: false}
  re = rm_request pak, '/projects.json',project: rmp
  msg='ERROR: could not create Redmine project for some reasons'
  chk (re.code!='201'), msg
  rmp = JSON.parse(re.body)['project'] rescue nil
  chk !rmp, msg

  # add rm project id to msp settings
  pak['redmine_project_id'] = rmp['id']
  settings_task.Notes = YAML.dump pak
  puts 'new Redmine project created'

  #---------------------------------------------------------------------
  # add tasks to Redmine project
  #---------------------------------------------------------------------


end



