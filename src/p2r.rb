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

def build_mst_url(pak, rmt_id); pak['task_redmine_url_field'] ? "http://#{pak['redmine_host']}:#{pak['redmine_port']}/issues/#{rmt_id}" : nil; end
def get_mst_url(pak, mst); pak['task_redmine_url_field'] ? eval("mst.#{pak['task_redmine_url_field']}") : nil; end
def set_mst_url(pak, mst, rmt_id)
  url=build_mst_url pak, rmt_id
  eval("mst.#{pak['task_redmine_url_field']} = '#{url}'") if url
  return url
end
def get_mst_redmine_id(pak, mst); eval("mst.#{pak['task_redmine_id_field']}"); end
def set_mst_redmine_id(pak, mst, rmt_id); eval("mst.#{pak['task_redmine_id_field']} = '#{rmt_id}'"); end

def process_issues pak, msp, rmp_id, force_task_creation = false
  rmt_id_field = "mst.#{pak['task_redmine_id_field']}"
  (1..msp.Tasks.Count).each do |i|
    next unless mst = msp.Tasks(i)
    rmt_id = eval(rmt_id_field)
    next if rmt_id.empty? # task not marked for sync
    rmt_id = rmt_id.to_i rescue 0
    if rmt_id == 0 || force_task_creation
      # create new task
      unless DRY_RUN
        rmt = {
            project_id: rmp_id, subject: mst.Name, description: "-----\nAutocreated by P2R from MSP task #{mst.ID} in MSP project #{msp.Name}\n-----\n",
            start_date: mst.Start.strftime('%Y-%m-%d'), due_date: mst.Finish.strftime('%Y-%m-%d')
        }
        rmt = rm_create pak, '/issues.json', 'issue', rmt,
                        "ERROR: could not create Redmine task from #{mst.ID} #{mst.Name} for some reasons"
        # write new task number to MSP
        set_mst_redmine_id pak, mst, rmt['id']
        set_mst_url pak, mst, rmt['id']
      else
        # keep task to be created
        puts "Will create task #{mst.ID} #{mst.Name}"
      end
    else
      # update existing task
      #   check task availability
      rmt = rm_get pak, "/issues/#{rmt_id}.json", 'issue', "ERROR: could not find Redmine task ##{rmt_id} for #{mst.ID} #{mst.Name}"
      #   check for changes
      #     subject - Name, start_date - Start, due_date - Finish
      changes={}
      changes['subject'] = mst.Name if rmt['subject'] != mst.Name.encode('UTF-8')
      d = mst.Start.strftime('%Y-%m-%d')
      changes['start_date'] = d if rmt['start_date'] != d
      d = mst.Finish.strftime('%Y-%m-%d')
      changes['due_date'] = d if rmt['due_date'] != d
      if changes.empty?
        puts "No changes for Task Redmine ##{rmt_id} from MSP #{mst.ID} #{mst.Name}"
      else
        changelist = changes.keys.join(', ')
        changes['notes'] = "Autoupdated by P2P at #{Time.now.strftime '%Y-%m-%d %H:%M'} (#{changelist})"
        if DRY_RUN
          puts "Will update task Redmine ##{rmt_id} from MSP #{mst.ID} #{mst.Name} (#{changelist})"
        else
          rm_update pak, "/issues/#{rmt['id']}.json", 'issue', {issue: changes},
                    "ERROR: could not update Redmine task ##{rmt['id']} from #{mst.ID} #{mst.Name} for some reasons"
          puts "Updated task Redmine ##{rmt_id} from MSP #{mst.ID} #{mst.Name} (#{changelist})"
        end
      end
      set_mst_url pak, mst, rmt['id']
    end
  end
end

if rmp_id
  #=====================================================================
  # existing Redmine project update
  #---------------------------------------------------------------------

  puts 'Existing Redmine project update'
  process_issues pak, msp, rmp_id, false

else
  #=====================================================================
  # new Redmine project creation
  #---------------------------------------------------------------------
  if DRY_RUN
    # project creation requested - exit on dry run
    chk true, "Will create new Redmine project #{pak['redmine_project_uuid']} from MSP project #{msp.Title}"
  end

  #---------------------------------------------------------------------
  # new Redmine project create
  #---------------------------------------------------------------------
  rmp = {name: msp.Title, identifier: pak['redmine_project_uuid'], is_public: false}
  rmp = rm_create pak, '/projects.json', 'project', rmp,
      'ERROR: could not create Redmine project for some reasons'

  # add rm project id to msp settings
  pak['redmine_project_id'] = rmp['id']
  settings_task.Notes = YAML.dump pak
  puts "Created new Redmine project #{pak['redmine_project_uuid']} ##{rmp['id']} from MSP project #{msp.Title}"

  #---------------------------------------------------------------------
  # add tasks to Redmine project
  #---------------------------------------------------------------------
  process_issues pak, msp, rmp['id'], true

end

puts "\n\n"


