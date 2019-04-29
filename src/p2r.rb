#=====================================================================
# Console Script for MS Project to Redmine synchronization
# (c) A. Siventsev 2019
#=====================================================================
VER = '0.2 26/04/19'
HDR = "Console Script for MS Project to Redmine synchronization v#{VER} (c) A. Siventsev 2019"

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
$msp = pserver.ActiveProject
chk !$msp,msg
$msp_name = $msp.Name.clone.encode 'UTF-8'

#---------------------------------------------------------------------
# find and process settings task
#---------------------------------------------------------------------
settings_task = nil
(1..$msp.Tasks.Count).each do |i|
  raw = $msp.Tasks(i)
  if raw && raw.Name == 'Redmine Synchronization'
    settings_task = raw
    break
  end
end
chk !settings_task, 'ERROR: task with name \'Redmine Sysncronization\' was not found in the project.'

begin
  $settings = YAML.load settings_task.Notes.to_s.gsub("\r", "\n")
rescue
  chk true, 'ERROR: could not extract settiings from Notes in \'Redmine Sysncronization\' task (YAML format expected)'
end

rmp_id = $settings.delete 'redmine_project_id'
missed_pars = %w(redmine_host redmine_api_key redmine_project_uuid task_redmine_id_field resource_redmine_id_field resource_default_redmine_role_id) - $settings.keys

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

$uuid = $settings['redmine_project_uuid']
project_path="/projects/#{$uuid}.json"
re = rm_request(project_path)

case re.code
  when '401'
    chk true, 'ERROR: not authorized by Redmine (maybe bad api key?)'
  when '404'
    if rmp_id # else proceed
      chk true, "ERROR: suppose project '#{$uuid}' has been published already (because redmine_project_id is provided) but have not found it"
    end
  when '403'
    chk true, "ERROR: access to project '#{$uuid}' in Redmine is forbidden, ask Redmine admin"
  when '200'
    begin
      rmp = JSON.parse(re.body)
    rescue
      chk true, "ERROR: wrong reply format to '/projects/#{$uuid}.json' (JSON expected)"
    end
    rmp = rmp['project']
    chk !rmp, "ERROR: wrong reply format to '/projects/#{$uuid}.json' ('project' key not found)"
    if rmp_id
      unless rmp_id == rmp['id'] # else proceed
        chk true, "ERROR: Redmine project id does not comply with redmine_project_id provided in settings"
      end
    else
      chk true, "ERROR: suppose have to create new project '#{$uuid}' (because redmine_project_id is not provided) but found the project has been published already"
    end
  else
    chk true, "ERROR: #{re.code} #{re.message}"
end

#---------------------------------------------------------------------
# check default tracker and role
#---------------------------------------------------------------------

if ($dflt_tracker_id = $settings['task_default_redmine_tracker_id'])
  chk !$dflt_tracker_id.is_a?(Integer), "ERROR: parameter task_default_redmine_tracker_id must be integer"
  trackers = rm_get '/trackers.json', 'trackers', 'ERROR: could not get Redmine tracker list'
  found = false
  trackers.each do |t|
    if t['id'] = $dflt_tracker_id
      found=true
      break
    end
  end
  chk !found, "ERROR: tracker not found for parameter task_default_redmine_tracker_id = #{$dflt_tracker_id}"
end

$dflt_role_id = $settings['resource_default_redmine_role_id']
chk !$dflt_tracker_id.is_a?(Integer), "ERROR: parameter task_default_redmine_tracker_id must be integer"
rm_get "/roles/#{$dflt_role_id}.json", 'role', "ERROR: could not get default team role resource_default_redmine_role_id=#{$dflt_role_id}"

#---------------------------------------------------------------------
# load project memberships
#---------------------------------------------------------------------

$team = {}
if rmp_id
  # for existing project only
  offset = 0
  loop do
    re = rm_request "/projects/#{$uuid}/memberships.json?offset=#{offset}"
    chk (re.code != '200'), 'ERROR: could not get team list of Redmine project'
    re = JSON.parse(re.body) rescue nil
    chk re.nil?, 'ERROR: could not parse reply of team list request'
    break if re['memberships'].empty? or re['limit'].nil?
    re['memberships'].each do |m|
      $team[m['user']['id']] = m
    end
    offset += re['limit']
  end
else
  # TODO - get current user and add to team in advance
  puts rm_get '/users/current.json', 'user', 'ERROR: could not get current user'
end
abort
#---------------------------------------------------------------------
# some utils for msp custom fields editing
#---------------------------------------------------------------------

def build_mst_url(rmt_id); $settings['task_redmine_url_field'] ? "http://#{$settings['redmine_host']}:#{$settings['redmine_port']}/issues/#{rmt_id}" : nil; end
def get_mst_url(mst); $settings['task_redmine_url_field'] ? eval("mst.#{$settings['task_redmine_url_field']}") : nil; end
def set_mst_url(mst, rmt_id)
  url=build_mst_url rmt_id
  eval("mst.#{$settings['task_redmine_url_field']} = '#{url}'") if url
  return url
end
def get_mst_redmine_id(mst); eval("mst.#{$settings['task_redmine_id_field']}"); end
def set_mst_redmine_id(mst, rmt_id); eval("mst.#{$settings['task_redmine_id_field']} = '#{rmt_id}'"); end

#---------------------------------------------------------------------
# task (issue) processing util
#---------------------------------------------------------------------

$rmts={} # issues processed
$rmus=[] # memberships processed

def process_issue rmp_id, mst, force_new_task = false

  mst_name = mst.Name.clone.encode 'UTF-8'
  rmt_id = eval "mst.#{$settings['task_redmine_id_field']}"
  return nil unless rmt_id =~ /^\s*\d+\s*$/ # task not marked for sync
  rmt_id = rmt_id.to_i

  if (rmt = $rmts[rmt_id])
    return rmt # already processed
  end

  # process task parent:
  mst_papa = mst.OutlineParent
  rmt_papa = nil
  unless mst_papa.UniqueID == 0 # suppose project summary task has UniqueID = 0
    rmt_papa_id = eval "mst_papa.#{$settings['task_redmine_id_field']}"
    if rmt_papa_id =~ /^\s*\d+\s*$/
      rmt_papa_id = rmt_papa_id.to_i
      rmt_papa = $rmts[rmt_papa_id]
      unless rmt_papa
        rmt_papa = process_issue rmp_id, mst_papa, force_new_task
      end
    else
      rmt_papa_id = 0
      eval "mst_papa.#{$settings['task_redmine_id_field']} = '0'"
      rmt_papa = process_issue rmp_id, mst_papa
    end

  end

  # check task resource appointment
  #   we expect not more than one synchronizable appointment
  rmu_id_field = "msr.#{$settings['resource_redmine_id_field']}"
  rmu_id_ok = nil
  msr_ok = nil
  (1..mst.Resources.Count).each do |j|
    next unless msr = mst.Resources(j)
    rmu_id = eval(rmu_id_field)
    next unless rmu_id =~ /^\s*\d+\s*$/ # resource not marked for sync
    chk rmu_id_ok, "ERROR: more than one sync resource for MSP task #{mst.ID} '#{mst_name}'"
    rmu_id = rmu_id.to_i

    if $rmus.include? rmu_id
      # resource already processed
      rmu_id_ok = rmu_id
      msr_ok = msr
    else
      member = $team[rmu_id]
      unless member
        # Redmine user is not team member - create new membership
        # check user availability
        re = rm_request "/users/#{rmu_id}.json"
        chk (re.code != '200'), "ERROR: Redmine user #{rmu_id} not found for resource in MSP task #{mst.ID} '#{mst_name}'"
        re = JSON.parse(re.body) rescue nil
        chk re.nil?, "ERROR: could not parse reply: Redmine user #{rmu_id} not found for resource in MSP task #{mst.ID} '#{mst_name}'"
        # create membership
        data = {user_id: rmu_id, role_ids: [$dflt_role_id]}
        puts data
        member = rm_create "/projects/#{$uuid}/memberships.json", 'membership', data,
                         "ERROR: could not create Redmine project membership for user #{rmu_id}"
        puts "New membership created: #{rmu_id}"
        $team[rmu_id] = member
      end
      rmu_id_ok = rmu_id
      msr_ok = msr
    end
  end
  if rmu_id_ok
    rmu_name = $team[rmu_id_ok]['user']['name']
    msr_name = msr_ok.Name.clone.encode 'UTF-8'
    unless rmu_name == msr_name
      puts "WARNING: membership ID=#{rmu_id_ok} name '#{rmu_name}' does not correspond to MSP resource name '#{msr_name}' (task ##{rmt_id} for #{mst.ID} '#{mst_name}')"
    end
  end

  if rmt_id == 0 || force_new_task

    # create new task
    unless DRY_RUN
      rmt = {
          project_id: rmp_id, subject: mst_name, description: "-----\nAutocreated by P2R from MSP task #{mst.ID} in MSP project #{$msp_name}\n-----\n",
          start_date: mst.Start.strftime('%Y-%m-%d'), due_date: mst.Finish.strftime('%Y-%m-%d'),
          assigned_to_id: rmu_id_ok, tracker_id: $dflt_tracker_id,
          parent_issue_id: (rmt_papa ? rmt_papa['id'] : '')
      }
      rmt = rm_create '/issues.json', 'issue', rmt,
                      "ERROR: could not create Redmine task from #{mst.ID} '#{mst_name}' for some reasons"
      # write new task number to MSP
      set_mst_redmine_id mst, rmt['id']
      set_mst_url mst, rmt['id']
      puts "Created task Redmine ##{rmt['id']} from MSP #{mst.ID} '#{mst_name}'"

      $rmts[rmt['id']] = rmt
      return rmt

    else
      # keep task to be created
      puts "Will create task #{mst.ID} '#{mst_name}'"

      return nil

    end

  else

    # update existing task
    #   check task availability
    rmt = rm_get "/issues/#{rmt_id}.json", 'issue', "ERROR: could not find Redmine task ##{rmt_id} for #{mst.ID} '#{mst_name}'"

    #   check for changes
    #     subject - Name, start_date - Start, due_date - Finish
    changes={}
    changes['assigned_to_id'] = (rmu_id_ok || '') if rmu_id_ok != (rmt['assigned_to'] ? rmt['assigned_to']['id'] : nil)
    changes['subject'] = mst_name if rmt['subject'] != mst_name
    d = mst.Start.strftime('%Y-%m-%d')
    changes['start_date'] = d if rmt['start_date'] != d
    d = mst.Finish.strftime('%Y-%m-%d')
    changes['due_date'] = d if rmt['due_date'] != d
    rmt_papa_id_old = (rmt['parent'] ? rmt['parent']['id'] : '')
    rmt_papa_id_new = (rmt_papa ? rmt_papa['id'] : '')
    changes['parent_issue_id'] = rmt_papa_id_new if rmt_papa_id_new != rmt_papa_id_old
    # apply changes
    if changes.empty?
      puts "No changes for Task Redmine ##{rmt_id} from MSP #{mst.ID} '#{mst_name}'"
    else
      # apply changes
      changelist = changes.keys.join(', ')
      changes['notes'] = "Autoupdated by P2P at #{Time.now.strftime '%Y-%m-%d %H:%M'} (#{changelist})"
      if DRY_RUN
        puts "Will update task Redmine ##{rmt_id} from MSP #{mst.ID} '#{mst_name}' (#{changelist})"
      else
        rm_update "/issues/#{rmt['id']}.json",  {issue: changes},
                  "ERROR: could not update Redmine task ##{rmt['id']} from #{mst.ID} '#{mst_name}' for some reasons"
        rmt = rm_get "/issues/#{rmt_id}.json", 'issue', "ERROR: could not find Redmine task ##{rmt_id} for #{mst.ID} '#{mst_name}'"
        puts "Updated task Redmine ##{rmt_id} from MSP #{mst.ID} '#{mst_name}' (#{changelist})"
      end
    end
    set_mst_url mst, rmt['id']

    $rmts[rmt['id']] = rmt
    return rmt

  end

end

#---------------------------------------------------------------------
# iterate over task list
#---------------------------------------------------------------------

def process_issues rmp_id, force_new_task = false

  (1..$msp.Tasks.Count).each do |i|

    # check msp task
    next unless mst = $msp.Tasks(i)

    process_issue rmp_id, mst, force_new_task

  end
end

#=====================================================================
# main work cycle
#=====================================================================

if rmp_id
  #=====================================================================
  # existing Redmine project update
  #---------------------------------------------------------------------

  puts 'Existing Redmine project update'
  process_issues rmp_id, false

else
  #=====================================================================
  # new Redmine project creation
  #---------------------------------------------------------------------
  if DRY_RUN
    # project creation requested - exit on dry run
    chk true, "Will create new Redmine project #{$uuid} from MSP project #{$msp_name}"
  end

  #---------------------------------------------------------------------
  # new Redmine project create
  #---------------------------------------------------------------------
  rmp = {name: $msp_name, identifier: $uuid, is_public: false}
  rmp = rm_create '/projects.json', 'project', rmp,
      'ERROR: could not create Redmine project for some reasons'

  # add rm project id to msp settings
  $settings['redmine_project_id'] = rmp['id']
  settings_task.Notes = YAML.dump $settings
  puts "Created new Redmine project #{$uuid} ##{rmp['id']} from MSP project #{$msp_name}"

  #---------------------------------------------------------------------
  # add tasks to Redmine project
  #---------------------------------------------------------------------
  process_issues rmp['id'], true

end

puts "\n\n"


