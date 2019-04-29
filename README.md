# P2R Console script for MS Project to Redmine synchronization

This script is intended to publish any exsisting MSP project to Redmine and get progress data back from Redmine.

Redmine is simple, free and convenient for project team to keep tasks progress and states and to make reports.

P2R provides task and appointment updates to Redmine and puts actual start and finish dates and spent hours back to MSP.
Start and finish are being obtained from task state changes.

P2R allows to choose tasks and team members (resources) to publish to Redmine. 

## MSP project preparation

To get msp project ready for publishing to Redmine or updating already published:
1. Choose Task custom text field for Redmine task ID, for example Text1.
2. Put 0 in this field for tasks to be published to Redmine.
3. Choose Resource custom text field for Redmine team member ID, for example, Text1. 
(ID will be used to identify Redmine project team members which correspond to MSP resources.)
(Please do not use localized names of custom fields, use english names like TextN.) 
4. Set Redmine user IDs for resources to be used in Redmine as team members.
(Memberships will be created authomatically. Ask Redmine admin for user IDs.)
5. Create special task in MSP project named 'Redmine Synchronization' and put Redmine synchronization 
settings to its Notes in YAML format according to the example:

<pre><code>redmine_host: 192.168.10.23
redmine_port: 3500
redmine_api_key: d1e693e7670b1fc79378441d485b9608e4c2dc4c
redmine_project_uuid: uute
task_redmine_id_field: Text1
task_redmine_url_field: Text2
task_default_redmine_tracker_id: 4
resource_redmine_id_field: Text1
resource_default_redmine_role_id: 4
</code></pre>

That's all. The meaning of the parameters is clear from their names. 

Parameter `redmine_project_uuid` must correspond Redmine rules which you could find on in settings pages of any Redmine project.
If an 'Unprocessable Entry' error occurs on project creation it is probably due to incorrect project uuid.

Parameter `task_default_redmine_tracker_id` is not required. If you omit it then Redmine's default tracker will be set.

Parameter `task_redmine_url_field` is not required. If you define it then URL to Remine issue will be put to this field for each sync task.

If you want to connect your MSP project to existing Redmine project instead of creating new one, please do the following:

1. Add parameter redmine_project_id to settings list. 
(It was not obvious how to get it from Redmine. Easy way is to get project in JSON format, for example, from URL `http://192.168.10.23:3500/projects/uute.json` for project described above.)
2. Put real Redmine task IDs instead of 0 for tasks to be synchronized.

## Some limitations

Redmine issue (task) can be appointed to one team member only so P2R expects no more than one synchronizable resourse per MSP project task. If more, an error will be reported and script stopped.

## Synchronization

1. Open your MSP project. (P2R authomatically connects to current active project via OLE.)
2. Run script without command line parameters. It will perform dry run and report all changes to be made. (Not necessary.)
3. Run script with -e parameter. It will perform real execution and make necessary changes to Redmine as well as MSP. 

Working directory does not matter. You can run script itself (if you have ruby installed) or compiled .EXE file from downloads folder 
(it does not require ruby installation). Compiled exe is just SFX archive with the same script plus necessary ruby binaries and libs.

After creating new Redmine project P2R authomatically adds `redmine_project_id` key to settings (do not forget to save the project!).

P2R sets 'Redmine Synchronization' task's dates to last synchronization dates on each run with execution. 
