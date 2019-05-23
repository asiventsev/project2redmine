# P2R Console script for MS Project to Redmine synchronization

This script is intended to publish any exsisting MSP project to Redmine and get progress data back from Redmine.

MS Project is very convenient for project creation:
* editing big timelines in document mode 
* managing inter-task links and relations
* auto arranging task dates based on relations and resource availability

But is not intended at all for tracking progress and team members performance in self-service mode. 
  
From this point of view Redmine seems to be a very good addition to MS Project. It is simple, free, convenient and customizable web tool for project team to keep tasks progress and to make personal workhours reports.

P2R is designed by the following use-case:
* project manager adds Redmine access settings to MS Project file 
* project manager marks tasks and resources (team members) in MS Project file for publishing to Redmine and then runs P2R.
* P2R creates new project in Redmine and adds team memberships for MSP project resources marked as Redmine users. (Connecting to existing project is also possible.) 
* P2R publishes new tasks with appointments to Redmine including start, finish and workhours estimate. Group tasks will also be published. (Connecting to existing tasks is also possible.)
* project team members put actual start, finish, estimate and hours spent to their tasks in Redmine
* project manager runs P2R periodically to get actual data back to MSP
* P2R sets start, finish, estimate and spent hours in MSP according to Redmine data
* project manager compares data with baselines and corrects remaining part of the projects as necessary

## MSP project preparation

To get msp project ready for publishing to Redmine or updating already published:
1. P2R will keep Redmine task IDs in the Hyperlink attribute of MSP tasks. Put 0 to this field for tasks to be published to Redmine. P2R will publish these tasks to Redmine and change zeros to actual IDs and also will generate task URLs in HyperlinkAddress attribute so the hyperlink will be clickable. Do not forget to save the project!
2. P2R will keep Redmine user IDs in the Hyperlink attribute of MSP tasks. Put Redmine user IDs to this field for Resources to be synchronized to Redmine. Redmine project memberships will be created authomatically. Ask Redmine admin for user IDs.
3. Create special task in MSP project named 'Redmine Synchronization' and put Redmine synchronization 
settings to its Notes in YAML format according to the example:

<pre><code>redmine_host: 192.168.10.23
redmine_port: 3500
redmine_api_key: d1e693e7670b1fc79378441d485b9608e4c2dc4c
redmine_project_uuid: uute
task_default_redmine_tracker_id: 4
resource_default_redmine_role_id: 4
</code></pre>

That's all. The meaning of the parameters is clear from their names. 

Parameter `redmine_project_uuid` must correspond Redmine rules which you could find on in settings pages of any Redmine project.
If an 'Unprocessable Entry' error occurs on project creation it is probably due to incorrect project uuid.

Parameter `task_default_redmine_tracker_id` is not required. If you omit it then Redmine's default tracker will be set.

If you want to connect your MSP project to existing Redmine project instead of creating new one, please do the following:

1. Add parameter redmine_project_id to settings list and set it to Redmine project ID (not uuid!). 
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

P2R sets 'Redmine Synchronization' task's dates to last synchronization dates on each run with execution. 
