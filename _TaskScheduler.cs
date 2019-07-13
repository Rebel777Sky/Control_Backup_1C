using System;
using System.Windows.Forms;

namespace TaskScheduler
{
    public partial class _TaskScheduler : Form
    {
        public _TaskScheduler()
        {
            InitializeComponent();
        }

        TaskScheduler objScheduler;
        //To hold Task Definition
        ITaskDefinition objTaskDef;
        //To hold Trigger Information
        ITimeTrigger objTrigger;
        //To hold Action Information
        IExecAction objAction;

        private void btnCreateTask_Click(object sender, EventArgs e)
        {
            try
            {
                objScheduler = new TaskScheduler();
                objScheduler.Connect();

                //Setting Task Definition
                SetTaskDefinition();
                //Setting Task Trigger Information
                SetTriggerInfo();
                //Setting Task Action Information
                SetActionInfo();

                //Getting the roort folder
                ITaskFolder root = objScheduler.GetFolder("\\");
                //Registering the task, if the task is already exist then it will be updated
                IRegisteredTask regTask = root.RegisterTaskDefinition("CreatBackupBase1C", objTaskDef, (int)_TASK_CREATION.TASK_CREATE_OR_UPDATE, null, null, _TASK_LOGON_TYPE.TASK_LOGON_INTERACTIVE_TOKEN, "");

                //To execute the task immediately calling Run()
                IRunningTask runtask = regTask.Run(null);

                MessageBox.Show("Task is created successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //Setting Task Definition
        private void SetTaskDefinition()
        {
            try
            {
                objTaskDef = objScheduler.NewTask(0);
                //Registration Info for task
                //Name of the task Author
                objTaskDef.RegistrationInfo.Author = "AuthorName";
                //Description of the task 
                objTaskDef.RegistrationInfo.Description = "CreatBackupBase1C";
                //Registration date of the task 
                objTaskDef.RegistrationInfo.Date = DateTime.Today.ToString("yyyy-MM-ddTHH:mm:ss"); //Date format 

                //Settings for task
                //Thread Priority
                objTaskDef.Settings.Priority = 7;
                //Enabling the task
                objTaskDef.Settings.Enabled = true;
                //To hide/show the task
                objTaskDef.Settings.Hidden = false;
                //Execution Time Lmit for task
                objTaskDef.Settings.ExecutionTimeLimit = "PT10M"; //10 minutes
                //Specifying no need of network connection
                objTaskDef.Settings.RunOnlyIfNetworkAvailable = false;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //Setting Task Trigger Information
        private void SetTriggerInfo()
        {
            try
            {
                //Trigger information based on time - TASK_TRIGGER_TIME
                objTrigger = (ITimeTrigger)objTaskDef.Triggers.Create(_TASK_TRIGGER_TYPE2.TASK_TRIGGER_TIME);
                //Trigger ID
                objTrigger.Id = "CreatBackupBase1CTrigger";
                //Start Time
                objTrigger.StartBoundary = "2014-01-09T10:10:00"; //yyyy-MM-ddTHH:mm:ss
                //End Time
                objTrigger.EndBoundary = "2020-01-01T07:30:00"; //yyyy-MM-ddTHH:mm:ss
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //Setting Task Action Information
        private void SetActionInfo()
        {
            try
            {
                //Action information based on exe- TASK_ACTION_EXEC
                objAction = (IExecAction)objTaskDef.Actions.Create(_TASK_ACTION_TYPE.TASK_ACTION_EXEC);
                //Action ID
                objAction.Id = "testAction1";
                //Set the path of the exe file to execute, Here mspaint will be opened
                objAction.Path = "mspaint";
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
