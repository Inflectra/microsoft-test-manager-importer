using System;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using System.IO;
using System.Net;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Threading;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;

using Inflectra.SpiraTest.AddOns.MTMImporter.SpiraImportExport;
using System.ServiceModel;

using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Proxy;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.Framework.Client;

namespace Inflectra.SpiraTest.AddOns.MTMImporter
{
    /// <summary>
    /// This is the code behind class for the utility that imports projects from
    /// HP Mercury Quality Center / TestDirector into Inflectra SpiraTest
    /// </summary>
    public class ProgressForm : System.Windows.Forms.Form
    {
        //Artifact Types
        public enum ArtifactType
        {
            None = -1,
            Requirement = 1,
            TestCase = 2,
            Incident = 3,
            Release = 4,
            TestRun = 5,
            Task = 6,
            TestStep = 7,
            TestSet = 8,
            AutomationHost = 9
        }

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnCancel;

        private const string DEFAULT_PASSWORD = "PleaseChange123ABC$!@";

        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;
        protected MainForm mainForm;
        protected ImportForm importForm;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Label lblProgress2;
        private System.Windows.Forms.Label lblProgress3;
        private System.Windows.Forms.Label lblProgress4;
        private System.Windows.Forms.Label lblProgress5;
        private System.Windows.Forms.Label lblProgress1;
        protected CookieContainer cookieContainer;
        protected Dictionary<int, int> releaseMapping = new Dictionary<int,int>();
        protected Dictionary<string, int> releasePathMapping = new Dictionary<string, int>();
        protected Dictionary<int, int> testPlanFolderMapping = new Dictionary<int, int>();
        protected Dictionary<int, int> testSuiteFolderMapping = new Dictionary<int,int>();
        protected Dictionary<int, int> testCaseMapping = new Dictionary<int,int>();
        protected Dictionary<int, int> testStepMapping = new Dictionary<int,int>();
        protected Dictionary<int, int> usersMapping = new Dictionary<int,int>();
        protected Dictionary<int, int> sharedStepMapping = new Dictionary<int, int>();
        protected Dictionary<string, int> areaCustomPropertyValueMapping = new Dictionary<string, int>();
        protected Dictionary<string, int> stateCustomPropertyValueMapping = new Dictionary<string, int>();
        protected List<MtmTestSuiteTestCaseMapping> testSuiteTestCaseMappings = new List<MtmTestSuiteTestCaseMapping>();
        protected Dictionary<int, int> testPlanTestSetFolderMapping = new Dictionary<int, int>();
        protected Dictionary<int, int> testSuiteTestSetMapping = new Dictionary<int, int>();
        protected Dictionary<int, int> testSuiteEntryTestSetTestCaseIdMapping = new Dictionary<int, int>();

        private Label lblProgress6;
        private ProgressBar progressBar1;

        protected const int SPIRA_ARTIFACT_ID_REQUIREMENT = 1;

        #region Properties

        public MainForm MainFormHandle
        {
            get
            {
                return this.mainForm;
            }
            set
            {
                this.mainForm = value;
            }
        }

        public ImportForm ImportFormHandle
        {
            get
            {
                return this.importForm;
            }
            set
            {
                this.importForm = value;
            }
        }

        #endregion

        public ProgressForm()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            // Add any event handlers
            this.Closing += new CancelEventHandler(ProgressForm_Closing);
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProgressForm));
            this.btnExit = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblProgress6 = new System.Windows.Forms.Label();
            this.lblProgress5 = new System.Windows.Forms.Label();
            this.lblProgress4 = new System.Windows.Forms.Label();
            this.lblProgress3 = new System.Windows.Forms.Label();
            this.lblProgress2 = new System.Windows.Forms.Label();
            this.lblProgress1 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(408, 286);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(96, 23);
            this.btnExit.TabIndex = 0;
            this.btnExit.Text = "Done";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(312, 286);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(88, 23);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(16, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(424, 23);
            this.label1.TabIndex = 6;
            this.label1.Text = "SpiraTest | Import From MS Test Manager";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.InitialImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.InitialImage")));
            this.pictureBox1.Location = new System.Drawing.Point(472, 8);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(40, 40);
            this.pictureBox1.TabIndex = 7;
            this.pictureBox1.TabStop = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblProgress6);
            this.groupBox1.Controls.Add(this.lblProgress5);
            this.groupBox1.Controls.Add(this.lblProgress4);
            this.groupBox1.Controls.Add(this.lblProgress3);
            this.groupBox1.Controls.Add(this.lblProgress2);
            this.groupBox1.Controls.Add(this.lblProgress1);
            this.groupBox1.ForeColor = System.Drawing.Color.Black;
            this.groupBox1.Location = new System.Drawing.Point(24, 61);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(480, 184);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Import Progress";
            // 
            // lblProgress6
            // 
            this.lblProgress6.Location = new System.Drawing.Point(40, 142);
            this.lblProgress6.Name = "lblProgress6";
            this.lblProgress6.Size = new System.Drawing.Size(368, 16);
            this.lblProgress6.TabIndex = 26;
            this.lblProgress6.Text = ">>    Incidents Imported";
            // 
            // lblProgress5
            // 
            this.lblProgress5.Location = new System.Drawing.Point(40, 118);
            this.lblProgress5.Name = "lblProgress5";
            this.lblProgress5.Size = new System.Drawing.Size(368, 16);
            this.lblProgress5.TabIndex = 25;
            this.lblProgress5.Text = ">>    Test Runs Imported";
            // 
            // lblProgress4
            // 
            this.lblProgress4.Location = new System.Drawing.Point(40, 94);
            this.lblProgress4.Name = "lblProgress4";
            this.lblProgress4.Size = new System.Drawing.Size(368, 16);
            this.lblProgress4.TabIndex = 24;
            this.lblProgress4.Text = ">>    Test Sets Imported";
            // 
            // lblProgress3
            // 
            this.lblProgress3.Location = new System.Drawing.Point(40, 70);
            this.lblProgress3.Name = "lblProgress3";
            this.lblProgress3.Size = new System.Drawing.Size(368, 16);
            this.lblProgress3.TabIndex = 22;
            this.lblProgress3.Text = ">>    Test Cases Imported";
            // 
            // lblProgress2
            // 
            this.lblProgress2.Location = new System.Drawing.Point(40, 47);
            this.lblProgress2.Name = "lblProgress2";
            this.lblProgress2.Size = new System.Drawing.Size(368, 16);
            this.lblProgress2.TabIndex = 21;
            this.lblProgress2.Text = ">>    Releases Imported";
            // 
            // lblProgress1
            // 
            this.lblProgress1.Location = new System.Drawing.Point(40, 24);
            this.lblProgress1.Name = "lblProgress1";
            this.lblProgress1.Size = new System.Drawing.Size(368, 16);
            this.lblProgress1.TabIndex = 26;
            this.lblProgress1.Text = ">>    Users Imported";
            // 
            // progressBar1
            // 
            this.progressBar1.ForeColor = System.Drawing.Color.OrangeRed;
            this.progressBar1.Location = new System.Drawing.Point(-3, 268);
            this.progressBar1.Maximum = 6;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(533, 10);
            this.progressBar1.TabIndex = 28;
            // 
            // ProgressForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 14);
            this.ClientSize = new System.Drawing.Size(526, 314);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnExit);
            this.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "ProgressForm";
            this.Text = "Microsoft Test Manager Importer for Inflectra SpiraTest";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }
        #endregion

        /// <summary>
        /// Closes the application
        /// </summary>
        /// <param name="sender">The sending object</param>
        /// <param name="e">The event arguments</param>
        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            //Hide the current form
            this.Hide();

            //Return to the main form
            MainFormHandle.Show();
        }

        /// <summary>
        /// Closes the application
        /// </summary>
        /// <param name="sender">The sending object</param>
        /// <param name="e">The event arguments</param>
        private void btnExit_Click(object sender, System.EventArgs e)
        {
            //Close the application
            this.MainFormHandle.Close();
        }

        /// <summary>
        /// Called if the form is closed
        /// </summary>
        /// <param name="sender">The sending object</param>
        /// <param name="e">The event arguments</param>
        private void ProgressForm_Closing(object sender, CancelEventArgs e)
        {
            /*
            //Disconnect/Logout if connected
            if (MainFormHandle.TdConnection != null)
            {
                if (MainFormHandle.TdConnection.Connected)
                {
                    MainFormHandle.TdConnection.Disconnect();
                }

                if (MainFormHandle.TdConnection.LoggedIn)
                {
                    MainFormHandle.TdConnection.Logout();
                }
            }*/
        }

        /// <summary>
        /// Starts the background thread for importing the data
        /// </summary>
        public void StartImport()
        {
            //Set the initial state of any buttons
            this.btnCancel.Enabled = true;
            this.btnExit.Enabled = false;

            //Clear the progress labels
            ProgressForm_OnProgressUpdate(0);

            //First change the cursor to an hourglass
            this.Cursor = Cursors.WaitCursor;

            //Start the background thread that performs the import
            ThreadPool.QueueUserWorkItem(new WaitCallback(this.ImportData));
        }

        /// <summary>
        /// Updates the form state when the import has finished
        /// </summary>
        public void ProgressForm_OnFinish()
        {
            //Change the cursor back to the default
            this.Cursor = Cursors.Default;

            //Enable the Exit button and disable cancel
            this.btnCancel.Enabled = false;
            this.btnExit.Enabled = true;

            //Display a message
            MessageBox.Show("SpiraTest Import Successful!", "Import", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Displays any errors raised by the import thread process
        /// </summary>
        /// <param name="exception">The exception raised</param>
        public void ProgressForm_OnError(Exception exception)
        {
            //Change the cursor back to the default
            this.Cursor = Cursors.Default;

            //Enable the Exit button and disable cancel
            this.btnCancel.Enabled = false;
            this.btnExit.Enabled = true;

            //Display the exception error message
            MessageBox.Show("SpiraTest Import Failed. Error: " + exception.Message, "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// Updates the progress display of the form
        /// </summary>
        /// <param name="progress">The progress state</param>
        public void ProgressForm_OnProgressUpdate(int progress)
        {
            //Make all the controls up to the specified one visible
            for (int i = 1; i <= 6; i++)
            {
                this.Controls.Find("lblProgress" + i.ToString(), true)[0].Visible = (i <= progress);
            }

            //Also update the progress bar
            this.progressBar1.Value = progress;
        }

        /// <summary>
        /// Imports all the test steps in a test case
        /// </summary>
        /// <param name="streamWriter"></param>
        /// <param name="mtmTestCase">The MTM test case</param>
        /// <param name="newTestCaseId">The ID of the corresponding SpiraTest test case</param>
        protected void ImportTestStepsInTestCase(StreamWriter streamWriter, ITestCase mtmTestCase, int newTestCaseId)
        {
            //Now get the test steps that belong to this test case
            if (mtmTestCase.Actions.Count > 0)
            {
                streamWriter.WriteLine("Getting Test Steps for test case " + mtmTestCase.Id);

                int position = 1;
                foreach (ITestAction testAction in mtmTestCase.Actions)
                {
                    //See if this is a shared step or not (need to handle differently)
                    if (testAction is ISharedStepReference)
                    {
                        ISharedStepReference mtmTestStepRef = (ISharedStepReference)testAction;
                        //Extract the test step info
                        int stepId = mtmTestStepRef.Id;
                        int mtmLinkedTestCaseId = mtmTestStepRef.SharedStepId;
                        streamWriter.WriteLine("Importing Shared Test Step: " + stepId.ToString());

                        //We need to find the matching Spira test case ID for the test step
                        if (this.sharedStepMapping.ContainsKey(mtmLinkedTestCaseId))
                        {
                            int linkedTestCaseId = this.sharedStepMapping[mtmLinkedTestCaseId];
                            try
                            {
                                RemoteTestStepParameter[] parameters = new RemoteTestStepParameter[0];
                                int newTestStepId = ImportFormHandle.SpiraImportProxy.TestCase_AddLink(newTestCaseId, position, linkedTestCaseId, parameters);

                                //Add to the mapping hashtable
                                if (!this.testStepMapping.ContainsKey(stepId))
                                {
                                    this.testStepMapping.Add(stepId, newTestStepId);
                                }

                                //The shared step reference doesn't have any attachments of its own
                            }
                            catch (Exception exception)
                            {
                                //If we have an error, log it and continue
                                streamWriter.WriteLine("Error adding shared test step " + stepId + " to SpiraTest (" + exception.Message + ")");
                            }
                            position++;
                        }
                        else
                        {
                            streamWriter.WriteLine("Warning: Could not find Spira test case that matches the MTM shared test step work item " + mtmLinkedTestCaseId + " so skipping this shared test step");
                        }
                    }
                    else if (testAction is ITestStep)
                    {
                        ITestStep mtmTestStep = (ITestStep)testAction;
                        //Extract the test step info
                        int stepId = mtmTestStep.Id;
                        streamWriter.WriteLine("Importing Test Step: " + stepId.ToString());

                        string stepTitle = MakeXmlSafe(mtmTestStep.Title);
                        string stepDescription = MakeXmlSafe(mtmTestStep.Description);
                        string stepExpectedResult = MakeXmlSafe(mtmTestStep.ExpectedResult);

                        //Convert any parameters
                        stepTitle = ConvertParameters(stepTitle, mtmTestCase.TestParameters);
                        stepDescription = ConvertParameters(stepDescription, mtmTestCase.TestParameters);
                        stepExpectedResult = ConvertParameters(stepExpectedResult, mtmTestCase.TestParameters);

                        try
                        {
                            RemoteTestStep remoteTestStep = new RemoteTestStep();
                            remoteTestStep.TestCaseId = newTestCaseId;
                            remoteTestStep.Position = position;
                            remoteTestStep.Description = stepTitle;
                            remoteTestStep.ExpectedResult = stepExpectedResult;
                            remoteTestStep.SampleData = stepDescription;
                            int newTestStepId = ImportFormHandle.SpiraImportProxy.TestCase_AddStep(remoteTestStep, newTestCaseId).TestStepId.Value;

                            //Add to the mapping hashtable
                            if (!this.testStepMapping.ContainsKey(stepId))
                            {
                                this.testStepMapping.Add(stepId, newTestStepId);
                            }

                            //Add attachments if requested
                            if (this.MainFormHandle.chkImportAttachments.Checked)
                            {
                                try
                                {
                                    ImportAttachments(streamWriter, mtmTestStep.Attachments, newTestStepId, ArtifactType.TestStep);
                                }
                                catch (Exception exception)
                                {
                                    streamWriter.WriteLine("Warning: Unable to import attachments for test step " + stepId + " (" + exception.Message + ")");
                                }
                            }
                        }
                        catch (Exception exception)
                        {
                            //If we have an error, log it and continue
                            streamWriter.WriteLine("Error adding test step " + stepId + " to SpiraTest (" + exception.Message + ")");
                        }
                        position++;
                    }
                }
            }
        }

        protected void ImportSharedStepsSteps(StreamWriter streamWriter, int projectId, string iterationPath, string areaPath, ITestManagementTeamProject mtmProject)
        {
            //Get all the shared steps in this area/iteration
            streamWriter.WriteLine(String.Format("Searching for the test steps of Shared Steps in Iteration='{0}', Area='{1}'", iterationPath, areaPath));
            string query = String.Format("Select * From WorkItems Where [System.IterationPath] = '{0}' And [System.AreaPath] = '{1}' Order By [System.Id]", iterationPath, areaPath);
            IEnumerable<ISharedStep> mtmSharedSteps = mtmProject.SharedSteps.Query(query);
            streamWriter.WriteLine(String.Format("Found {0} shared steps", mtmSharedSteps.Count()));

            //Reconnect to Spira
            ReconnectToSpira(projectId);

            foreach (ISharedStep mtmSharedStep in mtmSharedSteps)
            {
                int testId = mtmSharedStep.Id;
                if (this.sharedStepMapping.ContainsKey(testId))
                {
                    int newTestCaseId = this.sharedStepMapping[testId];
                    //Now get the test steps that belong to this shared test case
                    if (mtmSharedStep.Actions.Count > 0)
                    {
                        streamWriter.WriteLine("Getting Test Steps for shared test case " + testId);

                        int position = 1;
                        foreach (ITestAction testAction in mtmSharedStep.Actions)
                        {
                            //See if this is a shared step or not (need to handle differently)
                            if (testAction is ISharedStepReference)
                            {
                                ISharedStepReference mtmTestStepRef = (ISharedStepReference)testAction;
                                //Extract the test step info
                                int stepId = mtmTestStepRef.Id;
                                int mtmLinkedTestCaseId = mtmTestStepRef.SharedStepId;
                                streamWriter.WriteLine("Importing Shared Test Step: " + stepId.ToString());

                                //We need to find the matching Spira test case ID for the test step
                                if (this.sharedStepMapping.ContainsKey(mtmLinkedTestCaseId))
                                {
                                    int linkedTestCaseId = this.sharedStepMapping[mtmLinkedTestCaseId];
                                    try
                                    {
                                        RemoteTestStepParameter[] parameters = new RemoteTestStepParameter[0];
                                        int newTestStepId = ImportFormHandle.SpiraImportProxy.TestCase_AddLink(newTestCaseId, position, linkedTestCaseId, parameters);

                                        //Add to the mapping hashtable
                                        if (!this.testStepMapping.ContainsKey(stepId))
                                        {
                                            this.testStepMapping.Add(stepId, newTestStepId);
                                        }

                                        //The shared step reference doesn't have any attachments of its own
                                    }
                                    catch (Exception exception)
                                    {
                                        //If we have an error, log it and continue
                                        streamWriter.WriteLine("Error adding shared test step " + stepId + " to SpiraTest (" + exception.Message + ")");
                                    }
                                    position++;
                                }
                                else
                                {
                                    streamWriter.WriteLine("Warning: Could not find Spira test case that matches the MTM shared test step work item " + mtmLinkedTestCaseId + " so skipping this shared test step");
                                }
                            }
                            else if (testAction is ITestStep)
                            {
                                ITestStep mtmTestStep = (ITestStep)testAction;
                                //Extract the test step info
                                int stepId = mtmTestStep.Id;
                                streamWriter.WriteLine("Importing Test Step: " + stepId.ToString());

                                string stepTitle = MakeXmlSafe(mtmTestStep.Title);
                                string stepDescription = MakeXmlSafe(mtmTestStep.Description);
                                string stepExpectedResult = MakeXmlSafe(mtmTestStep.ExpectedResult);

                                //Convert any parameters
                                stepTitle = ConvertParameters(stepTitle, mtmSharedStep.TestParameters);
                                stepDescription = ConvertParameters(stepDescription, mtmSharedStep.TestParameters);
                                stepExpectedResult = ConvertParameters(stepExpectedResult, mtmSharedStep.TestParameters);

                                try
                                {
                                    RemoteTestStep remoteTestStep = new RemoteTestStep();
                                    remoteTestStep.TestCaseId = newTestCaseId;
                                    remoteTestStep.Position = position;
                                    remoteTestStep.Description = stepTitle;
                                    remoteTestStep.ExpectedResult = stepExpectedResult;
                                    remoteTestStep.SampleData = stepDescription;
                                    int newTestStepId = ImportFormHandle.SpiraImportProxy.TestCase_AddStep(remoteTestStep, newTestCaseId).TestStepId.Value;

                                    //Add to the mapping hashtable
                                    if (!this.testStepMapping.ContainsKey(stepId))
                                    {
                                        this.testStepMapping.Add(stepId, newTestStepId);
                                    }

                                    //Add attachments if requested
                                    if (this.MainFormHandle.chkImportAttachments.Checked)
                                    {
                                        try
                                        {
                                            ImportAttachments(streamWriter, mtmTestStep.Attachments, newTestStepId, ArtifactType.TestStep);
                                        }
                                        catch (Exception exception)
                                        {
                                            streamWriter.WriteLine("Warning: Unable to import attachments for test step " + stepId + " (" + exception.Message + ")");
                                        }
                                    }
                                }
                                catch (Exception exception)
                                {
                                    //If we have an error, log it and continue
                                    streamWriter.WriteLine("Error adding test step " + stepId + " to SpiraTest (" + exception.Message + ")");
                                }
                                position++;
                            }
                        }
                    }
                }
            }
        }

        protected void ImportedSharedSteps(StreamWriter streamWriter, int projectId, string iterationPath, string areaPath, ITestManagementTeamProject mtmProject, RemoteCustomList remoteCustomList_Area, RemoteCustomList remoteCustomList_State, int newSharedTestCasesFolderId)
        {
            //Get all the shared steps in this area/iteration
            streamWriter.WriteLine(String.Format("Searching for Shared Steps in Iteration='{0}', Area='{1}'", iterationPath, areaPath));
            string query = String.Format("Select * From WorkItems Where [System.IterationPath] = '{0}' And [System.AreaPath] = '{1}' Order By [System.Id]", iterationPath, areaPath);
            IEnumerable<ISharedStep> mtmSharedSteps = mtmProject.SharedSteps.Query(query);
            streamWriter.WriteLine(String.Format("Found {0} shared steps", mtmSharedSteps.Count()));

            //Reconnect to Spira
            ReconnectToSpira(projectId);

            foreach (ISharedStep mtmSharedStep in mtmSharedSteps)
            {
                //Extract the test info
                int testId = mtmSharedStep.Id;
                streamWriter.WriteLine("Importing Shared Step: " + testId.ToString());
                string testCaseName = MakeXmlSafe(mtmSharedStep.Title);
                string testCaseDescription = MakeXmlSafe(mtmSharedStep.Description);

                //Load the test case and capture the new id
                try
                {
                    //Populate the test case
                    RemoteTestCase remoteTestCase = new RemoteTestCase();
                    remoteTestCase.Name = testCaseName;
                    remoteTestCase.Description = testCaseDescription;
                    remoteTestCase.Active = (mtmSharedStep.State != "Closed");  //All map to active except Closed
                    if (mtmSharedStep.Priority >= 1 && mtmSharedStep.Priority <= 4)
                    {
                        remoteTestCase.TestCasePriorityId = mtmSharedStep.Priority;   //(1,2,3,4 same in both systems)
                    }
                    if (mtmSharedStep.Owner != null && usersMapping.ContainsKey(mtmSharedStep.Owner.UniqueUserId))
                    {
                        remoteTestCase.OwnerId = usersMapping[mtmSharedStep.Owner.UniqueUserId];
                    }

                    //Handle the Area and State custom properties
                    List<RemoteArtifactCustomProperty> remoteTestCaseCustomProperties = new List<RemoteArtifactCustomProperty>();
                    if (!String.IsNullOrEmpty(mtmSharedStep.Area))
                    {
                        if (this.areaCustomPropertyValueMapping.ContainsKey(mtmSharedStep.Area))
                        {
                            //We have already imported this area
                            int customPropertyValueId = this.areaCustomPropertyValueMapping[mtmSharedStep.Area];
                            remoteTestCaseCustomProperties.Add(new RemoteArtifactCustomProperty() { PropertyNumber = 1, IntegerValue = customPropertyValueId });
                        }
                        else
                        {
                            streamWriter.WriteLine("Adding new Area: '" + mtmSharedStep.Area + "' as SpiraTest custom property value");
                            RemoteCustomListValue remoteCustomListValue = new RemoteCustomListValue();
                            remoteCustomListValue.CustomPropertyListId = remoteCustomList_Area.CustomPropertyListId.Value;
                            remoteCustomListValue.Name = mtmSharedStep.Area;
                            remoteCustomListValue = ImportFormHandle.SpiraImportProxy.CustomProperty_AddCustomListValue(remoteCustomListValue);
                            this.areaCustomPropertyValueMapping.Add(mtmSharedStep.Area, remoteCustomListValue.CustomPropertyValueId.Value);
                            remoteTestCaseCustomProperties.Add(new RemoteArtifactCustomProperty() { PropertyNumber = 1, IntegerValue = remoteCustomListValue.CustomPropertyValueId.Value });
                        }
                    }
                    if (!String.IsNullOrEmpty(mtmSharedStep.State))
                    {
                        if (this.stateCustomPropertyValueMapping.ContainsKey(mtmSharedStep.State))
                        {
                            //We have already imported this area
                            int customPropertyValueId = this.stateCustomPropertyValueMapping[mtmSharedStep.State];
                            remoteTestCaseCustomProperties.Add(new RemoteArtifactCustomProperty() { PropertyNumber = 2, IntegerValue = customPropertyValueId });
                        }
                        else
                        {
                            streamWriter.WriteLine("Adding new State: '" + mtmSharedStep.State + "' as SpiraTest custom property value");
                            RemoteCustomListValue remoteCustomListValue = new RemoteCustomListValue();
                            remoteCustomListValue.CustomPropertyListId = remoteCustomList_State.CustomPropertyListId.Value;
                            remoteCustomListValue.Name = mtmSharedStep.State;
                            remoteCustomListValue = ImportFormHandle.SpiraImportProxy.CustomProperty_AddCustomListValue(remoteCustomListValue);
                            this.stateCustomPropertyValueMapping.Add(mtmSharedStep.State, remoteCustomListValue.CustomPropertyValueId.Value);
                            remoteTestCaseCustomProperties.Add(new RemoteArtifactCustomProperty() { PropertyNumber = 2, IntegerValue = remoteCustomListValue.CustomPropertyValueId.Value });
                        }
                    }
                    remoteTestCase.CustomProperties = remoteTestCaseCustomProperties.ToArray();

                    //Shared Steps will live in a special Shared Steps top-level folder
                    int newTestCaseId = ImportFormHandle.SpiraImportProxy.TestCase_Create(remoteTestCase, newSharedTestCasesFolderId).TestCaseId.Value;

                    //Add to the mapping dictionary
                    if (!this.sharedStepMapping.ContainsKey(testId))
                    {
                        this.sharedStepMapping.Add(testId, newTestCaseId);
                    }

                    //Add the input parameters if there are any
                    if (mtmSharedStep.TestParameters != null && mtmSharedStep.TestParameters.Count > 0)
                    {
                        foreach (ITestParameter mtmTestParameter in mtmSharedStep.TestParameters)
                        {
                            //Add the new test parameter to Spira
                            string parameterName = mtmTestParameter.Name;
                            string parameterDefaultValue = mtmTestParameter.Value;

                            if (!String.IsNullOrWhiteSpace(parameterName))
                            {
                                streamWriter.WriteLine("Adding parameter: '" + parameterName + "' to SpiraTest test case TC" + newTestCaseId);
                                RemoteTestCaseParameter remoteTestCaseParameter = new RemoteTestCaseParameter();
                                remoteTestCaseParameter.TestCaseId = newTestCaseId;
                                remoteTestCaseParameter.Name = parameterName;
                                remoteTestCaseParameter.DefaultValue = parameterDefaultValue;
                                ImportFormHandle.SpiraImportProxy.TestCase_AddParameter(remoteTestCaseParameter);
                            }
                        }
                    }

                    //Add attachments if requested
                    if (this.MainFormHandle.chkImportAttachments.Checked)
                    {
                        try
                        {
                            ImportAttachments(streamWriter, mtmSharedStep.Attachments, newTestCaseId, ArtifactType.TestCase);
                        }
                        catch (Exception exception)
                        {
                            streamWriter.WriteLine("Warning: Unable to import attachments for shared step " + testId + " (" + exception.Message + ")");
                        }
                    }

                    //See if we have an iteration path and add the test case to the release (unless already added)
                    if (mtmSharedStep.WorkItem != null && !String.IsNullOrEmpty(mtmSharedStep.WorkItem.IterationPath))
                    {
                        if (this.releasePathMapping.ContainsKey(mtmSharedStep.WorkItem.IterationPath))
                        {
                            int releaseId = this.releasePathMapping[mtmSharedStep.WorkItem.IterationPath];
                            string releaseAndTestCase = releaseId + ":" + newTestCaseId;
                            streamWriter.WriteLine(String.Format("Adding Test Case TC{0} to Release RL{1} in SpiraTest.", newTestCaseId, releaseId));

                            RemoteReleaseTestCaseMapping newMapping = new RemoteReleaseTestCaseMapping();
                            newMapping.ReleaseId = releaseId;
                            newMapping.TestCaseId = newTestCaseId;
                            ImportFormHandle.SpiraImportProxy.Release_AddTestMapping(newMapping);
                        }
                    }
                }
                catch (FaultException<ServiceFaultMessage> exception)
                {
                    //If we have an error, log it and continue
                    streamWriter.WriteLine("Error adding shared step " + testId + " to SpiraTest (" + exception.Message + ")");
                }
                catch (Exception exception)
                {
                    //If we have an error, log it and continue
                    streamWriter.WriteLine("Error adding shared step " + testId + " to SpiraTest (" + exception.Message + ")");
                }
            }
        }

        /// <summary>
        /// Imports an MTM test case into SpiraTest under the specified folder (or root if null)
        /// </summary>
        /// <param name="mtmTestCase">The MTM test case</param>
        /// <param name="testCaseFolderId">The Spira test case folder Id</param>
        protected void ImportTestCase(StreamWriter streamWriter, ITestCase mtmTestCase, RemoteCustomList remoteCustomList_Area, RemoteCustomList remoteCustomList_State)
        {
            //Extract the test info
            int testId = mtmTestCase.Id;
            streamWriter.WriteLine("Importing Test Case: " + testId.ToString());
            string testCaseName = MakeXmlSafe(mtmTestCase.Title);
            string testCaseDescription = MakeXmlSafe(mtmTestCase.Description);

            //Load the test case and capture the new id
            try
            {
                //Populate the test case
                RemoteTestCase remoteTestCase = new RemoteTestCase();
                remoteTestCase.Name = testCaseName;
                remoteTestCase.Description = testCaseDescription;
                remoteTestCase.Active = (mtmTestCase.State == "Ready");  //All map to inactive except Ready
                if (mtmTestCase.Priority >= 1 && mtmTestCase.Priority <= 4)
                {
                    remoteTestCase.TestCasePriorityId = mtmTestCase.Priority;   //(1,2,3,4 same in both systems)
                }
                if (mtmTestCase.Owner != null && usersMapping.ContainsKey(mtmTestCase.Owner.UniqueUserId))
                {
                    remoteTestCase.OwnerId = usersMapping[mtmTestCase.Owner.UniqueUserId];
                }

                //Handle the Area and State custom properties
                List<RemoteArtifactCustomProperty> remoteTestCaseCustomProperties = new List<RemoteArtifactCustomProperty>();
                if (!String.IsNullOrEmpty(mtmTestCase.Area))
                {
                    if (this.areaCustomPropertyValueMapping.ContainsKey(mtmTestCase.Area))
                    {
                        //We have already imported this area
                        int customPropertyValueId = this.areaCustomPropertyValueMapping[mtmTestCase.Area];
                        remoteTestCaseCustomProperties.Add(new RemoteArtifactCustomProperty() { PropertyNumber = 1, IntegerValue = customPropertyValueId });
                    }
                    else
                    {
                        streamWriter.WriteLine("Adding new Area: '" + mtmTestCase.Area + "' as SpiraTest custom property value");
                        RemoteCustomListValue remoteCustomListValue = new RemoteCustomListValue();
                        remoteCustomListValue.CustomPropertyListId = remoteCustomList_Area.CustomPropertyListId.Value;
                        remoteCustomListValue.Name = mtmTestCase.Area;
                        remoteCustomListValue = ImportFormHandle.SpiraImportProxy.CustomProperty_AddCustomListValue(remoteCustomListValue);
                        this.areaCustomPropertyValueMapping.Add(mtmTestCase.Area, remoteCustomListValue.CustomPropertyValueId.Value);
                        remoteTestCaseCustomProperties.Add(new RemoteArtifactCustomProperty() { PropertyNumber = 1, IntegerValue = remoteCustomListValue.CustomPropertyValueId.Value });
                    }
                }
                if (!String.IsNullOrEmpty(mtmTestCase.State))
                {
                    if (this.stateCustomPropertyValueMapping.ContainsKey(mtmTestCase.State))
                    {
                        //We have already imported this area
                        int customPropertyValueId = this.stateCustomPropertyValueMapping[mtmTestCase.State];
                        remoteTestCaseCustomProperties.Add(new RemoteArtifactCustomProperty() { PropertyNumber = 2, IntegerValue = customPropertyValueId });
                    }
                    else
                    {
                        streamWriter.WriteLine("Adding new State: '" + mtmTestCase.State + "' as SpiraTest custom property value");
                        RemoteCustomListValue remoteCustomListValue = new RemoteCustomListValue();
                        remoteCustomListValue.CustomPropertyListId = remoteCustomList_State.CustomPropertyListId.Value;
                        remoteCustomListValue.Name = mtmTestCase.State;
                        remoteCustomListValue = ImportFormHandle.SpiraImportProxy.CustomProperty_AddCustomListValue(remoteCustomListValue);
                        this.stateCustomPropertyValueMapping.Add(mtmTestCase.State, remoteCustomListValue.CustomPropertyValueId.Value);
                        remoteTestCaseCustomProperties.Add(new RemoteArtifactCustomProperty() { PropertyNumber = 2, IntegerValue = remoteCustomListValue.CustomPropertyValueId.Value });
                    }
                }
                remoteTestCase.CustomProperties = remoteTestCaseCustomProperties.ToArray();

                //Get the appropriate folder to put it under (maps to a MTM test suite)
                int? folderId = null;
                MtmTestSuiteTestCaseMapping mtmTestSuiteTestCaseMapping = this.testSuiteTestCaseMappings.FirstOrDefault(t => t.TestCaseId == mtmTestCase.Id);
                if (mtmTestSuiteTestCaseMapping != null)
                {
                    if (testSuiteFolderMapping.ContainsKey(mtmTestSuiteTestCaseMapping.TestSuiteId))
                    {
                        folderId = testSuiteFolderMapping[mtmTestSuiteTestCaseMapping.TestSuiteId];
                    }
                }
                /*
                if (!folderId.HasValue)
                {
                    streamWriter.WriteLine("Warning: Unable to locate test folder to add test " + testId + " to, so adding to unmapped folder");
                    folderId = newUnmappedTestCasesFolderId;
                }*/
                int newTestCaseId = ImportFormHandle.SpiraImportProxy.TestCase_Create(remoteTestCase, folderId).TestCaseId.Value;

                //Add to the mapping dictionary
                this.testCaseMapping.Add(testId, newTestCaseId);

                //Add the input parameters if there are any
                if (mtmTestCase.TestParameters != null && mtmTestCase.TestParameters.Count > 0)
                {
                    foreach (ITestParameter mtmTestParameter in mtmTestCase.TestParameters)
                    {
                        //Add the new test parameter to Spira
                        string parameterName = mtmTestParameter.Name;
                        string parameterDefaultValue = mtmTestParameter.Value;

                        if (!String.IsNullOrWhiteSpace(parameterName))
                        {
                            streamWriter.WriteLine("Adding parameter: '" + parameterName + "' to SpiraTest test case TC" + newTestCaseId);
                            RemoteTestCaseParameter remoteTestCaseParameter = new RemoteTestCaseParameter();
                            remoteTestCaseParameter.TestCaseId = newTestCaseId;
                            remoteTestCaseParameter.Name = parameterName;
                            remoteTestCaseParameter.DefaultValue = parameterDefaultValue;
                            ImportFormHandle.SpiraImportProxy.TestCase_AddParameter(remoteTestCaseParameter);
                        }
                    }
                }

                //Add attachments if requested
                if (this.MainFormHandle.chkImportAttachments.Checked)
                {
                    try
                    {
                        ImportAttachments(streamWriter, mtmTestCase.Attachments, newTestCaseId, ArtifactType.TestCase);
                    }
                    catch (Exception exception)
                    {
                        streamWriter.WriteLine("Warning: Unable to import attachments for test " + testId + " (" + exception.Message + ")");
                    }
                }

                //See if we have an iteration path and add the test case to the release (unless already added)
                if (mtmTestCase.WorkItem != null && !String.IsNullOrEmpty(mtmTestCase.WorkItem.IterationPath))
                {
                    if (this.releasePathMapping.ContainsKey(mtmTestCase.WorkItem.IterationPath))
                    {
                        int releaseId = this.releasePathMapping[mtmTestCase.WorkItem.IterationPath];
                        string releaseAndTestCase = releaseId + ":" + newTestCaseId;
                        streamWriter.WriteLine(String.Format("Adding Test Case TC{0} to Release RL{1} in SpiraTest.", newTestCaseId, releaseId));

                        RemoteReleaseTestCaseMapping newMapping = new RemoteReleaseTestCaseMapping();
                        newMapping.ReleaseId = releaseId;
                        newMapping.TestCaseId = newTestCaseId;
                        ImportFormHandle.SpiraImportProxy.Release_AddTestMapping(newMapping);
                    }
                }

                //Now get the test steps
                ImportTestStepsInTestCase(streamWriter, mtmTestCase, newTestCaseId);
            }
            catch (FaultException<ServiceFaultMessage> exception)
            {
                //If we have an error, log it and continue
                streamWriter.WriteLine("Error adding test case " + testId + " to SpiraTest (" + exception.Message + ")");
            }
            catch (Exception exception)
            {
                //If we have an error, log it and continue
                streamWriter.WriteLine("Error adding test case " + testId + " to SpiraTest (" + exception.Message + ")");
            }
        }

        /// <summary>
        /// Imports the test cases under a suite recursively
        /// </summary>
        protected void ImportTestSuiteTestCases(StreamWriter streamWriter, int projectId, IDynamicTestSuite mtmTestSuite, RemoteCustomList remoteCustomList_Area, RemoteCustomList remoteCustomList_State)
        {

            //Get the test cases to add to the test folder
            if (mtmTestSuite.TestCases != null && mtmTestSuite.TestCases.Count > 0)
            {
                streamWriter.WriteLine(String.Format("Importing {0} test cases in suite: {1}", mtmTestSuite.TestCases.Count, mtmTestSuite.Title));
                foreach (ITestSuiteEntry mtmTestSuiteEntry in mtmTestSuite.TestCases)
                {
                    if (mtmTestSuiteEntry.TestCase != null)
                    {
                        //Reconnect to Spira
                        ReconnectToSpira(projectId);

                        ImportTestCase(streamWriter, mtmTestSuiteEntry.TestCase, remoteCustomList_Area, remoteCustomList_State);
                    }
                }
            }

            //Dynamic suites don't have any sub-suites to worry about
        }

        /// <summary>
        /// Imports the test cases under a suite recursively
        /// </summary>
        protected void ImportTestSuiteTestCases(StreamWriter streamWriter, int projectId, IStaticTestSuite mtmTestSuite, RemoteCustomList remoteCustomList_Area, RemoteCustomList remoteCustomList_State)
        {
            //Get the test cases to add to the test folder
            if (mtmTestSuite.TestCases != null && mtmTestSuite.TestCases.Count > 0)
            {
                streamWriter.WriteLine(String.Format("Importing {0} test cases in suite: {1}", mtmTestSuite.TestCases.Count, mtmTestSuite.Title));
                foreach (ITestSuiteEntry mtmTestSuiteEntry in mtmTestSuite.TestCases)
                {
                    if (mtmTestSuiteEntry.TestCase != null)
                    {
                        //Reconnect to Spira
                        ReconnectToSpira(projectId);

                        ImportTestCase(streamWriter, mtmTestSuiteEntry.TestCase, remoteCustomList_Area, remoteCustomList_State);
                    }
                }
            }

            //Now we need to see if there are any child test suites under this
            if (mtmTestSuite.SubSuites != null && mtmTestSuite.SubSuites.Count > 0)
            {
                foreach (var mtmChildTestSuite in mtmTestSuite.SubSuites)
                {
                    if (mtmChildTestSuite is IStaticTestSuite)
                    {
                        ImportTestSuiteTestCases(streamWriter, projectId, (IStaticTestSuite)mtmChildTestSuite, remoteCustomList_Area, remoteCustomList_State);
                    }
                    else if (mtmChildTestSuite is IDynamicTestSuite)
                    {
                        ImportTestSuiteTestCases(streamWriter, projectId, (IDynamicTestSuite)mtmChildTestSuite, remoteCustomList_Area, remoteCustomList_State);
                    }
                }
            }
        }

        /// <summary>
        /// Imports the static test suites, test sets and test case mappings
        /// </summary>
        protected void ImportTestSuitesAsTestSets(StreamWriter streamWriter, IStaticTestSuite mtmTestSuite, int? parentTestSetFolderId = null)
        {
            //Load the test set and capture the new id
            RemoteTestSet remoteTestSet = new RemoteTestSet();
            remoteTestSet.Name = mtmTestSuite.Title;
            remoteTestSet.Description = mtmTestSuite.Description;
            remoteTestSet.TestSetStatusId = ConvertTestSetStatus(mtmTestSuite.State);
            int newTestSetId = ImportFormHandle.SpiraImportProxy.TestSet_Create(remoteTestSet, parentTestSetFolderId).TestSetId.Value;
            if (!testSuiteTestSetMapping.ContainsKey(mtmTestSuite.Id))
            {
                testSuiteTestSetMapping.Add(mtmTestSuite.Id, newTestSetId);
            }
            streamWriter.WriteLine("Imported test suite: " + mtmTestSuite.Title);

            //Get the test cases to add to the test set
            if (mtmTestSuite.TestCases != null && mtmTestSuite.TestCases.Count > 0)
            {
                foreach (ITestSuiteEntry mtmTestSuiteEntry in mtmTestSuite.TestCases)
                {
                    if (mtmTestSuiteEntry.TestCase != null)
                    {
                        //Add the test case under the test set
                        if (this.testCaseMapping.ContainsKey(mtmTestSuiteEntry.TestCase.Id))
                        {
                            int testCaseId = this.testCaseMapping[mtmTestSuiteEntry.TestCase.Id];
                            RemoteTestSetTestCaseMapping remoteTestSetTestCaseMapping = new RemoteTestSetTestCaseMapping();
                            remoteTestSetTestCaseMapping.TestSetId = newTestSetId;
                            remoteTestSetTestCaseMapping.TestCaseId = testCaseId;
                            remoteTestSetTestCaseMapping = ImportFormHandle.SpiraImportProxy.TestSet_AddTestMapping(remoteTestSetTestCaseMapping, null, null)[0];
                            int newTestSetTestCaseId = remoteTestSetTestCaseMapping.TestSetTestCaseId;
                            if (!testSuiteEntryTestSetTestCaseIdMapping.ContainsKey(mtmTestSuiteEntry.Id))
                            {
                                testSuiteEntryTestSetTestCaseIdMapping.Add(mtmTestSuiteEntry.Id, newTestSetTestCaseId);
                            }

                        }
                    }
                }
            }

            //Now we need to see if there are any child test suites under this
            //SpiraTest doesn't let you nest test sets under other test sets, so we'll import them all under the main test plan entry
            if (mtmTestSuite.SubSuites != null && mtmTestSuite.SubSuites.Count > 0)
            {
                foreach (var mtmChildTestSuite in mtmTestSuite.SubSuites)
                {
                    if (mtmChildTestSuite is IStaticTestSuite)
                    {
                        ImportTestSuitesAsTestSets(streamWriter, (IStaticTestSuite)mtmChildTestSuite, parentTestSetFolderId);
                    }
                    else if (mtmChildTestSuite is IDynamicTestSuite)
                    {
                        ImportTestSuitesAsTestSets(streamWriter, (IDynamicTestSuite)mtmChildTestSuite, parentTestSetFolderId);
                    }
                }
            }
        }

        /// <summary>
        /// Imports the test suites, test sets and test case mappings
        /// </summary>
        protected void ImportTestSuitesAsTestSets(StreamWriter streamWriter, IDynamicTestSuite mtmTestSuite, int? parentTestSetFolderId = null)
        {
            //Load the test set and capture the new id
            RemoteTestSet remoteTestSet = new RemoteTestSet();
            remoteTestSet.Name = mtmTestSuite.Title;
            remoteTestSet.Description = mtmTestSuite.Description;
            remoteTestSet.TestSetStatusId = ConvertTestSetStatus(mtmTestSuite.State); ;
            int newTestSetId = ImportFormHandle.SpiraImportProxy.TestSet_Create(remoteTestSet, parentTestSetFolderId).TestSetId.Value;
            if (!testSuiteTestSetMapping.ContainsKey(mtmTestSuite.Id))
            {
                testSuiteTestSetMapping.Add(mtmTestSuite.Id, newTestSetId);
            }
            streamWriter.WriteLine("Imported test suite: " + mtmTestSuite.Title);

            if (mtmTestSuite.TestCases != null && mtmTestSuite.TestCases.Count > 0)
            {
                foreach (ITestSuiteEntry mtmTestSuiteEntry in mtmTestSuite.TestCases)
                {
                    if (mtmTestSuiteEntry.TestCase != null)
                    {
                        //Add the test case under the test set
                        if (this.testCaseMapping.ContainsKey(mtmTestSuiteEntry.TestCase.Id))
                        {
                            int testCaseId = this.testCaseMapping[mtmTestSuiteEntry.TestCase.Id];
                            RemoteTestSetTestCaseMapping remoteTestSetTestCaseMapping = new RemoteTestSetTestCaseMapping();
                            remoteTestSetTestCaseMapping.TestSetId = newTestSetId;
                            remoteTestSetTestCaseMapping.TestCaseId = testCaseId;
                            remoteTestSetTestCaseMapping = ImportFormHandle.SpiraImportProxy.TestSet_AddTestMapping(remoteTestSetTestCaseMapping, null, null)[0];
                            int newTestSetTestCaseId = remoteTestSetTestCaseMapping.TestSetTestCaseId;
                            if (!testSuiteEntryTestSetTestCaseIdMapping.ContainsKey(mtmTestSuiteEntry.Id))
                            {
                                testSuiteEntryTestSetTestCaseIdMapping.Add(mtmTestSuiteEntry.Id, newTestSetTestCaseId);
                            }
                        }
                    }
                }
            }

            //Dynamic suites don't have any sub-suites to worry about
        }

        /// <summary>
        /// Imports the static test suites, test cases and test steps
        /// </summary>
        protected void ImportTestSuitesAsTestCaseFolders(StreamWriter streamWriter, IStaticTestSuite mtmTestSuite, int? parentTestFolderId = null)
        {
            //Load the test folder and capture the new id
            RemoteTestCase remoteTestFolder = new RemoteTestCase();
            remoteTestFolder.Name = mtmTestSuite.Title;
            remoteTestFolder.Description = mtmTestSuite.Description;
            int newTestFolderId = ImportFormHandle.SpiraImportProxy.TestCase_CreateFolder(remoteTestFolder, parentTestFolderId).TestCaseId.Value;
            if (!testSuiteFolderMapping.ContainsKey(mtmTestSuite.Id))
            {
                testSuiteFolderMapping.Add(mtmTestSuite.Id, newTestFolderId);
            }
            streamWriter.WriteLine("Imported test suite: " + mtmTestSuite.Title);

            //Get the test cases to add to the mapping collection
            if (mtmTestSuite.TestCases != null && mtmTestSuite.TestCases.Count > 0)
            {
                foreach (ITestSuiteEntry mtmTestSuiteEntry in mtmTestSuite.TestCases)
                {
                    if (mtmTestSuiteEntry.TestCase != null)
                    {
                        this.testSuiteTestCaseMappings.Add(new MtmTestSuiteTestCaseMapping(mtmTestSuite.Id, mtmTestSuiteEntry.TestCase.Id));
                    }
                }
            }

            //Now we need to see if there are any child test suites under this
            if (mtmTestSuite.SubSuites != null && mtmTestSuite.SubSuites.Count > 0)
            {
                foreach (var mtmChildTestSuite in mtmTestSuite.SubSuites)
                {
                    if (mtmChildTestSuite is IStaticTestSuite)
                    {
                        ImportTestSuitesAsTestCaseFolders(streamWriter, (IStaticTestSuite)mtmChildTestSuite, newTestFolderId);
                    }
                    else if (mtmChildTestSuite is IDynamicTestSuite)
                    {
                        ImportTestSuitesAsTestCaseFolders(streamWriter, (IDynamicTestSuite)mtmChildTestSuite, newTestFolderId);
                    }
                }
            }
        }

        /// <summary>
        /// Imports the test suites, test cases and test steps
        /// </summary>
        protected void ImportTestSuitesAsTestCaseFolders(StreamWriter streamWriter, IDynamicTestSuite mtmTestSuite, int? parentTestFolderId = null)
        {
            //Load the test folder and capture the new id
            RemoteTestCase remoteTestFolder = new RemoteTestCase();
            remoteTestFolder.Name = mtmTestSuite.Title;
            remoteTestFolder.Description = mtmTestSuite.Description;
            int newTestFolderId = ImportFormHandle.SpiraImportProxy.TestCase_CreateFolder(remoteTestFolder, parentTestFolderId).TestCaseId.Value;
            if (!testSuiteFolderMapping.ContainsKey(mtmTestSuite.Id))
            {
                testSuiteFolderMapping.Add(mtmTestSuite.Id, newTestFolderId);
            }
            streamWriter.WriteLine("Imported test suite: " + mtmTestSuite.Title);

            if (mtmTestSuite.TestCases != null && mtmTestSuite.TestCases.Count > 0)
            {
                foreach (ITestSuiteEntry mtmTestSuiteEntry in mtmTestSuite.TestCases)
                {
                    if (mtmTestSuiteEntry.TestCase != null)
                    {
                        this.testSuiteTestCaseMappings.Add(new MtmTestSuiteTestCaseMapping(mtmTestSuite.Id, mtmTestSuiteEntry.TestCase.Id));
                    }
                }
            }

            //Dynamic suites don't have any sub-suites to worry about
        }

        /// <summary>
        /// This method is responsible for actually importing the data
        /// </summary>
        /// <param name="stateInfo">State information handle</param>
        /// <remarks>This runs in background thread to avoid freezing the progress form</remarks>
        protected void ImportData(object stateInfo)
        {            
            //First open up the textfile that we will log information to (used for debugging purposes)
            string debugFile = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\Spira_MTMImport.log";
            StreamWriter streamWriter = File.CreateText(debugFile);

            try
            {
                streamWriter.WriteLine("Starting import at: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString());

                //Instantiate the connection to MTM
                //Configure the network credentials - used for accessing the MsTfs API
                //If we have a domain provided, use a NetworkCredential, otherwise use a TFS credential
                //TODO: Migrate code to VS2012 and add support for TfsClientCredentials
                ICredentials tfsCredential = null;
                NetworkCredential networkCredential = null;
                /*
                TfsClientCredentials tfsCredential = null;
                if (String.IsNullOrWhiteSpace(this.txtDomain.Text))
                {
                    SimpleWebTokenCredential simpleWebTokenCredential = new SimpleWebTokenCredential(this.externalLogin, this.externalPassword);
                    tfsCredential = new TfsClientCredentials(simpleWebTokenCredential);
                    tfsCredential.AllowInteractive = false;
                }
                else
                {*/
                //Windows credentials
                networkCredential = new NetworkCredential(Properties.Settings.Default.MtmUserName, Properties.Settings.Default.MtmPassword, Properties.Settings.Default.MtmDomain);
                /*}*/

                //Create a new TFS 2012 project collection instance and WorkItemStore instance
                //This requires that the URI includes the collection name not just the server name
                Uri tfsUri = new Uri(Properties.Settings.Default.MtmUrl);
                TfsTeamProjectCollection tfsTeamProjectCollection;
                if (tfsCredential == null)
                {
                    tfsTeamProjectCollection = new TfsTeamProjectCollection(tfsUri, networkCredential);
                }
                else
                {
                    tfsTeamProjectCollection = new TfsTeamProjectCollection(tfsUri, tfsCredential);
                }

                //Authenticate
                tfsTeamProjectCollection.Authenticate();

                //Get access to the work item, group security and test manager server services
                WorkItemServer workItemServer = tfsTeamProjectCollection.GetService<WorkItemServer>();
                TestManagementService testManagementService = tfsTeamProjectCollection.GetService<TestManagementService>();
                WorkItemStore workItemStore = new WorkItemStore(tfsTeamProjectCollection);

                //Locate the MTM project and TFS project
                string projectName = Properties.Settings.Default.MtmProjectName;
                streamWriter.WriteLine("Connecting to TFS/MTM project: " + projectName);
                ITestManagementTeamProject mtmProject = testManagementService.GetTeamProject(projectName);
                if (mtmProject == null)
                {
                    throw new ApplicationException(String.Format("Unable to access MTM project '{0}'", projectName));
                }
                Project workItemProject = workItemStore.Projects[projectName];
                if (workItemProject == null)
                {
                    throw new ApplicationException(String.Format("Unable to access TFS project '{0}'", projectName));
                }

                //Retrieve the TFS user list
                List<TeamFoundationIdentity> tfsUsers = Utils.ListContributors(tfsTeamProjectCollection, projectName);

                //Reconnect to Spira
                ImportFormHandle.SpiraImportProxy.Connection_Authenticate2(Properties.Settings.Default.SpiraUserName, Properties.Settings.Default.SpiraPassword, ImportForm.API_PLUGIN_NAME);

                //1) Create a new project
                RemoteProject remoteProject = new RemoteProject();
                remoteProject.Name = projectName;
                remoteProject.Description = projectName;
                remoteProject.Active = true;
                remoteProject = ImportFormHandle.SpiraImportProxy.Project_Create(remoteProject, null);
                streamWriter.WriteLine("New Project '" + projectName + "' Created");
                int projectId = remoteProject.ProjectId.Value;

                //Create a new list custom property for Area and State
                //Area
                RemoteCustomList remoteCustomList_Area = new RemoteCustomList();
                remoteCustomList_Area.Name = "Areas";
                remoteCustomList_Area.Active = true;
                remoteCustomList_Area = ImportFormHandle.SpiraImportProxy.CustomProperty_AddCustomList(remoteCustomList_Area);
                RemoteCustomProperty remoteCustomProperty_Area = new RemoteCustomProperty();
                remoteCustomProperty_Area.Name = "Area";
                remoteCustomProperty_Area.PropertyNumber = 1;
                remoteCustomProperty_Area.ProjectId = projectId;
                remoteCustomProperty_Area.CustomPropertyTypeId = 6;  //List
                remoteCustomProperty_Area.ArtifactTypeId = (int)ArtifactType.TestCase;
                ImportFormHandle.SpiraImportProxy.CustomProperty_AddDefinition(remoteCustomProperty_Area, remoteCustomList_Area.CustomPropertyListId);

                //State
                RemoteCustomList remoteCustomList_State = new RemoteCustomList();
                remoteCustomList_State.Name = "States";
                remoteCustomList_State.Active = true;
                remoteCustomList_State = ImportFormHandle.SpiraImportProxy.CustomProperty_AddCustomList(remoteCustomList_State);
                RemoteCustomProperty remoteCustomProperty_State = new RemoteCustomProperty();
                remoteCustomProperty_State.Name = "State";
                remoteCustomProperty_State.PropertyNumber = 2;
                remoteCustomProperty_State.ProjectId = projectId;
                remoteCustomProperty_State.CustomPropertyTypeId = 6;  //List
                remoteCustomProperty_State.ArtifactTypeId = (int)ArtifactType.TestCase;
                ImportFormHandle.SpiraImportProxy.CustomProperty_AddDefinition(remoteCustomProperty_State, remoteCustomList_State.CustomPropertyListId);

                //2) Get the users and import - if we don't want to import user, map all MTM users to single SpiraId
                int userId = -1;
                if (!this.MainFormHandle.chkImportUsers.Checked || tfsUsers == null || tfsUsers.Count == 0)
                {
                    RemoteUser remoteUser = new RemoteUser();
                    remoteUser.FirstName = "Microsoft";
                    remoteUser.LastName = "TestManager";
                    remoteUser.UserName = "mstestmanager";
                    remoteUser.EmailAddress = "mstestmanager@mycompany.com";
                    remoteUser.Active = true;
                    remoteUser.Admin = false;
                    userId = ImportFormHandle.SpiraImportProxy.User_Create(remoteUser, DEFAULT_PASSWORD, "What TFS URL was this project imported from?", Properties.Settings.Default.MtmUrl,  5).UserId.Value;
                    streamWriter.WriteLine("No MTM users found, so adding single MTM User in SpiraTest:");
                }
                else
                {
                    foreach (TeamFoundationIdentity tfsUser in tfsUsers)
                    {
                        if (!this.usersMapping.ContainsKey(tfsUser.UniqueUserId))
                        {
                            //Extract the user data
                            string userName = tfsUser.UniqueName;

                            //Get the first & last names
                            string[] names = tfsUser.DisplayName.Split(' ');
                            string firstName = names[0];
                            string lastName = "-";
                            if (names.Length > 1)
                            {
                                lastName = names[1];
                            }

                            //Default to observer role for all imports, for security reasons
                            RemoteUser remoteUser = new RemoteUser();
                            remoteUser.FirstName = firstName;
                            remoteUser.LastName = lastName;
                            remoteUser.UserName = tfsUser.UniqueName.Replace(" ", "-");
                            remoteUser.EmailAddress = remoteUser.UserName + "@mycompany.com";
                            remoteUser.Active = tfsUser.IsActive;
                            remoteUser.Admin = false;
                            userId = ImportFormHandle.SpiraImportProxy.User_Create(remoteUser, DEFAULT_PASSWORD, "What TFS URL was this project imported from?", Properties.Settings.Default.MtmUrl, 5).UserId.Value;
                            streamWriter.WriteLine("Added user: '" + tfsUser.UniqueName);

                            //Add the mapping to the hashtable for use later on
                            this.usersMapping.Add(tfsUser.UniqueUserId, userId);
                        }
                    }
                }

                //**** Show that we've imported users ****
                if (this.MainFormHandle.chkImportUsers.Checked)
                {
                    streamWriter.WriteLine("Users Imported");
                    this.ProgressForm_OnProgressUpdate(1);
                }

                //3) Get the releases (iterations) and areas.
                //Iterations
                foreach (Node iterationNode in workItemProject.IterationRootNodes)
                {
                    ImportIterationNode(iterationNode, streamWriter);
                }
                foreach (Node areaNode in workItemProject.AreaRootNodes)
                {
                    ImportAreaNode(remoteCustomList_Area, areaNode, streamWriter);
                }

                //**** Show that we've imported releases ****
                streamWriter.WriteLine("Releases Imported");
                this.ProgressForm_OnProgressUpdate(2);

                //4) Get the test cases and import

                //Add a top-level test folder for storing Shared Test Cases
                RemoteTestCase remoteSharedTestCasesFolder = new RemoteTestCase();
                remoteSharedTestCasesFolder.Name = "Shared Test Cases";
                remoteSharedTestCasesFolder.Active = true;
                remoteSharedTestCasesFolder.Description = "This folder contains all the imported shared test cases from MS Test Manager";
                int newSharedTestCasesFolderId = ImportFormHandle.SpiraImportProxy.TestCase_CreateFolder(remoteSharedTestCasesFolder, null).TestCaseId.Value;

                /*
                //Add a top-level test folder for storing Test Cases that are not part of a test plan / test suite
                RemoteTestCase remoteUnmappedTestCasesFolder = new RemoteTestCase();
                remoteUnmappedTestCasesFolder.Name = "Unmapped Test Cases";
                remoteUnmappedTestCasesFolder.Active = true;
                remoteUnmappedTestCasesFolder.Description = "This folder contains all the imported unmapped test cases from MS Test Manager";
                int newUnmappedTestCasesFolderId = ImportFormHandle.SpiraImportProxy.TestCase_CreateFolder(remoteUnmappedTestCasesFolder, null).TestCaseId.Value;
                */

                //First we need to build out the folders from the list of 'test plans'
                //The test plans are the top-level folders for test cases
                ITestPlanCollection plans = mtmProject.TestPlans.Query("Select * From TestPlan");
                streamWriter.WriteLine(String.Format("Found {0} test plans", plans.Count));
                foreach (ITestPlan mtmTestPlan in plans)
                {
                    //See if we are only importing a single test plan or not
                    if (Properties.Settings.Default.MtmTestPlanName == Utils.ALL_TEST_PLANS || mtmTestPlan.Name == Properties.Settings.Default.MtmTestPlanName)
                    {
                        //Load the test folder and capture the new id
                        RemoteTestCase remoteTestFolder = new RemoteTestCase();
                        remoteTestFolder.Name = mtmTestPlan.Name;
                        remoteTestFolder.Description = mtmTestPlan.Description;
                        remoteTestFolder.Active = true;
                        if (mtmTestPlan.Owner != null && usersMapping.ContainsKey(mtmTestPlan.Owner.UniqueUserId))
                        {
                            remoteTestFolder.OwnerId = usersMapping[mtmTestPlan.Owner.UniqueUserId];
                        }
                        int newTestFolderId = ImportFormHandle.SpiraImportProxy.TestCase_CreateFolder(remoteTestFolder, null).TestCaseId.Value;
                        if (!testPlanFolderMapping.ContainsKey(mtmTestPlan.Id))
                        {
                            testPlanFolderMapping.Add(mtmTestPlan.Id, newTestFolderId);
                        }
                        streamWriter.WriteLine("Imported test plan: " + mtmTestPlan.Name);

                        //Now we need to see if there are any test suites under the test plan
                        if (mtmTestPlan.RootSuite != null)
                        {
                            if (mtmTestPlan.RootSuite is IStaticTestSuite)
                            {
                                IStaticTestSuite mtmStaticTestSuite = (IStaticTestSuite)mtmTestPlan.RootSuite;
                                ImportTestSuitesAsTestCaseFolders(streamWriter, mtmStaticTestSuite, newTestFolderId);
                            }
                            foreach (var mtmTestSuite in mtmTestPlan.RootSuite.Entries)
                            {
                                if (mtmTestSuite is IStaticTestSuite)
                                {
                                    IStaticTestSuite mtmStaticTestSuite = (IStaticTestSuite)mtmTestSuite;
                                    ImportTestSuitesAsTestCaseFolders(streamWriter, mtmStaticTestSuite, newTestFolderId);
                                }
                                if (mtmTestSuite is IDynamicTestSuite)
                                {
                                    IDynamicTestSuite mtmDynamicTestSuite = (IDynamicTestSuite)mtmTestSuite;
                                    ImportTestSuitesAsTestCaseFolders(streamWriter, mtmDynamicTestSuite, newTestFolderId);
                                }
                            }
                        }
                        streamWriter.WriteLine("Imported test suites");
                    }
                }

                //First we import the shared steps (which don't belong to a suite)
                //We have to do this for each Area and Iteration path separately because otherwise we were getting OutOfMemory
                //errors from the TFS object model

                //Root Iteration (same as project name)
                string rootIterationPath = projectName;

                //Root Area (same as project name)
                string rootAreaPath = projectName;
                ImportedSharedSteps(streamWriter, projectId, rootIterationPath, rootAreaPath, mtmProject, remoteCustomList_Area, remoteCustomList_State, newSharedTestCasesFolderId);

                //Areas
                foreach (KeyValuePair<string, int> kvp2 in this.areaCustomPropertyValueMapping)
                {
                    string areaPath = kvp2.Key;
                    ImportedSharedSteps(streamWriter, projectId, rootIterationPath, areaPath, mtmProject, remoteCustomList_Area, remoteCustomList_State, newSharedTestCasesFolderId);
                }

                //Iterations
                foreach (KeyValuePair<string, int> kvp in this.releasePathMapping)
                {
                    string iterationPath = kvp.Key;

                    //Root Area (same as project name)
                    ImportedSharedSteps(streamWriter, projectId, iterationPath, rootAreaPath, mtmProject, remoteCustomList_Area, remoteCustomList_State, newSharedTestCasesFolderId);
                    
                    //Areas
                    foreach (KeyValuePair<string, int> kvp2 in this.areaCustomPropertyValueMapping)
                    {
                        string areaPath = kvp2.Key;
                        ImportedSharedSteps(streamWriter, projectId, iterationPath, areaPath, mtmProject, remoteCustomList_Area, remoteCustomList_State, newSharedTestCasesFolderId);
                    }
                }

                //Now the test steps of the 'Shared Test Cases' (which have their own steps)

                //Root Area (same as project name)
                rootAreaPath = projectName;
                ImportSharedStepsSteps(streamWriter, projectId, rootIterationPath, rootAreaPath, mtmProject);

                //Areas
                foreach (KeyValuePair<string, int> kvp2 in this.areaCustomPropertyValueMapping)
                {
                    string areaPath = kvp2.Key;
                    ImportSharedStepsSteps(streamWriter, projectId, rootIterationPath, areaPath, mtmProject);
                }

                //Iterations
                foreach (KeyValuePair<string, int> kvp in this.releasePathMapping)
                {
                    string iterationPath = kvp.Key;

                    //Areas
                    foreach (KeyValuePair<string, int> kvp2 in this.areaCustomPropertyValueMapping)
                    {
                        //Root Area (same as project name)
                        ImportSharedStepsSteps(streamWriter, projectId, iterationPath, rootAreaPath, mtmProject);

                        string areaPath = kvp2.Key;
                        ImportSharedStepsSteps(streamWriter, projectId, iterationPath, areaPath, mtmProject);
                    }
                }

                //Now we import the test cases themselves, we have to do a second separate pass for test steps
                //to avoid the issue of test cases that link to other test cases
                //We also have to do it by suite because the MTM API was throwing OutOfMemory exceptions
                //when you try and retrieve all the test cases in the project in one query
                foreach (ITestPlan mtmTestPlan in plans)
                {
                    //See if we are only importing a single test plan or not
                    if (Properties.Settings.Default.MtmTestPlanName == Utils.ALL_TEST_PLANS || mtmTestPlan.Name == Properties.Settings.Default.MtmTestPlanName)
                    {
                        //Now we need to see if there are any test suites under the test plan
                        if (mtmTestPlan.RootSuite != null)
                        {
                            if (mtmTestPlan.RootSuite is IStaticTestSuite)
                            {
                                IStaticTestSuite mtmStaticTestSuite = (IStaticTestSuite)mtmTestPlan.RootSuite;
                                ImportTestSuiteTestCases(streamWriter, projectId, mtmStaticTestSuite, remoteCustomList_Area, remoteCustomList_State);
                            }
                            foreach (var mtmTestSuite in mtmTestPlan.RootSuite.Entries)
                            {
                                if (mtmTestSuite is IStaticTestSuite)
                                {
                                    IStaticTestSuite mtmStaticTestSuite = (IStaticTestSuite)mtmTestSuite;
                                    ImportTestSuiteTestCases(streamWriter, projectId, mtmStaticTestSuite, remoteCustomList_Area, remoteCustomList_State);
                                }
                                if (mtmTestSuite is IDynamicTestSuite)
                                {
                                    IDynamicTestSuite mtmDynamicTestSuite = (IDynamicTestSuite)mtmTestSuite;
                                    ImportTestSuiteTestCases(streamWriter, projectId, mtmDynamicTestSuite, remoteCustomList_Area, remoteCustomList_State);
                                }
                            }
                        }
                        streamWriter.WriteLine("Imported test cases in suite");
                    }
                }

                //**** Show that we've imported test cases, test steps ****
                streamWriter.WriteLine("Test Cases and Test Steps Imported");
                this.ProgressForm_OnProgressUpdate(3);

                //5) Now re-import the test plans and test suites as test sets
                //First we need to build out the folders from the list of 'test plans'
                //The test plans are the top-level folders for test cases
                if (Properties.Settings.Default.TestSets)
                {
                    streamWriter.WriteLine("** Importing Test Sets into SpiraTest **");
                    plans = mtmProject.TestPlans.Query("Select * From TestPlan");
                    streamWriter.WriteLine(String.Format("Found {0} test plans", plans.Count));
                    foreach (ITestPlan mtmTestPlan in plans)
                    {
                        //See if we are only importing a single test plan or not
                        if (Properties.Settings.Default.MtmTestPlanName == Utils.ALL_TEST_PLANS || mtmTestPlan.Name == Properties.Settings.Default.MtmTestPlanName)
                        {
                            //Load the test folder and capture the new id
                            RemoteTestSet remoteTestSetFolder = new RemoteTestSet();
                            remoteTestSetFolder.Name = mtmTestPlan.Name;
                            remoteTestSetFolder.Description = mtmTestPlan.Description;
                            if (mtmTestPlan.Owner != null && usersMapping.ContainsKey(mtmTestPlan.Owner.UniqueUserId))
                            {
                                remoteTestSetFolder.OwnerId = usersMapping[mtmTestPlan.Owner.UniqueUserId];
                            }
                            int newTestSetFolderId = ImportFormHandle.SpiraImportProxy.TestSet_CreateFolder(remoteTestSetFolder, null).TestSetId.Value;
                            if (!testPlanFolderMapping.ContainsKey(mtmTestPlan.Id))
                            {
                                testPlanTestSetFolderMapping.Add(mtmTestPlan.Id, newTestSetFolderId);
                            }
                            streamWriter.WriteLine("Imported test plan: " + mtmTestPlan.Name);

                            //Now we need to see if there are any test suites under the test plan
                            if (mtmTestPlan.RootSuite != null)
                            {
                                if (mtmTestPlan.RootSuite is IStaticTestSuite)
                                {
                                    IStaticTestSuite mtmStaticTestSuite = (IStaticTestSuite)mtmTestPlan.RootSuite;
                                    ImportTestSuitesAsTestSets(streamWriter, mtmStaticTestSuite, newTestSetFolderId);
                                }
                                foreach (var mtmTestSuite in mtmTestPlan.RootSuite.Entries)
                                {
                                    if (mtmTestSuite is IStaticTestSuite)
                                    {
                                        IStaticTestSuite mtmStaticTestSuite = (IStaticTestSuite)mtmTestSuite;
                                        ImportTestSuitesAsTestSets(streamWriter, mtmStaticTestSuite, newTestSetFolderId);
                                    }
                                    if (mtmTestSuite is IDynamicTestSuite)
                                    {
                                        IDynamicTestSuite mtmDynamicTestSuite = (IDynamicTestSuite)mtmTestSuite;
                                        ImportTestSuitesAsTestSets(streamWriter, mtmDynamicTestSuite, newTestSetFolderId);
                                    }
                                }
                            }
                            streamWriter.WriteLine("Imported test suites");
                        }
                    }

                    //**** Show that we've imported test sets ****
                    streamWriter.WriteLine("Test Sets Imported");
                    this.ProgressForm_OnProgressUpdate(4);
                }

                //7) Get the test runs (including run steps) and import
                if (this.MainFormHandle.chkImportTestRuns.Checked)
                {
                    streamWriter.WriteLine("Importing Test Runs");
                    IEnumerable<ITestRun> mtmTestRuns = mtmProject.TestRuns.Query("Select * From TestRun");
                    if (mtmTestRuns == null)
                    {
                        streamWriter.WriteLine("A null run list was returned, so skipping test run import.");
                    }
                    else
                    {
                        streamWriter.WriteLine(String.Format("Found {0} test runs", mtmTestRuns.Count()));
                        foreach (ITestRun mtmTestRun in mtmTestRuns)
                        {
                            //Extract the test run info
                            int runId = mtmTestRun.Id;
                            int testPlanId = mtmTestRun.TestPlanId;
                            string iterationPath = mtmTestRun.Iteration;

                            streamWriter.WriteLine("Importing Test Run: " + mtmTestRun.Title + " (iteration path=" + iterationPath + ")");

                            //What SpiraTest considers 'test runs' are really the results of the test run
                            foreach (ITestCaseResult mtmTestResult in mtmTestRun.QueryResults())
                            {
                                int testResultId = mtmTestResult.Id.TestResultId;
                                int testId = mtmTestResult.TestCaseId;

                                streamWriter.WriteLine("Importing Test Run - Result ID: " + mtmTestResult.Id);

                                //MTM doesn't have test run steps, so we need to just create a single test run step that
                                //points to the first test run step

                                //Locate the test and get the corresponding SpiraTest test case id
                                if (this.testCaseMapping.ContainsKey(testId))
                                {
                                    int testCaseId = (int)this.testCaseMapping[testId];

                                    //Now create a new test run shell from this test case
                                    RemoteManualTestRun[] remoteTestRuns = ImportFormHandle.SpiraImportProxy.TestRun_CreateFromTestCases(new int[] { testCaseId }, null);
                                    if (remoteTestRuns.Length > 0)
                                    {
                                        //Update the test run information
                                        RemoteManualTestRun remoteTestRun = remoteTestRuns[0];
                                        remoteTestRun.ExecutionStatusId = ConvertExecutionStatus(mtmTestResult.Outcome);
                                        DateTime startDate = mtmTestResult.DateStarted;
                                        //Make sure the dates are greater than the Spira minimum
                                        if (startDate.Year < 1900)
                                        {
                                            startDate = DateTime.UtcNow.AddMinutes(-1);
                                        }
                                        DateTime endDate = mtmTestResult.DateCompleted;
                                        if (endDate.Year < 1900)
                                        {
                                            endDate = DateTime.UtcNow;
                                        }
                                        remoteTestRun.StartDate = startDate;
                                        remoteTestRun.EndDate = endDate;
                                        remoteTestRun.Name = mtmTestResult.TestCaseTitle;
                                        if (!String.IsNullOrEmpty(iterationPath) && releasePathMapping.ContainsKey(iterationPath))
                                        {
                                            remoteTestRun.ReleaseId = releasePathMapping[iterationPath];
                                        }
                                        //TODO: Implement once we determine the test suite entry of the test run
                                        //if (this.testSuiteEntryTestSetTestCaseIdMapping.ContainsKey(XXXXX))
                                        //{
                                        //    int testSetTestCaseId = this.testSuiteEntryTestSetTestCaseIdMapping[XXXXX];
                                        //    remoteTestRun.TestSetTestCaseId = testSetTestCaseId;
                                        //}
                                        if (mtmTestResult.Owner != null && usersMapping.ContainsKey(mtmTestResult.Owner.UniqueUserId))
                                        {
                                            remoteTestRun.TesterId = usersMapping[mtmTestResult.Owner.UniqueUserId];
                                        }

                                        //Populate the test steps
                                        if (remoteTestRun.TestRunSteps != null)
                                        {
                                            foreach (RemoteTestRunStep remoteTestRunStep in remoteTestRun.TestRunSteps)
                                            {
                                                remoteTestRun.TestRunSteps[0].ActualResult = mtmTestResult.ErrorMessage;
                                                remoteTestRun.TestRunSteps[0].ExecutionStatusId = ConvertExecutionStatus(mtmTestResult.Outcome);
                                            }
                                        }

                                        //Finally save this test run and get the updated list back
                                        try
                                        {
                                            DateTime endDate2 = mtmTestResult.DateCompleted;
                                            if (endDate2.Year < 1900)
                                            {
                                                endDate2 = DateTime.UtcNow;
                                            }
                                            remoteTestRuns = ImportFormHandle.SpiraImportProxy.TestRun_Save(remoteTestRuns, endDate2);
                                            if (remoteTestRuns.Length > 0 && remoteTestRun.TestRunSteps != null)
                                            {
                                                remoteTestRun = remoteTestRuns[0];

                                                //Add attachments if requested
                                                if (this.MainFormHandle.chkImportAttachments.Checked)
                                                {
                                                    try
                                                    {
                                                        ImportAttachments(streamWriter, mtmTestResult.Attachments, remoteTestRun.TestRunId.Value, ArtifactType.TestRun);
                                                    }
                                                    catch (Exception exception)
                                                    {
                                                        streamWriter.WriteLine("Warning: Unable to import attachments for run " + runId + " (" + exception.Message + ")");
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception exception)
                                        {
                                            //If we can't save the test run, proceed and log error
                                            streamWriter.WriteLine("*WARNING*: Unable to save Test Run " + runId + " in SpiraTest - continuing. Error message = '" + exception.Message + "' (" + exception.StackTrace + ")");
                                        }

                                        //Add this test case to the release as well, unless already added
                                        if (remoteTestRun.ReleaseId.HasValue)
                                        {
                                            string releaseAndTestCase = remoteTestRun.ReleaseId.Value + ":" + testCaseId;
                                            streamWriter.WriteLine(String.Format("Adding Test Case TC{0} to Release RL{1} in SpiraTest.", testCaseId, remoteTestRun.ReleaseId.Value));

                                            RemoteReleaseTestCaseMapping newMapping = new RemoteReleaseTestCaseMapping();
                                            newMapping.ReleaseId = remoteTestRun.ReleaseId.Value;
                                            newMapping.TestCaseId = testCaseId;
                                            ImportFormHandle.SpiraImportProxy.Release_AddTestMapping(newMapping);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //**** Show that we've imported test runs ****
                    streamWriter.WriteLine("Test Runs Imported");
                    this.ProgressForm_OnProgressUpdate(5);
                }

                //8) Get the incidents and import
                /*
                if (this.MainFormHandle.chkImportDefects.Checked)
                {
                    streamWriter.WriteLine("Importing Defects");
                    HP.QUalityCenter.BugFactory bugFactory = (HP.QUalityCenter.BugFactory)MainFormHandle.TdConnection.BugFactory;
                    HP.QUalityCenter.TDFilter tdFilter = (HP.QUalityCenter.TDFilter)bugFactory.Filter;
                    tdFilter.set_Order("bg_bug_id", 1);
                    HP.QUalityCenter.List bugList;
                    try
                    {
                        bugList = (HP.QUalityCenter.List)tdFilter.NewList();
                    }
                    catch (Exception)
                    {
                        //If there is an error try reconnecting
                        MainFormHandle.TryReconnect();
                        bugFactory = (HP.QUalityCenter.BugFactory)MainFormHandle.TdConnection.BugFactory;
                        tdFilter = (HP.QUalityCenter.TDFilter)bugFactory.Filter;
                        bugList = (HP.QUalityCenter.List)tdFilter.NewList();
                    }
                    if (bugList == null)
                    {
                        streamWriter.WriteLine("A null defect list was returned, so skipping defect import.");
                    }
                    else
                    {
                        foreach (HP.QUalityCenter.Bug bugObject in bugList)
                        {
                            bool objectExists = true;
                            int bugId = -1;
                            try
                            {
                                //Extract the incident/bug info
                                bugId = (int)bugObject.ID;
                                name = MakeXmlSafe(bugObject.Summary);
                            }
                            catch (Exception)
                            {
                                //Sometimes the underlying object doesn't exist, even though an ID is available
                                objectExists = false;
                                streamWriter.WriteLine("Skipping an empty defect record.");
                            }
                            if (objectExists)
                            {
                                streamWriter.WriteLine("Importing Defect: " + bugId.ToString());
                                string description = "Migrated from QualityCenter";
                                if (bugObject["bg_description"] != null)
                                {
                                    description = MakeXmlSafe(bugObject["bg_description"].ToString());
                                }
                                string resolution = "";
                                if (bugObject["bg_dev_comments"] != null)
                                {
                                    resolution = MakeXmlSafe(bugObject["bg_dev_comments"].ToString());
                                }
                                string bugSeverity = "";
                                if (bugObject["bg_severity"] != null)
                                {
                                    bugSeverity = bugObject["bg_severity"].ToString();
                                }
                                string bugPriority = bugObject.Priority;
                                string bugStatus = bugObject.Status;
                                string bugReproducibleYn = "N";
                                if (bugObject["bg_reproducible"] != null)
                                {
                                    bugReproducibleYn = bugObject["bg_reproducible"].ToString();
                                }

                                string openerName = bugObject.DetectedBy;
                                string ownerName = bugObject.AssignedTo;
                                Nullable<DateTime> closedDate = null;
                                if (bugObject["bg_closing_date"] != null)
                                {
                                    closedDate = DateTime.Parse(bugObject["bg_closing_date"].ToString());
                                }
                                Nullable<DateTime> creationDate = null;
                                if (bugObject["bg_detection_date"] != null)
                                {
                                    creationDate = DateTime.Parse(bugObject["bg_detection_date"].ToString());
                                }
                                //If we have a mapped test run step, lookup the SpiraTest id from the mapping hashtable
                                Nullable<int> testRunStepId = null;
                                if (this.MainFormHandle.chkImportTestRuns.Checked)
                                {
                                    testRunStepId = GetTestRunStepIdForBug(bugObject);
                                }

                                //Lookup the opener from the user mapping hashtable
                                Nullable<int> openerId = null;
                                if (this.usersMapping.ContainsKey(openerName))
                                {
                                    openerId = (int)this.usersMapping[openerName];
                                }
                                //Lookup the owner from the user mapping hashtable
                                Nullable<int> ownerId = null;
                                if (ownerName != null)
                                {
                                    if (this.usersMapping.ContainsKey(ownerName))
                                    {
                                        ownerId = (int)this.usersMapping[ownerName];
                                    }
                                }

                                //Convert the priority
                                Nullable<int> priorityId = ConvertPriority(bugPriority, ImportFormHandle.SpiraImportProxy);

                                //Convert the severity
                                Nullable<int> severityId = ConvertSeverity(bugSeverity, ImportFormHandle.SpiraImportProxy);

                                //Convert the status and type
                                Nullable<int> statusId = ConvertStatus(bugStatus, ImportFormHandle.SpiraImportProxy);

                                //Convert the various efforts from days to minutes
                                Nullable<int> estimatedEffort = null;
                                if (bugObject["bg_estimated_fix_time"] != null)
                                {
                                    //Assume 8 hour day
                                    estimatedEffort = ((int)bugObject["bg_estimated_fix_time"]) * 8 * 60;
                                }
                                Nullable<int> actualEffort = null;
                                if (bugObject["bg_actual_fix_time"] != null)
                                {
                                    //Assume 8 hour day
                                    actualEffort = ((int)bugObject["bg_actual_fix_time"]) * 8 * 60;
                                }

                                //QualityCenter doesn't have the concept of a defect type, so pass NULL to use the project's default

                                //Load the incident and capture the new id
                                Nullable<int> newIncidentId = null;
                                try
                                {
                                    //Load the incident data
                                    RemoteIncident remoteIncident = new RemoteIncident();
                                    remoteIncident.PriorityId = priorityId;
                                    remoteIncident.SeverityId = severityId;
                                    remoteIncident.IncidentStatusId = statusId;
                                    remoteIncident.Name = name;
                                    remoteIncident.Description = MakeXmlSafe(description);
                                    remoteIncident.TestRunStepId = testRunStepId;
                                    remoteIncident.OpenerId = openerId;
                                    remoteIncident.OwnerId = ownerId;
                                    remoteIncident.ClosedDate = closedDate;
                                    remoteIncident.CreationDate = creationDate;
                                    remoteIncident.EstimatedEffort = estimatedEffort;
                                    remoteIncident.ActualEffort = actualEffort;

                                    //Load any custom properties - QC stores them all as text
                                    //The text-10 custom property stores the QC bug ID
                                    remoteIncident.Text01 = MakeXmlSafe(bugObject["bg_user_01"]);
                                    remoteIncident.Text02 = MakeXmlSafe(bugObject["bg_user_02"]);
                                    remoteIncident.Text03 = MakeXmlSafe(bugObject["bg_user_03"]);
                                    remoteIncident.Text04 = MakeXmlSafe(bugObject["bg_user_04"]);
                                    remoteIncident.Text05 = MakeXmlSafe(bugObject["bg_user_05"]);
                                    remoteIncident.Text06 = MakeXmlSafe(bugObject["bg_user_06"]);
                                    remoteIncident.Text07 = MakeXmlSafe(bugObject["bg_user_07"]);
                                    remoteIncident.Text08 = MakeXmlSafe(bugObject["bg_user_08"]);
                                    remoteIncident.Text09 = MakeXmlSafe(bugObject["bg_user_09"]);
                                    remoteIncident.Text10 = "QC-" + bugId;
                                    newIncidentId = ImportFormHandle.SpiraImportProxy.Incident_Create(remoteIncident).IncidentId;
                                }
                                catch (Exception exception)
                                {
                                    //If we have an error, log it and continue
                                    streamWriter.WriteLine("Error adding defect " + bugId + " to SpiraTest (" + exception.Message + ")");
                                }

                                if (newIncidentId.HasValue)
                                {
                                    //Now add the resolution
                                    RemoteIncidentResolution remoteIncidentResolution = new RemoteIncidentResolution();
                                    remoteIncidentResolution.IncidentId = newIncidentId.Value;
                                    remoteIncidentResolution.Resolution = MakeXmlSafe(resolution);
                                    remoteIncidentResolution.CreationDate = (creationDate.HasValue) ? creationDate.Value : DateTime.Now;
                                    ImportFormHandle.SpiraImportProxy.Incident_AddResolutions(new RemoteIncidentResolution[] { remoteIncidentResolution });

                                    //Add attachments if requested
                                    if (this.MainFormHandle.chkImportAttachments.Checked)
                                    {
                                        try
                                        {
                                            ImportAttachments(streamWriter, "BUG", bugId, bugObject.Attachments, newIncidentId.Value, ArtifactType.Incident);
                                        }
                                        catch (Exception exception)
                                        {
                                            streamWriter.WriteLine("Warning: Unable to import attachments for bug " + bugId + " (" + exception.Message + ")");
                                        }
                                    }

                                    //Add to the incident mapping list
                                    incidentMapping.Add(bugId, newIncidentId.Value);
                                }
                            }
                        }
                    }

                    //**** Show that we've imported incidents ****
                    streamWriter.WriteLine("Defects/Incidents Imported");
                    this.ProgressForm_OnProgressUpdate(6);
                }*/

                //**** Mark the form as finished ****
                streamWriter.WriteLine("Import completed at: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString());
                this.ProgressForm_OnFinish();

                //Close the debugging file
                streamWriter.Close();
            }
            catch (Exception exception)
            {
                //Log the error
                streamWriter.WriteLine("*ERROR* Occurred during Import: '" + exception.Message + "' at " + exception.Source + " (" + exception.StackTrace + ")");
                streamWriter.Close();

                //Display the exception message
                ProgressForm_OnError(exception);
            }
        }

        /// <summary>
        /// Imports a list of attachments for an artifact from MTM
        /// </summary>
        private void ImportAttachments(StreamWriter streamWriter, IAttachmentCollection mtmAttachments, int artifactId, ArtifactType artifactType)
        {
            //Make sure we have a populated attachment factory
            if (mtmAttachments == null || mtmAttachments.Count < 1)
            {
                return;
            }
            streamWriter.WriteLine("Importing " + mtmAttachments.Count + " Attachments for " + artifactType + " with ID=" + artifactId);

            try
            {
                //Loop through the attachments
                foreach (ITestAttachment mtmAttachment in mtmAttachments)
                {
                    streamWriter.WriteLine("Importing Attachment: " + mtmAttachment.Name);
                    //Download the attachment
                    long length = mtmAttachment.Length;
                    byte[] attachmentData = new byte[length];
                    mtmAttachment.DownloadToArray(attachmentData, 0);

                    //Import the attachment
                    RemoteDocument remoteDocument = new RemoteDocument();
                    remoteDocument.FilenameOrUrl = mtmAttachment.Name;
                    remoteDocument.Description = MakeXmlSafe(mtmAttachment.Comment);
                    remoteDocument.ArtifactId = artifactId;
                    remoteDocument.ArtifactTypeId = (int)artifactType;
                    ImportFormHandle.SpiraImportProxy.Document_AddFile(remoteDocument, attachmentData);
                }
            }
            catch (Exception exception)
            {
                //Log and continue
                streamWriter.WriteLine("Warning: Unable to import attachments for " + artifactType + " artifact id '" + artifactId + "' - continuing with import (" + exception.Message + ").");
            }
        }

        /// <summary>
        /// Converts a MTM test set status for use in SpiraTest
        /// </summary>
        /// <param name="testSuiteState">The MTM test suite state</param>
        /// <returns>The SpiraTest test set status</returns>
        protected int ConvertTestSetStatus(TestSuiteState testSuiteState)
        {
            int statusId = (int)Utils.TestSetStatus.NotStarted; //Default to not started (no nulls allowed)
            switch (testSuiteState)
            {
                case TestSuiteState.None:
                case TestSuiteState.InPlanning:
                    //Not Started
                    statusId = (int)Utils.TestSetStatus.NotStarted;
                    break;

                case TestSuiteState.InProgress:
                    //In Progress
                    statusId = (int)Utils.TestSetStatus.InProgress;
                    break;

                case TestSuiteState.Completed:
                    //Completed
                    statusId = (int)Utils.TestSetStatus.Completed;
                    break;
            }
            return statusId;
        }

        /// <summary>
        /// Converts @parameter MTM format to ${parameter} spira format
        /// </summary>
        /// <param name="text"></param>
        /// <param name="testParameters"></param>
        /// <returns></returns>
        protected string ConvertParameters(string text, TestParameterCollection testParameters)
        {
            if (testParameters != null && testParameters.Count > 0)
            {
                foreach (ITestParameter testParameter in testParameters)
                {
                    string name = testParameter.Name;
                    text = text.Replace("@" + name, "${" + name + "}");
                }
            }

            return text;
        }

        /// <summary>
        /// Converts a MTM execution status for use in SpiraTest
        /// </summary>
        /// <param name="status">The MTM execution status</param>
        /// <returns>The SpiraTest execution status</returns>
        protected int ConvertExecutionStatus(TestOutcome status)
        {
            int statusId = 3; //Default to not run (no nulls allowed)
            switch (status)
            {
                case TestOutcome.Failed:
                case TestOutcome.Error:
                case TestOutcome.Timeout:
                    //Failed
                    statusId = 1;
                    break;
                case TestOutcome.Passed:
                    //Passed
                    statusId = 2;
                    break;
                case TestOutcome.Blocked:
                case TestOutcome.Aborted:
                    //Blocked
                    statusId = 5;
                    break;
                case TestOutcome.Warning:
                    //Caution
                    statusId = 6;
                    break;
            }
            return statusId;
        }

        /// <summary>
        /// Makes a string safe for use in XML (e.g. web service)
        /// </summary>
        /// <param name="input">The input string (as object)</param>
        /// <returns>The output string</returns>
        protected string MakeXmlSafe(object input)
        {
            //Handle null reference case
            if (input == null)
            {
                return "";
            }

            //Handle the case where the object is not a string
            string inputString;
            if (input.GetType() == typeof(string))
            {
                inputString = (string)input;
            }
            else
            {
                inputString = input.ToString();
            }

            //Handle empty string case
            if (inputString == "")
            {
                return inputString;
            }

            string output = inputString.Replace("\x00", "");
            output = output.Replace("\x01", "");
            output = output.Replace("\x02", "");
            output = output.Replace("\x03", "");
            output = output.Replace("\x04", "");
            output = output.Replace("\x05", "");
            output = output.Replace("\x06", "");
            output = output.Replace("\x07", "");
            output = output.Replace("\x08", "");
            output = output.Replace("\x0B", "");
            output = output.Replace("\x0C", "");
            output = output.Replace("\x0E", "");
            output = output.Replace("\x0F", "");
            output = output.Replace("\x10", "");
            output = output.Replace("\x11", "");
            output = output.Replace("\x12", "");
            output = output.Replace("\x13", "");
            output = output.Replace("\x14", "");
            output = output.Replace("\x15", "");
            output = output.Replace("\x16", "");
            output = output.Replace("\x17", "");
            output = output.Replace("\x18", "");
            output = output.Replace("\x19", "");
            output = output.Replace("\x1A", "");
            output = output.Replace("\x1B", "");
            output = output.Replace("\x1C", "");
            output = output.Replace("\x1D", "");
            output = output.Replace("\x1E", "");
            output = output.Replace("\x1F", "");
            return output;
        }

        /// <summary>
        /// Reconnects to a spira instance and project
        /// </summary>
        /// <param name="projectId">The id of the project</param>
        protected void ReconnectToSpira(int projectId)
        {
            //Reconnect to Spira
            ImportFormHandle.SpiraImportProxy.Connection_Authenticate2(Properties.Settings.Default.SpiraUserName, Properties.Settings.Default.SpiraPassword, ImportForm.API_PLUGIN_NAME);

            //Reconnect to Project
            ImportFormHandle.SpiraImportProxy.Connection_ConnectToProject(projectId);
        }

        /// <summary>
        /// Recursively imports a list of areas into SpiraTest as custom property values
        /// </summary>
        protected void ImportAreaNode(RemoteCustomList remoteCustomList_Area, Node areaNode, StreamWriter streamWriter)
        {
            RemoteCustomListValue remoteCustomListValue = new RemoteCustomListValue();
            remoteCustomListValue.CustomPropertyListId = remoteCustomList_Area.CustomPropertyListId.Value;
            remoteCustomListValue.Name = areaNode.Name;
            remoteCustomListValue = ImportFormHandle.SpiraImportProxy.CustomProperty_AddCustomListValue(remoteCustomListValue);
            this.areaCustomPropertyValueMapping.Add(areaNode.Path, remoteCustomListValue.CustomPropertyValueId.Value);
            streamWriter.WriteLine("Added area: '" + areaNode.Path);

            //Now get the children and load them under this node
            if (areaNode.HasChildNodes)
            {
                foreach (Node childNode in areaNode.ChildNodes)
                {
                    ImportAreaNode(remoteCustomList_Area, childNode, streamWriter);
                }
            }
        }

        /// <summary>
        /// Recursively imports a list of iterations into SpiraTest as releases
        /// </summary>
        protected void ImportIterationNode(Node iterationNode, StreamWriter streamWriter, int? parentReleaseId = null)
        {
            //Load the release and capture the ID
            RemoteRelease remoteRelease = new RemoteRelease();
            remoteRelease.Name = iterationNode.Name;
            remoteRelease.VersionNumber = iterationNode.Name.SafeSubstring(10);
            remoteRelease.StartDate = DateTime.UtcNow;
            remoteRelease.EndDate = DateTime.UtcNow.AddMonths(1);
            remoteRelease.ResourceCount = 1;
            int newReleaseId = ImportFormHandle.SpiraImportProxy.Release_Create(remoteRelease, parentReleaseId).ReleaseId.Value;
            streamWriter.WriteLine("Added iteration: '" + iterationNode.Path);

            //Add to the mapping hashtables
            this.releaseMapping.Add(iterationNode.Id, newReleaseId);
            this.releasePathMapping.Add(iterationNode.Path, newReleaseId);

            //Now get the children and load them under this node
            if (iterationNode.HasChildNodes)
            {
                foreach (Node childNode in iterationNode.ChildNodes)
                {
                    ImportIterationNode(childNode, streamWriter, newReleaseId);
                }
            }
        }
    }
}
