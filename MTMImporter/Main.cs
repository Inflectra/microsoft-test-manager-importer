using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Proxy;
using System.Net;
using System.Collections.ObjectModel;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading;

namespace Inflectra.SpiraTest.AddOns.MTMImporter
{
    /// <summary>
    /// This is the code behind class for the utility that imports projects from
    /// Microsoft Test Manager (MTM) into Inflectra SpiraTest
    /// </summary>
    public class MainForm : System.Windows.Forms.Form
    {
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cboProject;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.TextBox txtLogin;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnAuthenticate;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtServer;
        private System.Windows.Forms.Button btnNext;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

        protected ImportForm importForm;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        public System.Windows.Forms.CheckBox chkImportTestRuns;
        public System.Windows.Forms.CheckBox chkImportDefects;
        public System.Windows.Forms.CheckBox chkImportUsers;
        public CheckBox chkImportAttachments;
        private CheckBox chkPassword;
        private Label label7;
        private TextBox txtDomain;
        public CheckBox chkImportTestSets;
        private Label label4;
        private ComboBox cboTestPlans;
        protected ProgressForm progressForm;

        #region Properties

        /// <summary>
        /// The project name
        /// </summary>
        public string ProjectName
        {
            get
            {
                return this.cboProject.SelectedText;
            }
        }

        #endregion

        public MainForm()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            // Add any event handlers

            //Set the initial state of any buttons
            this.btnNext.Enabled = false;

            //Create the other forms and set a handle to this form and the import form
            this.importForm = new ImportForm();
            this.progressForm = new ProgressForm();
            this.importForm.MainFormHandle = this;
            this.importForm.ProgressFormHandle = this.progressForm;
            this.progressForm.MainFormHandle = this;
            this.progressForm.ImportFormHandle = this.importForm;
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.btnNext = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtDomain = new System.Windows.Forms.TextBox();
            this.chkPassword = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtServer = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnAuthenticate = new System.Windows.Forms.Button();
            this.cboProject = new System.Windows.Forms.ComboBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtLogin = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkImportAttachments = new System.Windows.Forms.CheckBox();
            this.chkImportUsers = new System.Windows.Forms.CheckBox();
            this.chkImportDefects = new System.Windows.Forms.CheckBox();
            this.chkImportTestRuns = new System.Windows.Forms.CheckBox();
            this.chkImportTestSets = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cboTestPlans = new System.Windows.Forms.ComboBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnNext
            // 
            this.btnNext.Location = new System.Drawing.Point(376, 64);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(96, 23);
            this.btnNext.TabIndex = 0;
            this.btnNext.Text = "Next >";
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(272, 64);
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
            this.label1.Size = new System.Drawing.Size(440, 23);
            this.label1.TabIndex = 6;
            this.label1.Text = "SpiraTest | Import From MS Test Manager";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.cboTestPlans);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.txtDomain);
            this.groupBox1.Controls.Add(this.chkPassword);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.txtServer);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnAuthenticate);
            this.groupBox1.Controls.Add(this.cboProject);
            this.groupBox1.Controls.Add(this.txtPassword);
            this.groupBox1.Controls.Add(this.txtLogin);
            this.groupBox1.ForeColor = System.Drawing.Color.Black;
            this.groupBox1.Location = new System.Drawing.Point(24, 48);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(480, 204);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "MS Test Manager Configuration";
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(6, 87);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(80, 16);
            this.label7.TabIndex = 24;
            this.label7.Text = "Domain:";
            this.label7.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // txtDomain
            // 
            this.txtDomain.Location = new System.Drawing.Point(96, 83);
            this.txtDomain.Name = "txtDomain";
            this.txtDomain.Size = new System.Drawing.Size(128, 20);
            this.txtDomain.TabIndex = 23;
            // 
            // chkPassword
            // 
            this.chkPassword.AutoSize = true;
            this.chkPassword.Location = new System.Drawing.Point(152, 137);
            this.chkPassword.Name = "chkPassword";
            this.chkPassword.Size = new System.Drawing.Size(126, 17);
            this.chkPassword.TabIndex = 22;
            this.chkPassword.Text = "Remember Password";
            this.chkPassword.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(24, 24);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(64, 16);
            this.label6.TabIndex = 21;
            this.label6.Text = "Server:";
            this.label6.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // txtServer
            // 
            this.txtServer.Location = new System.Drawing.Point(96, 24);
            this.txtServer.Name = "txtServer";
            this.txtServer.Size = new System.Drawing.Size(336, 20);
            this.txtServer.TabIndex = 20;
            this.txtServer.Text = "http://localhost:8080/tfs";
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(16, 113);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(128, 16);
            this.label5.TabIndex = 19;
            this.label5.Text = "Project:";
            this.label5.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(232, 56);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 16);
            this.label3.TabIndex = 17;
            this.label3.Text = "Password:";
            this.label3.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(8, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(80, 16);
            this.label2.TabIndex = 16;
            this.label2.Text = "User Name:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // btnAuthenticate
            // 
            this.btnAuthenticate.Location = new System.Drawing.Point(344, 80);
            this.btnAuthenticate.Name = "btnAuthenticate";
            this.btnAuthenticate.Size = new System.Drawing.Size(88, 23);
            this.btnAuthenticate.TabIndex = 14;
            this.btnAuthenticate.Text = "Authenticate";
            this.btnAuthenticate.Click += new System.EventHandler(this.btnAuthenticate_Click);
            // 
            // cboProject
            // 
            this.cboProject.DisplayMember = "FullPath";
            this.cboProject.Location = new System.Drawing.Point(152, 113);
            this.cboProject.Name = "cboProject";
            this.cboProject.Size = new System.Drawing.Size(280, 21);
            this.cboProject.TabIndex = 13;
            this.cboProject.SelectedIndexChanged += new System.EventHandler(this.cboProject_SelectedIndexChanged);
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(304, 56);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(128, 20);
            this.txtPassword.TabIndex = 11;
            // 
            // txtLogin
            // 
            this.txtLogin.Location = new System.Drawing.Point(96, 56);
            this.txtLogin.Name = "txtLogin";
            this.txtLogin.Size = new System.Drawing.Size(128, 20);
            this.txtLogin.TabIndex = 10;
            this.txtLogin.Text = "alex_qc";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(472, 8);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(40, 40);
            this.pictureBox1.TabIndex = 11;
            this.pictureBox1.TabStop = false;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkImportTestSets);
            this.groupBox2.Controls.Add(this.chkImportAttachments);
            this.groupBox2.Controls.Add(this.chkImportUsers);
            this.groupBox2.Controls.Add(this.btnCancel);
            this.groupBox2.Controls.Add(this.chkImportDefects);
            this.groupBox2.Controls.Add(this.chkImportTestRuns);
            this.groupBox2.Controls.Add(this.btnNext);
            this.groupBox2.ForeColor = System.Drawing.Color.Black;
            this.groupBox2.Location = new System.Drawing.Point(24, 258);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(480, 96);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Import Options";
            // 
            // chkImportAttachments
            // 
            this.chkImportAttachments.AutoSize = true;
            this.chkImportAttachments.Location = new System.Drawing.Point(304, 27);
            this.chkImportAttachments.Name = "chkImportAttachments";
            this.chkImportAttachments.Size = new System.Drawing.Size(85, 17);
            this.chkImportAttachments.TabIndex = 8;
            this.chkImportAttachments.Text = "Attachments";
            this.chkImportAttachments.UseVisualStyleBackColor = true;
            // 
            // chkImportUsers
            // 
            this.chkImportUsers.Checked = true;
            this.chkImportUsers.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkImportUsers.Location = new System.Drawing.Point(19, 23);
            this.chkImportUsers.Name = "chkImportUsers";
            this.chkImportUsers.Size = new System.Drawing.Size(88, 24);
            this.chkImportUsers.TabIndex = 6;
            this.chkImportUsers.Text = "Users";
            this.chkImportUsers.CheckedChanged += new System.EventHandler(this.chkImportUsers_CheckedChanged);
            // 
            // chkImportDefects
            // 
            this.chkImportDefects.Enabled = false;
            this.chkImportDefects.ForeColor = System.Drawing.Color.Gray;
            this.chkImportDefects.Location = new System.Drawing.Point(216, 23);
            this.chkImportDefects.Name = "chkImportDefects";
            this.chkImportDefects.Size = new System.Drawing.Size(136, 24);
            this.chkImportDefects.TabIndex = 5;
            this.chkImportDefects.Text = "Defects";
            // 
            // chkImportTestRuns
            // 
            this.chkImportTestRuns.Checked = true;
            this.chkImportTestRuns.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkImportTestRuns.Location = new System.Drawing.Point(113, 23);
            this.chkImportTestRuns.Name = "chkImportTestRuns";
            this.chkImportTestRuns.Size = new System.Drawing.Size(184, 24);
            this.chkImportTestRuns.TabIndex = 4;
            this.chkImportTestRuns.Text = "Test Runs";
            // 
            // chkImportTestSets
            // 
            this.chkImportTestSets.AutoSize = true;
            this.chkImportTestSets.Checked = true;
            this.chkImportTestSets.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkImportTestSets.Location = new System.Drawing.Point(19, 53);
            this.chkImportTestSets.Name = "chkImportTestSets";
            this.chkImportTestSets.Size = new System.Drawing.Size(71, 17);
            this.chkImportTestSets.TabIndex = 9;
            this.chkImportTestSets.Text = "Test Sets";
            this.chkImportTestSets.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(16, 160);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(128, 16);
            this.label4.TabIndex = 26;
            this.label4.Text = "Test Plan:";
            this.label4.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // cboTestPlans
            // 
            this.cboTestPlans.DisplayMember = "FullPath";
            this.cboTestPlans.Location = new System.Drawing.Point(152, 160);
            this.cboTestPlans.Name = "cboTestPlans";
            this.cboTestPlans.Size = new System.Drawing.Size(280, 21);
            this.cboTestPlans.TabIndex = 25;
            // 
            // MainForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(526, 366);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.Text = "Microsoft Test Manager Importer for Inflectra SpiraTest";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }
        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
            Application.ThreadException += new ThreadExceptionEventHandler(Application_ThreadException);

            //We need to catch these so that we can let the user know
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            //We handle the case of the TFS assemblies not being present
            if (e.Exception.GetType() == typeof(FileNotFoundException))
            {
                //This happens if the TFS client libraries are not available
                MessageBox.Show("Error Connecting to MTM. You need to make sure that this importer is installed on a computer that already has Visual Studio and/or TFS Team Explorer installed (" + e.Exception.Message + ")", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            MessageBox.Show("An unhandled exception has occured, please report this to Inflectra Support: " + e.Exception.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MiniDump.CreateMiniDump();
        }

        protected TestManagementService GetMtmServiceInstance(out TfsTeamProjectCollection tfsTeamProjectCollection)
        {
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
            networkCredential = new NetworkCredential(this.txtLogin.Text.Trim(), this.txtPassword.Text.Trim(), this.txtDomain.Text.Trim());
            /*}*/

            //Create a new TFS 2012 project collection instance and WorkItemStore instance
            //This requires that the URI includes the collection name not just the server name
            Uri tfsUri = new Uri(this.txtServer.Text.Trim());
            if (tfsCredential == null)
            {
                tfsTeamProjectCollection = new TfsTeamProjectCollection(tfsUri, networkCredential);
            }
            else
            {
                tfsTeamProjectCollection = new TfsTeamProjectCollection(tfsUri, tfsCredential);
            }

            try
            {
                //Now try authenticating
                try
                {
                    tfsTeamProjectCollection.Authenticate();
                }
                catch (Exception exception)
                {
                    MessageBox.Show("That domain/login/password combination was not valid for the instance of MTM:" + exception.Message, "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return  null;
                }

                //Get access to the work item and test manager server services
                WorkItemServer workItemServer = tfsTeamProjectCollection.GetService<WorkItemServer>();
                TestManagementService testManagementService = tfsTeamProjectCollection.GetService<TestManagementService>();

                //Make sure Team Manager is a supported feature
                if (!testManagementService.IsSupported())
                {
                    MessageBox.Show("This instance of TFS does not support Microsoft Test Management Services!", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return null;
                }
                return testManagementService;
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error Logging in to MTM: " + exception.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// Authenticates the user from the providing server/login/password information
        /// </summary>
        /// <param name="sender">The sending object</param>
        /// <param name="e">The event arguments</param>
        private void btnAuthenticate_Click(object sender, System.EventArgs e)
        {
            try
            {
                //Disable the login and next button
                this.btnNext.Enabled = false;

                //Make sure that a login was entered
                if (this.txtLogin.Text.Trim() == "")
                {
                    MessageBox.Show("You need to enter a MTM login", "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                //Make sure that a server was entered
                if (this.txtServer.Text.Trim() == "")
                {
                    MessageBox.Show("You need to enter a MTM server URL", "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                //Enable self-signed certificates
                PermissiveCertificatePolicy.Enact("");

                TfsTeamProjectCollection tfsTeamProjectCollection;
                TestManagementService testManagementService = GetMtmServiceInstance(out tfsTeamProjectCollection);

                if (testManagementService != null)
                {
                    MessageBox.Show("You have logged into MTM Successfully.\nPlease choose a Project from the list below:", "Authentication", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    //Now we need to populate the list of projects
                    ReadOnlyCollection<CatalogNode> projectNodes = tfsTeamProjectCollection.CatalogNode.QueryChildren(new[] { CatalogResourceTypes.TeamProject }, false, CatalogQueryOptions.None);

                    //Get the names into a simple list
                    List<string> projectNames = new List<string>();
                    foreach (CatalogNode projectNode in projectNodes)
                    {
                        if (projectNode.Resource != null)
                        {
                            string projectName = projectNode.Resource.DisplayName;
                            projectNames.Add(projectName);
                        }
                    }
                    this.cboProject.DataSource = projectNames;

                    //Also load in the test plans for this project
                    LoadTestPlansForProject();

                    //Enable the 'Next' button
                    this.btnNext.Enabled = true;
                }
            }
            catch (FileNotFoundException exception)
            {
                //This happens if the TFS client libraries are not available
                MessageBox.Show("Error Connecting to MTM. You need to make sure that this importer is installed on a computer that already has Visual Studio and/or TFS Team Explorer installed (" + exception.Message + ")", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            catch (TypeLoadException exception)
            {
                // The runtime coulnd't find referenced type in the assembly.
                MessageBox.Show("Error Connecting to MTM. You need to make sure that this importer is installed on a computer that already has Visual Studio and/or TFS Team Explorer installed (" + exception.Message + ")", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        protected void LoadTestPlansForProject()
        {
            //Create the list of plans
            List<string> testPlans = new List<string>();
            testPlans.Add(Utils.ALL_TEST_PLANS);

            //Get the current project name
            if (this.cboProject.SelectedValue != null)
            {
                string projectName = this.cboProject.SelectedValue.ToString();
                if (!String.IsNullOrWhiteSpace(projectName))
                {
                    //Get the list of plans for this project
                    TfsTeamProjectCollection tfsTeamProjectCollection;
                    TestManagementService testManagementService = GetMtmServiceInstance(out tfsTeamProjectCollection);
                    ITestManagementTeamProject mtmProject = testManagementService.GetTeamProject(projectName);
                    if (mtmProject != null)
                    {
                        ITestPlanCollection plans = mtmProject.TestPlans.Query("Select * From TestPlan");
                        foreach (ITestPlan mtmTestPlan in plans)
                        {
                            testPlans.Add(mtmTestPlan.Name);
                        }
                    }
                }
            }

            //Set the datasource
            this.cboTestPlans.DataSource = testPlans;
        }

        /// <summary>
        /// Closes the application
        /// </summary>
        /// <param name="sender">The sending object</param>
        /// <param name="e">The event arguments</param>
        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            //Close the application
            this.Close();
        }

        /// <summary>
        /// Called when the Next button is clicked. Switches to the second form
        /// </summary>
        /// <param name="sender">The sending object</param>
        /// <param name="e">The event arguments</param>
        private void btnNext_Click(object sender, System.EventArgs e)
        {
            //Store the info in settings for later
            Properties.Settings.Default.MtmUrl = this.txtServer.Text.Trim();
            Properties.Settings.Default.MtmUserName = this.txtLogin.Text.Trim();
            Properties.Settings.Default.MtmDomain = this.txtDomain.Text.Trim();
            if (chkPassword.Checked)
            {
                Properties.Settings.Default.MtmPassword = this.txtPassword.Text;  //Don't trim in case it contains a space
            }
            else
            {
                Properties.Settings.Default.MtmPassword = "";
            }
            Properties.Settings.Default.TestRuns = this.chkImportTestRuns.Checked;
            Properties.Settings.Default.TestSets = this.chkImportTestSets.Checked;
            Properties.Settings.Default.Defects = this.chkImportDefects.Checked;
            Properties.Settings.Default.Attachments = this.chkImportAttachments.Checked;
            Properties.Settings.Default.Users = this.chkImportUsers.Checked;
            Properties.Settings.Default.Save();

            //If we're not remembering the password, set the property again, now after the save for temporary access by the
            //import thread
            if (!chkPassword.Checked)
            {
                Properties.Settings.Default.MtmPassword = this.txtPassword.Text;
            }

            //Store (but don't save) the MTM project name and test plan name as well
            Properties.Settings.Default.MtmProjectName = this.cboProject.SelectedValue.ToString();
            Properties.Settings.Default.MtmTestPlanName = this.cboTestPlans.SelectedValue.ToString();

            //Hide the current form
            this.Hide();

            //Show the second page in the import wizard
            this.importForm.Show();
        }

        /// <summary>
        /// Populates the fields when the form is loaded
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_Load(object sender, EventArgs e)
        {
            this.txtServer.Text = Properties.Settings.Default.MtmUrl;
            this.txtLogin.Text = Properties.Settings.Default.MtmUserName;
            this.txtDomain.Text = Properties.Settings.Default.MtmDomain;
            if (String.IsNullOrEmpty(Properties.Settings.Default.MtmPassword))
            {
                this.chkPassword.Checked = false;
                this.txtPassword.Text = "";
            }
            else
            {
                this.chkPassword.Checked = true;
                this.txtPassword.Text = Properties.Settings.Default.SpiraPassword;
            }

            this.chkImportTestRuns.Checked = Properties.Settings.Default.TestRuns;
            this.chkImportTestSets.Checked = Properties.Settings.Default.TestSets;
            this.chkImportDefects.Checked = Properties.Settings.Default.Defects;
            this.chkImportAttachments.Checked = Properties.Settings.Default.Attachments;
            this.chkImportUsers.Checked = Properties.Settings.Default.Users;
        }

        private void chkImportUsers_CheckedChanged(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// When the project changes, change the list of test plans
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboProject_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadTestPlansForProject();
        }
    }
}
