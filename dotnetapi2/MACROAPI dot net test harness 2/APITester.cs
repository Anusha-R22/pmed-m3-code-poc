using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using InferMed.MACRO.API;
using MACROUserBS30;
using VBA;

namespace MACROAPI_dot_net_test_harness
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class APITester : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button cmdCreate;
		private System.Windows.Forms.Button cmdInput;
		private System.Windows.Forms.Button cmdGetData;
		private System.Windows.Forms.Button cmdChangeUserDetails;
		private System.Windows.Forms.Button cmdGetUserDetails;
		private System.Windows.Forms.Button cmdChangeUserPassword;
		private System.Windows.Forms.Button cmdRegisterSubject;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.ComboBox cboStudy;
		private System.Windows.Forms.ComboBox cboSite;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.TextBox txtXmlSubject;
		private System.Windows.Forms.TextBox txtXml;
		private System.Windows.Forms.Button cmdVisit;
		private System.Windows.Forms.Button cmdeForm;
		private System.Windows.Forms.Button cmdQuestion;
		private System.Windows.Forms.Label lblSubject;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.TextBox txtMsg;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.Button cmdClose;
		private System.Windows.Forms.TextBox txtSubject;
		private System.Windows.Forms.TextBox txtOldPassword;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtNewPassword;
		private System.Windows.Forms.Button btnImportCategories;
		private System.Windows.Forms.Button btnExportCategories;
        private Label label3;
        private TextBox txtUser;
        private Button btnResetPwd;
        private Label label4;
        private TextBox txtResetPwd;
        private Button btnExportRoles;
        private Button btnImportRoles;
		private string _serialisedUser;

		public APITester(string serialisedUser)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
			_serialisedUser = serialisedUser;

			cmdInput.Enabled = false;
			cmdGetData.Enabled = false;

			// load studies
			if( LoadStudies() )
			{
				// sites
				LoadSites();
			}
		}

		private bool LoadStudies()
		{
			bool bLoaded = false;

			MACROUserClass userObj = new MACROUserBS30.MACROUserClass();
			userObj.SetStateHex( ref _serialisedUser);

			this.cboStudy.Items.Clear();

			if( (userObj.GetAllStudies()).Count() == 0)
			{
				bLoaded = false;
			}
			else
			{

                // sql
                string sql = "SELECT * FROM ClinicalTrial ORDER BY CLINICALTRIALNAME";
                // SELECT * FROM ClinicalTrial ORDER BY CLINICALTRIALNAME
                DataTable dtStudies = DataAccess.GetDataSet(userObj.CurrentDBConString, sql).Tables[0];
                foreach (DataRow dr in dtStudies.Rows)
                {
                    StudyInfo study = new StudyInfo( dr["CLINICALTRIALNAME"].ToString(), (int)dr["CLINICALTRIALID"] );
                    int nItem = cboStudy.Items.Add( study );
                }

                //foreach( MACROUserBS30.StudyClass st in userObj.GetAllStudies() )
                //{
                //    StudyInfo study = new StudyInfo( st.StudyName, st.StudyId );
                //    int nItem = cboStudy.Items.Add( study );
                //}

				cboStudy.SelectedIndex = 0;
				bLoaded = true;
			}

			userObj = null;
			return bLoaded;
		}

		private bool LoadSites()
		{
			bool bLoaded = false;

			cboSite.Items.Clear();

			MACROUserClass userObj = new MACROUserBS30.MACROUserClass();
			userObj.SetStateHex( ref _serialisedUser);

			int nStudyId = -1;
			if( cboStudy.SelectedIndex > -1)
			{
				StudyInfo studyInfo = (StudyInfo)(cboStudy.SelectedItem);
				nStudyId = studyInfo.StudyId;
			}

			if( userObj.GetAllSites(ref nStudyId).Count() > 0)
			{

                // get sites for a study
                // SELECT TRIALSITE FROM TRIALSITE WHERE CLINICALTRIALID = 
                // sql
                string sql = "SELECT TRIALSITE FROM TRIALSITE WHERE CLINICALTRIALID = " + nStudyId.ToString() + " ORDER BY TRIALSITE";
                DataTable dtStudySites = DataAccess.GetDataSet(userObj.CurrentDBConString, sql).Tables[0];
                foreach (DataRow dr in dtStudySites.Rows)
                {
                    cboSite.Items.Add(dr["TRIALSITE"].ToString());
                    //    int nItem = cboStudy.Items.Add( study );
                }
                
                //foreach( MACROUserBS30.SiteClass site in userObj.GetAllSites(ref nStudyId) )
                //{
                //    // can't open subjects from Remote sites on the Server
                //    if( ! (userObj.DBIsServer && site.SiteLocation == 1) )
                //    {
                //        cboSite.Items.Add( site.Site );
                //        //mcolWritableSites.Add LCase(oSite.Site), LCase(oSite.Site)
                //    }
                //}
			}
			
			if( cboSite.Items.Count > 0 )
			{
				cboSite.SelectedIndex = 0;
				bLoaded = true;
			}
			else
			{
				cboSite.Items.Clear();
				//Call MsgBox("There are no subjects available in this study")
			}

			return bLoaded;
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(APITester));
            this.cmdCreate = new System.Windows.Forms.Button();
            this.cmdInput = new System.Windows.Forms.Button();
            this.cmdGetData = new System.Windows.Forms.Button();
            this.cmdChangeUserDetails = new System.Windows.Forms.Button();
            this.cmdGetUserDetails = new System.Windows.Forms.Button();
            this.cmdChangeUserPassword = new System.Windows.Forms.Button();
            this.cmdRegisterSubject = new System.Windows.Forms.Button();
            this.cboStudy = new System.Windows.Forms.ComboBox();
            this.cboSite = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.txtXmlSubject = new System.Windows.Forms.TextBox();
            this.txtXml = new System.Windows.Forms.TextBox();
            this.cmdVisit = new System.Windows.Forms.Button();
            this.cmdeForm = new System.Windows.Forms.Button();
            this.cmdQuestion = new System.Windows.Forms.Button();
            this.lblSubject = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btnExportRoles = new System.Windows.Forms.Button();
            this.btnImportRoles = new System.Windows.Forms.Button();
            this.btnExportCategories = new System.Windows.Forms.Button();
            this.btnImportCategories = new System.Windows.Forms.Button();
            this.txtMsg = new System.Windows.Forms.TextBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtResetPwd = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtUser = new System.Windows.Forms.TextBox();
            this.btnResetPwd = new System.Windows.Forms.Button();
            this.txtNewPassword = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtOldPassword = new System.Windows.Forms.TextBox();
            this.cmdClose = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmdCreate
            // 
            this.cmdCreate.Location = new System.Drawing.Point(8, 152);
            this.cmdCreate.Name = "cmdCreate";
            this.cmdCreate.Size = new System.Drawing.Size(136, 23);
            this.cmdCreate.TabIndex = 1;
            this.cmdCreate.Text = "Create Subject";
            this.cmdCreate.Click += new System.EventHandler(this.cmdCreate_Click);
            // 
            // cmdInput
            // 
            this.cmdInput.Location = new System.Drawing.Point(16, 16);
            this.cmdInput.Name = "cmdInput";
            this.cmdInput.Size = new System.Drawing.Size(136, 23);
            this.cmdInput.TabIndex = 2;
            this.cmdInput.Text = "Input Data";
            this.cmdInput.Click += new System.EventHandler(this.cmdInput_Click);
            // 
            // cmdGetData
            // 
            this.cmdGetData.Location = new System.Drawing.Point(168, 16);
            this.cmdGetData.Name = "cmdGetData";
            this.cmdGetData.Size = new System.Drawing.Size(136, 23);
            this.cmdGetData.TabIndex = 3;
            this.cmdGetData.Text = "Get Data";
            this.cmdGetData.Click += new System.EventHandler(this.cmdGetData_Click);
            // 
            // cmdChangeUserDetails
            // 
            this.cmdChangeUserDetails.Enabled = false;
            this.cmdChangeUserDetails.Location = new System.Drawing.Point(624, 16);
            this.cmdChangeUserDetails.Name = "cmdChangeUserDetails";
            this.cmdChangeUserDetails.Size = new System.Drawing.Size(136, 23);
            this.cmdChangeUserDetails.TabIndex = 4;
            this.cmdChangeUserDetails.Text = "Change User Details";
            // 
            // cmdGetUserDetails
            // 
            this.cmdGetUserDetails.Enabled = false;
            this.cmdGetUserDetails.Location = new System.Drawing.Point(472, 16);
            this.cmdGetUserDetails.Name = "cmdGetUserDetails";
            this.cmdGetUserDetails.Size = new System.Drawing.Size(136, 23);
            this.cmdGetUserDetails.TabIndex = 5;
            this.cmdGetUserDetails.Text = "Get User Details";
            // 
            // cmdChangeUserPassword
            // 
            this.cmdChangeUserPassword.Location = new System.Drawing.Point(16, 24);
            this.cmdChangeUserPassword.Name = "cmdChangeUserPassword";
            this.cmdChangeUserPassword.Size = new System.Drawing.Size(136, 23);
            this.cmdChangeUserPassword.TabIndex = 6;
            this.cmdChangeUserPassword.Text = "Change User Password";
            this.cmdChangeUserPassword.Click += new System.EventHandler(this.cmdChangeUserPassword_Click);
            // 
            // cmdRegisterSubject
            // 
            this.cmdRegisterSubject.Location = new System.Drawing.Point(320, 16);
            this.cmdRegisterSubject.Name = "cmdRegisterSubject";
            this.cmdRegisterSubject.Size = new System.Drawing.Size(136, 23);
            this.cmdRegisterSubject.TabIndex = 7;
            this.cmdRegisterSubject.Text = "Register Subject";
            this.cmdRegisterSubject.Click += new System.EventHandler(this.cmdRegisterSubject_Click);
            // 
            // cboStudy
            // 
            this.cboStudy.Location = new System.Drawing.Point(8, 16);
            this.cboStudy.Name = "cboStudy";
            this.cboStudy.Size = new System.Drawing.Size(121, 21);
            this.cboStudy.TabIndex = 8;
            this.cboStudy.SelectedIndexChanged += new System.EventHandler(this.cboStudy_SelectedIndexChanged);
            // 
            // cboSite
            // 
            this.cboSite.Location = new System.Drawing.Point(8, 16);
            this.cboSite.Name = "cboSite";
            this.cboSite.Size = new System.Drawing.Size(121, 21);
            this.cboSite.TabIndex = 9;
            this.cboSite.SelectedIndexChanged += new System.EventHandler(this.cboSite_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cboStudy);
            this.groupBox1.Location = new System.Drawing.Point(8, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(136, 48);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Study";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cboSite);
            this.groupBox2.Location = new System.Drawing.Point(8, 48);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(136, 48);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Site";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.txtSubject);
            this.groupBox3.Location = new System.Drawing.Point(8, 96);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(136, 48);
            this.groupBox3.TabIndex = 12;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Subject";
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(8, 16);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(120, 20);
            this.txtSubject.TabIndex = 13;
            this.txtSubject.TextChanged += new System.EventHandler(this.txtSubject_TextChanged);
            // 
            // txtXmlSubject
            // 
            this.txtXmlSubject.Location = new System.Drawing.Point(152, 8);
            this.txtXmlSubject.Name = "txtXmlSubject";
            this.txtXmlSubject.Size = new System.Drawing.Size(536, 20);
            this.txtXmlSubject.TabIndex = 13;
            // 
            // txtXml
            // 
            this.txtXml.Location = new System.Drawing.Point(152, 33);
            this.txtXml.Multiline = true;
            this.txtXml.Name = "txtXml";
            this.txtXml.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtXml.Size = new System.Drawing.Size(536, 175);
            this.txtXml.TabIndex = 14;
            this.txtXml.Text = resources.GetString("txtXml.Text");
            // 
            // cmdVisit
            // 
            this.cmdVisit.Location = new System.Drawing.Point(696, 40);
            this.cmdVisit.Name = "cmdVisit";
            this.cmdVisit.Size = new System.Drawing.Size(75, 23);
            this.cmdVisit.TabIndex = 9;
            this.cmdVisit.Text = "Visit";
            this.cmdVisit.Click += new System.EventHandler(this.cmdVisit_Click);
            // 
            // cmdeForm
            // 
            this.cmdeForm.Location = new System.Drawing.Point(696, 72);
            this.cmdeForm.Name = "cmdeForm";
            this.cmdeForm.Size = new System.Drawing.Size(75, 23);
            this.cmdeForm.TabIndex = 10;
            this.cmdeForm.Text = "eForm";
            this.cmdeForm.Click += new System.EventHandler(this.cmdeForm_Click);
            // 
            // cmdQuestion
            // 
            this.cmdQuestion.Location = new System.Drawing.Point(696, 104);
            this.cmdQuestion.Name = "cmdQuestion";
            this.cmdQuestion.Size = new System.Drawing.Size(75, 23);
            this.cmdQuestion.TabIndex = 11;
            this.cmdQuestion.Text = "Question";
            this.cmdQuestion.Click += new System.EventHandler(this.cmdQuestion_Click);
            // 
            // lblSubject
            // 
            this.lblSubject.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblSubject.Location = new System.Drawing.Point(8, 184);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(136, 23);
            this.lblSubject.TabIndex = 15;
            this.lblSubject.Text = "Study/Site/Label";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.btnExportRoles);
            this.groupBox4.Controls.Add(this.btnImportRoles);
            this.groupBox4.Controls.Add(this.btnExportCategories);
            this.groupBox4.Controls.Add(this.btnImportCategories);
            this.groupBox4.Controls.Add(this.txtMsg);
            this.groupBox4.Controls.Add(this.cmdInput);
            this.groupBox4.Controls.Add(this.cmdGetData);
            this.groupBox4.Controls.Add(this.cmdRegisterSubject);
            this.groupBox4.Controls.Add(this.cmdGetUserDetails);
            this.groupBox4.Controls.Add(this.cmdChangeUserDetails);
            this.groupBox4.Location = new System.Drawing.Point(8, 216);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(776, 328);
            this.groupBox4.TabIndex = 16;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "XML Tests";
            // 
            // btnExportRoles
            // 
            this.btnExportRoles.Location = new System.Drawing.Point(320, 48);
            this.btnExportRoles.Name = "btnExportRoles";
            this.btnExportRoles.Size = new System.Drawing.Size(136, 23);
            this.btnExportRoles.TabIndex = 11;
            this.btnExportRoles.Text = "Export User Roles";
            this.btnExportRoles.UseVisualStyleBackColor = true;
            this.btnExportRoles.Click += new System.EventHandler(this.btnExportRoles_Click);
            // 
            // btnImportRoles
            // 
            this.btnImportRoles.Location = new System.Drawing.Point(172, 47);
            this.btnImportRoles.Name = "btnImportRoles";
            this.btnImportRoles.Size = new System.Drawing.Size(136, 23);
            this.btnImportRoles.TabIndex = 10;
            this.btnImportRoles.Text = "Import User Roles";
            this.btnImportRoles.UseVisualStyleBackColor = true;
            this.btnImportRoles.Click += new System.EventHandler(this.btnImportRoles_Click);
            // 
            // btnExportCategories
            // 
            this.btnExportCategories.Location = new System.Drawing.Point(624, 48);
            this.btnExportCategories.Name = "btnExportCategories";
            this.btnExportCategories.Size = new System.Drawing.Size(136, 23);
            this.btnExportCategories.TabIndex = 9;
            this.btnExportCategories.Text = "Export Categories";
            this.btnExportCategories.Click += new System.EventHandler(this.btnExportCategories_Click);
            // 
            // btnImportCategories
            // 
            this.btnImportCategories.Location = new System.Drawing.Point(472, 48);
            this.btnImportCategories.Name = "btnImportCategories";
            this.btnImportCategories.Size = new System.Drawing.Size(136, 23);
            this.btnImportCategories.TabIndex = 8;
            this.btnImportCategories.Text = "Import Categories";
            this.btnImportCategories.Click += new System.EventHandler(this.btnImportCategories_Click);
            // 
            // txtMsg
            // 
            this.txtMsg.Location = new System.Drawing.Point(8, 80);
            this.txtMsg.Multiline = true;
            this.txtMsg.Name = "txtMsg";
            this.txtMsg.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtMsg.Size = new System.Drawing.Size(760, 240);
            this.txtMsg.TabIndex = 0;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.label4);
            this.groupBox5.Controls.Add(this.txtResetPwd);
            this.groupBox5.Controls.Add(this.label3);
            this.groupBox5.Controls.Add(this.txtUser);
            this.groupBox5.Controls.Add(this.btnResetPwd);
            this.groupBox5.Controls.Add(this.txtNewPassword);
            this.groupBox5.Controls.Add(this.label2);
            this.groupBox5.Controls.Add(this.label1);
            this.groupBox5.Controls.Add(this.txtOldPassword);
            this.groupBox5.Controls.Add(this.cmdChangeUserPassword);
            this.groupBox5.Location = new System.Drawing.Point(8, 552);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(776, 100);
            this.groupBox5.TabIndex = 17;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Changing User Details";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(520, 54);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(80, 13);
            this.label4.TabIndex = 15;
            this.label4.Text = "New password:";
            // 
            // txtResetPwd
            // 
            this.txtResetPwd.Location = new System.Drawing.Point(606, 50);
            this.txtResetPwd.Name = "txtResetPwd";
            this.txtResetPwd.Size = new System.Drawing.Size(140, 20);
            this.txtResetPwd.TabIndex = 14;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(568, 24);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(32, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "User:";
            // 
            // txtUser
            // 
            this.txtUser.Location = new System.Drawing.Point(606, 24);
            this.txtUser.Name = "txtUser";
            this.txtUser.Size = new System.Drawing.Size(140, 20);
            this.txtUser.TabIndex = 12;
            // 
            // btnResetPwd
            // 
            this.btnResetPwd.Location = new System.Drawing.Point(441, 24);
            this.btnResetPwd.Name = "btnResetPwd";
            this.btnResetPwd.Size = new System.Drawing.Size(109, 23);
            this.btnResetPwd.TabIndex = 11;
            this.btnResetPwd.Text = "Reset Password";
            this.btnResetPwd.UseVisualStyleBackColor = true;
            this.btnResetPwd.Click += new System.EventHandler(this.btnResetPwd_Click);
            // 
            // txtNewPassword
            // 
            this.txtNewPassword.Location = new System.Drawing.Point(272, 51);
            this.txtNewPassword.Name = "txtNewPassword";
            this.txtNewPassword.PasswordChar = '*';
            this.txtNewPassword.Size = new System.Drawing.Size(144, 20);
            this.txtNewPassword.TabIndex = 10;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(160, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 23);
            this.label2.TabIndex = 9;
            this.label2.Text = "New Password";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(160, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 23);
            this.label1.TabIndex = 8;
            this.label1.Text = "Old Password";
            // 
            // txtOldPassword
            // 
            this.txtOldPassword.Location = new System.Drawing.Point(272, 24);
            this.txtOldPassword.Name = "txtOldPassword";
            this.txtOldPassword.PasswordChar = '*';
            this.txtOldPassword.Size = new System.Drawing.Size(144, 20);
            this.txtOldPassword.TabIndex = 7;
            // 
            // cmdClose
            // 
            this.cmdClose.Location = new System.Drawing.Point(704, 664);
            this.cmdClose.Name = "cmdClose";
            this.cmdClose.Size = new System.Drawing.Size(75, 23);
            this.cmdClose.TabIndex = 18;
            this.cmdClose.Text = "Close";
            this.cmdClose.Click += new System.EventHandler(this.cmdClose_Click);
            // 
            // APITester
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(792, 694);
            this.Controls.Add(this.cmdClose);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.lblSubject);
            this.Controls.Add(this.txtXml);
            this.Controls.Add(this.txtXmlSubject);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.cmdCreate);
            this.Controls.Add(this.cmdQuestion);
            this.Controls.Add(this.cmdeForm);
            this.Controls.Add(this.cmdVisit);
            this.Name = "APITester";
            this.Text = "API Tester";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		private void cboStudy_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if( cboStudy.SelectedIndex > -1)
			{
				StudyInfo studyInfo = (StudyInfo)(cboStudy.SelectedItem);
				LoadSites();
			}
			DoSubjectSpec();
		}

		private void cmdVisit_Click(object sender, System.EventArgs e)
		{
			txtXml.Text = " <Visit Code = 'vvv' Cycle = '1'>\r\n</Visit>\r\n" + txtXml.Text ;
		}

		private void cmdeForm_Click(object sender, System.EventArgs e)
		{
			txtXml.Text = "  <Eform Code = 'eee' Cycle = '1'>\r\n</Eform>\r\n" + txtXml.Text;
		}

		private void cmdQuestion_Click(object sender, System.EventArgs e)
		{
			txtXml.Text = "   <Question Code = 'qqq' Cycle = '1'/>\r\n" + txtXml.Text;
		}

		private void cmdClose_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void DoSubjectSpec()
		{
			// get study / site / subject
			string studyName = "";
			if( cboStudy.SelectedIndex > -1)
			{
				StudyInfo studyInfo = (StudyInfo)(cboStudy.SelectedItem);
				studyName = studyInfo.StudyName;
			}
			string siteName = "";
			if( cboSite.SelectedIndex > -1)
			{
				siteName = (string)cboSite.SelectedItem;
			}
			string subjLabel = txtSubject.Text;
			if(( siteName != "" ) && ( subjLabel != "" ) )
			{
				lblSubject.Text = studyName + "/" + siteName + "/" + subjLabel;
				DoSubjectXml(studyName, siteName, subjLabel);
				cmdInput.Enabled = true;
				cmdGetData.Enabled = true;
			}
			else
			{
				lblSubject.Text = "";
				txtXmlSubject.Text = "";
				cmdInput.Enabled = false;
				cmdGetData.Enabled = false;
			}
		}

		private void txtSubject_TextChanged(object sender, System.EventArgs e)
		{
			DoSubjectSpec();
		}

		private void cboSite_SelectedIndexChanged(object sender, System.EventArgs e)
		{

			DoSubjectSpec();
		}

		private void DoSubjectXml(string study, string site, string subject)
		{
			string subjXml = "<MACROSubject";
			subjXml += " Study = \"" + study + "\"";
			subjXml += " Site = \"" + site + "\"";
			subjXml += " Label = \"" + subject + "\">";
			txtXmlSubject.Text = subjXml;
		}

		private string GetCommitXml()
		{
			string xmlToCommit = "<?xml version=\"1.0\"?>\r\n";
			if( txtXmlSubject.Text != "" )
			{
				xmlToCommit += txtXmlSubject.Text + "\r\n" + txtXml.Text + "\r\n</MACROSubject>";
			}
			return xmlToCommit;
		}

		// Create subject
		private void cmdCreate_Click(object sender, System.EventArgs e)
		{
			string studyName = "";
			int studyId = -1;
			if( cboStudy.SelectedIndex > -1)
			{
				StudyInfo studyInfo = (StudyInfo)(cboStudy.SelectedItem);
				studyName = studyInfo.StudyName;
				studyId = studyInfo.StudyId;
			}
			string siteName = "";
			if( cboSite.SelectedIndex > -1)
			{
				siteName = (string)cboSite.SelectedItem;
			}
			if( MessageBox.Show( "Create a subject for " + studyName + "/" + siteName + "?", "Create Subject", MessageBoxButtons.YesNoCancel ) == DialogResult.Yes )
			{	
				this.Cursor = Cursors.WaitCursor;
				string message = "";
				txtMsg.Text = "Creating subject for " + studyName + "/" + siteName + "... Please wait...\r\n";
				DateTime dtCreateSubjectStart = new DateTime(DateTime.Now.Ticks);
				int subjectId = API.CreateSubject( _serialisedUser, studyId, siteName, ref message );
				DateTime dtCreateSubjectEnd = new DateTime(DateTime.Now.Ticks);
				TimeSpan tsCreateSubject = new TimeSpan( dtCreateSubjectEnd.Ticks - dtCreateSubjectStart.Ticks );
				txtMsg.Text += "Subject id = " + subjectId.ToString();
				if(subjectId < 1)
				{
					txtMsg.Text += "\r\n" + message;
				}
				txtMsg.Text += "\r\n" + "Time taken = " + tsCreateSubject.Seconds.ToString() + " second(s)";
				this.Cursor = Cursors.Default;
			}
		}

		// input xml data
		private void cmdInput_Click(object sender, System.EventArgs e)
		{
			if( MessageBox.Show( "Are you sure you wish to input data via the API?", "Input Xml", MessageBoxButtons.YesNoCancel ) == DialogResult.Yes )
			{
				this.Cursor = Cursors.WaitCursor;
				string message = "";
				txtMsg.Text = "Input Xml via API ... Please wait...\r\n";
				DateTime dtInputXmlStart = new DateTime(DateTime.Now.Ticks);
				// InputXmlSubjectData
				API.DataInputResult diResult = API.InputXMLSubjectData( _serialisedUser, GetCommitXml(), ref message );
				DateTime dtInputXmlEnd = new DateTime(DateTime.Now.Ticks);
				TimeSpan tsInputXml = new TimeSpan( dtInputXmlEnd.Ticks - dtInputXmlStart.Ticks );
				txtMsg.Text += "Input result - " + diResult.ToString() + "\r\nMessage - " + message + "\r\n";
				txtMsg.Text += "\r\n" + "Time taken = " + tsInputXml.Seconds.ToString() + " second(s)";
				this.Cursor = Cursors.Default;
			}
		}

		// retrieve xml data
		private void cmdGetData_Click(object sender, System.EventArgs e)
		{
			if( MessageBox.Show( "Are you sure you wish to retrieve data via the API?", "Retrieve Xml", MessageBoxButtons.YesNoCancel ) == DialogResult.Yes )
			{
				this.Cursor = Cursors.WaitCursor;
				string message = "";
				txtMsg.Text = "Retrieve Xml via API ... Please wait...\r\n";
				DateTime dtRequestXmlStart = new DateTime(DateTime.Now.Ticks);
				// GetXMLSubjectData
				API.DataRequestResult drResult = API.GetXMLSubjectData( _serialisedUser, GetCommitXml(), ref message );
				DateTime dtRequestXmlEnd = new DateTime(DateTime.Now.Ticks);
				TimeSpan tsRequestXml = new TimeSpan( dtRequestXmlEnd.Ticks - dtRequestXmlStart.Ticks );
				txtMsg.Text += "Retrieve result - " + drResult.ToString() + "\r\nMessage - " + message + "\r\n";
				txtMsg.Text += "\r\n" + "Time taken = " + tsRequestXml.Seconds.ToString() + " second(s)";
				this.Cursor = Cursors.Default;
			}
		}

		// register Subject
		private void cmdRegisterSubject_Click(object sender, System.EventArgs e)
		{
			string studyName = "";
			int studyId = -1;
			if( cboStudy.SelectedIndex > -1)
			{
				StudyInfo studyInfo = (StudyInfo)(cboStudy.SelectedItem);
				studyName = studyInfo.StudyName;
				studyId = studyInfo.StudyId;
			}
			string siteName = "";
			if( cboSite.SelectedIndex > -1)
			{
				siteName = (string)cboSite.SelectedItem;
			}
			string subject = txtSubject.Text;
			if( (studyName != "") && (siteName != "") && (subject != "") )
			{
				if( MessageBox.Show( "Are you sure you wish to register " + studyName + "/" + siteName + "/" + subject + " via the API?", "Register Subject", MessageBoxButtons.YesNoCancel ) == DialogResult.Yes )
				{
					this.Cursor = Cursors.WaitCursor;
					string regID = "";
					txtMsg.Text = "Registering subject via API ... Please wait...\r\n";
					DateTime dtRegisterXmlStart = new DateTime(DateTime.Now.Ticks);
					// RegisterSubject
					API.APIRegResult registerResult = API.RegisterSubject( _serialisedUser, studyName, siteName, subject, ref regID);
					DateTime dtRegisterXmlEnd = new DateTime(DateTime.Now.Ticks);
					TimeSpan tsRegisterXml = new TimeSpan( dtRegisterXmlEnd.Ticks - dtRegisterXmlStart.Ticks );
					txtMsg.Text += "Register result code - " + registerResult.ToString() + "\r\nReg ID - " + regID + "\r\n";
					txtMsg.Text += "\r\n" + "Time taken = " + tsRegisterXml.Seconds.ToString() + " second(s)";
					this.Cursor = Cursors.Default;
				}

			}
		}

		/*
		// Get user details
		private void cmdGetUserDetails_Click(object sender, System.EventArgs e)
		{
			string userName = "rde";

			object message = null;
			VBA.Collection userDetails = API.GetUsersDetails( _serialisedUser, userName, ref message );
			API.UserDetail userDetail;
			//userDetails.Count  
		}

		// change user details
		private void cmdChangeUserDetails_Click(object sender, System.EventArgs e)
		{
			string message = "";
			bool bChanged = API.ChangeUserDetails( _serialisedUser,  , message );
		}
		*/

		// change user password
		private void cmdChangeUserPassword_Click(object sender, System.EventArgs e)
		{
			string newPassword = txtNewPassword.Text;
			string oldPassword = txtOldPassword.Text;

			if( MessageBox.Show( "Are you sure you wish to change your user password via the API?", "Change User Password", MessageBoxButtons.YesNoCancel ) == DialogResult.Yes )
			{
				this.Cursor = Cursors.WaitCursor;
				string message = "";
				txtMsg.Text = "Changing password via API ... Please wait...\r\n";
				DateTime dtChangePasswordStart = new DateTime(DateTime.Now.Ticks);
				// ChangeUserPassword
				bool bChangePassword = API.ChangeUserPassword( ref _serialisedUser, newPassword, oldPassword, ref message );
				DateTime dtChangePasswordEnd = new DateTime(DateTime.Now.Ticks);
				TimeSpan tsChangePassword = new TimeSpan( dtChangePasswordEnd.Ticks - dtChangePasswordStart.Ticks );
				if( bChangePassword )
				{
					txtMsg.Text += "Password change Success.";
				}
				else
				{
					txtMsg.Text += "Password change Failed.\r\nMessage - " + message + "\r\n";
				}
				txtMsg.Text += "\r\n" + "Time taken = " + tsChangePassword.Seconds.ToString() + " second(s)";
				this.Cursor = Cursors.Default;
			}
		}

		private void btnImportCategories_Click(object sender, System.EventArgs e)
		{
			if( MessageBox.Show( "Are you sure you wish to input categories via the API?", "Input category Xml", MessageBoxButtons.YesNoCancel ) == DialogResult.Yes )
			{
				this.Cursor = Cursors.WaitCursor;
				string message = "";
				txtMsg.Text = "Input Category Xml via API ... Please wait...\r\n";
				DateTime dtInputXmlStart = new DateTime(DateTime.Now.Ticks);
				// InputXmlSubjectData
				API.ImportResult result = API.ImportCategories( _serialisedUser, txtXml.Text, ref message);
				DateTime dtInputXmlEnd = new DateTime(DateTime.Now.Ticks);
				TimeSpan tsInputXml = new TimeSpan( dtInputXmlEnd.Ticks - dtInputXmlStart.Ticks );
				txtMsg.Text += "Input result - " + result.ToString() + "\r\nMessage - " + message + "\r\n";
				txtMsg.Text += "\r\n" + "Time taken = " + tsInputXml.Seconds.ToString() + " second(s)";
				this.Cursor = Cursors.Default;
			}
		}

		private void btnExportCategories_Click(object sender, System.EventArgs e)
		{
			if( MessageBox.Show( "Are you sure you wish to export categories via the API?", "Export category Xml", MessageBoxButtons.YesNoCancel ) == DialogResult.Yes )
			{
				this.Cursor = Cursors.WaitCursor;
				string message = "";
				txtMsg.Text = "Exporting Categories via API ... Please wait...\r\n";
				DateTime dtInputXmlStart = new DateTime(DateTime.Now.Ticks);
				// InputXmlSubjectData
				bool result = API.ExportCategories( _serialisedUser, txtXml.Text, ref message);
				DateTime dtInputXmlEnd = new DateTime(DateTime.Now.Ticks);
				TimeSpan tsInputXml = new TimeSpan( dtInputXmlEnd.Ticks - dtInputXmlStart.Ticks );
				txtMsg.Text += "Input result - " + result.ToString() + "\r\nMessage - " + message + "\r\n";
				txtMsg.Text += "\r\n" + "Time taken = " + tsInputXml.Seconds.ToString() + " second(s)";
				this.Cursor = Cursors.Default;
			}
		}

        private void btnResetPwd_Click(object sender, EventArgs e)
        {
            string user = txtUser.Text.Trim();
            string pwd = txtResetPwd.Text.Trim();
            if (user == "" || pwd == "")
            {
                MessageBox.Show("Please enter a user name and password!", "Reset Password");
                return;
            }
            if (MessageBox.Show("Are you sure you wish to reset a password via the API?", "Reset Password", MessageBoxButtons.YesNoCancel) == DialogResult.Yes)
            {
                this.Cursor = Cursors.WaitCursor;
                string message = "";
                txtMsg.Text = "Resetting password via API ... Please wait...\r\n";
                DateTime dtInputXmlStart = new DateTime(DateTime.Now.Ticks);
                // Reset password
                API.PasswordResult result = API.ResetPassword(ref _serialisedUser, user, pwd, ref message);
                DateTime dtInputXmlEnd = new DateTime(DateTime.Now.Ticks);
                TimeSpan tsInputXml = new TimeSpan(dtInputXmlEnd.Ticks - dtInputXmlStart.Ticks);
                txtMsg.Text += "Password result - " + result.ToString() + "\r\nMessage - " + message + "\r\n";
                txtMsg.Text += "\r\n" + "Time taken = " + tsInputXml.Seconds.ToString() + " second(s)";
                this.Cursor = Cursors.Default;
            }

        }

        private void btnExportRoles_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you wish to export user roles via the API?", "Export User Roles", MessageBoxButtons.YesNoCancel) == DialogResult.Yes)
            {
                this.Cursor = Cursors.WaitCursor;
                string message = "";
                txtMsg.Text = "Exporting User Roles via API ... Please wait...\r\n";
                DateTime dtInputXmlStart = new DateTime(DateTime.Now.Ticks);
                // Export Associations
                bool result = API.ExportAssociations(_serialisedUser, txtXml.Text, ref message);
                DateTime dtInputXmlEnd = new DateTime(DateTime.Now.Ticks);
                TimeSpan tsInputXml = new TimeSpan(dtInputXmlEnd.Ticks - dtInputXmlStart.Ticks);
                txtMsg.Text += "Export result - " + result.ToString() + "\r\nOutput - " + message + "\r\n";
                txtMsg.Text += "\r\n" + "Time taken = " + tsInputXml.Seconds.ToString() + " second(s)";
                this.Cursor = Cursors.Default;
            }
        }

        private void btnImportRoles_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you wish to import user roles via the API?", "Import User Roles", MessageBoxButtons.YesNoCancel) == DialogResult.Yes)
            {
                this.Cursor = Cursors.WaitCursor;
                string message = "";
                txtMsg.Text = "Importing User Roles via API ... Please wait...\r\n";
                DateTime dtInputXmlStart = new DateTime(DateTime.Now.Ticks);
                // Import Associations
                API.ImportResult result = API.ImportAssociations(_serialisedUser, txtXml.Text, ref message);
                DateTime dtInputXmlEnd = new DateTime(DateTime.Now.Ticks);
                TimeSpan tsInputXml = new TimeSpan(dtInputXmlEnd.Ticks - dtInputXmlStart.Ticks);
                txtMsg.Text += "Import result - " + result.ToString() + "\r\nOutput - " + message + "\r\n";
                txtMsg.Text += "\r\n" + "Time taken = " + tsInputXml.Seconds.ToString() + " second(s)";
                this.Cursor = Cursors.Default;
            }
        }
	}

	class StudyInfo
	{
		public StudyInfo(string studyName, int studyId)
		{
			_studyName = studyName;
			_studyId = studyId;
		}

		private string _studyName;
		private int _studyId;

		public override string ToString()
		{
			return _studyName;
		}

		public string StudyName
		{
			get
			{
				return _studyName;
			}
		}

		public int StudyId
		{
			get
			{
				return _studyId;
			}
		}

	}
}

