using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using InferMed.Components;
using System.Data.OracleClient;
using System.Data.SqlClient;
using log4net.Config;
using log4net;
using System.IO;
using MACROLOCKBS30;
using System.Diagnostics;
//using MACROAZRBBS30;


//----------------------------------------------------------------------
// 28/04/2006	bug 2725 incorrect matching of eform elements
// 28/04/2006 bug 2726 incorrect matching and display of rqg elements
// 10/05/2006 rebuild arezzo - fix to allow 3.0.72 version to work without 3.0.75 sd module
// 07/09/2006	bug 2864 prevent different datatypes being matched
// 07/09/2006 bug 2890 modifications to auto matching for reporting purposes
// 21/09/2006 bug 2805 add study difference reporting functionality (not for release)
// 14/10/2006 bug 2819 use .net login component instead of stub module
//----------------------------------------------------------------------

namespace InferMed.MACRO.StudyCopy
{
	/// <summary>
	/// Main user interface for study copying utility
	/// </summary>
	public class MainForm : System.Windows.Forms.Form
	{
		//settings file
		private const string _SETTINGS_FILE = "MACROSettings30.txt";

		//study definition function code
		private const string _FN_STUDYDEFINITION = "F1004";

		//currently logged into database connection string
		private string _dbCon = "";
		private string _dbCode = "";

		//security db connection string
		private string _secCon = "";

		//user
		private string _userName = "";
		private string _userNameFull = "";
		private string _MACRORole = "";

		//database connection object
		public static IDbConnection _dbConn = null;

		//settings
		private SettingsForm _settingsForm = null;

		//confirm copy
		private ConfirmCopyForm _copyForm = null;

		//study state
		private StudyState _state = null;

		//logging
		private static readonly ILog log = LogManager.GetLogger( typeof( MainForm ) );

		//study locking
		private string _lockToken = "";
		private long _lockStudy = 0;


		private System.Windows.Forms.StatusBar statusBar1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.ComboBox cboStudiesD;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.ComboBox cboStudiesS;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.CheckBox chkCLP;
		private System.Windows.Forms.CheckBox chkR;
		private System.Windows.Forms.Splitter splitter1;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.ListView lvwEforms;
		private System.Windows.Forms.MainMenu mainMenu1;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem2;
		private System.Windows.Forms.ListView lvwEformElements;
		private System.Windows.Forms.ContextMenu contextMenu1;
		private System.Windows.Forms.ContextMenu contextMenu2;
		private System.Windows.Forms.MenuItem menuItem3;
		private System.Windows.Forms.MenuItem menuItem4;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.MenuItem menuItem7;
		private System.Windows.Forms.MenuItem menuItem8;
		private System.Windows.Forms.Button btnCopyEform;
		private System.Windows.Forms.Button btnCopyEformElements;
		private System.Windows.Forms.PictureBox picStudyLocked;
		private System.Windows.Forms.Button btnSelectAllElements;
		private System.Windows.Forms.MenuItem menuItem5;
		private InferMed.Components.IMEDLogin.IMEDLogin imedLogin1;
		private System.ComponentModel.IContainer components;

		private bool PrevInstance()
		{
			if( Process.GetProcessesByName( Process.GetCurrentProcess().ProcessName ).Length > 1 )
			{
				return( true );
			}
			else
			{
				return( false );
			}
		}

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new MainForm());
		}

		public MainForm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//initialise logging
			XmlConfigurator.Configure( new FileInfo( Path.GetDirectoryName( AppDomain.CurrentDomain.BaseDirectory ) 
				+ @"\log4netconfig.xml" ) );
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MainForm));
			this.statusBar1 = new System.Windows.Forms.StatusBar();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.picStudyLocked = new System.Windows.Forms.PictureBox();
			this.cboStudiesD = new System.Windows.Forms.ComboBox();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.cboStudiesS = new System.Windows.Forms.ComboBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.btnCopyEform = new System.Windows.Forms.Button();
			this.lvwEforms = new System.Windows.Forms.ListView();
			this.contextMenu1 = new System.Windows.Forms.ContextMenu();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.splitter1 = new System.Windows.Forms.Splitter();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.imedLogin1 = new InferMed.Components.IMEDLogin.IMEDLogin();
			this.btnSelectAllElements = new System.Windows.Forms.Button();
			this.btnCopyEformElements = new System.Windows.Forms.Button();
			this.chkCLP = new System.Windows.Forms.CheckBox();
			this.chkR = new System.Windows.Forms.CheckBox();
			this.lvwEformElements = new System.Windows.Forms.ListView();
			this.contextMenu2 = new System.Windows.Forms.ContextMenu();
			this.mainMenu1 = new System.Windows.Forms.MainMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.menuItem8 = new System.Windows.Forms.MenuItem();
			this.menuItem5 = new System.Windows.Forms.MenuItem();
			this.menuItem7 = new System.Windows.Forms.MenuItem();
			this.menuItem2 = new System.Windows.Forms.MenuItem();
			this.menuItem3 = new System.Windows.Forms.MenuItem();
			this.menuItem4 = new System.Windows.Forms.MenuItem();
			this.groupBox2.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.groupBox4.SuspendLayout();
			this.groupBox5.SuspendLayout();
			this.SuspendLayout();
			// 
			// statusBar1
			// 
			this.statusBar1.Location = new System.Drawing.Point(0, 609);
			this.statusBar1.Name = "statusBar1";
			this.statusBar1.Size = new System.Drawing.Size(920, 20);
			this.statusBar1.TabIndex = 3;
			this.statusBar1.Text = "statusBar1";
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.picStudyLocked);
			this.groupBox2.Controls.Add(this.cboStudiesD);
			this.groupBox2.Location = new System.Drawing.Point(4, 4);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(376, 52);
			this.groupBox2.TabIndex = 5;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Destination study";
			// 
			// picStudyLocked
			// 
			this.picStudyLocked.Image = ((System.Drawing.Image)(resources.GetObject("picStudyLocked.Image")));
			this.picStudyLocked.Location = new System.Drawing.Point(356, 24);
			this.picStudyLocked.Name = "picStudyLocked";
			this.picStudyLocked.Size = new System.Drawing.Size(16, 16);
			this.picStudyLocked.TabIndex = 8;
			this.picStudyLocked.TabStop = false;
			this.picStudyLocked.Visible = false;
			// 
			// cboStudiesD
			// 
			this.cboStudiesD.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.cboStudiesD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboStudiesD.Location = new System.Drawing.Point(8, 20);
			this.cboStudiesD.Name = "cboStudiesD";
			this.cboStudiesD.Size = new System.Drawing.Size(348, 21);
			this.cboStudiesD.TabIndex = 7;
			this.cboStudiesD.SelectedIndexChanged += new System.EventHandler(this.cboStudiesD_SelectedIndexChanged);
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Controls.Add(this.cboStudiesS);
			this.groupBox3.Location = new System.Drawing.Point(384, 4);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(532, 52);
			this.groupBox3.TabIndex = 11;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Source study";
			// 
			// cboStudiesS
			// 
			this.cboStudiesS.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.cboStudiesS.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboStudiesS.Location = new System.Drawing.Point(8, 20);
			this.cboStudiesS.Name = "cboStudiesS";
			this.cboStudiesS.Size = new System.Drawing.Size(516, 21);
			this.cboStudiesS.TabIndex = 2;
			this.cboStudiesS.SelectedIndexChanged += new System.EventHandler(this.cboStudiesS_SelectedIndexChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.groupBox4);
			this.groupBox1.Controls.Add(this.splitter1);
			this.groupBox1.Controls.Add(this.groupBox5);
			this.groupBox1.Location = new System.Drawing.Point(0, 60);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(920, 547);
			this.groupBox1.TabIndex = 12;
			this.groupBox1.TabStop = false;
			// 
			// groupBox4
			// 
			this.groupBox4.Controls.Add(this.btnCopyEform);
			this.groupBox4.Controls.Add(this.lvwEforms);
			this.groupBox4.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox4.Location = new System.Drawing.Point(3, 16);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(914, 204);
			this.groupBox4.TabIndex = 17;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "eForms";
			// 
			// btnCopyEform
			// 
			this.btnCopyEform.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCopyEform.Enabled = false;
			this.btnCopyEform.Location = new System.Drawing.Point(829, 172);
			this.btnCopyEform.Name = "btnCopyEform";
			this.btnCopyEform.Size = new System.Drawing.Size(72, 24);
			this.btnCopyEform.TabIndex = 8;
			this.btnCopyEform.Text = "Copy";
			this.btnCopyEform.Click += new System.EventHandler(this.btnCopyEform_Click);
			// 
			// lvwEforms
			// 
			this.lvwEforms.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lvwEforms.ContextMenu = this.contextMenu1;
			this.lvwEforms.HideSelection = false;
			this.lvwEforms.Location = new System.Drawing.Point(4, 16);
			this.lvwEforms.Name = "lvwEforms";
			this.lvwEforms.Size = new System.Drawing.Size(905, 148);
			this.lvwEforms.SmallImageList = this.imageList1;
			this.lvwEforms.TabIndex = 1;
			this.lvwEforms.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lvwEforms_onKeyDown);
			this.lvwEforms.SelectedIndexChanged += new System.EventHandler(this.lvwEforms_SelectedIndexChanged);
			// 
			// contextMenu1
			// 
			this.contextMenu1.Popup += new System.EventHandler(this.contextMenu1_Popup);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(12, 12);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// splitter1
			// 
			this.splitter1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.splitter1.Location = new System.Drawing.Point(3, 220);
			this.splitter1.MinExtra = 150;
			this.splitter1.MinSize = 150;
			this.splitter1.Name = "splitter1";
			this.splitter1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.splitter1.Size = new System.Drawing.Size(914, 8);
			this.splitter1.TabIndex = 16;
			this.splitter1.TabStop = false;
			// 
			// groupBox5
			// 
			this.groupBox5.Controls.Add(this.imedLogin1);
			this.groupBox5.Controls.Add(this.btnSelectAllElements);
			this.groupBox5.Controls.Add(this.btnCopyEformElements);
			this.groupBox5.Controls.Add(this.chkCLP);
			this.groupBox5.Controls.Add(this.chkR);
			this.groupBox5.Controls.Add(this.lvwEformElements);
			this.groupBox5.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.groupBox5.Location = new System.Drawing.Point(3, 228);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(914, 316);
			this.groupBox5.TabIndex = 15;
			this.groupBox5.TabStop = false;
			this.groupBox5.Text = "eForm Elements";
			// 
			// imedLogin1
			// 
			this.imedLogin1.ApplicationPermissionCheck = "";
			this.imedLogin1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("imedLogin1.BackgroundImage")));
			this.imedLogin1.DataDBSetting = "";
			this.imedLogin1.InitialSettingsFile = "";
			this.imedLogin1.Location = new System.Drawing.Point(44, 44);
			this.imedLogin1.MandatoryMACROLogin = true;
			this.imedLogin1.MandatorySecurityDB = true;
			this.imedLogin1.Name = "imedLogin1";
			this.imedLogin1.SecurityDbLoginOnly = false;
			this.imedLogin1.Size = new System.Drawing.Size(40, 40);
			this.imedLogin1.TabIndex = 15;
			this.imedLogin1.Visible = false;
			// 
			// btnSelectAllElements
			// 
			this.btnSelectAllElements.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSelectAllElements.Location = new System.Drawing.Point(756, 285);
			this.btnSelectAllElements.Name = "btnSelectAllElements";
			this.btnSelectAllElements.Size = new System.Drawing.Size(68, 24);
			this.btnSelectAllElements.TabIndex = 14;
			this.btnSelectAllElements.Text = "Select All";
			this.btnSelectAllElements.Click += new System.EventHandler(this.btnSelectAllElements_Click);
			// 
			// btnCopyEformElements
			// 
			this.btnCopyEformElements.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCopyEformElements.Enabled = false;
			this.btnCopyEformElements.Location = new System.Drawing.Point(833, 285);
			this.btnCopyEformElements.Name = "btnCopyEformElements";
			this.btnCopyEformElements.Size = new System.Drawing.Size(68, 24);
			this.btnCopyEformElements.TabIndex = 13;
			this.btnCopyEformElements.Text = "Copy";
			this.btnCopyEformElements.Click += new System.EventHandler(this.btnCopyEformElements_Click);
			// 
			// chkCLP
			// 
			this.chkCLP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.chkCLP.Checked = true;
			this.chkCLP.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCLP.Location = new System.Drawing.Point(100, 289);
			this.chkCLP.Name = "chkCLP";
			this.chkCLP.Size = new System.Drawing.Size(152, 16);
			this.chkCLP.TabIndex = 11;
			this.chkCLP.Text = "Comments/Lines/Pictures";
			this.chkCLP.CheckedChanged += new System.EventHandler(this.chkCLP_CheckedChanged);
			// 
			// chkR
			// 
			this.chkR.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.chkR.Checked = true;
			this.chkR.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkR.Location = new System.Drawing.Point(8, 289);
			this.chkR.Name = "chkR";
			this.chkR.Size = new System.Drawing.Size(92, 16);
			this.chkR.TabIndex = 10;
			this.chkR.Text = "Responses";
			this.chkR.CheckedChanged += new System.EventHandler(this.chkR_CheckedChanged);
			// 
			// lvwEformElements
			// 
			this.lvwEformElements.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lvwEformElements.ContextMenu = this.contextMenu2;
			this.lvwEformElements.HideSelection = false;
			this.lvwEformElements.Location = new System.Drawing.Point(4, 16);
			this.lvwEformElements.Name = "lvwEformElements";
			this.lvwEformElements.Size = new System.Drawing.Size(905, 260);
			this.lvwEformElements.SmallImageList = this.imageList1;
			this.lvwEformElements.TabIndex = 7;
			this.lvwEformElements.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lvwEformElements_onKeyDown);
			this.lvwEformElements.SelectedIndexChanged += new System.EventHandler(this.lvwEformElements_SelectedIndexChanged);
			// 
			// contextMenu2
			// 
			this.contextMenu2.Popup += new System.EventHandler(this.contextMenu2_Popup);
			// 
			// mainMenu1
			// 
			this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																							this.menuItem1,
																																							this.menuItem3});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																							this.menuItem8,
																																							this.menuItem5,
																																							this.menuItem7,
																																							this.menuItem2});
			this.menuItem1.Text = "&File";
			// 
			// menuItem8
			// 
			this.menuItem8.Index = 0;
			this.menuItem8.Text = "&Settings...";
			this.menuItem8.Click += new System.EventHandler(this.menuItem8_Click);
			// 
			// menuItem5
			// 
			this.menuItem5.Enabled = false;
			this.menuItem5.Index = 1;
			this.menuItem5.Text = "Reports...";
			this.menuItem5.Visible = false;
			this.menuItem5.Click += new System.EventHandler(this.menuItem5_Click);
			// 
			// menuItem7
			// 
			this.menuItem7.Index = 2;
			this.menuItem7.Text = "-";
			// 
			// menuItem2
			// 
			this.menuItem2.Index = 3;
			this.menuItem2.Text = "E&xit";
			this.menuItem2.Click += new System.EventHandler(this.menuItem2_Click);
			// 
			// menuItem3
			// 
			this.menuItem3.Index = 1;
			this.menuItem3.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																							this.menuItem4});
			this.menuItem3.Text = "&Help";
			// 
			// menuItem4
			// 
			this.menuItem4.Index = 0;
			this.menuItem4.Text = "&About";
			this.menuItem4.Click += new System.EventHandler(this.menuItem4_Click);
			// 
			// MainForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(920, 629);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.statusBar1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Menu = this.mainMenu1;
			this.Name = "MainForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "MACRO Study Copy Tool";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.MainForm_Closing);
			this.Load += new System.EventHandler(this.MainForm_Load);
			this.groupBox2.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			this.groupBox4.ResumeLayout(false);
			this.groupBox5.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Show progress
		/// </summary>
		/// <param name="prog"></param>
		private void ShowProgress( string prog, bool msg, StudyCopyGlobal.LogPriority p )
		{
			statusBar1.Text = prog;
			Application.DoEvents();
			if( msg ) MessageBox.Show( prog );

			//logging
			log.Debug( "User status message :" + prog );
		}

		/// <summary>
		/// Lock macro study
		/// </summary>
		/// <param name="clinicalTrialId"></param>
		/// <returns></returns>
		private string LockStudy( long clinicalTrialId )
		{
			DBLockClass l = new DBLockClass();
			float wait = 0;
			string token;

//			try
//			{
				token = l.LockStudy( _dbCon, _userName, System.Convert.ToInt32( clinicalTrialId ), wait );
				switch( token )
				{
					case "0":
					case "1":
					case "2":
						token = "";
						break;
					default:
						break;
				}
				return( token );
//			}
//			catch
//			{
//				return( "" );
//			}
		}

		/// <summary>
		/// Rebuild arezzo after study changes
		/// rebuild arezzo - fix to allow 3.0.72 version to work without 3.0.75 sd module
		/// ic 10/07/06 commented out for recompile with 3.0.75
		/// </summary>
		private void RebuildArezzo()
		{
//			AzRebuildClass rb = new AzRebuildClass();
//			
//			try
//			{
//				ShowProgress( "Updating Arezzo...", false, StudyCopyGlobal.LogPriority.Normal );
//				rb.DoAREZZOUpdates( _dbCon, System.Convert.ToInt16( _lockStudy ) );
//				ShowProgress( "Ready", false, StudyCopyGlobal.LogPriority.Normal );
//			}
//#if( !DEBUG )
//			catch( Exception ex )
//			{
//				//logging
//				log.Fatal( "Exception", ex );
//
//				MessageBox.Show( "An exception occurred while updating Arezzo : " + ex.Message + " : " + ex.InnerException, 
//					"MACRO Study Copy Tool", MessageBoxButtons.OK, MessageBoxIcon.Error );
//			}
//#endif
//			finally
//			{
//			}
		}

		/// <summary>
		/// Unlock macro study
		/// </summary>
		/// <param name="token"></param>
		/// <param name="clinicalTrialId"></param>
		private void UnlockStudy( string token, long clinicalTrialId )
		{
			DBLockClass l = new DBLockClass();
			
			try
			{
				l.UnlockStudy( _dbCon, token, System.Convert.ToInt32( clinicalTrialId ) );
			}
			catch
			{
			}
		}

		/// <summary>
		/// Load studies combo
		/// </summary>
		private void LoadStudies()
		{
			DataSet ds = new DataSet();

			try
			{
				//logging
				log.Debug( "Loading studies" );

				//disable form
				Processing( true );

				//get a list of studies
				ds = MACRO30.GetStudyList( _dbCon );
	
				//clear study combos
				cboStudiesS.Items.Clear();
				cboStudiesD.Items.Clear();

				//create a blank first row in source combo
				cboStudiesD.Items.Add( new ComboItem( "", "" ) );
				cboStudiesS.Items.Add( new ComboItem( "", "" ) );

				//populate the comboboxes with studies
				foreach( DataRow row in ds.Tables[0].Rows )
				{
					cboStudiesS.Items.Add( new ComboItem( row[0].ToString(), row[1].ToString() ) );
					cboStudiesD.Items.Add( new ComboItem( row[0].ToString(), row[1].ToString() ) );
				}

				//if there are any studies, select the first
				if( cboStudiesD.Items.Count > 0 )
				{
					cboStudiesS.SelectedIndex = 0;
					cboStudiesD.SelectedIndex = 0;
				}

				//logging
				log.Debug( "Study combos loaded" );
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
#endif
			finally
			{
				//enable form
				Processing( false );

				ds.Dispose();
			}
		}

		/// <summary>
		/// Destination study combo has changed
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void cboStudiesD_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			DataSet ds = new DataSet();

			try
			{
				//logging
				log.Debug( "Loading destination eforms" );

				//disable form
				Processing( true );

				//disable reports
				menuItem5.Enabled = false;

				//unlock study
				if( _lockToken != "" )
				{
					//rebuild arezzo - fix to allow 3.0.72 version to work without 3.0.75 sd module
					RebuildArezzo();

					UnlockStudy( _lockToken, _lockStudy );

					//logging
					log.Debug( _lockStudy + " study unlocked (" + _lockToken + ")" );

					_lockToken = "";
					_lockStudy = 0;

					//lock icon
					picStudyLocked.Visible = false;
				}

				//clear the eform and eform element lists
				lvwEforms.Items.Clear();
				lvwEformElements.Items.Clear();

				//clear the application state
				if (_state != null) _state.Clear();

				//reset the source study selection
				cboStudiesS.SelectedIndex = 0;

				//get the newly selected destination study
				ComboItem ciSt = ( ComboItem )cboStudiesD.SelectedItem;

				
				//only do this if study is not blank
				if( ciSt.Code != "" )
				{
					//attempt to lock the study
					_lockToken = LockStudy( System.Convert.ToInt32( ciSt.Code ) );
	
					if( _lockToken == "" )
					{
						//study cannt be locked
						MessageBox.Show( "Study '" + ciSt.Value  + "' is currently being edited by another user", "Study Locked",
 						 MessageBoxButtons.OK, MessageBoxIcon.Information );
		
						//set the selected item to nothing
						cboStudiesD.SelectedIndex = 0;

						//logging
						log.Debug( "Destination eform already locked" );
					}
					else
					{
						//remember locked study id
						_lockStudy = System.Convert.ToInt32( ciSt.Code );

						//get a list of eforms for this study
						ds = MACRO30.GetEforms( _dbConn, ciSt.Code );

						//loop through all eforms
						foreach( DataRow row in ds.Tables[0].Rows )
						{
							//add the eform as a new item
							lvwEforms.Items.Add( new StudyListViewItem( StudyCopyGlobal.ElementType.eForm, row ) );

							//add the form to the study state object
							_state.AddEform( new Eform( row["CRFPAGEID"].ToString() ) );

							//lock icon
							picStudyLocked.Visible = true;
						}

						//logging
						log.Debug( "Destination eform locked and loaded" );
					}
				}
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
#endif
			finally
			{
				//dispose of eforms dataset
				ds.Dispose();

				//reset status
				ShowProgress( "Ready", false, StudyCopyGlobal.LogPriority.Normal );

				//enable form
				Processing( false );
			}
		}

		private void Processing( bool on )
		{
			//enable/disable the form
			this.Cursor = ( on ) ? Cursors.WaitCursor : Cursors.Default;
			cboStudiesS.Enabled = !on;
			cboStudiesD.Enabled = !on;
			groupBox1.Enabled = !on;
		}

		/// <summary>
		/// Source study combo has changed
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void cboStudiesS_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			DataSet ds = new DataSet();

			try
			{
				//logging
				log.Debug( "Loading source eforms" );

				//disable form
				Processing( true );

				if( ( cboStudiesD.SelectedIndex != 0 ) && ( cboStudiesS.SelectedIndex != 0 ) )
				{
					//enable reports
					menuItem5.Enabled = true;
				}
				else
				{
					//disable reports
					menuItem5.Enabled = false;
				}

				//clear the details of any previous eform matches in the eform listview
				foreach( StudyListViewItem lvi in lvwEforms.Items )
				{
					lvi.SourceStudyElementRow = null;
					lvi.Selected = false;
				}

				//unmatch everything in the application state
				if (_state != null) _state.Unmatch();

				//clear the context menu of options
				contextMenu1.MenuItems.Clear();

				//add the 'unmatch' option and a line to the context menu
				contextMenu1.MenuItems.Add( new StudyMenuItem( StudyCopyGlobal.ElementType.eFormElement, 
					"Unmatch eForms", new EventHandler( mnuUnmatchEform_Click ) ) );
				contextMenu1.MenuItems.Add( new StudyMenuItem( StudyCopyGlobal.ElementType.eFormElement, "-", null ) );
		
				//get the newly selected study
				ComboItem ciSt = ( ComboItem )cboStudiesS.SelectedItem;
			
				//only do this if study is not blank
				if( ciSt.Code != "" )
				{
					//get a list of eforms for this study
					ds = MACRO30.GetEforms( _dbConn, ciSt.Code );

					//logging
					log.Debug( "Looping source eforms" );

					//loop through all new eforms
					foreach( DataRow row in ds.Tables[0].Rows )
					{
						//get a match if one exists
						StudyListViewItem lvi = GetMatchingEformLvi( row );

						if( lvi != null )
						{
							lvi.SourceStudyElementRow = row;

							//match the efrom in the state object
							_state.GetEform( lvi.DestinationStudyElementRow["CRFPAGEID"].ToString() ).SourceId = row["CRFPAGEID"].ToString();

							//perform automatch on all eform elements
							AutoMatchEformElements( lvi.DestinationStudyElementRow["CRFPAGEID"].ToString(), row["CRFPAGEID"].ToString());
						}
						else
						{
							//if the item was not matched up add it to the context menu
							contextMenu1.MenuItems.Add( new StudyMenuItem( StudyCopyGlobal.ElementType.eForm, row, 
								new EventHandler( mnuMatchEform_Click ) ) );
						}
					}
				}

				//logging
				log.Debug( "Source eforms loaded" );
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
#endif
			finally
			{
				//dispose of eforms dataset
				ds.Dispose();

				//reset status
				ShowProgress( "Ready", false, StudyCopyGlobal.LogPriority.Low );

				//enable form
				Processing( false );
			}
		}

		private void AutoMatchEformElements( string eformIdD, string eformIdS )
		{
			DataSet dsS = new DataSet();
			DataSet dsD = new DataSet();

			try
			{
				//logging
				log.Debug( "auto matching eform elements" );

				//get the selected studies
				ComboItem ciStS = ( ComboItem )cboStudiesS.SelectedItem;
				ComboItem ciStD = ( ComboItem )cboStudiesD.SelectedItem;

				//get eform elements for source and destination
				dsS = MACRO30.GetEformElements( _dbConn, ciStS.Code, eformIdS );
				dsD = MACRO30.GetEformElements( _dbConn, ciStD.Code, eformIdD );
				
				//get the eform state object
				Eform ef = _state.GetEform( eformIdD );

				//logging
				log.Debug( "Looping destination elements" );

				//loop through each destination element
				foreach( DataRow row in dsD.Tables[0].Rows )
				{
					//add the element to the state object
					ef.AddElement( new EformElement( row["CRFELEMENTID"].ToString() ) );
				}

				//logging
				log.Debug( "Looping source elements" );

				//loop through all source eform elements
				foreach( DataRow row in dsS.Tables[0].Rows )
				{
					//find match if one exists
					AutoMatchEformElement( dsD, row, ef );
				}

				//logging
				log.Debug( "Eform elements auto matched" );
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
#endif
			finally
			{
				//dispose of element datasets
				dsS.Dispose();
				dsD.Dispose();
			}
		}

		private void AutoMatchEformElement( DataSet dsD, DataRow rowS, Eform ef )
		{
			bool matched = false;
			short index = 0;

			try
			{
				//logging
				log.Debug( "Auto matching eform element" );

				//update status
				ShowProgress("Auto-matching source element " + rowS["CRFELEMENTID"].ToString(), false, StudyCopyGlobal.LogPriority.Low);

				//try to match with an eform element in the source study
				while( ( index < dsD.Tables[0].Rows.Count ) && ( !matched ) )
				{
					DataRow rowD = dsD.Tables[0].Rows[index];

					if( !ef.GetElementByDestination( rowD["CRFELEMENTID"].ToString() ).Matched )
					{
						//if source and destination element are of the same controltype
						if( StudyCopyGlobal.GetControlType( rowD["CONTROLTYPE"].ToString() )
							== StudyCopyGlobal.GetControlType( rowS["CONTROLTYPE"].ToString() ) 
							&& rowD["DATATYPE"].ToString() == rowS["DATATYPE"].ToString()
							)
						{
							switch( rowS["CONTROLTYPE"].ToString() )
							{
								case StudyCopyGlobal._QGROUP:
									//if controltype is 0, check to see if it is a question group, try to match on QGROUPCODE
									if( ( rowS["QGROUPID"].ToString() != "0" ) 
										&& ( rowS["QGROUPCODE"].ToString() == rowD["QGROUPCODE"].ToString() ) ) matched = true;
									break;

								case StudyCopyGlobal._LINE:
									//if element is a line, try to match on x and y
									if( ( rowS["X"].ToString() == rowD["X"].ToString() ) 
										&& ( rowS["Y"].ToString() == rowD["Y"].ToString() ) ) matched = true;
									break;

								case StudyCopyGlobal._TEXTCOMMENT:
								case StudyCopyGlobal._PICTURE:
								case StudyCopyGlobal._HOTLINK:
									//if element is comment/picture/hotlink, try to match on caption alone
									if( rowS["CAPTION"].ToString() == rowD["CAPTION"].ToString() ) matched = true;
									break;

								default:

									//if an enterable element, match up on configured match criteria
									if ( ( _settingsForm.ElementMatchOperator != "" ) && ( _settingsForm.ElementMatchCriteria2 != "" ) )
									{
										//there is a 2nd match criteria
										if( _settingsForm.ElementMatchOperator == "AND" )
										{
											//row mustnt be already matched AND both criteria must match, logical AND
											if( ( rowS[_settingsForm.ElementMatchCriteria1].ToString() != "" )
												&& ( rowD[_settingsForm.ElementMatchCriteria1].ToString() != "" )
												&& ( rowS[_settingsForm.ElementMatchCriteria2].ToString() != "" )
												&& ( rowD[_settingsForm.ElementMatchCriteria2].ToString() != "" )
												&& ( rowS[_settingsForm.ElementMatchCriteria1].ToString() 
												== rowD[_settingsForm.ElementMatchCriteria1].ToString() ) 
												&& ( rowS[_settingsForm.ElementMatchCriteria2].ToString()  
												== rowD[_settingsForm.ElementMatchCriteria2].ToString() ) 
												)
											{
												matched = true;
											}
										}
										else
										{
											//only one criteria must match, logical OR
											if( ( ( rowS[_settingsForm.ElementMatchCriteria1].ToString() != "" )
												&& ( rowD[_settingsForm.ElementMatchCriteria1].ToString() != "" )
												&& ( rowS[_settingsForm.ElementMatchCriteria1].ToString() 
												== rowD[_settingsForm.ElementMatchCriteria1].ToString() ) )
												|| 
												( ( rowS[_settingsForm.ElementMatchCriteria2].ToString() != "" )
												&& ( rowD[_settingsForm.ElementMatchCriteria2].ToString() != "" )
												&& ( rowS[_settingsForm.ElementMatchCriteria2].ToString()  
												== rowD[_settingsForm.ElementMatchCriteria2].ToString() ) )
												)
											{
												matched = true;
											}
										}
									}
									else
									{
										//there is only 1 match criteria, just compare on 1st criteria
										if( ( rowS[_settingsForm.ElementMatchCriteria1].ToString() != "" )
											&& ( rowD[_settingsForm.ElementMatchCriteria1].ToString() != "" ) 
											&& ( rowS[_settingsForm.ElementMatchCriteria1].ToString() 
											== rowD[_settingsForm.ElementMatchCriteria1].ToString() ) )
										{
											matched = true;
										}
									}
									break;
							}
						}
					}
					index++;
				}

				if( matched )
				{
					//add the match to the state object
					ef.GetElementByDestination( dsD.Tables[0].Rows[index-1]["CRFELEMENTID"].ToString() ).SourceId = rowS["CRFELEMENTID"].ToString();
				}

				//logging
				log.Debug( "Element matching complete" );
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );

//				return( lviMatched );
			}
#endif
			finally
			{
			}
		}

		/// <summary>
		/// Load and match eform elements for a pair of matched eforms
		/// </summary>
		private void LoadEformElements()
		{
			DataSet dsS = new DataSet();
			DataSet dsD = new DataSet();

			try
			{
				//logging
				log.Debug( "Loading eform elements" );

				//clear element list
				lvwEformElements.Items.Clear();
				contextMenu2.MenuItems.Clear();
				btnCopyEformElements.Enabled = false;

				//only display the element list for a single eform
				if( ( lvwEforms.SelectedItems.Count != 1 )
					|| ( ( ( StudyListViewItem )lvwEforms.SelectedItems[0] ).SourceStudyElementRow == null ) )
				{
					return;
				}

				//disable form
				Processing( true );

				//get the selected studies
				ComboItem ciStS = ( ComboItem )cboStudiesS.SelectedItem;
				ComboItem ciStD = ( ComboItem )cboStudiesD.SelectedItem;

				//get eform elements for source and destination
				dsS = MACRO30.GetEformElements( _dbConn, ciStS.Code, ( ( StudyListViewItem )lvwEforms.SelectedItems[0]).SourceStudyElementId );
				dsD = MACRO30.GetEformElements( _dbConn, ciStD.Code, ( ( StudyListViewItem )lvwEforms.SelectedItems[0]).DestinationStudyElementId );

				//add the default menu items to the eformelements context menu
				contextMenu2.MenuItems.Add( new StudyMenuItem( StudyCopyGlobal.ElementType.eFormElement,
					"Unmatch eForm elements", new EventHandler( mnuUnmatchEformElement_Click ) ) );
				contextMenu2.MenuItems.Add( new StudyMenuItem( StudyCopyGlobal.ElementType.eFormElement,
					"Add all unmatched eForm elements", new EventHandler( mnuAddEformElements_Click ) ) );
				contextMenu2.MenuItems.Add( new StudyMenuItem( StudyCopyGlobal.ElementType.eFormElement, "-", null ) );
				
				//get the eform state object
				Eform ef = _state.GetEform( ( ( StudyListViewItem )lvwEforms.SelectedItems[0]).DestinationStudyElementId );

				//logging
				log.Debug( "Looping destination elements" );

				//first add a row for each destination element
				foreach( DataRow row in dsD.Tables[0].Rows )
				{
					if( ( ( System.Convert.ToInt64( row["CONTROLTYPE"].ToString() ) > 10000 ) && ( chkCLP.Checked ) )
						|| ( ( System.Convert.ToInt64( row["CONTROLTYPE"].ToString() ) < 10000 ) && ( chkR.Checked ) ) )
					{
						lvwEformElements.Items.Add( new StudyListViewItem( StudyCopyGlobal.ElementType.eFormElement, row ) );
					}
				}

				//logging
				log.Debug( "Looping source elements" );

				//loop through all source eform elements
				foreach( DataRow row in dsS.Tables[0].Rows )
				{
					//get a match if one exists
					StudyListViewItem lvi = GetMatchingElementLvi( row, ef );

					if( lvi != null )
					{
						//add the matched source element details to the destination element row
						lvi.SourceStudyElementRow = row;

						lvi.SetIcon( ( ef.GetElementBySource( row["CRFELEMENTID"].ToString() ).Copied ) ? StudyListViewItem.ItemIcon.Copied :
							StudyListViewItem.ItemIcon.None );
					}
					else
					{
						//if the item wasnt matched, and is of a type not filtered out by the checkboxes, add to the context menu
						if( ( ( System.Convert.ToInt64( row["CONTROLTYPE"].ToString() ) > 10000 ) && ( chkCLP.Checked ) )
							|| ( ( System.Convert.ToInt64( row["CONTROLTYPE"].ToString() ) < 10000 ) && ( chkR.Checked ) ) )
						{
							contextMenu2.MenuItems.Add( new StudyMenuItem( StudyCopyGlobal.ElementType.eFormElement, row,
								new EventHandler( mnuMatchEformElement_Click ) ) );
						}
					}
				}

				//add or remove '+' row
				AddRemovePlusRow();

				//logging
				log.Debug( "Eform elements loaded" );
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
#endif
			finally
			{
				//dispose of element datasets
				dsS.Dispose();
				dsD.Dispose();

				//reset status
				ShowProgress("Ready", false, StudyCopyGlobal.LogPriority.Low);
			
				//enable form
				Processing( false );
			}
		}

		/// <summary>
		/// Get listviewitem that matches this row
		/// </summary>
		/// <param name="row"></param>
		/// <returns></returns>
		/// 28/04/2006	bug 2725 incorrect matching of eform elements
		private StudyListViewItem GetMatchingEformLvi( DataRow row )
		{
			StudyListViewItem lviMatched = null;
			short lviIndex = 0;

			try
			{
				//logging
				log.Debug( "Matching eform lvi" );

				//try to match with an eform in the source study
				while( ( lviIndex < lvwEforms.Items.Count ) && ( lviMatched == null ) )
				{
					StudyListViewItem lvi = ( StudyListViewItem )lvwEforms.Items[lviIndex];

					if ( ( _settingsForm.EformMatchOperator != "" ) && ( _settingsForm.EformMatchCriteria2 != "" ) )
					{
						//there is a 2nd match criteria
						if( _settingsForm.EformMatchOperator == "AND" )
						{
							//row mustnt be already matched AND both criteria must match, logical AND
							if( ( lvi.SourceStudyElementRow == null )
								&& (
								( row[_settingsForm.EformMatchCriteria1].ToString() != "" )
								&& ( lvi.DestinationStudyElementRow[_settingsForm.EformMatchCriteria1].ToString() != "" )
								&& ( row[_settingsForm.EformMatchCriteria2].ToString() != "" )
								&& ( lvi.DestinationStudyElementRow[_settingsForm.EformMatchCriteria2].ToString() != "" )
								&& ( row[_settingsForm.EformMatchCriteria1].ToString() 
								== lvi.DestinationStudyElementRow[_settingsForm.EformMatchCriteria1].ToString() ) 
								&& ( row[_settingsForm.EformMatchCriteria2].ToString()  
								== lvi.DestinationStudyElementRow[_settingsForm.EformMatchCriteria2].ToString() ) 
								) )
							{
								//both criteria match and none of the values being compared are ""
								lviMatched = lvi;
							}
						}
						else
						{
							//row mustnt be already matched AND only one criteria must match, logical OR
							if( ( lvi.SourceStudyElementRow == null )
								&& (
								( ( row[_settingsForm.EformMatchCriteria1].ToString() != "" )
								&& ( lvi.DestinationStudyElementRow[_settingsForm.EformMatchCriteria1].ToString() != "" )
								&& ( row[_settingsForm.EformMatchCriteria1].ToString() 
								== lvi.DestinationStudyElementRow[_settingsForm.EformMatchCriteria1].ToString() ) )
								|| 
								( ( row[_settingsForm.EformMatchCriteria2].ToString() != "" )
								&& ( lvi.DestinationStudyElementRow[_settingsForm.EformMatchCriteria2].ToString() != "" )
								&& ( row[_settingsForm.EformMatchCriteria2].ToString()  
								== lvi.DestinationStudyElementRow[_settingsForm.EformMatchCriteria2].ToString() ) )
								) )
							{
								//one of the criteria match and the values being matched are not ""
								lviMatched = lvi;
							}
						}
					}
					else
					{
						//row mustnt be already matched AND 
						//there is only 1 match criteria, just compare on 1st criteria
						if( ( lvi.SourceStudyElementRow == null )
								&& (
							( row[_settingsForm.EformMatchCriteria1].ToString() != "" )
							&& ( lvi.DestinationStudyElementRow[_settingsForm.EformMatchCriteria1].ToString() != "" ) 
							&& ( row[_settingsForm.EformMatchCriteria1].ToString() 
							== lvi.DestinationStudyElementRow[_settingsForm.EformMatchCriteria1].ToString() ) ) )
						{
							//the first criteria matched and the values compared are not ""
							lviMatched = lvi;
						}
					}

					lviIndex++;
				}

				//logging
				log.Debug( "Eform lvi matching complete" );

				return( lviMatched );
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );

				return( lviMatched );
			}
#endif
			finally
			{
			}
		}

		/// <summary>
		/// Get listviewitem that matches this row
		/// </summary>
		/// <param name="row"></param>
		/// <param name="ef"></param>
		/// <returns></returns>
		/// 28/04/2006	bug 2725 incorrect matching of eform elements
		/// 28/04/2006  bug 2726 incorrect matching and display of rqg elements
		private StudyListViewItem GetMatchingElementLvi( DataRow row, Eform ef )
		{
			StudyListViewItem lviMatched = null;
			short lviIndex = 0;

			try
			{
				//logging
				log.Debug( "Matching element lvi" );

				//read state from state object
				EformElement el = ef.GetElementBySource( row["CRFELEMENTID"].ToString() );
				
				if( el != null )
				{
					//this element is matched in the state object
					while( ( lviIndex < lvwEformElements.Items.Count ) && ( lviMatched == null ) )
					{
						StudyListViewItem lvi = ( StudyListViewItem )lvwEformElements.Items[lviIndex];
						if( lvi.DestinationStudyElementRow["CRFELEMENTID"].ToString() == el.DestinationId ) lviMatched = lvi;
						lviIndex++;
					}
				}

				//logging
				log.Debug( "Element lvi matching complete" );

				return( lviMatched );
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );

				return( lviMatched );
			}
#endif
			finally
			{
			}
		}

		private bool MatchWarningAccepted( string dataItemCodeD, string dataItemCodeS, string dataTypeS, string dataTypeD )
		{
			//warn if datatypes are different
			if( dataTypeS != dataTypeD )
			{
				MessageBox.Show( "You are attempting to match questions of different datatypes:\n\n"
					+	dataItemCodeS + " (" + StudyCopyGlobal.GetDataType( dataTypeS ) + ") -> " 
					+	dataItemCodeD + " (" + StudyCopyGlobal.GetDataType( dataTypeD ) + ")"
					+ "\n\nDatatypes cannot be converted automatically by this tool.",
					"DataType Mismatch", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			}

			return( false );
		}

		/// <summary>
		/// Initialise listview controls
		/// </summary>
		private void InitListView()
		{
			//eforms listview
			lvwEforms.View = View.Details;
			lvwEforms.CheckBoxes = false;
			lvwEforms.GridLines = true;
			lvwEforms.MultiSelect = true;
			lvwEforms.FullRowSelect = true;
			lvwEforms.Columns.Add("", 15, HorizontalAlignment.Left);
			lvwEforms.Columns.Add("Dest.PageId", 50, HorizontalAlignment.Left);
			lvwEforms.Columns.Add("Dest.Title", 120, HorizontalAlignment.Left);
			lvwEforms.Columns.Add("Dest.Label", 130, HorizontalAlignment.Left);
			lvwEforms.Columns.Add("Src.PageId", 50, HorizontalAlignment.Left);
			lvwEforms.Columns.Add("Src.Title", 120, HorizontalAlignment.Left);
			lvwEforms.Columns.Add("Src.Label", 130, HorizontalAlignment.Left);


			//eform elements listview
			lvwEformElements.View = View.Details;
			lvwEformElements.CheckBoxes = false;
			lvwEformElements.GridLines = true;
			lvwEformElements.MultiSelect = true;
			lvwEformElements.FullRowSelect = true;
			lvwEformElements.Columns.Add("", 15, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Dest.ElementId", 50, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Dest.Caption", 120, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Dest.Control Type", 75, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Dest.DataItem Code", 50, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Dest.DataItem Name", 130, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Dest.Data Type", 60, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Dest.xy", 40, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Dest.QGroup Id", 40, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Src.ElementId", 50, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Src.Caption", 120, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Src.Control Type", 75, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Src.DataItem Code", 50, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Src.DataItem Name", 130, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Src.Data Type", 60, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Src.xy", 40, HorizontalAlignment.Left);
			lvwEformElements.Columns.Add("Src.QGroup Id", 40, HorizontalAlignment.Left);
		}

		/// <summary>
		/// Filter on response checkbox has changed
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void chkR_CheckedChanged(object sender, System.EventArgs e)
		{
			LoadEformElements();
		}

		/// <summary>
		/// Filter on comment/line/picture checkbox has changed
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void chkCLP_CheckedChanged(object sender, System.EventArgs e)
		{
			LoadEformElements();
		}

		/// <summary>
		/// Copy eform properties
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnCopyEform_Click(object sender, System.EventArgs e)
		{
			DataSet ds = new DataSet();

			try
			{
				if( _copyForm.ShowDialog( StudyCopyGlobal.ElementType.eForm ) == DialogResult.Yes )
				{
					if( _copyForm.ElementCopyProperties.Count > 0 )
					{
						//logging
						log.Debug( "Copying eforms" );

						//disable form
						Processing( true );

						foreach( StudyListViewItem lvi in lvwEforms.Items )
						{
							if( ( lvi.Selected ) && ( lvi.MatchedElement ) )
							{
								//get the selected studies
								ComboItem ciStS = ( ComboItem )cboStudiesS.SelectedItem;
								ComboItem ciStD = ( ComboItem )cboStudiesD.SelectedItem;

								//copy eform properties
								MACRO30.CopyEform( _dbConn, ciStS.Code, ciStD.Code, lvi.SourceStudyElementRow, 
									lvi.DestinationStudyElementRow, _copyForm.ElementCopyProperties );
						
								//get the updated row from the db
								ds = MACRO30.GetEforms( _dbConn, ciStD.Code, lvi.DestinationStudyElementRow["CRFPAGEID"].ToString() );
								if( ds.Tables[0].Rows.Count == 1 )
								{
									//update listview row
									lvi.DestinationStudyElementRow = ds.Tables[0].Rows[0];

									//set eform as copied in state object
									_state.GetEform( lvi.DestinationStudyElementRow["CRFPAGEID"].ToString() ).Copied = true;

									//set eform icon
									lvi.SetIcon( StudyListViewItem.ItemIcon.Copied );
								}
							}
						}

						//logging
						log.Debug( "Eforms copied" );
					}
				}
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
#endif
			finally
			{
				//enable form
				Processing( false );

				//reset status
				ShowProgress("Ready", false, StudyCopyGlobal.LogPriority.Normal);
			}
		}

		/// <summary>
		/// Match up eforms
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void mnuMatchEform_Click(object sender, System.EventArgs e)
		{
			try
			{
				//logging
				log.Debug( "Matching eforms" );

				//remove any already matching eform
				UnmatchEform();

				//get the menu item
				StudyMenuItem m = ( StudyMenuItem ) sender;

				//remove the item from the menu as it is now matched
				contextMenu1.MenuItems.Remove( m );
			
				//add the eform details to the listview
				( ( StudyListViewItem )lvwEforms.SelectedItems[0] ).SourceStudyElementRow = m.StudyElementRow;

				//logging
				log.Info( "Matching eforms:"
					+ " study:" + ( ( ComboItem )cboStudiesS.SelectedItem ).Code + "->" + ( ( ComboItem )cboStudiesS.SelectedItem ).Code 
					+ " eform:" + m.StudyElementRow["CRFPAGEID"].ToString() + "->" + ( ( StudyListViewItem )lvwEforms.SelectedItems[0] ).DestinationStudyElementRow["CRFPAGEID"].ToString() );

				AutoMatchEformElements( ( ( StudyListViewItem )lvwEforms.SelectedItems[0] ).DestinationStudyElementRow["CRFPAGEID"].ToString(),
					m.StudyElementRow["CRFPAGEID"].ToString() );

				//load any elements
				LoadEformElements();

				//logging
				log.Debug( "Eforms matched" );
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
#endif
			finally
			{
			}
		}

		/// <summary>
		/// Unmatch eform element source/destination
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void mnuUnmatchEformElement_Click(object sender, System.EventArgs e)
		{
			if( MessageBox.Show( "Are you sure you want to unmatch the selected item(s)?", "Unmatch Items", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question ) == DialogResult.Yes )
			{
				UnmatchEformElement();
			}
		}

		/// <summary>
		/// Unmatch eform elements
		/// </summary>
		private void UnmatchEformElement()
		{
			try
			{
				//logging
				log.Debug( "Unmatching elements" );

				//get the eform state object
				Eform ef = _state.GetEform( ( ( StudyListViewItem )lvwEforms.SelectedItems[0]).DestinationStudyElementId );

				foreach( StudyListViewItem lvi in lvwEformElements.Items )
				{
					if( ( lvi.Selected ) && ( lvi.MatchedElement ) )
					{
						//logging
						log.Info( "Unmatching element:"
							+ " study:" + ( ( ComboItem )cboStudiesS.SelectedItem ).Code + "->" + ( ( ComboItem )cboStudiesS.SelectedItem ).Code 
							+ " eform:" + ef.SourceId + "->" + ef.DestinationId
							+ " element:" + lvi.SourceStudyElementRow["CRFELEMENTID"].ToString() + "->" + lvi.DestinationStudyElementRow["CRFELEMENTID"].ToString() );

						//add to context menu
						contextMenu2.MenuItems.Add( new StudyMenuItem( StudyCopyGlobal.ElementType.eFormElement,
							lvi.SourceStudyElementRow, new EventHandler( mnuMatchEformElement_Click ) ) );
					
						//unmatch the element in the listview
						lvi.SourceStudyElementRow = null;

						//unmatch the element in the state object
						ef.GetElementByDestination( lvi.DestinationStudyElementRow["CRFELEMENTID"].ToString() ).UnMatch();

						//add or remove '+' row
						AddRemovePlusRow();
					}
				}

				//logging
				log.Debug( "Elements unmatched" );
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
#endif
			finally
			{
			}
		}

		/// <summary>
		/// Add all unmatched elements
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void mnuAddEformElements_Click(object sender, System.EventArgs e)
		{
			//match the eform element
			MatchEformElements( -1 );
		}

		/// <summary>
		/// Match up or add a new eform element
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void mnuMatchEformElement_Click(object sender, System.EventArgs e)
		{
			//match the eform element
			MatchEformElements( ( ( StudyMenuItem )sender ).Index );
		}

		/// <summary>
		/// Check for questionable question matches and warn user
		/// </summary>
		/// <param name="menuIndex"></param>
		/// <returns></returns>
		private bool MatchWarningAccepted( int menuIndex )
		{
			//get context menu item - source
			StudyMenuItem m = ( StudyMenuItem )contextMenu2.MenuItems[menuIndex];

			//get lvi item - destination
			StudyListViewItem lvi = ( StudyListViewItem )lvwEformElements.SelectedItems[0];

			//if adding a new row, dont need to check
			if( lvi.DestinationStudyElementRow == null )
			{
				return( true );
			}

			return( MatchWarningAccepted( m.StudyElementRow["DATAITEMCODE"].ToString(), lvi.DestinationStudyElementRow["DATAITEMCODE"].ToString(),
				m.StudyElementRow["DATATYPE"].ToString(), lvi.DestinationStudyElementRow["DATATYPE"].ToString() ) );
		}

		/// <summary>
		/// Match or add elements
		/// </summary>
		/// <param name="menuIndex"></param>
		private void MatchEformElements( int menuIndex )
		{
			bool asked = false;
			string question = "";
			int startIndex = 3, endIndex = contextMenu2.MenuItems.Count - 1;
			DialogResult r = DialogResult.No;

			try
			{
				//logging
				log.Debug( "Matching elements" );

				//disable form
				Processing( true );

				//> 0  means matching or adding single element. -1 means adding element(s)
				if( menuIndex > 0 )
				{
					if( !MatchWarningAccepted( menuIndex ) ) return;
					
					//selected element
					startIndex = menuIndex;
					endIndex = menuIndex;

					//remove any already matching element
					UnmatchEformElement();
				}

				//process all context menu items within the bounds passed, starting at the end
				for( int n = endIndex; n >= startIndex; n-- )
				{
					//get context menu item - source
					StudyMenuItem m = ( StudyMenuItem )contextMenu2.MenuItems[n];

					StudyListViewItem lvi = null;
					if( menuIndex > 0 )
					{
						//get the selected listview row - destination
						lvi = ( StudyListViewItem )lvwEformElements.SelectedItems[0];
					}
					else
					{
						//get the '+' row - destination
						lvi = ( StudyListViewItem )lvwEformElements.Items[lvwEformElements.Items.Count - 1];
					}

					//get the eform state object so we can update it when needed
					Eform ef = _state.GetEform( ( ( StudyListViewItem )lvwEforms.SelectedItems[0]).DestinationStudyElementId );

					//if there is currently no destination element in this listview row - adding an element
					if( lvi.DestinationStudyElementRow == null )
					{
						if( !asked )
						{
							//if we havent already asked the user if theyre sure they want to add element(s)
							question = "Are you sure you want to add ";
							question += ( menuIndex > 0 ) ? "the element " + m.StudyElementRow["CRFELEMENTID"].ToString() + "?" 
								: "all unmatched elements? (" + ( endIndex - startIndex + 1 ) + " items)";
							r = MessageBox.Show( question, "Add Element", MessageBoxButtons.YesNo, MessageBoxIcon.Question );
						
							//remember we have asked the user
							asked = true;
						}

						if( r == DialogResult.Yes )
						{
							//add the element
							DataRow newRow = AddNewElement( lvi, m.StudyElementRow );

							//if the element was added successfully
							if( newRow != null )
							{
								//add the element row to the listview item
								lvi.DestinationStudyElementRow = newRow;

								//mark row as copied
								lvi.SetIcon( StudyListViewItem.ItemIcon.Copied );

								//add the new element to the state object and set to copied
								EformElement el = new EformElement( lvi.DestinationStudyElementRow["CRFELEMENTID"].ToString() );
								ef.AddElement( el );
								el.Copied = true;
							}
						}
					}

					if( lvi.DestinationStudyElementRow != null )
					{
						//remove the item from the menu as it is now matched
						contextMenu2.MenuItems.Remove( m );

						//logging
						log.Info( "Matching elements:"
							+ " study:" + ( ( ComboItem )cboStudiesS.SelectedItem ).Code + "->" + ( ( ComboItem )cboStudiesS.SelectedItem ).Code 
							+ " eform:" + ef.SourceId + "->" + ef.DestinationId 
							+ " elements:" + m.StudyElementRow["CRFELEMENTID"].ToString() + "->" + lvi.DestinationStudyElementRow["CRFELEMENTID"].ToString() );
			
						//set the source row
						lvi.SourceStudyElementRow = m.StudyElementRow;

						//add match to state object
						ef.GetElementByDestination( lvi.DestinationStudyElementRow["CRFELEMENTID"].ToString() ).SourceId = lvi.SourceStudyElementRow["CRFELEMENTID"].ToString();

						//add or remove '+' row
						AddRemovePlusRow();
					}
				}

				//logging
				log.Debug( "Elements matched" );
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
#endif
			finally
			{
				//enable form
				Processing( false );

				//reset status
				ShowProgress("Ready", false, StudyCopyGlobal.LogPriority.Normal);
			}
		}

		/// <summary>
		/// Add or remove the '+' row
		/// </summary>
		private void AddRemovePlusRow()
		{
			try
			{
				if( lvwEformElements.Items.Count > 0 )
				{
					//if there are any elements in the list
					if( lvwEformElements.Items[lvwEformElements.Items.Count -1].ImageIndex == StudyListViewItem._ADD )
					{
						//if there is currently an '+' row
						if( contextMenu2.MenuItems.Count <= 3 )
						{
							//remove '+' row
							lvwEformElements.Items.Remove( lvwEformElements.Items[lvwEformElements.Items.Count -1] );
						}
					}
					else
					{
						//there is not currently a '+' row
						if ( contextMenu2.MenuItems.Count > 3 )
						{
							//add '+' row
							lvwEformElements.Items.Add( new StudyListViewItem( StudyCopyGlobal.ElementType.eFormElement, 
								StudyListViewItem.ItemIcon.Add ) );
						}
					}
				}
				else
				{
					//if there are no elements in the list
					if ( contextMenu2.MenuItems.Count > 3 )
					{
						//add '+' row
						lvwEformElements.Items.Add( new StudyListViewItem( StudyCopyGlobal.ElementType.eFormElement, 
							StudyListViewItem.ItemIcon.Add ) );
					}
				}
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
#endif
			finally
			{
			}
		}

		/// <summary>
		/// Add a new element to the destination study
		/// </summary>
		/// <param name="lvi"></param>
		/// <param name="rowS"></param>
		/// <returns></returns>
		private DataRow AddNewElement( StudyListViewItem lvi, DataRow rowS )
		{
			DataRow newRow = null;
			string qGroupCode = "", qGroupIdD = "0";

			try
			{
				//logging
				log.Debug( "Adding element" );

				//get the selected studies
				ComboItem ciStS = ( ComboItem )cboStudiesS.SelectedItem;
				ComboItem ciStD = ( ComboItem )cboStudiesD.SelectedItem;
	
				//we are adding an element
				string dataItemId = "0", dataItemCode = "";

				if( ( rowS["DATAITEMID"].ToString() != "0" ) && ( rowS["DATAITEMID"].ToString() != "" ) )
				{
					//this is an enterable element, not a line/comment/picture
					DialogResult r = DialogResult.No;

					if( MACRO30.DataItemCodeExists( _dbConn, ciStD.Code, rowS["DATAITEMCODE"].ToString() ) )
					{
						//if the dataitemcode already exists in the destination study - the user cant add it and 
						//must choose a dataitem from the dataitems already in the study
						MessageBox.Show( "The DATAITEMCODE of the dataitem associated with the element being added already exists\n"
							+ "in the destination study. In order to add this element you must associate the element with\n"
							+ "an existing dataitem ", "DATAITEMCODE '" + rowS["DATAITEMCODE"].ToString() + "' Already Exists", MessageBoxButtons.OK, 
							MessageBoxIcon.Information );

						//remember the dataitemcode so we can highlight it in the selection list
						dataItemCode = rowS["DATAITEMCODE"].ToString();
					}
					else
					{
						//the dataitemcode doesnt exist in the destination study - the user can add it or choose 
						//one from those already in the study
						r = MessageBox.Show( "Do you want to create a new dataitem for added element " + rowS["CRFELEMENTID"].ToString() 
							+ "?", "Create DataItem", MessageBoxButtons.YesNo, MessageBoxIcon.Question );
					}
				
					if( r == DialogResult.No )
					{
						//get user choice of dataitem
						DataItemsForm f = new DataItemsForm( _dbConn, ciStD.Code, rowS["DATATYPE"].ToString(), dataItemCode );
						f.ShowDialog();

						//remember the selected dataitemid
						dataItemId = f.SelectedDataItemId;
						f.Dispose();

						if( dataItemId == "0" )
						{
							//logging
							log.Debug( "User failed to select a valid dataitem" );

							//the user cancelled out of the dataitem choice form - exit from the function
							return( newRow );
						}
					}

					if( dataItemId == "0" )
					{
						//insert a new dataitem
						dataItemId = MACRO30.InsertDataItem( _dbConn, ciStS.Code, ciStD.Code, rowS, true, ciStS.Value );
					
						if( rowS["DATATYPE"].ToString() == StudyCopyGlobal._CATEGORY )
						{
							//element is category question - add category values
							MACRO30.InsertCategoryValues( _dbConn, ciStS.Code, ciStD.Code, rowS, dataItemId );
						}
					}
				}

				if( ( rowS["QGROUPID"].ToString() != "0" ) && ( rowS["QGROUPID"].ToString() != "" ) )
				{
					//if we are trying to add a question group - check that the qgroupcode isnt already taken
					if( MACRO30.QGroupCodeExists( _dbConn, ciStD.Code, ciStS.Code, rowS["QGROUPID"].ToString(), out qGroupCode, out qGroupIdD ) )
					{
						//this code is already taken so the question group cannot be added
						MessageBox.Show( "The QGROUPCODE of the QGROUP associated with the element already exists in the\n"
							+ "destination study and cannot be added.", "QGROUPCODE '" + qGroupCode + "' Already Exists", 
							MessageBoxButtons.OK, MessageBoxIcon.Information );
					}
					else
					{
						//insert a new question group, it will have no questions attached
						MACRO30.InsertQGroup( _dbConn, ciStD.Code, ciStS.Code, rowS, out qGroupIdD ); 
					}

					if( MACRO30.EformQGroupExists( _dbConn, ciStD.Code, 
						( ( StudyListViewItem )lvwEforms.SelectedItems[0] ).DestinationStudyElementId, qGroupIdD ) )
					{
						//this eform group is already on the eform - cannot be added
						MessageBox.Show( "The EFORMQGROUP already exists on the destination eform and cannot be added.", 
							"EFORMQGROUP '" + qGroupIdD + "' Already Exists", MessageBoxButtons.OK, MessageBoxIcon.Information );

						//exit
						return( newRow );
					}
				}

				//insert the element
				newRow = MACRO30.InsertEformElement( _dbConn, rowS, ciStD.Code, ciStS.Code, 
					( ( StudyListViewItem )lvwEforms.SelectedItems[0] ).DestinationStudyElementId, dataItemId, qGroupIdD );
			

				//logging
				log.Debug( "Element added" );

				//return the new element
				return( newRow );
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );

				return( newRow );
			}
#endif
			finally
			{
			}
		}	

		/// <summary>
		/// Unmatch eForm source/destination row
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void mnuUnmatchEform_Click(object sender, System.EventArgs e)
		{
			if( MessageBox.Show( "Are you sure you want to unmatch the selected item(s)?", "Unmatch Items", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question ) == DialogResult.Yes )
			{
				//unmatch the eform
				UnmatchEform();

				//reload any elements
				LoadEformElements();
			}
		}

		/// <summary>
		/// Unmatch eForm source/destination row
		/// </summary>
		private void UnmatchEform()
		{
			try
			{
				//logging
				log.Debug( "Unmatching eform" );

				foreach( StudyListViewItem lvi in lvwEforms.Items )
				{
					if( ( lvi.Selected ) && ( lvi.MatchedElement ) )
					{
						//logging
						log.Info( "Unmatching eform:"
							+ " study:" + ( ( ComboItem )cboStudiesS.SelectedItem ).Code + "->" + ( ( ComboItem )cboStudiesS.SelectedItem ).Code 
							+ " eform:" + lvi.SourceStudyElementRow["CRFPAGEID"].ToString() + "->" + lvi.DestinationStudyElementRow["CRFPAGEID"].ToString() );

						//add to context menu
						contextMenu1.MenuItems.Add( new StudyMenuItem( StudyCopyGlobal.ElementType.eForm,
							lvi.SourceStudyElementRow, new EventHandler( mnuMatchEform_Click ) ) );
					
						//unmatch in the listview
						lvi.SourceStudyElementRow = null;

						//unmatch in the state object
						_state.GetEform( lvi.DestinationStudyElementRow["CRFPAGEID"].ToString() ).UnMatch();
					}
				}

				//logging
				log.Debug( "Eform unmatched" );
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
#endif
			finally
			{
			}
		}

		/// <summary>
		/// Popup eForm context menu
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void contextMenu1_Popup(object sender, System.EventArgs e)
		{
			//unmatch option
			contextMenu1.MenuItems[0].Enabled = ( lvwEforms.SelectedItems.Count > 0 ) ? true : false;

			//element options
			for( int n = 2; n < contextMenu1.MenuItems.Count; n++ )
			{
				contextMenu1.MenuItems[n].Enabled = ( lvwEforms.SelectedItems.Count == 1 ) ? true : false;
			}
		}

		/// <summary>
		/// Selected item in eforms listview has changed
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void lvwEforms_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if( lvwEforms.SelectedItems.Count > 0 )
			{
				btnCopyEform.Enabled = true;
			}
			else
			{
				btnCopyEform.Enabled = false;
			}
		
			LoadEformElements();
		}

		/// <summary>
		/// Exit application
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void menuItem2_Click(object sender, System.EventArgs e)
		{
			Application.Exit();
		}

		/// <summary>
		/// Eform element menu has popped up
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void contextMenu2_Popup(object sender, System.EventArgs e)
		{
			if( contextMenu2.MenuItems.Count > 0 )
			{
				//unmatch option
				contextMenu2.MenuItems[0].Enabled = ( lvwEformElements.SelectedItems.Count > 0 );

				//add all unmatched items option
				contextMenu2.MenuItems[1].Enabled = ( contextMenu2.MenuItems.Count > 3 );

				//element options
				for( int n = 3; n < contextMenu2.MenuItems.Count; n++ )
				{
					contextMenu2.MenuItems[n].Enabled = ( lvwEformElements.SelectedItems.Count == 1 ) ? true : false;
				}
			}
		}

		/// <summary>
		/// About dialog
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void menuItem4_Click(object sender, System.EventArgs e)
		{
			MessageBox.Show("(c) " + Application.CompanyName + "\n" + "MACRO Study Copy Tool " 
				+ Application.ProductVersion, "MACRO Study Copy Tool", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		/// <summary>
		/// Copy eform element properties
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnCopyEformElements_Click(object sender, System.EventArgs e)
		{
			DataSet ds = new DataSet();

			try
			{
				if( _copyForm.ShowDialog( StudyCopyGlobal.ElementType.eFormElement ) == DialogResult.Yes )
				{
					//logging
					log.Debug( "Copying elements" );

					//disable form
					Processing( true );

					//get the eform state object
					Eform ef = _state.GetEform( ( ( StudyListViewItem )lvwEforms.SelectedItems[0]).DestinationStudyElementId );

					foreach( StudyListViewItem lvi in lvwEformElements.Items )
					{
						if( ( lvi.Selected ) && ( lvi.MatchedElement ) )
						{
							//get the selected studies
							ComboItem ciStS = ( ComboItem )cboStudiesS.SelectedItem;
							ComboItem ciStD = ( ComboItem )cboStudiesD.SelectedItem;

							if( !_state.DataItemCopied( lvi.DestinationStudyElementRow["DATAITEMID"].ToString() ) 
								&& ( ( _copyForm.DataItemCopyProperties.Count > 0 ) || (_copyForm.DataItemValidation) ) )
							{
								//dataitem hasnt been copied yet - copy dataitem
								MACRO30.CopyDataItem( _dbConn, ciStS.Code, ciStD.Code, lvi.SourceStudyElementRow, 
									lvi.DestinationStudyElementRow, _copyForm.DataItemCopyProperties, _copyForm.DataItemValidation );
							}

							if( ( !_state.QGroupCopied( lvi.DestinationStudyElementRow["QGROUPID"].ToString() ) ) 
								&& ( _copyForm.QGCopyProperties.Count > 0 ) )
							{
								//qgroup hasnt been copied yet - copy qgroup
								MACRO30.CopyQGroup( _dbConn, ciStD.Code, ciStS.Code, lvi.SourceStudyElementRow, lvi.DestinationStudyElementRow,
									_copyForm.QGCopyProperties );
							}

							if( ( lvi.SourceStudyElementRow["QGROUPID"].ToString() != "0" ) 
								&& ( lvi.SourceStudyElementRow["QGROUPID"].ToString() != "" ) 
								&& ( _copyForm.EformQGCopyProperties.Count > 0 ) )
							{
								//this element is a question group - copy eform group properties
								MACRO30.CopyEformQGroup( _dbConn, ciStD.Code, ciStS.Code, lvi.SourceStudyElementRow, lvi.DestinationStudyElementRow,
									_copyForm.EformQGCopyProperties );
							}

							if( _copyForm.ElementCopyProperties.Count > 0 )
							{
								//copy element
								MACRO30.CopyEformElement( _dbConn, ciStD.Code, ciStS.Code, lvi.SourceStudyElementRow, 
									lvi.DestinationStudyElementRow, _copyForm.ElementCopyProperties );
							}

							//get the new row from the db
							ds = MACRO30.GetEformElements( _dbConn, ciStD.Code, lvi.DestinationStudyElementRow["CRFPAGEID"].ToString(), 
								lvi.DestinationStudyElementRow["CRFELEMENTID"].ToString() );
							if( ds.Tables[0].Rows.Count == 1 )
							{
								//set the listviewitem datarow to the newly copied row
								lvi.DestinationStudyElementRow = ds.Tables[0].Rows[0];

								//set the listview icon to copied
								lvi.SetIcon( StudyListViewItem.ItemIcon.Copied );

								//set the element as copied in the state object
								ef.GetElementByDestination( lvi.DestinationStudyElementRow["CRFELEMENTID"].ToString() ).Copied = true;
							}
						}
					}

					//logging
					log.Debug( "Elements copied" );
				}
			}
#if( !DEBUG )
			catch (Exception ex)
			{
				//logging
				log.Error( "Error", ex );

				//catch any exception
				MessageBox.Show("An exception occurred :\n" + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
#endif
			finally
			{
				ds.Dispose();

				//enable form
				Processing( false );

				//reset status
				ShowProgress("Ready", false, StudyCopyGlobal.LogPriority.Normal);
			}
		}

		/// <summary>
		/// Eform element row selection has changed
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void lvwEformElements_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if( lvwEformElements.SelectedItems.Count > 0 )
			{
				//enable copy button
				btnCopyEformElements.Enabled = true;
			}
			else
			{
				//disable copy button
				btnCopyEformElements.Enabled = false;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void menuItem8_Click(object sender, System.EventArgs e)
		{
			//show the settings dialog
			_settingsForm.ShowDialog();
		}

		/// <summary>
		/// get ctrl-a and select all
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void lvwEformElements_onKeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if((e.Control)&&(e.KeyCode == Keys.A))
			{
				// Select all elements
				for( int n = 0; n < lvwEformElements.Items.Count; n++)
				{
					lvwEformElements.Items[n].Selected = true;
				}
			}
		}

		/// <summary>
		/// get ctrl-a and select all
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void lvwEforms_onKeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if((e.Control)&&(e.KeyCode == Keys.A))
			{
				// Select all elements
				for( int n = 0; n < lvwEforms.Items.Count; n++)
				{
					lvwEforms.Items[n].Selected = true;
				}
			}
		}

		/// <summary>
		/// select all eform elements
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSelectAllElements_Click(object sender, System.EventArgs e)
		{
			// Select all elements
			for( int n = 0; n < lvwEformElements.Items.Count; n++)
			{
				lvwEformElements.Items[n].Selected = true;
			}
		}

		private void MainForm_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			Processing( true );

			//unlock study
			if( _lockToken != "" )
			{
				//rebuild arezzo - fix to allow 3.0.72 version to work without 3.0.75 sd module
				RebuildArezzo();

				UnlockStudy( _lockToken, _lockStudy );

				//logging
				log.Debug( _lockStudy + " study unlocked (" + _lockToken + ")" );

				_lockToken = "";
				_lockStudy = 0;
			}

			//dispose of global resources
			if (_dbConn != null) 
			{
				_dbConn.Close();
				_dbConn.Dispose();
				_settingsForm.Dispose();
				_copyForm.Dispose();
			}

			Processing( false );
		}

		private void menuItem5_Click(object sender, System.EventArgs e)
		{
			ToolsForm f = new ToolsForm( _dbConn, ( ( ComboItem )cboStudiesD.SelectedItem ).Code,
				( ( ComboItem )cboStudiesD.SelectedItem ).Value,
				( ( ComboItem )cboStudiesS.SelectedItem ).Code, ( ( ComboItem )cboStudiesS.SelectedItem ).Value, _state );
				
			f.ShowDialog();
			f.Dispose();
		}

		private void MainForm_Load(object sender, System.EventArgs e)
		{
			IDbConnection secDbConn = null;

			try
			{
	
				if( PrevInstance() )
				{
					MessageBox.Show( "A previous instance of this application was detected.", 
						"InferMed MACRO Study Copy", MessageBoxButtons.OK, MessageBoxIcon.Stop );
					this.Close();
				}
				else
				{
					imedLogin1.MandatoryMACROLogin = true;
					imedLogin1.MandatorySecurityDB = true;
					imedLogin1.SecurityDbLoginOnly = false;

					if( imedLogin1.Login( _SETTINGS_FILE ) )
					{
						//logging
						log.Info( "-----------------------------------------------------------------------");
						log.Info( "Application starting up. Database " + imedLogin1.DatabaseCode + " User " + imedLogin1.UsernameFull );

						//set properties
						_secCon = IMEDEncryption.DecryptString(imedLogin1.SecDBConnectionString);
						_dbCon = IMEDEncryption.DecryptString(imedLogin1.DBConnectionString);
						_dbCode = imedLogin1.DatabaseCode;
						_userName = imedLogin1.Username;
						_userNameFull = imedLogin1.UsernameFull;
						_MACRORole = imedLogin1.MACRORole;

						//open a connection object
						switch( IMEDDataAccess.CalculateConnectionType( _secCon ) )
						{
							case IMEDDataAccess.ConnectionType.Oracle:
								_secCon = IMEDDataAccess.FormatConnectionString(IMEDDataAccess.ConnectionType.Oracle, _secCon);
								secDbConn = new OracleConnection( _secCon );
								break;
							case IMEDDataAccess.ConnectionType.SQLServer:
								_secCon = IMEDDataAccess.FormatConnectionString(IMEDDataAccess.ConnectionType.SQLServer, _secCon);
								secDbConn = new SqlConnection( _secCon );
								break;
						}
						secDbConn.Open();

						if (!MACRO30.CheckPermission(secDbConn, _MACRORole, _FN_STUDYDEFINITION))
						{
							Exception ex = new Exception("User does not have permission to access module");
							throw (ex);
						}

						_settingsForm = new SettingsForm( "" );
						_copyForm = new ConfirmCopyForm( "" );
						_state = new StudyState( "" );

						//set interface caption
						this.Text = "MACRO Study Copy Tool [" + imedLogin1.DatabaseCode + "]";

						string connectionString = "";

						//logging
						log.Debug( "Initialising application" );

						ShowProgress( "Initialising...", false, StudyCopyGlobal.LogPriority.Normal );

						//create listview headers
						InitListView();

						//set callback events
						MACRO30.DShowProgress DProg = new MACRO30.DShowProgress( ShowProgress );
						MACRO30.DShowProgressEvent += DProg;
				
						//open a connection object
						switch( IMEDDataAccess.CalculateConnectionType( _dbCon ) )
						{
							case IMEDDataAccess.ConnectionType.Oracle:
								connectionString = IMEDDataAccess.FormatConnectionString(IMEDDataAccess.ConnectionType.Oracle, _dbCon);
								_dbConn = new OracleConnection( connectionString );
								break;
							case IMEDDataAccess.ConnectionType.SQLServer:
								connectionString = IMEDDataAccess.FormatConnectionString(IMEDDataAccess.ConnectionType.SQLServer, _dbCon);
								_dbConn = new SqlConnection( connectionString );
								break;
						}
						_dbConn.Open();

						//load studies combos
						LoadStudies();
				
						//reset status
						ShowProgress( "Ready", false, StudyCopyGlobal.LogPriority.Normal );

						//logging
						log.Debug( "Application initialisation complete" );
					}
					else
					{
						this.Close();
					}
				}
			}
#if (!DEBUG)
			catch( Exception ex )
			{
				MessageBox.Show( "An exception occurred while initialising : " + ex.Message + " : " + ex.InnerException, 
					"InferMed MACRO Study Copy", MessageBoxButtons.OK, MessageBoxIcon.Error );
				Application.Exit();
			}
#endif
			finally
			{
				if (secDbConn != null)
				{
					//logging
					log.Debug( "Security connection closed" );

					secDbConn.Close();
					secDbConn.Dispose();

					imedLogin1.Dispose();
				}
			}
		}
	}
}
