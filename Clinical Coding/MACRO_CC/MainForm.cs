using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using InferMed.Components;
using System.Diagnostics;
using InferMed.MACRO.ClinicalCoding;
using MACROLOCKBS30;
using log4net.Config;
using log4net;
using System.IO;
using System.Data.OracleClient;
using System.Data.SqlClient;

namespace InferMed.MACRO.ClinicalCoding.MACRO_CC
{
	/// <summary>
	/// Summary description for MainForm.
	/// </summary>
	public class MainForm : System.Windows.Forms.Form
	{
		//logging interface
		private static readonly ILog log = LogManager.GetLogger( typeof( MainForm ) );

		//function code
		private const string _FN_CODECLINICALRESPONSE = "F6009";
		private const string _FN_IMPORTCLINICALDICTIONARY = "F6012";
		private const string _FN_AUTOENCODECLINICALRESPONSE = "F6013";

		//currently logged into database connection string
		private string _dbCon = "";
		private string _dbCode = "";
		//security db connection string
		private string _secCon = "";
		//user
		private string _userName = "";
		private string _userNameFull = "";
		private string _MACRORole = "";
		//permissions
		private bool _codeResponse = false;
		private bool _autoEncode = false;
		private bool _importDictionary = false;
		//preload list
		private string _plList = "";
		//collection of dictionaries
		MACROCCBS30.Dictionaries _dictionaries = null;
		//settings
		private const string _SETTINGS_FILE = "MACROSettings30.txt";
		private const string _SECURITY_PATH = "securitypath";
		private const string _PRELOADLIST = "meddrapreloadlist";


		//private InferMed.Components.IMEDLogin.IMEDLogin imedLogin1;
		private System.Windows.Forms.MainMenu mainMenu1;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem5;
		private System.Windows.Forms.MenuItem menuItem10;
		private System.Windows.Forms.MenuItem menuItem11;
		private System.Windows.Forms.StatusBar statusBar1;
		private System.Windows.Forms.ContextMenu contextMenu1;
		private System.Windows.Forms.MenuItem menuItem12;
		private System.Windows.Forms.MenuItem menuItem13;
		private System.Windows.Forms.ContextMenu contextMenu2;
		private System.Windows.Forms.MenuItem menuItem6;
		private System.Windows.Forms.MenuItem menuItem7;
		private System.Windows.Forms.MenuItem menuItem14;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.TabPage tabPage2;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.ListView lvwTerms;
		//private InferMed.Components.IMEDLogin.IMEDLogin imedLogin2;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.ListView lvwTermsU;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.ComboBox cboSites;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.ComboBox cboStudies;
		private System.Windows.Forms.Button btnRun;
		private System.Windows.Forms.Button btnCommit;
		private System.Windows.Forms.GroupBox groupBox6;
		private System.Windows.Forms.Button btnMapPath;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.TextBox txtMapPath;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnPluginPath;
		private System.Windows.Forms.TextBox txtPluginPath;
		private System.Windows.Forms.Button btnImport;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.ListView lvwDictionaries;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.MenuItem menuItem2;
		private InferMed.Components.IMEDLogin.IMEDLogin imedLogin1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		static void Main()
		{
			Application.Run( new MainForm() );
		}

		public MainForm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//set interface caption
			this.Text = "InferMed MACRO Clinical Coding Console";
			
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MainForm));
			this.mainMenu1 = new System.Windows.Forms.MainMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.menuItem5 = new System.Windows.Forms.MenuItem();
			this.menuItem10 = new System.Windows.Forms.MenuItem();
			this.menuItem11 = new System.Windows.Forms.MenuItem();
			this.statusBar1 = new System.Windows.Forms.StatusBar();
			this.contextMenu1 = new System.Windows.Forms.ContextMenu();
			this.menuItem12 = new System.Windows.Forms.MenuItem();
			this.menuItem13 = new System.Windows.Forms.MenuItem();
			this.menuItem7 = new System.Windows.Forms.MenuItem();
			this.menuItem14 = new System.Windows.Forms.MenuItem();
			this.menuItem2 = new System.Windows.Forms.MenuItem();
			this.contextMenu2 = new System.Windows.Forms.ContextMenu();
			this.menuItem6 = new System.Windows.Forms.MenuItem();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.lvwDictionaries = new System.Windows.Forms.ListView();
			this.groupBox6 = new System.Windows.Forms.GroupBox();
			this.btnImport = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.btnPluginPath = new System.Windows.Forms.Button();
			this.txtPluginPath = new System.Windows.Forms.TextBox();
			this.btnMapPath = new System.Windows.Forms.Button();
			this.label5 = new System.Windows.Forms.Label();
			this.txtMapPath = new System.Windows.Forms.TextBox();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.lvwTerms = new System.Windows.Forms.ListView();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.lvwTermsU = new System.Windows.Forms.ListView();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnCommit = new System.Windows.Forms.Button();
			this.btnRun = new System.Windows.Forms.Button();
			this.cboSites = new System.Windows.Forms.ComboBox();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cboStudies = new System.Windows.Forms.ComboBox();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.imedLogin1 = new InferMed.Components.IMEDLogin.IMEDLogin();
			this.tabControl1.SuspendLayout();
			this.tabPage1.SuspendLayout();
			this.groupBox5.SuspendLayout();
			this.groupBox6.SuspendLayout();
			this.tabPage2.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.groupBox4.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// mainMenu1
			// 
			this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																							this.menuItem1,
																																							this.menuItem10});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																							this.menuItem5});
			this.menuItem1.Text = "&File";
			// 
			// menuItem5
			// 
			this.menuItem5.Index = 0;
			this.menuItem5.Text = "Exit";
			this.menuItem5.Click += new System.EventHandler(this.menuItem5_Click);
			// 
			// menuItem10
			// 
			this.menuItem10.Index = 1;
			this.menuItem10.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																							 this.menuItem11});
			this.menuItem10.Text = "&Help";
			// 
			// menuItem11
			// 
			this.menuItem11.Index = 0;
			this.menuItem11.Text = "&About";
			this.menuItem11.Click += new System.EventHandler(this.menuItem11_Click);
			// 
			// statusBar1
			// 
			this.statusBar1.Location = new System.Drawing.Point(0, 513);
			this.statusBar1.Name = "statusBar1";
			this.statusBar1.Size = new System.Drawing.Size(852, 24);
			this.statusBar1.TabIndex = 14;
			this.statusBar1.Text = "statusBar1";
			// 
			// contextMenu1
			// 
			this.contextMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								 this.menuItem12,
																																								 this.menuItem13,
																																								 this.menuItem7,
																																								 this.menuItem14,
																																								 this.menuItem2});
			// 
			// menuItem12
			// 
			this.menuItem12.Enabled = false;
			this.menuItem12.Index = 0;
			this.menuItem12.Text = "View tree";
			this.menuItem12.Click += new System.EventHandler(this.menuItem12_Click);
			// 
			// menuItem13
			// 
			this.menuItem13.Enabled = false;
			this.menuItem13.Index = 1;
			this.menuItem13.Text = "Browser...";
			this.menuItem13.Click += new System.EventHandler(this.menuItem13_Click);
			// 
			// menuItem7
			// 
			this.menuItem7.Index = 2;
			this.menuItem7.Text = "-";
			// 
			// menuItem14
			// 
			this.menuItem14.Index = 3;
			this.menuItem14.Text = "Check All";
			this.menuItem14.Click += new System.EventHandler(this.menuItem14_Click_1);
			// 
			// menuItem2
			// 
			this.menuItem2.Index = 4;
			this.menuItem2.Text = "Uncheck All";
			this.menuItem2.Click += new System.EventHandler(this.menuItem2_Click);
			// 
			// contextMenu2
			// 
			this.contextMenu2.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								 this.menuItem6});
			// 
			// menuItem6
			// 
			this.menuItem6.Enabled = false;
			this.menuItem6.Index = 0;
			this.menuItem6.Text = "Browser...";
			this.menuItem6.Click += new System.EventHandler(this.menuItem6_Click);
			// 
			// tabControl1
			// 
			this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabControl1.Controls.Add(this.tabPage1);
			this.tabControl1.Controls.Add(this.tabPage2);
			this.tabControl1.Location = new System.Drawing.Point(4, 8);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(844, 500);
			this.tabControl1.TabIndex = 15;
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.groupBox5);
			this.tabPage1.Controls.Add(this.groupBox6);
			this.tabPage1.Location = new System.Drawing.Point(4, 22);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(836, 474);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "Dictionaries";
			// 
			// groupBox5
			// 
			this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox5.Controls.Add(this.imedLogin1);
			this.groupBox5.Controls.Add(this.lvwDictionaries);
			this.groupBox5.Location = new System.Drawing.Point(4, 4);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(832, 352);
			this.groupBox5.TabIndex = 13;
			this.groupBox5.TabStop = false;
			// 
			// lvwDictionaries
			// 
			this.lvwDictionaries.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lvwDictionaries.CheckBoxes = true;
			this.lvwDictionaries.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.lvwDictionaries.Location = new System.Drawing.Point(4, 12);
			this.lvwDictionaries.Name = "lvwDictionaries";
			this.lvwDictionaries.Size = new System.Drawing.Size(824, 332);
			this.lvwDictionaries.TabIndex = 15;
			// 
			// groupBox6
			// 
			this.groupBox6.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox6.Controls.Add(this.btnImport);
			this.groupBox6.Controls.Add(this.label1);
			this.groupBox6.Controls.Add(this.btnPluginPath);
			this.groupBox6.Controls.Add(this.txtPluginPath);
			this.groupBox6.Controls.Add(this.btnMapPath);
			this.groupBox6.Controls.Add(this.label5);
			this.groupBox6.Controls.Add(this.txtMapPath);
			this.groupBox6.Location = new System.Drawing.Point(4, 360);
			this.groupBox6.Name = "groupBox6";
			this.groupBox6.Size = new System.Drawing.Size(832, 108);
			this.groupBox6.TabIndex = 12;
			this.groupBox6.TabStop = false;
			this.groupBox6.Text = "Import Dictionary";
			// 
			// btnImport
			// 
			this.btnImport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImport.Enabled = false;
			this.btnImport.Location = new System.Drawing.Point(756, 76);
			this.btnImport.Name = "btnImport";
			this.btnImport.Size = new System.Drawing.Size(64, 24);
			this.btnImport.TabIndex = 12;
			this.btnImport.Text = "Import";
			this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 16);
			this.label1.TabIndex = 10;
			this.label1.Text = "Plugin Path";
			// 
			// btnPluginPath
			// 
			this.btnPluginPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnPluginPath.Location = new System.Drawing.Point(800, 44);
			this.btnPluginPath.Name = "btnPluginPath";
			this.btnPluginPath.Size = new System.Drawing.Size(24, 20);
			this.btnPluginPath.TabIndex = 8;
			this.btnPluginPath.Text = "...";
			this.btnPluginPath.Click += new System.EventHandler(this.btnPluginPath_Click);
			// 
			// txtPluginPath
			// 
			this.txtPluginPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtPluginPath.Location = new System.Drawing.Point(84, 44);
			this.txtPluginPath.Name = "txtPluginPath";
			this.txtPluginPath.Size = new System.Drawing.Size(712, 20);
			this.txtPluginPath.TabIndex = 7;
			this.txtPluginPath.TabStop = false;
			this.txtPluginPath.Text = "";
			this.txtPluginPath.TextChanged += new System.EventHandler(this.Import_TextChanged);
			// 
			// btnMapPath
			// 
			this.btnMapPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnMapPath.Location = new System.Drawing.Point(800, 20);
			this.btnMapPath.Name = "btnMapPath";
			this.btnMapPath.Size = new System.Drawing.Size(24, 20);
			this.btnMapPath.TabIndex = 6;
			this.btnMapPath.Text = "...";
			this.btnMapPath.Click += new System.EventHandler(this.btnMapPath_Click);
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(8, 24);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(64, 16);
			this.label5.TabIndex = 5;
			this.label5.Text = "Map path";
			// 
			// txtMapPath
			// 
			this.txtMapPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtMapPath.Location = new System.Drawing.Point(84, 20);
			this.txtMapPath.Name = "txtMapPath";
			this.txtMapPath.Size = new System.Drawing.Size(712, 20);
			this.txtMapPath.TabIndex = 4;
			this.txtMapPath.TabStop = false;
			this.txtMapPath.Text = "";
			this.txtMapPath.TextChanged += new System.EventHandler(this.Import_TextChanged);
			// 
			// tabPage2
			// 
			this.tabPage2.Controls.Add(this.groupBox2);
			this.tabPage2.Controls.Add(this.groupBox1);
			this.tabPage2.Location = new System.Drawing.Point(4, 22);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Size = new System.Drawing.Size(836, 474);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "Autoencoder";
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.groupBox4);
			this.groupBox2.Controls.Add(this.groupBox3);
			this.groupBox2.Location = new System.Drawing.Point(0, 112);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(836, 360);
			this.groupBox2.TabIndex = 19;
			this.groupBox2.TabStop = false;
			// 
			// groupBox4
			// 
			this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox4.Controls.Add(this.lvwTerms);
			this.groupBox4.Location = new System.Drawing.Point(4, 16);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(828, 220);
			this.groupBox4.TabIndex = 13;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "Matched Terms (0)";
			// 
			// lvwTerms
			// 
			this.lvwTerms.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lvwTerms.CheckBoxes = true;
			this.lvwTerms.ContextMenu = this.contextMenu1;
			this.lvwTerms.FullRowSelect = true;
			this.lvwTerms.GridLines = true;
			this.lvwTerms.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.lvwTerms.Location = new System.Drawing.Point(4, 16);
			this.lvwTerms.Name = "lvwTerms";
			this.lvwTerms.Size = new System.Drawing.Size(820, 204);
			this.lvwTerms.TabIndex = 10;
			this.lvwTerms.View = System.Windows.Forms.View.List;
			this.lvwTerms.SelectedIndexChanged += new System.EventHandler(this.lvwTerms_SelectedIndexChanged);
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Controls.Add(this.lvwTermsU);
			this.groupBox3.Location = new System.Drawing.Point(4, 240);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(828, 116);
			this.groupBox3.TabIndex = 11;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Unmatched Terms (0)";
			// 
			// lvwTermsU
			// 
			this.lvwTermsU.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lvwTermsU.ContextMenu = this.contextMenu2;
			this.lvwTermsU.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.lvwTermsU.Location = new System.Drawing.Point(4, 16);
			this.lvwTermsU.MultiSelect = false;
			this.lvwTermsU.Name = "lvwTermsU";
			this.lvwTermsU.Size = new System.Drawing.Size(820, 96);
			this.lvwTermsU.TabIndex = 11;
			this.lvwTermsU.SelectedIndexChanged += new System.EventHandler(this.lvwTermsU_SelectedIndexChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.btnCommit);
			this.groupBox1.Controls.Add(this.btnRun);
			this.groupBox1.Controls.Add(this.cboSites);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.cboStudies);
			this.groupBox1.Location = new System.Drawing.Point(0, 6);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(836, 104);
			this.groupBox1.TabIndex = 18;
			this.groupBox1.TabStop = false;
			// 
			// btnCommit
			// 
			this.btnCommit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCommit.Location = new System.Drawing.Point(760, 72);
			this.btnCommit.Name = "btnCommit";
			this.btnCommit.Size = new System.Drawing.Size(64, 24);
			this.btnCommit.TabIndex = 10;
			this.btnCommit.Text = "Commit";
			this.btnCommit.Click += new System.EventHandler(this.btnCommit_Click);
			// 
			// btnRun
			// 
			this.btnRun.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnRun.Location = new System.Drawing.Point(688, 72);
			this.btnRun.Name = "btnRun";
			this.btnRun.Size = new System.Drawing.Size(64, 24);
			this.btnRun.TabIndex = 9;
			this.btnRun.Text = "Run";
			this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
			// 
			// cboSites
			// 
			this.cboSites.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.cboSites.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboSites.DropDownWidth = 720;
			this.cboSites.Location = new System.Drawing.Point(60, 40);
			this.cboSites.Name = "cboSites";
			this.cboSites.Size = new System.Drawing.Size(768, 20);
			this.cboSites.TabIndex = 8;
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(4, 40);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(56, 16);
			this.label3.TabIndex = 7;
			this.label3.Text = "Site";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(4, 12);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(56, 16);
			this.label2.TabIndex = 6;
			this.label2.Text = "Study";
			// 
			// cboStudies
			// 
			this.cboStudies.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.cboStudies.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboStudies.DropDownWidth = 720;
			this.cboStudies.Enabled = false;
			this.cboStudies.Location = new System.Drawing.Point(60, 12);
			this.cboStudies.Name = "cboStudies";
			this.cboStudies.Size = new System.Drawing.Size(768, 20);
			this.cboStudies.TabIndex = 1;
			this.cboStudies.SelectedIndexChanged += new System.EventHandler(this.cboStudies_SelectedIndexChanged);
			// 
			// imedLogin1
			// 
			this.imedLogin1.ApplicationPermissionCheck = "";
			this.imedLogin1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("imedLogin1.BackgroundImage")));
			this.imedLogin1.DataDBSetting = "";
			this.imedLogin1.InitialSettingsFile = "";
			this.imedLogin1.Location = new System.Drawing.Point(16, 28);
			this.imedLogin1.MandatoryMACROLogin = true;
			this.imedLogin1.MandatorySecurityDB = true;
			this.imedLogin1.Name = "imedLogin1";
			this.imedLogin1.SecurityDbLoginOnly = false;
			this.imedLogin1.Size = new System.Drawing.Size(36, 36);
			this.imedLogin1.TabIndex = 16;
			this.imedLogin1.Visible = false;
			// 
			// MainForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(852, 537);
			this.Controls.Add(this.tabControl1);
			this.Controls.Add(this.statusBar1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Menu = this.mainMenu1;
			this.Name = "MainForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "InferMed MACRO Clinical Coding Console";
			this.Load += new System.EventHandler(this.MainForm_Load);
			this.tabControl1.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			this.groupBox5.ResumeLayout(false);
			this.groupBox6.ResumeLayout(false);
			this.tabPage2.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.groupBox4.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Initialise listview controls
		/// </summary>
		private void InitListView()
		{
			//matched terms listview
			lvwTerms.View = View.Details;
			lvwTerms.CheckBoxes = true;
			lvwTerms.GridLines = true;
			lvwTerms.MultiSelect = true;
			lvwTerms.FullRowSelect = true;
			lvwTerms.Columns.Add("Response Value", 100, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("Possible matches", 50, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("Dictionary", 60, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("Low Level Term", 100, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("LLT Key", 60, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("Preferred Term", 100, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("High Level Term", 100, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("High Level Group Term", 100, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("System Organ Class", 100, HorizontalAlignment.Left);

			//unmatched terms listview
			lvwTermsU.View = View.Details;
			lvwTermsU.GridLines = true;
			lvwTermsU.MultiSelect = false;
			lvwTermsU.FullRowSelect = true;
			lvwTermsU.Columns.Add("Response Value", 100, HorizontalAlignment.Left);
			lvwTermsU.Columns.Add("Possible matches", 50, HorizontalAlignment.Left);
			lvwTermsU.Columns.Add("Dictionary", 60, HorizontalAlignment.Left);

			//dictionaries listview
			lvwDictionaries.View = View.Details;
			lvwDictionaries.GridLines = true;
			lvwDictionaries.MultiSelect = false;
			lvwDictionaries.FullRowSelect = true;
			lvwDictionaries.Columns.Add("Preload", 25, HorizontalAlignment.Left);
			lvwDictionaries.Columns.Add("Id", 35, HorizontalAlignment.Left);
			lvwDictionaries.Columns.Add("Name", 90, HorizontalAlignment.Left);
			lvwDictionaries.Columns.Add("Version", 70, HorizontalAlignment.Left);
			lvwDictionaries.Columns.Add("Connection", 600, HorizontalAlignment.Left);
		}

		/// <summary>
		/// Load dictionaries from database
		/// </summary>
		private void LoadDictionaries()
		{
			_plList =  MedDRA.GetSetting( _PRELOADLIST, "" );

			//set and load dictionary object
			_dictionaries = new MACROCCBS30.Dictionaries();
			_dictionaries.Init( _secCon );

			//loop through dictionaries, adding each to the dictionary listview
			foreach( MACROCCBS30.Dictionary d in _dictionaries.DictionaryList )
			{
				ListViewItem lvi = new ListViewItem();
				lvi.SubItems.Add( d.Id.ToString() );
				lvi.SubItems.Add( d.Name );
				lvi.SubItems.Add( d.Version );
				lvi.SubItems.Add( d.Connection );
				if( IsPreloaded( d.Id.ToString() ) ) lvi.Checked = true;
				lvwDictionaries.Items.Add( lvi );
			}
		}

		/// <summary>
		/// Show progress
		/// </summary>
		/// <param name="prog"></param>
		private void ShowProgress( string prog, bool msg )
		{
			statusBar1.Text = prog;
			Application.DoEvents();
			if( msg ) MessageBox.Show( prog );
		}

		/// <summary>
		/// Add a passed term to the listview
		/// </summary>
		/// <param name="t"></param>
		private void AddTermToList( AutoCodedTermHistory t )
		{
			//create listview item
			ListViewItem lvi = new ListViewItem();
			//tag on the coded term object
			lvi.Tag = ( object )t;
			//set the common columns
			lvi.Text = t.ResponseValue;
			lvi.SubItems.Add( t.Matches.ToString() );
			lvi.SubItems.Add( t.CCDictionary.Name + " " + t.CCDictionary.Version );

			//if it is edited, it has been autocoded
			if( t.IsEdited )
			{
				//set the autocoded columns
				string dic, soc, socKey, hlgt, hlgtKey, hlt, hltKey, pt, ptKey, llt, lltKey;
				Plugins.CCXml.MedDRAUnwrapXmlNames( t.CodingDetails, out dic, out soc, out hlgt, out hlt, out pt, out llt );
				Plugins.CCXml.MedDRAUnwrapXmlKeys( t.CodingDetails, out dic, out socKey, out hlgtKey, out hltKey, out ptKey, out lltKey );

				lvi.SubItems.Add( llt );
				lvi.SubItems.Add( lltKey );
				lvi.SubItems.Add( pt );
				lvi.SubItems.Add( hlt );
				lvi.SubItems.Add( hlgt );
				lvi.SubItems.Add( soc );
				lvwTerms.Items.Add( lvi );
			}
			else
			{
				lvwTermsU.Items.Add( lvi );
			}
			//update the counters
			SetCount();
			Application.DoEvents();
		}

		/// <summary>
		/// Load studies combo
		/// </summary>
		private void LoadStudies()
		{
			DataSet ds = new DataSet();

			try
			{
				this.Cursor = Cursors.WaitCursor;
				cboStudies.Enabled = false;

				ds = DataAccess.GetStudyList( _dbCon );
	
				//clear, then fill the combobox with studies
				cboStudies.Items.Clear();
				foreach( DataRow row in ds.Tables[0].Rows )
				{
					cboStudies.Items.Add( new ComboItem( row[0].ToString(), row[1].ToString() ) );
				}
				if( cboStudies.Items.Count > 0 )
				{
					cboStudies.Enabled = true;
					cboStudies.SelectedIndex = 0;
				}
				else
				{
					cboSites.Items.Clear();
					cboSites.Enabled = false;
				}
				this.Cursor = Cursors.Default;
			}
			finally
			{
				ds.Dispose();
				
			}
		}

		/// <summary>
		/// Load sites combo
		/// </summary>
		private void LoadSites()
		{
			DataSet ds = new DataSet();

			try
			{
				this.Cursor = Cursors.WaitCursor;
				cboSites.Enabled = false;
	
				ComboItem ciSt = ( ComboItem )cboStudies.SelectedItem;
			
				ds = DataAccess.GetSiteList( _dbCon, ciSt.Value );

				//clear, then fill the combobox with sites
				cboSites.Items.Clear();
				foreach( DataRow row in ds.Tables[0].Rows )
				{
					cboSites.Items.Add( new ComboItem( row[0].ToString(), row[0].ToString() ) );
				}
				if( cboSites.Items.Count > 0 )
				{
					cboSites.Items.Insert( 0, new ComboItem( "All sites", "" ) );
					cboSites.Enabled = true;
					cboSites.Enabled = true;
					cboSites.SelectedIndex = 0;
					btnRun.Enabled = true;
				}
				else
				{
					btnRun.Enabled = false;
				}
				this.Cursor = Cursors.Default;
			}
			finally
			{
				ds.Dispose();
			}
		}

		/// <summary>
		/// Studies combo has changed
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void cboStudies_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			LoadSites();
		}
		
		/// <summary>
		/// Check for previous app instance
		/// </summary>
		/// <returns></returns>
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
		/// Set counters
		/// </summary>
		private void SetCount()
		{
			groupBox3.Text = "Unmatched Terms (" + lvwTermsU.Items.Count.ToString() + ")";
			groupBox4.Text = "Matched Terms (" + lvwTerms.Items.Count.ToString() + ")";
		}

		/// <summary>
		/// Enable or disable form fields
		/// </summary>
		/// <param name="enable"></param>
		private void EnableForm( bool enable )
		{
			tabControl1.Enabled = enable;
			menuItem5.Enabled = enable;
			menuItem11.Enabled = enable;
		}

		/// <summary>
		/// Lock macro subject
		/// </summary>
		/// <param name="clinicalTrialId"></param>
		/// <param name="trialSite"></param>
		/// <param name="personId"></param>
		/// <returns></returns>
		private string LockSubject( long clinicalTrialId, string trialSite, long personId )
		{
			DBLockClass l = new DBLockClass();
			float wait = 0;
			string token;

			try
			{
				token = l.LockSubject( _dbCon, _userName, System.Convert.ToInt32( clinicalTrialId ), 
					trialSite, System.Convert.ToInt32( personId ), wait );
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
			}
			catch
			{
				return( "" );
			}
		}

		/// <summary>
		/// Unlock macro subject
		/// </summary>
		/// <param name="token"></param>
		/// <param name="clinicalTrialId"></param>
		/// <param name="trialSite"></param>
		/// <param name="personId"></param>
		private void UnLockSubject( string token, long clinicalTrialId, string trialSite, long personId )
		{
			DBLockClass l = new DBLockClass();

			try
			{
				l.UnlockSubject( _dbCon, token, System.Convert.ToInt32( clinicalTrialId ), trialSite, 
					System.Convert.ToInt32( personId ) );
			}
			catch
			{
			}
		}

		/// <summary>
		/// Display coding tree
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void menuItem12_Click(object sender, System.EventArgs e)
		{
			if( lvwTerms.SelectedItems.Count > 0 )
			{
				//get the attached object
				ListViewItem lvi = lvwTerms.SelectedItems[0];
				AutoCodedTermHistory t = ( AutoCodedTermHistory )lvi.Tag;
				//display the tree dialog
				Plugins.MedDRATree mt = new Plugins.MedDRATree( t.CodingDetails );
				mt.ShowDialog();
				mt.Dispose();
			}
		}

		/// <summary>
		/// Re-code an autoencoded term
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void menuItem13_Click(object sender, System.EventArgs e)
		{
			if( lvwTerms.SelectedItems.Count > 0 )
			{
				string rfc = "";
				bool rfcOK = true;
				//get attached object
				ListViewItem lvi = lvwTerms.SelectedItems[0];
				AutoCodedTermHistory t = ( AutoCodedTermHistory )lvi.Tag;
				string codedValue = "", responseValue = t.ResponseValue;
				//create browser object
				Plugins.MedDRA m = new Plugins.MedDRA();
				//call code method
				m.Code( t.CCDictionary.Name, t.CCDictionary.Version, t.CCDictionary.Custom, ref responseValue, ref codedValue );

				//if coded value is returned
				if( ( codedValue != "" ) && ( codedValue != t.CodingDetails ) )
				{
					if( t.RequiresRFC )
					{
						RFCForm f = new RFCForm();
						f.ShowDialog();
						rfc = f.RFC;
						f.Dispose();
						if( rfc == "" ) rfcOK = false;
					}

					if( rfcOK )
					{
						//set the code
						t.SetCode( t.CCDictionary.Name, t.CCDictionary.Version, codedValue, _userName, _userNameFull,
							t.ResponseValue, t.ResponseTimeStamp, t.ResponseTimeStamp_TZ, "", false );

						string dic, soc, socKey, hlgt, hlgtKey, hlt, hltKey, pt, ptKey, llt, lltKey;
						Plugins.CCXml.MedDRAUnwrapXmlNames( t.CodingDetails, out dic, out soc, out hlgt, out hlt, out pt, out llt );
						Plugins.CCXml.MedDRAUnwrapXmlKeys( t.CodingDetails, out dic, out socKey, out hlgtKey, out hltKey, out ptKey, out lltKey );

						//update the listview
						lvi.SubItems[3].Text = llt;
						lvi.SubItems[4].Text = lltKey;
						lvi.SubItems[5].Text = pt;
						lvi.SubItems[6].Text =  hlt;
						lvi.SubItems[7].Text = hlgt;
						lvi.SubItems[8].Text = soc;
					}
				}
			}
		}

		/// <summary>
		/// Enable or disable context menu items
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void lvwTerms_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if( lvwTerms.SelectedItems.Count == 1 )
			{
				contextMenu1.MenuItems[0].Enabled = true;
				contextMenu1.MenuItems[1].Enabled = ( _codeResponse ) ? true: false;
			}
			else
			{
				contextMenu1.MenuItems[0].Enabled = false;
				contextMenu1.MenuItems[1].Enabled = false;
			}
		}

		/// <summary>
		/// Exit
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void menuItem5_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		/// <summary>
		/// About dialog
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void menuItem11_Click(object sender, System.EventArgs e)
		{
			MessageBox.Show("(c) " + Application.CompanyName + "\n" + "MACRO Clinical Coding Console " 
				+ Application.ProductVersion, "MACRO Clinical Coding Console", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		/// <summary>
		/// Code an unmatched term
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void menuItem6_Click(object sender, System.EventArgs e)
		{
			if( lvwTermsU.SelectedItems.Count > 0 )
			{
				string rfc = "";
				bool rfcOK = true;
				//get attached object
				ListViewItem lvi = lvwTermsU.SelectedItems[0];
				AutoCodedTermHistory t = ( AutoCodedTermHistory )lvi.Tag;
				string codedValue = "", responseValue = t.ResponseValue;
				//create browser object
				Plugins.MedDRA m = new Plugins.MedDRA();
				//call code method
				m.Code( t.CCDictionary.Name, t.CCDictionary.Version, t.CCDictionary.Custom, ref responseValue, ref codedValue );

				//if a coded value is returned
				if( codedValue != "" )
				{
					if( t.RequiresRFC )
					{
						RFCForm f = new RFCForm();
						f.ShowDialog();
						rfc = f.RFC;
						f.Dispose();
						if( rfc == "" ) rfcOK = false;
					}

					if( rfcOK )
					{
						//set the code
						t.SetCode( t.CCDictionary.Name, t.CCDictionary.Version, codedValue, _userName, _userNameFull,
							t.ResponseValue, t.ResponseTimeStamp, t.ResponseTimeStamp_TZ, rfc, false );

						//add the coded term to the matched term list and remove it from the unmatched term list
						AddTermToList( t );
						lvwTermsU.Items.Remove( lvi );
						SetCount();
					}
				}
			}
		}

		/// <summary>
		/// Check all
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void menuItem14_Click_1(object sender, System.EventArgs e)
		{
			foreach( ListViewItem lvi in lvwTerms.Items )
			{
				lvi.Checked = true;
			}
		}

		/// <summary>
		/// Uncheck all
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void menuItem2_Click(object sender, System.EventArgs e)
		{
			foreach( ListViewItem lvi in lvwTerms.Items )
			{
				lvi.Checked = false;
			}
		}

		/// <summary>
		/// Enable or disable context menu items
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void lvwTermsU_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if( ( lvwTermsU.SelectedItems.Count > 0 ) && ( _codeResponse ) )
			{
				contextMenu2.MenuItems[0].Enabled = true;
			}
			else
			{
				contextMenu2.MenuItems[0].Enabled = false;
			}
		}

		/// <summary>
		/// Autoencode
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnRun_Click(object sender, System.EventArgs e)
		{
			try
			{
				ArrayList alSite = new ArrayList();

				EnableForm( false );
				this.Cursor = Cursors.WaitCursor;

				contextMenu1.MenuItems[0].Enabled = false;
				contextMenu1.MenuItems[1].Enabled = false;
				contextMenu2.MenuItems[0].Enabled = false;
	
				//clear matched and unmatched lists
				lvwTerms.Items.Clear();
				lvwTermsU.Items.Clear();
				//get database, study and site 
				ComboItem ciSt = ( ComboItem )cboStudies.SelectedItem;
				ComboItem ciSi = ( ComboItem )cboSites.SelectedItem;
				if( ciSi.Code == "" )
				{
					//all sites
					for( int n = 1; n < cboSites.Items.Count; n++ )
					{
						ciSi = ( ComboItem )cboSites.Items[n];
						alSite.Add( ciSi.Code );
					}
				}
				else
				{
					//single site
					alSite.Add( ciSi.Code );
				}
		
				//run autoencoder
				MedDRA.AutoEncode( _dbCon, _userName, _userNameFull, 
					System.Convert.ToInt32( ciSt.Code ), alSite );

			}
			catch( Exception ex )
			{
				MessageBox.Show( "MACRO encountered a problem while autoencoding : " + ex.Message + " : " + ex.InnerException,
					"Autoencoding", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				this.Cursor = Cursors.Default;
				EnableForm( true );
				ShowProgress( "Ready", false );
			}
		}

		private void btnCommit_Click(object sender, System.EventArgs e)
		{
			try
			{
				if( MessageBox.Show( "Are you sure you want to commit all checked codes to the database?", "Commit Codes", 
					MessageBoxButtons.YesNo, MessageBoxIcon.Question ) == DialogResult.Yes )
				{
					bool lockFail = false;
					EnableForm( false );
					this.Cursor = Cursors.WaitCursor;

					lvwTerms.SelectedItems.Clear();
					contextMenu1.MenuItems[0].Enabled = false;
					contextMenu1.MenuItems[1].Enabled = false;

					//loop through all matched items
					foreach( ListViewItem lvi in lvwTerms.Items )
					{
						//if they are checked
						if( lvi.Checked )
						{
							//get the attached object
							AutoCodedTermHistory t = ( AutoCodedTermHistory )lvi.Tag;
							//get subject lock
							ShowProgress( "Locking subject " + t.PersonId, false );
							string token = LockSubject( t.ClinicalTrialId, t.TrialSite, t.PersonId );
							//save it
							if( token != "" )
							{
								ShowProgress( "Committing code " + t.PersonId + ", " + t.ResponseTaskId + " [" + t.Repeat + "]", false );
								t.Save( _dbCon, t.VisitId, t.VisitCycle, t.CrfPageId, t.CrfPageCycle );
								//remove it from the list
								lvwTerms.Items.Remove( lvi );
								//release lock
								ShowProgress( "Unlocking subject " + t.PersonId, false );
								UnLockSubject( token, t.ClinicalTrialId, t.TrialSite, t.PersonId );
								SetCount();
							}
							else
							{
								lockFail = true;
							}
						}
					}

					if( lockFail ) MessageBox.Show( "Some terms could not be committed as another user is editing the subject", 
						"Commit Codes", MessageBoxButtons.OK, MessageBoxIcon.Warning );
				}
			}
			catch( Exception ex )
			{
				MessageBox.Show( "MACRO encountered a problem while committing : " + ex.Message + " : " + ex.InnerException,
					"Commit Codes", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				this.Cursor = Cursors.Default;
				EnableForm( true );
				ShowProgress( "Ready", false );
			}
		}

		/// <summary>
		/// Browse for plugin
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnPluginPath_Click(object sender, System.EventArgs e)
		{
			openFileDialog1.Filter = "Dll files|*.dll";
			if( openFileDialog1.ShowDialog() == DialogResult.OK )
			{
				txtPluginPath.Text = openFileDialog1.FileName;
			}
			openFileDialog1.Dispose();
		}

		/// <summary>
		/// Get filename from path
		/// </summary>
		/// <param name="fullPath"></param>
		/// <returns></returns>
		private string GetFileName( string fullPath )
		{
			return( Path.GetFileName( fullPath ) );
		}

		/// <summary>
		/// Get path without filename from path
		/// </summary>
		/// <param name="fullPath"></param>
		/// <returns></returns>
		private string GetFilePath( string fullPath )
		{
			return( fullPath.Substring( 0,  ( fullPath.LastIndexOf( @"\" ) ) ) + @"\" );
		}

		/// <summary>
		/// Browse for cc map file
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnMapPath_Click(object sender, System.EventArgs e)
		{
			openFileDialog1.Filter = "XML files|*.xml";
			if( openFileDialog1.ShowDialog() == DialogResult.OK )
			{
				txtMapPath.Text = openFileDialog1.FileName;
			}
			openFileDialog1.Dispose();
		}

		/// <summary>
		/// Import a dictionary
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnImport_Click(object sender, System.EventArgs e)
		{
			string dName, dVersion, stlDName, prefix, pluginNamespace;

			try
			{
				//confirm action
				if( MessageBox.Show( "Are you sure you want to import this dictionary?", "Import Dictionary", 
					MessageBoxButtons.YesNo, MessageBoxIcon.Question ) == DialogResult.Yes )
				{
					DateTime s = DateTime.Now;

					//disable the form
					EnableForm( false );

					//get the clinical coding database connection string from the settings file,
					//or from the user
					string ccCon = MedDRA.GetDBConnectionString();
				
					//check for necessary upgrade
					DBUpgrade dbUpgrade = new DBUpgrade( ccCon );
					if( dbUpgrade.UpgradeRequired )
					{
						if( MessageBox.Show( "Do you want to create the Medicoder database structure?", "Database Upgrade Required",
							MessageBoxButtons.YesNo, MessageBoxIcon.Question ) == DialogResult.Yes )
						{
							ShowProgress( "Running creation scripts...", false );
							dbUpgrade.Upgrade();
						}
						else
						{
							Exception ex = new Exception( "Dictionaries cannot be imported until the Medicoder database is created" );
							throw ex;
						}
					}

					if( ccCon != "" )
					{
						//import the dictionary
						MedDRA.Import( ccCon, txtMapPath.Text, GetFilePath( txtMapPath.Text ), out dName, out dVersion, out stlDName, 
							out prefix, out pluginNamespace );
						
						//get the custom xml string for the dictionary
						string custom = Plugins.CCXml.MedDRAWrapXmlCustom( stlDName, "", ccCon, ccCon, prefix, 
							GetFilePath( txtPluginPath.Text ), "MedDRAPluginPreferences.txt" );
						//get the connection xml string for the dictionary
						string xmlCon = Plugins.CCXml.WrapXmlConnection( dName, dVersion, pluginNamespace, 
							GetFilePath( txtPluginPath.Text ), GetFileName( txtPluginPath.Text ), custom );

						log.Info( "Registering dictionary" );
						//add the dictionary to the macro database
						ShowProgress( "Registering dictionary", false );
						_dictionaries.AddNew( _secCon, dName, dVersion, xmlCon );

						MACROCCBS30.Dictionary d = new MACROCCBS30.Dictionary();
						d = _dictionaries.DictionaryFromVersion( dName, dVersion );

						//and to the listview
						ListViewItem lvi = new ListViewItem();
						lvi.SubItems.Add( d.Id.ToString() );
						lvi.SubItems.Add( d.Name);
						lvi.SubItems.Add( d.Version);
						lvi.SubItems.Add( d.Connection );
						lvwDictionaries.Items.Add( lvi );

						DateTime el = DateTime.Now;
						TimeSpan t = el.Subtract( s );
						MessageBox.Show( "Import completed successfully in " + t.Hours.ToString() + "h:" + t.Minutes.ToString() 
							+ "m:" + t.Seconds.ToString() + "s", "Import Dictionary", MessageBoxButtons.OK, MessageBoxIcon.Information );
					}
					else
					{
						MessageBox.Show( "The dictionary cannot be imported without a valid database", "Import Dictionary", 
							MessageBoxButtons.OK, MessageBoxIcon.Exclamation );
					}
				}
			}
			catch( Exception ex )
			{
				MessageBox.Show( "MACRO encountered a problem while importing : " + ex.Message + " : " + ex.InnerException, "Import Dictionary", 
					MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
			finally
			{
				EnableForm( true );
				txtMapPath.Text = "";
				txtPluginPath.Text = "";
				ShowProgress( "Ready", false );
			}
		}

		/// <summary>
		/// Enable/disable the import button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void Import_TextChanged(object sender, System.EventArgs e)
		{
			if( ( txtMapPath.Text != "" ) && ( txtPluginPath.Text != "" ) )
			{
				btnImport.Enabled = true;
			}
			else
			{
				btnImport.Enabled = false;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="dId"></param>
		/// <returns></returns>
		private bool IsPreloaded( string dId )
		{
			//set delimiter
			char[] del1 = "|".ToCharArray();

			//split on delimiter into dictionaries: from 1|2|3 to 1 2 3
			string[] plDictionaries = _plList.Split( del1 );
			
			//loop through the dictionary ids
			for( int dic = 0; dic < plDictionaries.Length; dic++ )
			{
				if( plDictionaries[dic].ToString() == dId.ToString() ) return( true );
			}

			return( false );
		}

		private void SetPreloadList( string dId, bool on )
		{
			_plList = GetPreloadList( dId, on );
			MedDRA.SetSetting( _PRELOADLIST, _plList );
		}

		private string GetPreloadList( string dId, bool on )
		{
			string plList = "";
			char[] del1 = "|".ToCharArray();

			//if there are any dictionaries in the current preloadlist
			if( _plList != "" )
			{
				//split on delimiter into dictionaries: from 1|2|3 to 1 2 3
				string[] plDictionaries = _plList.Split( del1 );

				for( int dic = 0; dic < plDictionaries.Length; dic++ )
				{
					//create a list of all but the passed dictionary id
					if( plDictionaries[dic].ToString() != dId )
					{
						plList += plDictionaries[dic].ToString() + "|";
					}
				}
			}

			//now decide if the passed dictionary id should be added to the end
			if( on )
			{
				plList += dId;
			}

			//remove trailing delimiter, if any
			if( plList.EndsWith( "|" ) )
			{
				plList = plList.Substring(0, (plList.Length - 1 ) );
			}

			return( plList );
		}

		private void lvwDictionaries_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
			SetPreloadList( lvwDictionaries.Items[e.Index].SubItems[1].Text, ( e.NewValue == CheckState.Checked ) );
		}

		/// <summary>
		/// Load main form, perform login
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void MainForm_Load(object sender, System.EventArgs e)
		{
			IDbConnection secDbConn = null;

			try
			{
				if( PrevInstance() )
				{
					MessageBox.Show( "A previous instance of this application was detected.", 
						"MACRO Clinical Coding Console", MessageBoxButtons.OK, MessageBoxIcon.Stop );
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
								secDbConn = new OracleConnection( IMEDDataAccess.FormatConnectionString(IMEDDataAccess.ConnectionType.Oracle, _secCon) );
								break;
							case IMEDDataAccess.ConnectionType.SQLServer:
								secDbConn = new SqlConnection( IMEDDataAccess.FormatConnectionString(IMEDDataAccess.ConnectionType.SQLServer, _secCon) );
								break;
						}
						secDbConn.Open();

						//set permissions
						_codeResponse = DataAccess.CheckPermission(secDbConn, _MACRORole, _FN_CODECLINICALRESPONSE);
						_importDictionary = DataAccess.CheckPermission(secDbConn, _MACRORole, _FN_IMPORTCLINICALDICTIONARY);
						_autoEncode = DataAccess.CheckPermission(secDbConn, _MACRORole, _FN_AUTOENCODECLINICALRESPONSE);

						ShowProgress( "Initialising...", false );

						//set callback events
						MedDRA.DShowProgress DProg = new MedDRA.DShowProgress( ShowProgress );
						MedDRA.DShowProgressEvent += DProg;
						MedDRA.DAddTermToList DTerm = new MedDRA.DAddTermToList( AddTermToList );
						MedDRA.DAddTermToListEvent += DTerm;

						DBUpgrade dbUpgrade = new DBUpgrade( _secCon, _dbCon );
						if( dbUpgrade.UpgradeRequired )
						{
							if( MessageBox.Show( "Do you want to upgrade your database to allow clinical coding?", "Database Upgrade Required",
								MessageBoxButtons.YesNo, MessageBoxIcon.Question ) == DialogResult.Yes )
							{
								ShowProgress( "Running upgrade scripts...", false );
								dbUpgrade.Upgrade();
							}
							else
							{
								Exception ex = new Exception( "This application cannot be run against this database version" );
								throw ex;
							}
						}

						ShowProgress( "Initialising...", false );

						//initialise listview
						InitListView();

						if( _autoEncode )
						{
							//load and enable autoencode console if permission
							LoadStudies();
							groupBox1.Enabled = true;
						}
						else
						{
							groupBox1.Enabled = false;
						}
						LoadDictionaries();
						if( _importDictionary )
						{
							//enable import if permission
							groupBox6.Enabled = true;
						}
						else
						{
							groupBox6.Enabled = false;
						}

						//attach event after initialising to avoid initial triggers
						this.lvwDictionaries.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lvwDictionaries_ItemCheck);
				
						ShowProgress( "Ready", false );
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
					"MACRO Clinical Coding Console", MessageBoxButtons.OK, MessageBoxIcon.Error );
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
