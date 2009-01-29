using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;


//----------------------------------------------------------------------
// 28/04/2006	bug 2725 incorrect matching of eform elements
//
//----------------------------------------------------------------------

namespace InferMed.MACRO.StudyCopy
{
	/// <summary>
	/// Settings for copy study tool
	/// </summary>
	public class SettingsForm : System.Windows.Forms.Form
	{
		private string _settingsFile = "";

		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabMatch;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.ComboBox cboE1;
		private System.Windows.Forms.ComboBox cboE2;
		private System.Windows.Forms.ComboBox cboE3;
		private System.Windows.Forms.ComboBox cboEl3;
		private System.Windows.Forms.ComboBox cboEl2;
		private System.Windows.Forms.ComboBox cboEl1;
		private System.Windows.Forms.Button btnClose;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public SettingsForm( string settingsFile )
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//settings file
			_settingsFile = settingsFile;

			//initialise combo options
			InitCombos();

			//set defaults
			cboE1.SelectedIndex = 0;
			cboE2.SelectedIndex = 0;
			cboE3.SelectedIndex = 0;
			cboEl1.SelectedIndex = 0;
			cboEl2.SelectedIndex = 0;
			cboEl3.SelectedIndex = 0;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(SettingsForm));
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabMatch = new System.Windows.Forms.TabPage();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.cboEl3 = new System.Windows.Forms.ComboBox();
			this.cboEl2 = new System.Windows.Forms.ComboBox();
			this.cboEl1 = new System.Windows.Forms.ComboBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cboE3 = new System.Windows.Forms.ComboBox();
			this.cboE2 = new System.Windows.Forms.ComboBox();
			this.cboE1 = new System.Windows.Forms.ComboBox();
			this.btnClose = new System.Windows.Forms.Button();
			this.tabControl1.SuspendLayout();
			this.tabMatch.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// tabControl1
			// 
			this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabControl1.Controls.Add(this.tabMatch);
			this.tabControl1.Location = new System.Drawing.Point(4, 4);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(440, 448);
			this.tabControl1.TabIndex = 0;
			// 
			// tabMatch
			// 
			this.tabMatch.Controls.Add(this.groupBox2);
			this.tabMatch.Controls.Add(this.groupBox1);
			this.tabMatch.Location = new System.Drawing.Point(4, 22);
			this.tabMatch.Name = "tabMatch";
			this.tabMatch.Size = new System.Drawing.Size(432, 422);
			this.tabMatch.TabIndex = 0;
			this.tabMatch.Text = "Match Criteria";
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.cboEl3);
			this.groupBox2.Controls.Add(this.cboEl2);
			this.groupBox2.Controls.Add(this.cboEl1);
			this.groupBox2.Location = new System.Drawing.Point(4, 72);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(424, 56);
			this.groupBox2.TabIndex = 1;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "eForm Elements";
			// 
			// cboEl3
			// 
			this.cboEl3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboEl3.Location = new System.Drawing.Point(244, 24);
			this.cboEl3.Name = "cboEl3";
			this.cboEl3.Size = new System.Drawing.Size(172, 21);
			this.cboEl3.TabIndex = 10;
			// 
			// cboEl2
			// 
			this.cboEl2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboEl2.Location = new System.Drawing.Point(184, 24);
			this.cboEl2.Name = "cboEl2";
			this.cboEl2.Size = new System.Drawing.Size(56, 21);
			this.cboEl2.TabIndex = 9;
			// 
			// cboEl1
			// 
			this.cboEl1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboEl1.Location = new System.Drawing.Point(8, 24);
			this.cboEl1.Name = "cboEl1";
			this.cboEl1.Size = new System.Drawing.Size(172, 21);
			this.cboEl1.TabIndex = 8;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.cboE3);
			this.groupBox1.Controls.Add(this.cboE2);
			this.groupBox1.Controls.Add(this.cboE1);
			this.groupBox1.Location = new System.Drawing.Point(4, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(424, 52);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "eForms";
			// 
			// cboE3
			// 
			this.cboE3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboE3.Location = new System.Drawing.Point(244, 20);
			this.cboE3.Name = "cboE3";
			this.cboE3.Size = new System.Drawing.Size(172, 21);
			this.cboE3.TabIndex = 2;
			// 
			// cboE2
			// 
			this.cboE2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboE2.Location = new System.Drawing.Point(184, 20);
			this.cboE2.Name = "cboE2";
			this.cboE2.Size = new System.Drawing.Size(56, 21);
			this.cboE2.TabIndex = 1;
			// 
			// cboE1
			// 
			this.cboE1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboE1.Location = new System.Drawing.Point(8, 20);
			this.cboE1.Name = "cboE1";
			this.cboE1.Size = new System.Drawing.Size(172, 21);
			this.cboE1.TabIndex = 0;
			// 
			// btnClose
			// 
			this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnClose.Location = new System.Drawing.Point(368, 460);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(72, 24);
			this.btnClose.TabIndex = 1;
			this.btnClose.Text = "Close";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// SettingsForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(450, 492);
			this.Controls.Add(this.btnClose);
			this.Controls.Add(this.tabControl1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "SettingsForm";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Settings";
			this.tabControl1.ResumeLayout(false);
			this.tabMatch.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Initialise combos
		/// </summary>
		private void InitCombos()
		{
			//eform match
			cboE1.Items.Add( new ComboItem( "CRFTITLE", "CRFTITLE" ) );
			cboE1.Items.Add( new ComboItem( "CRFPAGECODE", "CRFPAGECODE" ) );
			cboE1.Items.Add( new ComboItem( "CRFPAGEID", "CRFPAGEID" ) );

			cboE2.Items.Add( new ComboItem( "", "" ) );
			cboE2.Items.Add( new ComboItem( "AND", "AND" ) );
			cboE2.Items.Add( new ComboItem( "OR", "OR" ) );

			cboE3.Items.Add( new ComboItem( "", "" ) );
			cboE3.Items.Add( new ComboItem( "CRFTITLE", "CRFTITLE" ) );
			cboE3.Items.Add( new ComboItem( "CRFPAGECODE", "CRFPAGECODE" ) );
			cboE3.Items.Add( new ComboItem( "CRFPAGEID", "CRFPAGEID" ) );

			//element match
			cboEl1.Items.Add( new ComboItem( "DATAITEMNAME", "DATAITEMNAME" ) );
			cboEl1.Items.Add( new ComboItem( "DATAITEMCODE", "DATAITEMCODE" ) );
			cboEl1.Items.Add( new ComboItem( "ELEMENTCAPTION", "CAPTION" ) );
	
			cboEl2.Items.Add( new ComboItem( "", "" ) );
			cboEl2.Items.Add( new ComboItem( "AND", "AND" ) );
			cboEl2.Items.Add( new ComboItem( "OR", "OR" ) );

			cboEl3.Items.Add( new ComboItem( "", "" ) );
			cboEl3.Items.Add( new ComboItem( "DATAITEMNAME", "DATAITEMNAME" ) );
			cboEl3.Items.Add( new ComboItem( "DATAITEMCODE", "DATAITEMCODE" ) );
			cboEl3.Items.Add( new ComboItem( "ELEMENTCAPTION", "CAPTION" ) );
		}

		/// <summary>
		/// Eform match criteria 1
		/// </summary>
		public string EformMatchCriteria1
		{
			get{ return( ( ( ComboItem )cboE1.SelectedItem ).Code ); }
		}

		/// <summary>
		/// Close form
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		/// <summary>
		/// Eform match criteria 2
		/// </summary>
		public string EformMatchCriteria2
		{
			get{ return( ( ( ComboItem )cboE3.SelectedItem ).Code ); }
		}

		/// <summary>
		/// Eform match operator
		/// </summary>
		public string EformMatchOperator
		{
			get{ return( ( ( ComboItem )cboE2.SelectedItem ).Code ); }
		}

		/// <summary>
		/// Element match criteria 1
		/// </summary>
		/// 28/04/2006	bug 2725 incorrect matching of eform elements
		public string ElementMatchCriteria1
		{
			get{ return( ( ( ComboItem )cboEl1.SelectedItem ).Code ); }
		}

		/// <summary>
		/// Element match criteria 2
		/// </summary>
		/// 28/04/2006	bug 2725 incorrect matching of eform elements
		public string ElementMatchCriteria2
		{
			get{ return( ( ( ComboItem )cboEl3.SelectedItem ).Code ); }
		}

		/// <summary>
		/// Element match operator
		/// </summary>
		public string ElementMatchOperator
		{
			get{ return( ( ( ComboItem )cboEl2.SelectedItem ).Code ); }
		}
	}
}
