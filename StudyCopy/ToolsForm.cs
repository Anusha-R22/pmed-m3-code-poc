using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Text;
using System.IO;
using log4net;

namespace InferMed.MACRO.StudyCopy
{
	/// <summary>
	/// Summary description for StudyComparison.
	/// </summary>
	public class ToolsForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.PictureBox pictureBox1;
		
		private System.Windows.Forms.Button btnCreateReport;

		private IDbConnection _dbConn = null;
		private string _studyIdD = "";
		private string _studyIdS = "";
		private string _studyCodeD = "";
		private string _studyCodeS = "";
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.CheckBox chkEforms;
		private System.Windows.Forms.CheckBox chkDataitems;
		private StudyState _state = null;

		//logging
		private static readonly ILog log = LogManager.GetLogger( typeof( ToolsForm ) );

		public ToolsForm( IDbConnection dbConn, string studyIdD, string studyCodeD, string studyIdS, string studyCodeS, 
			StudyState state )
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			_dbConn = dbConn;
			_studyIdD = studyIdD;
			_studyIdS = studyIdS;
			_studyCodeD = studyCodeD;
			_studyCodeS = studyCodeS;
			_state = state;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(ToolsForm));
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.label1 = new System.Windows.Forms.Label();
			this.btnCreateReport = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.chkEforms = new System.Windows.Forms.CheckBox();
			this.chkDataitems = new System.Windows.Forms.CheckBox();
			this.tabControl1.SuspendLayout();
			this.tabPage1.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// tabControl1
			// 
			this.tabControl1.Controls.Add(this.tabPage1);
			this.tabControl1.Location = new System.Drawing.Point(4, 4);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(440, 448);
			this.tabControl1.TabIndex = 0;
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.groupBox1);
			this.tabPage1.Location = new System.Drawing.Point(4, 22);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(432, 422);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "Reporting";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.chkDataitems);
			this.groupBox1.Controls.Add(this.chkEforms);
			this.groupBox1.Controls.Add(this.pictureBox1);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.btnCreateReport);
			this.groupBox1.Location = new System.Drawing.Point(4, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(424, 140);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Study Differences";
			// 
			// pictureBox1
			// 
			this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
			this.pictureBox1.Location = new System.Drawing.Point(12, 24);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(36, 28);
			this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pictureBox1.TabIndex = 4;
			this.pictureBox1.TabStop = false;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(60, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(356, 32);
			this.label1.TabIndex = 3;
			this.label1.Text = "Creates a list of differences between the eForms and eForm elements in the two se" +
				"lected studies.";
			// 
			// btnCreateReport
			// 
			this.btnCreateReport.Location = new System.Drawing.Point(280, 104);
			this.btnCreateReport.Name = "btnCreateReport";
			this.btnCreateReport.Size = new System.Drawing.Size(132, 24);
			this.btnCreateReport.TabIndex = 2;
			this.btnCreateReport.Text = "Create HTML Report";
			this.btnCreateReport.Click += new System.EventHandler(this.btnCreateReport_Click);
			// 
			// btnClose
			// 
			this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnClose.Location = new System.Drawing.Point(368, 460);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(72, 24);
			this.btnClose.TabIndex = 3;
			this.btnClose.Text = "Close";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// chkEforms
			// 
			this.chkEforms.Checked = true;
			this.chkEforms.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkEforms.Location = new System.Drawing.Point(68, 60);
			this.chkEforms.Name = "chkEforms";
			this.chkEforms.Size = new System.Drawing.Size(208, 20);
			this.chkEforms.TabIndex = 5;
			this.chkEforms.Text = "Include eForms and CRFElements";
			// 
			// chkDataitems
			// 
			this.chkDataitems.Checked = true;
			this.chkDataitems.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDataitems.Location = new System.Drawing.Point(68, 80);
			this.chkDataitems.Name = "chkDataitems";
			this.chkDataitems.Size = new System.Drawing.Size(208, 20);
			this.chkDataitems.TabIndex = 6;
			this.chkDataitems.Text = "Include dataitems";
			// 
			// ToolsForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(448, 490);
			this.Controls.Add(this.btnClose);
			this.Controls.Add(this.tabControl1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "ToolsForm";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Tools";
			this.tabControl1.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void Processing( bool on )
		{
			//enable/disable the form
			this.Cursor = ( on ) ? Cursors.WaitCursor : Cursors.Default;
			groupBox1.Enabled = !on;
		}

		private void btnCreateReport_Click(object sender, System.EventArgs e)
		{
			try
			{
				Processing( true );

				saveFileDialog1.Filter = "html files (*.html)|*.html"  ;
				if( saveFileDialog1.ShowDialog() == DialogResult.OK )
				{
					string fileName = saveFileDialog1.FileName;

					StringBuilder sb = HTML.CreateReport( _dbConn, _studyIdD, _studyCodeD, _studyIdS, _studyCodeS, _state, 
						chkEforms.Checked, chkDataitems.Checked );

					StreamWriter sw = new StreamWriter( fileName );
					sw.Write( sb );
					sw.Flush();
					sw.Close();

					MessageBox.Show( "The report '" + fileName + "' was created successfully", 
						"MACRO Study Copy Tool", MessageBoxButtons.OK, MessageBoxIcon.Information );
				}
			}
#if( !DEBUG )
			catch( Exception ex )
			{
				//logging
				log.Error( ex.Message + " : " + ex.InnerException );
					MessageBox.Show( "An exception occurred : " + ex.Message + " : " + ex.InnerException, 
					"MACRO Study Copy Tool", MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
#endif
			finally
			{
				Processing( false );
				MACRO30.ShowProgress( "Ready" );
			}
		}
	}
}
