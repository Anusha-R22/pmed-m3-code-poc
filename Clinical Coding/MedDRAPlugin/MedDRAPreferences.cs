using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace InferMed.MACRO.ClinicalCoding.Plugins
{
	/// <summary>
	/// MedDRA preferences editing form
	/// </summary>
	public class MedDRAPreferences : System.Windows.Forms.Form
	{
		private MedDRAPreference _pref = null;

		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.CheckBox chkLLTKey;
		private System.Windows.Forms.CheckBox chkWeight;
		private System.Windows.Forms.CheckBox chkFullMatch;
		private System.Windows.Forms.CheckBox chkPartMatch;
		private System.Windows.Forms.CheckBox chkPTKey;
		private System.Windows.Forms.CheckBox chkHLTKey;
		private System.Windows.Forms.CheckBox chkHLGTKey;
		private System.Windows.Forms.CheckBox chkPT;
		private System.Windows.Forms.CheckBox chkHLT;
		private System.Windows.Forms.CheckBox chkHLGT;
		private System.Windows.Forms.CheckBox chkSOC;
		private System.Windows.Forms.CheckBox chkSOCKey;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.ComboBox comboBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.CheckBox chkLegend;
		private System.Windows.Forms.CheckBox chkPrimary;
		private System.Windows.Forms.CheckBox chkCurrent;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public MedDRAPreferences( ref MedDRAPreference pref )
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			_pref = pref;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MedDRAPreferences));
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkCurrent = new System.Windows.Forms.CheckBox();
			this.chkPrimary = new System.Windows.Forms.CheckBox();
			this.chkSOCKey = new System.Windows.Forms.CheckBox();
			this.chkSOC = new System.Windows.Forms.CheckBox();
			this.chkHLGTKey = new System.Windows.Forms.CheckBox();
			this.chkHLGT = new System.Windows.Forms.CheckBox();
			this.chkHLTKey = new System.Windows.Forms.CheckBox();
			this.chkHLT = new System.Windows.Forms.CheckBox();
			this.chkPTKey = new System.Windows.Forms.CheckBox();
			this.chkPT = new System.Windows.Forms.CheckBox();
			this.chkPartMatch = new System.Windows.Forms.CheckBox();
			this.chkFullMatch = new System.Windows.Forms.CheckBox();
			this.chkWeight = new System.Windows.Forms.CheckBox();
			this.chkLLTKey = new System.Windows.Forms.CheckBox();
			this.btnOK = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.comboBox1 = new System.Windows.Forms.ComboBox();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.chkLegend = new System.Windows.Forms.CheckBox();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.chkCurrent);
			this.groupBox1.Controls.Add(this.chkPrimary);
			this.groupBox1.Controls.Add(this.chkSOCKey);
			this.groupBox1.Controls.Add(this.chkSOC);
			this.groupBox1.Controls.Add(this.chkHLGTKey);
			this.groupBox1.Controls.Add(this.chkHLGT);
			this.groupBox1.Controls.Add(this.chkHLTKey);
			this.groupBox1.Controls.Add(this.chkHLT);
			this.groupBox1.Controls.Add(this.chkPTKey);
			this.groupBox1.Controls.Add(this.chkPT);
			this.groupBox1.Controls.Add(this.chkPartMatch);
			this.groupBox1.Controls.Add(this.chkFullMatch);
			this.groupBox1.Controls.Add(this.chkWeight);
			this.groupBox1.Controls.Add(this.chkLLTKey);
			this.groupBox1.Location = new System.Drawing.Point(4, 4);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(300, 136);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Columns";
			// 
			// chkCurrent
			// 
			this.chkCurrent.Location = new System.Drawing.Point(152, 116);
			this.chkCurrent.Name = "chkCurrent";
			this.chkCurrent.Size = new System.Drawing.Size(144, 16);
			this.chkCurrent.TabIndex = 13;
			this.chkCurrent.Text = "Current";
			// 
			// chkPrimary
			// 
			this.chkPrimary.Location = new System.Drawing.Point(8, 116);
			this.chkPrimary.Name = "chkPrimary";
			this.chkPrimary.Size = new System.Drawing.Size(144, 16);
			this.chkPrimary.TabIndex = 12;
			this.chkPrimary.Text = "Primary";
			// 
			// chkSOCKey
			// 
			this.chkSOCKey.Location = new System.Drawing.Point(152, 100);
			this.chkSOCKey.Name = "chkSOCKey";
			this.chkSOCKey.Size = new System.Drawing.Size(144, 16);
			this.chkSOCKey.TabIndex = 11;
			this.chkSOCKey.Text = "SOC Key";
			// 
			// chkSOC
			// 
			this.chkSOC.Location = new System.Drawing.Point(152, 84);
			this.chkSOC.Name = "chkSOC";
			this.chkSOC.Size = new System.Drawing.Size(144, 16);
			this.chkSOC.TabIndex = 10;
			this.chkSOC.Text = "System Organ Class";
			// 
			// chkHLGTKey
			// 
			this.chkHLGTKey.Location = new System.Drawing.Point(152, 68);
			this.chkHLGTKey.Name = "chkHLGTKey";
			this.chkHLGTKey.Size = new System.Drawing.Size(144, 16);
			this.chkHLGTKey.TabIndex = 9;
			this.chkHLGTKey.Text = "HLGT Key";
			// 
			// chkHLGT
			// 
			this.chkHLGT.Location = new System.Drawing.Point(152, 52);
			this.chkHLGT.Name = "chkHLGT";
			this.chkHLGT.Size = new System.Drawing.Size(144, 16);
			this.chkHLGT.TabIndex = 8;
			this.chkHLGT.Text = "High Level Group Term";
			// 
			// chkHLTKey
			// 
			this.chkHLTKey.Location = new System.Drawing.Point(152, 36);
			this.chkHLTKey.Name = "chkHLTKey";
			this.chkHLTKey.Size = new System.Drawing.Size(144, 16);
			this.chkHLTKey.TabIndex = 7;
			this.chkHLTKey.Text = "HLT Key";
			// 
			// chkHLT
			// 
			this.chkHLT.Location = new System.Drawing.Point(152, 20);
			this.chkHLT.Name = "chkHLT";
			this.chkHLT.Size = new System.Drawing.Size(144, 16);
			this.chkHLT.TabIndex = 6;
			this.chkHLT.Text = "High Level Term";
			// 
			// chkPTKey
			// 
			this.chkPTKey.Location = new System.Drawing.Point(8, 100);
			this.chkPTKey.Name = "chkPTKey";
			this.chkPTKey.Size = new System.Drawing.Size(144, 16);
			this.chkPTKey.TabIndex = 5;
			this.chkPTKey.Text = "PT Key";
			// 
			// chkPT
			// 
			this.chkPT.Location = new System.Drawing.Point(8, 84);
			this.chkPT.Name = "chkPT";
			this.chkPT.Size = new System.Drawing.Size(144, 16);
			this.chkPT.TabIndex = 4;
			this.chkPT.Text = "Preferred Term";
			// 
			// chkPartMatch
			// 
			this.chkPartMatch.Location = new System.Drawing.Point(8, 68);
			this.chkPartMatch.Name = "chkPartMatch";
			this.chkPartMatch.Size = new System.Drawing.Size(144, 16);
			this.chkPartMatch.TabIndex = 3;
			this.chkPartMatch.Text = "Part Match";
			// 
			// chkFullMatch
			// 
			this.chkFullMatch.Location = new System.Drawing.Point(8, 52);
			this.chkFullMatch.Name = "chkFullMatch";
			this.chkFullMatch.Size = new System.Drawing.Size(144, 16);
			this.chkFullMatch.TabIndex = 2;
			this.chkFullMatch.Text = "Full Match";
			// 
			// chkWeight
			// 
			this.chkWeight.Location = new System.Drawing.Point(8, 36);
			this.chkWeight.Name = "chkWeight";
			this.chkWeight.Size = new System.Drawing.Size(144, 16);
			this.chkWeight.TabIndex = 1;
			this.chkWeight.Text = "Weight";
			// 
			// chkLLTKey
			// 
			this.chkLLTKey.Location = new System.Drawing.Point(8, 20);
			this.chkLLTKey.Name = "chkLLTKey";
			this.chkLLTKey.Size = new System.Drawing.Size(144, 16);
			this.chkLLTKey.TabIndex = 0;
			this.chkLLTKey.Text = "LLT Key";
			// 
			// btnOK
			// 
			this.btnOK.Location = new System.Drawing.Point(224, 236);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(72, 24);
			this.btnOK.TabIndex = 1;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label2);
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Controls.Add(this.comboBox1);
			this.groupBox2.Location = new System.Drawing.Point(4, 144);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(300, 44);
			this.groupBox2.TabIndex = 2;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Results";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(156, 20);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(52, 16);
			this.label2.TabIndex = 2;
			this.label2.Text = "matches";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(12, 20);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(52, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Display";
			// 
			// comboBox1
			// 
			this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBox1.Location = new System.Drawing.Point(68, 16);
			this.comboBox1.Name = "comboBox1";
			this.comboBox1.Size = new System.Drawing.Size(80, 21);
			this.comboBox1.TabIndex = 0;
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.chkLegend);
			this.groupBox3.Location = new System.Drawing.Point(4, 192);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(300, 36);
			this.groupBox3.TabIndex = 3;
			this.groupBox3.TabStop = false;
			// 
			// chkLegend
			// 
			this.chkLegend.Location = new System.Drawing.Point(12, 12);
			this.chkLegend.Name = "chkLegend";
			this.chkLegend.Size = new System.Drawing.Size(136, 16);
			this.chkLegend.TabIndex = 0;
			this.chkLegend.Text = "Show legend";
			// 
			// MedDRAPreferences
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(306, 268);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.groupBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "MedDRAPreferences";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Preferences";
			this.Load += new System.EventHandler(this.MedDRAPreferences_Load);
			this.groupBox1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// OK
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnOK_Click(object sender, System.EventArgs e)
		{
			_pref._lltKey = chkLLTKey.Checked;
			_pref._weight = chkWeight.Checked;
			_pref._fullMatch = chkFullMatch.Checked;
			_pref._partMatch = chkPartMatch.Checked;
			_pref._primary = chkPrimary.Checked;
			_pref._current = chkCurrent.Checked;
			_pref._pt = chkPT.Checked;
			_pref._ptKey = chkPTKey.Checked;
			_pref._hlt = chkHLT.Checked;
			_pref._hltKey = chkHLTKey.Checked;
			_pref._hlgt = chkHLGT.Checked;
			_pref._hlgtKey = chkHLGTKey.Checked;
			_pref._soc = chkSOC.Checked;
			_pref._socKey = chkSOCKey.Checked;
			ComboItem ci = ( ComboItem )comboBox1.SelectedItem;
			_pref._result = System.Convert.ToInt32( ci.Code );
			_pref._legend = chkLegend.Checked;

			this.Close();
		}

		/// <summary>
		/// Set form fields
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void MedDRAPreferences_Load(object sender, System.EventArgs e)
		{
			comboBox1.Items.Add( new ComboItem( "top 50", "50" ) );
			comboBox1.Items.Add( new ComboItem( "top 100", "100" ) );
			comboBox1.Items.Add( new ComboItem( "top 500", "500" ) );
			comboBox1.Items.Add( new ComboItem( "top 1000", "1000" ) );
			comboBox1.Items.Add( new ComboItem( "All", "0" ) );

			chkLLTKey.Checked = _pref._lltKey;
			chkWeight.Checked = _pref._weight;
			chkFullMatch.Checked = _pref._fullMatch;
			chkPartMatch.Checked = _pref._partMatch;
			chkPrimary.Checked = _pref._primary;
			chkCurrent.Checked = _pref._current;
			chkPT.Checked = _pref._pt;
			chkPTKey.Checked = _pref._ptKey;
			chkHLT.Checked = _pref._hlt;
			chkHLTKey.Checked = _pref._hltKey;
			chkHLGT.Checked = _pref._hlgt;
			chkHLGTKey.Checked = _pref._hlgtKey;
			chkSOC.Checked = _pref._soc;
			chkSOCKey.Checked = _pref._socKey;
			chkLegend.Checked = _pref._legend;

			switch( _pref._result )
			{
				case 50:
					comboBox1.SelectedIndex = 0;
					break;
				case 100:
					comboBox1.SelectedIndex = 1;
					break;
				case 1000:
					comboBox1.SelectedIndex = 3;
					break;
				case 0:
					comboBox1.SelectedIndex = 4;
					break;
				default:
					comboBox1.SelectedIndex = 2;
					break;

			}
		}
	}
}
