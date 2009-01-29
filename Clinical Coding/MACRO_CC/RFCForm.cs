using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace InferMed.MACRO.ClinicalCoding.MACRO_CC
{
	/// <summary>
	/// Summary description for RFCForm.
	/// </summary>
	public class RFCForm : System.Windows.Forms.Form
	{
		public const string _FORBIDDEN_CHARS = "`¬|~\"";
		private string _rfc = "";

		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtRFC;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RFCForm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(RFCForm));
			this.button1 = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.txtRFC = new System.Windows.Forms.TextBox();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(212, 80);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(72, 24);
			this.button1.TabIndex = 0;
			this.button1.Text = "OK";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.txtRFC);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Location = new System.Drawing.Point(4, 4);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(284, 72);
			this.groupBox1.TabIndex = 1;
			this.groupBox1.TabStop = false;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(268, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Please enter a reason for changing the coding value.";
			// 
			// txtRFC
			// 
			this.txtRFC.Location = new System.Drawing.Point(8, 40);
			this.txtRFC.Name = "txtRFC";
			this.txtRFC.Size = new System.Drawing.Size(268, 20);
			this.txtRFC.TabIndex = 1;
			this.txtRFC.Text = "";
			// 
			// RFCForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 108);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.button1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "RFCForm";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Reason For Change";
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void button1_Click(object sender, System.EventArgs e)
		{
			if( ( txtRFC.Text.Length > 0 ) && ( txtRFC.Text.Length < 255 ) )
			{
				if( CharIsOK( txtRFC.Text ) )
				{
					_rfc = txtRFC.Text;
					this.Close();
				}
				else
				{
					MessageBox.Show( "A reason for change may not contain the following characters: " + _FORBIDDEN_CHARS );
				}
			}
			else
			{
				MessageBox.Show( "Please enter a reason for change no longer than 255 characters" );
			}
		}

		public string RFC
		{
			get { return( _rfc ); }
		}

		private static bool CharIsOK( string s )
		{
			for( int n = 0; n < _FORBIDDEN_CHARS.Length; n++ )
			{
				if( s == _FORBIDDEN_CHARS[n].ToString() ) return false;
			}
			return( true );
		}
	}
}
