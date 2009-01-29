using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace InferMed.MACRO.ClinicalCoding.MACRO_CC
{
	/// <summary>
	/// Summary description for DBConnectionForm.
	/// </summary>
	public class DBConnectionForm : System.Windows.Forms.Form
	{
		private string _dbCon = "";

		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.RadioButton optORA;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.TextBox txtUID;
		private System.Windows.Forms.TextBox txtPassword;
		private System.Windows.Forms.TextBox txtDBName;
		private System.Windows.Forms.TextBox txtServer;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.TextBox txtDBAlias;
		private System.Windows.Forms.RadioButton optSQL;
		private System.Windows.Forms.Button btnCancel;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DBConnectionForm()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(DBConnectionForm));
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label5 = new System.Windows.Forms.Label();
			this.txtDBAlias = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtServer = new System.Windows.Forms.TextBox();
			this.txtDBName = new System.Windows.Forms.TextBox();
			this.txtPassword = new System.Windows.Forms.TextBox();
			this.txtUID = new System.Windows.Forms.TextBox();
			this.optORA = new System.Windows.Forms.RadioButton();
			this.optSQL = new System.Windows.Forms.RadioButton();
			this.btnOK = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.groupBox2);
			this.groupBox1.Controls.Add(this.optORA);
			this.groupBox1.Controls.Add(this.optSQL);
			this.groupBox1.Location = new System.Drawing.Point(4, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(284, 212);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label5);
			this.groupBox2.Controls.Add(this.txtDBAlias);
			this.groupBox2.Controls.Add(this.label4);
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.label2);
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Controls.Add(this.txtServer);
			this.groupBox2.Controls.Add(this.txtDBName);
			this.groupBox2.Controls.Add(this.txtPassword);
			this.groupBox2.Controls.Add(this.txtUID);
			this.groupBox2.Location = new System.Drawing.Point(4, 64);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(276, 140);
			this.groupBox2.TabIndex = 2;
			this.groupBox2.TabStop = false;
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(8, 112);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(72, 16);
			this.label5.TabIndex = 11;
			this.label5.Text = "Server";
			// 
			// txtDBAlias
			// 
			this.txtDBAlias.Location = new System.Drawing.Point(112, 64);
			this.txtDBAlias.Name = "txtDBAlias";
			this.txtDBAlias.Size = new System.Drawing.Size(156, 20);
			this.txtDBAlias.TabIndex = 2;
			this.txtDBAlias.Text = "";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(8, 88);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(100, 12);
			this.label4.TabIndex = 8;
			this.label4.Text = "DB Name";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(8, 64);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(72, 12);
			this.label3.TabIndex = 7;
			this.label3.Text = "DB Alias";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(72, 16);
			this.label2.TabIndex = 6;
			this.label2.Text = "Password";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 16);
			this.label1.TabIndex = 5;
			this.label1.Text = "UID";
			// 
			// txtServer
			// 
			this.txtServer.Location = new System.Drawing.Point(112, 112);
			this.txtServer.Name = "txtServer";
			this.txtServer.Size = new System.Drawing.Size(156, 20);
			this.txtServer.TabIndex = 4;
			this.txtServer.Text = "";
			// 
			// txtDBName
			// 
			this.txtDBName.Location = new System.Drawing.Point(112, 88);
			this.txtDBName.Name = "txtDBName";
			this.txtDBName.Size = new System.Drawing.Size(156, 20);
			this.txtDBName.TabIndex = 3;
			this.txtDBName.Text = "";
			// 
			// txtPassword
			// 
			this.txtPassword.Location = new System.Drawing.Point(112, 40);
			this.txtPassword.Name = "txtPassword";
			this.txtPassword.PasswordChar = '*';
			this.txtPassword.Size = new System.Drawing.Size(156, 20);
			this.txtPassword.TabIndex = 1;
			this.txtPassword.Text = "";
			// 
			// txtUID
			// 
			this.txtUID.Location = new System.Drawing.Point(112, 16);
			this.txtUID.Name = "txtUID";
			this.txtUID.Size = new System.Drawing.Size(156, 20);
			this.txtUID.TabIndex = 0;
			this.txtUID.Text = "";
			// 
			// optORA
			// 
			this.optORA.Location = new System.Drawing.Point(8, 40);
			this.optORA.Name = "optORA";
			this.optORA.Size = new System.Drawing.Size(132, 20);
			this.optORA.TabIndex = 1;
			this.optORA.Text = "ORACLE";
			this.optORA.CheckedChanged += new System.EventHandler(this.optORA_CheckedChanged);
			// 
			// optSQL
			// 
			this.optSQL.Checked = true;
			this.optSQL.Location = new System.Drawing.Point(8, 20);
			this.optSQL.Name = "optSQL";
			this.optSQL.Size = new System.Drawing.Size(132, 20);
			this.optSQL.TabIndex = 0;
			this.optSQL.TabStop = true;
			this.optSQL.Text = "SQL Server";
			// 
			// btnOK
			// 
			this.btnOK.Location = new System.Drawing.Point(216, 228);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(64, 24);
			this.btnOK.TabIndex = 2;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// btnCancel
			// 
			this.btnCancel.Location = new System.Drawing.Point(148, 228);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(64, 24);
			this.btnCancel.TabIndex = 1;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// DBConnectionForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 260);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.groupBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "DBConnectionForm";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Database Connection";
			this.groupBox1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void optORA_CheckedChanged(object sender, System.EventArgs e)
		{
			if( optORA.Checked )
			{
				label5.Visible = false;
				txtServer.Visible = false;
				label4.Text = "Net service name";
			}
			else
			{
				label5.Visible = true;
				txtServer.Visible = true;
				label4.Text = "DB Name";
			}
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			string con = "";

			if( optORA.Checked )
			{
				con += "PROVIDER=MSDAORA;"
					+  "DATA SOURCE=" + txtDBName.Text + ";"
					+  "DATABASE=" + txtDBAlias.Text + ";"
					+  "USER ID=" + txtUID.Text + ";"
					+	 "PASSWORD=" + txtPassword.Text + ";";
			}
			else
			{
				con += "PROVIDER=SQLOLEDB;"
					+  "DATA SOURCE=" + txtServer.Text + ";"
					+  "DATABASE=" + txtDBName.Text + ";"
					+  "USER ID=" + txtUID.Text + ";"
					+	 "PASSWORD=" + txtPassword.Text + ";";
			}

			if( !DBUpgrade.DatabaseExists( con ) )
			{
				MessageBox.Show( "A connection to the specified database cannot be established. Please check the details", 
					"Database Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation );
			}
			else
			{
				_dbCon = con;
				this.Close();
			}
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		public string DBCon
		{
			get{ return( _dbCon ); }
		}
	}
}
