using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using InferMed.MACRO.API;

namespace MACROAPI_dot_net_test_harness
{
	/// <summary>
	/// Summary description for LoginAPI.
	/// </summary>
	public class LoginAPI : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblUsername;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Button cmdLogin;
		private System.Windows.Forms.Button cmdCancel;
		private System.Windows.Forms.TextBox txtUsername;
		private System.Windows.Forms.TextBox txtPassword;
		private System.Windows.Forms.TextBox txtDatabase;
		private System.Windows.Forms.TextBox txtRole;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox txtSecConn;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public LoginAPI()
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
			this.lblUsername = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtUsername = new System.Windows.Forms.TextBox();
			this.txtPassword = new System.Windows.Forms.TextBox();
			this.txtDatabase = new System.Windows.Forms.TextBox();
			this.txtRole = new System.Windows.Forms.TextBox();
			this.cmdLogin = new System.Windows.Forms.Button();
			this.cmdCancel = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.txtSecConn = new System.Windows.Forms.TextBox();
			this.SuspendLayout();
			// 
			// lblUsername
			// 
			this.lblUsername.Location = new System.Drawing.Point(32, 16);
			this.lblUsername.Name = "lblUsername";
			this.lblUsername.Size = new System.Drawing.Size(104, 23);
			this.lblUsername.TabIndex = 0;
			this.lblUsername.Text = "Username";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(32, 48);
			this.label1.Name = "label1";
			this.label1.TabIndex = 1;
			this.label1.Text = "Password";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(32, 80);
			this.label2.Name = "label2";
			this.label2.TabIndex = 2;
			this.label2.Text = "Database";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(32, 112);
			this.label3.Name = "label3";
			this.label3.TabIndex = 3;
			this.label3.Text = "Role";
			// 
			// txtUsername
			// 
			this.txtUsername.Location = new System.Drawing.Point(144, 16);
			this.txtUsername.Name = "txtUsername";
			this.txtUsername.Size = new System.Drawing.Size(152, 20);
			this.txtUsername.TabIndex = 4;
			this.txtUsername.Text = "rde";
			// 
			// txtPassword
			// 
			this.txtPassword.Location = new System.Drawing.Point(144, 48);
			this.txtPassword.Name = "txtPassword";
			this.txtPassword.PasswordChar = '*';
			this.txtPassword.Size = new System.Drawing.Size(152, 20);
			this.txtPassword.TabIndex = 5;
			this.txtPassword.Text = "macrotm";
			// 
			// txtDatabase
			// 
			this.txtDatabase.Location = new System.Drawing.Point(144, 80);
			this.txtDatabase.Name = "txtDatabase";
			this.txtDatabase.Size = new System.Drawing.Size(152, 20);
			this.txtDatabase.TabIndex = 6;
			this.txtDatabase.Text = "MACRODPH3";
			// 
			// txtRole
			// 
			this.txtRole.Location = new System.Drawing.Point(144, 112);
			this.txtRole.Name = "txtRole";
			this.txtRole.Size = new System.Drawing.Size(152, 20);
			this.txtRole.TabIndex = 7;
			this.txtRole.Text = "MACROUser";
			// 
			// cmdLogin
			// 
			this.cmdLogin.Location = new System.Drawing.Point(120, 192);
			this.cmdLogin.Name = "cmdLogin";
			this.cmdLogin.TabIndex = 8;
			this.cmdLogin.Text = "Login";
			this.cmdLogin.Click += new System.EventHandler(this.cmdLogin_Click);
			// 
			// cmdCancel
			// 
			this.cmdCancel.Location = new System.Drawing.Point(216, 192);
			this.cmdCancel.Name = "cmdCancel";
			this.cmdCancel.TabIndex = 9;
			this.cmdCancel.Text = "Cancel";
			this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(32, 144);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(100, 32);
			this.label4.TabIndex = 10;
			this.label4.Text = "Security Database conn string";
			// 
			// txtSecConn
			// 
			this.txtSecConn.Location = new System.Drawing.Point(144, 144);
			this.txtSecConn.Name = "txtSecConn";
			this.txtSecConn.Size = new System.Drawing.Size(152, 20);
			this.txtSecConn.TabIndex = 11;
			this.txtSecConn.Text = "Provider=SQLOLEDB;Data Source=HOOKD;Initial Catalog=MainSecurity;User ID=sa;pwd=m" +
				"acrotm;";
			// 
			// LoginAPI
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(336, 230);
			this.Controls.Add(this.txtSecConn);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.cmdCancel);
			this.Controls.Add(this.cmdLogin);
			this.Controls.Add(this.txtRole);
			this.Controls.Add(this.txtDatabase);
			this.Controls.Add(this.txtPassword);
			this.Controls.Add(this.txtUsername);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblUsername);
			this.Name = "LoginAPI";
			this.Text = "API Login Page";
			this.ResumeLayout(false);

		}
		#endregion

		private void cmdLogin_Click(object sender, System.EventArgs e)
		{
			string userName = txtUsername.Text;
			string passWord = txtPassword.Text;
			string macroDb = txtDatabase.Text;
			string macroRole = txtRole.Text;
			string message = "";
			string serialisedUser = "";
			string userNameFull = "";
			string secCon = txtSecConn.Text;

			// Login result
			API.LoginResult loginResult = API.LoginResult.Failed;
			// if security connection is opened
			if( secCon != "" )
			{
				// security db login
				loginResult = API.LoginSecurity( userName, passWord, macroDb, macroRole, secCon,
								ref message, ref userNameFull, ref serialisedUser );
			}
			else
			{
				// normal login
				loginResult = API.Login( userName, passWord, macroDb, macroRole, 
					ref message, ref userNameFull, ref serialisedUser );
			}
			// open 
			if ( loginResult == API.LoginResult.Success )
			{
				// open API tester eForm
				APITester apiTester = new APITester(serialisedUser);
				apiTester.ShowDialog();
				this.Close();
			}
			else
			{
				MessageBox.Show( "Login Failed - " + message );
				// close the eForm
				this.Close();
			}
		}

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new LoginAPI());
		}

		private void cmdCancel_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}
	}
}
