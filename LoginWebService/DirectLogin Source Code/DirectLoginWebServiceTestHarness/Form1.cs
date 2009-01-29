using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using MACRO30.DirectLogin;

namespace MACRODirectLoginTestHarness
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MACRODirectLogin x = new MACRODirectLogin();
            x.Url = txtURL.Text;
            string userXML = "";
            string errorXML = "";
            int result = x.Login(txtUser.Text, txtPassword.Text, out userXML, out errorXML);
            StringBuilder sbResult = new StringBuilder();
            switch (result)
            {
                case 0:
                    {
                        sbResult.Append("Success");
                        break;
                    }
                case 1:
                    {
                        sbResult.Append("Account Disabled");
                        break;
                    }
                case 2:
                    {
                        sbResult.Append("Failed");
                        break;
                    }
                case 3:
                    {
                        sbResult.Append("Change Password");
                        break;
                    }
                case 4:
                    {
                        sbResult.Append("Password Expired");
                        break;
                    }
            }
            txtResult.Text = sbResult.ToString();
            rtbXml.Text = userXML;
        }
    }
}