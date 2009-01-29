using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using InferMed.Components;
using InferMed.MACRO.ClinicalCoding;
using System.IO;
using log4net.Config;
using log4net;

namespace InferMed.MACRO.ClinicalCoding.MedDRAPreloader
{
	/// <summary>
	/// Dictionary preloader splashscreen
	/// </summary>
	public class LoadForm : System.Windows.Forms.Form
	{

		//logging
		private static readonly ILog log = LogManager.GetLogger( typeof( LoadForm ) );

		private bool _loading = false;

		private const string _DEF_USER_SETTINGS_FILE = @"C:\Program Files\InferMed\MACRO 3.0\MACROUserSettings30.txt";
		private const string _DEF_MACRO_PATH = @"C:\Program Files\InferMed\MACRO 3.0\";
		private const string _SETTINGS_FILE = @"MACROSettings30.txt";
		private const string _USER_SETTINGS_FILE = "settingsfile";
		private const string _PRELOADLIST = "meddrapreloadlist";
		private const string _SECDB = "securityPath";

		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label lblLoading;
		private System.Windows.Forms.Label lblApp;
		private System.Windows.Forms.Label lblCompany;
		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.Label lblDescription;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public LoadForm()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(LoadForm));
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.lblLoading = new System.Windows.Forms.Label();
			this.lblApp = new System.Windows.Forms.Label();
			this.lblCompany = new System.Windows.Forms.Label();
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.lblDescription = new System.Windows.Forms.Label();
			this.groupBox2.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox2
			// 
			this.groupBox2.BackColor = System.Drawing.Color.White;
			this.groupBox2.Controls.Add(this.lblDescription);
			this.groupBox2.Controls.Add(this.lblLoading);
			this.groupBox2.Controls.Add(this.lblApp);
			this.groupBox2.Controls.Add(this.lblCompany);
			this.groupBox2.Location = new System.Drawing.Point(180, 0);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(212, 132);
			this.groupBox2.TabIndex = 5;
			this.groupBox2.TabStop = false;
			// 
			// lblLoading
			// 
			this.lblLoading.Location = new System.Drawing.Point(4, 76);
			this.lblLoading.Name = "lblLoading";
			this.lblLoading.Size = new System.Drawing.Size(204, 44);
			this.lblLoading.TabIndex = 2;
			this.lblLoading.Click += new System.EventHandler(this.LoadForm_Click);
			// 
			// lblApp
			// 
			this.lblApp.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblApp.Location = new System.Drawing.Point(4, 16);
			this.lblApp.Name = "lblApp";
			this.lblApp.Size = new System.Drawing.Size(204, 16);
			this.lblApp.TabIndex = 1;
			this.lblApp.Click += new System.EventHandler(this.LoadForm_Click);
			// 
			// lblCompany
			// 
			this.lblCompany.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblCompany.Location = new System.Drawing.Point(4, 32);
			this.lblCompany.Name = "lblCompany";
			this.lblCompany.Size = new System.Drawing.Size(204, 16);
			this.lblCompany.TabIndex = 0;
			this.lblCompany.Click += new System.EventHandler(this.LoadForm_Click);
			// 
			// pictureBox1
			// 
			this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
			this.pictureBox1.Location = new System.Drawing.Point(4, 4);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(176, 128);
			this.pictureBox1.TabIndex = 4;
			this.pictureBox1.TabStop = false;
			this.pictureBox1.Click += new System.EventHandler(this.LoadForm_Click);
			// 
			// lblDescription
			// 
			this.lblDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescription.Location = new System.Drawing.Point(4, 44);
			this.lblDescription.Name = "lblDescription";
			this.lblDescription.Size = new System.Drawing.Size(204, 16);
			this.lblDescription.TabIndex = 3;
			// 
			// LoadForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.BackColor = System.Drawing.Color.White;
			this.ClientSize = new System.Drawing.Size(396, 136);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.pictureBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "LoadForm";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Form1";
			this.Click += new System.EventHandler(this.LoadForm_Click);
			this.Load += new System.EventHandler(this.LoadForm_Load);
			this.Activated += new System.EventHandler(this.LoadForm_Activate);
			this.groupBox2.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new LoadForm());
		}

		/// <summary>
		/// Returns the MACRO usersettings file path from the settings file
		/// </summary>
		/// <returns>User settings file path</returns>
		private static string GetUserSettingsFilePath()
		{
			string settingsFile = AppDomain.CurrentDomain.BaseDirectory  + _SETTINGS_FILE;
			if( !File.Exists( settingsFile ) )
			{
				Exception ex = new Exception( "Settings file '" + settingsFile + "' not found" );
				throw ex;
			}
			IMEDSettings20 iset = new IMEDSettings20( settingsFile );
			string userSettingsFile = iset.GetKeyValue( _USER_SETTINGS_FILE, _DEF_USER_SETTINGS_FILE );
			if( !File.Exists( userSettingsFile ) )
			{
				Exception ex = new Exception( "User settings file '" + userSettingsFile + "' not found" );
				throw ex;
			}
			return( userSettingsFile );
		}

		/// <summary>
		/// Returns a setting from the MACRO 3.0 user settings file
		/// </summary>
		/// <param name="key">Setting key name</param>
		/// <param name="defaultVal">Default value to be returned if the key is not found</param>
		/// <returns>MACRO setting value</returns>
		private static string GetSetting(string key, string defaultVal)
		{
			IMEDSettings20 iset = new IMEDSettings20( GetUserSettingsFilePath() );
			return( iset.GetKeyValue( key, defaultVal ) );
		}

		/// <summary>
		/// Hide the splashscreen if clicked
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void LoadForm_Click(object sender, System.EventArgs e)
		{
			this.Visible = false;
		}

		/// <summary>
		/// Show current progress
		/// </summary>
		/// <param name="prog"></param>
		private void ShowProgress( string prog )
		{
			lblLoading.Text = prog;
			Application.DoEvents();
		}

		/// <summary>
		/// Preload
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void LoadForm_Load(object sender, System.EventArgs e)
		{
			//initialise logging
			XmlConfigurator.Configure( new FileInfo( Path.GetDirectoryName( AppDomain.CurrentDomain.BaseDirectory ) 
				+ @"\log4netconfig.xml" ) );

			//set the labels
			lblCompany.Text = Application.CompanyName + " (c) 2006";
			lblApp.Text = "MedDRA Preloader v" + Application.ProductVersion;
			lblDescription.Text = "MedDRA dictionary preloading service";
		}

		/// <summary>
		/// Load
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void LoadForm_Activate(object sender, System.EventArgs e)
		{
			//dont 'activate' more than once
			if( _loading ) return;
			_loading = true;

			//set delimiter
			char[] del1 = "|".ToCharArray();

			Plugins.MedDRA m = null;

			try
			{
				//get the preload list from the settings file : meddrapreloadlist=SQL_M3|1|3~IC_M30|2
				ShowProgress( "Reading initialisation settings" );
				log.Info( "Reading initialisation settings" );
				string plList = GetSetting( _PRELOADLIST, "" );

				//if the setting was found
				if( plList == "" )
				{
					log.Info( "No preload list" );
				}
				else
				{
					log.Debug( "plList=" + plList );
					log.Info( "Getting security database connection string" );

					//get the security database connection string from the settings file
					string secDb = IMEDEncryption.DecryptString( GetSetting( _SECDB, "" ) );

					//if the connection was found
					if( secDb != "" )
					{
						//split on delimiter into dictionaries: from 1|2|3 to 1 2 3
						string[] plDictionaries = plList.Split( del1 );
								
						//if there are any delimited dictionaries in the list
						if( plDictionaries.Length > 0 )
						{
							//load the dictionaries object
							MACROCCBS30.Dictionaries dictionaries = new InferMed.MACRO.ClinicalCoding.MACROCCBS30.Dictionaries();
							dictionaries.Init( secDb );

							//if this database has dictionaries installed
							if( dictionaries.Count > 0 )
							{
								//if this is the first time in this loop
								if( m == null )
								{
									log.Info( "Connecting to server" );

									//initialise the meddra server object
									MACROCCBS30.Dictionary tempD = ( MACROCCBS30.Dictionary )dictionaries.DictionaryList[0];
									ShowProgress( "Connecting to server..." );
									m = new InferMed.MACRO.ClinicalCoding.Plugins.MedDRA( tempD.Custom );
								}

								//loop through the dictionary ids, loading each one
								for( int dic = 0; dic < plDictionaries.Length; dic++ )
								{
									//get this dictionary object
									MACROCCBS30.Dictionary dictionary = dictionaries.DictionaryFromId( System.Convert.ToInt32( plDictionaries[dic].ToString() ) );
									//if it exists
									if( dictionary != null )
									{
										log.Debug( "Loading dictionary " + dictionary.Name + dictionary.Version );

										//load the dictionary into the server
										ShowProgress( "Loading dictionary " + dictionary.Name + dictionary.Version + "..." );
										m.SetDictionary( dictionary.Custom );
									}
								}
							}
						}
					}
				}
			}
			catch( Exception ex )
			{
				log.Error( ex.Message, ex );
				this.Close();
			}									
			finally
			{
				if( m != null )
				{
					log.Info( "Disconnecting from server" );
					ShowProgress( "Disconnecting from server..." );
					m.Dispose();
				}
				Application.Exit();
			}
		}
	}
}
