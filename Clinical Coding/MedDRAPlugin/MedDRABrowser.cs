using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using log4net;

namespace InferMed.MACRO.ClinicalCoding.Plugins
{
	/// <summary>
	/// MedDRA browsing form
	/// </summary>
	public class MedDRABrowser : System.Windows.Forms.Form
	{
		//logging
		private static readonly ILog log = LogManager.GetLogger( typeof( MedDRABrowser ) );

		//imagelist images
		private const int _IMAGEBESTMATCH = 0;
		private const int _IMAGEHISTORICALMATCH = 1;
		private const int _IMAGEBESTHISTORICALMATCH = 2;
		//has a term been accepted by the user
		public bool _accepted = false;
		//MedDRA browser properties
		private string _custom = "";
		public string _dictionary = "";
		public string _dCon = "";
		public string _originalTerm = "";
		public string _primary = "";
		public string _current = "";
		public string _lastSearch = "";
		public MedDRATerm _term = null;
		private MedDRAPreference _pref = null;
		//server object
		private CCDataServer _ds = null;
		//statusbar panels
		private enum statusPanel
		{
			status, results, database
		}
		//is the form loaded
		private bool _loaded = false;

		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.StatusBar statusBar1;
		private System.Windows.Forms.Button btnSearch;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.TextBox txtSearch;
		private System.Windows.Forms.ListView lvwTerms;
		private System.Windows.Forms.MainMenu mainMenu1;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem2;
		private System.Windows.Forms.ContextMenu contextMenu1;
		private System.Windows.Forms.MenuItem menuItem3;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.PictureBox pictureBox2;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.PictureBox pictureBox3;
		private System.Windows.Forms.Label label3;
		private System.ComponentModel.IContainer components;

		public MedDRABrowser(string custom, string responseValue)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			_custom = custom;
			txtSearch.Text = responseValue;
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			_ds.Dispose();

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MedDRABrowser));
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnSearch = new System.Windows.Forms.Button();
			this.txtSearch = new System.Windows.Forms.TextBox();
			this.statusBar1 = new System.Windows.Forms.StatusBar();
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnOK = new System.Windows.Forms.Button();
			this.lvwTerms = new System.Windows.Forms.ListView();
			this.contextMenu1 = new System.Windows.Forms.ContextMenu();
			this.menuItem3 = new System.Windows.Forms.MenuItem();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.mainMenu1 = new System.Windows.Forms.MainMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.menuItem2 = new System.Windows.Forms.MenuItem();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label3 = new System.Windows.Forms.Label();
			this.pictureBox3 = new System.Windows.Forms.PictureBox();
			this.label2 = new System.Windows.Forms.Label();
			this.pictureBox2 = new System.Windows.Forms.PictureBox();
			this.label1 = new System.Windows.Forms.Label();
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.btnSearch);
			this.groupBox1.Controls.Add(this.txtSearch);
			this.groupBox1.Location = new System.Drawing.Point(4, 4);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(908, 52);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Search";
			// 
			// btnSearch
			// 
			this.btnSearch.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.btnSearch.Location = new System.Drawing.Point(828, 20);
			this.btnSearch.Name = "btnSearch";
			this.btnSearch.Size = new System.Drawing.Size(72, 24);
			this.btnSearch.TabIndex = 1;
			this.btnSearch.Text = "Search";
			this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
			// 
			// txtSearch
			// 
			this.txtSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtSearch.Location = new System.Drawing.Point(4, 20);
			this.txtSearch.MaxLength = 255;
			this.txtSearch.Name = "txtSearch";
			this.txtSearch.Size = new System.Drawing.Size(816, 20);
			this.txtSearch.TabIndex = 0;
			this.txtSearch.Text = "txtSearch";
			// 
			// statusBar1
			// 
			this.statusBar1.Location = new System.Drawing.Point(0, 546);
			this.statusBar1.Name = "statusBar1";
			this.statusBar1.ShowPanels = true;
			this.statusBar1.Size = new System.Drawing.Size(916, 20);
			this.statusBar1.TabIndex = 2;
			this.statusBar1.Text = "statusBar1";
			// 
			// btnCancel
			// 
			this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancel.Location = new System.Drawing.Point(836, 512);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(72, 24);
			this.btnCancel.TabIndex = 3;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// btnOK
			// 
			this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnOK.Enabled = false;
			this.btnOK.Location = new System.Drawing.Point(752, 512);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(72, 24);
			this.btnOK.TabIndex = 4;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// lvwTerms
			// 
			this.lvwTerms.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lvwTerms.ContextMenu = this.contextMenu1;
			this.lvwTerms.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.lvwTerms.Location = new System.Drawing.Point(4, 56);
			this.lvwTerms.Name = "lvwTerms";
			this.lvwTerms.Size = new System.Drawing.Size(908, 444);
			this.lvwTerms.SmallImageList = this.imageList1;
			this.lvwTerms.TabIndex = 5;
			this.lvwTerms.View = System.Windows.Forms.View.SmallIcon;
			this.lvwTerms.DoubleClick += new System.EventHandler(this.lvwTerms_DoubleClick);
			this.lvwTerms.SelectedIndexChanged += new System.EventHandler(this.lvwTerms_SelectedIndexChanged);
			// 
			// contextMenu1
			// 
			this.contextMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								 this.menuItem3});
			// 
			// menuItem3
			// 
			this.menuItem3.Index = 0;
			this.menuItem3.Text = "View MedDRA Tree...";
			this.menuItem3.Click += new System.EventHandler(this.menuItem3_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(12, 12);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// mainMenu1
			// 
			this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																							this.menuItem1});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																							this.menuItem2});
			this.menuItem1.Text = "Settings";
			// 
			// menuItem2
			// 
			this.menuItem2.Index = 0;
			this.menuItem2.Text = "Preferences...";
			this.menuItem2.Click += new System.EventHandler(this.menuItem2_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.pictureBox3);
			this.groupBox2.Controls.Add(this.label2);
			this.groupBox2.Controls.Add(this.pictureBox2);
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Controls.Add(this.pictureBox1);
			this.groupBox2.Location = new System.Drawing.Point(8, 504);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(732, 36);
			this.groupBox2.TabIndex = 6;
			this.groupBox2.TabStop = false;
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(236, 16);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(128, 12);
			this.label3.TabIndex = 5;
			this.label3.Text = "Exact && Historical Match";
			// 
			// pictureBox3
			// 
			this.pictureBox3.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox3.Image")));
			this.pictureBox3.Location = new System.Drawing.Point(216, 16);
			this.pictureBox3.Name = "pictureBox3";
			this.pictureBox3.Size = new System.Drawing.Size(16, 12);
			this.pictureBox3.TabIndex = 4;
			this.pictureBox3.TabStop = false;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(120, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(132, 12);
			this.label2.TabIndex = 3;
			this.label2.Text = "Historical Match";
			// 
			// pictureBox2
			// 
			this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
			this.pictureBox2.Location = new System.Drawing.Point(100, 16);
			this.pictureBox2.Name = "pictureBox2";
			this.pictureBox2.Size = new System.Drawing.Size(16, 12);
			this.pictureBox2.TabIndex = 2;
			this.pictureBox2.TabStop = false;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(24, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(68, 12);
			this.label1.TabIndex = 1;
			this.label1.Text = "Exact Match";
			// 
			// pictureBox1
			// 
			this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
			this.pictureBox1.Location = new System.Drawing.Point(8, 16);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(12, 12);
			this.pictureBox1.TabIndex = 0;
			this.pictureBox1.TabStop = false;
			// 
			// MedDRABrowser
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(916, 566);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.lvwTerms);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.statusBar1);
			this.Controls.Add(this.groupBox1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Menu = this.mainMenu1;
			this.MinimizeBox = false;
			this.Name = "MedDRABrowser";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "InferMed MACRO MedDRA Browser";
			this.Activated += new System.EventHandler(this.MedDRABrowser_Activated);
			this.groupBox1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Initialise the form
		/// </summary>
		/// <returns></returns>
		public bool Init()
		{
			string dictionary, ip, cCon, dCon, dPrefix, preferencesFile;

			CCXml.MedDRAUnwrapXmlCustom(_custom, out dictionary, out ip, out cCon, out dCon, out dPrefix, out preferencesFile);

			_ds = new CCDataServer( CCDataServer.DicType.MEDDRA, dictionary, ip, cCon, dCon, dPrefix );
			if( _ds.Status == CCDataServer.ServerStatus.Ready )
			{
				_dictionary = dictionary;
				_dCon = dCon;
				_pref = new MedDRAPreference( preferencesFile );
				InitListView();
				ShowPreferences();
				InitStatusBar();
				btnOK.Enabled = false;
				contextMenu1.MenuItems[0].Enabled = false;
				log.Debug( "DICTIONARY=" + _dictionary );
				return( true );
			}
			else
			{
				log.Error( _ds.Error );
				MessageBox.Show( "The plugin encountered a problem while initialising : " + _ds.Error, "MedDRA Plugin" );
				return( false );
			}
		}

		/// <summary>
		/// Cancel
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		/// <summary>
		/// OK
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnOK_Click(object sender, System.EventArgs e)
		{
			Accept();
		}

		private void Accept()
		{
			if( lvwTerms.SelectedItems.Count > 0 )
			{
				_accepted = true;
			}
			
			this.Close();
		}

		/// <summary>
		/// Initialise the statusbar
		/// </summary>
		private void InitStatusBar()
		{
			statusBar1.ShowPanels = true;

			statusBar1.Panels.Add( "Ready" );
			statusBar1.Panels[0].Width = 150;

			statusBar1.Panels.Add( "" );
			statusBar1.Panels[1].Width = 400;

			statusBar1.Panels.Add( _dictionary );
			statusBar1.Panels[2].Width = 200;
		}

		/// <summary>
		/// Initialise the listview
		/// </summary>
		private void InitListView()
		{
			lvwTerms.View = View.Details;
			lvwTerms.GridLines = true;
			lvwTerms.MultiSelect = false;
			lvwTerms.FullRowSelect = true;

			lvwTerms.Columns.Add("", 15, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("Low Level Term", 200, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("LLT Key", 60, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("Weight", 40, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("Full Match", 30, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("Part Match", 30, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("Primary", 30, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("Current", 30, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("Preferred Term", 140, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("PT Key", 60, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("High Level Term", 140, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("HLT Key", 60, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("High Level Group Term", 140, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("HLGT Key", 60, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("System Organ Class", 140, HorizontalAlignment.Left);
			lvwTerms.Columns.Add("SOC Key", 60, HorizontalAlignment.Left);
		}

		/// <summary>
		/// Show/hide depending on preferences
		/// </summary>
		private void ShowPreferences()
		{
			lvwTerms.Columns[2].Width = ( _pref._lltKey ) ? 60 : 0;
			lvwTerms.Columns[3].Width = ( _pref._weight ) ? 40 : 0;
			lvwTerms.Columns[4].Width = ( _pref._fullMatch ) ? 30 : 0;
			lvwTerms.Columns[5].Width = ( _pref._partMatch ) ? 30 : 0;
			lvwTerms.Columns[6].Width = ( _pref._primary ) ? 30 : 0;
			lvwTerms.Columns[7].Width = ( _pref._current ) ? 30 : 0;
			lvwTerms.Columns[8].Width = ( _pref._pt ) ? 140 : 0;
			lvwTerms.Columns[9].Width = ( _pref._ptKey ) ? 60 : 0;
			lvwTerms.Columns[10].Width = ( _pref._hlt ) ? 140 : 0;
			lvwTerms.Columns[11].Width = ( _pref._hltKey ) ? 60 : 0;
			lvwTerms.Columns[12].Width = ( _pref._hlgt ) ? 140 : 0;
			lvwTerms.Columns[13].Width = ( _pref._hlgtKey ) ? 60 : 0;
			lvwTerms.Columns[14].Width = ( _pref._soc ) ? 140 : 0;
			lvwTerms.Columns[15].Width = ( _pref._socKey ) ? 60 : 0;

			groupBox2.Visible = ( _pref._legend );
		}

		/// <summary>
		/// Load results into listview
		/// </summary>
		/// <param name="results"></param>
		/// <param name="historicalMatches"></param>
		/// <param name="totalMatches"></param>
		private void LoadListView(string[,] results, MedDRATerm[] historicalMatches, int totalMatches)
		{
			lvwTerms.Items.Clear();

			if ( results == null )
			{
				UpdateStatusBar( statusPanel.results, "0 rows returned of 0 matches" );
			}
			else
			{
				UpdateStatusBar( statusPanel.results, results.GetLength(0) + " rows returned of " + totalMatches + " matches" );

				log.Info( "LOADING LISTVIEW" );
				//loop through results array
				for(int n = 0; n < results.GetLength(0); n++)
				{
					//create listview item and meddraterm object
					ListViewItem lvi = new ListViewItem();
					MedDRATerm t = new MedDRATerm( results[n, (int)CCDataServer.ResultCol.socKey], results[n, (int)CCDataServer.ResultCol.soc],
						results[n, (int)CCDataServer.ResultCol.socAbbrev], results[n, (int)CCDataServer.ResultCol.hlgtKey],
						results[n, (int)CCDataServer.ResultCol.hlgt], results[n, (int)CCDataServer.ResultCol.hltKey],
						results[n, (int)CCDataServer.ResultCol.hlt], results[n, (int)CCDataServer.ResultCol.ptKey],
						results[n, (int)CCDataServer.ResultCol.pt], results[n, (int)CCDataServer.ResultCol.lltKey],
						results[n, (int)CCDataServer.ResultCol.llt] );
					//attach the meddraterm object to the listview item
					lvi.Tag = ( object )t;
					//set the bestmatch icon
					if( MedDRATerm.IsBestMatch( results[n, (int)CCDataServer.ResultCol.autoencoder] ) ) 
					{
						lvi.ImageIndex = _IMAGEBESTMATCH;
					}
					//set the historicalmatch icon
					if( IsHistoricalMatch( historicalMatches, results[n, (int)CCDataServer.ResultCol.lltKey],
						results[n, (int)CCDataServer.ResultCol.ptKey], results[n, (int)CCDataServer.ResultCol.hltKey],
						results[n, (int)CCDataServer.ResultCol.hlgtKey], results[n, (int)CCDataServer.ResultCol.socKey] ) )
					{
						if( lvi.ImageIndex == _IMAGEBESTMATCH )
						{
							lvi.ImageIndex = _IMAGEBESTHISTORICALMATCH;
						}
						else
						{
							lvi.ImageIndex = _IMAGEHISTORICALMATCH;
						}
					}
					//set the column values
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.llt] );
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.lltKey] );
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.weight] );
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.fullMatch] );
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.partMatch] );
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.primary] );
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.current] );
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.pt] );
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.ptKey] );
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.hlt] );
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.hltKey] );
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.hlgt] );
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.hlgtKey] );
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.soc] );
					lvi.SubItems.Add( results[n, (int)CCDataServer.ResultCol.socKey] );
					//add the listview item to the listview
					lvwTerms.Items.Add( lvi );
				}
			}
		}

		/// <summary>
		/// Is the row the bestmatch
		/// </summary>
		/// <param name="autoencoder"></param>
		/// <returns></returns>
		private bool IsBestMatch( string autoencoder )
		{
			if( autoencoder.Length != 2 ) return false;
			if( autoencoder.Substring( 0, 1 ) != "Y" ) return false;
			return true;
		}

		/// <summary>
		/// is the row a historicalmatch
		/// </summary>
		/// <param name="historicalMatches"></param>
		/// <param name="lltKey"></param>
		/// <param name="ptKey"></param>
		/// <param name="hltKey"></param>
		/// <param name="hlgtKey"></param>
		/// <param name="socKey"></param>
		/// <returns></returns>
		private bool IsHistoricalMatch( MedDRATerm[] historicalMatches, string lltKey, string ptKey, string hltKey,
			string hlgtKey, string socKey)
		{
			for( int n = 0; n < historicalMatches.Length; n++ )
			{
				if( ( historicalMatches[n]._lltKey == lltKey ) && ( historicalMatches[n]._ptKey == ptKey )
					&& ( historicalMatches[n]._hltKey == hltKey ) && ( historicalMatches[n]._hlgtKey == hlgtKey )
					&& ( historicalMatches[n]._socKey == socKey ) )
				{
					return( true );
				}
			}
			return( false );
		}

		/// <summary>
		/// Search
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSearch_Click(object sender, System.EventArgs e)
		{
			Search();
		}

		/// <summary>
		/// Search the dictionary for the search text
		/// </summary>
		private void Search()
		{
			int totalMatches = 0;
			string[,] results;
			MedDRATerm[] historicalMatches = null;

			try
			{
				if( txtSearch.Text != MedDRATerm.Clean( txtSearch.Text ) )
				{
					MessageBox.Show("Search term may not contain the following characters: " + MedDRATerm._FORBIDDEN_CHARS, "MedDRA Plugin");
					txtSearch.Focus();
					return;
				}
				this.Enabled = false;
				btnOK.Enabled = false;
				this.Cursor = Cursors.WaitCursor;
				_originalTerm = txtSearch.Text;

				log.Info( "FINDING TERM" );
				UpdateStatusBar( statusPanel.status, "Searching dictionary..." );
				results = _ds.FindTerm( _originalTerm, CCDataServer.SearchType.llt, _pref._result, ref totalMatches );

				log.Info( "FINDING HISTORICAL MATCHES" );
				UpdateStatusBar( statusPanel.status, "Searching historical matches..." );
				historicalMatches = MedDRATerm.Matches( _dCon, _dictionary, _originalTerm );

				log.Info( "LOADING LISTVIEW" );
				UpdateStatusBar( statusPanel.status, "Loading results pane..." );
				LoadListView( results, historicalMatches, totalMatches );
			}
			catch( Exception ex )
			{
				log.Error( ex.Message + " : " + ex.InnerException );
				MessageBox.Show( "The plugin encountered a problem while searching : " + ex.Message + " : " + ex.InnerException, "MedDRA Plugin" );
			}
			finally
			{
				UpdateStatusBar( statusPanel.status, "Ready" );
				this.Cursor = Cursors.Default;
				this.Enabled = true;
			}
		}

		/// <summary>
		/// Set statusbar panel text
		/// </summary>
		/// <param name="p"></param>
		/// <param name="s"></param>
		private void UpdateStatusBar( statusPanel p, string s )
		{
			statusBar1.Panels[( int )p].Text = s;
			Application.DoEvents();
		}

		/// <summary>
		/// Set the currently selected meddraterm 
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void lvwTerms_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if( lvwTerms.SelectedItems.Count > 0 )
			{
				btnOK.Enabled = true;
				_term = ( MedDRATerm )lvwTerms.SelectedItems[0].Tag;
				contextMenu1.MenuItems[0].Enabled = true;
			}
			else
			{
				btnOK.Enabled = false;
				contextMenu1.MenuItems[0].Enabled = false;
				_term = null;
			}
		}

		/// <summary>
		/// Load the dictionary and run a search on initial activation
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void MedDRABrowser_Activated(object sender, System.EventArgs e)
		{
			bool dictLoaded = false;

			if( _loaded ) return;
			
			UpdateStatusBar( statusPanel.status, "Loading dictionary..." );
			this.Enabled = false;
			this.Cursor = Cursors.WaitCursor;
			dictLoaded = _ds.LoadDictionary();
			this.Cursor = Cursors.Default;
			this.Enabled = true;
			UpdateStatusBar( statusPanel.status, "Ready" );

			if( dictLoaded )
			{
				Search();
				_loaded = true;
			}
			else
			{
				MessageBox.Show( "Unable to load MedDRA dictionary", "MedDRA Plugin" );
				this .Close();
			}

			
		}

		/// <summary>
		/// Preferences
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void menuItem2_Click(object sender, System.EventArgs e)
		{
			try
			{
				MedDRAPreferences mp = new MedDRAPreferences( ref _pref );
				mp.ShowDialog();
				_pref.Save();
				mp.Dispose();
				ShowPreferences();
			}
			catch( Exception ex )
			{
				MessageBox.Show( "The plugin encountered a problem while saving preferences : " + ex.Message +  " : " 
					+ ex.InnerException, "MedDRA Plugin" );
			}
		}

		/// <summary>
		/// Tree
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void menuItem3_Click(object sender, System.EventArgs e)
		{
			if( _term != null )
			{
				MedDRATree mt = new MedDRATree( _term );
				mt.ShowDialog();
				mt.Dispose();
			}
		}

		/// <summary>
		/// Double click a row
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void lvwTerms_DoubleClick(object sender, EventArgs e)
		{
			Accept();
		}
	}
}
