using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using InferMed.Components;
using System.Data;

namespace InferMed.MACRO.StudyCopy
{
	/// <summary>
	/// Display dataitems for user selection
	/// </summary>
	public class DataItemsForm : System.Windows.Forms.Form
	{
		//database connection object
		public static IDbConnection _dbConn = null;

		//dataitem identifiers
		private string _studyId = "";
		private string _dataType = "";
		private string _dataItemCode = "";

		//loaded flag
		private bool _loaded = false;

		//selection
		private string _selectedDataItemId = "0";

		//column constants
		private const short _DATAITEMID_COL = 1;
		private const short _DATAITEMCODE_COL = 2;
		private const short _DATAITEMNAME_COL = 3;
		private const short _DATAITEMFORMAT_COL = 4;
		private const short _DATAITEMLENGTH_COL = 5;
		private const short _DERIVATION_COL = 6;

		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.ListView lvwDataItems;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DataItemsForm( IDbConnection dbConn, string studyId, string dataType, string dataItemCode )
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			_dbConn = dbConn;
			_studyId = studyId;
			_dataType = dataType;
			_dataItemCode = dataItemCode;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(DataItemsForm));
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnOK = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.lvwDataItems = new System.Windows.Forms.ListView();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.btnOK);
			this.groupBox1.Controls.Add(this.btnCancel);
			this.groupBox1.Controls.Add(this.lvwDataItems);
			this.groupBox1.Location = new System.Drawing.Point(4, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(526, 318);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// btnOK
			// 
			this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnOK.Enabled = false;
			this.btnOK.Location = new System.Drawing.Point(358, 286);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(72, 24);
			this.btnOK.TabIndex = 2;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// btnCancel
			// 
			this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancel.Location = new System.Drawing.Point(442, 286);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(72, 24);
			this.btnCancel.TabIndex = 1;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// lvwDataItems
			// 
			this.lvwDataItems.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lvwDataItems.HideSelection = false;
			this.lvwDataItems.Location = new System.Drawing.Point(4, 12);
			this.lvwDataItems.Name = "lvwDataItems";
			this.lvwDataItems.Size = new System.Drawing.Size(518, 266);
			this.lvwDataItems.TabIndex = 0;
			this.lvwDataItems.SelectedIndexChanged += new System.EventHandler(this.lvwDataItems_SelectedIndexChanged);
			// 
			// DataItemsForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(534, 324);
			this.Controls.Add(this.groupBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "DataItemsForm";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Select DataItem";
			this.Activated += new System.EventHandler(this.DataItemsForm_Activated);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Initialise form
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void DataItemsForm_Activated(object sender, System.EventArgs e)
		{
			if( _loaded ) return;
			_loaded = true;

			try
			{
				//create listview headers
				InitListView();
				//load data
				LoadListView();
			}
			
#if( !DEBUG )
			catch( Exception ex )
			{
					MessageBox.Show( "An exception occurred while initialising : " + ex.Message + " : " + ex.InnerException, 
						"MACRO Study Merge", MessageBoxButtons.OK, MessageBoxIcon.Error );
			}
#endif
			finally
			{
			}
		}

		/// <summary>
		/// User selection
		/// </summary>
		public string SelectedDataItemId
		{
			get{ return( _selectedDataItemId ); }
		}

		/// <summary>
		/// Load dataitems into listview
		/// </summary>
		private void LoadListView()
		{
			DataSet ds = null;

			try
			{
				//get appropriate dataitems
				ds = MACRO30.GetDataItems( _dbConn, _studyId, _dataType );

				//load listview
				foreach( DataRow row in ds.Tables[0].Rows )
				{
					ListViewItem lvi = new ListViewItem();
					lvi.Text = "";
					lvi.SubItems.Add( row["DATAITEMID"].ToString() );
					lvi.SubItems.Add( row["DATAITEMCODE"].ToString() );
					lvi.SubItems.Add( row["DATAITEMNAME"].ToString() );
					lvi.SubItems.Add( row["DATAITEMFORMAT"].ToString() );
					lvi.SubItems.Add( row["DATAITEMLENGTH"].ToString() );
					lvi.SubItems.Add( row["DERIVATION"].ToString() );
					lvwDataItems.Items.Add( lvi );

					//select a dataitem if one set
					if( row["DATAITEMCODE"].ToString() == _dataItemCode )
					{
						lvi.Selected = true;
					}
				}

				//enable buttons
				if( lvwDataItems.Items.Count > 0 )
				{
					btnOK.Enabled = true;
					if( lvwDataItems.SelectedItems.Count > 0 ) lvwDataItems.SelectedItems[0].EnsureVisible();
				}
			}
			finally
			{
				ds.Dispose();
			}
		}

		/// <summary>
		/// Initialise listview control
		/// </summary>
		private void InitListView()
		{
			//eform elements listview
			lvwDataItems.View = View.Details;
			lvwDataItems.CheckBoxes = false;
			lvwDataItems.GridLines = true;
			lvwDataItems.MultiSelect = false;
			lvwDataItems.FullRowSelect = true;
			lvwDataItems.Columns.Add("", 0, HorizontalAlignment.Left);
			lvwDataItems.Columns.Add("DataItem Id", 50, HorizontalAlignment.Left);
			lvwDataItems.Columns.Add("DataItem Code", 60, HorizontalAlignment.Left);
			lvwDataItems.Columns.Add("DataItem Name", 110, HorizontalAlignment.Left);
			lvwDataItems.Columns.Add("DataItem Format", 60, HorizontalAlignment.Left);
			lvwDataItems.Columns.Add("DataItem Length", 50, HorizontalAlignment.Left);
			lvwDataItems.Columns.Add("Derivation", 150, HorizontalAlignment.Left);
		}

		private void lvwDataItems_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if( lvwDataItems.SelectedItems.Count > 0 )
			{
				_selectedDataItemId = lvwDataItems.SelectedItems[0].SubItems[_DATAITEMID_COL].Text;
			}
			else
			{
				_selectedDataItemId = "0";
			}
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			_selectedDataItemId = "0";
			this.Close();
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			if( _selectedDataItemId == "0" )
			{
				MessageBox.Show( "Please select a dataitem" );
				return;
			}
			this.Close();
		}
	}
}
