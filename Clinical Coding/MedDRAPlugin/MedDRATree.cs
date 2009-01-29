using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace InferMed.MACRO.ClinicalCoding.Plugins
{
	/// <summary>
	/// Summary description for MedDRATree.
	/// </summary>
	public class MedDRATree : System.Windows.Forms.Form
	{
		private System.Windows.Forms.TreeView treeView1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="term"></param>
		public MedDRATree( MedDRATerm term )
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			LoadTree( term );
		}

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="codedValue"></param>
		public MedDRATree( string codedValue )
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			LoadTree( codedValue );
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MedDRATree));
			this.treeView1 = new System.Windows.Forms.TreeView();
			this.SuspendLayout();
			// 
			// treeView1
			// 
			this.treeView1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.treeView1.ImageIndex = -1;
			this.treeView1.Location = new System.Drawing.Point(0, 0);
			this.treeView1.Name = "treeView1";
			this.treeView1.SelectedImageIndex = -1;
			this.treeView1.Size = new System.Drawing.Size(296, 218);
			this.treeView1.TabIndex = 0;
			// 
			// MedDRATree
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(296, 218);
			this.Controls.Add(this.treeView1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "MedDRATree";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Tree";
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Load a MedDRA path
		/// </summary>
		/// <param name="term"></param>
		private void LoadTree( MedDRATerm term )
		{
			LoadTreeView( term._soc, term._socKey, term._hlgt, term._hlgtKey, term._hlt, term._hltKey, term._pt,
				term._ptKey, term._llt, term._lltKey );
		}

		/// <summary>
		/// Load a MedDRA path
		/// </summary>
		/// <param name="codedValue"></param>
		private void LoadTree( string codedValue )
		{
			string dic, soc, socKey, hlgt, hlgtKey, hlt, hltKey, pt, ptKey, llt, lltKey;

			CCXml.MedDRAUnwrapXmlNames( codedValue, out dic, out soc, out hlgt, out hlt, out pt, out llt );
			CCXml.MedDRAUnwrapXmlKeys( codedValue,  out dic, out socKey, out hlgtKey, out hltKey, out ptKey, out lltKey );

			LoadTreeView( soc, socKey, hlgt, hlgtKey, hlt, hltKey, pt, ptKey, llt, lltKey );
		}

		/// <summary>
		/// Load a MedDRA path into a treeview control
		/// </summary>
		/// <param name="soc"></param>
		/// <param name="socKey"></param>
		/// <param name="hlgt"></param>
		/// <param name="hlgtKey"></param>
		/// <param name="hlt"></param>
		/// <param name="hltKey"></param>
		/// <param name="pt"></param>
		/// <param name="ptKey"></param>
		/// <param name="llt"></param>
		/// <param name="lltKey"></param>
		private void LoadTreeView( string soc, string socKey, string hlgt, string hlgtKey, string hlt, string hltKey, string pt,
			string ptKey, string llt, string lltKey )
		{
			treeView1.Nodes.Clear();
			treeView1.Nodes.Add( soc + " [" + socKey + "]" );
			treeView1.Nodes[0].Nodes.Add( hlgt + " [" + hlgtKey + "]" );
			treeView1.Nodes[0].Nodes[0].Nodes.Add( hlt + " [" + hltKey + "]" );
			treeView1.Nodes[0].Nodes[0].Nodes[0].Nodes.Add( pt + " [" + ptKey + "]" );
			treeView1.Nodes[0].Nodes[0].Nodes[0].Nodes[0].Nodes.Add( llt + " [" + lltKey + "]" );
			treeView1.ExpandAll();
		}
	}
}
