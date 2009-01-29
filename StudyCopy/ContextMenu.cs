using System;
using System.Windows.Forms;
using System.Drawing;

namespace InferMed.MACRO.StudyMerge
{
	/// <summary>
	/// Summary description for ContextMenu.
	/// </summary>
	public class ContextMenu : MenuItem
	{
		private Image itemImage;

		public ContextMenu(string text, System.EventHandler onClick) : base()
		{
			this.OwnerDraw = true;
			this.Text = text;
		}

		public Image MenuImage
		{
			get { return( itemImage ); }
			set { itemImage = value; }
		}
		
		protected override void OnDrawItem(DrawItemEventArgs e)
		{
			e.DrawBackground();
			Pen menuPen = new Pen(Color.DarkRed);
			//Rectangle rect = new Rectangle( e.Bounds.X, e.Bounds.Y, e.Bounds.Width+500, e.Bounds.Height+500 );
			Rectangle rect = new Rectangle( 1000, 1000, 5000, 5000 );
			Graphics g = e.Graphics;
			g.DrawRectangle( menuPen, rect );
			//e.Graphics.DrawString( this.Text );
			//g.DrawLine(menuPen, 1000, 1000, 2000, 2000);
			e.DrawFocusRectangle();

			base.OnDrawItem (e);
		}

	}
}
