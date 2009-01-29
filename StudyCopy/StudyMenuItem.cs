using System;
using System.Windows.Forms;
using System.Data;


//----------------------------------------------------------------------
// 28/04/2006 bug 2726 incorrect matching and display of rqg elements
//
//----------------------------------------------------------------------

namespace InferMed.MACRO.StudyCopy
{
	/// <summary>
	/// Study copy menu item object
	/// </summary>
	public class StudyMenuItem : MenuItem
	{
		private DataRow _StudyElementRow = null;

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="eType"></param>
		/// <param name="text"></param>
		/// <param name="onClick"></param>
		public StudyMenuItem( StudyCopyGlobal.ElementType eType, string text, System.EventHandler onClick )
		{
			this.Text = text;
			if( onClick != null ) this.Click += onClick;
		}

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="eType"></param>
		/// <param name="row"></param>
		/// <param name="onClick"></param>
		public StudyMenuItem( StudyCopyGlobal.ElementType eType, DataRow row, System.EventHandler onClick )
		{
			string description = "", type = "", id = "";
			_StudyElementRow = row;
			this.Click += onClick;

			switch( eType )
			{
				case StudyCopyGlobal.ElementType.eForm:
					id = " " + row["CRFPAGEID"].ToString();
					description = " " + row["CRFTITLE"].ToString();
					break;

				case StudyCopyGlobal.ElementType.eFormElement:

					type = StudyCopyGlobal.GetControlType( row["CONTROLTYPE"].ToString() );
					switch( row["CONTROLTYPE"].ToString() )
					{
						case StudyCopyGlobal._QGROUP:
							id = " Id " + row["QGROUPCODE"].ToString();
							break;
						case StudyCopyGlobal._LINE:
							description = " " + row["X"].ToString() + "x" + row["Y"].ToString() + "y";
							break;
						default:
							id = " [" + row["CRFELEMENTID"].ToString() + 
								( ( row["DATAITEMCODE"].ToString() != "" ) ? "/" + row["DATAITEMCODE"].ToString() : "" ) + "]";
							description = " " + row["CAPTION"].ToString();
							break;
					}
					break;
			}

			description = ( description.Length > 50 ) ? description.Substring(0, 50 ) + "..." : description;
			this.Text = type + id + description;
		}

		/// <summary>
		/// Get study element row
		/// </summary>
		public DataRow StudyElementRow
		{
			get{ return( _StudyElementRow ); }
			set{ _StudyElementRow = value; }
		}
	}
}
