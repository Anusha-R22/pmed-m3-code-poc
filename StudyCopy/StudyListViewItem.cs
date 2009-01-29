using System;
using System.Data;
using System.Windows.Forms;


//----------------------------------------------------------------------
// 28/04/2006 bug 2726 incorrect matching and display of rqg elements
//
//----------------------------------------------------------------------

namespace InferMed.MACRO.StudyCopy
{
	/// <summary>
	/// Study copy listviewitem object
	/// </summary>
	public class StudyListViewItem : ListViewItem
	{
		private StudyCopyGlobal.ElementType _ElementType;
		private DataRow _SourceStudyElementRow = null;
		private DataRow _DestinationStudyElementRow = null;

		//image constants
		public const short _COPIED= 0;
		public const short _COPIEDADD = 1;
		public const short _PARTCOPIED = 2;
		public const short _PARTCOPIEDADD = 3;
		public const short _ADD = 4;

		//column constants
		private const short _DEFORMID_COL = 1;
		private const short _DEFORMTITLE_COL = 2;
		private const short _DEFORMLABEL_COL = 3;
		private const short _SEFORMID_COL = 4;
		private const short _SEFORMTITLE_COL = 5;
		private const short _SEFORMLABEL_COL = 6;

		private const short _DELEMENTID_COL = 1;
		private const short _DELEMENTCAPTION_COL = 2;
		private const short _DELEMENTCONTROLTYPE_COL = 3;
		private const short _DDATAITEMCODE_COL = 4;
		private const short _DDATAITEMNAME_COL = 5;
		private const short _DDATATYPE_COL = 6;
		private const short _DXY_COL = 7;
		private const short _DGROUPID_COL = 8;
		private const short _SELEMENTID_COL = 9;
		private const short _SELEMENTCAPTION_COL = 10;
		private const short _SELEMENTCONTROLTYPE_COL = 11;
		private const short _SDATAITEMCODE_COL = 12;
		private const short _SDATAITEMNAME_COL = 13;
		private const short _SDATATYPE_COL = 14;
		private const short _SXY_COL = 15;
		private const short _SGROUPID_COL = 16;

		//item display icon
		public enum ItemIcon
		{
			None, Copied, CopiedAdd, PartCopied, PartCopiedAdd, Add
		}

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="eType"></param>
		/// <param name="icon"></param>
		public StudyListViewItem( StudyCopyGlobal.ElementType eType, ItemIcon icon )
		{
			this.Text = "";
			_ElementType = eType;

			switch( eType )
			{
				case StudyCopyGlobal.ElementType.eForm:
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					break;

				case StudyCopyGlobal.ElementType.eFormElement:
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					break;
			}
			SetIcon( icon );
		}

		/// <summary>
		/// Constructor
		/// </summary>
		/// 28/04/2006 bug 2726 incorrect matching and display of rqg elements
		public StudyListViewItem( StudyCopyGlobal.ElementType eType, DataRow row )
		{
			this.Text = "";
			_ElementType = eType;
			_DestinationStudyElementRow = row;

			switch( eType )
			{
				case StudyCopyGlobal.ElementType.eForm:
					this.SubItems.Add( row["CRFPAGEID"].ToString() );
					this.SubItems.Add( row["CRFTITLE"].ToString() );
					this.SubItems.Add( row["CRFPAGELABEL"].ToString() );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					break;
				case StudyCopyGlobal.ElementType.eFormElement:
					this.SubItems.Add( row["CRFELEMENTID"].ToString() );
					this.SubItems.Add( row["CAPTION"].ToString() );
					this.SubItems.Add( StudyCopyGlobal.GetControlType( row["CONTROLTYPE"].ToString() ) );
					this.SubItems.Add( ( row["DATAITEMCODE"].ToString() == "0" ) ? "" : row["DATAITEMCODE"].ToString() );
					this.SubItems.Add( row["DATAITEMNAME"].ToString() );
					this.SubItems.Add( StudyCopyGlobal.GetDataType( row["DATATYPE"].ToString() ) );
					this.SubItems.Add( ( ( row["X"].ToString() == "0" ) && ( row["Y"].ToString() == "0" ) ) ? "" : row["X"].ToString() + "x" + row["Y"].ToString() + "y" );
					this.SubItems.Add( ( row["QGROUPID"].ToString() == "0" ) ? "" : row["QGROUPCODE"].ToString() );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					this.SubItems.Add( "" );
					break;
			}
		}

		/// <summary>
		/// Sets/gets destination element row
		/// </summary>
		/// 28/04/2006 bug 2726 incorrect matching and display of rqg elements
		public DataRow DestinationStudyElementRow
		{
			get { return  _DestinationStudyElementRow; }
			set 
			{
				if( value == null )
				{
					_DestinationStudyElementRow = null;

					switch( _ElementType )
					{
						case StudyCopyGlobal.ElementType.eForm:
							this.SubItems[_DEFORMID_COL].Text = "";
							this.SubItems[_DEFORMTITLE_COL].Text = "";
							this.SubItems[_DEFORMLABEL_COL].Text = "";
							break;
						case StudyCopyGlobal.ElementType.eFormElement:
							this.SubItems[_DELEMENTID_COL].Text = "";
							this.SubItems[_DELEMENTCAPTION_COL].Text = "";
							this.SubItems[_DELEMENTCONTROLTYPE_COL].Text = "";
							this.SubItems[_DDATAITEMCODE_COL].Text = "";
							this.SubItems[_DDATAITEMNAME_COL].Text = "";
							this.SubItems[_DDATATYPE_COL].Text = "";
							this.SubItems[_DXY_COL].Text = "";
							this.SubItems[_DGROUPID_COL].Text = "";
							break;
					}
					SetIcon( ItemIcon.None );
				}
				else
				{
					_DestinationStudyElementRow = value;

					switch( _ElementType )
					{
						case StudyCopyGlobal.ElementType.eForm:
							this.SubItems[_DEFORMID_COL].Text = value["CRFPAGEID"].ToString();
							this.SubItems[_DEFORMTITLE_COL].Text = value["CRFTITLE"].ToString();
							this.SubItems[_DEFORMLABEL_COL].Text = value["CRFPAGELABEL"].ToString();
							break;
						case StudyCopyGlobal.ElementType.eFormElement:
							this.SubItems[_DELEMENTID_COL].Text = value["CRFELEMENTID"].ToString();
							this.SubItems[_DELEMENTCAPTION_COL].Text = value["CAPTION"].ToString();
							this.SubItems[_DELEMENTCONTROLTYPE_COL].Text = StudyCopyGlobal.GetControlType( value["CONTROLTYPE"].ToString() );
							this.SubItems[_DDATAITEMCODE_COL].Text = ( value["DATAITEMCODE"].ToString() == "0" ) ? "" : value["DATAITEMCODE"].ToString();
							this.SubItems[_DDATAITEMNAME_COL].Text = value["DATAITEMNAME"].ToString();
							this.SubItems[_DDATATYPE_COL].Text = StudyCopyGlobal.GetDataType( value["DATATYPE"].ToString() );
							this.SubItems[_DXY_COL].Text = ( ( value["X"].ToString() == "0" ) && ( value["Y"].ToString() == "0" ) ) ? "" : value["X"].ToString() + "x" + value["Y"].ToString() + "y" ;
							this.SubItems[_DGROUPID_COL].Text = ( value["QGROUPID"].ToString() == "0" ) ? "" : value["QGROUPCODE"].ToString();
							break;
					}
				}
			}
		}
		
		/// <summary>
		/// Sets/gets source element row
		/// </summary>
		/// 28/04/2006 bug 2726 incorrect matching and display of rqg elements
		public DataRow SourceStudyElementRow
		{
			get { return  _SourceStudyElementRow; }
			set
			{
				if( value == null )
				{
					_SourceStudyElementRow = null;

					switch( _ElementType )
					{
						case StudyCopyGlobal.ElementType.eForm:
							this.SubItems[_SEFORMID_COL].Text = "";
							this.SubItems[_SEFORMTITLE_COL].Text = "";
							this.SubItems[_SEFORMLABEL_COL].Text = "";
							break;
						case StudyCopyGlobal.ElementType.eFormElement:
							this.SubItems[_SELEMENTID_COL].Text = "";
							this.SubItems[_SELEMENTCAPTION_COL].Text = "";
							this.SubItems[_SELEMENTCONTROLTYPE_COL].Text = "";
							this.SubItems[_SDATAITEMCODE_COL].Text = "";
							this.SubItems[_SDATAITEMNAME_COL].Text = "";
							this.SubItems[_SDATATYPE_COL].Text = "";
							this.SubItems[_SXY_COL].Text = "";
							this.SubItems[_SGROUPID_COL].Text = "";
							break;
					}
					SetIcon( ItemIcon.None );
				}
				else
				{
					_SourceStudyElementRow = value;

					switch( _ElementType )
					{
						case StudyCopyGlobal.ElementType.eForm:
							this.SubItems[_SEFORMID_COL].Text = value["CRFPAGEID"].ToString();
							this.SubItems[_SEFORMTITLE_COL].Text = value["CRFTITLE"].ToString();
							this.SubItems[_SEFORMLABEL_COL].Text = value["CRFPAGELABEL"].ToString();
							break;
						case StudyCopyGlobal.ElementType.eFormElement:
							this.SubItems[_SELEMENTID_COL].Text = value["CRFELEMENTID"].ToString();
							this.SubItems[_SELEMENTCAPTION_COL].Text = value["CAPTION"].ToString();
							this.SubItems[_SELEMENTCONTROLTYPE_COL].Text = StudyCopyGlobal.GetControlType( value["CONTROLTYPE"].ToString() );
							this.SubItems[_SDATAITEMCODE_COL].Text = ( value["DATAITEMCODE"].ToString() == "0" ) ? "" : value["DATAITEMCODE"].ToString();
							this.SubItems[_SDATAITEMNAME_COL].Text = value["DATAITEMNAME"].ToString();
							this.SubItems[_SDATATYPE_COL].Text = StudyCopyGlobal.GetDataType( value["DATATYPE"].ToString() );
							this.SubItems[_SXY_COL].Text = ( ( value["X"].ToString() == "0" ) && ( value["Y"].ToString() == "0" ) ) ? "" : value["X"].ToString() + "x" + value["Y"].ToString() + "y";
							this.SubItems[_SGROUPID_COL].Text = ( value["QGROUPID"].ToString() == "0" ) ? "" : value["QGROUPCODE"].ToString();
							break;
					}
				}
			}
		}

		/// <summary>
		/// Gets source study element id
		/// </summary>
		public string SourceStudyElementId
		{
			get
			{ 
				string Id = "";

				switch( _ElementType )
				{
					case StudyCopyGlobal.ElementType.eForm:
						Id = ( _SourceStudyElementRow == null ) ? "" : _SourceStudyElementRow["CRFPAGEID"].ToString();
						break;
					case StudyCopyGlobal.ElementType.eFormElement:
						Id = ( _SourceStudyElementRow == null ) ? "" : _SourceStudyElementRow["CRFELEMENTID"].ToString();
						break;
				}
				return( Id );
			}
		}

		/// <summary>
		/// Gets destination study element id
		/// </summary>
		public string DestinationStudyElementId
		{
			get
			{
				string Id = "";

				switch( _ElementType )
				{
					case StudyCopyGlobal.ElementType.eForm:
						Id = ( _DestinationStudyElementRow == null ) ? "" : _DestinationStudyElementRow["CRFPAGEID"].ToString();
						break;
					case StudyCopyGlobal.ElementType.eFormElement:
						Id = ( _DestinationStudyElementRow == null ) ? "" : _DestinationStudyElementRow["CRFELEMENTID"].ToString();
						break;
				}
				return( Id );
			}
		}

		/// <summary>
		/// Sets item icon
		/// </summary>
		/// <param name="icon"></param>
		public void SetIcon( ItemIcon icon )
		{
			switch( icon )
			{
				case ItemIcon.None:
					this.ImageIndex = -1;
					break;
				case ItemIcon.Copied:
					this.ImageIndex = _COPIED;
					break;
				case ItemIcon.CopiedAdd:
					this.ImageIndex = _COPIEDADD;
					break;
				case ItemIcon.PartCopied:
					this.ImageIndex = _PARTCOPIED;
					break;
				case ItemIcon.PartCopiedAdd:
					this.ImageIndex = _PARTCOPIEDADD;
					break;
				case ItemIcon.Add:
					this.ImageIndex = _ADD;
					break;
			}
		}

		/// <summary>
		/// Is the item matched
		/// </summary>
		public bool MatchedElement
		{
			get{ return( ( _DestinationStudyElementRow != null ) && ( _SourceStudyElementRow != null ) ); }
		}

		/// <summary>
		/// Is the item copied
		/// </summary>
		public bool CopiedElement
		{
			get{ return( ( this.ImageIndex == _COPIED ) || ( this.ImageIndex == _PARTCOPIED ) ); }
		}
	}
}
