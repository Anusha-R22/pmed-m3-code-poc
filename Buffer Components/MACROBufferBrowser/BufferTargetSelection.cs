using System;
using System.Text;
using System.Collections;
using System.Data;
using InferMed.Components;
using log4net;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for BufferTargetSelection.
	/// </summary>
	class BufferTargetSelection
	{
		// private members
		// scope
		private int _scope;
		// type of action
		private int _action;
		// commit identifier
		private string _identStudy;
		private string _identSite;
		private string _identSubject;
		private string _identVisitId;
		private string _identVisitCycle;
		private string _identEformId;
		private string _identEformCycle;
		private string _identEformTaskId;
		private string _identEformElementId;
		private string _identResponseCycle;
		private string _identResponseTaskId;
		private string _identBufferResponseId;
		// back link detail
		private string _backlinkStudyId;
		private string _backlinkSite;
		private string _backlinkSubjectId;
		// Datatable to hold display matches
		private DataTable _dtPossibleMatches;
		// bookmark tracker
		private int _bookMark;
		// raw form members
		private string _rawIdentifier;
		private string _rawType;
		private string _rawBack;
		private string _rawScope;

		// log4net
		private static readonly ILog log = LogManager.GetLogger( typeof(BufferDataBrowserSave) );

		/// <summary>
		/// 
		/// </summary>
		/// <param name="formData"></param>
		public BufferTargetSelection(string formData, string bookMark)
		{
			log.Debug( "formData = " + formData );
			
			// initialize member variables
			_identStudy = "";
			_identSite = "";
			_identSubject = "";
			_identVisitId = "";
			_identVisitCycle = "";
			_identEformId = "";
			_identEformCycle = "";
			_identEformTaskId = "";
			_identEformElementId = "";
			_identResponseCycle = "";
			_identResponseTaskId = "";
			_identBufferResponseId = "";
			_backlinkStudyId = "";
			_backlinkSite = "";
			_backlinkSubjectId = "";
			// store raw identifier
			_rawIdentifier = "";
			_rawType = "";
			_rawBack = "";
			_rawScope = "";

			ExtractFormData( formData );

			// store bookmark
			_bookMark = 0;
			try
			{
				_bookMark = Convert.ToInt32( bookMark );
			}
			catch{}
		}

		/// <summary>
		/// Extract the passed in form data to the objects member variables
		/// </summary>
		/// <param name="formData"></param>
		private void ExtractFormData(string formData)
		{
			// extract form data to private members
			log.Info( "Starting ExtractFormData" );
			// delimiters as char[]
			char[] chMainDelim = { System.Convert.ToChar("&") };
			char[] chSubDelim = { System.Convert.ToChar("=") };
			char[] chDataPartDelim = { System.Convert.ToChar("`") };
			ArrayList alFormData = new ArrayList();
			alFormData.AddRange( formData.Split( chMainDelim ) );
			// loop through individual data
			foreach( string sData in alFormData )
			{
				log.Debug( "form data = " + sData );

				ArrayList alIndivData = new ArrayList();
				alIndivData.AddRange( sData.Split( chSubDelim ) );
				// handle as required - 1st part is data name
				switch( alIndivData[0].ToString() )
				{
					case "bidentifier":
					{
						// store row identifying data
						// item detail - study id`site`subject id`visit id`visit cycle`eform id`eform cycle`eform task id`eform element id`response cycle`response task id`buffer response id
						string sRowIdentifier = BufferBrowser.ReplaceHtmlCharacters( alIndivData[1].ToString() );
						// store raw
						_rawIdentifier = sRowIdentifier;

						ArrayList alIdentify = new ArrayList();
						alIdentify.AddRange( sRowIdentifier.Split( chDataPartDelim ) );
						for(int i=0; i < alIdentify.Count; i++)
						{
							switch( i )
							{
								case 0:
								{
									// study id
									_identStudy = alIdentify[ i ].ToString();
									break;
								}
								case 1:
								{
									// site
									_identSite = alIdentify[ i ].ToString();
									break;
								}
								case 2:
								{
									// subject id
									_identSubject = alIdentify[ i ].ToString();
									break;
								}
								case 3:
								{
									// visit id
									_identVisitId = alIdentify[ i ].ToString();
									break;
								}
								case 4:
								{
									// visit cycle
									_identVisitCycle = alIdentify[ i ].ToString();
									break;
								}
								case 5:
								{
									// eform id
									_identEformId = alIdentify[ i ].ToString();
									break;
								}
								case 6:
								{
									// eform cycle
									_identEformCycle = alIdentify[ i ].ToString();
									break;
								}
								case 7:
								{
									// eform task id
									_identEformTaskId = alIdentify[ i ].ToString();
									break;
								}
								case 8:
								{
									// eform element id
									_identEformElementId = alIdentify[ i ].ToString();
									break;
								}
								case 9:
								{
									// response cycle
									_identResponseCycle = alIdentify[ i ].ToString();
									break;
								}
								case 10:
								{
									// response task id
									_identResponseTaskId = alIdentify[ i ].ToString();
									break;
								}
								case 11:
								{
									// buffer response id
									_identBufferResponseId = alIdentify[ i ].ToString();
									break;
								}
							}
						}
						break;
					}
					case "btype":
					{
						// type of save
						string sSaveType = BufferBrowser.ReplaceHtmlCharacters( alIndivData[1].ToString() );
						// store raw
						_rawType = sSaveType;
						// don't use (!)
						if( sSaveType != "" )
						{
							_action = Convert.ToInt32( sSaveType );
						}
						break;
					}
					case "bback":
					{
						// return to buffer browser info
						string sReturnIdentifier = BufferBrowser.ReplaceHtmlCharacters( alIndivData[1].ToString() );
						// store raw
						_rawBack = sReturnIdentifier;

						ArrayList alBackInfo = new ArrayList();
						alBackInfo.AddRange( sReturnIdentifier.Split( chDataPartDelim ) );
						// should be 3 parts
						_backlinkStudyId = alBackInfo[0].ToString();
						_backlinkSite = alBackInfo[1].ToString();
						_backlinkSubjectId = alBackInfo[2].ToString();
						break;
					}
					case "bscope":
					{
						// scope of save - not used
						string sSaveScope = BufferBrowser.ReplaceHtmlCharacters( alIndivData[1].ToString() );
						// store raw
						_rawScope = sSaveScope;
						// don't use
						if ( sSaveScope != "" )
						{
							_scope = Convert.ToInt32( sSaveScope );
						}
						break;
					}
				}
			}
		}
	
		/// <summary>
		/// 
		/// </summary>
		/// <param name="bufferUser"></param>
		/// <returns></returns>
		public string RenderBufferTargetPage( BufferMACROUser bufferUser )
		{
			log.Info( "Starting RenderBufferTargetPage" );
			
			// calculate possible matches
			_dtPossibleMatches = GetPossibleMatches( bufferUser );

			// calculate
			DataTable dtBufferRow = GetBufferRowData( bufferUser, _identBufferResponseId );

			// render page
			StringBuilder sbPageHtml = new StringBuilder();

			// PageSize from MACRO setting
			string sUserSetting = BufferBrowser._SETTING_PAGE_LENGTH;
			object nDefaultPageLength = 50;
			int nPageSize = Convert.ToInt32( bufferUser.MACROUser.UserSettings.GetSetting(ref sUserSetting, ref nDefaultPageLength) );
			int nStart = 0;
			int nEnd = 0;

			// rowspan variables
			int nRowSpanSubject = 0;
			int nRowSpanVisit = 0;
			int nRowSpanEform = 0;

			// calculate starting row from bookmark
			if( ( _bookMark > _dtPossibleMatches.Rows.Count ) || ( _bookMark < 0 ) )
			{
				nStart = 0;
				_bookMark = 0;
			}
			else
			{
				nStart = _bookMark;
			}
			// calculate end row from given
			if( nStart + nPageSize >= _dtPossibleMatches.Rows.Count )
			{
				nEnd = _dtPossibleMatches.Rows.Count - 1;
			}
			else
			{
				nEnd = (nStart + nPageSize) - 1;
			}

			// header
			sbPageHtml.Append( "<body>" );
			// internal form
			sbPageHtml.Append( "<form name='FormDR' action='" );
			// url
			sbPageHtml.Append( "BufferTargetSelection.asp?" );
			sbPageHtml.Append( "bookmark=" );
			sbPageHtml.Append( _bookMark );
			sbPageHtml.Append( "' method='post'>" );
			// store initial data in hidden fields
			sbPageHtml.Append( "<input type='hidden' name='bidentifier' value='" );
			sbPageHtml.Append( _rawIdentifier );
			sbPageHtml.Append( "'>" );
			sbPageHtml.Append( "<input type='hidden' name='btype' value='" );
			sbPageHtml.Append( _rawType );
			sbPageHtml.Append( "'>" );
			sbPageHtml.Append( "<input type='hidden' name='bback' value='" );
			sbPageHtml.Append( _rawBack );
			sbPageHtml.Append( "'>" );
			sbPageHtml.Append( "<input type='hidden' name='bscope' value='" );
			sbPageHtml.Append( _rawScope );
			sbPageHtml.Append( "'>" );
			sbPageHtml.Append( "</form>" );
			sbPageHtml.Append( "<table style='cursor:default;' width='100%' class='clsTabletext' cellpadding='0' cellspacing='0' border='1' ID='Table1'>" );
			sbPageHtml.Append( "<tr height='30'>" );
			sbPageHtml.Append( "<td colspan='6' align='left'>Record(s) " );
			sbPageHtml.Append( ( nStart + 1 ) );
			sbPageHtml.Append( " to " );
			sbPageHtml.Append( ( nEnd + 1 ) );
			sbPageHtml.Append( " of " );
			sbPageHtml.Append( _dtPossibleMatches.Rows.Count );
			sbPageHtml.Append( "&nbsp;&nbsp;" );
			// paging icons & control
			int nPageBack = nStart - nPageSize;
			int nPageForward = nEnd + 1;
			// page back link
			if( nPageBack >= 0 )
			{
				// link
				sbPageHtml.Append( "<a onclick=\"fnTargetNewPage('" );
				sbPageHtml.Append( nPageBack );
				sbPageHtml.Append( "');\">" );
				// back image
				sbPageHtml.Append( @"<img src='../img/ico_backon.gif' border='0'" );
				// tooltip
				sbPageHtml.Append( " alt='previous page'" );
			}
			else
			{
				// back image
				sbPageHtml.Append( @"<img src='../img/ico_back.gif' border='0'" );
			}
			sbPageHtml.Append( ">" );
			if( nPageBack >= 0 )
			{
				// end link
				sbPageHtml.Append( "</a>" );
			}
			sbPageHtml.Append( "&nbsp;" );
			// page forward link
			if( nPageForward < (_dtPossibleMatches.Rows.Count - 1) )
			{
				// link
				sbPageHtml.Append( "<a onclick=\"fnTargetNewPage('" );
				sbPageHtml.Append( nPageForward );
				sbPageHtml.Append( "');\">" );
				sbPageHtml.Append( @"<img src='../img/ico_forwardon.gif' border='0'" );
				// tooltip
				sbPageHtml.Append( " alt='next page'" );
			}
			else
			{
				sbPageHtml.Append( @"<img src='../img/ico_forward.gif' border='0'" );
			}
			sbPageHtml.Append( ">" );
			// end link
			if( nPageForward < (_dtPossibleMatches.Rows.Count - 1) )
			{
				sbPageHtml.Append( "</a>" );
			}
			// page return to data browser icon & print
			// link
			sbPageHtml.Append( @"<a href='./BufferDataBrowser.asp?fltSt=" );
			sbPageHtml.Append( _backlinkStudyId );
			sbPageHtml.Append( "&fltSi=" );
			sbPageHtml.Append( _backlinkSite );
			sbPageHtml.Append( "&fltSj=" );
			sbPageHtml.Append( _backlinkSubjectId );
			sbPageHtml.Append( "&bookmark=0" );
			sbPageHtml.Append( "'>" );
			// back image
			sbPageHtml.Append( @"<img src='../img/ico_previouson.gif' border='0'" );
			// tooltip
			sbPageHtml.Append( " alt='Return to Buffer Data Browser'" );
			sbPageHtml.Append( ">" );
			// end link
			sbPageHtml.Append( "</a>" );
			sbPageHtml.Append( "&nbsp;&nbsp;&nbsp;" );
			// print icon
			sbPageHtml.Append( @"<a href='javascript:window.print();'><img src='../img/ico_print.gif' border='0' alt='Print listing'></a>" );
			sbPageHtml.Append( "</td>" );
			// end first row
			//sbPageHtml.Append( "<td colspan='4'>&nbsp;</td>" );
			sbPageHtml.Append( "</tr>" );
			// render buffer data item are selecting target for
			sbPageHtml.Append( "<td colspan='6'>" );
			sbPageHtml.Append( "<table style='cursor:default;' width='100%' class='clsTabletext' cellpadding='0' cellspacing='0' border='1' ID='Table3'>" );
			sbPageHtml.Append( "<tr height='20' class='clsTableHeaderText'><td colspan='6'><b>Buffer data item selecting target for</b></td></tr>" );
			sbPageHtml.Append( "<tr height='20' class='clsTableHeaderText'>" );
			sbPageHtml.Append( "<td>Study/Site/Subject</td><td>Visit</td><td>eForm</td>" );
			sbPageHtml.Append( "<td>Question</td><td>Value</td><td>Order Date</td></tr>" );
			sbPageHtml.Append( "<tr><td valign='top'>" );
			// study/site/subject
			sbPageHtml.Append( BufferSummary.SubjectLabel( bufferUser, Convert.ToInt32(_identStudy), _identSite, Convert.ToInt32( _identSubject ) ) );
			sbPageHtml.Append( "</td>" );
			// visit
			sbPageHtml.Append( "<td valign='top'>" );
			if ( dtBufferRow.Rows[0]["VISITNAME"] != System.DBNull.Value )
			{
				sbPageHtml.Append( dtBufferRow.Rows[0]["VISITNAME"].ToString() );
				if( _identVisitCycle != "" )
				{	
					sbPageHtml.Append( " [" );
					sbPageHtml.Append( _identVisitCycle );
					sbPageHtml.Append( "]" );
				}
				else
				{
					// DPH 17/03/2006 - Show that cycle is unknown
					sbPageHtml.Append( " [Unknown]" );
				}
			}
			else
			{
				sbPageHtml.Append( "Unknown" );
			}
			sbPageHtml.Append( "</td>" );
			// eForm
			sbPageHtml.Append( "<td valign='top'>" );
			if ( dtBufferRow.Rows[0]["CRFTITLE"] != System.DBNull.Value )
			{
				sbPageHtml.Append( dtBufferRow.Rows[0]["CRFTITLE"].ToString() );
				if( _identEformCycle != "" )
				{	
					sbPageHtml.Append( " [" );
					sbPageHtml.Append( _identEformCycle );
					sbPageHtml.Append( "]" );
				}
				else
				{
					// DPH 17/03/2006 - Show that cycle is unknown
					sbPageHtml.Append( " [Unknown]" );
				}
			}
			else
			{
				sbPageHtml.Append( "Unknown" );
			}
			sbPageHtml.Append( "</td>" );
			// Question
			sbPageHtml.Append( "<td valign='top'>" );
			if ( dtBufferRow.Rows[0]["DATAITEMNAME"] != System.DBNull.Value )
			{
				sbPageHtml.Append( dtBufferRow.Rows[0]["DATAITEMNAME"].ToString() );
				if( _identResponseCycle != "" )
				{	
					if( Convert.ToInt32( _identResponseCycle ) > 1 )
					{
						sbPageHtml.Append( " [" );
						sbPageHtml.Append( _identResponseCycle );
						sbPageHtml.Append( "]" );
					}
				}
			}
			else
			{
				sbPageHtml.Append( "Unknown" );
			}
			sbPageHtml.Append( "</td>" );
			// value
			sbPageHtml.Append( "<td valign='top'>" );
			sbPageHtml.Append( dtBufferRow.Rows[0]["RESPONSEVALUE"].ToString() );
			sbPageHtml.Append( "</td>" );
			// order date
			sbPageHtml.Append( "<td valign='top'>" );
			if ( dtBufferRow.Rows[0]["ORDERDATETIME"] != System.DBNull.Value )
			{
				DateTime dtOrderDate = DateTime.FromOADate( Convert.ToDouble( dtBufferRow.Rows[0]["ORDERDATETIME"] ) );
				sbPageHtml.Append( dtOrderDate.ToString( "dd/MM/yyyy" ) );
			}
			else
			{
				sbPageHtml.Append( "&nbsp;" );
			}
			sbPageHtml.Append( "</td>" );
			sbPageHtml.Append( "</tr></table></td></tr>" );
			// spacer
			sbPageHtml.Append( "<tr height='20'><td colspan='6'>&nbsp;</td></tr>" );
			// main body of data header
			sbPageHtml.Append( "<tr height='20' class='clsTableHeaderText'><td colspan='6'><b>Please select a target from the following list</b></td></tr>" );
			sbPageHtml.Append( "<tr height='20' class='clsTableHeaderText'>" );
			sbPageHtml.Append( "<td>Study/Site/Subject</td><td>Visit</td><td>eForm</td>" );
			sbPageHtml.Append( "<td>Question</td><td>Target Value</td><td>Target Status</td></tr>" );

			// no data detail
			if( _dtPossibleMatches.Rows.Count == 0 )
			{
				// render - There is no buffer data available for the current criteria
				sbPageHtml.Append( "<tr>" );
				sbPageHtml.Append( "<td valign='top' colspan='6'>" );
				sbPageHtml.Append( "There is no possible target data available for the currently selected buffer data question." );
				sbPageHtml.Append( "</td>" );
				sbPageHtml.Append( "</tr>" );
			}
			else
			{
				// set up main loop
				// run from 
				for(int nDataRow = nStart; nDataRow <= nEnd; nDataRow++)
				{
					// new row
					sbPageHtml.Append( "<tr>" );

					// nRowSpanSubject variable holds the number of rows the current column row will span
					if( nRowSpanSubject == 0 )
					{
						// calculate the new span, if there is one, by counting
						// the number of following rows where trial/site/subject/label values 
						// are the same as this one
						// can assume all rows will be for same trial/site/subject
						ArrayList alSubjectRows = new ArrayList();
						alSubjectRows.AddRange( _dtPossibleMatches.Rows );
						nRowSpanSubject = _dtPossibleMatches.Rows.Count; 

						// calculate for page (some rows may have been on last page)
						int nSectionRow = 0;
						for(nSectionRow = 0; nSectionRow < alSubjectRows.Count; nSectionRow++)
						{
							// loop through rows in section
							// Match on - studyid, site, subject, visitid, visitcycle, eformid, eformcycle,
							//					questionid, questioncycle
							if( (((DataRow)alSubjectRows[nSectionRow])["VISITID"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["VISITID"].ToString()) ||
								(((DataRow)alSubjectRows[nSectionRow])["VISITCYCLENUMBER"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["VISITCYCLENUMBER"].ToString()) ||
								(((DataRow)alSubjectRows[nSectionRow])["CRFPAGEID"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["CRFPAGEID"].ToString()) ||
								(((DataRow)alSubjectRows[nSectionRow])["CRFPAGECYCLENUMBER"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"].ToString()) ||
								(((DataRow)alSubjectRows[nSectionRow])["DATAITEMID"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["DATAITEMID"].ToString()) ||
								(((DataRow)alSubjectRows[nSectionRow])["REPEATNUMBER"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["REPEATNUMBER"].ToString())
								)
							{
								break;
							}
						}
						// adjust rowspan
						nRowSpanSubject -= nSectionRow;
 
						// only need to write this <td> once per subject on the page
						// link
						sbPageHtml.Append( GetJavascriptMenuDetail( nDataRow, 0 ) );

						sbPageHtml.Append( "<td valign='top'" );
						// do we need a rowspan?
						if( nRowSpanSubject > 1 )
						{
							sbPageHtml.Append( "rowspan='" );
							sbPageHtml.Append( nRowSpanSubject );
							sbPageHtml.Append( "'" );
						}
						sbPageHtml.Append( ">" );
						sbPageHtml.Append( BufferSummary.SubjectLabel( bufferUser, Convert.ToInt32( _identStudy ),
							_identSite, Convert.ToInt32( _identSubject ) ) );
						sbPageHtml.Append( "</td>" );
						// close link
						sbPageHtml.Append( "</a>" );
					}
					// decrement subject rowspan
					nRowSpanSubject--;

					// visit column
					if( nRowSpanVisit == 0 )
					{
						// calculate the new span, if there is one, by counting
						// the number of following rows where trial/site/subject/label & visit values 
						// are the same as this one
						// do this via a filter on the current rows values
						// assume trial/site/subject are same 
						StringBuilder sbFilter = new StringBuilder();
						// visit
						sbFilter.Append( "(VISITID " );
						if( _dtPossibleMatches.Rows[ nDataRow ]["VISITID"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// visit id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtPossibleMatches.Rows[ nDataRow ]["VISITID"].ToString() );
						}
						sbFilter.Append( ") AND " );
						// visit cycle
						sbFilter.Append( "(VISITCYCLENUMBER " );
						if( _dtPossibleMatches.Rows[ nDataRow ]["VISITCYCLENUMBER"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// visit id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtPossibleMatches.Rows[ nDataRow ]["VISITCYCLENUMBER"].ToString() );
						}
						sbFilter.Append( " )" );

						// put into arraylist & count
						ArrayList alVisitRows = new ArrayList();
						alVisitRows.AddRange( _dtPossibleMatches.Select( sbFilter.ToString() ) );
						nRowSpanVisit = alVisitRows.Count; 

						// calculate for page (some rows may have been on last page)
						int nSectionRow = 0;
						for(nSectionRow = 0; nSectionRow < alVisitRows.Count; nSectionRow++)
						{
							// loop through rows in section
							if( (((DataRow)alVisitRows[nSectionRow])["VISITID"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["VISITID"].ToString()) ||
								(((DataRow)alVisitRows[nSectionRow])["VISITCYCLENUMBER"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["VISITCYCLENUMBER"].ToString()) ||
								(((DataRow)alVisitRows[nSectionRow])["CRFPAGEID"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["CRFPAGEID"].ToString()) ||
								(((DataRow)alVisitRows[nSectionRow])["CRFPAGECYCLENUMBER"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"].ToString()) ||
								(((DataRow)alVisitRows[nSectionRow])["DATAITEMID"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["DATAITEMID"].ToString()) ||
								(((DataRow)alVisitRows[nSectionRow])["REPEATNUMBER"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["REPEATNUMBER"].ToString())
								)
							{
								break;
							}
						}
						// adjust rowspan
						nRowSpanVisit -= nSectionRow;

						// only need to write this <td> once per visit on the page
						// link
						sbPageHtml.Append( GetJavascriptMenuDetail( nDataRow, 1 ) );
						sbPageHtml.Append( "<td valign='top'" );
						// do we need a rowspan?
						if( nRowSpanVisit > 1 )
						{
							sbPageHtml.Append( "rowspan='" );
							sbPageHtml.Append( nRowSpanVisit );
							sbPageHtml.Append( "'" );
						}
						sbPageHtml.Append( ">" );
						// visit description
						if( _dtPossibleMatches.Rows[ nDataRow ]["VISITNAME"] != System.DBNull.Value )
						{
							sbPageHtml.Append( _dtPossibleMatches.Rows[ nDataRow ]["VISITNAME"].ToString() );
							if( _dtPossibleMatches.Rows[ nDataRow ]["VISITCYCLENUMBER"] != System.DBNull.Value )
							{	
								sbPageHtml.Append( " [" );
								sbPageHtml.Append( _dtPossibleMatches.Rows[ nDataRow ]["VISITCYCLENUMBER"].ToString() );
								sbPageHtml.Append( "]" );
							}
						}
						else
						{
							sbPageHtml.Append( "Unknown" );
						}
						sbPageHtml.Append( "</td>" );
						// close link
						sbPageHtml.Append( "</a>" );
					}
					// decrement visit rowspan
					nRowSpanVisit--;

					// eform column
					if( nRowSpanEform == 0 )
					{
						// calculate the new span, if there is one, by counting
						// the number of following rows where trial/site/subject/label & visit values 
						// are the same as this one
						// do this via a filter on the current rows values
						// can assume all rows will be for same trial/site/subject
						StringBuilder sbFilter = new StringBuilder();
						// trial id
						// visit
						sbFilter.Append( "(VISITID " );
						if( _dtPossibleMatches.Rows[ nDataRow ]["VISITID"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// visit id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtPossibleMatches.Rows[ nDataRow ]["VISITID"].ToString() );
						}
						sbFilter.Append( ") AND " );
						// visit cycle
						sbFilter.Append( "(VISITCYCLENUMBER " );
						if( _dtPossibleMatches.Rows[ nDataRow ]["VISITCYCLENUMBER"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// visit id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtPossibleMatches.Rows[ nDataRow ]["VISITCYCLENUMBER"].ToString() );
						}
						sbFilter.Append( " ) AND " );
						// eform
						sbFilter.Append( "(CRFPAGEID " );
						if( _dtPossibleMatches.Rows[ nDataRow ]["CRFPAGEID"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// crfpage id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtPossibleMatches.Rows[ nDataRow ]["CRFPAGEID"].ToString() );
						}
						sbFilter.Append( " ) AND " );
						// eform cycle
						sbFilter.Append( "(CRFPAGECYCLENUMBER " );
						if( _dtPossibleMatches.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// crfpage id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtPossibleMatches.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"].ToString() );
						}
						sbFilter.Append( " ) " );

						// put into arraylist & count
						ArrayList alEformRows = new ArrayList();
						alEformRows.AddRange( _dtPossibleMatches.Select( sbFilter.ToString() ) );
						nRowSpanEform = alEformRows.Count; 

						// calculate for page (some rows may have been on last page)
						int nSectionRow = 0;
						for(nSectionRow = 0; nSectionRow < alEformRows.Count; nSectionRow++)
						{
							// loop through rows in section
							if( (((DataRow)alEformRows[nSectionRow])["VISITID"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["VISITID"].ToString()) ||
								(((DataRow)alEformRows[nSectionRow])["VISITCYCLENUMBER"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["VISITCYCLENUMBER"].ToString()) ||
								(((DataRow)alEformRows[nSectionRow])["CRFPAGEID"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["CRFPAGEID"].ToString()) ||
								(((DataRow)alEformRows[nSectionRow])["CRFPAGECYCLENUMBER"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"].ToString()) ||
								(((DataRow)alEformRows[nSectionRow])["DATAITEMID"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["DATAITEMID"].ToString()) ||
								(((DataRow)alEformRows[nSectionRow])["REPEATNUMBER"].ToString() == 
								_dtPossibleMatches.Rows[ nDataRow ]["REPEATNUMBER"].ToString())
								)
							{
								break;
							}
						}
						// adjust rowspan
						nRowSpanEform -= nSectionRow;

						// only need to write this <td> once per visit on the page
						// link
						sbPageHtml.Append( GetJavascriptMenuDetail( nDataRow, 2 ) );
						sbPageHtml.Append( "<td valign='top'" );
						// do we need a rowspan?
						if( nRowSpanEform > 1 )
						{
							sbPageHtml.Append( "rowspan='" );
							sbPageHtml.Append( nRowSpanEform );
							sbPageHtml.Append( "'" );
						}
						sbPageHtml.Append( ">" );
						// eform description
						if( _dtPossibleMatches.Rows[ nDataRow ]["CRFTITLE"] != System.DBNull.Value )
						{
							sbPageHtml.Append( _dtPossibleMatches.Rows[ nDataRow ]["CRFTITLE"].ToString() );
							if( _dtPossibleMatches.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"] != System.DBNull.Value )
							{	
								sbPageHtml.Append( " [" );
								sbPageHtml.Append( _dtPossibleMatches.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"].ToString() );
								sbPageHtml.Append( "]" );
							}
						}
						else
						{
							sbPageHtml.Append( "Unknown" );
						}
						sbPageHtml.Append( "</td>" );
						// close link
						sbPageHtml.Append( "</a>" );
					}
					// decrement eform rowspan
					nRowSpanEform--;

					// response detail
					// link
					sbPageHtml.Append( GetJavascriptMenuDetail( nDataRow, 3 ) );
					// question name
					sbPageHtml.Append( "<td valign='top'>" );
					if( _dtPossibleMatches.Rows[ nDataRow ]["DATAITEMNAME"] != System.DBNull.Value )
					{
						sbPageHtml.Append( _dtPossibleMatches.Rows[ nDataRow ]["DATAITEMNAME"].ToString() );
						// if greater than 1 Display response repeat number 
						if( _dtPossibleMatches.Rows[ nDataRow ]["REPEATNUMBER"] != System.DBNull.Value )
						{	
							if( Convert.ToInt32( _dtPossibleMatches.Rows[ nDataRow ]["REPEATNUMBER"] ) > 1 )
							{
								sbPageHtml.Append( " [" );
								sbPageHtml.Append( _dtPossibleMatches.Rows[ nDataRow ]["REPEATNUMBER"].ToString() );
								sbPageHtml.Append( "]" );
							}
						}
					}
					else
					{
						sbPageHtml.Append( "&nbsp;" );
					}
					sbPageHtml.Append( "</td>" );
					// target value
					sbPageHtml.Append( "<td valign='top'>" );
					if( _dtPossibleMatches.Rows[ nDataRow ]["RESPONSEVALUE"] != System.DBNull.Value )
					{
						sbPageHtml.Append( _dtPossibleMatches.Rows[ nDataRow ]["RESPONSEVALUE"].ToString() );
					}
					else
					{
						sbPageHtml.Append( "&nbsp;" );
					}
					sbPageHtml.Append( "</td>" );
					// target status
					sbPageHtml.Append( "<td valign='top'>" );
					if( _dtPossibleMatches.Rows[ nDataRow ]["RESPONSESTATUS"] != System.DBNull.Value )
					{
						// get response status & display text
						int nResponseStatus = Convert.ToInt32( _dtPossibleMatches.Rows[ nDataRow ]["RESPONSESTATUS"] );
						int nLockStatus = 0;
						if(	_dtPossibleMatches.Rows[ nDataRow ]["LOCKSTATUS"] != System.DBNull.Value )
						{
							nLockStatus = Convert.ToInt32( _dtPossibleMatches.Rows[ nDataRow ]["LOCKSTATUS"] );
						}
						int nSdvStatus = 0;
						if(	_dtPossibleMatches.Rows[ nDataRow ]["SDVSTATUS"] != System.DBNull.Value )
						{
							nSdvStatus = Convert.ToInt32( _dtPossibleMatches.Rows[ nDataRow ]["SDVSTATUS"] );
						}
						int nDiscStatus = 0;
						if(	_dtPossibleMatches.Rows[ nDataRow ]["DISCREPANCYSTATUS"] != System.DBNull.Value )
						{
							nDiscStatus = Convert.ToInt32( _dtPossibleMatches.Rows[ nDataRow ]["DISCREPANCYSTATUS"] );
						}
						bool bNote = false;
						if( _dtPossibleMatches.Rows[ nDataRow ]["NOTESTATUS"] != null )
						{
							bNote = (Convert.ToInt32( _dtPossibleMatches.Rows[ nDataRow ]["NOTESTATUS"] ) > 0)?true:false;
						}
						bool bComment = false;
						if( _dtPossibleMatches.Rows[ nDataRow ]["COMMENTS"] != null )
						{
							bComment = (Convert.ToString( _dtPossibleMatches.Rows[ nDataRow ]["COMMENTS"] ).Length > 0)?true:false;
						}
						sbPageHtml.Append( BufferBrowser.ResponseStatusIcon( nResponseStatus, nLockStatus, nSdvStatus, nDiscStatus, bNote, bComment ) );
					}
					else
					{
						sbPageHtml.Append( "&nbsp;" );
					}
					sbPageHtml.Append( "</td>" );
					// close link
					sbPageHtml.Append( "</a>" );
					// end row
					sbPageHtml.Append( "</tr>" );
				}
			}

			// close main table
			sbPageHtml.Append( "</table>" );
			// close body tag
			sbPageHtml.Append( "</body>" );

			return sbPageHtml.ToString();
		}

		/// <summary>
		/// Get possible matches sql
		/// </summary>
		/// <param name="bufferUser"></param>
		/// <returns></returns>
		private DataTable GetPossibleMatches( BufferMACROUser bufferUser )
		{
			// log
			log.Info( "GetPossibleMatches" );

			// get connection string
			string sDbConn = bufferUser.MACROUser.CurrentDBConString;

			// are using sql server MACRO database
			bool bSqlServer = (bufferUser.MACROUser.Database.DatabaseType == MACROUserBS30.eMACRODatabaseType.mdtSQLServer )?true:false;

			// loop through dataset and attempt to get exact matches for unknown items
			ArrayList alDataItems = new ArrayList();

			// identify unknown items data item ids
			// get data item id
			long lDataItemId = Convert.ToInt32( _identEformElementId );

			// create one datatable containing all unknown data items in visit / eform order
			StringBuilder sbSql = new StringBuilder();

			sbSql.Append( "SELECT bufdat.* FROM " );
			sbSql.Append( "( " );
			// known matching data items
			sbSql.Append( "SELECT CRFPAGEINSTANCE.CLINICALTRIALID, CRFPAGEINSTANCE.TRIALSITE, CRFPAGEINSTANCE.PERSONID, CRFPAGEINSTANCE.VISITID, " );
			sbSql.Append( "CRFPAGEINSTANCE.CRFPAGEID, CRFPAGEINSTANCE.VISITCYCLENUMBER, CRFPAGEINSTANCE.CRFPAGECYCLENUMBER, CRFPAGEINSTANCE.CRFPAGETASKID, " );
			sbSql.Append( "CRFPAGEINSTANCE.CRFPAGEDATE, CRFPAGEINSTANCE.CRFPAGESTATUS, CRFPAGEINSTANCE.CRFPAGEINSTANCELABEL, " );
			sbSql.Append( "DATAITEMRESPONSE.DATAITEMID, DATAITEMRESPONSE.RESPONSEVALUE, DATAITEMRESPONSE.RESPONSETASKID, DATAITEMRESPONSE.REPEATNUMBER, " );
			sbSql.Append( "DATAITEMRESPONSE.RESPONSESTATUS, DATAITEMRESPONSE.LOCKSTATUS, DATAITEMRESPONSE.SDVSTATUS, DATAITEMRESPONSE.DISCREPANCYSTATUS, DATAITEMRESPONSE.NOTESTATUS, DATAITEMRESPONSE.COMMENTS, " );
			sbSql.Append( "CLINICALTRIAL.CLINICALTRIALNAME, STUDYVISIT.VISITCODE, CRFPAGE.CRFPAGECODE, STUDYVISIT.VISITNAME, CRFPAGE.CRFTITLE, DATAITEM.DATAITEMNAME, STUDYVISIT.VISITORDER, CRFPAGE.CRFPAGEORDER, CRFELEMENT.FIELDORDER " );
			sbSql.Append( "FROM DATAITEMRESPONSE, CRFELEMENT, CRFPAGEINSTANCE, CRFPAGE, DATAITEM, STUDYVISIT, CLINICALTRIAL " );
			sbSql.Append( "WHERE (DATAITEMRESPONSE.CLINICALTRIALID = CRFELEMENT.CLINICALTRIALID) AND (DATAITEMRESPONSE.CRFPAGEID = CRFELEMENT.CRFPAGEID) " );
			sbSql.Append( "AND (DATAITEMRESPONSE.CRFELEMENTID = CRFELEMENT.CRFELEMENTID) " );
			sbSql.Append( "AND (DATAITEMRESPONSE.CLINICALTRIALID = CRFPAGEINSTANCE.CLINICALTRIALID) AND (DATAITEMRESPONSE.TRIALSITE = CRFPAGEINSTANCE.TRIALSITE) " );
			sbSql.Append( "AND (DATAITEMRESPONSE.PERSONID = CRFPAGEINSTANCE.PERSONID) AND (DATAITEMRESPONSE.CRFPAGETASKID = CRFPAGEINSTANCE.CRFPAGETASKID) " );
			sbSql.Append( "AND (DATAITEMRESPONSE.CLINICALTRIALID = CRFPAGE.CLINICALTRIALID) AND (CRFPAGEINSTANCE.CRFPAGEID = CRFPAGE.CRFPAGEID) " );
			sbSql.Append( "AND (DATAITEMRESPONSE.CLINICALTRIALID = STUDYVISIT.CLINICALTRIALID) AND (CRFPAGEINSTANCE.VISITID = STUDYVISIT.VISITID) " );
			sbSql.Append( "AND (DATAITEMRESPONSE.CLINICALTRIALID = DATAITEM.CLINICALTRIALID) AND (DATAITEMRESPONSE.DATAITEMID = DATAITEM.DATAITEMID) " );
			sbSql.Append( "AND (DATAITEMRESPONSE.CLINICALTRIALID = CLINICALTRIAL.CLINICALTRIALID) " );
			sbSql.Append( "AND (CRFPAGEINSTANCE.CRFPAGESTATUS > -10 ) AND ( (CRFPAGEINSTANCE.CRFPAGEDATE = 0) OR (CRFPAGEINSTANCE.CRFPAGEDATE <  " );
			sbSql.Append( IMEDFunctions20.LocalNumToStandard( DateTime.Now.ToOADate().ToString(), false ) );
			sbSql.Append( " ) )");
			sbSql.Append( "AND (DATAITEMRESPONSE.CLINICALTRIALID = " );
			sbSql.Append( _identStudy );
			sbSql.Append( " ) AND (DATAITEMRESPONSE.TRIALSITE = '" );
			sbSql.Append( _identSite );
			sbSql.Append( "') AND ( DATAITEMRESPONSE.PERSONID = " );
			sbSql.Append( _identSubject );
			sbSql.Append( " )  AND ( DATAITEMRESPONSE.DATAITEMID = " );
			sbSql.Append( lDataItemId );
			sbSql.Append( " ) " );
			sbSql.Append( "UNION " );
			// unknown matching responses (not yet created)
			sbSql.Append( "SELECT CRFPAGEINSTANCE.CLINICALTRIALID, CRFPAGEINSTANCE.TRIALSITE, CRFPAGEINSTANCE.PERSONID, CRFPAGEINSTANCE.VISITID, " );
			sbSql.Append( "CRFPAGEINSTANCE.CRFPAGEID, CRFPAGEINSTANCE.VISITCYCLENUMBER, CRFPAGEINSTANCE.CRFPAGECYCLENUMBER, CRFPAGEINSTANCE.CRFPAGETASKID, " );
			sbSql.Append( "CRFPAGEINSTANCE.CRFPAGEDATE, CRFPAGEINSTANCE.CRFPAGESTATUS, CRFPAGEINSTANCE.CRFPAGEINSTANCELABEL, " );
			sbSql.Append( "DATAITEM.DATAITEMID, " );
			// sql server
			if( bSqlServer )
			{
				sbSql.Append( "null RESPONSEVALUE, null RESPONSETASKID, 1 REPEATNUMBER, " );
				sbSql.Append( "null RESPONSESTATUS, null LOCKSTATUS, null SDVSTATUS, null DISCREPANCYSTATUS, null NOTESTATUS, null COMMENTS, " );
			}
			else
			{
				// oracle
				sbSql.Append( "'' RESPONSEVALUE, to_number(null) RESPONSETASKID, to_number(1) REPEATNUMBER, " );
				sbSql.Append( "to_number(null) RESPONSESTATUS, to_number(null) LOCKSTATUS, to_number(null) SDVSTATUS, to_number(null) DISCREPANCYSTATUS, to_number(null) NOTESTATUS, '' COMMENTS, " );
			}
			sbSql.Append( "CLINICALTRIAL.CLINICALTRIALNAME, STUDYVISIT.VISITCODE, CRFPAGE.CRFPAGECODE, STUDYVISIT.VISITNAME, CRFPAGE.CRFTITLE, DATAITEM.DATAITEMNAME, STUDYVISIT.VISITORDER, CRFPAGE.CRFPAGEORDER, CRFELEMENT.FIELDORDER " );
			sbSql.Append( "FROM CRFPAGEINSTANCE, CRFELEMENT, CRFPAGE, DATAITEM, STUDYVISIT, CLINICALTRIAL " );
			sbSql.Append( "WHERE (CRFPAGEINSTANCE.CLINICALTRIALID = CRFELEMENT.CLINICALTRIALID) AND (CRFPAGEINSTANCE.CRFPAGEID = CRFELEMENT.CRFPAGEID) " );
			sbSql.Append( "AND (CRFPAGEINSTANCE.CLINICALTRIALID = CRFPAGE.CLINICALTRIALID) AND (CRFPAGEINSTANCE.CRFPAGEID = CRFPAGE.CRFPAGEID) " );
			sbSql.Append( "AND (CRFPAGEINSTANCE.CLINICALTRIALID = DATAITEM.CLINICALTRIALID) AND (CRFELEMENT.DATAITEMID = DATAITEM.DATAITEMID) " );
			sbSql.Append( "AND (CRFPAGEINSTANCE.CLINICALTRIALID = STUDYVISIT.CLINICALTRIALID) AND (CRFPAGEINSTANCE.VISITID = STUDYVISIT.VISITID) " );
			sbSql.Append( "AND (CRFPAGEINSTANCE.CLINICALTRIALID = CLINICALTRIAL.CLINICALTRIALID) " );
			sbSql.Append( "AND (CRFPAGEINSTANCE.CRFPAGESTATUS = -10 ) AND ( (CRFPAGEINSTANCE.CRFPAGEDATE = 0) OR (CRFPAGEINSTANCE.CRFPAGEDATE <  " );
			sbSql.Append( IMEDFunctions20.LocalNumToStandard( DateTime.Now.ToOADate().ToString(), false ) );
			sbSql.Append( " ) )");
			sbSql.Append( "AND (CRFPAGEINSTANCE.CLINICALTRIALID = " );
			sbSql.Append( _identStudy );
			sbSql.Append( ") AND (CRFPAGEINSTANCE.TRIALSITE = '" );
			sbSql.Append( _identSite );
			sbSql.Append( "') AND (CRFPAGEINSTANCE.PERSONID = " );
			sbSql.Append( _identSubject );
			sbSql.Append( ") AND (DATAITEM.DATAITEMID = " );
			sbSql.Append( lDataItemId );
			sbSql.Append( ") AND (CRFELEMENT.HIDDEN = 0) " );
			sbSql.Append( ") bufdat " );
			// order by
			sbSql.Append( "ORDER BY bufdat.VISITORDER, bufdat.VISITID, bufdat.VISITCYCLENUMBER, " );
			sbSql.Append( "bufdat.CRFPAGEORDER, bufdat.CRFPAGEID, bufdat.CRFPAGECYCLENUMBER, bufdat.FIELDORDER, bufdat.REPEATNUMBER " );

			// log sql
			log.Debug ( "GetPossibleMatches sql = " + sbSql.ToString() );

			// get datatable
			DataTable dtPossibleMatchList = BufferSummary.GetDataTable(sDbConn, sbSql.ToString() );

			// return it
			return dtPossibleMatchList;
		}

		/// <summary>
		/// Get unique buffer row
		/// </summary>
		/// <param name="bufferUser"></param>
		/// <param name="bufferId"></param>
		/// <returns></returns>
		private DataTable GetBufferRowData( BufferMACROUser bufferUser, string bufferId )
		{
			// log
			log.Info( "Starting GetBufferRowData" );

			// get connection string
			string sDbConn = bufferUser.MACROUser.CurrentDBConString;

			// are using sql server MACRO database
			bool bSqlServerDb = (bufferUser.MACROUser.Database.DatabaseType == MACROUserBS30.eMACRODatabaseType.mdtSQLServer )?true:false;

			// create one datatable containing all unknown data items in visit / eform order
			StringBuilder sbSql = new StringBuilder();

			sbSql.Append( "SELECT BUFFERRESPONSEDATA.RESPONSEVALUE, BUFFERRESPONSEDATA.ORDERDATETIME, STUDYVISIT.VISITNAME, " );
			sbSql.Append( "CRFPAGE.CRFTITLE, DATAITEM.DATAITEMNAME " );

			if( bSqlServerDb )
			{
				// sql server
				sbSql.Append( "FROM BUFFERRESPONSEDATA " );
				sbSql.Append( "LEFT JOIN STUDYVISIT ON (BUFFERRESPONSEDATA.CLINICALTRIALID = STUDYVISIT.CLINICALTRIALID) AND (BUFFERRESPONSEDATA.VISITID = STUDYVISIT.VISITID) " );
				sbSql.Append( "LEFT JOIN CRFPAGE ON (BUFFERRESPONSEDATA.CLINICALTRIALID = CRFPAGE.CLINICALTRIALID) AND (BUFFERRESPONSEDATA.CRFPAGEID = CRFPAGE.CRFPAGEID) " );
				sbSql.Append( "LEFT JOIN DATAITEM ON (BUFFERRESPONSEDATA.CLINICALTRIALID = DATAITEM.CLINICALTRIALID) AND (BUFFERRESPONSEDATA.DATAITEMID = DATAITEM.DATAITEMID) " );
				sbSql.Append( "WHERE " );
			}
			else
			{
				// oracle
				sbSql.Append( "FROM BUFFERRESPONSEDATA, STUDYVISIT, CRFPAGE, DATAITEM " );
				sbSql.Append( "WHERE (BUFFERRESPONSEDATA.CLINICALTRIALID = STUDYVISIT.CLINICALTRIALID(+)) AND (BUFFERRESPONSEDATA.VISITID = STUDYVISIT.VISITID(+)) " );
				sbSql.Append( "AND (BUFFERRESPONSEDATA.CLINICALTRIALID = CRFPAGE.CLINICALTRIALID(+)) AND (BUFFERRESPONSEDATA.CRFPAGEID = CRFPAGE.CRFPAGEID(+)) " );
				sbSql.Append( "AND (BUFFERRESPONSEDATA.CLINICALTRIALID = DATAITEM.CLINICALTRIALID(+)) AND (BUFFERRESPONSEDATA.DATAITEMID = DATAITEM.DATAITEMID(+)) " );
				sbSql.Append( "AND " );
			}
			sbSql.Append( "(BUFFERRESPONSEDATA.BUFFERRESPONSEDATAID = '" );
			sbSql.Append( bufferId );
			sbSql.Append( "')" );

			// log sql
			log.Debug ( "GetBufferRowData sql = " + sbSql.ToString() );

			// get datatable
			DataTable dtBufferRow = BufferSummary.GetDataTable(sDbConn, sbSql.ToString() );

			// return it
			return dtBufferRow;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="targetRow"></param>
		/// <param name="scope"></param>
		/// <returns></returns>
		private string GetJavascriptMenuDetail(int targetRow, int scope)
		{
			// fnSel - button, item detail *, back detail +,  name, value, save flag, goto eform flag
			// * item detail - study id`site`subject id`visit id`visit cycle`visit code`eform id`eform task id`eform cycle`eform code`eform element id`response cycle`response task id`buffer response id
			// + back detail - study id`site`subject id
			// i.e. <a onmouseup="fnSel(event.button,'10`site1`1','10`site1`1','ABC 1','',0,0);">
			StringBuilder sbJS = new StringBuilder();

			// button
			sbJS.Append( "<a onmouseup=\"fnSel(event.button,'" );
			// * item detail - study id`site`subject id`visit id`visit cycle`visit code`eform id`eform task id`eform cycle`eform code`eform element id`response cycle`response task id`buffer response id
			// study id
			sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["CLINICALTRIALID"] );
			sbJS.Append( "`" );
			// site
			sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["TRIALSITE"] );
			sbJS.Append( "`" );
			// subject id
			sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["PERSONID"] );
			sbJS.Append( "`" );
			// visit info
			if( scope > 0 )
			{
				// visit id
				if( _dtPossibleMatches.Rows[ targetRow ]["VISITID"] != System.DBNull.Value )
				{
					// visit id value
					sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["VISITID"] );
				}
				sbJS.Append( "`" );
				// visit cycle
				if( _dtPossibleMatches.Rows[ targetRow ]["VISITCYCLENUMBER"] != System.DBNull.Value )
				{
					sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["VISITCYCLENUMBER"] );
				}
				sbJS.Append( "`" );
				// visit code
				if( _dtPossibleMatches.Rows[ targetRow ]["VISITCODE"] != System.DBNull.Value )
				{
					sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["VISITCODE"] );
				}
				sbJS.Append( "`" );
			}
			else
			{
				sbJS.Append( "```" );
			}
			// eform info
			if( scope > 1 )
			{
				// eform id
				if( _dtPossibleMatches.Rows[ targetRow ]["CRFPAGEID"] != System.DBNull.Value )
				{
					sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["CRFPAGEID"] );
				}
				sbJS.Append( "`" );
				// eform task id
				if( _dtPossibleMatches.Rows[ targetRow ]["CRFPAGETASKID"] != System.DBNull.Value )
				{
					sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["CRFPAGETASKID"] );
				}
				sbJS.Append( "`" );
				// eform cycle
				if( _dtPossibleMatches.Rows[ targetRow ]["CRFPAGECYCLENUMBER"] != System.DBNull.Value )
				{
					sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["CRFPAGECYCLENUMBER"] );
				}
				sbJS.Append( "`" );
				// eform code
				if( _dtPossibleMatches.Rows[ targetRow ]["CRFPAGECODE"] != System.DBNull.Value )
				{
					sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["CRFPAGECODE"] );
				}
				sbJS.Append( "`" );
			}
			else
			{
				sbJS.Append( "````" );
			}
			// question level
			if( scope > 2 )
			{
				// eform element id
				if( _dtPossibleMatches.Rows[ targetRow ]["DATAITEMID"] != System.DBNull.Value )
				{
					sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["DATAITEMID"] );
				}
				sbJS.Append( "`" );
				// response cycle
				if( _dtPossibleMatches.Rows[ targetRow ]["REPEATNUMBER"] != System.DBNull.Value )
				{
					sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["REPEATNUMBER"] );
				}
				sbJS.Append( "`" );
				// response task id  - may not know
				sbJS.Append( "`" );
				// individual row guid - OF Buffer Item changing / selecting target of
				sbJS.Append( _identBufferResponseId );
				sbJS.Append( "" );
			}
			else
			{
				sbJS.Append( "```" );
			}
			sbJS.Append( "'" );
			// back detail - study id`site`subject id
			sbJS.Append( ",'" );
			// study id
			sbJS.Append( _backlinkStudyId );
			sbJS.Append( "`" );
			// site
			sbJS.Append( _backlinkSite );
			sbJS.Append( "`" );
			// subject id
			sbJS.Append( _backlinkSubjectId );
			// name
			sbJS.Append( "','" );
			switch ( scope )
			{
				case 0:
				{
					// subject
					sbJS.Append( BufferBrowser.SubjectLabel( _dtPossibleMatches.Rows[ targetRow ]["CLINICALTRIALNAME"].ToString(), 
						_dtPossibleMatches.Rows[ targetRow ]["TRIALSITE"].ToString(), 
						_dtPossibleMatches.Rows[ targetRow ]["PERSONID"].ToString() ,
						Convert.ToInt32( _dtPossibleMatches.Rows[ targetRow ]["PERSONID"] ) ) );
					break;
				}
				case 1:
				{
					// visit
					if( _dtPossibleMatches.Rows[ targetRow ]["VISITNAME"] != System.DBNull.Value )
					{
						sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["VISITNAME"] );
					}
					break;
				}
				case 2:
				{
					// eform
					if( _dtPossibleMatches.Rows[ targetRow ]["CRFTITLE"] != System.DBNull.Value )
					{
						sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["CRFTITLE"] );
					}
					break;
				}
				case 3:
				{
					// question
					if( _dtPossibleMatches.Rows[ targetRow ]["DATAITEMNAME"] != System.DBNull.Value )
					{
						sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["DATAITEMNAME"] );
					}
					break;
				}
			}
			sbJS.Append( "'," );
			// Value
			sbJS.Append( "'" );
			if( scope == 3)
			{
				if( _dtPossibleMatches.Rows[ targetRow ]["RESPONSEVALUE"] != System.DBNull.Value )
				{
					sbJS.Append( _dtPossibleMatches.Rows[ targetRow ]["RESPONSEVALUE"] );
				}
			}
			sbJS.Append( "'," );
			// save flag
			if( scope == 3 )
			{
				// target save flag true
				sbJS.Append ( "1," );
			}
			else
			{
				// target save flag false
				sbJS.Append ( "0," );
			}
			// goto eform flag
			if( (scope == 2) || (scope == 3) )
			{
				sbJS.Append( "1" );
			}
			else
			{
				sbJS.Append( "0" );
			}
			// close javascript
			sbJS.Append( ");\">" );

			// return
			return sbJS.ToString();
		}
	}
}
