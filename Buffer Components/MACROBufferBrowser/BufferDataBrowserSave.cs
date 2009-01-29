using System;
using System.Text;
using System.Collections;
using System.Data;
using log4net;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using InferMed.MACRO.API;
using InferMed.Components;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for BufferDataBrowserSave.
	/// No paging in first version - possible to use serialised results string stored in page?
	/// </summary>
	class BufferDataBrowserSave
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
		// Buffer Response Ids committed list
		private ArrayList _alBufferResponseList;
		// Datatable to hold diplay results
		private DataTable _dtDisplayResults;

		// log4net
		private static readonly ILog log = LogManager.GetLogger( typeof(BufferDataBrowserSave) );

		/// <summary>
		/// 
		/// </summary>
		/// <param name="bufferUser"></param>
		/// <param name="formData"></param>
		public BufferDataBrowserSave(string formData)
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

			ExtractFormData( formData );
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
						// scope of save 0 - subject 1 - visit 2 - eform 3 - question
						string sSaveScope = BufferBrowser.ReplaceHtmlCharacters( alIndivData[1].ToString() );
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
		public string RenderBufferSaveResultsPage(BufferMACROUser bufferUser)
		{
			log.Info( "Starting RenderBufferSaveResultsPage" );

			StringBuilder sbPageHtml = new StringBuilder();

			// store results to MACRO database via MACRO API
			CommitDataToMACRO( bufferUser );

			if( _action == 1 )
			{
				// show save results
				// now display results
				// rowspan variables
				int nRowSpanSubject = 0;
				int nRowSpanVisit = 0;
				int nRowSpanEform = 0;

				// all data results should be ordered in _dtDisplayResults datatable
				// body tag
				// DPH 02/08/2006 - Added page load event to allow for 'please wait' onscreen
				sbPageHtml.Append( "<body onload=\"fnPageLoaded();\">" );

				// get success count
				int nSuccessful = 0;
				foreach(BufferResponseItem bufferResponseItem in _alBufferResponseList)
				{
					if( bufferResponseItem.BufferCommitStatus == BufferBrowser.BufferCommitStatus.Success)
					{
						nSuccessful++;
					}
				}

				// main table
				sbPageHtml.Append( "<table style='cursor:default;' width='100%' class='clsTabletext' cellpadding='0' cellspacing='0' border='1' ID=\"Table1\">" );
				// info bar row
				sbPageHtml.Append( "<tr height='30'>" );
				sbPageHtml.Append( "<td colspan='7' align='left' valign='middle'>Record(s) successfully committed " );
				// no of records successful
				sbPageHtml.Append( nSuccessful );
				sbPageHtml.Append( " of " );
				// total attempted save on page
				sbPageHtml.Append( _dtDisplayResults.Rows.Count );
				sbPageHtml.Append( "&nbsp;&nbsp;" );
				// page return to data browser icon & print
				// page return to data browser link
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
				sbPageHtml.Append( "</tr>" );

				// Header and columns
				sbPageHtml.Append( "<tr height='20' class='clsTableHeaderText'>" );
				sbPageHtml.Append( "<td>Study/Site/Subject</td>" );
				sbPageHtml.Append( "<td>Visit</td>" );
				sbPageHtml.Append( "<td>eForm</td>" );
				sbPageHtml.Append( "<td>Question</td>" );
				sbPageHtml.Append( "<td>Value</td>" );
				sbPageHtml.Append( "<td>Status</td>" );
				sbPageHtml.Append( "<td>Warning / Reason for Fail</td>" );
				sbPageHtml.Append( "</tr>" );

				// set up main loop - loop through all
				for(int nDataRow = 0; nDataRow < _dtDisplayResults.Rows.Count; nDataRow++)
				{
					// new row
					sbPageHtml.Append( "<tr>" );
				
					// find matching results row
					// store result / failure result
					BufferBrowser.BufferCommitStatus eResponseCommitStatus = BufferBrowser.BufferCommitStatus.NotCommitted;
					string sFailRejectMessage = "";
					foreach( BufferResponseItem bufferResponseItem in _alBufferResponseList)
					{
						// if is a matching response row
						if( _dtDisplayResults.Rows[ nDataRow ]["BUFFERRESPONSEDATAID"].ToString() == bufferResponseItem.BufferResponseId )
						{
							// get result
							eResponseCommitStatus = bufferResponseItem.BufferCommitStatus;
							// extract message (if any)
							sFailRejectMessage = bufferResponseItem.FailRejectWarnMessage;
						}
					}

					// nRowSpanSubject variable holds the number of rows the current column row will span
					if( nRowSpanSubject == 0 )
					{
						// calculate the new span, if there is one, by counting
						// the number of following rows where trial/site/subject/label values 
						// are the same as this one
						// do this via a filter on the current rows values
						StringBuilder sbFilter = new StringBuilder();
						// trial id
						sbFilter.Append( "( CLINICALTRIALID = " );
						sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["CLINICALTRIALID"] );
						sbFilter.Append( ") AND " );
						// site
						sbFilter.Append( "(SITE = '" );
						sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["SITE"] );
						sbFilter.Append( "') AND " );
						// subject
						sbFilter.Append( "(PERSONID = " );
						sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["PERSONID"] );
						sbFilter.Append( ")" );

						// put into arraylist & count
						ArrayList alSubjectRows = new ArrayList();
						alSubjectRows.AddRange( _dtDisplayResults.Select( sbFilter.ToString() ) );
						nRowSpanSubject = alSubjectRows.Count; 

						// calculate for page (some rows may have been on last page)
						/*
						int nSectionRow = 0;
						for(nSectionRow = 0; nSectionRow < alSubjectRows.Count; nSectionRow++)
						{
							// loop through rows in section
							if( ((DataRow)alSubjectRows[nSectionRow])["BUFFERRESPONSEDATAID"].ToString() == 
								_dtBuffer.Rows[ nDataRow ]["BUFFERRESPONSEDATAID"].ToString() )
							{
								break;
							}
						}
						// adjust rowspan
						nRowSpanSubject -= nSectionRow;
						*/

						// only need to write this <td> once per subject on the page
						sbPageHtml.Append( "<td valign='top'" );
						// do we need a rowspan?
						if( nRowSpanSubject > 1 )
						{
							sbPageHtml.Append( "rowspan='" );
							sbPageHtml.Append( nRowSpanSubject );
							sbPageHtml.Append( "'" );
						}
						sbPageHtml.Append( ">" );
						sbPageHtml.Append( BufferBrowser.SubjectLabel( _dtDisplayResults.Rows[ nDataRow ]["CLINICALTRIALNAME"].ToString(), 
							_dtDisplayResults.Rows[ nDataRow ]["SITE"].ToString(), 
							_dtDisplayResults.Rows[ nDataRow ]["SUBJECTLABEL"].ToString(),
							Convert.ToInt32( _dtDisplayResults.Rows[ nDataRow ]["PERSONID"] ) ) );
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
						StringBuilder sbFilter = new StringBuilder();
						// trial id
						sbFilter.Append( "( CLINICALTRIALID = " );
						sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["CLINICALTRIALID"] );
						sbFilter.Append( ") AND " );
						// site
						sbFilter.Append( "(SITE = '" );
						sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["SITE"] );
						sbFilter.Append( "') AND " );
						// subject
						sbFilter.Append( "(PERSONID = " );
						sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["PERSONID"] );
						sbFilter.Append( ") AND " );
						// visit
						sbFilter.Append( "(VISITID " );
						if( _dtDisplayResults.Rows[ nDataRow ]["VISITID"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// visit id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["VISITID"].ToString() );
						}
						sbFilter.Append( ") AND " );
						// visit cycle
						sbFilter.Append( "(VISITCYCLENUMBER " );
						if( _dtDisplayResults.Rows[ nDataRow ]["VISITCYCLENUMBER"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// visit id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["VISITCYCLENUMBER"].ToString() );
						}
						sbFilter.Append( " )" );

						// put into arraylist & count
						ArrayList alVisitRows = new ArrayList();
						alVisitRows.AddRange( _dtDisplayResults.Select( sbFilter.ToString() ) );
						nRowSpanVisit = alVisitRows.Count; 

						// calculate for page (some rows may have been on last page)
						/*
						int nSectionRow = 0;
						for(nSectionRow = 0; nSectionRow < alVisitRows.Count; nSectionRow++)
						{
							// loop through rows in section
							if( ((DataRow)alVisitRows[nSectionRow])["BUFFERRESPONSEDATAID"].ToString() == 
								_dtDisplayResults.Rows[ nDataRow ]["BUFFERRESPONSEDATAID"].ToString() )
							{
								break;
							}
						}
						// adjust rowspan
						nRowSpanVisit -= nSectionRow;
						*/

						// only need to write this <td> once per visit on the page
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
						if( _dtDisplayResults.Rows[ nDataRow ]["VISITNAME"] != System.DBNull.Value )
						{
							sbPageHtml.Append( _dtDisplayResults.Rows[ nDataRow ]["VISITNAME"].ToString() );
							if( _dtDisplayResults.Rows[ nDataRow ]["VISITCYCLENUMBER"] != System.DBNull.Value )
							{	
								sbPageHtml.Append( " [" );
								sbPageHtml.Append( _dtDisplayResults.Rows[ nDataRow ]["VISITCYCLENUMBER"].ToString() );
								sbPageHtml.Append( "]" );
							}
						}
						else
						{
							sbPageHtml.Append( "Unknown" );
						}
						sbPageHtml.Append( "</td>" );
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
						StringBuilder sbFilter = new StringBuilder();
						// trial id
						sbFilter.Append( "( CLINICALTRIALID = " );
						sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["CLINICALTRIALID"] );
						sbFilter.Append( ") AND " );
						// site
						sbFilter.Append( "(SITE = '" );
						sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["SITE"] );
						sbFilter.Append( "') AND " );
						// subject
						sbFilter.Append( "(PERSONID = " );
						sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["PERSONID"] );
						sbFilter.Append( ") AND " );
						// visit
						sbFilter.Append( "(VISITID " );
						if( _dtDisplayResults.Rows[ nDataRow ]["VISITID"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// visit id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["VISITID"].ToString() );
						}
						sbFilter.Append( ") AND " );
						// visit cycle
						sbFilter.Append( "(VISITCYCLENUMBER " );
						if( _dtDisplayResults.Rows[ nDataRow ]["VISITCYCLENUMBER"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// visit id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["VISITCYCLENUMBER"].ToString() );
						}
						sbFilter.Append( " ) AND " );
						// eform
						sbFilter.Append( "(CRFPAGEID " );
						if( _dtDisplayResults.Rows[ nDataRow ]["CRFPAGEID"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// crfpage id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["CRFPAGEID"].ToString() );
						}
						sbFilter.Append( " ) AND " );
						// eform cycle
						sbFilter.Append( "(CRFPAGECYCLENUMBER " );
						if( _dtDisplayResults.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// crfpage id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtDisplayResults.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"].ToString() );
						}
						sbFilter.Append( " ) " );

						// put into arraylist & count
						ArrayList alEformRows = new ArrayList();
						alEformRows.AddRange( _dtDisplayResults.Select( sbFilter.ToString() ) );
						nRowSpanEform = alEformRows.Count; 

						// calculate for page (some rows may have been on last page)
						/*
						int nSectionRow = 0;
						for(nSectionRow = 0; nSectionRow < alEformRows.Count; nSectionRow++)
						{
							// loop through rows in section
							if( ((DataRow)alEformRows[nSectionRow])["BUFFERRESPONSEDATAID"].ToString() == 
								_dtDisplayResults.Rows[ nDataRow ]["BUFFERRESPONSEDATAID"].ToString() )
							{
								break;
							}
						}
						// adjust rowspan
						nRowSpanEform -= nSectionRow;
						*/

						// only need to write this <td> once per visit on the page
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
						if( _dtDisplayResults.Rows[ nDataRow ]["CRFTITLE"] != System.DBNull.Value )
						{
							sbPageHtml.Append( _dtDisplayResults.Rows[ nDataRow ]["CRFTITLE"].ToString() );
							if( _dtDisplayResults.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"] != System.DBNull.Value )
							{	
								sbPageHtml.Append( " [" );
								sbPageHtml.Append( _dtDisplayResults.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"].ToString() );
								sbPageHtml.Append( "]" );
							}
						}
						else
						{
							sbPageHtml.Append( "Unknown" );
						}
						sbPageHtml.Append( "</td>" );
					}
					// decrement eform rowspan
					nRowSpanEform--;

					// response detail
					// mark as an error row?
					bool bFailed = (eResponseCommitStatus==BufferBrowser.BufferCommitStatus.Success)?false:true;
					// question name
					sbPageHtml.Append( "<td valign='top'" );
					if(bFailed)
					{
						sbPageHtml.Append( " bgcolor='Yellow'" );
					}
					sbPageHtml.Append( ">" );
					if( _dtDisplayResults.Rows[ nDataRow ]["DATAITEMNAME"] != System.DBNull.Value )
					{
						sbPageHtml.Append( _dtDisplayResults.Rows[ nDataRow ]["DATAITEMNAME"].ToString() );
						// if greater than 1 Display response repeat number 
						if( _dtDisplayResults.Rows[ nDataRow ]["RESPONSEREPEATNUMBER"] != System.DBNull.Value )
						{	
							if( Convert.ToInt32( _dtDisplayResults.Rows[ nDataRow ]["RESPONSEREPEATNUMBER"] ) > 1 )
							{
								sbPageHtml.Append( " [" );
								sbPageHtml.Append( _dtDisplayResults.Rows[ nDataRow ]["RESPONSEREPEATNUMBER"].ToString() );
								sbPageHtml.Append( "]" );
							}
						}
					}
					else
					{
						sbPageHtml.Append( "&nbsp;" );
					}
					sbPageHtml.Append( "</td>" );
					// value
					sbPageHtml.Append( "<td valign='top'" );
					if(bFailed)
					{
						sbPageHtml.Append( " bgcolor='Yellow'" );
					}
					sbPageHtml.Append( ">" );
					if( _dtDisplayResults.Rows[ nDataRow ]["RESPONSEVALUE"] != System.DBNull.Value )
					{
						sbPageHtml.Append( _dtDisplayResults.Rows[ nDataRow ]["RESPONSEVALUE"].ToString() );
					}
					else
					{
						sbPageHtml.Append( "&nbsp;" );
					}
					sbPageHtml.Append( "</td>" );
					// target status
					sbPageHtml.Append( "<td valign='top'" );
					if(bFailed)
					{
						sbPageHtml.Append( " bgcolor='Yellow'" );
					}
					sbPageHtml.Append( ">" );
					if(bFailed)
					{
						// error icon
						sbPageHtml.Append( "<img alt='Buffer Save Failure' src='../img/ico_error.gif'>" );
					}
					else
					{
						if( _dtDisplayResults.Rows[ nDataRow ]["DIRESPONSESTATUS"] != System.DBNull.Value )
						{
							// get response status & display text
							int nResponseStatus = Convert.ToInt32( _dtDisplayResults.Rows[ nDataRow ]["DIRESPONSESTATUS"] );
							int nLockStatus = 0;
							if(	_dtDisplayResults.Rows[ nDataRow ]["LOCKSTATUS"] != System.DBNull.Value )
							{
								nLockStatus = Convert.ToInt32( _dtDisplayResults.Rows[ nDataRow ]["LOCKSTATUS"] );
							}
							int nSdvStatus = 0;
							if(	_dtDisplayResults.Rows[ nDataRow ]["SDVSTATUS"] != System.DBNull.Value )
							{
								nSdvStatus = Convert.ToInt32( _dtDisplayResults.Rows[ nDataRow ]["SDVSTATUS"] );
							}
							int nDiscStatus = 0;
							if(	_dtDisplayResults.Rows[ nDataRow ]["DISCREPANCYSTATUS"] != System.DBNull.Value )
							{
								nDiscStatus = Convert.ToInt32( _dtDisplayResults.Rows[ nDataRow ]["DISCREPANCYSTATUS"] );
							}
							bool bNote = false;
							if( _dtDisplayResults.Rows[ nDataRow ]["NOTESTATUS"] != null )
							{
								bNote = (Convert.ToInt32( _dtDisplayResults.Rows[ nDataRow ]["NOTESTATUS"] ) > 0)?true:false;
							}
							bool bComment = false;
							if( _dtDisplayResults.Rows[ nDataRow ]["NOTESTATUS"] != null )
							{
								bComment = (Convert.ToString( _dtDisplayResults.Rows[ nDataRow ]["COMMENTS"] ).Length > 0)?true:false;
							}
							sbPageHtml.Append( BufferBrowser.ResponseStatusIcon( nResponseStatus, nLockStatus, nSdvStatus, nDiscStatus, bNote, bComment ) );
						}
						else
						{
							sbPageHtml.Append( "&nbsp;" );
						}
					}
					sbPageHtml.Append( "</td>" );
					// Warning / reason for fail
					sbPageHtml.Append( "<td valign='top'" );
					if(bFailed)
					{
						sbPageHtml.Append( " bgcolor='Yellow'" );
					}
					sbPageHtml.Append( ">" );
					// get reason for fail
					if(sFailRejectMessage != "")
					{
						sbPageHtml.Append( sFailRejectMessage );
					}
					else
					{
						sbPageHtml.Append( "&nbsp;" );
					}
					sbPageHtml.Append( "</td>" );
					// end row
					sbPageHtml.Append( "</tr>" );
				}
				// close main table
				sbPageHtml.Append( "</table>" );
				// close body tag
				sbPageHtml.Append( "</body>" );
			}
			else
			{
				// discard
				// write javascript to automatically move page back to BufferDataBrowser
				sbPageHtml.Append( "<script language=\"javascript\">" );
				sbPageHtml.Append( "window.location=\"" );
				sbPageHtml.Append( @"./BufferDataBrowser.asp?fltSt=" );
				sbPageHtml.Append( _backlinkStudyId );
				sbPageHtml.Append( "&fltSi=" );
				sbPageHtml.Append( _backlinkSite );
				sbPageHtml.Append( "&fltSj=" );
				sbPageHtml.Append( _backlinkSubjectId );
				sbPageHtml.Append( "&bookmark=0" );
				sbPageHtml.Append( "\";" );
				sbPageHtml.Append( "</script>" );
				sbPageHtml.Append( "<body></body>" );
			}

			return sbPageHtml.ToString();
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="bufferUser"></param>
		private void CommitDataToMACRO(BufferMACROUser bufferUser)
		{
			log.Info( "Starting CommitDataToMACRO" );

			// get connection string
			string sDbConn = bufferUser.MACROUser.CurrentDBConString;
			// are using sql server MACRO database?
			bool bSqlServerDb = (bufferUser.MACROUser.Database.DatabaseType == MACROUserBS30.eMACRODatabaseType.mdtSQLServer )?true:false;

			// collect sql to extract rows to commit
			string sTrialDbColumn = "BUFFERRESPONSEDATA.CLINICALTRIALID";
			string sSiteDbColumn = "BUFFERRESPONSEDATA.SITE";
			// DPH 21/11/2007 change to user object since MACRO v3.0.76 led to this change
			MACROUserBS30.MACROUser blankUser = null;
			string sStudySiteSQL = bufferUser.MACROUser.DataLists.StudiesSitesWhereSQL( ref sTrialDbColumn, ref sSiteDbColumn, ref blankUser );

			// extract data to be committed given private members
			StringBuilder sbSql = new StringBuilder();
			// get basic sql
			if( bSqlServerDb )
			{
				// sql server
				sbSql.Append( GetBufferDataSqlServer( sStudySiteSQL ) );
			}
			else
			{
				// oracle
				sbSql.Append( GetBufferDataSqlOracle( sStudySiteSQL ) );
			}
			// pre-commit where clause
			sbSql.Append( GetPreCommitWhereClause() );
			// order by
			sbSql.Append( GetOrderBy() );

			log.Debug( "Collect Data to commit=" + sbSql.ToString() );
			// get DataTable
			DataTable dtDataToCommit = BufferSummary.GetDataTable( sDbConn, sbSql.ToString() );

			// reset Buffer response id list
			_alBufferResponseList = new ArrayList();

			// commit or reject
			switch( _action )
			{
				case 1:
				{
					// commit data
					// form xml string of data in MACRO API format
					// get subject object
					MACROSubject macroSubject = GetMACROSubject( dtDataToCommit );

					// serialize to API compliant xml
					string xmlAPI = SubjectToString( macroSubject );
				
					// open API & commit data
					string MACROCommitResults = "";
					if( CommitToMACRO( bufferUser, xmlAPI, ref MACROCommitResults ) )
					{
						// all done successfully
						SetAllResponseStatuses( BufferBrowser.BufferCommitStatus.Success ); 
					}
					else
					{
						// some failures have occurred
						// DPH 17/03/2006 - set all to success then mark just failures accordingly
						SetAllResponseStatuses( BufferBrowser.BufferCommitStatus.Success ); 
						// review & store fail results
						ProcessAPIResults( MACROCommitResults );
					}

					// update BufferResponseData table with collated results
					UpdateBufferResponse( bufferUser );

					// get datatable to use for displayresults
					sbSql = new StringBuilder();
					// get basic sql
					if( bSqlServerDb )
					{
						// sql server
						sbSql.Append( GetBufferDataSqlServer( sStudySiteSQL ) );
					}
					else
					{
						// oracle
						sbSql.Append( GetBufferDataSqlOracle( sStudySiteSQL ) );
					}
					// pre-commit where clause
					sbSql.Append( GetPostCommitWhereClause() );
					// order by
					sbSql.Append( GetOrderBy() );

					// get DataTable for use when rendering page
					_dtDisplayResults = BufferSummary.GetDataTable( sDbConn, sbSql.ToString() );
					break;
				}
				case 2:
				{
					// discard the data
					// create response list
					CreateBufferResponseList( dtDataToCommit );
					// set all responses 
					SetAllResponseStatuses( BufferBrowser.BufferCommitStatus.DiscardedByUser );
					// update on BufferResponseData tables
					UpdateBufferResponse( bufferUser );
					break;
				}
			}
		}

		/// <summary>
		/// Basic sql to retrieve Buffer response data with additional display info for sql server
		/// </summary>
		/// <param name="sStudySiteSQL"></param>
		/// <returns></returns>
		private string GetBufferDataSqlServer( string sStudySiteSQL )
		{
			log.Info( "Starting GetBufferDataSqlServer" );

			StringBuilder sbSql = new StringBuilder();
			// Full buffer data inc. visit description / eform description / question description 
			// matching response value / status
			// in visit / eForm order
			sbSql.Append( "SELECT BUFFERRESPONSEDATA.BUFFERRESPONSEDATAID, BUFFERRESPONSEDATA.BUFFERRESPONSEID, BUFFERRESPONSEDATA.BUFFERCOMMITSTATUS, BUFFERRESPONSEDATA.BUFFERCOMMITSTATUSTIMESTAMP, " );
			sbSql.Append( "BUFFERRESPONSEDATA.BUFFERCOMMITSTATUSTIMESTAMP_TZ, BUFFERRESPONSEDATA.CLINICALTRIALNAME, BUFFERRESPONSEDATA.CLINICALTRIALID, BUFFERRESPONSEDATA.SITE, BUFFERRESPONSEDATA.SUBJECTLABEL, " );
			sbSql.Append( "BUFFERRESPONSEDATA.PERSONID, BUFFERRESPONSEDATA.VISITCODE, BUFFERRESPONSEDATA.VISITID, BUFFERRESPONSEDATA.VISITCYCLENUMBER, BUFFERRESPONSEDATA.CRFPAGECODE, " );
			sbSql.Append( "BUFFERRESPONSEDATA.CRFPAGEID, BUFFERRESPONSEDATA.CRFPAGECYCLENUMBER, BUFFERRESPONSEDATA.DATAITEMCODE, BUFFERRESPONSEDATA.DATAITEMID, BUFFERRESPONSEDATA.RESPONSEREPEATNUMBER, " );
			sbSql.Append( "BUFFERRESPONSEDATA.RESPONSEVALUE, BUFFERRESPONSEDATA.ORDERDATETIME, BUFFERRESPONSEDATA.USERNAME, BUFFERRESPONSEDATA.USERNAMEFULL, " );
			sbSql.Append( "DATAITEMRESPONSE.RESPONSEVALUE AS DIRESPONSEVALUE, DATAITEMRESPONSE.RESPONSESTATUS AS DIRESPONSESTATUS, DATAITEMRESPONSE.LOCKSTATUS, DATAITEMRESPONSE.DISCREPANCYSTATUS, DATAITEMRESPONSE.SDVSTATUS, DATAITEMRESPONSE.NOTESTATUS, " );
			sbSql.Append( "DATAITEMRESPONSE.COMMENTS, STUDYVISIT.VISITNAME, CRFPAGE.CRFTITLE, DATAITEM.DATAITEMNAME, STUDYVISIT.VISITORDER, CRFPAGE.CRFPAGEORDER, CRFELEMENT.FIELDORDER, " );
			sbSql.Append( "DATAITEMRESPONSE.RESPONSETASKID, DATAITEMRESPONSE.CRFPAGETASKID " );
			sbSql.Append( "FROM BUFFERRESPONSEDATA " );
			sbSql.Append( "LEFT JOIN DATAITEMRESPONSE ON (BUFFERRESPONSEDATA.CLINICALTRIALID = DATAITEMRESPONSE.CLINICALTRIALID) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.SITE = DATAITEMRESPONSE.TRIALSITE) AND (BUFFERRESPONSEDATA.PERSONID = DATAITEMRESPONSE.PERSONID) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.VISITID = DATAITEMRESPONSE.VISITID) AND (BUFFERRESPONSEDATA.VISITCYCLENUMBER = DATAITEMRESPONSE.VISITCYCLENUMBER) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.CRFPAGEID = DATAITEMRESPONSE.CRFPAGEID) AND (BUFFERRESPONSEDATA.CRFPAGECYCLENUMBER = DATAITEMRESPONSE.CRFPAGECYCLENUMBER) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.DATAITEMID = DATAITEMRESPONSE.DATAITEMID) AND (BUFFERRESPONSEDATA.RESPONSEREPEATNUMBER = DATAITEMRESPONSE.REPEATNUMBER) " );
			sbSql.Append( "LEFT JOIN CRFELEMENT ON (BUFFERRESPONSEDATA.CLINICALTRIALID = CRFELEMENT.CLINICALTRIALID) AND (BUFFERRESPONSEDATA.CRFPAGEID = CRFELEMENT.CRFPAGEID) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.DATAITEMID = CRFELEMENT.DATAITEMID) " );
			sbSql.Append( "LEFT JOIN DATAITEM ON (BUFFERRESPONSEDATA.CLINICALTRIALID = DATAITEM.CLINICALTRIALID) AND (BUFFERRESPONSEDATA.DATAITEMID = DATAITEM.DATAITEMID) " );
			sbSql.Append( "LEFT JOIN CRFPAGE ON (BUFFERRESPONSEDATA.CLINICALTRIALID = CRFPAGE.CLINICALTRIALID) AND (BUFFERRESPONSEDATA.CRFPAGEID = CRFPAGE.CRFPAGEID) " );
			sbSql.Append( "LEFT JOIN STUDYVISIT ON (BUFFERRESPONSEDATA.CLINICALTRIALID = STUDYVISIT.CLINICALTRIALID) AND (BUFFERRESPONSEDATA.VISITID = STUDYVISIT.VISITID) " );

			// include where clause containing allowed study / site permissions
			sbSql.Append( "WHERE " );
			sbSql.Append( sStudySiteSQL );
			sbSql.Append( " " );

			return sbSql.ToString();
		}

		/// <summary>
		/// Basic sql to retrieve Buffer response data with additional display info for oracle
		/// </summary>
		/// <param name="sStudySiteSQL"></param>
		/// <returns></returns>
		private string GetBufferDataSqlOracle( string sStudySiteSQL )
		{
			log.Info( "Starting GetBufferDataSqlOracle" );

			StringBuilder sbSql = new StringBuilder();
			// Full buffer data inc. visit description / eform description / question description 
			// matching response value / status
			// in visit / eForm order
			sbSql.Append( "SELECT BUFFERRESPONSEDATA.BUFFERRESPONSEDATAID, BUFFERRESPONSEDATA.BUFFERRESPONSEID, BUFFERRESPONSEDATA.BUFFERCOMMITSTATUS, BUFFERRESPONSEDATA.BUFFERCOMMITSTATUSTIMESTAMP, " );
			sbSql.Append( "BUFFERRESPONSEDATA.BUFFERCOMMITSTATUSTIMESTAMP_TZ, BUFFERRESPONSEDATA.CLINICALTRIALNAME, BUFFERRESPONSEDATA.CLINICALTRIALID, BUFFERRESPONSEDATA.SITE, BUFFERRESPONSEDATA.SUBJECTLABEL, " );
			sbSql.Append( "BUFFERRESPONSEDATA.PERSONID, BUFFERRESPONSEDATA.VISITCODE, BUFFERRESPONSEDATA.VISITID, BUFFERRESPONSEDATA.VISITCYCLENUMBER, BUFFERRESPONSEDATA.CRFPAGECODE, " );
			sbSql.Append( "BUFFERRESPONSEDATA.CRFPAGEID, BUFFERRESPONSEDATA.CRFPAGECYCLENUMBER, BUFFERRESPONSEDATA.DATAITEMCODE, BUFFERRESPONSEDATA.DATAITEMID, BUFFERRESPONSEDATA.RESPONSEREPEATNUMBER, " );
			sbSql.Append( "BUFFERRESPONSEDATA.RESPONSEVALUE, BUFFERRESPONSEDATA.ORDERDATETIME, BUFFERRESPONSEDATA.USERNAME, BUFFERRESPONSEDATA.USERNAMEFULL, " );
			sbSql.Append( "DATAITEMRESPONSE.RESPONSEVALUE AS DIRESPONSEVALUE, DATAITEMRESPONSE.RESPONSESTATUS AS DIRESPONSESTATUS, DATAITEMRESPONSE.LOCKSTATUS, DATAITEMRESPONSE.DISCREPANCYSTATUS, DATAITEMRESPONSE.SDVSTATUS, DATAITEMRESPONSE.NOTESTATUS, " );
			sbSql.Append( "DATAITEMRESPONSE.COMMENTS, STUDYVISIT.VISITNAME, CRFPAGE.CRFTITLE, DATAITEM.DATAITEMNAME, STUDYVISIT.VISITORDER, CRFPAGE.CRFPAGEORDER, CRFELEMENT.FIELDORDER, " );
			sbSql.Append( "DATAITEMRESPONSE.RESPONSETASKID, DATAITEMRESPONSE.CRFPAGETASKID " );
			sbSql.Append( "FROM BUFFERRESPONSEDATA, DATAITEMRESPONSE, CRFELEMENT, DATAITEM, CRFPAGE, STUDYVISIT " );
			sbSql.Append( "WHERE " );
			// oracle specific left joins
			sbSql.Append( "(BUFFERRESPONSEDATA.CLINICALTRIALID = DATAITEMRESPONSE.CLINICALTRIALID(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.SITE = DATAITEMRESPONSE.TRIALSITE(+)) AND (BUFFERRESPONSEDATA.PERSONID = DATAITEMRESPONSE.PERSONID(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.VISITID = DATAITEMRESPONSE.VISITID(+)) AND (BUFFERRESPONSEDATA.VISITCYCLENUMBER = DATAITEMRESPONSE.VISITCYCLENUMBER(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.CRFPAGEID = DATAITEMRESPONSE.CRFPAGEID(+)) AND (BUFFERRESPONSEDATA.CRFPAGECYCLENUMBER = DATAITEMRESPONSE.CRFPAGECYCLENUMBER(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.DATAITEMID = DATAITEMRESPONSE.DATAITEMID(+)) AND (BUFFERRESPONSEDATA.RESPONSEREPEATNUMBER = DATAITEMRESPONSE.REPEATNUMBER(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.CLINICALTRIALID = CRFELEMENT.CLINICALTRIALID(+)) AND (BUFFERRESPONSEDATA.CRFPAGEID = CRFELEMENT.CRFPAGEID(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.DATAITEMID = CRFELEMENT.DATAITEMID(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.CLINICALTRIALID = DATAITEM.CLINICALTRIALID(+)) AND (BUFFERRESPONSEDATA.DATAITEMID = DATAITEM.DATAITEMID(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.CLINICALTRIALID = CRFPAGE.CLINICALTRIALID(+)) AND (BUFFERRESPONSEDATA.CRFPAGEID = CRFPAGE.CRFPAGEID(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.CLINICALTRIALID = STUDYVISIT.CLINICALTRIALID(+)) AND (BUFFERRESPONSEDATA.VISITID = STUDYVISIT.VISITID(+)) " );

			// include where clause containing allowed study / site permissions
			sbSql.Append( " AND " );
			sbSql.Append( sStudySiteSQL );
			sbSql.Append( " " );

			return sbSql.ToString();
		}

		/// <summary>
		/// Create where clause based on member variable data
		/// </summary>
		/// <returns></returns>
		private string GetPreCommitWhereClause()
		{
			log.Info( "Starting GetPreCommitWhereClause" );

			StringBuilder sbWhere = new StringBuilder();
			if( _scope == 3 )
			{
				// can use row id
				sbWhere.Append( "AND BUFFERRESPONSEDATA.BUFFERRESPONSEDATAID = '" );
				sbWhere.Append( _identBufferResponseId );
				sbWhere.Append( "' " );
			}
			else
			{
				// committing a group of data
				// need to calculate what to include in group
				// only commitable statuses
				sbWhere.Append( "AND BUFFERRESPONSEDATA.BUFFERCOMMITSTATUS " );
				// DPH 24/03/2006 - store buffer sql clause in a constant
				sbWhere.Append( BufferBrowser._DISPLAY_DATA_SQL_CLAUSE );
				// study id
				sbWhere.Append( "AND BUFFERRESPONSEDATA.CLINICALTRIALID = " );
				sbWhere.Append( _identStudy );
				// site
				sbWhere.Append( " AND BUFFERRESPONSEDATA.SITE = '" );
				sbWhere.Append( _identSite );
				// subject
				sbWhere.Append( "' AND BUFFERRESPONSEDATA.PERSONID = " );
				sbWhere.Append( _identSubject );
				sbWhere.Append( " " );
				if( _scope > 0 )
				{
					// include visit search criteria
					// visit id
					sbWhere.Append( " AND ( BUFFERRESPONSEDATA.VISITID " );
					if( _identVisitId != "" )
					{
						sbWhere.Append( "= " );
						sbWhere.Append( _identVisitId );
					}
					else
					{
						sbWhere.Append( "IS NULL" );
					}
					// visit cycle
					sbWhere.Append( " ) AND ( BUFFERRESPONSEDATA.VISITCYCLENUMBER " );
					if( _identVisitCycle != "" )
					{
						sbWhere.Append( "= " );
						sbWhere.Append( _identVisitCycle );
					}
					else
					{
						sbWhere.Append( "IS NULL" );
					}
					sbWhere.Append( " ) " );
				}
				if( _scope > 1 )
				{
					// include eform search criteria
					// eform id
					sbWhere.Append( " AND ( BUFFERRESPONSEDATA.CRFPAGEID " );
					if( _identEformId != "" )
					{
						sbWhere.Append( "= " );
						sbWhere.Append( _identEformId );
					}
					else
					{
						sbWhere.Append( "IS NULL" );
					}
					// eform cycle
					sbWhere.Append( " ) AND ( BUFFERRESPONSEDATA.CRFPAGECYCLENUMBER " );
					if( _identEformCycle != "" )
					{
						sbWhere.Append( "= " );
						sbWhere.Append( _identEformCycle );
					}
					else
					{
						sbWhere.Append( "IS NULL" );
					}
					sbWhere.Append( " ) " );
				}
				// just collect known target data
				//sbWhere.Append( " AND ( (BUFFERRESPONSEDATA.VISITID IS NOT NULL) OR (BUFFERRESPONSEDATA.VISITCYCLENUMBER IS NOT NULL) OR (BUFFERRESPONSEDATA.CRFPAGEID IS NOT NULL) OR (BUFFERRESPONSEDATA.CRFPAGECYCLENUMBER IS NOT NULL) ) " );
			}

			return sbWhere.ToString();
		}

		/// <summary>
		/// Create where clause based on BufferDataItemResponse unique row ids
		/// </summary>
		/// <param name="alBufferDataResponseID"></param>
		/// <returns></returns>
		private string GetPostCommitWhereClause()
		{
			log.Info( "Starting GetPostCommitWhereClause" );

			StringBuilder sbWhere = new StringBuilder();
			sbWhere.Append( "AND BUFFERRESPONSEDATA.BUFFERRESPONSEDATAID IN ( " );
			int nCount = 0;
			foreach(BufferResponseItem bufferResponse in _alBufferResponseList)
			{
				if( nCount > 0 )
				{
					sbWhere.Append( " , " );
				}
				sbWhere.Append( "'" );
				sbWhere.Append( bufferResponse.BufferResponseId );
				sbWhere.Append( "'" );
				nCount++;
			}
			sbWhere.Append( ") " );

			return sbWhere.ToString();
		}

		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		private string GetOrderBy()
		{
			log.Info( "Starting GetOrderBy" );

			StringBuilder sbOrderBy = new StringBuilder();

			sbOrderBy.Append( "ORDER BY BUFFERRESPONSEDATA.PERSONID, STUDYVISIT.VISITORDER, BUFFERRESPONSEDATA.VISITID, BUFFERRESPONSEDATA.VISITCYCLENUMBER, CRFPAGE.CRFPAGEORDER, " );
			// DPH 17/03/2006 - Include repeating question group ordering info - 
			sbOrderBy.Append( "BUFFERRESPONSEDATA.CRFPAGEID, BUFFERRESPONSEDATA.CRFPAGECYCLENUMBER, CRFELEMENT.FIELDORDER, BUFFERRESPONSEDATA.RESPONSEREPEATNUMBER, CRFELEMENT.QGROUPFIELDORDER " );

			return sbOrderBy.ToString();
		}

		/// <summary>
		/// Generate a MACRO subject 
		/// TODO: if 2 questions are exactly the same need to generate multiple 'subjects' 
		/// and commit separately - collating the results
		/// </summary>
		/// <param name="dtDataToCommit"></param>
		/// <returns></returns>
		private MACROSubject GetMACROSubject( DataTable dtDataToCommit )
		{
			log.Info( "Starting GetMACROSubject" );

			MACROSubject subject = new MACROSubject();
			MACROSubjectVisit visit = null;

			// store previous visit / eform combination
			int nVisitPrev = BufferBrowser._DEFAULT_MISSING_NUMERIC;
			int nVisitCyclePrev = BufferBrowser._DEFAULT_MISSING_NUMERIC;

			// loop through rows
			for( int i=0; i < dtDataToCommit.Rows.Count; i++ )
			{
				string commitStudyName = dtDataToCommit.Rows[i]["CLINICALTRIALNAME"].ToString();
				int commitStudyId = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["CLINICALTRIALID"] != System.DBNull.Value )
				{
					commitStudyId = Convert.ToInt32( dtDataToCommit.Rows[i]["CLINICALTRIALID"] );
				}
				string commitSite = dtDataToCommit.Rows[i]["SITE"].ToString();
				string commitSubjectLabel = dtDataToCommit.Rows[i]["SUBJECTLABEL"].ToString();
				int commitSubjectId = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["PERSONID"] != System.DBNull.Value )
				{
					commitSubjectId = Convert.ToInt32( dtDataToCommit.Rows[i]["PERSONID"] );
				}

				if ( i == 0 )
				{
					// if 1st row
					// subject detail
					subject.Study = commitStudyName;
					subject.Site = commitSite;
					subject.Label = commitSubjectLabel;
				}

				// visit info
				int commitVisitId = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["VISITID"] != System.DBNull.Value )
				{
					commitVisitId = Convert.ToInt32( dtDataToCommit.Rows[i]["VISITID"] );
				}
				int commitVisitCycle = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["VISITCYCLENUMBER"] != System.DBNull.Value )
				{
					commitVisitCycle = Convert.ToInt32( dtDataToCommit.Rows[i]["VISITCYCLENUMBER"] );
				}
				string commitVisitCode = "";
				if( dtDataToCommit.Rows[i]["VISITCODE"] != System.DBNull.Value )
				{
					commitVisitCode = dtDataToCommit.Rows[i]["VISITCODE"].ToString();
				}

				// check if new visit
				if((nVisitPrev != commitVisitId)||(nVisitCyclePrev != commitVisitCycle))
				{
					// create new visit
					visit = new MACROSubjectVisit();
					// add to subject
					subject.Visit.Add(visit);
					visit.Code = commitVisitCode;
					visit.Cycle = Convert.ToString(commitVisitCycle);
					// set prevs
					nVisitPrev = commitVisitId;
					nVisitCyclePrev = commitVisitCycle;
				}

				// eform info
				int commitEformId = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["CRFPAGEID"] != System.DBNull.Value )
				{
					commitEformId = Convert.ToInt32( dtDataToCommit.Rows[i]["CRFPAGEID"] );
				}
				int commitEformCycle = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["CRFPAGECYCLENUMBER"] != System.DBNull.Value )
				{
					commitEformCycle = Convert.ToInt32( dtDataToCommit.Rows[i]["CRFPAGECYCLENUMBER"] );
				}
				string commitEformCode = "";
				if( dtDataToCommit.Rows[i]["CRFPAGECODE"] != System.DBNull.Value )
				{
					commitEformCode = dtDataToCommit.Rows[i]["CRFPAGECODE"].ToString();
				}

				// check if eform exists
				int nEformPos = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				for(int nEform=0; nEform < visit.Eform.Count; nEform++)
				{
					MACROSubjectVisitEform visEform = (MACROSubjectVisitEform)visit.Eform[nEform];
					if(visEform != null)
					{
						if((visEform.Code == commitEformCode)&&(visEform.Cycle == commitEformCycle.ToString()))
						{
							nEformPos = nEform;
							break;
						}
					}
				}

				// get eform from object (if exists) or create new one
				MACROSubjectVisitEform eform = null;
				if(nEformPos != BufferBrowser._DEFAULT_MISSING_NUMERIC)
				{
					eform = (MACROSubjectVisitEform)visit.Eform[nEformPos];
				}
				if (eform == null)
				{
					eform = new MACROSubjectVisitEform();
					eform.Code = commitEformCode;
					eform.Cycle = Convert.ToString(commitEformCycle);
					visit.Eform.Add(eform);
				}

				// response info
				int commitQuestionId = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["DATAITEMID"] != System.DBNull.Value )
				{
					commitQuestionId = Convert.ToInt32( dtDataToCommit.Rows[i]["DATAITEMID"] );
				}
				int commitQuestionCycle = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["RESPONSEREPEATNUMBER"] != System.DBNull.Value )
				{
					commitQuestionCycle = Convert.ToInt32( dtDataToCommit.Rows[i]["RESPONSEREPEATNUMBER"] );
				}
				string commitQuestionCode = dtDataToCommit.Rows[i]["DATAITEMCODE"].ToString();
				string commitResponseValue = dtDataToCommit.Rows[i]["RESPONSEVALUE"].ToString();
				string commitBufferResponseId = dtDataToCommit.Rows[i]["BUFFERRESPONSEDATAID"].ToString();
				BufferBrowser.BufferCommitStatus eCommitStatus = (BufferBrowser.BufferCommitStatus)Convert.ToInt32( dtDataToCommit.Rows[i]["BUFFERCOMMITSTATUS"] );
				double commitOrderDate = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["ORDERDATETIME"] != System.DBNull.Value )
				{
					commitOrderDate = Convert.ToDouble( dtDataToCommit.Rows[i]["ORDERDATETIME"] );
				}
				string commitBufferParentResponseId = dtDataToCommit.Rows[i]["BUFFERRESPONSEID"].ToString();
				// add response to eform
				MACROSubjectVisitEformQuestion qu = new MACROSubjectVisitEformQuestion();
				qu.Code = commitQuestionCode;
				qu.Cycle = Convert.ToString( commitQuestionCycle );
				qu.Value = commitResponseValue;
				eform.Question.Add(qu);

				// Create BufferResponseItem and store in list
				BufferResponseItem bufferResponse = new BufferResponseItem( commitStudyName, commitStudyId,
						commitSite, commitSubjectLabel, commitSubjectId, commitVisitCode, commitVisitId, commitVisitCycle,
						commitEformCode, commitEformId, commitEformCycle, commitQuestionCode, commitQuestionId, commitQuestionCycle,
						commitResponseValue, commitOrderDate, eCommitStatus, commitBufferResponseId, commitBufferParentResponseId);

				// add Buffer response id to list
				_alBufferResponseList.Add( bufferResponse );
			}

			return subject;
		}

		/// <summary>
		/// Create list of bufferresponses only when discarding data
		/// </summary>
		/// <param name="dtDataToCommit"></param>
		private void CreateBufferResponseList( DataTable dtDataToCommit )
		{
			log.Info( "Starting CreateBufferResponseList" );

			// loop through rows
			for( int i=0; i < dtDataToCommit.Rows.Count; i++ )
			{
				string commitStudyName = dtDataToCommit.Rows[i]["CLINICALTRIALNAME"].ToString();
				int commitStudyId = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["CLINICALTRIALID"] != System.DBNull.Value )
				{
					commitStudyId = Convert.ToInt32( dtDataToCommit.Rows[i]["CLINICALTRIALID"] );
				}
				string commitSite = dtDataToCommit.Rows[i]["SITE"].ToString();
				string commitSubjectLabel = dtDataToCommit.Rows[i]["SUBJECTLABEL"].ToString();
				int commitSubjectId = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["PERSONID"] != System.DBNull.Value )
				{
					commitSubjectId = Convert.ToInt32( dtDataToCommit.Rows[i]["PERSONID"] );
				}

				// visit info
				int commitVisitId = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["VISITID"] != System.DBNull.Value )
				{
					commitVisitId = Convert.ToInt32( dtDataToCommit.Rows[i]["VISITID"] );
				}
				int commitVisitCycle = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["VISITCYCLENUMBER"] != System.DBNull.Value )
				{
					commitVisitCycle = Convert.ToInt32( dtDataToCommit.Rows[i]["VISITCYCLENUMBER"] );
				}
				string commitVisitCode = "";
				if( dtDataToCommit.Rows[i]["VISITCODE"] != System.DBNull.Value )
				{
					commitVisitCode = dtDataToCommit.Rows[i]["VISITCODE"].ToString();
				}

				// eform info
				int commitEformId = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["CRFPAGEID"] != System.DBNull.Value )
				{
					commitEformId = Convert.ToInt32( dtDataToCommit.Rows[i]["CRFPAGEID"] );
				}
				int commitEformCycle = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["CRFPAGECYCLENUMBER"] != System.DBNull.Value )
				{
					commitEformCycle = Convert.ToInt32( dtDataToCommit.Rows[i]["CRFPAGECYCLENUMBER"] );
				}
				string commitEformCode = "";
				if( dtDataToCommit.Rows[i]["CRFPAGECODE"] != System.DBNull.Value )
				{
					commitEformCode = dtDataToCommit.Rows[i]["CRFPAGECODE"].ToString();
				}

				// response info
				int commitQuestionId = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["DATAITEMID"] != System.DBNull.Value )
				{
					commitQuestionId = Convert.ToInt32( dtDataToCommit.Rows[i]["DATAITEMID"] );
				}
				int commitQuestionCycle = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["RESPONSEREPEATNUMBER"] != System.DBNull.Value )
				{
					commitQuestionCycle = Convert.ToInt32( dtDataToCommit.Rows[i]["RESPONSEREPEATNUMBER"] );
				}
				string commitQuestionCode = dtDataToCommit.Rows[i]["DATAITEMCODE"].ToString();
				string commitResponseValue = dtDataToCommit.Rows[i]["RESPONSEVALUE"].ToString();
				string commitBufferResponseId = dtDataToCommit.Rows[i]["BUFFERRESPONSEDATAID"].ToString();
				BufferBrowser.BufferCommitStatus eCommitStatus = (BufferBrowser.BufferCommitStatus)Convert.ToInt32( dtDataToCommit.Rows[i]["BUFFERCOMMITSTATUS"] );
				double commitOrderDate = BufferBrowser._DEFAULT_MISSING_NUMERIC;
				if( dtDataToCommit.Rows[i]["ORDERDATETIME"] != System.DBNull.Value )
				{
					commitOrderDate = Convert.ToDouble( dtDataToCommit.Rows[i]["ORDERDATETIME"] );
				}
				string commitBufferParentResponseId = dtDataToCommit.Rows[i]["BUFFERRESPONSEID"].ToString();

				// Create BufferResponseItem and store in list
				BufferResponseItem bufferResponse = new BufferResponseItem( commitStudyName, commitStudyId,
					commitSite, commitSubjectLabel, commitSubjectId, commitVisitCode, commitVisitId, commitVisitCycle,
					commitEformCode, commitEformId, commitEformCycle, commitQuestionCode, commitQuestionId, commitQuestionCycle,
					commitResponseValue, commitOrderDate, eCommitStatus, commitBufferResponseId, commitBufferParentResponseId);

				// add Buffer response id to list
				_alBufferResponseList.Add( bufferResponse );
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="subject"></param>
		/// <returns></returns>
		private string SubjectToString(MACROSubject subject)
		{
			log.Info( "Starting SubjectToString" );

			log.Info( "XmlSerializer x = new XmlSerializer(typeof(MACROSubject));" );
			XmlSerializer x = new XmlSerializer(typeof(MACROSubject));
			log.Info( "MemoryStream stream = new MemoryStream();" );
			MemoryStream stream = new MemoryStream();
			log.Info( "TextWriter writer = new StreamWriter(stream);" );
			TextWriter writer = new StreamWriter(stream);
			log.Info( "x.Serialize(writer, subject);" );
			x.Serialize(writer, subject);
			log.Info( "writer.Close();" );
			writer.Close();
			log.Info( "string xml=new System.Text.UTF8Encoding().GetString(stream.ToArray());" );
			string xml=new System.Text.UTF8Encoding().GetString(stream.ToArray());
			log.Info( "stream.Close();" );
			stream.Close();
			log.Debug("SubjectToString = " + xml);
			return xml;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="bufferUser"></param>
		/// <param name="sXmlToCommit"></param>
		/// <param name="sMACROCommitResults"></param>
		/// <returns></returns>
		private bool CommitToMACRO( BufferMACROUser bufferUser, string sXmlToCommit, ref string sMACROCommitResults)
		{
			bool bCommit = false;

			// collect serialised user so no need to log into API
			bool bNoDbChosen = false;
			string sSerialisedUser = bufferUser.MACROUser.GetStateHex( ref bNoDbChosen );

			// log
			log.Info( "Starting CommitToMACRO" );

			// commit data
			log.Debug("sXmlToCommit = " + sXmlToCommit );

			API.DataInputResult inputResult = API.InputXMLSubjectData( sSerialisedUser, sXmlToCommit, ref sMACROCommitResults );
			log.Debug("Committing data to MACRO through COM API - " + Convert.ToInt16(inputResult));

			// if a success
			if (inputResult==API.DataInputResult.Success )
			{
				bCommit = true;
			} 
			else 
			{
				log.Warn(sMACROCommitResults);
			}

			return bCommit;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sXmlCommitMsg"></param>
		private void ProcessAPIResults(string sXmlCommitMsg)
		{
			// log
			log.Info( "Starting ProcessAPIResults" );

			try
			{
				if(sXmlCommitMsg != "")
				{
					// load xml string
					XmlDocument commitStatusXML = new XmlDocument();
					commitStatusXML.LoadXml(sXmlCommitMsg);

					// loop through DataErrMsg
					foreach ( XmlNode nodedataErrMsg in commitStatusXML.SelectNodes( "//MACROInputErrors/DataErrMsg" ) )
					{
						// get detail from node
						// msgtype
						BufferBrowser.BufferCommitStatus eCommitStatus = (BufferBrowser.BufferCommitStatus)Convert.ToInt16(nodedataErrMsg.Attributes["MsgType"].Value);
						// description
						string sMsgDesc = nodedataErrMsg.Attributes["MsgDesc"].Value.ToString();
						// log errors
						log.Debug( sMsgDesc );

						// parse description
						if((eCommitStatus!=BufferBrowser.BufferCommitStatus.InvalidXML)
								&&(eCommitStatus!=BufferBrowser.BufferCommitStatus.Success))
						{
							// DPH 22/03/2006 - convert a warning level API status to a success
							// for saving to database purposes - but want to collect reject message
							// NotCommitted - equivalent to ValueWarning in API enum eDataInputError
							if( eCommitStatus == BufferBrowser.BufferCommitStatus.NotCommitted )
							{
								eCommitStatus = BufferBrowser.BufferCommitStatus.Success;
							}
							// read to first space
							int nSpace = sMsgDesc.IndexOf(" ");
							if(nSpace > -1)
							{
								string sDetail = sMsgDesc.Substring(0, nSpace);
								string sFailWarnRejectMessage = "";
								if( ( nSpace + 1 ) < sMsgDesc.Length )
								{
									sFailWarnRejectMessage = sMsgDesc.Substring( (nSpace + 1),  ( sMsgDesc.Length - (nSpace + 1) ) );
								}
								// extract - should be in format visit[cycle]:eform[cycle]:question[cycle]
								string sMainSep=":";
								string[] asDetail = sDetail.Split(sMainSep.ToCharArray());
								// DPH 20/03/2006 - Look out for whole subject error
								bool bSubjectError = false;
								string sVisitCode = "";
								int nVisitCycle = BufferBrowser._DEFAULT_MISSING_NUMERIC;
								string sEFormCode = "";
								int nEFormCycle = BufferBrowser._DEFAULT_MISSING_NUMERIC;
								string sQuestionCode = "";
								int nQuestionCycle = BufferBrowser._DEFAULT_MISSING_NUMERIC;
								// get visit/eform/question detail
								switch(asDetail.Length)
								{
									case 1:
									{
										// DPH 20/03/2006 - subject or visit
										// visit
										GetCodeAndCycle(asDetail[0], ref sVisitCode, ref nVisitCycle);
										// if visit = "Subject" & cycle = -1 i.e. subject level error
										if( (sVisitCode.ToLower() == "subject") && ( nVisitCycle == BufferBrowser._DEFAULT_MISSING_NUMERIC ) )
										{
											bSubjectError = true;
										}
										break;
									}
									case 2:
									{
										// visit
										GetCodeAndCycle(asDetail[0], ref sVisitCode, ref nVisitCycle);
										// eform
										GetCodeAndCycle(asDetail[1], ref sEFormCode, ref nEFormCycle);
										break;
									}
									case 3:
									{
										// visit
										GetCodeAndCycle(asDetail[0], ref sVisitCode, ref nVisitCycle);
										// eform
										GetCodeAndCycle(asDetail[1], ref sEFormCode, ref nEFormCycle);
										//question
										GetCodeAndCycle(asDetail[2], ref sQuestionCode, ref nQuestionCycle);
										break;
									}
								}
								// loop through Buffer Responses items
								for( int i=0; i < _alBufferResponseList.Count; i++)
								{
									BufferResponseItem bufferResponse = (BufferResponseItem)_alBufferResponseList[i];
									if( ! bSubjectError )
									{
										// match response
										if((sVisitCode == bufferResponse.VisitCode)&&(nVisitCycle == bufferResponse.VisitCycle))
										{
											if(sEFormCode != "")
											{
												if((sEFormCode == bufferResponse.EformCode)&&(nEFormCycle == bufferResponse.EformCycle))
												{
													if(sQuestionCode != "")
													{
														if((sQuestionCode == bufferResponse.QuestionCode)&&(nQuestionCycle == bufferResponse.QuestionCycle))
														{
															// set status
															bufferResponse.BufferCommitStatus = eCommitStatus;
															// set warning message
															bufferResponse.FailRejectWarnMessage = sFailWarnRejectMessage;
														}
													}
													else
													{
														// set status
														bufferResponse.BufferCommitStatus = eCommitStatus;
														// set warning message
														bufferResponse.FailRejectWarnMessage = sFailWarnRejectMessage;
													}
												}
											}
											else
											{
												// set status of response
												bufferResponse.BufferCommitStatus = eCommitStatus;
												// set warning message
												bufferResponse.FailRejectWarnMessage = sFailWarnRejectMessage;
											}
										}
									}
									else
									{
										// DPH 20/03/2006 - Subject error
										// set status of response
										bufferResponse.BufferCommitStatus = eCommitStatus;
										// set warning message
										bufferResponse.FailRejectWarnMessage = "Subject " + sFailWarnRejectMessage;
									}
								}
							}
						}
					}
				}
			}
			catch(Exception ex)
			{
				log.Error( "Error reading commit response XML.", ex );
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sCodeAndCycle"></param>
		/// <param name="sCode"></param>
		/// <param name="nCycle"></param>
		private static void GetCodeAndCycle(string sCodeAndCycle, ref string sCode, ref int nCycle)
		{
			string sSepStart = "[";
			string sSepEnd = "]";
			string[] asCodeCycle = sCodeAndCycle.Split(sSepStart.ToCharArray());
			if( asCodeCycle.Length == 2)
			{
				// get code 
				sCode = asCodeCycle[0];
				// get cycle
				string sCycle = (asCodeCycle[1].Split(sSepEnd.ToCharArray()))[0];
				try
				{
					nCycle = Convert.ToInt16(sCycle);
				}
				catch
				{}
			}
			else
			{
				// get code 
				sCode = asCodeCycle[0];
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="eCommitStatus"></param>
		private void SetAllResponseStatuses(BufferBrowser.BufferCommitStatus eCommitStatus)
		{
			// log
			log.Info( "Starting SetAllResponseStatuses" );
			
			// loop through responses
			foreach( BufferResponseItem respItem in _alBufferResponseList )
			{
				respItem.BufferCommitStatus = eCommitStatus;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="bufferUser"></param>
		// REVISIONS
		// DPH 21/11/2007 - Added LocalNumToStandard use for dates on non-uk systems
		private void UpdateBufferResponse( BufferMACROUser bufferUser )
		{
			// log
			log.Info( "Starting UpdateBufferResponse" );
			
			// collect current user detail
			string sUser = bufferUser.MACROUser.UserName;
			string sUsernameFull = bufferUser.MACROUser.UserNameFull;
			// database connection
			string sDbConnection = bufferUser.MACROUser.CurrentDBConString;

			// timestamp
			double dblTimestamp = DateTime.Now.ToOADate();
			// timezone
			int nTimezone = BufferBrowser.GetLocalMACROTimezone();

			try
			{
				// loop through BufferResponseItem list
				foreach( BufferResponseItem bufferResponseItem in _alBufferResponseList )
				{
					// create update sql
					// need to update 
					// BUFFERCOMMITSTATUS, BUFFERCOMMITSTATUSTIMESTAMP, BUFFERCOMMITSTATUSTIMESTAMP_TZ
					// USERNAME, USERNAMEFULL
					StringBuilder sbUpdateSql = new StringBuilder();
					// BUFFERCOMMITSTATUS
					sbUpdateSql.Append( "UPDATE BUFFERRESPONSEDATA SET BUFFERCOMMITSTATUS = " );
					sbUpdateSql.Append( Convert.ToInt32( bufferResponseItem.BufferCommitStatus ) );
					// BUFFERCOMMITSTATUSTIMESTAMP
					sbUpdateSql.Append( ", BUFFERCOMMITSTATUSTIMESTAMP = " );
					// DPH 21/11/2007 - LocalNumToStandard double date for SQL
					sbUpdateSql.Append( IMEDFunctions20.LocalNumToStandard(dblTimestamp.ToString(), false) );
					// BUFFERCOMMITSTATUSTIMESTAMP_TZ
					sbUpdateSql.Append( ", BUFFERCOMMITSTATUSTIMESTAMP_TZ = " );
					sbUpdateSql.Append( nTimezone );
					// USERNAME
					sbUpdateSql.Append( ", USERNAME = '" );
					sbUpdateSql.Append( sUser );
					// USERNAMEFULL
					sbUpdateSql.Append( "', USERNAMEFULL = '" );
					sbUpdateSql.Append( sUsernameFull );
					sbUpdateSql.Append( "' " );
					// where
					sbUpdateSql.Append( "WHERE BUFFERRESPONSEDATAID = '" );
					sbUpdateSql.Append( bufferResponseItem.BufferResponseId );
					sbUpdateSql.Append( "'" );

					// TODO: create insert sql
					StringBuilder sbInsertSql = new StringBuilder();
					sbInsertSql.Append( "INSERT INTO BUFFERRESPONSEDATAHISTORY (BUFFERRESPONSEDATAID, BUFFERRESPONSEID, BUFFERCOMMITSTATUS, BUFFERCOMMITSTATUSTIMESTAMP, BUFFERCOMMITSTATUSTIMESTAMP_TZ, ");
					sbInsertSql.Append("CLINICALTRIALNAME, CLINICALTRIALID, SITE, SUBJECTLABEL, PERSONID, VISITCODE, VISITID, VISITCYCLENUMBER, CRFPAGECODE, CRFPAGEID, CRFPAGECYCLENUMBER, ");
					sbInsertSql.Append("DATAITEMCODE, DATAITEMID, RESPONSEREPEATNUMBER, RESPONSEVALUE, ORDERDATETIME, USERNAME, USERNAMEFULL) VALUES ('");
					//BUFFERRESPONSEDATAID
					sbInsertSql.Append( bufferResponseItem.BufferResponseId );
					sbInsertSql.Append("','");
					//BUFFERRESPONSEID - link to complete buffermessage
					sbInsertSql.Append( bufferResponseItem.BufferParentId );
					sbInsertSql.Append("',");
					//BUFFERCOMMITSTATUS
					sbInsertSql.Append( Convert.ToInt32( bufferResponseItem.BufferCommitStatus ) );
					sbInsertSql.Append(",");
					//BUFFERCOMMITSTATUSTIMESTAMP
					// DPH 21/11/2007 - LocalNumToStandard double date for SQL
					sbInsertSql.Append( IMEDFunctions20.LocalNumToStandard( dblTimestamp.ToString(), false ) );
					sbInsertSql.Append(",");
					//BUFFERCOMMITSTATUSTIMESTAMP_TZ
					sbInsertSql.Append( nTimezone );
					sbInsertSql.Append(",'");
					//CLINICALTRIALNAME
					sbInsertSql.Append( bufferResponseItem.StudyName );
					sbInsertSql.Append("',");
					//CLINICALTRIALID
					sbInsertSql.Append( bufferResponseItem.StudyId );
					sbInsertSql.Append(",'");
					//SITE
					sbInsertSql.Append( bufferResponseItem.Site );
					sbInsertSql.Append("','");
					//SUBJECTLABEL
					sbInsertSql.Append( bufferResponseItem.SubjectLabel );
					sbInsertSql.Append("',");
					//PERSONID
					sbInsertSql.Append( bufferResponseItem.SubjectId );
					sbInsertSql.Append(",");
					//VISITCODE
					sbInsertSql.Append("'");
					sbInsertSql.Append( bufferResponseItem.VisitCode );
					sbInsertSql.Append("'");
					sbInsertSql.Append(",");
					//VISITID
					sbInsertSql.Append( bufferResponseItem.VisitId );
					sbInsertSql.Append(",");
					//VISITCYCLENUMBER
					sbInsertSql.Append( bufferResponseItem.VisitCycle );
					sbInsertSql.Append(",");
					//CRFPAGECODE
					sbInsertSql.Append("'");
					sbInsertSql.Append( bufferResponseItem.EformCode );
					sbInsertSql.Append("'");
					sbInsertSql.Append(",");
					//CRFPAGEID
					sbInsertSql.Append( bufferResponseItem.EformId );
					sbInsertSql.Append(",");
					//CRFPAGECYCLENUMBER
					sbInsertSql.Append( bufferResponseItem.EformCycle );
					sbInsertSql.Append(",");
					//DATAITEMCODE
					sbInsertSql.Append("'");
					sbInsertSql.Append( bufferResponseItem.QuestionCode );
					sbInsertSql.Append("'");
					sbInsertSql.Append(",");
					//DATAITEMID
					sbInsertSql.Append( bufferResponseItem.QuestionId );
					sbInsertSql.Append(",");
					//RESPONSEREPEATNUMBER
					sbInsertSql.Append( bufferResponseItem.QuestionCycle );
					sbInsertSql.Append(",");
					//RESPONSEVALUE
					sbInsertSql.Append("'");
					sbInsertSql.Append( bufferResponseItem.ResponseValue );
					sbInsertSql.Append("'");
					sbInsertSql.Append(",");
					// ORDERDATETIME
					if( bufferResponseItem.OrderDate != BufferBrowser._DEFAULT_MISSING_NUMERIC)
					{
						// DPH 21/11/2007 - LocalNumToStandard double date for SQL
						sbInsertSql.Append( IMEDFunctions20.LocalNumToStandard ( bufferResponseItem.OrderDate.ToString(), false ) );
					}
					else
					{
						sbInsertSql.Append( "null" );
					}
					sbInsertSql.Append(",");
					// USERNAME
					sbInsertSql.Append("'");
					sbInsertSql.Append( sUser );
					// USERNAMEFULL
					sbInsertSql.Append( "', '" );
					sbInsertSql.Append( sUsernameFull );
					sbInsertSql.Append( "')" );

					// write to BufferResponseData table
					// loop through responses to save and store on db
					// declare connection / datareader interface
					IDbConnection dbConn = null;
					IDbCommand dbCommand = null;

					// open ImedDataAccess class
					IMEDDataAccess imedData = new IMEDDataAccess();
					IMEDDataAccess.ConnectionType imedConnType;

					// get connection type
					imedConnType = IMEDDataAccess.CalculateConnectionType( sDbConnection );

					// get connection to db
					dbConn = imedData.GetConnection( imedConnType, sDbConnection );
					
					// open connection
					dbConn.Open();

					// get command object 
					dbCommand = imedData.GetCommand(dbConn);

					// set command sql - Update BufferResponseData
					dbCommand.CommandText = sbUpdateSql.ToString();
					// log
					log.Debug(sbUpdateSql.ToString());
					// execute command - BufferResponseData
					dbCommand.ExecuteNonQuery();

					// set command sql - Insert BufferResponseDataHistory
					dbCommand.CommandText = sbInsertSql.ToString();
					// log
					log.Debug(sbInsertSql.ToString());
					// execute command - BufferResponseDataHistory
					dbCommand.ExecuteNonQuery();

					// dispose command
					dbCommand.Dispose();

					// close connection
					dbConn.Close();

					// dispose connection
					dbConn.Dispose();
				}
			}
			catch(Exception ex)
			{
				// store error in log
				log.Error("Error saving to BufferResponseData table.", ex);
				// throw again
				throw ( new Exception( "Error saving to BufferResponsedata table.", ex) );
			}
		}

	}

	#region BufferResponseItem
	/// <summary>
	/// BufferResponseItem object - store identifiable features of a Buffer Response
	/// </summary>
	class BufferResponseItem
	{
		/// <summary>
		/// 
		/// </summary>
		/// <param name="buffStudyName"></param>
		/// <param name="buffStudyId"></param>
		/// <param name="buffSite"></param>
		/// <param name="buffSubject"></param>
		/// <param name="buffSubjectId"></param>
		/// <param name="buffVisitCode"></param>
		/// <param name="buffVisitId"></param>
		/// <param name="buffVisitCycle"></param>
		/// <param name="buffEformCode"></param>
		/// <param name="buffEformId"></param>
		/// <param name="buffEformCycle"></param>
		/// <param name="buffQuestionCode"></param>
		/// <param name="buffQuestionId"></param>
		/// <param name="buffQuestionCycle"></param>
		/// <param name="buffResponseValue"></param>
		/// <param name="buffOrderDate"></param>
		/// <param name="buffCommitStatus"></param>
		/// <param name="buffBufferResponseId"></param>
		/// <param name="buffBufferParentId"></param>
		public BufferResponseItem(string buffStudyName, int buffStudyId, string buffSite, string buffSubject, 
			int buffSubjectId, string buffVisitCode, int buffVisitId,
			int buffVisitCycle, string buffEformCode, int buffEformId, int buffEformCycle, 
			string buffQuestionCode, int buffQuestionId, int buffQuestionCycle, 
			string buffResponseValue, double buffOrderDate, BufferBrowser.BufferCommitStatus buffCommitStatus,
			string buffBufferResponseId, string buffBufferParentId )
		{
			_buffStudyName = buffStudyName;
			_buffStudyId = buffStudyId;
			_buffSite = buffSite;
			_buffSubject = buffSubject;
			_buffSubjectId = buffSubjectId;
			_buffVisitCode = buffVisitCode;
			_buffVisitId = buffVisitId;
			_buffVisitCycle = buffVisitCycle;
			_buffEformCode = buffEformCode;
			_buffEformId = buffEformId;
			_buffEformCycle = buffEformCycle;
			_buffQuestionCode = buffQuestionCode;
			_buffQuestionId = buffQuestionId;
			_buffQuestionCycle = buffQuestionCycle;
			_buffResponseValue = buffResponseValue;
			_buffOrderDate = buffOrderDate;
			_buffBufferResponseId = buffBufferResponseId;
			_buffBufferParentId = buffBufferParentId;
			_buffCommitStatus = buffCommitStatus;
			_buffFailRejectWarnMessage = "";
		}

		// private members
		private string _buffStudyName;
		private int _buffStudyId;
		private string _buffSite;
		private string _buffSubject;
		private int _buffSubjectId;
		private string _buffVisitCode;
		private int _buffVisitId;
		private int _buffVisitCycle;
		private string _buffEformCode;
		private int _buffEformId;
		private int _buffEformCycle;
		private string _buffQuestionCode;
		private int _buffQuestionId;
		private int _buffQuestionCycle;
		private string _buffResponseValue;
		private double _buffOrderDate;
		private string _buffBufferResponseId;
		private string _buffBufferParentId;
		private BufferBrowser.BufferCommitStatus _buffCommitStatus;
		private string _buffFailRejectWarnMessage;

		public string StudyName
		{
			get
			{
				return _buffStudyName;
			}
		}

		public int StudyId
		{
			get
			{
				return _buffStudyId;
			}
		}

		public string Site
		{
			get
			{
				return _buffSite;
			}
		}

		public string SubjectLabel
		{
			get
			{
				return _buffSubject;
			}
		}

		public int SubjectId
		{
			get
			{
				return _buffSubjectId;
			}
		}

		public string VisitCode
		{
			get
			{
				return _buffVisitCode;
			}
		}

		public int VisitId
		{
			get
			{
				return _buffVisitId;
			}
		}

		public int VisitCycle
		{
			get
			{
				return _buffVisitCycle;
			}
		}

		public string EformCode
		{
			get
			{
				return _buffEformCode;
			}
		}

		public int EformId
		{
			get
			{
				return _buffEformId;
			}
		}

		public int EformCycle
		{
			get
			{
				return _buffEformCycle;
			}
		}

		public string QuestionCode
		{
			get
			{
				return _buffQuestionCode;
			}
		}

		public int QuestionId
		{
			get
			{
				return _buffQuestionId;
			}
		}

		public int QuestionCycle
		{
			get
			{
				return _buffQuestionCycle;
			}
		}

		public string ResponseValue
		{
			get
			{
				return _buffResponseValue;
			}
		}

		public double OrderDate
		{
			get
			{
				return _buffOrderDate;
			}
		}

		public string BufferResponseId
		{
			get
			{
				return _buffBufferResponseId;
			}
		}

		public string BufferParentId
		{
			get
			{
				return _buffBufferParentId;
			}
		}

		public BufferBrowser.BufferCommitStatus BufferCommitStatus
		{
			get
			{
				return _buffCommitStatus;
			}
			set
			{
				_buffCommitStatus = value;
			}
		}

		public string FailRejectWarnMessage
		{
			get
			{
				return _buffFailRejectWarnMessage;
			}
			set
			{
				_buffFailRejectWarnMessage = value;
			}
		}
	}
	#endregion
}
