using System;
using System.Text;
using System.Collections;
using System.Data;
using InferMed.Components;
using log4net;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for BufferTargetSelectionSave.
	/// </summary>
	class BufferTargetSelectionSave
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
		private string _identVisitCode;
		private string _identEformId;
		private string _identEformCycle;
		private string _identEformCode;
		private string _identEformTaskId;
		private string _identEformElementId;
		private string _identResponseCycle;
		private string _identResponseTaskId;
		// buffer response identifier
		private string _identBufferResponseId;
		// back link detail
		private string _backlinkStudyId;
		private string _backlinkSite;
		private string _backlinkSubjectId;

		// log4net
		private static readonly ILog log = LogManager.GetLogger( typeof(BufferDataBrowserSave) );

		/// <summary>
		/// 
		/// </summary>
		/// <param name="formData"></param>
		public BufferTargetSelectionSave(string formData)
		{
			log.Debug( "formData = " + formData );
			
			// initialize member variables
			_identStudy = "";
			_identSite = "";
			_identSubject = "";
			_identVisitId = "";
			_identVisitCycle = "";
			_identVisitCode = "";
			_identEformId = "";
			_identEformCycle = "";
			_identEformCode = "";
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
									// visit code
									_identVisitCode = alIdentify[ i ].ToString();
									break;
								}
								case 6:
								{
									// eform id
									_identEformId = alIdentify[ i ].ToString();
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
									// eform cycle
									_identEformCycle = alIdentify[ i ].ToString();
									break;
								}
								case 9:
								{
									// eform code
									_identEformCode = alIdentify[ i ].ToString();
									break;
								}
								case 10:
								{
									// eform element id
									_identEformElementId = alIdentify[ i ].ToString();
									break;
								}
								case 11:
								{
									// response cycle
									_identResponseCycle = alIdentify[ i ].ToString();
									break;
								}
								case 12:
								{
									// response task id
									_identResponseTaskId = alIdentify[ i ].ToString();
									break;
								}
								case 13:
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
						// scope of save - not used
						string sSaveScope = BufferBrowser.ReplaceHtmlCharacters( alIndivData[1].ToString() );
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
		public string SaveBufferTarget( BufferMACROUser bufferUser )
		{
			log.Info( "Starting SaveBufferTarget" );

			// save buffer target changes
			CommitSave ( bufferUser );

			StringBuilder sbPageHtml = new StringBuilder();
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

			return sbPageHtml.ToString();
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="bufferUser"></param>
		// REVISIONS
		// DPH 21/11/2007 - Added LocalNumToStandard use for dates on non-uk systems
		private void CommitSave( BufferMACROUser bufferUser )
		{
			try
			{
				log.Info( "Starting CommitSave" );

				// database connection
				string sDbConnection = bufferUser.MACROUser.CurrentDBConString;

				// save passed detail to buffer response table - sql update
				StringBuilder sbSql = new StringBuilder();

				// timestamp + timezone info
				double dblTimeStamp = DateTime.Now.ToOADate();
				int nTimeZone = BufferBrowser.GetLocalMACROTimezone();

				sbSql.Append( "UPDATE BUFFERRESPONSEDATA SET BUFFERCOMMITSTATUS = " );
				sbSql.Append( Convert.ToInt32( BufferBrowser.BufferCommitStatus.NotCommitted ) );
				sbSql.Append( ", BUFFERCOMMITSTATUSTIMESTAMP = " );
				// DPH 21/11/2007 - LocalNumToStandard double date for SQL
				sbSql.Append( IMEDFunctions20.LocalNumToStandard(dblTimeStamp.ToString(), false) );
				sbSql.Append( ", BUFFERCOMMITSTATUSTIMESTAMP_TZ = " );
				sbSql.Append( nTimeZone );
				sbSql.Append( ", VISITCODE = '" );
				sbSql.Append( _identVisitCode );
				sbSql.Append( "', VISITID = " );
				sbSql.Append( _identVisitId );
				sbSql.Append( ", VISITCYCLENUMBER = " );
				sbSql.Append( _identVisitCycle );
				sbSql.Append( ", CRFPAGECODE = '" );
				sbSql.Append( _identEformCode );
				sbSql.Append( "', CRFPAGEID = " );
				sbSql.Append( _identEformId );
				sbSql.Append( ", CRFPAGECYCLENUMBER = " );
				sbSql.Append( _identEformCycle );
				sbSql.Append( ", RESPONSEREPEATNUMBER = " );
				sbSql.Append( _identResponseCycle );
				sbSql.Append( ", USERNAME = '" );
				sbSql.Append( bufferUser.MACROUser.UserName );
				sbSql.Append( "', USERNAMEFULL = '" );
				sbSql.Append( bufferUser.MACROUser.UserNameFull );
				sbSql.Append( "' WHERE BUFFERRESPONSEDATAID = '" );
				sbSql.Append( _identBufferResponseId );
				sbSql.Append( "'" );

				// write to BufferResponseData table
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
				dbCommand.CommandText = sbSql.ToString();
				// log
				log.Debug(sbSql.ToString());

				// execute command - BufferResponseData
				dbCommand.ExecuteNonQuery();

				// copy row to bufferresponsehistory
				StringBuilder sbHistorySql = new StringBuilder();
				sbHistorySql.Append( "INSERT INTO BUFFERRESPONSEDATAHISTORY (BUFFERRESPONSEDATAID, BUFFERRESPONSEID, BUFFERCOMMITSTATUS, BUFFERCOMMITSTATUSTIMESTAMP, BUFFERCOMMITSTATUSTIMESTAMP_TZ, ");
				sbHistorySql.Append("CLINICALTRIALNAME, CLINICALTRIALID, SITE, SUBJECTLABEL, PERSONID, VISITCODE, VISITID, VISITCYCLENUMBER, CRFPAGECODE, CRFPAGEID, CRFPAGECYCLENUMBER, ");
				sbHistorySql.Append("DATAITEMCODE, DATAITEMID, RESPONSEREPEATNUMBER, RESPONSEVALUE, ORDERDATETIME, USERNAME, USERNAMEFULL) ");
				// select from
				sbHistorySql.Append( "SELECT BUFFERRESPONSEDATAID, BUFFERRESPONSEID, BUFFERCOMMITSTATUS, BUFFERCOMMITSTATUSTIMESTAMP, BUFFERCOMMITSTATUSTIMESTAMP_TZ, ");
				sbHistorySql.Append("CLINICALTRIALNAME, CLINICALTRIALID, SITE, SUBJECTLABEL, PERSONID, VISITCODE, VISITID, VISITCYCLENUMBER, CRFPAGECODE, CRFPAGEID, CRFPAGECYCLENUMBER, ");
				sbHistorySql.Append("DATAITEMCODE, DATAITEMID, RESPONSEREPEATNUMBER, RESPONSEVALUE, ORDERDATETIME, USERNAME, USERNAMEFULL FROM BUFFERRESPONSEDATA WHERE BUFFERRESPONSEDATAID = '");
				sbHistorySql.Append( _identBufferResponseId );
				sbHistorySql.Append( "'" );

				// set command sql - Insert BufferResponseDataHistory
				dbCommand.CommandText = sbHistorySql.ToString();
				// log
				log.Debug(sbHistorySql.ToString());

				// execute command - BufferResponseDataHistory
				dbCommand.ExecuteNonQuery();

				// dispose command
				dbCommand.Dispose();

				// close connection
				dbConn.Close();

				// dispose connection
				dbConn.Dispose();
			}
			catch(Exception ex)
			{
				log.Error( "Error committing target change to database.", ex );
				throw ( new Exception( ex.Message ) );
			}
		}
	}
}
