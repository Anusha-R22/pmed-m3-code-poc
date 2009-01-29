using System;
using System.Text;
using log4net;
using InferMed.Components;
using System.Data;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Render Buffer Summary page
	/// </summary>
	class BufferSummary
	{
		private BufferSummary()
		{}

		// log4net
		private static readonly ILog log = LogManager.GetLogger( typeof(BufferSummary) );

		/// <summary>
		/// Render buffer summary page
		/// </summary>
		/// <param name="bufferMACROUser">MACRO user object</param>
		/// <param name="studyId">study id</param>
		/// <param name="site">site code</param>
		/// <param name="subjectNo">subject no</param>
		/// <returns>buffer summary page html</returns>
		public static string RenderBufferPage(BufferMACROUser bufferMACROUser, int studyId, string site, int subjectNo)
		{
			StringBuilder sbBufferPage = new StringBuilder();
			string sSubjectLabel = "";
			bool bPatientAware = false;
			bool bAllStudyRendered = false;

			log.Info( "Starting RenderBufferPage" );

			// check if in patient aware mode
			if( (studyId != BufferBrowser._DEFAULT_MISSING_NUMERIC) && (site != "") && (subjectNo != BufferBrowser._DEFAULT_MISSING_NUMERIC) )
			{
				// set flag
				bPatientAware = true;
				// collect subject label
				// DPH 07/12/2006 - get label dependent upon permission 
				sSubjectLabel = SubjectLabel( bufferMACROUser, studyId, site, subjectNo, true );
			}

			// form title
			string sPageTitle = "";
			// if patient aware include subject label
			if( bPatientAware )
			{
				sPageTitle += "Patient Aware Mode - ";
				if( sSubjectLabel != "" )
				{
					sPageTitle += sSubjectLabel;
				}
				else
				{
					sPageTitle += "Unknown Subject";
				}
			}

			// render header
			sbBufferPage.Append( RenderHeader(sPageTitle) );

			// calculate for logged in user - study / all
			DataTable dtStudies = GetStudyData( bufferMACROUser );

			// all studies total
			long lAllStudiesData = GetAllStudiesDataCount( dtStudies );

			// if patient aware mode calculate records for subject
			// DPH 07/12/2006 - only display section if know subject identity (label)
			if( ( bPatientAware ) && (sSubjectLabel != "") )
			{
				// get subject data
				long lSubjectData = GetSubjectBufferDataCount( bufferMACROUser, studyId, site, subjectNo );

				// render subject sections
				sbBufferPage.Append( "<tr>" );

				// subject
				sbBufferPage.Append( RenderSubjectSection(sSubjectLabel, lSubjectData, studyId, site, subjectNo) );

				// all
				sbBufferPage.Append( RenderSection( BufferBrowser._DEFAULT_MISSING_NUMERIC, lAllStudiesData, "All data" ) );
				bAllStudyRendered = true;

				sbBufferPage.Append( "</tr>" );
				sbBufferPage.Append( "<tr><td height=\"50px\" colspan=2></td></tr>" );
			}
			
			if( dtStudies.Rows.Count > 0 )
			{
				// loop through studies
				// if already rendered "all" do not render again else render next to first study
				foreach(DataRow drStudy in dtStudies.Rows)
				{
					// render subject sections
					sbBufferPage.Append( "<tr>" );

					// collect sum of studies 
					int lStudyId = Convert.ToInt32( drStudy["CLINICALTRIALID"] );
					long lStudyCount = Convert.ToInt64( drStudy["DATACOUNT"] );
					string sStudyName = drStudy["CLINICALTRIALNAME"].ToString();

					// render study
					sbBufferPage.Append( RenderSection( lStudyId, lStudyCount, sStudyName + " Study" ) );

					// check if need to render all section
					if(!bAllStudyRendered)
					{
						// all
						sbBufferPage.Append( RenderSection( BufferBrowser._DEFAULT_MISSING_NUMERIC, lAllStudiesData, "All data" ) );
						bAllStudyRendered = true;
					}
					else
					{
						// complete row
						sbBufferPage.Append( "<td>&nbsp;</td>" );
					}

					sbBufferPage.Append( "</tr>" );
					// separator
					sbBufferPage.Append( "<tr><td height=\"50px\" colspan=2></td></tr>" );
				}
			}
			else
			{
				// no study data
				// if not patient aware mode then need to display a no data message on the page
				if( !bPatientAware )
				{
					sbBufferPage.Append( "<tr>" );
					// all
					sbBufferPage.Append( RenderSection( BufferBrowser._DEFAULT_MISSING_NUMERIC, lAllStudiesData, "All data" ) );
					sbBufferPage.Append( "</tr>" );
				}
			}

			// close page
			sbBufferPage.Append( "</table></body>" );

			// return page
			return sbBufferPage.ToString();
		}

		/// <summary>
		/// retrieve Subject label from database
		/// </summary>
		/// <param name="bufferMACROUser"></param>
		/// <param name="studyId"></param>
		/// <param name="site"></param>
		/// <param name="subjectNo"></param>
		/// <returns></returns>
		public static string SubjectLabel(BufferMACROUser bufferMACROUser, int studyId, string site, int subjectNo)
		{
			StringBuilder sbSubjectLabel = new StringBuilder();

			// get connection string
			string sDbConn = bufferMACROUser.MACROUser.CurrentDBConString;

			// extract subject label
			StringBuilder sbSql = new StringBuilder();
			sbSql.Append("SELECT TRIALSUBJECT.LOCALIDENTIFIER1, CLINICALTRIAL.CLINICALTRIALNAME ");
			sbSql.Append( "FROM TRIALSUBJECT, CLINICALTRIAL " );
			sbSql.Append( "WHERE TRIALSUBJECT.CLINICALTRIALID = CLINICALTRIAL.CLINICALTRIALID " );
			sbSql.Append( "AND TRIALSUBJECT.CLINICALTRIALID = " );
			sbSql.Append(studyId);
			sbSql.Append(" AND TRIALSUBJECT.TRIALSITE = '");
			sbSql.Append(site);
			sbSql.Append("' AND TRIALSUBJECT.PERSONID = ");
			sbSql.Append(subjectNo);

			// log sql
			log.Debug ( "SubjectLabel sql=" + sbSql.ToString() );

			// get data table
			DataTable dataTable = GetDataTable(sDbConn, sbSql.ToString());

			// should only be 1 matching row - else searching for a non-existant subject
			if( dataTable.Rows.Count == 1)
			{
				// collect subject label (if is one)
				try
				{
					// study
					if( dataTable.Rows[0]["CLINICALTRIALNAME"] != System.DBNull.Value )
					{
						sbSubjectLabel.Append( dataTable.Rows[0]["CLINICALTRIALNAME"].ToString() );
						sbSubjectLabel.Append( @"\" );
					}
					// site
					sbSubjectLabel.Append( site );
					sbSubjectLabel.Append( @"\" );
					// subject
					string sSubjectLabel = "";
					if( dataTable.Rows[0]["LOCALIDENTIFIER1"] != System.DBNull.Value )
					{
						sSubjectLabel = dataTable.Rows[0]["LOCALIDENTIFIER1"].ToString();
					}
					if(sSubjectLabel != "")
					{
						// from database
						sbSubjectLabel.Append( sSubjectLabel );
					}
					else
					{
						// subject number
						sbSubjectLabel.Append( "(" );
						sbSubjectLabel.Append( subjectNo );
						sbSubjectLabel.Append( ")" );
					}
				}
				catch(Exception ex)
				{
					// log error
					StringBuilder sbError = new StringBuilder();
					sbError.Append( "Error trying to retrieve subject label - study=");
					sbError.Append( studyId );
					sbError.Append( " site=" );
					sbError.Append( site );
					sbError.Append( " subject=" );
					sbError.Append( subjectNo );
					sbError.Append( "\nError Desc=" );
					sbError.Append( ex.Message );
					log.Error( sbError.ToString() );
				}
			}
			else
			{
				// log warning as searching for non-existant subject
				StringBuilder sbWarn = new StringBuilder();
				sbWarn.Append( "Could not retrieve subject label - study=");
				sbWarn.Append( studyId );
				sbWarn.Append( " site=" );
				sbWarn.Append( site );
				sbWarn.Append( " subject=" );
				sbWarn.Append( subjectNo );
				log.Warn( sbWarn.ToString() );
			}

			// return 
			return sbSubjectLabel.ToString();
		}

		/// <summary>
		/// Check if subject label is permissible to be viewed by user - return label if so
		/// </summary>
		/// <param name="bufferMACROUser"></param>
		/// <param name="studyId"></param>
		/// <param name="site"></param>
		/// <param name="subjectNo"></param>
		/// <param name="bPermissionCheck"></param>
		/// <returns></returns>
		public static string SubjectLabel(BufferMACROUser bufferMACROUser, int studyId, string site, int subjectNo, bool bPermissionCheck)
		{
			StringBuilder sbSubjectLabel = new StringBuilder();

			// get connection string
			string sDbConn = bufferMACROUser.MACROUser.CurrentDBConString;
			// study / site sql
			string sTrialDbColumn = "TRIALSUBJECT.CLINICALTRIALID";
			string sSiteDbColumn = "TRIALSUBJECT.TRIALSITE";
			// DPH 21/11/2007 change to user object since MACRO v3.0.76 led to this change
			MACROUserBS30.MACROUser blankUser = null;
			string sStudySiteSQL = bufferMACROUser.MACROUser.DataLists.StudiesSitesWhereSQL( ref sTrialDbColumn, ref sSiteDbColumn, ref blankUser);
			// permissions boolean
			bool bExtractLabel = true;

			// extract subject label
			StringBuilder sbSql = new StringBuilder();
			sbSql.Append("SELECT TRIALSUBJECT.LOCALIDENTIFIER1, CLINICALTRIAL.CLINICALTRIALNAME ");
			sbSql.Append( "FROM TRIALSUBJECT, CLINICALTRIAL " );
			sbSql.Append( "WHERE TRIALSUBJECT.CLINICALTRIALID = CLINICALTRIAL.CLINICALTRIALID " );
			sbSql.Append( "AND TRIALSUBJECT.CLINICALTRIALID = " );
			sbSql.Append(studyId);
			sbSql.Append(" AND TRIALSUBJECT.TRIALSITE = '");
			sbSql.Append(site);
			sbSql.Append("' AND TRIALSUBJECT.PERSONID = ");
			sbSql.Append(subjectNo);
			// DPH 07/12/2006 - check if user can get to subject
			if( bPermissionCheck )
			{
				sbSql.Append( " AND " );
				sbSql.Append( sStudySiteSQL );
			}

			// log sql
			if( bPermissionCheck )
			{
				log.Debug ( "SubjectLabel with permissions sql=" + sbSql.ToString() );
			}
			else
			{
				log.Debug ( "SubjectLabel (no permission check) sql=" + sbSql.ToString() );
			}

			// get data table
			DataTable dataTable = GetDataTable(sDbConn, sbSql.ToString());

			// if zero rows searching for a non-existant subject / subject not have access to
			if( dataTable.Rows.Count == 0)
			{
				// can't extract label
				bExtractLabel = false;

				// log warning as searching for non-existant subject
				StringBuilder sbWarn = new StringBuilder();
				sbWarn.Append( "Could not retrieve subject label - study=");
				sbWarn.Append( studyId );
				sbWarn.Append( " site=" );
				sbWarn.Append( site );
				sbWarn.Append( " subject=" );
				sbWarn.Append( subjectNo );
				sbWarn.Append( " - this may be due to user permissions" );
				log.Warn( sbWarn.ToString() );
			}

			// if have label permission get label content
			if( bExtractLabel )
			{
				// have permission so get label !
				sbSubjectLabel.Append( SubjectLabel( bufferMACROUser, studyId, site, subjectNo ) );
			}

			// return 
			return sbSubjectLabel.ToString();
		}

		/// <summary>
		/// render page header
		/// </summary>
		/// <param name="sHeader"></param>
		/// <returns></returns>
		private static string RenderHeader(string sHeader)
		{
			StringBuilder sbHeader = new StringBuilder();

			sbHeader.Append( "<body>" );
			sbHeader.Append( "<table width=75% border=\"0\" cellpadding=\"0\" cellspacing=\"2\" id=\"Table0\">" );
			sbHeader.Append( "<tr><td colspan=2 align=center><font color=\"#336699\" style=\"font-family:verdana,arial,helvetica;font-size:14pt;\">" );
			sbHeader.Append( "MACRO Buffer Summary" );
			sbHeader.Append( "</font></td></tr>" );
			if( sHeader != "" )
			{
				sbHeader.Append( "<tr><td colspan=2 align=center><font color=\"#336699\" style=\"font-family:verdana,arial,helvetica;font-size:14pt;\">" );
				// header
				sbHeader.Append( sHeader );
				sbHeader.Append( "</font></td></tr>" );
			}
			// spacer
			sbHeader.Append( "<tr><td height=\"50px\" colspan=2></td></tr>" );

			return sbHeader.ToString();
		}

		/// <summary>
		/// render subject specific section on page
		/// </summary>
		/// <param name="sSubjectLabel"></param>
		/// <param name="lSubjectData"></param>
		/// <returns></returns>
		private static string RenderSubjectSection(string sSubjectLabel, long lSubjectData, int studyId, string site, int subjectNo)
		{
			StringBuilder sbSection = new StringBuilder();

			sbSection.Append( "<td>" );
			sbSection.Append( "<table width=\"300px\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" ID=\"Table1\">" );
			sbSection.Append( "<tr>" );
			sbSection.Append( "<td>" );
			sbSection.Append( "<table width=\"300px\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" style=\"font-family:verdana,arial,helvetica;font-size:8pt;\" ID=\"Table4\">" );
			sbSection.Append( "<tr bgcolor=\"#CCCCCC\">" );
			sbSection.Append( "<td>" );
			sbSection.Append( "<img src=\".." + @"\" + "img" + @"\" + "curve.gif\" border=\"0\">" );
			sbSection.Append( "</td>" );
			sbSection.Append( "<td>" );
			sbSection.Append( "<font color=\"#336699\">" );
			sbSection.Append( "<b>" );
			if( sSubjectLabel != "" )
			{
				sbSection.Append( sSubjectLabel );
			}
			else
			{
				sbSection.Append( "Unknown subject" );
			}
			sbSection.Append( "</b></font>" );
			sbSection.Append( "</td>" );
			sbSection.Append( "</tr>" );
			sbSection.Append( "</table>" );
			sbSection.Append( "<div style=\"width:100%;border-bottom:#cccccc 1px solid;border-left:#cccccc 1px solid;border-right:#cccccc 1px solid;background-color:#F0F0F0\">" );
			sbSection.Append( "<table cellspacing=\"0\" cellpadding=\"0\" border=\"0\" width=100% style=\"font-family:verdana,arial,helvetica;font-size:8pt;\" ID=\"Table5\">" );
			sbSection.Append( "<tr><td width=20px>&nbsp;</td><td align=left>" );
			if(lSubjectData > 0)
			{
				sbSection.Append( "<a style=\"cursor:hand;color:#0000ff\" onclick=\"javascript:NavigateToPage('" );
				sbSection.Append( "BufferDataBrowser.asp?fltSt=" );
				sbSection.Append( studyId );
				sbSection.Append( "&fltSi=" );
				sbSection.Append( site );
				sbSection.Append( "&fltSj=" );
				sbSection.Append( subjectNo );
				sbSection.Append( "');\">Data to Review (" );
				sbSection.Append( lSubjectData );
				sbSection.Append( ")</a>" );
			}
			else
			{
				sbSection.Append( "There is no data to review" );
			}
			sbSection.Append( "</td></tr>" );
			sbSection.Append( "<tr height='1'><td colspan=\"2\" align=left bgcolor='#C0C0C0'></td></tr>" );
			sbSection.Append( "<tr><td width=20px>&nbsp;</td><td align=left>" );
			if( sSubjectLabel != "" )
			{
				sbSection.Append( "<a style=\"cursor:hand;color:#0000ff\" onclick=\"javascript:NavigateToPage('" );
				sbSection.Append( "Schedule.asp?fltSt=" );
				sbSection.Append( studyId );
				sbSection.Append( "&fltSi=" );
				sbSection.Append( site );
				sbSection.Append( "&fltSj=" );
				sbSection.Append( subjectNo );
				sbSection.Append( "&new=0');\">Go to Subject schedule</a>" );
			}
			else
			{
				sbSection.Append( "Go to Subject schedule" );
			}
			sbSection.Append( "</td></tr></table>" );
			sbSection.Append( "</div>" );
			sbSection.Append( "</td>" );
			sbSection.Append( "</tr>" );
			sbSection.Append( "</table>" );
			sbSection.Append( "</td>" );

			return sbSection.ToString();
		}

		/// <summary>
		/// render 'standard' section on page
		/// </summary>
		/// <param name="studyId"></param>
		/// <param name="dataTotal"></param>
		/// <param name="sTitle"></param>
		/// <returns></returns>
		private static string RenderSection(int studyId, long dataTotal, string sTitle)
		{
			StringBuilder sbSection = new StringBuilder();
			
			sbSection.Append( "<td>" );
			sbSection.Append( "<table width=\"300px\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" ID=\"Table" );
			sbSection.Append( Guid.NewGuid().ToString() );
			sbSection.Append( "\">" );
			sbSection.Append( "<tr>" );
			sbSection.Append( "<td>" );
			sbSection.Append( "<table width=\"300px\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" style=\"font-family:verdana,arial,helvetica;font-size:8pt;\" ID=\"Table" );
			sbSection.Append( Guid.NewGuid().ToString() );
			sbSection.Append( "\">" );
			sbSection.Append( "<tr bgcolor=\"#CCCCCC\">" );
			sbSection.Append( "<td>" );
			sbSection.Append( "<img src=\".." + @"\" + "img" + @"\" + "curve.gif\" border=\"0\">" );
			sbSection.Append( "</td>" );
			sbSection.Append( "<td>" );
			sbSection.Append( "<font color=\"#336699\">" );
			sbSection.Append( "<b>" );
			sbSection.Append( sTitle );
			sbSection.Append( "</b></font>" );
			sbSection.Append( "</td>" );
			sbSection.Append( "</tr>" );
			sbSection.Append( "</table>" );
			sbSection.Append( "<div style=\"width:100%;border-bottom:#cccccc 1px solid;border-left:#cccccc 1px solid;border-right:#cccccc 1px solid;background-color:#F0F0F0\">" );
			sbSection.Append( "<table cellspacing=\"0\" cellpadding=\"0\" border=\"0\" width=100% style=\"font-family:verdana,arial,helvetica;font-size:8pt;\" ID=\"Table" );
			sbSection.Append( Guid.NewGuid().ToString() );
			sbSection.Append( "\">" );
			sbSection.Append( "<tr><td width=20px>&nbsp;</td><td align=left>" );
			if( dataTotal > 0 )
			{
				sbSection.Append( "<a style=\"cursor:hand;color:#0000ff\" onclick=\"javascript:NavigateToPage('" );
				sbSection.Append( "BufferDataBrowser.asp?fltSt=" );
				sbSection.Append( studyId );
				sbSection.Append( "&fltSi=&fltSj=-1" );
				sbSection.Append( "');\">Data to Review (" );
				sbSection.Append( dataTotal );
				sbSection.Append( ")</a>" );
			}
			else
			{
				sbSection.Append( "There is no data to review" );
			}
			sbSection.Append( "</td></tr>" );
			sbSection.Append( "</table>" );
			sbSection.Append( "</div>" );
			sbSection.Append( "</td>" );
			sbSection.Append( "</tr>" );
			sbSection.Append( "</table>" );
			sbSection.Append( "</td>" );

			return sbSection.ToString();
		}

		/// <summary>
		/// collect subject buffer data count
		/// </summary>
		/// <param name="bufferMACROUser"></param>
		/// <param name="studyId"></param>
		/// <param name="site"></param>
		/// <param name="subjectNo"></param>
		/// <returns></returns>
		private static long GetSubjectBufferDataCount(BufferMACROUser bufferMACROUser, int studyId, string site, int subjectNo)
		{
			// subject responses counter
			long lSubjectResponses = 0;

			// get connection string
			string sDbConn = bufferMACROUser.MACROUser.CurrentDBConString;

			// get sql 
			string sTrialDbColumn = "BUFFERRESPONSEDATA.CLINICALTRIALID";
			string sSiteDbColumn = "BUFFERRESPONSEDATA.SITE";
			// DPH 21/11/2007 change to user object since MACRO v3.0.76 led to this change
			MACROUserBS30.MACROUser blankUser = null;
			string sStudySiteSQL = bufferMACROUser.MACROUser.DataLists.StudiesSitesWhereSQL( ref sTrialDbColumn, ref sSiteDbColumn, ref blankUser );

			// extract number of buffer responses available for subject
			StringBuilder sbSql = new StringBuilder();
			sbSql.Append("SELECT CLINICALTRIALID, CLINICALTRIALNAME, SITE, PERSONID, COUNT(BUFFERRESPONSEDATAID) AS DATACOUNT ");
			sbSql.Append( "FROM BUFFERRESPONSEDATA " );
			sbSql.Append( "WHERE BUFFERCOMMITSTATUS " );
			// DPH 24/03/2006 - store buffer sql clause in a constant
			sbSql.Append( BufferBrowser._DISPLAY_DATA_SQL_CLAUSE );
			sbSql.Append( "AND CLINICALTRIALID = " );
			sbSql.Append(studyId);
			sbSql.Append(" AND SITE = '");
			sbSql.Append(site);
			sbSql.Append("' AND PERSONID = ");
			sbSql.Append(subjectNo);
			// include where clause containing allowed study / site permissions
			sbSql.Append( " AND " );
			sbSql.Append( sStudySiteSQL );
			// group by
			sbSql.Append( " GROUP BY CLINICALTRIALID, CLINICALTRIALNAME, SITE, PERSONID " );

			// log sql
			log.Debug ( "GetSubjectBufferDataCount sql=" + sbSql.ToString() );

			// get data table
			DataTable dataTable = GetDataTable(sDbConn, sbSql.ToString());

			// should only be 1 matching subject row - may be 0 (if no bufferdata)
			if( dataTable.Rows.Count == 1)
			{
				lSubjectResponses = Convert.ToInt64( dataTable.Rows[0]["DATACOUNT"] );
			}

			return lSubjectResponses;
		}

		/// <summary>
		/// return count for all studies block
		/// </summary>
		/// <param name="dtStudyData"></param>
		/// <returns></returns>
		private static long GetAllStudiesDataCount(DataTable dtStudyData)
		{
			long lAllStudiesCount = 0;

			// loop through studies
			foreach(DataRow drStudy in dtStudyData.Rows)
			{
				// collect sum of studies
				lAllStudiesCount += Convert.ToInt64( drStudy["DATACOUNT"] );
			}

			return lAllStudiesCount;
		}

		/// <summary>
		/// return study buffer data count datatable from database
		/// </summary>
		/// <param name="bufferMACROUser">MACRO user object</param>
		/// <returns></returns>
		private static DataTable GetStudyData(BufferMACROUser bufferMACROUser)
		{
			// get connection string
			string sDbConn = bufferMACROUser.MACROUser.CurrentDBConString;

			// get sql 
			string sTrialDbColumn = "BUFFERRESPONSEDATA.CLINICALTRIALID";
			string sSiteDbColumn = "BUFFERRESPONSEDATA.SITE";
			// DPH 21/11/2007 change to user object since MACRO v3.0.76 led to this change
			MACROUserBS30.MACROUser blankUser = null;
			string sStudySiteSQL = bufferMACROUser.MACROUser.DataLists.StudiesSitesWhereSQL( ref sTrialDbColumn, ref sSiteDbColumn, ref blankUser );

			// extract number of buffer responses available for subject
			StringBuilder sbSql = new StringBuilder();
			sbSql.Append("SELECT CLINICALTRIALID, CLINICALTRIALNAME, COUNT(BUFFERRESPONSEDATAID) AS DATACOUNT ");
			sbSql.Append( "FROM BUFFERRESPONSEDATA " );
			sbSql.Append( "WHERE BUFFERCOMMITSTATUS " );
			// DPH 24/03/2006 - store buffer sql clause in a constant
			sbSql.Append( BufferBrowser._DISPLAY_DATA_SQL_CLAUSE );
			// include where clause containing allowed study / site permissions
			sbSql.Append( " AND " );
			sbSql.Append( sStudySiteSQL );
			// group by
			sbSql.Append( " GROUP BY CLINICALTRIALID, CLINICALTRIALNAME " );

			// log sql
			log.Debug ( "GetStudyData sql=" + sbSql.ToString() );

			// get data table
			DataTable dtStudyData = GetDataTable(sDbConn, sbSql.ToString());

			return dtStudyData;
		}

		/// <summary>
		/// Return DataTable given Connection string & sql
		/// </summary>
		/// <param name="sDbConnection"></param>
		/// <param name="sSql"></param>
		/// <returns></returns>
		public static DataTable GetDataTable(string sDbConnection, string sSql)
		{
			DataTable dataTable = null;

			try
			{
				// connect to macro db
				IMEDDataAccess imedDb = new IMEDDataAccess();
				// macro connection
				IMEDDataAccess.ConnectionType eConnType = IMEDDataAccess.CalculateConnectionType(sDbConnection);
				IDbConnection dbConn = imedDb.GetConnection( eConnType, sDbConnection );
				IDbCommand dbCommand = null;
				IDbDataAdapter dbDataAdapter = null;
				DataSet dataDB = new DataSet();

				// get command
				dbCommand = imedDb.GetCommand( dbConn );
				// create command text
				dbCommand.CommandText = sSql;
				// get data adaptor
				dbDataAdapter = imedDb.GetDataAdapter(eConnType,dbCommand);
				// fill dataset
				dbDataAdapter.Fill(dataDB);
				// get datatable
				dataTable = dataDB.Tables[0];
				// dispose command
				dbCommand.Dispose();
				// close connection
				dbConn.Close();
			}
			catch(Exception ex)
			{
				log.Error( "Data table exception.\nSql=" + sSql, ex );
				throw ( new Exception( ex.Message ) );
			}
			// return DataSet
			return dataTable;
		}
	}
}
