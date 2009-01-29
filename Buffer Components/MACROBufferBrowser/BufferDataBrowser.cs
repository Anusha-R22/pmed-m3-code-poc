using System;
using System.Text;
using log4net;
using InferMed.Components;
using System.Data;
using System.Collections;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for BufferDataBrowser.
	/// </summary>
	class BufferDataBrowser
	{
		public BufferDataBrowser()
		{
			_dtBuffer = null;
			_bookMark = 0;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="bufferUser"></param>
		/// <param name="studyId"></param>
		/// <param name="site"></param>
		/// <param name="subjectNo"></param>
		/// <param name="bookMark"></param>
		public BufferDataBrowser(BufferMACROUser bufferUser, int studyId, string site, int subjectNo,
			string bookMark)
		{
			// calculate available unknown data
			CalculateAvailableBufferData(true, bufferUser, studyId, site, subjectNo);

			// calculate matches for unknown target data
			CalculateMatches(bufferUser, studyId, site, subjectNo);

			// calculate data for display
			CalculateAvailableBufferData(false, bufferUser, studyId, site, subjectNo);

			// store bookmark
			_bookMark = 0;
			try
			{
				_bookMark = Convert.ToInt32( bookMark );
			}
			catch{}
		}

		// private members
		DataTable _dtBuffer;
		int _bookMark;

		// log4net
		private static readonly ILog log = LogManager.GetLogger( typeof(BufferSummary) );

		// properties
		public string BookMark
		{
			get
			{
				return _bookMark.ToString();
			}
		}

		// functions
		/// <summary>
		/// calculate available data
		/// </summary>
		/// <param name="bUnknown"></param>
		/// <param name="bufferUser"></param>
		/// <param name="studyId"></param>
		/// <param name="site"></param>
		/// <param name="subjectNo"></param>
		private void CalculateAvailableBufferData(bool bUnknown, BufferMACROUser bufferUser, int studyId, string site, 
										int subjectNo)
		{
			// calculate data available to user / dependent upon mode - patient aware / study / all
			// store in private member dataset
			
			// log
			log.Info( "CalculateAvailableBufferData" );

			// get connection string
			string sDbConn = bufferUser.MACROUser.CurrentDBConString;

			// get sql where clause to determine studies / sites available to user
			string sTrialDbColumn = "BUFFERRESPONSEDATA.CLINICALTRIALID";
			string sSiteDbColumn = "BUFFERRESPONSEDATA.SITE";
			// DPH 21/11/2007 change to user object since MACRO v3.0.76 led to this change
			MACROUserBS30.MACROUser blankUser = null;
			string sStudySiteSQL = bufferUser.MACROUser.DataLists.StudiesSitesWhereSQL( ref sTrialDbColumn, ref sSiteDbColumn, ref blankUser );

			// extract number of buffer responses available for subject
			StringBuilder sbSql = new StringBuilder();
			// are using sql server MACRO database
			bool bSqlServerDb = (bufferUser.MACROUser.Database.DatabaseType == MACROUserBS30.eMACRODatabaseType.mdtSQLServer )?true:false;
																															
			// if just collecting unknown rows for matching
			if( bUnknown )
			{
				// just collect unknown target data
				sbSql.Append( GetUnknownDataItemSql( studyId, site, subjectNo, sStudySiteSQL, bSqlServerDb ) );
				// order by
				sbSql.Append( " ORDER BY BUFFERRESPONSEDATA.CLINICALTRIALNAME, BUFFERRESPONSEDATA.SITE, BUFFERRESPONSEDATA.PERSONID, BUFFERRESPONSEDATA.VISITID, BUFFERRESPONSEDATA.CRFPAGEID, BUFFERRESPONSEDATA.DATAITEMID " );
			}
			else
			{
				// collect unknown AND all known target data
				sbSql.Append( "SELECT bufdat.* FROM ( " );
				// known data				
				if( bSqlServerDb )
				{
					// sql server specific
					sbSql.Append( GetBufferResponsesSqlServer( studyId, site, subjectNo, sStudySiteSQL ) );
				}
				else
				{
					// Oracle specific
					sbSql.Append( GetBufferResponsesSqlOracle( studyId, site, subjectNo, sStudySiteSQL ) );
				}
				// union
				sbSql.Append( " UNION " );
				// unknown target data
				sbSql.Append( GetUnknownDataItemSql( studyId, site, subjectNo, sStudySiteSQL, bSqlServerDb ) );
				sbSql.Append( " ) bufdat " );
				// order by
				sbSql.Append( "ORDER BY lower( bufdat.CLINICALTRIALNAME ), bufdat.SITE, bufdat.PERSONID, bufdat.VISITORDER, bufdat.VISITID, bufdat.VISITCYCLENUMBER, " );
				sbSql.Append( "bufdat.CRFPAGEORDER, bufdat.CRFPAGEID, bufdat.CRFPAGECYCLENUMBER, bufdat.FIELDORDER, bufdat.RESPONSEREPEATNUMBER, bufdat.QGROUPFIELDORDER " );
			}

			// log sql
			log.Debug ( "CalculateAvailableBufferData sql = " + sbSql.ToString() );

			// get data table
			_dtBuffer = BufferSummary.GetDataTable(sDbConn, sbSql.ToString());
		}

		// calculate exact matches for data
		/// <summary>
		/// 
		/// </summary>
		/// <param name="bufferUser"></param>
		/// <param name="studyId"></param>
		/// <param name="site"></param>
		/// <param name="subjectNo"></param>
		private void CalculateMatches(BufferMACROUser bufferUser, int studyId, string site, int subjectNo)
		{
			// log
			log.Info( "CalculateMatches" );

			// get connection string
			string sDbConn = bufferUser.MACROUser.CurrentDBConString;

			// loop through dataset and attempt to get exact matches for unknown items
			ArrayList alDataItems = new ArrayList();

			// collect specified range
			int nTargetRange = Convert.ToInt32( BufferBrowser.GetSetting(BufferBrowser._BUFFER_TARGET_RANGE, "0" ) );
			log.Debug( "Working directory=" + System.IO.Path.GetDirectoryName( AppDomain.CurrentDomain.BaseDirectory ) );
			log.Debug( "Target range = " + nTargetRange.ToString() );

			// identify unknown items data item ids
			foreach(DataRow drBuffer in _dtBuffer.Rows)
			{
				// get data item id
				long lDataItemId = Convert.ToInt64( drBuffer["DATAITEMID"] );
				// add to unique list
				if( ! alDataItems.Contains( lDataItemId ) )
				{
					alDataItems.Add( lDataItemId );
				}
			}

			// create one datatable containing all unknown data items in visit / eform order
			/*
			StringBuilder sbSql = new StringBuilder();

			sbSql.Append( "SELECT CRFPAGEINSTANCE.CLINICALTRIALID, CRFPAGEINSTANCE.TRIALSITE, CRFPAGEINSTANCE.PERSONID, CRFPAGEINSTANCE.VISITID, " );
			sbSql.Append( "CRFPAGEINSTANCE.CRFPAGEID, CRFPAGEINSTANCE.VISITCYCLENUMBER, CRFPAGEINSTANCE.CRFPAGECYCLENUMBER, " );
			sbSql.Append( "CRFPAGEINSTANCE.CRFPAGEDATE, CRFPAGEINSTANCE.CRFPAGESTATUS, CRFPAGEINSTANCE.CRFPAGEINSTANCELABEL, " );
			sbSql.Append( "DATAITEMRESPONSE.DATAITEMID, DATAITEMRESPONSE.RESPONSEVALUE, DATAITEMRESPONSE.RESPONSETASKID, DATAITEMRESPONSE.REPEATNUMBER, " );
			sbSql.Append( "STUDYVISIT.VISITCODE, CRFPAGE.CRFPAGECODE " );
			sbSql.Append( "FROM DATAITEMRESPONSE, CRFELEMENT, CRFPAGEINSTANCE, CRFPAGE, STUDYVISIT " );
			sbSql.Append( "WHERE (DATAITEMRESPONSE.CLINICALTRIALID = CRFELEMENT.CLINICALTRIALID) AND (DATAITEMRESPONSE.CRFPAGEID = CRFELEMENT.CRFPAGEID) " );
			sbSql.Append( "AND (DATAITEMRESPONSE.CRFELEMENTID = CRFELEMENT.CRFELEMENTID) " );
			sbSql.Append( "AND (DATAITEMRESPONSE.CLINICALTRIALID = CRFPAGEINSTANCE.CLINICALTRIALID) AND (DATAITEMRESPONSE.TRIALSITE = CRFPAGEINSTANCE.TRIALSITE) " );
			sbSql.Append( "AND (DATAITEMRESPONSE.PERSONID = CRFPAGEINSTANCE.PERSONID) AND (DATAITEMRESPONSE.CRFPAGETASKID = CRFPAGEINSTANCE.CRFPAGETASKID) " );
			sbSql.Append( "AND (DATAITEMRESPONSE.CLINICALTRIALID = CRFPAGE.CLINICALTRIALID) AND (CRFPAGEINSTANCE.CRFPAGEID = CRFPAGE.CRFPAGEID) " );
			sbSql.Append( "AND (DATAITEMRESPONSE.CLINICALTRIALID = STUDYVISIT.CLINICALTRIALID) AND (CRFPAGEINSTANCE.VISITID = STUDYVISIT.VISITID) " );
			sbSql.Append( "AND (DATAITEMRESPONSE.CLINICALTRIALID = " );
			sbSql.Append( studyId );
			sbSql.Append( ") " );
			if( alDataItems.Count > 0 )
			{
				sbSql.Append( " AND (DATAITEMRESPONSE.DATAITEMID IN ( " );
				int nDataItemCounter = 0;
				foreach( long lDataItemId in alDataItems )
				{
					if( nDataItemCounter > 0 )
					{
						sbSql.Append( " , " );
					}
					sbSql.Append( lDataItemId );
					nDataItemCounter++;
				}
				sbSql.Append( " ) ) " );
			}
			// trialsite
			sbSql.Append( "AND (DATAITEMRESPONSE.TRIALSITE = '" );
			sbSql.Append( site );
			sbSql.Append( "') " );
			// personid
			sbSql.Append( "AND (DATAITEMRESPONSE.PERSONID = " );
			sbSql.Append( subjectNo );
			sbSql.Append( ") " );

			sbSql.Append( "ORDER BY CRFPAGEINSTANCE.PERSONID, STUDYVISIT.VISITORDER, CRFPAGEINSTANCE.VISITID, CRFPAGEINSTANCE.VISITCYCLENUMBER, CRFPAGE.CRFPAGEORDER, " );
			sbSql.Append( "CRFPAGEINSTANCE.CRFPAGEID, CRFPAGEINSTANCE.CRFPAGECYCLENUMBER, CRFELEMENT.FIELDORDER " );

			// log sql
			log.Debug ( "CalculateMatches sql = " + sbSql.ToString() );

			DataTable dtPossibleMatchList = BufferSummary.GetDataTable(sDbConn, sbSql.ToString() );
			*/

			DataTable dtPossibleMatchList = GetPossibleMatches( bufferUser, studyId, site, subjectNo, alDataItems );

			// loop through dataset
			foreach(DataRow drBuffer in _dtBuffer.Rows)
			{
				// ready array lists
				ArrayList alCloseMatch = new ArrayList();
				ArrayList alNormalMatch = new ArrayList();

				// filter possible matches on current data item
				StringBuilder sbFilter = new StringBuilder();
				// studyid
				sbFilter.Append( "( CLINICALTRIALID = " );
				sbFilter.Append( drBuffer["CLINICALTRIALID"].ToString() );
				sbFilter.Append( " ) " );
				// site
				sbFilter.Append( "AND ( TRIALSITE = '" );
				sbFilter.Append( drBuffer["SITE"].ToString() );
				sbFilter.Append( "' ) " );
				// personid
				sbFilter.Append( "AND ( PERSONID = " );
				sbFilter.Append( drBuffer["PERSONID"].ToString() );
				sbFilter.Append( " ) " );
				// DPH 17/03/2006 - Include visit & eform (if known to narrow down possible targets)
				// visit data
				if( drBuffer["VISITID"] != System.DBNull.Value )
				{
					sbFilter.Append( "AND ( VISITID = " );
					sbFilter.Append( drBuffer["VISITID"].ToString() );
					sbFilter.Append( " ) " );
					// include visit cycle if it exists 
					if( drBuffer["VISITCYCLENUMBER"] != System.DBNull.Value )
					{
						sbFilter.Append( "AND ( VISITCYCLENUMBER = " );
						sbFilter.Append( drBuffer["VISITCYCLENUMBER"].ToString() );
						sbFilter.Append( " ) " );
					}
				}
				// eform data
				if( drBuffer["CRFPAGEID"] != System.DBNull.Value )
				{
					sbFilter.Append( "AND ( CRFPAGEID = " );
					sbFilter.Append( drBuffer["CRFPAGEID"].ToString() );
					sbFilter.Append( " ) " );
					// include eForm cycle if it exists 
					if( drBuffer["CRFPAGECYCLENUMBER"] != System.DBNull.Value )
					{
						sbFilter.Append( "AND ( CRFPAGECYCLENUMBER = " );
						sbFilter.Append( drBuffer["CRFPAGECYCLENUMBER"].ToString() );
						sbFilter.Append( " ) " );
					}
				}
				// dataitemid
				sbFilter.Append( "AND ( DATAITEMID = " );
				sbFilter.Append( drBuffer["DATAITEMID"].ToString() );
				sbFilter.Append( " ) " );
				//string sFilter = "DATAITEMID = " + drBuffer["DATAITEMID"];

				// log filter
				log.Debug ( "CalculateMatches sFilter = " + sbFilter.ToString() );

				foreach( DataRow drPossMatch in dtPossibleMatchList.Select( sbFilter.ToString() ) )
				{
					// target match values
					int nTargetResponseTaskId = BufferBrowser._DEFAULT_MISSING_NUMERIC;
					if( drPossMatch["RESPONSETASKID"] != System.DBNull.Value )
					{
						nTargetResponseTaskId = Convert.ToInt32( drPossMatch["RESPONSETASKID"] );
					}
					int nTargetRepeatNumber = BufferBrowser._DEFAULT_MISSING_NUMERIC;
					if( drPossMatch["REPEATNUMBER"] != System.DBNull.Value )
					{
						nTargetRepeatNumber = Convert.ToInt32( drPossMatch["REPEATNUMBER"] );
					}
					string sTargetVisitCode = "";
					if( drPossMatch["VISITCODE"] != System.DBNull.Value )
					{
						sTargetVisitCode = drPossMatch["VISITCODE"].ToString();
					}
					int nTargetVisitId = BufferBrowser._DEFAULT_MISSING_NUMERIC;
					if( drPossMatch["VISITID"] != System.DBNull.Value )
					{
						nTargetVisitId = Convert.ToInt32( drPossMatch["VISITID"] );
					}
					int nTargetVisitCycleNumber = BufferBrowser._DEFAULT_MISSING_NUMERIC;
					if( drPossMatch["VISITCYCLENUMBER"] != System.DBNull.Value )
					{
						nTargetVisitCycleNumber = Convert.ToInt32( drPossMatch["VISITCYCLENUMBER"] );
					}
					string nTargetCrfPageCode = "";
					if( drPossMatch["CRFPAGECODE"] != System.DBNull.Value )
					{
						nTargetCrfPageCode = drPossMatch["CRFPAGECODE"].ToString();
					}
					int nTargetCrfPageId = BufferBrowser._DEFAULT_MISSING_NUMERIC;
					if( drPossMatch["CRFPAGEID"] != System.DBNull.Value )
					{
						nTargetCrfPageId = Convert.ToInt32( drPossMatch["CRFPAGEID"] );
					}
					int nTargetCrfPageCycleNumber = BufferBrowser._DEFAULT_MISSING_NUMERIC;
					if( drPossMatch["CRFPAGECYCLENUMBER"] != System.DBNull.Value )
					{
						nTargetCrfPageCycleNumber = Convert.ToInt32( drPossMatch["CRFPAGECYCLENUMBER"] );
					}

					// does buffer data item have an order date & possible match have an eform date
					if( ( drBuffer["ORDERDATETIME"] != System.DBNull.Value) && ( drPossMatch["CRFPAGEDATE"] != System.DBNull.Value) )
					{
						// order date YES - order date = eform date?
						DateTime dtOrder = DateTime.FromOADate( Convert.ToDouble( drBuffer["ORDERDATETIME"] ) );
						DateTime dtEform = DateTime.FromOADate( Convert.ToDouble( drPossMatch["CRFPAGEDATE"] ) );

						// if either date is equal to 0 do not do date compare
						if( (dtOrder.ToOADate() != 0) && (dtEform.ToOADate() != 0 ) )
						{
							if( dtOrder == dtEform )
							{
								// order date = eform date YES - add data item to close match list with a 'zero' rating
								TargetDataItemMatch targetDataItem = new TargetDataItemMatch(studyId, site, subjectNo, nTargetResponseTaskId, 
									nTargetRepeatNumber, 0, sTargetVisitCode, nTargetVisitId, nTargetVisitCycleNumber,
									nTargetCrfPageCode, nTargetCrfPageId, nTargetCrfPageCycleNumber);

								// add to close match list
								alCloseMatch.Add( targetDataItem );
							}
							else
							{
								// order date = eform date NO - does order date fall within specified range?
								int nCloseRating = System.Math.Abs( Convert.ToInt32( DateDiff( "d", dtOrder, dtEform ) ) );
							
								if( nCloseRating <= nTargetRange )
								{
									// does order date fall within specified range YES - add data item to close match list with a difference rating - abs(date diff)
									TargetDataItemMatch targetDataItem = new TargetDataItemMatch(studyId, site, subjectNo, 
										nTargetResponseTaskId, nTargetRepeatNumber, nCloseRating, sTargetVisitCode, nTargetVisitId, 
										nTargetVisitCycleNumber, nTargetCrfPageCode, nTargetCrfPageId, nTargetCrfPageCycleNumber );

									// add to close match list
									alCloseMatch.Add( targetDataItem );
								}
								else
								{
									// does order date fall within specified range NO - add to 'normal' matches list
									TargetDataItemMatch targetDataItem = new TargetDataItemMatch(studyId, site, subjectNo, 
										nTargetResponseTaskId, nTargetRepeatNumber, sTargetVisitCode, nTargetVisitId, 
										nTargetVisitCycleNumber, nTargetCrfPageCode, nTargetCrfPageId, nTargetCrfPageCycleNumber );

									// add to normal match list
									alNormalMatch.Add( targetDataItem );
								}
							}
						}
						else
						{
							// order date NO - add to 'normal' matches list
							TargetDataItemMatch targetDataItem = new TargetDataItemMatch(studyId, site, subjectNo, nTargetResponseTaskId, 
								nTargetRepeatNumber, sTargetVisitCode, nTargetVisitId, nTargetVisitCycleNumber,
								nTargetCrfPageCode, nTargetCrfPageId, nTargetCrfPageCycleNumber );

							// add to normal match list
							alNormalMatch.Add( targetDataItem );
						}
					}
					else
					{
						// order date NO - add to 'normal' matches list
						TargetDataItemMatch targetDataItem = new TargetDataItemMatch(studyId, site, subjectNo, nTargetResponseTaskId, 
													nTargetRepeatNumber, sTargetVisitCode, nTargetVisitId, nTargetVisitCycleNumber,
													nTargetCrfPageCode, nTargetCrfPageId, nTargetCrfPageCycleNumber );

						// add to normal match list
						alNormalMatch.Add( targetDataItem );
					}
				}

				// do items exist in close matches list?
				if( alCloseMatch.Count > 0 )
				{
					// do items exist in close matches list YES
					// if close match count = 1
					if( alCloseMatch.Count == 1 )
					{
						// if close match count = 1 TRUE - store data item as match
						StoreBufferTarget( bufferUser, drBuffer["BUFFERRESPONSEDATAID"].ToString() , (TargetDataItemMatch)alCloseMatch[0] );
					}
					else
					{
						// if close match count = 1 FALSE - loop through close matches list
						// collect 1st match with lowest close rating. Store data item as match
						int nClosestMatch = BufferBrowser._DEFAULT_MISSING_NUMERIC;
						// find lowest close rating
						foreach(TargetDataItemMatch targetMatch in alCloseMatch)
						{
							if( targetMatch.CloseRating != nClosestMatch)
							{
								if( nClosestMatch == BufferBrowser._DEFAULT_MISSING_NUMERIC) 
								{
									nClosestMatch = targetMatch.CloseRating;
								}
								else
								{
									if( targetMatch.CloseRating < nClosestMatch )
									{
										nClosestMatch = targetMatch.CloseRating;
									}
								}
							}
						}
						// get 1st 'lowest' close rating
						foreach(TargetDataItemMatch targetMatch in alCloseMatch)
						{
							if( targetMatch.CloseRating == nClosestMatch )
							{
								// store data item as match
								StoreBufferTarget( bufferUser, drBuffer["BUFFERRESPONSEDATAID"].ToString(), targetMatch );
								break;
							}
						}
					}
				}
				else
				{
					// do items exist in close matches list NO
					// if normal match list count =  1?
					if( alNormalMatch.Count == 1 )
					{
						// if normal match list count =  1 TRUE - store data item as match
						StoreBufferTarget( bufferUser, drBuffer["BUFFERRESPONSEDATAID"].ToString(), (TargetDataItemMatch)alNormalMatch[0] );
					}
					else
					{
						// if normal match list count =  1 FALSE - no exact match found - no action
					}
				}
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="bufferUser"></param>
		/// <param name="sUniqueRowId"></param>
		/// <param name="targetDataItemMatch"></param>
		private void StoreBufferTarget(BufferMACROUser bufferUser, string sUniqueRowId, TargetDataItemMatch targetDataItemMatch)
		{
			// connection string
			string sDbConnection = bufferUser.MACROUser.CurrentDBConString;
			// update sql
			StringBuilder sbUpdateSql = new StringBuilder();
			sbUpdateSql.Append( "UPDATE BUFFERRESPONSEDATA SET VISITCODE = " );
			if( targetDataItemMatch.VisitCode != "" )
			{
				sbUpdateSql.Append( "'" );
				sbUpdateSql.Append( targetDataItemMatch.VisitCode );
				sbUpdateSql.Append( "'" );
			}
			else
			{
				sbUpdateSql.Append( "null" );
			}
			sbUpdateSql.Append( ", VISITID = " );
			if( targetDataItemMatch.VisitId != BufferBrowser._DEFAULT_MISSING_NUMERIC )
			{
				sbUpdateSql.Append( targetDataItemMatch.VisitId );
			}
			else
			{
				sbUpdateSql.Append( "null" );
			}
			sbUpdateSql.Append( ", VISITCYCLENUMBER = " );
			if( targetDataItemMatch.VisitCycle != BufferBrowser._DEFAULT_MISSING_NUMERIC )
			{
				sbUpdateSql.Append( targetDataItemMatch.VisitCycle );
			}
			else
			{
				sbUpdateSql.Append( "null" );
			}
			sbUpdateSql.Append( ", CRFPAGECODE = " );
			if( targetDataItemMatch.CrfPageCode != "" )
			{
				sbUpdateSql.Append( "'" );
				sbUpdateSql.Append( targetDataItemMatch.CrfPageCode );
				sbUpdateSql.Append( "'" );
			}
			else
			{
				sbUpdateSql.Append( "null" );
			}
			sbUpdateSql.Append( ", CRFPAGEID = " );
			if( targetDataItemMatch.CrfPageId != BufferBrowser._DEFAULT_MISSING_NUMERIC )
			{
				sbUpdateSql.Append( targetDataItemMatch.CrfPageId );
			}
			else
			{
				sbUpdateSql.Append( "null" );
			}
			sbUpdateSql.Append( ", CRFPAGECYCLENUMBER = " );
			if( targetDataItemMatch.CrfPageCycle != BufferBrowser._DEFAULT_MISSING_NUMERIC )
			{
				sbUpdateSql.Append( targetDataItemMatch.CrfPageCycle );
			}
			else
			{
				sbUpdateSql.Append( "null" );
			}
			sbUpdateSql.Append( " WHERE BUFFERRESPONSEDATAID = '" );
			sbUpdateSql.Append( sUniqueRowId );
			sbUpdateSql.Append( "'" );

			// write to db
			// connect to macro db
			IMEDDataAccess imedDb = new IMEDDataAccess();
			// macro connection
			IMEDDataAccess.ConnectionType eConnType = IMEDDataAccess.CalculateConnectionType(sDbConnection);
			IDbConnection dbConn = imedDb.GetConnection( eConnType, sDbConnection );
			IDbCommand dbCommand = null;

			// open connection
			dbConn.Open();

			// get command object 
			dbCommand = imedDb.GetCommand(dbConn);
			
			// add to command object
			dbCommand.CommandText = sbUpdateSql.ToString();

			// execute command
			dbCommand.ExecuteNonQuery();

			// dispose command
			dbCommand.Dispose();
			
			// close connection
			dbConn.Close();
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="studyId"></param>
		/// <param name="site"></param>
		/// <param name="subjectNo"></param>
		/// <param name="sStudySiteSQL"></param>
		/// <returns></returns>
		private string GetUnknownDataItemSql(int studyId, string site, int subjectNo, string sStudySiteSQL, bool bSqlServer)
		{
			// extract number of buffer responses available for subject
			StringBuilder sbSql = new StringBuilder();
			sbSql.Append( "SELECT BUFFERRESPONSEDATA.BUFFERRESPONSEDATAID, BUFFERRESPONSEDATA.BUFFERRESPONSEID, BUFFERRESPONSEDATA.BUFFERCOMMITSTATUS, BUFFERRESPONSEDATA.BUFFERCOMMITSTATUSTIMESTAMP, " );
			sbSql.Append( "BUFFERRESPONSEDATA.BUFFERCOMMITSTATUSTIMESTAMP_TZ, BUFFERRESPONSEDATA.CLINICALTRIALNAME, BUFFERRESPONSEDATA.CLINICALTRIALID, BUFFERRESPONSEDATA.SITE, BUFFERRESPONSEDATA.SUBJECTLABEL, " );
			sbSql.Append( "BUFFERRESPONSEDATA.PERSONID, BUFFERRESPONSEDATA.VISITCODE, BUFFERRESPONSEDATA.VISITID, BUFFERRESPONSEDATA.VISITCYCLENUMBER, BUFFERRESPONSEDATA.CRFPAGECODE, " );
			sbSql.Append( "BUFFERRESPONSEDATA.CRFPAGEID, BUFFERRESPONSEDATA.CRFPAGECYCLENUMBER, BUFFERRESPONSEDATA.DATAITEMCODE, BUFFERRESPONSEDATA.DATAITEMID, BUFFERRESPONSEDATA.RESPONSEREPEATNUMBER, " );
			sbSql.Append( "BUFFERRESPONSEDATA.RESPONSEVALUE, BUFFERRESPONSEDATA.ORDERDATETIME, BUFFERRESPONSEDATA.USERNAME, BUFFERRESPONSEDATA.USERNAMEFULL, " );

			// sql server
			if( bSqlServer )
			{
				sbSql.Append( "null AS DIRESPONSEVALUE, null AS DIRESPONSESTATUS, null AS DISCREPANCYSTATUS, null AS LOCKSTATUS, null AS SDVSTATUS, null AS NOTESTATUS, null AS COMMENTS, null AS VISITNAME, null AS CRFTITLE, " );
				sbSql.Append( "DATAITEM.DATAITEMNAME, null AS VISITORDER, null AS CRFPAGEORDER, CRFELEMENT.FIELDORDER, null AS RESPONSETASKID, null AS CRFPAGETASKID, CRFELEMENT.QGROUPFIELDORDER " );
			}
			else
			{
				// oracle
				sbSql.Append( "'' DIRESPONSEVALUE, to_number(null) DIRESPONSESTATUS, to_number(null) DISCREPANCYSTATUS, to_number(null) LOCKSTATUS, to_number(null) SDVSTATUS, to_number(null) NOTESTATUS, '' COMMENTS, '' VISITNAME, '' CRFTITLE, " );
				sbSql.Append( "DATAITEM.DATAITEMNAME, to_number(null) VISITORDER, to_number(null) CRFPAGEORDER, CRFELEMENT.FIELDORDER, to_number(null) RESPONSETASKID, to_number(null) CRFPAGETASKID, CRFELEMENT.QGROUPFIELDORDER " );
			}
			sbSql.Append( "FROM BUFFERRESPONSEDATA " );
			if( bSqlServer )
			{
				// sql server
				sbSql.Append( "LEFT JOIN DATAITEM ON (BUFFERRESPONSEDATA.CLINICALTRIALID = DATAITEM.CLINICALTRIALID) " );
				sbSql.Append( "AND (BUFFERRESPONSEDATA.DATAITEMID = DATAITEM.DATAITEMID) " );
				sbSql.Append( "LEFT JOIN CRFELEMENT ON (BUFFERRESPONSEDATA.CLINICALTRIALID = CRFELEMENT.CLINICALTRIALID) AND (BUFFERRESPONSEDATA.CRFPAGEID = CRFELEMENT.CRFPAGEID) " );
				sbSql.Append( "AND (BUFFERRESPONSEDATA.DATAITEMID = CRFELEMENT.DATAITEMID) " );
				sbSql.Append( "WHERE " );
			}
			else
			{
				// oracle
				sbSql.Append( ", DATAITEM, CRFELEMENT " );
				sbSql.Append( "WHERE (BUFFERRESPONSEDATA.CLINICALTRIALID = DATAITEM.CLINICALTRIALID(+)) AND " );
				sbSql.Append( "(BUFFERRESPONSEDATA.DATAITEMID = DATAITEM.DATAITEMID(+)) AND " );
				sbSql.Append( "(BUFFERRESPONSEDATA.CLINICALTRIALID = CRFELEMENT.CLINICALTRIALID(+)) AND (BUFFERRESPONSEDATA.CRFPAGEID = CRFELEMENT.CRFPAGEID(+)) " );
				sbSql.Append( "AND (BUFFERRESPONSEDATA.DATAITEMID = CRFELEMENT.DATAITEMID(+)) AND " );
			}
			sbSql.Append( "BUFFERRESPONSEDATA.BUFFERCOMMITSTATUS " );
			// DPH 24/03/2006 - store buffer sql clause in a constant
			sbSql.Append( BufferBrowser._DISPLAY_DATA_SQL_CLAUSE );
			// check have study id
			if( studyId != BufferBrowser._DEFAULT_MISSING_NUMERIC )
			{
				sbSql.Append( "AND BUFFERRESPONSEDATA.CLINICALTRIALID = " );
				sbSql.Append(studyId);
			}
			// check if is subject specific
			if( (site != "") && ( subjectNo != BufferBrowser._DEFAULT_MISSING_NUMERIC ) )
			{
				sbSql.Append(" AND BUFFERRESPONSEDATA.SITE = '");
				sbSql.Append(site);
				sbSql.Append("' AND BUFFERRESPONSEDATA.PERSONID = ");
				sbSql.Append(subjectNo);
			}
			// if just collecting unknown rows for matching
			// just collect unknown target data
			sbSql.Append( " AND ( (BUFFERRESPONSEDATA.VISITID IS NULL) OR (BUFFERRESPONSEDATA.VISITCYCLENUMBER IS NULL) OR (BUFFERRESPONSEDATA.CRFPAGEID IS NULL) OR (BUFFERRESPONSEDATA.CRFPAGECYCLENUMBER IS NULL) ) " );

			// include where clause containing allowed study / site permissions
			sbSql.Append( " AND " );
			sbSql.Append( sStudySiteSQL );

			return sbSql.ToString();	
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="studyId"></param>
		/// <param name="site"></param>
		/// <param name="subjectNo"></param>
		/// <param name="sStudySiteSQL"></param>
		/// <returns></returns>
		private string GetBufferResponsesSqlServer(int studyId, string site, int subjectNo, string sStudySiteSQL)
		{
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
			sbSql.Append( "DATAITEMRESPONSE.RESPONSETASKID, CRFPAGEINSTANCE.CRFPAGETASKID, CRFELEMENT.QGROUPFIELDORDER " );
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
			sbSql.Append( "LEFT JOIN CRFPAGEINSTANCE ON (BUFFERRESPONSEDATA.CLINICALTRIALID = CRFPAGEINSTANCE.CLINICALTRIALID) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.SITE = CRFPAGEINSTANCE.TRIALSITE) AND (BUFFERRESPONSEDATA.PERSONID = CRFPAGEINSTANCE.PERSONID) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.VISITID = CRFPAGEINSTANCE.VISITID) AND (BUFFERRESPONSEDATA.VISITCYCLENUMBER = CRFPAGEINSTANCE.VISITCYCLENUMBER) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.CRFPAGEID = CRFPAGEINSTANCE.CRFPAGEID) AND (BUFFERRESPONSEDATA.CRFPAGECYCLENUMBER = CRFPAGEINSTANCE.CRFPAGECYCLENUMBER) " );
			sbSql.Append( "WHERE BUFFERCOMMITSTATUS " );
			// DPH 24/03/2006 - store buffer sql clause in a constant
			sbSql.Append( BufferBrowser._DISPLAY_DATA_SQL_CLAUSE );
			// check have study id
			if( studyId != BufferBrowser._DEFAULT_MISSING_NUMERIC )
			{
				sbSql.Append( "AND BUFFERRESPONSEDATA.CLINICALTRIALID = " );
				sbSql.Append(studyId);
			}
			// check if is subject specific
			if( (site != "") && ( subjectNo != BufferBrowser._DEFAULT_MISSING_NUMERIC ) )
			{
				sbSql.Append( " AND BUFFERRESPONSEDATA.SITE = '" );
				sbSql.Append( site );
				sbSql.Append( "' AND BUFFERRESPONSEDATA.PERSONID = " );
				sbSql.Append( subjectNo );
				sbSql.Append( " " );
			}
			// just collect known target data
			sbSql.Append( " AND ( (BUFFERRESPONSEDATA.VISITID IS NOT NULL) AND (BUFFERRESPONSEDATA.VISITCYCLENUMBER IS NOT NULL) AND (BUFFERRESPONSEDATA.CRFPAGEID IS NOT NULL) AND (BUFFERRESPONSEDATA.CRFPAGECYCLENUMBER IS NOT NULL) ) " );

			// include where clause containing allowed study / site permissions
			sbSql.Append( " AND " );
			sbSql.Append( sStudySiteSQL );

			// order by (not used as in union sql called)
			// ORDER BY CRFPAGEINSTANCE.PERSONID, STUDYVISIT.VISITORDER, CRFPAGEINSTANCE.VISITID, CRFPAGEINSTANCE.VISITCYCLENUMBER, CRFPAGE.CRFPAGEORDER,
			// CRFPAGEINSTANCE.CRFPAGEID, CRFPAGEINSTANCE.CRFPAGECYCLENUMBER, CRFELEMENT.FIELDORDER

			return sbSql.ToString();
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="studyId"></param>
		/// <param name="site"></param>
		/// <param name="subjectNo"></param>
		/// <param name="sStudySiteSQL"></param>
		/// <returns></returns>
		private string GetBufferResponsesSqlOracle(int studyId, string site, int subjectNo, string sStudySiteSQL)
		{
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
			sbSql.Append( "DATAITEMRESPONSE.RESPONSETASKID, CRFPAGEINSTANCE.CRFPAGETASKID, CRFELEMENT.QGROUPFIELDORDER " );
			sbSql.Append( "FROM BUFFERRESPONSEDATA, DATAITEMRESPONSE, CRFELEMENT, DATAITEM, CRFPAGEINSTANCE, CRFPAGE, STUDYVISIT " );
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
			// DPH 18/05/2005 - missing clinicaltrialid key link
			sbSql.Append( "AND (BUFFERRESPONSEDATA.CLINICALTRIALID = CRFPAGEINSTANCE.CLINICALTRIALID(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.SITE = CRFPAGEINSTANCE.TRIALSITE(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.PERSONID = CRFPAGEINSTANCE.PERSONID(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.VISITID = CRFPAGEINSTANCE.VISITID(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.VISITCYCLENUMBER = CRFPAGEINSTANCE.VISITCYCLENUMBER(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.CRFPAGEID = CRFPAGEINSTANCE.CRFPAGEID(+)) " );
			sbSql.Append( "AND (BUFFERRESPONSEDATA.CRFPAGECYCLENUMBER = CRFPAGEINSTANCE.CRFPAGECYCLENUMBER(+))" );
			sbSql.Append( "AND BUFFERCOMMITSTATUS " );
			// DPH 24/03/2006 - store buffer sql clause in a constant
			sbSql.Append( BufferBrowser._DISPLAY_DATA_SQL_CLAUSE );
			// check have study id
			if( studyId != BufferBrowser._DEFAULT_MISSING_NUMERIC )
			{
				sbSql.Append( "AND BUFFERRESPONSEDATA.CLINICALTRIALID = " );
				sbSql.Append(studyId);
			}
			// check if is subject specific
			if( (site != "") && ( subjectNo != BufferBrowser._DEFAULT_MISSING_NUMERIC ) )
			{
				sbSql.Append( " AND BUFFERRESPONSEDATA.SITE = '" );
				sbSql.Append( site );
				sbSql.Append( "' AND BUFFERRESPONSEDATA.PERSONID = " );
				sbSql.Append( subjectNo );
				sbSql.Append( " " );
			}
			// just collect known target data
			sbSql.Append( " AND ( (BUFFERRESPONSEDATA.VISITID IS NOT NULL) AND (BUFFERRESPONSEDATA.VISITCYCLENUMBER IS NOT NULL) AND (BUFFERRESPONSEDATA.CRFPAGEID IS NOT NULL) AND (BUFFERRESPONSEDATA.CRFPAGECYCLENUMBER IS NOT NULL) ) " );

			// include where clause containing allowed study / site permissions
			sbSql.Append( " AND " );
			sbSql.Append( sStudySiteSQL );

			// order by (not used as in union sql called)
			// ORDER BY CRFPAGEINSTANCE.PERSONID, STUDYVISIT.VISITORDER, CRFPAGEINSTANCE.VISITID, CRFPAGEINSTANCE.VISITCYCLENUMBER, CRFPAGE.CRFPAGEORDER,
			// CRFPAGEINSTANCE.CRFPAGEID, CRFPAGEINSTANCE.CRFPAGECYCLENUMBER, CRFELEMENT.FIELDORDER

			return sbSql.ToString();
		}

		/// <summary>
		/// same common params as the VBScript DateDiff
		/// </summary>
		/// <param name="comparisontype"></param>
		/// <param name="startDate"></param>
		/// <param name="endDate"></param>
		/// <returns></returns>
		private double DateDiff(string comparisontype, DateTime
			startDate, DateTime endDate) 
		{
			double diff=0;
			try 
			{
				TimeSpan TS = new
					TimeSpan( startDate.Ticks - endDate.Ticks );
				#region conversion options
				switch (comparisontype.ToLower()) 
				{
					case "m":
						diff = Convert.ToDouble(TS.TotalMinutes);
						break;
					case "s":
						diff = Convert.ToDouble(TS.TotalSeconds);
						break;
					case "t":
						diff = Convert.ToDouble(TS.Ticks);
						break;
					case "mm":
						diff = Convert.ToDouble(TS.TotalMilliseconds);
						break;
					case "yyyy":
						diff = Convert.ToDouble(TS.TotalDays/365);
						break;
					case "q":
						diff = Convert.ToDouble((TS.TotalDays/365)/4);
						break;
					default:
						//d
						diff = Convert.ToDouble(TS.TotalDays);
						break;
				}
				#endregion
			} 
			catch
			{
				diff = -1;
			}
			return diff;
		}

		/// <summary>
		/// Render HTML page body
		/// </summary>
		/// <param name="studyId"></param>
		/// <param name="site"></param>
		/// <param name="subjectNo"></param>
		/// <returns></returns>
		public string RenderBufferBrowserPage(BufferMACROUser bufferUser, int studyId, string site, int subjectNo)
		{
			log.Info( "Starting RenderBufferBrowserPage" );

			// render page given all data
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
			if( ( _bookMark > _dtBuffer.Rows.Count ) || ( _bookMark < 0 ) )
			{
				nStart = 0;
				_bookMark = 0;
			}
			else
			{
				nStart = _bookMark;
			}
			// calculate end row from given
			if( nStart + nPageSize >= _dtBuffer.Rows.Count )
			{
				nEnd = _dtBuffer.Rows.Count - 1;
			}
			else
			{
				nEnd = (nStart + nPageSize) - 1;
			}

			// all data should be ordered in _dtBuffer datatable
			// body tag
			sbPageHtml.Append( "<body>" );
			// internal form
			sbPageHtml.Append( "<form name='FormDR' action='" );
			// url
			sbPageHtml.Append( "BufferDataBrowser.asp?fltSt=" );
			sbPageHtml.Append( studyId );
			sbPageHtml.Append( "&fltSi=" );
			sbPageHtml.Append( site );
			sbPageHtml.Append( "&fltSj=" );
			sbPageHtml.Append( subjectNo );
			sbPageHtml.Append( "&bookmark=" );
			sbPageHtml.Append( _bookMark );
			sbPageHtml.Append( "' method='post'>" );
			sbPageHtml.Append( "<input type='hidden' name='bidentifier'>" );
			sbPageHtml.Append( "<input type='hidden' name='btype'>" );
			sbPageHtml.Append( "<input type='hidden' name='bback'>" );
			sbPageHtml.Append( "<input type='hidden' name='bscope'>" );
			sbPageHtml.Append( "</form>" );
			// main table
			sbPageHtml.Append( "<table style='cursor:default;' width='100%' class='clsTabletext' cellpadding='0' cellspacing='0' border='1' ID=\"Table1\">" );
			// info bar row
			sbPageHtml.Append( "<tr height='30'>" );
			sbPageHtml.Append( "<td colspan='8' align='left'>Record(s) " );
			// no of records
			sbPageHtml.Append( nStart + 1 );
			sbPageHtml.Append( " to " );
			// total on page
			sbPageHtml.Append( nEnd + 1 );
			sbPageHtml.Append( " of " );
			// grand total
			sbPageHtml.Append( _dtBuffer.Rows.Count );
			sbPageHtml.Append( "&nbsp;&nbsp;" );
			// paging icons & control
			int nPageBack = nStart - nPageSize;
			int nPageForward = nEnd + 1;
			// page back link
			if( nPageBack >= 0 )
			{
				// link
				sbPageHtml.Append( @"<a href='./BufferDataBrowser.asp?fltSt=" );
				sbPageHtml.Append( studyId );
				sbPageHtml.Append( "&fltSi=" );
				sbPageHtml.Append( site );
				sbPageHtml.Append( "&fltSj=" );
				sbPageHtml.Append( subjectNo );
				sbPageHtml.Append( "&bookmark=" );
				sbPageHtml.Append( nPageBack );
				sbPageHtml.Append( "'>" );
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
			if( nPageForward <= (_dtBuffer.Rows.Count - 1) )
			{
				sbPageHtml.Append( @"<a href='./BufferDataBrowser.asp?fltSt=" );
				sbPageHtml.Append( studyId );
				sbPageHtml.Append( "&fltSi=" );
				sbPageHtml.Append( site );
				sbPageHtml.Append( "&fltSj=" );
				sbPageHtml.Append( subjectNo );
				sbPageHtml.Append( "&bookmark=" );
				sbPageHtml.Append( nPageForward );
				sbPageHtml.Append( "'>" );
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
			if( nPageForward <= (_dtBuffer.Rows.Count - 1) )
			{
				sbPageHtml.Append( "</a>" );
			}
			sbPageHtml.Append( "&nbsp;&nbsp;" );
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
			sbPageHtml.Append( "<td>Order Date</td>" );
			sbPageHtml.Append( "<td>Target Value</td>" );
			sbPageHtml.Append( "<td>Target Status</td>" );
			sbPageHtml.Append( "</tr>" );

			// no data detail
			if( _dtBuffer.Rows.Count == 0 )
			{
				// render - There is no buffer data available for the current criteria
				sbPageHtml.Append( "<tr>" );
				sbPageHtml.Append( "<td valign='top' colspan='8'>" );
				sbPageHtml.Append( "There is no buffer data available for the current criteria" );
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
						// do this via a filter on the current rows values
						StringBuilder sbFilter = new StringBuilder();
						// trial id
						sbFilter.Append( "( CLINICALTRIALID = " );
						sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["CLINICALTRIALID"] );
						sbFilter.Append( ") AND " );
						// site
						sbFilter.Append( "(SITE = '" );
						sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["SITE"] );
						sbFilter.Append( "') AND " );
						// subject
						sbFilter.Append( "(PERSONID = " );
						sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["PERSONID"] );
						sbFilter.Append( ")" );

						// put into arraylist & count
						ArrayList alSubjectRows = new ArrayList();
						alSubjectRows.AddRange( _dtBuffer.Select( sbFilter.ToString() ) );
						nRowSpanSubject = alSubjectRows.Count; 

						// calculate for page (some rows may have been on last page)
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
 
						// only need to write this <td> once per subject on the page
						// link
						sbPageHtml.Append( GetJavascriptMenuDetail( nDataRow, 0, studyId, site, subjectNo ) );

						sbPageHtml.Append( "<td valign='top'" );
						// do we need a rowspan?
						if( nRowSpanSubject > 1 )
						{
							sbPageHtml.Append( "rowspan='" );
							sbPageHtml.Append( nRowSpanSubject );
							sbPageHtml.Append( "'" );
						}
						sbPageHtml.Append( ">" );
						sbPageHtml.Append( BufferBrowser.SubjectLabel( _dtBuffer.Rows[ nDataRow ]["CLINICALTRIALNAME"].ToString(), 
							_dtBuffer.Rows[ nDataRow ]["SITE"].ToString(), 
							_dtBuffer.Rows[ nDataRow ]["SUBJECTLABEL"].ToString(),
							Convert.ToInt32( _dtBuffer.Rows[ nDataRow ]["PERSONID"] ) ) );
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
						sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["CLINICALTRIALID"] );
						sbFilter.Append( ") AND " );
						// site
						sbFilter.Append( "(SITE = '" );
						sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["SITE"] );
						sbFilter.Append( "') AND " );
						// subject
						sbFilter.Append( "(PERSONID = " );
						sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["PERSONID"] );
						sbFilter.Append( ") AND " );
						// visit
						sbFilter.Append( "(VISITID " );
						if( _dtBuffer.Rows[ nDataRow ]["VISITID"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// visit id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["VISITID"].ToString() );
						}
						sbFilter.Append( ") AND " );
						// visit cycle
						sbFilter.Append( "(VISITCYCLENUMBER " );
						if( _dtBuffer.Rows[ nDataRow ]["VISITCYCLENUMBER"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// visit cycle value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["VISITCYCLENUMBER"].ToString() );
						}
						sbFilter.Append( " )" );

						// put into arraylist & count
						ArrayList alVisitRows = new ArrayList();
						alVisitRows.AddRange( _dtBuffer.Select( sbFilter.ToString() ) );
						nRowSpanVisit = alVisitRows.Count; 

						// calculate for page (some rows may have been on last page)
						int nSectionRow = 0;
						for(nSectionRow = 0; nSectionRow < alVisitRows.Count; nSectionRow++)
						{
							// loop through rows in section
							if( ((DataRow)alVisitRows[nSectionRow])["BUFFERRESPONSEDATAID"].ToString() == 
								_dtBuffer.Rows[ nDataRow ]["BUFFERRESPONSEDATAID"].ToString() )
							{
								break;
							}
						}
						// adjust rowspan
						nRowSpanVisit -= nSectionRow;

						// only need to write this <td> once per visit on the page
						// link
						sbPageHtml.Append( GetJavascriptMenuDetail( nDataRow, 1, studyId, site, subjectNo ) );
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
						if( _dtBuffer.Rows[ nDataRow ]["VISITNAME"] != System.DBNull.Value )
						{
							sbPageHtml.Append( _dtBuffer.Rows[ nDataRow ]["VISITNAME"].ToString() );
							if( _dtBuffer.Rows[ nDataRow ]["VISITCYCLENUMBER"] != System.DBNull.Value )
							{	
								sbPageHtml.Append( " [" );
								sbPageHtml.Append( _dtBuffer.Rows[ nDataRow ]["VISITCYCLENUMBER"].ToString() );
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
						StringBuilder sbFilter = new StringBuilder();
						// trial id
						sbFilter.Append( "( CLINICALTRIALID = " );
						sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["CLINICALTRIALID"] );
						sbFilter.Append( ") AND " );
						// site
						sbFilter.Append( "(SITE = '" );
						sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["SITE"] );
						sbFilter.Append( "') AND " );
						// subject
						sbFilter.Append( "(PERSONID = " );
						sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["PERSONID"] );
						sbFilter.Append( ") AND " );
						// visit
						sbFilter.Append( "(VISITID " );
						if( _dtBuffer.Rows[ nDataRow ]["VISITID"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// visit id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["VISITID"].ToString() );
						}
						sbFilter.Append( ") AND " );
						// visit cycle
						sbFilter.Append( "(VISITCYCLENUMBER " );
						if( _dtBuffer.Rows[ nDataRow ]["VISITCYCLENUMBER"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// visit id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["VISITCYCLENUMBER"].ToString() );
						}
						sbFilter.Append( " ) AND " );
						// eform
						sbFilter.Append( "(CRFPAGEID " );
						if( _dtBuffer.Rows[ nDataRow ]["CRFPAGEID"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// crfpage id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["CRFPAGEID"].ToString() );
						}
						sbFilter.Append( " ) AND " );
						// eform cycle
						sbFilter.Append( "(CRFPAGECYCLENUMBER " );
						if( _dtBuffer.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"] == System.DBNull.Value )
						{
							// null
							sbFilter.Append( " IS NULL " );
						}
						else
						{
							// crfpage id value
							sbFilter.Append( " = " );
							sbFilter.Append( _dtBuffer.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"].ToString() );
						}
						sbFilter.Append( " ) " );

						// put into arraylist & count
						ArrayList alEformRows = new ArrayList();
						alEformRows.AddRange( _dtBuffer.Select( sbFilter.ToString() ) );
						nRowSpanEform = alEformRows.Count; 

						// calculate for page (some rows may have been on last page)
						int nSectionRow = 0;
						for(nSectionRow = 0; nSectionRow < alEformRows.Count; nSectionRow++)
						{
							// loop through rows in section
							if( ((DataRow)alEformRows[nSectionRow])["BUFFERRESPONSEDATAID"].ToString() == 
								_dtBuffer.Rows[ nDataRow ]["BUFFERRESPONSEDATAID"].ToString() )
							{
								break;
							}
						}
						// adjust rowspan
						nRowSpanEform -= nSectionRow;

						// only need to write this <td> once per visit on the page
						// link
						sbPageHtml.Append( GetJavascriptMenuDetail( nDataRow, 2, studyId, site, subjectNo ) );
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
						if( _dtBuffer.Rows[ nDataRow ]["CRFTITLE"] != System.DBNull.Value )
						{
							sbPageHtml.Append( _dtBuffer.Rows[ nDataRow ]["CRFTITLE"].ToString() );
							if( _dtBuffer.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"] != System.DBNull.Value )
							{
								sbPageHtml.Append( " [" );
								sbPageHtml.Append( _dtBuffer.Rows[ nDataRow ]["CRFPAGECYCLENUMBER"].ToString() );
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
					sbPageHtml.Append( GetJavascriptMenuDetail( nDataRow, 3, studyId, site, subjectNo ) );
					// question name
					sbPageHtml.Append( "<td valign='top'>" );
					if( _dtBuffer.Rows[ nDataRow ]["DATAITEMNAME"] != System.DBNull.Value )
					{
						sbPageHtml.Append( _dtBuffer.Rows[ nDataRow ]["DATAITEMNAME"].ToString() );
						// include response repeat number if greater than 1
						if( _dtBuffer.Rows[ nDataRow ]["RESPONSEREPEATNUMBER"] != System.DBNull.Value )
						{
							if( Convert.ToInt32( _dtBuffer.Rows[ nDataRow ]["RESPONSEREPEATNUMBER"] ) > 1 )
							{
								sbPageHtml.Append( " [" );
								sbPageHtml.Append( _dtBuffer.Rows[ nDataRow ]["RESPONSEREPEATNUMBER"].ToString() );
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
					sbPageHtml.Append( "<td valign='top'>" );
					if( _dtBuffer.Rows[ nDataRow ]["RESPONSEVALUE"] != System.DBNull.Value )
					{
						sbPageHtml.Append( _dtBuffer.Rows[ nDataRow ]["RESPONSEVALUE"].ToString() );
					}
					else
					{
						sbPageHtml.Append( "&nbsp;" );
					}
					sbPageHtml.Append( "</td>" );
					// order date
					sbPageHtml.Append( "<td valign='top'>" );
					if( _dtBuffer.Rows[ nDataRow ]["ORDERDATETIME"] != System.DBNull.Value )
					{
						DateTime dtOrder = DateTime.FromOADate( Convert.ToDouble( _dtBuffer.Rows[ nDataRow ]["ORDERDATETIME"].ToString() ) );
						sbPageHtml.Append( dtOrder.ToString( "dd/MM/yyyy" ) );
					}
					else
					{
						sbPageHtml.Append( "&nbsp;" );
					}
					sbPageHtml.Append( "</td>" );
					// target value
					sbPageHtml.Append( "<td valign='top'>" );
					if( _dtBuffer.Rows[ nDataRow ]["DIRESPONSEVALUE"] != System.DBNull.Value )
					{
						sbPageHtml.Append( _dtBuffer.Rows[ nDataRow ]["DIRESPONSEVALUE"].ToString() );
					}
					else
					{
						sbPageHtml.Append( "&nbsp;" );
					}
					sbPageHtml.Append( "</td>" );
					// target status
					sbPageHtml.Append( "<td valign='top'>" );
					if( _dtBuffer.Rows[ nDataRow ]["DIRESPONSESTATUS"] != System.DBNull.Value )
					{
						// get response status & display text
						int nResponseStatus = Convert.ToInt32( _dtBuffer.Rows[ nDataRow ]["DIRESPONSESTATUS"] );
						int nLockStatus = 0;
						if(	_dtBuffer.Rows[ nDataRow ]["LOCKSTATUS"] != System.DBNull.Value )
						{
							nLockStatus = Convert.ToInt32( _dtBuffer.Rows[ nDataRow ]["LOCKSTATUS"] );
						}
						int nSdvStatus = 0;
						if(	_dtBuffer.Rows[ nDataRow ]["SDVSTATUS"] != System.DBNull.Value )
						{
							nSdvStatus = Convert.ToInt32( _dtBuffer.Rows[ nDataRow ]["SDVSTATUS"] );
						}
						int nDiscStatus = 0;
						if(	_dtBuffer.Rows[ nDataRow ]["DISCREPANCYSTATUS"] != System.DBNull.Value )
						{
							nDiscStatus = Convert.ToInt32( _dtBuffer.Rows[ nDataRow ]["DISCREPANCYSTATUS"] );
						}
						bool bNote = false;
						if( _dtBuffer.Rows[ nDataRow ]["NOTESTATUS"] != null )
						{
							bNote = (Convert.ToInt32( _dtBuffer.Rows[ nDataRow ]["NOTESTATUS"] ) > 0)?true:false;
						}
						bool bComment = false;
						if( _dtBuffer.Rows[ nDataRow ]["COMMENTS"] != null )
						{
							bComment = (Convert.ToString( _dtBuffer.Rows[ nDataRow ]["COMMENTS"] ).Length > 0)?true:false;
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
		/// Collect Javascript menu detail
		/// </summary>
		/// <param name="bufferRow"></param>
		/// <param name="scope"></param>
		/// <returns></returns>
		private string GetJavascriptMenuDetail(int bufferRow, int scope, 
											int studyId, string site, int subjectNo)
		{
			// fnM - button, item detail *, scope 0 - subject 1 - visit 2- eform 3 - question, back detail +
			// name, value, commit flag, reject flag, 
			// target select flag, target change flag, goto eform flag
			// * item detail - study id`site`subject id`visit id`visit cycle`eform id`eform cycle`eform task id`eform element id`response cycle`response task id`buffer response id
			// + back detail - study id`site`subject id
			// i.e. <a onmouseup="fnM(event.button,'10`site1`1',1,'10`site1`1','ABC 1','',0,0,0,0,0);">
			StringBuilder sbJS = new StringBuilder();
			// calculate if unknown data item 
			bool bUnknown = false;
			// for this row
			if ( ( _dtBuffer.Rows[ bufferRow ]["VISITID"] == System.DBNull.Value ) || ( _dtBuffer.Rows[ bufferRow ]["VISITCYCLENUMBER"] == System.DBNull.Value )
				|| ( _dtBuffer.Rows[ bufferRow ]["CRFPAGEID"] == System.DBNull.Value ) || ( _dtBuffer.Rows[ bufferRow ]["CRFPAGECYCLENUMBER"] == System.DBNull.Value ) )
			{
				bUnknown = true;
			}
			// work out for whole block
			if( scope < 2)
			{
				// subject or visit
				// calculate the new span, if there is one, by counting
				// the number of rows where trial/site/subject/label values 
				// are the same as this one
				// do this via a filter on the current rows values
				StringBuilder sbFilter = new StringBuilder();
				// trial id
				sbFilter.Append( "( CLINICALTRIALID = " );
				sbFilter.Append( _dtBuffer.Rows[ bufferRow ]["CLINICALTRIALID"] );
				sbFilter.Append( ") AND " );
				// site
				sbFilter.Append( "(SITE = '" );
				sbFilter.Append( _dtBuffer.Rows[ bufferRow ]["SITE"] );
				sbFilter.Append( "') AND " );
				// subject
				sbFilter.Append( "(PERSONID = " );
				sbFilter.Append( _dtBuffer.Rows[ bufferRow ]["PERSONID"] );
				sbFilter.Append( ")" );

				// visit scope
				if( scope == 1 )
				{
					// visit
					sbFilter.Append( " AND (VISITID " );
					if( _dtBuffer.Rows[ bufferRow ]["VISITID"] == System.DBNull.Value )
					{
						// null
						sbFilter.Append( " IS NULL " );
					}
					else
					{
						// visit id value
						sbFilter.Append( " = " );
						sbFilter.Append( _dtBuffer.Rows[ bufferRow ]["VISITID"].ToString() );
					}
					sbFilter.Append( ") AND " );
					// visit cycle
					sbFilter.Append( "(VISITCYCLENUMBER " );
					if( _dtBuffer.Rows[ bufferRow ]["VISITCYCLENUMBER"] == System.DBNull.Value )
					{
						// null
						sbFilter.Append( " IS NULL " );
					}
					else
					{
						// visit cycle value
						sbFilter.Append( " = " );
						sbFilter.Append( _dtBuffer.Rows[ bufferRow ]["VISITCYCLENUMBER"].ToString() );
					}
					sbFilter.Append( " )" );
				}

				// put into arraylist & count
				ArrayList alSectionRows = new ArrayList();
				alSectionRows.AddRange( _dtBuffer.Select( sbFilter.ToString() ) );
				
				// check if there is an 'unknown' data item in the section
				foreach( DataRow drSection in alSectionRows )
				{
					if ( ( drSection["VISITID"] == System.DBNull.Value ) || ( drSection["VISITCYCLENUMBER"] == System.DBNull.Value )
						|| ( drSection["CRFPAGEID"] == System.DBNull.Value ) || ( drSection["CRFPAGECYCLENUMBER"] == System.DBNull.Value ) )
					{
						bUnknown = true;
						break;
					}
				}
			}
			// button
			sbJS.Append( "<a onmouseup=\"fnM(event.button,'" );
			// item detail - study id`site`subject id`visit id`visit cycle``eform id`eform cycle`eform task id`eform element id`response cycle`response task id
			// study id
			sbJS.Append( _dtBuffer.Rows[ bufferRow ]["CLINICALTRIALID"] );
			sbJS.Append( "`" );
			// site
			sbJS.Append( _dtBuffer.Rows[ bufferRow ]["SITE"] );
			sbJS.Append( "`" );
			// subject id
			sbJS.Append( _dtBuffer.Rows[ bufferRow ]["PERSONID"] );
			sbJS.Append( "`" );
			// visit info
			if( scope > 0 )
			{
				// visit id
				if( _dtBuffer.Rows[ bufferRow ]["VISITID"] != System.DBNull.Value )
				{
					// visit id value
					sbJS.Append( _dtBuffer.Rows[ bufferRow ]["VISITID"] );
				}
				sbJS.Append( "`" );
				// visit cycle
				if( _dtBuffer.Rows[ bufferRow ]["VISITCYCLENUMBER"] != System.DBNull.Value )
				{
					sbJS.Append( _dtBuffer.Rows[ bufferRow ]["VISITCYCLENUMBER"] );
				}
				sbJS.Append( "`" );
			}
			else
			{
				sbJS.Append( "``" );
			}
			// eform info
			if( scope > 1 )
			{
				// eform id
				if( _dtBuffer.Rows[ bufferRow ]["CRFPAGEID"] != System.DBNull.Value )
				{
					sbJS.Append( _dtBuffer.Rows[ bufferRow ]["CRFPAGEID"] );
				}
				sbJS.Append( "`" );
				// eform cycle
				if( _dtBuffer.Rows[ bufferRow ]["CRFPAGECYCLENUMBER"] != System.DBNull.Value )
				{
					sbJS.Append( _dtBuffer.Rows[ bufferRow ]["CRFPAGECYCLENUMBER"] );

				}
				sbJS.Append( "`" );
				// eform task id
				if( _dtBuffer.Rows[ bufferRow ]["CRFPAGETASKID"] != System.DBNull.Value )
				{
					sbJS.Append( _dtBuffer.Rows[ bufferRow ]["CRFPAGETASKID"] );
				}
				sbJS.Append( "`" );
			}
			else
			{
				sbJS.Append( "```" );
			}
			// question level
			if( scope > 2 )
			{
				// eform element id
				if( _dtBuffer.Rows[ bufferRow ]["DATAITEMID"] != System.DBNull.Value )
				{
					sbJS.Append( _dtBuffer.Rows[ bufferRow ]["DATAITEMID"] );
				}
				sbJS.Append( "`" );
				// response cycle
				if( _dtBuffer.Rows[ bufferRow ]["RESPONSEREPEATNUMBER"] != System.DBNull.Value )
				{
					sbJS.Append( _dtBuffer.Rows[ bufferRow ]["RESPONSEREPEATNUMBER"] );
				}
				sbJS.Append( "`" );
				// response task id
				if( _dtBuffer.Rows[ bufferRow ]["RESPONSETASKID"] != System.DBNull.Value )
				{
					sbJS.Append( _dtBuffer.Rows[ bufferRow ]["RESPONSETASKID"] );
				}
				sbJS.Append( "`" );
				// individual row guid
				sbJS.Append( _dtBuffer.Rows[ bufferRow ]["BUFFERRESPONSEDATAID"] );
				sbJS.Append( "" );
			}
			else
			{
				sbJS.Append( "```" );
			}
			sbJS.Append( "'," );
			// scope
			sbJS.Append( scope );
			// back detail - study id`site`subject id
			sbJS.Append( ",'" );
			// study id
			sbJS.Append( studyId );
			sbJS.Append( "`" );
			// site
			sbJS.Append( site );
			sbJS.Append( "`" );
			// subject id
			sbJS.Append( subjectNo );
			// name
			sbJS.Append( "','" );
			switch ( scope )
			{
				case 0:
				{
					// subject
					sbJS.Append( BufferBrowser.SubjectLabel( _dtBuffer.Rows[ bufferRow ]["CLINICALTRIALNAME"].ToString(), 
						_dtBuffer.Rows[ bufferRow ]["SITE"].ToString(), 
						_dtBuffer.Rows[ bufferRow ]["SUBJECTLABEL"].ToString(),
						Convert.ToInt32( _dtBuffer.Rows[ bufferRow ]["PERSONID"] ) ) );
					break;
				}
				case 1:
				{
					// visit
					if( _dtBuffer.Rows[ bufferRow ]["VISITNAME"] != System.DBNull.Value )
					{
						sbJS.Append( _dtBuffer.Rows[ bufferRow ]["VISITNAME"] );
					}
					break;
				}
				case 2:
				{
					// eform
					if( _dtBuffer.Rows[ bufferRow ]["CRFTITLE"] != System.DBNull.Value )
					{
						sbJS.Append( _dtBuffer.Rows[ bufferRow ]["CRFTITLE"] );
					}
					break;
				}
				case 3:
				{
					// question
					if( _dtBuffer.Rows[ bufferRow ]["DATAITEMNAME"] != System.DBNull.Value )
					{
						sbJS.Append( _dtBuffer.Rows[ bufferRow ]["DATAITEMNAME"] );
					}
					break;
				}
			}
			sbJS.Append( "'," );
			// Value
			sbJS.Append( "'" );
			if( scope == 3)
			{
				if( _dtBuffer.Rows[ bufferRow ]["RESPONSEVALUE"] != System.DBNull.Value )
				{
					sbJS.Append( _dtBuffer.Rows[ bufferRow ]["RESPONSEVALUE"] );
				}
			}
			sbJS.Append( "'," );
			// commit flag
			if( ! bUnknown )
			{
				// not locked or frozen
				if( _dtBuffer.Rows[ bufferRow ]["LOCKSTATUS"] != System.DBNull.Value )
				{
					if( Convert.ToInt32( _dtBuffer.Rows[ bufferRow ]["LOCKSTATUS"] ) > 4 )
					{
						// is locked or frozen
						sbJS.Append( "0," );
					}
					else
					{
						// can be committed
						sbJS.Append( "1," );
					}
				}
				else
				{
					sbJS.Append( "1," );
				}
			}
			else
			{
				sbJS.Append( "0," );
			}
			// reject flag
			sbJS.Append ( "1," );
			// target select flag
			// if unknown target
			if( bUnknown )
			{
				if( scope == 3 )
				{
					// target select flag true
					sbJS.Append ( "1," );
				}
				else
				{
					// target select flag false
					sbJS.Append ( "0," );
				}
				// target change flag false
				sbJS.Append ( "0," );
			}
			else
			{
				// target select flag false
				sbJS.Append ( "0," );
				if( scope == 3 )
				{
					// target change flag true
					sbJS.Append ( "1," );
				}
				else
				{
					// target change flag false
					sbJS.Append ( "0," );
				}
			}
			// goto eform flag
			if( (scope == 2) || (scope == 3) )
			{
				// check is a 'known' target item
				if ( ! bUnknown )
				{
					sbJS.Append( "1" );
				}
				else
				{
					sbJS.Append( "0" );
				}
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

		/// <summary>
		/// Get DataTable containing possible matches
		/// </summary>
		/// <param name="bufferUser"></param>
		/// <param name="studyId"></param>
		/// <param name="site"></param>
		/// <param name="subjectNo"></param>
		/// <returns></returns>
		private DataTable GetPossibleMatches( BufferMACROUser bufferUser, int studyId, string site, int subjectNo, ArrayList alDataItems )
		{
			// log
			log.Info( "GetPossibleMatches" );

			// get connection string
			string sDbConn = bufferUser.MACROUser.CurrentDBConString;

			// are using sql server MACRO database
			bool bSqlServer = (bufferUser.MACROUser.Database.DatabaseType == MACROUserBS30.eMACRODatabaseType.mdtSQLServer )?true:false;

			// create one datatable containing all unknown data items in visit / eform order
			StringBuilder sbSql = new StringBuilder();

			sbSql.Append( "SELECT bufdat.* FROM " );
			sbSql.Append( "( " );
			// known matching data items
			sbSql.Append( "SELECT CRFPAGEINSTANCE.CLINICALTRIALID, CRFPAGEINSTANCE.TRIALSITE, CRFPAGEINSTANCE.PERSONID, CRFPAGEINSTANCE.VISITID, " );
			sbSql.Append( "CRFPAGEINSTANCE.CRFPAGEID, CRFPAGEINSTANCE.VISITCYCLENUMBER, CRFPAGEINSTANCE.CRFPAGECYCLENUMBER, " );
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
			sbSql.Append( IMEDFunctions20.LocalNumToStandard( DateTime.Now.ToOADate().ToString() , false ) );
			sbSql.Append( " ) ) ");
			sbSql.Append( "AND (DATAITEMRESPONSE.CLINICALTRIALID = " );
			sbSql.Append( studyId );
			sbSql.Append( " )" );
			if( site != "" )
			{
				sbSql.Append( " AND (DATAITEMRESPONSE.TRIALSITE = '" );
				sbSql.Append( site );
				sbSql.Append( "')" );
			}
			if( subjectNo != BufferBrowser._DEFAULT_MISSING_NUMERIC )
			{
				sbSql.Append( " AND ( DATAITEMRESPONSE.PERSONID = " );
				sbSql.Append( subjectNo );
				sbSql.Append( " ) " );
			}
			if( alDataItems.Count > 0 )
			{
				sbSql.Append( " AND (DATAITEMRESPONSE.DATAITEMID IN ( " );
				int nDataItemCounter = 0;
				foreach( long lDataItemId in alDataItems )
				{
					if( nDataItemCounter > 0 )
					{
						sbSql.Append( " , " );
					}
					sbSql.Append( lDataItemId );
					nDataItemCounter++;
				}
				sbSql.Append( " ) ) " );
			}
			sbSql.Append( "UNION " );
			// unknown matching responses (not yet created)
			sbSql.Append( "SELECT CRFPAGEINSTANCE.CLINICALTRIALID, CRFPAGEINSTANCE.TRIALSITE, CRFPAGEINSTANCE.PERSONID, CRFPAGEINSTANCE.VISITID, " );
			sbSql.Append( "CRFPAGEINSTANCE.CRFPAGEID, CRFPAGEINSTANCE.VISITCYCLENUMBER, CRFPAGEINSTANCE.CRFPAGECYCLENUMBER, " );
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
			sbSql.Append( " ) ) ");
			sbSql.Append( "AND (CRFPAGEINSTANCE.CLINICALTRIALID = " );
			sbSql.Append( studyId );
			sbSql.Append( ")" );
			if( site != "" )
			{
				sbSql.Append( " AND (CRFPAGEINSTANCE.TRIALSITE = '" );
				sbSql.Append( site );
				sbSql.Append( "')" );
			}
			if( subjectNo != BufferBrowser._DEFAULT_MISSING_NUMERIC )
			{
				sbSql.Append( " AND (CRFPAGEINSTANCE.PERSONID = " );
				sbSql.Append( subjectNo );
				sbSql.Append( ") " );
			}
			if( alDataItems.Count > 0 )
			{
				sbSql.Append( " AND (DATAITEM.DATAITEMID IN ( " );
				int nDataItemCounter = 0;
				foreach( long lDataItemId in alDataItems )
				{
					if( nDataItemCounter > 0 )
					{
						sbSql.Append( " , " );
					}
					sbSql.Append( lDataItemId );
					nDataItemCounter++;
				}
				sbSql.Append( " ) ) " );
			}
			sbSql.Append( "AND (CRFELEMENT.HIDDEN = 0) " );
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
	}

#region TargetDataItemMatch
	/// <summary>
	/// Store unique details for a (potential / exact) data item match
	/// </summary>
	class TargetDataItemMatch
	{
		// private members - unique 
		private int _studyId;
		private string _site;
		private int _subjectId;
		private int _responseTaskId;
		private int _repeatNumber;
		private bool _closeMatch;
		private int _closeRating;
		// extra info for update
		private string _visitCode;
		private int _visitId;
		private int _visitCycleNumber;
		private string _crfPageCode;
		private int _crfPageId;
		private int _crfPageCycleNumber;

		public TargetDataItemMatch()
		{}

		public TargetDataItemMatch(int studyId, string site, int subjectId, int responseTaskId, 
			int repeatNumber, string visitCode, int visitId, int visitCycleNumber,
			string crfPageCode, int crfPageId, int crfPageCycleNumber)
		{
			_studyId = studyId;
			_site = site;
			_subjectId = subjectId;
			_responseTaskId = responseTaskId;
			_repeatNumber = repeatNumber;
			_closeMatch = false;
			_closeRating = BufferBrowser._DEFAULT_MISSING_NUMERIC;
			_visitCode = visitCode;
			_visitId = visitId;
			_visitCycleNumber = visitCycleNumber;
			_crfPageCode = crfPageCode;
			_crfPageId = crfPageId;
			_crfPageCycleNumber = crfPageCycleNumber;
		}

		public TargetDataItemMatch(int studyId, string site, int subjectId, int responseTaskId, 
			int repeatNumber, int closeRating, string visitCode, int visitId, int visitCycleNumber,
			string crfPageCode, int crfPageId, int crfPageCycleNumber)
		{
			_studyId = studyId;
			_site = site;
			_subjectId = subjectId;
			_responseTaskId = responseTaskId;
			_repeatNumber = repeatNumber;
			_closeMatch = true;
			_closeRating = closeRating;
			_visitCode = visitCode;
			_visitId = visitId;
			_visitCycleNumber = visitCycleNumber;
			_crfPageCode = crfPageCode;
			_crfPageId = crfPageId;
			_crfPageCycleNumber = crfPageCycleNumber;
		}

		// properties
		public int StudyId
		{
			get
			{
				return _studyId;
			}
		}

		public string Site
		{
			get
			{
				return _site;
			}
		}

		public int SubjectId
		{
			get
			{
				return _subjectId;
			}
		}

		public int ResponseTaskId
		{
			get
			{
				return _responseTaskId;
			}
		}

		public int RepeatNumber
		{
			get
			{
				return _repeatNumber;
			}
		}

		public bool CloseMatch
		{
			get
			{
				return _closeMatch;
			}
			set
			{
				_closeMatch = value;
			}
		}

		public int CloseRating
		{
			get
			{
				return _closeRating;
			}
			set
			{
				_closeRating = value;
			}
		}

		public string VisitCode
		{
			get
			{
				return _visitCode;
			}
			set
			{
				_visitCode = value;
			}
		}

		public int VisitId
		{
			get
			{
				return _visitId;
			}
			set
			{
				_visitId = value;
			}
		}

		public int VisitCycle
		{
			get
			{
				return _visitCycleNumber;
			}
			set
			{
				_visitCycleNumber = value;
			}
		}

		public string CrfPageCode
		{
			get
			{
				return _crfPageCode;
			}
			set
			{
				_crfPageCode = value;
			}
		}

		public int CrfPageId
		{
			get
			{
				return _crfPageId;
			}
			set
			{
				_crfPageId = value;
			}
		}

		public int CrfPageCycle
		{
			get
			{
				return _crfPageCycleNumber;
			}
			set
			{
				_crfPageCycleNumber = value;
			}
		}
	}
	#endregion
}
