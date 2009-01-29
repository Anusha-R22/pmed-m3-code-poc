using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.OracleClient;
using System.IO;
using System.Text;
using log4net;
using InferMed.Components;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for BufferResponseDb.
	/// </summary>
	class BufferResponseDb
	{
		// member variables
		private string _sEncryptSecurityDbConn;
		private string _sSecurityDbConn;
		private string _sMACROdb;
		private string _sMACROdbConn;
		private bool _bConnOk;

		// log4net
		private static readonly ILog log = LogManager.GetLogger( typeof(BufferAPI) );

		private static int _BUFFERSIZE = 2048;

		/// <summary>
		/// 
		/// </summary>
		public BufferResponseDb()
		{
			// buffer flag
			_bConnOk = false;
			_sMACROdbConn = "";
			_sSecurityDbConn = "";
			// Encrypted Security DB string
			_sEncryptSecurityDbConn = BufferAPI.GetSetting( "securitypath" , "" );
			if(_sEncryptSecurityDbConn != "")
			{
				// security connection
				_sSecurityDbConn = IMEDEncryption.DecryptString( _sEncryptSecurityDbConn );
				// macro db to connect to
				_sMACROdb = BufferAPI.GetSetting( BufferAPI._MACRO_CONN_DB, "" );
				// log
				log.Debug( "Macro db setting (bufferdb) " + _sMACROdb );
				// CreateDbConnection
				CreateDbConnection();
				// have collected data so ok
				_bConnOk = true;
			}
			else
			{
				throw( new Exception("Failed to initialise MACRO database connection.") );
			}
		}

		/// <summary>
		/// 
		/// </summary>
		public bool ConnectionOk
		{
			get
			{
				return _bConnOk;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		private void CreateDbConnection()
		{
			// connect to security db
			IMEDDataAccess imedDb = new IMEDDataAccess();
			// security connection
			IMEDDataAccess.ConnectionType eConnType = IMEDDataAccess.CalculateConnectionType(_sSecurityDbConn);
			IDbConnection dbConn = imedDb.GetConnection( eConnType, _sSecurityDbConn );
			IDbCommand dbCommand = null;
			IDbDataAdapter dbDataAdapter = null;
			DataSet dataDB = new DataSet();
			DataTable dataTable = null;

			// open connection
			dbConn.Open();

			// get MACRO database connection from DATABASES
			// OR use one connected to now if MACROdb not given
			if(_sMACROdb != "")
			{
				// get command
				dbCommand = imedDb.GetCommand( dbConn );
				// databases sql
				string sSQL = "SELECT DATABASECODE, SERVERNAME, DATABASETYPE, NAMEOFDATABASE, DATABASEUSER, DATABASEPASSWORD FROM DATABASES WHERE DATABASECODE = '" + _sMACROdb + "'";
				// create command text
				dbCommand.CommandText = sSQL;
				// get data adaptor
				dbDataAdapter = imedDb.GetDataAdapter(eConnType,dbCommand);
				// fill dataset
				dbDataAdapter.Fill(dataDB);
				// get datatable
				dataTable = dataDB.Tables[0];

				// get data
				string sServerName = dataTable.Rows[0]["SERVERNAME"].ToString();
				string sDbName = dataTable.Rows[0]["NAMEOFDATABASE"].ToString();
				int nDbType = Convert.ToInt32(dataTable.Rows[0]["DATABASETYPE"]);
				string sEncryptDbUser = dataTable.Rows[0]["DATABASEUSER"].ToString();
				string sEncryptDbPwd = dataTable.Rows[0]["DATABASEPASSWORD"].ToString();
				// formulate connection string
				if(nDbType == 1)
				{
					// sql server connection string
					_sMACROdbConn=IMEDDataAccess.CreateDBConnectionString(IMEDDataAccess.ConnectionType.SQLServer, sServerName,
						sDbName, IMEDEncryption.DecryptString(sEncryptDbUser), IMEDEncryption.DecryptString(sEncryptDbPwd));
				}
				else
				{
					// oracle database connection string
					_sMACROdbConn=IMEDDataAccess.CreateDBConnectionString(IMEDDataAccess.ConnectionType.Oracle, sDbName,
						"", IMEDEncryption.DecryptString(sEncryptDbUser), IMEDEncryption.DecryptString(sEncryptDbPwd));
				}

				// dispose of command
				dbCommand.Dispose();
			}
			else
			{
				// use current connection
				_sMACROdbConn = _sSecurityDbConn;
			}

			// close connection to security db
			dbConn.Close();
		}

		/// <summary>
		/// Return DataTable given Connection string & sql
		/// </summary>
		/// <param name="sDbConnection"></param>
		/// <param name="sSql"></param>
		/// <returns></returns>
		private DataTable GetDataSet(string sDbConnection, string sSql)
		{
			// connect to security db
			IMEDDataAccess imedDb = new IMEDDataAccess();
			// security connection
			IMEDDataAccess.ConnectionType eConnType = IMEDDataAccess.CalculateConnectionType(sDbConnection);
			IDbConnection dbConn = imedDb.GetConnection( eConnType, sDbConnection );
			IDbCommand dbCommand = null;
			IDbDataAdapter dbDataAdapter = null;
			DataSet dataDB = new DataSet();
			DataTable dataTable = null;

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

			// return DataSet
			return dataTable;
		}

		/// <summary>
		/// Get study definition id
		/// </summary>
		/// <param name="sStudyCode"></param>
		/// <returns></returns>
		public int GetStudyDefinitionId(string sStudyCode)
		{
			int nStudyDefId = BufferAPI._DEFAULT_MISSING_NUMERIC;

			// sql
			string sSql = "SELECT CLINICALTRIALID FROM CLINICALTRIAL WHERE CLINICALTRIALNAME = '" + sStudyCode + "'";

			// log
			log.Debug( "sSql= " + sSql );

			// get data table
			DataTable dataTable = GetDataSet(_sMACROdbConn, sSql);

			// will only be 1 matching row
			nStudyDefId = Convert.ToInt32(dataTable.Rows[0]["CLINICALTRIALID"]);

			return nStudyDefId;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="nStudyId"></param>
		/// <param name="sSite"></param>
		/// <returns></returns>
		public bool CheckStudySite(int nStudyId, string sSite)
		{
			bool bStudySiteOk = false;

			// sql
			StringBuilder sbSql = new StringBuilder();
			sbSql.Append("SELECT CLINICALTRIALID, TRIALSITE, STUDYVERSION FROM TRIALSITE WHERE CLINICALTRIALID = ");
			sbSql.Append(nStudyId);
			sbSql.Append(" AND TRIALSITE = '");
			sbSql.Append(sSite);
			sbSql.Append("'");

			// log
			log.Debug( "sSql= " + sbSql.ToString() );

			// get data table
			DataTable dataTable = GetDataSet(_sMACROdbConn, sbSql.ToString());

			// check there is a matching row
			bStudySiteOk = (dataTable.Rows[0]["TRIALSITE"].ToString()==sSite)?true:false;

			return bStudySiteOk;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="nStudyId"></param>
		/// <param name="sSite"></param>
		/// <param name="sSubjectLabel"></param>
		/// <returns></returns>
		public int GetSubjectId(int nStudyId, string sSite, string sSubjectLabel)
		{
			int nSubjectId = BufferAPI._DEFAULT_MISSING_NUMERIC;

			// sql
			StringBuilder sbSql = new StringBuilder();
			sbSql.Append("SELECT PERSONID FROM TRIALSUBJECT WHERE CLINICALTRIALID = ");
			sbSql.Append(nStudyId);
			sbSql.Append(" AND TRIALSITE = '");
			sbSql.Append(sSite);
			sbSql.Append("' AND LOCALIDENTIFIER1 = '");
			sbSql.Append(sSubjectLabel);
			sbSql.Append("'");

			// log
			log.Debug( "sSql= " + sbSql.ToString() );

			// get data table
			DataTable dataTable = GetDataSet(_sMACROdbConn, sbSql.ToString());

			// should only be 1 matching row
			if( dataTable.Rows.Count == 0)
			{
				// no matches on subject label
				// if input is numeric check against MACRO subject
				int nSubjLabel = BufferAPI._DEFAULT_MISSING_NUMERIC;
				try
				{
					nSubjLabel = Convert.ToInt32( sSubjectLabel );
				}
				catch{}

				if(nSubjLabel != BufferAPI._DEFAULT_MISSING_NUMERIC)
				{
					// search on numeric
					sbSql = new StringBuilder();
					sbSql.Append("SELECT PERSONID FROM TRIALSUBJECT WHERE CLINICALTRIALID = ");
					sbSql.Append(nStudyId);
					sbSql.Append(" AND TRIALSITE = '");
					sbSql.Append(sSite);
					sbSql.Append("' AND PERSONID = ");
					sbSql.Append(nSubjLabel);

					// get data table
					DataTable dataTable2 = GetDataSet(_sMACROdbConn, sbSql.ToString());

					// log
					log.Debug( "sSql= " + sbSql.ToString() );

					// should be exactly 1 match
					nSubjectId = Convert.ToInt32(dataTable2.Rows[0]["PERSONID"]);
				}
			}
			else if( dataTable.Rows.Count > 1)
			{
				// more than 1 match - throw exception
				throw (new Exception( "More than 1 subject label match." ));
			}
			else
			{
				nSubjectId = Convert.ToInt32(dataTable.Rows[0]["PERSONID"]);
			}

			// check subjectid found
			if(nSubjectId == BufferAPI._DEFAULT_MISSING_NUMERIC)
			{
				throw ( new Exception( "No matching subject found." ) );
			}

			return nSubjectId;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="nStudyId"></param>
		/// <param name="sDataItemCode"></param>
		/// <returns></returns>
		public int GetDataItemId(int nStudyId, string sDataItemCode)
		{
			int nDataItemId = BufferAPI._DEFAULT_MISSING_NUMERIC;

			// sql
			StringBuilder sbSql = new StringBuilder();
			sbSql.Append( "SELECT DATAITEMID FROM DATAITEM WHERE CLINICALTRIALID = " );
			sbSql.Append( nStudyId );
			sbSql.Append( " AND DATAITEMCODE = '");
			sbSql.Append( sDataItemCode );
			sbSql.Append( "'" );

			// log
			log.Debug( "sSql= " + sbSql.ToString() );

			// get data table
			DataTable dataTable = GetDataSet(_sMACROdbConn, sbSql.ToString() );

			// will only be 1 matching row
			nDataItemId = Convert.ToInt32(dataTable.Rows[0]["DATAITEMID"]);

			return nDataItemId;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="nStudyId"></param>
		/// <param name="nDataItemId"></param>
		/// <returns></returns>
		public ResponseDataItem GetDataItemInfo(int nStudyId, int nDataItemId)
		{
			// get data item data type
			// sql
			StringBuilder sbSql = new StringBuilder();
			sbSql.Append( "SELECT DATAITEMCODE, DATATYPE, DATAITEMFORMAT, DATAITEMLENGTH FROM DATAITEM WHERE CLINICALTRIALID = " );
			sbSql.Append( nStudyId );
			sbSql.Append( " AND DATAITEMID = ");
			sbSql.Append( nDataItemId );

			// get data table
			DataTable dataTable = GetDataSet(_sMACROdbConn, sbSql.ToString() );

			// should be just 1 row
			ResponseDataItem rdi = new ResponseDataItem();
			rdi.Id = nDataItemId;
			rdi.Code = dataTable.Rows[0]["DATAITEMCODE"].ToString();
			rdi.MACRODataType = (BufferAPI.MACRODataTypes)Convert.ToInt16( dataTable.Rows[0]["DATATYPE"] );
			rdi.Format = (!Convert.IsDBNull(dataTable.Rows[0]["DATAITEMFORMAT"]))?dataTable.Rows[0]["DATAITEMFORMAT"].ToString():"";
			rdi.Length = Convert.ToInt16( dataTable.Rows[0]["DATAITEMLENGTH"] );

			if(rdi.MACRODataType == BufferAPI.MACRODataTypes.Category)
			{
				// categories 
				sbSql = new StringBuilder();
				sbSql.Append( "SELECT VALUECODE, ITEMVALUE, ACTIVE FROM VALUEDATA WHERE CLINICALTRIALID = " );
				sbSql.Append( nStudyId );
				sbSql.Append( " AND DATAITEMID = ");
				sbSql.Append( nDataItemId );

				// log
				log.Debug( "sSql= " + sbSql.ToString() );

				// get data table
				DataTable dataCatTable = GetDataSet(_sMACROdbConn, sbSql.ToString() );

				// loop through category rows
				foreach(DataRow dr in dataCatTable.Rows)
				{
					// use only active categories
					if( Convert.ToInt16(dr["ACTIVE"]) == 1)
					{
						// get code & value
						string sCatCode = dr["VALUECODE"].ToString();
						string sCatValue = dr["ITEMVALUE"].ToString();
						// create new category
						CategoryItem catItem = new CategoryItem(sCatCode, sCatValue);
						// add to responsedata item
						rdi.AddCategory( catItem );
					}
				}
			}

			return rdi;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="nStudyId"></param>
		/// <param name="sVisitCode"></param>
		/// <returns></returns>
		public int GetVisitId(int nStudyId, string sVisitCode)
		{
			int nVisitId = BufferAPI._DEFAULT_MISSING_NUMERIC;

			if(sVisitCode != "")
			{
				// sql
				StringBuilder sbSql = new StringBuilder();
				sbSql.Append( "SELECT VISITID FROM STUDYVISIT WHERE CLINICALTRIALID = " );
				sbSql.Append( nStudyId );
				sbSql.Append( " AND VISITCODE = '");
				sbSql.Append( sVisitCode );
				sbSql.Append( "'" );

				// log
				log.Debug( "sSql= " + sbSql.ToString() );

				// get data table
				DataTable dataTable = GetDataSet(_sMACROdbConn, sbSql.ToString() );

				// will only be 1 matching row - not a basic test requirement so catch error
				try
				{
					nVisitId = Convert.ToInt32(dataTable.Rows[0]["VISITID"]);
				}
				catch(Exception ex)
				{
					log.Info("Visit Id cannot be found.",ex);
				}
			}

			return nVisitId;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="nStudyId"></param>
		/// <param name="sEformCode"></param>
		/// <returns></returns>
		public int GetEformId(int nStudyId, string sEformCode)
		{
			int nEformId = BufferAPI._DEFAULT_MISSING_NUMERIC;

			if(sEformCode != "")
			{
				// sql
				StringBuilder sbSql = new StringBuilder();
				sbSql.Append( "SELECT CRFPAGEID FROM CRFPAGE WHERE CLINICALTRIALID = " );
				sbSql.Append( nStudyId );
				sbSql.Append( " AND CRFPAGECODE = '");
				sbSql.Append( sEformCode );
				sbSql.Append( "'" );

				// log
				log.Debug( "sSql= " + sbSql.ToString() );

				// get data table
				DataTable dataTable = GetDataSet(_sMACROdbConn, sbSql.ToString() );

				// will only be 1 matching row - not a basic test requirement so catch error
				try
				{
					nEformId = Convert.ToInt32(dataTable.Rows[0]["CRFPAGEID"]);
				}
				catch(Exception ex)
				{
					log.Info("eForm Id cannot be found.",ex);
				}
			}

			return nEformId;
		}

		/// <summary>
		/// get study information
		/// </summary>
		/// <param name="nStudyId"></param>
		/// <returns></returns>
		public StudyInfo GetStudyInfo(int nStudyId)
		{
			// create studyinfo object
			StudyInfo studyInfo = new StudyInfo();

			// Visit info
			StringBuilder sbSql = new StringBuilder();
			sbSql.Append( "SELECT VISITID, VISITCODE FROM STUDYVISIT WHERE CLINICALTRIALID = " );
			sbSql.Append( nStudyId );

			// get visit info data table
			DataTable dataTable = GetDataSet(_sMACROdbConn, sbSql.ToString() );
			foreach(DataRow dataRow in dataTable.Rows)
			{
				// get visit info
				int nVisitId = Convert.ToInt32( dataRow["VISITID"] );
				string sVisitCode = dataRow["VISITCODE"].ToString();

				// write to studyinfo object
				studyInfo.AddVisit(nVisitId, sVisitCode);
			}
			dataTable.Dispose();

			// eform info
			sbSql = new StringBuilder();
			sbSql.Append( "SELECT CRFPAGEID, CRFPAGECODE FROM CRFPAGE WHERE CLINICALTRIALID = " );
			sbSql.Append( nStudyId );

			// get eform info datatable
			dataTable = GetDataSet(_sMACROdbConn, sbSql.ToString() );
			foreach(DataRow dataRow in dataTable.Rows)
			{
				// get visit info
				int nEformId = Convert.ToInt32( dataRow["CRFPAGEID"] );
				string sEformCode = dataRow["CRFPAGECODE"].ToString();

				// write to studyinfo object
				studyInfo.AddEForm( nEformId, sEformCode);
			}
			dataTable.Dispose();

			// dataitem info
			sbSql = new StringBuilder();
			sbSql.Append( "SELECT DATAITEMID, DATAITEMCODE, DATATYPE, DATAITEMFORMAT, DATAITEMLENGTH FROM DATAITEM WHERE CLINICALTRIALID = " );
			sbSql.Append( nStudyId );

			// get dataitem info datatable
			dataTable = GetDataSet(_sMACROdbConn, sbSql.ToString() );

			// categories datatable
			sbSql = new StringBuilder();
			sbSql.Append( "SELECT VALUECODE, ITEMVALUE, DATAITEMID FROM VALUEDATA WHERE CLINICALTRIALID = " );
			sbSql.Append( nStudyId );
			sbSql.Append( " AND ACTIVE = 1 ");
			DataTable dataCategoryTable = GetDataSet(_sMACROdbConn, sbSql.ToString() );

			foreach(DataRow dataRow in dataTable.Rows)
			{
				// get data item info
				ResponseDataItem rdi = new ResponseDataItem(
					Convert.ToInt32( dataRow["DATAITEMID"] ),
					dataRow["DATAITEMCODE"].ToString(),
					(BufferAPI.MACRODataTypes)Convert.ToInt16( dataRow["DATATYPE"] ),
					(!Convert.IsDBNull(dataRow["DATAITEMFORMAT"]))?dataRow["DATAITEMFORMAT"].ToString():"",
					Convert.ToInt16( dataRow["DATAITEMLENGTH"] ) );

				// if a category question
				if(rdi.MACRODataType == BufferAPI.MACRODataTypes.Category)
				{
					// put together filter expression
					StringBuilder sbFilterExp = new StringBuilder();
					sbFilterExp.Append("DATAITEMID = ");
					sbFilterExp.Append(rdi.Id.ToString());
					string sSortBy = "DATAITEMID";
 
					// filter category types for this data item and add to definition
					foreach(DataRow drCat in dataCategoryTable.Select(sbFilterExp.ToString(),sSortBy))
					{
						// get code & value
						string sCatCode = drCat["VALUECODE"].ToString();
						string sCatValue = drCat["ITEMVALUE"].ToString();
						// create new category
						CategoryItem catItem = new CategoryItem(sCatCode, sCatValue);
						// add to responsedataitem
						rdi.AddCategory(catItem);
					}
				}

				// write to studyinfo object
				studyInfo.AddDataItem( rdi );
			}

			return studyInfo;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sBufferMessageXML"></param>
		/// <param name="subjDetail"></param>
		public void WriteBufferMessages(ref string sBufferMessageXML, ref SubjectDetails subjDetail)
		{
			// create guid for row identifier
			string sGuid = Guid.NewGuid().ToString();
			// store in subjectdetails
			subjDetail.Guid = sGuid;
			// create timestamp
			double dblResponseTimestamp = DateTime.Now.ToOADate();

			// need to perform separately for sql server and oracle due to binary type being used
			if(IMEDDataAccess.CalculateConnectionType(_sMACROdbConn)==IMEDDataAccess.ConnectionType.SQLServer)
			{
				// sql server
				SaveToSqlServerDb(ref sBufferMessageXML, ref subjDetail, sGuid, dblResponseTimestamp);
			}
			else
			{
				// oracle
				SaveToOracleDb(ref sBufferMessageXML, ref subjDetail, sGuid, dblResponseTimestamp);
			}

			// if no errors save to responsebufferdata tables
			if(subjDetail.SubjectResponseStatus == BufferAPI.BufferResponseStatus.Success)
			{
				SaveBufferResponses(ref subjDetail, sGuid, dblResponseTimestamp);
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sBufferMessageXML"></param>
		/// <param name="subjDetail"></param>
		/// <param name="sGuid"></param>
		/// <param name="dblResponseTimestamp"></param>
		private void SaveToSqlServerDb(ref string sBufferMessageXML, ref SubjectDetails subjDetail, string sGuid, double dblResponseTimestamp)
		{
			try
			{
				// declare connection / datareader interface
				IDbConnection dbConn = null;
				IDbCommand dbCommand = null;

				// open ImedDataAccess class
				IMEDDataAccess imedData = new IMEDDataAccess();
				IMEDDataAccess.ConnectionType imedConnType;

				// get connection type
				imedConnType = IMEDDataAccess.CalculateConnectionType(_sMACROdbConn);

				// get connection to db
				dbConn = imedData.GetConnection(imedConnType,_sMACROdbConn);
				
				// open connection
				dbConn.Open();

				// get command object 
				dbCommand = imedData.GetCommand(dbConn);

				// put together insert command
				// insert row but leave binary to be inserted by streaming
				StringBuilder sbInsertSql = new StringBuilder();
				// insert
				sbInsertSql.Append("INSERT INTO BUFFERRESPONSE (BUFFERRESPONSEID, BUFFERMESSAGE, BUFFERRESPONSESTATUS, BUFFERRESPONSESTATTIMESTAMP, ");
				sbInsertSql.Append("BUFFERRESPONSESTATTIMESTAMP_TZ, CLINICALTRIALNAME, CLINICALTRIALID, SITE, SUBJECTLABEL, PERSONID) VALUES (");
				// parameters
				// NOTE: FOR ORACLE use ':' instead of '@' in parameters
				sbInsertSql.Append("@bufferresponseid, @buffermessage, @bufferresponsestatus, @bufferresponsestattimestamp");
				sbInsertSql.Append(", @bufferresponsestattimestamp_tz, @clinicaltrialname, @clinicaltrialid, @site, @subjectlabel, @personid)");

				// add to command object
				dbCommand.CommandText = sbInsertSql.ToString();

				// now add parameter values
				// bufferresponseid
				IDbDataParameter dbParameter = imedData.AddDataParameter(ref dbCommand, "@bufferresponseid", DbType.AnsiString);
				dbParameter.Value = sGuid;
				dbParameter.Size = sGuid.Length;
				// buffermessage - enter dummy value
				// quirk with this method - can't update a NULL value
				dbParameter = imedData.AddDataParameter(ref dbCommand, "@buffermessage", DbType.String);
				dbParameter.Value = "x";
				dbParameter.Size = 1;
				//dbParameter.Size = byteDummy.Length;
				
				// bufferresponsestatus
				dbParameter = imedData.AddDataParameter(ref dbCommand, "@bufferresponsestatus", DbType.UInt16);
				dbParameter.Value = Convert.ToInt16(subjDetail.SubjectResponseStatus);
				dbParameter.Size = 2;
				// bufferresponsestattimestamp
				dbParameter = imedData.AddDataParameter(ref dbCommand, "@bufferresponsestattimestamp", DbType.Double);
				dbParameter.Value = dblResponseTimestamp;
				dbParameter.Size = 8;
				// bufferresponsestattimestamp_tz
				dbParameter = imedData.AddDataParameter(ref dbCommand, "@bufferresponsestattimestamp_tz", DbType.UInt16);
				dbParameter.Value = Convert.ToInt16(GetLocalMACROTimezone());
				dbParameter.Size = 2;
				// clinicaltrialname
				dbParameter = imedData.AddDataParameter(ref dbCommand, "@clinicaltrialname", DbType.AnsiString);
				if(subjDetail.ClinicalTrialName != "")
				{
					dbParameter.Value = subjDetail.ClinicalTrialName;
					dbParameter.Size = subjDetail.ClinicalTrialName.Length;
				}
				else
				{
					dbParameter.Value = System.DBNull.Value;
					dbParameter.Size = 0;
				}
				// clinicaltrialid
				dbParameter = imedData.AddDataParameter(ref dbCommand, "@clinicaltrialid", DbType.UInt32);
				if(subjDetail.ClinicalTrialId != BufferAPI._DEFAULT_MISSING_NUMERIC)
				{
					dbParameter.Value = subjDetail.ClinicalTrialId;
					dbParameter.Size = 4;
				}
				else
				{
					dbParameter.Value = System.DBNull.Value;
					dbParameter.Size = 0;
				}
				// site
				dbParameter = imedData.AddDataParameter(ref dbCommand, "@site", DbType.AnsiString);
				if(subjDetail.Site != "")
				{
					dbParameter.Value = subjDetail.Site;
					dbParameter.Size = subjDetail.Site.Length;
				}
				else
				{
					dbParameter.Value = System.DBNull.Value;
					dbParameter.Size = 0;
				}
				// subjectlabel
				dbParameter = imedData.AddDataParameter(ref dbCommand, "@subjectlabel", DbType.AnsiString);
				if(subjDetail.SubjectLabel != "")
				{
					dbParameter.Value = subjDetail.SubjectLabel;
					dbParameter.Size = subjDetail.SubjectLabel.Length;
				}
				else
				{
					dbParameter.Value = System.DBNull.Value;
					dbParameter.Size = 0;
				}
				// personid
				dbParameter = imedData.AddDataParameter(ref dbCommand, "@personid", DbType.UInt32);
				if(subjDetail.SubjectId != BufferAPI._DEFAULT_MISSING_NUMERIC)
				{
					dbParameter.Value = subjDetail.SubjectId;
					dbParameter.Size = 4;
				}
				else
				{
					dbParameter.Value = System.DBNull.Value;
					dbParameter.Size = 0;
				}

				// execute command
				dbCommand.ExecuteNonQuery();

				// now obtain reference to BUFFERMESSAGE field in table
				// refresh command object
				dbCommand = imedData.GetCommand(dbConn);

				// reference sql
				StringBuilder sbReferenceSql = new StringBuilder();
				sbReferenceSql.Append("SELECT @pointer=TEXTPTR(BUFFERMESSAGE) FROM BUFFERRESPONSE WHERE BUFFERRESPONSEID='");
				sbReferenceSql.Append(sGuid);
				sbReferenceSql.Append("'");

				// set command object with reference sql
				dbCommand.CommandText = sbReferenceSql.ToString();

				// now add parameter detail to command
				// attempting to obtain a pointer to the binary field
				SqlParameter sqlOutPointerParam = ((SqlCommand)dbCommand).Parameters.Add("@pointer", SqlDbType.VarBinary, 100);
				sqlOutPointerParam.Direction = ParameterDirection.Output;
				// execute command
				dbCommand.ExecuteNonQuery();

				// refresh command object
				dbCommand = imedData.GetCommand(dbConn);

				// updatetext command sql (way to stream binary fields)
				string sUpdateTextSql = "UPDATETEXT BUFFERRESPONSE.BUFFERMESSAGE @pointer @offset @delete WITH LOG @xmldata";

				// add to command object
				dbCommand.CommandText = sUpdateTextSql;

				// now add parameter values to command
				// pointer
				IDbDataParameter dbPointerParameter = imedData.AddDataParameter(ref dbCommand, "@pointer", DbType.Binary);
				dbPointerParameter.Size = 16;
				// offset
				IDbDataParameter dbOffsetParameter = imedData.AddDataParameter(ref dbCommand, "@offset", DbType.Int32);
				// delete - to delete existing dummy character of length 1
				IDbDataParameter dbDeleteParameter = imedData.AddDataParameter(ref dbCommand, "@delete", DbType.Int16);
				dbDeleteParameter.Value = 1;
				// xmldata
				IDbDataParameter dbXmlStringParameter = imedData.AddDataParameter(ref dbCommand, "@xmldata", DbType.String);
				dbXmlStringParameter.Size = _BUFFERSIZE;

				// set up string array position
				int nBufferPosition = 0;
				// set up completed flag
				bool bCompleteRead = false;
				// continue reading string until have reached end of file
				while(!bCompleteRead)
				{
					// set offset parameter (position to set string from)
					dbOffsetParameter.Value = nBufferPosition;
					// read part of string at a time to send in chunks
					int nBufferRead = _BUFFERSIZE;
					// calculate next buffer block size
					if((nBufferPosition + _BUFFERSIZE) > sBufferMessageXML.Length)
					{
						// check if read is complete
						if(nBufferPosition == sBufferMessageXML.Length)
						{
							nBufferRead = 0;
						}
						else
						{
							nBufferRead = sBufferMessageXML.Length - nBufferPosition;
						}
					}

					// if have retrieved bytes need to store in db
					if(nBufferRead > 0)
					{
						// get buffer string
						string sBuffer = sBufferMessageXML.Substring(nBufferPosition, nBufferRead);
						// write to db
						// pointerparam
						dbPointerParameter.Value = sqlOutPointerParam.Value;
						// fill string parameter with contents of buffer string
						dbXmlStringParameter.Value = sBuffer;
						// execute command
						dbCommand.ExecuteNonQuery();

						// set delete parameter not to delete again
						dbDeleteParameter.Value = 0;

						// update position counter
						nBufferPosition += nBufferRead;
					}
					else
					{
						// have read last block of bytes
						bCompleteRead = true;
					}
				}

				// close connection
				dbConn.Close();
			}
			catch(Exception ex)
			{
				// store error in log
				log.Error("Error saving to BufferResponse table.", ex);
				// throw again
				throw ( new Exception( "Error saving to BufferResponse table.", ex) );
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sBufferMessageXML"></param>
		/// <param name="subjDetail"></param>
		/// <param name="sGuid"></param>
		/// <param name="dblResponseTimestamp"></param>
		/// <returns></returns>
		private bool SaveToOracleDb(ref string sBufferMessageXML, ref SubjectDetails subjDetail, string sGuid, double dblResponseTimestamp)
		{
			try
			{
				OracleConnection connection = new OracleConnection(IMEDDataAccess.FormatConnectionString(IMEDDataAccess.ConnectionType.Oracle, _sMACROdbConn));
				connection.Open();

				// Declare Oracle objects
				// transaction object
				OracleTransaction txn;
				// command
				OracleCommand cmd = new OracleCommand("",connection);

				// Start a transaction
				txn = connection.BeginTransaction();

				// set command text to store temporary nclob in oracle db server memory
				cmd.CommandText = "declare xx nclob; begin dbms_lob.createtemporary(xx, false, 0); :tempclob:= xx; end;";
			
				// command has transaction
				cmd.Transaction = txn;

				// create OracleParameter for command
				OracleParameter opClob = cmd.Parameters.Add("tempclob", OracleType.NClob);
				// getting a reference to the blob through the output parameter
				opClob.Direction = ParameterDirection.Output;
	
				// execute - create temporary workspace on server
				cmd.ExecuteNonQuery();

				// get temporary lob reference
				OracleLob oracTempLob = (OracleLob)cmd.Parameters[0].Value;
				// begin batch transaction
				oracTempLob.BeginBatch(OracleLobOpenMode.ReadWrite);

				// set up string position
				int nBufferPosition = 0;
				// set up completed flag
				bool bCompleteRead = false;
				// continue reading string until have reached end 
				while(!bCompleteRead)
				{
					// read part of string at a time to send in chunks
					int nBufferRead = _BUFFERSIZE;
					// calculate next buffer block size
					if((nBufferPosition + _BUFFERSIZE) > sBufferMessageXML.Length)
					{
						// check if read is complete
						if(nBufferPosition == sBufferMessageXML.Length)
						{
							nBufferRead = 0;
						}
						else
						{
							nBufferRead = sBufferMessageXML.Length - nBufferPosition;
						}
					}

					// if have retrieved bytes need to store in db
					if(nBufferRead > 0)
					{
						// get buffer string
						string sBuffer = sBufferMessageXML.Substring(nBufferPosition, nBufferRead);
						
						byte[] byteBuffer = Encoding.Unicode.GetBytes(sBuffer);

						// write to blob on server within transaction
						oracTempLob.Write(byteBuffer, 0, byteBuffer.Length);
	
						// update position counter
						nBufferPosition += nBufferRead;
					}
					else
					{
						// have read last block of bytes
						bCompleteRead = true;
					}

				}

				// commit batch
				oracTempLob.EndBatch();
				// remove parameters object from command
				cmd.Parameters.Clear();

				// ready for stored procedure to insert all detail to BUFFERRESPONSE table
				cmd.CommandText = "INSERTBUFFERMESSAGE";
				cmd.CommandType = CommandType.StoredProcedure;
				// add parameters
				// bufferresponseid
				OracleParameter oracGuidParam = cmd.Parameters.Add("bufferresponseid", OracleType.VarChar);
				oracGuidParam.Value = sGuid;
				oracGuidParam.Size = sGuid.Length;
				// buffermessage (nclob) - set value equal to that of OracleLob placed temporarily on server
				OracleParameter oracBinaryParam = cmd.Parameters.Add("buffermessage", OracleType.NClob);
				oracBinaryParam.Value = oracTempLob;
				// bufferresponsestatus
				OracleParameter oracStatParam = cmd.Parameters.Add("bufferresponsestatus", OracleType.Int16);
				oracStatParam.Value = Convert.ToInt16(subjDetail.SubjectResponseStatus);
				oracStatParam.Size = 2;
				// bufferresponsestattimestamp
				OracleParameter oracTimeParam = cmd.Parameters.Add("bufferresponsestattimestamp", OracleType.Double);
				oracTimeParam.Value = dblResponseTimestamp;
				oracTimeParam.Size = 8;
				// bufferresponsestattimestamp_tz
				OracleParameter oracTimeZoneParam = cmd.Parameters.Add("bufferresponsestattimestamp_tz", OracleType.Int16);
				oracTimeZoneParam.Value = GetLocalMACROTimezone();
				oracTimeZoneParam.Size = 2;
				// clinicaltrialname
				OracleParameter oracTrialNameParam = cmd.Parameters.Add("clinicaltrialname", OracleType.VarChar);
				if(subjDetail.ClinicalTrialName != "")
				{
					oracTrialNameParam.Value = subjDetail.ClinicalTrialName;
					oracTrialNameParam.Size = subjDetail.ClinicalTrialName.Length;
				}
				else
				{
					oracTrialNameParam.Value = System.DBNull.Value;
					oracTrialNameParam.Size = 0;
				}
				// clinicaltrialid
				OracleParameter oracTrialIdParam = cmd.Parameters.Add("clinicaltrialid", OracleType.Int32);
				if( subjDetail.ClinicalTrialId != BufferAPI._DEFAULT_MISSING_NUMERIC )
				{
					oracTrialIdParam.Value = Convert.ToInt32(subjDetail.ClinicalTrialId);
					oracTrialIdParam.Size = 4;
				}
				else
				{
					oracTrialIdParam.Value = System.DBNull.Value;
					oracTrialIdParam.Size = 0;
				}
				// site
				OracleParameter oracSiteParam = cmd.Parameters.Add("site", OracleType.VarChar);
				if(subjDetail.Site != "")
				{
					oracSiteParam.Value = subjDetail.Site;
					oracSiteParam.Size = subjDetail.Site.Length;
				}
				else
				{
					oracSiteParam.Value = System.DBNull.Value;
					oracSiteParam.Size = 0;
				}
				// subjectlabel
				OracleParameter oracSubjLabelParam = cmd.Parameters.Add("subjectlabel", OracleType.VarChar);
				if(subjDetail.SubjectLabel != "")
				{
					oracSubjLabelParam.Value = subjDetail.SubjectLabel;
					oracSubjLabelParam.Size = subjDetail.SubjectLabel.Length;
				}
				else
				{
					oracSubjLabelParam.Value = System.DBNull.Value;
					oracSubjLabelParam.Size = 0;
				}
				// personid
				OracleParameter oracPersonIdParam = cmd.Parameters.Add("personid", OracleType.Int32);
				if( subjDetail.SubjectId != BufferAPI._DEFAULT_MISSING_NUMERIC )
				{
					oracPersonIdParam.Value = Convert.ToInt32(subjDetail.SubjectId);
					oracPersonIdParam.Size = 4;
				}
				else
				{
					oracPersonIdParam.Value = System.DBNull.Value;
					oracPersonIdParam.Size = 0;
				}

				// execute stored procedure
				cmd.ExecuteNonQuery();

				// commit transaction
				txn.Commit();

				// close objects
				connection.Close();
			}
			catch(Exception ex)
			{
				// store error in log
				log.Error("Error saving to BufferResponse table.", ex);
				// throw again
				throw ( new Exception( "Error saving to BufferResponse table.", ex) );
			}
			return true;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="subjDetail"></param>
		/// <param name="sGuid"></param>
		/// <param name="dblResponseTimestamp"></param>
		// REVISIONS:
		// DPH 21/11/2007 - Standardize double dates
		private void SaveBufferResponses(ref SubjectDetails subjDetail, string sGuid, double dblResponseTimestamp)
		{
			try
			{
				// loop through responses to save and store on db
				// declare connection / datareader interface
				IDbConnection dbConn = null;
				IDbCommand dbCommand = null;

				// open ImedDataAccess class
				IMEDDataAccess imedData = new IMEDDataAccess();
				IMEDDataAccess.ConnectionType imedConnType;

				// get connection type
				imedConnType = IMEDDataAccess.CalculateConnectionType(_sMACROdbConn);

				// get connection to db
				dbConn = imedData.GetConnection(imedConnType,_sMACROdbConn);
					
				// open connection
				dbConn.Open();

				// get command object 
				dbCommand = imedData.GetCommand(dbConn);

				foreach(ResponseDetails responseDet in subjDetail.Responses)
				{
					// create insert sql
					StringBuilder sbSql = new StringBuilder();

					// bufferresponsedata
					sbSql.Append("INSERT INTO BUFFERRESPONSEDATA (BUFFERRESPONSEDATAID, BUFFERRESPONSEID, BUFFERCOMMITSTATUS, BUFFERCOMMITSTATUSTIMESTAMP, BUFFERCOMMITSTATUSTIMESTAMP_TZ, ");
					sbSql.Append("CLINICALTRIALNAME, CLINICALTRIALID, SITE, SUBJECTLABEL, PERSONID, VISITCODE, VISITID, VISITCYCLENUMBER, CRFPAGECODE, CRFPAGEID, CRFPAGECYCLENUMBER, ");
					sbSql.Append("DATAITEMCODE, DATAITEMID, RESPONSEREPEATNUMBER, RESPONSEVALUE, ORDERDATETIME, USERNAME, USERNAMEFULL) VALUES ('");
					//BUFFERRESPONSEDATAID - new guid
					string sResponseGuid = Guid.NewGuid().ToString();
					responseDet.Guid = sResponseGuid;
					sbSql.Append(sResponseGuid);
					sbSql.Append("','");
					//BUFFERRESPONSEID - link to complete buffermessage
					sbSql.Append(sGuid);
					sbSql.Append("',");
					//BUFFERCOMMITSTATUS
					sbSql.Append((int)BufferAPI.BufferCommitStatus.NotCommitted);
					sbSql.Append(",");
					//BUFFERCOMMITSTATUSTIMESTAMP
					// DPH 21/11/2007 - Convert double to Standardized format for SQL
					sbSql.Append(IMEDFunctions20.LocalNumToStandard( dblResponseTimestamp.ToString(), false));
					sbSql.Append(",");
					//BUFFERCOMMITSTATUSTIMESTAMP_TZ
					sbSql.Append(GetLocalMACROTimezone());
					sbSql.Append(",'");
					//CLINICALTRIALNAME
					sbSql.Append(subjDetail.ClinicalTrialName);
					sbSql.Append("',");
					//CLINICALTRIALID
					sbSql.Append(subjDetail.ClinicalTrialId);
					sbSql.Append(",'");
					//SITE
					sbSql.Append(subjDetail.Site);
					sbSql.Append("','");
					//SUBJECTLABEL
					sbSql.Append(subjDetail.SubjectLabel);
					sbSql.Append("',");
					//PERSONID
					sbSql.Append(subjDetail.SubjectId);
					sbSql.Append(",");
					//VISITCODE
					if(responseDet.VisitCode != "")
					{
						sbSql.Append("'");
						sbSql.Append(responseDet.VisitCode);
						sbSql.Append("'");
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//VISITID
					if(responseDet.VisitId != BufferAPI._DEFAULT_MISSING_NUMERIC)
					{
						sbSql.Append(responseDet.VisitId);
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//VISITCYCLENUMBER
					if(responseDet.VisitCycle != BufferAPI._DEFAULT_MISSING_NUMERIC)
					{
						sbSql.Append(responseDet.VisitCycle);
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//CRFPAGECODE
					if(responseDet.CRFPageCode != "")
					{
						sbSql.Append("'");
						sbSql.Append(responseDet.CRFPageCode);
						sbSql.Append("'");
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//CRFPAGEID
					if(responseDet.CRFPageId != BufferAPI._DEFAULT_MISSING_NUMERIC)
					{
						sbSql.Append(responseDet.CRFPageId);
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//CRFPAGECYCLENUMBER
					if(responseDet.CRFPageCycle != BufferAPI._DEFAULT_MISSING_NUMERIC)
					{
						sbSql.Append(responseDet.CRFPageCycle);
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//DATAITEMCODE
					if(responseDet.DataItemCode != "")
					{
						sbSql.Append("'");
						sbSql.Append(responseDet.DataItemCode);
						sbSql.Append("'");
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//DATAITEMID
					if(responseDet.DataItemId != BufferAPI._DEFAULT_MISSING_NUMERIC)
					{
						sbSql.Append(responseDet.DataItemId);
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//RESPONSEREPEATNUMBER
					if(responseDet.DataItemRepeatNo != BufferAPI._DEFAULT_MISSING_NUMERIC)
					{
						sbSql.Append(responseDet.DataItemRepeatNo);
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//RESPONSEVALUE
					if(responseDet.ResponseValue != "")
					{
						sbSql.Append("'");
						sbSql.Append(FormatMACROResponseData(subjDetail.ClinicalTrialId,responseDet.DataItemId,responseDet.ResponseValue));
						sbSql.Append("'");
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					// ORDERDATETIME
					if(responseDet.DataCompareDate != DateTime.FromOADate(0))
					{
						// DPH 21/11/2007 - Convert double to Standardized format for SQL
						sbSql.Append(IMEDFunctions20.LocalNumToStandard(responseDet.DataCompareDate.ToOADate().ToString(),false));
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//USERNAME + USERNAMEFULL
					sbSql.Append("null,null)");

					// set command sql - BufferResponseData
					dbCommand.CommandText = sbSql.ToString();
					// log
					log.Debug(sbSql.ToString());
					// execute command - BufferResponseData
					dbCommand.ExecuteNonQuery();

					// change sql to also write to bufferresponsedatahistory table
					sbSql.Replace("INSERT INTO BUFFERRESPONSEDATA", "INSERT INTO BUFFERRESPONSEDATAHISTORY");

					// set command sql - BufferResponseDataHistory
					dbCommand.CommandText = sbSql.ToString();
					// execute command - BufferResponseDataHistory
					dbCommand.ExecuteNonQuery();
				}

				// dispose command
				dbCommand.Dispose();

				// close connection
				dbConn.Close();

				// dispose connection
				dbConn.Dispose();

			}
			catch(Exception ex)
			{
				// store error in log
				log.Error("Error saving to BufferResponseData table.", ex);
				// throw again
				throw ( new Exception( "Error saving to BufferResponsedata table.", ex) );
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="nStudyId"></param>
		/// <param name="nDataItemId"></param>
		/// <param name="sDataItemValue"></param>
		/// <returns></returns>
		private string FormatMACROResponseData(int nStudyId, int nDataItemId, string sDataItemValue)
		{
			// Get DataItem info from Db
			ResponseDataItem responseDataItem = this.GetDataItemInfo(nStudyId, nDataItemId);

			// obtain Formatted Value
			string sFormattedValue = sDataItemValue;

			// if a format exists
			if(responseDataItem.Format != "")
			{
				switch(responseDataItem.MACRODataType)
				{
					case BufferAPI.MACRODataTypes.Text:
					case BufferAPI.MACRODataTypes.Category:
					case BufferAPI.MACRODataTypes.Multimedia:
					case BufferAPI.MACRODataTypes.Thesaurus:
					{
						// store as is
						break;
					}
					case BufferAPI.MACRODataTypes.Date:
					{
						// DPH 22/11/2007 - set culture as local culture is disrupting output of DateTime.ParseExact
						IFormatProvider culture = new System.Globalization.CultureInfo("en-GB", true);
						// pass in GB culture (as expect dd/MM/yyyy format)
						// read date into datetime
						DateTime dtDate = DateTime.ParseExact(sDataItemValue, BufferBasicChecks.DateFormatString(sDataItemValue), culture); 

						// convert MACRO date into microsoft format date
						string sFormatDate = BufferBasicChecks.FormatMACROToMicrosoftDate(responseDataItem.Format);

						// write date using new format
						sFormattedValue = dtDate.ToString(sFormatDate);

						break;
					}
					case BufferAPI.MACRODataTypes.IntegerData:
					case BufferAPI.MACRODataTypes.LabTest:
					case BufferAPI.MACRODataTypes.Real:
					{
						// convert into a double
						double dblResponseValue = Convert.ToDouble(sDataItemValue);

						//  modify MACRO numeric format to Microsoft format - 9's to 0's
						string sFormatNumeric = responseDataItem.Format.Replace("9", "0");

						// convert to a string
						sFormattedValue = dblResponseValue.ToString(sFormatNumeric);

						break;
					}
				}
			}

			return sFormattedValue;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="subjDetail"></param>
		public void UpdateBufferResponses(ref SubjectDetails subjDetail)
		{
			// loop through responses and set relevant response
			// BufferAPI.BufferCommitStatus
			try
			{
				// loop through responses to save and store on db
				// declare connection / datareader interface
				IDbConnection dbConn = null;
				IDbCommand dbCommand = null;

				// open ImedDataAccess class
				IMEDDataAccess imedData = new IMEDDataAccess();
				IMEDDataAccess.ConnectionType imedConnType;

				// get connection type
				imedConnType = IMEDDataAccess.CalculateConnectionType(_sMACROdbConn);

				// get connection to db
				dbConn = imedData.GetConnection(imedConnType,_sMACROdbConn);
					
				// open connection
				dbConn.Open();

				// get command object 
				dbCommand = imedData.GetCommand(dbConn);

				foreach(ResponseDetails responseDet in subjDetail.Responses)
				{
					// create update sql
					StringBuilder sbUpdateSql = new StringBuilder();
					sbUpdateSql.Append( "UPDATE BUFFERRESPONSEDATA SET BUFFERCOMMITSTATUS = " );
					// BUFFERCOMMITSTATUS
					sbUpdateSql.Append( (int)responseDet.BufferCommitStatus );
					// BUFFERCOMMITSTATUSTIMESTAMP
					double dblResponseTimestamp = DateTime.Now.ToOADate();
					sbUpdateSql.Append( " , BUFFERCOMMITSTATUSTIMESTAMP = " );
					sbUpdateSql.Append( dblResponseTimestamp );
					// BUFFERCOMMITSTATUSTIMESTAMP_TZ
					sbUpdateSql.Append( " , BUFFERCOMMITSTATUSTIMESTAMP_TZ = " );
					sbUpdateSql.Append( GetLocalMACROTimezone() );
					// USERNAME
					if(responseDet.Username != "")
					{
						sbUpdateSql.Append ( " , USERNAME = '" );
						sbUpdateSql.Append( responseDet.Username );
						sbUpdateSql.Append ( "'" );
					}
					// USERNAMEFULL
					if(responseDet.UsernameFull != "")
					{	
						sbUpdateSql.Append( " , USERNAMEFULL = '" );
						sbUpdateSql.Append( responseDet.UsernameFull );
						sbUpdateSql.Append ( "'" );
					}
					// WHERE ...
					// BUFFERRESPONSEDATAID
					sbUpdateSql.Append( " WHERE BUFFERRESPONSEDATAID = '" );
					sbUpdateSql.Append( responseDet.Guid );
					// BUFFERRESPONSEID
					sbUpdateSql.Append( "' AND BUFFERRESPONSEID = '" );
					sbUpdateSql.Append( subjDetail.Guid );
					sbUpdateSql.Append( "'" );

					// create insert sql
					StringBuilder sbSql = new StringBuilder();

					// bufferresponsedatahistory
					sbSql.Append("INSERT INTO BUFFERRESPONSEDATAHISTORY (BUFFERRESPONSEDATAID, BUFFERRESPONSEID, BUFFERCOMMITSTATUS, BUFFERCOMMITSTATUSTIMESTAMP, BUFFERCOMMITSTATUSTIMESTAMP_TZ, ");
					sbSql.Append("CLINICALTRIALNAME, CLINICALTRIALID, SITE, SUBJECTLABEL, PERSONID, VISITCODE, VISITID, VISITCYCLENUMBER, CRFPAGECODE, CRFPAGEID, CRFPAGECYCLENUMBER, ");
					sbSql.Append("DATAITEMCODE, DATAITEMID, RESPONSEREPEATNUMBER, RESPONSEVALUE, ORDERDATETIME, USERNAME, USERNAMEFULL) VALUES ('");
					//BUFFERRESPONSEDATAID 
					sbSql.Append(responseDet.Guid);
					sbSql.Append("','");
					//BUFFERRESPONSEID - link to complete buffermessage
					sbSql.Append(subjDetail.Guid);
					sbSql.Append("',");
					//BUFFERCOMMITSTATUS
					sbSql.Append((int)responseDet.BufferCommitStatus);
					sbSql.Append(",");
					//BUFFERCOMMITSTATUSTIMESTAMP
					sbSql.Append(dblResponseTimestamp.ToString());
					sbSql.Append(",");
					//BUFFERCOMMITSTATUSTIMESTAMP_TZ
					sbSql.Append(GetLocalMACROTimezone());
					sbSql.Append(",'");
					//CLINICALTRIALNAME
					sbSql.Append(subjDetail.ClinicalTrialName);
					sbSql.Append("',");
					//CLINICALTRIALID
					sbSql.Append(subjDetail.ClinicalTrialId);
					sbSql.Append(",'");
					//SITE
					sbSql.Append(subjDetail.Site);
					sbSql.Append("','");
					//SUBJECTLABEL
					sbSql.Append(subjDetail.SubjectLabel);
					sbSql.Append("',");
					//PERSONID
					sbSql.Append(subjDetail.SubjectId);
					sbSql.Append(",");
					//VISITCODE
					if(responseDet.VisitCode != "")
					{
						sbSql.Append("'");
						sbSql.Append(responseDet.VisitCode);
						sbSql.Append("'");
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//VISITID
					if(responseDet.VisitId != BufferAPI._DEFAULT_MISSING_NUMERIC)
					{
						sbSql.Append(responseDet.VisitId);
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//VISITCYCLENUMBER
					if(responseDet.VisitCycle != BufferAPI._DEFAULT_MISSING_NUMERIC)
					{
						sbSql.Append(responseDet.VisitCycle);
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//CRFPAGECODE
					if(responseDet.CRFPageCode != "")
					{
						sbSql.Append("'");
						sbSql.Append(responseDet.CRFPageCode);
						sbSql.Append("'");
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//CRFPAGEID
					if(responseDet.CRFPageId != BufferAPI._DEFAULT_MISSING_NUMERIC)
					{
						sbSql.Append(responseDet.CRFPageId);
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//CRFPAGECYCLENUMBER
					if(responseDet.CRFPageCycle != BufferAPI._DEFAULT_MISSING_NUMERIC)
					{
						sbSql.Append(responseDet.CRFPageCycle);
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//DATAITEMCODE
					if(responseDet.DataItemCode != "")
					{
						sbSql.Append("'");
						sbSql.Append(responseDet.DataItemCode);
						sbSql.Append("'");
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//DATAITEMID
					if(responseDet.DataItemId != BufferAPI._DEFAULT_MISSING_NUMERIC)
					{
						sbSql.Append(responseDet.DataItemId);
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//RESPONSEREPEATNUMBER
					if(responseDet.DataItemRepeatNo != BufferAPI._DEFAULT_MISSING_NUMERIC)
					{
						sbSql.Append(responseDet.DataItemRepeatNo);
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//RESPONSEVALUE
					if(responseDet.ResponseValue != "")
					{
						sbSql.Append("'");
						sbSql.Append(FormatMACROResponseData(subjDetail.ClinicalTrialId,responseDet.DataItemId,responseDet.ResponseValue));
						sbSql.Append("'");
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					// ORDERDATETIME
					if(responseDet.DataCompareDate != DateTime.FromOADate(0))
					{
						sbSql.Append(responseDet.DataCompareDate.ToOADate());
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					//USERNAME 
					if(responseDet.Username != "")
					{
						sbSql.Append("'");
						sbSql.Append( responseDet.Username );
						sbSql.Append("'");
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(",");
					// USERNAMEFULL
					if(responseDet.UsernameFull != "")
					{
						sbSql.Append("'");
						sbSql.Append( responseDet.UsernameFull );
						sbSql.Append("'");
					}
					else
					{
						sbSql.Append("null");
					}
					sbSql.Append(")");

					// set command sql - BufferResponseData
					dbCommand.CommandText = sbUpdateSql.ToString();
					// log
					log.Debug(sbUpdateSql.ToString());
					// execute command - BufferResponseData
					dbCommand.ExecuteNonQuery();

					// set command sql - BufferResponseDataHistory
					dbCommand.CommandText = sbSql.ToString();
					// log
					log.Debug(sbSql.ToString());
					// execute command - BufferResponseDataHistory
					dbCommand.ExecuteNonQuery();
				}

				// dispose command
				dbCommand.Dispose();

				// close connection
				dbConn.Close();

				// dispose connection
				dbConn.Dispose();

			}
			catch(Exception ex)
			{
				// store error in log
				log.Error("Error saving to BufferResponseData table.", ex);
				// throw again
				throw ( new Exception( "Error saving to BufferResponsedata table.", ex) );
			}
		}

		/// <summary>
		/// Gets local Timezone in MACRO format
		/// </summary>
		/// <returns></returns>
		private static int GetLocalMACROTimezone()
		{
			// get local timezone
			TimeZone tzLocal = TimeZone.CurrentTimeZone;
			// get timespan of UTC offset
			TimeSpan tsLocal = tzLocal.GetUtcOffset(DateTime.Now);
			// calculate offset MACRO style - offset in minutes UTC (GMT)
			int nOffset = -1 * Convert.ToInt32(tsLocal.TotalMinutes);
			return nOffset;
		}

		/// <summary>
		/// check if tables exist
		/// </summary>
		/// <returns></returns>
		public bool DoBufferAPITablesExist()
		{
			bool bExist=true;
			StringBuilder sbSQL = new StringBuilder();

			// declare connection interface
			IDbConnection dbConn = null;
			IDbCommand dbCommand = null;
			IDbDataAdapter dbDataAdapter = null;
			DataSet dataDB = new DataSet();

			try
			{
				// open ImedDataAccess class
				IMEDDataAccess imedData = new IMEDDataAccess();
				IMEDDataAccess.ConnectionType imedConnType;

				// calculate connection type
				imedConnType=IMEDDataAccess.CalculateConnectionType(_sMACROdbConn);

				// get connection
				dbConn = imedData.GetConnection(imedConnType,_sMACROdbConn);

				// open connection
				dbConn.Open();

				// get command object
				dbCommand = imedData.GetCommand(dbConn);
				// set command text
				sbSQL.Append("SELECT TABLE_NAME FROM ");
				if(imedConnType==IMEDDataAccess.ConnectionType.SQLServer)
				{
					// sql server
					sbSQL.Append("INFORMATION_SCHEMA.Tables ");
				}
				else
				{
					// oracle
					sbSQL.Append("USER_TABLES ");
				}
				sbSQL.Append("WHERE TABLE_NAME = 'BUFFERRESPONSE'");
				dbCommand.CommandText=sbSQL.ToString();

				// get data adaptor
				dbDataAdapter = imedData.GetDataAdapter(dbCommand);

				// fill dataset
				dbDataAdapter.Fill(dataDB);

				// get datatable
				DataTable dataTable = dataDB.Tables[0];

				// check if tables exist
				if(dataTable.Rows.Count == 0)
				{
					bExist=false;
				}
			}
			catch
			{
				return false;
			}
			finally
			{
				// in case of db problem
				// close db
				dbConn.Close();
			}

			// return
			return bExist;
		}

		/// <summary>
		/// create tables in connected to database
		/// </summary>
		public bool CreateBufferAPITables()
		{
			// created tables successfully
			bool bOk = true;

			// declare connection interface
			IDbConnection dbConn = null;
			IDbCommand dbCommand = null;

			try
			{
				// open ImedDataAccess class
				IMEDDataAccess imedData = new IMEDDataAccess();
				IMEDDataAccess.ConnectionType imedConnType;

				// calculate connection type
				imedConnType=IMEDDataAccess.CalculateConnectionType(_sMACROdbConn);

				// get connection
				dbConn = imedData.GetConnection(imedConnType,_sMACROdbConn);

				// open connection
				dbConn.Open();

				// get command object
				dbCommand = imedData.GetCommand(dbConn);

				// retrieve SQL table script
				string sSQLFile="";

				if(imedConnType==IMEDDataAccess.ConnectionType.SQLServer)
				{
					// SQL Server
					sSQLFile=IMEDFunctions20.ExtractAllTextFromFile(IMEDFunctions20.GetAppPath() + "MSSQL Buffer API tables.sql");
				}
				else
				{
					// Oracle
					sSQLFile=IMEDFunctions20.ExtractAllTextFromFile(IMEDFunctions20.GetAppPath() + "ORA Buffer API Tables.sql");
				}

				// if no sql returned then failed
				if(sSQLFile=="")
				{
					bOk = false;
				}

				// split file and execute a row at a time
				char[] chCrLf = {System.Convert.ToChar("\r"), System.Convert.ToChar("\n")};
				string[] aSql = sSQLFile.Split(chCrLf);

				// loop through and execute a line at a time
				foreach(string sSql in aSql)
				{
					if(sSql!="")
					{
						// set text
						dbCommand.CommandText=sSql;
						// execute
						dbCommand.ExecuteNonQuery();
					}
				}
			}
			catch(Exception ex)
			{
				bOk = false;
				throw ( new Exception( ex.Message ) );
			}
			finally
			{
				dbConn.Close();
			}

			return bOk;
		}
	}
}
