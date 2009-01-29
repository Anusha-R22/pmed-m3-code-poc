using System;
using System.Data;
using System.Collections;
using InferMed.Components;
using log4net;

namespace InferMed.MACRO.ClinicalCoding.MACROCCBS30
{
	/// <summary>
	/// Data access for clinical coding business services
	/// </summary>
	public class CCDataAccess
	{
		//logging
		private static readonly ILog log = LogManager.GetLogger( typeof( CCDataAccess ) );

		/// <summary>
		/// Constructor
		/// </summary>
		private CCDataAccess()
		{
		}

		/// <summary>
		/// Convert a column value from null
		/// </summary>
		/// <param name="r"></param>
		/// <param name="col"></param>
		/// <returns></returns>
		public static string ConvertFromNull( DataRow r, string col )
		{
			return( ( ( r.IsNull( col ) ) || ( r[col] == System.DBNull.Value ) ) ? "" : r[col].ToString() );
		}

		/// <summary>
		/// Prepare a string for inclusion in an sql string
		/// </summary>
		/// <param name="term"></param>
		/// <returns></returns>
		public static string SQLTerm( string term )
		{
			return( term.Replace( "'", "''" ) );
		}

		/// <summary>
		/// Returns a dataset from the database
		/// </summary>
		/// <param name="connectionString"></param>
		/// <param name="sql"></param>
		/// <returns></returns>
		public static DataSet GetDataSet(string connectionString, string sql)
		{
			IMEDDataAccess imedData = new IMEDDataAccess();
			IDbConnection dbConn = null;
			
			try
			{
				// open ImedDataAccess class
				IMEDDataAccess.ConnectionType imedConnType;
				// calculate connection type
				imedConnType=IMEDDataAccess.CalculateConnectionType( connectionString );
				// get connection & open
				dbConn = imedData.GetConnection( imedConnType, connectionString );
				dbConn.Open();
				
				return( GetDataSet( dbConn, sql ) );
			}
			finally
			{
				//clean up objects
				dbConn.Close();
				dbConn.Dispose();
			}
		}

		/// <summary>
		/// Returns a dataset from the database
		/// </summary>
		/// <param name="dbConn">Database connection object</param>
		/// <param name="sql">SQL to run against the MACRO database</param>
		/// <returns>Dataset containing MACRO data</returns>
		public static DataSet GetDataSet(IDbConnection dbConn, string sql)
		{
			IMEDDataAccess imedData = new IMEDDataAccess();
			IDbCommand dbCommand = null;
			IDbDataAdapter dbDataAdapter = null;
			DataSet dataDB = new DataSet();

			try
			{
				// get command
				dbCommand = imedData.GetCommand( dbConn );
				// create command text
				dbCommand.CommandText = sql;
				// get data adaptor
				dbDataAdapter = imedData.GetDataAdapter( dbCommand );
				// fill dataset
				dbDataAdapter.Fill( dataDB );
				
				return( dataDB );
			}
			finally
			{
				//clean up objects
				dbCommand.Dispose();
			}
		}

		/// <summary>
		/// Run an sql string against a database
		/// </summary>
		/// <param name="connectionString"></param>
		/// <param name="sql"></param>
		public static void RunSQL( string connectionString, string sql )
		{
			IMEDDataAccess imedData = new IMEDDataAccess();
			IDbConnection dbConn = null;
			IDbCommand dbCommand = null;

			try
			{
				// open ImedDataAccess class
				IMEDDataAccess.ConnectionType imedConnType;
				// calculate connection type
				imedConnType=IMEDDataAccess.CalculateConnectionType( connectionString );
				// get connection & open
				dbConn = imedData.GetConnection( imedConnType, connectionString );
				dbConn.Open();

				// get command
				dbCommand = imedData.GetCommand( dbConn );
				// create command text
				dbCommand.CommandText = sql;
				dbCommand.ExecuteNonQuery();
			}
			finally
			{
				//clean up objects
				dbCommand.Dispose();
			}
		}

		public static string Sproc_SP_MACRO_CODING_UPDATE_DIR( IMEDDataAccess.ConnectionType cType, int clinicalTrialId, 
			string trialSite, int personId, int responseTaskId, int repeat, string dictionaryName, string dictionaryVersion, 
			CodedTerm.eCodingStatus codingStatus, string codingDetails )
		{
			string sql = "";

			switch( cType )
			{
				case IMEDDataAccess.ConnectionType.SQLServer:
					sql = "EXEC SP_MACRO_CODING_UPDATE_DIR "
						+ "@clinicaltrialid = " + clinicalTrialId + ", "
						+ "@trialsite = '" + trialSite + "', "
						+ "@personid = " + personId + ", "
						+ "@responsetaskid = " + responseTaskId + ", "
						+ "@repeatnumber = " + repeat + ", " 
						+ "@dictionaryname = " + ( ( dictionaryName == "" ) ? "null" : "'" + dictionaryName + "'" ) + ", "
						+ "@dictionaryversion = " + ( ( dictionaryVersion == "" ) ? "null" : "'" + dictionaryVersion + "'" ) + ", "
						+ "@codingstatus = " + System.Convert.ToInt32( codingStatus ) + ", "
						+ "@codingdetails = " + ( ( codingDetails == "" ) ? "null" : "'" + SQLTerm( codingDetails ) + "'" );
					break;
				case IMEDDataAccess.ConnectionType.Oracle:
					sql = "CALL SP_MACRO_CODING_UPDATE_DIR("
						+ clinicalTrialId + ", "
						+ "'" + trialSite + "', "
						+ personId + ", "
						+ responseTaskId + ", "
						+ repeat + ", " 
						+ ( ( dictionaryName == "" ) ? "null" : "'" + dictionaryName + "'" ) + ", "
						+ ( ( dictionaryVersion == "" ) ? "null" : "'" + dictionaryVersion + "'" ) + ", "
						+ System.Convert.ToInt32( codingStatus ) + ", "
						+ ( ( codingDetails == "" ) ? "null" : "'" + SQLTerm( codingDetails ) + "')" );
					break;
			}

			return( sql );
		}

		/// <summary>
		/// Run a clinical coding stored procedure
		/// </summary>
		/// <param name="cType"></param>
		/// <param name="clinicalTrialId"></param>
		/// <param name="trialSite"></param>
		/// <param name="personId"></param>
		/// <param name="visitId"></param>
		/// <param name="visitCycle"></param>
		/// <param name="crfPageId"></param>
		/// <param name="crfPageCycle"></param>
		/// <param name="responseTaskId"></param>
		/// <param name="repeat"></param>
		/// <param name="dictionaryName"></param>
		/// <param name="dictionaryVersion"></param>
		/// <param name="codingStatus"></param>
		/// <param name="codingDetails"></param>
		/// <param name="codingTimeStamp"></param>
		/// <param name="codingTimeStamp_TZ"></param>
		/// <param name="responseValue"></param>
		/// <param name="responseTimeStamp"></param>
		/// <param name="responseTimeStamp_TZ"></param>
		/// <param name="userName"></param>
		/// <param name="userNameFull"></param>
		/// <param name="reasonForChange"></param>
		/// <returns></returns>
		public static string Sproc_SP_MACRO_CODING_UPDATE( IMEDDataAccess.ConnectionType cType, int clinicalTrialId, string trialSite, int personId, 
			int visitId, int visitCycle, int crfPageId, int crfPageCycle, int responseTaskId, int repeat, string dictionaryName, 
			string dictionaryVersion, CodedTerm.eCodingStatus codingStatus, 
			string codingDetails, double codingTimeStamp, short codingTimeStamp_TZ, string responseValue, double responseTimeStamp, 
			short responseTimeStamp_TZ, string userName, string userNameFull, string reasonForChange )
		{
			string sql = "";

			//MLM 15/11/07: Added LocalNumToStandards around concatenated timestamps below..
			switch( cType )
			{
				case IMEDDataAccess.ConnectionType.SQLServer:
					sql = "EXEC SP_MACRO_CODING_UPDATE "
						+ "@clinicaltrialid = " + clinicalTrialId + ", "
						+ "@trialsite = '" + trialSite + "', "
						+ "@personid = " + personId + ", "
						+ "@responsetaskid = " + responseTaskId + ", "
						+ "@repeatnumber = " + repeat + ", " 
						+ "@responsetimestamp = " + IMEDFunctions20.LocalNumToStandard(responseTimeStamp.ToString(), false) + ", "
						+ "@dictionaryname = " + ( ( dictionaryName == "" ) ? "null" : "'" + dictionaryName + "'" ) + ", "
						+ "@dictionaryversion = " + ( ( dictionaryVersion == "" ) ? "null" : "'" + dictionaryVersion + "'" ) + ", "
						+ "@codingstatus = " + System.Convert.ToInt32( codingStatus ) + ", "
						+ "@codingdetails = " + ( ( codingDetails == "" ) ? "null" : "'" + SQLTerm( codingDetails ) + "'" ) + ", " 
						+ "@visitid = " + visitId + ", "
						+ "@visitcyclenumber = " + visitCycle + ", "
						+ "@crfpageid = " + crfPageId + ", "
						+ "@crfpagecyclenumber = " + crfPageCycle + ", "
						+ "@username = '" + userName + "', "
						+ "@usernamefull = '" + userNameFull + "', " 
						+ "@responsetimestamp_tz = " + responseTimeStamp_TZ + ", "
						+ "@codingtimestamp = " + IMEDFunctions20.LocalNumToStandard(codingTimeStamp.ToString(), false) + ", "
						+ "@codingtimestamp_tz = " + codingTimeStamp_TZ + ", "
						+ "@reasonforchange = " + ( ( reasonForChange == "" ) ? "null" : "'" + SQLTerm( reasonForChange ) + "'" ) + ", "
						+ "@responsevalue = " + ( ( responseValue == "" ) ? "null" : "'" + SQLTerm( responseValue ) + "'" );
					break;
				case IMEDDataAccess.ConnectionType.Oracle:
					sql = "CALL SP_MACRO_CODING_UPDATE("
						+ clinicalTrialId + ", "
						+ "'" + trialSite + "', "
						+ personId + ", "
						+ responseTaskId + ", "
						+ repeat + ", " 
						+ IMEDFunctions20.LocalNumToStandard(responseTimeStamp.ToString(), false) + ", "
						+ ( ( dictionaryName == "" ) ? "null" : "'" + dictionaryName + "'" ) + ", "
						+ ( ( dictionaryVersion == "" ) ? "null" : "'" + dictionaryVersion + "'" ) + ", "
						+ System.Convert.ToInt32( codingStatus ) + ", "
						+ ( ( codingDetails == "" ) ? "null" : "'" + SQLTerm( codingDetails ) + "'" ) + ", " 
						+ visitId + ", "
						+ visitCycle + ", "
						+ crfPageId + ", "
						+ crfPageCycle + ", "
						+ "'" + userName + "', "
						+ "'" + userNameFull + "', " 
						+ responseTimeStamp_TZ + ", "
						+ IMEDFunctions20.LocalNumToStandard(codingTimeStamp.ToString(), false) + ", "
						+ codingTimeStamp_TZ + ", "
						+ ( ( reasonForChange == "" ) ? "null" : "'" + SQLTerm( reasonForChange ) + "'" ) + ", "
						+ ( ( responseValue == "" ) ? "null" : "'" + SQLTerm( responseValue ) + "'" )
						+ ")";
					break;
			}
			
			log.Info(sql);

			return( sql );
		}
	}
}
