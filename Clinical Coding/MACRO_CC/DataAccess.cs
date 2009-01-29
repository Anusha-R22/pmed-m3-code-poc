using System;
using System.Data;
using InferMed.Components;

namespace InferMed.MACRO.ClinicalCoding.MACRO_CC
{
	/// <summary>
	/// Summary description for DataAccess.
	/// </summary>
	public class DataAccess
	{
		public DataAccess()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		/// <summary>
		/// Check for a permission belonging to a role
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="roleCode"></param>
		/// <param name="functionCode"></param>
		/// <returns></returns>
		public static bool CheckPermission(IDbConnection dbConn, string roleCode, string functionCode)
		{
			DataSet ds = new DataSet();

			try
			{
				string sql = "SELECT COUNT(*) AS MATCHES "
					+ "FROM ROLEFUNCTION "
					+ "WHERE ROLECODE = '" + roleCode + "' "
					+ "AND FUNCTIONCODE = '" + functionCode + "' ";

				//execute the query
				ds = GetDataSet( dbConn, sql );
				//return whether the dataitemcode was found or not 
				return( ( System.Convert.ToInt32( ( ds.Tables[0].Rows[0]["MATCHES"].ToString() ) ) > 0 ) );
			}
			finally
			{
				ds.Dispose();
			}
		}

		public static DataSet GetDatabaseList(string connectionString)
		{
			string sql = "SELECT DATABASECODE, DATABASETYPE, DATABASEUSER, DATABASEPASSWORD, SERVERNAME, "
				+ "NAMEOFDATABASE FROM DATABASES";
			return( GetDataSet( connectionString, sql ) );
		}

		public static DataSet GetStudyList(string connectionString)
		{
			string sql = "SELECT CLINICALTRIALNAME, CLINICALTRIALID FROM CLINICALTRIAL WHERE CLINICALTRIALID > 0 "
				+ "ORDER BY CLINICALTRIALNAME";
			return( GetDataSet(connectionString, sql) );
		}

		public static DataSet GetSiteList(string connectionString, string siteCode)
		{
			string sql = "SELECT TRIALSITE "
				+ "FROM CLINICALTRIAL, TRIALSITE "
				+ "WHERE CLINICALTRIAL.CLINICALTRIALID = TRIALSITE.CLINICALTRIALID "
				+ "AND CLINICALTRIALNAME = '" + siteCode + "' "
				+ "ORDER BY TRIALSITE";
			return( GetDataSet( connectionString, sql) );
		}

		/// <summary>
		/// Prepare a string for adding to a SQL string
		/// </summary>
		/// <param name="term"></param>
		/// <returns></returns>
		public static string SQLTerm( string term )
		{
			return( term.Replace( "'", "''" ) );
		}

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

		public static void RunSQL(string connectionString, string sql)
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
				
				RunSQL( dbConn, sql );
			}
			finally
			{
				//clean up objects
				dbConn.Close();
				dbConn.Dispose();
			}
		}

		public static void RunSQL( IDbConnection dbConn, string sql )
		{
			IMEDDataAccess imedData = new IMEDDataAccess();
			IDbCommand dbCommand = null;

			try
			{
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
	}
}
