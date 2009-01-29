using System;
using System.Data;
using InferMed.Components;

namespace InferMed.MACRO.ClinicalCoding.MedDRAPreloader
{
	/// <summary>
	/// Summary description for DataAccess.
	/// </summary>
	public class DataAccess
	{
		private DataAccess()
		{
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
	}
}
