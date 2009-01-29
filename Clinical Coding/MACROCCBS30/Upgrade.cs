using System;
using System.Data;
using InferMed.Components;

namespace InferMed.MACRO.ClinicalCoding.MACROCCBS30
{
	/// <summary>
	/// Summary description for Upgrade.
	/// </summary>
	public class Upgrade
	{
		public Upgrade()
		{
			//
			// TODO: Add constructor logic here
			//
		}

//		/// <summary>
//		/// check if tables exist
//		/// </summary>
//		/// <returns></returns>
//		public bool DoCCTablesExist( string con )
//		{
//			bool bExist=true;
//			StringBuilder sbSQL = new StringBuilder();
//
//			// declare connection interface
//			IDbConnection dbConn = null;
//			IDbCommand dbCommand = null;
//			IDbDataAdapter dbDataAdapter = null;
//			DataSet dataDB = new DataSet();
//
//			try
//			{
//				// open ImedDataAccess class
//				IMEDDataAccess imedData = new IMEDDataAccess();
//				IMEDDataAccess.ConnectionType imedConnType;
//
//				// calculate connection type
//				imedConnType=IMEDDataAccess.CalculateConnectionType(_sMACROdbConn);
//
//				// get connection
//				dbConn = imedData.GetConnection(imedConnType,_sMACROdbConn);
//
//				// open connection
//				dbConn.Open();
//
//				// get command object
//				dbCommand = imedData.GetCommand(dbConn);
//				// set command text
//				sbSQL.Append("SELECT TABLE_NAME FROM ");
//				if(imedConnType==IMEDDataAccess.ConnectionType.SQLServer)
//				{
//					// sql server
//					sbSQL.Append("INFORMATION_SCHEMA.Tables ");
//				}
//				else
//				{
//					// oracle
//					sbSQL.Append("USER_TABLES ");
//				}
//				sbSQL.Append("WHERE TABLE_NAME = 'CODINGHISTORY'");
//				dbCommand.CommandText=sbSQL.ToString();
//
//				// get data adaptor
//				dbDataAdapter = imedData.GetDataAdapter(dbCommand);
//
//				// fill dataset
//				dbDataAdapter.Fill(dataDB);
//
//				// get datatable
//				DataTable dataTable = dataDB.Tables[0];
//
//				// check if tables exist
//				if(dataTable.Rows.Count == 0)
//				{
//					bExist=false;
//				}
//			}
//			catch
//			{
//				return false;
//			}
//			finally
//			{
//				// in case of db problem
//				// close db
//				dbConn.Close();
//			}
//
//			// return
//			return bExist;
//		}
//
//		/// <summary>
//		/// create tables in connected to database
//		/// </summary>
//		public bool RunSQLFile( string fileName )
//		{
//			// created tables successfully
//			bool bOk = true;
//
//			// declare connection interface
//			IDbConnection dbConn = null;
//			IDbCommand dbCommand = null;
//
//			try
//			{
//				// open ImedDataAccess class
//				IMEDDataAccess imedData = new IMEDDataAccess();
//				IMEDDataAccess.ConnectionType imedConnType;
//
//				// calculate connection type
//				imedConnType=IMEDDataAccess.CalculateConnectionType(_sMACROdbConn);
//
//				// get connection
//				dbConn = imedData.GetConnection(imedConnType,_sMACROdbConn);
//
//				// open connection
//				dbConn.Open();
//
//				// get command object
//				dbCommand = imedData.GetCommand(dbConn);
//
//				// retrieve SQL table script
//				string sSQLFile="";
//
//				//get filename
//				sSQLFile=IMEDFunctions20.ExtractAllTextFromFile(IMEDFunctions20.GetAppPath() + fileName );
//				
//
//				// if no sql returned then failed
//				if(sSQLFile=="")
//				{
//					bOk = false;
//				}
//
//				// split file and execute a row at a time
//				char[] chCrLf = {System.Convert.ToChar("\r"), System.Convert.ToChar("\n")};
//				string[] aSql = sSQLFile.Split(chCrLf);
//
//				// loop through and execute a line at a time
//				foreach(string sSql in aSql)
//				{
//					if(sSql!="")
//					{
//						// set text
//						dbCommand.CommandText=sSql;
//						// execute
//						dbCommand.ExecuteNonQuery();
//					}
//				}
//			}
//			catch
//			{
//				bOk = false;
//				throw;
//			}
//			finally
//			{
//				dbConn.Close();
//			}
//
//			return bOk;
//		}
	}
}
