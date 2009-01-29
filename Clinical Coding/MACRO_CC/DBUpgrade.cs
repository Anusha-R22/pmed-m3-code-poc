using System;
using System.Data;
using System.Text;
using System.Windows.Forms;
using InferMed.Components;

namespace InferMed.MACRO.ClinicalCoding.MACRO_CC
{
	/// <summary>
	/// Summary description for DBUpgrade.
	/// </summary>
	public class DBUpgrade
	{
		//tables
		private const string _DICTIONARIES_TABLE = "DICTIONARIES";
		private const string _CODINGHISTORY_TABLE = "CODINGHISTORY";
		private const string _DICTIONARYINFO_TABLE = "DICTIONARY_INFO";
		//upgraded flags
		private bool _dbUpgraded = false;
		private bool _secDBUpgraded = false;
		private bool _medicoderDBUpgraded = false;
		//connection strings
		private string _secCon = "";
		private string _dbCon = "";
		private string _medicoderCon = "";

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="secCon"></param>
		/// <param name="dbCon"></param>
		public DBUpgrade( string secCon, string dbCon )
		{
			_secCon = secCon;
			_dbCon = dbCon;
			_secDBUpgraded = DBUpgrade.CCTableExists( _secCon, _DICTIONARIES_TABLE );
			_dbUpgraded = DBUpgrade.CCTableExists( _dbCon, _CODINGHISTORY_TABLE );
		}

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="medicoderCon"></param>
		public DBUpgrade( string medicoderCon )
		{
			_medicoderCon = medicoderCon;
			_medicoderDBUpgraded = DBUpgrade.CCTableExists( _medicoderCon, _DICTIONARYINFO_TABLE );
		}

		/// <summary>
		/// Does the database exist
		/// </summary>
		public static bool DatabaseExists( string con )
		{
			return( true );
		}

		/// <summary>
		/// Is a database upgrade required
		/// </summary>
		public bool UpgradeRequired
		{
			get
			{ 
				if( _medicoderCon != "" )
				{
					return( !_medicoderDBUpgraded );
				}
				else
				{
					return( !( _dbUpgraded && _secDBUpgraded ) ); }
			}
		}

		/// <summary>
		/// Check if CC tables exist in given db connection string
		/// </summary>
		/// <param name="sDbConn"></param>
		/// <returns></returns>
		private static bool CCTableExists(string sDbConn, string ccTableName)
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
				imedConnType=IMEDDataAccess.CalculateConnectionType(sDbConn);

				// get connection
				dbConn = imedData.GetConnection(imedConnType,sDbConn);

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
				sbSQL.Append("WHERE TABLE_NAME = '" + ccTableName + "'");
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
		/// Upgrade database to clinical coding version
		/// </summary>
		/// <param name="sDbConn"></param>
		public void Upgrade()
		{
			// declare connection interface
			IDbConnection dbCon = null;
			IMEDDataAccess imedData = new IMEDDataAccess();
			

			try
			{
				if( _medicoderCon != "" )
				{
					if( !_medicoderDBUpgraded )
					{
						string sqlStringsDB = "";

						if( IMEDDataAccess.CalculateConnectionType( _medicoderCon ) == IMEDDataAccess.ConnectionType.SQLServer )
						{
							// SQL Server
							sqlStringsDB=IMEDFunctions20.ExtractAllTextFromFile(IMEDFunctions20.GetAppPath() + "SQL_CC_MEDICODER.sql");
						}
						else
						{
							// Oracle
							sqlStringsDB=IMEDFunctions20.ExtractAllTextFromFile(IMEDFunctions20.GetAppPath() + "ORA_CC_MEDICODER.sql");
						}
						dbCon = GetDbConnection( _medicoderCon );
						RunSQLStrings( ( imedData.GetCommand( dbCon ) ), sqlStringsDB );
						dbCon.Close();

					}
				}
				else
				{
					if( !_secDBUpgraded )
					{
						string sqlStringsSec = "";

						if( IMEDDataAccess.CalculateConnectionType( _secCon ) == IMEDDataAccess.ConnectionType.SQLServer )
						{
							// SQL Server
							sqlStringsSec=IMEDFunctions20.ExtractAllTextFromFile(IMEDFunctions20.GetAppPath() + "SQL_CC_MACROSEC.sql");
						}
						else
						{
							// Oracle
							sqlStringsSec=IMEDFunctions20.ExtractAllTextFromFile(IMEDFunctions20.GetAppPath() + "ORA_CC_MACROSEC.sql");
						}
						dbCon = GetDbConnection( _secCon );
						RunSQLStrings( ( imedData.GetCommand( dbCon ) ), sqlStringsSec );
						dbCon.Close();
					}

					if( !_dbUpgraded )
					{
						string sqlStringsDB = "";

						if( IMEDDataAccess.CalculateConnectionType( _dbCon ) == IMEDDataAccess.ConnectionType.SQLServer )
						{
							// SQL Server
							sqlStringsDB=IMEDFunctions20.ExtractAllTextFromFile(IMEDFunctions20.GetAppPath() + "SQL_CC_MACRO.sql");
						}
						else
						{
							// Oracle
							sqlStringsDB=IMEDFunctions20.ExtractAllTextFromFile(IMEDFunctions20.GetAppPath() + "ORA_CC_MACRO.sql");
						}
						dbCon = GetDbConnection( _dbCon );
						RunSQLStrings( ( imedData.GetCommand( dbCon ) ), sqlStringsDB );
						dbCon.Close();

					}
				}
			}
			finally
			{
				if( dbCon.State == ConnectionState.Open )
				{
					dbCon.Close();
				}
			}
		}

		/// <summary>
		/// Return a database connection
		/// </summary>
		/// <param name="con"></param>
		/// <returns></returns>
		private IDbConnection GetDbConnection( string con )
		{
			IDbConnection dbCon = null;
			IMEDDataAccess imedData = new IMEDDataAccess();

			// open ImedDataAccess class
			IMEDDataAccess.ConnectionType imedConnType;

			// calculate connection type
			imedConnType=IMEDDataAccess.CalculateConnectionType( con );

			// get connection
			dbCon = imedData.GetConnection( imedConnType, con );

			// open connection
			dbCon.Open();

			// return connection
			return( dbCon );
		}

		/// <summary>
		/// Run crlf delimited sql strings
		/// </summary>
		/// <param name="dbCommand"></param>
		/// <param name="sqlStrings"></param>
		private void RunSQLStrings( IDbCommand dbCommand, string sqlStrings )
		{
			// if no sql returned then failed
			if( sqlStrings == "" )
			{
				Exception ex = new Exception( "No sql found in file" );
				throw ex;
			}

			// split file and execute a row at a time
			char[] chCrLf = {System.Convert.ToChar("\r"), System.Convert.ToChar("\n")};
			string[] aSql = sqlStrings.Split(chCrLf);

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
	}
}
