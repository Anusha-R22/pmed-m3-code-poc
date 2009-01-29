using System;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.OracleClient;
using System.Collections.Generic;
using System.Text;

namespace MACROAPI_dot_net_test_harness
{
	public class DataAccess //: IDisposable 
	{
		/// <summary>
        /// DataAccess class 
		/// </summary>
        //private IDbConnection idbConn = null;
        //private IDbDataAdapter idbAdapter = null;
        //private IDbCommand idbCommand = null;
        //private IDataReader idbReader = null;
        //private IDbDataParameter idbParameter = null;

        // microsoft OLEDB provider
		private const string _CONNECTION_MSDAORA = "MSDAORA";
        // oracle OLEDB provider
        private const string _CONNECTION_ORAOLEDB = "ORAOLEDB.ORACLE";
		//sql server provider keyword
		private const string _CONNECTION_SQLOLEDB = "SQLOLEDB";
		//oledb provider
		private const string _CONNECTION_PROVIDER = "PROVIDER";
		//tnsname for oracle, server for SQL Server
		private const string _CONNECTION_DATASOURCE = "DATA SOURCE";
		//database for SQL server only
		private const string _CONNECTION_DATABASE = "DATABASE";
		//user id for oracle and SQL server
		private const string _CONNECTION_USERID = "USER ID";
		//password for SQL server and Oracle
		private const string _CONNECTION_PASSWORD = "PASSWORD";

        // parameter name used in out cursors in oracle stored procedures
        private const string _ORACLE_OUTCURSOR_PARAM_NAME = "OUTCURSOR";

		public enum ConnectionType
		{
			Unknown = -1,
			Oracle = 1,
			SQLServer = 2,
			OleDb = 3
		}

        public enum MACROConnectionType
        {
            SQLServer = 1,
            Oracle = 3
        }

        /// <summary>
        /// constructor
        /// </summary>
		private DataAccess()
		{
		}

		/// <summary>
		/// attempt to calculate connection type
		/// </summary>
		/// <param name="sConString"></param>
		/// <returns>ConnectionType</returns>
		public static ConnectionType CalculateConnectionType(string sConString)
		{
			// default connection type to OleDb
			ConnectionType ConnType=ConnectionType.OleDb;
			// check for specific provider types
			// oracle
            if ((sConString.IndexOf(_CONNECTION_MSDAORA) > 0) || (sConString.IndexOf(_CONNECTION_ORAOLEDB) > 0) || (sConString.IndexOf("ORAOLEDB.ORACLE.1") > 0))
			{
				ConnType = ConnectionType.Oracle;
			}
			// sql server
            if (sConString.IndexOf(_CONNECTION_SQLOLEDB) > 0)
			{
				ConnType = ConnectionType.SQLServer;
			}
			return ConnType;
		}

        /// <summary>
        /// attempt to calculate connection type from Connection object
        /// </summary>
        /// <param name="DbConnection"></param>
        /// <returns></returns>
		public static ConnectionType CalculateConnectionType(IDbConnection DbConnection)
		{
			ConnectionType connType;
			switch(DbConnection.GetType().FullName)
			{
				case "System.Data.SqlClient.SqlConnection":
				{
					connType = ConnectionType.SQLServer;
					break;
				}
				case "System.Data.OracleClient.OracleConnection":
				{
					connType = ConnectionType.Oracle;
					break;
				}
				default:
				{
					connType = ConnectionType.OleDb;
					break;
				}
			}
			return connType;
		}

        /// <summary>
        /// Convert a MACRO connection type to standard connection type
        /// </summary>
        /// <param name="macroConnectionType"></param>
        /// <returns></returns>
        public static ConnectionType ConvertFromMACROConnectionType(string macroConnectionType)
        {
            switch ((MACROConnectionType)Convert.ToInt16(macroConnectionType))
            {
                case MACROConnectionType.SQLServer:
                    return ConnectionType.SQLServer;
                case MACROConnectionType.Oracle:
                    return ConnectionType.Oracle;
                default:
                    return ConnectionType.Unknown;
            }
        }

        /// <summary>
        /// create connection string from params
        /// </summary>
        /// <param name="conType"></param>
        /// <param name="sServerName"></param>
        /// <param name="sDbName"></param>
        /// <param name="sDbUser"></param>
        /// <param name="sDbPassword"></param>
        /// <returns></returns>
		public static string CreateDBConnectionString(ConnectionType conType, string sServerName, 
			string sDbName, string sDbUser, 
			string sDbPassword)
		{
			StringBuilder sbConn = new System.Text.StringBuilder();
			// 'oledb provider - all
			const string sCONNECTION_PROVIDER = "PROVIDER";
			//tnsname for oracle, server for SQL Server
			const string sCONNECTION_DATASOURCE = "DATA SOURCE";
			//database for SQL server only
			const string sCONNECTION_DATABASE = "DATABASE";
			//user id for oracle and SQL server
			const string sCONNECTION_USERID = "USER ID";
			//password for SQL server and Oracle
			const string sCONNECTION_PASSWORD = "PASSWORD";
			// connection strings
			const string sCONNECTION_SQLOLEDB = "SQLOLEDB";
			const string sCONNECTION_MSDAORA = "MSDAORA";

			switch(conType)
			{
				case ConnectionType.SQLServer:
				{
					//sql server
					sbConn.Append(sCONNECTION_PROVIDER + "=" + sCONNECTION_SQLOLEDB + ";");
					sbConn.Append(sCONNECTION_DATASOURCE + "=" + sServerName + ";");
					sbConn.Append(sCONNECTION_DATABASE + "=" + sDbName + ";");
					sbConn.Append(sCONNECTION_USERID + "=" + sDbUser + ";");
					sbConn.Append(sCONNECTION_PASSWORD + "=" + sDbPassword + ";");
					break;					
				}
				case ConnectionType.Oracle:
				{
					//oracle oledb native provider
					sbConn.Append(sCONNECTION_PROVIDER + "=" + sCONNECTION_MSDAORA + ";");
					sbConn.Append(sCONNECTION_DATASOURCE + "=" + sServerName + ";");
					sbConn.Append(sCONNECTION_USERID + "=" + sDbUser + ";");
					sbConn.Append(sCONNECTION_PASSWORD + "=" + sDbPassword + ";");
					break;
				}
				default:
				{
					break;
				}
			}

			return sbConn.ToString();
		}

        /// <summary>
        /// check connection string and make sure in right format for connection type
        /// </summary>
        /// <param name="conType"></param>
        /// <param name="sConString"></param>
        /// <returns></returns>
		public static string FormatConnectionString(ConnectionType conType, string sConString)
		{
			string sFormatted;
			// remove unwanted info
			char[] chDelim = { System.Convert.ToChar(";") };
			string[] aString = sConString.Split(chDelim);
			StringBuilder sbConn = new System.Text.StringBuilder();
			foreach(string sStringPart in aString)
			{
				switch(conType)
				{
					case ConnectionType.Oracle:
					{
						// remove provider=x; from string
						// remove database=x; for oracle
						if(((sStringPart.Replace(" ", "").ToLower()).IndexOf("provider=")==-1)&&
							((sStringPart.Replace(" ", "").ToLower()).IndexOf("database=")==-1))
						{
							sbConn.Append(sStringPart + ";");
						}
						break;
					}
					case ConnectionType.SQLServer:
					{
						// remove provider=x; from string
                        if ((sStringPart.Replace(" ", "").ToLower()).IndexOf("provider=") == -1)
						{
							sbConn.Append(sStringPart + ";");
						}
						break;
					}
					default:
					{
						// just append
						sbConn.Append(sStringPart + ";");
						break;
					}
				}
			}
			sFormatted = sbConn.ToString();
			return sFormatted;
		}

        /// <summary>
        /// GetConnection returns IDbConnection
        /// </summary>
        /// <param name="conType"></param>
        /// <param name="sConString"></param>
        /// <returns></returns>
		public static IDbConnection GetConnection(ConnectionType conType, 
			string sConString)
		{
            IDbConnection idbConn = null;

			// Format connection string for connection type
			sConString = FormatConnectionString(conType, sConString);
			switch(conType) 
			{
				case ConnectionType.Oracle: // Oracle Data Provider
					idbConn = new OracleConnection(sConString);
					break;
				case ConnectionType.SQLServer: // Sql Data Provider
					idbConn = new SqlConnection(sConString);
					break;
				case ConnectionType.OleDb: // OleDb Data Provider
					idbConn = new OleDbConnection(sConString);
					break;
				default:
					break;
			}
			return idbConn;
		}

        /// <summary>
        /// GetConnection returns IDbConnection
        /// </summary>
        /// <param name="sConString"></param>
        /// <returns></returns>
        public static IDbConnection GetConnection(string sConString)
        {
            IDbConnection idbConn = null;

            DataAccess.ConnectionType ct = DataAccess.CalculateConnectionType(sConString);

            // Format connection string for connection type
            sConString = FormatConnectionString(ct, sConString);
            switch (ct)
            {
                case ConnectionType.Oracle: // Oracle Data Provider
                    idbConn = new OracleConnection(sConString);
                    break;
                case ConnectionType.SQLServer: // Sql Data Provider
                    idbConn = new SqlConnection(sConString);
                    break;
                case ConnectionType.OleDb: // OleDb Data Provider
                    idbConn = new OleDbConnection(sConString);
                    break;
                default:
                    break;
            }
            return idbConn;
        }

        /// <summary>
        /// GetCommand returns IDbCommand
		/// 2 overloaded versions - one with sql, one without
        /// </summary>
        /// <param name="conType"></param>
        /// <param name="DbConnection"></param>
        /// <param name="sSql"></param>
        /// <returns></returns>
		public static IDbCommand GetCommand(ConnectionType conType, 
			IDbConnection DbConnection, string sSql)
		{
            IDbCommand idbCommand = null;

			switch(conType) 
			{
				case ConnectionType.Oracle: // Oracle Data Provider
					idbCommand = new OracleCommand(sSql, (OracleConnection)DbConnection);
					break;
				case ConnectionType.SQLServer: // Sql Data Provider
					idbCommand = new SqlCommand(sSql,(SqlConnection)DbConnection);
					break;
				case ConnectionType.OleDb: // OleDb Data Provider
					idbCommand = new OleDbCommand(sSql, (OleDbConnection)DbConnection);
					break;
				default:
					break;
			}

			return idbCommand;
		}

        /// <summary>
        /// one with sql & calculate connection type
        /// </summary>
        /// <param name="DbConnection"></param>
        /// <param name="sSql"></param>
        /// <returns></returns>
		public static IDbCommand GetCommand(IDbConnection DbConnection,
			string sSql)
		{
            IDbCommand idbCommand = null;

			// calculate connection type
			ConnectionType conType = CalculateConnectionType(DbConnection);
			switch(conType) 
			{
				case ConnectionType.Oracle: // Oracle Data Provider
					idbCommand = new OracleCommand(sSql, (OracleConnection)DbConnection);
					break;
				case ConnectionType.SQLServer: // Sql Data Provider
					idbCommand = new SqlCommand(sSql,(SqlConnection)DbConnection);
					break;
				case ConnectionType.OleDb: // OleDb Data Provider
					idbCommand = new OleDbCommand(sSql, (OleDbConnection)DbConnection);
					break;
				default:
					break;
			}

			return idbCommand;
		}

        /// <summary>
        /// GetCommand returns IDbCommand
        /// no SQL
        /// </summary>
        /// <param name="conType"></param>
        /// <param name="DbConnection"></param>
        /// <returns></returns>
		public static IDbCommand GetCommand(ConnectionType conType, 
			IDbConnection DbConnection)
		{
            IDbCommand idbCommand = null;

			switch(conType) 
			{
				case ConnectionType.Oracle: // Oracle Data Provider
					idbCommand = ((OracleConnection)DbConnection).CreateCommand();
					break;
				case ConnectionType.SQLServer: // Sql Data Provider
					idbCommand = ((SqlConnection)DbConnection).CreateCommand();
					break;
				case ConnectionType.OleDb: // OleDb Data Provider
					idbCommand = ((OleDbConnection)DbConnection).CreateCommand();
					break;
				default:
					break;
			}

			return idbCommand;
		}

        /// <summary>
        /// no SQL & calculate connection type
        /// </summary>
        /// <param name="DbConnection"></param>
        /// <returns></returns>
		public static IDbCommand GetCommand(IDbConnection DbConnection)
		{
            IDbCommand idbCommand = null;

			// calculate connection type
			ConnectionType conType = CalculateConnectionType(DbConnection);
			switch(conType) 
			{
				case ConnectionType.Oracle: // Oracle Data Provider
					idbCommand = ((OracleConnection)DbConnection).CreateCommand();
					break;
				case ConnectionType.SQLServer: // Sql Data Provider
					idbCommand = ((SqlConnection)DbConnection).CreateCommand();
					break;
				case ConnectionType.OleDb: // OleDb Data Provider
					idbCommand = ((OleDbConnection)DbConnection).CreateCommand();
					break;
				default:
					break;
			}

			return idbCommand;
		}

        /// <summary>
        /// no SQL & calculate connection type
        /// </summary>
        /// <param name="DbConnection"></param>
        /// <returns></returns>
        public static IDbCommand GetCommand(string conString)
        {
            IDbConnection idbConnection = DataAccess.GetConnection(conString);
            IDbCommand idbCommand = null;

            // calculate connection type
            ConnectionType conType = CalculateConnectionType(idbConnection);
            switch (conType)
            {
                case ConnectionType.Oracle: // Oracle Data Provider
                    idbCommand = ((OracleConnection)idbConnection).CreateCommand();
                    break;
                case ConnectionType.SQLServer: // Sql Data Provider
                    idbCommand = ((SqlConnection)idbConnection).CreateCommand();
                    break;
                case ConnectionType.OleDb: // OleDb Data Provider
                    idbCommand = ((OleDbConnection)idbConnection).CreateCommand();
                    break;
                default:
                    break;
            }

            return idbCommand;
        }

        /// <summary>
        /// GetDataAdapter returns IDbDataAdapter
        /// overloaded 3 times
        /// Connection type, connection string, sql
        /// </summary>
        /// <param name="conType"></param>
        /// <param name="sConString"></param>
        /// <param name="sSql"></param>
        /// <returns></returns>
		public static IDbDataAdapter GetDataAdapter(ConnectionType conType, 
			string sConString, string sSql)
		{
            IDbDataAdapter idbAdapter = null;

			// Format connection string for connection type
			sConString = FormatConnectionString(conType, sConString);
			switch(conType) 
			{
				case ConnectionType.Oracle: // Oracle Data Provider
					idbAdapter = new OracleDataAdapter(sSql, sConString);
					break;
				case ConnectionType.OleDb: // OleDb Data Provider
					idbAdapter = new OleDbDataAdapter(sSql, sConString);
					break;
				case ConnectionType.SQLServer: // Sql Data Provider
					idbAdapter = new SqlDataAdapter(sSql, sConString); 
					break;
				default:
					break;
			}
			return idbAdapter;
		}

        /// <summary>
        /// Connection type, connection object, sql
        /// </summary>
        /// <param name="conType"></param>
        /// <param name="DbConnection"></param>
        /// <param name="sSql"></param>
        /// <returns></returns>
		public static IDbDataAdapter GetDataAdapter(ConnectionType conType, 
			IDbConnection DbConnection, string sSql)
		{
            IDbDataAdapter idbAdapter = null;

			switch(conType) 
			{
				case ConnectionType.Oracle: // Oracle Data Provider
					idbAdapter = new OracleDataAdapter(sSql, (OracleConnection)DbConnection);
					break;
				case ConnectionType.OleDb: // OleDb Data Provider
					idbAdapter = new OleDbDataAdapter(sSql, (OleDbConnection)DbConnection);

					break;
				case ConnectionType.SQLServer: // Sql Data Provider
					idbAdapter = new SqlDataAdapter(sSql, (SqlConnection)DbConnection);
					break;
				default:
					break;
			}
			return idbAdapter;
		}

        /// <summary>
        /// connection object, sql
        /// </summary>
        /// <param name="DbConnection"></param>
        /// <param name="sSql"></param>
        /// <returns></returns>
		public static IDbDataAdapter GetDataAdapter(IDbConnection DbConnection, string sSql)
		{
            IDbDataAdapter idbAdapter = null;

			// calculate connection type
			ConnectionType conType = CalculateConnectionType(DbConnection);
			switch(conType) 
			{
				case ConnectionType.Oracle: // Oracle Data Provider
					idbAdapter = new OracleDataAdapter(sSql, (OracleConnection)DbConnection);
					break;
				case ConnectionType.OleDb: // OleDb Data Provider
					idbAdapter = new OleDbDataAdapter(sSql, (OleDbConnection)DbConnection);

					break;
				case ConnectionType.SQLServer: // Sql Data Provider
					idbAdapter = new SqlDataAdapter(sSql, (SqlConnection)DbConnection);
					break;
				default:
					break;
			}
			return idbAdapter;
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="conType"></param>
        /// <param name="DbCommand"></param>
        /// <returns></returns>
		public static IDbDataAdapter GetDataAdapter(ConnectionType conType, 
			IDbCommand DbCommand)
		{
            IDbDataAdapter idbAdapter = null;

			switch(conType) 
			{
				case ConnectionType.Oracle: // Oracle Data Provider
					idbAdapter = new OracleDataAdapter((OracleCommand)DbCommand);
					break;
				case ConnectionType.OleDb: // OleDb Data Provider
					idbAdapter = new OleDbDataAdapter((OleDbCommand)DbCommand);
					break;
				case ConnectionType.SQLServer: // Sql Data Provider
					idbAdapter = new SqlDataAdapter((SqlCommand)DbCommand); 
					break;
				default:
					break;
			}
			return idbAdapter;
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="DbCommand"></param>
        /// <returns></returns>
		public static IDbDataAdapter GetDataAdapter(IDbCommand DbCommand)
		{
            IDbDataAdapter idbAdapter = null;

			// calculate connection type
			ConnectionType conType = CalculateConnectionType(DbCommand.Connection);
			switch(conType) 
			{
				case ConnectionType.Oracle: // Oracle Data Provider
					idbAdapter = new OracleDataAdapter((OracleCommand)DbCommand);
					break;
				case ConnectionType.OleDb: // OleDb Data Provider
					idbAdapter = new OleDbDataAdapter((OleDbCommand)DbCommand);
					break;
				case ConnectionType.SQLServer: // Sql Data Provider
					idbAdapter = new SqlDataAdapter((SqlCommand)DbCommand); 
					break;
				default:
					break;
			}
			return idbAdapter;
		}

        /// <summary>
        /// GetDataReader returns IDataReader
        /// </summary>
        /// <param name="conType"></param>
        /// <param name="dbCommand"></param>
        /// <returns></returns>
		public static IDataReader GetDataReader(ConnectionType conType, 
			IDbCommand dbCommand)
		{
            IDataReader idbReader = null;

			try
			{
				switch(conType) 
				{
					case ConnectionType.Oracle: // Oracle Data Provider
						idbReader = ((OracleCommand)dbCommand).ExecuteReader();
						break;
					case ConnectionType.OleDb: // OleDb Data Provider
						idbReader = ((OleDbCommand)dbCommand).ExecuteReader();
						break;
					case ConnectionType.SQLServer: // Sql Data Provider
						idbReader = ((SqlCommand)dbCommand).ExecuteReader();
						break;
					default:
						break;
				}

				return idbReader;
			}
			finally
			{
			}
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dbCommand"></param>
        /// <returns></returns>
		public static IDataReader GetDataReader(IDbCommand dbCommand)
		{
            IDataReader idbReader = null;

			try
			{
				// calculate connection type
				ConnectionType conType = CalculateConnectionType(dbCommand.Connection);

				switch(conType) 
				{
					case ConnectionType.Oracle: // Oracle Data Provider
						idbReader = ((OracleCommand)dbCommand).ExecuteReader();
						break;
					case ConnectionType.OleDb: // OleDb Data Provider
						idbReader = ((OleDbCommand)dbCommand).ExecuteReader();
						break;
					case ConnectionType.SQLServer: // Sql Data Provider
						idbReader = ((SqlCommand)dbCommand).ExecuteReader();
						break;
					default:
						break;
				}

				return idbReader;
			}
			finally
			{
			}
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dbCommand"></param>
        /// <param name="cbToExecute"></param>
        /// <returns></returns>
		public static IDataReader GetDataReader(IDbCommand dbCommand, CommandBehavior cbToExecute)
		{
            IDataReader idbReader = null;

			try
			{
				// calculate connection type
				ConnectionType conType = CalculateConnectionType(dbCommand.Connection);

				switch(conType) 
				{
					case ConnectionType.Oracle: // Oracle Data Provider
						idbReader = ((OracleCommand)dbCommand).ExecuteReader(cbToExecute);
						break;
					case ConnectionType.OleDb: // OleDb Data Provider
						idbReader = ((OleDbCommand)dbCommand).ExecuteReader(cbToExecute);
						break;
					case ConnectionType.SQLServer: // Sql Data Provider
						idbReader = ((SqlCommand)dbCommand).ExecuteReader(cbToExecute);
						break;
					default:
						break;
				}

				return idbReader;
			}
			finally
			{
			}
		}


		public static IDbDataParameter AddDataParameter(ref IDbCommand dbCommand, string sParameterName, DbType eDbDataType, 
                                    ParameterDirection ParamDirection,  object ParamValue)
		{
            IDbDataParameter idbParameter = null;

			try
			{
				// calculate connection type
				ConnectionType conType = CalculateConnectionType(dbCommand.Connection);

				switch(conType) 
				{
                    case ConnectionType.Oracle: // Oracle Data Provider
                        {
                            OracleType oraType = DbTypeTranslationToOracle(eDbDataType);
                            idbParameter = ((OracleCommand)dbCommand).Parameters.Add(sParameterName, oraType);
                            break;
                        }
                    case ConnectionType.OleDb: // OleDb Data Provider
                        {
                            OleDbType oleDbType = DbTypeTranslationToOleDb(eDbDataType);
                            idbParameter = ((OleDbCommand)dbCommand).Parameters.Add(sParameterName, oleDbType);
                            break;
                        }
                    case ConnectionType.SQLServer: // Sql Data Provider
                        {
                            SqlDbType sqlDbType = DbTypeTranslationToSqlServer(eDbDataType);
                            idbParameter = ((SqlCommand)dbCommand).Parameters.Add(sParameterName, sqlDbType);
                            break;
                        }
                    default:
                        {
                            break;
                        }
				}
                // direction
                idbParameter.Direction = ParamDirection;
                // value
                idbParameter.Value = ParamValue;

				return idbParameter;
			}
			finally
			{
			}
		}

        /// <summary>
        /// Returns a SQL Server or ORACLE connection string
        /// </summary>
        /// <param name="cType"></param>
        /// <param name="sServer"></param>
        /// <param name="sNameOf"></param>
        /// <param name="sName"></param>
        /// <param name="sUser"></param>
        /// <param name="sPassword"></param>
        /// <returns></returns>
        public static string GetConnectionString(ConnectionType cType, string sServer, string sNameOf, string sName, 
			string sUser, string sPassword)
		{
			string sConn = "";

			//create connection string for selected database
			switch (cType)
			{
				case ConnectionType.SQLServer:
					//sql server
					sConn = _CONNECTION_PROVIDER + "=" + _CONNECTION_SQLOLEDB + ";"
						+ _CONNECTION_DATASOURCE + "=" + sServer + ";"
						+ _CONNECTION_DATABASE + "=" + sName + ";"
						+ _CONNECTION_USERID + "=" + sUser + ";"
						+ _CONNECTION_PASSWORD + "=" + sPassword + ";";
					break;
				case ConnectionType.Oracle:
					//oracle or microsoft oledb native provider
					sConn = _CONNECTION_PROVIDER + "=" + _CONNECTION_MSDAORA + ";"
						+ _CONNECTION_DATASOURCE + "=" + sNameOf + ";"
						+ _CONNECTION_USERID + "=" + sUser + ";"
						+ _CONNECTION_PASSWORD + "=" + sPassword + ";";
					break;
			}
			return( sConn );
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static int RunSQL(IDbConnection dbConn, string sql)
        {
            //DataAccess imedData = new DataAccess();
            IDbCommand dbCommand = null;

            try
            {
                // check is open connection
                if (dbConn.State != ConnectionState.Open)
                {
                    dbConn.Open();
                }
                // get command
                dbCommand = DataAccess.GetCommand(dbConn);
                // create command text
                dbCommand.CommandText = sql;
                return dbCommand.ExecuteNonQuery();
            }
            finally
            {
                //clean up objects
                dbCommand.Dispose();
            }
        }

        public static int RunSQL(string dbConnectionString, string sql)
        {
            // create connection
            IDbConnection idbConnection = DataAccess.GetConnection(dbConnectionString);
            // return result from overloaded function
            int rowsReturned = RunSQL(idbConnection, sql);
            // check is closed connection
            if (idbConnection.State == ConnectionState.Open)
            {
                idbConnection.Close();
            }
            // tidy
            idbConnection.Dispose();
            // return
            return rowsReturned;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="sql"></param>
        /// <param name="paramList"></param>
        /// <param name="storedProc"></param>
        /// <returns></returns>
        public static int RunSQL(IDbConnection dbConn, string sql, List<ParameterInfo> paramList, bool storedProc)
        {
            IDbCommand dbCommand = null;

            try
            {
                // check is open connection
                if (dbConn.State != ConnectionState.Open)
                {
                    dbConn.Open();
                }
                // get command
                dbCommand = DataAccess.GetCommand(dbConn);
                // create command text
                dbCommand.CommandText = sql;
                if (storedProc)
                {
                    dbCommand.CommandType = CommandType.StoredProcedure;
                }
                // add parameters to command
                AddParametersToCommand(ref dbCommand, paramList, storedProc, false);
                // execute & return rows affected
                return dbCommand.ExecuteNonQuery();
            }
            finally
            {
                //clean up objects
                dbCommand.Dispose();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dbConnectionString"></param>
        /// <param name="sql"></param>
        /// <param name="paramList"></param>
        /// <param name="storedProc"></param>
        /// <returns></returns>
        public static int RunSQL(string dbConnectionString, string sql, List<ParameterInfo> paramList, bool storedProc)
        {
            // create connection
            IDbConnection idbConnection = DataAccess.GetConnection(dbConnectionString);
            // return result from overloaded function
            int rowsReturned = RunSQL(idbConnection, sql, paramList, storedProc);
            // check is closed connection
            if (idbConnection.State == ConnectionState.Open)
            {
                idbConnection.Close();
            }
            // tidy
            idbConnection.Dispose();
            // return
            return rowsReturned;
        }

        /// <summary>
        /// Returns a dataset from the database
        /// </summary>
        /// <param name="connectionString">MACRO database connection string</param>
        /// <param name="sql">SQL to run against the MACRO database</param>
        /// <returns>Dataset containing MACRO data</returns>
        public static DataSet GetDataSet(string connectionString, string sql)
        {
            //DataAccess imedData = new DataAccess();
            IDbConnection dbConn = null;

            try
            {
                // open ImedDataAccess class
                DataAccess.ConnectionType imedConnType;
                // calculate connection type
                imedConnType = DataAccess.CalculateConnectionType(connectionString);
                // get connection & open
                dbConn = DataAccess.GetConnection(imedConnType, connectionString);
                dbConn.Open();

                return (GetDataSet(dbConn, sql));
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
            //DataAccess imedData = new DataAccess();
            IDbCommand dbCommand = null;
            IDbDataAdapter dbDataAdapter = null;
            DataSet dataDB = new DataSet();

            try
            {
                // get command
                dbCommand = DataAccess.GetCommand(dbConn);
                // create command text
                dbCommand.CommandText = sql;
                // get data adaptor
                dbDataAdapter = DataAccess.GetDataAdapter(dbCommand);
                // fill dataset
                dbDataAdapter.Fill(dataDB);

                return (dataDB);
            }
            finally
            {
                //clean up objects
                dbCommand.Dispose();
            }
        }

        public static DataSet GetDataSet(IDbConnection dbConn, string sql, 
                                            List<ParameterInfo> paramList, bool storedProc)
        {
            IDbCommand dbCommand = null;
            IDbDataAdapter dbDataAdapter = null;
            DataSet dataDB = new DataSet();

            try
            {
                if (dbConn.State != ConnectionState.Open)
                {
                    // open
                    dbConn.Open();
                }
                // get command
                dbCommand = DataAccess.GetCommand(dbConn);
                // create command text
                dbCommand.CommandText = sql;
                // is it stored procedure
                if(storedProc)
                {
                    dbCommand.CommandType = CommandType.StoredProcedure;
                }
                // deal with parameters
                AddParametersToCommand(ref dbCommand, paramList, storedProc, true);
                // get data adaptor
                dbDataAdapter = DataAccess.GetDataAdapter(dbCommand);
                // fill dataset
                dbDataAdapter.Fill(dataDB);
                // return datase
                return dataDB;
            }
            finally
            {
                // clean up
                dbCommand.Dispose();
            }
        }

        public static DataSet GetDataSet(string dbConnectionString, string sql, 
                                            List<ParameterInfo> paramList, bool storedProc)
        {
            // create connection
            IDbConnection idbConnection = DataAccess.GetConnection(dbConnectionString);
            // return result from overloaded function
            return GetDataSet(idbConnection, sql, paramList, storedProc);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dbComm"></param>
        /// <param name="paramList"></param>
        /// <param name="storedProc"></param>
        public static void AddParametersToCommand(ref IDbCommand dbCommand, List<ParameterInfo> paramList,
                                    bool storedProc, bool retrieveData)
        {
            // calculate connection type
            ConnectionType conType = DataAccess.CalculateConnectionType(dbCommand.Connection);
            
            // loop through parameters adding to command 
            for (int i = 0; i < paramList.Count; i++)
            {
                // get parameter information
                ParameterInfo parameterInfo = paramList[i];

                // idbparameter
                IDbDataParameter idbParameter;

                switch (conType)
                {
                    case ConnectionType.Oracle: // Oracle Data Provider
                        {
                            OracleType oracleType = DbTypeTranslationToOracle(parameterInfo.DbGenericType);
                            // add parameter
                            idbParameter = ((OracleCommand)dbCommand).Parameters.Add(parameterInfo.ParameterName, oracleType);
                            // direction
                            idbParameter.Direction = parameterInfo.ParamDirection;
                            // value
                            idbParameter.Value = parameterInfo.ParamValue;
                            break;
                        }
                    case ConnectionType.OleDb: // OleDb Data Provider
                        {
                            OleDbType oleDbType = DbTypeTranslationToOleDb(parameterInfo.DbGenericType);
                            // add parameter
                            idbParameter = ((OleDbCommand)dbCommand).Parameters.Add(parameterInfo.ParameterName, oleDbType);
                            // direction
                            idbParameter.Direction = parameterInfo.ParamDirection;
                            // value
                            idbParameter.Value = parameterInfo.ParamValue;
                            break;
                        }
                    case ConnectionType.SQLServer: // Sql Data Provider
                        {
                            SqlDbType sqlDbType = DbTypeTranslationToSqlServer(parameterInfo.DbGenericType);
                            // add parameter
                            idbParameter = ((SqlCommand)dbCommand).Parameters.Add(parameterInfo.ParameterName, sqlDbType);
                            // direction
                            idbParameter.Direction = parameterInfo.ParamDirection;
                            // value
                            idbParameter.Value = parameterInfo.ParamValue;
                            break;
                        }
                    default:
                        break;
                }
            }

            // if is stored procedure on an oracle database
            if ((retrieveData) && (storedProc) && (conType == ConnectionType.Oracle))
            {
                // add final parameter for oracle databases
                // idbparameter
                IDbDataParameter idbParameter;
                // oracle data type
                OracleType oracleType = OracleType.Cursor;
                // add parameter
                idbParameter = ((OracleCommand)dbCommand).Parameters.Add(_ORACLE_OUTCURSOR_PARAM_NAME, oracleType);
                // direction
                idbParameter.Direction = ParameterDirection.Output;
            }
        }

        /// <summary>
        /// Translate the generic database datatype to an oracle specific one
        /// </summary>
        /// <param name="dbType"></param>
        /// <returns></returns>
        public static OracleType DbTypeTranslationToOracle(DbType dbType)
        {
            OracleType oracleType = OracleType.VarChar;

            switch (dbType)
            {
                case DbType.AnsiString:
                    {
                        // non-unicode
                        oracleType = OracleType.VarChar;
                        break;
                    }
                case DbType.Binary:
                    {
                        // binary stream
                        oracleType = OracleType.Blob;
                        break;
                    }
                case DbType.Boolean:
                    {
                        // boolean
                        oracleType = OracleType.Int16;
                        break;
                    }
                case DbType.Byte:
                    {
                        // byte
                        oracleType = OracleType.Byte;
                        break;
                    }
                case DbType.Currency:
                    {
                        // currency
                        oracleType = OracleType.Number;
                        break;
                    }
                case DbType.Date:
                    {
                        // date 
                        oracleType = OracleType.Timestamp;
                        break;
                    }
                case DbType.DateTime:
                    {
                        // date/time
                        oracleType = OracleType.DateTime;
                        break;
                    }
                case DbType.Decimal:
                    {
                        oracleType = OracleType.Number;
                        break;
                    }
                case DbType.Double:
                    {
                        oracleType = OracleType.Double;
                        break;
                    }
                case DbType.Guid:
                    {
                        oracleType = OracleType.RowId;
                        break;
                    }
                case DbType.Int16:
                    {
                        // short
                        oracleType = OracleType.Int16;
                        break;
                    }
                case DbType.Int32:
                    {
                        // integer
                        oracleType = OracleType.Int32;
                        break;
                    }
                case DbType.Int64:
                    {
                        oracleType = OracleType.Number;
                        break;
                    }
                case DbType.Object:
                    {
                        oracleType = OracleType.Cursor;
                        break;
                    }
                case DbType.SByte:
                    {
                        oracleType = OracleType.SByte;
                        break;
                    }
                case DbType.Single:
                    {
                        oracleType = OracleType.Number;
                        break;
                    }
                case DbType.String:
                    {
                        // unicode string
                        oracleType = OracleType.NVarChar;
                        break;
                    }
                case DbType.Time:
                    {
                        oracleType = OracleType.IntervalDayToSecond;
                        break;
                    }
                case DbType.UInt16:
                    {
                        // unisigned short
                        oracleType = OracleType.UInt16;
                        break;
                    }
                case DbType.UInt32:
                    {
                        // unsigned integer
                        oracleType = OracleType.UInt32;
                        break;
                    }
                case DbType.UInt64:
                    {
                        oracleType = OracleType.Number;
                        break;
                    }
                case DbType.VarNumeric:
                    {
                        oracleType = OracleType.Number;
                        break;
                    }
                case DbType.Xml:
                    {
                        oracleType = OracleType.NVarChar;
                        break;
                    }
            }
            return oracleType;
        }

        /// <summary>
        /// Translate the generic database datatype to an sql server specific one
        /// </summary>
        /// <param name="dbType"></param>
        /// <returns></returns>
        public static SqlDbType DbTypeTranslationToSqlServer(DbType dbType)
        {
            SqlDbType sqlType = SqlDbType.VarChar;

            switch (dbType)
            {
                case DbType.AnsiString:
                    {
                        // non-unicode
                        sqlType = SqlDbType.VarChar;
                        break;
                    }
                case DbType.Binary:
                    {
                        // binary stream
                        sqlType = SqlDbType.Binary;
                        break;
                    }
                case DbType.Boolean:
                    {
                        // boolean
                        sqlType = SqlDbType.Bit;
                        break;
                    }
                case DbType.Byte:
                    {
                        // byte
                        sqlType = SqlDbType.Variant;
                        break;
                    }
                case DbType.Currency:
                    {
                        // currency
                        sqlType = SqlDbType.Money;
                        break;
                    }
                case DbType.Date:
                    {
                        // date 
                        sqlType = SqlDbType.DateTime;
                        break;
                    }
                case DbType.DateTime:
                    {
                        // date/time
                        sqlType = SqlDbType.Variant;
                        break;
                    }
                case DbType.Decimal:
                    {
                        sqlType = SqlDbType.Decimal;
                        break;
                    }
                case DbType.Double:
                    {
                        sqlType = SqlDbType.Real;
                        break;
                    }
                case DbType.Guid:
                    {
                        sqlType = SqlDbType.UniqueIdentifier;
                        break;
                    }
                case DbType.Int16:
                    {
                        // short
                        sqlType = SqlDbType.SmallInt;
                        break;
                    }
                case DbType.Int32:
                    {
                        // integer
                        sqlType = SqlDbType.Int;
                        break;
                    }
                case DbType.Int64:
                    {
                        sqlType = SqlDbType.BigInt;
                        break;
                    }
                case DbType.SByte:
                    {
                        sqlType = SqlDbType.Variant;
                        break;
                    }
                case DbType.Single:
                    {
                        sqlType = SqlDbType.Variant;
                        break;
                    }
                case DbType.String:
                    {
                        // unicode string
                        sqlType = SqlDbType.NVarChar;
                        break;
                    }
                case DbType.Time:
                    {
                        sqlType = SqlDbType.Variant;
                        break;
                    }
                case DbType.UInt16:
                    {
                        // unisigned short
                        sqlType = SqlDbType.Variant;
                        break;
                    }
                case DbType.UInt32:
                    {
                        // unsigned integer
                        sqlType = SqlDbType.Variant;
                        break;
                    }
                case DbType.UInt64:
                    {
                        sqlType = SqlDbType.Variant;
                        break;
                    }
                case DbType.VarNumeric:
                    {
                        sqlType = SqlDbType.Variant;
                        break;
                    }
                case DbType.Xml:
                    {
                        sqlType = SqlDbType.Xml;
                        break;
                    }
            }

            return sqlType;
        }

        public static OleDbType DbTypeTranslationToOleDb(DbType dbType)
        {
            OleDbType oleDbType = OleDbType.VarChar;

            switch (dbType)
            {
                case DbType.AnsiString:
                    {
                        // non-unicode
                        oleDbType = OleDbType.VarChar;
                        break;
                    }
                case DbType.Binary:
                    {
                        // binary stream
                        oleDbType = OleDbType.Binary;
                        break;
                    }
                case DbType.Boolean:
                    {
                        // boolean
                        oleDbType = OleDbType.Boolean;
                        break;
                    }
                case DbType.Byte:
                    {
                        // byte
                        oleDbType = OleDbType.VarBinary;
                        break;
                    }
                case DbType.Currency:
                    {
                        // currency
                        oleDbType = OleDbType.Currency;
                        break;
                    }
                case DbType.Date:
                    {
                        // date 
                        oleDbType = OleDbType.Date;
                        break;
                    }
                case DbType.DateTime:
                    {
                        // date/time
                        oleDbType = OleDbType.DBDate;
                        break;
                    }
                case DbType.Decimal:
                    {
                        oleDbType = OleDbType.Decimal;
                        break;
                    }
                case DbType.Double:
                    {
                        oleDbType = OleDbType.Double;
                        break;
                    }
                case DbType.Guid:
                    {
                        oleDbType = OleDbType.Guid;
                        break;
                    }
                case DbType.Int16:
                    {
                        // short
                        oleDbType = OleDbType.SmallInt;
                        break;
                    }
                case DbType.Int32:
                    {
                        // integer
                        oleDbType = OleDbType.Integer;
                        break;
                    }
                case DbType.Int64:
                    {
                        oleDbType = OleDbType.BigInt;
                        break;
                    }
                case DbType.SByte:
                    {
                        oleDbType = OleDbType.Variant;
                        break;
                    }
                case DbType.Single:
                    {
                        oleDbType = OleDbType.Single;
                        break;
                    }
                case DbType.String:
                    {
                        // unicode string
                        oleDbType = OleDbType.VarWChar;
                        break;
                    }
                case DbType.Time:
                    {
                        oleDbType = OleDbType.DBTime;
                        break;
                    }
                case DbType.UInt16:
                    {
                        // unisigned short
                        oleDbType = OleDbType.UnsignedSmallInt;
                        break;
                    }
                case DbType.UInt32:
                    {
                        // unsigned integer
                        oleDbType = OleDbType.UnsignedInt;
                        break;
                    }
                case DbType.UInt64:
                    {
                        oleDbType = OleDbType.UnsignedBigInt;
                        break;
                    }
                case DbType.VarNumeric:
                    {
                        oleDbType = OleDbType.VarNumeric;
                        break;
                    }
                case DbType.Xml:
                    {
                        oleDbType = OleDbType.VarChar;
                        break;
                    }
            }
            return oleDbType;
        }
    }
}
