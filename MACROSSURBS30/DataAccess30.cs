/* ------------------------------------------------------------------------
 * File: DataAcess30.cs
 * Author: Nicky Johns
 * Copyright: InferMed, 2008, All Rights Reserved
 * Purpose: Data Access class copied from MACROCATBS30 for the SSURIO project for the MACRO 3.0 API
 * Date: November 2007
 * ------------------------------------------------------------------------
 * Revisions:
 * NCJ 3 March 2008 - Copied from DataAcess30 for MACROCATBS30
 * ------------------------------------------------------------------------*/
using System;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.OracleClient;
using System.Text;

namespace MACROSSURBS30
{
    public class DataAccess : IDisposable
    {
        /// <summary>
        /// DataAccess class 
        /// </summary>
        private IDbDataAdapter idbAdapter = null;
        private IDataReader idbReader = null;
        private IDbDataParameter idbParameter = null;


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

        public enum ConnectionType
        {
            Unknown = -1,
            Oracle = 1,
            SQLServer = 2,
            OleDb = 3
        }

        /// <summary>
        /// constructor
        /// </summary>
        public DataAccess()
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
            ConnectionType ConnType = ConnectionType.OleDb;
            // check for specific provider types
            // oracle
            if ((sConString.IndexOf(_CONNECTION_MSDAORA) > 0) || (sConString.IndexOf(_CONNECTION_ORAOLEDB) > 0) || (sConString.IndexOf("ORAOLEDB.ORACLE.1") > 0))
            {
                ConnType = ConnectionType.Oracle;
            }
            // sql server
            // NCJ 29 Nov 07 - Changed to _CONNECTION_SQLOLEDB
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
            switch (DbConnection.GetType().FullName)
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

            switch (conType)
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
            foreach (string sStringPart in aString)
            {
                switch (conType)
                {
                    case ConnectionType.Oracle:
                        {
                            // remove provider=x; from string
                            // remove database=x; for oracle
                            // NCJ 29 Nov 07 - Remove spaces!
                            if (((sStringPart.Replace(" ", "").ToLower()).IndexOf("provider=") == -1) &&
                                ((sStringPart.Replace(" ", "").ToLower()).IndexOf("database=") == -1))
                            {
                                sbConn.Append(sStringPart + ";");
                            }
                            break;
                        }
                    case ConnectionType.SQLServer:
                        {
                            // remove provider=x; from string
                            if ((sStringPart.ToLower()).IndexOf("provider=") == -1)
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
            switch (conType)
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
            switch (conType)
            {
                case ConnectionType.Oracle: // Oracle Data Provider
                    idbCommand = new OracleCommand(sSql, (OracleConnection)DbConnection);
                    break;
                case ConnectionType.SQLServer: // Sql Data Provider
                    idbCommand = new SqlCommand(sSql, (SqlConnection)DbConnection);
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
            // calculate connection type
            ConnectionType conType = CalculateConnectionType(DbConnection);
            IDbCommand idbCommand = null;
            switch (conType)
            {
                case ConnectionType.Oracle: // Oracle Data Provider
                    idbCommand = new OracleCommand(sSql, (OracleConnection)DbConnection);
                    break;
                case ConnectionType.SQLServer: // Sql Data Provider
                    idbCommand = new SqlCommand(sSql, (SqlConnection)DbConnection);
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
            switch (conType)
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
            // calculate connection type
            ConnectionType conType = CalculateConnectionType(DbConnection);
            IDbCommand idbCommand = null;
            switch (conType)
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
        /// GetDataAdapter returns IDbDataAdapter
        /// overloaded 3 times
        /// Connection type, connection string, sql
        /// </summary>
        /// <param name="conType"></param>
        /// <param name="sConString"></param>
        /// <param name="sSql"></param>
        /// <returns></returns>
        public IDbDataAdapter GetDataAdapter(ConnectionType conType,
            string sConString, string sSql)
        {
            // Format connection string for connection type
            sConString = FormatConnectionString(conType, sConString);
            switch (conType)
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
        public IDbDataAdapter GetDataAdapter(ConnectionType conType,
            IDbConnection DbConnection, string sSql)
        {
            switch (conType)
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
        public IDbDataAdapter GetDataAdapter(IDbConnection DbConnection, string sSql)
        {
            // calculate connection type
            ConnectionType conType = CalculateConnectionType(DbConnection);
            switch (conType)
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
        public IDbDataAdapter GetDataAdapter(ConnectionType conType,
            IDbCommand DbCommand)
        {
            switch (conType)
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
        public IDbDataAdapter GetDataAdapter(IDbCommand DbCommand)
        {
            // calculate connection type
            ConnectionType conType = CalculateConnectionType(DbCommand.Connection);
            switch (conType)
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
        /// <param name="dbConn"></param>
        /// <param name="sql"></param>
        public static void RunSQL(IDbConnection dbConn, string sql)
        {
            IDbCommand dbCommand = null;

            try
            {
                // get command
                dbCommand = GetCommand(dbConn);
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

        public static void RunSQL(string dbConnectionString, string sql)
        {
            // create connection
            IDbConnection idbConnection = DataAccess.GetConnection(dbConnectionString);
            idbConnection.Open();
            // return result from overloaded function
            RunSQL(idbConnection, sql);
            // check is closed connection
            if (idbConnection.State == ConnectionState.Open)
            {
                idbConnection.Close();
            }
            // tidy
            idbConnection.Dispose();
        }

        /// <summary>
        /// Returns a dataset from the database
        /// </summary>
        /// <param name="connectionString">MACRO database connection string</param>
        /// <param name="sql">SQL to run against the MACRO database</param>
        /// <returns>Dataset containing MACRO data</returns>
        public static DataSet GetDataSet(string connectionString, string sql)
        {
            IDbConnection dbConn = null;

            try
            {
                // open ImedDataAccess class
                DataAccess.ConnectionType imedConnType;
                // calculate connection type
                imedConnType = DataAccess.CalculateConnectionType(connectionString);
                // get connection & open
                dbConn = GetConnection(imedConnType, connectionString);
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
            DataAccess imedData = new DataAccess();
            IDbCommand dbCommand = null;
            IDbDataAdapter dbDataAdapter = null;
            DataSet dataDB = new DataSet();

            try
            {
                // get command
                dbCommand = GetCommand(dbConn);
                // create command text
                dbCommand.CommandText = sql;
                // get data adaptor
                dbDataAdapter = imedData.GetDataAdapter(dbCommand);
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

        #region IDisposable Members

        public void Dispose()
        {
            // empty / close objects that may be open
            if (idbParameter != null)
            {
                idbParameter = null;
            }
            if (idbReader != null)
            {
                if (!idbReader.IsClosed)
                {
                    idbReader.Close();
                }
            }
            if (idbAdapter != null)
            {
                idbAdapter = null;
            }
        }

        #endregion
    }
}
