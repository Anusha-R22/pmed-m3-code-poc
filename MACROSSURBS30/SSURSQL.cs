/* ------------------------------------------------------------------------
 * File: SSURSQL.cs
 * Author: Nicky Johns
 * Copyright: InferMed, 2008, All Rights Reserved
 * Purpose: SQL for Study/Site/User/Role associations for MACRO 3.0 API
 * Date: March 2008
 * ------------------------------------------------------------------------
 * Revisions:
 * 
 * ------------------------------------------------------------------------*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace MACROSSURBS30
{
    /// <summary>
    /// Contains database access routines for MACRO User Roles
    /// </summary>
    public static class SSURSQL
    {
        public const string ALL_STUDIES = "AllStudies";
        public const string ALL_SITES = "AllSites";

        /// <summary>
        /// Get all the Study/Site/User/Role associations matching the given parameters (case-insensitive)
        /// from the MACRO USERROLE table
        /// </summary>
        /// <param name="dbcon">MACRO database connection string</param>
        /// <param name="study">Study name (may be "")</param>
        /// <param name="site">Site code (may be "")</param>
        /// <param name="user">User name (may be "")</param>
        /// <param name="role">Role code (may be "")</param>
        /// <returns>Datatable containing matching database rows</returns>
        public static DataTable GetAssocs(string dbcon, string study, string site, string user, string role)
        {
            // We need to handle case-sensitivity for Oracle
            DataAccess.ConnectionType conType = DataAccess.CalculateConnectionType(dbcon);

            string sql = GetAssocsSQL(conType, study, site, user, role);
            DataSet ds = DataAccess.GetDataSet(dbcon, sql);

            // Did we get anything?
            if (ds == null) return null;
            if (ds.Tables[0].Rows.Count == 0) return null;
            return ds.Tables[0];
        }

        // Get the SQL to retrieve associations matching the given parameters
        private static string GetAssocsSQL(DataAccess.ConnectionType conType,
                            string study, string site, string user, string role)
        {
            StringBuilder sqlSB = new StringBuilder("");
            sqlSB.Append("SELECT STUDYCODE, SITECODE, USERNAME, ROLECODE FROM USERROLE");

            // For the first one we use WHERE
            string whereOrAnd = " WHERE ";
            sqlSB.Append(GetSQLFilter("STUDYCODE", study, conType, ref whereOrAnd));
            sqlSB.Append(GetSQLFilter("SITECODE", site, conType, ref whereOrAnd));
            sqlSB.Append(GetSQLFilter("USERNAME", user, conType, ref whereOrAnd));
            sqlSB.Append(GetSQLFilter("ROLECODE", role, conType, ref whereOrAnd));

            return sqlSB.ToString();
        }

        // Get an SQL filter using WHERE or AND
        private static string GetSQLFilter(string field, string val,
                                    DataAccess.ConnectionType conType, ref string whereOrAnd)
        {
            string sql = "";
            if (val != "")
            {
                sql = whereOrAnd + FieldEquals(field, val, conType);
                // For subsequent filters we'll use AND
                whereOrAnd = " AND ";
            }
            return sql;
        }

        /// <summary>
        /// Does this user exist in the given Security Database?
        /// </summary>
        /// <param name="secDBCon">Security database connection string</param>
        /// <param name="user">User code (case-insensitive)</param>
        /// <returns>True if user exists, with user = official value; false otherwise, with user unchanged</returns>
        public static Boolean UserExists(string secDBCon, ref string user)
        {
            // We need to handle case-sensitivity for Oracle
            DataAccess.ConnectionType conType = DataAccess.CalculateConnectionType(secDBCon);

            string sql = "SELECT USERNAME FROM MACROUSER WHERE "
                + FieldEquals("USERNAME", user, conType);

            string found = GetSingleDBField(secDBCon, sql);
            if (found != "")
            {
                // Send back the "real" value
                user = found;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Does this role exist in the given Security Database?
        /// </summary>
        /// <param name="secDBCon">Security database connection string</param>
        /// <param name="role">Role code (case-insensitive)</param>
        /// <returns>True if role exists, with role = official value; false otherwise, with role unchanged</returns>
        public static Boolean RoleExists(string secDBCon, ref string role)
        {
            // We need to handle case-sensitivity for Oracle
            DataAccess.ConnectionType conType = DataAccess.CalculateConnectionType(secDBCon);

            string sql = "SELECT ROLECODE FROM ROLE WHERE "
                + FieldEquals("ROLECODE", role, conType);

            string found = GetSingleDBField(secDBCon, sql);
            if (found != "")
            {
                // Send back the "real" value
                role = found;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Does this study exist in the given Database?
        /// </summary>
        /// <param name="dbCon">MACRO database connection string</param>
        /// <param name="study">Study name (case-insensitive)</param>
        /// <returns>True if site exists, with study = official value; otherwise false with study unchanged</returns>
        public static Boolean StudyExists(string dbCon, ref string study)
        {
            // Special case - AllStudies
            if (study.ToUpper() == ALL_STUDIES.ToUpper())
            {
                study = ALL_STUDIES;
                return true;
            }

            // We need to handle case-sensitivity for Oracle
            DataAccess.ConnectionType conType = DataAccess.CalculateConnectionType(dbCon);

            string sql = "SELECT CLINICALTRIALNAME FROM CLINICALTRIAL WHERE "
                + FieldEquals("CLINICALTRIALNAME", study, conType);

            string foundStudy = GetSingleDBField(dbCon, sql);
            if (foundStudy != "")
            {
                // Send back the "real" value
                study = foundStudy;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Does this site exist in the given Database?
        /// </summary>
        /// <param name="dbCon">MACRO database connection string</param>
        /// <param name="site">Site code (case-insensitive)</param>
        /// <returns>True if site exists, with site = official value; otherwise false with site unchanged</returns>
        public static Boolean SiteExists(string dbCon, ref string site)
        {
            // Special case - AllSites
            if (site.ToUpper() == ALL_SITES.ToUpper())
            {
                site = ALL_SITES;
                return true;
            }

             // We need to handle case-sensitivity for Oracle
            DataAccess.ConnectionType conType = DataAccess.CalculateConnectionType(dbCon);

            string sql = "SELECT SITE FROM SITE WHERE "
                 + FieldEquals("SITE", site, conType);

            string found = GetSingleDBField(dbCon, sql);
            if (found != "")
            {
                // Send back the "real" value
                site = found;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Does this study/site association exist in the given Database?
        /// </summary>
        /// <param name="dbCon">MACRO database connection string</param>
        /// <param name="study">Study name - assumed correct case</param>
        /// <param name="site">Site code - assumed correct case</param>
        /// <returns>True if site participatig in study; false if not</returns>
        public static Boolean StudySiteExists(string dbCon, string study, string site)
        {
            // Special case - AllStudies, AllSites
            if (site == ALL_SITES || study == ALL_STUDIES)
                return true;

            string sql = "SELECT * FROM TRIALSITE, CLINICALTRIAL WHERE "
                + FieldEquals("TRIALSITE", site)
                + " AND TRIALSITE.CLINICALTRIALID = CLINICALTRIAL.CLINICALTRIALID AND "
                + FieldEquals("CLINICALTRIAL.CLINICALTRIALNAME", study);

            return RowsExist(dbCon, sql);
        }

        /// <summary>
        /// Does this association exist in the given MACRO Database?
        /// All parameters assumed to be correct case
        /// </summary>
        /// <param name="dbCon">DB connection string</param>
        /// <param name="study">Study name</param>
        /// <param name="site">Site code</param>
        /// <param name="user">User name</param>
        /// <param name="role">Role code</param>
        /// <returns>True if row already exists in DB</returns>
        public static Boolean AssocExists(string dbCon, string study, string site, string user, string role)
        {
            string sql = "SELECT * FROM USERROLE WHERE "
                + FieldEquals("STUDYCODE", study)
                + " AND " + FieldEquals("SITECODE", site)
                + " AND " + FieldEquals("USERNAME", user)
                + " AND " + FieldEquals("ROLECODE", role);

            return RowsExist(dbCon, sql);
        }

        /// <summary>
        /// Does this user have any roles in this DB?
        /// </summary>
        /// <param name="secDBCon">MACRO database connection string</param>
        /// <param name="userName">User name - assumed correct case</param>
        /// <returns>True if user has at least one role in this DB; false otherwise</returns>
        public static Boolean UserHasRolesInDB(string dbCon, string userName)
        {
            string sql = "SELECT * FROM USERROLE WHERE "
                + FieldEquals("USERNAME", userName);

            return RowsExist(dbCon, sql);
        }

        /// <summary>
        /// Do we get any rows returned from this SQL query?
        /// </summary>
        /// <param name="dbCon">Database connection string</param>
        /// <param name="sql">SQL query</param>
        /// <returns>True if rows exist, False if not</returns>
        private static Boolean RowsExist(string dbCon, string sql)
        {
            DataSet ds = DataAccess.GetDataSet(dbCon, sql);
            // Did we get anything?
            if (ds == null) return false;
            if (ds.Tables[0].Rows.Count == 0) return false;
            // Looks like we found something
            return true;
        }

        /// <summary>
        /// Get a single field value from the database
        /// </summary>
        /// <param name="dbCon">Database connection string</param>
        /// <param name="sql">SQL to retrieve single value</param>
        /// <returns>Single value as string</returns>
        private static string GetSingleDBField(string dbCon, string sql)
        {
            DataSet ds = DataAccess.GetDataSet(dbCon, sql);
            // Did we get anything?
            if (ds == null) return "";
            if (ds.Tables[0].Rows.Count == 0) return "";
            // Looks like we found something - return the first item from the first row
            return ds.Tables[0].Rows[0][0].ToString();
        }

        /// <summary>
        /// Retrieve all the fields from the MACROUser table for this user
        /// </summary>
        /// <param name="secDBCon">Security database connection string</param>
        /// <param name="userName">User name - assumed correct case</param>
        /// <returns>Data table of results</returns>
        public static DataTable GetUserDetails(string secDBCon, string userName)
        {
            string sql = "SELECT * FROM MACROUSER WHERE "
                + FieldEquals( "USERNAME", userName);
            DataSet ds = DataAccess.GetDataSet(secDBCon, sql);

            // Did we get anything?
            if (ds == null) return null;
            if (ds.Tables[0].Rows.Count == 0) return null;
            return ds.Tables[0];
        }

        /// <summary>
        /// Delete a user role association from the MACRO database
        /// All parameters assumed to be correct case
        /// </summary>
        /// <param name="dbCon">DB connection string</param>
        /// <param name="study">Study name</param>
        /// <param name="site">Site code</param>
        /// <param name="user">User name</param>
        /// <param name="role">Role code</param>
        public static void DeleteRole(string dbCon, string study, string site, string user, string role)
        {
            string sql = "DELETE FROM USERROLE WHERE "
                + FieldEquals("STUDYCODE", study)
                + " AND " + FieldEquals("SITECODE", site)
                + " AND " + FieldEquals("USERNAME", user)
                + " AND " + FieldEquals("ROLECODE", role);

            DataAccess.RunSQL(dbCon, sql);
        }

        /// <summary>
        /// Get the SQL string for case-insensitive 'Field = Value' depending on database type
        /// For Oracle we convert to upper case
        /// </summary>
        /// <param name="field">Database field (column) name</param>
        /// <param name="value">Value to be compared</param>
        /// <param name="conType">Database connection type (Oracle or SQLServer)</param>
        /// <returns>Appropriate SQL fragment</returns>
        private static string FieldEquals(string field, string value, DataAccess.ConnectionType conType)
        {
            // Handle case-sensitivity for Oracle
            if (conType == DataAccess.ConnectionType.Oracle)
                return "NLS_UPPER(" + field + ") = NLS_UPPER('" + value + "')";
            else
                return field + " = '" + value + "'";
        }

        /// <summary>
        /// Get SQL fragment for case-sensitive Field = Value 
        /// </summary>
        /// <param name="field">Database column name</param>
        /// <param name="value">String value to compare with</param>
        /// <returns>SQL string</returns>
        private static string FieldEquals(string field, string value)
        {
            return field + " = '" + value + "'";
        }

        /// <summary>
        /// Insert a new user role into the USERROLE table in the MACRO database
        /// </summary>
        /// <param name="dbCon">MACRO Database connection string</param>
        /// <param name="study">Study name</param>
        /// <param name="site">Site code</param>
        /// <param name="user">User name</param>
        /// <param name="role">Role code</param>
        public static void InsertRole(string dbCon, string study, string site, string user, string role)
        {
            string sql = "INSERT INTO USERROLE "
                + "(USERNAME, ROLECODE, STUDYCODE, SITECODE, TYPEOFINSTALLATION) "
                + "VALUES ('" + user + "', '" + role + "', '"
                + study + "', '" + site + "', 1)";

            DataAccess.RunSQL(dbCon, sql);
        }

        /// <summary>
        /// Delete row from the USERDATABASES table for this Database Code and User.
        /// </summary>
        /// <param name="secDBCon">Security database connection string</param>
        /// <param name="dbCode">MACRO database code (assumed correct case)</param>
        /// <param name="user">User name (assumed correct case)</param>
        public static void DeleteDBUser(string secDBCon, string dbCode, string user)
        {
            string sql = "DELETE FROM USERDATABASE WHERE "
                + FieldEquals("DATABASECODE", dbCode)
                + " AND " + FieldEquals("USERNAME", user);

            DataAccess.RunSQL(secDBCon, sql);
        }

        /// <summary>
        /// Add a row to the USERDATABASES table for this Database Code and User.
        /// Does nothing if row already exists.
        /// </summary>
        /// <param name="secDBCon">Security database connection string</param>
        /// <param name="dbCode">MACRO database code (assumed correct case)</param>
        /// <param name="user">User name (assumed correct case)</param>
        public static void AddDBUser(string secDBCon, string dbCode, string user)
        {
            string sql = "SELECT * FROM USERDATABASE WHERE "
                + FieldEquals("DATABASECODE", dbCode)
                + " AND " + FieldEquals("USERNAME", user);

            // Only add row if it doesn't already exist
            if (!RowsExist(secDBCon, sql))
            {
                sql = "INSERT INTO USERDATABASE (DATABASECODE, USERNAME) "
                    + "VALUES ('" + dbCode + "', '" + user + "')";

                DataAccess.RunSQL(secDBCon, sql);
            }
        }

    }
}
