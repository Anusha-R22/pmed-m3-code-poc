/* ------------------------------------------------------------------------
 * File: CatSQL.cs
 * Author: Nicky Johns
 * Copyright: InferMed, 2007, All Rights Reserved
 * Purpose: DB Access routines for categories for the MACRO 3.0 (and M4) API
 * Date: November 2007
 * ------------------------------------------------------------------------
 * Revisions:
 *  NCJ 28 Nov 07 - Deal with Oracle (case-sensitive) and SQL Server
 *  NCJ 29 Nov 07 - Fixing bugs during WBT-ing
 *  NCJ 5 Dec 07 - Use dbCommands for when transaction control is required.
 * * ------------------------------------------------------------------------*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace MACROCATBS30
{
    /// <summary>
    /// Handles all the SQL database access for Category handling for MACRO 3.0 API
    /// </summary>
    public static class CatSQL
    {
        /// <summary>
        /// Get the study ID for the given StudyName
        /// </summary>
        /// <param name="dbcon">DB connection string</param>
        /// <param name="studyname">Study name</param>
        /// <returns>ClinicalTrialId, or 0 if study not found</returns>
        public static int GetStudyId(string dbcon, string studyname)
        {
            string sql = CatSQL.StudyIdSQL(dbcon, studyname);

            DataSet ds = DataAccess.GetDataSet(dbcon, sql);

            // Did we get anything?
            if (ds == null) return 0;
            if (ds.Tables[0].Rows.Count == 0) return 0;

            return Convert.ToInt32(ds.Tables[0].Rows[0][0].ToString());
        }

        /// <summary>
        /// Get data item id for given data item code
        /// </summary>
        /// <param name="dbcon">DB connection string</param>
        /// <param name="studyid">Study ID</param>
        /// <param name="qcode">Data Item Code</param>
        /// <returns>DataItemId, or 0 if data item not found</returns>
        public static int GetDataItemId(string dbcon, int studyid, string qcode)
        {
            string sql = CatSQL.QuestionIdSQL(dbcon, studyid, qcode);

            DataSet ds = DataAccess.GetDataSet(dbcon, sql);

            // Did we get anything?
            if (ds == null) return 0;
            if (ds.Tables[0].Rows.Count == 0) return 0;
            return Convert.ToInt32(ds.Tables[0].Rows[0]["DataItemId"].ToString());
        }

        // Get data item id and type for given data item code
        // Returns datatype in dtype (-1 if data item not found)
        public static int GetDataItemId(string dbcon, int studyid, string qname, out int dtype)
        {
            dtype = -1;
            string sql = CatSQL.QuestionIdSQL(dbcon, studyid, qname);

            DataSet ds = DataAccess.GetDataSet(dbcon, sql);

            // Did we get anything?
            if (ds == null) return 0;
            if (ds.Tables[0].Rows.Count == 0) return 0;

            dtype = Convert.ToInt32(ds.Tables[0].Rows[0]["DataType"].ToString());
            return Convert.ToInt32(ds.Tables[0].Rows[0]["DataItemId"].ToString());
        }

        /// <summary>
        /// Get a list of codes of category questions in the given study
        /// </summary>
        /// <param name="dbcon">Database connection string</param>
        /// <param name="studyId">Study ID</param>
        /// <returns>Data table of question codes</returns>
        public static DataTable GetCatQuestionCodes(string dbcon, int studyId)
        {
            string sql = GetDItemsSQL(studyId, 1);
            DataSet ds = DataAccess.GetDataSet(dbcon, sql);

            // Did we get anything?
            if (ds == null) return null;
            if (ds.Tables[0].Rows.Count == 0) return null;
            return ds.Tables[0];
        }

        private static string StudyIdSQL(string dbcon, string studyName)
        {
            string whereName;
            DataAccess.ConnectionType conType = DataAccess.CalculateConnectionType(dbcon);
            // Oracle is case-sensitive
            if (conType == DataAccess.ConnectionType.Oracle)
                whereName = "NLS_UPPER(ClinicalTrialName) = NLS_UPPER('" + studyName + "')";
            else
                whereName = "ClinicalTrialName = '" + studyName + "'";
            return "SELECT ClinicalTrialId FROM ClinicalTrial WHERE " + whereName;
        }

        private static string QuestionIdSQL(string dbcon, int studyid, string qname)
        {
            string whereName;
            DataAccess.ConnectionType conType = DataAccess.CalculateConnectionType(dbcon);
            // Oracle is case-sensitive
            if (conType == DataAccess.ConnectionType.Oracle)
                whereName = "NLS_UPPER(DataItemCode) = NLS_UPPER('" + qname + "')";
            else
                whereName = "DataItemCode = '" + qname + "'";

            string sql = "SELECT DataItemId, DataType FROM DataItem"
                    + " WHERE " + whereName
                    + " AND ClinicalTrialId = " + studyid;

            return sql;
        }

        /// <summary>
        /// SQL to retrieve all category values for the given study
        /// </summary>
        /// <param name="studyId">Study Id</param>
        /// <returns>SQL to retrieve data set</returns>
        public static string GetCatsSQL(int studyId)
        {
            string sql = "SELECT DataItemCode, ValueCode, ItemValue, Active, ValueOrder"
            + " FROM ValueData cat, DataItem di "
            + " WHERE cat.ClinicalTrialId = " + studyId
            + " AND di.ClinicalTrialId = cat.ClinicalTrialId"
            + " AND di.DataItemId = cat.DataItemId"
            + " ORDER BY DataItemCode, ValueOrder";

            return sql;
        }

        // Get SQL to retrieve all category items for given question in specified study
        private static string GetCatsSQL(int studyId, int qid)
        {
            string sql = "SELECT ValueCode, ItemValue, Active, ValueOrder FROM ValueData "
            + " WHERE ClinicalTrialId = " + studyId
            + " AND DataItemId = " + qid
            + " ORDER BY ValueOrder";

            return sql;
        }

        // Get SQL to retrieve all question codes of given datatype in specified study
        private static string GetDItemsSQL(int studyId, int dataType)
        {
            return "SELECT DataItemCode FROM DataItem"
              + " WHERE ClinicalTrialId = " + studyId
              + " AND DataType = " + dataType;
        }

        // Delete ALL categories for the given question in the study
        // using an open database command object
        public static void DelCats(IDbCommand dbComm, int studyid, int qid)
        {
            dbComm.CommandText = "DELETE FROM VALUEDATA WHERE ClinicalTrialId = " + studyid
                    + " AND DataItemId = " + qid;
            dbComm.ExecuteNonQuery();
        }

        // Save a category row to the VALUEDATA table
        // using an open database command object
        public static void SaveCat(IDbCommand dbComm, int studyid, int qid,
                                    string ccode, string val, int active, int order)
        {
            // We use the Order as the ValueID
            dbComm.CommandText = "INSERT INTO VALUEDATA "
                + "(ClinicalTrialId, VersionId, DataItemId, ValueId, ValueCode, ItemValue, Active, ValueOrder) "
                + "VALUES (" + studyid + ", 1, " + qid + ", " + order
                + ", '" + ccode + "', '" + ReplaceQuotes(val) + "', " + active + ", " + order
                + ")";
            dbComm.ExecuteNonQuery();
        }

        /// <summary>
        /// Update the DataItemLength for a question using an open database command object
        /// </summary>
        /// <param name="idbcon">DB command object</param>
        /// <param name="studyid">Study ID</param>
        /// <param name="qid">Data item ID</param>
        /// <param name="length">Data item length</param>
        public static void SaveDataItemLength(IDbCommand dbComm, int studyid, int qid, int length)
        {
            dbComm.CommandText = "UPDATE DATAITEM SET DATAITEMLENGTH = " + length
                    + " WHERE ClinicalTrialId = " + studyid
                    + " AND DataItemId = " + qid;
            dbComm.ExecuteNonQuery();
        }

        /// <summary>
        /// Replace single quotes with double single quotes in string
        /// </summary>
        /// <param name="s">String</param>
        /// <returns>String with each single quote replaced with two single quotes</returns>
        public static string ReplaceQuotes(string s)
        {
            return s.Replace("'", "''");
        }

        // Get a data table containing rows of category items for ALL questions in this study
        public static DataTable GetCats(string dbcon, int studyid)
        {
            string sql = GetCatsSQL(studyid);

            DataSet ds = DataAccess.GetDataSet(dbcon, sql);

            // Did we get anything?
            if (ds == null) return null;
            if (ds.Tables[0].Rows.Count == 0) return null;
            return ds.Tables[0];
        }

        // Get a data table containing rows of category items for this study/question
        public static DataTable GetCats(string dbcon, int studyid, int qid)
        {
            string sql = GetCatsSQL(studyid, qid);

            DataSet ds = DataAccess.GetDataSet(dbcon, sql);

            // Did we get anything?
            if (ds == null) return null;
            if (ds.Tables[0].Rows.Count == 0) return null;
            return ds.Tables[0];
        }

        // Execute an SQL statement (not SELECT)
        private static void RunSQL(string dbcon, string sql)
        {
            IDbConnection idbcon = DataAccess.GetConnection(DataAccess.CalculateConnectionType(dbcon), dbcon);
            idbcon.Open();
            DataAccess.RunSQL(idbcon, sql);
            idbcon.Close();
        }

        /// <summary>
        /// Get a database connection object based on the given connection string
        /// </summary>
        /// <param name="dbcon">Connection string</param>
        /// <returns>An open database connection object</returns>
        public static IDbConnection GetConnection(string dbcon)
        {
            IDbConnection idbcon = DataAccess.GetConnection(DataAccess.CalculateConnectionType(dbcon), dbcon);
            idbcon.Open();
            return idbcon;
        }

    }
}
