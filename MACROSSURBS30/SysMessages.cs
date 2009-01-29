/* ------------------------------------------------------------------------
 * File: SysMessages.cs
 * Author: Nicky Johns
 * Copyright: InferMed, 2008, All Rights Reserved
 * Purpose: Message handling for Study/Site/User/Role associations for MACRO 3.0 API
 * Date: March 2008
 * ------------------------------------------------------------------------
 * Revisions:
 * NCJ 17 Mar 08 - Must set CursorLocation for ADODB connection
 * ------------------------------------------------------------------------*/
using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using MACROSysDataXfer30;

namespace MACROSSURBS30
{
    /// <summary>
    /// Implements writing to the MESSAGE table in MACRO for User Role Associations
    /// </summary>
    public static class SysMessages
    {
        // Constants copied from MACRO 3.0
        private const short MSG_USER_MESSAGE = 32;
        private const short MSG_USERROLE_MESSAGE = 33;

        private const string MSG_SEP = "*";
        private const string MSG_ROLE_ADD = "1";
        private const string MSG_ROLE_ADD_TXT = "New User Role";
        private const string MSG_ROLE_REMOVE = "0";
        private const string MSG_ROLE_REMOVE_TXT = "Delete User Role";
        private const string MSG_INSTALL_TYPE = "1";
        private const string MSG_USER_TXT = "User Details";
        private const string MSG_USER_EDIT = "1";

        /// <summary>
        /// Store message in MACRO MESSAGE table for New User Role
        /// </summary>
        /// <param name="dbCon">MACRO database connection string</param>
        /// <param name="secDBCon">MACRO database connection string</param>
        /// <param name="apiUser">User name of current API user</param>
        /// <param name="assoc">SSUR association object</param>
        public static void SendNewUserRoleMessage(string dbCon, string secDBCon, string apiUser, SSURAssoc assoc)
        {
            // Get parameters for adding a role
            string msgParams = RoleParams(assoc, true);
            short msgType = MSG_USERROLE_MESSAGE;
            string msgBody = MSG_ROLE_ADD_TXT;
            string sysUser = apiUser;
            string rUser = assoc.User;
            string site = assoc.Site;
            // NB This is not the SSUR role code! (Which is included in the msgParams)
            string roleCode = "";

            // Hook up to MACROSysDataXfer30
            MACROSysDataXfer30.SysMessages messager = new MACROSysDataXfer30.SysMessages();

            // Create ADODB database connection
            // Pass empty strings as extra params to macroCon.Open
            // http://www.dotnetcoders.com/web/Articles/ShowArticle.aspx?article=54
            ADODB.Connection macroCon = new ADODB.Connection();
            macroCon.Open(dbCon, "", "", -1);
            // Must set CursorLocation (otherwise Oracle doesn't work)
            macroCon.CursorLocation = ADODB.CursorLocationEnum.adUseClient;

            messager.AddNewSystemMessage(ref macroCon, ref msgType, ref sysUser, ref rUser,
                            ref msgBody, ref msgParams, ref site, ref roleCode);

            // Get parameters for User Details
            msgParams = UserParams(secDBCon, assoc.User);
            msgType = MSG_USER_MESSAGE;
            msgBody = MSG_USER_TXT;
            // For a User Details message, pass "" for "AllSites"
            site = (assoc.Site == SSURSQL.ALL_SITES ? "" : assoc.Site);
            messager.AddNewSystemMessage(ref macroCon, ref msgType, ref sysUser, ref rUser,
                            ref msgBody, ref msgParams, ref site, ref roleCode);

            // Close ADODB Connection
            macroCon.Close();
            macroCon = null;
        }

        /// <summary>
        /// Store message in MACRO MESSAGE table for Delete User Role
        /// </summary>
        /// <param name="dbCon">MACRO database connection string</param>
        /// <param name="apiUser">User name of current API user</param>
        /// <param name="assoc">SSUR association object</param>
        public static void SendRemoveRoleMessage(string dbCon, string apiUser, SSURAssoc assoc)
        {
            // Get parameters for deleting a role
            string msgParams = RoleParams(assoc, false);
            short msgType = MSG_USERROLE_MESSAGE;
            string msgBody = MSG_ROLE_REMOVE_TXT;
            string sysUser = apiUser;
            string rUser = assoc.User;
            string site = assoc.Site;
            string roleCode = "";

            // Hook up to MACROSysDataXfer30
            MACROSysDataXfer30.SysMessages messager = new MACROSysDataXfer30.SysMessages();

            // Create ADODB database connection
            ADODB.Connection macroCon = new ADODB.Connection();
            macroCon.Open(dbCon, "", "", -1);
            // Must set CursorLocation (otherwise Oracle doesn't work)
            macroCon.CursorLocation = ADODB.CursorLocationEnum.adUseClient;

            messager.AddNewSystemMessage(ref macroCon, ref msgType, ref sysUser, ref rUser,
                            ref msgBody, ref msgParams, ref site, ref roleCode);

            // Close ADODB Connection
            macroCon.Close();
            macroCon = null;
        }

        /// <summary>
        /// Get the message parameters for a Role message (add or delete)
        /// </summary>
        /// <param name="assoc">The SSUR association</param>
        /// <param name="add">True for adding role; False for removing role</param>
        /// <returns>Parameter string</returns>
        private static string RoleParams(SSURAssoc assoc, Boolean add)
        {
            return assoc.User + MSG_SEP
                + assoc.Role + MSG_SEP
                + assoc.Study + MSG_SEP
                + assoc.Site + MSG_SEP
                + MSG_INSTALL_TYPE + MSG_SEP
                + (add ? MSG_ROLE_ADD : MSG_ROLE_REMOVE);
        }

        /// <summary>
        /// Get the message parameters for a "User Details" message
        /// </summary>
        /// <param name="secDBCon">Security database connection string</param>
        /// <param name="user">User name</param>
        /// <returns>Parameter string</returns>
        private static string UserParams(string secDBCon, string user)
        {
            // Retrieve this user's info from the MACROUSER table
            DataTable dt = SSURSQL.GetUserDetails(secDBCon, user);
            if (dt == null)
                return "";      // Shouldn't happen!
            // Assume just one row
            DataRow dr = dt.Rows[0];

            return dr["USERNAME"].ToString() + MSG_SEP
                + dr["USERNAMEFULL"].ToString() + MSG_SEP
                + dr["USERPASSWORD"].ToString() + MSG_SEP
                + dr["ENABLED"].ToString() + MSG_SEP
                + dr["LASTLOGIN"].ToString() + MSG_SEP
                + dr["FIRSTLOGIN"].ToString() + MSG_SEP
                + dr["FAILEDATTEMPTS"].ToString() + MSG_SEP
                + dr["PASSWORDCREATED"].ToString() + MSG_SEP
                + dr["SYSADMIN"].ToString() + MSG_SEP
                + MSG_USER_EDIT;
        }

    }
}