/* ------------------------------------------------------------------------
 * File: API.cs
 * Author: David Hook
 * Copyright: InferMed, 2007-8, All Rights Reserved
 * Purpose: Interface to MACRO 3.0 API
 * Date: Dec 2007
 * ------------------------------------------------------------------------
 * Revisions:
 * NCJ 19 Mar 08 - Added new functionality for Patch 3.0.82 (ARFG 3 & 4)
 * ------------------------------------------------------------------------*/

using System;
using MACROAPI30;

namespace InferMed.MACRO.API
{
	/// <summary>
	/// The API class for MACRO 3.0
	/// </summary>
	public class API
	{
		// Enumerations
		public enum LoginResult
        {Success = 0,AccountDisabled = 1,Failed = 2,ChangePassword = 3,PasswordExpired = 4,InvalidSecurityDb = 5};
		
        public enum DataRequestResult
        {Success = 0,InvalidXML = 1,SubjectNotExist = 2,SubjectNotOpened = 3};
		
        public enum DataInputResult
        {Success = 0,InvalidXML = 1,SubjectNotExist = 2,SubjectNotOpened = 3,DataNotAdded = 4};
		
        public enum APIRegResult
        {Success = 0,AlreadyRegistered = 1,NotReady = 2,Ineligible = 3,NotUnique = 4,MissingInfo = 5,UnknownError = 10,SubjectNotOpened = 11,SubjectReadOnly = 12};
        
        public enum ImportResult
        { Error = -1, Success = 0, InvalidXML = 1, NotAllDone = 2, PermissionDenied = 3 };

        public enum PasswordResult
        {
            Error = -1,
            Success = 0,
            AccountDisabled = 1,
            UnknownUser = 2,
            InvalidPassword = 3,
            PermissionDenied = 4
        };

		private API(){/*prevent instances of class*/}

		#region COM _MACROAPI Members

		/// <summary>
		/// allow external application to log in
		/// </summary>
		/// <param name="userName"></param>
		/// <param name="password"></param>
		/// <param name="databaseCode"></param>
		/// <param name="userRole"></param>
		/// <param name="message">returns by ref any failure message</param>
		/// <param name="userNameFull"></param>
		/// <param name="serialisedUser">returns by ref a serialised version of the user for use in subsequent API calls</param>
		/// <returns>login success status</returns>
		public static LoginResult Login(string userName, string password, string databaseCode, string userRole, ref string message, ref string userNameFull, ref string serialisedUser)
		{
			return  (LoginResult)(new MACROAPIClass().Login(userName,password,databaseCode,userRole, ref message, ref userNameFull, ref serialisedUser));
		}

		/// <summary>
		/// allow external application to create subject
		/// </summary>
		/// <param name="serialisedUser"></param>
		/// <param name="studyId"></param>
		/// <param name="site"></param>
		/// <param name="message">returns be ref any failure message</param>
		/// <returns>the personid or -1 if failed</returns>
		public static int CreateSubject(string serialisedUser, int studyId, string site, ref string message)
		{
			return (int)(new MACROAPIClass().CreateSubject(ref serialisedUser, studyId, site, ref message));
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="serialisedUser"></param>
		/// <param name="dataInputXml"></param>
		/// <param name="reportXml"></param>
		/// <returns></returns>
		public static DataInputResult InputXMLSubjectData(string serialisedUser, string dataInputXml, ref string reportXml)
		{
			return (DataInputResult)(new MACROAPIClass().InputXMLSubjectData(serialisedUser,dataInputXml,ref reportXml));
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="serialisedUser"></param>
		/// <param name="dataRequestXml"></param>
		/// <param name="returnedDataXml"></param>
		/// <returns></returns>
		public static DataRequestResult GetXMLSubjectData(string serialisedUser, string dataRequestXml, ref string returnedDataXml)
		{
			return (DataRequestResult)(new MACROAPIClass().GetXMLSubjectData(serialisedUser,dataRequestXml,ref returnedDataXml));
		}

		/// <summary>
		/// Change password of currently logged-in user
		/// </summary>
        /// <param name="serialisedUser">User details as returned from Login or LoginSecurity</param>
		/// <param name="newPassword">New password</param>
		/// <param name="oldPassword">Old password</param>
		/// <param name="message">Details of error if call fails</param>
		/// <returns>Whether password change was successful</returns>
		public static bool ChangeUserPassword(ref string serialisedUser, string newPassword, string oldPassword, ref string message)
		{
			return (bool)(new MACROAPIClass().ChangeUserPassword(ref serialisedUser, newPassword, oldPassword, ref message));
		}

		/// <summary>
		/// Register subject using registration as defined in study definition
		/// </summary>
        /// <param name="serialisedUser">User details as returned from Login or LoginSecurity</param>
		/// <param name="study">Study name</param>
		/// <param name="site">Site code</param>
		/// <param name="subject">Subject, either as label or subject ID</param>
		/// <param name="regID">Returned registration ID</param>
		/// <returns>Result code indicating outcome of registration</returns>
		public static APIRegResult RegisterSubject(string serialisedUser, string study, string site, string subject, ref string regID)
		{
			return (APIRegResult)(new MACROAPIClass().RegisterSubject(serialisedUser, study, site, subject, ref regID));
		}

		/// <summary>
		/// Allows login to a selected security database
		/// </summary>
		/// <param name="userName">User name as stored in MACRO database</param>
		/// <param name="password">Password</param>
		/// <param name="databaseCode">MACRO database code</param>
		/// <param name="userRole">Role code</param>
		/// <param name="securityCon">Connection string for MACRO security database</param>
		/// <param name="message">Returns by ref any failure message</param>
        /// <param name="userNameFull">If successful, returns by ref full name of user</param>
		/// <param name="serialisedUser">Returns by ref a serialised version of the user for use in subsequent API calls</param>
		/// <returns>Result code indication outcome of Login</returns>
		public static LoginResult LoginSecurity(string userName, string password, 
			string databaseCode, string userRole, string securityCon,
			ref string message, 
			ref string userNameFull, ref string serialisedUser)
		{
			return (LoginResult)(new MACROAPIClass().LoginSecurity(userName,password,databaseCode,userRole,securityCon, ref message, ref userNameFull, ref serialisedUser));
		}

        /// <summary>
        /// Reset another user's password
        /// </summary>
        /// <param name="serialisedUser">Details of currently logged in user as returned from Login or LoginSecurity</param>
        /// <param name="userName">Name of user for whom password is to be reset</param>
        /// <param name="newPassword">New password</param>
        /// <param name="message">Returned message if call fails</param>
        /// <returns>Result code indicating outcome</returns>
        public static PasswordResult ResetPassword(ref string serialisedUser, string userName, 
                            string newPassword, ref string message)
        {
            return (PasswordResult)(new MACROAPIClass().ResetPassword(ref serialisedUser, userName, newPassword, ref message));
        }

		/// <summary>
		/// Export category question details as XML
		/// </summary>
        /// <param name="serialisedUser">User details as returned from Login or LoginSecurity</param>
		/// <param name="xmlCatRequest">xml category request</param>
		/// <param name="xmlReport">returned xml report</param>
        /// <returns>True if export succeeded, or False if export could not be done</returns>
		public static bool ExportCategories(string serialisedUser, string xmlCatRequest, ref string xmlReport)
		{
			return (bool)(new MACROAPIClass().ExportCategories(serialisedUser, xmlCatRequest, ref xmlReport));
		}

		/// <summary>
        /// Import categories specified in XML string
		/// </summary>
        /// <param name="serialisedUser">User details as returned from Login or LoginSecurity</param>
        /// <param name="xmlCatsInput">XML string containing category specifications</param>
        /// <param name="xmlReport">Error messages as XML</param>
        /// <returns>Result code indicating outcome</returns>
        public static ImportResult ImportCategories(string serialisedUser, string xmlCatsInput, ref string xmlReport)
		{
            return (ImportResult)(new MACROAPIClass().ImportCategories(serialisedUser, xmlCatsInput, ref xmlReport));
		}

        /// <summary>
        /// Import User Role associations specified in XML string
        /// </summary>
        /// <param name="serialisedUser">User details as returned from Login or LoginSecurity</param>
        /// <param name="associationsXml">XML string containing user role associations</param>
        /// <param name="reportXml">Error messages as XML</param>
        /// <returns>Result code indicating outcome</returns>
        public static ImportResult ImportAssociations(string serialisedUser, string associationsXml, ref string reportXml)
        {
            return (ImportResult)(new MACROAPIClass().ImportAssociations(serialisedUser, associationsXml, ref reportXml));
        }

        /// <summary>
        /// Export user role associations as an XML string
        /// </summary>
        /// <param name="serialisedUser">User details as returned from Login or LoginSecurity</param>
        /// <param name="assocRequestXml">XML string containing details of user roles to be exported</param>
        /// <param name="reportXml">XML string containing user role associations</param>
        /// <returns>True if export succeeded, or False if export could not be done</returns>
        public static bool ExportAssociations(string serialisedUser, string assocRequestXml, ref string reportXml)
        {
            return (new MACROAPIClass().ExportAssociations(serialisedUser, assocRequestXml, ref reportXml));
        }

		#region not implemented from COM version
        /*
		/// <summary>
		/// ChangeUserDetails - Implemented 21/08/2006 by DPH
		/// </summary>
		/// <param name="serialisedUser"></param>
		/// <param name="newDetails"></param>
		/// <param name="message"></param>
		/// <returns></returns>
		//public static bool ChangeUserDetails(string serialisedUser, ref UserDetail newDetails, ref object message)
		//{
		//	return (bool)(new MACROAPIClass().ChangeUserDetails(serialisedUser, ref newDetails, ref message)); 
		//}

		/// <summary>
		/// GetUsersDetails - Implemented 21/08/2006 by DPH
		/// </summary>
		/// <param name="serialisedUser"></param>
		/// <param name="userName"></param>
		/// <param name="message"></param>
		/// <returns></returns>
		//public static VBA.Collection GetUsersDetails(string serialisedUser, string userName, ref object message)
		//{
		//	return (VBA.Collection)(new MACROAPIClass().GetUsersDetails(serialisedUser, userName, ref message));
		//}

		public int LoginForASP(string sUserName, string sPassword, string sDatabaseCode, string sUserRole, ref object vMessage, ref object vUserNameFull, ref object vSerialisedUser)
		{
			// TODO:  Add API.LoginForASP implementation
			return 0;
		}

		public int ChangeUserPasswordForASP(ref object vSerialisedUser, string sNewPassword, string sOldPassword, ref object vMessage)
		{
			// TODO:  Add API.ChangeUserPasswordForASP implementation
			return 0;
		}
		
		public int CreateSubjectForASP(ref object vSerialisedUser , int nStudyId, string sSite, ref object vMessage)
		{
			// TODO:  Add API.CreateSubjectForASP implementation
			return 0;		
		}
		
		public int RegisterSubjectForASP(object sSerialisedUser, string sStudyName, string sSite, string sSubject, ref object vRegID)
		{
			// TODO:  Add API.RegisterSubjectForASP implementation
			return 0;		
		}
		
		public static int LoginSecurityForASP(string userName, string password, 
			string databaseCode, string userRole, string securityCon,
			ref object message, 
			ref object userNameFull, ref object serialisedUser)
		{
			return 0;
		}

		public static bool ExportCategoriesForASP(string serialisedUser, string xmlCatRequest,
			ref object xmlReport)
		{
			return 0;
		}
		
		public static int ImportCategoriesForASP(string serialisedUser, string xmlCatsInput,
			ref object xmlReport)
		{
			return 0;
		}
*/
        #endregion

        #endregion
    }
}
