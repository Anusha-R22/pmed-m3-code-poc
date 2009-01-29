using System;
using MACROAPI30;

namespace InferMed.MACRO.API
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	public class API
	{
		
		public enum LoginResult{Success = 0,AccountDisabled = 1,Failed = 2,ChangePassword = 3,PasswordExpired = 4};
		public enum DataRequestResult{Success = 0,InvalidXML = 1,SubjectNotExist = 2,SubjectNotOpened = 3};
		public enum DataInputResult{Success = 0,InvalidXML = 1,SubjectNotExist = 2,SubjectNotOpened = 3,DataNotAdded = 4};
		public enum APIRegResult{Success = 0,AlreadyRegistered = 1,NotReady = 2,Ineligible = 3,NotUnique = 4,MissingInfo = 5,UnknownError = 10,SubjectNotOpened = 11,SubjectReadOnly = 12};

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

		/// <summary>
		/// ChangeUserPassword - Implemented 21/08/2006 by DPH
		/// </summary>
		/// <param name="serialisedUser"></param>
		/// <param name="newPassword"></param>
		/// <param name="oldPassword"></param>
		/// <param name="message"></param>
		/// <returns></returns>
		public static bool ChangeUserPassword(ref string serialisedUser, string newPassword, string oldPassword, ref string message)
		{
			return (bool)(new MACROAPIClass().ChangeUserPassword(ref serialisedUser, newPassword, oldPassword, ref message));
		}

		/// <summary>
		/// RegisterSubject - Implemented 21/08/2006 by DPH
		/// </summary>
		/// <param name="serialisedUser"></param>
		/// <param name="study"></param>
		/// <param name="site"></param>
		/// <param name="subject"></param>
		/// <param name="regID"></param>
		/// <returns></returns>
		public static APIRegResult RegisterSubject(string serialisedUser, string study, string site, string subject, ref string regID)
		{
			return (APIRegResult)(new MACROAPIClass().RegisterSubject(serialisedUser, study, site, subject, ref regID));
		}

		#region not implemented from COM version
/*
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
*/
		#endregion

		#endregion
	}
}
