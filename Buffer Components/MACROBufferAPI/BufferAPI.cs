using System;
using log4net;
using InferMed.Components;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for BufferAPI.
	/// </summary>
	class BufferAPI
	{
		private BufferAPI()
		{
			// can't instantiate
		}

		public const string _USER_SETTINGS_FILE = "settingsfile";
		public const string _MACRO_CONN_DB = "bufferdb";
		private const string _DEF_USER_SETTINGS_FILE = @"C:\Program Files\InferMed\MACRO 3.0\MACROUserSettings30.txt";
		private const string _SETTINGS_FILE = @"MACROSettings30.txt";
		
		public const int _DEFAULT_MISSING_NUMERIC = -1;

		// log4net
		private static readonly ILog log = LogManager.GetLogger( typeof(BufferAPI) );

		public enum BufferResponseStatus
		{
			Success = 0,
			FailParse = 1,
			FailClinicalTrial = 2,
			FailSite = 3,
			FailSubjectLabel = 4,
			FailDataItemCode = 5,
			FailResponseValue = 6,
			FailDbConnection = 7
		};

		public enum BufferCommitStatus
		{
			Success = 0,
			InvalidXML = 1,
			SubjectNotExist = 2,
			SubjectNotLoaded = 3,
			VisitNotExist = 4,
			EformNotExist = 5,
			QuestionNotExist = 6,
			EformInUse = 7,
			VisitLockedFrozen = 8,
			EformLockedFrozen = 9,
			QuestionNotEnterable = 10,
			NoVisitEformDate = 11,
			NoLockForSave = 12,
			ValueRejected = 13,
			NotCommitted = 14,
			LoginFailed = 15
		};

		public enum MACRODataTypes
		{
			Text = 0,
			Category = 1,
			IntegerData = 2,
			Real = 3,
			Date = 4,
			Multimedia = 5,
			LabTest = 6,
			Thesaurus = 8
		};

		/// <summary>
		/// Returns the MACRO usersettings file path from the settings file
		/// </summary>
		/// <returns>User settings file path</returns>
		public static string GetUserSettingsFilePath()
		{
			IMEDSettings20 iset = new IMEDSettings20( BufferAPI._SETTINGS_FILE );
			return( iset.GetKeyValue( BufferAPI._USER_SETTINGS_FILE, BufferAPI._DEF_USER_SETTINGS_FILE ) );
		}

		/// <summary>
		/// Returns a setting from the MACRO 3.0 user settings file
		/// </summary>
		/// <param name="key">Setting key name</param>
		/// <param name="defaultVal">Default value to be returned if the key is not found</param>
		/// <returns>MACRO setting value</returns>
		public static string GetSetting(string key, string defaultVal)
		{
			IMEDSettings20 iset = new IMEDSettings20( GetUserSettingsFilePath() );
			return( iset.GetKeyValue( key, defaultVal ) );
		}

	}
}
