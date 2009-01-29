using System;

namespace InferMed.MACRO.ClinicalCoding.MACROCCBS30
{
	/// <summary>
	/// Coded term object
	/// </summary>
	public class CodedTerm
	{
		public enum eCodingStatus
		{
			Empty = 0, NotCoded = 1, Coded = 2, PendingNewCode = 3, AutoEncoded = 4, Validated = 5, DoNotCode = 6
		}

		public enum eStatus
		{
			History = 0, Current = 1
		}

		protected string _dictionaryName = "";
		protected string _dictionaryVersion = "";
		protected eCodingStatus _codingStatus = eCodingStatus.Empty;
		protected string _codingDetails = "";
		protected double _codingTimeStamp = 0;
		protected short _codingTimeStamp_TZ = 0;
		protected string _responseValue = "";
		protected double _responseTimeStamp = 0;
		protected short _responseTimeStamp_TZ = 0;
		protected string _userName = "";
		protected string _userNameFull = "";
		protected string _reasonForChange = "";


		public CodedTerm()
		{
			//use Init() as constructor because this is compiled for use in a vb6 application
			//and com does not support arguments in constructors
			//when this object is no longer being used by vb6, this constructor can be replaced 
			//and overloaded by methods below
		}

		/// <summary>
		/// Manual load of object
		/// </summary>
		/// <param name="dictionaryName"></param>
		/// <param name="dictionaryVersion"></param>
		/// <param name="codingStatus"></param>
		/// <param name="codingDetails"></param>
		/// <param name="codingTimeStamp"></param>
		/// <param name="codingTimeStamp_TZ"></param>
		/// <param name="responseTimeStamp"></param>
		/// <param name="responseTimeStamp_TZ"></param>
		/// <param name="userName"></param>
		/// <param name="userNameFull"></param>
		/// <param name="reasonForChange"></param>
		public void Init( string dictionaryName, string dictionaryVersion, eCodingStatus codingStatus, 
			string codingDetails, double  codingTimeStamp, short codingTimeStamp_TZ, string responseValue, double responseTimeStamp, 
			short responseTimeStamp_TZ, string userName, string userNameFull, string reasonForChange )
		{
			_dictionaryName = dictionaryName;
			_dictionaryVersion = dictionaryVersion;
			_codingStatus = codingStatus;
			_codingDetails = codingDetails;
			_codingTimeStamp = codingTimeStamp;
			_codingTimeStamp_TZ = codingTimeStamp_TZ;
			_responseValue = responseValue;
			_responseTimeStamp = responseTimeStamp;
			_responseTimeStamp_TZ = responseTimeStamp_TZ;
			_userName = userName;
			_userNameFull = userNameFull;
			_reasonForChange = reasonForChange;
		}

		public static eStatus GetStatus( string status )
		{
			return( ( eStatus ) System.Convert.ToInt32( status ) );
		}

		public static eCodingStatus GetCodingStatus( string codingStatus )
		{
			return( ( codingStatus == "" ) ? eCodingStatus.Empty : ( eCodingStatus ) System.Convert.ToInt32( codingStatus ) );
		}

		public string DictionaryName
		{
			get { return( _dictionaryName ); }
		}

		public string DictionaryVersion
		{
			get { return( _dictionaryVersion ); }
		}

		public eCodingStatus CodingStatus
		{
			get { return( _codingStatus ); }
		}

		public string CodingDetails
		{
			get { return( _codingDetails ); }
		}

		public double CodingTimeStamp
		{
			get { return( _codingTimeStamp ); }
		}

		public short CodingTimeStamp_TZ
		{
			get { return( _codingTimeStamp_TZ ); }
		}

		public string ResponseValue
		{
			get { return( _responseValue ); }
		}

		public double ResponseTimeStamp
		{
			get { return( _responseTimeStamp ); }
		}

		public short ResponseTimeStamp_TZ
		{
			get { return( _responseTimeStamp_TZ ); }
		}

		public string UserName
		{
			get { return( _userName ); }
		}

		public string UserNameFull
		{
			get { return( _userNameFull ); }
		}

		public string ReasonForChange
		{
			get { return( _reasonForChange ); }
		}
	}
}
