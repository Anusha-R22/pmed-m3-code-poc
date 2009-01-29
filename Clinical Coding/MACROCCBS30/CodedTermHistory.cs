using System;
//using System.Runtime.InteropServices;
using System.Data;
using System.Collections;
using InferMed.Components;

namespace InferMed.MACRO.ClinicalCoding.MACROCCBS30
{
//	[ComVisible(true)]
//	[Guid("B8E160BA-F8E7-48D5-98E9-CE682161A4E4")]
//	public interface ICodedTermHistory
//	{
//		[ComVisible(true)]
//			void InitEmpty( int clinicalTrialId, string trialSite, int personId, int responseTaskId, 
//			short repeat );
//		[ComVisible(true)]
//			void InitManual( int clinicalTrialId, string trialSite, int personId, int responseTaskId, 
//			short repeat, string dictionaryName, string dictionaryVersion, CodedTerm.eCodingStatus codingStatus, 
//			string codingDetails, double codingTimeStamp, short codingTimeStamp_TZ, string responseValue, double responseTimeStamp, 
//			short responseTimeStamp_TZ, string userName, string userNameFull);
//		[ComVisible(true)]
//			void InitAuto( string con, int clinicalTrialId, string trialSite, int personId, int responseTaskId, 
//			short repeat );
//		[ComVisible(true)]
//			void Save( string con, int visitId, short visitCycle, int crfPageId, short crfPageCycle );
//		[ComVisible(true)]
//			void SetCode( string dictionaryName, string dictionaryVersion, string codingDetails, string userName,
//			string userNameFull, string responseValue, double responseTimeStamp, short responseTimeStamp_TZ, string reasonForChange,
//			bool autoEncode);
//		[ComVisible(true)]
//			void SetStatus( CodedTerm.eCodingStatus codingStatus, string userName, string userNameFull, 
//			string responseValue, double responseTimeStamp, short responseTimeStamp_TZ );
//		[ComVisible(true)]
//			void AddHistoryTerm( ref CodedTerm c );
//	}

	/// <summary>
	/// Object holding clinical coding history of a response
	/// </summary>
//	[ComVisible(true)]
//	[Guid("3AF497FB-EC2F-40DA-87B3-C2BCCF41837A")]
//	[ClassInterface(ClassInterfaceType.None)]
	public class CodedTermHistory : CodedTerm //, ICodedTermHistory
	{
		private bool _edited = false;

		private int _clinicalTrialId;
		private string _trialSite;
		private int _personId;
		private int _responseTaskId;
		private short _repeat;
		protected ArrayList _history = new ArrayList();

		public CodedTermHistory()
		{
			//use Init() as constructor because this is compiled for use in a vb6 application
			//and com does not support arguments in constructors
			//when this object is no longer being used by vb6, this constructor can be replaced 
			//and overloaded by methods below
		}

		/// <summary>
		/// Initialise an empty object
		/// </summary>
		/// <param name="clinicalTrialId"></param>
		/// <param name="trialSite"></param>
		/// <param name="personId"></param>
		/// <param name="responseTaskId"></param>
		/// <param name="repeat"></param>
//		[ComVisible(true)]
		public void InitEmpty( int clinicalTrialId, string trialSite, int personId, int responseTaskId, 
			short repeat )
		{
			_clinicalTrialId = clinicalTrialId;
			_trialSite = trialSite;
			_personId = personId;
			_responseTaskId = responseTaskId;
			_repeat = repeat;
		}

		/// <summary>
		/// Manual load of object
		/// </summary>
		/// <param name="clinicalTrialId"></param>
		/// <param name="trialSite"></param>
		/// <param name="personId"></param>
		/// <param name="responseTaskId"></param>
		/// <param name="repeat"></param>
		/// <param name="dictionaryName"></param>
		/// <param name="dictionaryVersion"></param>
		/// <param name="codingStatus"></param>
		/// <param name="codingDetails"></param>
		/// <param name="codingTimeStamp"></param>
		/// <param name="responseTimeStamp"></param>
		/// <param name="responseTimeStamp_TZ"></param>
		/// <param name="userName"></param>
		/// <param name="userNameFull"></param>
		/// <param name="codingTimeStamp_TZ"></param>
		/// <param name="reasonForChange"></param>
//		[ComVisible(true)]
		public void InitManual( int clinicalTrialId, string trialSite, int personId, int responseTaskId, 
			short repeat, string dictionaryName, string dictionaryVersion, CodedTerm.eCodingStatus codingStatus, 
			string codingDetails, double codingTimeStamp, short codingTimeStamp_TZ, string responseValue, double responseTimeStamp, 
			short responseTimeStamp_TZ, string userName, string userNameFull)
		{
			_clinicalTrialId = clinicalTrialId;
			_trialSite = trialSite;
			_personId = personId;
			_responseTaskId = responseTaskId;
			_repeat = repeat;
			base.Init( dictionaryName, dictionaryVersion, codingStatus, codingDetails, codingTimeStamp, 
				codingTimeStamp_TZ, responseValue, responseTimeStamp, responseTimeStamp_TZ, userName, userNameFull, "" );
		}

		/// <summary>
		/// Automatic load of object
		/// </summary>
		/// <param name="con"></param>
		/// <param name="clinicalTrialId"></param>
		/// <param name="trialSite"></param>
		/// <param name="personId"></param>
		/// <param name="responseTaskId"></param>
		/// <param name="repeat"></param>
//		[ComVisible(true)]
		public void InitAuto( string con, int clinicalTrialId, string trialSite, int personId, int responseTaskId, 
			short repeat )
		{
			DataSet ds = null;
			
			try
			{
				_clinicalTrialId = clinicalTrialId;
				_trialSite = trialSite;
				_personId = personId;
				_responseTaskId = responseTaskId;
				_repeat = repeat;

				//create sql to get coding history for a response
				string sql = "SELECT * FROM CODINGHISTORY "
					+ "WHERE CLINICALTRIALID = " + clinicalTrialId + " "
					+ "AND TRIALSITE = '" + trialSite + "' " 
					+ "AND PERSONID = " + personId + " "
					+ "AND RESPONSETASKID = " + responseTaskId + " " 
					+ "AND REPEATNUMBER = " + repeat;

				//get the coding history, if any
				ds = CCDataAccess.GetDataSet( con, sql );

				foreach( DataRow r in ds.Tables[0].Rows )
				{
					if( GetStatus( CCDataAccess.ConvertFromNull( r, "Status" ) ) == eStatus.Current )
					{
						//load current codingdetails
						base.Init( CCDataAccess.ConvertFromNull( r, "DictionaryName" ),	
							CCDataAccess.ConvertFromNull( r, "DictionaryVersion" ), 
							GetCodingStatus( CCDataAccess.ConvertFromNull( r, "CodingStatus" ) ),
							CCDataAccess.ConvertFromNull( r, "CodingDetails" ), 
							System.Convert.ToDouble( CCDataAccess.ConvertFromNull( r, "CodingTimeStamp" ) ),	
							System.Convert.ToInt16( CCDataAccess.ConvertFromNull( r, "CodingTimeStamp_TZ" ) ),
							CCDataAccess.ConvertFromNull( r, "ResponseValue" ),
							System.Convert.ToDouble( CCDataAccess.ConvertFromNull( r, "ResponseTimeStamp" ) ), 
							System.Convert.ToInt16( CCDataAccess.ConvertFromNull( r, "ResponseTimeStamp_TZ" ) ), 
							CCDataAccess.ConvertFromNull( r, "UserName" ),
							CCDataAccess.ConvertFromNull( r, "UserNameFull" ), 
							"" );
					}

					//save the coding details in the history
					CodedTerm t = new CodedTerm();

					t.Init( CCDataAccess.ConvertFromNull( r, "DictionaryName" ),	
					CCDataAccess.ConvertFromNull( r, "DictionaryVersion" ), 
					GetCodingStatus( CCDataAccess.ConvertFromNull( r, "CodingStatus" ) ),
					CCDataAccess.ConvertFromNull( r, "CodingDetails" ), 
					System.Convert.ToDouble( CCDataAccess.ConvertFromNull( r, "CodingTimeStamp" ) ),	
					System.Convert.ToInt16( CCDataAccess.ConvertFromNull( r, "CodingTimeStamp_TZ" ) ),
					CCDataAccess.ConvertFromNull( r, "ResponseValue" ),
					System.Convert.ToDouble( CCDataAccess.ConvertFromNull( r, "ResponseTimeStamp" ) ), 
					System.Convert.ToInt16( CCDataAccess.ConvertFromNull( r, "ResponseTimeStamp_TZ" ) ), 
					CCDataAccess.ConvertFromNull( r, "UserName" ),
					CCDataAccess.ConvertFromNull( r, "UserNameFull" ), 
					CCDataAccess.ConvertFromNull( r, "ReasonForChange" ) );
					
					AddHistoryTerm( ref t );
				}
			}
			finally
			{
				ds.Dispose();
			}

		}

		/// <summary>
		/// Save the coded term if it has changed
		/// </summary>
		/// <param name="con"></param>
		/// <param name="visitId"></param>
		/// <param name="visitCycle"></param>
		/// <param name="crfPageId"></param>
		/// <param name="crfPageCycle"></param>
//		[ComVisible(true)]
		public void Save( string con, int visitId, short visitCycle, int crfPageId, short crfPageCycle )
		{
			if( IsEdited )
			{
				CCDataAccess.RunSQL( con, CCDataAccess.Sproc_SP_MACRO_CODING_UPDATE( IMEDDataAccess.CalculateConnectionType( con ), this.ClinicalTrialId , 
					this.TrialSite, this.PersonId, visitId , visitCycle , crfPageId, crfPageCycle , this.ResponseTaskId,
					this.Repeat, this.DictionaryName, this.DictionaryVersion, this.CodingStatus, this.CodingDetails, this.CodingTimeStamp,
					this.CodingTimeStamp_TZ, this.ResponseValue, this.ResponseTimeStamp, this.ResponseTimeStamp_TZ, this.UserName,
					this.UserNameFull, this.ReasonForChange ) );

				CodedTerm t = new CodedTerm();
				t.Init( this.DictionaryName, this.DictionaryVersion, this.CodingStatus, this.CodingDetails, 
					this.CodingTimeStamp,	this.CodingTimeStamp_TZ, this.ResponseValue, this.ResponseTimeStamp, this.ResponseTimeStamp_TZ, 
					this.UserName, this.UserNameFull, this.ReasonForChange );
				
				AddHistoryTerm( ref t );

				MarkAsUnchanged();
			}
		}

		/// <summary>
		/// Get now
		/// </summary>
		/// <returns></returns>
		private double ImedNow()
		{
			return( DateTime.Now.ToOADate() );
		}

		/// <summary>
		/// Get timezone
		/// </summary>
		/// <returns></returns>
		private short GMTDiff()
		{
			// get local timezone
			TimeZone tzLocal = TimeZone.CurrentTimeZone;
			// get timespan of UTC offset
			TimeSpan tsLocal = tzLocal.GetUtcOffset(DateTime.Now);
			// calculate offset MACRO style - offset in minutes UTC (GMT)
			short nOffset = Convert.ToInt16( -1 * Convert.ToInt32(tsLocal.TotalMinutes) );
			return nOffset;
		}

		public bool IsEdited
		{
			get { return( _edited ); }
		}

		/// <summary>
		/// Does the term require a reason for change
		/// </summary>
		public bool RequiresRFC
		{
			get 
			{ 
				if( NoHistory )
				{
					return( false );
				}
				else
				{
					if( _history.Count == 1 )
					{
						CodedTerm t = ( CodedTerm )_history[0];
						return( ( ( t.CodingStatus != eCodingStatus.DoNotCode ) 
							&& ( t.CodingStatus != eCodingStatus.NotCoded  ) ) );
					}
					else
					{
						return( true );
					}
				}
			}
		}

		/// <summary>
		/// Does the term have a history
		/// </summary>
		/// <returns></returns>
		public bool NoHistory
		{
			get
			{
				if( _history.Count == 0 )
				{
					return( true );
				}
				else
				{
					return( false );
				}
			}
		}

		/// <summary>
		/// Set the code
		/// </summary>
		/// <param name="dictionaryName"></param>
		/// <param name="dictionaryVersion"></param>
		/// <param name="codingDetails"></param>
		/// <param name="userName"></param>
		/// <param name="userNameFull"></param>
		/// <param name="responseValue"></param>
		/// <param name="responseTimeStamp"></param>
		/// <param name="responseTimeStamp_TZ"></param>
		/// <param name="reasonForChange"></param>
		/// <param name="autoEncode"></param>
//		[ComVisible(true)]
		public void SetCode( string dictionaryName, string dictionaryVersion, string codingDetails, string userName,
			string userNameFull, string responseValue, double responseTimeStamp, short responseTimeStamp_TZ, string reasonForChange,
			bool autoEncode)
		{
			_dictionaryName = dictionaryName;
			_dictionaryVersion = dictionaryVersion;
			_codingDetails = codingDetails;
			_codingStatus = ( autoEncode ) ? eCodingStatus.AutoEncoded : eCodingStatus.Coded;
			_userName = userName;
			_userNameFull = userNameFull;
			_responseValue = responseValue;
			_responseTimeStamp = responseTimeStamp;
			_responseTimeStamp_TZ = responseTimeStamp_TZ;
			_reasonForChange = reasonForChange;
			MarkAsChanged();
		}

		/// <summary>
		/// Set the status
		/// </summary>
		/// <param name="codingStatus"></param>
		/// <param name="userName"></param>
		/// <param name="userNameFull"></param>
		/// <param name="responseValue"></param>
		/// <param name="responseTimeStamp"></param>
		/// <param name="responseTimeStamp_TZ"></param>
//		[ComVisible(true)]
		public void SetStatus( CodedTerm.eCodingStatus codingStatus, string userName, string userNameFull, 
			string responseValue, double responseTimeStamp, short responseTimeStamp_TZ )
		{
			if(_codingStatus == codingStatus)
			{
				//if we are trying to set the coded term status to the status that it is already, dont bother
			}
			else if( ( NoHistory ) && ( responseValue == "" ) && ( codingStatus == CodedTerm.eCodingStatus.NotCoded ) )
			{
				//if the term has no history, the responsevalue is "" and the status has been set to not coded,
				//the unsaved response has just had its value cleared - set the term back to empty
				_codingStatus = CodedTerm.eCodingStatus.Empty;
				_userName = "";
				_userNameFull = "";
				_responseValue = "";
				_responseTimeStamp = 0;
				_responseTimeStamp_TZ = 0;
				MarkAsUnchanged();
			}
			else
			{
				_codingStatus = codingStatus;
				_userName = userName;
				_userNameFull = userNameFull;
				_responseValue = responseValue;
				_responseTimeStamp = responseTimeStamp;
				_responseTimeStamp_TZ = responseTimeStamp_TZ;
				//if status is changing to not coded but the term has a code, clear it automatically
				if( codingStatus == eCodingStatus.NotCoded )
				{
					if( _codingDetails != "" )
					{
						_dictionaryName = "";
						_dictionaryVersion = "";
						_codingDetails = "";
						_reasonForChange = "*** Coding details cleared automatically due to status change to not coded";
					}
				}
			
				MarkAsChanged();
			}
		}

		/// <summary>
		/// Mark the term as changed and update timestamp
		/// </summary>
		private void MarkAsChanged()
		{
			_edited = true;
			_codingTimeStamp = ImedNow();
			_codingTimeStamp_TZ = GMTDiff();
		}

		/// <summary>
		/// Mark term as not changed
		/// </summary>
		private void MarkAsUnchanged()
		{
			_reasonForChange = "";
			_edited = false;
		}

		/// <summary>
		/// Add a history term
		/// </summary>
		/// <param name="c"></param>
//		[ComVisible(true)]
		public void AddHistoryTerm( ref CodedTerm c )
		{
			_history.Add( c );
		}
		
		public int ClinicalTrialId
		{
			get { return( _clinicalTrialId ); }
		}

		public string TrialSite
		{
			get { return( _trialSite ); }
		}

		public int PersonId
		{
			get { return( _personId ); }
		}

		public int ResponseTaskId
		{
			get { return( _responseTaskId ); }
		}

		public int Repeat
		{
			get { return( _repeat ); }
		}

		public ArrayList History
		{
			get{ return _history; }
		}
	}
}
