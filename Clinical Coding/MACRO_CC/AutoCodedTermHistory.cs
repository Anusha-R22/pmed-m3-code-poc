using System;
using InferMed.MACRO.ClinicalCoding.MACROCCBS30;

namespace InferMed.MACRO.ClinicalCoding.MACRO_CC
{
	/// <summary>
	/// Encapsulates an auto coded term
	/// </summary>
	public class AutoCodedTermHistory : CodedTermHistory
	{
		//response details
		private int _visitId;
		private short _visitCycle;
		private int _crfPageId;
		private short _crfPageCycle;
		//total number of matches found in dictionary
		private int _matches = 0;
		//dictionary used
		private Dictionary _ccDictionary = null;

		public AutoCodedTermHistory()
		{
		}

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="clinicalTrialId"></param>
		/// <param name="trialSite"></param>
		/// <param name="personId"></param>
		/// <param name="visitId"></param>
		/// <param name="visitCycle"></param>
		/// <param name="crfPageId"></param>
		/// <param name="crfPageCycle"></param>
		/// <param name="responseTaskId"></param>
		/// <param name="repeat"></param>
		/// <param name="responseValue"></param>
		/// <param name="responseTimeStamp"></param>
		/// <param name="responseTimeStamp_TZ"></param>
		/// <param name="ccDictionary"></param>
		public void InitEmpty( int clinicalTrialId, string trialSite, int personId, int visitId, short visitCycle, 
			int crfPageId, short crfPageCycle, int responseTaskId, short repeat, string responseValue, double responseTimeStamp,
			short responseTimeStamp_TZ, Dictionary ccDictionary )
		{
			_visitId = visitId;
			_visitCycle = visitCycle;
			_crfPageId = crfPageId;
			_crfPageCycle = crfPageCycle;
			_ccDictionary = ccDictionary;
			_responseValue = responseValue;
			_responseTimeStamp = responseTimeStamp;
			_responseTimeStamp_TZ = responseTimeStamp_TZ;
			base.InitEmpty( clinicalTrialId, trialSite, personId, responseTaskId, repeat );
		}

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="con"></param>
		/// <param name="clinicalTrialId"></param>
		/// <param name="trialSite"></param>
		/// <param name="personId"></param>		
		/// <param name="visitId"></param>
		/// <param name="visitCycle"></param>
		/// <param name="crfPageId"></param>
		/// <param name="crfPageCycle"></param>
		/// <param name="responseTaskId"></param>
		/// <param name="repeat"></param>
		/// <param name="ccDictionary"></param>
		public void InitAuto(  string con, int clinicalTrialId, string trialSite, int personId, int visitId, short visitCycle, 
			int crfPageId, short crfPageCycle, int responseTaskId, short repeat, Dictionary ccDictionary )
		{
			_visitId = visitId;
			_visitCycle = visitCycle;
			_crfPageId = crfPageId;
			_crfPageCycle = crfPageCycle;
			_ccDictionary = ccDictionary;
			base.InitAuto( con, clinicalTrialId, trialSite, personId, responseTaskId, repeat );
		}

		public int VisitId
		{
			get { return( _visitId ); }
		}

		public short VisitCycle
		{
			get { return( _visitCycle ); }
		}

		public int CrfPageId
		{
			get { return( _crfPageId ); }
		}
		
		public short CrfPageCycle
		{
			get { return( _crfPageCycle ); }
		}

		public int Matches
		{
			get { return( _matches ); }
			set { _matches = value; }
		}

		public Dictionary CCDictionary
		{
			get { return( _ccDictionary ); }
		}
	}
}
