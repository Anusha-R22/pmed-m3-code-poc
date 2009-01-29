using System;
using System.Collections;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for SubjectDetails.
	/// </summary>
	class SubjectDetails
	{
		// member variables
		private string _sGuid;
		private string _sClinicalTrialName;
		private int _nClinicalTrialId;
		private string _sSiteCode;
		private string _sSubjectLabel;
		private int _nSubjectId;
		private ArrayList _alResponses;
		private BufferAPI.BufferResponseStatus _eBufferSubjectStatus;

		public SubjectDetails()
		{
			// default constructor
			_sGuid = "";
			_sClinicalTrialName = "";
			_nClinicalTrialId = -1;
			_sSiteCode = "";
			_sSubjectLabel = "";
			_nSubjectId = -1;
			_eBufferSubjectStatus = BufferAPI.BufferResponseStatus.Success;
			_alResponses = new ArrayList();
		}

		// properties
		public string Guid
		{
			get
			{
				return _sGuid;
			}
			set
			{
				_sGuid = value;
			}
		}

		public string ClinicalTrialName
		{
			get
			{
				return _sClinicalTrialName;
			}
			set
			{
				_sClinicalTrialName = value;
			}
		}

		public int ClinicalTrialId
		{
			get
			{
				return _nClinicalTrialId;
			}
			set
			{
				_nClinicalTrialId = value;
			}
		}

		public string Site
		{
			get
			{
				return _sSiteCode;
			}
			set
			{
				_sSiteCode = value;
			}
		}

		public string SubjectLabel
		{
			get
			{
				return _sSubjectLabel;
			}
			set
			{
				_sSubjectLabel = value;
			}
		}

		public int SubjectId
		{
			get
			{
				return _nSubjectId;
			}
			set
			{
				_nSubjectId = value;
			}
		}

		public ArrayList Responses
		{
			get
			{
				return _alResponses;
			}
		}

		public BufferAPI.BufferResponseStatus SubjectResponseStatus
		{
			get
			{
				return _eBufferSubjectStatus;
			}
			set
			{
				_eBufferSubjectStatus = value;
			}
		}

		public void AddResponse(ResponseDetails responseDetails)
		{
			_alResponses.Add(responseDetails);
		}
	}
}
