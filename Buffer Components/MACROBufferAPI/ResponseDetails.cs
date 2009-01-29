using System;


namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for ResponseDetails.
	/// </summary>
	class ResponseDetails
	{
		// member variables
		private string _sGuid;
		private string _sVisitCode;
		private int _nVisitId;
		private int _nVisitCycle;
		private string _sCRFPageCode;
		private int _nCRFPageId;
		private int _nCRFPageCycle;
		private string _sDataItemCode;
		private int _nDataItemId;
		private int _nDataItemRepeatNumber;
		private string _sResponseValue;
		private DateTime _dtCompareDate;
		private BufferAPI.BufferResponseStatus _eBufferResponseStatus;
		private BufferAPI.BufferCommitStatus _eBufferCommitStatus;
		private string _sUserName;
		private string _sUserNameFull;

		public ResponseDetails()
		{
			// default constructor
			_sGuid = "";
			_sVisitCode = "";
			_nVisitId = -1;
			_nVisitCycle = -1;
			_sCRFPageCode = "";
			_nCRFPageId = -1;
			_nCRFPageCycle = -1;
			_sDataItemCode = "";
			_nDataItemId = -1;
			_nDataItemRepeatNumber = 1;
			_sResponseValue = "";
			_dtCompareDate = DateTime.FromOADate(0);
			_eBufferResponseStatus = BufferAPI.BufferResponseStatus.Success;
			_eBufferCommitStatus = BufferAPI.BufferCommitStatus.NotCommitted;
			_sUserName = "";
			_sUserNameFull = "";
		}

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

		public string VisitCode
		{
			get
			{
				return _sVisitCode;
			}
			set
			{
				_sVisitCode = value;
			}
		}

		public int VisitId
		{
			get
			{
				return _nVisitId;
			}
			set
			{
				_nVisitId = value;
			}
		}

		public int VisitCycle
		{
			get
			{
				return _nVisitCycle;
			}
			set
			{
				_nVisitCycle = value;
			}
		}

		public string CRFPageCode
		{
			get
			{
				return _sCRFPageCode;
			}
			set
			{
				_sCRFPageCode = value;
			}
		}

		public int CRFPageId
		{
			get
			{
				return _nCRFPageId;
			}
			set
			{
				_nCRFPageId = value;
			}
		}

		public int CRFPageCycle
		{
			get
			{
				return _nCRFPageCycle;
			}
			set
			{
				_nCRFPageCycle = value;
			}
		}

		public string DataItemCode
		{
			get
			{
				return _sDataItemCode;
			}
			set
			{
				_sDataItemCode = value;
			}
		}

		public int DataItemId
		{
			get
			{
				return _nDataItemId;
			}
			set
			{
				_nDataItemId = value;
			}
		}

		public int DataItemRepeatNo
		{
			get
			{
				return _nDataItemRepeatNumber;
			}
			set
			{
				_nDataItemRepeatNumber = value;
			}
		}

		public string ResponseValue
		{
			get
			{
				return _sResponseValue;
			}
			set
			{
				_sResponseValue = value;
			}
		}

		public DateTime DataCompareDate
		{
			get
			{
				return _dtCompareDate;
			}
			set
			{
				_dtCompareDate = value;
			}
		}

		public BufferAPI.BufferResponseStatus BufferResponseStatus
		{
			get
			{
				return _eBufferResponseStatus;
			}
			set
			{
				_eBufferResponseStatus = value;
			}
		}

		public BufferAPI.BufferCommitStatus BufferCommitStatus
		{
			get
			{
				return _eBufferCommitStatus;
			}
			set
			{
				_eBufferCommitStatus = value;
			}
		}

		public string Username
		{
			get
			{
				return _sUserName;
			}
			set
			{
				_sUserName = value;
			}
		}

		public string UsernameFull
		{
			get
			{
				return _sUserNameFull;
			}
			set
			{
				_sUserNameFull = value;
			}
		}
	}
}
