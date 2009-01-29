using System;
using System.Collections;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for StudyInfo.
	/// </summary>
	class StudyInfo
	{
		// private members
		private ArrayList _alVisit;
		private ArrayList _alEform;
		private ArrayList _alDataItem;

		public StudyInfo()
		{
			_alVisit = new ArrayList();
			_alEform = new ArrayList();
			_alDataItem = new ArrayList();
		}	

		// Add Visit
		public void AddVisit(int nVisitId, string sVisitCode)
		{
			VisitInfo visitInfo = new VisitInfo(nVisitId, sVisitCode);
			_alVisit.Add(visitInfo);
		}

		// Add Eform
		public void AddEForm(int nEformId, string sEformCode)
		{
			EformInfo eformInfo = new EformInfo(nEformId, sEformCode);
			_alEform.Add(eformInfo);
		}

		// Add DataItem
		public void AddDataItem(ResponseDataItem dataItemInfo)
		{
			_alDataItem.Add(dataItemInfo);
		}

		// functions
		/// <summary>
		/// Get visit id
		/// </summary>
		/// <param name="sVisitCode"></param>
		/// <returns></returns>
		public int GetVisitId(string sVisitCode)
		{
			int nVisitId = BufferAPI._DEFAULT_MISSING_NUMERIC;
			// loop through VisitInfo
			foreach(VisitInfo visitInfo in _alVisit)
			{
				if(visitInfo.VisitCode == sVisitCode)
				{
					nVisitId = visitInfo.VisitId;
					break;
				}
			}
			return nVisitId;
		}

		/// <summary>
		/// get eform id
		/// </summary>
		/// <param name="sEformCode"></param>
		/// <returns></returns>
		public int GetEformId(string sEformCode)
		{
			int nEformId = BufferAPI._DEFAULT_MISSING_NUMERIC;
			// loop through EformInfo
			foreach(EformInfo eformInfo in _alEform)
			{
				if(eformInfo.EFormCode == sEformCode)
				{
					// have match
					nEformId = eformInfo.EFormId;
					break;
				}
			}
			return nEformId;
		}

		/// <summary>
		/// get data item id
		/// </summary>
		/// <param name="sDataItemCode"></param>
		/// <returns></returns>
		public int GetDataItemId(string sDataItemCode)
		{
			int nDataItemId = BufferAPI._DEFAULT_MISSING_NUMERIC;
			// loop through dataitem
			foreach(ResponseDataItem responseDataItem in _alDataItem)
			{
				if(responseDataItem.Code == sDataItemCode)
				{
					// have a match
					nDataItemId = responseDataItem.Id;
					break;
				}
			}
			return nDataItemId;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="nDataItemId"></param>
		/// <returns></returns>
		public ResponseDataItem GetResponseDataItem(int nDataItemId)
		{
			ResponseDataItem responseDataItem = new ResponseDataItem();

			// loop through dataitem
			foreach(ResponseDataItem rdi in _alDataItem)
			{
				if(rdi.Id == nDataItemId)
				{
					// have a match
					responseDataItem = new ResponseDataItem(rdi.Id, rdi.Code, rdi.MACRODataType, rdi.Format, rdi.Length);
					foreach(CategoryItem catItem in rdi.Categories)
					{
						responseDataItem.AddCategory(catItem);
					}
					break;
				}
			}
			return responseDataItem;
		}
	}

	// store a visit id / visit code pairing 
	class VisitInfo
	{
		// member variables
		private int _nVisitId;
		private string _sVisitCode;

		public VisitInfo(int nVisitId, string sVisitCode)
		{
			_nVisitId = nVisitId;
			_sVisitCode = sVisitCode;
		}

		// properties
		public int VisitId
		{
			get
			{
				return _nVisitId;
			}
		}

		public string VisitCode
		{
			get
			{
				return _sVisitCode;
			}
		}
	}

	// store an eform id / eform code pairing
	class EformInfo
	{
		// member variables
		private int _nEformId;
		private string _sEformCode;

		public EformInfo(int nEformId, string sEformCode)
		{
			_nEformId= nEformId;
			_sEformCode = sEformCode;
		}

		// member variables
		public int EFormId
		{
			get
			{
				return _nEformId;
			}
		}

		public string EFormCode
		{
			get
			{
				return _sEformCode;
			}
		}
	}
}
