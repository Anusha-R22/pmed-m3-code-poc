using System;
using System.Collections;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for ResponseDataItem.
	/// </summary>
	class ResponseDataItem
	{
		// member variables
		private int _nDataItemId;
		private string _sDataItemCode;
		private BufferAPI.MACRODataTypes _eMACRODataType;
		private string _sDataItemFormat;
		private int _nDataItemLength;
		private ArrayList _alCategories;

		public ResponseDataItem()
		{
			_nDataItemId = BufferAPI._DEFAULT_MISSING_NUMERIC;
			_sDataItemCode = "";
			_eMACRODataType = BufferAPI.MACRODataTypes.Text;
			_sDataItemFormat = "";
			_nDataItemLength = BufferAPI._DEFAULT_MISSING_NUMERIC;
			_alCategories = new ArrayList();
		}

		public ResponseDataItem(int nId, string sCode, BufferAPI.MACRODataTypes eMACRODataType, string sFormat, int nLength)
		{	
			_nDataItemId = nId;
			_sDataItemCode = sCode;
			_eMACRODataType = eMACRODataType;
			_sDataItemFormat = sFormat;
			_nDataItemLength = nLength;
			_alCategories = new ArrayList();
		}

		// properties
		public int Id
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

		public string Code
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

		public BufferAPI.MACRODataTypes MACRODataType
		{
			get
			{
				return _eMACRODataType;
			}
			set
			{
				_eMACRODataType = value;
			}
		}

		public string Format
		{
			get
			{
				return _sDataItemFormat;
			}
			set
			{
				_sDataItemFormat = value;
			}
		}

		public int Length
		{
			get
			{
				return _nDataItemLength;
			}
			set
			{
				_nDataItemLength = value;
			}
		}

		public ArrayList Categories
		{
			get
			{
				return _alCategories;
			}
		}

		public void AddCategory(CategoryItem catItem)
		{
			_alCategories.Add(catItem);
		}
	}
}
