using System;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for CategoryItem.
	/// </summary>
	class CategoryItem
	{
		// member variables
		private string _sCategoryCode;
		private string _sCategoryValue;

		public CategoryItem()
		{}

		public CategoryItem(string sCategoryCode, string sCategoryValue)
		{
			_sCategoryCode = sCategoryCode;
			_sCategoryValue = sCategoryValue;
		}

		// properties
		public string Code
		{
			get
			{
				return _sCategoryCode;
			}
			set
			{
				_sCategoryCode = value;
			}
		}

		public string Value
		{
			get
			{
				return _sCategoryValue;
			}
			set
			{
				_sCategoryValue = value;
			}
		}
	}
}
