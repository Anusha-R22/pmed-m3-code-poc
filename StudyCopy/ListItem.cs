using System;

namespace InferMed.MACRO.StudyMerge
{
	/// <summary>
	/// Code / value object
	/// </summary>
	public class ComboItem
	{
		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="sValue">Value parameter</param>
		/// <param name="sCode">Code parameter</param>
		public ComboItem(string sValue, string sCode)
		{
			_value = sValue;
			_code = sCode;
		}

		/// <summary>
		/// Returns value
		/// </summary>
		public String Value
		{
			get { return  _value; }
		}
		
		/// <summary>
		/// Returns code
		/// </summary>
		public String Code
		{
			get { return  _code; }
		}

		/// <summary>
		/// Returns value
		/// </summary>
		/// <returns></returns>
		override public string ToString()
		{
			return _value;
		}

		private string _value;
		private string _code;
	}
}
