using System;

namespace InferMed.MACRO.StudyCopy
{
	/// <summary>
	/// Global study copy objects
	/// </summary>
	public class StudyCopyGlobal
	{
		private StudyCopyGlobal()
		{
		}

		//macro element types
		public const string _QGROUP = "0";
		public const string _TEXTBOX = "1";
		public const string _OPTIONBUTTONS = "2";
		public const string _POPUPLIST = "4";
		public const string _CALENDAR = "8";
		public const string _ATTACHMENT = "32";
		public const string _LINE = "16385";
		public const string _TEXTCOMMENT = "16386";
		public const string _PICTURE = "16388";
		public const string _HOTLINK = "16390";

		public const string _TEXT = "0";
		public const string _CATEGORY = "1";
		public const string _INTEGER = "2";
		public const string _REAL = "3";
		public const string _DATE = "4";
		public const string _MULTIMEDIA = "5";
		public const string _LABTEST = "6";
		public const string _MULTILINETEXT = "7";
		public const string _THESAURUS = "8";

		//logging priority
		public enum LogPriority
		{
			Low, Normal, High
		}

		//macro element types
		public enum ElementType
		{
			eForm, eFormElement, DataItem
		}

		public static string ReplaceControlChars(string s)
		{
			s = s.Replace("\n", " "); //newline
			s = s.Replace("\r", " "); //carriage return

			return (s);
		}

		/// <summary>
		/// Return control type string
		/// </summary>
		/// <param name="cType"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		/// <returns></returns>
		public static string GetControlType( string cType )
		{
			string controlType = "";

			switch( cType )
			{
				case _QGROUP: controlType = "Question group"; break;
				case _TEXTBOX: controlType = "Text box"; break;
				case _OPTIONBUTTONS: controlType = "Option buttons"; break;
				case _POPUPLIST: controlType = "Popup list"; break;
				case _CALENDAR: controlType = "Calendar"; break;
				case _ATTACHMENT: controlType = "Attachment"; break;
				case _LINE: controlType = "Line"; break;
				case _TEXTCOMMENT: controlType = "Text comment"; break;
				case _PICTURE: controlType = "Picture"; break;
				case _HOTLINK: controlType = "Hotlink"; break;
				default: controlType = "Unknown (" + cType + ")"; break;
			}

			return( controlType );
		}

		/// <summary>
		/// Return data type string
		/// </summary>
		/// <param name="dType"></param>
		/// <returns></returns>
		public static string GetDataType( string dType )
		{
			string dataType = "";

			switch( dType )
			{
				case _TEXT: dataType = "Text"; break;
				case _CATEGORY: dataType = "Category"; break;
				case _INTEGER: dataType = "Integer"; break;
				case _REAL: dataType = "Real"; break;
				case _DATE: dataType = "Date"; break;
				case _MULTIMEDIA: dataType = "Multimedia"; break;
				case _LABTEST: dataType = "Lab test"; break;
				case _MULTILINETEXT: dataType = "Multiline text"; break;
				case _THESAURUS: dataType = "Thesaurus"; break;
				case "": dataType = ""; break;
				default: dataType = "Unknown (" + dType + ")"; break;
			}

			return( dataType );
		}
	}
}
