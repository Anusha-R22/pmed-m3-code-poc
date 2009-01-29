using System;

namespace InferMed.MACRO.ClinicalCoding.Interface
{
	/// <summary>
	/// Interface class between MACRO and plugins
	/// </summary>
	public interface IPlugin
	{
		void Code(string name, string version, string custom, ref string responseValue, ref string codedValue);
		string[,] FindTerm(string name, string version, string custom, string responseValue, int maxReturn, ref int totalMatches);
		string ToText(string name, string version, string custom, string codedValue);
		string ToHTML(string name, string version, string custom, string codedValue);
		string ToSingleLineText(string name, string version, string custom, string codedValue);
		string ToXmlTree(string name, string version, string custom, string codedValue);
		void ToTree(string name, string version, string custom, string codedValue);
	}
}
