using System;
using System.Runtime.InteropServices;

namespace InferMed.MACRO.ClinicalCoding.MACRO_CC
{
	[ComVisible(true)]
	[Guid("9e83f4db-896c-4c87-8c1f-eada0bbd51ac")]
	public interface IMACRO_CC
	{
		[ComVisible(true)]
		void Init( string secCon, string dbCon, string dbCode, string userName, string userNameFull, 
			bool codeResponse, bool importDictionary, bool autoEncode );
	}

	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	[ComVisible(true)]
	[Guid("0864cd5e-7c3b-429b-b0f5-a1a406b23195")]
	[ClassInterface(ClassInterfaceType.None)]
	public class MACRO_CC : IMACRO_CC
	{
		public MACRO_CC()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		[ComVisible(true)]
		public void Init( string secCon, string dbCon, string dbCode, string userName, string userNameFull, 
			bool codeResponse, bool importDictionary, bool autoEncode )
		{
//			MainForm f = new MainForm( secCon, dbCon, dbCode, userName, userNameFull, codeResponse, importDictionary, autoEncode );
//			f.ShowDialog();
//			f.Dispose();
		}
	}
}
