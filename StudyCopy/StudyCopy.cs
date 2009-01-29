using System;
using System.Runtime.InteropServices;

namespace InferMed.MACRO.StudyCopy
{
	[ComVisible(true)]
	[Guid("EDA2F3E7-7759-4271-91A5-5AA4AC936E25")]
	public interface IStudyCopy
	{
		[ComVisible(true)]
		void Init( string secCon, string dbCon, string dbCode, string userName, string userNameFull );
	}

	/// <summary>
	/// VB access point for study copy tool
	/// </summary>
	[ComVisible(true)]
	[Guid("C658C64F-1283-4979-8289-2CA7235E4480")]
	[ClassInterface(ClassInterfaceType.None)]
	public class StudyCopy : IStudyCopy
	{
		public StudyCopy()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		[ComVisible(true)]
		public void Init( string secCon, string dbCon, string dbCode, string userName, string userNameFull )
		{
//			MainForm f = new MainForm( secCon, dbCon, dbCode, userName, userNameFull );
//			f.ShowDialog();
//			f.Dispose();
		}
	}
}
