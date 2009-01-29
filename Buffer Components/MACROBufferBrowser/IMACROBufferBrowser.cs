using System;
using System.Runtime.InteropServices;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// MACRO Buffer Data Browser Class Public Interface for COM
	/// </summary>
	[ComVisible(true)]
	[Guid("ef78dc22-3973-4fcb-a255-269056cde57e")]
	public interface IMACROBufferBrowser
	{
		[ComVisible(true)]
		string BufferSummaryPage(string serialisedUser, bool isUserHex, int studyId, string site, int subjectNo);
		[ComVisible(true)]
		string LoadBufferDataBrowser(string serialisedUser, bool isUserHex, int studyId, string site, int subjectNo,
			string bookMark);
		[ComVisible(true)]
		string GetBufferSaveResultsPage(string serialisedUser, bool isUserHex, string formData);
		[ComVisible(true)]
		string WorkingDirectory();
		[ComVisible(true)]
		string GetBufferTargetSelectionPage(string serialisedUser, bool isUserHex, string formData, string bookMark);
		[ComVisible(true)]
		string SaveBufferTargetSelection( string serialisedUser, bool isUserHex, string formData );
	}
}
