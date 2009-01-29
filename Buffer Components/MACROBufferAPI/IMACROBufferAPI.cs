using System;
using System.Runtime.InteropServices;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Interface for MACROBufferAPI - make COM visible & gen GUID
	/// </summary>
	[ComVisible(true)]
	[Guid("b992e709-bc44-400e-b47f-1e6f7f7644ad")]
	public interface IMACROBufferAPI
	{
		[ComVisible(true)]
		string WriteBufferMessage(string sBufferMessageXML, bool bCommitToMACRO);		
	}
}
