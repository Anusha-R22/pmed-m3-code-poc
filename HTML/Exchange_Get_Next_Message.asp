<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=StringUtilities.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Exchange_Get_Next_MIMessage.asp
'   Author:     Mo Morris, 18 May 2000
'   Purpose:    Used by TrialOffice for the purpose of downloading entries in the Message table
'		on a Server to a calling Client site.
'		This script is called once for each message and has the dual functionality of
'		marking the previous message as having been sent, using the sent time that is
'`		sent out to a Client site and read back with the next call to this script
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'	DPH 01/05/2002 - Added extra error handling
'	DPH 27/08/2002 - Study Versioning Changes - Store MessageReceivedTimeStamp
'	REM 18/12/02 - Added Messagetype < 32, so don't get system messages
'	ic 24/05/2004 added variable checking
'	MLM 21/06/05: Store MessageReceivedTimeStamp_TZ
'-----------------------------------------------------------------------------------------------'
on error resume next
dim sDblDate

'validate previousmessageid
if (not fnNumeric(request.querystring("PreviousMessageID"))) then
	Response.Write("ERROR:The PreviousMessageID '" & request.querystring("PreviousMessageID") & "' is not valid")
	Response.End 
end if
'validate site
if (not fnValidateSite(request.querystring("Site"))) then
	Response.Write("ERROR:The site '" & request.querystring("Site") & "' does not exist")
	Response.End 
end if


if request.querystring("PreviousMessageId") > "" then
	' DPH 27/08/2002 - Set MessageReceivedTimeStamp to Now
	' MLM 21/06/05: Also set the MessageReceivedTimeStamp_TZ
	sDblDate = CStr(CDbl(Now))
	sDblDate = ConvertLocalNumToStandard(sDblDate)

	msSQL = "UPDATE Message SET MessageReceived = 1, MessageReceivedTimeStamp = " & sDblDate & _
		", MessageReceivedTimeStamp_TZ = " & session("strTimeZone")
	msSQL = msSQL & " WHERE TrialSite = '" & request.querystring("site") & "' AND MessageId = " & request.querystring("PreviousMessageId") 
	MACROCnn.Execute(msSQL)

	if err.number <> 0 then
		response.write "ERROR1:" & err.number
		Response.End 
	end if
end if

msSQL = "SELECT * FROM Message WHERE TrialSite = '" & request.querystring("site") & "' AND MessageReceived = 0 AND MessageDirection = 0  AND MessageType < 32"

Set rsMessageList = MACROCnn.Execute(msSQL) 

if err.number <> 0 then
	response.write "ERROR2:" & err.number
	Response.End 
end if

if rsMessageList.eof then
		response.write "."
else

		response.write rsMessageList("MessageId")
		response.write "<br>"
		response.write rsMessageList("TrialSite")
		response.write "<br>"
		response.write rsMessageList("ClinicalTrialId")
		response.write "<br>"
		response.write rsMessageList("MessageType")
		response.write "<br>"
		response.write rsMessageList("MessageParameters")
		response.write "<br>"

end if

rsMessageList.Close
set rsMessageList = Nothing

%>

<!--#include file=CloseDataConnection.txt-->
