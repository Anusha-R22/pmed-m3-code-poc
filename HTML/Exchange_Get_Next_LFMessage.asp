<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=StringUtilities.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Exchange_Get_Next_LFMessage.asp
'   Author:     Nicky Johns, December 2002
'   Purpose:    Download entries from the LFMessage table on a Server to a calling Client site.
'		This script is called once for each message and has the dual functionality of
'		marking the previous message as having been sent, using the sent time that is
'`		sent out to a Client site and read back with the next call to this script
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'		NCJ 19 Dec 02 - File created, based on Exchange_Get_Next_MIMessage.asp
'		NCJ 3 Jan 03 - Added DataItemId
'		ic 24/05/2004 added variable checking
'-----------------------------------------------------------------------------------------------'
dim sSQL
dim rsRecordSet
dim dblSentTime

'validate previousmessageid
if (not fnNumeric(request.querystring("PreviousMessageID"))) then
	Response.Write("ERROR:The PreviousMessageID '" & request.querystring("PreviousMessageID") & "' is not valid")
	Response.End 
end if
'validate previousmessagesent
if (not fnNumeric(request.querystring("PreviousMessageSent"))) then
	Response.Write("ERROR:The PreviousMessageSent '" & request.querystring("PreviousMessageSent") & "' is not valid")
	Response.End 
end if
'validate previousmessagesenttz
if (not fnNumeric(request.querystring("PreviousMessageSentTZ"))) then
	Response.Write("ERROR:The PreviousMessageSentTZ '" & request.querystring("PreviousMessageSentTZ") & "' is not valid")
	Response.End 
end if
'validate site
if (not fnValidateSite(request.querystring("Site"))) then
	Response.Write("ERROR:The site '" & request.querystring("Site") & "' does not exist")
	Response.End 
end if

on error resume next

'check to see if a PreviousMessageID has been passed for the purpose of setting its MessageSent field
if request.querystring("PreviousMessageID") > "" then
	sSQL = "UPDATE LFMessage SET " _
		& " SentTimeStamp = " & request.querystring("PreviousMessageSent") & ", " _
		& " SentTimestamp_TZ = " & request.querystring("PreviousMessageSentTZ") _
		& " WHERE MessageId = " & request.querystring("PreviousMessageID")  _ 
		& " AND TrialSite = '" & request.querystring("site") & "'" _
		& " AND Source = 0"
	MACROCnn.Execute(sSQL)

	if err.number <> 0 then
		response.write "ERROR1:" & err.number
		Response.End 
	end if
end if

' Select the unsent LFMessages from the Server (i.e. Source = 0)
' We MUST send them in the order of creation
sSQL = "SELECT * FROM LFMessage " _
	& "WHERE TrialSite = '" & request.querystring("site") & "' " _
	& " AND Source = 0" _
	& " AND SentTimeStamp = 0" _
	& " ORDER BY ClinicalTrialName, PersonId, MessageID"

Set rsRecordSet = MACROCnn.Execute(sSQL) 

if err.number <> 0 then
	response.write "ERROR2:" & err.number
	Response.End 
end if

'Either write the contents of a message back or a single full stop when there are no more messages to send
if rsRecordSet.eof then
	response.write "."
else
	'set the Sent time on the message being sent out
	'note that this same time will come back with the PreviousMessageID
	'so that it can be set into the database on the server
	dblSentTime=CDbl(Now)

	response.write rsRecordSet("MessageID")
	response.write "<br>"
	response.write rsRecordSet("ClinicalTrialName")
	response.write "<br>"
	response.write rsRecordSet("TrialSite")
	response.write "<br>"
	response.write rsRecordSet("PersonId")
	response.write "<br>"
	response.write rsRecordSet("VisitId")
	response.write "<br>"
	response.write rsRecordSet("VisitCycleNumber")
	response.write "<br>"
	response.write rsRecordSet("CRFPageId")
	response.write "<br>"
	response.write rsRecordSet("CRFPageCycleNumber")
	response.write "<br>"
	response.write rsRecordSet("ResponseTaskId")
	response.write "<br>"
	response.write rsRecordSet("RepeatNumber")
	response.write "<br>"
	response.write rsRecordSet("DataItemId")
	response.write "<br>"
	response.write rsRecordSet("Source")
	response.write "<br>"
	response.write rsRecordSet("MsgType")
	response.write "<br>"
	response.write rsRecordSet("Scope")
	response.write "<br>"
	response.write rsRecordSet("UserName")
	response.write "<br>"
	response.write rsRecordSet("UserNameFull")
	response.write "<br>"
	response.write rsRecordSet("RollbackSource")
	response.write "<br>"
	response.write rsRecordSet("RollbackMessageId")
	response.write "<br>"
	response.write ConvertLocalNumToStandard(CStr(rsRecordSet("MsgCreatedTimestamp")))
	response.write "<br>"
	response.write ConvertLocalNumToStandard(CStr(rsRecordSet("MsgCreatedTimestamp_TZ")))
	response.write "<br>"
	response.write ConvertLocalNumToStandard(CStr(dblSentTime))
	response.write "<br>"
	response.write session("strTimeZone")
	response.write "<br>"

end if

rsRecordSet.Close
set rsRecordSet = Nothing

%>

<!--#include file=CloseDataConnection.txt-->
