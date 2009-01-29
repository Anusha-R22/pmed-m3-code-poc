<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=StringUtilities.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Exchange_Get_Next_MIMessage.asp
'   Author:     Mo Morris, 18 May 2000
'   Purpose:    Used by TrialOffice for the purpose of downloading entries in the MIMessage table
'		on a Server to a calling Client site.
'		This script is called once for each message and has the dual functionality of
'		marking the previous message as having been sent, using the sent time that is
'`		sent out to a Client site and read back with the next call to this script
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'		Mo Morris 22/11/00, Changed for new field MIMessageResponseTimeStamp
'		Nicky Johns 25/10/01,	MLM's 2.1 changes added (regional settings)
'		DPH 01/05/2002 - Added Error handling
'		NCJ 15 Jan 03 - Added new MIMessage fields for 3.0
'		NCJ 20 Jan 03 - Send messages in Study/Subject order
'		ic 24/05/2004 added variable checking
'		MLM 21/06/05: bug 2565: Update MIMessageSent_TZ
'-----------------------------------------------------------------------------------------------'
dim sSQL
dim rsRecordSet
dim dblSentTime

on error resume next


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
'validate site
if (not fnValidateSite(request.querystring("Site"))) then
	Response.Write("ERROR:The site '" & request.querystring("Site") & "' does not exist")
	Response.End 
end if


'check to see if a PreviousMessageID has been passed for the purpose of setting its MIMessageSent field
' MLM 21/06/05: bug 2565: Update time zone too
if request.querystring("PreviousMessageID") > "" then
	sSQL = "UPDATE MIMessage SET MIMessageSent = " & request.querystring("PreviousMessageSent") & ", " _
		& "MIMessageSent_TZ = " & session("strTimeZone") & " " _
		& "WHERE MIMessageId = " & request.querystring("PreviousMessageID") & " " _ 
		& "AND MIMessageSite = '" & request.querystring("site") & "' " _
		& "AND MIMessageSource = 0"
	MACROCnn.Execute(sSQL)

	if err.number <> 0 then
		response.write "ERROR1:" & err.number
		Response.End 
	end if
end if

' NCJ 20 Jan 03 - Added ORDER BY MIMessageTrialName, MIMessagePersonID
sSQL = "SELECT * FROM MIMessage " _
	& "WHERE MIMessageSite = '" & request.querystring("site") & "' " _
	& "AND MIMessageSource = 0 " _
	& "AND MIMessageSent = 0" _
	& "ORDER BY MIMessageTrialName,MIMessagePersonID "

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

	response.write rsRecordSet("MIMessageID")
	response.write "<br>"
	response.write rsRecordSet("MIMessageSite")
	response.write "<br>"
	response.write rsRecordSet("MIMessageSource")
	response.write "<br>"
	response.write rsRecordSet("MIMessageType")
	response.write "<br>"
	response.write rsRecordSet("MIMessageScope")
	response.write "<br>"
	response.write rsRecordSet("MIMessageObjectID")
	response.write "<br>"
	response.write rsRecordSet("MIMessageObjectSource")
	response.write "<br>"
	response.write rsRecordSet("MIMessagePriority")
	response.write "<br>"
	response.write rsRecordSet("MIMessageTrialName")
	response.write "<br>"
	response.write rsRecordSet("MIMessagePersonID")
	response.write "<br>"
	response.write rsRecordSet("MIMessageVisitID")
	response.write "<br>"
	response.write rsRecordSet("MIMessageVisitCycle")
	response.write "<br>"
	response.write rsRecordSet("MIMessageCRFPageTaskID")
	response.write "<br>"
	response.write rsRecordSet("MIMessageResponseTaskID")
	response.write "<br>"
	response.write rsRecordSet("MIMessageResponseValue")
	response.write "<br>"
	response.write rsRecordSet("MIMessageOCDiscrepancyID")
	response.write "<br>"
	' NCJ 25/10/01 - Convert the next three values
	response.write ConvertLocalNumToStandard(CStr(rsRecordSet("MIMessageCreated")))
	response.write "<br>"
	response.write ConvertLocalNumToStandard(CStr(dblSentTime))
	response.write "<br>"
	response.write ConvertLocalNumToStandard(CStr(rsRecordSet("MIMessageReceived")))
	response.write "<br>"
	response.write rsRecordSet("MIMessageHistory")
	response.write "<br>"
	response.write rsRecordSet("MIMessageProcessed")
	response.write "<br>"
	response.write rsRecordSet("MIMessageStatus")
	response.write "<br>"
	response.write rsRecordSet("MIMessageText")
	response.write "<br>"
	' DPH 17/1/2002 - MIMessageUserCode replaced with MIMessageUserNameFull
	'response.write rsRecordSet("MIMessageUserCode")
	response.write rsRecordSet("MIMessageUserName")
	response.write "<br>"
	response.write rsRecordSet("MIMessageUserNameFull")
	response.write "<br>"
	' NCJ 25/10/01 - Convert the ResponseTimeStamp
	response.write ConvertLocalNumToStandard(CStr(rsRecordSet("MIMessageResponseTimeStamp")))
	response.write "<br>"
	' NCJ 15 Jan 03 - New fields added for MACRO 3.0
	response.write rsRecordSet("MIMessageResponseCycle")
	response.write "<br>"
	response.write rsRecordSet("MIMessageCreated_TZ")
	response.write "<br>"
	response.write rsRecordSet("MIMessageReceived_TZ")
	response.write "<br>"
	' MLM 21/06/05: The MIMessageSent_TZ
	response.write session("strTimeZone")
	response.write "<br>"
	response.write rsRecordSet("MIMessageCRFPageCycle")
	response.write "<br>"
	response.write rsRecordSet("MIMessageCRFPageId")
	response.write "<br>"
	response.write rsRecordSet("MIMessageDataItemId")
	response.write "<br>"
	
end if

rsRecordSet.Close
set rsRecordSet = Nothing

%>

<!--#include file=CloseDataConnection.txt-->
