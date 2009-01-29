<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Exchange_MIMessage_List.asp
'   Author:     Mo Morris, 18 May 2000
'   Purpose:    Used by TrialOffice for the purpose of returning a text message that contains
'		the number of Discrepancy, Message, Note and SDV messages that are waiting on
'		the Server to be sent to the calling site
'
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'		Mo Morris 6/6/2000 SR 3554, cast variables into Integers when reading back from Oracle
'		DPH 16/04/2002 - Removed vbCrLf from end of "No Messages" message
'		ic 24/05/2004 added variable checking
'-----------------------------------------------------------------------------------------------'

dim sSQL
dim nDiscrepancyCount
dim nMessageCount
dim nNoteCount
dim nSDVCount
dim rsRecordSet

'validate site
if (not fnValidateSite(request.querystring("Site"))) then
	Response.Write("ERROR:The site '" & request.querystring("Site") & "' does not exist")
	Response.End 
end if


'Get the number of unsent Discrepancy messages
sSQL = "SELECT count (MIMessageID) as MIMCount FROM MIMessage " _
	& "WHERE MIMessageSite = '" & request.querystring("site") & "' " _
	& "AND MIMessageSource = 0 " _
	& "AND MIMessageSent = 0 " _
	& "AND MIMessageType = 0"

Set rsRecordSet = MACROCnn.Execute(sSQL) 

if rsRecordSet.eof then
	nDiscrepancyCount=0
else
	nDiscrepancyCount=cInt(rsRecordSet("MIMCount"))
end if

'Get the number of unsent Message messages
sSQL = "SELECT count (MIMessageID) as MIMCount FROM MIMessage " _
	& "WHERE MIMessageSite = '" & request.querystring("site") & "' " _
	& "AND MIMessageSource = 0 " _
	& "AND MIMessageSent = 0 " _
	& "AND MIMessageType = 1"

Set rsRecordSet = MACROCnn.Execute(sSQL) 

if rsRecordSet.eof then
	nMessageCount=0
else
	nMessageCount=cInt(rsRecordSet("MIMCount"))
end if

'Get the number of unsent Note messages
sSQL = "SELECT count (MIMessageID) as MIMCount FROM MIMessage " _
	& "WHERE MIMessageSite = '" & request.querystring("site") & "' " _
	& "AND MIMessageSource = 0 " _
	& "AND MIMessageSent = 0 " _
	& "AND MIMessageType = 2"

Set rsRecordSet = MACROCnn.Execute(sSQL) 

if rsRecordSet.eof then
	nNoteCount=0
else
	nNoteCount=cInt(rsRecordSet("MIMCount"))
end if

'Get the number of unsent SDV messages
sSQL = "SELECT count (MIMessageID) as MIMCount FROM MIMessage " _
	& "WHERE MIMessageSite = '" & request.querystring("site") & "' " _
	& "AND MIMessageSource = 0 " _
	& "AND MIMessageSent = 0 " _
	& "AND MIMessageType = 3"

Set rsRecordSet = MACROCnn.Execute(sSQL) 

if rsRecordSet.eof then
	nSDVCount=0
else
	nSDVCount=cInt(rsRecordSet("MIMCount"))
end if

'Create a message to send back to calling client installation
if nDiscrepancyCount + nMessageCount + nNoteCount + nSDVCount = 0 then
	response.write "No user messages to download" '& chr(13) & chr(10)
else
	response.write "There are user messages to download:-" & chr(13) & chr(10) & nDiscrepancyCount
	if nDiscrepancyCount = 1 then
		response.write " Discrepancy"
	else
		response.write " Discrepancies"
	end if
	response.write chr(13) & chr(10) & nMessageCount
	if nMessageCount = 1 then
		response.write " Message"
	else
		response.write " Messages"
	end if
	response.write chr(13) & chr(10) & nNoteCount
	if nNoteCount = 1 then
		response.write " Note"
	else
		response.write " Notes"
	end if
	response.write chr(13) & chr(10) & nSDVCount
	if nSDVCount = 1 then
		response.write " SDV"
	else
		response.write " SDVs"
	end if
end if

rsRecordSet.Close
set rsRecordSet = Nothing

%>

<!--#include file=CloseDataConnection.txt-->