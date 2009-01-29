<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Exchange_LFMessage_List.asp
'   Author:     Nicky Johns, December 2002
'   Purpose:    Used by TrialOffice to return a text message that contains
'		the number of Lock/Freeze messages that are waiting on
'		the Server to be sent to the calling site
'
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'		NCJ 20 Dec 02 - Initial development, based on Exchange_MIMessage_List.asp
'		ic 24/05/2004 added variable checking
'-----------------------------------------------------------------------------------------------'

dim sSQL
dim nMessageCount
dim rsRecordSet

'validate site
if (not fnValidateSite(request.querystring("Site"))) then
	Response.Write("ERROR:The site '" & request.querystring("Site") & "' does not exist")
	Response.End 
end if

'Get the number of unsent LF messages
sSQL = "SELECT count (MessageID) as MsgCount FROM LFMessage " _
	& "WHERE TrialSite = '" & request.querystring("site") & "' " _
	& "AND Source = 0 " _
	& "AND SentTimestamp = 0 "

Set rsRecordSet = MACROCnn.Execute(sSQL) 

if rsRecordSet.eof then
	nMessageCount=0
else
	nMessageCount=cInt(rsRecordSet("MsgCount"))
end if


'Create a message to send back to calling client installation
if nMessageCount = 0 then
	' NB The exact wording of this message must match that in frmDataTransfer!
	response.write "No Lock/Freeze messages to download" '& chr(13) & chr(10)
else
	response.write "There are Lock/Freeze messages to download:-" & chr(13) & chr(10) & nMessageCount
	if nMessageCount= 1 then
		response.write " message"
	else
		response.write " messages"
	end if
end if

rsRecordSet.Close
set rsRecordSet = Nothing

%>

<!--#include file=CloseDataConnection.txt-->