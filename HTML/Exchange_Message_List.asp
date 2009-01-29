<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Exchange_Message_List.asp
'   Author:     Andrew Newbigging, 1999
'   Purpose:    Used by TrialOffice for the purpose of returning a text message that contains
'		the body text of all the study messages that are waiting on the Server to be sent
'		to the calling site
'
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:	Mo Morris 19/5/2000 added study to the message "No study messages to download"
'				REM 18/12/02 - added MessageType < 32 so doesn't return system messages
'				ic 24/05/2004 added variable checking
'-----------------------------------------------------------------------------------------------'

'validate site
if (not fnValidateSite(request.querystring("Site"))) then
	Response.Write("ERROR:The site '" & request.querystring("Site") & "' does not exist")
	Response.End 
end if

msSQL = "SELECT MessageTimestamp,MessageBody FROM Message WHERE TrialSite = '" & request.querystring("site") & "' AND MessageReceived = 0 AND MessageDirection = 0  AND MessageType < 32"

Set rsMessageList = MACROCnn.Execute(msSQL) 

if rsMessageList.eof then
	response.write "No study messages to download"
else
	Response.Write "There are study messages to download:-" & Chr(13) & Chr(10)
	Do While not rsMessageList.eof
		response.write rsMessageList("MessageBody") & " (" & cDate(rsMessageList("MessageTimestamp")) & ")" & Chr(13) & Chr(10)
		rsMessageList.MoveNext
	Loop

end if

rsMessageList.Close
set rsMessageList = Nothing

%>

<!--#include file=CloseDataConnection.txt-->