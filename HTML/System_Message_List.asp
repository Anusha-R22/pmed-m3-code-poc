<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Exchange_Message_List.asp
'   Author:     Richard Meinesz, 2002
'   Purpose:    Used by TrialOffice for the purpose of returning a text message that contains
'				the body text of all the system messages that are waiting on the Server to be sent
'				to the calling site
'
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'	ic 24/05/2004 added variable checking
'-----------------------------------------------------------------------------------------------'

'validate site
if (not fnValidateSite(Request.Form("site"))) then
	Response.Write("ERROR:The site '" & Request.Form("site") & "' does not exist")
	Response.End 
end if


msSQL = "SELECT MessageTimestamp,MessageBody FROM Message WHERE TrialSite = '" & Request.Form("site") & "' AND MessageReceived = 0 AND MessageDirection = 0 AND (MessageType >= 32 AND MessageType < 50)"

Set rsMessageList = MACROCnn.Execute(msSQL) 

if rsMessageList.eof then
	response.write "There are no system messages to download"
else
	Do While not rsMessageList.eof
		response.write rsMessageList("MessageBody") & " (" & cDate(rsMessageList("MessageTimestamp")) & ")" & chr(13) & chr(10) & vbtab
		rsMessageList.MoveNext
	Loop

end if

rsMessageList.Close
set rsMessageList = Nothing

%>

<!--#include file=CloseDataConnection.txt-->
