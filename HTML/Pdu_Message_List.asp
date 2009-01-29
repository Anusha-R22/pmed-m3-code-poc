<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Pdu_Message_List.asp
'   Author:     David Hook, 2005
'   Purpose:    Used by TrialOffice for the purpose of returning a text message that contains
'				the body text of all the PDU messages that are waiting on the Server to be sent
'				to the calling site
'
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'-----------------------------------------------------------------------------------------------'

'validate site
if (not fnValidateSite(request.querystring("Site"))) then
	Response.Write("ERROR:The site '" & request.querystring("Site") & "' does not exist")
	Response.End 
end if

msSQL = "SELECT MessageTimestamp,MessageBody FROM Message WHERE TrialSite = '" & Request.querystring("site") & "' AND MessageReceived = 0 AND MessageDirection = 0 AND MessageType IN (50, 51)"

Set rsMessageList = MACROCnn.Execute(msSQL) 

if rsMessageList.eof then
	response.write "There are no PDU messages to download"
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
