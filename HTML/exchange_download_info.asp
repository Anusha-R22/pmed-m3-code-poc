<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Exchange_Message_Info.asp
'   Author:     David Hook, 2002
'   Purpose:    Used by Data Transfer for the purpose of returning info on the data to be sent
'				to the calling site
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
' REVISIONS
'	DPH 01/05/2002 - Changed error handling to <> 0
'	ic 24/05/2004 added variable checking
'-----------------------------------------------------------------------------------------------'
on error resume next

'validate site
if (not fnValidateSite(request.querystring("Site"))) then
	Response.Write("ERROR:The site '" & request.querystring("Site") & "' does not exist")
	Response.End 
end if


msSQL = "SELECT MessageBody FROM Message WHERE TrialSite = '" & request.querystring("site") & "' AND MessageReceived = 0 AND MessageDirection = 0"

Set rsMessageList = MACROCnn.Execute(msSQL) 

if err.number <> 0 then
	response.write "ERROR" & chr(13) & chr(10)
	Response.End 
end if

if rsMessageList.eof then
	response.write "No study messages to download" & chr(13) & chr(10)
else
	Do While not rsMessageList.eof
		response.write rsMessageList("MessageBody") & chr(13) & chr(10)
		rsMessageList.MoveNext
	Loop

end if

rsMessageList.Close
set rsMessageList = Nothing

sSQL = "SELECT count (MIMessageID) as MIMCount FROM MIMessage " _
	& "WHERE MIMessageSite = '" & request.querystring("site") & "' " _
	& "AND MIMessageSource = 0 " _
	& "AND MIMessageSent = 0 " _
	& "AND MIMessageType = 0"

Set rsRecordSet = MACROCnn.Execute(sSQL) 

if err.number <> 0 then
	response.write "ERROR" & chr(13) & chr(10)
	Response.End 
end if

if rsRecordSet.eof then
	nDiscrepancyCount=0
else
	nDiscrepancyCount=cInt(rsRecordSet("MIMCount"))
end if

rsRecordSet.close
set rsRecordSet = nothing
 
Response.Write "MIMessages " & nDiscrepancyCount & chr(13) & chr(10)

%>

<!--#include file=CloseDataConnection.txt-->