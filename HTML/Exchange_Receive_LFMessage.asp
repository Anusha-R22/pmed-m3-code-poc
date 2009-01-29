<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=StringUtilities.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Exchange_Receive_LFMessage.asp
'   Author:     Nicky Johns, InferMed, December 2002
'   Purpose:    Store entries in the LFMessage table sent from a calling Client site to a receiving Server.
'		This script is called once for each message
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'		NCJ 19 Dec 02 - Initial development, based on Exchange_receive_MIMessage.asp
'		NCJ 3 Jan 03 - Added QuestionId
'		ic 24/05/2004 added variable checking
'-----------------------------------------------------------------------------------------------'

On Error Resume Next

Dim lMessageID
Dim sTrialName
Dim lTrialID
Dim sTrialSite
Dim lPersonID
Dim nSource
Dim nMessageType
Dim nScope
Dim lVisitID
Dim nVisitCycle
Dim lCRFPageID
Dim nCRFPageCycle
Dim lResponseID
Dim nResponseCycle
Dim lQuestionID

Dim dMessageCreated
Dim nMessageCreated_TZ
Dim dMessageSent
Dim nMessageSent_TZ
Dim dMessageReceived
Dim nMessageReceived_TZ

Dim sUserName
Dim sUserNameFull

Dim nRBSource
Dim lRBMessageId

Dim nMessageProcessed

Dim sSQL
Dim bInsertedBefore
Dim lSequenceNo

'store the passed parameters
lMessageID = request.form("ID")
sTrialName = request.form("TrialName")
sTrialSite = request.form("Site")
lPersonID = request.form("PersonID")

lVisitID = request.form("VisitID")
nVisitCycle = request.form("VisitCycle")
lCRFPageID = request.form("CRFPageID")
nCRFPageCycle = request.form("CRFPageCycle")
lResponseID = request.form("ResponseTaskID")
nResponseCycle = request.form("ResponseCycle")
lQuestionID = request.form("QuestionID")

nSource = request.form("Source")
nMessageType = request.form("ActionType")
nScope = request.form("Scope")

sUserName = Replace(request.form("UserName"), "'", "''")
sUserNameFull = Replace(request.form("UserNameFull"), "'", "''")

nRBSource = request.form("RollbackSource")
lRBMessageID = request.form("RollbackMessageID")

dMessageCreated = request.form("MsgTimeStamp")
nMessageCreated_TZ = request.form("MsgTimeStampTZ")
dMessageSent = request.form("SentTimeStamp")
nMessageSent_TZ = request.form("SentTimeStampTZ")

'time stamp the message received flag
dMessageReceived = ConvertLocalNumToStandard(CStr(CDbl(Now)))

' REM 19/02/03 - Added Time Zone
nMessageReceived_TZ = session("strTimeZone")

'set the message processed flag to 0
nMessageProcessed = 0


'validate input
'----------------------------------------------------------------------------
'validate id
if (not fnNumeric(lMessageID)) then
	Response.Write("ERROR:The ID '" & lMessageID & "' is not valid")
	Response.End 
end if
'validate study
if (not fnValidateStudy(sTrialName)) then
	Response.Write("ERROR:The study '" & sTrialName & "' does not exist")
	Response.End 
end if
'validate site
if (not fnValidateSite(sTrialSite)) then
	Response.Write("ERROR:The site '" & sTrialSite & "' does not exist")
	Response.End 
end if
'validate personid
if (not fnNumeric(lPersonID)) then
	Response.Write("ERROR:The PersonID '" & lPersonID & "' is not valid")
	Response.End 
end if
'validate visitid
if (not fnNumeric(lVisitID)) then
	Response.Write("ERROR:The VisitID '" & lVisitID & "' is not valid")
	Response.End 
end if
'validate visitcycle
if (not fnNumeric(nVisitCycle)) then
	Response.Write("ERROR:The VisitCycle '" & nVisitCycle & "' is not valid")
	Response.End 
end if
'validate crfpageid
if (not fnNumeric(lCRFPageID)) then
	Response.Write("ERROR:The CRFPageID '" & lCRFPageID & "' is not valid")
	Response.End 
end if
'validate crfpagecycle
if (not fnNumeric(nCRFPageCycle)) then
	Response.Write("ERROR:The CRFPageCycle '" & nCRFPageCycle & "' is not valid")
	Response.End 
end if
'validate responsetaskid
if (not fnNumeric(lResponseID)) then
	Response.Write("ERROR:The ResponseTaskID '" & lResponseID & "' is not valid")
	Response.End 
end if
'validate responsecycle
if (not fnNumeric(nResponseCycle)) then
	Response.Write("ERROR:The ResponseCycle '" & nResponseCycle & "' is not valid")
	Response.End 
end if
'validate questionid
if (not fnNumeric(lQuestionID)) then
	Response.Write("ERROR:The QuestionID '" & lQuestionID & "' is not valid")
	Response.End 
end if
'validate source
if (not fnNumeric(nSource)) then
	Response.Write("ERROR:The Source '" & nSource & "' is not valid")
	Response.End 
end if
'validate actiontype
if (not fnNumeric(nMessageType)) then
	Response.Write("ERROR:The ActionType '" & nMessageType & "' is not valid")
	Response.End 
end if
'validate scope
if (not fnNumeric(nScope)) then
	Response.Write("ERROR:The Scope '" & nScope & "' is not valid")
	Response.End 
end if
'validate username
if not fnValidateUsername(sUserName) then
	Response.Write("ERROR:The UserName '" & sUserName & "' is not valid")
	Response.End 
end if
''validate fullusername
'if (not fnAlphaNumeric(sUserNameFull) or not fnLengthBetween(sUserNameFull, 1, 100)) then
'	Response.Write("ERROR:The UserNameFull '" & sUserNameFull & "' is not valid")
'	Response.End 
'end if
'validate rollbacksource
if (not fnNumeric(nRBSource)) then
	Response.Write("ERROR:The RollbackSource '" & nRBSource & "' is not valid")
	Response.End 
end if
'validate rollbackmessageid
if (not fnNumeric(lRBMessageID)) then
	Response.Write("ERROR:The RollbackMessageID '" & lRBMessageID & "' is not valid")
	Response.End 
end if
'validate msgtimestamp
if (not fnNumeric(dMessageCreated)) then
	Response.Write("ERROR:The MsgTimeStamp '" & dMessageCreated & "' is not valid")
	Response.End 
end if
'validate msgtimestamptz
if (not fnNumeric(nMessageCreated_TZ)) then
	Response.Write("ERROR:The MsgTimeStampTZ '" & nMessageCreated_TZ & "' is not valid")
	Response.End 
end if
'validate senttimestamp
if (not fnNumeric(dMessageSent)) then
	Response.Write("ERROR:The SentTimeStamp '" & dMessageSent & "' is not valid")
	Response.End 
end if
'validate senttimestamptz
if (not fnNumeric(nMessageSent_TZ)) then
	Response.Write("ERROR:The nMessageSent_TZ '" & nMessageSent_TZ & "' is not valid")
	Response.End 
end if
'----------------------------------------------------------------------------
         

' DPH 20/11/2002 Check if this exact message has been received previously. 
' If it has do not write to database again as will cause DB Key violation
sSQL = "SELECT Count(*) AS LFCount FROM LFMessage" _
	& " WHERE MessageId = " & lMessageID _
        & " AND TrialSite ='" & sTrialSite & "'" _
        & " AND Source = " & nSource
Set rsTemp = MACROCnn.Execute(sSQL)

if err.number <> 0 then
	response.write "ERROR1: " & err.number & ", " & err.description & ", in Exchange_Receive_LFMessage.asp"
	Response.End
end if 

' NCJ 15 Jan 03 - Force to Longs for correct comparison
If CLng(rsTemp(0).Value) > CLng(0) Then
	bInsertedBefore = true
Else
	bInsertedBefore = false
End If
rsTemp.Close
set rsTemp = nothing
	
If Not bInsertedBefore Then

	' Get the TrialId
	sSQL = "SELECT ClinicalTrialId FROM ClinicalTrial WHERE ClinicalTrialName = '" & sTrialName & "'"
	Set rsTemp = MACROCnn.Execute(sSQL)
	if err.number <> 0 then
		Response.write "ERROR READING FROM ClinicalTrial:" & err.number
		Response.End
	end if 
	lTrialID = rsTemp("ClinicalTrialId")
	rsTemp.Close
	set rsTemp = nothing


	' Get the next Sequence No
	sSQL = "SELECT max(SequenceNo) as MaxSeqNo FROM LFMessage "
	Set rsTemp = MACROCnn.Execute(sSQL)
	if err.number <> 0 then
		Response.write "ERROR READING FROM LFMESSAGE:" & err.number & ", " & err.description & ", in Exchange_Receive_LFMessage.asp"
		Response.End
	end if 
	If IsNull(rsTemp("MaxSeqNo")) Then
		lSequenceNo = 1
	Else
		lSequenceNo = Clng(rsTemp("MaxSeqNo")) + 1
	End If
	rsTemp.Close
	set rsTemp = nothing


	'store the new message
    	sSQL = "INSERT INTO LFMessage "
    	sSQL = sSQL & "( ClinicalTrialName, ClinicalTrialId, TrialSite, PersonId, " _
                & " Source, MessageID, Scope, "
    	sSQL = sSQL & " VisitId, VisitCycleNumber, " _
                & " CRFPageId, CRFPageCycleNumber, " _
                & " ResponseTaskId, RepeatNumber, " _
                & " DataItemId, " _
                & " UserName, UserNameFull, "
    	sSQL = sSQL & " MsgType, ProcessedStatus, " _
                & " RollbackSource, RollbackMessageID, " _
                & " MsgCreatedTimeStamp, MsgCreatedTimeStamp_TZ, " _
                & " ProcessedTimeStamp, ProcessedTimeStamp_TZ, " _
                & " SentTimeStamp, SentTimeStamp_TZ, " _
                & " ReceivedTimeStamp, ReceivedTimeStamp_TZ, " _
                & " SequenceNo "
    	sSQL = sSQL & " ) "
    	sSQL = sSQL & "VALUES ('" & sTrialName & "', " & lTrialID & ", '" & sTrialSite & "'," & lPersonID & ", " _
		& nSource & ", " & lMessageID & ", " & nScope & ", " _
		& lVisitID & ", " & nVisitCycle & ", " _
		& lCRFPageID & ", " & nCRFPageCycle & ", " _
		& lResponseID & ", " & nResponseCycle & ", " _
		& lQuestionID & ", " _
		& "'" & sUserName & "', '" & sUserNameFull & "', "
    	sSQL = sSQL & nMessageType & ", " & nMessageProcessed & ", " _
		& nRBSource & ", " & lRBMessageId & ", " _
		& dMessageCreated & ", " & nMessageCreated_TZ & ", " _
		& "0, 0, " _
		& dMessageSent & ", " & nMessageSent_TZ & ", " _
		& dMessageReceived & ", " & nMessageReceived_TZ & ", " _
		& lSequenceNo _
		& ")"

	MACROCnn.Execute sSQL
			
	if err.number <> 0 then
		response.write "ERROR4: " & err.number & ", " & err.description & ", in Exchange_Receive_LFMessage.asp"
		Response.End
	end if

end if


if err.number <> 0 then
	response.write "ERROR5: " & err.number & ", " & err.description & ", in Exchange_Receive_LFMessage.asp"
else
	response.write "SUCCESS"
end if

%>

<!--#include file=CloseDataConnection.txt-->