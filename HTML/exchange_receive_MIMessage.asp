<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=StringUtilities.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Exchange_receive_MIMessage.asp
'   Author:     Mo Morris, 19 May 2000
'   Purpose:    Used by TrialOffice for the purpose of storing entries in the MIMessage table
'		that have been sent from a calling Client site to a receiving Server.
'		This script is called once for each message
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'		Mo Morris 22/11/00, Changed for new field MIMessageResponseTimeStamp
'		Mo Morris 23/11/00,	On Error Resume Next line added
'		Nicky Johns 25/10/01,	MLM's 2.1 changes added (regional settings)
'		DPH 04/04/2002 - Backwards Compatibility Changes
'		DPH 18/11/2002 - Stop "historic" mimessages changing current mimessage status
'		NCJ 15 Jan 03 - Added new MIMessage fields for 3.0
'		TA  12/03/2003 - Update MIMsgStatus in subject data tables
'		ic 24/05/2004 added variable checking
'-----------------------------------------------------------------------------------------------'

On Error Resume Next

Dim lMIMessageID
Dim sMIMessageSite
Dim nMIMessageSource
Dim nMIMessageType
Dim nMIMessageScope
Dim lMIMessageObjectID
Dim nMIMessageObjectSource
Dim nMIMessagePriority
Dim sMIMessageTrialName
Dim lMIMessagePersonID
Dim lMIMessageVisitId
Dim nMIMessageVisitCycle
Dim lMIMessageCRFPageTaskID
Dim lMIMessageResponseTaskID
Dim sMIMessageResponseValue
Dim lMIMessageOCDiscrepancyID
Dim dMIMessageCreated
Dim dMIMessageSent
Dim dMIMessageReceived
Dim nMIMessageHistory
Dim nMIMessageProcessed
Dim nMIMessageStatus
Dim sMIMessageText
'Dim sMIMessageUserCode
Dim sMIMessageUserName
Dim sMIMessageUserNameFull
Dim dMIMessageResponseTimeStamp
Dim sSQL
Dim bNotaPriorityChangeMessage
Dim bInsertedBefore

' NCJ 15 Jan 03 - New fields added for MACRO 3.0
Dim nMIMessageResponseCycle
Dim nMIMessageCreated_TZ
Dim nMIMessageReceived_TZ
Dim nMIMessageSent_TZ

Dim lMIMessageCRFPageCycle
Dim lMIMessageCRFPageID
Dim lMIMessageDataItemID

'TA 12/03/2003: object for updating MIMsgStatus
Dim MIMsgStatic
dim sCon 'for connection string

'store the passed parameters
lMIMessageID = request.form("ID")
sMIMessageSite = request.form("Site")
nMIMessageSource = request.form("Source")
nMIMessageType = request.form("Type")
nMIMessageScope = request.form("Scope")
lMIMessageObjectID = request.form("ObjectID")
nMIMessageObjectSource = request.form("ObjectSource")
nMIMessagePriority = request.form("Priority")
sMIMessageTrialName = request.form("TrialName")
lMIMessagePersonID = request.form("PersonID")
lMIMessageVisitId = request.form("VisitId")
nMIMessageVisitCycle = request.form("VisitCycle")
lMIMessageCRFPageTaskID = request.form("CRFPageTaskID")
' New for 3.0
lMIMessageCRFPageID = request.form("CRFPageID")
lMIMessageCRFPageCycle = request.form("CRFPageCycle")
lMIMessageDataItemID = request.form("DataItemID")
nMIMessageResponseCycle = request.form("ResponseCycle")
' '''''''''''''''''''
lMIMessageResponseTaskID = request.form("ResponseTaskID")
sMIMessageResponseValue = Replace(request.form("ResponseValue"), "'", "''")
lMIMessageOCDiscrepancyID = request.form("OCDiscrepancyID")
dMIMessageCreated = request.form("Created")
dMIMessageSent = request.form("Sent")
dMIMessageReceived = request.form("Received")
nMIMessageHistory = request.form("History")
nMIMessageProcessed = request.form("Processed")
nMIMessageStatus = request.form("Status")
sMIMessageText = Replace(request.form("Text"), "'", "''")
' DPH 17/1/2002 - Replaced MIMessageUserCode with MIMessageUserNameFull
'sMIMessageUserCode = request.form("UserCode")
' DPH 04/04/2002 - Backwards Compatibility Changes - User Name / Description
If request.form("UserCode") <> "" Then
	sMIMessageUserName = Replace(request.form("UserCode"), "'", "''")
	sMIMessageUserNameFull = Replace(request.form("UserName"), "'", "''")
Else
	sMIMessageUserName = Replace(request.form("UserName"), "'", "''")
	sMIMessageUserNameFull = Replace(request.form("UserNameFull"), "'", "''")
End If
' DPH 04/04/2002 - Backwards Compatibility Changes - ResponseTimeStamp
If request.form("ResponseTimeStamp") <> "" Then
	dMIMessageResponseTimeStamp=request.form("ResponseTimeStamp")
Else
	dMIMessageResponseTimeStamp = 0
End If

nMIMessageCreated_TZ = request.form("Created_TZ")
nMIMessageSent_TZ = request.form("Sent_TZ")

'set the message processed flag to 0
nMIMessageProcessed = 0

' REM 19/02/03 - added Time Zone
nMIMessageReceived_TZ = session("strTimeZone")

'initialise the Priority Change Message flag
bNotaPriorityChangeMessage = 1


'validate input
'----------------------------------------------------------------------------
'validate messageid
if (not fnNumeric(lMIMessageID)) then
	Response.Write("ERROR:The ID '" & lMIMessageID & "' is not valid")
	Response.End 
end if
'validate site
if (not fnValidateSite(sMIMessageSite)) then
	Response.Write("ERROR:The site '" & sMIMessageSite & "' does not exist")
	Response.End 
end if
'validate source
if (not fnNumeric(nMIMessageSource)) then
	Response.Write("ERROR:The Source '" & nMIMessageSource & "' is not valid")
	Response.End 
end if
'validate type
if (not fnNumeric(nMIMessageType)) then
	Response.Write("ERROR:The Type '" & nMIMessageType & "' is not valid")
	Response.End 
end if
'validate scope
if (not fnNumeric(nMIMessageScope)) then
	Response.Write("ERROR:The Scope '" & nMIMessageScope & "' is not valid")
	Response.End 
end if
'validate objectid
if (not fnNumeric(lMIMessageObjectID)) then
	Response.Write("ERROR:The ObjectID '" & lMIMessageObjectID & "' is not valid")
	Response.End 
end if
'validate objectsource
if (not fnNumeric(nMIMessageObjectSource)) then
	Response.Write("ERROR:The ObjectSource '" & nMIMessageObjectSource & "' is not valid")
	Response.End 
end if
'validate priority
if (not fnNumeric(nMIMessagePriority)) then
	Response.Write("ERROR:The Priority '" & nMIMessagePriority & "' is not valid")
	Response.End 
end if
'validate study
if (not fnValidateStudy(sMIMessageTrialName)) then
	Response.Write("ERROR:The study '" & sMIMessageTrialName & "' does not exist")
	Response.End 
end if
'validate personid
if (not fnNumeric(lMIMessagePersonID)) then
	Response.Write("ERROR:The PersonID '" & lMIMessagePersonID & "' is not valid")
	Response.End 
end if
'validate visitid
if (not fnNumeric(lMIMessageVisitId)) then
	Response.Write("ERROR:The VisitID '" & lMIMessageVisitId & "' is not valid")
	Response.End 
end if
'validate visitcycle
if (not fnNumeric(nMIMessageVisitCycle)) then
	Response.Write("ERROR:The VisitCycle '" & nMIMessageVisitCycle & "' is not valid")
	Response.End 
end if
'validate crfpagetaskid
if (not fnNumeric(lMIMessageCRFPageTaskID)) then
	Response.Write("ERROR:The CRFPageTaskID '" & lMIMessageCRFPageTaskID & "' is not valid")
	Response.End 
end if
'validate crfpageid
if (not fnNumeric(lMIMessageCRFPageID)) then
	Response.Write("ERROR:The CRFPageID '" & lMIMessageCRFPageID & "' is not valid")
	Response.End 
end if
'validate crfpagecycle
if (not fnNumeric(lMIMessageCRFPageCycle)) then
	Response.Write("ERROR:The CRFPageCycle '" & lMIMessageCRFPageCycle & "' is not valid")
	Response.End 
end if
'validate dataitemid
if (not fnNumeric(lMIMessageDataItemID)) then
	Response.Write("ERROR:The DatItemID '" & lMIMessageDataItemID & "' is not valid")
	Response.End 
end if
'validate responsecycle
if (not fnNumeric(nMIMessageResponseCycle)) then
	Response.Write("ERROR:The ResponseCycle '" & nMIMessageResponseCycle & "' is not valid")
	Response.End 
end if
'validate responsetaskid
if (not fnNumeric(lMIMessageResponseTaskID)) then
	Response.Write("ERROR:The ResponseTaskID '" & lMIMessageResponseTaskID & "' is not valid")
	Response.End 
end if
'validate ocdiscrepancyid
if (not fnNumeric(lMIMessageOCDiscrepancyID)) then
	Response.Write("ERROR:The OCDiscrepancyID '" & lMIMessageOCDiscrepancyID & "' is not valid")
	Response.End 
end if
'validate created
if (not fnNumeric(dMIMessageCreated)) then
	Response.Write("ERROR:The Created '" & dMIMessageCreated & "' is not valid")
	Response.End 
end if
'validate sent
if (not fnNumeric(dMIMessageSent)) then
	Response.Write("ERROR:The Sent '" & dMIMessageSent & "' is not valid")
	Response.End 
end if
'validate received
if (not fnNumeric(dMIMessageReceived)) then
	Response.Write("ERROR:The Received '" & dMIMessageReceived & "' is not valid")
	Response.End 
end if
'time stamp the message received flag
' NCJ 25/10/01 - Convert local to standard
' dMIMessageReceived = CDbl(Now)
dMIMessageReceived = ConvertLocalNumToStandard(CStr(CDbl(Now)))

'validate history
if (not fnNumeric(nMIMessageHistory)) then
	Response.Write("ERROR:The History '" & nMIMessageHistory & "' is not valid")
	Response.End 
end if
''validate processed
'if (not fnNumeric(nMIMessageProcessed)) then
'	Response.Write("ERROR:The Processed '" & nMIMessageProcessed & "' is not valid")
'	Response.End 
'end if
'validate status
if (not fnNumeric(nMIMessageStatus)) then
	Response.Write("ERROR:The Status '" & nMIMessageStatus & "' is not valid")
	Response.End 
end if
'validate username
if not fnValidateUsername(sMIMessageUserName) then
	Response.Write("ERROR:The UserName '" & sMIMessageUserName & "' is not valid")
	Response.End 
end if
''validate fullusername
'if (not fnAlphaNumeric(sMIMessageUserNameFull) or not fnLengthBetween(sMIMessageUserNameFull, 1, 100)) then
'	Response.Write("ERROR:The UserNameFull '" & sMIMessageUserNameFull & "' is not valid")
'	Response.End 
'end if
'validate responsetimestamp
if (not fnNumeric(dMIMessageResponseTimeStamp)) then
	Response.Write("ERROR:The ResponseTimeStamp '" & dMIMessageResponseTimeStamp & "' is not valid")
	Response.End 
end if
'validate created_tz
if (not fnNumeric(nMIMessageCreated_TZ)) then
	Response.Write("ERROR:The Created_TZ '" & nMIMessageCreated_TZ & "' is not valid")
	Response.End 
end if
'validate sent_tz
if (not fnNumeric(nMIMessageSent_TZ)) then
	Response.Write("ERROR:The Sent_TZ '" & nMIMessageSent_TZ & "' is not valid")
	Response.End 
end if
'-------------------------------------------------------------------------------------


'Priority Change Messages.
'Data Monitors are allowed to change the Priority of a Discrepancy, this
'is done in conjunction with setting its message Sent field back to 0,
'which has the effect of causing the message to be retransmitted.
'The above activity normally happens on a Server, but could happen on a Client
'database so the ability to transfer a priority Change message from Client 
'database to a Server has been placed here.

'Check for a Discrepancy message with status Raised,
'because it might be a priority change message that only requires
'MIMessagePriority and MIMessageReceived to be updated
If nMIMessageType = 0 And nMIMessageStatus = 0 Then
	sSQL = "SELECT MIMessageID FROM MIMessage" _
		& " WHERE MIMessageId = " & lMIMessageID _
                & " AND MIMessageSite ='" & sMIMessageSite & "'" _
                & " AND MIMessageSource = " & nMIMessageSource
	Set rsTemp = MACROCnn.Execute(sSQL)

	if err.number <> 0 then
		response.write "ERROR1:" & err.number
		Response.End
	end if 

	Do While not rsTemp.EOF
                'Update the priority and set the new Received time
		sSQL = "UPDATE MIMessage " _
			& " SET MIMessagePriority = " & nMIMessagePriority & "," _
			& " MIMessageReceived = " & dMIMessageReceived _
			& " WHERE MIMessageId = " & lMIMessageID _
                	& " AND MIMessageSite ='" & sMIMessageSite & "'" _
                	& " AND MIMessageSource = " & nMIMessageSource
		MACROCnn.Execute sSQL

		if err.number <> 0 then
			response.write "ERROR2:" & err.number
			Response.End
		end if 

        'No more processing required for this priority changing message
        bNotaPriorityChangeMessage = 0
		rsTemp.MoveNext
	Loop
	rsTemp.Close
	Set rsTemp = Nothing
End If
            
If bNotaPriorityChangeMessage = 1 Then
	' DPH 20/11/2002 Check if this exact message has been received previously. 
	'	If we are here it is not a PriorityChange Message
	'	If it has do not write to database again as will cause DB Key violation
	sSQL = "SELECT Count(*) AS MICount FROM MIMessage" _
		& " WHERE MIMessageId = " & lMIMessageID _
                & " AND MIMessageSite ='" & sMIMessageSite & "'" _
                & " AND MIMessageSource = " & nMIMessageSource
	Set rsTemp = MACROCnn.Execute(sSQL)

	if err.number <> 0 then
		response.write "ERROR6:" & err.number
		Response.End
	end if 

	If CLng(rsTemp(0).Value) > CLng(0) Then
		bInsertedBefore = true
	Else
		bInsertedBefore = false
	End If
	rsTemp.Close
	set rsTemp = nothing
	
	If Not bInsertedBefore Then
		'Set the history flag on the previous message of a Discrepancies or SDV.
		'Discrepancies and SDVs have a MessageObjectID, Messages and Notes do not
		' DPH 08/11/2002 Make sure MIMessage is a current one before setting
		'	others of the same object ID to be a history mimessage
		If lMIMessageObjectID > 0 And nMIMessageHistory = 0 Then
		            sSQL = "UPDATE MIMessage " _
		                	& " SET MIMessageHistory = 1 " _
				& " WHERE MIMessageSite = '" & sMIMessageSite & "'" _
		                	& " AND MIMessageObjectId = " & lMIMessageObjectID _
		                	& " AND MIMessageObjectSource = " & nMIMessageObjectSource _
		                	& " AND MIMessageHistory = 0"
		            MACROCnn.Execute sSQL

					if err.number <> 0 then
						response.write "ERROR3:" & err.number
						Response.End
					end if 

		End If

		'store the new message
		' Replaced MIMessageUserCode with MIMessageUserNameFull
		sSQL = "INSERT INTO MIMessage (MIMessageID,MIMessageSite,MIMessageSource,MIMessageType," _
			& "MIMessageScope,MIMessageObjectID,MIMessageObjectSource,MIMessagePriority," _
			& "MIMessageTrialName,MIMessagePersonID,MIMessageVisitId,MIMessageVisitCycle," _
		  		& "MIMessageCRFPageTaskID,MIMessageResponseTaskID,MIMessageResponseValue," _
		    	& "MIMessageOCDiscrepancyID,MIMessageCreated,MIMessageSent,MIMessageReceived," _
		    	& "MIMessageHistory,MIMessageProcessed,MIMessageStatus,MIMessageText," _
		    	& "MIMessageUserName,MIMessageUserNameFull,MIMessageResponseTimeStamp, " _
		    	& "MIMessageResponseCycle,MIMessageCreated_TZ,MIMessageReceived_TZ, " _
		    	& "MIMessageSent_TZ,MIMessageCRFPageCycle,MIMessageCRFPageID,MIMessageDataItemId" _
		    	& ") "
		sSQL = sSQL & "VALUES (" & lMIMessageID & ",'" & sMIMessageSite & "'," & nMIMessageSource & "," & nMIMessageType & "," _
		    	& nMIMessageScope & "," & lMIMessageObjectID & "," & nMIMessageObjectSource & "," & nMIMessagePriority & ",'" _
		    	& sMIMessageTrialName & "'," & lMIMessagePersonID & "," & lMIMessageVisitId & "," & nMIMessageVisitCycle & "," _
		    	& lMIMessageCRFPageTaskID & "," & lMIMessageResponseTaskID & ",'" & sMIMessageResponseValue & "'," _
		    	& lMIMessageOCDiscrepancyID & "," & dMIMessageCreated & "," _
		    	& dMIMessageSent & "," & dMIMessageReceived & "," _
		    	& nMIMessageHistory & "," & nMIMessageProcessed & "," & nMIMessageStatus & ",'" _
		    	& sMIMessageText & "','" & sMIMessageUserName & "','" & sMIMessageUserNameFull & "'," & dMIMessageResponseTimeStamp & ","
		 sSQL = sSQL & nMIMessageResponseCycle & "," & nMIMessageCreated_TZ & "," & nMIMessageReceived_TZ & "," _
				& nMIMessageSent_TZ & "," & lMIMessageCRFPageCycle & ","  & lMIMessageCRFPageID & "," & lMIMessageDataItemID _
				& ")"

			MACROCnn.Execute sSQL
			
	
		if err.number <> 0 then
			response.write "ERROR4:" & err.number
			Response.End
		end if 
			
		'TA 12/02/2003: updateMIMsgStatus	- a clinicaltrial id of -1 forces it to be calculated
		' all vaues passed by value to avoid type mismatches
		sCon = Session("strConn")
		Set MIMsgStatic = Createobject("MACROMIMsgBS30.MIMsgStatic")
		If nMIMessageType = 2 then
			'note
			MIMsgStatic.UpdateNoteStatusInDB (sCon),(nMIMessageScope),(sMIMessageTrialName),(-1),(sMIMessageSite),(lMIMessagePersonID), _
												(lMIMessageVisitId),(nMIMessageVisitCycle), (lMIMessageCRFPageTaskID) , (lMIMessageResponseTaskID), (nMIMessageResponseCycle) 
		Else
			'SDV,discrepancy
			MIMsgStatic.UpdateMIMsgStatusInDB (0),(0),(0),(0),(sCon),(nMIMessageType),(sMIMessageTrialName),(-1),(sMIMessageSite),(lMIMessagePersonID), _
												(lMIMessageVisitId),(nMIMessageVisitCycle), (lMIMessageCRFPageTaskID) , (lMIMessageResponseTaskID), (nMIMessageResponseCycle) 
		End if
		set MIMsgStatic=nothing
		
			
		if err.number <> 0 then
			response.write "ERROR4.5:" & err.number
			Response.End
		end if 
	end if
end if

if err.number <> 0 then
	response.write "ERROR5:" & err.number
else
	response.write "SUCCESS"
end if

%>

<!--#include file=CloseDataConnection.txt-->