Attribute VB_Name = "basMessageFunctions"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       basMessageFunctions.bas
'   Author:     Andrew Newbigging
'   Purpose:    Saves messages for remote sites to pick up later
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
 '  1   Andrew Newbigging
'   2   Andrew Newbigging   18/2/99
'       MSMQ not used any more.  Messages are held in the main macro database.
'   3  PN  10/09/99     Upgrade from DAO to ADO and updated code to conform
'                       to VB standards doc version 1.0
'   4  PN  15/09/99     Changed call to ADODBConnection() to MacroADODBConnection()
'   ATN 16/12/99        Changed integers to longs
'   Mo Morris 25/4/00   New subs added CreateLockUnlockFreezeMessage,CreateTrialSubjectLockStatusMessage,
'                       CreateVisitInstanceLockStatusMessage, CreateCRFPageInstanceLockStatusMessage,
'                       CreateDataItemLockStatusMessage, CreateTrialSubjectUnLockMessage,
'                       CreateVisitInstanceUnLockMessage, CreateCRFPageInstanceUnLockMessage,
'                       RemoteSetTrialSubjectLockStatus, RemoteSetVisitInstanceLockStatus,
'                       RemoteSetCRFPageInstanceLockStatus, RemoteSetDataItemLockStatus,
'                       RemoteUnlockTrialSubject, RemoteUnlockVisitInstance, RemoteUnlockCRFPageInstance
'   Mo Morris 4/5/00    SR 3406, adding ClinicalTrialName to Message's parameter field
' WillC SR3534 4/8/00   Changed the word Trial in messages to study.
'   NCJ 29/9/00     Added spaces to "Study" in messages
'   TA 17/6/02: CBB 2.2.16.18 - linking site to a study now puts a message to send a zip file of the study
' ASH 23/08/2002 - Added RepeatNumber to RemoteSetDataItemLockStatus and CreateDataItemLockStatusMessage
' DPH 23/08/2002 - Added Study versioning features
'   NCJ 18 Dec 02 - Added comma in sSQL in CreateStatusMessage
'----------------------------------------------------------------------------------------'

Option Explicit
Option Base 0
Option Compare Binary

'----------------------------------------------------------------------------------------'
Public Sub CreateStatusMessage(vMessageType As Integer, _
                                vClinicalTrialId As Long, _
                                vClinicalTrialName As String, _
                                vTrialSite As String, _
                                Optional sCabFileName As String, _
                                Optional sStudyDescription As String = "")
'----------------------------------------------------------------------------------------'
'   ATN 18/2/99
'   Messages added to main MACRO database now.
'Mo Morris 23/2/00  sCabFileName added as a parameter for us on NewTrial and NewVersion messages
'Mo Morris 16/5/00 SR 3422, Set up Inpreparation status messages
' WillC SR3534 4/8/00   Changed the word Trial in messages to Study.
'   NCJ 29/9/00 - Added spaces too! Also assume vClinicalTrialName has no leading/trailing spaces
'   DPH 20/08/2002 - New Message Field MessageReceivedTimeStamp / optional parameter
'   DPH 03/09/2002 - Changed optional parameter to take version description
'   RS 29/10/2002   -   Add MessageCreatedTimezone column
'   NCJ 18 Dec 02 - Added comma in sSQL for MessageTimestamp_TZ
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim sExistingCabFile As String
Dim lMessageId As Long
Dim sVersionDesc As String
Dim oTimezone As New TimeZone
Dim rsSite As ADODB.Recordset
Dim sSiteSQL As String
Dim sSiteParameters As String

On Error GoTo ErrHandler

    ' PN 10/09/99 - updated code from using dao to use ado
    'changed GetTimeStamp to Format(NOW, "mm/dd/yyyy hh:mm:ss")
    'changed # to ' arround date
    
    ' WillC 4/2/00 Changed to SQLStandardNow from Cdbl(now) to cope with regional settings
    'Changed Mo Morris 26/4/00, MessageReceived enumeration now being used
    'Mo Morris 30/8/01 Db Audit (UserId to UserName, MessageId no longer an autonumber)
    ' RS 29/10/2002: Add Timezone column
    lMessageId = NextMessageId
    sSQL = "INSERT INTO Message (MessageId, TrialSite, ClinicalTrialId, MessageType, " _
        & " MessageTimestamp, MessageTimestamp_TZ, " _
        & " UserName, MessageReceived, MessageDirection, MessageReceivedTimeStamp,MessageBody,MessageParameters) " _
        & " VALUES (" & lMessageId & ",'" & vTrialSite & "'," & vClinicalTrialId & "," & vMessageType & "," & SQLStandardNow & "," & oTimezone.TimezoneOffset & ",'" _
        & goUser.UserName & "'," & MessageReceived.NotYetReceived & "," & MessageDirection.MessageOut & ",0,'"
    
    Select Case vMessageType
    Case ExchangeMessageType.NewTrial
        '   message body
        ' WillC SR3534 4/8/00   Changed Trial in message to study.
        sSQL = sSQL & "Site " & vTrialSite & " has been added to study " & vClinicalTrialName
        'changed Mo Morris 25/2/00
        'check to see wether a distribution cab file exists for this trial (TRIALNAME.zip)
        'if one exists append its name as the message parameter
        'otherwise leave message parameters blank
        
        'TA 17/6/02: CBB 2.2.16.18 - make sure it checks for zip files as well as cab
        sExistingCabFile = Dir(goUser.Database.HTMLLocation & vClinicalTrialName & ".zip")
        If sExistingCabFile = "" Then
            'no zip file - check for cab file (old style study distribution
            sExistingCabFile = Dir(goUser.Database.HTMLLocation & vClinicalTrialName & ".cab")
        End If
        
        'REM 17/01/03 - Add all the site paramaters to MessageParameter field
        sSiteSQL = "SELECT * FROM Site" _
                & " WHERE Site = '" & vTrialSite & "'"
        Set rsSite = New ADODB.Recordset
        rsSite.Open sSiteSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        
        'ASH 20/2/2003 Removed SiteLocale and SiteTimezone
        sSiteParameters = vTrialSite & gsSEPARATOR & ReplaceQuotes(rsSite!SiteDescription) & gsSEPARATOR & rsSite!SiteStatus & gsSEPARATOR & rsSite!SiteLocation & gsSEPARATOR & rsSite!SiteCountry
        
        sSQL = sSQL & "','" & sExistingCabFile & gsMSGSEPARATOR & sSiteParameters & "')"
        
'        If sExistingCabFile = "" Then
'            ' message parameters = ClinicalTrialName
'            sSQL = sSQL & "','')"
'        Else
'            ' message parameters = ClinicalTrialName.cab or zip
'            sSQL = sSQL & "','" & sExistingCabFile & "')"
'        End If
        
    'Mo Morris 16/5/00 SR 3422
    Case ExchangeMessageType.InPreparation
        '   message body
        sSQL = sSQL & "Study " & vClinicalTrialName _
                & " has been set to In preparation."
        ' message parameters = TrialName*TrialStatus
        sSQL = sSQL & "','" & vClinicalTrialName & "*" & vMessageType & "')"
        
    Case ExchangeMessageType.NewVersion
        '   message body
        sVersionDesc = "A new version "
        If sStudyDescription <> "" Then
            sVersionDesc = sVersionDesc & "(" & sStudyDescription & ") "
        End If
        sVersionDesc = sVersionDesc & "of study " & vClinicalTrialName _
                & " has been distributed."
        ' make sure description is not too long
        sVersionDesc = Left(sVersionDesc, 255)
        ' message parameters = ClinicalTrialName
        sSQL = sSQL & sVersionDesc & "','" & sCabFileName & "')"
        
    Case ExchangeMessageType.TrialOpen
        '   message body
        sSQL = sSQL & "Study " & vClinicalTrialName _
                & " has been opened."
        'Changed Mo Morris 4/5/00 SR3406
        ' message parameters = TrialName*TrialStatus
        sSQL = sSQL & "','" & vClinicalTrialName & "*" & vMessageType & "')"
        
    Case ExchangeMessageType.TrialSuspended
        '   message body
        sSQL = sSQL & "Study " & vClinicalTrialName _
                & " has been suspended."
        'Changed Mo Morris 4/5/00 SR3406
        ' message parameters = TrialName*TrialStatus
        sSQL = sSQL & "','" & vClinicalTrialName & "*" & vMessageType & "')"
        
    Case ExchangeMessageType.ClosedRecruitment
        '   message body
        sSQL = sSQL & "Study " & vClinicalTrialName _
                & " has been closed to recruitment."
        'Changed Mo Morris 4/5/00 SR3406
        ' message parameters = TrialName*TrialStatus
        sSQL = sSQL & "','" & vClinicalTrialName & "*" & vMessageType & "')"
        
    Case ExchangeMessageType.ClosedFollowUp
        '   message body
        sSQL = sSQL & "Study " & vClinicalTrialName _
                & " has been closed to follow up."
        'Changed Mo Morris 4/5/00 SR3406
        ' message parameters = TrialName*TrialStatus
        sSQL = sSQL & "','" & vClinicalTrialName & "*" & vMessageType & "')"
        
    End Select
    
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "CreateStatusMessage", "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub


'--------------------------------------------------------------------------------
Public Function GetStudyVersionFromParameterField(sParam As String, sStudyName As String) As Long
'--------------------------------------------------------------------------------
' Extract Studyname from recordset parameter 'X.zip'
'--------------------------------------------------------------------------------
On Error GoTo ErrorHandler

    GetStudyVersionFromParameterField = CLng(Left(sParam, (InStr(1, sParam, sStudyName & ".zip", vbBinaryCompare)) - 1))

Exit Function

ErrorHandler:
    GetStudyVersionFromParameterField = -1
End Function

'--------------------------------------------------------------------------------
Public Function GetStudyNameFromParameterField(sParam As String) As String
'--------------------------------------------------------------------------------
' Extract Studyname from recordset parameter 'XStudyname.zip'
'--------------------------------------------------------------------------------
Dim sStudyName As String
Dim nI As Integer
Dim nLen As Integer

On Error GoTo ErrorHandler

    ' firstly strip to XStudyname
    sStudyName = Mid(sParam, 1, InStr(sParam, ".") - 1)
    ' Remove numerics
    If IsNumeric(Left(sStudyName, 1)) Then
        nLen = Len(sStudyName)
        For nI = 1 To nLen
            If IsNumeric(Left(sStudyName, 1)) Then
                sStudyName = Right(sStudyName, Len(sStudyName) - 1)
            End If
        Next
    End If
    
    GetStudyNameFromParameterField = sStudyName

Exit Function

ErrorHandler:
    GetStudyNameFromParameterField = ""
End Function

'---------------------------------------------------------------------
Public Sub CreateLockUnlockFreezeMessage(ByVal nMessageType As Integer, _
                                ByVal lClinicalTrialId As Long, _
                                ByVal sTrialSite As String, _
                                ByVal sMessageBody As String, _
                                ByVal sMessageParameters As String)
'---------------------------------------------------------------------
Dim sSQL As String
Dim lMessageId As Long
Dim oTimezone As New TimeZone

    'Mo Morris 30/8/01 Db Audit (UserId to UserName, MessageId no longer an autonumber)
    lMessageId = NextMessageId
    sSQL = "INSERT INTO Message (MessageId, TrialSite,ClinicalTrialId,MessageType,MessageTimestamp,MessageTimestamp_TZ,UserName,MessageReceived,MessageDirection,MessageBody,MessageParameters) " _
        & "  VALUES (" & lMessageId & ",'" & sTrialSite & "'," & lClinicalTrialId & "," & nMessageType & "," & SQLStandardNow & "," & oTimezone.TimezoneOffset & ",'" _
        & goUser.UserName & "'," & MessageReceived.NotYetReceived & "," & MessageDirection.MessageOut & ",'" _
        & sMessageBody & "','" & sMessageParameters & "')"
        
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "CreateLockUnlockFreezeMessage", "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub CreateTrialSubjectLockStatusMessage(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal nLockSetting As LockStatus, _
                            ByVal sTimeStamp As String)
'---------------------------------------------------------------------
'Changed Mo Morris 4/5/00 SR3406
'---------------------------------------------------------------------
Dim sMessageBody As String
Dim sMessageParameters As String

    Select Case nLockSetting
    Case LockStatus.lsLocked
        sMessageBody = "Locking Subject: " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
        sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & LockStatus.lsLocked _
            & "*" & nPersonId & "*" & sTimeStamp
    Case LockStatus.lsUnlocked
        sMessageBody = "UnLocking Subject: " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
        sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & LockStatus.lsUnlocked _
            & "*" & nPersonId & "*" & sTimeStamp
    Case LockStatus.lsFrozen
        sMessageBody = "Freezing Subject: " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
        sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & LockStatus.lsFrozen _
            & "*" & nPersonId & "*" & sTimeStamp
    End Select
    
    Call CreateLockUnlockFreezeMessage(ExchangeMessageType.TrialSubjectLockStatus, _
            lClinicalTrialId, sTrialSite, sMessageBody, sMessageParameters)

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "CreateTrialSubjectLockStatusMessage", "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub CreateVisitInstanceLockStatusMessage(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lVisitId As Long, _
                            ByVal nVisitCycleNumber As Integer, _
                            ByVal nLockSetting As LockStatus, _
                            ByVal sTimeStamp As String)
'---------------------------------------------------------------------
'Changed Mo Morris 4/5/00 SR3406
'---------------------------------------------------------------------
Dim sMessageBody As String
Dim sMessageParameters As String

    Select Case nLockSetting
    Case LockStatus.lsLocked
        sMessageBody = "Locking Visit: " & VisitCodeFromId(lClinicalTrialId, lVisitId) _
            & " on subject " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
        sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & LockStatus.lsLocked _
            & "*" & nPersonId & "*" & lVisitId & "*" & nVisitCycleNumber & "*" & sTimeStamp
    Case LockStatus.lsUnlocked
        sMessageBody = "UnLocking Visit: " & VisitCodeFromId(lClinicalTrialId, lVisitId) _
            & " on subject " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
        sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & LockStatus.lsUnlocked _
            & "*" & nPersonId & "*" & lVisitId & "*" & nVisitCycleNumber & "*" & sTimeStamp
    Case LockStatus.lsFrozen
        sMessageBody = "Freezing Visit: " & VisitCodeFromId(lClinicalTrialId, lVisitId) _
            & " on subject " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
        sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & LockStatus.lsFrozen _
            & "*" & nPersonId & "*" & lVisitId & "*" & nVisitCycleNumber & "*" & sTimeStamp
    End Select
    
    Call CreateLockUnlockFreezeMessage(ExchangeMessageType.VisitInstanceLockStatus, _
            lClinicalTrialId, sTrialSite, sMessageBody, sMessageParameters)

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "CreateVisitInstanceLockStatusMessage", "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub CreateCRFPageInstanceLockStatusMessage(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lCRFPageTaskId As Long, _
                            ByVal nLockSetting As LockStatus, _
                            ByVal sTimeStamp As String)
'---------------------------------------------------------------------
'Changed Mo Morris 4/5/00 SR3406
'---------------------------------------------------------------------
Dim sMessageBody As String
Dim sMessageParameters As String

    Select Case nLockSetting
    Case LockStatus.lsLocked
        sMessageBody = "Locking Form: " & CRFPageCodeFromTaskId(lClinicalTrialId, sTrialSite, nPersonId, lCRFPageTaskId) _
            & " on subject " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
        sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & LockStatus.lsLocked _
            & "*" & nPersonId & "*" & lCRFPageTaskId & "*" & sTimeStamp
    Case LockStatus.lsUnlocked
        sMessageBody = "UnLocking Form: " & CRFPageCodeFromTaskId(lClinicalTrialId, sTrialSite, nPersonId, lCRFPageTaskId) _
            & " on subject " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
        sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & LockStatus.lsUnlocked _
            & "*" & nPersonId & "*" & lCRFPageTaskId & "*" & sTimeStamp
    Case LockStatus.lsFrozen
        sMessageBody = "Freezing Form: " & CRFPageCodeFromTaskId(lClinicalTrialId, sTrialSite, nPersonId, lCRFPageTaskId) _
            & " on subject " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
        sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & LockStatus.lsFrozen _
            & "*" & nPersonId & "*" & lCRFPageTaskId & "*" & sTimeStamp
    End Select
    
    Call CreateLockUnlockFreezeMessage(ExchangeMessageType.CRFPageInstanceLockStatus, _
            lClinicalTrialId, sTrialSite, sMessageBody, sMessageParameters)

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "CreateCRFPageInstanceLockStatusMessage", "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub CreateDataItemLockStatusMessage(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lResponseTaskId As Long, _
                            ByVal nRepeatNumber As Integer, _
                            ByVal nLockSetting As LockStatus, _
                            ByVal sTimeStamp As String)
'---------------------------------------------------------------------
'Changed Mo Morris 4/5/00 SR3406
'---------------------------------------------------------------------
Dim sMessageBody As String
Dim sMessageParameters As String

    Select Case nLockSetting
    Case LockStatus.lsLocked
        sMessageBody = "Locking Question: " & DataItemCodeFromTaskId(lClinicalTrialId, sTrialSite, nPersonId, lResponseTaskId, nRepeatNumber) _
            & " on subject " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
        sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & LockStatus.lsLocked _
            & "*" & nPersonId & "*" & lResponseTaskId & "*" & nRepeatNumber & "*" & sTimeStamp
    Case LockStatus.lsUnlocked
        sMessageBody = "UnLocking Question: " & DataItemCodeFromTaskId(lClinicalTrialId, sTrialSite, nPersonId, lResponseTaskId, nRepeatNumber) _
            & " on subject " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
        sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & LockStatus.lsUnlocked _
            & "*" & nPersonId & "*" & lResponseTaskId & "*" & nRepeatNumber & "*" & sTimeStamp
    Case LockStatus.lsFrozen
        sMessageBody = "Freezing Question: " & DataItemCodeFromTaskId(lClinicalTrialId, sTrialSite, nPersonId, lResponseTaskId, nRepeatNumber) _
            & " on subject " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
        sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & LockStatus.lsFrozen _
            & "*" & nPersonId & "*" & lResponseTaskId & "*" & nRepeatNumber & "*" & sTimeStamp
    End Select
    
    Call CreateLockUnlockFreezeMessage(ExchangeMessageType.DataItemLockStatus, _
            lClinicalTrialId, sTrialSite, sMessageBody, sMessageParameters)

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "CreateDataItemLockStatusMessage", "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub CreateTrialSubjectUnLockMessage(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer)
'---------------------------------------------------------------------
'Changed Mo Morris 4/5/00 SR3406
'---------------------------------------------------------------------
Dim sMessageBody As String
Dim sMessageParameters As String

    sMessageBody = "UnLocking Subject : " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
    sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & nPersonId
    
    Call CreateLockUnlockFreezeMessage(ExchangeMessageType.TrialSubjectUnLock, _
            lClinicalTrialId, sTrialSite, sMessageBody, sMessageParameters)

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "CreateTrialSubjectUnLockMessage", "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub CreateVisitInstanceUnLockMessage(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lVisitId As Long, _
                            ByVal nVisitCycleNumber As Integer)
'---------------------------------------------------------------------
'Changed Mo Morris 4/5/00 SR3406
'---------------------------------------------------------------------
Dim sMessageBody As String
Dim sMessageParameters As String

    sMessageBody = "UnLocking Visit : " & VisitCodeFromId(lClinicalTrialId, lVisitId) _
            & " on subject " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
    sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & nPersonId & "*" & lVisitId & "*" & nVisitCycleNumber
    
    Call CreateLockUnlockFreezeMessage(ExchangeMessageType.VisitInstanceUnLock, _
            lClinicalTrialId, sTrialSite, sMessageBody, sMessageParameters)

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "CreateVisitInstanceUnLockMessage", "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub CreateCRFPageInstanceUnLockMessage(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lCRFPageTaskId As Long, _
                            ByVal lVisitId As Long, _
                            ByVal nVisitCycleNumber As Integer)
'---------------------------------------------------------------------
'Changed Mo Morris 4/5/00 SR3406
'---------------------------------------------------------------------
Dim sMessageBody As String
Dim sMessageParameters As String

    sMessageBody = "UnLocking Form : " & CRFPageCodeFromTaskId(lClinicalTrialId, sTrialSite, nPersonId, lCRFPageTaskId) _
            & " on subject " & nPersonId & " in " & TrialNameFromId(lClinicalTrialId)
    sMessageParameters = TrialNameFromId(lClinicalTrialId) & "*" & nPersonId & "*" & lCRFPageTaskId _
            & "*" & lVisitId & "*" & nVisitCycleNumber
    
    Call CreateLockUnlockFreezeMessage(ExchangeMessageType.CRFPageInstanceUnLock, _
            lClinicalTrialId, sTrialSite, sMessageBody, sMessageParameters)

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "CreateCRFPageInstanceUnLockMessage", "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub


'------------------------------------------------------------
Public Sub RemoteSetTrialSubjectLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal nLockSetting As LockStatus, _
                            ByVal sTimeStamp As String)
'------------------------------------------------------------
'Similar to SetTrialSubjectLockStatus which performs lock changes on a server database
'Lock, Unlock or Freeze a trial subject
'i.e. apply setting to subject and to all its visits, forms and data items that are not frozen
'Don't change items that already have this status
'Don't set the changed flag, because we do not want the data to be transmitted back to the server
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
    
    'Ignore the LockStatus change message if the Trialsubject contains
    'changed data that has not been exported yet
    sSQL = "SELECT Changed FROM TrialSubject" & _
        " WHERE ClinicalTrialId = " & lClinicalTrialId & _
        " AND TrialSite = '" & sTrialSite & "'" & _
        " AND PersonId = " & nPersonId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTemp.RecordCount = 0 Then
        rsTemp.Close
        Set rsTemp = Nothing
        Exit Sub
    End If
    If rsTemp!Changed = Changed.Changed Then
        rsTemp.Close
        Set rsTemp = Nothing
        Exit Sub
    End If

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus and the Changed flag
    sSQLSetLockWhere = " LockStatus = " & nLockSetting & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND LockStatus <> " & LockStatus.lsFrozen & _
            " AND LockStatus <> " & nLockSetting

    ' Set lock on Trial Subject
    sSQL = "UPDATE TrialSubject SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL
    
    ' Set lock on all the visit instances
    sSQL = "UPDATE VisitInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL
    
    ' Set lock on all the CRF Pages
    sSQL = "UPDATE CRFPageInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL

    ' Set lock on all the Data Items
    'Mo Morris 30/8/01 Db Audit (UserId to UserName)
    sSQL = "UPDATE DataItemResponse SET " & _
           " ResponseTimestamp = " & sTimeStamp & ", " & _
           " UserName = '" & goUser.UserName & "', "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL

    ' Now transfer the stuff into DataItemResponseHistory
    sSQL = SQLToUpdateDataItemResponseHistory
    ' Pick off the items we've just created, matching on the TimeStamp
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND ResponseTimeStamp = " & sTimeStamp
    
    MacroADODBConnection.Execute sSQL
            
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, _
                                      Err.Description, _
                                      "RemoteSetTrialSubjectLockStatus", _
                                      "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------------------------
Public Sub RemoteSetVisitInstanceLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lVisitId As Long, _
                            ByVal nVisitCycleNumber As Integer, _
                            ByVal nLockSetting As LockStatus, _
                            ByVal sTimeStamp As String)
'------------------------------------------------------------
'Similar to SetVisitInstanceLockStatus which performs lock changes on a server database
'Lock, Unlock or Freeze a visit instance
'i.e. apply setting to visit and to all its forms and data items that are not frozen
'Don't change items that already have this status
'Don't set the changed flag, because we do not want the data to be transmitted back to the server
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
    
    'Ignore the LockStatus change message if the VisitInstance contains
    'changed data that has not been exported yet
    sSQL = "SELECT Changed FROM VisitInstance" & _
        " WHERE ClinicalTrialId = " & lClinicalTrialId & _
        " AND TrialSite = '" & sTrialSite & "'" & _
        " AND PersonId = " & nPersonId & _
        " AND VisitId = " & lVisitId & _
        " AND VisitCycleNumber = " & nVisitCycleNumber
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTemp.RecordCount = 0 Then
        rsTemp.Close
        Set rsTemp = Nothing
        Exit Sub
    End If
    If rsTemp!Changed = Changed.Changed Then
        rsTemp.Close
        Set rsTemp = Nothing
        Exit Sub
    End If

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus and the Changed flag
    sSQLSetLockWhere = " LockStatus = " & nLockSetting & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND VisitId = " & lVisitId & _
            " AND VisitCycleNumber = " & nVisitCycleNumber & _
            " AND LockStatus <> " & LockStatus.lsFrozen & _
            " AND LockStatus <> " & nLockSetting

    ' Lock the visit instance
    sSQL = "UPDATE VisitInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL
    
    ' Lock all the CRF Pages
    sSQL = "UPDATE CRFPageInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL

    ' Lock all the Data Items
    'Mo Morris 30/8/01 Db Audit (UserId to UserName)
    sSQL = "UPDATE DataItemResponse SET " & _
           " ResponseTimestamp = " & sTimeStamp & ", " & _
           " UserName = '" & goUser.UserName & "', "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL

    ' Now transfer the stuff into DataItemResponseHistory
    sSQL = SQLToUpdateDataItemResponseHistory
    ' Pick off the items we've just created, matching on the TimeStamp
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND ResponseTimeStamp = " & sTimeStamp
    
    MacroADODBConnection.Execute sSQL
            
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, _
                                      Err.Description, _
                                      "RemoteSetVisitInstanceLockStatus", _
                                      "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------------------------
Public Sub RemoteSetCRFPageInstanceLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lCRFPageTaskId As Long, _
                            ByVal nLockSetting As LockStatus, _
                            ByVal sTimeStamp As String)
'------------------------------------------------------------
'Similar to SetCRFPageInstanceLockStatus which performs lock changes on a server database
'Lock, Unlock or Freeze a CRF page
'i.e. apply setting to page and to all its data items that are not frozen
'Don't change items that already have this status
'Don't set the changed flag, because we do not want the data to be transmitted back to the server
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
    
    'Ignore the LockStatus change message if the CRFPageInstance contains
    'changed data that has not been exported yet
    sSQL = "SELECT Changed FROM CRFPageInstance" & _
        " WHERE ClinicalTrialId = " & lClinicalTrialId & _
        " AND TrialSite = '" & sTrialSite & "'" & _
        " AND PersonId = " & nPersonId & _
        " AND CRFPageTaskId = " & lCRFPageTaskId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTemp.RecordCount = 0 Then
        rsTemp.Close
        Set rsTemp = Nothing
        Exit Sub
    End If
    If rsTemp!Changed = Changed.Changed Then
        rsTemp.Close
        Set rsTemp = Nothing
        Exit Sub
    End If

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus and the Changed flag
    sSQLSetLockWhere = " LockStatus = " & nLockSetting & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND CRFPageTaskId = " & lCRFPageTaskId & _
            " AND LockStatus <> " & LockStatus.lsFrozen & _
            " AND LockStatus <> " & nLockSetting

    ' Lock the CRF Page
    sSQL = "UPDATE CRFPageInstance SET"
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL

    ' Lock all the Data Items
    'Mo Morris 30/8/01 Db Audit (UserId to UserName)
    sSQL = "UPDATE DataItemResponse SET " & _
           " ResponseTimestamp = " & sTimeStamp & ", " & _
           " UserName = '" & goUser.UserName & "', "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL

    ' Now transfer the stuff into DataItemResponseHistory
    sSQL = SQLToUpdateDataItemResponseHistory
    ' Pick off the items we've just created, matching on the TimeStamp
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND ResponseTimeStamp = " & sTimeStamp
    
    MacroADODBConnection.Execute sSQL
            
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, _
                                      Err.Description, _
                                      "RemoteSetCRFPageInstanceLockStatus", _
                                      "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------------------------
Public Sub RemoteSetDataItemLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lResponseTaskId As Long, _
                            ByVal nRepeatNumber As Integer, _
                            ByVal nLockSetting As LockStatus, _
                            ByVal sTimeStamp As String)
'------------------------------------------------------------
'Similar to SetDataItemLockStatus which performs lock changes on a server database
'Lock, Freeze or Unlock a data item
'Don't change items that already have this status
'Don't set the changed flag, because we do not want the data to be transmitted back to the server
'------------------------------------------------------------
Dim sSQL  As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
    
    'Ignore the LockStatus change message if the DataItemResponse contains
    'changed data that has not been exported yet
    sSQL = "SELECT Changed FROM DataItemResponse" & _
        " WHERE ClinicalTrialId = " & lClinicalTrialId & _
        " AND TrialSite = '" & sTrialSite & "'" & _
        " AND PersonId = " & nPersonId & _
        " AND ResponseTaskId = " & lResponseTaskId & _
        " AND RepeatNumber = " & nRepeatNumber
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTemp.RecordCount = 0 Then
        rsTemp.Close
        Set rsTemp = Nothing
        Exit Sub
    End If
    If rsTemp!Changed = Changed.Changed Then
        rsTemp.Close
        Set rsTemp = Nothing
        Exit Sub
    End If

    ' Lock the Data Item - set the LockStatus and the Changed flag
    'Mo Morris 30/8/01 Db Audit (UserId to UserName)
    sSQL = "UPDATE DataItemResponse " & _
           " SET ResponseTimestamp = " & sTimeStamp & ", " & _
           " UserName = '" & goUser.UserName & "', " & _
           " LockStatus = " & nLockSetting
    sSQL = sSQL & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND ResponseTaskId = " & lResponseTaskId & _
            " AND RepeatNumber = " & nRepeatNumber & _
            " AND LockStatus <> " & LockStatus.lsFrozen & _
            " AND LockStatus <> " & nLockSetting

    MacroADODBConnection.Execute sSQL

    ' Now transfer the stuff into DataItemResponseHistory
    sSQL = SQLToUpdateDataItemResponseHistory
    ' Pick off the item we've just created, matching on the TimeStamp
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND ResponseTaskId = " & lResponseTaskId & _
            " AND RepeatNumber = " & nRepeatNumber & _
            " AND ResponseTimeStamp = " & sTimeStamp
    
    MacroADODBConnection.Execute sSQL
            
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, _
                                      Err.Description, _
                                      "RemoteSetDataItemLockStatus", _
                                      "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------------------------
Public Sub RemoteUnlockTrialSubject(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer)
'------------------------------------------------------------
'Similar to UnlockTrialSubject which performs lock changes on a server database
'Unlock a trial subject
'Apply setting only to subject if not frozen and not already unlocked
'Don't set the changed flag, because we do not want the data to be transmitted back to the server
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrHandler

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus, but only if it is locked (not frozen or unlocked)
    sSQLSetLockWhere = " LockStatus = " & LockStatus.lsUnlocked & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND LockStatus = " & LockStatus.lsLocked

    sSQL = "UPDATE TrialSubject SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, _
                                      Err.Description, _
                                      "RemoteUnlockTrialSubject", _
                                      "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------------------------
Public Sub RemoteUnlockVisitInstance(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lVisitId As Long, _
                            ByVal nVisitCycleNumber As Integer)
'------------------------------------------------------------
'Similar to UnlockVisitInstance which performs lock changes on a server database
'Unlock a Visit Instance if not frozen and not already unlocked
'and ALSO unlock the Trial Subject
'Don't set the changed flag, because we do not want the data to be transmitted back to the server
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrHandler

    ' Set the LockStatus, but only if it is locked (not frozen or unlocked)
    sSQLSetLockWhere = " LockStatus = " & LockStatus.lsUnlocked & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND VisitId = " & lVisitId & _
            " AND VisitCycleNumber = " & nVisitCycleNumber & _
            " AND LockStatus = " & LockStatus.lsLocked

    ' Lock the visit instance
    sSQL = "UPDATE VisitInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL

    ' Make sure the Trial Subject is also unlocked
    Call RemoteUnlockTrialSubject(lClinicalTrialId, sTrialSite, nPersonId)
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, _
                                      Err.Description, _
                                      "RemoteUnlockVisitInstance", _
                                      "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------------------------
Public Sub RemoteUnlockCRFPageInstance(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lCRFPageTaskId As Long, _
                            ByVal lVisitId As Long, _
                            ByVal nVisitCycleNumber As Integer)
'------------------------------------------------------------
'Similar to UnlockCRFPageInstance which performs lock changes on a server database
'Unlock a CRF Page Instance if not frozen and not already unlocked
'and ALSO unlock its Visit and Trial Subject
'Don't set the changed flag, because we do not want the data to be transmitted back to the server
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrHandler

    ' Set the LockStatus and the Changed flag, but only if it is locked (not frozen or unlocked)
    sSQLSetLockWhere = " LockStatus = " & LockStatus.lsUnlocked & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND CRFPageTaskId = " & lCRFPageTaskId & _
            " AND LockStatus <> " & LockStatus.lsLocked

    ' Lock the CRF Page
    sSQL = "UPDATE CRFPageInstance SET"
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL

    ' Make sure the Visit and Trial subject are also unlocked
    Call RemoteUnlockVisitInstance(lClinicalTrialId, sTrialSite, nPersonId, _
                            lVisitId, nVisitCycleNumber)

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, _
                                      Err.Description, _
                                      "RemoteUnlockCRFPageInstance", _
                                      "MessageFunctions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub


