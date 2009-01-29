VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LockFreeze"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'----------------------------------------------------------------------------------------'
'   File:       LockFreeze.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Nicky Johns, December 2002
'   Purpose:    Class to carry out Lock/Freeze operations for a subject in MACRO 3.0
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 4-19 Dec 02 - Initial development
'   NCJ 24 Dec 02 - Added CanLockFreezeOnServer
'   NCJ 7 Jan 02 - Added ActionText, ScopeText, StatusText
'   NCJ 17 Jan 03 - Added Web versions of CanLockFreezeOnServer and DoLockFreeze which take connection STRINGs
'   DPH 17/01/2003 Need adOpenKeyset to use RecordCount correctly in CanLockFreezeOnServer
'   DPH 29/05/2003 - Use client cursor in CanLockFreezeOnServerWWW
'----------------------------------------------------------------------------------------'

Option Explicit

' The scope of a Lock/Freeze operation
Public Enum LFScope
    lfscStudy = 0
    lfscSubject = 1
    lfscVisit = 2
    lfscEForm = 3
    lfscQuestion = 4
End Enum

' The types of Lock/Freeze operation
Public Enum LFAction
    lfaUnlock = 0
    lfaLock = 1
    lfaFreeze = 2
    lfaUnfreeze = 3
    lfaRollback = 4
End Enum

' The "processed status" of an LF Message
Public Enum LFProcessStatus
    lfpUnProcessed = 0
    lfpProcessed = 1
    lfpRefused = 2
End Enum

'----------------------------------------------------------------------------------------'
Public Function DoLockFreeze(oDBCon As ADODB.Connection, oLFObj As LFObject, enAction As LFAction, _
                nSource As Integer, sUserName As String, sUserNameFull As String) As LFMessage
'----------------------------------------------------------------------------------------'
' Perform a Lock/Freeze operation
' and return the LFMessage representing it
'----------------------------------------------------------------------------------------'
Dim oLFMsg As LFMessage

    On Error GoTo ErrHandler
    
    ' Create the LF Message
    Set oLFMsg = New LFMessage
    Call oLFMsg.Init(oDBCon, oLFObj, enAction, _
                            nSource, sUserName, sUserNameFull)
    
    Call oLFMsg.DoAction(oDBCon, False)
    
    ' Now say we've done it and save the message in the LFMessage table
    Call oLFMsg.SetAsProcessed
    Call oLFMsg.Save(oDBCon)
    
    Set DoLockFreeze = oLFMsg
    Set oLFMsg = Nothing
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "LockFreeze.DoLockFreeze"

End Function

'----------------------------------------------------------------------------------------'
Public Function DoLockFreezeWWW(sConnection As String, oLFObj As LFObject, enAction As LFAction, _
                nSource As Integer, sUserName As String, sUserNameFull As String) As LFMessage
'----------------------------------------------------------------------------------------'
' Perform a Lock/Freeze operation
' and return the LFMessage representing it
' WEB version with connection string as parameter
'----------------------------------------------------------------------------------------'
' REVISIONS
' DPH 16/09/2003 - Use client cursor
'------------------------------------------------------------------------------
Dim oDBCon As ADODB.Connection

    On Error GoTo ErrHandler
    
    Set oDBCon = New ADODB.Connection
    Call oDBCon.Open(sConnection)
    oDBCon.CursorLocation = adUseClient
    
    Set DoLockFreezeWWW = DoLockFreeze(oDBCon, oLFObj, enAction, nSource, sUserName, sUserNameFull)
    
    Set oDBCon = Nothing
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "LockFreeze.DoLockFreezeWWW"

End Function

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'
    
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'


End Sub

'-------------------------------------------------------------------------------------
Public Function GetMessagesToTransfer(oDBCon As ADODB.Connection, _
                        nSource As Integer, sSite As String) As Collection
'-------------------------------------------------------------------------------------
' Get unsent messages as collection of LFMessage objects
'-------------------------------------------------------------------------------------

    Set GetMessagesToTransfer = GetLFMessagesToTransfer(oDBCon, nSource, sSite)

End Function

'-------------------------------------------------------------------------------------
Public Function RefuseMessage(oDBCon As ADODB.Connection, oLFMsg As LFMessage) As LFMessage
'-------------------------------------------------------------------------------------
' Refuse this message, i.e. Create a rollback message for this one
' and return the (saved) Rollback message created
' Assume this oLFMsg is NOT of type Rollback!
'-------------------------------------------------------------------------------------
Dim oRBMsg As LFMessage
Dim oTimeZone As TimeZone
Dim sTimestamp As String

    On Error GoTo ErrHandler
    
    Set oRBMsg = New LFMessage
    Set oTimeZone = New TimeZone
    sTimestamp = IMedNow
    
    ' Set up with same details as message being rolled back
    ' Assume we're always creating rollbacks at sites
    Call oRBMsg.Init(oDBCon, oLFMsg.MessageObject, lfaRollback, TypeOfInstallation.RemoteSite, _
                oLFMsg.UserName, oLFMsg.UserNameFull)
    ' Set the study name since we know it
    oRBMsg.StudyName = oLFMsg.StudyName
    ' Set the rollback details
    oRBMsg.RollbackMsgId = oLFMsg.MessageId
    oRBMsg.RollbackSource = oLFMsg.MsgSource
    ' Say we've processed it here (it will be set to unprocessed when it arrives on the Server)
    oRBMsg.ProcessedStatus = lfpProcessed
    oRBMsg.MsgCreatedTimestamp = sTimestamp
    oRBMsg.MsgCreatedTimestamp_TZ = oTimeZone.TimezoneOffset
    oRBMsg.ProcessedTimestamp = sTimestamp
    oRBMsg.ProcessedTimestamp_TZ = oTimeZone.TimezoneOffset
    
    Call oRBMsg.Save(oDBCon)
    
    ' Now set the original to Refused
    Call oLFMsg.Refuse(oDBCon)
    
    ' Return the Rollback message just created
    Set RefuseMessage = oRBMsg
    
    Set oRBMsg = Nothing
    Set oTimeZone = Nothing
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|LockFreeze.RefuseMessage"

End Function

'------------------------------------------------------------------------------
Public Sub ProcessSubjectRollBacks(oDBCon As ADODB.Connection, _
                sStudyName As String, sSite As String, lSubjectId As Long)
'------------------------------------------------------------------------------
' Get all the messages to be rolled back for a particular subject
' and execute the rollbacks
' This is called from AutoImport to hoover up any unprocessed Rollbacks
' before processing other pending LF messages
' Assumes a Subject Database lock in place!
'------------------------------------------------------------------------------
Dim oLFRBMsg As LFMessage

    On Error GoTo ErrHandler
    
    For Each oLFRBMsg In GetSubjectRollbacks(oDBCon, sStudyName, sSite, lSubjectId)
        Call HandleRollBack(oDBCon, oLFRBMsg)
    Next

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|LockFreeze.ProcessSubjectRollBacks"

End Sub

'------------------------------------------------------------------------------
Public Sub HandleRollBack(oDBCon As ADODB.Connection, oLFRBMsg As LFMessage)
'------------------------------------------------------------------------------
' Assume oLFRBMsg is a message of type "Rollback"
' Get the message to be rolled back, and execute the rollback
'------------------------------------------------------------------------------
Dim oLFMsgToRollBack As LFMessage

    On Error GoTo ErrHandler
    
    Set oLFMsgToRollBack = New LFMessage
    ' Get the message to rollback
    Call oLFMsgToRollBack.Load(oDBCon, oLFRBMsg.Site, oLFRBMsg.RollbackSource, oLFRBMsg.RollbackMsgId)
    Call oLFMsgToRollBack.DoRollback(oDBCon)
    ' Set the Rollback message to Processed
    Call oLFRBMsg.ChangeProcessStatus(oDBCon, LFProcessStatus.lfpProcessed)

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|LockFreeze.HandleRollBack"

End Sub

'------------------------------------------------------------------------------
Public Function CanLockFreezeOnServer(oDBCon As ADODB.Connection, _
                        sStudyName As String, sSite As String, lSubjectId As Long) As Boolean
'------------------------------------------------------------------------------
' Returns TRUE if there is no subject data waiting for import
' and no unprocessed Site LF messages for subject
' Only relevant for lock/freeze operations on the Server
'------------------------------------------------------------------------------
' REVISIONS
' DPH 17/01/2003 Need adOpenKeyset to use RecordCount correctly
'------------------------------------------------------------------------------
Dim sSQL As String
Dim sSQLWhere As String
Dim rsTemp As ADODB.Recordset
Dim bCanLockFreeze As Boolean

    On Error GoTo ErrHandler
    
    bCanLockFreeze = True
    
    Set rsTemp = New ADODB.Recordset
    
    sSQLWhere = " WHERE ClinicalTrialName = '" & sStudyName & "'" _
        & " AND TrialSite = '" & sSite & "'" _
        & " AND PersonId = " & lSubjectId
    
    ' First check for unprocessed site LF Messages
    sSQL = "SELECT * FROM LFMessage" & sSQLWhere _
        & " AND Source = " & TypeOfInstallation.RemoteSite _
        & " AND ProcessedStatus = " & LFProcessStatus.lfpUnProcessed
    rsTemp.Open sSQL, oDBCon, adOpenKeyset, adLockReadOnly, adCmdText
    bCanLockFreeze = (rsTemp.RecordCount = 0)
    rsTemp.Close
    
    ' If there weren't any, check for unimported data
    If bCanLockFreeze Then
        sSQL = "SELECT * FROM DataImport" & sSQLWhere
        rsTemp.Open sSQL, oDBCon, adOpenKeyset, adLockReadOnly, adCmdText
        bCanLockFreeze = (rsTemp.RecordCount = 0)
        rsTemp.Close
    End If
    
    Set rsTemp = Nothing

    CanLockFreezeOnServer = bCanLockFreeze
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|LockFreeze.CanLockFreezeOnServer"

End Function

'------------------------------------------------------------------------------
Public Function CanLockFreezeOnServerWWW(sConnection As String, _
                        sStudyName As String, sSite As String, lSubjectId As Long) As Boolean
'------------------------------------------------------------------------------
' Returns TRUE if there is no subject data waiting for import
' and no unprocessed Site LF messages for subject
' Only relevant for lock/freeze operations on the Server
' This is the Web version which takes connection STRING
'------------------------------------------------------------------------------
' REVISIONS
' DPH 29/05/2003 - Use client cursor
'------------------------------------------------------------------------------

Dim oDBCon As ADODB.Connection

    On Error GoTo ErrHandler
    
    Set oDBCon = New ADODB.Connection
    Call oDBCon.Open(sConnection)
    oDBCon.CursorLocation = adUseClient
    
    CanLockFreezeOnServerWWW = CanLockFreezeOnServer(oDBCon, sStudyName, sSite, lSubjectId)
    
    Set oDBCon = Nothing
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "LockFreeze.CanLockFreezeOnServerWWW"

End Function

'------------------------------------------------------------------------------
Public Sub DealWithUnprocessedMessages(oDBCon As ADODB.Connection, _
            sStudyName As String, sSite As String, lSubjectId As Long, lLastLFMessageId As Long)
'------------------------------------------------------------------------------
' Deal with unprocessed messages on the Server after an AutoImport,
' processing Site messages for this subject where MessageId <= lLastLFMessageId
' Assumes a subject DB lock is in place
'------------------------------------------------------------------------------
Dim colMsgs As Collection
Dim oLFMsg As LFMessage

    On Error GoTo ErrHandler
    
    Set colMsgs = GetUnprocessedSiteMessages(oDBCon, sStudyName, sSite, lSubjectId, lLastLFMessageId)
    
    For Each oLFMsg In colMsgs
        ' Execute the action and save as Processed
        Call oLFMsg.DoAction(oDBCon, True)
    Next
    
    Set colMsgs = Nothing

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|LockFreeze.DealWithUnprocessedMessages"

End Sub

'------------------------------------------------------------------------------
Public Function LastUsedLFMessageId(oDBCon As ADODB.Connection, _
            lStudyId As Long, sSite As String, lSubjectId As Long) As Long
'------------------------------------------------------------------------------
' Return the last used Site LF Message Id for this subject (excluding Rollbacks)
' Returns 0 if no LF messages for this subject
'------------------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT MAX(MessageID) as MaxMessageId FROM LFMessage " _
        & " WHERE ClinicalTrialId = " & lStudyId _
        & " AND TrialSite = '" & sSite & "'" _
        & " AND PersonId = " & lSubjectId _
        & " AND Source = " & TypeOfInstallation.RemoteSite _
        & " AND MsgType <> " & LFAction.lfaRollback
        
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, oDBCon, adOpenStatic, adLockReadOnly, adCmdText
    
    If IsNull(rsTemp!MaxMessageId) Then     ' No LF messages for this subject
        LastUsedLFMessageId = 0
    Else
        LastUsedLFMessageId = rsTemp!MaxMessageId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|LockFreeze.LastUsedLFMessageId"

End Function

'------------------------------------------------------------------------------
Public Property Get ProcessableMessages(oDBCon As ADODB.Connection, colIgnoreSubjects As Collection) As Collection
'------------------------------------------------------------------------------
' NCJ 14 Jan 03 - Get the "processable" site messages that aren't for subjects in colIgnoreSubjects
' colIgnoreSubjects is a collection with items keyed by "ClinicalTrialName|TrialSite|PersonId"
'------------------------------------------------------------------------------

    Set ProcessableMessages = GetProcessableLFMsgs(oDBCon, colIgnoreSubjects)

End Property

'------------------------------------------------------------------------------
Public Property Get RollbackMessages(oDBCon As ADODB.Connection) As Collection
'------------------------------------------------------------------------------
' NCJ 14 Jan 03 - Get all the unprocessed rollback messages from Sites
'------------------------------------------------------------------------------

    Set RollbackMessages = GetAllRollbacks(oDBCon)

End Property

'------------------------------------------------------------------------------
Public Function ActionText(enAction As LFAction) As String
'------------------------------------------------------------------------------
' Return appropriate text for given LF action
'------------------------------------------------------------------------------

    ActionText = GetActionText(enAction)

End Function

'------------------------------------------------------------------------------
Public Function ScopeText(enScope As LFScope) As String
'------------------------------------------------------------------------------
' Return appropriate text for given LF Scope
'------------------------------------------------------------------------------

    ScopeText = GetScopeText(enScope)

End Function

'------------------------------------------------------------------------------
Public Function StatusText(enStatus As LFProcessStatus) As String
'------------------------------------------------------------------------------
' Return appropriate text for given LF ProcessStatus
'------------------------------------------------------------------------------

    StatusText = GetProcessStatusText(enStatus)
    
End Function

