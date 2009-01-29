VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LFMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   File:       LFMessage.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Nicky Johns, December 2002
'   Purpose:    Represents a Lock/Freeze message as stored in MACRO 3.0 LFMessage table
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 6-12 Dec 02 - Initial development
'   NCJ 3 Jan 03 - Added QuestionId
'   NCJ 14 Feb 03 - Make sure message saving works for regional settings
'   NCJ 13 Apr 04 - Convert timestamps to Standard in SetAsSent and ChangeProcessStatus
'----------------------------------------------------------------------------------------'

Option Explicit

'***********************************************************************
' The repository of values reflecting data in the LFMessage table
Private msStudyName As String

Public MsgSource As Integer     ' Can't use TypeOfInstallation enum here
Public MessageId As Long
Public ActionType As LFAction
'Public Scope As LFScope
Public ProcessedStatus As LFProcessStatus
Public RollbackSource As Integer     ' Can't use TypeOfInstallation enum here
Public RollbackMsgId As Long

Public UserName As String
Public UserNameFull As String

' Multiple Timestamps
Public MsgCreatedTimestamp As Double
Public MsgCreatedTimestamp_TZ As Integer
Public ProcessedTimestamp As Double
Public ProcessedTimestamp_TZ As Integer
Public SentTimestamp As Double
Public SentTimestamp_TZ As Integer
Public ReceivedTimestamp As Double
Public ReceivedTimestamp_TZ As Integer

Public SequenceNo As Long
'*****************************************************************************

' The object to which this message refers
Private moLFObject As LFObject

'----------------------------------------------------------------------------------------'
Public Sub Init(oDBCon As ADODB.Connection, oLFObj As LFObject, enAction As LFAction, _
                nSource As Integer, sUserName As String, sUserNameFull As String)
'----------------------------------------------------------------------------------------'
' Initialise LFMessage for the specified LF Object
'----------------------------------------------------------------------------------------'
Dim oTimeZone As TimeZone

    Set oTimeZone = New TimeZone
    
    ' Store the object
    Set moLFObject = oLFObj
    
'    Scope = oLFObj.Scope
    ActionType = enAction
    MsgSource = nSource
    UserName = sUserName
    UserNameFull = sUserNameFull
    msStudyName = ""        ' Initialise to "" until we need it
    
    ProcessedStatus = LFProcessStatus.lfpUnProcessed
    MsgCreatedTimestamp = IMedNow
    MsgCreatedTimestamp_TZ = oTimeZone.TimezoneOffset
    
    ' Initialise rollback info to none
    RollbackMsgId = 0
    RollbackSource = 0
    
    ' Not processed, saved, sent or received yet
    ProcessedTimestamp = 0
    ProcessedTimestamp_TZ = 0
    SentTimestamp = 0
    SentTimestamp_TZ = 0
    ReceivedTimestamp = 0
    ReceivedTimestamp_TZ = 0
    
    ' No MessageId yet
    MessageId = 0
    
    ' Make sure we have both the study Name and ID
    Call GetStudyDetails(oDBCon)
    
    Set oTimeZone = Nothing
    
End Sub

'--------------------------------------------------------------
Friend Sub SetAsProcessed()
'--------------------------------------------------------------
Dim oTimeZone As TimeZone

    Set oTimeZone = New TimeZone
        
    ProcessedStatus = LFProcessStatus.lfpProcessed
    ProcessedTimestamp = IMedNow
    ProcessedTimestamp_TZ = oTimeZone.TimezoneOffset

    Set oTimeZone = Nothing
    
End Sub

'--------------------------------------------------------------
Public Sub Save(oDBCon As ADODB.Connection)
'--------------------------------------------------------------
' Insert this as a new message in the LFMessage table
' NCJ 14 Feb 03 - Must convert timestamps to standard nums
'--------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    ' Make sure we have both the study Name and ID
    Call GetStudyDetails(oDBCon)
    
    ' Generate new messageID if it doesn't have one
    ' (NB received messages will already have a MessageID)
    If Me.MessageId = 0 Then
        Me.MessageId = NextMessageId(oDBCon)
    End If
    ' Always generate a new sequence number
    Me.SequenceNo = NextSeqId(oDBCon)
    
    With Me
        sSQL = SQLInsertFields
                    
        sSQL = sSQL & " VALUES ( '" & .StudyName & "', " & .StudyId & ", '" & .Site & "', " & .SubjectId _
                    & ", " & .MsgSource & ", " & .MessageId & ", " & .Scope
        sSQL = sSQL & ", " & .VisitId & ", " & .VisitCycle _
                    & ", " & .EFormId & ", " & .EFormCycle _
                    & ", " & .ResponseId & ", " & .ResponseCycle _
                    & ", " & .QuestionId _
                    & ", '" & .UserName & "', '" & .UserNameFull & "'"
        sSQL = sSQL & ", " & .ActionType & ", " & .ProcessedStatus _
                    & ", " & .RollbackSource & ", " & .RollbackMsgId _
                    & ", " & LocalNumToStandard(CStr(.MsgCreatedTimestamp)) & ", " & .MsgCreatedTimestamp_TZ _
                    & ", " & LocalNumToStandard(CStr(.ProcessedTimestamp)) & ", " & .ProcessedTimestamp_TZ _
                    & ", " & LocalNumToStandard(CStr(.SentTimestamp)) & ", " & .SentTimestamp_TZ _
                    & ", " & LocalNumToStandard(CStr(.ReceivedTimestamp)) & ", " & .ReceivedTimestamp_TZ _
                    & ", " & .SequenceNo
        sSQL = sSQL & " )"
        
    End With
    
    oDBCon.Execute sSQL
  
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "LFMessage.Save"

End Sub

'--------------------------------------------------------------
Private Function SQLInsertFields() As String
'--------------------------------------------------------------
' SQL to INSERT the correct list of fields in the LFMessage table
'--------------------------------------------------------------
Dim sSQL As String

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
    sSQL = sSQL & " )"

    SQLInsertFields = sSQL

End Function

'--------------------------------------------------------------
Public Function Load(oDBCon As ADODB.Connection, _
                    ByVal sSite As String, ByVal nSource As Integer, ByVal lMsgId As Long) As Boolean
'--------------------------------------------------------------
' Load a given message
' Returns FALSE if message was not found (in this case object will not have any values)
'--------------------------------------------------------------
Dim sSQL As String
Dim rsLFMsg As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Set rsLFMsg = New ADODB.Recordset
    
    sSQL = "SELECT * FROM LFMessage WHERE " _
            & " TrialSite = '" & sSite & "'" _
            & " AND Source = " & nSource _
            & " AND MessageID = " & lMsgId
    rsLFMsg.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    If rsLFMsg.RecordCount > 0 Then
        ' We found the message
        Call LoadFromRS(rsLFMsg)
        Load = True
    Else
        ' No message found
        Load = False
    End If
    
    rsLFMsg.Close
    Set rsLFMsg = Nothing
        
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "LFMessage.Load"

End Function

'--------------------------------------------------------------
Friend Sub LoadFromRS(rsLFMsg As ADODB.Recordset)
'--------------------------------------------------------------
' Load all values from recordset of values from the LFMessage table
'--------------------------------------------------------------

    On Error GoTo ErrHandler
    
    msStudyName = rsLFMsg!ClinicalTrialName
    
    ' Update our LFObject
    Set moLFObject = New LFObject
    Call moLFObject.LoadFromRS(rsLFMsg, rsLFMsg!Scope)
     
    MsgSource = rsLFMsg!Source
    MessageId = rsLFMsg!MessageId
    ActionType = rsLFMsg!MsgType
    ProcessedStatus = rsLFMsg!ProcessedStatus
    RollbackSource = rsLFMsg!RollbackSource
    RollbackMsgId = rsLFMsg!RollbackMessageId

    UserName = rsLFMsg!UserName
    UserNameFull = rsLFMsg!UserNameFull

    MsgCreatedTimestamp = rsLFMsg!MsgCreatedTimestamp
    MsgCreatedTimestamp_TZ = rsLFMsg!MsgCreatedTimestamp_TZ
    ProcessedTimestamp = rsLFMsg!ProcessedTimestamp
    ProcessedTimestamp_TZ = rsLFMsg!ProcessedTimestamp_TZ
    SentTimestamp = rsLFMsg!SentTimestamp
    SentTimestamp_TZ = rsLFMsg!SentTimestamp_TZ
    ReceivedTimestamp = rsLFMsg!ReceivedTimestamp
    ReceivedTimestamp_TZ = rsLFMsg!ReceivedTimestamp_TZ

    SequenceNo = rsLFMsg!SequenceNo
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "LFMessage.LoadFromRS"

End Sub

'----------------------------------------------------------------------------------------'
Public Sub DoRollback(oDBCon As ADODB.Connection)
'----------------------------------------------------------------------------------------'
' Roll back this LF message, i.e. undo what it did and mark it as Refused
'----------------------------------------------------------------------------------------'

    Call ExecuteRollback(oDBCon, Me)
    Call Refuse(oDBCon)
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub DoAction(oDBCon As ADODB.Connection, Optional bSaveAsProcessed As Boolean = False)
'----------------------------------------------------------------------------------------'
' Actually perform the lock/freeze operation represented by this message
' NB We don't do Rollbacks here (see DoRollback)
' If bSaveAsProcessed = TRUE, we update the message as Processed in the DB
'----------------------------------------------------------------------------------------'
Dim enLockSetting As LockStatus

    On Error GoTo ErrHandler
    
    If Me.ActionType = lfaUnfreeze Then
        ' Deal with Unfreeze separately
        ' NCJ 9 Jan 03 - New method which "rolls back" freeze
        Call UnfreezeObject(oDBCon, moLFObject)
    Else
        ' It's Lock, Unlock or Freeze
        Select Case Me.ActionType
        Case LFAction.lfaLock
            enLockSetting = LockStatus.lsLocked
        Case LFAction.lfaUnlock
            enLockSetting = LockStatus.lsUnlocked
        Case LFAction.lfaFreeze
            enLockSetting = LockStatus.lsFrozen
        End Select
    
        Select Case moLFObject.Scope
        Case LFScope.lfscQuestion
            Call SetDataItemLockStatus(oDBCon, enLockSetting, moLFObject)
            If Me.ActionType = lfaUnlock Then
                ' Ensure eForm, visit & subject are unlocked
                Call UnlockCRFPageInstance(oDBCon, moLFObject)
            End If
            
        Case LFScope.lfscEForm
            Call SetCRFPageInstanceLockStatus(oDBCon, enLockSetting, moLFObject)
            If Me.ActionType = lfaUnlock Then
                ' Ensure visit & subject are unlocked
                Call UnlockVisitInstance(oDBCon, moLFObject)
            End If
            
        Case LFScope.lfscVisit
            Call SetVisitInstanceLockStatus(oDBCon, enLockSetting, moLFObject)
            If Me.ActionType = lfaUnlock Then
                ' Ensure subject is unlocked
                Call UnlockTrialSubject(oDBCon, moLFObject)
            End If
            
        Case LFScope.lfscSubject
            Call SetTrialSubjectLockStatus(oDBCon, enLockSetting, moLFObject)

        End Select
    End If
    
    ' Do we want to update the DB with Processed status?
    If bSaveAsProcessed Then
        Call ChangeProcessStatus(oDBCon, LFProcessStatus.lfpProcessed)
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "LFMessage.DoAction"

End Sub

'--------------------------------------------------------------
Friend Sub Refuse(oDBCon As ADODB.Connection)
'--------------------------------------------------------------
' Set this message to Refused and save to DB
'--------------------------------------------------------------

    Call ChangeProcessStatus(oDBCon, LFProcessStatus.lfpRefused)
    
End Sub

'--------------------------------------------------------------
Public Sub ChangeProcessStatus(oDBCon As ADODB.Connection, enNewStatus As LFProcessStatus)
'--------------------------------------------------------------
' Change Processed status and save to DB
' NCJ 13 Apr 04 - Convert timestamp to Standard
'--------------------------------------------------------------
Dim oTimeZone As TimeZone
Dim sSQL As String

    On Error GoTo ErrHandler
    
    Set oTimeZone = New TimeZone
        
    ProcessedStatus = enNewStatus
    ProcessedTimestamp = IMedNow
    ProcessedTimestamp_TZ = oTimeZone.TimezoneOffset

    Set oTimeZone = Nothing

    sSQL = "UPDATE LFMessage SET ProcessedStatus = " & enNewStatus _
            & ", ProcessedTimeStamp = " & LocalNumToStandard(CStr(ProcessedTimestamp)) _
            & ", ProcessedTimestamp_TZ = " & ProcessedTimestamp_TZ _
            & SQLMessageWhere
    oDBCon.Execute sSQL
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "LFMessage.ChangeProcessStatus"

End Sub

'--------------------------------------------------------------
Public Sub SetAsSent(oDBCon As ADODB.Connection)
'--------------------------------------------------------------
' Set this message as sent
' by updating its Sent time stamps
' NCJ 13 Apr 04 - Convert timestamp to Standard
'--------------------------------------------------------------
Dim oTimeZone As TimeZone
Dim sSQL As String

    On Error GoTo ErrHandler
    
    Set oTimeZone = New TimeZone
        
    SentTimestamp = IMedNow
    SentTimestamp_TZ = oTimeZone.TimezoneOffset

    Set oTimeZone = Nothing

    sSQL = "UPDATE LFMessage SET SentTimeStamp = " & LocalNumToStandard(CStr(SentTimestamp)) _
            & ", SentTimestamp_TZ = " & SentTimestamp_TZ _
            & SQLMessageWhere
    oDBCon.Execute sSQL

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "LFMessage.SetAsSent"

End Sub

'--------------------------------------------------------------
Private Function SQLMessageWhere() As String
'--------------------------------------------------------------
' SQL WHERE clause to identify current message
'--------------------------------------------------------------
Dim sSQL As String
    
    sSQL = sSQL & " WHERE " _
            & " TrialSite = '" & Me.Site & "'" _
            & " AND Source = " & Me.MsgSource _
            & " AND MessageID = " & Me.MessageId

    SQLMessageWhere = sSQL
    
End Function

'--------------------------------------------------------------
Public Property Get MessageObject() As LFObject
'--------------------------------------------------------------
' Return the "object" to which this message refers
'--------------------------------------------------------------

    Set MessageObject = moLFObject

End Property

'------------------------------------------------------------
Public Property Get StudyName() As String
'------------------------------------------------------------
' Study name
'------------------------------------------------------------
    
    StudyName = msStudyName
    
End Property

'------------------------------------------------------------
Public Property Let StudyName(sStudyName As String)
'------------------------------------------------------------
' Let Study name if known
'------------------------------------------------------------

    msStudyName = sStudyName
    
End Property

'------------------------------------------------------------
Public Property Get StudyId() As Long
'------------------------------------------------------------
' The Study ID
'------------------------------------------------------------
    
    StudyId = moLFObject.StudyId
    
End Property

'------------------------------------------------------------
Public Property Let StudyId(lStudyId As Long)
'------------------------------------------------------------
' The Study ID
'------------------------------------------------------------
    
    moLFObject.StudyId = lStudyId

End Property

'------------------------------------------------------------
Private Sub GetStudyDetails(oDBCon As ADODB.Connection)
'------------------------------------------------------------
' Make sure we have a Study Name AND a Study ID
' assuming we have at least one of them!
'------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
    
    If moLFObject.StudyId = 0 Then
        ' Assume we have a study name
        sSQL = "SELECT ClinicalTrialId from ClinicalTrial WHERE ClinicalTrialName = '" & msStudyName & "'"
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, oDBCon, adOpenForwardOnly, adLockReadOnly, adCmdText
        moLFObject.StudyId = rsTemp!ClinicalTrialId
        rsTemp.Close
        Set rsTemp = Nothing
    End If
    
    If msStudyName = "" Then
        ' Assume we have a Study ID
        sSQL = "SELECT ClinicalTrialName from ClinicalTrial WHERE ClinicalTrialId = " & moLFObject.StudyId
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, oDBCon, adOpenForwardOnly, adLockReadOnly, adCmdText
        msStudyName = rsTemp!ClinicalTrialName
        rsTemp.Close
        Set rsTemp = Nothing
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "LFMessage.GetStudyDetails"

End Sub

'------------------------------------------------------------
Public Function CanExecute(oDBCon As ADODB.Connection) As Boolean
'------------------------------------------------------------
' Returns TRUE if we can execute this LFMessage on the specified database
' i.e. if target data not changed and no dependencies on previously refused actions
'------------------------------------------------------------

    CanExecute = CanPerformLFAction(oDBCon, Me)

End Function

'------------------------------------------------------------
' PROPERTIES TO RETURN STORED OBJECT VALUES
'------------------------------------------------------------

'------------------------------------------------------------
Public Property Get Site() As String
    Site = moLFObject.Site
End Property
Public Property Let Site(sSite As String)
    moLFObject.Site = sSite
End Property

'------------------------------------------------------------
Public Property Get SubjectId() As Long
    SubjectId = moLFObject.SubjectId
End Property
Public Property Let SubjectId(lSubjID As Long)
    moLFObject.SubjectId = lSubjID
End Property

'------------------------------------------------------------
Public Property Get VisitId() As Long
    VisitId = moLFObject.VisitId
End Property
Public Property Let VisitId(lVisitId As Long)
    moLFObject.VisitId = lVisitId
End Property

'------------------------------------------------------------
Public Property Get VisitCycle() As Integer
    VisitCycle = moLFObject.VisitCycle
End Property
Public Property Let VisitCycle(nCycle As Integer)
    moLFObject.VisitCycle = nCycle
End Property

'------------------------------------------------------------
Public Property Get EFormId() As Long
    EFormId = moLFObject.EFormId
End Property
Public Property Let EFormId(lEFormId As Long)
    moLFObject.EFormId = lEFormId
End Property

'------------------------------------------------------------
Public Property Get EFormCycle() As Integer
    EFormCycle = moLFObject.EFormCycle
End Property
Public Property Let EFormCycle(nCycle As Integer)
    moLFObject.EFormCycle = nCycle
End Property

'------------------------------------------------------------
Public Property Get ResponseId() As Long
    ResponseId = moLFObject.ResponseId
End Property
Public Property Let ResponseId(lResponseId As Long)
    moLFObject.ResponseId = lResponseId
End Property

'------------------------------------------------------------
Public Property Get ResponseCycle() As Integer
    ResponseCycle = moLFObject.ResponseCycle
End Property
Public Property Let ResponseCycle(nCycle As Integer)
    moLFObject.ResponseCycle = nCycle
End Property

'------------------------------------------------------------
Public Property Get QuestionId() As Long
    QuestionId = moLFObject.QuestionId
End Property
Public Property Let QuestionId(lQuestionId As Long)
    moLFObject.QuestionId = lQuestionId
End Property

'------------------------------------------------------------
Public Property Get Scope() As LFScope
    Scope = moLFObject.Scope
End Property
Public Property Let Scope(enScope As LFScope)
    moLFObject.Scope = enScope
End Property

'------------------------------------------------------------
Public Property Get ScopeText() As String
    ScopeText = GetScopeText(Scope)
End Property

'------------------------------------------------------------
Public Property Get ActionText() As String
    ActionText = GetActionText(ActionType)
End Property

'------------------------------------------------------------
Private Sub Class_Initialize()
'------------------------------------------------------------

    Set moLFObject = New LFObject

End Sub

'------------------------------------------------------------
Private Sub Class_Terminate()
'------------------------------------------------------------

    Set moLFObject = Nothing

End Sub

'--------------------------------------------------------------
Private Function NextMessageId(oDBCon As ADODB.Connection) As Long
'--------------------------------------------------------------
' Get the next MessageID for the LFMessage table
'--------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    sSQL = "SELECT max(MessageID) as MaxMessageId FROM LFMessage "
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, oDBCon, adOpenForwardOnly, adLockReadOnly, adCmdText
                                            
    If IsNull(rsTemp!MaxMessageId) Then
        NextMessageId = 1
    Else
        NextMessageId = rsTemp!MaxMessageId + 1
    End If

    rsTemp.Close
    Set rsTemp = Nothing
    
End Function

'--------------------------------------------------------------
Private Function NextSeqId(oDBCon As ADODB.Connection) As Long
'--------------------------------------------------------------
' Get the next sequence number for the LFMessage table
'--------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    sSQL = "SELECT max(SequenceNo) as MaxSeqNo FROM LFMessage "
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, oDBCon, adOpenForwardOnly, adLockReadOnly, adCmdText

    If IsNull(rsTemp!MaxSeqNo) Then
        NextSeqId = 1
    Else
        NextSeqId = rsTemp!MaxSeqNo + 1
    End If

    rsTemp.Close
    Set rsTemp = Nothing
    
End Function


