VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIDiscrepancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: MIDiscrepancy.cls
' Toby Aldridge, InferMed, Nov 2001
' Discrepancy class for MACRO
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' TA Nov 2001 - Initial development
'   NCJ 14 Oct 02 - Changed enumeration names
' NCJ 5 Nov 02 - Fixed bugs due to Timezone additions
'TA 18/11/2002: Added new fields
' DPH 26/03/2004 - Tidy up on class termination / tidy objects
'----------------------------------------------------------------------------------------'

Option Explicit

Public Enum eDiscrepancyMIMStatus
    dsRaised = 0
    dsResponded = 1
    dsClosed = 2
End Enum

Private mcolMIMsgStates As Collection
Private mcolMIMessages As Collection

'connection string
Private msCon As String

'refrence to current message
Private moCurrentMsg As MIMsg

'----------------------------------------------------------------------------------------'
Public Sub ReLoad()
'----------------------------------------------------------------------------------------'
'ReLoad an discrepancy set of MIMessages according to the curently loaded site
'       , object source and object id.
'----------------------------------------------------------------------------------------'
Dim sCon As String
Dim lObjectId As Long
Dim nObjectSource As Integer
Dim sSite As String

    sCon = msCon
    lObjectId = Me.DiscrepancyID
    nObjectSource = Me.DiscrepancySource
    sSite = Me.Site
    Call Load(sCon, lObjectId, nObjectSource, Site)
  
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIDiscrepancy.ReLoad"
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Load(sCon As String, lObjectId As Long, nObjectSource As Integer, sSite As String)
Attribute Load.VB_Description = "Load an discrepancy set of MIMessages according to site, object source and object id."
'----------------------------------------------------------------------------------------'
'Load an discrepancy set of MIMessages according to site, object source and object id.
'----------------------------------------------------------------------------------------'
' REVISIONS
' DPH 26/03/2004 - tidy up objects
'----------------------------------------------------------------------------------------'
Dim oMsg As MIMsg
Dim oDS As MIMsgDS
Dim colSerStates As Collection
Dim oState As MIMsgState
Dim i As Long
   
    On Error GoTo ErrLabel
    
    Set oDS = New MIMsgDS
    msCon = sCon
    Set mcolMIMsgStates = CollectionDeSerialise(oDS.Load(sCon, lObjectId, nObjectSource, sSite, False))
    Set oDS = Nothing
    
    Set mcolMIMessages = New Collection
    For i = 1 To mcolMIMsgStates.Count
        Set oState = mcolMIMsgStates(i)
        Set oMsg = New MIMsg
        oMsg.SetState oState
        mcolMIMessages.Add oMsg
        If oMsg.History = mimhCurrent Then
            Set moCurrentMsg = oMsg
        End If
    Next
    
    Set oMsg = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIDiscrepancy.Load"
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Save()
Attribute Save.VB_Description = "Save the current discrepancy set of MIMessages"
'----------------------------------------------------------------------------------------'
'Save the current discrepancy set of MIMessages
'----------------------------------------------------------------------------------------'
' REVISIONS
' DPH 26/03/2004 - tidy up objects
'----------------------------------------------------------------------------------------'
Dim oDS As MIMsgDS
Dim colSerStates As Collection
Dim oState As MIMsgState
Dim i As Long
Dim sUpdated
Dim oMsg As MIMsg

    On Error GoTo ErrLabel
    
    Set oDS = New MIMsgDS
    sUpdated = oDS.Save(msCon, CollectionSerialise(mcolMIMsgStates))
    Set mcolMIMsgStates = CollectionDeSerialise(sUpdated)
    Set oDS = Nothing
    
    Set mcolMIMessages = New Collection
    For i = 1 To mcolMIMsgStates.Count
        Set oState = mcolMIMsgStates(i)
        Set oMsg = New MIMsg
        oMsg.SetState oState
        mcolMIMessages.Add oMsg
        If oMsg.History = mimhCurrent Then
            Set moCurrentMsg = oMsg
        End If
    Next

    Set oMsg = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIDiscrepancy.Save"

End Sub

'----------------------------------------------------------------------------------------'
Public Function Raise(sCon As String, _
                    sText As String, nPriority As Integer, lExternalId As Long, _
                    sUserName As String, sUserNameFull As String, nSource As MIMsgSource, _
                    nScope As MIMsgScope, _
                    sStudyName As String, sSite As String, lSubjectId As Long, _
                    lVisitId As Long, nVisitCycle As Integer, lEFormTaskId As Long, _
                    lResponseTaskId As Long, nResponseCycle As Integer, dblResponseTimeStamp As Double, sResponseValue As String, _
                        lEFormId As Long, nEFormCycle As Integer, lQuestionId As Long, sDataUsername As String, _
                    dblTimestamp As Double, nTimezoneOffset As Integer) As Integer
'----------------------------------------------------------------------------------------'
'Raise an discrepancy (create first MIMessage row)
'Raises mimeInvalidForThisStatus error if not allowed for current status.
'Funtion returns index in messages collection of create message row.
' RS 30/09/2002: Added Timestamp & Timezone offset parameters
'----------------------------------------------------------------------------------------'
Dim oMsg As MIMsg
Dim oMsgState As MIMsgState
Dim colSerStates As Collection
Dim oState As MIMsgState

    
    If Not mcolMIMessages Is Nothing Then    'already raised
        Err.Raise MIMsgErrors.mimeInvalidForThisStatus, , "Discrepancy is already raised"
    End If
 
    On Error GoTo ErrLabel
    msCon = sCon
    Set mcolMIMsgStates = New Collection
    Set mcolMIMessages = New Collection
    ' RS 30/09/2002: Added Timestamp & Timezone offset parameters
    Raise = Add(sText, nPriority, lExternalId, eDiscrepancyMIMStatus.dsRaised, _
                sUserName, sUserNameFull, nSource, _
                nScope, sStudyName, sSite, lSubjectId, lVisitId, nVisitCycle, lEFormTaskId, _
                lResponseTaskId, nResponseCycle, dblResponseTimeStamp, sResponseValue, _
                lEFormId, nEFormCycle, lQuestionId, sDataUsername, _
                dblTimestamp, nTimezoneOffset)
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIDiscrepancy.Save"

End Function

'----------------------------------------------------------------------------------------'
Public Function CloseDown(sText As String, sUserName As String, sUserNameFull As String, nSource As MIMsgSource, _
                            dblResponseTimeStamp As Double, sResponseValue As String, _
                            dblTimestamp As Double, nTimezoneOffset As Integer) As Integer
'----------------------------------------------------------------------------------------'
'Close a discrepancy (by adding a closed MIMessage row)
'Raises mimeInvalidForThisStatus error if not allowed for current status.
'Funtion returns index in messages collection of create message row.
' NCJ 5 Nov 02 - Added TimeStamp and TimezoneOffset
'----------------------------------------------------------------------------------------'


    
    If CurrentStatus = eDiscrepancyMIMStatus.dsClosed Then   'not closed, can close
        Err.Raise MIMsgErrors.mimeInvalidForThisStatus, , "Discrepancy already closed"
    End If
    On Error GoTo ErrLabel
    CloseDown = Add(sText, CurrentMessage.Priority, ExternalId, eDiscrepancyMIMStatus.dsClosed, _
                sUserName, sUserNameFull, nSource, _
                Scope, StudyName, Site, SubjectId, VisitId, VisitCycle, EFormTaskId, _
                ResponseTaskId, ResponseCycle, dblResponseTimeStamp, sResponseValue, _
                EFormId, EFormCycle, QuestionId, DataUserName, _
                dblTimestamp, nTimezoneOffset, _
                DiscrepancyID, DiscrepancySource)
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIDiscrepancy.CloseDown"
    
End Function

'----------------------------------------------------------------------------------------'
Public Function Respond(sText As String, sUserName As String, sUserNameFull As String, nSource As MIMsgSource, _
                            dblResponseTimeStamp As Double, sResponseValue As String, _
                            dblTimestamp As Double, nTimezoneOffset As Integer) As Integer
'----------------------------------------------------------------------------------------'
'Respond to a discrepancy
'Raises mimeInvalidForThisStatus error if not allowed for current status.
'Function returns index in messages collection of create message row.
' NCJ 5 Nov 02 - Added TimeStamp and TimezoneOffset
'----------------------------------------------------------------------------------------'
    
    If CurrentStatus <> eDiscrepancyMIMStatus.dsRaised Then 'if raised, can respond
        Err.Raise MIMsgErrors.mimeInvalidForThisStatus, , "Discrepancy is not raised"
    End If

    On Error GoTo ErrLabel
    Respond = Add(sText, CurrentMessage.Priority, ExternalId, eDiscrepancyMIMStatus.dsResponded, _
                sUserName, sUserNameFull, nSource, _
                Scope, StudyName, Site, SubjectId, VisitId, VisitCycle, EFormTaskId, _
                ResponseTaskId, ResponseCycle, dblResponseTimeStamp, sResponseValue, _
                EFormId, EFormCycle, QuestionId, DataUserName, _
                dblTimestamp, nTimezoneOffset, _
                DiscrepancyID, DiscrepancySource)

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIDiscrepancy.Respond"
    
End Function

'----------------------------------------------------------------------------------------'
Public Function ReRaise(sText As String, sUserName As String, sUserNameFull As String, nSource As MIMsgSource, _
                            dblResponseTimeStamp As Double, sResponseValue As String, _
                            dblTimestamp As Double, nTimezoneOffset As Integer) As Integer
'----------------------------------------------------------------------------------------'
'Reraise a discrepancy (by adding a new 'raised' MIMessage row).
'Raises mimeInvalidForThisStatus error if not allowed for current status.
'Function returns index in messages collection of create message row.
' NCJ 5 Nov 02 - Added TimeStamp and TimezoneOffset
'----------------------------------------------------------------------------------------'
    
    If CurrentStatus <> eDiscrepancyMIMStatus.dsResponded Then  'only if responded, can reraise
        Err.Raise MIMsgErrors.mimeInvalidForThisStatus, , "Only responded discrepancies can be reraised"
    End If

    On Error GoTo ErrLabel
    
    ReRaise = Add(sText, CurrentMessage.Priority, ExternalId, eDiscrepancyMIMStatus.dsRaised, _
                sUserName, sUserNameFull, nSource, _
                Scope, StudyName, Site, SubjectId, VisitId, VisitCycle, EFormTaskId, _
                ResponseTaskId, ResponseCycle, dblResponseTimeStamp, sResponseValue, _
                EFormId, EFormCycle, QuestionId, DataUserName, _
                dblTimestamp, nTimezoneOffset, _
                DiscrepancyID, DiscrepancySource)

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIDiscrepancy.ReRaise"
    
End Function

'----------------------------------------------------------------------------------------'
Private Function Add(sText As String, nPriority As Integer, lExternalId As Long, _
                    nStatus As eDiscrepancyMIMStatus, _
                    sUserName As String, sUserNameFull As String, nSource As MIMsgSource, _
                    nScope As MIMsgScope, _
                    sStudyName As String, sSite As String, lSubjectId As Long, _
                    lVisitId As Long, nVisitCycle As Integer, lEFormTaskId As Long, _
                    lResponseTaskId As Long, nResponseCycle As Integer, dblResponseTimeStamp As Double, sResponseValue As String, _
                    lEFormId As Long, nEFormCycle As Integer, lQuestionId As Long, sDataUsername As String, _
                    dblTimestamp As Double, nTimezoneOffset As Integer, _
                    Optional lObjectId As Long = -1, Optional nObjectSource As Integer = -1) As Integer
'----------------------------------------------------------------------------------------'
' general purpose rotuine to add a MIMessage row to a discrepancy.
'Funtion returns index in messages collection of create message row.
' RS 30/09/2002:    Added Timestamp & TimezoneOffset parameter
'TA 18/11/2002: Added new fields
'----------------------------------------------------------------------------------------'
Dim oMsg As MIMsg
Dim oMsgState As MIMsgState
    
    On Error GoTo ErrLabel
    
    Set oMsgState = New MIMsgState
    mcolMIMsgStates.Add oMsgState
    Set oMsg = New MIMsg
    oMsg.SetState oMsgState
    Set oMsgState = Nothing
    oMsg.Init sUserName, sUserNameFull, nSource, MIMsgType.mimtDiscrepancy, nStatus, _
                            nScope, sStudyName, sSite, lSubjectId, lVisitId, nVisitCycle, lEFormTaskId, _
                            lResponseTaskId, nResponseCycle, dblResponseTimeStamp, sResponseValue, _
                            dblTimestamp, nTimezoneOffset, _
                            lEFormId, nEFormCycle, lQuestionId, sDataUsername, _
                            lObjectId, nObjectSource
                            
    'set other values
    oMsg.Text = sText
    oMsg.Priority = nPriority
    oMsg.ExternalId = lExternalId
    
    'uncurrent the current message
    If Not moCurrentMsg Is Nothing Then
        moCurrentMsg.History = MIMsgHistory.mimhNotCurrent
        moCurrentMsg.Processed = mimpProcessed
    End If
    mcolMIMessages.Add oMsg
    Set moCurrentMsg = oMsg
    
    Set oMsg = Nothing
    'return the new index
    Add = mcolMIMessages.Count

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIDiscrepancy.Add"
    
End Function


'----------------------------------------------------------------------------------------'
Public Sub SetText(sText As String, sUserName As String)
Attribute SetText.VB_Description = "Set the text for the current message row. mimeInvalidUserName,mimeAlreadyProcessed,mimeInvalidForThisStatus errors can occur."
'----------------------------------------------------------------------------------------'
'Set the text for the current message row.
'mimeInvalidUserName,mimeAlreadySent,mimeInvalidForThisStatus errors can occur.
'----------------------------------------------------------------------------------------'
    
    If sUserName <> CurrentMessage.UserName Then
        Err.Raise MIMsgErrors.mimeInvalidUserName, , "This message can only be edited by " & CurrentMessage.UserName
    End If
    If CurrentMessage.TimeSent <> 0 Then
        Err.Raise MIMsgErrors.mimeAlreadySent, , "This message has already been sent"
    End If

    On Error GoTo ErrLabel
    CurrentMessage.Text = sText

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIDiscrepancy.SetText"
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub SetPriority(nPriority As Integer, nSource As MIMsgSource)
Attribute SetPriority.VB_Description = "Set the priority for the current message row. mimeNotCreatedHere,mimeInvalidForThisStatus errors can occur."
'----------------------------------------------------------------------------------------'
'Set the priority for the current message row.
'mimeNotCreatedHere,mimeInvalidForThisStatus errors can occur.
'----------------------------------------------------------------------------------------'
    
    If nSource <> CurrentMessage.Source Then
        Err.Raise MIMsgErrors.mimeNotCreatedHere, , "This message was not created here"
    End If
    If CurrentStatus <> dsRaised Then
        Err.Raise MIMsgErrors.mimeInvalidForThisStatus, , "Priority can only be changed on raised discrepancies"
    End If

    On Error GoTo ErrLabel
    CurrentMessage.Priority = nPriority
    CurrentMessage.TimeSent = 0

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIDiscrepancy.SetPriority"
    
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get CurrentMessage() As MIMsg
Attribute CurrentMessage.VB_Description = "Return the current message row for the discrepancy."
'----------------------------------------------------------------------------------------'
'Return the current message row for the discrepancy.
'----------------------------------------------------------------------------------------'

    If mcolMIMessages.Count > 0 Then
        Set CurrentMessage = mcolMIMessages(mcolMIMessages.Count)
    End If

End Property

'----------------------------------------------------------------------------------------'
Public Property Get Messages() As Collection
Attribute Messages.VB_Description = "Collection of all the message rows."
'----------------------------------------------------------------------------------------'
'Collection of all the message rows.
'----------------------------------------------------------------------------------------'

    Set Messages = mcolMIMessages
    
End Property


' Properties common to the discrepancy (all MIMessage rows)

Public Property Get MIMessageType() As MIMsgType
    MIMessageType = mimtDiscrepancy
End Property

Public Property Get ExternalId() As Long
    ExternalId = moCurrentMsg.ExternalId
End Property

Public Property Get Scope() As MIMsgScope
    Scope = moCurrentMsg.Scope
End Property

Public Property Get StudyName() As String
    StudyName = moCurrentMsg.StudyName
End Property

Public Property Get Site() As String
    Site = moCurrentMsg.Site
End Property

Public Property Get SubjectId() As Long
    SubjectId = moCurrentMsg.SubjectId
End Property

Public Property Get VisitId() As Long
    VisitId = moCurrentMsg.VisitId
End Property

Public Property Get VisitCycle() As Integer
    VisitCycle = moCurrentMsg.VisitCycle
End Property

Public Property Get EFormTaskId() As Long
    EFormTaskId = moCurrentMsg.EFormTaskId
End Property

Public Property Get ResponseTaskId() As Long
    ResponseTaskId = moCurrentMsg.ResponseTaskId
End Property

Public Property Get ResponseCycle() As Integer
    ResponseCycle = moCurrentMsg.ResponseCycle
End Property


Public Property Get DiscrepancyID() As Long
    DiscrepancyID = moCurrentMsg.MIMessageId
End Property

Public Property Get DiscrepancySource() As MIMsgSource
    DiscrepancySource = moCurrentMsg.MIMessagesSource
End Property

Public Property Get CurrentStatus() As eDiscrepancyMIMStatus
    CurrentStatus = moCurrentMsg.Status
End Property

 
Public Property Get EFormId() As Long
    EFormId = moCurrentMsg.EFormId
End Property
 
Public Property Get EFormCycle() As Integer
    EFormCycle = moCurrentMsg.EFormCycle
End Property

Public Property Get QuestionId() As Long
    QuestionId = moCurrentMsg.QuestionId
End Property

Public Property Get DataUserName() As String
    DataUserName = moCurrentMsg.DataUserName
End Property

'----------------------------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------------------------
' DPH 26/03/2004 - Tidy up on termination
'----------------------------------------------------------------------
Dim oMIMsg As MIMsg

    ' remove references to MIMsgState objects
    Call CollectionRemoveAll(mcolMIMsgStates)
    ' terminate MIMsg Objects
    If Not (mcolMIMessages Is Nothing) Then
        For Each oMIMsg In mcolMIMessages
            oMIMsg.Terminate
        Next
        
        Call CollectionRemoveAll(mcolMIMessages)
    End If
    
    Set moCurrentMsg = Nothing
    Set mcolMIMsgStates = Nothing
    Set mcolMIMessages = Nothing

End Sub
