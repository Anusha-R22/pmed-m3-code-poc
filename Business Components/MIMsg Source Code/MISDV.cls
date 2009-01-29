VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MISDV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: MISDV.cls
' Toby Aldridge, InferMed, Nov 2001
' SDV class for MACRO
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' TA Nov 2001 - Initial development
' NCJ 11 Oct 02 - Added Queried and Cancelled as new SDV statuses
' NCJ 14-15 Oct 02 - Coding for new status changes; bug fixing of timestamps/timezones
' NCJ 5 Nov 02 - We allow ANY status changes
' TA 18/11/2002: Added new fields
' NCJ 20 Nov 02 - Corrected MIMessageType property
' DPH 26/03/2004 - Tidy up on class termination / tidy objects
'----------------------------------------------------

Option Explicit

'SDV statuses
Public Enum eSDVMIMStatus
    ssPlanned = 0
    ssQueried = 1 'new in 3.0
    ssDone = 2
    ssCancelled = 3 'new in 3.0
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
'ReLoad an SDV set of MIMessages according to the curently loaded site
'       , object source and object id.
'----------------------------------------------------------------------------------------'
Dim sCon As String
Dim lObjectId As Long
Dim nObjectSource As Integer
Dim sSite As String

    sCon = msCon
    lObjectId = Me.SDVID
    nObjectSource = Me.SDVSource
    sSite = Me.Site
    Call Load(sCon, lObjectId, nObjectSource, Site)
  
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MISDV.ReLoad"
    
End Sub


'----------------------------------------------------------------------------------------'
Public Sub Load(sCon As String, lObjectId As Long, nObjectSource As Integer, sSite As String)
Attribute Load.VB_Description = "Load an SDV set of MIMessages according to site, object source and object id."
'----------------------------------------------------------------------------------------'
'Load an SDV set of MIMessages according to site, object source and object id.
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
    Err.Raise Err.Number, , Err.Description & "|MISDV.Load"

End Sub

'----------------------------------------------------------------------------------------'
Public Sub Save()
Attribute Save.VB_Description = "Save the current SDV set of MIMessages"
'----------------------------------------------------------------------------------------'
'Save the current SDV set of MIMessages
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
    Err.Raise Err.Number, , Err.Description & "|MISDV.Save"

End Sub

'----------------------------------------------------------------------------------------'
Public Function Raise(sCon As String, _
                    sText As String, nStatus As eSDVMIMStatus, _
                    sUserName As String, sUserNameFull As String, nSource As MIMsgSource, _
                    dblTimestamp As Double, nTimezoneOffset As Integer, _
                    nScope As MIMsgScope, _
                    sStudyName As String, sSite As String, lSubjectId As Long, _
                    lVisitId As Long, nVisitCycle As Integer, lEFormTaskId As Long, _
                    lResponseTaskId As Long, nResponseCycle As Integer, _
                        lEFormId As Long, nEFormCycle As Integer, lQuestionId As Long, sDataUsername As String, _
                    dblResponseTimeStamp As Double, sResponseValue As String) As Integer
'----------------------------------------------------------------------------------------'
'Raise an SDV (create first MIMessage row).
'Raises mimeInvalidForThisStatus error if already a current row.
'Function returns index in messages collection of create message row.
' NCJ 14 Oct 02 - Depending on Scope, some arguments may be 0
'----------------------------------------------------------------------------------------'
Dim oMsg As MIMsg
Dim oMsgState As MIMsgState
Dim colSerStates As Collection
Dim oState As MIMsgState
    
    If Not mcolMIMessages Is Nothing Then     'already raised
        Err.Raise MIMsgErrors.mimeInvalidForThisStatus, , "SDV is already created"
    End If

    On Error GoTo ErrLabel

    msCon = sCon
    Set mcolMIMsgStates = New Collection
    Set mcolMIMessages = New Collection
    Raise = Add(sText, nStatus, _
                sUserName, sUserNameFull, nSource, _
                dblTimestamp, nTimezoneOffset, _
                nScope, sStudyName, sSite, lSubjectId, _
                lVisitId, nVisitCycle, lEFormTaskId, _
                lResponseTaskId, nResponseCycle, dblResponseTimeStamp, sResponseValue, _
                lEFormId, nEFormCycle, lQuestionId, sDataUsername)
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MISDV.Raise"

End Function

'----------------------------------------------------------------------------------------'
Public Function Done(sText As String, sUserName As String, sUserNameFull As String, _
                        nSource As MIMsgSource, _
                        dblTimestamp As Double, nTimezoneOffset As Integer, _
                        Optional dblResponseTimeStamp As Double = 0, _
                        Optional sResponseValue As String = "") As Integer
'----------------------------------------------------------------------------------------'
' Mark an SDV as done (by adding a 'done' MIMessage row)
'Raises mimeInvalidForThisStatus error if not allowed for current status.
'Funtion returns index in messages collection of create message row.
' NCJ 14 oct 02 - Call generic ChangeStatus routine
' NCJ 15 Oct 02 - Added dblTimestamp and nTimezoneOffset parameters
'----------------------------------------------------------------------------------------'

    Done = ChangeStatus(ssDone, sText, sUserName, sUserNameFull, nSource, _
                        dblTimestamp, nTimezoneOffset, _
                        dblResponseTimeStamp, sResponseValue)
    
End Function

'----------------------------------------------------------------------------------------'
Public Function ChangeStatus(enNewStatus As eSDVMIMStatus, _
                        sText As String, sUserName As String, sUserNameFull As String, _
                        nSource As MIMsgSource, _
                        dblTimestamp As Double, nTimezoneOffset As Integer, _
                        Optional dblResponseTimeStamp As Double = 0, _
                        Optional sResponseValue As String = "") As Integer
'----------------------------------------------------------------------------------------'
' Change the status of this SDV
' Raises mimeInvalidForThisStatus error if enNewStatus not allowed for current status.
' Returns index in Messages collection of newly created message row.
' NCJ 15 Oct 02 - Added dblTimestamp and nTimezoneOffset parameters
'----------------------------------------------------------------------------------------'
    

    
    If Not ValidStatusTransition(enNewStatus) Then
        Err.Raise MIMsgErrors.mimeInvalidForThisStatus, , "Status has not changed"
    End If
    On Error GoTo ErrLabel
    ChangeStatus = Add(sText, enNewStatus, _
                sUserName, sUserNameFull, nSource, _
                dblTimestamp, nTimezoneOffset, _
                Scope, StudyName, Site, SubjectId, VisitId, VisitCycle, EFormTaskId, _
                ResponseTaskId, ResponseCycle, dblResponseTimeStamp, sResponseValue, _
                EFormId, EFormCycle, QuestionId, DataUserName, _
                SDVID, SDVSource)

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MISDV.ChangeStatus"

End Function

'----------------------------------------------------------------------------------------'
Private Function ValidStatusTransition(enNewStatus As eSDVMIMStatus) As Boolean
'----------------------------------------------------------------------------------------'
' Returns TRUE if NewStatus is a valid status transition for this SDV
' NCJ 5 Nov 02 - We now allow any status transitions (except to itself)
'----------------------------------------------------------------------------------------'

    ValidStatusTransition = (CurrentStatus <> enNewStatus)

'    ValidStatusTransition = False
'
'    ' Mustn't be the same status
'    If CurrentStatus = enNewStatus Then Exit Function
'
'    Select Case CurrentStatus
'    Case eSDVMIMStatus.ssPlanned, eSDVMIMStatus.ssQueried
'        ' Can go to anything (except the same, already screened out)
'        ValidStatusTransition = True
'
'    Case eSDVMIMStatus.ssDone
'        ' Can only go back to Planned
'        If enNewStatus = ssPlanned Then
'            ValidStatusTransition = True
'        End If
'
'    Case eSDVMIMStatus.ssCancelled
'        ' No status transitions allowed
'
'    End Select
 
End Function

'----------------------------------------------------------------------------------------'
Private Function Add(sText As String, nStatus As eSDVMIMStatus, _
                    sUserName As String, sUserNameFull As String, nSource As MIMsgSource, _
                    dblTimestamp As Double, nTimezoneOffset As Integer, _
                    nScope As MIMsgScope, _
                    sStudyName As String, sSite As String, lSubjectId As Long, _
                    lVisitId As Long, nVisitCycle As Integer, lEFormTaskId As Long, _
                    lResponseTaskId As Long, nResponseCycle As Integer, _
                    dblResponseTimeStamp As Double, sResponseValue As String, _
                     lEFormId As Long, nEFormCycle As Integer, lQuestionId As Long, sDataUsername As String, _
                    Optional lObjectId As Long = -1, Optional nObjectSource As Integer = -1) As Integer
'----------------------------------------------------------------------------------------'
' general purpose routine to add a MIMessage row to an SDV.
' RS 30/09/2002: Added Timestamp & TimezoneOffset parameters
' NCJ 15 Oct - Moved Timestamp & TimezoneOffset parameters
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
    
    oMsg.Init sUserName, sUserNameFull, nSource, MIMsgType.mimtSDVMark, nStatus, _
                    nScope, sStudyName, sSite, lSubjectId, _
                    lVisitId, nVisitCycle, lEFormTaskId, _
                    lResponseTaskId, nResponseCycle, _
                    dblResponseTimeStamp, sResponseValue, _
                    dblTimestamp, nTimezoneOffset, _
                    lEFormId, nEFormCycle, lQuestionId, sDataUsername, _
                    lObjectId, nObjectSource
                            
    oMsg.Text = sText
                            
    'uncurrent the current message
    If Not moCurrentMsg Is Nothing Then
        moCurrentMsg.Processed = mimpProcessed
        moCurrentMsg.History = MIMsgHistory.mimhNotCurrent
    End If
    mcolMIMessages.Add oMsg
    Set moCurrentMsg = oMsg
    
    Set oMsg = Nothing
    'return the new index
    Add = mcolMIMessages.Count

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MISDV.Add"

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
    Err.Raise Err.Number, , Err.Description & "|MISDV.SetText"

End Sub

'----------------------------------------------------------------------------------------'
Public Property Get CurrentMessage() As MIMsg
Attribute CurrentMessage.VB_Description = "Return the current message row for the SDV."
'----------------------------------------------------------------------------------------'
'Return the current message row for the SDV.
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


'properties common to the SDV (all MIMessage rows)

Public Property Get MIMessageType() As MIMsgType
    MIMessageType = MIMsgType.mimtSDVMark
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

Public Property Get SDVID() As Long
    SDVID = moCurrentMsg.MIMessageId
End Property

Public Property Get SDVSource() As MIMsgSource
    SDVSource = moCurrentMsg.MIMessagesSource
End Property

Public Property Get CurrentStatus() As eSDVMIMStatus
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

