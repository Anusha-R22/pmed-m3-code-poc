VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MINote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: MINote.cls
' Toby Aldridge, InferMed, Nov 2001
' Note class for MACRO
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' TA Nov 2001 - Initial development
'   NCJ 14 Oct 02 - Changed enumeration names
' DPH 26/03/2004 - Tidy up on class termination / tidy objects
'---------------------------------------------------------

Option Explicit

'note statuses
Public Enum eNoteMIMStatus
    nsPublic = 0
    nsPrivate = 1
End Enum

Private mcolMIMsgStates As Collection
Private mcolMIMessages As Collection

'connection string
Private msCon As String

'refrence to current message
Private moCurrentMsg As MIMsg

'----------------------------------------------------------------------------------------'
Public Sub Load(sCon As String, lId As Long, nSource As Integer, sSite As String)
Attribute Load.VB_Description = "Load a note"
'----------------------------------------------------------------------------------------'
'Load a note
'----------------------------------------------------------------------------------------'
' REVISIONS
' DPH 26/03/2004 - tidy up objects
'----------------------------------------------------------------------------------------'
Dim oMsg As MIMsg
Dim oDS As MIMsgDS
Dim colSerStates As Collection
Dim oState As MIMsgState
Dim i As Long
   
    Set oDS = New MIMsgDS
    msCon = sCon
    Set mcolMIMsgStates = CollectionDeSerialise(oDS.Load(sCon, lId, nSource, sSite, True))
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
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Save()
Attribute Save.VB_Description = "Save a note"
'----------------------------------------------------------------------------------------'
'Save a note
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
End Sub

'----------------------------------------------------------------------------------------'
Public Function Init(sCon As String, _
                    sText As String, _
                    sUserName As String, sUserNameFull As String, nSource As MIMsgSource, _
                    nScope As MIMsgScope, _
                    sStudyName As String, sSite As String, lSubjectId As Long, _
                    lVisitId As Long, nVisitCycle As Integer, lEFormTaskId As Long, _
                    lResponseTaskId As Long, nResponseCycle As Integer, dblResponseTimeStamp As Double, sResponseValue As String, _
                        lEFormId As Long, nEFormCycle As Integer, lQuestionId As Long, sDataUsername As String, _
                    dblTimestamp As Double, nTimezoneOffset As Integer, nNoteStatus As eNoteMIMStatus) As Integer
'----------------------------------------------------------------------------------------'
' Create a new note.
'Raises mimeInvalidForThisStatus error if already a current row.
' RS 30/09/2002: Added Timestamp & TimezoneOffset parameters
'----------------------------------------------------------------------------------------'
Dim oMsg As MIMsg
Dim oMsgState As MIMsgState
Dim colSerStates As Collection
Dim oState As MIMsgState

    If Not mcolMIMessages Is Nothing Then     'already raised
        Err.Raise MIMsgErrors.mimeInvalidForThisStatus, , "Operation invalid for current status"
    End If

    msCon = sCon
    Set mcolMIMsgStates = New Collection
    Set mcolMIMessages = New Collection
    Init = Add(sText, nNoteStatus, _
                sUserName, sUserNameFull, nSource, _
                nScope, sStudyName, sSite, lSubjectId, lVisitId, nVisitCycle, lEFormTaskId, _
                lResponseTaskId, nResponseCycle, dblResponseTimeStamp, sResponseValue, _
                lEFormId, nEFormCycle, lQuestionId, sDataUsername, _
                dblTimestamp, nTimezoneOffset)
    
End Function



'----------------------------------------------------------------------------------------'
Private Function Add(sText As String, nStatus As eNoteMIMStatus, _
                    sUserName As String, sUserNameFull As String, nSource As MIMsgSource, _
                    nScope As MIMsgScope, _
                    sStudyName As String, sSite As String, lSubjectId As Long, _
                    lVisitId As Long, nVisitCycle As Integer, lEFormTaskId As Long, _
                    lResponseTaskId As Long, nResponseCycle As Integer, dblResponseTimeStamp As Double, sResponseValue As String, _
                        lEFormId As Long, nEFormCycle As Integer, lQuestionId As Long, sDataUsername As String, _
                    dblTimestamp As Double, nTimezoneOffset As Integer) As Integer
'----------------------------------------------------------------------------------------'
' general purpose rotuine to add a note.
'----------------------------------------------------------------------------------------'
Dim oMsg As MIMsg
Dim oMsgState As MIMsgState
    
    If Not moCurrentMsg Is Nothing Then
        'we only have one note per object
        Err.Raise vbObjectError + 1002, , "only one note allowed"
    End If
    
    Set oMsgState = New MIMsgState
    mcolMIMsgStates.Add oMsgState
    Set oMsg = New MIMsg
    oMsg.SetState oMsgState
    Set oMsgState = Nothing
    ' RS 30/09/2002: Added Timestamp & TimezoneOffset parameters
    oMsg.Init sUserName, sUserNameFull, nSource, MIMsgType.mimtNote, nStatus, _
                            nScope, sStudyName, sSite, lSubjectId, lVisitId, nVisitCycle, lEFormTaskId, _
                            lResponseTaskId, nResponseCycle, dblResponseTimeStamp, sResponseValue, _
                            dblTimestamp, nTimezoneOffset, _
                            lEFormId, nEFormCycle, lQuestionId, sDataUsername, _
                            0, 0 ' objectid and source are 0 for notes
                            
    oMsg.Text = sText
    mcolMIMessages.Add oMsg
    Set moCurrentMsg = oMsg
    
    Set oMsg = Nothing
    'return the new index
    Add = mcolMIMessages.Count

End Function

'----------------------------------------------------------------------------------------'
Public Function SetText(sText As String, sUserName As String)
Attribute SetText.VB_Description = "Set the text for the current message row. mimeInvalidUserName,mimeAlreadyProcessed errors can occur."
'----------------------------------------------------------------------------------------'
'Set the text for the current message row.
'mimeInvalidUserName,mimeAlreadySent errors can occur.
'----------------------------------------------------------------------------------------'

    If sUserName <> CurrentMessage.UserName Then
        Err.Raise MIMsgErrors.mimeInvalidUserName, , "This message can only be edited by " & CurrentMessage.UserName
    End If
    If CurrentMessage.TimeSent <> 0 Then
        Err.Raise MIMsgErrors.mimeAlreadySent, , "This message has already been sent"
    End If
    CurrentMessage.Text = sText
    
End Function

'----------------------------------------------------------------------------------------'
Public Property Get CurrentMessage() As MIMsg
Attribute CurrentMessage.VB_Description = "Return the MIMessage object corresponding to this note"
'----------------------------------------------------------------------------------------'
' Return the MIMessage object corresponding to this note
'----------------------------------------------------------------------------------------'
    Set CurrentMessage = mcolMIMessages(1)

End Property


'properties common to the note
Public Property Get MIMessageType() As MIMsgType
    MIMessageType = mimtNote
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

Public Property Get ResponseTimeStamp() As Double
    ResponseTimeStamp = moCurrentMsg.ResponseTimeStamp
End Property

Public Property Get ResponseValue() As String
    ResponseValue = moCurrentMsg.ResponseValue
End Property

 
Public Property Get CurrentStatus() As eNoteMIMStatus
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

