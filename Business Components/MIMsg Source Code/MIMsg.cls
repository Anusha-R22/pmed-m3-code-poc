VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: MIMsg.cls
' Toby Aldridge, InferMed, Nov 2001
' Generic MIMessage class for MACRO
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' TA Nov 2001 - Initial development
'   NCJ 15 Oct 02 - Added ScopeText
' TA 18/11/2002: Added new fields
' NCJ 20-22 Nov 02 - Added (then removed) Freezer message type
' DPH 26/03/2004 - Tidy up on class termination
' ic 27/02/2007 issue 2114, added GMT to timestamps
'----------------------------------------------------

Option Explicit

'List of MIMessage state object properties for reference
'MIMessageId
'MIMessageSite
'MIMessageSource
'MIMessageType
'MIMessageScope
'MIMessageObjectID
'MIMessageObjectSource
'MIMessagePriority
'MIMessageTrialName
'MIMessagePersonId
'MIMessageVisitId
'MIMessageVisitCycle
'MIMessageCRFPageTaskID
'MIMessageResponseTaskId
'MIMessageResponseValue
'MIMessageOCDiscrepancyID
'MIMessageCreated
'MIMessageSent
'MIMessageReceived
'MIMessageHistory
'MIMessageProcessed
'MIMessageStatus
'MIMessageText
'MIMessageUserCode
'MIMessageUserName
'MIMessageResponseTimeStamp
'MIMessageResponseCycle

'MIMessage type enum
Public Enum MIMsgType
    mimtDiscrepancy = 0
    mimtNote = 2
    mimtSDVMark = 3
End Enum

'message rpocessed enum
Public Enum MIMsgProcessed
    mimpProcessed = 1
    mimpUnprocessed = 0
End Enum

' msg scope enum
Public Enum MIMsgScope
    mimscStudy = 0
    mimscSubject = 1
    mimscVisit = 2
    mimscEForm = 3
    mimscQuestion = 4
End Enum
        
' msg source enum
Public Enum MIMsgSource
    mimsServer = 0
    mimsSite = 1
End Enum

'message history enum
Public Enum MIMsgHistory
    mimhCurrent = 0
    mimhNotCurrent = 1
End Enum
        
       
Private Const m_UNKNOWN_VALUE = -1
        
Private moState As MIMsgState

'---------------------------------------------------------------------
Friend Sub SetState(oState As MIMsgState)
'---------------------------------------------------------------------
'set this message's state object
'---------------------------------------------------------------------

    Set moState = oState

End Sub

'---------------------------------------------------------------------
Public Property Get Key() As String
Attribute Key.VB_Description = "return a unique key string for a msg"
'---------------------------------------------------------------------
' return a unique key string for a msg
'---------------------------------------------------------------------

    Key = moState.Key
    
End Property

'---------------------------------------------------------------------
Friend Sub Init(sUserName As String, sUserNameFull As String, _
                        nSource As MIMsgSource, nType As MIMsgType, _
                        ByVal nStatus As Integer, _
                        nScope As MIMsgScope, _
                        sStudyName As String, sSite As String, lSubjectId As Long, _
                        lVisitId As Long, nVisitCycle As Integer, lEFormTaskId As Long, _
                        lResponseTaskId As Long, nResponseCycle As Integer, dblResponseTimeStamp As Double, sResponseValue As String, _
                        dblTimestamp As Double, nTimezoneOffset As Integer, _
                        lEFormId As Long, nEFormCycle As Integer, lQuestionId As Long, sDataUsername As String, _
                        Optional lObjectId As Long = m_UNKNOWN_VALUE, Optional nObjectSource As Integer = m_UNKNOWN_VALUE)
'---------------------------------------------------------------------
' create a new message row.
' RS 30/09/2002 Added Timestamp & TimezoneOffset parameters
'---------------------------------------------------------------------

    With moState
        .UserName = sUserName
        .UserNameFull = sUserNameFull
        .MIMessageType = nType
        .Status = nStatus
        
        'scope
        .Scope = nScope
        .StudyName = sStudyName
        .Site = sSite
        .SubjectId = lSubjectId
        .VisitId = lVisitId
        .VisitCycle = nVisitCycle
        .EFormTaskId = lEFormTaskId
        .ResponseTaskId = lResponseTaskId
        .ResponseCycle = nResponseCycle
        .ResponseTimeStamp = dblResponseTimeStamp
        .ResponseValue = sResponseValue
    
        'TA 18/11/2002: New fields
        .EFormId = lEFormId
        .EFormCycle = nEFormCycle
        .QuestionId = lQuestionId
        .DataUserName = sDataUsername
        
        'the object identifier or -1 if not known
        .ObjectId = lObjectId
        .ObjectSource = nObjectSource
        
        'identifier to be calculated when saved
        .Id = m_UNKNOWN_VALUE
        .Source = nSource

        ' RS 30/09/2002: No longer use timestamp from machine, but value passed in parameter
        ' .TimeCreated = CDbl(Now)
        .TimeCreated = dblTimestamp
        .TimeCreatedTimezoneOffset = nTimezoneOffset
        
        'fill in default values
        .TimeReceived = 0
        .TimeSent = 0
        ' RS 30/09/2002: Default values for timezoneoffset
        .TimeReceivedTimezoneOffset = 0
        .TimeSentTimezoneOffset = 0
        
        'will be current (newest) record
        .History = MIMsgHistory.mimhCurrent
        'not responded to
        .Processed = MIMsgProcessed.mimpProcessed

        .IsNew = True
    End With
        
                    
End Sub

'---------------------------------------------------------------------
Public Property Get StatusText() As String
'---------------------------------------------------------------------
' The status of the MIMessage as a string
'---------------------------------------------------------------------

    StatusText = GetStatusText(MIMessageType, Status)

End Property

'---------------------------------------------------------------------
Public Property Get ScopeText() As String
'---------------------------------------------------------------------
' The scope of the MIMessage as a string
'---------------------------------------------------------------------

    ScopeText = GetScopeText(Scope)

End Property

'---------------------------------------------------------------------
Friend Property Let Priority(nPriority As Integer)
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    If nPriority < 1 Or nPriority > 10 Then
        Err.Raise mimeInvalidPriority, , "Invalid prioirty"
    End If
    
    moState.Priority = nPriority
    moState.IsChanged = True

End Property

'---------------------------------------------------------------------
Friend Property Let ExternalId(lExternalId As Long)
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    moState.ExternalId = lExternalId
    moState.IsChanged = True

End Property

'---------------------------------------------------------------------
Friend Property Let Text(sText As String)
'---------------------------------------------------------------------

'---------------------------------------------------------------------
Dim sValidationChars As String
Dim nIndex As Integer
    'note that Chr(34)= "
    'note that Chr(124)= |
    'note that Chr(126)= ~
    'note that the ! character is not been checked for,
    'it is here to reverse the manner in which the Like function operates
    
    If Len(sText) > 2000 Then
        Err.Raise mimeInvalidText, , "Message text may not more than 2000 characters."
    End If
    sValidationChars = "[!`" & Chr(34) & Chr(124) & Chr(126) & "]"
    For nIndex = 1 To Len(sText)
        If Not Mid(sText, nIndex, 1) Like sValidationChars Then
            Err.Raise mimeInvalidText, , "Message text may not contain double or backward quotes, tildes or the | character."
        End If
    Next

    moState.Text = sText
    moState.IsChanged = True
    
End Property

'---------------------------------------------------------------------
Friend Property Let History(nHistory As MIMsgHistory)
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    moState.History = nHistory
    moState.IsChanged = True
    
End Property

'---------------------------------------------------------------------
Friend Property Let Status(nStatus As Integer)
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    moState.Status = nStatus
    moState.IsChanged = True
End Property

'---------------------------------------------------------------------
Friend Property Let Processed(nProcessed As MIMsgProcessed)
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    moState.Processed = nProcessed
    moState.IsChanged = True
    
End Property

'---------------------------------------------------------------------
Friend Property Let TimeSent(dblTime As Double)
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    moState.TimeSent = dblTime
    moState.IsChanged = True
    
End Property

'general property gets - just passed on to state object

Public Property Get Id() As Long
    Id = moState.Id
End Property

Public Property Get Source() As MIMsgSource
    Source = moState.Source
End Property

Public Property Get Priority() As Integer
    Priority = moState.Priority
End Property

Public Property Get TimeCreated() As Double
    TimeCreated = moState.TimeCreated
End Property

Public Property Get TimeSent() As Double
    TimeSent = moState.TimeSent
End Property

Public Property Get TimeReceived() As Double
    TimeReceived = moState.TimeReceived
End Property

Public Property Get History() As MIMsgHistory
    History = moState.History
End Property
Public Property Get Processed() As MIMsgProcessed
    Processed = moState.Processed
End Property

Public Property Get Status() As Integer
    Status = moState.Status
End Property

Public Property Get Text() As String
    Text = moState.Text
End Property

Public Property Get UserName() As String
    UserName = moState.UserName
End Property

Public Property Get UserNameFull() As String
    UserNameFull = moState.UserNameFull
End Property

Public Property Get ResponseValue() As String
    ResponseValue = moState.ResponseValue
End Property

Public Property Get ResponseTimeStamp() As Double
    ResponseTimeStamp = moState.ResponseTimeStamp
End Property

'friend properties - propeties that this message that are identical to all other messages in the set

Friend Property Get MIMessageType() As MIMsgType
    MIMessageType = moState.MIMessageType
End Property

Friend Property Get Scope() As MIMsgScope
    Scope = moState.Scope
End Property

Friend Property Get MIMessageId() As Long
    MIMessageId = moState.ObjectId
End Property

Friend Property Get MIMessagesSource() As MIMsgSource
    MIMessagesSource = moState.ObjectSource
End Property

Friend Property Get StudyName() As String
    StudyName = moState.StudyName
End Property

Friend Property Get Site() As String
    Site = moState.Site
End Property

Friend Property Get SubjectId() As Long
    SubjectId = moState.SubjectId
End Property

Friend Property Get VisitId() As Long
    VisitId = moState.VisitId
End Property

Friend Property Get VisitCycle() As Integer
    VisitCycle = moState.VisitCycle
End Property

Friend Property Get EFormTaskId() As Long
    EFormTaskId = moState.EFormTaskId
End Property

Friend Property Get ResponseTaskId() As Long
    ResponseTaskId = moState.ResponseTaskId
End Property

Friend Property Get ResponseCycle() As Integer
    ResponseCycle = moState.ResponseCycle
End Property

Friend Property Get ExternalId() As Long
    ExternalId = moState.ExternalId
End Property


Friend Property Get EFormId() As Long
    EFormId = moState.EFormId
End Property
 
Friend Property Get EFormCycle() As Integer
    EFormCycle = moState.EFormCycle
End Property

Friend Property Get QuestionId() As Long
    QuestionId = moState.QuestionId
End Property

Friend Property Get DataUserName() As String
    DataUserName = moState.DataUserName
End Property

'ic 27/02/2007 issue 2114, added GMT to timestamps
Public Property Get TimeCreatedTimezoneOffset() As Integer
    TimeCreatedTimezoneOffset = moState.TimeCreatedTimezoneOffset
End Property

'----------------------------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------------------------
' DPH 26/03/2004 - Call terminate
'----------------------------------------------------------------------

    Call Terminate

End Sub

'----------------------------------------------------------------------
Friend Sub Terminate()
'----------------------------------------------------------------------
' DPH 26/03/2004 - Tidy up on termination - remove state object
'----------------------------------------------------------------------

    Set moState = Nothing

End Sub
