VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMIMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsMIMessage.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Nicky Johns, May 2000
'   Purpose:    Message class for Bidirectional Communication
'               using the MIMessage table
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ/WillC 18-22 May - Initial development
'   NCJ 21/11/00 - Added ResponseTimeStamp
' NCJ 1 Mar 01 - Corrected use of ConvertLocalNumToStandard in SQL
' TA 02/10/2002: Removed unused code
' DPH 07/11/2003 - remove val function when handling timestamp in Load
'----------------------------------------------------------------------------------------'

Option Explicit

' Private Values to represent MIMessage table fields
Private mlMsgId As Long
Private msMsgSite As String
Private mnMsgSource As Integer
Private mnMsgType As Integer
Private mnMsgScope As Integer
Private mlMsgObjectID As Long
Private mnMsgObjectSource As Integer
Private mnMsgPriority As Integer
Private msMsgTrialName As String
Private mlMsgPersonId As Long
Private mlMsgVisitId As Long
Private mnMsgVisitCycle As Integer
Private mlMsgCRFPageId As Long
Private mlMsgCRFPageTaskId As Long
Private mlMsgResponseTaskId As Long
Private msMsgResponseValue As String
Private mdblResponseTimeStamp As Double     ' NCJ 21/11/00
Private mlMsgOCDiscrepancyID As Long
Private mdblMsgCreated As Double
Private mdblMsgSent As Double
Private mdblMsgReceived As Double
Private mnMsgHistory As Integer
Private mnMsgProcessed As Integer
Private mnMsgStatus As Integer
Private msMsgText As String
Private msMsgUserName As String
Private msMsgUserNameFull As String
Private mnThisInstallation As Integer


'-----------------------------------------------------
Public Property Get MessageSite() As String
'-----------------------------------------------------
' MIMessageSite - Read-only
'-----------------------------------------------------

    MessageSite = msMsgSite
    
End Property

'-----------------------------------------------------
Public Property Get MessagePriority() As Integer
'-----------------------------------------------------
' MIMessagePriority
'-----------------------------------------------------

    MessagePriority = mnMsgPriority
    
End Property



'-----------------------------------------------------
Public Property Get MessageText() As String
'-----------------------------------------------------
' MIMessageText
'-----------------------------------------------------

    MessageText = msMsgText
    
End Property


'-----------------------------------------------------
Public Property Get MessageTrialName() As String
'-----------------------------------------------------
' MIMessageTrialName
'-----------------------------------------------------

    MessageTrialName = msMsgTrialName
    
End Property


'-----------------------------------------------------
Public Property Get MessagePersonID() As Long
'-----------------------------------------------------
' MIMessagePersonId
'-----------------------------------------------------

    MessagePersonID = mlMsgPersonId
    
End Property



'-----------------------------------------------------
Public Property Get MessageVisitID() As Long
'-----------------------------------------------------
' MIMessageVisitId
'-----------------------------------------------------

    MessageVisitID = mlMsgVisitId
    
End Property



'-----------------------------------------------------
Public Property Get MessageVisitCycle() As Integer
'-----------------------------------------------------
' MIMessageVisitCycle
'-----------------------------------------------------

    MessageVisitCycle = mnMsgVisitCycle
    
End Property



'-----------------------------------------------------
Public Property Get MessageCRFPageTaskId() As Long
'-----------------------------------------------------
' MIMessageCRFPageTaskId
'-----------------------------------------------------

    MessageCRFPageTaskId = mlMsgCRFPageTaskId
    
End Property



'-----------------------------------------------------
Public Property Get MessageResponseTimeStamp() As Double
'-----------------------------------------------------
' NCJ 21 Nov 00
' MIMessageResponseTimeStamp
'-----------------------------------------------------

    MessageResponseTimeStamp = mdblResponseTimeStamp

End Property


'-----------------------------------------------------
Public Property Get MessageResponseTaskId() As Long
'-----------------------------------------------------
' MIMessageResponseTaskId
'-----------------------------------------------------

    MessageResponseTaskId = mlMsgResponseTaskId
    
End Property


'-----------------------------------------------------
Public Property Get MessageResponseValue() As String
'-----------------------------------------------------
' MIMessageResponseValue
'-----------------------------------------------------

    MessageResponseValue = msMsgResponseValue
    
End Property

'-----------------------------------------------------
Public Property Get MessageOCDiscrepancyId() As Long
'-----------------------------------------------------
' MIMessageOCDiscrepancyId
'-----------------------------------------------------

    MessageOCDiscrepancyId = mlMsgOCDiscrepancyID
    
End Property

'-----------------------------------------------------
Public Property Get MessageCreated() As Double
'-----------------------------------------------------
' MIMessageCreated - Read-only
'-----------------------------------------------------

    MessageCreated = mdblMsgCreated
    
End Property

'-----------------------------------------------------
Public Property Get MessageSent() As Double
'-----------------------------------------------------
' MIMessageSent - Read-only
'-----------------------------------------------------

    MessageSent = mdblMsgSent
    
End Property

'-----------------------------------------------------
Public Property Get MessageReceived() As Double
'-----------------------------------------------------
' MIMessageReceived - Read-only
'-----------------------------------------------------

    MessageReceived = mdblMsgReceived
    
End Property

'-----------------------------------------------------
Public Property Get MessageHistory() As Integer
'-----------------------------------------------------
' MIMessageHistory - Read-only
'-----------------------------------------------------

    MessageHistory = mnMsgHistory
    
End Property

'-----------------------------------------------------
Public Property Get MessageProcessed() As Integer
'-----------------------------------------------------
' MIMessageProcessed - Read-only
'-----------------------------------------------------

    MessageProcessed = mnMsgProcessed
    
End Property



'-----------------------------------------------------
Public Property Get MessageStatus() As Integer
'-----------------------------------------------------
' MIMessageStatus
'-----------------------------------------------------

    MessageStatus = mnMsgStatus
    
End Property

'-----------------------------------------------------
Public Property Get MessageUserName() As String
'-----------------------------------------------------
' MIMessageUserCode - Read-only
'-----------------------------------------------------

    MessageUserName = msMsgUserName
    
End Property

'-----------------------------------------------------
Public Property Get MessageUserNameFull() As String
'-----------------------------------------------------
' MIMessageUserName - Read-only
'-----------------------------------------------------

    MessageUserNameFull = msMsgUserNameFull
    
End Property


'-----------------------------------------------------
Private Sub Class_Initialize()
'-----------------------------------------------------
' Initialise a new MIMessage object with default values
'-----------------------------------------------------

    ' Set up where we are
    If gblnRemoteSite Then
        mnThisInstallation = TypeOfInstallation.RemoteSite
    Else
        mnThisInstallation = TypeOfInstallation.Server
    End If
    
End Sub


'-----------------------------------------------------
Public Sub Load(lMessageId As Long, sSite As String, nSource As Integer)
'-----------------------------------------------------
' Load the values from this DB record into the module level variables
' NCJ 21/11/00 - Include ResponseTimeStamp
' DPH 07/11/2003 - remove val function when handling timestamp
'-----------------------------------------------------
Dim sSQL As String
Dim rsLoadData As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    sSQL = " SELECT * FROM MIMessage "
    sSQL = sSQL & " WHERE MIMessageID = " & lMessageId
    sSQL = sSQL & " AND MIMessageSite = '" & sSite & "'"
    sSQL = sSQL & " AND MIMessageSource = " & nSource
    
    Set rsLoadData = New ADODB.Recordset
    
    rsLoadData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    If rsLoadData.RecordCount > 0 Then
        mlMsgId = rsLoadData!MIMessageID
        msMsgSite = rsLoadData!MIMessageSite
        mnMsgSource = rsLoadData!MIMessageSource
        mnMsgType = rsLoadData!MIMessageType
        mnMsgScope = rsLoadData!MIMessageScope
        mlMsgObjectID = rsLoadData!MIMessageObjectId
        mnMsgObjectSource = rsLoadData!MIMessageObjectSource
        mnMsgPriority = rsLoadData!MIMessagePriority
        msMsgTrialName = rsLoadData!MIMessageTrialName
        mlMsgPersonId = rsLoadData!MIMessagePersonId
        mlMsgVisitId = rsLoadData!MIMessageVisitId
        mnMsgVisitCycle = rsLoadData!MIMessageVisitCycle
        mlMsgCRFPageTaskId = rsLoadData!MIMessageCRFPageTaskId
        mlMsgResponseTaskId = rsLoadData!MIMessageResponseTaskId
        ' DPH 07/11/2003 - remove val function when handling timestamp
        ' Val function recognizes only the period (.) as a valid decimal separator
        'mdblResponseTimeStamp = Val(RemoveNull(rsLoadData!MIMessageResponseTimeStamp))
        mdblResponseTimeStamp = IIf(RemoveNull(rsLoadData!MIMessageResponseTimeStamp) <> "", CDbl(RemoveNull(rsLoadData!MIMessageResponseTimeStamp)), 0)
        msMsgResponseValue = RemoveNull(rsLoadData!MIMessageResponseValue)
        mlMsgOCDiscrepancyID = rsLoadData!MIMessageOCDiscrepancyID
        mdblMsgCreated = rsLoadData!MIMessageCreated
        mdblMsgSent = rsLoadData!MIMessageSent
        mdblMsgReceived = rsLoadData!MIMessageReceived
        mnMsgHistory = rsLoadData!MIMessageHistory
        mnMsgProcessed = rsLoadData!MIMessageProcessed
        mnMsgStatus = rsLoadData!MIMessageStatus
        msMsgText = RemoveNull(rsLoadData!MIMessageText)
        msMsgUserName = rsLoadData!MIMessageUserName
        msMsgUserNameFull = RemoveNull(rsLoadData!MIMessageUserNameFull)
    End If
    
    Set rsLoadData = Nothing
 
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Load", "clsMIMessage")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

 
End Sub





