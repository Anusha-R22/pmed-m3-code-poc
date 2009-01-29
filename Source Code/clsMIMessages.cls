VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMIMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsMIMessages.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Nicky Johns, May 2000
'   Purpose:    Collection of clsMIMessage objects class for
'               Bidirectional Communication using the MIMessage table
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 22 May - Initial development
'   NCJ 21 Nov 2000 - Do not use module-level object variables
'   NCJ 16 Oct 02 - Minor changes for consolidation of MIMsg enumerations
'   Mo  25/10/2007  Bug 2950, Add an Order By MIMessageCreated to PopulateCollection, which
'                   is called by frmViewDiscrepancies.cmdPrintDCF_Click and
'                   frmViewDiscrepancies.cmdPrintListing_Click
'----------------------------------------------------------------------------------------'

Option Explicit

' Collection of MIMessage objects
Private mcolMsgs As Collection

'Private moMIMsg As clsMIMessage

' Local variables

Private mnObjectStatus As Integer
Private mnPriority As Integer
Private mnCurrentIndex As Integer

'----------------------------------------------------------------------------------------'
Public Property Get CurrentRecord() As clsMIMessage
'----------------------------------------------------------------------------------------'
' The "current" record in this collection
'----------------------------------------------------------------------------------------'

    If mcolMsgs.Count > 0 Then
        Set CurrentRecord = mcolMsgs.Item(mnCurrentIndex)
    End If

End Property

'----------------------------------------------------------------------------------------'
Public Property Get MessageStatus() As Integer
'----------------------------------------------------------------------------------------'
' The status (SDV status or Discrepancy status) of the collection
' i.e. the status of the "current" record
'----------------------------------------------------------------------------------------'
    
    MessageStatus = mnObjectStatus
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get MessagePriority()
'----------------------------------------------------------------------------------------'
' The Priority of the collection
' i.e. the priority of the "current" record
'----------------------------------------------------------------------------------------'

    MessagePriority = mnPriority
    
End Property

'----------------------------------------------------------------------------------------'
Public Sub PopulateCollection(lMsgId As Long, sSite As String, nSource As Integer)
'----------------------------------------------------------------------------------------'
' Create all the messages that match this message object
' i.e. all the messages which match on ObjectId, ObjectSource and Site
' If ObjectId = 0 (a Message or a Note), collection will only contain one record
'---------------------------------------------------------------------------------
Dim sSQL As String
Dim rsLoadData As ADODB.Recordset
Dim lObjectId As Long
Dim nObjectSource As Integer
Dim nIndex As Long
Dim oMIMsg As clsMIMessage

    On Error GoTo ErrHandler
    
    lObjectId = 0
    ' Reset collection
    Call InitialiseMe
    
    ' Select the initial record
    sSQL = "SELECT MIMessageObjectID, MIMessageObjectSource FROM MIMessage "
    sSQL = sSQL & " WHERE MIMessageID = " & lMsgId
    sSQL = sSQL & " AND MIMessageSite = '" & sSite & "'"
    sSQL = sSQL & " AND MIMessageSource = " & nSource

    Set rsLoadData = New ADODB.Recordset
    rsLoadData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    ' Get its ObjectID and ObjectSource
    If rsLoadData.RecordCount > 0 Then
        lObjectId = rsLoadData!MIMessageObjectId
        nObjectSource = rsLoadData!MIMessageObjectSource
    End If
    rsLoadData.Close
    
    ' Check if ObjectId is zero (for Notes/Messages)
    If lObjectId = 0 Then
        ' Single record only
        Set oMIMsg = New clsMIMessage
        Call oMIMsg.Load(lMsgId, sSite, nSource)
        mcolMsgs.Add oMIMsg
        ' This is the current record so set Status and Priority
        mnPriority = oMIMsg.MessagePriority
        mnObjectStatus = oMIMsg.MessageStatus
        mnCurrentIndex = 1
   Else
        ' Select the matching records for SDV/Discrepancy
        sSQL = "SELECT MIMessageID, MIMessageSource FROM MIMessage "
        sSQL = sSQL & " WHERE MIMessageObjectID = " & lObjectId
        sSQL = sSQL & " AND MIMessageObjectSource = " & nObjectSource
        sSQL = sSQL & " AND MIMessageSite = '" & sSite & "'"
        'Mo 25/10/2007 Bug 2950, add ORDER BY MIMessageCreated
        sSQL = sSQL & " ORDER BY MIMessageCreated"
        Set rsLoadData = New ADODB.Recordset
        rsLoadData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        ' Keep track of the indexes
        nIndex = 1
        Do While Not rsLoadData.EOF
            ' Add each record to the collection
            Set oMIMsg = New clsMIMessage
            Call oMIMsg.Load(rsLoadData!MIMessageID, _
                                sSite, _
                                rsLoadData!MIMessageSource)
            mcolMsgs.Add oMIMsg
            If oMIMsg.MessageHistory = MIMsgHistory.mimhCurrent Then
                ' This is the current record so set Status and Priority
                mnPriority = oMIMsg.MessagePriority
                mnObjectStatus = oMIMsg.MessageStatus
                mnCurrentIndex = nIndex
            End If
            nIndex = nIndex + 1
            rsLoadData.MoveNext
        Loop
        rsLoadData.Close
    End If
    
    Set rsLoadData = Nothing
    Set oMIMsg = Nothing
    
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PopulateCollection", "clsMIMessages")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub AddItem(oMIMsg As clsMIMessage)
'----------------------------------------------------------------------------------------'
' Add a new message object to the collection
'----------------------------------------------------------------------------------------'

    Call mcolMsgs.Add(oMIMsg)
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub RemoveItem(lIndex As Long)
'----------------------------------------------------------------------------------------'
' Remove item from collection
'----------------------------------------------------------------------------------------'

    Call mcolMsgs.Remove(lIndex)
    
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get Item(lIndex As Long) As clsMIMessage
'----------------------------------------------------------------------------------------'
' Get item from collection
'----------------------------------------------------------------------------------------'

    Set Item = mcolMsgs.Item(lIndex)

End Property

'----------------------------------------------------------------------------------------'
Public Function Count() As Integer
'----------------------------------------------------------------------------------------'
' Count of items in collection
'----------------------------------------------------------------------------------------'

    Count = mcolMsgs.Count
    
End Function

'----------------------------------------------------------------------------------------'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
' The enumerator to allow For...Each iterations
'----------------------------------------------------------------------------------------'
    Set NewEnum = mcolMsgs.[_NewEnum]
End Function

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'

    Call InitialiseMe
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'

    Set mcolMsgs = Nothing

End Sub

'----------------------------------------------------------------------------------------'
Private Sub InitialiseMe()
'----------------------------------------------------------------------------------------'
    
    Set mcolMsgs = Nothing
    Set mcolMsgs = New Collection
    mnPriority = 0
    mnObjectStatus = 5
    mnCurrentIndex = 1

End Sub
