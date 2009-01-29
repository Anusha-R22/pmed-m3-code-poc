VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCommunicationHistory 
   Caption         =   "Communication History"
   ClientHeight    =   5745
   ClientLeft      =   4650
   ClientTop       =   2880
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   7125
   Begin VB.Frame fraOptions 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5220
      Width           =   4335
      Begin VB.OptionButton optSubjMsgs 
         Caption         =   "Subject data transfer"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   180
         Width           =   1995
      End
      Begin VB.OptionButton optSysMsgs 
         Caption         =   "System messages"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5820
      TabIndex        =   0
      Tag             =   "KeepBottomRight"
      Top             =   5280
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwMessages 
      Height          =   5175
      Left            =   60
      TabIndex        =   1
      Tag             =   "Resize"
      Top             =   0
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   9128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Study"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Direction"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmCommunicationHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmCommunicationHistory.frm
'   Author:     Paul Norris 22/07/99
'   Purpose:    To display the communication history recorded in the Message table in Macro.mdb
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   PN  22/09/99    Added moFormResize to handle resizing of window
'   WillC 11/11/99  Added error handler
'   Mo Morris   18/11/99    DAO to ADO conversion
'  WillC   11/12/99
'          Changed the Following where present from Integer to Long  ClinicalTrialId
'           CRFPageId,VisitId,CRFElementID
'   NCJ 20 Jan 00 - SR 2672 Changed "Trial" to "Study"
'   TA 25/04/2000   subclassing removed
'   TA 26/04/2000   Resizing now done within the form
'   Mo Morris   21/6/00 SR 3637, display of Parameters removed from lvwMessages
'   NCJ 1 Sept 03 - Changed history view to show either subjects transferred or system messages
'------------------------------------------------------------------------------------

Option Explicit

' message received constants
Private Const mcRECEIVED = "Received"
Private Const mcNOT_YET_RECEIVED = "Not yet received"

' message direction constants
Private Const mcINCOMING = "Incoming"
Private Const mcOUTGOING = "Outgoing"

'' message type constants
'' NCJ 20/1/00 - Changed "Trial" to "Study"
'Private Const mcNEW_TRIAL = "New study"
'Private Const mcNEW_TRIAL_VERSION = "New version of a study"
'Private Const mcTRIAL_OPENED = "Study opened"
'Private Const mcTRIAL_SUSPENDED = "Study suspended"
'Private Const mcRECRUITMENT_CLOSED = "Study closed to recruitment"
'Private Const mcFOLLOW_UP_CLOSED = "Study closed to follow up"
'Private Const mcPATIENT_DATA_TRANSMISSION = "Subject data transmission"
'Private Const mcMAIL = "Mail"

'form initial height and width
Private mlHeight As Long
Private mlWidth As Long

'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------
' PN  22/09/99 - created
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Me.Height >= mlHeight Then
        lvwMessages.Height = Me.ScaleHeight - cmdClose.Height - 120
        cmdClose.Top = Me.ScaleHeight - cmdClose.Height - 60
        fraOptions.Top = lvwMessages.Top + lvwMessages.Height
    Else
'        Me.Height = mlHeight
    End If
    
    If Me.Width >= mlWidth Then
        lvwMessages.Width = Me.ScaleWidth - 120
        cmdClose.Left = lvwMessages.Left + lvwMessages.Width - cmdClose.Width
    Else
'        Me.Width = mlWidth
    End If
    
ErrHandler:
    ' Do nothing

End Sub

'---------------------------------------------------------------------
Private Sub cmdClose_Click()
'---------------------------------------------------------------------

    Unload Me

End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
  
    On Error GoTo ErrHandler

    Me.Icon = frmMenu.Icon
    
  ' Trigger the loading of Subject messages
    optSubjMsgs.Value = True
    
    Me.BackColor = glFormColour
    
    mlHeight = Me.Height
    mlWidth = Me.Width
    
    FormCentre Me
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Load", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

''---------------------------------------------------------------------
'Private Function GetMessageTypeText(eType As ExchangeMessageType) As String
''---------------------------------------------------------------------
'
'    On Error GoTo ErrHandler
'
'    Select Case eType
'    Case ExchangeMessageType.NewTrial
'        GetMessageTypeText = mcNEW_TRIAL
'
'    Case ExchangeMessageType.NewVersion
'        GetMessageTypeText = mcNEW_TRIAL_VERSION
'
'    Case ExchangeMessageType.TrialOpen
'        GetMessageTypeText = mcTRIAL_OPENED
'
'    Case ExchangeMessageType.TrialSuspended
'        GetMessageTypeText = mcTRIAL_SUSPENDED
'
'    Case ExchangeMessageType.ClosedRecruitment
'        GetMessageTypeText = mcRECRUITMENT_CLOSED
'
'    Case ExchangeMessageType.ClosedFollowUp
'        GetMessageTypeText = mcFOLLOW_UP_CLOSED
'
'    Case ExchangeMessageType.PatientData
'        GetMessageTypeText = mcPATIENT_DATA_TRANSMISSION
'
'    Case ExchangeMessageType.Mail
'        GetMessageTypeText = mcMAIL
'
'    Case ExchangeMessageType.PasswordChange
'        GetMessageTypeText = mcPWD_CHANGE
'
'    End Select
'
'Exit Function
'ErrHandler:
'    Err.Raise Err.Number, , Err.Description & "|frmCommunicationHistory.GetMessageTypeText"
'
'End Function

''---------------------------------------------------------------------
'Private Sub LoadMessages()
''---------------------------------------------------------------------
'Dim rsMessages As ADODB.Recordset
'Dim sSQL As String
'Dim oItem As ListItem
'
'    On Error GoTo ErrHandler
'
'
'    'Changed Mo Morris 21/6/00 SR 3637, display of Parameters removed from lvwMessages
'    'Mo Morris 30/8/01 Db Audit (UserId to UserName)
'    sSQL = "SELECT Message.trialsite, Site.sitedescription, Message.messagetimestamp, ClinicalTrial.ClinicalTrialname, "
'    sSQL = sSQL & "Message.messagedirection, Message.username, Message.messagebody, "
'    sSQL = sSQL & "Message.messagereceived, Message.messagetype "
'    sSQL = sSQL & "FROM Message, ClinicalTrial, Site "
'    sSQL = sSQL & "WHERE ClinicalTrial.ClinicalTrialID = Message.ClinicalTrialID "
'    sSQL = sSQL & "and Site.site = Message.trialsite"
'    Set rsMessages = New ADODB.Recordset
'    rsMessages.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    With rsMessages
'        Do While Not .EOF
'            Set oItem = lvwMessages.ListItems.Add(, , .Fields("TrialSite"))
'            ' Cast MessageTimeStamp as a date - NCJ 2 Dec 99
'            oItem.SubItems(1) = .Fields("sitedescription")
'            'changed Mo Morris 16/6/00
'            If Not IsNull(.Fields("MessageTimeStamp")) Then
'                oItem.SubItems(2) = Format$(CDate(.Fields("MessageTimeStamp")), "yyyy/mm/dd hh:mm:ss")
'            End If
'            'changed Mo Morris 16/6/00, stop displaying library (id = 0) for patient data messages
'            If LCase(.Fields("ClinicalTrialName")) <> "library" Then
'                oItem.SubItems(3) = .Fields("ClinicalTrialName")
'            End If
'            oItem.SubItems(4) = IIf(.Fields("MessageDirection") = MessageIn, mcINCOMING, mcOUTGOING)
'            'changed Mo Morris 16/6/2000, RemoveNull added where nulls might occur
'            'Mo Morris 30/8/01 Db Audit (UserId to UserName)
'            oItem.SubItems(5) = RemoveNull(.Fields("UserName"))
'            oItem.SubItems(6) = RemoveNull(.Fields("MessageBody"))
'            'hanged Mo Morris 26/4/00, enumeration changed from EReceived to Received
'            oItem.SubItems(7) = IIf(.Fields("MessageReceived") = Received, mcRECEIVED, mcNOT_YET_RECEIVED)
'            'changed Mo Morris 21/6/00 SR 3637, display of Parameters removed from lvwMessages
'            'oItem.SubItems(8) = .Fields("MessageParameters")
'            oItem.SubItems(8) = GetMessageTypeText(.Fields("MessageType"))
'
'            .MoveNext
'        Loop
'    End With
'    rsMessages.Close
'    Set rsMessages = Nothing
'
'    ' set all column widths to autoresize
'    Call LVSetAllColWidths(lvwMessages, LVSCW_AUTOSIZE_USEHEADER)
'    Call LVSetStyleEx(lvwMessages, LVSTHeaderDragDrop Or LVSTFullRowSelect, True)
'
'Exit Sub
'ErrHandler:
'    Err.Raise Err.Number, , Err.Description & "|frmCommunicationHistory.LoadMessages"
'
'End Sub

'---------------------------------------------------------------------
Private Sub LoadSystemMessages()
'---------------------------------------------------------------------
Dim rsMessages As ADODB.Recordset
Dim sSQL As String
Dim oItem As ListItem
Dim sTimestamp As String

    On Error GoTo ErrHandler
    
    lvwMessages.ListItems.Clear
    
    lvwMessages.ColumnHeaders(1).Text = "Date"
    lvwMessages.ColumnHeaders(2).Text = "Direction"
    lvwMessages.ColumnHeaders(3).Text = "User"
    lvwMessages.ColumnHeaders(4).Text = "Type"
    lvwMessages.ColumnHeaders(5).Text = "Received"

    'NCJ 1 Sept 03 - Load System message history (ClinicalTrialId = -1)
    sSQL = "SELECT MessageTimeStamp, MessageDirection, Username, MessageBody, "
    sSQL = sSQL & "MessageReceived, Messagetype "
    sSQL = sSQL & "FROM Message "
    sSQL = sSQL & "WHERE Message.ClinicalTrialID = -1 "
    Set rsMessages = New ADODB.Recordset
    rsMessages.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With rsMessages
        Do While Not .EOF
            If Not IsNull(.Fields("MessageTimeStamp")) Then
                sTimestamp = Format$(CDate(.Fields("MessageTimeStamp")), "yyyy/mm/dd hh:mm:ss")
            Else
                sTimestamp = ""
            End If
            Set oItem = lvwMessages.ListItems.Add(, , sTimestamp)
            oItem.SubItems(1) = IIf(.Fields("MessageDirection") = MessageIn, mcINCOMING, mcOUTGOING)
            oItem.SubItems(2) = RemoveNull(.Fields("UserName"))
            oItem.SubItems(3) = RemoveNull(.Fields("MessageBody"))
            oItem.SubItems(4) = IIf(.Fields("MessageReceived") = MessageReceived.Received, mcRECEIVED, mcNOT_YET_RECEIVED)
            
            .MoveNext
        Loop
    End With
    rsMessages.Close
    Set rsMessages = Nothing
    
    ' set all column widths to autoresize
    Call LVSetAllColWidths(lvwMessages, LVSCW_AUTOSIZE_USEHEADER)
    Call LVSetStyleEx(lvwMessages, LVSTHeaderDragDrop Or LVSTFullRowSelect, True)
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationHistory.LoadSystemMessages"
   
End Sub

'---------------------------------------------------------------------
Private Sub LoadSubjectMessages()
'---------------------------------------------------------------------
Dim rsMessages As ADODB.Recordset
Dim sSQL As String
Dim oItem As ListItem
Dim sTimestamp As String

    On Error GoTo ErrHandler
    
    lvwMessages.ListItems.Clear
    
    lvwMessages.ColumnHeaders(1).Text = "Date"
    lvwMessages.ColumnHeaders(2).Text = "User"
    lvwMessages.ColumnHeaders(3).Text = "Details"
    lvwMessages.ColumnHeaders(4).Text = ""
    lvwMessages.ColumnHeaders(5).Text = ""
    
    'NCJ 1 Sept 03 - Load Subject transfer message history (MessageId = 41)
    sSQL = "SELECT MessageTimeStamp, Username, MessageBody "
    sSQL = sSQL & "FROM Message "
    sSQL = sSQL & "WHERE Message.MessageType = " & ExchangeMessageType.PatientDataSent
    Set rsMessages = New ADODB.Recordset
    rsMessages.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With rsMessages
        Do While Not .EOF
            If Not IsNull(.Fields("MessageTimeStamp")) Then
                sTimestamp = Format$(CDate(.Fields("MessageTimeStamp")), "yyyy/mm/dd hh:mm:ss")
            Else
                sTimestamp = ""
            End If
            Set oItem = lvwMessages.ListItems.Add(, , sTimestamp)
            oItem.SubItems(1) = RemoveNull(.Fields("UserName"))
            oItem.SubItems(2) = RemoveNull(.Fields("MessageBody"))
            
            .MoveNext
        Loop
    End With
    rsMessages.Close
    Set rsMessages = Nothing
    
    ' set all column widths to autoresize
    Call LVSetAllColWidths(lvwMessages, LVSCW_AUTOSIZE_USEHEADER)
    Call LVSetStyleEx(lvwMessages, LVSTHeaderDragDrop Or LVSTFullRowSelect, True)
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationHistory.LoadSystemMessages"
   
End Sub

'---------------------------------------------------------------------
Private Sub lvwMessages_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'---------------------------------------------------------------------

    With ColumnHeader
        Call SortList(.Index - 1, Abs(lvwMessages.SortOrder - 1))
    End With

End Sub

'---------------------------------------------------------------------
Private Sub SortList(iSortKey As Integer, iSortOrder As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    With lvwMessages
        .SortKey = iSortKey
        .SortOrder = iSortOrder
        
        Select Case .ColumnHeaders(.SortKey + 1).Text
        Case "Date"
            Call SortListview(Me.lvwMessages, iSortKey, .SortOrder, LVTDate)
        Case Else
            .Sorted = True
        End Select
    End With
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationHistory.SortList"
   
End Sub

'---------------------------------------------------------------------
Private Sub optSubjMsgs_Click()
'---------------------------------------------------------------------
' Display the subject data transfer messages
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    Call LoadSubjectMessages

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "optSubjMsgs_Click", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub optSysMsgs_Click()
'---------------------------------------------------------------------
' Display the system messages
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    Call LoadSystemMessages

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "optSysMsgs_Click", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
End Sub
