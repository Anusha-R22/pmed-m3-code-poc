VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCommunicationsLog 
   Caption         =   "Communications Log"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   9420
      TabIndex        =   13
      Top             =   4980
      Width           =   1125
   End
   Begin MSComctlLib.ListView lvwCommunications 
      Height          =   3735
      Left            =   60
      TabIndex        =   14
      Top             =   1200
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6588
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame fraSelection 
      Caption         =   "Selection Criteria"
      Height          =   1155
      Left            =   60
      TabIndex        =   15
      Top             =   0
      Width           =   10455
      Begin VB.Frame fraOptions 
         Height          =   975
         Left            =   5340
         TabIndex        =   16
         Top             =   120
         Width           =   3615
         Begin VB.OptionButton OptRCBoth 
            Caption         =   "Both"
            Height          =   195
            Left            =   2520
            TabIndex        =   10
            Top             =   720
            Width           =   915
         End
         Begin VB.OptionButton optCreated 
            Caption         =   "Created"
            Height          =   255
            Left            =   2520
            TabIndex        =   9
            Top             =   420
            Width           =   975
         End
         Begin VB.OptionButton optReceived 
            Caption         =   "Received"
            Height          =   255
            Left            =   2520
            TabIndex        =   8
            Top             =   120
            Width           =   1035
         End
         Begin MSMask.MaskEdBox mskToDateTime 
            Height          =   315
            Left            =   1200
            TabIndex        =   7
            Top             =   600
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFromDateTime 
            Height          =   315
            Left            =   1200
            TabIndex        =   6
            Top             =   120
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   " (dd/mm/yyyy)"
            Height          =   195
            Left            =   60
            TabIndex        =   19
            Top             =   420
            Width           =   1095
         End
         Begin VB.Label lblFromDate 
            Alignment       =   1  'Right Justify
            Caption         =   "From Date"
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   180
            Width           =   735
         End
         Begin VB.Label lblTodate 
            Alignment       =   1  'Right Justify
            Caption         =   "To Date"
            Height          =   195
            Left            =   480
            TabIndex        =   17
            Top             =   660
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   345
         Left            =   9240
         TabIndex        =   12
         Top             =   720
         Width           =   1125
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   345
         Left            =   9240
         TabIndex        =   11
         Top             =   300
         Width           =   1125
      End
      Begin VB.CommandButton cmdSite 
         Caption         =   "Site..."
         Height          =   345
         Left            =   3900
         TabIndex        =   5
         Top             =   300
         Width           =   1125
      End
      Begin VB.CommandButton cmdType 
         Caption         =   "Type..."
         Height          =   345
         Left            =   2700
         TabIndex        =   4
         Top             =   300
         Width           =   1125
      End
      Begin VB.CommandButton cmdStatus 
         Caption         =   "Status..."
         Height          =   345
         Left            =   1500
         TabIndex        =   3
         Top             =   300
         Width           =   1125
      End
      Begin VB.OptionButton OptBoth 
         Caption         =   "Both"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   900
         Width           =   795
      End
      Begin VB.OptionButton optOutgoing 
         Caption         =   "Outgoing"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optIncoming 
         Caption         =   "Incoming"
         Height          =   195
         Left            =   240
         TabIndex        =   0
         Top             =   300
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmCommunicationsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       frmCommunicationsLog.frm
'   Author:     Ashitei Trebi-Ollennu, October 2002
'   Purpose:    Monitors Communications between sites.
'------------------------------------------------------------------------------

Option Explicit
Private mColStatus As Collection
Private mColStatusOriginal As Collection
Private mColType As Collection
Private mColTypeOriginal As Collection
Private mColSites As Collection
Private mColSiteOriginal As Collection
Private mColClinicalTrials As Collection
Private Const msDATE_DISPLAY_FORMAT = "yyyy/mm/dd hh:mm:ss"
Private mbResetButtonClicked As Boolean
Private Const msDateMaskDefault = "__/__/____"
Private Const msSetDateMask = "##/##/####"
Private Const msMidnight = ".9999884259"

'------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------
    
    mbResetButtonClicked = False
    Unload Me

End Sub

'-----------------------------------------------------------------------------
Private Sub cmdRefresh_Click()
'-----------------------------------------------------------------------------
'
'-----------------------------------------------------------------------------
    
    LoadListView

End Sub

'---------------------------------------------------------------------------
Private Sub cmdReset_Click()
'---------------------------------------------------------------------------
'
'---------------------------------------------------------------------------
    
    mbResetButtonClicked = True
    mskFromDateTime.Text = msDateMaskDefault
    mskToDateTime.Text = msDateMaskDefault
    Display

End Sub

'-----------------------------------------------------------------------------
Private Sub cmdSite_Click()
'-----------------------------------------------------------------------------
'
'-----------------------------------------------------------------------------
Dim Left As Long
Dim Top As Long
    
    Left = Me.Left + cmdSite.Left + fraSelection.Left
    Top = Me.Top + cmdSite.Top + fraSelection.Top + cmdSite.Height + (Me.Height - Me.ScaleHeight)
    Call frmCommunicationSites.Display(Left, Top, mColSites, mColSiteOriginal)

End Sub

'-----------------------------------------------------------------------------
Private Sub cmdStatus_Click()
'-----------------------------------------------------------------------------
'
'-----------------------------------------------------------------------------
Dim Left As Long
Dim Top As Long
    
    Left = Me.Left + cmdStatus.Left + fraSelection.Left
    Top = Me.Top + cmdStatus.Top + fraSelection.Top + cmdStatus.Height + (Me.Height - Me.ScaleHeight)
    Call frmCommunicationStatus.Display(Left, Top, mColStatus, mColStatusOriginal)
    
    
End Sub

'------------------------------------------------------------------------------
Private Sub cmdType_Click()
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim Left As Long
Dim Top As Long

    Left = Me.Left + cmdType.Left + fraSelection.Left
    Top = Me.Top + cmdType.Top + fraSelection.Top + cmdType.Height + (Me.Height - Me.ScaleHeight)
    Call frmCommunicationType.Display(Left, Top, mColType, mColTypeOriginal)
       
End Sub

'-------------------------------------------------------------------------------
Private Sub Form_Load()
'-------------------------------------------------------------------------------
'
'-------------------------------------------------------------------------------

    Me.Icon = frmMenu.Icon
    mskFromDateTime.Mask = msSetDateMask
    mskToDateTime.Mask = msSetDateMask
    FormCentre Me

End Sub

'-----------------------------------------------------------------------------
Public Sub Display()
'-----------------------------------------------------------------------------
'
'-----------------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    Call AddToStatusCollection
    Call AddToSiteCollection
    Call AddToTypeCollection
    Call BuildTrialCollection
    OptRCBoth.Value = True
    OptBoth.Value = True
    BuildHeaders
   
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.Display"
End Sub

'----------------------------------------------------------------------------
Private Sub LoadListView()
'----------------------------------------------------------------------------
'loads the listview with the data returned from the database tables MIMessage, Message
'----------------------------------------------------------------------------
Dim sSite As String
Dim sStatus As String
Dim sTypeString As String
Dim sType As String
Dim sMIMEssageType As String
Dim j As Integer
Dim vType As Variant
Dim sMsg As String

    On Error GoTo ErrHandler
    
    'inform if  to-date > from-date
    sMsg = "The To date entered is earlier then From date"
    If mskToDateTime.Text <> msDateMaskDefault Then
        If mskFromDateTime.Text <> msDateMaskDefault Then
            If ConvertLocalNumToStandard(CStr(CDbl((CDate(mskToDateTime.Text))))) < _
                ConvertLocalNumToStandard(CStr(CDbl((CDate(mskFromDateTime.Text))))) Then
                Call DialogInformation(sMsg, "Date Error")
                Exit Sub
            End If
        End If
    End If
    
    'clear the list view
    lvwCommunications.ListItems.Clear
    
    'initialise variables
    sType = ""
    sMIMEssageType = ""
    
    'build the selected site(s) for use in SQL
    sSite = MakeSQLSiteString
    
    'build the selected status(es) for use in SQL
    sStatus = MakeSQLStatusString
   
    'build the selected type(s) for use in SQL
    If mColType.Count <> 0 Then
        For j = 1 To mColType.Count
            'if item in collection then it does not belong to MIMessageENUMS
            If mColType.Item(j) <> "20|0" And mColType.Item(j) <> "18|3" And mColType.Item(j) <> "19|2" Then
                If sTypeString = "" Then
                    sTypeString = mColType.Item(j)
                Else
                    sTypeString = sTypeString & "," & mColType.Item(j)
                End If
            Else
                'discrepancy
                If mColType.Item(j) = "20|0" Then
                    If sMIMEssageType = "" Then
                        sMIMEssageType = "0"
                    Else
                        sMIMEssageType = sMIMEssageType & "," & "0"
                    End If
                End If
                'notes
                If mColType.Item(j) = "19|2" Then
                    If sMIMEssageType = "" Then
                        sMIMEssageType = "2"
                    Else
                        sMIMEssageType = sMIMEssageType & "," & "2"
                    End If
                End If
                'SDV
                If mColType.Item(j) = "18|3" Then
                    If sMIMEssageType = "" Then
                        sMIMEssageType = "3"
                    Else
                        sMIMEssageType = sMIMEssageType & "," & "3"
                    End If
                End If
            End If
        Next
    End If
    
    'remove single quotes from end of string
    sType = Replace(sTypeString, Chr(39), "")
    
    'load listview with records from MIMessage Table
    If sMIMEssageType <> "" Then
        Call LoadMIMessageRecords(sMIMEssageType, sSite, sStatus)
    End If
    
    'load listview with records from Message Table
     If sType <> "" Then
        Call LoadMessageRecords(sType, sSite, sStatus)
     End If

    Call lvw_SetAllColWidths(lvwCommunications, LVSCW_AUTOSIZE_USEHEADER)

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.LoadListView"
End Sub

'------------------------------------------------------------------------------------
Private Sub AddToStatusCollection()
'------------------------------------------------------------------------------------
'builds collection for status enums
'this adds the row number(listindex) as item in collection and enum as key the collection
'when new enums are added this will need to be modified
'------------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Set mColStatus = New Collection
    Set mColStatusOriginal = New Collection
    
    mColStatus.Add 2, "0": mColStatusOriginal.Add 2, "0"
    mColStatus.Add 3, "1": mColStatusOriginal.Add 3, "1"
    mColStatus.Add 0, "2": mColStatusOriginal.Add 0, "2"
    mColStatus.Add 1, "3": mColStatusOriginal.Add 1, "3"
    mColStatus.Add 4, "4": mColStatusOriginal.Add 4, "4"
    mColStatus.Add 5, "5": mColStatusOriginal.Add 5, "5"
    mColStatus.Add 6, "6": mColStatusOriginal.Add 6, "6"
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.AddToStatusCollection"
End Sub

'-------------------------------------------------------------------------------------
Private Sub AddToSiteCollection()
'-------------------------------------------------------------------------------------
'builds collection for sites
'adds site to site collections
'-------------------------------------------------------------------------------------
Dim rsSites As ADODB.Recordset
Dim sSQL As String
Dim nNum As Integer

    On Error GoTo ErrHandler

    Set mColSites = New Collection
    Set mColSiteOriginal = New Collection

    sSQL = "SELECT * FROM Site"
    Set rsSites = New ADODB.Recordset
    rsSites.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    If rsSites.RecordCount <= 0 Then Exit Sub
    rsSites.MoveFirst
    Do Until rsSites.EOF
        mColSites.Add rsSites("Site").Value, rsSites("Site").Value
        mColSiteOriginal.Add rsSites("Site").Value, rsSites("Site").Value
        rsSites.MoveNext
    Loop

    Set rsSites = Nothing

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.AddToSiteCollection"
End Sub


'------------------------------------------------------------------------------------
Private Sub AddToTypeCollection()
'------------------------------------------------------------------------------------
'builds collection for type enums
'the items are the enums and the key is the listindex from routine loadlistbox
'------------------------------------------------------------------------------------

On Error GoTo ErrHandler
    
    Set mColType = New Collection
    Set mColTypeOriginal = New Collection
    
    'messagetypes
    mColType.Add "0", "0": mColTypeOriginal.Add "0", "0"
    mColType.Add "1", "1": mColTypeOriginal.Add "1", "1"
    mColType.Add "2", "2": mColTypeOriginal.Add "2", "2"
    mColType.Add "3", "3": mColTypeOriginal.Add "3", "3"
    mColType.Add "4", "4": mColTypeOriginal.Add "4", "4"
    mColType.Add "5", "5": mColTypeOriginal.Add "5", "5"
    mColType.Add "8", "6": mColTypeOriginal.Add "8", "6"
    mColType.Add "10", "7": mColTypeOriginal.Add "10", "7"
    mColType.Add "11", "8": mColTypeOriginal.Add "11", "8"
    mColType.Add "16,17,18,19", "9": mColTypeOriginal.Add "16,17,18,19", "9"
    mColType.Add "20,21,22", "10": mColTypeOriginal.Add "20,21,22", "10"
    mColType.Add "30", "11": mColTypeOriginal.Add "30", "11"
    mColType.Add "31", "12": mColTypeOriginal.Add "31", "12"
    mColType.Add "32", "13": mColTypeOriginal.Add "32", "13"
    mColType.Add "33", "14": mColTypeOriginal.Add "33", "14"
    mColType.Add "34", "15": mColTypeOriginal.Add "34", "15"
    mColType.Add "35", "16": mColTypeOriginal.Add "35", "16"
    mColType.Add "40", "17": mColTypeOriginal.Add "40", "17"
    'mimessagetypes
    mColType.Add "18|3", "18": mColTypeOriginal.Add "18|3", "18"
    mColType.Add "19|2", "19": mColTypeOriginal.Add "19|2", "19"
    mColType.Add "20|0", "20": mColTypeOriginal.Add "20|0", "20"
    'new messagetypes
    mColType.Add "36", "21": mColTypeOriginal.Add "36", "21"
    mColType.Add "37", "22": mColTypeOriginal.Add "37", "22"
    mColType.Add "38", "23": mColTypeOriginal.Add "38", "23"

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.AddToCollection"
End Sub

'--------------------------------------------------------------------------------
Private Sub BuildHeaders()
'--------------------------------------------------------------------------------
'builds column headers for the listview
'--------------------------------------------------------------------------------
Dim colmX As MSComctlLib.ColumnHeader

    On Error GoTo ErrHandler

    'clear listview
    lvwCommunications.ListItems.Clear
    
    'do not rebuild headers when the Reset button is clicked
    If mbResetButtonClicked Then Exit Sub
    
    'add column headers with widths
    Set colmX = lvwCommunications.ColumnHeaders.Add(, , "Study", 1000)
    Set colmX = lvwCommunications.ColumnHeaders.Add(, , "Site", 1000)
    Set colmX = lvwCommunications.ColumnHeaders.Add(, , "Subject", 1000)
    Set colmX = lvwCommunications.ColumnHeaders.Add(, , "Visit", 1000)
    Set colmX = lvwCommunications.ColumnHeaders.Add(, , "eForm", 1000)
    Set colmX = lvwCommunications.ColumnHeaders.Add(, , "Question", 1000)
    Set colmX = lvwCommunications.ColumnHeaders.Add(, , "User", 1000)
    Set colmX = lvwCommunications.ColumnHeaders.Add(, , "Direction", 1000)
    Set colmX = lvwCommunications.ColumnHeaders.Add(, , "Type", 1000)
    Set colmX = lvwCommunications.ColumnHeaders.Add(, , "Transfer Status", 1500)
    Set colmX = lvwCommunications.ColumnHeaders.Add(, , "Created", 1000)
    Set colmX = lvwCommunications.ColumnHeaders.Add(, , "Received", 1000)
    Set colmX = lvwCommunications.ColumnHeaders.Add(, , "Text", 1000)
 
    'set view type
    lvwCommunications.View = lvwReport
    'set initial sort to ascending on column 0 (study)
    lvwCommunications.SortKey = 0
    lvwCommunications.SortOrder = lvwAscending
    
    Me.Show vbModal

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.BuildHeaders"
End Sub

'------------------------------------------------------------------------------

Private Sub LoadMIMessageRecords(ByVal sType As String, _
                                ByVal sSite As String, _
                                ByVal sStatus As String)
'------------------------------------------------------------------------------
'gets records from the MIMEssage table to load the listview
'------------------------------------------------------------------------------
Dim itmX As MSComctlLib.ListItem
Dim rsMIMessages As ADODB.Recordset
Dim sSQL As String
Dim sVist As String

    On Error GoTo ErrHandler
    
    'exit if statuses are not for MIMessages
    If Not IfMISearchCriteria(sStatus, "0") And Not IfMISearchCriteria(sStatus, "1") Then Exit Sub
    
    sSQL = GetMISearchCriteria(sType, sSite, sStatus)
    
    'get selected option button for either Incoming,Outgoing or both
    sSQL = sSQL & GetMITransferOption
    
    'get which dates to filter
    sSQL = sSQL & GetMIDatesCrireria

    Set rsMIMessages = New ADODB.Recordset
    rsMIMessages.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    If rsMIMessages.RecordCount <= 0 Then Exit Sub
    
    lvwCommunications.ListItems.Clear
    rsMIMessages.MoveFirst
    Do Until rsMIMessages.EOF
        
        Set itmX = lvwCommunications.ListItems.Add(, , rsMIMessages!MIMessageTrialName)
            
        If Not IsNull(rsMIMessages!MIMessageSite) Then
            itmX.SubItems(1) = rsMIMessages!MIMessageSite
        End If
        
        If Not IsNull(rsMIMessages!MIMessagePersonId) Then
            itmX.SubItems(2) = rsMIMessages!MIMessagePersonId
        End If
        
        If Not IsNull(rsMIMessages!MIMessageVisitId) Then
            itmX.SubItems(3) = rsMIMessages!MIMessageVisitId
        End If
        
        If Not IsNull(rsMIMessages!MIMessageCRFPageTaskId) Then
            itmX.SubItems(4) = rsMIMessages!MIMessageCRFPageTaskId
        End If
            
        If Not IsNull(rsMIMessages!MIMessageResponseTaskId) Then
            itmX.SubItems(5) = rsMIMessages!MIMessageResponseTaskId
        End If
        
        If Not IsNull(rsMIMessages!MIMessageUserName) Then
            itmX.SubItems(6) = rsMIMessages!MIMessageUserName
        End If
        
        If Not IsNull(rsMIMessages!MIMessageSource) Then
            If rsMIMessages!MIMessageSource = 0 Then
                itmX.SubItems(7) = "Outgoing"
            Else
                itmX.SubItems(7) = "Incoming"
            End If
        End If
            
        If Not IsNull(rsMIMessages!MIMessageType) Then
            itmX.SubItems(8) = GetMIMessageText(rsMIMessages!MIMessageType)
        End If
        
        If rsMIMessages!MIMessageSent <> 0 Then
            itmX.SubItems(9) = "Received"
        Else
            itmX.SubItems(9) = "Not Received"
        End If
            
        If Not IsNull(rsMIMessages!MIMessageCreated) And rsMIMessages!MIMessageCreated <> 0 Then
            itmX.SubItems(10) = Format$(rsMIMessages![MIMessageCreated], msDATE_DISPLAY_FORMAT)
        End If
        
        If Not IsNull(rsMIMessages!MIMessageReceived) And rsMIMessages!MIMessageReceived <> 0 Then
            itmX.SubItems(11) = Format$(rsMIMessages![MIMessageReceived], msDATE_DISPLAY_FORMAT)
        End If
        
        If Not IsNull(rsMIMessages!MIMessageText) Then
            itmX.SubItems(12) = rsMIMessages!MIMessageText
        End If
        
        rsMIMessages.MoveNext
    Loop

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.LoadMIMessageRecords"
End Sub

'------------------------------------------------------------------------------
Private Sub LoadMessageRecords(ByVal sType As String, _
                                ByVal sSite As String, ByVal sStatus As String)
'------------------------------------------------------------------------------
'gets records from the Message table
'------------------------------------------------------------------------------
Dim itmX  As MSComctlLib.ListItem
Dim rsMessages As ADODB.Recordset
Dim sTrialName As String
Dim sSQL As String


    On Error GoTo ErrHandler
    
    'get which search criteria
    sSQL = GetMeSearchCriteria(sType, sSite, sStatus)

    'get selected option button for either Incoming,Outgoing or both
    sSQL = sSQL & GetMTransferOption
   
   'get which dates to filter
    sSQL = sSQL & GetMESDatesCriteria

    Set rsMessages = New ADODB.Recordset
    rsMessages.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    If rsMessages.RecordCount <= 0 Then Exit Sub
    
    rsMessages.MoveFirst
    Do Until rsMessages.EOF
        If Not IsNull(rsMessages!ClinicalTrialId) And rsMessages!ClinicalTrialId <> -1 Then
            sTrialName = mColClinicalTrials.Item(CStr(rsMessages!ClinicalTrialId))
        Else
            sTrialName = ""
        End If
        
        Set itmX = lvwCommunications.ListItems.Add(, , sTrialName)
            
        If Not IsNull(rsMessages!TrialSite) Then
            itmX.SubItems(1) = rsMessages!TrialSite
        End If
        
        If Not IsNull(rsMessages!UserName) Then
            itmX.SubItems(6) = rsMessages!UserName
        End If
        
        If Not IsNull(rsMessages!MessageDirection) Then
            If rsMessages!MessageDirection = 0 Then
                itmX.SubItems(7) = "Outgoing"
            Else
                itmX.SubItems(7) = "Incoming"
            End If
        End If
        
        If Not IsNull(rsMessages!MessageType) Then
            itmX.SubItems(8) = GetMessageTypeText(rsMessages!MessageType)
        End If
        
        If Not IsNull(rsMessages!MessageReceived) Then
            Select Case rsMessages!MessageReceived
            Case MessageReceived.Error
                sStatus = "Error"
            Case MessageReceived.Locked
                sStatus = "Locked"
            Case MessageReceived.NotYetReceived
                sStatus = "Not Received"
            Case MessageReceived.PendingOverRule
                sStatus = "Pending Overrule"
            Case MessageReceived.Received
                sStatus = "Received"
            Case MessageReceived.Skipped
                sStatus = "Skipped"
            Case MessageReceived.Superceeded
                sStatus = "Superseded"
            End Select
            itmX.SubItems(9) = sStatus
        End If
        
        If Not IsNull(rsMessages!MessageTimeStamp) And rsMessages!MessageTimeStamp <> 0 Then
            itmX.SubItems(10) = Format$(rsMessages![MessageTimeStamp], msDATE_DISPLAY_FORMAT)
        End If
            
        If Not IsNull(rsMessages!MessageReceivedTimeStamp) And rsMessages!MessageReceivedTimeStamp <> 0 Then
            itmX.SubItems(11) = Format$(rsMessages![MessageReceivedTimeStamp], msDATE_DISPLAY_FORMAT)
        End If
        
        If Not IsNull(rsMessages!MessageBody) Then
            itmX.SubItems(12) = rsMessages!MessageBody
        End If
        
        rsMessages.MoveNext
    Loop

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.LoadMessageRecords"
End Sub

'-------------------------------------------------------------------------------------
Private Function GetMessageTypeText(ByVal nNum As String) As String
'-------------------------------------------------------------------------------------
'returns the text for a given Message type enum
'-------------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Select Case nNum
        Case Is = "0": GetMessageTypeText = "New Trial"
        Case Is = "1": GetMessageTypeText = "In Preparation"
        Case Is = "2": GetMessageTypeText = "Trial Open"
        Case Is = "3": GetMessageTypeText = "Closed Recruitment"
        Case Is = "4": GetMessageTypeText = "Closed FollowUp"
        Case Is = "5": GetMessageTypeText = "Trial Suspended"
        Case Is = "8": GetMessageTypeText = "New Version"
        Case Is = "10": GetMessageTypeText = "Patient Data"
        Case Is = "11": GetMessageTypeText = "Mail"
        Case Is = "16", "17", "18", "19": GetMessageTypeText = "Locking/Freezing"
        Case Is = "20", "21", "22": GetMessageTypeText = "Unlocking"
        Case Is = "30": GetMessageTypeText = "Lab Definition Server To Site"
        Case Is = "31": GetMessageTypeText = "Lab Definition Site To Server"
        Case Is = "32": GetMessageTypeText = "User"
        Case Is = "33": GetMessageTypeText = "User Role"
        Case Is = "34": GetMessageTypeText = "Password Change"
        Case Is = "35": GetMessageTypeText = "Role"
        Case Is = "36": GetMessageTypeText = "System Log"
        Case Is = "37": GetMessageTypeText = "User Log"
        Case Is = "38": GetMessageTypeText = "Restore User Role"
        Case Is = "40": GetMessageTypeText = "Password Policy"
    End Select

Exit Function:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.GetMessageTypeText"
End Function

'---------------------------------------------------------------------------------
Private Sub BuildTrialCollection()
'---------------------------------------------------------------------------------
'receives study name and or description and returns clinicaltrial ID
'---------------------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
    
    Set mColClinicalTrials = New Collection

    sSQL = "SELECT * FROM ClinicalTrial WHERE ClinicalTrialID > 0"
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    rsTemp.MoveFirst
    Do Until rsTemp.EOF
        mColClinicalTrials.Add rsTemp("ClinicalTrialName").Value, CStr(rsTemp("ClinicalTrialId").Value)
        rsTemp.MoveNext
    Loop
    
    rsTemp.Close
    Set rsTemp = Nothing
        
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.BuildTrialCollection"
End Sub

'-------------------------------------------------------------------------
Private Sub Form_Resize()
'-------------------------------------------------------------------------
'
'-------------------------------------------------------------------------
On Error Resume Next

    If Me.ScaleWidth > 10275 + 120 Then
        fraSelection.Width = Me.ScaleWidth - 120
    Else
        fraSelection.Width = 10275
    End If
    cmdRefresh.Left = fraSelection.Width - cmdRefresh.Width - 120
    cmdReset.Left = fraSelection.Width - cmdReset.Width - 120
    lvwCommunications.Width = Me.ScaleWidth - 120
    cmdOK.Left = Me.ScaleWidth - cmdOK.Width - 60
    
    lvwCommunications.Height = Me.ScaleHeight - fraSelection.Top - fraSelection.Height - cmdOK.Height - 120
    cmdOK.Top = lvwCommunications.Top + lvwCommunications.Height + 60

End Sub

'------------------------------------------------------------------------------------------
Private Function GetMIDatesCrireria() As String
'------------------------------------------------------------------------------------------
'returns the dates from the MIMessage Table to be used in the SQL to load the listview
'------------------------------------------------------------------------------------------
Dim sSQL As String
Dim sToDate As String
Dim sFromDate As String
        
    On Error GoTo ErrHandler
        
    GetMIDatesCrireria = ""
    
    If mskToDateTime.Text <> msDateMaskDefault Then
        sToDate = ConvertLocalNumToStandard(CStr(CDbl((CDate(mskToDateTime.Text)))))
    End If

    If mskFromDateTime.Text <> msDateMaskDefault Then
        sFromDate = ConvertLocalNumToStandard(CStr(CDbl((CDate(mskFromDateTime.Text)))))
    End If
    
    If optReceived Then
        'filter on date received
        'both date fields empty
        If mskToDateTime = msDateMaskDefault And mskFromDateTime = msDateMaskDefault Then
            GetMIDatesCrireria = ""
            Exit Function
        'only from date entered
         ElseIf mskFromDateTime <> msDateMaskDefault And mskToDateTime = msDateMaskDefault Then
            sSQL = sSQL & " AND MIMessageReceived >= " & sFromDate
        'only to date entered
         ElseIf mskFromDateTime = msDateMaskDefault And mskToDateTime <> msDateMaskDefault Then
            sSQL = sSQL & " AND MIMessageReceived <= " & sToDate & msMidnight
        'both dates entered
        Else
            sSQL = sSQL & " AND MIMessageReceived >= " & sFromDate & " AND MIMessageReceived <= " & sToDate & msMidnight
        End If
    ElseIf optCreated Then
        'filter on date created
        'both date fields empty
        If mskToDateTime = msDateMaskDefault And mskFromDateTime = msDateMaskDefault Then
            GetMIDatesCrireria = ""
            Exit Function
        'only from date entered
        ElseIf mskFromDateTime <> msDateMaskDefault And mskToDateTime = msDateMaskDefault Then
            sSQL = sSQL & " AND MIMessageCreated >= " & sFromDate
         'only to date entered
        ElseIf mskFromDateTime = msDateMaskDefault And mskToDateTime <> msDateMaskDefault Then
             sSQL = sSQL & " AND MIMessageCreated <= " & sToDate & msMidnight
         'both dates entered
        Else
            sSQL = sSQL & " AND MIMessageCreated >= " & sFromDate & " AND MIMessageCreated <= " & sToDate & msMidnight
        End If
    Else
        'filter on both
        'both date fields empty
        If mskToDateTime = msDateMaskDefault And mskFromDateTime = msDateMaskDefault Then
            GetMIDatesCrireria = ""
            Exit Function
        'only from date entered
        ElseIf mskFromDateTime <> msDateMaskDefault And mskToDateTime = msDateMaskDefault Then
            sSQL = sSQL & " AND (MIMessageCreated >= " & sFromDate & " AND MIMessageReceived >= " & sFromDate & ")"
        'only to date entered
        ElseIf mskFromDateTime = msDateMaskDefault And mskToDateTime <> msDateMaskDefault Then
             sSQL = sSQL & " AND (MIMessageCreated <= " & sToDate & msMidnight & " AND MIMessageReceived <= " & sToDate & msMidnight & ")"
        'both dates entered
        Else
            sSQL = sSQL & " AND (MIMessageCreated >= " & sFromDate & " AND MIMessageCreated <= " & sToDate & msMidnight
            sSQL = sSQL & " AND MIMessageReceived >= " & sFromDate & " AND MIMessageReceived <= " & sToDate & msMidnight & ")"
        End If
    End If

    GetMIDatesCrireria = sSQL

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.GetMIDatesCrireria"
End Function

'------------------------------------------------------------------------------------------
Private Function GetMESDatesCriteria() As String
'------------------------------------------------------------------------------------------
'returns the dates from the Message Table to be used in the SQL to load the listview
'------------------------------------------------------------------------------------------
Dim sSQL As String
Dim sToDate As String
Dim sFromDate As String
        
    On Error GoTo ErrHandler
    
    GetMESDatesCriteria = ""
        
    If mskToDateTime.Text <> msDateMaskDefault Then
        sToDate = ConvertLocalNumToStandard(CStr(CDbl((CDate(mskToDateTime.Text)))))
    End If
    
    If mskFromDateTime.Text <> msDateMaskDefault Then
        sFromDate = ConvertLocalNumToStandard(CStr(CDbl((CDate(mskFromDateTime.Text)))))
    End If
    
    If optReceived Then
        'filter on date received
        'both date fields empty
        If mskToDateTime = msDateMaskDefault And mskFromDateTime = msDateMaskDefault Then
            GetMESDatesCriteria = ""
            Exit Function
        'only from-date entered
        ElseIf mskFromDateTime <> msDateMaskDefault And mskToDateTime = msDateMaskDefault Then
            sSQL = sSQL & " AND MessageReceivedTimeStamp >= " & sFromDate
        'only to-date entered
        ElseIf mskFromDateTime = msDateMaskDefault And mskToDateTime <> msDateMaskDefault Then
            sSQL = sSQL & " AND MessageReceivedTimeStamp <= " & sToDate & msMidnight
        'both dates entered
        Else
            sSQL = sSQL & " AND MessageReceivedTimeStamp >= " & sFromDate & " AND MessageReceivedTimeStamp <= " & sToDate & msMidnight
        End If
    ElseIf optCreated Then
        'filter on date created
        'both date fields empty
        If mskToDateTime = msDateMaskDefault And mskFromDateTime = msDateMaskDefault Then
            GetMESDatesCriteria = ""
            Exit Function
        'only from-date entered
        ElseIf mskFromDateTime <> msDateMaskDefault And mskToDateTime = msDateMaskDefault Then
            sSQL = sSQL & " AND MessageTimeStamp >= " & sFromDate
        'only to-date entered
        ElseIf mskFromDateTime = msDateMaskDefault And mskToDateTime <> msDateMaskDefault Then
            sSQL = sSQL & " AND MessageTimeStamp <= " & sToDate & msMidnight
        'both dates entered
        Else
            sSQL = sSQL & " AND MessageTimeStamp >= " & sFromDate & " AND MessageTimeStamp <= " & sToDate & msMidnight
        End If
    Else
        'filter on both date created and received
        'both date fields empty
        If mskToDateTime = msDateMaskDefault And mskFromDateTime = msDateMaskDefault Then
            GetMESDatesCriteria = ""
            Exit Function
        'only from-date entered
        ElseIf mskFromDateTime <> msDateMaskDefault And mskToDateTime = msDateMaskDefault Then
            sSQL = sSQL & " AND (MessageTimeStamp >= " & sFromDate & " AND MessageReceivedTimeStamp >= " & sFromDate & ")"
        'only to-date entered
       ElseIf mskFromDateTime = msDateMaskDefault And mskToDateTime <> msDateMaskDefault Then
            sSQL = sSQL & " AND (MessageTimeStamp <= " & sToDate & msMidnight & " AND MessageReceivedTimeStamp <= " & sToDate & msMidnight & ")"
        'both dates entered
        Else
            sSQL = sSQL & " AND (MessageTimeStamp >= " & sFromDate & " AND MessageTimeStamp <= " & sToDate & msMidnight
            sSQL = sSQL & " AND MessageReceivedTimeStamp >= " & sFromDate & " AND MessageReceivedTimeStamp <= " & sToDate & msMidnight & ")"
        End If
    End If

    GetMESDatesCriteria = sSQL

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.GetMESDatesCriteria"
End Function

'-------------------------------------------------------------------------------------
Private Function GetMTransferOption() As String
'-------------------------------------------------------------------------------------
'returns the selected Message Direction option to be used in the SQL to load the
'listview from the Message Table
'-------------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    GetMTransferOption = ""
    
    'filter on incoming transfers
    If optIncoming Then
        sSQL = sSQL & " AND MessageDirection = 1"
    'filter on outgoing transfers
    ElseIf optOutgoing Then
        sSQL = sSQL & " AND MessageDirection = 0"
    Else
    'filter on both outgoing and incoming transfers
        sSQL = sSQL & " AND (MessageDirection = 0 OR MessageDirection = 1)"
    End If
    
    GetMTransferOption = sSQL

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.GetMTransferOption"
End Function

'------------------------------------------------------------------------------------------
Private Function GetMITransferOption()
'------------------------------------------------------------------------------------------
'returns the selected Message Direction option to be used in the SQL to load
'the listview from the MIMessage Table
'------------------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    GetMITransferOption = ""
    
    'filter on incoming transfers
    If optIncoming Then
        sSQL = sSQL & " AND MIMessageSource = 1"
    'filter on outgoing transfers
    ElseIf optOutgoing Then
        sSQL = sSQL & " AND MIMessageSource = 0"
    Else
    'filter on both outgoing and incoming transfers
        sSQL = sSQL & " AND (MIMessageSource = 0 OR MIMessageSource = 1)"
    End If
    
    GetMITransferOption = sSQL

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.GetMITransferOption"
End Function

'------------------------------------------------------------------------
Private Function MakeSQLStatusString() As String
'------------------------------------------------------------------------
'creates and returns the statuses to be used in the SQL to load the listview
'------------------------------------------------------------------------
Dim sStatusString As String
Dim sStatus As String
Dim vStatus As Variant
Dim n As Integer

    On Error GoTo ErrHandler
    
    MakeSQLStatusString = ""
    sStatusString = ""
    sStatus = ""
    
    'build the statuses for sql
    If mColStatus.Count <> 0 Then
        For n = 1 To mColStatus.Count
            If sStatusString = "" Then
                sStatusString = mColStatus.Item(n)
            Else
                sStatusString = sStatusString & "," & mColStatus.Item(n)
            End If
        Next
    End If
    'remove single quotes string
    sStatus = Replace(sStatusString, Chr(39), "")
    
    MakeSQLStatusString = sStatus
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.MakeSQLStatusString"
End Function

'------------------------------------------------------------------------
Private Function MakeSQLSiteString() As String
'------------------------------------------------------------------------
'creates and returns the sites to be used in the SQL to load the listview
'------------------------------------------------------------------------
Dim sSiteString As String
Dim sSite As String
Dim vSite As Variant
Dim i As Integer

    On Error GoTo ErrHandler
    
    MakeSQLSiteString = ""
    
    'build the sites for sql
    If mColSites.Count <> 0 Then
        For i = 1 To mColSites.Count
            If sSiteString = "" Then
                sSiteString = mColSites.Item(i)
            Else
                sSiteString = sSiteString & "," & mColSites.Item(i)
            End If
        Next
    End If
    
    'add quotes around each element for the sql
    vSite = Split(sSiteString, ",")
    For i = 0 To UBound(vSite)
        If sSite = "" Then
            sSite = Chr(39) & vSite(i) & Chr(39)
        Else
            sSite = sSite & "," & Chr(39) & vSite(i) & Chr(39)
        End If
    Next
    
    MakeSQLSiteString = sSite

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.MakeSQLSiteString"
End Function

'------------------------------------------------------------------------------------------
Private Sub lvwCommunications_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'------------------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------------------
On Error GoTo ErrHandler
    
    Call lvw_Sort(lvwCommunications, ColumnHeader)

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.GetMIMessageText"
End Sub

'----------------------------------------------------------
Private Sub mskFromDateTime_LostFocus()
'----------------------------------------------------------
'
'----------------------------------------------------------
Dim sMsg As String

     On Error GoTo ErrHandler

    sMsg = "The date " & mskFromDateTime.Text & " is not a valid date"
    If mskFromDateTime.Text <> msDateMaskDefault Then
        If Not IsDate(mskFromDateTime.Text) Then
            Call DialogInformation(sMsg, "Date Error")
            mskFromDateTime.SetFocus
            Exit Sub
        End If
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.mskFromDateTime_LostFocus"
End Sub

'----------------------------------------------------------
Private Sub mskToDateTime_LostFocus()
'----------------------------------------------------------
'
'----------------------------------------------------------
Dim sMsg As String

    On Error GoTo ErrHandler
    
    sMsg = "The date " & mskToDateTime.Text & " is not a valid date"
    If mskToDateTime.Text <> msDateMaskDefault Then
        If Not IsDate(mskToDateTime.Text) Then
            Call DialogInformation(sMsg, "Date Error")
            mskToDateTime.SetFocus
            Exit Sub
        End If
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.mskToDateTime_LostFocus"
End Sub

'---------------------------------------------------------------------
Private Function GetMIMessageText(ByVal nNum As String) As String
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Select Case nNum
        Case Is = "3": GetMIMessageText = "SDV"
        Case Is = "2": GetMIMessageText = "Note"
        Case Is = "0": GetMIMessageText = "Discrepancy"
    End Select

Exit Function:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.GetMIMessageText"
End Function

'--------------------------------------------------------------------------------
Private Function GetMISearchCriteria(ByVal sType As String, _
                                    ByVal sSite As String, _
                                    ByVal sStatus As String) As String
'--------------------------------------------------------------------------------
'returns string for which criteria to be used in SQL in LoadMIMessageRecords
'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    GetMISearchCriteria = ""
    
    sSQL = "SELECT * FROM MIMessage"
    If Not sType = "" Then
        sSQL = sSQL & " WHERE MIMessageType IN (" & sType & ")"
    End If
    
    If Not sSite = "" Then
        sSQL = sSQL & " AND MIMessageSite IN ( " & sSite & ")"
    End If
    
    'since mimessage statuses are not the same as message statuses, we can only
    'use the mimessagesent column in the mimessage table to determine if a mimessage
    'is received or not yet received.
    'mimessage received is when mimessagesent column = 0
    'mimessage not yet received is when mimessagesent column <> 0
    'filter on both received and not received statuses
    If IfMISearchCriteria(sStatus, "0") And IfMISearchCriteria(sStatus, "1") Then
        sSQL = sSQL & " AND MIMESSAGESENT >= 0"
    'filter on not received status
    ElseIf IfMISearchCriteria(sStatus, "0") Then
        sSQL = sSQL & " AND MIMESSAGESENT = 0"
    'filter on received status
    ElseIf IfMISearchCriteria(sStatus, "1") Then
        sSQL = sSQL & " AND MIMESSAGESENT > 0"
    End If
    
    GetMISearchCriteria = sSQL

Exit Function:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.GetMISearchCriteria"
End Function

'-----------------------------------------------------------------------------------
Private Function GetMeSearchCriteria(ByVal sType As String, _
                                    ByVal sSite As String, _
                                    ByVal sStatus As String) As String
'-----------------------------------------------------------------------------------
'filter criteria for records from Message table.
'-----------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    GetMeSearchCriteria = ""
    
    sSQL = "SELECT * FROM Message"
    If sType <> "" Then
        sSQL = sSQL & " WHERE MessageType IN (" & sType & ")"
    End If
    If sSite <> "" Then
        sSQL = sSQL & " AND TrialSite IN ( " & sSite & ")"
    End If
    If sStatus <> "" Then
        sSQL = sSQL & " AND MESSAGERECEIVED IN ( " & sStatus & ")"
    End If
    
    GetMeSearchCriteria = sSQL
    
Exit Function:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCommunicationsLog.GetMeSearchCriteria"
End Function

'---------------------------------------------------------------------------------------
Private Function IfMISearchCriteria(sStatus As String, sIsStatus As String) As Boolean
'---------------------------------------------------------------------------------------
'finds if the search criteria includes MIMessage statuses.
'---------------------------------------------------------------------------------------
Dim bIsIncluded As Boolean
Dim vStatus As Variant
Dim nI As Integer
    
    bIsIncluded = False
    
    vStatus = Split(sStatus, ",")
    For nI = 0 To UBound(vStatus)
        If vStatus(nI) = sIsStatus Then
            bIsIncluded = True
        End If
    Next
    
    IfMISearchCriteria = bIsIncluded
End Function
