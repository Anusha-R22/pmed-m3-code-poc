VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log  File"
   ClientHeight    =   5910
   ClientLeft      =   645
   ClientTop       =   2595
   ClientWidth     =   14010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5910
   ScaleWidth      =   14010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&OK"
      Height          =   375
      Left            =   12720
      TabIndex        =   13
      Top             =   5460
      Width           =   1215
   End
   Begin VB.CommandButton cmdCulling 
      Caption         =   "&Culling"
      Height          =   372
      Left            =   4020
      TabIndex        =   11
      Top             =   5460
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CommandButton cmdPrinting 
      Caption         =   "&Printing"
      Height          =   372
      Left            =   2700
      TabIndex        =   10
      Top             =   5460
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Frame fraSearchCriteria 
      Caption         =   "Search Criteria"
      Height          =   1100
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   13890
      Begin VB.ComboBox cboTaskId 
         Height          =   315
         Left            =   2220
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtLocation 
         Height          =   315
         Left            =   12480
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskLogDateTime 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton optOnDate 
         Caption         =   "On"
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   800
      End
      Begin VB.OptionButton optAfter 
         Caption         =   "After"
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   720
         Width           =   800
      End
      Begin VB.OptionButton optBefore 
         Caption         =   "Before"
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   800
      End
      Begin VB.TextBox txtLogNumber 
         Height          =   285
         Left            =   13440
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txtUserId 
         Height          =   315
         Left            =   11520
         TabIndex        =   5
         Top             =   480
         Width           =   780
      End
      Begin VB.TextBox txtLogMessage 
         Height          =   315
         Left            =   4920
         TabIndex        =   4
         Top             =   480
         Width           =   6435
      End
      Begin VB.Label Label6 
         Caption         =   "Location"
         Height          =   255
         Left            =   12480
         TabIndex        =   20
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "User Name"
         Height          =   255
         Left            =   11520
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Log message"
         Height          =   255
         Left            =   4920
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Task Id:"
         Height          =   255
         Left            =   2220
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   " (dd/mm/yyyy)"
         Height          =   200
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdRefreshAll 
      Caption         =   "&Refresh All"
      Height          =   372
      Left            =   60
      TabIndex        =   7
      Top             =   5460
      Width           =   1212
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   372
      Left            =   1380
      TabIndex        =   9
      Top             =   5460
      Width           =   1212
   End
   Begin VB.Data datLogDetails 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6660
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5460
      Visible         =   0   'False
      Width           =   4095
   End
   Begin MSComctlLib.ListView lvwLogDetails 
      Height          =   4095
      Left            =   60
      TabIndex        =   6
      Top             =   1260
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   7223
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmLogDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmLogDetails.frm
'   Author:     Mo Morris, July 1997
'   Purpose:    Displays MACRO log file entries
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   1   Mo Morris           4/07/97
'   2   Andrew Newbiggin    1/07/97
'   3   Andrew Newbiggin    10/09/97
'   4   Mo Morris            11/09/97
'   5   Mo Morris            11/09/97
'   6   Mo Morris            18/09/97
'   7   Mo Morris            26/09/97
'   8   Andrew Newbiggin    27/11/97
'       Mo Morris           12/11/98    Adjustments to the form
'   9   PN           14/09/99 Upgrade from DAO to ADO and updated code to conform
'                             to VB standards doc version 1.0
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   PN  01/10/99    Added shortcuts to buttons on this form, and set visibility of
'                   buttons that do nothing (Printing and Culling) to False
''  WillC 23/2/2000  Added Constant the value is the Cdbl(23:59:59) so that we can search accurately
'                  on dates, before dates or after dates
''  WillC 23/2/2000  Changed txtLogDateTime text box to MaskEdit input box to stop spurious date input.
'   TA 08/05/2000   removed subclassing
'   WillC SR3403 30/5/00 sort the listview items by date after a search
'   'WillC 21/6/00 SR3619 allow searches on underscores in the message field
'   NCJ 19/10/00 SR3618 Revisited - date input now works correctly!
'   NCJ 30/10/00 - Added handling for searches for Oracle (only Access and SQL Server were dealt with)
'   ASH 11/06/2002 Bug 2.2.14 no. 24
'   REM 18/10/02 Created display routine and pass in whether to display the LogDetails from the MACRO database
            'or the Login details from the security database
'   REM 30/10/02 Added GMT to the DateTime field and added new field called location
'   REM 31/10/02 - added check for LogDetails, to decide if need LogDetails from MACRO DB or Login Log from Security DB
'                   and added ability to connect to other databases to view their LogDetails table
'   REM 05/11/02 Made taskId text box into a combo box for seaching
'   NCJ 21 Jun 06 - Issue 2745 - Added TaskIDs for study opening/closing
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
Option Explicit
Option Compare Binary
Option Base 0

'Private Const msDATE_DISPLAY_FORMAT = "dd/mm/yyyy hh:mm:ss"
'ASH 11/06/2002
Private Const msDATE_DISPLAY_FORMAT = "yyyy/mm/dd hh:mm:ss"
'WillC 23/2/2000 Added Constant the value is the Cdbl(23:59:59) so that we can search accurately
Private Const msMidnight = ".9999884259"
' Mask for date field
Private Const msDateMaskDefault = "__/__/____"
Private Const msSetDateMask = "##/##/####"

Private mbLogDetails As Boolean
Private msDatabaseCode As String
Private mconMACRO As ADODB.Connection

'---------------------------------------------------------------------
Public Sub Display(bLogDetails As Boolean, Optional sDatabaseCode As String = "")
'---------------------------------------------------------------------
'REM 18/10/02
'Moved load into display routine
'---------------------------------------------------------------------
Dim sSQL As String
Dim oDatabase As MACROUserBS30.Database
Dim sMessage As String
Dim sConnection As String

    On Error GoTo ErrHandler
       
    Me.Icon = frmMenu.Icon
    
    mbLogDetails = bLogDetails
    msDatabaseCode = sDatabaseCode
    
    'REM 31/10/02 - added ability to connect to other databases to view their LogDetails table
    If msDatabaseCode <> "" Then
        Set oDatabase = New MACROUserBS30.Database
        Call oDatabase.Load(SecurityADODBConnection, goUser.UserName, sDatabaseCode, "", False, sMessage)
        sConnection = oDatabase.ConnectionString
        Set mconMACRO = New ADODB.Connection
        mconMACRO.Open sConnection
        mconMACRO.CursorLocation = adUseClient
    End If
    
    'WillC 23/2/2000 changed txtLogDateTime to a mask edit Control to stop spurious input
    mskLogDateTime.Mask = msSetDateMask
    
    'Create the required ColumnHeaders with details
    'Mo Morris 12/10/01, Column width reset since listview control has been upgraded
    'widths changed from 1400,400,100,5750,700 to 1700,700,1500,6000,900
    'REM 30/10/02 - added new field called Location, i.e. for site or server.  This field will always be filled with
    'the word 'Local' until it is transfered back to the server where the field will be filled with the sites name
    With lvwLogDetails
        .ColumnHeaders.Add 1, , "Date & Time", 2800
        .ColumnHeaders.Add , , "LogNo", 700
        .ColumnHeaders.Add , , "Task Id", 1500
        .ColumnHeaders.Add , , "Log Message", 6000
        .ColumnHeaders.Add , , "User Id", 900
        .ColumnHeaders.Add , , "Location", 900
    
        'Set View property to Report
        .View = lvwReport
    
        'Populate the List
        If bLogDetails Then
            sSQL = "SELECT * FROM LogDetails ORDER BY LogDateTime DESC, LogNumber DESC"
        Else
            sSQL = "SELECT * FROM LoginLog ORDER BY LogDateTime DESC, LogNumber DESC"
        End If
        
        ' PN 14/09/99 use new generic PopulateListView() procedure
        ' to load listview
        Call PopulateListView(sSQL)
        
        'REM 05/11/02 - Fill drop down with values
        Call PopulateDropdowns
        
        'Set initial Sort to ascending on column 0 (Date & Time)
        .SortKey = 1
        .SortOrder = lvwAscending
        .Sorted = True
        
        .SortKey = 0
        .SortOrder = lvwAscending
        .Sorted = True
    End With
    
    'Disable the Search Button (until some search criteria has been entered)
    cmdSearch.Enabled = False
    
    'WillC SR3195 sort the listview items by date
    Call SortListview(lvwLogDetails, 0, 1, LVTDate)
    
    Me.Show vbModal
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Form_Load")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub cmdClose_Click()
'---------------------------------------------------------------------
' Close the form
'---------------------------------------------------------------------
 
    Unload Me

End Sub

'---------------------------------------------------------------------
Private Sub lvwLogDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------

    lvwLogDetails.LabelEdit = lvwManual
    
End Sub

'---------------------------------------------------------------------
Private Sub PopulateDropdowns()
'---------------------------------------------------------------------
'REM 05/11/02
'Routine loads all the values into the drop downs for searching
' NCJ 21 Jun 06 - Issue 2745 - Added study opening/closing
'---------------------------------------------------------------------
    
    'clear combo box before loading entries
    cboTaskId.Clear
    
    'if mbLogDetails is true then load up the task ids for the system log
    If mbLogDetails Then
        With cboTaskId
            .AddItem gsAUTOIMPORT
            .AddItem gsAUTO_IMPORT_LLD
            .AddItem gsDEL_TRIAL_PRD
            .AddItem gsDEL_TRIAL_SD
            .AddItem gsIMPORT_SDD
            .AddItem gsEXPORT_SDD
            .AddItem gsIMPORT_PRD
            .AddItem gsEXPORT_PRD
            .AddItem gsEXPORT_PAT_CAB
            .AddItem gsIMPORT_UPGRADE
            .AddItem gsEXPORT_STUDY_CAB
            .AddItem gsIMPORT_STUDY_CAB
            .AddItem gsIMPORT_PAT_CAB
            .AddItem gsAUTO_EXPORT_PRD
            .AddItem gsAUTO_IMPORT_PRD
            .AddItem gsCLEAR_CABEXTR_FOLDER
            .AddItem gsIMPORT_DOC
            .AddItem gsIMPORT_DOC_AND_GRAPHICS
            .AddItem gsEXPORT_LDD
            .AddItem gsEXPORT_LDD_CAB
            .AddItem gsIMPORT_LDD
            .AddItem gsCLEANUP_PRD
            .AddItem gsIMPORT_PAT_ZIP
            .AddItem gsIMPORT_STUDY_ZIP
            .AddItem gsIMPORT_LDD_ZIP
            .AddItem gsEXPORT_PAT_ZIP
            .AddItem gsEXPORT_STUDY_ZIP
            .AddItem gsEXPORT_LDD_ZIP
            .AddItem gsHEX_ENCODE
            .AddItem gsHEX_DECODE
            .AddItem gsVALIDATE_ZIP
            .AddItem gsDOWNLOAD_MESG
            .AddItem gsDOWNLOAD_MIMESG
            .AddItem gsDATA_INTEG_COMMS
            .AddItem gsSYS_TIMEOUT
            .AddItem gsSYSMSG_SEND_ERR
            .AddItem gsSYSMSG_DOWNLOAD_ERR
            .AddItem gsDOWNLOAD_LFMESG
            .AddItem gsREPORT_XFER
            .AddItem gsREPORT_XFER_SITE
            .AddItem gsREPORT_XFER_SERVER
            .AddItem gsREPORT_XFER_ERR
            .AddItem gsCANCEL_TRANSFER
            .AddItem gsCONNECT_FAIL
            .AddItem gsPATDATA_SEND
            .AddItem gsOPEN_TRIAL_SD        ' NCJ 21 Jun 06
            .AddItem gsCLOSE_TRIAL_SD        ' NCJ 21 Jun 06
            .AddItem gsNEW_TRIAL_SD        ' NCJ 21 Jun 06
            .AddItem gsCOPY_TRIAL_SD        ' NCJ 21 Jun 06
        End With
        
    Else 'load up task ids for the user log
    
        With cboTaskId
            .AddItem gsCREATE_DB
            .AddItem gsNEW_ROLE
            .AddItem gsEDIT_ROLE
            .AddItem gsCHANGE_PSWD
            .AddItem gsNEW_USER_ROLE
            .AddItem gsDEL_USER_ROLE
            .AddItem gsUSER_ENABLED
            .AddItem gsUSER_DISABLED
            .AddItem gsUSER_UNLOCKED
            .AddItem gsCHANGE_USERNAME_FULL
            .AddItem gsCREATE_NEW_USER
            .AddItem gsLOGIN
            .AddItem gsLOGOFF
            .AddItem gsCHANGE_SYSADMIN_STATUS
            .AddItem gsUSERNAME_CONFLICT
        End With
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub PopulateListView(sSQL As String)
'---------------------------------------------------------------------
' PN 14/09/99 new procedure to load list view
' this routine loads the listview with the data in the recordset passed in
' it handles locking the screen and setting the mouse pointer
' REM 31/10/02 - added check for LogDetails, to decide if need LogDetails from MACRO DB or Login Log from Security DB
'---------------------------------------------------------------------
Dim rsLogRecords As ADODB.Recordset
Dim oListItem As ListItem
Dim sDate As String
Dim iMousePointer As Integer

    On Error GoTo ErrHandler
    ' save the pointer to replace it later
    iMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    ' load the data
    Set rsLogRecords = New ADODB.Recordset
    'REM 31/10/02 - added check for LogDetails, to decide if need LogDetails from MACRO DB or Login Log from Security DB
    If mbLogDetails Then 'get log details from MACRO database
        If msDatabaseCode = "" Then 'if no database code then current database user is logged into
            rsLogRecords.Open sSQL, MacroADODBConnection, adOpenDynamic, adLockReadOnly, adCmdText
            frmLogDetails.Caption = "Log  File for " & goUser.DatabaseCode
        Else 'else id database that user right clicked on in the tree view
            rsLogRecords.Open sSQL, mconMACRO, adOpenDynamic, adLockReadOnly, adCmdText
            frmLogDetails.Caption = "Log  File for " & msDatabaseCode
        End If
    Else ' get login log details from security database
        rsLogRecords.Open sSQL, SecurityADODBConnection, adOpenDynamic, adLockReadOnly, adCmdText
        frmLogDetails.Caption = "User Log  File"
    End If
    
    ' prevent updates to the screen while loading
    Call LockWindow(Me.lvwLogDetails)
    
 
    With lvwLogDetails
        'Remove existing items from list
        .ListItems.Clear
         
         If rsLogRecords.RecordCount >= -1 Then

        'Populate the List
        Do While Not rsLogRecords.EOF
            'REM 02/09/03 - Added new timezone routine
            sDate = DisplayGMTTime(rsLogRecords![LogDateTime], msDATE_DISPLAY_FORMAT, rsLogRecords![LogDateTime_TZ])
'            sDate = Format$(rsLogRecords![LogDateTime], msDATE_DISPLAY_FORMAT)
'            'REM 30/10/02 - added in time zome to date/time
'            sDate = sDate & " (GMT" & IIf(rsLogRecords![LogDateTime_TZ] < 0, "+", "") & _
'                                -rsLogRecords![LogDateTime_TZ] \ 60 & ":" & _
'                                Format(Abs(rsLogRecords![LogDateTime_TZ]) Mod 60, "00") & ")"
            Set oListItem = .ListItems.Add(, , sDate)
            
            With oListItem
                .SubItems(1) = Right$("00" & RemoveNull(rsLogRecords![LogNumber]), 2)
                .SubItems(2) = RemoveNull(rsLogRecords![TaskId])
                .SubItems(3) = RemoveNull(rsLogRecords![LogMessage])
                'Changed Mo Morris, 12/10/01 Db Audit (UserId to UserName)
                .SubItems(4) = RemoveNull(rsLogRecords![UserName])
                'REM 30/10/02 - added Location field
                .SubItems(5) = RemoveNull(rsLogRecords![Location])
            End With
            rsLogRecords.MoveNext
            
        Loop
        
        ' destroy recordset
        rsLogRecords.Close
        Set rsLogRecords = Nothing
       End If
    End With
    
    
    ' allow the screen to paint
    Call UnlockWindow
    
    'Mo Morris 12/10/01, Call to LVSetStyleEx replaced by setting lvwLogDetails.FullRowSelect to True
    'Call LVSetStyleEx(lvwLogDetails, LVSTFullRowSelect, True)

    ' reset mouse pointer
    Screen.MousePointer = iMousePointer
    
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PopulateListView")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'---------------------------------------------------------------------
Private Sub cmdRefreshAll_Click()
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    ' read the data
    If mbLogDetails Then
        sSQL = "SELECT * FROM LogDetails ORDER BY LogDateTime DESC, LogNumber DESC"
    Else
        sSQL = "SELECT * FROM LoginLog ORDER BY LogDateTime DESC, LogNumber DESC"
    End If
    ' PN 14/09/99 use new generic PopulateListView() procedure
    ' to load listview
    Call PopulateListView(sSQL)
    
    'REM 05/11/02 - Fill drop down with values
    Call PopulateDropdowns
    
    'Clear the search text boxes
    mskLogDateTime.Text = msDateMaskDefault
'    txtLogNumber = vbNullString
    'txtTaskId.Text = vbNullString
    txtLogMessage.Text = vbNullString
    txtUserId.Text = vbNullString
    cboTaskId.Text = vbNullString
    
   'WillC SR3195 10/2/00 sort the listview items by date
    Call SortListview(lvwLogDetails, 0, 1, LVTDate)
    
            
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdRefreshAll_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'---------------------------------------------------------------------
Private Sub cmdSearch_Click()
'---------------------------------------------------------------------
' Refresh the list view based on the selected filters
'---------------------------------------------------------------------
Dim sSQL As String
Dim sSearchLogDateTime As String
'Dim sSearchLogNumber As String
Dim sSearchTaskId As String
Dim sSearchLogMessage As String
Dim sSearchUserId As String
Dim sSearchLocation As String
Dim sDateTime As String
Dim bSecurity As Boolean

'on dates, before dates or after dates

    On Error GoTo ErrHandler

    'Prepare search words for use by the Like Operator
    'The "*" at either end of the user entered search word stands for
    'any number of wildcard characters
    ' PN 14/09/99 replaced * in Like operator with ANSI standard SQL % operator
    
    ' WillC 23/2/2000 SR 2342 added check to see if mask edit date field is empty
    If mskLogDateTime.Text = msDateMaskDefault Then
        ' if date is not filled in set it to WildCards
        sSearchLogDateTime = "'%%'"
    Else
        ' Remove underscores from mask
        sDateTime = Replace(mskLogDateTime.Text, "_", "")
        If gIsDate(sDateTime) Then
            sSearchLogDateTime = ConvertLocalNumToStandard(CStr(CDbl((CDate(sDateTime)))))
            'Replace text with correctly formatted version
            mskLogDateTime.Text = Format(sDateTime, "dd\/mm\/yyyy")
        Else
            Call DialogError("The date " & sDateTime & " is not a valid date.")
            mskLogDateTime.Mask = ""
            mskLogDateTime.Text = ""
            mskLogDateTime.Mask = msSetDateMask
            Exit Sub
        End If
    End If
    
'    sSearchLogNumber = "%" & txtLogNumber.Text & "%"
    sSearchTaskId = Trim(cboTaskId.Text)
    sSearchLogMessage = "%" & Trim(txtLogMessage.Text) & "%"
    sSearchUserId = "%" & Trim(txtUserId.Text) & "%"
    sSearchLocation = "%" & Trim(txtLocation.Text) & "%"
    
    'WillC 21/6/00 SR3619
    If InStr(1, sSearchLogMessage, "_") Then
        sSearchLogMessage = Replace(sSearchLogMessage, "_", "[_]")
    End If
    'Create SQL statement to perform the Search based on contents of search text boxes
    'Changed by Mo Morris 1/12/99
    'Note that access is forgiving about a data item's type, where as SQL Server is
    'unforgiving and requires the data items to be converted to varchars
    
    If mbLogDetails Then
    
        sSQL = "SELECT * FROM LogDetails"
    Else
        sSQL = "SELECT * FROM LoginLog"
    End If
        ' Allow for the search without date ie wildcards
     If sSearchLogDateTime = "'%%'" Then
         sSQL = sSQL & " WHERE LogDateTime Like '%%'"
     Else
         If optBefore.Value = True Then
             sSQL = sSQL & " WHERE LogDateTime <= " & sSearchLogDateTime
         ElseIf optOnDate.Value = True Then           ' Date at 00:00:00                     and Add the Constant for midnight
             sSQL = sSQL & " WHERE LogDateTime > " & sSearchLogDateTime & " AND LogDateTime < " & sSearchLogDateTime & msMidnight
         Else 'optAfter.value = true
             sSQL = sSQL & " WHERE LogDateTime >= " & sSearchLogDateTime & msMidnight
         End If
    End If
    
    If Not mbLogDetails Then
        bSecurity = True
    Else
        bSecurity = False
    End If
    
    ' NCJ 30/10/00 - Use new GetSQLStringLike
    ' REM 17/02/03 - don't want to do LIKE search on TaskId
    'sSQL = sSQL & " AND " & GetSQLStringLike("TaskId", sSearchTaskId)
    If sSearchTaskId <> "" Then
        sSQL = sSQL & " AND TaskId ='" & sSearchTaskId & "'"
    End If
    sSQL = sSQL & " AND " & GetSQLStringLike("LogMessage", sSearchLogMessage, bSecurity)
    'Changed Mo Morris, 12/10/01 Db Audit (UserId to UserName)
    sSQL = sSQL & " AND " & GetSQLStringLike("UserName", sSearchUserId, bSecurity)
    sSQL = sSQL & " AND " & GetSQLStringLike("Location", sSearchLocation, bSecurity)
    sSQL = sSQL & " ORDER BY LogDateTime DESC, LogNumber DESC"
    
'    Select Case gUser.DatabaseType
'    Case 0
'        'Access database
''        sSQL = sSQL & " AND LogNumber Like '" & sSearchLogNumber & "'"
'        sSQL = sSQL & " AND TaskId Like '" & sSearchTaskId & "'"
'        sSQL = sSQL & " AND LogMessage Like '" & sSearchLogMessage & "'"
'        sSQL = sSQL & " AND UserId Like '" & sSearchUserId & "'"
'        sSQL = sSQL & " ORDER BY LogDateTime DESC, LogNumber DESC"
'
'    Case 1
'        'SQL Server database
''        sSQL = sSQL & " AND CONVERT(VARCHAR,LogNumber) Like '" & sSearchLogNumber & "'"
'        sSQL = sSQL & " AND CONVERT(VARCHAR,TaskId) Like '" & sSearchTaskId & "'"
'        sSQL = sSQL & " AND CONVERT(VARCHAR,LogMessage) Like '" & sSearchLogMessage & "'"
'        sSQL = sSQL & " AND CONVERT(VARCHAR,UserId) Like '" & sSearchUserId & "'"
'        sSQL = sSQL & " ORDER BY LogDateTime DESC, LogNumber DESC"
'    End Select
        
    ' PN 14/09/99 use new generic PopulateListView() procedure
    ' to load listview
    Call PopulateListView(sSQL)
    
   'WillC SR3403 30/5/00 sort the listview items by date
    Call SortListview(lvwLogDetails, 0, 1, LVTDate)
    
    cmdSearch.Enabled = False
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdSearch_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------

   
End Sub

'---------------------------------------------------------------------
Private Sub lvwLogDetails_ColumnClick(ByVal ColumnHeader As ColumnHeader)
'---------------------------------------------------------------------

Dim lNewSortOrder As Long

 On Error GoTo ErrHandler
    With lvwLogDetails
        'Set the Sortkey to the ColumnHeader item that has been clicked
        .SortKey = ColumnHeader.Index - 1
        lNewSortOrder = Abs(.SortOrder - 1)
        
        Select Case ColumnHeader.Text
        Case "Date & Time"
            ' special sorting for date columns
            Call SortListview(lvwLogDetails, .SortKey, lNewSortOrder, LVTDate)
        
        Case "LogNo", "Task Id", "User Id"
            ' special sorting for numeric columns
            Call SortListview(lvwLogDetails, .SortKey, lNewSortOrder, LVTNumber)
        
        Case Else
            ' sorting for all other columns
            .SortOrder = lNewSortOrder
            .Sorted = True
            
        End Select
        
    End With
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lvwLogDetails_ColumnClick")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'---------------------------------------------------------------------
Private Sub optAfter_Click()
'---------------------------------------------------------------------
' Make sure theres something to search on
'---------------------------------------------------------------------

   Call CheckSearchText
   
End Sub

'---------------------------------------------------------------------
Private Sub optBefore_Click()
'---------------------------------------------------------------------
' Make sure theres something to search on
'---------------------------------------------------------------------

   Call CheckSearchText

End Sub

'---------------------------------------------------------------------
Private Sub optOnDate_Click()
'---------------------------------------------------------------------
' Make sure theres something to search on
'---------------------------------------------------------------------

   Call CheckSearchText

End Sub

'---------------------------------------------------------------------
Private Sub mskLogDateTime_Change()
'---------------------------------------------------------------------
    
    Call CheckSearchText
   
End Sub

'---------------------------------------------------------------------
Private Sub txtLogMessage_Change()
'---------------------------------------------------------------------
    
    Call CheckSearchText
    
End Sub

''---------------------------------------------------------------------
'Private Sub txtLogNumber_Change()
''---------------------------------------------------------------------
'
'    Call CheckSearchText
'
'End Sub

'---------------------------------------------------------------------
Private Sub txtLocation_Change()
'---------------------------------------------------------------------

    Call CheckSearchText
    
End Sub

'---------------------------------------------------------------------
Private Sub cboTaskId_Change()
'---------------------------------------------------------------------
    
    Call CheckSearchText
    
End Sub

'---------------------------------------------------------------------
Private Sub cboTaskId_Click()
'---------------------------------------------------------------------

    Call CheckSearchText
    
End Sub

'---------------------------------------------------------------------
Private Sub txtUserId_Change()
'---------------------------------------------------------------------
 On Error GoTo ErrHandler
    
    Call CheckSearchText
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtUserId_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'---------------------------------------------------------------------
Private Sub CheckSearchText()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
 
    If mskLogDateTime.Text = msDateMaskDefault And _
       cboTaskId.Text = vbNullString And _
       txtLogMessage.Text = vbNullString And _
       txtUserId.Text = vbNullString And _
       txtLocation.Text = vbNullString Then
        cmdSearch.Enabled = False
    Else
        cmdSearch.Enabled = True
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "CheckSearchText")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

