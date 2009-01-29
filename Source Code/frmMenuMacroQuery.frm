VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMenu 
   Caption         =   "MACRO Query Module"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   -735
   ClientWidth     =   11355
   Icon            =   "frmMenuMacroQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   11355
   Begin VB.PictureBox SelectFilterBar 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   100
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   11175
      TabIndex        =   11
      ToolTipText     =   "Slide bar between Select and Filter area"
      Top             =   2700
      Width           =   11235
   End
   Begin VB.PictureBox FilterOutputBar 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   100
      Left            =   100
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   11175
      TabIndex        =   12
      ToolTipText     =   "Slide bar between Filter and Output area"
      Top             =   4500
      Width           =   11235
   End
   Begin VB.Frame fraDisplayOptions 
      Caption         =   "Display Options"
      Height          =   815
      Left            =   1440
      TabIndex        =   40
      Top             =   4600
      Width           =   2000
      Begin VB.OptionButton optDoNotDisplayOutput 
         Caption         =   "Do not Display Output"
         Height          =   255
         Left            =   60
         TabIndex        =   42
         Top             =   500
         Width           =   1860
      End
      Begin VB.OptionButton optDisplayOutput 
         Caption         =   "Display Output"
         Height          =   195
         Left            =   60
         TabIndex        =   41
         Top             =   240
         Width           =   1800
      End
   End
   Begin VB.CommandButton cmdCancelRun 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   8760
      TabIndex        =   39
      Top             =   4680
      Width           =   1200
   End
   Begin VB.TextBox txtProgress 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7380
      TabIndex        =   37
      Top             =   5040
      Width           =   3945
   End
   Begin VB.CommandButton cmdSaveOutPut 
      Caption         =   "Save Output"
      Height          =   315
      Left            =   6120
      TabIndex        =   36
      Top             =   4680
      Width           =   1200
   End
   Begin MSDataGridLib.DataGrid grdOutPut 
      Height          =   2295
      Left            =   120
      TabIndex        =   35
      Top             =   5460
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboCatCodes 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   4080
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   4620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtBandType 
      Height          =   315
      Left            =   5940
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   2880
      Width           =   5475
   End
   Begin VB.Frame fraBandType 
      Caption         =   "Band Type"
      Height          =   735
      Left            =   120
      TabIndex        =   30
      Top             =   3660
      Width           =   1250
      Begin VB.OptionButton optORBands 
         Caption         =   "OR-Bands"
         Height          =   255
         Left            =   60
         TabIndex        =   32
         Top             =   450
         Width           =   1100
      End
      Begin VB.OptionButton optAndBands 
         Caption         =   "AND-Bands"
         Height          =   195
         Left            =   60
         TabIndex        =   31
         Top             =   240
         Width           =   1150
      End
   End
   Begin VB.TextBox txtBandNo 
      Height          =   315
      Left            =   960
      TabIndex        =   29
      Top             =   3300
      Width           =   400
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4920
      TabIndex        =   27
      Top             =   2880
      Width           =   800
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   315
      Left            =   1500
      TabIndex        =   26
      Top             =   2880
      Width           =   800
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   315
      Left            =   4080
      TabIndex        =   22
      Top             =   2880
      Width           =   800
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   315
      Left            =   3180
      TabIndex        =   21
      Top             =   2880
      Width           =   800
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   2400
      TabIndex        =   20
      Top             =   2880
      Width           =   800
   End
   Begin VB.TextBox txtCriteria 
      Height          =   315
      Left            =   2280
      TabIndex        =   19
      Top             =   4080
      Width           =   3550
   End
   Begin VB.ComboBox cboOperator 
      Height          =   315
      Left            =   2250
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3705
      Width           =   3555
   End
   Begin VB.ComboBox cboOperand 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3300
      Width           =   3550
   End
   Begin VB.ListBox lstFilterText 
      Height          =   1035
      Left            =   5880
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   3300
      Width           =   5475
   End
   Begin VB.TextBox txtRecordsCount 
      Height          =   315
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   5040
      Width           =   800
   End
   Begin VB.TextBox txtResultsCount 
      Height          =   315
      Left            =   5145
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4620
      Width           =   800
   End
   Begin VB.CommandButton cmdRunQuery 
      Caption         =   "Run Query"
      Height          =   315
      Left            =   7440
      TabIndex        =   6
      Top             =   4680
      Width           =   1200
   End
   Begin MSComctlLib.ImageList imgTreeViewIcons16 
      Left            =   6600
      Top             =   4500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMacroQuery.frx":08CA
            Key             =   "unticked"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMacroQuery.frx":0A24
            Key             =   "ticked"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboStudies 
      Height          =   315
      Left            =   2300
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   100
      Width           =   3495
   End
   Begin VB.TextBox txtQueryText 
      Height          =   2500
      Left            =   5880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   100
      Width           =   5400
   End
   Begin MSComctlLib.TreeView trwQuestions 
      Height          =   2100
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3704
      _Version        =   393217
      Indentation     =   706
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Timer tmrSystemIdleTimeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   7140
      Top             =   4560
   End
   Begin MSComctlLib.StatusBar sbrMenu 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7860
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Key             =   "UserKey"
            Object.ToolTipText     =   "Name of current user"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Key             =   "RoleKey"
            Object.ToolTipText     =   "Role of current user"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Key             =   "UserDatabase"
            Object.ToolTipText     =   "Name of current Database"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   10080
      TabIndex        =   0
      Top             =   4680
      Width           =   1200
   End
   Begin VB.Label lblProgress 
      Caption         =   "Progress:"
      Height          =   195
      Left            =   6660
      TabIndex        =   38
      Top             =   5040
      Width           =   675
   End
   Begin VB.Label lblBandNo 
      Caption         =   "Band No."
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   3300
      Width           =   705
   End
   Begin VB.Label lblCriteria 
      Caption         =   "Criteria:"
      Height          =   195
      Left            =   1470
      TabIndex        =   25
      Top             =   4080
      Width           =   705
   End
   Begin VB.Label lblOperator 
      Caption         =   "Operator:"
      Height          =   195
      Left            =   1470
      TabIndex        =   24
      Top             =   3720
      Width           =   705
   End
   Begin VB.Label lblOperand 
      Caption         =   "Operand:"
      Height          =   195
      Left            =   1470
      TabIndex        =   23
      Top             =   3300
      Width           =   700
   End
   Begin VB.Label lblOutput 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " OUTPUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   4700
      Width           =   1250
   End
   Begin VB.Label lblFilter 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FILTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   2900
      Width           =   1250
   End
   Begin VB.Label lblSelect 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SELECT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   13
      Top             =   105
      Width           =   1250
   End
   Begin VB.Label lblNoRecords 
      Caption         =   "No. Subject Records:"
      Height          =   195
      Left            =   3540
      TabIndex        =   10
      Top             =   5055
      Width           =   1550
   End
   Begin VB.Label lblNoResults 
      Caption         =   "No. Results:"
      Height          =   195
      Left            =   4260
      TabIndex        =   8
      Top             =   4620
      Width           =   945
   End
   Begin VB.Label lblStudy 
      Caption         =   "Study:"
      Height          =   195
      Left            =   1680
      TabIndex        =   5
      Top             =   120
      Width           =   500
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFNew 
         Caption         =   "&New Query"
      End
      Begin VB.Menu mnuFOpen 
         Caption         =   "&Open Query"
      End
      Begin VB.Menu mnuSeparator0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFNewBatchedQuery 
         Caption         =   "New &Batched Query"
      End
      Begin VB.Menu mnuFOpenBatchedQuery 
         Caption         =   "Open Batched &Query"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuSeparato3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuVOutputOptions 
         Caption         =   "&Output Options"
      End
      Begin VB.Menu mnuVCollapse 
         Caption         =   "&Collapse all nodes in treeview"
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run"
      Begin VB.Menu mnuRRunQuery 
         Caption         =   "&Run Query"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHUserGuide 
         Caption         =   "&User Guide"
      End
      Begin VB.Menu mnuSeparato4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAboutMacro 
         Caption         =   "&About MACRO"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         frmMenuMacroQuery
' Copyright:    InferMed Ltd. 2000-2005. All Rights Reserved
' Author:       Mo Morris, December 2001
' Purpose:      Contains the main form of the Macro Query Module
'----------------------------------------------------------------------------------------'
'   Revisions:
'   Mo  29/4/2002   changes made to SaveQuery, SR4693
'   Mo  7/5/2002    changes stemming from Label/PersonId being switched
'                   from a single field to 2 separate fields
'                   local recordset grsOutPut changed, field Subject replaced by fields Label and PersonId
'   Mo  27/5/2002   CBB 2.2.13.25, minor change to cmdRunQuery_Click
'   Mo  27/1/2003   Changes throughout Query module as RQGs are incorporated.
'   Mo  1/5/2003    Bug 1356, minor change to LoadOutPut, additional call to DisplayProgressMessage added
'   Mo  15/7/2003   Minor changes (mainly in PrepareOutPut) around the introduction of output
'                   in STATA format
'   Mo  17/11/2004  Bug 2411 - There are now 2 forms of STATA output Standard and Float.
'                   "Standard"  Uses ddmmmyyyy Standard dates (e.g. 01jan2004)
'                   "Float"     Uses ddmmyyyy Float dates (e.g. 01012004 for 1 January 2004)
'                   Changes have been made to cmdSaveOutput_Click, SaveQuery and OpenQuery.
'   Mo  25/1/2005   Bug 2510, Site/PersonId concatenation bug. "-" inserted between Site and
'                   PersonId when concatenating.
'   Mo  25/10/2005  COD0040 - Changes around the new Thesaurus Data Item Type
'   NCJ 19-20 Dec 05 - Check Partial Date Flag where relevant
'   Mo  25/5/2006   Bug 2666 - Filtering of Reals corrected for SQL Server databases
'   Mo  30/5/2006   Bug 2668 - Option to exclude Subject Label from saved output files
'   Mo  2/6/2006    Bug 2737 - Add Question Short Code length to the Options Window
'   Mo  1/8/2006    Bug 2775 - Checking for single digit numeric questions and setting them to two digit fields
'                   that are capable of holding a special value (-1 to -9) when saved in QM output files (STATA, SAS and csv)
'   Mo 18/10/2006   Bug 2822 - Make MACRO Query Module comply with Partial Dates.
'                   Changes made to FormatOutPut, it now checks Partial Date Flag when
'                   deciding if a date should appear in a Date or String field.
'   Mo  1/11/2006   Bug 2795 - "Precede SAS informats with colons" option added.
'   Mo  31/1/2007   Bug 2873 - Real and LabTest response data to be placed in Decimal not Single fields
'                   a change from adSingle to adDecimal
'   Mo  2/4/2007    MRC15022007 - Query Module Batch Facilities
'----------------------------------------------------------------------------------------'

Option Explicit

Private Const msImageUnTicked = "unticked"
Private Const msImageTicked = "ticked"
Private mbComments As Boolean
Private mbCTCGrade As Boolean
Private mbLabResult As Boolean
Private mbStatus As Boolean
Private mbTimeStamp As Boolean
Private mbUserName As Boolean
Private mbValueCode As Boolean

Private mbStudyAllQuestions As Boolean

Private mcolSelectedForms As Collection
Private mcolSelectedVisits As Collection
Private mcolQuestionAttributes As Collection

Private mcolVFQLookUp As Collection

Private msCurrentQueryPathName As String

Private mrsDataItemNames As ADODB.Recordset
Public mrsData As ADODB.Recordset

Private mcolColumnWidths As Collection
Private mbQueryRunning As Boolean

'--------------------------------------------------------------------
Private Sub cboCatCodes_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    EnableAddButton
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboCatCodes_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cboOperand_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'check for a selected item
    If cboOperand.ListIndex = -1 Then
        Exit Sub
    End If
    
    'Changed Mo 5/9/2002, changes around handling of new Site, Subject Label & PersonIds Filter Functions
    Select Case cboOperand.Text
    Case "[SubjectId]"
        Call PopulatecboCatCodesWithPersonIds
        txtCriteria.Text = ""
        txtCriteria.Visible = False
        cboCatCodes.Visible = True
    Case "[Site]"
        Call PopulatecboCatCodesWithSites
        txtCriteria.Text = ""
        txtCriteria.Visible = False
        cboCatCodes.Visible = True
    Case "[SubjectLabel]"
        Call PopulatecboCatCodesWithSubjectLabels
        txtCriteria.Text = ""
        txtCriteria.Visible = False
        cboCatCodes.Visible = True
    Case Else
        'If a question of type category has been selected then populate cboCatCodes
        'with the relevant Codes/Values, make cboCatCodes visable and txtCriteria invisable
        If DataTypeFromId(glSelectedTrialId, cboOperand.ItemData(cboOperand.ListIndex)) = DataType.Category Then
            'Add the codes/values for the selected question into cboCatCodes
            PopulatecboCatCodes (cboOperand.ItemData(cboOperand.ListIndex))
            'Clear txtCriteria, make txtCriteria invisable and cboCatCodes visable
            txtCriteria.Text = ""
            txtCriteria.Visible = False
            cboCatCodes.Visible = True
        Else
            'Make cboCatCodes invisable and txtCriteria visible and clear them both
            txtCriteria.Visible = True
            cboCatCodes.Visible = False
            txtCriteria.Text = ""
            cboCatCodes.Clear
        End If
    End Select
    
    EnableAddButton
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboOperand_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cboOperator_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    EnableAddButton
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboOperator_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cboStudies_click()
'--------------------------------------------------------------------
Dim i As Integer

    On Error GoTo Errhandler

    'exit if no item is currently selected
    If cboStudies.ListIndex < 0 Then Exit Sub
    
    'exit if currently selected item matches glSelectedTrialId
    If cboStudies.ItemData(cboStudies.ListIndex) = glSelectedTrialId Then Exit Sub
    
    If Not gbBatchQueryMode = True Then
        'If there is already a currently selected study check that the user wants to change their query
        If glSelectedTrialId > 0 Then
            If DialogQuestion("Selecting a new Study will clear down your current Query" _
                & vbNewLine & "Do you want to continue?", "Change Query") <> vbYes Then
                    'reset cboStudies to its previous entry
                    For i = 0 To cboStudies.ListCount - 1
                        If cboStudies.ItemData(i) = glSelectedTrialId Then
                            cboStudies.ListIndex = i
                            Exit For
                        End If
                    Next
                    Exit Sub
            End If
            'check that the current query does not need saving
            Call SaveCheck
        End If
    End If
    
    Call ClearQueryAndReset
    
    'Mo 2/4/2007 MRC15022007
    'set gbQuerySaved to false, the new query has yet to be saved with a name
    gbQuerySaved = False
    
    glSelectedTrialId = cboStudies.ItemData(cboStudies.ListIndex)
    
    Call TreeViewThisStudy((cboStudies.ItemData(cboStudies.ListIndex)), cboStudies.Text)
    
    'load the Filter Operands combo
    PopulateOperandCombo
    
    'Add an initial study label to the query text
    txtQueryText.Text = "[S]" & cboStudies.Text
    
    'enable the AND/OR Filter Band Type option buttons and set an initial Band Type
    optAndBands.Enabled = True
    optORBands.Enabled = True
    'The option will remain unchanged from the previous query
    'but the correct text has to be placred in txBandType.Text
    If optAndBands.Value = True Then
        txtBandType.Text = "[AND] Filter Bands connected by OR"
    Else
        txtBandType.Text = "[OR] Filter Bands connected by AND"
    End If

    'enable the cmdNew Filter button
    cmdNew.Enabled = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboStudies_click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cmdAdd_Click()
'--------------------------------------------------------------------
Dim sFilter As String
Dim sCriteria As String

    On Error GoTo Errhandler
    
    'Validate and build the Filter Element Text
    Select Case cboOperand.Text
    Case "[SubjectId]"
        'No need to validate criteria, because it can only be a valid entry from cboCatCodes
        'construct the Filter Element Text
        sFilter = "[" & Trim(txtBandNo.Text) & "]" & Mid(cboOperand.Text, 2, Len(cboOperand.Text) - 2) & Mid(cboOperator.Text, 1, InStr(cboOperator.Text, "(") - 1) & cboCatCodes.Text
    Case "[Site]", "[SubjectLabel]"
        sFilter = "[" & Trim(txtBandNo.Text) & "]" & Mid(cboOperand.Text, 2, Len(cboOperand.Text) - 2) & Mid(cboOperator.Text, 1, InStr(cboOperator.Text, "(") - 1) & "'" & cboCatCodes.Text & "'"
    Case Else
        'validate txtCriteria if it exists (when operator is not either of the NULL functions)
        If ((cboOperator.Text <> " IS NOT NULL (Response Exists)") And (cboOperator.Text <> " IS NULL (No Response)")) Then
            'Validate the filter txtCriteria.text based on the question type
            Select Case DataTypeFromId(glSelectedTrialId, cboOperand.ItemData(cboOperand.ListIndex))
            'Mo 25/10/2005 COD0040
            Case DataType.Text, DataType.Multimedia, DataType.Thesaurus
                If Not gblnValidString(txtCriteria.Text, valOnlySingleQuotes) Then
                    MsgBox "Filter criteria " & gsCANNOT_CONTAIN_INVALID_CHARS, _
                            vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
                    Exit Sub
                ElseIf Not gblnValidString(txtCriteria.Text, valAlpha + valNumeric + valSpace + valDateSeperators) Then
                    MsgBox "Text Filter criteria can only contain alphanumeric characters, spaces or date separators :/.", _
                            vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
                    Exit Sub
                ElseIf Len(txtCriteria.Text) > 50 Then
                    MsgBox "Filter criteria can not be more than 50 characters", _
                            vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
                    Exit Sub
                End If
            Case DataType.Date
                If Not gblnValidString(txtCriteria.Text, valOnlySingleQuotes) Then
                    MsgBox "Filter criteria " & gsCANNOT_CONTAIN_INVALID_CHARS, _
                            vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
                    Exit Sub
                ElseIf Not gblnValidString(txtCriteria.Text, valNumeric + valDateSeperators) Then
                    MsgBox "Date Filter criteria can only contain numeric characters or date separators :/.", _
                            vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
                    Exit Sub
                ElseIf Len(txtCriteria.Text) > 50 Then
                    MsgBox "Filter criteria can not be more than 50 characters", _
                            vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
                    Exit Sub
                End If
            Case DataType.IntegerData, DataType.Real, DataType.LabTest
                If Not gblnValidString(txtCriteria.Text, valOnlySingleQuotes) Then
                    MsgBox "Filter criteria " & gsCANNOT_CONTAIN_INVALID_CHARS, _
                            vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
                    Exit Sub
                ElseIf Not gblnValidString(txtCriteria.Text, valNumeric + valDecimalPoint) Then
                    MsgBox "Numeric Filter criteria can only contain numeric characters and decimal points.", _
                            vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
                    Exit Sub
                ElseIf Len(txtCriteria.Text) > 50 Then
                    MsgBox "Filter criteria can not be more than 50 characters", _
                            vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
                    Exit Sub
                End If
            Case DataType.Category
                'no validation required for criteria of questions of type category,
                'because it can only be a valid entry from cboCatCodes
            End Select
            'Using the question's data type determine the addition of single quotes around the criteria
            Select Case DataTypeFromId(glSelectedTrialId, cboOperand.ItemData(cboOperand.ListIndex))
            'Mo 25/10/2005 COD0040
            Case DataType.Text, DataType.Date, DataType.Multimedia, DataType.Thesaurus
                sCriteria = "'" & txtCriteria.Text & "'"
            Case DataType.IntegerData, DataType.Real, DataType.LabTest
                sCriteria = txtCriteria.Text
            Case DataType.Category
                sCriteria = "'" & Mid(cboCatCodes.Text, 1, InStr(cboCatCodes.Text, " ") - 1) & "'"
            End Select
        Else
            sCriteria = ""
        End If
        'construct the Filter Element Text
        sFilter = "[" & Trim(txtBandNo.Text) & "]" & cboOperand.Text & Mid(cboOperator.Text, 1, InStr(cboOperator.Text, "(") - 1) & sCriteria
    End Select
    
    If cmdAdd.Caption = "Change" Then
        'Remove Filter from  listbox lstFilterText
        lstFilterText.RemoveItem lstFilterText.ListIndex
    End If
    
    'add new Filter to listbox
    lstFilterText.AddItem sFilter
    'store the selected questions DataItemID as ItemData
    lstFilterText.ItemData(lstFilterText.NewIndex) = cboOperand.ItemData(cboOperand.ListIndex)
    
    'Disable the Save Output button
    cmdSaveOutPut.Enabled = False
    
    'flag query as changed
    gbQueryChanged = True
    
    'clear down the elements that made up the new Filter
    cboOperand.ListIndex = -1
    cboOperand.Enabled = False
    cboOperator.ListIndex = -1
    cboOperator.Enabled = False
    txtCriteria.Text = ""
    txtCriteria.Enabled = False
    txtCriteria.Visible = True
    cboCatCodes.Clear
    cboCatCodes.Enabled = False
    cboCatCodes.Visible = False
    txtBandNo.Text = ""
    txtBandNo.Enabled = False
    
    If cmdAdd.Caption = "Change" Then
        'The edit is over. Put cmdAdd's caption back to 'add'
        'and disable the edit command button
        cmdAdd.Caption = "Add"
        cmdEdit.Enabled = False
    End If
    
    'enable/disable the relevant controls
    cmdAdd.Enabled = False
    cmdNew.Enabled = True
    cmdCancel.Enabled = False
    lstFilterText.Enabled = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdAdd_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'make sur that cmdAdd's caption is 'Add"
    cmdAdd.Caption = "Add"
    
    'enable/disable the relevant controls
    cboOperand.ListIndex = -1
    cboOperand.Enabled = False
    cboOperator.ListIndex = -1
    cboOperator.Enabled = False
    txtCriteria.Text = ""
    txtCriteria.Enabled = False
    txtCriteria.Visible = True
    cboCatCodes.Clear
    cboCatCodes.Enabled = False
    cboCatCodes.Visible = False
    txtBandNo.Text = ""
    txtBandNo.Enabled = False
    cmdCancel.Enabled = False
    cmdAdd.Enabled = False
    lstFilterText.ListIndex = -1
    lstFilterText.Enabled = True
    cmdEdit.Enabled = False
    cmdNew.Enabled = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdCancel_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cmdCancelRun_Click()
'--------------------------------------------------------------------

    gbCancelled = True

End Sub

'--------------------------------------------------------------------
Private Sub cmdDelete_Click()
'--------------------------------------------------------------------
    
    On Error GoTo Errhandler
    
    'remove from the Filter listbox
    lstFilterText.RemoveItem lstFilterText.ListIndex
    
    'Disable the Save Output button
    cmdSaveOutPut.Enabled = False
    
    'flag query as changed
    gbQueryChanged = True
    
    'enable/disable the relevant controls
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdDelete_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cmdEdit_Click()
'--------------------------------------------------------------------
Dim sFilterText As String
Dim sOperand As String
Dim sOperator As String
Dim sCriteria As String
Dim sBandNo As String
Dim i As Integer

    On Error GoTo Errhandler

    'change cmdAdd caption to change
    cmdAdd.Caption = "Change"
    
    'Extract the Filter elements from the filter string
    'which is of the form [BAND NUMBER]OPERAND OPERATOR CRITERIA, with delimiter spaces
    'or of the form [BAND NUMBER]OPERAND IS NULL/IS NOT NULL
    'strip of leading "["
    sFilterText = Mid(lstFilterText.Text, 2)
    sBandNo = Mid(sFilterText, 1, InStr(sFilterText, "]") - 1)
    sFilterText = Mid(sFilterText, InStr(sFilterText, "]") + 1)
    sOperand = Mid(sFilterText, 1, InStr(sFilterText, " ") - 1)
    sFilterText = Mid(sFilterText, InStr(sFilterText, " ") + 1)
    If sFilterText = "IS NULL " Or sFilterText = "IS NOT NULL " Then
        sOperator = " " & sFilterText
        sCriteria = ""
    Else
        sOperator = " " & Mid(sFilterText, 1, InStr(sFilterText, " "))
        sCriteria = Mid(sFilterText, InStr(sFilterText, " ") + 1)
    End If
    
    'Check for a Special Filter Function (negative itemdata of -1, -2, -3 etc) and add square brackets
    If lstFilterText.ItemData(lstFilterText.ListIndex) < 0 Then
        sOperand = "[" & sOperand & "]"
    End If
    
    'Strip off any single quotes from around sCriteria
    If sCriteria <> "" Then
        If Mid(sCriteria, 1, 1) = "'" Then
            sCriteria = Mid(sCriteria, 2)
        End If
        If Mid(sCriteria, Len(sCriteria)) = "'" Then
            sCriteria = Mid(sCriteria, 1, Len(sCriteria) - 1)
        End If
    End If

    'Place Filter elements in the appropriate combos/textboxes
    cboOperand.Enabled = True
    For i = 0 To cboOperand.ListCount - 1
        If cboOperand.List(i) = sOperand Then
            cboOperand.ListIndex = i
            Exit For
        End If
    Next
    
    cboOperator.Enabled = True
    For i = 0 To cboOperator.ListCount - 1
        If Mid(cboOperator.List(i), 1, InStr(cboOperator.List(i), "(") - 1) = sOperator Then
            cboOperator.ListIndex = i
            Exit For
        End If
    Next
    
    'Populate cboCatCodes/txtCriteria correctly, checking for Special Filter Functions first
    Select Case lstFilterText.ItemData(lstFilterText.ListIndex)
    Case -1
        Call PopulatecboCatCodesWithPersonIds
        'select the required element in cboCatCodes
        For i = 0 To cboCatCodes.ListCount - 1
            If cboCatCodes.List(i) = sCriteria Then
                cboCatCodes.ListIndex = i
                Exit For
            End If
        Next
        cboCatCodes.Visible = True
        txtCriteria.Text = ""
        txtCriteria.Visible = False
    Case -2
        Call PopulatecboCatCodesWithSites
        'select the required element in cboCatCodes
        For i = 0 To cboCatCodes.ListCount - 1
            If cboCatCodes.List(i) = sCriteria Then
                cboCatCodes.ListIndex = i
                Exit For
            End If
        Next
        cboCatCodes.Visible = True
        txtCriteria.Text = ""
        txtCriteria.Visible = False
    Case -3
        Call PopulatecboCatCodesWithSubjectLabels
        'select the required element in cboCatCodes
        For i = 0 To cboCatCodes.ListCount - 1
            If cboCatCodes.List(i) = sCriteria Then
                cboCatCodes.ListIndex = i
                Exit For
            End If
        Next
        cboCatCodes.Visible = True
        txtCriteria.Text = ""
        txtCriteria.Visible = False
    Case Else
        If sCriteria <> "" Then
            'Extract DataItemId from lstFilterText.ItemDat and check DataItemType for category questions
            If DataTypeFromId(glSelectedTrialId, lstFilterText.ItemData(lstFilterText.ListIndex)) = DataType.Category Then
                PopulatecboCatCodes (lstFilterText.ItemData(lstFilterText.ListIndex))
                'select the required element in cboCatCodes
                For i = 0 To cboCatCodes.ListCount - 1
                    If Mid(cboCatCodes.List(i), 1, InStr(cboCatCodes.List(i), " ") - 1) = sCriteria Then
                        cboCatCodes.ListIndex = i
                        Exit For
                    End If
                Next
                cboCatCodes.Visible = True
                txtCriteria.Text = ""
                txtCriteria.Visible = False
            Else
                txtCriteria.Text = sCriteria
                txtCriteria.Visible = True
                cboCatCodes.Visible = False
            End If
        End If
    End Select
    
    txtBandNo.Enabled = True
    txtBandNo.Text = sBandNo
        
    'enable/disable the relevant controls
    txtCriteria.Enabled = True
    cboCatCodes.Enabled = True
    cmdDelete.Enabled = False
    cmdNew.Enabled = False
    cmdEdit.Enabled = False
    lstFilterText.Enabled = False
    cmdCancel.Enabled = True
    cmdAdd.Enabled = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdEdit_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cmdExit_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'check that the current query does not need saving
    Call SaveCheck

    Call ExitMACRO
    Call MACROEnd

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdExit_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Public Sub InitialiseMe()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'load the forms study combo with the studies in the current database
    LoadStudyCombo
    
    'set an initial Band Type and clear txBandType.Text (because query text areas should be clear at this stage)
    optAndBands.Value = True
    txtBandType.Text = ""
    
    Call ClearQueryAndReset

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "InitialiseMe")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Public Sub CheckUserRights()
'--------------------------------------------------------------------

End Sub

'--------------------------------------------------------------------
Private Sub cmdNew_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'enable/disable the relevant controls
    lstFilterText.ListIndex = -1
    lstFilterText.Enabled = False
    cmdNew.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cboOperand.Enabled = True
    cboOperator.Enabled = True
    txtCriteria.Enabled = True
    cboCatCodes.Enabled = True
    txtBandNo.Enabled = True
    cmdCancel.Enabled = True
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdNew_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cmdRunQuery_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    Call DisableRunSaveDisplay

    Call HourglassOn
    
    'Clear current grid and remove any splits
    Call ClearGrid
    
    'Clear the Retrieved Result and Records text boxes
    txtResultsCount.Text = ""
    txtRecordsCount.Text = ""
    DoEvents
    
    'Query the database for the specified response data
    Call QueryDB
    
    'Check for Call to QueryDB having being cancelled
    If gbCancelled Then
        Call HourglassOff
        Exit Sub
    End If
    
    If gbDisplayOutPut Then
        gbNotDisplayedNotSaved = False
        If mrsData.RecordCount > 0 Then
            Call DisplayOutPut
        End If
    Else
        gbNotDisplayedNotSaved = True
    End If
    
    'Check for Call to DisplayOutPut having being cancelled
    If gbCancelled Then
        Call HourglassOff
        Exit Sub
    End If
    
    Call EnableRunSaveDisplay
    
    Call HourglassOff
  
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdRunQuery_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cmdSaveOutput_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    Call HourglassOn

    Call DisableRunSaveDisplay
    
    'Check for the current query not having been displayed
    'If the current query has not been displayed the recordset grsOutPut will not have been created
    If Not gbDisplayOutPut Then
        'If the not displayed query has already been saved the boolean gbNotDisplayedNotSaved will have been set
        If gbNotDisplayedNotSaved Then
            If gbUseShortCodes Then
                Set gColQuestionCodes = New Collection
            End If
            Call PrepareOutPut
            Call LoadOutPut
            gbNotDisplayedNotSaved = False
        End If
    End If
    
    'Check for actions having been cancelled
    If gbCancelled Then
        Call HourglassOff
        Exit Sub
    End If
    
    'Mo 1/11/2006 Bug 2795
    Select Case gnOutPutType
    Case eOutPutType.CSV
        Call OutputToCSV
    Case eOutPutType.Access
        Call OutputToAccess
    Case eOutPutType.SPSS
        Call DialogInformation("SPSS Output facilities are not yet available in this module.", "MACRO Query Module")
    Case eOutPutType.SAS, eOutPutType.SASColons
        Call OutputToSAS
    Case eOutPutType.STATA
        Call OutputToSTATA("Float")
    Case eOutPutType.MACROBD
        Call OutputToMACROBD
    Case eOutPutType.STATAStandardDates
        Call OutputToSTATA("Standard")
    End Select

    Call EnableRunSaveDisplay
    
    Call HourglassOff
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdSaveOutput_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub FilterOutputBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'If Left Button is down update Top Property of FilterOutputBar and call refresh
    If (Button = vbLeftButton) Then
        FilterOutputBar.Top = FilterOutputBar.Top + Y
        'Me.Refresh
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "FilterOutputBar_MouseMove")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub FilterOutputBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'Check that FilterOutputBar has not been dragged to far to the Top
    If (FilterOutputBar.Top < SelectFilterBar.Top + 1860) Then
        FilterOutputBar.Top = SelectFilterBar.Top + 1860
    End If
    
    'Check that FilterOutputBar has not been dragged to far to the Bottom
    If (FilterOutputBar.Top > Me.ScaleHeight - 2200) Then
        FilterOutputBar.Top = Me.ScaleHeight - 2200
    End If
    
    lstFilterText.Height = FilterOutputBar.Top - lstFilterText.Top - 100
    lblOutput.Top = FilterOutputBar.Top + 200
    fraDisplayOptions.Top = lblOutput.Top - 90
    lblNoResults.Top = lblOutput.Top
    txtResultsCount.Top = lblOutput.Top
    txtRecordsCount.Top = txtResultsCount.Top + txtResultsCount.Height + 100
    lblNoRecords.Top = txtRecordsCount.Top
    cmdExit.Top = lblOutput.Top
    cmdCancelRun.Top = lblOutput.Top
    cmdRunQuery.Top = lblOutput.Top
    cmdSaveOutPut.Top = lblOutput.Top
    txtProgress.Top = cmdExit.Top + cmdExit.Height + 100
    lblProgress.Top = txtProgress.Top
    grdOutPut.Top = lblOutput.Top + lblOutput.Height + 500
    grdOutPut.Height = Me.ScaleHeight - grdOutPut.Top - 500

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "FilterOutputBar_MouseUp")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    GetRegistrySettings
    
    If gbDisplayOutPut Then
        optDisplayOutput.Value = True
    Else
        optDoNotDisplayOutput.Value = True
    End If
    
    SelectFilterBar.Top = gnSelectFilterBarTop
    FilterOutputBar.Top = gnFilterOutputBarTop
    Me.Left = gnRegFormLeft
    Me.Top = gnRegFormTop
    Me.Width = gnRegFormWidth
    Me.Height = gnRegFormHeight
    'The above lines will trigger a call to Form_Resize
    
    'set cbocatcodes to invisible
    cboCatCodes.Visible = False
    
    'Load the cboOpererators combo
    cboOperator.AddItem " = (Equal To)"
    cboOperator.AddItem " > (Greater Than)"
    cboOperator.AddItem " < (Less Than)"
    cboOperator.AddItem " >= (Greater Than or Equal To)"
    cboOperator.AddItem " <= (Less Than or Equal To)"
    cboOperator.AddItem " <> (Not Equal To)"
    cboOperator.AddItem " IS NOT NULL (Response Exists)"
    cboOperator.AddItem " IS NULL (No Response)"
    'cboOperator.AddItem " IN (Set Operator)"
    
    Set trwQuestions.ImageList = imgTreeViewIcons16
    
Exit Sub
Errhandler:
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

'--------------------------------------------------------------------
Private Sub Form_Resize()
'--------------------------------------------------------------------
Dim nFormWidth As Integer

    On Error GoTo Errhandler

    'check that form has not been minimised
    If Me.WindowState = vbMinimized Then Exit Sub
     
    If Me.Width < mnMINFORMWIDTH Then Me.Width = mnMINFORMWIDTH

    If Me.Height < mnMINFORMHEIGHT Then Me.Height = mnMINFORMHEIGHT
    
    If Me.Height < FilterOutputBar.Top + 2000 Then
        FilterOutputBar.Top = 4495
        SelectFilterBar.Top = 2635
    End If
    
    nFormWidth = Me.ScaleWidth
    
    lblSelect.Left = 100
    lblSelect.Top = 100
    
    lblStudy.Left = 1700
    lblStudy.Top = 100
    
    cboStudies.Left = 2200
    cboStudies.Top = 100
    cboStudies.Width = ((nFormWidth - 300) / 2) - 2100
    
    trwQuestions.Left = 100
    trwQuestions.Top = 500
    trwQuestions.Width = (nFormWidth - 300) / 2
    trwQuestions.Height = SelectFilterBar.Top - trwQuestions.Top - 100
    
    txtQueryText.Left = ((nFormWidth - 300) / 2) + 200
    txtQueryText.Top = 100
    txtQueryText.Width = (nFormWidth - 300) / 2
    txtQueryText.Height = SelectFilterBar.Top - txtQueryText.Top - 100
    
    SelectFilterBar.Left = 100
    'Note that SelectFilterBar.Top is set in Form_Load
    SelectFilterBar.Width = nFormWidth - 200
    
    lblFilter.Left = 100
    lblFilter.Top = SelectFilterBar.Top + 200
    
    lblBandNo.Left = 100
    lblBandNo.Top = lblFilter.Top + lblFilter.Height + 100
    
    txtBandNo.Left = 950
    txtBandNo.Top = lblBandNo.Top
    
    lblOperand.Left = 1450
    lblOperand.Top = lblBandNo.Top
    
    cboOperand.Left = 2250
    cboOperand.Top = lblOperand.Top
    cboOperand.Width = trwQuestions.Width - 2150

    cmdCancel.Left = cboOperand.Width + 1450
    cmdCancel.Top = lblFilter.Top
    
    cmdEdit.Left = cmdCancel.Left - 900
    cmdEdit.Top = lblFilter.Top
    
    cmdDelete.Left = cmdEdit.Left - 900
    cmdDelete.Top = lblFilter.Top
    
    cmdAdd.Left = cmdDelete.Left - 900
    cmdAdd.Top = lblFilter.Top
    
    cmdNew.Left = cmdAdd.Left - 900
    cmdNew.Top = lblFilter.Top
    
    cboOperator.Left = 2250
    cboOperator.Top = cboOperand.Top + cboOperand.Height + 100
    cboOperator.Width = cboOperand.Width
    
    lblOperator.Left = 1450
    lblOperator.Top = cboOperator.Top
    
    txtCriteria.Left = 2250
    txtCriteria.Top = cboOperator.Top + cboOperator.Height + 100
    txtCriteria.Width = cboOperand.Width
    cboCatCodes.Left = 2250
    cboCatCodes.Top = txtCriteria.Top
    cboCatCodes.Width = txtCriteria.Width
    
    lblCriteria.Left = 1450
    lblCriteria.Top = txtCriteria.Top
    
    fraBandType.Left = 100
    fraBandType.Top = lblOperator.Top
    
    txtBandType.Left = txtQueryText.Left
    txtBandType.Top = lblFilter.Top
    txtBandType.Width = txtQueryText.Width
    
    lstFilterText.Left = txtQueryText.Left
    lstFilterText.Top = cboOperand.Top
    lstFilterText.Width = txtQueryText.Width
    lstFilterText.Height = FilterOutputBar.Top - lstFilterText.Top - 100 'HERE
    
    FilterOutputBar.Left = 100
    'Note that FilterOutputBar.Top is set in Form_Load
    FilterOutputBar.Width = nFormWidth - 200
    
    lblOutput.Left = 100
    lblOutput.Top = FilterOutputBar.Top + 200
    
    fraDisplayOptions.Left = 1450
    fraDisplayOptions.Top = lblOutput.Top - 90
    
    lblNoResults.Left = 3800
    lblNoResults.Top = lblOutput.Top
    
    txtResultsCount.Left = 5400
    txtResultsCount.Top = lblOutput.Top
    
    txtRecordsCount.Left = 5400
    txtRecordsCount.Top = txtResultsCount.Top + txtResultsCount.Height + 100
    
    lblNoRecords.Left = 3800
    lblNoRecords.Top = txtRecordsCount.Top
    
    cmdExit.Left = nFormWidth - 1300
    cmdExit.Top = lblOutput.Top
    
    cmdCancelRun.Left = cmdExit.Left - 1300
    cmdCancelRun.Top = lblOutput.Top
    
    cmdRunQuery.Left = cmdCancelRun.Left - 1300
    cmdRunQuery.Top = lblOutput.Top
    
    cmdSaveOutPut.Left = cmdRunQuery.Left - 1300
    cmdSaveOutPut.Top = lblOutput.Top
    
    txtProgress.Left = cmdRunQuery.Left
    txtProgress.Width = (cmdExit.Left - cmdRunQuery.Left) + cmdRunQuery.Width
    txtProgress.Top = cmdExit.Top + cmdExit.Height + 100
    
    lblProgress.Left = txtProgress.Left - lblProgress.Width - 50
    lblProgress.Top = txtProgress.Top

    grdOutPut.Left = 100
    grdOutPut.Top = lblOutput.Top + lblOutput.Height + 500
    grdOutPut.Width = nFormWidth - 200
    grdOutPut.Height = Me.ScaleHeight - grdOutPut.Top - 500
    'if a split exists then reset the width to the right of the split
    If grdOutPut.Splits.Count > 1 Then
        grdOutPut.Splits(1).Size = grdOutPut.Width - grdOutPut.Splits(0).Size - 70
    End If
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Form_Resize")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'--------------------------------------------------------------------

    'Check for a Query being run
    If mbQueryRunning Then
        Cancel = 1
        Exit Sub
    End If
    
    'check that the current query does not need saving
    Call SaveCheck
    
    Call ExitMACRO
    Call MACROEnd

End Sub

'--------------------------------------------------------------------
Private Sub grdOutPut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--------------------------------------------------------------------
'The code in this sub is activated when 8 character Question codes are being displayed.
'The grids Tool Tip text is changed to the Long format (VisitCode/eFormCode/QuestionCode)
'of the currently hovered over question/column.
'Note that this facility is not required for columns 0 to 5 (the identification fields)
'--------------------------------------------------------------------

    'On Error Resume Next is set because the X co-ordinate is not always contained by a column
    On Error Resume Next
    
    'Exit if 8 character Question Codes are not being used
    If Not gbUseShortCodes Then
        grdOutPut.ToolTipText = ""
        Exit Sub
    End If

    'Exit if current column is 0 to 5
    If grdOutPut.Columns(grdOutPut.ColContaining(X)).ColIndex > 5 Then
        grdOutPut.ToolTipText = gColQuestionCodes(grdOutPut.Columns(grdOutPut.ColContaining(X)).Caption)
    Else
        grdOutPut.ToolTipText = ""
    End If
    
    'Clear any error that might have occured
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub

'--------------------------------------------------------------------
Private Sub lstFilterText_Click()
'--------------------------------------------------------------------
    
    On Error GoTo Errhandler
    
    'Don't enable Delete or Edit if the AND/OR Band Filter Type entry (index = 0) is clicked
    If lstFilterText.ListIndex <> -1 Then
        'disable the add Filter button
        cmdAdd.Enabled = False
        'make sure its caption is 'Add'
        cmdAdd.Caption = "Add"
        'enable the delete and edit candidate buttons
        cmdDelete.Enabled = True
        cmdEdit.Enabled = True
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lstFilterText_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub mnuFExit_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'check that the current query does not need saving
    Call SaveCheck

    Call ExitMACRO
    Call MACROEnd

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuFExit_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub mnuFNew_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'check that the current query does not need saving
    Call SaveCheck
    
    Call ClearQueryAndReset
    
    'Unselect the previously selected study
    cboStudies.ListIndex = -1
    
    'set gbQuerySaved to false, the new query has yet to be saved with a name
    gbQuerySaved = False
    
    'set focus to cboStudies
    cboStudies.SetFocus

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuFNew_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'Mo 2/4/2007 MRC15022007
'--------------------------------------------------------------------
Private Sub mnuFNewBatchedQuery_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'check that the current query does not need saving
    Call SaveCheck
    
    Call ClearQueryAndReset
    
    'Unselect the previously selected study
    cboStudies.ListIndex = -1
    
    'set gbBatchQuerySaved to false, the new batch query has yet to be saved with a name
    gbBatchQuerySaved = False
    
    gbBatchQueryMode = True
    
    'Open the Batched Query Window
    frmBatchedQuery.Show vbModal

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuFNewBatchedQuery_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub mnuFOpen_Click()
'--------------------------------------------------------------------
Dim sQueryPathName As String

    'check that the current query does not need saving
    Call SaveCheck
    
    'clear everything down prior to opening query
    Call ClearQueryAndReset
    
    'Unselect the previously selected study
    cboStudies.ListIndex = -1

    On Error GoTo CancelOpen
    With CommonDialog1
        .DialogTitle = "Open MACRO Query"
        .InitDir = gsOUT_FOLDER_LOCATION
        .DefaultExt = "txt"
        .Filter = "Text file (*.txt)|*.txt"
        .CancelError = True
        .ShowOpen
  
        sQueryPathName = .FileName
    End With
    
    'open the selected query
    Call OpenQuery(sQueryPathName)
    
    'Store the name of the opened query
    msCurrentQueryPathName = sQueryPathName
    
    'set Query has been saved with a name flag
    gbQuerySaved = True

    
CancelOpen:

End Sub

'Mo 2/4/2007 MRC15022007
'--------------------------------------------------------------------
Private Sub mnuFOpenBatchedQuery_Click()
'--------------------------------------------------------------------

    'check that the current query does not need saving
    Call SaveCheck
    
    'clear everything down prior to opening query
    Call ClearQueryAndReset
    
    'Unselect the previously selected study
    cboStudies.ListIndex = -1
    
    gbBatchQueryMode = True
    
    'Open the Batched Query Window
    frmBatchedQuery.OpenBatchQuery
    
End Sub

'--------------------------------------------------------------------
Private Sub mnuFPrint_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    Call DialogInformation("Print facilities are not yet available in this module.", "MACRO Query Module")

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuFPrint_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub mnuFSave_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'if the query has not been named then a call to Save As is required
    If gbQuerySaved = False Then
        mnuFSaveAs_Click
        Exit Sub
    End If

    Call SaveQuery(msCurrentQueryPathName)

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuFSave_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub mnuFSaveAs_Click()
'--------------------------------------------------------------------
Dim sQueryName As String
Dim sQueryPathName As String

    'Check for the current query already having a name
    If gbQuerySaved = False Then
        sQueryName = "MQ " & cboStudies.Text & " (" & Format(Now, "yyyy mm dd hh mm") & ").txt"
        sQueryPathName = gsOUT_FOLDER_LOCATION & sQueryName
    Else
        sQueryPathName = msCurrentQueryPathName
    End If
    
    On Error GoTo CancelSaveAs
    With CommonDialog1
        .DialogTitle = "Save MACRO Query As"
        .CancelError = True
        .Filter = "Text file (*.txt)|*.txt"
        .DefaultExt = "txt"
        .Flags = cdlOFNOverwritePrompt
        .FileName = sQueryPathName
        .ShowSave
  
        sQueryPathName = .FileName
    End With
    
    'save the query
    Call SaveQuery(sQueryPathName)
    
    'Store the name of the saved query
    msCurrentQueryPathName = sQueryPathName
    
    'set Query has been saved with a name flag
    gbQuerySaved = True
    
    'set changed flag false
    gbQueryChanged = False

CancelSaveAs:

End Sub

'--------------------------------------------------------------------
Private Sub mnuHAboutMacro_Click()
'--------------------------------------------------------------------

    frmAbout.Display

End Sub

'--------------------------------------------------------------------
Private Sub mnuHUserGuide_Click()
'--------------------------------------------------------------------

    'Mo 3/7/2002, CBB 2.2.18.6
    Call MACROHelp(Me.hWnd, App.Title)
    
End Sub

'--------------------------------------------------------------------
Private Sub LoadStudyCombo()
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsStudies As ADODB.Recordset

    On Error GoTo Errhandler

    sSQL = "SELECT ClinicalTrialId, ClinicalTrialName " _
        & "FROM ClinicalTrial " _
        & "WHERE ClinicalTrialID > 0 " _
        & "ORDER BY ClinicalTrialName"
    Set rsStudies = New ADODB.Recordset
    rsStudies.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do Until rsStudies.EOF
        cboStudies.AddItem rsStudies!ClinicalTrialName
        cboStudies.ItemData(cboStudies.NewIndex) = rsStudies!ClinicalTrialId
        rsStudies.MoveNext
    Loop
    
    rsStudies.Close
    Set rsStudies = Nothing

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadStudyCombo")
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
Private Sub mnuRRunQuery_Click()
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    cmdRunQuery_Click

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuRRunQuery_Click")
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
Private Sub mnuVCollapse_Click()
'---------------------------------------------------------------------
Dim oNode As MSComctlLib.Node

    On Error GoTo Errhandler
    
    For Each oNode In trwQuestions.Nodes
        oNode.Expanded = False
    Next
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuVCollapse_Click")
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
Private Sub TreeViewThisStudy(ByVal lTrialId As Long, _
                                ByVal sTrialName As String)
'---------------------------------------------------------------------
Dim nodX As MSComctlLib.Node
Dim sRootKey As String
Dim sVisitsRootKey As String
Dim sFormsRootKey As String
Dim sQuestionsRootKey As String

    On Error GoTo Errhandler

    HourglassOn
    trwQuestions.Visible = False
    
    'Clear the selected nodes collections
    Set mcolSelectedForms = Nothing
    Set mcolSelectedForms = New Collection
    Set mcolSelectedVisits = Nothing
    Set mcolSelectedVisits = New Collection
    Set mcolQuestionAttributes = Nothing
    Set mcolQuestionAttributes = New Collection
    'initialise the Study.All Questions ticked boolean
    mbStudyAllQuestions = False
    
    'Add Selected study name as root node
    sRootKey = "S"
    Set nodX = trwQuestions.Nodes.Add(, , sRootKey, sTrialName, msImageUnTicked)
    nodX.Tag = msImageUnTicked
    trwQuestions.Nodes(sRootKey).Expanded = True
    
    'Add Visits node
    sVisitsRootKey = sRootKey & "|V"
    Set nodX = trwQuestions.Nodes.Add(sRootKey, tvwChild, sVisitsRootKey, "Visits", msImageUnTicked)
    nodX.Tag = msImageUnTicked
    'add all visits and question under each visit
    Call TreeViewVisits(lTrialId, sVisitsRootKey)
    
    'Add Forms node
    sFormsRootKey = sRootKey & "|F"
    Set nodX = trwQuestions.Nodes.Add(sRootKey, tvwChild, sFormsRootKey, "Forms", msImageUnTicked)
    nodX.Tag = msImageUnTicked
    'add all forms and question under each form
    Call TreeViewForms(lTrialId, sFormsRootKey)
    
    'Add Questions node
    sQuestionsRootKey = sRootKey & "|Q"
    Set nodX = trwQuestions.Nodes.Add(sRootKey, tvwChild, sQuestionsRootKey, "Questions", msImageUnTicked)
    nodX.Tag = msImageUnTicked
    'add all questions
    Call TreeViewQuestions(lTrialId, sQuestionsRootKey)
    
    trwQuestions.Visible = True
    HourglassOff

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "TreeViewThisStudy")
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
Private Sub TreeViewVisits(ByVal lTrialId As Long, _
                            ByVal sVisitsRootKey As String)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsVisits As ADODB.Recordset
Dim rsQuestions As ADODB.Recordset
Dim nodX As MSComctlLib.Node
Dim sVisitKey As String
Dim sQuestionKey As String

    On Error GoTo Errhandler

    'create recordset of the studies visits
    sSQL = "SELECT VisitId, VisitCode " _
        & "FROM StudyVisit " _
        & "WHERE ClinicalTrialId = " & lTrialId & " " _
        & "ORDER BY VisitOrder"
    Set rsVisits = New ADODB.Recordset
    rsVisits.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    'Loop through the studies visits
    Do Until rsVisits.EOF
        'create a unique key for this visit
        sVisitKey = sVisitsRootKey & "|" & rsVisits!VisitId
        'Add visit to the treeview
        Set nodX = trwQuestions.Nodes.Add(sVisitsRootKey, tvwChild, sVisitKey, rsVisits!VisitCode, msImageUnTicked)
        nodX.Tag = msImageUnTicked
        'Create an initial entry in the selected form nodes collection
        mcolSelectedVisits.Add 0, Str(rsVisits!VisitId)
        'Create a recordset of the questions within this visit
        sSQL = "SELECT DISTINCT DataItem.DataItemId, DataItem.DataItemCode " _
            & "FROM DataItem, CRFElement, StudyVisitCRFPage " _
            & "WHERE CRFElement.ClinicalTrialId = DataItem.ClinicalTrialId " _
            & "AND CRFElement.DataItemId = DataItem.DataItemId " _
            & "AND CRFElement.ClinicalTrialId = StudyVisitCRFPage.ClinicalTrialId " _
            & "AND CRFElement.CRFPageId = StudyVisitCRFPage.CRFPageId " _
            & "AND StudyVisitCRFPage.ClinicalTrialId = " & lTrialId & " " _
            & "AND StudyVisitCRFPage.VisitId = " & rsVisits!VisitId & " " _
            & "ORDER BY DataItem.DataItemCode"
        Set rsQuestions = New ADODB.Recordset
        rsQuestions.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        'Loop through the questions for this visit
        Do Until rsQuestions.EOF
            'create a unique key for this Question
            sQuestionKey = sVisitKey & "|" & rsQuestions!DataItemId
            'Add question to the treeview
            Set nodX = trwQuestions.Nodes.Add(sVisitKey, tvwChild, sQuestionKey, rsQuestions!DataItemCode, msImageUnTicked)
            nodX.Tag = msImageUnTicked
            rsQuestions.MoveNext
        Loop
        rsQuestions.Close
        Set rsQuestions = Nothing
        rsVisits.MoveNext
    Loop
        
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "TreeViewVisits")
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
Private Sub TreeViewForms(ByVal lTrialId As Long, _
                            ByVal sFormsRootKey As String)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsForms As ADODB.Recordset
Dim rsQuestions As ADODB.Recordset
Dim nodX As MSComctlLib.Node
Dim sFormKey As String
Dim sQuestionKey As String

    On Error GoTo Errhandler

    'create recordset of the studies Forms
    sSQL = "SELECT CRFPageId, CRFPageCode " _
        & "FROM CRFPage " _
        & "WHERE ClinicalTrialId = " & lTrialId & " " _
        & "ORDER BY CRFPageOrder"
    Set rsForms = New ADODB.Recordset
    rsForms.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    'Loop through the studies Forms
    Do Until rsForms.EOF
        'create a unique key for this Form
        sFormKey = sFormsRootKey & "|" & rsForms!CRFPageId
        'Add Form to the treeview
        Set nodX = trwQuestions.Nodes.Add(sFormsRootKey, tvwChild, sFormKey, rsForms!CRFPageCode, msImageUnTicked)
        nodX.Tag = msImageUnTicked
        'Create an initial entry in the selected form nodes collection
        mcolSelectedForms.Add 0, Str(rsForms!CRFPageId)
        'Create a recordset of the questions on this form
        sSQL = "SELECT DataItem.DataItemId, DataItem.DataItemCode " _
            & "FROM DataItem, CRFElement " _
            & "WHERE CRFElement.ClinicalTrialId = DataItem.ClinicalTrialId " _
            & "AND CRFElement.DataItemId = DataItem.DataItemId " _
            & "AND CRFElement.ClinicalTrialId = " & lTrialId & " " _
            & "AND CRFElement.CRFPageId = " & rsForms!CRFPageId & " " _
            & "ORDER BY CRFElement.FieldOrder, CRFElement.QGroupFieldOrder"
        Set rsQuestions = New ADODB.Recordset
        rsQuestions.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        'Loop through the questions for this form
        Do Until rsQuestions.EOF
            'create a unique key for this Question
            sQuestionKey = sFormKey & "|" & rsQuestions!DataItemId
            'Add question to the treeview
            Set nodX = trwQuestions.Nodes.Add(sFormKey, tvwChild, sQuestionKey, rsQuestions!DataItemCode, msImageUnTicked)
            nodX.Tag = msImageUnTicked
            rsQuestions.MoveNext
        Loop
        rsQuestions.Close
        Set rsQuestions = Nothing
        rsForms.MoveNext
    Loop
    rsForms.Close
    Set rsForms = Nothing
        
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "TreeViewForms")
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
Private Sub TreeViewQuestions(ByVal lTrialId As Long, _
                            ByVal sQuestionsRootKey As String)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsQuestions As ADODB.Recordset
Dim nodX As MSComctlLib.Node
Dim sQuestionKey As String

    On Error GoTo Errhandler

    'create recordset of the studies Questions
    sSQL = "SELECT DataItemId, DataItemCode " _
        & "FROM DataItem " _
        & "WHERE ClinicalTrialId = " & lTrialId & " " _
        & "ORDER BY DataItemCode"
    Set rsQuestions = New ADODB.Recordset
    rsQuestions.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    'Loop through the studies Questions
    Do Until rsQuestions.EOF
        'create a unique key for this Question
        sQuestionKey = sQuestionsRootKey & "|" & rsQuestions!DataItemId
        'Add Question to the treeview
        Set nodX = trwQuestions.Nodes.Add(sQuestionsRootKey, tvwChild, sQuestionKey, rsQuestions!DataItemCode, msImageUnTicked)
        nodX.Tag = msImageUnTicked
        'Create an initial entry in the selected Question Attributes collection
        mcolQuestionAttributes.Add 0, Str(rsQuestions!DataItemId)
        'Add the Question Attributes to the treeview
        Set nodX = trwQuestions.Nodes.Add(sQuestionKey, tvwChild, sQuestionKey & "|Comments", "Comments", msImageUnTicked)
        nodX.Tag = msImageUnTicked
        Set nodX = trwQuestions.Nodes.Add(sQuestionKey, tvwChild, sQuestionKey & "|CTCGrade", "CTCGrade", msImageUnTicked)
        nodX.Tag = msImageUnTicked
        Set nodX = trwQuestions.Nodes.Add(sQuestionKey, tvwChild, sQuestionKey & "|LabResult", "LabResult", msImageUnTicked)
        nodX.Tag = msImageUnTicked
        Set nodX = trwQuestions.Nodes.Add(sQuestionKey, tvwChild, sQuestionKey & "|Status", "Status", msImageUnTicked)
        nodX.Tag = msImageUnTicked
        Set nodX = trwQuestions.Nodes.Add(sQuestionKey, tvwChild, sQuestionKey & "|TimeStamp", "TimeStamp", msImageUnTicked)
        nodX.Tag = msImageUnTicked
        Set nodX = trwQuestions.Nodes.Add(sQuestionKey, tvwChild, sQuestionKey & "|UserName", "UserName", msImageUnTicked)
        nodX.Tag = msImageUnTicked
        Set nodX = trwQuestions.Nodes.Add(sQuestionKey, tvwChild, sQuestionKey & "|ValueCode", "ValueCode", msImageUnTicked)
        nodX.Tag = msImageUnTicked
        rsQuestions.MoveNext
    Loop

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "TreeViewQuestions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub mnuVOutputOptions_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    frmQueryOptions.Show vbModal

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuVOutputOptions_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub optAndBands_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    If optAndBands.Value = True Then
        txtBandType.Text = "[AND] Filter Bands connected by OR"
        'flag query as changed
        gbQueryChanged = True
    End If
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optAndBands_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub optDisplayOutput_Click()
'--------------------------------------------------------------------
    
    On Error GoTo Errhandler
    
    gbDisplayOutPut = True
    gbNotDisplayedNotSaved = False

    If txtResultsCount.Text <> "" Then
        If CLng(txtResultsCount.Text) > 0 Then
            Call HourglassOn
            Call DisableRunSaveDisplay
            Call DisplayOutPut
            If gbCancelled Then
                Call HourglassOff
                Exit Sub
            End If
            Call EnableRunSaveDisplay
            Call HourglassOff
        End If
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optDisplayOutput_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub optDoNotDisplayOutput_Click()
'--------------------------------------------------------------------
    
    On Error GoTo Errhandler
    
    gbDisplayOutPut = False
    Call ClearGrid

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optDoNotDisplayOutput_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub optORBands_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    If optORBands.Value = True Then
        txtBandType.Text = "[OR] Filter Bands connected by AND"
        'flag query as changed
        gbQueryChanged = True
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optORBands_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub SelectFilterBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'If Left Button is down update Top Property of SelectFilterBar and call refresh
    If (Button = vbLeftButton) Then
        SelectFilterBar.Top = SelectFilterBar.Top + Y
        'Me.Refresh
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SelectFilterBar_MouseMove")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub SelectFilterBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'Check that SelectFilterBar has not been dragged to far to the Top
    If (SelectFilterBar.Top < 1500) Then
        SelectFilterBar.Top = 1500
    End If
    
    'Check that SelectFilTerBar has not been dragged to far to the bottom
    If (SelectFilterBar.Top > FilterOutputBar.Top - 1860) Then
        SelectFilterBar.Top = FilterOutputBar.Top - 1860
    End If
    
    trwQuestions.Height = SelectFilterBar.Top - trwQuestions.Top - 100
    txtQueryText.Height = SelectFilterBar.Top - txtQueryText.Top - 100
    lblFilter.Top = SelectFilterBar.Top + 200
    cmdNew.Top = lblFilter.Top
    cmdAdd.Top = lblFilter.Top
    cmdDelete.Top = lblFilter.Top
    cmdEdit.Top = lblFilter.Top
    cmdCancel.Top = lblFilter.Top
    txtBandType.Top = lblFilter.Top
    lblBandNo.Top = lblFilter.Top + lblFilter.Height + 100
    lstFilterText.Top = lblBandNo.Top
    txtBandNo.Top = lblBandNo.Top
    lblOperand.Top = lblBandNo.Top
    cboOperand.Top = lblOperand.Top
    cboOperator.Top = cboOperand.Top + cboOperand.Height + 100
    lblOperator.Top = cboOperator.Top
    fraBandType.Top = cboOperator.Top
    txtCriteria.Top = cboOperator.Top + cboOperator.Height + 100
    cboCatCodes.Top = txtCriteria.Top
    lblCriteria.Top = txtCriteria.Top
    lstFilterText.Height = FilterOutputBar.Top - lstFilterText.Top - 100

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SelectFilterBar_MouseUp")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub trwQuestions_NodeClick(ByVal Node As MSComctlLib.Node)
'--------------------------------------------------------------------
'This sub controls the toggling of a nodes .tag and .image between unticked and ticked.
'While doing this it checks for inconsistant selections:-
'   All questions on a study in combination with any other selection
'   All questions on a visit in combination with a specific question on that visit
'   All questions on a Form in combination with a specific question on that Form
'The consistency checks are achieved by using the collections mcolSelectedForms and mcolSelectedVisits
'--------------------------------------------------------------------
Dim asNodeKeys() As String
Dim nQuestionCount As Integer
Dim nPrevQAMask As Integer
Dim sAttribute As String

    On Error GoTo Errhandler

    'Place the node key into an array
    asNodeKeys = Split(Node.Key, "|")
    'check for the the nodes that can never be ticked "S|V", "S|F", "S|Q"
    If UBound(asNodeKeys) = 1 Then Exit Sub
    
    'Must handle clicks on whole study first
    If Node.Key = "S" Then
        If Node.Tag = msImageTicked Then
            'Allow Study.All questions to be unticked
            Node.Image = msImageUnTicked
            Node.Tag = msImageUnTicked
            'flag query as changed
            gbQueryChanged = True
            mbStudyAllQuestions = False
            RefreshQueryText
            Exit Sub
        Else
            'Check for no other nodes having been ticked BY LOOPING THROUGH TREE LOOKING FOR Ticked Node
            If TreeNodesClicked Then
                'disallow ticking of Study.all questions when something below has been ticked
                Exit Sub
            Else
                Node.Image = msImageTicked
                Node.Tag = msImageTicked
                'flag query as changed
                gbQueryChanged = True
                mbStudyAllQuestions = True
                RefreshQueryText
                Exit Sub
            End If
        End If
    Else
        'Check for Study.All Questions being set and disallow all other ticks
        If mbStudyAllQuestions Then
            Exit Sub
        End If
    End If
    
    'Validate/Process clicks on Visits and Visit Questions
    If asNodeKeys(1) = "V" Then
        nQuestionCount = mcolSelectedVisits(Str(asNodeKeys(2)))
        'distinquish between a Visit.All questions and Visit.Specific question click
        If UBound(asNodeKeys) = 2 Then
            'node.key is of form "S|V|VisitId"
            If nQuestionCount = 0 Or nQuestionCount = 999 Then
                'allow the ticking/unticking of a Visit
                If Node.Tag = msImageUnTicked Then
                    Node.Image = msImageTicked
                    Node.Tag = msImageTicked
                    'flag query as changed
                    gbQueryChanged = True
                    mcolSelectedVisits.Remove Str(asNodeKeys(2))
                    mcolSelectedVisits.Add 999, Str(asNodeKeys(2))
                Else
                    Node.Image = msImageUnTicked
                    Node.Tag = msImageUnTicked
                    'flag query as changed
                    gbQueryChanged = True
                    mcolSelectedVisits.Remove Str(asNodeKeys(2))
                    mcolSelectedVisits.Add 0, Str(asNodeKeys(2))
                End If
            Else
                'disallow ticking, because Questions have been ticked below
                Exit Sub
            End If
        Else
            'node.key is of form "S|V|VisitId|QuestionId"
            If nQuestionCount = 999 Then
                'disallow ticking of questions below a ticked Visit
                Exit Sub
            Else
                'allow ticking/unticking of question below Visit
                If Node.Tag = msImageUnTicked Then
                    Node.Image = msImageTicked
                    Node.Tag = msImageTicked
                    'flag query as changed
                    gbQueryChanged = True
                    mcolSelectedVisits.Remove Str(asNodeKeys(2))
                    mcolSelectedVisits.Add nQuestionCount + 1, Str(asNodeKeys(2))
                Else
                    Node.Image = msImageUnTicked
                    Node.Tag = msImageUnTicked
                    'flag query as changed
                    gbQueryChanged = True
                    mcolSelectedVisits.Remove Str(asNodeKeys(2))
                    mcolSelectedVisits.Add nQuestionCount - 1, Str(asNodeKeys(2))
                End If
            End If
        End If
    End If
    
    'Validate/Process clicks on Forms and Form Questions
    If asNodeKeys(1) = "F" Then
        nQuestionCount = mcolSelectedForms(Str(asNodeKeys(2)))
        'distinquish between a Form.All questions and Form.Specific question click
        If UBound(asNodeKeys) = 2 Then
            'node.key is of form "S|F|FormId"
            If nQuestionCount = 0 Or nQuestionCount = 999 Then
                'allow the ticking/unticking of a form
                If Node.Tag = msImageUnTicked Then
                    Node.Image = msImageTicked
                    Node.Tag = msImageTicked
                    'flag query as changed
                    gbQueryChanged = True
                    mcolSelectedForms.Remove Str(asNodeKeys(2))
                    mcolSelectedForms.Add 999, Str(asNodeKeys(2))
                Else
                    Node.Image = msImageUnTicked
                    Node.Tag = msImageUnTicked
                    'flag query as changed
                    gbQueryChanged = True
                    mcolSelectedForms.Remove Str(asNodeKeys(2))
                    mcolSelectedForms.Add 0, Str(asNodeKeys(2))
                End If
            Else
                'disallow ticking, because Questions have been ticked below
                Exit Sub
            End If
        Else
            'node.key is of form "S|F|FormId|QuestionId"
            If nQuestionCount = 999 Then
                'disallow ticking of questions below a ticked form
                Exit Sub
            Else
                'allow ticking/unticking of question below form
                If Node.Tag = msImageUnTicked Then
                    Node.Image = msImageTicked
                    Node.Tag = msImageTicked
                    'flag query as changed
                    gbQueryChanged = True
                    mcolSelectedForms.Remove Str(asNodeKeys(2))
                    mcolSelectedForms.Add nQuestionCount + 1, Str(asNodeKeys(2))
                Else
                    Node.Image = msImageUnTicked
                    Node.Tag = msImageUnTicked
                    'flag query as changed
                    gbQueryChanged = True
                    mcolSelectedForms.Remove Str(asNodeKeys(2))
                    mcolSelectedForms.Add nQuestionCount - 1, Str(asNodeKeys(2))
                End If
            End If
        End If
    End If
    
    'Process clicks on Questions and Question attributes
    If asNodeKeys(1) = "Q" Then
        nPrevQAMask = mcolQuestionAttributes(Str(asNodeKeys(2)))
        If UBound(asNodeKeys) = 2 Then
            'node.key is of form "S|Q|DataItemId"
            If Node.Tag = msImageUnTicked Then
                Node.Image = msImageTicked
                Node.Tag = msImageTicked
                'flag query as changed
                gbQueryChanged = True
                mcolQuestionAttributes.Remove Str(asNodeKeys(2))
                mcolQuestionAttributes.Add nPrevQAMask + mnMASK_RESPONSEVALUE, Str(asNodeKeys(2))
            Else
                'only allow the question to be unticked if no attributes have been selected.
                'This will only occur when nPrevQAMASK = 1
                If nPrevQAMask = 1 Then
                    Node.Image = msImageUnTicked
                    Node.Tag = msImageUnTicked
                    'flag query as changed
                    gbQueryChanged = True
                    mcolQuestionAttributes.Remove Str(asNodeKeys(2))
                    mcolQuestionAttributes.Add nPrevQAMask - mnMASK_RESPONSEVALUE, Str(asNodeKeys(2))
                End If
            End If
        Else
            'node.key is of form "S|Q|DataItemId|Attribute"
            sAttribute = asNodeKeys(3)
            If Node.Tag = msImageUnTicked Then
                Node.Image = msImageTicked
                Node.Tag = msImageTicked
                'flag query as changed
                gbQueryChanged = True
                mcolQuestionAttributes.Remove Str(asNodeKeys(2))
                Select Case sAttribute
                Case "Comments"
                    mcolQuestionAttributes.Add nPrevQAMask + mnMASK_COMMENTS, Str(asNodeKeys(2))
                Case "CTCGrade"
                    mcolQuestionAttributes.Add nPrevQAMask + mnMASK_CTCGRADE, Str(asNodeKeys(2))
                Case "LabResult"
                    mcolQuestionAttributes.Add nPrevQAMask + mnMASK_LABRESULT, Str(asNodeKeys(2))
                Case "Status"
                    mcolQuestionAttributes.Add nPrevQAMask + mnMASK_STATUS, Str(asNodeKeys(2))
                Case "TimeStamp"
                    mcolQuestionAttributes.Add nPrevQAMask + mnMASK_TIMESTAMP, Str(asNodeKeys(2))
                Case "UserName"
                    mcolQuestionAttributes.Add nPrevQAMask + mnMASK_USERNAME, Str(asNodeKeys(2))
                Case "ValueCode"
                    mcolQuestionAttributes.Add nPrevQAMask + mnMASK_VALUECODE, Str(asNodeKeys(2))
                End Select
                'Automatically tick a question's ResponseValue if it has not already been ticked
                If (mcolQuestionAttributes(Str(asNodeKeys(2))) And mnMASK_RESPONSEVALUE) = 0 Then
                    trwQuestions.Nodes("S|Q|" & asNodeKeys(2)).Image = msImageTicked
                    trwQuestions.Nodes("S|Q|" & asNodeKeys(2)).Tag = msImageTicked
                    nPrevQAMask = mcolQuestionAttributes(Str(asNodeKeys(2)))
                    mcolQuestionAttributes.Remove Str(asNodeKeys(2))
                    mcolQuestionAttributes.Add nPrevQAMask + mnMASK_RESPONSEVALUE, Str(asNodeKeys(2))
                End If
            Else
                Node.Image = msImageUnTicked
                Node.Tag = msImageUnTicked
                'flag query as changed
                gbQueryChanged = True
                mcolQuestionAttributes.Remove Str(asNodeKeys(2))
                Select Case sAttribute
                Case "Comments"
                    mcolQuestionAttributes.Add nPrevQAMask - mnMASK_COMMENTS, Str(asNodeKeys(2))
                Case "CTCGrade"
                    mcolQuestionAttributes.Add nPrevQAMask - mnMASK_CTCGRADE, Str(asNodeKeys(2))
                Case "LabResult"
                    mcolQuestionAttributes.Add nPrevQAMask - mnMASK_LABRESULT, Str(asNodeKeys(2))
                Case "Status"
                    mcolQuestionAttributes.Add nPrevQAMask - mnMASK_STATUS, Str(asNodeKeys(2))
                Case "TimeStamp"
                    mcolQuestionAttributes.Add nPrevQAMask - mnMASK_TIMESTAMP, Str(asNodeKeys(2))
                Case "UserName"
                    mcolQuestionAttributes.Add nPrevQAMask - mnMASK_USERNAME, Str(asNodeKeys(2))
                Case "ValueCode"
                    mcolQuestionAttributes.Add nPrevQAMask - mnMASK_VALUECODE, Str(asNodeKeys(2))
                End Select
            End If
        End If
    End If
  
    RefreshQueryText

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "trwQuestions_NodeClick")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub RefreshQueryText()
'--------------------------------------------------------------------
'There are 8 types of Select statement in the Select Text:-
'   [S]STUDYNAME
'   [S]STUDYNAME.All Questions
'   [V]VISITCODE.All Questions
'   [V]VISITCODE.QUESTIONNAME
'   [F]FORMCODE.All Questions
'   [F]FORMCODE.QUESTIONNAME
'   [Q]QUESTIONCODE.ResponseValue
'   [Q]QUESTIONCODE.Other AttributeName
'They will always appear in the above order
'--------------------------------------------------------------------
Dim nodX As MSComctlLib.Node
Dim sQueryText As String
Dim sQueryLine As String
Dim asNodeKeys() As String
   
    On Error GoTo Errhandler
    
    'disable the cmdRunQuery button. It will be enabled as soon as a ticked tree node is found
    cmdRunQuery.Enabled = False
    mnuRRunQuery.Enabled = False
    cmdSaveOutPut.Enabled = False
    mnuFSave.Enabled = False
    mnuFSaveAs.Enabled = False
    mnuFPrint.Enabled = False
    
    sQueryText = ""
    For Each nodX In trwQuestions.Nodes
        'check for the need of a Study label if study.all questions is not set
        If nodX.Index = 1 And nodX.Tag = msImageUnTicked Then
            sQueryLine = "[S]" & nodX.Text
            sQueryText = sQueryText & sQueryLine & vbNewLine
        End If
            
        If nodX.Tag = msImageTicked Then
            'enable the cmdRunQuery button
            cmdRunQuery.Enabled = True
            mnuRRunQuery.Enabled = True
            mnuFSave.Enabled = True
            mnuFSaveAs.Enabled = True
            mnuFPrint.Enabled = True
            asNodeKeys = Split(nodX.Key, "|")
            If UBound(asNodeKeys) = 0 Then
                'nodX.key = "S"
                sQueryLine = "[S]" & nodX.Text & ".All Questions"
            ElseIf asNodeKeys(1) = "V" Then
                'nodeX.Key starts "S|V"
                'distinquish between Visit.All Questions and Visit.Specific Question
                If UBound(asNodeKeys) = 2 Then
                    'nodX.Key is "S|V|VisitId"
                    sQueryLine = "[V]" & nodX.Text & ".All Questions"
                Else
                    'nodX.Key is "S|V|VisitId|QuestionId"
                    sQueryLine = "[V]" & nodX.Parent & "." & nodX.Text
                End If
            ElseIf asNodeKeys(1) = "F" Then
                'nodeX.Key starts "S|F"
                'distinquish between Form.All Questions and Form.Specific Question
                If UBound(asNodeKeys) = 2 Then
                    'nodX.Key is "S|F|FormId"
                    sQueryLine = "[F]" & nodX.Text & ".All Questions"
                Else
                    'nodX.Key is "S|F|FormId|QuestionId"
                    sQueryLine = "[F]" & nodX.Parent & "." & nodX.Text
                End If
            ElseIf asNodeKeys(1) = "Q" Then
                'nodeX.Key starts "S|Q" Then
                'distinquish between a Question and a Question|Attribute
                If UBound(asNodeKeys) = 2 Then
                    'nodX.Key is "S|Q|QuestionId"
                    sQueryLine = "[Q]" & nodX.Text & ".ResponseValue"
                Else
                    'nodX.Key is "S|Q|QuestionId|Attribute"
                    sQueryLine = "[Q]" & nodX.Parent & "." & nodX.Text
                    'assess which result attributes might have been selected
                    Select Case nodX.Text
                    Case "Comments"
                        mbComments = True
                    Case "CTCGrade"
                        mbCTCGrade = True
                    Case "LabResult"
                        mbLabResult = True
                    Case "Status"
                        mbStatus = True
                    Case "TimeStamp"
                        mbTimeStamp = True
                    Case "UserName"
                        mbUserName = True
                    Case "ValueCode"
                        mbValueCode = True
                    End Select
                End If
            End If

            sQueryText = sQueryText & sQueryLine & vbNewLine
        End If
    Next nodX
    
    'remove the  last Carriage Return/Line feed from squerytext
    sQueryText = Mid(sQueryText, 1, Len(sQueryText) - 2)
       
    txtQueryText.Text = sQueryText

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshQueryText")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub CreateQuerySQL(ByRef sDataSQL As String, _
                            ByRef sDataItemNamesSQL As String, _
                            ByRef sSubjectRecordsSQL As String)
'--------------------------------------------------------------------
'This sub creates 2 strings:-
'   sDataSQL will retrieve the currently specified data
'   sDataItemNamesSQL will retrieve the DataItemNames of the specified data
'   sSubjectRecordsSQL will retrive the number of Subject records
'
'Changed Mo Morris 22/8/2002, Stemming from SR4874
'For performance reasons sDataItemNamesSQL has changed. It used to
'accesss DataItemResponse and be a SELECT DISTINCT query, it now no
'longer accesses DataItemResponse and runs many times faster.
'sSQLNames has been added. It is similar to sSQL but the table names are different
'Changed Mo 27/1/2003 RQG changes
' NCJ 20 Dec 05 - Added DataItemCase (the Partial Date flag)
'--------------------------------------------------------------------
Dim sQueryFor As String
Dim sSQL As String
Dim sSQLNames As String
Dim nodX As MSComctlLib.Node
Dim asNodeKeys() As String
Dim nCount As Integer
Dim anVisitIDs() As Integer
Dim anFormIDs() As Integer
Dim anQuestionIDs() As Integer
Dim i As Integer
Dim nQuestionsVisitId As Integer
Dim nPrevQuestionsVisitId As Integer
Dim nQuestionsFormId As Integer
Dim nPrevQuestionsFormId As Integer
Dim bFirstSelection As Boolean
Dim nPrevQuestionId As Integer

    On Error GoTo Errhandler

    'Create and build up a QueryFor string based on which result attributes have been used
    sQueryFor = "TrialSite, PersonId, VisitId, VisitCycleNumber, CRFPageId, CRFPageCycleNumber, DataItemId, ResponseValue, ValueCode, ResponseStatus, RepeatNumber"
    If mbComments Then sQueryFor = sQueryFor & ", Comments"
    If mbCTCGrade Then sQueryFor = sQueryFor & ", CTCGrade"
    If mbLabResult Then sQueryFor = sQueryFor & ", LabResult"
    If mbStatus Then sQueryFor = sQueryFor & ", ResponseStatus"
    If mbTimeStamp Then sQueryFor = sQueryFor & ", ResponseTimeStamp"
    If mbUserName Then sQueryFor = sQueryFor & ", UserName"
    If mbValueCode Then sQueryFor = sQueryFor & ", ValueCode"
    
    sDataSQL = "SELECT " & sQueryFor & " FROM DataItemResponse " _
        & "WHERE ClinicalTrialId = " & glSelectedTrialId

    ' NCJ 20 Dec 05 - Added DataItemCase (the Partial Date flag)
    sDataItemNamesSQL = "SELECT StudyVisit.VisitId, StudyVisit.VisitCode, StudyVisit.VisitOrder, " _
        & "CRFPage.CRFPageId, CRFPage.CRFPageCode, CRFPage.CRFPageOrder, " _
        & "DataItem.DataItemId, DataItem.DataItemCode, CRFElement.FieldOrder, CRFElement.QGroupFieldOrder, " _
        & "DataItem.DataType, DataItem.DataItemFormat, DataItem.DataItemLength, DataItem.DataItemCase " _
        & "FROM StudyVisit, StudyVisitCRFPage, CRFPage, DataItem, CRFElement " _
        & "WHERE StudyVisit.ClinicalTrialId = StudyVisitCRFPage.ClinicalTrialId " _
        & "AND StudyVisit.VisitId = StudyVisitCRFPage.VisitId " _
        & "AND StudyVisit.ClinicalTrialId = CRFPage.ClinicalTrialId " _
        & "AND StudyVisitCRFPage.CRFPageId = CRFPage.CRFPageId " _
        & "AND StudyVisit.ClinicalTrialId = DataItem.ClinicalTrialId " _
        & "AND StudyVisit.ClinicalTrialId = CRFElement.ClinicalTrialId " _
        & "AND CRFPage.CRFPageId = CRFElement.CRFPageId " _
        & "AND DataItem.DataItemId = CRFElement.DataItemId " _
        & "AND StudyVisit.ClinicalTrialId = " & glSelectedTrialId
        'note that the linkage to table CRFElement is based on
        'ClinicalTrialId, CRFPageId and DataItemId (not CRFElementId, which can change when a study is edited)
        
    sSubjectRecordsSQL = "SELECT DISTINCT TrialSite, PersonId, VisitCycleNumber, CRFPageCycleNumber, RepeatNumber " _
        & "FROM DataItemResponse " _
        & "WHERE ClinicalTrialId = " & glSelectedTrialId
    
    'Check for Study.All Questions
    If mbStudyAllQuestions Then
        'Add Filter elements to sDataSQL and sSubjectRecordsSQL
        Call FilterThisSQL(sDataSQL, sSubjectRecordsSQL)
        sDataSQL = sDataSQL & " ORDER BY TrialSite, PersonId, VisitCycleNumber, CRFPageCycleNumber, RepeatNumber"
        sDataItemNamesSQL = sDataItemNamesSQL & " ORDER BY StudyVisit.VisitOrder, CRFPage.CRFPageOrder, CRFElement.FieldOrder, CRFElement.QGroupFieldOrder"
        Exit Sub
    End If
        
    'initialise bFirstSelection which controls the switch from AND to OR
    bFirstSelection = True
    
    sSQL = ""
    sSQLNames = ""

    'Check for Visit.All Questions
    nCount = 0
    For Each nodX In trwQuestions.Nodes
        If nodX.Tag = msImageTicked Then
            asNodeKeys = Split(nodX.Key, "|")
            If asNodeKeys(1) = "V" And UBound(asNodeKeys) = 2 Then
                nCount = nCount + 1
                ReDim Preserve anVisitIDs(nCount)
                anVisitIDs(nCount) = asNodeKeys(2)
            End If
        End If
    Next nodX
    If nCount > 0 Then
        If bFirstSelection Then
            sSQL = sSQL & " AND (("
            sSQLNames = sSQLNames & " AND (("
        Else
            sSQL = sSQL & " OR ("
            sSQLNames = sSQLNames & " OR ("
        End If
        'set the first selection boolean to false
        bFirstSelection = False
        If nCount = 1 Then
            sSQL = sSQL & "VisitId = " & anVisitIDs(1)
            sSQLNames = sSQLNames & "StudyVisit.VisitId = " & anVisitIDs(1)
        Else
            sSQL = sSQL & "VisitId IN ("
            sSQLNames = sSQLNames & "StudyVisit.VisitId IN ("
            For i = 1 To nCount
                sSQL = sSQL & anVisitIDs(i)
                sSQLNames = sSQLNames & anVisitIDs(i)
                If i <> nCount Then
                    sSQL = sSQL & ","
                    sSQLNames = sSQLNames & ","
                End If
            Next i
            sSQL = sSQL & ")"
            sSQLNames = sSQLNames & ")"
        End If
        sSQL = sSQL & ")"
        sSQLNames = sSQLNames & ")"
    End If
    
    'Check for Visit.Single/Several Questions
    'initialize the previous visitid
    nPrevQuestionsVisitId = 0
    For Each nodX In trwQuestions.Nodes
        If nodX.Tag = msImageTicked Then
            asNodeKeys = Split(nodX.Key, "|")
            'search for ticked questions within a visit node
            If asNodeKeys(1) = "V" And UBound(asNodeKeys) = 3 Then
                'extract the visitid
                nQuestionsVisitId = asNodeKeys(2)
                'initialization after the first found tick
                If nPrevQuestionsVisitId = 0 Then
                    nPrevQuestionsVisitId = nQuestionsVisitId
                    nCount = 0
                End If
                'When the visitId changes process the previous one
                If nQuestionsVisitId <> nPrevQuestionsVisitId Then
                    'decide on use of AND or OR
                    If bFirstSelection Then
                        sSQL = sSQL & " AND (("
                        sSQLNames = sSQLNames & " AND (("
                    Else
                        sSQL = sSQL & " OR ("
                        sSQLNames = sSQLNames & " OR ("
                    End If
                    'set the first selection boolean to false
                    bFirstSelection = False
                    sSQL = sSQL & "VisitId = " & nPrevQuestionsVisitId & " AND DataItemId "
                    sSQLNames = sSQLNames & "StudyVisit.VisitId = " & nPrevQuestionsVisitId & " AND DataItem.DataItemId "
                    If nCount = 1 Then
                        sSQL = sSQL & "= " & anQuestionIDs(1)
                        sSQLNames = sSQLNames & "= " & anQuestionIDs(1)
                    Else
                        sSQL = sSQL & "IN ("
                        sSQLNames = sSQLNames & "IN ("
                        For i = 1 To nCount
                            sSQL = sSQL & anQuestionIDs(i)
                            sSQLNames = sSQLNames & anQuestionIDs(i)
                            If i <> nCount Then
                                sSQL = sSQL & ","
                                sSQLNames = sSQLNames & ","
                            End If
                        Next i
                        sSQL = sSQL & ")"
                        sSQLNames = sSQLNames & ")"
                    End If
                    sSQL = sSQL & ")"
                    sSQLNames = sSQLNames & ")"
                    'prepare the PrevVisit variable and initialize the question counter
                    nPrevQuestionsVisitId = nQuestionsVisitId
                    nCount = 0
                End If
                'increment the questions counter and store the currently ticked questionid
                nCount = nCount + 1
                ReDim Preserve anQuestionIDs(nCount)
                anQuestionIDs(nCount) = asNodeKeys(3)
            End If
        End If
    Next nodX
    'Process the last found visit
    If nPrevQuestionsVisitId <> 0 Then
        If bFirstSelection Then
            sSQL = sSQL & " AND (("
            sSQLNames = sSQLNames & " AND (("
        Else
            sSQL = sSQL & " OR ("
            sSQLNames = sSQLNames & " OR ("
        End If
        'set the first selection boolean to false
        bFirstSelection = False
        sSQL = sSQL & "VisitId = " & nQuestionsVisitId & " AND DataItemId "
        sSQLNames = sSQLNames & "StudyVisit.VisitId = " & nQuestionsVisitId & " AND DataItem.DataItemId "
        If nCount = 1 Then
            sSQL = sSQL & "= " & anQuestionIDs(1)
            sSQLNames = sSQLNames & "= " & anQuestionIDs(1)
        Else
            sSQL = sSQL & "IN ("
            sSQLNames = sSQLNames & "IN ("
            For i = 1 To nCount
                sSQL = sSQL & anQuestionIDs(i)
                sSQLNames = sSQLNames & anQuestionIDs(i)
                If i <> nCount Then
                    sSQL = sSQL & ","
                    sSQLNames = sSQLNames & ","
                End If
            Next i
            sSQL = sSQL & ")"
            sSQLNames = sSQLNames & ")"
        End If
        sSQL = sSQL & ")"
        sSQLNames = sSQLNames & ")"
    End If
    
    'Check for Form.All Questions
    nCount = 0
    For Each nodX In trwQuestions.Nodes
        If nodX.Tag = msImageTicked Then
            asNodeKeys = Split(nodX.Key, "|")
            If asNodeKeys(1) = "F" And UBound(asNodeKeys) = 2 Then
                nCount = nCount + 1
                ReDim Preserve anFormIDs(nCount)
                anFormIDs(nCount) = asNodeKeys(2)
            End If
        End If
    Next nodX
    If nCount > 0 Then
        If bFirstSelection Then
            sSQL = sSQL & " AND (("
            sSQLNames = sSQLNames & " AND (("
        Else
            sSQL = sSQL & " OR ("
            sSQLNames = sSQLNames & " OR ("
        End If
        'set the first selection boolean to false
        bFirstSelection = False
        If nCount = 1 Then
            sSQL = sSQL & "CRFPageId = " & anFormIDs(1)
            sSQLNames = sSQLNames & "CRFPage.CRFPageId = " & anFormIDs(1)
        Else
            sSQL = sSQL & "CRFPageId IN ("
            sSQLNames = sSQLNames & "CRFPage.CRFPageId IN ("
            For i = 1 To nCount
                sSQL = sSQL & anFormIDs(i)
                sSQLNames = sSQLNames & anFormIDs(i)
                If i <> nCount Then
                    sSQL = sSQL & ","
                    sSQLNames = sSQLNames & ","
                End If
            Next i
            sSQL = sSQL & ")"
            sSQLNames = sSQLNames & ")"
        End If
        sSQL = sSQL & ")"
        sSQLNames = sSQLNames & ")"
    End If
    
    'Check for Form.Single/Several Questions
    'initialize the previous Formid
    nPrevQuestionsFormId = 0
    For Each nodX In trwQuestions.Nodes
        If nodX.Tag = msImageTicked Then
            asNodeKeys = Split(nodX.Key, "|")
            'search for ticked questions within a Form node
            If asNodeKeys(1) = "F" And UBound(asNodeKeys) = 3 Then
                'extract the Formid
                nQuestionsFormId = asNodeKeys(2)
                'initialization after the first found tick
                If nPrevQuestionsFormId = 0 Then
                    nPrevQuestionsFormId = nQuestionsFormId
                    nCount = 0
                End If
                'When the FormId changes process the previous one
                If nQuestionsFormId <> nPrevQuestionsFormId Then
                    'decide on use of AND or OR
                    If bFirstSelection Then
                        sSQL = sSQL & " AND (("
                        sSQLNames = sSQLNames & " AND (("
                    Else
                        sSQL = sSQL & " OR ("
                        sSQLNames = sSQLNames & " OR ("
                    End If
                    'set the first selection boolean to false
                    bFirstSelection = False
                    sSQL = sSQL & "CRFPageId = " & nPrevQuestionsFormId & " AND DataItemId "
                    sSQLNames = sSQLNames & "CRFPage.CRFPageId = " & nPrevQuestionsFormId & " AND DataItem.DataItemId "
                    If nCount = 1 Then
                        sSQL = sSQL & "= " & anQuestionIDs(1)
                        sSQLNames = sSQLNames & "= " & anQuestionIDs(1)
                    Else
                        sSQL = sSQL & "IN ("
                        sSQLNames = sSQLNames & "IN ("
                        For i = 1 To nCount
                            sSQL = sSQL & anQuestionIDs(i)
                            sSQLNames = sSQLNames & anQuestionIDs(i)
                            If i <> nCount Then
                                sSQL = sSQL & ","
                                sSQLNames = sSQLNames & ","
                            End If
                        Next i
                        sSQL = sSQL & ")"
                        sSQLNames = sSQLNames & ")"
                    End If
                    sSQL = sSQL & ")"
                    sSQLNames = sSQLNames & ")"
                    'prepare the PrevForm variable and initialize the question counter
                    nPrevQuestionsFormId = nQuestionsFormId
                    nCount = 0
                End If
                'increment the questions counter and store the currently ticked questionid
                nCount = nCount + 1
                ReDim Preserve anQuestionIDs(nCount)
                anQuestionIDs(nCount) = asNodeKeys(3)
            End If
        End If
    Next nodX
    'Process the last found form
    If nPrevQuestionsFormId <> 0 Then
        If bFirstSelection Then
            sSQL = sSQL & " AND (("
            sSQLNames = sSQLNames & " AND (("
        Else
            sSQL = sSQL & " OR ("
            sSQLNames = sSQLNames & " OR ("
        End If
        'set the first selection boolean to false
        bFirstSelection = False
        sSQL = sSQL & "CRFPageId = " & nQuestionsFormId & " AND DataItemId "
        sSQLNames = sSQLNames & "CRFPage.CRFPageId = " & nQuestionsFormId & " AND DataItem.DataItemId "
        If nCount = 1 Then
            sSQL = sSQL & "= " & anQuestionIDs(1)
            sSQLNames = sSQLNames & "= " & anQuestionIDs(1)
        Else
            sSQL = sSQL & "IN ("
            sSQLNames = sSQLNames & "IN ("
            For i = 1 To nCount
                sSQL = sSQL & anQuestionIDs(i)
                sSQLNames = sSQLNames & anQuestionIDs(i)
                If i <> nCount Then
                    sSQL = sSQL & ","
                    sSQLNames = sSQLNames & ","
                End If
            Next i
            sSQL = sSQL & ")"
            sSQLNames = sSQLNames & ")"
        End If
        sSQL = sSQL & ")"
        sSQLNames = sSQLNames & ")"
    End If
    
    'Check for Question.Single/Several
    nCount = 0
    nPrevQuestionId = 0
    For Each nodX In trwQuestions.Nodes
        If nodX.Tag = msImageTicked Then
            asNodeKeys = Split(nodX.Key, "|")
            'search for ticked question or ticked question attribute
            If asNodeKeys(1) = "Q" Then
                If asNodeKeys(2) <> nPrevQuestionId Then
                    nPrevQuestionId = asNodeKeys(2)
                    'increment the questions counter and store the ticked questionid
                    nCount = nCount + 1
                    ReDim Preserve anQuestionIDs(nCount)
                    anQuestionIDs(nCount) = asNodeKeys(2)
                End If
            End If
        End If
    Next nodX
    If nCount > 0 Then
        If bFirstSelection Then
            sSQL = sSQL & " AND (("
            sSQLNames = sSQLNames & " AND (("
        Else
            sSQL = sSQL & " OR ("
            sSQLNames = sSQLNames & " OR ("
        End If
        'set the first selection boolean to false
        bFirstSelection = False
        If nCount = 1 Then
            sSQL = sSQL & "DataItemId = " & anQuestionIDs(1)
            sSQLNames = sSQLNames & "DataItem.DataItemId = " & anQuestionIDs(1)
        Else
            sSQL = sSQL & "DataItemId IN ("
            sSQLNames = sSQLNames & "DataItem.DataItemId IN ("
            For i = 1 To nCount
                sSQL = sSQL & anQuestionIDs(i)
                sSQLNames = sSQLNames & anQuestionIDs(i)
                If i <> nCount Then
                    sSQL = sSQL & ","
                    sSQLNames = sSQLNames & ","
                End If
            Next i
            sSQL = sSQL & ")"
            sSQLNames = sSQLNames & ")"
        End If
        sSQL = sSQL & ")"
        sSQLNames = sSQLNames & ")"
    End If
    
    'Check for nothing being ticked
    
    'Add a closing parenthesis
    sSQL = sSQL & ")"
    sSQLNames = sSQLNames & ")"
    
    sDataSQL = sDataSQL & sSQL
    sSubjectRecordsSQL = sSubjectRecordsSQL & sSQL

    'Add Filter elements to sDataSQL and sSubjectRecordsSQL
    Call FilterThisSQL(sDataSQL, sSubjectRecordsSQL)
    
    sDataSQL = sDataSQL & " ORDER BY TrialSite, PersonId, VisitCycleNumber, CRFPageCycleNumber, RepeatNumber"
    sDataItemNamesSQL = sDataItemNamesSQL & sSQLNames & " ORDER BY StudyVisit.VisitOrder, CRFPage.CRFPageOrder, CRFElement.FieldOrder, CRFElement.QGroupFieldOrder"
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "CreateQuerySQL")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Function TreeNodesClicked() As Boolean
'--------------------------------------------------------------------
Dim nodX As MSComctlLib.Node
Dim bTickedNode As Boolean

    On Error GoTo Errhandler

    'intialize boolean to false
    bTickedNode = False
    
    For Each nodX In trwQuestions.Nodes
        If nodX.Tag = msImageTicked Then
            bTickedNode = True
            Exit For
        End If
    Next nodX
    
    If bTickedNode Then
        TreeNodesClicked = True
    Else
        TreeNodesClicked = False
    End If

Exit Function
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "TreeNodesClicked")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'---------------------------------------------------------------------
Public Sub GetRegistrySettings()
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    gnRegFormLeft = GetSetting("MACRO_QM", "Options", "FormLeft", Abs(Screen.Width - mnMINFORMWIDTH) / 2)
    gnRegFormTop = GetSetting("MACRO_QM", "Options", "FormTop", Abs(Screen.Height - mnMINFORMHEIGHT) / 2)
    gnRegFormWidth = GetSetting("MACRO_QM", "Options", "FormWidth", mnMINFORMWIDTH)
    gnRegFormHeight = GetSetting("MACRO_QM", "Options", "FormHeight", mnMINFORMHEIGHT)
    gnSelectFilterBarTop = 2635
    gnFilterOutputBarTop = 4495
    gsSVMissing = ""
    gsSVUnobtainable = ""
    gsSVNotApplicable = ""
    gnOutPutType = 0
    gbOutputCategoryCodes = True
    gbDisplayStudyName = False
    gbDisplaySiteCode = True
    gbDisplayLabel = True
    gbDisplayPersonId = True
    gbDisplayVisitCycle = True
    gbDisplayFormCycle = True
    gbDisplayRepeatNumber = True
    gbSplitGrid = False
    gbUseShortCodes = True
    gbDisplayOutPut = False
    'Mo 30/5/2006 Bug 2668
    gbExcludeLabel = False
    'Mo 2/6/2006 Bug 2737
    gnShortCodeLength = 8
    
    'Mo 2/4/2007 MRC15022007
    gbSASInformatColons = False
    gsFileNamePath = ""
    gsFileNameText = ""
    gsFileNameStamp = "DATE"
    
'    gnRegFormLeft = GetSetting("MACRO_QM", "Options", "FormLeft", Abs(Screen.Width / 10))
'    gnRegFormTop = GetSetting("MACRO_QM", "Options", "FormTop", Abs(Screen.Height / 10))
'    gnRegFormWidth = GetSetting("MACRO_QM", "Options", "FormWidth", Abs(Screen.Width / 10 * 8))
'    gnRegFormHeight = GetSetting("MACRO_QM", "Options", "FormHeight", Abs(Screen.Height / 10 * 8))
'    mnRegFormWindowState = GetSetting("IMedCUI", "Options", "WindowState", 0)
'    gnRegFormLeftMax = GetSetting("IMedCUI", "Options", "FormLeftMax", 0)
'    gnRegFormTopMax = GetSetting("IMedCUI", "Options", "FormTopMax", 0)
'    gnRegFormWidthMax = GetSetting("IMedCUI", "Options", "FormWidthMax", Abs(Screen.Width))
'    gnRegFormHeightMax = GetSetting("IMedCUI", "Options", "FormHeightMax", Abs(Screen.Height))

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetRegistrySettings")
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
Private Sub PopulateOperandCombo()
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsDataItems As ADODB.Recordset

    On Error GoTo Errhandler
    
    cboOperand.Clear
    
    'Changed Mo 5/9/2002, Add Site, Subject Label & PersonIds Filter Functions
    cboOperand.AddItem "[SubjectId]"
    cboOperand.ItemData(cboOperand.NewIndex) = -1
    cboOperand.AddItem "[Site]"
    cboOperand.ItemData(cboOperand.NewIndex) = -2
    cboOperand.AddItem "[SubjectLabel]"
    cboOperand.ItemData(cboOperand.NewIndex) = -3
    
    sSQL = "SELECT DataItemId, DataItemCode " _
        & "FROM DataItem " _
        & "WHERE ClinicalTrialId = " & glSelectedTrialId _
        & " ORDER BY DataItemCode"
        
    Set rsDataItems = New ADODB.Recordset
    rsDataItems.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Do Until rsDataItems.EOF
        cboOperand.AddItem rsDataItems!DataItemCode
        cboOperand.ItemData(cboOperand.NewIndex) = rsDataItems!DataItemId
        rsDataItems.MoveNext
    Loop

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PopulateOperandCombo")
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
Private Sub EnableAddButton()
'---------------------------------------------------------------------
' Enable the Add/Change Factor button (assume in Edit mode)
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    '"IS NULL" and "IS NOT NULL" special validation
    If ((cboOperator.Text = " IS NOT NULL (Response Exists)") Or (cboOperator.Text = " IS NULL (No Response)")) Then
        If ((cboOperand.Text = "[SubjectId]") Or (cboOperand.Text = "[Site]") Or (cboOperand.Text = "[SubjectLabel]")) Then
            'Disallow "IS NULL" and "IS NOT NULL" for operands "[SubjectId]", "[Site]", and "[SubjectLabel]"
            cboOperator.ListIndex = -1
        Else
            'If "IS NULL" or "IS NOT NULL" have been corretly selected clear the Criteria/CatCodes controls
            txtCriteria.Text = ""
            cboCatCodes.ListIndex = -1
        End If
    End If
    
    If ((((txtCriteria.Visible = True) And (Trim(txtCriteria.Text) <> "")) Or ((cboCatCodes.Visible = True) And (cboCatCodes.ListIndex <> -1))) _
        And (Trim(txtBandNo.Text) <> "") _
        And (cboOperand.ListIndex <> -1) _
        And (cboOperator.ListIndex <> -1)) _
        Then
            cmdAdd.Enabled = True
    ElseIf ((cboOperator.Text = " IS NOT NULL (Response Exists)") Or (cboOperator.Text = " IS NULL (No Response)")) _
        And (Trim(txtBandNo.Text) <> "") _
        And (cboOperand.ListIndex <> -1) _
        Then
            cmdAdd.Enabled = True
    Else
            cmdAdd.Enabled = False
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "EnableAddButton")
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
Private Sub txtBandNo_Change()
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'Check for the entry of a single digit numeric
    Select Case txtBandNo.Text
    Case 1, 2, 3, 4, 5, 6, 7, 8, 9
        'Do nothing its a valid entry
        'check the add butttons enabled status
        EnableAddButton
    Case Else
        'clear the invalid entry down
        txtBandNo.Text = ""
        'check the add butttons enabled status
        EnableAddButton
    End Select

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtBandNo_Change")
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
Private Sub txtCriteria_Change()
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    EnableAddButton
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtCriteria_Change")
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
Private Sub FilterThisSQL(ByRef sDataSQL As String, _
                            ByRef sSubjectRecordsSQL As String)
'---------------------------------------------------------------------
'The general approach to filtering is that of querying the DataItemResponse Table
'for TrialSite & PersonId combinations that fullfill the filtering criteria.
'Some of the Special Filter Functions [SubjectId],[Site],[SuvbjectLabel] query other tables.
'The resulting list of TrialSite & PersonIds, combined into a single field is
'then added to the SELECT query in the form:-
'   AND TrialSite & PersonID IN(SiteID,SiteID,SiteID.....)
'
'A filter element is of the form OPERAND OPERATOR CRITERIA
'   (e.g. Age > 30, Sex = 'm')
'In SQL terms the above examples become
'   (DataItemId = 10123 AND ResponseValue > 30), (DataItemID = 10056 AND ResponseValue = 'm')
'
'Each filter element is within a band.
'If AND-Bands have been selected each filter within a band will be connected by ANDs.
'e.g. ((Filter) AND (Filter) AND (Filter))
'
'If OR-Bands have been selected each filter within a band will be connected by ORs.
'e.g. ((Filter) OR (Filter) OR (Filter))
'
'AND-Bands are connected by ORs
'e.g. ((Band1) OR (Band2) OR (Band3))
'
'OR-Bands are connected by ANDs
'e.g. ((Band1) AND (Band2) AND (Band3))
'
'Each filter element is run as a separate SQL statement and placed in its
'own list of SitePersonIds. The individual lists of SitePersonIds are then combined into a
'single list containing SitePersonIds that occur in all of the individual lists.
'AND-Band element lists are combined using AND logic.
'OR-Band element lists are combined using OR logic.
'
'AND-Band SitePersonID lists are combined using OR logic.
'(i.e. the combined list encompasses all of the entries in the individual lists)
'
'OR-Band SitePersonID lists are combined using AND logic.
'(i.e. only SitePersonIDs that occur in all lists get into the combined list)
'
'---------------------------------------------------------------------
Dim sSQL As String
Dim sSQLStart As String
Dim i As Integer
Dim sFilterText As String
Dim sBandNo As String
Dim sPrevBandNo As String
Dim sFilter As String
Dim sInternalBandConnector As String
Dim sExternalBandConnector As String
Dim rsFilteredSubjects As ADODB.Recordset
Dim sInString As String
Dim bANDBands As Boolean
Dim colSingleElementSiteIds As Collection
Dim colSingleBandSiteIds As Collection
Dim colAllBandsSiteIds As Collection
Dim bFirstBand As Boolean
Dim bFirstElement As Boolean
Dim bFirstFilterElement As Boolean
Dim vSitePersonId As Variant

    On Error GoTo Errhandler

    'Exit if no Filtering elements exist
    If lstFilterText.ListCount = 0 Then Exit Sub

    'Clear the SitePersonIds collections
    Set colSingleElementSiteIds = New Collection
    Set colSingleBandSiteIds = New Collection
    Set colAllBandsSiteIds = New Collection

    'set up the AND/OR Internal/External band connectors
    If optAndBands.Value = True Then
        bANDBands = True
        sInternalBandConnector = " AND "
        sExternalBandConnector = " OR "
    Else
        bANDBands = False
        sInternalBandConnector = " OR "
        sExternalBandConnector = " AND "
    End If
    
    'Question based filtering SQL statements to start as follows
    sSQLStart = "SELECT DISTINCT TrialSite, PersonId FROM DataItemResponse " _
        & "WHERE ClinicalTrialId = " & glSelectedTrialId & " AND "
    
    bFirstBand = True
    bFirstElement = True
    bFirstFilterElement = True
    sPrevBandNo = "0"
    'Read through and process the filter elements
    For i = 0 To lstFilterText.ListCount - 1
        sFilterText = lstFilterText.List(i)
        'Extract band number from filter Text
        sBandNo = Mid(sFilterText, 2, InStr(sFilterText, "]") - 2)
        'Put filter in SQL format and place within parenthesis
        Select Case lstFilterText.ItemData(i)
        Case -1
            'Its a [SubjectId] Filter Function
            sFilter = "SELECT DISTINCT TrialSite, PersonId FROM TrialSubject " _
                & "WHERE ClinicalTrialId = " & glSelectedTrialId & " AND " _
                & "PersonId " & Mid(sFilterText, InStr(sFilterText, " ") + 1)
        Case -2
            'Its a [Site] Filter Function
            sFilter = "SELECT DISTINCT TrialSite, PersonId FROM TrialSubject " _
                & "WHERE ClinicalTrialId = " & glSelectedTrialId & " AND " _
                & "TrialSite " & Mid(sFilterText, InStr(sFilterText, " ") + 1)
        Case -3
            'Its a [SubjectLabel] Filter Function
           sFilter = "SELECT DISTINCT TrialSite, PersonId FROM TrialSubject " _
                & "WHERE ClinicalTrialId = " & glSelectedTrialId & " AND " _
                & "LocalIdentifier1 " & Mid(sFilterText, InStr(sFilterText, " ") + 1)
        Case Else
            'Using the question's data type determine how the filter element should be executed
            'Numerics require type conversion
            'Category questions require ValueCode
            Select Case DataTypeFromId(glSelectedTrialId, lstFilterText.ItemData(i))
            'Mo 25/10/2005 COD0040
            Case DataType.Text, DataType.Date, DataType.Multimedia, DataType.Thesaurus
                sFilter = "(DataItemId = " & lstFilterText.ItemData(i) & " AND ResponseValue " & Mid(sFilterText, InStr(sFilterText, " ")) & ")"
            Case DataType.IntegerData, DataType.Real, DataType.LabTest
                'Decide between a NULL/NON-NULL and a Numeric Operator
                If (InStr(sFilterText, "IS NULL") = 0) And (InStr(sFilterText, "IS NOT NULL") = 0) Then
                    'the Val function will crash on an empty response so an '0' is concatenated to it
                    'following SQl statement is database specific
                    Select Case goUser.Database.DatabaseType
                    Case MACRODatabaseType.Access
                        sFilter = "(DataItemId = " & lstFilterText.ItemData(i) & " AND VAL('0' & ResponseValue) " & Mid(sFilterText, InStr(sFilterText, " ")) & ")"
                    Case MACRODatabaseType.Oracle80
                        sFilter = "(DataItemId = " & lstFilterText.ItemData(i) & " AND TO_NUMBER('0' || ResponseValue) " & Mid(sFilterText, InStr(sFilterText, " ")) & ")"
                    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
                        'Mo 25/5/2006 Bug 2666, CAST parameter changd from DECIMAL to REAL
                        sFilter = "(DataItemId = " & lstFilterText.ItemData(i) & " AND CAST('0' + ResponseValue as REAL) " & Mid(sFilterText, InStr(sFilterText, " ")) & ")"
                    End Select
                Else
                    sFilter = "(DataItemId = " & lstFilterText.ItemData(i) & " AND ResponseValue " & Mid(sFilterText, InStr(sFilterText, " ")) & ")"
                End If
            Case DataType.Category
                sFilter = "(DataItemId = " & lstFilterText.ItemData(i) & " AND ValueCode " & Mid(sFilterText, InStr(sFilterText, " ")) & ")"
            End Select
            sFilter = sSQLStart & sFilter
'            Debug.Print "Filter sql = " & sFilter
        End Select
        'Has the Filter band number changed from previous filter element
        If sBandNo <> sPrevBandNo Then
            'Check for Band Type
            If bANDBands Then
                'Its an AND-Band, is it the first band
                If sPrevBandNo <> "0" Then
                    If bFirstBand Then
                        'copy colSingleBandSiteIds to colAllBandsSiteIds
                        Set colAllBandsSiteIds = colSingleBandSiteIds
                        Set colSingleBandSiteIds = New Collection
                        bFirstBand = False
                    Else
                        Call JoinListsOR(colSingleBandSiteIds, colAllBandsSiteIds)
                        Set colSingleBandSiteIds = New Collection
                    End If
                End If
            Else
                'Its an OR-BAND, is it the first band
                If sPrevBandNo <> "0" Then
                    If bFirstBand Then
                        'copy colSingleBandSiteIds to colAllBandsSiteIds
                        Set colAllBandsSiteIds = colSingleBandSiteIds
                        Set colSingleBandSiteIds = New Collection
                        bFirstBand = False
                    Else
                        Call JoinListsAND(colSingleBandSiteIds, colAllBandsSiteIds)
                        Set colSingleBandSiteIds = New Collection
                    End If
                End If
            End If
            'Set Previous Band Number to current Band Number
            sPrevBandNo = sBandNo
            'Check for Band Type
            If bANDBands Then
                'Its first element of a New AND-BAND
                sSQL = sFilter
                'Convert the AND-BAND filter element into a collection of SitePersonIds
                Set colSingleBandSiteIds = FilterSqlToList(sSQL)
            Else
                'Its first element of a New OR-BAND
                sSQL = sFilter
                Set colSingleBandSiteIds = FilterSqlToList(sSQL)
            End If
        Else
            'Check for Band Type
            If bANDBands Then
                'Its an AND-BAND
                sSQL = sFilter
                'Convert the AND-BAND filter element into a collection of SitePersonIds
                Set colSingleElementSiteIds = FilterSqlToList(sSQL)
                'ADD colSingleElementSiteIds to colSingleBandSiteIds using AND LOGIC
                Call JoinListsAND(colSingleElementSiteIds, colSingleBandSiteIds)
            Else
                'Its an OR-BAND
                sSQL = sFilter
                'Convert the OR-BAND filter element into a collection of SitePersonIds
                Set colSingleElementSiteIds = FilterSqlToList(sSQL)
                'ADD colSingleElementSiteIds to colSingleBandSiteIds using OR LOGIC
                Call JoinListsOR(colSingleElementSiteIds, colSingleBandSiteIds)
            End If
        End If
    Next
    
    'act on the last filter element
    If bANDBands Then
        If bFirstBand Then
            'copy colSingleBandSiteIds to colAllBandsSiteIds
            Set colAllBandsSiteIds = colSingleBandSiteIds
            Set colSingleBandSiteIds = New Collection
            bFirstBand = False
        Else
            Call JoinListsOR(colSingleBandSiteIds, colAllBandsSiteIds)
            'Set colSingleBandSiteIds = New Collection
        End If
    Else
        If bFirstBand Then
            Set colAllBandsSiteIds = colSingleBandSiteIds
            Set colSingleBandSiteIds = New Collection
            bFirstBand = False
        Else
            Call JoinListsAND(colSingleBandSiteIds, colAllBandsSiteIds)
        End If
    End If
    
    'if colAllBandsSiteIds contains anything then loop through colAllBandsSiteIds creating an In String
    If colAllBandsSiteIds.Count > 0 Then
        'following SQl statement is database specific
        'Mo 25/1/2005 Bug 2510, Site/PersonId concatenation bug, "-" added.
        Select Case goUser.Database.DatabaseType
        Case MACRODatabaseType.Access
            sInString = " AND TrialSite & '-' & PersonId IN("
        Case MACRODatabaseType.Oracle80
            sInString = " AND TrialSite || '-' || PersonId IN("
        Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
            sInString = " AND TrialSite + '-' + CAST(PersonId as varchar(10)) IN("
        End Select
        For Each vSitePersonId In colAllBandsSiteIds
            sInString = sInString & "'" & vSitePersonId & "'" & ","
        Next
        sInString = Mid(sInString, 1, Len(sInString) - 1) & ")"
    Else
        'create an InString that contains no SitePersonIds
        'following SQl statement is database specific
        'Mo 25/1/2005 Bug 2510, Site/PersonId concatenation bug, "-" added.
        Select Case goUser.Database.DatabaseType
        Case MACRODatabaseType.Access
            sInString = " AND TrialSite & '-' & PersonId IN('zzz')"
        Case MACRODatabaseType.Oracle80
            sInString = " AND TrialSite || '-' || PersonId IN('zzz')"
        Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
            sInString = " AND TrialSite + '-' + CAST(PersonId as varchar(10)) IN('zzz')"
        End Select
    End If
    
    'Add the Filtered Subjects InString onto the supplied SQL statements
    sDataSQL = sDataSQL & sInString
    sSubjectRecordsSQL = sSubjectRecordsSQL & sInString

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "FilterThisSQL")
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
Public Sub ClearQueryAndReset()
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'Clear current grid and remove any splits
    Call ClearGrid
    
    'Clear the treeview down
    trwQuestions.Nodes.Clear
    
    'clear down the current query text areas/selections
    txtQueryText.Text = ""
    txtBandType.Text = ""
    DoEvents
    lstFilterText.Clear
    cboOperand.ListIndex = -1
    cboOperator.ListIndex = -1
    txtCriteria.Text = ""
    cboCatCodes.Clear
    txtBandNo.Text = ""
    txtResultsCount.Text = ""
    txtRecordsCount.Text = ""
    
    'initialise result attribute booleans
    mbComments = False
    mbCTCGrade = False
    mbLabResult = False
    mbStatus = False
    mbTimeStamp = False
    mbUserName = False
    mbValueCode = False
    
    'set initial Control states
    cmdNew.Enabled = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    cmdCancel.Enabled = False
    cmdRunQuery.Enabled = False
    mnuRRunQuery.Enabled = False
    cmdSaveOutPut.Enabled = False
    cmdCancelRun.Enabled = False
    mnuFSave.Enabled = False
    mnuFSaveAs.Enabled = False
    mnuFPrint.Enabled = False
    cboOperand.Enabled = False
    cboOperator.Enabled = False
    txtCriteria.Enabled = False
    txtCriteria.Visible = True
    cboCatCodes.Enabled = False
    cboCatCodes.Visible = False
    txtBandNo.Enabled = False
    optAndBands.Enabled = False
    optORBands.Enabled = False
    
    'initialise the QueryChanged flag
    gbQueryChanged = False
    
    'initialise the SelectedTrialId
    glSelectedTrialId = 0
    
    'Mo 2/4/2007 MRC15022007
    'Reset the study defaults
    gsSVMissing = ""
    gsSVUnobtainable = ""
    gsSVNotApplicable = ""
    gnOutPutType = 0
    gbOutputCategoryCodes = True
    gbDisplayStudyName = False
    gbDisplaySiteCode = True
    gbDisplayLabel = True
    gbDisplayPersonId = True
    gbDisplayVisitCycle = True
    gbDisplayFormCycle = True
    gbDisplayRepeatNumber = True
    gbSplitGrid = False
    gbUseShortCodes = True
    gbDisplayOutPut = False
    gbExcludeLabel = False
    gnShortCodeLength = 8
    gbSASInformatColons = False
    gsFileNamePath = ""
    gsFileNameText = ""
    gsFileNameStamp = "DATE"
    
    'clear any name that might have been placed in forms caption
    ' NCJ 7 Jan 04 - Changed to upper case MACRO
    Me.Caption = "MACRO Query Module"

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ClearQueryAndReset")
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
Private Sub SaveQuery(ByVal sQueryPathName As String)
'---------------------------------------------------------------------
'Changed Mo Morris 29/4/2002, SR4693 Regional settings no longer affects
'   the storage of True and False alongside query options.
'   (i.e if regional settinga are German it will not save '[Subject Label]Falsch'
'   but will save '[Subject Label]False', as it will for all regional settings)
'---------------------------------------------------------------------
Dim nIOFileNumber As Integer
Dim i As Integer
Dim sQueryName As String
Dim sOutPutType As String

    On Error GoTo Errhandler

    'open the output file
    nIOFileNumber = FreeFile
    Open sQueryPathName For Output As #nIOFileNumber
    
    'write the [SELECT] section label and Time Stamp to file
    Print #nIOFileNumber, "[SELECT] Saved " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    'write the contents of txtQueryText to file
    Print #nIOFileNumber, txtQueryText.Text
    
    'write the [FILTER] section label and Filter Band Type to file
    Print #nIOFileNumber, "[FILTER]" & txtBandType.Text
    
    'write the contents of lstFilterText to file
    For i = 0 To lstFilterText.ListCount - 1
        Print #nIOFileNumber, lstFilterText.List(i)
    Next
    
    'write the [OUTPUT] section label to file
    Print #nIOFileNumber, "[OUTPUT]"
    
    'Mo 1/11/2006 Bug 2795
    'write output type to file
    Select Case gnOutPutType
    Case eOutPutType.CSV
        sOutPutType = "CSV"
    Case eOutPutType.Access
        sOutPutType = "Access"
    Case eOutPutType.SPSS
        sOutPutType = "SPSS"
    Case eOutPutType.SAS
        sOutPutType = "SAS"
    Case eOutPutType.STATA
        sOutPutType = "STATA"
    Case eOutPutType.MACROBD
        sOutPutType = "MACROBD"
    Case eOutPutType.STATAStandardDates
        sOutPutType = "STATAStandardDates"
    Case eOutPutType.SASColons
        sOutPutType = "SASColons"
    End Select
    Print #nIOFileNumber, "[Format]" & sOutPutType
    
    'Mo 2/4/2007 MRC15022007
    'Write File Name Path to file
    Print #nIOFileNumber, "[File Name Path]" & gsFileNamePath
    
    'Write File Name Text to file
    Print #nIOFileNumber, "[File Name Text]" & gsFileNameText
    
    'Write File Name Stamp to file
    Print #nIOFileNumber, "[File Name Stamp]" & gsFileNameStamp
    
    'Write use of category codes to file
    If gbOutputCategoryCodes Then
        Print #nIOFileNumber, "[Use Category Codes]True"
    Else
        Print #nIOFileNumber, "[Use Category Codes]False"
    End If
    
    'Mo 2/6/2006 Bug 2737
    'Write use of  short Question codes
    If gbUseShortCodes Then
        Print #nIOFileNumber, "[Use Short Codes]" & CStr(gnShortCodeLength)
    Else
        Print #nIOFileNumber, "[Use Short Codes]False"
    End If

    'Write identification fields to be displayed to file
    If gbDisplayStudyName Then
        Print #nIOFileNumber, "[Display Study Name]True"
    Else
        Print #nIOFileNumber, "[Display Study Name]False"
    End If
    If gbDisplaySiteCode Then
        Print #nIOFileNumber, "[Display Site Code]True"
    Else
        Print #nIOFileNumber, "[Display Site Code]False"
    End If
    If gbDisplayLabel Then
        Print #nIOFileNumber, "[Display Label]True"
    Else
        Print #nIOFileNumber, "[Display Label]False"
    End If
    If gbDisplayPersonId Then
        Print #nIOFileNumber, "[Display PersonId]True"
    Else
        Print #nIOFileNumber, "[Display PersonId]False"
    End If
    If gbDisplayVisitCycle Then
        Print #nIOFileNumber, "[Display Visit Cycle]True"
    Else
        Print #nIOFileNumber, "[Display Visit Cycle]False"
    End If
    If gbDisplayFormCycle Then
        Print #nIOFileNumber, "[Display Form Cycle]True"
    Else
        Print #nIOFileNumber, "[Display Form Cycle]False"
    End If
    If gbDisplayRepeatNumber Then
        Print #nIOFileNumber, "[Display Repeat Number]True"
    Else
        Print #nIOFileNumber, "[Display Repeat Number]False"
    End If
    
    'Write use of split bar in grid to file
    If gbSplitGrid Then
        Print #nIOFileNumber, "[Grid Split bar]True"
    Else
        Print #nIOFileNumber, "[Grid Split bar]False"
    End If
    
    'Mo 30/5/2006 Bug 2668
    If gbExcludeLabel Then
        Print #nIOFileNumber, "[Exclude Subject Label]True"
    Else
        Print #nIOFileNumber, "[Exclude Subject Label]False"
    End If
    
    'Write Special Values to file
    Print #nIOFileNumber, "[SV Missing]" & gsSVMissing
    Print #nIOFileNumber, "[SV Unobtainable]" & gsSVUnobtainable
    Print #nIOFileNumber, "[SV NotApplicable]" & gsSVNotApplicable
    
    'close the output file
    Close #nIOFileNumber
    
    'extract Query name from QueryPathName
    sQueryName = StripFileNameFromPath(sQueryPathName)
    
    'place name of output file in forms caption
    Me.Caption = "MACRO Query Module : " & sQueryName
    
    'set the query has changed flag to false
    gbQueryChanged = False

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SaveQuery")
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
Private Sub SaveCheck()
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'Check for there being something to save
    'and prompt the user as to wether they want to save it
    If gbQueryChanged = True Then
        If DialogQuestion("Changes have been made to the current Query" _
        & vbNewLine & "Do you want to save the changes?", "Save Query Changes") <> vbYes Then
            'set mbquery to false so that this question is not asked again
            gbQueryChanged = False
            'User does not want to save changes
            Exit Sub
        End If
    End If

    'If the query has not been named then a call to Save As is required
    'But only do this if changes have been made to the unnamed query
    If ((gbQuerySaved = False) And (gbQueryChanged = True)) Then
        mnuFSaveAs_Click
        Exit Sub
    End If
    
    'if the query contains changes it needs to be saved
    If gbQueryChanged = True Then
        Call SaveQuery(msCurrentQueryPathName)
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SaveCheck")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'Mo 2/4/2007 MRC15022007, Changed from Sub to function
'---------------------------------------------------------------------
Public Function OpenQuery(ByVal sQueryPathName As String) As Boolean
'---------------------------------------------------------------------
'While reading through the query it will be validated.
'Each line must start with one of the following labels:-
'   [SELECT], Selection header record,                      mandatory
'   [S], Study plus Study selection,                        mandatory
'   [V], Visit selections,                          optional
'   [F], Form selections,                           optional
'   [Q], Question selections,                       optional
'   [FILTER], Filter header plus AND/OR Filter Band type,   mandatory
'   [1] to [10], Filter element with a filter band number
'   [OUTPUT], Output header plus Output details,            mandatory
'   [Format],CSV,Access,SPSS,SAS,STATA                      mandatory
'   [File Name Path] Blank for App.Path/Out Folder or user specified path
'                                                           mandatory
'   [File Name Text] Blank for Study Name or user specified text
'                                                           mandatory
'   [File Name Stamp]   "DATE" for a yyyymmdd stamp,
'                       "DATETIME" for a yyyymmddhhmmss stamp
'                       or "" for no stamp                  mandatory
'   [Use Category Codes] True or False                      mandatory
'   [Use Short Codes] False or Number or (True)             mandatory
'   [Display Study Name] True or False                      mandatory
'   [Display Site Code] True or False                       mandatory
'   [Display Label] True or False                           mandatory
'   [Display PersonId] True or False                        mandatory
'   [Display Visit Cycle] True or False                     mandatory
'   [Display Form Cycle] True or False                      mandatory
'   [Display Repeat Number] True or False                   mandatory
'   [Grid Split bar] True or False                          mandatory
'   [Exclude Subject Label] True or False                   mandatory
'   [SV Missing] Special Value for status Missing           mandatory
'   [SV Unobtainable] Special Value for status Unobtainable mandatory
'   [SV NotApplicable] Special Value for status NotApplicable mandatory
'
'Opening the query will be aborted if:-
'   Not all of the mandatory lines occur.
'   Invalid labels occur.
'   The query contains a Study name that does not exist in the current DB.
'   The query contains Visit, Form or Question codes that do not exist in the specified study.
'   Mo 2/6/2006 Bug 2737 - Add Question Short Code length to the Options Window
'   The [Use Short Codes] label should now be followed by an integer number that represents the Short Code length.
'   From now on this label will be set to False or a Number. For historical reasons True will still be read as a
'   valid entry, but it will be assessed as a a code length of 8 (the default value).
'   Mo  2/4/2007    MRC15022007 - Query Module Batch Facilities
'   3 new labels  [File Name Path], [File Name Text] and [File Name Stamp] added
'---------------------------------------------------------------------
Dim nIOFileNumber As Integer
Dim sQueryName As String
Dim sQueryLine As String
Dim sLabel As String
Dim sStudy As String
Dim lClinicalTrialId As Long
Dim lVisitId As Long
Dim lFormId As Long
Dim lQuestionId As Long
Dim sStudyQuestion As String
Dim sVisit As String
Dim sVisitQuestion As String
Dim sForm As String
Dim sFormQuestion As String
Dim sQuestion As String
Dim sQuestionAttribute As String
Dim sFilterQuestion As String
Dim sFilterANDOR As String
Dim nMandatoryLinesCount As Integer
Dim bContainsInvalidLabels As Boolean
Dim bStudyDoesNotExist As Boolean
Dim bStudyFoundInCombo As Boolean
Dim bVisitDoesNotExist As Boolean
Dim bFormDoesNotExist As Boolean
Dim bQuestionDoesNotExist As Boolean
Dim bTreeNodeDoesNotExist As Boolean
Dim bContainsInvalidOptions As Boolean
Dim sNodeKey As String
Dim i As Integer
Dim sInvalidMessage As String
'Mo 2/4/2007 MRC15022007
Dim bValid As Boolean

    On Error GoTo Errhandler
    
    Call HourglassOn
    
    'open the input file
    nIOFileNumber = FreeFile
    Open sQueryPathName For Input As #nIOFileNumber
    
    nMandatoryLinesCount = 0
    bContainsInvalidLabels = False
    bStudyDoesNotExist = False
    bVisitDoesNotExist = False
    bFormDoesNotExist = False
    bQuestionDoesNotExist = False
    bStudyFoundInCombo = False
    bTreeNodeDoesNotExist = False
    bContainsInvalidOptions = False
    'Read the query line by line
    Do While Not EOF(nIOFileNumber)
        Line Input #nIOFileNumber, sQueryLine
        'process non blank lines
        If Trim(sQueryLine) <> "" Then
            'Strip label from line
            sLabel = Mid(sQueryLine, 1, InStr(sQueryLine, "]"))
            Select Case sLabel
            Case "[SELECT]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
            Case "[S]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                'Distinquish between 'Study' and 'Study.All Questions'
                If InStr(sQueryLine, ".") > 0 Then
                    sStudy = Mid(sQueryLine, 4, InStr(sQueryLine, ".") - 4)
                    sStudyQuestion = Mid(sQueryLine, InStr(sQueryLine, ".") + 1)
                Else
                    sStudy = Mid(sQueryLine, 4)
                    sStudyQuestion = ""
                End If
                'Check study exists, note the return of the found ClinicalTrialId
                If Not StudyExists(sStudy, lClinicalTrialId) Then
                    bStudyDoesNotExist = True
                    Exit Do
                End If
                'Select study in cboStudies, which will have the effect of
                'triggering the building of the treeview
                For i = 0 To cboStudies.ListCount - 1
                    If cboStudies.List(i) = sStudy Then
                        cboStudies.ListIndex = i
                        bStudyFoundInCombo = True
                        Exit For
                    End If
                Next
                'Check that study was found in cboStudies
                If Not bStudyFoundInCombo Then
                    bStudyDoesNotExist = True
                    Exit Do
                End If
                'If its 'Study.All Questions' click the relevant node
                If sStudyQuestion = "All Questions" Then
                    sNodeKey = "S"
                    If Not NodeClickedInTreeView(sNodeKey) Then
                        bTreeNodeDoesNotExist = True
                        Exit Do
                    End If
                End If
            Case "[V]"
                sVisit = Mid(sQueryLine, 4, InStr(sQueryLine, ".") - 4)
                sVisitQuestion = Mid(sQueryLine, InStr(sQueryLine, ".") + 1)
                'Check Visit exists in specified study
                If Not VisitExists(sVisit, lClinicalTrialId, lVisitId) Then
                    bVisitDoesNotExist = True
                    Exit Do
                End If
                'Distinquish between "All Questions" and specific question
                If sVisitQuestion <> "All Questions" Then
                    'Its a specific question, Check it exists in specified study
                    If Not QuestionExists(sVisitQuestion, lClinicalTrialId, lQuestionId) Then
                        bQuestionDoesNotExist = True
                        Exit Do
                    End If
                    sNodeKey = "S|V|" & lVisitId & "|" & lQuestionId
                Else
                    sNodeKey = "S|V|" & lVisitId
                End If
                'Select node in treeview, which in turn calls RefreshQueryText
                If Not NodeClickedInTreeView(sNodeKey) Then
                    bTreeNodeDoesNotExist = True
                    Exit Do
                End If
            Case "[F]"
                sForm = Mid(sQueryLine, 4, InStr(sQueryLine, ".") - 4)
                sFormQuestion = Mid(sQueryLine, InStr(sQueryLine, ".") + 1)
                'Check Form exists in specified study
                If Not FormExists(sForm, lClinicalTrialId, lFormId) Then
                    bFormDoesNotExist = True
                    Exit Do
                End If
                'Distinquish between "All Questions" and specific question
                If sFormQuestion <> "All Questions" Then
                    'Its a specific question, Check it exists in specified study
                    If Not QuestionExists(sFormQuestion, lClinicalTrialId, lQuestionId) Then
                        bQuestionDoesNotExist = True
                        Exit Do
                    End If
                    sNodeKey = "S|F|" & lFormId & "|" & lQuestionId
                Else
                    sNodeKey = "S|F|" & lFormId
                End If
                'Select node in treeview, which in turn calls RefreshQueryText
                If Not NodeClickedInTreeView(sNodeKey) Then
                    bTreeNodeDoesNotExist = True
                    Exit Do
                End If
            Case "[Q]"
                sQuestion = Mid(sQueryLine, 4, InStr(sQueryLine, ".") - 4)
                sQuestionAttribute = Mid(sQueryLine, InStr(sQueryLine, ".") + 1)
                'Check Question exists in specified study
                If Not QuestionExists(sQuestion, lClinicalTrialId, lQuestionId) Then
                    bQuestionDoesNotExist = True
                    Exit Do
                End If
                'Distinquish between "ResponseValue" and the other question attributes
                Select Case sQuestionAttribute
                Case "ResponseValue"
                    sNodeKey = "S|Q|" & lQuestionId
                Case "Comments"
                    sNodeKey = "S|Q|" & lQuestionId & "|Comments"
                Case "CTCGrade"
                    sNodeKey = "S|Q|" & lQuestionId & "|CTCGrade"
                Case "LabResult"
                    sNodeKey = "S|Q|" & lQuestionId & "|LabResult"
                Case "Status"
                    sNodeKey = "S|Q|" & lQuestionId & "|Status"
                Case "TimeStamp"
                    sNodeKey = "S|Q|" & lQuestionId & "|TimeStamp"
                Case "UserName"
                    sNodeKey = "S|Q|" & lQuestionId & "|UserName"
                Case "ValueCode"
                    sNodeKey = "S|Q|" & lQuestionId & "|ValueCode"
                End Select
                'Select node in treeview, which in turn calls RefreshQueryText
                If Not NodeClickedInTreeView(sNodeKey) Then
                    bTreeNodeDoesNotExist = True
                    Exit Do
                End If
            Case "[FILTER]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                sQueryLine = Mid(sQueryLine, InStr(sQueryLine, "]") + 1)
                sFilterANDOR = Mid(sQueryLine, 2, InStr(sQueryLine, "]") - 2)
                'Set relevant AND/OR Filter Band type option button
                If sFilterANDOR = "AND" Then
                    'set to false first so that the setting to true definately triggers optANDBands_Click,
                    'which in turn places the AND/OR filter option text in txtBandType.text
                    optAndBands.Value = False
                    optAndBands.Value = True
                Else
                    optORBands.Value = False
                    optORBands.Value = True
                End If
            Case "[1]", "[2]", "[3]", "[4]", "[5]", "[6]", "[7]", "[8]", "[9]"
                sFilterQuestion = Mid(sQueryLine, 4, InStr(sQueryLine, " ") - 4)
                Select Case sFilterQuestion
                Case "SubjectId"
                    lQuestionId = -1
                Case "Site"
                    lQuestionId = -2
                Case "SubjectLabel"
                    lQuestionId = -3
                Case Else
                    'Check Filter question exists in specified study
                    If Not QuestionExists(sFilterQuestion, lClinicalTrialId, lQuestionId) Then
                        bQuestionDoesNotExist = True
                        Exit Do
                    End If
                End Select
                'Add filter element to lstFilterText
                lstFilterText.AddItem sQueryLine
                'Store the questions DatatItemID as ItemData
                lstFilterText.ItemData(lstFilterText.NewIndex) = lQuestionId
            Case "[OUTPUT]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
            Case "[Format]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                gbSASInformatColons = False
                'extract output format and set gnOutPutType
                Select Case Mid(sQueryLine, 9)
                Case "CSV"
                    gnOutPutType = eOutPutType.CSV
                Case "Access"
                    gnOutPutType = eOutPutType.Access
                Case "SPSS"
                    gnOutPutType = eOutPutType.SPSS
                Case "SAS"
                    gnOutPutType = eOutPutType.SAS
                Case "STATA"
                    gnOutPutType = eOutPutType.STATA
                Case "MACROBD"
                    gnOutPutType = eOutPutType.MACROBD
                Case "STATAStandardDates"
                    gnOutPutType = eOutPutType.STATAStandardDates
                Case "SASColons"
                    gnOutPutType = eOutPutType.SASColons
                    gbSASInformatColons = True
                Case Else
                    bContainsInvalidOptions = True
                End Select
            'Mo 2/4/2007 MRC15022007
            Case "[File Name Path]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                If Trim(Mid(sQueryLine, 17)) = "" Then
                    gsFileNamePath = ""
                Else
                    gsFileNamePath = Mid(sQueryLine, 17)
                End If
            Case "[File Name Text]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                If Trim(Mid(sQueryLine, 17)) = "" Then
                    gsFileNameText = ""
                Else
                    gsFileNameText = Mid(sQueryLine, 17)
                End If
            Case "[File Name Stamp]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                gsFileNameStamp = Mid(sQueryLine, 18)
            Case "[Use Category Codes]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                'extract setting for gbOutputCategoryCodes
                Select Case Mid(sQueryLine, 21)
                Case "True"
                    gbOutputCategoryCodes = True
                Case "False"
                    gbOutputCategoryCodes = False
                Case Else
                    bContainsInvalidOptions = True
                End Select
            Case "[Use Short Codes]"
                'Mo 2/6/2006 Bug 2737
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                'extract setting for gbUseShortCodes & gnShortCodeLength
                Select Case Mid(sQueryLine, 18)
                Case "True"
                    gbUseShortCodes = True
                    gnShortCodeLength = 8
                Case "False"
                    gbUseShortCodes = False
                    gnShortCodeLength = 8
                Case "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18"
                    gbUseShortCodes = True
                    gnShortCodeLength = CInt(Mid(sQueryLine, 18))
                Case Else
                    bContainsInvalidOptions = True
                End Select
            Case "[Display Study Name]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                'extract setting for gbDisplayStudyName
                Select Case Mid(sQueryLine, 21)
                Case "True"
                    gbDisplayStudyName = True
                Case "False"
                    gbDisplayStudyName = False
                Case Else
                    bContainsInvalidOptions = True
                End Select
            Case "[Display Site Code]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                'extract setting for gbDisplaySiteCode
                Select Case Mid(sQueryLine, 20)
                Case "True"
                    gbDisplaySiteCode = True
                Case "False"
                    gbDisplaySiteCode = False
                Case Else
                    bContainsInvalidOptions = True
                End Select
            Case "[Display Label]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                'extract setting for gbDisplayLabel
                Select Case Mid(sQueryLine, 16)
                Case "True"
                    gbDisplayLabel = True
                Case "False"
                    gbDisplayLabel = False
                Case Else
                    bContainsInvalidOptions = True
                End Select
            Case "[Display PersonId]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                'extract setting for gbDisplayPersonId
                Select Case Mid(sQueryLine, 19)
                Case "True"
                    gbDisplayPersonId = True
                Case "False"
                    gbDisplayPersonId = False
                Case Else
                    bContainsInvalidOptions = True
                End Select
            Case "[Display Visit Cycle]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                'extract setting for gbDisplayVisitCycle
                Select Case Mid(sQueryLine, 22)
                Case "True"
                    gbDisplayVisitCycle = True
                Case "False"
                    gbDisplayVisitCycle = False
                Case Else
                    bContainsInvalidOptions = True
                End Select
            Case "[Display Form Cycle]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                'extract setting for gbDisplayFormCycle
                Select Case Mid(sQueryLine, 21)
                Case "True"
                    gbDisplayFormCycle = True
                Case "False"
                    gbDisplayFormCycle = False
                Case Else
                    bContainsInvalidOptions = True
                End Select
            Case "[Display Repeat Number]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                'extract setting for gbDisplayRepeatNumber
                Select Case Mid(sQueryLine, 24)
                Case "True"
                    gbDisplayRepeatNumber = True
                Case "False"
                    gbDisplayRepeatNumber = False
                Case Else
                    bContainsInvalidOptions = True
                End Select
            Case "[Grid Split bar]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                'extract setting for gbSplitGrid
                Select Case Mid(sQueryLine, 17)
                Case "True"
                    gbSplitGrid = True
                Case "False"
                    gbSplitGrid = False
                Case Else
                    bContainsInvalidOptions = True
                End Select
            'Mo 30/5/2006 Bug 2668
            Case "[Exclude Subject Label]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                'extract setting for gbExcludeLabel
                Select Case Mid(sQueryLine, 24)
                Case "True"
                    gbExcludeLabel = True
                Case "False"
                    gbExcludeLabel = False
                Case Else
                    bContainsInvalidOptions = True
                End Select
            Case "[SV Missing]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                gsSVMissing = Mid(sQueryLine, 13)
            Case "[SV Unobtainable]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                gsSVUnobtainable = Mid(sQueryLine, 18)
            Case "[SV NotApplicable]"
                nMandatoryLinesCount = nMandatoryLinesCount + 1
                gsSVNotApplicable = Mid(sQueryLine, 19)
            Case Else
                'set the invalid label flag
                bContainsInvalidLabels = True
            End Select
        End If
    Loop
    
    'close the input file
    Close #nIOFileNumber
    
    
    'Mo 30/5/2006 Bug 2668, if nMandatoryLinesCount = 18 then it is because it is
    'a query that has been saved before [Exclude Subject Label] was added.
    If nMandatoryLinesCount = 18 Then
        'increment nMandatoryLinesCount to 19
        nMandatoryLinesCount = 19
        'set gbExcludeLabel to the default value of false
        gbExcludeLabel = False
    End If
    'Mo 2/4/2007 MRC15022007
    If nMandatoryLinesCount = 19 Then
        'increment nMandatoryLinesCount to 22
        nMandatoryLinesCount = 22
        'set gsFileNamePath, gsFileNameText & gsFileNameStamp to the default value of false
        gsFileNamePath = ""
        gsFileNameText = ""
        gsFileNameStamp = "DATE"
    End If
    'Check for any query validation errors
    'Mo 2/4/2007 MRC15022007
    If nMandatoryLinesCount <> 22 Or _
    bContainsInvalidLabels Or bStudyDoesNotExist Or bVisitDoesNotExist Or _
    bFormDoesNotExist Or bQuestionDoesNotExist Or bTreeNodeDoesNotExist Or bContainsInvalidOptions Then
        'Check the individual errors and build up an error message
        sInvalidMessage = "Opening of Query aborted." & vbNewLine
        If bStudyDoesNotExist Then
            sInvalidMessage = sInvalidMessage & vbNewLine & "The query was written for a study that does not exist in the current database."
        End If
        If bVisitDoesNotExist Or bFormDoesNotExist Or bQuestionDoesNotExist Then
            sInvalidMessage = sInvalidMessage & vbNewLine & "The query contains Visit, Form or Question codes that no longer exist in the specified study."
        End If
        If bTreeNodeDoesNotExist Then
            sInvalidMessage = sInvalidMessage & vbNewLine & "The query contains visit/question or form/question " _
            & vbNewLine & "combinations that are no longer valid in the specified study."
        End If
        If bContainsInvalidLabels Then
            sInvalidMessage = sInvalidMessage & vbNewLine & "The query contains invalid [LABELS]."
        End If
        If bContainsInvalidOptions Then
            sInvalidMessage = sInvalidMessage & vbNewLine & "The query contains invalid output options."
        End If
        'The nMandatoryLinesCount check tends to be invalid if other errors that cause
        'an 'Exit DO' are reached. Only display this error if the others have not occured
        If nMandatoryLinesCount <> 4 And Not (bStudyDoesNotExist Or bVisitDoesNotExist Or bFormDoesNotExist Or bQuestionDoesNotExist Or bTreeNodeDoesNotExist Or bContainsInvalidOptions) Then
            sInvalidMessage = sInvalidMessage & vbNewLine & "The query does not contain all of the mandatory labels [SELECT],[S],[FILTER] etc."
        End If
        
        'Display the compond message
        Call DialogInformation(sInvalidMessage, "MACRO Query error: " & sQueryPathName)
        'clear down the newly opened query
        Call ClearQueryAndReset
        'Unselect the previously selected study
        cboStudies.ListIndex = -1
        bValid = False
    Else
        'Mo 2/4/2007 MRC15022007
        bValid = True
    
        'extract Query name from QueryPathName
        sQueryName = StripFileNameFromPath(sQueryPathName)
        
        'place name of newly open query in forms caption
        Me.Caption = "MACRO Query Module : " & sQueryName
        
        'set the query has changed flag to false
        gbQueryChanged = False
    End If
    
    Call HourglassOff
    
    'Mo 2/4/2007 MRC15022007
    If bValid Then
        OpenQuery = True
    Else
        OpenQuery = False
    End If
    
Exit Function
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "OpenQuery")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'---------------------------------------------------------------------
Private Function StudyExists(ByVal sClinicalTrialName As String, _
                            ByRef lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo Errhandler

    sSQL = "SELECT ClinicalTrialID " _
        & "FROM ClinicalTrial " _
        & "WHERE ClinicalTrialName = '" & sClinicalTrialName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        StudyExists = False
    Else
        lClinicalTrialId = rsTemp!ClinicalTrialId
        StudyExists = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
        
Exit Function
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "StudyExists")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'---------------------------------------------------------------------
Private Function VisitExists(ByVal sVisitCode As String, _
                            ByVal lClinicalTrialId As Long, _
                            ByRef lVisitIdId As Long) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo Errhandler

    sSQL = "SELECT VisitID " _
        & "FROM StudyVisit " _
        & "WHERE ClinicalTrialId = " & lClinicalTrialId & " " _
        & "AND VisitCode = '" & sVisitCode & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        VisitExists = False
    Else
        lVisitIdId = rsTemp!VisitId
        VisitExists = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
        
Exit Function
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "VisitExists")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'---------------------------------------------------------------------
Private Function FormExists(ByVal sFormCode As String, _
                            ByVal lClinicalTrialId As Long, _
                            ByRef lFormId As Long) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo Errhandler

    sSQL = "SELECT CRFPageId " _
        & "FROM CRFPage " _
        & "WHERE ClinicalTrialId = " & lClinicalTrialId & " " _
        & "AND CRFPageCode = '" & sFormCode & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        FormExists = False
    Else
        lFormId = rsTemp!CRFPageId
        FormExists = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
        
Exit Function
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "FormExists")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'---------------------------------------------------------------------
Private Function QuestionExists(ByVal sQuestionCode As String, _
                            ByVal lClinicalTrialId As Long, _
                            ByRef lQuestionId As Long) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo Errhandler

    sSQL = "SELECT DataItemID " _
        & "FROM DataItem " _
        & "WHERE ClinicalTrialId = " & lClinicalTrialId & " " _
        & "AND DataItemCode = '" & sQuestionCode & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        QuestionExists = False
    Else
        lQuestionId = rsTemp!DataItemId
        QuestionExists = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
        
Exit Function
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "QuestionExists")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'---------------------------------------------------------------------
Private Function NodeClickedInTreeView(ByVal sNodeKey As String) As Boolean
'---------------------------------------------------------------------
Dim oNode As MSComctlLib.Node

    On Error GoTo NodeNotFound

    Set oNode = trwQuestions.Nodes(sNodeKey)
    Call trwQuestions_NodeClick(oNode)
    
    NodeClickedInTreeView = True
    
Exit Function

NodeNotFound:
    NodeClickedInTreeView = False

End Function

'---------------------------------------------------------------------
Private Function FilterSqlToList(ByVal sSQL As String) As Collection
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim colSitePersonIds As Collection

    On Error GoTo Errhandler

    Set colSitePersonIds = New Collection

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Do While Not rsTemp.EOF
        'Mo 25/1/2005 Bug 2510, Site/PersonId concatenation bug, "-" added.
        colSitePersonIds.Add rsTemp!TrialSite & "-" & rsTemp!PersonID, rsTemp!TrialSite & "-" & rsTemp!PersonID
        rsTemp.MoveNext
    Loop
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    Set FilterSqlToList = colSitePersonIds
    
    Set colSitePersonIds = Nothing

Exit Function
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "FilterSqlToList")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'---------------------------------------------------------------------
Private Sub JoinListsAND(ByRef colToBeAdded As Collection, _
                            ByRef colMaster As Collection)
'---------------------------------------------------------------------
'This sub joins 2 collections of SitePersonIds using AND logic
'(i.e. a SitePersonId has to occur in both lists to get into the returned list)
'---------------------------------------------------------------------
Dim vSitePersonId As Variant
Dim colNewMaster As Collection
Dim svar As String
    
    'initialise the new master collection of SitePersonIds
    Set colNewMaster = New Collection
    
    'Compare contents of colToBeAdded with ColMaster
    'If an entry from colToBeAdded exists in ColMaster then add it to colNewMaster
    'If an entry from colToBeAdded does not exist in ColMaster an error will be generated and it is not added to colNewMaster+-
    On Error Resume Next
    For Each vSitePersonId In colToBeAdded
        svar = colMaster.Item(vSitePersonId)
        If Err.Number = 0 Then
            colNewMaster.Add vSitePersonId, vSitePersonId
        Else
            'clear error
            Err.Clear
        End If
    Next
    
    'overwrite the old Master collection with the new Master collection
    Set colMaster = colNewMaster
    
    'clear down colNewMaster
    Set colNewMaster = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub JoinListsOR(ByRef colToBeAdded As Collection, _
                            ByRef colMaster As Collection)
'---------------------------------------------------------------------
'This sub joins 2 collections of SitePersonIds using OR logic
'(i.e. a SitePersonId only has to occur in one of the lists to get into the returned list)
'---------------------------------------------------------------------
Dim vSitePersonId As Variant
Dim colNewMaster As Collection
Dim svar As String
    
    'initialise the new master collection of SitePersonIds
    Set colNewMaster = New Collection
    'copy contents of current colMaster into colNewMaster
    Set colNewMaster = colMaster
    
    'Read contents of colToBeAdded
    'If an entry from colToBeAdded already exists in ColMaster then it does not need to be added to colNewMaster
    'If an entry from colToBeAdded does not exist in ColMaster an error will be generated and it should be added to colNewMaster
    On Error Resume Next
    For Each vSitePersonId In colToBeAdded
        svar = colMaster.Item(vSitePersonId)
        If Err.Number <> 0 Then
            colNewMaster.Add vSitePersonId, vSitePersonId
            'clear error
            Err.Clear
        End If
    Next
    
    'overwrite the old Master collection with the new Master collection
    Set colMaster = colNewMaster
    
    'clear down colNewMaster
    Set colNewMaster = Nothing
    
End Sub

'---------------------------------------------------------------------
Private Sub PopulatecboCatCodes(lDataItemId As Long)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo Errhandler
    
    cboCatCodes.Clear
    sSQL = "SELECT ValueCode, ItemValue FROM ValueData " _
        & "WHERE ClinicalTrialId = " & glSelectedTrialId & " " _
        & "AND DataItemId = " & cboOperand.ItemData(cboOperand.ListIndex) & " " _
        & "ORDER BY ValueOrder"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do Until rsTemp.EOF
        cboCatCodes.AddItem rsTemp!ValueCode & " - " & rsTemp!ItemValue
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PopulatecboCatCodes")
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
Private Sub ClearGrid()
'---------------------------------------------------------------------
Dim i As Integer
Dim rsTemp As ADODB.Recordset

    On Error GoTo Errhandler

    'give the grid an empty datasource
    Set rsTemp = New ADODB.Recordset
    Set grdOutPut.Datasource = rsTemp
    Set rsTemp = Nothing
    
    'remove any splits that might exist
    For i = 1 To grdOutPut.Splits.Count - 1
        grdOutPut.Splits.Remove i
    Next i

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ClearGrid")
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
Private Sub GridSplitAdd()
'---------------------------------------------------------------------
Dim nSplitWidth As Integer
Dim i As Integer
Dim nGridWidth As Integer

    On Error GoTo Errhandler

    'The width of the split will be based on the total width of the columns
    'used by the identification fields that have been requested
    nSplitWidth = 0
    If gbDisplayStudyName Then
        nSplitWidth = nSplitWidth + grdOutPut.Columns(0).Width
    Else
        grdOutPut.Splits(0).Columns(0).Visible = False
    End If
    
    If gbDisplaySiteCode Then
        nSplitWidth = nSplitWidth + grdOutPut.Columns(1).Width
    Else
        grdOutPut.Splits(0).Columns(1).Visible = False
    End If

    If gbDisplayLabel Then
        nSplitWidth = nSplitWidth + grdOutPut.Columns(2).Width
    Else
        grdOutPut.Splits(0).Columns(2).Visible = False
    End If
    
    If gbDisplayPersonId Then
        nSplitWidth = nSplitWidth + grdOutPut.Columns(3).Width
    Else
        grdOutPut.Splits(0).Columns(3).Visible = False
    End If
    
    If gbDisplayVisitCycle Then
        nSplitWidth = nSplitWidth + grdOutPut.Columns(4).Width
    Else
        grdOutPut.Splits(0).Columns(4).Visible = False
    End If
    
    If gbDisplayFormCycle Then
        nSplitWidth = nSplitWidth + grdOutPut.Columns(5).Width
    Else
        grdOutPut.Splits(0).Columns(5).Visible = False
    End If
    
    If gbDisplayRepeatNumber Then
        nSplitWidth = nSplitWidth + grdOutPut.Columns(6).Width
    Else
        grdOutPut.Splits(0).Columns(6).Visible = False
    End If
    
    'hide the question response fields that will be to the left of the split
    For i = 7 To grdOutPut.Splits(0).Columns.Count - 1
        grdOutPut.Splits(0).Columns(i).Visible = False
    Next i
    
    'retrieve width of grid
    nGridWidth = grdOutPut.Width
    
    'create the split
    grdOutPut.Splits.Add (1)
    'as per Q306886
    grdOutPut.Splits(0).ScrollBars = dbgBoth
    grdOutPut.Splits(0).SizeMode = dbgExact
    '550 is a fudge factor for width of select column
    grdOutPut.Splits(0).Size = nSplitWidth + 550
    'prevent the split from being resized
    grdOutPut.Splits(0).AllowSizing = False
    
    'set the width to the right of split, with 620 being a fudge factor for width of select column
    grdOutPut.Splits(1).SizeMode = dbgExact
    grdOutPut.Splits(1).Size = nGridWidth - nSplitWidth - 620
    
    'hide the identification fields to the right of the split
    For i = 0 To 6
        grdOutPut.Splits(1).Columns(i).Visible = False
    Next i
    
    'display the question fields to the right of the split
    For i = 7 To grdOutPut.Splits(1).Columns.Count - 1
        grdOutPut.Splits(1).Columns(i).Visible = True
    Next i
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GridSplitAdd")
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
Private Function FormatOutPut(ByVal lDataItemId As Long, _
                        ByVal vResponseValue As Variant, _
                        ByVal vValueCode As Variant, _
                        ByVal nStatus As Integer) As Variant
'---------------------------------------------------------------------
' NCJ 19 Dec 05 - Added Partial Date flag
'---------------------------------------------------------------------
'Dim vResponse As Variant
Dim sDataItemFormat As String
Dim nPDFlag As Integer

    On Error GoTo Errhandler
    
    'Responses should either exist or be null. An empty string is an invalid response that is set to null.
    If Not IsNull(vResponseValue) Then
        If vResponseValue = "" Then
            vResponseValue = Null
        End If
    End If

    If IsNull(vResponseValue) Then
        Select Case nStatus
        Case Status.Missing
            If gsSVMissing = "" Then
                FormatOutPut = Null
            Else
                Select Case DataTypeFromId(glSelectedTrialId, lDataItemId)
                'If is a date the negative special value has to be converted into a string
                Case DataType.Date
                    'Mo 18/10/2006 Bug 2822, check the Partial Dates Flag when deciding Date or String
                    sDataItemFormat = DataItemFormatFromId(glSelectedTrialId, lDataItemId, nPDFlag)
                    If nPDFlag = 0 And DateFormatCanBeConverted(sDataItemFormat) Then
                        FormatOutPut = Format(CDate(gsSVMissing), "yyyy/mm/dd")
                    Else
                        FormatOutPut = gsSVMissing
                    End If
                Case Else
                    FormatOutPut = gsSVMissing
                End Select
            End If
        Case Status.NotApplicable
            If gsSVNotApplicable = "" Then
                FormatOutPut = Null
            Else
                Select Case DataTypeFromId(glSelectedTrialId, lDataItemId)
                'If is a date the negative special value has to be converted into a string
                Case DataType.Date
                    'Mo 18/10/2006 Bug 2822, check the Partial Dates Flag when deciding Date or String
                    sDataItemFormat = DataItemFormatFromId(glSelectedTrialId, lDataItemId, nPDFlag)
                    If nPDFlag = 0 And DateFormatCanBeConverted(sDataItemFormat) Then
                        FormatOutPut = Format(CDate(gsSVNotApplicable), "yyyy/mm/dd")
                    Else
                        FormatOutPut = gsSVNotApplicable
                    End If
                Case Else
                    FormatOutPut = gsSVNotApplicable
                End Select
            End If
        Case Status.Unobtainable
            If gsSVUnobtainable = "" Then
                FormatOutPut = Null
            Else
                Select Case DataTypeFromId(glSelectedTrialId, lDataItemId)
                'If is a date the negative special value has to be converted into a string
                Case DataType.Date
                    'Mo 18/10/2006 Bug 2822, check the Partial Dates Flag when deciding Date or String
                    sDataItemFormat = DataItemFormatFromId(glSelectedTrialId, lDataItemId, nPDFlag)
                    If nPDFlag = 0 And DateFormatCanBeConverted(sDataItemFormat) Then
                        FormatOutPut = Format(CDate(gsSVUnobtainable), "yyyy/mm/dd")
                    Else
                        FormatOutPut = gsSVUnobtainable
                    End If
                Case Else
                    FormatOutPut = gsSVUnobtainable
                End Select
            End If
        Case Else
            FormatOutPut = Null
        End Select
    Else
        Select Case DataTypeFromId(glSelectedTrialId, lDataItemId)
        'Mo 25/10/2005 COD0040
        Case DataType.Text, DataType.Multimedia, DataType.Thesaurus
            FormatOutPut = vResponseValue
        Case DataType.IntegerData
            'Changed Mo 27/6/2002, CBB 2.2.18.4, CInt replacing CLng to stop overflow error on long integer data
            'Changed Mo 3/12/2002, SR4982/SR5111 Regional settings
            FormatOutPut = CLng(ConvertStandardToLocalNum(CStr(vResponseValue)))
        Case DataType.Real, DataType.LabTest
            'Changed Mo 3/12/2002, SR4982/SR5111 Regional settings
            'Mo 31/1/2007 Bug 2873, Cast Real and LabTest responses using CDec(), used to be CSing()
            FormatOutPut = CDec(ConvertStandardToLocalNum(CStr(vResponseValue)))
        Case DataType.Category
            'check for codes or values being required
            If gbOutputCategoryCodes Then
                FormatOutPut = vValueCode
                'vResponse = vValueCode
            Else
                FormatOutPut = vResponseValue
                'vResponse = vResponseValue
            End If
        Case DataType.Date
            ' NCJ 19 Dec 05 - Check partial date flag (can't convert partial dates)
            sDataItemFormat = DataItemFormatFromId(glSelectedTrialId, lDataItemId, nPDFlag)
            If nPDFlag = 0 And DateFormatCanBeConverted(sDataItemFormat) Then
                FormatOutPut = yyyymmddhhmmssDateFormat(sDataItemFormat, vResponseValue)
            Else
                FormatOutPut = vResponseValue
            End If
        Case Else
            FormatOutPut = vResponseValue
        End Select
    End If
    
Exit Function
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "FormatOutPut")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'---------------------------------------------------------------------
Private Function QuestionExistsInStudy(ByVal sVFQKey As String, _
                                        ByRef sVisitFormQuestion As String) As Boolean
'---------------------------------------------------------------------

    On Error Resume Next
    
    sVisitFormQuestion = mcolVFQLookUp(sVFQKey)
    
    If Err.Number <> 0 Then
        Err.Clear
        QuestionExistsInStudy = False
    Else
        QuestionExistsInStudy = True
    End If
    
End Function

'---------------------------------------------------------------------
Public Sub PopulatecboCatCodesWithSites()
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo Errhandler
    
    cboCatCodes.Clear
    sSQL = "SELECT TrialSite FROM TrialSite " _
        & "WHERE ClinicalTrialId = " & glSelectedTrialId & " " _
        & "ORDER BY TrialSite"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do Until rsTemp.EOF
        cboCatCodes.AddItem rsTemp!TrialSite
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PopulatecboCatCodesWithSites")
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
Public Sub PopulatecboCatCodesWithPersonIds()
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo Errhandler
    
    cboCatCodes.Clear
    sSQL = "SELECT DISTINCT PersonId FROM TrialSubject " _
        & "WHERE ClinicalTrialId = " & glSelectedTrialId & " " _
        & "ORDER BY PersonId"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do Until rsTemp.EOF
        cboCatCodes.AddItem rsTemp!PersonID
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PopulatecboCatCodesWithPersonIds")
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
Public Sub PopulatecboCatCodesWithSubjectLabels()
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo Errhandler
    
    cboCatCodes.Clear
    sSQL = "SELECT DISTINCT LocalIdentifier1 FROM TrialSubject " _
        & "WHERE ClinicalTrialId = " & glSelectedTrialId & " " _
        & "ORDER BY LocalIdentifier1"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do Until rsTemp.EOF
        If rsTemp!LocalIdentifier1 <> "" Then
            cboCatCodes.AddItem rsTemp!LocalIdentifier1
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PopulatecboCatCodesWithSubjectLabels")
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
Private Sub DisplayOutPut()
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    If gbUseShortCodes Then
        Set gColQuestionCodes = New Collection
    End If
    
    'PrepareOutPut creates the recordset grsOutPut with the required fields/Questions
    Call PrepareOutPut
    
    If QueryCancelled Then Exit Sub
    
    'LoadOutPut places the contents of  mrsData into the recordset grsOutPut
    Call LoadOutPut
    
    If QueryCancelled Then Exit Sub
    
    'GridOutPut makes the recordset grsOutPut the Datasource of the grid grdOutPut
    Call GridOutPut

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "DisplayOutPut")
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
Public Sub QueryDB()
'---------------------------------------------------------------------
Dim sDataSQL As String
Dim sDataItemNamesSQL As String
Dim sSubjectRecordsSQL As String
Dim rsSubjectRecords As ADODB.Recordset

    On Error GoTo Errhandler

    Call DisplayProgressMessage("Preparing to Query Database")
    
    If QueryCancelled Then Exit Sub
    
    'CreateQuery returns 3 SQL statements:-
    'sDataSQL will retrieve the currently specified data
    'sDataItemNamesSQL will retrieve the DataItemNames of the specified data
    'sSubjectRecordsSQL will retrieve the number of Subject Records
    Call CreateQuerySQL(sDataSQL, sDataItemNamesSQL, sSubjectRecordsSQL)
    
    If QueryCancelled Then Exit Sub
    
    Call DisplayProgressMessage("Querying Database for Question Details.")
    'create a recordset of the currently specified DataItemNames
    Set mrsDataItemNames = New ADODB.Recordset

    mrsDataItemNames.Open sDataItemNamesSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    'Check the RecordCount does not exeed the limit of 2040 (2048-8 for the record identification fields)
    If mrsDataItemNames.RecordCount > 2040 Then
        Call DialogInformation("You have selected " & mrsDataItemNames.RecordCount & " Questions," & vbNewLine & _
            "which is more than the 2040 limit." & vbNewLine & _
            "Your Query has been terminated.", "MACRO Query Module")
        mrsDataItemNames.Close
        Set mrsDataItemNames = Nothing
        gbCancelled = True
    End If
    
    If QueryCancelled Then Exit Sub
    
    Call DisplayProgressMessage("Querying Database for Question Responses.")
    'Create a recordset of the required data
    Set mrsData = New ADODB.Recordset
    mrsData.Open sDataSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    'Display the number of Results
    txtResultsCount.Text = mrsData.RecordCount
    DoEvents    'to allow txtResultsCount to get updated
    
    If QueryCancelled Then Exit Sub
    
    Call DisplayProgressMessage("Querying Database for Subject Details.")
    'Create a recordset of Subject Records
    Set rsSubjectRecords = New ADODB.Recordset
    rsSubjectRecords.Open sSubjectRecordsSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    'Display the number of Subject Records
    txtRecordsCount.Text = rsSubjectRecords.RecordCount
    DoEvents    'to allow txtRecordsCount to get updated
    
    If QueryCancelled Then Exit Sub
    
    Call DisplayProgressMessage("Querying of Database Completed.")

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "QueryDB")
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
Private Function QueryCancelled() As Boolean
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    If gbCancelled Then
        Call DisplayProgressMessage("Query action has been Cancelled.")
        'Clear current grid and remove any splits
        Call ClearGrid
        'Clear the Retrieved Result and Records text boxes
        txtResultsCount.Text = ""
        txtRecordsCount.Text = ""
        DoEvents
        Call EnableRunSaveDisplay
        'Force cmdSaveOutPut to be disabled
        cmdSaveOutPut.Enabled = False
    End If

    QueryCancelled = gbCancelled

Exit Function
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "QueryCancelled")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'---------------------------------------------------------------------
Public Sub PrepareOutPut()
'---------------------------------------------------------------------
Dim sVisitFormQuestion As String
Dim sVFQKey As String
Dim nQuestionAttributes As Integer
Dim sVFQLongCode As String
Dim i As Integer
Dim sDataItemFormat As String
Dim nCatCodeLength As Integer
Dim bCatCodesNumeric As Boolean

    On Error GoTo Errhandler

    Call DisplayProgressMessage("Preparing Column Headers")
    
    'initialise the VisitFormQuestionLookup collection
    Set mcolVFQLookUp = New Collection
    
    Set grsOutPut = New ADODB.Recordset
    With grsOutPut
        'set the cursor location to Client
        .CursorLocation = adUseClient
        'Add the Subject identification fields
        .Fields.Append "Trial", adVarChar, 15
        .Fields.Append "Site", adVarChar, 8
        .Fields.Append "Label", adVarChar, 50
        .Fields.Append "PersonId", adInteger, -1, adFldIsNullable
        'Store the DataItemlength under Precision for the numeric fields
        .Fields.Item("PersonId").Precision = 10
        .Fields.Append "VisitCycle", adSmallInt, -1, adFldIsNullable
        .Fields.Item("VisitCycle").Precision = 5
        .Fields.Append "FormCycle", adSmallInt, -1, adFldIsNullable
        .Fields.Item("FormCycle").Precision = 5
        .Fields.Append "RepeatNumber", adSmallInt
        .Fields.Item("RepeatNumber").Precision = 5
        'Loop through the specified DataItemNames recordset, adding each DataItem/Question to grsOutPut
        mrsDataItemNames.MoveFirst
        Do While Not mrsDataItemNames.EOF
            'Prepare a question header
            If gbUseShortCodes Then
                sVFQLongCode = mrsDataItemNames!VisitCode & "/" & mrsDataItemNames!CRFPageCode & "/" & mrsDataItemNames!DataItemCode
                sVisitFormQuestion = CreateUniqueQuestion(mrsDataItemNames!DataItemCode, sVFQLongCode)
            Else
                sVisitFormQuestion = mrsDataItemNames!VisitCode & "/" & mrsDataItemNames!CRFPageCode & "/" & mrsDataItemNames!DataItemCode
            End If
            'Prepare VisitFormQuestion Lookup key
            sVFQKey = mrsDataItemNames!VisitId & "|" & mrsDataItemNames!CRFPageId & "|" & mrsDataItemNames!DataItemId
            'assess the data type of the required question
            Select Case mrsDataItemNames!DataType
            'Mo 25/10/2005 COD0040
            Case DataType.Text, DataType.Thesaurus
                'Force single character fields to be two charater fields that are
                'capable of holding a special value string of "-1" to "-9"
                If mrsDataItemNames!DataItemLength = 1 Then
                    .Fields.Append sVisitFormQuestion, adVarChar, 2, adFldIsNullable
                Else
                    .Fields.Append sVisitFormQuestion, adVarChar, mrsDataItemNames!DataItemLength, adFldIsNullable
                End If
            Case DataType.Category
                If gbOutputCategoryCodes Then
                    Call AssessCategoryCodes(glSelectedTrialId, 1, mrsDataItemNames!DataItemId, bCatCodesNumeric, nCatCodeLength)
                    'Force single character fields to be two character fields that are
                    'capable of holding a special value string of "-1" to "-9"
                    If nCatCodeLength = 1 Then
                        nCatCodeLength = 2
                    End If
                    If bCatCodesNumeric Then
                        'Mo Morris 1/10/03 changed from adSmallInt to adInteger, this stops
                        'MACRO_QM crashing on category response data with long numeric codes
                        .Fields.Append sVisitFormQuestion, adInteger, -1, adFldIsNullable
                        'Store nCatCodeLength under Precision
                        .Fields.Item(sVisitFormQuestion).Precision = nCatCodeLength
                    Else
                        .Fields.Append sVisitFormQuestion, adVarChar, nCatCodeLength, adFldIsNullable
                    End If
                Else
                    'Treat Category Questions with category Values as a Text Question of Length mrsDataItemNames!DataItemLength
                    'Mo 1/8/2006 Bug 2775, Force single character fields to be two character fields
                    'that are capable of holding a special value string of "-1" to "-9"
                    If mrsDataItemNames!DataItemLength = 1 Then
                        .Fields.Append sVisitFormQuestion, adVarChar, 2, adFldIsNullable
                    Else
                        .Fields.Append sVisitFormQuestion, adVarChar, mrsDataItemNames!DataItemLength, adFldIsNullable
                    End If
                End If
            Case DataType.Multimedia
                'Multimedia fields are always 36 char long
                .Fields.Append sVisitFormQuestion, adVarChar, 36, adFldIsNullable
            Case DataType.Date
                sDataItemFormat = mrsDataItemNames!DataItemFormat
                ' NCJ 20 Dec 05 - Added Partial Date flag check
                If CInt(RemoveNull(mrsDataItemNames!DataItemCase)) = 0 And DateFormatCanBeConverted(sDataItemFormat) Then
                    .Fields.Append sVisitFormQuestion, adDBTimeStamp, -1, adFldIsNullable
                    'Use Precision to store content type. Date/Time = 1, Date = 2, Time = 3
                    Select Case sDataItemFormat
                    Case "d/m/y/h/m", "m/d/y/h/m", "y/m/d/h/m", "d/m/y/h/m/s", "m/d/y/h/m/s", "y/m/d/h/m/s"
                        .Fields.Item(sVisitFormQuestion).Precision = 1
                    Case "d/m/y", "m/d/y", "y/m/d"
                        .Fields.Item(sVisitFormQuestion).Precision = 2
                    Case "h/m", "h/m/s"
                        .Fields.Item(sVisitFormQuestion).Precision = 3
                    End Select
                Else
                    ' Not a "standard" date
                    .Fields.Append sVisitFormQuestion, adVarChar, mrsDataItemNames!DataItemLength, adFldIsNullable
                End If
            Case DataType.IntegerData
                .Fields.Append sVisitFormQuestion, adInteger, -1, adFldIsNullable
                'Store the DataItemlength under Precision
                'Mo 1/8/2006 Bug 2775, Force single character fields to be two character fields
                'that are capable of holding a special value string of "-1" to "-9"
                If mrsDataItemNames!DataItemLength = 1 Then
                    .Fields.Item(sVisitFormQuestion).Precision = 2
                Else
                    .Fields.Item(sVisitFormQuestion).Precision = mrsDataItemNames!DataItemLength
                End If
            Case DataType.Real, DataType.LabTest
                'Mo 31/1/2007 Bug 2873, store Real and LabTest responses in Decimal fields, changed from Single
                .Fields.Append sVisitFormQuestion, adDecimal, -1, adFldIsNullable
                'Store the DataItemlength under Precision
                .Fields.Item(sVisitFormQuestion).Precision = mrsDataItemNames!DataItemLength
                'store the number of decimal places in the fields NumericScale property that is not used
                If InStr(mrsDataItemNames!DataItemFormat, ".") = 0 Then
                    .Fields.Item(sVisitFormQuestion).NumericScale = 0
                Else
                    .Fields.Item(sVisitFormQuestion).NumericScale = Len(mrsDataItemNames!DataItemFormat) - InStr(mrsDataItemNames!DataItemFormat, ".")
                End If
            End Select
            'add to VisitFormQuiestionName LookUp collection
            mcolVFQLookUp.Add sVisitFormQuestion, sVFQKey
            'Check for any question attributes having been selected and create appropriate heading
            If mcolQuestionAttributes(Str(mrsDataItemNames!DataItemId)) > 0 Then
                nQuestionAttributes = mcolQuestionAttributes(Str(mrsDataItemNames!DataItemId))
                If (nQuestionAttributes And mnMASK_COMMENTS) > 0 Then
                    If gbUseShortCodes Then
                        sVisitFormQuestion = CreateUniqueQuestion(mrsDataItemNames!DataItemCode, sVFQLongCode & "/Comments")
                        .Fields.Append sVisitFormQuestion, adVarChar, 255, adFldIsNullable
                        mcolVFQLookUp.Add sVisitFormQuestion, sVFQKey & "|" & mnMASK_COMMENTS
                    Else
                        .Fields.Append sVisitFormQuestion & "/Comments", adVarChar, 255, adFldIsNullable
                        mcolVFQLookUp.Add sVisitFormQuestion & "/Comments", sVFQKey & "|" & mnMASK_COMMENTS
                    End If
                End If
                If (nQuestionAttributes And mnMASK_CTCGRADE) > 0 Then
                    If gbUseShortCodes Then
                        sVisitFormQuestion = CreateUniqueQuestion(mrsDataItemNames!DataItemCode, sVFQLongCode & "/CTCGrade")
                        .Fields.Append sVisitFormQuestion, adVarChar, 1, adFldIsNullable
                        mcolVFQLookUp.Add sVisitFormQuestion, sVFQKey & "|" & mnMASK_CTCGRADE
                    Else
                        .Fields.Append sVisitFormQuestion & "/CTCGrade", adVarChar, 1, adFldIsNullable
                        mcolVFQLookUp.Add sVisitFormQuestion & "/CTCGrade", sVFQKey & "|" & mnMASK_CTCGRADE
                    End If
                End If
                If (nQuestionAttributes And mnMASK_LABRESULT) > 0 Then
                    If gbUseShortCodes Then
                        sVisitFormQuestion = CreateUniqueQuestion(mrsDataItemNames!DataItemCode, sVFQLongCode & "/LabResult")
                        .Fields.Append sVisitFormQuestion, adVarChar, 1, adFldIsNullable
                        mcolVFQLookUp.Add sVisitFormQuestion, sVFQKey & "|" & mnMASK_LABRESULT
                    Else
                        .Fields.Append sVisitFormQuestion & "/LabResult", adVarChar, 1, adFldIsNullable
                        mcolVFQLookUp.Add sVisitFormQuestion & "/LabResult", sVFQKey & "|" & mnMASK_LABRESULT
                    End If
                End If
                If (nQuestionAttributes And mnMASK_STATUS) > 0 Then
                    If gbUseShortCodes Then
                        sVisitFormQuestion = CreateUniqueQuestion(mrsDataItemNames!DataItemCode, sVFQLongCode & "/Status")
                        .Fields.Append sVisitFormQuestion, adVarChar, 15, adFldIsNullable
                        mcolVFQLookUp.Add sVisitFormQuestion, sVFQKey & "|" & mnMASK_STATUS
                    Else
                        .Fields.Append sVisitFormQuestion & "/Status", adVarChar, 15, adFldIsNullable
                        mcolVFQLookUp.Add sVisitFormQuestion & "/Status", sVFQKey & "|" & mnMASK_STATUS
                    End If
                End If
                If (nQuestionAttributes And mnMASK_TIMESTAMP) > 0 Then
                    If gbUseShortCodes Then
                        sVisitFormQuestion = CreateUniqueQuestion(mrsDataItemNames!DataItemCode, sVFQLongCode & "/TimeStamp")
                        .Fields.Append sVisitFormQuestion, adDBTimeStamp, -1, adFldIsNullable
                        '.Fields.Append sVisitFormQuestion, adVarChar, 255, adFldIsNullable
                        mcolVFQLookUp.Add sVisitFormQuestion, sVFQKey & "|" & mnMASK_TIMESTAMP
                    Else
                        .Fields.Append sVisitFormQuestion & "/TimeStamp", adDBTimeStamp, -1, adFldIsNullable
                        '.Fields.Append sVisitFormQuestion & "/TimeStamp", adVarChar, 255, adFldIsNullable
                        mcolVFQLookUp.Add sVisitFormQuestion & "/TimeStamp", sVFQKey & "|" & mnMASK_TIMESTAMP
                    End If
                End If
                If (nQuestionAttributes And mnMASK_USERNAME) > 0 Then
                    If gbUseShortCodes Then
                        sVisitFormQuestion = CreateUniqueQuestion(mrsDataItemNames!DataItemCode, sVFQLongCode & "/UserName")
                        .Fields.Append sVisitFormQuestion, adVarChar, 20, adFldIsNullable
                        mcolVFQLookUp.Add sVisitFormQuestion, sVFQKey & "|" & mnMASK_USERNAME
                    Else
                        .Fields.Append sVisitFormQuestion & "/UserName", adVarChar, 20, adFldIsNullable
                        mcolVFQLookUp.Add sVisitFormQuestion & "/UserName", sVFQKey & "|" & mnMASK_USERNAME
                    End If
                End If
                If (nQuestionAttributes And mnMASK_VALUECODE) > 0 Then
                    If gbUseShortCodes Then
                        sVisitFormQuestion = CreateUniqueQuestion(mrsDataItemNames!DataItemCode, sVFQLongCode & "/ValueCode")
                        .Fields.Append sVisitFormQuestion, adVarChar, 15, adFldIsNullable
                        mcolVFQLookUp.Add sVisitFormQuestion, sVFQKey & "|" & mnMASK_VALUECODE
                    Else
                        .Fields.Append sVisitFormQuestion & "/ValueCode", adVarChar, 15, adFldIsNullable
                        mcolVFQLookUp.Add sVisitFormQuestion & "/ValueCode", sVFQKey & "|" & mnMASK_VALUECODE
                    End If
                End If
            End If
            'Check for cancel having been hit
            DoEvents
            If QueryCancelled Then Exit Do
            mrsDataItemNames.MoveNext
        Loop
        .Open
    End With
    
    'Check for a Cancel in the above loop
    If gbCancelled Then
        Exit Sub
    End If
    
    'Place the widths of the grid's/recordset's column headers into the Column Widths Collection
    Set mcolColumnWidths = New Collection
    For i = 0 To grsOutPut.Fields.Count - 1
        mcolColumnWidths.Add TextWidth(grsOutPut.Fields(i).Name & "  "), grsOutPut.Fields(i).Name
    Next i
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PrepareOutPut")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'Mo 2/4/2007 MRC15022007, Optional parameter bShowDialogs added
'---------------------------------------------------------------------
Public Sub LoadOutPut(Optional bShowDialogs As Boolean = True)
'---------------------------------------------------------------------
Dim nOldData As Integer
Dim sPrevKey As String
Dim sKeyCollection As String
Dim sVFQKey As String
Dim sVFQAKey As String
Dim sVisitFormQuestion As String
Dim nQuestionAttributes As Integer
Dim sMessage As String
Dim sTrialName As String
Dim sSite As String
Dim sLabel As String
Dim vResponse As Variant
Dim lSubRecNo As Long

    On Error GoTo Errhandler

    Call DisplayProgressMessage("Loading Responses")
    
    'Load the response data from mrsData into grsOutPut
    nOldData = 0
    sPrevKey = ""
    lSubRecNo = 0
    mrsData.MoveFirst
    Do While Not mrsData.EOF
        sKeyCollection = mrsData!TrialSite & "|" & mrsData!PersonID & "|" _
            & mrsData!VisitCycleNumber & "|" & mrsData!CRFPageCycleNumber & "|" & mrsData!RepeatNumber
        If sKeyCollection <> sPrevKey Then
            sPrevKey = sKeyCollection
            grsOutPut.AddNew
            lSubRecNo = lSubRecNo + 1
            Call DisplayProgressMessage("Loading Subject Record " & lSubRecNo)
            sTrialName = TrialNameFromId(glSelectedTrialId)
            grsOutPut.Fields("Trial") = sTrialName
            Call WidthCheck(sTrialName, "Trial")
            sSite = mrsData!TrialSite
            grsOutPut.Fields("Site") = sSite
            Call WidthCheck(sSite, "Site")
            sLabel = SubjectLabelFromTrialSiteId(glSelectedTrialId, mrsData!TrialSite, mrsData!PersonID)
            grsOutPut.Fields("Label") = sLabel
            Call WidthCheck(sLabel, "Label")
            grsOutPut.Fields("PersonId") = mrsData!PersonID
            grsOutPut.Fields("VisitCycle") = mrsData!VisitCycleNumber
            grsOutPut.Fields("FormCycle") = mrsData!CRFPageCycleNumber
            grsOutPut.Fields("RepeatNumber") = mrsData!RepeatNumber
        End If

        sVFQKey = mrsData!VisitId & "|" & mrsData!CRFPageId & "|" & mrsData!DataItemId
        'check that the visit/form/question of this response still exists in the study definition
        'Note that QuestionExistsInStudy updates sVisitFormQuestion
        If Not QuestionExistsInStudy(sVFQKey, sVisitFormQuestion) Then
            nOldData = nOldData + 1
        Else
            'format the response and place into the OutPut recordset
            vResponse = FormatOutPut(mrsData!DataItemId, mrsData!ResponseValue, mrsData!ValueCode, mrsData!ResponseStatus)
             grsOutPut.Fields(sVisitFormQuestion) = vResponse
            Call WidthCheck(vResponse, sVisitFormQuestion)
            'Check for any question attributes having been selected
            If mcolQuestionAttributes(Str(mrsData!DataItemId)) > 0 Then
                nQuestionAttributes = mcolQuestionAttributes(Str(mrsData!DataItemId))
                If (nQuestionAttributes And mnMASK_COMMENTS) > 0 Then
                    sVFQAKey = sVFQKey & "|" & mnMASK_COMMENTS
                    sVisitFormQuestion = mcolVFQLookUp(sVFQAKey)
                    grsOutPut.Fields(sVisitFormQuestion) = mrsData!Comments
                    Call WidthCheck(mrsData!Comments, sVisitFormQuestion)
                End If
                If (nQuestionAttributes And mnMASK_CTCGRADE) > 0 Then
                    sVFQAKey = sVFQKey & "|" & mnMASK_CTCGRADE
                    sVisitFormQuestion = mcolVFQLookUp(sVFQAKey)
                    Select Case mrsData!CTCGrade
                    Case 1, 2, 3, 4, 5
                        grsOutPut.Fields(sVisitFormQuestion) = mrsData!CTCGrade
                    End Select
                End If
                If (nQuestionAttributes And mnMASK_LABRESULT) > 0 Then
                    sVFQAKey = sVFQKey & "|" & mnMASK_LABRESULT
                    sVisitFormQuestion = mcolVFQLookUp(sVFQAKey)
                    Select Case mrsData!LabResult
                    Case LabResult.Low
                        grsOutPut.Fields(sVisitFormQuestion) = "L"
                    Case LabResult.Normal
                        grsOutPut.Fields(sVisitFormQuestion) = "N"
                    Case LabResult.High
                        grsOutPut.Fields(sVisitFormQuestion) = "H"
                    End Select
                End If
                If (nQuestionAttributes And mnMASK_STATUS) > 0 Then
                    sVFQAKey = sVFQKey & "|" & mnMASK_STATUS
                    sVisitFormQuestion = mcolVFQLookUp(sVFQAKey)
                    Select Case mrsData!ResponseStatus
                    Case Status.Requested
                        grsOutPut.Fields(sVisitFormQuestion) = "Requested"
                        Call WidthCheck("Requested", sVisitFormQuestion)
                    Case Status.CancelledByUser
                        grsOutPut.Fields(sVisitFormQuestion) = "Cancelled"
                        Call WidthCheck("Cancelled", sVisitFormQuestion)
                    Case Status.NotApplicable
                        grsOutPut.Fields(sVisitFormQuestion) = "Not Applicable"
                        Call WidthCheck("Not Applicable", sVisitFormQuestion)
                    Case Status.Unobtainable
                        grsOutPut.Fields(sVisitFormQuestion) = "Unobtainable"
                        Call WidthCheck("Unobtainable", sVisitFormQuestion)
                    Case Status.Success
                        grsOutPut.Fields(sVisitFormQuestion) = "OK"
                        Call WidthCheck("OK", sVisitFormQuestion)
                    Case Status.Missing
                        grsOutPut.Fields(sVisitFormQuestion) = "Missing"
                        Call WidthCheck("Missing", sVisitFormQuestion)
                    Case Status.Inform
                        grsOutPut.Fields(sVisitFormQuestion) = "Inform"
                        Call WidthCheck("Inform", sVisitFormQuestion)
                    Case Status.OKWarning
                        grsOutPut.Fields(sVisitFormQuestion) = "OK Warning"
                        Call WidthCheck("OK Warning", sVisitFormQuestion)
                    Case Status.Warning
                        grsOutPut.Fields(sVisitFormQuestion) = "Warning"
                        Call WidthCheck("Warning", sVisitFormQuestion)
                    Case Status.InvalidData
                        grsOutPut.Fields(sVisitFormQuestion) = "Invalid"
                        Call WidthCheck("Invalid", sVisitFormQuestion)
                    End Select
                End If
                If (nQuestionAttributes And mnMASK_TIMESTAMP) > 0 Then
                    sVFQAKey = sVFQKey & "|" & mnMASK_TIMESTAMP
                    sVisitFormQuestion = mcolVFQLookUp(sVFQAKey)
                    grsOutPut.Fields(sVisitFormQuestion) = Format(CDate(mrsData!ResponseTimeStamp), "yyyy/mm/dd hh:mm:ss")
                    Call WidthCheck(Format(CDate(mrsData!ResponseTimeStamp), "yyyy/mm/dd hh:mm:ss"), sVisitFormQuestion)
                End If
                If (nQuestionAttributes And mnMASK_USERNAME) > 0 Then
                    sVFQAKey = sVFQKey & "|" & mnMASK_USERNAME
                    sVisitFormQuestion = mcolVFQLookUp(sVFQAKey)
                    grsOutPut.Fields(sVisitFormQuestion) = mrsData!UserName
                    Call WidthCheck(mrsData!UserName, sVisitFormQuestion)
                End If
                If (nQuestionAttributes And mnMASK_VALUECODE) > 0 Then
                    sVFQAKey = sVFQKey & "|" & mnMASK_VALUECODE
                    sVisitFormQuestion = mcolVFQLookUp(sVFQAKey)
                    grsOutPut.Fields(sVisitFormQuestion) = mrsData!ValueCode
                    Call WidthCheck(mrsData!ValueCode, sVisitFormQuestion)
                End If
            End If
        End If
        'Check for cancel having been hit
        DoEvents
        If QueryCancelled Then Exit Do
        mrsData.MoveNext
    Loop
    
    'Check for a Cancel in the above loop
    If gbCancelled Then
        Exit Sub
    Else
        Call DisplayProgressMessage("Loading of Subject Records Completed")
    End If

    If bShowDialogs Then
        'Check for responses that no longer exist in the study having been found
        If nOldData > 0 Then
            sMessage = nOldData & " question response"
            If nOldData > 1 Then sMessage = sMessage & "s"
            sMessage = sMessage & " that no longer exist in the study "
            If nOldData = 1 Then
                sMessage = sMessage & "was"
            Else
                sMessage = sMessage & "were"
            End If
            sMessage = sMessage & " not retrieved."
            Call DialogInformation(sMessage, "MACRO Query")
        End If
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadOutPut")
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
Public Sub WidthCheck(vContents As Variant, sColumnName As String)
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    If TextWidth(vContents & "  ") > mcolColumnWidths(sColumnName) Then
        mcolColumnWidths.Remove sColumnName
        mcolColumnWidths.Add TextWidth(vContents & "  "), sColumnName
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "WidthCheck")
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
Private Sub GridOutPut()
'---------------------------------------------------------------------
Dim i As Integer

    On Error GoTo Errhandler

    Set grdOutPut.Datasource = grsOutPut

    'Set the Max column Widths
    For i = 0 To grdOutPut.Columns.Count - 1
        grdOutPut.Columns(i).Width = mcolColumnWidths(grdOutPut.Columns(i).Caption) + 10
    Next i

    'Add a datagrid split if it is requested
    If gbSplitGrid Then
        GridSplitAdd
    Else
        'Hide the unwanted identification fields
        'If a split exists then GridSplitAdd will handle the hiding of unwanted identification fields
        If Not gbDisplayStudyName Then
            grdOutPut.Splits(0).Columns(0).Visible = False
        End If
        If Not gbDisplaySiteCode Then
            grdOutPut.Splits(0).Columns(1).Visible = False
        End If
        If Not gbDisplayLabel Then
            grdOutPut.Splits(0).Columns(2).Visible = False
        End If
        If Not gbDisplayPersonId Then
            grdOutPut.Splits(0).Columns(3).Visible = False
        End If
        If Not gbDisplayVisitCycle Then
            grdOutPut.Splits(0).Columns(4).Visible = False
        End If
        If Not gbDisplayFormCycle Then
            grdOutPut.Splits(0).Columns(5).Visible = False
        End If
        If Not gbDisplayRepeatNumber Then
            grdOutPut.Splits(0).Columns(6).Visible = False
        End If
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GridOutPut")
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
Private Sub DisableRunSaveDisplay()
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbQueryRunning = True
    
    'Disable the RunQuery SaveOutput buttons while the query is running
    cmdRunQuery.Enabled = False
    cmdSaveOutPut.Enabled = False
    
    'Disable the Display options while the query is running
    fraDisplayOptions.Enabled = False
    
    'Disable the Studies combo
    cboStudies.Enabled = False
    
    'Disable the exit button
    cmdExit.Enabled = False
    
    'Disable the Menu items
    mnuFile.Enabled = False
    mnuView.Enabled = False
    mnuRun.Enabled = False
    mnuHelp.Enabled = False
    
    'Enable the CancelRun button
    cmdCancelRun.Enabled = True
    gbCancelled = False

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "DisableRunSaveDisplay")
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
Private Sub EnableRunSaveDisplay()
'---------------------------------------------------------------------
    
    On Error GoTo Errhandler
    
    mbQueryRunning = False
    
    'If there is some output enable the save output command button
    If Not mrsData Is Nothing Then
        If mrsData.RecordCount > 0 Then
            cmdSaveOutPut.Enabled = True
        End If
    Else
        cmdSaveOutPut.Enabled = False
    End If
    
    'Enable the RunQuery button when the query is complete
    cmdRunQuery.Enabled = True
    
    'Enable the Display options when the query is complete
    fraDisplayOptions.Enabled = True
    
    'Enable the Studies combo
    cboStudies.Enabled = True
    
    'Enable the exit button
    cmdExit.Enabled = True
    
    'Enable the Menu items
    mnuFile.Enabled = True
    mnuView.Enabled = True
    mnuRun.Enabled = True
    mnuHelp.Enabled = True
    
    'Disable the CancelRun button
    cmdCancelRun.Enabled = False
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "EnableRunSaveDisplay")
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
Public Function ForgottenPassword(sSecurityCon As String, sUsername As String, sPassword As String, ByRef sErrMsg As String) As eDTForgottenPassword
'---------------------------------------------------------------------
'REM 06/12/02
'---------------------------------------------------------------------

    'dummy routine

End Function
