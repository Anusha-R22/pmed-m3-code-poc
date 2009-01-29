VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   Caption         =   "Create Data Views"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frmMenuCreateDataViews.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStudy 
      Caption         =   "Study Selection"
      Height          =   1095
      Left            =   120
      TabIndex        =   36
      Top             =   2160
      Width           =   3495
      Begin VB.OptionButton optAllStudies 
         Caption         =   "All studies"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.OptionButton optSingleStudy 
         Caption         =   "Only this study:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cboStudies 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCreateTriggers 
      Caption         =   "Create Triggers"
      Height          =   375
      Left            =   3840
      TabIndex        =   35
      Top             =   7800
      Width           =   1600
   End
   Begin VB.CommandButton cmdCreateDataViews 
      Caption         =   "Create Data Views"
      Height          =   375
      Left            =   2040
      TabIndex        =   34
      Top             =   7800
      Width           =   1600
   End
   Begin VB.CommandButton cmdCreateViewNames 
      Caption         =   "Create View Names"
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   7800
      Width           =   1635
   End
   Begin VB.Frame Frame7 
      Caption         =   "Missing Response Special Values"
      Height          =   2000
      Left            =   4920
      TabIndex        =   29
      Top             =   120
      Width           =   2800
      Begin VB.TextBox txtMissing 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2600
      End
      Begin VB.TextBox txtUnobtainable 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Width           =   2600
      End
      Begin VB.TextBox txtNotApplicable 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1600
         Width           =   2600
      End
      Begin VB.Label Label4 
         Caption         =   "For Status Missing insert"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   300
         Width           =   2205
      End
      Begin VB.Label Label1 
         Caption         =   "For Status Unobtainable insert"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   2205
      End
      Begin VB.Label Label3 
         Caption         =   "For Status NotApplicable insert"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   1400
         Width           =   2205
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Category Codes or Values"
      Height          =   1095
      Left            =   3720
      TabIndex        =   28
      Top             =   2160
      Width           =   3975
      Begin VB.OptionButton optCatCodesTyped 
         Caption         =   "As Codes in Numeric columns when applicable."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   3720
      End
      Begin VB.OptionButton optCatValues 
         Caption         =   "As Values in Text columns (e.g. Female, Male)."
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3720
      End
      Begin VB.OptionButton optCatCodes 
         Caption         =   "As Codes in Text columns (e.g. F,M or 1,2)."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3600
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "Date Last Run"
      Height          =   615
      Left            =   120
      TabIndex        =   27
      Top             =   7155
      Width           =   5415
      Begin VB.TextBox txtCreateTriggersDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtCreateDataViewsDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtCreateViewNamesDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data View Table Names"
      Height          =   3135
      Left            =   120
      TabIndex        =   25
      Top             =   4005
      Width           =   7620
      Begin MSComctlLib.ListView lvwTables 
         Height          =   1695
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   2990
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   5460
         TabIndex        =   11
         Top             =   2040
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Edit Data View Table Name"
         Height          =   675
         Left            =   120
         TabIndex        =   26
         Top             =   2340
         Width           =   6435
         Begin VB.CommandButton cmdChange 
            Caption         =   "Change"
            Height          =   375
            Left            =   5340
            TabIndex        =   14
            Top             =   180
            Width           =   975
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   4260
            TabIndex        =   13
            Top             =   200
            Width           =   975
         End
         Begin VB.TextBox txtTableName 
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   4035
         End
      End
   End
   Begin VB.TextBox txtProgressMessage 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3585
      Width           =   7620
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data View Types (select one or both options)"
      Height          =   930
      Left            =   120
      TabIndex        =   21
      Top             =   1180
      Width           =   4700
      Begin VB.CheckBox chkResponseValuePlus 
         Caption         =   "Response Value plus attributes (one question per row)."
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   4215
      End
      Begin VB.CheckBox chkResponseValueOnly 
         Caption         =   "Response Value only (one subject per row)."
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Use of Visits"
      Height          =   940
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   4700
      Begin VB.OptionButton optVisitTogether 
         Caption         =   "One Data View per eForm irrespective of the Visit."
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3915
      End
      Begin VB.OptionButton optVisitSeparate 
         Caption         =   "Separate Data Views for each Visit that an eForm occurs in."
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   4515
      End
   End
   Begin VB.Timer tmrSystemIdleTimeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5640
      Top             =   7320
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      Top             =   7800
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar sbrMenu 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   8280
      Width           =   7845
      _ExtentX        =   13838
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
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Key             =   "UserDatabase"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Create Data Views Progress:"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   3345
      Width           =   2115
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmMenu (frmCreateDataViews.frm)
'   Author:     Mo Morris, January 2001
'   Purpose:    Started life as frmMenu in the Macro Stub. Now handles all the processing
'               required by the Create Data Views executable.
'----------------------------------------------------------------------------------------'
'  RJCW 07/09/2001  Changes for SQL Server Triggers
'  Modified Procs:  RefreshDataViewDetails
'                   CmdCreateTriggers_Click
'                   InitialiseMe
'  New Procs:       CheckSQLServerTriggerTableStructure
'                   AlterDataViewDetails
'
'   Mo 16/4/2002    Changes to CreateRODataViewTable and CreateAndPopulateDataViews, data view
'                   question response columns are now typed based on DataItem!DataType.
'
'                   Responses to category questions can now optionally be Codes (e.g. M, F)
'                   or Values (e.g. Male, Female), with Value being the default.
'
'                   Field OutputCategoryValues has been added to table DataViewDetails to hold the
'                   category question's Values or Codes option. Values is the default.
'                   mbOutputCategoryValues stores this setting optCatCodes and optCatValues control its changing.
'   Mo 30/4/20002   A ReplaceQuotes added to ResponseValue in CreateAndPopulateDataViews
'   Mo 18/6/2002    Standard dates/times are now converted into DateTime fields
'                   in RO (Result Only) Data View Tables.
'                   Changes to CreateAndpopulateDataViews and CreateRODataViewTable
'                   New functions DateFormatCanBeConverted and UniversalDateFormatString added.
'   ASH 23/07/2002  Added RQGs to dataviews and restructured most of routines .
'   Mo  3/9/2002    Changes around column widths within lvwTables now that the new column
'                   "Group" has been added. Changes to Form_Load, Form_Resize, CreateTableNames
'                   and PopulateListView, which now call new sub SetColumnWidths
'   ATO 4/9/2002    Added ShowColumnHeaders and FixTableNameLength routines
'   NCJ 23 Apr 03 - Use generic help for MACRO 3.0
'   TA 07/10/2004   number(16,10) datatype changed to float in RO dataviews
'   TA 07/10/2004   integer and number(11) datatype changed to BIGINT and number(15) in RO dataviews
'   Mo 24/11/2004 - Bug 2413, New Typed codes (numeric or alpha) option added.
'                   New Subs/functions rsDataValues, AssessCategoryCodes, optCatCodesTyped_Click added.
'                   Subs/functions changed RefreshDataViewDetails, DataBaseSpecificColumnFormatting,
'                   FormatResponse, CreateRODataViewTable, CreateRQGDataViewTable, PopulateROTable,
'                   optCatCodes_Click, optCatValues_Click, StoreDataViewOptions
'   Mo 3/12/2004    Bug 2446 - Adding Special Values facilities to Create Data Views.
'                   new frame of controls (Frame7) added.
'                   Changes to CreateDataViewDetails, UpdateTriggerTableAndCategoryValues, Form_Resize,
'                   RefreshDataViewDetails, StoreDataViewOptions, PopulateROTable and FormatResponse
'                   New Function DataViewDetailsContainsSpecialValues and sub AddSpecialValuesToDataViewDetails
'   MLM 25/04/05    Added study-specific data views.
'   MLM 13/09/05    bug 2632: Set default values of null for "special values"
'   Mo  25/10/2005  COD0030 - Changes around the new Thesaurus Data Item Type
'   Mo  24/10/2006  Bug 2824 - Make MACRO Create Data Views Module comply with Partial Dates.
'                   When deciding as to wether a date question should occupy a date or string field the
'                   Partial date flag now needs to be checked. Partial date question to occupy string fields.
'                   Changes made to functions DataBaseSpecificColumnFormatting and FormatResponse
'                   Corrections to some of the field types for Oracle databases have been made under this bug number
'                   Subs CreateRODataViewTable, CreateWADataViewTable and CreateRQGDataViewTable
'                   have been corrected and now have an Oracle database option.
'----------------------------------------------------------------------------------------'

Option Explicit

Private mcolTableNames As Collection
Private mnSelListViewEntry As Integer
Private msSelListViewEntryText As String
Private mbDataViewTableNamesExist As Boolean
Private mbDataViewTablesExist As Boolean
Private mbDataViewTriggersExist As Boolean
Private mbDataViewNameEditsEnabled As Boolean
Private mbDataViewSeparateVisits As Boolean
Private mbDataViewRORequired As Boolean
Private mbDataViewWARequired As Boolean
'Changed Mo Morris 16/4/2002
Private mnOutputCategoryValues As Integer
'Changed ASH 25/07/2002
Private mbDataViewAllVisits As Boolean
Private mcolMacroTableNames As Collection
'ATO 3/09/2002 Used in sub FixTableNameLength
Private mnTableNameLength As Integer
'MLM 15/04/05:
Private mlStudyId As Long

'---------------------------------------------------------------------
Private Sub cboStudies_Click()
'---------------------------------------------------------------------
' MLM 15/04/05: Copied from MACRO 2.1. Filter on the selected study.
'---------------------------------------------------------------------
    
    'when the form loads, the first study is automatically selected from the list, although "all studies" is selected.
    If Not cboStudies.Enabled Then
        Exit Sub
    End If
    
    mlStudyId = cboStudies.ItemData(cboStudies.ListIndex)
    
    'retrieve status details from table DataViewDetails
    RefreshDataViewDetails
    
    'if Data View Names exist then populate lvwTables with them
    PopulateListView

End Sub


'---------------------------------------------------------------------
Private Sub chkResponseValueOnly_Click()
'---------------------------------------------------------------------
'ASH 25/07/2002 - Used module level variables to replace object names
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
        mbDataViewRORequired = (chkResponseValueOnly.Value = vbChecked)
    
        If Not mbDataViewRORequired And Not mbDataViewWARequired Then
            cmdCreateViewNames.Enabled = False
        Else
            cmdCreateViewNames.Enabled = True
        End If

Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "chkResponseValueOnly_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub chkResponseValuePlus_Click()
'---------------------------------------------------------------------
'ASH 25/07/2002 - Used module level variables to replace object name
'--------------------------------------------------------------------

    On Error GoTo ErrLabel
    
        mbDataViewWARequired = (chkResponseValuePlus.Value = vbChecked)
    
        If Not mbDataViewRORequired And Not mbDataViewWARequired Then
            cmdCreateViewNames.Enabled = False
        Else
            cmdCreateViewNames.Enabled = True
        End If

Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "chkResponseValuePlus_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    txtTableName.Text = ""
    cmdCancel.Enabled = False
    cmdChange.Enabled = False
    lvwTables.Enabled = True
    
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdCancel_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub cmdChange_Click()
'---------------------------------------------------------------------
'Mo Morris 2/4/01   Table names checked for being an empty string.
'                   Alphanumeric/underscore message re-worded
'                   Data view table names now checked against the table names that already exist in Macro
'ASH 26/07/2002     Added function IsTableNameValid
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sTableName As String
    
    On Error GoTo ErrLabel
    
    sTableName = Trim(txtTableName.Text)
        
    If IsTableNameValid(sTableName) Then
        'place the edited tablename back into lvwTables
        lvwTables.ListItems(mnSelListViewEntry).Text = sTableName
        'Update the edited DataViewName in table DataViewTables
        sSQL = "UPDATE DataViewTables SET DataViewName = '" & sTableName & "'" _
            & " WHERE DataViewName = '" & msSelListViewEntryText & "'"
        MacroADODBConnection.Execute sSQL
    End If
    
    txtTableName.Text = ""
    cmdCancel.Enabled = False
    cmdChange.Enabled = False
    lvwTables.Enabled = True

Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdChange_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub cmdCreateTriggers_Click()
'---------------------------------------------------------------------
'Note that this sub requires the The 2 text files:-
'   DVROScript.txt
'   DVWAScript.txt
'Which should be in Macro's app.path
'---------------------------------------------------------------------
Dim sSQL As String
Dim nFileNumber As Integer
Dim sFileName As String
Dim sPrefix As String

    On Error GoTo ErrLabel
    
    HourglassOn

    '   RJCW 05/09/2001   Enable SQL Server Triggers
    '                     the SQL Server script file extension is ".sql"
    If goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        sPrefix = "ORA_"
    Else
        '   all sql server triggers must be dropped before creation
        RemoveTriggers
        sPrefix = "MSSQL_"
    End If

    If mbDataViewRORequired Then
        sFileName = gsAppPath & "Database Scripts\CDV\" & sPrefix & "DVROScript.sql"
        nFileNumber = FreeFile
        Open sFileName For Input As #nFileNumber
        sSQL = Input(LOF(nFileNumber), #nFileNumber)
        Close #nFileNumber
        MacroADODBConnection.Execute sSQL
    End If
    
    If mbDataViewWARequired Then
        sFileName = gsAppPath & "Database Scripts\CDV\" & sPrefix & "DVWAScript.sql"
        nFileNumber = FreeFile
        Open sFileName For Input As #nFileNumber
        sSQL = Input(LOF(nFileNumber), #nFileNumber)
        Close #nFileNumber
        MacroADODBConnection.Execute sSQL
    End If
    
    'Set Triggers created flag in table DataViewDetails
    sSQL = "UPDATE DataViewDetails Set DataViewTriggersExist = " & DataViewState.Created & "," _
        & " DataViewTriggersDate = " & SQLStandardNow
    'MLM 15/04/05:
    If mlStudyId > 0 Then
        sSQL = sSQL & " WHERE ClinicalTrialId = " & mlStudyId
    End If
    MacroADODBConnection.Execute sSQL
        
    RefreshDataViewDetails
    
    HourglassOff

Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdCreateTriggers_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub cmdEdit_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    txtTableName.Enabled = True
    mnSelListViewEntry = lvwTables.SelectedItem.Index
    msSelListViewEntryText = lvwTables.SelectedItem.Text
    txtTableName.Text = lvwTables.SelectedItem.Text
    lvwTables.Enabled = False
    cmdCancel.Enabled = True
    cmdChange.Enabled = True
    cmdEdit.Enabled = False

Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdEdit_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub
'---------------------------------------------------------------------
Private Sub cmdCreateViewNames_Click()
'---------------------------------------------------------------------
'Initiates and calls routunes to create dataview table names
'---------------------------------------------------------------------
Dim sTitle As String
Dim sPrompt As String
Dim nResponse As Integer

    On Error GoTo ErrLabel
    
    If mbDataViewTableNamesExist Then
        sTitle = "Table Name Creation"
        sPrompt = "Table Names already exist." _
            & vbNewLine & "If you continue the existing names together with" _
            & vbNewLine & "any Data View Tables and Triggers will be removed."
        nResponse = DialogWarning(sPrompt, sTitle, True)
        If nResponse = vbCancel Then
            Exit Sub
        End If
    End If
    
    HourglassOn
    
    'MLM 13/03/06: Moved outside of if
    'store the Data View options that will be used to create names
    StoreDataViewOptions
    'remove any Triggers that might exist
    RemoveTriggers
    'remove the Data View Names from DataViewTables together with the actual Data View Tables
    RemoveDataViews
    RemoveDataViewTables
    
    cmdEdit.Enabled = False
    txtTableName.Text = ""
    cmdCancel.Enabled = False
    cmdChange.Enabled = False
    lvwTables.Enabled = True
        
    'ASH 26/07/2002 disable controls while CreateAndPopulateDataViews is running
    Call EnableControls(False)
    
    CreateTableNames
    
    'ASH 26/07/2002 Enable controls after CreateAndPopulateDataViews has completed
    Call EnableControls(True)
    
    HourglassOff
        
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdCreateViewNames_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub cmdExit_Click()
'---------------------------------------------------------------------
   
    Call ExitMACRO
    Call MACROEnd

End Sub

'---------------------------------------------------------------------
Public Sub InitialiseMe()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    'Changed Mo Morris 6/4/01, link to Create Data Views help added
    ' NCJ 23 Apr 03 - Use generic help for MACRO 3.0
'    gsMACROUserGuidePath = gsMACROUserGuidePath & "CDV\Contents.htm"
    
    'RJCW 26/11/01 call to populate Keywords Dictionary object
    FetchKeywords
    
    ' ATO 4/9/2002 get max length for tablename
    Call FixTableNameLength
    
    
    'ASH 25/07/2002 creates a collection of Macro tables
    CreateMacroTableCollection

   If QuestionNamesNotOK Then
        Call ExitMACRO
        Call MACROEnd
        Exit Sub
    End If
    
    'ASH 26/07/2002 set the default options
    SetControlsDefaultOptions
    
    'Check to see if table DataViewTables exists
    '(i.e. is this application being run for the first time)
    If Not TableExists("DataViewTables") Then
        CreateDataViewTables
        CreateDataViewDetails
    End If
    
    'ASH 26/07/2002 Updates SQL server triggers and category values
    UpdateTriggerTableAndCategoryValue
    
    'MLM 15/04/05: Here's a further database upgrade for data views
    UpgradeTo3072
    'populate the drop-down list of studies
    PopulateStudyList
    MaintainDataViewDetails
    
    'retrieve status details from table DataViewDetails
    RefreshDataViewDetails
    
    'if Data View Names exist then populate lvwTables with them
    If mbDataViewTableNamesExist Then
        PopulateListView
        'if the tables have not been created then they can still be edited
        'i.e lvwTables can be enabled
        If Not mbDataViewTablesExist Then
            lvwTables.Enabled = True
        End If
    End If

Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "InitialiseMe", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Public Sub CheckUserRights()
'---------------------------------------------------------------------
'Do not rewmove this Sub
'---------------------------------------------------------------------

End Sub

'---------------------------------------------------------------------
Private Sub CreateAndPopulateDataViews()
'---------------------------------------------------------------------
'This will read through the contents of table DataViewTables, creating
'each named Data View Table and populating it.
'Mo Morris 16/4/2002
'   Changes made so that the question responses are cast into different
'   types based on the question type:-
'
'   Question Type   Casting
'   Text            Keep as string
'   Category        Keep as string
'   Multimedia      Keep as string
'   Date            Most cast into DateTime
'   IntegerData     Cast into integer
'   Real            Cast into a single
'   LabTest         Cast into a single
'Mo Morris  18/6/2002, standard dates/times are now converted into DateTime fields
'   in RO (Result Only) Data View Tables
'ASH 23/07/2002 - Broken down into several new routines.
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsDataViewTables As ADODB.Recordset

    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM DataViewTables"
    'MLM 15/04/05:
    If mlStudyId > 0 Then
        sSQL = sSQL & " WHERE ClinicalTrialId = " & mlStudyId
    End If
    Set rsDataViewTables = New ADODB.Recordset
    rsDataViewTables.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Do Until rsDataViewTables.EOF
        
        'place progress message in txtProgressMessage
        txtProgressMessage.Text = "Creating " & rsDataViewTables!DataViewName
        txtProgressMessage.Refresh
        'create and populate a data view table of type WA (With Attrtibutes)
        If rsDataViewTables!DataViewType = "WA" Then
            Call CreateWADataViewTable(rsDataViewTables!DataViewName)
                
            Call PopulateWATable _
                    (rsDataViewTables!ClinicalTrialId, rsDataViewTables!CRFPageID, _
                    rsDataViewTables!DataViewName, ConvertFromNull(rsDataViewTables!VisitId, vbLong))
        End If
        
        'create and populate a data view table of type RO (Response Only)
        'also creates and populates RORQGs tables
        If rsDataViewTables!DataViewType = "RO" Then
            'Changed Mo Morris, 29/8/2002, QGroupID is now 0 for NON-RGQ RO Data View Tables, not Null
            If rsDataViewTables!QGroupID = 0 Then
                Call CreateRODataViewTable(rsDataViewTables!DataViewName, _
                        rsDataViewTables!ClinicalTrialId, rsDataViewTables!CRFPageID)
            Else
                Call CreateRQGDataViewTable _
                    (rsDataViewTables!DataViewName, rsDataViewTables!ClinicalTrialId, _
                    rsDataViewTables!CRFPageID, rsDataViewTables!QGroupID)
            End If

            Call PopulateROTable _
                    (rsDataViewTables!ClinicalTrialId, rsDataViewTables!CRFPageID, _
                    rsDataViewTables!DataViewName, ConvertFromNull(rsDataViewTables!VisitId, vbLong), _
                    rsDataViewTables!QGroupID)
        End If
        
        rsDataViewTables.MoveNext
        
    Loop
    rsDataViewTables.Close
    Set rsDataViewTables = Nothing
    
    'place progress message in txtProgressMessage
    txtProgressMessage.Text = "Create Data Views Completed"
    DoEvents
    
    'Set Data View Tables created flag in table DataViewDetails
    sSQL = "UPDATE DataViewDetails Set DataViewTablesExist = " & DataViewState.Created & "," _
        & " DataViewTablesDate = " & SQLStandardNow
    'MLM 15/04/05:
    If mlStudyId > 0 Then
        sSQL = sSQL & " WHERE ClinicalTrialId = " & mlStudyId
    End If
    MacroADODBConnection.Execute sSQL
    
    RefreshDataViewDetails

Exit Sub
ErrHandler:
    'The following error traps handle the situation which occurs when trying to place a DataItem
    'response that no longer exists on a form into a newly created data view that does not have
    'a column for the old dataitem.
    'With an Access database error number 3600 'No value given for one or more required parameters' gets generated.
    'With an Oracle or SQLServer database error number 3604 'Invalid column name' gets generated.
    'In Access -2147217904 has a description of 'No value given for one or more required parameters.'
    If Err.Number - vbObjectError = 3600 Then Resume Next
    If Err.Number - vbObjectError = 3604 Then Resume Next
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CreateAndPopulateDataViews"
End Sub

'---------------------------------------------------------------------
Private Sub cmdCreateDataViews_Click()
'---------------------------------------------------------------------
'Initiates and calls routines to create and populate dataview tables
'---------------------------------------------------------------------
Dim sTitle As String
Dim sPrompt As String
Dim nResponse As Integer

    On Error GoTo ErrLabel
    
    If mbDataViewTablesExist Then
        sTitle = "Data View Creation"
        sPrompt = "Populated Data View Tables already exist." _
            & vbNewLine & "If you continue the existing Data View Tables will" _
            & vbNewLine & "be removed, prior to being re-created and re-populated." _
            & vbNewLine & "The Data Views will be based on the settings in place" _
            & vbNewLine & "when the Data View Table names were created."
        nResponse = DialogWarning(sPrompt, sTitle, True)
        If nResponse = vbCancel Then
            Exit Sub
        End If
    Else
        sTitle = "Data View Creation"
        sPrompt = "The Data Views will be based on the settings in place" _
            & vbNewLine & "when the Data View Table names were created."
        nResponse = DialogWarning(sPrompt, sTitle, True)
        If nResponse = vbCancel Then
            Exit Sub
        End If
        'retrieve the stored settings
        RefreshDataViewDetails
    End If
    
    HourglassOn
    
    'MLM 13/03/06: Moved here to always attempt removal of old data views.
    'remove any Triggers that might exist
    RemoveTriggers
    'remove the Data View Tables, but leave the table DataViewTables intact
    RemoveDataViews
    
    'ASH 26/07/2002 disable controls while CreateAndPopulateDataViews is running
    Call EnableControls(False)
    
    'Create Data views and populate them
    CreateAndPopulateDataViews
    
    'ASH 26/07/2002 Enable controls after CreateAndPopulateDataViews has completed
    Call EnableControls(True)

    HourglassOff
        
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdCreateDataViews_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub cmdHelp_Click()
'---------------------------------------------------------------------
'Help added by Mo Morris 6/4/01
'---------------------------------------------------------------------
            
    'REM 07/12/01 - New Call to MACRO Help
    Call MACROHelp(Me.hWnd, App.Title)
            
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
Dim colmX As MSComctlLib.ColumnHeader

    On Error GoTo ErrLabel
    
    'clear listview of protocols
    lvwTables.ListItems.Clear
    
    'add column headers with widths that are re-calculated in Form_Resize
    'Changed Mo 3/9/2002, column width changes
    Set colmX = lvwTables.ColumnHeaders.Add(, , "Table Name", 2500)
    Set colmX = lvwTables.ColumnHeaders.Add(, , "Type", 500)
    Set colmX = lvwTables.ColumnHeaders.Add(, , "Study", 900)
    Set colmX = lvwTables.ColumnHeaders.Add(, , "Visit", 900)
    Set colmX = lvwTables.ColumnHeaders.Add(, , "Form", 900)
    Set colmX = lvwTables.ColumnHeaders.Add(, , "Group", 900)
 
    'set view type
    lvwTables.View = lvwReport
    'set initial sort to ascending on column 0 (file name)
    lvwTables.SortKey = 0
    lvwTables.SortOrder = lvwAscending
    lvwTables.Sorted = True
    
    optCatCodes.Caption = "As Codes in Text columns. (e.g. F,M or 1,2)"
    optCatValues.Caption = "As Values in Text columns. (e.g. Female, Male)"
    optCatCodesTyped.Caption = "As Codes in Numeric columns when applicable."
    
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Load", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------
' MLM 15/04/05: Changed resizing to accommodate the extra "study selection" frame.
'---------------------------------------------------------------------
Dim lHeight As Long
Dim lWidth As Long

    On Error GoTo ErrLabel
    
    'detect a minimize call and exit the rest of Resize
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If

    'force a minimum hieght to prevent control height being set to negative values
    If Me.Height < 9000 Then Me.Height = 9000 '8500
    
    'force a minimum width to prevent parts of window disapearing
    If Me.Width < 7900 Then Me.Width = 7900 '10700
    
    lWidth = Me.ScaleWidth
    
    Frame1.Width = lWidth - Frame7.Width - 240 '5800
    Frame2.Width = Frame1.Width 'lWidth - 5800
    fraStudy.Width = lWidth - Frame6.Width - 240
    
    Frame6.Left = fraStudy.Left + fraStudy.Width + 100

    Frame7.Left = Frame1.Left + Frame1.Width + 100
    txtProgressMessage.Width = lWidth - 240
    Frame3.Width = lWidth - 240
    lvwTables.Width = Frame3.Width - 240

    
    cmdEdit.Left = Frame3.Width - (cmdEdit.Width + 120)
    Frame4.Width = Frame3.Width - 240
    txtTableName.Width = Frame4.Width - (cmdCancel.Width + cmdChange.Width + 480)
    cmdCancel.Left = txtTableName.Width + 240
    cmdChange.Left = cmdCancel.Left + cmdChange.Width + 120
    cmdExit.Left = lWidth - (cmdExit.Width + 120)
    cmdHelp.Left = cmdExit.Left
    
    lHeight = Me.ScaleHeight
    
    Frame3.Height = lHeight - 5700 '4250
    lvwTables.Height = Frame3.Height - 1500
    cmdEdit.Top = lvwTables.Height + 320
    Frame4.Top = cmdEdit.Top + 370
    Frame5.Top = Frame3.Top + Frame3.Height + 140
    cmdHelp.Top = Frame5.Top - 50
    cmdExit.Top = cmdHelp.Top + 500
    
    cmdCreateViewNames.Top = Frame5.Top + Frame5.Height + 140
    cmdCreateDataViews.Top = cmdCreateViewNames.Top
    cmdCreateTriggers.Top = cmdCreateViewNames.Top
    
    
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Resize", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub CreateTableNames()
'---------------------------------------------------------------------
'ASH 23/07/2002 - Broken down into manageable routines
'Creates dataview table names to be stored in DataViewTables table.
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsUniqueForms As ADODB.Recordset

    On Error GoTo ErrHandler
    
    'Clear the unique names collection
    Set mcolTableNames = Nothing
    Set mcolTableNames = New Collection
    
    'create a recordset from which the names of the individual Data Views will be created
    sSQL = "SELECT DISTINCT StudyVisitCRFPage.ClinicalTrialId, ClinicalTrial.ClinicalTrialName,"
    If optVisitSeparate.Value = True Then
        sSQL = sSQL & " StudyVisitCRFPage.VisitId, StudyVisit.VisitCode,"
    End If
    sSQL = sSQL & " StudyVisitCRFPage.CRFPageID, CRFPage.CRFPageCode" _
        & " FROM StudyVisitCRFPage, ClinicalTrial, CRFPage"
    If optVisitSeparate.Value = True Then
        sSQL = sSQL & ", StudyVisit"
    End If
    sSQL = sSQL & " WHERE StudyVisitCRFPage.ClinicalTrialId = ClinicalTrial.ClinicalTrialId"
    If optVisitSeparate.Value = True Then
        sSQL = sSQL & " AND StudyVisitCRFPage.ClinicalTrialId = StudyVisit.ClinicalTrialId" _
            & " AND StudyVisitCRFPage.VisitId = StudyVisit.VisitId"
    End If
    sSQL = sSQL & " AND StudyVisitCRFPage.ClinicalTrialId = CRFPage.ClinicalTrialId" _
        & " AND StudyVisitCRFPage.CRFPageId = CRFPage.CRFPageId"
    'MLM 15/04/05:
    If mlStudyId > 0 Then
        sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = " & mlStudyId
    End If
    sSQL = sSQL & " ORDER BY ClinicalTrial.ClinicalTrialName,"
    If optVisitSeparate.Value = True Then
        sSQL = sSQL & " StudyVisit.VisitCode,"
    End If
    sSQL = sSQL & " CRFPage.CRFPageCode"

    Set rsUniqueForms = New ADODB.Recordset
    rsUniqueForms.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    'Fix recordcount problem (Q181479)
    If rsUniqueForms.RecordCount < 2 Then
        If rsUniqueForms.RecordCount = 1 Then
            On Error Resume Next
            'Mo Morris 10/9/01 VisitId changed to CRFPageID
            If Not rsUniqueForms!CRFPageID > 0 Then
                Call DialogError("The Database contains no active forms.")
                On Error GoTo ErrHandler
                Exit Sub
            End If
        Else
            Call DialogError("The Database contains no active forms.")
            Exit Sub
        End If
    End If
    
   
    lvwTables.ListItems.Clear
    
   'populate lvwTables with the Data Views to be created
    Do Until rsUniqueForms.EOF
        
        If optVisitSeparate.Value = True Then
            
            Call CreateROTableTypeName(rsUniqueForms!ClinicalTrialId, _
                                    rsUniqueForms!ClinicalTrialname, _
                                    rsUniqueForms!CRFPageCode, _
                                    rsUniqueForms!CRFPageID, _
                                    rsUniqueForms!VisitId, _
                                    rsUniqueForms!VisitCode)
            
            Call CreateRORQGTableTypeName(rsUniqueForms!ClinicalTrialId, _
                                    rsUniqueForms!ClinicalTrialname, _
                                    rsUniqueForms!CRFPageCode, _
                                    rsUniqueForms!CRFPageID, _
                                    rsUniqueForms!VisitId, _
                                    rsUniqueForms!VisitCode)
            Call CreateWATableTypeName(rsUniqueForms!ClinicalTrialId, _
                                    rsUniqueForms!ClinicalTrialname, _
                                    rsUniqueForms!CRFPageCode, _
                                    rsUniqueForms!CRFPageID, _
                                    rsUniqueForms!VisitId, _
                                    rsUniqueForms!VisitCode)
        Else
            Call CreateROTableTypeName(rsUniqueForms!ClinicalTrialId, _
                                    rsUniqueForms!ClinicalTrialname, _
                                    rsUniqueForms!CRFPageCode, _
                                    rsUniqueForms!CRFPageID)
            
            Call CreateRORQGTableTypeName(rsUniqueForms!ClinicalTrialId, _
                                    rsUniqueForms!ClinicalTrialname, _
                                    rsUniqueForms!CRFPageCode, _
                                    rsUniqueForms!CRFPageID)
                                    
            Call CreateWATableTypeName(rsUniqueForms!ClinicalTrialId, _
                                    rsUniqueForms!ClinicalTrialname, _
                                    rsUniqueForms!CRFPageCode, _
                                    rsUniqueForms!CRFPageID)

        End If
        
        rsUniqueForms.MoveNext
    
    Loop 'on rsUniqueForms
    rsUniqueForms.Close
    Set rsUniqueForms = Nothing
    
    'Set Data View Names created flag in table DataViewDetails
    sSQL = "UPDATE DataViewDetails Set DataViewNamesExist = " & DataViewState.Created & "," _
        & " DataViewNamesDate = " & SQLStandardNow
    'MLM 15/04/05:
    If mlStudyId > 0 Then
        sSQL = sSQL & " WHERE ClinicalTrialId = " & mlStudyId
    End If
    MacroADODBConnection.Execute sSQL
    
    StoreDataViewOptions
    
    RefreshDataViewDetails
    
    'ATO 4/09/2002 Resize listview to text length
    Call lvw_SetAllColWidths(lvwTables, LVSCW_AUTOSIZE_USEHEADER)
        
    'ATO 4/09/2002
    Call ShowColumnHeaders(mbDataViewSeparateVisits, False)

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CreateTableNames"
End Sub

'----------------------------------------------------------------------------------------
Private Function CreateTableName(ByVal sTrialName As String, _
                            ByVal sFormCode As String, _
                            ByVal sTableType As String, _
                            Optional ByVal sVisitCode As String = "", _
                            Optional ByVal sGroupCode As String = "") As String
'----------------------------------------------------------------------------------------
''This function creates a table name from the Trial/Visit/Form/Group codes
'Checks the name length for not being greater than max characters for database.
'Truncates name lengths greater than this and adds a numeric suffix
'until a unique table name is created.
'The Global collection gColTableNames is used to check for uniqueness.
'ASH 23/07/2002 - Modified to take account of RQGs names
'-----------------------------------------------------------------------------------------
Dim sTableName As String
Dim nSuffix As Integer
Dim sSuffixedName As String
Dim bNameAddedToCollection As Boolean

    On Error GoTo ErrHandler
        
    If mbDataViewRORequired And sTableType = "RO" Then
        sTableName = "R" & "_" & sTrialName
    Else
        sTableName = "W" & "_" & sTrialName
    End If
        
    If sVisitCode <> "" Then
        sTableName = sTableName & "_" & sVisitCode
    End If

    sTableName = sTableName & "_" & sFormCode
        
    If sGroupCode <> "" Then
        sTableName = sTableName & "_" & sGroupCode
    End If
    
    If Len(sTableName) > mnTableNameLength Then
    'uses 2 spaces at end for number suffix
        sTableName = Mid(sTableName, 1, (mnTableNameLength - 2))
        bNameAddedToCollection = False
        nSuffix = 1
        On Error Resume Next
        Do
            sSuffixedName = sTableName & CStr(nSuffix)
            mcolTableNames.Add sSuffixedName, sSuffixedName
            If Err.Number = 0 Then
                bNameAddedToCollection = True
            Else
                nSuffix = nSuffix + 1
                Err.Clear
            End If
        Loop Until bNameAddedToCollection
        sTableName = sSuffixedName
    Else
        mcolTableNames.Add sTableName, sTableName
    End If
    
    CreateTableName = sTableName

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CreateTableName"
End Function

'---------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------------------

    Call ExitMACRO
    Call MACROEnd

End Sub

'---------------------------------------------------------------------
Private Sub lvwTables_DblClick()
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    If mbDataViewNameEditsEnabled Then
        cmdEdit_Click
    End If

Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "lvwTables_DblClick", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub CreateDataViewTables()
'---------------------------------------------------------------------
'Creates the main DataviewTable which stores all the dataview table names.
'ASH 23/07/2002 - Added QGroupID and QGroupCode
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        sSQL = "CREATE TABLE DataViewTables (DataViewName TEXT(28)," _
            & " DataViewType TEXT(2),  ClinicalTrialName TEXT(15)," _
            & " VisitCode TEXT(15), CRFPageCode TEXT(15)," _
            & " QGroupCode TEXT(15)," _
            & " ClinicalTrialId INTEGER, VisitId INTEGER, CRFPageId INTEGER," _
            & " QGroupID INTEGER," _
            & " CONSTRAINT PrimaryKey PRIMARY KEY" _
            & " (DataViewName))"
    Else
        sSQL = "CREATE TABLE DataViewTables (DataViewName VARCHAR(255)," _
            & " DataViewType VARCHAR(2),  ClinicalTrialName VARCHAR(15)," _
            & " VisitCode VARCHAR(15), CRFPageCode VARCHAR(15)," _
            & " QGroupCode VARCHAR(15)," _
            & " ClinicalTrialId INTEGER, VisitId INTEGER, CRFPageId INTEGER," _
            & " QGroupID INTEGER," _
            & " CONSTRAINT PKDataViewTables PRIMARY KEY" _
            & " (DataViewName))"
    End If
    
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CreateDataViewTables"
End Sub

'---------------------------------------------------------------------
Private Sub CreateDataViewDetails()
'---------------------------------------------------------------------
'Mo 17/4/2002   Field OutputCategoryValues added to table DataViewDetails
'               OutputCategoryValues default value is 1 (DataViewOption.Required) optCatValues.Value = True
'               OutputCategoryValues value = 0 (DataViewOption.NotRequired) optCatCodes.Value = True
'Mo 29/11/2004  Bug 2446 - Adding Special Values facilities to CDV
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'RJCW 06/09/2001   Update table DataViewDetails to include DataViewTrigCalc
    If goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        sSQL = "CREATE TABLE DataViewDetails (" _
            & " DataViewNamesExist INTEGER, DataViewNamesDate DECIMAL(16,10)," _
            & " DataViewTablesExist INTEGER, DataViewTablesDate DECIMAL(16,10)," _
            & " DataViewTriggersExist INTEGER,DataViewTriggersDate DECIMAL(16,10)," _
            & " DataViewSeparateVisits INTEGER, DataViewRO INTEGER, DataViewWA INTEGER, OutputCategoryValues INTEGER," _
            & " SpecialValueMissing VARCHAR2(2), SpecialValueUnobtainable VARCHAR2(2), SpecialValueNotApplicable VARCHAR2(2))"
    Else
        sSQL = "CREATE TABLE DataViewDetails (" _
            & " DataViewNamesExist INTEGER, DataViewNamesDate DECIMAL(16,10)," _
            & " DataViewTablesExist INTEGER, DataViewTablesDate DECIMAL(16,10)," _
            & " DataViewTriggersExist INTEGER,DataViewTriggersDate DECIMAL(16,10)," _
            & " DataViewSeparateVisits INTEGER, DataViewRO INTEGER, DataViewWA INTEGER," _
            & " DataViewTrigCalc VARCHAR(1), OutputCategoryValues INTEGER," _
            & " SpecialValueMissing VARCHAR(2), SpecialValueUnobtainable VARCHAR(2), SpecialValueNotApplicable VARCHAR(2))"
    End If
    
    MacroADODBConnection.Execute sSQL
    
    sSQL = "INSERT INTO DataViewDetails (DataViewNamesExist,DataViewNamesDate," _
        & " DataViewTablesExist,DataViewTablesDate,DataViewTriggersExist," _
        & " DataViewTriggersDate,DataViewSeparateVisits,DataViewRO,DataViewWA,OutputCategoryValues," _
        & " SpecialValueMissing,SpecialValueUnobtainable,SpecialValueNotApplicable)" _
        & " VALUES(" & DataViewState.NotCreated & ",0," _
        & DataViewState.NotCreated & ",0," _
        & DataViewState.NotCreated & ",0," _
        & DataViewOption.NotRequired & "," & DataViewOption.Required & "," _
        & DataViewOption.NotRequired & "," & DataViewOption.Required & ",null,null,null)"
        
    MacroADODBConnection.Execute sSQL
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CreateDataViewDetails"
End Sub

'---------------------------------------------------------------------
Private Sub MaintainDataViewDetails()
'---------------------------------------------------------------------
' MLM 15/04/05: Copied from 2.1. Ensure that the ClinicalTrialIds in DataViewDetails
' correspond to those in ClinicalTrial. Called from InitialiseMe
' MLM 13/09/05: bug 2632: Set default values of null for "special values"
' MLM 13/03/06: Really set them to null this time.
'---------------------------------------------------------------------

Dim sSQL As String
Dim dicStudies As Scripting.Dictionary
Dim rsDataViewDetails As ADODB.Recordset
Dim rsStudy As ADODB.Recordset
Dim lNamesExist As Long
Dim dblNamesDate As Double
Dim lTablesExist As Long
Dim dblTablesDate As Double
Dim lTriggersExist As Long
Dim dblTriggersDate As Double
Dim lSeparateVisits As Long
Dim lRO As Long
Dim lWA As Long
Dim lOutputCategoryValues As Long
Dim vSpecialValueMissing As Variant
Dim vSpecialValueUnobtainable As Variant
Dim vSpecialValueNotApplicable As Variant

    On Error GoTo ErrHandler
    
    'default to "response only" data views (other parameters all have defaults of 0)
    lRO = 1
    ' MLM 13/09/05: bug 2632: also need to set these default values:
    vSpecialValueMissing = Null
    vSpecialValueUnobtainable = Null
    vSpecialValueNotApplicable = Null
    
    'start by building a collection of the clinicaltrialid we already have CDV details for
    'meanwhile, if we notice a null clinicaltrialid, it means we're upgrading:
    'remember these details so that they can be applied to existing studies.
    Set rsDataViewDetails = New ADODB.Recordset
    Set dicStudies = New Scripting.Dictionary
    With rsDataViewDetails
        .Open "SELECT * FROM DataViewDetails", MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
        Do Until .EOF
            If IsNull(.Fields("ClinicalTrialId").Value) Then
                lNamesExist = .Fields("DataViewNamesExist").Value
                dblNamesDate = .Fields("DataViewNamesDate").Value
                lTablesExist = .Fields("DataViewTablesExist").Value
                dblTablesDate = .Fields("DataViewTablesDate").Value
                lTriggersExist = .Fields("DataViewTriggersExist").Value
                dblTriggersDate = .Fields("DataViewTriggersDate").Value
                lSeparateVisits = .Fields("DataViewSeparateVisits").Value
                lRO = .Fields("DataViewRO").Value
                lWA = .Fields("DataViewWA").Value
                lOutputCategoryValues = .Fields("OutputCategoryValues").Value
                vSpecialValueMissing = .Fields("SpecialValueMissing").Value
                vSpecialValueUnobtainable = .Fields("SpecialValueUnobtainable").Value
                vSpecialValueNotApplicable = .Fields("SpecialValueNotApplicable").Value
                
                dicStudies.Add "", ""
            Else
                dicStudies.Add .Fields("ClinicalTrialId").Value, .Fields("ClinicalTrialId").Value
            End If
            .MoveNext
        Loop
        .Close
    End With
    
    'now loop through the studies in this database, and decide what action to take in DataViewDetails for each one.
    Set rsStudy = New ADODB.Recordset
    With rsStudy
        .Open "SELECT ClinicalTrialId FROM Clinicaltrial WHERE ClinicalTrialId > 0", MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
        Do Until .EOF
            If dicStudies.Exists(.Fields(0).Value) Then
                'this study already represented in DataViewDetails
            ElseIf dicStudies.Exists("") Then
                sSQL = "UPDATE DataViewDetails SET ClinicalTrialId = " & .Fields(0).Value & _
                    " WHERE ClinicalTrialId IS NULL"
                MacroADODBConnection.Execute sSQL
                dicStudies.Remove ""
            Else
                sSQL = "INSERT INTO DataViewDetails (DataViewNamesExist, DataViewNamesDate, DataViewTablesExist, DataViewTablesDate, DataViewTriggersExist, DataViewTriggersDate, " & _
                    "DataViewSeparateVisits, DataViewRO, DataViewWA, OutputCategoryValues, SpecialValueMissing, SpecialValueUnobtainable, SpecialValueNotApplicable, ClinicalTrialId)" & _
                    " VALUES(" & lNamesExist & ", " & _
                    ConvertLocalNumToStandard(CStr(dblNamesDate)) & ", " & _
                    lTablesExist & ", " & _
                    ConvertLocalNumToStandard(CStr(dblTablesDate)) & ", " & _
                    lTriggersExist & ", " & _
                    ConvertLocalNumToStandard(CStr(dblTriggersDate)) & ", " & _
                    lSeparateVisits & ", " & _
                    lRO & ", " & _
                    lWA & ", " & _
                    lOutputCategoryValues & ", " & _
                    SQL_ValueToStringValue(vSpecialValueMissing) & ", " & _
                    SQL_ValueToStringValue(vSpecialValueUnobtainable) & ", " & _
                    SQL_ValueToStringValue(vSpecialValueNotApplicable) & ", " & _
                    .Fields(0).Value & ")"
                    
                MacroADODBConnection.Execute sSQL
            End If
            .MoveNext
        Loop
        .Close
    End With
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.MaintainDataViewDetails"
End Sub

'---------------------------------------------------------------------
Private Sub RefreshDataViewDetails()
'---------------------------------------------------------------------
'Refreshes the DataViewDetails table
'---------------------------------------------------------------------
'   Mo  5/2/2003    SR 4341 Fixed
'Mo Morris  24/11/2004 - Bug 2413 Category codes in Numeric Fields option
' MLM 15/04/05: Study-specific data views: This is now has aspects of the old 2.1 version mixed in.
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim rsDate As ADODB.Recordset

    On Error GoTo ErrHandler

    cmdCreateTriggers.Enabled = False

    sSQL = "SELECT MIN(DataViewNamesExist), " & _
        "MIN(DataViewTablesExist), " & _
        "MIN(DataViewTriggersExist), " & _
        "MIN(DataViewSeparateVisits), " & _
        "MIN(DataViewRO), " & _
        "MIN(DataViewWA), " & _
        "MIN(OutputCategoryValues), " & _
        "MIN(SpecialValueMissing), " & _
        "MIN(SpecialValueUnobtainable), " & _
        "MIN(SpecialValueNotApplicable) " & _
        "FROM DataViewDetails"
    If mlStudyId > 0 Then
        sSQL = sSQL & " WHERE ClinicalTrialId = " & mlStudyId
    End If
    'sSQL = "SELECT * FROM DataViewDetails"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    'prepare a 2nd query that will discover what we should display for the various created dates
    sSQL = "SELECT DISTINCT ? FROM DataViewDetails"
    If mlStudyId > 0 Then
        sSQL = sSQL & " WHERE ClinicalTrialId = " & mlStudyId
    End If
    Set rsDate = New ADODB.Recordset

    
    If rsTemp.Fields(0).Value = DataViewState.Created Then
        mbDataViewTableNamesExist = True
        mbDataViewNameEditsEnabled = True
        cmdCreateDataViews.Enabled = True
        rsDate.Open Replace(sSQL, "?", "DataViewNamesDate"), MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        If rsDate.RecordCount = 1 Then
            txtCreateViewNamesDate.Text = Format(rsDate!DataViewNamesDate, "yyyy/mm/dd")
        Else
            txtCreateViewNamesDate.Text = "Various"
        End If
        rsDate.Close
    Else
        mbDataViewTableNamesExist = False
        mbDataViewNameEditsEnabled = False
        txtCreateViewNamesDate.Text = "Not Run"
        cmdCreateDataViews.Enabled = False
    End If
    
    If rsTemp.Fields(1).Value = DataViewState.Created Then
        mbDataViewTablesExist = True
        rsDate.Open Replace(sSQL, "?", "DataViewTablesDate"), MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        If rsDate.RecordCount = 1 Then
            txtCreateDataViewsDate.Text = Format(rsDate!DataViewTablesDate, "yyyy/mm/dd")
        Else
            txtCreateDataViewsDate.Text = "Various"
        End If
        rsDate.Close
        cmdCreateDataViews.Enabled = True
        'Only enable the creation of triggers for Oracle databases
        '   RJCW 05/09/2001   Enable SQL Server Triggers
        If goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Or goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
            cmdCreateTriggers.Enabled = True
        End If
    Else
        mbDataViewTablesExist = False
        txtCreateDataViewsDate.Text = "Not Run"
    End If
    
    If rsTemp.Fields(2).Value = DataViewState.Created Then
        mbDataViewTriggersExist = True
        rsDate.Open Replace(sSQL, "?", "DataViewTriggersDate"), MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        If rsDate.RecordCount = 1 Then
            txtCreateTriggersDate.Text = Format(rsDate!DataViewTriggersDate, "yyyy/mm/dd")
        Else
            txtCreateTriggersDate.Text = "Various"
        End If
        rsDate.Close
    Else
        mbDataViewTriggersExist = False
        txtCreateTriggersDate.Text = "Not Run"
    End If
    
    If rsTemp.Fields(3).Value = DataViewOption.Required Then
        mbDataViewSeparateVisits = True
        optVisitSeparate.Value = True
    Else
        mbDataViewSeparateVisits = False
        optVisitTogether.Value = True
    End If
        
    If rsTemp.Fields(4).Value = DataViewOption.Required Then
        mbDataViewRORequired = True
        chkResponseValueOnly.Value = vbChecked
    Else
        mbDataViewRORequired = False
        chkResponseValueOnly.Value = vbUnchecked
    End If
    
    If rsTemp.Fields(5).Value = DataViewOption.Required Then
        mbDataViewWARequired = True
        chkResponseValuePlus.Value = vbChecked
    Else
        mbDataViewWARequired = False
        chkResponseValuePlus.Value = vbUnchecked
    End If
    
    'Mo 24/11/2004 - Bug 2413, using new enumeration for Create Data Views
    Select Case rsTemp.Fields(6).Value
    Case DataViewCategoryOptions.Codes
        mnOutputCategoryValues = DataViewCategoryOptions.Codes
        optCatCodes.Value = True
    Case DataViewCategoryOptions.Values
        mnOutputCategoryValues = DataViewCategoryOptions.Values
        optCatValues.Value = True
    Case DataViewCategoryOptions.TypedCodes
        mnOutputCategoryValues = DataViewCategoryOptions.TypedCodes
        optCatCodesTyped.Value = True
    End Select
    
    txtMissing.Text = RemoveNull(rsTemp.Fields(7).Value)
    txtUnobtainable.Text = RemoveNull(rsTemp.Fields(8).Value)
    txtNotApplicable.Text = RemoveNull(rsTemp.Fields(9).Value)
    
    rsTemp.Close
    
    'MLM 15/04/05: Added. Having populated the UI with the dataview settings,
    ' warn if they aren't representative of all dataviews currently created.
    If mlStudyId = 0 Then
        sSQL = "SELECT DISTINCT DataViewSeparateVisits, DataViewRO, DataViewWA, OutputCategoryValues, " & _
            "SpecialValueMissing, SpecialValueUnobtainable, SpecialValueNotApplicable " & _
            "FROM DataViewDetails WHERE DataViewNamesExist = " & DataViewState.Created
        rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        If rsTemp.RecordCount > 1 Then
            MsgBox """All studies"" is selected and data views exist using different options for different studies." & vbCrLf & _
                "In this mode, the displayed options will replace the current options for all studies if you recreate the data views." & vbCrLf & _
                "To view the options used by existing data views or to recreate the data views using these options, select a specific study.", , _
                "Create Data Views"
        End If
        rsTemp.Close
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.RefreshDataViewDetails"
End Sub

'---------------------------------------------------------------------
Private Sub RemoveDataViews()
'---------------------------------------------------------------------
'Removes dataview names from DataViewTables when new dataviews names are to be created
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
    
    HourglassOn

    sSQL = "SELECT DataViewName FROM DataViewTables"
    'MLM 15/04/05:
    If mlStudyId > 0 Then
        sSQL = sSQL & " WHERE ClinicalTrialId = " & mlStudyId
    End If

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    'the resume next is to capture Drop Errors if the table does not exist
    On Error GoTo 0
    On Error Resume Next
    
    Do While Not rsTemp.EOF
        sSQL = "DROP TABLE " & rsTemp!DataViewName
        MacroADODBConnection.Execute sSQL
        rsTemp.MoveNext
    Loop
    
    On Error GoTo ErrHandler
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    'Update Data View Tables created flag in table DataViewDetails
    sSQL = "UPDATE DataViewDetails Set DataViewTablesExist = " & DataViewState.NotCreated & "," _
        & " DataViewTablesDate = 0"
    'MLM 15/04/05:
    If mlStudyId > 0 Then
        sSQL = sSQL & " WHERE ClinicalTrialId = " & mlStudyId
    End If

    MacroADODBConnection.Execute sSQL
    
    RefreshDataViewDetails
    
    HourglassOff

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.RemoveDataViews"
End Sub

'---------------------------------------------------------------------
Private Sub RemoveDataViewTables()
'---------------------------------------------------------------------
'Removes DataViewTables when new dataviews are to be created
'--------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    'MLM 15/04/05: Don't delete the whole DataViewTables table, as it might contain data for studies we want.
    sSQL = "DELETE FROM DataViewTables"
    If mlStudyId > 0 Then
        sSQL = sSQL & " WHERE ClinicalTrialId = " & mlStudyId
    End If

    'remove DataViewTables and its contents and then re-ceate it
    'sSQL = "DROP TABLE DataViewTables"
    MacroADODBConnection.Execute sSQL
    
    'CreateDataViewTables
    
    'Update Data View Names created flag in table DataViewDetails
    sSQL = "UPDATE DataViewDetails Set DataViewNamesExist = " & DataViewState.NotCreated & "," _
        & " DataViewNamesDate = 0"
    'MLM 15/04/05:
    If mlStudyId > 0 Then
        sSQL = sSQL & " WHERE ClinicalTrialId = " & mlStudyId
    End If
    MacroADODBConnection.Execute sSQL
    
    'Set Data View Options to those currently being displayed
    StoreDataViewOptions
    
    RefreshDataViewDetails

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.RemoveDataViewTables"
End Sub

'---------------------------------------------------------------------
Private Sub RemoveTriggers()
'---------------------------------------------------------------------
'Removes triggers. Not applicable to Access database.
'---------------------------------------------------------------------
Dim sSQL As String

    On Error Resume Next

    MacroADODBConnection.Execute "DROP TRIGGER MACRO_DVROTRIGGER"
    
    MacroADODBConnection.Execute "DROP TRIGGER MACRO_DVWATRIGGER"
    
    'Update Data View Names created flag in table DataViewDetails
    sSQL = "UPDATE DataViewDetails Set DataViewTriggersExist = " & DataViewState.NotCreated & "," _
        & " DataViewTriggersDate = 0"
    'MLM 15/04/05:
    If mlStudyId > 0 Then
        sSQL = sSQL & " WHERE ClinicalTrialId = " & mlStudyId
    End If
    MacroADODBConnection.Execute sSQL
        
    RefreshDataViewDetails

End Sub

'---------------------------------------------------------------------
Private Sub PopulateStudyList()
'---------------------------------------------------------------------
' MLM 15/04/05: Copied from 2.1. Fill cboStudies with list of studies from the db.
'---------------------------------------------------------------------

Dim sSQL As String
Dim rsStudies As ADODB.Recordset

    sSQL = "SELECT ClinicalTrialId, ClinicalTrialName FROM ClinicalTrial WHERE ClinicalTrialId > 0 ORDER BY ClinicalTrialName"
    Set rsStudies = New ADODB.Recordset
    With rsStudies
        .Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
        'don't let the user select a study if there aren't any
        If .EOF Then
            optSingleStudy.Enabled = False
            Exit Sub
        End If
        Do Until .EOF
            cboStudies.AddItem .Fields("ClinicalTrialName").Value
            cboStudies.ItemData(cboStudies.ListCount - 1) = .Fields("ClinicalTrialId").Value
            .MoveNext
        Loop
        .Close
    End With
    'select the first study in the list
    cboStudies.ListIndex = 0

End Sub

'---------------------------------------------------------------------
Private Sub PopulateListView()
'---------------------------------------------------------------------
'Populates the GUI with selected options.
'ATO 4/9/2002 added group name/code to listview during population
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim itmX As MSComctlLib.ListItem
Dim bVisitCodeNotAssesed As Boolean
Dim bVisitCodeExists As Boolean

    On Error GoTo ErrHandler

    bVisitCodeNotAssesed = True
    
    'clear listview of protocols
    lvwTables.ListItems.Clear
    
    'MLM 14/07/03: this moved inside PopulateListView from InitialiseMe
    'If Not gbDataViewTableNamesExist Then Exit Sub ???
    
    'Clear the unique names collection
    Set mcolTableNames = Nothing
    Set mcolTableNames = New Collection

    sSQL = "SELECT * FROM DataViewTables"
    'MLM 15/04/05:
    If mlStudyId > 0 Then
        sSQL = sSQL & " WHERE ClinicalTrialId = " & mlStudyId
    End If

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Do While Not rsTemp.EOF
'        If bVisitCodeNotAssesed Then
'            If IsNull(rsTemp!VisitCode) Then
'                bVisitCodeExists = False
'
'            Else
'                bVisitCodeExists = True
'
'            End If
'            bVisitCodeNotAssesed = False
'        End If
        
        Set itmX = lvwTables.ListItems.Add(, , rsTemp!DataViewName)
        'add the DataViewName to the collection gcolTableNames, which is used for validating uniqueness
        mcolTableNames.Add rsTemp!DataViewName, rsTemp!DataViewName
        itmX.SubItems(1) = rsTemp!DataViewType
        itmX.SubItems(2) = rsTemp!ClinicalTrialname
        'If bVisitCodeExists Then
        If Not IsNull(rsTemp!VisitCode) Then
            itmX.SubItems(3) = RemoveNull(rsTemp!VisitCode)
            bVisitCodeExists = True
        End If
        itmX.SubItems(4) = rsTemp!CRFPageCode
        'ATO 4/9/2002 added group name/code
        itmX.SubItems(5) = RemoveNull(rsTemp!QGroupCode)
        rsTemp.MoveNext
    Loop
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    'ATO 4/09/2002 Resize listview to text length
    Call lvw_SetAllColWidths(lvwTables, LVSCW_AUTOSIZE_USEHEADER)
    
    'ATO 4/09/2002
    Call ShowColumnHeaders(bVisitCodeExists, False)
    
'    ???
'    'MLM 14/07/03: this moved inside PopulateListView from InitialiseMe
'    If Not gbDataViewTablesExist Then
'        lvwTables.Enabled = True
'    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.PopulateListView"
End Sub

'---------------------------------------------------------------------
Private Sub lvwTables_ItemClick(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If mbDataViewNameEditsEnabled Then
        cmdEdit.Enabled = True
    End If

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lvwTables_ItemClick")
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
Private Sub CreateWADataViewTable(ByVal sTableName As String)
'---------------------------------------------------------------------
'Creates Response Value Plus attributes table.
'ASH 23/07/2002 - Added OwnerQGroupID and RepeatNumber
'Mo 24/10/2006 Bug 2824 - Corrected for Oracle databases.
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Access
        sSQL = "CREATE TABLE " & sTableName & "(ClinicalTrialId INTEGER," _
            & " Site TEXT(8)," _
            & " PersonId INTEGER," _
            & " VisitId INTEGER," _
            & " VisitCycleNumber SMALLINT," _
            & " CRFPageId INTEGER," _
            & " CRFPageCycleNumber SMALLINT," _
            & " OwnerQGroupID SMALLINT," _
            & " DataItemCode TEXT(15)," _
            & " RepeatNumber SMALLINT," _
            & " ResponseTimeStamp TEXT(19)," _
            & " DataType SMALLINT," _
            & " ResponseValue TEXT(255)," _
            & " Units TEXT(15)," _
            & " ValueCode TEXT(15)," _
            & " ResponseStatus SMALLINT," _
            & " LabResult TEXT(1)," _
            & " CTCGrade SMALLINT," _
            & " CONSTRAINT PrimaryKey PRIMARY KEY" _
            & " (ClinicalTrialId,Site,PersonId,VisitId,VisitCycleNumber,CRFPageId,CRFPageCycleNumber,OwnerQGroupID,DataItemCode,RepeatNumber))"
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sSQL = "CREATE TABLE " & sTableName & "(ClinicalTrialId INTEGER," _
            & " Site VARCHAR(8)," _
            & " PersonId INTEGER," _
            & " VisitId INTEGER," _
            & " VisitCycleNumber INTEGER," _
            & " CRFPageId INTEGER," _
            & " CRFPageCycleNumber INTEGER," _
            & " OwnerQGroupID INTEGER," _
            & " DataItemCode VARCHAR(15)," _
            & " RepeatNumber INTEGER," _
            & " ResponseTimeStamp VARCHAR(19)," _
            & " DataType INTEGER," _
            & " ResponseValue VARCHAR(255)," _
            & " Units VARCHAR(15)," _
            & " ValueCode VARCHAR(15)," _
            & " ResponseStatus INTEGER," _
            & " LabResult VARCHAR(1)," _
            & " CTCGrade INTEGER," _
            & " CONSTRAINT PK" & sTableName & " PRIMARY KEY" _
            & " (ClinicalTrialId,Site,PersonId,VisitId,VisitCycleNumber,CRFPageId,CRFPageCycleNumber,OwnerQGroupID,DataItemCode,RepeatNumber))"
    Case MACRODatabaseType.Oracle80
        'Mo 24/10/2006 Bug 2824, newly added Oracle option, Oracle used to use the SQLServer option
        sSQL = "CREATE TABLE " & sTableName & "(ClinicalTrialId NUMBER(11)," _
            & " Site VARCHAR2(8)," _
            & " PersonId NUMBER(11)," _
            & " VisitId NUMBER(11)," _
            & " VisitCycleNumber NUMBER(11)," _
            & " CRFPageId NUMBER(11)," _
            & " CRFPageCycleNumber NUMBER(11)," _
            & " OwnerQGroupID NUMBER(11)," _
            & " DataItemCode VARCHAR2(15)," _
            & " RepeatNumber NUMBER(11)," _
            & " ResponseTimeStamp VARCHAR2(19)," _
            & " DataType NUMBER(11)," _
            & " ResponseValue VARCHAR2(255)," _
            & " Units VARCHAR2(15)," _
            & " ValueCode VARCHAR2(15)," _
            & " ResponseStatus NUMBER(11)," _
            & " LabResult VARCHAR2(1)," _
            & " CTCGrade NUMBER(11)," _
            & " CONSTRAINT PK" & sTableName & " PRIMARY KEY" _
            & " (ClinicalTrialId,Site,PersonId,VisitId,VisitCycleNumber,CRFPageId,CRFPageCycleNumber,OwnerQGroupID,DataItemCode,RepeatNumber))"
    End Select
    
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CreateWADataViewTable"
End Sub

'---------------------------------------------------------------------
Private Sub CreateRODataViewTable(ByVal sTableName As String, _
                                ByVal lClinicalTrialId As Long, _
                                ByVal lCRFPageId As Long)
'---------------------------------------------------------------------
'Creates a response Value Only dataview table
'
'Mo Morris 16/4/2002
'   Changes made so that the question columns in the dataview tables
'   are typed to reflect their content:-
'
'   Question Type   in Access DB    in SQL Server DB   in Oracle DB
'   Text            TEXT(255)       VARCHAR(255)        VARCHAR2(255)
'   Category        TEXT(255)       VARCHAR(255)        VARCHAR2(255)
'   Multimedia      TEXT(255)       VARCHAR(255)        VARCHAR2(255)
'   Date            DATETIME        DATETIME            DATE
'   IntegerData     INTEGER         INTEGER             NUMBER(11)
'   Real            DOUBLE          NUMERIC(16,10)      NUMBER(16,10)
'   LabTest         DOUBLE          NUMERIC(16,10)      NUMBER(16,10)
'Mo Morris  18/6/2002
'   Standard dates/times are now converted to DateTime fields
'
'ASH 23/07/2002 - Now calls DataBaseSpecificColumnFormatting
'Mo Morris  24/11/2004 - Bug 2413 Category codes in Numeric Fields option
'Mo 24/10/2006 Bug 2824 - Corrected for Oracle databases.
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsFormDataItemCodes As ADODB.Recordset
Dim sSQLDataItemCodes As String

    On Error GoTo ErrHandler

    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Access
        sSQL = "CREATE TABLE " & sTableName & "(ClinicalTrialId INTEGER," _
            & " Site TEXT(8)," _
            & " PersonId INTEGER," _
            & " VisitId INTEGER," _
            & " VisitCycleNumber SMALLINT," _
            & " CRFPageId INTEGER," _
            & " CRFPageCycleNumber SMALLINT,"
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sSQL = "CREATE TABLE " & sTableName & "(ClinicalTrialId INTEGER," _
            & " Site VARCHAR(8)," _
            & " PersonId INTEGER," _
            & " VisitId INTEGER," _
            & " VisitCycleNumber INTEGER," _
            & " CRFPageId INTEGER," _
            & " CRFPageCycleNumber INTEGER,"
    Case MACRODatabaseType.Oracle80
        'Mo 24/10/2006 Bug 2824, newly added Oracle option, Oracle used to use the SQLServer option
        sSQL = "CREATE TABLE " & sTableName & "(ClinicalTrialId NUMBER(11)," _
            & " Site VARCHAR2(8)," _
            & " PersonId NUMBER(11)," _
            & " VisitId NUMBER(11)," _
            & " VisitCycleNumber NUMBER(11)," _
            & " CRFPageId NUMBER(11)," _
            & " CRFPageCycleNumber NUMBER(11),"
    End Select
    
    'create a recordset of the questions within form lCRFPageId
    'Changed Mo 18/6/2002, select DataItemFormat as well
    'Mo 24/10/2006 Bug 2824, DataItemCase added to following SQL
    sSQLDataItemCodes = "SELECT DataItem.DataItemId, DataItem.DataItemCode, DataItem.DataType, DataItem.DataItemFormat, DataItem.DataItemCase" _
        & " FROM CRFElement, DataItem" _
        & " WHERE CRFElement.ClinicalTrialId = " & lClinicalTrialId _
        & " AND CRFElement.CRFPageId = " & lCRFPageId _
        & " AND CRFElement.ClinicalTrialId = DataItem.ClinicalTrialId" _
        & " AND CRFElement.DataItemId = DataItem.DataItemId" _
        & " AND CRFElement.OwnerQGroupID = 0" _
        & " ORDER BY CRFElement.FieldOrder"
    Set rsFormDataItemCodes = New ADODB.Recordset
    rsFormDataItemCodes.Open sSQLDataItemCodes, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        
    Do Until rsFormDataItemCodes.EOF
      
        'Mo 24/10/2006 Bug 2824, DataItemCase added to call to DataBaseSpecificColumnFormatting
        sSQL = sSQL & DataBaseSpecificColumnFormatting(lClinicalTrialId, _
                rsFormDataItemCodes!DataItemId, rsFormDataItemCodes!DataItemCode, CInt(RemoveNull(rsFormDataItemCodes!DataItemCase)), _
                ConvertFromNull(rsFormDataItemCodes!DataItemFormat, vbString), _
                ConvertFromNull(rsFormDataItemCodes!DataType, vbInteger))
        rsFormDataItemCodes.MoveNext
    
    Loop
    
    rsFormDataItemCodes.Close
    Set rsFormDataItemCodes = Nothing
        
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY"
    Else
        sSQL = sSQL & " CONSTRAINT PK" & sTableName & " PRIMARY KEY"
    End If
    
    sSQL = sSQL & " (ClinicalTrialId,Site,PersonId,VisitId,VisitCycleNumber,CRFPageId,CRFPageCycleNumber))"

    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CreateRODataViewTable"
End Sub

'---------------------------------------------------------------------
Private Sub optAllStudies_Click()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    cboStudies.Enabled = False
    mlStudyId = 0
    
    'refresh form to show all studies
    'retrieve status details from table DataViewDetails
    RefreshDataViewDetails
    
    'if Data View Names exist then populate lvwTables with them
    PopulateListView

End Sub

'---------------------------------------------------------------------
Private Sub optSingleStudy_Click()
'---------------------------------------------------------------------

    cboStudies.Enabled = True
    
    'filter everything by the selected study:
    cboStudies_Click

End Sub

'---------------------------------------------------------------------
Private Sub optCatCodes_Click()
'---------------------------------------------------------------------
'Mo Morris  24/11/2004 - Bug 2413 Category codes in Numeric Fields option
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    If optCatCodes.Value = True Then
        mnOutputCategoryValues = DataViewCategoryOptions.Codes
    End If
    
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "optCatCodes_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub optCatCodesTyped_Click()
'---------------------------------------------------------------------
'Mo Morris  24/11/2004 - Bug 2413 Category codes in Numeric Fields option
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    If optCatCodesTyped.Value = True Then
        mnOutputCategoryValues = DataViewCategoryOptions.TypedCodes
    End If
    
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "optCatCodesTyped_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub optCatValues_Click()
'---------------------------------------------------------------------
'Mo Morris  24/11/2004 - Bug 2413 Category codes in Numeric Fields option
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    If optCatValues.Value = True Then
        mnOutputCategoryValues = DataViewCategoryOptions.Values
    End If
    
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "optCatValues_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub optVisitSeparate_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
     
    If optVisitSeparate.Value = True Then
        mbDataViewSeparateVisits = True
    Else
        mbDataViewSeparateVisits = False
    End If
        
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "optVisitSeparate_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub optVisitTogether_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    If optVisitTogether.Value = True Then
        mbDataViewSeparateVisits = False
    Else
        mbDataViewSeparateVisits = True
    End If
    
Exit Sub
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "optVisitTogether_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub StoreDataViewOptions()
'---------------------------------------------------------------------
'Mo Morris  24/11/2004 - Bug 2413 Category codes in Numeric Fields option
'---------------------------------------------------------------------
Dim nDataViewSeparateVisits As Integer
Dim nDataViewRO As Integer
Dim nDataViewWA  As Integer
Dim sSQL As String
Dim sSVMissing As String
Dim sSVUnobtainable As String
Dim sSVNotApplicable As String

    On Error GoTo ErrHandler

    If mbDataViewSeparateVisits Then
        nDataViewSeparateVisits = DataViewOption.Required
    Else
        nDataViewSeparateVisits = DataViewOption.NotRequired
    End If
    
    If mbDataViewRORequired Then
        nDataViewRO = DataViewOption.Required
    Else
        nDataViewRO = DataViewOption.NotRequired
    End If
    
    If mbDataViewWARequired Then
        nDataViewWA = DataViewOption.Required
    Else
        nDataViewWA = DataViewOption.NotRequired
    End If
    
    If txtMissing.Text = "" Then
        sSVMissing = "null"
    Else
        sSVMissing = txtMissing.Text
    End If
    
    If txtUnobtainable.Text = "" Then
        sSVUnobtainable = "null"
    Else
        sSVUnobtainable = txtUnobtainable.Text
    End If
    
    If txtNotApplicable.Text = "" Then
        sSVNotApplicable = "null"
    Else
        sSVNotApplicable = txtNotApplicable.Text
    End If
      
    'Changed mo 17/4/2002, OutputCategoryValues now updated
    sSQL = "UPDATE DataViewDetails Set DataViewSeparateVisits = " & nDataViewSeparateVisits & "," _
        & " DataViewRO = " & nDataViewRO & "," _
        & " DataViewWA = " & nDataViewWA & "," _
        & " OutputCategoryValues = " & mnOutputCategoryValues & "," _
        & " SpecialValueMissing = " & sSVMissing & "," _
        & " SpecialValueUnobtainable = " & sSVUnobtainable & "," _
        & " SpecialValueNotApplicable = " & sSVNotApplicable
    'MLM 15/04/05:
    If mlStudyId > 0 Then
        sSQL = sSQL & " WHERE ClinicalTrialId = " & mlStudyId
    End If
        
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.StoreDataViewOptions"
End Sub

'---------------------------------------------------------------------
Private Function QuestionNamesNotOK() As Boolean
'---------------------------------------------------------------------
'Check the current data base for invalid questions that might have been
'created in an earlier version of Macro before Macro's reserved words list
'was extended iwth words like 'site', 'personid' etc.
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bInvalidQuestionsFound As Boolean
Dim sPrompt As String

    On Error GoTo ErrHandler

    bInvalidQuestionsFound = False
    
    sSQL = "SELECT DataItem.DataItemCode, ClinicalTrial.ClinicalTrialName " _
        & " FROM DataItem, ClinicalTrial " _
        & " WHERE DataItem.ClinicalTrialId = ClinicalTrial.ClinicalTrialId " _
        & " ORDER BY ClinicalTrial.ClinicalTrialName, DataItem.DataItemCode"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Do Until rsTemp.EOF
        If Not gblnNotAReservedWord(rsTemp!DataItemCode) Then
            sPrompt = "The question '" & rsTemp!DataItemCode & "' in Study '" & rsTemp!ClinicalTrialname & "' is no" _
                & vbNewLine & "longer valid, since it is now one of Macro's reserved words." _
                & vbNewLine & "You must replace it before Create Data Views can be run." _
                & vbNewLine & vbNewLine & "To replace it, copy the original question and then replace" _
                & vbNewLine & "all occurrences of the original question with the new one," _
                & vbNewLine & "prior to deleting the original question."
            Call DialogInformation(sPrompt, "Invalid Question Name")
            bInvalidQuestionsFound = True
        End If
        rsTemp.MoveNext
    Loop
    
    If bInvalidQuestionsFound Then
        QuestionNamesNotOK = True
    Else
        QuestionNamesNotOK = False
    End If
        
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.QuestionNamesNotOK"
End Function

'---------------------------------------------------------------------
Private Function CheckSQLServerTriggerTableStructure() As Boolean
'---------------------------------------------------------------------
' RJCW 06/09/2001
' This sub checks for the presents of the DataViewTrigCalc field
' in the table DataViewDetails if it does exist the function returns
' true else False is returned
'---------------------------------------------------------------------

Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    CheckSQLServerTriggerTableStructure = False
    
    sSQL = "UPDATE DataViewDetails SET DataViewTrigCalc = ''"
        
    MacroADODBConnection.Execute sSQL

    CheckSQLServerTriggerTableStructure = True


Exit Function
ErrHandler:
'With an Oracle or SQLServer database error number 3604 'Invalid column name' gets generated.
If vbObjectError - 3604 Then Exit Function
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CheckSQLServerTriggerTableStructure"
End Function

'---------------------------------------------------------------------
Private Sub AlterDataViewDetails()
'---------------------------------------------------------------------
' RJCW 06/09/2001
' This sub adds the DataViewTrigCalc field to the table DataViewDetails
'---------------------------------------------------------------------

Dim sSQL As String
    
    On Error GoTo ErrHandler

    sSQL = sSQL & " ALTER TABLE  DataViewDetails  ADD  DataViewTrigCalc VARCHAR(1)  NULL"

    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.AlterDataViewDetails"
End Sub

'---------------------------------------------------------------------
Private Function DataViewDetailsContainsOutputCategoryValues() As Boolean
'---------------------------------------------------------------------
'Checks DataViewDetails table for column OutputCategoryValues
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
    
    DataViewDetailsContainsOutputCategoryValues = False
    
    sSQL = "SELECT OutputCategoryValues FROM DataViewDetails"
        
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    rsTemp.Close
    Set rsTemp = Nothing

    DataViewDetailsContainsOutputCategoryValues = True

Exit Function
ErrHandler:
    'The expected errors are as follows:-
    'With an Access database error number 3600 'No value given for one or more required parameters' gets generated.
    'With an Oracle or SQLServer database error number 3604 'Invalid column name' gets generated.
    If Err.Number - vbObjectError = 3600 Then Exit Function
    If Err.Number - vbObjectError = 3604 Then Exit Function
    Err.Raise Err.Number, , Err.Description & "|frmMenu.DataViewDetailsContainsOutputCategoryValues"
End Function

'---------------------------------------------------------------------
Private Sub AddOutputCategoryValuesToDataViewDetails()
'---------------------------------------------------------------------
'Adds column OutputCategoryValues to DataViewDetails table if missing
'---------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Access
        sSQL = "ALTER Table DataViewDetails ADD COLUMN OutputCategoryValues SMALLINT"
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sSQL = "ALTER Table DataViewDetails ADD OutputCategoryValues INTEGER"
    Case MACRODatabaseType.Oracle80
        sSQL = "ALTER Table DataViewDetails ADD OutputCategoryValues INTEGER"
    End Select

    MacroADODBConnection.Execute sSQL
    
    sSQL = "UPDATE DataViewDetails Set OutputCategoryValues = " & DataViewOption.Required
        
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.AddOutputCategoryValuesToDataViewDetails"
End Sub

'---------------------------------------------------------------------
Private Function DateFormatCanBeConverted(ByRef sDateFormat As String) As Boolean
'---------------------------------------------------------------------
'This function is used to decide which date/time questions can be converted
'from string fields to DateTime fields.
'
'This function takes a date/time format string and standardizes its format using
'several Replace statements.
'There are 20 standard formats.
'11 of the standard formats will be converted into DateTime fields:-
'   d/m/y       d/m/y/h/m       d/m/y/h/m/s
'   m/d/y       m/d/y/h/m       m/d/y/h/m/s
'   y/m/d       y/m/d/h/m       y/m/d/h/m/s
'   h/m
'   h/m/s
'9 of the formats will not be converted into DateTime fields, they will remain as strings
'   y/m     y/m/h/m     y/m/h/m/s
'   m/y     m/y/h/m     m/y/h/m/s
'   y/d/m   y/d/m/h/m   y/d/m/h/m/s
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'Replace double format characters with single format characters
    sDateFormat = Replace(sDateFormat, "dd", "d")
    sDateFormat = Replace(sDateFormat, "mm", "m")
    sDateFormat = Replace(sDateFormat, "hh", "h")
    sDateFormat = Replace(sDateFormat, "ss", "s")
    'Replace yyyy with y
    sDateFormat = Replace(sDateFormat, "yyyy", "y")
    'Replace all Date/Time Separators with "/"
    sDateFormat = Replace(sDateFormat, ":", "/")
    sDateFormat = Replace(sDateFormat, ".", "/")
    sDateFormat = Replace(sDateFormat, "-", "/")
    sDateFormat = Replace(sDateFormat, " ", "/")
    
    Select Case sDateFormat
    Case "d/m/y", "m/d/y", "y/m/d", "h/m", "h/m/s", "d/m/y/h/m", "m/d/y/h/m", "y/m/d/h/m", "d/m/y/h/m/s", "m/d/y/h/m/s", "y/m/d/h/m/s"
        DateFormatCanBeConverted = True
    Case Else
        'date formats y/m, m/y and y/d/m (with or without time elements) are not converted
        DateFormatCanBeConverted = False
    End Select

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.DateFormatCanBeConverted"
End Function

'---------------------------------------------------------------------
Private Function UniversalDateFormatString(ByVal sDateFormat As String, ByVal sResponse As String) As String
'---------------------------------------------------------------------
'This function is passed a standardized date/time format string together
'with a corresponding response.
'The contents of sResponse will have all of its Separators Replaced by "/"
'This function will then extract the d/m/y/h/m/s elements and then construct a
'Universal Date Format string in the form m/d/yyyy h:m:s
'It is the Universal Date Format string that is used to write the response
'into a DateTime field.
'---------------------------------------------------------------------
Dim sDay As String
Dim sMonth As String
Dim sYear As String
Dim sHour As String
Dim sMin As String
Dim sSec As String
Dim asElements() As String
Dim sUDF As String

    On Error GoTo ErrHandler
    
    'Replace all Separators with "/"
    sResponse = Replace(sResponse, ":", "/")
    sResponse = Replace(sResponse, ".", "/")
    sResponse = Replace(sResponse, "-", "/")
    sResponse = Replace(sResponse, " ", "/")
    
    Select Case sDateFormat
    Case "d/m/y"
        asElements = Split(sResponse, "/")
        sDay = asElements(0)
        sMonth = asElements(1)
        sYear = asElements(2)
        sUDF = sMonth & "/" & sDay & "/" & sYear
    Case "m/d/y"
        asElements = Split(sResponse, "/")
        sMonth = asElements(0)
        sDay = asElements(1)
        sYear = asElements(2)
        sUDF = sMonth & "/" & sDay & "/" & sYear
    Case "y/m/d"
        asElements = Split(sResponse, "/")
        sYear = asElements(0)
        sMonth = asElements(1)
        sDay = asElements(2)
        sUDF = sMonth & "/" & sDay & "/" & sYear
    Case "h/m"
        asElements = Split(sResponse, "/")
        sHour = asElements(0)
        sMin = asElements(1)
        sUDF = sHour & ":" & sMin
    Case "h/m/s"
        asElements = Split(sResponse, "/")
        sHour = asElements(0)
        sMin = asElements(1)
        sSec = asElements(2)
        sUDF = sHour & ":" & sMin & ":" & sSec
    Case "d/m/y/h/m"
        asElements() = Split(sResponse, "/")
        sDay = asElements(0)
        sMonth = asElements(1)
        sYear = asElements(2)
        sHour = asElements(3)
        sMin = asElements(4)
        sUDF = sMonth & "/" & sDay & "/" & sYear & " " & sHour & ":" & sMin
    Case "m/d/y/h/m"
        asElements = Split(sResponse, "/")
        sMonth = asElements(0)
        sDay = asElements(1)
        sYear = asElements(2)
        sHour = asElements(3)
        sMin = asElements(4)
        sUDF = sMonth & "/" & sDay & "/" & sYear & " " & sHour & ":" & sMin
    Case "y/m/d/h/m"
        asElements = Split(sResponse, "/")
        sYear = asElements(0)
        sMonth = asElements(1)
        sDay = asElements(2)
        sHour = asElements(3)
        sMin = asElements(4)
        sUDF = sMonth & "/" & sDay & "/" & sYear & " " & sHour & ":" & sMin
    Case "d/m/y/h/m/s"
        asElements = Split(sResponse, "/")
        sDay = asElements(0)
        sMonth = asElements(1)
        sYear = asElements(2)
        sHour = asElements(3)
        sMin = asElements(4)
        sSec = asElements(5)
        sUDF = sMonth & "/" & sDay & "/" & sYear & " " & sHour & ":" & sMin & ":" & sSec
    Case "m/d/y/h/m/s"
        asElements = Split(sResponse, "/")
        sMonth = asElements(0)
        sDay = asElements(1)
        sYear = asElements(2)
        sHour = asElements(3)
        sMin = asElements(4)
        sSec = asElements(5)
        sUDF = sMonth & "/" & sDay & "/" & sYear & " " & sHour & ":" & sMin & ":" & sSec
    Case "y/m/d/h/m/s"
        asElements = Split(sResponse, "/")
        sYear = asElements(0)
        sMonth = asElements(1)
        sDay = asElements(2)
        sHour = asElements(3)
        sMin = asElements(4)
        sSec = asElements(5)
        sUDF = sMonth & "/" & sDay & "/" & sYear & " " & sHour & ":" & sMin & ":" & sSec
    End Select
    
    UniversalDateFormatString = sUDF

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.UniversalDateFormatString"
End Function

'-----------------------------------------------------------------------
Private Sub CreateRQGDataViewTable(ByVal sTableName As String, _
                                    ByVal lClinicalTrialId As Long, _
                                    ByVal lCRFPageId As Long, _
                                    ByVal lGroupID As Long)
'-----------------------------------------------------------------------
'ASH 16/07/2002
'Based on CreateRODataViewTable
'Creates Repeating question group table/structure
'Mo Morris  24/11/2004 - Bug 2413 Category codes in Numeric Fields option
'Mo 24/10/2006 Bug 2824 - Corrected for Oracle databases.
'-----------------------------------------------------------------------
Dim sSQL As String
Dim rsFormDataItemRQGCodes As ADODB.Recordset
Dim sSQLDataItemRQGCodes As String

    On Error GoTo ErrHandler

    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Access
        sSQL = "CREATE TABLE " & sTableName & "(ClinicalTrialId INTEGER," _
            & " Site TEXT(8)," _
            & " PersonId INTEGER," _
            & " VisitId INTEGER," _
            & " VisitCycleNumber SMALLINT," _
            & " CRFPageId INTEGER," _
            & " CRFPageCycleNumber SMALLINT," _
            & " OwnerQGroupID SMALLINT," _
            & " RepeatNumber SMALLINT,"
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sSQL = "CREATE TABLE " & sTableName & "(ClinicalTrialId INTEGER," _
            & " Site VARCHAR(8)," _
            & " PersonId INTEGER," _
            & " VisitId INTEGER," _
            & " VisitCycleNumber INTEGER," _
            & " CRFPageId INTEGER," _
            & " CRFPageCycleNumber INTEGER," _
            & " OwnerQGroupID INTEGER," _
            & " RepeatNumber INTEGER,"
    Case MACRODatabaseType.Oracle80
        'Mo 24/10/2006 Bug 2824, newly added Oracle option, Oracle used to use the SQLServer option
        sSQL = "CREATE TABLE " & sTableName & "(ClinicalTrialId NUMBER(11)," _
            & " Site VARCHAR2(8)," _
            & " PersonId NUMBER(11)," _
            & " VisitId NUMBER(11)," _
            & " VisitCycleNumber NUMBER(11)," _
            & " CRFPageId NUMBER(11)," _
            & " CRFPageCycleNumber NUMBER(11)," _
            & " OwnerQGroupID NUMBER(11)," _
            & " RepeatNumber NUMBER(11),"
    End Select
    
    'create a recordset of the questions within form lCRFPageId that belong to a RQG with OwnerQGroupID equal lGroupID
    'Mo 2/5/2003, bug 1424 , SQL now correctly gets the questions within a specific RQG on a specific eForm
    'Mo 24/10/2006 Bug 2824, DataItemCase added to following SQL
    sSQLDataItemRQGCodes = "SELECT DataItem.DataItemId, DataItem.DataItemCode, DataItem.DataType, DataItem.DataItemFormat, DataItem.DataItemCase" _
        & " FROM CRFElement, DataItem, QGroupQuestion" _
        & " WHERE CRFElement.ClinicalTrialId = DataItem.ClinicalTrialId " _
        & " AND CRFElement.DataItemId = DataItem.DataItemId " _
        & " AND CRFElement.ClinicalTrialId = QGroupQuestion.ClinicalTrialId " _
        & " AND CRFElement.OwnerQGroupID = QGroupQuestion.QGroupId " _
        & " AND CRFElement.DataItemId = QGroupQuestion.DataItemId " _
        & " AND CRFElement.ClinicalTrialId = " & lClinicalTrialId _
        & " AND CRFElement.OwnerQGroupID = " & lGroupID _
        & " AND CRFElement.CRFPageId = " & lCRFPageId _
        & " ORDER BY QGroupQuestion.QOrder"
    
    Set rsFormDataItemRQGCodes = New ADODB.Recordset
    rsFormDataItemRQGCodes.Open sSQLDataItemRQGCodes, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    
    Do Until rsFormDataItemRQGCodes.EOF
        
        'Mo 24/10/2006 Bug 2824, DataItemCase added to call to DataBaseSpecificColumnFormatting
        sSQL = sSQL & DataBaseSpecificColumnFormatting(lClinicalTrialId, _
                rsFormDataItemRQGCodes!DataItemId, rsFormDataItemRQGCodes!DataItemCode, CInt(RemoveNull(rsFormDataItemRQGCodes!DataItemCase)), _
                ConvertFromNull(rsFormDataItemRQGCodes!DataItemFormat, vbString), _
                ConvertFromNull(rsFormDataItemRQGCodes!DataType, vbInteger))
        rsFormDataItemRQGCodes.MoveNext
    
    Loop
    
    rsFormDataItemRQGCodes.Close
    Set rsFormDataItemRQGCodes = Nothing
        
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY"
    Else
        sSQL = sSQL & " CONSTRAINT PK" & sTableName & " PRIMARY KEY"
    End If
    
    sSQL = sSQL & " (ClinicalTrialId,Site,PersonId,VisitId,VisitCycleNumber,CRFPageId,CRFPageCycleNumber,OwnerQGroupID,RepeatNumber))"

    MacroADODBConnection.Execute sSQL

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CreateRQGDataViewTable"
End Sub

'-------------------------------------------------------------------------------------
Private Sub PopulateROTable(ByVal lTrialID As Long, _
                            ByVal lPageID As Long, _
                            ByVal sTableName As String, _
                            ByVal lVisitID As Long, _
                            ByVal lGroupID As Long)
'-------------------------------------------------------------------------------------
'ASH 22/07/2002 - Populates Response value only table
'Formerly part of CreateAndPopulateDataviews routine
'LVisitID could have the value Null_Long
'Mo Morris  24/11/2004 - Bug 2413 Category codes in Numeric Fields option
'-------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsFormData As ADODB.Recordset
Dim sRowKey As String
Dim sPrevRowKey As String
Dim sResponse As String

    On Error GoTo ErrHandler

    'place progress message in txtProgressMessage
    txtProgressMessage.Text = "Populating " & sTableName
    DoEvents
    'Mo 24/10/2006 Bug 2824, DataItemCase added to following SQL
    sSQL = "SELECT  DataItemResponse.TrialSite,"
    sSQL = sSQL & " DataItemResponse.PersonId, DataItemResponse.VisitId,"
    sSQL = sSQL & " DataItemResponse.VisitCycleNumber,"
    sSQL = sSQL & " DataItemResponse.CRFPageCycleNumber, DataItem.DataItemCode, DataItem.DataType, DataItem.DataItemCase,"
    sSQL = sSQL & " DataItem.DataItemFormat, DataItemResponse.ResponseValue, DataItemResponse.ResponseStatus, DataItemResponse.RepeatNumber, DataItemResponse.ValueCode"
    sSQL = sSQL & " FROM DataItemResponse, DataItem, CRFElement"
    sSQL = sSQL & " WHERE DataItemResponse.ClinicalTrialId = " & lTrialID
    
    'If visit code is null then the Data View will be of the type "One Data View per eForm irrespective of the Visit"
    If lVisitID <> NULL_LONG Then
        sSQL = sSQL & " AND DataItemResponse.VisitId = " & lVisitID
    End If
    
    sSQL = sSQL & " AND DataItemResponse.CRFPageId = " & lPageID
    sSQL = sSQL & " AND CRFElement.CRFPageId = " & lPageID
    sSQL = sSQL & " AND CRFElement.OwnerQGroupID = " & lGroupID
    sSQL = sSQL & " AND DataItemResponse.ClinicalTrialId = DataItem.ClinicalTrialId"
    sSQL = sSQL & " AND DataItemResponse.DataItemId = DataItem.DataItemId"
    sSQL = sSQL & " AND DataItemResponse.ClinicalTrialId = CRFElement.ClinicalTrialId"
    sSQL = sSQL & " AND DataItemResponse.DataItemId = CRFElement.DataItemId"
    sSQL = sSQL & " ORDER BY DataItemResponse.TrialSite,"
    sSQL = sSQL & " DataItemResponse.PersonId, DataItemResponse.VisitId,"
    sSQL = sSQL & " DataItemResponse.VisitCycleNumber, DataItemResponse.CRFPageID,"
    sSQL = sSQL & " DataItemResponse.CRFPageCycleNumber"
         
     'If QGroupId is > 0 then the Data View will for RQGs
    If lGroupID > 0 Then
        sSQL = sSQL & " ,DataItemResponse.RepeatNumber , DataItem.DataItemCode "
    Else
         sSQL = sSQL & " ,DataItem.DataItemCode"
    End If

    Set rsFormData = New ADODB.Recordset
    rsFormData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    'loop through the data adding it to the Data view table
    sPrevRowKey = ""
    Do Until rsFormData.EOF
        'Test for a new Row and ceate a record for it.
        sRowKey = rsFormData!TrialSite & "|" & rsFormData!PersonId & "|" _
            & rsFormData!VisitId & "|" & rsFormData!VisitCycleNumber & "|" _
            & rsFormData!CRFPageCycleNumber
            
        If lGroupID > 0 Then
           sRowKey = sRowKey & "|" & rsFormData!RepeatNumber
        End If

        If sRowKey <> sPrevRowKey Then
        'Need to insert new row
            sPrevRowKey = sRowKey
            sSQL = "INSERT INTO " & sTableName & " (ClinicalTrialId, Site, PersonId,"
            sSQL = sSQL & " VisitId, VisitCycleNumber, CRFPageID, CRFPageCycleNumber"
            
            If lGroupID > 0 Then
            'for RQGs we need to include GroupID and RepeatNumber
                sSQL = sSQL & " ,OwnerQGroupID,RepeatNumber)"
            Else
                sSQL = sSQL & ")"
            End If

            sSQL = sSQL & " VALUES (" & lTrialID & ",'" & rsFormData!TrialSite & "',"
            sSQL = sSQL & rsFormData!PersonId & "," & rsFormData!VisitId & "," & rsFormData!VisitCycleNumber & ","
            sSQL = sSQL & lPageID & "," & rsFormData!CRFPageCycleNumber
            
            If lGroupID > 0 Then
                sSQL = sSQL & "," & lGroupID & "," & rsFormData!RepeatNumber & ")"
            Else
                sSQL = sSQL & ")"
            End If

            MacroADODBConnection.Execute sSQL
        End If
        
        'Format Responses
        'Mo 24/10/2006 Bug 2824, DataItemCase added to call to FormatResponse
        sResponse = FormatResponse(ConvertFromNull(rsFormData!ResponseValue, vbString), _
                    sTableName, rsFormData!DataItemCode, rsFormData!ResponseStatus, CInt(RemoveNull(rsFormData!DataItemCase)), _
                    ConvertFromNull(rsFormData!DataItemFormat, vbString), _
                    ConvertFromNull(rsFormData!DataType, vbInteger), _
                    ConvertFromNull(rsFormData!ValueCode, vbString))
    
        'update the current dataitem
        sSQL = "UPDATE " & sTableName & " SET "
        sSQL = sSQL & rsFormData!DataItemCode & " = " & sResponse
        sSQL = sSQL & " WHERE ClinicalTrialId = " & lTrialID
        sSQL = sSQL & " AND Site = '" & rsFormData!TrialSite & "'"
        sSQL = sSQL & " AND PersonId = " & rsFormData!PersonId
        sSQL = sSQL & " AND VisitId = " & rsFormData!VisitId
        sSQL = sSQL & " AND VisitCycleNumber = " & rsFormData!VisitCycleNumber
        sSQL = sSQL & " AND CRFPageID = " & lPageID
        sSQL = sSQL & " AND CRFPageCycleNumber = " & rsFormData!CRFPageCycleNumber
        
        'If QGroupId is > 0 then the Data View will for RQGs
        If lGroupID > 0 Then
            sSQL = sSQL & " AND OwnerQGroupID = " & lGroupID
            sSQL = sSQL & " AND RepeatNumber = " & rsFormData!RepeatNumber
        End If
    
        MacroADODBConnection.Execute sSQL
        
        rsFormData.MoveNext
    Loop
    
    rsFormData.Close
    Set rsFormData = Nothing

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.PopulateROTable"
End Sub

'-------------------------------------------------------------------------------------
Private Sub PopulateWATable(ByVal lTrialID As Long, _
                            ByVal lPageID As Long, _
                            ByVal sTableName As String, _
                            ByVal lVisitID As Long)
'-------------------------------------------------------------------------------------
'ASH 22/07/2002 - Populates Response value Plus attributes tables
'Formerly part of CreateAndPopulateDataviews routine.Populates WA tables.
'-------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsFormData As ADODB.Recordset
Dim sResponse As String

    On Error GoTo ErrHandler

    'place progress message in txtProgressMessage
    txtProgressMessage.Text = "Populating " & sTableName
    DoEvents
    sSQL = "SELECT DataItemResponse.ClinicalTrialId, DataItemResponse.TrialSite,"
    sSQL = sSQL & " DataItemResponse.PersonId, DataItemResponse.VisitId,"
    sSQL = sSQL & " DataItemResponse.VisitCycleNumber, DataItemResponse.CRFPageID,"
    sSQL = sSQL & " DataItemResponse.CRFPageCycleNumber, DataItem.DataItemCode,"
    sSQL = sSQL & " DataItemResponse.ResponseTimeStamp, DataItem.DataType,"
    sSQL = sSQL & " DataItemResponse.ResponseValue, DataItemResponse.UnitOfMeasurement,"
    sSQL = sSQL & " DataItemResponse.ValueCode, DataItemResponse.ResponseStatus,"
    sSQL = sSQL & " DataItemResponse.LabResult, DataItemResponse.CTCGrade,"
    sSQL = sSQL & " DataItemResponse.RepeatNumber, CRFElement.OwnerQGroupID"
    sSQL = sSQL & " FROM DataItemResponse, DataItem, CRFElement"
    sSQL = sSQL & " WHERE DataItemResponse.ClinicalTrialId = " & lTrialID
    'If visit code is null then the Data View will be of the type "One Data View per eForm irrespective of the Visit"
    If lVisitID <> NULL_LONG Then
        sSQL = sSQL & " AND DataItemResponse.VisitId = " & lVisitID
    End If
    sSQL = sSQL & " AND DataItemResponse.CRFPageId = " & lPageID
    sSQL = sSQL & " AND CRFElement.CRFPageId = " & lPageID
    sSQL = sSQL & " AND DataItemResponse.ClinicalTrialId = DataItem.ClinicalTrialId"
    sSQL = sSQL & " AND DataItemResponse.ClinicalTrialId = CRFElement.ClinicalTrialId"
    sSQL = sSQL & " AND DataItemResponse.CRFElementID = CRFElement.CRFElementID"
    sSQL = sSQL & " AND DataItemResponse.DataItemId = DataItem.DataItemId"
    sSQL = sSQL & " ORDER BY DataItemResponse.TrialSite,"
    sSQL = sSQL & " DataItemResponse.PersonId, DataItemResponse.VisitId,"
    sSQL = sSQL & " DataItemResponse.VisitCycleNumber, DataItemResponse.CRFPageID,"
    sSQL = sSQL & " DataItemResponse.CRFPageCycleNumber,CRFElement.OwnerQGroupID,"
    sSQL = sSQL & " DataItem.DataItemCode,DataItemResponse.RepeatNumber"

    Set rsFormData = New ADODB.Recordset
      
    rsFormData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

    'loop through the data adding it to the Data view table
    Do Until rsFormData.EOF
        'Changed Mo Morris SR5019 17/1/2003, The Null test fails to pick up empty strings from converted databases
        'If Not IsNull(rsFormData!ResponseValue) Then
        If Len(Trim(rsFormData!ResponseValue)) > 0 Then
            sResponse = "'" & ReplaceQuotes(rsFormData!ResponseValue) & "'"
        Else
            sResponse = "null"
        End If
        sSQL = "INSERT INTO " & sTableName & " VALUES(" _
            & rsFormData!ClinicalTrialId & ",'" & rsFormData!TrialSite & "'," _
            & rsFormData!PersonId & "," & rsFormData!VisitId & "," _
            & rsFormData!VisitCycleNumber & "," & rsFormData!CRFPageID & "," _
            & rsFormData!CRFPageCycleNumber & "," & rsFormData!OwnerQGroupID & ",'" _
            & rsFormData!DataItemCode & "'," & rsFormData!RepeatNumber & ",'" _
            & Format(rsFormData!ResponseTimeStamp, "yyyy/mm/dd hh:mm:ss") & "'," & rsFormData!DataType & "," _
            & sResponse & ",'" & rsFormData!UnitOfMeasurement & "','" & rsFormData!ValueCode & "'," _
            & rsFormData!ResponseStatus & ",'" & rsFormData!LabResult & "'," _
            & VarianttoString(rsFormData!CTCGrade, True) & ")"
          
        MacroADODBConnection.Execute sSQL
        rsFormData.MoveNext
    Loop
    rsFormData.Close
    Set rsFormData = Nothing

Exit Sub
ErrHandler:
Err.Raise Err.Number, , Err.Description & "|frmMenu.PopulateWATable"
End Sub

'----------------------------------------------------------------------------------------------
Private Function FormatResponse(ByVal sResponseToFormat As String, _
                                ByVal sTableName As String, _
                                ByVal sDataItemCode As String, _
                                ByVal nStatus As Integer, _
                                ByVal nDataItemCase As Integer, _
                                Optional ByVal sDataItemToFormat As String = "", _
                                Optional ByVal sDataType As Integer = DataType.Text, _
                                Optional ByVal sValueCode As String = "") As String
'----------------------------------------------------------------------------------------------
'ASH 22/07/2002 - Formerly part of CreateAndPopulateDataviews routine
'Formats a response and returns it as  a  formatted string ready for inclusion in SQL
'including all necessary quotes
'Mo Morris  24/11/2004 - Bug 2413 Category codes in Numeric Fields option
'Mo Morris  25/11/2004 - Bug 2268 Oracle Data View tables not being populated with Times.
'Mo 24/10/2006 Bug 2824, DataItemCase added to FormatResponse
'----------------------------------------------------------------------------------------------
Dim sResponse As String
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
    
    If Len(Trim(sResponseToFormat)) = 0 Then
        'its an empty response, check for the existence of special values
        Select Case nStatus
        Case Status.Missing
            If txtMissing.Text = "" Then
                sResponse = "null"
            Else
                Select Case sDataType
                Case DataType.Date
                    'Mo 24/10/2006 Bug 2824, check Partial Dates flag before deciding on date format
                    If nDataItemCase = 0 And DateFormatCanBeConverted(sDataItemToFormat) Then
                        'Special Value dates being converted into special dates
                        Select Case goUser.Database.DatabaseType
                        Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
                            sResponse = "CONVERT(DATETIME,'" & Format(CDate(txtMissing.Text), "mm/dd/yyyy") & "',101)"
                        Case MACRODatabaseType.Oracle80
                            sResponse = "to_date('" & Format(CDate(txtMissing.Text), "mm/dd/yyyy") & "','mm/dd/yyyy hh24:mi:ss')"
                        End Select
                    Else
                        sResponse = txtMissing.Text
                    End If
                Case Else
                    sResponse = txtMissing.Text
                End Select
            End If
        Case Status.NotApplicable
            If txtNotApplicable.Text = "" Then
                sResponse = "null"
            Else
                Select Case sDataType
                Case DataType.Date
                    'Mo 24/10/2006 Bug 2824, check Partial Dates flag before deciding on date format
                    If nDataItemCase = 0 And DateFormatCanBeConverted(sDataItemToFormat) Then
                        'Special Value dates being converted into special dates
                        Select Case goUser.Database.DatabaseType
                        Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
                            sResponse = "CONVERT(DATETIME,'" & Format(CDate(txtNotApplicable.Text), "mm/dd/yyyy") & "',101)"
                        Case MACRODatabaseType.Oracle80
                            sResponse = "to_date('" & Format(CDate(txtNotApplicable.Text), "mm/dd/yyyy") & "','mm/dd/yyyy hh24:mi:ss')"
                        End Select
                    Else
                        sResponse = txtNotApplicable.Text
                    End If
                Case Else
                    sResponse = txtNotApplicable.Text
                End Select
            End If
        Case Status.Unobtainable
            If txtUnobtainable.Text = "" Then
                sResponse = "null"
            Else
                Select Case sDataType
                Case DataType.Date
                    'Mo 24/10/2006 Bug 2824, check Partial Dates flag before deciding on date format
                    If nDataItemCase = 0 And DateFormatCanBeConverted(sDataItemToFormat) Then
                        'Special Value dates being converted into special dates
                        Select Case goUser.Database.DatabaseType
                        Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
                            sResponse = "CONVERT(DATETIME,'" & Format(CDate(txtUnobtainable.Text), "mm/dd/yyyy") & "',101)"
                        Case MACRODatabaseType.Oracle80
                            sResponse = "to_date('" & Format(CDate(txtUnobtainable.Text), "mm/dd/yyyy") & "','mm/dd/yyyy hh24:mi:ss')"
                        End Select
                    Else
                        sResponse = txtUnobtainable.Text
                    End If
                Case Else
                    sResponse = txtUnobtainable.Text
                End Select
            End If
        Case Else
            sResponse = "null"
        End Select
    Else
        Select Case sDataType
            'Mo 25/10/2005 COD0030
            Case DataType.Text, DataType.Multimedia, DataType.Thesaurus
                sResponse = "'" & ReplaceQuotes(sResponseToFormat) & "'"
            'Changed Mo 18/6/2002, standard dates/times will be converted to DateTime fields
            Case DataType.Date
                'Mo 24/10/2006 Bug 2824, check Partial Dates flag before deciding on date format
                If nDataItemCase = 0 And DateFormatCanBeConverted(sDataItemToFormat) Then
                    Select Case goUser.Database.DatabaseType
                        Case MACRODatabaseType.Access
                            sResponse = "#" & UniversalDateFormatString(sDataItemToFormat, sResponseToFormat) & "#"
                        Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
                                'The SQL Server CONVERT function is set to use style 101 which is mm/dd/yyyy
                                'Note that times without a date is always given the date 01/01/1900
                                sResponse = "CONVERT(DATETIME,'" & UniversalDateFormatString(sDataItemToFormat, sResponseToFormat) & "',101)"
                        Case MACRODatabaseType.Oracle80
                            If sDataItemToFormat = "h/m" Or sDataItemToFormat = "h/m/s" Then
                                'For times without a date the date 01/01/1900 is added (so that Oracle behaves the same as SQL Server)
                                sResponse = "to_date('01/01/1900 " & UniversalDateFormatString(sDataItemToFormat, sResponseToFormat) & "','mm/dd/yyyy hh24:mi:ss')"
                            Else
                                sResponse = "to_date('" & UniversalDateFormatString(sDataItemToFormat, sResponseToFormat) & "','mm/dd/yyyy hh24:mi:ss')"
                            End If
                    End Select
                Else
                    sResponse = "'" & sResponseToFormat & "'"
                End If
            Case DataType.Category
                Select Case mnOutputCategoryValues
                Case DataViewCategoryOptions.Codes
                    sResponse = "'" & sValueCode & "'"
                Case DataViewCategoryOptions.Values
                    sResponse = "'" & ReplaceQuotes(sResponseToFormat) & "'"
                Case DataViewCategoryOptions.TypedCodes
                    Select Case goUser.Database.DatabaseType
                    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
                        sSQL = "SELECT data_type FROM information_Schema.columns " _
                            & "WHERE upper(table_name) = upper('" & sTableName & "') " _
                            & "AND upper(column_name)= upper('" & sDataItemCode & "')"
                        Set rsTemp = New ADODB.Recordset
                        rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                        If rsTemp!data_type = "varchar" Then
                            sResponse = "'" & sValueCode & "'"
                        Else
                            sResponse = sValueCode
                        End If
                    Case MACRODatabaseType.Oracle80
                        sSQL = "SELECT data_type FROM user_tab_columns " _
                            & "WHERE upper(table_name) = upper('" & sTableName & "') " _
                            & "AND upper(column_name) = upper('" & sDataItemCode & "')"
                        Set rsTemp = New ADODB.Recordset
                        rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                        If rsTemp!data_type = "VARCHAR2" Then
                            sResponse = "'" & sValueCode & "'"
                        Else
                            sResponse = sValueCode
                        End If
                    End Select
                End Select
            Case DataType.IntegerData, DataType.Real, DataType.LabTest
                sResponse = sResponseToFormat
        End Select
    End If
    
    FormatResponse = sResponse

Exit Function:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.FormatResponse"
End Function

'-------------------------------------------------------------------------------------------------
Private Function DataBaseSpecificColumnFormatting(ByVal lClinicalTrialId As Long, _
                                                ByVal lDataItemId As Long, _
                                                ByVal sDataItemCode As String, _
                                                ByVal nDataItemCase As Integer, _
                                                Optional ByVal sDataItemToFormat As String = "", _
                                                Optional ByVal nDataType As Integer = 0) As String
'-------------------------------------------------------------------------------------------------
'ASH 22/07/2002 - Formerly part of CreateAndPopulateDataviews routine
'ASH 25/07/2002 - Creates specific columns types for different database types
'Mo Morris 16/4/2002
'   Changes made so that the question columns in the dataview tables
'   are typed to reflect their content:-
'
'   Question Type   in Access DB    in SQL Server DB     in Oracle DB
'   Text            TEXT(255)       VARCHAR(255)        VARCHAR2(255)
'   Category        TEXT(255)       VARCHAR(255)        VARCHAR2(255)
'   Multimedia      TEXT(255)       VARCHAR(255)        VARCHAR2(255)
'   Thesaurus       TEXT(255)       VARCHAR(255)        VARCHAR2(255)
'   Date            DATETIME        DATETIME            DATE
'   IntegerData     INTEGER         INTEGER             NUMBER(11)    TA - now BIGINT and FLOAT respectively
'   Real            DOUBLE          NUMERIC(16,10)      NUMBER(16,10) TA - now FLOAT
'   LabTest         DOUBLE          NUMERIC(16,10)      NUMBER(16,10) TA - now FLOAT
'Mo Morris  24/11/2004 - Bug 2413 Category codes in Numeric Fields option
'Mo 24/10/2006 Bug 2824, DataItemCase added to DataBaseSpecificColumnFormatting
'--------------------------------------------------------------------------------------------------
Dim sSQL As String
Dim sDataItemFormat As String
Dim bCatCodesNumeric As Boolean
Dim nCatCodeLength As Integer

    On Error GoTo ErrHandler
    sSQL = ""
    sSQL = sSQL & " " & sDataItemCode
    
    Select Case goUser.Database.DatabaseType
        Case MACRODatabaseType.Access
            Select Case nDataType
                'Mo 25/10/2005 COD0030
                Case DataType.Text, DataType.Multimedia, DataType.Thesaurus
                    sSQL = sSQL & " TEXT(255),"
                Case DataType.Category
                    If mnOutputCategoryValues = DataViewCategoryOptions.TypedCodes Then
                        Call AssessCategoryCodes(lClinicalTrialId, 1, lDataItemId, bCatCodesNumeric, nCatCodeLength)
                        If bCatCodesNumeric Then
                            sSQL = sSQL & " INTEGER,"
                        Else
                            sSQL = sSQL & " TEXT(255),"
                        End If
                    Else
                        sSQL = sSQL & " TEXT(255),"
                    End If
                Case DataType.Date
                    sDataItemFormat = sDataItemToFormat
                    'Mo 24/10/2006 Bug 2824, check Partial Dates flag before deciding on date format
                    If nDataItemCase = 0 And DateFormatCanBeConverted(sDataItemFormat) Then
                        sSQL = sSQL & " DATETIME,"
                    Else
                        sSQL = sSQL & " TEXT(255),"
                    End If
                Case DataType.IntegerData
                    sSQL = sSQL & " INTEGER,"
                Case DataType.Real, DataType.LabTest
                    sSQL = sSQL & " DOUBLE,"
            End Select
        
        Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
            Select Case nDataType
                'Mo 25/10/2005 COD0030
                Case DataType.Text, DataType.Multimedia, DataType.Thesaurus
                    sSQL = sSQL & " VARCHAR(255),"
                Case DataType.Category
                    If mnOutputCategoryValues = DataViewCategoryOptions.TypedCodes Then
                        Call AssessCategoryCodes(lClinicalTrialId, 1, lDataItemId, bCatCodesNumeric, nCatCodeLength)
                        If bCatCodesNumeric Then
                            'TA 23/02/2005: use setting to decide whether > 9 digits allowed
                            If LCase(GetMACROSetting(MACRO_SETTING_CDV_BIGINT, "true")) = "true" Then
                                sSQL = sSQL & " BIGINT,"
                            Else
                                sSQL = sSQL & " INTEGER,"
                            End If
                        Else
                            sSQL = sSQL & " VARCHAR(255),"
                        End If
                    Else
                        sSQL = sSQL & " VARCHAR(255),"
                    End If
                Case DataType.Date
                    sDataItemFormat = sDataItemToFormat
                    'Mo 24/10/2006 Bug 2824, check Partial Dates flag before deciding on date format
                    If nDataItemCase = 0 And DateFormatCanBeConverted(sDataItemFormat) Then
                        sSQL = sSQL & " DATETIME,"
                    Else
                        sSQL = sSQL & " VARCHAR(255),"
                    End If
                Case DataType.IntegerData
                    'TA 23/02/2005: use setting to decide whether > 9 digits allowed
                    If LCase(GetMACROSetting(MACRO_SETTING_CDV_BIGINT, "true")) = "true" Then
                        sSQL = sSQL & " BIGINT,"
                    Else
                        sSQL = sSQL & " INTEGER,"
                    End If
                Case DataType.Real, DataType.LabTest
                'TA 07/10/2004: changed to float
                    sSQL = sSQL & " FLOAT,"
            End Select
        
        Case MACRODatabaseType.Oracle80
            Select Case nDataType
                'Mo 25/10/2005 COD0030
                Case DataType.Text, DataType.Multimedia, DataType.Thesaurus
                    sSQL = sSQL & " VARCHAR2(255),"
                Case DataType.Category
                    If mnOutputCategoryValues = DataViewCategoryOptions.TypedCodes Then
                        Call AssessCategoryCodes(lClinicalTrialId, 1, lDataItemId, bCatCodesNumeric, nCatCodeLength)
                        If bCatCodesNumeric Then
                            'TA 23/02/2005: use setting to decide whether > 9 digits allowed
                            If LCase(GetMACROSetting(MACRO_SETTING_CDV_BIGINT, "true")) = "true" Then
                                sSQL = sSQL & " NUMBER(15),"
                            Else
                                sSQL = sSQL & " NUMBER(9),"
                            End If
                        Else
                            sSQL = sSQL & " VARCHAR2(255),"
                        End If
                    Else
                        sSQL = sSQL & " VARCHAR2(255),"
                    End If
                Case DataType.Date
                    sDataItemFormat = sDataItemToFormat
                    'Mo 24/10/2006 Bug 2824, check Partial Dates flag before deciding on date format
                    If nDataItemCase = 0 And DateFormatCanBeConverted(sDataItemFormat) Then
                        sSQL = sSQL & " DATE,"
                    Else
                        sSQL = sSQL & " VARCHAR2(255),"
                    End If
                Case DataType.IntegerData
                    'TA 23/02/2005: use setting to decide whether > 9 digits allowed
                    If LCase(GetMACROSetting(MACRO_SETTING_CDV_BIGINT, "true")) = "true" Then
                        sSQL = sSQL & " NUMBER(15),"
                    Else
                        sSQL = sSQL & " NUMBER(9),"
                    End If
                Case DataType.Real, DataType.LabTest
                    'TA 07/10/2004: changed to float
                    sSQL = sSQL & " FLOAT,"
            End Select
        
        End Select
        
        DataBaseSpecificColumnFormatting = sSQL

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.DataBaseSpecificColumnFormatting"
End Function

'---------------------------------------------------------------------------------
Private Sub CreateROTableTypeName(ByVal lTrialID As Long, _
                                ByVal sTrialName As String, _
                                ByVal sPageCode As String, _
                                ByVal lPageID As Long, _
                                Optional ByVal lVisitID As Long, _
                                Optional ByVal sVisitCode As String = "")
'---------------------------------------------------------------------------------
'ASH 23/07/2002 - Formerly part of CreateTableNames routine
'Generates names for Response Value Only dataview tables and inserts it into the DataViewTables table
'---------------------------------------------------------------------------------
Dim itmX As MSComctlLib.ListItem
Dim sTableName As String
Dim sSQL As String

    On Error GoTo ErrHandler
        
    'Create Response Value Only Tables
    If mbDataViewRORequired Then
        'set-up the name of a data view table
        If optVisitSeparate.Value Then
            sTableName = CreateTableName(sTrialName, _
                sPageCode, "RO", sVisitCode)
        Else
            sTableName = CreateTableName(sTrialName, _
                sPageCode, "RO")
        End If
        'Add Data View Name to ListView
        Set itmX = lvwTables.ListItems.Add(, , sTableName)
        itmX.SubItems(1) = "RO"
        itmX.SubItems(2) = sTrialName
        If optVisitSeparate.Value Then
            itmX.SubItems(3) = sVisitCode
        End If
        itmX.SubItems(4) = sPageCode
        
        'Add Data View Name to table DataViewTables
        'Changed Mo Morris, 29/8/2002, QGroupID of 0 added to Insert
        sSQL = "INSERT INTO DataViewTables (DataViewName,DataViewType,ClinicalTrialName,ClinicalTrialId,CRFPageCode,CRFPageId,QGroupID"
        If optVisitSeparate.Value = True Then
            sSQL = sSQL & ",VisitCode,VisitId"
        End If
        sSQL = sSQL & ") VALUES('" & sTableName & "','RO','" & sTrialName & "'," _
            & lTrialID & ",'" & sPageCode & "'," & lPageID & ",0"
        If optVisitSeparate.Value Then
            sSQL = sSQL & ",'" & sVisitCode & "'," & lVisitID
        End If
        sSQL = sSQL & ")"
        MacroADODBConnection.Execute sSQL
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CreateROTableTypeName"
End Sub

'--------------------------------------------------------------------------------------
Private Sub CreateRORQGTableTypeName(ByVal lTrialID As Long, _
                                    ByVal sTrialName As String, _
                                    ByVal sPageCode As String, _
                                    ByVal lPageID As Long, _
                                    Optional ByVal lVisitID As Long, _
                                    Optional ByVal sVisitCode As String = "")
'--------------------------------------------------------------------------------------
'ASH 23/07/2002 - Formerly part of CreateTableNames routine
'Generates names for Repeating Question Groups  dataview tables and inserts it into
'the DataViewTables table
'--------------------------------------------------------------------------------------
Dim itmX As MSComctlLib.ListItem
Dim rsRQG As ADODB.Recordset
Dim sTableName As String
Dim sSQL As String
    
    On Error GoTo ErrHandler
        
    If mbDataViewRORequired Then
        
    'ASH 11/07/2002
    'Now check to see if and how many RQGs exist on the eForm
    sSQL = " SELECT DISTINCT CRFElement.OwnerQGroupID, QGroup.QGroupID,"
    sSQL = sSQL & " QGroup.QGroupCode,QGroup.QGroupName "
    sSQL = sSQL & " FROM QGroup,CRFElement "
    sSQL = sSQL & " WHERE QGroup.ClinicalTrialId = " & lTrialID
    sSQL = sSQL & " AND CRFElement.ClinicalTrialId = " & lTrialID
    sSQL = sSQL & " AND QGroup.QGroupID = CRFElement.QGroupID"
    sSQL = sSQL & " AND CRFElement.CRFPageId = " & lPageID

    Set rsRQG = New ADODB.Recordset
    rsRQG.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    Do Until rsRQG.EOF
        'set-up the name of a data view table
        If optVisitSeparate.Value Then
            sTableName = CreateTableName _
                (sTrialName, sPageCode, "RO", sVisitCode, rsRQG!QGroupCode)
        Else
            sTableName = CreateTableName _
                (sTrialName, sPageCode, "RO", , rsRQG!QGroupCode)
        End If
        'Add Data View Name to ListView
        Set itmX = lvwTables.ListItems.Add(, , sTableName)
        itmX.SubItems(1) = "RO"
        itmX.SubItems(2) = sTrialName
        If optVisitSeparate.Value Then
            itmX.SubItems(3) = sVisitCode
        End If
        itmX.SubItems(4) = sPageCode
        'ASH 11/07/2002 add group code to listview
        itmX.SubItems(5) = rsRQG!QGroupCode
        
        'Add Data View Name to table DataViewTables
        sSQL = "INSERT INTO DataViewTables (DataViewName,DataViewType,ClinicalTrialName,ClinicalTrialId,CRFPageCode,CRFPageId,QGroupCode,QGroupID"
        If optVisitSeparate.Value Then
            sSQL = sSQL & ",VisitCode,VisitId"
        End If
        sSQL = sSQL & ") VALUES('" & sTableName & "','RO','" & sTrialName & "'," _
            & lTrialID & ",'" & sPageCode & "'," & lPageID & " ,'" & rsRQG!QGroupCode & "'," & rsRQG!QGroupID & ""
        If optVisitSeparate.Value Then
            sSQL = sSQL & ",'" & sVisitCode & "'," & lVisitID
        End If
        sSQL = sSQL & ")"
        
        MacroADODBConnection.Execute sSQL
        
        rsRQG.MoveNext
    Loop 'on rsRQG

End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CreateRORQGTableTypeName"
End Sub

'----------------------------------------------------------------------------------------------
Private Sub CreateWATableTypeName(ByVal lTrialID As Long, _
                                ByVal sTrialName As String, _
                                ByVal sPageCode As String, _
                                ByVal lPageID As Long, _
                                Optional ByVal lVisitID As Long, _
                                Optional ByVal sVisitCode As String = "")
'----------------------------------------------------------------------------------------------
'ASH 23/07/2002 - Formerly part of CreateTableNames routine
'Generates names for Response Value Plus Attributes Dataview tables and inserts it into
'the DataViewTables table
'----------------------------------------------------------------------------------------------
Dim itmX As MSComctlLib.ListItem
Dim sSQL As String
Dim sTableName As String

    On Error GoTo ErrHandler
    'Create Response Value With Attributes Tables
    If mbDataViewWARequired Then
        'set-up the name of a data view table
        If optVisitSeparate.Value Then
            sTableName = CreateTableName(sTrialName, sPageCode, "WA", sVisitCode)
        Else
            sTableName = CreateTableName(sTrialName, sPageCode, "WA")
        End If
        'Add Data View Name to ListView
        Set itmX = lvwTables.ListItems.Add(, , sTableName)
        itmX.SubItems(1) = "WA"
        itmX.SubItems(2) = sTrialName
        If optVisitSeparate.Value Then
            itmX.SubItems(3) = sVisitCode
        End If
        itmX.SubItems(4) = sPageCode
        
        'Add Data View Name to table DataViewTables
        sSQL = "INSERT INTO DataViewTables (DataViewName,DataViewType,ClinicalTrialName,ClinicalTrialId,CRFPageCode,CRFPageId"
        If optVisitSeparate.Value Then
            sSQL = sSQL & ",VisitCode,VisitId"
        End If
        sSQL = sSQL & ") VALUES('" & sTableName & "','WA','" & sTrialName & "'," _
            & lTrialID & ",'" & sPageCode & "'," & lPageID
        If optVisitSeparate.Value Then
            sSQL = sSQL & ",'" & sVisitCode & "'," & lVisitID
        End If
        sSQL = sSQL & ")"
        MacroADODBConnection.Execute sSQL
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CreateWATableTypeName"
End Sub

'---------------------------------------------------------------------------------
Private Function IsTableNameValid(ByVal sTableName As String) As Boolean
'---------------------------------------------------------------------------------
'ASH 25/07/2002 - Checks validity of table name. Moved from cmdChange_Click
'and adds it to mcolTableNames if valid.
'---------------------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
    sSQL = ""
    IsTableNameValid = True
    
    'Check for any changes having been made
    If msSelListViewEntryText <> sTableName Then
        'check for a non empty string
        If sTableName = "" Then
            Call DialogError("You have not entered a table name.", "Invalid Table Name")
            IsTableNameValid = False
            Exit Function
        End If
        'check the length of the edited table name
        If Len(sTableName) > mnTableNameLength Then
            Call DialogError("'" & sTableName & "' is over  " & mnTableNameLength & " characters in length.", "Invalid Table Name")
            IsTableNameValid = False
            Exit Function
        End If
        'Check the table name for characters that are not Alphanumerics or underscores
        If Not gblnValidString(sTableName, valAlpha + valNumeric + valUnderscore) Then
            Call DialogError("Table names can only contain alphanumeric and underscore characters.", "Invalid Table Name")
            IsTableNameValid = False
            Exit Function
        End If
        'Mo 2/5/2003, Bug 1568, Table Names checked for not being one of Macro's reserved word
        'Check that the table name is not a reserved word
        If Not gblnNotAReservedWord(sTableName) Then
            Call DialogError("'" & sTableName & "' is one of Macro's Reserved Words.", "Invalid Table Name")
            IsTableNameValid = False
            Exit Function
        End If
        'checks for duplicate MACRO table names before adding to collection
        If CollectionMember(mcolMacroTableNames, sTableName) = True Then
            Call DialogError("'" & sTableName & "' already exists as a table name within Macro.", "Invalid Table Name")
            IsTableNameValid = False
            Exit Function
        End If

        On Error GoTo 0
        
        On Error Resume Next
        'validate the edited table name
        mcolTableNames.Add sTableName, sTableName
        If Err.Number <> 0 Then
            Call DialogError("'" & sTableName & "' is not a unique name.", "Invalid Name")
            Err.Clear
            IsTableNameValid = False
            Exit Function
        End If
        
    End If
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.IsTableNameValid"
End Function

'------------------------------------------------------------------------------
Private Sub CreateMacroTableCollection()
'------------------------------------------------------------------------------
'ASH 25/07/2002
'Creates a collection of current Macro tables to be used in IsTableNameValid
'Called in InitializeMe routine
'------------------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
    sSQL = ""
    Set mcolMacroTableNames = Nothing
    Set mcolMacroTableNames = New Collection

        sSQL = "SELECT * FROM MACROTable "
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        Do Until rsTemp.EOF
            mcolMacroTableNames.Add rsTemp!TableName, rsTemp!TableName
            rsTemp.MoveNext
        Loop
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CreateMacroTableCollection"
End Sub

'----------------------------------------------------------------------------------
Private Sub UpdateTriggerTableAndCategoryValue()
'----------------------------------------------------------------------------------
'ASH 25/07/2002
'Updates triggers for SQL Server database types. Also updates Category Values
'Moved from InitialiseMe
'----------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    'RJCW 06/09/2001   As the table structure has changed to accommodate triggers
    'in SQL Server a check is needed to see if the db has the new structure
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
        If Not CheckSQLServerTriggerTableStructure Then
            '   RJCW 06/09/2001   Update table DataViewDetails to new structure
            AlterDataViewDetails
        End If
    End If
    
    'Changed mo 17/4/2002, check for table DataViewDetails containing
    'newly added field OutputCategoryValues
    If Not DataViewDetailsContainsOutputCategoryValues Then
        'call AddOutputCategoryValuesToDataViewDetails to add the field OutputCategoryValues to table DataViewDetails
        Call AddOutputCategoryValuesToDataViewDetails
    End If
    
    'Check for table dataViewDetails containing Special Values fields
    If Not DataViewDetailsContainsSpecialValues Then
        'call AddSpecialValuesToDataViewDetails to add the 3 Special Values fields to table DataViewDetails
        Call AddSpecialValuesToDataViewDetails
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.UpdateTriggerTableAndCategoryValue"
End Sub

'------------------------------------------------------------------------------------
Private Sub EnableControls(bEnable As Boolean)
'------------------------------------------------------------------------------------
'ASH 26/07/2002.Enables selected controls
'Moved from cmdCreateDataViews
'------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    cmdCreateViewNames.Enabled = bEnable
    cmdCreateDataViews.Enabled = bEnable
    cmdExit.Enabled = bEnable
    optVisitSeparate.Enabled = bEnable
    optVisitTogether.Enabled = bEnable
    chkResponseValueOnly.Enabled = bEnable
    chkResponseValuePlus.Enabled = bEnable

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.EnableControls"
End Sub

'------------------------------------------------------------------------------------
Private Sub SetControlsDefaultOptions()
'------------------------------------------------------------------------------------
'ASH 26/07/2002
'Sets the default options for the controls/buttons
'------------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    optVisitTogether.Value = True
    chkResponseValueOnly.Value = vbChecked
    cmdEdit.Enabled = False
    txtTableName.Enabled = False
    cmdCancel.Enabled = False
    cmdChange.Enabled = False
    'Changed Mo Morris 17/4/2002 default of category output being Values added
    optCatValues.Value = True

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.SetControlsDefaultOptions"
End Sub

'----------------------------------------------------------------------
Private Sub FixTableNameLength()
'----------------------------------------------------------------------
'ATO 4/09/2002
'checks database type and decides the max length allowed for table name
'----------------------------------------------------------------------
On Error GoTo ErrHandler
    
    If goUser.Database.DatabaseType = Oracle80 Then
        'Changed Mo Morris 21/1/2003
        'max table name length allowed for oracle is 30
        'By restricting table name to 28 the Primary key ("PK" & TableName) will never be greater than 30
        mnTableNameLength = 28 ' max table name length allowed for oracle
    Else
        mnTableNameLength = 255 ' max table name length allowed for SQL
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.FixTableNameLength"
End Sub

'------------------------------------------------------------------------
Private Sub ShowColumnHeaders(bShowVisitCol As Boolean, _
                                bShowGroupCol As Boolean)
'------------------------------------------------------------------------
'ATO 4/09/2002 removes visit and or Group column if not needed
'------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    'Do if both RO and WA checkboxes have been selected
    If mbDataViewWARequired And mbDataViewRORequired Then
        If Not bShowVisitCol Then
            lvwTables.ColumnHeaders(4).Width = 0
        End If
    End If
    
    'Do if only WA checkbox selected
    If mbDataViewWARequired And Not mbDataViewRORequired Then
        'remove Group header
        lvwTables.ColumnHeaders(6).Width = 0
        If Not bShowVisitCol Then
            lvwTables.ColumnHeaders(4).Width = 0
        End If
    End If
    
    'Do if only RO checkbox selected
    If Not mbDataViewWARequired And mbDataViewRORequired Then
        If Not bShowVisitCol Then
            lvwTables.ColumnHeaders(4).Width = 0
        End If
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.ShowColumnHeaders"
End Sub


'---------------------------------------------------------------------
Public Function ForgottenPassword(sSecurityCon As String, sUsername As String, sPassword As String, ByRef sErrMsg As String) As eDTForgottenPassword
'---------------------------------------------------------------------
'REM 06/12/02
'---------------------------------------------------------------------

    'dummy routine

End Function

'---------------------------------------------------------------------
Private Sub AssessCategoryCodes(ByVal lClinicalTrialId As Long, _
                                ByVal nVersion As Integer, _
                                ByVal lDataItemId As Long, _
                                ByRef bCatCodesNumeric As Boolean, _
                                ByRef nCatCodeLength As Integer)
'---------------------------------------------------------------------
'Sub to assess the usage of Numeric or Alpha codes
'Mo Morris  24/11/2004 - Bug 2413 Category codes in Numeric Fields option
'---------------------------------------------------------------------
Dim rsValueData As ADODB.Recordset

    On Error GoTo ErrHandler

    Set rsValueData = New ADODB.Recordset
    Set rsValueData = rsDataValues(lClinicalTrialId, nVersion, lDataItemId)
    'Loop through the category codes assesing the type (numeric or string) and the length
    nCatCodeLength = 0
    bCatCodesNumeric = True
    If rsValueData.RecordCount = 0 Then
        bCatCodesNumeric = False
        Exit Sub
    End If
    rsValueData.MoveFirst
    Do While Not rsValueData.EOF
        If Len(rsValueData!ValueCode) > nCatCodeLength Then
            nCatCodeLength = Len(rsValueData!ValueCode)
        End If
        If Not IsNumeric(rsValueData!ValueCode) Then
            bCatCodesNumeric = False
            Exit Do
        End If
        rsValueData.MoveNext
    Loop

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.AssessCategoryCodes"
End Sub

'---------------------------------------------------------------------
Private Function rsDataValues(lClinicalTrialId As Long, nVersionId As Integer, lDataItemId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
'Creates a ReadOnly recordset, that supports RecordCount.
'Mo Morris  24/11/2004 - Bug 2413 Category codes in Numeric Fields option
'---------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    sSQL = "SELECT ValueData.* FROM ValueData " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND VersionId = " & nVersionId _
        & " AND DataItemId = " & lDataItemId _
        & " AND Active = 1" _
        & " ORDER BY ValueOrder"
    Set rsDataValues = New ADODB.Recordset
    rsDataValues.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.rsDataValues"
End Function

'---------------------------------------------------------------------
Private Sub txtMissing_Change()
'---------------------------------------------------------------------
Dim sText As String

    On Error GoTo ErrHandler
    
    'mbSVMissingChanged = True
    sText = txtMissing.Text
    If sText <> "" And sText <> "-" And sText <> "-1" And sText <> "-2" _
    And sText <> "-3" And sText <> "-4" And sText <> "-5" And sText <> "-6" _
    And sText <> "-7" And sText <> "-8" And sText <> "-9" Then
        txtMissing.Text = ""
        Call DialogInformation("Special values can only be negative numbers between -1 and -9", "Special Values")
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.txtMissing_Change"
End Sub

'--------------------------------------------------------------------
Private Sub txtMissing_LostFocus()
'--------------------------------------------------------------------
Dim sText As String

    On Error GoTo ErrHandler
    
    'mbSVMissingChanged = True
    sText = txtMissing.Text
    If sText <> "" And sText <> "-1" And sText <> "-2" _
    And sText <> "-3" And sText <> "-4" And sText <> "-5" And sText <> "-6" _
    And sText <> "-7" And sText <> "-8" And sText <> "-9" Then
        txtMissing.Text = ""
        Call DialogInformation("Special values can only be negative numbers between -1 and -9", "Special Values")
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.txtMissing_LostFocus"
End Sub

'---------------------------------------------------------------------
Private Sub txtNotApplicable_Change()
'--------------------------------------------------------------------
Dim sText As String

    On Error GoTo ErrHandler
    
    'mbSVNotApplicableChanged = True
    sText = txtNotApplicable.Text
    If sText <> "" And sText <> "-" And sText <> "-1" And sText <> "-2" _
    And sText <> "-3" And sText <> "-4" And sText <> "-5" And sText <> "-6" _
    And sText <> "-7" And sText <> "-8" And sText <> "-9" Then
        txtNotApplicable.Text = ""
        Call DialogInformation("Special values can only be negative numbers between -1 and -9", "Special Values")
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.txtNotApplicable_Change"
End Sub

'--------------------------------------------------------------------
Private Sub txtNotApplicable_LostFocus()
'--------------------------------------------------------------------
Dim sText As String

    On Error GoTo ErrHandler
    
    'mbSVNotApplicableChanged = True
    sText = txtNotApplicable.Text
    If sText <> "" And sText <> "-1" And sText <> "-2" _
    And sText <> "-3" And sText <> "-4" And sText <> "-5" And sText <> "-6" _
    And sText <> "-7" And sText <> "-8" And sText <> "-9" Then
        txtNotApplicable.Text = ""
        Call DialogInformation("Special values can only be negative numbers between -1 and -9", "Special Values")
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.txtNotApplicable_LostFocus"
End Sub

'---------------------------------------------------------------------
Private Sub txtUnobtainable_Change()
'--------------------------------------------------------------------
Dim sText As String

    On Error GoTo ErrHandler
    
    'mbSVUnobtainableChanged = True
    sText = txtUnobtainable.Text
    If sText <> "" And sText <> "-" And sText <> "-1" And sText <> "-2" _
    And sText <> "-3" And sText <> "-4" And sText <> "-5" And sText <> "-6" _
    And sText <> "-7" And sText <> "-8" And sText <> "-9" Then
        txtUnobtainable.Text = ""
        Call DialogInformation("Special values can only be negative numbers between -1 and -9", "Special Values")
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.txtUnobtainable_Change"
End Sub

'--------------------------------------------------------------------
Private Sub txtUnobtainable_LostFocus()
'--------------------------------------------------------------------
Dim sText As String

    On Error GoTo ErrHandler
    
    'mbSVUnobtainableChanged = True
    sText = txtUnobtainable.Text
    If sText <> "" And sText <> "-1" And sText <> "-2" _
    And sText <> "-3" And sText <> "-4" And sText <> "-5" And sText <> "-6" _
    And sText <> "-7" And sText <> "-8" And sText <> "-9" Then
        txtUnobtainable.Text = ""
        Call DialogInformation("Special values can only be negative numbers between -1 and -9", "Special Values")
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.txtUnobtainable_LostFocus"
End Sub

'---------------------------------------------------------------------
Private Function DataViewDetailsContainsSpecialValues()
'---------------------------------------------------------------------
'Checks DataViewDetails table for column SpecialValueMissing
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
    
    DataViewDetailsContainsSpecialValues = False
    
    sSQL = "SELECT SpecialValueMissing FROM DataViewDetails"
        
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    rsTemp.Close
    Set rsTemp = Nothing

    DataViewDetailsContainsSpecialValues = True

Exit Function
ErrHandler:
    'The expected errors are as follows:-
    'With an Access database error number 3600 'No value given for one or more required parameters' gets generated.
    'With an Oracle or SQLServer database error number 3604 'Invalid column name' gets generated.
    If Err.Number - vbObjectError = 3600 Then Exit Function
    If Err.Number - vbObjectError = 3604 Then Exit Function
    Err.Raise Err.Number, , Err.Description & "|frmMenu.DataViewDetailsContainsSpecialValues"
End Function

'---------------------------------------------------------------------
Private Sub AddSpecialValuesToDataViewDetails()
'---------------------------------------------------------------------
'Adds the 3 Special Value columns to DataViewDetails table if missing
'---------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sSQL = "ALTER Table DataViewDetails ADD SpecialValueMissing VARCHAR(2)," _
            & "SpecialValueUnobtainable VARCHAR(2), SpecialValueNotApplicable VARCHAR(2)"
        MacroADODBConnection.Execute sSQL
    Case MACRODatabaseType.Oracle80
        sSQL = "ALTER Table DataViewDetails ADD SpecialValueMissing VARCHAR2(2)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataViewDetails ADD SpecialValueUnobtainable VARCHAR2(2)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataViewDetails ADD SpecialValueNotApplicable VARCHAR2(2)"
        MacroADODBConnection.Execute sSQL
    End Select


Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.AddSpecialValuesToDataViewDetails"
End Sub

'---------------------------------------------------------------------
Private Sub UpgradeTo3072()
'---------------------------------------------------------------------
' MLM 15/04/05: Created. Add a ClinicalTrialId field to the DataViewDetails table
'---------------------------------------------------------------------

Dim sSQL As String

    'if an error occurs, assume it's because the column exists already
    On Error Resume Next
    
    sSQL = "ALTER TABLE DataViewDetails ADD ClinicalTrialId "
    
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Access, MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sSQL = sSQL & "INTEGER"
    Case MACRODatabaseType.Oracle80
        sSQL = sSQL & "NUMBER(11)"
    End Select
    
    MacroADODBConnection.Execute sSQL

End Sub
