VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   Caption         =   "MACRO Batch Data Entry"
   ClientHeight    =   6675
   ClientLeft      =   3045
   ClientTop       =   4020
   ClientWidth     =   14805
   Icon            =   "frmMenuBatchDataEntry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   14805
   Begin VB.Frame fraDisplayOptions 
      Caption         =   "Display Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   120
      TabIndex        =   45
      Top             =   3600
      Width           =   2000
      Begin VB.TextBox txtBufferCount 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   47
         Top             =   360
         Width           =   800
      End
      Begin VB.CheckBox chkDisplay 
         Caption         =   "Display Buffer"
         Height          =   375
         Left            =   180
         TabIndex        =   46
         Top             =   300
         Width           =   825
      End
      Begin VB.Label lblBufferCount 
         Caption         =   "No. Entries"
         Height          =   195
         Left            =   1080
         TabIndex        =   48
         Top             =   180
         Width           =   800
      End
   End
   Begin VB.TextBox txtProgress 
      Enabled         =   0   'False
      Height          =   315
      Left            =   12060
      TabIndex        =   42
      Top             =   3780
      Width           =   2500
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "&Upload"
      Height          =   315
      Left            =   12120
      TabIndex        =   21
      Top             =   4200
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   13440
      TabIndex        =   20
      Top             =   4200
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11400
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrSystemIdleTimeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   11820
      Top             =   3960
   End
   Begin VB.Frame fraEntry 
      Caption         =   "Entry Commands"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   7620
      TabIndex        =   26
      Top             =   3660
      Width           =   4035
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   2760
         TabIndex        =   19
         Top             =   300
         Width           =   1200
      End
      Begin VB.CommandButton cmdAddClear 
         Caption         =   "Add && Clea&r"
         Height          =   315
         Left            =   1400
         TabIndex        =   18
         Top             =   300
         Width           =   1200
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   180
         TabIndex        =   17
         Top             =   300
         Width           =   1200
      End
   End
   Begin VB.Frame fraEditDelete 
      Caption         =   "Edit/Delete Commands"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   2160
      TabIndex        =   25
      Top             =   3600
      Width           =   5355
      Begin VB.CommandButton cmdClearBuffer 
         Caption         =   "Clear &Buffer"
         Height          =   315
         Left            =   100
         TabIndex        =   49
         Top             =   300
         Width           =   1150
      End
      Begin VB.CommandButton cmdCancelEdit 
         Caption         =   "Ca&ncel Edit"
         Height          =   315
         Left            =   4050
         TabIndex        =   16
         Top             =   300
         Width           =   1150
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "Chan&ge"
         Height          =   315
         Left            =   3150
         TabIndex        =   15
         Top             =   300
         Width           =   800
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   2250
         TabIndex        =   14
         Top             =   300
         Width           =   800
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   300
         Width           =   800
      End
   End
   Begin VB.Frame fraBREI 
      Caption         =   "Batch Response Entry Interface"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Left            =   120
      TabIndex        =   24
      Top             =   4500
      Width           =   14595
      Begin VB.CheckBox chkUnobtainable 
         Caption         =   "Set Status to Unobtainable"
         Height          =   435
         Left            =   13080
         TabIndex        =   51
         Top             =   480
         Width           =   1300
      End
      Begin VB.TextBox txtRepeatNumber 
         Height          =   315
         Left            =   9240
         MaxLength       =   5
         TabIndex        =   44
         Top             =   1140
         Width           =   1000
      End
      Begin VB.ComboBox cboStudy 
         Height          =   315
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   540
         Width           =   1500
      End
      Begin VB.ComboBox cboSite 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   540
         Width           =   1000
      End
      Begin VB.TextBox txtPersonId 
         Height          =   315
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   2
         Top             =   540
         Width           =   1000
      End
      Begin VB.TextBox txtLabel 
         Height          =   315
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1140
         Width           =   1000
      End
      Begin VB.ComboBox cboQuestion 
         Height          =   315
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   540
         Width           =   1500
      End
      Begin VB.ComboBox cboEForm 
         Height          =   315
         Left            =   6540
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   540
         Width           =   1500
      End
      Begin VB.TextBox txtEFormCycle 
         Height          =   315
         Left            =   8100
         MaxLength       =   5
         TabIndex        =   8
         Top             =   540
         Width           =   1000
      End
      Begin VB.TextBox txtEFormDate 
         Height          =   315
         Left            =   8160
         MaxLength       =   12
         TabIndex        =   9
         ToolTipText     =   "Enter eForm Date in dd/mm/yyyy format"
         Top             =   1140
         Width           =   1000
      End
      Begin VB.ComboBox cboVisit 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   555
         Width           =   1500
      End
      Begin VB.TextBox txtVisitCycle 
         Height          =   315
         Left            =   5370
         MaxLength       =   5
         TabIndex        =   5
         Top             =   555
         Width           =   1000
      End
      Begin VB.TextBox txtVisitDate 
         Height          =   315
         Left            =   5355
         MaxLength       =   12
         TabIndex        =   6
         ToolTipText     =   "Enter Visit Date in dd/mm/yyyy format"
         Top             =   1140
         Width           =   1000
      End
      Begin VB.TextBox txtResponse 
         Height          =   315
         Left            =   10860
         MaxLength       =   255
         TabIndex        =   11
         Top             =   540
         Width           =   2000
      End
      Begin VB.ComboBox cboCatCodes 
         Height          =   315
         Left            =   10860
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1140
         Width           =   2000
      End
      Begin VB.Label lblUnobtainable 
         Caption         =   "Response Status"
         Height          =   195
         Left            =   13080
         TabIndex        =   50
         Top             =   300
         Width           =   1300
      End
      Begin VB.Label lblRepeatNumber 
         Caption         =   "Repeat Number"
         Height          =   195
         Left            =   9240
         TabIndex        =   36
         Top             =   900
         Width           =   1200
      End
      Begin VB.Label lblResponseCodes 
         Caption         =   "Response  Codes"
         Height          =   195
         Left            =   10860
         TabIndex        =   40
         Top             =   900
         Width           =   1900
      End
      Begin VB.Label lblStudy 
         Caption         =   "Study"
         Height          =   195
         Left            =   60
         TabIndex        =   39
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label lblSite 
         Caption         =   "Site"
         Height          =   195
         Left            =   1650
         TabIndex        =   38
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label lblSubjectId 
         Caption         =   "Subject ID"
         Height          =   195
         Left            =   2775
         TabIndex        =   37
         Top             =   300
         Width           =   795
      End
      Begin VB.Label lblQuestion 
         Caption         =   "Question"
         Height          =   195
         Left            =   9240
         TabIndex        =   35
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label lblForm 
         Caption         =   "eForm"
         Height          =   195
         Left            =   6540
         TabIndex        =   34
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label lblFormCycle 
         Caption         =   "eForm Cycle"
         Height          =   195
         Left            =   8100
         TabIndex        =   33
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label lblVisit 
         Caption         =   "Visit"
         Height          =   195
         Left            =   3840
         TabIndex        =   32
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label lblVisitCycle 
         Caption         =   "Visit Cycle"
         Height          =   195
         Left            =   5400
         TabIndex        =   31
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label lblResponse 
         Caption         =   "Response Text"
         Height          =   195
         Left            =   10860
         TabIndex        =   30
         Top             =   300
         Width           =   1995
      End
      Begin VB.Label lblSubjectLabel 
         Caption         =   "or Subject Label"
         Height          =   195
         Left            =   2775
         TabIndex        =   29
         Top             =   900
         Width           =   1200
      End
      Begin VB.Label lblFormDate 
         Caption         =   "or eForm Date"
         Height          =   195
         Left            =   8160
         TabIndex        =   28
         Top             =   900
         Width           =   1000
      End
      Begin VB.Label lblVisitDate 
         Caption         =   "or Visit Date"
         Height          =   195
         Left            =   5355
         TabIndex        =   27
         Top             =   915
         Width           =   1005
      End
   End
   Begin VB.Frame fraResponseBuffer 
      Caption         =   "Response Buffer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   13000
      Begin MSComctlLib.ListView lvwBuffer 
         Height          =   3075
         Left            =   180
         TabIndex        =   22
         Top             =   300
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   5424
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.StatusBar sbrMenu 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   41
      Top             =   6300
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   900
            MinWidth        =   882
            Key             =   "UserKey"
            Object.ToolTipText     =   "Name of current user"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   900
            MinWidth        =   882
            Key             =   "RoleKey"
            Object.ToolTipText     =   "Role of current user"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   900
            MinWidth        =   882
            Key             =   "UserDatabase"
            Object.ToolTipText     =   "Name of current Database"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblProgress 
      Caption         =   "Progress"
      Height          =   195
      Left            =   11880
      TabIndex        =   43
      Top             =   3540
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFImport 
         Caption         =   "&Import file of Batch Responses"
      End
      Begin VB.Menu mnuFUpload 
         Caption         =   "&Upload contents of Response Buffer"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFGenerateSubjects 
         Caption         =   "&Generate Subjects"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnlockBatchDataEntryUpload 
         Caption         =   "Un&Lock Batch Data Entry Upload"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuFHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHUserGuide 
         Caption         =   "&User Guide"
      End
      Begin VB.Menu mnuSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAboutMacro 
         Caption         =   "&About Macro"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         frmMenuBatchDataEntry.frm
' Copyright:    InferMed Ltd. 2000. All Rights Reserved
' Author:       Mo Morris, July 2002
' Purpose:      Contains the main form of the Macro Batch Data Entry Module
'----------------------------------------------------------------------------------------'
'   Revisions:
'   Mo 27/1/2003    CreateBatchResponseDataTable removed (now part of standard database)
'                   Multi-line delete facilities added.
'                   Checkbox option to not display the ResponseBuffer added
'                   Textbox to display Number of entries in ResponseBuffer added
'                   Load Response Buffer progress messages added.
' NCJ 29 Jan 03 - Get Prolog switches from new ArezzoMemory class
' Mo 6/6/2003   Bug 1844, Changes to textbox.maxlength and textbox_lostfocus subs so that
'               PersonId is restricted to a max of 9 digits
'               Suject Label is restricted to 50 chars
'               VisitCycle is restricted to a max of 4 digits
'               VisitCycleDate is restricted to 12 chars
'               eFormCycle is restricted to a max of 4 digits
'               eFormCycleDate is restricted to 12 chars
'               Response is restricted to 255 chars
'               RepeatNumber is restricted to a max of 4 digits
' NCJ 9 Mar 04 - Changed count from integer to long in ListViewSelectedCount
'TA 20/04/2004: remove null on the response value becasue code can't handle nulls
' NCJ 30 Jun 04 - Do not set up global StudyDef here
' Mo 31/10/2006 Bug 2799, Provide Buffer Delete command button. cmdClearBuffer added.
' Mo 10/1/2007  Bug 2865, Prevent the overflow error when deleting buffer entries when
'               there are more than 32767 entries in the buffer file. Change the i variable
'               from an Integer to a LONG in cmdDelete_Click
' Mo 17/10/2007 Bug 2875, restrict the list of studies that the user can choose from
'               to studies that the user has permissions for.
'               Changes made to LoadStudyCombo.
' Mo 19/10/2007 Bug 2691, restrict the list of sites that the user can choose from to
'               sites that the user has permissions for.
'               Changes made to LoadSiteCombo, which now calls MACROUser.GetNewSubjectSites
' Mo 5/2/2008   Bug 3010. "Unlock Batch Data Entry Upload" added to file menu together with
'               mnuUnlockBatchDataEntryUpload's call to UnlockBatchUpload in modBatchDataEntry.
' Mo 26/6/2008  WO-080002 - Bug 3042 - Unobtainable status Changes
'               New controls:-
'                   lblUnobtainable
'                   chkUnobtainable
'               New subroutines:-
'                   chkUnobtainable_Click
'                   Changed subroutines:-
'                   cboEForm_Click
'                   cboQuestion_Click
'                   cboVisit_Click
'                   cmdChange_Click
'                   cmdEdit_Click
'                   Form_Load
'                   Form_Resize
'                   IsEntryComplete
'                   ClearAllSelections
'                   EnableBatchEntryControls
'                   LoadResponseBuffer
'                   AddNewBatchResponse
'----------------------------------------------------------------------------------------'

Option Explicit

'--------------------------------------------------------------------
Private Sub cboCatCodes_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'exit if no item is currently selected
    If cboCatCodes.ListIndex = -1 Then
        gsSelCatCode = ""
        Exit Sub
    End If
    
    'exit if currently selected item matches gsSelCatCode
    If Mid(cboCatCodes.Text, InStr(cboCatCodes.Text, " ") - 1) = gsSelCatCode Then Exit Sub
    
    'Store selected Category Code
    gsSelCatCode = Mid(cboCatCodes.Text, 1, InStr(cboCatCodes.Text, " ") - 1)
    
    Call IsEntryComplete
        
Exit Sub
ErrHandler:
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
Private Sub cboEForm_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'exit if no item is currently selected
    If cboEForm.ListIndex = -1 Then
        glSelCRFPageId = 0
        Exit Sub
    End If
    
    'exit if currently selected item matches glSelCRFPageId
    If cboEForm.ItemData(cboEForm.ListIndex) = glSelCRFPageId Then Exit Sub
    
    'Store selected CRFPageId
    glSelCRFPageId = cboEForm.ItemData(cboEForm.ListIndex)
    
    'Clear Question and Response controls
    cboQuestion.Clear
    cboQuestion.ListIndex = -1
    cboQuestion.Enabled = False
    glSelDataItemId = 0
    txtResponse.Text = ""
    txtResponse.Enabled = False
    gsSelResponse = ""
    cboCatCodes.Clear
    cboCatCodes.ListIndex = -1
    cboCatCodes.Enabled = False
    'Mo 27/6/2008 - WO-080002
    chkUnobtainable.Value = 0
    gnSelUnobtainable = 0
    chkUnobtainable.Enabled = False
    
    Call LoadQuestionCombo
    
    Call IsEntryComplete

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboEForm_Click")
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
Private Sub cboQuestion_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'exit if no item is currently selected
    If cboQuestion.ListIndex = -1 Then
        glSelDataItemId = 0
        Exit Sub
    End If
    
    'exit if currently selected item matches glSelDataItemId
    If cboQuestion.ItemData(cboQuestion.ListIndex) = glSelDataItemId Then Exit Sub
    
    'Store selected DataItemId
    glSelDataItemId = cboQuestion.ItemData(cboQuestion.ListIndex)
    
    'If question is of type category then populate cboCatCodes
    If DataTypeFromId(glSelTrialId, glSelDataItemId) = DataType.Category Then
        LoadCatCodes
        'clear any Response entry that might exist and disable txtResponse
        txtResponse.Text = ""
        gsSelResponse = ""
        txtResponse.Enabled = False
    Else
        'enable txtResponse
        txtResponse.Enabled = True
        'Clear any previous text response
        txtResponse.Text = ""
        'clear any CatCode entry that might exist and disable cboCatCodes
        cboCatCodes.Clear
        cboCatCodes.ListIndex = -1
        cboCatCodes.Enabled = False
        gsSelCatCode = ""
    End If
    
    'Mo 27/6/2008 - WO-080002
    chkUnobtainable.Value = 0
    gnSelUnobtainable = 0
    chkUnobtainable.Enabled = True
    
    'If Question is a RQG then enable txtRepeatNumber otherwise force a RepeatNumber of 1
    If QuestionIsRQG Then
        txtRepeatNumber.Enabled = True
        txtRepeatNumber.Text = ""
        glSelRepeatNumber = 0
    Else
        txtRepeatNumber.Enabled = False
        txtRepeatNumber.Text = 1
        glSelRepeatNumber = 1
    End If
    
    Call IsEntryComplete
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboQuestion_Click")
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
Private Sub cboSite_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler

    'exit if no item is currently selected
    If cboSite.ListIndex = -1 Then
        gsSelSite = ""
        Exit Sub
    End If
    
    'exit if currently selected item matches gsSelSite
    If cboSite.Text = gsSelSite Then Exit Sub
    
    'Store the selected Site
    gsSelSite = cboSite.Text
    
    Call IsEntryComplete

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboSite_Click")
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
Private Sub cboStudy_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler

    'exit if no item is currently selected
    If cboStudy.ListIndex = -1 Then Exit Sub
    
    'exit if currently selected item matches glSelTrialId
    If cboStudy.ItemData(cboStudy.ListIndex) = glSelTrialId Then Exit Sub
    
    'initialize the Batch Response Entry controls and current selection variables
    Call ClearAllSelections
    
    'Store selected ClinicalTrialId
    glSelTrialId = cboStudy.ItemData(cboStudy.ListIndex)
    
    'enable the Batch Response Entry controls, which are disabled until a Study has been selected
    Call EnableBatchEntryControls
    
    Call LoadSiteCombo
    Call LoadVisitCombo
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboStudy_Click")
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
Private Sub cboVisit_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'exit if no item is currently selected
    If cboVisit.ListIndex <= -1 Then
        glSelVisitId = 0
        Exit Sub
    End If
    
    'exit if currently selected item matches glSelVisitId
    If cboVisit.ItemData(cboVisit.ListIndex) = glSelVisitId Then Exit Sub
    
    'Store selected VisitId
    glSelVisitId = cboVisit.ItemData(cboVisit.ListIndex)
    
    'Clear Question and Response controls
    cboQuestion.Clear
    cboQuestion.ListIndex = -1
    cboQuestion.Enabled = False
    glSelDataItemId = 0
    txtResponse.Text = ""
    txtResponse.Enabled = False
    gsSelResponse = ""
    cboCatCodes.Clear
    cboCatCodes.ListIndex = -1
    cboCatCodes.Enabled = False
    'Mo 27/6/2008 - WO-080002
    chkUnobtainable.Value = 0
    gnSelUnobtainable = 0
    chkUnobtainable.Enabled = False
    
    Call LoadEFormCombo
    
    Call IsEntryComplete

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboVisit_Click")
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
Private Sub chkDisplay_Click()
'--------------------------------------------------------------------

    'Clear and reload the Response buffer
    txtBufferCount.Text = ""
    txtBufferCount.Refresh
    lvwBuffer.ListItems.Clear
    Call LoadResponseBuffer

End Sub

'--------------------------------------------------------------------
Private Sub chkUnobtainable_Click()
'--------------------------------------------------------------------

    If chkUnobtainable.Value = 1 Then
        gnSelUnobtainable = 1
    Else
        gnSelUnobtainable = 0
    End If
    
    Call IsEntryComplete

End Sub

'--------------------------------------------------------------------
Private Sub cmdCancelEdit_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'Clear the controls in the Batch Response Entry Interface
    cboStudy.ListIndex = -1
    glSelTrialId = 0
    ClearAllSelections
    cmdCancelEdit.Enabled = False
    cmdChange.Enabled = False
    lvwBuffer.Enabled = True
    cmdClear.Enabled = True
    cmdUpload.Enabled = True
    'Unset the Edit Taking Place boolean that is used by IsEntryComplete
    gbEditTakingPlace = False
    'unselect the selected item
    lvwBuffer.SelectedItem.Selected = False

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdCancelEdit_Click")
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
Private Sub cmdChange_Click()
'--------------------------------------------------------------------
Dim sSQL As String
Dim sResponse As String
Dim sSubject As String
Dim sVisitCycle As String
Dim seFormCycle As String

    On Error GoTo ErrHandler

    'put quotes around gsSelLabel if its not null
    If gsSelLabel <> "Null" Then
        gsSelLabel = "'" & gsSelLabel & "'"
    End If
    
    'choose between a response from txtResponse or cboCatCodes
    If gsSelResponse = "" Then
        sResponse = gsSelCatCode
    Else
        sResponse = gsSelResponse
    End If

    'Update the entry in table BatchResponseData
    'Mo 27/6/2008 - WO-080002, Unobtainable added to sql
    sSQL = "UPDATE BatchResponseData " _
        & "SET ClinicalTrialId = " & glSelTrialId _
        & ",Site = '" & gsSelSite _
        & "',PersonId = " & glSelPersonId _
        & ",SubjectLabel = " & gsSelLabel _
        & ",VisitId = " & glSelVisitId _
        & ",VisitCycleNumber = " & glSelVisitCycle _
        & ",VisitCycleDate = " & gdblSelVisitDate _
        & ",CRFPageID = " & glSelCRFPageId _
        & ",CRFPageCycleNumber = " & glSelEFormCycle _
        & ",CRFPageCycleDate = " & gdblSelEFormDate _
        & ",DataItemId = " & glSelDataItemId _
        & ",RepeatNumber = " & glSelRepeatNumber _
        & ",Response = '" & ReplaceQuotes(sResponse) & "'" _
        & ",Unobtainable = " & gnSelUnobtainable _
        & ",UserName = '" & goUser.UserName & "'" _
        & " WHERE BatchResponseId = " & lvwBuffer.SelectedItem.Text
    MacroADODBConnection.Execute sSQL
    
    'Decide between PersonId or Subject Label
    If glSelPersonId = 0 Then
        sSubject = txtLabel.Text
        'Mo 27/6/2008 - WO-080002
        lvwBuffer.SelectedItem.SubItems(14) = 0
    Else
        sSubject = CStr(glSelPersonId)
        'Mo 27/6/2008 - WO-080002
        lvwBuffer.SelectedItem.SubItems(14) = 1
    End If
    
    'Decide between Visit Cycle Number or visit Cycle Date
    If glSelVisitCycle = 0 Then
        sVisitCycle = Format(CDate(gdblSelVisitDate), "dd/mm/yyyy")
        'Mo 27/6/2008 - WO-080002
        lvwBuffer.SelectedItem.SubItems(15) = 0
    Else
        sVisitCycle = glSelVisitCycle
        'Mo 27/6/2008 - WO-080002
        lvwBuffer.SelectedItem.SubItems(15) = 1
    End If
    
    'Decide between eForm Cycle Number or eForm Cycle Date
    If glSelEFormCycle = 0 Then
        seFormCycle = Format(CDate(gdblSelEFormDate), "dd/mm/yyyy")
        'Mo 27/6/2008 - WO-080002
        lvwBuffer.SelectedItem.SubItems(16) = 0
    Else
        seFormCycle = glSelEFormCycle
        'Mo 27/6/2008 - WO-080002
        lvwBuffer.SelectedItem.SubItems(16) = 1
    End If
    
    'Put the entry back into the Response Buffer
    lvwBuffer.SelectedItem.SubItems(1) = cboStudy.Text
    lvwBuffer.SelectedItem.SubItems(2) = gsSelSite
    lvwBuffer.SelectedItem.SubItems(3) = sSubject
    lvwBuffer.SelectedItem.SubItems(4) = cboVisit.Text
    lvwBuffer.SelectedItem.SubItems(5) = sVisitCycle
    lvwBuffer.SelectedItem.SubItems(6) = cboEForm.Text
    lvwBuffer.SelectedItem.SubItems(7) = seFormCycle
    lvwBuffer.SelectedItem.SubItems(8) = cboQuestion.Text
    lvwBuffer.SelectedItem.SubItems(9) = glSelRepeatNumber
    lvwBuffer.SelectedItem.SubItems(10) = sResponse
    'Mo 27/6/2008 - WO-080002
    If gnSelUnobtainable = 1 Then
        lvwBuffer.SelectedItem.SubItems(11) = 1
    Else
        lvwBuffer.SelectedItem.SubItems(11) = ""
    End If
    lvwBuffer.SelectedItem.SubItems(12) = goUser.UserName
    
    'Clear the controls in the Batch Response Entry Interface
    cboStudy.ListIndex = -1
    glSelTrialId = 0
    Call ClearAllSelections
    'enable/disable the relevant controls
    cmdCancelEdit.Enabled = False
    cmdChange.Enabled = False
    lvwBuffer.Enabled = True
    cmdClear.Enabled = True
    cmdUpload.Enabled = True
    'Unset the Edit Taking Place boolean that is used by IsEntryComplete
    gbEditTakingPlace = False
    'unselect the selected item
    lvwBuffer.SelectedItem.Selected = False

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdChange_Click")
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
Private Sub cmdClear_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    cboStudy.ListIndex = -1
    glSelTrialId = 0
    'note that ClearAllSelections will disable the "Add" and "Add & Clear" command buttons
    Call ClearAllSelections
    Call DisableBatchEntryControls

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdClear_Click")
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
Private Sub cmdClearBuffer_Click()
'--------------------------------------------------------------------
' Mo 31/10/2006 Bug 2799, Provide Buffer Delete command button. cmdClearBuffer added.
'--------------------------------------------------------------------
Dim sMSG As String
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sMSG = "Are you sure that you want to remove all of the entries in the Batch Data Entry Response Buffer?"
    If DialogQuestion(sMSG) = vbYes Then
        Call HourglassOn
        'remove from database
        sSQL = "DELETE FROM BatchResponseData"
        MacroADODBConnection.Execute sSQL
        'remove from the Response Buffer
        lvwBuffer.ListItems.Clear
        'set txtBufferCount to zero
        txtBufferCount.Text = 0
        txtBufferCount.Refresh
        Call HourglassOff
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdClearBuffer_Click")
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
Private Sub cmdDelete_Click()
'--------------------------------------------------------------------
Dim sSQL As String
Dim olvwItem As ListItem
'Mo 10/1/2007, Bug 2865, change i from an INTEGER to a LONG
Dim i As Long

    On Error GoTo ErrHandler
    
    Call HourglassOn

    For i = lvwBuffer.ListItems.Count To 1 Step -1
        Set olvwItem = lvwBuffer.ListItems(i)
        If olvwItem.Selected Then
            'Remove from the database
            sSQL = "DELETE FROM BatchResponseData " _
                & "WHERE BatchResponseId = " & olvwItem.Text
            MacroADODBConnection.Execute sSQL
            'remove from the Response Buffer
            lvwBuffer.ListItems.Remove i
            'Decrement txtBufferCount
            txtBufferCount.Text = txtBufferCount.Text - 1
            txtBufferCount.Refresh
        End If
    Next

    'enable/disable the relevant controls
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    
    Call HourglassOff

Exit Sub
ErrHandler:
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
Dim i As Integer

    On Error GoTo ErrHandler

    'initialize the Batch Response Entry controls and current selection variables
    cboStudy.ListIndex = -1
    glSelTrialId = 0
    Call ClearAllSelections

    'set the Edit Taking Place boolean that is used by IsEntryComplete
    gbEditTakingPlace = True
    
    lvwBuffer.Enabled = False
    cmdCancelEdit.Enabled = True
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdClear.Enabled = False
    cmdUpload.Enabled = False

    'Load up the controls in the Batch Response Entry Interface with the contents of the entry to be edited
    'Load cboStudy
    For i = 0 To cboStudy.ListCount - 1
        If cboStudy.List(i) = lvwBuffer.SelectedItem.SubItems(1) Then
            cboStudy.ListIndex = i
            Exit For
        End If
    Next
    'Load cboSite
    For i = 0 To cboSite.ListCount - 1
        If cboSite.List(i) = lvwBuffer.SelectedItem.SubItems(2) Then
            cboSite.ListIndex = i
            Exit For
        End If
    Next
    'Load Subject Id/Label
    'Mo 27/6/2008 - WO-080002
    If lvwBuffer.SelectedItem.SubItems(14) = 1 Then
        txtPersonId.Text = lvwBuffer.SelectedItem.SubItems(3)
    Else
        txtLabel.Text = lvwBuffer.SelectedItem.SubItems(3)
    End If
    'Load cboVisit
    For i = 0 To cboVisit.ListCount - 1
        If cboVisit.List(i) = lvwBuffer.SelectedItem.SubItems(4) Then
            cboVisit.ListIndex = i
            Exit For
        End If
    Next
    'Load Visit Cycle/date
    'Mo 27/6/2008 - WO-080002
    If lvwBuffer.SelectedItem.SubItems(15) = 1 Then
        txtVisitCycle.Text = lvwBuffer.SelectedItem.SubItems(5)
    Else
        txtVisitDate.Text = lvwBuffer.SelectedItem.SubItems(5)
    End If
    'Load cboeForm
    For i = 0 To cboEForm.ListCount - 1
        If cboEForm.List(i) = lvwBuffer.SelectedItem.SubItems(6) Then
            cboEForm.ListIndex = i
            Exit For
        End If
    Next
    'Load eForm Cycle/date
    'Mo 27/6/2008 - WO-080002
    If lvwBuffer.SelectedItem.SubItems(16) = 1 Then
        txtEFormCycle.Text = lvwBuffer.SelectedItem.SubItems(7)
    Else
        txtEFormDate.Text = lvwBuffer.SelectedItem.SubItems(7)
    End If
    'Load cboQuestion
    For i = 0 To cboQuestion.ListCount - 1
        If cboQuestion.List(i) = lvwBuffer.SelectedItem.SubItems(8) Then
            cboQuestion.ListIndex = i
            Exit For
        End If
    Next
    'Load RepeatNumber
    txtRepeatNumber.Text = lvwBuffer.SelectedItem.SubItems(9)
    'Load Response. Having loaded cboQuestion, cboQuestion_Click would have been
    'called and txtResponse would have been enabled or disabled based on wether the
    'selected question is of type category
    If txtResponse.Enabled = True Then
        txtResponse.Text = lvwBuffer.SelectedItem.SubItems(10)
    Else
        For i = 0 To cboCatCodes.ListCount - 1
            If Mid(cboCatCodes.List(i), 1, InStr(cboCatCodes.List(i), " ") - 1) = lvwBuffer.SelectedItem.SubItems(10) Then
                cboCatCodes.ListIndex = i
                Exit For
            End If
        Next
    End If
    'Mo 27/6/2008 - WO-080002
    'Load Unobtainable status
    If lvwBuffer.SelectedItem.SubItems(11) = "1" Then
        chkUnobtainable.Value = 1
    Else
        chkUnobtainable.Value = 0
    End If

Exit Sub
ErrHandler:
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

    Call ExitBDE
    
End Sub

'--------------------------------------------------------------------
Public Sub InitialiseMe()
'--------------------------------------------------------------------
Dim oArezzoMemory As clsAREZZOMemory

    On Error GoTo ErrHandler
    
    'The following Doevents prevents command buttons ghosting during form load
    DoEvents
    
    'initialize the Batch Response Entry controls and current selection variables
    Call ClearAllSelections
    
    Call LoadStudyCombo
    
    Call LoadResponseBuffer
    
    'disable the Batch Respone Entry controls until a study has been selected
    DisableBatchEntryControls
    'Disable the command buttons
    cmdAdd.Enabled = False
    cmdAddClear.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdChange.Enabled = False
    cmdCancelEdit.Enabled = False
    
    'Initialize the Edit Taking Place boolean
    gbEditTakingPlace = False
    
    'Create and initialise a new Arezzo instance
    Set goArezzo = New Arezzo_DM
    
    ' NCJ 29 Jan 03 - Get prolog switches from new ArezzoMemory class
    Set oArezzoMemory = New clsAREZZOMemory
    Call oArezzoMemory.Load(0, goUser.CurrentDBConString)
    'Get the Prolog memory settings using GetPrologSwitches
    Call goArezzo.Init(gsTEMP_PATH, oArezzoMemory.AREZZOSwitches)
    Set oArezzoMemory = Nothing
    
    'Create instance of StudyRo object
    ' NCJ 30 Jun 04 - This is now done in modBatchDataEntry
'    Set goStudyDef = New StudyDefRO
    
    'Check for a Batch Import command line switches.
    If UCase(Left(Command, 3)) = "/BI" Then
        If ValidBatchResponseFile(gsBDCLPathAndNameOfFile) Then
            Call ImportBatchResponseFile(gsBDCLPathAndNameOfFile)
            'Call UploadBatchResponses after the import of a valid Batch Response File
            Call UploadBatchResponses
        End If
    
        Call cmdExit_Click
    End If
    
    'Check for a Batch Upload command line switches.
    If UCase(Left(Command, 3)) = "/BU" Then
        Call UploadBatchResponses
        Call cmdExit_Click
    End If

Exit Sub
ErrHandler:
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
Private Sub LoadStudyCombo()
'--------------------------------------------------------------------
'Mo 17/10/2007 - Bug 2875, rewritten, only adds studies that user has permissions for
'--------------------------------------------------------------------
Dim colstudies As Collection
Dim oStudy As Study

    On Error GoTo ErrHandler

    'Clear current contents of cboStudy
    cboStudy.Clear
    
    Set colstudies = goUser.GetNewSubjectStudies

    For Each oStudy In colstudies
        cboStudy.AddItem oStudy.StudyName
        cboStudy.ItemData(cboStudy.NewIndex) = oStudy.StudyId
    Next

Exit Sub
ErrHandler:
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

'--------------------------------------------------------------------
Private Sub cmdAddClear_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call cmdAdd_Click
    Call cmdClear_Click

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdAddClear_Click")
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

    On Error GoTo ErrHandler
    
    'If Response Buffer is not currently being displayed then set chkDisplay to checked
    'and chkDisplay_Click will be called, which in turn calls LoadResponseBuffer
    If chkDisplay.Value <> 1 Then
        chkDisplay.Value = 1
    End If
    
    'disable the "Add" and "Add & Clear" command buttons until a field
    'has been changed from the currently added record of entries
    cmdAdd.Enabled = False
    cmdAddClear.Enabled = False
    
    Call AddNewBatchResponse
    
Exit Sub
ErrHandler:
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
Private Sub cmdUpload_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call UploadBatchResponses
    
    'Clear and reload the Response buffer
    txtBufferCount.Text = ""
    txtBufferCount.Refresh
    lvwBuffer.ListItems.Clear
    Call LoadResponseBuffer

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdUpload_Click")
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
Dim colmX As MSComctlLib.ColumnHeader
Dim i As Integer

    On Error GoTo ErrHandler

    'set form size to the maximum size that can be displayed in 800*600
    Me.Width = gnMINFORMWIDTH
    Me.Height = gnMINFORMHEIGHT
    
    FormCentre Me
    
    'clear listview
    lvwBuffer.ListItems.Clear
    'add column headers with widths that are re-calculated by auto sizing later on
    'The first column is not visible and is used to store the BatchResponseId
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "Id", 0)
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "Study", 1000)
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "Site", 1000)
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "Subject", 1000)
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "Visit", 1000)
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "Visit Cycle", 1000)
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "eForm", 1000)
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "eForm Cycle", 1000)
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "Question", 1000)
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "Repeat Number", 1000)
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "Response", 1000)
    'Mo 27/6/2008 - WO-080002
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "Unobtainable", 1000)
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "User Name", 1000)
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "Upload Message", 1000)
    'The last 3 columns are not visible and are used to store "1" or "0" indicators.
    'If "Subject" contains a SubjectId column(14) is set to 1
    'If "Subject" contains a SubjectLabel column(14) is set to 0
    'If "Visit Cycle" contains a Cycle Number column(15) is set to 1
    'If "Visit Cycle" contains a Cycle Date column(15) is set to 0
    'If "eForm Cycle" contains a Cycle Number column(16) is set to 1
    'If "eForm Cycle" contains a Cycle Date column(16) is set to 0
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "Subject Id or Label", 0)
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "Visit Cycle or Date", 0)
    Set colmX = lvwBuffer.ColumnHeaders.Add(, , "eForm Cycle or Date", 0)
    
    'Mo 27/6/2008 - WO-080002
    'Auto size the column headers (1 to 13)
    For i = 2 To 14
        Call lvw_SetColWidth(lvwBuffer, i, LVSCW_AUTOSIZE_USEHEADER)
    Next i

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

'--------------------------------------------------------------------
Private Sub LoadSiteCombo()
'--------------------------------------------------------------------
'Mo 19/10/2007 - Bug 2691, rewritten, only adds sites that user has permissions for
'--------------------------------------------------------------------
Dim colSites As Collection
Dim oSite As Site

    On Error GoTo ErrHandler
    
    cboSite.Enabled = True
    'Clear current contents of cboSite
    cboSite.Clear
    
    Set colSites = goUser.GetNewSubjectSites(glSelTrialId)
    
    For Each oSite In colSites
        cboSite.AddItem oSite.Site
    Next

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadSiteCombo")
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
Private Sub LoadQuestionCombo()
'--------------------------------------------------------------------
'This sub is called when a ClinicalTrialId and a CRFPageId have been selected.
'
'It retrieves all questions within the selected eForn.
'
'If cboQuestion contains a single entry then this automaticaly
'becomes the selected Question.
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsQuestion As ADODB.Recordset

    On Error GoTo ErrHandler
    
    cboQuestion.Enabled = True
    'Clear current contents of cboQuestion
    cboQuestion.Clear
    glSelDataItemId = 0
  
    'retrieve all questions within the selected eForm
    sSQL = "SELECT DataItem.DataItemId, DataItem.DataItemCode " _
        & "FROM DataItem, CRFElement " _
        & "WHERE CRFElement.ClinicalTrialId = DataItem.ClinicalTrialId " _
        & "AND CRFElement.DataItemId = DataItem.DataItemId " _
        & "AND CRFElement.ClinicalTrialId = " & glSelTrialId & " " _
        & "AND CRFElement.CRFPageId = " & glSelCRFPageId & " " _
        & "ORDER BY DataItem.DataItemCode"
    
    Set rsQuestion = New ADODB.Recordset
    rsQuestion.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do Until rsQuestion.EOF
        cboQuestion.AddItem rsQuestion!DataItemCode
        cboQuestion.ItemData(cboQuestion.NewIndex) = rsQuestion!DataItemId
        rsQuestion.MoveNext
    Loop
    
    rsQuestion.Close
    Set rsQuestion = Nothing
    
    'If there is only one question in cboQuestion then Automatically select it
    If cboQuestion.ListCount = 1 Then
        cboQuestion.ListIndex = 0
    End If
        
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadQuestionCombo")
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
Private Sub LoadEFormCombo()
'--------------------------------------------------------------------
'This sub is called when a ClinicalTrialId and a VisitId have been selected.
'
'It retrieves all eForms within the selected Visit.
'
'If cboEForm contains a single entry then this automaticaly
'becomes the selected EForm.
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsEForm As ADODB.Recordset

    On Error GoTo ErrHandler
    
    cboEForm.Enabled = True
    'Clear current contents of cboEForm
    cboEForm.Clear
    glSelCRFPageId = 0
    
    'retrieve all eForms within the selected Visit
    sSQL = "SELECT DISTINCT CRFPage.CRFPageId, CRFPage.CRFPageCode " _
        & "FROM CRFPage, StudyVisitCRFPage " _
        & "WHERE CRFPage.ClinicalTrialId = StudyVisitCRFPage.ClinicalTrialId " _
        & "AND CRFPage.CRFPageId = StudyVisitCRFPage.CRFPageId " _
        & "AND CRFPage.ClinicalTrialId = " & glSelTrialId & " " _
        & "AND StudyVisitCRFPage.VisitId = " & glSelVisitId & " " _
        & "ORDER BY CRFPage.CRFPageCode"
    
    Set rsEForm = New ADODB.Recordset
    rsEForm.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do Until rsEForm.EOF
        cboEForm.AddItem rsEForm!CRFPageCode
        cboEForm.ItemData(cboEForm.NewIndex) = rsEForm!CRFPageId
        rsEForm.MoveNext
    Loop
    
    rsEForm.Close
    Set rsEForm = Nothing
    
    'If there is only one Form in cboEForm then Automatically select it
    If cboEForm.ListCount = 1 Then
        cboEForm.ListIndex = 0
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadEFormCombo")
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
Private Sub LoadVisitCombo()
'--------------------------------------------------------------------
'This sub is called when only a ClinicalTrialId has been selected.
'
'It retrieves all visits in the Study.
'
'If cboVisit contains a single entry then this automaticaly
'becomes the selected Visit.
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsVisit As ADODB.Recordset

    On Error GoTo ErrHandler
    
    'Check for a visit already having been opened
    If glSelVisitId > 0 Then
        Exit Sub
    End If
    
    cboVisit.Enabled = True
    'Clear current contents of cboVisit
    cboVisit.Clear

    'retrieve all visits within the selected study
    sSQL = "SELECT VisitId, VisitCode " _
        & "FROM StudyVisit " _
        & "WHERE ClinicalTrialId = " & glSelTrialId & " " _
        & "ORDER BY VisitCode"
    
    Set rsVisit = New ADODB.Recordset
    rsVisit.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do Until rsVisit.EOF
        cboVisit.AddItem rsVisit!VisitCode
        cboVisit.ItemData(cboVisit.NewIndex) = rsVisit!VisitId
        rsVisit.MoveNext
    Loop
    
    rsVisit.Close
    Set rsVisit = Nothing
    
    'If there is only one Visit in cboVisit then Automatically select it
    If cboVisit.ListCount = 1 Then
        cboVisit.ListIndex = 0
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadVisitCombo")
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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'--------------------------------------------------------------------

    Call ExitBDE

End Sub

'--------------------------------------------------------------------
Private Sub Form_Resize()
'--------------------------------------------------------------------
Dim lFormWidth As Long
Dim lFormHeight As Long
Dim l24th As Long

    On Error GoTo ErrHandler
    
    'check that form has not been minimised
    If Me.WindowState = vbMinimized Then Exit Sub
     
    If Me.Width < gnMINFORMWIDTH Then
        Me.Width = gnMINFORMWIDTH
    End If

    If Me.Height < gnMINFORMHEIGHT Then
        Me.Height = gnMINFORMHEIGHT
    End If
    
    lFormWidth = Me.ScaleWidth
    lFormHeight = Me.ScaleHeight
    
    fraResponseBuffer.Left = 100
    fraResponseBuffer.Top = 0
    fraResponseBuffer.Width = lFormWidth - 200
    fraResponseBuffer.Height = lFormHeight - 3150
    
    lvwBuffer.Left = 100
    lvwBuffer.Top = 300
    lvwBuffer.Width = fraResponseBuffer.Width - 200
    lvwBuffer.Height = fraResponseBuffer.Height - 400
    
    fraDisplayOptions.Left = 100
    fraDisplayOptions.Top = fraResponseBuffer.Height + 200
    fraDisplayOptions.Width = 2000
    fraDisplayOptions.Height = 825
    
    lblBufferCount.Left = 1100
    lblBufferCount.Top = 220
    txtBufferCount.Left = 1100
    txtBufferCount.Top = 420
    
    fraEditDelete.Left = 2200
    fraEditDelete.Top = fraDisplayOptions.Top
    fraEditDelete.Width = 5300
    fraEditDelete.Height = 825
    
    'Mo 31/10/2006 Bug 2799
    cmdClearBuffer.Left = 100
    cmdClearBuffer.Top = 300
    cmdEdit.Left = 1350
    cmdEdit.Top = 300
    cmdDelete.Left = 2250
    cmdDelete.Top = 300
    cmdChange.Left = 3150
    cmdChange.Top = 300
    cmdCancelEdit.Left = 4050
    cmdCancelEdit.Top = 300
    
    fraEntry.Left = 7600
    fraEntry.Top = fraDisplayOptions.Top
    fraEntry.Width = 4000
    fraEntry.Height = 825
    
    cmdAdd.Left = 100
    cmdAdd.Top = 300
    cmdAddClear.Left = 1400
    cmdAddClear.Top = 300
    cmdClear.Left = 2700
    cmdClear.Top = 300
    
    lblProgress.Top = fraEntry.Top - 150
    lblProgress.Left = fraResponseBuffer.Width - 2400
    txtProgress.Top = fraEntry.Top + 100
    txtProgress.Left = lblProgress.Left
    cmdUpload.Top = txtProgress.Top + 400
    cmdUpload.Left = lblProgress.Left
    cmdExit.Left = cmdUpload.Left + 1300
    cmdExit.Top = cmdUpload.Top

    fraBREI.Left = 100
    fraBREI.Top = fraDisplayOptions.Top + 900
    fraBREI.Width = lFormWidth - 200
    fraBREI.Height = 1600
    
    'Mo 26/6/2008 - WO-080002
    'The width of the controls within fraBREI is worked out as follows:-
    'There are 10 columns of controls ( that makes 11 gaps of 100 twips between them)
    'chkUnobtainable is given a fixed width of 1300, the remaining controls are allocated widths as follows
    'The width is split into 24 parts and each column is allocated a width of 2, 3 or 4 24ths
    'Note that in the following line XXX stands for 11 * 100 (for gaps) + 1300 (fixed width of chkUnobtainable)
    l24th = (fraBREI.Width - 2400) / 24
    
    lblUnobtainable.Left = fraBREI.Width - 1400
    lblUnobtainable.Top = 300
    lblUnobtainable.Width = 1300
    chkUnobtainable.Left = fraBREI.Width - 1400
    chkUnobtainable.Top = 480
    chkUnobtainable.Width = 1300
    
    lblStudy.Left = 100
    lblStudy.Top = 300
    lblStudy.Width = l24th * 3
    cboStudy.Left = 100
    cboStudy.Top = 550
    cboStudy.Width = l24th * 3
    
    lblSite.Left = lblStudy.Left + lblStudy.Width + 100
    lblSite.Top = 300
    lblSite.Width = l24th * 2
    cboSite.Left = lblSite.Left
    cboSite.Top = 550
    cboSite.Width = l24th * 2
    
    lblSubjectId.Left = lblSite.Left + lblSite.Width + 100
    lblSubjectId.Top = 300
    lblSubjectId.Width = l24th * 2
    txtPersonId.Left = lblSubjectId.Left
    txtPersonId.Top = 550
    txtPersonId.Width = l24th * 2
    lblSubjectLabel.Left = lblSubjectId.Left
    lblSubjectLabel.Top = 900
    lblSubjectLabel.Width = 1200
    txtLabel.Left = lblSubjectId.Left
    txtLabel.Top = 1150
    txtLabel.Width = l24th * 2
    
    lblVisit.Left = lblSubjectId.Left + lblSubjectId.Width + 100
    lblVisit.Top = 300
    lblVisit.Width = l24th * 3
    cboVisit.Left = lblVisit.Left
    cboVisit.Top = 550
    cboVisit.Width = l24th * 3
    
    lblVisitCycle.Left = lblVisit.Left + lblVisit.Width + 100
    lblVisitCycle.Top = 300
    lblVisitCycle.Width = l24th * 2
    txtVisitCycle.Left = lblVisitCycle.Left
    txtVisitCycle.Top = 550
    txtVisitCycle.Width = l24th * 2
    lblVisitDate.Left = lblVisitCycle.Left
    lblVisitDate.Top = 900
    lblVisitDate.Width = l24th * 2
    txtVisitDate.Left = lblVisitCycle.Left
    txtVisitDate.Top = 1150
    txtVisitDate.Width = l24th * 2
    
    lblForm.Left = lblVisitCycle.Left + lblVisitCycle.Width + 100
    lblForm.Top = 300
    lblForm.Width = l24th * 3
    cboEForm.Left = lblForm.Left
    cboEForm.Top = 550
    cboEForm.Width = l24th * 3
    
    lblFormCycle.Left = lblForm.Left + lblForm.Width + 100
    lblFormCycle.Top = 300
    lblFormCycle.Width = l24th * 2
    txtEFormCycle.Left = lblFormCycle.Left
    txtEFormCycle.Top = 550
    txtEFormCycle.Width = l24th * 2
    lblFormDate.Left = lblFormCycle.Left
    lblFormDate.Top = 900
    lblFormDate.Width = 1000
    txtEFormDate.Left = lblFormCycle.Left
    txtEFormDate.Top = 1150
    txtEFormDate.Width = l24th * 2
    
    lblQuestion.Left = lblFormCycle.Left + lblFormCycle.Width + 100
    lblQuestion.Top = 300
    lblQuestion.Width = l24th * 3
    cboQuestion.Left = lblQuestion.Left
    cboQuestion.Top = 550
    cboQuestion.Width = l24th * 3
    lblRepeatNumber.Left = lblQuestion.Left + lblQuestion.Width - 1150
    lblRepeatNumber.Top = 900
    lblRepeatNumber.Width = 1150
    txtRepeatNumber.Left = lblQuestion.Left + lblQuestion.Width - (l24th * 1)
    txtRepeatNumber.Top = 1150
    txtRepeatNumber.Width = l24th * 1
    
    lblResponse.Left = lblQuestion.Left + lblQuestion.Width + 100
    lblResponse.Top = 300
    lblResponse.Width = l24th * 4
    txtResponse.Left = lblResponse.Left
    txtResponse.Top = 550
    txtResponse.Width = l24th * 4
    lblResponseCodes.Left = lblResponse.Left
    lblResponseCodes.Top = 900
    lblResponseCodes.Width = l24th * 4
    cboCatCodes.Left = lblResponse.Left
    cboCatCodes.Top = 1150
    cboCatCodes.Width = l24th * 4
    
Exit Sub
ErrHandler:
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

    Call ExitBDE
    
End Sub

'--------------------------------------------------------------------
Private Sub lvwBuffer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call lvw_Sort(lvwBuffer, ColumnHeader)

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lvwBuffer_ColumnClick")
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
Private Sub lvwBuffer_DblClick()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'Changed Mo 13/5/2003, bug 1707
    If lvwBuffer.ListItems.Count > 0 Then
        cmdEdit_Click
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lvwBuffer_DblClick")
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
Private Sub lvwBuffer_ItemClick(ByVal Item As MSComctlLib.ListItem)
'--------------------------------------------------------------------

    On Error GoTo ErrHandler

    Select Case ListViewSelectedCount
    Case 0
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    Case 1
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
    Case Else
        cmdEdit.Enabled = False
        cmdDelete.Enabled = True
    End Select
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lvwBuffer_ItemClick")
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

    Call ExitBDE

End Sub

'--------------------------------------------------------------------
Private Sub ExitBDE()
'--------------------------------------------------------------------
' This is where we tidy up before we go home
'--------------------------------------------------------------------

    ' Ignore errors here
    On Error Resume Next
    
    ' Tidy up the import/upload things
    Call TidyUpBDE
    
    ' Only shut down the ALM if it has been started
    If Not goArezzo Is Nothing Then
        goArezzo.Finish
        Set goArezzo = Nothing
    End If
    
    Set goUser = Nothing
    
    Call ExitMACRO
    Call MACROEnd

End Sub

'--------------------------------------------------------------------
Private Sub mnuFGenerateSubjects_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    frmSubjectGenerator.Show vbModal

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuFGenerateSubjects_Click")
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
Private Sub mnuFImport_Click()
'--------------------------------------------------------------------
Dim sBatchResponseFile As String

    On Error GoTo CancelOpen
    
    With CommonDialog1
        .DialogTitle = "Import Batch Response File"
        .InitDir = gsIN_FOLDER_LOCATION
        .DefaultExt = "csv"
        .Filter = "Comma Separated Values (*.csv)|*.csv|Text file (*.txt)|*.txt"
        .CancelError = True
        .ShowOpen
  
        sBatchResponseFile = .FileName
    End With

    If ValidBatchResponseFile(sBatchResponseFile) Then
        Call ImportBatchResponseFile(sBatchResponseFile)
    End If
    
CancelOpen:

End Sub

'--------------------------------------------------------------------
Private Sub mnuFUpload_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call UploadBatchResponses
    
    'Clear and reload the Response buffer
    txtBufferCount.Text = ""
    txtBufferCount.Refresh
    lvwBuffer.ListItems.Clear
    Call LoadResponseBuffer

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuFUpload_Click")
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
Private Sub mnuHAboutMacro_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    frmAbout.Display

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuHAboutMacro_Click")
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
Private Sub mnuHUserGuide_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call MACROHelp(Me.hWnd, App.Title)

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuHUserGuide_Click")
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
Private Sub mnuUnlockBatchDataEntryUpload_Click()
'--------------------------------------------------------------------

    Call UnlockBatchUpload

End Sub

'--------------------------------------------------------------------
Private Sub txtEFormCycle_Change()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If txtEFormCycle.Text = "" Then
        'The eForm Cycle number has been cleared down
        txtEFormDate.Enabled = True
        glSelEFormCycle = 0
    ElseIf ((Not gblnValidString(txtEFormCycle.Text, valNumeric) Or Len(txtEFormCycle.Text) > 4)) Then
        'The entered eForm Cycle number is invalid
        Call DialogError("Invalid eForm Cycle Number." & vbCrLf & "A valid Cycle Number is a 1 to 4 digit integer.", "Invalid Cycle Number")
        txtEFormCycle.Text = ""
    Else
        'A valid eForm Cycle Number has been entered
        txtEFormDate.Enabled = False
        gdblSelEFormDate = 0
        glSelEFormCycle = CLng(txtEFormCycle.Text)
    End If
    
    Call IsEntryComplete

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtEFormCycle_Change")
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
Private Sub txtEFormDate_Change()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If txtEFormDate.Text = "" Then
        'The eForm Date has been cleared down
        txtEFormCycle.Enabled = True
        gdblSelEFormDate = 0
    Else
        'An eform Date is being entered, it will be validated by txtEFormDate_LostFocus
        txtEFormCycle.Enabled = False
        glSelEFormCycle = 0
    End If
   
    Call IsEntryComplete

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtEFormDate_Change")
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
Private Sub txtEFormDate_LostFocus()
'--------------------------------------------------------------------
Dim sMSG As String

    On Error GoTo ErrHandler
    
    If txtEFormDate.Text <> "" Then
        'validate newly entered eForm Date
        sMSG = ValidateDate(txtEFormDate.Text, gdblSelEFormDate)
        If sMSG > "" Then
            'The eForm Date was invalid
            Call DialogError(sMSG, "Invalid Cycle Date")
            txtEFormDate.SetFocus
        End If
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtEFormDate_LostFocus")
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
Private Sub txtLabel_Change()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If txtLabel.Text = "" Then
        'The SubjectLabel has been cleared down
        txtPersonId.Enabled = True
        gsSelLabel = "Null"
    ElseIf Not gblnValidString(txtLabel.Text, valOnlySingleQuotes) Then
        'The entered Subject Label is invalid
        Call DialogError("Invalid Subject Label." & vbCrLf & "A valid Subject Label" & gsCANNOT_CONTAIN_INVALID_CHARS, "Invalid Subject Label")
        txtLabel.Text = ""
    Else
        'A valid SubjectLabel has been entered
        txtPersonId.Enabled = False
        glSelPersonId = 0
        gsSelLabel = txtLabel.Text
    End If
    
    Call IsEntryComplete

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtLabel_Change")
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
Private Sub txtPersonId_Change()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If txtPersonId.Text = "" Then
        'The PersonId has been cleared down
        txtLabel.Enabled = True
        glSelPersonId = 0
    ElseIf ((Not gblnValidString(txtPersonId.Text, valNumeric) Or Len(txtPersonId.Text) > 9)) Then
        'The entered PersonId is invalid
        Call DialogError("Invalid PersonId." & vbCrLf & "A valid PersonId is a 1 to 9 digit integer.", "Invalid PersonId")
        txtPersonId.Text = ""
    Else
        'A valid PersonId has been entered
        txtLabel.Enabled = False
        gsSelLabel = "Null"
        glSelPersonId = CLng(txtPersonId.Text)
    End If
    
    Call IsEntryComplete

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtPersonId_Change")
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
Private Sub txtRepeatNumber_Change()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If txtRepeatNumber.Text <> "" Then
        If ((Not gblnValidString(txtRepeatNumber.Text, valNumeric) Or Len(txtRepeatNumber.Text) > 4)) Then
            'The entered Repeat Number is invalid
            Call DialogError("Invalid Repeat Number." & vbCrLf & "A Repeat Number is a 1 to 4 digit integer.", "Invalid Repeat Number")
            txtRepeatNumber.Text = ""
        Else
            'A valid Repeat Number has been entered
            glSelRepeatNumber = CLng(txtRepeatNumber.Text)
        End If
    End If
    
    Call IsEntryComplete

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtRepeatNumber_Change")
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
Private Sub txtResponse_Change()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Not gblnValidString(txtResponse.Text, valOnlySingleQuotes) Then
        'The entered Response is invalid
        Call DialogError("Invalid Response." & vbCrLf & "A valid Response" & gsCANNOT_CONTAIN_INVALID_CHARS, "Invalid Response")
        txtResponse.Text = ""
    Else
        'A valid Response has been entered
        gsSelResponse = txtResponse.Text
    End If
    
    Call IsEntryComplete

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtResponse_Change")
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
Private Sub txtVisitCycle_Change()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler

    If txtVisitCycle.Text = "" Then
        'The Visit Cycle number has been cleared down
        txtVisitDate.Enabled = True
        glSelVisitCycle = 0
    ElseIf ((Not gblnValidString(txtVisitCycle.Text, valNumeric) Or Len(txtVisitCycle.Text) > 4)) Then
        'The entered Visit Cycle number is invalid
        Call DialogError("Invalid Visit Cycle Number." & vbCrLf & "A valid Cycle Number is a 1 to 4 digit integer.", "Invalid Cycle Number")
        txtVisitCycle.Text = ""
    Else
        'A valid Visit Cycle Number has been entered
        txtVisitDate.Enabled = False
        gdblSelVisitDate = 0
        glSelVisitCycle = CLng(txtVisitCycle.Text)
    End If
    
    Call IsEntryComplete

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtVisitCycle_Change")
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
Private Sub txtVisitDate_Change()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If txtVisitDate.Text = "" Then
        'The Visit Date has been cleared down
        txtVisitCycle.Enabled = True
        gdblSelVisitDate = 0
    Else
        'A Visit Date is being entered, it will be validated by txtVisitDate_LostFocus
        txtVisitCycle.Enabled = False
        glSelVisitCycle = 0
    End If
    
    Call IsEntryComplete

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtVisitdate_Change")
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
Private Sub txtVisitdate_LostFocus()
'---------------------------------------------------------------------
Dim sMSG As String

    On Error GoTo ErrHandler
    
    If txtVisitDate.Text <> "" Then
        'validate newly entered Visit Date
        sMSG = ValidateDate(txtVisitDate.Text, gdblSelVisitDate)
        If sMSG > "" Then
            'The Visit Date was invalid
            Call DialogError(sMSG, "Invalid Cycle Date")
            txtVisitDate.SetFocus
        End If
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtVisitdate_LostFocus")
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
Private Sub LoadCatCodes()
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
    
    cboCatCodes.Enabled = True
    cboCatCodes.Clear
    sSQL = "SELECT ValueCode, ItemValue FROM ValueData " _
        & "WHERE ClinicalTrialId = " & glSelTrialId & " " _
        & "AND DataItemId = " & glSelDataItemId & " " _
        & "ORDER BY ValueOrder"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do Until rsTemp.EOF
        'Display both Category Code and Value in combo
        cboCatCodes.AddItem rsTemp!ValueCode & " - " & rsTemp!ItemValue
        'Store Category Code as itemdata
        'cboCatCodes.ItemData(cboCatCodes.NewIndex) = rsTemp!ValueCode
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadCatCodes")
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
Private Sub IsEntryComplete()
'---------------------------------------------------------------------
'This sub is run to see whether the "Add" and "Add & Clear" command
'buttons can be enabled. It checks to see that something has been entered
'in each of the required fields.
'When a Response Buffer entry is being edited it concerns itself with the "Change" command
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'Mo 26/6/2008 - WO-080002, Disable chkUnobtainable if txtResponse.text and cboCatCodes.text are blank
    If glSelDataItemId <> 0 Then
        If (gsSelResponse = "") And (gsSelCatCode = "") Then
            chkUnobtainable.Enabled = True
        Else
            chkUnobtainable.Enabled = False
        End If
        'Disable txtResponse and cboCatCodes it chkUnobtainable has been set
        If gnSelUnobtainable = 1 Then
            txtResponse.Enabled = False
            cboCatCodes.Enabled = False
        Else
            If DataTypeFromId(glSelTrialId, glSelDataItemId) = DataType.Category Then
                cboCatCodes.Enabled = True
                txtResponse.Enabled = False
            Else
                txtResponse.Enabled = True
                cboCatCodes.Enabled = False
            End If
        End If
    End If
    
    'Start by disabling the command buttons
    If gbEditTakingPlace Then
        'Disable the edits "Change" command button
        cmdChange.Enabled = False
    Else
        'Disable the "Add" and "Add & Clear" command buttons
        cmdAdd.Enabled = False
        cmdAddClear.Enabled = False
    End If
    
    If (cboSite.ListIndex = -1) Or _
    (cboQuestion.ListIndex = -1) Or _
    (cboEForm.ListIndex = -1) Or _
    (cboVisit.ListIndex = -1) Then
        Exit Sub
    End If
    
    If (txtPersonId.Text = "") And (txtLabel.Text = "") Then
        Exit Sub
    End If
    
    If (txtEFormCycle.Text = "") And (txtEFormDate.Text = "") Then
        Exit Sub
    End If
    
    If (txtVisitCycle.Text = "") And (txtVisitDate.Text = "") Then
        Exit Sub
    End If
    
    'Changed Mo 7/5/2003, a NULL response is now a valid Batch Response so txtResponse.Text
    'and cboCatCodes.ListIndex are no longer part of the IsEntryComplete check
'    If (txtResponse.Text = "") And (cboCatCodes.ListIndex = -1) Then
'        Exit Sub
'    End If
    
    If txtRepeatNumber.Text = "" Then
        Exit Sub
    End If
    
    If gbEditTakingPlace Then
        'Enable the edits "Change" command button
        cmdChange.Enabled = True
    Else
        'Enable the "Add" and "Add & Clear" command buttons
        cmdAdd.Enabled = True
        cmdAddClear.Enabled = True
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "IsEntryComplete")
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
Private Sub ClearAllSelections()
'---------------------------------------------------------------------
'This sub clears all of the Batch Response entry controls as well as
'the global variables that hold the current selection.
'This sub is called when a new study is selected and after the
'"Add and Clear" command button has been clicked.
'
'Note that if you clear a combo prio to setting ListIndex to -1 the
'cboControlName_Click sub will not be called.
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    cboSite.Clear
    cboSite.ListIndex = -1
    cboSite.Enabled = False
    gsSelSite = ""
    
    txtPersonId.Text = ""
    glSelPersonId = 0
    
    txtLabel.Text = ""
    gsSelLabel = "Null"
    
    cboVisit.Clear
    cboVisit.ListIndex = -1
    cboVisit.Enabled = False
    glSelVisitId = 0
    
    txtVisitCycle.Text = ""
    glSelVisitCycle = 0
    
    txtVisitDate.Text = ""
    gdblSelVisitDate = 0
    
    cboEForm.Clear
    cboEForm.ListIndex = -1
    cboEForm.Enabled = False
    glSelCRFPageId = 0
    
    txtEFormCycle.Text = ""
    glSelEFormCycle = 0
    
    txtEFormDate.Text = ""
    gdblSelEFormDate = 0
    
    cboQuestion.Clear
    cboQuestion.ListIndex = -1
    cboQuestion.Enabled = False
    glSelDataItemId = 0
    
    txtRepeatNumber.Text = ""
    glSelRepeatNumber = 0
    
    txtResponse.Text = ""
    gsSelResponse = ""
    
    cboCatCodes.Clear
    cboCatCodes.ListIndex = -1
    cboCatCodes.Enabled = False
    gsSelCatCode = ""
    
    'Mo 26/6/2008 - WO-080002
    chkUnobtainable.Value = 0
    gnSelUnobtainable = 0
    
    'Disable the  "Change", "Add" and "Add & Clear" command buttons
    cmdChange.Enabled = False
    cmdAdd.Enabled = False
    cmdAddClear.Enabled = False

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ClearAllSelections")
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
Private Sub DisableBatchEntryControls()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    txtPersonId.Enabled = False
    txtLabel.Enabled = False
    txtVisitCycle.Enabled = False
    txtVisitDate.Enabled = False
    txtEFormCycle.Enabled = False
    txtEFormDate.Enabled = False
    txtRepeatNumber.Enabled = False
    txtResponse.Enabled = False
    'Mo 27/6/2008 - WO-080002
    chkUnobtainable.Enabled = False

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "DisableBatchEntryControls")
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
Public Sub EnableBatchEntryControls()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    txtPersonId.Enabled = True
    txtLabel.Enabled = True
    txtVisitCycle.Enabled = True
    txtVisitDate.Enabled = True
    txtEFormCycle.Enabled = True
    txtEFormDate.Enabled = True
    txtRepeatNumber.Enabled = True
    'Mo 27/6/2008 - WO-080002
    'txtResponse.Enabled = True

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "EnableBatchEntryControls")
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
Private Sub LoadResponseBuffer()
'---------------------------------------------------------------------
'Read the contents of the BatchResponseData Table, convert Ids into
'names /codes and place into the Batch Response Buffer (lvwBuffer)
'---------------------------------------------------------------------
'Revisions
'---------------------------------------------------------------------
'TA 20/04/2004: remove null on the response value becasue code can't handle nulls


Dim sSQL As String
Dim rsBatchResponses As ADODB.Recordset
Dim sClinicalTrial As String
Dim sVisit As String
Dim seForm As String
Dim sQuestion As String
Dim sSubject As String
Dim sVisitCycle As String
Dim seFormCycle As String
Dim i As Integer
Dim itmX As MSComctlLib.ListItem
Dim lDataType As Long
Dim lCount As Long
Dim lTotal As Long

    On Error GoTo ErrHandler
    
    Call HourglassOn

    sSQL = "SELECT * FROM BatchResponseData " _
        & "ORDER BY BatchResponseId"
    Set rsBatchResponses = New ADODB.Recordset
    rsBatchResponses.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    lTotal = rsBatchResponses.RecordCount
    txtBufferCount.Text = lTotal
    txtBufferCount.Refresh
    
    If chkDisplay.Value = 1 Then
        lCount = 0
        Do Until rsBatchResponses.EOF
            lCount = lCount + 1
            txtProgress.Text = "Loading Response " & lCount & " of " & lTotal
            txtProgress.Refresh
            sClinicalTrial = TrialNameFromId(rsBatchResponses!ClinicalTrialId)
            sVisit = VisitCodeFromId(rsBatchResponses!ClinicalTrialId, rsBatchResponses!VisitId)
            seForm = CRFPageCodeFromId(rsBatchResponses!ClinicalTrialId, rsBatchResponses!CRFPageId)
            sQuestion = DataItemCodeFromId(rsBatchResponses!ClinicalTrialId, rsBatchResponses!DataItemId)
            'load entry into Buffer
            Set itmX = lvwBuffer.ListItems.Add(, , rsBatchResponses!BatchResponseId)
            itmX.SubItems(1) = sClinicalTrial
            itmX.SubItems(2) = rsBatchResponses!Site
            If rsBatchResponses!PersonId = 0 Then
                sSubject = rsBatchResponses!SubjectLabel
                'Mo 27/6/2008 - WO-080002
                itmX.SubItems(14) = 0
            Else
                sSubject = rsBatchResponses!PersonId
                'Mo 27/6/2008 - WO-080002
                itmX.SubItems(14) = 1
            End If
            itmX.SubItems(3) = sSubject
            itmX.SubItems(4) = sVisit
            If rsBatchResponses!VisitCycleNumber = 0 Then
                sVisitCycle = Format(CDate(rsBatchResponses!VisitCycleDate), "dd/mm/yyyy")
                'Mo 27/6/2008 - WO-080002
                itmX.SubItems(15) = 0
            Else
                sVisitCycle = rsBatchResponses!VisitCycleNumber
                'Mo 27/6/2008 - WO-080002
                itmX.SubItems(15) = 1
            End If
            itmX.SubItems(5) = sVisitCycle
            itmX.SubItems(6) = seForm
            If rsBatchResponses!CRFPageCycleNumber = 0 Then
                seFormCycle = Format(CDate(rsBatchResponses!CRFPageCycleDate), "dd/mm/yyyy")
                'Mo 27/6/2008 - WO-080002
                itmX.SubItems(16) = 0
            Else
                seFormCycle = rsBatchResponses!CRFPageCycleNumber
                'Mo 27/6/2008 - WO-080002
                itmX.SubItems(16) = 1
            End If
            itmX.SubItems(7) = seFormCycle
            itmX.SubItems(8) = sQuestion
            itmX.SubItems(9) = rsBatchResponses!RepeatNumber
            'if its a numeric question run ConvertStandardToLocalNum over the response
            lDataType = DataTypeFromId(rsBatchResponses!ClinicalTrialId, rsBatchResponses!DataItemId)
            Select Case lDataType
            'TA 20/04/2004: remove null on the response value becasue code can't handle nulls
            Case DataType.IntegerData, DataType.LabTest, DataType.Real
                itmX.SubItems(10) = ConvertStandardToLocalNum(RemoveNull(rsBatchResponses!Response))
            Case Else
                itmX.SubItems(10) = RemoveNull(rsBatchResponses!Response)
            End Select
            'Mo 27/6/2008 - WO-080002
            If rsBatchResponses!Unobtainable = 1 Then
                itmX.SubItems(11) = 1
            End If
            If Not IsNull(rsBatchResponses!UserName) Then
                'Mo 27/6/2008 - WO-080002
                itmX.SubItems(12) = rsBatchResponses!UserName
            End If
            If Not IsNull(rsBatchResponses!UploadMessage) Then
                'Mo 27/6/2008 - WO-080002
                itmX.SubItems(13) = rsBatchResponses!UploadMessage
            End If
            rsBatchResponses.MoveNext
        Loop
        
        If rsBatchResponses.RecordCount > 0 Then
            'unselect the last entered item
            lvwBuffer.SelectedItem.Selected = False
        End If
        
        'Set the Max Column Widths
        'Mo 27/6/2008 - WO-080002
        For i = 2 To 14
            Call lvw_SetColWidth(lvwBuffer, i, LVSCW_AUTOSIZE_USEHEADER)
        Next i
    End If
    
    'Close down the recordset
    rsBatchResponses.Close
    Set rsBatchResponses = Nothing
    
    txtProgress.Text = "Response Buffer Loaded"
    txtProgress.Refresh

    Call HourglassOff

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadResponseBuffer")
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
Private Sub AddNewBatchResponse()
'---------------------------------------------------------------------
'This sub will add the newly entered Batch Response to the BatchResponseData Table
'and to the Response Buffer (lvwBuffer).
'Note that the fields have been previously validated.
'---------------------------------------------------------------------
Dim sSQL As String
Dim lNextBatchId As Long
Dim sResponse As String
Dim sDBResponse As String
Dim sSubject As String
Dim sVisitCycle As String
Dim seFormCycle As String
Dim itmX As MSComctlLib.ListItem
Dim i As Integer
Dim sLabel As String
Dim lDataType As Long

    On Error GoTo ErrHandler
    
    lNextBatchId = GetNextBatchId
    
    'put quotes around gsSelLabel if its not null
    If gsSelLabel <> "Null" Then
        sLabel = "'" & gsSelLabel & "'"
    Else
        sLabel = gsSelLabel
    End If
    
    'choose between a response from txtResponse or cboCatCodes
    If gsSelResponse = "" Then
        sResponse = gsSelCatCode
    Else
        sResponse = gsSelResponse
    End If
    
    'if its a numeric question run ConvertLocalNumToStandard over the response
    lDataType = DataTypeFromId(glSelTrialId, glSelDataItemId)
    Select Case lDataType
    Case DataType.IntegerData, DataType.LabTest, DataType.Real
        sDBResponse = ConvertLocalNumToStandard(sResponse)
    Case Else
        sDBResponse = sResponse
    End Select

    'Add new entry to table BatchResponseData
    'Mo 27/6/2008 - WO-080002, Unobtainable added to sql
    sSQL = "INSERT INTO BatchResponseData (BatchResponseId, ClinicalTrialId, Site, PersonId, SubjectLabel, " _
        & "VisitId, VisitCycleNumber, VisitCycleDate, CRFPageID, CRFPageCycleNumber, CRFPageCycleDate, " _
        & "DataItemId, RepeatNumber, Response, UserName, Unobtainable) " _
        & "VALUES (" & lNextBatchId & "," & glSelTrialId & ",'" & gsSelSite & "'," & glSelPersonId & "," & sLabel & "," _
        & glSelVisitId & "," & glSelVisitCycle & "," & gdblSelVisitDate & "," & glSelCRFPageId & "," & glSelEFormCycle & "," & gdblSelEFormDate & "," _
        & glSelDataItemId & "," & glSelRepeatNumber & ",'" & ReplaceQuotes(sDBResponse) & "','" & goUser.UserName & "'," & gnSelUnobtainable & ")"
    MacroADODBConnection.Execute sSQL
    
    'Increment txtBufferCount
    txtBufferCount.Text = txtBufferCount.Text + 1
    txtBufferCount.Refresh
    
    'Add new entry to Response Buffer (lvwBuffer)
    Set itmX = lvwBuffer.ListItems.Add(, , lNextBatchId)
    itmX.SubItems(1) = cboStudy.Text
    itmX.SubItems(2) = gsSelSite
    'Decide between PersonId or Subject Label
    If glSelPersonId = 0 Then
        sSubject = txtLabel.Text
        'Mo 27/6/2008 - WO-080002
        itmX.SubItems(14) = 0
    Else
        sSubject = CStr(glSelPersonId)
        'Mo 27/6/2008 - WO-080002
        itmX.SubItems(14) = 1
    End If
    itmX.SubItems(3) = sSubject
    itmX.SubItems(4) = cboVisit.Text
    'Decide between Visit Cycle Number or visit Cycle Date
    If glSelVisitCycle = 0 Then
        sVisitCycle = Format(CDate(gdblSelVisitDate), "dd/mm/yyyy")
        'Mo 27/6/2008 - WO-080002
        itmX.SubItems(15) = 0
    Else
        sVisitCycle = glSelVisitCycle
        'Mo 27/6/2008 - WO-080002
        itmX.SubItems(15) = 1
    End If
    itmX.SubItems(5) = sVisitCycle
    itmX.SubItems(6) = cboEForm.Text
    'Decide between eForm Cycle Number or eForm Cycle Date
    If glSelEFormCycle = 0 Then
        seFormCycle = Format(CDate(gdblSelEFormDate), "dd/mm/yyyy")
        'Mo 27/6/2008 - WO-080002
        itmX.SubItems(16) = 0
    Else
        seFormCycle = glSelEFormCycle
        'Mo 27/6/2008 - WO-080002
        itmX.SubItems(16) = 1
    End If
    itmX.SubItems(7) = seFormCycle
    itmX.SubItems(8) = cboQuestion.Text
    itmX.SubItems(9) = glSelRepeatNumber
    itmX.SubItems(10) = sResponse
    'Mo 27/6/2008 - WO-080002
    If gnSelUnobtainable = 1 Then
        itmX.SubItems(11) = 1
    Else
        itmX.SubItems(11) = ""
    End If
    itmX.SubItems(12) = goUser.UserName
    
    'Make sure last entry is visible
    lvwBuffer.ListItems(itmX.Index).EnsureVisible

    'Set the Max Column Widths
    'Mo 27/6/2008 - WO-080002
    For i = 2 To 14
        Call lvw_SetColWidth(lvwBuffer, i, LVSCW_AUTOSIZE_USEHEADER)
    Next i
    
    'unselect the entered item
    lvwBuffer.SelectedItem.Selected = False
    
    'disable the Edit and Delete buttons
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "AddNewBatchResponse")
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
Public Function ForgottenPassword(sSecurityCon As String, sUserName As String, sPassword As String, sErrMsg As String) As Boolean
'---------------------------------------------------------------------
'dummy function for frmNewLogin to compile
'---------------------------------------------------------------------


End Function

'---------------------------------------------------------------------
Private Function ListViewSelectedCount() As Long
'---------------------------------------------------------------------
' NCJ 9 Mar 04 - Roche UAT Bug 44 - Change nCount as integer to lCount as long
'---------------------------------------------------------------------
Dim olvwItem As ListItem
Dim lCount As Long

    On Error GoTo ErrHandler
    
    lCount = 0
    For Each olvwItem In lvwBuffer.ListItems
        If olvwItem.Selected Then lCount = lCount + 1
    Next
    
    ListViewSelectedCount = lCount

Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ListViewSelectedCount")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function
