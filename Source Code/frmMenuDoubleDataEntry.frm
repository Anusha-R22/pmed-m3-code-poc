VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   Caption         =   "MACRO Double Data Entry"
   ClientHeight    =   10920
   ClientLeft      =   3045
   ClientTop       =   3780
   ClientWidth     =   11625
   Icon            =   "frmMenuDoubleDataEntry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10920
   ScaleWidth      =   11625
   Begin VB.Timer tmrRTE365 
      Left            =   9960
      Top             =   1920
   End
   Begin VB.PictureBox picDD 
      Height          =   3615
      Left            =   100
      ScaleHeight     =   3555
      ScaleWidth      =   7875
      TabIndex        =   30
      Top             =   2000
      Width           =   7935
      Begin VB.PictureBox picDDEForm 
         BackColor       =   &H00FFFFFF&
         Height          =   2445
         Left            =   0
         ScaleHeight     =   2385
         ScaleWidth      =   5700
         TabIndex        =   31
         Top             =   0
         Width           =   5760
         Begin VB.CommandButton cmdPreviousRepeat 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Previous Repeat"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   600
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.CommandButton cmdPreviousEForm 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Previous eForm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1080
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.CommandButton cmdNextEForm 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Next eForm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1080
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.CommandButton cmdEndSession 
            BackColor       =   &H00FFFFFF&
            Caption         =   "End Session"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1560
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.TextBox txtQuestionResponse 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CommandButton cmdAnotherRQGRow 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Another row"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   120
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.CommandButton cmdRepeatEForm 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Repeat eForm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   600
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.Label lblCatCodes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   34
            Top             =   1920
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label lblHeader 
            BackColor       =   &H8000000E&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1680
            Width           =   2535
         End
         Begin VB.Label lblQuestionLabel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   18
            Top             =   1200
            Visible         =   0   'False
            Width           =   2655
         End
      End
      Begin VB.VScrollBar VScroll 
         Height          =   2385
         LargeChange     =   900
         Left            =   6960
         SmallChange     =   30
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.HScrollBar HScroll 
         Height          =   300
         LargeChange     =   900
         Left            =   120
         SmallChange     =   30
         TabIndex        =   28
         Top             =   3000
         Visible         =   0   'False
         Width           =   2595
      End
   End
   Begin VB.PictureBox picFontSizing 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      ScaleHeight     =   555
      ScaleWidth      =   1275
      TabIndex        =   29
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fraESD 
      Caption         =   "Entry Session Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1750
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   10400
      Begin VB.CommandButton cmdSecondPass 
         Caption         =   "Second Pass"
         Height          =   315
         Left            =   8500
         TabIndex        =   17
         Top             =   1300
         Width           =   1800
      End
      Begin VB.CommandButton cmdFirstPass 
         Caption         =   "First Pass"
         Height          =   315
         Left            =   8500
         TabIndex        =   16
         Top             =   800
         Width           =   1800
      End
      Begin VB.ComboBox cboEFormInstance 
         Height          =   315
         Left            =   6700
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1300
         Width           =   1500
      End
      Begin VB.ComboBox cboVisitInstance 
         Height          =   315
         Left            =   6700
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   800
         Width           =   1500
      End
      Begin VB.CommandButton cmdGenerateSubject 
         Caption         =   "Generate Subject"
         Height          =   315
         Left            =   8800
         TabIndex        =   7
         Top             =   300
         Width           =   1500
      End
      Begin VB.ComboBox cboSubjects 
         Height          =   315
         Left            =   6200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   300
         Width           =   2500
      End
      Begin VB.ComboBox cboVisit 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   800
         Width           =   5000
      End
      Begin VB.ComboBox cboEForm 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1320
         Width           =   5000
      End
      Begin VB.ComboBox cboSite 
         Height          =   315
         Left            =   3200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   1500
      End
      Begin VB.ComboBox cboStudy 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   2000
      End
      Begin VB.Label lblVisitCycle 
         Caption         =   "Visit Cycle"
         Height          =   195
         Left            =   5900
         TabIndex        =   10
         Top             =   800
         Width           =   795
      End
      Begin VB.Label lblVisit 
         Caption         =   "Visit"
         Height          =   195
         Left            =   100
         TabIndex        =   8
         Top             =   800
         Width           =   300
      End
      Begin VB.Label lblFormCycle 
         Caption         =   "eForm Cycle"
         Height          =   195
         Left            =   5760
         TabIndex        =   14
         Top             =   1300
         Width           =   900
      End
      Begin VB.Label lblForm 
         Caption         =   "eForm"
         Height          =   195
         Left            =   100
         TabIndex        =   12
         Top             =   1300
         Width           =   495
      End
      Begin VB.Label lblSubjectId 
         Caption         =   "Subject Label/ID"
         Height          =   195
         Left            =   4920
         TabIndex        =   5
         Top             =   300
         Width           =   1300
      End
      Begin VB.Label lblSite 
         Caption         =   "Site"
         Height          =   195
         Left            =   2800
         TabIndex        =   3
         Top             =   300
         Width           =   300
      End
      Begin VB.Label lblStudy 
         Caption         =   "Study"
         Height          =   195
         Left            =   100
         TabIndex        =   1
         Top             =   300
         Width           =   400
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   9960
      TabIndex        =   32
      Top             =   10200
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrSystemIdleTimeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   9480
      Top             =   1920
   End
   Begin MSComctlLib.StatusBar sbrMenu 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   33
      Top             =   10545
      Width           =   11625
      _ExtentX        =   20505
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFDataVerification 
         Caption         =   "Data &Verification"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFUnlockBatchDataEntryUpload 
         Caption         =   "Un&Lock Batch Data Entry Upload"
      End
      Begin VB.Menu mnuSeparato3 
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
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuVChangeFont 
         Caption         =   "&Change eForm Font Size"
         Begin VB.Menu mnuVFont8 
            Caption         =   "Arial 8 pt"
         End
         Begin VB.Menu mnuVFont9 
            Caption         =   "Arial 9 pts"
         End
         Begin VB.Menu mnuVFont10 
            Caption         =   "Arial 10 pts"
         End
         Begin VB.Menu mnuVFont12 
            Caption         =   "Arial 12 pts"
         End
         Begin VB.Menu mnuVFont14 
            Caption         =   "Arial 14 pts"
         End
         Begin VB.Menu mnuVFont16 
            Caption         =   "Arial 16 pts"
         End
         Begin VB.Menu mnuVFont18 
            Caption         =   "Arial 18 pts"
         End
      End
      Begin VB.Menu mnuVChangeColour 
         Caption         =   "Change Colour &Scheme"
         Begin VB.Menu mnuVChangeGreys 
            Caption         =   "&Greys"
         End
         Begin VB.Menu mnuVChangeGreens 
            Caption         =   "&Greens"
         End
         Begin VB.Menu mnuVChangeBlues 
            Caption         =   "&Blues"
         End
         Begin VB.Menu mnuVChangePurples 
            Caption         =   "&Purples"
         End
         Begin VB.Menu mnuVChangeReds 
            Caption         =   "&Reds"
         End
      End
      Begin VB.Menu mnuVDisplayCategoryCodes 
         Caption         =   "&Display Category Codes"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         frmMenuDoubleDataEntry.frm
' Copyright:    InferMed Ltd. 2003. All Rights Reserved
' Author:       Mo Morris, September 2006
' Purpose:      Contains the main form of the MACRO Double Data Entry Module
'----------------------------------------------------------------------------------------'
'   Revisions:
'----------------------------------------------------------------------------------------'

Option Explicit

Private moArezzo As Arezzo_DM

Private mbEFormsBeingDisplayed As Boolean

Private mbEFormLoading As Boolean

Private mbVisitDateEform As Boolean
Private mbVisitDateEformExists As Boolean
Private mlVisitDateEformId As Long

'--------------------------------------------------------------------
Private Sub cboEForm_Click()
'--------------------------------------------------------------------
Dim sText As String

    On Error GoTo Errhandler
    
    'exit if no item is currently selected
    If cboEForm.ListIndex <= -1 Then
        glSelCRFPageId = 0
        glFirstSelCRFPageId = 0
        gsSelCRFPageCode = ""
        Exit Sub
    End If
    
    'exit if currently selected item matches glSelCRFPageId
    If cboEForm.ItemData(cboEForm.ListIndex) = glSelCRFPageId Then Exit Sub
    
    Call ClearFormInstancecbo
    Call ClearFirstSecondButtons
    tmrRTE365.Enabled = True
    
    'Store selected CRFPageId
    sText = Mid(cboEForm.Text, InStr(cboEForm.Text, "(") + 1)
    gsSelCRFPageCode = Mid(sText, 1, Len(sText) - 1)
    glSelCRFPageId = cboEForm.ItemData(cboEForm.ListIndex)
    'Store the select eForm again in glFirstSelCRFPageId
    glFirstSelCRFPageId = glSelCRFPageId
    
    Call LoadEFormInstanceCombo
    
    'enable form instance combo
    cboEFormInstance.Enabled = True

Exit Sub
Errhandler:
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
Private Sub cboEFormInstance_Click()
'--------------------------------------------------------------------
Dim bDIRDataExists As Boolean
Dim bFirstPassExists As Boolean
Dim bSecondPassExists As Boolean

    On Error GoTo Errhandler
    
    'exit if no item is currently selected
    If cboEFormInstance.ListIndex <= -1 Then
        glSelCRFPageCycleNumber = 0
        Exit Sub
    End If
    
    'exit if currently selected item matches glSelCRFPageCycleNumber
    If cboEFormInstance.ItemData(cboEFormInstance.ListIndex) = glSelCRFPageCycleNumber Then Exit Sub
    
    'Store the selected FormCycleNumber
    glSelCRFPageCycleNumber = cboEFormInstance.ItemData(cboEFormInstance.ListIndex)
    
    Call ClearFirstSecondButtons
    tmrRTE365.Enabled = True
    
    'call AssessPass to enable cmdFirstPass/cmdSecondPass
    Call AssessPass(bDIRDataExists, bFirstPassExists, bSecondPassExists)
    
    If bDIRDataExists Then
        Call DialogError("Data has already been entered for this eForm." & vbCrLf & "Double Data Entry disallowed.", "Data already entered")
        Exit Sub
    End If
    
    If Not bFirstPassExists And Not bSecondPassExists Then
        cmdFirstPass.Caption = "CREATE First Pass"
        cmdFirstPass.Enabled = True
    End If
    
    If bFirstPassExists And Not bSecondPassExists Then
        cmdFirstPass.Caption = "EDIT First Pass"
        cmdFirstPass.Enabled = True
        cmdSecondPass.Caption = "CREATE Second Pass"
        cmdSecondPass.Enabled = True
    End If
    
    If bFirstPassExists And bSecondPassExists Then
        cmdFirstPass.Caption = "EDIT First Pass"
        cmdFirstPass.Enabled = True
        cmdSecondPass.Caption = "EDIT Second Pass"
        cmdSecondPass.Enabled = True
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboEFormInstance_Click")
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

    On Error GoTo Errhandler

    'exit if no item is currently selected
    If cboSite.ListIndex = -1 Then
        gsSelSite = ""
        Exit Sub
    End If
    
    'exit if currently selected item matches gsSelSite
    If cboSite.Text = gsSelSite Then Exit Sub
    
    Call ClearSubjectscbo
    cboVisit.ListIndex = -1
    cboVisit.Enabled = False
    glSelVisitId = 0
    gsSelVisitCode = ""
    Call ClearVisitInstancecbo
    Call ClearFormcbo
    Call ClearFormInstancecbo
    Call ClearFirstSecondButtons
    tmrRTE365.Enabled = True
    
    'Store the selected Site
    gsSelSite = cboSite.Text
    
    Call LoadSubjectCombo
    
    'enable subjects combo & Generate Subject Command button
    cboSubjects.Enabled = True
    cmdGenerateSubject.Enabled = True

Exit Sub
Errhandler:
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

    On Error GoTo Errhandler

    'exit if no item is currently selected
    If cboStudy.ListIndex = -1 Then Exit Sub
    
    'exit if currently selected item matches glSelTrialId
    If cboStudy.ItemData(cboStudy.ListIndex) = glSelTrialId Then Exit Sub
    
    Call ClearSitecbo
    Call ClearSubjectscbo
    Call ClearVisitcbo
    Call ClearVisitInstancecbo
    Call ClearFormcbo
    Call ClearFormInstancecbo
    Call ClearFirstSecondButtons
    tmrRTE365.Enabled = True
    
    'Store selected ClinicalTrialName & ClinicalTrialId
    gsSelTrialName = Trim(cboStudy.Text)
    glSelTrialId = cboStudy.ItemData(cboStudy.ListIndex)
    
    'enable the Batch Response Entry controls, which are disabled until a Study has been selected
    'Call EnableBatchEntryControls
    
    Call LoadSiteCombo
    Call LoadVisitCombo
    
    'enable site combo
    cboSite.Enabled = True
    
Exit Sub
Errhandler:
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
Private Sub cboSubjects_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'exit if no item is currently selected
    If cboSubjects.ListIndex = -1 Then Exit Sub
    
    'exit if currently selected item matches glSelPersonId
    If cboSubjects.ItemData(cboSubjects.ListIndex) = glSelPersonId Then Exit Sub
    
    cboVisit.ListIndex = -1
    cboVisit.Enabled = False
    glSelVisitId = 0
    gsSelVisitCode = ""
    Call ClearVisitInstancecbo
    Call ClearFormcbo
    Call ClearFormInstancecbo
    Call ClearFirstSecondButtons
    tmrRTE365.Enabled = True
    
    'Store selected PersonId
    glSelPersonId = cboSubjects.ItemData(cboSubjects.ListIndex)
    
    'disable cmdGenerateSubject
    cmdGenerateSubject.Enabled = False
    
    'enable visits combo
    cboVisit.Enabled = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboSubjects_Click")
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
Dim sText As String

    On Error GoTo Errhandler
    
    'exit if no item is currently selected
    If cboVisit.ListIndex <= -1 Then
        glSelVisitId = 0
        gsSelVisitCode = ""
        Exit Sub
    End If
    
    'exit if currently selected item matches glSelVisitId
    If cboVisit.ItemData(cboVisit.ListIndex) = glSelVisitId Then Exit Sub
    
    Call ClearVisitInstancecbo
    Call ClearFormcbo
    Call ClearFormInstancecbo
    Call ClearFirstSecondButtons
    tmrRTE365.Enabled = True
    
    'Store selected VisitId
    sText = Mid(cboVisit.Text, InStr(cboVisit.Text, "(") + 1)
    gsSelVisitCode = Mid(sText, 1, Len(sText) - 1)
    glSelVisitId = cboVisit.ItemData(cboVisit.ListIndex)
    
    Call LoadVisitInstanceCombo
    
    Call LoadEFormCombo
    
    'enable visit instance combo
    cboVisitInstance.Enabled = True

Exit Sub
Errhandler:
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
Private Sub cboVisitInstance_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'exit if no item is currently selected
    If cboVisitInstance.ListIndex <= -1 Then
        glSelVisitCycleNumber = 0
        Exit Sub
    End If
    
    'exit if currently selected item matches glSelVisitCycleNumber
    If cboVisitInstance.ItemData(cboVisitInstance.ListIndex) = glSelVisitCycleNumber Then Exit Sub
    
    cboEForm.ListIndex = -1
    cboEForm.Enabled = False
    glSelCRFPageId = 0
    glFirstSelCRFPageId = 0
    gsSelCRFPageCode = ""
    Call ClearFormInstancecbo
    Call ClearFirstSecondButtons
    tmrRTE365.Enabled = True
    
    'Store selected VisitCycleNumber
    glSelVisitCycleNumber = cboVisitInstance.ItemData(cboVisitInstance.ListIndex)
    
    'enable forms combo
    cboEForm.Enabled = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboVisitInstance_Click")
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
Private Sub cmdAnotherRQGRow_Click(Index As Integer)
'--------------------------------------------------------------------
Dim oControl As Control
Dim oLabel As Label
Dim oTextBox As TextBox
Dim lQGroupId As Long
Dim lRQGRowHeight As Long
Dim nRepeatNumber As Integer
Dim asTag() As String
Dim sSQL As String
Dim rsGroupQuestions As ADODB.Recordset
Dim lTopCount As Long
Dim sDataItemName As String
Dim nDataItemType As Integer
Dim nDataItemLength As Integer
Dim nTabIndex As Integer
Dim nMaxRepeats As Integer
Dim nFirstQuestionIndex As Integer
Dim bGotFirstQuestionIndex As Boolean
Dim sDDResponse As String
Dim nFieldOrder As Integer
Dim rsCatCodes As ADODB.Recordset
Dim sCatCodeText As String
Dim sLineCatCodeText As String
Dim sAllCatCodeText As String
Dim nCatCodeLines As Integer
Dim lMaxWidth As Long

    'extract details from cmdAnotherRQGRow's Tag
    asTag = Split(cmdAnotherRQGRow.Item(Index).Tag, "|")
    lQGroupId = CLng(asTag(0))
    nFieldOrder = CInt(asTag(1))
    lRQGRowHeight = CLng(asTag(2))
    nRepeatNumber = CInt(asTag(3)) + 1
    lMaxWidth = CLng(asTag(4))
    nCatCodeLines = 0
    
    'Increment the height of the current DDeForm by the current RQG Row Height
    picDDEForm.Height = picDDEForm.Height + lRQGRowHeight
    
    'extract current DDeForm position from cmdAnotherRQGRow
    lTopCount = cmdAnotherRQGRow(Index).Top
    'extract current TabIndex from cmdAnotherRQGRow
    nTabIndex = cmdAnotherRQGRow(Index).TabIndex
    
    'Loop through all the controls on this DDeForm and move everything below
    'the current point down by the RQG Row Height
    For Each oControl In Me.Controls
        If oControl.Name = "lblQuestionLabel" Or oControl.Name = "txtQuestionResponse" Or oControl.Name = "cmdAnotherRQGRow" _
        Or oControl.Name = "cmdRepeatEForm" Or oControl.Name = "cmdNextEForm" Or oControl.Name = "cmdEndSession" _
        Or oControl.Name = "cmdPreviousRepeat" Or oControl.Name = "cmdPreviousEForm" Or oControl.Name = "lblCatCodes" Then
            If oControl.Top > lTopCount Then
                oControl.Top = oControl.Top + lRQGRowHeight
            End If
        End If
    Next
    
    'create recordset of the individual questions within this RQG
    sSQL = "SELECT DataItemId, QOrder FROM QGroupQuestion" _
        & " WHERE ClinicalTrialId = " & glSelTrialId _
        & " AND VersionId = 1" _
        & " AND QGroupId = " & lQGroupId _
        & " ORDER BY QOrder"
    Set rsGroupQuestions = New ADODB.Recordset
    rsGroupQuestions.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    bGotFirstQuestionIndex = False
    'loop through the RQG questions, placing the next row on the DDeForm
    Do Until rsGroupQuestions.EOF
        'Skip derived questions
        If Not QuestionIsDerived(glSelTrialId, rsGroupQuestions!DataItemId) Then
            Call GetDataItemDetails(glSelTrialId, rsGroupQuestions!DataItemId, sDataItemName, nDataItemType, nDataItemLength)
            'Create and position a question label
            Load Me.lblQuestionLabel(gnLabelIndex)
            Set oLabel = Me.lblQuestionLabel(gnLabelIndex)
            gnLabelIndex = gnLabelIndex + 1
            sDataItemName = " [" & nFieldOrder & "." & rsGroupQuestions!QOrder & "." & nRepeatNumber & "] " & sDataItemName
            With oLabel
                .Caption = sDataItemName
                .Top = lTopCount
                .Left = 100
                .Width = picFontSizing.TextWidth(sDataItemName & "   ")
                .Height = (gnTextBoxHeight + 100)
                .Visible = True
                .TabIndex = nTabIndex
            End With
            nTabIndex = nTabIndex + 1
            
            'Create and position a question textbox
            'Store the index of the first question textbox to be created for current row
            If Not bGotFirstQuestionIndex Then
                nFirstQuestionIndex = gnTextBoxIndex
                bGotFirstQuestionIndex = True
            End If
            Load Me.txtQuestionResponse(gnTextBoxIndex)
            Set oTextBox = Me.txtQuestionResponse(gnTextBoxIndex)
            sDDResponse = GetCurrentDDResponse(gnPassNumber, rsGroupQuestions!DataItemId, nRepeatNumber)
            With oTextBox
                .Text = sDDResponse
                .Top = lTopCount
                .Left = 600 + gn50CharWidth
                .Width = picFontSizing.TextWidth(String(nDataItemLength + 3, "_"))
                .Height = (gnTextBoxHeight + 100)
                .Tag = rsGroupQuestions!DataItemId & "|" & nFieldOrder & "|" & nRepeatNumber & "|" & rsGroupQuestions!QOrder & "|" & nDataItemType
                .Enabled = True
                .Visible = True
                .TabIndex = nTabIndex
            End With
            nTabIndex = nTabIndex + 1
            
            'perform an initial save of a blank response to the DoubleData table
            Call SaveUpdateResponse(gnTextBoxIndex)
            
            'Display category codes if required
            If gbRegDisplayCategoryCodes Then
                'If the question is a category code question
                If nDataItemType = eDataType.Category Then
                    Load Me.lblCatCodes(gnCatCodeIndex)
                    Set oLabel = Me.lblCatCodes(gnCatCodeIndex)
                    gnCatCodeIndex = gnCatCodeIndex + 1
                    'get a recordset of this questions category codes
                    Set rsCatCodes = New ADODB.Recordset
                    Set rsCatCodes = GetCatCodes(rsGroupQuestions!DataItemId)
                    sAllCatCodeText = ""
                    sLineCatCodeText = ""
                    nCatCodeLines = 1
                    Do Until rsCatCodes.EOF
                        sCatCodeText = "[" & rsCatCodes!ValueCode & " - " & rsCatCodes!ItemValue & "] "
                        'can this category code fit onto current line
                        If (picFontSizing.TextWidth(sLineCatCodeText) + picFontSizing.TextWidth(sCatCodeText) + 600 + gn50CharWidth) > lMaxWidth Then
                            'add linefeed then caption and increment line counter
                            sAllCatCodeText = sAllCatCodeText & vbNewLine & sCatCodeText
                            nCatCodeLines = nCatCodeLines + 1
                            sLineCatCodeText = sCatCodeText
                        Else
                            'just add caption
                            sAllCatCodeText = sAllCatCodeText & sCatCodeText
                            sLineCatCodeText = sLineCatCodeText & sCatCodeText
                        End If
                        rsCatCodes.MoveNext
                    Loop
                    rsCatCodes.Close
                    Set rsCatCodes = Nothing
                    With oLabel
                        .Caption = sAllCatCodeText
                        .Top = lTopCount + (1.5 * gnTextBoxHeight)
                        .Left = 600 + gn50CharWidth
                        .Width = lMaxWidth - (600 + gn50CharWidth)
                        .Height = nCatCodeLines * gnTextBoxHeight
                        .Visible = True
                    End With
                End If
            End If
            
            gnTextBoxIndex = gnTextBoxIndex + 1
            
            'increment lTopCount
            lTopCount = lTopCount + (2 * gnTextBoxHeight) + (nCatCodeLines * gnTextBoxHeight)
            nCatCodeLines = 0
        End If
        rsGroupQuestions.MoveNext
    Loop
    
    'Based on the Max number of repeats for this RQG
    'either reposition cmdAnotherRQGRow below the newly added rows
    'or remove cmdAnotherRQGRow because the max number of repeats has been reached
    'reposition the another row button
    nMaxRepeats = GetRQGMaxRepeats(glSelTrialId, glSelCRFPageId, lQGroupId)
    If (nMaxRepeats - nRepeatNumber) > 1 Then
        cmdAnotherRQGRow.Item(Index).Top = lTopCount
        'increment the RepeatNumber to be used in another row
        cmdAnotherRQGRow.Item(Index).Tag = lQGroupId & "|" & nFieldOrder & "|" & lRQGRowHeight & "|" & nRepeatNumber & "|" & lMaxWidth
    ElseIf (nMaxRepeats - nRepeatNumber) = 1 Then
        cmdAnotherRQGRow.Item(Index).Top = lTopCount
        'increment the RepeatNumber to be used on last row minus space for the non required "Another Row" command button
        'when "Another Row" is clicked for the last time
        cmdAnotherRQGRow.Item(Index).Tag = lQGroupId & "|" & nFieldOrder & "|" & (lRQGRowHeight - (2 * gnTextBoxHeight)) & "|" & nRepeatNumber & "|" & lMaxWidth
    Else
        Unload cmdAnotherRQGRow.Item(Index)
    End If
    
    'call DDeFormChecks which checks for increasing width and/or height
    Call DDeFormChecks
    
    'call DDeFormScrollbars which checks the need for scrollbars
    Call DDeFormScrollbars
    
    'set the focus to the first question of the newly added row
    txtQuestionResponse(nFirstQuestionIndex).SetFocus

End Sub

'---------------------------------------------------------------------
Private Sub cmdAnotherRQGRow_GotFocus(Index As Integer)
'---------------------------------------------------------------------

    Call CheckVerticalScroll(cmdAnotherRQGRow(Index).Top, cmdAnotherRQGRow(Index).Height)
    Call CheckHorizontalScroll(cmdAnotherRQGRow(Index).Left, cmdAnotherRQGRow(Index).Width)
    
End Sub

'--------------------------------------------------------------------
Private Sub cmdEndSession_Click(Index As Integer)
'--------------------------------------------------------------------

    Call ClearDDeForm
    
    'set the eForms being displayed flag to false
    mbEFormsBeingDisplayed = False
    
    Call ClearSitecbo
    Call ClearSubjectscbo
    Call ClearVisitcbo
    Call ClearVisitInstancecbo
    Call ClearFormcbo
    Call ClearFormInstancecbo
    Call ClearFirstSecondButtons
    cboStudy.Enabled = True
    cboStudy.ListIndex = -1
    glSelTrialId = 0

End Sub

'---------------------------------------------------------------------
Private Sub cmdEndSession_GotFocus(Index As Integer)
'---------------------------------------------------------------------

    Call CheckVerticalScroll(cmdEndSession(Index).Top, cmdEndSession(Index).Height)
    Call CheckHorizontalScroll(cmdEndSession(Index).Left, cmdEndSession(Index).Width)
    
End Sub

'--------------------------------------------------------------------
Private Sub cmdExit_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    ' Only shut down the ALM if it has been started
    If Not moArezzo Is Nothing Then
        moArezzo.Finish
        Set moArezzo = Nothing
    End If

    Call ExitMACRO
    Call MACROEnd

Exit Sub
Errhandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdExit_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Public Sub InitialiseMe()
'--------------------------------------------------------------------
Dim oArezzoMemory As clsAREZZOMemory
Dim ltemp As Long

    On Error GoTo Errhandler

    'The following Doevents prevents command buttons ghosting during form load
    DoEvents
    
    'Clear and disable combos etc.
    glSelTrialId = 0
    Call ClearSitecbo
    Call ClearSubjectscbo
    Call ClearVisitcbo
    Call ClearVisitInstancecbo
    Call ClearFormcbo
    Call ClearFormInstancecbo
    Call ClearFirstSecondButtons
    Call ClearDDeForm
    
    'set the eForms being displayed flag to false
    mbEFormsBeingDisplayed = False

    Call LoadStudyCombo
        
    'Create and initialise a new Arezzo instance
    Set moArezzo = New Arezzo_DM
    
    ' NCJ 29 Jan 03 - Get prolog switches from new ArezzoMemory class
    Set oArezzoMemory = New clsAREZZOMemory
    Call oArezzoMemory.Load(0, goUser.CurrentDBConString)
    'Get the Prolog memory settings using GetPrologSwitches
    Call moArezzo.Init(gsTEMP_PATH, oArezzoMemory.AREZZOSwitches)
    Set oArezzoMemory = Nothing
    
    'Set the appropriate Font Size menu item
    Select Case gnRegDDFontSize
    Case 8
        mnuVFont8.Checked = True
    Case 9
        mnuVFont9.Checked = True
    Case 10
        mnuVFont10.Checked = True
    Case 12
        mnuVFont12.Checked = True
    Case 14
        mnuVFont14.Checked = True
    Case 16
        mnuVFont16.Checked = True
    Case 18
        mnuVFont18.Checked = True
    End Select
    
    'Set the appropriate Colour Scheme menu item
    Select Case gsRegDDColourScheme
    Case "Greys"
        mnuVChangeGreys.Checked = True
    Case "Greens"
        mnuVChangeGreens.Checked = True
    Case "Blues"
        mnuVChangeBlues.Checked = True
    Case "Purples"
        mnuVChangePurples.Checked = True
    Case "Reds"
        mnuVChangeReds.Checked = True
    End Select
    
    'Set the appropriate FDisplay Category Codes menu item
    If gbRegDisplayCategoryCodes Then
        mnuVDisplayCategoryCodes.Checked = True
    Else
        mnuVDisplayCategoryCodes.Checked = False
    End If
    
Exit Sub
Errhandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "InitialiseMe", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Public Sub CheckUserRights()
'--------------------------------------------------------------------
' Dummy routine which gets called during MACRO initialisation
'--------------------------------------------------------------------

End Sub

'--------------------------------------------------------------------
Private Sub cmdFirstPass_Click()
'--------------------------------------------------------------------

    gnPassNumber = ePassNumber.First
    
    Call DisableCombos
    
    Call AssessVisitDates
    
    Call BuildDDeForm

End Sub

'--------------------------------------------------------------------
Private Sub cmdGenerateSubject_Click()
'--------------------------------------------------------------------
Dim oStudyDef As StudyDefRO
Dim oSubject As StudySubject
Dim sCountry As String
Dim sToken As String
Dim sErrMsg As String

    'Create the required Study object
    Set oStudyDef = New StudyDefRO
    oStudyDef.Load gsADOConnectString, glSelTrialId, 1, moArezzo
    
    sCountry = goUser.GetAllSites.Item(gsSelSite).CountryName
    
    'Place a lock on new subjects being generated for the selected Study & Site
    sToken = LockSubjectGeneration(glSelTrialId, gsSelSite, sErrMsg)
    'Check for lock having worked
    If sToken = "" Then
        'Lock failed
        DialogError "Unable to Generate subject." & vbCrLf & sErrMsg
        Exit Sub
    End If

    'Generate a new subject for the specified study and site
    Set oSubject = oStudyDef.NewSubject(gsSelSite, goUser.UserName, sCountry, goUser.UserNameFull, goUser.UserRole)
    
    'Add new subject to cboSubjects and select it
    cboSubjects.AddItem "(" & oSubject.PersonID & ")"
    cboSubjects.ItemData(cboSubjects.NewIndex) = oSubject.PersonID
    cboSubjects.ListIndex = cboSubjects.NewIndex
    
    'Clean up after generating subject
    oStudyDef.Terminate
    Set oStudyDef = Nothing
    Set oSubject = Nothing
    
    Call UnLockSubjectGeneration(glSelTrialId, gsSelSite, sToken)
    
End Sub

'--------------------------------------------------------------------
Private Sub cmdNextEForm_Click(Index As Integer)
'--------------------------------------------------------------------

    'extract the CRFPageId of the next eForm
    glSelCRFPageId = CLng(cmdNextEForm(1).Tag)
    gsSelCRFPageCode = CRFPageCodeFromId(glSelTrialId, glSelCRFPageId)
    glSelCRFPageCycleNumber = 1
    mbVisitDateEform = False
    Call BuildDDeForm

End Sub

'---------------------------------------------------------------------
Private Sub cmdNextEForm_GotFocus(Index As Integer)
'---------------------------------------------------------------------

    Call CheckVerticalScroll(cmdNextEForm(Index).Top, cmdNextEForm(Index).Height)
    Call CheckHorizontalScroll(cmdNextEForm(Index).Left, cmdNextEForm(Index).Width)
    
End Sub

'--------------------------------------------------------------------
Private Sub cmdPreviousEForm_Click(Index As Integer)
'--------------------------------------------------------------------

    'extract the CRFPageId of the prev eForm
    glSelCRFPageId = CLng(cmdPreviousEForm(1).Tag)
    gsSelCRFPageCode = CRFPageCodeFromId(glSelTrialId, glSelCRFPageId)
    glSelCRFPageCycleNumber = 1
    mbVisitDateEform = False
    Call BuildDDeForm

End Sub

'---------------------------------------------------------------------
Private Sub cmdPreviousEForm_GotFocus(Index As Integer)
'---------------------------------------------------------------------

    Call CheckVerticalScroll(cmdPreviousEForm(Index).Top, cmdPreviousEForm(Index).Height)
    Call CheckHorizontalScroll(cmdPreviousEForm(Index).Left, cmdPreviousEForm(Index).Width)
    
End Sub

'--------------------------------------------------------------------
Private Sub cmdPreviousRepeat_Click(Index As Integer)
'--------------------------------------------------------------------

    'Decrement the eForm Cycle Number
    glSelCRFPageCycleNumber = glSelCRFPageCycleNumber - 1
    mbVisitDateEform = False
    Call BuildDDeForm

End Sub

'---------------------------------------------------------------------
Private Sub cmdPreviousRepeat_GotFocus(Index As Integer)
'---------------------------------------------------------------------

    Call CheckVerticalScroll(cmdPreviousRepeat(Index).Top, cmdPreviousRepeat(Index).Height)
    Call CheckHorizontalScroll(cmdPreviousRepeat(Index).Left, cmdPreviousRepeat(Index).Width)
    
End Sub

'--------------------------------------------------------------------
Private Sub cmdRepeatEForm_Click(Index As Integer)
'--------------------------------------------------------------------

    'Increment the eForm Cycle Number
    glSelCRFPageCycleNumber = glSelCRFPageCycleNumber + 1
    mbVisitDateEform = False
    Call BuildDDeForm

End Sub

'---------------------------------------------------------------------
Private Sub cmdRepeatEForm_GotFocus(Index As Integer)
'---------------------------------------------------------------------

    Call CheckVerticalScroll(cmdRepeatEForm(Index).Top, cmdRepeatEForm(Index).Height)
    Call CheckHorizontalScroll(cmdRepeatEForm(Index).Left, cmdRepeatEForm(Index).Width)
    
End Sub

'--------------------------------------------------------------------
Private Sub cmdSecondPass_Click()
'--------------------------------------------------------------------

    gnPassNumber = ePassNumber.Second
    
    Call AssessVisitDates
    
    Call BuildDDeForm

End Sub

'--------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------
Dim oRequest As Object

    On Error GoTo Errhandler
    
    Call GetRegistrySettings
    
    mbEFormLoading = True
    
    'Initially always set form left,top,width & height to the non-max saved settings
    Me.Left = gnRegFormLeft
    Me.Top = gnRegFormTop
    Me.Width = gnRegFormWidth
    Me.Height = gnRegFormHeight
    Me.WindowState = gnRegFormWindowState
    If gnRegFormWindowState = 2 Then
        'If the window is supposed to be in a max state, then
        'set it to its non-max dimensions and then max it
        Me.WindowState = gnRegFormWindowState
    End If
    
    mbEFormLoading = False
    
    picDDEForm.Visible = False
    
    tmrRTE365.Enabled = False
    tmrRTE365.Interval = 10

Exit Sub
Errhandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Load", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub Form_Resize()
'--------------------------------------------------------------------
Dim nFormWidth As Integer
Dim nFormHeight As Integer

    On Error GoTo Errhandler

    'check that form has not been minimised
    If Me.WindowState = vbMinimized Then Exit Sub
    
    'force a minimum hieght for the form
    If Me.Height < mnMINFORMHEIGHT Then Me.Height = mnMINFORMHEIGHT
    
    'force a minimum width for the form
    If Me.Width < mnMINFORMWIDTH Then Me.Width = mnMINFORMWIDTH
    
    nFormWidth = Me.Width
    nFormHeight = Me.Height

    picDD.Width = nFormWidth - 300
    picDD.Height = nFormHeight - 3200
    picDDEForm.Left = 0
    
    'Assess the need for vertical scrollbar
    If picDDEForm.Height > picDD.Height Then
        Set VScroll.Container = picDD
        VScroll.Max = (picDDEForm.Height - picDD.Height) / 10
        VScroll.Top = 0
        If picDDEForm.Width > picDD.Width Then
            VScroll.Height = picDD.ScaleHeight - HScroll.Height
        Else
            VScroll.Height = picDD.ScaleHeight
        End If
        VScroll.Left = picDD.ScaleWidth - VScroll.Width
        VScroll.LargeChange = CInt(((VScroll.Max / 2) / 10) + 1)
        VScroll.SmallChange = CInt(((VScroll.Max / 10) / 10) + 1)
        VScroll.Value = 0
        VScroll.Visible = True
    Else
        VScroll.Visible = False
    End If
    
    'Assess the need for horizontal scrollbar
    If picDDEForm.Width > picDD.Width Then
        Set HScroll.Container = picDD
        HScroll.Max = (picDDEForm.Width - picDD.Width) / 10
        HScroll.Top = picDD.ScaleHeight - HScroll.Height
        If VScroll.Visible = True Then
            HScroll.Width = picDD.ScaleWidth - VScroll.Width
        Else
            HScroll.Width = picDD.ScaleWidth
        End If
        HScroll.Left = 0
        HScroll.LargeChange = CInt(((HScroll.Max / 2) / 10) + 1)
        HScroll.SmallChange = CInt(((HScroll.Max / 10) / 10) + 1)
        HScroll.Value = 0
        HScroll.Visible = True
    Else
        HScroll.Visible = False
    End If
    
    cmdExit.Top = nFormHeight - 1130
    cmdExit.Left = nFormWidth - 1400

    'store changed settings
    If Not mbEFormLoading Then
        If Me.WindowState = 2 Then
            gnRegFormLeftMax = Me.Left
            gnRegFormTopMax = Me.Top
            gnRegFormWidthMax = Me.Width
            gnRegFormHeightMax = Me.Height
        Else
            gnRegFormLeft = Me.Left
            gnRegFormTop = Me.Top
            gnRegFormWidth = Me.Width
            gnRegFormHeight = Me.Height
        End If
        gnRegFormWindowState = Me.WindowState
    End If
    
Exit Sub
Errhandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Resize", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'save Registry settings
    Call SaveRegistrySettings
    
    ' Only shut down the ALM if it has been started
    If Not moArezzo Is Nothing Then
        moArezzo.Finish
        Set moArezzo = Nothing
    End If
    
    Call ExitMACRO
    Call MACROEnd

Exit Sub
Errhandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Unload", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub HScroll_Change()
'--------------------------------------------------------------------

    picDDEForm.Left = CSng(-HScroll.Value) * 10

End Sub

'--------------------------------------------------------------------
Private Sub HScroll_Scroll()
'--------------------------------------------------------------------

    picDDEForm.Left = CSng(-HScroll.Value) * 10

End Sub

'--------------------------------------------------------------------
Private Sub mnuFDataVerification_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    frmDataVerification.Show vbModal

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuFDataVerification_Click")
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

    Call cmdExit_Click

End Sub

'--------------------------------------------------------------------
Private Sub mnuFUnlockBatchDataEntryUpload_Click()
'--------------------------------------------------------------------

    Call UnlockBatchUpload

End Sub

'--------------------------------------------------------------------
Private Sub mnuHAboutMacro_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    frmAbout.Display

Exit Sub
Errhandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuHAboutMacro_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub mnuHUserGuide_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    Call MACROHelp(Me.hWnd, App.Title)

Exit Sub
Errhandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuHUserGuide_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Public Function ForgottenPassword(sSecurityCon As String, sUserName As String, sPassword As String, sErrMsg As String) As Boolean
'---------------------------------------------------------------------
'dummy function for frmNewLogin to compile
'---------------------------------------------------------------------


End Function

'--------------------------------------------------------------------
Private Sub LoadStudyCombo()
'--------------------------------------------------------------------
Dim colstudies As Collection
Dim oStudy As Study

    On Error GoTo Errhandler

    'Clear current contents of cboStudy
    cboStudy.Clear
    glSelTrialId = 0
    
    Set colstudies = goUser.GetNewSubjectStudies

    For Each oStudy In colstudies
        cboStudy.AddItem oStudy.StudyName
        cboStudy.ItemData(cboStudy.NewIndex) = oStudy.StudyId
    Next

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

'--------------------------------------------------------------------
Private Sub LoadSiteCombo()
'--------------------------------------------------------------------
Dim colSites As Collection
Dim oSite As Site

    On Error GoTo Errhandler
    
    'Clear current contents of cboSite
    cboSite.Clear
    gsSelSite = ""
    
    Set colSites = goUser.GetNewSubjectSites(glSelTrialId)
    
    For Each oSite In colSites
        cboSite.AddItem oSite.Site
    Next
    
Exit Sub
Errhandler:
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
Private Sub LoadVisitCombo()
'--------------------------------------------------------------------
'This sub is called when only a ClinicalTrialId has been selected.
'
'It retrieves all visits in the Study.
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsVisit As ADODB.Recordset

    On Error GoTo Errhandler
    
    'Clear current contents of cboVisit
    cboVisit.Clear
    glSelVisitId = 0
    gsSelVisitCode = ""

    'retrieve all visits within the selected study
    sSQL = "SELECT VisitId, VisitCode, VisitName, VisitOrder " _
        & "FROM StudyVisit " _
        & "WHERE ClinicalTrialId = " & glSelTrialId & " " _
        & "ORDER BY VisitOrder"
    
    Set rsVisit = New ADODB.Recordset
    rsVisit.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do Until rsVisit.EOF
        cboVisit.AddItem rsVisit!VisitName & " (" & rsVisit!VisitCode & ")"
        cboVisit.ItemData(cboVisit.NewIndex) = rsVisit!VisitId
        rsVisit.MoveNext
    Loop
    
    rsVisit.Close
    Set rsVisit = Nothing

Exit Sub
Errhandler:
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
Private Sub LoadSubjectCombo()
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsSubjects As ADODB.Recordset

    On Error GoTo Errhandler
    
    'Clear current contents of cboSite
    cboSubjects.Clear
    
    sSQL = "SELECT PersonId, LocalIdentifier1 " _
        & "FROM TrialSubject " _
        & "WHERE ClinicalTrialID = " & glSelTrialId _
        & " AND TrialSite = '" & gsSelSite & "'" _
        & " AND " & goUser.DataLists.StudiesSitesWhereSQL("TrialSubject.ClinicalTrialID", "TrialSubject.TrialSite") _
        & " ORDER BY LocalIdentifier1, PersonId"
    Set rsSubjects = New ADODB.Recordset
    rsSubjects.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do Until rsSubjects.EOF
        If IsNull(rsSubjects!LocalIdentifier1) Then
            cboSubjects.AddItem "(" & rsSubjects!PersonID & ")"
        Else
            If rsSubjects!LocalIdentifier1 = "" Then
                cboSubjects.AddItem "(" & rsSubjects!PersonID & ")"
            Else
                cboSubjects.AddItem rsSubjects!LocalIdentifier1
            End If
        End If
        cboSubjects.ItemData(cboSubjects.NewIndex) = rsSubjects!PersonID
        rsSubjects.MoveNext
    Loop
    
    rsSubjects.Close
    Set rsSubjects = Nothing

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadSubjectCombo")
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
'It retrieves all (non-visit-date) eForms within the selected Visit.
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsEForm As ADODB.Recordset

    On Error GoTo Errhandler
    
    'cboEForm.Enabled = True
    'Clear current contents of cboEForm
    cboEForm.Clear
    glSelCRFPageId = 0
    glFirstSelCRFPageId = 0
    gsSelCRFPageCode = ""
    
    'retrieve all eForms within the selected Visit (apart from visit date eforms)
    sSQL = "SELECT CRFPage.CRFPageId, CRFPage.CRFPageOrder, CRFPage.CRFPageCode, CRFPage.CRFTitle " _
        & "FROM CRFPage, StudyVisitCRFPage " _
        & "WHERE CRFPage.ClinicalTrialId = StudyVisitCRFPage.ClinicalTrialId " _
        & "AND CRFPage.CRFPageId = StudyVisitCRFPage.CRFPageId " _
        & "AND CRFPage.ClinicalTrialId = " & glSelTrialId & " " _
        & "AND StudyVisitCRFPage.VisitId = " & glSelVisitId & " " _
        & "AND StudyVisitCRFPage.eFormUse = 0 " _
        & "ORDER BY CRFPage.CRFPageOrder"
    
    Set rsEForm = New ADODB.Recordset
    rsEForm.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do Until rsEForm.EOF
        cboEForm.AddItem rsEForm!CRFTitle & " (" & rsEForm!CRFPageCode & ")"
        cboEForm.ItemData(cboEForm.NewIndex) = rsEForm!CRFPageId
        rsEForm.MoveNext
    Loop
    
    rsEForm.Close
    Set rsEForm = Nothing

Exit Sub
Errhandler:
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
Private Sub LoadVisitInstanceCombo()
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsVisitInstances As ADODB.Recordset

    On Error GoTo Errhandler
    
    'cboVisitInstance.Enabled = True
    'Clear current contents of cboVisitInstance
    cboVisitInstance.Clear
    glSelVisitCycleNumber = 0
    
    'retrieve all Visit Instances for the currently selected Visit
    sSQL = "SELECT VisitCycleNumber, VisitDate " _
        & "FROM VisitInstance " _
        & "WHERE ClinicalTrialId = " & glSelTrialId & " " _
        & "AND TrialSite = '" & gsSelSite & "'" _
        & "AND PersonId = " & glSelPersonId & " " _
        & "AND VisitId = " & glSelVisitId & " " _
        & "ORDER BY VisitCycleNumber"
    Set rsVisitInstances = New ADODB.Recordset
    rsVisitInstances.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsVisitInstances.RecordCount = 0 Then
        'If no instances exist offer user instance 1
        cboVisitInstance.AddItem "(1)"
        cboVisitInstance.ItemData(cboVisitInstance.NewIndex) = 1
    Else
        Do Until rsVisitInstances.EOF
            If rsVisitInstances!VisitDate = 0 Then
                cboVisitInstance.AddItem "(" & rsVisitInstances!VisitCycleNumber & ")"
            Else
                cboVisitInstance.AddItem "(" & rsVisitInstances!VisitCycleNumber & ") " & CDate(rsVisitInstances!VisitDate)
            End If
            cboVisitInstance.ItemData(cboVisitInstance.NewIndex) = rsVisitInstances!VisitCycleNumber
            rsVisitInstances.MoveNext
        Loop
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadVisitInstanceCombo")
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
Private Sub LoadEFormInstanceCombo()
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsEFormInstances As ADODB.Recordset

    On Error GoTo Errhandler
    
    cboEFormInstance.Enabled = True
    'Clear current contents of cboEFormInstance
    cboEFormInstance.Clear
    glSelCRFPageCycleNumber = 0
    
    'retrieve all eForm Instances for the currently selected eForm
    sSQL = "SELECT CRFPageCycleNumber, CRFPageDate " _
        & "FROM CRFPageInstance " _
        & "WHERE ClinicalTrialId = " & glSelTrialId & " " _
        & "AND TrialSite = '" & gsSelSite & "'" _
        & "AND PersonId = " & glSelPersonId & " " _
        & "AND VisitId = " & glSelVisitId & " " _
        & "AND VisitCycleNumber = " & glSelVisitCycleNumber & " " _
        & "AND CRFPageId = " & glSelCRFPageId & " " _
        & "ORDER BY CRFPageCycleNumber"
    Set rsEFormInstances = New ADODB.Recordset
    rsEFormInstances.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsEFormInstances.RecordCount = 0 Then
        'If no instances exist offer user instance 1
        cboEFormInstance.AddItem "(1)"
        cboEFormInstance.ItemData(cboEFormInstance.NewIndex) = 1
    Else
        Do Until rsEFormInstances.EOF
            If rsEFormInstances!CRFPageDate = 0 Then
                cboEFormInstance.AddItem "(" & rsEFormInstances!CRFPageCycleNumber & ")"
            Else
                cboEFormInstance.AddItem "(" & rsEFormInstances!CRFPageCycleNumber & ") " & CDate(rsEFormInstances!CRFPageDate)
            End If
            cboEFormInstance.ItemData(cboEFormInstance.NewIndex) = rsEFormInstances!CRFPageCycleNumber
            rsEFormInstances.MoveNext
        Loop
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadEFormInstanceCombo")
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
Private Function LockSubjectGeneration(ByVal lClinicalTrialId As Long, _
                                        ByVal sSite As String, _
                                        ByRef sMessage As String) As String
'--------------------------------------------------------------------
Dim sToken As String

    On Error GoTo ErrLabel

    'Calling LockSubject with NULL_INTEGER as the SubjectId has the effect
    'of locking New Subject Generation for the specified Study and Site
    sToken = LockSubject(goUser.UserName, lClinicalTrialId, sSite, NULL_INTEGER, sMessage)

    LockSubjectGeneration = sToken

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.LockSubjectGeneration"

End Function

'--------------------------------------------------------------------
Private Sub UnLockSubjectGeneration(ByVal lClinicalTrialId As Long, _
                                        ByVal sSite As String, _
                                        ByVal sToken As String)
'--------------------------------------------------------------------

    On Error GoTo ErrLabel

    Call UnlockSubject(lClinicalTrialId, sSite, NULL_INTEGER, sToken)
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.UnlockSubjectGeneration"

End Sub

'---------------------------------------------------------------------
Public Function LockSubject(ByVal sUser As String, _
                            ByVal lStudyId As Long, _
                            ByVal sSite As String, _
                            ByVal lSubjectId As Long, _
                            ByRef sMessage As String) As String
'---------------------------------------------------------------------
'Lock a subject. Based on MACRO_DM's modDataEntry.LockSubject
'Returns a token if lock is successful or empty string if not.
'If the Subject can not be locked sMessage is set to the reason.
' NCJ 1 Jul 04 - Return more meaningful error messages
'---------------------------------------------------------------------
Dim sToken As String
Dim sLockDetails As String

    On Error GoTo ErrLabel
    
    sToken = MACROLOCKBS30.LockSubject(gsADOConnectString, sUser, lStudyId, sSite, lSubjectId)
    Select Case sToken
    Case MACROLOCKBS30.DBLocked.dblStudy
        sLockDetails = MACROLOCKBS30.LockDetailsStudy(gsADOConnectString, lStudyId)
        If sLockDetails = "" Then
            sMessage = "This study is currently being edited by another user."
        Else
            sMessage = "This study is currently being edited by " & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblSubject
        sLockDetails = MACROLOCKBS30.LockDetailsSubject(gsADOConnectString, lStudyId, sSite, lSubjectId)
        If sLockDetails = "" Then
            sMessage = "This subject is currently being edited by another user."
        Else
            sMessage = "This subject is currently being edited by " & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblEFormInstance
        ' An eForm is in use, but we don't know which one, so give a generic message
        sMessage = "This subject is currently being edited by another user."
        sToken = ""
    End Select
    
    LockSubject = sToken
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.LockSubject"

End Function

'---------------------------------------------------------------------
Public Sub UnlockSubject(ByVal lStudyId As Long, _
                        ByVal sSite As String, _
                        ByVal lSubjectId As Long, _
                        ByVal sToken As String)
'---------------------------------------------------------------------
'Unlock a subject. Based on MACRO_DM's modDataEntry.UnlockSubject
'---------------------------------------------------------------------

    On Error GoTo ErrLabel

    If sToken <> "" Then
        'if no gsStudyToken then UnlockSubject is being called without a corresponding LockSubject being called first
        MACROLOCKBS30.UnlockSubject gsADOConnectString, sToken, lStudyId, sSite, lSubjectId
        'always set this to empty string for same reason as above
        sToken = ""
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.UnlockSubject"

End Sub

'---------------------------------------------------------------------
Private Sub ClearSitecbo()
'---------------------------------------------------------------------

    cboSite.Clear
    cboSite.ListIndex = -1
    cboSite.Enabled = False
    gsSelSite = ""

End Sub

'---------------------------------------------------------------------
Private Sub ClearSubjectscbo()
'---------------------------------------------------------------------

    cboSubjects.Clear
    cboSubjects.ListIndex = -1
    cboSubjects.Enabled = False
    cmdGenerateSubject.Enabled = False
    gsSelSubject = ""
    glSelPersonId = 0

End Sub

'---------------------------------------------------------------------
Private Sub ClearVisitcbo()
'---------------------------------------------------------------------

    cboVisit.Clear
    cboVisit.ListIndex = -1
    cboVisit.Enabled = False
    glSelVisitId = 0
    gsSelVisitCode = ""

End Sub

'---------------------------------------------------------------------
Private Sub ClearVisitInstancecbo()
'---------------------------------------------------------------------

    cboVisitInstance.Clear
    cboVisitInstance.ListIndex = -1
    cboVisitInstance.Enabled = False
    glSelVisitCycleNumber = 0

End Sub

'---------------------------------------------------------------------
Private Sub ClearFormcbo()
'---------------------------------------------------------------------

    cboEForm.Clear
    cboEForm.ListIndex = -1
    cboEForm.Enabled = False
    glSelCRFPageId = 0
    glFirstSelCRFPageId = 0
    gsSelCRFPageCode = ""

End Sub

'---------------------------------------------------------------------
Private Sub ClearFormInstancecbo()
'---------------------------------------------------------------------

    cboEFormInstance.Clear
    cboEFormInstance.ListIndex = -1
    cboEFormInstance.Enabled = False
    glSelCRFPageCycleNumber = 0

End Sub

'---------------------------------------------------------------------
Private Sub AssessPass(ByRef bDIRDataExists As Boolean, _
                        ByRef bFirstPassExists As Boolean, _
                        ByRef bSecondPassExists As Boolean)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsData As ADODB.Recordset

    'query for data already existing in DataItemResponse table for selected eForm
    sSQL = "SELECT ResponseValue FROM DataItemResponse " _
        & "WHERE ClinicalTrialId = " & glSelTrialId & " " _
        & "AND TrialSite = '" & gsSelSite & "' " _
        & "AND PersonId = " & glSelPersonId & " " _
        & "AND VisitId = " & glSelVisitId & " " _
        & "AND VisitCycleNumber = " & glSelVisitCycleNumber & " " _
        & "AND CRFPageId = " & glSelCRFPageId & " " _
        & "AND CRFPageCycleNumber = " & glSelCRFPageCycleNumber & " " _
        
    Set rsData = New ADODB.Recordset
    rsData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    bDIRDataExists = False
    
    If rsData.RecordCount > 0 Then
        Do Until rsData.EOF
            If Not IsNull(rsData!ResponseValue) Then
                bDIRDataExists = True
                Exit Sub
            End If
            rsData.MoveNext
        Loop
    End If
        
    'query for First pass data.
    sSQL = "SELECT Response FROM DoubleData " _
        & "WHERE ClinicalTrialId = " & glSelTrialId & " " _
        & "AND TrialSite = '" & gsSelSite & "' " _
        & "AND PersonId = " & glSelPersonId & " " _
        & "AND VisitId = " & glSelVisitId & " " _
        & "AND VisitCycleNumber = " & glSelVisitCycleNumber & " " _
        & "AND CRFPageId = " & glSelCRFPageId & " " _
        & "AND CRFPageCycleNumber = " & glSelCRFPageCycleNumber & " " _
        & "AND Pass = 1"
    Set rsData = New ADODB.Recordset
    rsData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsData.RecordCount > 0 Then
        bFirstPassExists = True
    Else
        bFirstPassExists = False
    End If
    
    'query for Second pass data.
    sSQL = "SELECT Response FROM DoubleData " _
        & "WHERE ClinicalTrialId = " & glSelTrialId & " " _
        & "AND TrialSite = '" & gsSelSite & "' " _
        & "AND PersonId = " & glSelPersonId & " " _
        & "AND VisitId = " & glSelVisitId & " " _
        & "AND VisitCycleNumber = " & glSelVisitCycleNumber & " " _
        & "AND CRFPageId = " & glSelCRFPageId & " " _
        & "AND CRFPageCycleNumber = " & glSelCRFPageCycleNumber & " " _
        & "AND Pass = 2"
    Set rsData = New ADODB.Recordset
    rsData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsData.RecordCount > 0 Then
        bSecondPassExists = True
    Else
        bSecondPassExists = False
    End If

End Sub

'---------------------------------------------------------------------
Private Sub BuildDDeForm()
'---------------------------------------------------------------------
Dim oControl As Control
Dim oLabel As Label
Dim oButton As CommandButton
Dim rseFormElements As ADODB.Recordset
Dim lTopCount As Long
Dim lMaxWidth As Long
Dim oTextBox As TextBox
Dim sQGroupName As String
Dim sDataItemName As String
Dim nDataItemType As Integer
Dim nDataItemLength As Integer
Dim bRQGFlag As Boolean
Dim nRQGRepeatNumber As Integer
Dim lRQGId As Long
Dim sHeader As String
Dim sDDResponse As String
Dim lRQGStart As Long
Dim lRQGHeight As Long
Dim nMaxRepeats As Integer
Dim nLastRepeatNumber As Integer
Dim lNexteFormId As Long
Dim lPreveFormId As Long
Dim nButtonWidth As Integer
Dim lPreviousRepeatTop As Long
Dim lPreviouseFormTop As Long
Dim nFieldOrder As Integer
Dim rsCatCodes As ADODB.Recordset
Dim sCatCodeText As String
Dim sLineCatCodeText As String
Dim sAllCatCodeText As String
Dim nCatCodeLines As Integer

    Call ClearDDeForm
    
    Call SetFontAndColour
    
    bRQGFlag = False
    gnLabelIndex = 1
    gnTextBoxIndex = 1
    gnCommandIndex = 1
    gnCatCodeIndex = 1
    nCatCodeLines = 0
    lTopCount = 100
    picFontSizing.FontSize = gnRegDDFontSize
    'note that Question Names are upto 50 chars long
    gn50CharWidth = picFontSizing.TextWidth(String(50, "X"))
    gnTextBoxHeight = picFontSizing.TextHeight("X")
    nButtonWidth = picFontSizing.TextWidth("  XxxxxxxxXxxxxxxxX  ")
    'If category codes are to be displayed give lMaxWidth an intial value of 1.5 * gn50CharWidth
    If gbRegDisplayCategoryCodes Then
        lMaxWidth = gn50CharWidth * 1.5
    Else
        lMaxWidth = 0
    End If
    
    'Place a header label on the form
    If gnPassNumber = ePassNumber.First Then
        sHeader = " [First Pass]"
    Else
        sHeader = " [Second Pass]"
    End If
    sHeader = sHeader & "    Subject: " & gsSelTrialName & "/" & gsSelSite & "/" & glSelPersonId _
        & "    Visit: " & VisitNameFromId(glSelTrialId, glSelVisitId, glSelVisitCycleNumber) & " (" & gsSelVisitCode & ")[" & glSelVisitCycleNumber & "]" _
        & "    eForm: " & CRFPageNameFromId(glSelTrialId, glSelCRFPageId) & " (" & gsSelCRFPageCode & ")[" & glSelCRFPageCycleNumber & "]    "
    With lblHeader
        .Caption = sHeader
        .Top = lTopCount
        .Left = 100
        .Width = picFontSizing.TextWidth(sHeader & "               ")
        .Height = (gnTextBoxHeight + 100)
        .Visible = True
        'Check lMaxWidth is wide enough for lblHeader
        If lMaxWidth < (.Left + .Width) Then
            lMaxWidth = .Left + .Width
        End If
    End With
    
    'increment lTopCount
    lTopCount = lTopCount + (2 * gnTextBoxHeight)
    
    'loop through the form contents to build the DDeForm
    Set rseFormElements = New ADODB.Recordset
    Set rseFormElements = GetEFormQuestionsAndData()
    Do Until rseFormElements.EOF
        'Skip derived questions
        If Not QuestionIsDerived(glSelTrialId, rseFormElements!DataItemId) Then
            'Check for the need to put "Another row" button on the screen
            'because the previous question was the last question in a RQG
            If (bRQGFlag) And (rseFormElements!OwnerQGroupId = 0) Then
                'If the Max number of repeats for this RQG has been exceeded don't add an "Another row" button
                nMaxRepeats = GetRQGMaxRepeats(glSelTrialId, glSelCRFPageId, lRQGId)
                If nMaxRepeats > nLastRepeatNumber Then
                    Load Me.cmdAnotherRQGRow(gnCommandIndex)
                    Set oButton = Me.cmdAnotherRQGRow(gnCommandIndex)
                    gnCommandIndex = gnCommandIndex + 1
                    With oButton
                        .Top = lTopCount
                        .Left = 600 + gn50CharWidth
                        .Width = nButtonWidth
                        .Height = (gnTextBoxHeight + 100)
                        .Visible = True
                    End With
                    'Store Question Groups Id (QGroupID), Height of RQG and RepeatNumber for next row
                    If nMaxRepeats - nLastRepeatNumber = 1 Then
                        'The following settings would be used to create the last row of
                        'a Repeating Question Group, which would not require "Another Row"
                        'command button, hence the height is (2 * gnTextBoxHeight) shorter
                        oButton.Tag = lRQGId & "|" & nFieldOrder & "|" & lTopCount - lRQGStart - (2 * gnTextBoxHeight) & "|" & nLastRepeatNumber & "|" & lMaxWidth
                    Else
                        oButton.Tag = lRQGId & "|" & nFieldOrder & "|" & lTopCount - lRQGStart & "|" & nLastRepeatNumber & "|" & lMaxWidth
                    End If
                    bRQGFlag = False
                    'increment lTopCount
                    lTopCount = lTopCount + (2 * gnTextBoxHeight)
                End If
            End If
            'Store FieldOrder that will be used in cmdAnotherRQGRow.tags
            nFieldOrder = rseFormElements!FieldOrder
            'Check for a Repeating Question Group
            If rseFormElements!QGroupID > 0 Then
                'Create and position Repeating Question Group Label
                Load Me.lblQuestionLabel(gnLabelIndex)
                Set oLabel = Me.lblQuestionLabel(gnLabelIndex)
                gnLabelIndex = gnLabelIndex + 1
                sQGroupName = " [Question Group] " & QGroupNameFromId(glSelTrialId, rseFormElements!QGroupID)
                With oLabel
                    .Caption = sQGroupName
                    .Top = lTopCount
                    .Left = 100
                    .Width = picFontSizing.TextWidth(sQGroupName & "    ")
                    .Height = (gnTextBoxHeight + 100)
                    .Visible = True
                End With
                'Store the Repeating Question Group Id
                lRQGId = rseFormElements!QGroupID
            End If
            'check for a question
            If rseFormElements!DataItemId > 0 Then
                Call GetDataItemDetails(glSelTrialId, rseFormElements!DataItemId, sDataItemName, nDataItemType, nDataItemLength)
                'Create Question Label
                Load Me.lblQuestionLabel(gnLabelIndex)
                Set oLabel = Me.lblQuestionLabel(gnLabelIndex)
                gnLabelIndex = gnLabelIndex + 1
                'check for a question that is part of a repeating question group
                If rseFormElements!OwnerQGroupId > 0 Then
                    If IsNull(rseFormElements!RepeatNumber) Then
                        sDataItemName = " [" & rseFormElements!FieldOrder & "." & rseFormElements!QGroupFieldOrder & ".1] " & sDataItemName
                        'Store the RepeatNumber
                        nLastRepeatNumber = 1
                    Else
                        sDataItemName = " [" & rseFormElements!FieldOrder & "." & rseFormElements!QGroupFieldOrder & "." & rseFormElements!RepeatNumber & "] " & sDataItemName
                        'Store the RepeatNumber
                        nLastRepeatNumber = rseFormElements!RepeatNumber
                    End If
                    'reset the bRQGFlag flag everytime a new cycle of the RQG is started, so that
                    'lRQGStart is measured at the start of the RQG's last row.
                    If bRQGFlag And nLastRepeatNumber <> nRQGRepeatNumber Then
                        bRQGFlag = False
                    End If
                    'Is this the first question in a Repeating question Group
                    If Not bRQGFlag Then
                        'Set the its a Repeating Question Group Flag
                        bRQGFlag = True
                        'Store repeating question group starting point
                        lRQGStart = lTopCount
                        'Store the RepeatNumber for this cycle of the RQG
                        nRQGRepeatNumber = nLastRepeatNumber
                    End If
                Else
                    sDataItemName = " [" & rseFormElements!FieldOrder & "] " & sDataItemName
                End If
                With oLabel
                    .Left = 100
                    .Caption = sDataItemName
                    .Top = lTopCount
                    .Width = picFontSizing.TextWidth(sDataItemName & "   ")
                    .Height = (gnTextBoxHeight + 100)
                    .Visible = True
                End With
                
                'Create Question TextBox
                Load Me.txtQuestionResponse(gnTextBoxIndex)
                Set oTextBox = Me.txtQuestionResponse(gnTextBoxIndex)
                
                'get previously entered response for displaying in TextBox
                If IsNull(rseFormElements!Response) Then
                    sDDResponse = ""
                Else
                    sDDResponse = rseFormElements!Response
                End If
                With oTextBox
                    .Text = sDDResponse
                    .Top = lTopCount
                    .Left = 600 + gn50CharWidth
                    .Width = picFontSizing.TextWidth(String(nDataItemLength + 3, "_"))
                    .Height = (gnTextBoxHeight + 100)
                    If lMaxWidth < (.Left + .Width) Then
                        lMaxWidth = .Left + .Width
                    End If
                    'store DataItemId|FieldOrder|RepeatNumber|QGroupFieldOrder|DataItemType in tag
                    If IsNull(rseFormElements!RepeatNumber) Then
                        .Tag = rseFormElements!DataItemId & "|" & rseFormElements!FieldOrder & "|1|" & rseFormElements!QGroupFieldOrder & "|" & nDataItemType
                    Else
                        .Tag = rseFormElements!DataItemId & "|" & rseFormElements!FieldOrder & "|" & rseFormElements!RepeatNumber & "|" & rseFormElements!QGroupFieldOrder & "|" & nDataItemType
                    End If
                    'Only enable the TextBox if the response still has status entered
                    If Not IsNull(rseFormElements!Status) Then
                        If rseFormElements!Status = eDoubleDataStatus.Entered Then
                            .Enabled = True
                        End If
                    Else
                        .Enabled = True
                    End If
                    'disable textbox if its a multimedia question
                    If nDataItemType = eDataType.Multimedia Then
                        .Enabled = False
                    End If
                    .Visible = True
                End With
                'If this eForm is being opened in Double Data Entry for the first time then
                'perform an initial save of a blank response to the DoubleData table
                If IsNull(rseFormElements!RepeatNumber) Then
                    Call SaveUpdateResponse(gnTextBoxIndex)
                End If
                
                'Display category codes if required
                If gbRegDisplayCategoryCodes Then
                    'If the question is a category code question
                    If nDataItemType = eDataType.Category Then
                        Load Me.lblCatCodes(gnCatCodeIndex)
                        Set oLabel = Me.lblCatCodes(gnCatCodeIndex)
                        gnCatCodeIndex = gnCatCodeIndex + 1
                        'get a recordset of this questions category codes
                        Set rsCatCodes = New ADODB.Recordset
                        Set rsCatCodes = GetCatCodes(rseFormElements!DataItemId)
                        sAllCatCodeText = ""
                        sLineCatCodeText = ""
                        nCatCodeLines = 1
                        Do Until rsCatCodes.EOF
                            sCatCodeText = "[" & rsCatCodes!ValueCode & " - " & rsCatCodes!ItemValue & "] "
                            If lMaxWidth < (600 + gn50CharWidth + picFontSizing.TextWidth(sCatCodeText)) Then
                                lMaxWidth = 600 + gn50CharWidth + picFontSizing.TextWidth(sCatCodeText)
                            End If
                            'can this category code fit onto current line
                            If (picFontSizing.TextWidth(sLineCatCodeText) + picFontSizing.TextWidth(sCatCodeText) + 600 + gn50CharWidth) > lMaxWidth Then
                                'add linefeed then caption and increment line counter
                                sAllCatCodeText = sAllCatCodeText & vbNewLine & sCatCodeText
                                nCatCodeLines = nCatCodeLines + 1
                                sLineCatCodeText = sCatCodeText
                            Else
                                'just add caption
                                sAllCatCodeText = sAllCatCodeText & sCatCodeText
                                sLineCatCodeText = sLineCatCodeText & sCatCodeText
                            End If
                            rsCatCodes.MoveNext
                        Loop
                        rsCatCodes.Close
                        Set rsCatCodes = Nothing
                        With oLabel
                            .Caption = sAllCatCodeText
                            .Top = lTopCount + (1.5 * gnTextBoxHeight)
                            .Left = 600 + gn50CharWidth
                            .Width = lMaxWidth - (600 + gn50CharWidth)
                            .Height = nCatCodeLines * gnTextBoxHeight
                            .Visible = True
                        End With
                    End If
                End If
                
                'increment gnTextBoxIndex
                gnTextBoxIndex = gnTextBoxIndex + 1
            End If
            
            'Increment lTopCount
            lTopCount = lTopCount + (2 * gnTextBoxHeight) + (nCatCodeLines * gnTextBoxHeight)
            nCatCodeLines = 0
        End If
        rseFormElements.MoveNext
    Loop
    
    rseFormElements.Close
    Set rseFormElements = Nothing
    
    'Check for the need to put "Another row" button on the screen
    If bRQGFlag Then
        'If the Max number of repeats for this RQG has been exceeded don't add an "Another row" button
        nMaxRepeats = GetRQGMaxRepeats(glSelTrialId, glSelCRFPageId, lRQGId)
        If nMaxRepeats > nLastRepeatNumber Then
            Load Me.cmdAnotherRQGRow(gnCommandIndex)
            Set oButton = Me.cmdAnotherRQGRow(gnCommandIndex)
            gnCommandIndex = gnCommandIndex + 1
            With oButton
                .Top = lTopCount
                .Left = 600 + gn50CharWidth
                .Width = nButtonWidth
                .Height = (gnTextBoxHeight + 100)
                .Visible = True
            End With
            'Store Question Groups Id (QGroupID), Height of RQG and RepeatNumber for next row
            If nMaxRepeats - nLastRepeatNumber = 1 Then
                'The following settings would be used to create the last row of
                'a Repeating Question Group, which would not require "Another Row"
                'command button, hence the height is (2 * gnTextBoxHeight) shorter
                oButton.Tag = lRQGId & "|" & nFieldOrder & "|" & lTopCount - lRQGStart - (2 * gnTextBoxHeight) & "|" & nLastRepeatNumber & "|" & lMaxWidth
            Else
                oButton.Tag = lRQGId & "|" & nFieldOrder & "|" & lTopCount - lRQGStart & "|" & nLastRepeatNumber & "|" & lMaxWidth
            End If
            'increment lTopCount
            lTopCount = lTopCount + gnTextBoxHeight + 100
        End If
    End If
    
    'Place the required command buttons at the bottom of the DDeForm
    'Note that the buttons are loaded in the order of their required tab order:-
    '   Another Row     1
    '   Repeat eForm    2
    '   Next eForm      3
    '   End Session     4
    '   Previous Repeat 5
    '   Previous eForm  6
    If EFormIsRepeating Then
        'position and display the Repeat eForm command button
        Load Me.cmdRepeatEForm(1)
        Set oButton = Me.cmdRepeatEForm(1)
        With oButton
            .Top = lTopCount
            .Left = 600 + gn50CharWidth
            .Width = nButtonWidth
            .Height = (gnTextBoxHeight + 100)
            .Visible = True
        End With
        'store lPreviousRepeatTop
        lPreviousRepeatTop = lTopCount
        lTopCount = lTopCount + gnTextBoxHeight + 100
    End If
    
    'If the currently created eForm is the VisitDateEForm
    If glSelCRFPageId = mlVisitDateEformId Then
        mbVisitDateEform = True
        glCRFPageIdAfterVisitDate = glFirstSelCRFPageId
    End If
    
    'Add the Next eForm command button
    Load Me.cmdNextEForm(1)
    Set oButton = Me.cmdNextEForm(1)
    With oButton
        .Top = lTopCount
        .Left = 600 + gn50CharWidth
        .Width = nButtonWidth
        .Height = (gnTextBoxHeight + 100)
        .Visible = True
        If glCRFPageIdAfterVisitDate = 0 Then
            If NexteFormExists(glSelTrialId, glSelVisitId, glSelCRFPageId, lNexteFormId) Then
                .Tag = lNexteFormId
                .Enabled = True
            Else
                .Tag = ""
                .Enabled = False
            End If
        Else
            .Tag = glCRFPageIdAfterVisitDate
            .Enabled = True
            'set glCRFPageIdAfterVisitDate back to 0 (its done its job of making
            'sure that the CRFPageId of the form that was to be built before
            'the Visit-date-eForm is placed in the tag of the cmdNextEForm button)
            glCRFPageIdAfterVisitDate = 0
        End If
    End With
    'store lPreviouseFormTop
    lPreviouseFormTop = lTopCount
    lTopCount = lTopCount + gnTextBoxHeight + 100
    
    'Add the End session command button
    Load Me.cmdEndSession(1)
    Set oButton = Me.cmdEndSession(1)
    With oButton
        .Top = lTopCount
        .Left = 600 + gn50CharWidth
        .Width = nButtonWidth
        .Height = (gnTextBoxHeight + 100)
        .Visible = True
        'Check lMaxWidth is wide enough for command buttons
        If lMaxWidth < (.Left + .Width) Then
            lMaxWidth = .Left + .Width
        End If
    End With
    lTopCount = lTopCount + gnTextBoxHeight + 100
    
    If EFormIsRepeating Then
        'position and display the Previous Repeat command button
        Load Me.cmdPreviousRepeat(1)
        Set oButton = Me.cmdPreviousRepeat(1)
        With oButton
            .Top = lPreviousRepeatTop
            .Left = 600 + gn50CharWidth - nButtonWidth
            .Width = nButtonWidth
            .Height = (gnTextBoxHeight + 100)
            .Visible = True
            If glSelCRFPageCycleNumber > 1 Then
                .Enabled = True
            Else
                .Enabled = False
            End If
        End With
    End If
        
    'Add the Previous eForm command button
    Load Me.cmdPreviousEForm(1)
    Set oButton = Me.cmdPreviousEForm(1)
    With oButton
        .Top = lPreviouseFormTop
        .Left = 600 + gn50CharWidth - nButtonWidth
        .Width = nButtonWidth
        .Height = (gnTextBoxHeight + 100)
        .Visible = True
        If PreveFormExists(glSelTrialId, glSelVisitId, glSelCRFPageId, lPreveFormId) Then
            .Tag = lPreveFormId
            .Enabled = True
        ElseIf Not mbVisitDateEform And mbVisitDateEformExists Then
            .Tag = mlVisitDateEformId
            .Enabled = True
        Else
            .Tag = ""
            .Enabled = False
        End If
    End With
    
    'Set size of picDDEForm
    picDDEForm.Height = lTopCount + 500
    picDDEForm.Width = lMaxWidth + 500
    'Position top left corner of
    picDDEForm.Top = 0
    picDDEForm.Left = 0
    
    'call DDeFormChecks which checks for increasing width and/or height
    Call DDeFormChecks
    
    'call DDeFormScrollbars which checks the need for scrollbars
    Call DDeFormScrollbars
    
    picDDEForm.Visible = True
    
    'set the eForms being displayed flag to True
    mbEFormsBeingDisplayed = True
    
    If txtQuestionResponse.Count > 1 Then
        'set the focus to the first question of the newly added row
        If txtQuestionResponse(1).Enabled = True Then
            txtQuestionResponse(1).SetFocus
        End If
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVChangeBlues_Click()
'---------------------------------------------------------------------

    mnuVChangeGreys.Checked = False
    mnuVChangeGreens.Checked = False
    mnuVChangeBlues.Checked = True
    mnuVChangePurples.Checked = False
    mnuVChangeReds.Checked = False
    
    gsRegDDColourScheme = "Blues"
    glRegDDLightColour = RGB(238, 244, 251)
    glRegDDMediumColour = RGB(213, 228, 244)
    glRegDDDarKColour = RGB(188, 213, 237)
    
    Call SetMACROSetting("DDColourScheme", gsRegDDColourScheme)
    Call SetMACROSetting("DDLightColour", glRegDDLightColour)
    Call SetMACROSetting("DDMediumColour", glRegDDMediumColour)
    Call SetMACROSetting("DDDarkColour", glRegDDDarKColour)

    If mbEFormsBeingDisplayed Then
        If mbVisitDateEform Then
            glCRFPageIdAfterVisitDate = cmdNextEForm(1).Tag
        End If
        Call BuildDDeForm
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVChangeGreens_Click()
'---------------------------------------------------------------------

    mnuVChangeGreys.Checked = False
    mnuVChangeGreens.Checked = True
    mnuVChangeBlues.Checked = False
    mnuVChangePurples.Checked = False
    mnuVChangeReds.Checked = False
    
    gsRegDDColourScheme = "Greens"
    glRegDDLightColour = RGB(238, 251, 244)
    glRegDDMediumColour = RGB(213, 244, 228)
    glRegDDDarKColour = RGB(188, 237, 213)
    
    Call SetMACROSetting("DDColourScheme", gsRegDDColourScheme)
    Call SetMACROSetting("DDLightColour", glRegDDLightColour)
    Call SetMACROSetting("DDMediumColour", glRegDDMediumColour)
    Call SetMACROSetting("DDDarkColour", glRegDDDarKColour)

    If mbEFormsBeingDisplayed Then
        If mbVisitDateEform Then
            glCRFPageIdAfterVisitDate = cmdNextEForm(1).Tag
        End If
        Call BuildDDeForm
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVChangePurples_Click()
'---------------------------------------------------------------------

    mnuVChangeGreys.Checked = False
    mnuVChangeGreens.Checked = False
    mnuVChangeBlues.Checked = False
    mnuVChangePurples.Checked = True
    mnuVChangeReds.Checked = False
    
    gsRegDDColourScheme = "Purples"
    glRegDDLightColour = RGB(244, 238, 251)
    glRegDDMediumColour = RGB(228, 213, 244)
    glRegDDDarKColour = RGB(213, 188, 237)
    
    Call SetMACROSetting("DDColourScheme", gsRegDDColourScheme)
    Call SetMACROSetting("DDLightColour", glRegDDLightColour)
    Call SetMACROSetting("DDMediumColour", glRegDDMediumColour)
    Call SetMACROSetting("DDDarkColour", glRegDDDarKColour)

    If mbEFormsBeingDisplayed Then
        If mbVisitDateEform Then
            glCRFPageIdAfterVisitDate = cmdNextEForm(1).Tag
        End If
        Call BuildDDeForm
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVChangeReds_Click()
'---------------------------------------------------------------------

    mnuVChangeGreys.Checked = False
    mnuVChangeGreens.Checked = False
    mnuVChangeBlues.Checked = False
    mnuVChangePurples.Checked = False
    mnuVChangeReds.Checked = True
    
    gsRegDDColourScheme = "Reds"
    glRegDDLightColour = RGB(251, 238, 241)
    glRegDDMediumColour = RGB(244, 213, 221)
    glRegDDDarKColour = RGB(237, 188, 201)
    
    Call SetMACROSetting("DDColourScheme", gsRegDDColourScheme)
    Call SetMACROSetting("DDLightColour", glRegDDLightColour)
    Call SetMACROSetting("DDMediumColour", glRegDDMediumColour)
    Call SetMACROSetting("DDDarkColour", glRegDDDarKColour)

    If mbEFormsBeingDisplayed Then
        If mbVisitDateEform Then
            glCRFPageIdAfterVisitDate = cmdNextEForm(1).Tag
        End If
        Call BuildDDeForm
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVChangeGreys_Click()
'---------------------------------------------------------------------

    mnuVChangeGreys.Checked = True
    mnuVChangeGreens.Checked = False
    mnuVChangeBlues.Checked = False
    mnuVChangePurples.Checked = False
    mnuVChangeReds.Checked = False
    
    gsRegDDColourScheme = "Greys"
    glRegDDLightColour = RGB(244, 244, 244)
    glRegDDMediumColour = RGB(228, 228, 228)
    glRegDDDarKColour = RGB(212, 212, 212)
    
    Call SetMACROSetting("DDColourScheme", gsRegDDColourScheme)
    Call SetMACROSetting("DDLightColour", glRegDDLightColour)
    Call SetMACROSetting("DDMediumColour", glRegDDMediumColour)
    Call SetMACROSetting("DDDarkColour", glRegDDDarKColour)

    If mbEFormsBeingDisplayed Then
        If mbVisitDateEform Then
            glCRFPageIdAfterVisitDate = cmdNextEForm(1).Tag
        End If
        Call BuildDDeForm
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVDisplayCategoryCodes_Click()
'---------------------------------------------------------------------
    
    If mnuVDisplayCategoryCodes.Checked Then
        mnuVDisplayCategoryCodes.Checked = False
        gbRegDisplayCategoryCodes = False
    Else
        mnuVDisplayCategoryCodes.Checked = True
        gbRegDisplayCategoryCodes = True
    End If
    Call SetMACROSetting("DDDispayCategoryCodes", gbRegDisplayCategoryCodes)
    
    If mbEFormsBeingDisplayed Then
        If mbVisitDateEform Then
            glCRFPageIdAfterVisitDate = cmdNextEForm(1).Tag
        End If
        Call BuildDDeForm
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVFont10_Click()
'---------------------------------------------------------------------

    mnuVFont8.Checked = False
    mnuVFont9.Checked = False
    mnuVFont12.Checked = False
    mnuVFont14.Checked = False
    mnuVFont16.Checked = False
    mnuVFont18.Checked = False
    mnuVFont10.Checked = True
    gnRegDDFontSize = 10
    Call SetMACROSetting("DDFontSize", 10)
    
    If mbEFormsBeingDisplayed Then
        If mbVisitDateEform Then
            glCRFPageIdAfterVisitDate = cmdNextEForm(1).Tag
        End If
        Call BuildDDeForm
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVFont12_Click()
'---------------------------------------------------------------------

    mnuVFont8.Checked = False
    mnuVFont9.Checked = False
    mnuVFont10.Checked = False
    mnuVFont14.Checked = False
    mnuVFont16.Checked = False
    mnuVFont18.Checked = False
    mnuVFont12.Checked = True
    gnRegDDFontSize = 12
    Call SetMACROSetting("DDFontSize", 12)
    
    If mbEFormsBeingDisplayed Then
        If mbVisitDateEform Then
            glCRFPageIdAfterVisitDate = cmdNextEForm(1).Tag
        End If
        Call BuildDDeForm
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVFont14_Click()
'---------------------------------------------------------------------

    mnuVFont8.Checked = False
    mnuVFont9.Checked = False
    mnuVFont10.Checked = False
    mnuVFont12.Checked = False
    mnuVFont16.Checked = False
    mnuVFont18.Checked = False
    mnuVFont14.Checked = True
    gnRegDDFontSize = 14
    Call SetMACROSetting("DDFontSize", 14)
    
    If mbEFormsBeingDisplayed Then
        If mbVisitDateEform Then
            glCRFPageIdAfterVisitDate = cmdNextEForm(1).Tag
        End If
        Call BuildDDeForm
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVFont16_Click()
'---------------------------------------------------------------------

    mnuVFont8.Checked = False
    mnuVFont9.Checked = False
    mnuVFont10.Checked = False
    mnuVFont12.Checked = False
    mnuVFont14.Checked = False
    mnuVFont18.Checked = False
    mnuVFont16.Checked = True
    gnRegDDFontSize = 16
    Call SetMACROSetting("DDFontSize", 16)
    
    If mbEFormsBeingDisplayed Then
        If mbVisitDateEform Then
            glCRFPageIdAfterVisitDate = cmdNextEForm(1).Tag
        End If
        Call BuildDDeForm
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVFont18_Click()
'---------------------------------------------------------------------

    mnuVFont8.Checked = False
    mnuVFont9.Checked = False
    mnuVFont10.Checked = False
    mnuVFont12.Checked = False
    mnuVFont14.Checked = False
    mnuVFont16.Checked = False
    mnuVFont18.Checked = True
    gnRegDDFontSize = 18
    Call SetMACROSetting("DDFontSize", 18)
    
    If mbEFormsBeingDisplayed Then
        If mbVisitDateEform Then
            glCRFPageIdAfterVisitDate = cmdNextEForm(1).Tag
        End If
        Call BuildDDeForm
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVFont8_Click()
'---------------------------------------------------------------------

    mnuVFont9.Checked = False
    mnuVFont10.Checked = False
    mnuVFont12.Checked = False
    mnuVFont14.Checked = False
    mnuVFont16.Checked = False
    mnuVFont18.Checked = False
    mnuVFont8.Checked = True
    gnRegDDFontSize = 8
    Call SetMACROSetting("DDFontSize", 8)
    
    If mbEFormsBeingDisplayed Then
        If mbVisitDateEform Then
            glCRFPageIdAfterVisitDate = cmdNextEForm(1).Tag
        End If
        Call BuildDDeForm
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVFont9_Click()
'---------------------------------------------------------------------

    mnuVFont8.Checked = False
    mnuVFont10.Checked = False
    mnuVFont12.Checked = False
    mnuVFont14.Checked = False
    mnuVFont16.Checked = False
    mnuVFont18.Checked = False
    mnuVFont9.Checked = True
    gnRegDDFontSize = 9
    Call SetMACROSetting("DDFontSize", 9)
    
    If mbEFormsBeingDisplayed Then
        If mbVisitDateEform Then
            glCRFPageIdAfterVisitDate = cmdNextEForm(1).Tag
        End If
        Call BuildDDeForm
    End If

End Sub

'---------------------------------------------------------------------
Private Function QGroupNameFromId(ByVal lClinicalTrialId As Long, _
                            ByVal lQGroupId As Long) As String
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sTemp As String

    On Error GoTo Errhandler
    
    sTemp = ""
    sSQL = "SELECT QGroupName FROM QGroup " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND QGroupId = " & lQGroupId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        sTemp = ""
    Else
        sTemp = RemoveNull(rsTemp!QGroupName)
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

    QGroupNameFromId = sTemp
    
Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "QGroupNameFromId", "basTrialData")
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
Private Sub tmrRTE365_Timer()
'---------------------------------------------------------------------

    tmrRTE365.Enabled = False
    Call ClearDDeForm

End Sub

'---------------------------------------------------------------------
Private Sub txtQuestionResponse_GotFocus(Index As Integer)
'---------------------------------------------------------------------

    Call CheckVerticalScroll(txtQuestionResponse(Index).Top, txtQuestionResponse(Index).Height)
    Call CheckHorizontalScroll(txtQuestionResponse(Index).Left, txtQuestionResponse(Index).Width)

End Sub

'---------------------------------------------------------------------
Private Sub txtQuestionResponse_KeyPress(Index As Integer, KeyAscii As Integer)
'---------------------------------------------------------------------
    
    'Read Return as an entry terminator and turn it into a tab
    If (KeyAscii = vbKeyReturn) Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

'---------------------------------------------------------------------
Private Sub txtQuestionResponse_LostFocus(Index As Integer)
'---------------------------------------------------------------------

    'validate and save the currently entered/edited response
    Call SaveUpdateResponse(Index)

End Sub

'---------------------------------------------------------------------
Private Sub VScroll_Change()
'---------------------------------------------------------------------

    picDDEForm.Top = CSng(-VScroll.Value) * 10

End Sub

'---------------------------------------------------------------------
Private Sub VScroll_Scroll()
'---------------------------------------------------------------------

    picDDEForm.Top = CSng(-VScroll.Value) * 10
    
End Sub

'---------------------------------------------------------------------
Private Sub SaveUpdateResponse(ByVal nIndex As Integer)
'---------------------------------------------------------------------
Dim sResponse As String
Dim asTag() As String
Dim sSQL As String
Dim lDataItemId As Long
Dim nFieldOrder As Integer
Dim nRepeatNumber As Integer
Dim nQGroupFieldOrder As Integer
Dim nDataItemType As Integer
Dim rsDoubleData As ADODB.Recordset

    sResponse = txtQuestionResponse(nIndex).Text
    asTag = Split(txtQuestionResponse(nIndex).Tag, "|")
    lDataItemId = CLng(asTag(0))
    nFieldOrder = CLng(asTag(1))
    nRepeatNumber = CInt(asTag(2))
    nQGroupFieldOrder = CLng(asTag(3))
    nDataItemType = CInt(asTag(4))
            
    If txtQuestionResponse(nIndex).Text <> "" Then
        'Check response for invalid characters and length
        If (InStr(sResponse, "`") > 0) Or (InStr(sResponse, """") > 0) Or (InStr(sResponse, "|") > 0) Or (InStr(sResponse, "~") > 0) Then
            Call DialogError("Question response may not contain double or backwards quotes or the | or ~ characters.", "Invalid Response")
            txtQuestionResponse(nIndex).SetFocus
            Exit Sub
        End If
        If Len(sResponse) > 255 Then
            Call DialogError("Question response may not be longer than 255 characters.", "Invalid Response")
            txtQuestionResponse(nIndex).SetFocus
            Exit Sub
        End If
        'Validate the response based on DataItemType
        Select Case nDataItemType
        Case DataType.Text, DataType.Category, DataType.Thesaurus
            'no additional validation required
        Case DataType.IntegerData
            'check for numbers only
            If Not gblnValidString(sResponse, valNumeric) Then
                Call DialogError("This Integer question can only contain numeric characters.", "Invalid Response")
                txtQuestionResponse(nIndex).SetFocus
                Exit Sub
            End If
        Case DataType.Real, DataType.LabTest
            'check for numbers and decimal point
            If Not gblnValidString(sResponse, valNumeric + valDecimalPoint) Then
                Call DialogError("This Real/Lab question can only contain numeric characters and decimal points", "Invalid Response")
                txtQuestionResponse(nIndex).SetFocus
                Exit Sub
            End If
        Case DataType.Date
            'check for numbers and date separators /.:- and space
            If Not gblnValidString(sResponse, valNumeric + valDateSeperators) Then
                Call DialogError("This Date/Time question can only contain numeric characters and date separators .:-/", "Invalid Response")
                txtQuestionResponse(nIndex).SetFocus
                Exit Sub
            End If
        End Select
    End If
    
    'Check for the response already existing or being saved for the first time
    sSQL = "SELECT * FROM DoubleData " _
        & " WHERE ClinicalTrialId = " & glSelTrialId _
        & " AND TrialSite = '" & gsSelSite & "'" _
        & " AND PersonId = " & glSelPersonId _
        & " AND VisitId = " & glSelVisitId _
        & " AND VisitCycleNumber = " & glSelVisitCycleNumber _
        & " AND CRFPageID = " & glSelCRFPageId _
        & " AND CRFPageCycleNumber = " & glSelCRFPageCycleNumber _
        & " AND DataItemId = " & lDataItemId _
        & " AND RepeatNumber = " & nRepeatNumber _
        & " AND Pass = " & gnPassNumber
    Set rsDoubleData = New ADODB.Recordset
    rsDoubleData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    If rsDoubleData.RecordCount = 0 Then
        rsDoubleData.AddNew
        rsDoubleData.Fields(0) = glSelTrialId
        rsDoubleData.Fields(1) = gsSelSite
        rsDoubleData.Fields(2) = glSelPersonId
        rsDoubleData.Fields(3) = glSelVisitId
        rsDoubleData.Fields(4) = glSelVisitCycleNumber
        rsDoubleData.Fields(5) = glSelCRFPageId
        rsDoubleData.Fields(6) = glSelCRFPageCycleNumber
        rsDoubleData.Fields(7) = lDataItemId
        rsDoubleData.Fields(8) = nFieldOrder
        rsDoubleData.Fields(9) = nRepeatNumber
        rsDoubleData.Fields(10) = nQGroupFieldOrder
        rsDoubleData.Fields(11) = gnPassNumber
        rsDoubleData.Fields(12) = sResponse
        rsDoubleData.Fields(13) = goUser.UserName
        rsDoubleData.Fields(14) = SQLStandardNow
        rsDoubleData.Fields(15) = eDoubleDataStatus.Entered
        rsDoubleData.Update
    ElseIf (rsDoubleData.Fields(12) <> sResponse) Or (IsNull(rsDoubleData.Fields(12)) And (sResponse <> "")) Then
        rsDoubleData.Fields(12) = sResponse
        rsDoubleData.Fields(13) = goUser.UserName
        rsDoubleData.Fields(14) = SQLStandardNow
        rsDoubleData.Fields(15) = eDoubleDataStatus.Entered
        rsDoubleData.Update
    End If
   
    rsDoubleData.Close
    Set rsDoubleData = Nothing
    
End Sub

'---------------------------------------------------------------------
Private Function GetCurrentDDResponse(ByVal nPass As Integer, _
                                    ByVal lDataItemId As Long, _
                                    ByVal nRepeatNumber As Integer) As String
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsDoubleData As ADODB.Recordset
Dim sDDResponse As String

    sSQL = "SELECT Response FROM DoubleData " _
        & " WHERE ClinicalTrialId = " & glSelTrialId _
        & " AND TrialSite = '" & gsSelSite & "'" _
        & " AND PersonId = " & glSelPersonId _
        & " AND VisitId = " & glSelVisitId _
        & " AND VisitCycleNumber = " & glSelVisitCycleNumber _
        & " AND CRFPageId = " & glSelCRFPageId _
        & " AND CRFPageCycleNumber = " & glSelCRFPageCycleNumber _
        & " AND DataItemId = " & lDataItemId _
        & " AND RepeatNumber = " & nRepeatNumber _
        & " AND Pass = " & nPass
    Set rsDoubleData = New ADODB.Recordset
    rsDoubleData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    If rsDoubleData.RecordCount = 0 Then
        sDDResponse = ""
    Else
        sDDResponse = rsDoubleData.Fields(0)
    End If
    
    rsDoubleData.Close
    Set rsDoubleData = Nothing
    
    GetCurrentDDResponse = sDDResponse

End Function

'---------------------------------------------------------------------
Private Function GetRQGMaxRepeats(ByVal lClinicalTrialId As Long, _
                                ByVal lCRFPageId As Long, _
                                ByVal lQGroupId As Long) As Integer
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sMaxRepeats As String

    sSQL = "SELECT MaxRepeats FROM eFormQGroup " _
        & "WHERE ClinicalTrialId = " & lClinicalTrialId & " " _
        & "AND CRFPageId = " & lCRFPageId & " " _
        & "AND QGroupId = " & lQGroupId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 1 Then
        sMaxRepeats = rsTemp!MaxRepeats
    Else
        sMaxRepeats = 0
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

    GetRQGMaxRepeats = CInt(sMaxRepeats)

End Function

'---------------------------------------------------------------------
Private Function GetEFormQuestionsAndData() As ADODB.Recordset
'---------------------------------------------------------------------
'This function returns the definition of a single eForm (glSelCRFPageId) in the form of
'DataItemIds,QGroupIds & OwnerQGroupIds sorted on FieldOrder,RepeatNumber & QGroupFieldOrder.
'
'The eForm definition is bolted to the DoubleData table using a Left Join that will place
'Nulls in the returned DoubleData.RepeatNumber and DoubleData.First/SecondPassResponse when
'nothing exists in the DoubleData table for the specified PersonId. With or without data
'in the DoubleData table this function will always return an eForm definition.
'Note that hidden questions are filtered out.
'---------------------------------------------------------------------
Dim sSQL As String

    If goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        sSQL = "SELECT CRFElement.FieldOrder, DoubleData.RepeatNumber, CRFElement.QGroupFieldOrder, " _
            & "CRFElement.DataItemId, CRFElement.OwnerQGroupId, CRFElement.QGroupId, DoubleData.Response, " _
            & "DoubleData.Status " _
            & "FROM CRFElement, DoubleData " _
            & "WHERE DoubleData.ClinicalTrialId (+) = CRFElement.ClinicalTrialId " _
            & "AND DoubleData.CRFPageId (+) = CRFElement.CRFPageId " _
            & "AND DoubleData.DataItemId (+) =  CRFElement.DataItemId " _
            & "AND DoubleData.TrialSite (+) = '" & gsSelSite & "' " _
            & "AND DoubleData.PersonId (+) = " & glSelPersonId & " " _
            & "AND DoubleData.VisitId (+) = " & glSelVisitId & " " _
            & "AND DoubleData.VisitCycleNumber (+) = " & glSelVisitCycleNumber & " " _
            & "AND DoubleData.CRFPageCycleNumber (+) = " & glSelCRFPageCycleNumber & " " _
            & "AND DoubleData.Pass (+) = " & gnPassNumber & " " _
            & "AND CRFElement.ClinicalTrialId = " & glSelTrialId & " " _
            & "AND CRFElement.CRFPageId = " & glSelCRFPageId & " " _
            & "AND CRFElement.ControlType < " & gnVISUAL_ELEMENT & " " _
            & "AND CRFElement.Hidden = 0 " _
            & "ORDER BY CRFElement.FieldOrder, NVL(DoubleData.RepeatNumber,0), CRFElement.QGroupFieldOrder"
            'The above ORDER BY line contains a NVL funtion call around RepeatNumber, this has the effect of
            'replacing a null RepeatNumber with a zero. This has been done because Oracle places nulls below
            'values when sorting, whereas SQL Server places nulls above values. Having the records with null
            'Repeatnumbers above those with nonnull RepeatNumbers is the desired outcome for this query
    Else
        sSQL = "SELECT CRFElement.FieldOrder, DoubleData.RepeatNumber, CRFElement.QGroupFieldOrder, " _
            & "CRFElement.DataItemId, CRFElement.OwnerQGroupId, CRFElement.QGroupId, DoubleData.Response, " _
            & "DoubleData.Status " _
            & "FROM CRFElement LEFT JOIN DoubleData " _
            & "ON CRFElement.ClinicalTrialId = DoubleData.ClinicalTrialId " _
            & "AND CRFElement.CRFPageId = DoubleData.CRFPageId " _
            & "AND CRFElement.DataItemId = DoubleData.DataItemId " _
            & "AND DoubleData.TrialSite = '" & gsSelSite & "' " _
            & "AND DoubleData.PersonId = " & glSelPersonId & " " _
            & "AND DoubleData.VisitId = " & glSelVisitId & " " _
            & "AND DoubleData.VisitCycleNumber = " & glSelVisitCycleNumber & " " _
            & "AND DoubleData.CRFPageCycleNumber = " & glSelCRFPageCycleNumber & " " _
            & "AND DoubleData.Pass = " & gnPassNumber & " " _
            & "WHERE CRFElement.ClinicalTrialId = " & glSelTrialId & " " _
            & "AND CRFElement.CRFPageId = " & glSelCRFPageId & " " _
            & "AND CRFElement.ControlType < " & gnVISUAL_ELEMENT & " " _
            & "AND CRFElement.Hidden = 0 " _
            & "ORDER BY CRFElement.FieldOrder, DoubleData.RepeatNumber, CRFElement.QGroupFieldOrder"
    End If
        
    Set GetEFormQuestionsAndData = New ADODB.Recordset
    GetEFormQuestionsAndData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

End Function

'---------------------------------------------------------------------
Private Function EFormIsRepeating()
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bItsRepeating As Boolean

    sSQL = "SELECT Repeating FROM StudyVisitCRFPage " _
        & "WHERE ClinicalTrialId = " & glSelTrialId & " " _
        & "AND VisitId = " & glSelVisitId & " " _
        & "AND CRFPageId = " & glSelCRFPageId
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    bItsRepeating = False
    If rsTemp.RecordCount = 1 Then
        If rsTemp!Repeating = 1 Then
            bItsRepeating = True
        End If
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

    EFormIsRepeating = bItsRepeating
    
End Function

'--------------------------------------------------------------------
Private Sub ClearDDeForm()
'--------------------------------------------------------------------
Dim oControl As Control

    picDDEForm.Visible = False
    VScroll.Visible = False
    HScroll.Visible = False

    'clear previous DDeForm content
    For Each oControl In Me.Controls
        If oControl.Name = "lblQuestionLabel" Or oControl.Name = "txtQuestionResponse" Or oControl.Name = "cmdAnotherRQGRow" _
        Or oControl.Name = "cmdRepeatEForm" Or oControl.Name = "cmdNextEForm" Or oControl.Name = "cmdEndSession" _
        Or oControl.Name = "cmdPreviousRepeat" Or oControl.Name = "cmdPreviousEForm" Or oControl.Name = "lblCatCodes" Then
            If oControl.Index > 0 Then
                Unload oControl
                DoEvents
            End If
        End If
    Next

End Sub

'--------------------------------------------------------------------
Private Sub ClearFirstSecondButtons()
'--------------------------------------------------------------------

    cmdFirstPass.Caption = "First Pass"
    cmdFirstPass.Enabled = False
    cmdSecondPass.Caption = "Second Pass"
    cmdSecondPass.Enabled = False
    
End Sub

'--------------------------------------------------------------------
Private Sub CheckVerticalScroll(ByVal sglTop As Single, _
                                ByVal sglHeight As Single)
'--------------------------------------------------------------------
Dim nScrollAmount As Integer

    If VScroll.Visible Then
        If sglTop + (sglHeight * 2) + picDDEForm.Top > picDD.Height Then
            nScrollAmount = (sglTop + (sglHeight * 2) - picDD.Height) / 10
            If nScrollAmount < VScroll.Max Then
                VScroll.Value = nScrollAmount
            Else
                VScroll.Value = VScroll.Max
            End If
        ElseIf sglTop < -picDDEForm.Top Then
            If sglTop < picDD.Height Then
                VScroll.Value = VScroll.Min
            Else
                VScroll.Value = sglTop / 10
            End If
        End If
    End If

End Sub

'--------------------------------------------------------------------
Private Sub CheckHorizontalScroll(ByVal sglLeft As Single, _
                                    ByVal sglWidth As Single)
'--------------------------------------------------------------------
Dim nScrollAmount As Integer

    If HScroll.Visible Then
        If sglLeft + (sglWidth * 1.2) + picDDEForm.Left > picDD.Width Then
            nScrollAmount = (sglLeft + (sglWidth * 1.2) - picDD.Width) / 10
            If nScrollAmount < HScroll.Max Then
                HScroll.Value = nScrollAmount
            Else
                HScroll.Value = HScroll.Max
            End If
        ElseIf sglLeft < -picDDEForm.Left Then
            If sglLeft < picDD.Width Then
                HScroll.Value = HScroll.Min
            Else
                HScroll.Value = sglLeft / 10
            End If
        End If
    End If

End Sub

'--------------------------------------------------------------------
Private Sub SetFontAndColour()
'--------------------------------------------------------------------
Dim oControl As Control

    For Each oControl In Me.Controls
        If oControl.Name = "lblQuestionLabel" Or oControl.Name = "txtQuestionResponse" Or oControl.Name = "cmdAnotherRQGRow" _
        Or oControl.Name = "lblHeader" Or oControl.Name = "cmdRepeatEForm" Or oControl.Name = "cmdNextEForm" Or oControl.Name = "lblCatCodes" _
        Or oControl.Name = "cmdEndSession" Or oControl.Name = "cmdPreviousRepeat" Or oControl.Name = "cmdPreviousEForm" Then
            oControl.FontSize = gnRegDDFontSize
            Select Case oControl.Name
            Case "txtQuestionResponse"
                oControl.BackColor = glRegDDLightColour
            Case "lblHeader", "lblQuestionLabel", "lblCatCodes"
                oControl.BackColor = glRegDDMediumColour
            Case "cmdAnotherRQGRow", "cmdPreviousRepeat", "cmdRepeatEForm", "cmdPreviousEForm", "cmdNextEForm", "cmdEndSession"
                oControl.BackColor = glRegDDDarKColour
            End Select
        End If
    Next
    
    picDDEForm.BackColor = glRegDDMediumColour

End Sub

'--------------------------------------------------------------------
Private Sub AssessVisitDates()
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim lVisitDateeFormId As Long
Dim lVisitDateQuestionId As Long

    glCRFPageIdAfterVisitDate = 0
    mbVisitDateEform = False
    mbVisitDateEformExists = False
    mlVisitDateEformId = 0

    'Query for a Visit-Date-eForm existing in the current visit
    sSQL = "SELECT CRFPageId " _
        & "FROM StudyVisitCRFPage " _
        & "WHERE ClinicalTrialId = " & glSelTrialId & " " _
        & "AND VisitId = " & glSelVisitId & " " _
        & "AND eFormUse = 1"
        
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        'No Visit Date eForm exists for this visit
        rsTemp.Close
        Set rsTemp = Nothing
        Exit Sub
    Else
        lVisitDateeFormId = rsTemp!CRFPageId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    'Query for the Visit-Date-Question on the Visit-Date-eForm
    sSQL = "SELECT DataItemId " _
        & "FROM CRFElement " _
        & "WHERE ClinicalTrialId = " & glSelTrialId & " " _
        & "AND CRFPageId = " & lVisitDateeFormId & " " _
        & "AND ElementUse = 1"

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        'This should not occur,
        'If a Visit-Date-eForm exists there should be a Visit-Date-Question
        rsTemp.Close
        Set rsTemp = Nothing
        Exit Sub
    Else
        lVisitDateQuestionId = rsTemp!DataItemId
    End If

    rsTemp.Close
    Set rsTemp = Nothing
    
    If QuestionIsDerived(glSelTrialId, lVisitDateQuestionId) Then
        'Its a derived Visit-Date-Question, nothing needs to be done
        Exit Sub
    End If
    
    'Check for the Visit-Date-Question being already verified
    sSQL = "SELECT Status FROM DoubleData " _
        & " WHERE ClinicalTrialId = " & glSelTrialId _
        & " AND TrialSite = '" & gsSelSite & "'" _
        & " AND PersonId = " & glSelPersonId _
        & " AND VisitId = " & glSelVisitId _
        & " AND VisitCycleNumber = " & glSelVisitCycleNumber _
        & " AND CRFPageId = " & lVisitDateeFormId _
        & " AND CRFPageCycleNumber = 1" _
        & " AND DataItemId = " & lVisitDateQuestionId _
        & " AND RepeatNumber = 1"
        
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount > 0 Then
        If rsTemp!Status = eDoubleDataStatus.Verified Then
            'The Visit-Date-Question has already been entered and verified
            'So exit without causing it to be asked again
            rsTemp.Close
            Set rsTemp = Nothing
            Exit Sub
        End If
    End If
   
    rsTemp.Close
    Set rsTemp = Nothing
    
    'store CRFPageId of the form that was going to be built
    'and will now be built after the Visit-Date-EForm
    glCRFPageIdAfterVisitDate = glSelCRFPageId
    mbVisitDateEform = True
    mbVisitDateEformExists = True
    mlVisitDateEformId = lVisitDateeFormId
    'overwrite glSelCRFPageId with the Visit-Date-EForm Id
    glSelCRFPageId = lVisitDateeFormId
    gsSelCRFPageCode = CRFPageCodeFromId(glSelTrialId, glSelCRFPageId)

End Sub

'--------------------------------------------------------------------
Private Sub DisableCombos()
'--------------------------------------------------------------------

    cboStudy.Enabled = False
    cboSite.Enabled = False
    cboSubjects.Enabled = False
    cboVisit.Enabled = False
    cboVisitInstance.Enabled = False
    cboEForm.Enabled = False
    cboEFormInstance.Enabled = False
    cmdFirstPass.Enabled = False
    cmdSecondPass.Enabled = False

End Sub

'---------------------------------------------------------------------
Public Sub GetRegistrySettings()
'---------------------------------------------------------------------

    'Get the window state of the last opened Double Date Entry window
    gnRegFormWindowState = GetMACROSetting("DDWindowState", 0)
    
    'Get non-max co-ordinates for the Double Date Entry window
    gnRegFormLeft = GetMACROSetting("DDFormLeft", Abs(Screen.Width - mnMINFORMWIDTH) / 2)
    gnRegFormTop = GetMACROSetting("DDFormTop", Abs(Screen.Height - mnMINFORMHEIGHT) / 2)
    gnRegFormWidth = GetMACROSetting("DDFormWidth", mnMINFORMWIDTH)
    gnRegFormHeight = GetMACROSetting("DDFormHeight", mnMINFORMHEIGHT)
    
    'Get max co-ordinates for the Double Date Entry window
    gnRegFormLeftMax = GetMACROSetting("DDFormLeftMax", 0)
    gnRegFormTopMax = GetMACROSetting("DDFormTopMax", 0)
    gnRegFormWidthMax = GetMACROSetting("DDFormWidthMax", Abs(Screen.Width))
    gnRegFormHeightMax = GetMACROSetting("DDFormHeightMax", Abs(Screen.Height))
    
    'Get the saved colours for the Double Data Entry window, else default grey colours
    glRegDDLightColour = GetMACROSetting("DDLightColour", RGB(244, 244, 244))
    glRegDDMediumColour = GetMACROSetting("DDMediumColour", RGB(228, 228, 228))
    glRegDDDarKColour = GetMACROSetting("DDDarkColour", RGB(212, 212, 212))
    
    'get the saved fontsize for the Double Data Entry window
    gnRegDDFontSize = GetMACROSetting("DDFontSize", 8)
    
    'Get the saved colour scheme
    gsRegDDColourScheme = GetMACROSetting("DDColourScheme", "Greys")
    
    'Get the Display Catagory Codes setting
    gbRegDisplayCategoryCodes = GetMACROSetting("DDDispayCategoryCodes", True)
    
 End Sub

'---------------------------------------------------------------------
Public Sub SaveRegistrySettings()
'---------------------------------------------------------------------

    Call SetMACROSetting("DDWindowState", gnRegFormWindowState)
    Call SetMACROSetting("DDFormLeft", gnRegFormLeft)
    Call SetMACROSetting("DDFormTop", gnRegFormTop)
    Call SetMACROSetting("DDFormWidth", gnRegFormWidth)
    Call SetMACROSetting("DDFormHeight", gnRegFormHeight)
    Call SetMACROSetting("DDFormLeftMax", gnRegFormLeftMax)
    Call SetMACROSetting("DDFormTopMax", gnRegFormTopMax)
    Call SetMACROSetting("DDFormWidthMax", gnRegFormWidthMax)
    Call SetMACROSetting("DDFormHeightMax", gnRegFormHeightMax)

End Sub

'---------------------------------------------------------------------
Private Sub DDeFormChecks()
'---------------------------------------------------------------------

    If Me.WindowState = vbMaximized Then
        Exit Sub
    End If

    'Increase width of form if required
    If picDDEForm.Width > picDD.Width Then
        If picDDEForm.Width + 350 > (Screen.Width) Then
            'Window set to Max width
            picDD.Width = Screen.Width - 300
            frmMenu.Width = picDD.Width + 300
            DDFormCentre Me
        Else
            'Window increased in width
            picDD.Width = picDDEForm.Width + 50
            frmMenu.Width = picDD.Width + 300
            DDFormCentre Me
        End If
    End If
    
    'Increase height of form if required
    If picDDEForm.Height > picDD.Height Then
        If picDDEForm.Height + 3250 > (Screen.Height - 400) Then
            'Window set to Max height
            picDD.Height = Screen.Height - 3600
            frmMenu.Height = picDD.Height + 3200
            DDFormCentre Me
        Else
            'Window increased in height
            picDD.Height = picDDEForm.Height + 50
            frmMenu.Height = picDD.Height + 3200
            DDFormCentre Me
        End If
    End If

End Sub

'---------------------------------------------------------------------
Public Sub DDeFormScrollbars()
'---------------------------------------------------------------------

    'Assess the need for vertical scrollbar
    If VScroll.Visible = True Then
        VScroll.Max = (picDDEForm.Height - picDD.Height) / 10
        VScroll.LargeChange = CInt(((VScroll.Max / 2) / 10) + 1)
        VScroll.SmallChange = CInt(((VScroll.Max / 10) / 10) + 1)
    Else
        If picDDEForm.Height > picDD.Height Then
            Set VScroll.Container = picDD
            VScroll.Max = (picDDEForm.Height - picDD.Height) / 10
            VScroll.Top = 0
            If picDDEForm.Width > picDD.Width Then
                VScroll.Height = picDD.ScaleHeight - HScroll.Height
            Else
                VScroll.Height = picDD.ScaleHeight
            End If
            VScroll.Left = picDD.ScaleWidth - VScroll.Width
            VScroll.LargeChange = CInt(((VScroll.Max / 2) / 10) + 1)
            VScroll.SmallChange = CInt(((VScroll.Max / 10) / 10) + 1)
            VScroll.Value = 0
            VScroll.Visible = True
        End If
    End If
    
    'Assess the need for horizontal scrollbar
    If picDDEForm.Width > picDD.Width Then
        Set HScroll.Container = picDD
        HScroll.Max = (picDDEForm.Width - picDD.Width) / 10
        HScroll.Top = picDD.ScaleHeight - HScroll.Height
        If VScroll.Visible = True Then
            HScroll.Width = picDD.ScaleWidth - VScroll.Width
        Else
            HScroll.Width = picDD.ScaleWidth
        End If
        HScroll.Left = 0
        HScroll.LargeChange = CInt(((HScroll.Max / 2) / 10) + 1)
        HScroll.SmallChange = CInt(((HScroll.Max / 10) / 10) + 1)
        HScroll.Value = 0
        HScroll.Visible = True
    End If

End Sub

'---------------------------------------------------------------------
Public Sub UnlockBatchUpload()
'---------------------------------------------------------------------
Dim sSQL As String

    sSQL = "DELETE FROM MACROUserSettings " _
        & "WHERE UserName = 'BUL' " _
        & "AND UserSetting = 'Batch Upload Lock'"
    MacroADODBConnection.Execute sSQL
        
End Sub

'---------------------------------------------------------------------
Private Function GetCatCodes(lDataItemId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
Dim sSQL

    sSQL = "SELECT ValueCode, ItemValue FROM ValueData" _
        & " WHERE ClinicalTrialId = " & glSelTrialId _
        & " AND VersionId = 1" _
        & " AND DataItemId = " & lDataItemId _
        & " AND Active = 1" _
        & " ORDER BY ValueOrder"
    Set GetCatCodes = New ADODB.Recordset
    GetCatCodes.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
End Function
