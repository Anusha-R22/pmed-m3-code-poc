VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCUIFunctionEditor 
   Caption         =   "Expression Builder"
   ClientHeight    =   8385
   ClientLeft      =   1710
   ClientTop       =   1680
   ClientWidth     =   11085
   Icon            =   "frmCUIFunctionEditor.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8385
   ScaleWidth      =   11085
   Begin VB.PictureBox picBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   10890
      TabIndex        =   41
      Top             =   3120
      Width           =   10890
   End
   Begin VB.Frame fraControlsContainer 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   10935
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5880
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdMatchBrackets 
         Caption         =   "Match brackets"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2160
         TabIndex        =   2
         Top             =   3480
         Width           =   1365
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "Check"
         Height          =   330
         Left            =   9915
         TabIndex        =   3
         Top             =   3480
         Width           =   1000
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   51
         Left            =   6480
         TabIndex        =   34
         Tag             =   "3"
         ToolTipText     =   "3"
         Top             =   720
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   57
         Left            =   5640
         TabIndex        =   33
         Tag             =   "1"
         ToolTipText     =   "1"
         Top             =   720
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   58
         Left            =   6060
         TabIndex        =   32
         Tag             =   "2"
         ToolTipText     =   "2"
         Top             =   720
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   70
         Left            =   6060
         TabIndex        =   31
         Tag             =   "+"
         ToolTipText     =   "Numerical addition (do not use with strings)"
         Top             =   1935
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   5640
         TabIndex        =   30
         Tag             =   "-"
         ToolTipText     =   "Numerical subtraction (do not use with dates)"
         Top             =   1935
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   5640
         TabIndex        =   29
         Tag             =   "*"
         ToolTipText     =   "Multiplication"
         Top             =   1590
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   6060
         TabIndex        =   28
         Tag             =   "/"
         ToolTipText     =   "Division"
         Top             =   1590
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   64
         Left            =   6480
         TabIndex        =   27
         Tag             =   "9"
         ToolTipText     =   "9"
         Top             =   0
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   44
         Left            =   6060
         TabIndex        =   42
         Tag             =   "8"
         ToolTipText     =   "8"
         Top             =   0
         Width           =   350
      End
      Begin VB.ComboBox cboDecisions 
         Height          =   315
         Left            =   6915
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   216
         Width           =   4000
      End
      Begin VB.ComboBox cboTimeUnits 
         Height          =   315
         Left            =   6915
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2985
         Width           =   4000
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Index           =   59
         Left            =   6480
         TabIndex        =   24
         Tag             =   ":"
         ToolTipText     =   "Separator for compound data names"
         Top             =   1935
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   56
         Left            =   5640
         TabIndex        =   23
         Tag             =   "0"
         ToolTipText     =   "0"
         Top             =   1080
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   55
         Left            =   6060
         TabIndex        =   22
         Tag             =   "."
         ToolTipText     =   "decimal point"
         Top             =   1080
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   54
         Left            =   6480
         TabIndex        =   21
         Tag             =   "6"
         ToolTipText     =   "6"
         Top             =   360
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   53
         Left            =   6060
         TabIndex        =   20
         Tag             =   "5"
         ToolTipText     =   "5"
         Top             =   360
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   52
         Left            =   5640
         TabIndex        =   19
         Tag             =   "4"
         ToolTipText     =   "4"
         Top             =   360
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   ","
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   50
         Left            =   6480
         TabIndex        =   18
         Tag             =   ","
         ToolTipText     =   "Comma, list separator"
         Top             =   1080
         Width           =   350
      End
      Begin VB.ComboBox cboDataValues 
         Height          =   315
         Left            =   6915
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2424
         Width           =   4000
      End
      Begin VB.ComboBox cboOtherTasks 
         Height          =   315
         Left            =   6915
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1320
         Width           =   4000
      End
      Begin VB.ComboBox cboCandidates 
         Height          =   315
         Left            =   6915
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   768
         Width           =   4000
      End
      Begin VB.ComboBox cboDataItems 
         Height          =   315
         Left            =   6915
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1860
         Width           =   4005
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   14
         Left            =   5640
         TabIndex        =   43
         Tag             =   "7"
         ToolTipText     =   "7"
         Top             =   0
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   6060
         TabIndex        =   13
         Tag             =   " )"
         ToolTipText     =   "Right parenthesis"
         Top             =   2280
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "("
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   5640
         TabIndex        =   12
         Tag             =   "( "
         ToolTipText     =   "Left parenthesis"
         Top             =   2280
         Width           =   350
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   330
         Left            =   0
         TabIndex        =   0
         Top             =   3480
         Width           =   1000
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   1080
         TabIndex        =   1
         Top             =   3480
         Width           =   1000
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   6480
         TabIndex        =   11
         Tag             =   "_"
         ToolTipText     =   "Underscore"
         Top             =   1590
         Width           =   350
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "me:value"
         Height          =   300
         Index           =   4
         Left            =   5640
         TabIndex        =   10
         Tag             =   "me:value"
         ToolTipText     =   "Current question"
         Top             =   2640
         Width           =   1185
      End
      Begin VB.Frame fraButtonContainer 
         Caption         =   "Caption "
         Height          =   660
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   5295
         Begin VB.CommandButton cmdFunction 
            Appearance      =   0  'Flat
            Caption         =   "button"
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Tag             =   "function( "
            ToolTipText     =   "Used for multiple conditional evaluation"
            Top             =   240
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin MSComctlLib.TabStrip tabFunctionSelector 
         Height          =   3300
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   5821
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblDecisionsOrVisits 
         Caption         =   "Decisions"
         Height          =   195
         Left            =   6915
         TabIndex        =   40
         Top             =   0
         Width           =   4005
      End
      Begin VB.Label Label5 
         Caption         =   "Time Units"
         Height          =   195
         Left            =   6915
         TabIndex        =   39
         Top             =   2760
         Width           =   4005
      End
      Begin VB.Label Label4 
         Caption         =   "Data Values"
         Height          =   195
         Left            =   6915
         TabIndex        =   38
         Top             =   2208
         Width           =   4005
      End
      Begin VB.Label Label3 
         Caption         =   "Other Tasks"
         Height          =   195
         Left            =   6915
         TabIndex        =   37
         Top             =   1104
         Width           =   4005
      End
      Begin VB.Label lblCandidatesOrForms 
         Caption         =   "Candidates"
         Height          =   195
         Left            =   6915
         TabIndex        =   36
         Top             =   552
         Width           =   4005
      End
      Begin VB.Label Label1 
         Caption         =   "Questions"
         Height          =   195
         Left            =   6915
         TabIndex        =   35
         Top             =   1656
         Width           =   4005
      End
   End
   Begin VB.TextBox txtHelp 
      BackColor       =   &H80000018&
      Height          =   1140
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3240
      Width           =   10890
   End
   Begin VB.TextBox txtExpression 
      Height          =   2985
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   120
      Width           =   10890
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditOpt 
         Caption         =   "&Cut"
         Index           =   0
      End
      Begin VB.Menu mnuEditOpt 
         Caption         =   "C&opy"
         Index           =   1
      End
      Begin VB.Menu mnuEditOpt 
         Caption         =   "&Paste"
         Index           =   2
      End
   End
   Begin VB.Menu mnuCheck 
      Caption         =   "&Check"
      Begin VB.Menu mnuCheckOpt 
         Caption         =   "Chec&k"
         Index           =   0
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsOpt 
         Caption         =   "&Add Right Brackets"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuOptionsOpt 
         Caption         =   "Add Co&lons"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuOptionsFont 
         Caption         =   "Font..."
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmCUIFunctionEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999-2007. All Rights Reserved
'   File:       frmCUIFunctionEditor.frm
'   Language:   VB6
'   Author:     Mo Morris 1999, Robert Williams 2002
'   Purpose:    Allows user to build expressions and conditions.
'               THIS VERSION FOR MACRO 3.0
'
'   Constraints: This form is a COPY of the smae form in the Arezzo Composer project (CUI.vpb)
'               We should aim to keep it in sync. with its sibling.
'               This form uses the goALM variable, assumed to represent an ALM5 instance
'--------------------------------------------------------------------'
'   Mo Morris       14/9/99
'   Form Load changed to handle differences between standalone CUI and
'   Macro/MTMCUI, namely the Decisions/Visits combo, the Candidates/Forms
'   combo and forms Top/Left position.
'   Mo Morris       16/9/99
'   cmdOK_Click now checks for a non-blank expression/condition prior to
'   calling CLM.DLL to validate it.
'   Mo Morris       22/10/99
'   command buttons for datetime, datediff and quotes(") removed.
'   command buttons timenow, datenow, date_diff, time_diff, abs, between,
'   status, case, else and ':' added together with their relevant help text
'   information in sub DisplayHelp
'   willC  11/10/99   Added the error handlers
'   NCJ 2 Dec 99    Rewrote cmdOK_Click
'   Mo Morris       13/12/99
'   'date' and 'time' function buttons added
'   NCJ 26/9/00 - Set mctlSourceOfExpression to Nothing in Form_Unload
'   Mo Morris   4/12/00 This form has been re-linked to the CUI 2.1 version of this form
'               Tab order changed.
'   Mo Morris   5/12/00 Form_Load changed so that the population of cboOtherTasks is now governed
'               by the conditional compilation.
'               In MACRO and MTMCUI.exe (Not StandAlone) it will be populated with all of the
'               non-Visit and non-Form tasks.
'               In CUI.exe (StandAlone) it will be populated with all of the non-Decision tasks.
' NCJ 19 Feb 01 - Changed help text for "like"
'---------------------------------------------------------------------
' CUI 5.0
'   RCW 05 Mar 02 - Changed calls to getsetting to getCUIsetting
'   RCW 06 Mar 02 - Changed calls to savesetting to saveCUIsetting
'
' New Expression Editor
'   RCW 29 Jul 02 - 07 Aug 02 Added new tabbing mechanism for functions
'   RCW 08 Aug 02 New help mechanism for tabs and buttons via Prolog
'   RCW 09 Aug 02 Double click mechanism
'
' MACRO 3.0
'   NCJ 30 Aug 02 - This is now a COPY of original in CUI 5.0, for MACRO 3.0
'   NCJ 31 Oct 02 - Filter DataItems combo on selected eForm
'
'   RCW Aug - Sep 02 Addition of Expression Editor Assistance features
'  (E.A) Oct 22 - Added extra expression/condition checks for the Expression Editor
'   RCW 05 Dec 02 - Tab order
'   RCW 18 Dec 02 - Prevent large text being sent to ALM
'   RCW 18 Dec 02 - Prevent backwards quotes being sent to ALM
'   RCW 19 Dec 02 - Centre screen on load & fixed MatchBracket problem
'                   also removed dbl click for matchbracket
'
'   NCJ 16/17 Jan 03 - Incorporated latest Composer changes into this MACRO version
'                   Changed calls to GetCUIsetting to GetMACROsetting
'                   Changed calls to SaveCUIsetting to SetMACROsetting
'   NCJ 23 Jan 03 - Initialise the focus to the txtExpression field for MACRO
'   NCJ 27 Jan 03 - Added sArezzoType to Initialise routine
'   NCJ 23/24 Apr 03 - Get combo boxes working properly by adding KeyPress, KeyDown and GotFocus events
'   NCJ 7 May 03 - Disallow usual MACRO disallowed characters
'   NCJ 14 May 03 - MACRO limit on size of expressions changed from 2000 to 4000
'   NCJ 2 Jul 03 - Fixed bug in call to new term checker
'   NCJ 3 Jan 06 - Allow user to select font
'   NCJ 19 Jan 06 - Only set font details in Form_Load if there are some
'   NCJ 22 Feb 06 - Avoid crash in Font details if Font name is empty!!
'   NCJ 19 Sept 06 - Implemented "read-only" mode, based on StudyAccessMode
'   NCJ 20 Sept 07 - Issue 2947, HotH 1450 (KKS) - Store font size as "standard" (may be fractional)
'--------------------------------------------------------------------'

Option Explicit

Dim mlngRegThisFormLeft As Long 'renamed from gnRegThisFormLeft and _Top to
Dim mlngRegThisFormTop As Long  'conform closer to coding standards

Private mcolTabs As Collection 'contains compressed string representation of tabs

Option Compare Binary
Option Base 0

Private mctlSourceOfExpression As Control   'control to which result will be returned
Private msValidationType As String

Private msglFrameWidth As Single
Private mbMouseDrag As Boolean

Private msLastButton As String
Private msDClickString As String

Private mnTabIndex As Integer

'audit trail for undo
Private mnCutOpt As Integer
Private mnCopyOpt As Integer
Private mnPasteOpt As Integer
Private msKeyPressed As String

'expression editor assistance variables
Private mbSetofFlag As Boolean
Private mbBracketOK As Boolean
Private mbColonOK As Boolean
Private msCurrentButton As String

'Variables used to expand text to include brackets
Private msLastSelection As String
Private mlngLastSelStart As Long
Private mlngLastSelLength As Long
Private msLastExpansion As String

' NCJ 23 Apr 03 - Get combo boxes working properly
Private mbComboKeyPressed As Boolean
Private msLastComboItem As String

Private msErrorMsg As String

'constants for screen building
Private Const mnSPACER = 50
Private Const mnVERTICALSPACER = 20
Private Const mnBUTTONHEIGHT = 300
Private Const mnFRAMEPADDING = 350
Private Const msglFRAMEWIDTHPADDING = 150
Private Const msglCONTROLBOXHEIGHTFROMBOTTOM = 3960 'distance of top of control
                                                    'container from bottom of form
Private Const msglDISTHELPBTMFORMFORMBTM = 4850 'distance of bottom of help area
                                                'from bottom of form
Private Const msSIZE_ERROR = "Expression is too large to be processed"

' NCJ 27 Jan 03
Private msTermType As String

' NCJ 19 Sept 06 - Readonlyness
Private mbCanEdit As Boolean

'---------------------------------------------------------------------
Public Sub Initialise(ByRef vControl As Control, ByVal ValidationType As String, _
                        ByVal bCanEdit As Boolean, Optional sTitle As String, _
                        Optional sArezzoType As String = "")
'---------------------------------------------------------------------
' Initialises the expression or condition to be edited
' NCJ 26/9/00 - Changed ByVal to ByRef
' NCJ 27 Jan 03 - Added sArezzoType to specify a term type, e.g. "string", "numeric", "temporal"
' NCJ 21 Sept 06 - Set up mbCanEdit based on new bCanEdit argument
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    ' Can edit if R/W or Full Control
'    mbCanEdit = (neAccessMode >= sdReadWrite)
    mbCanEdit = bCanEdit
    
    'copy the validation type into module variable ("Expression" or "Condition")
    msValidationType = ValidationType
    msErrorMsg = "This is not a valid "
    If msValidationType = "Expression" Then
        msErrorMsg = msErrorMsg & "expression"
    Else
        msErrorMsg = msErrorMsg & "condition"
    End If
    
    'copy the contents of the calling control into the expression window
    Set mctlSourceOfExpression = vControl
    
    ' NCJ 27 Jan 03 - Remember the term type if valid
    Select Case sArezzoType
    Case "string", "numeric", "temporal"
        msTermType = sArezzoType
    Case Else
        msTermType = ""
    End Select
    
    txtExpression.Text = vControl.Text
    txtExpression.SelStart = Len(txtExpression.Text)
    
    If mbCanEdit Then
        Me.Caption = "Editing: " & sTitle       'set title
    Else
        Me.Caption = "Viewing: " & sTitle       'set title
    End If
    txtExpression.Locked = Not mbCanEdit
    cmdOK.Enabled = mbCanEdit
   
    'set edit menu option numbers
    mnCutOpt = 0
    mnCopyOpt = 1
    mnPasteOpt = 2
    
    'initialise assistance variables
    mbBracketOK = True 'this makes it possible to turn off suggested text
                            'such as includes or = after dataitems
    mbColonOK = True
    mlngLastSelStart = 0 'this is the start of any expanded text
    
    'get autocomplete value from settings file
    mbBracketOK = (GetMACROSetting("AutoCompleteBracket", "On") = "On")
    mbColonOK = (GetMACROSetting("AutoCompleteColon", "On") = "On")
    mnuOptionsOpt(0).Checked = mbBracketOK
    mnuOptionsOpt(1).Checked = mbColonOK

    Me.Show vbModal
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Initialise")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub AddTextToExpression(ByRef sText As String)
'---------------------------------------------------------------------
' NCJ 19 Sept 06 - Only if they can edit
'---------------------------------------------------------------------

    If Not mbCanEdit Then Exit Sub
    
    'Replaces the currently selected text in the text box with the new string.
    txtExpression.SelText = sText
    
    'add closing bracket
    If InStr(sText, "(") > 0 And mbBracketOK Then
        Call CloseBracket("button", "(")
    End If
    If InStr(sText, "[") > 0 And mbBracketOK Then
        Call CloseBracket("button", "[")
    End If

    txtExpression.SetFocus
    
    ' allow button to be pressed again
    msLastButton = sText
 
End Sub

'---------------------------------------------------------------------
Private Sub cboCandidates_Click()
'---------------------------------------------------------------------
' NB This is the "eForms" combo in MACRO
' NCJ 31 Oct 02 - Update DataItems combo with just the questions on the selected eForm
'---------------------------------------------------------------------

#If Not StandAloneCUI Then
Dim rsQuestions As ADODB.Recordset
#End If

    On Error GoTo ErrHandler

    ' NCJ 23 Apr 03 - Do not process the Click event if they've selected with a key
    If mbComboKeyPressed = True Then
        mbComboKeyPressed = False
        msLastComboItem = cboCandidates.Text
        Exit Sub
    End If
    
    If cboCandidates.Text = "" Then
        ' Repopulate the Questions combo
        Call FillDataItemsCombo
        Exit Sub
    End If
    
    AddTextToExpression cboCandidates.Text
    
    #If Not StandAloneCUI Then
        
        'also add the colon - RCW 2/9/02
        If mbColonOK Then
            AddTextToExpression ":"
        End If
        
        'change the contents of cboDataItems to those within clicked eForm
        Set rsQuestions = gdsCRFPageDataItems(frmMenu.ClinicalTrialId, frmMenu.VersionId, _
                            cboCandidates.ItemData(cboCandidates.ListIndex))
        cboDataItems.Clear
        If rsQuestions.RecordCount > 0 Then
            rsQuestions.MoveFirst
            While Not rsQuestions.EOF
                If rsQuestions!DataItemId > 0 Then
                    cboDataItems.AddItem rsQuestions!DataItemCode
                    cboDataItems.ItemData(cboDataItems.NewIndex) = rsQuestions!DataItemId
                End If
                rsQuestions.MoveNext   'get next record
            Wend
        End If
        Set rsQuestions = Nothing
    #End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboCandidates_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub cboCandidates_GotFocus()
'---------------------------------------------------------------------

    msLastComboItem = cboCandidates.Text

End Sub

Private Sub cboCandidates_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
' Prevent automatic selection with arrow keys
'---------------------------------------------------------------------

    Select Case KeyCode
    Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight
        ' Remember they selected with a key
        mbComboKeyPressed = True
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub cboCandidates_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' Let them select with a key but don't treat as a Click event unless it's Return
'---------------------------------------------------------------------

    If KeyAscii = vbKeyReturn Then
        ' Simulate a click if it's the same as they last selected
        If msLastComboItem = cboCandidates.Text Then
            mbComboKeyPressed = False
            Call cboCandidates_Click
        End If
    Else
        ' Remember they selected with a key
        mbComboKeyPressed = True
    End If

End Sub

'---------------------------------------------------------------------
Private Sub cboDataItems_Click()
'---------------------------------------------------------------------
Dim msDataItemKey As String
Dim clmDataItem As DataItem
Dim clmCollection As Collection
Dim msValue As Variant

    On Error GoTo ErrHandler
    
    ' NCJ 23 Apr 03 - Do not process the Click event if they've selected with a key
    If mbComboKeyPressed = True Then
        msLastComboItem = cboDataItems.Text
        mbComboKeyPressed = False
        Exit Sub
    End If
    
    If cboDataItems.Text = "" Then Exit Sub
    
    AddTextToExpression cboDataItems.Text


    'change the contents of cboDataValues to values pertaining to clicked data item
    msDataItemKey = gclmGuideline.colDataItems.GetDataItemKey(cboDataItems.Text)
    Set clmDataItem = gclmGuideline.colDataItems.Item(msDataItemKey)
    cboDataValues.Clear
    If clmDataItem.DataItemType = "boolean" Then
        cboDataValues.AddItem clmDataItem.TrueValue
        cboDataValues.AddItem clmDataItem.FalseValue
    Else
        'copy contents of range of values to combo
        Set clmCollection = clmDataItem.RangeValues
        For Each msValue In clmCollection
            cboDataValues.AddItem msValue
        Next
        If clmCollection.Count = 0 And clmDataItem.DefaultValue <> "" Then
            cboDataValues.AddItem clmDataItem.DefaultValue
        End If
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboDataItems_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub cboDataItems_GotFocus()
'---------------------------------------------------------------------

    ' Remember the last combo item selected
    msLastComboItem = cboDataItems.Text

End Sub

'---------------------------------------------------------------------
Private Sub cboDataItems_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
' Prevent automatic selection with arrow keys
'---------------------------------------------------------------------

    Select Case KeyCode
    Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight
        ' Remember they selected with a key
        mbComboKeyPressed = True
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub cboDataItems_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' Let them select with a key but don't treat as a Click event unless it's Return
'---------------------------------------------------------------------

    If KeyAscii = vbKeyReturn Then
        ' Simulate a click if it's the same as they last selected
        If msLastComboItem = cboDataItems.Text Then
            mbComboKeyPressed = False
            Call cboDataItems_Click
        End If
    Else
        ' Remember they selected with a key
        mbComboKeyPressed = True
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub cboDataValues_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    ' NCJ 23 Apr 03 - Do not process the Click event if they've selected with a key
    If mbComboKeyPressed = True Then
        mbComboKeyPressed = False
        msLastComboItem = cboDataValues.Text
        Exit Sub
    End If
    
    If cboDataValues.Text = "" Then Exit Sub
    
    AddTextToExpression cboDataValues.Text
    
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboDataValues_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub cboDataValues_GotFocus()
'---------------------------------------------------------------------

    msLastComboItem = cboDataValues.Text

End Sub

'---------------------------------------------------------------------
Private Sub cboDataValues_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
' Prevent automatic selection with arrow keys
'---------------------------------------------------------------------

    Select Case KeyCode
    Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight
        ' Remember they selected with a key
        mbComboKeyPressed = True
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub cboDataValues_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' Let them select with a key but don't treat as a Click event unless it's Return
'---------------------------------------------------------------------

    If KeyAscii = vbKeyReturn Then
        ' Simulate a click if it's the same as they last selected
        If msLastComboItem = cboDataValues.Text Then
            mbComboKeyPressed = False
            Call cboDataValues_Click
        End If
    Else
        ' Remember they selected with a key
        mbComboKeyPressed = True
    End If

End Sub

'---------------------------------------------------------------------
Private Sub cboDecisions_Click()
'---------------------------------------------------------------------
' NB This is the "Visits" combo within MACRO
'---------------------------------------------------------------------
Dim clmtask As Task
Dim clmCollection As Collection
Dim vntItem As Variant
Dim clmVisitTask As Task
Dim sComponentKey As Variant

    On Error GoTo ErrHandler

    ' NCJ 23 Apr 03 - Do not process the Click event if they've selected with a key
    If mbComboKeyPressed = True Then
        mbComboKeyPressed = False
        msLastComboItem = cboDecisions.Text
        Exit Sub
    End If
    
    If cboDecisions.Text = "" Then Exit Sub
    
    AddTextToExpression cboDecisions.Text
    
    cboCandidates.Clear
    
    #If Not StandAloneCUI Then
        
        'also add the colon - RCW 2/9/02
        If mbColonOK Then
            AddTextToExpression ":"
        End If
        
        'change the contents of cboCandidates to those forms within clicked visit
        Set clmVisitTask = gclmGuideline.colTasks.Item(cboDecisions.ItemData(cboDecisions.ListIndex))
        For Each sComponentKey In clmVisitTask.Components
            Set clmtask = gclmGuideline.colTasks.Item(sComponentKey)
            'check for plans that are Macro created (i.e. locked) with 'Form' at the start of their caption
            If clmtask.TaskType = "plan" Then
                If clmtask.Locked Then
                    If Mid(clmtask.Caption, 1, 4) = "Form" Then
                        cboCandidates.AddItem clmtask.Name
                        ' NCJ 31 Oct 02 - Need to store TaskKey too
                        cboCandidates.ItemData(cboCandidates.NewIndex) = clmtask.TaskKey
                    End If
                End If
            End If
        Next
        ' Select no eForm and repopulate DataItems combo
        cboCandidates.ListIndex = -1
        Call FillDataItemsCombo
    #Else
        'change the contents of cboCandidates to those pertaining to the clicked decision
        Set clmtask = gclmGuideline.colTasks.Item(cboDecisions.ItemData(cboDecisions.ListIndex))
        'populate candidates listbox
        Set clmCollection = clmtask.Candidates
        For Each vntItem In clmCollection
            cboCandidates.AddItem vntItem
        Next
    #End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboDecisions_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub cboDecisions_GotFocus()
'---------------------------------------------------------------------

    msLastComboItem = cboDecisions.Text

End Sub

'---------------------------------------------------------------------
Private Sub cboDecisions_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
' Prevent automatic selection with arrow keys
'---------------------------------------------------------------------

    Select Case KeyCode
    Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight
        ' Remember they selected with a key
        mbComboKeyPressed = True
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub cboDecisions_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' Let them select with a key but don't treat as a Click event unless it's Return
'---------------------------------------------------------------------

    If KeyAscii = vbKeyReturn Then
        ' Simulate a click if it's the same as they last selected
        If msLastComboItem = cboDecisions.Text Then
            mbComboKeyPressed = False
            Call cboDecisions_Click
        End If
    Else
        ' Remember they selected with a key
        mbComboKeyPressed = True
    End If

End Sub

'---------------------------------------------------------------------
Private Sub cboOtherTasks_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' NCJ 23 Apr 03 - Do not process the Click event if they've selected with a key
    If mbComboKeyPressed = True Then
        mbComboKeyPressed = False
        msLastComboItem = cboOtherTasks.Text
        Exit Sub
    End If
    
    If cboOtherTasks.Text = "" Then Exit Sub
    
    AddTextToExpression cboOtherTasks.Text
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboOtherTasks_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    

End Sub

'---------------------------------------------------------------------
Private Sub cboOtherTasks_GotFocus()
'---------------------------------------------------------------------

    msLastComboItem = cboOtherTasks.Text

End Sub

'---------------------------------------------------------------------
Private Sub cboOtherTasks_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
' Prevent automatic selection with arrow keys
'---------------------------------------------------------------------

    Select Case KeyCode
    Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight
        ' Remember they selected with a key
        mbComboKeyPressed = True
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub cboOtherTasks_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' Let them select with a key but don't treat as a Click event unless it's Return
'---------------------------------------------------------------------

    If KeyAscii = vbKeyReturn Then
        ' Simulate a click if it's the same as they last selected
        If msLastComboItem = cboOtherTasks.Text Then
            mbComboKeyPressed = False
            Call cboOtherTasks_Click
        End If
    Else
        ' Remember they selected with a key
        mbComboKeyPressed = True
    End If

End Sub

'---------------------------------------------------------------------
Private Sub cboTimeUnits_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' NCJ 23 Apr 03 - Do not process the Click event if they've selected with a key
    If mbComboKeyPressed = True Then
        mbComboKeyPressed = False
        msLastComboItem = cboTimeUnits.Text
        Exit Sub
    End If
    
    If cboTimeUnits.Text = "" Then Exit Sub
    
    AddTextToExpression cboTimeUnits.Text
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboTimeUnits_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    

End Sub

'---------------------------------------------------------------------
Private Sub cboTimeUnits_GotFocus()
'---------------------------------------------------------------------
    
    msLastComboItem = cboTimeUnits.Text

End Sub

'---------------------------------------------------------------------
Private Sub cboTimeUnits_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
' Prevent automatic selection with arrow keys
'---------------------------------------------------------------------

    Select Case KeyCode
    Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight
        ' Remember they selected with a key
        mbComboKeyPressed = True
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub cboTimeUnits_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' Let them select with a key but don't treat as a Click event unless it's Return
'---------------------------------------------------------------------

    If KeyAscii = vbKeyReturn Then
        ' Simulate a click if it's the same as they last selected
        If msLastComboItem = cboTimeUnits.Text Then
            mbComboKeyPressed = False
            Call cboTimeUnits_Click
        End If
    Else
        ' Remember they selected with a key
        mbComboKeyPressed = True
    End If

End Sub

'---------------------------------------------------------------------
Private Sub cmdButton_Click(Index As Integer)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    If mbCanEdit Then
        Call CondAddTextToExpression(cmdButton(Index).Caption, cmdButton(Index).Tag)
    End If
    cmdButton(Index).Value = False
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdButton_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub cmdButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
    
    Call DisplayHelp(cmdButton(Index).Caption)
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdCheck_Click()
'---------------------------------------------------------------------

    Call PerformChecks
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdFunction_Click(Index As Integer)
'---------------------------------------------------------------------
' NCJ 19 Sept 06 - Only for read-write mode
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If mbCanEdit Then
        Call CondAddTextToExpression(cmdFunction(Index).Caption, cmdFunction(Index).Tag)
    End If
    cmdFunction(Index).Value = False
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdButton_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub cmdFunction_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------

    Call DisplayHelp(cmdFunction(Index).Caption)

End Sub

'---------------------------------------------------------------------
Private Sub FillEFormsCombo()
'---------------------------------------------------------------------
' Populate the "Candidates" combo with eForms (for MACRO)
'---------------------------------------------------------------------
Dim clmtask As Task

    cboCandidates.Clear
    'populate candidates combo with all the Macro created Form tasks within the current protocol
    For Each clmtask In gclmGuideline.colTasks
        'check for plans that are Macro created (i.e. locked) with 'Form' at the start of their caption
        If clmtask.TaskType = "plan" Then
            If clmtask.Locked Then
                If Mid(clmtask.Caption, 1, 4) = "Form" Then
                    cboCandidates.AddItem clmtask.Name
                    ' NCJ 31 Oct 02 - Need to store TaskKey too
                    cboCandidates.ItemData(cboCandidates.NewIndex) = clmtask.TaskKey
                End If
            End If
        End If
    Next

End Sub

'---------------------------------------------------------------------
Private Sub FillDataItemsCombo()
'---------------------------------------------------------------------
' Fill Data Items and Values combo with all the AREZZO data items
'---------------------------------------------------------------------
Dim oDataItem As DataItem
Dim vDataItemValue As Variant
Dim colDataItemValues As Collection

    cboDataItems.Clear      ' NCJ 23 Apr 03
    'populate data items combo
    For Each oDataItem In gclmGuideline.colDataItems
        cboDataItems.AddItem oDataItem.Name
        cboDataItems.ItemData(cboDataItems.NewIndex) = oDataItem.DataItemKey
    Next
    
    cboDataValues.Clear
    'populate data values combo
    Set colDataItemValues = gclmGuideline.DataItemValues
    For Each vDataItemValue In colDataItemValues
        cboDataValues.AddItem vDataItemValue
    Next

End Sub

'---------------------------------------------------------------------
Private Sub cmdMatchBrackets_Click()
'---------------------------------------------------------------------
' Match pairs of matching brackets
'---------------------------------------------------------------------

    txtExpression.SetFocus
    Call MatchBrackets
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------------------
' NCJ 23 Jan 03 - Initialise the focus to the txtExpression field for MACRO
'---------------------------------------------------------------------

    txtExpression.SetFocus
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
'Mo Morris 5/12/00  Population of cboOtherTasks is now governed by conditional compilation.
'In MACRO and MTMCUI.exe (Not StandAlone) it will be populated with all of the non-Visit and
'non-Form tasks that have been created in Arezzo, The Top-Level plan and the Data Entry tasks
'will not be displayed.
'In CUI.exe (StandAlone) it will be populated with all of the non-Decision tasks.
'---------------------------------------------------------------------
Dim clmDataItem As DataItem
Dim clmtask As Task
Dim vntCandidate As Variant
Dim sTopLevelPlanKey As String
Dim clmTopLevelPlanTask As Task
Dim sComponentKey As Variant
Dim sFontSize As String

    On Error GoTo ErrHandler

    ' NCJ 19 Jan 06 - Only set font details if there are some
    If GetMACROSetting("FunctionEditorFontName", "") > "" Then
        txtExpression.FontName = GetMACROSetting("FunctionEditorFontName", txtExpression.FontName)
        ' NCJ 20 Sept 07 - Bug 2947 - Ensure numbers are interpreted according to local settings
        ' Settings file always stores in standard format
        sFontSize = LocalNumToStandard(CStr(txtExpression.FontSize))
        txtExpression.FontSize = StandardNumToLocal(GetMACROSetting("FunctionEditorFontSize", sFontSize))
        txtExpression.FontItalic = GetMACROSetting("FunctionEditorFontItalic", "False")
        txtExpression.FontBold = GetMACROSetting("FunctionEditorFontBold", "False")
    End If
    
    lblDecisionsOrVisits.Caption = "Visits"
    lblCandidatesOrForms.Caption = "Forms"
    'populate Decisions combo with all the Macro created Visit tasks within the current protocol
    'Note that these will all be components of the TopLevelPlan
    sTopLevelPlanKey = gclmGuideline.TopLevelPlanKey
    Set clmTopLevelPlanTask = gclmGuideline.colTasks.Item(sTopLevelPlanKey)
    For Each sComponentKey In clmTopLevelPlanTask.Components
        Set clmtask = gclmGuideline.colTasks.Item(sComponentKey)
        'check for plans that are Macro created (i.e. locked) with 'Visit' at the start of their caption
        If clmtask.TaskType = "plan" Then
            If clmtask.Locked Then
                If Mid(clmtask.Caption, 1, 5) = "Visit" Then
                    cboDecisions.AddItem clmtask.Name
                    cboDecisions.ItemData(cboDecisions.NewIndex) = clmtask.TaskKey
                End If
            End If
        End If
    Next
    
    Call FillEFormsCombo
    
    'populate Other Tasks combo
    For Each clmtask In gclmGuideline.colTasks
    'Add all the non-Macro created (not Locked) tasks to the Other combo
    'Note that the Top-level plan and the 'Data Entry' tasks will not appear anywhere
        If Not clmtask.Locked Then
            cboOtherTasks.AddItem clmtask.Name
            cboOtherTasks.ItemData(cboOtherTasks.NewIndex) = clmtask.TaskKey
        End If
    Next
    
    Call ConstructFunctionTabs("MACRO")   'RCW 29/7/02 Added for Expression Editor revamp
    Call SetUpMACROSpecificButtons("MACRO")
        
    ' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True
    
    Call FillDataItemsCombo
    
    'populate time units combo
    cboTimeUnits.AddItem "year"
    cboTimeUnits.AddItem "years"
    cboTimeUnits.AddItem "month"
    cboTimeUnits.AddItem "months"
    cboTimeUnits.AddItem "week"
    cboTimeUnits.AddItem "weeks"
    cboTimeUnits.AddItem "day"
    cboTimeUnits.AddItem "days"
    cboTimeUnits.AddItem "hour"
    cboTimeUnits.AddItem "hours"
    cboTimeUnits.AddItem "minute"
    cboTimeUnits.AddItem "minutes"
    cboTimeUnits.AddItem "second"
    cboTimeUnits.AddItem "seconds"
    
    'display initial help text - moved to tab help RCW 9/8/02
'    txtHelp.Text = "Right mouse click a button - for help text" & vbNewLine _
'        & " Left mouse click a button - for entry and help text"
    
    'initialise double click mechanism
    Call SetupDClickString
    
    'enable matchbracket button if there is text present
    Call EnableMatchBracket
    
    'centre form in screen
    Call CentreForm
    
    msLastComboItem = ""
    mbComboKeyPressed = False
    
    Me.Refresh
    txtExpression.SelStart = Len(txtExpression.Text)
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Form_Load")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
'Validate the expression\condition against the relevant CLM.DLL call
'If valid, the expression will be accepted by CLM.DLL, placed in the
'originating control and the function editor will close down.
'If invalid, the user is informed and the function editor kept open.
'---------------------------------------------------------------------
Dim sReturn As String
Dim bOK As Boolean
Dim sErrMsg As String

    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass
    
    ' NCJ 19/2/01 Initialise to TRUE to prevent empty strings being rejected
    bOK = True
    
    'changed by Mo Morris 16/9/99
    'If an expression/condition exists then validate it before leaving the
    'function editor via a call to CLM.DLL. Knowledge of the type of validation
    '(expression or condition) is held in the calling parameter ValidationType
    If txtExpression.Text <> vbNullString Then
    
        'check that expression is not too large
        bOK = CheckSizeOfExpression
    
        'RCW 5/9/02 - previous contents of IF..THEN moved to
        'CheckExpressionSyntax so that it can be shared
        If bOK Then
            bOK = CheckExpressionSyntax(sErrMsg)
        Else
            txtExpression.SelStart = 1
            DialogWarning msSIZE_ERROR
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
    End If
    
    If bOK Then
        'place the contents of txtExpression into sReturn prior to unloading
        'the function editor form
        sReturn = txtExpression.Text
        Screen.MousePointer = vbDefault
        Me.Hide
        ' Ignore errors occurring in source control - NCJ 2 Dec 99
        On Error Resume Next
        'set the focus back to the calling/originating control
        mctlSourceOfExpression.SetFocus
        'and then place the possibly changed expression back it
        mctlSourceOfExpression.Text = sReturn
        On Error GoTo 0
        Unload Me
    Else        ' Expression is not OK
        txtExpression.SelStart = Len(txtExpression.Text)
        Call DialogWarning(sErrMsg)
        Screen.MousePointer = vbDefault
    End If
    
Exit Sub

ErrHandler:
    DialogWarning "This is not a valid AREZZO term" & vbCrLf & Err.Description
    txtExpression.SelStart = 0
    txtExpression.SelLength = 0
    Screen.MousePointer = vbDefault
        
End Sub


'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------

    On Error Resume Next    ' Necessary for read-only mode where mctlSourceOfExpression might not be enabled
    
    Me.Hide
    'set the focus back to the calling/originating control
    mctlSourceOfExpression.SetFocus
    Unload Me
   
End Sub

'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------

    Call ResizeFormElements

End Sub

'---------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------------------
' NCJ 26/9/00 - Set mctlSourceOfExpression to Nothing to release reference
'---------------------------------------------------------------------

    ' NCJ 26/9/00
    Set mctlSourceOfExpression = Nothing
    
    'save form positional details in registry
    mlngRegThisFormLeft = Me.Left
    SetMACROSetting "FunctionEditorLeft", mlngRegThisFormLeft
    mlngRegThisFormTop = Me.Top
    SetMACROSetting "FunctionEditorTop", mlngRegThisFormTop

End Sub

'---------------------------------------------------------------------
Private Sub DisplayHelp(ByRef sButtonTitle As String)
'---------------------------------------------------------------------
'Changed by Mo Morris 22/10/99
'Helptext for datetime and datediff commented out.
'Helptext added for ':', abs, between, case, datenow, date_diff, else,
'time_diff, timenow and status
' NCJ 19 Feb 01 - Changed help text for "like"
'
' RCW 8/8/02 - Changed to use Prolog for help text. Part of new look
'              for expression editor. Accepts button caption from
'              buttons inside and outside tabbed area instead of
'              button index.
'-----------------------------------------------------------

Dim sQuery As String
Dim sResultCode As String
Dim sHelpText As String

    On Error GoTo ErrHandler

    sQuery = "cmp_get_help_text( button, `" & sButtonTitle & "` ). "
    sHelpText = goALM.GetPrologResult(sQuery, sResultCode)
    
    If sHelpText <> " " Then
        txtHelp.Text = sHelpText
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "DisplayHelp")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    

End Sub

'---------------------------------------------------------------------
Private Sub ConstructFunctionTabs(ByRef sApplication As String)
'---------------------------------------------------------------------
'
' RCW 29/7/02
' Created as part of the overhaul of the expression editor.
' ConstructFunctionTabs builds the function tab area
'
'---------------------------------------------------------------------

Dim nTabCounter As Integer
Dim colTabs As Collection
Dim nTabIndex As Integer
Dim varTabIndex As Variant
Dim sTabDetails As String
Dim sTabDetailArr As Variant
Dim sCaption As String
Dim sToolTipText As String
Dim sHelpText As String
Dim sFrameList As String
Dim sApplicationFlag As String
Dim sTabID As String


    On Error GoTo Error
    
    'Get details for all tabs from Prolog
    'set up mcolTabs
    Call GetAllTabDetails
    
    nTabCounter = 1 'keeps count of tabs created. Remember that as some
                    'tabs are not always visible nTabCounter may not
                    'be the same as nTabIndex
                    
    'record width of first frame before doing anything
    msglFrameWidth = fraButtonContainer(0).Width
    
    'cycle through tabs in collection
    For nTabIndex = 1 To (mcolTabs.Count - 1)
        'get details of each tab
        Call UnpackTab(nTabIndex, sTabID, sApplicationFlag, sCaption, sToolTipText, sFrameList)
        
        'create tab and set details
        Call AddTab(nTabCounter, nTabIndex, sCaption, sToolTipText, sApplicationFlag, sApplication, sFrameList, sTabID)
    
    Next

    'populate first tab
    Call SelectTab(1)

    Exit Sub
Error:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Initialise")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select


End Sub


'---------------------------------------------------------------------
Private Sub SetTabDetails(ByRef nTabNumber As Integer, _
                          ByVal sCaption As String, _
                          ByVal sToolTipText As String, _
                          ByVal sFrameList As String, _
                          ByVal sTabID As Integer)
'---------------------------------------------------------------------
'
' RCW 29/7/02
' Created as part of the overhaul of the expression editor.
' SetTabDetails initialises a specific tab
'
'---------------------------------------------------------------------
   
    On Error GoTo Err
    
    With frmCUIFunctionEditor.tabFunctionSelector.Tabs(nTabNumber)
        .Caption = sCaption
        .Key = sCaption
        .ToolTipText = sToolTipText
        'use tag as convenient place to store tab index for Prolog as
        'well as details of tab contents
        .Tag = sTabID & "|" & sFrameList
    End With
    
    Exit Sub

Err:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Initialise")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub AddTab(ByRef nTabNumber As Integer, _
                   ByVal nTabIndex As Integer, _
                   ByVal sCaption As String, _
                   ByVal sToolTipText As String, _
                   ByVal sApplicationFlag As String, _
                   ByVal sApplication As String, _
                   ByVal sFrameList As String, _
                   ByVal sTabID As Integer)
'---------------------------------------------------------------------
'
' RCW 29/7/02
' Created as part of the overhaul of the expression editor.
' AddTab adds a specific tab to the tab area
'
'---------------------------------------------------------------------

    'set tab details if applicable for the application
    If ((sApplication = "MACRO" _
         And (sApplicationFlag = "M" Or sApplicationFlag = "B")) _
        Or ((sApplication = "AREZZO" _
             And (sApplicationFlag = "A" Or sApplicationFlag = "B")))) Then
    
        'One tab is already present so the first tab should not be added
        If nTabIndex <> 1 Then
            frmCUIFunctionEditor.tabFunctionSelector.Tabs.Add nTabNumber
        End If
    
        Call SetTabDetails(nTabNumber, sCaption, sToolTipText, sFrameList, sTabID)
        
        'Increment tab number to show actual number of displayed tabs
        nTabNumber = nTabNumber + 1
    
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub SelectTab(ByVal nTabNumber As Integer)
'---------------------------------------------------------------------
' RCW 29/7/02
' Created as part of the overhaul of the expression editor.
' SelectTab populates a tab with appropriate buttons and frames
'---------------------------------------------------------------------
Dim sFrameList As String
Dim vFrameList As Variant


    'get frame list
    Call UnpackGeneric(tabFunctionSelector.Tabs(nTabNumber).Tag, "|", vFrameList)
    sFrameList = vFrameList(1)
    
    Call UnpackAndPlaceFrames(nTabNumber, sFrameList)
    
    Call DisplayTabHelp(nTabNumber)

    Call SetTabOrder

End Sub

'---------------------------------------------------------------------
Private Sub SetUpFrame(nTabNumber As Integer, nFrameNumber As Integer, sFrame As String)
'---------------------------------------------------------------------
'
' RCW 31/7/02
' Created as part of the overhaul of the expression editor.
' SetUpFrame sets attributes of frame and puts buttons in it
'
'---------------------------------------------------------------------

Dim sFrameCaption As String
Dim sButtons As String  'list of button definitions
    
    'get frame details
    Call UnpackFrame(sFrame, sFrameCaption, sButtons)
             
    'assign frame details
    frmCUIFunctionEditor.fraButtonContainer(nFrameNumber).Caption = sFrameCaption
    frmCUIFunctionEditor.fraButtonContainer(nFrameNumber).Width = msglFrameWidth
        
    'put buttons in frame
    Call PutButtonsInFrame(nTabNumber, sButtons, nFrameNumber)
    
End Sub

'---------------------------------------------------------------------
Private Sub PutButtonsInFrame(nTabNumber As Integer, _
                              sButtons As String, _
                              nFrameNumber As Integer)
'---------------------------------------------------------------------
'
' RCW 31/7/02
' Created as part of the overhaul of the expression editor.
' PutButtonsInFrame populates a frame with appropriate buttons
'
'---------------------------------------------------------------------
Dim sglOriginLeft As Single
Dim sglOriginTop As Single
Dim sglButtonTop As Single
Dim sglButtonLeft As Single
Dim sglButtonRight As Single
Dim sglFrameHeight As Single
Dim sglFrameWidth As Single
Dim nRowCount As Integer

Dim nButtonID As Integer
Dim sCaption As String
Dim sTag As String
Dim sToolTipText As String
Dim vButton As Variant
Dim nButton As Integer
Dim sButton As String

    On Error GoTo Error
    
    'Set origin coordinates
    sglOriginLeft = cmdFunction(0).Left
    sglOriginTop = cmdFunction(0).Top
    nRowCount = 1
    sglFrameWidth = 0
    
    'Set button coords
    sglButtonLeft = sglOriginLeft + mnSPACER
    sglButtonTop = sglOriginTop + mnSPACER
    
    'get button details for this frame
    Call UnpackButtonList(sButtons, vButton)
    
    For nButton = 0 To UBound(vButton) - 2
        sButton = CStr(vButton(nButton))
        
        Call CreateButton(nFrameNumber, _
                          sglButtonTop, _
                          sglButtonLeft, _
                          sButton, _
                          nButtonID)
                
        'adjust button co-ords
        sglButtonRight = sglButtonLeft + _
                       CInt(frmCUIFunctionEditor.cmdFunction(nButtonID).Width) + _
                       mnSPACER
                       
        If sglButtonRight > msglFrameWidth Then
            'Button won't fit on this line
            sglButtonTop = sglButtonTop + CInt(frmCUIFunctionEditor.cmdFunction(0).Height) + mnSPACER
            sglButtonLeft = sglOriginLeft + mnSPACER
            nRowCount = nRowCount + 1
            With cmdFunction(nButtonID)
                .Top = sglButtonTop
                .Left = sglButtonLeft
            End With
        Else
            'collect information for resizing frame
            If sglButtonRight > sglFrameWidth Then
                sglFrameWidth = sglButtonRight
            End If
        End If
        'set position for next button
        sglButtonLeft = sglButtonLeft + mnSPACER + CInt(cmdFunction(nButtonID).Width)
    Next
        
    'resize frame when finished
    Call ResizeFrame(nFrameNumber, nRowCount, sglFrameWidth)
     
    Exit Sub
Error:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Initialise")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub


'---------------------------------------------------------------------
Private Sub ResizeFrame(ByRef nFrameNumber As Integer, _
                        ByRef nRowCount As Integer, _
                        ByRef sglFrameWidth As Single)
'---------------------------------------------------------------------
'
' RCW 5/8/02
' Resize frame after buttons have been placed in it
'
'---------------------------------------------------------------------

Dim sglFrameHeight As Single
Dim sglCaptionWidth As Single
Dim sglContentsWidth As Single

    sglFrameHeight = (mnBUTTONHEIGHT + mnSPACER) * nRowCount + mnFRAMEPADDING
    fraButtonContainer(nFrameNumber).Height = sglFrameHeight
    sglContentsWidth = sglFrameWidth + msglFRAMEWIDTHPADDING
    sglCaptionWidth = CInt(Me.TextWidth(fraButtonContainer(nFrameNumber).Caption)) _
                    + (2 * msglFRAMEWIDTHPADDING)
    If sglContentsWidth > sglCaptionWidth Then
        fraButtonContainer(nFrameNumber).Width = sglContentsWidth
    Else
        fraButtonContainer(nFrameNumber).Width = sglCaptionWidth
    End If

End Sub

'---------------------------------------------------------------------
Private Sub PositionFrames(nTabNumber As Integer)
'---------------------------------------------------------------------
'
' RCW 31/7/02
' Created as part of the overhaul of the expression editor.
' PositionFrames arranges frames logically within a tab
'
'---------------------------------------------------------------------
Dim nFrameNumber As Integer
Dim nLastFrame As Integer
Dim nButton As Integer

Const sglFRAMEPADDING = 120

    'make sure default frame is visible
    fraButtonContainer(0).Visible = True
    
    'cycle through rest of frames
    If fraButtonContainer.Count > 1 Then
        For nFrameNumber = 1 To fraButtonContainer.Count - 1
            nLastFrame = nFrameNumber - 1
            
            'try to position frame to the right of the last frame
            fraButtonContainer(nFrameNumber).Top = _
                fraButtonContainer(nLastFrame).Top
            fraButtonContainer(nFrameNumber).Left = _
                fraButtonContainer(nLastFrame).Left + _
                fraButtonContainer(nLastFrame).Width + _
                (2 * mnSPACER)
            
            If (fraButtonContainer(nFrameNumber).Left _
                + fraButtonContainer(nFrameNumber).Width _
                + sglFRAMEPADDING) _
                > tabFunctionSelector.Width Then
                'position current frame underneath last frame
                fraButtonContainer(nFrameNumber).Top = _
                    CInt(fraButtonContainer(nLastFrame).Top) + _
                    CInt(fraButtonContainer(nLastFrame).Height) + _
                    mnVERTICALSPACER
                fraButtonContainer(nFrameNumber).Left = frmCUIFunctionEditor.fraButtonContainer(0).Left
            End If
            
           'bring to top and make visible
            fraButtonContainer(nFrameNumber).ZOrder 0
            fraButtonContainer(nFrameNumber).Visible = True
        Next
    End If
   
End Sub

'---------------------------------------------------------------------
Private Sub mnuCheckOpt_Click(Index As Integer)
'---------------------------------------------------------------------
    
    Select Case Index
    
        Case 0 'Check
            Call PerformChecks
            
    End Select
        
End Sub

'---------------------------------------------------------------------
Private Sub mnuEdit_Click()
'---------------------------------------------------------------------

    Call SetEditMenuOpts

End Sub

'---------------------------------------------------------------------
Private Sub mnuEditOpt_Click(Index As Integer)
'---------------------------------------------------------------------

    Select Case Index
    
        Case 0 'cut
            Call Cut
        Case 1 'copy
            Call Copy
        Case 2 'paste
            Call Paste
           
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuOptionsFont_Click(Index As Integer)
'---------------------------------------------------------------------
' NCJ 3 Jan 06 - Allow a different font! Hooray!
'---------------------------------------------------------------------
    
    On Error Resume Next
    
    With txtExpression
        CommonDialog1.FontName = .FontName
        CommonDialog1.FontBold = .FontBold
        CommonDialog1.FontItalic = .FontItalic
        CommonDialog1.FontSize = .FontSize
    End With
    
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlCFBoth Or cdlCFScalableOnly Or cdlCFWYSIWYG
    CommonDialog1.ShowFont
    
    'Changed by Mo Morris   29/4/99
    'Because the fonts dialog is restricted to fonts that are available on the printer as well
    'as the computer, if no printers are installed then the user will have been informed by the
    'unhelpful message 'THERE ARE NO FONTS INSTALLED OPEN THE FONTS FOLDER etc'
    'The following code traps the generated error code and displays a more helpful message
    If Err.Number = 24574 Then
        MsgBox ("You will not be able to edit fonts until you have installed a default printer.")
        Exit Sub
    End If
    
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        On Error GoTo 0
    End If
    
    With txtExpression
        ' NCJ 22 Feb 06 - Avoid crash if Font name is empty!! (It can happen...)
        If CommonDialog1.FontName > "" Then .FontName = CommonDialog1.FontName
        .FontSize = CommonDialog1.FontSize
        .FontBold = CommonDialog1.FontBold
        .FontItalic = CommonDialog1.FontItalic
    End With
    
    SetMACROSetting "FunctionEditorFontName", txtExpression.FontName
    ' NCJ 20 Sept 07 - Issue 2947 - Store font size as "standard" (may be fractional)
    SetMACROSetting "FunctionEditorFontSize", LocalNumToStandard(CStr(txtExpression.FontSize))
    SetMACROSetting "FunctionEditorFontItalic", txtExpression.FontItalic
    SetMACROSetting "FunctionEditorFontBold", txtExpression.FontBold

End Sub

'---------------------------------------------------------------------
Private Sub mnuOptionsOpt_Click(Index As Integer)
'---------------------------------------------------------------------

    'Autocomplete on/off
    Call AutoCompleteToggle(Index)

End Sub

'---------------------------------------------------------------------
Private Sub picBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
'
' RCW 9/8/02
' Created as part of the overhaul of the expression editor.
' Indicate that clicking on the border between the edit and help text
' box windows will allow the resizing of the two windows.
'
'---------------------------------------------------------------------

    frmCUIFunctionEditor.MousePointer = vbSizeNS
    mbMouseDrag = True

End Sub


'---------------------------------------------------------------------
Private Sub picBorder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
'
' RCW 9/8/02
' Created as part of the overhaul of the expression editor.
' Move the picture box = resize of text boxes.
'
'---------------------------------------------------------------------

    If mbMouseDrag Then
        Call MoveBorder(Y)
    End If

End Sub

'---------------------------------------------------------------------
Private Sub picBorder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
'
' RCW 9/8/02
' Created as part of the overhaul of the expression editor.
' Indicate that clicking on the border between the edit and help text
' box windows will allow the resizing of the two windows.
'
'---------------------------------------------------------------------

    frmCUIFunctionEditor.MousePointer = vbDefault
    mbMouseDrag = False

End Sub

'---------------------------------------------------------------------
Private Sub tabFunctionSelector_Click()
'---------------------------------------------------------------------
Dim nTabSelected As Integer

    nTabSelected = CInt(tabFunctionSelector.SelectedItem.Index)
    
    Call ClearTabs
    Call SelectTab(nTabSelected)

End Sub

'---------------------------------------------------------------------
Private Sub GetAllTabDetails()
'---------------------------------------------------------------------
'
' RCW 2/8/02
' Created as part of the overhaul of the expression editor.
' Gets the details of all tabs in one collection
'
'---------------------------------------------------------------------
Dim sQuery As String
Dim sResultCode As String
Dim vTab As Variant
Dim sTab As String
Dim vTabAttrs As Variant
Dim sText As String
Dim sCaption As String
Dim nTab As Integer

    On Error GoTo Error
    
    sQuery = "cmp_get_tabs. "
    Set mcolTabs = goALM.GetPrologList(sQuery, sResultCode)
    
    Exit Sub
Error:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Initialise")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub UnpackGeneric(sSource As String, sDelimiter As String, vParsedString As Variant)
'---------------------------------------------------------------------
'
' RCW 5/8/02
' Created as part of the overhaul of the expression editor.
' Given a specified delimiter, unpacks elements from a string
'
'---------------------------------------------------------------------

    vParsedString = Split(sSource, sDelimiter)

End Sub

'---------------------------------------------------------------------
Private Sub UnpackTab(ByRef nTabIndex As Integer, _
                      ByRef sTabID As String, _
                      ByRef sApplicationFlag As String, _
                      ByRef sCaption As String, _
                      ByRef sToolTipText As String, _
                      ByRef sFrameList As String)
'---------------------------------------------------------------------
'
' RCW 5/8/02
' Created as part of the overhaul of the expression editor.
' Unpacks the details for a tab from a collection element to produce
'   Caption
'   Tool Tip Text
'   Application Flag (M for MACRO only, A for AREZZO only and B for both
'   List of frames
'
'---------------------------------------------------------------------
Dim vTabContent As Variant

    Call UnpackGeneric(CStr(mcolTabs(nTabIndex)), Chr(164), vTabContent)
    
    sTabID = CStr(vTabContent(0))
    sApplicationFlag = CStr(vTabContent(1)) ' MACRO only or AREZZO only or Both
    sCaption = CStr(vTabContent(2))
    sToolTipText = CStr(vTabContent(3))
    sFrameList = CStr(vTabContent(4))
                      
End Sub

'---------------------------------------------------------------------
Private Sub UnpackAndPlaceFrames(ByRef nTabNumber As Integer, _
                                 ByVal sFrameList As String)
'---------------------------------------------------------------------
'
' RCW 5/8/02
' Created as part of the overhaul of the expression editor.
' Unpacks the details for a frame and arranges buttons
'
'---------------------------------------------------------------------
Dim vFrame As Variant   'array returned by UnpackGeneric
Dim nFrame As Integer   'Frame counter
Dim sFrame As String    'Frame definition containing packed buttons and frame caption

    On Error GoTo Error

    Call UnpackGeneric(sFrameList, Chr(165), vFrame)
    
    'First frame is already present
    sFrame = CStr(vFrame(0))
    
    'set up tab index counter for new buttons
    mnTabIndex = 7
    Call SetUpFrame(nTabNumber, 0, sFrame)
    

    If UBound(vFrame) > 1 Then
        For nFrame = 1 To UBound(vFrame) - 2
            Load fraButtonContainer(nFrame)
            Set fraButtonContainer(nFrame).Container = fraControlsContainer
                        
            sFrame = CStr(vFrame(nFrame))
            Call SetUpFrame(nTabNumber, nFrame, sFrame)
        Next
    End If

    Call PositionFrames(nTabNumber)

    Exit Sub
    
Error:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Initialise")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub UnpackFrame(sFrame As String, sCaption As String, sButtons As String)
'---------------------------------------------------------------------
'
' RCW 5/8/02
' Created as part of the overhaul of the expression editor.
' Extracts frame attributes from string
'
'---------------------------------------------------------------------

Dim vFrameAttr As Variant

    Call UnpackGeneric(sFrame, Chr(166), vFrameAttr)
    
    sCaption = CStr(vFrameAttr(1))
    sButtons = CStr(vFrameAttr(2))

End Sub

'---------------------------------------------------------------------
Private Sub UnpackButtonList(sButtons As String, vButton As Variant)
'---------------------------------------------------------------------
'
' RCW 5/8/02
' Created as part of the overhaul of the expression editor.
' Extracts the list of buttons from string
'
'---------------------------------------------------------------------

    Call UnpackGeneric(sButtons, Chr(167), vButton)
    
End Sub

'---------------------------------------------------------------------
Private Sub UnpackButton(ByVal sButton As String, _
                         ByRef nButtonID As Integer, _
                         ByRef sCaption As String, _
                         ByRef sTag As String, _
                         ByRef sToolTipText As String)
'---------------------------------------------------------------------
'
' RCW 5/8/02
' Created as part of the overhaul of the expression editor.
' Gets button details from string
'
'---------------------------------------------------------------------
Dim vButtonAttr As Variant

    Call UnpackGeneric(sButton, Chr(168), vButtonAttr)
    
    nButtonID = CInt(vButtonAttr(0))
    sCaption = CStr(vButtonAttr(1))
    sTag = CStr(vButtonAttr(2))
    sToolTipText = CStr(vButtonAttr(3))

End Sub

'---------------------------------------------------------------------
Private Sub ClearTabs()
'---------------------------------------------------------------------
'
' RCW 5/8/02
' Created as part of the overhaul of the expression editor.
' Gets button details from string
'
'---------------------------------------------------------------------
Dim nButton As Integer
Dim nFrame As Integer
Dim cmdFnButton As CommandButton
Dim bFirst As Boolean

    'clear previously selected tab contents
    'delete buttons
    bFirst = True
    For Each cmdFnButton In cmdFunction
        If bFirst Then
            bFirst = False 'first button in collection is not dynamic
                           'and can't be deleted
        Else
            Unload cmdFnButton
        End If
    Next
    
    'delete frames
    If CInt(fraButtonContainer.Count) >= 2 Then
        For nFrame = 2 To CInt(fraButtonContainer.Count)
            Unload fraButtonContainer(nFrame - 1)
        Next
    End If

    'restore size of first frame
    fraButtonContainer(0).Width = msglFrameWidth
    
End Sub

'---------------------------------------------------------------------
Private Sub CreateButton(ByRef nFrameNumber As Integer, _
                         ByRef sglButtonTop As Single, _
                         ByRef sglButtonLeft As Single, _
                         ByRef sButton As String, _
                         ByRef nButtonID As Integer)
'---------------------------------------------------------------------
'
' RCW 5/8/02
' Created as part of the overhaul of the expression editor.
' Gets button details from string
'
'---------------------------------------------------------------------
Dim sCaption As String
Dim sTag As String
Dim sToolTipText As String

        'extract button details from string
        Call UnpackButton(sButton, nButtonID, sCaption, sTag, sToolTipText)
        
        'create new button
        Load cmdFunction(nButtonID)
        
        'adjust its properties
        With cmdFunction(nButtonID)
            'attach button to frame
            Set .Container = fraButtonContainer(nFrameNumber)
            'set caption, tooltiptext and tag
            .Caption = sCaption
            .ToolTipText = sToolTipText
            .Tag = sTag
            'physical position
            .Top = sglButtonTop
            .Left = sglButtonLeft
            'size
            .Width = Me.TextWidth(sCaption) + 200
            'make visible
            .Visible = True
            'bring to top
            .ZOrder 0
            .TabIndex = mnTabIndex
        End With
        
        mnTabIndex = mnTabIndex + 1

End Sub

'---------------------------------------------------------------------
Private Sub DisplayTabHelp(ByRef nTabNumber As Integer)
'---------------------------------------------------------------------
'
' RCW 8/8/02
' Created as part of the overhaul of the expression editor.
' Gets help text for a tab and displays it in the help window.
' Also displays left and right mouse button help message.
'
'---------------------------------------------------------------------

Dim nTabPrologIndex As Integer
Dim sQuery As String
Dim sResultCode As String
Dim sHelpText As String
Dim vTabTag As Variant

    Call UnpackGeneric(tabFunctionSelector.Tabs(nTabNumber).Tag, "|", vTabTag)
    nTabPrologIndex = vTabTag(0)
    
    sQuery = "cmp_get_help_text( tab, " & nTabPrologIndex & " ). "
    sHelpText = goALM.GetPrologResult(sQuery, sResultCode)
    sHelpText = sHelpText & vbNewLine & vbNewLine & _
                            "Right mouse click a button - for help text" & _
                            vbNewLine & _
                            "Left mouse click a button - for entry and help text"
    
    txtHelp.Text = sHelpText
    
End Sub

'---------------------------------------------------------------------
Private Sub ResizeFormElements()
'---------------------------------------------------------------------
'
' RCW 8/8/02
' Created as part of the overhaul of the expression editor.
' Resizes edit, help and fields
'
'---------------------------------------------------------------------
    
    If (Me.WindowState <> vbMaximized) And _
       (Me.WindowState <> vbMinimized) Then
        Call FormResizeVertical
        Call FormResizeHorizontal
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub MoveBorder(ByRef sglUpOrDownShift As Single)
'---------------------------------------------------------------------
'
' RCW 9/8/02
' Created as part of the overhaul of the expression editor.
' Indicate that clicking on the border between the edit and help text
' box windows will allow the resizing of the two windows
'
'---------------------------------------------------------------------
Dim sglCombinedHeight As Single
Dim sglEditHeight As Single
Dim sglHelpHeight As Single
Dim bUp As Boolean

Const sglMINIMUMWINDOWHEIGHT = 500

    'determine direction
    bUp = (sglUpOrDownShift <= 0)
    
    'find combined height of edit window, help window and border
    sglEditHeight = txtExpression.Height
    sglHelpHeight = txtHelp.Height
    sglCombinedHeight = sglEditHeight + _
                        sglHelpHeight + _
                        picBorder.Height
        
    If (sglEditHeight > sglMINIMUMWINDOWHEIGHT And bUp) Or _
       (sglHelpHeight > sglMINIMUMWINDOWHEIGHT And Not (bUp)) Then
    
        'adjust height of text window by shift
        txtExpression.Height = sglEditHeight + sglUpOrDownShift
        
        'calculate new border top
        picBorder.Top = picBorder.Top + sglUpOrDownShift
        
        'calculate new top of help window
        txtHelp.Top = txtHelp.Top + sglUpOrDownShift
        
        'and its height
        txtHelp.Height = sglCombinedHeight _
                         - txtExpression.Height _
                         - picBorder.Height
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub FormResizeHorizontal()
'---------------------------------------------------------------------
'
' RCW 9/8/02
' Created as part of the overhaul of the expression editor.
' Resizes edit, help and fields
'
'---------------------------------------------------------------------
Dim sglWindowWidth As Single
Dim sglComboBoxWidth As Single
Dim sglFormWidth As Single
Dim sglContainerFrameWidth As Single
Const sglRESIZEWIDTHLIMIT = 8400    'limit of how small form can get

    sglFormWidth = frmCUIFunctionEditor.Width
    If sglFormWidth >= sglRESIZEWIDTHLIMIT Then
        'resize edit and help windows laterally
        sglWindowWidth = frmCUIFunctionEditor.Width - 390
        txtExpression.Width = sglWindowWidth
        txtHelp.Width = sglWindowWidth
        picBorder.Width = sglWindowWidth
        
        'resize controls container frame
        sglContainerFrameWidth = frmCUIFunctionEditor.Width - 345
        fraControlsContainer.Width = sglContainerFrameWidth
        
        'resize combobox fields laterally
        sglComboBoxWidth = frmCUIFunctionEditor.Width - 7280
        cboCandidates.Width = sglComboBoxWidth
        cboDataItems.Width = sglComboBoxWidth
        cboDataValues.Width = sglComboBoxWidth
        cboDecisions.Width = sglComboBoxWidth
        cboOtherTasks.Width = sglComboBoxWidth
        cboTimeUnits.Width = sglComboBoxWidth
        
        'keep check button aligned with right edge of combobox fields
        cmdCheck.Left = cboCandidates.Left + _
                        cboCandidates.Width - _
                        cmdCheck.Width
                        
    Else
        'need to stop resizing and keep size at minimum
        'this keeps behavior consistent with main Composer form
        frmCUIFunctionEditor.Width = sglRESIZEWIDTHLIMIT
    End If

End Sub

'---------------------------------------------------------------------
Private Sub FormResizeVertical()
'---------------------------------------------------------------------
'
' RCW 9/8/02
' Created as part of the overhaul of the expression editor.
' Resizes edit and help windows
'
'---------------------------------------------------------------------

Dim sglFormHeight As Single

Const sglRESIZEHEIGHTLIMIT = 7000
    
    sglFormHeight = frmCUIFunctionEditor.Height
    
    If sglFormHeight > sglRESIZEHEIGHTLIMIT Then
        'resize components of form dependent on current form height
        Call ResizeFormDependent(sglFormHeight)
    Else
        'need to stop resizing and keep size at minimum
        'this keeps behavior consistent with main Composer form
        frmCUIFunctionEditor.Height = sglRESIZEHEIGHTLIMIT
        
        'resize components of form dependent on minimum form height
        Call ResizeFormDependent(sglRESIZEHEIGHTLIMIT)
    End If
End Sub

'---------------------------------------------------------------------
Private Function DoubleAllowed(ByRef sText As String) As Boolean
'---------------------------------------------------------------------
'
' RCW 9/8/02
' Created as part of the overhaul of the expression editor.
' Returns true if doubleclicks are allowed on a particular button.
' Doubleclicks are allowed for functions held in msDClickString
'
'---------------------------------------------------------------------
Dim bDoubleClick As Boolean
Dim bInPermittedDoubleList As Boolean
Dim nPos As Integer

    bDoubleClick = (sText = msLastButton)
    
    If bDoubleClick Then
        If Len(sText) < Len(msDClickString) Then
            nPos = 0
            nPos = InStr(msDClickString, sText)
            bInPermittedDoubleList = (nPos > 0)
        Else
            bInPermittedDoubleList = False
        End If
        
        DoubleAllowed = bInPermittedDoubleList
        
    Else
    
        DoubleAllowed = True
    
    End If
    
End Function

'---------------------------------------------------------------------
Private Sub SetupDClickString()
'---------------------------------------------------------------------
'
' RCW 9/8/02
' Created as part of the overhaul of the expression editor.
' Setup up msDClickString which holds the functions that are allowed
' to be clicked on more than once.
'
'---------------------------------------------------------------------
Dim sQuery As String
Dim sResultCode As String

    sQuery = "cmp_get_allowed_doubles. "
    msDClickString = goALM.GetPrologResult(sQuery, sResultCode)
    
    msDClickString = msDClickString & "0123456789()"

End Sub

'---------------------------------------------------------------------
Private Sub CondAddTextToExpression(ByRef sCaption As String, _
                                    ByRef sTag As String)
'---------------------------------------------------------------------
'
' RCW 9/8/02
' Created as part of the overhaul of the expression editor.
' Conditional layer to AddTextToExpression
'
'---------------------------------------------------------------------

    If DoubleAllowed(sCaption) Then
        Call AddTextToExpression(sTag)
    End If
    msLastButton = sCaption

End Sub

'---------------------------------------------------------------------
Private Sub SetUpMACROSpecificButtons(ByRef sApplicationName As String)
'---------------------------------------------------------------------
'
' RCW 16/8/02
' Created as part of the overhaul of the expression editor.
' Makes colon button larger for MACRO as well as me:value visible
'
'---------------------------------------------------------------------

    If sApplicationName = "MACRO" Then
        cmdButton(59).Height = 650  'larger colon button
        cmdButton(4).Visible = True 'me:value visible
    Else
        cmdButton(59).Height = 300  'smaller colon button for AREZZO
        cmdButton(4).Visible = False 'no me:value for AREZZO
    End If

End Sub

'---------------------------------------------------------------------
Private Sub ResizeFormDependent(ByVal sglFormHeight As Single)
'---------------------------------------------------------------------
'
' RCW 27/8/02
' Created as part of the overhaul of the expression editor.
' Resizes edit and help windows relative to a specified form height.
' Used both when form being resized and when form has reached lowest height.
'
'---------------------------------------------------------------------
Dim sglCombinedHeight As Single
Dim sglEditHeight As Single
Dim sglHelpHeight As Single
Dim sglHelpTop As Single
Dim sglNewCombinedHeight As Single
Dim sglContainerFrameTop As Single

        'move top of control container
        sglContainerFrameTop = frmCUIFunctionEditor.ScaleHeight _
                               - msglCONTROLBOXHEIGHTFROMBOTTOM
        fraControlsContainer.Top = sglContainerFrameTop
    
        'find combined height of edit window, help window and border
        sglEditHeight = txtExpression.Height
        sglHelpHeight = txtHelp.Height
        sglCombinedHeight = sglEditHeight + _
                            sglHelpHeight + _
                            picBorder.Height
        
        'calculate new combined height from new form height
        sglNewCombinedHeight = sglFormHeight _
                               - txtExpression.Top _
                               - msglDISTHELPBTMFORMFORMBTM
        
        'recalculate help window height as same proportion
        'of combined height
        txtHelp.Height = sglNewCombinedHeight * (sglHelpHeight / _
                                                 sglCombinedHeight)
                                                 
        'calculate edit window height as what's left
        txtExpression.Height = sglNewCombinedHeight - _
                               txtHelp.Height - _
                               picBorder.Height
                               
        'move border to bottom of edit window
        picBorder.Top = txtExpression.Top + _
                        txtExpression.Height
                        
        'move help to bottom of border
        txtHelp.Top = picBorder.Top + _
                      picBorder.Height

End Sub





'---------------------------------------------------------------------
Private Sub CloseBracket(ByRef sType As String, ByVal sBracketType As String)
'---------------------------------------------------------------------
'
' RCW 2/9/02
' Add closing bracket
'
' 25/9/02 - include [] as well
'
'---------------------------------------------------------------------
Dim lngPos As Long
Dim sQuery As String
Dim sPostText As String
Dim sResultCode As String
Dim sClosingBracket As String

    Select Case sBracketType
        Case "("
            sClosingBracket = ")"
        Case "["
            sClosingBracket = "]"
    End Select
        
    lngPos = txtExpression.SelStart
    If sType = "key" Then ' close bracket after adding an extra space
        txtExpression.SelText = "  " & sClosingBracket
        txtExpression.SelStart = lngPos + 1
    Else
        'spaces are already included in function name
        txtExpression.SelText = " " & sClosingBracket
        txtExpression.SelStart = lngPos
    End If

End Sub


'---------------------------------------------------------------------
Private Sub Copy()
'---------------------------------------------------------------------
'
' RCW 4/9/02
' Copy text from expression to clipboard
'
'---------------------------------------------------------------------
    
    Call Clipboard.Clear
    Call Clipboard.SetText(txtExpression.SelText)
    
End Sub

'---------------------------------------------------------------------
Private Sub Cut()
'---------------------------------------------------------------------
'
' RCW 2/9/02
' Cut text from expression to clipboard
'
'---------------------------------------------------------------------
   
    Call Clipboard.Clear
    Call Clipboard.SetText(txtExpression.SelText)
    txtExpression.SelText = ""
    
End Sub

'---------------------------------------------------------------------
Private Sub Paste()
'---------------------------------------------------------------------
' RCW 2/9/02
' Paste text into expression from clipboard
' NCJ 19 Sept 06 - Only do it if in Edit mode
'---------------------------------------------------------------------
    
    If mbCanEdit Then txtExpression.SelText = Clipboard.GetText()

End Sub

'---------------------------------------------------------------------
Private Sub PerformChecks()
'---------------------------------------------------------------------
'
' RCW 5/9/02
' Run syntax and content checks on expression
' (E.A) Nov 2002 Added extra condition and expression check
' NCJ Jun 03 - New term checker
'---------------------------------------------------------------------
Dim sQuery As String
Dim sResultCode As String
Dim sCheckText As String

    Screen.MousePointer = vbHourglass
    
    If CheckSizeOfExpression Then
        ' Check syntax first
'        If CheckExpressionSyntax(sCheckText) Then
            ' If the user is currently editing a condition argument
            If msValidationType = "Condition" Then
                
                'Check if there is a problem with the condition
'                sQuery = "clmc_check_condition( `" & txtExpression.Text & "` ). "
                sQuery = "pdl_check_cond( `" & txtExpression.Text & "` ). "
                sCheckText = goALM.GetPrologResult(sQuery, sResultCode)
            
            ' Else if the user is currenly editing an expression
            ElseIf msValidationType = "Expression" Then
            
                'Check if there is a problem with the expression
'                sQuery = "clmc_check_expression( `" & txtExpression.Text & "`"
                sQuery = "pdl_check_expr( `" & txtExpression.Text & "`"
                ' NCJ 27 Jan 03 - Add term type if given
                If msTermType <> "" Then
                    sQuery = sQuery & ", " & msTermType
                End If
                sQuery = sQuery & " ). "
                
                sCheckText = goALM.GetPrologResult(sQuery, sResultCode)
            
            End If
            If sCheckText = "" Then
                sCheckText = "No errors or warnings were found"
            End If
'        End If
    Else
        ' Expression too big
        sCheckText = msSIZE_ERROR
    End If
    
    DialogWarning sCheckText
    Screen.MousePointer = vbDefault

End Sub

'---------------------------------------------------------------------
Private Function CheckExpressionSyntax(ByRef sErrMsg As String) As Boolean
'---------------------------------------------------------------------
'
' RCW 5/9/02
' Run syntax check on expression.
'
' RCW 18/12/02 - Add additional check for backwards quotes
' NCJ 7 May 03 - Disallow MACRO forbidden keys here
'       Return error message in sErrMsg as appropriate
'---------------------------------------------------------------------
Dim bOK As Boolean

    bOK = False
    sErrMsg = msErrorMsg
    
    ' NCJ 7 May 03
    If gblnValidString(txtExpression.Text, valOnlySingleQuotes) Then

        If msValidationType = "Condition" Then
'            If gclmGuideline.IsValidCondition(txtExpression.Text) = False Then GoTo ErrorHandler
            bOK = gclmGuideline.IsValidCondition(txtExpression.Text)
        ElseIf msValidationType = "Expression" Then
'            If gclmGuideline.IsValidExpression(txtExpression.Text) = False Then GoTo ErrorHandler
            bOK = gclmGuideline.IsValidExpression(txtExpression.Text)
        End If
    Else
        sErrMsg = msValidationType & "s" & gsCANNOT_CONTAIN_INVALID_CHARS

    End If
    
    CheckExpressionSyntax = bOK

End Function

'---------------------------------------------------------------------
Private Sub MatchBrackets()
'---------------------------------------------------------------------
'
' RCW 2/9/02
' Expand selection to include brackets
'
'---------------------------------------------------------------------
Dim sText As String
Dim lngTextLen As Long
Dim lngTextStart As Long
Dim nLBracketCount As Integer
Dim nRBracketCount As Integer
Dim nBracketBalance As Integer

    'record current selected text & its length
    sText = txtExpression.SelText
    lngTextLen = txtExpression.SelLength
    lngTextStart = txtExpression.SelStart
    
    'record number of brackets
    nLBracketCount = CountChar("(", sText)
    nRBracketCount = CountChar(")", sText)
    
    nBracketBalance = Abs(nLBracketCount - nRBracketCount)
    
    Select Case nLBracketCount
        Case Is < nRBracketCount
            'more right brackets than left
            'expand selection to the left only
            Call MatchBracketsOneWay("left", _
                                     nBracketBalance, _
                                     lngTextStart, _
                                     lngTextLen)
        Case Is > nRBracketCount
            'more left brackets than right
            'expand selection to the right only
            Call MatchBracketsOneWay("right", _
                                     nBracketBalance, _
                                     lngTextStart, _
                                     lngTextLen)
        Case Else
            'brackets are balanced
            'expand to next level of brackets
            'left first
            Call MatchBracketsOneWay("left", _
                                     1, _
                                     lngTextStart, _
                                     lngTextLen)
            'then right
            'reset selected text length and start variables
            lngTextLen = txtExpression.SelLength
            lngTextStart = txtExpression.SelStart
            Call MatchBracketsOneWay("right", _
                                     1, _
                                     lngTextStart, _
                                     lngTextLen)

    End Select
        
End Sub


'---------------------------------------------------------------------
Private Function CountChar(ByRef sChar As String, _
                           ByRef sText As String) As Integer
'---------------------------------------------------------------------
'
' RCW 2/9/02
' Count particular characters in a given string
'
'---------------------------------------------------------------------
Dim sSearchChar As String
Dim nCount As Integer
Dim lngPos As Long
    
    sSearchChar = Left(sChar, 1)
    nCount = 0
    
    If Len(sText) > 0 Then
        lngPos = FindChar(sChar, sText, "right", 1)
        
        Do While lngPos > 0
            nCount = nCount + 1
            lngPos = FindChar(sChar, sText, "right", lngPos + 1)
        Loop
    End If
        
    CountChar = nCount
    
End Function


'---------------------------------------------------------------------
Private Function FindChar(ByRef sChar As String, _
                          ByRef sText As String, _
                          ByRef sDirection As String, _
                          ByRef lngStart As Long) As Integer
'---------------------------------------------------------------------
'
' RCW 2/9/02
' Bidirectional Instr
'
'---------------------------------------------------------------------
                          
    If sDirection = "left" Then
        FindChar = InStrRev(sText, sChar, lngStart)
    Else
        FindChar = InStr(lngStart, sText, sChar)
    End If

End Function


'---------------------------------------------------------------------
Private Sub MatchBracketsOneWay(ByRef sDirection As String, _
                                ByRef nBracketCount As Integer, _
                                ByRef lngSelStart As Long, _
                                ByRef lngSelLength As Long)
'---------------------------------------------------------------------
'
' RCW 3/9/02
' Expanded selected text in txtExpression in one direction
'
'---------------------------------------------------------------------
Dim lngSearchStart As Long
Dim lngBracketPos As Long
Dim lngNewPoint As Long
Dim lngOffset As Long
Dim sBracket As String
Dim sOtherBracket As String

    'set some variables according to the direction to be searched
    Call ExpansionInitialise(sDirection, _
                             lngSelStart, _
                             lngSelLength, _
                             lngOffset, _
                             lngSearchStart, _
                             sBracket, _
                             sOtherBracket, _
                             lngNewPoint)

    'different searches needs to be made if we are trying to balance brackets
    'than when we are not
    Call ExpansionSearch(nBracketCount, _
                         lngSelStart, _
                         lngBracketPos, _
                         sBracket, _
                         sOtherBracket, _
                         sDirection, _
                         lngSearchStart, _
                         lngOffset)
    
    'change selected text
    Call ExpansionChange(sDirection, _
                         lngBracketPos, _
                         lngNewPoint, _
                         lngSelStart, _
                         lngSelLength)
    
    
End Sub

'---------------------------------------------------------------------
Private Sub ExpansionInitialise(ByVal sDirection As String, _
                                ByVal lngSelStart As Long, _
                                ByVal lngSelLength As Long, _
                                ByRef lngOffset As Long, _
                                ByRef lngSearchStart As Long, _
                                ByRef sBracket As String, _
                                ByRef sOtherBracket As String, _
                                ByRef lngNewPoint As Long)
'---------------------------------------------------------------------
'
' RCW 3/9/02
' Initialise variables for expansion according to the direction
'
'---------------------------------------------------------------------
    
    If sDirection = "left" Then
        lngOffset = -1
        lngSearchStart = lngSelStart - 1
        sBracket = "("
        sOtherBracket = ")"
        lngNewPoint = 0     'start of txtExpression.text
    Else
        lngOffset = 1
        lngSearchStart = lngSelStart + lngSelLength
        sBracket = ")"
        sOtherBracket = "("
        lngNewPoint = Len(txtExpression.Text)
    End If

End Sub


'---------------------------------------------------------------------
Private Sub ExpansionChange(ByRef sDirection As String, _
                            ByRef lngBracketPos As Long, _
                            ByRef lngNewPoint As Long, _
                            ByVal lngSelStart As Long, _
                            ByVal lngSelLength As Long)
'---------------------------------------------------------------------
'
' RCW 3/9/02
' Make change to expression text
'
'---------------------------------------------------------------------
Dim lngStart As Long
Dim lngLength As Long

    'decide on lngNewPoint if anything found
    If sDirection = "left" Then
        'lngNewPoint is biggest of current lngNewPoint and bracket position
        If lngBracketPos > lngNewPoint Then
            lngNewPoint = lngBracketPos
        End If
        
        'decide new start and end points for selected text
        lngStart = lngNewPoint
        lngLength = (lngSelStart + lngSelLength) - lngStart
    Else
        'lngNewPoint is smallest of space and bracket position
        lngNewPoint = Min0Long(lngNewPoint, lngBracketPos)
        
        'decide new start and end points for selected text
        lngStart = lngSelStart
        lngLength = lngNewPoint - lngStart
    End If
    
    'change selected text
    txtExpression.SelStart = lngStart
    txtExpression.SelLength = lngLength
    
End Sub

'---------------------------------------------------------------------
Private Sub ExpansionSearch(ByRef nBracketCount As Integer, _
                            ByRef lngSelStart As Long, _
                            ByRef lngBracketPos As Long, _
                            ByRef sBracket As String, _
                            ByRef sOtherBracket, _
                            ByRef sDirection As String, _
                            ByRef lngSearchStart As Long, _
                            ByRef lngOffset As Long)
'---------------------------------------------------------------------
'
' RCW 3/9/02
' Find brackets outside current expression
'
'---------------------------------------------------------------------
Dim lngCharPos As Long
Dim lngEndPos As Long
Dim sSearchText As String

    If sDirection = "right" Then
        lngEndPos = Len(txtExpression.Text)
    Else
        lngEndPos = 0
    End If
    
    lngCharPos = lngSearchStart + 1
    sSearchText = txtExpression.Text
    
    If (nBracketCount > 0) And (lngSearchStart <> lngEndPos) Then
        'search for matching brackets
        'Comparison of SearchStart and EndPos added because if they
        'were equal there would be an infinite loop with CharPos
        'increasing from EndPos + 1
        'RCW 19/12/02
        Do While (Abs(lngCharPos - lngEndPos) > 0) _
             And (nBracketCount > 0)

            If Mid(sSearchText, lngCharPos, 1) = sBracket Then
                nBracketCount = nBracketCount - 1
            End If
            If Mid(sSearchText, lngCharPos, 1) = sOtherBracket Then
                nBracketCount = nBracketCount + 1
            End If
            
            lngCharPos = lngCharPos + lngOffset
            
        Loop
    Else
    
        lngCharPos = lngEndPos
        
    End If

    lngBracketPos = lngCharPos
    
End Sub

'---------------------------------------------------------------------
Private Function Min0Long(ByRef lngFirstNum As Long, _
                          ByRef lngSecondNum As Long) As Long
'---------------------------------------------------------------------
'
' RCW 3/9/02
' Long minimum unless one or both of the arguments is 0
'
'---------------------------------------------------------------------

Dim lngMin As Long

    If (lngFirstNum = 0) And (lngSecondNum = 0) Then
        lngMin = 0
    Else
        If lngFirstNum = 0 Then
            lngMin = lngSecondNum
        Else
            If lngSecondNum = 0 Then
                lngMin = lngFirstNum
            Else
                If lngFirstNum > lngSecondNum Then
                    lngMin = lngSecondNum
                Else
                    lngMin = lngFirstNum
                End If
            End If
        End If
    End If
    
    Min0Long = lngMin
            
End Function

'---------------------------------------------------------------------
Private Sub AutoCompleteToggle(nIndex As Integer)
'---------------------------------------------------------------------
'
' RCW 23/9/02
' Toggle autocomplete from off to on or on to off
'
'---------------------------------------------------------------------

    mnuOptionsOpt(nIndex).Checked = Not (mnuOptionsOpt(nIndex).Checked)
    mbBracketOK = mnuOptionsOpt(0).Checked
    mbColonOK = mnuOptionsOpt(1).Checked
    
    If mbBracketOK Then
        Call SetMACROSetting("AutoCompleteBracket", "On")
    Else
        Call SetMACROSetting("AutoCompleteBracket", "Off")
    End If
    If mbColonOK Then
        Call SetMACROSetting("AutoCompleteColon", "On")
    Else
        Call SetMACROSetting("AutoCompleteColon", "Off")
    End If

End Sub

'---------------------------------------------------------------------
Private Sub SetEditMenuOpts()
'---------------------------------------------------------------------
'
' RCW 26/9/02
' Make sure cut, copy and paste are appropriately enabled
' NCJ 20 Sept 06 - Take edit mode into account
'---------------------------------------------------------------------
Dim nSelectedText As Integer

    ' Cut and Copy should only be enabled if something is selected
    ' NCJ 20 Sept 06 - Only allow Cut if can edit
    nSelectedText = txtExpression.SelLength
    If nSelectedText > 0 Then
        mnuEditOpt(mnCutOpt).Enabled = mbCanEdit
        mnuEditOpt(mnCopyOpt).Enabled = True
    Else
        mnuEditOpt(mnCutOpt).Enabled = False
        mnuEditOpt(mnCopyOpt).Enabled = False
    End If
    
    'paste should only be enabled if there is something
    'in the clipboard and it is text
    If Clipboard.GetFormat(1) Then 'clipboard contains text
        If Len(Clipboard.GetText) > 0 Then
            ' NCJ 20 Sept 06 - Only allow Paste if can edit
            mnuEditOpt(mnPasteOpt).Enabled = mbCanEdit
        End If
    Else
        mnuEditOpt(mnPasteOpt).Enabled = False
    End If

End Sub

'---------------------------------------------------------------------
Private Sub txtExpression_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
    
    Call RecordKeyPress(KeyCode, Shift)
    
    'if text has been added then the last button pressed cannot
    'be double-clicked
    msLastButton = ""
    
End Sub

'---------------------------------------------------------------------
Private Sub RecordKeyPress(ByVal nKeyCode As Integer, nShift As Integer)
'---------------------------------------------------------------------
'
' RCW 27/8/02
' Record key pressed
'
'---------------------------------------------------------------------
    
    'The function of this procedure is to correct for the odd feature
    'in VB that records the key pressed a & A as A and 9 & ( as 9.
    If (nShift = vbShiftMask) Then
        'Shift key pressed
        msKeyPressed = Chr(nKeyCode)
        Select Case nKeyCode
            Case 48
                msKeyPressed = ")"  '9 -> (
            Case 57
                msKeyPressed = "("  '0 -> )
'            Case Else
'                msKeyPressed = Chr(nKeyCode)
        End Select
    Else
        If nShift = 0 Then
            'nothing else pressed (shift, ctrl or alt)
            Select Case nKeyCode
'                Case 65 To 90
'                    msKeyPressed = Chr(nKeyCode + 32) 'A -> a
                Case 219
                    msKeyPressed = "["
'                Case Else
'                    msKeyPressed = Chr(nKeyCode)
            End Select
        End If
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub txtExpression_Change()
'---------------------------------------------------------------------
Dim sBracketType As String

    'complete brackets typed
    If (msKeyPressed = "(" Or msKeyPressed = "[") _
    And mbBracketOK Then
        sBracketType = msKeyPressed
        msKeyPressed = ""
        Call CloseBracket("key", sBracketType)
    End If

    'enable or disable match bracket button
    'according to whether there is text present
    Call EnableMatchBracket
    
End Sub

'---------------------------------------------------------------------
Private Sub SetTabOrder()
'---------------------------------------------------------------------
'
' RCW 05/12/02
' Reset tab order of controls outside the tabbed area
'
'---------------------------------------------------------------------
Dim oButton As CommandButton

    'reset the tab-index for each button in the numeric keypad area
    cmdButton(14).TabIndex = NextTab
    cmdButton(44).TabIndex = NextTab
    cmdButton(64).TabIndex = NextTab
    cmdButton(52).TabIndex = NextTab
    cmdButton(53).TabIndex = NextTab
    cmdButton(54).TabIndex = NextTab
    cmdButton(57).TabIndex = NextTab
    cmdButton(58).TabIndex = NextTab
    cmdButton(51).TabIndex = NextTab
    cmdButton(56).TabIndex = NextTab
    cmdButton(55).TabIndex = NextTab
    cmdButton(50).TabIndex = NextTab
    cmdButton(2).TabIndex = NextTab
    cmdButton(3).TabIndex = NextTab
    cmdButton(0).TabIndex = NextTab
    cmdButton(1).TabIndex = NextTab
    cmdButton(70).TabIndex = NextTab
    cmdButton(59).TabIndex = NextTab
    cmdButton(12).TabIndex = NextTab
    cmdButton(13).TabIndex = NextTab
    cmdButton(4).TabIndex = NextTab
    
    'reset the tab index for
    cboDecisions.TabIndex = NextTab
    cboCandidates.TabIndex = NextTab
    cboOtherTasks.TabIndex = NextTab
    cboDataItems.TabIndex = NextTab
    cboDataValues.TabIndex = NextTab
    cboTimeUnits.TabIndex = NextTab
    
    'the OK, cancel, match brackets and check buttons, the edit and
    'help windows and the button tab will always be 0-6 in the tab
    'index
    cmdOK.TabIndex = 0
    cmdCancel.TabIndex = 1
    cmdMatchBrackets.TabIndex = 2
    cmdCheck.TabIndex = 3
    txtExpression.TabIndex = 4
    txtHelp.TabIndex = 5
    tabFunctionSelector.TabIndex = 6
    
End Sub

'---------------------------------------------------------------------
Private Function NextTab() As Integer
'---------------------------------------------------------------------
'
' RCW 05/12/02
' Provide the next tab index for the form and reset the running
' tab index
'
'---------------------------------------------------------------------
    
    mnTabIndex = mnTabIndex + 1
    NextTab = mnTabIndex
    
End Function

'---------------------------------------------------------------------
Private Function CheckSizeOfExpression() As Boolean
'---------------------------------------------------------------------
' RCW 18/12/02
' Check that text for expression is not larger than the ALM can handle
' NCJ 24 Jan 02 - We have a much more strict limit for MACRO
' NCJ 14 May 03 - MACRO limit changed from 2000 to 4000 (use MACRO global const.)
'---------------------------------------------------------------------

'    CheckSizeOfExpression = (Len(txtExpression.Text) < goALM.GoalSizeLimit)
    If glMAX_AREZZO_EXPR_LEN < goALM.GoalSizeLimit Then
        CheckSizeOfExpression = (Len(txtExpression.Text) <= glMAX_AREZZO_EXPR_LEN)
    Else
        CheckSizeOfExpression = (Len(txtExpression.Text) < goALM.GoalSizeLimit)
    End If
    
End Function

'---------------------------------------------------------------------
Private Sub CentreForm()
'---------------------------------------------------------------------
'
' RCW 19/12/02
' Centre form in screen
'
'---------------------------------------------------------------------
Dim lngFormHeight As Long
Dim lngFormWidth As Long
Dim lngScreenHeight As Long
Dim lngScreenWidth As Long
Dim lngNewTop As Long

Dim lngFormBtm As Long
Dim lngFormRight As Long

Const lngTASKBARHEIGHT = 500    'estimated allowance for the size of
                                'the task bar at the bottom of the
                                'screen which is not allowed for in
                                'screen.height

    'get recorded position of screen from registry
    'use -1 as default position rather than a calculated value
    'as a value is calculated afterwards
    mlngRegThisFormLeft = GetMACROSetting("FunctionEditorLeft", -1)
    mlngRegThisFormTop = GetMACROSetting("FunctionEditorTop", -1)
    
    'if the form can fit on the screen using the recorded position
    'then use that otherwise try centering or maximise if all else
    'fails
    lngFormHeight = Me.Height
    lngFormWidth = Me.Width
    lngScreenHeight = Screen.Height
    lngScreenWidth = Screen.Width
    
    lngFormBtm = mlngRegThisFormTop + lngFormHeight
    lngFormRight = mlngRegThisFormLeft + lngFormWidth
    
    If (lngFormBtm > lngScreenHeight) _
    Or (lngFormRight > lngScreenWidth) _
    Or (mlngRegThisFormLeft <= 0) _
    Or (mlngRegThisFormTop <= 0) Then
        'reposition form
        lngNewTop = ((lngScreenHeight - lngFormHeight) / 2) - lngTASKBARHEIGHT
        
        If (lngNewTop > 0) And _
           (lngFormWidth < lngScreenWidth) Then
           'available screen is large enough for form
           'we can position form at the centre of the screen
            Me.Top = lngNewTop
            Me.Left = (lngScreenWidth - lngFormWidth) / 2
        Else
            'screen is not large enough for form so maximise
            Me.WindowState = vbMaximized
        End If
    Else
        'use stored screen position
        Me.Top = mlngRegThisFormTop
        Me.Left = mlngRegThisFormLeft
    End If

End Sub

'---------------------------------------------------------------------
Private Sub EnableMatchBracket()
'---------------------------------------------------------------------
'
' RCW 20/12/02
' Enable match bracket button if there is text in the editor
' NCJ 28 Jan 03 - Also do the Check button and menu option
' NCJ 19 Sept 06 - Also consider mbCanEdit
'---------------------------------------------------------------------
Dim bText As Boolean

    bText = (Trim(txtExpression.Text) <> "")
    cmdMatchBrackets.Enabled = bText
    cmdCheck.Enabled = bText And mbCanEdit
    mnuCheckOpt(0).Enabled = bText And mbCanEdit  ' 0 is "Check" option

End Sub

