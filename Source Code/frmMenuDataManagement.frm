VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMenu 
   AutoShowChildren=   0   'False
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5010
   ClientLeft      =   2595
   ClientTop       =   3120
   ClientWidth     =   13305
   Icon            =   "frmMenuDataManagement.frx":0000
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSymbols 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   0
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   13305
      TabIndex        =   22
      Top             =   4665
      Visible         =   0   'False
      Width           =   13305
      Begin MACRODataManagement.StatusIcons StatusIcons2 
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":08CA
      End
      Begin MACRODataManagement.StatusIcons StatusIcons2 
         Height          =   330
         Index           =   1
         Left            =   1200
         TabIndex        =   24
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":08E6
      End
      Begin MACRODataManagement.StatusIcons StatusIcons2 
         Height          =   330
         Index           =   2
         Left            =   2400
         TabIndex        =   25
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":0902
      End
      Begin MACRODataManagement.StatusIcons StatusIcons2 
         Height          =   330
         Index           =   3
         Left            =   3600
         TabIndex        =   26
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":091E
      End
      Begin MACRODataManagement.StatusIcons StatusIcons2 
         Height          =   330
         Index           =   4
         Left            =   4800
         TabIndex        =   27
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":093A
      End
      Begin MACRODataManagement.StatusIcons StatusIcons2 
         Height          =   330
         Index           =   5
         Left            =   6000
         TabIndex        =   28
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":0956
      End
      Begin MACRODataManagement.StatusIcons StatusIcons2 
         Height          =   330
         Index           =   6
         Left            =   7200
         TabIndex        =   29
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":0972
      End
      Begin MACRODataManagement.StatusIcons StatusIcons2 
         Height          =   330
         Index           =   7
         Left            =   8400
         TabIndex        =   30
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":098E
      End
   End
   Begin VB.PictureBox picUsedForPrinting 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   13245
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   13305
   End
   Begin VB.PictureBox picFunctions 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   13305
      TabIndex        =   1
      Top             =   3810
      Visible         =   0   'False
      Width           =   13305
      Begin MACRODataManagement.StatusFunction StatusFunction 
         Height          =   500
         Index           =   1
         Left            =   1020
         TabIndex        =   2
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   873
         Caption         =   "F3- Previous eForm"
      End
      Begin MACRODataManagement.StatusFunction StatusFunction 
         Height          =   500
         Index           =   2
         Left            =   2160
         TabIndex        =   3
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   873
         Caption         =   "F4- Next eForm"
      End
      Begin MACRODataManagement.StatusFunction StatusFunction 
         Height          =   500
         Index           =   3
         Left            =   3240
         TabIndex        =   12
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   873
         Caption         =   "F5- Print eForm"
      End
      Begin MACRODataManagement.StatusFunction StatusFunction 
         Height          =   500
         Index           =   4
         Left            =   4320
         TabIndex        =   4
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   873
         Caption         =   "F6- Save and return"
      End
      Begin MACRODataManagement.StatusFunction StatusFunction 
         Height          =   500
         Index           =   5
         Left            =   5400
         TabIndex        =   5
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   873
         Caption         =   "F7- Save eForm"
      End
      Begin MACRODataManagement.StatusFunction StatusFunction 
         Height          =   500
         Index           =   6
         Left            =   6480
         TabIndex        =   6
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   873
         Caption         =   "F8- Cancel and return"
      End
      Begin MACRODataManagement.StatusFunction StatusFunction 
         Height          =   500
         Index           =   7
         Left            =   7560
         TabIndex        =   7
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   873
         Caption         =   "F9- Clear response"
      End
      Begin MACRODataManagement.StatusFunction StatusFunction 
         Height          =   500
         Index           =   8
         Left            =   8610
         TabIndex        =   8
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   873
         Caption         =   "F10- Question menu"
      End
      Begin MACRODataManagement.StatusFunction StatusFunction 
         Height          =   500
         Index           =   9
         Left            =   9720
         TabIndex        =   9
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   873
         Caption         =   "F11- New comment"
      End
      Begin MACRODataManagement.StatusFunction StatusFunction 
         Height          =   500
         Index           =   10
         Left            =   10800
         TabIndex        =   10
         Top             =   0
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   873
         Caption         =   "F12- Remove comments"
      End
      Begin MACRODataManagement.StatusFunction StatusFunction 
         Height          =   500
         Index           =   0
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   873
         Caption         =   "F1- Help"
      End
   End
   Begin VB.PictureBox picSymbols 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   1
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   13305
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   13305
      Begin MACRODataManagement.StatusIcons StatusIcons 
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":09AA
      End
      Begin MACRODataManagement.StatusIcons StatusIcons 
         Height          =   330
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":09C6
      End
      Begin MACRODataManagement.StatusIcons StatusIcons 
         Height          =   330
         Index           =   2
         Left            =   2400
         TabIndex        =   15
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":09E2
      End
      Begin MACRODataManagement.StatusIcons StatusIcons 
         Height          =   330
         Index           =   3
         Left            =   3600
         TabIndex        =   16
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":09FE
      End
      Begin MACRODataManagement.StatusIcons StatusIcons 
         Height          =   330
         Index           =   4
         Left            =   4800
         TabIndex        =   17
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":0A1A
      End
      Begin MACRODataManagement.StatusIcons StatusIcons 
         Height          =   330
         Index           =   5
         Left            =   6000
         TabIndex        =   18
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":0A36
      End
      Begin MACRODataManagement.StatusIcons StatusIcons 
         Height          =   330
         Index           =   6
         Left            =   7200
         TabIndex        =   19
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":0A52
      End
      Begin MACRODataManagement.StatusIcons StatusIcons 
         Height          =   330
         Index           =   7
         Left            =   8400
         TabIndex        =   20
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":0A6E
      End
      Begin MACRODataManagement.StatusIcons StatusIcons 
         Height          =   330
         Index           =   8
         Left            =   9660
         TabIndex        =   31
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":0A8A
      End
      Begin MACRODataManagement.StatusIcons StatusIcons 
         Height          =   330
         Index           =   9
         Left            =   10860
         TabIndex        =   32
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "Label1"
         Picture         =   "frmMenuDataManagement.frx":0AA6
      End
   End
   Begin VB.Timer tmrSystemIdleTimeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   360
      Top             =   1680
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   900
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "MS Sans Serif"
   End
   Begin MSComctlLib.ImageList imglistSmallIcons 
      Left            =   1140
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuDataManagement.frx":0AC2
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuDataManagement.frx":0DDC
            Key             =   "TwoPeople"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglistSmallIconsOLD 
      Left            =   180
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuDataManagement.frx":10F6
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuDataManagement.frx":1208
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuDataManagement.frx":131A
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuDataManagement.frx":142C
            Key             =   "Hide"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuDataManagement.frx":1746
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuDataManagement.frx":1A60
            Key             =   "TwoPeople"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopupContainer"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpSubItem 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998-2006. All Rights Reserved
'   File:       frmMenuDataManagement.frm
'   Author:     Andrew Newbigging, November 1997
'   Purpose:    Main menu used in MACRO Data Management
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   ATN 11/12/99    Added Transfer Data to Communication menu
'                   Modified InitialiseMe to check for command line switches:
'                   /ai     do an autoimport of messages not yet processed
'                           sub 'AutoImport' created to do this.
'                   /tr     do a send/receive communication with MACRO Exchange server
'                   Modified _queryunload to NOT try to shut down the PLM
'                   , if one of the above switches has been used.
' NCJ/TA Sept 01: Various changes to integrate new MACRODEBS component
' TA 27/9/01: Study properties stored in frmMenu removed
' TA 2/10/01: Better error handling
' NCJ 4 Oct 01 - Reinstated PSS (but it's only used for DeleteProformaTrial)
' TA 11/10/01: Fixes to View menu display and tidying
' DPH 18/10/2001 - Remove References to RemoveLockOnRecord
' NCJ 18 Oct 01 - goArezzo.Init now uses new GetPrologSwitches
' DPH 25/10/2001 - Removed Reports menu item and references in code
' TA 08/11/2001 - Ensure that Discrepancies and Notes are availaible on the View Menu
' DPH 17/01/2002 - Returning add comments functionality
' NCJ 22 Mar 02 - Only show ReadOnly message if it's non-empty
' TA 03/04/02 - Added parameter to SubjectSelectandOpen to prevent open form being shown
' TA 03/04/02 - Added OC discrepancy code done in Palo Alto
' TA 10/04/02 - Only show OC Discrepancy menu item if the view exists
' DPH 11/04/2002 - AutoImport Changes
' DPH 15/04/2002 - Making Data Transfer modal
' DPH 16/05/2002 - Added command line data transfer with conditional compilation arguements
' DPH 14/06/2002 - Removed AutoImport functionality from MACRO_DM
' MLM 10/07/02: CBB 2.2.19/11: Removed check of "Monitor/review data" permission in MIMessageOpen
'TA CBB 2.2.20.3 18/7/02: ensure our row in ArezzoToken table is cleared when we close MACRO
'ASH 10/11/2002 - Added call to InitialiseSettingsFile.
' NCJ 18 Sept 02 - Deal with failure of frmEFormDataEntry.Display in EFIOpen
' ZA 24/09/2004 - Remove reference to PSS
'   TA 26/09/02: Changes for New UI -opening new web forms and resizing of them
'   TA 01/10/2002: Split Subject SelectAndOPen into ShowSubjectList and SubjectOpen
'   TA 04/10/2002: Added ShowPopUpMenu that is easier to use than ShowPopup
' TA 15/10/2002: Mores changes to the User Interface
' NCJ 16 Oct 02 - Added new Subject SDVs menu
'TA 22/10/2002: Added more error handlers
' NCJ 28 Nov 02 - Get Arezzo switches using ArezzoMemory class in InitialiseMe
' NCJ 24 Dec 02 - Removed AutoImport routine since it's no longer used here
'TA 09/01/2003 - RefreshSearchResults to simulate the refresh button in the left hand menu being pressed
'TA 14/01/2003: Check permissions and study/site allowed in switch user
'TA 19/01/2003: Unload subject when closing eform when it wasn't opened from the schedule
'RS 20/01/2003: Added ToggleLocalFormat function (Local Date/Time format)
' NCJ 21 Jan 03 - Added call to LocalDateFormat dialog
' NCJ 22 Jan 03 - Only offer to transfer data on exit if they have Transfer Data permission
' NCJ 23 Jan 03 - Removed erroneous message about changing local date formats when changing split screen
' TA 23/01/2003 - Close subject (unload from memory so data xfer works) when home form , subject list and new subject form displayed
' TA 24/01/2003 - Removed all studies option on study dropdown combo
' NCJ 28 Jan 03 - Pass DB connection string to clsArezzoMemory
' TA 13/02/2003 - Do not unload subject when closing the eform if we are switching user
' RS 18/02/2003 - Moved Communication Menu into HTML Panel, changed menu methods to public methods called from JSfunctions
' TA 05/03/2003 - Allowed toggling of the left hand menu being shown
' TA 05/03/03 - new background for login code
' NCJ 20 Mar 03 - Prevent crash in MDIForm_QueryUnload if it's a Timeout exit
' TA 03/04/2003 - No need to unload webbrowser forms any more before refreshing them
' TA 08/04/2003: Status controls are now part of DM and status functions show double lines of text
' NCJ 13 May 03 - Must refresh the schedule and the menu after successful registration
' TA 21/05/2003 - We now check for data transfer when logout is pressed
' TA 21/05/2003: changed message wording in ConfirmDataXfer
' TA 29/05/2003: minor changes to laod order in InitWebForms
' NCJ 10 Jun 03 - Got Templates working correctly
' ic 23/06/2003 added bInit parameter to LoadLhCombos()
' NCJ 25 Jun 03 - MACRO 3.0 Bug 1830 - Sorted out permissions for View Discs/SDVs
' ic 13/08/2003 bug 1946, initialise web eform top height in InitialiseMe()
' NCJ 2 Sept 03 - Bug 1989 - In UserSwitch make sure we recognise change of ChangeData permissions
' REM 12/03/04 - For MACRO Desktop Edition only, in routine EnableDisableTaskListItems. Add disable of Create new subject menu item is database is bigger than 1 gig
' NCJ 22 Mar 06 - Issue 2690 - Added AREZZO Setting in InitialiseMe
' NCJ 7 Jun 06 - Issue 2690 - Change default validation ordering
' NCJ 21 Nov 06 - Added Partial Dates setting to AREZZO in InitialiseMe (but not yet used)
'------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------'

Option Explicit
Option Compare Binary
Option Base 0

'public instance of OC class for copy/paste of OC discrepancies
' this is public so that frmNewDiscrepancy can use it
Public gOC As clsOC

'   ATN 10/12/99 - General object for remote communication settings
Public gTrialOffice As New clsCommunication

'hold a reference to the data entry form to catch its events
Private WithEvents mofrmEFormDataEntry As frmEFormDataEntry
Attribute mofrmEFormDataEntry.VB_VarHelpID = -1

''hold a reference to the open subject form to catch its events
'Private WithEvents mofrmSubjectList As frmSubjectList

'hold a reference to the new subject form to catch its events
Private WithEvents mofrmNewSubject As frmNewSubject
Attribute mofrmNewSubject.VB_VarHelpID = -1

'TA 19/04/2000 - store popup item selected
Private mlPopUpItem As Long

 'TA 25/09/2002: New UI code
Private mofrmMenuLh As frmWebBrowser
Private mofrmMenuTop As frmWebBrowser
Private mofrmMenuHome As frmWebBrowser
Private mofrmHeaderLh As frmWebBrowser
Private mofrmFooterLh As frmWebBrowser
Private mofrmSchedule As frmWebBrowser
Private mofrmeFormTop As frmWebBrowser
Private mofrmeFormLh As frmWebBrowser

'display background pic for mdi form
Private mofrmMDIBackGRound As frmWebBrowser

'new subject list
Private mofrmSubjectList As frmWebBrowser

'public for use by the hourglass form
Public mofrmBorderTop As frmBorder
Public moFrmBorderBottom As frmBorder

'are we in split screen mode
Public SplitScreen As Boolean

' NCJ 22 Jan 03 - Made this private because it's only used in frmMenu
Private mbDataBeingTransferred As Boolean

'indicate whther we are in process of switching user
Private mbSwitchingUser As Boolean

'stores width of left hand panes (can be 0 or WEB_LH_WIDTH)
Private mlWebLhWidth As Long
'stores height of headerlh (can be WEB_APP_MENU_TOP_HEIGHT or WEB_APP_HEADER_LH_HEIGHT)
Private mlWebAppHeaderLhHeight As Long


'---------------------------------------------------------------------
Public Sub SetUpTrialOffice()
'---------------------------------------------------------------------
' NCJ 7/3/00
' Use new clsComms and extract "current" record
'---------------------------------------------------------------------
Dim ocolComms As clsComms

    Set ocolComms = New clsComms
    ' Get all the settings
    ocolComms.Load
    ' Pick off the current one (i.e. whose dates include today)
    Set gTrialOffice = ocolComms.GetCurrentRecord
    
    ' Don't need this any more
    Set ocolComms = Nothing
    
    ' Check that there was a record, otherwise set to empty object
    If gTrialOffice Is Nothing Then
        Set gTrialOffice = New clsCommunication
    End If

End Sub

'---------------------------------------------------------------------
Private Sub MDIForm_Unload(Cancel As Integer)
'---------------------------------------------------------------------

    SaveFormDimensions Me

End Sub

'----------------------------------------------------------------------------------------------
Public Sub ToggleUserIntervention(bEnable As Boolean)
'----------------------------------------------------------------------------------------------
' NCJ 24 Jan 03 - Disable/Enable user intervention
' For use when e.g. opening eForms
'----------------------------------------------------------------------------------------------

    mofrmMenuHome.Enabled = bEnable
    mofrmFooterLh.Enabled = bEnable
    mofrmHeaderLh.Enabled = bEnable
    mofrmMenuLh.Enabled = bEnable
    mofrmMenuTop.Enabled = bEnable

End Sub

'---------------------------------------------------------------------
Public Sub ResetTransferStatus()
'---------------------------------------------------------------------
' Set all "Changed" flags to "Changed"
' NCJ 6/6/00 - Added confirmatory message box
' NCJ 10/10/00 - SR 3972 Moved to Communication Menu from File Menu
' RS 18/02/2003 - Moved to HTML Panel
'---------------------------------------------------------------------
Dim sMsg As String

    sMsg = "This will mark ALL existing data as 'Not exported', "
    sMsg = sMsg & "meaning it will all be exported next time you do data transfer."
    sMsg = sMsg & vbNewLine & vbNewLine & "Are you sure you wish to continue?"
    If MsgBox(sMsg, vbYesNo, "MACRO Data Transfer Status") = vbYes Then
        Call ResetTransferFlags
    End If
    
End Sub

'---------------------------------------------------------------------
Public Sub CommunicationStatusReport()
'---------------------------------------------------------------------
' Show communication transfer report, i.e. how many records sent etc.
' NCJ 10/10/00 - Moved to Communication menu from File menu
' RS 18/02/2003 - Moved to HTML Panel
'---------------------------------------------------------------------
    
    Call ShowDocument(Me.hWnd, CheckDataItemResponseHistory)

End Sub

'---------------------------------------------------------------------
Public Sub RemoteTimeSynchronisation()
'---------------------------------------------------------------------
'WillC 3/8/00 added the time synchronisation  form to the project and
'added the new menu option to call it.
' RS 18/02/2003 - Moved to HTML Panel
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    frmTimeSynch.Show vbModal

Exit Sub
ErrHandler:
    
    mbDataBeingTransferred = False
    
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuCTimeSynch_Click")
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
Public Sub TransferData()
'---------------------------------------------------------------------
' REVISIONS
' DPH 15/04/2002 - One transfer call rather than two
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If IsSubjectOpen Then
        DialogInformation "Data transfer cannot occur while a subject is open"
    Else
        'changed Mo Morris 21/3/00, SR 3191
        mbDataBeingTransferred = True
    
        ' DPH 15/04/2002 - To make transfer screen modal need just one call
        RemoteCommsTransfer
        
        'changed Mo Morris 21/3/00, SR 3191
        DoEvents
        mbDataBeingTransferred = False
        
        Call goUser.ReloadStudySitePermissions
        RefreshAfterUserChange
    
    End If

Exit Sub
ErrHandler:
    
    mbDataBeingTransferred = False
    
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuCTransfer_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select


End Sub


Public Sub ViewLockFreeze()
    frmViewLockFreeze.Display
End Sub

'---------------------------------------------------------------------
Public Sub Register()
'---------------------------------------------------------------------
' NCJ 23/11/00
' Attempt registration of subject
'---------------------------------------------------------------------

    If IsSubjectOpen Then
        If EnableRegistrationMenu Then
            If RegisterSubject Then
                If Not FormIsLoaded(g_DATAENTRY_FORM_NAME) Then
                    ' NCJ 13 May 03 - Must refresh the schedule in case eForms are dependent on registration
                    Call RefreshSchedule
                End If
                ' And update menu options
                Call EnableDisableTaskListItems
            End If
        Else
            DialogInformation "Cannot register subject at this time"
        End If
    Else
        DialogInformation "You have no subject open for registration"
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuPopUpSubItem_Click(Index As Integer)
'---------------------------------------------------------------------
'TA 19/04/2000 - store item clicked on user-defined menu
'---------------------------------------------------------------------

     mlPopUpItem = Index + 1
    
End Sub
 
'---------------------------------------------------------------------
Public Sub ToggleFunctionKeys(bShow As Boolean)
'---------------------------------------------------------------------
' toggle display of function keys along the bottom
'---------------------------------------------------------------------
'Dim bShow As Boolean

    'bShow = Not picFunctions.Visible
    If bShow <> goUser.UserSettings.GetSetting(SETTING_VIEW_FUNCTION_KEYS, False) Then
        picFunctions.Visible = bShow
        'Save status bar setting so it is re-instated when user next used the app.
        goUser.UserSettings.SetSetting SETTING_VIEW_FUNCTION_KEYS, bShow
        MDIForm_Resize
    End If
        
End Sub

'---------------------------------------------------------------------
Public Sub ToggleSymbols(bShow As Boolean)
'---------------------------------------------------------------------
' toggle display of stuses along the bottom
'---------------------------------------------------------------------
'Dim bShow As Boolean

    'bShow = Not picSymbols.Visible
    
    If bShow <> goUser.UserSettings.GetSetting(SETTING_VIEW_SYMBOLS, False) Then
        picSymbols(0).Visible = bShow
        picSymbols(1).Visible = bShow
        'Save status bar setting so it is re-instated when user next used the app.
        goUser.UserSettings.SetSetting SETTING_VIEW_SYMBOLS, bShow
        
        MDIForm_Resize
    End If
        
End Sub

'---------------------------------------------------------------------
Public Sub ToggleSplitScreen(bSplit As Boolean)
'---------------------------------------------------------------------
' toggle display of stuses along the bottom
'---------------------------------------------------------------------
    
    If bSplit <> goUser.UserSettings.GetSetting(SETTING_SPLIT_SCREEN, False) Then
        goUser.UserSettings.SetSetting SETTING_SPLIT_SCREEN, bSplit
        SplitScreen = bSplit
        If bSplit Then
            moFrmBorderBottom.Display False
            'send to back
            moFrmBorderBottom.ZOrder 1
'            CloseWinForm wfDataBrowser
'            CloseWinForm wfDiscepancies
'            CloseWinForm wfNotes
'            CloseWinForm wfSDV
            EnsureCaptionsCorrect
        Else
            'unload data browser and MIMessage browser if open
            moFrmBorderBottom.Hide
            If FormIsLoaded("frmDataItemResponse") Then
                Unload frmDataItemResponse
            End If
            If FormIsLoaded("frmViewDiscrepancies") Then
                Unload frmViewDiscrepancies
                
            End If
            'bring eform forms to the top if open
            If FormIsLoaded(g_DATAENTRY_FORM_NAME) Then
                frmEFormDataEntry.ZOrder
                mofrmeFormTop.ZOrder
                mofrmeFormLh.ZOrder
            End If
        End If
        MDIForm_Resize
    End If
        
End Sub

'---------------------------------------------------------------------
Public Sub ToggleSameEform(bSame As Boolean)
'---------------------------------------------------------------------
' toggle opening of subjects at the same eform
'---------------------------------------------------------------------
    
    If bSame <> goUser.UserSettings.GetSetting(SETTING_SAME_EFORM, False) Then
        goUser.UserSettings.SetSetting SETTING_SAME_EFORM, bSame
    End If
        
End Sub

'---------------------------------------------------------------------
Public Sub ToggleLocalFormat(bLocal As Boolean)
'---------------------------------------------------------------------
' Toggle use & format of local date/time
' RS 20/01/2003
' NOTE: This function sets SETTING_LOCAL_FORMAT and SETTING_LOCAL_DATE_FORMAT
'---------------------------------------------------------------------
Dim sLocalDateFormat As String
Dim sLocalTimeFormat As String

' NB - LH Panel checkbox
'//check/uncheck option checkboxes - optFunctionKeys,optSymbols,optDateFormat,optSplitScreen,optSameForm
'Function fnSetOptionChecked(sName, bChecked)

    If bLocal <> goUser.UserSettings.GetSetting(SETTING_LOCAL_FORMAT, False) Then

        ' Can't change things while a subject is open
        If IsSubjectOpen Then
            DialogInformation "Local date formats cannot be changed while a subject is open"
            ' Toggle the check box back
            Call mofrmMenuLh.ExecuteJavaScript.fnSetOptionChecked(("optDateFormat"), (Not bLocal))
            Exit Sub
        End If
        
        ' Toggle the "local" setting
        goUser.UserSettings.SetSetting SETTING_LOCAL_FORMAT, bLocal
        
        If bLocal Then
            ' Ask the user to enter/edit the date format
            sLocalDateFormat = goUser.UserSettings.GetSetting(SETTING_LOCAL_DATE_FORMAT, "dd/mm/yyyy")
            
            If frmLocalFormats.Display(sLocalDateFormat, goArezzo) Then
                If sLocalDateFormat <> "" Then
                    goUser.UserSettings.SetSetting SETTING_LOCAL_DATE_FORMAT, sLocalDateFormat
                    ' Update the Options panel
                    Call mofrmMenuLh.ExecuteJavaScript.fnSetDateFormatLabel((sLocalDateFormat))
                End If
            End If
        End If

    End If
        
End Sub

'---------------------------------------------------------------------
Public Property Get DefaultDateFormat() As String
'---------------------------------------------------------------------
' NCJ 9 Feb 00 - default date format for study
' NCJ 24 Sep 01 - Read from StudyDef object
'---------------------------------------------------------------------
Const sFormat = "dd/mm/yyyy"

    If goStudyDef Is Nothing Then
        ' Arbitrary date format to use if none other available
        DefaultDateFormat = sFormat
    Else
        If goStudyDef.DateFormat > "" Then
            DefaultDateFormat = goStudyDef.DateFormat
        Else
            DefaultDateFormat = sFormat
        End If
    End If
    
End Property

'---------------------------------------------------------------------
Public Sub InitialiseMe()
'---------------------------------------------------------------------
' This gets called from Main (in MainMACROModule) at startup
' to do things specific to Data Management module
' NCJ 15 Sept 99
' NCJ 17 Jan 01 - Do not call InitialiseCollections here
' DPH 15/04/2002 - Replaced Import / Export with 1 call
' DPH 14/06/2002 - Removed AutoImport functionality from MACRO_DM
' NCJ 28 Nov 02 - Use ArezzoMemory class to get PrologSwitches
' ic 13/08/2003 bug 1946, initialise web eform top height
' NCJ 22 Mar 06 - Added AREZZO Setting for issue 2690
' NCJ 21 Nov 06 - Added AREZZO setting for PDs (but not yet used in AREZZO)
'---------------------------------------------------------------------
Dim oArezzoMemory As clsAREZZOMemory

    ' NCJ 13/1/00 - Use pointer to new 2.0 Help Files
    ' NCJ 23 Apr 03 - Use generic help path (already set up) for MACRO 3.0 Help
'    If LCase$(Command) = "review" Then
'        gsMACROUserGuidePath = gsMACROUserGuidePath & "DR\Contents.htm"
'    Else
'        gsMACROUserGuidePath = gsMACROUserGuidePath & "DE\Contents.htm"
'    End If
    
    'ic 13/08/2003 bug 1946, initialise web eform top height
    WEB_EFORM_TOP_HEIGHT = DEFAULT_WEB_EFORM_TOP_HEIGHT
    
    '  Set up communication settings in the global object
    Call SetUpTrialOffice
    
    ' Set remote site - NCJ 1/12/99, ATN 11/12/99
    If Me.gTrialOffice.TrialOffice = "" Then
        gblnRemoteSite = False
    Else
        gblnRemoteSite = True
    End If

    mbDataBeingTransferred = False
    
    '   ATN 11/12/99
    '   Check for command line switches.
    If UCase(Left(Command, Len(gsAUTO_IMPORT))) = UCase(gsAUTO_IMPORT) Then
        ' DPH 14/06/2002 - Removed AutoImport functionality from MACRO_DM
        gLog gsAUTOIMPORT, "AutoImport functionality has been removed from MACRO Data Entry. Please use the MACRO AutoImport module."
        ExitMACRO
        
    ElseIf UCase(Left(Command, Len(gsTRANSFER_DATA))) = UCase(gsTRANSFER_DATA) Then
        ' DPH 15/04/2002 - Replaced Import / Export with 1 call
#If BackgroundXfer Then
        RemoteCommsModelessTransfer
#Else
        RemoteCommsTransfer
#End If
        ExitMACRO
    Else
        ' No command line parameters
        
        'init split screen flag
        SplitScreen = goUser.UserSettings.GetSetting(SETTING_SPLIT_SCREEN, False)
        
        
        'Reset status bars to how they were when user exited the app.
        picFunctions.Visible = goUser.UserSettings.GetSetting(SETTING_VIEW_FUNCTION_KEYS, False)
        picSymbols(0).Visible = goUser.UserSettings.GetSetting(SETTING_VIEW_SYMBOLS, False)
        picSymbols(1).Visible = goUser.UserSettings.GetSetting(SETTING_VIEW_SYMBOLS, False)
        Call InitWebForms
        
          ' Create and initialise a new Arezzo instance
        Set goArezzo = New Arezzo_DM
        ' NCJ 18/10/01 - Get the Prolog memory settings using GetPrologSwitches
        ' NCJ 28 Nov 02 - Get switches using ArezzoMemory class
        Set oArezzoMemory = New clsAREZZOMemory
        ' Using TrialId = 0 gets max. values for all studies
        Call oArezzoMemory.Load(0, goUser.CurrentDBConString)
        Call goArezzo.Init(gsTEMP_PATH, oArezzoMemory.AREZZOSwitches)
        Set oArezzoMemory = Nothing
        
        ' NCJ 22 Mar 06 - Add AREZZO setting for validation ordering (default to "old")
        ' NCJ 7 Jun 06 - Default is now "idorder" (so users must explicitly switch off new ordering)
        Call SetAREZZOSetting("warningorder", GetMACROSetting("warningorder", "idorder"))
        ' NCJ 21 Nov 06 - Added Partial Dates = No (but this isn't used in 3.0.76)
        Call SetAREZZOSetting("partialdates", GetMACROSetting("partialdates", "no"))
        
        MDIForm_Resize
'        'set the variable for the last browser shown
'        gLastBrowser = lbNeither
        
'REM 24/04/02 - load last StudyDef, which stored in the registry
#If VTRACK <> 1 Then
        'call loadstudy with StudyId of 0 and GetFromReg = true
        LoadStudy goStudyDef, gsADOConnectString, 0, 1, goArezzo, True
#End If
    End If

'REM 12/03/04 - For MACRO Desktop Edition only.  Warn user that database is approaching the max size of 1 GB.
#If DESKTOP = 1 Then
    If (DatabaseSize > 820) Or (DatabaseSize < 1024) Then
        MsgBox "Your MACRO database is approaching the maximum size of 1 GB.", vbInformation, "MACRO Database"
    End If
#End If

End Sub

'---------------------------------------------------------------------
Private Sub InitWebForms()
'---------------------------------------------------------------------
'initialise web formas
'---------------------------------------------------------------------
On Error GoTo ErrLabel
    
    'initialise sizes so that left hand pane is showing
    mlWebLhWidth = WEB_LH_WIDTH
    mlWebAppHeaderLhHeight = WEB_APP_HEADER_LH_HEIGHT
    
    Set mofrmHeaderLh = New frmWebBrowser
    Set mofrmMenuLh = New frmWebBrowser
    Set mofrmFooterLh = New frmWebBrowser
    Set mofrmMenuTop = New frmWebBrowser
    Set mofrmMenuHome = New frmWebBrowser
    Set mofrmSchedule = New frmWebBrowser
    Set mofrmeFormTop = New frmWebBrowser
    Set mofrmeFormLh = New frmWebBrowser
    Set mofrmSubjectList = New frmWebBrowser
    
    Set mofrmBorderTop = New frmBorder
    Set moFrmBorderBottom = New frmBorder

    With mofrmMenuTop
        .Top = 0
        .Left = 0
        .Height = WEB_APP_MENU_TOP_HEIGHT
        .Width = Me.ScaleWidth
    End With

    With mofrmHeaderLh
        .Top = mofrmMenuTop.Height
        .Left = 0
        .Width = mlWebLhWidth
        .Height = WEB_APP_HEADER_LH_HEIGHT
    End With

    With mofrmMenuLh
        .Top = mofrmHeaderLh.Top + mofrmHeaderLh.Height
        .Left = 0
        .Width = mlWebLhWidth
        .Height = Me.ScaleHeight - mlWebAppHeaderLhHeight - WEB_APP_FOOTER_LH_HEIGHT
    End With
          
    With mofrmFooterLh
        .Top = Me.ScaleHeight - WEB_APP_FOOTER_LH_HEIGHT
        .Left = 0
        .Width = mlWebLhWidth
        .Height = WEB_APP_FOOTER_LH_HEIGHT
    End With
    
    mofrmSchedule.Visible = False
    mofrmBorderTop.Visible = False
    moFrmBorderBottom.Visible = False
    
    Call ResizeBottomRightHandForms
    
    'TA 06/03/03 - new background for login
    'unlaod previsou background
    Unload mofrmMDIBackGRound
    'refesh screen so it goes white
    DoEvents
    
    'display expand button if left hand menu pane not shown
    Call mofrmHeaderLh.Display(wdtHTML, AppHeaderLhHTML(mlWebLhWidth = 0), "no")

    Call mofrmMenuLh.Display(wdtHTML, AppMenuLhHTML, "auto")
    mofrmMenuLh.ExecuteJavaScript.fnPageLoaded
    
    
    Call mofrmFooterLh.Display(wdtHTML, AppFooterLhHTML, "no")
    
    Call mofrmMenuTop.Display(wdtHTML, AppMenuTopHTML(goUser), "no")
    mofrmMenuTop.ExecuteJavaScript.fnPageLoaded
    
    Call mofrmSchedule.Display(wdtUrl, "", "no")

    Call mofrmeFormTop.Display(wdtHTML, eFormTopHTML, "no")
    Call mofrmeFormLh.Display(wdtHTML, eFormLhHTML, "auto")



    'show the border
    mofrmBorderTop.Display True
    
    If SplitScreen Then
        moFrmBorderBottom.Display False
    End If

    Call LoadLhCombos
    
    
    'TA 30/11/2003: no need to update disc count as this isdone in AppMenuLhHTML
    'put SDV.disc count in task pane
'    UpdateDiscCount
    'enable/disable taskteims
    EnableDisableTaskListItems
    
' TA 29/05/2003: minor changes to laod order  - showhome done last
    Call ShowHome
    Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|frmMenu.InitWebForms"
    
End Sub

'---------------------------------------------------------------------
Public Sub LoadLhCombos(Optional lStudyId As Long = -1)
'---------------------------------------------------------------------

'---------------------------------------------------------------------
Dim oJs As JSFunctions
Dim sList As String
Dim oStudy As Study

    On Error GoTo ErrLabel

    HourglassOn
    
    Set oJs = New JSFunctions
    
    If lStudyId = -1 Then
        'we need to load studies
    '//      fnLoadSelect('fltSt',lstStudies,false,true);
        sList = "" '"`All studies" ' -1 signifies all studies
        If goUser.GetAllSites.Count > 0 Then
            For Each oStudy In goUser.GetAllStudies
                sList = sList & "|" & oStudy.StudyId & "`" & oStudy.StudyName '& "`" & oStudy.StudyId
            Next
        End If
        'ic 23/06/2003 added bInit parameter
        Call oJs.fnLoadSelect(mofrmMenuLh, "fltSt", sList, False)
        
        If goUser.GetAllStudies.Count = 0 Then
            DialogWarning "There are no studies available to you"
            HourglassOff
            Exit Sub
        Else
            lStudyId = goUser.GetAllStudies(1).StudyId
        End If
        
    End If
    
    'fill the rest according to our choice
    Call oJs.LoadLhCombos(mofrmMenuLh, lStudyId)
    
    HourglassOff
    
    Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|frmMenu.LoadLhCombos"
End Sub

'---------------------------------------------------------------------
Public Sub ShowHome()
'---------------------------------------------------------------------
'display the home page
'---------------------------------------------------------------------
Dim sHomeURL As String

    On Error GoTo ErrLabel


    'try to read home url from settings file
    sHomeURL = GetMACROSetting("homeurl", "")
    
    If sHomeURL = "" Then
        'no defined url use the default one if it exists
        If FileExists(gsDOCUMENTS_PATH & WEB_appMenuHome_URL) Then
            'only show this file if it exists
            sHomeURL = gsDOCUMENTS_PATH & WEB_appMenuHome_URL
        Else
            sHomeURL = ""
        End If
    End If


    If LCase(Right(sHomeURL, 4)) = ".asp" Then
        'assume they are going to the home page for reports so give them user details
        Call mofrmMenuHome.Display(wdtUrl, sHomeURL, , , , goUser.GetStateHex(False))
    Else

        Call mofrmMenuHome.Display(wdtUrl, sHomeURL)
    End If
    
    If IsSubjectOpen Then
        'close subject
        CloseSubject goStudyDef, True, False, False
    End If
    
    OpenWinForm wfHome
    
    Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|frmMenu.ShowHome"
    
End Sub


'---------------------------------------------------------------------
Private Sub MDIForm_Load()
'---------------------------------------------------------------------
' Turn on key preview for form, so that F1 (Help) can be trapped by form
'---------------------------------------------------------------------
Dim btnX As Button
   
  On Error GoTo ErrHandler
     
   
    SetFormDimensions Me
    
    Me.Caption = GetApplicationTitle

    Call SetupSymbols

    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "MDIForm_Load")
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
Public Sub DisplayMDIBackGround()
'---------------------------------------------------------------------
'diaply background that appears behing login screen
'---------------------------------------------------------------------

' TA 05/03/03 - new background for login code
     Set mofrmMDIBackGRound = New frmWebBrowser
     'make much higher and wider the mdi form so that no white border is shown
     With mofrmMDIBackGRound
        .Top = -240
        .Left = -240
        .Width = Me.ScaleWidth + 480
        .Height = Me.ScaleHeight + 480
        'show background pic
        .Display wdtHTML, "<img height='100%' width='100%' src='" & App.Path & "/www/img/bg.jpg'>", "no"
    End With

End Sub

'---------------------------------------------------------------------
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---------------------------------------------------------------------
' NCJ 22 Jan 03 - Check for new LF message to be sent
' NCJ 20 Mar 03 - Only do a full shutdown if the user has clicked the Close button
'---------------------------------------------------------------------


    Set gOC = Nothing
'Changed Mo Morris 21/3/00, SR 3191
    'Stop the QueryUnload if Data Transfer is taking place is active
    If mbDataBeingTransferred Then
        Cancel = 1
        Exit Sub
    End If
    
    HourglassOn
    
    'TA CBB 2.2.20.3 18/7/02: ensure our row in ArezzoToken table is cleared when we close MACRO
    Call DeleteArezzoToken
    
    ' NCJ 20 Mar 03 - Check for type of close down
    If UnloadMode = vbFormControlMenu Then
        ' User has clicked Close box - do a tidy up before exiting
        
        '  If a record is open, save the details
        If IsSubjectOpen Then
            ' Prompt user to close down data entry form if its still open
            If Not CloseSubject(goStudyDef, True, True, True) Then
                HourglassOff
                Cancel = 1
                Exit Sub
            End If
        End If
        
        
        Call ConfirmDataXfer(Cancel)
        If Cancel = 1 Then
            Exit Sub
        End If
    
    Else
        ' We're being shut down unexpectedly - forms may have already been unloaded
        ' so do a "minimum" close subject
        If IsSubjectOpen Then
            If Not goStudyDef Is Nothing Then
                goStudyDef.RemoveSubject
                goStudyDef.Terminate
                Set goStudyDef = Nothing
            End If
        End If
        
    End If
    
    If UCase(Left(Command, Len(gsAUTO_IMPORT))) <> UCase(gsAUTO_IMPORT) _
    And UCase(Left(Command, Len(gsTRANSFER_DATA))) <> UCase(gsTRANSFER_DATA) Then
        
        ' Only shut down the ALM if it has been started
        If Not goArezzo Is Nothing Then
            goArezzo.Finish
            Set goArezzo = Nothing
        End If
    
    End If
    
    HourglassOff
    
     If UnloadMode = vbFormControlMenu Then
        'TA 07/05/2003:
        'only call this if they clicked close button otherwise we end in a loop
        ExitMACRO
    End If

End Sub

''---------------------------------------------------------------------
' NCJ 10 Jun 03 - Commented out gFindForm because unused
''---------------------------------------------------------------------

'---------------------------------------------------------------------
Private Sub MDIForm_Resize()
'---------------------------------------------------------------------
Dim intCount As Integer
Dim intQuarterWidth As Integer
    
    'SDM 02/12/99   Allow the app to resize to anysize
    On Error Resume Next


    If Me.WindowState <> vbMinimized Then

        For intCount = 0 To StatusFunction.UBound
            StatusFunction(intCount).Width = Me.ScaleWidth / (StatusFunction.UBound + 1)
            If intCount > 0 Then
                StatusFunction(intCount).Left = StatusFunction(intCount - 1).Left + StatusFunction(intCount - 1).Width
            End If
        Next intCount
        
       
        For intCount = 0 To StatusIcons.UBound
            StatusIcons(intCount).Width = Me.ScaleWidth / (StatusIcons.UBound + 1)
            If intCount > 0 Then
                StatusIcons(intCount).Left = StatusIcons(intCount - 1).Left + StatusIcons(intCount - 1).Width
            End If
        Next intCount
        
        
        'second row
        For intCount = 0 To StatusIcons2.UBound
            StatusIcons2(intCount).Width = Me.ScaleWidth / (StatusIcons2.UBound + 1)
            If intCount > 0 Then
                StatusIcons2(intCount).Left = StatusIcons2(intCount - 1).Left + StatusIcons2(intCount - 1).Width
            End If
        Next intCount
        
        
    End If
    
    If Not mofrmMenuLh Is Nothing Then
        'TA 25/09/2002: New UI code
        mofrmHeaderLh.Height = WEB_APP_HEADER_LH_HEIGHT
        mofrmMenuLh.Height = Me.ScaleHeight - (mofrmHeaderLh.Top + mofrmHeaderLh.Height) - WEB_APP_FOOTER_LH_HEIGHT
        mofrmFooterLh.Top = Me.ScaleHeight - WEB_APP_FOOTER_LH_HEIGHT
                
        mofrmMenuLh.Width = mlWebLhWidth
        mofrmFooterLh.Width = mlWebLhWidth
        mofrmHeaderLh.Width = mlWebLhWidth
        
        With mofrmMenuTop
            .Left = 0
            .Height = WEB_APP_MENU_TOP_HEIGHT
            .Width = Me.ScaleWidth
        End With
        
        Call ResizeBottomRightHandForms
    End If
    
End Sub


'---------------------------------------------------------------------
Public Sub ResizeBottomRightHandForms(Optional sFormName As String = "")
'---------------------------------------------------------------------
'resizes and positions following forms
'   Home
'   Schedule
'   eForm Top
'   eFormLH
'   frmSubjectList (if loaded)
'   frmNewSubject (if loaded)
'   frmeFormDataEntry (if loaded)
'   frmViewDiscrepancies (if loaded)
'   frmDataItemResponse (if loaded)

'optionally pass in a formname will result in only that form being resized
'(single form option only available for mofrmBorderTop and frmDataItemREsponse)
'---------------------------------------------------------------------
Dim lTop As Long
Dim lLeft As Long
Dim lHeight As Long
Dim lWidth As Long

Dim lBottomTop As Long
Dim lBottomHeight As Long

    On Error Resume Next
    'ignore erros as this is resizing code

    lTop = mofrmMenuTop.Top + mofrmMenuTop.Height
    lLeft = mofrmMenuLh.Width
    lWidth = Me.ScaleWidth - mofrmMenuLh.Width
    If SplitScreen Then
        lHeight = (frmMenu.ScaleHeight - lTop - WEB_INNER_BORDER) / 2
    Else
        lHeight = (frmMenu.ScaleHeight - lTop - WEB_INNER_BORDER)
    End If

    If sFormName = "" Or sFormName = "frmBorder" Then
        With mofrmBorderTop
            .Top = lTop
            .Left = lLeft
            .Height = lHeight
            .Width = lWidth
        End With
    End If
    
    If SplitScreen Then
        lBottomTop = lHeight + WEB_BORDER_TITLE_HEIGHT + WEB_BORDER_HEIGHT + WEB_INNER_BORDER
        lBottomHeight = (frmMenu.ScaleHeight - lBottomTop - WEB_INNER_BORDER)   'full height
        If sFormName = "" Or sFormName = "frmBorder" Then
            With moFrmBorderBottom
                .Top = lBottomTop
                .Left = lLeft
                .Height = lBottomHeight
                .Width = lWidth
            End With
        End If
    Else
        lBottomTop = lTop
        lBottomHeight = lHeight
    End If
    

    
    lTop = lTop + WEB_BORDER_HEIGHT + WEB_BORDER_TITLE_HEIGHT
    lLeft = lLeft + WEB_BORDER_WIDTH + WEB_INNER_BORDER
    lHeight = lHeight - (WEB_BORDER_HEIGHT) - WEB_BORDER_TITLE_HEIGHT - 2 * WEB_INNER_BORDER
    lWidth = lWidth - (2 * WEB_BORDER_WIDTH) - (2 * WEB_INNER_BORDER)
            
    lBottomTop = lBottomTop + WEB_BORDER_HEIGHT + WEB_BORDER_TITLE_HEIGHT
    lBottomHeight = lBottomHeight - (WEB_BORDER_HEIGHT) - WEB_BORDER_TITLE_HEIGHT - 2 * WEB_INNER_BORDER - 30
            
            
    If sFormName = "" Then
        With mofrmMenuHome
            .Top = lTop
            .Left = lLeft
            .Height = lHeight
            .Width = lWidth
        End With
        
        With mofrmSchedule
            .Top = lTop
            .Left = lLeft
            .Height = lHeight
            .Width = lWidth
        End With
        
        'should we resize schedule?
        'schedule is visible if "Subject" is the last top form opened and eform is not open
        If Not FormIsLoaded(g_DATAENTRY_FORM_NAME) Then
            'eform not open
            If Not gcolTopFormOrder Is Nothing Then
                If gcolTopFormOrder(gcolTopFormOrder.Count) = wfSubject Then
                    'if the subject (schedule or eform is topmost)
                    With mofrmSchedule
                        .Top = lTop
                        .Left = lLeft
                        .Height = lHeight
                        .Width = lWidth
                    End With
                    'execute javascript resize event (ignoring errors)
'                    Debug.Print Timer
                    mofrmSchedule.ExecuteJavaScript.SizeHeaders
'                    Debug.Print Timer
                End If
            End If
        End If
        
        'If FormIsLoaded("frmSubjectList") Then
            With mofrmSubjectList
                .Top = lTop
                .Left = lLeft
                .Height = lHeight
                .Width = lWidth
            End With
        'End If
        
        If FormIsLoaded("frmNewSubject") Then
            With frmNewSubject
                .Top = lTop
                .Left = lLeft
                .Height = lHeight
                .Width = lWidth
            End With
        End If
        
        If FormIsLoaded("frmViewDiscrepancies") Then
            With frmViewDiscrepancies
                .Top = lBottomTop
                .Left = lLeft
                .Height = lBottomHeight
                .Width = lWidth
                'if loaded and the last one browser displayed then bring to top
'                If gLastBrowser = lbMIMessage Then
'                    If SplitScreen Then
'                        .ZOrder
'                        moFrmBorderBottom.SetCaption "MIMessage Browser"
'                    End If
'                End If
            End With
        End If
    End If
    
    If (sFormName = "") Or (sFormName = "frmDataItemResponse") Then
        If FormIsLoaded("frmDataItemResponse") Then
            With frmDataItemResponse
                .Top = lBottomTop
                .Left = lLeft
                .Height = lBottomHeight
                .Width = lWidth
                'if loaded and the last one browser displayed then bring to top
'                If gLastBrowser = lbDataBrowser Then
'                    If SplitScreen Then
'                        .ZOrder
'                        moFrmBorderBottom.SetCaption "Data Browser"
'                    End If
'                End If
            End With
        End If
    End If
    
    If sFormName = "" Then
        'the follwing three have slightly different dimensions
        With mofrmeFormTop
            .Top = lTop
            .Left = lLeft
            .Height = WEB_EFORM_TOP_HEIGHT
            .Width = lWidth

        End With
        
         With mofrmeFormLh
            .Top = mofrmeFormTop.Top + mofrmeFormTop.Height '- 10 'TA we have to do this or you can see between the forms, I don't know why
            .Left = lLeft
            .Height = lHeight - WEB_EFORM_TOP_HEIGHT
            .Width = WEB_EFORM_LH_WIDTH
        End With
        
        If FormIsLoaded(g_DATAENTRY_FORM_NAME) Then
            With frmEFormDataEntry
                .Top = mofrmeFormTop.Top + mofrmeFormTop.Height ' - 10 'TA we have to do this or you can see between the forms, I don't know why
                .Left = mofrmeFormLh.Left + mofrmeFormLh.Width + 10   'TA 16/10/2003 : add a few pixels to avoid the being-able-to-click-on-the-form-behind bug
                .Height = lHeight - WEB_EFORM_TOP_HEIGHT
                .Width = lWidth - mofrmeFormLh.Width
                'TA 27/03/2003 - let the resize work
                mofrmeFormTop.ExecuteJavaScript.fnResize
            End With
        End If
    End If
    
End Sub

'---------------------------------------------------------------------
Public Sub ShowCommsSettings()
'---------------------------------------------------------------------

    If Not IsDataEntryFormLoaded Then
        frmCommunicationConfigurationList.Show vbModal
    End If

End Sub

'---------------------------------------------------------------------
Public Sub CommunicationHistory()
' RS 18/02/2003 - Moved from menu to HTML Panel
'---------------------------------------------------------------------

    If Not IsDataEntryFormLoaded Then
        frmCommunicationHistory.Show vbModal
    End If

End Sub


'---------------------------------------------------------------------
Public Sub Templates()
'---------------------------------------------------------------------
' Ask the user to select a WORD templates and fill it in with this subject's values
' Assume a subject is open
'---------------------------------------------------------------------
Dim mWordApp As New Word.Application
Dim wdFormField As Variant
Dim sResult As String

    On Error GoTo ErrHandler

    ' Can't do it without AREZZO
    If goArezzo Is Nothing Then Exit Sub
    
    CommonDialog1.Flags = cdlOFNExplorer + cdlOFNFileMustExist
    CommonDialog1.Filter = "Word templates|*.dot"
    On Error Resume Next
    ' NCJ 11 Jun 03 - Initialise directory to MACRO "Documents" folder (Bug 1790)
    CommonDialog1.InitDir = gsDOCUMENTS_PATH
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen
    
    If Err.Number <> cdlCancel Then
    
        On Error GoTo ErrHandler
        
        Call HourglassOn
        
        Set mWordApp = New Word.Application
        ' NCJ 10 Jun 03 - Changed Open to Add so we can use the template correctly
        mWordApp.Documents.Add CommonDialog1.FileName
        
        For Each wdFormField In mWordApp.Documents(1).FormFields
        
            'TA 31/10/2000 - additional code to check Arezzo hasnn't returned an error
            sResult = goArezzo.EvaluateExpression(LCase$(wdFormField.Result))
            If Not goArezzo.ResultOK(sResult) Then
                'there is an Arezzo error so the field should be an empty string
                sResult = ""
            End If
            wdFormField.Result = sResult
    
        Next
        
        Call HourglassOff
        ' Show the document
        mWordApp.Visible = True
    End If

    Err.Clear
    Exit Sub

ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Templates", Err.Source) = Retry Then
        Resume
    End If
   
End Sub

'---------------------------------------------------------------------
Private Sub mofrmEFormDataEntry_SubjectLabelChanged()
'---------------------------------------------------------------------
'refresh quicklist
'---------------------------------------------------------------------

    If goUser.CheckPermission(gsFnViewQuickList) Then
        'TA 13/01/2003: ensure left hand panel quick listis updated with the new subject
        Call mofrmMenuLh.ExecuteJavaScript.fnReloadQuickList((GetWinIO.GetDelimitedSubjectList(goUser)))
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mofrmEFormDataEntry_Unload()
'---------------------------------------------------------------------
'This occurs every time the data entry form is unloaded
'---------------------------------------------------------------------

    On Error GoTo Errorlabel
       
    If Not RefreshSchedule Then
        'schedule not open
        
        'no shcedule - must unload subject
        
        'REM 25/04/02 - changed the last boolean parameter to False, i.e. don't unload study on close
        'don't check if eform open as we are in its unload event
        If Not mbSwitchingUser Then
            CloseSubject goStudyDef, True, False, False
        End If
        
    End If
    
        'hide eform borders- moved here from eFormBuilder.Terminate
        mofrmeFormTop.Hide
        mofrmeFormLh.Hide
    
Exit Sub
  
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mofrmEFormDataEntry_Unload", Err.Source) = Retry Then
        Resume
    End If
End Sub


'---------------------------------------------------------------------
Private Sub mofrmSubjectList_SubjectSelected(lStudyId As Long, sSite As String, lSubjectId As Long)
'---------------------------------------------------------------------
' OK clicked on subject form
'---------------------------------------------------------------------
    
    On Error GoTo Errorlabel

    Call SubjectOpen(lStudyId, sSite, lSubjectId)
    
Exit Sub
  
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mofrmSubjectList_SubjectSelected", Err.Source) = Retry Then
        Resume
    End If
    
End Sub


'---------------------------------------------------------------------
Private Sub mofrmNewSubject_Selected(lStudyId As Long, sSite As String)
'---------------------------------------------------------------------
' OK clicked on subject form
'---------------------------------------------------------------------

    On Error GoTo Errorlabel

    Call SubjectOpen(lStudyId, sSite, g_NEW_SUBJECT_ID)
    
    If goUser.CheckPermission(gsFnViewQuickList) Then
        'TA 13/01/2003: ensure left hand panel quick listis updated with the new subject
        Call mofrmMenuLh.ExecuteJavaScript.fnReloadQuickList((GetWinIO.GetDelimitedSubjectList(goUser)))
    End If
    

Exit Sub
  
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mofrmNewSubject_Selected", Err.Source) = Retry Then
        Resume
    End If

    
End Sub


'---------------------------------------------------------------------
Public Sub UserLogOut()
'---------------------------------------------------------------------
' Routine created NCJ 6/3/00
' Log the user out
' NCJ 6/3/00 SR3010 Include confirmatory message
' NCJ 16/3/00 SR3010 Check if data needs saving on Data Entry form first
'---------------------------------------------------------------------
Dim sMsg As String
Dim Cancel As Integer

    On Error GoTo Errorlabel
    
    If Not IsDataEntryFormLoaded Then

        sMsg = "Are you sure you wish to logout?" & vbNewLine
        
        If DialogQuestion(sMsg) = vbYes Then
            Set gOC = Nothing
            Call CloseSubject(goStudyDef, True, True, True)
            'TA 21/05/2003: prompt for data xfer
            Call ConfirmDataXfer(Cancel)
            If Cancel = 1 Then
                Exit Sub
            End If
            Call UnloadAllChildForms(False)
            
            'cant log out for active directory, end app
            If (gbIsActiveDirectoryLogin) Then
                MACROEnd
            Else
                Call Main
            End If
        End If
    
    End If
Exit Sub
  
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "UserLogOut", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Public Sub UserSwitch()
'---------------------------------------------------------------------
' TA 07/01/2003: Switch user - staying on same eform and question
' NCJ 2 Sept 03 - Farmed out some of the code to SwitchTheUser
' so that ChangeData rights are correctly dealt with (Bug 1989)
'---------------------------------------------------------------------
Dim sMsg As String
Dim oUser As MACROUser
Dim lEFITaskId As Long
Dim lResponseTaskId As Long
Dim nRptNo As Integer
Dim bDoSwitch As Boolean

    On Error GoTo Errorlabel
        
    mbSwitchingUser = True
    Set gOC = Nothing
    
    ' Initialise EFI
    lEFITaskId = 0
    bDoSwitch = False
    
    If FormIsLoaded(g_DATAENTRY_FORM_NAME) Then
        sMsg = "Are you sure you wish to change users and log in at the same eForm?"
        If frmEFormDataEntry.SaveNeeded Then
            sMsg = sMsg & vbCrLf & "Changed data for this eForm will be saved."
        End If
        
        If DialogQuestion(sMsg) = vbYes Then
            ' Check we can save & move on OK
            If Not IsDataEntryFormLoaded(False, lEFITaskId, lResponseTaskId, nRptNo) Then
                bDoSwitch = True
            End If
        End If
    Else
        ' No eForm loaded so go ahead with switch
        bDoSwitch = True
    End If
    
    If bDoSwitch Then
        ' Log in a new user
        Set oUser = frmNewLogin.Display(InitializeSecurityADODBConnection, goUser.Database.DatabaseCode, True)
        If Not oUser Is Nothing Then
            Set goUser = oUser
            frmDataItemResponse.Hide
            frmViewDiscrepancies.Hide
            frmNewSubject.Hide
            mofrmSubjectList.Hide
            
            Call SwitchTheUser(lEFITaskId, lResponseTaskId, nRptNo)
            Call RefreshAfterUserChange
    
        Else
            ' The login failed
            ' Just unload rather than exitMACRO
            Unload Me
            MACROEnd
        End If
    End If
    
    mbSwitchingUser = False
    
Exit Sub
  
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "UserSwitch", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub SwitchTheUser(ByVal lEFITaskId As Long, ByVal lResponseTaskId As Long, ByVal nRptNo As Integer)
'---------------------------------------------------------------------
' NCJ 2 Sept 03 - Code extracted from UserSwitch
' Switch to a new user
' Reload eForm if lEFITaskId > 0, and go to specified response/rptno
'---------------------------------------------------------------------
Dim lStudyId As Long
Dim sSite As String
Dim lSubjectId As Long
Dim oStudySite As StudySite
Dim bAllowedThisStudySite As Boolean
Dim bNeedToReload As Boolean

    On Error GoTo Errorlabel
            
    ' Have we got a subject currently loaded?
    ' If not, just reset screen and nothing more to do
    If Not IsSubjectOpen Then
        Call ShowHome
        Exit Sub
    End If
    
    ' We have a subject loaded
    'Store which subject is open
    lStudyId = goStudyDef.StudyId
    sSite = goStudyDef.Subject.Site
    lSubjectId = goStudyDef.Subject.PersonId
    
    'check they have permission to access this study/site
    bAllowedThisStudySite = False
    For Each oStudySite In goUser.GetStudiesSites
        If oStudySite.StudyId = lStudyId And oStudySite.Site = sSite Then
            bAllowedThisStudySite = True
            Exit For
        End If
    Next
    
    If Not bAllowedThisStudySite Then
        'not allowed to open this subject - close subject
        CloseSubject goStudyDef, False, False, True
    Else
        ' Check for a change in read/write permission
        If goUser.CheckPermission(gsFnChangeData) Then
            ' Reload if subject is Read Only
            bNeedToReload = goStudyDef.Subject.ReadOnly
        Else
            ' User can't change data
            ' Reload if subject is not Read Only
            bNeedToReload = Not goStudyDef.Subject.ReadOnly
        End If
        If bNeedToReload Then
            ' Unload and reload subject
            SubjectOpen lStudyId, sSite, lSubjectId
        Else
            ' Just point DEBS and AREZZO at correct user
            goStudyDef.Subject.UserName = goUser.UserName
            goStudyDef.Subject.SetUserProperties goUser.UserNameFull, goUser.UserRole

        End If
        
        ' Do we want to reload the eForm?
        If lEFITaskId > 0 Then
            With mofrmMenuHome

                If frmEFormDataEntry.Display(goUser, goStudyDef.Subject.eFIByTaskId(lEFITaskId), _
                                        .Left + WEB_EFORM_LH_WIDTH, _
                                        .Top + WEB_EFORM_TOP_HEIGHT, _
                                        .Width - WEB_EFORM_LH_WIDTH, _
                                        .Height - WEB_EFORM_TOP_HEIGHT, _
                                        mofrmeFormTop, mofrmeFormLh, _
                                        lResponseTaskId, nRptNo) Then
                    MDIForm_Resize
                    frmEFormDataEntry.SetFocus

                End If
            End With
        End If
        
    End If

Exit Sub
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.SwitchTheUser"

End Sub

'---------------------------------------------------------------------
Public Sub RefreshAfterUserChange()
'---------------------------------------------------------------------
'rest left hand side after user switch
'---------------------------------------------------------------------

     DoEvents
     
    mlWebLhWidth = WEB_LH_WIDTH
    mlWebAppHeaderLhHeight = WEB_APP_HEADER_LH_HEIGHT
     
    'do appheaderlh
     'reset username in top left hand header

     With mofrmMenuTop
         .Top = 0
         .Left = 0
         .Width = Me.ScaleWidth
         .Height = WEB_APP_MENU_TOP_HEIGHT
     End With
     
    Call mofrmMenuTop.Display(wdtHTML, AppMenuTopHTML(goUser), "no")
    mofrmMenuTop.ExecuteJavaScript.fnPageLoaded

     'reset left heand menu

     With mofrmMenuLh
         .Top = mofrmHeaderLh.Top + mofrmHeaderLh.Height
         .Left = 0
         .Width = mlWebLhWidth
         .Height = Me.ScaleHeight - (mofrmHeaderLh.Top + mofrmHeaderLh.Height) - WEB_APP_FOOTER_LH_HEIGHT
     End With
     Call mofrmMenuLh.Display(wdtHTML, AppMenuLhHTML, "auto")
     mofrmMenuLh.ExecuteJavaScript.fnPageLoaded
    
     
     Call LoadLhCombos
     
    'init split screen flag
    SplitScreen = goUser.UserSettings.GetSetting(SETTING_SPLIT_SCREEN, False)
        
    'Reset status bars to how they were when user exited the app.
    picFunctions.Visible = goUser.UserSettings.GetSetting(SETTING_VIEW_FUNCTION_KEYS, False)
    picSymbols(0).Visible = goUser.UserSettings.GetSetting(SETTING_VIEW_SYMBOLS, False)
    picSymbols(1).Visible = goUser.UserSettings.GetSetting(SETTING_VIEW_SYMBOLS, False)
    
    'update discrepancy list
    UpdateDiscCount
    'enable/disable taskteims
    EnableDisableTaskListItems
    
    
    MDIForm_Resize

End Sub


'---------------------------------------------------------------------
Private Sub RefreshTrialDocumentList()
'---------------------------------------------------------------------
' Create a variable to add Menu objects and receive the list of trial documents
Dim mMenu As Menu
Dim mDocumentNumber As Integer
Dim rsReferences As ADODB.Recordset
Dim sSQL As String
  
    On Error GoTo ErrHandler

    sSQL = "SELECT DocumentPath " _
            & " FROM StudyDocument " _
            & " WHERE ClinicalTrialId = " & goStudyDef.StudyId _
            & "   AND VersionId = " & goStudyDef.Version
    
    ' Get the list of trials
    Set rsReferences = New ADODB.Recordset
    rsReferences.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    mDocumentNumber = 1
    
    ' While the record is not the last record, add a ListItem object.
    While Not rsReferences.EOF
    
        
        rsReferences.MoveNext   ' Move to next record.
    Wend
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshTrialDocumentList")
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
Private Sub tmrSystemIdleTimeout_Timer()
'---------------------------------------------------------------------
' when the timer goes off it must be time to lock the system
'  it prompts the user to enter the password or wxit MACRO
' the system is then either closed in a controlled way or resets the timer
' NCJ 17/3/00 - Tidied up and simplified (SR 3015)
'TA 27/04/2000: new timeout handling
' DPH 18/10/2001 - Remove References to RemoveLockOnRecord
'---------------------------------------------------------------------
    
    'new timeout handling
    glSystemIdleTimeoutCount = glSystemIdleTimeoutCount + 1
    If glSystemIdleTimeout = glSystemIdleTimeoutCount Then
        ' set the couter to 0 and disable the timer until the user logs in
        glSystemIdleTimeoutCount = 0
        
        tmrSystemIdleTimeout.Enabled = False
        
        ' NCJ 28 May 03 - Reset dirty data in eForm
        If FormIsLoaded(g_DATAENTRY_FORM_NAME) Then
            Call frmEFormDataEntry.ResetUnvalidatedData
        End If
        
        'display splash - we will prompt when exit AMCRO chosen if data entry form is loaded
        If frmTimeOutSplash.Display(FormIsLoaded(g_DATAENTRY_FORM_NAME)) Then
            'password correctly entered
            tmrSystemIdleTimeout.Enabled = True
        Else
        'exit MACRO chosen
    
        ' unload all forms and exit
        Call UnloadAllForms
        End If

    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub SetupSymbols()
'---------------------------------------------------------------------
Dim str As Variant
Dim nCount As Integer

    nCount = 0
    For Each str In Array(DM30_ICON_INVALID, _
                                    DM30_ICON_WARNING, _
                                    DM30_ICON_OK_WARNING, _
                                    DM30_ICON_INFORM, _
                                    DM30_ICON_OK, _
                                    DM30_ICON_MISSING, _
                                    DM30_ICON_UNOBTAINABLE, _
                                    DM30_ICON_NA, _
                                    DM30_ICON_LOCKED, _
                                    DM30_ICON_FROZEN)
        StatusIcons(nCount).Picture = str
        StatusIcons(nCount).Caption = Array("Invalid", _
                                                "Warning", _
                                                "OK Warning", _
                                                "Inform", _
                                                "OK", _
                                                "Missing", _
                                                "Unobtainable", _
                                                "Not applicable", _
                                                "Locked", _
                                                "Frozen")(nCount)
        
        StatusIcons(nCount).Height = picSymbols(0).ScaleHeight
        nCount = nCount + 1
    Next
    

    nCount = 0
    For Each str In Array(DM30_ICON_CHANGE_COUNTALL, _
                                    DM30_ICON_NOTE, _
                                    DM30_ICON_COMMENT, _
                                    DM30_ICON_RAISED_DISC, _
                                    DM30_ICON_RESPONDED_DISC, _
                                    DM30_ICON_QUERIED_SDV, _
                                    DM30_ICON_PLANNED_SDV, _
                                    DM30_ICON_DONE_SDV)
        StatusIcons2(nCount).Picture = str
        StatusIcons2(nCount).Caption = Array("Previous values", _
                                                "Note", _
                                                "Comment", _
                                                "Raised discrepancy", _
                                                "Responded discrepancy", _
                                                "Queried SDV", _
                                                "Planned SDV", _
                                                "Done SDV")(nCount)
        
        StatusIcons2(nCount).Height = picSymbols(1).ScaleHeight
        nCount = nCount + 1
    Next
'Public Const DM30_ICON_CHANGE_COUNT1 = "DM30_ChangeCount1"
'Public Const DM30_ICON_CHANGE_COUNT2 = "DM30_ChangeCount2"
'Public Const DM30_ICON_CHANGE_COUNT3 = "DM30_ChangeCount3"
'Public Const DM30_ICON_NOTE = "DM30_Note"
'Public Const DM30_ICON_COMMENT = "DM30_Comment"
'Public Const DM30_ICON_NOTE_COMMENT = "DM30_NoteComment"
'Public Const DM30_ICON_RAISED_DISC = "DM30_RaisedDisc"
'Public Const DM30_ICON_RESPONDED_DISC = "DM30_RespondedDisc"
'Public Const DM30_ICON_FROZEN = "DM30_Frozen"
'Public Const DM30_ICON_INFORM = "DM30_Inform"
'Public Const DM30_ICON_LOCKED = "DM30_Locked"
'Public Const DM30_ICON_MISSING = "DM30_Missing"
'Public Const DM30_ICON_NA = "DM30_NA"
'Public Const DM30_ICON_OK = "DM30_OK"
'Public Const DM30_ICON_OK_WARNING = "DM30_OKWarning"
'Public Const DM30_ICON_UNOBTAINABLE = "DM30_Unobtainable"
'Public Const DM30_ICON_WARNING = "DM30_Warning"
'Public Const DM30_ICON_INVALID = "DM30_Invalid"
''TA 21/10/200 New SDV icons
'Public Const DM30_ICON_QUERIED_SDV = "DM30_QueriedSDV"
'Public Const DM30_ICON_PLANNED_SDV = "DM30_PlannedSDV"

    
End Sub


'---------------------------------------------------------------------
Public Function ShowPopUp(sOption As String, Optional sEnabledChecked As String = "") As Long
'---------------------------------------------------------------------
' TA 19/04/2000 - Show a user-defined popup menu
' Input:
'       sOption - "|" delimited string of menu options
'       sEnabledChecked - "|" delimited string of codes to determine item appearance
'                           "*" for disabled, "#" for checked"
' Output:
'       function - item selected (-1 if nothing selected)
'---------------------------------------------------------------------

Dim vOption As Variant
Dim vEnabledChecked As Variant
Dim sStatus As String
Dim i As Long

    'temp work around for optional item
    If sEnabledChecked = "" Then
        sEnabledChecked = sOption
    End If
    
    'set default choice to unspecified
    mlPopUpItem = -1
    'fill string array with choices
    vOption = Split(sOption, "|")
    vEnabledChecked = Split(sEnabledChecked, "|")
    For i = 0 To UBound(vOption)
        If i <> 0 Then
            'not intial menu item so create new one
            Load Me.mnuPopUpSubItem(i)
        End If
        sStatus = vEnabledChecked(i)
        With mnuPopUpSubItem(i)
            'disable if contains '*'
            .Enabled = Not CBool(InStr(1, sStatus, "*"))
            'check if contains '#'
            .Checked = CBool(InStr(1, sStatus, "#"))
            .Caption = vOption(i)
        End With
    Next
    'show menu
    PopupMenu mnuPopUp
    
    'unload controls created at run time
    For i = 1 To UBound(vOption)
        Unload mnuPopUpSubItem(i)
    Next
    
    'return user's choice
    ShowPopUp = mlPopUpItem
    
End Function

'---------------------------------------------------------------------
Public Function ShowPopUpMenu(oMenuItems As clsMenuItems, _
                                Optional ByVal X As Single = -1, Optional ByVal Y As Single = -1) As String
'---------------------------------------------------------------------
' TA 19/04/2000 - Show a user-defined popup menu
' Input: collection of clsMenuItem objects
'
' Output:
'       function - key item selected ("" if nothing selected)
'---------------------------------------------------------------------
Dim i As Long

    For i = 0 To oMenuItems.Count - 1
        If i <> 0 Then
            Load Me.mnuPopUpSubItem(i)
        End If
        With mnuPopUpSubItem(i)
            .Enabled = oMenuItems.Item(i).Enabled
            .Checked = oMenuItems.Item(i).Checked
            .Caption = oMenuItems.Item(i).Caption
        End With
    Next
    
    'set default choice to unspecified
    mlPopUpItem = -1
    
    If X = -1 And Y = -1 Then
        'use default coordinates
        'show menu
        If oMenuItems.DefaultItemIndex = -1 Then
            'show with no default
            PopupMenu mnuPopUp
        Else
            PopupMenu mnuPopUp, , , , mnuPopUpSubItem(oMenuItems.DefaultItemIndex)
        End If
    Else
        'use given coordinates
        'show menu
        If oMenuItems.DefaultItemIndex = -1 Then
            'show with no default
            PopupMenu mnuPopUp, , X, Y
        Else
            PopupMenu mnuPopUp, , X, Y, mnuPopUpSubItem(oMenuItems.DefaultItemIndex)
        End If
    End If
    
    'unload controls created at run time (except 0 element)
    For i = 1 To mnuPopUpSubItem.Count - 1
        Unload mnuPopUpSubItem(i)
    Next
    
    'return user's choice
    If mlPopUpItem = -1 Then
        ShowPopUpMenu = ""
    Else
        ShowPopUpMenu = oMenuItems.Item(mlPopUpItem - 1).Key
    End If
    
End Function

'---------------------------------------------------------------------
Public Sub ShowSubjectList(sStudyName As String, sSite As String, sLabel As String, _
                                    sId As String, sOrderBy As String, sAscend As String)
'---------------------------------------------------------------------
'show the subject list
'---------------------------------------------------------------------

    On Error GoTo Errorlabel
    
   If GUISubjectClose(True, True) Then
'        Set mofrmSubjectList = frmSubjectList
        frmHourglass.Display "Opening subject list", True
        
        If IsSubjectOpen Then
            'close subject
            CloseSubject goStudyDef, True, False, False
        End If

        With mofrmMenuHome
            OpenWinForm wfOpenSubject
            mofrmSubjectList.Display wdtHTML, SubjectListHTML(goUser, sStudyName, sSite, sLabel, sId, sOrderBy, sAscend, 0), "auto"
            mofrmSubjectList.Left = .Left
            mofrmSubjectList.Top = .Top
            mofrmSubjectList.Height = .Height
            mofrmSubjectList.Width = .Width
        End With
        UnloadfrmHourglass
        
    End If
Exit Sub
  
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "ShowSubjectList", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Public Sub SubjectOpen(lStudyId As Long, sSite As String, lSubjectId As Long)
'---------------------------------------------------------------------
' Open existing subject.
' Function returns true if subject successfully open.
'---------------------------------------------------------------------
Dim oEFI As EFormInstance
    
    On Error GoTo Errorlabel
    
    If Not IsDataEntryFormLoaded Then
        HourglassOn
        'hide what's going on by bringing border to the top
        mofrmBorderTop.ZOrder
        If LoadSubject(goStudyDef, goArezzo, gsSubjectToken, lStudyId, sSite, lSubjectId) Then
            GUISubjectOpen
            If goUser.UserSettings.GetSetting(SETTING_SAME_EFORM, False) And lSubjectId <> g_NEW_SUBJECT_ID Then
                'open at the last used form
                
                Set oEFI = goStudyDef.Subject.GetEFIbyAllSubjectsKey(goUser.UserSettings.GetSetting(SETTING_LAST_USED_EFORM, "1|1|1|1"))
                If Not oEFI Is Nothing Then
                    ' NCJ 11 Mar 02 - Include nResponseRptNo
                    ' NCJ 18 Sept 02 - Display may fail
                    With mofrmMenuHome
                        GUISubjectOpen
                        If frmEFormDataEntry.Display(goUser, oEFI, .Left + WEB_EFORM_LH_WIDTH, .Top + WEB_EFORM_TOP_HEIGHT, .Width - WEB_EFORM_LH_WIDTH, _
                                                            .Height - WEB_EFORM_TOP_HEIGHT, _
                                                             mofrmeFormTop, mofrmeFormLh) Then
                            MDIForm_Resize
                        End If
                    End With
                Else
                
                    
                    Call ShowSchedule(goUser, goStudyDef, mofrmSchedule, mofrmeFormTop, mofrmeFormLh)
                 End If
            Else
                Call ShowSchedule(goUser, goStudyDef, mofrmSchedule, mofrmeFormTop, mofrmeFormLh)
            End If
        End If
        'let them see again
        mofrmBorderTop.ZOrder 1
        HourglassOff
    End If
  
Exit Sub
  
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "SubjectOpen", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Public Sub ShowNewSubject()
'---------------------------------------------------------------------
' Open a new trial subject
'---------------------------------------------------------------------
Dim lStudyId As Long
Dim sSite As String
Dim lSubjectId As Long
  
On Error GoTo Errorlabel
    
    If GUISubjectClose(True, True) Then
        HourglassOn
        
        If IsSubjectOpen Then
            'close subject
            CloseSubject goStudyDef, True, False, False
        End If
        
        
        Set mofrmNewSubject = frmNewSubject
        With mofrmMenuHome
            'call with  0 as study id nd "" as site to display all
            OpenWinForm wfNewSubject
            mofrmNewSubject.Display goUser, 0, "", .Top, .Left, .Height, .Width
        End With
        HourglassOff
    End If
    
  
Exit Sub
  
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "ShowNewSubject", Err.Source) = Retry Then
        Resume
    End If


End Sub

Private Function ScheduleOpen() As Boolean
'---------------------------------------------------------------------
' Show the schedule - this can only be called if the schedule has previously been shown
'---------------------------------------------------------------------
    
On Error GoTo Errorlabel

    If GUISubjectClose(True, True) Then
        'display schedule
        GUISubjectOpen
        Call ShowSchedule(goUser, goStudyDef, mofrmSchedule, mofrmeFormTop, mofrmeFormLh)
        ScheduleOpen = True
    Else
        ScheduleOpen = False
    End If
  
Exit Function
  
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "ScheduleOpen", Err.Source) = Retry Then
        Resume
    End If
    
End Function

'---------------------------------------------------------------------
Private Sub GUISubjectOpen()
'---------------------------------------------------------------------
' enable /disable menus etc when subject opened
'---------------------------------------------------------------------
' REVISIONS
' DPH 15/04/2002 - Disallow data transfer when a subject is open
'---------------------------------------------------------------------
Dim sMsg As String

    RefreshTrialDocumentList '


    If goStudyDef.Subject.ReadOnly Then
        sMsg = goStudyDef.Subject.ReadOnlyReason
        ' NCJ 22 Mar 02 - Only show message if we have something to say
        ' (e.g. we don't give a message if user doesn't have Change Data rights)
        If sMsg > "" Then
            sMsg = goStudyDef.Subject.ReadOnlyReason & vbCrLf & vbCrLf & "No further data changes may be made to this subject, although you can view and print the data."
            DialogInformation sMsg
        End If
    Else
        'they can update
        'TA 5/10/01: i'm not sure what the following two lines are for or when they were commented out
        'mnuRDocument(1).Caption = "1"
        'mnuRDocument(2).Caption = "2"
    End If
    
    OpenWinForm wfSubject
    
    'create a reference to the form to catch its events
    Set mofrmEFormDataEntry = frmEFormDataEntry
    
    'enable/disable taskteims
    EnableDisableTaskListItems
    
    
End Sub

'---------------------------------------------------------------------
Public Function GUISubjectClose(bPrompt As Boolean, bCheckEformOpen As Boolean) As Boolean
'---------------------------------------------------------------------
' Close trial subject
' Doesn't remove object from memory but unloads DataEntry and Schedule Form
' If bPrompt is true the user can cancel this action
' If false any changed data will be lost
' note this does not remove the study and subject business objects from memory
'should only be called from close subject
'---------------------------------------------------------------------
' REVISIONS
' DPH 15/04/2002 - Allow data transfer when a subject is closed
'---------------------------------------------------------------------
Dim lStudyId As Long
Dim sSite As String
Dim lSubjectId As Long
Dim beFormLoaded As Boolean


On Error GoTo Errorlabel
    
    If bCheckEformOpen Then
        If IsDataEntryFormLoaded(bPrompt) Then
            GUISubjectClose = False
'EXIT FUNCTION HERE
            Exit Function
        End If
    End If
            
    'close schedule
    Call CloseSchedule
    
    GUISubjectClose = True
    
    'gui changes

    CloseWinForm wfSubject
    
    'lose reference to data entry form
    Set mofrmEFormDataEntry = Nothing
                      
    'enable/disable taskteims
    EnableDisableTaskListItems
                      
Exit Function
  
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmMenu.GUISubjectClose"

        
End Function

'---------------------------------------------------------------------
Public Sub EFIOpen(lStudyId As Long, sSite As String, lSubjectId As Long, _
                        lEFITaskId As Long, lResponseTaskId As Long, _
                        nResponseRptNo As Integer, _
                        sReturnFormName As String, _
                        bFocusToQuestion As Boolean)
'---------------------------------------------------------------------
' Open an eForm instance, loading the study and subject first.
' To be called from frmDiscrepancies and Data Browser.
' NCJ 18 Sept 02 - Deal with Display failing
' MLM 20/01/03: If the requested eForm is a visit eForm, open the first normal form in the visit instead.
'---------------------------------------------------------------------
Dim oEFI As EFormInstance
'MLM 20/01/03:
Dim oVisitEFI As EFormInstance

    On Error GoTo Errorlabel
    
    ' NCJ 30 Jun 03 - Don't show the eForm unless they have View Data permission
    If Not (goUser.CheckPermission(gsFnViewData) Or goUser.CheckPermission(gsFnViewSubjectData)) Then Exit Sub

    If GUISubjectClose(True, True) Then
        'hide what's going on by bringing border to the top
        mofrmBorderTop.ZOrder
        HourglassOn
        If LoadSubject(goStudyDef, goArezzo, gsSubjectToken, lStudyId, sSite, lSubjectId) Then
            Set oEFI = goStudyDef.Subject.eFIByTaskId(lEFITaskId)
            'MLM 20/01/03: now that the subject's loaded, look at the form's purpose
            Set oVisitEFI = oEFI.VisitInstance.VisitEFormInstance
            If Not oVisitEFI Is Nothing Then
                If oEFI.EFormTaskId = oVisitEFI.EFormTaskId Then
                    'the user asked for the visit eform
                    Set oEFI = goStudyDef.Subject.GetFirstVisitForm(oEFI.VisitInstance, Not goStudyDef.Subject.ReadOnly)
                End If
            End If
            Set oVisitEFI = Nothing
            
            'MLM 20/01/03: With the new code above, oEFI could come out as Nothing (if there's a visit date but all other forms in the visit are requested, and the user cannot change data), so:
            If Not oEFI Is Nothing Then
                ' NCJ 11 Mar 02 - Include nResponseRptNo
                ' NCJ 18 Sept 02 - Display may fail
                With mofrmMenuHome
                    GUISubjectOpen
                    If frmEFormDataEntry.Display(goUser, oEFI, .Left + WEB_EFORM_LH_WIDTH, .Top + WEB_EFORM_TOP_HEIGHT, .Width - WEB_EFORM_LH_WIDTH, _
                                                        .Height - WEB_EFORM_TOP_HEIGHT, _
                                                         mofrmeFormTop, mofrmeFormLh, _
                                                        lResponseTaskId, nResponseRptNo) Then
                        MDIForm_Resize
                    End If
                End With
            End If
        End If
        HourglassOff
        'let them see again
        mofrmBorderTop.ZOrder 1
        
    End If
  
Exit Sub
  
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "EFIOpen", Err.Source) = Retry Then
        Resume
    End If
   
End Sub

'---------------------------------------------------------------------
Public Function IsDataEntryFormLoaded(Optional bPromptUser As Boolean = True, _
                                    Optional ByRef lEFITaskId As Long, Optional ByRef lResponseTaskId As Long, Optional ByRef nRptNo As Integer) As Boolean
'---------------------------------------------------------------------
' Is the data entry form open?
' This is the only function that should call frmEFormDataEntry.ClosedSuccessfully
'---------------------------------------------------------------------

    On Error GoTo Errorlabel

    If FormIsLoaded(g_DATAENTRY_FORM_NAME) Then
        IsDataEntryFormLoaded = Not frmEFormDataEntry.ClosedSuccessfully(bPromptUser, lEFITaskId, lResponseTaskId, nRptNo)
    Else
        IsDataEntryFormLoaded = False
    End If
  
Exit Function
  
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "IsDataEntryFormLoaded", Err.Source) = Retry Then
        Resume
    End If

End Function

'---------------------------------------------------------------------
Private Function IsSubjectOpen() As Boolean
'---------------------------------------------------------------------
' Is a study subject currently open?
'---------------------------------------------------------------------

    IsSubjectOpen = False
    If Not goStudyDef Is Nothing Then
        If Not goStudyDef.Subject Is Nothing Then
            IsSubjectOpen = True
        End If
    End If
    
End Function

'---------------------------------------------------------------------
Private Sub RemoteCommsTransfer()
'---------------------------------------------------------------------
' Added in as part of making the transfer screen modal 15/04/2002
'---------------------------------------------------------------------
  
    On Error GoTo ErrHandler
    
    If Me.gTrialOffice.TransferData = 1 Then

        frmDataTransfer.MACROServerDesc = frmMenu.gTrialOffice.TrialOffice
        frmDataTransfer.HTTPAddress = frmMenu.gTrialOffice.HTTPAddress
        frmDataTransfer.Site = frmMenu.gTrialOffice.Site
        ' Show form modally
        frmDataTransfer.Show vbModal
    
        'refresh discrepancy and sdv count in task list
        UpdateDiscCount
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RemoteCommsTransfer")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

' DPH 16/05/2002 - Added command line data transfer with conditional compilation arguements
#If BackgroundXfer Then
'---------------------------------------------------------------------
Private Sub RemoteCommsModelessTransfer()
'---------------------------------------------------------------------
' Created so can run command line data transfer
'---------------------------------------------------------------------
  
    On Error GoTo ErrHandler
    
    If Me.gTrialOffice.TransferData = 1 Then

        frmDataTransfer.MACROServerDesc = frmMenu.gTrialOffice.TrialOffice
        frmDataTransfer.HTTPAddress = frmMenu.gTrialOffice.HTTPAddress
        frmDataTransfer.Site = frmMenu.gTrialOffice.Site
        ' Start modeless transfer
        Call frmDataTransfer.BackGroundDisplay
    
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RemoteCommsTransfer")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub
#End If

'---------------------------------------------------------------------
Public Function ForgottenPassword(sSecurityCon As String, sUserName As String, sPassword As String, ByRef sErrMsg As String) As eDTForgottenPassword
'---------------------------------------------------------------------
'REM 06/12/02
'Used to get new password from server for users on a site who have forgotten their password
'---------------------------------------------------------------------
Dim sDatabaseCode As String
Dim sSiteCode As String
Dim sHTTPAddress As String
Dim nPortNumber As Integer
Dim sIISUserName As String
Dim sIISPassword As String

    sDatabaseCode = frmDBList.Display(sSecurityCon, sUserName, sSiteCode, sHTTPAddress, sIISUserName, sIISPassword, nPortNumber, sErrMsg)
    
    If sErrMsg <> "" Then
        ForgottenPassword = pError
        
    ElseIf sDatabaseCode = "" Then
        ForgottenPassword = pNoDatabases
        sErrMsg = "User " & sUserName & " does not have access to any databases!"
    Else
        ForgottenPassword = frmDataTransfer.ForgottenPassword(sSecurityCon, sDatabaseCode, sUserName, sPassword, sSiteCode, sHTTPAddress, sIISUserName, sIISPassword, nPortNumber, sErrMsg)
    End If
    
End Function

'---------------------------------------------------------------------
Public Sub Resize()
'---------------------------------------------------------------------
'make resize event publically available
    MDIForm_Resize
End Sub

'---------------------------------------------------------------------
Public Function RefreshSearchResults()
'---------------------------------------------------------------------
'TA 09/01/2003
'Simulates the refresh button in the left hand menu being pressed
'---------------------------------------------------------------------

    mofrmMenuLh.ExecuteJavaScript.fnSearch (goUser.CheckPermission(gsFnMonitorDataReviewData))

End Function

'---------------------------------------------------------------------
Public Function ToggleLeftHandPane()
'---------------------------------------------------------------------
'hides/show the left hand task pane
'---------------------------------------------------------------------

    If mlWebLhWidth = WEB_LH_WIDTH Then
        mlWebLhWidth = 0
        mlWebAppHeaderLhHeight = WEB_APP_MENU_TOP_HEIGHT
    Else
        mlWebLhWidth = WEB_LH_WIDTH
        mlWebAppHeaderLhHeight = WEB_APP_HEADER_LH_HEIGHT
    End If
    
    'let resize code sort everything out
    Call Resize

End Function

'---------------------------------------------------------------------
Public Sub ViewOCDiscs()
'---------------------------------------------------------------------
' Show OC Discrepancies form
'---------------------------------------------------------------------

    On Error GoTo Errorlabel

    If gOC Is Nothing Then
        'only set instance on first attempt
        Set gOC = New clsOC
    End If
    gOC.ShowForm
    
Exit Sub
  
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "ViewOCDiscs", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Public Sub UpdateDiscCount()
'---------------------------------------------------------------------

'---------------------------------------------------------------------
'//function updates task list item counter text
'//sName arg is <td> id tag without 'td' prefix
'Function fnSetTaskListItemCounter(sName, sNum)
Dim lRaisedDisc As Long
Dim lRespondedDisc As Long
Dim lClosed As Long
Dim lPlanned As Long
Dim lQueried As Long

Dim oMIMList As MIDataLists

    On Error GoTo Errorlabel

    ' NCJ 25 Jun 03 - MACRO 3.0 Bug 1830 - Sorted out permissions a bit better
   If goUser.CheckPermission(gsFnViewDiscrepancies) Or goUser.CheckPermission(gsFnViewSDV) Then
    
        Set oMIMList = New MIDataLists
        'TA 30/11/2003: use MIMESSAGETRIALNAME not ClinicalTrial.ClinicalTrialId
        Call oMIMList.GetMIMsgStatusCount(goUser.CurrentDBConString, _
                            goUser.DataLists.StudiesSitesWhereSQL("MIMESSAGETRIALNAME", "MIMessageSite"), _
                            lRaisedDisc, lRespondedDisc, lClosed, lPlanned, lQueried)
                            
        If goUser.CheckPermission(gsFnViewDiscrepancies) Then
            mofrmMenuLh.ExecuteJavaScript.fnSetTaskListItemCounter CStr(gsVIEW_RAISED_DISCREPANCIES_MENUID), CStr(lRaisedDisc)
            mofrmMenuLh.ExecuteJavaScript.fnSetTaskListItemCounter CStr(gsVIEW_RESPONDED_DISCREPANCIES_MENUID), CStr(lRespondedDisc)
        End If
        
        If goUser.CheckPermission(gsFnViewSDV) Then
            mofrmMenuLh.ExecuteJavaScript.fnSetTaskListItemCounter CStr(gsVIEW_PLANNED_SDV_MARKS_MENUID), CStr(lPlanned)
        End If
        
    End If

Exit Sub
  
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "UpdateDiscCount", Err.Source) = Retry Then
        Resume
    End If
    

End Sub

'---------------------------------------------------------------------
Public Sub EnableDisableTaskListItems()
'---------------------------------------------------------------------
'REVISIONS:
'REM 12/03/04 - For MACRO Desktop Edition only. Add disable of Create new subject menu item is database is bigger than 1 gig
'---------------------------------------------------------------------
'//function enables/disables tasklist items
'//sName arg is <td> id tag without 'td' prefix
'var aJS = new Array();
'Function fnEnableTaskListItem(sName, bEnable)

'in task list
'Create New Subject
'View subject list
'View raised discrepancies
'View responded discrepancies
'View Oracle Clinical discrepancies
'View planned SDV marks
'Laboratories and normal ranges
'View changes since last session
'Templates
'Register Subject
'View lock/freeze history
'Database lock administration
'Change Password
Dim bSubjectOpen As Boolean
Dim bRegisterSubject As Boolean
Dim bDatabaseLockAdmin As Boolean
Dim bTransferData As Boolean

    On Error GoTo Errorlabel

    bSubjectOpen = IsSubjectOpen
    
    bRegisterSubject = False
    If IsSubjectOpen Then
        If EnableRegistrationMenu Then
             bRegisterSubject = True
        End If
    End If
    
    bDatabaseLockAdmin = Not bSubjectOpen
    
    bTransferData = Not bSubjectOpen
    
    mofrmMenuLh.ExecuteJavaScript.fnEnableTaskListItem (gsREGISTER_SUBJECT_MENUID), (bRegisterSubject)
    mofrmMenuLh.ExecuteJavaScript.fnEnableTaskListItem (gsDB_LOCK_ADMIN_MENUID), (bDatabaseLockAdmin)
    ' NCJ 10 Jun 03 - Enable Templates only if subject is open
    ' NCJ 18 Jun 03 - And if they have Templates permission
    mofrmMenuLh.ExecuteJavaScript.fnEnableTaskListItem (gsTEMPLATES_MENUID), _
                                        (bSubjectOpen And goUser.CheckPermission(gsFnWordTemplates))
    
    If Not goUser.DBIsServer Then
        'if at the site
        mofrmMenuLh.ExecuteJavaScript.fnEnableTaskListItem (gsTRANSFER_DATA_MENUID), (bTransferData)
    End If
    
    'REM 12/03/04 - For MACRO Desktop Edition only. Add disable of Create new subject menu item is database is bigger than 1 GB
    #If DESKTOP = 1 Then
        'check database size is less than 1 GB
        If DatabaseSize > 1024 Then
            mofrmMenuLh.ExecuteJavaScript.fnEnableTaskListItem (gsCREATE_NEW_SUBJECT_MENUID), (False)
        End If
    #End If
    
Exit Sub
  
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "EnableDisableTaskListItems", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Public Function DatabaseSize() As Double
'---------------------------------------------------------------------
'REM 12/03/04
'Function returns the size (in MegaBytes) of the MSDE data file stored in the Database folder for MACRO Desktop Edition
'---------------------------------------------------------------------
Dim oFSO As FileSystemObject
Dim oFile As File
Dim vFileSize As Variant
    
    On Error GoTo ErrLabel

    'create new filesystem object
    Set oFSO = New FileSystemObject
    'get MSDE data file
    Set oFile = oFSO.GetFile(App.Path & "\Database\MACRO30.mdf")
    
    'get size of MSDE
    vFileSize = oFile.SIZE
    
    'convert return (which is in bytes) to megabytes
    DatabaseSize = (vFileSize / 1024) / 1024
    
    Set oFile = Nothing
    Set oFSO = Nothing

Exit Function
ErrLabel:
    DatabaseSize = 0
End Function

'---------------------------------------------------------------------
Private Sub ConfirmDataXfer(ByRef Cancel As Integer)
'---------------------------------------------------------------------
'TA 21/05/2003: Moved here gotm MDIForm_QueryUnload
'---------------------------------------------------------------------

Dim sSQL As String
Dim sMsg As String
Dim rsNewMessages As ADODB.Recordset
Dim rsChangedRecords As ADODB.Recordset
   
Dim bNewLFMessages As Boolean
Dim oLF As LockFreeze


        'if the user has the rights to have changed data and the app is running at a remote site prompt
        'to upload the changed data.
        ' NCJ 22 Jan 03 - Only offer to transfer data if they have Transfer Data permission
        ' (and don't look at ChangeData permission - irrelevant for MIMessages etc.)
        ' AND also check the Lock/Freeze messages
    '    If goUser.CheckPermission(gsFnChangeData) And gblnRemoteSite Then
        If goUser.CheckPermission(gsFnTransferData) And gblnRemoteSite Then
            Set rsChangedRecords = New ADODB.Recordset
            Set rsNewMessages = New ADODB.Recordset
    
            sSQL = "SELECT PersonId FROM TrialSubject WHERE Changed = " & Changed.Changed
            Set rsChangedRecords = New ADODB.Recordset
            rsChangedRecords.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
            sSQL = "SELECT * FROM MIMessage " _
            & "WHERE MIMessageSource = " & TypeOfInstallation.RemoteSite _
            & " AND MIMessageSent = 0"
    
            Set rsNewMessages = New ADODB.Recordset
            rsNewMessages.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
            ' NCJ 22 Jan 03 - Include search for LF messages to transfer
            Set oLF = New LockFreeze
            bNewLFMessages = (oLF.GetMessagesToTransfer(MacroADODBConnection, _
                                TypeOfInstallation.RemoteSite, gTrialOffice.Site).Count > 0)
            Set oLF = Nothing
            
            'TA 21/05/2003: changed message wording
            If bNewLFMessages Or rsChangedRecords.RecordCount > 0 Or rsNewMessages.RecordCount > 0 Then
                sMsg = "Data has changed or new messages have been created since data transfer was last run." & vbCrLf _
                    & " Do you wish to upload the changed data now?"
                Select Case DialogQuestion(sMsg, , True)
                        Case vbYes
                            HourglassOff
                            Cancel = 1
                            Call TransferData
                            Exit Sub
                        Case vbNo
                            'Do nothing just shut down
                        Case vbCancel
                            HourglassOff
                            Cancel = 1
                            Exit Sub
                End Select
            End If
    
            Set rsChangedRecords = Nothing
            Set rsNewMessages = Nothing
        End If

End Sub
