VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMenu 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   ClientHeight    =   7875
   ClientLeft      =   2220
   ClientTop       =   4800
   ClientWidth     =   11880
   Icon            =   "frmMenuStudyDefinition.frx":0000
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tlbMenu 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      Begin VB.PictureBox picSelectedItem 
         Height          =   375
         Left            =   6240
         ScaleHeight     =   315
         ScaleWidth      =   6945
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   7000
         Begin VB.Label lblSelectedItem 
            AutoSize        =   -1  'True
            Caption         =   "lblSelectedItem"
            Height          =   195
            Left            =   0
            TabIndex        =   2
            Top             =   60
            Width           =   1080
         End
      End
   End
   Begin VB.Timer tmrSystemIdleTimeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   600
      Top             =   1680
   End
   Begin MSComctlLib.StatusBar sbrMenu 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   7560
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Key             =   "UserKey"
            Object.ToolTipText     =   "Name of current user"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Text            =   "Role"
            TextSave        =   "Role"
            Key             =   "RoleKey"
            Object.ToolTipText     =   "Name of current users role"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Key             =   "UserDatabase"
            Object.ToolTipText     =   "Name of current database."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   "17/04/2007"
            Object.ToolTipText     =   "Date"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   "15:32"
            Object.ToolTipText     =   "Time"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   7.15654e-38
   End
   Begin MSComctlLib.ImageList imglistSmallIcons 
      Left            =   240
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuFNew 
         Caption         =   "&New..."
      End
      Begin VB.Menu mnuFOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFExport 
         Caption         =   "Export..."
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFArezzoSettings 
         Caption         =   "AREZZO Memory Settings..."
      End
      Begin VB.Menu mnuFS4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFDelete 
         Caption         =   "&Delete Study..."
      End
      Begin VB.Menu mnuFCopy 
         Caption         =   "Copy S&tudy..."
      End
      Begin VB.Menu mnuFArezzoUpdate 
         Caption         =   "AREZZO Update for Clinical Gateway"
      End
      Begin VB.Menu mnuFDL1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDdatabaseLockAdministration 
         Caption         =   "&Database Lock Administration..."
      End
      Begin VB.Menu mnuFS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Begin VB.Menu mnuFPrintCRFPage 
            Caption         =   "&eForm"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuFPrintAllCRFPages 
            Caption         =   "&All eForms"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuFPrintSetup 
         Caption         =   "Print &Setup..."
      End
      Begin VB.Menu mnuFSLockMACRO 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFPageBreaks 
         Caption         =   "View Page &Breaks"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFSPB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLockMACRO 
         Caption         =   "&Stand by"
      End
      Begin VB.Menu mnuFS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuE 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEData 
         Caption         =   "&Question Definition..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuECRF 
         Caption         =   "e&Form..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuECaption 
         Caption         =   "&Caption"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuES1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEQGroups 
         Caption         =   "Question Groups...."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuES3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEUnusedQuestions 
         Caption         =   "Delete &Unused Questions"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEDeleteUnusedQGroups 
         Caption         =   "Delete Unused Question Groups"
      End
   End
   Begin VB.Menu mnuV 
      Caption         =   "&View"
      Begin VB.Menu mnuVStudyDefinition 
         Caption         =   "&Study Details"
      End
      Begin VB.Menu mnuVDataList 
         Caption         =   "&Question List"
      End
      Begin VB.Menu mnuVCRF 
         Caption         =   "&eForms"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuVVisits 
         Caption         =   "&Visits"
      End
      Begin VB.Menu mnuVAREZZO 
         Caption         =   "&AREZZO"
      End
      Begin VB.Menu mnuVReferences 
         Caption         =   "&References"
      End
      Begin VB.Menu mnuVS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVLibrary 
         Caption         =   "&Library"
      End
      Begin VB.Menu mnuVTrialList 
         Caption         =   "List of &Other Studies..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuVS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVDataCRFPage 
         Caption         =   "&Group Question List by eForm"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuVDisplayDataByName 
         Caption         =   "D&isplay Question List by Name"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuVDisplayFormsAlphabetically 
         Caption         =   "Display Question List eForms Alphabetically"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuVS4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVAREZZOReport 
         Caption         =   "AREZZO Terms Report"
      End
   End
   Begin VB.Menu mnuI 
      Caption         =   "&Insert"
      Begin VB.Menu mnuIDataItem 
         Caption         =   "&Question..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuIQuestionGroup 
         Caption         =   "Question Group..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuICRFPage 
         Caption         =   "eFor&m..."
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuIVisit 
         Caption         =   "&Visit..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuR 
      Caption         =   "Fo&rmat"
      Begin VB.Menu mnuRFont 
         Caption         =   "&Font..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRFontColour 
         Caption         =   "Font &Colour..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCRFBackgroundColour 
         Caption         =   "&Background Colour..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRChangeTo 
         Caption         =   "Change &to"
         Enabled         =   0   'False
         Begin VB.Menu mnuRTextBox 
            Caption         =   "&Text Box"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuRFreeText 
            Caption         =   "&Free Text"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuROptionButtons 
            Caption         =   "&Option Buttons"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuROptionBoxes 
            Caption         =   "Option Bo&xes"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuRPopupList 
            Caption         =   "&Drop-down List"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuRAttachment 
            Caption         =   "&Attachment"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuRCalendar 
            Caption         =   "&Calendar"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuRS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRDefaults 
         Caption         =   "&Defaults"
         Begin VB.Menu mnuRDefaultFont 
            Caption         =   "&Font"
         End
         Begin VB.Menu mnuRDefaultFontColour 
            Caption         =   "Font &Colour"
         End
         Begin VB.Menu mnuRDefaultBackgroundColour 
            Caption         =   "&Background Colour"
         End
      End
   End
   Begin VB.Menu mnuO 
      Caption         =   "&Options"
      Begin VB.Menu mnuOAutoCaption 
         Caption         =   "&Auto Caption"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOCRFGrid 
         Caption         =   "eForm &Grid"
         Begin VB.Menu mnuOCRFGridNone 
            Caption         =   "&None"
         End
         Begin VB.Menu mnuOCRFGridSmall 
            Caption         =   "&Small"
         End
         Begin VB.Menu mnuOCRFGridMedium 
            Caption         =   "&Medium"
         End
         Begin VB.Menu mnuOCRFGridLarge 
            Caption         =   "&Large"
         End
         Begin VB.Menu mnuOCRFGridShowGrid 
            Caption         =   "Sho&w Grid"
         End
      End
      Begin VB.Menu mnuOCombinedMovement 
         Caption         =   "&Combined Caption/Control Movement"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutomaticNumbering 
         Caption         =   "Au&tomatic Numbering"
      End
      Begin VB.Menu mnuHI1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHideIcons 
         Caption         =   "&Hide Icons"
      End
      Begin VB.Menu mnuHideRQGIcons 
         Caption         =   "Hide Icons in &RQGs"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuHI2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUseOptionButton 
         Caption         =   "&Use option buttons when number of values in category list is"
         Begin VB.Menu mnuCatValues 
            Caption         =   "1"
            Index           =   0
         End
         Begin VB.Menu mnuCatValues 
            Caption         =   "2"
            Index           =   1
         End
         Begin VB.Menu mnuCatValues 
            Caption         =   "3"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu mnuCatValues 
            Caption         =   "4"
            Index           =   3
         End
         Begin VB.Menu mnuCatValues 
            Caption         =   "5"
            Index           =   4
         End
         Begin VB.Menu mnuCatValues 
            Caption         =   "6"
            Index           =   5
         End
         Begin VB.Menu mnuCatValues 
            Caption         =   "7"
            Index           =   6
         End
         Begin VB.Menu mnuCatValues 
            Caption         =   "8"
            Index           =   7
         End
         Begin VB.Menu mnuCatValues 
            Caption         =   "9"
            Index           =   8
         End
         Begin VB.Menu mnuCatValues 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnuCatValues 
            Caption         =   "&Never use option buttons"
            Index           =   10
         End
         Begin VB.Menu mnuCatValues 
            Caption         =   "&Always use option buttons"
            Index           =   11
         End
      End
      Begin VB.Menu mnuDefaultRFC 
         Caption         =   "Reason &For Change"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuP 
      Caption         =   "&Parameters"
      Visible         =   0   'False
      Begin VB.Menu mnuPStandardFormats 
         Caption         =   "&Standard Formats"
      End
      Begin VB.Menu mnuPTrialPhases 
         Caption         =   "Study &Phases"
      End
      Begin VB.Menu mnuPUnitMaintenance 
         Caption         =   "&Units and Conversion Factors"
      End
      Begin VB.Menu mnuPValidationType 
         Caption         =   "&Validation Type"
      End
      Begin VB.Menu mnuPTrialType 
         Caption         =   "Study &Type"
      End
      Begin VB.Menu mnuPLab 
         Caption         =   "&Normal Ranges and CTC"
      End
   End
   Begin VB.Menu mnuW 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuH 
      Caption         =   "&Help"
      Begin VB.Menu mnuHUserGuide 
         Caption         =   "&User Guide"
      End
      Begin VB.Menu mnuHS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAboutMACRO 
         Caption         =   "&About MACRO"
      End
   End
   Begin VB.Menu mnuFLD 
      Caption         =   "Question"
      Visible         =   0   'False
      Begin VB.Menu mnuFLDEdit 
         Caption         =   "&Edit Question Definition"
      End
      Begin VB.Menu mnuFLDEditCaption 
         Caption         =   "Edit Captio&n"
      End
      Begin VB.Menu mnuFLDHideCaption 
         Caption         =   "&Hide Caption"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFLDEditCRFPage 
         Caption         =   "&Edit eForm"
      End
      Begin VB.Menu mnuFLDEditQGroup 
         Caption         =   "Edit Question Group"
      End
      Begin VB.Menu mnuFLDEditEFG 
         Caption         =   "Edit EForm Group"
      End
      Begin VB.Menu mnuFLDRenumber 
         Caption         =   "Re&number All Items"
      End
      Begin VB.Menu mnuFLDS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFLDDeleteField 
         Caption         =   "&Delete Question"
      End
      Begin VB.Menu mnuFLDDeleteCRFPage 
         Caption         =   "&Delete eForm"
      End
      Begin VB.Menu mnuFLDS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFLDInsertDataDefinition 
         Caption         =   "Insert &Question Definition"
      End
      Begin VB.Menu mnuFLDInsertEFQG 
         Caption         =   "Insert eForm Question &Group"
      End
      Begin VB.Menu mnuFLDS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFLDPaste 
         Caption         =   "Paste Selected Items"
      End
      Begin VB.Menu mnuFLDS5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFLDInsertSpace 
         Caption         =   "&Insert Space"
      End
      Begin VB.Menu mnuFLDRemoveSpace 
         Caption         =   "&Remove Space"
      End
      Begin VB.Menu mnuFLDS6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFLDFont 
         Caption         =   "&Font"
      End
      Begin VB.Menu mnuFLDFontColour 
         Caption         =   "Font &Colour"
      End
      Begin VB.Menu mnuFLDBackgroundColour 
         Caption         =   "Change &Background Colour"
      End
      Begin VB.Menu mnuFLDS4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFLDChangeTo 
         Caption         =   "Change &to"
         Begin VB.Menu mnuFLDTextBox 
            Caption         =   "&Text Box"
         End
         Begin VB.Menu mnuFLDFreeText 
            Caption         =   "F&ree Text"
         End
         Begin VB.Menu mnuFLDOptionButtons 
            Caption         =   "&Option Buttons"
         End
         Begin VB.Menu mnuFLDOptionBoxes 
            Caption         =   "Option Bo&xes"
         End
         Begin VB.Menu mnuFLDPopupList 
            Caption         =   "&Drop-down List"
         End
         Begin VB.Menu mnuFLDAttachment 
            Caption         =   "&Attachment"
         End
         Begin VB.Menu mnuFLDCalendar 
            Caption         =   "&Calendar"
         End
      End
      Begin VB.Menu mnuFLDEPROSep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFLDEPRO 
         Caption         =   "EPRO..."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuPDataList 
      Caption         =   "Data List"
      Visible         =   0   'False
      Begin VB.Menu mnuPDataListEditDataD 
         Caption         =   "&Edit Question Definition"
      End
      Begin VB.Menu mnuPDataListDeleteDataD 
         Caption         =   "&Delete Question Definition"
      End
      Begin VB.Menu mnuPDataListDuplicateDataD 
         Caption         =   "D&uplicate Question Definition"
      End
      Begin VB.Menu mnuPDataListS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPDataListEditForm 
         Caption         =   "&Edit eForm Definition"
      End
      Begin VB.Menu mnuPDataListViewForm 
         Caption         =   "&View eForm"
      End
      Begin VB.Menu mnuPDataListDeleteForm 
         Caption         =   "&Delete eForm"
      End
      Begin VB.Menu mnuPDataListDuplicateForm 
         Caption         =   "D&uplicate eForm"
      End
      Begin VB.Menu mnuPDataListS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPDataListEditQGroup 
         Caption         =   "&Edit Question Group"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPDataListEditEFG 
         Caption         =   "Edit e&Form Group"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPDataListDeleteQGroup 
         Caption         =   "&Delete Question Group"
      End
      Begin VB.Menu mnuPDataListDuplicateQGroup 
         Caption         =   "D&uplicate Question Group"
      End
      Begin VB.Menu mnuPDataListS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPDataListInsertdataD 
         Caption         =   "Insert &Question Definition"
      End
      Begin VB.Menu mnuPDataListInsertForm 
         Caption         =   "Insert e&Form"
      End
      Begin VB.Menu mnuPDatalistInsertQGroup 
         Caption         =   "Insert Question Group"
      End
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Status"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuFStatus 
      Caption         =   "Form Status"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuFStatus2 
      Caption         =   "Form Status 2"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuPStudyVisit 
      Caption         =   "Study Visit"
      Visible         =   0   'False
      Begin VB.Menu mnuPStudyVisitBackgroundColour 
         Caption         =   "Set Background Colour"
      End
      Begin VB.Menu mnuPStudyVisitDeleteVisit 
         Caption         =   "Delete Visit"
      End
      Begin VB.Menu mnuPStudyVisitInsertVisit 
         Caption         =   "Insert Visit"
      End
      Begin VB.Menu mnuPStudyVisitSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPStudyVisitAllowMultipleForms 
         Caption         =   "Allow Multiple eForms"
      End
      Begin VB.Menu mnuPStudyVisitAllowSingleForm 
         Caption         =   "Allow Single eForm"
      End
      Begin VB.Menu mnuPStudyVisitRemoveFormFromVisit 
         Caption         =   "Remove eForm from Visit"
      End
      Begin VB.Menu mnuUseFormForDateValidation 
         Caption         =   "Set as Visit eForm"
      End
      Begin VB.Menu mnuPStudyVisitSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPStudyVisitViewEform 
         Caption         =   "View eForm"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpContainer"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpsubItem 
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

'   Copyright:  InferMed Ltd. 1998-2006. All Rights Reserved
'   File:       frmMenu.frm (frmMenuStudyDefinition.frm)
'   Author:     Andrew Newbigging June 1997
'   Purpose:    Main menu used in MACRO StudyDefinition
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   TA 01/08/2000: previous comments in 2.1 comment archive
'   NCJ 20/12/99    Removed GetTrialStatus and put code inline in Delete
'                   Changed Trial to Study in text messages
'   NCJ 7 Jan 00, SR 2583 - Refresh current CRF after changing default colour & font
'   NCJ 13/1/00     Set gsMACROUserGuidePath to new value in InitialiseMe
'   Mo Morris   17/1/00, changes made to LaunchMTMCUI
'   Mo Morris   18/1/00, SR2744, several hot key changes made to main menu and
'                   right mouse menus
'   MO Morris   27/1/00   Changes made to LaunchMTMCUI
'   Mo Morris   14/2/00 Sections of NewTrial re-written
'   TA 14/3/2000   SR3224  Default grid registry setting set to None
'   TA 29/04/2000   subclassing removed
'   TA 09/05/2000   replaced scree.mousepointer with new hourglass functions
'   TA 09/05/2000  SR2795,SR3288 redid functions to display and hide DataList form
'   TA 01/08/2000 SR3544 : code to keep menu caption consistent
'   WillC 30/8/00 SR3601 I have changed the pop-menu so that it takes into consideration what you wish
'                   to delete and displays the correct caption.
'   TA 10/10/2000 SR3307: print all eForms menu item
'   NCJ 16/10/00    Fix for SR3980 (see RefreshCurrentCRFPage)
'   NCJ 23/10/00 SR3951 - Added "View Page Breaks" menu item
'   TA 16/11/2000 Copied ShowPopup from Data Management and changed all menu captions to title case
'   NCJ 16/2/2001 Added "integrity" checks on study before closing
'   NCJ 12/4/2001 - SR 3963 Delete lock file when closing study
'   NCJ 17 May 01 - Changed to use new ALM4 (Arezzo Logic Module)
'  ATO 12 July 01 - Added CopyStudy Menu item
'   DPH 15/10/2001 - Permission For Deleting Unused Questions Check Added
'   DPH 25/10/2001 - Removed Reports menu item and references in code
'   MACRO 3.0
' NCJ 1 Nov 01 - Changed MACROLOCKBS22 to MACROLOCKBS30
' REM 18/11/01 - Added the load for Repeating Question Groups
' RJCW 26/11/01 call to populate Keywords Dictionary object
' NCJ 27 Nov 01 - Added Public property for QuestionGroups collection
' NCJ 30 Nov 01 - Removed references to global CRFElements
' NCJ/REM 6-10 Dec 01 - Editing Question Groups
' NCJ 18 Dec 01 - Putting eForm Group on eForm now in frmCRFDesign
'               Added new group items to mnuPDatalist
' NCJ 3 Jan 02 - Implemented EditEFG from mnuPDataList
' NCJ 7 Jan 02 - Added Insert QGroup to Data List menu
' ZA 14/02/02 - Mark study loaded by ACM, invalid. In OpenTrial
' NCJ 14/15 Jan 02 - Use gFindQuestionList for data list menu items
'                   Brought forward 2.2 Bug fix to LaunchMTMCUI
' DPH 15/01/2002 - Unload Trial list after copy so refreshes next time - Bug fix from 2.2
' DPH 27/05/2002 - Has study been distributed to a remote site check
' REM 01/03/02 - Added Duplicate QGroup to the question list right click menu
' TA 16/04/2002 - Return whether insert was successful in NewTrial
' DPH 27/05/2002 - If study has been distributed then disallow deletion in Delete routine
' MLM 18/06/02: CBB 2.2.15/17 Deleting a study requires a study lock in delete routine
' ASH 21/06/2002 - Bug 2.2.16 no.11 in OpenTrial
' ASH 21/06/2002 - added confirmation message and mousepointer settings SR 4638
'TA 09/07/02: CBB2.2.19.5 & SR4819 unload StudyVisit form rather than just hide it in HideVisits
' MLM 11/07/02: CBB 2.2.18/3 Rebuild Protocols collection after deleting study.
' ZA 20/08/2002 - open Study only if Arezzo status is updated successfully, if set to 1 in ArezzoUpdateStatus column
' RS 21/08/2002 - Added Routine mnuPopupSubItem_Click (needed to make ShowPopup work)
' RS 06/09/2002 - Added Paste Selected Items Menu (for MultiSelect)
' ZA 09/09/2002 - Added new menu items under Options menu
' ZA 23/09/2002 - Remove call to ProtocolStorage
' ASH 30/10/2002 -Changed submenu text to "Hide Icons in RQGs" under Options menu
' NCJ 6-7 Nov 02 - Added new Hotlink button, and Hotlink handling
' NCJ 17 Jan 03 - Start up the CLM for each study with its own memory settings
' NCJ 31 Jan 03 - Added CheckStudyOK when exiting study
'               Close all windows BEFORE closing a study (to avoid lost_focus problems)
' TA  11/03/2003 - added ellipses to approriate menu items
' NCJ 26 Mar 03 - Changed order of items in Insert menu
' NCJ 28 Mar 03 - Enabled Edit Caption for rt. mouse menu on QuGroup
' MLM 01/04/04: Added "Duplicate eForm" menu item.
' NCJ 2 Apr 03 - Bug 1437 - Do not offer Font options if clicking on a question group
' NCJ 3 Apr 03 - Synchronisation of visit cycles on return from AREZZO (for Bug 870)
' NCJ 8 Apr 03 - Disable "Copy Study" in Library Management (Bug 1467)
' NCJ 23 Apr 03 - Changed Arezzo to AREZZO (Bug 1633)
' NCJ 30 Apr 03 - Fixed bugs in Duplicate eForm
' Mo 16/5/2003  Bug 1577, mnuFPrintCRFPage and mnuFPrintAllCRFPages are now enabled/disabled
'               by the Activate/QueryUnload subroutines of frmCRFDesign. They are also enabled/disabled
'               by the showing/hiding of frmCRFDesign in sub tlbMenu_ButtonClick.
'               mnuFPrintCRFPage and mnuFPrintAllCRFPages are no longer enabled/disabled
'               (by wether an eform has the focus or not) by the ChangeSelectedItem called
'               routines DisableAllMenuItems, EnableMenusForCRFPage and EnableMenusForCRFElements.
' NCJ 29 May 03 - Enable Paste menu item according to contents of paste buffer (BUG 1815)
' NCJ 5-9 Jun 03 - Added "View/AREZZO Terms Report" menu item
' NCJ 25 Nov 04 - Handle invalid lock tokens more elegantly in CloseTrial (Issue 2471)
' NCJ 9 Mar 05 - Issue 2539 - Prefill input box in mnuPDataListDuplicateDataD_Click
' TA 25/04/2005 - Added for menu item for patient diaries (mnuFLDEPRO)- will remain invisible in 3.0
' NCJ 3 Jan 06 - Added Partial Dates property
' NCJ 11 May 06 - Added study/form access modes
' NCJ Jun 06 - MUSD stuff
' NCJ 27 Jun 06 - Issue 2744 - Added "Export..." menu item and SaveExportFile
' NCJ 12 Jul 06 - Added Multi-User switch
' NCJ 5-13 Sept 06 - Fixing MUSD issues
' NCJ 25 Sept 06 - More fixes as result of WBT
' NCJ 28 Sept 06 - New mnuPStudyVisitViewEForm menu item to view eForm from schedule
' NCJ 25 Oct 06 - Tidying up as result of MUSD WBT
'------------------------------------------------------------------------------------'

Option Explicit
Option Compare Binary
Option Base 0

Private mnPersonId As Integer
'Private msMode As String
Private msUpdateMode As String
Private mlClinicalTrialId As Long
Private mnVersionId As Integer
Private msClinicalTrialName As String
Private mnTrialStatus As Integer

'WillC 30/11/99
Private mbAutoCaption As Boolean
Private mbCaptionControlMove As Boolean

Private Const msSELECTED_ITEM = "SEL"

Private mbSystemLocked As Boolean

'TA 19/04/2000 - store popup item selected
Private mlPopUpItem As Long

'REM 20/11/01 - Stores all the question groups
Private moQuestionGroups As QuestionGroups

' NCJ 3 Jan 06 - Partial Dates flag
Private mbAllowPDs As Boolean
' NCJ 12 Jul 06 - Multi User flag
Private mbAllowMU As Boolean

' NCJ 10 May 06 - Study access mode and eForm lock token
Private meStudyAccess As eSDAccessMode
Private msEFormLockToken As String
Private mlLockedEFormID As Long
Private msCacheToken As String
Private msTrialDetails As String
Private mlLastLockFailure As Long

'---------------------------------------------------------------------
Public Property Get StudyAccessMode() As eSDAccessMode
'---------------------------------------------------------------------
' NCJ 10 May 06 - The user's Study access mode
'---------------------------------------------------------------------
    
    StudyAccessMode = meStudyAccess
    
End Property

'---------------------------------------------------------------------
Public Property Get eFormAccessMode() As eSDAccessMode
'---------------------------------------------------------------------
' NCJ 10 May 06 - The user's access mode for the current eForm
' Assume there is one!
'---------------------------------------------------------------------
 
    eFormAccessMode = frmCRFDesign.AccessMode
 
End Property

'---------------------------------------------------------------------
Public Property Get AllowPDs() As Boolean
'---------------------------------------------------------------------
' NCJ 3 Jan 06 - Are we allowed to use Partial Dates?
'---------------------------------------------------------------------
 
    AllowPDs = mbAllowPDs
    
 End Property

'---------------------------------------------------------------------
Public Property Get AllowMU() As Boolean
'---------------------------------------------------------------------
' NCJ 12 Jul 06 - Are we allowed to operate MultiUser mode?
'---------------------------------------------------------------------
 
    AllowMU = mbAllowMU
    
 End Property

'---------------------------------------------------------------------
Public Property Get QuestionGroups() As QuestionGroups
'---------------------------------------------------------------------
' The study's collection of defined Question Groups
'---------------------------------------------------------------------

    Set QuestionGroups = moQuestionGroups

End Property

'---------------------------------------------------------------------
Public Sub RefreshQuestionGroups(Optional ByVal lDataItemId As Long = 0)
'---------------------------------------------------------------------
' Refresh the Question group objects for this study
' Should be called after a question has been deleted
' If lDataItemId is given, only do the refresh if the question belongs to a group
'---------------------------------------------------------------------
Dim oQG As QuestionGroup
Dim bDoRefresh As Boolean

    If lDataItemId = 0 Then
        ' Unconditional refresh
        bDoRefresh = True
    Else
        ' See if the question is in a group
        For Each oQG In moQuestionGroups
            If oQG.QuestionExists(lDataItemId) Then
                bDoRefresh = True
                Exit For
            End If
        Next
    End If
    If bDoRefresh Then
        Set moQuestionGroups = Nothing
        Set moQuestionGroups = New QuestionGroups
        moQuestionGroups.Load Me.ClinicalTrialId, Me.VersionId
    End If
    
    Set oQG = Nothing
    
End Sub

'---------------------------------------------------------------------
Public Property Let CombinedCaptionControlMove(bCaptionControlMove As Boolean)
'---------------------------------------------------------------------
' Toggle the CombinedCaptionControlMove on/off depending on registry setting
'---------------------------------------------------------------------
    mbCaptionControlMove = bCaptionControlMove
    
    Select Case bCaptionControlMove
        Case False
            mbCaptionControlMove = False
            Me.mnuOCombinedMovement.Checked = False
        Case True
            Me.mnuOCombinedMovement.Checked = True
            mbCaptionControlMove = True
    End Select

End Property

'---------------------------------------------------------------------
Public Property Get AutoCaption() As Boolean
'---------------------------------------------------------------------
'To Toggle the AutoCaption on/off depending on registry setting
'---------------------------------------------------------------------
    
    AutoCaption = mbAutoCaption

End Property

'---------------------------------------------------------------------
Public Property Let AutoCaption(bAutoCaption As Boolean)
'---------------------------------------------------------------------
' Toggle the AutoCaption on/off depending on registry setting
'---------------------------------------------------------------------

    mbAutoCaption = bAutoCaption
    
    Select Case bAutoCaption
        Case False
            mbAutoCaption = False
            Me.mnuOAutoCaption.Checked = False
        Case True
            mbAutoCaption = True
            Me.mnuOAutoCaption.Checked = True
    End Select
    
End Property

'---------------------------------------------------------------------
Public Property Get ClinicalTrialId() As Long
'---------------------------------------------------------------------

    ClinicalTrialId = mlClinicalTrialId

End Property

'---------------------------------------------------------------------
Public Property Let ClinicalTrialId(ByVal vClinicalTrialId As Long)
'---------------------------------------------------------------------

    mlClinicalTrialId = vClinicalTrialId

End Property

'---------------------------------------------------------------------
Public Property Get VersionId() As Integer
'---------------------------------------------------------------------

    VersionId = mnVersionId

End Property

'---------------------------------------------------------------------
Public Property Let VersionId(ByVal vVersionId As Integer)
'---------------------------------------------------------------------

    mnVersionId = vVersionId

End Property

'---------------------------------------------------------------------
Public Property Get ClinicalTrialName() As String
'---------------------------------------------------------------------

    ClinicalTrialName = msClinicalTrialName

End Property

'---------------------------------------------------------------------
Public Property Let ClinicalTrialName(ByVal vClinicalTrialName As String)
'---------------------------------------------------------------------

    msClinicalTrialName = vClinicalTrialName
    If IsAppInLibraryMode Then
        Me.Caption = GetApplicationTitle
    Else
        Me.Caption = GetApplicationTitle & " - " & vClinicalTrialName
    End If
    
End Property

'---------------------------------------------------------------------
Public Property Get PersonId() As Integer
'---------------------------------------------------------------------

    PersonId = mnPersonId

End Property

'---------------------------------------------------------------------
Public Property Let PersonId(ByVal vPersonId As Integer)
'---------------------------------------------------------------------

    mnPersonId = vPersonId

End Property

''---------------------------------------------------------------------
'Public Property Get Mode() As String
''---------------------------------------------------------------------
'
'    Mode = msMode
'
'End Property
'
''---------------------------------------------------------------------
'Public Property Let Mode(ByVal vMode As String)
'
'    msMode = vMode
'
'End Property

'---------------------------------------------------------------------
Public Property Get UpdateMode() As String
'---------------------------------------------------------------------

    UpdateMode = msUpdateMode

End Property

'---------------------------------------------------------------------
Public Property Let UpdateMode(tmpMode As String)
'---------------------------------------------------------------------

    msUpdateMode = tmpMode

End Property


'---------------------------------------------------------------------
Public Property Get TrialStatus() As Integer
'---------------------------------------------------------------------

    TrialStatus = mnTrialStatus

End Property

'---------------------------------------------------------------------
Public Property Let TrialStatus(tmpTrialStatus As Integer)
'---------------------------------------------------------------------

    mnTrialStatus = tmpTrialStatus

End Property

'---------------------------------------------------------------------
Public Sub InitialiseMe()
'---------------------------------------------------------------------
' Initialisations specific to Study Definition
' This gets called from Main at startup
' NCJ 15 Sept 1999
'RJCW 26/11/01 call to populate Keywords Dictionary object
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    ' NCJ 22 Feb 00 - Switch CLM saves on
    gbDoCLMSave = True
     
     'RJCW 26/11/01 call to populate Keywords Dictionary object
     Call FetchKeywords

    ' MLM 18/01/02: Commented out gsMACROUserGuidePath as there is now a .chm
    '               Also use IsAppInLibraryMode
    If IsAppInLibraryMode Then
        ' NCJ 10 May 06 - Request access mode of R/W
        OpenTrial 0, "Library", sdReadWrite
    End If
    
    ' NCJ 23/10/00 - Iniialise Page Breaks variable
    gbShowPageBreaksOnly = False
    
    'ZA 10/09/2002 - set up menus under options based on values from MACRO Settings
     LoadMACROSettingMenus
     
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "InitialiseMe")
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
Public Sub CheckUserRights()
'---------------------------------------------------------------------
' NCJ 9 Dec 99
' Do any enabling/disabling of menu options etc.
' according to current user's rights
' This gets called from OpenTrial
' NCJ 10 May 06 - Added Study Access Mode
'---------------------------------------------------------------------

    If goUser.CheckPermission(gsFnCreateStudy) Then
        mnuFNew.Enabled = True
        mnuFCopy.Enabled = True     ' NCJ 21 Jun 06
        tlbMenu.Buttons(gsTRIAL_LABEL).Enabled = True
    Else
        mnuFNew.Enabled = False
        mnuFCopy.Enabled = False     ' NCJ 21 Jun 06
        tlbMenu.Buttons(gsTRIAL_LABEL).Enabled = False
    End If
    
    If goUser.CheckPermission(gsFnDelStudy) And meStudyAccess >= sdReadWrite Then
        mnuFDelete.Enabled = True
    Else
        mnuFDelete.Enabled = False
    End If
    
    mnuFExport.Enabled = (meStudyAccess > sdReadOnly)       ' NCJ 26 Jun 06

    If goUser.CheckPermission(gsFnCreateQuestion) And meStudyAccess >= sdReadWrite Then
        mnuIDataItem.Enabled = True
    Else
        mnuIDataItem.Enabled = False
    End If

    If goUser.CheckPermission(gsFnCreateVisit) And meStudyAccess >= sdReadWrite Then
        mnuIVisit.Enabled = True
    Else
        mnuIVisit.Enabled = False
    End If

    If goUser.CheckPermission(gsFnCreateEForm) And meStudyAccess >= sdReadWrite Then
        mnuICRFPage.Enabled = True
    Else
        mnuICRFPage.Enabled = False
    End If

    If goUser.CheckPermission(gsFnMaintEForm) And meStudyAccess >= sdReadWrite Then
        mnuRDefaultFont.Enabled = True
        mnuRDefaultFontColour.Enabled = True
        mnuRDefaultBackgroundColour.Enabled = True
    Else
        mnuRDefaultFont.Enabled = False
        mnuRDefaultFontColour.Enabled = False
        mnuRDefaultBackgroundColour.Enabled = False
    End If

    If goUser.CheckPermission(gsFnAmendArezzo) And meStudyAccess >= sdReadWrite Then
        tlbMenu.Buttons(gsPROFORMA_LABEL).Enabled = True
        mnuVAREZZO.Enabled = True
    Else
        tlbMenu.Buttons(gsPROFORMA_LABEL).Enabled = False
        mnuVAREZZO.Enabled = False
    End If
    
    If goUser.CheckPermission(gsFnCopyQuestionFromLib) And meStudyAccess >= sdReadWrite Then
        tlbMenu.Buttons(gsLIBRARY_LABEL).Enabled = True
        mnuVLibrary.Enabled = True
    Else
        tlbMenu.Buttons(gsLIBRARY_LABEL).Enabled = False
        mnuVLibrary.Enabled = False
    End If
    
    If goUser.CheckPermission(gsFnEditStudyDetails) Then
        tlbMenu.Buttons(gsTRIAL_LABEL).Enabled = True
        mnuVStudyDefinition.Enabled = True
    Else
        tlbMenu.Buttons(gsTRIAL_LABEL).Enabled = False
        mnuVStudyDefinition.Enabled = False
    End If
    
    If goUser.CheckPermission(gsFnCopyQuestionFromStudy) And meStudyAccess >= sdReadWrite Then
        mnuVTrialList.Enabled = True
    Else
        mnuVTrialList.Enabled = False
    End If
    
    ' NCJ 26/5/00 - MAke GGB Arezzo Update invisible for non-GGB-users
     If goUser.CheckPermission(gsFnGGBArezzoUpdate) And meStudyAccess >= sdReadWrite Then
        mnuFArezzoUpdate.Visible = True
    Else
        mnuFArezzoUpdate.Visible = False
    End If

    ' REM 13/12/01 - Permission for editing and creating question groups
    If goUser.CheckPermission(gsFnMaintainQGroups) And meStudyAccess >= sdReadWrite Then
        mnuEQGroups.Enabled = True
        mnuIQuestionGroup.Enabled = True
    Else
        mnuEQGroups.Enabled = False
        mnuIQuestionGroup.Enabled = False
    End If
    
    ' ASH 16/12/2002 Enabling / disabling lock administration
    If (goUser.CheckPermission(gsFnRemoveOwnLocks) Or goUser.CheckPermission(gsFnRemoveAllLocks)) And _
        msClinicalTrialName = "" Then
        mnuDdatabaseLockAdministration.Enabled = True
    Else
         mnuDdatabaseLockAdministration.Enabled = False
    End If


End Sub

'---------------------------------------------------------------------
Public Sub OpenTrial(ByVal lClinicalTrialId As Long, _
                     ByVal sClinicalTrialName As String, _
                     ByVal enRequestedMode As eSDAccessMode)
'---------------------------------------------------------------------
' NCJ 12 Jan 00 - Check that trial isn't already open in DM
' REM 18/11/01 - Added the load for Repeating Question Groups
' ASH 21/06/2002 Bug 2.2.16 no.11
' ZA 14/02/02 - Mark study loaded by ACM, invalid.
' NCJ 10 May 06 - Added Requested access mode
' NCJ 19 Jun 06 - Added Cache tokens
' NCJ 27 Jun 06 - Added check for permission to open LO
' NCJ 12 Jul 06 - Check "multi user" switch
' NCJ 25 Oct 06 - Check for non-MUSD if requested access is RO
'---------------------------------------------------------------------
Dim rsTrialDetails As ADODB.Recordset
Dim btnX As Button
Dim sLockDetails As String
Dim sMSG As String
Dim enRequiredMode As eSDAccessMode

    On Error GoTo ErrHandler
    
    UpdateMode = gsREAD
    
    'TA 04/07/2001: new locking model
    ' NCJ 6 Jun 06 - New locking model again
    gsStudyToken = ""
    sMSG = ""
    ' Reset last eForm lock failure
    mlLastLockFailure = 0
    ' In MUSD, default to requested access mode, otherwise go for FC
    If mbAllowMU Then
        enRequiredMode = enRequestedMode
    Else
        enRequiredMode = eSDAccessMode.sdFullControl
    End If
    
    ' Assume for now that we'll get what we want
    meStudyAccess = enRequiredMode

    ' Only need locks for RW or FC; need open permission for LO
    If enRequiredMode > sdReadOnly Then
        Select Case enRequiredMode
        Case eSDAccessMode.sdFullControl
            ' Get a global study lock
            gsStudyToken = MACROLOCKBS30.LockStudy(gsADOConnectString, goUser.UserName, lClinicalTrialId)
        Case eSDAccessMode.sdReadWrite
            ' Get a lock for RW access
            gsStudyToken = MACROLOCKBS30.LockStudyRW(gsADOConnectString, goUser.UserName, lClinicalTrialId)
        Case eSDAccessMode.sdLayoutOnly
            ' Check we're allowed to open LO
            ' (Returns "" if it's OK)
            gsStudyToken = MACROLOCKBS30.OpenStudyLO(gsADOConnectString, goUser.UserName, lClinicalTrialId)
        End Select
        Select Case gsStudyToken
        Case MACROLOCKBS30.DBLocked.dblStudy
            sLockDetails = MACROLOCKBS30.LockDetailsStudy(gsADOConnectString, lClinicalTrialId)
            If sLockDetails = "" Then
                sMSG = "Another user has this study open for editing."
            Else
                sMSG = "This study definition is currently being used by " & Split(sLockDetails, "|")(0) & "."
            End If
        Case MACROLOCKBS30.DBLocked.dblSubject
            sLockDetails = MACROLOCKBS30.LockDetailsStudy(gsADOConnectString, lClinicalTrialId)
            If sLockDetails = "" Then
                sMSG = "Another user currently has this study open."
            Else
                sMSG = "User '" & Split(sLockDetails, "|")(0) & "' currently has this study open for data entry."
            End If
        Case Else
            'lock successful
        End Select
        If sMSG > "" Then
            gsStudyToken = ""
            If mbAllowMU Then
                ' We haven't been able to get desired access, so assign RO access
                meStudyAccess = sdReadOnly
                DialogInformation sMSG & vbCrLf & "You may not make any changes to this study."
            Else
                ' Not in Multi User mode, so deny access altogether
                DialogInformation sMSG
' *** EXIT HERE if no study lock
                Exit Sub
            End If
        End If
    End If
   
   ' NCJ 21 Jun 06 - Issue 2745 - Log study opening
    msTrialDetails = " study [" & sClinicalTrialName & "]"
    If mbAllowMU Then
        msTrialDetails = msTrialDetails & " with access mode " & GetAccessModeString(meStudyAccess, mbAllowMU)
    End If
    Call gLog(gsOPEN_TRIAL_SD, "Open" & msTrialDetails)
   
    ' We're OK for editing
    ' NCJ 6 June 06 - Not sure about this UpdateMode...
    UpdateMode = gsUPDATE
    
'    If meStudyAccess < sdReadWrite Then
'        ' Ensure no saving of the CLM file in non-RW modes
'        gbDoCLMSave = False
'    End If
    ' Only allow saving of the CLM file in RW modes
    gbDoCLMSave = (meStudyAccess >= sdReadWrite)
    
    ' NCJ 17 Jan 03 - We start up the CLM for each study with its own memory settings
    Call StartUpCLM(lClinicalTrialId)
    
    '   ATN 26/4/99 SR 730
    '   Check that a proforma definition exists in the protedit database (unless its the library).
    '   If not, then the study can only be used for data entry and cannot be modified.
    '   Mo Morris 3/2/00, changed from Me.ClinicalTrialId to lClinicalTrialId
    If lClinicalTrialId > 0 Then
        ' Pass TrialName rather than ID - NCJ 10/8/99
        ' NCJ 19 Jun 06 - Pass in access mode
        If CheckProteditForTrial(sClinicalTrialName, (meStudyAccess >= sdReadWrite)) = False Then
            ' Unlock if necessary
            If gsStudyToken > "" Then
                On Error Resume Next
                Call MACROLOCKBS30.UnlockStudyRW(gsADOConnectString, gsStudyToken, lClinicalTrialId)
                gsStudyToken = ""
            End If
            Close
            Exit Sub
        End If
    End If
    
    ' Show the toolbar
    tlbMenu.Visible = True
    
    HourglassOn
    
    ' NCJ 19 Jun 06 - Get ourselves an AREZZO token
    msCacheToken = MACROLOCKBS30.CacheAddStudyRow(gsADOConnectString, lClinicalTrialId)
    
    ' open the study definition
    Me.ClinicalTrialId = lClinicalTrialId
    Me.ClinicalTrialName = sClinicalTrialName
    Me.VersionId = gnCurrentVersionId(lClinicalTrialId)
    ' NCJ 10 May 06 - Mode not used any more
'    Me.Mode = gsSTUDY_DEFINITION_MODE
    Me.TrialStatus = gdsTrialDetails(lClinicalTrialId)!statusId
    
    ' Get the 'Single use data items' flag - NCJ 4/10/99
    Set rsTrialDetails = New ADODB.Recordset
    Set rsTrialDetails = gdsStudyDefinition(lClinicalTrialId, Me.VersionId)
    gbSingleUseDataItems = (rsTrialDetails.Fields("SingleUseDataItems") = 1)
    rsTrialDetails.Close
    Set rsTrialDetails = Nothing
    
    ' Call LoadProformaTrial instead - NCJ 10/8/99
    LoadProformaTrial sClinicalTrialName
    
   'ZA 20/08/2002 - check if Study requires Arezzo to be updated and if it is successful
   If meStudyAccess >= sdReadWrite Then
        If UpdateArezzoStatus(lClinicalTrialId) Then
            'carry on
        Else
            ' TODO, shall we call the closetrial and exit this routine or we need to
            'do any more rollback
            CloseTrial
            HourglassOff
            Exit Sub
        End If
    End If
            
    'REM 18/11/01 - added the Question Groups Load
    ' NCJ 10/12/01 - Changed to RefreshQuestionGroups
    Call RefreshQuestionGroups
    
    'set all buttons to Unpressed
    'changed by Mo Morris   7/1/99  SR 654
    For Each btnX In tlbMenu.Buttons
        btnX.Value = tbrUnpressed
    Next
    
    Call EnableMenusWhenTrialOpen
    
    'ASH 21/06/2002 Bug 2.2.16 no.11
    Call EnableUnusedQuestionsMenu(mlClinicalTrialId, mnVersionId)
    'REM 19/07/02
    Call EnableUnusedQGroupsMenu(mlClinicalTrialId, mnVersionId)
    
    If Me.ClinicalTrialId > 0 Then      ' if not library
        Me.Caption = GetApplicationTitle & " - " & Me.ClinicalTrialName
        If mbAllowMU Then
            ' In Multi-user mode, show access
            Me.Caption = Me.Caption & " (" & GetAccessModeString(meStudyAccess, True) & ")"
        Else
            Me.Caption = Me.Caption & " (Update)"
        End If
        ' NCJ 10/12/99 - Menus already enabled/disabled according to user's rights
        mnuP.Visible = False
        
    Else
        ' We're in library mode
        Me.Caption = GetApplicationTitle          '   SR 881
        tlbMenu.Buttons(gsTRIAL_LABEL).Visible = False
        tlbMenu.Buttons(gsDATA_ITEM_LABEL).Visible = True
        tlbMenu.Buttons(gsCRF_PAGE_LABEL).Visible = True
        tlbMenu.Buttons(gsVISIT_LABEL).Visible = False
        tlbMenu.Buttons(gsPROFORMA_LABEL).Visible = False
        tlbMenu.Buttons(gsLIBRARY_LABEL).Visible = False
        tlbMenu.Buttons(gsDOCUMENT_LABEL).Visible = False
        ' DPH 25/10/2001 - Removed Reports menu item and references in code
        'tlbMenu.Buttons(gsREPORT_LABEL).Visible = False

        mnuVStudyDefinition.Enabled = False
        mnuVDataList.Enabled = True
        mnuVCRF.Enabled = True
        mnuVVisits.Enabled = False
        mnuVAREZZO.Enabled = False
        mnuVLibrary.Enabled = False
    '    mnuVReferences.Enabled = True
    '   ATN 1/12/99
    '   Added new menu item to view report definitions, disabled in the library
        mnuVReferences.Enabled = False
        ' DPH 25/10/2001 - Removed Reports menu item and references in code
        'mnuVReports.Enabled = False
        mnuIVisit.Enabled = False
        mnuP.Visible = True
        'mnuFPrintVisit.Visible = False
        
    '   ATN 3/5/99  SR 881
    '   Disable file menu options in the library
        mnuFNew.Enabled = False
        mnuFOpen.Enabled = False
        mnuFClose.Enabled = False
        mnuFDelete.Enabled = False
        mnuFCopy.Enabled = False
        mnuFExport.Enabled = False      ' NCJ 26 Jun 06
        
    '   ATN 5/3/99  SR 878
        mnuRS1.Visible = False
        
        frmCRFDesign.DefaultFontName = gDefaultFontName
        frmCRFDesign.DefaultFontSize = gDefaultFontSize
        frmCRFDesign.DefaultFontBold = gDefaultFontBold
        frmCRFDesign.DefaultFontItalic = gDefaultFontItalic
        frmCRFDesign.DefaultFontColour = gDefaultFontColour
        frmCRFDesign.DefaultCRFColour = gDefaultCRFPageColour
        
        Call ViewQuestions(gsUPDATE, Me.ClinicalTrialId, Me.ClinicalTrialName)
        
        ViewCRF
        
    End If
    
    'ZA 14/02/02 - Mark study loaded by ACM, invalid.
    ' NCJ 19 Jun 06 - Only invalidate Cache for Full Control and R/W users
    If meStudyAccess >= sdReadWrite Then
        MarkStudyAsChanged
    End If
    
    Me.Show
    
    HourglassOff
    
    Exit Sub

ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "OpenTrial", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub EnableMenusWhenTrialOpen()
'---------------------------------------------------------------------
' Enable the main menus when a trial is open
'---------------------------------------------------------------------
    
    ' Do general enabling before user-specific enabling
    mnuV.Enabled = True
    mnuI.Enabled = True
    mnuR.Enabled = True
    mnuO.Enabled = True
    mnuE.Enabled = True
    mnuFClose.Enabled = True
    mnuFPrint.Enabled = True
    
    'Ash added 03/08/2001
    mnuRChangeTo.Enabled = False
    
    ' PN change
    mnuVDisplayDataByName.Checked = GetSetting(App.Title, "Settings", "ViewDataListName") = "-1"
    mnuVDisplayDataByName.Enabled = True
    mnuVDisplayFormsAlphabetically.Checked = GetSetting(App.Title, "Settings", "DataListFormOrderAlphabetic") = "-1"
    mnuVDisplayFormsAlphabetically.Enabled = True
    
    mnuVDataCRFPage.Checked = True
    mnuVDataCRFPage.Enabled = True

    ' This will be made visible or invisible later
    mnuFArezzoUpdate.Enabled = True
    
    ' Now adjust other items according to current user's rights
    Call CheckUserRights

 End Sub

'---------------------------------------------------------------------
Public Sub CloseTrial()
'---------------------------------------------------------------------
' Close the current trial
' NCJ 12/4/01 - Make sure we delete the lock file
' REM 18/11/01 - added the set moQuestionGroups = nothing
' NCJ 31 Jan 03 - Close all windows BEFORE closing study (to avoid lost_focus problems)
' NCJ 25 Nov 04 - Handle lock token erros more elegantly (to avoid losing CLM file)
' NCJ 6 Jun 06 - Release RW and eForm locks
'---------------------------------------------------------------------
Dim oForm As Form
     
     On Error GoTo ErrHandler

    HourglassOn "Closing study"
    
    If Me.ClinicalTrialId > 0 Then
        
        ' NCJ 31 Jan 03 - Must close down windows here BEFORE deleting stuff
        For Each oForm In Forms
            If oForm.Name <> "frmMenu" Then
                If oForm.MDIChild = True Then
                    ' NCJ 14 Jan 02 - Only call HideWindow if frmDataDefinition currently visible
                    If oForm.Name = "frmDataDefinition" And oForm.Visible Then
                        Call oForm.HideWindow
                    End If
                    Unload oForm
                End If
            End If
        Next
        
        CloseProformaTrial Me.ClinicalTrialId, Me.VersionId, Me.ClinicalTrialName
        
        'REM 18/11/01 - set the object to nothing to remove the collection
        Set moQuestionGroups = Nothing
        
        'this will not occur in Library Management
        Me.Caption = GetApplicationTitle            'SR 881
    
    End If
        
    ' Show the toolbar
    tlbMenu.Visible = False
    
    '   ATN 29/4/99 SR 807
    '   Close down the file used to lock out a study from other users
    Close
    
    ' NCJ 6 Jun 06 - Release eForm lock if any
    Call UnlockEForm
    
     'TA 04/07/2001: new locking model
     ' NCJ 6 Jun 06 - New again!
     If gsStudyToken <> "" Then
        ' NCJ 25 Nov 04 - Handle invalid tokens more elegantly (Issue 2471)
        On Error Resume Next
        'if no gsStudyToken then CloseTrial is being called without a corresponding OpenTrial being called first
        '(when frmMenu is loaded)
        ' NCJ 12 Jun 06 - Make relevant Unlock Study call
        If meStudyAccess = sdFullControl Then
            ' Release global lock
            MACROLOCKBS30.UnlockStudy gsADOConnectString, gsStudyToken, Me.ClinicalTrialId
        Else
            ' Release RW lock
            MACROLOCKBS30.UnlockStudyRW gsADOConnectString, gsStudyToken, Me.ClinicalTrialId
        End If
        If Err.Number <> 0 Then
            ' NCJ 25 Nov 04 - Lock token error
            Call DialogWarning("There was a problem releasing the lock when closing this study. " & vbCrLf _
                            & "Another user may have already deleted the lock in the Lock Administration window. " & vbCrLf _
                            & "Please note that deleting the lock for study which is open in SD" & vbCrLf _
                            & "can cause damage to study and subject data.")
        End If
        ' Always set this to empty string for same reason as above
        ' (and to make sure we don't get Lock Token errors twice)
        gsStudyToken = ""
     End If
    
    ' NCJ 19 Jun 06 - Release our Cache token
    If msCacheToken > "" Then Call MACROLOCKBS30.CacheRemoveSubjectRow(gsADOConnectString, msCacheToken)
    msCacheToken = ""
    
    ' NCJ 21 Jun 06 - Issue 2745 - Log study closing
    If msTrialDetails > "" Then Call gLog(gsCLOSE_TRIAL_SD, "Close" & msTrialDetails)
    msTrialDetails = ""
    
    ' Reset to "normal" error handler
    On Error GoTo ErrHandler
    
    Me.ClinicalTrialId = -1
    
    Call DisableMenusIfNoTrial
        
    HourglassOff
    
    Exit Sub

ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "CloseTrial", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub DisableMenusIfNoTrial()
'---------------------------------------------------------------------
' Disable all the menus (except File) if there is no trial loaded
' NCJ 19 Sept 06 - Ensure that File/Delete is re-enabled in MUSD
'---------------------------------------------------------------------
    
    mnuO.Enabled = False
    mnuE.Enabled = False
    mnuV.Enabled = False
    mnuI.Enabled = False
    mnuR.Enabled = False
    
    mnuVDataCRFPage.Checked = False
    mnuP.Visible = False
    ' NCJ - Don't disable individual items
    ' because the whole menu is disabled

    mnuFPrint.Enabled = False
    mnuFClose.Enabled = False
    mnuFExport.Enabled = False  ' NCJ 27 Jun 06
    mnuFArezzoUpdate.Enabled = False
    
    ' NCJ 31 Jan 03 - Moved here from EnableMenusWhenTrialClosed
    mnuDdatabaseLockAdministration.Enabled = True

    ' NCJ 19 Sept 06 - Delete might have been disabled in OpenTrial for RO studies
    If Not goUser Is Nothing Then
        ' During SD start up, goUser might not yet exist
        mnuFDelete.Enabled = goUser.CheckPermission(gsFnDelStudy)
    End If
    
End Sub

'---------------------------------------------------------------------
Public Sub ViewStudyDetails()
'---------------------------------------------------------------------
      
    On Error GoTo ErrHandler
    
    frmStudyDefinition.ClinicalTrialId = Me.ClinicalTrialId
    frmStudyDefinition.VersionId = Me.VersionId
    frmStudyDefinition.ClinicalTrialName = Me.ClinicalTrialName
    
    mnuVStudyDefinition.Checked = True
    tlbMenu.Buttons(gsTRIAL_LABEL).Value = tbrPressed
    frmStudyDefinition.UpdateMode = UpdateMode
    frmStudyDefinition.RefreshTrialDetails

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ViewStudyDetails")
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
Public Sub HideStudyDefinition(Optional bSaveWithoutPrompt As Boolean = True)
'---------------------------------------------------------------------
Dim oForm As Form
     
     On Error GoTo ErrHandler

    For Each oForm In Forms
        If oForm.Name = "frmStudyDefinition" Then
            If oForm.SaveChanges(bSaveWithoutPrompt) Then
                ' PN 22/09/99
                ' replace hide call with unload to ensure that the form
                ' is properly unloaded
                Unload oForm
                mnuVStudyDefinition.Checked = False
                tlbMenu.Buttons(gsTRIAL_LABEL).Value = tbrUnpressed
            Else
                mnuVStudyDefinition.Checked = True
                tlbMenu.Buttons(gsTRIAL_LABEL).Value = tbrPressed
            End If
            Exit For
        End If
    Next oForm

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "HideStudyDefinition")
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
Private Sub ViewVisits()
'---------------------------------------------------------------------
     
     On Error GoTo ErrHandler

    mnuVVisits.Checked = True
    tlbMenu.Buttons(gsVISIT_LABEL).Value = tbrPressed
    
    frmStudyVisits.ClinicalTrialId = Me.ClinicalTrialId
    frmStudyVisits.VersionId = Me.VersionId
    frmStudyVisits.ClinicalTrialName = Me.ClinicalTrialName
    frmStudyVisits.PersonId = Me.PersonId
    'Following line added by Mo Morris 8/1/99 SR 631
    frmStudyVisits.Show
    frmStudyVisits.RefreshStudyVisits

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ViewVisits")
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
Public Sub HideVisits()
'---------------------------------------------------------------------
' hide schedule when visits button is unclicked
'---------------------------------------------------------------------
        
Dim tmpForm As Form
     
     On Error GoTo ErrHandler
    For Each tmpForm In Forms
        If tmpForm.Name = "frmStudyVisits" Then
            tmpForm.WindowState = vbNormal
            'TA 09/07/02: CBB2.2.19.5 & SR4819 unload form rather than just hide it
            ' so that its query_unload event cannot be called later while it is hidden
            Unload tmpForm
        End If
    Next
    
    mnuVVisits.Checked = False
    tlbMenu.Buttons(gsVISIT_LABEL).Value = tbrUnpressed
    
    'Mo Morris 26/2/99 SR 694, dummy call to change menu options and clear the
    'currently selected item
    ChangeSelectedItem "", ""

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "HideVisits")
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
Public Sub ViewCRF(Optional lCRFPageId As Long = 0)
'---------------------------------------------------------------------
'   SDM 22/02/00 Updated to show grid
'   NCJ 6 June 06 - EForm locking and access modes
'---------------------------------------------------------------------
Dim sMSG As String

    On Error GoTo ErrHandler

    HourglassOn
    
    ' NCJ 6 Jun 06 - Lock eForm if necessary
    Call LockEForm(lCRFPageId)
    
    mnuVCRF.Checked = True
    tlbMenu.Buttons(gsCRF_PAGE_LABEL).Value = tbrPressed
    
    frmCRFDesign.ClinicalTrialId = Me.ClinicalTrialId
    frmCRFDesign.VersionId = Me.VersionId
    frmCRFDesign.ClinicalTrialName = Me.ClinicalTrialName
    frmCRFDesign.UpdateMode = gsUPDATE
        
    'SDM 22/02/00 Updated to show grid
    Select Case True
        Case mnuOCRFGridNone.Checked
            frmCRFDesign.Grid = "None"
        Case mnuOCRFGridSmall.Checked
            frmCRFDesign.Grid = "Small"
        Case mnuOCRFGridMedium.Checked
            frmCRFDesign.Grid = "Medium"
        Case mnuOCRFGridLarge.Checked
            frmCRFDesign.Grid = "Large"
    End Select
    frmCRFDesign.GridDisplay = mnuOCRFGridShowGrid.Checked
    frmCRFDesign.DisplayGrid
    
    frmCRFDesign.Show

    Call frmCRFDesign.DisplayCRFPage(lCRFPageId)

    'Changed by Mo Morris 7/1/99 SR 656
    'Make additional toolbar buttons visible and Unpressed
    ' NCJ 10/12/99 - Only show buttons if user has appropriate access
    ' NCJ 6 Nov 02 - Added LINK button
    ' NCJ 11 May 06 - Consider Access Mode too
    If goUser.CheckPermission(gsFnMaintEForm) And frmCRFDesign.AccessMode <> sdReadOnly Then
        tlbMenu.Buttons(gsLINE_LABEL).Visible = True
        tlbMenu.Buttons(gsLINE_LABEL).Value = tbrUnpressed
        tlbMenu.Buttons(gsCOMMENT_LABEL).Visible = True
        tlbMenu.Buttons(gsCOMMENT_LABEL).Value = tbrUnpressed
        tlbMenu.Buttons(gsPICTURE_LABEL).Visible = True
        tlbMenu.Buttons(gsPICTURE_LABEL).Value = tbrUnpressed
        tlbMenu.Buttons(gsLINK_LABEL).Visible = True
        tlbMenu.Buttons(gsLINK_LABEL).Value = tbrUnpressed
    End If
    
    'ASH 29/1/2003 Show eform on top even if schedule open
    frmCRFDesign.ZOrder
    
    HourglassOff
    
    Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ViewCRF")
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
Public Sub LockEForm(lCRFPageId As Long)
'---------------------------------------------------------------------
' NCJ 6 Jun 06 - Attempt to get an eForm lock based on current study access mode.
' The EFormID and its lock token (if any) are stored here in frmMenu.
' NCJ 24 Oct 06 - Try and display user name in message when lock already exists
'---------------------------------------------------------------------
Dim sMSG As String
Dim sLockDetails As String

    On Error GoTo ErrHandler

    If lCRFPageId > 0 And lCRFPageId <> mlLockedEFormID Then
        ' Unlock previous eForm if necessary
        Call UnlockEForm
        ' Default access mode is same as study access mode
        frmCRFDesign.AccessMode = meStudyAccess
        ' Don't need a lock for Read Only or Full Control
        If meStudyAccess = sdLayoutOnly Or meStudyAccess = sdReadWrite Then
            ' Try to get an eForm lock
            msEFormLockToken = MACROLOCKBS30.LockEForm(gsADOConnectString, goUser.UserName, mlClinicalTrialId, lCRFPageId)
            Select Case msEFormLockToken
            Case MACROLOCKBS30.DBLocked.dblStudy
                sLockDetails = MACROLOCKBS30.LockDetailsStudy(gsADOConnectString, mlClinicalTrialId)
                If sLockDetails = "" Then
                    sMSG = "Another user is editing this study."
                Else
                    sMSG = "This study definition is currently being used by " & Split(sLockDetails, "|")(0) & "."
                End If
            Case MACROLOCKBS30.DBLocked.dblEForm
                sLockDetails = MACROLOCKBS30.LockDetailsEForm(gsADOConnectString, mlClinicalTrialId, lCRFPageId)
                If sLockDetails = "" Then
                    sMSG = "Another user is editing this eForm: " & CRFPageCodeFromId(mlClinicalTrialId, lCRFPageId)
                Else
                    sMSG = "The eForm: " & CRFPageCodeFromId(mlClinicalTrialId, lCRFPageId) _
                        & vbCrLf & "is currently being edited by " & Split(sLockDetails, "|")(0) & "."
                End If
            Case MACROLOCKBS30.DBLocked.dblSubject, MACROLOCKBS30.DBLocked.dblEFormInstance
                sLockDetails = MACROLOCKBS30.LockDetailsStudy(gsADOConnectString, mlClinicalTrialId)
                If sLockDetails = "" Then
                    sMSG = "Another user currently has this study open for data entry."
                Else
                    sMSG = "User '" & Split(sLockDetails, "|")(0) & "' currently has this study open for data entry."
                End If
            Case Else
                ' We got the eForm lock - all OK
                mlLockedEFormID = lCRFPageId
                mlLastLockFailure = 0
            End Select
            If sMSG > "" Then
                ' Couldn't get the lock
                msEFormLockToken = ""
                mlLockedEFormID = 0
                If mlLastLockFailure <> lCRFPageId Then
                    Call DialogInformation(sMSG & vbCrLf & "You may not make any changes to this eForm.")
                    ' Remember we've already given the user this message
                    mlLastLockFailure = lCRFPageId
                End If
                frmCRFDesign.AccessMode = sdReadOnly
            End If
        End If
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.LockEForm"
    
End Sub

'---------------------------------------------------------------------
Private Sub UnlockEForm()
'---------------------------------------------------------------------
' NCJ 6 June 06 - Remove current eForm lock, if any
'---------------------------------------------------------------------

    If msEFormLockToken > "" Then
        Call MACROLOCKBS30.UnlockEForm(gsADOConnectString, msEFormLockToken, mlClinicalTrialId, mlLockedEFormID)
        msEFormLockToken = ""
    End If
    mlLockedEFormID = 0
    
End Sub

'---------------------------------------------------------------------
Public Sub HideCRF()
'---------------------------------------------------------------------
Dim tmpForm As Form
     
     On Error GoTo ErrHandler
     
    For Each tmpForm In Forms
        If tmpForm.Name = "frmCRFDesign" Then
            tmpForm.Hide
        End If
    Next
    
    mnuVCRF.Checked = False
    tlbMenu.Buttons(gsCRF_PAGE_LABEL).Value = tbrUnpressed
    
    'Changed by Mo Morris 7/1/99 SR 656
    'Make toolbar buttons that are additional to form design invisible
    tlbMenu.Buttons(gsLINE_LABEL).Visible = False
    tlbMenu.Buttons(gsCOMMENT_LABEL).Visible = False
    tlbMenu.Buttons(gsPICTURE_LABEL).Visible = False
    tlbMenu.Buttons(gsLINK_LABEL).Visible = False
    
    'Mo Morris 26/2/99 SR 694, dummy call to change menu options and clear the
    'currently selected item
    ChangeSelectedItem "", ""

    ' NCJ 6 Jun 06 - Release lock
    Call UnlockEForm
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "HideCRF")
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
Private Sub ViewReferences()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    HourglassOn "Retrieving references..."
    
    frmReferences.Show
    
    mnuVReferences.Checked = True
    tlbMenu.Buttons(gsDOCUMENT_LABEL).Value = tbrPressed
    
    HourglassOff
    
Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ViewReferences")
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
Public Sub HideReferences()
'---------------------------------------------------------------------
     
    On Error GoTo ErrHandler
                
    Unload frmReferences
    
    mnuVReferences.Checked = False
    tlbMenu.Buttons(gsDOCUMENT_LABEL).Value = tbrUnpressed
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "HideReferences")
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
Private Sub DisableAllMenuItems()
'---------------------------------------------------------------------
' Do a blanket disable of all menu items
' prior to enabling the ones we want
' Copied from ChangeSelectedItem, NCJ 9 Dec 99
'---------------------------------------------------------------------
    
    'disable all options
    mnuEDelete.Enabled = False
    mnuEData.Enabled = False
    mnuECRF.Enabled = False
    mnuECaption.Enabled = False
    'Mo 16/5/2003  Bug 1577, mnuFPrintCRFPage and mnuFPrintAllCRFPages no longer disabled here
    'TA 10/10/2000 SR3307 - new menu item
    ' NCJ 23/10/00 SR3951
    mnuFPageBreaks.Enabled = False
    mnuCRFBackgroundColour.Enabled = False
    
    mnuRTextBox.Enabled = False
    mnuRFreeText.Enabled = False
    mnuRPopupList.Enabled = False
    mnuROptionButtons.Enabled = False
    mnuROptionBoxes.Enabled = False
    mnuRCalendar.Enabled = False
    mnuRAttachment.Enabled = False
    
    mnuRFont.Enabled = False
    mnuRFontColour.Enabled = False
   
    
    mnuFLDFreeText.Enabled = False
    mnuFLDTextBox.Enabled = False
    mnuFLDPopupList.Enabled = False
    mnuFLDOptionButtons.Enabled = False
    mnuFLDOptionBoxes.Enabled = False
    mnuFLDCalendar.Enabled = False
    mnuFLDAttachment.Enabled = False
    mnuFLDBackgroundColour.Enabled = False
    
    mnuFLDTextBox.Checked = False
    mnuFLDFreeText.Checked = False
    mnuFLDPopupList.Checked = False
    mnuFLDOptionBoxes.Checked = False
    mnuFLDOptionButtons.Checked = False
    mnuFLDCalendar.Checked = False
    mnuFLDAttachment.Checked = False
    
    mnuFLDEdit.Enabled = False
    mnuFLDEditCaption.Enabled = False
    mnuFLDEdit.Visible = True
    mnuFLDDeleteField.Enabled = False
    mnuFLDDeleteCRFPage.Enabled = False
    mnuFLDEditCRFPage.Enabled = False
    mnuFLDEditEFG.Enabled = False       ' NCJ 3 Dec 01
    mnuFLDEditQGroup.Enabled = False    ' REM 6 Dec 01
    mnuFLDRenumber.Enabled = False
    mnuFLDInsertDataDefinition.Enabled = False
    mnuFLDInsertSpace.Enabled = False
    mnuFLDRemoveSpace.Enabled = False
    mnuFLDPaste.Enabled = False         ' RS 6/9/2002
    mnuFLDInsertEFQG.Enabled = False
    mnuFLDFont.Enabled = False
    mnuFLDFontColour.Enabled = False
    
    mnuRTextBox.Checked = False
    mnuRFreeText.Checked = False
    mnuRPopupList.Checked = False
    mnuROptionBoxes.Checked = False
    mnuROptionButtons.Checked = False
    mnuRCalendar.Checked = False
    mnuRAttachment.Checked = False

    'some options need to be enabled
    mnuRChangeTo.Visible = True
    mnuFLDChangeTo.Visible = True

End Sub

'---------------------------------------------------------------------
Public Sub ChangeSelectedItem(ByVal vType As String, _
                              ByVal vName As String, _
                              Optional ByVal bIsGroupCaption As Boolean = True)
'---------------------------------------------------------------------
'Mo Morris  26/2/99 SR 694
'The main conceptual change is that ChangeSelectedItem is now called to
'unselect items when frmCRFDesign, frmStudyVisits, frmDataList are closed
'via calls from HideDatalist, HideCRF and HideVisits.
'
'The disabling section at the start of the sub has been expanded to incorporate
'the unsetting of all options.
'
'mnuFLDPicture and mnuRPicture removed
' NCJ 9 Dec 99 - Moved menu disabling code to DisableAllMenuItems
' NCJ 2 Apr 03 - Added optional bIsGroupCaption argument (for a group, says whether we're on its Caption)
'---------------------------------------------------------------------
Dim nCRFElementID As Integer
Dim oCRFElement As CRFElement

     On Error GoTo ErrHandler
    
    Call DisableAllMenuItems

    ' Display the currently selected item or clear it
    If vName = "" Then
        lblSelectedItem.Caption = ""
        lblSelectedItem.ToolTipText = ""
        picSelectedItem.Visible = False
    Else
        lblSelectedItem.Caption = vName
        'WillC 25/2/2000 SR2876 show the entire selected item in the tool tip
        lblSelectedItem.ToolTipText = vName
        picSelectedItem.Visible = True
    End If
    
    ' Don't do anything if nothing selected
    If vType = "" Then
        Exit Sub
    End If
    
    ' Sort out what we've got
    If InStr(vType, gsCRF_ELEMENT_LABEL) > 0 Then
        ' It's a CRFElement
        nCRFElementID = Mid(vType, Len(gsCRF_ELEMENT_LABEL) + 1)
        Set oCRFElement = frmCRFDesign.CRFElementById(nCRFElementID)
        ' NCJ 2 Apr 03 - Added bIsGroupCaption argument
        Call EnableMenusForCRFElement(oCRFElement, bIsGroupCaption)
        
    ElseIf vType = gsCRF_PAGE_LABEL Then
        ' It's a CRFPage
        EnableMenusForCRFPage
        
    ElseIf vType = gsDATA_ITEM_LABEL Then
        Call EnableMenusForDataItem
        
    ElseIf vType = gsVISIT_LABEL Then
        Call EnableMenusForVisit
    
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ChangeSelectedItem")
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
Private Sub EnableMenusForVisit()
'---------------------------------------------------------------------
' A visit has been selected
' Enable the menu options for it
'---------------------------------------------------------------------

    If goUser.CheckPermission(gsFnDelVisit) And meStudyAccess = sdFullControl Then
        mnuEDelete.Enabled = True
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub EnableMenusForDataItem()
'---------------------------------------------------------------------
' A data item has been selected
' Enable the menu options for it
'---------------------------------------------------------------------
    
    If goUser.CheckPermission(gsFnDelQuestion) And meStudyAccess = sdFullControl Then
        mnuEDelete.Enabled = True
    End If
    
    If goUser.CheckPermission(gsFnAmendQuestion) Then
        mnuEData.Enabled = True
    End If

End Sub

'---------------------------------------------------------------------
Private Sub EnableMenusForCRFPage()
'---------------------------------------------------------------------
' CRF page has been selected
' Enable the menu options for it
' NCJ 10 May 06 - Consider access mode
'---------------------------------------------------------------------
Dim bCanEditeForm As Boolean

    bCanEditeForm = (Me.eFormAccessMode <> sdReadOnly)
    
    'TA 10/10/2000 SR3307 - new menu item
    'Mo 16/5/2003  Bug 1577, mnuFPrintCRFPage and mnuFPrintAllCRFPages no longer enabled here
    mnuFLDChangeTo.Enabled = False
    
    'Ash added 02/08/2001 new
    mnuRChangeTo.Enabled = False
    
    ' NCJ 23/10/00 SR3951
    mnuFPageBreaks.Enabled = True

    ' Access rights check - NCJ 9 Dec 99
    ' NCJ 10 May 06 - Access mode also checked
    If goUser.CheckPermission(gsFnMaintEForm) Then
        mnuCRFBackgroundColour.Enabled = bCanEditeForm
        mnuECRF.Enabled = True
        mnuFLDEditCRFPage.Enabled = True
        If bCanEditeForm Then
            mnuFLDEditCRFPage.Caption = "Edit eForm"
        Else
            ' They can only view (not edit) the dialog
            mnuFLDEditCRFPage.Caption = "View eForm definition"
        End If
        mnuFLDRenumber.Enabled = bCanEditeForm
        mnuFLDBackgroundColour.Enabled = bCanEditeForm
        mnuFLDInsertSpace.Enabled = bCanEditeForm
        mnuFLDRemoveSpace.Enabled = bCanEditeForm
        ' NCJ 29 May 03 - Enable Paste according to contents of paste buffer
        mnuFLDPaste.Enabled = frmCRFDesign.PasteBufferAvailable And bCanEditeForm
    End If
    
    ' NCJ 15 Jun 06 - Check study access
    If meStudyAccess = sdFullControl Then
        If goUser.CheckPermission(gsFnDelEForm) Then
            mnuEDelete.Enabled = True
            mnuFLDDeleteCRFPage.Enabled = True
        End If
    End If
    If goUser.CheckPermission(gsFnCreateQuestion) Then
        If Me.StudyAccessMode >= sdReadWrite Then
            mnuFLDInsertDataDefinition.Enabled = True
        End If
        If Me.eFormAccessMode >= sdLayoutOnly Then
            mnuFLDInsertEFQG.Enabled = True
        End If
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub EnableMenusForCRFElement(oCRFElement As CRFElement, ByVal bIsGroupCaption As Boolean)
'---------------------------------------------------------------------
' A CRF Element has been selected
' Enable the menus for it
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    ' Set the check marks (not dependent on user access rights)
    If oCRFElement.DataItemId > 0 Then
        Select Case oCRFElement.ControlType
        Case gn_TEXT_BOX
            mnuFLDTextBox.Checked = True
            mnuRTextBox.Checked = True
        Case gn_RICH_TEXT_BOX
            mnuFLDFreeText.Checked = True
            mnuRFreeText.Checked = True
        Case gn_POPUP_LIST
            mnuFLDPopupList.Checked = True
            mnuRPopupList.Checked = True
        Case gn_PUSH_BUTTONS
            mnuFLDOptionBoxes.Checked = True
            mnuROptionBoxes.Checked = True
        Case gn_OPTION_BUTTONS
            mnuFLDOptionButtons.Checked = True
            mnuROptionButtons.Checked = True
        Case gn_CALENDAR
            mnuFLDCalendar.Checked = True
            mnuRCalendar.Checked = True
        Case gn_ATTACHMENT
            mnuFLDAttachment.Checked = True
            mnuRAttachment.Checked = True
        End Select
    End If
    
    'REM 07/03/02 - Renaming the delete field according to CRFElement type
    If oCRFElement.QGroupID > 0 Or oCRFElement.OwnerQGroupID > 0 Then
        mnuFLDDeleteField.Caption = "&Remove EForm Group"
    ElseIf (oCRFElement.DataItemId > 0) Then
        mnuFLDDeleteField.Caption = "&Remove Question"
    ElseIf oCRFElement.ControlType = gn_COMMENT Then
        mnuFLDDeleteField.Caption = "&Delete Comment"
    ElseIf oCRFElement.ControlType = gn_LINE Then
       mnuFLDDeleteField.Caption = "&Delete Line"
    ' It 's a picture
    ElseIf oCRFElement.ControlType = gn_PICTURE Then
        mnuFLDDeleteField.Caption = "&Delete Picture"
    ElseIf oCRFElement.ControlType = gn_HOTLINK Then
        mnuFLDDeleteField.Caption = "&Delete Hotlink"
    End If
    
    ' NCJ 14 Jun 06 - Check eform Access mode too
    If goUser.CheckPermission(gsFnMaintEForm) And Me.eFormAccessMode > sdReadOnly Then
        mnuFLDDeleteField.Enabled = True
    End If

    ' NCJ 23/10/00 - Enable CRFPage printing if they're on a CRFElement
    ' (because they must be on a CRFPage!)
    'Mo 16/5/2003  Bug 1577, mnuFPrintCRFPage and mnuFPrintAllCRFPages no longer enabled here
    mnuFPageBreaks.Enabled = True

    ' NCJ 3 Dec 01 - Allow eForm Group editing
    ' REM 13/12/01 - check permission for user
    If oCRFElement.QGroupID > 0 Or oCRFElement.OwnerQGroupID > 0 Then
        mnuFLDEditQGroup.Enabled = True
        mnuFLDEditEFG.Enabled = True
        ' NB Editing RQG defn. depends on study access mode
        If goUser.CheckPermission(gsFnMaintainQGroups) And (Me.StudyAccessMode >= sdReadWrite) Then
            mnuFLDEditQGroup.Caption = "Edit Question Group"
        Else
            mnuFLDEditQGroup.Caption = "View Question Group"
        End If
        If (Me.eFormAccessMode = sdReadOnly) Then
            mnuFLDEditEFG.Caption = "View EForm Group"
        Else
            mnuFLDEditEFG.Caption = "Edit EForm Group"
        End If
        
        'REM 06/03/02 - Enabled font for question group captions
        ' NCJ 2 Apr 03 - Do not enable for a question group unless it's the caption
        ' NCJ 5 Sept 06 - Check the eForm access mode too
        If (Me.eFormAccessMode >= sdReadWrite) Then
            If (oCRFElement.QGroupID = 0 Or (oCRFElement.QGroupID > 0 And bIsGroupCaption)) Then
                mnuFLDFont.Enabled = True
                mnuFLDFontColour.Enabled = True
                mnuRFont.Enabled = True
                mnuRFontColour.Enabled = True
            End If
        End If
        
        ' NCJ 28 Mar 03 - Enable "Edit Caption" for qu grps.
        If oCRFElement.QGroupID > 0 Then
            mnuFLDEditCaption.Enabled = (Me.eFormAccessMode > sdReadOnly)
            mnuFLDEditCaption.Caption = "&Edit Caption"
        End If
        
    End If
    
    ' NCJ 11 May 06 - Consider eForm/study access mode too, and allow question viewing rather than editing
    If (oCRFElement.DataItemId > 0) Then
        mnuFLDEdit.Enabled = True
        ' NB Editing qu. defn. depends on study access mode
        If (Not goUser.CheckPermission(gsFnAmendQuestion)) Or _
         (Me.StudyAccessMode = sdReadOnly) Or (Me.StudyAccessMode = sdLayoutOnly) Then
            mnuFLDEdit.Caption = "View Question Definition"
        Else
            mnuFLDEdit.Caption = "Edit Question Definition"
        End If
    End If
    
    ' Allow hotlink viewing
    If oCRFElement.ControlType = gn_HOTLINK Then
        mnuECaption.Enabled = True
        If (Not goUser.CheckPermission(gsFnAmendQuestion)) Or (Me.eFormAccessMode = sdReadOnly) Then
            mnuFLDEditCaption.Caption = "View Hotlink"
            mnuFLDEditCaption.Enabled = True
        Else
            mnuFLDEditCaption.Caption = "&Edit Hotlink"
        End If
    End If
    
    ' If the user can't edit the question there's no more to do
    If Not goUser.CheckPermission(gsFnAmendQuestion) Or Me.eFormAccessMode = sdReadOnly Then
        Exit Sub
    End If
    
    If (oCRFElement.DataItemId > 0) Then
        ' It's a question field
        mnuEData.Enabled = True
        mnuECaption.Enabled = True
        
        mnuFLDEdit.Enabled = True
        mnuFLDEditCaption.Enabled = True
        mnuFLDHideCaption.Enabled = True
        mnuFLDFont.Enabled = True
        mnuFLDFontColour.Enabled = True
        
        mnuRFont.Enabled = True
        mnuRFontColour.Enabled = True
        
       'ash added 03/08/2001
        mnuRChangeTo.Enabled = True
       
       
        mnuFLDChangeTo.Enabled = True
    
        Select Case oCRFElement.DataItemType
        Case DataType.Text  'Text
            mnuFLDTextBox.Enabled = True
            mnuRTextBox.Enabled = True
            mnuFLDFreeText.Enabled = True
            mnuRFreeText.Enabled = True
        Case DataType.Category  'Category
            mnuFLDTextBox.Enabled = True
            mnuRTextBox.Enabled = True
            mnuFLDPopupList.Enabled = True
            mnuRPopupList.Enabled = True
            mnuFLDOptionButtons.Enabled = True
            mnuROptionButtons.Enabled = True
            mnuFLDOptionBoxes.Enabled = True
            mnuROptionBoxes.Enabled = True
        Case DataType.IntegerData, DataType.Real  'Integers/Real numbers
            mnuFLDTextBox.Enabled = True
            mnuRTextBox.Enabled = True
        Case DataType.Date  'Date
            mnuFLDTextBox.Enabled = True
            mnuRTextBox.Enabled = True
            mnuFLDCalendar.Enabled = True
            mnuRCalendar.Enabled = True
        Case DataType.Multimedia  'Multimedia
            mnuFLDAttachment.Enabled = True
            mnuRAttachment.Enabled = True
        End Select
        
        mnuFLDEditCaption.Caption = "&Edit Caption"

        
    ' It's a comment or hotlink
    ElseIf oCRFElement.ControlType = gn_COMMENT Or oCRFElement.ControlType = gn_HOTLINK Then
        mnuECaption.Enabled = True
        If oCRFElement.ControlType = gn_COMMENT Then
            mnuFLDEditCaption.Caption = "&Edit Comment"
        Else
            mnuFLDEditCaption.Caption = "&Edit Hotlink"
        End If
        mnuFLDEditCaption.Enabled = True
        mnuRFont.Enabled = True
        mnuRFontColour.Enabled = True
        mnuFLDFont.Enabled = True
        mnuFLDFontColour.Enabled = True
        
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "EnableMenusForCRFElement")
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
Private Sub MDIForm_Load()
'---------------------------------------------------------------------
Dim oButton As Button
Dim sRegGridSize As String
Dim bRegAutoCaption As Boolean
Dim bRegCombinedCaptionControlMove As Boolean

     On Error GoTo ErrHandler
   
    With imglistSmallIcons.ListImages
        .Add , gsTRIAL_LABEL, LoadResPicture(gsTRIAL_LABEL, vbResIcon)
        .Add , gsDATA_ITEM_LABEL, LoadResPicture(gsDATA_ITEM_LABEL, vbResIcon)
        .Add , gsQGROUP_LABEL, LoadResPicture(gsQGROUP_LABEL, vbResIcon)
        .Add , gsDATA_COMMENT_LABEL, LoadResPicture(gsDATA_COMMENT_LABEL, vbResIcon)
        .Add , gsLINE_LABEL, LoadResPicture(gsLINE_LABEL, vbResIcon)
        .Add , gsCOMMENT_LABEL, LoadResPicture(gsCOMMENT_LABEL, vbResIcon)
        .Add , gsPICTURE_LABEL, LoadResPicture(gsPICTURE_LABEL, vbResIcon)
        .Add , gsCRF_PAGE_LABEL, LoadResPicture(gsCRF_PAGE_LABEL, vbResIcon)
        .Add , gsVISIT_LABEL, LoadResPicture(gsVISIT_LABEL, vbResIcon)
        .Add , gsPROFORMA_LABEL, LoadResPicture(gsPROFORMA_LABEL, vbResIcon)
        .Add , gsLIBRARY_LABEL, LoadResPicture(gsLIBRARY_LABEL, vbResIcon)
        ' DPH 25/10/2001 - Removed Reports menu item and references in code
        '.Add , gsREPORT_LABEL, LoadResPicture(gsREPORT_LABEL, vbResIcon)
        .Add , gsDOCUMENT_LABEL, LoadResPicture(gsDOCUMENT_LABEL, vbResIcon)
        ' NCJ 6 Nov 02 - Added Link icon for Hotlink (its image is stored in frmSDImages)
        .Add , gsLINK_LABEL, frmSDImages.imgList.ListImages(gsLINK_LABEL).Picture
    End With
    
    'tlbMenu.Height = 495
    ' Add buttons to toolbar
    
    ' Hide menu bar while loading
    tlbMenu.Visible = False
    tlbMenu.ImageList = imglistSmallIcons
    
    Set oButton = tlbMenu.Buttons.Add(, , , tbrSeparator)
    Set oButton = tlbMenu.Buttons.Add(, , , tbrSeparator)
    Set oButton = tlbMenu.Buttons.Add(, gsTRIAL_LABEL, "Study", tbrCheck, gsTRIAL_LABEL)
    oButton.ToolTipText = "Study details"
    Set oButton = tlbMenu.Buttons.Add(, gsDATA_ITEM_LABEL, "Questions", tbrCheck, gsDATA_ITEM_LABEL)
    oButton.ToolTipText = "List of questions"
    Set oButton = tlbMenu.Buttons.Add(, gsCRF_PAGE_LABEL, "eForms", tbrCheck, gsCRF_PAGE_LABEL)
    oButton.ToolTipText = "eForms"
    Set oButton = tlbMenu.Buttons.Add(, gsVISIT_LABEL, "Visits", tbrCheck, gsVISIT_LABEL)
    oButton.ToolTipText = "Schedule of eforms and visits"
    Set oButton = tlbMenu.Buttons.Add(, gsDOCUMENT_LABEL, "References", tbrCheck, gsDOCUMENT_LABEL)
    oButton.ToolTipText = "Reference documents"
    ' DPH 25/10/2001 - Removed Reports menu item and references in code
    Set oButton = tlbMenu.Buttons.Add(, gsPROFORMA_LABEL, "AREZZO", tbrCheck, gsPROFORMA_LABEL)
    oButton.ToolTipText = "AREZZO Composer"
    Set oButton = tlbMenu.Buttons.Add(, , , tbrSeparator)
    Set oButton = tlbMenu.Buttons.Add(, gsLIBRARY_LABEL, "Library", tbrCheck, gsLIBRARY_LABEL)
    oButton.ToolTipText = "Library"
    Set oButton = tlbMenu.Buttons.Add(, , , tbrSeparator)
    'SDM 28/01/00 SR2051
    Set oButton = tlbMenu.Buttons.Add(, gsLINE_LABEL, "Line", tbrCheck, gsLINE_LABEL)
    oButton.ToolTipText = "Add a line to the eform"
    oButton.Visible = False
    'SDM 28/01/00 SR2051
    Set oButton = tlbMenu.Buttons.Add(, gsCOMMENT_LABEL, "Comment", tbrCheck, gsCOMMENT_LABEL)
    oButton.ToolTipText = "Add a comment to the eform"
    oButton.Visible = False
    'SDM 28/01/00 SR2051
    Set oButton = tlbMenu.Buttons.Add(, gsPICTURE_LABEL, "Picture", tbrCheck, gsPICTURE_LABEL)
    oButton.ToolTipText = "Add a picture to the eform"
    oButton.Visible = False
    
    ' NCJ 6 Nov 02 - Hotlink button
    Set oButton = tlbMenu.Buttons.Add(, gsLINK_LABEL, "Hotlink", tbrCheck, gsLINK_LABEL)
    oButton.ToolTipText = "Add a hotlink to the eform"
    oButton.Visible = False

    Set oButton = tlbMenu.Buttons.Add(, msSELECTED_ITEM, , tbrPlaceholder)
    
    picSelectedItem.Left = oButton.Left + (tlbMenu.ButtonWidth * 4) + 400
    picSelectedItem.Top = (tlbMenu.Height - picSelectedItem.Height) / 2

    
    ' NCJ 3 Jan 06 - Check whether we're using Partial Dates
    mbAllowPDs = (LCase(GetMACROSetting(MACRO_SETTING_PARTIAL_DATES, "no")) = "yes")
    
    ' NCJ 12 Jul 06 - Check whether we're using Multi User
    mbAllowMU = (LCase(GetMACROSetting(MACRO_SETTING_MUSD, "no")) = "yes")

    Me.Caption = GetApplicationTitle

    'TA 10/10/2000 SR 3958: hide menu items that do not apply to Study Defintion
    ' Need to reactivate LibraryData and PhaseValidation when we've done the reports
    'mnuFPrintUnitConversion.Visible = IsAppInLibraryMode
    'mnuFPrintLibraryData.Visible = False
    'mnuFPrintPhaseValidation.Visible = False
        
    '   ATN 1/4/99
    '   Call CloseTrial to ensure that menu options are disabled
    Call CloseTrial
    
    'TAA 14/3/00 SR 3224 provide default setting of None
    sRegGridSize = GetSetting("IMedMACRO", "Options", "GridSize", "None")
    Select Case sRegGridSize
        Case "None"
            Me.mnuOCRFGridNone.Checked = True
            frmCRFDesign.Grid = "None"
        Case "Small"
            Me.mnuOCRFGridSmall.Checked = True
            frmCRFDesign.Grid = "Small"
        Case "Medium"
            Me.mnuOCRFGridMedium.Checked = True
            frmCRFDesign.Grid = "Medium"
        Case "Large"
            Me.mnuOCRFGridLarge.Checked = True
            frmCRFDesign.Grid = "Large"
        Case Else
            'TA 27/11/2000: case else added to correct a setting changed manually in RegEdit
            'defualt is "None"
            Me.mnuOCRFGridNone.Checked = True
            frmCRFDesign.Grid = "None"
    End Select
    Me.mnuOCRFGridShowGrid.Checked = GetSetting("IMedMACRO", "Options", "GridDisplay", False)
    bRegAutoCaption = GetSetting("ImedMACRO", "Options", "AutoCaption", True)
    Select Case bRegAutoCaption
            Case False
                Me.AutoCaption = False
            Case True
                Me.AutoCaption = True
    End Select
    
    bRegCombinedCaptionControlMove = GetSetting("ImedMACRO", "Options", "CombinedCaptionControlMove", True)
    Select Case bRegCombinedCaptionControlMove
            Case False
                Me.CombinedCaptionControlMove = False
            Case True
                Me.CombinedCaptionControlMove = True
    End Select
    
    Set oButton = Nothing

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "MDIForm_Load")
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
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---------------------------------------------------------------------
' NCJ 16/2/01 - Check their study's OK before letting them get out
'---------------------------------------------------------------------

    If UnloadMode = vbFormControlMenu Then
        ' They clicked the close box
        If Not CheckStudyOK Then
            ' Stop the app closing
            Cancel = 1
        End If
    End If

End Sub

'---------------------------------------------------------------------
Private Sub MDIForm_Unload(Cancel As Integer)
'---------------------------------------------------------------------
Dim sRegGridSize As String

    ' Forget any errors during the unload - NCJ 9 Nov 99
    On Error Resume Next
    
    SaveSetting "ImedMACRO", "Options", "AutoCaption", Me.mnuOAutoCaption.Checked
    SaveSetting "ImedMACRO", "Options", "CombinedCaptionControlMove", Me.mnuOCombinedMovement.Checked
       
  ' WillC 22/2/2000 SR2597 Use the SaveSetting and if statement to save the grid size from session
  ' to session.
    If Me.mnuOCRFGridNone.Checked = True Then
             sRegGridSize = "None"
        ElseIf Me.mnuOCRFGridSmall.Checked = True Then
             sRegGridSize = "Small"
        ElseIf Me.mnuOCRFGridMedium.Checked = True Then
             sRegGridSize = "Medium"
        ElseIf Me.mnuOCRFGridLarge.Checked = True Then
             sRegGridSize = "Large"
    End If

    SaveSetting "ImedMACRO", "Options", "GridSize", sRegGridSize
    SaveSetting "ImedMACRO", "Options", "GridDisplay", Me.mnuOCRFGridShowGrid.Checked

    Call CloseTrial
    
    ' Close any files
    Close
    
    Unload Me

    ' Shut down the CLM
    Call ShutDownCLM
    
    Call ExitMACRO
    
End Sub

'---------------------------------------------------------------------
Public Function IsAppInLibraryMode() As Boolean
'---------------------------------------------------------------------
' PN 17/09/99 - created
' this function will return true if the application is being run in library mode
' or false if not
'---------------------------------------------------------------------
      
     On Error GoTo ErrHandler
   If LCase$(Command) = "library" Then
        IsAppInLibraryMode = True
    Else
        IsAppInLibraryMode = False
    End If

Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "IsAppInLibraryMode")
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
Private Sub mnuAutomaticNumbering_Click()
'---------------------------------------------------------------------
'ZA 09/09/2002 - sets the automatic numbering on/off based on user selection
'---------------------------------------------------------------------
    'toggle between checked and unchecked state
    mnuAutomaticNumbering.Checked = Not mnuAutomaticNumbering.Checked
    
    If mnuAutomaticNumbering.Checked Then
        SetMACROSetting gs_AUTOMATIC_NUMBERING, CStr(eAutoNumbering.NumberingOn)
    Else
        SetMACROSetting gs_AUTOMATIC_NUMBERING, CStr(eAutoNumbering.NumberingOff)
    End If
End Sub

'------------------------------------------------------------------------------
Private Sub mnuCatValues_Click(Index As Integer)
'------------------------------------------------------------------------------
'Decided whether to display option buttons or drop down list
'------------------------------------------------------------------------------

    'Clears all previous selection
    ClearOptionChoices
    'Select the current user option
    mnuCatValues(Index).Checked = Not mnuCatValues(Index).Checked
    
    Select Case Index
        Case 0 To 8
            'User selection to use option buttons between 1 to 8
            gnUseOptionButton = Val(mnuCatValues(Index).Caption)
        Case gn_NEVER_USE_OPTION_MENU
            'User selected not to use option buttons
            gnUseOptionButton = gn_NEVER_USE_OPTION_BUTTONS
        Case gn_ALWAYS_USE_OPTION_MENU
            'User selected always use option buttons
            gnUseOptionButton = gn_ALWAYS_USE_OPTION_BUTTONS
    End Select
    
    SetMACROSetting gs_USE_OPTION_BUTTONS, CStr(gnUseOptionButton)
    
End Sub

'--------------------------------------------------------------------
Private Sub mnuDdatabaseLockAdministration_Click()
'--------------------------------------------------------------------
'ASH 16/12/2002
'--------------------------------------------------------------------

    Call frmLocksAdmin.Display(goUser, goUser.DatabaseCode)

End Sub

'---------------------------------------------------------------------
Private Sub mnuDefaultRFC_Click()
'---------------------------------------------------------------------
Dim enDefaultRFC As eRFCDefault
    
    'Toggle the check state of this menu item
    mnuDefaultRFC.Checked = Not mnuDefaultRFC.Checked
    
    'save user selection for Default RFC
    If mnuDefaultRFC.Checked Then
        enDefaultRFC = eRFCDefault.RFCDefaultOn
    Else
        enDefaultRFC = eRFCDefault.RFCDefaultOff
    End If
    
    'save this key/value for future use
    SetMACROSetting gs_DEFAULT_RFC, CStr(enDefaultRFC)
End Sub

'---------------------------------------------------------------------
Private Sub mnuECaption_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").EditCaption
End Sub

'---------------------------------------------------------------------
Private Sub mnuECRF_Click()
'---------------------------------------------------------------------
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").ShowCRFPageDefinition

End Sub

'---------------------------------------------------------------------
Private Sub mnuEData_Click()
'---------------------------------------------------------------------
' Edit the currently selected question
'---------------------------------------------------------------------
Dim oElement As CRFElement

    If ActiveForm.Name = "frmCRFDesign" Then
        Set oElement = frmCRFDesign.CurrentCRFElement
        If oElement.DataItemId > 0 Then
            frmDataDefinition.ShowDataDefinition Me.ClinicalTrialId, _
                                            Me.VersionId, _
                                            Me.ClinicalTrialName, _
                                            oElement.DataItemId, _
                                            Me.StudyAccessMode
        End If
        Set oElement = Nothing
    ElseIf ActiveForm.Name = "frmDataList" Then
        ' NCJ 15 Jan 02 - Check it's not read-only
        If ActiveForm.UpdateMode <> gsREAD And ActiveForm.SelectedDataItemId > 0 Then
            frmDataDefinition.ShowDataDefinition Me.ClinicalTrialId, _
                                            Me.VersionId, _
                                            Me.ClinicalTrialName, _
                                            ActiveForm.SelectedDataItemId, _
                                            Me.StudyAccessMode
        End If
    End If
End Sub

'---------------------------------------------------------------------
Private Sub mnuEDelete_Click()
'---------------------------------------------------------------------

    If ActiveForm.Name = "frmStudyVisits" Then
        ActiveForm.Delete
    ElseIf ActiveForm.Name = "frmDataList" Then
        ' NCJ 15 Jan 02 - Check it's not read-only
        If ActiveForm.UpdateMode <> gsREAD Then
            ActiveForm.DeleteDataItem
        End If
    ElseIf ActiveForm.Name = "frmCRFDesign" Then
        ActiveForm.Delete
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuEDeleteUnusedQGroups_Click()
'---------------------------------------------------------------------
'REM 19/07/02
'Delete Unused Question Groups
'---------------------------------------------------------------------

    If DialogQuestion("Are you sure you want to delete all unused question groups?") = vbYes Then
        HourglassOn
        'delete all unused QGroups
        moQuestionGroups.DeleteUnusedQGroups
        'refreshes question list / treeview
        Call RefreshQuestionLists(Me.ClinicalTrialId)
        Call MarkStudyAsChanged         ' NCJ 5 Sept 06
        HourglassOff
        DialogInformation "All unused question groups deleted"
    End If
    
End Sub

'-----------------------------------------------------------------------
Public Sub EditQuestionGroup(oQGroup As QuestionGroup)
'-----------------------------------------------------------------------
' Show the Edit window for the given question group
' NCJ 21 Sept 06 - Don't allow editing of a group on a locked eForm
'-----------------------------------------------------------------------
Dim colDataItemIDs As Collection
Dim oForm As Form
Dim bCanEdit As Boolean

    On Error GoTo ErrLabel
    
    ' NCJ 21 Sept 06 - Check for locked forms
    If IsRQGOnLockedForm(oQGroup.QGroupID) Then
        Call DialogInformation("You may not edit this question group because" & vbCrLf _
                            & "it is on an eForm currently being edited by another user.")
        bCanEdit = False
    Else
        bCanEdit = (meStudyAccess >= sdReadWrite)
    End If
    
    Set colDataItemIDs = DataItemIdsNotAllowedInQGroup(Me.ClinicalTrialId, _
                                        Me.VersionId, oQGroup.QGroupID)
    'Display selected QGroup for editing
    If frmQGroupDefinition.Display(oQGroup, colDataItemIDs, bCanEdit) Then
        ' Update current eForm if any
        Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
        If Not oForm Is Nothing Then
            'REM 31/01/02 - added check for CRFPage
            If frmCRFDesign.CRFPageId > 0 Then
                oForm.RefreshEFormGroup (oQGroup.QGroupID)
            End If
        End If
        Call RefreshQuestionLists(Me.ClinicalTrialId)
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, Err.Source, Err.Description & "|frmmenu.EditQuestionGroup"
End Sub


Private Sub mnuFLDEPRO_Click()
'-----------------------------------------------------------------------
' TA 25/04/2005 - Added for Patient diaries - will remain invisible in 3.0
'-----------------------------------------------------------------------
    gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign").EPRO
End Sub

'-----------------------------------------------------------------------
Private Sub mnuEQGroups_Click()
'-----------------------------------------------------------------------
' REM 06/12/01 - Edit a Question Group
'-----------------------------------------------------------------------
Dim oQGroup As QuestionGroup
Dim oForm As Form
Dim nReturn As Integer
Dim colDataItemIDs As Collection

    'Select which QGroup to edit
    nReturn = frmSelectQGroup.Display(moQuestionGroups, oQGroup, True)
    
    If nReturn = EditQGroup.Edit Then
    
        Call EditQuestionGroup(oQGroup)
        
    ElseIf nReturn = EditQGroup.Delete Then
            ' Update current eForm if any
            Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
            If Not oForm Is Nothing Then
                oForm.RefreshMe
            End If
            ' Update current question list if any
            Call RefreshQuestionLists(Me.ClinicalTrialId)
            ' NCJ 19 Jun 06
            Call MarkStudyAsChanged
            ' Go back and let them do another one if they want
            Call mnuEQGroups_Click

    End If
    
End Sub

'-----------------------------------------------------------------------
Private Sub mnuEUnusedQuestions_Click()
'-----------------------------------------------------------------------
'calls routine that deletes unused questions for a particular study
'-----------------------------------------------------------------------
'Revisions:
'ASH 21/06/2002 added confirmation message and mousepointer settings SR 4638
'-----------------------------------------------------------------------

    If DialogQuestion("Are you sure you want to delete all unused questions?") = vbYes Then
        HourglassOn
        Call DeleteUnusedQuestions(Me.ClinicalTrialId, Me.VersionId)
        Call MarkStudyAsChanged         ' NCJ 5 Sept 06
        HourglassOff
    End If

End Sub

'------------------------------------------------------------------------
Private Sub mnuFArezzoSettings_Click()
'------------------------------------------------------------------------
' ASH 22/11/2002 Loads AREZZO Settings form
' NCJ 24 Jun 03 - Ensure it works properly in the Library
'------------------------------------------------------------------------
Dim sMSG As String
    
    On Error GoTo ErrHandler
     
    If RefreshIsNeeded Then Exit Sub
    
    If Me.ClinicalTrialId > 0 Then
        sMSG = "Any changes made to the settings will not be effective" & vbCrLf & _
                "until the study is closed and re-opened"
        ' Study already open
        Call frmArezzoSettings.Display(Me.ClinicalTrialId, _
                Me.ClinicalTrialName, meStudyAccess, sMSG)
    Else
        ' NCJ 24 Jun 03 - Don't close trial etc. if no study loaded! (Bug 1871)
'        If Not CheckStudyOK Then
'            Exit Sub
'        End If
'        CloseTrial

        ' NCJ 9 May 06 - Use mode = gsREAD to select a trial
        frmTrialList.Mode = gsREAD
        frmTrialList.Show vbModal
        ' NCJ 13 Jun 06 - Always allow them to change settings from here
        If frmTrialList.SelectedClinicalTrialId > 0 Then
            Call frmArezzoSettings.Display(frmTrialList.SelectedClinicalTrialId, _
                    frmTrialList.SelectedClinicalTrialName, sdFullControl)
        End If
   End If
        
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuFArezzoSettings_Click")
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
Private Sub mnuFArezzoUpdate_Click()
'---------------------------------------------------------------------
' NCJ 26/5/00 - Update Arezzo after a GGB import
' Assume there is a trial open
'---------------------------------------------------------------------

    Call UpdateAllArezzo(Me.ClinicalTrialId)
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuFClose_Click()
'---------------------------------------------------------------------
' Close trial
' NCJ 16/2/01 - Perform study check before closing
'---------------------------------------------------------------------

    If CheckStudyOK Then
        CloseTrial
    End If
    
End Sub

'--------------------------------------------------------------------
Private Sub mnuFCopy_Click()
'---------------------------------------------------------------------
' Copy a complete study definition
' REVISIONS
' DPH 15/01/2002 - Unload Trial list after copy so refreshes next time
' NCJ 21 Jun 06 - Log the copy operation
'---------------------------------------------------------------------
Dim sNewName As String

    On Error GoTo ErrHandler
    
    ' NCJ 16/2/01
    If Not CheckStudyOK Then
        Exit Sub
    End If
        
    CloseTrial
    ' NCJ 9 May 06 - Use mode = gsREAD to select a trial to copy
    frmTrialList.Mode = gsREAD
    frmTrialList.Show vbModal

    If frmTrialList.SelectedClinicalTrialId > 0 Then
        sNewName = CopyStudy(frmTrialList.SelectedClinicalTrialId, frmTrialList.SelectedClinicalTrialName)
        ' NCJ 21 Jun 06 - Log study copy
        If sNewName > "" Then
            Call gLog(gsCOPY_TRIAL_SD, "Study [" & frmTrialList.SelectedClinicalTrialName _
                                        & "] copied to new study [" & sNewName & "]")
        End If
        ' DPH 15/01/2002 Unload frmTrialList so will display copied trial without clicking on redisplay
        Unload frmTrialList
    End If

Exit Sub
ErrHandler:
        If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuFCopy_Click", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub mnuFExit_Click()
'---------------------------------------------------------------------
' Exit from SD
' NCJ 16/2/01 - Perform study check before closing
'---------------------------------------------------------------------

    If CheckStudyOK Then
        ' This goes to the MDIForm_Unload event
        Unload Me
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDAttachment_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldInputTool gnATTACHMENT

End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDBackgroundColour_Click()
'---------------------------------------------------------------------
     
    On Error GoTo ErrHandler
    
    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateCRFPageBackgroundColour

    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuFLDBackgroundColour_Click")
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
Private Sub mnuFLDCalendar_Click()
'---------------------------------------------------------------------
     
    On Error GoTo ErrHandler
    
    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldInputTool gnCALENDAR

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuFLDCalendar_Click")
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
Private Sub mnuFLDDeleteCRFPage_Click()
'---------------------------------------------------------------------
'added by Mo Morris 24/8/99
'---------------------------------------------------------------------

    Call AttemptDeleteCRFPage(UpdateMode, frmCRFDesign.tabCRF.SelectedItem.Caption, _
                ClinicalTrialId, VersionId, frmCRFDesign.CRFPageId)

End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDEdit_Click()
'---------------------------------------------------------------------
' Edit the currently selected data item definition
'---------------------------------------------------------------------
Dim oElement As CRFElement

    On Error GoTo ErrHandler
    
    ' Assume active form is CRFDesign
    Set oElement = frmCRFDesign.CurrentCRFElement
    
    If oElement.DataItemId > 0 Then
        frmDataDefinition.ShowDataDefinition Me.ClinicalTrialId, _
                                        Me.VersionId, _
                                        Me.ClinicalTrialName, _
                                        oElement.DataItemId, _
                                        Me.StudyAccessMode
    End If
    
    Set oElement = Nothing

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuFLDEdit_Click")
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
Private Sub mnuFLDEditCaption_Click()
'---------------------------------------------------------------------
Dim oForm As Form
    
    For Each oForm In Forms
        If oForm.Name = "frmCRFDesign" Then
            oForm.Show
        End If
    Next

    Set oForm = Nothing
    
    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").EditCaption
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDEditCRFPage_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").ShowCRFPageDefinition

End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDEditEFG_Click()
'---------------------------------------------------------------------
' Edit an eForm Group
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").EditEFG

End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDEditQGroup_Click()
'---------------------------------------------------------------------
' REM 06/12/01 - Edit a Question Group from the current eForm
'---------------------------------------------------------------------
Dim oQGroup As QuestionGroup
Dim lQGroupId As Long
Dim colDataItemIDs As New Collection

    On Error GoTo ErrLabel
    
    ' The eForm's CurrentElement may be a group OR a group member
    ' First see if it's a group
    lQGroupId = frmCRFDesign.CurrentCRFElement.QGroupID
    
    ' If not a group, it must be a group member so pick up its Owner group
    If lQGroupId = 0 Then
        lQGroupId = frmCRFDesign.CurrentCRFElement.OwnerQGroupID
    End If

    'Set the QGroup object to the one identified by the QGroupId
    Set oQGroup = moQuestionGroups.GroupById(lQGroupId)
    
    ' Display the Group Definition dialog
    Call EditQuestionGroup(oQGroup)
    
    Set oQGroup = Nothing
    
Exit Sub
ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                            "mnuFLDEditQGroup_Click", Err.Source)
    Case OnErrorAction.Retry
        Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDFreeText_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldInputTool gnRICH_TEXT_BOX

End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDHideCaption_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").HideCaption

End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDInsertDataDefinition_Click()
'---------------------------------------------------------------------

    Call InsertNewDataItem

End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDInsertEFQG_Click()
'---------------------------------------------------------------------
' Put an eFormGroup on the eForm
' Assume there's a current eForm, and get the eForm to do all the work
'---------------------------------------------------------------------
Dim oQGroup As QuestionGroup
Dim nReturn As Integer

        nReturn = frmSelectQGroup.Display(moQuestionGroups, oQGroup, False)
        ' See if they selected one
        If Not oQGroup Is Nothing Then
            Call frmCRFDesign.DropEFormGroupOnCRF(oQGroup.QGroupID)
        End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDInsertSpace_Click()
'---------------------------------------------------------------------

    AdjustSpaceOnForm 720
   
End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDOptionBoxes_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldInputTool gnPUSH_BUTTONS + gnOPTION_BUTTONS
End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDPaste_Click()
'---------------------------------------------------------------------
' Paste items in MultiSelect PasteBuffer onto active CRFPage
' RS 6/9/2002
'---------------------------------------------------------------------

    If ActiveForm.Name = "frmCRFDesign" Then
        ActiveForm.PasteSelection
    End If


End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDRemoveSpace_Click()
'---------------------------------------------------------------------

    AdjustSpaceOnForm -720

End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDRenumber_Click()
'---------------------------------------------------------------------
Dim sMSG As String

    'TA 23/11/2000: user now prompted for confirmation
    sMSG = "Do you wish to renumber all questions to reflect their positions on the eForm?"
    If DialogQuestion(sMSG) = vbYes Then
        gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign").Renumber
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuFNew_Click()
'---------------------------------------------------------------------
' Create a new trial
' NCJ 16/2/01 - Perform validity checks on current trial
'---------------------------------------------------------------------

    If CheckStudyOK Then
        Call NewTrial
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuFOpen_Click()
'---------------------------------------------------------------------
' Open an existing trial
' NCJ 16/2/01 - Perform validity checks on current trial
'---------------------------------------------------------------------
     
    On Error GoTo ErrHandler

    If CheckStudyOK Then
        CloseTrial
        ' NCJ 9 May 06 - Use mode = gsUPDATE to select a trial to edit
        frmTrialList.Mode = gsUPDATE
        frmTrialList.Show vbModal
        
        If frmTrialList.SelectedClinicalTrialId > 0 Then
            ' NCJ 10 May 06 - Added StudyAccessMode
            Me.OpenTrial frmTrialList.SelectedClinicalTrialId, _
                        frmTrialList.SelectedClinicalTrialName, _
                        frmTrialList.StudyAccessMode
        End If
    End If
    

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuFOpen_Click")
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
Public Sub mnuFDelete_Click()
'---------------------------------------------------------------------
' Delete a trial, after first closing current trial
' NCJ 16/2/01 - Perform validity checks on current trial before closing
'---------------------------------------------------------------------
Dim lClinicalTrialIdToDelete As Integer
Dim sClinicalTrialNameToDelete As String
     
    On Error GoTo ErrHandler
    
    ' NCJ 16/2/01
    If Not CheckStudyOK Then
        Exit Sub
    End If
    
    If Me.ClinicalTrialId > 0 Then
        'changed by Mo Morris 1/2/99 SR 548
        'If MsgBox("Do you want to delete this trial ?", vbYesNoCancel + vbApplicationModal + vbQuestion) = vbYes Then
            lClinicalTrialIdToDelete = Me.ClinicalTrialId
            sClinicalTrialNameToDelete = Me.ClinicalTrialName
            
            CloseTrial
            Delete lClinicalTrialIdToDelete, sClinicalTrialNameToDelete
        'End If
    Else
    
        CloseTrial
        ' NCJ 9 May 06 - Use mode = gsREAD to select a trial to delete
        frmTrialList.Mode = gsREAD
        frmTrialList.Show vbModal
    
        If frmTrialList.SelectedClinicalTrialId > 0 Then
            Delete frmTrialList.SelectedClinicalTrialId, frmTrialList.SelectedClinicalTrialName
        End If
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuFDelete_Click")
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
Private Function HasStudyBeenDistributed(lTrialId As Long) As Boolean
'---------------------------------------------------------------------
' DPH 27/05/2002 - Has study been distributed to a remote site check
'---------------------------------------------------------------------
Dim rsStudyCheck As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "SELECT COUNT(*) FROM Message WHERE ClinicalTrialId = " & lTrialId & " AND MessageType = " & _
            ExchangeMessageType.NewVersion

    Set rsStudyCheck = New ADODB.Recordset
    rsStudyCheck.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsStudyCheck.Fields(0).Value > 0 Then
        ' Distribute messages exist
        HasStudyBeenDistributed = True
    Else
        HasStudyBeenDistributed = False
    End If
    
    rsStudyCheck.Close
    Set rsStudyCheck = Nothing

Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "HasStudyBeenDistributed")
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
Private Sub mnuFPageBreaks_Click()
'---------------------------------------------------------------------
' NCJ 23/10/00 - View page breaks on the current CRF
'---------------------------------------------------------------------

    ' Set Show Page Breaks only
    gbShowPageBreaksOnly = True
    PrintCRFForm frmCRFDesign
    'Called this routine to displaygrid when printing
    'Ash 24/07/2001
    'Call frmCRFDesign.DisplayGrid
    ' NB PrintCRFForm automatically resets gbShowPageBreaksOnly
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuFPrintAllCRFPages_Click()
'---------------------------------------------------------------------
'   TA 10/10/2000 SR3307: print all eForms
'   by creating and printing each form in turn
'---------------------------------------------------------------------
Dim oTab As MSComctlLib.Tab
Dim oOriginalTab As MSComctlLib.Tab

    On Error GoTo ErrHandler

    If IsPrinterInstalled = True Then
        Set oOriginalTab = frmCRFDesign.tabCRF.SelectedItem
        
        'get user to open printer
        On Error Resume Next
        frmMenu.CommonDialog1.ShowPrinter
        'check for errors in ShowPrinter (incuding a Cancel)
        If Err.Number > 0 Then Exit Sub
        On Error GoTo ErrHandler
        HourglassOn
        For Each oTab In frmCRFDesign.tabCRF.Tabs
            'selecting a tab causes the tabCRF_Click event which builds the new eForm
            frmCRFDesign.tabCRF.SelectedItem = oTab
            PrintCRFForm frmCRFDesign
        Next
        If frmCRFDesign.tabCRF.SelectedItem.Index <> oOriginalTab.Index Then
            'return to the orginal if not already there
            frmCRFDesign.tabCRF.SelectedItem = oOriginalTab
        End If
        HourglassOff
    Else
        MsgBox "You have no printer installed on your machine.", vbInformation, "Print All eForms"
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuFPrintAllCRFPages_Click()")
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
Private Sub mnuFPrintDataDefinitions_Click()
'---------------------------------------------------------------------
    
    'Ash 11/09/2001
    'Call Print_QuestionDefinitionReport


End Sub

'---------------------------------------------------------------------
Private Sub mnuFPrintFormData_Click()
'---------------------------------------------------------------------
    
    'Ash 11/09/2001
    'Call Print_QuestionWithinEformReport

End Sub

'--------------------------------------------------------------------
Private Sub mnuFPrintLibraryData_Click()
'---------------------------------------------------------------------
' WillC added this 8/8/00
'--------------------------------------------------------------------

'TA 27/04/2001: removed until replaced by non crystal code

'    If IsPrinterInstalled = True Then
'        frmLibraryParametersRpt.Show
'    Else
'        MsgBox "You have no printer installed on your machine.", vbInformation, "MACRO"
'    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuFPrintPhaseValidation_Click()
'---------------------------------------------------------------------
' WillC added this 9/8/00
'--------------------------------------------------------------------

'TA 27/04/2001: removed until replaced by non crystal code

'    If IsPrinterInstalled = True Then
'        frmPhaseValidationRpt.Show
'    Else
'        MsgBox "You have no printer installed on your machine.", vbInformation, "MACRO"
'    End If
End Sub

'---------------------------------------------------------------------
Private Sub mnuFPrintTrialList_Click()
'--------------------------------------------------------------------
    
    'Ash 11/09/2001
    'Call Print_StudyDefinitionReport

End Sub

'---------------------------------------------------------------------
Private Sub mnuFPrintCRFPage_Click()
'---------------------------------------------------------------------
' Print current eForm
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If IsPrinterInstalled = True Then
        'get user to open printer
        On Error Resume Next
        frmMenu.CommonDialog1.ShowPrinter
        'check for errors in ShowPrinter (incuding a Cancel)
        If Err.Number > 0 Then Exit Sub
        On Error GoTo ErrHandler
        PrintCRFForm frmCRFDesign
    Else
        Call DialogInformation("You have no printer installed on your machine")
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuFPrintCRFPage_Click()")
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
Private Sub mnuFPrintDataList_Click()
'---------------------------------------------------------------------
' Print Question List Report
'---------------------------------------------------------------------

    'Ash 11/09/2001
    'Call Print_QuestionListReport

End Sub

'---------------------------------------------------------------------
Private Sub mnuFPrintSetup_Click()
'---------------------------------------------------------------------
    
    On Error Resume Next
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter
    'restore normal error trapping
    On Error GoTo 0

End Sub

'---------------------------------------------------------------------
Private Sub mnuFPrintCategory_Click()
'---------------------------------------------------------------------
' Find out using the record count if there is any thing to print.
'---------------------------------------------------------------------
    
    'Ash 11/09/2001
    'Call Print_CategoryQuestionValues

End Sub

'---------------------------------------------------------------------
Private Sub mnuFPrintUnitConversion_Click()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    'Ash 11/09/2001
    'Call Print_UnitConversion

End Sub

'---------------------------------------------------------------------
Private Sub mnuFPrintVisit_Click()
'---------------------------------------------------------------------

    'Call Print_ScheduleVisitReport

End Sub

'---------------------------------------------------------------------
Private Sub mnuHAboutMACRO_Click()
'---------------------------------------------------------------------
    ' PN 26/09/99 display form non modally
    frmAbout.Display

End Sub

'---------------------------------------------------------------------
Private Sub mnuHideIcons_Click()
'---------------------------------------------------------------------
'ZA 09/09/2002 - Display or hide question status icon based on user selection
'---------------------------------------------------------------------
    'toggles the checked and unchecked state of this menu
    mnuHideIcons.Checked = Not mnuHideIcons.Checked
        
    'save user selection in MACRO settings file
    If mnuHideIcons.Checked = True Then
        SetMACROSetting gs_HIDE_QUESTION_STATUS_ICON, CStr(eStatusFlag.Hide)
    Else
        SetMACROSetting gs_HIDE_QUESTION_STATUS_ICON, CStr(eStatusFlag.Show)
    End If
End Sub

'---------------------------------------------------------------------
Private Sub mnuHideRQGIcons_Click()
'---------------------------------------------------------------------
'ZA 09/09/2002 - Display or hide RQG status icon based on user selection
'---------------------------------------------------------------------
    
    'toggles the checked and unchecked state of this menu
    mnuHideRQGIcons.Checked = Not mnuHideRQGIcons.Checked
    
    'save user selection in MACRO settings file
    If mnuHideRQGIcons.Checked = True Then
        SetMACROSetting gs_HIDE_RQG_STATUS_ICON, CStr(eStatusFlag.Hide)
    Else
        SetMACROSetting gs_HIDE_RQG_STATUS_ICON, CStr(eStatusFlag.Show)
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuHUserGuide_Click()
'---------------------------------------------------------------------

    'Call ShowDocument(Me.hwnd, gsMACROUserGuidePath)
    
    'REM 07/12/01 - New Call for the MACRO Help
    Call MACROHelp(Me.hWnd, App.Title)

End Sub

'---------------------------------------------------------------------
Private Sub mnuIQuestionGroup_Click()
'---------------------------------------------------------------------
' REM 20/11/01 create a new Question Group
'---------------------------------------------------------------------
Dim oQG As QuestionGroup
Dim sQGroupCode As String
Dim colDataItemIDs As New Collection
Dim oForm As Form

    'Check to see if the question list is currently displayed and if not then display
    Set oForm = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    If oForm Is Nothing Then
        ' Display current question list
        Call ViewQuestions(gsUPDATE, Me.ClinicalTrialId, Me.ClinicalTrialName)
        Set oForm = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    End If

    'Get the Question Group code from the user
    sQGroupCode = GetItemCode("Question Group", "New " & "Question Group" & " code:")
    
    If sQGroupCode = "" Then    ' if cancel, then return control to user
        Exit Sub
    End If
    
    
    'Create a new Question Group
    Set oQG = moQuestionGroups.NewGroup(sQGroupCode)

    ' Delete the new question group if they pressed cancel, else update question list
    ' NCJ 14 Jun 06 - Allow editing on a new group
    If Not frmQGroupDefinition.Display(oQG, colDataItemIDs, True) Then
        moQuestionGroups.Delete (oQG.QGroupID)
    Else
        Call RefreshQuestionLists(Me.ClinicalTrialId)
        ' NCJ 20 Jun 06 - Mark study as changed
        Call MarkStudyAsChanged
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuLockMACRO_Click()
'---------------------------------------------------------------------
'TA 27/04/2000 Lock application from use - password to re-enter
'---------------------------------------------------------------------
    
    If Not frmTimeOutSplash.Display(False) Then
        ' unload all forms and exit
        Call UnloadAllForms
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuOAutoCaption_Click()
'---------------------------------------------------------------------

    mnuOAutoCaption.Checked = Not mnuOAutoCaption.Checked

End Sub

'---------------------------------------------------------------------
Private Sub mnuOCombinedMovement_Click()
'---------------------------------------------------------------------
' Toggle Combined Movement of captions and fields
'Added by Mo Morris 20/8/99
'---------------------------------------------------------------------

    mnuOCombinedMovement.Checked = Not mnuOCombinedMovement.Checked

End Sub

'---------------------------------------------------------------------
Private Sub mnuOCRFGridLarge_Click()
'---------------------------------------------------------------------
'   SDM 22/02/00 Updated to show grid
'---------------------------------------------------------------------
Dim oForm As Form

    On Error GoTo ErrHandler
    
    mnuOCRFGridSmall.Checked = False
    mnuOCRFGridMedium.Checked = False
    mnuOCRFGridLarge.Checked = True
    mnuOCRFGridNone.Checked = False
     
    HourglassOn
    For Each oForm In Forms
        If oForm.Name = "frmCRFDesign" Then
            oForm.GridDisplay = False
            Call oForm.DisplayGrid
            Call oForm.picCRFPage.Refresh
            oForm.Grid = "Large"
            If mnuOCRFGridShowGrid.Checked Then
                oForm.GridDisplay = True
                Call oForm.DisplayGrid
                Call oForm.picCRFPage.Refresh
            End If
        End If
    Next
    HourglassOff
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuOCRFGridLarge_Click")
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
Private Sub mnuOCRFGridMedium_Click()
'---------------------------------------------------------------------
'   SDM 22/02/00 Updated to show grid
'---------------------------------------------------------------------
Dim oForm As Form

    On Error GoTo ErrHandler
    
    mnuOCRFGridSmall.Checked = False
    mnuOCRFGridMedium.Checked = True
    mnuOCRFGridLarge.Checked = False
    mnuOCRFGridNone.Checked = False
     
    HourglassOn
    For Each oForm In Forms
        If oForm.Name = "frmCRFDesign" Then
            oForm.GridDisplay = False
            Call oForm.DisplayGrid
            Call oForm.picCRFPage.Refresh
            oForm.Grid = "Medium"
            If mnuOCRFGridShowGrid.Checked Then
                oForm.GridDisplay = True
                Call oForm.DisplayGrid
                Call oForm.picCRFPage.Refresh
            End If
        End If
    Next
    HourglassOff
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuOCRFGridMedium_Click")
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
Private Sub mnuOCRFGridNone_Click()
'---------------------------------------------------------------------
'   SDM 22/02/00 Updated to show grid
'---------------------------------------------------------------------
Dim oForm As Form
     
    On Error GoTo ErrHandler
    
    mnuOCRFGridSmall.Checked = False
    mnuOCRFGridMedium.Checked = False
    mnuOCRFGridLarge.Checked = False
    mnuOCRFGridNone.Checked = True
    mnuOCRFGridShowGrid.Checked = False
    
    HourglassOn
    For Each oForm In Forms
        If oForm.Name = "frmCRFDesign" Then
            oForm.GridDisplay = False
            Call oForm.DisplayGrid
            Call oForm.picCRFPage.Refresh
            oForm.Grid = "None"
        End If
    Next
    HourglassOff

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuOCRFGridNone_Click")
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
Private Sub mnuOCRFGridShowGrid_Click()
'---------------------------------------------------------------------
'   SDM 22/02/00 Show the grid
'---------------------------------------------------------------------
Dim oForm As Form

    On Error GoTo ErrHandler
    
    HourglassOn
    mnuOCRFGridShowGrid.Checked = Not mnuOCRFGridShowGrid.Checked
    For Each oForm In Forms
        If oForm.Name = "frmCRFDesign" Then
            oForm.GridDisplay = mnuOCRFGridShowGrid.Checked
            Call oForm.DisplayGrid
            Call oForm.picCRFPage.Refresh
        End If
    Next
    HourglassOff
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuOCRFGridLarge_Click")
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
Private Sub mnuOCRFGridSmall_Click()
'---------------------------------------------------------------------
'   SDM 22/02/00 Updated to show grid
'---------------------------------------------------------------------
Dim oForm As Form

    On Error GoTo ErrHandler
    
    mnuOCRFGridSmall.Checked = True
    mnuOCRFGridMedium.Checked = False
    mnuOCRFGridLarge.Checked = False
    mnuOCRFGridNone.Checked = False
     
    HourglassOn
    For Each oForm In Forms
        If oForm.Name = "frmCRFDesign" Then
            oForm.GridDisplay = False
            Call oForm.DisplayGrid
            Call oForm.picCRFPage.Refresh
            oForm.Grid = "Small"
            If mnuOCRFGridShowGrid.Checked Then
                oForm.GridDisplay = True
                Call oForm.DisplayGrid
                Call oForm.picCRFPage.Refresh
            End If
        End If
    Next
    HourglassOff

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuOCRFGridSmall_Click")
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
Private Sub mnuPDataListDeleteDataD_Click()
'---------------------------------------------------------------------
' NCJ 10/1/02 SR 3498 Make sure we get correct instance of frmDataList
'---------------------------------------------------------------------
Dim oForm As Form

    Set oForm = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    If Not oForm Is Nothing Then
        Call oForm.DeleteDataItem
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuPDataListDeleteForm_Click()
'---------------------------------------------------------------------
' NCJ 14/1/00 Make sure we get correct instance of frmDataList
'---------------------------------------------------------------------
Dim oForm As Form

    Set oForm = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    If Not oForm Is Nothing Then

        Call AttemptDeleteCRFPage(UpdateMode, oForm.SelectedItemName, _
                ClinicalTrialId, VersionId, oForm.SelectedCRFFormId)
        Set oForm = Nothing
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuPDataListDeleteQGroup_Click()
'---------------------------------------------------------------------
'REM 22/07/02 - Delete a Question Group with right click menu
'---------------------------------------------------------------------
Dim oQuestionList As Form
Dim oForm As Form
Dim lQGroupId As Long

    On Error GoTo ErrHandler

    'Make sure we get correct instance of frmDataList
    Set oQuestionList = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    
    If Not oQuestionList Is Nothing Then

        'gets groupID of the QGroup selected in the list box
        lQGroupId = oQuestionList.GetIdFromSelectedItemKey(QGroupID)
          
        'Check that there is a QGroupId
        If lQGroupId > 0 Then
            'Check to see if the selected QGroup is on any EForms
            If moQuestionGroups.IsOnEForm(lQGroupId) = True Then
                If DialogQuestion("This question group is used on one or more EForms. Are you sure you want to delete it?") = vbNo Then
                    Exit Sub
                End If
            Else
                If DialogQuestion("Are you sure you want to delete this question group?") = vbNo Then
                    Exit Sub
                End If
            End If
        
            'Delete the Question Group identified by its QGroupID
            moQuestionGroups.Delete (lQGroupId)
            Call MarkStudyAsChanged
        End If
        
    End If
    
    'if eForm is displayed then refresh
    Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
    If Not oForm Is Nothing Then
        oForm.RefreshMe
    End If

    Set oQuestionList = Nothing
    Set oForm = Nothing
    
    'Refresh question list
    Call RefreshQuestionLists(Me.ClinicalTrialId)

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenuStudyDefintion.mnuPDataListDeleteQGroup_Click"

End Sub

'---------------------------------------------------------------------
Private Sub mnuPDataListDuplicateDataD_Click()
'---------------------------------------------------------------------
' NCJ 10/1/02 Make sure we get correct instance of frmDataList
' MLM 11/06/03: 3.0 bug 1524: refresh questions lists after duplicating question.
'---------------------------------------------------------------------
Dim sCode As String
Dim iDataItemId As Long
Dim rsDataItem As ADODB.Recordset
Dim sSQL As String
Dim oForm As Form
Dim oCurrentForm As Form

    On Error GoTo ErrHandler
    
    ' NCJ 10/1/02
    Set oCurrentForm = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    If Not oCurrentForm Is Nothing Then
    
        With oCurrentForm
            'NCJ 9 Mar 05 - Issue 2539 - Prefill input box with code of original question
            'TA 28/03/2000 - call new function to get code
            sCode = GetItemCode(gsITEM_TYPE_QUESTION, gsITEM_TYPE_QUESTION & " code: ", _
                                    DataItemCodeFromId(.ClinicalTrialId, .SelectedDataItemId))
            
            If sCode <> vbNullString Then
                ' now simply copy the data item
                iDataItemId = CopyDataItem(.ClinicalTrialId, .VersionId, .SelectedDataItemId, _
                                        .ClinicalTrialId, .VersionId, sCode)
                     
                ' the id must be used to obtain the name of the data item
                sSQL = "select DataItemName, DataItemCode from DataItem where dataitemid=" & iDataItemId
                sSQL = sSQL & "and CLinicalTrialID=" & .ClinicalTrialId
                sSQL = sSQL & "and VersionID=" & .VersionId
                Set rsDataItem = New ADODB.Recordset
                rsDataItem.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                
                ' NCJ 10 Jan 02 - Add node to ALL question lists currently displayed
                For Each oForm In Forms
                    If oForm.Name = "frmDataList" Then
                        If oForm.ClinicalTrialId = ClinicalTrialId Then
                            Call oForm.AddDataItemToList(iDataItemId, rsDataItem.Fields(0), rsDataItem.Fields(1))
                            oForm.SelectDataItem iDataItemId, True
                        End If
                    End If
                Next
                
                rsDataItem.Close
                Set rsDataItem = Nothing
            End If
            
        End With
        
        'MLM 11/06/03:
        Call RefreshQuestionLists(Me.ClinicalTrialId)

        ' NCJ 11 Sept 06 - Mark study as changed
        Call MarkStudyAsChanged
        
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuPDataListDuplicateDataD_Click")
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
Private Sub mnuPDataListDuplicateForm_Click()
'---------------------------------------------------------------------
' MLM 31/03/03: Added. Create a new eForm containing all the same
'   CRFElements as the selected one.
' NCJ 30 Apr 03 - Bug fixes (bugs 1651,1656)
'---------------------------------------------------------------------

Dim oForm As Form
Dim lEFormId As Long
Dim sSQL As String
Dim rsElementIds As ADODB.Recordset
Dim colElementIds As Collection

    Set oForm = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    If Not oForm Is Nothing Then
        'create a copy of just the CRFPage row (not the CRFElements)
        lEFormId = DuplicateCRFPage(Me.ClinicalTrialId, Me.VersionId, oForm.SelectedCRFFormId, _
            Me.ClinicalTrialId, Me.VersionId, "Duplicating eForm - " & oForm.SelectedItemName, "")
        If lEFormId = 0 Then
            'cancelled
            Exit Sub
        End If
        
        ' Create a copy of the CRFElements
        ' NCJ 30 Apr 03 - Do NOT include elements from inside groups (Bug 1656)
        sSQL = "SELECT CRFElementId FROM CRFElement WHERE " & _
            " ClinicalTrialId = " & Me.ClinicalTrialId & _
            " AND VersionId = " & Me.VersionId & _
            " AND CRFPageId = " & oForm.SelectedCRFFormId & _
            " AND OwnerQGroupID = 0 " & _
            " ORDER BY CRFElementId"
        Set colElementIds = New Collection
        Set rsElementIds = New ADODB.Recordset
        rsElementIds.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        ' Create a collection to pass through to PasteCRFElements
        Do Until rsElementIds.EOF
            colElementIds.Add rsElementIds.Fields(0).Value
            rsElementIds.MoveNext
        Loop
        rsElementIds.Close
        If colElementIds.Count > 0 Then
            ' NB PasteCRFElements doesn't expect to be passed group questions (only the groups themselves)
            PasteCRFElements Me.ClinicalTrialId, Me.VersionId, oForm.SelectedCRFFormId, lEFormId, colElementIds
        End If
    End If
    
    Call RefreshQuestionLists(Me.ClinicalTrialId)
    
    ' NCJ 30 Apr 03 - Must also refresh schedule and eForm tabs (bug 1651)
    If Not frmMenu.gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmStudyVisits") Is Nothing Then
        If frmStudyVisits.Visible Then
            Call frmStudyVisits.RefreshStudyVisits
        End If
    End If
    
    'if eForm is displayed then refresh
    Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
    If Not oForm Is Nothing Then
        Call oForm.RefreshCRF
    End If
    
    Set oForm = Nothing

    ' NCJ 11 Sept 06 - Mark study as changed
    Call MarkStudyAsChanged

Exit Sub

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.mnuPDataListDuplicateForm_Click"

End Sub

'---------------------------------------------------------------------
Private Sub mnuPDataListDuplicateQGroup_Click()
'---------------------------------------------------------------------
' Copying a question group definition within a study
'---------------------------------------------------------------------
Dim oQG As QuestionGroup
Dim oNewQG As QuestionGroup
Dim sQGroupCode As String
Dim vDataItemId As Variant
Dim oForm As Form

    On Error GoTo ErrLabel

    sQGroupCode = GetItemCode("Question Group", "New " & "Question Group" & " code:")
    If sQGroupCode = "" Then    ' if cancel, then return control to user
        Exit Sub
    End If
    
    Set oForm = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    If Not oForm Is Nothing Then
            
        Set oQG = QuestionGroups.GroupById(oForm.SelectedQGroupID)
        Set oNewQG = QuestionGroups.NewGroup(sQGroupCode)
        oNewQG.Store
        oNewQG.QGroupName = oQG.QGroupName
        oNewQG.DisplayType = oQG.DisplayType
        
        For Each vDataItemId In oQG.Questions
            oNewQG.AddQuestion CLng(vDataItemId)
        Next
                    
        oNewQG.Save
        
    End If
    
    ' NCJ 20 Jun 06 - Mark study as changed
    Call MarkStudyAsChanged
    
    Call RefreshQuestionLists(Me.ClinicalTrialId)
    
    Set oQG = Nothing
    Set oNewQG = Nothing
    Set oForm = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.mnuPDataListDuplicateQGroup_Click"
End Sub

'---------------------------------------------------------------------
Private Sub mnuPDataListEditDataD_Click()
'---------------------------------------------------------------------
Dim oForm As Form

    ' NCJ 14/1/00
    Set oForm = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    If Not oForm Is Nothing Then
        oForm.ShowSelectedDataItem
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuPDataListEditEFG_Click()
'---------------------------------------------------------------------
' NCJ 3 Jan 02
' User wants to edit an eForm Group from the Question List
' NCJ 13 Jun 06 - Only if we have a lock on its eForm
'---------------------------------------------------------------------
Dim oForm As Form
Dim oEFG As EFormGroupSD
Dim lQGroupId As Long
Dim lCRFPageId As Long
Dim oEFGs As EFormGroupsSD
Dim bOnCurrentForm As Boolean
Dim bCanEdit As Boolean

    On Error GoTo ErrLabel
    
    ' Store whether eFormGroup being edited is on currently displayed eForm
    bOnCurrentForm = False
    
    Set oForm = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    If Not oForm Is Nothing Then
        lQGroupId = oForm.SelectedQGroupID
        lCRFPageId = oForm.SelectedCRFFormId
        ' Get the eFormGroup
        ' We see if it's on the currently loaded eForm
        ' and if not, we must load it from the DB
        Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
        If Not oForm Is Nothing Then
            ' See if this is our page and get the eForm group if poss.
            If oForm.CRFPageId = lCRFPageId Then
                bOnCurrentForm = True
                Set oEFG = frmCRFDesign.EFormGroupById(lQGroupId)
            End If
        End If
        ' Did we find the eFormGroup?
        If oEFG Is Nothing Then
            ' Need to load eFormGroups collection from DB
            Set oEFGs = New EFormGroupsSD
            Call oEFGs.Load(Me.ClinicalTrialId, Me.VersionId, lCRFPageId)
            Set oEFG = oEFGs.EFormGroupById(lQGroupId)
        End If
        
        ' Now let the user edit the details
        ' NCJ 13 Jun 06 - Only if we have the eForm lock
        bCanEdit = (meStudyAccess = sdFullControl) Or (lCRFPageId = mlLockedEFormID)
        If frmEFormQGroupDefinition.Display(oEFG, Me.QuestionGroups.GroupById(lQGroupId), bCanEdit) Then
            ' Redraw the current eForm if they did anything
            If bOnCurrentForm Then
                Call RefreshCurrentCRFPage
            End If
        End If
        
        ' Tidy up and throw away the objects we created
        Set oEFG = Nothing
        Set oEFGs = Nothing
        Set oForm = Nothing
    End If
    
Exit Sub
ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                                "mnuPDataListEditEFG_Click", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuPDataListEditForm_Click()
'---------------------------------------------------------------------
' User wants to edit an eForm from the Question List
'---------------------------------------------------------------------
Dim oForm As Form

    Set oForm = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    If Not oForm Is Nothing Then
    'ASH 29/1/2003 Show the form to be editted before showing the
    'page definition form
        Call ViewCRF(oForm.SelectedCRFFormId)
        ' NCJ 15 May 06 - Include eForm access mode
        Call frmCRFPageDefinition.Display(ClinicalTrialId, VersionId, _
                            ClinicalTrialName, oForm.SelectedCRFFormId, Me.eFormAccessMode)
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuPDataListEditQGroup_Click()
'---------------------------------------------------------------------
' User wants to edit a Question Group from the Question List
'---------------------------------------------------------------------
Dim lQGroupId As Long
Dim oForm As Form
Dim oQGroup As QuestionGroup

    Set oForm = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    If Not oForm Is Nothing Then
        Set oQGroup = Me.QuestionGroups.GroupById(oForm.SelectedQGroupID)
        Call EditQuestionGroup(oQGroup)
    End If

    Set oForm = Nothing
    Set oQGroup = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub mnuPDataListInsertdataD_Click()
'---------------------------------------------------------------------

    Call InsertNewDataItem
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuPDataListInsertForm_Click()
'---------------------------------------------------------------------

    Call InsertNewCRFPage

End Sub

'---------------------------------------------------------------------
Private Sub mnuPDatalistInsertQGroup_Click()
'---------------------------------------------------------------------
' Insert Question Group from Data List menu
' Same as from "Insert" menu
'---------------------------------------------------------------------

    Call mnuIQuestionGroup_Click

End Sub

'---------------------------------------------------------------------
Private Sub mnuPDataListViewForm_Click()
'---------------------------------------------------------------------
' ViewCRF for selected data item's CRF page (if any)
'---------------------------------------------------------------------
Dim oForm As Form

    ' NCJ 7 Sept 06 - Check for study updates
    If RefreshIsNeeded Then Exit Sub
    
    Set oForm = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    If Not oForm Is Nothing Then
        ' Changed from SelectedItemID - NCJ 18 Oct 99
        Call ViewCRF(oForm.SelectedCRFFormId)
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuPLab_Click()
'---------------------------------------------------------------------

    frmLabDefinitions.Display

End Sub

'---------------------------------------------------------------------
Private Sub mnuPopUpSubItem_Click(Index As Integer)
'---------------------------------------------------------------------
'TA 19/04/2000 - store item clicked on user-defined menu
'RS 21/08/2002 - This routine was forgotten. Copied from Data Management menu
'---------------------------------------------------------------------
     mlPopUpItem = Index + 1
End Sub

'---------------------------------------------------------------------
Private Sub mnuPStandardFormats_Click()
'---------------------------------------------------------------------

    frmStandardFormatDefinition.Show

End Sub

'---------------------------------------------------------------------
Private Sub ChangeStudyVisitCRFPage(eSelectionType As CellSelectionType)
'---------------------------------------------------------------------
' implemented for new menu option
'---------------------------------------------------------------------
Dim oForm As Form
     
    On Error GoTo ErrHandler
    
    ' add the form selected to the visit
    Set oForm = gFindForm(ClinicalTrialId, VersionId, "frmStudyVisits")
    If Not oForm Is Nothing Then
        oForm.ChangeFormVisitDetails (eSelectionType)
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.ChangeStudyVisitCRFPage"
    
End Sub

'---------------------------------------------------------------------
Private Sub RefreshSchedule()
'---------------------------------------------------------------------
' NCJ 3 Apr 03 - Refresh the schedule if it's showing
' e.g. if they changed any visit cycles in AREZZO
'---------------------------------------------------------------------
Dim oForm As Form
     
    On Error GoTo ErrHandler
    
    ' See if the schedule is showing
    Set oForm = gFindForm(ClinicalTrialId, VersionId, "frmStudyVisits")
    If Not oForm Is Nothing Then
        If oForm.Visible Then
            Call oForm.BuildStudyVisits
        End If
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.RefreshSchedule"
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuPStudyVisitAllowMultipleForms_Click()
'---------------------------------------------------------------------

    Call ChangeStudyVisitCRFPage(AllowRepeating)

End Sub

'---------------------------------------------------------------------
Private Sub mnuPStudyVisitAllowSingleForm_Click()
'---------------------------------------------------------------------

    Call ChangeStudyVisitCRFPage(AllowSingle)

End Sub

'---------------------------------------------------------------------
Private Sub mnuPStudyVisitBackgroundColour_Click()
'---------------------------------------------------------------------
' update the background colour of the selected column
'---------------------------------------------------------------------
Dim oForm As Form
    
    On Error GoTo CancelSelected
    CommonDialog1.CancelError = True
    CommonDialog1.ShowColor

    Set oForm = gFindForm(ClinicalTrialId, VersionId, "frmStudyVisits")
    If Not oForm Is Nothing Then
        oForm.SetBackgroungColour CommonDialog1.Color
    End If
    
CancelSelected:

End Sub

'---------------------------------------------------------------------
Private Sub mnuPStudyVisitDeleteVisit_Click()
'---------------------------------------------------------------------
Dim oForm As Form
     
     On Error GoTo ErrHandler
    ' delete the visit is entirely handled by the studyvisit form
    Set oForm = gFindForm(ClinicalTrialId, VersionId, "frmStudyVisits")
    If Not oForm Is Nothing Then
        oForm.Delete
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuPStudyVisitDeleteVisit_Click")
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
Private Sub mnuPStudyVisitInsertVisit_Click()
'---------------------------------------------------------------------
Dim oForm As Form
     
     On Error GoTo ErrHandler
    ' delete the visit is entirely handled by the studyvisit form
    Set oForm = gFindForm(ClinicalTrialId, VersionId, "frmStudyVisits")
    If Not oForm Is Nothing Then
        oForm.InsertVisit
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuPStudyVisitInsertVisit_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'ZA 28/08/2002 - don't need it any more
'---------------------------------------------------------------------
'Private Sub mnuPStudyVisitPromptUser_Click()
''---------------------------------------------------------------------
'' TA    17/03/2000
''---------------------------------------------------------------------
'Dim oForm As Form
'
'     On Error GoTo ErrHandler
'    ' toggle the date user prompt is entirely handled by the studyvisit form
'    Set oForm = gFindForm(ClinicalTrialId, VersionId, "frmStudyVisits")
'    If Not oForm Is Nothing Then
'        oForm.ToggleDatePrompt
'    End If
'
'Exit Sub
'ErrHandler:
'    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
'                                    "mnuPStudyVisitPromptUser_Click")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'    End Select
'End Sub

'---------------------------------------------------------------------
Private Sub mnuPStudyVisitRemoveFormFromVisit_Click()
'---------------------------------------------------------------------

    Call ChangeStudyVisitCRFPage(NotAssigned)

End Sub

'---------------------------------------------------------------------
Private Sub mnuPStudyVisitViewEform_Click()
'---------------------------------------------------------------------
' NCJ 28 Sept 06 - New Schedule menu item to view eForm from schedule
'---------------------------------------------------------------------

    Call frmStudyVisits.GridDblClick

End Sub

'---------------------------------------------------------------------
Private Sub mnuPTrialPhases_Click()
'---------------------------------------------------------------------

    ' NCJ 8 Dec 99 - Use TrialType form with type "Phase"
    frmTrialType.FormType = "Phases"
    frmTrialType.Show vbModal

End Sub

'---------------------------------------------------------------------
Private Sub mnuPTrialType_Click()
'---------------------------------------------------------------------

    ' NCJ 8 Dec 99 - Set to edit Trial Types
    frmTrialType.FormType = "Types"
    frmTrialType.Show vbModal

End Sub

'---------------------------------------------------------------------
Private Sub mnuPUnitMaintenance_Click()
'---------------------------------------------------------------------

    frmUnitMaintenance.Show vbModal

End Sub

'---------------------------------------------------------------------
Private Sub mnuPValidationType_Click()
'---------------------------------------------------------------------

    frmValidationType.Show vbModal
 
End Sub

'---------------------------------------------------------------------
Private Sub mnuRAttachment_Click()
'---------------------------------------------------------------------

gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldInputTool gnATTACHMENT

End Sub

'---------------------------------------------------------------------
Private Sub mnuRCalendar_Click()
'---------------------------------------------------------------------

gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldInputTool gnCALENDAR

End Sub

'---------------------------------------------------------------------
Private Sub mnuRDefaultBackgroundColour_Click()
'---------------------------------------------------------------------
' Change the default background colour for forms
' NCJ 7 Jan 00, SR 2583 - Refresh current CRF if there is one
'---------------------------------------------------------------------
Dim oForm As Form
Dim tmpDefaultFontItalic As Integer
Dim tmpDefaultFontBold As Integer

    On Error Resume Next
    
    CommonDialog1.Color = frmCRFDesign.DefaultCRFColour
    CommonDialog1.Flags = &H1&
    CommonDialog1.ShowColor
    
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        On Error GoTo 0
    End If
    
    frmCRFDesign.DefaultCRFColour = CommonDialog1.Color
    
    
    If frmCRFDesign.DefaultFontBold Then
        tmpDefaultFontBold = 1
    Else
        tmpDefaultFontBold = 0
    End If
    
    If frmCRFDesign.DefaultFontItalic Then
        tmpDefaultFontItalic = 1
    Else
        tmpDefaultFontItalic = 0
    End If

    gdsUpdateStudyDefinition ClinicalTrialId, VersionId, _
                frmCRFDesign.DefaultFontColour, frmCRFDesign.DefaultCRFColour, frmCRFDesign.DefaultFontName, tmpDefaultFontBold, _
                tmpDefaultFontItalic, frmCRFDesign.DefaultFontSize
                
    ' NCJ 7 Jan 00 - Rebuild current page if loaded & visible
    ' NCJ 16/10/00 - To fix SR3980
    Call RefreshCurrentCRFPage
    
End Sub

'---------------------------------------------------------------------
Public Function RefreshIsNeeded(Optional bForceRefresh As Boolean = False) As Boolean
'---------------------------------------------------------------------
' NCJ 7 Sept 06
' Do a complete refresh if the study def has been changed by another user,
' or if bForceRefresh = TRUE
' Returns TRUE if things were updated; returns FALSE if nothing changed
'---------------------------------------------------------------------
Dim lThisCRFPageID As Long
Dim oForm As Form
Dim rsTrialDetails As ADODB.Recordset

    On Error GoTo ErrLabel
    
    RefreshIsNeeded = False
    
    ' Check for loaded study
    If Me.ClinicalTrialId < 1 Then Exit Function
    
    ' Has the study changed?
    If StudyHasChanged Or bForceRefresh Then
        RefreshIsNeeded = True
        
        ' Check the study still exists!
        ' read clinical trial table
        Set rsTrialDetails = New ADODB.Recordset
        Set rsTrialDetails = gdsTrialDetails(Me.ClinicalTrialId)
        If rsTrialDetails.EOF Then
            DialogInformation "This study has been deleted by another user."
           ' BUT NOW WHAT?? Difficult to jump out of everything from here.... but we'll try
           Call CloseTrial
           Exit Function
        End If
        
        ' Do the question lists and RQGs
        Call RefreshQuestionLists(Me.ClinicalTrialId)
        Call RefreshQuestionGroups
        
        ' Do the schedule if it's showing
        Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmStudyVisits")
        If Not oForm Is Nothing Then
            Call oForm.RefreshStudyVisits
        End If
        
        ' Do the eForms if they're showing
        Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
        If Not oForm Is Nothing Then
            ' Try to preserve the current eForm
            If oForm.Visible Then
                ' Do the Full Monty
                Call oForm.RefreshCRFPageList(oForm.CRFPageId)
                Call oForm.RefreshMe
            Else
                ' Just redo the tabs to make sure they're consistent
                Call oForm.RefreshCRF
            End If
        End If
            
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.RefreshIsNeeded"
    
End Function

'---------------------------------------------------------------------
Private Sub RefreshCurrentCRFPage()
'---------------------------------------------------------------------
' NCJ 16/10/00, This is to fix SR3980 when needed
' Refresh current CRF page if there is one and it's visible
'---------------------------------------------------------------------
Dim oForm As Form

    Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
    If Not oForm Is Nothing Then
        If oForm.Visible Then
            ' NCJ 16/10/00 SR 3980 Check there is an eForm to build!
            If oForm.CRFPageId > 0 Then
                Call oForm.RefreshMe
            End If
        End If
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuRDefaultFont_Click()
'---------------------------------------------------------------------
'Changed by  Mo Morris   7/9/98     (change made in Released and Developed versions)
'   CommonDialog.ShowFont flags now set to cdlCFBoth Or cdlCFScalableOnly Or cdlCFWYSIWYG
'instead of cdlCFPrinterFonts. This restricts dialog to scalable fonts that are available
'on both Screen and Printer.
' NCJ 7 Jan 00, SR 2583 - Rebuild current form if visible
'---------------------------------------------------------------------
Dim oForm As Form
Dim tmpDefaultFontItalic As Integer
Dim tmpDefaultFontBold As Integer
    
    On Error Resume Next
    
    CommonDialog1.FontName = frmCRFDesign.DefaultFontName
    CommonDialog1.FontBold = frmCRFDesign.DefaultFontBold
    CommonDialog1.FontItalic = frmCRFDesign.DefaultFontItalic
    CommonDialog1.FontSize = frmCRFDesign.DefaultFontSize
    
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
    
    frmCRFDesign.DefaultFontName = CommonDialog1.FontName
    frmCRFDesign.DefaultFontSize = CommonDialog1.FontSize
    frmCRFDesign.DefaultFontBold = CommonDialog1.FontBold
    frmCRFDesign.DefaultFontItalic = CommonDialog1.FontItalic
    
    If frmCRFDesign.DefaultFontBold Then
        tmpDefaultFontBold = 1
    Else
        tmpDefaultFontBold = 0
    End If
    
    If frmCRFDesign.DefaultFontItalic Then
        tmpDefaultFontItalic = 1
    Else
        tmpDefaultFontItalic = 0
    End If
    
    gdsUpdateStudyDefinition ClinicalTrialId, VersionId, _
                frmCRFDesign.DefaultFontColour, frmCRFDesign.DefaultCRFColour, frmCRFDesign.DefaultFontName, tmpDefaultFontBold, _
                tmpDefaultFontItalic, frmCRFDesign.DefaultFontSize
                
    ' NCJ 7 Jan 00 - Rebuild current page if loaded & visible
    ' NCJ 16/10/00 - To fix SR3980
    Call RefreshCurrentCRFPage
'    Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
'    If Not oForm Is Nothing Then
'        If oForm.Visible Then
'            Call BuildCRFPage(oForm, oForm.CRFPageId)
'        End If
'    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuRDefaultFontColour_Click()
'---------------------------------------------------------------------
' NCJ 13 Dec 99, SR 2375 - Save the new value in frmCRFDesign
' NCJ 7 Jan 00 - Refresh current CRF if any
'---------------------------------------------------------------------
Dim tmpDefaultFontItalic As Integer
Dim tmpDefaultFontBold As Integer
Dim lNewColour As Long
Dim oForm As Form

    On Error GoTo ErrHandler

    On Error Resume Next
    
    CommonDialog1.Color = frmCRFDesign.DefaultFontColour
    CommonDialog1.Flags = &H1&
    CommonDialog1.ShowColor
    
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        On Error GoTo 0
    End If
    
    ' NCJ 13 Dec 99
    lNewColour = CommonDialog1.Color
    frmCRFDesign.DefaultFontColour = lNewColour
        
    If frmCRFDesign.DefaultFontBold Then
        tmpDefaultFontBold = 1
    Else
        tmpDefaultFontBold = 0
    End If
    
    If frmCRFDesign.DefaultFontItalic Then
        tmpDefaultFontItalic = 1
    Else
        tmpDefaultFontItalic = 0
    End If
    
    gdsUpdateStudyDefinition ClinicalTrialId, VersionId, _
                lNewColour, frmCRFDesign.DefaultCRFColour, _
                frmCRFDesign.DefaultFontName, tmpDefaultFontBold, _
                tmpDefaultFontItalic, frmCRFDesign.DefaultFontSize
    
     ' NCJ 7 Jan 00 - Rebuild current page if loaded & visible
    ' NCJ 16/10/00 - To fix SR3980
    Call RefreshCurrentCRFPage
'    Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
'    If Not oForm Is Nothing Then
'        If oForm.Visible Then
'            Call BuildCRFPage(oForm, oForm.CRFPageId)
'        End If
'    End If
   
    Exit Sub
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuRDefaultFontColour_Click")
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
Private Sub mnuRFont_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldFont
End Sub

'---------------------------------------------------------------------
Private Sub mnuRFontColour_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldFontColour

End Sub

'---------------------------------------------------------------------
Private Sub mnuRFreeText_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldInputTool gnRICH_TEXT_BOX

End Sub

'---------------------------------------------------------------------
Private Sub mnuROptionBoxes_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldInputTool gnPUSH_BUTTONS

End Sub

'---------------------------------------------------------------------
Private Sub mnuROptionButtons_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldInputTool gnOPTION_BUTTONS

End Sub

'---------------------------------------------------------------------
Private Sub mnuRPopupList_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldInputTool gnPOPUP_LIST

End Sub

'---------------------------------------------------------------------
Private Sub mnuRTextBox_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldInputTool gnTEXT_BOX

End Sub

'---------------------------------------------------------------------
Private Sub mnuUseFormForDateValidation_Click()
'---------------------------------------------------------------------
'---------------------------------------------------------------------
    
    Call ChangeStudyVisitCRFPage(CellSelectionType.VisitForm)

End Sub

'---------------------------------------------------------------------
Private Sub mnuVAREZZOReport_Click()
'---------------------------------------------------------------------
' NCJ 5 June 03
' Show a report on all AREZZO terms
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call frmAREZZOReport.Display(Me.ClinicalTrialId, Me.ClinicalTrialName)

ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuVAREZZOReport_Click", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVCRF_Click()
'---------------------------------------------------------------------
     
    On Error GoTo ErrHandler
    
    ' NCJ 7 Sept 06 - Check for study updates
    If RefreshIsNeeded Then Exit Sub
    
    If mnuVCRF.Checked = True Then
        mnuVCRF.Checked = False
        HideCRF
    Else
        mnuVCRF.Checked = True
        ViewCRF
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mnuVCRF_Click")
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
Private Sub mnuVDataCRFPage_Click()
'---------------------------------------------------------------------
' View data by CRF Page toggle
'---------------------------------------------------------------------
     
    On Error GoTo ErrLabel
    
    ' NCJ 7 Sept 06 - Check for study updates
    If RefreshIsNeeded Then Exit Sub
    
    If mnuVDataCRFPage.Checked = True Then
        mnuVDataCRFPage.Checked = False
    Else
        mnuVDataCRFPage.Checked = True
    End If
    
    If mnuVDataList.Checked = True Then
        ' Update all current question lists
        Call RefreshDatalistWindows
    End If

Exit Sub
ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuVDataCRFPage_Click", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuVDataList_Click()
'---------------------------------------------------------------------
     
     On Error GoTo ErrLabel
     
    ' NCJ 7 Sept 06 - Check for study updates
    If RefreshIsNeeded Then Exit Sub
    
    If mnuVDataList.Checked = True Then
        Call HideQuestions(gsUPDATE, Me.ClinicalTrialId, True)
    Else
        Call ViewQuestions(gsUPDATE, Me.ClinicalTrialId, Me.ClinicalTrialName)
    End If

Exit Sub
ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuVDataList_Click", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuVDisplayDataByName_Click()
'---------------------------------------------------------------------
' Display question list by Name
'---------------------------------------------------------------------
      
    On Error GoTo ErrLabel
    
    ' NCJ 7 Sept 06 - Check for study updates
    If RefreshIsNeeded Then Exit Sub
    
    If mnuVDisplayDataByName.Checked Then
        mnuVDisplayDataByName.Checked = False
        Call SaveSetting(App.Title, "Settings", "ViewDataListName", "0")
    Else
        mnuVDisplayDataByName.Checked = True
        Call SaveSetting(App.Title, "Settings", "ViewDataListName", "-1")
    End If
    
    Call RefreshDatalistWindows

Exit Sub
ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuVDisplayDataByName_Click", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuVDisplayFormsAlphabetically_Click()
'---------------------------------------------------------------------
' Display eForms alphabetically in Question List windows
'---------------------------------------------------------------------
      
    On Error GoTo ErrLabel
    
    ' NCJ 7 Sept 06 - Check for study updates
    If RefreshIsNeeded Then Exit Sub
    
    If mnuVDisplayFormsAlphabetically.Checked Then
        mnuVDisplayFormsAlphabetically.Checked = False
        Call SaveSetting(App.Title, "Settings", "DataListFormOrderAlphabetic", "0")
    Else
        mnuVDisplayFormsAlphabetically.Checked = True
        Call SaveSetting(App.Title, "Settings", "DataListFormOrderAlphabetic", "-1")
    End If
    
    ' NCJ 15 Jan 02 - Refresh ALL data list windows
    Call RefreshDatalistWindows

Exit Sub
ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuVDisplayFormsAlphabetically_Click", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub RefreshDatalistWindows()
'---------------------------------------------------------------------
' Refresh ALL currently displayed data list windows
' (See also RefreshQuestionLists to refresh for a specific study)
'---------------------------------------------------------------------
Dim oForm As Form
     
    On Error GoTo ErrLabel
    
    For Each oForm In Forms   'iterate through the forms collection
        If oForm.Name = "frmDataList" Then
            Call oForm.RefreshDataList
        End If
    Next

    Set oForm = Nothing

Exit Sub
ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "RefreshDatalistWindows", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuVLibrary_Click()
'---------------------------------------------------------------------

    If mnuVLibrary.Checked = True Then
        Call HideQuestions(gsREAD, 0, True)
    Else
        Call ViewQuestions(gsREAD, 0, gsLIBRARY_LABEL)
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVArezzo_Click()
'---------------------------------------------------------------------

    'ViewProforma
    LaunchMTMCUI

End Sub

'---------------------------------------------------------------------
Private Sub mnuVReferences_Click()
'---------------------------------------------------------------------

    ' NCJ 7 Sept 06 - Check for study updates
    If RefreshIsNeeded Then Exit Sub
    
    If mnuVReferences.Checked = True Then
        mnuVReferences.Checked = False
        HideReferences
    Else
        mnuVReferences.Checked = True
        ViewReferences
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVStudyDefinition_Click()
'---------------------------------------------------------------------

    ' NCJ 7 Sept 06 - Check for study updates
    If RefreshIsNeeded Then Exit Sub
    
    If mnuVStudyDefinition.Checked = True Then
        mnuVStudyDefinition.Checked = False
        HideStudyDefinition
    Else
        mnuVStudyDefinition.Checked = True
        ViewStudyDetails
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuVTrialList_Click()
'---------------------------------------------------------------------
     
    On Error GoTo ErrHandler
    
    ' NCJ 9 May 06 - Use mode = gsREAD to select a trial to view
    frmTrialList.Mode = gsREAD
    frmTrialList.Show vbModal
    
    'stop outline of form lingering
    DoEvents
    
'   ATN 1/5/99
'   Check that a trial has been selected before opening the data list window

    If frmTrialList.SelectedClinicalTrialId > 0 Then
            Call ViewQuestions(gsREAD, frmTrialList.SelectedClinicalTrialId, _
                                frmTrialList.SelectedClinicalTrialName)
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "mnuVTrialList_Click")
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
Private Sub mnuVVisits_Click()
'---------------------------------------------------------------------

    ' NCJ 7 Sept 06 - Check for study updates
    If RefreshIsNeeded Then Exit Sub
    
    If mnuVVisits.Checked = True Then
        mnuVVisits.Checked = False
        HideVisits
    Else
        mnuVVisits.Checked = True
        ViewVisits
    End If

End Sub

'---------------------------------------------------------------------
Private Sub tmrSystemIdleTimeout_Timer()
'---------------------------------------------------------------------
' nb This timer event should never occur unless DevMode = 0
' when the timer goes off it must be time to lock the system
'  it prompts the user to enter the password or wxit MACRO
' the system is then either closed in a controlled way or resets the timer
' NCJ 17/3/00 - Tidied up and simplified (SR 3015)
'TA 27/04/2000: new timeout handling
'---------------------------------------------------------------------
    
    'new timeout handling
    glSystemIdleTimeoutCount = glSystemIdleTimeoutCount + 1
    If glSystemIdleTimeout = glSystemIdleTimeoutCount Then
        ' set the couter to 0 and disable the timer until the user logs in
        glSystemIdleTimeoutCount = 0
        tmrSystemIdleTimeout.Enabled = False
        If frmTimeOutSplash.Display(False) Then
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
Private Sub tlbMenu_ButtonClick(ByVal Button As Button)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    ' NCJ 7 Sept 06 - Check for study updates
    Call RefreshIsNeeded
    
    ' NCJ 17 Apr 07 - Check study still exists! (in case deleted)
    If Me.ClinicalTrialId < 0 Then Exit Sub
    
    Select Case Button.Key
    Case gsTRIAL_LABEL
        If Button.Value = tbrPressed Then
            ViewStudyDetails
        Else
            ' PN change
            ' pass PromptBeforeSave=True so that the
            ' user is prompted to save changes
            HideStudyDefinition (False)
        End If
    Case gsLIBRARY_LABEL
        If Button.Value = tbrPressed Then
            Call ViewQuestions(gsREAD, 0, gsLIBRARY_LABEL)
        Else
            Call HideQuestions(gsREAD, 0, True)
        End If
' NCJ 10 May 06 - Trial List button no longer used
'    Case gsTRIAL_LIST_LABEL
'        frmTrialList.Show vbModal
    Case gsDATA_ITEM_LABEL
        If Button.Value = tbrPressed Then
            Call ViewQuestions(gsUPDATE, Me.ClinicalTrialId, Me.ClinicalTrialName)
        Else
            Call HideQuestions(gsUPDATE, Me.ClinicalTrialId, True)
        End If
    Case gsCRF_PAGE_LABEL
        If Button.Value = tbrPressed Then
            ViewCRF
            'Mo 16/5/2003  Bug 1577
            mnuFPrintAllCRFPages.Enabled = True
            mnuFPrintCRFPage.Enabled = True
        Else
            HideCRF
            'Mo 16/5/2003  Bug 1577
            mnuFPrintAllCRFPages.Enabled = False
            mnuFPrintCRFPage.Enabled = False
        End If
    Case gsVISIT_LABEL
        If Button.Value = tbrPressed Then
            ViewVisits
        Else
            HideVisits
        End If
    Case gsPROFORMA_LABEL
        'ViewProforma
        LaunchMTMCUI
    
    Case gsDOCUMENT_LABEL
        If Button.Value = tbrPressed Then
            ViewReferences
        Else
            HideReferences
        End If
    Case gsLINE_LABEL
        tlbMenu.Buttons(gsCOMMENT_LABEL).Value = tbrUnpressed
        tlbMenu.Buttons(gsPICTURE_LABEL).Value = tbrUnpressed
        tlbMenu.Buttons(gsLINK_LABEL).Value = tbrUnpressed
        
    Case gsCOMMENT_LABEL
        tlbMenu.Buttons(gsLINE_LABEL).Value = tbrUnpressed
        tlbMenu.Buttons(gsPICTURE_LABEL).Value = tbrUnpressed
        tlbMenu.Buttons(gsLINK_LABEL).Value = tbrUnpressed
        
    Case gsPICTURE_LABEL
        tlbMenu.Buttons(gsLINE_LABEL).Value = tbrUnpressed
        tlbMenu.Buttons(gsCOMMENT_LABEL).Value = tbrUnpressed
        tlbMenu.Buttons(gsLINK_LABEL).Value = tbrUnpressed
    
    Case gsLINK_LABEL   ' NCJ 6 Nov 02
        tlbMenu.Buttons(gsLINE_LABEL).Value = tbrUnpressed
        tlbMenu.Buttons(gsCOMMENT_LABEL).Value = tbrUnpressed
        tlbMenu.Buttons(gsPICTURE_LABEL).Value = tbrUnpressed
        
    End Select

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "tlbMenu_ButtonClick", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuCRFBackgroundColour_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateCRFPageBackgroundColour

End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDDeleteField_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").DeleteCRFElement

End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDFont_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldFont
          
End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDFontColour_Click()
'---------------------------------------------------------------------

    gFindForm(Me.ClinicalTrialId, _
          Me.VersionId, _
          "frmCRFDesign").UpdateFieldFontColour
          
End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDOptionButtons_Click()
'---------------------------------------------------------------------
Dim oForm As Form
    
    Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
    If Not oForm Is Nothing Then
        Call oForm.UpdateFieldInputTool(gnOPTION_BUTTONS)
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDPopupList_Click()
'---------------------------------------------------------------------
Dim oForm As Form
    
    Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
    If Not oForm Is Nothing Then
        Call oForm.UpdateFieldInputTool(gnPOPUP_LIST)
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuFLDTextBox_Click()
'---------------------------------------------------------------------
Dim oForm As Form

    Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
    If Not oForm Is Nothing Then
        Call oForm.UpdateFieldInputTool(gnTEXT_BOX)
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuICRFPage_Click()
'---------------------------------------------------------------------
    
    Call InsertNewCRFPage

End Sub

'---------------------------------------------------------------------
Private Sub InsertNewCRFPage()
'---------------------------------------------------------------------
Dim oForm As Form

    ' Find form CRFDesign if it exists
    Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
    ' If not, create it
    If oForm Is Nothing Then
        ViewCRF
        Set oForm = gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
    End If
    Call oForm.InsertCRFPage

End Sub

'---------------------------------------------------------------------
Private Sub InsertNewDataItem()
'---------------------------------------------------------------------
' Insert a new data item
'---------------------------------------------------------------------
Dim oForm As Form

    Set oForm = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    If oForm Is Nothing Then
        ' Display current question list
        Call ViewQuestions(gsUPDATE, Me.ClinicalTrialId, Me.ClinicalTrialName)
        Set oForm = gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
    End If
    Call oForm.InsertDataItem
    
    Set oForm = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub mnuIDataItem_Click()
'---------------------------------------------------------------------
    
    Call InsertNewDataItem
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuIVisit_Click()
'---------------------------------------------------------------------

    If gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmStudyVisits") Is Nothing Then
        ViewVisits
    End If
    gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmStudyVisits").InsertVisit

End Sub

'---------------------------------------------------------------------
Public Function gFindForm(vClinicalTrialId As Long, _
                          vVersionId As Integer, _
                          vFormName As String) As Form
'---------------------------------------------------------------------
' Find form of given name if it's loaded
' Make sure TrialId and VersionId match
' If form isn't currently loaded, return Nothing
' NB This only ever returns the FIRST matching form
'   For frmDataList, use gFindQuestionList instead
'---------------------------------------------------------------------
Dim oForm As Form
     
    On Error GoTo ErrHandler
    
    For Each oForm In Forms   'iterate through the forms collection
        If oForm.Name = vFormName Then    'same name and trial key'
            If oForm.ClinicalTrialId = vClinicalTrialId _
            And oForm.VersionId = vVersionId Then
                Set gFindForm = oForm     'return form
                Exit Function                   'exit from loop
            End If
        End If
    Next
    
    Set gFindForm = Nothing
    
Exit Function
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "gFindForm", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If

End Function

'---------------------------------------------------------------------
Public Function gFindQuestionList(vClinicalTrialId As Long, _
                          vVersionId As Integer) As Form
'---------------------------------------------------------------------
' Find current updateable Question List for given study
' if it's loaded
' Make sure TrialId and VersionId match
' If form isn't currently loaded, return Nothing
'---------------------------------------------------------------------
Dim oForm As Form
     
    On Error GoTo ErrLabel
    
    For Each oForm In Forms   'iterate through the forms collection
        If oForm.Name = "frmDataList" Then    'same name and trial key'
            If oForm.ClinicalTrialId = vClinicalTrialId _
            And oForm.VersionId = vVersionId And oForm.UpdateMode <> gsREAD Then
                Set gFindQuestionList = oForm     'return form
                Exit Function                   'exit from loop
            End If
        End If
    Next
    
    Set gFindQuestionList = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.gFindQuestionList"
    
End Function

'---------------------------------------------------------------------
Public Sub NewTrial(Optional ByVal vCopyFromClinicalTrialId As Variant)
'---------------------------------------------------------------------
'Mo Morris  14/2/00 sections re-written
'TA 16/04/2002: Return whether insert was successful
'---------------------------------------------------------------------
Dim sClinicalTrialName As String
Dim lClinicalTrialId As Long
Dim nVersionId As Integer
     
    On Error GoTo ErrHandler

    ' Get valid name from user
    Do
        sClinicalTrialName = InputBox("Study name: ", gsDIALOG_TITLE, sClinicalTrialName)
        If sClinicalTrialName = "" Then    ' if cancel, then exit from this form
            Exit Sub
        End If
    Loop Until ValidateTrialName(sClinicalTrialName)
    
    ' NCJ 17 Jan 03 - Must close previous trial before creating new one
    ' (This is important to do here, now that we load AREZZO with each trial)
    Call CloseTrial
    
    HourglassOn
    
    ' NCJ 19 Sept 06 - For new trial, must ensure CLMSave flag is set!
    gbDoCLMSave = True
    
    ' NCJ 17 Jan 03 - Start up CLM for new trial with default memory settings (pass -1 as study id)
    Call StartUpCLM(-1)
    
    'TA 16/04/2002: Return whether insert was successful
    If InsertTrial(lClinicalTrialId, nVersionId, sClinicalTrialName, vCopyFromClinicalTrialId) Then
        ' NCJ 21 Jun 06 - Issue 2745 - Log study created
        Call gLog(gsNEW_TRIAL_SD, "New study created [" & sClinicalTrialName & "]")
        'successful
        ' Display the trial and study definition details
        ' NCJ 14 Jun 06 - Access mode for new trial is Full Control
        OpenTrial lClinicalTrialId, sClinicalTrialName, eSDAccessMode.sdFullControl
              
        ' NCJ 13 Dec 99 - Check rights to edit study details (SR 98)
        If goUser.CheckPermission(gsFnEditStudyDetails) Then
            ViewStudyDetails
            
            '   ATN 29/4/99 SR834
            '   refresh study details window with new details for the new trial
            '   ATN 27/1/2000   SR2511
            '   Set the flag to indicate this is a new trial
            frmStudyDefinition.IsNew = True
            frmStudyDefinition.RefreshTrialDetails
        End If
    End If
    
    HourglassOff
    
Exit Sub

ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "NewTrial", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
End Sub


'---------------------------------------------------------------------
Public Sub CopyTrial(Optional ByVal vCopyFromClinicalTrialId As Variant)
'---------------------------------------------------------------------

    Call NewTrial(vCopyFromClinicalTrialId)

End Sub

'---------------------------------------------------------------------
Public Sub Delete(ByVal lClinicalTrialId As Long, _
                  ByVal sClinicalTrialName As String)
'---------------------------------------------------------------------
' Inquires if the trial is in Preparation if so it allows its deletion if not then
' the user is not allowed to delete the trial...
' NCJ 30/3/01 - Changed message box
' DPH 27/05/2002 - If study has been distributed then disallow deletion
' MLM 18/06/02: CBB 2.2.15/17 Deleting a study requires a study lock.
' MLM 11/07/02: CBB 2.2.18/3 Rebuild Protocols collection after deleting study.
' NCJ 25 Nov 04 - Ignore errors when releasing study lock
'---------------------------------------------------------------------
Dim rsTrialStatus As ADODB.Recordset
Dim sSQL As String
Dim nStatus As Integer
Dim sLockDetails As String

    On Error GoTo ErrHandler

    ' NCJ 20/12/99 - Changed to get trial status directly
    Set rsTrialStatus = New ADODB.Recordset
    sSQL = "SELECT StatusId from ClinicalTrial" _
        & " WHERE ClinicalTrialName = '" & sClinicalTrialName & "'"
    rsTrialStatus.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    nStatus = rsTrialStatus!statusId
    
    rsTrialStatus.Close
    Set rsTrialStatus = Nothing
    
    If nStatus <> eTrialStatus.InPreparation Then
        Call DialogInformation("You may not delete a study which is not in preparation.")
        Exit Sub
    End If
    
    ' DPH 27/05/2002 - If study has been distributed then disallow deletion
    If HasStudyBeenDistributed(lClinicalTrialId) Then
        Call DialogInformation("You cannot delete this study as it has been distributed to remote sites." & vbCrLf & "Please contact your MACRO systems administrator.")
        Exit Sub
    End If

    If DialogQuestion("Do you want to delete the study " & sClinicalTrialName & "?") <> vbYes Then
        Exit Sub
    End If
    
    'MLM 18/06/02: Use locking
    gsStudyToken = MACROLOCKBS30.LockStudy(gsADOConnectString, goUser.UserName, lClinicalTrialId)
    Select Case gsStudyToken
    Case MACROLOCKBS30.DBLocked.dblStudy, MACROLOCKBS30.DBLocked.dblEForm
        gsStudyToken = ""
        sLockDetails = MACROLOCKBS30.LockDetailsStudy(gsADOConnectString, lClinicalTrialId)
        If sLockDetails = "" Then
            DialogInformation "This study definition is currently being edited by another user. "
        Else
            DialogInformation "This study definition is currently being edited by " & Split(sLockDetails, "|")(0) & "."
        End If
        Exit Sub
    ' NCJ 13 Sept 06 - Corrected dblEForm to dblEFormInstance
    Case MACROLOCKBS30.DBLocked.dblSubject, MACROLOCKBS30.DBLocked.dblEFormInstance
        gsStudyToken = ""
        DialogInformation "Another user currently has this study open for data entry."
        Exit Sub
    Case Else
        'lock successful
    End Select

    HourglassOn
    
    DeleteTrialPRD lClinicalTrialId, sClinicalTrialName
    DeleteTrialSD lClinicalTrialId, sClinicalTrialName

    Unload frmTrialList
    
    'TA 07/04/2003: remove cache entries so reload is forced and study no longer existing discovered
    Call MarkStudyInvalid(lClinicalTrialId)
    
    'MLM 18/06/02: Remove lock corresponding to deleted study.
    ' NCJ 25 Nov 04 - Ignore errors here (in case a user has deleted the lock(!!))
    On Error Resume Next
    If gsStudyToken <> "" Then
        MACROLOCKBS30.UnlockStudy gsADOConnectString, gsStudyToken, lClinicalTrialId
        gsStudyToken = ""
    End If
    On Error GoTo ErrHandler
    
    'Display message to user
    'changed Mo Morris 10/1/00, from Me.ClinicalTrialName to sClinicalTrialName
    ' NCJ 30/3/01 - Use standard MACRO dialog box
    Call DialogInformation("Study " & sClinicalTrialName & " deleted.")
    
    HourglassOff
        
    Exit Sub
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Delete")
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
Private Sub LaunchMTMCUI()
'---------------------------------------------------------------------
' Launch the MACRO AREZZO Composer
' NCJ 3 Apr 03 - Synchronise visit cycles on our return
'---------------------------------------------------------------------
Dim sCallParameters As String
Dim nPreviousWindowState As Integer

    On Error GoTo ErrHandler
 
    'Launch the MTM Composer application

    mnuVAREZZO.Checked = True
    tlbMenu.Buttons(gsPROFORMA_LABEL).Value = tbrPressed
    
    ' Note that CLM.DLL has to be closed down prior to launching the MTMCUI,
    ' and started up again afterwards
    ' Must ensure current state of guideline is saved to PSS first
    SaveCLMGuideline
    ShutDownCLM
    
'   ATN 16/12/99 SR 2370
'   Minimize the study definition window before calling Arezzo Composer
    nPreviousWindowState = Me.WindowState
    Me.WindowState = vbMinimized

    ' Call MTM Composer and pass it the name of the trial together with
    ' the DATABASETYPE|DATABASEPATH|SERVERNAME|DATABASENAME of the current database
    'changed Mo Morris 17/1/00
    'DatabaseUser, DatabasePassword, MacroUserName and MacroPassword are
    'additionally passed as part of the sCallParameters string, for the
    'purpose of handling password protected databases (as per MACRO)
    'changed Mo Morris 27/1/00, added temp path for MTMCUI's use of IMedCLM2.dll
    'to the beginning of sCallParameters, which is now
    'gsTEMP_PATH|DATABASETYPE|DATABASEPATH|SERVERNAME|DATABASENAME|DATABASEUSER|DATABASEPASSWORD|MACROUSERNAME|MACROPASSWORD'
    'Mo Morris 10/10/01, pass DatabaseName as ServerName for Oracle databases
    If goUser.Database.DatabaseType = MACRODatabaseType.oracle80 Then
        sCallParameters = gsTEMP_PATH & "|" & goUser.Database.DatabaseType & "|" & goUser.Database.DatabaseLocation _
            & "|" & goUser.Database.NameOfDatabase & "|" & goUser.Database.NameOfDatabase _
            & "|" & goUser.Database.DatabaseUser & "|" & goUser.Database.DatabasePassword _
            & "|" & goUser.UserName & "|" & "" 'goUser.UserPassword - TA 7/1/03 doesn't need the user's password
    Else
        sCallParameters = gsTEMP_PATH & "|" & goUser.Database.DatabaseType & "|" & goUser.Database.DatabaseLocation _
            & "|" & goUser.Database.ServerName & "|" & goUser.Database.NameOfDatabase _
            & "|" & goUser.Database.DatabaseUser & "|" & goUser.Database.DatabasePassword _
            & "|" & goUser.UserName & "|" & "" 'goUser.UserPassword - TA 7/1/03 doesn't need the user's password
    End If
    
    ExecCmd "" & gsAppPath & "MTMCUI.exe " _
               & Me.ClinicalTrialName & "|" & sCallParameters & ""
    
    ' On returning to MACRO must start the CLM again...
    ' NCJ 17 Jan 03 - Pass in the study ID
    Call StartUpCLM(Me.ClinicalTrialId)
    
    ' ... and reload the current guideline back into Arezzo
    ' NCJ 17/5/01 - Use new ALM
    ' MLM/NCJ 15/01/02 Use LoadProformaTrial instead, because this also refreshes ProformaEditor.mclmTopLevelPlan
    'goALM.ArezzoFile = gpssProtocol.ArezzoFile
    LoadProformaTrial Me.ClinicalTrialName
    
    ' NCJ 3 Apr 03 - Synchronise visit cycles in case they changed any
    If SynchroniseVisitCycles(Me.ClinicalTrialId) Then
        ' They changed a visit cycle - update Schedule if it's showing
        Call RefreshSchedule
    End If
    
    mnuVAREZZO.Checked = False
    tlbMenu.Buttons(gsPROFORMA_LABEL).Value = tbrUnpressed

'   ATN 16/12/99 SR 2370
'   Return the window to its previous state
    Me.WindowState = nPreviousWindowState
    
Exit Sub
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "LaunchMTMCUI")
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
Public Sub ViewQuestions(sMode As String, lStudyId As Long, sStudyName As String)
'---------------------------------------------------------------------
' Show Data List form
'Input:
'   sMode - updatable form or read only
'   lStudyId - Study id
'   sStudyName - Study Name
'---------------------------------------------------------------------
Dim frmQuestion As frmDataList
Dim frmForm As Form
Dim nFormCount As Integer
Dim lWidth As Long
    
    On Error GoTo ErrHandler

    HourglassOn
    'count the number of datalist forms that are already visible
    nFormCount = 0
    For Each frmForm In Forms
        If frmForm.Name = "frmDataList" Then
            If frmForm.Visible Then
                If frmForm.ClinicalTrialId = lStudyId And frmForm.UpdateMode = sMode Then
                    'already showing in same mode - activate and exit
                    frmForm.SetFocus
                    HourglassOff
                    Exit Sub
                End If
                If frmForm.UpdateMode = gsREAD Then
                    nFormCount = nFormCount + 1
                End If
            End If
        End If
    Next
    
    
    Set frmQuestion = New frmDataList
    
    frmQuestion.ClinicalTrialId = lStudyId
    frmQuestion.VersionId = gnCurrentVersionId(lStudyId)
    frmQuestion.ClinicalTrialName = sStudyName
    frmQuestion.UpdateMode = sMode
        
    lWidth = (frmMenu.Width \ 4) - 50
    frmQuestion.Width = lWidth
    
    If sMode = gsUPDATE Then
        'show open study
        mnuVDataList.Checked = True
        tlbMenu.Buttons(gsDATA_ITEM_LABEL).Value = tbrPressed
        frmQuestion.Top = 0
        frmQuestion.Left = 0
        If mbAllowMU Then
            ' If MultiUser show access mode
            frmQuestion.Caption = "Question List (" & GetAccessModeString(Me.StudyAccessMode, mbAllowMU) & ")"
        Else
            ' Default to "Update"
            frmQuestion.Caption = "Question List (Update)"
        End If
    Else
        'show read only study (including library)
        frmQuestion.Top = 1000 * ((nFormCount) \ 3)  ' offset
        frmQuestion.Left = lWidth + lWidth * (nFormCount Mod 3) + (1000 * (nFormCount \ 3))  ' offset
        frmQuestion.Caption = sStudyName & " (" & GetAccessModeString(eSDAccessMode.sdReadOnly) & ")"
        If lStudyId = 0 Then
            'Show library
            mnuVLibrary.Checked = True
            tlbMenu.Buttons(gsLIBRARY_LABEL).Value = tbrPressed
        End If
    End If
    
    With frmMenu
        frmQuestion.Height = (.Height - .tlbMenu.Height - .sbrMenu.Height) - 750
    End With
    
    frmQuestion.Show
          
    HourglassOff

Exit Sub

ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "ViewQuestions", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If

End Sub

'---------------------------------------------------------------------
Public Sub HideQuestions(sMode As String, lStudyId As Long, bFromMenu As Boolean)
'---------------------------------------------------------------------
' Unload Data List form
'Input:
'   sMode - updatable form or read only
'   lStudyId - Study id
'   bFromMenu - from menu rather than clicking close
'---------------------------------------------------------------------

Dim frmForm As Form

    On Error GoTo ErrHandler
    
    If sMode = gsREAD And lStudyId = 0 Then
        'closed read only library
        mnuVLibrary.Checked = False
        tlbMenu.Buttons(gsLIBRARY_LABEL).Value = tbrUnpressed
    End If
    
    If sMode = gsUPDATE And lStudyId = Me.ClinicalTrialId Then
        'closed open study form
        mnuVDataList.Checked = False
        tlbMenu.Buttons(gsDATA_ITEM_LABEL).Value = tbrUnpressed
    End If
    
    If bFromMenu Then
        'if close wasn't clicked
        For Each frmForm In Forms
            If frmForm.Name = "frmDataList" Then
                If frmForm.ClinicalTrialId = lStudyId And frmForm.UpdateMode = sMode Then
                    Unload frmForm
                    Exit For
                End If
            End If
        Next
    End If
    
    'Mo Morris 26/2/99 SR 694, dummy call to change menu options and clear the
    'currently selected item
    ChangeSelectedItem "", ""
    
    
Exit Sub

ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "HideQuestions", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If

End Sub


'---------------------------------------------------------------------
Public Function ShowPopup(sOption As String, Optional sEnabledChecked As String = "") As Long
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
            Load Me.mnuPopUpsubItem(i)
        End If
        sStatus = vEnabledChecked(i)
        With mnuPopUpsubItem(i)
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
        Unload mnuPopUpsubItem(i)
    Next
    
    'return user's choice
    ShowPopup = mlPopUpItem
    
End Function

'----------------------------------------------------------------------------
Public Sub DeleteUnusedQuestions(ByVal lClinicalTrialId As Long, ByVal nVersionId As Integer)
'----------------------------------------------------------------------------
'Deletes unused questions from the appropriate tables in the database
'REM 06/03/02 - assigned recordset to an array then loop through array to delete unused questions
'----------------------------------------------------------------------------
Dim rsDelete As ADODB.Recordset
Dim sSQL As String
Dim vData As Variant
Dim i As Integer
        
    On Error GoTo ErrHandler

    'REM 10/01/02 - 'returns recorset of unused questions, i.e. questions not on eforms or in question groups
    Set rsDelete = New ADODB.Recordset
    Set rsDelete = UnusedQuestionList(lClinicalTrialId, nVersionId)
    
    'checks if records exist
    If rsDelete.RecordCount > 0 Then
   
        'REM 06/03/02 - assign recordset to an array
        vData = rsDelete.GetRows
        rsDelete.Close
        Set rsDelete = Nothing

    
        'looping thru array to perform delete
        For i = 0 To UBound(vData, 2)
            'deletes from Arezzo file
            DeleteProformaDataItem vData(0, i)
            'deletes from appropriate tables
            Call gdsDeleteDataItem(lClinicalTrialId, nVersionId, CLng(vData(0, i)))
        Next
        
        'refreshes question list / treeview
        Call RefreshQuestionLists(lClinicalTrialId)
        DialogInformation "All unused questions deleted"
    Else
        rsDelete.Close
        Set rsDelete = Nothing
    
    End If

    ' NCJ 19 Jun 06 - Mark study as changed
    Call MarkStudyAsChanged
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "DeleteUnusedQuestions", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If

End Sub

'---------------------------------------------------------------------
Public Sub RefreshQuestionLists(ByVal lClinicalTrialId As Long)
'---------------------------------------------------------------------
' ASH 18/07/2001
' Refresh ALL question lists currently showing for this Study
' This was copied from frmdataDefinition to help refresh question list
' (See also RefreshDatalistWindows)
' NCJ 14 Jan 02 - Routine made public so that it can be called
'           whenever data lists need updating
' ASH 21/06/2002 Bug 2.2.16 no.11
'---------------------------------------------------------------------
Dim oForm As Form
     
    On Error GoTo ErrLabel
    
    For Each oForm In Forms   'iterate through the forms collection
        If oForm.Name = "frmDataList" Then
            ' Check its trial id
            If oForm.ClinicalTrialId = lClinicalTrialId Then
                oForm.RefreshDataList
            End If
        End If
    Next

    Set oForm = Nothing
    
    'ASH 21/06/2002 Bug 2.2.16 no.11
    Call EnableUnusedQuestionsMenu(mlClinicalTrialId, mnVersionId)
    'REM 19/07/02
    Call EnableUnusedQGroupsMenu(mlClinicalTrialId, mnVersionId)

Exit Sub
ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                                "RefreshQuestionLists", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub MarkStudyInvalid(ByVal lClinicalTrialId As Long)
'---------------------------------------------------------------------
'Delete all rows that belongs to this Study as user might change this
'Study making the same Study loaded by Subject Cache Manager invalid.
'---------------------------------------------------------------------

    MacroADODBConnection.Execute "Delete from ArezzoToken where ClinicalTrialID = " & lClinicalTrialId

End Sub

'------------------------------------------------------------------------------
Public Sub EnableUnusedQuestionsMenu(ByVal lClinicalTrialId As Long, _
                                    ByVal nVersionId As Integer)
'------------------------------------------------------------------------------
''ASH 20/06/2002 bug 2.2.16 no.11
' NCJ 10 May 06 - Look at StudyAccess too
'------------------------------------------------------------------------------

    If frmDataList.DoUnusedQuestionsExist(mlClinicalTrialId, mnVersionId) _
                And goUser.CheckPermission(gsFnDelQuestion) And meStudyAccess = sdFullControl Then
        mnuEUnusedQuestions.Enabled = True
    Else
        mnuEUnusedQuestions.Enabled = False
    End If

End Sub

'------------------------------------------------------------------------------
Public Sub EnableUnusedQGroupsMenu(ByVal lClinicalTrialId As Long, _
                                    ByVal nVersionId As Integer)
'------------------------------------------------------------------------------
' REM 19/07/02
' NCJ 10 May 06 - Look at StudyAccess too
'------------------------------------------------------------------------------

    If frmDataList.DoUnusedQGroupsExist(lClinicalTrialId, nVersionId) _
                And meStudyAccess = eSDAccessMode.sdFullControl Then
        mnuEDeleteUnusedQGroups.Enabled = True
    Else
        mnuEDeleteUnusedQGroups.Enabled = False
    End If
    

End Sub

'------------------------------------------------------------------------------
Private Sub ClearOptionChoices()
'------------------------------------------------------------------------------
' ZA 09/09/2002 - finds and uncheck a checked sub menu for using option buttons
'------------------------------------------------------------------------------
Dim n As Integer
    
    For n = 0 To mnuCatValues.Count - 1
        mnuCatValues(n).Checked = False
    Next n
End Sub

'------------------------------------------------------------------------------
Private Sub LoadMACROSettingMenus()
'------------------------------------------------------------------------------
' ZA 01/09/2002 - loads values from MACROSetting files to set up menus
'------------------------------------------------------------------------------
Dim enAutoNumbering As eAutoNumbering
Dim enHideQuestionStatus As eStatusFlag
Dim enHideRQGStatus As eStatusFlag
Dim enRFCDefault As eRFCDefault
    
    On Error GoTo ErrLabel
    
    'get the value of Automatic numbering
    enAutoNumbering = CInt(GetMACROSetting(gs_AUTOMATIC_NUMBERING, eAutoNumbering.NumberingOff))
    'check this menu item only if value is 1
    If enAutoNumbering = eAutoNumbering.NumberingOn Then
        mnuAutomaticNumbering.Checked = True
    Else
        mnuAutomaticNumbering.Checked = False
    End If
    
    'get the value of hide question status icon
    enHideQuestionStatus = CInt(GetMACROSetting(gs_HIDE_QUESTION_STATUS_ICON, eStatusFlag.Show))
    'check this menu item only if the value is 1
    If enHideQuestionStatus = eStatusFlag.Show Then
        mnuHideIcons.Checked = False
    Else
        mnuHideIcons.Checked = True
    End If
    
    'get the value of RQG status icon
    enHideRQGStatus = CInt(GetMACROSetting(gs_HIDE_RQG_STATUS_ICON, eStatusFlag.Hide))
    'check this menu item if the value is 1
    If enHideRQGStatus = eStatusFlag.Hide Then
        mnuHideRQGIcons.Checked = True
    Else
        mnuHideRQGIcons.Checked = False
    End If
    
    'get the value for using option buttons on CRF page
    'the default value is 3 e.g. use option buttons if the items are <= 3
    gnUseOptionButton = CInt(GetMACROSetting(gs_USE_OPTION_BUTTONS, "3"))
    
    'clear all Option buttons menu items
    ClearOptionChoices
    
    Select Case gnUseOptionButton
        Case 1 To 9
            'the index for selected menu item is based on caption
            mnuCatValues(gnUseOptionButton - 1).Checked = True
        Case gn_ALWAYS_USE_OPTION_BUTTONS
            'Always use option buttons
            mnuCatValues(gn_ALWAYS_USE_OPTION_MENU).Checked = True
        Case gn_NEVER_USE_OPTION_BUTTONS
            'Never use option buttons
            mnuCatValues(gn_NEVER_USE_OPTION_MENU).Checked = True
        Case Else
            'if nothing found, set this value to 3
            mnuCatValues(gn_DEFAULT_OPTION_MENU_VALUE).Checked = True
        
    End Select
    
    'get the value of RFC
    enRFCDefault = CInt(GetMACROSetting(gs_DEFAULT_RFC, "1"))
    'check this menu item only if the vlaue is 1
    If enRFCDefault = eRFCDefault.RFCDefaultOn Then
        mnuDefaultRFC.Checked = True
    Else
        mnuDefaultRFC.Checked = False
    End If
    
    Exit Sub
    
ErrLabel:
     Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, _
                            "LoadMACROSettingMenu", Err.Source)
    Case OnErrorAction.Retry
        Resume
    End Select
End Sub


'---------------------------------------------------------------------
Public Function ForgottenPassword(sSecurityCon As String, sUsername As String, sPassword As String, ByRef sErrMsg As String) As eDTForgottenPassword
'---------------------------------------------------------------------
'REM 06/12/02
'---------------------------------------------------------------------

    'dummy routine

End Function

'----------------------------------------------------------------------------
Private Function CheckStudyOK() As Boolean
'----------------------------------------------------------------------------
' NCJ 31 Jan 03 - Check that the study is OK before closing it
' We check that there is a schedule, and then that there are no infinitely cycling visits
' Returns FALSE if we should NOT close the study
'----------------------------------------------------------------------------
Dim bOK As Boolean
    
    bOK = True
    ' Check the Schedule
    bOK = ScheduleOK
    If bOK Then
        ' Check for infinitely cycling visits
        bOK = CheckStudyHoldsWater
    End If
    CheckStudyOK = bOK
        
End Function

'----------------------------------------------------------------------------
Public Function ScheduleOK(Optional bShowMessage As Boolean = True) As Boolean
'----------------------------------------------------------------------------
' NCJ 31 Jan 03 - Check that the study contains something in its schedule
' A dialog will be shown asking if the user wants to close if bShowMessage = true
'----------------------------------------------------------------------------
Dim bOK As Boolean
Dim sSQL As String
Dim rsVEFs As ADODB.Recordset
Dim sMSG As String

    bOK = True
    ' Only check if a study is open
    ' NCJ 10 May 06 - And if study is RW
    If Me.ClinicalTrialId > 0 And meStudyAccess >= sdReadWrite Then
        ' See if there are any eForms in Visits
        sSQL = "SELECT COUNT(*) FROM StudyVisitCRFPage WHERE ClinicalTrialId = " & Me.ClinicalTrialId
        Set rsVEFs = New ADODB.Recordset
        rsVEFs.Open sSQL, MacroADODBConnection
        bOK = (rsVEFs.Fields(0) > 0)
        If (Not bOK) And bShowMessage Then
            sMSG = "This study's Schedule is empty, and this study cannot be used for data entry."
            sMSG = sMSG & vbCrLf & vbCrLf & "Are you sure you wish to close this study?"
            bOK = (DialogWarning(sMSG, , True) = vbOK)
        End If
        rsVEFs.Close
        Set rsVEFs = Nothing
    End If
    
    ScheduleOK = bOK

End Function

'----------------------------------------------------------------------------
Public Function IsQuestionOnLockedForm(lDataItemId As Long) As Boolean
'----------------------------------------------------------------------------
' NCJ 12 Jun 06 - Is this question on an eForm that's currently locked by another user?
'----------------------------------------------------------------------------
Dim colQEForms As Collection
Dim colLockedEForms As Collection
Dim bOnLockedEForm As Boolean
Dim vQCRFPageId As Variant
Dim vLCRFPageId As Variant

    On Error GoTo ErrLabel
    
    bOnLockedEForm = False
    
    ' Get all the EForms this question is on
    Set colQEForms = CRFPagesForDataItem(mlClinicalTrialId, mnVersionId, lDataItemId)
    If colQEForms.Count > 0 Then
        ' Get the eForms that are locked by another user
        Set colLockedEForms = MACROLOCKBS30.LockedEForms(gsADOConnectString, goUser.UserName, mlClinicalTrialId)
        For Each vQCRFPageId In colQEForms
            For Each vLCRFPageId In colLockedEForms
                If CLng(vQCRFPageId) = CLng(vLCRFPageId) Then
                    bOnLockedEForm = True
                    Exit For
                End If
            Next
            If bOnLockedEForm Then Exit For     ' Don't bother continuing if we have a result
        Next
        Set colLockedEForms = Nothing
    End If
    
    Set colQEForms = Nothing
    
    IsQuestionOnLockedForm = bOnLockedEForm
  
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.IsQuestionOnLockedForm"
    
End Function

'----------------------------------------------------------------------------
Public Function IsRQGOnLockedForm(lQGroupId As Long) As Boolean
'----------------------------------------------------------------------------
' NCJ 21 Sept 06 - Is this RQG on an eForm that's currently locked by another user?
'----------------------------------------------------------------------------
Dim colQEForms As Collection
Dim colLockedEForms As Collection
Dim bOnLockedEForm As Boolean
Dim vQCRFPageId As Variant
Dim vLCRFPageId As Variant

    On Error GoTo ErrLabel
    
    bOnLockedEForm = False
    
    ' Get all the EForms this question is on
    Set colQEForms = CRFPagesForQGroup(mlClinicalTrialId, mnVersionId, lQGroupId)
    If colQEForms.Count > 0 Then
        ' Get the eForms that are locked by another user
        Set colLockedEForms = MACROLOCKBS30.LockedEForms(gsADOConnectString, goUser.UserName, mlClinicalTrialId)
        For Each vQCRFPageId In colQEForms
            For Each vLCRFPageId In colLockedEForms
                If CLng(vQCRFPageId) = CLng(vLCRFPageId) Then
                    bOnLockedEForm = True
                    Exit For
                End If
            Next
            If bOnLockedEForm Then Exit For     ' Don't bother continuing if we have a result
        Next
        Set colLockedEForms = Nothing
    End If
    
    Set colQEForms = Nothing
    
    IsRQGOnLockedForm = bOnLockedEForm
  
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.IsRQGOnLockedForm"
    
End Function

'----------------------------------------------------------------------------
Public Sub MarkStudyAsChanged()
'----------------------------------------------------------------------------
' NCJ 19 Jun 06 - Mark study as changed (so other users know to refresh)
'----------------------------------------------------------------------------

    ' Make sure we leave our own token untouched
    Call MACROLOCKBS30.CacheInvalidateStudy(gsADOConnectString, mlClinicalTrialId, msCacheToken)
    
End Sub

'----------------------------------------------------------------------------
Public Function StudyHasChanged() As Boolean
'----------------------------------------------------------------------------
' NCJ 19 Jun 06 - Has another user changed the study? (Do we need to refresh?)
' If so, get ourselves a new cache token
'----------------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    If msCacheToken > "" And MACROLOCKBS30.CacheEntryStillValid(gsADOConnectString, msCacheToken) Then
        ' OK
        StudyHasChanged = False
    Else
        StudyHasChanged = True
        ' Check the study hasn't been deleted
        
        ' Get ourselves a new token
        msCacheToken = MACROLOCKBS30.CacheAddStudyRow(gsADOConnectString, mlClinicalTrialId)
        DialogInformation "Another user has made changes to this study, and the display will now be updated."
    End If
   
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.StudyHasChanged"

End Function

'-----------------------------------------------------------------------
Private Sub mnuFExport_Click()
'-----------------------------------------------------------------------
' NCJ 26 Jun 06 - Issue 2744 - Allow Export of study
'-----------------------------------------------------------------------

    Call SaveExportFile
    
End Sub

'----------------------------------------------------------------------------
Private Function SaveExportFile() As String
'----------------------------------------------------------------------------
' NCJ 26 Jun 06 - Issue 2744 - Save study as export file to selected location
'----------------------------------------------------------------------------
Dim oExchange As clsExchange
Dim sFileSpec As String

    On Error GoTo ErrHandler
    
    With CommonDialog1
        .CancelError = True
        .DialogTitle = "Save study export file"
        .Filter = "Export (*.zip)|*.zip"
        .FilterIndex = 1
        .InitDir = gsOUT_FOLDER_LOCATION
        .FileName = msClinicalTrialName & "_" & Format(Now, "yyyymmddhhmm")
        .Flags = cdlOFNPathMustExist + cdlOFNOverwritePrompt
    End With
    ' Ignore "cancel" error
    On Error GoTo CancelError
    
    CommonDialog1.ShowSave
    
    On Error GoTo ErrHandler
    
    sFileSpec = CommonDialog1.FileName
    
    Set oExchange = New clsExchange
    Call oExchange.ExportNamedSDD(mlClinicalTrialId, msClinicalTrialName, mnVersionId, msClinicalTrialName, sFileSpec)
    Set oExchange = Nothing
    
Exit Function
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "SaveExportFile", Err.Source) = OnErrorAction.Retry Then
        Resume
    End If
CancelError:
    ' Do nothing if they cancel the File Dialog
End Function
