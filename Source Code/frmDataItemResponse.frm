VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDataItemResponse 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Browser"
   ClientHeight    =   3765
   ClientLeft      =   7395
   ClientTop       =   4635
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3765
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picStatusIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   1140
      ScaleHeight     =   255
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.ImageList imgStatusIcons 
      Left            =   120
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   22
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":0000
            Key             =   "K0old"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":036C
            Key             =   "K1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":06FA
            Key             =   "K2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":0AB0
            Key             =   "K3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":0E43
            Key             =   "K4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":11E7
            Key             =   "K5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":1570
            Key             =   "K7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":1900
            Key             =   "K9"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":1CB1
            Key             =   "K6"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":2423
            Key             =   "K10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":27B0
            Key             =   "K11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":2B49
            Key             =   "K12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":2F14
            Key             =   "K13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":32E5
            Key             =   "K16"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":365E
            Key             =   "K32"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":39E0
            Key             =   "K48"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":3D6D
            Key             =   "K64"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":4103
            Key             =   "K128"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":449E
            Key             =   "Old256"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":452F
            Key             =   "K256"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":48A1
            Key             =   "K512"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataItemResponse.frx":4C1B
            Key             =   "K768"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid flexData 
      Height          =   2325
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4101
      _Version        =   393216
      BackColorBkg    =   -2147483643
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   3
      GridLinesFixed  =   1
      SelectionMode   =   1
      MergeCells      =   4
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lcmdPrint 
      Caption         =   "Print"
      Height          =   240
      Left            =   3360
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   120
      Width           =   675
   End
   Begin VB.Label lcmdClose 
      Caption         =   "Close"
      Height          =   240
      Left            =   4140
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   120
      Width           =   675
   End
   Begin VB.Label lblTotals 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblCommentSize 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   3180
      Width           =   1680
   End
End
Attribute VB_Name = "frmDataItemResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmDataItemResponse.frm
'   Author:     Mo Morris July 1997
'   Purpose:    Allows user to enter selection criteria and
'               display sets of data, including audit information.  Can be called in monitor mode (which allows cross-trial selection), or for a single patient.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
 '  1 - 18  Mo Morris, Andrew Newbigging 28/07/97 to 17/07/98
'   21  Mo Morris               11/9/98     SPR 426
'                                           Changes to Public Sub PrintData & PrintForms.
'                                           Cancel Print (cdlCancel) now detected
'   22  Mo Morris               22/9/98     SPR 435
'                                           When switching from individual Patient Data view to
'                                           Monitor view all of the forms combos have to be re-loaded.
'       Mo Morris               24/11/98    SPR 587
'                                           PrintData and PrintForms changed to handle all errors
'                                           generated from a CommonDialog.ShowPrinter, including Cancel
'       Mo Morris               11/1/99     SR 628
'       When displaying forms as opposed to data items 4 of flexData's columns become hidden by
'       having their width set to 0. When adjusting flexData's column widths by dragging the
'       sides of the column headers it is possible to widen one of these 0 width columns.
'       Unfortunately there does not appear to be a place where this can be prevented from happening,
'       so a check has been placed in flexData_MouseMove, the event most likely to occur afterwards.
'       Andrew Newbigging       26/4/99     SR 483
'       Added a new column to store index values of the record.
'       Added functionality to add new comments and swith to a data entry form.
'       Added explicit ordering on combo boxes and main list view.
'       Added record count in the status bar.
'   PN  26/09/99    Amended PopulateList(), txtDate_lostfocus() for mtm1.6 changes
'  WillC    11/11/99    Added the error handlers
'   NCJ 17/11/99    Localise number values before printing/displaying
'   Mo Morris   18/11/99    DAO to ADO conversion
'   NCJ 30/11/99 - Calculate CRF Page Date in ShowForm
'                   Use Doubles for dates
'   NCJ 11/12/99 - Changed clsLockedFrozenCoords to clsCoords
'  WillC   11/12/99
'          Changed the Following where present from Integer to Long  ClinicalTrialId
'           CRFPageId,VisitId,CRFElementID
'   SDM November - December
'                           Redesigned the data browser. Changes include the selection
'                           of trials, visits, forms and data items. Updating the new
'                           lock and freeze so that items can be OK, missing, etc and
'                           still be locked or frozen. Colour coded the grid to allow
'                           the user to easily identify what is lock and frozen. Also
'                           adjusted the positioning of the popup menu when the right
'                           mouse button is pressed. Adjusted the page layout so that
'                           the search criteria could be seen when viewed in 800x600.
'   NCJ 16 Dec 99 - Disallow Locking/Freezing if user is not authorised
'   NCJ 12 Jan 00 - Build schedule when opening form from Data Browser
'   NCJ 21 Jan 00 - SR2741 Block right mouse menu in Read mode
'                   SR2726 Removed Error check box
'   NCJ 2 Feb 00 - Set up schedule's Visit ID in ShowForm
'   NCJ 4 Feb 00 - Make sure schedule is refreshed in ShowForm
'                   Changed routine names in error handlers from "ShowForm"
'   ATN 16 Feb 00 - PopulateImageList uses 16x16 bitmaps from the resource file
'   NCJ 1/3/00 SR 3121 - Major rewriting of setting Lock/Freeze/Unlock
'               Changed all Row & Column variables from Integer to Long
'   NCJ 2/3/00 - Removed unused commented code
'   NCJ 3/3/00 SR3125 - Changed status menu options to work correctly
'               SR3169 - New AdjustRowHeight routine
'   Mo Morris   6/3/00  SR 3016. cboUser is now populated with user codes via calls to
'   LoadUserCombo and LoadUserComboMonitor. cboUser_Click now sets new mudule based
'   variable msComboSelectedUser and PopulateList uses it when constructing its SQL
'   statement. bRecordCountIsReallyZero added to populateList and is used to make sure that
'   an erroneous recordcount of 1 is not displayed.
'   NCJ 9/3/00  SR3190 Use LaunchMode property to determine "Monitor" mode or "Subject" mode
'       Make sure that Clicks are always done before double clicks
'   Mo Morris   16/3/00 SR1185
'   PopulateList changed so that it now places the changed status (NoChange, Changed or Imported)
'   into the new field.
'   Mo Morris   17/3/00 SR 3016     cboUser chaged from style 0-dropdown combo to style 2-dropdown list
'   Mo Morris   23/3/00 SR 3261     See PrintData for details
'   NCJ 21/4/00 - Alterations to ChangeStatusInBrowser
'                   Use HourglassOn and HourglassOff instead of mousepointer
'                   Removed imgOK, imgWarning and imgInform
'                   Show inform icons if user has correct permissions
'   NCJ 25/4/00 - SR3239 Implemented hierarchical unlocking
'               Commented out subclassing calls
'   Mo Morris 26/4/00 changes made to ChangeLockUnlockFreeze arround creating REmote Site Lock Messages
'   TA 28/04/2000 SR3276: prevent users selecting column headers
'   TA 3/5/2000 SR3398 Moved disabling frames codes in PopulateList later in code so they are disabled after initial checks
'   NCJ 9/5/2000 SR3405 Now correctly handle frozen/locked colourings in Audit Trail view
'               SR3429 Changed tooltip text for "New" checkbox
'   NCJ 19/5/00 New MIMessage handling from right-mouse menu
'   NCJ 31/5/00 SR3522 Ensure that ChangeData and AddComment permissions go together
'               Removed unused EnableQuestionStatusMenu
'   TA 05/06/2000: Launchmode property now read/write
' NCJ 15/6/00 SR3611 Must include VisitCycleNumber in SQL joins in PopulateList
'   Mo Morris 13/6/00,  Performance changes
'               new function TooManyRecords created
'   NCJ 1/9/00 - Made changes in TooManyRecords and PopulateList to match 2.0.41
'   NCJ 3/10/00 - SR 3311 Removed mnuStatusMissing and mnuStatusUnobtainable and ChangeStatusInBrowser
'   NCJ 6/10/00 - Added NRStatus and CTCGrade text to Status column
'   NCJ 17/10/00 SR3864 - Added Validation message column
'                SR3870 - Order rows by VisitOrder, CRFPageOrder and FieldOrder rather than by name
'   NCJ 18/10/00 SR3866 - Replace cboPerson with text field for subject label
'   NCJ 30/10/00 - Set mlSelectedPersonID in Subject mode, but set it to 0 in Monitor mode
'   NCJ 6/11/00 - Make sure row height is never 0! (see AdjustRowHeight)
'   Mo  1/12/00     The following changes have been made to the data browser printed listing:-
'               The layout has been adapted for use on Letter as well as A4 paper
'               The Print date nad page are now right justified in the header
'               The data item Response no longer overlaps fields to its left
' NCJ 16 Jan 01 - Create private visits/pages/dataitems collections
'               for the single selected study
' TA  19 Jan 01 - sql for populate list now in clsDataReviewSQL
'                   (old code can be found in Projects\Macro\Release 2.1\copy of frmDataItemResponse 19-1-2001)
'   NCJ 30 Jan 01 - In PopulateList only adjust row height ONCE per row
'   TA  31 Jan 01 - New progress bar used to improve performance
'   TA  31 Jan 01 - Tidied up PopulateList
'   TA 09/05/2001 - New grid population subroutine as used for Roche 2.0
'   TA 8/8/2001 - DataBrowser object not passed statuses as a string
'   TA 13/08/2001 - Printing now uses the mvData array
'   NCJ 25 Sep 01 - Use goArezzo instead of frmArezzo for MACRO 2.2
'   TA Sept 01 - Changes to incorporate the new business objects
'   NCJ 2 Oct 01 - Removed reference to frmStudyVisits
'   TA 03/10/2001: Sites in combo and subject data return is now linited by the site user table
'   DPH 11/01/2002 - Comment out Add Comment functionality in EnableStatusMenu (for V-Track temporarily)
'   DPH 17/1/2002 - Add Comment functionality put back in for Macro v2.2.8
'   ZA 23/04/04 - Changed cboUser style for list to a combo, allowing user to type user name
'   ZA 24/04/2002 - Copied IsValidString function, added change event for cboUser
'   Mo 24/5/2002 -  (Stemming from CBB 2.2.8.7) Changes made to SplitTextLine so that it now
'                   splits a strint that does not contain any spaces (instead of recursively
'                   calling itself until there is a call stack overflow)
'   DPH 25/5/2002   UlockSubject call added to ChangeLockUnLockFreeze
'   TA 17/07/2002   Added Close button
'   ATO 20/08/2002  Changes made on ChangeLockUnlockFreeze (Added RepeatNumber)
'   TA 26/09/02: Changes for New UI - no title bar, not maximised , added close button
'   TA 02/10/2002: Some of populatelist put in SetUpGrid so that it can be called through DisplayNew when vData is already known
'    TA 02/10/2002: IMEDNow used instead of Now in QNewMIMessage
'   TA 02/10/2002: ChangeCommentInBrowser is a bit dangerous to be done without going through the business layer
'                        - so is commented out until further notice
'   NCJ 16 Oct 02 - Minor changes for new SDV/MIMsg handling
'   NCJ 17 Oct 02 - New RightMouseMenu routine; removed some unused code
'   TA 01/11/2002 - Changed to new icons
'   NCJ 5 Nov 02 - Do not allow more than one SDV per object
'   TA 18/11/2002: removed code that used index stored in hidden column in grid - now use mvData directly
'   NCJ 23/24 Dec 02 - New Locking & Freezing stuff
'   NCJ 9 Jan 03 - Must refresh Data Browser when doing Unfreezing
'   TA 19/01/2003: Subject locking now done when creating mimessages
'   RS 22/01/2003: Generate combined status Icon, for each possible status combination
'   RS 13/02/3004: Reset the min/max click values when drawing new grid
'   IC 14/02/2003  bug 807, remove '/Label' from column header (Form_Load)
'                           display label or id in brackets (PopulateGridSection)
'                           remove text on statuses (PopulateGridSection)
'   RS 18/02/2003: Adjust Rowheight by 1.22 for long comments
'   RS 12/03/2003: Adjust Rowheight for single-row-subjects to make sure label is displayed correctly
'                  Adjust RowWidth for some columns to display more columns
'   NCJ 20 Mar 03 - DisplayNew can now take Null data
'                   Added some calls to RestartSystemIdleTimer at appropriate places
'   TA 26/03/2003: Added database timestamp and database time zone columns
'   Mo 14/5/2003    The following changes have been made to the data browser printed listing:-
'                   The Transfer field (Not exported/New/Exported) has been removed from the listing.
'                   The printed Date & Time now includes the timezone + offset.
'                   User Name field expanded to take 20 character User Names.
'                   Comment field reduced in width.
'                   A Key of User Name Codes - Full User Names now appears at the end of the listing.
'   TA 22/05/2003:  MOivedthe formating of GTM timezone offsets into own subroutine
'   TA 27/05/2003:  Added eFormLabel to display and speeded and PopulateGrid so make scrolling somoother
'   MLM 04/06/03: 3.0 buglist 731: Replaced K6 in imgStatusIcons to fix white bg in OKWarnings.
'   MLM 06/06/03: 3.0 buglist 1799: Left hand columns stop growing when they reach a certain size, to
'                   ensure they fit on 800x600 screen.
'   DPH 23/09/2003 - keep and use a collection of drawn images to improve performance
'   NCJ 24 Aug 04 - Changed gsFnLockData to gsFnUnLockData in "UNLOCK" right mouse menu (Bug 2368)
'   TA 08/11/2004: CBD 2425 CRM 990 - reinstated max col width check for response value so the row height sizing works
'   Mo 16/11/2004   Bug 2417 - Identical eForm Labels printing problem fixed
'   ic 28/07/2005   added clinical coding
'------------------------------------------------------------------------------------

Option Explicit
Option Base 0
Option Compare Binary

'constants for grid cols
Private Const mnTRIALSITEPERSON_COL = 0
Private Const mnVISIT_COL = 1
Private Const mnFORM_COL = 2
Private Const mnDATAITEM_COL = 3
Private Const mnDATARESPONSE_COL = 4
Private Const mnSTATUS_COL = 5
Private Const mnNEW_COL = 6 'SDM SR1185 30/11/99
Private Const mnTIMESTAMP_COL = 7
'TA 26/03/2003: db time stamp column
Private Const mnDB_TIMESTAMP_COL = 8
Private Const mnUSERID_COL = 9
Private Const mnUSERNAMEFULL_COL = 10
Private Const mnCOMMENT_COL = 11
Private Const mnRFC_COL = 12    'SDM SR2408
Private Const mnOVERRULE_COL = 13    'NCJ - SR 3454
Private Const mnVAL_MESSAGE_COL = 14    'NCJ 17/10/00 SR3864, Validation message

'ic 28/07/2005 added clinical coding columns
Private Const mnDICTIONARY_COL = 15
Private Const mnCODINGSTATUS_COL = 16
Private Const mnCODINGDETAILS_COL = 17

Private Const mnNUM_OF_CCCOLS = 3

Private Const mnNUM_OF_COLS = 15


Private Const msINDEX_SEPARATOR = ","

Private Const mnLOCKED_COLOUR = &H80&
Private Const mnFROZEN_COLOUR = &H808000

' Store row and column clicked
Private mlColumnClicked As Long
Private mlRowClicked As Long
' Store range of rows covered by click
Private mlRowClickedMin As Long
Private mlRowClickedMax As Long

' NCJ 29/2/00 - Changed to use new clsColCoords
' Each collection class contains the coordinates of cells that are locked or frozen
Private mcolLocked As clsColCoords
Private mcolFrozen As clsColCoords

' NCJ 9/3/00 - Store our "launch mode", i.e. either Monitor or Subject
Private mnLaunchMode As eMACROWindow

'Arrays to hold recordset data and whether a row has been shown
Dim mvData As Variant
Dim mvRowShown As Variant

'number of rows to put in grid in one go
Private Const m_ROW_BUFFER = 100

'query type when refresh was last pressed
Private mDRType As eDataBrowserType

' RS Store index of last populated row
Private mlLastPopulatedRow As Long

' DPH 23/09/2003 - Added collection to hold images
Private mcolDrawnImages As Collection

'ic 28/07/2005 form-wide collection of dictionaries
Private moDictionaries As MACROCCBS30.Dictionaries

'---------------------------------------------------------------------
Private Sub cmdCloseForm_Click()
'---------------------------------------------------------------------
'unload form
'---------------------------------------------------------------------
    
    Unload Me

End Sub

'---------------------------------------------------------------------
Private Sub flexData_Click()
'---------------------------------------------------------------------
'   SDM 01/12/99
'   Highlights the grid
'   NCJ 9/5/00 - Changed numbers to column constant values
'---------------------------------------------------------------------
Dim lEnteredCol As Long
Dim lEnteredRow As Long
Dim lRowCount As Long
Dim oCoordinates As clsCoords
Dim cntRow As Long
Dim lRowClickedMin As Long
Dim lRowClickedMax As Long

    On Error GoTo ErrHandler
    
    With flexData
        'MLM 27/06/03: Record the position of the mouse as early as possible,
        'otherwise the wrong thing will be selected if the user moves their mouse around after clicking.
        lRowClickedMin = mlRowClickedMin
        lRowClickedMax = mlRowClickedMax
        mlColumnClicked = .MouseCol
        mlRowClicked = .MouseRow
        mlRowClickedMin = .MouseRow
        mlRowClickedMax = .MouseRow
        lEnteredCol = .MouseCol
        lEnteredRow = .MouseRow
        .FillStyle = flexFillRepeat

        ' NCJ 20 Mar 03 - Reset timer
        Call RestartSystemIdleTimer
        
        ' If no data showing, do nothing
        If flexData.Rows = 1 Then Exit Sub
        
        .Visible = False
        
        ' RS 23/01/2003 Redraw icons of selected items
        If lRowClickedMin > 0 Then
            For cntRow = lRowClickedMin To lRowClickedMax
               .Row = cntRow
               ' picStatusIcon.BackColor = vbWindowBackground
               Call SetIcon(mnTRIALSITEPERSON_COL, .Row, vbWindowBackground)
               Call SetIcon(mnVISIT_COL, .Row, vbWindowBackground)
               Call SetIcon(mnFORM_COL, .Row, vbWindowBackground)
               Call SetIcon(mnSTATUS_COL, .Row, vbWindowBackground)
            Next
        End If
        
        'Deselect all cells - set to default black unselected
        
        .Row = 1
        .Col = 0
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .CellBackColor = vbWindowBackground
        .CellForeColor = vbWindowText
        
        ' Now reset the colour of all the locked items
        For Each oCoordinates In mcolLocked
            .Col = oCoordinates.Col
            .Row = oCoordinates.Row
            .CellForeColor = mnLOCKED_COLOUR
        Next oCoordinates

        ' And reset the colour of all the frozen items
        For Each oCoordinates In mcolFrozen
            .Col = oCoordinates.Col
            .Row = oCoordinates.Row
            .CellForeColor = mnFROZEN_COLOUR
        Next oCoordinates
        
        Select Case lEnteredCol
            Case mnTRIALSITEPERSON_COL
                ' Trial subject column
                For lRowCount = 1 To .Rows - 1
                    If .TextMatrix(lRowCount, mnTRIALSITEPERSON_COL) = .TextMatrix(lEnteredRow, mnTRIALSITEPERSON_COL) Then
                        'Select
                        .Row = lRowCount
                        .Col = mnTRIALSITEPERSON_COL
                        .ColSel = .Cols - 1
                        .CellBackColor = vbHighlight
                        .CellForeColor = vbHighlightText
                        If mlRowClickedMin > .Row Then
                            mlRowClickedMin = .Row
                        End If
                        If mlRowClickedMax < .Row Then
                            mlRowClickedMax = .Row
                        End If
                        ' Set Icons in all columns
                        'picStatusIcon.BackColor = vbHighlight
                        Call SetIcon(mnTRIALSITEPERSON_COL, .Row, vbHighlight)
                        Call SetIcon(mnVISIT_COL, .Row, vbHighlight)
                        Call SetIcon(mnFORM_COL, .Row, vbHighlight)
                        Call SetIcon(mnSTATUS_COL, .Row, vbHighlight)
                    End If
                Next lRowCount
            Case mnVISIT_COL
                ' Visit column
                For lRowCount = 1 To .Rows - 1
                    If .TextMatrix(lRowCount, mnVISIT_COL) = .TextMatrix(lEnteredRow, mnVISIT_COL) And _
                       .TextMatrix(lRowCount, mnTRIALSITEPERSON_COL) = .TextMatrix(lEnteredRow, mnTRIALSITEPERSON_COL) Then
                        'Select
                        .Row = lRowCount
                        .Col = mnVISIT_COL
                        .ColSel = .Cols - 1
                        .CellBackColor = vbHighlight
                        .CellForeColor = vbHighlightText
                        If mlRowClickedMin > .Row Then
                            mlRowClickedMin = .Row
                        End If
                        If mlRowClickedMax < .Row Then
                            mlRowClickedMax = .Row
                        End If
                        'picStatusIcon.BackColor = vbHighlight
                        Call SetIcon(mnVISIT_COL, .Row, vbHighlight)
                        Call SetIcon(mnFORM_COL, .Row, vbHighlight)
                        Call SetIcon(mnSTATUS_COL, .Row, vbHighlight)
                    End If
                Next lRowCount
            Case mnFORM_COL
                ' CRF Page (eForm) column
                For lRowCount = 1 To .Rows - 1
                    If .TextMatrix(lRowCount, mnFORM_COL) = .TextMatrix(lEnteredRow, mnFORM_COL) And _
                       .TextMatrix(lRowCount, mnVISIT_COL) = .TextMatrix(lEnteredRow, mnVISIT_COL) And _
                       .TextMatrix(lRowCount, mnTRIALSITEPERSON_COL) = .TextMatrix(lEnteredRow, mnTRIALSITEPERSON_COL) Then
                        'Select
                        .Row = lRowCount
                        .Col = mnFORM_COL
                        .ColSel = .Cols - 1
                        .CellBackColor = vbHighlight
                        .CellForeColor = vbHighlightText
                        If mlRowClickedMin > .Row Then
                            mlRowClickedMin = .Row
                        End If
                        If mlRowClickedMax < .Row Then
                            mlRowClickedMax = .Row
                        End If
                        
                        'picStatusIcon.BackColor = vbHighlight
                        Call SetIcon(mnFORM_COL, .Row, vbHighlight)
                        Call SetIcon(mnSTATUS_COL, .Row, vbHighlight)
                    End If
                Next lRowCount
            Case Else
                ' Question column
                'Select
                'TA 28/04/2000 SR3276
                If lEnteredRow > 0 Then
                    'only if not header row
                    .Row = lEnteredRow
                    .RowSel = lEnteredRow
                    .Col = mnDATAITEM_COL
                    .ColSel = .Cols - 1
                    .CellBackColor = vbHighlight
                    .CellForeColor = vbHighlightText
                    
                    'picStatusIcon.BackColor = vbHighlight
                    Call SetIcon(mnSTATUS_COL, .Row, vbHighlight)
                End If
        End Select
        
        .FillStyle = flexFillSingle
        
        .Refresh
        .Visible = True
        .Col = lEnteredCol
        .ColSel = lEnteredCol
        .Row = lEnteredRow
        .RowSel = lEnteredRow
    End With
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "flexData_Click", Err.Source) = Retry Then
        Resume
    End If
   
End Sub



'---------------------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------------------
'decide on which format of printing is required
'---------------------------------------------------------------------
    
    'TA 13/08/2001: changed to use mvData
       
        PrintData
    
End Sub



'---------------------------------------------------------------------
Private Sub flexData_DblClick()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    ' NCJ 9/3/00 - We sometimes get a DblClick without a click
    ' so try and remedy this...
    ' Is the current row the same as the row last clicked?
    If mlRowClicked <> flexData.Row Then
        Call flexData_Click
    End If
    
    ' NCJ 29/3/00  SR3213 - Ignore double clicks in Subject or Visit columns
    'MLM 21/01/03: Also ignore double clicks in the header row
    If flexData.Col > mnVISIT_COL And flexData.Row > 0 Then
        If mDRType = dbtDataItemResponse Then
            'MLM 27/06/03: 3.0 bug list 1709: While waiting for the selected form to appear, don't let the user do anything in the data browser.
            Me.Enabled = False
            Call ShowForm
            Me.Enabled = True
        End If
    End If
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "flexData_DblClick", Err.Source) = Retry Then
        Resume
    End If
   
End Sub

'---------------------------------------------------------------------
Private Sub flexData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
'Changed by Mo Morris 11/1/99 SR 628
'When displaying forms as opposed to data items 4 of flexData's columns become hidden by
'having their width set to 0. When adjusting flexData's column widths by dragging the
'sides of the column headers it is possible to widen one of these 0 width columns.
'Unfortunately there does not appear to be a place where this can be prevented from happening,
'so a check has been placed here on the event most likely to occur afterwards
'---------------------------------------------------------------------
    
    If mDRType = dbteForms And flexData.ColWidth(mnDATAITEM_COL) <> 0 Then
        flexData.ColWidth(mnDATAITEM_COL) = 0
    End If

End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
' NB Note that Form_Load can get called AFTER Init or Monitor
' REVISIONS
' ic 14/02/2003 bug 807, removed '/Label' from column header
' DPH 23/09/2003 - initialise collection of drawn images
' ic 28/07/2005 added clinical coding
'---------------------------------------------------------------------
Dim oCon As Control

    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon
    
    ' DPH 23/09/2003 - initialise collection of drawn images
    Set mcolDrawnImages = New Collection
    
    imgStatusIcons.MaskColor = vbTransparent
    
    ' lblCommentSize is only used for getting the size of comments and RFC
    lblCommentSize.Visible = False
    Me.BackColor = eMACROColour.emcBackground
    lblTotals.BackColor = eMACROColour.emcBackground
    
'TA 26/09/2002: Part of New UI  - ebsure all backgrounds are white
    For Each oCon In Me.Controls
        On Error Resume Next
        oCon.BackColor = eMACROColour.emcBackground
    Next
    
    On Error GoTo ErrHandler
    
    
    With lcmdClose
        .BackColor = eMACROColour.emcBackground
        .Font = MACRO_DE_FONT
        .FontSize = 8
        .ForeColor = eMACROColour.emdEnabledText
        .FontUnderline = True
        .MouseIcon = frmImages.CursorHandPoint.MouseIcon
        .MousePointer = 99 'custom
    End With
    
    
    With lcmdPrint
        .BackColor = eMACROColour.emcBackground
        .Font = MACRO_DE_FONT
        .FontSize = 8
        .ForeColor = eMACROColour.emdEnabledText
        .FontUnderline = True
        .MouseIcon = frmImages.CursorHandPoint.MouseIcon
        .MousePointer = 99 'custom
    End With
    
    
    With flexData
'        .Cols = 12
        ' NCJ 19/5/00 - mnINDEX_COL is last column (and they start at 0)
        If gbClinicalCoding Then
            .Cols = (mnNUM_OF_COLS + mnNUM_OF_CCCOLS)
        Else
            .Cols = mnNUM_OF_COLS
        End If
        .MergeCol(mnTRIALSITEPERSON_COL) = True
        .MergeCol(mnVISIT_COL) = True
        .MergeCol(mnFORM_COL) = True
        .FixedCols = 0
        .GridLines = flexGridFlat
        
        'ic 14/02/2003 bug 807, removed '/Label'
        .Row = 0
        .Col = mnTRIALSITEPERSON_COL
        .Text = "Study/Site/Subject"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)
        '.ColWidth(mnTRIALSITEPERSON_COL) = 1600
        .ColAlignment(mnTRIALSITEPERSON_COL) = flexAlignLeftCenter
        
        .Col = mnVISIT_COL
        .Text = "Visit"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)
'        .ColWidth(mnVISIT_COL) = 1200
        .ColAlignment(mnVISIT_COL) = flexAlignLeftCenter
        
        .Col = mnFORM_COL
        .Text = "eForm"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)
'        .ColWidth(mnFORM_COL) = 1000
        .ColAlignment(mnFORM_COL) = flexAlignLeftCenter
        
        .Col = mnDATAITEM_COL
        .Text = "Question"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)
'        .ColWidth(mnDATAITEM_COL) = 1500
        .ColAlignment(mnDATAITEM_COL) = flexAlignLeftCenter
        
        .Col = mnDATARESPONSE_COL
        .Text = "Value"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)
'        .ColWidth(mnDATARESPONSE_COL) = 1000
        .ColAlignment(mnDATARESPONSE_COL) = flexAlignLeftCenter
        
        .Col = mnSTATUS_COL
        .Text = "Status"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)
'        .ColWidth(mnSTATUS_COL) = 1300
        .ColAlignment(mnSTATUS_COL) = flexAlignLeftCenter
        
        'SDM SR1185 30/11/99
        ' NCJ 29/3/00 - Renamed from "New" to "Transfer"
        .Col = mnNEW_COL
        .Text = "Transfer"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)
        .ColWidth(mnSTATUS_COL) = 1300
'        .ColAlignment(mnSTATUS_COL) = flexAlignLeftCenter
        
        .Col = mnTIMESTAMP_COL
        .Text = "Date and time"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)
'        .ColWidth(mnTIMESTAMP_COL) = 1750
        .ColAlignment(mnTIMESTAMP_COL) = flexAlignLeftCenter
        
        .Col = mnDB_TIMESTAMP_COL
        .Text = "Database date and time"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)

        .ColAlignment(mnDB_TIMESTAMP_COL) = flexAlignLeftCenter
        
        .Col = mnUSERID_COL
        .Text = "User Name"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)
'        .ColWidth(mnUSERID_COL) = 1000
        .ColAlignment(mnUSERID_COL) = flexAlignLeftCenter
               
        .Col = mnUSERNAMEFULL_COL
        .Text = "Full User Name"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)
'        .ColWidth(mnUSERID_COL) = 1000
        .ColAlignment(mnUSERNAMEFULL_COL) = flexAlignLeftCenter
        
        
        .Col = mnCOMMENT_COL
        .Text = "Comment"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)
'        .ColWidth(mnCOMMENT_COL) = 6000
        .ColAlignment(mnCOMMENT_COL) = flexAlignLeftCenter
        
        'SDM SR2408
        .Col = mnRFC_COL
        .Text = "Reason For Change"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)
        .ColAlignment(mnRFC_COL) = flexAlignLeftCenter
        
        ' NCJ 19/5/00 SR3454
        .Col = mnOVERRULE_COL
        .Text = "Overrule Reason"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)
        .ColAlignment(mnOVERRULE_COL) = flexAlignLeftCenter
        
        ' NCJ 17/10/00 SR3864
        .Col = mnVAL_MESSAGE_COL
        .Text = "Validation Message"
        .ColWidth(.Col) = 100
        AdjustColumnWidth (.Col)
        .ColAlignment(mnVAL_MESSAGE_COL) = flexAlignLeftCenter
        
        'ic 28/07/2005 added clinical coding columns
        If gbClinicalCoding Then
            .Col = mnDICTIONARY_COL
            .Text = "Dictionary"
            .ColWidth(.Col) = 100
            AdjustColumnWidth (.Col)
            .ColAlignment(mnDICTIONARY_COL) = flexAlignLeftCenter
            
            .Col = mnCODINGSTATUS_COL
            .Text = "Coding status"
            .ColWidth(.Col) = 100
            AdjustColumnWidth (.Col)
            .ColAlignment(mnCODINGSTATUS_COL) = flexAlignLeftCenter
            
            .Col = mnCODINGDETAILS_COL
            .Text = "Code"
            .ColWidth(.Col) = 100
            AdjustColumnWidth (.Col)
            .ColAlignment(mnCODINGDETAILS_COL) = flexAlignLeftCenter
        End If

        .Visible = False
    End With
    
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
Public Sub DisplayNew(vData As Variant, enDRType As eDataBrowserType)
'---------------------------------------------------------------------
'Show From Set up grid once the vData is known
'This will eventually replace Display
' NB vData may be Null
'---------------------------------------------------------------------
    
    Load Me
    
    Me.WindowState = vbNormal
    
    frmMenu.ResizeBottomRightHandForms Me.Name
        
    mnLaunchMode = eMACROWindow.MonitorBrowser

    Call SetUpGrid(vData, enDRType)

    Me.Show vbModeless

    Me.ZOrder

    
End Sub

'---------------------------------------------------------------------
Private Sub SetUpGrid(vData As Variant, enDRType As eDataBrowserType)
'---------------------------------------------------------------------
' Set up grid once the vData is known
' NB vData may be Null
' ic 28/07/2005 added clinical coding
'---------------------------------------------------------------------
Dim lRecordCount As Long
Dim i As Long

    HourglassOn
    
    mvData = vData
    mDRType = enDRType
    
    flexData.Visible = False
    flexData.Rows = 1
      
    If Not IsNull(vData) Then
        '2nd dimension is number of rows
        lRecordCount = UBound(mvData, 2) + 1
    
        ReDim mvRowShown(lRecordCount - 1) As Long
    Else
        lRecordCount = 0
    End If
    
    'add 1 because for the header row
    flexData.Rows = lRecordCount + 1
    
    ' NCJ 29/2/00 - New classes containing collection of Coords objects
    ' (These store which cells are locked or frozen)
    Set mcolLocked = New clsColCoords
    Set mcolFrozen = New clsColCoords
    ' Set the fixed integer (for calculating item keys)
    mcolLocked.FixedInteger = flexData.Cols
    mcolFrozen.FixedInteger = flexData.Cols

    flexData.Row = 0
    
    If gbClinicalCoding Then
        For i = 0 To (mnNUM_OF_COLS + mnNUM_OF_CCCOLS) - 1
            flexData.ColWidth(i) = 0
            'adjust to header width
            AdjustColumnWidth (i)
            
        Next
    Else
        For i = 0 To mnNUM_OF_COLS - 1
            flexData.ColWidth(i) = 0
            'adjust to header width
            AdjustColumnWidth (i)
            
        Next
    End If
    
    'hide/show cols according to forms or data view
    If mDRType = eDataBrowserType.dbteForms Then
       flexData.TextMatrix(0, mnTIMESTAMP_COL) = "Date"
        flexData.ColWidth(mnDATAITEM_COL) = 0
        flexData.ColWidth(mnDATARESPONSE_COL) = 0
        flexData.ColWidth(mnSTATUS_COL) = 0
        flexData.ColWidth(mnNEW_COL) = 0
        flexData.ColWidth(mnUSERID_COL) = 0
        flexData.ColWidth(mnCOMMENT_COL) = 0
        flexData.ColWidth(mnRFC_COL) = 0    'SDM SR2408
        flexData.ColWidth(mnOVERRULE_COL) = 0    'NCJ SR3454
        flexData.ColWidth(mnVAL_MESSAGE_COL) = 0    'NCJ 17/10/00 SR3864
        
        If gbClinicalCoding Then
            flexData.ColWidth(mnDICTIONARY_COL) = 0
            flexData.ColWidth(mnCODINGSTATUS_COL) = 0
            flexData.ColWidth(mnCODINGDETAILS_COL) = 0
        End If
        
    Else
        'set row to 0 so header is used for resizing column
        flexData.Row = 0
        flexData.TextMatrix(0, mnTIMESTAMP_COL) = "Date and time"
        flexData.TextMatrix(0, mnDB_TIMESTAMP_COL) = "Database date and time"
        AdjustColumnWidth (mnDATAITEM_COL)
        AdjustColumnWidth (mnDATARESPONSE_COL)
        AdjustColumnWidth (mnSTATUS_COL)
        AdjustColumnWidth (mnNEW_COL)
        AdjustColumnWidth (mnUSERID_COL)
        AdjustColumnWidth (mnCOMMENT_COL)
        AdjustColumnWidth (mnRFC_COL)
        AdjustColumnWidth (mnOVERRULE_COL)
        AdjustColumnWidth (mnVAL_MESSAGE_COL)
        
        'ic 28/07/2005 added clinical coding columns
        If gbClinicalCoding Then
            If (enDRType = eDataBrowserType.dbtDataItemResponse) Then
                AdjustColumnWidth (mnDICTIONARY_COL)
                AdjustColumnWidth (mnCODINGSTATUS_COL)
                AdjustColumnWidth (mnCODINGDETAILS_COL)
            Else
                flexData.ColWidth(mnDICTIONARY_COL) = 0
                flexData.ColWidth(mnCODINGSTATUS_COL) = 0
                flexData.ColWidth(mnCODINGDETAILS_COL) = 0
            End If
        End If
    End If

    'TA 28/02/2001: new code to fill grid
    Call PopulateGridSection(1)

    'focus on first row
    If flexData.Rows > 2 Then
        flexData.Row = 1
        flexData.Col = mnDATAITEM_COL
        flexData.TopRow = 1
    End If

    lblTotals.Caption = lRecordCount & " records"

    flexData.Visible = True

    HourglassOff
    
End Sub


'---------------------------------------------------------------------
Private Sub flexdata_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
'   If right mouse click and the browser is in data mode (rather than form or audit trail)
'   and the mouse position is on a valid data item response row
'   then display the popup list to allow changes to the status or a comment to be added
'   TA 01/02/2001 SR3391: store leftmost column so that grid doesn't move about
'   ic 28/07/2005 added clinical coding
'---------------------------------------------------------------------
Dim lMouseCol As Long
Dim lCurRow As Long
Dim lLeftCol As Long

    On Error GoTo ErrHandler
    
    lLeftCol = flexData.LeftCol
    
    'Debug.Print "Mousedown on Col = " & flexData.Col
    'Debug.Print "Mousedown on Row = " & flexData.Row
    'Debug.Print "ColSel = " & flexData.ColSel
    'Debug.Print "RowSel = " & flexData.RowSel
 
    
    If flexData.MouseRow >= flexData.Rows Then Exit Sub
    
    'SDM 01/12/99 SR1355
    'Allows the right mouse button to select the row
    If Button = vbRightButton Then
        flexData.Row = flexData.MouseRow
        Call flexData_Click
    End If

    ' NCJ 1/3/00 - Store these values locally
    'MLM 27/06/03: Display pop-up menu based on where the mouse was when flexData_Click was called, rather than where it is now
'    lMouseCol = flexData.MouseCol
'    lCurRow = flexData.Row
    lMouseCol = mlColumnClicked
    lCurRow = mlRowClicked
    
    ' If mouse is outside current cell then do nothing
    If Y < flexData.CellTop Or Y > flexData.CellTop + flexData.CellHeight Then
        Exit Sub
    End If
    
    'REM 01/09/03 - If user right clicks on a column header then do nothing
    If lCurRow = 0 Then
        Exit Sub
    End If
    
    ' NCJ 17 Oct 02 - Let disabled right-mouse menu be shown for forzen items
    ' If item is frozen and it's not a question, do nothing
'    If mcolFrozen.IsItem(lCurRow, lMouseCol) And lMouseCol < mnDATAITEM_COL Then
'        Exit Sub
'    End If
    
    ' Prepare and show right mouse menu to change statuses etc.
    'ic 28/07/2005 handle clinical coding columns

    If Button = vbRightButton And (mDRType = eDataBrowserType.dbtDataItemResponse) And lMouseCol <= mnCODINGSTATUS_COL Then
        ' NCJ 17 Oct 02 - New right mouse menu to handle popup and everything
        Call RightMouseMenu(lCurRow, lMouseCol)
    End If
    
    ' RS 23/01/2003: Do not reset, as required by Click event to redraw status icons of selected area
    'mlRowClickedMin = 0
    'mlRowClickedMax = 0
    
    'ensure same leftmost col is leftmost
    flexData.LeftCol = lLeftCol
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "RightMouseMenu", Err.Source) = Retry Then
        Resume
    End If
   
End Sub

'---------------------------------------------------------------------
Private Function RightMouseMenu(lRow As Long, lCol As Long) As Boolean
'---------------------------------------------------------------------
' NCJ 16 Oct 02 - Rewritten based on original EnableStatusMenu
' Do the Right Mouse Menu for the Data Browser,
' depending on the Row and Column clicked
' Assume that frozen Subject, Visit and Form clicks don't come here
' ic 12/09/2005 added clinical coding
'---------------------------------------------------------------------
Dim oMenuItems As clsMenuItems
Dim bLocked As Boolean
Dim bUnLocked As Boolean
Dim bFrozen As Boolean
Dim sItem As String
Dim bCanUnfreeze As Boolean
Dim nCodingStatus As Integer


    On Error GoTo ErrHandler
    
    ' See if we have a locked or frozen item
    bFrozen = mcolFrozen.IsItem(lRow, lCol)
    bLocked = mcolLocked.IsItem(lRow, lCol)
    bUnLocked = Not bLocked And Not bFrozen
    
    If gbClinicalCoding Then
        If (RemoveNull(mvData(dbcCodingStatus, lRow - 1)) = "") Then
            nCodingStatus = eCodingStatus.csEmpty
        Else
            nCodingStatus = Val(RemoveNull(mvData(dbcCodingStatus, lRow - 1)))
        End If
    End If
    
    ' NCJ 23 Dec 02 - Can only unfreeze if it's the subject OR our parent is not frozen
    ' (User's Unfreeze permission is checked later)
    If bFrozen Then
        If lCol = mnTRIALSITEPERSON_COL Then
            ' We can unfreeze the subject
            bCanUnfreeze = True
        Else
            ' Our parent (i.e. the column to the left) must not be frozen
            bCanUnfreeze = Not (mcolFrozen.IsItem(lRow, lCol - 1))
        End If
    Else
        ' Can't unfreeze if not frozen!
        bCanUnfreeze = False
    End If
    
    Set oMenuItems = New clsMenuItems
    
    ' Create the menu items
    Call oMenuItems.Add("LOCK", "Lock", _
                goUser.CheckPermission(gsFnLockData) And bUnLocked)
    ' NCJ 24 Aug 04 - Changed gsFnLockData to gsFnUnLockData
    Call oMenuItems.Add("UNLOCK", "Unlock", _
                goUser.CheckPermission(gsFnUnLockData) And bLocked)
    Call oMenuItems.Add("FREEZE", "Freeze", _
                goUser.CheckPermission(gsFnFreezeData) And Not bFrozen)
    Call oMenuItems.Add("UNFREEZE", "Unfreeze", _
                goUser.CheckPermission(gsFnUnFreezeData) And bCanUnfreeze)
    Call oMenuItems.AddSeparator
    Call oMenuItems.Add("SDV", "New SDV Mark...", _
                goUser.CheckPermission(gsFnCreateSDV) And Not bFrozen)
    
    If lCol >= mnDATAITEM_COL And lCol <= mnOVERRULE_COL Then
        ' It's a question - add Notes and Discrepancies
        lCol = mnDATAITEM_COL
        Call oMenuItems.Add("NOTE", "New Note...", Not bFrozen)
        Call oMenuItems.Add("DISC", "New Discrepancy...", _
                goUser.CheckPermission(gsFnCreateDiscrepancy) And Not bFrozen)
    End If
    
    If gbClinicalCoding Then
        If (nCodingStatus <> eCodingStatus.csEmpty) And (lCol >= mnDATAITEM_COL) Then
            Call oMenuItems.AddSeparator
            
            'set to validated or return to coded
            If nCodingStatus = eCodingStatus.csValidated Then
                Call oMenuItems.Add("CODESTATUSCODE", "Return coding status to Coded")
                oMenuItems.KeyedItem("CODESTATUSCODE").Enabled = goUser.CheckPermission(gsFnValidateClinicalCode)
            Else
                Call oMenuItems.Add("CODESTATUSVALID", "Set coding status to Validated")
                oMenuItems.KeyedItem("CODESTATUSVALID").Enabled = goUser.CheckPermission(gsFnValidateClinicalCode) _
                    And (nCodingStatus = eCodingStatus.csCoded _
                    Or nCodingStatus = eCodingStatus.csAutoEncoded)
            End If
                
            'set to pending or return to coded
            If nCodingStatus = eCodingStatus.csPendingNewCode Then
                Call oMenuItems.Add("CODESTATUSCODE", "Return coding status to Coded")
                oMenuItems.KeyedItem("CODESTATUSCODE").Enabled = goUser.CheckPermission(gsFnChangeClinicalStatus)
            Else
                Call oMenuItems.Add("CODESTATUSPEND", "Set coding status to Pending New Code")
                oMenuItems.KeyedItem("CODESTATUSPEND").Enabled = goUser.CheckPermission(gsFnChangeClinicalStatus) _
                    And (nCodingStatus = eCodingStatus.csCoded _
                    Or nCodingStatus = eCodingStatus.csAutoEncoded _
                    Or nCodingStatus = eCodingStatus.csValidated)
            End If
            
            'set to not coded
            Call oMenuItems.Add("CODESTATUSNOT", "Set coding status to Not Coded")
            oMenuItems.KeyedItem("CODESTATUSNOT").Enabled = goUser.CheckPermission(gsFnChangeClinicalStatus) _
                And (nCodingStatus <> eCodingStatus.csNotCoded And nCodingStatus <> eCodingStatus.csEmpty)
                
            'set to do not code
            Call oMenuItems.Add("CODESTATUSDONT", "Set coding status to Do Not Code")
            oMenuItems.KeyedItem("CODESTATUSDONT").Enabled = goUser.CheckPermission(gsFnChangeClinicalStatus) _
                And (nCodingStatus = eCodingStatus.csNotCoded Or nCodingStatus = eCodingStatus.csEmpty)
        End If
    End If

    ' Show the popup menu
    sItem = frmMenu.ShowPopUpMenu(oMenuItems)
   
    RightMouseMenu = True
    
    Select Case sItem
        ' NCJ 23 Dec 02 - Pass new LFAction rather than LockStatus
        Case "LOCK"
            Call ChangeLockUnlockFreeze(LFAction.lfaLock)
        Case "UNLOCK"
            Call ChangeLockUnlockFreeze(LFAction.lfaUnlock)
        Case "FREEZE"
            Call ChangeLockUnlockFreeze(LFAction.lfaFreeze)
        Case "UNFREEZE"
            Call ChangeLockUnlockFreeze(LFAction.lfaUnfreeze)
            
        Case "SDV"
            Call NewMIMessage(lCol, MIMsgType.mimtSDVMark)
        Case "NOTE"
            Call NewMIMessage(lCol, MIMsgType.mimtNote)
        Case "DISC"
            Call NewMIMessage(lCol, MIMsgType.mimtDiscrepancy)
        
        'ic 27/10/2005 added clinical coding
        Case "CODESTATUSVALID"
            Call ChangeCodingStatus(lRow, csValidated)
        Case "CODESTATUSCODE"
            Call ChangeCodingStatus(lRow, csCoded)
        Case "CODESTATUSPEND"
            Call ChangeCodingStatus(lRow, csPendingNewCode)
        Case "CODESTATUSNOT"
            Call ChangeCodingStatus(lRow, csNotCoded)
        Case "CODESTATUSDONT"
            Call ChangeCodingStatus(lRow, csDoNotCode)
            
        Case Else
            ' They didn't select anything
            RightMouseMenu = False
    End Select

    Set oMenuItems = Nothing
    
Exit Function
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "RightMouseMenu", Err.Source) = Retry Then
        Resume
    End If

End Function

'---------------------------------------------------------------------
Private Sub ChangeCodingStatus(lRow As Long, eNewStatus As eCodingStatus)
'---------------------------------------------------------------------
' ic 12/09/2005
' change the coding status of a response
'---------------------------------------------------------------------
Dim bOK As Boolean
Dim bRefresh As Boolean
Dim sToken As String
Dim oCodedTermHistory As MACROCCBS30.CodedTermHistory


    bRefresh = False
    Select Case eNewStatus
    Case eCodingStatus.csValidated, eCodingStatus.csDoNotCode, eCodingStatus.csCoded
        'status change to validated/do not code/coded - just change the status
        bRefresh = True
        
    Case eCodingStatus.csPendingNewCode
        'status change to pending new - warn the user that the existing code will be lost
        If DialogQuestion("The current coding details for this question will be overwritten - are you sure you wish to continue?") = vbYes Then
            bRefresh = True
        End If
        
    Case eCodingStatus.csNotCoded
        'status change to not coded - warn the user that the existing code, if any, will be lost
        bOK = False
        
        If Val(RemoveNull(mvData(dbcCodingStatus, lRow - 1))) = eCodingStatus.csDoNotCode Then
            bOK = True
        Else
            If DialogQuestion("The current coding details for this question will be cleared - are you sure you wish to continue?") = vbYes Then
                bOK = True
            End If
        End If
        If bOK Then
            'if user accepted, clear coding and dictionary details
            bRefresh = True
        End If
    End Select

    If bRefresh Then
        sToken = LockSubject(goUser.UserName, CLng(mvData(dbcStudyId, lRow - 1)), CStr(mvData(dbcSite, lRow - 1)), _
            CLng(mvData(dbcSubjectId, lRow - 1)))
        If sToken = "" Then
            ' this subject is currently locked (Message already given to User in LockSubject)
            Exit Sub
        End If
        
        'load the coded term
        Set oCodedTermHistory = New MACROCCBS30.CodedTermHistory
        Call oCodedTermHistory.InitAuto(goUser.CurrentDBConString, CLng(mvData(dbcStudyId, lRow - 1)), CStr(mvData(dbcSite, lRow - 1)), _
            CLng(mvData(dbcSubjectId, lRow - 1)), CLng(mvData(dbcResponseTaskId, lRow - 1)), _
            CInt(mvData(dbcResponseCycleNumber, lRow - 1)))
        'set the new status
        Call oCodedTermHistory.SetStatus(CInt(eNewStatus), goUser.UserName, goUser.UserNameFull, _
            ConvertFromNull(mvData(dbcResponseValue, lRow - 1), vbString), CDbl(mvData(dbcResponseTimeStamp, lRow - 1)), _
            CInt(mvData(dbcResponseTimestamp_TZ, lRow - 1)))
        'save the changed value
        Call oCodedTermHistory.Save(goUser.CurrentDBConString, CLng(mvData(dbcVisitId, lRow - 1)), _
            CInt(mvData(dbcVisitCycleNumber, lRow - 1)), CLng(mvData(dbcEFormId, lRow - 1)), _
            CInt(mvData(dbcEFormCycleNumber, lRow - 1)))
        Set oCodedTermHistory = Nothing
  
        'unlock the subject
        UnlockSubject CLng(mvData(dbcStudyId, lRow - 1)), CStr(mvData(dbcSite, lRow - 1)), CLng(mvData(dbcSubjectId, lRow - 1)), sToken
        'refresh the results pane
        Call frmMenu.RefreshSearchResults
    End If
End Sub


'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------

    On Error Resume Next

    lblTotals.Left = 60 'picSelection.Left + picSelection.Width
    lblTotals.Width = Me.ScaleWidth - lblTotals.Left - 120
    lblTotals.Top = Me.ScaleHeight - lblTotals.Height '- 350
    
    lcmdClose.Left = Me.ScaleWidth - lcmdClose.Width - 120
    lcmdClose.Top = 0
    lcmdPrint.Top = 0
    lcmdPrint.Left = lcmdClose.Left - lcmdPrint.Width
       
    flexData.Left = 60 'picSelection.Left + picSelection.Width
    flexData.Width = Me.ScaleWidth - flexData.Left - 120
    flexData.Top = lcmdClose.Height
    flexData.Height = lblTotals.Top - flexData.Top

End Sub

'---------------------------------------------------------------------
Public Sub AutoClose()
'---------------------------------------------------------------------
' This is called when MACRO is doing an automatic shut down
' e.g. on a Log Out or a time out
'---------------------------------------------------------------------

    ' NCJ 29/3/00 SR3300 - Reset Launch Mode ready for restart
    mnLaunchMode = eMACROWindow.None
    
End Sub

'---------------------------------------------------------------------
Private Sub NewMIMessage(lCol As Long, enType As MIMsgType)
'---------------------------------------------------------------------
' NCJ 17 Oct 02 - Create new MIMsg of given type
' (based on original QNewMIMessage)
' NCJ 5 Nov 02 - Check for already existing SDV marks
'ASH 16/01/2003 added RemoveNull(mvData(dbcResponseValue, lRow)) to avoid MACRO crashing when
'discrepancies are raised on missing missing items in databrowser
'---------------------------------------------------------------------
Dim enScope As MIMsgScope
Dim lRow As Long
Dim sIndexValues As String
Dim lStudyId As Long
Dim lPersonId As Long
Dim sSite As String
Dim lResponseTaskId As Long
Dim lVisitId As Long
Dim lCRFPageTaskId As Long
Dim nVisitCycleNumber As Integer
Dim nResponseCycle As Integer
Dim dblResponseTimeStamp As Double
Dim sResponseValue  As String
Dim oTimezone As TimeZone               ' RS 30/09/2002
Dim bCreateMessage As Boolean
Dim lEFormId As Long
Dim nEFormCycle As Integer
Dim lQuestionId As Long
Dim sUserName As String

    On Error GoTo ErrLabel
    
    bCreateMessage = True
    
    ' Determine the scope of the MIMsg
    Select Case lCol
    Case mnTRIALSITEPERSON_COL  ' Subject
        enScope = MIMsgScope.mimscSubject
    Case mnVISIT_COL        ' Visit
        enScope = MIMsgScope.mimscVisit
    Case mnFORM_COL     ' eForm
        enScope = MIMsgScope.mimscEForm
    Case mnDATAITEM_COL     ' Question
        enScope = MIMsgScope.mimscQuestion
    End Select
    
    ' Get the row
    lRow = mlRowClickedMin - 1
    
        'nb we must zero values not needed
    
    lStudyId = mvData(dbcStudyId, lRow)
    lPersonId = mvData(dbcSubjectId, lRow)
    sSite = mvData(dbcSite, lRow)
    lVisitId = mvData(dbcVisitId, lRow)
    nVisitCycleNumber = mvData(dbcVisitCycleNumber, lRow)
    lCRFPageTaskId = mvData(dbcEFormTaskID, lRow)
    lEFormId = mvData(dbcEFormId, lRow)
    nEFormCycle = mvData(dbcEFormCycleNumber, lRow)
    
    ' Get the last two values if needed and then zero out the ones we don't need
    If enScope = MIMsgScope.mimscQuestion Then
        lResponseTaskId = mvData(dbcResponseTaskId, lRow)
        dblResponseTimeStamp = mvData(dbcResponseTimeStamp, lRow)
        nResponseCycle = mvData(dbcResponseCycleNumber, lRow)
        sResponseValue = RemoveNull(mvData(dbcResponseValue, lRow))
        lQuestionId = mvData(dbcQuestionId, lRow)
        sUserName = mvData(dbcUserName, lRow)
    Else
        ' It's either eForm, Visit or Subject
        sUserName = "" 'no datausername for forms/visits/subjects
        dblResponseTimeStamp = 0
        nResponseCycle = 0
        lResponseTaskId = 0
        sResponseValue = ""
        If enScope <> mimscEForm Then
            ' It's either Visit or Subject
            lCRFPageTaskId = 0
            lEFormId = 0
            nEFormCycle = 0
            If enScope <> mimscVisit Then
                ' It's a Subject
                lVisitId = 0
                nVisitCycleNumber = 0
            End If
        End If
    End If
    
    ' For SDVs we only allow one per object
    If enType = MIMsgType.mimtSDVMark Then
        ' Check to see if there's one already
        If SDVExists(enScope, _
                        CStr(mvData(dbcStudyName, lRow)), sSite, lPersonId, _
                        lVisitId, nVisitCycleNumber, _
                        lCRFPageTaskId, lResponseTaskId) Then
            DialogInformation "An SDV Mark already exists for this " & GetScopeText(enScope)
            bCreateMessage = False
        End If
    End If
    
    If bCreateMessage Then
        Set oTimezone = New TimeZone
        ' Create the new object
        ' RS 23/01/2003: Update Icon only if a message was created
        If CreateNewMIMessage(enType, enScope, _
                        IMedNow, oTimezone.TimezoneOffset, _
                        sSite, _
                        lStudyId, _
                        lPersonId, _
                        Nothing, _
                        lVisitId, _
                        nVisitCycleNumber, _
                        lCRFPageTaskId, _
                        lResponseTaskId, _
                        nResponseCycle, _
                        sResponseValue, _
                        dblResponseTimeStamp, _
                        lEFormId, nEFormCycle, lQuestionId, sUserName) Then
    
            ' A message was actually created: Update Icon
            ' First update the mv array according to scope (propagate to left)
            Call NewMIstatus(lStudyId, sSite, lPersonId, lVisitId, nVisitCycleNumber, lCRFPageTaskId, lResponseTaskId, nResponseCycle, enType)
            
            
            
            
            
            Select Case enType
                Case MIMsgType.mimtDiscrepancy
                Case MIMsgType.mimtNote
                Case MIMsgType.mimtSDVMark
            End Select
            
        End If
    
        Set oTimezone = Nothing
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataItemResponse.NewMIMessage"
    
End Sub

'---------------------------------------------------------------------
Private Sub ChangeLockUnlockFreeze(ByVal enAction As LFAction)
'---------------------------------------------------------------------
'   SDM 07/12/99
' CHange the lock status of an object
' TA 5/10/2201: Subject is now locked during the update
' ATO 20/08/2002 Added RepeatNumber
' NCJ 23 Dec 02 - Now takes LockFreeze Action rather than LockStatus
' NCJ 3 Jan 03 - Use QuestionId too
' NCJ 9 Jan03 - Do a Refresh for Unfreeze (because we don't know what the final statuses are going to be)
'---------------------------------------------------------------------
Dim lRow As Long
Dim sIndexValues As String
Dim lCol As Long       ' NCJ 8/5/00
Dim lStudyId As Long
Dim lSubjectId As Long
Dim sSite As String
Dim lResponseTaskId As Long
Dim lVisitId As Long
Dim lCRFPageTaskId As Long
Dim nVisitCycleNumber As Integer
Dim nRepeatNumber As Integer 'ATO 21/08/2002
Dim sTimestamp As String 'ATO 21/08/2002

Dim sMsg As String
Dim sToken As String
Dim lMVDataIndex As Long

' NCJ 23 Dec 02 - New Lock/Freeze objects
Dim oLFObj As LFObject
Dim oFlocker As LockFreeze
Dim nSource As Integer
Dim enStatus As LockStatus
Dim lCRFPageID As Long
Dim nCRFPageCycleNumber As Integer
Dim sStudyName As String
Dim lQuestionId As Long
Dim oTimezone As TimeZone

    On Error GoTo ErrLabel
    
    HourglassOn
    
    Set oTimezone = New TimeZone
    
    ' Index into mvData is one less than grid row no.
    lMVDataIndex = mlRowClickedMin - 1
       
    lStudyId = mvData(dbcStudyId, lMVDataIndex)
    lSubjectId = mvData(dbcSubjectId, lMVDataIndex)
    sSite = mvData(dbcSite, lMVDataIndex)
    lVisitId = mvData(dbcVisitId, lMVDataIndex)
    nVisitCycleNumber = mvData(dbcVisitCycleNumber, lMVDataIndex)
    lCRFPageTaskId = mvData(dbcEFormTaskID, lMVDataIndex)
    lResponseTaskId = mvData(dbcResponseTaskId, lMVDataIndex)
    nRepeatNumber = mvData(dbcResponseCycleNumber, lMVDataIndex)
    lQuestionId = mvData(dbcQuestionId, lMVDataIndex)
    
    ' NCJ 23 Dec 02 - Use PageID and CycleNo
    lCRFPageID = mvData(dbcEFormId, lMVDataIndex)
    nCRFPageCycleNumber = mvData(dbcEFormCycleNumber, lMVDataIndex)
    
    HourglassOff
    
    sStudyName = mvData(dbcStudyName, lMVDataIndex)
    If Not CanDoLFAOnServer(sStudyName, sSite, lSubjectId) Then
        DialogInformation "Lock/freeze operations may not be carried out on this subject because there is unimported site data"
        Exit Sub
    End If
    
    'TA 27/9/01: this code checks to see if the subject is locked
    'nb The subject could be locked by by the current user having the schedule open
    sToken = LockSubject(goUser.UserName, lStudyId, sSite, lSubjectId)
    If sToken = "" Then
        ' this subject is currently locked (Message already given to User in LockSubject)
        Exit Sub
    End If
    
    ' NCJ 23 Dec 02 - You CAN Unfreeze in MACRO 3.0!
'    ' SDM 07/12/99   Warn user of permanence of freezing.

    HourglassOn
    
    flexData.Visible = False
    Set oLFObj = New LFObject
    Select Case mlColumnClicked
    Case mnTRIALSITEPERSON_COL
        ' Apply to the Trial Subject
        Call oLFObj.Init(LFScope.lfscSubject, lStudyId, sSite, lSubjectId)
    Case mnVISIT_COL
        ' Apply to the Visit
        Call oLFObj.Init(LFScope.lfscVisit, lStudyId, sSite, lSubjectId, _
                        lVisitId, nVisitCycleNumber)
    Case mnFORM_COL
        ' Apply to the eForm
        Call oLFObj.Init(LFScope.lfscEForm, lStudyId, sSite, lSubjectId, _
                        lVisitId, nVisitCycleNumber, _
                        lCRFPageID, nCRFPageCycleNumber)
    Case Is > mnFORM_COL
         ' Apply to single data item
        Call oLFObj.Init(LFScope.lfscQuestion, lStudyId, sSite, lSubjectId, _
                        lVisitId, nVisitCycleNumber, _
                        lCRFPageID, nCRFPageCycleNumber, _
                        lResponseTaskId, nRepeatNumber, lQuestionId)
    End Select
    
    ' NCJ 23 Dec 02 - Get the LockFreeze object to do all the work
    Set oFlocker = New LockFreeze
    If gblnRemoteSite Then
        nSource = TypeOfInstallation.RemoteSite
    Else
        nSource = TypeOfInstallation.Server
    End If
    Call oFlocker.DoLockFreeze(MacroADODBConnection, oLFObj, enAction, nSource, _
                                goUser.UserName, goUser.UserNameFull)
    
    ' Update grid so that the user does not have to refresh
    ' Mimic what's just been done in the SQL,
    ' i.e. change status of each row included in the selection
    ' and ripple the new status to the right
    ' NCJ 9 Jan 03 - Must Refresh for an Unfreeze
    Select Case enAction
    Case LFAction.lfaLock
        enStatus = LockStatus.lsLocked
    Case LFAction.lfaUnlock
        enStatus = LockStatus.lsUnlocked
    Case LFAction.lfaFreeze
        enStatus = LockStatus.lsFrozen
    Case LFAction.lfaUnfreeze
        ' NCJ 9 Jan 03 - Must Refresh for an Unfreeze
        Call frmMenu.RefreshSearchResults
    End Select
    
    ' Don't do cell colouring for an Unfreeze
    If enAction <> LFAction.lfaUnfreeze Then
        ' Now colour the grid cells accordingly
        For lRow = mlRowClickedMin To mlRowClickedMax
            flexData.Row = lRow
            ' Do columns up to the Question column
            If mlColumnClicked < mnDATAITEM_COL Then
                For lCol = mlColumnClicked To mnDATAITEM_COL
                    flexData.Col = lCol
                    ' If unfreezing, must allow changes to Frozen cells
                    Call SetLockUnlockFreeze(enStatus, (enAction = LFAction.lfaUnfreeze))
                    
                    Select Case enStatus
                        Case LockStatus.lsLocked:
                            If (lCol = 0 And mvData(dbcSubjectLockStatus, lRow - 1) = LockStatus.lsFrozen) Or (lCol = 1 And mvData(dbcVisitLockStatus, lRow - 1) = LockStatus.lsFrozen) Or (lCol = 2 And mvData(dbcEFormLockStatus, lRow - 1) = LockStatus.lsFrozen) Then
                                SetIcon lCol, lRow, vbHighlight
                            Else
                                SetIcon lCol, lRow, mnLOCKED_COLOUR
                            End If
                                
                        Case LockStatus.lsUnlocked: SetIcon lCol, lRow, vbHighlight
                        Case LockStatus.lsFrozen:   SetIcon lCol, lRow, mnFROZEN_COLOUR
                    End Select
                    
                Next lCol
            Else
                ' Single call automatically deals with columns to right of question
                flexData.Col = mlColumnClicked
                Call SetLockUnlockFreeze(enStatus, (enAction = LFAction.lfaUnfreeze))
            End If
            flexData.Col = mnTIMESTAMP_COL
            
            ' flexData.Text = Format(Now, "yyyy/mm/dd hh:mm:ss")
            ' RS 27/01/2003, 08/10/2002 Add Timezone Information, or convert timestamp to local format
            ' Use lRow - 1 as lRow indicates the grid, corresponding array row one less
            If GetMACROSetting("timestampdisplay", "storedvalue") = "storedvalue" Then
                ' Display the stored value, add offset to GMT in brackets
                ' Original Format/Value
                'flexData.Text = Format(mvData(dbcResponseTimeStamp, lRow - 1), "yyyy/mm/dd hh:mm:ss") & _
                                    " (GMT" & IIf(mvData(dbcResponseTimestamp_TZ, lRow - 1) < 0, "+", "") & -mvData(dbcResponseTimestamp_TZ, lRow - 1) \ 60 & ":" & Format(Abs(mvData(dbcResponseTimestamp_TZ, lRow - 1)) Mod 60, "00") & ")"
                'TA 22/05/2003: use function
                flexData.Text = DisplayGMTTime(mvData(dbcResponseTimeStamp, lRow - 1), "yyyy/mm/dd hh:mm:ss", mvData(dbcResponseTimestamp_TZ, lRow - 1))
                                                            
            Else
                ' Convert the stored value to local time
                flexData.Text = Format(oTimezone.ConvertDateTimeToLocal(mvData(dbcResponseTimeStamp, lRow - 1), mvData(dbcResponseTimestamp_TZ, lRow - 1)), "yyyy/mm/dd hh:mm:ss")
            End If

            'TA db timestamp (copied from above)
            flexData.Col = mnDB_TIMESTAMP_COL
            ' RS 27/01/2003, 08/10/2002 Add Timezone Information, or convert timestamp to local format
            If GetMACROSetting("timestampdisplay", "storedvalue") = "storedvalue" Then
                ' Display the stored value, add offset to GMT in brackets
                ' Original Format/Value
'                flexData.Text = Format(mvData(dbcDatabaseTimeStamp, lRow - 1), "yyyy/mm/dd hh:mm:ss") & _
                                    " (GMT" & IIf(mvData(dbcDatabaseTimestamp_TZ, lRow - 1) < 0, "+", "") & -mvData(dbcDatabaseTimestamp_TZ, lRow - 1) \ 60 & ":" & Format(Abs(mvData(dbcDatabaseTimestamp_TZ, lRow - 1)) Mod 60, "00") & ")"
                 'TA 22/05/2003: use function
                flexData.Text = DisplayGMTTime(mvData(dbcDatabaseTimeStamp, lRow - 1), "yyyy/mm/dd hh:mm:ss", mvData(dbcDatabaseTimestamp_TZ, lRow - 1))
            
            Else
                ' Convert the stored value to local time
                flexData.Text = Format(oTimezone.ConvertDateTimeToLocal(mvData(dbcDatabaseTimeStamp, lRow - 1), mvData(dbcDatabaseTimestamp_TZ, lRow - 1)), "yyyy/mm/dd hh:mm:ss")
            End If
            
            
            flexData.Col = mnUSERID_COL
            flexData.Text = goUser.UserName
        Next lRow
    End If
    
    ' NCJ 25/4/00 - If we unlocked, then the SQL unlocks will have rippled leftwards
    ' so refresh all rows that match this ClinicalTrialId, Person and Site
    If enAction = LFAction.lfaUnlock Then
        ' We'll stop when we've found our subject
        ' First of all go from where we are forwards
        For lRow = mlRowClickedMin To flexData.Rows - 1
            If Not RefreshUnlocks(lRow, lStudyId, _
                                sSite, lSubjectId, lCRFPageTaskId, _
                                lVisitId, nVisitCycleNumber) Then Exit For
        Next
        
        ' Then go from where we are backwards
        If mlRowClickedMin > 1 Then
            For lRow = mlRowClickedMin - 1 To 1 Step -1
                If Not RefreshUnlocks(lRow, lStudyId, _
                                sSite, lSubjectId, lCRFPageTaskId, _
                                lVisitId, nVisitCycleNumber) Then Exit For
            Next
        End If
        
    End If
    
    ' unlock the subject (this is the database lock!)
    UnlockSubject lStudyId, sSite, lSubjectId, sToken
    
    HourglassOff
    
    Set oLFObj = Nothing
    Set oFlocker = Nothing
    
    flexData.Refresh
    flexData.Visible = True
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataItemResponse.ChangeLockUnlockFreeze"

End Sub

'---------------------------------------------------------------------
Private Function RefreshUnlocks(ByVal lRow As Long, _
                    ByVal lStudyId As Long, _
                    ByVal sSite As String, _
                    ByVal lSubjectId As Long, _
                    ByVal lCRFPageTaskId As Long, _
                    ByVal lVisitId As Long, _
                    ByVal nVisitCycleNumber As Integer) As Boolean
'---------------------------------------------------------------------
' NCJ 25 Apr 00 - Unlock the current cell if it's to the left of mlColumnClicked
' and it matches the given details
' Return FALSE is this cell doesn't match the Trial/Site/Subject
' or TRUE if it does
' NCJ 23 Dec 02 - mvData is indexed by Row-1
'---------------------------------------------------------------------
Dim lThisTrialId As Long
Dim lThisPersonId As Long
Dim sThisSite As String
Dim lThisResponseTaskId As Long
Dim lThisVisitId As Long
Dim lThisCRFPageTaskId As Long
Dim nThisVisitCycleNumber As Integer
Dim bFoundSubject As Boolean
Dim lMVDataIndex As Long

    On Error GoTo ErrLabel
    
    bFoundSubject = False
    
    ' NCJ 23 Dec 02
    lMVDataIndex = lRow - 1
    
    lThisTrialId = mvData(dbcStudyId, lMVDataIndex)
    If lThisTrialId = lStudyId Then
        ' cell matches required TrialID
        lThisPersonId = mvData(dbcSubjectId, lMVDataIndex)
        sThisSite = mvData(dbcSite, lMVDataIndex)
        If lThisPersonId = lSubjectId And sThisSite = sSite Then
            ' We've found one we want to update
            bFoundSubject = True
            lThisResponseTaskId = mvData(dbcResponseTaskId, lMVDataIndex)
            lThisVisitId = mvData(dbcVisitId, lMVDataIndex)
            lThisCRFPageTaskId = mvData(dbcEFormTaskID, lMVDataIndex)
            nThisVisitCycleNumber = mvData(dbcVisitCycleNumber, lMVDataIndex)
            ' Update lock settings for page, visit and subject columns
            ' to the left of the column clicked
            flexData.Row = lRow
            If mlColumnClicked > mnTRIALSITEPERSON_COL Then
                flexData.Col = mnTRIALSITEPERSON_COL
                Call SetCellLockUnlockFreeze(LockStatus.lsUnlocked)
            End If
            If mlColumnClicked > mnVISIT_COL _
             And lThisVisitId = lVisitId _
             And nThisVisitCycleNumber = nVisitCycleNumber Then
                flexData.Col = mnVISIT_COL
                Call SetCellLockUnlockFreeze(LockStatus.lsUnlocked)
            End If
            If mlColumnClicked > mnFORM_COL And lThisCRFPageTaskId = lCRFPageTaskId Then
                flexData.Col = mnFORM_COL
                Call SetCellLockUnlockFreeze(LockStatus.lsUnlocked)
            End If
        End If
    End If

    RefreshUnlocks = bFoundSubject
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataItemResponse.RefreshUnlocks"
   
End Function

'---------------------------------------------------------------------
Private Sub ShowForm()
'---------------------------------------------------------------------
'   ATN 26/4/99
'   Subroutine to launch a data entry form from the browser
'   NCJ 30/11/99 - Calculate form date after opening form
'   NCJ 12/1/00, SR 2644 - Ensure Visit schedule is created when showing eForm
' NCJ 24/5/00 - Code moved to new routine ShowEForm
'---------------------------------------------------------------------
Dim vValues As Variant
Dim lStudyId As Long
Dim sSite As String
Dim lSubjectId As Long
Dim lEFITaskId As Long
Dim lResponseTaskId As Long
Dim sIndexValues As String
Dim bFocusToQuestion As Boolean
Dim nRptNo As Integer
Dim lRow As Long

    On Error GoTo ErrLabel

    HourglassOn
    
    lRow = flexData.Row - 1
    
    lStudyId = mvData(dbcStudyId, lRow)
    lSubjectId = mvData(dbcSubjectId, lRow)
    sSite = mvData(dbcSite, lRow)
    lResponseTaskId = mvData(dbcResponseTaskId, lRow)
    
    lEFITaskId = mvData(dbcEFormTaskID, lRow)

    ' NCJ 11 Mar 02 - Get ResponseCycleNumber too
    nRptNo = mvData(dbcResponseCycleNumber, lRow)

    ' NCJ 29/3/00 SR3213 - If in form column, show first element on form
    bFocusToQuestion = (flexData.Col <> mnFORM_COL)
    
    ' NCJ 11 Mar 02 - Added nRptNo
    Call frmMenu.EFIOpen(lStudyId, sSite, lSubjectId, lEFITaskId, _
                        lResponseTaskId, nRptNo, "", bFocusToQuestion)
    
    HourglassOff
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataItemResponse.ShowForm"
   
End Sub

'---------------------------------------------------------------------
Private Sub PrintData()
'---------------------------------------------------------------------
'Changed by  Mo Morris   11/9/98     SPR 426 (change made in Released and Developed versions)
'CancelError on CommonDialog1 enabled and tested for after the ShowPrinter call.
' NCJ 30/11/99 - Expect timestamps to be doubles
'SDM 02/12/99   Changed to Private
'Mo Morris  23/3/00 SR 3261
'   'New' and 'Reason For change added to pintout.
'   'Reason for Change' is handled in the same manner as comments because
'   the text can go over onto another line.
'Mo Morris  1/12/00 In adapting the form to work with Letter paper as well as A4 paper
'   The position and width of 'Comment' and 'Reason for change' now varies
'   ResponseValue text now wraps
'   variable names changed
'Mo Morris  2/2/01 DataItemName text no wraps. changes made around keeping track of currentY
'   at the end of a line.
'Mo Morris  19/7/02, CBB 2.2.15.5
'   Subject Label, Overrule Reason and Warning Message (Validation Message) added.
'   Manner of printing Comments/Reason For Change/Overrule reason/Warning Message changed,
'   instead of these test field wrapping from line to line within a 1 inch column,
'   they will share a single 2.5 inch column and have identifying labels placed above them.
'   Column headings in this printout have been made the same as in the Data Browser.
'   Minor column width adjustments to the left of comment fields.
'   Study, Visit and form headings now included in the logic of when to throw a page.
'   Page width check added to prevent printing in Portrait when the setting to Landscape
'   has not been picked up by the common dialog control when running under Windows 2000
'
'   Column widths (in inches/twips) changed as follows:-
'   FIELD           WIDTH       START AT
'   Question        1.5         .5/720
'   Response        1.5         2/2880
'   Status          0.75        3.5/5040
'   Transfer        0.75        4.25/6120
'   Date & Time     1.125       5.0/7200
'   User Name       0.625       6.125/8820
'   Comments        2.25        6.75/9720
'Mo Morris  14/5/2003
'   Transfer field removed from listing
'   Date & Time field expanded so that the timezone gap can be printed.
'   User Name field expanded to take 20 character User Names.
'   Comment field reduced in width.
'   Listing now has a "User Name - Full User Name" key printed at the end.
'
'   New Column widths (in inches/twips) as follows:-
'   FIELD           WIDTH       START AT
'   Question        1.5         .5/720
'   Response        1.5         2/2880
'   Status          0.75        3.5/5040
'   Date & Time     1.75        4.25/6120
'   User Name       1.25        6.0/8640
'   Comments        2.25        7.25/10440
'---------------------------------------------------------------------
Dim sPersonKey As String
Dim sPreviousPersonKey As String
Dim sHeadingText As String
Dim sVisitName As String
Dim sPreviousVisitName As String
Dim sForm As String
Dim sFormPrint As String
Dim sPreviousForm As String
Dim lYStorage As Long
Dim nHeadingWidth As Integer
Dim i As Integer
Dim sCommentline As String
Dim sChar As String
Dim sPrintText As String
Dim lYAfterDataItemName As Long
Dim lYAfterResponseValue As Long
Dim nAdditionalLines As Integer
Dim nHeadingLines As Integer
Dim nDataItemResponseLines As Integer
Dim nCommentLines As Integer
Dim lPrintingWidth As Long
Dim lRow As Long
Dim nCommentsWidth As Integer
Dim sWarningMessage As String
Const nSTARTCOMMENT As Integer = 10440
Dim colUserNames As Collection

    On Error Resume Next
    
    Set colUserNames = New Collection
    
    CommonDialog1.CancelError = True
    'Changed Mo 24/5/2002, Stemming from CBB 2.2.8.7
    'Printer.Orientation = vbPRORLandscape
    CommonDialog1.Orientation = cdlLandscape
    
    'WillC 10/5/00  SR3434 Added in following line of code to allow the user to choose a
    'printer from a number of printers.  Microsoft  Article ID Q254925
    Printer.TrackDefault = True
    
    CommonDialog1.ShowPrinter
    'check for errors in ShowPrinter (incuding a Cancel)
    If Err.Number > 0 Then Exit Sub
    
    'restore normal error trapping
    On Error GoTo PrinterError
    
    'Changed Mo 24/5/2002, Stemming from CBB 2.2.8.7
    Printer.Orientation = CommonDialog1.Orientation

    'set printer scalemode to twips
    Printer.ScaleMode = vbTwips
    Printer.FontSize = 8
    Printer.ScaleTop = -720
    Printer.ScaleLeft = -720
    
    'Note that on Windows 2000 systems the ealier "CommonDialog1.Orientation = cdlLandscape" line gets ignored
    'MLM 29/11/02: Changes to make printing possible on printer/paper combinations down to 10" printable area.
    If Printer.ScaleWidth > 15840 Then
        'printable area is > 11"; allow a .5" margin inside both sides of the printable area, and size the output to fill remaining space
        Printer.ScaleLeft = -720
        lPrintingWidth = Printer.ScaleWidth - 1440
    ElseIf Printer.ScaleWidth < 14410 Then
        'printable area < 10"; do not allow print.
        Call DialogError("The selected paper Size and Orientation are not wide enough for this listing." _
            & vbNewLine & "Make sure you have selected Landscape Orientation.", "Paper Width Problem")
        Exit Sub
    Else
        '10" < printable area < 11"; fix output at 10", centred in printable area
        lPrintingWidth = 14400
        Printer.ScaleLeft = -10 - (Printer.ScaleWidth - 14400) \ 2
    End If
    
    'Assess the width to be used by Comment/Reason For Change/Overrule Reason/Warning Message
    nCommentsWidth = lPrintingWidth - nSTARTCOMMENT
    Call PrintHeaderData(lPrintingWidth)
    
    sPreviousPersonKey = ""
    sPreviousVisitName = ""
    sPreviousForm = ""
    
    For lRow = 0 To UBound(mvData, 2)
        'The general approach here is to assess the number of lines to print a response (including
        'Subject, Visit and Form Heading lines) and then call PageEndcheck to see if there is enough space
        'on the current page to print it
        nHeadingLines = 0
        sPersonKey = mvData(dbcStudyName, lRow) & "/" _
            & mvData(dbcSite, lRow) & "/" & mvData(dbcSubjectId, lRow)
        If sPreviousPersonKey <> sPersonKey Then
            nHeadingLines = nHeadingLines + 1
            sPreviousVisitName = ""
            sPreviousForm = ""
        End If
        sVisitName = mvData(dbcVisitName, lRow) & "[" & mvData(dbcVisitCycleNumber, lRow) & "]"
        If sVisitName <> sPreviousVisitName Then
            nHeadingLines = nHeadingLines + 1
        End If
        'Mo 16/11/2004   Bug 2417
        sForm = mvData(dbcEFormTitle, lRow) & "[" & mvData(MACRODBBS30.dbcEFormCycleNumber, lRow) & "]"
        sFormPrint = eFormTitleLabel(mvData(dbcEFormTitle, lRow), RemoveNull(mvData(dbcEFormLabel, lRow)), mvData(MACRODBBS30.dbcEFormCycleNumber, lRow))
        If sForm <> sPreviousForm Then
            nHeadingLines = nHeadingLines + 1
            sPreviousForm = ""
        End If

        'Estimate the number of additional lines required to print the DataItemName, ResponseValue
        nDataItemResponseLines = 0
        If RemoveNull(mvData(dbcDataItemName, lRow)) <> "" Then
            'Check DataItemName width against 2160 (1.5 inches)
            If nDataItemResponseLines < Printer.TextWidth(mvData(dbcDataItemName, lRow)) \ 2160 Then
                nDataItemResponseLines = Printer.TextWidth(mvData(dbcDataItemName, lRow)) \ 2160
            End If
        End If
        If RemoveNull(mvData(dbcResponseValue, lRow)) <> "" Then
            'check ResponseValue width against 2160 (1.5 inch)
            If nDataItemResponseLines < Printer.TextWidth(mvData(dbcResponseValue, lRow)) \ 2160 Then
                nDataItemResponseLines = Printer.TextWidth(mvData(dbcResponseValue, lRow)) \ 2160
            End If
        End If
        
        'Estimate the number of lines required to print Comment/Reason For Change/Overrule Reason/Warning Message
        nCommentLines = 0
        'Assess Comment length
        If RemoveNull(mvData(dbcComments, lRow)) <> "" Then
            'Add a line for the comment heading
            nCommentLines = nCommentLines + 1
            sCommentline = ""
            For i = 1 To Len(mvData(dbcComments, lRow))
                sChar = Mid(mvData(dbcComments, lRow), i, 1)
                If Asc(sChar) = 13 Then
                    'assess width of this part of the comment, which might wrap onto another line
                    nCommentLines = nCommentLines + (Printer.TextWidth(sCommentline) \ nCommentsWidth) + 1
                    'after detecting a CR the LF can be skipped
                    i = i + 1
                    sCommentline = ""
                Else
                    sCommentline = sCommentline & sChar
                End If
            Next
        End If
        'Assess Reason For Change length
        If RemoveNull(mvData(dbcReasonForChange, lRow)) <> "" Then
            '1 is added for the Reason For Change heading
            nCommentLines = nCommentLines + (Printer.TextWidth(mvData(dbcReasonForChange, lRow)) \ nCommentsWidth) + 2
        End If
        'Assess Overrule Reason length
        If RemoveNull(mvData(dbcOverruleReason, lRow)) <> "" Then
            '1 is added for the Overrule Reason heading
            nCommentLines = nCommentLines + (Printer.TextWidth(mvData(dbcOverruleReason, lRow)) \ nCommentsWidth) + 2
        End If
        'Assess Validation Message length
        If RemoveNull(mvData(dbcValMessage, lRow)) <> "" Then
            '1 is added for the Validation Message heading
            nCommentLines = nCommentLines + (Printer.TextWidth(mvData(dbcValMessage, lRow)) \ nCommentsWidth) + 2
        End If
        
        'set nAdditionalLines to the greater of nDataItemResponseLines and nCommentLines
        nAdditionalLines = nDataItemResponseLines
        If (nCommentLines - 1) > nDataItemResponseLines Then
            nAdditionalLines = nCommentLines - 1
        End If
        'Add the number of Heading Lines
        nAdditionalLines = nAdditionalLines + nHeadingLines
        
        If nAdditionalLines > 0 Then
            Call PageEndCheck(lPrintingWidth, nAdditionalLines)
        End If
        
        'Print Header Lines
        If sPreviousPersonKey <> sPersonKey Then
            'Format and display Person Heading
            sPreviousPersonKey = sPersonKey
            lYStorage = Printer.CurrentY
            'Mo Morris  19/7/02, CBB 2.2.15.5, Subject Label added to sHeadingText
            sHeadingText = "Study: " & mvData(dbcStudyName, lRow) _
                 & "     Site: " & mvData(dbcSite, lRow) _
                 & "     Subject: " & mvData(dbcSubjectId, lRow) _
                 & "     Label: " & RemoveNull(mvData(dbcSubjectLabel, lRow))
            Printer.FontBold = True
            nHeadingWidth = Printer.TextWidth(sHeadingText)
            Printer.Line (-10, lYStorage)-(nHeadingWidth + 20, lYStorage + 240), , B
            Printer.CurrentY = lYStorage + 30
            Printer.CurrentX = 0
            Printer.Print sHeadingText
            Printer.FontBold = False
            Printer.CurrentY = lYStorage + 270
        End If
        If sVisitName <> sPreviousVisitName Then
            'Format and display Visit Heading
            sPreviousVisitName = sVisitName
            lYStorage = Printer.CurrentY
            sHeadingText = "Visit: " & sVisitName
            Printer.FontBold = True
            nHeadingWidth = Printer.TextWidth(sHeadingText)
            Printer.Line (230, lYStorage)-(nHeadingWidth + 260, lYStorage + 240), , B
            Printer.CurrentY = lYStorage + 30
            Printer.CurrentX = 240
            Printer.Print sHeadingText
            Printer.FontBold = False
            Printer.CurrentY = lYStorage + 270
        End If
        If sForm <> sPreviousForm Then
            'Format and display Form Heading
            sPreviousForm = sForm
            lYStorage = Printer.CurrentY
            'Mo 16/11/2004   Bug 2417
            sHeadingText = "eForm: " & sFormPrint
            Printer.FontBold = True
            nHeadingWidth = Printer.TextWidth(sHeadingText)
            Printer.Line (470, lYStorage)-(nHeadingWidth + 500, lYStorage + 240), , B
            Printer.CurrentY = lYStorage + 30
            Printer.CurrentX = 480
            Printer.Print sHeadingText
            Printer.FontBold = False
            Printer.CurrentY = lYStorage + 270
        End If
        
        'Store Y position at start of line
        lYStorage = Printer.CurrentY
        
        'Changed Mo Morris 2/2/01
        If RemoveNull(mvData(dbcDataItemName, lRow)) <> "" Then
            sPrintText = mvData(dbcDataItemName, lRow)
            'include cycle number if its a Repeating Question Group(OwnerQGroupId > 0)
            If mvData(dbcOwnerQGroupId, lRow) <> 0 Then
                sPrintText = sPrintText & "[" & mvData(dbcResponseCycleNumber, lRow) & "]"
            End If
            
            'Check DataItemName width against 2160 (1.5 inches)
            If Printer.TextWidth(sPrintText) > 2160 Then sPrintText = SplitTextLine(sPrintText, 2160, 720)
            Printer.CurrentX = 720                          '1/2 inch
            Printer.Print sPrintText
        End If
        'Store Y position after printing DataItemName
        lYAfterDataItemName = Printer.CurrentY
        
        'reset the Y position after printing DataItemName
        Printer.CurrentY = lYStorage
        
        ' Localise value before display - NCJ 17/11/99
        ' NB This is not formatted
        'changed Mo Morris 3/3/00 don't call LocaliseStandardValue with Null values
        If Not IsNull(mvData(dbcResponseValue, lRow)) Then
            sPrintText = LocaliseStandardValue((mvData(dbcResponseValue, lRow)), (mvData(dbcDataType, lRow)))
            'check ResponseValue width against 2160 (1.5 inches)
            If Printer.TextWidth(sPrintText) > 1980 Then sPrintText = SplitTextLine(sPrintText, 2160, 2880)
            Printer.CurrentX = 2880
            Printer.Print sPrintText
        End If
        'Store Y position after printing ResponseValue
        lYAfterResponseValue = Printer.CurrentY
        
        'reset the Y position after printing ResponseValue
        Printer.CurrentY = lYStorage
        
        Printer.CurrentX = 5040                         '3.5 inch
        Printer.Print GetStatusText((mvData(dbcResponseStatus, lRow))), ;
        Printer.CurrentX = 6120                         '4.25 inch
        'Cast timestamp into a date and add the timezone gap
        
        'Printer.Print Format(CDate(mvData(dbcResponseTimeStamp, lRow)), "yyyy/mm/dd hh:mm:ss") & _
            " (GMT" & IIf(mvData(dbcResponseTimestamp_TZ, lRow) < 0, "+", "") & -mvData(dbcResponseTimestamp_TZ, lRow) \ 60 & ":" & Format(Abs(mvData(dbcResponseTimestamp_TZ, lRow)) Mod 60, "00") & ")", ;
        'TA 22/05/2003: use function
        Printer.Print DisplayGMTTime(mvData(dbcResponseTimeStamp, lRow), "yyyy/mm/dd hh:mm:ss", mvData(dbcResponseTimestamp_TZ, lRow)), ;

        Printer.CurrentX = 8640                         '6.0 inch
        Printer.Print mvData(dbcUserName, lRow), ;
        
        'Add "UserName - FullUserName" to the Collection of UserNames
        On Error Resume Next
        colUserNames.Add mvData(dbcUserName, lRow) & " - " & mvData(dbcFullUserName, lRow), mvData(dbcUserName, lRow)
        If Err.Number <> 0 Then
            'clear the already in collection error
            Err.Clear
        End If
        'restore normal error trapping
        On Error GoTo PrinterError
        
        'Print Comment
        'Note that comments are delimited by CR (Asci 13) and LF (Asci 10) characters
        'Note that comments end with a CR/LF
        If RemoveNull(mvData(dbcComments, lRow)) <> "" Then
            lYStorage = Printer.CurrentY
            Printer.FontBold = True
            nHeadingWidth = Printer.TextWidth("Comment:-")
            Printer.Line (nSTARTCOMMENT - 10, lYStorage - 10)-(nSTARTCOMMENT + 10 + nHeadingWidth, lYStorage + 190), , B
            Printer.CurrentY = lYStorage
            Printer.CurrentX = nSTARTCOMMENT
            Printer.Print "Comment:-"
            Printer.FontBold = False
            sCommentline = ""
            For i = 1 To Len(mvData(dbcComments, lRow))
                sChar = Mid(mvData(dbcComments, lRow), i, 1)
                If Asc(sChar) = 13 Then
                    'check for comment line being greater than 2 nCommenWidth
                    If Printer.TextWidth(sCommentline) > nCommentsWidth Then sCommentline = SplitTextLine(sCommentline, nCommentsWidth, nSTARTCOMMENT)
                    Printer.CurrentX = nSTARTCOMMENT             '6.75 inch
                    Printer.Print sCommentline
                    'after detecting a CR the LF can be skipped
                    i = i + 1
                    sCommentline = ""
                Else
                    sCommentline = sCommentline & sChar
                End If
            Next
        End If
        
        'Print Reason For Change
        If RemoveNull(mvData(dbcReasonForChange, lRow)) <> "" Then
            lYStorage = Printer.CurrentY
            Printer.FontBold = True
            nHeadingWidth = Printer.TextWidth("Reason For Change:-")
            Printer.Line (nSTARTCOMMENT - 10, lYStorage - 10)-(nSTARTCOMMENT + 10 + nHeadingWidth, lYStorage + 190), , B
            Printer.CurrentY = lYStorage
            Printer.CurrentX = nSTARTCOMMENT
            Printer.Print "Reason For Change:-"
            Printer.FontBold = False
            sPrintText = mvData(dbcReasonForChange, lRow)
            'check for Reason for change line being greater than nCommentsWidth
            If Printer.TextWidth(sPrintText) > nCommentsWidth Then sPrintText = SplitTextLine(sPrintText, nCommentsWidth, nSTARTCOMMENT)
            Printer.CurrentX = nSTARTCOMMENT
            Printer.Print sPrintText
        End If
        
        'Print Overrule Reason
        If RemoveNull(mvData(dbcOverruleReason, lRow)) <> "" Then
            lYStorage = Printer.CurrentY
            Printer.FontBold = True
            nHeadingWidth = Printer.TextWidth("Overrule Reason:-")
            Printer.Line (nSTARTCOMMENT - 10, lYStorage - 10)-(nSTARTCOMMENT + 10 + nHeadingWidth, lYStorage + 190), , B
            Printer.CurrentY = lYStorage
            Printer.CurrentX = nSTARTCOMMENT
            Printer.Print "Overrule Reason:-"
            Printer.FontBold = False
            sPrintText = mvData(dbcOverruleReason, lRow)
            'check for Overrule Reason line being greater than nCommentsWidth
            If Printer.TextWidth(sPrintText) > nCommentsWidth Then sPrintText = SplitTextLine(sPrintText, nCommentsWidth, nSTARTCOMMENT)
            Printer.CurrentX = nSTARTCOMMENT
            Printer.Print sPrintText
        End If
        
        'Print Validation Message
        'Note that LabTest Warning messages can contain CR/LF, but unlike comments do not end with CR/LF
        If RemoveNull(mvData(dbcValMessage, lRow)) <> "" Then
            'add a CR/LF to WarningMessage so that CR/LF processing code works
            sWarningMessage = mvData(dbcValMessage, lRow) & vbNewLine
            lYStorage = Printer.CurrentY
            Printer.FontBold = True
            nHeadingWidth = Printer.TextWidth("Validation Message:-")
            Printer.Line (nSTARTCOMMENT - 10, lYStorage - 10)-(nSTARTCOMMENT + 10 + nHeadingWidth, lYStorage + 190), , B
            Printer.CurrentY = lYStorage
            Printer.CurrentX = nSTARTCOMMENT
            Printer.Print "Validation Message:-"
            Printer.FontBold = False
            sCommentline = ""
            For i = 1 To Len(sWarningMessage)
                sChar = Mid(sWarningMessage, i, 1)
                If Asc(sChar) = 13 Then
                    'check for comment line being greater than 2 nCommenWidth
                    If Printer.TextWidth(sCommentline) > nCommentsWidth Then sCommentline = SplitTextLine(sCommentline, nCommentsWidth, nSTARTCOMMENT)
                    Printer.CurrentX = nSTARTCOMMENT             '6.75 inch
                    Printer.Print sCommentline
                    'after detecting a CR the LF can be skipped
                    i = i + 1
                    sCommentline = ""
                Else
                    sCommentline = sCommentline & sChar
                End If
            Next

        End If
        
        'CurrentY includes the number of lines to print Comment/Reason For Change/Overrule Reason/Warning Message
        'set CurrentY to the greater of CurrentY, lYAfterDataItemName, lYAfterResponseValue & lYAfterComment
        If Printer.CurrentY < lYAfterDataItemName Then
            Printer.CurrentY = lYAfterDataItemName
        End If
        If Printer.CurrentY < lYAfterResponseValue Then
            Printer.CurrentY = lYAfterResponseValue
        End If
        
        Call PageEndCheck(lPrintingWidth)
        
    Next
    
    'Print the "User Name - Full User Name" key that is contained in colUserNames
    'Check for the need to throw a new page
    If (Printer.CurrentY + ((colUserNames.Count + 2) * 300)) > 9780 Then
        Printer.NewPage
        Call PrintHeaderData(lPrintingWidth)
    End If
    Printer.Print
    Printer.FontBold = True
    lYStorage = Printer.CurrentY
    nHeadingWidth = Printer.TextWidth("User Name - Full User Name Key")
    Printer.Line (-10, lYStorage)-(nHeadingWidth + 20, lYStorage + 240), , B
    Printer.CurrentY = lYStorage + 30
    Printer.CurrentX = 0
    Printer.Print "User Name - Full User Name Key"
    Printer.FontBold = False
    Printer.CurrentY = lYStorage + 270
    For i = 1 To colUserNames.Count
        Printer.Print colUserNames.Item(i)
    Next
    
    Printer.EndDoc
    
Exit Sub

PrinterError:

MsgBox "A printer error has occurred.  The error number is " & Err.Number, vbOKOnly + vbInformation

End Sub

'---------------------------------------------------------------------
Private Sub PrintHeaderData(PrintingWidth As Long)
'---------------------------------------------------------------------
'Mo Morris  19/7/02, CBB 2.2.15.5
'Mo Morris  14/5/2003, minor heading changes
'---------------------------------------------------------------------
Dim nCurrentY As Integer
Dim sHeaderText As String

    On Error GoTo ErrHandler

    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.Print "MACRO - Subject Data";
    sHeaderText = "Printed " & Format(Now, "yyyy/mm/dd hh:mm:ss") & "    Page " & Printer.Page
    Printer.CurrentX = PrintingWidth - Printer.TextWidth(sHeaderText)
    Printer.Print sHeaderText
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.DrawWidth = 4
    Printer.CurrentY = Printer.CurrentY + 30
    nCurrentY = Printer.CurrentY
    Printer.Line (0, nCurrentY)-(PrintingWidth, nCurrentY)
    Printer.DrawWidth = 1
    Printer.CurrentY = Printer.CurrentY + 60
    nCurrentY = Printer.CurrentY
    'Heading banner from 1/2 inch to 10 1/4 inch on A4 paper
    Printer.Line (710, nCurrentY)-(PrintingWidth, nCurrentY + 240), , B
    Printer.CurrentY = nCurrentY + 30
    Printer.CurrentX = 720                              '1/2 inch
    Printer.Print "Question", ;
    Printer.CurrentX = 2880                             '2 inch
    Printer.Print "Value", ;
    Printer.CurrentX = 5040                             '3 1/2 inch
    Printer.Print "Status", ;
    Printer.CurrentX = 6120                             '4 1/4 inch
    Printer.Print "Date & Time", ;
    Printer.CurrentX = 8640                             '6 inch
    Printer.Print "User Name", ;
    Printer.CurrentX = 10440                            '7 1/4 inch
    Printer.Print "Comment/RFC/Overrule Reason/Validation Message"
    Printer.CurrentY = nCurrentY + 270
    
Exit Sub
ErrHandler:
    'Changed 22/6/00 SR 3640
    Select Case Err.Number
        Case 482
            MsgBox "Printer error number 482 has occurred.", vbInformation, "MACRO"
        Case Else
            Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PrintHeaderData")
                Case OnErrorAction.Ignore
                    Resume Next
                Case OnErrorAction.Retry
                    Resume
                Case OnErrorAction.QuitMACRO
                    Call ExitMACRO
                    Call MACROEnd
            End Select
    End Select
   
End Sub

'---------------------------------------------------------------------
Private Sub PageEndCheck(PrintingWidth As Long, Optional ByVal ExtraLines As Integer = 0)
'---------------------------------------------------------------------
'Mo Morris  19/7/02, CBB 2.2.15.5
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    'note that 9780 is 10080-300
    'i.e. 7 inches (7*1440 twips) - space required to print a line (300 twips)
    If (Printer.CurrentY + (ExtraLines * 300)) > (9780) Then
        Printer.NewPage
        If mDRType = dbteForms Then
            Call PrintHeaderForms(PrintingWidth)
        Else
            Call PrintHeaderData(PrintingWidth)
        End If
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PageEndCheck")
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
Private Sub PrintForms()
'---------------------------------------------------------------------
'Changed by  Mo Morris   11/9/98     SPR 426 (change made in Released and Developed versions)
'CancelError on CommonDialog1 enabled and tested for after the ShowPrinter call.
' NCJ 30/11/99 - Expect timestamps to be doubles
'SDM 02/12/99   Changed to Private
'Mo Morris  19/7/02, CBB 2.2.15.5, minor clean up changes made to the forms listing.
'   VisitCycleNumber added to sPreviousVisit.
'   CRFPageCyccleNumber added to sPreviousForm.
'   Subject Label added to listing.
'---------------------------------------------------------------------
Dim sPreviousTrial As String
Dim sPreviousSite As String
Dim lPreviousPersonId As Long
Dim sPreviousVisitCycle As String
Dim sPreviousFormCycle As String
Dim lPrintingWidth As Long
Dim lRow As Long
Dim sCurrentVisitCycle As String
Dim sCurrentFormCycle As String

    On Error Resume Next
    
    CommonDialog1.CancelError = True
    
    'Changed Mo 24/5/2002, Stemming from CBB 2.2.8.7
    'Printer.Orientation = vbPRORLandscape
    CommonDialog1.Orientation = cdlLandscape
    'WillC 10/5/00  SR3434 Added in following line of code to allow the user to choose a
    'printer from a number of printers.  Microsoft  Article ID Q254925
    Printer.TrackDefault = True

    CommonDialog1.ShowPrinter
    'check for errors in ShowPrinter (incuding a Cancel)
    If Err.Number > 0 Then Exit Sub
    
    'restore normal error trapping
    On Error GoTo 0
    
    'Changed Mo 24/5/2002, Stemming from CBB 2.2.8.7
    Printer.Orientation = CommonDialog1.Orientation
    
    'set printer scalemode to twips
    Printer.ScaleMode = vbTwips
    Printer.FontSize = 8
    'extend the top and left margins by 1/2 inch to 3/4 inch approx
    Printer.ScaleLeft = -720
    Printer.ScaleTop = -720
    
    'Detext the printing width of the paper on the selected printer minus 2 * 1/2 inch borders (1440 twips)
    lPrintingWidth = Printer.ScaleWidth - 1440
    'Changed Mo Morris 22/7/2002
    'Check that width of the selected paper size/orienation is wide enough minimum of 10 inches (14,400 twips)
    'Note that on Windows 2000 systems the ealier "CommonDialog1.Orientation = cdlLandscape" line gets ignored
    If lPrintingWidth < 14400 Then
        Call DialogError("The selected paper Size and Orientation are not wide enough for this listing." _
            & vbNewLine & "Make sure you have selected Landscape Orientation.", "Paper Width Problem")
        Exit Sub
    End If
    Call PrintHeaderForms(lPrintingWidth)
    
    sPreviousTrial = ""
    sPreviousSite = ""
    lPreviousPersonId = 0
    sPreviousVisitCycle = ""
    sPreviousFormCycle = ""
    For lRow = 0 To UBound(mvData, 2)
        If mvData(dbcStudyName, lRow) <> sPreviousTrial Then
            sPreviousTrial = mvData(dbcStudyName, lRow)
            sPreviousSite = ""
            lPreviousPersonId = 0
            sPreviousVisitCycle = ""
            sPreviousFormCycle = ""
            Printer.CurrentX = 0
            Printer.Print mvData(dbcStudyName, lRow), ;
        End If
        If mvData(dbcSite, lRow) <> sPreviousSite Then
            sPreviousSite = mvData(dbcSite, lRow)
            lPreviousPersonId = 0
            sPreviousVisitCycle = ""
            sPreviousFormCycle = ""
            Printer.CurrentX = 1620                            '1 1/8 inch
            Printer.Print mvData(dbcSite, lRow), ;
        End If
        If mvData(dbcSubjectId, lRow) <> lPreviousPersonId Then
            lPreviousPersonId = mvData(dbcSubjectId, lRow)
            sPreviousVisitCycle = ""
            sPreviousFormCycle = ""
            Printer.CurrentX = 2520                            '1 3/4 inch
            Printer.Print mvData(dbcSubjectId, lRow) & "/" & RemoveNull(mvData(dbcSubjectLabel, lRow)), ;
        End If
        sCurrentVisitCycle = mvData(dbcVisitName, lRow) & "|" & mvData(dbcVisitCycleNumber, lRow)
        If sCurrentVisitCycle <> sPreviousVisitCycle Then
            sPreviousVisitCycle = sCurrentVisitCycle
            sPreviousFormCycle = ""
            Printer.CurrentX = 6480                           '4 1/2 inch
            Printer.Print mvData(dbcVisitName, lRow) & "[" & mvData(dbcVisitCycleNumber, lRow) & "]", ;
        End If
        sCurrentFormCycle = mvData(dbcEFormTitle, lRow) & "|" & mvData(dbcEFormCycleNumber, lRow)
        If mvData(dbcEFormTitle, lRow) <> sPreviousFormCycle Then
            sPreviousFormCycle = sCurrentFormCycle
            Printer.CurrentX = 10440                             '7 1/4 inch
            Printer.Print eFormTitleLabel(mvData(dbcEFormTitle, lRow), RemoveNull(mvData(dbcEFormLabel, lRow)), mvData(MACRODBBS30.dbcEFormCycleNumber, lRow)), ;
        End If
        Printer.CurrentX = 11880                               '8 1/4 inch
        Printer.Print GetStatusText((mvData(dbcEFormStatus, lRow))), ;
        Printer.CurrentX = 12960                                '9 inch
        ' Cast timestamp into a date - NCJ 30/11/99
        'Mo Morris 1/12/00, check for the existence of a date
        If Not IsNull(mvData(dbcResponseTimeStamp, lRow)) Then
            If Val(mvData(dbcResponseTimeStamp, lRow)) > 0 Then
                Printer.Print Format(CDate(mvData(dbcResponseTimeStamp, lRow)), "yyyy/mm/dd hh:mm:ss"), ;
            End If
        End If
        'end the current line
        Printer.Print ""
        Call PageEndCheck(lPrintingWidth)
        
    Next

    Printer.EndDoc

End Sub

'---------------------------------------------------------------------
Private Sub PrintHeaderForms(PrintingWidth As Long)
'---------------------------------------------------------------------
'Mo Morris  19/7/02, CBB 2.2.15.5
'   Changes to column names
'   Changes to column widths
'   Label added alongside subject (PersonId)
'---------------------------------------------------------------------
Dim nCurrentY As Integer
Dim sHeaderText As String

    On Error GoTo ErrHandler

    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.Print "MACRO - Subject Data";
    sHeaderText = "Printed " & Format(Now, "yyyy/mm/dd hh:mm:ss") & "    Page " & Printer.Page
    Printer.CurrentX = PrintingWidth - Printer.TextWidth(sHeaderText)
    Printer.Print sHeaderText
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.DrawWidth = 4
    Printer.CurrentY = Printer.CurrentY + 30
    nCurrentY = Printer.CurrentY
    Printer.Line (0, nCurrentY)-(PrintingWidth, nCurrentY)
    Printer.DrawWidth = 1
    Printer.CurrentY = Printer.CurrentY + 60
    nCurrentY = Printer.CurrentY
    'Heading banner from 1/2 inch to 10 1/4 inch
    Printer.Line (-10, nCurrentY)-(PrintingWidth, nCurrentY + 240), , B
    Printer.CurrentY = nCurrentY + 30
    Printer.CurrentX = 0
    Printer.Print "Study", ;
    Printer.CurrentX = 1620                             '1 1/8 inch
    Printer.Print "Site", ;
    Printer.CurrentX = 2520                             '1 3/4 inch
    Printer.Print "Subject/Label", ;
    Printer.CurrentX = 6480                             '4 1/2 inch
    Printer.Print "Visit", ;
    Printer.CurrentX = 10440                            '7 1/4 inch
    Printer.Print "eForm", ;
    Printer.CurrentX = 11880                            '8 1/4 inch
    Printer.Print "Status", ;
    Printer.CurrentX = 12960                            '9 inch
    Printer.Print "Date"
    Printer.CurrentY = nCurrentY + 270
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PrintHeaderForms")
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
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------------------
'reset all module level variables when unloading
'---------------------------------------------------------------------

    ' Store row and column clicked
    mlColumnClicked = 0
    mlRowClicked = 0
    ' Store range of rows covered by click
    mlRowClickedMin = 0
    mlRowClickedMax = 0
    ' NCJ 29/2/00 - Changed to use new clsColCoords
    ' Each collection class contains the coordinates of cells that are locked or frozen
    Set mcolLocked = Nothing
    Set mcolFrozen = Nothing

    ' NCJ 9/3/00 - Store our "launch mode", i.e. either Monitor or Subject
    mnLaunchMode = eMACROWindow.None
    'Arrays to hold recordset data and whether a row has been shown
     mvData = ""
     mvRowShown = ""
    mDRType = MACRODBBS30.eDataBrowserType.dbtDataItemResponse
    'inform everyone that i'm closed
    CloseWinForm wfDataBrowser
    
End Sub

'---------------------------------------------------------------------
Private Sub lcmdclose_Click()
'---------------------------------------------------------------------
    
    Unload Me

End Sub


'---------------------------------------------------------------------
Private Sub AdjustRowHeight(ByVal RowIndex As Long)
'---------------------------------------------------------------------
' NCJ 3/3/00
' Adjust row height according to height of Comments & RFC
' NB Use lblCommentSize for "sizing" the text
' Min. height is height of question status icon
' NCJ 19/5/00 - Include Overrule reason in calculations
' NCJ 17/10/00 SR3864 - Include Validation Message in calculations
' NCJ 6/11/00 - Make sure row height is never 0!
' RS 18/02/2003 - Adjust rowheight for multiple comments
'---------------------------------------------------------------------
Dim sComments As String
Dim sRFC As String
Dim sOverrule As String
Dim sValMessage As String
Dim lCurCol As Long
Dim lCurRow As Long
Dim sSubject As String
Dim sResponse As String

    
    lblCommentSize.WordWrap = True
    
    With flexData
        
        ' RS Set Form Font for TextHeight properties
        'Me.FontName = flexData.Font.Name
        'Me.FontSize = flexData.Font.SIZE
        'Me.FontBold = flexData.Font.Bold
        'Me.FontItalic = flexData.Font.Italic
        
        ' Save current settings
        lCurRow = .Row
        lCurCol = .Col
        
        .Row = RowIndex
        .Col = mnTRIALSITEPERSON_COL
        
        ' Default height
        ' NCJ 6/11/00 - PictureHeight may be 0 in Forms view with no patient data
        ' RS 22/01/2003 - Adjust actual picture Height, otherwise all rows too high
        If .CellPicture.Height > 0 Then
            .RowHeight(RowIndex) = .CellPicture.Height * 0.7
        Else
            ' NCJ Set height to text height plus a bit (times 1.5 seems good)
            .RowHeight(RowIndex) = CLng(TextHeight("X") * 1.5)
        End If
        
        ' RS 12/03.2003: If there is only a single dataitem for a subject, make the rowheight at least 2 lines
        ' to make sure that the subject label is displayed.
        ' RS 14/03/2003: Section moved to just before Exit of PopulateGridSection, as the grid is not completely filled at this
        ' point and it isnot possible to compare current row with next row

        sComments = .TextMatrix(RowIndex, mnCOMMENT_COL)
        sRFC = .TextMatrix(RowIndex, mnRFC_COL)
        sOverrule = .TextMatrix(RowIndex, mnOVERRULE_COL)
        sValMessage = .TextMatrix(RowIndex, mnVAL_MESSAGE_COL)
        sResponse = .TextMatrix(RowIndex, mnDATARESPONSE_COL)
        
        'TA WE NEED SONETHING LIKE THE FOLLOWING CODE TO ENSURE FULL SUBJECT LABEL IS SEEN IN
'        ' TA 26/02/2003: include trial/site/subject - only want to do this if there is one row per subject
'       sSubject = .TextMatrix(RowIndex, mnTRIALSITEPERSON_COL)
'        If sSubject > "" Then
'            lblCommentSize.Width = .ColWidth(mnTRIALSITEPERSON_COL)
'            lblCommentSize.Caption = trial / Site / Subject
'            ' Enlarge to fit Comments
'            If .RowHeight(RowIndex) < lblCommentSize.Height * 1.22 Then
'                .RowHeight(RowIndex) = lblCommentSize.Height * 1.22
'            End If
'        End If
        
        ' Check Comments, RFC, Overrule and Validation Message
        ' RS multiply height by 1.22, otherwise not all lines are displayed
        If sComments > "" Then
            lblCommentSize.Width = .ColWidth(mnCOMMENT_COL)
            lblCommentSize.Caption = sComments
            ' Enlarge to fit Comments
            If .RowHeight(RowIndex) < lblCommentSize.Height * 1.22 Then
                .RowHeight(RowIndex) = lblCommentSize.Height * 1.22
            End If
        End If
        
        If sRFC > "" Then
            lblCommentSize.Width = .ColWidth(mnRFC_COL)
            lblCommentSize.Caption = sRFC
            ' Enlarge to fit RFC
            If .RowHeight(RowIndex) < lblCommentSize.Height Then
                .RowHeight(RowIndex) = lblCommentSize.Height
            End If
        End If
        
        ' NCJ 19/5/00
        If sOverrule > "" Then
            lblCommentSize.Width = .ColWidth(mnOVERRULE_COL)
            lblCommentSize.Caption = sOverrule
            ' Enlarge to fit Overrule
            If .RowHeight(RowIndex) < lblCommentSize.Height Then
                .RowHeight(RowIndex) = lblCommentSize.Height
            End If
        End If
        
        ' NCJ 17/10/00
        If sValMessage > "" Then
            lblCommentSize.Width = .ColWidth(mnVAL_MESSAGE_COL)
            lblCommentSize.Caption = sValMessage
            ' Enlarge to fit Validation Message
            If .RowHeight(RowIndex) < lblCommentSize.Height Then
                .RowHeight(RowIndex) = lblCommentSize.Height
            End If
        End If
     
        ' TA 05/11/2004: make sure height of row shows full response
        If sResponse > "" Then
            lblCommentSize.Width = .ColWidth(mnDATARESPONSE_COL)
            lblCommentSize.Caption = sResponse
            ' Enlarge to fit response
            If .RowHeight(RowIndex) < lblCommentSize.Height Then
                .RowHeight(RowIndex) = lblCommentSize.Height
            End If
        End If
     
     
     
        ' Restore current settings
        .Row = lCurRow
        .Col = lCurCol
        
    End With    ' flexData
    
End Sub

'-------------------------------------------------------------------------------------------------
Private Sub AdjustLastRowHeight(ByRef lRow As Long)
'-------------------------------------------------------------------------------------------------
' MLM 05/06/02: added.
'-------------------------------------------------------------------------------------------------

Dim lTextHeight As Long

    With flexData
        lTextHeight = TextHeight(.TextMatrix(lRow, mnTRIALSITEPERSON_COL)) * 1.1
        If .RowHeight(lRow) < lTextHeight Then .RowHeight(lRow) = lTextHeight
    End With
    
End Sub
'---------------------------------------------------------------------
Private Sub AdjustColumnWidth(ByVal ColumnIndex As Long)
'---------------------------------------------------------------------
'   SDM 06/12/99
'---------------------------------------------------------------------
Dim lWidth As Long
Dim lMaxWidth
    On Error GoTo ErrHandler
    
    With flexData
        .Col = ColumnIndex

        'MLM 06/06/03: This moved to GetMaxColumnTextWidth
        'TA 08/11/2004: CBD 2425 CRM 990 - reinstated for response value so the row height sizing works
        Select Case ColumnIndex
    '    Case mnTRIALSITEPERSON_COL: lMaxWidth = 2000
    '    Case mnVISIT_COL: lMaxWidth = 2000
    '    Case mnFORM_COL: lMaxWidth = 2000
    '    Case mnDATAITEM_COL: lMaxWidth = 2000
        Case mnDATARESPONSE_COL: lMaxWidth = 2000
    '    Case mnSTATUS_COL: lMaxWidth = 700
        Case Else: lMaxWidth = 10000
        End Select

        'calculate width based on text passed in (taking into account max width
        lWidth = Min(TextWidth(.Text & "    "), lMaxWidth)
        If .ColWidth(ColumnIndex) < lWidth Then
            .ColWidth(ColumnIndex) = lWidth
        End If
    End With
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDataItemResponse.AdjustColumnWidth"

End Sub

'---------------------------------------------------------------------
Private Function GetMaxColumnTextWidth(lColumnIndex As Long)
'---------------------------------------------------------------------
' RS 12/03/2003
'
' Calculate the width of the widest text in the given columnm,
' checking all rows, including header row
'
' Called just before exiting PopulateGridSection
'
'MLM 06/06/03: Consideration of max colum widths moved here from AdjustColumnWidth
'---------------------------------------------------------------------
Dim lMaxWidth As Long
Dim lWidth As Long
Dim lRow As Long

    lWidth = 0
    For lRow = 0 To flexData.Rows - 1
        If TextWidth(flexData.TextMatrix(lRow, lColumnIndex)) > lWidth Then
            lWidth = TextWidth(flexData.TextMatrix(lRow, lColumnIndex))
        End If
    Next
    
    Select Case lColumnIndex
    Case mnTRIALSITEPERSON_COL: lMaxWidth = 2000
    Case mnVISIT_COL: lMaxWidth = 2000
    Case mnFORM_COL: lMaxWidth = 2000
    Case mnDATAITEM_COL: lMaxWidth = 2000
    Case mnDATARESPONSE_COL: lMaxWidth = 2000
    Case mnSTATUS_COL: lMaxWidth = 700
    Case Else: lMaxWidth = 10000
    End Select
    
    If lMaxWidth < lWidth Then
        GetMaxColumnTextWidth = lMaxWidth
    Else
        GetMaxColumnTextWidth = lWidth
    End If
    
End Function

'---------------------------------------------------------------------
Private Sub SetCellLockUnlockFreeze(ByVal nSetting As LockStatus, Optional bUnfreezing As Boolean = False)
'---------------------------------------------------------------------
' NCJ 4th May 2000
' Set the Lock/Unlock/Freeze status for the SINGLE current flexData cell ONLY
' (i.e. don't ripple effects through to the left or right)
' Code based on what was in SetLockUnlockFreeze
' Do not ever change the status of a frozen cell - UNLESS bUnfreezing = TRUE
' NCJ 23 Dec 02 - Added bUnfreezing, and dealt with changes from Frozen
'---------------------------------------------------------------------
Dim lCol As Long
Dim lRow As Long

    On Error GoTo ErrHandler
    
    lRow = flexData.Row
    lCol = flexData.Col
    
    ' If it's frozen, do not change its status unless we're Unfreezing
    If mcolFrozen.IsItem(lRow, lCol) And Not bUnfreezing Then
        SetIcon lCol, lRow, vbHighlight
        Exit Sub
    End If
    
    Select Case nSetting
        Case LockStatus.lsUnlocked
            If flexData.CellBackColor = vbHighlight Then
                flexData.CellForeColor = vbHighlightText
            
            Else
                flexData.CellForeColor = vbWindowText
            End If
            ' If it's locked, remove from our "locked" collection
            If mcolLocked.IsItem(lRow, lCol) Then
                Call mcolLocked.Remove(lRow, lCol)
            End If
            ' If it was frozen, remove from our "frozen" collection
            If mcolFrozen.IsItem(lRow, lCol) Then
                Call mcolFrozen.Remove(lRow, lCol)
            End If

            ' RS 24/01/2003 Update Array & StatusIcons
            Select Case lCol
                Case mnTRIALSITEPERSON_COL:   mvData(dbcSubjectLockStatus, lRow - 1) = LockStatus.lsUnlocked
                Case mnVISIT_COL:             mvData(dbcVisitLockStatus, lRow - 1) = LockStatus.lsUnlocked
                Case mnFORM_COL:              mvData(dbcEFormLockStatus, lRow - 1) = LockStatus.lsUnlocked
                Case mnSTATUS_COL:            mvData(dbcDataItemLockStatus, lRow - 1) = LockStatus.lsUnlocked
            End Select
            SetIcon lCol, lRow, flexData.CellBackColor

        Case LockStatus.lsLocked
        
            'Set colour of text
            If flexData.CellBackColor = vbHighlight Then
                flexData.CellForeColor = vbHighlightText
                flexData.CellBackColor = mnLOCKED_COLOUR
                
            
            Else
                flexData.CellForeColor = mnLOCKED_COLOUR
            End If
            
            ' RS 24/01/2003 Update Array & StatusIcons
            Select Case lCol
                Case mnTRIALSITEPERSON_COL:   mvData(dbcSubjectLockStatus, lRow - 1) = LockStatus.lsLocked
                Case mnVISIT_COL:             mvData(dbcVisitLockStatus, lRow - 1) = LockStatus.lsLocked
                Case mnFORM_COL:              mvData(dbcEFormLockStatus, lRow - 1) = LockStatus.lsLocked
                Case mnSTATUS_COL:            mvData(dbcDataItemLockStatus, lRow - 1) = LockStatus.lsLocked
            End Select
            SetIcon lCol, lRow, flexData.CellBackColor
            
            ' If it's not already locked, add to our "locked" collection
            If mcolLocked.IsItem(lRow, lCol) Then
                ' OK, it's already there
            Else
                ' Add to collection
                Call mcolLocked.AddItem(lRow, lCol)
            End If
            ' And if it was frozen, remove from our "frozen" collection
            If mcolFrozen.IsItem(lRow, lCol) Then
                Call mcolFrozen.Remove(lRow, lCol)
            End If
            
        Case LockStatus.lsFrozen
            If flexData.CellBackColor = vbHighlight Then
                flexData.CellForeColor = vbHighlightText
                flexData.CellBackColor = mnFROZEN_COLOUR
                'flexData.CellForeColor = vbWindowBackground
            Else
                flexData.CellForeColor = mnFROZEN_COLOUR
            End If
            'Add coordinates to collection
            If mcolFrozen.IsItem(lRow, lCol) Then
                ' OK, it's already there
            Else
                ' Add it to "Frozen" collection
                Call mcolFrozen.AddItem(lRow, lCol)
            End If
            ' If it was locked, remove from our "locked" collection
            If mcolLocked.IsItem(lRow, lCol) Then
                Call mcolLocked.Remove(lRow, lCol)
            End If
    
            ' RS 24/01/2003 Update Array & StatusIcons
            Select Case lCol
                Case mnTRIALSITEPERSON_COL:   mvData(dbcSubjectLockStatus, lRow - 1) = LockStatus.lsFrozen
                Case mnVISIT_COL:             mvData(dbcVisitLockStatus, lRow - 1) = LockStatus.lsFrozen
                Case mnFORM_COL:              mvData(dbcEFormLockStatus, lRow - 1) = LockStatus.lsFrozen
                Case mnSTATUS_COL:            mvData(dbcDataItemLockStatus, lRow - 1) = LockStatus.lsFrozen
            End Select
            SetIcon lCol, lRow, flexData.CellBackColor
    
    End Select

Exit Sub
ErrHandler:
  Err.Raise Err.Number, , Err.Description & "|frmDataItemResponse.SetCellLockUnlockFreeze"

End Sub
 
'---------------------------------------------------------------------
Private Sub SetLockUnlockFreeze(ByVal nSetting As LockStatus, Optional bUnfreezing As Boolean = False)
'---------------------------------------------------------------------
' Set the Lock/Unlock/Freeze status for the current flexData cell
' AND all the cells to the right if it's a question cell
' NCJ 29/2/00 - Use new clsColCoords
' NCJ 8/5/00 - Use new SetCellLockUnlockFreeze routine
' NCJ 23 Dec 02 - If bUnfreezing = TRUE, allow changes to Frozen cells
'---------------------------------------------------------------------
Dim lColClicked As Long
Dim lRowClicked As Long
Dim lColCount As Long
Dim lRowCount As Long

    On Error GoTo ErrHandler
    
    lColClicked = flexData.Col
    If flexData.Col >= mnDATAITEM_COL Then
    
        ' Do all columns to the right of the question col too
        For lColCount = mnDATAITEM_COL To flexData.Cols - 1
            flexData.Col = lColCount
            
            Call SetCellLockUnlockFreeze(nSetting, bUnfreezing)
        Next lColCount
                
    Else
        ' For Subject, Visit or Form just do single cell
        Call SetCellLockUnlockFreeze(nSetting, bUnfreezing)
    End If
    flexData.Col = lColClicked

Exit Sub
ErrHandler:
  Err.Raise Err.Number, , Err.Description & "|frmDataItemResponse.SetLockUnlockFreeze"

End Sub

'----------------------------------------------------------------------------------------
Private Function IsParentLocked(ByVal lRow As Long, ByVal lCol As Long) As Boolean
'----------------------------------------------------------------------------------------
' Return TRUE if "parent" object is locked (cell immediately to left)
' i.e. if question, test if form is locked;
' if form, test if visit is locked;
' if visit, test if trial subject is locked.
'----------------------------------------------------------------------------------------

    IsParentLocked = False
    If lCol > mnTRIALSITEPERSON_COL Then
        If mcolLocked.IsItem(lRow, lCol - 1) Then
            IsParentLocked = True
        End If
    End If
    
End Function




'---------------------------------------------------------------------
Private Function SplitTextLine(ByRef sTextLine As String, _
                                    nWidth As Integer, _
                                    nPrintFrom As Integer) As String
'---------------------------------------------------------------------
'Changed Mo 24/5/2002
'   This function now checks for the situation where sTextLine contains
'   no spaces and instead splits the string at the approriate character position
'   instaed of a space.
'   This function is virtually identical to frmViewDiscrepancies.SplitMessageLine
'---------------------------------------------------------------------
Dim i As Integer
Dim sPart As String
Dim sPreviousPart
Dim sChar As String
Dim q As Integer

    On Error GoTo ErrHandler

    'to handle the manner in which this function works a space is added to the Textline
    'unless there is one already
    If Mid(sTextLine, Len(sTextLine), 1) <> " " Then
        sTextLine = sTextLine & " "
    End If
    
    sPart = ""
    sPreviousPart = ""
    For i = 1 To Len(sTextLine)
        sChar = Mid(sTextLine, i, 1)
        sPart = sPart & sChar
        If sChar = " " Then
            If Printer.TextWidth(sPart) > nWidth Then
                Printer.CurrentX = nPrintFrom
                'Check for the situation where no spaces have been reached and the textwidth is beyond nWidth.
                'In this situation sPreviousPart would be empty and would need to have a truncated section
                'of sTextLine placed in it.
                If sPreviousPart = "" Then
                    q = Len(sTextLine)
                    Do
                        q = q - 1
                        sPreviousPart = Mid(sTextLine, 1, q)
                    Loop Until Printer.TextWidth(sPreviousPart) < nWidth
                End If
                Printer.Print sPreviousPart
                sTextLine = Mid(sTextLine, Len(sPreviousPart) + 1)
                Exit For
            Else
                sPreviousPart = sPart
            End If
        End If
    Next
    
    'Strip off the Space that was added to the end
    sTextLine = Mid(sTextLine, 1, Len(sTextLine) - 1)
    
    'check to see whether the remaining part of TextLine requires a recursive call to SplitTextLine
    If Printer.TextWidth(sTextLine) < nWidth Then
        SplitTextLine = sTextLine
    Else
        SplitTextLine = SplitTextLine(sTextLine, nWidth, nPrintFrom)
    End If
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDataItemResponse.SplitTextLine"
   
End Function

'---------------------------------------------------------------------
Private Sub PopulateGridSection(lGridStartRow As Long)
'---------------------------------------------------------------------
'TA 28/02/2001: Fill in grid from the row "lGridStartRow"
'REVISIONS
'ic 14/02/2003 bug 807 display label or id in brackets
'                      remove text on statuses
' NCJ 20 Mar 03 - Note that mvData may be Null if there are no matching records
' MLM 06/06/03: Make row taller if a subject, visit or eForm has only one response.
' ic 28/07/2005 added clinical coding
'---------------------------------------------------------------------

Dim sDateFormat As String
Dim sNRStatus As String
Dim lRow As Long
Dim lGridRow As Long
Dim bNoMoreRows As Boolean
Dim sLastPatient As String
Dim sDataItemName As String
Dim oTimezone As TimeZone           ' Used to convert timestamp to local time
Dim lTextHeight As Long             ' Used to adjust RowHeight for single-row-subjects
Dim lMaxRow As Long
'MLM 05/06/02:
Dim sLastEForm As String
Dim bSingleResponse As Boolean
Dim oDictionary As MACROCCBS30.Dictionary
Dim sCodingDetails As String
Dim sCodingStatus As String
Dim sError As String


    On Error GoTo ErrHandler
    
    ' NCJ 20 Mar 03 - Reset timeout timer
    Call RestartSystemIdleTimer
        
    ' NCJ 20 Mar 03 - mvData may be Null
    If Not IsNull(mvData) Then
        'TA 27/05/2003: check we haven't already displayed everything for smoother scrolling
        If AreAllRowsShown Then
            'we have already hsown all rows - let's exit
    'EXIT SUB
            Exit Sub
        End If
        lMaxRow = UBound(mvData, 2)
    Else
        lMaxRow = -1
    End If
    
    ' RS 13/02/2003 Reset click values for new grid
    mlRowClickedMin = 0
    mlRowClickedMax = 0
    
    picStatusIcon.BackColor = vbWindowBackground
    Set oTimezone = New TimeZone
    
    ' TA 18/10/2001 - In Forms view, show date only
    If mDRType = eDataBrowserType.dbteForms Then
        sDateFormat = "yyyy/mm/dd"
    Else
        sDateFormat = "yyyy/mm/dd hh:mm:ss"
    End If
    
    'this becomes true when no more rows are to be displayed
    bNoMoreRows = False
    
    If gbClinicalCoding Then
        If (mDRType = eDataBrowserType.dbtDataItemResponse) Then
            If (moDictionaries Is Nothing) Then
                'get the dictionary for this question
                Set moDictionaries = New MACROCCBS30.Dictionaries
                Call moDictionaries.Init(SecurityDatabasePath)
            End If
        End If
    End If
    
    lGridRow = lGridStartRow
    bSingleResponse = True
    HourglassOn
    Do While Not bNoMoreRows
        lRow = lGridRow - 1
        If lMaxRow = -1 Or (lRow > lMaxRow) Then
            ' That's it - finish off here
            If Not flexData.Redraw Then
                flexData.Redraw = True
            End If
            
            ' RS 12/03/2003: Adjust column widths to widest text
            With flexData
                .ColWidth(mnVISIT_COL) = GetMaxColumnTextWidth(mnVISIT_COL) * 1.3
                .ColWidth(mnFORM_COL) = GetMaxColumnTextWidth(mnFORM_COL) * 1.2
                .ColWidth(mnDATAITEM_COL) = GetMaxColumnTextWidth(mnDATAITEM_COL) * 1.2
                .ColWidth(mnDATARESPONSE_COL) = GetMaxColumnTextWidth(mnDATARESPONSE_COL) * 1.2
                'MLM 05/06/03: Also size status column to allow room for NR/CTC.
                .ColWidth(mnSTATUS_COL) = GetMaxColumnTextWidth(mnSTATUS_COL) * 1.2
            End With
            
            If bSingleResponse Then
                AdjustLastRowHeight lRow
            End If
            
            'MLM 06/06/03: Don't do it like this. It only accounts for subjects with 1 response; need to consider visits and eforms too.
            ' RS 12/03.2003: If there is only a single dataitem for a subject, make the rowheight at least 2 lines
            ' to make sure that the subject label is displayed.
'            With flexData
'                For lRow = 1 To flexData.Rows - 1
'                    If lRow > 0 And (.TextMatrix(lRow, 0) = .TextMatrix(lRow - 1, 0)) Then
'                        ' Previous line same subject: no action
'                    Else
'                        ' Previous line is different subject
'                        If lRow < .Rows - 1 Then
'                            If (.TextMatrix(lRow, 0) = .TextMatrix(lRow + 1, 0)) Then
'                                ' Next line is same subject: MULTI ROW SUBJECT
'                            Else
'                                ' Previous row different subject, Next row also different: SINGLE ROW SUBJECT
'                                lTextHeight = TextHeight(.TextMatrix(lRow, 0)) * 1.2
'                                If .RowHeight(lRow) < lTextHeight Then .RowHeight(lRow) = lTextHeight
'                            End If
'                        Else
'                            ' Previous row different subject, and this is the last row: SINGLE ROW SUBJECT
'                            lTextHeight = TextHeight(.TextMatrix(lRow, 0)) * 1.2
'                            If .RowHeight(lRow) < lTextHeight Then .RowHeight(lRow) = lTextHeight
'                        End If
'                    End If
'                Next
'            End With
            
            
'=== EXIT SUB ==='
            HourglassOff
            Exit Sub
        End If
        If mvRowShown(lRow) = 0 Then
            'if row hasn't been shown already
            'store that this row has been shown
            mvRowShown(lRow) = 1
            If flexData.Redraw Then
                'turn of writing to screen
                flexData.Redraw = False
            End If
            
            flexData.Row = lGridRow
            mlLastPopulatedRow = lGridRow

            'ic 14/02/2003 display label or id in brackets
            'Note that in Individual PERSON VIEW Column(mnSitePersonCol) is not visible, but this
            'data needs to be placed here to make MergeCells work correctly
            flexData.Col = mnTRIALSITEPERSON_COL

            flexData.Text = String(4, vbCrLf) & mvData(dbcStudyName, lRow) & "/" & _
                            mvData(dbcSite, lRow) & "/" & vbCrLf & _
                            RtnSubjectText(mvData(dbcSubjectId, lRow), mvData(dbcSubjectLabel, lRow))
    
            ' RS 22/01/2003: Display Combined status icon
            SetIcon flexData.Col, flexData.Row, vbWindowBackground, True
            'Call SetStatusImage(mvData(dbcSubjectStatus, lRow))
            
            
            Call SetLockUnlockFreeze(Val(RemoveNull(mvData(dbcSubjectLockStatus, lRow))))
            
            flexData.Col = mnVISIT_COL
            flexData.Text = String(3, vbCrLf) & mvData(dbcVisitName, lRow) & "[" & mvData(dbcVisitCycleNumber, lRow) & "]"
    
            ' RS 22/01/2003: Display Combined status icon
            SetIcon flexData.Col, flexData.Row, vbWindowBackground, True
            'Call SetStatusImage(mvData(dbcVisitStatus, lRow))
            Call SetLockUnlockFreeze(Val(RemoveNull(mvData(dbcVisitLockStatus, lRow))))
            
            flexData.Col = mnFORM_COL
            flexData.Text = String(3, vbCrLf) & eFormTitleLabel(mvData(dbcEFormTitle, lRow), RemoveNull(mvData(dbcEFormLabel, lRow)), mvData(MACRODBBS30.dbcEFormCycleNumber, lRow))
            
            ' RS 22/01/2003: Display Combined status icon
            SetIcon flexData.Col, flexData.Row, vbWindowBackground, True
            'Call SetStatusImage(mvData(dbcEFormStatus, lRow))
            Call SetLockUnlockFreeze(Val(RemoveNull(mvData(dbcEFormLockStatus, lRow))))
            
            If mDRType <> dbteForms Then
                flexData.Col = mnDATAITEM_COL
                sDataItemName = mvData(dbcDataItemName, lRow)
                'add cycle number it has a ownerqgroupid
                If mvData(dbcOwnerQGroupId, lRow) <> 0 Then
                    sDataItemName = sDataItemName & "[" & mvData(dbcResponseCycleNumber, lRow) & "]"
                End If
                flexData.Text = sDataItemName
                If lGridStartRow = 1 Then
                    AdjustColumnWidth (flexData.Col)
                End If
                
                flexData.Col = mnDATARESPONSE_COL
                If IsNull(mvData(dbcResponseValue, lRow)) Then
                    flexData.Text = ""
                Else
                    If (mvData(dbcDataType, lRow)) = DataType.Multimedia Then
                        flexData.Text = "Attached"
                    Else
                        ' Localise value before display - NCJ 17/11/99
                        flexData.Text = LocaliseStandardValue((mvData(dbcResponseValue, lRow)), (mvData(dbcDataType, lRow)))
                    End If
                End If
                If lGridStartRow = 1 Then
                    AdjustColumnWidth (flexData.Col)
                End If
                
                flexData.Col = mnSTATUS_COL
                'ic 14/02/2003 bug 807 remove text on statuses
                'flexData.Text = Trim(GetStatusText((mvData(dbcResponseStatus, lRow))))
    
                ' NCJ 6/10/00 - Include LabResult and CTCGrade
                If mvData(dbcDataType, lRow) = DataType.LabTest Then
                    sNRStatus = GetNRCTCText(mvData(dbcLabResult, lRow), mvData(dbcCTCGrade, lRow))
                    If sNRStatus > "" Then
                        flexData.Text = flexData.Text & "          [" & sNRStatus & "]"
                    End If
                End If
    
                ' RS 22/01/2003: Display Combined status icon
                 SetIcon flexData.Col, flexData.Row, vbWindowBackground, True
                'Call SetStatusImage(mvData(dbcResponseStatus, lRow))
                
                Call SetLockUnlockFreeze(Val(RemoveNull(mvData(dbcDataItemLockStatus, lRow))))
            
                flexData.Col = mnNEW_COL
                flexData.Text = GetTransferStatusText((mvData(dbcChanged, lRow)))
                If lGridStartRow = 1 Then
                    AdjustColumnWidth (flexData.Col)
                End If
            
            End If
            
            'Ensure timestamp is a date - NCJ 30/11/99
            flexData.Col = mnTIMESTAMP_COL
            
            'TA 3/7/02: 'isnull check removed - ResponseTimeStamp can never be null
            'If Not IsNull(mvData(dbcResponseTimeStamp, lRow)) Then
                If Val(mvData(dbcResponseTimeStamp, lRow)) > 0 Then
                    
                    ' RS 08/10/2002 Add Timezone Information, or convert timestamp to local format
                    If GetMACROSetting("timestampdisplay", "storedvalue") = "storedvalue" Then
                        ' Display the stored value, add offset to GMT in brackets
                        ' Original Format/Value
                        'flexData.Text = Format(CDate(mvData(dbcResponseTimeStamp, lRow)), sDateFormat) & _
                                            " (GMT" & IIf(mvData(dbcResponseTimestamp_TZ, lRow) < 0, "+", "") & -mvData(dbcResponseTimestamp_TZ, lRow) \ 60 & ":" & Format(Abs(mvData(dbcResponseTimestamp_TZ, lRow)) Mod 60, "00") & ")"
                        'TA 22/05/2003: use function
                        flexData.Text = DisplayGMTTime(mvData(dbcResponseTimeStamp, lRow), "yyyy/mm/dd hh:mm:ss", mvData(dbcResponseTimestamp_TZ, lRow))
                    Else
                        ' Convert the stored value to local time
                        flexData.Text = Format(oTimezone.ConvertDateTimeToLocal(mvData(dbcResponseTimeStamp, lRow), mvData(dbcResponseTimestamp_TZ, lRow)), sDateFormat)
                    End If
                    
                End If
                
                
            If lGridStartRow = 1 Then
                AdjustColumnWidth (flexData.Col)
            End If
                
              flexData.Col = mnDB_TIMESTAMP_COL
              
                'TA do database time stamp
                If Val(mvData(dbcDatabaseTimeStamp, lRow)) > 0 Then
                    
                    ' RS 08/10/2002 Add Timezone Information, or convert timestamp to local format
                    If GetMACROSetting("timestampdisplay", "storedvalue") = "storedvalue" Then
                        ' Display the stored value, add offset to GMT in brackets
                        ' Original Format/Value
                        'flexData.Text = Format(CDate(mvData(dbcDatabaseTimeStamp, lRow)), sDateFormat) & _
                                            " (GMT" & IIf(mvData(dbcDatabaseTimestamp_TZ, lRow) < 0, "+", "") & -mvData(dbcDatabaseTimestamp_TZ, lRow) \ 60 & ":" & Format(Abs(mvData(dbcDatabaseTimestamp_TZ, lRow)) Mod 60, "00") & ")"
                        'TA 22/05/2003: use function
                        flexData.Text = DisplayGMTTime(mvData(dbcDatabaseTimeStamp, lRow), sDateFormat, mvData(dbcDatabaseTimestamp_TZ, lRow))
                        
                    
                    Else
                        ' Convert the stored value to local time
                        flexData.Text = Format(oTimezone.ConvertDateTimeToLocal(mvData(dbcDatabaseTimeStamp, lRow), mvData(dbcDatabaseTimestamp_TZ, lRow)), sDateFormat)
                    End If
                    
                End If
                
                
            'End If
           ' If lGridStartRow = 1 Then
                AdjustColumnWidth (flexData.Col)
            'End If
            
            If mDRType <> dbteForms Then
                flexData.Col = mnUSERID_COL
                flexData.Text = mvData(dbcUserName, lRow)
                If lGridStartRow = 1 Then
                    AdjustColumnWidth (flexData.Col)
                End If
                
                flexData.Col = mnUSERNAMEFULL_COL
                flexData.Text = ConvertFromNull(mvData(dbcFullUserName, lRow), vbString)
                If lGridStartRow = 1 Then
                    AdjustColumnWidth (flexData.Col)
                End If

                'TA 02/06/2003: only show comments if they have permission
                If goUser.CheckPermission(gsFnViewIComments) Then
                    flexData.Col = mnCOMMENT_COL
                    flexData.Text = RemoveNull(mvData(dbcComments, lRow))
                    If lGridStartRow = 1 Then
                        AdjustColumnWidth (flexData.Col)
                    End If
                End If

                
                'SDM SR2408
                flexData.Col = mnRFC_COL
                flexData.Text = RemoveNull(mvData(dbcReasonForChange, lRow))
                If lGridStartRow = 1 Then
                    AdjustColumnWidth (flexData.Col)
                End If

                
                ' NCJ 19/5/00 SR3454
                flexData.Col = mnOVERRULE_COL
                flexData.Text = RemoveNull(mvData(dbcOverruleReason, lRow))
                If lGridStartRow = 1 Then
                    AdjustColumnWidth (flexData.Col)
                End If

               
            ' NCJ 17/10/00 SR3864 Show Validation message
            ' For Inform status, only show if user has correct rights
            flexData.Col = mnVAL_MESSAGE_COL
            If mvData(dbcResponseStatus, lRow) = Status.Inform Then
                If goUser.CheckPermission(gsFnMonitorDataReviewData) Then
                    flexData.Text = RemoveNull(mvData(dbcValMessage, lRow))
                End If
            Else
                flexData.Text = RemoveNull(mvData(dbcValMessage, lRow))
            End If
                If lGridStartRow = 1 Then
                    AdjustColumnWidth (flexData.Col)
                End If

            'ic 28/07/2005 added clinical coding
            If gbClinicalCoding Then
                If (mDRType = eDataBrowserType.dbtDataItemResponse) Then
                    flexData.Col = mnDICTIONARY_COL
                    flexData.Text = RemoveNull(mvData(dbcDictionaryName, lRow)) & " " & RemoveNull(mvData(dbcDictionaryVersion, lRow))
                    If lGridStartRow = 1 Then
                        AdjustColumnWidth (flexData.Col)
                    End If
                    
                    If (RemoveNull(mvData(dbcCodingStatus, lRow)) > "") Then
                        sCodingStatus = GetCodingStatusString(RemoveNull(mvData(dbcCodingStatus, lRow)))
                    Else
                        sCodingStatus = ""
                    End If
                    flexData.Col = mnCODINGSTATUS_COL
                    flexData.Text = sCodingStatus
                    If lGridStartRow = 1 Then
                        AdjustColumnWidth (flexData.Col)
                    End If
                    
                    If RemoveNull(mvData(dbcCodingDetails, lRow)) > "" Then
                        Set oDictionary = moDictionaries.DictionaryFromVersion(RemoveNull(mvData(dbcDictionaryName, lRow)), RemoveNull(mvData(dbcDictionaryVersion, lRow)))
                        If Not (oDictionary Is Nothing) Then
                            If Not oDictionary.ToText(RemoveNull(mvData(dbcCodingDetails, lRow)), sCodingDetails, sError) Then
                                sCodingDetails = ""
                                Call DialogError("MACRO plugin '" & oDictionary.Id & "' encountered errors :" & vbCrLf & sError)
                            End If
                        Else
                            sCodingDetails = "Dictionary not found"
                        End If
                    Else
                        sCodingDetails = ""
                    End If
                    flexData.Col = mnCODINGDETAILS_COL
                    flexData.Text = sCodingDetails
                    If lGridStartRow = 1 Then
                        AdjustColumnWidth (flexData.Col)
                    End If
                End If
            End If

            End If
        
        
            ' NCJ 3/3/00 Adjust height for Comments and RFC and Overrule
            AdjustRowHeight (lGridRow)
        End If
        If lGridRow >= lGridStartRow + m_ROW_BUFFER Then
            'over a m_ROW_BUFFER then stop cycling once our patient is differenct from the last
            If mvData(dbcStudyName, lRow) & "|" _
                    & mvData(dbcSite, lRow) & "|" _
                    & mvData(dbcSubjectId, lRow) <> sLastPatient Then
                        bNoMoreRows = True
            End If
        End If
        
        'store patient key
        sLastPatient = mvData(dbcStudyName, lRow) & "|" _
            & mvData(dbcSite, lRow) & "|" _
            & mvData(dbcSubjectId, lRow)
            
        'MLM 05/06/03:
        If lGridRow > lGridStartRow Then
            If sLastEForm <> sLastPatient & "|" & mvData(dbcVisitId, lRow) & _
                "|" & mvData(dbcVisitCycleNumber, lRow) & "|" & mvData(dbcEFormTaskID, lRow) Then
                If bSingleResponse Then
                    AdjustLastRowHeight lRow
                End If
                bSingleResponse = True
            Else
                bSingleResponse = False
            End If
        End If
        sLastEForm = sLastPatient & "|" & mvData(dbcVisitId, lRow) & _
            "|" & mvData(dbcVisitCycleNumber, lRow) & "|" & mvData(dbcEFormTaskID, lRow)
        
        lGridRow = lGridRow + 1
    
    Loop
    HourglassOff

    
    
    
    If Not flexData.Redraw Then
        flexData.Redraw = True
    End If

    Set oTimezone = Nothing
    Set oDictionary = Nothing

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDataItemResponse.PopulateGridSection"

End Sub

'---------------------------------------------------------------------
Private Function AreAllRowsShown() As Boolean
'---------------------------------------------------------------------
'have all rows been displayed
'---------------------------------------------------------------------
Dim i As Long

    For i = 0 To UBound(mvRowShown)
        If mvRowShown(i) = 0 Then
            'this row hasn't been shown therefore not all rows shown
            AreAllRowsShown = False
            Exit Function
        End If
    Next
    
    AreAllRowsShown = True

End Function

'---------------------------------------------------------------------
Private Sub flexData_Scroll()
'---------------------------------------------------------------------
'TA 28/02/2001: catch when grid is scrolled
'---------------------------------------------------------------------
Static bWorking As Boolean

    'TA 27/05/2003: check we are already processing event to stop jerky behaviour

    If Not bWorking Then
        'not already trying something
        bWorking = True
        'stop other events being sent
        'fill in visible grid
        Call PopulateGridSection(flexData.TopRow)
        bWorking = False
    End If

End Sub

'---------------------------------------------------------------------
' NCJ 23 Dec 02 - Removed unused IsValidString function
'---------------------------------------------------------------------

'---------------------------------------------------------------------
Private Function CanDoLFAOnServer(sStudyName As String, sSite As String, lSubjectId As Long) As Boolean
'---------------------------------------------------------------------
' Returns TRUE if Lock/Freeze operations can be done on this subject
' i.e. if we're a Server, there must be no unprocessed LFMessages and no data awaiting import
'---------------------------------------------------------------------
Dim oLF As LockFreeze

    CanDoLFAOnServer = True
    
    ' We don't need to do checks if we're at the site
    If gblnRemoteSite Then Exit Function

    ' We're on the Server
    Set oLF = New LockFreeze
    CanDoLFAOnServer = oLF.CanLockFreezeOnServer(MacroADODBConnection, _
                                    sStudyName, sSite, lSubjectId)
    Set oLF = Nothing
    
End Function

'---------------------------------------------------------------------
Private Sub lcmdPrint_Click()
'---------------------------------------------------------------------
'print
'---------------------------------------------------------------------

    If mDRType = dbteForms Then
        PrintForms
    Else
        PrintData
    End If
    
End Sub


''-----------------------------------------------------------------------------
'' GenerateStatusIcon
'' RS 22/01/2002
''
'' Combines the different status values to generate a single icon showing each
'' status information. If the combined Icon does not exist, it is added to the
'' imagelist.
''-----------------------------------------------------------------------------
'Private Sub GenerateStatusIcon(lBasicStatus, lCommentStatus, lChangeCount, lNoteStatus, lDiscrepancyStatus, lSDVStatus, lLockStatus)
'Dim picCombinedIcon As StdPicture
'Dim lKeyBasic As Long
'Dim lKeyComment As Long
'Dim lKeyChangeCount As Long
'Dim lKeyNote As Long
'Dim lKeyDisc As Long
'Dim lKeySDV As Long
'Dim lCurrentKey As Long
'Static Keys As String
'
'    Set picCombinedIcon = MakeIcon(lBasicStatus, lCommentStatus, lChangeCount, lNoteStatus, lDiscrepancyStatus, lSDVStatus, lLockStatus)
'
'    ' Set the icon in the grid
'    flexData.CellPictureAlignment = flexAlignLeftCenter
'    Set flexData.CellPicture = picCombinedIcon
'    flexData.Text = Space(flexData.CellPicture.Width / 80) & flexData.Text
'    Call AdjustColumnWidth(flexData.Col)
'    Exit Sub
'
''    On Error GoTo ErrHandler
''
''    ' Determine which basic Icon to use
''    Select Case lBasicStatus
''        Case Status.Success
''            lKeyBasic = 3
''        Case Status.Missing
''            lKeyBasic = 4
''        Case Status.Warning
''            lKeyBasic = 7
''        Case Status.OKWarning
''            lKeyBasic = 6
''        Case Status.Inform
''            If goUser.CheckPermission(gsFnMonitorDataReviewData) Then
''                lKeyBasic = 5
''            Else
''                lKeyBasic = 3
''            End If
''        Case Status.NotApplicable
''            lKeyBasic = 1
''        Case Status.Unobtainable
''            lKeyBasic = 2
''        Case Else
''            lKeyBasic = 0
''    End Select
''
''    ' Override Icon with Discrepancy / SDV Status
''    If lDiscrepancyStatus = eDiscrepancyStatus.dsRaised Then lKeyBasic = 12
''    If lDiscrepancyStatus = eDiscrepancyStatus.dsResponded Then lKeyBasic = 13
''    If lSDVStatus = eSDVStatus.ssQueried Then lKeyBasic = 11
''    If lLockStatus = LockStatus.lsLocked Then lKeyBasic = 9
''    If lLockStatus = LockStatus.lsFrozen Then lKeyBasic = 10
''
''
''    lCurrentKey = lKeyBasic
''
''
''    ' If a comment, combine with comment
''    If lCommentStatus <> 0 Then
''        lKeyComment = 128
''
''        ' Only generate new image if not yet available
''        AddCombinedIcon lCurrentKey, lKeyComment
''        lCurrentKey = lCurrentKey + lKeyComment
''
''    End If
''
''
''    ' ---- ChangeCount ----
''
''    If lChangeCount > 1 Then
''        If lChangeCount = 2 Then
''            lKeyChangeCount = 16
''        Else
''            If lChangeCount = 3 Then
''                lKeyChangeCount = 32
''            Else
''                lKeyChangeCount = 48
''            End If
''        End If
''
''        AddCombinedIcon lCurrentKey, lKeyChangeCount
''        lCurrentKey = lCurrentKey + lKeyChangeCount
''    End If
''
''
''    ' NoteStatus
''    If lNoteStatus <> 0 Then
''        lKeyNote = 64
''        AddCombinedIcon lCurrentKey, lKeyNote
''        lCurrentKey = lCurrentKey + lKeyNote
''    End If
''
''
''    ' SDV Status
''    If lSDVStatus = eSDVStatus.ssPlanned Then
''        lKeySDV = 256
''        AddCombinedIcon lCurrentKey, lKeySDV
''        lCurrentKey = lCurrentKey + lKeySDV
''    End If
''
''
''    Set picCombinedIcon = imgStatusIcons.ListImages("K" & lCurrentKey).Picture
''
''    ' Set the icon in the grid
''    flexData.CellPictureAlignment = flexAlignLeftCenter
''    Set flexData.CellPicture = picCombinedIcon
''    flexData.Text = Space(flexData.CellPicture.Width / 60) & flexData.Text
''    Call AdjustColumnWidth(flexData.Col)
''
''Exit Sub
''ErrHandler:
''    Err.Raise Err.Number, , Err.Description & "|frmDataItemResponse.GenerateStatusIcon"
'
'End Sub
''-----------------------------------------------------------------------------
'' AddCombinedIcon
'' IN: the keys of the two existing icons to combine
'' It is assumed that the two icons already exist
'' The combined icon is added to the imagelist with the combined key
''-----------------------------------------------------------------------------
'Private Sub AddCombinedIcon(Key1 As Long, Key2 As Long)
'Dim picCombinedIcon As StdPicture
'Dim sKey As String
'
'
'        ' Try to get the Key of the Icon that should be created. If this
'        ' succeeds, then the Icon already exists, no further action required
'        On Error GoTo IconMissing:
'        sKey = imgStatusIcons.ListImages.Item("K" & (Key1 + Key2)).Key
'        Exit Sub
'
'IconMissing:
'        ' The Combined Icon does not exist. Combine the images and add to imagelist
'        Set picCombinedIcon = imgStatusIcons.Overlay("K" & Key1, "K" & Key2)
'        On Error Resume Next
'        imgStatusIcons.ListImages.Add , "K" & (Key1 + Key2), picCombinedIcon
'        'Debug.Print "Added " & "K" & (Key1 + Key2) & " Height = " & imgStatusIcons.ListImages.Item("K" & (Key1 + Key2)).Picture.Height
'
'
'End Sub

'-----------------------------------------------------------------------------
Private Function MakeIcon(lBasicStatus, lCommentStatus, lChangeCount, lNoteStatus, lDiscrepancyStatus, lSDVStatus, lLockStatus) As StdPicture
'-----------------------------------------------------------------------------
' Make the required Icon
'-----------------------------------------------------------------------------
' REVISIONS
' DPH 23/09/2003 - Keep Collection of Icons and check if already made/stored
'-----------------------------------------------------------------------------
Dim lKeyBasic As Long
Dim lKeyComment As Long
Dim lKeyChangeCount As Long
Dim lKeyNote As Long
Dim lKeyDisc As Long
Dim lKeySDV As Long
Dim lCurrentKey As Long
Dim sCollectionKey As String

    ' DPH 23/09/2003 - Reset variables
    lKeyBasic = 0
    lKeyComment = 0
    lKeyChangeCount = 0
    lKeyNote = 0
    lKeyDisc = 0
    lKeySDV = 0
    lCurrentKey = 0

    ' Determine which basic Icon to use
    Select Case lBasicStatus
        Case Status.Success
            lKeyBasic = 3
        Case Status.Missing
            lKeyBasic = 4
        Case Status.Warning
            lKeyBasic = 7
        Case Status.OKWarning
            lKeyBasic = 6
        Case Status.Inform
            If goUser.CheckPermission(gsFnMonitorDataReviewData) Then
                lKeyBasic = 5
            Else
                lKeyBasic = 3
            End If
        Case Status.NotApplicable
            lKeyBasic = 1
        Case Status.Unobtainable
            lKeyBasic = 2
        Case Else
            lKeyBasic = 0
    End Select
    
    ' Override Icon with Discrepancy / SDV Status
    If lDiscrepancyStatus = eDiscrepancyStatus.dsRaised Then lKeyBasic = 12
    If lDiscrepancyStatus = eDiscrepancyStatus.dsResponded Then lKeyBasic = 13
    'If lSDVStatus = eSDVStatus.ssQueried Then lKeyBasic = 11
    If lLockStatus = LockStatus.lsLocked Then lKeyBasic = 9
    If lLockStatus = LockStatus.lsFrozen Then lKeyBasic = 10

  
    ' If a comment, combine with comment
    If lCommentStatus <> 0 Then
        lKeyComment = 128
    End If
    
    ' ---- ChangeCount ----
    If lChangeCount > 1 Then
        If lChangeCount = 2 Then
            lKeyChangeCount = 16
        Else
            If lChangeCount = 3 Then
                lKeyChangeCount = 32
            Else
                lKeyChangeCount = 48
            End If
        End If
    End If
        
    
    ' NoteStatus
    If lNoteStatus <> 0 Then
        lKeyNote = 64
    End If
    
    
    ' SDV Status
    If lSDVStatus = eSDVStatus.ssPlanned Then
        lKeySDV = 256
    End If
    If lSDVStatus = eSDVStatus.ssQueried Then
        lKeySDV = 512
    End If
    If lSDVStatus = eSDVStatus.ssComplete Then
        lKeySDV = 768
    End If
    


    Me.Visible = True
    picStatusIcon.Picture = Nothing
    
    picStatusIcon.AutoRedraw = True
    picStatusIcon.Height = 390
    'picStatusIcon.BackColor = flexData.BackColorSel
    
    ' DPH 23/09/2003 - added collection checking section for existing pictures
    sCollectionKey = "K" & lKeySDV & "-" & lKeyBasic & "-" & lKeyComment & "-" & lKeyChangeCount & "-" & lKeyNote & "-" & Abs(picStatusIcon.BackColor)
    ' Check if picture exists in collection else create
    If Not CollectionMember(mcolDrawnImages, sCollectionKey, True) Then
        ' create required image
        If lKeySDV > 0 Then imgStatusIcons.ListImages("K" & lKeySDV).Draw picStatusIcon.hDc, 0, 0, imlTransparent
        picStatusIcon.Picture = picStatusIcon.Image
        If lKeyBasic > 0 Then imgStatusIcons.ListImages("K" & lKeyBasic).Draw picStatusIcon.hDc, 0, 0, imlTransparent
        picStatusIcon.Picture = picStatusIcon.Image
        If lKeyComment > 0 Then imgStatusIcons.ListImages("K" & lKeyComment).Draw picStatusIcon.hDc, 0, 0, imlTransparent
        picStatusIcon.Picture = picStatusIcon.Image
        If lKeyChangeCount > 0 Then imgStatusIcons.ListImages("K" & lKeyChangeCount).Draw picStatusIcon.hDc, 0, 0, imlTransparent
        picStatusIcon.Picture = picStatusIcon.Image
        If lKeyNote > 0 Then imgStatusIcons.ListImages("K" & lKeyNote).Draw picStatusIcon.hDc, 0, 0, imlTransparent
        picStatusIcon.Picture = picStatusIcon.Image
        ' now add to collection for possible use later
        Call CollectionAddAnyway(mcolDrawnImages, picStatusIcon.Image, sCollectionKey)
    Else
        ' get image from collection
        picStatusIcon.Picture = mcolDrawnImages(sCollectionKey)
    End If
    
    ' Now the picture in picStatusIcon is correct
    Set MakeIcon = picStatusIcon.Picture


    ' Set the icon in the grid
    'flexData.CellPictureAlignment = flexAlignLeftCenter
    'Set flexData.CellPicture = picStatusIcon.Picture
    'flexData.Text = Space(flexData.CellPicture.Width / 60) & flexData.Text
    'Call AdjustColumnWidth(flexData.Col)


End Function


'Private Sub DumpIcons()
'Dim a As Variant
'Dim b As Variant
'Dim c As Variant
'Dim d As Variant
'Dim e As Variant
'
'Dim cntA As Integer
'Dim cntB As Integer
'Dim cntC As Integer
'Dim cntD As Integer
'Dim cntE As Integer
'
'    a = Array(1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 13)
'    b = Array(0, 16, 32, 48)
'    c = Array(0, 64)
'    d = Array(0, 128)
'    e = Array(0, 256)
'
'    For cntA = LBound(a) To UBound(a)
'        For cntB = LBound(b) To UBound(b)
'            For cntC = LBound(c) To UBound(c)
'                For cntD = LBound(d) To UBound(d)
'                    For cntE = LBound(e) To UBound(e)
'                        SavePicture MakeIcon(a(cntA), b(cntB), c(cntC), d(cntD), e(cntE), "C:\Temp\NewIcons\Normal - K" & (a(cntA) + b(cntB) + c(cntC) + d(cntD) + e(cntE))
'                        'Debug.Print a(cntA), b(cntB), c(cntC), d(cntD), e(cntE)
'                    Next
'                Next
'            Next
'        Next
'    Next
'
'End Sub


Private Function SetIcon(lColumn As Long, lRow As Long, Optional lBackground As Long = -1, Optional bPopulateGrid As Boolean = False) As StdPicture
Dim lCurrentCol As Long         ' Save current Col
Dim lCurrentRow As Long         ' Save current Row


' Only update icon if requested column actually contains an Icon
If lColumn = mnTRIALSITEPERSON_COL Or lColumn = mnVISIT_COL Or lColumn = mnFORM_COL Or lColumn = mnSTATUS_COL Then

    ' Save current Row/Col setting
    lCurrentCol = flexData.Col
    lCurrentRow = flexData.Row

    ' Position on Row/Column to update
    flexData.Col = lColumn
    flexData.Row = lRow

    ' Set specific Background colour if requested, otherwise use what is already there
    If lBackground <> -1 And lBackground <> 0 Then picStatusIcon.BackColor = lBackground

    ' Make the icon based on the different parts of the status
    Select Case lColumn
        Case mnTRIALSITEPERSON_COL
            Set flexData.CellPicture = MakeIcon(mvData(dbcsubjectStatus, lRow - 1), 0, 0, mvData(dbcSubjectNoteStatus, lRow - 1), mvData(dbcSubjectDiscStatus, lRow - 1), mvData(dbcSubjectSDVStatus, lRow - 1), mvData(dbcSubjectLockStatus, lRow - 1))
        Case mnVISIT_COL
            Set flexData.CellPicture = MakeIcon(mvData(dbcVisitStatus, lRow - 1), 0, 0, mvData(dbcVisitNoteStatus, lRow - 1), mvData(dbcVisitDiscStatus, lRow - 1), mvData(dbcVisitSDVStatus, lRow - 1), mvData(dbcVisitLockStatus, lRow - 1))
        Case mnFORM_COL
            Set flexData.CellPicture = MakeIcon(mvData(dbcEFormStatus, lRow - 1), 0, 0, mvData(dbcEFormNoteStatus, lRow - 1), mvData(dbcEFormDiscStatus, lRow - 1), mvData(dbcEFormSDVStatus, lRow - 1), mvData(dbcEFormLockStatus, lRow - 1))
        Case mnSTATUS_COL
            Set flexData.CellPicture = MakeIcon(mvData(dbcResponseStatus, lRow - 1), mvData(dbcComments, lRow - 1) <> "", mvData(dbcChangeCount, lRow - 1), mvData(dbcDataItemNoteStatus, lRow - 1), mvData(dbcDataItemDiscStatus, lRow - 1), mvData(dbcDataItemSDVStatus, lRow - 1), mvData(dbcDataItemLockStatus, lRow - 1))
    End Select

    ' If populating the grid (first time, or Refresh), adjust column width
    If bPopulateGrid Then
      '  Select Case lCurrentCol
     '   Case 0, 1, 2
            flexData.CellPictureAlignment = flexAlignLeftCenter

            flexData.Text = flexData.Text
      '  Case Else
       '     flexData.CellPictureAlignment = flexAlignLeftCenter
        '    flexData.Text = Space(flexData.CellPicture.Width / 80) & flexData.Text
        'End Select
        Call AdjustColumnWidth(flexData.Col)
    End If

    ' Restore Row/Col setting
    flexData.Col = lCurrentCol
    flexData.Row = lCurrentRow

End If
End Function

Private Sub NewMIstatus(lStudyId, sSite, lPersonId, lVisitId, nVisitCycleNumber, lCRFPageTaskId, lResponseTaskId, nResponseCycle, enType)
Dim lRow As Long
Dim lTopRow As Long

    lTopRow = flexData.TopRow
    Call frmMenu.RefreshSearchResults
    If flexData.Rows >= lTopRow Then flexData.TopRow = lTopRow
    Exit Sub

    If lVisitId = 0 Then
        ' MIMessage for Subject
        For lRow = 1 To flexData.Rows - 1
            If mvData(dbcStudyId, lRow - 1) = lStudyId And mvData(dbcSite, lRow - 1) = sSite And mvData(dbcSubjectId, lRow - 1) = lPersonId Then
                Select Case enType
                    Case MIMsgType.mimtSDVMark: mvData(dbcSubjectSDVStatus, lRow - 1) = eSDVStatus.ssPlanned
                End Select
                SetIcon mnTRIALSITEPERSON_COL, lRow
            End If
        Next
    End If
    
    If lCRFPageTaskId = 0 Then
        ' MIMessage for Visit
        ' MIMessage for Subject
        For lRow = 1 To flexData.Rows - 1
            If mvData(dbcStudyId, lRow - 1) = lStudyId And mvData(dbcSite, lRow - 1) = sSite And mvData(dbcSubjectId, lRow - 1) = lPersonId _
                And mvData(dbcVisitId, lRow - 1) = lVisitId And mvData(dbcVisitCycleNumber, lRow - 1) = nVisitCycleNumber Then
                SetIcon mnTRIALSITEPERSON_COL, lRow
                SetIcon mnVISIT_COL, lRow
            End If
        Next
    End If

    If lResponseTaskId = 0 Then
        ' MIMessage for Form
    End If
    
    
    
    
    ' For each Row in the grid
    'Debug.Print "NewMIStatus: ", lStudyId, sSite, lPersonId, lVisitId, nVisitCycleNumber, lCRFPageTaskId, lResponseTaskId, nResponseCycle, enType
    
End Sub

'---------------------------------------------------------------------
Private Function DisplayGMTTime(ByVal dblTime As Double, ByVal sDateFormat As String, ByVal vTimezoneOffset As Variant) As String
'---------------------------------------------------------------------
'TA 22/05/2003: Function to convert a time and timezone offset into GMT
'should be static method in TimeZone class
'---------------------------------------------------------------------

    DisplayGMTTime = Format(CDate(dblTime), sDateFormat) & " " & DisplayGMTTimeZoneOffset(vTimezoneOffset)

End Function

'---------------------------------------------------------------------
Private Function DisplayGMTTimeZoneOffset(ByVal vTimezoneOffset As Variant) As String
'---------------------------------------------------------------------
'TA 22/05/2003: Function to convert a timezoneoffset as returned by TimeZone class into GMT offset
'should be static method in TimeZone class
'---------------------------------------------------------------------
Dim sText As String

    If IsNull(vTimezoneOffset) Then
        sText = ""
    Else
        sText = "(GMT"
        If vTimezoneOffset <> 0 Then
            If vTimezoneOffset < 0 Then
                sText = sText & "+"
            End If
            sText = sText & -vTimezoneOffset \ 60 & ":" & Format(Abs(vTimezoneOffset) Mod 60, "00")
        End If
                                                
        sText = sText & ")"
    End If

     DisplayGMTTimeZoneOffset = sText
 
End Function


'---------------------------------------------------------------------
Private Function eFormTitleLabel(ByVal sTitle As String, ByVal sLabel As String, ByVal sCycle As String) As String
'---------------------------------------------------------------------
'TA 27/05/2003: return label if it exists or title if not
'---------------------------------------------------------------------

    If sLabel = "" Then
        eFormTitleLabel = sTitle & "[" & sCycle & "]"
    Else
        eFormTitleLabel = sLabel
    End If

End Function




