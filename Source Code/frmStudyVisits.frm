VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStudyVisits 
   Caption         =   "Study Visits"
   ClientHeight    =   7080
   ClientLeft      =   75
   ClientTop       =   1440
   ClientWidth     =   11835
   Icon            =   "frmStudyVisits.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   11835
   Begin VB.CommandButton cmdValidationExpression 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   0
      Picture         =   "frmStudyVisits.frx":0A4A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Timer tmrGridClick 
      Enabled         =   0   'False
      Interval        =   125
      Left            =   120
      Top             =   6120
   End
   Begin VB.TextBox txtCycleVisit 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      MaxLength       =   255
      TabIndex        =   2
      Text            =   "This text box is used for repeating cycle"
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox txtVisitName 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      MaxLength       =   50
      TabIndex        =   1
      Text            =   "Used for Editing of Visit Name"
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid flxVisits 
      Height          =   6855
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12091
      _Version        =   393216
      BackColorBkg    =   16777215
   End
   Begin MSComctlLib.ImageList imgLargeIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudyVisits.frx":0D54
            Key             =   "RepeatingCRFpage"
            Object.Tag             =   "gsREPEATING_CRF_PAGE_LABEL"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudyVisits.frx":11A6
            Key             =   "Visit_eForm"
            Object.Tag             =   "gsVISIT_EFORM"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudyVisits.frx":15F8
            Key             =   "CRFpageLarge"
            Object.Tag             =   "gsLARGE_CRF_PAGE_LABEL"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmStudyVisits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998 - 2006. All Rights Reserved
'   File:       frmStudyVisits.frm
'   Author:     Andrew Newbigging June 1997
'   Purpose:    Maintenance of Schedule in Study Definition
'               (NB See file frmRunSchedule.frm for DM version of frmStudyVisits)
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   26  Joanne Lau              02/10/98 SPR 527
'                                        Changes made to the flxVisits mousedown event,
'                                        to prevent a form being inserted or deleted
'                                        from the schedule when a click is made outside
'                                        the grid.See Sub for explanation.
'   26  Andrew Newbigging       13/10/98
'       New row for visit date included (for MACRO_DM only).  Change made to:
'       flxvisits_click,RefreshRows,RefreshStudyVisits, txtCycleVisit_LostFocus
'   27  Andrew newbigging       15/10/98
'       Correction to previous change - reference to frmMenu.TrialSubject changed
'       to frmMenu.TrialSubject
'   28  Andrew Newbigging       9/11/98 SPR 578
'       RefreshStudyVisits modified to prevent user resizing of grid
'   29  Andrew Newbigging       12/11/98
'       Added clinicaltrialid parameter to DeleteProformaVisit
'   30  Andrew Newbigging       24/11/98    SR 616
'       Added validation to check that code starts with alphabetic character (required by Prolog)
'       on InsertVisit
'   31  Andrew Newbigging       24/11/98    SR 617
'       Displays visit name in message during delete
'   32  Andrew Newbigging       24/11/98    SR 580
'       Sets selected cell in form activate to be the first available form and allows
'       user to select a form for data entry by pressing Return
'   33  Andrew newbigging       4/12/1998   SR 645
'       Modified ShowForm to call frmCRFDate if no date has been entered for the form
'       Mo Morris               10/12/98    SR 556 & 658
'       Form_Unload added so that frmMenu.HideVisits is called when frmStudyVisits is closed
'       via the forms Close box.
'       Mo Morris               10/12/98    SR 630
'       flxVisits_MouseUp changed so that when a CRF Page order position is changed a search
'       is made for frmCRFDesign. If it exists then frmCRFDesign.RefreshCRF is called to change
'       the order of the CRF page tabs in the Tabstrip.
'       Mo Morris               6/1/99      SR 668
'       InsertVisit now makes an additional call to gblnNotAReservedWord
'       Mo Morris               8/1/99      SR 631
'       RefreshStudyVisits used to include a me.show call that was not required if frmStudyVisits
'       was not visible. The me.show call has now been removed. When RefreshStudyVisits is
'       called from frmMenu.ViewVisits, frmStudyVisits.InsertVisit and frmStudyVisits.Delete
'       a frmStudyVisits/me.show is neccessary so a frmStudyVisits/me.show has been placed
'       before their call to RefreshStudyVisits.
'       Andrew Newbigging       22/2/99
'       Modified ShowForm to set the CRFPageDate property of frmCRFDataEntry and to use
'       the new function FormatDate to set the date format for the appropriate database type
'       Andrew Newbigging       5/3/99
'       Modified ShowForm to call the SetFormFocus routine to set the focus to the first
'       available item.
'       Andrew Newbigging       1/5/99  SR 875,876
'       Modified Delete to ensure that txtVisitName is not visible after deleting a visit
'       Mo Morris               4/6/99  SR 947
'       This form can no longer be minimized (form attribute Controlbox set to false)
'       NCJ 13/8/99 - Arezzo updating
'       NCJ 27/8/99 - Set Arezzo cycle info for repeating forms
'   Paul Norris             30/08/99
'       Amended InsertVisit() to saving visit name when inserting new visit
'       Ameded flxVisits_Click() and flxVisits_DblClick() to enable timer so that when
'       an event is fired only one action takes place, ie one or the other
'   Paul Norris             02/09/99    SR 415, 1310, 291, 1491, 1396, 1711
'       Amended RefreshColumns(), flxVisits_Click(), flxVisits_MouseUp(), txtVisitName_LostFocus()
'       Added variaous routines to upgrade to MTM2.0
'   Paul Norris             06/09/99    SR 1396
'       Added SetBackgroundColour() to set back colour of grid
'   NCJ 17 Sept 99
'       Removed old unused code for DataManagement
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   NCJ 27/9/99     Added SetCyclingTask calls to InsertStudyVisitCRFPage
'                   Removed superfluous arguments in RefreshCells
'   PN  28/09/99    Amended ShowPopupMenu() and flxVisits_MouseUp() to setup menu options correctly
'                   Amended SetColumnWidth() to specify a min column width
'                   Amended txtCycleVisit_LostFocus() to allow and save null string
'   NCJ 9 Nov 99    Hide Visits schedule before showing CRFPage in GridDblClick (SR 2065)
'   WillC   10/11/99 Added the error handlers
'   Mo Morris   16/11/99    DAO to ADO conversion'
'   NCJ 30/11/99    VisitDate is expression not condition
'   NCJ 3 Dec 99 - Use single quotes in SQL strings
'   NCJ 7 Dec 99 - Bug fix to SetColumnWidth SR 2085
'   NCJ 10 Dec 99 - Added checks for user access rights
'   NCJ 13 Dec 99 - Ids to Long
'   NCJ 18/1/00 SR 2586 Stop user doing silly drags
'   Mo Morris   14/2/00 sections of InsertVisit code re-writen
'   WillC       28/2/2000 SR2796 Added a tool tip for row which allows editing.
'   TA 17/03/2000   SR3207    put in right click option of prompting user for a visit date
'   TA 29/04/2000   subclassing removed
' NCJ 10 Jan 02 - Make sure ALL question lists are updated in ReorderCRFPages
' ZA 23/08/2002 - Added visit date check to updateStudyVisitCRFPage function
' ZA 09/09/2002 - Allow user to create cycling visits from schedule
' ZA 11/09/2002 - call SaveVisitName & SaveVisitCycle routine when this form closes
' ASH 4/11/2002 - Changed nRepeatValue to lRepeatValue in sub SaveVisitCycle
' NCJ 31 Jan 03 - Accept Return key to enter Visit Name or Visit Cycles
' REM 17/08/04 - in routine SaveVisitCycle and SaveVisitName check to see if active control is nothing
' NCJ 2 Nov 04 - In ChangeFormVisitDetails, check for accidental changes to an "Open" study (Issue 2405)
' NCJ Jun 06 - MUSD - Consider access modes
' NCJ 24 Aug 06 - Check for study updates on any schedule click; include Access Mode in status string
' NCJ 28 Sept 06 - New mnuPStudyVisitViewEForm menu item to view eForm from schedule
'-------------------------------------------------------------------------------------'
Option Explicit
Option Compare Binary
Option Base 0

' PN 02/09/99
'ZA 22/08/2002 - new option to set form for validation
Public Enum CellSelectionType
    AllowSingle
    AllowRepeating
    NotAssigned
    VisitForm
End Enum

' PN 02/09/99
Private Enum MouseAction
    SingleClick
    DoubleClick
    LeftClick
    RightClick
End Enum

Private mnClinicalTrialId As Long
Private mnVersionId As Integer
Private msClinicalTrialName As String
Private msTrialSite As String
Private mnPersonId As Integer
Private mnCRFPageTaskId() As Long
Private mnCurrentCRFPageTaskId As Long
Private mblnRepeating() As Boolean

'TA 17/03/2000
'additional const for user prompting of dates
Private Const mnVISIT_DATE_PROMPT_ON = 1
Private Const mnVISIT_DATE_PROMPT_OFF = 0
Private Const msPROMPT_USER = "Prompt user for date"
Private Const mSVISIT_CYCLE = "1"

Private Const mROW_VISITID = 0
Private Const mROW_VISITORDER = 1
Private Const mROW_VISITCODE = 2
Private Const mROW_VISITNAME = 3
'ZA 05/09/2002 - added for visit cycle
Private Const mROW_VISITCYCLE = 4

'NCJ 11 Sep 02 - VISITCYCLENUMBER AND VISITTASKID rows are not being used
' but leave them in for now because the constants ARE being used (oh yuk)
Private Const mROW_VISITCYCLENUMBER = 5
Private Const mROW_VISITTASKID = 5

Private Const mnCRFPageIdColumn = 0
Private Const mnCRFPageOrderColumn = 1
Private Const mnCRFPageCodeColumn = 2
Private Const mnCRFPageNameColumn = 3
Private Const mnCRFPageCycleNumberColumn = 4

Private Const msEMPTYCELL = vbNullString
Private Const msSINGLECELL = "."
Private Const msREPEATINGCELL = ".."
Private Const msVISIT_EFORMCELL = "..."

'ZA 06/09/2002 - Unlimited number of visits
Private Const msUNLIMITED_VISIT_REPEATS = "Unlimited"
Private Const mnUNLIMITED_VISIT_VALUE = -1
Private gbNotInCell As Boolean

' PN 01/09/99
' udt variable used to cache the selected row at mousedown
' for selection of correct fixed col and row
Private meSelectedGridCell As SelectedGridCell

' PN 01/09/99
' keep track of which mouse button was selected and whether the action was a single
' or double click - this is because a double click mouse action will
' fire both click and dblclick events
' but on a dblclick only the dblclick action is required
Private meMouseButton As MouseAction    ' was the button a left or right button
Private meMouseAction As MouseAction    ' was the action a click or dblclick

'   SR 580  24/11/98    ATN
'   New variables to store the location of the first available form
Public mnFirstFormCol As Integer
Public mnFirstFormRow As Integer

'ZA 27/08/2002 - New variable to store the location of visit form
Private mnVisitFormRow As Integer
Private menCellSelected As CellSelectionType

' NCJ 16 Apr 07 - Store Visit Name and Cycle to detect whether changed

'---------------------------------------------------------------------
Private Sub flxVisits_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
'WillC 28/2/2000 SR2796 Added a tooltip for row which allows editing.
'---------------------------------------------------------------------
' Not working in all screen resolutions

    ' NCJ 14 Jun 06 - Ignore if not in edit mode
    If frmMenu.StudyAccessMode < sdReadWrite Then Exit Sub
    
'    If y >= 370 And y <= 740 And x >= 1500 Then
    With flxVisits
        'check the mouse is at least in the first visible column
        If .MouseCol > mnCRFPageCycleNumberColumn And .MouseCol < .Cols Then
            If .MouseRow = mROW_VISITNAME Then
                .ToolTipText = "Single click to edit visit caption"
            'ZA 28/08/2002 - comment out this line
            ElseIf .MouseRow = mROW_VISITCYCLE Then
                'ZA 06/09/2002 - just use the tooltiptext with out any condition
                'If .TextMatrix(mROW_VISITCYCLE, .MouseCol) = mSVISIT_CYCLE Then
                '    .ToolTipText = "Use right mouse menu to switch to editable cycle"
                'Else
                   .ToolTipText = "Single click to edit number of visit cycle"
                'End If
            Else
                .ToolTipText = vbNullString
            End If
        Else
               .ToolTipText = vbNullString
        End If
    End With


End Sub

Private Sub flxVisits_Scroll()
    'SDM 08/12/99 SR2251
    txtCycleVisit.Visible = False
    txtVisitName.Visible = False
    cmdValidationExpression.Visible = False
End Sub

'---------------------------------------------------------------------
Public Property Get ClinicalTrialId() As Long
'---------------------------------------------------------------------

    ClinicalTrialId = mnClinicalTrialId

End Property

'---------------------------------------------------------------------
Public Property Let ClinicalTrialId(ByVal vClinicalTrialId As Long)
'---------------------------------------------------------------------

    mnClinicalTrialId = vClinicalTrialId

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
    Me.Caption = vClinicalTrialName & " schedule"

End Property

'---------------------------------------------------------------------
Public Property Get TrialSite() As String
'---------------------------------------------------------------------

    TrialSite = msTrialSite

End Property

'---------------------------------------------------------------------
Public Property Let TrialSite(ByVal vTrialSite As String)
'---------------------------------------------------------------------

    msTrialSite = vTrialSite

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

'---------------------------------------------------------------------
Public Property Get CurrentCRFPageTaskId() As Long
'---------------------------------------------------------------------

    PersonId = mnCurrentCRFPageTaskId

End Property

'---------------------------------------------------------------------
Public Property Let CurrentCRFPageTaskId(ByVal lCRFPageTaskId As Long)
'---------------------------------------------------------------------

    mnCurrentCRFPageTaskId = lCRFPageTaskId

End Property

'---------------------------------------------------------------------
Public Sub InsertVisit()
'---------------------------------------------------------------------
' Create a new visit
' Mo Morris 14/2/00  sections of code re-writen
'---------------------------------------------------------------------
Dim sVisitCode As String
Dim lVisitId As Long
Dim sSQL As String
Dim sMSG As String
    
    On Error GoTo ErrHandler

    ' PN 30/08/99
    ' if a previous form name was being edited then save changes
    If txtVisitName.Visible Then
        Call txtVisitName_LostFocus
    End If

    'TA 28/03/2000 - call new function to get code
    sVisitCode = GetItemCode(gsITEM_TYPE_VISIT, "New " & gsITEM_TYPE_VISIT & " code:")
    If sVisitCode = "" Then    ' if cancel, then return control to user
        Exit Sub
    End If

    'Begin transaction
    TransBegin

    ' Create an Arezzo plan for the visit - NCJ 10/8/99
    lVisitId = gnNewCLMPlan(gsCLMVisitName(sVisitCode))

    'Changed by Mo Morris 6/8/99
    'VisitDateLabel and VisitBackgroundColour added
    'TA 20/03/2000 VisitDatePrompt Default of 0 appended
    sSQL = "INSERT INTO StudyVisit " _
            & "( ClinicalTrialId, VersionId, VisitId, " _
            & "VisitCode, VisitName, VisitOrder, VisitDateLabel, VisitBackgroundColour, VisitDatePrompt)" _
            & " VALUES (" _
            & ClinicalTrialId & "," _
            & VersionId & "," _
            & lVisitId & ",'" _
            & sVisitCode & "','" _
            & sVisitCode & "'," _
            & flxVisits.Cols - mnCRFPageCycleNumberColumn & ",'',0,0)"

    MacroADODBConnection.Execute sSQL

    ' This now completes the Arezzo side of things - NCJ 10/8/99
    InsertProformaVisit ClinicalTrialId, lVisitId, flxVisits.Cols - mnCRFPageCycleNumberColumn
    
    'End transaction
    TransCommit
    
    ' NCJ 19 Jun 06 - Mark as changed
    Call frmMenu.MarkStudyAsChanged

    frmStudyVisits.ClinicalTrialId = Me.ClinicalTrialId
    frmStudyVisits.VersionId = Me.VersionId
    frmStudyVisits.ClinicalTrialName = Me.ClinicalTrialName
    'Following line added by Mo Morris 8/1/99 SR 631
    frmStudyVisits.Show
    frmStudyVisits.RefreshStudyVisits
    
    '  Simulate click to edit name
    flxVisits.Col = flxVisits.Cols - 1
    flxVisits.Row = mROW_VISITNAME
    
    ' PN 01/09/99
    ' force the click event to work by setting the
    ' mouse button variable meMouseButton to LeftClick
    meMouseButton = LeftClick
    'flxVisits_Click
    
Exit Sub
ErrHandler:
    'RollBack transaction
    TransRollBack
    
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "InsertVisit")
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
Public Function mblnVisitExists(ByVal vClinicalTrialId As Long, _
                                ByVal vVersionId As Integer, _
                                ByVal vVisitCode As String) As Boolean
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT VisitId FROM StudyVisit " _
        & " WHERE VisitCode = '" & vVisitCode _
        & "' AND ClinicalTrialId = " & vClinicalTrialId _
        & " AND VersionId = " & vVersionId

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

    If rsTemp.RecordCount = 0 Then
        mblnVisitExists = False
    Else
        mblnVisitExists = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "mblnVisitExists")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Function

'---------------------------------------------------------------------
Public Sub Delete()
'---------------------------------------------------------------------
Dim mnResponse As Single
Dim msSQL As String
Dim mForm As Form
  On Error GoTo ErrHandler

If flxVisits.Col > mnCRFPageCycleNumberColumn Then
    Set mForm = frmMenu
    If mForm.TrialStatus > 1 Then
        MsgBox "You cannot delete a visit once a study has been opened", _
            vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        Exit Sub
    End If


'   SR 617      24/11/98    ATN
'   Displays visit name in message
'SR3664 ASH 2/08/2001
'Reworded message to match that of deleting questions and eforms
    mnResponse = MsgBox("Are you sure you want to delete visit : " & flxVisits.TextMatrix(mROW_VISITNAME, flxVisits.Col) & " ?", _
            vbYesNo + vbExclamation + vbDefaultButton2 + vbApplicationModal, gsDIALOG_TITLE)
    If mnResponse = vbYes Then
    
        'Begin transaction
        TransBegin
    
        msSQL = "DELETE FROM StudyVisit " _
                & "WHERE ClinicalTrialId = " & ClinicalTrialId _
                & " AND VersionId = " & VersionId _
                & " AND VisitId = " & flxVisits.TextMatrix(mROW_VISITID, flxVisits.Col)
                            
        MacroADODBConnection.Execute msSQL
        
        msSQL = "DELETE FROM StudyVisitCRFPage " _
                & "WHERE ClinicalTrialId = " & ClinicalTrialId _
                & " AND VersionId = " & VersionId _
                & " AND VisitId = " & flxVisits.TextMatrix(mROW_VISITID, flxVisits.Col)
    
        MacroADODBConnection.Execute msSQL
        
'   ATN 12/11/98
'   Added clinicaltrialid parameter to DeleteProformaVisit
        DeleteProformaVisit flxVisits.TextMatrix(mROW_VISITID, flxVisits.Col), Me.ClinicalTrialId
        
'   ATN 1/5/99  SR 875,876
'   Ensure the txtVisitName doesn't have the focus

        txtVisitName.Visible = False
        flxVisits.SetFocus
        
        ReorderVisits
        
        'End transaction
        TransCommit
        
        ' NCJ 19 Jun 06 - Mark as changed
        Call frmMenu.MarkStudyAsChanged

        'Following line added by Mo Morris 8/1/99 SR 631
        Me.Show
        Me.RefreshStudyVisits
        frmMenu.ChangeSelectedItem "", ""
    End If
End If
    

Exit Sub
ErrHandler:
    'RollBack transaction
    TransRollBack
    
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Delete")
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
Private Sub ReorderVisits()
'---------------------------------------------------------------------
Dim nCol As Integer
Dim sSQL As String
  
    On Error GoTo ErrHandler

    'Begin transaction
    TransBegin
    
    For nCol = mnCRFPageCycleNumberColumn + 1 To flxVisits.Cols - 1
        If flxVisits.TextMatrix(mROW_VISITORDER, nCol) <> nCol - mnCRFPageCycleNumberColumn Then
            
            flxVisits.TextMatrix(mROW_VISITORDER, nCol) = nCol - mnCRFPageCycleNumberColumn
            
            sSQL = "UPDATE StudyVisit " _
            & " SET VisitOrder = " & flxVisits.TextMatrix(mROW_VISITORDER, nCol) _
            & " WHERE ClinicalTrialId = " & ClinicalTrialId _
            & " AND VersionId = " & VersionId _
            & " AND VisitId = " & flxVisits.TextMatrix(mROW_VISITID, nCol)
            
            MacroADODBConnection.Execute sSQL
    
        End If
    Next
    
    'End transaction
    TransCommit
    ' NCJ 20 Jun 06 - Mark study as changed
    Call frmMenu.MarkStudyAsChanged

Exit Sub
ErrHandler:
    'RollBack transaction
    TransRollBack
    
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ReorderVisits")
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
Private Sub ReorderCRFPages()
'---------------------------------------------------------------------
' Reorder CRF pages based on ordering in flex grid
' NCJ 10 Jan 02 - Make sure ALL question lists are updated
'---------------------------------------------------------------------
Dim nRow As Integer
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'Begin transaction
    TransBegin
    
    '   15/10/98
    '   Changed start of For..next to be after the mROW_VISITTASKID
    For nRow = mROW_VISITTASKID + 1 To flxVisits.Rows - 1
        If flxVisits.TextMatrix(nRow, mnCRFPageOrderColumn) <> nRow - mROW_VISITCYCLENUMBER Then
            
            flxVisits.TextMatrix(nRow, mnCRFPageOrderColumn) = nRow - mROW_VISITCYCLENUMBER
            
            sSQL = "UPDATE CRFPage " _
            & " SET CRFPageOrder = " & flxVisits.TextMatrix(nRow, mnCRFPageOrderColumn) _
            & " WHERE ClinicalTrialId = " & ClinicalTrialId _
            & " AND VersionId = " & VersionId _
            & " AND CRFPageId = " & flxVisits.TextMatrix(nRow, mnCRFPageIdColumn)
            
            MacroADODBConnection.Execute sSQL
    
        End If
    Next
    
     'WillC 24/2/2000 SR2797 let the datalist reflect the change in eForm order
    ' NCJ 10 Jan 02 - Refresh ALL question lists
    ' frmDataList.RefreshDataList
    Call frmMenu.RefreshQuestionLists(Me.ClinicalTrialId)
    
    'End transaction
    TransCommit
    ' NCJ 20 Jun 06 - Mark study as changed
    Call frmMenu.MarkStudyAsChanged

Exit Sub
ErrHandler:
    'RollBack transaction
    TransRollBack
    
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ReorderCRFPages")
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
Private Function DeleteStudyVisitCRFPage() As Boolean
'---------------------------------------------------------------------
' Delete a study visit crf page
' NCJ 12 Jan 00 - Check study not open and return FALSE if change not done
' otherwise return TRUE
'---------------------------------------------------------------------
Dim sSQL As String
Dim sMSG As String
Dim sSQLExtended As String

    On Error GoTo ErrHandler

    DeleteStudyVisitCRFPage = True
    
    ' NCJ 12 Jan 00 - Check trial status
    If frmMenu.TrialStatus > eTrialStatus.InPreparation Then
        sMSG = "You may not remove an eForm from a Visit once a study has been opened"
        MsgBox sMSG, vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        DeleteStudyVisitCRFPage = False
    Else
        
        If IsValidationFormAlreadyThere(flxVisits.Col) And menCellSelected = CellSelectionType.VisitForm Then
            sSQLExtended = flxVisits.TextMatrix(mnVisitFormRow, mnCRFPageIdColumn)
        Else
            sSQLExtended = flxVisits.TextMatrix(flxVisits.Row, mnCRFPageIdColumn)
            
        End If
        sSQL = "DELETE FROM StudyVisitCRFPage " _
                & "WHERE ClinicalTrialId = " & ClinicalTrialId _
                & " AND VersionId = " & VersionId _
                & " AND VisitId = " & flxVisits.TextMatrix(mROW_VISITID, flxVisits.Col) _
                & " And CRFPageId = " & sSQLExtended
              '  & " AND CRFPageId = " & flxVisits.TextMatrix(flxVisits.Row, mnCRFPageIdColumn)
        
        MacroADODBConnection.Execute sSQL
        
        'ZA .....
        If IsValidationFormAlreadyThere(flxVisits.Col) And menCellSelected = CellSelectionType.VisitForm Then
            DeleteProformaStudyVisitCRFPage flxVisits.TextMatrix(mROW_VISITID, flxVisits.Col), _
                flxVisits.TextMatrix(mnVisitFormRow, mnCRFPageIdColumn)
            'mnVisitFormRow is the row where we need to remove the Visit Form
            flxVisits.Row = mnVisitFormRow
        Else
            DeleteProformaStudyVisitCRFPage flxVisits.TextMatrix(mROW_VISITID, flxVisits.Col), _
                flxVisits.TextMatrix(flxVisits.Row, mnCRFPageIdColumn)
            
        End If
        
        'ZA 27/08/2002 - change the current row to the row that contains Visit Form
        Set flxVisits.CellPicture = Nothing
        'Now change the row back to where the mouse click has occurred
        flxVisits.Row = meSelectedGridCell.Row
    End If
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "DeleteStudyVisitCRFPage")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

'---------------------------------------------------------------------
Private Sub UpdateStudyVisitCRFPage(nRepeating As Integer, _
                                    nVisitID As Long, _
                                    nCRFPageID As Long, _
                                    bVisitDateForm As Boolean)
'---------------------------------------------------------------------
' update a study visit crf page
'ZA 23/08/2002 - Added check for visit date validation
'---------------------------------------------------------------------
Dim sSQL As String
  
    On Error GoTo ErrHandler
    
    sSQL = "UPDATE StudyVisitCRFPage " _
        & " SET Repeating = " & nRepeating & ", " _
        & " EFormUse = " & IIf(bVisitDateForm, eEFormUse.VisitEForm, eEFormUse.User) _
        & " WHERE ClinicalTrialId = " & ClinicalTrialId _
        & " AND VersionId = " & VersionId _
        & " AND VisitId = " & nVisitID _
        & " AND CRFPageId = " & nCRFPageID
    
    MacroADODBConnection.Execute sSQL
   
    ' Insert Proforma editor update call here
    ' Add cycling information to the CRF plan inside the Visit plan
    ' NCJ 27/8/99
    'ZA 23/08/2002 - check for Visit eform
    If bVisitDateForm Then
        SetCyclingTask nVisitID, nCRFPageID, False
        Set flxVisits.CellPicture = imgLargeIcons.ListImages(gsVISIT_EFORM).Picture
        
    Else
        If nRepeating = 0 Then
            SetCyclingTask nVisitID, nCRFPageID, False
            Set flxVisits.CellPicture = imgLargeIcons.ListImages(gsLARGE_CRF_PAGE_LABEL).Picture
        Else
            SetCyclingTask nVisitID, nCRFPageID, True
            Set flxVisits.CellPicture = imgLargeIcons.ListImages(gsREPEATING_CRF_PAGE_LABEL).Picture
        End If
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "UpdateStudyVisitCRFPage")
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
Private Sub InsertStudyVisitCRFPage(iRepeating As Integer, _
                                    nVisitID As Long, _
                                    nCRFPageID As Long, _
                                    enEFormUse As eEFormUse)
'---------------------------------------------------------------------
' insert a study visit crf page
'ZA 22/08/2002 - added enFormUse parameter for validation
'---------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    'changed by Mo Morris 6/8/99 mnNodeTag replaced by 0
    ' Page code and NodeTag removed NCJ 13/8/99
    InsertProformaStudyVisitCRFPage _
        ClinicalTrialId, _
        nVisitID, _
        nCRFPageID, _
        flxVisits.TextMatrix(flxVisits.Row, mnCRFPageOrderColumn)
    
    'changed by Mo Morris 6/8/99, NodeTag removed from sql statement
    sSQL = "INSERT INTO StudyVisitCRFPage " _
            & "VALUES (" & ClinicalTrialId & "," _
            & VersionId & "," _
            & nVisitID & "," _
            & nCRFPageID & "," & iRepeating & "," & enEFormUse & ")"
            
    MacroADODBConnection.Execute sSQL
    
    ' Added calls to SetCyclingTask here - NCJ 27 Sept 99
    If iRepeating = 0 And enEFormUse = eEFormUse.User Then   ' Not repeating
        Set flxVisits.CellPicture = imgLargeIcons.ListImages(gsLARGE_CRF_PAGE_LABEL).Picture
        SetCyclingTask nVisitID, nCRFPageID, False
    'ZA 22/08/2002 - check if this form is being used as visit eForm
    ElseIf enEFormUse = eEFormUse.VisitEForm Then
        Set flxVisits.CellPicture = imgLargeIcons.ListImages(gsVISIT_EFORM).Picture
    Else
    ' Repeating form
        Set flxVisits.CellPicture = imgLargeIcons.ListImages(gsREPEATING_CRF_PAGE_LABEL).Picture
        SetCyclingTask nVisitID, nCRFPageID, True
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "InsertStudyVisitCRFPage")
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
Public Sub ChangeFormVisitDetails(eNewCellData As CellSelectionType)
'---------------------------------------------------------------------
' this function will manage all types of change to the selection of the
' studyvisit crfpage cell
' ie from any CellSelectionType to any other CellSelectionType
' NCJ 10/12/99 - Check user function access here
' (may be called via clicks or menu options)
' NCJ 13/1/00 - Make sure we do Exit Sub if user does not have access rights
'               Also exit on result of DeleteStudyVisitCRFPage as necessary
' NCJ 2 Nov 04 - Check for accidental changes to an "Open" study (Issue 2405)
' MLM 24/06/05: bug 2544: Tidied up transaction and error handling.
'---------------------------------------------------------------------
Dim lVisitId As Long
Dim lCRFPageId As Long
Dim iMousePointer As Integer
Dim sMSG As String
Dim nErr As Integer
Dim sErrDesc As String

   On Error GoTo ErrHandler
      
    ' save the pointer -  to restore later
    iMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    With flxVisits
        If .TextMatrix(.Row, mnCRFPageIdColumn) <> vbNullString Then
        
            'Begin transaction
            TransBegin
            On Error GoTo ErrHandlerRollback
            
            lCRFPageId = .TextMatrix(.Row, mnCRFPageIdColumn)
            lVisitId = .TextMatrix(mROW_VISITID, .Col)
            
            'ZA 27/08/2002 - keep cellselection value to be used in DeleteVisitCRFPage routine
            menCellSelected = eNewCellData
            ' determine the state of the cell now
            Select Case flxVisits
            Case msEMPTYCELL
                ' NCJ 2 Nov 04 - Check permission here (rather than for each case)
                If Not goUser.CheckPermission(gsFnAddEFormToVisit) Then
                    TransRollBack
                    Screen.MousePointer = iMousePointer
                    Exit Sub
                End If
                ' NCJ 2 Nov 04 - Check first for an "Open" study
                If frmMenu.TrialStatus > 1 Then
                    sMSG = "Are you sure you wish to add an eForm to this visit?" & vbCrLf _
                            & "Once added it may not be removed (because the study has been opened)"
                    If DialogQuestion(sMSG) = vbNo Then
                        ' They've thought better of it
                        TransRollBack
                        Screen.MousePointer = iMousePointer
                        Exit Sub
                    End If
                End If
                Select Case eNewCellData
                Case AllowSingle
                    ' was an empty cell now set it to a single cell
'                    ' NCJ Must check user's access rights here
'                    If goUser.CheckPermission(gsFnAddEFormToVisit) Then
                        Call InsertStudyVisitCRFPage(0, lVisitId, lCRFPageId, eEFormUse.User)
'                    Else
'                        Screen.MousePointer = iMousePointer
'                        Exit Sub
'                    End If
                Case AllowRepeating
                    ' was an empty cell now set it to a repeating cell
'                    If goUser.CheckPermission(gsFnAddEFormToVisit) Then
                        Call InsertStudyVisitCRFPage(1, lVisitId, lCRFPageId, eEFormUse.User)
'                    Else
'                        Screen.MousePointer = iMousePointer
'                        Exit Sub
'                    End If
                'ZA 22/08/2002 -
                Case VisitForm
'                    If goUser.CheckPermission(gsFnAddEFormToVisit) Then
                        If IsValidationFormAlreadyThere(flxVisits.Col) Then
                            If DeleteStudyVisitCRFPage = False Then
                                TransRollBack
                                Screen.MousePointer = iMousePointer
                                Exit Sub
                            End If
                        End If
                        Call InsertStudyVisitCRFPage(0, lVisitId, lCRFPageId, eEFormUse.VisitEForm)
'                    Else
'                        Screen.MousePointer = iMousePointer
'                    End If
                
                End Select
        
            Case msSINGLECELL
                Select Case eNewCellData
                Case NotAssigned
                    ' was a single cell now set it to an empty
                    ' Only if trial status = in preparation
                    If goUser.CheckPermission(gsFnRemoveEFormFromVisit) Then
                        If DeleteStudyVisitCRFPage = False Then
                            TransRollBack
                            Screen.MousePointer = iMousePointer
                            Exit Sub
                        End If
                    Else
                        TransRollBack
                        Screen.MousePointer = iMousePointer
                        Exit Sub
                    End If
                    
                Case AllowRepeating
                    ' was a single cell now set it to a repeating cell
                    If goUser.CheckPermission(gsFnAddEFormToVisit) Then
                        ' NCJ 2 Nov 04 - Check first for an "Open" study
                         If frmMenu.TrialStatus > 1 Then
                             sMSG = "Are you sure you wish to change this eForm to be repeating?"
                             If DialogQuestion(sMSG) = vbNo Then
                                 ' They've thought better of it
                                 TransRollBack
                                 Screen.MousePointer = iMousePointer
                                 Exit Sub
                             End If
                         End If
                        Call UpdateStudyVisitCRFPage(1, lVisitId, lCRFPageId, False)
                    Else
                        TransRollBack
                        Screen.MousePointer = iMousePointer
                        Exit Sub
                    End If
                'ZA 22/08/2002 - change existing cell to form validation
                Case VisitForm
                    'clear the existing cell and put form validation cell
                    If goUser.CheckPermission(gsFnAddEFormToVisit) Then
                        'delete the previous Visit Form if found on this column
                        If IsValidationFormAlreadyThere(meSelectedGridCell.Col) Then
                            If DeleteStudyVisitCRFPage = False Then
                                TransRollBack
                                Screen.MousePointer = iMousePointer
                                Exit Sub
                            End If
                        End If
                        Call UpdateStudyVisitCRFPage(0, lVisitId, lCRFPageId, True)
                    Else
                        Screen.MousePointer = iMousePointer
                    End If
                End Select
           
        
            Case msREPEATINGCELL
                Select Case eNewCellData
                Case NotAssigned
                    ' was a repeating cell now set it to an empty cell
                    If goUser.CheckPermission(gsFnRemoveEFormFromVisit) Then
                        If DeleteStudyVisitCRFPage = False Then
                            TransRollBack
                            Screen.MousePointer = iMousePointer
                            Exit Sub
                        End If
                    Else
                        TransRollBack
                        Screen.MousePointer = iMousePointer
                        Exit Sub
                    End If
                
                Case AllowSingle
                    ' was a repeating cell now set it to a single cell
                    If goUser.CheckPermission(gsFnRemoveEFormFromVisit) Then
                        Call UpdateStudyVisitCRFPage(0, lVisitId, lCRFPageId, False)
                    Else
                        TransRollBack
                        Screen.MousePointer = iMousePointer
                        Exit Sub
                    End If
              
                'ZA 22/08/2002
                Case VisitForm
                    'clear the existing cell and put form validation cell
                    If goUser.CheckPermission(gsFnAddEFormToVisit) Then
                        If IsValidationFormAlreadyThere(meSelectedGridCell.Col) Then
                            If DeleteStudyVisitCRFPage = False Then
                                TransRollBack
                                Screen.MousePointer = iMousePointer
                                Exit Sub
                            End If
                        End If
                        Call UpdateStudyVisitCRFPage(0, lVisitId, lCRFPageId, True)
                    Else
                        Screen.MousePointer = iMousePointer
                    End If
                    
                Case AllowRepeating
                    ' was a single cell now set it to a repeating cell
                    If goUser.CheckPermission(gsFnAddEFormToVisit) Then
                        Call UpdateStudyVisitCRFPage(1, lVisitId, lCRFPageId, False)
                    Else
                        TransRollBack
                        Screen.MousePointer = iMousePointer
                        Exit Sub
                    End If
                    
                End Select
            'ZA 23/08/2002 - TODO cycling of form in visits
            Case msVISIT_EFORMCELL
                'user clicked on Visit Form cell
                Select Case eNewCellData
                'clear out the cell
                Case NotAssigned
                    If goUser.CheckPermission(gsFnRemoveEFormFromVisit) Then
                        If DeleteStudyVisitCRFPage = False Then
                            TransRollBack
                            Screen.MousePointer = iMousePointer
                            Exit Sub
                        End If
                    Else
                        TransRollBack
                        Screen.MousePointer = iMousePointer
                        Exit Sub
                    End If
                    
                Case AllowSingle
                    If goUser.CheckPermission(gsFnAddEFormToVisit) Then
                        Call UpdateStudyVisitCRFPage(0, lVisitId, lCRFPageId, False)
                    Else
                        TransRollBack
                        Screen.MousePointer = iMousePointer
                        Exit Sub
                    End If
                    
                Case AllowRepeating
                    ' was a single cell now set it to a repeating cell
                    If goUser.CheckPermission(gsFnAddEFormToVisit) Then
                        Call UpdateStudyVisitCRFPage(1, lVisitId, lCRFPageId, False)
                    Else
                        TransRollBack
                        Screen.MousePointer = iMousePointer
                        Exit Sub
                    End If

                End Select
            End Select
            
            'End transaction
            TransCommit
            On Error GoTo ErrHandler
            ' NCJ 20 Jun 06 - Mark study as changed
            Call frmMenu.MarkStudyAsChanged

            .CellPictureAlignment = flexAlignCenterCenter
            .CellAlignment = flexAlignCenterCenter
            
            If eNewCellData = AllowRepeating Then
                .Text = msREPEATINGCELL
            
            ElseIf eNewCellData = AllowSingle Then
                .Text = msSINGLECELL
            'ZA 22/08/2002 -
            ElseIf eNewCellData = VisitForm Then
                If IsValidationFormAlreadyThere(meSelectedGridCell.Col) And mnVisitFormRow <> meSelectedGridCell.Row Then
                    .TextMatrix(mnVisitFormRow, meSelectedGridCell.Col) = msEMPTYCELL
                End If
                .Text = msVISIT_EFORMCELL
            Else
                .Text = msEMPTYCELL
            
            End If
        
        End If
    
    End With
    
    ' restore mouse pointer
    Screen.MousePointer = iMousePointer

Exit Sub
        
ErrHandlerRollback:
    nErr = Err.Number
    sErrDesc = Err.Description & "|frmStudyVisits.ChangeFormVisitDetails"
    'RollBack transaction
    TransRollBack
    Err.Raise nErr, , sErrDesc
Exit Sub

ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmStudyVisits.ChangeFormVisitDetails"
End Sub

'---------------------------------------------------------------------
Private Sub GridClick()
'---------------------------------------------------------------------
' grid click is only handled in response to the timer event
' this prevents the click and dblclick events both being raised at each dblclick
'---------------------------------------------------------------------
Dim bDisableEdit As Boolean
Dim txtTextBox As TextBox
 
    On Error GoTo ErrHandler
    
    ' NCJ 20 Jun 06 - Check for study updates
    ' NCJ 24 Aug 06 - Do this before anything else
    If frmMenu.RefreshIsNeeded Then Exit Sub
    
    If gbNotInCell Then Exit Sub 'Set if the cursor in a cell when clicked
    
    ' NCJ 14 Jun 06 - Disallow changes if not read-write
    If frmMenu.StudyAccessMode < sdReadWrite Then Exit Sub
    
    If MousePointer <> vbDefault Then
        MousePointer = vbDefault
        Exit Sub
    End If
    
    With flxVisits
        If .Col <= mnCRFPageCycleNumberColumn Then
            ' Select the row
            .RowSel = .Rows - 1
        
        'SDM 26/01/00 SR2783
        'ZA 06/09/2002 - check for visit name or visit cycle
        ElseIf meSelectedGridCell.Row = mROW_VISITCYCLE Or _
               meSelectedGridCell.Row = mROW_VISITNAME Then
      '  ElseIf .MouseRow = mROW_VISITCYCLE Or _
      '         .MouseRow = mROW_VISITNAME Then
            ' Select the column
            .ColSel = .Cols - 1
            .Row = meSelectedGridCell.Row
            .Col = meSelectedGridCell.Col

            ' Only allow editing here if user can maintain the visit - NCJ 10/12/99
            If goUser.CheckPermission(gsFnMaintVisit) Then
                If meSelectedGridCell.Row = mROW_VISITCYCLE Then
                    Set txtTextBox = txtCycleVisit
                Else
                    'visit name
                    Set txtTextBox = txtVisitName
                End If
                    
                'ZA 28/08/2002 - don't need this any more
                If Not bDisableEdit Then
                ' Prepare text box to edit the visit name/date
                    txtTextBox.Move .Left + .CellLeft, _
                            .Top + .CellTop, _
                            .CellWidth, _
                            .CellHeight
    
                        With txtTextBox
                            .Text = flxVisits.Text
                            .Tag = flxVisits.Col
                            .Enabled = True
                            .Visible = True
                            .SetFocus
    
                        End With
                    End If
            End If      ' If user can edit name or date
            
        ' NB User access checks exist in ChangeFormVisitDetails
        Else
            If flxVisits = msEMPTYCELL Then
                ' If it's empty, change it to a single non-repeating form
                Call ChangeFormVisitDetails(AllowSingle)
            
            ElseIf flxVisits.Text = msSINGLECELL Then
                ' If it's a single non-repeating form, change it to a repeating form
                Call ChangeFormVisitDetails(AllowRepeating)
            
            ElseIf flxVisits.Text = msREPEATINGCELL Then
                ' If it's a repeating form then change it to an empty cell
                Call ChangeFormVisitDetails(NotAssigned)
                
            ElseIf flxVisits.Text = msVISIT_EFORMCELL Then
                'if is is a validtion form, change it to an empty cell
                Call ChangeFormVisitDetails(NotAssigned)
            End If
        End If
        
    End With

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GridClick")
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
Private Sub cmdValidationExpression_Click()
'---------------------------------------------------------------------
' Changed from condition to expression - NCJ 30/11/99
' NCJ 21 Sept 06 - No longer used
'---------------------------------------------------------------------

'    On Error GoTo ErrHandler
'
'    frmCUIFunctionEditor.Initialise txtCycleVisit, _
'                        "Expression", "Visit Date Expression"
'    cmdValidationExpression.Visible = False 'SDM 08/12/99 SR2251
'Exit Sub
'
'ErrHandler:
'    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
'                                    "cmdValidationExpression_Click")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            End
'    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub flxVisits_Click()
'---------------------------------------------------------------------
' If the right mouse was clicked then ignore the event
' NCJ 18/1/00 SR 2586 Check they've clicked in the same place
'---------------------------------------------------------------------
   
   On Error GoTo ErrHandler

    ' NCJ meSelectedGridCell is where the mouse down occurred
   If meSelectedGridCell.Col = flxVisits.MouseCol _
     And meSelectedGridCell.Row = flxVisits.MouseRow Then
   
        If meMouseButton = LeftClick Then
            meMouseAction = SingleClick

            tmrGridClick.Enabled = True
'            Screen.MousePointer = vbHourglass
        End If
    Else
        MousePointer = vbNormal
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "flxVisits_Click")
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
Private Sub flxVisits_DblClick()
'---------------------------------------------------------------------
' if the right mouse was clicked then ignore the event
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If meMouseButton = LeftClick Then
        meMouseAction = DoubleClick
        tmrGridClick.Enabled = True
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "flxVisits_DblClick")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'--------------------------------------------------------
Public Sub GridDblClick()
'--------------------------------------------------------
' PN 02/09/99
' Respond to a double click - show the selected CRF page
' NCJ 28 Sept 06 - Made Public so that it can be called from frmMenu
'--------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' NCJ 20 Jun 06 - Check for study updates
    If frmMenu.RefreshIsNeeded Then Exit Sub
    
    If IsEditableCellSelected Then
        ' a cell that could have a crf page in has been selected
        ' so show the form
        With flxVisits
            Call frmMenu.ViewCRF(.TextMatrix(.Row, mnCRFPageIdColumn))
        End With
        ' Hide the Visits schedule first (SR 2065) - NCJ 9 Nov 99
        frmMenu.HideVisits
  
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GridDblClick")
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
Private Function IsEditableCellSelected() As Boolean
'---------------------------------------------------------------------
' is the cell currently selected editable
'---------------------------------------------------------------------
    On Error GoTo ErrHandler
    
  With flxVisits
        If .Col > mnCRFPageCycleNumberColumn And _
            .Row > mROW_VISITTASKID Then
            IsEditableCellSelected = True
            
        End If
    End With

Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "IsEditableCellSelected")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

'---------------------------------------------------------------------
Private Function IsEditableColumnSelected() As Boolean
'---------------------------------------------------------------------

     On Error GoTo ErrHandler
    
 If flxVisits.Col > mnCRFPageCycleNumberColumn Then
        IsEditableColumnSelected = True

    End If

Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "IsEditableColumnSelected")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

'---------------------------------------------------------------------
Private Sub flxVisits_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
    ' SPR 527
    ' A variable gbNotInCell is set depending on if the cursor is in the cell
    ' on the FlexGrid when clicked. If it is'nt then the cell remains unchanged.
    ' This is to prevent any new forms being inserted in the Schedule or being deleted.
    ' However the Click event is still being triggered on a cell even though the click
    ' occured outside the grid.
'---------------------------------------------------------------------
Dim sVisitName As String

    On Error GoTo ErrHandler
        
    meSelectedGridCell.Col = flxVisits.MouseCol
    meSelectedGridCell.Row = flxVisits.MouseRow
    
    If Button = vbLeftButton Then
                
        flxVisits.Tag = gsEMPTY_STRING
        If X > flxVisits.CellLeft + flxVisits.CellWidth Or _
                Y > flxVisits.CellTop + flxVisits.CellHeight Then
        
            gbNotInCell = True
            flxVisits.FocusRect = flexFocusNone
        Else
            gbNotInCell = False
            flxVisits.FocusRect = flexFocusLight
        End If
        
        ' NCJ 22 Dec 99, SR 1767 Do not select visit if in form column
        sVisitName = flxVisits.TextMatrix(mROW_VISITNAME, flxVisits.MouseCol)
        If sVisitName > "" Then
            frmMenu.ChangeSelectedItem gsVISIT_LABEL, _
                "Visit : " & sVisitName & " (" & GetAccessModeString(frmMenu.StudyAccessMode, frmMenu.AllowMU) & ")"
        Else
            frmMenu.ChangeSelectedItem "", ""
        End If
        
        If frmMenu.StudyAccessMode >= sdReadWrite Then
            ' NCJ 14 Jun 06 - No reordering visits/forms if read only
            If flxVisits.MouseRow = mROW_VISITCODE _
            And flxVisits.MouseCol > mnCRFPageCycleNumberColumn _
            And gbNotInCell = False Then
                flxVisits.Tag = CStr(flxVisits.MouseCol)
                MousePointer = vbSizeWE
            '   15/10/98
            '   Changed row test to be after the mROW_VISITTASKID
            ElseIf flxVisits.MouseCol = mnCRFPageNameColumn _
            And flxVisits.MouseRow > mROW_VISITTASKID _
            And gbNotInCell = False Then
                flxVisits.Tag = CStr(flxVisits.MouseRow)
                MousePointer = vbSizeNS
            End If
        Else
            ' NCJ 7 Sept 06 - Check for study changes and refresh if necessary
            Call frmMenu.RefreshIsNeeded
        End If
    
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "flxVisits_MouseDown")
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
Private Function IsUserPrompt() As Boolean
'---------------------------------------------------------------------
' TA 17/03/2000
' returns true if the prompt for date option is set
'---------------------------------------------------------------------
   ' IsUserPrompt = flxVisits.TextMatrix(mROW_VISITCYCLE, flxVisits.Col) = msPROMPT_USER
    
End Function

'---------------------------------------------------------------------
Private Function IsNotEditMode() As Boolean
'---------------------------------------------------------------------
' TA 17/03/2000
' returns false only if txtCycleVisit visible and it is the current column
'---------------------------------------------------------------------

 '   IsNotEditMode = Not (txtCycleVisit.Visible And (Val(txtCycleVisit.Tag) = flxVisits.Col))
    
End Function


'---------------------------------------------------------------------
Private Sub ShowPopupMenu()
'---------------------------------------------------------------------
' PN 02/09/99
' display the popup menu options
'---------------------------------------------------------------------
Dim bShowMenu As Boolean
  
  On Error GoTo ErrHandler
    'ASH 17/12/2002
    'called here to avoid MACRO crashing when visits deleted when
    'visit cycle field still has focus
    If Me.ActiveControl.Name = "txtCycleVisit" Then
        txtCycleVisit_LostFocus
    End If
    
    ' PN 26/09/99 show the InsertVisit menu option for clicking in white space
    ' first apply the context
    If gbNotInCell Then
        ' the mouse was clicked outside any editable cells and outside
        ' any columns
        Call SetupMenuOptions(False, True, False, False, False, False, False, False, False)
        bShowMenu = True
        
    Else
        If IsEditableCellSelected Then
            ' an editable cell was clicked
            'ZA 23/08/2002 - only allow one form for validation in a Visit
            If flxVisits = msSINGLECELL Then
                ' NCJ 24 Oct 06 - Corrected handling of IsValidationFormAlreadyThere
                Call SetupMenuOptions(False, False, True, False, True, False, False, _
                                    Not IsValidationFormAlreadyThere(flxVisits.MouseCol), True)
'                If IsValidationFormAlreadyThere(flxVisits.MouseCol) Then
'                   Call SetupMenuOptions(False, False, True, False, True, False, False, True, True)
'                Else
'                    Call SetupMenuOptions(False, False, True, False, True, False, False, True, True)
'                End If
                bShowMenu = True
            
            ElseIf flxVisits = msEMPTYCELL Then
                ' NCJ 24 Oct 06 - Corrected handling of IsValidationFormAlreadyThere
                Call SetupMenuOptions(False, False, True, True, False, False, False, _
                                    Not IsValidationFormAlreadyThere(flxVisits.MouseCol), True)
'                If IsValidationFormAlreadyThere(flxVisits.MouseCol) Then
'                    Call SetupMenuOptions(False, False, True, True, False, False, False, True, True)
'                Else
'                    Call SetupMenuOptions(False, False, True, True, False, False, False, True, True)
'                End If
                bShowMenu = True
                
            ElseIf flxVisits = msREPEATINGCELL Then
                Call SetupMenuOptions(False, False, False, True, True, False, False, _
                                    Not IsValidationFormAlreadyThere(flxVisits.MouseCol), True)
                bShowMenu = True
            ElseIf flxVisits = msVISIT_EFORMCELL Then
                Call SetupMenuOptions(False, False, True, True, True, False, False, False, True)
                bShowMenu = True
            End If
        
        ElseIf IsEditableColumnSelected Then
            ' a column was clicked
            ' NCJ 24 Oct 06 - Use False rather than defunct IsNotEditMode, IsUserPrompt
            Call SetupMenuOptions(True, True, False, False, False, False, False, False, False)
'            Call SetupMenuOptions(True, True, False, False, False, IsNotEditMode, IsUserPrompt, False, False)
            bShowMenu = True
            
        Else
            bShowMenu = False
            
        End If
        
    End If
    
    ' then show the menu options
    If bShowMenu Then
        Call PopupMenu(frmMenu.mnuPStudyVisit, vbPopupMenuRightButton)
    
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ShowPopupMenu")
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
Private Sub SetupMenuOptions(bEnableStudyOptions As Boolean, _
                            bEnableInsertOption As Boolean, _
                            bEnableMultiple As Boolean, _
                            bEnableSingle As Boolean, _
                            bEnableRemove As Boolean, bEnablePrompt As Boolean, _
                            bCheckPrompt As Boolean, bDateValidation As Boolean, _
                            bEnableViewEForm As Boolean)
'---------------------------------------------------------------------
' enable/disable apropriate menu options
' NCJ 10/12/99 - Added user rights access checking
' NCJ 27 Jun 06 - Disallow Delete unless Full Control access
' NCJ 28 Sept 06 - Added bEnableViewEForm
'---------------------------------------------------------------------
  On Error GoTo ErrHandler
    
    With frmMenu
        ' Disable them all first
        .mnuPStudyVisitBackgroundColour.Enabled = False
        .mnuPStudyVisitDeleteVisit.Enabled = False
        
        .mnuPStudyVisitInsertVisit.Enabled = False
        
        .mnuPStudyVisitAllowMultipleForms.Enabled = False
        .mnuPStudyVisitAllowSingleForm.Enabled = False
        .mnuPStudyVisitRemoveFormFromVisit.Enabled = False
        .mnuUseFormForDateValidation = False
        
        ' NCJ 28 Sept 06 - New "View eForm"
        .mnuPStudyVisitViewEform.Enabled = bEnableViewEForm
        
        ' Now enable the ones we want
        If bEnableStudyOptions And goUser.CheckPermission(gsFnMaintVisit) Then
            .mnuPStudyVisitBackgroundColour.Enabled = True
        End If
        
        ' NCJ 27 Jun 06 - Only allow Delete for Full Control users
        If bEnableStudyOptions And goUser.CheckPermission(gsFnDelVisit) And (frmMenu.StudyAccessMode = sdFullControl) Then
            .mnuPStudyVisitDeleteVisit.Enabled = True
        End If
        
        If bEnableInsertOption And goUser.CheckPermission(gsFnCreateVisit) Then
            .mnuPStudyVisitInsertVisit.Enabled = bEnableInsertOption
        End If
        
        If bEnableMultiple And goUser.CheckPermission(gsFnAddEFormToVisit) Then
            .mnuPStudyVisitAllowMultipleForms.Enabled = True
        End If
        ' NB Here we don't know whether allowing single
        ' is adding or deleting a form...
        ' Must let it be trapped when action is performed
        If bEnableSingle And goUser.CheckPermission(gsFnAddEFormToVisit) _
         Or goUser.CheckPermission(gsFnRemoveEFormFromVisit) Then
            .mnuPStudyVisitAllowSingleForm.Enabled = bEnableSingle
        End If
        If bEnableRemove And goUser.CheckPermission(gsFnRemoveEFormFromVisit) Then
            .mnuPStudyVisitRemoveFormFromVisit.Enabled = bEnableRemove
        End If
        
        'ZA 22/08/2002 -
        If bDateValidation And goUser.CheckPermission(gsFnAddEFormToVisit) Then
            .mnuUseFormForDateValidation.Enabled = True
        End If
        
    End With

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "SetupMenuOptions")
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
Private Function MouseUpIsOK(nMouseDownCol As Integer, nMouseDownRow As Integer) As Boolean
'---------------------------------------------------------------------
' NCJ 18/1/00 SR 2586
' Check to see if the mouse up has occurred at a "sensible" position
' in relation to the location of the mouse down
' i.e. only allow dragging up and down form column
' or left and right along visit code row
'---------------------------------------------------------------------
Dim nMouseColNow As Integer
Dim nMouseRowNow As Integer

    MouseUpIsOK = False
    
    nMouseColNow = flxVisits.MouseCol
    nMouseRowNow = flxVisits.MouseRow
    
    If nMouseDownCol = nMouseColNow And nMouseDownRow = nMouseRowNow Then
        ' They've clicked in the same place
        MouseUpIsOK = True
    
    ElseIf nMouseDownCol = mnCRFPageNameColumn And nMouseDownRow > mROW_VISITTASKID Then
        ' They started in the CRFPage column
        ' they must finish in the same column and on a sensible row
        If nMouseColNow = nMouseDownCol And nMouseRowNow > mROW_VISITTASKID Then
            MouseUpIsOK = True
        End If
        
    ElseIf nMouseDownRow = mROW_VISITCODE And nMouseDownCol > mnCRFPageCycleNumberColumn Then
        ' They started on a visit code
        ' Only allow dragging from visit column to visit column within this row
        If nMouseRowNow = nMouseDownRow And nMouseColNow > mnCRFPageCycleNumberColumn Then
            MouseUpIsOK = True
        End If
    
    End If
    
    
End Function

'---------------------------------------------------------------------
Private Sub flxVisits_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' NCJ 26 Jun 06 - MUSD - Ignore if no write access
' NCJ 16 Apr 07 - MUSD - Don't reorder visits unless they dragged to a new cell
'---------------------------------------------------------------------
Dim oForm As Form

    On Error GoTo ErrHandler

    ' NCJ 28 Sept 06 - Must set meMouseButton here so we can do Double Clicks in RO mode
    If Button = vbLeftButton Then
        meMouseButton = LeftClick
    Else
        meMouseButton = RightClick
    End If
    
    ' NCJ 26 Jun 06 - Ignore if not in edit mode
    If frmMenu.StudyAccessMode < sdReadWrite Then Exit Sub
    
    ' NCJ 18/1/00 SR2586 Ignore "silly" drags
    ' meSelectedGridCell is where the mouse down occurred
    If Not MouseUpIsOK(meSelectedGridCell.Col, meSelectedGridCell.Row) Then
        Exit Sub
    End If
    
    With flxVisits
        
        If Button = vbLeftButton Then
            ' NCJ 16 Apr 07 - Check that they actually moved the mouse to a new cell
            If .Tag <> gsEMPTY_STRING _
            And MousePointer = vbSizeWE _
            And .MouseCol > mnCRFPageCycleNumberColumn _
            And .MouseCol <> meSelectedGridCell.Col Then
                .Redraw = False
                .ColPosition(Val(flxVisits.Tag)) = .MouseCol
                .Redraw = True
                ReorderVisits
            '   15/10/98
            '   Changed row test to be after the mROW_VISITTASKID
            ' NCJ 16 Apr 07 - Check that they actually moved the mouse to a new cell
            ElseIf .Tag <> gsEMPTY_STRING _
            And MousePointer = vbSizeNS _
            And .MouseRow > mROW_VISITTASKID _
            And .MouseRow <> meSelectedGridCell.Row Then
                .Redraw = False
                .RowPosition(Val(.Tag)) = .MouseRow
                .Redraw = True
                ReorderCRFPages
                'If frmCRFDesign is open then change the form tabs order
                For Each oForm In Forms
                    If oForm.Name = "frmCRFDesign" Then
                        oForm.RefreshCRF
                    End If
                Next
                MousePointer = vbNormal
            '   ATN 9/11/98
            '   Added additional clause to reset mousepointer to normal
            '   if the cursor is on one of the top (fixed) rows
            ElseIf MousePointer = vbSizeNS Then
                MousePointer = vbNormal
            End If
            
            ' PN 01/09/99
            ' NCJ 28 Sept 06 - This is now done earlier
'            meMouseButton = LeftClick
            
        Else
            ' PN 01/09/99
            ' the right mouse was clicked so ONLY menu options apply
            
            ' bring focus to the newly selected cell
            .Col = .MouseCol
            .Row = .MouseRow
            
            ' PN 28/09/99 set flag for click in cell
            If X > flxVisits.CellLeft + flxVisits.CellWidth Or _
                    Y > flxVisits.CellTop + flxVisits.CellHeight Then
            
                gbNotInCell = True
                flxVisits.FocusRect = flexFocusNone
            Else
                gbNotInCell = False
                flxVisits.FocusRect = flexFocusLight
            End If
            
            ' NCJ 28 Sept 06 - This is now done earlier
'            meMouseButton = RightClick
            Call ShowPopupMenu
            
        End If
        
    End With

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "flxVisits_MouseUp")
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
Private Sub Form_Activate()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    frmMenu.ChangeSelectedItem "", ""

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Form_Activate")
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
Private Sub Form_Load()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon
    
    ' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True
    
    Me.Top = 0
    Me.Left = (frmMenu.Width / 4)
    Me.Width = ((frmMenu.Width / 4) * 3) - 200
    Me.Height = frmMenu.Height - frmMenu.tlbMenu.Height * 2.25 - frmMenu.sbrMenu.Height
    
    flxVisits.Top = 50
    flxVisits.Left = 50
    flxVisits.Width = Me.Width - 200
    flxVisits.Height = Me.Height - 500
    Me.Caption = Me.ClinicalTrialName & " schedule"
    
    mnCurrentCRFPageTaskId = 0
    
    'ASH 4/11/2002
    txtCycleVisit.MaxLength = Len(msUNLIMITED_VISIT_REPEATS)
    
    ' NCJ 20 Jan 00
    'MLM 07/01/03: Don't load images from the resource file, as the image list is populated at design time.
    'Call PopulateImageList
    
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
  
  On Error GoTo ErrHandler
    
    If KeyCode = vbKeyF1 Then               ' Show user guide
        'ShowDocument Me.hWnd, gsMACROUserGuidePath
        
        'REM 07/12/01 - New Call to MACRO Help
        Call MACROHelp(Me.hWnd, App.Title)
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Form_KeyDown")
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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---------------------------------------------------------------------
'ZA 11/09/2002 - save the visit name/cycle data if lost focus was not
'fired because form was closed
'---------------------------------------------------------------------
     
     SaveVisitCycle False
     SaveVisitName False
     
End Sub

'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Me.Height > 500 Then
        flxVisits.Top = 50
        flxVisits.Left = 50
        flxVisits.Width = Me.Width - 200
        flxVisits.Height = Me.Height - 500
    End If
    

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Form_Resize")
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
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------------------
'This sub added by Mo Morris 10/12/98 (SPR 556 & 658)
'---------------------------------------------------------------------
    ' Forget any errors during the unload - NCJ 9 Nov 99
    ' On Error GoTo ErrHandler
    
    frmMenu.HideVisits
    
End Sub

'---------------------------------------------------------------------
Private Sub tmrGridClick_Timer()
'---------------------------------------------------------------------
' timer event ensures the handling of only one mouse action in the case
' of a dblclick
'---------------------------------------------------------------------
    
    tmrGridClick.Enabled = False
    
    If meMouseAction = DoubleClick Then
        Call GridDblClick
    ElseIf meMouseAction = SingleClick Then
        Call GridClick
    End If
    
    Screen.MousePointer = vbDefault

End Sub

'---------------------------------------------------------------------
Private Sub txtCycleVisit_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
' NCJ 31 Jan 03 - Accept Return key
'---------------------------------------------------------------------

    If KeyCode = vbKeyReturn Then
        Call SaveVisitCycle(True)
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub txtCycleVisit_LostFocus()
'---------------------------------------------------------------------
'ZA 11/09/2002 - calls SaveVisitCycle
'---------------------------------------------------------------------
    
    SaveVisitCycle True

End Sub

'---------------------------------------------------------------------
Private Sub txtVisitName_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
' NCJ 31 Jan 03 - Accept Return key
'---------------------------------------------------------------------

    If KeyCode = vbKeyReturn Then
        Call SaveVisitName(True)
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub txtVisitName_LostFocus()
'---------------------------------------------------------------------
'ZA 11/09/2002 - call SaveVisitName routine
'---------------------------------------------------------------------
    
    SaveVisitName True

End Sub

'---------------------------------------------------------------------
Private Sub SetColumnWidth(iColumn As Integer)
'---------------------------------------------------------------------
' function to handle logic for setting a column width
' called when column data is edited and when grid is loaded
'---------------------------------------------------------------------
    
  Dim lMaxColWidth As Long
    On Error GoTo ErrHandler
    ' PN 28/09/99 specify a minimum width for the columns
    lMaxColWidth = 1200
    With flxVisits
        ' get the column width by reading the width of cells
        ' mROW_VISITCYCLE and mROW_VISITNAME
        ' the max value dictates the column width
        If GetTextWidth(.TextMatrix(mROW_VISITNAME, iColumn)) > lMaxColWidth Then
            ' Bug fix - NCJ 7 Dec 99
            lMaxColWidth = GetTextWidth(.TextMatrix(mROW_VISITNAME, iColumn))
            'lMaxColWidth = GetTextWidth(.TextMatrix(mROW_VISITCODE, iColumn))
        End If
        If GetTextWidth(.TextMatrix(mROW_VISITCODE, iColumn)) > lMaxColWidth Then
            lMaxColWidth = GetTextWidth(.TextMatrix(mROW_VISITCODE, iColumn))
        End If
        'ZA 06/09/2002 - visit cycle column
        If GetTextWidth(.TextMatrix(mROW_VISITCYCLE, iColumn)) > lMaxColWidth Then
            lMaxColWidth = GetTextWidth(.TextMatrix(mROW_VISITCYCLE, iColumn))
        End If
        If .ColWidth(iColumn) > 0 Then
            .ColWidth(iColumn) = lMaxColWidth
        End If
    
    End With

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "SetColumnWidth")
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
Private Function GetTextWidth(sText As String) As Long
'---------------------------------------------------------------------
' obtain the width of the text fro the form control
' this works since the text box and form hasve the same font properties
'---------------------------------------------------------------------
    
  On Error GoTo ErrHandler
    
    GetTextWidth = TextWidth(String(Len(sText) + 4, "_"))

Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GetTextWidth")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

'---------------------------------------------------------------------
Public Sub RefreshStudyVisits()
'---------------------------------------------------------------------
    
    Call BuildStudyVisits

End Sub

'---------------------------------------------------------------------
Public Sub BuildStudyVisits()
'---------------------------------------------------------------------
' NCJ 31/8/99 - Added two extra optional arguments for use by DM
'---------------------------------------------------------------------

    On Error Resume Next

    With flxVisits
        'ZA 28/08/2002 - reduced to 6
        .Rows = 6
        .Cols = 5
        'ZA 28/08/2002 - reduced to 5 as date prompt row is not required
        .FixedRows = 5
        .FixedCols = 4
        '.TextMatrix(5, 4) = "Visit cycles"
    
        .Visible = False
    
        '   SPR 578 ATN 9/11/98
        '   Disable user resizing of the grid
        .AllowUserResizing = flexResizeNone
    
        .MergeCells = flexMergeRestrictAll
        .MergeCol(mnCRFPageIdColumn) = True
        .MergeCol(mnCRFPageOrderColumn) = True
        .MergeCol(mnCRFPageCodeColumn) = True
        .MergeCol(mnCRFPageNameColumn) = True
        .MergeRow(mROW_VISITID) = True
        .MergeRow(mROW_VISITORDER) = True
        .MergeRow(mROW_VISITCODE) = True
        .MergeRow(mROW_VISITNAME) = True
        .ColWidth(mnCRFPageIdColumn) = 0
        .ColWidth(mnCRFPageOrderColumn) = 0
        .ColWidth(mnCRFPageCycleNumberColumn) = 0
        .ColWidth(mnCRFPageCodeColumn) = 0
        .RowHeight(mROW_VISITID) = 0
        .RowHeight(mROW_VISITORDER) = 0
        .RowHeight(mROW_VISITTASKID) = 0
    '    .RowHeight(mROW_VISITCYCLENUMBER) = 0
        .RowHeight(mROW_VISITCODE) = 375
        .ColWidth(mnCRFPageNameColumn) = 1500
        .RowHeight(mROW_VISITNAME) = 375
        .FocusRect = flexFocusLight
        .WordWrap = True
    
        .RowHeight(mROW_VISITCYCLE) = 350
        .BackColor = vbWhite
        .GridLines = flexGridNone
        
        RefreshRows
        
        RefreshColumns
        
        RefreshCells
        
        .Visible = True
    End With

    If frmMenu.StudyAccessMode < sdReadWrite Then
        ' Disable stuff
    End If
    
    'Changed by Mo Morris 8/1/99 SR 631
    'Me.Show
End Sub

'---------------------------------------------------------------------
Private Sub RefreshColumns()
'---------------------------------------------------------------------
    
Dim sSQL As String
Dim rsStudyVisits As ADODB.Recordset
Dim bEditVisitDate As Boolean

    On Error GoTo ErrHandler
    
    '   13/10/98    Visit date added
    sSQL = "SELECT  StudyVisit.*,NULL as VisitCycleNumber,NULL as VisitDate FROM StudyVisit " _
        & " WHERE ClinicalTrialId   = " & ClinicalTrialId _
        & " AND   VersionId         = " & VersionId _
        & " ORDER BY StudyVisit.VisitOrder"
    
    Set rsStudyVisits = New ADODB.Recordset
    rsStudyVisits.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With flxVisits
        While Not rsStudyVisits.EOF
            If IsNull(rsStudyVisits!VisitCycleNumber) Then
                .Cols = .Cols + 1
                .Col = .Cols - 1
                .ColWidth(.Col) = 1000
            Else
                If rsStudyVisits!VisitCycleNumber > 1 Then
                    .Cols = .Cols + 1
                    .Col = .Cols - 1
                    .ColWidth(.Col) = 1000
                End If
            End If
        
            .Row = mROW_VISITID
            .Text = rsStudyVisits!VisitId
            .CellAlignment = flexAlignCenterCenter
            .CellPictureAlignment = flexAlignCenterCenter
        
            .Row = mROW_VISITORDER
            .Text = rsStudyVisits!VisitOrder
            .CellAlignment = flexAlignCenterCenter
            .CellPictureAlignment = flexAlignCenterCenter
        
            .Row = mROW_VISITCODE
            .Text = rsStudyVisits!VisitCode
            .CellAlignment = flexAlignCenterCenter
            .CellPictureAlignment = flexAlignCenterCenter
        
            .Row = mROW_VISITNAME
            .Text = rsStudyVisits!VisitName
            .CellAlignment = flexAlignCenterCenter
            .CellPictureAlignment = flexAlignCenterCenter
            .CellBackColor = .BackColorFixed
            
           '  new Visit cycles row
            .Row = mROW_VISITCYCLE
            
            If Not IsNull(rsStudyVisits!Repeating) Then
                'ZA 06/09/2002 - put unlimited if the value is -1
                If rsStudyVisits.Fields("Repeating").Value = mnUNLIMITED_VISIT_VALUE Then
                    .Text = msUNLIMITED_VISIT_REPEATS
                Else
                    .Text = rsStudyVisits!Repeating
                End If
            Else
                .Text = mSVISIT_CYCLE
            End If
            
            Call SetColumnWidth(.Col)
            
            .CellAlignment = flexAlignCenterCenter
            .CellPictureAlignment = flexAlignCenterCenter
            .CellBackColor = .BackColorFixed
            
            .Row = mROW_VISITCYCLENUMBER
            .Text = msEMPTYCELL
            .CellAlignment = flexAlignCenterCenter
            .CellPictureAlignment = flexAlignCenterCenter
            
            ' PN 02/09/99
            ' set this column's width
            Call SetColumnWidth(.Col)
            
            ' PN 06/09/99
            ' set this column's colour
            Call SetBackgroungColour(rsStudyVisits.Fields("VisitBackgroundColour"), False)
            
            rsStudyVisits.MoveNext
        Wend
        
    End With
    rsStudyVisits.Close
    Set rsStudyVisits = Nothing

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshColumns")
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
Private Sub RefreshRows()
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsCRFPages As ADODB.Recordset

    On Error GoTo ErrHandler
  
    sSQL = "SELECT crfpage.*, NULL as CRFPageCycleNumber " _
        & " FROM CRFPage " _
        & " WHERE   ClinicalTrialId     =  " & ClinicalTrialId _
        & " AND     VersionId           =  " & VersionId _
        & " ORDER BY CRFPage.CRFPageOrder"
    
    Set rsCRFPages = New ADODB.Recordset
    rsCRFPages.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With flxVisits
        While Not rsCRFPages.EOF
            If IsNull(rsCRFPages!CRFPageCycleNumber) Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .RowHeight(.Row) = 750
            Else
                If rsCRFPages!CRFPageCycleNumber > 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .RowHeight(.Row) = 750
                End If
            End If
            
            .Col = mnCRFPageIdColumn
            .Text = rsCRFPages!CRFPageId
            .Col = mnCRFPageOrderColumn
            .Text = rsCRFPages!CRFPageOrder
            .Col = mnCRFPageCodeColumn
            .Text = rsCRFPages!CRFPageCode
            .Col = mnCRFPageNameColumn
        '    .Text = rsCRFPages!CRFPageCode & " : " & rsCRFPages!CRFTitle
            .Text = rsCRFPages!CRFTitle
            .CellAlignment = flexAlignLeftCenter
        
            .Col = mnCRFPageCycleNumberColumn
            .Text = msEMPTYCELL
            .CellAlignment = flexAlignLeftCenter
        
            
            rsCRFPages.MoveNext
        Wend
    
    End With
    rsCRFPages.Close
    Set rsCRFPages = Nothing

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshRows")
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
Private Sub RefreshCells()
'---------------------------------------------------------------------
' Refresh cells in schedule
' NCJ 31/8/99 - Added two extra optional arguments for use by DM
' NCJ 27/9/99 - Removed two extra optional arguments (only used by DM)
'               Changed sub to Private (only used in this module)
'---------------------------------------------------------------------
    
Dim sSQL As String
Dim rsInstances As ADODB.Recordset
Dim nCol As Integer
Dim nRow As Integer
  On Error GoTo ErrHandler
'   SR 580  24/11/98    ATN
'   Initialise location of first available form
 mnFirstFormCol = 32000
 mnFirstFormRow = 32000

With flxVisits
    For nCol = mnCRFPageCycleNumberColumn + 1 To .Cols - 1
        .Col = nCol
        sSQL = "SELECT StudyVisitCRFPage.* " _
            & " FROM StudyVisitCRFPage " _
            & " WHERE StudyVisitCRFPage.ClinicalTrialId = " & ClinicalTrialId _
            & " AND StudyVisitCRFPage.VersionId = " & VersionId _
            & " AND VisitId = " & .TextMatrix(mROW_VISITID, nCol)
            
        Set rsInstances = New ADODB.Recordset
        rsInstances.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
        While Not rsInstances.EOF
            For nRow = mROW_VISITCYCLENUMBER + 1 To .Rows - 1
                .Row = nRow
                If CStr(.TextMatrix(nRow, mnCRFPageIdColumn)) = CStr(rsInstances!CRFPageId) Then
                    'TA 24/08/2002: check to se if it is the visit eform (for visit date)
                    If rsInstances!eFormUse = eEFormUse.VisitEForm Then
                        .Text = msVISIT_EFORMCELL
                        Set flxVisits.CellPicture = imgLargeIcons.ListImages(gsVISIT_EFORM).Picture
                    Else
                        If rsInstances!Repeating = 0 Then
                            .Text = msSINGLECELL
                            Set .CellPicture = imgLargeIcons.ListImages(gsLARGE_CRF_PAGE_LABEL).Picture
                            ' Set .CellPicture = LoadResPicture(gsLARGE_CRF_PAGE_LABEL, vbResIcon)
                        Else
                            .Text = msREPEATINGCELL
                             Set .CellPicture = imgLargeIcons.ListImages(gsREPEATING_CRF_PAGE_LABEL).Picture
                           ' Set .CellPicture = LoadResPicture(gsREPEATING_CRF_PAGE_LABEL, vbResIcon)
                        End If
                    End If
                    .CellAlignment = flexAlignCenterCenter
                    .CellPictureAlignment = flexAlignCenterCenter
                '   SR 580  24/11/98    ATN
                '   If not already set, set the location of the first available form
                    If .Col <= mnFirstFormCol And .Row < mnFirstFormRow Then
                         mnFirstFormCol = .Col
                         mnFirstFormRow = .Row
                    End If
                    Exit For
                End If

            Next nRow
            
            rsInstances.MoveNext
        Wend
    Next
    'ASH 17/12/2002 Give row name in schedule
    .TextMatrix(4, 3) = "Visit Cycles"
End With

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshCells")
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
Public Sub SetBackgroungColour(lColour As Long, Optional bSave As Boolean = True)
'---------------------------------------------------------------------
    
Dim sSQL As String
Dim lIndex As Long
Dim lColIndex As Long
Dim eSelectedCell As SelectedGridCell
  On Error GoTo ErrHandler
    ' set the background colour to the value passed in
    With flxVisits
        ' first save the selecetd cell
        eSelectedCell.Row = .Row
        eSelectedCell.Col = .Col
        
        ' then copy the colour in
        'SDM SR2104 23/11/99
        'For lIndex = 1 To .Rows - 1
        For lIndex = mROW_VISITCYCLENUMBER To .Rows - 1
            .Row = lIndex
            .CellBackColor = lColour
        Next lIndex
        
        ' then reset the selected cell
        .Row = eSelectedCell.Row
        .Col = eSelectedCell.Col
        
        If bSave Then
            ' then save the colour change
            sSQL = "UPDATE StudyVisit "
            sSQL = sSQL & " SET VisitBackgroundColour = " & lColour
            sSQL = sSQL & " WHERE ClinicalTrialId = " & ClinicalTrialId
            sSQL = sSQL & " AND VersionId = " & VersionId
            sSQL = sSQL & " AND VisitId = " & .TextMatrix(mROW_VISITID, .Col)
    
            MacroADODBConnection.Execute sSQL
            
        End If
        
    End With

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "SetBackgroungColour")
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
Private Sub PopulateImageList()
'---------------------------------------------------------------------
' NCJ 20 Jan 00
' Read the images into the imagelist
'   NOTE:   Taking images from the resouce file everytime instead of
'           using an image list will eat up memory.
'
' MLM 07/01/03: NB: This is no longer used as the icons have been added to the ImageList at
'   design time. I've haven't deleted this sub because it might be useful again, e.g. for
'   populating the ImageList from a 'master' ImageList on another form.
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    With imgLargeIcons
        .ListImages.Add , gsLARGE_CRF_PAGE_LABEL, LoadResPicture(gsLARGE_CRF_PAGE_LABEL, vbResIcon)
        .ListImages.Add , gsREPEATING_CRF_PAGE_LABEL, LoadResPicture(gsREPEATING_CRF_PAGE_LABEL, vbResIcon)
        .ListImages.Add , gsVISIT_EFORM, LoadResPicture(gsVISIT_EFORM, vbResIcon)
    End With

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PopulateImageList")
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
Private Sub SaveVisitCycle(ByVal bForLostFocusEvent As Boolean)
'---------------------------------------------------------------------
' ZA 09/09/2002 - saves Visit cycle in StudyVisit table
' True - if it is called by LostFocus event otherwise false
' save the visit cycle label data if lost focus event was fired
' or if this form is being closed
' ASH 4/11/2002 Changed nRepeatValue to lRepeatValue
' REM 17/08/04 - check to see if active control is nothing
' NCJ 16 Apr 07 - Set dirty flag for MUSD when visit cycles value changes; and only save if changed
'---------------------------------------------------------------------
Dim sSQL As String
Dim bSave As Boolean
Dim lRepeatValue As Long 'ash 4/11/2002
Dim nVisitColumn As Integer
Dim sRepeatValue As String

    On Error GoTo ErrHandler
    
    'if lost focus event is not fired and user was not editing vist cycle
    'then don't need to process any further
    'REM 17/08/04 - check to see if active control is nothing
    If Not bForLostFocusEvent And Me.ActiveControl Is Nothing Then Exit Sub
    If Not bForLostFocusEvent And Me.ActiveControl.Name <> "txtCycleVisit" Then Exit Sub
    
    bSave = False
       
    sRepeatValue = Trim(txtCycleVisit.Text)
    
   'ZA 06/09/2002 - if we have -1 or unlimited, then save the value as -1
    If LCase(sRepeatValue) = LCase(msUNLIMITED_VISIT_REPEATS) Or (Val(sRepeatValue) = mnUNLIMITED_VISIT_VALUE) Then
        lRepeatValue = mnUNLIMITED_VISIT_VALUE
        sRepeatValue = msUNLIMITED_VISIT_REPEATS
        txtCycleVisit.Visible = False
        bSave = True
    ElseIf sRepeatValue = "" Then
    'if no repeat then default to 1
        lRepeatValue = 1
        sRepeatValue = 1
        txtCycleVisit.Visible = False
        bSave = True
    Else
        lRepeatValue = Val(sRepeatValue)
        'check the range between 1 to 999 valid integer
        If (lRepeatValue < 1 Or lRepeatValue > 999) Or Not gblnValidString(sRepeatValue, valNumeric) Then
            'only display warning dialog if this is called by txtVisitCycle_lostfocus event
            If bForLostFocusEvent Then
                Call DialogWarning("Number of repeats must be between 1 and 999, or -1 for unlimited repeats")
                txtCycleVisit.SetFocus
                Exit Sub
            End If
        Else
            'value entered is a valid value
            txtCycleVisit.Visible = False
            bSave = True
        End If
    End If
   
   
    'get the current column
    nVisitColumn = Val(txtCycleVisit.Tag)
      
    ' NCJ 16 Apr 07 - Only save if cycle value has changed
    If bSave And flxVisits.TextMatrix(mROW_VISITCYCLE, nVisitColumn) <> sRepeatValue Then
        If bForLostFocusEvent Then
            'display visit cycle value if this function is fired by lost focus event of txtVisitCycle
            flxVisits.TextMatrix(mROW_VISITCYCLE, nVisitColumn) = sRepeatValue
        End If

        sSQL = "UPDATE StudyVisit " _
           & " SET Repeating = " & lRepeatValue _
            & " WHERE ClinicalTrialId = " & ClinicalTrialId _
            & " AND VersionId = " & VersionId _
            & " AND VisitId = " & flxVisits.TextMatrix(mROW_VISITID, nVisitColumn)

        MacroADODBConnection.Execute sSQL
        
        'update AREZZO visit repeats
        Call SetVisitRepeats(flxVisits.TextMatrix(mROW_VISITID, nVisitColumn), lRepeatValue)
        
        ' PN 02/09/99
        ' now set the column width correctly
        Call SetColumnWidth(flxVisits.Col)
        
        ' NCJ 16 Apr 07 - Set dirty flag for MUSD
        Call frmMenu.MarkStudyAsChanged
    End If
   
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SaveVisitCycle()")
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
Private Sub SaveVisitName(ByVal bForLostFocusEvent As Boolean)
'---------------------------------------------------------------------
' ZA 28/08/2002 - saves Visit name in StudyVisit table
' True - if it is called by LostFocus event otherwise false
' save the visit Name label data if lost focus event was fired
' or if this form is being closed
' REM 17/08/04 - check to see if active control is nothing
' NCJ 16 Apr 07 - Set dirty flag for MUSD when visit name changes; and only save if changed
'---------------------------------------------------------------------
Dim sSQL As String
Dim bSave As Boolean
Dim nVisitNameColumn As Integer
Dim sVisitName As String


    On Error GoTo ErrHandler
    
    'if lost focus event is not fired and user was not editing vist cycle
    'then don't need to process any further
    'REM 17/08/04 - check to see if active control is nothing
    If Not bForLostFocusEvent And Me.ActiveControl Is Nothing Then Exit Sub
    If Not bForLostFocusEvent And Me.ActiveControl.Name <> "txtVisitName" Then Exit Sub
        
    sVisitName = Trim(txtVisitName.Text)
    bSave = False
    
    'display warning dialog only if the call came from LostFocus event
    If bForLostFocusEvent Then
        'if cancel, then return control to user
        If sVisitName = gsEMPTY_STRING Then
            txtVisitName.Visible = False
            Exit Sub
        'show warning message if user entered invalied characters
        ElseIf Not gblnValidString(sVisitName, valOnlySingleQuotes) Then
            Call DialogWarning("Visit names" & gsCANNOT_CONTAIN_INVALID_CHARS)
            txtVisitName.SetFocus
            Exit Sub
        'show warning message if more than 50 characters
        ElseIf Len(sVisitName) > 50 Then
            Call DialogWarning("Visit names cannot be more than 50 characters long")
            txtVisitName.SetFocus
            Exit Sub
        Else
            'every thing seems OK
            txtVisitName.Visible = False
            bSave = True
        End If
    Else
        'a call came when this form is being closed
        If sVisitName <> gsEMPTY_STRING Then
            If Not gblnValidString(sVisitName, valOnlySingleQuotes) Or Len(sVisitName) > 50 Then
            Else
                bSave = True
            End If
        Else
            'exit this routine if no value entered by user
            Exit Sub
        End If
    End If
    
    'get the column where user was editing visit name
    nVisitNameColumn = Val(txtVisitName.Tag)
    'save this visit name only if there is a change
    ' NCJ 16 Apr 07 - Check it has really changed
    If bSave And flxVisits.TextMatrix(mROW_VISITNAME, nVisitNameColumn) <> sVisitName Then
        If bForLostFocusEvent Then
            flxVisits.TextMatrix(mROW_VISITNAME, nVisitNameColumn) = sVisitName
        End If
        ' NCJ 3 Dec 99 - Use single quotes in SQL string
        sSQL = "UPDATE StudyVisit " _
            & " SET VisitName = '" & ReplaceQuotes(sVisitName) & "' " _
            & " WHERE ClinicalTrialId = " & ClinicalTrialId _
            & " AND VersionId = " & VersionId _
            & " AND VisitId = " & flxVisits.TextMatrix(mROW_VISITID, nVisitNameColumn)
            
            MacroADODBConnection.Execute sSQL
            
            ' PN 02/09/99
            ' now set the column width correctly
            Call SetColumnWidth(flxVisits.Col)
            
            ' NCJ 16 Apr 07 - Set dirty flag for MUSD
            Call frmMenu.MarkStudyAsChanged
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SaveVisitName()")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'-----------------------------------------------------------------------------
Private Function IsValidationFormAlreadyThere(ByVal nCol As Integer) As Boolean
'-----------------------------------------------------------------------------
'ZA 23/08/2002 - check if a visit contains a date validation form
'-----------------------------------------------------------------------------
Dim i As Integer

    IsValidationFormAlreadyThere = False
    
    For i = mROW_VISITTASKID To flxVisits.Rows - 1
        
        If flxVisits.TextMatrix(i, nCol) = msVISIT_EFORMCELL Then
            IsValidationFormAlreadyThere = True
            mnVisitFormRow = i
            Exit Function
        End If
    Next
    
End Function
