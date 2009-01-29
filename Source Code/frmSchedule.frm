VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSchedule 
   BorderStyle     =   0  'None
   Caption         =   "Schedule"
   ClientHeight    =   4500
   ClientLeft      =   5910
   ClientTop       =   4965
   ClientWidth     =   7125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTitleBar 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   7095
      TabIndex        =   2
      Top             =   15
      Width           =   7095
   End
   Begin MSComctlLib.ImageList imglstStatus 
      Left            =   4680
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":0000
            Key             =   "DM30_RaisedDisc"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":00CF
            Key             =   "DM30_Warning"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":0167
            Key             =   "DM30_NA"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":04DF
            Key             =   "DM30_InactiveForm"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":087E
            Key             =   "DM30_NewForm"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":0930
            Key             =   "DM30_RespondedDisc"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":09FC
            Key             =   "DM30_Frozen"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":0A93
            Key             =   "DM30_Inform"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":0E0A
            Key             =   "DM30_Locked"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":0EBD
            Key             =   "DM30_Missing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":124D
            Key             =   "DM30_OK"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":12E4
            Key             =   "DM30_OKWarning"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":1391
            Key             =   "DM30_Unobtainable"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxSchedule 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   6588
      _Version        =   393216
      BorderStyle     =   0
   End
   Begin VB.Shape shpSDVStatus 
      Height          =   495
      Index           =   0
      Left            =   5700
      Top             =   3540
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "this invisible label is used to calculate cell heights"
      Height          =   345
      Left            =   60
      TabIndex        =   1
      Top             =   3900
      Visible         =   0   'False
      Width           =   4455
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------
' File: frmSchedule.frm
' Copyright: InferMed 2001, All Rights Reserved
' Author: Nicky Johns/Toby Aldridge, InferMed, May 2001
' Form to display Subject Schedule classes for MACRO 2.2
'----------------------------------------------------
' REVISIONS
'   NCJ 30/8/01 - Some tidying up
'   DPH 12/10/2001 - CanOpenEform & call to it in flxSchedule_DblClick to
'       control opening new forms
'   DPH 24/10/2001 - flxSchedule_MouseMove edited to correct tooltip problem
'   NCJ 20 Mar 02 - Ensure visit/form dates only editable when appropriate
'   NCJ 15 May 02 - Conditionally compiled code for Multi User support
'   za 22/05/02, only call RedrawGrid if there is a Subject
'   TA 11/07/02: CBB 2.2.19.12 Allow changing of status by right-clicking an eForm in the schedule
'   NCJ 14 Aug 02 - Changed ToggleEFIStatus to use new Response status routine
'   MLM 30/08/02: Don't check for a visit date in flxSchedule_DblClick()
'   NCJ 18 Sept 02 - frmEFormDataEntry.Display may fail if we failed to get Responses
'   NCJ 20 Sept 02 - Added extra arg. to RemoveResponses
'   TA 26/09/02: Changes for New UI - no title bar, not maximised etc
'   NCJ 26 Sept 02 - Removed EditDate & associated stuff; added check for subject updates when double clicking
'   NCJ 30 Sept 02 - Fixed bug in MUScheduleUpdated
'----------------------------------------------------

Option Explicit

'hold reference to eform to catch events
Private WithEvents moEFormDataEntry As Form
Attribute moEFormDataEntry.VB_VarHelpID = -1

Private moSubject As StudySubject
Private moUser As MACROUserBS30.User
Private mlTotalColumnsWidth As Long
Private mlTotalRowsHeight As Long
Private Const m_DEFAULT_COLWIDTH = 1300
Private Const m_DEFAULT_ROWHEIGHT = 700
Private Const m_COL_PADDING = 240
Private mnCol As Integer
Private mnRow As Integer

'----------------------------------------------------------------
Public Function Display(oUSer As MACROUserBS30.User, oStudyDef As StudyDefRO)
'----------------------------------------------------------------
' Display an schedule for a preloaded subject
'----------------------------------------------------------------

    Load Me
    
    Me.BackColor = eMACROColour.emcBackGround
    picTitleBar.Height = eMACROLength.emlTitleBarHeight
    picTitleBar.BackColor = eMACROColour.emcTitlebar
    picTitleBar.Left = 120
    
    flxSchedule.Top = picTitleBar.Height
    flxSchedule.Left = 120
    
    flxSchedule.BackColorBkg = eMACROColour.emcBackGround
    
'    Me.WindowState = vbMaximized
       
    'create  a reference to the subject
    Set moSubject = oStudyDef.Subject
    Set moUser = oUSer
    
    Me.Caption = "MACRO Schedule [" & oStudyDef.Name & "] [" & oStudyDef.Subject.Label & "]"
    
    RefreshGrid
    
    Me.Show vbModeless
    
End Function

'--------------------------------------------
Private Sub RefreshGrid()
'--------------------------------------------
' Refresh the schedule grid display.
' ReUse RedrewGrid if the form is visible
'--------------------------------------------
Dim os As ScheduleGrid
Dim oCell As GridCell
Dim lRow As Long
Dim lCol As Long
Dim lGridRow As Long 'flexgrid row not the same as row (two header rows)
Dim s As String
Dim OSV As ScheduleVisit
Dim sEFIText As String
Dim lRowHeight As Long
Dim sVisitDate As String
Dim sVisitName As String
Dim oControl As Control

    On Error GoTo ErrLabel
    
    HourglassOn

'    'unload any previously loaded shapes
'    For Each oControl In Me.Controls
'        If TypeOf oControl Is Shape Then
'            If oControl.Index <> 0 Then
'                Unload oControl
'            End If
'        End If
'    Next

    Set os = moSubject.ScheduleGrid
    
    flxSchedule.Visible = False

    flxSchedule.GridLines = flexGridNone
    flxSchedule.Clear
    flxSchedule.FixedCols = 1
    flxSchedule.FixedRows = 1
    flxSchedule.Cols = os.ColMax + 1
    flxSchedule.Rows = os.RowMax + 1 + 1 ' add one for visit date
                
    flxSchedule.MergeCells = flexMergeRestrictAll
    flxSchedule.MergeCol(0) = True
                
    flxSchedule.ColWidth(0) = m_DEFAULT_COLWIDTH
    flxSchedule.RowHeight(0) = m_DEFAULT_ROWHEIGHT
    flxSchedule.RowHeight(1) = 300
       
    'put in visit and visit date row headers
    flxSchedule.Col = 0
    flxSchedule.Row = 0
    flxSchedule.CellAlignment = flexAlignRightCenter
    flxSchedule.Text = "Visit  "
    flxSchedule.Row = 1
    flxSchedule.CellAlignment = flexAlignRightCenter
    flxSchedule.Text = "Visit Date   "

    'put in visit names
    For lCol = 1 To os.ColMax
        flxSchedule.ColWidth(lCol) = m_DEFAULT_COLWIDTH
        flxSchedule.Col = lCol
        flxSchedule.Row = 0
        flxSchedule.CellAlignment = flexAlignCenterCenter
        
        Set oCell = os.Cells(0, lCol)
        
        'set visit name
'        flxSchedule.Text = oCell.Visit.Name
        sVisitName = oCell.Visit.Name
        
        ' check to see if there is a corresponding visit instance and get the VisitDateString accordingly
        ' NCJ 17 Jan 02 - Also include cycle number in name if > 1 (Current Buglist 2.2.3 Bug 24)
        If oCell.VisitInst Is Nothing Then
            sVisitDate = ""
        Else
            sVisitDate = oCell.VisitInst.VisitDateString
            If oCell.VisitInst.CycleNo > 1 Then
                sVisitName = sVisitName & " [" & oCell.VisitInst.CycleNo & "]"
            End If
        End If
        
        flxSchedule.Text = sVisitName
        
        'set the width
        If (Me.TextWidth(sVisitName) + m_COL_PADDING > m_DEFAULT_COLWIDTH) Or (Me.TextWidth(sVisitDate) + m_COL_PADDING > m_DEFAULT_COLWIDTH) Then
            'visit name or visit date greater than the default
            If Me.TextWidth(sVisitDate) > Me.TextWidth(sVisitName) Then
                'visit date longer
                flxSchedule.ColWidth(lCol) = Me.TextWidth(sVisitDate) + m_COL_PADDING
            Else
                'visit name longer
                flxSchedule.ColWidth(lCol) = Me.TextWidth(sVisitName) + m_COL_PADDING
            End If
        Else
            'visit name and date shorter than the default
            flxSchedule.ColWidth(lCol) = m_DEFAULT_COLWIDTH
        End If
        
        flxSchedule.Row = 1
        'set to grey
        flxSchedule.CellBackColor = vbButtonFace
        flxSchedule.CellAlignment = flexAlignCenterCenter
        'set visit date
        flxSchedule.Text = sVisitDate
         
    Next
    
    'do rows
    For lRow = 1 To os.RowMax
        lGridRow = lRow + 1 ' add 1 for the visit date taking up a row
        flxSchedule.RowHeight(lGridRow) = m_DEFAULT_ROWHEIGHT
        flxSchedule.Row = lGridRow ' visit date
        'first column is the eForm name
        flxSchedule.Col = 0
        flxSchedule.Text = os.Cells(lRow, 0).eForm.Name
        flxSchedule.WordWrap = True
        
        For lCol = 1 To os.ColMax
            flxSchedule.Col = lCol
            
            flxSchedule.CellAlignment = flexAlignCenterBottom
            flxSchedule.CellPictureAlignment = flexAlignCenterTop
            flxSchedule.WordWrap = True
            ' NCJ 31/8/01 - Initialise text here
            flxSchedule.Text = ""
            sEFIText = ""
            
            Set oCell = os.Cells(lRow, lCol)
            
            flxSchedule.CellBackColor = os.Cells(0, lCol).Visit.BackgroundColour
            Select Case oCell.CellType
            Case Blank
                'leave blank
            Case Inactive
                'set picture to a grey form (-100 is a dummy status value)
                Set flxSchedule.CellPicture = imglstStatus.ListImages(GetEFormIconByStatus(-100, 0, 0)).Picture
            Case Active
'                Load shpSDVStatus(oCell.eFormInst.EFormTaskId)
'                Call SetSDVStatusGraphics(flxSchedule, shpSDVStatus(oCell.eFormInst.EFormTaskId), oCell.eFormInst.SDVStatus)
                
                'set the picture according to status
                Set flxSchedule.CellPicture = imglstStatus.ListImages(GetEFormIconByStatus(oCell.eFormInst.Status, oCell.eFormInst.DiscrepancyStatus, oCell.eFormInst.LockStatus)).Picture
                'label
                If oCell.eFormInst.eFormLabel <> "" Then
                    sEFIText = oCell.eFormInst.eFormLabel & vbCrLf
                End If
                'date
                If oCell.eFormInst.eFormDateString <> "" Then
                   sEFIText = sEFIText & oCell.eFormInst.eFormDateString & vbCrLf
                End If
                
                If sEFIText <> "" Then
                    'there is eFormInstance text
                    flxSchedule.Text = sEFIText
                    'use lblSize to work out the height of the text when wrapped
                    lblSize.Caption = sEFIText
                    lblSize.Width = flxSchedule.ColWidth(lCol)
                    lRowHeight = m_DEFAULT_ROWHEIGHT + lblSize.Height - 100 'from version 2.1
                    If lRowHeight > flxSchedule.RowHeight(lGridRow) Then
                        flxSchedule.RowHeight(lGridRow) = lRowHeight
                    End If
                End If
            Case Else
                'should never happen
            End Select
        Next
    Next
    
    'set focus on top left
    flxSchedule.Row = 2
    flxSchedule.Col = 1
    
    flxSchedule.Visible = True
    
    Call CalcTotalColRowWidth
    HourglassOff
    Exit Sub
    
ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "RefreshGrid", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select

End Sub

'--------------------------------------------
Private Sub ReDrawGrid()
'--------------------------------------------
' Redraw the schedule grid with the flexgrid not visible
'--------------------------------------------

    flxSchedule.Redraw = False
    RefreshGrid
    flxSchedule.Redraw = True

End Sub

'--------------------------------------------
Private Sub flxSchedule_DblClick()
'--------------------------------------------
' Show the selected data entry form
'
' NCJ 15 May 02 - Support for Multi User mode (NOT for release)
' MLM 30/08/02: Removed check for visit date (this will now be done before saving the form)
'--------------------------------------------
Dim lVisit As Long
Dim lEForm As Long
Dim oEFI As EFormInstance
Dim oCell As GridCell
Dim dblVisitDate As Double
Dim sErrorMessage As String

    lVisit = flxSchedule.Col
    lEForm = flxSchedule.Row - 1 ' allow for visit date row
    
    If lEForm < 1 Or lVisit < 1 Then
        Exit Sub
    End If
    
    HourglassOn

    ' If the schedule's changed we have to start again
    If MUScheduleUpdated("Please reselect the eForm to be opened.") Then
        HourglassOff
        Exit Sub
    End If

    Set oCell = moSubject.ScheduleGrid.Cells(lEForm, lVisit)
    If oCell.CellType = Active Then
        Set oEFI = oCell.eFormInst
    End If
    
    If Not oEFI Is Nothing Then
        ' DPH 12/10/2001 - Check if can open Eform
        If CanOpenEform(oEFI, sErrorMessage) Then
            ' Display may fail if we failed to get Responses
            With Me
                If frmEFormDataEntry.Display(moUser, oEFI, .Left, .Top, .Width, .Height) Then
                    ' form is displayed
                    Set moEFormDataEntry = frmEFormDataEntry
                End If
            End With
        Else
            Call DialogError(sErrorMessage)
        End If
    End If
    
    HourglassOff
    
Exit Sub

End Sub
'--------------------------------------------------------------------------
Private Sub flxSchedule_KeyPress(KeyAscii As Integer)
'--------------------------------------------------------------------------
'Allow user to select a form for data entry by pressing the Return key
'--------------------------------------------------------------------------
 On Error GoTo ErrHandler

    ' NCJ 2/2/00 - Same as flxSchedule_DblClick
    If KeyAscii = 13 Then
    
        Call flxSchedule_DblClick
        
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "flxSchedule_KeyPress")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'----------------------------------------------------------------------------------------------
Private Sub flxSchedule_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'----------------------------------------------------------------------------------------------
'Display pop up menu for editing eForm or visit dates
' NCJ 20 Mar 02 - Do not allow editing of dates if visit/form is locked/frozen
' or user doesn't have change data rights
'   TA 11/07/02: CBB 2.2.19.12 Allow changing of status by right-clicking an eForm in the schedule
' NCJ 26 Sept 02 - No longer allow date editing here (removed various bits of code)
'----------------------------------------------------------------------------------------------
Dim oGrid As ScheduleGrid
Dim oCell As GridCell
Dim lVisit As Long
Dim lEForm As Long
Dim oEFI As EFormInstance
Dim sEnableStatus As String

    On Error GoTo ErrLabel
    
    flxSchedule.Row = mnRow
    flxSchedule.Col = mnCol
    
    lVisit = flxSchedule.Col
    lEForm = flxSchedule.Row - 1
    
    ' NCJ 4/10/01 - Make sure we've got valid values
    If lEForm < 0 Or lVisit < 1 Then Exit Sub
    
    Set oGrid = moSubject.ScheduleGrid
    Set oCell = moSubject.ScheduleGrid.Cells(lEForm, lVisit)
    
    If oCell.CellType = Active Then
        Set oEFI = oCell.eFormInst
    End If
    
    'check if the user has pressed right mouse button
    If Button = vbRightButton Then
        
        'if the row is 1 then it is visit date
        If flxSchedule.Row = 1 Then
            ' NCJ 26 Sept 02 - Nothing to do here now!
        Else
            ' It is an eForm
            If Not oEFI Is Nothing Then
                If Not CanMakeEFormUnobtainable(oEFI) Then
                    'not allowed to unobtainablise eForm
                    sEnableStatus = "*"
                End If
                'add separator
                sEnableStatus = sEnableStatus & "|"
                If Not CanMakeEFormMissing(oEFI) Then
                    'not allowed to make eForm Missing
                    sEnableStatus = sEnableStatus & "*"
                End If
                'show popup mennu
                Select Case frmMenu.ShowPopUp("Unobtainable|Missing", sEnableStatus)
                Case 1 'set eform to unobtainable
                    Call ToggleEFIStatus(oEFI, eStatus.Unobtainable)
                Case 2 'set eform to missing
                    Call ToggleEFIStatus(oEFI, eStatus.Missing)
                End Select
            End If
    
        End If
    End If
    
    Set oEFI = Nothing
    Set oCell = Nothing
    Set oGrid = Nothing
    Exit Sub
    
ErrLabel:
    'destroy all the objects since we are have got an error
    Set oEFI = Nothing
    Set oGrid = Nothing
    Set oCell = Nothing
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "flxSchedule_MouseUp", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
End Sub

'----------------------------------------------------------------------------------------------
Private Sub flxSchedule_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'----------------------------------------------------------------------------------------------
'TA 11/07/2002: all context sensitive popup menu processing now done in MouseUp event
'----------------------------------------------------------------------------------------------
Dim oGrid As ScheduleGrid
Dim oCell As GridCell
Dim lVisit As Long
Dim lEForm As Long
Dim oEFI As EFormInstance

    flxSchedule.Row = mnRow
    flxSchedule.Col = mnCol

'TA 11/07/2002: all context sensitive popup menu processing now done in MouseUp event

    Exit Sub

ErrLabel:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "flxSchedule_MouseDown", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
End Sub

'--------------------------------------------------------------
Private Function CanMakeEFormUnobtainable(oEFI As EFormInstance) As Boolean
'--------------------------------------------------------------
' Returns TRUE if current user can unobtainablise the given eForm instance
'--------------------------------------------------------------

    CanMakeEFormUnobtainable = False
    
    ' User must be able to change data
    If Not gUser.CheckFunctionAccess(gsFnChangeData) Then Exit Function
    
    ' Subject must not be Read-Only
    If moSubject.ReadOnly Then Exit Function
    
    ' Form must not be locked or frozen
    If oEFI.LockStatus <> eLockStatus.lsUnlocked Then Exit Function
    
    ' Form must have status of missing
    If oEFI.Status <> eStatus.Missing Then Exit Function
    
    ' If we get here, all is OK
    CanMakeEFormUnobtainable = True

End Function

'--------------------------------------------------------------
Private Function CanMakeEFormMissing(oEFI As EFormInstance) As Boolean
'--------------------------------------------------------------
' Returns TRUE if current user can make missing the given eForm instance
'--------------------------------------------------------------

    CanMakeEFormMissing = False
    
    ' User must be able to change data
    If Not gUser.CheckFunctionAccess(gsFnChangeData) Then Exit Function
    
    ' Subject must not be Read-Only
    If moSubject.ReadOnly Then Exit Function
    
    ' Form must not be locked or frozen
    If oEFI.LockStatus <> eLockStatus.lsUnlocked Then Exit Function
    
    ' Form must have status of unobtanable
    If oEFI.Status <> eStatus.Unobtainable Then Exit Function

    ' If we get here, all is OK
    CanMakeEFormMissing = True

End Function

'--------------------------------------------------------------
Private Sub ToggleEFIStatus(oEFI As EFormInstance, nToStatus As eStatus)
'--------------------------------------------------------------
'Make an eForm unobtainable by changing all its 'Missing' responses to 'Unobtainable'
' or vice versa
' no permissions etc are checked in here as they were checked to enable the menu items
' nb we change derived questions
' NCJ 14 Aug 02 - Use new SetStatusFromSchedule routine of Response object
'--------------------------------------------------------------
Dim oResponse As Response
Dim lChanged As Long
Dim nOldStatus As eStatus
Dim sStatus As String
Dim sMsg As String
Dim sLockErrMsg As String
Dim sEFILockToken As String
Dim sVEFILockToken As String

    On Error GoTo ErrLabel
    
    HourglassOn
    
    ' NCJ 26 Sept 02 - Check for subject data updates first
    ' and don't continue if there are any
    If MUScheduleUpdated("Please try again") Then
        HourglassOff
        Exit Sub
    End If
    
    'choose the statuses to look for
    If nToStatus = eStatus.Missing Then
        nOldStatus = eStatus.Unobtainable
    Else
        nOldStatus = eStatus.Missing
    End If
        
    'load efi's responses
    'we don't need to hold onto the EFILock Token or VEFILockToken
    If moSubject.LoadResponses(oEFI, sLockErrMsg, sEFILockToken, sVEFILockToken) <> lrrReadWrite Then
        DialogError sLockErrMsg
        HourglassOff
'EXIT SUB HERE
        Exit Sub
    End If
    
    lChanged = 0
    'loop through each response
    For Each oResponse In oEFI.Responses
        If oResponse.Status = nOldStatus And oResponse.LockStatus = eLockStatus.lsUnlocked Then
            'this response is unlocked so toggle this response's status. nb we change derived questions
'            oResponse.Status = nToStatus
            ' NCJ 14 Aug 02 - Use new call
            Call oResponse.SetStatusFromSchedule(nToStatus)
            'increase the count of responses changed
            lChanged = lChanged + 1
        End If
    Next

    sMsg = "Are you sure you wish to change " & lChanged & " " & GetStatusText((nOldStatus)) & " response"
    If lChanged > 1 Then
        'if more than one change, make plural
        sMsg = sMsg & "s"
    End If
    sMsg = sMsg & " on '" & oEFI.eForm.Name & "' to " & GetStatusText((nToStatus)) & "?"
    
    'Ask to confirm
    If DialogQuestion(sMsg) = vbYes Then
        'save the responses - this will save the subject
        Select Case moSubject.SaveResponses(oEFI, sLockErrMsg)
        Case srrNoLockForSaving
            DialogError sLockErrMsg
        Case srrSubjectReloaded
            ' we'll just try again...
            If moSubject.SaveResponses(oEFI, sLockErrMsg) <> srrSuccess Then
                DialogError "Unable to save changes because another user is editing this subject"
            End If
        Case srrSuccess
            ' OK
        End Select
        'ensure schedule is refreshed
        ReDrawGrid
    End If

    'remove the response from memory
    Call moSubject.RemoveResponses(oEFI, True)
    
    HourglassOff
        
Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|frmSchedule.ToggleEFIStatus"
    
End Sub

'----------------------------------------------------------------------------------------------
Private Sub flxSchedule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'----------------------------------------------------------------------------------------------
'Display tooltip text
' DPH 24/10/2001 - Moved save row/col info until after difference check to fix tooltip problem
'----------------------------------------------------------------------------------------------
Dim oEFI As EFormInstance
Dim oCell As GridCell
Dim lVisit As Long
Dim lEForm As Long
Dim nCount As Integer

    On Error GoTo ErrLabel
    
    ' DPH 24/10/2001 - Moved current row save until after same cell check
    'save current row and column for future inquiry
'    mnRow = flxSchedule.MouseRow
'    mnCol = flxSchedule.MouseCol

    'If the same cell exit
    If (mnRow = flxSchedule.MouseRow And mnCol = flxSchedule.MouseCol) And _
       (X < mlTotalColumnsWidth And Y < mlTotalRowsHeight) And _
       Len(flxSchedule.ToolTipText) > 0 Then
        Exit Sub
    End If

    'save current row and column for future inquiry
    mnRow = flxSchedule.MouseRow
    mnCol = flxSchedule.MouseCol

    'clear any previous tooltiptext value
    flxSchedule.ToolTipText = ""
    
    'exit if we are not on the active grid
    If mnRow <= 1 Or mnCol <= 0 Then Exit Sub
    
    lEForm = mnRow - 1 ' visit date
    lVisit = mnCol
    
    Set oCell = moSubject.ScheduleGrid.Cells(lEForm, lVisit)
    
     'Exit if not over the grid (MouseRow is never > Rows)
    If X > mlTotalColumnsWidth Or Y > mlTotalRowsHeight Then
        Exit Sub
    End If
       
   'display tooltip text depending upon the celltype
    Select Case oCell.CellType
        Case Active
            Set oEFI = oCell.eFormInst
            If Not oEFI Is Nothing Then
                flxSchedule.ToolTipText = "Status = " & oEFI.StatusString
                ' Only show lock status if not Unlocked
                If oEFI.LockStatus <> LockStatus.lsUnlocked Then
                    flxSchedule.ToolTipText = flxSchedule.ToolTipText & ", " & oEFI.LockStatusString
                End If
            End If
        Case Inactive
            flxSchedule.ToolTipText = "Inactive"
        Case Else
            flxSchedule.ToolTipText = ""
    End Select
    
    Set oEFI = Nothing
    Set oCell = Nothing
    Exit Sub
    
ErrLabel:
    'destroy all the objects since we are have got an error
    Set oEFI = Nothing
    Set oCell = Nothing
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "flxSchedule_MouseMove", Err.Source)
        Case OnErrorAction.Retry
            Resume
    End Select
End Sub


'--------------------------------------------
Private Sub Form_Resize()
'--------------------------------------------
'resizing code
'--------------------------------------------

    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.ScaleHeight < 1000 Then Exit Sub
    If Me.ScaleWidth < 1000 Then Exit Sub
    
    
    picTitleBar.Width = Me.ScaleWidth - 240
    With flxSchedule
        .Width = Me.ScaleWidth - 240
        .Height = Me.ScaleHeight - 240 - picTitleBar.Height
    End With
    

End Sub

'--------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'--------------------------------------------
'clear up
'--------------------------------------------
        
    If Not moSubject Is Nothing Then
        Set moSubject = Nothing
        Set moUser = Nothing
    End If

End Sub

'---------------------------------------------------------------------
Private Sub CalcTotalColRowWidth()
'---------------------------------------------------------------------
'   SDM 24/02/00 Calculation for use in flxVisits_MouseMove
'   Mo Morris 1/3/00, SR3122
'   NCJ 20/4/00 SR3346 - Changed nCols to nRows in second loop
'---------------------------------------------------------------------
Dim nCols As Integer
Dim nRows As Integer
    
    mlTotalColumnsWidth = 0
    mlTotalRowsHeight = 0
    For nCols = 0 To flxSchedule.Cols - 1
        mlTotalColumnsWidth = mlTotalColumnsWidth + flxSchedule.ColWidth(nCols)
    Next nCols
    For nRows = 0 To flxSchedule.Rows - 1
        mlTotalRowsHeight = mlTotalRowsHeight + flxSchedule.RowHeight(nRows)
    Next nRows
   
End Sub

'---------------------------------------------------------------------
Private Sub moEFormDataEntry_Unload(Cancel As Integer)
'---------------------------------------------------------------------
' Refresh schedule when eform unloaded.
'---------------------------------------------------------------------

    ' Lose reference as form is unloaded
    Set moEFormDataEntry = Nothing
    
    'za 22/05/02, only call RedrawGrid, if there is a Subject
    'As we exit from DM, frmSchedule can be unloaded before moFormDataEntry_unload event
    If Not moSubject Is Nothing Then
        ReDrawGrid
    End If
    
End Sub

'---------------------------------------------------------------------
Private Function CanOpenEform(oEFI As EFormInstance, ByRef sMessage As String) As Boolean
'---------------------------------------------------------------------
' Check if user can open eform they have clicked on
'---------------------------------------------------------------------
    
    CanOpenEform = False
    ' If cannot view data exit
    If Not moUser.CheckPermission(gsFnViewData) Then
        sMessage = "You do not have permission to view subject data"
        Exit Function
    End If
    ' If not a requested status eform then is OK and exit
    If oEFI.Status <> eStatus.requested Then
        CanOpenEform = True
        Exit Function
    End If
    ' Check permission to change data, subject not readonly
    If Not moUser.CheckPermission(gsFnChangeData) Or moSubject.ReadOnly Then
        sMessage = "You may not enter new data for this subject"
        Exit Function
    End If
    ' Make sure visit is unlocked
    If Not oEFI.VisitInstance.LockStatus = eLockStatus.lsUnlocked Then
        sMessage = "This visit is " & oEFI.LockStatusString & " and new eForms cannot be opened"
        Exit Function
    End If
    CanOpenEform = True

End Function

'---------------------------------------------------------------------
Private Function GetEFormIconByStatus(nFormStatus As Integer, enDiscrepStatus As MACRODEBS30.eDiscrepancyStatus, enLockStatus As LockStatus) As String
'---------------------------------------------------------------------
' Get the label for the form icon corresponding to the given form status
'---------------------------------------------------------------------


        Select Case enLockStatus
        Case eLockStatus.lsFrozen
            GetEFormIconByStatus = DM30_ICON_FROZEN
        Case eLockStatus.lsLocked
            GetEFormIconByStatus = DM30_ICON_LOCKED
        Case eLockStatus.lsUnlocked
            Select Case enDiscrepStatus
            Case MACRODEBS30.eDiscrepancyStatus.dsRaised
                GetEFormIconByStatus = DM30_ICON_RAISED_DISC
            Case MACRODEBS30.eDiscrepancyStatus.dsResponded
                GetEFormIconByStatus = DM30_ICON_RESPONDED_DISC
            Case Else
                Select Case nFormStatus
                Case eStatus.Inform
                    ' Yellow form
                    GetEFormIconByStatus = DM30_ICON_OK
            '   Case Status.Inform
            '        GetEFormIconByStatus = gsINFORM_CRF_PAGE_LABEL
                Case eStatus.requested
                    GetEFormIconByStatus = DM30_ICON_NEW_FORM
                Case eStatus.Missing
                    GetEFormIconByStatus = DM30_ICON_MISSING
                Case eStatus.InvalidData
                    GetEFormIconByStatus = DM30_ICON_WARNING 'need circle with diagonal
                Case eStatus.Success
                    ' Yellow form with tick
                    GetEFormIconByStatus = DM30_ICON_OK
                Case eStatus.Warning
                    GetEFormIconByStatus = DM30_ICON_WARNING
                Case eStatus.OKWarning
                    GetEFormIconByStatus = DM30_ICON_OK_WARNING
                Case eStatus.Unobtainable
                    GetEFormIconByStatus = DM30_ICON_UNOBTAINABLE
                Case eStatus.NotApplicable
                    GetEFormIconByStatus = DM30_ICON_NA
                Case Else
                    ' No status - display as inactive form
                    GetEFormIconByStatus = DM30_ICON_INACTIVE_FORM
                End Select
            End Select
        End Select

End Function

'---------------------------------------------------------------------
Private Function MUScheduleUpdated(ByVal sRedoMsg As String) As Boolean
'---------------------------------------------------------------------
' See if the schedule has been updated by other users
' Returns TRUE if schedule updated OR if there was an update problem,
' or FALSE if nothing's changed and all is OK (so it's safe to continue what we were doing)
' Gives user message if the grid has changed
' sRedoMsg is text for user to tell them to redo what they were doing
'---------------------------------------------------------------------
Dim sMsg As String
Dim oGrid As ScheduleGrid
Dim sLockErrMsg As String

    On Error GoTo ErrHandler
    
    MUScheduleUpdated = False
    
    ' Receive all updates from other users
    ' Call new StudySubject.Reload routine
    If moSubject.Reload(sLockErrMsg) Then
        ' There's been a reload so refresh the schedule
        Call RefreshGrid
        sMsg = "This subject has been updated by another user."
        sMsg = sMsg & vbCrLf & sRedoMsg
        DialogInformation sMsg
        MUScheduleUpdated = True
    Else
        ' No reload - was it because of a lock violation?
        ' In this case we don't ask them to try again
        If sLockErrMsg > "" Then
            DialogInformation sLockErrMsg
            MUScheduleUpdated = True
        End If
    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmSchedule.MUScheduleUpdated"

End Function


