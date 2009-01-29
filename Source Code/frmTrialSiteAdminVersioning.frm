VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTrialSiteAdminVersioning 
   Caption         =   "Study Site Administration"
   ClientHeight    =   6540
   ClientLeft      =   8310
   ClientTop       =   5385
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10035
   StartUpPosition =   1  'CenterOwner
   Tag             =   "KeepBottomRight"
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Clear All"
      Height          =   345
      Index           =   1
      Left            =   1230
      TabIndex        =   6
      Top             =   6180
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   8820
      TabIndex        =   9
      Top             =   6180
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   7620
      TabIndex        =   8
      Top             =   6180
      Width           =   1125
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "&History"
      Height          =   345
      Left            =   2400
      TabIndex        =   7
      Top             =   6180
      Width           =   1125
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Select All"
      Height          =   345
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   6180
      Width           =   1125
   End
   Begin TabDlg.SSTab tabAssocDistribute 
      Height          =   5355
      Left            =   60
      TabIndex        =   3
      Top             =   780
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   9446
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Participation"
      TabPicture(0)   =   "frmTrialSiteAdminVersioning.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lvwList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Distribution"
      TabPicture(1)   =   "frmTrialSiteAdminVersioning.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblWarning"
      Tab(1).Control(1)=   "lvwDistribute"
      Tab(1).Control(2)=   "cmdDistribute"
      Tab(1).Control(3)=   "cboVersion"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Recall"
      TabPicture(2)   =   "frmTrialSiteAdminVersioning.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvwRecall"
      Tab(2).Control(1)=   "cmdRecall"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdRecall 
         Caption         =   "&Recall"
         Height          =   345
         Left            =   -66480
         TabIndex        =   16
         Top             =   4920
         Width           =   1125
      End
      Begin MSComctlLib.ListView lvwRecall 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   15
         Top             =   540
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7646
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
      Begin VB.ComboBox cboVersion 
         Height          =   315
         Left            =   -71640
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   4920
         Width           =   4815
      End
      Begin VB.CommandButton cmdDistribute 
         Caption         =   "&Distribute"
         Height          =   345
         Left            =   -66480
         TabIndex        =   12
         Top             =   4920
         Width           =   1125
      End
      Begin MSComctlLib.ListView lvwList 
         Height          =   3555
         Left            =   240
         TabIndex        =   4
         Tag             =   "resize"
         Top             =   540
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   6271
         View            =   1
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Site"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "TrialSite"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwDistribute 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   10
         Tag             =   "resize"
         Top             =   540
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7646
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Site"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "TrialSite"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblWarning 
         Caption         =   "Please select a distributable study version"
         Height          =   315
         Left            =   -74760
         TabIndex        =   14
         Top             =   4980
         Width           =   5595
      End
   End
   Begin VB.OptionButton optSitesbyThing 
      Caption         =   "Sites by Study"
      Height          =   255
      Left            =   4740
      TabIndex        =   1
      Top             =   60
      Value           =   -1  'True
      Width           =   1755
   End
   Begin VB.OptionButton optThingsbySite 
      Caption         =   "Studies by Site"
      Height          =   255
      Left            =   4740
      TabIndex        =   2
      Top             =   420
      Width           =   1755
   End
   Begin VB.Frame fraCombo 
      Caption         =   "Studies"
      Height          =   735
      Left            =   60
      TabIndex        =   13
      Top             =   0
      Width           =   4575
      Begin VB.ComboBox cboSelect 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmTrialSiteAdminVersioning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998-2004. All Rights Reserved
'   File:       frmTrialSiteAdminVersioning.frm
'   Author:     David Hook, August 2002
'   Purpose:    Maintain cross-reference between trials and sites
'               Replaces frmTrialSiteAdmin
'--------------------------------------------------------------------------------
'Revisions
'   ZA 18/09/2002   update List_sites.js file when saving site details
' DPH 13/01/2003 Save when changing view
' NCJ 18 Nov 04 - Issue 2424 - Disable OK button once it's pressed
' MLM 11/04/07: Rewrote RefreshList, RefreshDistributionList, RefreshRecallList for performance reasons.
'--------------------------------------------------------------------------------


Option Explicit
Option Compare Binary
Option Base 0

'Private Const m_ICON_UNCHECKED = 1
'Private Const m_ICON_CHECKED = 2

'min form height and width
Private mlHeight As Long
Private mlWidth As Long
Private meDisplay As eDisplayType
Private mlSelectedTrialId As Long
Private msComboSelectedText As String
Private mbChangedParticipation As Boolean
Private mbFormSetup As Boolean
Private colChangedParticipation As New Collection
'ASH 11/12/2002
Private oDatabase As MACROUserBS30.Database
Private bLoad As Boolean
Private msConnectionstring As String
Private sMessage As String
Private mconMACRO As ADODB.Connection
Private msDatabase As String
' DPH 13/01/2003 - If saved from a change then use 'previous' value
Private mlPrevTrialId As Long
Private msPrevSite As String
Private msPrevStudyName As String


'---------------------------------------------------------------------
Public Function Display(sDatabase As String, _
                        eType As eDisplayType, _
                        Optional sName As String = "")
'---------------------------------------------------------------------
' display the form according to the eDisplayType
' sName will be selected in combo if passed through
'---------------------------------------------------------------------

Dim columnHdr As ColumnHeader

    On Error GoTo ErrHandler
    
    msDatabase = sDatabase
    mbFormSetup = True
    meDisplay = eType

    mlHeight = 4000
    mlWidth = 4000
    Me.Icon = frmMenu.Icon
    FormCentre Me

    Call DisplayData

    If meDisplay = DisplayTrialsBySite Or meDisplay = DisplayLabsBySite Or meDisplay = DisplayUsersBySite Then
        optThingsbySite.Value = True
    End If

    'set the combo if something passed through
    If sName <> "" Then
        cboSelect.Text = sName
    End If

    mbFormSetup = False

    Me.Show vbModal

Exit Function

ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Display")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Function

'---------------------------------------------------------------------
Private Sub cboSelect_Click()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------
' REVISIONS
' DPH 13/01/2003 - Save details on cboSelect / store previous
'---------------------------------------------------------------------

On Error GoTo ErrHandler

    If Not (mlSelectedTrialId = 0 And cboSelect.Text = "") Then
        If Not cboSelect.ListIndex = -1 Then
        
            ' DPH 13/01/2003 - Save details on cboSelect
            If mbChangedParticipation Then
                ' save detail
                Call SaveParticipationDetail(True)
                
                'ZA 18/09/2002 - updates List_sites.js
                CreateSitesList
            End If

            mlSelectedTrialId = cboSelect.ItemData(cboSelect.ListIndex)
            msComboSelectedText = cboSelect.Text
            Call RefreshList
            ' refresh version combo
            Call SetupVersionCombo
            ' refresh distribution listview
            Call RefreshDistributionList
            ' Refresh Recall list
            Call RefreshRecallList
        End If
        
        ' DPH 13/01/2003 - store previous
        If cboSelect.ListCount > 0 Then
            Call StorePreviousStudySite
        Else
            msPrevSite = ""
            mlPrevTrialId = -1
            msPrevStudyName = ""
        End If

    End If

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboSelect_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub cboVersion_Click()
'---------------------------------------------------------------------
' If a specific version is selected only show those versions in listview
'---------------------------------------------------------------------

    If Not mbFormSetup Then
        Call RefreshDistributionList
    End If

End Sub

'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------
' Cancel form - Checking if Cancel firstly
'---------------------------------------------------------------------

On Error GoTo ErrorHandler

    If mbChangedParticipation Then
        Select Case DialogQuestion("Do you want to save participating site changes?", "Save Changes", True)
        Case vbYes
            ' save detail
            Call SaveParticipationDetail
        Case vbNo
            ' do nothing
        Case vbCancel
            ' back to form
            Exit Sub
        End Select
    End If

    Unload Me
Exit Sub

ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdCancel_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub cmdDistribute_Click()
'---------------------------------------------------------------------
' Distribute Selected site/studies
'---------------------------------------------------------------------
Dim bDistribute As Boolean
Dim lVersion As Long
Dim oLIDistribute As ListItem
Dim lStudyId As Long
Dim sSiteCode As String
Dim sStudyName As String
Dim sVersionDesc As String
Dim sDistributeResults As String

On Error GoTo ErrorHandler

    ' get details based on view
    Select Case meDisplay
        Case DisplayTrialsBySite
            sSiteCode = cboSelect.List(cboSelect.ListIndex)
        Case DisplaySitesByTrial
            lStudyId = cboSelect.ItemData(cboSelect.ListIndex)
            sStudyName = GetTrialNameFromId(lStudyId)
            If cboVersion.Visible Then
                lVersion = cboVersion.ItemData(cboVersion.ListIndex)
                ' if 'Latest Version' get real latest version value
                If lVersion = 10000 Then
                    lVersion = GetLatestVersionOfTrialAvailable(lStudyId)
                End If
            Else
                lVersion = GetLatestVersionOfTrialAvailable(lStudyId)
            End If
            ' Get version description
            sVersionDesc = GetVersionDescription(lStudyId, lVersion)
    End Select

    ' confirm distribution through dialog
    For Each oLIDistribute In lvwDistribute.ListItems
        If oLIDistribute.Selected Then
            ' get row detail
            Select Case meDisplay
                Case DisplayTrialsBySite
                    sStudyName = oLIDistribute.Tag
                    lStudyId = TrialIdFromName(sStudyName)
                    ' Get latest version number
                    lVersion = GetLatestVersionOfTrialAvailable(lStudyId)
                    ' Get version description
                    sVersionDesc = GetVersionDescription(lStudyId, lVersion)
                Case DisplaySitesByTrial
                    sSiteCode = oLIDistribute.Tag
            End Select
            
            ' put details in confirmation dialog
            frmDistributeConfirm.Caption = "Study Distribution Confirmation"
            frmDistributeConfirm.lblConfirm = "Please confirm distribution of the following study version(s) :"
            Call frmDistributeConfirm.AddDetailToListView(sSiteCode, sStudyName, lVersion, sVersionDesc)
        End If
    Next
    frmDistributeConfirm.Show vbModal
    bDistribute = frmDistributeConfirm.mbDistribute
    Unload frmDistributeConfirm
    
    ' if distibution took place
    If bDistribute Then
        ' Distribute study versions
        For Each oLIDistribute In lvwDistribute.ListItems
            If oLIDistribute.Selected Then
                ' get row detail
                Select Case meDisplay
                    Case DisplayTrialsBySite
                        sStudyName = oLIDistribute.Tag
                        lStudyId = TrialIdFromName(sStudyName)
                        ' Get latest version number
                        lVersion = GetLatestVersionOfTrialAvailable(lStudyId)
                    Case DisplaySitesByTrial
                        sSiteCode = oLIDistribute.Tag
                End Select
                
                ' Distribute message function
                If sDistributeResults <> "" Then
                    sDistributeResults = sDistributeResults & vbCrLf
                End If
                sDistributeResults = sDistributeResults & DistributeVersionMessage(lStudyId, sStudyName, sSiteCode, lVersion)
                ' Mark Associated Listitem so is updated (using copied key on distribute list)
                Call CollectionAddAnyway(colChangedParticipation, lvwList.ListItems(oLIDistribute.Key), "K|" & lvwList.ListItems(oLIDistribute.Key).Index)
            End If
        Next
        
        ' show results dialog
        Call frmDistributionResults.InitialiseMe(sDistributeResults, True)
        frmDistributionResults.Show vbModal
        
        ' save associated site info to update version number
        Call SaveParticipationDetail
        
        ' refresh main listview
        Call RefreshList
    
        ' refresh distribution listview
        Call RefreshDistributionList
        
        ' Refresh Recall list
        Call RefreshRecallList
        
    End If
    
Exit Sub
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdDistribute_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub cmdHistory_Click()
'---------------------------------------------------------------------
' Open History form
'---------------------------------------------------------------------
Dim oLI As ListItem
Dim sSiteName As String
Dim sStudyName As String
Dim lStudyId As Long

On Error GoTo ErrorHandler
    
    ' get selected row and extract detail
    If Me.tabAssocDistribute.Tab = 0 Then
        For Each oLI In lvwList.ListItems
            If oLI.Selected Then
                Select Case meDisplay
                    Case DisplaySitesByTrial
                        lStudyId = GetcboSelectItemData(cboSelect.ListIndex)
                        sStudyName = cboSelect.List(cboSelect.ListIndex)
                        sSiteName = oLI.Tag
                    Case DisplayTrialsBySite
                        sSiteName = cboSelect.List(cboSelect.ListIndex)
                        sStudyName = oLI.Tag
                        lStudyId = TrialIdFromName(sStudyName)
                End Select
                ' Show History form
                Call frmStudySiteHistory.InitialiseMe(meDisplay, sSiteName, lStudyId, sStudyName)
                frmStudySiteHistory.Show vbModal
                Exit For
            End If
        Next
    Else
        For Each oLI In lvwDistribute.ListItems
            If oLI.Selected Then
                Select Case meDisplay
                    Case DisplaySitesByTrial
                        lStudyId = cboSelect.ItemData(cboSelect.ListIndex)
                        sStudyName = cboSelect.List(cboSelect.ListIndex)
                        sSiteName = oLI.Text
                    Case DisplayTrialsBySite
                        sSiteName = cboSelect.List(cboSelect.ListIndex)
                        sStudyName = oLI.Text
                        lStudyId = TrialIdFromName(sStudyName)
                End Select
                ' Show History form
                Call frmStudySiteHistory.InitialiseMe(meDisplay, sSiteName, lStudyId, sStudyName)
                frmStudySiteHistory.Show vbModal
                Exit For
            End If
        Next
    End If
    
Exit Sub
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdHistory_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click(Index As Integer)
'---------------------------------------------------------------------
' Complete form - Save firstly if necessary
'---------------------------------------------------------------------
    
    ' NCJ 18 Nov 04 - Disable OK button to prevent damage by further clicking
    cmdOK(Index).Enabled = False
    
    If mbChangedParticipation Then
        ' save detail
        Call SaveParticipationDetail
        
        'ZA 18/09/2002 - updates List_sites.js
        CreateSitesList
        
        Call frmMenu.RefereshTreeView
    End If
    
    Unload Me
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdRecall_Click()
'---------------------------------------------------------------------
' Recall Selected Messages to Site
'---------------------------------------------------------------------
Dim bRecall As Boolean
Dim lVersion As Long
Dim oLIRecall As ListItem
Dim lStudyId As Long
Dim sSiteCode As String
Dim sStudyName As String
Dim sVersionDesc As String
Dim sRecallResults As String

On Error GoTo ErrorHandler

    ' get details based on view
    Select Case meDisplay
        Case DisplayTrialsBySite
            sSiteCode = cboSelect.List(cboSelect.ListIndex)
        Case DisplaySitesByTrial
            lStudyId = cboSelect.ItemData(cboSelect.ListIndex)
            sStudyName = GetTrialNameFromId(lStudyId)
    End Select

    ' confirm distribution through dialog
    For Each oLIRecall In lvwRecall.ListItems
        If oLIRecall.Selected Then
            ' get row detail
            Select Case meDisplay
                Case DisplayTrialsBySite
                    sStudyName = oLIRecall.Tag
                    lStudyId = TrialIdFromName(sStudyName)
                Case DisplaySitesByTrial
                    sSiteCode = oLIRecall.Tag
            End Select
            
            ' Get version in column to recall
            lVersion = RemoveNull(oLIRecall.SubItems(LatestVersionSubItemIndex))
            ' Get version description
            sVersionDesc = GetVersionDescription(lStudyId, lVersion)
            
            ' put details in confirmation dialog
            frmDistributeConfirm.Caption = "Study Version Recall Confirmation"
            frmDistributeConfirm.lblConfirm = "Please confirm recall of the following study version(s) :"
            Call frmDistributeConfirm.AddDetailToListView(sSiteCode, sStudyName, lVersion, sVersionDesc)
        End If
    Next
    frmDistributeConfirm.Show vbModal
    bRecall = frmDistributeConfirm.mbDistribute
    Unload frmDistributeConfirm
    
    ' if distibution took place
    If bRecall Then
        ' Distribute study versions
        For Each oLIRecall In lvwRecall.ListItems
            If oLIRecall.Selected Then
                ' get row detail
                Select Case meDisplay
                    Case DisplayTrialsBySite
                        sStudyName = oLIRecall.Tag
                        lStudyId = TrialIdFromName(sStudyName)
                    Case DisplaySitesByTrial
                        sSiteCode = oLIRecall.Tag
                End Select
                
                ' Get version in column to recall
                lVersion = RemoveNull(oLIRecall.SubItems(LatestVersionSubItemIndex))
                
                ' Distribute message function
                If sRecallResults <> "" Then
                    sRecallResults = sRecallResults & vbCrLf
                End If
                
                sRecallResults = sRecallResults & RecallVersionMessage(lStudyId, sStudyName, sSiteCode, lVersion)
                ' Mark Associated Listitem so is updated (using copied key on distribute list)
                Call CollectionAddAnyway(colChangedParticipation, lvwList.ListItems(oLIRecall.Key), "K|" & lvwList.ListItems(oLIRecall.Key).Index)
            End If
        Next
        
        ' show results dialog
        Call frmDistributionResults.InitialiseMe(sRecallResults, False)
        frmDistributionResults.Show vbModal
        
        ' save associated site info to update version number
        Call SaveParticipationDetail
        
        ' refresh main listview
        Call RefreshList
    
        ' refresh distribution listview
        Call RefreshDistributionList
        
        ' Refresh Recall list
        Call RefreshRecallList
        
    End If
    
Exit Sub
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdRecall_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub cmdSelectAll_Click(Index As Integer)
'---------------------------------------------------------------------
' select all / clear all items
'---------------------------------------------------------------------
Dim oListItem As MSComctlLib.ListItem

On Error GoTo ErrorHandler
    
    HourglassOn
    ' Dependant on which tab
    If tabAssocDistribute.Tab = 0 Then
        For Each oListItem In lvwList.ListItems
            If (Not oListItem.Checked And Index = 0) Or (oListItem.Checked And Index = 1) Then
                'if unchecked and index is 0 (Select) OR checked and index = 1 (Clear) then click it
                oListItem.Checked = Not oListItem.Checked
                Call lvwList_ItemCheck(oListItem)
            End If
        Next
    Else
        For Each oListItem In lvwDistribute.ListItems
            If (Not oListItem.Selected And Index = 0) Or (oListItem.Selected And Index = 1) Then
                'if not selected and index is 0 (Select) OR selected and index = 1 (Clear) then click it
                oListItem.Selected = Not oListItem.Selected
            End If
        Next
        ' if cleared all on distribute form disable distribute button
        If Index = 1 Then
            cmdDistribute.Enabled = False
        Else
            ' check for rows
            If lvwDistribute.ListItems.Count > 0 Then
                cmdDistribute.Enabled = True
            End If
        End If
    End If

    HourglassOff

Exit Sub
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdSelectAll_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Sub

Private Sub Form_Resize()

On Error GoTo ErrHandler

'    If Me.WindowState <> vbMinimized Then
'        If Me.Height >= mlHeight Then
'            fraSites.Height = Me.ScaleHeight - cmdOK(0).Height - fraCombo.Height - 240
'            lvwList.Height = fraSites.Height - 360
'            cmdOK(0).Top = fraSites.Top + fraSites.Height + 120
'            cmdOK(1).Top = cmdOK(0).Top
'            cmdOK(2).Top = cmdOK(0).Top
'            '   Bug Fix Macro 2.2 008 RJCW 15/10/2001
'            cmdAll(0).Top = cmdOK(0).Top
'            cmdAll(1).Top = cmdOK(0).Top
'        End If
'        If Me.Width >= mlWidth Then
'            fraSites.Width = Me.ScaleWidth - 120
'            lvwList.Width = fraSites.Width - 240
'            cmdOK(0).Left = fraSites.Left + fraSites.Width - cmdOK(0).Width
''            cmdOK(1).Left = fraSites.Left + fraSites.Width - cmdOK(1).Width
''            cmdOK(2).Left = fraSites.Left + fraSites.Width - cmdOK(2).Width
'        End If
'    End If

ErrHandler:

End Sub

'---------------------------------------------------------------------
Public Sub RefreshList()
'---------------------------------------------------------------------
' MLM 16/04/07: Rewritten to retreive all data for the listview from the database
'   in one query for performance reasons.
'---------------------------------------------------------------------
Dim sSQL As String
Dim sThisSite
Dim bCurrentSite As Boolean
Dim rsListview As ADODB.Recordset
Dim skey As String
Dim sField As String
Dim oLIList As ListItem
Dim eSiteLocation As SiteLocation

    On Error GoTo ErrHandler

    ' Reset Participation dirty flag and collection
    mbChangedParticipation = False
    Set colChangedParticipation = New Collection

    ' lock the window to prevent updates
    Call LockWindow(lvwList)

    With lvwList
        .ListItems.Clear
        .SortKey = 0 'sort using the listitem object's text property.
        .SortOrder = lvwAscending
        .Sorted = True
    End With

    Select Case meDisplay
    Case DisplayTrialsBySite
'        sSQL = "SELECT ClinicalTrialName, ClinicalTrialDescription FROM ClinicalTrial WHERE ClinicalTrialId > 0 ORDER BY ClinicalTrialName"
        eSiteLocation = GetSiteLocation(cboSelect.List(cboSelect.ListIndex))
'    Case DisplaySitesByTrial
'        sSQL = "SELECT Site, SiteDescription FROM Site WHERE SiteStatus = 0"
    End Select


    Set rsListview = New ADODB.Recordset
    With rsListview
        .Open GetListViewSql(), mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText
        While Not .EOF
            ' "K|" is added to the key because the site might be numeric
            ' and listviews can't have numeric Keys
            skey = "K|" & .Fields(0)
            sThisSite = .Fields(0)

            'bCurrentSite = IsMappedSiteOrTrial(sThisSite)

            With lvwList.ListItems.Add(, skey, CStr(.Fields(0).Value))
                .Checked = Not IsNull(rsListview.Fields(2).Value)
                'TA 29/09/2000: tag is id
                .Tag = Format(rsListview.Fields(0).Value)
                Select Case meDisplay
                Case DisplayTrialsBySite
                    .SubItems(1) = CStr(rsListview.Fields(1).Value) 'ClinicalTrialId
                    If eSiteLocation = SiteLocation.ESiteServer And .Checked Then
                        .SubItems(2) = "Latest"
                    End If
                Case DisplaySitesByTrial
                    Select Case rsListview.Fields(1).Value
                    Case SiteLocation.ESiteRemote
                        .SubItems(1) = "Remote"
                    Case SiteLocation.ESiteServer
                        .SubItems(1) = "Server"
                    End Select
                End Select
                If Not IsNull(rsListview.Fields("distver").Value) Then 'has been deployed
                    .SubItems(2) = rsListview.Fields("distver").Value
                    .SubItems(3) = Format(rsListview.Fields("distts").Value, "yyyy/MM/dd hh:mm")
                End If
                If Not IsNull(rsListview.Fields("recver").Value) Then 'has been deployed
                    .SubItems(4) = rsListview.Fields("recver").Value
                    .SubItems(5) = Format(rsListview.Fields("rects").Value, "yyyy/MM/dd hh:mm")
                End If
            End With
            .MoveNext
        Wend
        .Close
    End With

    lvwList.Sorted = True

    ' get history detail
    'Call PopulateListViewWithHistoryDetail

    ' History Button
    If lvwList.ListItems.Count > 0 Then
        For Each oLIList In lvwList.ListItems
            If oLIList.Selected Then
                Call lvwList_ItemClick(oLIList)
            End If
        Next
    End If
    
    ' unlock the window for updates
    Call UnlockWindow
    Set rsListview = Nothing
    
    lvw_SetAllColWidths lvwList, LVSCW_AUTOSIZE_USEHEADER
    'MLM 22/04/07: don't auto-size the new hidden column
    If meDisplay = DisplayTrialsBySite Then
        lvwList.ColumnHeaders.Item(2).Width = 0
    End If

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshList")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Sub

'---------------------------------------------------------------------
Private Function GetListViewSql() As String
'---------------------------------------------------------------------
' MLM 16/04/07: Created. Retrieve appropriate SQL based on db type
'   and selected study or site. SQL returns all sites/studies, whether
'   they pertain to the selected study or site, the max version ever
'   distributed, the max time at which this version was distributed,
'   the max version received and the max time for that received version.
'   Exclude inactive sites and the library.
'---------------------------------------------------------------------
Dim sSQL As String

    Select Case meDisplay
    Case DisplayTrialsBySite
        Select Case Connection_Property(CONNECTION_PROVIDER, msConnectionstring)
        Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE 'Oracle
            sSQL = "select clinicaltrial.clinicaltrialname, clinicaltrial.clinicaltrialid, trialsite.trialsite," & _
                " max(to_number(substr(distver.messageparameters, 1, instr(distver.messageparameters, clinicaltrial.clinicaltrialname) - 1))) distver," & _
                " max(to_number(substr(recver.messageparameters, 1, instr(recver.messageparameters, clinicaltrial.clinicaltrialname) - 1))) recver," & _
                " max(distts.messagetimestamp) distts, max(rects.messagereceivedtimestamp) rects" & _
                " from clinicaltrial, trialsite, message distver, message distts, message recver, message rects" & _
                " where clinicaltrial.clinicaltrialid = trialsite.clinicaltrialid(+)" & _
                " and distver.clinicaltrialid(+) = trialsite.clinicaltrialid" & _
                " and distver.trialsite(+) = trialsite.trialsite" & _
                " and distts.clinicaltrialid(+) = trialsite.clinicaltrialid" & _
                " and distts.trialsite(+) = trialsite.trialsite" & _
                " and recver.clinicaltrialid(+) = trialsite.clinicaltrialid" & _
                " and recver.trialsite(+) = trialsite.trialsite" & _
                " and rects.clinicaltrialid(+) = trialsite.clinicaltrialid" & _
                " and rects.trialsite(+) = trialsite.trialsite" & _
                " and clinicaltrial.clinicaltrialid <> 0" & _
                " and trialsite.trialsite(+) = '" & cboSelect.List(cboSelect.ListIndex) & _
                "' and distver.messagetype(+) = 8 and distts.messagetype(+) = 8" & _
                " and recver.messagetype(+) = 8 and rects.messagetype(+) = 8" & _
                " and recver.messagereceivedtimestamp(+) <> 0 and rects.messagereceivedtimestamp(+) <> 0"
            sSQL = sSQL & " group by clinicaltrial.clinicaltrialid, clinicaltrial.clinicaltrialname, trialsite.trialsite," & _
                " substr(distts.messageparameters, 1, instr(distts.messageparameters, clinicaltrial.clinicaltrialname) - 1)," & _
                " substr(rects.messageparameters, 1, instr(rects.messageparameters, clinicaltrial.clinicaltrialname) - 1)" & _
                " having (substr(distts.messageparameters, 1, instr(distts.messageparameters, clinicaltrial.clinicaltrialname) - 1) is null" & _
                " or max(to_number(substr(distver.messageparameters, 1, instr(distver.messageparameters, clinicaltrial.clinicaltrialname) - 1))) =" & _
                " substr(distts.messageparameters, 1, instr(distts.messageparameters, clinicaltrial.clinicaltrialname) - 1))" & _
                " and (substr(rects.messageparameters, 1, instr(rects.messageparameters, clinicaltrial.clinicaltrialname) - 1) is null" & _
                " or max(to_number(substr(recver.messageparameters, 1, instr(recver.messageparameters, clinicaltrial.clinicaltrialname) - 1))) =" & _
                " substr(rects.messageparameters, 1, instr(rects.messageparameters, clinicaltrial.clinicaltrialname) - 1))" & _
                " order by clinicaltrial.clinicaltrialname"
        Case Else ' assume SQL Server
            sSQL = "select clinicaltrial.clinicaltrialname, clinicaltrial.clinicaltrialid, trialsite.trialsite," & _
                " max(cast(substring(distver.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, distver.messageparameters) - 1) as int)) distver," & _
                " max(cast(substring(recver.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, recver.messageparameters) - 1) as int)) recver," & _
                " max(distts.messagetimestamp) distts, max(rects.messagereceivedtimestamp) rects" & _
                " from ((((clinicaltrial left join trialsite on clinicaltrial.clinicaltrialid = trialsite.clinicaltrialid" & _
                " and trialsite.trialsite = '" & cboSelect.List(cboSelect.ListIndex) & _
                "') left join message distver on distver.clinicaltrialid = trialsite.clinicaltrialid" & _
                " and distver.trialsite = trialsite.trialsite and distver.messagetype = 8)" & _
                " left join message distts on distts.clinicaltrialid = trialsite.clinicaltrialid" & _
                " and distts.trialsite = trialsite.trialsite and distts.messagetype = 8)" & _
                " left join message recver on recver.clinicaltrialid = trialsite.clinicaltrialid" & _
                " and recver.trialsite = trialsite.trialsite and recver.messagetype = 8" & _
                " and recver.messagereceivedtimestamp <> 0)" & _
                " left join message rects on rects.clinicaltrialid = trialsite.clinicaltrialid" & _
                " and rects.trialsite = trialsite.trialsite and rects.messagetype = 8" & _
                " and rects.messagereceivedtimestamp <> 0 where clinicaltrial.ClinicalTrialId <> 0"
            sSQL = sSQL & " group by clinicaltrial.clinicaltrialid, clinicaltrial.clinicaltrialname, trialsite.trialsite," & _
                " substring(distts.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, distts.messageparameters) - 1)," & _
                " substring(rects.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, rects.messageparameters) - 1)" & _
                " having (substring(distts.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, distts.messageparameters) - 1) is null" & _
                " or max(cast(substring(distver.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, distver.messageparameters) - 1) as int)) =" & _
                " substring(distts.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, distts.messageparameters) - 1))" & _
                " and (substring(rects.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, rects.messageparameters) - 1) is null" & _
                " or max(cast(substring(recver.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, recver.messageparameters) - 1) as int)) =" & _
                " substring(rects.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, rects.messageparameters) - 1))" & _
                " order by clinicaltrial.clinicaltrialname"
        End Select
    Case DisplaySitesByTrial
        Select Case Connection_Property(CONNECTION_PROVIDER, msConnectionstring)
        Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE 'Oracle
            sSQL = "select site.site, site.sitelocation, clinicaltrial.clinicaltrialname," & _
                " max(to_number(substr(distver.messageparameters, 1, instr(distver.messageparameters, clinicaltrial.clinicaltrialname) - 1))) distver," & _
                " max(to_number(substr(recver.messageparameters, 1, instr(recver.messageparameters, clinicaltrial.clinicaltrialname) - 1))) recver," & _
                " max(distts.messagetimestamp) distts, max(rects.messagereceivedtimestamp) rects" & _
                " from site, clinicaltrial, trialsite, message distver, message distts, message recver, message rects" & _
                " where trialsite.trialsite(+) = site.site" & _
                " and clinicaltrial.clinicaltrialid(+) = trialsite.clinicaltrialid" & _
                " and distver.clinicaltrialid(+) = trialsite.clinicaltrialid" & _
                " and distver.trialsite(+) = trialsite.trialsite" & _
                " and distts.clinicaltrialid(+) = trialsite.clinicaltrialid" & _
                " and distts.trialsite(+) = trialsite.trialsite" & _
                " and recver.clinicaltrialid(+) = trialsite.clinicaltrialid" & _
                " and recver.trialsite(+) = trialsite.trialsite" & _
                " and rects.clinicaltrialid(+) = trialsite.clinicaltrialid" & _
                " and rects.trialsite(+) = trialsite.trialsite" & _
                " and distver.messagetype(+) = 8 and distts.messagetype(+) = 8" & _
                " and recver.messagetype(+) = 8 and rects.messagetype(+) = 8" & _
                " and site.SiteStatus = 0 and trialsite.clinicaltrialid(+) = " & GetcboSelectItemData(cboSelect.ListIndex) & _
                " and recver.messagereceivedtimestamp(+) <> 0 and rects.messagereceivedtimestamp(+) <> 0"
            sSQL = sSQL & " group by site.site, site.sitelocation, clinicaltrial.clinicaltrialname," & _
                " substr(distts.messageparameters, 1, instr(distts.messageparameters, clinicaltrial.clinicaltrialname) - 1)," & _
                " substr(rects.messageparameters, 1, instr(rects.messageparameters, clinicaltrial.clinicaltrialname) - 1)" & _
                " having (substr(distts.messageparameters, 1, instr(distts.messageparameters, clinicaltrial.clinicaltrialname) - 1) is null" & _
                " or max(to_number(substr(distver.messageparameters, 1, instr(distver.messageparameters, clinicaltrial.clinicaltrialname) - 1))) =" & _
                " substr(distts.messageparameters, 1, instr(distts.messageparameters, clinicaltrial.clinicaltrialname) - 1))" & _
                " and (substr(rects.messageparameters, 1, instr(rects.messageparameters, clinicaltrial.clinicaltrialname) - 1) is null" & _
                " or max(to_number(substr(recver.messageparameters, 1, instr(recver.messageparameters, clinicaltrial.clinicaltrialname) - 1))) =" & _
                " substr(rects.messageparameters, 1, instr(rects.messageparameters, clinicaltrial.clinicaltrialname) - 1))" & _
                " order by site.site"
        Case Else ' assume SQL Server
            sSQL = "select site.site, site.sitelocation, clinicaltrial.clinicaltrialname," & _
                " max(cast(substring(distver.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, distver.messageparameters) - 1) as int)) distver," & _
                " max(cast(substring(recver.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, recver.messageparameters) - 1) as int)) recver," & _
                " max(distts.messagetimestamp) distts, max(rects.messagereceivedtimestamp) rects" & _
                " from (((((site left join trialsite on trialsite.trialsite = site.site" & _
                " and trialsite.clinicaltrialid = " & GetcboSelectItemData(cboSelect.ListIndex) & _
                ") left join clinicaltrial on clinicaltrial.clinicaltrialid = trialsite.clinicaltrialid)" & _
                " left join message distver on distver.clinicaltrialid = trialsite.clinicaltrialid" & _
                " and distver.trialsite = trialsite.trialsite and distver.messagetype = 8)" & _
                " left join message distts on distts.clinicaltrialid = trialsite.clinicaltrialid" & _
                " and distts.trialsite = trialsite.trialsite and distts.messagetype = 8)" & _
                " left join message recver on recver.clinicaltrialid = trialsite.clinicaltrialid" & _
                " and recver.trialsite = trialsite.trialsite and recver.messagetype = 8" & _
                " and recver.messagereceivedtimestamp <> 0)" & _
                " left join message rects on rects.clinicaltrialid = trialsite.clinicaltrialid" & _
                " and rects.trialsite = trialsite.trialsite and rects.messagetype = 8" & _
                " and rects.messagereceivedtimestamp <> 0 where Site.SiteStatus = 0"
            sSQL = sSQL & " group by site.site, site.sitelocation, clinicaltrial.clinicaltrialname," & _
                " substring(distts.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, distts.messageparameters) - 1)," & _
                " substring(rects.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, rects.messageparameters) - 1)" & _
                " having (substring(distts.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, distts.messageparameters) - 1) is null" & _
                " or max(cast(substring(distver.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, distver.messageparameters) - 1) as int)) =" & _
                " substring(distts.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, distts.messageparameters) - 1))" & _
                " and (substring(rects.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, rects.messageparameters) - 1) is null" & _
                " or max(cast(substring(recver.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, recver.messageparameters) - 1) as int)) =" & _
                " substring(rects.messageparameters, 1, charindex(clinicaltrial.clinicaltrialname, rects.messageparameters) - 1))" & _
                " order by site.site"
        End Select
    End Select
    GetListViewSql = sSQL

End Function

'---------------------------------------------------------------------
Private Sub lvwDistribute_ItemClick(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------
' If an item is clicked upon enable cmdDistribute button
'---------------------------------------------------------------------

    cmdDistribute.Enabled = True
    CmdHistory.Enabled = True
    
End Sub

'---------------------------------------------------------------------
Private Sub lvwList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------
' Mark the Participating site as participating or not
'---------------------------------------------------------------------

    ' if BY site/study exists
    If cboSelect.Text > "" Then
        mbChangedParticipation = True
        Call CollectionAddAnyway(colChangedParticipation, Item, "K|" & Item.Index)
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------
' Enable the history button if remote site/study combination
'---------------------------------------------------------------------
Dim sSite As String

    ' if BY site/study exists
    If cboSelect.Text > "" Then
        Select Case meDisplay
            Case DisplayTrialsBySite
                sSite = cboSelect.List(cboSelect.ListIndex)
            Case DisplaySitesByTrial
                sSite = Item.Tag
        End Select
    
        ' Enable the history button if a remote site
        If GetSiteLocation(sSite) = SiteLocation.ESiteRemote Then
            CmdHistory.Enabled = True
        Else
            CmdHistory.Enabled = False
        End If
    Else
        CmdHistory.Enabled = False
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub lvwRecall_ItemClick(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------
' If an item is clicked upon enable cmdDistribute button
'---------------------------------------------------------------------

    cmdRecall.Enabled = True
    
End Sub

'---------------------------------------------------------------------
Private Sub optSitesbyThing_Click()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    Call ChangeRound

End Sub

'---------------------------------------------------------------------
Private Sub optThingsbySite_Click()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    Call ChangeRound

End Sub

'MLM 16/04/07: Removed:
''---------------------------------------------------------------------
'Public Function IsMappedSiteOrTrial(ByVal sTrialOrSite As String) As Boolean
''---------------------------------------------------------------------
'    'Returns TRUE if Site exists in trial table.
'    'Means that this site is associated with the trial.
''---------------------------------------------------------------------
'Dim sSQL As String
'Dim sField As String
'Dim rsMapped As ADODB.Recordset
'
'    On Error GoTo ErrHandler
'    'sSQL returns a recordset of all sites used for this trial.
'    Select Case meDisplay
'    Case DisplaySitesByTrial
'        sSQL = "SELECT  TrialSite FROM TrialSite, ClinicalTrial " _
'                & "WHERE TrialSite.ClinicalTrialId = ClinicalTrial.ClinicalTrialId" _
'                & " AND ClinicalTrialName = '" & RTrim(msComboSelectedText) & "'"
'    Case DisplayTrialsBySite
'        sSQL = "SELECT ClinicalTrial.clinicaltrialname FROM ClinicalTrial, TrialSite" _
'                & " WHERE TrialSite.ClinicalTrialId = ClinicalTrial.ClinicalTrialId" _
'                & " AND TrialSite.TrialSite = '" & RTrim(msComboSelectedText) & "'"
'    End Select
'
'    Set rsMapped = New ADODB.Recordset
'    With rsMapped
'
'       .Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'        'Compare each(active) site (ThisSite - in the listview) with the set of sites used for
'        'this trial.If a match is found ,return true. The site is 'checked' in the listview.
'        IsMappedSiteOrTrial = False
'        Do While Not .EOF
'            If sTrialOrSite = .Fields(0) Then
'                IsMappedSiteOrTrial = True
'                Exit Do
'            End If
'            .MoveNext
'        Loop
'
'        .Close
'    End With
'
'    Set rsMapped = Nothing
'
'Exit Function
'ErrHandler:
'  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "IsMappedSiteOrTrial")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Unload frmMenu
'   End Select
'
'End Function

'---------------------------------------------------------------------
Private Sub ChangeRound()
'---------------------------------------------------------------------
' switch between things by Site and Sites by things
'---------------------------------------------------------------------
' REVISIONS
' DPH 13/01/2003 - Save details on change of optionButton
'---------------------------------------------------------------------

    ' DPH 13/01/2003 - Save details on change of optionButton
    If mbChangedParticipation Then
        ' save detail
        Call SaveParticipationDetail(True)
        
        'ZA 18/09/2002 - updates List_sites.js
        CreateSitesList
    End If

    If optThingsbySite.Value Then
        Select Case meDisplay
        Case DisplaySitesByTrial: meDisplay = DisplayTrialsBySite
        Case DisplaySitesByUser: meDisplay = DisplayUsersBySite
        End Select
    Else
        Select Case meDisplay
        Case DisplayTrialsBySite: meDisplay = DisplaySitesByTrial
        Case DisplayUsersBySite: meDisplay = DisplaySitesByUser
        End Select
    End If

    Call DisplayData

End Sub

'---------------------------------------------------------------------
Private Sub DisplayData()
'---------------------------------------------------------------------
' redisplay all data
'---------------------------------------------------------------------
' DPH 13/01/2003 - Store 'previous' study/site so can save when a change occurs
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    ' reset Participation details changed flag
    mbChangedParticipation = False

    ' study / sites combo
    Call LoadCombo

    Me.Caption = "Study Site Administration " & "[" & goUser.DatabaseCode & "]"
    optThingsbySite.Caption = "Studies by Site"
    optSitesbyThing.Caption = "Sites by Study"

    ' setup columnheaders
    Call SetupColumnHeaders

    ' load listview
    Call LoadListView

    ' If they are in trial/site admin and they can't add or remove trials/sites then disable the listview
    lvwList.Enabled = goUser.CheckPermission(gsFnAddSiteToTrialOrTrialToSite) Or goUser.CheckPermission(gsFnRemoveSite)

    If cboSelect.ListCount > 0 Then
        cboSelect.ListIndex = 0
        
        Call StorePreviousStudySite
    Else
        msPrevSite = ""
        mlPrevTrialId = -1
        msPrevStudyName = ""
    End If

    ' set version combo
    Call SetupVersionCombo

    ' Check tab
    Call VersionComboControlByTab

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "DisplayData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
End Sub

'---------------------------------------------------------------------
Public Sub LoadCombo()
'---------------------------------------------------------------------

'---------------------------------------------------------------------
Dim rsCombo As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    cboSelect.Clear
    
    Set oDatabase = New MACROUserBS30.Database
    bLoad = oDatabase.Load(SecurityADODBConnection, goUser.UserName, msDatabase, "", False, sMessage)
    msConnectionstring = oDatabase.ConnectionString
    Set mconMACRO = New ADODB.Connection
    mconMACRO.Open msConnectionstring
    mconMACRO.CursorLocation = adUseClient

    
    Set rsCombo = New ADODB.Recordset

    With rsCombo
        Select Case meDisplay
            Case DisplaySitesByTrial
                sSQL = "SELECT ClinicalTrialName, ClinicalTrialId from ClinicalTrial WHERE ClinicalTrialId > 0 ORDER BY ClinicalTrialName"
                .Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText
            Case DisplayTrialsBySite, DisplayLabsBySite, DisplayUsersBySite
                 sSQL = "SELECT Site FROM Site WHERE SiteStatus = 0"
                .Open sSQL, mconMACRO
        End Select

        Do While Not .EOF
            cboSelect.AddItem .Fields(0)
            If .Fields.Count > 1 Then
                cboSelect.ItemData(cboSelect.NewIndex) = .Fields(1)
            End If
            .MoveNext
        Loop

        .Close
    End With

    Set rsCombo = Nothing

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadCombo")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Sub

'---------------------------------------------------------------------
Public Sub LoadListView()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------
Dim rsListview As ADODB.Recordset
Dim sSQL As String
Dim sField As String

    On Error GoTo ErrHandler

    lvwList.ListItems.Clear
    lvwList.View = lvwReport
    
    Set oDatabase = New MACROUserBS30.Database
    bLoad = oDatabase.Load(SecurityADODBConnection, goUser.UserName, msDatabase, "", False, sMessage)
    msConnectionstring = oDatabase.ConnectionString
    Set mconMACRO = New ADODB.Connection
    mconMACRO.Open msConnectionstring
    mconMACRO.CursorLocation = adUseClient


    Select Case meDisplay
    Case DisplayTrialsBySite
        sSQL = "SELECT ClinicalTrialName, ClinicalTrialDescription from ClinicalTrial where clinicaltrialid > 0 ORDER BY ClinicalTrialName"
    Case DisplaySitesByTrial
        sSQL = "SELECT Site, SiteDescription FROM Site WHERE SiteStatus = 0"
    End Select

    Set rsListview = New ADODB.Recordset
    With rsListview
        .Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not .EOF
            With lvwList.ListItems.Add(, , .Fields(0))
                '.SubItems(1) = rsListview.Fields(1)
            End With
            .MoveNext
        Loop
        .Close
    End With

    Set rsListview = Nothing

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadListView")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub SetupColumnHeaders()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------
' REVISIONS
' DPH 03/09/2002 - New Headers for Recall Listview
'---------------------------------------------------------------------
On Error GoTo ErrorHandler

    Select Case meDisplay
        Case DisplayTrialsBySite
            ' Set headers on Associated list
            With lvwList
                .ColumnHeaders.Clear
                fraCombo.Caption = "Site"
                Call .ColumnHeaders.Add(, , "Studies", 2000, lvwColumnLeft)
                'MLM 11/04/07: hidden column to store ClinicalTrialId
                Call .ColumnHeaders.Add(, , "Study ID", 0)

                'Sort on first column
                .SortKey = 0 'sort using the listitem object's text property.
                .SortOrder = lvwAscending
                .Sorted = True
            End With

            ' set headers on distribute list
            With lvwDistribute
                .ColumnHeaders.Clear

                Call .ColumnHeaders.Add(, , "Studies", 2000, lvwColumnLeft)
                'MLM 11/04/07: hidden column to store ClinicalTrialId
                Call .ColumnHeaders.Add(, , "Study ID", 0)

                .MultiSelect = True

                'Sort on first column
                .SortKey = 0 'sort using the listitem object's text property.
                .SortOrder = lvwAscending
                .Sorted = True

            End With

            ' set headers on distribute list
            With lvwRecall
                .ColumnHeaders.Clear

                Call .ColumnHeaders.Add(, , "Studies", 2000, lvwColumnLeft)
                'MLM 11/04/07: hidden column to store ClinicalTrialId
                Call .ColumnHeaders.Add(, , "Study ID", 0)

                .MultiSelect = True

                'Sort on first column
                .SortKey = 0 'sort using the listitem object's text property.
                .SortOrder = lvwAscending
                .Sorted = True

            End With
        
        Case DisplaySitesByTrial

            With lvwList
                .ColumnHeaders.Clear
                fraCombo.Caption = "Study"

                Call .ColumnHeaders.Add(, , "Active Sites", 2000, lvwColumnLeft)

                ' Other Headers
                Call .ColumnHeaders.Add(, , "Type", 800, lvwColumnLeft)

                .MultiSelect = True

                'Sort on first column
                .SortKey = 0 'sort using the listitem object's text property.
                .SortOrder = lvwAscending
                .Sorted = True
            End With

            With lvwDistribute
                .ColumnHeaders.Clear

                Call .ColumnHeaders.Add(, , "Active Sites", 2000, lvwColumnLeft)

                ' Other Headers
                Call .ColumnHeaders.Add(, , "Type", 800, lvwColumnLeft)

                .MultiSelect = True

                'Sort on first column
                .SortKey = 0 'sort using the listitem object's text property.
                .SortOrder = lvwAscending
                .Sorted = True
            End With

            With lvwRecall
                .ColumnHeaders.Clear

                Call .ColumnHeaders.Add(, , "Active Sites", 2000, lvwColumnLeft)

                ' Other Headers
                Call .ColumnHeaders.Add(, , "Type", 800, lvwColumnLeft)

                .MultiSelect = True

                'Sort on first column
                .SortKey = 0 'sort using the listitem object's text property.
                .SortOrder = lvwAscending
                .Sorted = True
            End With

        End Select

    ' Add columns on both listviews
    With lvwList
        Call .ColumnHeaders.Add(, , "Deployed Version", 1500, lvwColumnLeft)
        Call .ColumnHeaders.Add(, , "Deployed Date", 2000, lvwColumnLeft)
        Call .ColumnHeaders.Add(, , "Received Version", 1500, lvwColumnLeft)
        Call .ColumnHeaders.Add(, , "Received Date", 2000, lvwColumnLeft)
    End With

    With lvwDistribute
        Call .ColumnHeaders.Add(, , "Deployed Version", 1500, lvwColumnLeft)
        Call .ColumnHeaders.Add(, , "Deployed Date", 2000, lvwColumnLeft)
        Call .ColumnHeaders.Add(, , "Received Version", 1500, lvwColumnLeft)
        Call .ColumnHeaders.Add(, , "Received Date", 2000, lvwColumnLeft)
    End With

    With lvwRecall
        Call .ColumnHeaders.Add(, , "Deployed Version", 1500, lvwColumnLeft)
        Call .ColumnHeaders.Add(, , "Deployed Date", 2000, lvwColumnLeft)
        Call .ColumnHeaders.Add(, , "Received Version", 1500, lvwColumnLeft)
        Call .ColumnHeaders.Add(, , "Received Date", 2000, lvwColumnLeft)
    End With

Exit Sub
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SetupColumnHeaders")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Sub

'---------------------------------------------------------------------
Private Sub SetupVersionCombo()
'---------------------------------------------------------------------
'  Set up version numbers if in sitesbystudy mode
'  else hide combo and show different message on label
'---------------------------------------------------------------------
' REVISIONS
' DPH 03/09/2002 - Added Description to Version combo
'---------------------------------------------------------------------
Dim rsCombo As ADODB.Recordset
Dim sSQL As String
Dim sComboVersion As String

    On Error GoTo ErrHandler
    cboVersion.Clear

    Select Case meDisplay
        Case DisplaySitesByTrial

            cboVersion.Visible = True

            Set rsCombo = New ADODB.Recordset

            With rsCombo
                ' Set default selection
                cboVersion.AddItem "Latest Version"
                ' Default Latest Version to 10000
                cboVersion.ItemData(cboVersion.NewIndex) = 10000

                sSQL = "SELECT StudyVersion, VersionDescription FROM StudyVersion WHERE ClinicalTrialId = " & mlSelectedTrialId _
                    & " ORDER BY StudyVersion"
                .Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

                Do While Not .EOF
                    sComboVersion = CStr(.Fields(0))
                    ' if description not null then display
                    If Not IsNull(.Fields(1)) Then
                        sComboVersion = sComboVersion & " (" & ReplaceCharsForCombo(.Fields(1)) & ")"
                    End If
                    cboVersion.AddItem sComboVersion
                    cboVersion.ItemData(cboVersion.NewIndex) = .Fields(0)
                    .MoveNext
                Loop

                .Close

                cboVersion.ListIndex = 0
                lblWarning.Caption = "Please select a distributable study version"

            End With

        Case DisplayTrialsBySite

            lblWarning.Caption = "In this view the latest version of the studies will be distributed by default."
            cboVersion.Visible = False

    End Select


    Set rsCombo = Nothing

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SetupVersionCombo")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub tabAssocDistribute_Click(PreviousTab As Integer)
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    ' Check if version control combo needs to be shown
    Call VersionComboControlByTab

    ' refresh distribution list
    Select Case tabAssocDistribute.Tab
        Case 1
            Call RefreshDistributionList
        Case 2
            Call RefreshRecallList
        Case Else
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub VersionComboControlByTab()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

On Error GoTo ErrorHandler

    If cboVersion.Visible = True Or mbFormSetup = True Then
        If tabAssocDistribute.Tab = 0 Then
            cboVersion.Enabled = False
        Else
            cboVersion.Enabled = True
        End If
    End If

Exit Sub
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "VersionComboControlByTab")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Sub

'---------------------------------------------------------------------
Private Sub RefreshDistributionList()
'---------------------------------------------------------------------
' Show available distributable studies given Participating studies/sites
' and version chosen from combo (if applicable)
'---------------------------------------------------------------------
Dim lVersion As Long
Dim oLIMainListItem As ListItem
Dim oLIDistribute As ListItem
Dim nSubItems As Integer
Dim sSite As String
Dim lStudyId As Long
Dim sStudyName As String

On Error GoTo ErrorHandler

    ' get site chosen details from cboSelect (if appropriate)
    Select Case meDisplay
        Case DisplayTrialsBySite
            sSite = cboSelect.List(cboSelect.ListIndex)
        Case DisplaySitesByTrial
            lStudyId = GetcboSelectItemData(cboSelect.ListIndex)
            sStudyName = GetTrialNameFromId(lStudyId)
    End Select

    ' Collect selected version (if applicable)
    If cboVersion.Visible Then
        lVersion = cboVersion.ItemData(cboVersion.ListIndex)
        ' if 'Latest Version' get real latest version value
        If lVersion = 10000 Then
            lVersion = GetLatestVersionOfTrialAvailable(lStudyId)
        End If
    Else
        lVersion = 10000
    End If

    ' Clear listview
    lvwDistribute.ListItems.Clear

    ' Copy relevant columns from main listview to distribute listview
    For Each oLIMainListItem In lvwList.ListItems
        ' If is a participating study/site
        If oLIMainListItem.Checked Then
            Select Case meDisplay
                Case DisplayTrialsBySite
                    sStudyName = oLIMainListItem.Tag
                    lStudyId = CLng(oLIMainListItem.SubItems(1)) 'TrialIdFromName(sStudyName)
                    ' Get latest version number
                    lVersion = GetLatestVersionOfTrialAvailable(lStudyId)
                Case DisplaySitesByTrial
                    sSite = oLIMainListItem.Tag
            End Select
            ' Check Version
            If ShowVersionInListView(lVersion, RemoveNull(oLIMainListItem.SubItems(LatestVersionSubItemIndex))) _
                And GetSiteLocation(sSite) = SiteLocation.ESiteRemote Then
                ' Copy to Distribute listview
                Set oLIDistribute = lvwDistribute.ListItems.Add(, , oLIMainListItem.Text)
                ' Copy Subitems
                For nSubItems = 1 To oLIMainListItem.ListSubItems.Count
                    oLIDistribute.SubItems(nSubItems) = oLIMainListItem.SubItems(nSubItems)
                Next
                oLIDistribute.Tag = oLIMainListItem.Tag
                oLIDistribute.Key = oLIMainListItem.Key
            End If
        End If
    Next

    ' Enable Distribution button if something to select
    If lvwDistribute.ListItems.Count > 0 Then
        cmdDistribute.Enabled = True
    Else
        cmdDistribute.Enabled = False
    End If
    
    lvw_SetAllColWidths lvwDistribute, LVSCW_AUTOSIZE_USEHEADER
    'MLM 22/04/07: don't auto-size the new hidden column
    If meDisplay = DisplayTrialsBySite Then
        lvwDistribute.ColumnHeaders.Item(2).Width = 0
    End If
    
Exit Sub
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshDistributionList")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Sub

'--------------------------------------------------------------------------------
Private Function ShowVersionInListView(lSelectedVersion As Long, sListViewVersion As String) As Boolean
'--------------------------------------------------------------------------------
' Complete listview with History Detail
'--------------------------------------------------------------------------------
Dim lListViewVersion As Long

On Error GoTo ErrorHandler

    If sListViewVersion = "" Or sListViewVersion = "Latest" Then
        lListViewVersion = 0
    Else
        lListViewVersion = CLng(sListViewVersion)
    End If

    If lListViewVersion < lSelectedVersion Then
        ShowVersionInListView = True
    Else
        ShowVersionInListView = False
    End If

Exit Function
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ShowVersionInListView")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Function

'--------------------------------------------------------------------------------
Private Function GetSiteLocation(sSiteCode As String) As SiteLocation
'--------------------------------------------------------------------------------
' Get site detail (server/remote)
'--------------------------------------------------------------------------------
Dim rsSite As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrorHandler

    GetSiteLocation = SiteLocation.ESiteNoLocation

    sSQL = "SELECT SiteLocation FROM Site WHERE Site = '" & sSiteCode & "'"

    Set rsSite = New ADODB.Recordset
    With rsSite
        .Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText

        If Not rsSite.EOF Then
            If IsNull(rsSite(0).Value) Then
                GetSiteLocation = SiteLocation.ESiteNoLocation
            Else
                GetSiteLocation = rsSite(0).Value
            End If
        End If

        rsSite.Close

    End With
    Set rsSite = Nothing

Exit Function
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetSiteLocation")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Function

''--------------------------------------------------------------------------------
'Private Sub PopulateListViewWithHistoryDetail()
''--------------------------------------------------------------------------------
'' Complete listview with History Detail
''--------------------------------------------------------------------------------
'Dim oLVParticipationItem As ListItem
'Dim sSite As String
'Dim lStudyId As Long
'Dim sStudyName As String
'Dim rsMessage As ADODB.Recordset
'Dim sSQL As String
'Dim dblDistributeDate As Double
'Dim dblReceivedDate As Double
'Dim lDistVersion As Long
'Dim lReceivedVersion As Long
'Dim lTempVersion As Long
'Dim nFirstDataItem As Long
'Dim eSiteLocation As SiteLocation
'
'    On Error GoTo ErrHandler
'
'    ' get site/study chosen details from cboSelect
'    ' Depending on view subitems are different ...
'    Select Case meDisplay
'        Case DisplayTrialsBySite
'            sSite = cboSelect.List(cboSelect.ListIndex)
'            eSiteLocation = GetSiteLocation(sSite)
'            nFirstDataItem = 1
'        Case DisplaySitesByTrial
'            lStudyId = GetcboSelectItemData(cboSelect.ListIndex)
'            sStudyName = cboSelect.List(cboSelect.ListIndex)
'            nFirstDataItem = 2
'    End Select
'
'    ' Go through each listitem retrieving history detail
'    For Each oLVParticipationItem In lvwList.ListItems
'        ' get site/study details from listview item
'        Select Case meDisplay
'            Case DisplayTrialsBySite
'                sStudyName = oLVParticipationItem.Tag
'                lStudyId = TrialIdFromName(sStudyName)
'            Case DisplaySitesByTrial
'                sSite = oLVParticipationItem.Tag
'                ' set location of site in listview
'                eSiteLocation = GetSiteLocation(sSite)
'                Select Case eSiteLocation
'                    Case SiteLocation.ESiteServer
'                        oLVParticipationItem.SubItems(1) = "Server"
'                    Case SiteLocation.ESiteRemote
'                        oLVParticipationItem.SubItems(1) = "Remote"
'                End Select
'        End Select
'
'        ' DPH 03/09/2002 - set Server version to "Latest"
'        If eSiteLocation = SiteLocation.ESiteServer And oLVParticipationItem.Checked Then
'            oLVParticipationItem.SubItems(nFirstDataItem) = "Latest"
'        End If
'
'        ' Reset Working version variables
'        dblDistributeDate = 0
'        dblReceivedDate = 0
'        lDistVersion = 0
'        lReceivedVersion = 0
'
'        ' Get Message sent / received dates
'        sSQL = "SELECT MessageParameters, MessageTimeStamp, MessageReceived, MessageReceivedTimeStamp FROM Message WHERE TrialSite = '" & sSite & "'" _
'            & " AND ClinicalTrialId = " & lStudyId _
'            & " AND MessageType = 8" _
'            & " ORDER BY MessageTimeStamp DESC"
'
'        Set rsMessage = mconMACRO.Execute(sSQL, -1, adCmdText)
'        Do While Not rsMessage.EOF
'
'            ' if Version is set (number before parameter name...)
'            lTempVersion = GetStudyVersionFromParameterField(rsMessage("MessageParameters"), sStudyName)
'            If lTempVersion > 0 Then
'                ' If Version distributed and later than currently highest version in variables
'                If lDistVersion < lTempVersion Then
'                    lDistVersion = lTempVersion
'                    If Not IsNull(rsMessage("MessageTimeStamp")) Then
'                        dblDistributeDate = rsMessage("MessageTimeStamp")
'                    End If
'                End If
'
'                ' Check Version received and store if latest for site/study version
'                If Not IsNull(rsMessage("MessageReceivedTimeStamp")) Then
'                    If lReceivedVersion < lTempVersion And rsMessage("MessageReceivedTimeStamp") > 0 Then
'                        lReceivedVersion = lTempVersion
'                        dblReceivedDate = rsMessage("MessageReceivedTimeStamp")
'                    End If
'                End If
'            End If
'
'            rsMessage.MoveNext
'        Loop
'        rsMessage.Close
'        Set rsMessage = Nothing
'
'        ' Fill data row with detail
'        If lDistVersion > 0 Then
'            oLVParticipationItem.SubItems(nFirstDataItem) = lDistVersion
'            oLVParticipationItem.SubItems(nFirstDataItem + 1) = Format(dblDistributeDate, "yyyy/MM/dd hh:mm")
'        End If
'        If lReceivedVersion > 0 Then
'            oLVParticipationItem.SubItems(nFirstDataItem + 2) = lReceivedVersion
'            oLVParticipationItem.SubItems(nFirstDataItem + 3) = Format(dblReceivedDate, "yyyy/MM/dd hh:mm")
'        End If
'    Next
'
'Exit Sub
'ErrHandler:
'  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PopulateListViewWithHistoryDetail")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Unload frmMenu
'   End Select
'
'End Sub

'--------------------------------------------------------------------------------
Private Sub SaveParticipationDetail(Optional bChangeView As Boolean = False)
'--------------------------------------------------------------------------------
' Save Changed participation details
'--------------------------------------------------------------------------------
' REVISIONS
' DPH 13/01/2003 - If saved from a change then use 'previous' value
'--------------------------------------------------------------------------------
Dim nChanged As Integer
Dim lIndex As Long
Dim sSite As String
Dim sStudyName As String
Dim lStudyId As Long
Dim lVersion As Long
Dim bChecked As Boolean
Dim sSQL As String
Dim lUpdate As Long
Dim oLIChanged As ListItem
' DPH 13/01/2003 - If saved from a change then use 'previous' value

    On Error GoTo ErrorHandler
    
    If Not bChangeView Then
        Select Case meDisplay
            Case DisplayTrialsBySite
                sSite = cboSelect.List(cboSelect.ListIndex)
            Case DisplaySitesByTrial
                lStudyId = GetcboSelectItemData(cboSelect.ListIndex)
                sStudyName = GetTrialNameFromId(lStudyId)
        End Select
    Else
        Select Case meDisplay
            Case DisplayTrialsBySite
                sSite = msPrevSite
            Case DisplaySitesByTrial
                lStudyId = mlPrevTrialId
                sStudyName = msPrevStudyName
        End Select
    End If
    
    ' loop through collection of changed (or 'dirty') rows and write to db
    For nChanged = 1 To colChangedParticipation.Count
        ' extract changed listitem from collection
        Set oLIChanged = colChangedParticipation.Item(nChanged)
        
        ' use index to retrieve list data
        Select Case meDisplay
            Case DisplayTrialsBySite
                sStudyName = oLIChanged.Tag
                'lStudyId = TrialNameFromId(sStudyName)
                lStudyId = GetTrialIdFromTrialName(sStudyName)
            Case DisplaySitesByTrial
                sSite = oLIChanged.Tag
        End Select
        
        ' get version info (need to find distributed version)
        ' as constantly updated get from database
        lVersion = GetStudySiteLatestVersion(lStudyId, sStudyName, sSite)
        
        ' checked status
        bChecked = oLIChanged.Checked
        
        ' Now have all relevant data so save to TrialSite table
        If bChecked Then
            ' Insert or Update row
            ' try update
            sSQL = "UPDATE TrialSite SET StudyVersion = " & lVersion & _
                " WHERE TrialSite = '" & sSite _
                & "' AND ClinicalTrialId = " & lStudyId
            mconMACRO.Execute sSQL, lUpdate, adCmdText
            
            ' insert if no update
            If lUpdate = 0 Then
                sSQL = "INSERT INTO TrialSite (TrialSite,ClinicalTrialId,StudyVersion) " _
                    & "VALUES ('" & sSite & "'," & lStudyId & "," & lVersion & ")"
                mconMACRO.Execute sSQL, -1, adCmdText
                
                'REM 17/01/03 - Create status message in message table to distribute study
                CreateStatusMessage ExchangeMessageType.NewTrial, _
                                    lStudyId, _
                                    sStudyName, _
                                    sSite
            End If
        Else
            ' delete row
            sSQL = "DELETE FROM TrialSite WHERE TrialSite = '" & sSite _
                & "' AND ClinicalTrialId = " & lStudyId
            mconMACRO.Execute sSQL, -1, adCmdText
        End If
    Next
    
    'REM 21/05/03 - Update User Study/Site permissions
    goUser.ReloadStudySitePermissions
    
Exit Sub
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SaveParticipationDetail")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Sub

'--------------------------------------------------------------------------------
Private Function GetLatestVersionOfTrialAvailable(lTrialId As Long) As Integer
'--------------------------------------------------------------------------------
' Extract
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsVersion As ADODB.Recordset
Dim nVersion As Integer

    On Error GoTo ErrorHandler
    
    sSQL = "SELECT Max(StudyVersion) AS MaxVersion FROM StudyVersion WHERE ClinicalTrialId = " & lTrialId
    
    Set rsVersion = New ADODB.Recordset
    rsVersion.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    If Not rsVersion.EOF Then
        If IsNull(rsVersion("MaxVersion")) Then
            nVersion = 0
        Else
            nVersion = rsVersion("MaxVersion")
        End If
    Else
        nVersion = 0
    End If
    rsVersion.Close
    Set rsVersion = Nothing
    
    GetLatestVersionOfTrialAvailable = nVersion
    
Exit Function
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetLatestVersionOfTrialAvailable")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Function

'--------------------------------------------------------------------------------
Private Function LatestVersionSubItemIndex() As Integer
'--------------------------------------------------------------------------------
' Return latest version column depending on view
'--------------------------------------------------------------------------------

    ' Find latest version column number depending on view
    Select Case meDisplay
        Case eDisplayType.DisplayTrialsBySite
            ' Because of 'type' field
            LatestVersionSubItemIndex = 2 'MLM 23/04/07: Changed from 1 to 2 due to new hidden column.
        Case eDisplayType.DisplaySitesByTrial
            LatestVersionSubItemIndex = 2
        Case Else
            LatestVersionSubItemIndex = -1
    End Select

End Function

'--------------------------------------------------------------------------------
Private Function DistributeVersionMessage(lTrialId As Long, sTrialName As String, _
                        sSiteCode As String, lVersion As Long) As String
'--------------------------------------------------------------------------------
' Distribute version Message (checking distribution rules firstly)
'--------------------------------------------------------------------------------
' REVISIONS
' DPH 03/09/2002 - Changed to put version description in message
'--------------------------------------------------------------------------------
Dim sStudyFileName As String
Dim bDistribute As Boolean
Dim sSQL As String
Dim rsCheckMessages As ADODB.Recordset
Dim sMessageResult As String
Dim lStoredVersion As Long

    On Error GoTo ErrorHandler
    
    ' Set bDistribute to true
    bDistribute = True

    ' check distribution rules
    ' Go to database and check that no message with a higher version number has been sent
    ' should not happen (but may in multiuser environment)
    sSQL = "SELECT MessageId,MessageTimeStamp,MessageParameters,MessageReceived,MessageReceivedTimeStamp " _
        & " FROM Message WHERE TrialSite = '" & sSiteCode _
        & "' AND ClinicalTrialId = " & lTrialId & " AND MessageType = " & ExchangeMessageType.NewVersion _
        & " ORDER BY MessageTimeStamp"
    
    Set rsCheckMessages = New ADODB.Recordset
    rsCheckMessages.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly
    ' loop through study/site new versions looking for more recent version distribution
    Do While Not rsCheckMessages.EOF
        lStoredVersion = GetStudyVersionFromParameterField(rsCheckMessages("MessageParameters"), sTrialName)
        If lStoredVersion >= lVersion Then
            ' same or later version distributed so do not distribute
            bDistribute = False
            ' put together message to return from function
            sMessageResult = "Study " & sTrialName & " version " & lVersion _
                & " was not distributed to " & sSiteCode
            If rsCheckMessages("MessageReceivedTimeStamp") > 0 Then
                sMessageResult = sMessageResult & " as it was received by the site on " _
                    & Format(rsCheckMessages("MessageReceivedTimeStamp"), "yyyy/MM/dd hh:mm")
            Else
                sMessageResult = sMessageResult & " as it was made available for download to the site on " _
                    & Format(rsCheckMessages("MessageTimeStamp"), "yyyy/MM/dd hh:mm")
            End If
            ' exit loop as found failure
            Exit Do
        Else
            ' check if previous version was received, if not mark as superceeded
            If rsCheckMessages("MessageReceivedTimeStamp") = 0 Then
                ' DPH 04/09/2002 use MessageReceived.Superceeded status instead of 5
                sSQL = "UPDATE Message SET MessageReceived = " & MessageReceived.Superceeded & " WHERE MessageId = " & rsCheckMessages("MessageId")
                MacroADODBConnection.Execute sSQL, -1, adCmdText
            End If
        End If
        rsCheckMessages.MoveNext
    Loop
    rsCheckMessages.Close
    Set rsCheckMessages = Nothing
    
    ' distribute if OK
    If bDistribute Then
        sStudyFileName = lVersion & sTrialName & ".zip"
        
        ' Create Message
        Call CreateStatusMessage(ExchangeMessageType.NewVersion, lTrialId, _
                sTrialName, sSiteCode, sStudyFileName, ReplaceCharsForCombo(GetVersionDescription(lTrialId, lVersion)))
                
        ' complete status string
        sMessageResult = "Study " & sTrialName & " version " & lVersion _
            & " is available for " & sSiteCode & " to download"
    End If
    
    ' set function result
    DistributeVersionMessage = sMessageResult
    
Exit Function
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "DistributeVersionMessage")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Function

'--------------------------------------------------------------------------------
Private Function GetVersionDescription(lStudyId As Long, lVersion As Long) As String
'--------------------------------------------------------------------------------
' Collect study version description
'--------------------------------------------------------------------------------
Dim sVersionDescription As String
Dim sSQL As String
Dim rsVersion As ADODB.Recordset

    On Error GoTo ErrorHandler
    
    sVersionDescription = ""
    
    sSQL = "SELECT VersionDescription FROM StudyVersion WHERE ClinicalTrialId = " _
        & lStudyId & " AND StudyVersion = " & lVersion
    
    Set rsVersion = New ADODB.Recordset
    rsVersion.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly
    If Not rsVersion.EOF Then
        If Not IsNull(rsVersion("VersionDescription")) Then
            sVersionDescription = rsVersion("VersionDescription")
        End If
    End If
    rsVersion.Close
    Set rsVersion = Nothing
    
    GetVersionDescription = sVersionDescription
Exit Function
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetVersionDescription")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Function

'--------------------------------------------------------------------------------
Private Function GetStudySiteLatestVersion(lStudyId As Long, sStudyName As String, sSite As String) As Long
'--------------------------------------------------------------------------------
' Collect latest version number distributed to study/site
'--------------------------------------------------------------------------------
Dim lVersion As Long
Dim lTempVersion As Long
Dim sSQL As String
Dim rsVersion As ADODB.Recordset

    On Error GoTo ErrorHandler
    
    lVersion = 0
    lTempVersion = 0
    
    sSQL = "SELECT MessageParameters " _
        & " FROM Message WHERE TrialSite = '" & sSite _
        & "' AND ClinicalTrialId = " & lStudyId & " AND MessageType = " & ExchangeMessageType.NewVersion _
        & " ORDER BY MessageTimeStamp DESC"

    Set rsVersion = New ADODB.Recordset
    rsVersion.Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly
    Do While Not rsVersion.EOF
        If Not IsNull(rsVersion("MessageParameters")) Then
            lTempVersion = GetStudyVersionFromParameterField(rsVersion("MessageParameters"), sStudyName)
            If lTempVersion > lVersion Then
                lVersion = lTempVersion
            End If
        End If
        rsVersion.MoveNext
    Loop
    rsVersion.Close
    Set rsVersion = Nothing
    
    GetStudySiteLatestVersion = lVersion
    
Exit Function
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetStudySiteLatestVersion")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Function

'--------------------------------------------------------------------------------
Private Function GetcboSelectItemData(lIndex As Long) As Long
'--------------------------------------------------------------------------------
' Used to check itemdat in empty combos
'--------------------------------------------------------------------------------
On Error GoTo DefaultItemData
    If lIndex > -1 Then
        GetcboSelectItemData = cboSelect.ItemData(lIndex)
    End If
Exit Function
DefaultItemData:
    GetcboSelectItemData = 0
End Function

'--------------------------------------------------------------------------------
Private Function ReplaceCharsForCombo(sText As String) As String
'--------------------------------------------------------------------------------
' Replace Characters like New line etc as putting in a combo line
'--------------------------------------------------------------------------------
Dim nChar As Integer
Dim sModText As String

    sModText = sText
    For nChar = 1 To 31
        sModText = Replace(sModText, Chr(nChar), " ")
    Next
    
    ReplaceCharsForCombo = sModText
End Function

'---------------------------------------------------------------------
Private Sub RefreshRecallList()
'---------------------------------------------------------------------
' Show available recallable studies given Participating studies
'---------------------------------------------------------------------
Dim oLIMainListItem As ListItem
Dim oLIRecall As ListItem
Dim nSubItems As Integer
Dim sSite As String
Dim lStudyId As Long
Dim sStudyName As String

On Error GoTo ErrorHandler

    ' get site chosen details from cboSelect (if appropriate)
    Select Case meDisplay
        Case DisplayTrialsBySite
            sSite = cboSelect.List(cboSelect.ListIndex)
        Case DisplaySitesByTrial
            lStudyId = GetcboSelectItemData(cboSelect.ListIndex)
            sStudyName = GetTrialNameFromId(lStudyId)
    End Select

    ' Clear listview
    lvwRecall.ListItems.Clear

    ' Copy relevant columns from main listview to recall listview
    For Each oLIMainListItem In lvwList.ListItems
        ' If is a participating study/site
        If oLIMainListItem.Checked Then
            Select Case meDisplay
                Case DisplayTrialsBySite
                    sStudyName = oLIMainListItem.Tag
                    lStudyId = CLng(oLIMainListItem.SubItems(1))
                Case DisplaySitesByTrial
                    sSite = oLIMainListItem.Tag
            End Select
            ' Check if distributed version > receieved version
            If ShowRecallInListView(RemoveNull(oLIMainListItem.SubItems(LatestVersionSubItemIndex)), _
                RemoveNull(oLIMainListItem.SubItems(LatestVersionSubItemIndex + 2))) _
                And GetSiteLocation(sSite) = SiteLocation.ESiteRemote Then
                ' Copy to Recall listview
                Set oLIRecall = lvwRecall.ListItems.Add(, , oLIMainListItem.Text)
                ' Copy Subitems
                For nSubItems = 1 To oLIMainListItem.ListSubItems.Count
                    oLIRecall.SubItems(nSubItems) = oLIMainListItem.SubItems(nSubItems)
                Next
                oLIRecall.Tag = oLIMainListItem.Tag
                oLIRecall.Key = oLIMainListItem.Key
            End If
        End If
    Next

    ' Enable Recall button if something to select
    If lvwRecall.ListItems.Count > 0 Then
        cmdRecall.Enabled = True
    Else
        cmdRecall.Enabled = False
    End If
    
    lvw_SetAllColWidths lvwRecall, LVSCW_AUTOSIZE_USEHEADER
    'MLM 22/04/07: don't auto-size the new hidden column
    If meDisplay = DisplayTrialsBySite Then
        lvwRecall.ColumnHeaders.Item(2).Width = 0
    End If
    
Exit Sub
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshRecallList")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Sub

'--------------------------------------------------------------------------------
Private Function ShowRecallInListView(sDistributeListViewVersion As String, sReceivedListViewVersion As String) As Boolean
'--------------------------------------------------------------------------------
' Check if Should display row in Recall tab
'--------------------------------------------------------------------------------
Dim lDistListViewVersion As Long
Dim lReceivedListViewVersion As Long

On Error GoTo ErrorHandler

    If sDistributeListViewVersion = "" Or sDistributeListViewVersion = "Latest" Then
        lDistListViewVersion = 0
    Else
        lDistListViewVersion = CLng(sDistributeListViewVersion)
    End If

    If sReceivedListViewVersion = "" Or sReceivedListViewVersion = "Latest" Then
        lReceivedListViewVersion = 0
    Else
        lReceivedListViewVersion = CLng(sReceivedListViewVersion)
    End If

    If lDistListViewVersion > lReceivedListViewVersion Then
        ShowRecallInListView = True
    Else
        ShowRecallInListView = False
    End If

Exit Function
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ShowRecallInListView")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Function

'--------------------------------------------------------------------------------
Private Function RecallVersionMessage(lStudyId As Long, sStudyName As String, _
                                sSiteCode As String, lVersion As Long) As String
'--------------------------------------------------------------------------------
' Recall Version Message by deleting it and noting deletion to system log
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim lRecsAffected As Long
Dim sMessage As String

    On Error GoTo ErrorHandler
    
    ' Set up SQL to Delete Recalled row
    sSQL = "DELETE FROM Message WHERE ClinicalTrialId = " & lStudyId & " AND TrialSite = '" _
            & sSiteCode & "' AND MessageParameters = '" & lVersion & sStudyName & ".zip' " _
            & "AND MessageType = " & ExchangeMessageType.NewVersion & " AND MessageReceived = " _
            & MessageReceived.NotYetReceived
    ' Execute Delete
    MacroADODBConnection.Execute sSQL, lRecsAffected, adCmdText
    ' Check if any records affected by SQL call
    If lRecsAffected = 0 Then
        sMessage = "Unable to remove Message to distribute study " _
                & sStudyName & " (v" & lVersion & ") to site " & sSiteCode & " that failed"
        ' Set Message in system log
        gLog "StudyDistribution", sMessage
        ' Set return string
        RecallVersionMessage = sMessage
    Else
        sMessage = "Removed Message to distribute study " & sStudyName & " (v" & lVersion _
                & ") to site " & sSiteCode
        ' Set Message in system log
        gLog "StudyDistribution", sMessage
        ' Reset previos message if it has been set to Superceeded
        Call UnmarkPreviousSuperceededMessage(lStudyId, sSiteCode)
        ' Set return string
        RecallVersionMessage = sMessage
    End If

Exit Function
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RecallVersionMessage")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Function

'--------------------------------------------------------------------------------
Private Sub UnmarkPreviousSuperceededMessage(lStudyId As Long, sSiteCode As String)
'--------------------------------------------------------------------------------
' Find previous message & if superceeded unmark to be ready to receive
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsMessage As ADODB.Recordset

    ' Get Message sent / received dates
    sSQL = "SELECT MessageId, MessageReceived FROM Message WHERE TrialSite = '" & sSiteCode & "'" _
        & " AND ClinicalTrialId = " & lStudyId _
        & " AND MessageType = " & ExchangeMessageType.NewVersion _
        & " ORDER BY MessageTimeStamp DESC"

    Set rsMessage = New ADODB.Recordset
    With rsMessage
        .Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly
    End With

    ' check 1st record (last version distributed)
    If Not rsMessage.EOF Then
        ' If the Message is superceeded then reset to not received
        If rsMessage("MessageReceived") = MessageReceived.Superceeded Then
            sSQL = "UPDATE Message SET MessageReceived = " & MessageReceived.NotYetReceived _
                    & " WHERE MessageId = " & rsMessage("MessageId")
            MacroADODBConnection.Execute sSQL, -1, adCmdText
        End If
    End If
    
    rsMessage.Close
    Set rsMessage = Nothing
    
Exit Sub
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "UnmarkPreviousSuperceededMessage")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Sub

'---------------------------------------------------------------------
Private Function GetTrialNameFromId(ByVal lClinicalTrialId As Long) As String
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT ClinicalTrialName FROM ClinicalTrial " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, mconMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        GetTrialNameFromId = ""
    Else
        GetTrialNameFromId = rsTemp!ClinicalTrialName
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetTrialNameFromId", "TrialData")
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
Private Function GetTrialIdFromTrialName(ByVal sTrialName As String) As Long
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT ClinicalTrialId FROM ClinicalTrial " _
        & " WHERE ClinicalTrialName = '" & sTrialName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, mconMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        GetTrialIdFromTrialName = -1
    Else
        GetTrialIdFromTrialName = rsTemp!ClinicalTrialId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetTrialNameFromId", "TrialData")
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
Private Sub StorePreviousStudySite()
'---------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Select Case meDisplay
        Case DisplayTrialsBySite
            msPrevSite = cboSelect.List(cboSelect.ListIndex)
        Case DisplaySitesByTrial
            mlPrevTrialId = GetcboSelectItemData(cboSelect.ListIndex)
            msPrevStudyName = GetTrialNameFromId(mlPrevTrialId)
    End Select

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "StorePreviousStudySite")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Sub
