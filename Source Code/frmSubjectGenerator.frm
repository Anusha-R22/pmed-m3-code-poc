VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSubjectGenerator 
   Caption         =   "MACRO Subject Generator"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   5670
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraEntryControls 
      Caption         =   "Specify Study, Site and Numbers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   120
      TabIndex        =   6
      Top             =   3180
      Width           =   5475
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   2820
         TabIndex        =   16
         Top             =   1140
         Width           =   1200
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   4140
         TabIndex        =   15
         Top             =   1140
         Width           =   1200
      End
      Begin VB.ComboBox cboStudy 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   730
         Width           =   1500
      End
      Begin VB.ComboBox cboSite 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   730
         Width           =   1500
      End
      Begin VB.TextBox txtNextId 
         Height          =   315
         Left            =   3240
         TabIndex        =   8
         Top             =   730
         Width           =   1000
      End
      Begin VB.TextBox txtNumberOfSubjects 
         Height          =   315
         Left            =   4320
         MaxLength       =   7
         TabIndex        =   7
         Top             =   720
         Width           =   1000
      End
      Begin VB.Label lblStudy 
         Caption         =   "Study"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label lblSite 
         Caption         =   "Site"
         Height          =   195
         Left            =   1740
         TabIndex        =   13
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label lblNextId 
         Caption         =   "Next ID"
         Height          =   195
         Left            =   3240
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblNumberOfSubjects 
         Caption         =   "Number of Subjects"
         Height          =   375
         Left            =   4320
         TabIndex        =   11
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   315
      Left            =   3060
      TabIndex        =   5
      Top             =   5220
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Top             =   5220
      Width           =   1200
   End
   Begin VB.TextBox txtProgress 
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   5220
      Width           =   2500
   End
   Begin VB.Frame fraStudySiteList 
      Caption         =   "Subjects to be Generated"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   5475
      Begin MSComctlLib.ListView lvwStudySiteList 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
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
   Begin VB.Label lblProgress 
      Caption         =   "Progress"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4980
      Width           =   1335
   End
End
Attribute VB_Name = "frmSubjectGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         frmSubjectGenerator.frm
' Copyright:    InferMed Ltd October 2002
' Purpose:      USED to Generate New Subjects
'               Only used within MACRO_BD
'               called from frmMenu.mnuFGenerateSubjects_Click
'----------------------------------------------------------------------------------------'
'   Revisions:
'TA 28/10/2003 choose which country from the site when creating a new subject
' NCJ 6 May 03 - Added UserNameFull and UserRole to NewSubject
' NCJ 1 July 04 - Check returned lock token (Bug 2314)
' Mo 17/10/2007 - Bug 2875, Prevent the generation of subjects for studies with status "suspended",
'                 "closed to followup" or "closed to recruitment"
'                 In addition restrict the list of studies that the user can choose from
'                 to studies that the user has permissions for.
'                 Changes made to cboStudy_Click and LoadStudyCombo
' Mo 19/10/2007 - Bug 2691, Prevent the generation of subjects for remote sites unless the
'                 user is at the remote site.
'                 Change LoadSiteCombo so that it only offers remote site if you are at that remote site's database.
'                 In addition restrict the list of sites that the user can choose from
'                 to sites that the user has permissions for.
'                 Both of the above are achieved by LoadSiteCombo now calling MACROUser.GetNewSubjectSites
'----------------------------------------------------------------------------------------'

Option Explicit

Private mlSelTrialId As Long
Private msSelSite As String

Private mcolStudySites As Collection

'--------------------------------------------------------------------
Private Sub cboSite_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler

    'exit if no item is currently selected
    If cboSite.ListIndex = -1 Then Exit Sub
    
    'exit if currently selected item matches msSelSite
    If cboSite.Text = msSelSite Then Exit Sub
    
    'Store selected ClinicalTrialId
    msSelSite = cboSite.Text
    
    'Display the next available PersonId for selected Study/Site
    txtNextId = NextSubjectId
    
    'Clear down any previously entered NumberOfSubjects and make sure its enabled
    txtNumberOfSubjects.Text = ""
    txtNumberOfSubjects.Enabled = True
    
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
    
    'Mo 17/10/2007 Bug 2875
    'Check the status of the selected study
    Select Case GetStudyStatusId(cboStudy.ItemData(cboStudy.ListIndex))
    Case eTrialStatus.ClosedToRecruitment
        Call DialogError("This study is Closed to Recruitment." & vbCrLf & "You cannot generate new subjects for this study.", "Study Closed to Recruitment")
        cboStudy.ListIndex = -1
        Exit Sub
    Case eTrialStatus.ClosedToFollowUp
        Call DialogError("This study is Closed to Follow Up." & vbCrLf & "You cannot generate new subjects for this study.", "Study Closed to Follow Up")
        cboStudy.ListIndex = -1
        Exit Sub
    Case eTrialStatus.Suspended
        Call DialogError("This study is Suspended." & vbCrLf & "You cannot generate new subjects for this study.", "Study Suspended")
        cboStudy.ListIndex = -1
        Exit Sub
    End Select
    
    'exit if currently selected item matches mlSelTrialId
    If cboStudy.ItemData(cboStudy.ListIndex) = mlSelTrialId Then Exit Sub
    
    'Store selected ClinicalTrialId
    mlSelTrialId = cboStudy.ItemData(cboStudy.ListIndex)
    
    'Having selected a study enable and populate cboSite
    msSelSite = ""
    cboSite.Enabled = True
    Call LoadSiteCombo
    'and clear NextId and NumberOfSubjects text boxes
    txtNextId.Text = ""
    txtNumberOfSubjects.Text = ""
    txtNumberOfSubjects.Enabled = False
    
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
Private Sub cmdAdd_Click()
'--------------------------------------------------------------------
' NCJ 1 Jul 04 - Check the lock token!
'---------------------------------------------------------------------
Dim itmX As MSComctlLib.ListItem
Dim i As Integer
Dim sToken As String
Dim sErrMsg As String

    On Error GoTo ErrHandler
    
    'Place a lock on new subjects being generated for the selected Study & Site
    ' NCJ 1 Jul 04 - added error message argument
    sToken = LockSubjectGeneration(mlSelTrialId, cboSite.Text, sErrMsg)
    
    ' NCJ 1 Jul 04 - Cannot continue if we didn't get a token!
    If sToken = "" Then
        ' lock failed
        DialogError "Unable to add subjects." & vbCrLf & sErrMsg
        Exit Sub
    End If
    
    'Add the selected Study & Site to the listview
    Set itmX = lvwStudySiteList.ListItems.Add(, , cboStudy.Text)
    itmX.SubItems(1) = cboSite.Text
    itmX.SubItems(2) = CLng(txtNextId.Text)
    itmX.SubItems(3) = CLng(txtNextId.Text) + CLng(txtNumberOfSubjects.Text) - 1
    itmX.SubItems(4) = cboStudy.ItemData(cboStudy.ListIndex)
    
    'Make sure last entry is visible
    lvwStudySiteList.ListItems(itmX.Index).EnsureVisible

    'Set the Max Column Widths
    For i = 1 To 4
        Call lvw_SetColWidth(lvwStudySiteList, i, LVSCW_AUTOSIZE_USEHEADER)
    Next i
    
    'unselect the entered item
    lvwStudySiteList.SelectedItem.Selected = False
    
    'add the SudySite combination to the mcolStudySites collection
    mcolStudySites.Add sToken, cboStudy.Text & "|" & cboSite.Text

    'Having added Study/Site/First/Last to lvwStudySiteList clear down the current entry controls
    Call ClearEntryControls
    
    'enable the Generate button
    cmdGenerate.Enabled = True
    
    'disable the Delete button that might have been enabled
    cmdDelete.Enabled = False
    

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
Private Sub cmdDelete_Click()
'--------------------------------------------------------------------
Dim sToken As String

    On Error GoTo ErrHandler
    
    'retrieve the Lock token for this Study/Site entry and unlock Subject generation
    sToken = mcolStudySites.Item(lvwStudySiteList.SelectedItem.Text & "|" & lvwStudySiteList.SelectedItem.SubItems(1))
    Call UnLockSubjectGeneration(lvwStudySiteList.SelectedItem.SubItems(4), lvwStudySiteList.SelectedItem.SubItems(1), sToken)

     'Remove from the StudySites collection
     mcolStudySites.Remove lvwStudySiteList.SelectedItem.Text & "|" & lvwStudySiteList.SelectedItem.SubItems(1)
     
     'remove from the Listview
    lvwStudySiteList.ListItems.Remove (lvwStudySiteList.SelectedItem.Index)
    
    'Disable the delete button
    cmdDelete.Enabled = False
    
    'Disable the generate button if no entries remain in lvwStudySiteList
    If lvwStudySiteList.ListItems.Count = 0 Then
        cmdGenerate.Enabled = False
    End If
    
    'Clear down the entry controls, which might contain out of date study/site information
    Call ClearEntryControls

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
Private Sub cmdExit_Click()
'--------------------------------------------------------------------
Dim itmX As MSComctlLib.ListItem
Dim sStudyName As String
Dim sSite As String
Dim lClinicalTrialId As Long
Dim sToken As String

    On Error GoTo ErrHandler
    
    'Remove any Subject generation locks that might have been created, but never used (and removed)
    'For each study/site combination in lvwStudySiteList
    For Each itmX In lvwStudySiteList.ListItems
        sStudyName = itmX.Text
        sSite = itmX.SubItems(1)
        lClinicalTrialId = itmX.SubItems(4)
        
        'retrieve the Lock token for this Study/Site entry and remove the Subject Generation Lock
        sToken = mcolStudySites.Item(sStudyName & "|" & sSite)
        Call UnLockSubjectGeneration(lClinicalTrialId, sSite, sToken)

        'Remove from the StudySites collection
        mcolStudySites.Remove sStudyName & "|" & sSite
         
    Next itmX
    
    Unload Me

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdExit_Click")
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
Private Sub cmdGenerate_Click()
'--------------------------------------------------------------------
Dim itmX As MSComctlLib.ListItem
Dim lClinicalTrialId As Long
Dim sStudyName As String
Dim sSite As String
Dim lLastSubjectId As Long
Dim lLastGeneratedId As Long
Dim oStudyDef As StudyDefRO
Dim oSubject As StudySubject
Dim sToken As String
Dim sCountry As String



    On Error GoTo ErrHandler
    
    Call HourglassOn

    'For each study/site combination in lvwStudySiteList
    For Each itmX In lvwStudySiteList.ListItems
        sStudyName = itmX.Text
        sSite = itmX.SubItems(1)
        lLastSubjectId = itmX.SubItems(3)
        lClinicalTrialId = itmX.SubItems(4)
        
        'Create the required Study object
        Set oStudyDef = New StudyDefRO
        oStudyDef.Load gsADOConnectString, lClinicalTrialId, 1, goArezzo
        
        'Generate new subjects until lLastSubjectId is reached
        
        'TA 28/10/2003 choose which country we want
        sCountry = goUser.GetAllSites.Item(sSite).CountryName
        
        Do
            ' NCJ 6 May 03 - Added UserNameFull and UserRole
            Set oSubject = oStudyDef.NewSubject(sSite, goUser.UserName, sCountry, goUser.UserNameFull, goUser.UserRole)
            txtProgress.Text = sStudyName & " " & sSite & " " & oSubject.PersonId & " Generated"
            txtProgress.Refresh
            lLastGeneratedId = oSubject.PersonId
            oStudyDef.RemoveSubject
        Loop Until lLastGeneratedId >= lLastSubjectId
        
        'Clean up
        oStudyDef.Terminate
        Set oStudyDef = Nothing
        Set oSubject = Nothing
        
        'retrieve the Lock token for this Study/Site entry and remove the Subject Generation Lock
        sToken = mcolStudySites.Item(sStudyName & "|" & sSite)
        Call UnLockSubjectGeneration(lClinicalTrialId, sSite, sToken)

        'Remove from the StudySites collection
        mcolStudySites.Remove sStudyName & "|" & sSite
         
    Next itmX
    
    'Clear the contents of the listview
    lvwStudySiteList.ListItems.Clear
    
    'Disable the Generate button
    cmdGenerate.Enabled = False
    
    Call HourglassOff

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdGenerate_Click")
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
    
    Me.Icon = frmMenu.Icon
    
    'set initial size of form
    Me.Width = gnMINFORMWIDTHGENERATOR
    Me.Height = gnMINFORMHEIGHTGENERATOR
    
    'clear listview
    lvwStudySiteList.ListItems.Clear
    'add column headers with widths that are re-calculated by auto sizing

    Set colmX = lvwStudySiteList.ColumnHeaders.Add(, , "Study", 10)
    Set colmX = lvwStudySiteList.ColumnHeaders.Add(, , "Site", 10)
    Set colmX = lvwStudySiteList.ColumnHeaders.Add(, , "First Id", 10)
    Set colmX = lvwStudySiteList.ColumnHeaders.Add(, , "Last Id", 10)
    'trialId stored in a hidden column
    Set colmX = lvwStudySiteList.ColumnHeaders.Add(, , "TrialId", 0)
    
    'Auto size the column headers (1 to 4)
    For i = 1 To 4
        Call lvw_SetColWidth(lvwStudySiteList, i, LVSCW_AUTOSIZE_USEHEADER)
    Next i
    
    'Set initial disabled settings
    mlSelTrialId = 0
    msSelSite = ""
    cboSite.Enabled = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    txtNextId.Enabled = False
    txtNumberOfSubjects.Enabled = False
    cmdGenerate.Enabled = False
    
    'Clear the selected StudySites collections
    Set mcolStudySites = Nothing
    Set mcolStudySites = New Collection
    
    Call LoadStudyCombo
    
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
Private Sub LoadSiteCombo()
'--------------------------------------------------------------------
'Mo 19/10/2007 - Bug 2691, rewritten, only adds sites that user has permissions for
'--------------------------------------------------------------------
Dim colSites As Collection
Dim oSite As Site
Dim sVar As String

    On Error GoTo ErrHandler
    
    'Clear current contents of cboSite
    cboSite.Clear
    
    Set colSites = goUser.GetNewSubjectSites(mlSelTrialId)

    For Each oSite In colSites
        'check that the site in combination with the selected Study has not
        'already been added to the listview of subjects to be generated.
        'If it has don't add it to the combo
        On Error Resume Next
        sVar = mcolStudySites.Item(cboStudy.Text & "|" & oSite.Site)
        If Err.Number <> 0 Then
            'clear the not in collection error
            Err.Clear
            'and add site to combo
            cboSite.AddItem oSite.Site
        End If
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
Private Function NextSubjectId() As Long
'--------------------------------------------------------------------
Dim lPersonId As Long
Dim sSQL As String
Dim rsPerson As ADODB.Recordset

    On Error GoTo ErrHandler
    
    'Get the next PersonId for the selected Study/Site
    sSQL = "SELECT MAX(PersonId) as MaxPersonId " _
            & " FROM TrialSubject " _
            & " WHERE ClinicalTrialId = " & mlSelTrialId _
            & " AND TrialSite = '" & msSelSite & "'"
    
    Set rsPerson = New ADODB.Recordset
    rsPerson.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If IsNull(rsPerson!MaxPersonId) Then
        lPersonId = 1
    Else
        lPersonId = rsPerson!MaxPersonId + 1
    End If
    
    rsPerson.Close
    Set rsPerson = Nothing
    
    NextSubjectId = lPersonId

Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "NextSubjectId")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'--------------------------------------------------------------------
Private Sub Form_Resize()
'--------------------------------------------------------------------
Dim lFormWidth As Long
Dim lFormHeight As Long
Dim l10th As Long

    On Error GoTo ErrHandler
    
    If Me.Width < gnMINFORMWIDTHGENERATOR Then
        Me.Width = gnMINFORMWIDTHGENERATOR
    End If

    If Me.Height < gnMINFORMHEIGHTGENERATOR Then
        Me.Height = gnMINFORMHEIGHTGENERATOR
    End If
    
    lFormWidth = Me.ScaleWidth
    lFormHeight = Me.ScaleHeight
    
    fraStudySiteList.Left = 100
    fraStudySiteList.Top = 0
    fraStudySiteList.Width = lFormWidth - 200
    fraStudySiteList.Height = lFormHeight - 2600
    
    lvwStudySiteList.Left = 100
    lvwStudySiteList.Top = 300
    lvwStudySiteList.Width = fraStudySiteList.Width - 200
    lvwStudySiteList.Height = fraStudySiteList.Height - 400
    
    fraEntryControls.Left = 100
    fraEntryControls.Top = fraStudySiteList.Height + 200
    fraEntryControls.Width = fraStudySiteList.Width
    fraEntryControls.Height = 1600
    
    'The width of the controls within fraEntryControls is worked out as follows:-
    'There are 4 controls (that makes 5 gaps of 100 twips between them)
    'The width is split into 10 parts, the 2 combos get a width of 3 parts
    'and the text boxes get a width of 2 parts
    l10th = (fraEntryControls.Width - 500) / 10
    
    lblStudy.Left = 100
    lblStudy.Top = 480
    lblStudy.Width = l10th * 3
    cboStudy.Left = 100
    cboStudy.Top = 730
    cboStudy.Width = l10th * 3
    
    lblSite.Left = lblStudy.Left + lblStudy.Width + 100
    lblSite.Top = 480
    lblSite.Width = l10th * 3
    cboSite.Left = lblSite.Left
    cboSite.Top = 730
    cboSite.Width = l10th * 3
    
    lblNextId.Left = lblSite.Left + lblSite.Width + 100
    lblNextId.Top = 480
    lblNextId.Width = l10th * 2
    txtNextId.Left = lblNextId.Left
    txtNextId.Top = 730
    txtNextId.Width = l10th * 2
    
    lblNumberOfSubjects.Left = lblNextId.Left + lblNextId.Width + 100
    lblNumberOfSubjects.Top = 300
    lblNumberOfSubjects.Width = l10th * 2
    txtNumberOfSubjects.Left = lblNumberOfSubjects.Left
    txtNumberOfSubjects.Top = 730
    txtNumberOfSubjects.Width = l10th * 2
    
    cmdDelete.Left = fraEntryControls.Width - cmdDelete.Width - 100
    cmdDelete.Top = txtNumberOfSubjects.Top + txtNumberOfSubjects.Height + 100
    cmdAdd.Left = cmdDelete.Left - cmdAdd.Width - 100
    cmdAdd.Top = cmdDelete.Top
    
    lblProgress.Left = 100
    lblProgress.Top = fraEntryControls.Top + fraEntryControls.Height + 100
    lblProgress.Width = lFormWidth - 2800
    txtProgress.Left = 100
    txtProgress.Top = lblProgress.Top + 250
    txtProgress.Width = lblProgress.Width
    cmdExit.Left = lFormWidth - cmdExit.Width - 100
    cmdExit.Top = txtProgress.Top
    cmdGenerate.Left = cmdExit.Left - cmdGenerate.Width - 100
    cmdGenerate.Top = txtProgress.Top

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
Private Sub lvwStudySiteList_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If lvwStudySiteList.ListItems.Count > 0 Then
        'an item has been selected so enable the delete button
        cmdDelete.Enabled = True
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lvwStudySiteList_Click")
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
Private Sub lvwStudySiteList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call lvw_Sort(lvwStudySiteList, ColumnHeader)

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lvwStudySiteList_ColumnClick")
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
Private Sub txtNumberOfSubjects_Change()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If txtNumberOfSubjects.Text <> "" Then
        If ((Not gblnValidString(txtNumberOfSubjects.Text, valNumeric) Or Len(txtNumberOfSubjects.Text) > 6)) Then
            'The entered Number of Subjects is invalid
            Call DialogError("Invalid Number of Subjects." & vbCrLf & "Enter a number between 1 and 999,999.", "Invalid Number of Subjects")
            txtNumberOfSubjects.Text = ""
        Else
            'A valid Number of Subjects has been entered, enable the add button
            cmdAdd.Enabled = True
        End If
    Else
        'disable the Add button
        cmdAdd.Enabled = False
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtNumberOfSubjects_Change")
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
' NCJ 1 Jul 04 - We don't need to load the study def here!
' Added byref argument for failure error message
'--------------------------------------------------------------------
'Dim oStudyDef As StudyDefRO
Dim sToken As String

    On Error GoTo ErrLabel

'    Set oStudyDef = New StudyDefRO
'    oStudyDef.Load gsADOConnectString, lClinicalTrialId, 1, goArezzo

    'Calling LockSubject with NULL_INTEGER as the SubjectId has the effect
    'of locking New Subject Generation for the specified Study and Site
    sToken = LockSubject(goUser.UserName, lClinicalTrialId, sSite, NULL_INTEGER, sMessage)
    
'    Set oStudyDef = Nothing
    
    LockSubjectGeneration = sToken

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmSubjectGenerator.LockSubjectGeneration"

End Function

'--------------------------------------------------------------------
Private Sub UnLockSubjectGeneration(ByVal lClinicalTrialId As Long, _
                                        ByVal sSite As String, _
                                        ByVal sToken As String)
'--------------------------------------------------------------------
' NCJ 1 Jul 04 - We don;t need to load the study def here!
'--------------------------------------------------------------------
'Dim oStudyDef As StudyDefRO

    On Error GoTo ErrLabel

'    Set oStudyDef = New StudyDefRO
'    oStudyDef.Load gsADOConnectString, lClinicalTrialId, 1, goArezzo
    
    Call UnlockSubject(lClinicalTrialId, sSite, NULL_INTEGER, sToken)
    
'    Set oStudyDef = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmSubjectGenerator.UnlockSubjectGeneration"

End Sub

'--------------------------------------------------------------------
Private Sub ClearEntryControls()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'Clear down the Study/site/Next/Number of entry controls
    cboStudy.ListIndex = -1
    mlSelTrialId = 0
    cboSite.Enabled = False
    cboSite.Clear
    cboSite.ListIndex = -1
    msSelSite = ""
    txtNextId = ""
    txtNumberOfSubjects.Text = ""
    txtNumberOfSubjects.Enabled = False

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ClearEntryControls")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub
