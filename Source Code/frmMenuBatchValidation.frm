VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   Caption         =   "MACRO Batch Validation"
   ClientHeight    =   7605
   ClientLeft      =   3045
   ClientTop       =   4020
   ClientWidth     =   8385
   Icon            =   "frmMenuBatchValidation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   8385
   Begin VB.Frame fraForm 
      Caption         =   "eForm"
      Height          =   675
      Left            =   60
      TabIndex        =   19
      Top             =   1680
      Width           =   3195
      Begin VB.ComboBox cboEForm 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame fraRevalidation 
      Caption         =   "Revalidation"
      Height          =   3435
      Left            =   60
      TabIndex        =   12
      Top             =   3300
      Width           =   8235
      Begin VB.CheckBox chkSaveChanges 
         Caption         =   "Save changes"
         Height          =   255
         Left            =   4080
         TabIndex        =   22
         ToolTipText     =   "If ticked, all changes are automatically saved. If unticked, changes are reported but NOT saved."
         Top             =   310
         Width           =   1452
      End
      Begin VB.CheckBox chkSingleForm 
         Caption         =   "Single eForm only"
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   310
         Width           =   1695
      End
      Begin VB.CheckBox chkChangesOnly 
         Caption         =   "Log changes only"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Show only changes in the log file"
         Top             =   310
         Width           =   1575
      End
      Begin VB.CommandButton cmdRevalidate 
         Caption         =   "Revalidate"
         Height          =   345
         Left            =   5760
         TabIndex        =   17
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   6960
         TabIndex        =   16
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox txtMsg 
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "frmMenuBatchValidation.frx":08CA
         Top             =   720
         Width           =   8000
      End
      Begin MSComctlLib.ProgressBar pbBar 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   8000
         _ExtentX        =   14102
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgress 
         Caption         =   "Completed 0 of 100"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3060
         Width           =   2775
      End
   End
   Begin VB.Frame fraSubj 
      Caption         =   "Subject"
      Height          =   675
      Left            =   60
      TabIndex        =   9
      Top             =   2400
      Width           =   3195
      Begin VB.TextBox txtSubject 
         Height          =   315
         Left            =   120
         MaxLength       =   255
         TabIndex        =   10
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame fraSubjects 
      Caption         =   "Subjects"
      Height          =   3170
      Left            =   3360
      TabIndex        =   6
      Top             =   60
      Width           =   4935
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   345
         Left            =   3720
         TabIndex        =   8
         Top             =   2730
         Width           =   1125
      End
      Begin VB.ListBox lstSubjects 
         Height          =   2205
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label lblNSubjects 
         Caption         =   "0 subjects"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2820
         Width           =   2295
      End
   End
   Begin VB.Frame fraStudy 
      Caption         =   "Study"
      Height          =   675
      Left            =   60
      TabIndex        =   4
      Top             =   180
      Width           =   3210
      Begin VB.ComboBox cboStudy 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame fraSite 
      Caption         =   "Site"
      Height          =   675
      Left            =   60
      TabIndex        =   2
      Top             =   915
      Width           =   3195
      Begin VB.ComboBox cboSite 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   7170
      TabIndex        =   0
      Top             =   6840
      Width           =   1125
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6180
      Top             =   5100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrSystemIdleTimeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5700
      Top             =   5160
   End
   Begin MSComctlLib.StatusBar sbrMenu 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7230
      Width           =   8385
      _ExtentX        =   14790
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSeparator2 
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
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAboutMacro 
         Caption         =   "&About MACRO"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         frmMenuBatchValidation.frm
' Copyright:    InferMed Ltd. 2003-2007. All Rights Reserved
' Author:       Nicky Johns, February 2003
' Purpose:      Contains the main form of the MACRO 3.0 Batch Validation Module
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 21-28 Feb 03 - Initial development
'   NCJ 6 Mar 03 - Allow spec. of SubjectID rather than label; general UI tidying up
'               Make sure we only deal with "writable" sites
'   NCJ 11 Mar 03 - Ensure all controls disabled during revalidation
'   NCJ 25 Mar 03 - Made "subjects" label bigger
'   NCJ 18 Jun 03 - Bug 1856 - Get all studies/sites (which doesn't check the user's Open Subject permission)
'   NCJ 15 Sept 03 - Bug 2014 - The Subject Label maybe NULL in Oracle so must deal with this
'   NCJ 11 Nov 04 - Issue 2443 - Added "Changes only" check box for log file output
'   NCJ 6 Mar 07 - Issues 2102, 2871 - Fix "All Studies" and allow single eForm
'   NCJ 31 Aug 07 - Issue 2931 - Added "report only" mode where subject data is NOT saved
'----------------------------------------------------------------------------------------'

Option Explicit

Private moArezzo As Arezzo_DM
Private moRevalidator As Revalidator

Private mlSelStudyId As Long
Private msSelStudyName As String
Private msSelSite As String
' NCJ 28 Feb 07 - Selected eForm & ID
Private msSelEForm As String
Private mlSelEFormId As Long

' The subject list array
Private mvSubjects As Variant

Private mlTotalSubjects As Long

Private mbCancelled As Boolean
Private mbCancelMessageDisplayed As Boolean
Private mbRevalidating As Boolean
'Private mbAllStudies As Boolean     ' If they've selected the All Studies option button

Private mbLoading As Boolean

Private mlMinScaleHeight As Long
Private mlMinScaleWidth As Long

Private mcolWritableSites As Collection

Private moTimezone As Timezone

Private Const msALL_SITES = "All Sites"
' NCJ 28 Feb 07 - All Studies
Private Const msALL_STUDIES = "All Studies"
Private Const mnALL_STUDIES = -1

' The minimum top coord for the Revalidation frame
Private Const mlREVALIDATION_TOP As Long = 3000
' The gap between controls
Private Const mlGAP As Long = 60

'--------------------------------------------------------------------
Private Sub cboEForm_Click()
'--------------------------------------------------------------------
' They've selected an eForm
'--------------------------------------------------------------------
    
    If cboEForm.ListIndex > -1 Then
        ' Has it changed?
        If cboEForm.List(cboEForm.ListIndex) <> msSelEForm Then
            msSelEForm = cboEForm.List(cboEForm.ListIndex)
            mlSelEFormId = cboEForm.ItemData(cboEForm.ListIndex)
        End If
    Else
        ' Nothing selected
        If msSelEForm > "" Then
            msSelEForm = ""
            mlSelEFormId = 0
        End If
    End If

End Sub

'--------------------------------------------------------------------
Private Sub chkSingleForm_Click()
'--------------------------------------------------------------------
' Whether they want to do a single eForm
'--------------------------------------------------------------------

    ' Synchronise the eForms drop-down
    Call LoadEForms
    
End Sub

'--------------------------------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------
' They want to cancel the revalidation
'--------------------------------------------------------------------

    mbCancelled = True

End Sub

'--------------------------------------------------------------------
Private Sub cmdExit_Click()
'--------------------------------------------------------------------

    Call mnuFExit_Click

End Sub

'--------------------------------------------------------------------
Public Sub InitialiseMe()
'--------------------------------------------------------------------
' This gets called from Main when MACRO starts up
'--------------------------------------------------------------------
Dim oArezzoMemory As clsAREZZOMemory

    On Error GoTo ErrHandler
    
    mbLoading = True
    
    'The following Doevents prevents command buttons ghosting during form load
    DoEvents
        
    'Create and initialise a new Arezzo instance
    Set moArezzo = New Arezzo_DM
    
    ' NCJ 29 Jan 03 - Get Prolog switches from new ArezzoMemory class
    Set oArezzoMemory = New clsAREZZOMemory
    Call oArezzoMemory.Load(0, goUser.CurrentDBConString)
    Call moArezzo.Init(gsTEMP_PATH, oArezzoMemory.AREZZOSwitches)
    Set oArezzoMemory = Nothing
    
    ' Our "Revalidator" class
    Set moRevalidator = New Revalidator
    
    Set moTimezone = New Timezone
    
    Call ShowNoOfSubjects
    
    Call LoadStudies    ' Automatically triggers LoadSites and LoadEForms
'    If LoadStudies Then
'        Call LoadSites
'    End If
    txtSubject.Text = ""
    
    mbRevalidating = False
   
    mbLoading = False

Exit Sub
ErrHandler:
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
Private Function GetLogFileName() As String
'--------------------------------------------------------------------
' Get name of log file for results of revalidation
'--------------------------------------------------------------------
    
    GetLogFileName = App.Path & "\RVLog " & Format(Now, "d-mmm-yy hh-mm-ss") & ".txt"

End Function

'--------------------------------------------------------------------
Private Sub cmdRefresh_Click()
'--------------------------------------------------------------------
' Refresh the subject list and store the array in mvSubjects
' NCJ 5 Mar 07 - GetSelectedSubjects now does its own display
'--------------------------------------------------------------------
'Dim lStudyId As Long
'Dim sSite As String
'Dim lSubjectRow As Long
'Dim sSubjectRow As String

    On Error GoTo ErrHandler
    
    If mbLoading Then Exit Sub
    
    lstSubjects.Clear
    
    ' All Studies selected?
'    If mbAllStudies Then
    If msSelStudyName = msALL_STUDIES Then
        Call GetAllSubjects
    Else
        Call GetSelectedSubjects
    End If
    ' NCJ 5 Mar 07 - Moved list display code to GetSelectedSubjects
'    If GetSelectedSubjects Then
        ' The mvSubjects array will now have been filled in
'        For lSubjectRow = 0 To UBound(mvSubjects, 2)
'            ' Create Study/Site/Subject listing
'            ' Only take subjects from writable sites
'            If CollectionMember(mcolWritableSites, LCase(mvSubjects(eSubjectListCols.Site, lSubjectRow)), False) Then
'                sSubjectRow = mvSubjects(eSubjectListCols.StudyName, lSubjectRow) _
'                        & "/" & mvSubjects(eSubjectListCols.Site, lSubjectRow)
'                ' If subject label is available, use it, otherwise use subject ID
'                If mvSubjects(eSubjectListCols.SubjectLabel, lSubjectRow) > "" Then
'                    sSubjectRow = sSubjectRow & "/" & mvSubjects(eSubjectListCols.SubjectLabel, lSubjectRow)
'                Else
'                    sSubjectRow = sSubjectRow & "/" & mvSubjects(eSubjectListCols.SubjectId, lSubjectRow)
'                End If
'                lstSubjects.AddItem sSubjectRow
'            End If
'        Next
'        lblNSubjects.Caption = NSubjectsString(lstSubjects.ListCount)
'    End If
    
    ' Anything selected?
    If lstSubjects.ListCount > 0 Then
        cmdRevalidate.Enabled = True
    ElseIf Not mbLoading Then
        DialogWarning "There are no subjects available for revalidation matching your selection"
        cmdRevalidate.Enabled = False
    End If
    
    ListRefreshed = True
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdRefresh_Click", Err.Source) = Retry Then
        Resume
    End If

End Sub

'--------------------------------------------------------------------
Private Function GetSelectedSubjects() As Boolean
'--------------------------------------------------------------------
' Get the array of subjects into mvSubjects
' according to currently selected Study, Site and Subject,
' and update the subject list display.
' Returns FALSE if no subjects found
' NB If "All Sites", mvSubjects is NOT filtered on writable sites,
' but subject list display is filtered (using mcolWritableSites)
'--------------------------------------------------------------------
Dim sSite As String
Dim lSubjectId As Long
Dim sSubjectLabel As String
Dim lSubjectRow As Long
Dim sSubjectRow As String

    On Error GoTo ErrLabel
    
    ' The subject label
    sSubjectLabel = Trim(txtSubject.Text)
    
    If msSelStudyName > "" And msSelSite > "" Then
        If msSelSite = msALL_SITES Then
            sSite = ""
        Else
            sSite = msSelSite
        End If
        mvSubjects = goUser.DataLists.GetSubjectList(sSubjectLabel, msSelStudyName, sSite)
        If IsNull(mvSubjects) Then
            ' If the subject label is numeric, try again using it as a subject ID
            If IsNumeric(sSubjectLabel) Then
                lSubjectId = CLng(sSubjectLabel)
                ' Check it's a sensible subject id
                If lSubjectId > 0 And lSubjectId = CDbl(sSubjectLabel) Then
                    mvSubjects = goUser.DataLists.GetSubjectList(, msSelStudyName, sSite, lSubjectId)
                End If
            End If
        End If
        ' Did we eventually find any subjects?
        If Not IsNull(mvSubjects) Then
            ' Fill in the subject list display
            For lSubjectRow = 0 To UBound(mvSubjects, 2)
                ' Create Study/Site/Subject listing
                ' Only take subjects from writable sites
                If CollectionMember(mcolWritableSites, LCase(mvSubjects(eSubjectListCols.Site, lSubjectRow)), False) Then
                    sSubjectRow = mvSubjects(eSubjectListCols.StudyName, lSubjectRow) _
                            & "/" & mvSubjects(eSubjectListCols.Site, lSubjectRow)
                    ' If subject label is available, use it, otherwise use subject ID
                    If mvSubjects(eSubjectListCols.SubjectLabel, lSubjectRow) > "" Then
                        sSubjectRow = sSubjectRow & "/" & mvSubjects(eSubjectListCols.SubjectLabel, lSubjectRow)
                    Else
                        sSubjectRow = sSubjectRow & "/" & mvSubjects(eSubjectListCols.SubjectId, lSubjectRow)
                    End If
                    lstSubjects.AddItem sSubjectRow
                End If
            Next
        End If
        lblNSubjects.Caption = NSubjectsString(lstSubjects.ListCount)
    End If
        
    GetSelectedSubjects = Not IsNull(mvSubjects)

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmMenu.GetSelectedSubjects"

End Function

'--------------------------------------------------------------------
Private Sub ShowNoOfSubjects()
'--------------------------------------------------------------------
' Set the "no. of subjects" label
'--------------------------------------------------------------------

    lblNSubjects.Caption = NSubjectsString(lstSubjects.ListCount)

End Sub

'--------------------------------------------------------------------
Private Sub RevalidateSubjects()
'--------------------------------------------------------------------
' Revalidate subjects according to user's selection
' NCJ 5 Mar 07 - Issues 2102, 2871 - Now correctly handles All Studies; allow single eForm
' NCJ 31 Aug 07 - Issue 2931 - Added "save changes" mode
'--------------------------------------------------------------------
Dim sLogFile As String
Dim lSubjectsDone As Long
Dim sConfMsg As String
Dim lSubjects As Long
Dim sProgress As String

    On Error GoTo ErrLabel
    
'    If mbAllStudies Then
    If msSelStudyName = msALL_STUDIES Then
        lSubjects = mlTotalSubjects
        sConfMsg = "Revalidate all studies"
        If msSelSite = msALL_SITES Then
            sConfMsg = sConfMsg & "/sites"
        Else
            sConfMsg = sConfMsg & " at site " & msSelSite
        End If
        sConfMsg = sConfMsg & " - " & NSubjectsString(lSubjects)
    Else
        lSubjects = lstSubjects.ListCount
        sConfMsg = "Revalidate "
        If chkSingleForm.Value = 1 Then
            sConfMsg = "Revalidate single eForm" & vbCrLf & msSelEForm & vbCrLf & "for "
        End If
        sConfMsg = sConfMsg & NSubjectsString(lSubjects)
    End If
    
    sConfMsg = sConfMsg & "," & vbCrLf
    ' NCJ 31 Aug 07 - Issue 2931 - Consider "Save changes" mode
    If chkSaveChanges = vbUnchecked Then
        sConfMsg = sConfMsg & "without saving any subject data?"
    Else
        sConfMsg = sConfMsg & "and save all the subject data changes?"
    End If
    
    ' Check they want to do it
    If DialogQuestion(sConfMsg) = vbYes Then
    
        sLogFile = SelectLogFile
        
        If sLogFile > "" Then
            
            DisplayMsg vbCrLf & "Revalidation started " & WhatsTheTime
            ' NCJ 31 Aug 07 - Issue 2931 - Consider "Save changes" mode
            If chkSaveChanges.Value = vbUnchecked Then
                DisplayMsg "Reporting only - no subject data is being saved"
            Else
                DisplayMsg "All changed subject data is being automatically saved"
            End If
            DisplayMsg "Revalidation log file: " & sLogFile
            
            HourglassOn
            
            ' Remember we're busy
            mbRevalidating = True
            
            mbCancelled = False
            mbCancelMessageDisplayed = False
            ' Allow them to cancel
            cmdCancel.Enabled = True
            ' But stop them doing anything else!
            Call EnableButtons(False)
            
            pbBar.Min = 0
            pbBar.Max = lSubjects
            Call ShowProgress(0, lSubjects)
            lSubjectsDone = 0
            
            ' NCJ 11 Nov 04 - Set Verbose mode if they don't want "Changes only"
            moRevalidator.Verbose = (chkChangesOnly = vbUnchecked)
            ' NCJ 31 Aug 07 - Set SaveChanges mode
            moRevalidator.SaveChanges = (chkSaveChanges = vbChecked)

            ' NCJ 6 Mar 07 - Added msSelEForm (selected eForm - may be "")
            Call moRevalidator.InitRevalidation(sLogFile, goUser, msSelEForm)
            
            Me.Refresh
            DoEvents
                                 
'            If mbAllStudies Then
            If msSelStudyName = msALL_STUDIES Then
                Call RevalidateAllSubjects(lSubjects)
            Else
                ' mvSubjects already set up
                Call RevalidateSetOfSubjects(lSubjects, lSubjectsDone)
            End If
            
            Call moRevalidator.EndRevalidation
            
            cmdCancel.Enabled = False
            Call EnableButtons(True)
            If Not mbCancelMessageDisplayed Then
                DisplayMsg "Revalidation complete " & WhatsTheTime
            End If
            HourglassOff
        
            ' We're not busy any more
            mbRevalidating = False
        End If
        
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmMenu.RevalidateSubjects"

End Sub

'--------------------------------------------------------------------
Private Sub RevalidateSetOfSubjects(ByVal lTotalSubjs As Long, ByRef lSubjsDoneSoFar As Long)
'--------------------------------------------------------------------
' Revalidate the subjects contained in the data array mvSubjects (assumed non-empty)
' Assume we have lTotalSubjs to do, and we've done lSubjsDoneSoFar before coming into this routine
' Update lSubjsDoneSoFar with how many we did here
' NCJ 28 Feb 07 - Issue 2871 - Added selected EForm to Revalidate
'--------------------------------------------------------------------
Dim lSubjectRow As Long
Dim sSubjLabel As String

    On Error GoTo ErrLabel
    
    ' NCJ 28 Feb 07 - Do the check here for an empty list
    If IsNull(mvSubjects) Then Exit Sub

    ' Loop through revalidating all the subjects
    For lSubjectRow = 0 To UBound(mvSubjects, 2)
        
        ' See if they cancelled
        If UserCancelled Then Exit For
        
        ' Only take subjects from writable sites (previously set up)
        If CollectionMember(mcolWritableSites, LCase(mvSubjects(eSubjectListCols.Site, lSubjectRow)), False) Then
            ' Revalidate the next subject
            ' NCJ 15 Sept 03 - Bug 2014 - The SubjLabel maybe NULL in Oracle so must deal with this
            If mvSubjects(eSubjectListCols.SubjectLabel, lSubjectRow) <> "" Then
                sSubjLabel = mvSubjects(eSubjectListCols.SubjectLabel, lSubjectRow)
            Else
                sSubjLabel = ""
            End If
            ' NCJ 28 Feb 07 - Issue 2871 - Added selected EForm, mlSelEFormId
            Call moRevalidator.Revalidate(goUser, _
                                mvSubjects(eSubjectListCols.StudyId, lSubjectRow), _
                                mvSubjects(eSubjectListCols.Site, lSubjectRow), _
                                mvSubjects(eSubjectListCols.SubjectId, lSubjectRow), _
                                mvSubjects(eSubjectListCols.StudyName, lSubjectRow), _
                                sSubjLabel, _
                                moArezzo, mlSelEFormId)
            lSubjsDoneSoFar = lSubjsDoneSoFar + 1
            Call ShowProgress(lSubjsDoneSoFar, lTotalSubjs)
            
            Me.Refresh
            DoEvents
        End If
    
    Next

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmMenu.RevalidateSetOfSubjects"

End Sub

'--------------------------------------------------------------------
Private Function UserCancelled() As Boolean
'--------------------------------------------------------------------
' Has the user clicked the Cancel button?
' If so, display "Cancelled" message and say how far we'd got
'--------------------------------------------------------------------
        
    If mbCancelled And Not mbCancelMessageDisplayed Then
        DisplayMsg "Revalidation cancelled. (" & lblProgress.Caption & ") " & WhatsTheTime
        ' Remember we've already displayed the message
        mbCancelMessageDisplayed = True
    End If
    UserCancelled = mbCancelled
    
End Function

'--------------------------------------------------------------------
Private Sub RevalidateAllSubjects(ByVal lTotalSubjects As Long)
'--------------------------------------------------------------------
' Revalidate all the revalidatable subjects in the database
' NCJ 28 Feb 07 - Filter on Site if required
'--------------------------------------------------------------------
Dim colStudies As Collection
Dim oStudy As Study
Dim colSites As Collection
Dim oSite As Site
Dim lSubjsDoneSoFar As Long

    On Error GoTo ErrLabel
    
    lSubjsDoneSoFar = 0
    
    ' Get all the studies for this user
    Set colStudies = goUser.GetOpenSubjectStudies
    
    For Each oStudy In colStudies
        
        ' See if they cancelled the revalidation
        If UserCancelled Then Exit For
        
        If msSelSite = msALL_SITES Then
            ' We're doing ALL sites
            Set colSites = goUser.GetOpenSubjectSites(oStudy.StudyId)
            
            For Each oSite In colSites
                ' See if they cancelled
                If UserCancelled Then Exit For
            
                ' Can't open subjects from Remote sites on the Server
                If Not (goUser.DBIsServer And oSite.SiteLocation = TypeOfInstallation.RemoteSite) Then
                    mvSubjects = goUser.DataLists.GetSubjectList(, oStudy.StudyName, oSite.Site)
                    Call RevalidateSetOfSubjects(lTotalSubjects, lSubjsDoneSoFar)
                End If
            Next
        ElseIf msSelSite > "" Then
            ' They selected a specific site (assume writable!)
            mvSubjects = goUser.DataLists.GetSubjectList(, oStudy.StudyName, msSelSite)
            Call RevalidateSetOfSubjects(lTotalSubjects, lSubjsDoneSoFar)
        End If
    
    Next
    
    Set colStudies = Nothing
    Set colSites = Nothing
    Set oStudy = Nothing
    Set oSite = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmMenu.RevalidateAllSubjects"

End Sub

'--------------------------------------------------------------------
Private Sub cmdRevalidate_Click()
'--------------------------------------------------------------------
' Revalidate subjects
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call RevalidateSubjects

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdRevalidate_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub ShowProgress(ByVal nSubjsDone As Integer, ByVal nTotalSubjs As Integer)
'--------------------------------------------------------------------
' Display revalidation progress
'--------------------------------------------------------------------

    pbBar.Value = nSubjsDone
    lblProgress.Caption = "Completed " & nSubjsDone & " of " & nTotalSubjs

End Sub

'--------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------
' NCJ 27 Feb 07 - Removed All/Selected option buttons; added chkSingleEForm, cboEForm
' NCJ 31 Aug 07 - Added chkSaveChanges
'--------------------------------------------------------------------
    
    mbLoading = True
    
    FormCentre Me
    
    txtMsg.Text = ""
    lblProgress.Caption = ""
    
    cmdCancel.Enabled = False
    cmdRevalidate.Enabled = False
    cmdRefresh.Enabled = False
    
    pbBar.Min = 0
    pbBar.Value = 0
    
'    optSelected.Value = True
'    optAllStudies.Value = False
'    optSelected.Visible = False
'    optAllStudies.Visible = False
'    mbAllStudies = False
    
    ' Remember our "minimum" size (this is how we start off)
    mlMinScaleHeight = Me.ScaleHeight
    mlMinScaleWidth = Me.ScaleWidth
    
    ' Default to verbose (i.e. not "Changes only")
    chkChangesOnly.Value = vbUnchecked
    ' Default to all eForms (i.e. not single eForm)
    chkSingleForm.Value = vbUnchecked
    ' Initially disable the eForms combo
    cboEForm.Enabled = False
    ' Default to Save Changes
    chkSaveChanges.Value = vbChecked
    
    ' No selections yet
    msSelStudyName = ""
    mlSelStudyId = 0
    msSelSite = ""
    ' NCJ 27 Feb 07
    msSelEForm = ""
    mlSelEFormId = 0
    
    mbLoading = False

End Sub

'--------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If UnloadMode = vbFormControlMenu Then
        ' Don't let them out if we're busy revalidating
        If mbRevalidating Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    Call TidyUpOnExit
    
    Call ExitMACRO
    Call MACROEnd

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_QueryUnload", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub Form_Resize()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'check that form has not been minimised
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.ScaleWidth <= mlMinScaleWidth Then
        ' Fit to min. width if width less than minimum
        Call FitToWidth(mlMinScaleWidth)
    Else
        Call FitToWidth(Me.ScaleWidth)
    End If

    If Me.ScaleHeight <= mlMinScaleHeight Then
        ' Set to the "minimum" height
        Call FitToHeight(mlMinScaleHeight)
    Else
        Call FitToHeight(Me.ScaleHeight)
    End If

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Resize", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub FitToHeight(ByVal lWinHeight As Long)
'--------------------------------------------------------------------
' Fit the controls into the given window height
' Assume the height is not below the minimum
'--------------------------------------------------------------------

    ' Move the Revalidation area down (we don't change its height)
    cmdExit.Top = lWinHeight - sbrMenu.Height - mlGAP - cmdExit.Height
    fraRevalidation.Top = cmdExit.Top - mlGAP - fraRevalidation.Height
    
    ' Pull the subject list down to meet it
    fraSubjects.Height = fraRevalidation.Top - mlGAP - fraSubjects.Top
    ' Size the lstSubjects first because it won't be exact
    lstSubjects.Height = fraSubjects.Height - lstSubjects.Top - cmdRefresh.Height - 2 * mlGAP
    cmdRefresh.Top = lstSubjects.Top + lstSubjects.Height + mlGAP
    lblNSubjects.Top = cmdRefresh.Top
    
    ' Make sure status bar always sits on top
    sbrMenu.ZOrder
    
End Sub

'--------------------------------------------------------------------
Private Sub FitToWidth(ByVal lWinWidth As Long)
'--------------------------------------------------------------------
' Fit the controls into the given window width
' Assume the width is not below the minimum
'--------------------------------------------------------------------

    ' Expand the subject list area
    fraSubjects.Width = lWinWidth - fraSubjects.Left - mlGAP
    lstSubjects.Width = fraSubjects.Width - 2 * lstSubjects.Left
    cmdRefresh.Left = lstSubjects.Left + lstSubjects.Width - cmdRefresh.Width
    
    ' Expand the revalidation area
    fraRevalidation.Width = lWinWidth - fraRevalidation.Left - mlGAP
    txtMsg.Width = fraRevalidation.Width - 2 * txtMsg.Left
    cmdCancel.Left = txtMsg.Left + txtMsg.Width - cmdCancel.Width
    cmdRevalidate.Left = cmdCancel.Left - cmdRevalidate.Width - mlGAP
    pbBar.Width = txtMsg.Width
    
    ' Finally move the Exit button over
    cmdExit.Left = fraRevalidation.Left + fraRevalidation.Width - cmdExit.Width

End Sub

'--------------------------------------------------------------------
Private Sub mnuFExit_Click()
'--------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    Call TidyUpOnExit
    
    Call ExitMACRO
    Call MACROEnd

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdExit_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub mnuHAboutMacro_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    frmAbout.Display

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuHAboutMacro_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub mnuHUserGuide_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call MACROHelp(Me.hWnd, App.Title)

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuHUserGuide_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Public Function ForgottenPassword(sSecurityCon As String, sUsername As String, sPassword As String, sErrMsg As String) As Boolean
'---------------------------------------------------------------------
'dummy function for frmNewLogin to compile
'---------------------------------------------------------------------


End Function

'---------------------------------------------------------------------
Private Sub TidyUpOnExit()
'---------------------------------------------------------------------
' Tidy up when exiting
'---------------------------------------------------------------------
    
    Set moRevalidator = Nothing
    
    ' Only shut down the ALM if it has been started
    If Not moArezzo Is Nothing Then
        moArezzo.Finish
        Set moArezzo = Nothing
    End If

End Sub

'--------------------------------------------
Private Sub DisplayMsg(sText As String)
'--------------------------------------------
' Display a message in the Message Window, followed by CR
'--------------------------------------------
    
    txtMsg.Text = txtMsg.Text & vbCrLf & sText

End Sub

'--------------------------------------------
Public Function LoadStudies() As Boolean
'--------------------------------------------
' Populate the Study combo with studies the user has access to
' NCJ 27 Feb 07 - Issue 2102 - Added All Studies
'--------------------------------------------
Dim lRow As Long
Dim vStudies As Variant
Dim oStudy As Study
Dim colStudies As Collection

    HourglassOn
    
    cboStudy.Clear
    
    ' NCJ 18 Jun 03 - Bug 1856 - Get all studies (which doesn't check the user's Open Subject permission)
    Set colStudies = goUser.GetAllStudies
    
    ' Are there any studies?
    If colStudies.Count = 0 Then
        LoadStudies = False
    Else
        ' Add the studies to the combo
        ' and the study IDs to the ItemData array
        cboStudy.AddItem msALL_STUDIES
        cboStudy.ItemData(cboStudy.NewIndex) = mnALL_STUDIES
        For Each oStudy In colStudies
            cboStudy.AddItem oStudy.StudyName
            cboStudy.ItemData(cboStudy.NewIndex) = oStudy.StudyId
        Next
        cboStudy.ListIndex = 1  ' Bypass "All Studies" to begin with
        LoadStudies = True
    End If
 
    HourglassOff

End Function

'--------------------------------------------
Public Sub LoadSites()
'--------------------------------------------
' Populate the Sites combo with sites the user has access to
' according to the chosen study
'--------------------------------------------
Dim lRow As Long
Dim vSites As Variant

Dim colSites As Collection
Dim oSite As Site

    cboSite.Clear
    msSelSite = ""
    HourglassOn
    
    ' NCJ 18 Jun 03 - Bug 1856 - Get all sites (which doesn't check the user's Open Subject permission)
'    Set colSites = goUser.GetOpenSubjectSites(cboStudy.ItemData(cboStudy.ListIndex))
    Set colSites = goUser.GetAllSites(cboStudy.ItemData(cboStudy.ListIndex))
    Set mcolWritableSites = New Collection

    ' Are there any sites?
    If colSites.Count > 0 Then
        cboSite.AddItem msALL_SITES
        For Each oSite In colSites
            ' Can't open subjects from Remote sites on the Server
            If Not (goUser.DBIsServer And oSite.SiteLocation = TypeOfInstallation.RemoteSite) Then
                cboSite.AddItem oSite.Site
                mcolWritableSites.Add LCase(oSite.Site), LCase(oSite.Site)
            End If
        Next
    End If
    
    ' Any available sites for the study?
    If mcolWritableSites.Count > 0 Then
        ' Select "All sites"
        cboSite.ListIndex = 0
    Else
        ' No sites for this study
        cboSite.Clear
        Call ClearSubjects
        If Not mbLoading Then
            Call DialogWarning("There are no subjects available for revalidation in this study")
        End If
    End If
    
    
    HourglassOff
    
End Sub

'--------------------------------------------------------------------
Private Sub LoadEForms()
'--------------------------------------------------------------------
' NCJ 28 Feb 07 - Issue 2871
' Load the eForms combo according to the currently selected study
' Also enable subject spec accordingly
'--------------------------------------------------------------------
Dim colEForms As Collection
Dim oEForm As eFormRO

    On Error GoTo ErrLabel
    
    ' Clear combo and reset selected eForm
    cboEForm.Clear
    msSelEForm = ""
    mlSelEFormId = 0
    cboEForm.Enabled = False
    chkSingleForm.Enabled = False
    txtSubject.Enabled = False
    
    If mlSelStudyId < 1 Then
        ' Either no study or All Studies, so can't select Single EForm
        chkSingleForm.Value = 0
    Else
        ' Single study selected
        txtSubject.Enabled = True
        ' Make sure eForm check box is enabled
        chkSingleForm.Enabled = True
        If chkSingleForm.Value = 1 Then
            ' They want to select an eForm
            Set colEForms = moRevalidator.StudyEForms(mlSelStudyId, goUser, moArezzo)
            If Not colEForms Is Nothing Then
                ' Assume a study has some eForms!
                For Each oEForm In colEForms
                    cboEForm.AddItem oEForm.code
                    cboEForm.ItemData(cboEForm.NewIndex) = oEForm.EFormId
                Next
                ' Enable drop-down and select first form
                cboEForm.Enabled = True
                cboEForm.ListIndex = 0
                Set colEForms = Nothing
                Set oEForm = Nothing
            End If
        End If
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmMenu.LoadEForms"

End Sub

'--------------------------------------------------------------------
Private Sub cboSite_Click()
'--------------------------------------------------------------------
' They clicked on a Site
'--------------------------------------------
    
    If cboSite.ListIndex > -1 Then
        If cboSite.List(cboSite.ListIndex) <> msSelSite Then
            msSelSite = cboSite.List(cboSite.ListIndex)
            ListRefreshed = False
        End If
    Else
        If msSelSite > "" Then
            msSelSite = ""
            ListRefreshed = False
        End If
    End If
    
End Sub

'--------------------------------------------
Private Sub cboStudy_Click()
'--------------------------------------------
' They clicked on a Study
'--------------------------------------------

    ' Any study chosen?
    If cboStudy.ListIndex > -1 Then
        If cboStudy.List(cboStudy.ListIndex) <> msSelStudyName Then
            msSelStudyName = cboStudy.List(cboStudy.ListIndex)
            mlSelStudyId = cboStudy.ItemData(cboStudy.ListIndex)
            Call LoadSites
            ' NCJ 28 Feb 07 - Check whether we need to update the eForms list
            Call LoadEForms
            ListRefreshed = False
        End If
    Else
        If msSelStudyName > "" Then
            msSelStudyName = ""
            mlSelStudyId = 0
            ' NCJ 28 Feb 07 - Check whether we need to update the eForms list
            Call LoadEForms
            ListRefreshed = False
        End If
    End If

End Sub

'--------------------------------------------------------------------
Private Function SelectLogFile() As String
'--------------------------------------------------------------------
' Get the user to select a file to contain the revalidation results
'--------------------------------------------------------------------
Dim sLogFile As String

    On Error GoTo CancelOpen
    
    With CommonDialog1
        .DialogTitle = "MACRO Validation Log File"
        .InitDir = gsTEMP_PATH
        .DefaultExt = "txt"
        .Filter = "Text file (*.txt)|*.txt|Log file (*.log)|*.log|All files (*.*)|*.*"
        .FilterIndex = 1
        .CancelError = True
        .Flags = cdlOFNCreatePrompt + cdlOFNPathMustExist _
                    + cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
        .ShowSave
  
        sLogFile = .FileName
    End With

    SelectLogFile = sLogFile
    
CancelOpen:

End Function

'--------------------------------------------------------------------
Private Sub EnableButtons(ByVal bEnable As Boolean)
'--------------------------------------------------------------------
' Enable/disable buttons before or after revalidation
' NCJ 28 Feb 07 - Take into account current settings too
'--------------------------------------------------------------------

    cmdRefresh.Enabled = bEnable
    cmdExit.Enabled = bEnable
    
    mnuFile.Enabled = bEnable
    mnuFHelp.Enabled = bEnable
    
    cboStudy.Enabled = bEnable
    cboSite.Enabled = bEnable
    ' NCJ 6 Sep 07 - "Save changes" check box
    chkSaveChanges.Enabled = bEnable
    
    ' Only enable if there are some subjects to revalidate
    cmdRevalidate.Enabled = bEnable And lstSubjects.ListCount > 0
    
    ' Only enable Single EForm and Subject if single study selected
    chkSingleForm.Enabled = bEnable And mlSelStudyId > 0
    txtSubject.Enabled = bEnable And mlSelStudyId > 0
    
    'Only enable if eForms combo contains something
    cboEForm.Enabled = bEnable And cboEForm.ListCount > 0

End Sub

'--------------------------------------------------------------------
Private Function WhatsTheTime() As String
'--------------------------------------------------------------------
' The current time formatted as a string
'--------------------------------------------------------------------

    WhatsTheTime = Format(Now, "hh:mm:ss")

End Function

''--------------------------------------------------------------------
'Private Sub optAllStudies_Click()
''--------------------------------------------------------------------
'' They want to do ALL studies
''--------------------------------------------------------------------
'
'    mbAllStudies = True
'    ' Disable the study/site drop-downs
'    Call EnableStudySiteSubject(False)
'    ' Refresh the list box
'    Call cmdRefresh_Click
'
'End Sub
'
''--------------------------------------------------------------------
'Private Sub optSelected_Click()
''--------------------------------------------------------------------
'' They want to do selected studies
''--------------------------------------------------------------------
'
'    mbAllStudies = False
'    ' Enable the study/site drop-downs
'    Call EnableStudySiteSubject(True)
'    ' Refresh the list box
'    Call cmdRefresh_Click
'
'End Sub

'--------------------------------------------------------------------
Private Sub ClearSubjects()
'--------------------------------------------------------------------
' Clear the subject list display
'--------------------------------------------------------------------
    
    mlTotalSubjects = 0
    lstSubjects.Clear
    mvSubjects = Null
    lblNSubjects.Caption = NSubjectsString(mlTotalSubjects)
    cmdRevalidate.Enabled = False

End Sub

'--------------------------------------------------------------------
Private Function GetAllSubjects() As Long
'--------------------------------------------------------------------
' Get all potentially revalidatable subjects in the database
' (NCJ 28 Feb 07 - optionally filtered by site)
' but display just the study/site combinations in the listbox
'--------------------------------------------------------------------
Dim colStudies As Collection
Dim oStudy As Study
Dim colSites As Collection
Dim oSite As Site

    HourglassOn
    
    Call ClearSubjects
    
    Set colStudies = goUser.GetOpenSubjectStudies
    For Each oStudy In colStudies
        Set colSites = goUser.GetOpenSubjectSites(oStudy.StudyId)
        For Each oSite In colSites
            ' Check it's either the selected site OR we're doing all sites
            If msSelSite = oSite.Site Or msSelSite = msALL_SITES Then
                ' Also can't open subjects from Remote sites on the Server
                If Not (goUser.DBIsServer And oSite.SiteLocation = TypeOfInstallation.RemoteSite) Then
                    mvSubjects = goUser.DataLists.GetSubjectList(, oStudy.StudyName, oSite.Site)
                    If Not IsNull(mvSubjects) Then
                        mlTotalSubjects = mlTotalSubjects + UBound(mvSubjects, 2) + 1
                        lstSubjects.AddItem oStudy.StudyName & "/" & oSite.Site & " (" & UBound(mvSubjects, 2) + 1 & ")"
                    End If
                End If
            End If
        Next
    Next
    
    lblNSubjects.Caption = NSubjectsString(mlTotalSubjects)
    
    cmdRevalidate.Enabled = (mlTotalSubjects > 0)
    
    Set colStudies = Nothing
    Set colSites = Nothing
    Set oStudy = Nothing
    Set oSite = Nothing
    
    ListRefreshed = True
    
    HourglassOff

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmMenu.GetAllSubjects"

End Function

'--------------------------------------------------------------------
Private Property Get ListRefreshed() As Boolean
'--------------------------------------------------------------------
' Is the list up to date?
'--------------------------------------------------------------------

    ListRefreshed = cmdRefresh.Enabled
    
End Property

'--------------------------------------------------------------------
Private Property Let ListRefreshed(bRefreshed As Boolean)
'--------------------------------------------------------------------
' Whether the list of subjects is refreshed
' NCJ 28 Feb 07 - Disallow revalidation if list not up to date
'--------------------------------------------------------------------

    cmdRefresh.Enabled = Not bRefreshed And (msSelSite > "") And txtSubject.BackColor = vbWindowBackground
    ' Don't let them revalidate unless the list is up to date
    cmdRevalidate.Enabled = Not cmdRefresh.Enabled And lstSubjects.ListCount > 0

End Property

'--------------------------------------------------------------------
Private Sub txtSubject_Change()
'--------------------------------------------------------------------
' Enable list refresh when they change the subject label
'--------------------------------------------------------------------
Dim sSubjLabel As String

    sSubjLabel = Trim(txtSubject.Text)
    If Not gblnValidString(sSubjLabel, valOnlySingleQuotes) Then
        txtSubject.BackColor = vbYellow
        ' Prevent list refreshing
        ListRefreshed = True
    Else
        txtSubject.BackColor = vbWindowBackground
        ListRefreshed = False
    End If
    
End Sub

'--------------------------------------------------------------------
Private Sub txtSubject_KeyPress(KeyAscii As Integer)
'--------------------------------------------------------------------
' Intercept RETURN to mean cmdRefresh_Click
'--------------------------------------------------------------------

    If KeyAscii = vbKeyReturn Then
        If cmdRefresh.Enabled Then
            Call cmdRefresh_Click
        End If
    End If

End Sub

'--------------------------------------------------------------------
Private Function NSubjectsString(ByVal lSubjects As Long) As String
'--------------------------------------------------------------------
' Returns a string saying "N subjects" or "1 subject" if lSubjects = 1
'--------------------------------------------------------------------
Dim sSubjects As String

    sSubjects = lSubjects & " subject"
    If lSubjects = 1 Then
        ' Leave it as it is
    Else
        ' Add an "s"
        sSubjects = sSubjects & "s"
    End If

    NSubjectsString = sSubjects

End Function

