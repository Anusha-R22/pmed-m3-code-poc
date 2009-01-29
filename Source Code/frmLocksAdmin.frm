VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLocksAdmin 
   Caption         =   "Database Lock Administration"
   ClientHeight    =   4110
   ClientLeft      =   7695
   ClientTop       =   5130
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   8430
   Begin VB.Frame fraLocks 
      Caption         =   "Current Locks"
      Height          =   4035
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   6975
      Begin MSComctlLib.ListView lvwLocks 
         Height          =   3675
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6482
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   375
      Left            =   7140
      TabIndex        =   3
      Top             =   3660
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Default         =   -1  'True
      Height          =   375
      Left            =   7140
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   7140
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmLocksAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
'   File:       frmLocksAdmin.frm
'   Copyright:  InferMed Ltd. 2001-2006. All Rights Reserved
'   Author:     Matthew Martin, September 2001
'   Purpose:    User Interface for removal of expired locks.
'----------------------------------------------------------------------------------------
' Revisions:
' ZA 02/10/01 added error handling routines
' TA 04/10/01 Standardised GUI
' TA 05/10/01 Allowed multiple selction of locks
' MLM 28/05/02 Current Build Buglist 2.2.12 no. 26: handle "all studies" locks
' ASH 12/06/2002 Applied macro date format bug 2.2.14 no.24
'REM 18/10/02 Added DatabaseCode to the dislplay function so that the refresh function could
' creat ethe appropriate connection string to the database
' NCJ 25 Oct 06 - Check for an eForm lock in MUSD with SubjectID = 0
'----------------------------------------------------------------------------------------

Option Explicit

Dim moUser As MACROUser 'to remember the permissions of the user
Dim mvLocks As Variant 'to remember the various IDs of the locks while they're displayed to the user

Private msDatabaseCode As String

'----------------------------------------------------------------------------------------
Public Sub Display(oUser As MACROUser, sDatabaseCode As String)
'----------------------------------------------------------------------------------------
' Only method this form should be displayed.
'----------------------------------------------------------------------------------------

    Load Me
    Set moUser = oUser
    msDatabaseCode = sDatabaseCode
    FormCentre Me
    RefreshList
    Me.Show vbModal
    Set moUser = Nothing
    
End Sub

'----------------------------------------------------------------------------------------
Private Sub RefreshList()
'----------------------------------------------------------------------------------------
' MLM 12/09/01: Refresh the list of displayed locks from the database.
' MLM 28/05/02: handle new "all studies" lock type
' REM 18/10/02: added MACROUserBS30 reference so can get the connection string for the passed in DatabaseCode
' ASH 28/1/2003: Added code to display subject label if exists or still display subjectid
' NCJ 25 Oct 06 - Check for MUSD eForm locks where SubjectID = 0
'----------------------------------------------------------------------------------------

Dim oDBLock As DBLock
Dim lCount As Long
Dim oDatabase As MACROUserBS30.Database
Dim bLoad As Boolean
Dim sMessage As String
'ASH 28/1/2003
Dim sSubjectLabel As String
Dim lTrialId As Long
Dim lSubjectId As Long
Dim sSite As String
Dim vLocks As Variant

    On Error GoTo ErrHandler
    
    sSubjectLabel = ""
    
    Set oDBLock = New DBLock
    
    Set oDatabase = New MACROUserBS30.Database
    bLoad = oDatabase.Load(SecurityADODBConnection, moUser.UserName, msDatabaseCode, "", False, sMessage)
    
    If moUser.CheckPermission(gsFnRemoveAllLocks) Then
        mvLocks = oDBLock.AllLockDetails(oDatabase.ConnectionString)
    Else
        If moUser.CheckPermission(gsFnRemoveOwnLocks) Then
            mvLocks = oDBLock.AllLockDetails(oDatabase.ConnectionString, moUser.UserName)
        Else
            'no permissions so don't give them any locks back
            DialogInformation "You do not have permission to remove locks"
            mvLocks = Null
        End If
    End If
    
    vLocks = Null
    If Not IsNull(mvLocks) Then
        'prettify the data prior to display
        For lCount = 0 To UBound(mvLocks, 2)
            'convert double timestamps to formatted date strings
            'mvLocks(LockDetailColumn.ldcLockTimeStamp, lCount) = _
                'Format(CDate(mvLocks(LockDetailColumn.ldcLockTimeStamp, lCount)), "dd\/mm\/yyyy hh\:mm")
                'ASH 12/06/2002 Applied macro date format bug 2.2.14 no.24
                mvLocks(LockDetailColumn.ldcLockTimeStamp, lCount) = _
                Format(CDate(mvLocks(LockDetailColumn.ldcLockTimeStamp, lCount)), "yyyy/mm/dd hh:mm:ss")
               'MLM 28/05/02: handle new "all studies" lock type
            If mvLocks(LockDetailColumn.ldcStudyId, lCount) = -1 Then
                mvLocks(LockDetailColumn.ldcStudyName, lCount) = "All studies"
            End If
            ' NCJ 25 Oct 06 - Check for MUSD eForm locks where SubjectID = 0
            If mvLocks(LockDetailColumn.ldcSubjectId, lCount) = 0 And mvLocks(LockDetailColumn.ldcEFormInstanceId, lCount) > 0 Then
                mvLocks(LockDetailColumn.ldcEFormTitle, lCount) = mvLocks(LockDetailColumn.ldcEFormSDTitle, lCount)
            End If
        Next lCount
        
        vLocks = mvLocks
        For lCount = 0 To UBound(mvLocks, 2)
            'ASH 28/1/2003 Use subject label or IDs
            If vLocks(LockDetailColumn.ldcSubjectId, lCount) > 0 Then
                lTrialId = mvLocks(LockDetailColumn.ldcStudyId, lCount)
                lSubjectId = mvLocks(LockDetailColumn.ldcSubjectId, lCount)
                sSite = mvLocks(LockDetailColumn.ldcSite, lCount)
                sSubjectLabel = SubjectLabelFromTrialSiteId(lTrialId, sSite, lSubjectId)
                If sSubjectLabel <> "" Then
                    vLocks(LockDetailColumn.ldcSubjectId, lCount) = sSubjectLabel
                End If
            End If
        Next lCount
        
    End If

    lvw_FromArray lvwLocks, vLocks, Array("Study", "Site", "Subject", "eForm", "eForm Cycle", "User", "TimeStamp")
    
    If IsNull(mvLocks) Then
        'there are no locks
        cmdDelete.Enabled = False
    Else
        'select the 1st listed lock by default
        lvwLocks.ListItems(1).Selected = True
        cmdDelete.Enabled = True
    End If
    
    Exit Sub
    
ErrHandler:
    Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "RefreshList", Err.Source)
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

'----------------------------------------------------------------------------------------
Private Sub cmdRefresh_Click()
'----------------------------------------------------------------------------------------
' MLM 14/09/01: Refresh :P
'----------------------------------------------------------------------------------------
    
    RefreshList
 
   
End Sub

Private Sub Form_Load()
    
    Me.Icon = frmMenu.Icon
    
End Sub

'----------------------------------------------------------------------------------------
Private Sub cmdDelete_Click()
'----------------------------------------------------------------------------------------
' MLM 14/09/01: Delete the selected lock.
'----------------------------------------------------------------------------------------
Dim lRow As Long
Dim lLockIndex As Long
Dim oDBLock As DBLock
Dim lSelectCount As Long
Dim nAnswer As Integer

    On Error GoTo ErrHandler
    
        
    'count select rows
    lSelectCount = 0
    For lRow = 1 To lvwLocks.ListItems.Count
        If lvwLocks.ListItems(lRow).Selected Then
            lSelectCount = lSelectCount + 1
        End If
    Next
    
    If lSelectCount = 1 Then
        nAnswer = DialogQuestion("Are you sure you wish to delete this lock?", "Locks Administration")
    Else
        nAnswer = DialogQuestion("Are you sure you wish to delete these " & lSelectCount & " selected locks?", "Locks Administration")
    End If
    If nAnswer = vbYes Then
    
        'go down listview listitems and delete locks when selected
        For lRow = 1 To lvwLocks.ListItems.Count
            If lvwLocks.ListItems(lRow).Selected Then
                lLockIndex = lRow - 1
                Set oDBLock = New DBLock
                
                If Not IsNull(mvLocks(LockDetailColumn.ldcEFormInstanceId, lLockIndex)) Then
                    'the user is deleting a form lock
                    oDBLock.UnlockEFormInstance moUser.Database.ConnectionString, _
                        mvLocks(LockDetailColumn.ldcToken, lLockIndex), _
                        mvLocks(LockDetailColumn.ldcStudyId, lLockIndex), _
                        mvLocks(LockDetailColumn.ldcSite, lLockIndex), _
                        mvLocks(LockDetailColumn.ldcSubjectId, lLockIndex), _
                        mvLocks(LockDetailColumn.ldcEFormInstanceId, lLockIndex)
                ElseIf Not IsNull(mvLocks(lLockIndex, LockDetailColumn.ldcSubjectId, lLockIndex)) Then
                    'subject
                    oDBLock.UnlockSubject moUser.Database.ConnectionString, _
                        mvLocks(LockDetailColumn.ldcToken, lLockIndex), _
                        mvLocks(LockDetailColumn.ldcStudyId, lLockIndex), _
                        mvLocks(LockDetailColumn.ldcSite, lLockIndex), _
                        mvLocks(LockDetailColumn.ldcSubjectId, lLockIndex)
                Else
                    'study
                    oDBLock.UnlockStudy moUser.Database.ConnectionString, _
                        mvLocks(LockDetailColumn.ldcToken, lLockIndex), _
                        mvLocks(LockDetailColumn.ldcStudyId, lLockIndex)
                End If
            End If
        Next
    
        RefreshList
    End If
    
    Exit Sub
    
ErrHandler:
    Select Case Err.Number
        Case -2147221504
            'this error means that the lock has disappeared of its own accord before you managed to delete it, so that's ok
            Err.Clear
            RefreshList
        Case Else
            Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdDelete_Click", Err.Source)
                Case OnErrorAction.Retry
                    Resume
                Case OnErrorAction.QuitMACRO
                    Call ExitMACRO
                    Call MACROEnd
            End Select
            
        End Select
End Sub

'----------------------------------------------------------------------------------------
Private Sub Form_Resize()
'----------------------------------------------------------------------------------------
'Resixing of form
'----------------------------------------------------------------------------------------

    'just exit if there are any errors
    '(controls will stay put)
    On Error GoTo ErrLabel
    
    fraLocks.Height = Me.ScaleHeight - 60
    lvwLocks.Height = fraLocks.Height - 360
    
    cmdClose.Top = fraLocks.Top + fraLocks.Height - cmdClose.Height
    
    fraLocks.Width = Me.ScaleWidth - cmdDelete.Width - 240
    lvwLocks.Width = fraLocks.Width - 240
    
    cmdDelete.Left = fraLocks.Left + fraLocks.Width + 120
    cmdRefresh.Left = cmdDelete.Left
    cmdClose.Left = cmdDelete.Left
    
    Exit Sub
    
ErrLabel:
    
End Sub

