VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmViewLockFreeze 
   Caption         =   "Lock / Freeze History"
   ClientHeight    =   8085
   ClientLeft      =   3870
   ClientTop       =   3270
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   12510
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   11340
      TabIndex        =   20
      Top             =   7680
      Width           =   1125
   End
   Begin MSComctlLib.ListView lvwTransferStatus 
      Height          =   7515
      Left            =   2520
      TabIndex        =   19
      Top             =   60
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   13256
      LabelEdit       =   1
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
   Begin VB.Frame fraSelection 
      Height          =   7575
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   2415
      Begin VB.Frame fraSite 
         Caption         =   "Site"
         Height          =   615
         Left            =   60
         TabIndex        =   30
         Top             =   1140
         Width           =   1755
         Begin VB.ComboBox cboSite 
            Height          =   315
            Left            =   60
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   180
            Width           =   1575
         End
      End
      Begin VB.Frame fraStudy 
         Caption         =   "Study"
         Height          =   615
         Left            =   60
         TabIndex        =   29
         Top             =   540
         Width           =   1755
         Begin VB.ComboBox cboStudy 
            Height          =   315
            Left            =   60
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   180
            Width           =   1575
         End
      End
      Begin VB.Frame fraSource 
         Caption         =   "Source"
         Height          =   795
         Left            =   60
         TabIndex        =   28
         Top             =   4680
         Width           =   1755
         Begin VB.CheckBox chkRemoteSite 
            Caption         =   "Remote Site"
            Height          =   195
            Left            =   60
            TabIndex        =   13
            Top             =   480
            Width           =   1275
         End
         Begin VB.CheckBox chkServer 
            Caption         =   "Server"
            Height          =   195
            Left            =   60
            TabIndex        =   12
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.Frame fraScope 
         Caption         =   "Scope"
         Height          =   1275
         Left            =   60
         TabIndex        =   27
         Top             =   3360
         Width           =   1755
         Begin VB.CheckBox chkVisit 
            Caption         =   "Visit"
            Height          =   255
            Left            =   60
            TabIndex        =   10
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox chkSubject 
            Caption         =   "Subject"
            Height          =   195
            Left            =   60
            TabIndex        =   11
            Top             =   1020
            Width           =   915
         End
         Begin VB.CheckBox chkQuestion 
            Caption         =   "Question"
            Height          =   195
            Left            =   60
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkEform 
            Caption         =   "Eform"
            Height          =   195
            Left            =   60
            TabIndex        =   9
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame fraStatus 
         Caption         =   "Status"
         Height          =   975
         Left            =   60
         TabIndex        =   26
         Top             =   5520
         Width           =   1755
         Begin VB.CheckBox chkRefused 
            Caption         =   "Refused"
            Height          =   195
            Left            =   60
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkUnprocessed 
            Caption         =   "Unprocessed"
            Height          =   195
            Left            =   60
            TabIndex        =   16
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox chkProcessed 
            Caption         =   "Processed"
            Height          =   195
            Left            =   60
            TabIndex        =   15
            Top             =   480
            Width           =   1275
         End
      End
      Begin VB.Frame FraType 
         Caption         =   "Types"
         Height          =   1515
         Left            =   60
         TabIndex        =   25
         Top             =   1800
         Width           =   1755
         Begin VB.CheckBox chkRollback 
            Caption         =   "Rollback"
            Height          =   195
            Left            =   60
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkUnlock 
            Caption         =   "Unlock"
            Height          =   195
            Left            =   60
            TabIndex        =   7
            Top             =   1260
            Width           =   915
         End
         Begin VB.CheckBox chkLock 
            Caption         =   "Lock"
            Height          =   195
            Left            =   60
            TabIndex        =   5
            Top             =   1005
            Width           =   675
         End
         Begin VB.CheckBox chkUnfreeze 
            Caption         =   "Unfreeze"
            Height          =   195
            Left            =   60
            TabIndex        =   3
            Top             =   495
            Width           =   975
         End
         Begin VB.CheckBox chkFreeze 
            Caption         =   "Freeze"
            Height          =   195
            Left            =   60
            TabIndex        =   4
            Top             =   750
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Rese&t"
         Height          =   345
         Left            =   1200
         TabIndex        =   1
         Top             =   180
         Width           =   1125
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dates"
         Height          =   1035
         Left            =   60
         TabIndex        =   21
         Top             =   6480
         Width           =   2295
         Begin MSMask.MaskEdBox mskToDate 
            Height          =   375
            Left            =   1140
            TabIndex        =   18
            Top             =   540
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFromDate 
            Height          =   375
            Left            =   1140
            TabIndex        =   17
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "To date"
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   660
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "(dd/mm/yyyy)"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "From date"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   345
         Left            =   60
         TabIndex        =   0
         Top             =   180
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmViewLockFreeze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2003. All Rights Reserved
'   File:       frmViewLockFreeze.frm
'   Author:     Ashitei Trebi-Ollennu, January 2003
'   Purpose:    Viewer for lock/freeze messages
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   TA 8/1/03: resize code added
'   NCJ 14 Jan 03 - Don't display zero timestamps
'----------------------------------------------------------------------------------------'

Option Explicit

Private Const msDATE_DISPLAY_FORMAT = "yyyy/mm/dd hh:mm:ss"
Private Const msDateMaskDefault = "__/__/____"
Private Const msSetDateMask = "##/##/####"
Private Const msMidnight = ".9999884259"

Private Const msSOURCE_SERVER = "Server"
Private Const msSOURCE_REMOTESITE = "Remote Site"
Private Const msALL_SITES = "AllSites"
Private Const msALL_STUDIES = "AllStudies"

Private mbResetButtonClicked As Boolean
Private mColStatus As Collection
Private mColScope As Collection
Private mColSites As Collection
Private mColClinicalTrials As Collection
Private mColSource As Collection
Private mColType As Collection

'--------------------------------------------------------------------------
Private Sub BuildColumnHeaders()
'--------------------------------------------------------------------------
'builds column headers for listview
'--------------------------------------------------------------------------
Dim colmX As MSComctlLib.ColumnHeader
    
    On Error GoTo ErrHandler
    
    'clear listview
    lvwTransferStatus.ListItems.Clear
    
    'do not rebuild headers when the Reset button is clicked
    If mbResetButtonClicked Then Exit Sub
    
    'add column headers with widths
    Set colmX = lvwTransferStatus.ColumnHeaders.Add(, , "Study", 800)
    Set colmX = lvwTransferStatus.ColumnHeaders.Add(, , "Site", 800)
    Set colmX = lvwTransferStatus.ColumnHeaders.Add(, , "Subject", 1000)
    Set colmX = lvwTransferStatus.ColumnHeaders.Add(, , "Source", 1000)
    Set colmX = lvwTransferStatus.ColumnHeaders.Add(, , "Scope", 1000)
    Set colmX = lvwTransferStatus.ColumnHeaders.Add(, , "Message Type", 1500)
    Set colmX = lvwTransferStatus.ColumnHeaders.Add(, , "Visit", 800)
    Set colmX = lvwTransferStatus.ColumnHeaders.Add(, , "Eform", 800)
    Set colmX = lvwTransferStatus.ColumnHeaders.Add(, , "Question", 1000)
    Set colmX = lvwTransferStatus.ColumnHeaders.Add(, , "User Name", 1000)
    Set colmX = lvwTransferStatus.ColumnHeaders.Add(, , "Processed Timestamp", 2000)
    Set colmX = lvwTransferStatus.ColumnHeaders.Add(, , "Process Status", 1500)
 
    'set view type
    lvwTransferStatus.View = lvwReport
    'set initial sort to ascending on column 0 (study)
    lvwTransferStatus.SortKey = 0
    lvwTransferStatus.SortOrder = lvwAscending
    
Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.BuildColumnHeaders"
End Sub

'--------------------------------------------------------------------------
Private Sub LoadStudies()
'--------------------------------------------------------------------------
'adds available studies to the combo box and to the collection
'--------------------------------------------------------------------------
Dim oStudy As Study

    
    On Error GoTo ErrHandler
    
    Set mColClinicalTrials = New Collection
    
    If mbResetButtonClicked Then
        cboStudy.Clear
    End If
            
    If goUser.GetAllStudies.Count > 0 Then
    cboStudy.AddItem "AllStudies"
        For Each oStudy In goUser.GetAllStudies
            cboStudy.AddItem oStudy.StudyName
            mColClinicalTrials.Add oStudy.StudyName, oStudy.StudyName
        Next
    Else
        DialogWarning "There are no studies available to you"
        Exit Sub
    End If
    
    cboStudy.ListIndex = 0

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.LoadStudies"
End Sub

'--------------------------------------------------------------------------
Private Sub LoadSites()
'--------------------------------------------------------------------------
'initialises mColSites adds available sites to the combo box and to the collection
'--------------------------------------------------------------------------
Dim oSite As Site
    
    On Error GoTo ErrHandler
    
    Set mColSites = New Collection
    
    If mbResetButtonClicked Then
        cboSite.Clear
    End If
    
    If goUser.GetAllSites.Count > 0 Then
        cboSite.AddItem "AllSites"
        For Each oSite In goUser.GetAllSites
            cboSite.AddItem oSite.Site
            mColSites.Add oSite.Site, oSite.Site
        Next
    Else
        DialogWarning "There are no sites available to you"
        Exit Sub
    End If
    
    cboSite.ListIndex = 0

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.LoadSites"
End Sub

'--------------------------------------------------------------------------
Private Sub LoadSources()
'--------------------------------------------------------------------------
'initialises mColSource adds enums for source to mColSource collection
'--------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Set mColSource = New Collection
    
    mColSource.Add TypeOfInstallation.Server, msSOURCE_SERVER
    mColSource.Add TypeOfInstallation.RemoteSite, msSOURCE_REMOTESITE
    
Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.LoadSources"
End Sub

'--------------------------------------------------------------------------
Private Sub LoadScopes()
'--------------------------------------------------------------------------
'initialises mColScope adds enums for scopes to mColScope collection
'--------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Set mColScope = New Collection

    mColScope.Add LFScope.lfscEForm, str(LFScope.lfscEForm)
    mColScope.Add LFScope.lfscQuestion, str(LFScope.lfscQuestion)
    mColScope.Add LFScope.lfscSubject, str(LFScope.lfscSubject)
    mColScope.Add LFScope.lfscVisit, str(LFScope.lfscVisit)
    
Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.LoadScopes"
End Sub
'--------------------------------------------------------------------------
Public Sub Display()
'--------------------------------------------------------------------------
'adds to collections,sets default options and builds listview headers
'--------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    LoadScopes
    LoadSites
    LoadTypes
    LoadStudies
    LoadSources
    LoadStatus
    CheckAll
    BuildColumnHeaders
    
    'if reset button clicked do not show
    'form since its already being displayed.
    If Not mbResetButtonClicked Then
        Me.Show vbModal
    End If
        
Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.Display"
End Sub

'-----------------------------------------------------------------------
Private Sub chkEform_Click()
'-----------------------------------------------------------------------
'adds or removes from mColScope collection based on state of checkbox
'-----------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If chkEform.Value Then
        If Not CollectionMember(mColScope, str(LFScope.lfscEForm), False) Then
            mColScope.Add LFScope.lfscEForm, str(LFScope.lfscEForm)
        End If
    Else
         mColScope.Remove (str(LFScope.lfscEForm))
    End If
    
Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.chkEform_Click"
End Sub

'------------------------------------------------------------------
Private Sub chkFreeze_Click()
'------------------------------------------------------------------
'adds or removes from mColType collection based on state of checkbox
'------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If chkFreeze.Value Then
        If Not CollectionMember(mColType, str(LFAction.lfaFreeze), False) Then
            mColType.Add LFAction.lfaFreeze, str(LFAction.lfaFreeze)
        End If
    Else
         mColType.Remove (str(LFAction.lfaFreeze))
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.chkFreeze_Click"
End Sub

'-------------------------------------------------------------------
Private Sub chkLock_Click()
'-------------------------------------------------------------------
'adds or removes from mColType collection based on state of checkbox
'-------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If chkLock.Value Then
        If Not CollectionMember(mColType, str(LFAction.lfaLock), False) Then
            mColType.Add LFAction.lfaLock, str(LFAction.lfaLock)
        End If
    Else
         mColType.Remove (str(LFAction.lfaLock))
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.chkLock_Click"
End Sub

'-------------------------------------------------------------------
Private Sub chkProcessed_Click()
'-------------------------------------------------------------------
'adds or removes from mColStatus collection based on state of checkbox
'-------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If chkProcessed.Value Then
        If Not CollectionMember(mColStatus, str(LFProcessStatus.lfpProcessed), False) Then
            mColStatus.Add LFProcessStatus.lfpProcessed, str(LFProcessStatus.lfpProcessed)
        End If
    Else
         mColStatus.Remove (str(LFProcessStatus.lfpProcessed))
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.chkProcessed_Click"
End Sub

'--------------------------------------------------------------------
Private Sub chkQuestion_Click()
'--------------------------------------------------------------------
'adds or removes from mColScope collection based on state of checkbox
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If chkQuestion.Value Then
        If Not CollectionMember(mColScope, str(LFScope.lfscQuestion), False) Then
            mColScope.Add LFScope.lfscQuestion, str(LFScope.lfscQuestion)
        End If
    Else
         mColScope.Remove (str(LFScope.lfscQuestion))
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.chkQuestion_Click"
End Sub

'---------------------------------------------------------------------
Private Sub chkRefused_Click()
'---------------------------------------------------------------------
'adds or removes from mColStatus collection based on state of checkbox
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If chkRefused.Value Then
        If Not CollectionMember(mColStatus, str(LFProcessStatus.lfpRefused), False) Then
            mColStatus.Add LFProcessStatus.lfpRefused, str(LFProcessStatus.lfpRefused)
        End If
    Else
         mColStatus.Remove (str(LFProcessStatus.lfpRefused))
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.chkRefused_Click"
End Sub

'-------------------------------------------------------------------
Private Sub chkRemoteSite_Click()
'-------------------------------------------------------------------
'adds or removes from mColSource collection based on state of checkbox
'-------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If chkRemoteSite.Value Then
        If Not CollectionMember(mColSource, msSOURCE_REMOTESITE, False) Then
            mColSource.Add TypeOfInstallation.RemoteSite, msSOURCE_REMOTESITE
        End If
    Else
         mColSource.Remove (msSOURCE_REMOTESITE)
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.chkRemoteSite_Click"
End Sub

'---------------------------------------------------------------------
Private Sub chkRollback_Click()
'---------------------------------------------------------------------
'adds or removes from mColType collection based on state of checkbox
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If chkRollback.Value Then
        If Not CollectionMember(mColType, str(LFAction.lfaRollback), False) Then
            mColType.Add LFAction.lfaRollback, str(LFAction.lfaRollback)
        End If
    Else
         mColType.Remove (str(LFAction.lfaRollback))
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.chkRollback_Click"
End Sub

'--------------------------------------------------------------------
Private Sub chkServer_Click()
'--------------------------------------------------------------------
'adds or removes from mColSource collection based on state of checkbox
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If chkServer.Value Then
        If Not CollectionMember(mColSource, msSOURCE_SERVER, False) Then
            mColSource.Add TypeOfInstallation.Server, msSOURCE_SERVER
        End If
    Else
         mColSource.Remove (msSOURCE_SERVER)
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.chkServer_Click"
End Sub

'-----------------------------------------------------------------
Private Sub chkSubject_Click()
'-----------------------------------------------------------------
'adds or removes from mColScope collection based on state of checkbox
'-----------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If chkSubject.Value Then
        If Not CollectionMember(mColScope, str(LFScope.lfscSubject), False) Then
            mColScope.Add LFScope.lfscSubject, str(LFScope.lfscSubject)
        End If
    Else
         mColScope.Remove (str(LFScope.lfscSubject))
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.chkSubject_Click"
End Sub


'------------------------------------------------------------------
Private Sub chkUnfreeze_Click()
'-------------------------------------------------------------------
'adds or removes from mColType collection based on state of checkbox
'-------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If chkUnfreeze.Value Then
        If Not CollectionMember(mColType, str(LFAction.lfaUnfreeze), False) Then
            mColType.Add LFAction.lfaUnfreeze, str(LFAction.lfaUnfreeze)
        End If
    Else
         mColType.Remove (str(LFAction.lfaUnfreeze))
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.chkUnfreeze_Click"
End Sub

'-------------------------------------------------------------------
Private Sub chkUnlock_Click()
'-------------------------------------------------------------------
'adds or removes from mColType collection based on state of checkbox
'-------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If chkUnlock.Value Then
        If Not CollectionMember(mColType, str(LFAction.lfaUnlock), False) Then
            mColType.Add LFAction.lfaUnlock, str(LFAction.lfaUnlock)
        End If
    Else
         mColType.Remove (str(LFAction.lfaUnlock))
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.chkUnlock_Click"
End Sub

'------------------------------------------------------------------
Private Sub chkUnprocessed_Click()
'------------------------------------------------------------------
'adds or removes from mColStatus collection based on state of checkbox
'------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If chkUnprocessed.Value Then
        If Not CollectionMember(mColStatus, str(LFProcessStatus.lfpUnProcessed), False) Then
            mColStatus.Add LFProcessStatus.lfpUnProcessed, str(LFProcessStatus.lfpUnProcessed)
        End If
    Else
         mColStatus.Remove (str(LFProcessStatus.lfpUnProcessed))
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.chkUnprocessed_Click"
End Sub

'--------------------------------------------------------------------
Private Sub chkVisit_Click()
'--------------------------------------------------------------------
'adds or removes from mColScope collection based on state of checkbox
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If chkVisit.Value Then
        If Not CollectionMember(mColScope, str(LFScope.lfscVisit), False) Then
            mColScope.Add LFScope.lfscVisit, str(LFScope.lfscVisit)
        End If
    Else
         mColScope.Remove (str(LFScope.lfscVisit))
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.chkVisit_Click"
End Sub

'-------------------------------------------------------------------------
Private Sub cmdOK_Click()
'-------------------------------------------------------------------------
'
'-------------------------------------------------------------------------

    Unload Me

End Sub


'------------------------------------------------------------------------
Private Sub cmdRefresh_Click()
'------------------------------------------------------------------------
'
'------------------------------------------------------------------------

    LoadListView

End Sub

'-------------------------------------------------------------------------
Private Sub cmdReset_Click()
'-------------------------------------------------------------------------
'sets form and controls to default status
'-------------------------------------------------------------------------
    
    mbResetButtonClicked = True
    mskFromDate.Text = msDateMaskDefault
    mskToDate.Text = msDateMaskDefault
    cboStudy.ListIndex = 0
    cboSite.ListIndex = 0
    Display
    mbResetButtonClicked = False
    
End Sub

'--------------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
    
    Me.Icon = frmMenu.Icon
    mskFromDate.Mask = msSetDateMask
    mskToDate.Mask = msSetDateMask

        'set size, position and window state
    Call SetFormDimensions(Me)

End Sub

'--------------------------------------------------------------------------
Private Sub LoadTypes()
'--------------------------------------------------------------------------
'initialises mColType adds enums to mColType collection
'--------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Set mColType = New Collection
    
    mColType.Add LFAction.lfaFreeze, str(LFAction.lfaFreeze)
    mColType.Add LFAction.lfaLock, str(LFAction.lfaLock)
    mColType.Add LFAction.lfaRollback, str(LFAction.lfaRollback)
    mColType.Add LFAction.lfaUnfreeze, str(LFAction.lfaUnfreeze)
    mColType.Add LFAction.lfaUnlock, str(LFAction.lfaUnlock)
    
Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.LoadTypes"
End Sub
'---------------------------------------------------------------------------
Private Sub LoadStatus()
'---------------------------------------------------------------------------
'initialises mColStatus adds enums to mColStatus collection
'---------------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    Set mColStatus = New Collection
    
    mColStatus.Add LFProcessStatus.lfpProcessed, str(LFProcessStatus.lfpProcessed)
    mColStatus.Add LFProcessStatus.lfpRefused, str(LFProcessStatus.lfpRefused)
    mColStatus.Add LFProcessStatus.lfpUnProcessed, str(LFProcessStatus.lfpUnProcessed)

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.LoadStatus"
End Sub

'------------------------------------------------------------------------
Private Function CreateSQLSiteString() As String
'------------------------------------------------------------------------
'creates and returns the sites to be used in the SQL to load the listview
'------------------------------------------------------------------------
Dim sSiteString As String
Dim sSite As String
Dim vSite As Variant
Dim i As Integer

    On Error GoTo ErrHandler
    
    CreateSQLSiteString = ""
    sSiteString = ""
    
    'build the sites for sql
    If cboSite.Text = msALL_SITES Then
        If mColSites.Count <> 0 Then
            For i = 1 To mColSites.Count
                If sSiteString = "" Then
                    sSiteString = Chr(39) & mColSites.Item(i) & Chr(39)
                Else
                    sSiteString = sSiteString & "," & Chr(39) & mColSites.Item(i) & Chr(39)
                End If
            Next
        End If
    Else
        sSiteString = Chr(39) & Trim(cboSite.Text) & Chr(39)
    End If
    
    CreateSQLSiteString = sSiteString
            
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.CreateSQLSiteString"
End Function

'------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'------------------------------------------------------------------------
    'store window dimensions
'------------------------------------------------------------------------
    
    Call SaveFormDimensions(Me)

End Sub

'------------------------------------------------------------------------
Private Sub Form_Resize()
'------------------------------------------------------------------------

'------------------------------------------------------------------------

    On Error Resume Next
    
    lvwTransferStatus.Width = Me.ScaleWidth - fraSelection.Width - 180
    lvwTransferStatus.Height = Me.ScaleHeight - cmdOK.Height - 240
    cmdOK.Left = Me.ScaleWidth - cmdOK.Width - 60
    cmdOK.Top = lvwTransferStatus.Height + 120

End Sub

'------------------------------------------------------------------------------------------
Private Sub lvwTransferStatus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'------------------------------------------------------------------------------------------
'allows re-ordering of listview items when column header is clicked
'------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Call lvw_Sort(lvwTransferStatus, ColumnHeader)

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.lvwTransferStatus"
End Sub

'------------------------------------------------------------------------
Private Function CreateSQLStatusString() As String
'------------------------------------------------------------------------
'creates and returns the statuses to be used in the SQL to load the listview
'------------------------------------------------------------------------
Dim sStatus As String
Dim n As Integer

    On Error GoTo ErrHandler
    
    CreateSQLStatusString = ""
    sStatus = ""
    
    'build the statuses for sql
    If mColStatus.Count <> 0 Then
        For n = 1 To mColStatus.Count
            If sStatus = "" Then
                sStatus = mColStatus.Item(n)
            Else
                sStatus = sStatus & "," & mColStatus.Item(n)
            End If
        Next
    End If
    
    CreateSQLStatusString = sStatus
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.CreateSQLStatusString"
End Function

'---------------------------------------------------------------------------
Private Sub LoadListView()
'---------------------------------------------------------------------------
'loads listview by calling routines that create sql required to populate listview
' NCJ 14 Jan 03 - Don't show Processed timestamp if it's 0
'---------------------------------------------------------------------------
Dim itmX  As MSComctlLib.ListItem
Dim sSite As String
Dim sStudies As String
Dim sStatus As String
Dim sSource As String
Dim sScope As String
Dim sTypeString As String
Dim sType As String
Dim sDate As String
Dim sSQL As String
Dim rsLFRecords As ADODB.Recordset
Dim sMsg As String
Dim sVisit As String
Dim sQuestion As String
Dim sForm As String
Dim oLF As LockFreeze

    On Error GoTo ErrHandler
    
    Set oLF = New LockFreeze
    
    'inform if  to-date > from-date
    sMsg = "The To date entered is earlier than From date"
    If mskToDate.Text <> msDateMaskDefault Then
        If mskFromDate.Text <> msDateMaskDefault Then
            If ConvertLocalNumToStandard(CStr(CDbl((CDate(mskToDate.Text))))) < _
                ConvertLocalNumToStandard(CStr(CDbl((CDate(mskFromDate.Text))))) Then
                Call DialogInformation(sMsg, "Date Error")
                Exit Sub
            End If
        End If
    End If
    
    'clear the list view
    lvwTransferStatus.ListItems.Clear
    
    'initialise variables
    sType = ""
    sSite = ""
    sStudies = ""
    sSource = ""
    sScope = ""
    sStatus = ""
    
    'build the selected study(ies) for use in SQL
    sStudies = CreateSQLStudiesString
    
    'build the selected site(s) for use in SQL
    sSite = CreateSQLSiteString
    
    'build the selected status(es) for use in SQL
    sStatus = CreateSQLStatusString
    
    'build the selected source(s) for use in SQL
    sSource = CreateSQLSourceString
    
    'build the selected type(es) for use in SQL
    sType = CreateSQLTypeString
    
    'build the selected scope(s) for use in SQL
    sScope = CreateSQLScopeString
    
    'get dates if any entered
    sDate = CreateSQLDatesString
    
    sSQL = "SELECT * FROM LFMessage "
    sSQL = sSQL & "WHERE ClinicalTrialName IN (" & sStudies & ")"
    sSQL = sSQL & "AND TrialSite IN (" & sSite & ")"
    sSQL = sSQL & " AND " & goUser.DataLists.StudiesSitesWhereSQL("CLINICALTRIALID", "TRIALSITE")
    If sSource <> "" Then
        sSQL = sSQL & "AND SOURCE IN (" & sSource & ")"
    End If
    If sScope <> "" Then
        sSQL = sSQL & "AND SCOPE IN (" & sScope & ")"
    End If
    If sType <> "" Then
        sSQL = sSQL & "AND MSGTYPE IN (" & sType & ")"
    End If
    If sStatus <> "" Then
        sSQL = sSQL & "AND ProcessedStatus IN (" & sStatus & ")"
    End If
    If sDate <> "" Then
        sSQL = sSQL & sDate
    End If
    
    sSQL = sSQL & " ORDER BY ProcessedTimestamp DESC"
    
    Set rsLFRecords = New ADODB.Recordset
    rsLFRecords.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    If rsLFRecords.RecordCount <= 0 Then
        Call DialogInformation("No records retrieved for selected criteria.")
        Exit Sub
    End If
    
    rsLFRecords.MoveFirst
    Do Until rsLFRecords.EOF
        
        If Not IsNull(rsLFRecords!ClinicalTrialName) Then
            Set itmX = lvwTransferStatus.ListItems.Add(, , rsLFRecords!ClinicalTrialName)
        End If
        
        If Not IsNull(rsLFRecords!TrialSite) Then
            itmX.SubItems(1) = rsLFRecords!TrialSite
        End If
        
        If Not IsNull(rsLFRecords!PersonId) Then
            itmX.SubItems(2) = rsLFRecords!PersonId
        End If
        
        If Not IsNull(rsLFRecords!Source) Then
            itmX.SubItems(3) = DecodeSource(rsLFRecords!Source)
        End If
        
        If Not IsNull(rsLFRecords!Scope) Then
            itmX.SubItems(4) = oLF.ScopeText(rsLFRecords!Scope)
        End If
        
        If Not IsNull(rsLFRecords!MSGTYPE) Then
            itmX.SubItems(5) = oLF.ActionText(rsLFRecords!MSGTYPE)
        End If
        
        If Val(RemoveNull(rsLFRecords!VisitId)) <> 0 Then
            sVisit = goUser.DataLists.GetStudyItemCode(soVisit, rsLFRecords!ClinicalTrialId, rsLFRecords!VisitId)
            If rsLFRecords!VisitCycleNumber = 1 Then
                itmX.SubItems(6) = sVisit
            Else
                itmX.SubItems(6) = sVisit & "[" & rsLFRecords!VisitCycleNumber & "]"
            End If
        End If
        
        If Val(RemoveNull(rsLFRecords!crfpageid)) <> 0 Then
            sForm = goUser.DataLists.GetStudyItemCode(soeform, rsLFRecords!ClinicalTrialId, rsLFRecords!crfpageid)
            If rsLFRecords!CRFPageCycleNumber = 1 Then
                itmX.SubItems(7) = sForm
            Else
                itmX.SubItems(7) = sForm & "[" & rsLFRecords!CRFPageCycleNumber & "]"
            End If
        End If
        
        If Val(RemoveNull(rsLFRecords!DataItemId)) <> 0 Then
            sQuestion = goUser.DataLists.GetStudyItemCode(soQuestion, rsLFRecords!ClinicalTrialId, rsLFRecords!DataItemId)
            If rsLFRecords!RepeatNumber = 1 Then
                itmX.SubItems(8) = sQuestion
            Else
                itmX.SubItems(8) = sQuestion & "[" & rsLFRecords!RepeatNumber & "]"
            End If

        End If
       
        If Not IsNull(rsLFRecords!UserNameFull) Then
            itmX.SubItems(9) = rsLFRecords!UserNameFull
        End If
        
        If Val(RemoveNull(rsLFRecords!ProcessedTimestamp)) <> 0 Then
            ' Only show Processed date if non-zero
            itmX.SubItems(10) = Format$(rsLFRecords![ProcessedTimestamp], msDATE_DISPLAY_FORMAT)
        End If
        
        If Not IsNull(rsLFRecords!ProcessedStatus) Then
            itmX.SubItems(11) = oLF.StatusText(rsLFRecords!ProcessedStatus)
        End If
        rsLFRecords.MoveNext
    Loop
    
    Call lvw_SetAllColWidths(lvwTransferStatus, LVSCW_AUTOSIZE_USEHEADER)
    
    Set oLF = Nothing
Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.LoadListView"
End Sub

'----------------------------------------------------------------------
Private Function CreateSQLStudiesString() As String
'----------------------------------------------------------------------
'creates and returns the study(ies) to be used in the SQL to load the listview
'----------------------------------------------------------------------
Dim sStudyString As String
Dim sStudies As String
Dim vStudy As Variant
Dim i As Integer

    On Error GoTo ErrHandler
    
    CreateSQLStudiesString = ""
    sStudyString = ""
    
    'build the studies for sql
    If cboStudy.Text = msALL_STUDIES Then
        If mColClinicalTrials.Count <> 0 Then
            For i = 1 To mColClinicalTrials.Count
                If sStudyString = "" Then
                    sStudyString = Chr(39) & mColClinicalTrials.Item(i) & Chr(39)
                Else
                    sStudyString = sStudyString & "," & Chr(39) & mColClinicalTrials.Item(i) & Chr(39)
                End If
            Next
        End If
    Else
        sStudyString = Chr(39) & Trim(cboStudy.Text) & Chr(39)
    End If
    
    CreateSQLStudiesString = sStudyString
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.CreateSQLStudiesString"
End Function

'----------------------------------------------------------------------
Private Function CreateSQLSourceString() As String
'----------------------------------------------------------------------
'creates and returns the source(s) to be used in the SQL to load the listview
'----------------------------------------------------------------------
Dim sSourceString As String
Dim sSource As String
Dim i As Integer
    
    On Error GoTo ErrHandler
    
    CreateSQLSourceString = ""
    
    If mColSource.Count <> 0 Then
        For i = 1 To mColSource.Count
            If sSource = "" Then
                sSource = mColSource.Item(i)
            Else
                sSource = sSource & "," & mColSource.Item(i)
            End If
        Next
    End If
    
    CreateSQLSourceString = sSource

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.CreateSQLSourceString"
End Function

'----------------------------------------------------------------------
Private Function CreateSQLTypeString() As String
'----------------------------------------------------------------------
'creates and returns the type(s) enums to be used in the SQL to load the listview
'----------------------------------------------------------------------
Dim sTypeString As String
Dim sType As String
Dim i As Integer
    
    On Error GoTo ErrHandler
    
    CreateSQLTypeString = ""
    
    If mColType.Count <> 0 Then
        For i = 1 To mColType.Count
            If sType = "" Then
                sType = mColType.Item(i)
            Else
                sType = sType & "," & mColType.Item(i)
            End If
        Next
    End If
    
    CreateSQLTypeString = sType

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.CreateSQLTypeString"
End Function

'----------------------------------------------------------------------
Private Function CreateSQLScopeString() As String
'----------------------------------------------------------------------
'creates and returns the scope(s) enums to be used in the SQL to load the listview
'----------------------------------------------------------------------
Dim sScope As String
Dim i As Integer
    
    On Error GoTo ErrHandler
    
    CreateSQLScopeString = ""
    
    If mColScope.Count <> 0 Then
        For i = 1 To mColScope.Count
            If sScope = "" Then
                sScope = mColScope.Item(i)
            Else
                sScope = sScope & "," & mColScope.Item(i)
            End If
        Next
    End If
    
    CreateSQLScopeString = sScope

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.CreateSQLScopeString"
End Function

'--------------------------------------------------------------------------
Private Function DecodeSource(nNum As Integer) As String
'--------------------------------------------------------------------------
'decodes source enums so as to display corresponding text in listview
'--------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    DecodeSource = ""
    Select Case nNum
        Case TypeOfInstallation.Server
            DecodeSource = msSOURCE_SERVER
        Case TypeOfInstallation.RemoteSite
            DecodeSource = msSOURCE_REMOTESITE
        End Select
        
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.DecodeSource"
End Function

'-----------------------------------------------------------------------------------
Private Function CreateSQLDatesString() As String
'-----------------------------------------------------------------------------------
'creates and returns the processTimeStamp to be used in the SQL to load the listview
'-----------------------------------------------------------------------------------
Dim sSQL As String
Dim sToDate As String
Dim sFromDate As String
    
    On Error GoTo ErrHandler
    
    CreateSQLDatesString = ""
    
    If mskToDate.Text <> msDateMaskDefault Then
        sToDate = ConvertLocalNumToStandard(CStr(CDbl((CDate(mskToDate.Text)))))
    End If
    
    If mskFromDate.Text <> msDateMaskDefault Then
        sFromDate = ConvertLocalNumToStandard(CStr(CDbl((CDate(mskFromDate.Text)))))
    End If

    'both date fields empty
    If mskToDate = msDateMaskDefault And mskFromDate = msDateMaskDefault Then
        CreateSQLDatesString = ""
        Exit Function
    
    'only from-date entered
    ElseIf mskFromDate <> msDateMaskDefault And mskToDate = msDateMaskDefault Then
        sSQL = sSQL & " AND PROCESSEDTIMESTAMP >= " & sFromDate
        CreateSQLDatesString = sSQL
    
    'only to-date entered
    ElseIf mskFromDate = msDateMaskDefault And mskToDate <> msDateMaskDefault Then
        sSQL = sSQL & " AND PROCESSEDTIMESTAMP <= " & sToDate & msMidnight
        CreateSQLDatesString = sSQL
    'both dates entered
    Else
        sSQL = sSQL & " AND PROCESSEDTIMESTAMP >= " & sFromDate & " AND PROCESSEDTIMESTAMP <= " & sToDate & msMidnight
        CreateSQLDatesString = sSQL
    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.CreateSQLDatesString"
End Function

'--------------------------------------------------------------------------------
Private Sub CheckAll()
'--------------------------------------------------------------------------------
'default status of selection criteria
'--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    'scope
    chkEform.Value = 1
    chkQuestion.Value = 1
    chkVisit.Value = 1
    chkSubject.Value = 1
    'status
    chkProcessed.Value = 1
    chkUnprocessed.Value = 1
    chkRefused.Value = 1
    'types
    chkFreeze.Value = 1
    chkUnfreeze.Value = 1
    chkLock.Value = 1
    chkUnlock.Value = 1
    chkRollback.Value = 1
    'source
    chkServer.Value = 1
    chkRemoteSite.Value = 1

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.CheckAll"
End Sub

'--------------------------------------------------------------------------
Private Sub mskFromDate_LostFocus()
'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Dim sMsg As String

     On Error GoTo ErrHandler

    sMsg = "The date " & mskFromDate.Text & " is not a valid date"
    If mskFromDate.Text <> msDateMaskDefault Then
        If Not IsDate(mskFromDate.Text) Then
            Call DialogInformation(sMsg, "Date Error")
            mskFromDate.SetFocus
            Exit Sub
        End If
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.mskFromDate_LostFocus"
End Sub

'--------------------------------------------------------------------------------
Private Sub mskToDate_LostFocus()
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
Dim sMsg As String

    On Error GoTo ErrHandler
       
    sMsg = "The date " & mskToDate.Text & " is not a valid date"
    If mskToDate.Text <> msDateMaskDefault Then
        If Not IsDate(mskToDate.Text) Then
            Call DialogInformation(sMsg, "Date Error")
            mskToDate.SetFocus
            Exit Sub
        End If
    End If

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmViewLockFreeze.mskToDate_LostFocus"
End Sub
