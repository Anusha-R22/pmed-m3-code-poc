VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabDefinitions 
   Caption         =   "Normal Ranges and CTC"
   ClientHeight    =   7125
   ClientLeft      =   3840
   ClientTop       =   2865
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   6435
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4860
      TabIndex        =   9
      Top             =   5460
      Width           =   1215
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4860
      TabIndex        =   10
      Top             =   5940
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   375
      Left            =   4860
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   4860
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4860
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4860
      TabIndex        =   11
      Top             =   6660
      Width           =   1215
   End
   Begin TabDlg.SSTab ssTabLabResult 
      Height          =   6495
      Left            =   60
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   60
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Test Groups"
      TabPicture(0)   =   "frmLabDefinitions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDetails(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraList(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tests"
      TabPicture(1)   =   "frmLabDefinitions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraList(1)"
      Tab(1).Control(1)=   "fraDetails(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Laboratories"
      TabPicture(2)   =   "frmLabDefinitions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDetails(2)"
      Tab(2).Control(1)=   "fraList(2)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "CTC Schemes"
      TabPicture(3)   =   "frmLabDefinitions.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraList(3)"
      Tab(3).Control(1)=   "fraDetails(3)"
      Tab(3).ControlCount=   2
      Begin VB.Frame fraList 
         Caption         =   "CTC Schemes"
         Height          =   2895
         Index           =   3
         Left            =   -74880
         TabIndex        =   40
         Top             =   420
         Width           =   6015
         Begin VB.CommandButton cmdCTC 
            Caption         =   "Edit Criteria..."
            Height          =   375
            Left            =   4680
            TabIndex        =   1
            Top             =   240
            Width           =   1215
         End
         Begin MSComctlLib.ListView lvwList 
            Height          =   2535
            Index           =   3
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   4471
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
      Begin VB.Frame fraDetails 
         Caption         =   "CTC Scheme"
         Height          =   2955
         Index           =   3
         Left            =   -74880
         TabIndex        =   35
         Top             =   3420
         Width           =   6015
         Begin VB.TextBox txtCode 
            Height          =   315
            Index           =   3
            Left            =   720
            MaxLength       =   15
            TabIndex        =   37
            Top             =   300
            Width           =   1635
         End
         Begin VB.TextBox txtDescription 
            Height          =   1695
            Index           =   3
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   1140
            Width           =   4455
         End
         Begin VB.Label Label3 
            Caption         =   "Code"
            Height          =   315
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Description"
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   780
            Width           =   1335
         End
      End
      Begin VB.Frame fraList 
         Caption         =   "Laboratories"
         Height          =   2895
         Index           =   2
         Left            =   -74880
         TabIndex        =   33
         Top             =   420
         Width           =   6015
         Begin VB.CommandButton cmdNormalRange 
            Caption         =   "Edit &Ranges..."
            Height          =   375
            Left            =   4680
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
         Begin MSComctlLib.ListView lvwList 
            Height          =   2535
            Index           =   2
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   4471
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
      Begin VB.Frame fraDetails 
         Caption         =   "Laboratory"
         Height          =   2955
         Index           =   2
         Left            =   -74880
         TabIndex        =   28
         Top             =   3420
         Width           =   6015
         Begin VB.TextBox txtDescription 
            Height          =   1695
            Index           =   2
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   30
            Top             =   1140
            Width           =   4455
         End
         Begin VB.TextBox txtCode 
            Height          =   315
            Index           =   2
            Left            =   720
            MaxLength       =   15
            TabIndex        =   29
            Top             =   300
            Width           =   1635
         End
         Begin VB.Label Label1 
            Caption         =   "Description"
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label lblLAb 
            Caption         =   "Code"
            Height          =   315
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame fraList 
         Caption         =   "Tests"
         Height          =   2895
         Index           =   1
         Left            =   -74880
         TabIndex        =   26
         Top             =   420
         Width           =   6015
         Begin MSComctlLib.ListView lvwList 
            Height          =   2535
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   4471
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
      Begin VB.Frame fraDetails 
         Caption         =   "Test"
         Height          =   2955
         Index           =   1
         Left            =   -74880
         TabIndex        =   19
         Top             =   3420
         Width           =   6015
         Begin VB.TextBox txtDescription 
            Height          =   1215
            Index           =   1
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   1200
            Width           =   4395
         End
         Begin VB.TextBox txtCode 
            Height          =   315
            Index           =   1
            Left            =   720
            MaxLength       =   15
            TabIndex        =   20
            Top             =   360
            Width           =   1635
         End
         Begin VB.ComboBox cboGroups 
            Height          =   315
            Left            =   3240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   2460
            Width           =   1275
         End
         Begin VB.ComboBox cboUnits 
            Height          =   315
            Left            =   600
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2460
            Width           =   1155
         End
         Begin VB.CommandButton cmdUnits 
            Caption         =   "..."
            Height          =   315
            Left            =   1800
            TabIndex        =   7
            Top             =   2460
            Width           =   315
         End
         Begin VB.Label Label6 
            Caption         =   "Description"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Code"
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label8 
            Caption         =   "Group"
            Height          =   315
            Left            =   2640
            TabIndex        =   23
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "Units"
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Top             =   2520
            Width           =   495
         End
      End
      Begin VB.Frame fraList 
         Caption         =   "Test Groups"
         Height          =   2895
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   420
         Width           =   6015
         Begin MSComctlLib.ListView lvwList 
            Height          =   2535
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   4471
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
      Begin VB.Frame fraDetails 
         Caption         =   "Test Group"
         Height          =   2955
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   3420
         Width           =   6015
         Begin VB.TextBox txtCode 
            Height          =   315
            Index           =   0
            Left            =   720
            MaxLength       =   15
            TabIndex        =   12
            Top             =   300
            Width           =   1635
         End
         Begin VB.TextBox txtDescription 
            Height          =   1695
            Index           =   0
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   1200
            Width           =   4455
         End
         Begin VB.Label Label4 
            Caption         =   "Code"
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Description"
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   780
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmLabDefinitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       frmLabDefinitions.frm
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, September 2000
'   Purpose:    Form for administering Normal Ranges and CTC Criteria
'----------------------------------------------------------------------------------------'
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
'   TA 13/10/2000: Repopulating Groups combo sorted
'   TA 18/10/2000: Put Test Group and Test tabs before Lab and Scheme
'   TA 24/10/2000: Button order and position changed
'   TA 09/11/2000: Now in DM - if in DM vSite contains site code else vSite contains null
'   DPH 15/10/2001 - Check Permissions for Normal Range and CTC

Option Explicit

'constants coresponding to tabs for each onject type
Private Const m_TAB_CLINICALTESTGROUP = 0
Private Const m_TAB_CLINICALTEST = 1
Private Const m_TAB_LAB = 2
Private Const m_TAB_CTCSCHEME = 3

'contain the diffrent object collections
Private moLabs As clsLabs

'only allow schemes in SD/LM
#If SD Then
Private moSchemes As clsCTCSchemes
#End If

Private moClinicalTestGroups As clsClinTestGroups
Private moClinicalTests As clsClinicalTests

'hold whether description should be updated along with code
Private mbUpdateDescWithCode As Boolean

'can only edit labs and NR when site is not null
Private mvSite As Variant

'TA September 2000
'moCol is the generic collection that can be used to edit Laboratories, CTC Schemes, Clinical Test Groups and Clinical Tests
'it assumes that all these collection classes have Delete, Insert, Update, Item and PopulateListview methods
'Ideally a common interface should have been created for these classes (maybe not for Clinical Tests)
' but for now I have opted for late binding instead
Private moCol As Object

'create was pressed rather than edit
Private mbCreate As Boolean

'initial form height and width
Private mlHeight As Long
Private mlWidth As Long

'---------------------------------------------------------------------
Private Sub Clear()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    txtCode(ssTabLabResult.Tab).Text = ""
    txtDescription(ssTabLabResult.Tab).Text = ""
    If ssTabLabResult.Tab = m_TAB_CLINICALTEST Then
        SetUnitsComboText ""
        If cboGroups.ListCount > 0 Then
            cboGroups.ListIndex = 0
        End If
    End If
    
    'set up link between desc and code
    mbUpdateDescWithCode = True
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdCreate_Click()
'---------------------------------------------------------------------

'---------------------------------------------------------------------
    
    mbCreate = True
    Call EditMode(True)
    'creating so disable highlight in list view
    lvwList(ssTabLabResult.Tab).HideSelection = True
    Clear
    'turn on update description with code
    mbUpdateDescWithCode = True
    
End Sub


'only allow in SD/LM
#If SD Then
'---------------------------------------------------------------------
Private Sub cmdCTC_Click()
'---------------------------------------------------------------------
' display CTC form if there are clinical tests
'---------------------------------------------------------------------

    If moClinicalTests Is Nothing Then
        Set moClinicalTests = New clsClinicalTests
    End If

    If moClinicalTests.Count > 0 Then
        frmCTC.Display moSchemes.Item(lvwList(ssTabLabResult.Tab).SelectedItem.Tag), moClinicalTests, moClinicalTestGroups
    Else
        DialogError "You must set up at least one Clinical Test"
    End If
    lvwList(m_TAB_CTCSCHEME).SetFocus
End Sub
#End If

'---------------------------------------------------------------------
Private Sub cmdDelete_Click()
'---------------------------------------------------------------------
'delete a Test, Group, Scheme or Lab
'---------------------------------------------------------------------
Dim sMessage As String
    
    On Error GoTo ErrHandler
    
    With moCol
        If DialogQuestion("Do you wish to delete the item " & lvwList(ssTabLabResult.Tab).SelectedItem.Tag & "?") = vbYes Then
            If .Delete(lvwList(ssTabLabResult.Tab).SelectedItem.Tag, sMessage) Then
                'TA 3/10/00: following code overidden currently
                If lvwList(ssTabLabResult.Tab).ListItems.Count <= 1 Then
                    'no further listitems
                Else
                    If lvwList(ssTabLabResult.Tab).SelectedItem.Index = 1 Then
                        'first one deleted - selcet 2nd
                        lvwList(ssTabLabResult.Tab).SelectedItem = lvwList(ssTabLabResult.Tab).ListItems(2)
                    Else
                        'not first - select the one before it
                       lvwList(ssTabLabResult.Tab).SelectedItem = lvwList(ssTabLabResult.Tab).ListItems(lvwList(ssTabLabResult.Tab).SelectedItem.Index - 1)
                    End If
                End If
                txtCode(ssTabLabResult.Tab).Text = ""
                txtDescription(ssTabLabResult.Tab).Text = ""
                ReDisplay
            Else
                DialogError sMessage
            End If
        End If
    End With
    
Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdDelete_Click")
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
Private Sub cmdEdit_Click()
'---------------------------------------------------------------------

'---------------------------------------------------------------------
    mbCreate = False
    Call EditMode(True)
    'disallow editing of code
    txtCode(ssTabLabResult.Tab).Enabled = False
    'do we want to set the background color here?  lvwList(ssTabLabResult.Tab).backcolor=vbwindowsbackground

End Sub



'---------------------------------------------------------------------
Private Sub cmdNormalRange_Click()
'---------------------------------------------------------------------
' display CTC form if there are clinical tests
'---------------------------------------------------------------------
    If moClinicalTests Is Nothing Then
        Set moClinicalTests = New clsClinicalTests
    End If
    
    If moClinicalTests.Count > 0 Then
        'if user's site is same as lab site then allow edits
        Call frmNormalRanges.Display(moLabs.Item(lvwList(ssTabLabResult.Tab).SelectedItem.Tag), _
                moClinicalTests, moClinicalTestGroups, _
                (moLabs.Item(lvwList(m_TAB_LAB).SelectedItem.Tag).SiteText = RemoveNull(mvSite)))
    Else
        DialogError "You must set up at least one Clinical Test"
    End If
    lvwList(m_TAB_LAB).SetFocus
End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()

'---------------------------------------------------------------------
    Unload Me
End Sub

'---------------------------------------------------------------------
Private Sub cmdSave_Click()
'---------------------------------------------------------------------
'save or update data
'---------------------------------------------------------------------
Dim bError As Boolean
Dim sTag As String
    On Error GoTo ErrHandler
    
    sTag = ""
    
    Select Case ssTabLabResult.Tab
    Case m_TAB_CLINICALTEST
        If mbCreate Then
            bError = Not moClinicalTests.Insert(txtCode(ssTabLabResult.Tab), txtDescription(ssTabLabResult.Tab), cboGroups.Text, cboUnits.Text)
            'when creating store code in sTag for seletion afterwards
            sTag = txtCode(ssTabLabResult.Tab)
        Else
            bError = Not moClinicalTests.Update(txtCode(ssTabLabResult.Tab).Text, txtDescription(ssTabLabResult.Tab), cboGroups.Text, cboUnits.Text)
        End If
    Case m_TAB_LAB
        If mbCreate Then
            bError = Not moLabs.Insert(txtCode(ssTabLabResult.Tab), txtDescription(ssTabLabResult.Tab), mvSite)
            'when creating store code in sTag for seletion afterwards
            sTag = txtCode(ssTabLabResult.Tab)
        Else
            bError = Not moLabs.Update(txtCode(ssTabLabResult.Tab), txtDescription(ssTabLabResult.Tab))
        End If
    Case Else
        If mbCreate Then
            bError = Not moCol.Insert(txtCode(ssTabLabResult.Tab), txtDescription(ssTabLabResult.Tab))
            'when creating store code in sTag for seletion afterwards
            sTag = txtCode(ssTabLabResult.Tab)
        Else
            bError = Not moCol.Update(txtCode(ssTabLabResult.Tab), txtDescription(ssTabLabResult.Tab))
        End If
    End Select
    
    If bError Then
        DialogError "An item with the code " & txtCode(ssTabLabResult.Tab) & " already exists"
    Else
        Call ReDisplay(sTag)
        '   TA 13/10/2000: Make sure we repopulate groups combo
        If ssTabLabResult.Tab = m_TAB_CLINICALTESTGROUP Then
            moClinicalTestGroups.PopulateCombo cboGroups
            cboGroups.ListIndex = 0
        End If
    End If
    
Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdSave_Click")
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
Private Sub cmdUndo_Click()
'---------------------------------------------------------------------
'---------------------------------------------------------------------
    Clear
    ReDisplay
End Sub

'---------------------------------------------------------------------
Public Sub Display(Optional vSite As Variant = Null)
'---------------------------------------------------------------------
'---------------------------------------------------------------------
Dim i As Long
Dim nStartTab As Integer

    On Error GoTo ErrHandler
        
    Call HourglassOn
    
    If vSite = "" Then
        ' no site is passed in
        mvSite = Null
    Else
        mvSite = vSite
    End If
    
    lvwList(m_TAB_CLINICALTESTGROUP).Visible = False
    
    Me.Icon = frmMenu.Icon
    
    'initial form size
    mlWidth = Me.Width
    mlHeight = Me.Height

    Set moClinicalTestGroups = New clsClinTestGroups
    
    If AtRemoteSite Or InDM Then
        'in data management so we know site and are editing labs - show lab tab first
        nStartTab = m_TAB_LAB
        ssTabLabResult.Tab = m_TAB_LAB
        lvwList(m_TAB_LAB).Visible = True
    Else
        nStartTab = m_TAB_CLINICALTESTGROUP
        Call moClinicalTestGroups.PopulateListView(lvwList(m_TAB_CLINICALTESTGROUP))
        lvwList(m_TAB_CLINICALTESTGROUP).Visible = True
        Set moCol = moClinicalTestGroups
    End If

    moClinicalTestGroups.PopulateCombo cboGroups
    
    'pick first item
    If cboGroups.ListCount > 0 Then
        cboGroups.ListIndex = 0
    End If


    LoadUnitsCombo
    
    Call EditMode(False)
    
    If lvwList(nStartTab).ListItems.Count > 0 Then
        lvwList(nStartTab).SelectedItem = lvwList(nStartTab).ListItems(1)
        Call ListItemSelect(lvwList(nStartTab).SelectedItem.Tag)
    End If
    
    
    Call TabStopSettings 'rem 18/10/01 set TabStop sequence
    
    'set size, position and window state
    Call SetFormDimensions(Me)
    
    Call HourglassOff
    
    Me.Show vbModal
    
    'unreference collection classes when form unloaded
    'this should free up memory as these should be the last references to the objects
    Set moLabs = Nothing
'only allow schemes in SD/LM
#If SD Then
    Set moSchemes = Nothing
#End If
    Set moClinicalTestGroups = Nothing
    Set moClinicalTests = Nothing
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Display")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'only allow units in SD/LM
#If SD Then
'---------------------------------------------------------------------
Private Sub cmdUnits_Click()
'---------------------------------------------------------------------
' show units form so that they can add more units
'---------------------------------------------------------------------
Dim sCurrentUnit As String
    
    sCurrentUnit = cboUnits.Text
    frmUnitMaintenance.Show vbModal
    'repopulate listbox
    Call LoadUnitsCombo
    'on error reuired just in case sCurrentUnit no longer exists
    On Error Resume Next
    cboUnits.Text = sCurrentUnit
    On Error GoTo 0

End Sub
#End If

'---------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---------------------------------------------------------------------
' are they sure they want to close
'---------------------------------------------------------------------

    If cmdOK.Enabled = False Then
        'currently editing
        If DialogWarning("Closing this window will lose the changes you have just made to the current item", , True) = vbCancel Then
            'cancel form close
            Cancel = 1
        End If
    End If
    'store window dimensions
    Call SaveFormDimensions(Me)
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------

    If Me.Width >= mlWidth Then
        cmdEdit.Left = Me.ScaleWidth - cmdEdit.Width - 300
        cmdCreate.Left = cmdEdit.Left
        cmdDelete.Left = cmdEdit.Left
        cmdUndo.Left = cmdEdit.Left
        cmdSave.Left = cmdEdit.Left
        cmdOK.Left = cmdEdit.Left
        ssTabLabResult.Width = Me.ScaleWidth - 120
        fraList(ssTabLabResult.Tab).Width = ssTabLabResult.Width - 240
        fraDetails(ssTabLabResult.Tab).Width = fraList(ssTabLabResult.Tab).Width
        lvwList(ssTabLabResult.Tab).Width = fraList(ssTabLabResult.Tab).Width - cmdEdit.Width - 360
        txtDescription(ssTabLabResult.Tab).Width = lvwList(ssTabLabResult.Tab).Width
        cmdNormalRange.Left = fraList(ssTabLabResult.Tab).Width - cmdNormalRange.Width - 120
        cmdCTC.Left = cmdNormalRange.Left
    Else
        Me.Width = mlWidth
    End If
    
    If Me.Height >= mlHeight Then
        cmdOK.Top = Me.ScaleHeight - cmdOK.Height - 60
        cmdUndo.Top = cmdOK.Top - 660
        cmdSave.Top = cmdUndo.Top - cmdSave.Height - 90
        
        cmdDelete.Top = cmdUndo.Top - fraDetails(ssTabLabResult.Tab).Height - 90
        cmdEdit.Top = cmdDelete.Top - cmdEdit.Height - 90
        cmdCreate.Top = cmdEdit.Top - cmdCreate.Height - 90
        
        ssTabLabResult.Height = Me.ScaleHeight - cmdOK.Height - 180
        fraList(ssTabLabResult.Tab).Height = ssTabLabResult.Height - fraDetails(ssTabLabResult.Tab).Height - 600
        fraDetails(ssTabLabResult.Tab).Top = fraList(ssTabLabResult.Tab).Top + fraList(ssTabLabResult.Tab).Height + 60
        lvwList(ssTabLabResult.Tab).Height = fraList(ssTabLabResult.Tab).Height - 360
    Else
        Me.Height = mlHeight
    End If

End Sub

'---------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------------------
' clear all objects
'---------------------------------------------------------------------
    Set moLabs = Nothing
'only allow schemes in SD/LM
#If SD Then
    Set moSchemes = Nothing
#End If
    Set moClinicalTestGroups = Nothing
    Set moClinicalTests = Nothing
    Set moCol = Nothing
End Sub


'---------------------------------------------------------------------
Private Sub lvwList_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'---------------------------------------------------------------------

    Call ListViewSort(lvwList(Index), ColumnHeader)

End Sub

'---------------------------------------------------------------------
Private Sub lvwList_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    Call ListItemSelect(Item.Tag)
    
End Sub

'---------------------------------------------------------------------
Private Sub ListItemSelect(sTag As String)
'---------------------------------------------------------------------

'---------------------------------------------------------------------
    With moCol.Item(sTag)
        txtCode(ssTabLabResult.Tab).Text = .Code
        txtDescription(ssTabLabResult.Tab) = .Description
    End With

    Select Case ssTabLabResult.Tab
    Case m_TAB_CLINICALTEST
        With moClinicalTests.Item(sTag)
            Call SetUnitsComboText(.Unit)
            Call moClinicalTestGroups.PopulateCombo(cboGroups)
             SetComboDropdownWidth Me, cboGroups
             cboGroups.Text = .ClinicalTestGroupCode
        End With
    Case m_TAB_LAB
        'disable edit and delete if this labs site does not correspond to current site
        If Not (lvwList(m_TAB_LAB).SelectedItem Is Nothing) Then
        If moLabs.Item(lvwList(m_TAB_LAB).SelectedItem.Tag).SiteText <> RemoveNull(mvSite) Then
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
            Else
                cmdEdit.Enabled = cmdCreate.Enabled
                cmdDelete.Enabled = cmdCreate.Enabled
            End If
        End If
    End Select
    
End Sub
'---------------------------------------------------------------------
Private Sub ssTabLabResult_Click(PreviousTab As Integer)
'---------------------------------------------------------------------
' now fills collections the first time they click a tab
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'set object to have actions performed on it according to tab
    
    
    Select Case ssTabLabResult.Tab
    Case m_TAB_CLINICALTESTGROUP
        Call TabStopSettings 'rem 18/10/01 set Tabstop
        'always loaded
        Set moCol = moClinicalTestGroups
    Case m_TAB_CLINICALTEST
        Call TabStopSettings 'rem 18/10/01 set Tabstop
        If moClinicalTests Is Nothing Then
            '1st time this tab clicked
            Set moClinicalTests = New clsClinicalTests
        End If
        Set moCol = moClinicalTests
    Case m_TAB_LAB
        Call TabStopSettings 'rem 18/10/01 set Tabstop
        If moLabs Is Nothing Then
            '1st time this tab clicked
            Set moLabs = New clsLabs
            'load explicitly
            moLabs.Load
        End If
        Set moCol = moLabs
    Case m_TAB_CTCSCHEME
    Call TabStopSettings 'rem 18/10/01 set Tabstop
'only allow schemes in SD/LM
#If SD Then
        If moSchemes Is Nothing Then
            '1st time this tab clicked
            Set moSchemes = New clsCTCSchemes
        End If
        Set moCol = moSchemes
#End If
    End Select
    
    Form_Resize
    
    Call EditMode(False)
    Call ReDisplay
    
    'ensure visible
    lvwList(ssTabLabResult.Tab).Visible = True
    
    'if called before form is shown
    If lvwList(ssTabLabResult.Tab).Visible Then
        lvwList(ssTabLabResult.Tab).SetFocus
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ssTabLabResult_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub


'--------------------------------------------------------------------------------------------------
Private Sub TabStopSettings()
'--------------------------------------------------------------------------------------------------
'rem 18/10/01
'sets the tabstop to true or false for each tab on the form
'--------------------------------------------------------------------------------------------------
    Select Case ssTabLabResult.Tab
    Case m_TAB_CLINICALTESTGROUP
        lvwList(m_TAB_CLINICALTEST).TabStop = False
        lvwList(m_TAB_LAB).TabStop = False
        lvwList(m_TAB_CTCSCHEME).TabStop = False
        txtCode(m_TAB_CLINICALTEST).TabStop = False
        txtCode(m_TAB_LAB).TabStop = False
        txtCode(m_TAB_CTCSCHEME).TabStop = False
        txtDescription(m_TAB_CLINICALTEST).TabStop = False
        txtDescription(m_TAB_LAB).TabStop = False
        txtDescription(m_TAB_CTCSCHEME).TabStop = False
        cboUnits.TabStop = False
        cmdUnits.TabStop = False
        cboGroups.TabStop = False
        cmdNormalRange.TabStop = False
        cmdCTC.TabStop = False
        
        txtDescription(m_TAB_CLINICALTESTGROUP).TabStop = True
        lvwList(m_TAB_CLINICALTESTGROUP).TabStop = True
        txtCode(m_TAB_CLINICALTESTGROUP).TabStop = True
        
    Case m_TAB_CLINICALTEST
        lvwList(m_TAB_CLINICALTESTGROUP).TabStop = False
        lvwList(m_TAB_LAB).TabStop = False
        lvwList(m_TAB_CTCSCHEME).TabStop = False
        txtCode(m_TAB_CLINICALTESTGROUP).TabStop = False
        txtCode(m_TAB_LAB).TabStop = False
        txtCode(m_TAB_CTCSCHEME).TabStop = False
        txtDescription(m_TAB_CLINICALTESTGROUP).TabStop = False
        txtDescription(m_TAB_LAB).TabStop = False
        txtDescription(m_TAB_CTCSCHEME).TabStop = False
        cmdNormalRange.TabStop = False
        cmdCTC.TabStop = False
        
        txtCode(m_TAB_CLINICALTEST).TabStop = True
        lvwList(m_TAB_CLINICALTEST).TabStop = True
        txtDescription(m_TAB_CLINICALTEST).TabStop = True
        cboUnits.TabStop = True
        cmdUnits.TabStop = True
        cboGroups.TabStop = True
        
    Case m_TAB_LAB
        lvwList(m_TAB_CLINICALTEST).TabStop = False
        lvwList(m_TAB_CLINICALTESTGROUP).TabStop = False
        lvwList(m_TAB_CTCSCHEME).TabStop = False
        txtCode(m_TAB_CLINICALTEST).TabStop = False
        txtCode(m_TAB_CLINICALTESTGROUP).TabStop = False
        txtCode(m_TAB_CTCSCHEME).TabStop = False
        txtDescription(m_TAB_CLINICALTEST).TabStop = False
        txtDescription(m_TAB_CLINICALTESTGROUP).TabStop = False
        txtDescription(m_TAB_CTCSCHEME).TabStop = False
        cboUnits.TabStop = False
        cmdUnits.TabStop = False
        cboGroups.TabStop = False
        cmdCTC.TabStop = False
        
        cmdNormalRange.TabStop = True
        txtDescription(m_TAB_LAB).TabStop = True
        txtCode(m_TAB_LAB).TabStop = True
        lvwList(m_TAB_LAB).TabStop = True
        
        'Check sto see if the listview contains any entries, if not then sets tabstop to false
        If lvwList(m_TAB_LAB).ListItems.Count = 0 Then
            lvwList(m_TAB_LAB).TabStop = False
        Else
            lvwList(m_TAB_LAB).TabStop = True
        End If
        
    Case m_TAB_CTCSCHEME
        lvwList(m_TAB_CLINICALTEST).TabStop = False
        lvwList(m_TAB_LAB).TabStop = False
        lvwList(m_TAB_CLINICALTESTGROUP).TabStop = False
        txtCode(m_TAB_CLINICALTEST).TabStop = False
        txtCode(m_TAB_LAB).TabStop = False
        txtCode(m_TAB_CLINICALTESTGROUP).TabStop = False
        txtDescription(m_TAB_CLINICALTEST).TabStop = False
        txtDescription(m_TAB_LAB).TabStop = False
        txtDescription(m_TAB_CLINICALTESTGROUP).TabStop = False
        cboUnits.TabStop = False
        cmdUnits.TabStop = False
        cboGroups.TabStop = False
        cmdNormalRange.TabStop = False
        
        cmdCTC.TabStop = True
        txtDescription(m_TAB_CTCSCHEME).TabStop = True
        txtCode(m_TAB_CTCSCHEME).TabStop = True
        lvwList(m_TAB_CTCSCHEME).TabStop = True
        
    End Select


End Sub


'---------------------------------------------------------------------
Private Sub EditMode(bEnabled As Boolean)
'---------------------------------------------------------------------
' enable disable lower frame buttons
'---------------------------------------------------------------------
Dim i As Long

    On Error GoTo ErrHandler


    'always ensure this defaults to false when they're are any changes
    mbUpdateDescWithCode = False
    
    txtCode(ssTabLabResult.Tab).Enabled = bEnabled
    txtDescription(ssTabLabResult.Tab).Enabled = bEnabled
    cboUnits.Enabled = bEnabled
    cboGroups.Enabled = bEnabled
    cmdUndo.Enabled = bEnabled
    cmdSave.Enabled = bEnabled
    cmdUnits.Enabled = bEnabled
    
    cmdCreate.Enabled = Not bEnabled
    cmdDelete.Enabled = Not (bEnabled) And Not (lvwList(ssTabLabResult.Tab).SelectedItem Is Nothing)
    cmdEdit.Enabled = Not (bEnabled) And Not (lvwList(ssTabLabResult.Tab).SelectedItem Is Nothing)
    cmdOK.Enabled = Not bEnabled
    ' DPH 15/10/2001 - Allow for Permissions as well for Normal Range and CTC
    cmdNormalRange.Enabled = Not (bEnabled) And Not (lvwList(ssTabLabResult.Tab).SelectedItem Is Nothing) And (goUser.CheckPermission(gsFnMaintainNormalRanges))
    cmdCTC.Enabled = Not (bEnabled) And Not (lvwList(ssTabLabResult.Tab).SelectedItem Is Nothing) And (goUser.CheckPermission(gsFnMaintainCTC))
    
    fraList(ssTabLabResult.Tab).Enabled = Not bEnabled
    
    For i = 0 To ssTabLabResult.Tabs - 1
        'disable other tabs if appropriate
        ssTabLabResult.TabEnabled(i) = Not bEnabled
        'hide listview selection
        lvwList(i).HideSelection = False
    Next
    ssTabLabResult.TabEnabled(ssTabLabResult.Tab) = True

    ValidateData
    
    'disable Clinical Tests tab if there are no groups
    ssTabLabResult.TabEnabled(m_TAB_CLINICALTEST) = (moClinicalTestGroups.Count > 0) And Not (bEnabled)
    
    'disable create, edit and delete buttons if the user does not have appropriate permissions or we are in labs only mode
    With ssTabLabResult
        If ((.Tab = m_TAB_LAB) And NoLabEdit) _
            Or ((.Tab = m_TAB_CTCSCHEME) And NoSchemeEdit) _
            Or ((.Tab = m_TAB_CLINICALTESTGROUP Or .Tab = m_TAB_CLINICALTEST) And (NoTestEdit Or InDM)) Then
                cmdDelete.Enabled = False
                cmdCreate.Enabled = False
                cmdEdit.Enabled = False
        End If
    End With


    If bEnabled Then
        If mbCreate Then
            If txtCode(ssTabLabResult.Tab).Visible Then
                txtCode(ssTabLabResult.Tab).SetFocus
            End If
        Else
            If txtDescription(ssTabLabResult.Tab).Visible Then
                txtDescription(ssTabLabResult.Tab).SetFocus
            End If
        End If
    Else
        If lvwList(ssTabLabResult.Tab).Visible Then
            lvwList(ssTabLabResult.Tab).SetFocus
        End If
    End If
         
    'only allow viewing of CTC Schemes when out of library mangagement
    If InDM Then
        ssTabLabResult.TabEnabled(m_TAB_CTCSCHEME) = False
    End If
     
Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "EditMode")
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
Private Sub ReDisplay(Optional sTag As String = "")
'---------------------------------------------------------------------
'repopulate the listview and select the previously selected item
'---------------------------------------------------------------------
Dim olistItem As MSComctlLib.ListItem

    On Error GoTo ErrHandler
    
    'if tag not nknown and items exist then store tag of selected one
    If lvwList(ssTabLabResult.Tab).ListItems.Count > 0 And sTag = "" Then
        'if tag (code) isn't passed in then get it
        sTag = lvwList(ssTabLabResult.Tab).SelectedItem.Tag
    End If
    
    '(re)fill listview
    Call moCol.PopulateListView(lvwList(ssTabLabResult.Tab))
    Call EditMode(False)
    
    If lvwList(ssTabLabResult.Tab).Visible Then
        lvwList(ssTabLabResult.Tab).SetFocus
    End If
    
    If lvwList(ssTabLabResult.Tab).ListItems.Count > 0 And sTag = "" Then
         sTag = lvwList(ssTabLabResult.Tab).SelectedItem.Tag
    End If
    
    If sTag <> "" And lvwList(ssTabLabResult.Tab).ListItems.Count > 0 Then
        lvwList(ssTabLabResult.Tab).SelectedItem = ListItembyTag(lvwList(ssTabLabResult.Tab), sTag)
        lvwList_ItemClick ssTabLabResult.Tab, ListItembyTag(lvwList(ssTabLabResult.Tab), sTag)
        lvwList(ssTabLabResult.Tab).SelectedItem.EnsureVisible
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Redisplay")
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
Private Sub txtCode_Change(Index As Integer)
'---------------------------------------------------------------------
' enable/disable save button according to input text
'---------------------------------------------------------------------

    If mbUpdateDescWithCode Then
        'still updating desc with code
        txtDescription(Index).Text = txtCode(Index).Text
    End If
    
    ValidateData
    
        
End Sub

'---------------------------------------------------------------------
Private Sub txtDescription_KeyPress(Index As Integer, KeyAscii As Integer)
'---------------------------------------------------------------------
' when key is pressed undo link between code and description
'---------------------------------------------------------------------

    mbUpdateDescWithCode = False
    
End Sub


Private Sub txtDescription_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' when mouse is pressed undo link between code and description
'---------------------------------------------------------------------

    mbUpdateDescWithCode = False
    
End Sub

'---------------------------------------------------------------------
Private Sub txtDescription_Change(Index As Integer)
'---------------------------------------------------------------------
' enable/disable save button according to input text
'---------------------------------------------------------------------

    ValidateData
    
End Sub

'---------------------------------------------------------------------
Private Sub LoadUnitsCombo()
'---------------------------------------------------------------------
Dim rsTemp As adodb.Recordset

    On Error GoTo ErrHandler
    
    cboUnits.Clear
    Set rsTemp = New adodb.Recordset
    cboUnits.AddItem ""
    rsTemp.Open "SELECT DISTINCT Unit FROM Units ORDER BY Unit", MacroADODBConnection
    Do While Not rsTemp.EOF
        cboUnits.AddItem rsTemp.Fields(0).Value
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    SetComboDropdownWidth Me, cboUnits
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtDescription_Change")
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
Private Sub SetUnitsComboText(sUnit As String)
'---------------------------------------------------------------------
' set text of combo assuming the first item is an empty string
' TA 13/10/2000: error handling put in just in case unit has been deleted from units table
'---------------------------------------------------------------------
    If sUnit = "" Then
        cboUnits.ListIndex = 0
    Else
        On Error GoTo ErrHandler
        cboUnits.Text = sUnit
        On Error GoTo 0
    End If
    
    Exit Sub
    
ErrHandler:
    cboUnits.ListIndex = 0
    
End Sub

'---------------------------------------------------------------------
Private Sub ValidateData()
'---------------------------------------------------------------------
' enable/disable save button according to input text
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If cmdCreate.Enabled Then
        'not in edit mode
        txtCode(ssTabLabResult.Tab).BackColor = vbWindowBackground
        txtDescription(ssTabLabResult.Tab).BackColor = vbWindowBackground
        cmdSave.Enabled = False
    Else
        'validate values
        Call ValidTextBox(txtCode(ssTabLabResult.Tab), ttCode)
        Call ValidTextBox(txtDescription(ssTabLabResult.Tab), ttDesc)
        'save allowed if no invalid textboxes AND cmdUndo is enabled (TA 24/10/2000)
        cmdSave.Enabled = (txtCode(ssTabLabResult.Tab).BackColor = vbWindowBackground) _
                        And (txtDescription(ssTabLabResult.Tab).BackColor = vbWindowBackground) _
                        And cmdUndo.Enabled
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ValidateData")
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
Private Sub ControlLock(oCtrl As Control, bLock As Boolean)
'---------------------------------------------------------------------
' change control to edit/ non edit
'---------------------------------------------------------------------
    oCtrl.Locked = bLock
    If bLock Then
        'locked colour background
        oCtrl.BackColor = vbButtonFace
    Else
        'normal background
        oCtrl.BackColor = vbWindowBackground
    End If
End Sub

'---------------------------------------------------------------------
Private Function InDM() As Boolean
'---------------------------------------------------------------------
'is user in data management?
'---------------------------------------------------------------------

    InDM = (App.Title = "MACRO_DM")

End Function

'---------------------------------------------------------------------
Private Function AtRemoteSite() As Boolean
'---------------------------------------------------------------------
' is user at remote site?
'---------------------------------------------------------------------

    AtRemoteSite = (VarType(mvSite) <> vbNull)
    
End Function

'---------------------------------------------------------------------
Private Function NoTestEdit() As Boolean
'---------------------------------------------------------------------
'returns true if user can't edit tests and test groups
'---------------------------------------------------------------------

    NoTestEdit = AtRemoteSite Or Not (goUser.CheckPermission(gsFnMaintainClinicalTests))

End Function

'---------------------------------------------------------------------
Private Function NoSchemeEdit() As Boolean
'---------------------------------------------------------------------
'returns true if user can't edit schemes
'---------------------------------------------------------------------
    NoSchemeEdit = AtRemoteSite Or Not (goUser.CheckPermission(gsFnMaintainCTCSchemes))

End Function

'---------------------------------------------------------------------
Private Function NoLabEdit() As Boolean
'---------------------------------------------------------------------
'returns true if user can't edit labs
'---------------------------------------------------------------------

    NoLabEdit = Not (goUser.CheckPermission(gsFnMaintainLaboratories))

End Function
