VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNormalRanges 
   Caption         =   "Laboratory Normal Ranges"
   ClientHeight    =   7245
   ClientLeft      =   4680
   ClientTop       =   5295
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11355
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   9960
      TabIndex        =   25
      Top             =   6780
      Width           =   1215
   End
   Begin VB.Frame fraNR 
      Caption         =   "Normal Range"
      Height          =   3195
      Left            =   60
      TabIndex        =   26
      Top             =   3480
      Width           =   11235
      Begin VB.Frame Frame2 
         Caption         =   "Clinical Test"
         Height          =   1095
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   10995
         Begin VB.TextBox txtUnit 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   240
            Width           =   1275
         End
         Begin VB.TextBox txtDesc 
            BackColor       =   &H8000000F&
            Height          =   735
            Left            =   5580
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   240
            Width           =   5295
         End
         Begin VB.ComboBox cboTest 
            Height          =   315
            Left            =   780
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   660
            Width           =   1755
         End
         Begin VB.ComboBox cboGroup 
            Height          =   315
            Left            =   780
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Units"
            Height          =   255
            Left            =   2700
            TabIndex        =   46
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Description"
            Height          =   255
            Left            =   4500
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Test"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   660
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Group"
            Height          =   255
            Left            =   180
            TabIndex        =   41
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   9900
         TabIndex        =   24
         Top             =   2700
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   8595
         TabIndex        =   23
         Top             =   2700
         Width           =   1215
      End
      Begin VB.Frame fraValues 
         Caption         =   "Values"
         Height          =   1695
         Left            =   120
         TabIndex        =   34
         Top             =   1380
         Width           =   5535
         Begin VB.CheckBox chkAbsPercent 
            Caption         =   "Percent"
            Height          =   315
            Left            =   4620
            TabIndex        =   15
            Top             =   1260
            Width           =   855
         End
         Begin VB.TextBox txtNormMin 
            Height          =   315
            Left            =   900
            MaxLength       =   12
            TabIndex        =   8
            Top             =   420
            Width           =   1455
         End
         Begin VB.TextBox txtNormMax 
            Height          =   315
            Left            =   2700
            MaxLength       =   12
            TabIndex        =   9
            Top             =   420
            Width           =   1455
         End
         Begin VB.CheckBox chkFeasPercent 
            Caption         =   "Percent"
            Height          =   315
            Left            =   4620
            TabIndex        =   12
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtFeasMin 
            Height          =   315
            Left            =   900
            MaxLength       =   12
            TabIndex        =   10
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtFeasMax 
            Height          =   315
            Left            =   2700
            MaxLength       =   12
            TabIndex        =   11
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtAbsMin 
            Height          =   315
            Left            =   900
            MaxLength       =   12
            TabIndex        =   13
            Top             =   1260
            Width           =   1455
         End
         Begin VB.TextBox txtAbsMax 
            Height          =   315
            Left            =   2700
            MaxLength       =   12
            TabIndex        =   14
            Top             =   1260
            Width           =   1455
         End
         Begin VB.Label lblAbsPercent 
            Caption         =   "%"
            Height          =   255
            Index           =   1
            Left            =   4260
            TabIndex        =   50
            Top             =   1320
            Width           =   195
         End
         Begin VB.Label lblAbsPercent 
            Caption         =   "%"
            Height          =   255
            Index           =   0
            Left            =   2460
            TabIndex        =   49
            Top             =   1320
            Width           =   195
         End
         Begin VB.Label lblFeasPercent 
            Caption         =   "%"
            Height          =   255
            Index           =   1
            Left            =   4260
            TabIndex        =   48
            Top             =   900
            Width           =   195
         End
         Begin VB.Label lblFeasPercent 
            Caption         =   "%"
            Height          =   255
            Index           =   0
            Left            =   2460
            TabIndex        =   47
            Top             =   900
            Width           =   195
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Normal"
            Height          =   255
            Left            =   180
            TabIndex        =   39
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Min"
            Height          =   255
            Left            =   840
            TabIndex        =   38
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Max"
            Height          =   255
            Left            =   2700
            TabIndex        =   37
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Feasible"
            Height          =   255
            Left            =   180
            TabIndex        =   36
            Top             =   900
            Width           =   615
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Absolute"
            Height          =   255
            Left            =   180
            TabIndex        =   35
            Top             =   1320
            Width           =   615
         End
      End
      Begin VB.Frame fraSex 
         Caption         =   "Gender"
         Height          =   1095
         Left            =   5760
         TabIndex        =   33
         Top             =   1380
         Width           =   1455
         Begin VB.OptionButton optMale 
            Caption         =   "Male"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   735
         End
         Begin VB.OptionButton optFemale 
            Caption         =   "Female"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   915
         End
         Begin VB.OptionButton optNone 
            Caption         =   "Unspecified"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1155
         End
      End
      Begin VB.Frame fraAge 
         Caption         =   "Age Range"
         Height          =   1095
         Left            =   7320
         TabIndex        =   30
         Top             =   1380
         Width           =   1215
         Begin VB.TextBox txtAgeMin 
            Height          =   315
            Left            =   600
            MaxLength       =   3
            TabIndex        =   19
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtAgeMax 
            Height          =   315
            Left            =   600
            MaxLength       =   3
            TabIndex        =   20
            Top             =   660
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Min "
            Height          =   255
            Left            =   180
            TabIndex        =   32
            Top             =   300
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Max"
            Height          =   255
            Left            =   180
            TabIndex        =   31
            Top             =   660
            Width           =   315
         End
      End
      Begin VB.Frame fraDates 
         Caption         =   "Effective Dates"
         Height          =   1095
         Left            =   8640
         TabIndex        =   27
         Top             =   1380
         Width           =   2475
         Begin VB.TextBox txtEnd 
            Height          =   315
            Left            =   600
            MaxLength       =   12
            TabIndex        =   22
            Top             =   660
            Width           =   1755
         End
         Begin VB.TextBox txtStart 
            Height          =   315
            Left            =   600
            MaxLength       =   12
            TabIndex        =   21
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "End"
            Height          =   255
            Left            =   60
            TabIndex        =   29
            Top             =   660
            Width           =   435
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Start"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   300
            Width           =   375
         End
      End
   End
   Begin VB.Frame fraRows 
      Caption         =   "Normal Ranges"
      Height          =   3315
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   11235
      Begin VB.CommandButton cmdCreateLike 
         Caption         =   "Create &Like"
         Height          =   375
         Left            =   5940
         TabIndex        =   2
         Top             =   2820
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   9900
         TabIndex        =   5
         Top             =   2820
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   8580
         TabIndex        =   4
         Top             =   2820
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Create"
         Height          =   375
         Left            =   7260
         TabIndex        =   3
         Top             =   2820
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwRows 
         Height          =   2475
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   4366
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
End
Attribute VB_Name = "frmNormalRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       frmNormalRanges.frm
'   Copyright:  InferMed Ltd. 2000-2003. All Rights Reserved
'   Author:     Toby Aldridge, September 2000
'   Purpose:    Form for administering Normal Ranges
'----------------------------------------------------------------------------------------'
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
'   TA 24/10/2000: Button order and position changed
' NCJ 14 Feb 03 - Set MaxTextLength on range fields to be 12 (in form design mode)
'----------------------------------------------------------------------------------------'

Option Explicit


'form minmum height and width
Private mlMinHeight As Long
Private mlMinWidth As Long

Private moNormalRanges As clsNormalRanges
Private moClinicalTests As clsClinicalTests
Private moClinicalTestGroups As clsClinTestGroups

Private molab As clsLab

Private mbAllowEdit As Boolean

Private mlIndex As Long
'add or edit
Private mbCreate As Boolean

'----------------------------------------------------------------------------------------'
Public Sub Display(oLab As clsLab, oClinicalTests As clsClinicalTests, oClinicalTestGroups As clsClinTestGroups, bAllowEdit As Boolean)
'----------------------------------------------------------------------------------------'
' display form
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
        
    Set molab = oLab
        
    Call HourglassOn
    lvwRows.Visible = False
    
    mlMinWidth = Me.Width
    mlMinHeight = Me.Height
    
    mbAllowEdit = bAllowEdit
    
    'set tooltiptext for dates
    txtStart.ToolTipText = DEFAULT_DATE_FORMAT
    txtEnd.ToolTipText = DEFAULT_DATE_FORMAT
    
    Me.Icon = frmMenu.Icon
    Me.Caption = "Laboratory Normal Ranges (" & oLab.Code & ")"
    
    'set up ClinicalTest and groups for combos
    Set moClinicalTests = oClinicalTests
    Set moClinicalTestGroups = oClinicalTestGroups
    
    'set up normal rangess for this laboratory
    Set moNormalRanges = New clsNormalRanges
    Call moNormalRanges.Load(oLab.Code)

    'populate listview
    Call moNormalRanges.PopulateListView(lvwRows)
    'sort by clinical test group
    Call ListViewSort(lvwRows, lvwRows.ColumnHeaders(1))
    
    moClinicalTestGroups.PopulateCombo cboGroup
    'pick first item
    If cboGroup.ListCount > 0 Then
        cboGroup.ListIndex = 0
    End If
    
    EditMode False
    
    'select first item
    If lvwRows.ListItems.Count > 0 Then
        lvwRows_ItemClick lvwRows.ListItems(1)
    End If

    'set size, position and window state
    Call SetFormDimensions(Me)
        
    lvwRows.Visible = True
    Call HourglassOff
    
    Me.Show vbModal
    
    'unreference collection classes when close clicked
    Set moClinicalTests = Nothing
    Set moClinicalTestGroups = Nothing
    Set moNormalRanges = Nothing
    Set molab = Nothing
    
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



Private Sub cboTest_Click()

    If cboTest.ListCount > 0 Then
        txtDesc.Text = moClinicalTests.Item(cboTest.Text).Description
        txtUnit.Text = moClinicalTests.Item(cboTest.Text).Unit
    Else
        txtDesc.Text = ""
        txtUnit.Text = ""
    End If
    ValidateData

End Sub

Private Sub chkAbsPercent_Click()

    ValidateData

End Sub

Private Sub chkFeasPercent_Click()

    ValidateData

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'----------------------------------------------------------------------------------------'
' are they sure they want to close
'----------------------------------------------------------------------------------------'

    If cmdOK.Enabled = False Then
        'currently editing
        If DialogWarning("Closing this window will lose the changes you have just made to the current Normal Range", , True) = vbCancel Then
            'cancel form close
            Cancel = 1
        End If
    End If
    'store window dimensions
    Call SaveFormDimensions(Me)
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cboGroup_Click()
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    If moClinicalTests.PopulateCombo(cboTest, cboGroup.Text) > 0 Then
        cboTest.ListIndex = 0
    Else
        txtDesc.Text = ""
        txtUnit.Text = ""
    End If
    
    ValidateData
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboGroup_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdAdd_Click()
'----------------------------------------------------------------------------------------'
' create  a new normal range
'----------------------------------------------------------------------------------------'

    mbCreate = True
    Clear
    Call EditMode(True)
    lvwRows.HideSelection = True
    
End Sub


'----------------------------------------------------------------------------------------'
Private Sub cmdCreateLike_Click()
'----------------------------------------------------------------------------------------'
' create a new normal range base on the currently slected one
'----------------------------------------------------------------------------------------'

    mbCreate = True
    Call EditMode(True)
    lvwRows.HideSelection = True
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdDelete_Click()
'----------------------------------------------------------------------------------------'
' delete the currently selected item in the listview
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    With moNormalRanges.Item(mlIndex).ClinicalTest
        If DialogQuestion("Do you wish to delete this Normal Range for " & .ClinicalTestGroup.Code & "/" & .Code & "?") = vbYes Then
            Call moNormalRanges.Delete(mlIndex)
            'at this stage the deleted item still exists in the list view
            If lvwRows.ListItems.Count <= 1 Then
                'no further listitems
                mlIndex = 0
            Else
                If lvwRows.SelectedItem.Index = 1 Then
                    'first one deleted - select 2nd
                    mlIndex = lvwRows.ListItems(2).Tag
                Else
                    'not first - select the one before it
                    mlIndex = lvwRows.ListItems(lvwRows.SelectedItem.Index - 1).Tag
                End If
            End If

            txtNormMin.Text = ""
            txtNormMax.Text = ""
            txtFeasMin.Text = ""
            txtFeasMax.Text = ""
            txtAbsMin.Text = ""
            txtAbsMax.Text = ""
            chkFeasPercent.Value = False
            chkAbsPercent.Value = False
            optNone.Value = True
            txtStart.Text = ""
            txtEnd.Text = ""
            txtAgeMin = ""
            txtAgeMax.Text = ""
            
            ReDisplay
        End If
    End With
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdDelete_Click()")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
        
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdEdit_Click()
'----------------------------------------------------------------------------------------'
' edit the currently selected item in listview
'----------------------------------------------------------------------------------------'
    mbCreate = False
    Call EditMode(True)
    
End Sub

Private Sub cmdOK_Click()
'----------------------------------------------------------------------------------------'
' unload form
'----------------------------------------------------------------------------------------'

    Unload Me
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdSave_Click()
'----------------------------------------------------------------------------------------'
' update or insert the data into the database and collection if valid - or display an error
'----------------------------------------------------------------------------------------'
Dim lResult As Long
Dim lNewId As Long
Dim nGender As GenderCode

    On Error GoTo ErrHandler
    
    lNewId = mlIndex

    'determine gender
    If optMale.Value Then
        nGender = GenderCode.gMale
    End If
    If optFemale.Value Then
        nGender = GenderCode.gFemale
    End If
    If optNone.Value Then
        nGender = GenderCode.gNone
    End If
       
    If mbCreate Then
        'insert
        lResult = moNormalRanges.Insert(cboTest.Text, nGender, _
                                txtAgeMin.Text, txtAgeMax.Text, StringtoDate(txtStart.Text, ""), StringtoDate(txtEnd.Text, ""), _
                                txtNormMin.Text, txtNormMax.Text, txtFeasMin.Text, txtFeasMax.Text, _
                                txtAbsMin.Text, txtAbsMax.Text, (chkFeasPercent.Value = vbChecked), (chkAbsPercent.Value = vbChecked), lNewId)
    Else
        'update
        lResult = moNormalRanges.Update(mlIndex, cboTest.Text, nGender, _
                                txtAgeMin.Text, txtAgeMax.Text, StringtoDate(txtStart.Text, ""), StringtoDate(txtEnd.Text, ""), _
                                txtNormMin.Text, txtNormMax.Text, txtFeasMin.Text, txtFeasMax.Text, _
                                txtAbsMin.Text, txtAbsMax.Text, (chkFeasPercent.Value = vbChecked), (chkAbsPercent.Value = vbChecked))
    
    End If
              
        
    If lResult = ValidRangeStatus.vreOK Then
        'update succeeded
        mlIndex = lNewId
        'update lab's changed property to changed
        molab.Changed = Changed.Changed
        ReDisplay

    Else
        'update failed
        If lResult < 0 Then
            'invalid Data
            'TA have to change sign of result 'cos erros returned as negative
            DialogError "Unable to save." & vbCrLf & GetValidRangeStatusText(Abs(lResult))
        Else
            'conflicted with exisiting
            HighlightListItembyTag lvwRows, Format(lResult)
            DialogError "You have already defined a normal range for this test, gender, age and date"
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
            End
    End Select
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function ReDisplay()
'----------------------------------------------------------------------------------------'

    moNormalRanges.PopulateListView lvwRows
    EditMode False
    If lvwRows.ListItems.Count > 0 Then
    Set lvwRows.SelectedItem = ListItembyTag(lvwRows, Format(mlIndex))
        lvwRows_ItemClick ListItembyTag(lvwRows, Format(mlIndex))
    End If

End Function
'----------------------------------------------------------------------------------------'
Private Sub cmdUndo_Click()
'----------------------------------------------------------------------------------------'
' reset and return to top frame
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    Clear
    'renable top frame
    EditMode False
    'undo all bold
    HighlightListItembyTag lvwRows, -1
    'select previously selected listitem
    Set lvwRows.SelectedItem = ListItembyTag(lvwRows, Format(mlIndex))
    If lvwRows.ListItems.Count > 0 Then
        lvwRows_ItemClick ListItembyTag(lvwRows, Format(mlIndex))
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdUndo_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub


'----------------------------------------------------------------------------------------'
Private Sub Form_Resize()
'----------------------------------------------------------------------------------------'
' handle form resizing
'----------------------------------------------------------------------------------------'

    If Me.Width >= mlMinWidth Then
        fraRows.Width = Me.ScaleWidth - 120
        lvwRows.Width = fraRows.Width - 240
        cmdDelete.Left = fraRows.Width - cmdDelete.Width - 120
        cmdEdit.Left = cmdDelete.Left - cmdEdit.Width - 90
        cmdAdd.Left = cmdEdit.Left - cmdAdd.Width - 90
        cmdCreateLike.Left = cmdAdd.Left - cmdCreateLike.Width - 90
        cmdOK.Left = Me.ScaleWidth - cmdOK.Width - 180
    Else
        Me.Width = mlMinWidth
    End If
    
    If Me.Height >= mlMinHeight Then
        fraRows.Height = Me.ScaleHeight - fraNR.Height - cmdOK.Height - 240
        lvwRows.Height = fraRows.Height - cmdEdit.Height - 480
        cmdEdit.Top = fraRows.Height - cmdEdit.Height - 120
        cmdAdd.Top = cmdEdit.Top
        cmdCreateLike.Top = cmdEdit.Top
        cmdDelete.Top = cmdEdit.Top
        fraNR.Top = fraRows.Top + fraRows.Height + 60
        cmdOK.Top = fraNR.Top + fraNR.Height + 90
    Else
        Me.Height = mlMinHeight
    End If
    
        
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub lvwRows_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'----------------------------------------------------------------------------------------'

    Call ListViewSort(lvwRows, ColumnHeader)
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub lvwRows_ItemClick(ByVal Item As MSComctlLib.ListItem)
'----------------------------------------------------------------------------------------'
'select listitem
    'store currently selected index
    mlIndex = Val(Item.Tag)
    Call NormalRangeSelect(mlIndex)
End Sub

'----------------------------------------------------------------------------------------'
Private Sub NormalRangeSelect(lIndex As Long)
'----------------------------------------------------------------------------------------'
' listitem selected - fill the text boxes with that item's data
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    With moNormalRanges.Item(lIndex)
        cboGroup.Text = moClinicalTests.Item(.ClinicalTestCode).ClinicalTestGroupCode
        moClinicalTests.PopulateCombo cboTest, moClinicalTests.Item(.ClinicalTestCode).ClinicalTestGroupCode
        cboTest.Text = .ClinicalTestCode
        txtDesc.Text = moClinicalTests.Item(.ClinicalTestCode).Description
        txtUnit.Text = moClinicalTests.Item(.ClinicalTestCode).Unit
        txtNormMin.Text = .NormalMinText
        txtNormMax.Text = .NormalMaxText
        txtFeasMin.Text = .FeasibleMinText
        txtFeasMax.Text = .FeasibleMaxText
        txtAbsMin.Text = .AbsoluteMinText
        txtAbsMax.Text = .AbsolutemaxText
        chkFeasPercent.Value = Switch(.FeasiblePercent, vbChecked, Not .FeasiblePercent, vbUnchecked)
        chkAbsPercent.Value = Switch(.AbsolutePercent, vbChecked, Not .AbsolutePercent, vbUnchecked)
        Select Case .GenderCode
        Case GenderCode.gMale: optMale.Value = True
        Case GenderCode.gFemale: optFemale.Value = True
        Case Else: optNone.Value = True
        End Select
        txtStart.Text = .EffectiveStartText
        txtEnd.Text = .EffectiveEndText
        txtAgeMin = .AgeMinText
        txtAgeMax.Text = .AgeMaxText

    End With
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "NormalRangeSelect")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub

'----------------------------------------------------------------------------------------'
Private Sub EditMode(bEnabled As Boolean)
'----------------------------------------------------------------------------------------'
' enable and disble buttons according to mode
'----------------------------------------------------------------------------------------'
Dim ctrl As Control

    On Error GoTo ErrHandler
    
    lvwRows.HideSelection = False
    
    For Each ctrl In Me
        If (TypeOf ctrl Is TextBox) Or (TypeOf ctrl Is CheckBox) Or (TypeOf ctrl Is OptionButton) Or (TypeOf ctrl Is ComboBox) Then
            ctrl.Enabled = bEnabled
        End If
    Next

    cmdSave.Enabled = bEnabled
    cmdUndo.Enabled = bEnabled
    fraRows.Enabled = Not bEnabled
    cmdOK.Enabled = Not bEnabled
    cmdAdd.Enabled = Not bEnabled And mbAllowEdit
    cmdCreateLike.Enabled = Not (bEnabled) And Not (lvwRows.SelectedItem Is Nothing) And mbAllowEdit
    cmdDelete.Enabled = Not (bEnabled) And Not (lvwRows.SelectedItem Is Nothing) And mbAllowEdit
    cmdEdit.Enabled = Not (bEnabled) And Not (lvwRows.SelectedItem Is Nothing) And mbAllowEdit
    cmdOK.Enabled = Not bEnabled
    
    ValidateData
    
    'disable create, create like,   edit and delete buttons if the user does not have appropriate permissions
    If Not (goUser.CheckPermission(gsFnMaintainNormalRanges)) Then
        cmdDelete.Enabled = False
        cmdAdd.Enabled = False
        cmdCreateLike.Enabled = False
        cmdEdit.Enabled = False
    End If

    'set lower frame caption according to mode
    If bEnabled Then
        If mbCreate Then
            fraNR.Caption = "Normal Range (Create)"
        Else
            fraNR.Caption = "Normal Range (Edit)"
        End If
    Else
        fraNR.Caption = "Normal Range"
    End If

    'setfocus errors if form is not yet displayed
    On Error Resume Next
    If bEnabled Then
        cboGroup.SetFocus
    Else
        lvwRows.SetFocus
    End If
    On Error GoTo ErrHandler
    
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

'----------------------------------------------------------------------------------------'
Private Sub Clear()
'----------------------------------------------------------------------------------------'
'clear all the text boxes
'----------------------------------------------------------------------------------------'

    txtNormMin.Text = ""
    txtNormMax.Text = ""
    txtFeasMin.Text = ""
    txtFeasMax.Text = ""
    txtAbsMin.Text = ""
    txtAbsMax.Text = ""
    chkFeasPercent.Value = vbUnchecked
    chkAbsPercent.Value = vbUnchecked
    optNone.Value = True
    txtStart.Text = ""
    txtEnd.Text = ""
    txtAgeMin = ""
    txtAgeMax.Text = ""
    
End Sub

'---------------------------------------------------------------------
Private Sub ValidateData()
'---------------------------------------------------------------------
' enable/disable save button according to input text
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If cmdAdd.Enabled Or Not (mbAllowEdit) Then
        'not in edit mode
        txtNormMin.BackColor = vbWindowBackground
        txtNormMax.BackColor = vbWindowBackground
        txtFeasMin.BackColor = vbWindowBackground
        txtFeasMax.BackColor = vbWindowBackground
        txtAbsMin.BackColor = vbWindowBackground
        txtAbsMax.BackColor = vbWindowBackground
        txtAgeMin.BackColor = vbWindowBackground
        txtAgeMax.BackColor = vbWindowBackground
        cmdSave.Enabled = False
    Else
    
        'check lab
        If cboTest.ListIndex = -1 Then
            cboTest.BackColor = vbYellow
        Else
            cboTest.BackColor = vbWindowBackground
        End If
    
        'validate text box values
        Call ValidTextBox(txtNormMin, ttDouble)
        Call ValidTextBox(txtNormMax, ttDouble)
        Call ValidTextBox(txtFeasMin, ttDouble)
        Call ValidTextBox(txtFeasMax, ttDouble)
        Call ValidTextBox(txtAbsMin, ttDouble)
        Call ValidTextBox(txtAbsMax, ttDouble)
        Call ValidTextBox(txtAgeMin, ttAge)
        Call ValidTextBox(txtAgeMax, ttAge)
        
        'check dates
        If StringtoDate(txtStart.Text, "") <> -1 Then
            'valid
            txtStart.BackColor = vbWindowBackground
        Else
            txtStart.BackColor = vbYellow
        End If
        If StringtoDate(txtEnd.Text, "") <> -1 Then
            'valid
            txtEnd.BackColor = vbWindowBackground
        Else
            txtEnd.BackColor = vbYellow
        End If
        
        If txtNormMin.Text = "" And txtNormMax.Text = "" Then
            'normal range undefined - disallow
            txtNormMin.BackColor = vbYellow
            txtNormMax.BackColor = vbYellow
        End If
        
        'check valid ranges
        Call ValidRangeTextBoxes(txtNormMin, txtNormMax)
        If chkFeasPercent.Value = vbUnchecked Then
            'absolute value
            Call ValidRangeTextBoxes(txtFeasMin, txtFeasMax)
        End If
        If chkAbsPercent.Value = vbUnchecked Then
            'absolute value
            Call ValidRangeTextBoxes(txtAbsMin, txtAbsMax)
        End If
        Call ValidRangeTextBoxes(txtAgeMin, txtAgeMax)
        
        'check effective dates valid
        If Not ValidRange(StringtoDateVariant(txtStart.Text), StringtoDateVariant(txtEnd.Text)) Then
            'invalid
            txtStart.BackColor = vbYellow
            txtEnd.BackColor = vbYellow
        End If
        
        'do not allow negative value if feasible is a percetnage
        If chkFeasPercent.Value = vbChecked Then
            If Val(txtFeasMin) < 0 Then
                txtFeasMin.BackColor = vbYellow
            End If
            If Val(txtFeasMax) < 0 Then
                txtFeasMax.BackColor = vbYellow
            End If
        End If
        'do not allow negative value if feasible is a percetnage
        If chkAbsPercent.Value = vbChecked Then
            If Val(txtAbsMin) < 0 Then
                txtAbsMin.BackColor = vbYellow
            End If
            If Val(txtAbsMax) < 0 Then
                txtAbsMax.BackColor = vbYellow
            End If
        End If
        
        'save allowed if no invalid textboxes AND cmdUndo is enabled (TA 24/10/2000)
        cmdSave.Enabled = txtNormMin.BackColor = vbWindowBackground And txtNormMax.BackColor = vbWindowBackground _
                        And txtFeasMin.BackColor = vbWindowBackground And txtFeasMax.BackColor = vbWindowBackground _
                        And txtAbsMin.BackColor = vbWindowBackground And txtAbsMax.BackColor = vbWindowBackground _
                        And txtAgeMin.BackColor = vbWindowBackground And txtAgeMax.BackColor = vbWindowBackground _
                        And cboTest.BackColor = vbWindowBackground _
                        And cmdUndo.Enabled
    End If
           
    'turn on and off percent labels
    lblFeasPercent(0).Visible = chkFeasPercent.Value = vbChecked
    lblFeasPercent(1).Visible = chkFeasPercent.Value = vbChecked
    lblAbsPercent(0).Visible = chkAbsPercent.Value = vbChecked
    lblAbsPercent(1).Visible = chkAbsPercent.Value = vbChecked
    
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
Private Function StringtoDate(ByVal sDate As String, sFormattedDate As String) As Double
'---------------------------------------------------------------------
'return a double from a date string
'input:
'   sDate - string
'output:
'   -1 - invalid date
'   0 - empty string
'   positive double - date as double
'   sArezzoDate - formatted version of string ("" if invalid)
'---------------------------------------------------------------------
Dim sArezzoDate As String

    On Error GoTo ErrHandler

    If Trim(sDate) = "" Then
        'not entered - unspoecified
        StringtoDate = UNSPECIFIED_DATE
    Else
        'Read date using current default date format
        sFormattedDate = ReadValidDate(sDate, DEFAULT_DATE_FORMAT, sArezzoDate)
        If sFormattedDate = "" Then
            'invalid Date
            StringtoDate = -1
        Else
            ' Date was accepted OK
            ' Convert to Double value
            StringtoDate = ConvertDateFromArezzo(sArezzoDate)
        End If
    End If
    
Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "StringtoDate")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   

End Function


Public Function StringtoDateVariant(sString As String) As Variant
'----------------------------------------------------------------------------------------'
' convert a string into a date (double) variant or null if empty string or invalid date
'----------------------------------------------------------------------------------------'
Dim sArezzoDate As String   'dummy string passed through to CLM
Dim sDate As String
    
    sDate = ReadValidDate(Trim(sString), DEFAULT_DATE_FORMAT, sArezzoDate)

    If sDate = "" Then
        StringtoDateVariant = Null
    Else
        StringtoDateVariant = sDate
    End If
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub txtEnd_LostFocus()
'----------------------------------------------------------------------------------------'
' change the textbox to proper format
'----------------------------------------------------------------------------------------'
Dim sDate As String
    
    Call StringtoDate(txtEnd.Text, sDate)
    txtEnd.Text = sDate
    ValidateData
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtStart_LostFocus()
'----------------------------------------------------------------------------------------'
' change the textbox to proper format
'----------------------------------------------------------------------------------------'
Dim sDate As String
    Call StringtoDate(txtStart.Text, sDate)
    txtStart.Text = sDate
    ValidateData
    
End Sub

Private Sub txtAbsMax_Change()

    ValidateData

End Sub

Private Sub txtAbsMin_Change()

    ValidateData

End Sub

Private Sub txtAgeMax_Change()

    ValidateData

End Sub

Private Sub txtAgeMin_Change()

    ValidateData

End Sub

Private Sub txtEnd_Change()

    ValidateData

End Sub

Private Sub txtFeasMax_Change()

    ValidateData

End Sub

Private Sub txtFeasMin_Change()

    ValidateData

End Sub

Private Sub txtNormMax_Change()

    ValidateData

End Sub

Private Sub txtNormMin_Change()

    ValidateData

End Sub

Private Sub txtStart_Change()

    ValidateData

End Sub
