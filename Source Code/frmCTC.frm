VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCTC 
   Caption         =   "Common Toxicity Criteria"
   ClientHeight    =   7260
   ClientLeft      =   4230
   ClientTop       =   2175
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8985
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   6780
      Width           =   1215
   End
   Begin VB.Frame fraCTC 
      Caption         =   "Common Toxicity Criterion"
      Height          =   3435
      Left            =   60
      TabIndex        =   14
      Top             =   3240
      Width           =   8835
      Begin VB.Frame Frame2 
         Caption         =   "Test"
         Height          =   1095
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   8595
         Begin VB.TextBox txtDesc 
            BackColor       =   &H8000000F&
            Height          =   735
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   240
            Width           =   4695
         End
         Begin VB.ComboBox cboTest 
            Height          =   315
            Left            =   780
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   660
            Width           =   1755
         End
         Begin VB.ComboBox cboGroup 
            Height          =   315
            Left            =   780
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Description"
            Height          =   255
            Left            =   2580
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Test"
            Height          =   255
            Left            =   180
            TabIndex        =   20
            Top             =   660
            Width           =   435
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Group"
            Height          =   255
            Left            =   180
            TabIndex        =   19
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   7500
         TabIndex        =   11
         Top             =   2940
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   7500
         TabIndex        =   10
         Top             =   2475
         Width           =   1215
      End
      Begin VB.Frame fraValues 
         Caption         =   "Values"
         Height          =   1935
         Left            =   120
         TabIndex        =   15
         Top             =   1380
         Width           =   7275
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   315
            Left            =   2640
            TabIndex        =   32
            Top             =   1080
            Width           =   4575
            Begin VB.OptionButton optMax 
               Caption         =   "Upper Limit Normal"
               Height          =   315
               Index           =   2
               Left            =   2880
               TabIndex        =   35
               Top             =   0
               Width           =   1695
            End
            Begin VB.OptionButton optMax 
               Caption         =   "Lower Limit Normal"
               Height          =   315
               Index           =   1
               Left            =   1080
               TabIndex        =   34
               Top             =   0
               Width           =   1695
            End
            Begin VB.OptionButton optMax 
               Caption         =   "Absolute"
               Height          =   315
               Index           =   0
               Left            =   0
               TabIndex        =   33
               Top             =   0
               Width           =   915
            End
         End
         Begin VB.Frame fraMin 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   315
            Left            =   2640
            TabIndex        =   28
            Top             =   660
            Width           =   4575
            Begin VB.OptionButton optMin 
               Caption         =   "Absolute"
               Height          =   315
               Index           =   0
               Left            =   0
               TabIndex        =   31
               Top             =   0
               Width           =   915
            End
            Begin VB.OptionButton optMin 
               Caption         =   "Lower Limit Normal"
               Height          =   315
               Index           =   1
               Left            =   1080
               TabIndex        =   30
               Top             =   0
               Width           =   1695
            End
            Begin VB.OptionButton optMin 
               Caption         =   "Upper Limit Normal"
               Height          =   315
               Index           =   2
               Left            =   2880
               TabIndex        =   29
               Top             =   0
               Width           =   1695
            End
         End
         Begin VB.TextBox txtUnit 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   240
            Width           =   1755
         End
         Begin VB.TextBox txtExpr 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1500
            Width           =   4215
         End
         Begin VB.ComboBox cboGrade 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtMin 
            Height          =   315
            Left            =   960
            TabIndex        =   8
            Top             =   660
            Width           =   1455
         End
         Begin VB.TextBox txtMax 
            Height          =   315
            Left            =   960
            TabIndex        =   9
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Units"
            Height          =   255
            Left            =   2580
            TabIndex        =   26
            Top             =   300
            Width           =   435
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Grade"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Expression"
            Height          =   255
            Left            =   60
            TabIndex        =   23
            Top             =   1560
            Width           =   795
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Min"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Max"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   735
         End
      End
   End
   Begin VB.Frame fraRows 
      Caption         =   "Common Toxicity Criteria"
      Height          =   3135
      Left            =   60
      TabIndex        =   13
      Top             =   60
      Width           =   8835
      Begin VB.CommandButton cmdCreateLike 
         Caption         =   "Create &Like"
         Height          =   375
         Left            =   3540
         TabIndex        =   1
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   7500
         TabIndex        =   5
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   6180
         TabIndex        =   3
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Create"
         Height          =   375
         Left            =   4860
         TabIndex        =   2
         Top             =   2640
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwRows 
         Height          =   2295
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   4048
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
Attribute VB_Name = "frmCTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       frmNormalRanges.frm
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, September 2000
'   Purpose:    Form for administering Normal Ranges
'----------------------------------------------------------------------------------------'
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
'   TA 24/10/2000: Button order and position changed
'   TA 25/10/2000: Alterations to valid CTC checking (Sub ValidateCriterion)

Option Explicit


'minimum height and width
Private mlMinWidth As Long
Private mlMinHeight As Long

Private moCTC As clsCTCriteria
Private moClinicalTests As clsClinicalTests
Private moClinicalTestGroups As clsClinTestGroups

Private mlIndex As Long
'add or edit
Private mbCreate As Boolean

'----------------------------------------------------------------------------------------'
Public Sub Display(oScheme As clsCTCScheme, oClinicalTests As clsClinicalTests, oClinicalTestGroups As clsClinTestGroups)
'----------------------------------------------------------------------------------------'
' display form
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
        
    Call HourglassOn
    lvwRows.Visible = False
    
    'get min width from desgned size of form
    mlMinWidth = Me.Width
    mlMinHeight = Me.Height
    
    Me.Icon = frmMenu.Icon
    
    Me.Caption = "Common Toxicity Criteria (" & oScheme.Code & ")"
    
    'set up ClinicalTest and groups for combos
    Set moClinicalTests = oClinicalTests
    Set moClinicalTestGroups = oClinicalTestGroups
    
    'set up normal rangess for this laboratory
    Set moCTC = New clsCTCriteria
    Call moCTC.Load(oScheme.Code)

    'populate listview
    Call moCTC.PopulateListView(lvwRows)
    
    'TA 03/04/2003: no longer needed - inital sort done in sql
'    'sort by group
'    Call ListViewSort(lvwRows, lvwRows.ColumnHeaders(1))
    
    moClinicalTestGroups.PopulateCombo cboGroup
    'pick first item
    If cboGroup.ListCount > 0 Then
        cboGroup.ListIndex = 0
    End If

    FillGradeCombo
    
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
    
    'unreference collection classes when form unloaded
    Set moClinicalTests = Nothing
    Set moClinicalTestGroups = Nothing
    Set moCTC = Nothing
    
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

'----------------------------------------------------------------------------------------'
Private Sub cboGroup_Click()
'----------------------------------------------------------------------------------------'

    If moClinicalTests.PopulateCombo(cboTest, cboGroup.Text) > 0 Then
        cboTest.ListIndex = 0
    Else
        txtDesc.Text = ""
        txtUnit.Text = ""
    End If
    ValidateCriterion
    
End Sub



Private Sub cboTest_Click()

    If cboTest.ListCount > 0 Then
        txtDesc.Text = moClinicalTests.Item(cboTest.Text).Description
        txtUnit.Text = moClinicalTests.Item(cboTest.Text).Unit
    Else
        txtDesc.Text = ""
        txtUnit.Text = ""
    End If
    Call ValidateCriterion
    Call Evaluate

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
    
    With moCTC.Item(mlIndex).ClinicalTest
        If DialogQuestion("Do you wish to delete this Criterion for " & .ClinicalTestGroup.Code & "/" & .Code & "?") = vbYes Then

            Call moCTC.Delete(mlIndex)
            'at this stage the deleted item still exists in the list view
            If lvwRows.ListItems.Count <= 1 Then
                'no further listitems
                mlIndex = 0
            Else
                If lvwRows.SelectedItem.Index = 1 Then
                    'first one deleted - selcet 2nd
                    mlIndex = lvwRows.ListItems(2).Tag
                Else
                    'not first - select the one before it
                    mlIndex = lvwRows.ListItems(lvwRows.SelectedItem.Index - 1).Tag
                End If
            End If
                
            txtMin.Text = ""
            txtMax = ""
            optMin(NRFactor.nrfabsolute).Value = True
            optMax(NRFactor.nrfabsolute).Value = True
                
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
Dim sMessage As String
Dim bSaved As Boolean
Dim lId As Long
    If mbCreate Then
        'insert
        lId = moCTC.Insert(cboTest.Text, cboGrade.Text, txtMin.Text, txtMax.Text, MinType, MaxType, sMessage)
        If lId = -1 Then
            'conflicting CTC
            bSaved = False
        Else
            mlIndex = lId
            bSaved = True
        End If
    Else
        'update
        bSaved = moCTC.Update(mlIndex, cboTest.Text, cboGrade.Text, txtMin.Text, txtMax.Text, MinType, MaxType, sMessage)
                
    End If

    If bSaved Then
        ReDisplay
    Else
        DialogError sMessage
        
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function ReDisplay()
'----------------------------------------------------------------------------------------'

        moCTC.PopulateListView lvwRows

        EditMode False
        Set lvwRows.SelectedItem = ListItembyTag(lvwRows, Format(mlIndex))
        If lvwRows.ListItems.Count > 0 Then
            lvwRows_ItemClick ListItembyTag(lvwRows, Format(mlIndex))
        End If
        
End Function

'----------------------------------------------------------------------------------------'
Private Sub cmdUndo_Click()
'----------------------------------------------------------------------------------------'
' reset and return to top frame
'----------------------------------------------------------------------------------------'
    'renable top frame
    EditMode False
    'undo all bold
    HighlightListItembyTag lvwRows, -1
    'select previously selected listitem
    Set lvwRows.SelectedItem = ListItembyTag(lvwRows, Format(mlIndex))
    If lvwRows.ListItems.Count > 0 Then
        lvwRows_ItemClick ListItembyTag(lvwRows, Format(mlIndex))
    End If
    
End Sub


Private Sub Form_Resize()
' handle form resizing
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
        fraRows.Height = Me.ScaleHeight - fraCTC.Height - cmdOK.Height - 240
        lvwRows.Height = fraRows.Height - cmdEdit.Height - 480
        cmdEdit.Top = fraRows.Height - cmdEdit.Height - 120
        cmdAdd.Top = cmdEdit.Top
        cmdCreateLike.Top = cmdEdit.Top
        cmdDelete.Top = cmdEdit.Top
        fraCTC.Top = fraRows.Top + fraRows.Height + 60
        cmdOK.Top = fraCTC.Top + fraCTC.Height + 90
    Else
        Me.Height = mlMinHeight
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub lvwRows_ItemClick(ByVal Item As MSComctlLib.ListItem)
'----------------------------------------------------------------------------------------'
'select listitem
    'store currently selected index
    mlIndex = Val(Item.Tag)
    Call CTCRangeSelect(mlIndex)
End Sub

'----------------------------------------------------------------------------------------'
Private Sub lvwRows_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'----------------------------------------------------------------------------------------'

    Call ListViewSort(lvwRows, ColumnHeader)
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub CTCRangeSelect(lIndex As Long)
'----------------------------------------------------------------------------------------'
' listitem selected - fill the text boxes with that item's data
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler

    With moCTC.Item(lIndex)
        moClinicalTestGroups.PopulateCombo cboGroup
        cboGroup.Text = moClinicalTests.Item(.ClinicalTestCode).ClinicalTestGroupCode
        moClinicalTests.PopulateCombo cboTest, moClinicalTests.Item(.ClinicalTestCode).ClinicalTestGroupCode
         cboTest = .ClinicalTestCode
        txtDesc.Text = moClinicalTests.Item(.ClinicalTestCode).Description
        txtUnit.Text = moClinicalTests.Item(.ClinicalTestCode).Unit
        ListCtrl_Pick cboGrade, .Grade
        txtMin.Text = .MinText
        txtMax = .MaxText
        optMin(.MinType).Value = True
        optMax(.MaxType).Value = True
        
    End With
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "CTCRangeSelect")
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
    cmdAdd.Enabled = Not bEnabled
    cmdCreateLike.Enabled = Not (bEnabled) And Not (lvwRows.SelectedItem Is Nothing)
    cmdDelete.Enabled = Not (bEnabled) And Not (lvwRows.SelectedItem Is Nothing)
    cmdEdit.Enabled = Not (bEnabled) And Not (lvwRows.SelectedItem Is Nothing)
    cmdOK.Enabled = Not bEnabled
    
    ValidateCriterion
    
    'disable create, create like, edit and delete buttons if the user does not have appropriate permissions
    If Not (goUser.CheckPermission(gsFnMaintainCTC)) Then
        cmdDelete.Enabled = False
        cmdAdd.Enabled = False
        cmdCreateLike.Enabled = False
        cmdEdit.Enabled = False
    End If
    
    'set lower frame caption according to mode
    If bEnabled Then
        If mbCreate Then
            fraCTC.Caption = "Common Toxicity Criterion (Create)"
        Else
            fraCTC.Caption = "Common Toxicity Criterion (Edit)"
        End If
    Else
        fraCTC.Caption = "Common Toxicity Criterion"
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

    cboGrade.ListIndex = 0
    txtMin.Text = ""
    txtMax = ""
    optMin(NRFactor.nrfabsolute).Value = True
    optMax(NRFactor.nrfabsolute).Value = True
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'----------------------------------------------------------------------------------------'
' are they sure they want to close
'----------------------------------------------------------------------------------------'

    If cmdOK.Enabled = False Then
        'currently editing
        If DialogWarning("Closing this window will lose the changes you have just made to the current Criterion", , True) = vbCancel Then
            'cancel form close
            Cancel = 1
        End If
    End If

    'store window dimensions
    Call SaveFormDimensions(Me)
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function MinType() As NRFactor
'----------------------------------------------------------------------------------------'
'return min type from option buttons
'----------------------------------------------------------------------------------------'

    If optMin(NRFactor.nrfabsolute).Value = True Then
        MinType = NRFactor.nrfabsolute
        Exit Function
    End If
    
    If optMin(NRFactor.nrfLower).Value = True Then
        MinType = NRFactor.nrfLower
        Exit Function
    End If
    
    If optMin(NRFactor.nrfUpper).Value = True Then
        MinType = NRFactor.nrfUpper
        Exit Function
    End If
    
End Function

'----------------------------------------------------------------------------------------'
Private Function MaxType() As NRFactor
'----------------------------------------------------------------------------------------'
'return max type from option buttons
'----------------------------------------------------------------------------------------'

    If optMax(NRFactor.nrfabsolute).Value = True Then
        MaxType = NRFactor.nrfabsolute
        Exit Function
    End If
    
    If optMax(NRFactor.nrfLower).Value = True Then
        MaxType = NRFactor.nrfLower
        Exit Function
    End If
    
    If optMax(NRFactor.nrfUpper).Value = True Then
        MaxType = NRFactor.nrfUpper
        Exit Function
    End If
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub Evaluate()
'----------------------------------------------------------------------------------------'
'evalutate the string the expression according to controls
'----------------------------------------------------------------------------------------'
       
    txtExpr.Text = CTCExpr(txtMin.Text, txtMax.Text, MinType, MaxType, txtUnit.Text)

End Sub

'----------------------------------------------------------------------------------------'
Private Sub optMax_Click(Index As Integer)
'----------------------------------------------------------------------------------------'

    'if blank and LLN or ULN clicked the set to "1"
    If txtMax.Text = "" And (optMax(1).Value Or optMax(2).Value) Then
        txtMax.Text = "1"
    End If

    Evaluate
    ValidateCriterion

End Sub

'----------------------------------------------------------------------------------------'
Private Sub optMin_Click(Index As Integer)
'----------------------------------------------------------------------------------------'

    'if blank and LLN or ULN clicked the set to "1"
    If txtMin.Text = "" And (optMin(1).Value Or optMin(2).Value) Then
        txtMin.Text = "1"
    End If
    Evaluate
    ValidateCriterion

End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtMax_Change()
'----------------------------------------------------------------------------------------'

    Evaluate
    ValidateCriterion

End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtMin_Change()
'----------------------------------------------------------------------------------------'

    Evaluate
    ValidateCriterion
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub FillGradeCombo()
'----------------------------------------------------------------------------------------'

    With cboGrade
        .AddItem "1"
        .ItemData(.NewIndex) = 1
        .AddItem "2"
        .ItemData(.NewIndex) = 2
        .AddItem "3"
        .ItemData(.NewIndex) = 3
        .AddItem "4"
        .ItemData(.NewIndex) = 4
    End With
    
End Sub

'---------------------------------------------------------------------
Private Sub ValidateCriterion()
'---------------------------------------------------------------------
' enable/disable save button according to input text
'---------------------------------------------------------------------
 Dim bAbs As Boolean
 
    On Error GoTo ErrHandler
    
    If cmdAdd.Enabled Then
        'not in edit mode
        txtMin.BackColor = vbWindowBackground
        txtMax.BackColor = vbWindowBackground
        cmdSave.Enabled = False
    Else
        'check lab
        If cboTest.ListIndex = -1 Then
            cboTest.BackColor = vbYellow
        Else
            cboTest.BackColor = vbWindowBackground
        End If
        
        'validate values
        Call ValidTextBox(txtMin, ttDouble)
        Call ValidTextBox(txtMax, ttDouble)
        
        
        If (txtMin.Text = "" And txtMax.Text = "") Or ((Not (ValidRange(StringtoNumberVariant(txtMin.Text), StringtoNumberVariant(txtMax.Text)))) _
                                And optMin(0).Value And optMax(0).Value) Then
            'both blank or invalid range while both values absolute
            txtMin.BackColor = vbYellow
            txtMax.BackColor = vbYellow
        End If

        'TA 25/10/2000: only allow min to be 0 or blank when an absolute value
        If Val(txtMin.Text) = 0 And Not optMin(0).Value Then
            txtMin.BackColor = vbYellow
        End If
        'TA 25/10/2000: only allow min to be 0 or blank when an absolute value
        If Val(txtMax.Text) = 0 And Not optMax(0).Value Then
            txtMax.BackColor = vbYellow
        End If

        'save allowed if no invalid textboxes AND cmdUndo is enabled (TA 24/10/2000)
        cmdSave.Enabled = (txtMin.BackColor = vbWindowBackground) And (txtMax.BackColor = vbWindowBackground) _
                                And cboTest.BackColor = vbWindowBackground _
                                And cmdUndo.Enabled
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ValidateCriterion")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub


