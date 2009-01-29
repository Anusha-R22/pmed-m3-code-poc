VERSION 5.00
Begin VB.Form frmEFormQGroupDefinition 
   Caption         =   "EForm Group"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   3780
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3660
      TabIndex        =   6
      Top             =   3780
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   1035
      Left            =   2700
      TabIndex        =   14
      Top             =   60
      Width           =   3555
      Begin VB.TextBox txtGroupCode 
         Height          =   315
         Left            =   1800
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   220
         Width           =   1500
      End
      Begin VB.TextBox txtGroupName 
         Height          =   315
         Left            =   1800
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Group Code"
         Height          =   315
         Left            =   105
         TabIndex        =   18
         Top             =   220
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Group Name"
         Height          =   315
         Left            =   105
         TabIndex        =   15
         Top             =   600
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   2700
      TabIndex        =   9
      Top             =   1200
      Width           =   3555
      Begin VB.TextBox txtMaxRep 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   660
         Width           =   1500
      End
      Begin VB.CheckBox chkDisplayBorder 
         Alignment       =   1  'Right Justify
         Caption         =   "Display Border"
         Height          =   315
         Left            =   560
         TabIndex        =   1
         Top             =   240
         Width           =   1445
      End
      Begin VB.TextBox txtMinRep 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   1080
         Width           =   1500
      End
      Begin VB.TextBox txtInitRows 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   1500
         Width           =   1500
      End
      Begin VB.TextBox txtDisplayRows 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   1920
         Width           =   1500
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Max Repeats"
         Height          =   315
         Left            =   105
         TabIndex        =   13
         Top             =   660
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Min Repeats"
         Height          =   315
         Left            =   105
         TabIndex        =   12
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Initial No. of Rows"
         Height          =   315
         Left            =   105
         TabIndex        =   11
         Top             =   1500
         Width           =   1500
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "No. of Display Rows"
         Height          =   315
         Left            =   105
         TabIndex        =   10
         Top             =   1920
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Group Questions"
      Height          =   3555
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2480
      Begin VB.ListBox lstGroupQuest 
         Height          =   2985
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   2205
      End
   End
End
Attribute VB_Name = "frmEFormQGroupDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999-2006. All Rights Reserved
'   File:       frmEFormQGroupDefinition.frm
'   Author:     Richard Meinesz, November 2001
'   Purpose:    To allow the user to set the properties for an EForm Question Group.
'----------------------------------------------------------------------------------------'
'Revisions:
'   NCJ 28 Feb 02 - Added ValidInteger routine to screen for non-integers
'   NCJ 1 Nov 02 - Rows must now be less than 100 (new msPOSITIVE_ROW text)
'   NCJ 13 Jun 06 - Consider edit mode (for MUSD)
'----------------------------------------------------------------------------------------'

Option Explicit
Private mbOKClicked As Boolean
Private mbIsChanged As Boolean
Private mbIsLoading As Boolean
Private WithEvents moEFG As EFormGroupSD
Attribute moEFG.VB_VarHelpID = -1
Private moQG As QuestionGroup

' NCJ 28 Feb 02 - Error message for non-integer repeat values
Private Const msPOSITIVE_REPEAT = " must be an integer from 1 to 999"
' NCJ 1 Nov 02 - Error message for non-integer row values
Private Const msPOSITIVE_ROW = " must be an integer from 1 to 99"

'--------------------------------------------------------------------------------------------------
Public Function Display(oEFormGroup As EFormGroupSD, oQGroup As QuestionGroup, _
                    bEdit As Boolean) As Boolean
'--------------------------------------------------------------------------------------------------'
' REM 26/11/01
' Display form and load list and text boxes with current values
' NCJ 13 Jun 06 - Added edit mode
'--------------------------------------------------------------------------------------------------
Dim nBorder As Integer

    mbIsLoading = True

    Set moEFG = oEFormGroup
    Set moQG = oQGroup
    
    'Center form on the screen
    Call FormCentre(frmEFormQGroupDefinition)
    
    moEFG.Store
    
    mbOKClicked = False
    mbIsChanged = False

    'Change the border property from a boolean to an integer as a Checkbox requires an integer
    If oEFormGroup.Border = True Then
        nBorder = 1
    Else
        nBorder = 0
    End If

    'Set the text boxs to display the set values
    txtGroupCode.Text = oQGroup.QGroupCode
    txtGroupName.Text = oQGroup.QGroupName
    chkDisplayBorder.Value = nBorder
    txtDisplayRows.Text = oEFormGroup.DisplayRows
    txtInitRows.Text = oEFormGroup.InitialRows
    txtMaxRep.Text = oEFormGroup.MaxRepeats
    txtMinRep.Text = oEFormGroup.MinRepeats
    
    txtGroupCode.Enabled = False
    txtGroupName.Enabled = False

    'Fill the  list box with the Questions in the selected group
    Call FillQGroupList(oEFormGroup.StudyID, oEFormGroup.VersionId, oEFormGroup.QGroupID)

    Call EnableFields(bEdit)
    
    mbIsLoading = False

    Me.Show vbModal
    
    If mbOKClicked Then
        moEFG.Save
    Else
        moEFG.Restore
    End If
    
    Display = mbOKClicked
End Function

'--------------------------------------------------------------------------------------------------
Private Sub FillQGroupList(lStudyID As Long, nVersionId As Integer, lQGroupId As Long)
'--------------------------------------------------------------------------------------------------
' REM 20/11/01
' Fills the list box with all the QGroupQuestions for the particular Study, Version and QGroup
'--------------------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsQGroups As ADODB.Recordset
Dim vData As Variant
Dim i As Integer

    'Clear the list box before filling it
    lstGroupQuest.Clear

    'Select all the QGroupQuestion DataItemId's and accociated Codes from the QGroupQuestion and DataItem Table
    'for a specific StudyId, VersionID and QGroupID
    sSQL = "SELECT QGroupQuestion.DataItemId, DataItem.DataItemCode" & _
            " FROM DataItem, QGroupQuestion" & _
            " WHERE QGroupQuestion.DataItemId = DataItem.DataItemId" & _
            " AND QGroupQuestion.VersionID = DataItem.VersionID" & _
            " AND DataItem.ClinicalTrialID = QGroupQuestion.ClinicalTrialID" & _
            " AND QGroupQuestion.ClinicalTrialID = " & lStudyID & _
            " AND QGroupQuestion.VersionID = " & nVersionId & _
            " AND QGroupQuestion.QGroupID = " & lQGroupId & _
            " ORDER BY QGroupQuestion.QOrder"

    Set rsQGroups = New ADODB.Recordset
    rsQGroups.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

    'Check the recordset contains values and convert it into an array
    'Loop through the array and load it into the list box
    If rsQGroups.RecordCount > 0 Then
        vData = rsQGroups.GetRows
        For i = 0 To UBound(vData, 2)
            lstGroupQuest.AddItem vData(1, i) 'DataItemcode
            lstGroupQuest.ItemData(lstGroupQuest.NewIndex) = vData(0, i) 'DataItemId
        Next
    End If

    rsQGroups.Close
    Set rsQGroups = Nothing

    'makes sure nothing is selected
    lstGroupQuest.ListIndex = -1
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Function ValidInteger(sText As String) As Boolean
'--------------------------------------------------------------------------------------------------
' NCJ 28 Feb 02
' Validates sText - must be an integer and greater than 0
' Assume already trimmed
' Returns FALSE if not a valid integer
'--------------------------------------------------------------------------------------------------
        
    On Error GoTo NotAnInteger
    
    ValidInteger = False
    
    If sText = "" Then Exit Function
    
    If Not IsNumeric(sText) Then Exit Function
    
    If Val(sText) <> CInt(sText) Then Exit Function
    
    ValidInteger = True
        
NotAnInteger:

End Function

'--------------------------------------------------------------------------------------------------
Private Sub chkDisplayBorder_Click()
'--------------------------------------------------------------------------------------------------
' REM 28/11/01
' Changes the Border display from an integer to a boolean and adds it to the object
'--------------------------------------------------------------------------------------------------
Dim bBorder As Boolean

    If mbIsLoading = False Then

        If (chkDisplayBorder.Value = 1) Then
            bBorder = True
        Else
            bBorder = False
        End If
    
        moEFG.Border = bBorder
        
        mbIsChanged = True
        
    End If

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------------------------------------

    mbOKClicked = False
    Unload Me

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'--------------------------------------------------------------------------------------------------

    mbOKClicked = True
    Unload Me
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------------------------------------

    Me.Icon = frmMenu.Icon

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'--------------------------------------------------------------------------------------------------

    If (mbOKClicked = False) And (mbIsChanged = True) Then
        If DialogQuestion("Are you sure you want to cancel without saving?") = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub moEFG_DisplayRowsValid(bValid As Boolean)
'--------------------------------------------------------------------------------------------------

    Call Valid(bValid, txtDisplayRows)

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub moEFG_InitialRowsValid(bValid As Boolean)
'--------------------------------------------------------------------------------------------------

    Call Valid(bValid, txtInitRows)

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub moEFG_MaxRepeatsValid(bValid As Boolean)
'--------------------------------------------------------------------------------------------------

    Call Valid(bValid, txtMaxRep)
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub moEFG_MinRepeatsValid(bValid As Boolean)
'--------------------------------------------------------------------------------------------------

    Call Valid(bValid, txtMinRep)

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub Valid(bIsValid As Boolean, txtTextBox As TextBox)
'--------------------------------------------------------------------------------------------------
' REM 29/11/01
' Changes the text box background colour is invalid data is entered
'--------------------------------------------------------------------------------------------------

    'If bValid is true then text box background white else yellow
    If bIsValid Then
        txtTextBox.BackColor = vbWindowBackground
    Else
        txtTextBox.BackColor = g_INVALID_BACKCOLOUR
    End If

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub moEFG_IsValid(bValid As Boolean)
'--------------------------------------------------------------------------------------------------

    cmdOK.Enabled = bValid
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub txtInitRows_Change()
'--------------------------------------------------------------------------------------------------'
' REM 27/11/01
' Validate text entered and add it to the object
' NCJ 28 Feb 02 - Don't allow real nos.
'--------------------------------------------------------------------------------------------------
Dim sInitRows As String

    On Error GoTo InvalidValue
    
    If mbIsLoading = False Then
    
        sInitRows = Trim(txtInitRows.Text)
        
        If sInitRows = "" Then
            moEFG.InitialRows = 0
        ElseIf Not ValidInteger(sInitRows) Then
            ' Generate dummy error (with random errno)
            Err.Raise 98765, , "Not an integer"
        Else
            moEFG.InitialRows = sInitRows
        End If
        txtInitRows.Tag = txtInitRows.Text
        
        mbIsChanged = True
        
   End If
   
Exit Sub
InvalidValue:
    DialogError "Initial Rows" & msPOSITIVE_ROW
    ' Reset to previous value
    txtInitRows.Text = moEFG.InitialRows
    
End Sub

'--------------------------------------------------------------------------------------------------'
Private Sub txtInitRows_GotFocus()
'--------------------------------------------------------------------------------------------------'
' REM 27/11/01
' Validate text entered and add it to the object
'--------------------------------------------------------------------------------------------------
    
    txtInitRows.SelStart = 0
    txtInitRows.SelLength = Len(txtInitRows.Text)

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub txtMaxRep_Change()
'--------------------------------------------------------------------------------------------------
' REM 27/11/01
' Validate Text entered and add it to the object
' NCJ 28 Feb 02 - Don't allow real nos.
'--------------------------------------------------------------------------------------------------
Dim sMaxRep As String

    On Error GoTo InvalidValue
    
    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
        
        sMaxRep = Trim(txtMaxRep.Text)
        
        If sMaxRep = "" Then
            moEFG.MaxRepeats = 0
            
        ElseIf Not ValidInteger(sMaxRep) Then
            ' Generate dummy error (with random errno)
            Err.Raise 98765, , "Not an integer"
        
        Else
            moEFG.MaxRepeats = sMaxRep
        End If
        
        txtMaxRep.Tag = txtMaxRep.Text
        
        mbIsChanged = True
        
    End If
   
Exit Sub
InvalidValue:
    DialogError "Max Repeats" & msPOSITIVE_REPEAT
    ' Reset to previous value
    txtMaxRep.Text = moEFG.MaxRepeats
End Sub
    
'--------------------------------------------------------------------------------------------------
Private Sub txtMaxRep_GotFocus()
'--------------------------------------------------------------------------------------------------
' REM 11/12/01
' Selects text in the text box as it gets focus
'--------------------------------------------------------------------------------------------------
    
    txtMaxRep.SelStart = 0
    txtMaxRep.SelLength = Len(txtMaxRep.Text)
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub txtMinRep_Change()
'--------------------------------------------------------------------------------------------------
' REM 27/11/01
' Validate text entered and add it to the object
' NCJ 28 Feb 02 - Don't allow real nos.
'--------------------------------------------------------------------------------------------------
Dim sMinRep As String

    On Error GoTo InvalidValue
    
    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
    
        ' Trim spaces
        sMinRep = Trim(txtMinRep.Text)
        
        ' If empty, use value of 0 to get yellow background
        If sMinRep = "" Then
            moEFG.MinRepeats = 0
        ElseIf Not ValidInteger(sMinRep) Then
            ' Generate dummy error (with random errno)
            Err.Raise 98765, , "Not an integer"
        Else
            moEFG.MinRepeats = sMinRep
        End If
        
        txtMinRep.Tag = txtMinRep.Text
        
        mbIsChanged = True
        
    End If
   
Exit Sub
InvalidValue:
    DialogError "Min Repeats" & msPOSITIVE_REPEAT
    ' Reset to previous value
    txtMinRep.Text = moEFG.MinRepeats

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub txtMinRep_GotFocus()
'--------------------------------------------------------------------------------------------------
' REM 11/12/01
' Selects text in the text box as it gets focus
'--------------------------------------------------------------------------------------------------

    txtMinRep.SelStart = 0
    txtMinRep.SelLength = Len(txtMinRep.Text)

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub txtDisplayRows_Change()
'--------------------------------------------------------------------------------------------------
' REM 27/11/01
' Validate text entered and add it to the object
' NCJ 28 Feb 02 - Don't allow real nos.
'--------------------------------------------------------------------------------------------------
Dim sDispRows As String

    On Error GoTo InvalidValue
    
    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
    
        sDispRows = Trim(txtDisplayRows.Text)
        
        If sDispRows = "" Then
            moEFG.DisplayRows = 0
        ElseIf Not ValidInteger(sDispRows) Then
            ' Generate dummy error (with random errno)
            Err.Raise 98765, , "Not an integer"
        Else
            moEFG.DisplayRows = sDispRows
        End If
        txtDisplayRows.Tag = txtDisplayRows.Text
        
        mbIsChanged = True
        
   End If
   
Exit Sub
InvalidValue:
    DialogError "Display Rows" & msPOSITIVE_ROW
    ' Reset to previous value
    txtDisplayRows.Text = moEFG.DisplayRows

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub txtDisplayRows_GotFocus()
'--------------------------------------------------------------------------------------------------
' REM 11/12/01
' Selects text in the text box as it gets focus
'--------------------------------------------------------------------------------------------------

    txtDisplayRows.SelStart = 0
    txtDisplayRows.SelLength = Len(txtDisplayRows.Text)
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub EnableFields(bEdit As Boolean)
'--------------------------------------------------------------------------------------------------
' NCJ 13 Jun 06 - Enable fields depending on edit access
'--------------------------------------------------------------------------------------------------

    cmdOK.Enabled = bEdit
    chkDisplayBorder.Enabled = bEdit
    txtDisplayRows.Enabled = bEdit
    txtInitRows.Enabled = bEdit
    txtMaxRep.Enabled = bEdit
    txtMinRep.Enabled = bEdit
    
End Sub
