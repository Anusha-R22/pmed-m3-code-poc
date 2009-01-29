VERSION 5.00
Begin VB.Form frmSelectQGroup 
   Caption         =   "Select Question Group"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   2970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   1020
      TabIndex        =   3
      Top             =   3180
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2000
      TabIndex        =   2
      Top             =   3180
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   3180
      Width           =   900
   End
   Begin VB.ListBox lstQGroups 
      Height          =   2985
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2835
   End
End
Attribute VB_Name = "frmSelectQGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999-2006. All Rights Reserved
'   File:       frmSelectQGroup.frm
'   Author:     Richard Meinesz, December 2001
'   Purpose:    To allow the user to select a question group to edit or delete.
'----------------------------------------------------------------------------------------'
'Revisions:
' REM 30/01/02 - added modular level variable to return the QGroup code
' REM 30/01/02 - Added Question Group Code to the dialog question
' NCJ 15 Jun 06 - Consider study access mode
'----------------------------------------------------------------------------------------'

Option Explicit

Private mbOKClicked As Boolean
Private mbDeleteClicked As Boolean
Private mbCancelClicked As Boolean
Private moQGs As QuestionGroups
Private mlQGroupID As Long
Private msQGroupCode As String

'--------------------------------------------------------------------------------------------------
Public Function Display(oQGroups As QuestionGroups, ByRef oQGroup As QuestionGroup, bEdit As Boolean) As Integer
'--------------------------------------------------------------------------------------------------
' REM 28/11/01
' Display form
'--------------------------------------------------------------------------------------------------
Dim lGroupId As Long
    
    Set moQGs = oQGroups
    
    Call FormCentre(frmSelectQGroup)
    
    mbOKClicked = False
    mbDeleteClicked = False
    mbCancelClicked = False
    
    'If editing a QGroup then need an Edit and Delete command button
    If bEdit Then
        With frmSelectQGroup
            .Caption = "Edit Question Group"
            .cmdOK.Caption = "&Edit"
            .cmdDelete.Visible = True
        End With
    End If
    
    'Load all the Question Groups into the list box
    Call LoadQGroupList
    
    ' NCJ 15 Jun 06 - But only allow deleting if user has Full Control
    Call EnableOKDelete
'    'Check to see if an item has been selected, if not OK button disabled
'    If (lstQGroups.ListIndex = -1) Then
'        cmdOK.Enabled = False
'        cmdDelete.Enabled = False
'    End If

    Me.Show vbModal

    If mbOKClicked Then ' OK or Edit
        Set oQGroup = oQGroups.GroupById(mlQGroupID)
        Display = EditQGroup.Edit
    ElseIf mbDeleteClicked Then ' Delete
        Set oQGroup = Nothing
        Display = EditQGroup.Delete
    ElseIf mbCancelClicked Then
        Set oQGroup = Nothing
        Display = EditQGroup.Cancel ' Cancel
    End If

End Function

'--------------------------------------------------------------------------------------------------'
Private Sub EnableOKDelete()
'--------------------------------------------------------------------------------------------------'
' NCJ 15 Jun 06 - Enable OK and Delete buttons
'--------------------------------------------------------------------------------------------------'
    
    'Check to see if an item has been selected, and if they have full control
    cmdOK.Enabled = (lstQGroups.ListIndex > -1)
    cmdDelete.Enabled = (lstQGroups.ListIndex > -1) And (frmMenu.StudyAccessMode = sdFullControl)

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub LoadQGroupList()
'--------------------------------------------------------------------------------------------------'
' REM 28/11/01
' Retrieves all the Question groups from the QuestionGroups Object and loads them into a list box
'--------------------------------------------------------------------------------------------------
Dim oQGroup As QuestionGroup

    For Each oQGroup In moQGs
        lstQGroups.AddItem oQGroup.QGroupCode
        lstQGroups.ItemData(lstQGroups.NewIndex) = oQGroup.QGroupID
    Next

End Sub

'--------------------------------------------------------------------------------------------------
Private Function SelectQGroupID() As Long
'--------------------------------------------------------------------------------------------------'
' REM 28/11/01
' Loops through the list box to see which item was selected
'REVISIONS:
'REM 30/01/02 - added modular level variable to return the QGroup code
'--------------------------------------------------------------------------------------------------
Dim i As Integer
    
    'Loop through the list box and finds the one that was selected
    For i = 0 To lstQGroups.ListCount - 1
        If lstQGroups.Selected(i) Then
            SelectQGroupID = lstQGroups.ItemData(i)
            msQGroupCode = lstQGroups.List(i)
            Exit Function
        End If
    Next
    
End Function

'--------------------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------------------------------------'

    mbCancelClicked = True
    Unload Me

End Sub

'--------------------------------------------------------------------------------------------------'
Private Sub cmdDelete_Click()
'--------------------------------------------------------------------------------------------------'
'REM 28/11/01
' Deleting a Question Group from a Study.  First checks to see if the Question Group resides on any EForms.
'REViSIONS:
'REM 30/01/02 - Added Question Group Code to the dialog question
'--------------------------------------------------------------------------------------------------'
Dim i As Integer
Dim nSelected As Integer

    'gets groupID of the QGroup selected in the list box
    mlQGroupID = SelectQGroupID
            
    'Check to see if the selected QGroup is on any EForms
    If moQGs.IsOnEForm(mlQGroupID) = True Then
        If DialogQuestion("This question group '" & msQGroupCode & "' is used on one or more EForms. Are you sure you want to delete it?") = vbNo Then
            Exit Sub
        End If
    Else
        If DialogQuestion("Are you sure you want to delete the question group '" & msQGroupCode & "' ?") = vbNo Then
            Exit Sub
        End If
    End If
    

    mbDeleteClicked = True

    'Delete the Question Group identified by its QGroupID
    moQGs.Delete (mlQGroupID)
    
    'Remove the selected Question Group name from the list box
    For i = 0 To lstQGroups.ListCount - 1
        If lstQGroups.Selected(i) Then
            nSelected = lstQGroups.ListIndex
        End If
    Next
    lstQGroups.RemoveItem (nSelected)
    
    Unload Me

End Sub

'--------------------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'--------------------------------------------------------------------------------------------------'

    mlQGroupID = SelectQGroupID

    mbOKClicked = True
    Unload Me

End Sub

'--------------------------------------------------------------------------------------------------'
Private Sub Form_Load()
'--------------------------------------------------------------------------------------------------'
    
    Me.Icon = frmMenu.Icon

End Sub

'--------------------------------------------------------------------------------------------------'
Private Sub lstQGroups_Click()
'--------------------------------------------------------------------------------------------------'
' REM 29/11/01
' Enables the OK button if the user selects an item in the list box
'--------------------------------------------------------------------------------------------------'

    Call EnableOKDelete  ' NCJ 15 Jun 06
    
'    If (lstQGroups.ListIndex > -1) Then
'        cmdOK.Enabled = True
'        cmdDelete.Enabled = True
'    End If

End Sub

'--------------------------------------------------------------------------------------------------'
Private Sub lstQGroups_DblClick()
'--------------------------------------------------------------------------------------------------'
' Double click has same effect as clicking the OK button
'--------------------------------------------------------------------------------------------------'
    
    mlQGroupID = SelectQGroupID

    mbOKClicked = True
    Unload Me

End Sub

