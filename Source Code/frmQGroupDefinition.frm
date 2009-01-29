VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmQGroupDefinition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Question Group Definition"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkUnusedQuest 
      Caption         =   "Show unused questions only"
      Height          =   255
      Left            =   190
      TabIndex        =   17
      Top             =   4685
      Width           =   2535
   End
   Begin VB.Frame Frame4 
      Caption         =   "Group Questions"
      Height          =   3375
      Left            =   4380
      TabIndex        =   14
      Top             =   1250
      Width           =   2955
      Begin VB.ListBox lstGroupQuest 
         Height          =   2790
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   16
         Top             =   240
         Width           =   2470
      End
      Begin MSComCtl2.UpDown UpdownQG 
         Height          =   2985
         Left            =   2580
         TabIndex        =   15
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   5265
         _Version        =   393216
         BuddyControl    =   "lstGroupQuest"
         BuddyDispid     =   196611
         OrigLeft        =   1980
         OrigTop         =   300
         OrigRight       =   2220
         OrigBottom      =   2895
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   3975
      Begin VB.TextBox txtGroupCode 
         Height          =   315
         Left            =   1275
         TabIndex        =   11
         Top             =   180
         Width           =   2600
      End
      Begin VB.TextBox txtGroupName 
         Height          =   315
         Left            =   1275
         TabIndex        =   10
         Top             =   675
         Width           =   2600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Group Code"
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Group Name"
         Height          =   315
         Left            =   180
         TabIndex        =   12
         Top             =   675
         Width           =   900
      End
   End
   Begin VB.Frame frameQuestions 
      Caption         =   "All Questions"
      Height          =   3375
      Left            =   60
      TabIndex        =   7
      Top             =   1250
      Width           =   2775
      Begin VB.ListBox lstQuestions 
         Height          =   2790
         Left            =   120
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   2500
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   4740
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   4740
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "< Remove"
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add >"
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Display Type"
      Height          =   1095
      Left            =   4380
      TabIndex        =   0
      Top             =   60
      Width           =   2955
      Begin VB.OptionButton optUser 
         Caption         =   "User-defined"
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   630
         Width           =   1335
      End
      Begin VB.OptionButton optAutofit 
         Caption         =   "Autofit"
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   280
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmQGroupDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999-2006. All Rights Reserved
'   File:       frmQGroupDefinition.frm
'   Author:     Richard Meinesz, November 2001
'   Purpose:    To allow the user to add a new Question Group or edit current Question Groups.
'----------------------------------------------------------------------------------------'
'Revisions:
' REM 30/01/02 - If more than one item in the list is selected disable the updown buttons
' REM 04/09/02 - added check box so that user can either display all questions or just unused questionsin the question list box
' ASH 12/9/2002 - Registry keys replaced with calls to new Settings file in chkUnusedQuest_Click
' NCJ 14 Jun 06 - Added editability checks (for MUSD)
'----------------------------------------------------------------------------------------'

Option Explicit
'REM 04/09/02 - added constants for Unused question list
Private Const msQREGKEY = "GroupDefShowQuestions"
Private Const msALL = "All"
Private Const msUNUSED = "Unused"


Private mbOKClicked As Boolean
Private mbIsChanged As Boolean
Private mbIsLoading As Boolean
Private moQG As QuestionGroup
'REM 04/09/02 - added two modular level collection and Dictionary
Private mcolAllDataItemIds As Collection
Private mdicUnusedDataItemIds As Scripting.Dictionary
Private mdicAllDataitemIds As Scripting.Dictionary
Private mbCanEdit As Boolean    ' NCJ 14 Jun 06

'--------------------------------------------------------------------------------------------------
Public Function Display(oQGroup As QuestionGroup, colDataItemIDs As Collection, bEdit As Boolean) As Boolean
'--------------------------------------------------------------------------------------------------
' REM 22/11/01
' Dispay form, set form attributes and load list boxes
' NCJ 14 Jun 06 - Added bEdit
'--------------------------------------------------------------------------------------------------
Dim sValue As String
Dim sRegPath As String

    mbIsLoading = True
    mbCanEdit = bEdit
    
    Call FormCentre(frmQGroupDefinition)
    
    Set moQG = oQGroup
    
    Set mcolAllDataItemIds = New Collection
    Set mcolAllDataItemIds = colDataItemIDs
    
    Call SetUpAllQuestionDisctionary
    Call SetUpUnusedQuestionDictionary
    
    'Stores all the Questions from the group in the loaded order
    moQG.Store
    
    mbOKClicked = False
    mbIsChanged = False

    ' load the text boxs using the QuestionGroup object
    txtGroupCode.Text = oQGroup.QGroupCode
    txtGroupName.Text = oQGroup.QGroupName
    
    txtGroupCode.Enabled = False
    
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    
    ' Sets the Display Type option buttons
    Call ShowDisplayType(oQGroup.DisplayType)
    
    'REM 04/09/02 - added check box to either display all questions or just unused questions in the question list box
    'load the value from the registry

    'sValue = GetFromRegistry(GetMacroRegistryKey, msQREGKEY)
    
    'ASH 10/9/2002
    sValue = GetMACROSetting(msQREGKEY, "")

    If sValue = msUNUSED Then
        'Call SetUpUnusedQuestionDictionary
        'ensure the check box is selected
        chkUnusedQuest.Value = 1
        'fills the question list box with only the unused questions in the study
        Call FillUnusedQuestionList
    Else ' no registry key set or key set to All
        Call FillAllQuestionList(oQGroup, colDataItemIDs)
        chkUnusedQuest.Value = 0
    End If
    
    ' Fills the Question Group list box with all the Group Questions (only in edit mode, no GroupQuestion until after insert)
    Call FillQGroupList(oQGroup.StudyID, oQGroup.VersionId, oQGroup.QGroupID)
    
    mbIsLoading = False
    
    Call EnableFields(bEdit)
    
    Me.Show vbModal
    
    If mbOKClicked Then
        moQG.Save
        ' NCJ 20 Jun 06 - Mark study as changed
        Call frmMenu.MarkStudyAsChanged
    Else
        moQG.Restore
    End If
    
    Display = mbOKClicked
    
End Function

'--------------------------------------------------------------------------------------------------
Private Sub ReOrderGroupQuestions()
'--------------------------------------------------------------------------------------------------
' REM 22/11/01
' Removes all Questions from the group then reloads them in the order they are in the list box
'--------------------------------------------------------------------------------------------------
Dim lDataItemId As Long
Dim i As Integer
 
    On Error GoTo ErrLabel
 
    'Removes all the questions from the object so they can be reloaded in the correct order
    moQG.RemoveAllQuestions
    
    For i = 1 To lstGroupQuest.ListCount
        lDataItemId = lstGroupQuest.ItemData(i - 1)
        'add the question to the group object
        moQG.AddQuestion (lDataItemId)
    Next

    mbIsChanged = True

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmQGroupDefinition.ReOrderGroupQuestions"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub FillAllQuestionList(oQGroup As QuestionGroup, colDataItemIDs As Collection)
'--------------------------------------------------------------------------------------------------
' REM 20/11/01
' Fills the Question list box with all the questions for the particular Study and Version
'--------------------------------------------------------------------------------------------------
'Dim sSQL As String
'Dim rsQuestions As ADODB.Recordset
Dim vDataItemId As Variant
Dim i As Integer

    On Error GoTo ErrLabel
    
    'Clear the list box before filling it
    lstQuestions.Clear
    frameQuestions.Caption = "All Questions"

    'Select all the DataItem Id's and Codes from the DataItem table corrisponding to the Study and Version
'    sSQL = "SELECT DataItemId, DataItemCode" & _
'            " FROM DataItem" & _
'            " WHERE ClinicalTrialID = " & oQGroup.StudyID & _
'            " AND VersionID = " & oQGroup.VersionId
'
'    Set rsQuestions = New ADODB.Recordset
'    rsQuestions.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText


    If mdicAllDataitemIds.Count > 0 Then
        For Each vDataItemId In mdicAllDataitemIds
            If Not moQG.QuestionExists(vDataItemId) Then
                If Not CollectionMember(colDataItemIDs, Str(vDataItemId), False) Then
                    lstQuestions.AddItem mdicAllDataitemIds.Item(vDataItemId)
                    lstQuestions.ItemData(lstQuestions.NewIndex) = vDataItemId
                End If
            End If
        Next
    End If
    
    'Check the recordset contains values and convert it into an array
    'Loop through the array and load it into the list box
'    If rsQuestions.RecordCount > 0 Then
'        vData = rsQuestions.GetRows
'
'        For i = 0 To UBound(vData, 2)
'            'Check to see if the question is in  the group list, if so it must not be added to the all question list
'            If Not moQG.QuestionExists(vData(0, i)) Then
'                If Not CollectionMember(colDataItemIDs, Str(vData(0, i)), False) Then
'                    lstQuestions.AddItem vData(1, i) 'DataItemcode
'                    lstQuestions.ItemData(lstQuestions.NewIndex) = vData(0, i) 'DataItemId
'                End If
'            End If
'        Next
'    End If
'
'    rsQuestions.Close
'    Set rsQuestions = Nothing
    
    'makes sure nothing is selected
    lstQuestions.ListIndex = -1
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmQGroupDefinition.FillAllQuestionList"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub FillUnusedQuestionList()
'--------------------------------------------------------------------------------------------------
'REM 04/09/02
'Fills the question list box with all unused questions in thje study, i.e. questions that are not on an eForm
'or used in any Question Groups (which are used or unused)
'--------------------------------------------------------------------------------------------------
Dim i As Integer
Dim vDataItemId As Variant
    
    On Error GoTo ErrLabel
    
    'Clear the list box before filling it
    lstQuestions.Clear
    'Change the caption
    frameQuestions.Caption = "Unused Questions"
    
    'Check to see if there are any unused questions in the dictionary
    If mdicUnusedDataItemIds.Count > 0 Then
        'fill the list box
        For Each vDataItemId In mdicUnusedDataItemIds
            lstQuestions.AddItem mdicUnusedDataItemIds.Item(vDataItemId)
            lstQuestions.ItemData(lstQuestions.NewIndex) = vDataItemId
        Next
    
    End If

    lstQuestions.ListIndex = -1
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmQGroupDefinition.FillUnusedQuestionList"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub FillQGroupList(lStudyID As Long, nVersionId As Integer, lQGroupId As Long)
'--------------------------------------------------------------------------------------------------
' REM 20/11/01
' Fills the list box with all the QGroupQuestions for the particular Study, Version and QGroup
' Will only fill the list box in edit mode as QGroupQuestions are only added during an Insert
'--------------------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsQGroups As ADODB.Recordset
Dim vData As Variant
Dim i As Integer
    
    On Error GoTo ErrLabel
    
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
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmQGroupDefinition.FillQGroupList"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub AddQuestionToGroup()
'--------------------------------------------------------------------------------------------------
' REM 20/11/01
' Adds a question to the Question Group list box and removes it from the Question list box
'--------------------------------------------------------------------------------------------------
Dim sText As String
Dim lDataItemId As Long
Dim i As Integer
Dim j As Integer

    On Error GoTo ErrLabel

    'Loops through the list box to check for multi-selected items
    ' then add them to the Group Question list box
    For i = 0 To lstQuestions.ListCount - 1
    
        If lstQuestions.Selected(i) = True Then
            'Text from the list box
            sText = lstQuestions.List(i)
            'Associated DataItemId
            lDataItemId = lstQuestions.ItemData(i)
        
            'Add text from lstQuestions to lstGroupQuest
            lstGroupQuest.AddItem sText
            'add associated DataItemId to lstGroupQuest
            lstGroupQuest.ItemData(lstGroupQuest.NewIndex) = lDataItemId
            'added text is selected in the lstGroupQuest list box
            lstGroupQuest.Selected(lstGroupQuest.NewIndex) = True
        
            'add the question to the group object
            Call moQG.AddQuestion(lDataItemId)
            
            'REM 04/09/02
            'Checks to see if the question being added to the group exists in the unused question list, if so removes it
            If mdicUnusedDataItemIds.Exists(lDataItemId) Then
                'Add to unused question dictionary
                mdicUnusedDataItemIds.Remove lDataItemId
            End If

            'REM 05/09/02 - Remove question being added to group from all question list
            mdicAllDataitemIds.Remove lDataItemId
   
            'set list index to be the new index after being added to the list box
            lstGroupQuest.ListIndex = lstGroupQuest.NewIndex
            
        End If
    Next
    
    'Loop through all the questions and delete selected ones
    For j = lstQuestions.ListCount - 1 To 0 Step -1
        If lstQuestions.Selected(j) = True Then
            'Remove the text and DataItemId from lstQuestions
            Call lstQuestions.RemoveItem(j)
        End If
    Next

    'ensures that no items in the AllQuest list box are selected
    lstQuestions.ListIndex = -1

    'call the lstQuestions click event to set the command button enabled properties
    lstQuestions_Click
    
    mbIsChanged = True

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmQGroupDefinition.AddQuestionToGroup"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub RemoveQuestionFromGroup()
'--------------------------------------------------------------------------------------------------
' REM 20/11/01
' Removes a question from the Question Group List Box
'--------------------------------------------------------------------------------------------------
Dim sText As String
Dim lDataItemId As Long
Dim i As Integer

    On Error GoTo ErrLabel

    'Loops through the list box to check for multislected items
    For i = lstGroupQuest.ListCount - 1 To 0 Step -1
    
        If lstGroupQuest.Selected(i) = True Then
            'Text from the list box
            sText = lstGroupQuest.List(i)
            'Associated DataItemId
            lDataItemId = lstGroupQuest.ItemData(i)
        
            'Add text from lstGroupQuest to lstQuestions
            lstQuestions.AddItem sText
            'add associated DataItemId to lstGroupQuest
            lstQuestions.ItemData(lstQuestions.NewIndex) = lDataItemId
            'added text is selected in the lstQuestions list box
            lstQuestions.Selected(lstQuestions.NewIndex) = True
        
            'Remove the text and DataItemId from lstGroupQuest
            Call lstGroupQuest.RemoveItem(i)
            
            'removes the question from the object group
            Call moQG.RemoveQuestion(lDataItemId)
            
            ' REM 04/09/02 - Add to unused question dictionary regardless whether it is in fact an unused question
            ' reason for this is it must still be avaliable to the user to add back to the group if the user is
            ' only viewing the unused question list
            mdicUnusedDataItemIds.Add lDataItemId, sText
            
            'REM 05/09/02 - only add the question to the all question list if its not already there
            If Not mdicAllDataitemIds.Exists(lDataItemId) Then
                'add to the all question dictionary
                mdicAllDataitemIds.Add lDataItemId, sText
            End If
        
            'set list index to be the new index after being added to the list box
            lstQuestions.ListIndex = lstQuestions.NewIndex
            
        End If
    Next
    
    'ensures that no items in the GroupQuest listbox are selected
    lstGroupQuest.ListIndex = -1
        
    'call the lstGroupQuest click event to set the command button enabled properties
    lstGroupQuest_Click
    
    mbIsChanged = True
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmQGroupDefinition.RemoveQuestionFromGroup"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub chkUnusedQuest_Click()
'--------------------------------------------------------------------------------------------------
'REM 04/09/02
'By selecting the check box the list will only display unsed questions instead of all questions
'ASH 12/9/2002 - Registry keys replaced with calls to new Settings file
'--------------------------------------------------------------------------------------------------
Dim sRegPath As String

    On Error GoTo ErrLabel
    
    'get MACRO registry path
    'sRegPath = GetMacroRegistryKey
    
    If chkUnusedQuest.Value = 0 Then
        'save key values
        'SetKeyValue sRegPath, msQREGKEY, msALL, REG_SZ
        'ASH 10/9/2002
        Call SetMACROSetting(msQREGKEY, msALL)
        Call FillAllQuestionList(moQG, mcolAllDataItemIds)
    ElseIf chkUnusedQuest.Value = 1 Then
        'save key values
        'SetKeyValue sRegPath, msQREGKEY, msUNUSED, REG_SZ
        'ASH 10/9/2002
        Call SetMACROSetting(msQREGKEY, msUNUSED)
        'Call SetUpUnusedQuestionDictionary
        Call FillUnusedQuestionList
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmQGroupDefinition.chkUnusedQuest_Click"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------------------------------------

    Me.Icon = frmMenu.Icon
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub lstQuestions_Click()
'--------------------------------------------------------------------------------------------------
' REM 21/11/01
' Checks to make sure an item has been selected in the AllQuest list box before enabling the Add command button
'--------------------------------------------------------------------------------------------------
Dim i As Integer

    If lstQuestions.ListIndex > -1 And mbCanEdit Then
        cmdAdd.Enabled = True
    Else
        cmdAdd.Enabled = False
    End If
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub lstQuestions_DblClick()
'--------------------------------------------------------------------------------------------------
' REM 21/11/01
' Adds a question to the Group Question list and removes it from the AllQuest list box
'--------------------------------------------------------------------------------------------------
    
    If mbCanEdit Then Call AddQuestionToGroup
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub cmdAdd_Click()
'--------------------------------------------------------------------------------------------------
' REM 21/11/01
' Adds a question to the Group Question list
'--------------------------------------------------------------------------------------------------
    
    Call AddQuestionToGroup

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub lstGroupQuest_Click()
'--------------------------------------------------------------------------------------------------
' REM 21/11/01
' Checks to make sure an item has been selected in the GroupQuest list box before enabling the Remove
' command button and the updown buttons
'REVISIONS:
' REM 30/01/02 - If more than one item in the list is selected disable the updown buttons
'--------------------------------------------------------------------------------------------------
Dim i As Integer
Dim nSelected As Integer

    On Error GoTo ErrLabel
        
    If lstGroupQuest.ListIndex > -1 And mbCanEdit Then
        cmdRemove.Enabled = True
        'only enable the updown if there is more than one item in the list
        UpdownQG.Enabled = (lstGroupQuest.ListCount > 1)
        
        'REM 30/01/02 - If more than one item in the list is selected disable the updown buttons
        nSelected = 0
        For i = 0 To lstGroupQuest.ListCount - 1
            If lstGroupQuest.Selected(i) = True Then
                nSelected = nSelected + 1
                If nSelected > 1 Then
                    UpdownQG.Enabled = False
                    Exit For
                End If
            End If
        Next
    Else
        cmdRemove.Enabled = False
        UpdownQG.Enabled = False
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmQGroupDefinition.lstGroupQuest_Click"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub lstGroupQuest_DblClick()
'--------------------------------------------------------------------------------------------------
' REM 21/11/01
' Adds a question to the AllQuest list box and removes it from the GroupQuest list box
'--------------------------------------------------------------------------------------------------
    
    If mbCanEdit Then Call RemoveQuestionFromGroup
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------------------------------------
' REM 21/11/01
' cancel button
'--------------------------------------------------------------------------------------------------
  
    mbOKClicked = False
    Unload Me

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'--------------------------------------------------------------------------------------------------
' REM 21/11/01
' When the OK button is clicked the object group is saved
'--------------------------------------------------------------------------------------------------

    mbOKClicked = True
    Unload Me

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub cmdRemove_Click()
'--------------------------------------------------------------------------------------------------
' REM 21/11/01
' Adds a question to the Question list box and removes it from the GroupQuest list box
'--------------------------------------------------------------------------------------------------

    Call RemoveQuestionFromGroup

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub ShowDisplayType(nDisplayType As Integer)
'--------------------------------------------------------------------------------------------------
' REM 21/11/01
' Sets the DisplayType option buttons
'--------------------------------------------------------------------------------------------------
    'Set both the option buttons value to false first
    optAutofit.Value = False
    optUser.Value = False
    
    Select Case nDisplayType
    Case 0 ' Autofit
        optAutofit.Value = True
    Case 1 ' User-defined
        optUser.Value = True
    End Select
        
    'This is a temp measure for 3.0 as there is no User-defined option yet
    optUser.Enabled = False
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub txtGroupName_Change()
'--------------------------------------------------------------------------------------------------
' REM 23/11/01
' Validates the test typed into the GroupName text box
'--------------------------------------------------------------------------------------------------
Dim sName As String
Dim bValid As Boolean

    On Error GoTo ErrLabel

    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
    
        'trim all spaces
        sName = Trim(txtGroupName.Text)
        
        ' If there is no text or invalid text is typed
        ' then change back colour to yellow and disable the OK button
        ' NCJ 1 Mar 02 - trap text of more than 255 chars
        bValid = False
        If (sName = "") Then
            ' Not valid
        ElseIf (Len(sName) > 255) Then
            ' Not valid
        ElseIf (Not gblnValidString(sName, valOnlySingleQuotes)) Then
            ' Not valid
        Else
            bValid = True
        End If
        
        If bValid Then
            cmdOK.Enabled = True
            txtGroupName.BackColor = vbWindowBackground
            moQG.QGroupName = sName
        Else
            cmdOK.Enabled = False
            txtGroupName.BackColor = g_INVALID_BACKCOLOUR
        End If
        
        mbIsChanged = True
        
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmQGroupDefinition.txtGroupName_Change"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub UpdownQG_UpClick()
'--------------------------------------------------------------------------------------------------
' REM 22/11/01
' Moves the selected item in the GroupQuest list box up one item in the list at a time
'--------------------------------------------------------------------------------------------------
Dim sText As String
Dim lDataItemId As Long
Dim nIndex As Integer
Dim nNewIndex As Integer
    
    On Error GoTo ErrLabel

    'Checks to see if an item has been selected in the list box and that it is not the top one,
    'as the top one cannot be moved up
    If lstGroupQuest.ListIndex > 0 Then
        nIndex = lstGroupQuest.ListIndex
        'Text from the list box
        sText = lstGroupQuest.List(nIndex)
        'Associated DataItemId
        lDataItemId = lstGroupQuest.ItemData(nIndex)
        
        'Remove the text and DataItemId from lstGroupQuest
        lstGroupQuest.RemoveItem (nIndex)
        
        'Suntract 1 from the current index
        nNewIndex = (nIndex - 1)
        
        'Add text from lstGroupQuest back to lstGroupQuest
        lstGroupQuest.AddItem sText, nNewIndex
        'add associated DataItemId
        lstGroupQuest.ItemData(lstGroupQuest.NewIndex) = lDataItemId
        
        'Select the item that has just been moved
        lstGroupQuest.ListIndex = nNewIndex
        'REM 29/01/02 - highlights the moved item
        lstGroupQuest.Selected(nNewIndex) = True

        'reorders the GroupQuestions in the object
        Call ReOrderGroupQuestions
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmQGroupDefinition.UpdownQG_UpClick"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub UpdownQG_DownClick()
'--------------------------------------------------------------------------------------------------
' REM 22/11/01
' Moves the selected item in the GroupQuest list box down one item in the list at a time
'--------------------------------------------------------------------------------------------------
Dim sText As String
Dim lDataItemId As Long
Dim nIndex As Integer
Dim nNewIndex As Integer
Dim nCount As Integer

    On Error GoTo ErrLabel
    
    'Check to see if an item has been slected in the list box
    If lstGroupQuest.ListIndex > -1 Then
        nIndex = lstGroupQuest.ListIndex
        'Text from the list box
        sText = lstGroupQuest.List(nIndex)
        'Associated DataItemId
        lDataItemId = lstGroupQuest.ItemData(nIndex)
        
        nNewIndex = (nIndex + 1)
        
        'Count the number of items in the listbox
        nCount = lstGroupQuest.ListCount
        
        'If the item being moved reaches the end of the list then exit the sub
        If nNewIndex >= nCount Then
            Exit Sub
        End If
        
        'Remove the text and DataItemId from lstGroupQuest
        lstGroupQuest.RemoveItem (nIndex)
    
        'Add text back to lstGroupQuest at the New Index
        lstGroupQuest.AddItem sText, nNewIndex
        'add associated DataItemId
        lstGroupQuest.ItemData(lstGroupQuest.NewIndex) = lDataItemId
        
        'Select the item that has just been moved
        lstGroupQuest.ListIndex = nNewIndex
        'REM 29/01/02 - highlights the moved item
        lstGroupQuest.Selected(nNewIndex) = True

        'reorders the GroupQuestions in the object
        Call ReOrderGroupQuestions
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmQGroupDefinition.UpdownQG_DownClick"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'--------------------------------------------------------------------------------------------------
' REM 11/12/01
' Asks the user if they want to save after they click cancel
'--------------------------------------------------------------------------------------------------
    
    If (mbOKClicked = False) And mbIsChanged And (mbCanEdit = True) Then
        If DialogQuestion("Are you sure you want to cancel without saving?") = vbNo Then
            Cancel = 1
        End If
    End If

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub SetUpUnusedQuestionDictionary()
'--------------------------------------------------------------------------------------------------
'REM 04/09/02
'Sets up a dictionary of all the unused questions in a study
'--------------------------------------------------------------------------------------------------
Dim rsUnusedQuestions As ADODB.Recordset

    On Error GoTo ErrLabel

    Set rsUnusedQuestions = New ADODB.Recordset
    'recordset of all unused questions
    Set rsUnusedQuestions = UnusedQuestionList(moQG.StudyID, moQG.VersionId)
    
    Set mdicUnusedDataItemIds = New Scripting.Dictionary
    
    'add all unused questions to the dictionary
    Do While Not rsUnusedQuestions.EOF
        ' Add each unused question to the dictionary
        mdicUnusedDataItemIds.Add CLng(rsUnusedQuestions!DataItemId), CStr(rsUnusedQuestions!DataItemCode)

        rsUnusedQuestions.MoveNext
    Loop
    
    rsUnusedQuestions.Close
    Set rsUnusedQuestions = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmQGroupDefinition.UnusedQuestionDictionary"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub SetUpAllQuestionDisctionary()
'--------------------------------------------------------------------------------------------------
'REM 05/09/02
'Sets up a dictionary of all the questions in a study
'--------------------------------------------------------------------------------------------------
Dim rsAllQuestions As ADODB.Recordset

    On Error GoTo ErrLabel
    
    Set rsAllQuestions = New ADODB.Recordset
    
    Set rsAllQuestions = AllQuestionList(moQG.StudyID, moQG.VersionId)
    
    Set mdicAllDataitemIds = New Scripting.Dictionary
    
    Do While Not rsAllQuestions.EOF
        
        mdicAllDataitemIds.Add CLng(rsAllQuestions!DataItemId), CStr(rsAllQuestions!DataItemCode)
        
        rsAllQuestions.MoveNext
        
    Loop
    
    rsAllQuestions.Close
    Set rsAllQuestions = Nothing
    
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmQGroupDefinition.SetUpAllQuestionDisctionary"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub EnableFields(bEdit As Boolean)
'--------------------------------------------------------------------------------------------------
' NCJ 14 Jun 06 - Enable/disable fields according to whether the user can edit
'--------------------------------------------------------------------------------------------------

    cmdOK.Enabled = bEdit
    optAutofit.Enabled = bEdit
    txtGroupName.Enabled = bEdit
    chkUnusedQuest.Enabled = bEdit
    
End Sub
