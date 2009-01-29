VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QuestionGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------
'   File:       QuestionGroup.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Richard Meinesz, November 2001
'   Purpose:    Question Group Object for MACRO SD
'--------------------------------------------------------------------------------------------------
' Revisions:
'--------------------------------------------------------------------------------------------------

Option Explicit

Private mlStudyID As Long
Private mnVersionId As Integer
Private mlQGroupID As Long
Private msQGroupCode As String
Private msQGroupName As String
Private mnDisplayType As Integer
Private mcolQuestions As Collection
Private mbIsChanged As Boolean
Private mbIsNew As Boolean

Private msQGroupNameStored As String
Private mnDisplayTypeStored As Integer
Private mcolQuestionsStored As Collection

'--------------------------------------------------------------------------------------------------
Public Sub Restore()
'--------------------------------------------------------------------------------------------------
' REM 21/11/01
' Restores the values in the object to those that were stored
'--------------------------------------------------------------------------------------------------
Dim i As Integer

    Set mcolQuestions = New Collection
    
    For i = 1 To mcolQuestionsStored.Count
        mcolQuestions.Add mcolQuestionsStored(i), Str(mcolQuestionsStored(i))
    Next

    msQGroupName = msQGroupNameStored
    mnDisplayType = mnDisplayTypeStored
    
End Sub

'--------------------------------------------------------------------------------------------------
Public Sub Store()
'--------------------------------------------------------------------------------------------------
' REM 21/11/01
' Stores the values that are currently in the object
'--------------------------------------------------------------------------------------------------
Dim i As Integer

    Set mcolQuestionsStored = New Collection
    
    For i = 1 To mcolQuestions.Count
        mcolQuestionsStored.Add mcolQuestions(i), Str(mcolQuestions(i))
    Next
    
    msQGroupNameStored = msQGroupName
    mnDisplayTypeStored = mnDisplayType
    
End Sub

'--------------------------------------------------------------------------------------------------
Public Sub UpdateCRFElements()
'--------------------------------------------------------------------------------------------------
' REM 07/12/01
' Updates the CRFElements after they have been edited
'--------------------------------------------------------------------------------------------------
    
    Call RemoveDeletedCRFElemets
    
    Call AddNewCRFElements
    
    Call ReOrderCRFElements
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub RemoveDeletedCRFElemets()
'--------------------------------------------------------------------------------------------------
' REM 07/12/01
' Delets any CRFElements from the database that were removed during editing an EForm Question Group
'--------------------------------------------------------------------------------------------------
Dim i As Integer
    
    'Loop through the collection of stored Questions and compares them to the current questions
    'and removes any from the database that appear in the stored but not in the current collection
    For i = 1 To mcolQuestionsStored.Count
        If Not CollectionMember(mcolQuestions, Str(mcolQuestionsStored(i)), False) Then
        
            Call DBDeleteCRFGroupElements(mlStudyID, mnVersionId, mlQGroupID, mcolQuestionsStored(i))
        
        End If
    Next

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub AddNewCRFElements()
'--------------------------------------------------------------------------------------------------
' REM 07/12/01
' Adds any new CRFElements to the database
'--------------------------------------------------------------------------------------------------
Dim i As Integer

    'Loop through the collection of current Questions and compares them to the stored questions
    'and adds any questions to the database that appear in the current but not in the stored collection
    For i = 1 To mcolQuestions.Count
        If Not CollectionMember(mcolQuestionsStored, Str(mcolQuestions(i)), False) Then
            
            Call DBInsertCRFGroupMemberElements(mlStudyID, mnVersionId, mcolQuestions(i), mlQGroupID, i)
        
        End If
    Next
        
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub ReOrderCRFElements()
'--------------------------------------------------------------------------------------------------
' REM 07/12/01
' Sets the new order of CRFElements after editing
'--------------------------------------------------------------------------------------------------

    Call DBReorderCRFGroupElements(mlStudyID, mnVersionId, mlQGroupID)

End Sub

'--------------------------------------------------------------------------------------------------
Public Property Get StudyID() As Long
'--------------------------------------------------------------------------------------------------

    StudyID = mlStudyID

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get VersionId() As Integer
'--------------------------------------------------------------------------------------------------

    VersionId = mnVersionId

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get QGroupID() As Long
'--------------------------------------------------------------------------------------------------

    QGroupID = mlQGroupID

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get QGroupCode() As String
'--------------------------------------------------------------------------------------------------

    QGroupCode = msQGroupCode

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get QGroupName() As String
'--------------------------------------------------------------------------------------------------

    QGroupName = msQGroupName

End Property

'--------------------------------------------------------------------------------------------------
Public Property Let QGroupName(sQGroupName As String)
'--------------------------------------------------------------------------------------------------
    
    If sQGroupName <> msQGroupName Then
        msQGroupName = sQGroupName
        mbIsChanged = True
    End If

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get DisplayType() As Integer
'--------------------------------------------------------------------------------------------------

    DisplayType = mnDisplayType

End Property

'--------------------------------------------------------------------------------------------------
Public Property Let DisplayType(nDisplayType As Integer)
'--------------------------------------------------------------------------------------------------
    
    If nDisplayType <> mnDisplayType Then
        mnDisplayType = nDisplayType
        mbIsChanged = True
    End If

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get Questions() As Collection
'--------------------------------------------------------------------------------------------------

    Set Questions = mcolQuestions

End Property

'--------------------------------------------------------------------------------------------------
Public Property Get IsChanged() As Boolean
'--------------------------------------------------------------------------------------------------

    IsChanged = mbIsChanged

End Property

'--------------------------------------------------------------------------------------------------
Public Sub Init(lStudyID As Long, nVersionId As Integer, lQGroupId As Long, sQGroupCode As String, bIsNew As Boolean)
'--------------------------------------------------------------------------------------------------
' REM 16/11/01
' Initialises the modular level variables and set a new collection
'--------------------------------------------------------------------------------------------------

    mlStudyID = lStudyID
    mnVersionId = nVersionId
    mlQGroupID = lQGroupId
    msQGroupCode = sQGroupCode
    
    Set mcolQuestions = New Collection

    mbIsChanged = False
    
    mbIsNew = bIsNew
    
    If bIsNew Then
        msQGroupName = sQGroupCode
        mnDisplayType = 0
    End If
    
End Sub

'--------------------------------------------------------------------------------------------------
Public Sub Save()
'--------------------------------------------------------------------------------------------------
' REM 16/11/01
' Saves the data to the QGroup and QGroupQuestion Tables
'--------------------------------------------------------------------------------------------------
Dim sSQL As String
Dim i As Integer

    On Error GoTo ErrLabel
    
    'If its a new Question Group then insert a new one into the table
    If mbIsNew Then
        'Inserts a new Question Group into the QGroup table
        sSQL = "INSERT INTO QGroup" & _
            " (ClinicalTrialID, VersionID, QGroupID, QGroupCode, QGroupName, DisplayType)" & _
            " VALUES (" & mlStudyID & "," & mnVersionId & "," & mlQGroupID & ",'" & msQGroupCode & "','" & ReplaceQuotes(msQGroupName) & "'," & mnDisplayType & ")"
        MacroADODBConnection.Execute sSQL
        
        'set IsNew to flase
        mbIsNew = False
    
    'If its a Question Group that has been edited then Update the table
    ElseIf mbIsChanged Then
        'Update the Question group tables with the new values
        sSQL = "UPDATE QGroup" & _
               " SET QGroupName = '" & ReplaceQuotes(msQGroupName) & "'," & _
               " DisplayType = " & mnDisplayType & _
               " WHERE ClinicalTrialID = " & mlStudyID & _
               " AND VersionID = " & mnVersionId & _
               " AND QGroupID = " & mlQGroupID
        MacroADODBConnection.Execute sSQL
    
        ' Remove the current rows that correspond to the ones that are to be updated
        sSQL = "DELETE FROM QGroupQuestion" & _
               " WHERE ClinicalTrialID = " & mlStudyID & _
               " AND VersionID = " & mnVersionId & _
               " AND QGroupID = " & mlQGroupID
        MacroADODBConnection.Execute sSQL
        
        'Loop through the collection and inserts the values into the database
        For i = 1 To mcolQuestions.Count
            sSQL = "INSERT INTO QGroupQuestion" & _
                   " (ClinicalTrialID, VersionID, QGroupID, DataItemId, QOrder)" & _
                   " VALUES (" & mlStudyID & "," & mnVersionId & "," & mlQGroupID & "," & mcolQuestions(i) & "," & i & ")"
            MacroADODBConnection.Execute sSQL

        Next
        
        'UpdateCRFElemets if they have been edited
        Call UpdateCRFElements
        
    End If


    mbIsChanged = False
    
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|QuestionGroup.Save"

End Sub

'--------------------------------------------------------------------------------------------------
Public Function QuestionExists(ByVal lDataItemId As Long) As Boolean
'--------------------------------------------------------------------------------------------------
' REM 06/12/01
' Checks to see if a question exits in the collection
'--------------------------------------------------------------------------------------------------
Dim i As Integer

    QuestionExists = CollectionMember(mcolQuestions, Str(lDataItemId), False)

End Function

'--------------------------------------------------------------------------------------------------
Public Sub AddQuestion(lDataItemId As Long)
'--------------------------------------------------------------------------------------------------
' REM 16/11/01
' Adds a question to the collection
'--------------------------------------------------------------------------------------------------
    On Error Resume Next
    mcolQuestions.Add lDataItemId, Str(lDataItemId)
    
    If Err.Number = 0 Then
        mbIsChanged = True
    End If
    
End Sub

'--------------------------------------------------------------------------------------------------
Public Sub RemoveQuestion(lDataItemId As Long)
'--------------------------------------------------------------------------------------------------
' REM 16/11/01
'   Removes a single question  from the collection using the DataItemId
'--------------------------------------------------------------------------------------------------

    mcolQuestions.Remove Str(lDataItemId)
    mbIsChanged = True
    
End Sub

'--------------------------------------------------------------------------------------------------
Public Sub RemoveAllQuestions()
'--------------------------------------------------------------------------------------------------
' REM 21/11/01
' Removes all the questions from the collection
'--------------------------------------------------------------------------------------------------

    Set mcolQuestions = New Collection
    mbIsChanged = True
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
'--------------------------------------------------------------------------------------------------
    Set mcolQuestions = Nothing
End Sub
