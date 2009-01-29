VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QuestionGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------
'   File:       QuestionGroups.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Richard Meinesz, November 2001
'   Purpose:    Question Groups Collection Object for MACRO SD
'--------------------------------------------------------------------------------------------------
' Revisions:
'   NCJ 28 Nov 01 - Added Exit Function before error handler in QGroupById
' NCJ 25 Sept 02 - Must do case-insensitive comparison in CodeExists
'--------------------------------------------------------------------------------------------------

Option Explicit

Private mcolQGroups As Collection
Private mlStudyID As Long
Private mnVersionId As Integer

'--------------------------------------------------------------------------------------------------
Public Property Get Count() As Integer
'--------------------------------------------------------------------------------------------------
' REM 21/11/01
' Count property to count the number of records in the QGroups collection
'--------------------------------------------------------------------------------------------------
    
    Count = mcolQGroups.Count

End Property

'--------------------------------------------------------------------------------------------------
Public Sub Load(lStudyID As Long, nVersionId As Integer)
'--------------------------------------------------------------------------------------------------
' REM 16/11/01
' Loads the Question Group and assoacited Group Questions for a specific study and version
'--------------------------------------------------------------------------------------------------
Dim oQG As QuestionGroup
Dim sSQL As String
Dim rsQG As ADODB.Recordset
Dim rsDataItemId As ADODB.Recordset

    On Error GoTo ErrLabel
    
    mlStudyID = lStudyID
    mnVersionId = nVersionId
    
    Set mcolQGroups = New Collection

    'Selects the Question Group accociated with a specific study and version
    sSQL = "SELECT * FROM QGroup" & _
            " WHERE ClinicalTrialID = " & mlStudyID & _
            " AND VersionId = " & mnVersionId
     
    Set rsQG = New ADODB.Recordset
    rsQG.Open sSQL, MacroADODBConnection

    'Loop through the rescordset adding to the collection
    Do While Not rsQG.EOF
        Set oQG = New QuestionGroup
        
            With oQG
                .Init mlStudyID, mnVersionId, rsQG.Fields!QGroupID, rsQG.Fields!QGroupCode, False
                .QGroupName = rsQG.Fields!QGroupName
                .DisplayType = rsQG.Fields!DisplayType
            End With
        
        'Add to collection
        mcolQGroups.Add oQG, Str(oQG.QGroupID)
        
        'Selects all the Group Questions accoaciated with the study and version
        sSQL = "SELECT DataItemId FROM QGroupQuestion" & _
            " WHERE ClinicalTrialID = " & mlStudyID & _
            " AND VersionID = " & mnVersionId & _
            " AND QGroupID = " & oQG.QGroupID & _
            " ORDER BY QOrder"
            
        Set rsDataItemId = New ADODB.Recordset
        rsDataItemId.Open sSQL, MacroADODBConnection
        
        'Add the questions DataItemId to the object group
        Do While Not rsDataItemId.EOF
            oQG.AddQuestion rsDataItemId.Fields!DataItemId
            rsDataItemId.MoveNext
        Loop
        
        rsQG.MoveNext
    Loop
    
    rsQG.Close
    Set rsQG = Nothing
    
    'If the recordset was never set then don't try and close it
    If Not (rsDataItemId Is Nothing) Then
        rsDataItemId.Close
        Set rsDataItemId = Nothing
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|QuestionGroups.Load"
    
End Sub

'--------------------------------------------------------------------------------------------------
Public Function GroupById(lQGroupId As Long) As QuestionGroup
'--------------------------------------------------------------------------------------------------
' REM 16/11/01
' Returns the Question Group object identified by QGroupID
' Generates an error if it doesn't exist
'--------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrLabel
    
    Set GroupById = mcolQGroups.Item(Str(lQGroupId))

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|QuestionGroups.GroupById(" & lQGroupId & ")"
    
End Function

'--------------------------------------------------------------------------------------------------
Public Function IsOnEForm(ByVal lQGroupId As Long) As Boolean
'--------------------------------------------------------------------------------------------------
' REM 12/12/01
' A check to see if a specific question group is on any EForms
'--------------------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsQG As ADODB.Recordset

    sSQL = "SELECT COUNT(*) FROM EFormQGroup" & _
            " WHERE ClinicalTrialID = " & mlStudyID & _
            " AND VersionID = " & mnVersionId & _
            " AND QGroupID = " & lQGroupId

    Set rsQG = New ADODB.Recordset
    rsQG.Open sSQL, MacroADODBConnection

    'If Count is greater than 0 then the QGroup exists on at least one form
    If rsQG.Fields(0).Value > 0 Then
        IsOnEForm = True
    Else
        IsOnEForm = False
    End If
    
    rsQG.Close
    Set rsQG = Nothing

End Function

'--------------------------------------------------------------------------------------------------
Public Sub Delete(lQGroupId As Long)
'--------------------------------------------------------------------------------------------------
' REM 16/11/01
' Deletes a Question Group and its associated QGroupQuestions
' Generates an error if QGroup doesn't exist
'--------------------------------------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    'delete question group from collection
    mcolQGroups.Remove (Str(lQGroupId))

    Call DBDeleteQuestionGroup(mlStudyID, mnVersionId, lQGroupId)

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|QuestionGroups.Delete(" & lQGroupId & ")"

End Sub

'--------------------------------------------------------------------------------------------------
Public Sub DeleteUnusedQGroups()
'--------------------------------------------------------------------------------------------------
'REM 19/07/02
'Delete all unused question groups in the current study
'--------------------------------------------------------------------------------------------------
Dim rsUnusedQGroups As ADODB.Recordset

    On Error GoTo ErrLabel

    Set rsUnusedQGroups = New ADODB.Recordset
    Set rsUnusedQGroups = QGroupNotOnEForm(mlStudyID, mnVersionId)
    
    Do While Not rsUnusedQGroups.EOF
    
        'delete question group from collection
        mcolQGroups.Remove (Str(rsUnusedQGroups!QGroupID))
        
        Call DBDeleteQuestionGroup(mlStudyID, mnVersionId, rsUnusedQGroups!QGroupID)
        
        rsUnusedQGroups.MoveNext
    Loop

    rsUnusedQGroups.Close
    Set rsUnusedQGroups = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|QuestionGroups.DeleteUnusedGroups"
End Sub


'--------------------------------------------------------------------------------------------------
Public Function NewGroup(sQGroupCode As String) As QuestionGroup
'--------------------------------------------------------------------------------------------------
' REM 16/11/01
' Creates a new Question Group in the Question Group table
'--------------------------------------------------------------------------------------------------
Dim oQG As QuestionGroup
Dim lQGroupId As Long
Dim sSQL As String
Dim rsQGroupID As ADODB.Recordset

    On Error GoTo ErrLabel
    
    Set oQG = New QuestionGroup
    
    'SQL to find max QGroupID
    sSQL = "SELECT MAX (QGroupID) as MaxQGroupID FROM QGroup" & _
           " WHERE ClinicalTrialID = " & mlStudyID & _
           " AND VersionID = " & mnVersionId
    
    Set rsQGroupID = New ADODB.Recordset
    rsQGroupID.Open sSQL, MacroADODBConnection
    
    If IsNull(rsQGroupID.Fields!MaxQGroupID) Then
        'this is the first Group so assign it the QGroupID 1
        lQGroupId = gnFIRST_ID
    Else
        'Add 1 to current max QGroupID
        lQGroupId = rsQGroupID.Fields!MaxQGroupID + gnID_INCREMENT
    End If
    
    'Initialise and save the object
    With oQG
        .Init mlStudyID, mnVersionId, lQGroupId, sQGroupCode, True
        .Save
    End With
    
    'Add  it to the collection
    mcolQGroups.Add oQG, Str(lQGroupId)
    
    Set NewGroup = oQG
    
    Set oQG = Nothing
    rsQGroupID.Close
    Set rsQGroupID = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|QuestionGroups.NewGroup"
    
End Function

'----------------------------------------------------------------------------------------'
Public Function CodeExists(sQGroupCode As String) As Boolean
'----------------------------------------------------------------------------------------'
' REM 21/11/01
' Function to check whether the QGroupCode a user entered already exists
' NCJ 25 Sept 02 - Must do case-insensitive comparison
'----------------------------------------------------------------------------------------'
Dim oQG As QuestionGroup
    
    On Error GoTo ErrLabel
    
    'Loop through the collection of QGroups
    For Each oQG In mcolQGroups
        'If there is a match then the QGroupCode already exists so returns true and exits function
        If lCase(oQG.QGroupCode) = lCase(sQGroupCode) Then
            CodeExists = True
            Exit Function
        End If
    Next
    'Else the code doesn't exist so returns false
    CodeExists = False
   
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|QuestionGroups.CodeExists"
    
End Function

'----------------------------------------------------------------------------------------'
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
'----------------------------------------------------------------------------------------'

    Set NewEnum = mcolQGroups.[_NewEnum]
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'
    Set mcolQGroups = Nothing
End Sub


