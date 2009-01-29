VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EFormGroupsSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------
'   File:       EFormGroupsSD.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Richard Meinesz, November 2001
'   Purpose:    EForm Groups Collection Object for MACRO SD
'--------------------------------------------------------------------------------------------------
' Revisions:
'--------------------------------------------------------------------------------------------------

Option Explicit

Private mcolEFormGroups As Collection
Private mlStudyID As Long
Private mnVersionId As Integer
Private mlCRFPageId As Long

'--------------------------------------------------------------------------------------------------
Public Property Get Count() As Integer
'--------------------------------------------------------------------------------------------------

    Count = mcolEFormGroups.Count
    
End Property

'--------------------------------------------------------------------------------------------------
Public Sub Load(lStudyID As Long, nVersionId As Integer, lCRFPageId As Long)
'--------------------------------------------------------------------------------------------------
' REM 16/11/01
' Loads the EFormQuestion Groups for a specific Study, Version and CRFPage
'--------------------------------------------------------------------------------------------------
Dim oEFG As EFormGroupSD
Dim sSQL As String
Dim rsEFG As ADODB.Recordset
Dim rsElementId As ADODB.Recordset

    On Error GoTo ErrLabel

    mlStudyID = lStudyID
    mnVersionId = nVersionId
    mlCRFPageId = lCRFPageId
    
    Set mcolEFormGroups = New Collection
    
    'Select all from the EFormQgroup table
    sSQL = "SELECT * FROM EFormQGroup" & _
            " WHERE ClinicalTrialID = " & mlStudyID & _
            " AND VersionID = " & mnVersionId & _
            " AND CRFPageID = " & mlCRFPageId
     
    Set rsEFG = New ADODB.Recordset
    rsEFG.Open sSQL, MacroADODBConnection
    
    
    'Loop through the rescordset adding the field values to the collection
    Do While Not rsEFG.EOF
        Set oEFG = New EFormGroupSD
        
        With oEFG
            .Init mlStudyID, mnVersionId, mlCRFPageId, rsEFG.Fields!QGroupID, False
            .Border = rsEFG.Fields!Border
            .DisplayRows = rsEFG.Fields!DisplayRows
            .InitialRows = rsEFG.Fields!InitialRows
            .MinRepeats = rsEFG.Fields!MinRepeats
            .MaxRepeats = rsEFG.Fields!MaxRepeats
        End With
        
        'Add to our EFormGroups collection
        mcolEFormGroups.Add oEFG, Str(oEFG.QGroupID)
        
        'pick up all CRFElements which belong to this group
        sSQL = "SELECT CRFElementID FROM CRFElement" & _
            " WHERE ClinicalTrialID = " & mlStudyID & _
            " AND VersionID = " & mnVersionId & _
            " AND CRFPageID = " & mlCRFPageId & _
            " AND OwnerQGroupID = " & oEFG.QGroupID & _
            " ORDER BY QGroupFieldOrder"
            
        Set rsElementId = New ADODB.Recordset
        rsElementId.Open sSQL, MacroADODBConnection
        
        Do While Not rsElementId.EOF
            oEFG.AddElement rsElementId.Fields!CRFElementId
            rsElementId.MoveNext
        Loop
        
        'move to next record
        rsEFG.MoveNext
    Loop
    
    rsEFG.Close
    Set rsEFG = Nothing
    
    If Not (rsElementId Is Nothing) Then
        rsElementId.Close
        Set rsElementId = Nothing
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormGroupsSD.Load"
    
End Sub

'--------------------------------------------------------------------------------------------------
Public Function NewEFormGroup(lQGroupId As Long, lCRFPageId As Long) As EFormGroupSD
'--------------------------------------------------------------------------------------------------
' REM 17/11/01
' Creates a new EForm Group
'--------------------------------------------------------------------------------------------------
Dim oEFG As EFormGroupSD
Dim sSQL As String

    On Error GoTo ErrLabel

    Set oEFG = New EFormGroupSD
    
    'Initialise and save the new EForm Group
    With oEFG
        .Init mlStudyID, mnVersionId, lCRFPageId, lQGroupId, True
        .Save
    End With
    
    'Add it to the collection
    mcolEFormGroups.Add oEFG, Str(oEFG.QGroupID)
    
    Set NewEFormGroup = oEFG
    
    Set oEFG = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormGroupsSD.NewEFormGroup"

End Function

'--------------------------------------------------------------------------------------------------
Public Function EFormGroupById(lQGroupId As Long) As EFormGroupSD
'--------------------------------------------------------------------------------------------------
' REM 17/11/01
' Returns the EForm QGroup object identified by QGroupID
'--------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    Set EFormGroupById = mcolEFormGroups.Item(Str(lQGroupId))

End Function

'--------------------------------------------------------------------------------------------------
Public Sub Delete(lQGroupId As Long)
'--------------------------------------------------------------------------------------------------
' REM 17/11/01
' Deletes an EFormGroup from the Database and the collection
'--------------------------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel
 
    'Remove the QGroup from the collection
    mcolEFormGroups.Remove Str(lQGroupId)
    
    'Delete QGroup from EFormQGroup table
    sSQL = "DELETE FROM EFormQGroup" & _
            " WHERE ClinicalTrialID = " & mlStudyID & _
            " AND VersionID = " & mnVersionId & _
            " AND CRFPageID = " & mlCRFPageId & _
            " AND QGroupID = " & lQGroupId
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormGroupsSD.Delete"
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
'----------------------------------------------------------------------------------------'

    Set NewEnum = mcolEFormGroups.[_NewEnum]
    
End Function

'--------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
'--------------------------------------------------------------------------------------------------

    Set mcolEFormGroups = Nothing
    
End Sub
