Attribute VB_Name = "modCopyPatientData"
'------------------------------------------------------------------------------'
' File:         modCopyPatientData.bas
' Copyright:    InferMed Ltd. 2001. All Rights Reserved
' Author:       Ashitei Trebi-Ollennu, July 2001
' Purpose:      Contains  routines for copying Patient / Subject data
'------------------------------------------------------------------------------'
'   Revisions:
'   ZA 19/06/2002 - added new functions GetMaxMisseageID and GetMaxMIMessageObjectID
'   REM 03/07/02 - Added new routine for RQG's - CopyQGroupInstance
'-------------------------------------------------------------------------------
Option Explicit

'----------------------------------------------------------------------------------------
Public Sub CopyVisitInstances(ByVal lOldClinicalTrialId As Long, _
                            ByVal lNewClinicalTrialId As Long)
'-----------------------------------------------------------------------------------------
'duplicates existing VisitInstances rows with old clinical trial ID under the new ID
'------------------------------------------------------------------------------------------
Dim rsVisitInstances As ADODB.Recordset
Dim rsCopyVisitInstances As ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim sSQL As String
Dim sSQL1 As String

     On Error GoTo ErrLabel
     
    Screen.MousePointer = vbHourglass
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from VisitInstance " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
    'creates recordset to contain records to be copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from VisitInstance " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsVisitInstances = New ADODB.Recordset
    rsVisitInstances.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyVisitInstances = New ADODB.Recordset
    rsCopyVisitInstances.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
   
    'checks if records exist
    If rsVisitInstances.RecordCount <= 0 Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsVisitInstances.MoveFirst
    
    'begin record insertion
     For j = 1 To rsVisitInstances.RecordCount
        rsCopyVisitInstances.AddNew
        rsCopyVisitInstances.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsVisitInstances.Fields.Count - 1
                rsCopyVisitInstances.Fields(i).Value = rsVisitInstances.Fields(i).Value
            Next
        rsCopyVisitInstances.Update
        rsVisitInstances.MoveNext
    Next j

    rsVisitInstances.Close
    Set rsVisitInstances = Nothing
    rsCopyVisitInstances.Close
    Set rsCopyVisitInstances = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyPatientData.CopyVisitInstances"

End Sub

'----------------------------------------------------------------------------
Public Sub CopyDataItemResponses(ByVal lOldClinicalTrialId As Long, _
                                ByVal lNewClinicalTrialId As Long)
'-----------------------------------------------------------------------------
'duplicates dataitemresponses rows with old clinical trial ID under the new ID
'-----------------------------------------------------------------------------
Dim rsDataItemResponses As ADODB.Recordset
Dim rsCopyDataItemResponses As ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim sSQL As String
Dim sSQL1 As String

     On Error GoTo ErrLabel
     
    Screen.MousePointer = vbHourglass
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from DataItemResponse " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
    'creates recordset to contain records to be copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from DataItemResponse " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsDataItemResponses = New ADODB.Recordset
   rsDataItemResponses.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyDataItemResponses = New ADODB.Recordset
    rsCopyDataItemResponses.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
   
    'checks if records exist
    If rsDataItemResponses.RecordCount <= 0 Then
        'added 3/08/2001
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    'move to first record inrecordset
   rsDataItemResponses.MoveFirst
    
    'begin record insertion
     For j = 1 To rsDataItemResponses.RecordCount
        rsCopyDataItemResponses.AddNew
        rsCopyDataItemResponses.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsDataItemResponses.Fields.Count - 1
                rsCopyDataItemResponses.Fields(i).Value = rsDataItemResponses.Fields(i).Value
            Next
        rsCopyDataItemResponses.Update
        rsDataItemResponses.MoveNext
    Next j
    rsDataItemResponses.Close
    Set rsDataItemResponses = Nothing
    rsCopyDataItemResponses.Close
    Set rsCopyDataItemResponses = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyPatientData.CopyDataItemResponses"

End Sub

'-------------------------------------------------------------------------------
Public Sub CopyDataItemResponseHistorys(ByVal lOldClinicalTrialId As Long, _
                                ByVal lNewClinicalTrialId As Long)
'--------------------------------------------------------------------------------
'duplicates dataitemhistory rows with old clinical trial ID under the new ID
'---------------------------------------------------------------------------------
Dim rsDataItemResponseHistorys As ADODB.Recordset
Dim rsCopyDataItemResponseHistorys As ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim sSQL As String
Dim sSQL1 As String

     On Error GoTo ErrLabel
     
      Screen.MousePointer = vbHourglass
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from DataItemResponseHistory " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
    'creates recordset to contain records to be copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from DataItemResponseHistory " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsDataItemResponseHistorys = New ADODB.Recordset
   rsDataItemResponseHistorys.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyDataItemResponseHistorys = New ADODB.Recordset
    rsCopyDataItemResponseHistorys.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
   
    'checks if records exist
    If rsDataItemResponseHistorys.RecordCount <= 0 Then
        'added 3/08/2001
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    'move to first record inrecordset
   rsDataItemResponseHistorys.MoveFirst
    
    'begin record insertion
     For j = 1 To rsDataItemResponseHistorys.RecordCount
        rsCopyDataItemResponseHistorys.AddNew
        rsCopyDataItemResponseHistorys.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsDataItemResponseHistorys.Fields.Count - 1
                rsCopyDataItemResponseHistorys.Fields(i).Value = rsDataItemResponseHistorys.Fields(i).Value
            Next
        rsCopyDataItemResponseHistorys.Update
        rsDataItemResponseHistorys.MoveNext
    Next j
    rsDataItemResponseHistorys.Close
    Set rsDataItemResponseHistorys = Nothing
    rsCopyDataItemResponseHistorys.Close
    Set rsCopyDataItemResponseHistorys = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyPatientData.CopyDataItemResponseHistorys"

End Sub

'-------------------------------------------------------------------------------
Public Sub CopyTrialSubjects(ByVal lOldClinicalTrialId As Long, _
                            ByVal lNewClinicalTrialId As Long)
'-------------------------------------------------------------------------------
'dupilcates  trial subjects rows with old clinical trial ID under the new ID
'-------------------------------------------------------------------------------
Dim rsTrialSubjects As ADODB.Recordset
Dim rsCopyTrialSubjects As ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim sSQL As String
Dim sSQL1 As String

     On Error GoTo ErrLabel
     
      Screen.MousePointer = vbHourglass
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from TrialSubject " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
    'creates recordset to contain records to be copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from TrialSubject " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsTrialSubjects = New ADODB.Recordset
   rsTrialSubjects.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyTrialSubjects = New ADODB.Recordset
    rsCopyTrialSubjects.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
   
    'checks if records exist
    If rsTrialSubjects.RecordCount <= 0 Then
        'added 3/08/2001
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    'move to first record inrecordset
   rsTrialSubjects.MoveFirst
    
    'begin record insertion
     For j = 1 To rsTrialSubjects.RecordCount
        rsCopyTrialSubjects.AddNew
        rsCopyTrialSubjects.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsTrialSubjects.Fields.Count - 1
                rsCopyTrialSubjects.Fields(i).Value = rsTrialSubjects.Fields(i).Value
            Next
        rsCopyTrialSubjects.Update
        rsTrialSubjects.MoveNext
    Next j
    rsTrialSubjects.Close
    Set rsTrialSubjects = Nothing
    rsCopyTrialSubjects.Close
    Set rsCopyTrialSubjects = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyPatientData.CopyTrialSubjects"

End Sub

'----------------------------------------------------------------------------------
Public Sub CopyCRFPageInstances(ByVal lOldClinicalTrialId As Long, _
                                ByVal lNewClinicalTrialId As Long)
'-----------------------------------------------------------------------------------
'duplicates  crfpageinstances rows with old clinical trial ID under the new ID
'-----------------------------------------------------------------------------------
Dim rsCRFPageInstances As ADODB.Recordset
Dim rsCopyCRFPageInstances As ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim sSQL As String
Dim sSQL1 As String

     On Error GoTo ErrLabel
     
    Screen.MousePointer = vbHourglass
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from CRFPageInstance " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
    'creates recordset to contain records to be copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from CRFPageInstance " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsCRFPageInstances = New ADODB.Recordset
   rsCRFPageInstances.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyCRFPageInstances = New ADODB.Recordset
    rsCopyCRFPageInstances.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
   
    'checks if records exist
    If rsCRFPageInstances.RecordCount <= 0 Then
        'added 3/08/2001
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    'move to first record inrecordset
   rsCRFPageInstances.MoveFirst
    
    'begin record insertion
     For j = 1 To rsCRFPageInstances.RecordCount
        rsCopyCRFPageInstances.AddNew
        rsCopyCRFPageInstances.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsCRFPageInstances.Fields.Count - 1
                rsCopyCRFPageInstances.Fields(i).Value = rsCRFPageInstances.Fields(i).Value
            Next
        rsCopyCRFPageInstances.Update
        rsCRFPageInstances.MoveNext
    Next j
    rsCRFPageInstances.Close
    Set rsCRFPageInstances = Nothing
    rsCopyCRFPageInstances.Close
    Set rsCopyCRFPageInstances = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyPatientData.CopyCRFPageInstances"

End Sub

'------------------------------------------------------------------------------------
Public Sub CopyRSNextNumbers(ByVal sClinicalTrialName As String, _
                            ByVal sNewClinicalTrialName As String)
'-------------------------------------------------------------------------------------
'duplicates rows of rsnextnumbers under new clinical trial name
'-------------------------------------------------------------------------------------
Dim rsRSNextNumbers As ADODB.Recordset
Dim rsCopyRSNextNumbers As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer
    
    On Error GoTo ErrLabel
    
    Screen.MousePointer = vbHourglass
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from RSNextNumber " _
    & " Where ClinicalTrialName = '" & sClinicalTrialName & "'"
    
    'creates recordset to contain records to be copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from RSNextNumber " _
    & " WHERE ClinicalTrialName = '" & "'"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsRSNextNumbers = New ADODB.Recordset
    rsRSNextNumbers.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyRSNextNumbers = New ADODB.Recordset
    rsCopyRSNextNumbers.Open sSQL1, MacroADODBConnection, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    

    'checks if records exist
    If rsRSNextNumbers.RecordCount <= 0 Then
        'added 3/08/2001
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    'move to first record in recordset
    rsRSNextNumbers.MoveFirst
    
    'begin record insertion
    Do While Not rsRSNextNumbers.EOF And Not rsRSNextNumbers.BOF
        rsCopyRSNextNumbers.AddNew
        rsCopyRSNextNumbers.Fields(0).Value = sNewClinicalTrialName
        For i = 1 To rsRSNextNumbers.Fields.Count - 1
            rsCopyRSNextNumbers.Fields(i).Value = rsRSNextNumbers.Fields(i).Value
        Next
        rsCopyRSNextNumbers.Update
        rsRSNextNumbers.MoveNext
    Loop

    rsRSNextNumbers.Close
    Set rsRSNextNumbers = Nothing
    rsCopyRSNextNumbers.Close
    Set rsCopyRSNextNumbers = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyPatientData.CopyRSNextNumbers"

End Sub

'-------------------------------------------------------------------------------------
Public Sub CopyRSSubjectIdentifiers(ByVal sClinicalTrialName As String, _
                                ByVal sNewClinicalTrialName As String)
'-------------------------------------------------------------------------------------
'duplicates rows of rssubject identifiers under new study name
'--------------------------------------------------------------------------------------
Dim rsRSSubjectIdentifiers As ADODB.Recordset
Dim rsCopyRSSubjectIdentifiers As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer
    
    On Error GoTo ErrLabel
    
     Screen.MousePointer = vbHourglass
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from RSSubjectIdentifier " _
    & " WHERE ClinicalTrialName = '" & sClinicalTrialName & "'"
    
    'creates recordset to contain records to be copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from RSSubjectIdentifier " _
    & " WHERE ClinicalTrialName = '" & "'"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsRSSubjectIdentifiers = New ADODB.Recordset
    rsRSSubjectIdentifiers.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyRSSubjectIdentifiers = New ADODB.Recordset
    rsCopyRSSubjectIdentifiers.Open sSQL1, MacroADODBConnection, adOpenForwardOnly, adLockPessimistic, adCmdText

    'checks if records exist
    If rsRSSubjectIdentifiers.RecordCount <= 0 Then
        'added 3/08/2001
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    'move to first record in recordset
    rsRSSubjectIdentifiers.MoveFirst
    
    'begin record insertion
    Do While Not rsRSSubjectIdentifiers.EOF And Not rsRSSubjectIdentifiers.BOF
        rsCopyRSSubjectIdentifiers.AddNew
        rsCopyRSSubjectIdentifiers.Fields(0).Value = sNewClinicalTrialName
        For i = 1 To rsRSSubjectIdentifiers.Fields.Count - 1
            rsCopyRSSubjectIdentifiers.Fields(i).Value = rsRSSubjectIdentifiers.Fields(i).Value
        Next
        rsCopyRSSubjectIdentifiers.Update
        rsRSSubjectIdentifiers.MoveNext
    Loop

    rsRSSubjectIdentifiers.Close
    Set rsRSSubjectIdentifiers = Nothing
    rsCopyRSSubjectIdentifiers.Close
    Set rsCopyRSSubjectIdentifiers = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyPatientData.CopyRSSubjectIdentifiers"

End Sub

'---------------------------------------------------------------------------------------
Public Sub CopyRSUniquenessCheck(ByVal sClinicalTrialName As String, _
                                ByVal sNewClinicalTrialName As String)
'-------------------------------------------------------------------------------------
'duplicates rows of RSUniquenessCheck under new study name
'--------------------------------------------------------------------------------------
Dim rsRSUniquenessCheck As ADODB.Recordset
Dim rsCopyRSUniquenessCheck As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer
    
    On Error GoTo ErrLabel
     
    Screen.MousePointer = vbHourglass
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from RSUniquenessCheck " _
    & " WHERE ClinicalTrialName = '" & sClinicalTrialName & "'"
    
    'creates recordset to contain records to be copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from RSUniquenessCheck " _
    & " WHERE ClinicalTrialName = '" & "'"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsRSUniquenessCheck = New ADODB.Recordset
    rsRSUniquenessCheck.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyRSUniquenessCheck = New ADODB.Recordset
    rsCopyRSUniquenessCheck.Open sSQL1, MacroADODBConnection, adOpenForwardOnly, adLockPessimistic, adCmdText

    'checks if records exist
    If rsRSUniquenessCheck.RecordCount <= 0 Then
        'added 3/08/2001
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    'move to first record in recordset
    rsRSUniquenessCheck.MoveFirst
    
    'begin record insertion
    Do While Not rsRSUniquenessCheck.EOF And Not rsRSUniquenessCheck.BOF
        rsCopyRSUniquenessCheck.AddNew
        rsCopyRSUniquenessCheck.Fields(0).Value = sNewClinicalTrialName
        For i = 1 To rsRSUniquenessCheck.Fields.Count - 1
            rsCopyRSUniquenessCheck.Fields(i).Value = rsRSUniquenessCheck.Fields(i).Value
        Next
        rsCopyRSUniquenessCheck.Update
        rsRSUniquenessCheck.MoveNext
    Loop

    rsRSUniquenessCheck.Close
    Set rsRSUniquenessCheck = Nothing
    rsCopyRSUniquenessCheck.Close
    Set rsCopyRSUniquenessCheck = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyPatientData.CopyRSUniquenessCheck"

End Sub

'--------------------------------------------------------------------------------------
Public Sub CopyMIMessages(ByVal sClinicalTrialName As String, _
                        ByVal sNewClinicalTrialName As String)
'-------------------------------------------------------------------------------------
'duplicates rows of RSUniquenessCheck under new study name
'--------------------------------------------------------------------------------------
Dim rsMIMessages As ADODB.Recordset
Dim rsCopyMIMessages As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer
Dim lMaxId As Long
Dim lMaxObjectId As Long

    On Error GoTo ErrLabel
    
     Screen.MousePointer = vbHourglass
    'get the max values for MIMessageId and MIMessageObjectId
    lMaxId = GetMaxMIMessageId
    lMaxObjectId = GetMaxMIMessageObjectId
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from MIMessage " _
    & " WHERE MIMessageTrialName = '" & sClinicalTrialName & "'"
    
    'creates recordset to contain records to be copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from MIMessage " _
    & " WHERE MIMessageID = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsMIMessages = New ADODB.Recordset
    rsMIMessages.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyMIMessages = New ADODB.Recordset
    rsCopyMIMessages.Open sSQL1, MacroADODBConnection, adOpenForwardOnly, adLockPessimistic, adCmdText

    'checks if records exist
    If rsMIMessages.RecordCount <= 0 Then
        'added 3/08/2001
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    'move to first record in recordset
    rsMIMessages.MoveFirst
    
    'begin record insertion
    Do While Not rsMIMessages.EOF And Not rsMIMessages.BOF
        rsCopyMIMessages.AddNew
        For i = 0 To rsMIMessages.Fields.Count - 1
            'ZA 19/06/2002 - to ensure MIMessageID and MIMessageObjectID are different from values
            ' already in the table, the previous max values are added.
            'New clinical trial name is added in MIMIssageTrialName
            Select Case lCase(rsMIMessages.Fields(i).Name)
                Case "mimessageid"
                    'make new IDs unique by adding the current max to existing values
                    rsCopyMIMessages.Fields(i).Value = rsMIMessages.Fields(i).Value + lMaxId
                Case "mimessageobjectid"
                    'make new IDs unique by adding the current max to existing values
                    rsCopyMIMessages.Fields(i).Value = rsMIMessages.Fields(i).Value + lMaxObjectId
                Case "mimessagetrialname"
                    'add new trial name
                    rsCopyMIMessages.Fields(i).Value = sNewClinicalTrialName
                Case Else
                    rsCopyMIMessages.Fields(i).Value = rsMIMessages.Fields(i).Value
            End Select
            
        Next
        rsCopyMIMessages.Update
        rsMIMessages.MoveNext
    Loop

    rsMIMessages.Close
    Set rsMIMessages = Nothing
    rsCopyMIMessages.Close
    Set rsCopyMIMessages = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyPatientData.CopyMIMessages"

End Sub

'-------------------------------------------------------------------------------
Private Function GetMaxMIMessageId() As Long
'-------------------------------------------------------------------------------
' Get the largest value or MIMessageID from MIMessage table
'-------------------------------------------------------------------------------
Dim sSQL As String
Dim rs As ADODB.Recordset

    sSQL = "SELECT MAX(MIMessageID) FROM MIMessage"
    Set rs = New ADODB.Recordset
    rs.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic
    
    GetMaxMIMessageId = Val(RemoveNull(rs.Fields(0).Value))
    rs.Close
    Set rs = Nothing
    
    
End Function

'-------------------------------------------------------------------------------
Private Function GetMaxMIMessageObjectId() As Long
'-------------------------------------------------------------------------------
'Get the largest value of MIMessageObjectID from MIMessage table
'-------------------------------------------------------------------------------
Dim sSQL As String
Dim rs As ADODB.Recordset

    sSQL = "SELECT MAX(MIMessageObjectID) FROM MIMessage"
    Set rs = New ADODB.Recordset
    rs.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic
    GetMaxMIMessageObjectId = Val(RemoveNull(rs.Fields(0).Value))
    rs.Close
    Set rs = Nothing
    
    
End Function

'--------------------------------------------------------------------------------------
Public Sub CopyQGroupInstance(ByVal lOldClinicalTrialId As Long, _
                              ByVal lNewClinicalTrialId As Long)
'--------------------------------------------------------------------------------------
' REM 12/12/01
' duplicates existing QGroupInstance rows with old clinical trial ID under the new ID
'--------------------------------------------------------------------------------------
Dim rsQGroupInstance As ADODB.Recordset
Dim rsCopyQGroupInstance As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer

    On Error GoTo ErrLabel

    'creates recordset to contain records to be copied
    sSQL = "Select * from QGroupInstance " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
     'creates receiving recordset
    sSQL1 = "Select * from QGroupInstance " _
    & " WHERE ClinicalTrialId = -1"
    
    'setting and initialising recordset
    Set rsQGroupInstance = New ADODB.Recordset
    rsQGroupInstance.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyQGroupInstance = New ADODB.Recordset
    rsCopyQGroupInstance.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checks if records exist
    If rsQGroupInstance.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsQGroupInstance.MoveFirst
     
    'begin record insertion
     For j = 1 To rsQGroupInstance.RecordCount
        rsCopyQGroupInstance.AddNew
        rsCopyQGroupInstance.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsQGroupInstance.Fields.Count - 1
                rsCopyQGroupInstance.Fields(i).Value = rsQGroupInstance.Fields(i).Value
            Next
        rsCopyQGroupInstance.Update
        rsQGroupInstance.MoveNext
    Next j

    rsQGroupInstance.Close
    Set rsQGroupInstance = Nothing
    rsCopyQGroupInstance.Close
    Set rsCopyQGroupInstance = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyQGroupInstance"
End Sub


