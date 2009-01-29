Attribute VB_Name = "modCopyStudy"
'-------------------------------------------------------------------------------------------
' File:         modCopyStudy.bas
' Copyright:    InferMed Ltd. 2001-2006. All Rights Reserved
' Author:       Ashitei Trebi-Ollennu, July 2001
' Purpose:      Contains  routines for renaming or copying New Study Definitions
'-----------------------------------------------------------------------------------------
'   Revisions: 08/08/2001 ATO Added CopyTrilasites to module after consultation with NCJ
'   15/10/2001 Added transaction processing
'   12/11/01 ZA - changed nErrNum from integer type to long to accomodate long error types
'   03/07/02 REM - added 3 new Copytables for RQG's
'   NCJ 11 Feb 03 - Ensure we have the ALM running during CopyStudy
'   NCJ 30 Apr 03 - When copying a study set its status to 'In Preparation'
'   NCJ 19 May 03 - Fixed the 'In preparation' bug properly this time!
'   NCJ 21 Jun 06 - Return new study name in CopyStudy
'-----------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------
Public Function CopyStudy(ByVal lClinicalTrialId As Long, _
                  ByVal sClinicalTrialName As String) As String
'---------------------------------------------------------------------
'Duplicates a selected study under new trial ID
' NCJ 11 Feb 03 - Make sure we start the ALM before doing this
' NCJ 21 Jun 06 - Return new study name
'---------------------------------------------------------------------
Dim rsTrialStatus As ADODB.Recordset
Dim sSQL As String
Dim sNewClinicalTrialName As String
Dim lNewTrialId As Long
Dim lErrNum As Long
Dim sErrDesc As String
Dim bCopySubjectData As Boolean

    On Error GoTo ErrLabel
    
    CopyStudy = ""
    
    'checks and validates new study name
    Do
        sNewClinicalTrialName = InputBox(" New study name: ", gsDIALOG_TITLE, sClinicalTrialName)
        If sNewClinicalTrialName = "" Then
            Exit Function
        End If
        
    Loop Until ValidateTrialName(sNewClinicalTrialName)
    
    CopyStudy = sNewClinicalTrialName

    'TODO TA 10/12/2002: Copying patient data does not work - the proforma contains the study name
    bCopySubjectData = False '(DialogQuestion("Do you want to copy subject data" & "?") = vbYes)
    
    HourglassOn
    
    ' NCJ 11 Feb 03 - Must ensure we have AREZZO running
    Call StartUpCLM(lClinicalTrialId)
    
    'calculates new trial ID for study to
    'be copied since ID is not Autonumber
    lNewTrialId = NewClinicalTrialID 'From modSDTrialData
   'calls to various routines to do with the copying of record from selected tables
    
    'Begin transaction
    TransBegin
    Call CopyClinicalTrial(lClinicalTrialId, lNewTrialId, sClinicalTrialName, sNewClinicalTrialName)
    Call CopyDataItemValidations(lClinicalTrialId, lNewTrialId)
    Call CopyDataItems(lClinicalTrialId, lNewTrialId)
    Call CopyCRFElements(lClinicalTrialId, lNewTrialId)
    Call CopyCRFPages(lClinicalTrialId, lNewTrialId)
    Call CopyReasonForChanges(lClinicalTrialId, lNewTrialId)
    Call CopyStudyDefinition(lClinicalTrialId, lNewTrialId)
    Call CopyStudyDocuments(lClinicalTrialId, lNewTrialId)
    Call CopyStudyReports(lClinicalTrialId, lNewTrialId)
    Call CopyStudyReportDatas(lClinicalTrialId, lNewTrialId)
    Call CopyStudyVisits(lClinicalTrialId, lNewTrialId)
    Call CopyStudyVisitCRFPages(lClinicalTrialId, lNewTrialId)
    Call CopySubjectNumberings(lClinicalTrialId, lNewTrialId)
    Call CopySubjectEligibility(lClinicalTrialId, lNewTrialId)
    Call CopySubjectUniquenessCheck(lClinicalTrialId, lNewTrialId)
    Call CopyTrialStatusHistorys(lClinicalTrialId, lNewTrialId)
    Call CopyValueDatas(lClinicalTrialId, lNewTrialId)
    Call CopyProformaTrial(sClinicalTrialName, sNewClinicalTrialName)
    'REM 12/12/01 - added Copy for three new tables for RQG's
    Call CopyQGroup(lClinicalTrialId, lNewTrialId)
    Call CopyQGroupQuestion(lClinicalTrialId, lNewTrialId)
    Call CopyEFormQGroup(lClinicalTrialId, lNewTrialId)
    
    If bCopySubjectData Then
        Call CopyDataItemResponses(lClinicalTrialId, lNewTrialId)
        Call CopyDataItemResponseHistorys(lClinicalTrialId, lNewTrialId)
        Call CopyCRFPageInstances(lClinicalTrialId, lNewTrialId)
        Call CopyTrialSubjects(lClinicalTrialId, lNewTrialId)
        Call CopyVisitInstances(lClinicalTrialId, lNewTrialId)
        Call CopyRSNextNumbers(sClinicalTrialName, sNewClinicalTrialName)
        Call CopyRSSubjectIdentifiers(sClinicalTrialName, sNewClinicalTrialName)
        Call CopyRSUniquenessCheck(sClinicalTrialName, sNewClinicalTrialName)
        Call CopyMIMessages(sClinicalTrialName, sNewClinicalTrialName)
        'ASH 23/12/2002 Moved from copy study only above: Study Definition Bug 467
        Call CopyTrialSites(lClinicalTrialId, lNewTrialId)
        
        'TA 10/12/2002 ' no longer a qgroup instance table
        'Call CopyQGroupInstance(lClinicalTrialId, lNewTrialId)
    
        'informs user of successful completion of study copying
        DialogInformation "Copying of Study " & sNewClinicalTrialName & " and subject data complete"
    Else
        'informs user of successful completion of study copying
        DialogInformation "Copying of Study " & sNewClinicalTrialName & " complete"
    End If
    
    'commit transaction
    TransCommit

    Call ShutDownCLM
    HourglassOff
    
Exit Function
ErrLabel:
    'rollback transaction
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'just in case transaction not started
    On Error Resume Next
    TransRollBack
    Call ShutDownCLM
    HourglassOff
    Err.Raise lErrNum, , sErrDesc & "|" & "modCopyStudy.CopyStudy"

End Function

'-------------------------------------------------------------------
Public Sub CopyClinicalTrial(ByVal lOldClinicalTrialId As Long, _
                            ByVal lNewClinicalTrialId As Long, _
                            ByVal sClinicalTrialName As String, _
                            ByVal sNewClinicalTrialName As String)
'-----------------------------------------------------------------------
'duplicates existing trial with old clinical trial ID under the new ID
'-----------------------------------------------------------------------
Dim rsClinicalTrial As ADODB.Recordset
Dim rsCopyClinicalTrial As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer
    
    On Error GoTo ErrLabel
    
   'creates recordset to contain records to be copied
    sSQL = "Select * from ClinicalTrial " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId _
    & " And ClinicalTrialName = '" & sClinicalTrialName & "'"
    
    'creates recordset to contain records to be copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from ClinicalTrial " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsClinicalTrial = New ADODB.Recordset
    rsClinicalTrial.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyClinicalTrial = New ADODB.Recordset
    rsCopyClinicalTrial.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText

    'checks if records exist
    If rsClinicalTrial.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record in recordset
    rsClinicalTrial.MoveFirst
    
    'begin record insertion
    Do While Not rsClinicalTrial.EOF And Not rsClinicalTrial.BOF
        rsCopyClinicalTrial.AddNew
        rsCopyClinicalTrial.Fields(0) = lNewClinicalTrialId
        rsCopyClinicalTrial.Fields(1) = sNewClinicalTrialName
        For i = 2 To rsClinicalTrial.Fields.Count - 1
            rsCopyClinicalTrial.Fields(i).Value = rsClinicalTrial.Fields(i).Value
        Next
        ' NCJ 19 May 03 - Overwrite StatusId to be "In preparation"
        rsCopyClinicalTrial.Fields("StatusId").Value = eTrialStatus.InPreparation
        rsCopyClinicalTrial.Update
        rsClinicalTrial.MoveNext
    Loop

    rsClinicalTrial.Close
    Set rsClinicalTrial = Nothing
    rsCopyClinicalTrial.Close
    Set rsCopyClinicalTrial = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyClinicalTrial"

End Sub

'---------------------------------------------------------------------------------------------
Public Sub CopyCRFElements(ByVal lOldClinicalTrialId As Long, _
                            ByVal lNewClinicalTrialId As Long)
'---------------------------------------------------------------------------------------------
'duplicates existing CRFElement rows with old clinical trial ID under the new ID
'---------------------------------------------------------------------------------------------
Dim rsCRFElement As ADODB.Recordset
Dim rsCopyCRFElement As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer
    
    On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from CRFElement " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
    'creates recordset for recieving records being copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
     sSQL1 = "Select * from CRFElement " _
     & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsCRFElement = New ADODB.Recordset
    rsCRFElement.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising copy recordset
    Set rsCopyCRFElement = New ADODB.Recordset
    rsCopyCRFElement.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checking if records available
    If rsCRFElement.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsCRFElement.MoveFirst
    
    For j = 1 To rsCRFElement.RecordCount
        rsCopyCRFElement.AddNew
        rsCopyCRFElement.Fields(0) = lNewClinicalTrialId
        For i = 1 To rsCRFElement.Fields.Count - 1
            rsCopyCRFElement.Fields(i).Value = rsCRFElement.Fields(i).Value
        Next
        rsCopyCRFElement.Update
        rsCRFElement.MoveNext
    Next j
    
    rsCRFElement.Close
    Set rsCRFElement = Nothing
    rsCopyCRFElement.Close
    Set rsCopyCRFElement = Nothing


Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyCRFElements"

End Sub

'----------------------------------------------------------------------------------------
Public Sub CopyCRFPages(ByVal lOldClinicalTrialId As Long, _
                        ByVal lNewClinicalTrialId As Long)
'-----------------------------------------------------------------------------------------
'duplicates existing CRFPage rows with old clinical trial ID under the new ID
'-----------------------------------------------------------------------------------------
Dim rsCRFPage As ADODB.Recordset
Dim rsCopyCRFPage As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer
    
     On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from CRFPage " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
    'creates recordset for records being copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from CRFPage " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsCRFPage = New ADODB.Recordset
    rsCRFPage.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising copy recordset
    Set rsCopyCRFPage = New ADODB.Recordset
    rsCopyCRFPage.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checking for existence of records
    If rsCRFPage.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsCRFPage.MoveFirst
    
    For j = 1 To rsCRFPage.RecordCount
        rsCopyCRFPage.AddNew
        rsCopyCRFPage.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsCRFPage.Fields.Count - 1
                rsCopyCRFPage.Fields(i).Value = rsCRFPage.Fields(i).Value
            Next
        rsCopyCRFPage.Update
        rsCRFPage.MoveNext
    Next j
    
    rsCRFPage.Close
    Set rsCRFPage = Nothing
    rsCopyCRFPage.Close
    Set rsCopyCRFPage = Nothing
    

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyCRFPages"
End Sub

'----------------------------------------------------------------------------------------
Public Sub CopyDataItems(ByVal lOldClinicalTrialId As Long, _
                        ByVal lNewClinicalTrialId As Long)
'-----------------------------------------------------------------------------------------
'duplicates existing DataItem rows with old clinical trial ID under the new ID
'-----------------------------------------------------------------------------------------
Dim rsDataItem As ADODB.Recordset
Dim rsCopyDataItem As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer

     On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from DataItem " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from DataItem " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    Set rsDataItem = New ADODB.Recordset
    rsDataItem.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set rsCopyDataItem = New ADODB.Recordset
    rsCopyDataItem.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    If rsDataItem.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsDataItem.MoveFirst
     'loop to copy records into new trialID
    For j = 1 To rsDataItem.RecordCount
        rsCopyDataItem.AddNew
        rsCopyDataItem.Fields(0) = lNewClinicalTrialId
        For i = 1 To rsDataItem.Fields.Count - 1
            rsCopyDataItem.Fields(i).Value = rsDataItem.Fields(i).Value
        Next
        rsCopyDataItem.Update
        rsDataItem.MoveNext
    Next j
    
    rsDataItem.Close
    Set rsDataItem = Nothing
    rsCopyDataItem.Close
    Set rsCopyDataItem = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyDataItems"
End Sub

'------------------------------------------------------------------------------------------
Public Sub CopyDataItemValidations(ByVal lOldClinicalTrialId As Long, _
                                    ByVal lNewClinicalTrialId As Long)
'-------------------------------------------------------------------------------------------
'duplicates existing DataItemValidation rows with old clinical trial ID under the new ID
'-------------------------------------------------------------------------------------------
Dim rsDataItemValidation As ADODB.Recordset
Dim rsCopyDataItemValidation As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer
    
     On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from DataItemValidation " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
    'creates recieving recordset
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from DataItemValidation " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
   
    'setting and initialising recordset
    Set rsDataItemValidation = New ADODB.Recordset
    rsDataItemValidation.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyDataItemValidation = New ADODB.Recordset
    rsCopyDataItemValidation.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checks if recorsd exist
    If rsDataItemValidation.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsDataItemValidation.MoveFirst
     
    'begin record insertion
     For j = 1 To rsDataItemValidation.RecordCount
        rsCopyDataItemValidation.AddNew
        rsCopyDataItemValidation.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsDataItemValidation.Fields.Count - 1
                rsCopyDataItemValidation.Fields(i).Value = rsDataItemValidation.Fields(i).Value
            Next
        rsCopyDataItemValidation.Update
        rsDataItemValidation.MoveNext
    Next j
    
    rsDataItemValidation.Close
    Set rsDataItemValidation = Nothing
    rsCopyDataItemValidation.Close
    Set rsCopyDataItemValidation = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyDataItemValidations"

End Sub

'--------------------------------------------------------------------------------------------
Public Sub CopyReasonForChanges(ByVal lOldClinicalTrialId As Long, _
                                ByVal lNewClinicalTrialId As Long)
'---------------------------------------------------------------------------------------------
'duplicates existing ReasonForChange rows with old clinical trial ID under the new ID
'---------------------------------------------------------------------------------------------
Dim rsReasonForChange As ADODB.Recordset
Dim rsCopyReasonForChange As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer

     On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from ReasonForChange " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
    'creates recordset to contain records to be copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from ReasonForChange " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsReasonForChange = New ADODB.Recordset
    rsReasonForChange.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyReasonForChange = New ADODB.Recordset
    rsCopyReasonForChange.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checks if records exist
    If rsReasonForChange.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsReasonForChange.MoveFirst
    
    'begin record insertion
     For j = 1 To rsReasonForChange.RecordCount
        rsCopyReasonForChange.AddNew
        rsCopyReasonForChange.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsReasonForChange.Fields.Count - 1
                rsCopyReasonForChange.Fields(i).Value = rsReasonForChange.Fields(i).Value
            Next
        rsCopyReasonForChange.Update
        rsReasonForChange.MoveNext
    Next j
     
    rsReasonForChange.Close
    Set rsReasonForChange = Nothing
    rsCopyReasonForChange.Close
    Set rsCopyReasonForChange = Nothing


Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyReasonForChanges"

End Sub

'--------------------------------------------------------------------------------------
Public Sub CopyStudyDefinition(ByVal lOldClinicalTrialId As Long, _
                                ByVal lNewClinicalTrialId As Long)
'--------------------------------------------------------------------------------------
'duplicates existing StudyDefinition with old clinical trial ID under the new ID
'--------------------------------------------------------------------------------------

Dim rsStudyDefinition As ADODB.Recordset
Dim rsCopyStudyDefinition As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer

    On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from StudyDefinition " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
     'creates recordset for receiving recordset
     'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from StudyDefinition " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    Set rsStudyDefinition = New ADODB.Recordset
    rsStudyDefinition.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set rsCopyStudyDefinition = New ADODB.Recordset
    rsCopyStudyDefinition.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
   
    
    If rsStudyDefinition.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsStudyDefinition.MoveFirst
     
    'loop to copy records into new trialID
    For j = 1 To rsStudyDefinition.RecordCount
        rsCopyStudyDefinition.AddNew
        rsCopyStudyDefinition.Fields(0) = lNewClinicalTrialId
        For i = 1 To rsStudyDefinition.Fields.Count - 1
            rsCopyStudyDefinition.Fields(i).Value = rsStudyDefinition.Fields(i).Value
        Next
        rsCopyStudyDefinition.Update
        rsStudyDefinition.MoveNext
    Next j
    
    rsStudyDefinition.Close
    Set rsStudyDefinition = Nothing
    rsCopyStudyDefinition.Close
    Set rsCopyStudyDefinition = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyStudyDefinition"

End Sub

'--------------------------------------------------------------------------------------
Public Sub CopyStudyDocuments(ByVal lOldClinicalTrialId As Long, _
                                ByVal lNewClinicalTrialId As Long)
'--------------------------------------------------------------------------------------
'duplicates existing StudyDocument rows with old clinical trial ID under the new ID
'---------------------------------------------------------------------------------------

Dim rsStudyDocument As ADODB.Recordset
Dim rsCopyStudyDocument As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer
    
     On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from StudyDocument " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
     'creates receiving recordset
     'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from StudyDocument " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsStudyDocument = New ADODB.Recordset
    rsStudyDocument.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyStudyDocument = New ADODB.Recordset
    rsCopyStudyDocument.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checks if records exist
    If rsStudyDocument.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsStudyDocument.MoveFirst
     
    'begin record insertion
     For j = 1 To rsStudyDocument.RecordCount
        rsCopyStudyDocument.AddNew
        rsCopyStudyDocument.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsStudyDocument.Fields.Count - 1
                rsCopyStudyDocument.Fields(i).Value = rsStudyDocument.Fields(i).Value
            Next
        rsCopyStudyDocument.Update
        rsStudyDocument.MoveNext
    Next j

    rsStudyDocument.Close
    Set rsStudyDocument = Nothing
    rsCopyStudyDocument.Close
    Set rsCopyStudyDocument = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyStudyDocuments"

End Sub

'----------------------------------------------------------------------------------
Public Sub CopyStudyReports(ByVal lOldClinicalTrialId As Long, _
                            ByVal lNewClinicalTrialId As Long)
'----------------------------------------------------------------------------------
'duplicates existing StudyReport rows with old clinical trial ID under the new ID
'-----------------------------------------------------------------------------------
Dim rsStudyReport As ADODB.Recordset
Dim rsCopyStudyReport As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer

     On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from StudyReport " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
    'creates recordset to contain records to be copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from StudyReport " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsStudyReport = New ADODB.Recordset
    rsStudyReport.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyStudyReport = New ADODB.Recordset
    rsCopyStudyReport.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checks if records exist
    If rsStudyReport.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsStudyReport.MoveFirst
    
    'begin record insertion
     For j = 1 To rsStudyReport.RecordCount
        rsCopyStudyReport.AddNew
        rsCopyStudyReport.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsStudyReport.Fields.Count - 1
                rsCopyStudyReport.Fields(i).Value = rsStudyReport.Fields(i).Value
            Next
        rsCopyStudyReport.Update
        rsStudyReport.MoveNext
    Next j
     
    rsStudyReport.Close
    Set rsStudyReport = Nothing
    rsCopyStudyReport.Close
    Set rsCopyStudyReport = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyStudyReports"

End Sub

'--------------------------------------------------------------------------------------------
Public Sub CopyStudyReportDatas(ByVal lOldClinicalTrialId As Long, _
                                ByVal lNewClinicalTrialId As Long)
'---------------------------------------------------------------------------------------------
'duplicates existing StudyReportData rows with old clinical trial ID under the new ID
'---------------------------------------------------------------------------------------------
Dim rsStudyReportData As ADODB.Recordset
Dim rsCopyStudyReportData As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer

     On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from StudyReportData " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
    'creates receiving recordset
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from StudyReportData " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsStudyReportData = New ADODB.Recordset
    rsStudyReportData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyStudyReportData = New ADODB.Recordset
    rsCopyStudyReportData.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
   
    'checks if records exist
    If rsStudyReportData.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsStudyReportData.MoveFirst
    
    'begin record insertion
     For j = 1 To rsStudyReportData.RecordCount
        rsCopyStudyReportData.AddNew
        rsCopyStudyReportData.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsStudyReportData.Fields.Count - 1
                rsCopyStudyReportData.Fields(i).Value = rsStudyReportData.Fields(i).Value
            Next
        rsCopyStudyReportData.Update
        rsStudyReportData.MoveNext
    Next j
     
    rsStudyReportData.Close
    Set rsStudyReportData = Nothing
    rsCopyStudyReportData.Close
    Set rsCopyStudyReportData = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyStudyReportDatas"

End Sub

'-------------------------------------------------------------------
Public Sub CopyStudyVisits(ByVal lOldClinicalTrialId As Long, _
                            ByVal lNewClinicalTrialId As Long)
'-------------------------------------------------------------------
'duplicates existing StudyVisit rows with old clinical trial ID under the new ID
'--------------------------------------------------------------------
Dim rsStudyVisit As ADODB.Recordset
Dim rsCopyStudyVisit As ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim sSQL As String
Dim sSQL1 As String

     On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from StudyVisit " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
    'creates recordset to contain records to be copied
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from StudyVisit " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsStudyVisit = New ADODB.Recordset
    rsStudyVisit.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyStudyVisit = New ADODB.Recordset
    rsCopyStudyVisit.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
   
    'checks if records exist
    If rsStudyVisit.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsStudyVisit.MoveFirst
    
    'begin record insertion
     For j = 1 To rsStudyVisit.RecordCount
        rsCopyStudyVisit.AddNew
        rsCopyStudyVisit.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsStudyVisit.Fields.Count - 1
                rsCopyStudyVisit.Fields(i).Value = rsStudyVisit.Fields(i).Value
            Next
        rsCopyStudyVisit.Update
        rsStudyVisit.MoveNext
    Next j

    rsStudyVisit.Close
    Set rsStudyVisit = Nothing
    rsCopyStudyVisit.Close
    Set rsCopyStudyVisit = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyStudyVisits"

End Sub

'-----------------------------------------------------------------------------------------
Public Sub CopyStudyVisitCRFPages(ByVal lOldClinicalTrialId As Long, _
                                ByVal lNewClinicalTrialId As Long)
'------------------------------------------------------------------------------------------
'duplicates existing StudyVisitCRFPage rows with old clinical trial ID under the new ID
'------------------------------------------------------------------------------------------
Dim rsStudyVisitCRFPage As ADODB.Recordset
Dim rsCopyStudyVisitCRFPage As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer

     On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from StudyVisitCRFPage " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
     'creates receiving recordset
     'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from StudyVisitCRFPage " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsStudyVisitCRFPage = New ADODB.Recordset
    rsStudyVisitCRFPage.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising receiving recordset
    Set rsCopyStudyVisitCRFPage = New ADODB.Recordset
    rsCopyStudyVisitCRFPage.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checks if records exist
    If rsStudyVisitCRFPage.RecordCount <= 0 Then
        Exit Sub
    End If
   
   'move to first record inrecordset
    rsStudyVisitCRFPage.MoveFirst
    'begin record insertion
     For j = 1 To rsStudyVisitCRFPage.RecordCount
        rsCopyStudyVisitCRFPage.AddNew
        rsCopyStudyVisitCRFPage.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsStudyVisitCRFPage.Fields.Count - 1
                rsCopyStudyVisitCRFPage.Fields(i).Value = rsStudyVisitCRFPage.Fields(i).Value
            Next
        rsCopyStudyVisitCRFPage.Update
        rsStudyVisitCRFPage.MoveNext
    Next j
    
    rsStudyVisitCRFPage.Close
    Set rsStudyVisitCRFPage = Nothing
    rsCopyStudyVisitCRFPage.Close
    Set rsCopyStudyVisitCRFPage = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyStudyVisitCRFPages"

End Sub

'-----------------------------------------------------------------------
Public Sub CopySubjectNumberings(ByVal lOldClinicalTrialId As Long, _
                                ByVal lNewClinicalTrialId As Long)
'-----------------------------------------------------------------------
'duplicates existing SubjectNumbering rows with old clinical trial ID under the new ID
'-----------------------------------------------------------------------
Dim rsSubjectNumbering As ADODB.Recordset
Dim rsCopySubjectNumbering As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer

     On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from SubjectNumbering " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
    'creates receiving recordset
    'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from SubjectNumbering " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsSubjectNumbering = New ADODB.Recordset
    rsSubjectNumbering.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
   
    'setting and initialising receiving recordset
    Set rsCopySubjectNumbering = New ADODB.Recordset
    rsCopySubjectNumbering.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
   
   'checks if records exist
    If rsSubjectNumbering.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsSubjectNumbering.MoveFirst
    
    'begin record insertion
     For j = 1 To rsSubjectNumbering.RecordCount
        rsCopySubjectNumbering.AddNew
        rsCopySubjectNumbering.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsSubjectNumbering.Fields.Count - 1
                rsCopySubjectNumbering.Fields(i).Value = rsSubjectNumbering.Fields(i).Value
            Next
        rsCopySubjectNumbering.Update
        rsSubjectNumbering.MoveNext
    Next j
     
    rsSubjectNumbering.Close
    Set rsSubjectNumbering = Nothing
    rsCopySubjectNumbering.Close
    Set rsCopySubjectNumbering = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopySubjectNumberings"


End Sub

'---------------------------------------------------------------------------
Private Sub CopyTrialStatusHistorys(ByVal lOldClinicalTrialId As Long, _
                                    ByVal lNewClinicalTrialId As Long)
'---------------------------------------------------------------------------
' duplicates existing TrialStatusHistory rows with old clinical trial ID under the new ID
' NCJ 30 Apr 03 - Do not copy old history but set it as a "new" study (Bug 1649)
'---------------------------------------------------------------------------
Dim sSQL As String
Dim oTimezone As TimeZone

    On Error GoTo ErrLabel
    
    Set oTimezone = New TimeZone
    
    sSQL = "INSERT INTO TrialStatusHistory ( ClinicalTrialId, VersionId, " _
        & " TrialStatusChangeId, StatusId, UserName, " _
        & " StatusChangedTimestamp, StatusChangedTimestamp_TZ )" _
        & " VALUES (" & lNewClinicalTrialId & ", 1," _
        & " 1,1,'" & goUser.UserName & "'," _
        & SQLStandardNow & ", " & oTimezone.TimezoneOffset & ")"
    
    MacroADODBConnection.Execute sSQL
    
    Set oTimezone = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyTrialStatusHistorys"

End Sub

'------------------------------------------------------------------
Public Sub CopyValueDatas(ByVal lOldClinicalTrialId As Long, _
                            ByVal lNewClinicalTrialId As Long)
'------------------------------------------------------------------
'duplicates existing ValueData rows with old clinical trial ID under the new ID
'------------------------------------------------------------------
Dim rsValueData As ADODB.Recordset
Dim rsCopyValueData As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer

     On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from ValueData " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
     'creates receiving recordset
     'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from ValueData " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsValueData = New ADODB.Recordset
    rsValueData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising receiving recordset
    Set rsCopyValueData = New ADODB.Recordset
    rsCopyValueData.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checks if records exist
    If rsValueData.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsValueData.MoveFirst
    
    'begin record insertion
     For j = 1 To rsValueData.RecordCount
        rsCopyValueData.AddNew
        rsCopyValueData.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsValueData.Fields.Count - 1
                rsCopyValueData.Fields(i).Value = rsValueData.Fields(i).Value
            Next
        rsCopyValueData.Update
        rsValueData.MoveNext
    Next j
     
    rsValueData.Close
    Set rsValueData = Nothing
    rsCopyValueData.Close
    Set rsCopyValueData = Nothing


Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyValueDatas"

End Sub
'------------------------------------------------------------------------
Public Sub CopyTrialSites(ByVal lOldClinicalTrialId As Long, _
                            ByVal lNewClinicalTrialId As Long)
'-------------------------------------------------------------------------
'duplicates existing TrialSites rows with old clinicaltrial ID under the new ID
'added 08/08/2001 Ash
'-------------------------------------------------------------------------
Dim rsTrialSites As ADODB.Recordset
Dim rsCopyTrialSites As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer
    
     On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "Select * from TrialSite " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
     'creates receiving recordset
     'Changed Mo Morris 31/8/01, "Where true = false" does not work in SQl Server
    sSQL1 = "Select * from TrialSite " _
    & " WHERE ClinicalTrialId = -1"
    '& " Where true = false"
    
    'setting and initialising recordset
    Set rsTrialSites = New ADODB.Recordset
    rsTrialSites.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyTrialSites = New ADODB.Recordset
    rsCopyTrialSites.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checks if records exist
    If rsTrialSites.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsTrialSites.MoveFirst
     
    'begin record insertion
     For j = 1 To rsTrialSites.RecordCount
        rsCopyTrialSites.AddNew
        rsCopyTrialSites.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsTrialSites.Fields.Count - 1
                rsCopyTrialSites.Fields(i).Value = rsTrialSites.Fields(i).Value
            Next
        rsCopyTrialSites.Update
        rsTrialSites.MoveNext
    Next j

    rsTrialSites.Close
    Set rsTrialSites = Nothing
    rsCopyTrialSites.Close
    Set rsCopyTrialSites = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyTrialSites"

End Sub

'-----------------------------------------------------------------------
Public Sub CopyQGroup(ByVal lOldClinicalTrialId As Long, _
                      ByVal lNewClinicalTrialId As Long)
'-----------------------------------------------------------------------
' REM 12/12/01
' duplicates existing QGroup rows with old clinical trial ID under the new ID
'-----------------------------------------------------------------------
Dim rsQGroups As ADODB.Recordset
Dim rsCopyQGroups As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer

    On Error GoTo ErrLabel

    'creates recordset to contain records to be copied
    sSQL = "Select * from QGroup " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
     'creates receiving recordset
    sSQL1 = "Select * from QGroup " _
    & " WHERE ClinicalTrialId = -1"
    
    'setting and initialising recordset
    Set rsQGroups = New ADODB.Recordset
    rsQGroups.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyQGroups = New ADODB.Recordset
    rsCopyQGroups.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checks if records exist
    If rsQGroups.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record in recordset
    rsQGroups.MoveFirst
     
    'begin record insertion
     For j = 1 To rsQGroups.RecordCount
        rsCopyQGroups.AddNew
        rsCopyQGroups.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsQGroups.Fields.Count - 1
                rsCopyQGroups.Fields(i).Value = rsQGroups.Fields(i).Value
            Next
        rsCopyQGroups.Update
        rsQGroups.MoveNext
    Next j

    rsQGroups.Close
    Set rsQGroups = Nothing
    rsCopyQGroups.Close
    Set rsCopyQGroups = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyQGroup"
End Sub

'-----------------------------------------------------------------------
Public Sub CopyQGroupQuestion(ByVal lOldClinicalTrialId As Long, _
                              ByVal lNewClinicalTrialId As Long)
'-----------------------------------------------------------------------
' REM 12/12/01
' duplicates existing QGroupQuestion rows with old clinical trial ID under the new ID
'-----------------------------------------------------------------------
Dim rsQGroupQuestions As ADODB.Recordset
Dim rsCopyQGroupQuestions As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer

    On Error GoTo ErrLabel

    'creates recordset to contain records to be copied
    sSQL = "Select * from QGroupQuestion " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
     'creates receiving recordset
    sSQL1 = "Select * from QGroupQuestion " _
    & " WHERE ClinicalTrialId = -1"
    
    'setting and initialising recordset
    Set rsQGroupQuestions = New ADODB.Recordset
    rsQGroupQuestions.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyQGroupQuestions = New ADODB.Recordset
    rsCopyQGroupQuestions.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checks if records exist
    If rsQGroupQuestions.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsQGroupQuestions.MoveFirst
     
    'begin record insertion
     For j = 1 To rsQGroupQuestions.RecordCount
        rsCopyQGroupQuestions.AddNew
        rsCopyQGroupQuestions.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsQGroupQuestions.Fields.Count - 1
                rsCopyQGroupQuestions.Fields(i).Value = rsQGroupQuestions.Fields(i).Value
            Next
        rsCopyQGroupQuestions.Update
        rsQGroupQuestions.MoveNext
    Next j

    rsQGroupQuestions.Close
    Set rsQGroupQuestions = Nothing
    rsCopyQGroupQuestions.Close
    Set rsCopyQGroupQuestions = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyQGroupQuestion"
End Sub

'-----------------------------------------------------------------------
Public Sub CopyEFormQGroup(ByVal lOldClinicalTrialId As Long, _
                           ByVal lNewClinicalTrialId As Long)
'-----------------------------------------------------------------------
' REM 12/12/01
' duplicates existing EFormQGroup rows with old clinical trial ID under the new ID
'-----------------------------------------------------------------------
Dim rsEFormQGroup As ADODB.Recordset
Dim rsCopyEFormQGroup As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Integer
Dim j As Integer

    On Error GoTo ErrLabel

    'creates recordset to contain records to be copied
    sSQL = "Select * from EFormQGroup " _
    & " Where ClinicalTrialID = " & lOldClinicalTrialId
    
     'creates receiving recordset
    sSQL1 = "Select * from EFormQGroup " _
    & " WHERE ClinicalTrialId = -1"
    
    'setting and initialising recordset
    Set rsEFormQGroup = New ADODB.Recordset
    rsEFormQGroup.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyEFormQGroup = New ADODB.Recordset
    rsCopyEFormQGroup.Open sSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checks if records exist
    If rsEFormQGroup.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record inrecordset
    rsEFormQGroup.MoveFirst
     
    'begin record insertion
     For j = 1 To rsEFormQGroup.RecordCount
        rsCopyEFormQGroup.AddNew
        rsCopyEFormQGroup.Fields(0) = lNewClinicalTrialId
            For i = 1 To rsEFormQGroup.Fields.Count - 1
                rsCopyEFormQGroup.Fields(i).Value = rsEFormQGroup.Fields(i).Value
            Next
        rsCopyEFormQGroup.Update
        rsEFormQGroup.MoveNext
    Next j

    rsEFormQGroup.Close
    Set rsEFormQGroup = Nothing
    rsCopyEFormQGroup.Close
    Set rsCopyEFormQGroup = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopyEFormQGroup"
End Sub

'-----------------------------------------------------------------------
Public Sub CopySubjectUniquenessCheck(ByVal lOldClinicalTrialId As Long, _
                                ByVal lNewClinicalTrialId As Long)
'-----------------------------------------------------------------------
'duplicates existing SubjectUniquenessCheck rows with old clinical trial ID under the new ID
'-----------------------------------------------------------------------

Dim sSQL As String

     On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "INSERT INTO UniquenessCheck " _
    & " SELECT " & lNewClinicalTrialId & " AS ClinicalTrialId," _
    & " VersionId, CheckCode, Expression" _
    & " FROM UniquenessCheck " _
    & " WHERE ClinicalTrialId = " & lOldClinicalTrialId & "" _
    & " AND VersionId = 1"
               
    MacroADODBConnection.Execute sSQL
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopySubjectUniquenessCheck"


End Sub


'-----------------------------------------------------------------------
Public Sub CopySubjectEligibility(ByVal lOldClinicalTrialId As Long, _
                                ByVal lNewClinicalTrialId As Long)
'-----------------------------------------------------------------------
'duplicates existing SubjectEligibility rows with old clinical trial ID under the new ID
'-----------------------------------------------------------------------

Dim sSQL As String

     On Error GoTo ErrLabel
    
    'creates recordset to contain records to be copied
    sSQL = "INSERT INTO Eligibility " _
    & " SELECT " & lNewClinicalTrialId & " AS ClinicalTrialId," _
    & " VersionId, EligibilityCode, RandomisationCode, Flag, Condition" _
    & " FROM Eligibility " _
    & " WHERE ClinicalTrialId = " & lOldClinicalTrialId & "" _
    & " AND VersionId = 1"
               
    MacroADODBConnection.Execute sSQL
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modCopyStudy.CopySubjectEligibility"

End Sub

