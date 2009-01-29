Attribute VB_Name = "modDiagnostic"

'---------------------------------------------------------------------------------------
' File: modDiagnostic.bas
' Copyright:    InferMed Ltd. 2006-2008. All Rights Reserved
' Author:   Nicky Johns, InferMed, 21 February 2006
' Purpose:      Contains the main routines of the MACRO Diagnostic application
'----------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 5 Jun 07 - Added count of non-empty responses
'   NCJ 18 Jun 07 - Trap errors when loading PLM files
'   NCJ 21 Jun 07 - Check on eForms in visits
'   NCJ 25-28 Jun 07 - Keep track of things already reported to avoid duplicate reports
'   NCJ 18 Feb 08 - Must also pass Site and PersonId to ReportOnCRF
'
'----------------------------------------------------------------------------------------'

Private msLogFile As String
Public Const gsPLM_ERROR As String = "Error while loading PLM file for this subject"

' Remember what's been deleted
Private mcolDeletedVisits As Collection
Private mcolDeletedQuestions As Collection
Private mcolDelFromEForm As Collection  ' Questions removed from eForms
Private mcolDelFromVisit As Collection  ' eForms removed from visits

' Count "invalid" responses
Private mlInvalidResponses As Long

Option Explicit

'---------------------------------------------------------------------------------------
Public Function PatIntegrityReport(lTrialId As Long, sSite As String, lPersonId As Long) As String
'---------------------------------------------------------------------------------------
' Produce report on any deleted objects in this patient's data
'---------------------------------------------------------------------------------------
Dim sReport As String

    On Error GoTo ErrLabel
    
    ' Initialise collections of deleted things
    Set mcolDeletedQuestions = New Collection
    Set mcolDelFromEForm = New Collection
    Set mcolDelFromVisit = New Collection
    Set mcolDeletedVisits = New Collection
    
    ' Initialise number of invalid responses
    mlInvalidResponses = 0
    
    sReport = AllCounts(lTrialId, sSite, lPersonId)
    If sReport > "" Then
        sReport = sReport & vbCrLf & ReportOnVIs(lTrialId, sSite, lPersonId)
        sReport = sReport & vbCrLf & ReportOnEFIs(lTrialId, sSite, lPersonId)
        sReport = sReport & vbCrLf & ReportOnResponses(lTrialId, sSite, lPersonId)
    Else
        sReport = "Nothing found for this subject"
    End If
    
    ' Tidy up collections
    Set mcolDeletedQuestions = Nothing
    Set mcolDelFromEForm = Nothing
    Set mcolDelFromVisit = Nothing
    Set mcolDeletedVisits = Nothing
    
    PatIntegrityReport = sReport
    
Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.PatIntegrityReport"
End Function

'---------------------------------------------------------------------------------------
Private Function AllCounts(lTrialId As Long, sSite As String, lPersonId As Long) As String
'---------------------------------------------------------------------------------------
' Count all the Visit Instances, CRFPageInstances and DataItemReponses for this subject
'---------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsInsts As ADODB.Recordset
Dim sReport As String
Dim lCount As Long

    sReport = ""
    
    sSQL = "SELECT VisitTaskId from VisitInstance"
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lTrialId
    sSQL = sSQL & " AND TrialSite = '" & sSite & "'"
    sSQL = sSQL & " AND PersonId = " & lPersonId
  
    Set rsInsts = New ADODB.Recordset
    rsInsts.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' Anything there?
    If rsInsts.RecordCount > 0 Then
        sReport = rsInsts.RecordCount & " Visit instances" & vbCrLf
        
        sSQL = "SELECT CRFPageTaskId from CRFPageInstance"
        sSQL = sSQL & " WHERE ClinicalTrialId = " & lTrialId
        sSQL = sSQL & " AND TrialSite = '" & sSite & "'"
        sSQL = sSQL & " AND PersonId = " & lPersonId
      
        Set rsInsts = New ADODB.Recordset
        rsInsts.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
        sReport = sReport & rsInsts.RecordCount & " EForm instances" & vbCrLf
    
        sSQL = "SELECT ResponseTaskId from DataItemResponse"
        sSQL = sSQL & " WHERE ClinicalTrialId = " & lTrialId
        sSQL = sSQL & " AND TrialSite = '" & sSite & "'"
        sSQL = sSQL & " AND PersonId = " & lPersonId
      
        Set rsInsts = New ADODB.Recordset
        rsInsts.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
        sReport = sReport & rsInsts.RecordCount & " Response values" & vbCrLf
    
        rsInsts.Close
        Set rsInsts = Nothing
    End If

    AllCounts = sReport
    
Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.AllCounts"
End Function

'---------------------------------------------------------------------------------------
Private Function ReportOnVIs(lTrialId As Long, sSite As String, lPersonId As Long) As String
'---------------------------------------------------------------------------------------
' Return report of any visits present in VisitInstance table but not in StudyVisit table
'---------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsInsts As ADODB.Recordset
Dim sReport As String
Dim lCount As Long

    On Error GoTo ErrLabel
    
    sReport = ""
    
    sSQL = "SELECT DISTINCT VisitID from VisitInstance"
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lTrialId
    sSQL = sSQL & " AND TrialSite = '" & sSite & "'"
    sSQL = sSQL & " AND PersonId = " & lPersonId
  
    Set rsInsts = New ADODB.Recordset
    rsInsts.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    lCount = rsInsts.RecordCount
    
    If Not rsInsts.EOF Then rsInsts.MoveFirst
    
    Do While Not rsInsts.EOF
        sReport = sReport & ReportOnVisit(lTrialId, rsInsts.Fields("VisitId"))
        rsInsts.MoveNext
    Loop
    
    rsInsts.Close
    Set rsInsts = Nothing

    If sReport = "" Then
        sReport = "All Visit instances OK"
    End If
    ReportOnVIs = "--- VISITS ---" & vbCrLf _
                    & sReport & " (" & lCount & " visits)" & vbCrLf
    
Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.ReportOnVIs"
End Function

'---------------------------------------------------------------------------------------
Private Function ReportOnEFIs(lTrialId As Long, sSite As String, lPersonId As Long) As String
'---------------------------------------------------------------------------------------
' Return report of any eForms present in CRFPageInstance table but not in CRFPage table
'---------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsInsts As ADODB.Recordset
Dim sReport As String
Dim lCount As Long
Dim colCRFsDone As Collection

    On Error GoTo ErrLabel
    
    sReport = ""
    Set colCRFsDone = New Collection
    
    sSQL = "SELECT DISTINCT CRFPageID from CRFPageInstance"
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lTrialId
    sSQL = sSQL & " AND TrialSite = '" & sSite & "'"
    sSQL = sSQL & " AND PersonId = " & lPersonId
  
    Set rsInsts = New ADODB.Recordset
    rsInsts.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    lCount = rsInsts.RecordCount
    
    If Not rsInsts.EOF Then rsInsts.MoveFirst
    
    Do While Not rsInsts.EOF
        ' NCJ 18 Feb 08 - Must also pass Site and PersonId to ReportOnCRF
        sReport = sReport & ReportOnCRF(lTrialId, sSite, lPersonId, rsInsts.Fields("CRFPageId"))
        rsInsts.MoveNext
    Loop
    
    rsInsts.Close
    Set rsInsts = Nothing

    If sReport = "" Then
        sReport = "All EForm instances OK"
    End If
    ReportOnEFIs = "--- EFORMS ---" & vbCrLf & sReport _
                    & " (" & lCount & " eForms)" & vbCrLf
    
Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.ReportOnEFIs"
End Function

'------------------------------------------------------------------------
Private Function ReportOnResponses(lTrialId As Long, sSite As String, lPersonId As Long) As String
'------------------------------------------------------------------------
' Return report on whether the questions and CRFElements still exist for all responses for this patient
' NCJ 5 Jun 07 - Count non-empty responses separately
'------------------------------------------------------------------------
Dim sSQL As String
Dim rsInsts As ADODB.Recordset
Dim sReport As String
Dim lCount As Long
Dim lNonEmptyCount As Long

    On Error GoTo ErrLabel
    
    sReport = ""
    
    sSQL = "SELECT ResponseTaskId, CRFPageId, VisitId, CRFElementId, DataItemID, RESPONSEVALUE, RESPONSETIMESTAMP from DataItemResponse"
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lTrialId
    sSQL = sSQL & " AND TrialSite = '" & sSite & "'"
    sSQL = sSQL & " AND PersonId = " & lPersonId
  
    Set rsInsts = New ADODB.Recordset
    rsInsts.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' No. of responses
    lCount = rsInsts.RecordCount
    ' No. of non-empty responses
    lNonEmptyCount = 0
    
    If Not rsInsts.EOF Then rsInsts.MoveFirst
    
    Do While Not rsInsts.EOF
        If CollectionMember(mcolDeletedVisits, "K" & rsInsts.Fields("VisitId"), False) _
                Or CollectionMember(mcolDelFromVisit, PageDelFromVisitKey(rsInsts.Fields("CRFPageId"), rsInsts.Fields("VisitId")), False) Then
            ' It's in a deleted visit or on an eForm removed from a visit so ignore
            ' but update the baddie count
            mlInvalidResponses = mlInvalidResponses + 1
        Else
            sReport = sReport & ReportOnQuestion(lTrialId, rsInsts.Fields("DataItemId"), rsInsts.Fields("ResponseTaskId"), _
                        rsInsts.Fields("CRFPageId"), rsInsts.Fields("CRFElementId"), _
                        RemoveNull(rsInsts.Fields("RESPONSEVALUE")), rsInsts.Fields("RESPONSETIMESTAMP"))
        End If
        If RemoveNull(rsInsts.Fields("RESPONSEVALUE")) <> "" Then
            lNonEmptyCount = lNonEmptyCount + 1
        End If
        rsInsts.MoveNext
    Loop
    
    rsInsts.Close
    Set rsInsts = Nothing

    ' Check for error report and any invalid responses
    If sReport = "" And mlInvalidResponses = 0 Then
        sReport = "All Responses OK"
    End If
    sReport = "--- QUESTIONS ---" & vbCrLf & sReport _
                    & " (" & lCount & " response values, " & lNonEmptyCount & " non-empty responses)" & vbCrLf
    If mlInvalidResponses > 0 Then
        sReport = sReport & "*** " & mlInvalidResponses & " responses with invalid references"
    End If
    
    ReportOnResponses = sReport
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.ReportOnResponses"
End Function

'------------------------------------------------------------------------
Private Function ReportOnCRF(lTrialId As Long, sSite As String, lPersonId As Long, lPageId As Long) As String
'------------------------------------------------------------------------
' Return message if this eForm doesn't exist,
' or if it doesn't exist in its visit
' Return "" if it's OK
' NCJ 18 Feb 08 - Must use Site and PersonId in CRFPageInstance table
'------------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sReport As String
Dim sPageCode As String

    On Error GoTo ErrLabel
    
    sReport = ""
    
    ' First check the eForm exists
    sSQL = "SELECT CRFPageCode from CRFPage"
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lTrialId
    sSQL = sSQL & " AND CRFPageId = " & lPageId
  
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsTemp.EOF Then
        rsTemp.Close
        sReport = "* eForm " & lPageId & " does not exist" & vbCrLf
    Else
        sPageCode = rsTemp.Fields("CRFPageCode")
        rsTemp.Close
        ' Now check that it exists in all the visits
        sSQL = "SELECT DISTINCT VisitId from CRFPageInstance"
        sSQL = sSQL & " WHERE ClinicalTrialId = " & lTrialId
        sSQL = sSQL & " AND TrialSite = '" & sSite & "'"
        sSQL = sSQL & " AND PersonId = " & lPersonId
        sSQL = sSQL & " AND CRFPageId = " & lPageId
                
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                  
        Do While Not rsTemp.EOF
            sReport = sReport & ReportOnCRFInVisit(lTrialId, lPageId, sPageCode, rsTemp.Fields("VisitId"))
            rsTemp.MoveNext
        Loop
        rsTemp.Close
    End If
    
    Set rsTemp = Nothing
    ReportOnCRF = sReport
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.ReportOnCRF"

End Function

'------------------------------------------------------------------------
Private Function ReportOnCRFInVisit(lTrialId As Long, lPageId As Long, sPageCode As String, lVisitId As Long)
'------------------------------------------------------------------------
' Return message if this eForm doesn't exist in this visit
' Assume eForm exists
' Return "" if it's OK
'------------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sReport As String
Dim sVisit As String
Dim sPageDelKey As String

    On Error GoTo ErrLabel
    
    sReport = ""
    
    sSQL = "SELECT CRFPageId, VisitId from StudyVisitCRFPage"
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lTrialId
    sSQL = sSQL & " AND CRFPageId = " & lPageId
    sSQL = sSQL & " AND VisitId = " & lVisitId
  
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsTemp.EOF Then
        ' Does this visit actually exist? (If not, we've dealt with it already)
        sVisit = VisitCodeFromId(lTrialId, lVisitId)
        If sVisit <> "" Then
            ' eForm has been removed from visit
            sReport = "* eForm " & lPageId & " (" & sPageCode & ") does not exist in visit " & lVisitId
            sReport = sReport & " (" & sVisit & ")" & vbCrLf
            ' Add to collection of eForms removed from visits
            sPageDelKey = PageDelFromVisitKey(lPageId, lVisitId)
            Call CollectionAddAnyway(mcolDelFromVisit, sPageDelKey, sPageDelKey)
        End If
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    ReportOnCRFInVisit = sReport
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.ReportOnCRFInVisit"

End Function

'---------------------------------------------------------------------------------------------
Private Function PageDelFromVisitKey(lCRFPageId As Long, lVisitId As Long) As String
'---------------------------------------------------------------------------------------------
' Get the collection key for an eForm removed from a visit
'---------------------------------------------------------------------------------------------

    PageDelFromVisitKey = "K" & lCRFPageId & "-" & lVisitId

End Function

'---------------------------------------------------------------------------------------------
Private Function ReportOnVisit(lTrialId As Long, lVisitId As Long) As String
'---------------------------------------------------------------------------------------------
' Return message if this visit doesn't exist
' Return "" if it's OK
'------------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrLabel
    
    ReportOnVisit = ""
    
    sSQL = "SELECT VisitCode from StudyVisit"
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lTrialId
    sSQL = sSQL & " AND VisitId = " & lVisitId
  
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsTemp.EOF Then
        ReportOnVisit = "* Visit " & lVisitId & " does not exist" & vbCrLf
        ' Remember it as one of our deleted ones
        Call CollectionAddAnyway(mcolDeletedVisits, lVisitId, "K" & lVisitId)
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.ReportOnVisit"

End Function

'---------------------------------------------------------------------------------------------
Private Function ReportOnQuestion(lTrialId As Long, lDataItemId As Long, lResponseId As Long, _
                    lCRFPageId As Long, lCRFElementId As Long, _
                    sValue As String, dblTimeStamp As Double) As String
'---------------------------------------------------------------------------------------------
' Return message if this Question OR CRFElement doesn't exist
' Return "" if it's OK or if it's already been reported
' Update count of "Invalid responses" as appropriate
'------------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sReport As String
Dim sQCode As String
Dim sPageCode As String
Dim sNewQCode As String
Dim sDeletedQuestionKey As String
Dim sDelFromEFormKey As String

    On Error GoTo ErrLabel
    
    sReport = ""
    
    ' Keys for "deleted" collections
    sDeletedQuestionKey = "K" & lDataItemId
    sDelFromEFormKey = "K" & lDataItemId & "-" & lCRFPageId
    
    If CollectionMember(mcolDeletedQuestions, sDeletedQuestionKey, False) _
        Or CollectionMember(mcolDelFromEForm, sDelFromEFormKey, False) Then
        ' Already done this one
        ReportOnQuestion = ""
        ' Increment "baddie" count
        mlInvalidResponses = mlInvalidResponses + 1
        Exit Function
    End If
    
    ' Check to see if the DataItem is defined
    sSQL = "SELECT DataItemCode from DataItem"
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lTrialId
    sSQL = sSQL & " AND DataItemId = " & lDataItemId
  
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsTemp.EOF Then
        sReport = "* ResponseTaskID = " & lResponseId & ", DataItemID = " & lDataItemId & " - question does not exist"
        ' Add to collection of deleted questions
        Call CollectionAddAnyway(mcolDeletedQuestions, lDataItemId, sDeletedQuestionKey)
      Else
        sQCode = rsTemp.Fields("DataItemCode")
        rsTemp.Close
        ' Question still exists - is it still on the eForm?
        ' First check whether form exists
        sPageCode = CRFPageCodeFromId(lTrialId, lCRFPageId)
        If sPageCode <> "" Then
            ' The eForm still exists - check the CRFElement
            sSQL = "SELECT DataItemId from CRFElement "
            sSQL = sSQL & " WHERE CRFElement.ClinicalTrialId = " & lTrialId
            sSQL = sSQL & " AND CRFElement.CRFPageId = " & lCRFPageId
            sSQL = sSQL & " AND CRFElement.CRFElementId = " & lCRFElementId
        
            Set rsTemp = New ADODB.Recordset
            rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            If rsTemp.EOF Then
                sReport = "* ResponseTaskID = " & lResponseId & ", DataItemID = " & lDataItemId & " (" & sQCode _
                        & ") - question does not exist"
                ' Add to collection of els deleted from eForm
                Call CollectionAddAnyway(mcolDelFromEForm, sDelFromEFormKey, sDelFromEFormKey)
            ElseIf rsTemp.Fields("DataItemId") <> lDataItemId Then
                sReport = "* ResponseTaskID = " & lResponseId & ", CRFElementID = " & lCRFElementId _
                        & " - CRFElement has changed DataItemId from " & lDataItemId & " (" & sQCode _
                        & ") to " & rsTemp.Fields("DataItemId")
                ' Does new question exist?
                sNewQCode = DataItemCodeFromId(lTrialId, rsTemp.Fields("DataItemId"))
                If sNewQCode <> "" Then
                    sReport = sReport & " (" & sNewQCode & ")"
                End If
            End If
            If sReport <> "" Then
                ' Add on the eForm info
                sReport = sReport & " on eForm " & lCRFPageId & " (" & sPageCode & ")"
            End If
            rsTemp.Close
        Else
            ' The form has been deleted - already reported elsewhere,
            ' so ignore but increment "invalid response" count
            mlInvalidResponses = mlInvalidResponses + 1
        End If
    End If
    
    Set rsTemp = Nothing

    ' For an invalid response, report its value and timestamp
    If sReport > "" Then
            sReport = sReport & ", Value = " & sValue & _
                    ", Timestamp = " & Format(CDate(dblTimeStamp), "yyyy/mm/dd hh:mm:ss") _
                    & vbCrLf
            mlInvalidResponses = mlInvalidResponses + 1
    End If
    
    ReportOnQuestion = sReport
  
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.ReportOnQuestion"

End Function

'---------------------------------------------------------------------------------------------
Private Function DiagGetCLMFile(ByVal sStudyName As String) As String
'---------------------------------------------------------------------------------------------
' Get the CLM file as a string from the PROTOCOLS table for this study
'---------------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
    
    DiagGetCLMFile = ""
    sSQL = "SELECT ArezzoFile from Protocols WHERE FILENAME = '" & sStudyName & "'"
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    If Not rsTemp.EOF Then
        DiagGetCLMFile = rsTemp.Fields("ArezzoFile")
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

End Function

'---------------------------------------------------------------------------------------------
Public Function DiagSaveCLMFile(sStudyName As String) As String
'---------------------------------------------------------------------------------------------
' Save the AREZZO study def as a CLM file
'---------------------------------------------------------------------------------------------
Dim sCLMFile As String
Dim sFileName As String

    sCLMFile = DiagGetCLMFile(sStudyName)
    If sCLMFile <> "" Then
        sFileName = gsTEMP_PATH & sStudyName & ".clm"
        Call StringToFile(sFileName, sCLMFile)
        DiagSaveCLMFile = "CLM file saved as: " & sFileName & vbCrLf _
                & "Length of file: " & Len(sCLMFile)
    Else
        DiagSaveCLMFile = " *** Study not found: " & sStudyName
    End If
    
End Function

'---------------------------------------------------------------------------------------------
Private Function DiagGetPLMFile(ByVal lTrialId As Long, ByVal sSite As String, ByVal lPersonId As Long, _
            ByRef sMSG As String) As String
'---------------------------------------------------------------------------------------------
' Get the PLM file as a string from the TRIALSUBJECT table
' sMsg gives error message if not possible; sMsg = "" if all OK
'---------------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrLabel
    
    DiagGetPLMFile = ""
    sMSG = ""
    
    sSQL = "SELECT ProformaState from TrialSubject"
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lTrialId
    sSQL = sSQL & " AND TrialSite = '" & sSite & "'"
    sSQL = sSQL & " AND PersonId = " & lPersonId

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp.Fields("ProformaState")) Then
            DiagGetPLMFile = rsTemp.Fields("ProformaState")
        End If
    Else
        sMSG = "Subject not found"
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.DiagGetPLMFile"
End Function

'---------------------------------------------------------------------------------------------
Private Sub DiagUpdatePLMFile(ByVal sState As String, ByVal lTrialId As Long, ByVal sSite As String, ByVal lPersonId As Long)
'---------------------------------------------------------------------------------------------
' Update the PLM file in the TRIALSUBJECT table
'---------------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
    
    sSQL = "SELECT ProformaState from TrialSubject"
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lTrialId
    sSQL = sSQL & " AND TrialSite = '" & sSite & "'"
    sSQL = sSQL & " AND PersonId = " & lPersonId

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    rsTemp.Fields("ProformaState").Value = sState
    rsTemp.Update
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    
End Sub

'---------------------------------------------------------------------------------------------
Public Function DiagSavePLMFile(sTrialName As String, lTrialId As Long, sSite As String, lPersonId As Long) As String
'---------------------------------------------------------------------------------------------
' Save a subject state file as a PLM file
'---------------------------------------------------------------------------------------------
Dim sPLMFile As String
Dim sFileName As String
Dim sMSG As String

    On Error GoTo Muppeted
    
    sPLMFile = DiagGetPLMFile(lTrialId, sSite, lPersonId, sMSG)
    If sPLMFile <> "" Then
        sFileName = gsTEMP_PATH & sTrialName & "_" & sSite & "_" & lPersonId & ".plm"
        Call StringToFile(sFileName, sPLMFile)
        DiagSavePLMFile = "PLM file saved as: " & sFileName & vbCrLf _
                & "Length of file: " & Len(sPLMFile)
    ElseIf sMSG > "" Then
        DiagSavePLMFile = " *** " & sMSG & " & sTrialName & " / " & sSite & " / " & lPersonId"
    Else
        DiagSavePLMFile = " *** PLM file is NULL: " & sTrialName & "/" & sSite & "/" & lPersonId
    End If
    
    Exit Function
    
Muppeted:
    DiagSavePLMFile = " *** Error with subject state: " & lTrialId & "/" & sSite & "/" & lPersonId & vbCrLf _
            & Err.Number & "-" & Err.Description
End Function

'--------------------------------------------------------------------
Public Function LoadAREZZOSubject(oArezzo As Arezzo_DM, sStudyName As String, _
                            lTrialId As Long, sSite As String, lPersonId As Long) As Boolean
'--------------------------------------------------------------------
' This loads the CLM and PLM files into AREZZO
' Returns True if Ok, or False if there was a problem
'--------------------------------------------------------------------
Dim sMSG As String

    On Error GoTo ErrLabel
    
    oArezzo.ALM.ArezzoFile = DiagGetCLMFile(sStudyName)
    oArezzo.ALM.GuidelineInstance.Clear
    
    ' Ignore errors with PLM but report failure
    On Error GoTo PLMFailure
    oArezzo.ALM.SetState DiagGetPLMFile(lTrialId, sSite, lPersonId, sMSG)
    LoadAREZZOSubject = (sMSG = "")

Exit Function
PLMFailure:
    LoadAREZZOSubject = False

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.LoadAREZZOSubject"
    
End Function

'--------------------------------------------------------------------
Public Sub SaveAREZZOSubject(oArezzo As Arezzo_DM, lTrialId As Long, sSite As String, lPersonId As Long)
'--------------------------------------------------------------------
' This saves the current AREZZO Patient State in the MACRO DB
'--------------------------------------------------------------------

    Call DiagUpdatePLMFile(oArezzo.ALM.GetState, lTrialId, sSite, lPersonId)
    
End Sub

'---------------------------------------------------------------------------------------------
Public Function CLMMemory(sStudyName As String, oArezzo As Arezzo_DM) As String
'---------------------------------------------------------------------------------------------
' Report on AREZZO memory stats when this CLM file is loaded
'---------------------------------------------------------------------------------------------
Dim sReport As String

    On Error GoTo ErrLabel
    
    oArezzo.ALM.ArezzoFile = DiagGetCLMFile(sStudyName)
    sReport = sStudyName & " CLM file loaded" & vbCrLf
    CLMMemory = sReport & MemInfoReport(oArezzo)
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.CLMMemory"
    
End Function

'---------------------------------------------------------------------------------------------
Public Function PLMMemory(sSubjSpec As String, sStudyName As String, _
                            lTrialId As Long, sSite As String, lPersonId As Long, oArezzo As Arezzo_DM) As String
'---------------------------------------------------------------------------------------------
' Report on AREZZO memory stats when this subject is loaded
'---------------------------------------------------------------------------------------------
Dim sReport As String

    On Error GoTo ErrLabel
    
    sReport = "Subject: " & sSubjSpec
    If LoadAREZZOSubject(oArezzo, sStudyName, lTrialId, sSite, lPersonId) Then
        sReport = sReport & " CLM and PLM files loaded" & vbCrLf
        PLMMemory = sReport & MemInfoReport(oArezzo)
    Else
        PLMMemory = sReport & vbCrLf & gsPLM_ERROR
    End If
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.PLMMemory"
    
End Function

'---------------------------------------------------------------------------------------------
Private Function MemInfoReport(oArezzo As Arezzo_DM) As String
'---------------------------------------------------------------------------------------------
' Report on current AREZZO memory stats
'---------------------------------------------------------------------------------------------
Dim sReport As String
Dim sMemory As String
Dim vMemory As Variant
Dim sQuery As String
Dim sR As String
Dim sglPT As Single
Dim sglPF As Single
Dim sglTT As Single
Dim sglTF As Single

    sReport = ""
    
    sQuery = "memory_info. "
    sMemory = oArezzo.ALM.GetPrologResult(sQuery, sR)
    ' We get PT-PF-TT-TF
    vMemory = Split(sMemory, "-")
    sglPT = vMemory(0)
    sglPF = vMemory(1)
    sglTT = vMemory(2)
    sglTF = vMemory(3)
    sReport = sReport & "Program Space (K)" & vbCrLf & MemStats(sglPT, sglPF)
    sReport = sReport & "Text Space (K)" & vbCrLf & MemStats(sglTT, sglTF)
    
    MemInfoReport = sReport

End Function

'---------------------------------------------------------------------------------------------
Private Function MemStats(sglTotal As Single, sglFree As Single) As String
'---------------------------------------------------------------------------------------------
' Show the memory statistics nicely
'---------------------------------------------------------------------------------------------
Dim sMem As String

    sMem = "    TOTAL: " & CLng(sglTotal / 1024) & vbCrLf
    sMem = sMem & "    USED:  " & CLng((sglTotal - sglFree) / 1024) _
                & " (" & Format((100 * (sglTotal - sglFree) / sglTotal), "#0.0") & "%)" & vbCrLf
    sMem = sMem & "    FREE:  " & CLng(sglFree / 1024) _
                & " (" & Format(((sglFree / sglTotal) * 100), "#0.0") & "%)" & vbCrLf
    MemStats = sMem
    
End Function

'---------------------------------------------------------------------------------------------
Public Function CLMAnalyse(sStudyName As String, oArezzo As Arezzo_DM) As String
'---------------------------------------------------------------------------------------------

    oArezzo.ALM.ArezzoFile = DiagGetCLMFile(sStudyName)
    CLMAnalyse = sStudyName & " CLM File" & vbCrLf & CLMReport(oArezzo)

End Function

'---------------------------------------------------------------------------------------------
Private Function CLMReport(oArezzo As Arezzo_DM) As String
'---------------------------------------------------------------------------------------------
' See what's in the AREZZO CLM file
' Lists counts of: Tasks, DataItems, Internal triggers, Preconditions and Scheduling constraints
'---------------------------------------------------------------------------------------------
Dim sQuery As String
Dim sR As String

    sQuery = "allcounts. "
    CLMReport = oArezzo.ALM.GetPrologResult(sQuery, sR)
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
Public Function DataIntegrityReport(lTrialId As Long, sSite As String, lPersonId As Long, _
                                    ByVal sDataFileName As String, _
                                    ByRef sDeleteDataFileName As String, _
                                    ByRef bStuffToDelete As Boolean) As String
'-------------------------------------------------------------------------------------------------------------------------
' Read the AREZZO data values in the given file
' and check that each one exists in the DIR table
' Written for the DTU's FOURT study
' File expected to contain 6 comma-separated values per line
' ResponseTaskId, RptNo, DataValueKey, DataName, DateTimeStamp, DataValue
' including a header row of text labels
' File must have a .csv extension
'-------------------------------------------------------------------------------------------------------------------------
Dim nIOFileNumber As Integer
Dim sDataLine As String
Dim lLineCount As Long
Dim sWhere As String
Dim sReport As String
Dim lMissing As Long
Dim sPrologQuery As String

    On Error GoTo ErrLabel
    
    HourglassOn
    
    bStuffToDelete = False
    sDeleteDataFileName = ""
    
    sWhere = " WHERE ClinicaltrialId = " & lTrialId _
        & " AND TrialSite = '" & sSite & "'" _
        & " AND PersonId = " & lPersonId
    
    'open the File
    nIOFileNumber = FreeFile
    Open sDataFileName For Input As #nIOFileNumber
    
    lMissing = 0
    lLineCount = 0
    sReport = ""
    sPrologQuery = ""
    
    'Read the file line by line, ignoring the header row
    If Not EOF(nIOFileNumber) Then
        ' Gobble up the header row
        Line Input #nIOFileNumber, sDataLine
        ' Now read the rest
        Do While Not EOF(nIOFileNumber)
            Line Input #nIOFileNumber, sDataLine
            If Trim(sDataLine) > "" Then
                lLineCount = lLineCount + 1
                sReport = sReport & MatchDataValue(sDataLine, sWhere, lMissing, sPrologQuery)
            End If
        Loop
    End If
  
    ' Close the input file
    Close #nIOFileNumber
        
    If sReport = "" Then
        sReport = "No data mismatches found"
    Else
        bStuffToDelete = True
        ' Write Prolog query to a file (same as data file but with "_del" and .pl extension)
        ' Assume data file has .csv extension
        sDeleteDataFileName = Left(sDataFileName, Len(sDataFileName) - Len(".csv")) & "_del.pl"
        Call StringToFile(sDeleteDataFileName, sPrologQuery)
    End If
    
    sReport = sReport & vbCrLf & lLineCount & " AREZZO data values checked"
    sReport = sReport & vbCrLf & lMissing & " data values missing from MACRO database"
    
    DataIntegrityReport = sReport
    
    HourglassOff

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.DataIntegrityReport"
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
Private Function MatchDataValue(ByVal sDataLine As String, ByVal sWhere As String, _
    ByRef lMissing As Long, ByRef sPrologQuery As String) As String
'-------------------------------------------------------------------------------------------------------------------------
' sDataline contains the data details
' Update lMissing if this is missing from the DIR table
' sWhere is the SQL WHERE clause for study/site/personid
' sPrologQuery is where we add the psf_extra/1 calls
'-------------------------------------------------------------------------------------------------------------------------
Dim lResponseTaskId As Long
Dim nRptNo As Integer
Dim aDVArray() As String
Dim lDVKey As Long
Dim sDVName As String
Dim sVal As String
Dim sTime As String
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sMatch As String
Dim nIndex As Integer
Dim sComma As String

    sMatch = ""
    ' We expect 6 comma-separated elements per line (NCJ 29 Jan 07 - Added RptNo so it's now 6)
    aDVArray = Split(sDataLine, ",")
    lResponseTaskId = CLng(aDVArray(0))
    nRptNo = CInt(aDVArray(1))
    lDVKey = CLng(aDVArray(2))
    sDVName = aDVArray(3)
    ' Read the other array values only if we need them later
    
    ' Screen out V:F:date questions because these don't exist in MACRO
    If Right(sDVName, 5) <> ":date" Then
        sSQL = "SELECT ResponseTaskID from DataItemResponse " & sWhere _
                & " AND ResponseTaskID = " & lResponseTaskId _
                & " AND RepeatNumber = " & nRptNo
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
        ' Log it if there's no response AND the eForm is requested
        ' (this screens out form dates which are generally OK)
        'If rsTemp.EOF And EFormIsRequested(lResponseTaskId, sWhere) Then
        If rsTemp.EOF Then
            lMissing = lMissing + 1
            ' Retrieve the rest of the info on this data value
            sTime = aDVArray(4)
            ' The Value may contain commas, so we stitch together everything that's left
            sVal = ""
            sComma = ""     ' No comma first time through
            For nIndex = 5 To UBound(aDVArray)
                sVal = sVal & sComma & aDVArray(nIndex)
                sComma = ","    ' We want a comma from now on
            Next
            sVal = Replace(sVal, """", "")   ' Remove any surrounding double quotes
            sMatch = sDVName & " = " & sVal & ", " & lResponseTaskId & ", " & nRptNo & ", " & lDVKey & ", " & sTime & vbCrLf
            sPrologQuery = sPrologQuery & "psf_extra(" & lDVKey & "). " & vbCrLf
        End If
        
        rsTemp.Close
        Set rsTemp = Nothing
    End If
    
    MatchDataValue = sMatch

End Function

'-------------------------------------------------------------------------------------------------------------------------
Private Function EFormIsRequested(ByVal lResponseTaskId As String, ByVal sWhere As String) As Boolean
'-------------------------------------------------------------------------------------------------------------------------
' Return TRUE if this response's eForm is requested
'-------------------------------------------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim lCRFPageTaskId As Long
Dim bRequested As Boolean

    bRequested = True
    ' Get the CRFPageTaskID from the ResponseTaskID
    lCRFPageTaskId = CLng(Left(lResponseTaskId, 5))
    
    sSQL = "Select CRFPageStatus from CRFPAGEINSTANCE " & sWhere _
            & " AND CRFPAGETASKID = " & lCRFPageTaskId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not rsTemp.EOF Then
        If rsTemp!CRFPageStatus <> -10 Then
            ' It's not requested
            bRequested = False
        End If
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    EFormIsRequested = bRequested

End Function

''----------------------------------------------------------------------
'Private Sub EnsureLoaded(oArezzo As Arezzo_DM, sFile As String)
''----------------------------------------------------------------------
'' Load the specified file into Prolog
'' by calling ensure_loaded/1
''----------------------------------------------------------------------
'Dim sQuery As String
'Dim sR As String
'
'Const sZeros = ", write( '0000' ). "
'
'    sQuery = "ensure_loaded( '" & sFile & "' )" & sZeros
'    Call oArezzo.ALM.GetPrologResult(sQuery, sR)
'
'End Sub

'----------------------------------------------------------------------
Public Function DoTheDeletes(oArezzo As Arezzo_DM, sDelPLFile As String) As Boolean
'-----------------------------------------------------------------------
' Delete data from the AREZZO patient state
' Assume patient already loaded, and
' assume sDelPLFile is a Prolog file containing psf_extra/1 clauses
'-----------------------------------------------------------------------
Dim sQuery As String
Dim sR As String

Const sZeros = ", write( '0000' ). "

    DoTheDeletes = False
    
    On Error GoTo StuffHappens
    
    ' Ensure there are no stray psf_extra clauses hanging around
    sQuery = "retractall( psf_extra(_) )" & sZeros
    Call oArezzo.ALM.GetPrologResult(sQuery, sR)
    
    ' Load the pl file into Prolog
    sQuery = "ensure_loaded( '" & sDelPLFile & "' )" & sZeros
    Call oArezzo.ALM.GetPrologResult(sQuery, sR)
    ' Remove all the data specified by the psf_extra(DVKey) clauses
    sQuery = "forall( psf_extra(DVKey), plm_remove_data( DVKey ) )" & sZeros
    Call oArezzo.ALM.GetPrologResult(sQuery, sR)
    ' Abolish the file
    sQuery = "abolish_files( '" & sDelPLFile & "' )" & sZeros
    Call oArezzo.ALM.GetPrologResult(sQuery, sR)
    ' Success!
    DoTheDeletes = True
 
Exit Function
StuffHappens:
    ' Ensure we've unloaded the file
    On Error Resume Next
    If sDelPLFile <> "" Then
        sQuery = "abolish_files( '" & sDelPLFile & "' )" & sZeros
        Call oArezzo.ALM.GetPrologResult(sQuery, sR)
    End If
    
End Function

'---------------------------------------------------------------------
Public Sub LogToFile(ByVal sText As String)
'---------------------------------------------------------------------
' Add text to the log file
' Precede with carriage return
'---------------------------------------------------------------------
Dim n As Integer

    n = FreeFile
    Open LogFileName For Append As n
    Print #n, vbCrLf & sText
    Close n
    
End Sub

'---------------------------------------------------------------------
Private Function LogFileName() As String
'---------------------------------------------------------------------
' Get a suitable log file name in the MACRO Temp folder
' New one each day
'---------------------------------------------------------------------

    LogFileName = gsTEMP_PATH & "Diagnostic_" & Format(Now, "yyyymmdd") & ".log"

End Function

''---------------------------------------------------------------------
'Public Function GetTaskIDsFile(sTrialName As String, lTrialId As Long, sSite As String, lPersonId As Long) As String
''---------------------------------------------------------------------
'' Assert a Prolog file of TaskIds to be used when running this guideline
''---------------------------------------------------------------------
'Dim sClauses As String
'Dim sFileName As String
'
'    sClauses = GetAllTaskIDClauses(lTrialId, sSite, lPersonId)
'    sFileName = gsTEMP_PATH & sTrialName & "_" & sSite & "_" & lPersonId & "_TASKIDS.pl"
'    Call StringToFile(sFileName, sClauses)
'    GetTaskIDsFile = sFileName
'
'End Function

''---------------------------------------------------------------------
'Private Function GetAllTaskIDClauses(lTrialId As Long, sSite As String, lPersonId As Long) As String
''---------------------------------------------------------------------
'' Get the Prolog file of assertions of visit TaskIds and eForm TaskIDs
'' Returns a string with the relevant macro_predefinedID/4 clauses
''---------------------------------------------------------------------
'Dim sSQL As String
'Dim rsTemp As ADODB.Recordset
'Dim lCRFPageTaskId As Long
'Dim sAssert As String
'
'Const sPREDEFINED = "macro_predefinedID("
'
'    sAssert = ""
'
'    ' Do all the Visit Task Ids first
'    sSQL = "SELECT VisitTaskID, VisitCode, VisitCycleNumber " _
'            & " FROM VisitInstance, StudyVisit "
'    sSQL = sSQL & " WHERE VisitInstance.ClinicalTrialID = " & lTrialId _
'            & " AND VisitInstance.TrialSite = '" & sSite & "'" _
'            & " AND VisitInstance.PersonId = " & lPersonId _
'            & " AND StudyVisit.ClinicalTrialId = VisitInstance.ClinicalTrialID" _
'            & " AND StudyVisit.VisitId = VisitInstance.VisitID" _
'            & " ORDER BY VisitTaskId "
'
'    Set rsTemp = New ADODB.Recordset
'    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    Do While Not rsTemp.EOF
'        sAssert = sAssert & sPREDEFINED
'        ' NB Visit parent's TaskID is always 10000
'        sAssert = sAssert & LCase(rsTemp.Fields("VisitCode")) & ",10000,"
'        sAssert = sAssert & LCase(rsTemp.Fields("VisitCycleNumber")) & ","
'        sAssert = sAssert & LCase(rsTemp.Fields("VisitTaskID")) & "). " & vbCrLf
'        rsTemp.MoveNext
'    Loop
'
'    rsTemp.Close
'
'    ' Now do all the eForm Task Ids
'    sSQL = "SELECT CRFPageTaskID, CRFPageCode, VisitTaskID, CRFPageCycleNumber " _
'            & " FROM CRFPAGEInstance, CRFPage, VisitInstance "
'    sSQL = sSQL & " WHERE CRFPAGEInstance.ClinicalTrialID = " & lTrialId _
'            & " AND CRFPageInstance.TrialSite = '" & sSite & "'" _
'            & " AND CRFPageInstance.PersonId = " & lPersonId
'    sSQL = sSQL & " AND CRFPage.ClinicalTrialId = CRFPageInstance.ClinicalTrialID" _
'            & " AND CRFPage.CRFPageId = CRFPageInstance.CRFPageID" _
'            & " AND VisitInstance.ClinicalTrialId = CRFPageInstance.ClinicalTrialID" _
'            & " AND VisitInstance.TrialSite = CRFPageInstance.TrialSite" _
'            & " AND VisitInstance.PersonId = CRFPageInstance.PersonId" _
'            & " AND VisitInstance.VisitId = CRFPageInstance.VisitID" _
'            & " AND VisitInstance.VisitCycleNumber = CRFPageInstance.VisitCycleNumber" _
'            & " ORDER BY CRFPageTaskID"
'
'    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    Do While Not rsTemp.EOF
'        sAssert = sAssert & sPREDEFINED
'        sAssert = sAssert & LCase(rsTemp.Fields("CRFPageCode")) & ","
'        sAssert = sAssert & LCase(rsTemp.Fields("VisitTaskId")) & ","
'        sAssert = sAssert & LCase(rsTemp.Fields("CRFPageCycleNumber")) & ","
'        sAssert = sAssert & LCase(rsTemp.Fields("CRFPageTaskID")) & "). " & vbCrLf
'        rsTemp.MoveNext
'    Loop
'
'    rsTemp.Close
'    Set rsTemp = Nothing
'
'    GetAllTaskIDClauses = sAssert
'
'Exit Function
'
'ErrLabel:
'    Err.Raise Err.Number, , Err.Description & "|modDiagnostic.GetAllTaskIDClauses"
'End Function
'
'

