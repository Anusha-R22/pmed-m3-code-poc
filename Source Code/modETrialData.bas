Attribute VB_Name = "modETrialData"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       modETrialData.bas
'   Author      Paul Norris, 23/09/99
'   Purpose:    All common TrialData functions for the StudyDefintion project are in this module.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   PN  29/09/99    Conversion to ADO form DAO
'   Mo 13/12/99     Id's from integer to Long
'----------------------------------------------------------------------------------------'

Option Explicit

'---------------------------------------------------------------------
Public Function gdsTrialHistory(lClinicalTrialId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
Dim sSQL As String
' PN 29/09/99   convert routine to return an ADO recordset
Dim rsTrialHistory As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    'REM 01/09/03 - Changed SQL stament to display study version properly
    sSQL = "SELECT ClinicalTrialId, StatusId, StatusChangedTimestamp, UserName, " _
        & " (SELECT MAX(StudyVersion) FROM StudyVersion " _
        & " WHERE StudyVersion.ClinicalTrialId = TrialStatusHistory.ClinicalTrialId " _
        & " AND StudyVersion.VersionTimestamp < StatusChangedTimestamp) VersionId " _
        & " FROM TrialStatusHistory " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " ORDER BY TrialStatusChangeId"
    Set rsTrialHistory = New ADODB.Recordset
    rsTrialHistory.Open sSQL, MacroADODBConnection
    
'    'Mo Morris 30/8/01 Db Audit (UserId to UserName)
'    sSQL = "SELECT  StudyDefinition.UserName as StudyDefinitionUserName ,StudyDefinition.StudyDefinitionTimeStamp,'', TrialStatusHistory.* FROM StudyDefinition, TrialStatusHistory " _
'                        & " WHERE StudyDefinition.ClinicalTrialId = " & ClinicalTrialId _
'                        & " AND StudyDefinition.ClinicalTrialId = TrialStatusHistory.ClinicalTrialId " _
'                        & " AND StudyDefinition.VersionId = TrialStatusHistory.VersionId " _
'                        & " ORDER BY StudyDefinition.VersionId, TrialStatusHistory.TrialStatusChangeId"
'    Set rsTrialHistory = New ADODB.Recordset
'    rsTrialHistory.Open sSQL, MacroADODBConnection
    
    ' return the recordset then clean up
    Set gdsTrialHistory = rsTrialHistory
    Set rsTrialHistory = Nothing

Exit Function
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "gdsTrialHidstory", "modETrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
End Function


