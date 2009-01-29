Attribute VB_Name = "basVisitsData"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       basVisitSQL.bas
'   Author:     Andrew Newbigging, June 1997
'   Purpose:    Common SQL functions for study visits and the schedule.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
 '  1       Andrew Newbigging       4/06/97
'   2       Andrew Newbigging       4/07/97
'   3       Andrew Newbigging       11/07/97
'   4       Andrew Newbigging       17/07/97
'   5       SDM                     10/11/99    Copied in error handling rountines
'   Mo Morris   15/11/99    DAO to ADO conversion
'   NCJ 11 Dec 99 - Ids to Long
'   NCJ 20 Mar 00   SR 3207 Get and Set Visit Instance Dates
'   NCJ 12 Feb 01 - Set Changed flag in SaveVisitInstanceDate
'
'------------------------------------------------------------------------------------'
Option Explicit

'---------------------------------------------------------------------
Public Function gdsStudyVisitList(ClinicalTrialId As Long, VersionId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
'Creates a ReadOnly recordset
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT  * FROM StudyVisit " _
        & " WHERE ClinicalTrialId = " & ClinicalTrialId _
        & " AND VersionId = " & VersionId
    Set gdsStudyVisitList = New ADODB.Recordset
    gdsStudyVisitList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsStudyVisitList", "BASVISITSQL.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function gdsStudyVisitCRFPages(ClinicalTrialId As Long, VersionId As Integer, StudyVisitId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
'Creates a ReadOnly recordset
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT  StudyVisitCRFPage.*,CRFPage.CRFTitle FROM StudyVisitCRFPage, CRFPage " _
        & " WHERE StudyVisitCRFPage.ClinicalTrialId = " & ClinicalTrialId _
        & " AND StudyVisitCRFPage.VersionId = " & VersionId _
        & " AND StudyVisitCRFPage.ClinicalTrialId = CRFPage.ClinicalTrialId " _
        & " AND StudyVisitCRFPage.VersionId = CRFPage.VersionId" _
        & " AND StudyVisitCRFPage.CRFPageId = CRFPage.CRFPageId" _
        & " AND VisitId = " & Str(StudyVisitId)
    Set gdsStudyVisitCRFPages = New ADODB.Recordset
    gdsStudyVisitCRFPages.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsStudyVisitCRFPages", "BASVISITSQL.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'------------------------------------------------------------
Public Sub SaveVisitInstanceDate(lClinicalTrialId As Long, _
                                 nPersonId As Integer, _
                                 sTrialSite As String, _
                                 lVisitTaskId As Long, _
                                 dblDate As Double)
'------------------------------------------------------------
' Save visit instance date
' NCJ 12 Feb 01 - Set Changed flag when visit date changes
' NCJ 26 Feb 01 - Set Changed flag on TrialSubject too!
'------------------------------------------------------------
Dim sSQL As String

    sSQL = "UPDATE VisitInstance " _
            & " SET VisitDate = " & dblDate _
            & ", Changed = " & Changed.Changed _
            & " WHERE ClinicalTrialId = " & lClinicalTrialId _
            & " AND PersonId = " & nPersonId _
            & " AND TrialSite = '" & sTrialSite & "'" _
            & " AND VisitTaskId = " & lVisitTaskId
    MacroADODBConnection.Execute sSQL
    
    ' NCJ 26/2/01 - Change TrialSubject too
    sSQL = "UPDATE TrialSubject " _
            & " SET Changed = " & Changed.Changed _
            & " WHERE ClinicalTrialId = " & lClinicalTrialId _
            & " AND TrialSite = '" & sTrialSite _
            & "' AND PersonId = " & nPersonId
    MacroADODBConnection.Execute sSQL

End Sub

'--------------------------------------
Public Function GetVisitInstanceDate(lClinicalTrialId As Long, _
                                 nPersonId As Integer, _
                                 sTrialSite As String, _
                                 lVisitTaskId As Long) As Double
'--------------------------------------
' Get visit instance date as double
'--------------------------------------
Dim sSQL As String
Dim rsVisitDate As ADODB.Recordset

    sSQL = "SELECT VisitDate " _
        & " FROM VisitInstance " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND PersonId = " & nPersonId _
        & " AND TrialSite = '" & sTrialSite & "'" _
        & " AND VisitTaskId = " & lVisitTaskId
    Set rsVisitDate = New ADODB.Recordset
    rsVisitDate.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText

    If Not IsNull(rsVisitDate!VisitDate) Then
        GetVisitInstanceDate = rsVisitDate!VisitDate
    Else
        GetVisitInstanceDate = 0
    End If
    
    Set rsVisitDate = Nothing
    
End Function



