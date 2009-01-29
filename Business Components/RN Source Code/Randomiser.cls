VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Randomiser"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   File:       Randomiser.cls
'   Copyright:  InferMed Ltd. 2003. All Rights Reserved
'   Author:     Nicky Johns, September 2003
'   Purpose:    Class to carry out randomisation for a subject in MACRO 3.0
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 29 Sept - 6 Oct 03 - Initial development
'   NCJ 7 Oct 03 - CLASS NO LONGER USED (code copied and pasted directly into frmEFormDataEntry and modRegistrationWWW.bas)
'----------------------------------------------------------------------------------------'

Option Explicit

' Possible randomisation results
Public Enum RandomisationResult
    rrSuccess = 0
    rrCantStratify = 1
    rrNoTreatments = 2
    rrError = 3
End Enum

'----------------------------------------------------------------------------------------'
Public Function RandomiseSubject(ByVal sDBConnection As String, _
                oSubject As StudySubject, _
                ByRef colResults As Collection) As Boolean
'----------------------------------------------------------------------------------------'
' Randomise a subject and return results in colResults
'----------------------------------------------------------------------------------------'
Dim oDBCon As ADODB.Connection
Dim sSQL As String
Dim rsRands As ADODB.Recordset
Dim sRandCode As String
Dim sStratValue As String
Dim sTreatment As String
Dim bHaveRandomised As Boolean

    On Error GoTo ErrHandler
    
    ' Have we randomised?
    bHaveRandomised = False
    Set colResults = New Collection
    
    Set oDBCon = New ADODB.Connection
    Call oDBCon.Open(sDBConnection)
    oDBCon.CursorLocation = adUseClient
    
    ' Check we have the necessary randomisation tables
    If RandomisationAvailable(oDBCon) Then
        ' Get the possible randomisation codes
        sSQL = "SELECT RandomisationCode, RandomisationCond, StratificationExpr" _
            & " FROM Randomisation WHERE " _
            & " ClinicalTrialId = " & oSubject.StudyId
        Set rsRands = New ADODB.Recordset
        rsRands.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    
        ' See if any need processing
        If rsRands.RecordCount > 0 Then
            rsRands.MoveFirst
            Do While Not rsRands.EOF
                sRandCode = rsRands!RandomisationCode
                If Not IsSubjectRandomised(oSubject, oDBCon, sRandCode) Then
                    ' See if we need to process this randomisation
                    If ShouldRandomiseSubject(oSubject, rsRands!RandomisationCond) Then
                        ' We're going to attempt a randomisation
                        bHaveRandomised = True
                        sStratValue = EvaluateStratification(oSubject, rsRands!StratificationExpr)
                        If sStratValue > "" Then
                            ' Assign treatment
                            Call AssignTreatment(oSubject, oDBCon, colResults, sRandCode, sStratValue)
                        Else
                            ' Add a "Can't stratify" error for this randomisation
                            Call IncludeResult(colResults, rrCantStratify, sRandCode, "")
                        End If
                    End If
                End If
                rsRands.MoveNext
            Loop
        End If
        rsRands.Close
        Set rsRands = Nothing
    End If
    
    Set oDBCon = Nothing
    
    RandomiseSubject = bHaveRandomised
    
Exit Function
ErrHandler:
    ' Set the results to show an error
    Call IncludeResult(colResults, rrError, "", Err.Description)
    ' Signal that there's something to look at
    RandomiseSubject = True
    
End Function

'----------------------------------------------------------------------------------------'
Private Function IsSubjectRandomised(oSubject As StudySubject, _
                                    oDBCon As ADODB.Connection, _
                                    ByVal sRandCode As String) As Boolean
'----------------------------------------------------------------------------------------'
' Returns TRUE if subject is already randomised on this randomisation
'----------------------------------------------------------------------------------------'
Dim rsRands As ADODB.Recordset
Dim sSQL As String

    IsSubjectRandomised = False
    
    On Error GoTo ErrHandler
    
    ' See if this subject has been allocated a treatment on this randomisation
    sSQL = "SELECT Treatment FROM Treatments WHERE " _
        & " ClinicalTrialId = " & oSubject.StudyId _
        & " AND RandomisationCode = '" & sRandCode & "'" _
        & " AND TrialSite = '" & oSubject.Site & "'" _
        & " AND PersonId = " & oSubject.PersonID
    Set rsRands = New ADODB.Recordset
    rsRands.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    If rsRands.RecordCount > 0 Then
        ' They've been done
        IsSubjectRandomised = True
    End If
    rsRands.Close
    
    Set rsRands = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "Randomiser.IsSubjectRandomised"

End Function

'----------------------------------------------------------------------------------------'
Private Function ShouldRandomiseSubject(oSubject As StudySubject, _
                                        ByVal sRandCond As String) As Boolean
'----------------------------------------------------------------------------------------'
' Returns TRUE if the randomisation condition is true
'----------------------------------------------------------------------------------------'
Dim sResult As String

    sResult = oSubject.Arezzo.EvaluateExpression(sRandCond)
    ShouldRandomiseSubject = (sResult = "true")

End Function

'----------------------------------------------------------------------------------------'
Private Function TreatmentDataItem(ByVal sRandCode As String) As String
'----------------------------------------------------------------------------------------'
' Returns the AREZZO data item name of the treatment for this randomisation
'----------------------------------------------------------------------------------------'

    TreatmentDataItem = "person:" & LCase(sRandCode) & ":treatment"
    
End Function

'----------------------------------------------------------------------------------------'
Private Function EvaluateStratification(oSubject As StudySubject, _
                                    ByVal sStratExpr As String) As String
'----------------------------------------------------------------------------------------'
' Evaluate stratification expression
' Returns empty string if no go
'----------------------------------------------------------------------------------------'
Dim sStrat As String

    sStrat = oSubject.Arezzo.EvaluateExpression(sStratExpr)
    If Not oSubject.Arezzo.ResultOK(sStrat) Then
        sStrat = ""
    End If
    EvaluateStratification = sStrat

End Function

'----------------------------------------------------------------------------------------'
Private Sub IncludeResult(colResults As Collection, _
                            ByVal enResultType As RandomisationResult, _
                            ByVal sRandCode As String, _
                            ByVal sTreatment As String)
'----------------------------------------------------------------------------------------'
' Add a result to the results collection
'----------------------------------------------------------------------------------------'
Dim sResult As String
Const sSep = "|"

    sResult = enResultType & sSep & sRandCode & sSep & sTreatment
    colResults.Add sResult

End Sub

'----------------------------------------------------------------------------------------'
Private Sub AssignTreatment(oSubject As StudySubject, _
                                oDBCon As ADODB.Connection, _
                                colResults As Collection, _
                                ByVal sRandCode As String, _
                                ByVal sStratValue As String)
'----------------------------------------------------------------------------------------'
' Get the treatment for this subject
' and add an entry to the colResults collection accordingly
' NCJ 6 Oct 03 - Add AREZZO data first, before updating Treatments table
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim oTimezone As Timezone
Dim rsTreats As ADODB.Recordset
Dim sTreat As String

    On Error GoTo ErrHandler
    
    sTreat = ""
    ' Get the next unassigned treatment
    sSQL = "SELECT * FROM Treatments WHERE" _
        & " ClinicalTrialId = " & oSubject.StudyId _
        & " AND RandomisationCode = '" & sRandCode & "'" _
        & " AND StratificationValue = '" & sStratValue & "'" _
        & " AND TrialSite IS NULL" _
        & " ORDER BY TreatSeqNo "
    
    Set rsTreats = New ADODB.Recordset
    rsTreats.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    ' Hopefully there is a record
    If rsTreats.RecordCount > 0 Then
        rsTreats.MoveFirst
        ' Quickly mark this one as ours!
        rsTreats!TrialSite = oSubject.Site
        Call rsTreats.Update
        ' Haul out the treatment
        sTreat = rsTreats!Treatment
        ' Add the data to AREZZO first to make sure we can do it
        Call oSubject.Arezzo.AddData(TreatmentDataItem(sRandCode), sTreat)
        If oSubject.Save = eSaveResponsesResult.srrSuccess Then
            ' Fill in the rest of the subject details
            rsTreats!PersonID = oSubject.PersonID
            If oSubject.Label > "" Then
                rsTreats!SubjectLabel = oSubject.Label
            End If
            rsTreats!RandomiseDate = IMedNow
            Set oTimezone = New Timezone
            rsTreats!RandomiseDate_TZ = oTimezone.TimezoneOffset
            Set oTimezone = Nothing
            Call rsTreats.Update
            ' Add this successful result to the collection
            Call IncludeResult(colResults, rrSuccess, sRandCode, sTreat)
        Else
            ' Can't update patient state so give up
            ' Release this treatment
            rsTreats!TrialSite = Null
            Call rsTreats.Update
            Call IncludeResult(colResults, rrError, sRandCode, sTreat)
        End If
    Else
        ' If the recordset is empty, we either have an unknown stratification value
        ' or we've run out of treatments for this value
        Call IncludeResult(colResults, rrNoTreatments, sRandCode, "")
    End If

    rsTreats.Close
    Set rsTreats = Nothing

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "Randomiser.AssignTreatment"
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function RandomisationAvailable(oDBCon As ADODB.Connection) As Boolean
'----------------------------------------------------------------------------------------'
' Are the randomisation tables available in this database?
'----------------------------------------------------------------------------------------'
Dim rsRands As ADODB.Recordset

    RandomisationAvailable = True
    
    On Error GoTo NoTable
    
    Set rsRands = New ADODB.Recordset
    ' Just try and access the table with a dummy select
    rsRands.Open "SELECT RandomisationCode FROM Randomisation WHERE 1=2", oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    rsRands.Close
    rsRands.Open "SELECT RandomisationCode FROM Treatments WHERE 1=2", oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    rsRands.Close
    Set rsRands = Nothing

Exit Function
NoTable:
    ' Error reading from tables so return FALSE
    Set rsRands = Nothing
    RandomisationAvailable = False
End Function


