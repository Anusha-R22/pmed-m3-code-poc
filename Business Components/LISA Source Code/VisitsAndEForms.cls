VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VisitsAndEForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------------------------------------------------------
' File: VisitsAndEForms.cls
' Copyright: InferMed Ltd 2004 All Rights Reserved
' Author: Nicky Johns, April 2004
' Purpose: Class to retrieve visit and eForm info directly from DB using SQL
'   Fast access specifically for LISA Phase II project
'------------------------------------------------------------------
' REVISIONS
' NCJ 27 Apr 04 - Initial development
' NCJ 15 Jul 04 - Also store LastVisitCycles collection
' NCJ 19 Jul 04 -  Added AllEFIsInLastVisit
'------------------------------------------------------------------

Option Explicit

Private mlStudyId As Long
Private msSite As String
Private mlPersonId As Long

Private msConnectionString As String

Private msSQLSubjectWhere As String
Private msSQLStudyWhere As String

Private mcolVisitsByCode As Collection
Private mcolEFormsByCode As Collection
Private mcolLastVisitCycles As Collection       ' NCJ 15 Jul 04

'------------------------------------------------------------------
Public Sub Init(sDBConn As String, _
                lStudyId As Long, sSite As String, lSubjId As Long)
'------------------------------------------------------------------
' Initialise class for this DB and subject
'------------------------------------------------------------------

    msConnectionString = sDBConn
    
    mlStudyId = lStudyId
    msSite = sSite
    mlPersonId = lSubjId
    
    msSQLStudyWhere = " WHERE ClinicalTrialId = " & mlStudyId
    msSQLSubjectWhere = msSQLStudyWhere & _
                    " AND TrialSite = '" & msSite & "'" & _
                    " AND PersonId = " & mlPersonId

    Set mcolVisitsByCode = New Collection
    Set mcolEFormsByCode = New Collection
    Set mcolLastVisitCycles = New Collection
    
End Sub

'------------------------------------------------------------------
Public Function VisitIdByCode(ByVal sVisitCode As String) As Long
'------------------------------------------------------------------
' Return Visit Id for this visit code
'------------------------------------------------------------------
Dim sSQL As String
Dim oDBCon As ADODB.Connection
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrLabel
    
    ' Do we have it already?
    If CollectionMember(mcolVisitsByCode, sVisitCode, False) Then
        VisitIdByCode = mcolVisitsByCode.Item(sVisitCode)
    Else
        ' Go to the database
        Set oDBCon = New ADODB.Connection
        Call oDBCon.Open(msConnectionString)
        oDBCon.CursorLocation = adUseClient
        
        sSQL = "SELECT VisitId FROM StudyVisit " & msSQLStudyWhere _
                & " AND VisitCode = '" & sVisitCode & "'"
        
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, oDBCon, adOpenKeyset, adLockReadOnly, adCmdText
        
        If rsTemp.RecordCount = 0 Then
            VisitIdByCode = -1
        Else
            VisitIdByCode = rsTemp!VisitId
            ' Add it to our collection for future reference
            mcolVisitsByCode.Add CLng(rsTemp!VisitId), sVisitCode
        End If
        
        Call rsTemp.Close
        Set rsTemp = Nothing
        Call oDBCon.Close
        Set oDBCon = Nothing
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|VisitsAndEForms.VisitIdByCode"
    
End Function

'------------------------------------------------------------------
Public Function EFormIdByCode(ByVal sEFormCode As String) As Long
'------------------------------------------------------------------
' Return EForm Id for this eForm code
'------------------------------------------------------------------
Dim sSQL As String
Dim oDBCon As ADODB.Connection
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrLabel
    
    ' Do we have it already?
    If CollectionMember(mcolEFormsByCode, sEFormCode, False) Then
        EFormIdByCode = mcolEFormsByCode.Item(sEFormCode)
    Else
        ' Go to the database
        Set oDBCon = New ADODB.Connection
        Call oDBCon.Open(msConnectionString)
        oDBCon.CursorLocation = adUseClient
        
        sSQL = "SELECT CRFPageId FROM CRFPage " & msSQLStudyWhere _
                & " AND CRFPageCode = '" & sEFormCode & "'"
        
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, oDBCon, adOpenKeyset, adLockReadOnly, adCmdText
        
        If rsTemp.RecordCount = 0 Then
            EFormIdByCode = -1
        Else
            EFormIdByCode = rsTemp!CRFPageId
            ' Add it to our collection for future reference
            mcolEFormsByCode.Add CLng(rsTemp!CRFPageId), sEFormCode
        End If
        
        Call rsTemp.Close
        Set rsTemp = Nothing
        Call oDBCon.Close
        Set oDBCon = Nothing
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|VisitsAndEForms.EFormIdByCode"
    
End Function

'------------------------------------------------------------------
Public Function AllEFIsByVisitEForm(ByVal lVisitId As Long, ByVal lEformId As Long) As Collection
'------------------------------------------------------------------
' Return collection of all EFI TaskIds matching this visit and eForm
' Collection may be empty if no matching EFIs
'------------------------------------------------------------------
Dim sSQL As String
Dim oDBCon As ADODB.Connection
Dim rsTemp As ADODB.Recordset
Dim colEFTaskIds As Collection

    On Error GoTo ErrLabel
    
    Set colEFTaskIds = New Collection
    
    ' Go to the database
    Set oDBCon = New ADODB.Connection
    Call oDBCon.Open(msConnectionString)
    oDBCon.CursorLocation = adUseClient
    
    ' Retrieve all the EFIs for the current subject matching this visit ID and eFormID
    sSQL = "SELECT CRFPageTaskId FROM CRFPageInstance " _
            & msSQLSubjectWhere _
            & " AND CRFPageId = " & lEformId _
            & " AND VisitId = " & lVisitId
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, oDBCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    ' Add the task ids to collection
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            colEFTaskIds.Add CLng(rsTemp!CRFPageTaskId)
            rsTemp.MoveNext
        Loop
    End If
    
    Call rsTemp.Close
    Set rsTemp = Nothing
    
    Call oDBCon.Close
    Set oDBCon = Nothing

    Set AllEFIsByVisitEForm = colEFTaskIds
    Set colEFTaskIds = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|VisitsAndEForms.AllEFIsByVisitEForm"
    
End Function

'------------------------------------------------------------------
Public Function AllEFIsInLastVisit(ByVal lVisitId As Long, sEFormCodes As String) As Collection
'------------------------------------------------------------------
' Return collection of all EFI TaskIds matching this visit (last cycle only) and set of eForm codes
' sEFormCodes is comma-separated list of ready-quoted eForm codes
' Collection may be empty if no matching EFIs
'------------------------------------------------------------------
Dim sSQL As String
Dim oDBCon As ADODB.Connection
Dim rsTemp As ADODB.Recordset
Dim colEFTaskIds As Collection
Dim nCycleNo As Integer

    On Error GoTo ErrLabel
    
    Set colEFTaskIds = New Collection
    
    ' Go to the database
    Set oDBCon = New ADODB.Connection
    Call oDBCon.Open(msConnectionString)
    oDBCon.CursorLocation = adUseClient
    
    ' Get the last visit cycle no
    nCycleNo = LastVisitCycle(oDBCon, lVisitId)
    
    If nCycleNo > 0 Then
        ' Retrieve all the EFIs for the current subject matching this visit ID and set of eForm codes
        sSQL = "SELECT CRFPageTaskId FROM CRFPageInstance, CRFPage " _
                & " WHERE CRFPageInstance.ClinicalTrialId = " & mlStudyId _
                & " AND CRFPageInstance.TrialSite = '" & msSite & "'" _
                & " AND CRFPageInstance.PersonId = " & mlPersonId _
                & " AND CRFPageInstance.VisitId = " & lVisitId _
                & " AND VisitCycleNumber = " & nCycleNo _
                & " AND CRFPageInstance.CRFPageId = CRFPage.CRFPageId" _
                & " AND CRFPageInstance.ClinicalTrialId = CRFPage.ClinicalTrialId" _
                & " AND CRFPage.CRFPageCode IN (" & sEFormCodes & ") "

        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, oDBCon, adOpenKeyset, adLockReadOnly, adCmdText
        
        ' Add the task ids to collection
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                colEFTaskIds.Add CLng(rsTemp!CRFPageTaskId)
                rsTemp.MoveNext
            Loop
        End If
        
        Call rsTemp.Close
        Set rsTemp = Nothing
    End If
    
    Call oDBCon.Close
    Set oDBCon = Nothing

    Set AllEFIsInLastVisit = colEFTaskIds
    Set colEFTaskIds = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|VisitsAndEForms.AllEFIsInLastVisit"

End Function

'------------------------------------------------------------------
Private Function LastVisitCycle(oDBCon As ADODB.Connection, ByVal lVisitId As Long) As Long
'------------------------------------------------------------------
' Return last visit cycle of this visit
' Returns -1 if no visit cycles exist
' NCJ 14 Jul 04 - Keep collection of "done" visits
'------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim nCycle As Integer
Dim sVisitIdKey As String

    On Error GoTo ErrLabel
    
    sVisitIdKey = "K" & lVisitId
    ' Do we have it already?
    If CollectionMember(mcolLastVisitCycles, sVisitIdKey, False) Then
        LastVisitCycle = mcolLastVisitCycles.Item(sVisitIdKey)
    Else
    
        ' Get the max cycle no. for the current subject matching this visit ID
        sSQL = "SELECT MAX(VisitCycleNumber) as LastCycle FROM VisitInstance " _
                & msSQLSubjectWhere _
                & " AND VisitId = " & lVisitId
        
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, oDBCon, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If IsNull(rsTemp!LastCycle) Then
            LastVisitCycle = -1
        Else
            LastVisitCycle = rsTemp!LastCycle
            mcolLastVisitCycles.Add CLng(rsTemp!LastCycle), sVisitIdKey
        End If

        Call rsTemp.Close
        Set rsTemp = Nothing
    End If
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|VisitsAndEForms.LastVisitCycle"
    
End Function

'------------------------------------------------------------------
Public Sub Terminate()
'------------------------------------------------------------------
' Tidy up
'------------------------------------------------------------------

    Set mcolVisitsByCode = Nothing
    Set mcolEFormsByCode = Nothing
    Set mcolLastVisitCycles = Nothing

End Sub

'------------------------------------------------------------------
Private Sub Class_Terminate()
'------------------------------------------------------------------

    Call Terminate

End Sub
