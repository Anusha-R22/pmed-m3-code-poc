VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIMsgStatic"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Public static routines for MIMessage these form a public wrapper for general MIMessage routines"
'----------------------------------------------------------------------------------------'
'   File:       MIMsgStatic.cls
'   Copyright:  InferMed Ltd. 2001-2002. All Rights Reserved
'   Author:     Toby Aldridge, Nov 2001
'   Purpose:    Public static routines for MIMessage
'                   these form a public wrapper for general MIMessage routines
'----------------------------------------------------------------------------------------'
' REVISIONS
'   NCJ 16 Oct 02 - GetScopeText
'   TA  20 Jan 03 - GetMIMTypeText now has plural option
'   TA 12/03/2003 - MOved sql to update MIMsgStatus here
'   TA 24/08/2004 - Change hierachical SDV status inheritance (non, cancelled, done no bubbled up) - Organon request
'----------------------------------------------------------------------------------------'


Option Explicit

'public enum for erros - so that external clients can check
Public Enum MIMsgErrors
    mimeInvalidText = vbObjectError + 100009
    mimeInvalidPriority = vbObjectError + 100010
    mimeAlreadyProcessed = vbObjectError + 100011
    mimeInvalidUserName = vbObjectError + 100012
    mimeNotCreatedHere = vbObjectError + 100013
    mimeInvalidForThisStatus = vbObjectError + 100014
    mimeAlreadySent = vbObjectError + 100015
End Enum

'---------------------------------------------------------------------
Public Function GetStatusText(ByVal nMIMsgType As MIMsgType, ByVal nStatus As Integer) As String
Attribute GetStatusText.VB_Description = "Get the text that represents a MIMsg's status"
'---------------------------------------------------------------------
' Get the text that represents a MIMsg's status
'---------------------------------------------------------------------
    
    GetStatusText = modMIMsg.GetStatusText(nMIMsgType, nStatus)
    

End Function

'---------------------------------------------------------------------
Public Function GetMIMTypeText(ByVal nType As MIMsgType, Optional bPlural As Boolean = False) As String
Attribute GetMIMTypeText.VB_Description = "Get the text that represents a MIMessage type"
'---------------------------------------------------------------------
' Get the text that represents a MIMessage type
'---------------------------------------------------------------------

    GetMIMTypeText = modMIMsg.GetMIMTypeText(nType, bPlural)

End Function
 
'---------------------------------------------------------------------
Public Function GetScopeText(ByVal enScope As MIMsgScope) As String
'---------------------------------------------------------------------
' The scope of an MIMessage as a string
'---------------------------------------------------------------------

    GetScopeText = modMIMsg.GetScopeText(enScope)

End Function

'----------------------------------------------------------------------------------------'
Public Function IsDuplicateOCId(sCon As String, lOCId As Long, Optional sStudyName As String = "") As Boolean
'----------------------------------------------------------------------------------------'
'returns true if the specified OCId already exists in this study (or all studies if "")
'----------------------------------------------------------------------------------------'
Dim rs As Recordset
Dim oCon As Connection
Dim sSQL As String

    On Error GoTo ErrLabel
    
    Set oCon = New Connection
    oCon.Open sCon
    oCon.CursorLocation = adUseClient
    
    sSQL = "select Count(*) from MIMESSAGE where "
    If sStudyName <> "" Then
        sSQL = sSQL & " MIMESSAGETRIALNAME = '" & sStudyName & "'  and "
    End If
    sSQL = sSQL & " MIMESSAGEOCDISCREPANCYID = " & lOCId
    
    
    
    Set rs = New Recordset
    rs.Open sSQL, oCon
            
    IsDuplicateOCId = False
    If rs.Fields(0) > 0 Then
        IsDuplicateOCId = True
    End If
    rs.Close
    Set rs = Nothing

    oCon.Close
    Set oCon = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIMsgStatic.GetClinicalTrialId(" & sStudyName & ")"
    
End Function

'----------------------------------------------------------------------------------------'
Public Sub UpdateNoteStatusInDB(sCon As String, enScope As MIMsgScope, _
                                sStudyName As String, lClinicalTrialId As Long, sSite As String, lSubjectId As Long, _
                                Optional lVisitId As Long = 0, Optional nVisitCycle As Integer = 0, _
                                Optional lCRFPageTaskId As Long = 0, _
                                Optional lResponseTaskId As Long = 0, Optional nResponseCycle As Integer = 0)
Attribute UpdateNoteStatusInDB.VB_Description = "if -1 for lClinialtrialid then calculate from sStudyName"
'----------------------------------------------------------------------------------------'
'Update the MIMessageStatus in the subject data tables according the contents of the MIMessage table
'TA 12/02/2003 - if -1 for lClinialtrialid then calculate from sStudyName
'----------------------------------------------------------------------------------------'
Dim sWhereSQL As String
Dim oCon As ADODB.Connection

    Set oCon = New ADODB.Connection
    oCon.Open sCon
    oCon.BeginTrans
    
    If lClinicalTrialId = -1 Then
        'don't know trial id - calculate
        lClinicalTrialId = GetClinicalTrialId(oCon, sStudyName)
    End If
    
    sWhereSQL = " WHERE ClinicalTrialId = " & lClinicalTrialId _
                    & " AND TrialSite = '" & sSite & "'" _
                    & " AND PersonId = " & lSubjectId
                      

    Select Case enScope
    Case MIMsgScope.mimscSubject
        'update the subject data table
        oCon.Execute "Update TrialSubject set NoteStatus=1" & sWhereSQL
    Case MIMsgScope.mimscVisit
        'update the visit data table
        sWhereSQL = sWhereSQL & " AND VisitId=" & lVisitId & " AND VisitCycleNumber=" & nVisitCycle
        oCon.Execute "Update VisitInstance set NoteStatus=1" & sWhereSQL
    Case MIMsgScope.mimscEForm
        'update the eform data table
        sWhereSQL = sWhereSQL & " AND VisitId=" & lVisitId & " AND VisitCycleNumber=" & nVisitCycle
        sWhereSQL = sWhereSQL & " AND CRFPageTaskId=" & lCRFPageTaskId
        oCon.Execute "Update CRFPageInstance set NoteStatus=1" & sWhereSQL
    Case MIMsgScope.mimscQuestion
        'update the response data table
        sWhereSQL = sWhereSQL & " AND VisitId=" & lVisitId & " AND VisitCycleNumber=" & nVisitCycle
        sWhereSQL = sWhereSQL & " AND CRFPageTaskId=" & lCRFPageTaskId
        sWhereSQL = sWhereSQL & " AND ResponseTaskId=" & lResponseTaskId & " AND RepeatNumber=" & nResponseCycle
        oCon.Execute "Update DataItemResponse set NoteStatus=1" & sWhereSQL
    End Select
    
    oCon.CommitTrans
    oCon.Close
    Set oCon = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIMsgStatic.UpdateNoteStatusInDB"
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub SetImportedMIMsgStatus(ByRef oCon As ADODB.Connection, _
    ByRef enScope As MIMsgScope, ByRef sStudyName As String, _
    ByRef lClinicalTrialId As Long, ByRef sSite As String, ByRef lSubjectId As Long, _
    Optional ByRef lVisitId As Long = 0, Optional ByRef nVisitCycle As Integer = 0, _
    Optional ByRef lCRFPageTaskId As Long = 0, _
    Optional ByRef lResponseTaskId As Long = 0, Optional ByRef nResponseCycle As Integer = 0)
'----------------------------------------------------------------------------------------'
' MLM 29/09/06: Issue 2727: Added.
'----------------------------------------------------------------------------------------'
Dim sQuerySQL As String
Dim sUpdateSQL As String
Dim sTable As String
Dim rsNoteCount As ADODB.Recordset
Dim nNoteStatus

On Error GoTo ErrLabel

    'note
    Set rsNoteCount = New ADODB.Recordset
    sQuerySQL = "SELECT COUNT(*) FROM MIMessage WHERE MIMessageType = " & MIMsgType.mimtNote & _
        " AND MIMessageTrialName = '" & sStudyName & "' AND MIMessageSite = '" & sSite & "' AND MIMessagePersonId = " & lSubjectId
    sUpdateSQL = " WHERE ClinicalTrialId = " & lClinicalTrialId & " AND TrialSite = '" & sSite & "' AND PersonId = " & lSubjectId
    sTable = "TrialSubject"
    If lVisitId <> 0 Then
        sQuerySQL = sQuerySQL & " AND MIMessageVisitId = " & lVisitId & " AND MIMessageVisitCycle = " & nVisitCycle
        sUpdateSQL = sUpdateSQL & " AND VisitId = " & lVisitId & " AND VisitCycleNumber = " & nVisitCycle
        sTable = "VisitInstance"
    End If
    If lCRFPageTaskId <> 0 Then
        sQuerySQL = sQuerySQL & " AND MIMessageCRFPageTaskId = " & lCRFPageTaskId
        sUpdateSQL = sUpdateSQL & " AND CRFPageTaskId = " & lCRFPageTaskId
        sTable = "CRFPageInstance"
    End If
    If lResponseTaskId <> 0 Then
        sQuerySQL = sQuerySQL & " AND MIMessageResponseTaskId = " & lResponseTaskId & " AND MIMessageResponseCycle = " & nResponseCycle
        sUpdateSQL = sUpdateSQL & " AND ResponseTaskId = " & lResponseTaskId & " AND RepeatNumber = " & nResponseCycle
        sTable = "DataItemResponse"
    End If
    rsNoteCount.Open sQuerySQL, oCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsNoteCount.Fields(0).Value = 0 Then
        nNoteStatus = 0
    Else
        nNoteStatus = 1
    End If
    rsNoteCount.Close
    oCon.Execute "UPDATE " & sTable & " SET NoteStatus = " & nNoteStatus & sUpdateSQL
    
    'discrepancy
    Call SetIndividualStatus(oCon, enScope, MIMsgType.mimtDiscrepancy, sStudyName, lClinicalTrialId, _
        sSite, lSubjectId, lVisitId, nVisitCycle, lCRFPageTaskId, lResponseTaskId, nResponseCycle)
    'SDV
    Call SetIndividualStatus(oCon, enScope, MIMsgType.mimtSDVMark, sStudyName, lClinicalTrialId, _
        sSite, lSubjectId, lVisitId, nVisitCycle, lCRFPageTaskId, lResponseTaskId, nResponseCycle)

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIMsgStatic.SetImportedMIMsgStatus"

End Sub

'----------------------------------------------------------------------------------------'
Public Sub UpdateMIMsgStatusInDB(ByRef enSubjectStatus As Integer, ByRef enVisitStatus As Integer, _
                                        ByRef eneFormStatus As Integer, ByRef enResponseStatus As Integer, _
                                        sCon As String, enMIMsgType As MIMsgType, _
                                        sStudyName As String, lClinicalTrialId As Long, sSite As String, lSubjectId As Long, _
                                        Optional lVisitId As Long = 0, Optional nVisitCycle As Integer = 0, _
                                        Optional lCRFPageTaskId As Long = 0, _
                                        Optional lResponseTaskId As Long = 0, Optional nResponseCycle As Integer = 0)
Attribute UpdateMIMsgStatusInDB.VB_Description = "if -1 for lClinialtrialid then calculate from sStudyName"
'----------------------------------------------------------------------------------------'
' Update the MIMessageStatus in the subject data tables according the contents of the MIMessage table
'   and return the new statuses for subject/visit/eform/response
' NCJ 15 Oct 02 - Only update things for which we have info
' NCJ 21 Oct 02 - Beware of NULL statuses
'   TA following variables are returned byref to the calling function
' enSubjectStatus As MIMsgStatus
' enVisitStatus As MIMsgStatus
' eneFormStatus As MIMsgStatus
' enResponseStatus As MIMsgStatus
'TA 12/02/2003 - if -1 for lClinialtrialid then calcualte from sStudyName
' MLM 29/09/06: Issue 2727: Refactored to call new routine SetIndividualStatus
'----------------------------------------------------------------------------------------'
Dim sWhereSQL As String
Dim sStatusCol As String
Dim oCon As ADODB.Connection
'Dim bSQLServer As Boolean

    On Error GoTo ErrLabel

    'bSQLServer = (Connection_Property(CONNECTION_PROVIDER, sCon) = CONNECTION_SQLOLEDB)
    
    ' Initialise return values
    enSubjectStatus = 0
    eneFormStatus = 0
    eneFormStatus = 0
    enResponseStatus = 0
    
    Set oCon = New ADODB.Connection
    oCon.Open sCon
    'TA 24/08/2004: use client side cursor - it's better and allows a recordcount
    oCon.CursorLocation = adUseClient
    oCon.BeginTrans

    If lClinicalTrialId = -1 Then
        'don't know trial id - calculate
        lClinicalTrialId = GetClinicalTrialId(oCon, sStudyName)
    End If
        
    enSubjectStatus = SetIndividualStatus(oCon, MIMsgScope.mimscSubject, enMIMsgType, sStudyName, lClinicalTrialId, _
        sSite, lSubjectId)
    
    ' Only do Visit/eForm/Response updates if we have the relevant info
    ' NB SQL may return NULL if no MIMsgs for specified objects
    If lVisitId > 0 Then
        enVisitStatus = SetIndividualStatus(oCon, MIMsgScope.mimscVisit, enMIMsgType, sStudyName, lClinicalTrialId, _
            sSite, lSubjectId, lVisitId, nVisitCycle)
    
        If lCRFPageTaskId > 0 Then
            eneFormStatus = SetIndividualStatus(oCon, MIMsgScope.mimscEForm, enMIMsgType, sStudyName, lClinicalTrialId, _
                sSite, lSubjectId, lVisitId, nVisitCycle, lCRFPageTaskId)
                
            If lResponseTaskId > 0 Then
                enResponseStatus = SetIndividualStatus(oCon, MIMsgScope.mimscQuestion, enMIMsgType, sStudyName, lClinicalTrialId, _
                    sSite, lSubjectId, lVisitId, nVisitCycle, lCRFPageTaskId, lResponseTaskId, nResponseCycle)
            End If
        End If
    End If
    
    
    oCon.CommitTrans
    oCon.Close
    Set oCon = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIMsgStatic.MIMsgStatusInDB"
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function SetIndividualStatus(ByRef oCon As ADODB.Connection, ByRef enScope As MIMsgScope, _
    ByRef enMIMsgType As MIMsgType, ByRef sStudyName As String, _
    ByRef lClinicalTrialId As Long, ByRef sSite As String, ByRef lSubjectId As Long, _
    Optional ByRef lVisitId As Long = 0, Optional ByRef nVisitCycle As Integer = 0, _
    Optional ByRef lCRFPageTaskId As Long = 0, _
    Optional ByRef lResponseTaskId As Long = 0, Optional ByRef nResponseCycle As Integer = 0) _
    As Integer
'----------------------------------------------------------------------------------------'
' MLM 22/08/06: Created from UpdateMIMsgStatusInDb. Set the SDVStatus or DiscrepancyStatus
'   of one TrialSubject, VisitInstance, CRFPageInstance or DataItemResponse row.
'----------------------------------------------------------------------------------------'

Dim sStatusCol As String
Dim sWhereSQL As String
Dim sTable As String
Dim nStatus As Integer
Dim bSQLServer As Boolean

On Error GoTo ErrLabel

    'MLM 06/11/06: Use Instr to cope with SQL Server connection strings extracted from connection objects
    bSQLServer = (InStr(Connection_Property(CONNECTION_PROVIDER, oCon.ConnectionString), CONNECTION_SQLOLEDB) > 0)

    'decide on the column to be updated
    Select Case enMIMsgType
        Case MIMsgType.mimtDiscrepancy
            sStatusCol = "DiscrepancyStatus"
        Case MIMsgType.mimtSDVMark
            sStatusCol = "SDVStatus"
    End Select
    
    'decide on the rest of the SQL for updating the status
    'subject
    sWhereSQL = " WHERE ClinicalTrialId = " & lClinicalTrialId _
                & " AND TrialSite = '" & sSite & "'" _
                & " AND PersonId = " & lSubjectId
    sTable = "TrialSubject"
    'visit
    If enScope >= mimscVisit Then
        sTable = "VisitInstance"
        sWhereSQL = sWhereSQL & " AND VisitId=" & lVisitId & " AND VisitCycleNumber=" & nVisitCycle
    End If
    'eForm
    If enScope >= mimscEForm Then
        sTable = "CRFPageInstance"
        sWhereSQL = sWhereSQL & " AND CRFPageTaskId=" & lCRFPageTaskId
    End If
    'question
    If enScope = mimscQuestion Then
        sTable = "DataItemResponse"
        sWhereSQL = sWhereSQL & " AND ResponseTaskId=" & lResponseTaskId & " AND RepeatNumber=" & nResponseCycle
    End If

    nStatus = GetMIMsgHierachicalStatus(GetStatusSQL(enMIMsgType, bSQLServer, sStudyName, sSite, lSubjectId, _
                                    lVisitId, nVisitCycle, lCRFPageTaskId, _
                                    lResponseTaskId, nResponseCycle), oCon, enMIMsgType, enScope)
    oCon.Execute "UPDATE " & sTable & " SET " & sStatusCol & " = " & _
        nStatus & sWhereSQL
    SetIndividualStatus = nStatus

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIMsgStatic.SetIndividualStatus"
    

End Function

'----------------------------------------------------------------------------------------'
Public Function CreateDecodeSQL(bSQLServer As Boolean, sColumnName As String, _
                                    asOriginalValues As Variant, asNewValues As Variant, sElse As String) As String
'----------------------------------------------------------------------------------------'
'Create the Decode/Case SQL according to databases type
'values are sql string. ie put single quotes around a varchar datatype
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim i As Long

    On Error GoTo ErrLabel

    sSQL = "CASE " & sColumnName
    
    For i = 0 To UBound(asOriginalValues)
        sSQL = sSQL & " WHEN " & asOriginalValues(i) & " THEN " & asNewValues(i)
    Next
    sSQL = sSQL & " ELSE " & sElse & " END"

    If Not bSQLServer Then
        sSQL = Replace(sSQL, "CASE ", "DECODE(")
        sSQL = Replace(sSQL, " WHEN ", ",")
        sSQL = Replace(sSQL, " THEN ", ",")
        sSQL = Replace(sSQL, " ELSE ", ",")
        sSQL = Replace(sSQL, " END", ")")
    End If
    
    CreateDecodeSQL = sSQL
Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIMsgStatic.CreateDecodeSQL"
    
End Function
                                    
'----------------------------------------------------------------------------------------'
Public Function GetStatusSQL(enMIMsgType As MIMsgType, bSQLServer As Boolean, _
                                sStudyName As String, sSite As String, lSubjectId As Long, _
                                Optional lVisitId As Long = 0, Optional nVisitCycle As Integer = 0, _
                                Optional lCRFPageTaskId As Long = 0, _
                                Optional lResponseTaskId As Long = 0, Optional nResponseCycle As Integer = 0) As String
'----------------------------------------------------------------------------------------'
' Get the SQL to retrieve the current MIMessageStatus for the specified Subject/Visit/EForm or Response
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim vOld As Variant
Dim vNew As Variant
Dim sElse As String


    On Error GoTo ErrLabel

    sElse = "0"
    
    ' Convert the MIMessage table statuses to the DEBS30 hierarchical statuses
    Select Case enMIMsgType
    Case MIMsgType.mimtDiscrepancy
        ' raised, responded, closed
        vOld = Array("0", "1", "2")
        vNew = Array("30", "20", "10")
    Case MIMsgType.mimtSDVMark
        ' planned , complete(done)
        ' NCJ 14 Oct 02 - Planned, Queried, Done, Cancelled
        vOld = Array("0", "1", "2", "3")
        vNew = Array("30", "40", "20", "10")
        
    End Select

    ' get decode/case sql
    sSQL = "SELECT "
    
    'TA 24/08/2004 'group by' sql for non question specific SDVs
    If enMIMsgType = MIMsgType.mimtSDVMark Then
        If GetScope(lVisitId, lCRFPageTaskId, lResponseTaskId) <> mimscQuestion Then
            'question SDV can't inherit from below
            sSQL = sSQL & GetSDVDecodeSQL(bSQLServer, GetScope(lVisitId, lCRFPageTaskId, lResponseTaskId)) & ","
        End If
    End If
    sSQL = sSQL & "MAX(" _
                        & CreateDecodeSQL(bSQLServer, "MIMessageStatus", vOld, vNew, sElse) _
                        & ") FROM MIMESSAGE WHERE MIMessageType = " & enMIMsgType _
                        & " AND MIMessageHistory = 0"
                        
    'limit to this subject
    sSQL = sSQL & " AND MIMessageTrialName = '" & sStudyName & "'" _
            & " AND MIMessageSite = '" & sSite & "'" _
            & " AND MIMessagePersonId = " & lSubjectId
            
            
    If lVisitId <> 0 Then
        'if we are dealing with a visit/eform/response then add those columns
        sSQL = sSQL & " AND MIMessageVisitId = " & lVisitId & " AND MIMessageVisitCycle = " & nVisitCycle
        If lCRFPageTaskId <> 0 Then
            'if we are dealing with a eform/response then add those columns
            sSQL = sSQL & " AND MIMessageCRFPageTaskId = " & lCRFPageTaskId
            If lResponseTaskId <> 0 Then
                'if we are dealing with a response add those columns
                sSQL = sSQL & " AND MIMessageResponseTaskId = " & lResponseTaskId & " AND MIMessageResponseCycle = " & nResponseCycle
            End If
        End If
    End If

'TA 24/08/2004 'group by' sql for non question specific SDVs
    If enMIMsgType = MIMsgType.mimtSDVMark Then
        If GetScope(lVisitId, lCRFPageTaskId, lResponseTaskId) <> mimscQuestion Then
            'only subject, visit, eForm SDV can inherit from below
            sSQL = sSQL & " group by " & GetSDVDecodeSQL(bSQLServer, GetScope(lVisitId, lCRFPageTaskId, lResponseTaskId))
        End If
    End If

    GetStatusSQL = sSQL
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIMsgStatic.GetStatusSQL"
    
End Function

'----------------------------------------------------------------------------------------'
Public Function GetSDVDecodeSQL(bSQLServer As Boolean, enScope As MIMsgScope) As String
'----------------------------------------------------------------------------------------'
'TA 24/08/2004 - Return decode sql according to scope
'   the string 'this' will be returned for SDVs matching enScope
'   the string 'inherited' will be returned for SDV for lower scopes
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel
    GetSDVDecodeSQL = CreateDecodeSQL(bSQLServer, "MIMESSAGESCOPE", Array(CStr(enScope)), Array("'this'"), "'inherited'")
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIMsgStatic.GetSDVDecodeSQL"

End Function


'----------------------------------------------------------------------------------------'
Private Function GetScope(lVisitId As Long, lCRFPageTaskId As Long, lResponseTaskId As Long) As MIMsgScope
'----------------------------------------------------------------------------------------'
'TA 24/08/2004 - Return Scope of a MIMessage
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel
    If lVisitId = 0 Then GetScope = mimscSubject: Exit Function
    If lCRFPageTaskId = 0 Then GetScope = mimscVisit: Exit Function
    If lResponseTaskId = 0 Then GetScope = mimscEForm: Exit Function
    'must be mimessage on question
    GetScope = mimscQuestion
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIMsgStatic.GetScope"
    
End Function

'----------------------------------------------------------------------------------------'
Private Function GetMIMsgHierachicalStatus(sSQL As String, oCon As Connection, enMIMsgType As MIMsgType, enScope As MIMsgScope) As Integer
'----------------------------------------------------------------------------------------'
'Get Hierachical MIMsgs status from recordset
'----------------------------------------------------------------------------------------'
Dim enStatus As Integer
Dim nThis As Integer
Dim nInherited As Integer
Dim rs As Recordset

    On Error GoTo ErrLabel
    enStatus = 0
    Set rs = New ADODB.Recordset
    rs.Open sSQL, oCon
    If enMIMsgType <> mimtSDVMark Or enScope = mimscQuestion Then
        'not SDV or question SDV just return result or 0 if null
        'the record count will never be 0 for these
        If IsNull(rs.Fields(0).Value) Then
            enStatus = 0
        Else
            enStatus = rs.Fields(0).Value
        End If
    Else
        'SDV on subject, visit, or eForm
        If rs.RecordCount = 0 Then
            'no SDVs - return 0
            enStatus = 0
        Else
            nThis = 0
            nInherited = 0
            'loop through picking up values for "this" and "inherited" - they will appear 0 or 1 times each
            Do While Not rs.EOF
                Select Case rs.Fields(0).Value
                Case "this": nThis = rs.Fields(1).Value
                Case "inherited": nInherited = rs.Fields(1).Value
                Case Else: Err.Raise vbObjectError + 2001, , "SDVstatus recordset not return 'this' and 'inherited'" 'this will never happen
                End Select
                rs.MoveNext
            Loop
            'decide whether to inherit or not
            If (nInherited = 0) Or (nInherited = 10) Or (nInherited = 20) Then
                'none, cancelled or done - not inherited up - use own status
                enStatus = nThis
            Else
                'planned or queried - inherited up - take max status of own and inherited
                enStatus = Max((nThis), (nInherited))
            End If
        End If
    End If
    rs.Close
    Set rs = Nothing
    GetMIMsgHierachicalStatus = enStatus
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIMsgStatic.GetMIMsgHierachicalStatus"

End Function

'----------------------------------------------------------------------------------------'
Private Function GetClinicalTrialId(oCon As Connection, sStudyName As String) As Long
'----------------------------------------------------------------------------------------'
'calcualte clinicaltrialid from trialname
'----------------------------------------------------------------------------------------'
Dim rs As Recordset

  

    On Error GoTo ErrLabel
    Set rs = New Recordset
    rs.Open "select CLINICALTRIALID from CLINICALTRIAL where CLINICALTRIALNAME = '" & sStudyName & "'", oCon
    GetClinicalTrialId = rs.Fields(0)
    rs.Close
    Set rs = Nothing


Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIMsgStatic.GetClinicalTrialId(" & sStudyName & ")"
    
End Function


