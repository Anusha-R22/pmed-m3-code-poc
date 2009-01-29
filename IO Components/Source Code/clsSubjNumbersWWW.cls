VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubjNumbersWWW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'----------------------------------------------------------------------------------------'
'   File:       clsSubjNumbersWWW.cls
'   Copyright:  InferMed Ltd. 2003. All Rights Reserved
'   Author:     Nicky Johns, June 2003
'   Purpose:    Subject Numbering class for MACRO WWW
'               Based on clsSubjectNumbering from Windows DM
'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 11-12 Jun 03 - Initial development, based on copy of clsSubjectNumbering from Windows DM
'----------------------------------------------------------------------------------------'

Option Explicit

Private Const m_COL_CLINICALTRIALID = 1
Private Const m_COL_VERSIOINID = 2
Private Const m_COL_STARTNUMBER = 3
Private Const m_COL_NUMBERWIDTH = 4
Private Const m_COL_PREFIX = 5
Private Const m_COL_USEPREFIX = 6
Private Const m_COL_SUFFIX = 7
Private Const m_COL_USESUFFIX = 8
Private Const m_COL_TRIGGERVISITID = 9
Private Const m_COL_TRIGGERFORMID = 10
Private Const m_COL_USEREGISTRATION = 11

Private mlClinicalTrialId As Long
Private mnVersionId As Integer

Private mlStartNumber As Long
Private mnNumberWidth As Long
Private msPrefix As String
Private mnUsePrefix As Integer
Private msSuffix As String
Private mnUseSuffix As Integer
Private mlTriggerFormId As Long
Private mlTriggerVisitId As Long
Private mnUseRegistration As Integer

Private mtblRegCond As clsDataTable
Private mtblUniqueCheck As clsDataTable

'----------------------------------------------------------------------------------------'
Public Sub Load(ByVal sConnection As String, ByVal lClinicalTrialId As Long, ByVal nVersionId As Integer)
'----------------------------------------------------------------------------------------'
' Load data from SubjectNumbering DB table for this study
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim tblReg As clsDataTable

    On Error GoTo ErrHandler
    
    mlClinicalTrialId = lClinicalTrialId
    mnVersionId = nVersionId

    sSQL = "SELECT * FROM SubjectNumbering" _
                & " WHERE ClinicalTrialId = " & lClinicalTrialId _
                & " AND VersionId = " & nVersionId
    Set tblReg = TableFromSQL(sConnection, sSQL)
    Select Case tblReg.Rows
    Case 0
        'no registration details
    Case 1
        'details found
        mlStartNumber = tblReg(1, m_COL_STARTNUMBER)
        mnNumberWidth = tblReg(1, m_COL_NUMBERWIDTH)
        msPrefix = tblReg(1, m_COL_PREFIX)
        mnUsePrefix = tblReg(1, m_COL_USEPREFIX)
        msSuffix = tblReg(1, m_COL_SUFFIX)
        mnUseSuffix = tblReg(1, m_COL_USESUFFIX)
        mlTriggerFormId = tblReg(1, m_COL_TRIGGERFORMID)
        mlTriggerVisitId = tblReg(1, m_COL_TRIGGERVISITID)
        mnUseRegistration = tblReg(1, m_COL_USEREGISTRATION)


        sSQL = "SELECT EligibilityCode, Condition FROM Eligibility" _
                    & " WHERE ClinicalTrialId = " & lClinicalTrialId _
                    & " AND VersionId = " & nVersionId _
                    & " AND RandomisationCode IS NULL"
        Set mtblRegCond = TableFromSQL(sConnection, sSQL, RecordBuild("Code", "Condition"))
                            
        sSQL = "SELECT CheckCode, Expression FROM UniquenessCheck" _
                    & " WHERE ClinicalTrialId = " & lClinicalTrialId _
                    & " AND VersionId = " & nVersionId
        Set mtblUniqueCheck = TableFromSQL(sConnection, sSQL, RecordBuild("Code", "Expression"))
        
    Case Else
        'if this ever happens we have a corrupt database
    End Select
    
    
Exit Sub

ErrHandler:
  Err.Raise Err.Number, , Err.Description & "|clsSubjNumbersWWW.Load"
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'
'set up default values
'----------------------------------------------------------------------------------------'
    
    Set mtblRegCond = New clsDataTable
    Set mtblRegCond.Headings = RecordBuild("Code", "Condition")

    Set mtblUniqueCheck = New clsDataTable
    Set mtblUniqueCheck.Headings = RecordBuild("Code", "Expression")
    
    mlStartNumber = 1
    mnNumberWidth = 0
    msPrefix = ""
    mnUsePrefix = 0
    msSuffix = ""
    mnUseSuffix = 0
    mlTriggerFormId = 0
    mlTriggerVisitId = 0
    mnUseRegistration = 0
        
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get StartNumber() As Long
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    StartNumber = mlStartNumber

End Property

'----------------------------------------------------------------------------------------'
Public Property Get NumberWidth() As Integer
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    NumberWidth = mnNumberWidth

End Property

'----------------------------------------------------------------------------------------'
Public Property Get Prefix() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Prefix = msPrefix
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get UsePrefix() As Boolean
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    UsePrefix = (mnUsePrefix = 1)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get Suffix() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Suffix = msSuffix
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get UseSuffix() As Boolean
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    UseSuffix = (mnUseSuffix = 1)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get TriggerFormId() As Long
'----------------------------------------------------------------------------------------'
' The eForm ID after which registration happens
'----------------------------------------------------------------------------------------'

    TriggerFormId = mlTriggerFormId
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get TriggerVisitId() As Long
'----------------------------------------------------------------------------------------'
' The visit ID containing the eForm after which registration happens
'----------------------------------------------------------------------------------------'

    TriggerVisitId = mlTriggerVisitId
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get UseRegistration() As Boolean
'----------------------------------------------------------------------------------------'

    UseRegistration = (mnUseRegistration = 1)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get RegistrationConditions() As clsDataTable
'----------------------------------------------------------------------------------------'

    Set RegistrationConditions = mtblRegCond
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get UniquenessChecks() As clsDataTable
'----------------------------------------------------------------------------------------'

    Set UniquenessChecks = mtblUniqueCheck
    
End Property

'----------------------------------------------------------------------------------------'
Private Function TableFromSQL(sConnection As String, _
                                sSQL As String, _
                                Optional recHeadings As Variant = Empty) As clsDataTable
'----------------------------------------------------------------------------------------'
'Return a data table from SQL
'Input: optional RecHeadings - a data record of headings for each column,
'               if nothing is passed through the db column names will be used
'----------------------------------------------------------------------------------------'
Dim rs As ADODB.Recordset
Dim tblTable As clsDataTable
Dim recRow As clsDataRecord
Dim lCols As Long
Dim i As Long
Dim oDBCon As ADODB.Connection

    On Error GoTo ErrHandler
    
    Set oDBCon = New ADODB.Connection
    Call oDBCon.Open(sConnection)

    Set rs = New ADODB.Recordset
    rs.Open sSQL, oDBCon
    lCols = rs.Fields.Count
    Set tblTable = New clsDataTable
    
    If IsEmpty(recHeadings) Then
        Set tblTable.Headings = New clsDataRecord
        tblTable.Headings.Init lCols
        For i = 1 To lCols
            tblTable.Headings.Field(i) = rs.Fields(i - 1).Name
        Next
    Else
        Set tblTable.Headings = recHeadings
    End If
    
    Do While Not rs.EOF
        Set recRow = New clsDataRecord
        recRow.Init lCols
        For i = 1 To lCols
            recRow(i) = FieldToString(rs.Fields(i - 1).Value)
        Next
        tblTable.Add recRow
        Set recRow = Nothing
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing

    Set TableFromSQL = tblTable

Exit Function
ErrHandler:
  Err.Raise Err.Number, , Err.Description & "|clsSubjNumbersWWW.TableFromSQL"
    
End Function

'----------------------------------------------------------------------------------------'
Private Function RecordBuild(ParamArray vFields() As Variant) As clsDataRecord
'----------------------------------------------------------------------------------------'
'return a record from a list of parameter strings
'----------------------------------------------------------------------------------------'
Dim recRecord As clsDataRecord
Dim i As Long

    Set recRecord = New clsDataRecord
    
    recRecord.Init UBound(vFields) + 1
    
    For i = 0 To UBound(vFields)
        recRecord(i + 1) = vFields(i)
    Next

    Set RecordBuild = recRecord
    Set recRecord = Nothing
    
End Function

'----------------------------------------------------------------------------------------'
Private Function FieldToString(vValue As Variant)
'----------------------------------------------------------------------------------------'
'converts a variant to ""
'----------------------------------------------------------------------------------------'
    
    If VarType(vValue) = vbNull Then
        FieldToString = ""
    Else
        FieldToString = Format(vValue)
    End If

End Function

