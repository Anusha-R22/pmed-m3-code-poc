VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubjectNumbering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'----------------------------------------------------------------------------------------'
'   File:       clsSubjectNumbering.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, November 2000
'   Purpose:    Subject Numbering class for MACRO DM, SD
'               This is intended to be used as "read-only"
'               Do NOT set individual values in here!!! (See clsRegistration instead)
'----------------------------------------------------------------------------------------'
' REVISIONS:
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
Attribute mtblRegCond.VB_VarHelpID = -1
Private mtblUniqueCheck As clsDataTable
Attribute mtblUniqueCheck.VB_VarHelpID = -1

'updating or inserting when save occurs
Private mbUpdate As Boolean


'----------------------------------------------------------------------------------------'
Public Sub Load(lClinicalTrialId As Long, nVersionId As Integer)
'----------------------------------------------------------------------------------------'
' Load data from SubjectNumbering DB table for this trial
'----------------------------------------------------------------------------------------'

Dim sSQL As String
Dim tblReg As clsDataTable

    On Error GoTo ErrHandler
    
    mlClinicalTrialId = lClinicalTrialId
    mnVersionId = nVersionId
        
        
    sSQL = "SELECT * FROM SubjectNumbering" _
                & " WHERE ClinicalTrialId = " & lClinicalTrialId _
                & " AND VersionId = " & nVersionId
    Set tblReg = TableFromSQL(sSQL, , True)
    Select Case tblReg.Rows
    Case 0
        'no registration details
        mbUpdate = False
    Case 1
        'details found
        mbUpdate = True
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
        Set mtblRegCond = TableFromSQL(sSQL, RecordBuild("Code", "Condition"))
                            
        sSQL = "SELECT CheckCode, Expression FROM UniquenessCheck" _
                    & " WHERE ClinicalTrialId = " & lClinicalTrialId _
                    & " AND VersionId = " & nVersionId
        Set mtblUniqueCheck = TableFromSQL(sSQL, RecordBuild("Code", "Expression"))
        
    Case Else
        'if this ever happens we have a corrupt database
    End Select
    
    
Exit Sub

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Load", "clsSubjectNumbering")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Save()
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim i As Long
' TA 23/11/2000 all arezzos are trimmed

    On Error GoTo ErrHandler
    
    If mbUpdate Then
        sSQL = "UPDATE SubjectNumbering SET" _
                    & " StartNumber = " & mlStartNumber _
                    & ", NumberWidth = " & mnNumberWidth _
                    & ", Prefix = '" & ReplaceQuotes(Trim(msPrefix)) & "'" _
                    & ", UsePrefix = " & mnUsePrefix _
                    & ", Suffix = '" & ReplaceQuotes(Trim(msSuffix)) & "'" _
                    & ", UseSuffix = " & mnUseSuffix _
                    & ", TriggerVisitId = " & mlTriggerVisitId _
                    & ", TriggerFormId = " & mlTriggerFormId _
                    & ", UseRegistration = " & mnUseRegistration _
                    & " WHERE ClinicalTrialId = " & mlClinicalTrialId _
                    & " AND VersionId = " & mnVersionId
    Else
        sSQL = "INSERT INTO SubjectNumbering VALUES (" _
                    & mlClinicalTrialId _
                    & "," & mnVersionId _
                    & "," & mlStartNumber _
                    & ", " & mnNumberWidth _
                    & ", '" & ReplaceQuotes(Trim(msPrefix)) & "'" _
                    & ", " & mnUsePrefix _
                    & ", '" & ReplaceQuotes(Trim(msSuffix)) & "'" _
                    & ", " & mnUseSuffix _
                    & ", " & mlTriggerVisitId _
                    & ", " & mlTriggerFormId _
                    & ", " & mnUseRegistration & ")"
    End If
    MacroADODBConnection.Execute sSQL
    
    
    'registration conditions
    sSQL = "DELETE FROM Eligibility" _
                    & " WHERE ClinicalTrialId = " & mlClinicalTrialId _
                    & " AND VersionId = " & mnVersionId _
                    & " AND RandomisationCode IS NULL"
    MacroADODBConnection.Execute sSQL
    For i = 1 To mtblRegCond.Rows
        'note the flag is always set to 1 (eligibility condition) currently
        sSQL = mlClinicalTrialId & ", " & mnVersionId & ", '" & mtblRegCond(i, 1) & "', null, 1, '" & ReplaceQuotes(Trim(mtblRegCond(i, 2))) & "'"
        MacroADODBConnection.Execute "INSERT INTO Eligibility VALUES (" & sSQL & ")"
    Next
    
    sSQL = "DELETE FROM UniquenessCheck" _
           & " WHERE ClinicalTrialId = " & mlClinicalTrialId _
           & " AND VersionId = " & mnVersionId
        MacroADODBConnection.Execute sSQL
    For i = 1 To mtblUniqueCheck.Rows
        sSQL = mlClinicalTrialId & ", " & mnVersionId & ", '" & mtblUniqueCheck(i, 1) & "',  '" & ReplaceQuotes(Trim(mtblUniqueCheck(i, 2))) & "'"
        MacroADODBConnection.Execute "INSERT INTO UniquenessCheck VALUES (" & sSQL & ")"
    Next
                    
   
Exit Sub

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Save", "clsSubjectNumbering")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
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
Public Property Let StartNumber(lNumber As Long)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    mlStartNumber = lNumber
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get NumberWidth() As Integer
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    NumberWidth = mnNumberWidth
    

End Property

'----------------------------------------------------------------------------------------'
Public Property Let NumberWidth(nNumber As Integer)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    mnNumberWidth = nNumber
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get Prefix() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Prefix = msPrefix
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let Prefix(sPrefix As String)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    msPrefix = sPrefix
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get UsePrefix() As Boolean
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    UsePrefix = (mnUsePrefix = 1)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let UsePrefix(bUse As Boolean)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    If bUse Then
        mnUsePrefix = 1
    Else
        mnUsePrefix = 0
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get Suffix() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Suffix = msSuffix
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let Suffix(sSuffix As String)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    msSuffix = sSuffix
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get UseSuffix() As Boolean
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    UseSuffix = (mnUseSuffix = 1)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let UseSuffix(bUse As Boolean)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
    
    If bUse Then
        mnUseSuffix = 1
    Else
        mnUseSuffix = 0
    End If

End Property

'----------------------------------------------------------------------------------------'
Public Property Get TriggerFormId() As Long
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    TriggerFormId = mlTriggerFormId
    
End Property

Public Property Let TriggerFormId(lId As Long)

    mlTriggerFormId = lId

End Property

Public Property Get TriggerVisitId() As Long

    TriggerVisitId = mlTriggerVisitId
    
End Property

Public Property Let TriggerVisitId(lId As Long)

    mlTriggerVisitId = lId

End Property

Public Property Get UseRegistration() As Boolean

    UseRegistration = (mnUseRegistration = 1)
    
End Property

Public Property Let UseRegistration(bUse As Boolean)

    If bUse Then
        mnUseRegistration = 1
    Else
        mnUseRegistration = 0
    End If
    
End Property


Public Property Get RegistrationConditions() As clsDataTable

    Set RegistrationConditions = mtblRegCond
    
End Property


Public Property Set RegistrationConditions(tblRegistrationConditions As clsDataTable)

    Set mtblRegCond = tblRegistrationConditions
    
End Property

Public Property Get UniquenessChecks() As clsDataTable

    Set UniquenessChecks = mtblUniqueCheck
    
End Property

Public Property Set UniquenessChecks(tblUniquenessChecks As clsDataTable)

    Set mtblUniqueCheck = tblUniquenessChecks

End Property


