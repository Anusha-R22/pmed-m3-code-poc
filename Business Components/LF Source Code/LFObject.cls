VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LFObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   File:       LFObject.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Nicky Johns, December 2002
'   Purpose:    Represents an object specification for MACRO 3.0
'               i.e. a Question, EForm, Visit or Subject
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 10-12 Dec 02 - Initial development
'   NCJ 3 Jan 03 - Added QuestionId
'   NCJ 8 Jan 03 - Make sure SQL strings are empty if not in scope; added SQLObjectWhere
'----------------------------------------------------------------------------------------'

Option Explicit

Public StudyId As Long
Public Site As String
Public SubjectId As Long

Public Scope As LFScope

' Specify the object (values of 0 should be used for unused slots)
Public VisitId As Long
Public VisitCycle As Integer
Public EFormId As Long
Public EFormCycle As Integer
Public ResponseId As Long
Public ResponseCycle As Integer
Public QuestionId As Long

'--------------------------------------------------------------
Public Sub LoadFromRS(rsLFObj As ADODB.Recordset, enScope As LFScope)
'--------------------------------------------------------------
' Load all values from recordset of values
'--------------------------------------------------------------
    
    Scope = enScope
    
    StudyId = rsLFObj!ClinicalTrialId
    Site = rsLFObj!TrialSite
    SubjectId = rsLFObj!PersonId

    ' Set Visit details for Visits, EForms and Questions
    If enScope >= LFScope.lfscVisit Then
        VisitId = rsLFObj!VisitId
        VisitCycle = rsLFObj!VisitCycleNumber
    Else
        VisitId = 0
        VisitCycle = 0
    End If
    
    ' Set EForm details for EForms and Questions
    If enScope >= LFScope.lfscEForm Then
        EFormId = rsLFObj!CRFPageId
        EFormCycle = rsLFObj!CRFPageCycleNumber
    Else
        EFormId = 0
        EFormCycle = 0
    End If

    ' Set Question details for Questions
    If enScope = LFScope.lfscQuestion Then
        ResponseId = rsLFObj!ResponseTaskId
        ResponseCycle = rsLFObj!RepeatNumber
        QuestionId = rsLFObj!DataItemId
    Else
        ResponseId = 0
        ResponseCycle = 0
        QuestionId = 0
    End If

End Sub

'-------------------------------------------------------------------
Public Sub Init(enScope As LFScope, _
                lStudyId As Long, sSite As String, lSubjectId As Long, _
                Optional lVisitId As Long = 0, Optional nVisitCycle As Integer = 0, _
                Optional lEFormId As Long = 0, Optional nEformCycle As Integer = 0, _
                Optional lResponseId As Long = 0, Optional nResponseCycle As Integer = 0, _
                Optional lQuestionId As Long = 0)
'-------------------------------------------------------------------
' Initialise with explicit values
' Assume we haven't got values for non-required parameters,
' e.g. if Scope is Visit, there are no eForm or Question details
'-------------------------------------------------------------------
               
    Scope = enScope
    StudyId = lStudyId
    Site = sSite
    SubjectId = lSubjectId
    VisitId = lVisitId
    VisitCycle = nVisitCycle
    EFormId = lEFormId
    EFormCycle = nEformCycle
    ResponseId = lResponseId
    ResponseCycle = nResponseCycle
    QuestionId = lQuestionId
              
End Sub

'-------------------------------------------------------------------
Public Property Get SQLSubjectWhere() As String
'-------------------------------------------------------------------
' The SQL WHERE clause for this object's Subject
'-------------------------------------------------------------------
    
    SQLSubjectWhere = " WHERE ClinicalTrialId = " & StudyId & _
                    " AND TrialSite = '" & Site & "'" & _
                    " AND PersonId = " & SubjectId

End Property

'-------------------------------------------------------------------
Public Property Get SQLAndVisitWhere() As String
'-------------------------------------------------------------------
' The additional SQL WHERE clause for this object's Visit
'-------------------------------------------------------------------
    
    If Scope >= LFScope.lfscVisit Then
        SQLAndVisitWhere = " AND VisitId = " & VisitId & _
                    " AND VisitCycleNumber = " & VisitCycle
    Else
        SQLAndVisitWhere = ""
    End If

End Property

'-------------------------------------------------------------------
Public Property Get SQLAndEFormWhere() As String
'-------------------------------------------------------------------
' The additional SQL WHERE clause for this object's EForm
'-------------------------------------------------------------------
    
    If Scope >= LFScope.lfscEForm Then
        SQLAndEFormWhere = " AND CRFPageId = " & EFormId & _
                    " AND CRFPageCycleNumber = " & EFormCycle
    Else
        SQLAndEFormWhere = ""
    End If

End Property

'-------------------------------------------------------------------
Public Property Get SQLAndQuestionWhere() As String
'-------------------------------------------------------------------
' The additional SQL WHERE clause for this object's Question
'-------------------------------------------------------------------
    
    If Scope = LFScope.lfscQuestion Then
        SQLAndQuestionWhere = " AND ResponseTaskId = " & ResponseId & _
                    " AND RepeatNumber = " & ResponseCycle
    Else
        SQLAndQuestionWhere = ""
    End If

End Property

'-------------------------------------------------------------------
Public Property Get SQLObjectWhere() As String
'-------------------------------------------------------------------
' The SQL WHERE clause that identifies this object
' according to its scope
'-------------------------------------------------------------------

    ' NB The Visit, Form and Question WHERE's are empty if object not in their scope
    SQLObjectWhere = SQLSubjectWhere & SQLAndVisitWhere _
                    & SQLAndEFormWhere & SQLAndQuestionWhere

End Property

'--------------------------------------------------------------
Private Sub Class_Initialize()
'--------------------------------------------------------------

    ' Default to Subject scope
    Scope = LFScope.lfscSubject
    VisitId = 0
    VisitCycle = 0
    EFormId = 0
    EFormCycle = 0
    ResponseId = 0
    ResponseCycle = 0
    QuestionId = 0
    
End Sub
