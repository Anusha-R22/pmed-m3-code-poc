VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QuestionRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = ""
Attribute VB_Description = ""
Attribute VB_Description = ""
Attribute VB_Description = ""
Attribute VB_Description = ""
'----------------------------------------------------
' File: QuestionRO.cls
' Author: Nicky Johns
' Copyright: InferMed, June 2001, All Rights Reserved
' Question def Read-only Object for MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 15 Jun 01 - Initial development
'----------------------------------------------------

Option Explicit

Private moQuestionState As QuestionState

'----------------------------------------------------
Public Property Get Name() As String
Attribute Name.VB_Description = "Name of question"
'----------------------------------------------------
' Name of question
'----------------------------------------------------

    Name = moQuestionState.Name

End Property

'----------------------------------------------------
Public Property Get Code() As String
Attribute Code.VB_Description = "Code of question"
'----------------------------------------------------
' Code of question
'----------------------------------------------------

    Code = moQuestionState.Code

End Property

'----------------------------------------------------
Public Property Get QuestionId() As Long
Attribute QuestionId.VB_Description = ""
'----------------------------------------------------
'----------------------------------------------------

    QuestionId = moQuestionState.QuestionId

End Property

'----------------------------------------------------
Public Property Get DataType() As Integer
Attribute DataType.VB_Description = ""
'----------------------------------------------------
'----------------------------------------------------

    DataType = moQuestionState.DataType

End Property

'----------------------------------------------------
Public Property Get QuestionLength() As Integer
Attribute QuestionLength.VB_Description = ""
'----------------------------------------------------
'----------------------------------------------------

    QuestionLength = moQuestionState.QuestionLength

End Property

'----------------------------------------------------
Public Property Get Unit() As String
Attribute Unit.VB_Description = ""
'----------------------------------------------------
'----------------------------------------------------

    Unit = moQuestionState.Unit

End Property

'----------------------------------------------------
Public Property Get Format() As String
Attribute Format.VB_Description = ""
'----------------------------------------------------
'----------------------------------------------------

    Format = moQuestionState.Format

End Property

'----------------------------------------------------
Public Property Get Helptext() As String
Attribute Helptext.VB_Description = ""
'----------------------------------------------------
'----------------------------------------------------

    Helptext = moQuestionState.Helptext

End Property

'----------------------------------------------------
Public Property Get TextCase() As Integer
Attribute TextCase.VB_Description = ""
'----------------------------------------------------
'----------------------------------------------------

    TextCase = moQuestionState.TextCase

End Property

'----------------------------------------------------
Public Property Get ClinicalTest() As String
Attribute ClinicalTest.VB_Description = ""
'----------------------------------------------------
'----------------------------------------------------

    ClinicalTest = moQuestionState.ClinicalTest

End Property

'----------------------------------------------------
Public Property Get DerivationExpr() As String
Attribute DerivationExpr.VB_Description = ""
'----------------------------------------------------
'----------------------------------------------------

    DerivationExpr = moQuestionState.DerivationExpr

End Property


 
 
 
 
 
