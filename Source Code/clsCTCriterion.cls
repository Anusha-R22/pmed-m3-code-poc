VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCTCriterion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsCTCriterion.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, September 2000
'   Purpose:    CTCriterion object for MACRO
'----------------------------------------------------------------------------------------'
' Revisions:
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
'----------------------------------------------------------------------------------------'

Option Explicit

Public Id As Long
Public SchemeCode As String

Public Grade As Integer
Public Min As Variant
Public Max As Variant
Public MinType As NRFactor
Public MaxType As NRFactor

Private msClinicalTestCode As String
Private moClinicalTest As clsClinicalTest


'functions to retrieve corresponding tests
'----------------------------------------------------------------------------------------'
Public Property Get ClinicalTest() As clsClinicalTest
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
    
    Set ClinicalTest = moClinicalTest
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let ClinicalTestCode(sClinicalTestCode As String)
'----------------------------------------------------------------------------------------'
' set up linked Clinical Test
'----------------------------------------------------------------------------------------'

    msClinicalTestCode = sClinicalTestCode
    Set moClinicalTest = New clsClinicalTest
    moClinicalTest.Load msClinicalTestCode

End Property

'----------------------------------------------------------------------------------------'
Public Property Get ClinicalTestCode() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    ClinicalTestCode = msClinicalTestCode

End Property

'----------------------------------------------------------------------------------------'
Public Property Get Expr() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Expr = CTCExpr(VarianttoString(Min), VarianttoString(Max), MinType, MaxType, moClinicalTest.Unit)

End Property

'functions to return text values of properties

'----------------------------------------------------------------------------------------'
Public Property Get MinText()
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    MinText = VarianttoString(Min)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get MaxText()
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    MaxText = VarianttoString(Max)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get GradeText() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    GradeText = Format(Grade)

End Property

'----------------------------------------------------------------------------------------'
Public Property Get ListViewSubItems() As Collection
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Set ListViewSubItems = New Collection
    With ListViewSubItems
        .Add moClinicalTest.Code
        .Add GradeText
        .Add Expr
    End With
    
End Property


'----------------------------------------------------------------------------------------'
Public Sub CopyInto(oNewCTCriterion As clsCTCriterion)
'----------------------------------------------------------------------------------------'
'copy all this object's data into oCTCriterion's data
'----------------------------------------------------------------------------------------'

    With oNewCTCriterion
        .Id = Id
        .SchemeCode = SchemeCode
        .ClinicalTestCode = ClinicalTestCode
        .Min = Min
        .Max = Max
        .MinType = MinType
        .MaxType = MaxType
        .Grade = Grade
    End With

End Sub
