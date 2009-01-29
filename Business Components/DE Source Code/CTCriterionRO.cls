VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTCriterionRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: CTCRiterionRO.cls
' Author: Toby Aldridge
' Copyright: InferMed, 2001 - 2004, All Rights Reserved
' Purpose: CTC Object for MACRO 3.0 Data Entry
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'   TA Sept 2002 - Initial development
'   NCJ 24 Mar 04 - Added Terminate to clear moState object
'--------------------------------------

Option Explicit

Public Enum eCTCGrade
    ctcNone = -1
End Enum


Public Enum eNRFactor
    nrfAbsolute = 0
    nrfLower = 1
    nrfUpper = 2
End Enum

Private moState As CTCriterionState

'----------------------------------------------------------------------------------------'
Friend Sub SetState(oState As CTCriterionState)
Attribute SetState.VB_Description = "set up reference to state object"
'----------------------------------------------------------------------------------------'
' set up reference to state object
'----------------------------------------------------------------------------------------'

    Set moState = oState
    
End Sub

'all properties
'----------------------------------------------------
Public Property Get Id() As Long
Attribute Id.VB_Description = "Criterion ID."
'----------------------------------------------------
'Criterion ID.
'----------------------------------------------------

    Id = moState.Id
    
End Property

Public Property Get SchemeCode() As String
Attribute SchemeCode.VB_Description = "Criterion's CTC Scheme code."
'----------------------------------------------------
'Criterion's CTC Scheme code.
'----------------------------------------------------

    SchemeCode = moState.SchemeCode
    
End Property

'----------------------------------------------------
Public Property Get Grade() As Integer
Attribute Grade.VB_Description = "Criterion's CTC grade."
'----------------------------------------------------
'Criterion's CTC grade.
'----------------------------------------------------

    Grade = moState.Grade
    
End Property

Public Property Get Min() As Variant
Attribute Min.VB_Description = "Minimum value for this criterion."
'----------------------------------------------------
'Minimum value for this criterion.
'----------------------------------------------------

    Min = moState.Min
    
End Property

Public Property Get Max() As Variant
Attribute Max.VB_Description = "Maximum value for this criterion."
'----------------------------------------------------
'Maximum value for this criterion.
'----------------------------------------------------

    Max = moState.Max
    
End Property

Public Property Get MinType() As eNRFactor
Attribute MinType.VB_Description = "Minimum based on LLN, ULN or an absolute value?"
'----------------------------------------------------
'Minimum based on LLN, ULN or an absolute value?
'----------------------------------------------------

    MinType = moState.MinType
    
End Property

Public Property Get MaxType() As eNRFactor
Attribute MaxType.VB_Description = "Maximum based on LLN, ULN or an absolute value?"
'----------------------------------------------------
'Maximum based on LLN, ULN or an absolute value?
'----------------------------------------------------

    MaxType = moState.MaxType
    
End Property

Public Property Get ClinicalTestCode() As String
Attribute ClinicalTestCode.VB_Description = "Criterion's clinical test code."
'----------------------------------------------------
' Criterion's clinical test code.
'----------------------------------------------------

    ClinicalTestCode = moState.ClinicalTestCode
    
End Property
 
'----------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------
' NCJ 24 Mar 04 - Clear up moState
'----------------------------------------------------

    Set moState = Nothing

End Sub
