VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Validation"
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
' File: Validation.cls
' Author: Nicky Johns
' Copyright: InferMed, June 2001, All Rights Reserved
' eFormElement's Validation condition Read-only Object for MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 11 Sept 01 - Initial development
'----------------------------------------------------

Option Explicit

Private msValCond As String
Private msMsgExpr As String
Private mnValId As Integer
Private mnValType As Integer

'----------------------------------------------------
Friend Sub SetItem(ByVal nValId As Integer, _
                    ByVal nValType As Integer, _
                    ByVal sValCond As String, _
                    ByVal sMsgExpr As String)
'----------------------------------------------------
' Set the validation condition and message
' (only to be done within MACRO Schedule business object)
'----------------------------------------------------

    mnValId = nValId
    mnValType = nValType
    msValCond = sValCond
    msMsgExpr = sMsgExpr
    
End Sub

'----------------------------------------------------
Public Property Get ValidationId() As Integer
Attribute ValidationId.VB_Description = "The validation Id (read-only)"
'----------------------------------------------------
' The validation Id (read-only)
'----------------------------------------------------

    ValidationId = mnValId
    
End Property

'----------------------------------------------------
Public Property Get ValidationType() As Integer
Attribute ValidationType.VB_Description = "The validation Type (read-only)"
'----------------------------------------------------
' The validation Type (read-only)
'----------------------------------------------------

    ValidationType = mnValType
    
End Property

'----------------------------------------------------
Public Property Get ValidationCond() As String
Attribute ValidationCond.VB_Description = "The validation condition (read-only)"
'----------------------------------------------------
' The validation condition (read-only)
'----------------------------------------------------

    ValidationCond = msValCond
    
End Property

'----------------------------------------------------
Public Property Get MessageExpr() As String
Attribute MessageExpr.VB_Description = "The validation message expression (read-only)"
'----------------------------------------------------
' The validation message expression (read-only)
'----------------------------------------------------

    MessageExpr = msMsgExpr
    
End Property


 
 
 
 
 
