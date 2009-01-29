VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidationState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: ValidationState.cls
' Copyright InferMed Ltd 2001 All Rights Reserved
' Author: Nicky Johns, InferMed, July 2001
' Purpose: Validation condition state class for questions in MACRO DM 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 6 July 01 - Initial development
' NCJ 31 Jul 01 - Bug fixed Key property
'----------------------------------------------------

Option Explicit

Public QuestionId As Long
Public ValidationID As Long
Public ValidationType As Long
' The Arezzo expression defining the validation condition
Public ValidationCond As String
' The Arezzo expression defining the validation message
Public MessageExpr As String


'----------------------------------------------------
Public Property Get Key() As String
'----------------------------------------------------
' Get the key of this object (for DeSerialisation)
' NCJ 31 Jul 01 - Bug fixed
'----------------------------------------------------

    Key = Str(ValidationID)

End Property

'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Unwrap the class properties from the given property bag
' Ignore Key property
'----------------------------------------------------

    QuestionId = PropBag.ReadProperty("QuestionId")
    ValidationID = PropBag.ReadProperty("ValidationID")
    ValidationType = PropBag.ReadProperty("ValidationType")
    ValidationCond = PropBag.ReadProperty("ValidationCond")
    MessageExpr = PropBag.ReadProperty("MessageExpr")
    
End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Wrap up the class properties in the given property bag
'----------------------------------------------------

    PropBag.WriteProperty "QuestionId", QuestionId
    PropBag.WriteProperty "ValidationID", ValidationID
    PropBag.WriteProperty "ValidationType", ValidationType
    PropBag.WriteProperty "ValidationCond", ValidationCond
    PropBag.WriteProperty "MessageExpr", MessageExpr

End Sub

'--------------------------------------------
Public Property Get State() As String
'--------------------------------------------
' Get Validation state serialised as a string
'--------------------------------------------
Dim pbBag As PropertyBag

    Set pbBag = New PropertyBag
    Call Class_WriteProperties(pbBag)
    State = pbBag.Contents
    Set pbBag = Nothing
    
End Property

'--------------------------------------------
Public Property Let State(sByteArray As String)
'--------------------------------------------
' Set Validation state from serialised string
'--------------------------------------------
Dim pbBag As PropertyBag
Dim bites() As Byte

    bites = sByteArray
    Set pbBag = New PropertyBag
    pbBag.Contents = bites
    Call Class_ReadProperties(pbBag)
    Set pbBag = Nothing
    
End Property


