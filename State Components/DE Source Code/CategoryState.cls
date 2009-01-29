VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: CategoryState.cls
' Copyright InferMed Ltd 2001 All Rights Reserved
' Author: Nicky Johns, InferMed, July 2001
' Purpose: Category item state class for questions in MACRO DM 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 6 July 01 - Initial development
' NCJ 31 Jul 01 - Bug fixed Key property
'----------------------------------------------------

Option Explicit

Public QuestionId As Long
Public Code As String
Public Value As String
Public Order As Integer
Public Active As Boolean

'----------------------------------------------------
Public Property Get Key() As String
'----------------------------------------------------
' Get the key of this object (for DeSerialisation)
' NCJ 31 Jul 01 - Bug fixed
'----------------------------------------------------

    Key = LCase(Code)

End Property

'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Unwrap the class properties from the given property bag
' Ignore Key property
'----------------------------------------------------

    QuestionId = PropBag.ReadProperty("QuestionId")
    Code = PropBag.ReadProperty("Code")
    Value = PropBag.ReadProperty("Value")
    Order = PropBag.ReadProperty("Order")
    Active = PropBag.ReadProperty("Active")
    
End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Wrap up the class properties in the given property bag
'----------------------------------------------------

    PropBag.WriteProperty "QuestionId", QuestionId
    PropBag.WriteProperty "Code", Code
    PropBag.WriteProperty "Value", Value
    PropBag.WriteProperty "Order", Order
    PropBag.WriteProperty "Active", Active

End Sub

'--------------------------------------------
Public Property Get State() As String
'--------------------------------------------
' Get Category state serialised as a string
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
' Set Category state from serialised string
'--------------------------------------------
Dim pbBag As PropertyBag
Dim bites() As Byte

    bites = sByteArray
    Set pbBag = New PropertyBag
    pbBag.Contents = bites
    Call Class_ReadProperties(pbBag)
    Set pbBag = Nothing
    
End Property


