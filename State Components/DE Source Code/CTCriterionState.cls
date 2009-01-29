VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTCriterionState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: CTCriterionState.cls
' Toby Aldridge, InferMed, Sept 2001
' CT Criterion State class for MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------

'----------------------------------------------------

Option Explicit

Public Id As Long
Public SchemeCode As String
Public Grade As Integer
Public Min As Variant
Public Max As Variant
Public MinType As Integer
Public MaxType As Integer
Public ClinicalTestCode As String


'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Read class properties from a Property Bag
'----------------------------------------------------

    Id = PropBag.ReadProperty("1")
    SchemeCode = PropBag.ReadProperty("2")
    Grade = PropBag.ReadProperty("3")
    Min = PropBag.ReadProperty("4")
    Max = PropBag.ReadProperty("5")
    MinType = PropBag.ReadProperty("6")
    MaxType = PropBag.ReadProperty("7")
    ClinicalTestCode = PropBag.ReadProperty("8")
    
End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Write class properties to a Property Bag
'----------------------------------------------------

    PropBag.WriteProperty "1", Id
    PropBag.WriteProperty "2", SchemeCode
    PropBag.WriteProperty "3", Grade
    PropBag.WriteProperty "4", Min
    PropBag.WriteProperty "5", Max
    PropBag.WriteProperty "6", MinType
    PropBag.WriteProperty "7", MaxType
    PropBag.WriteProperty "8", ClinicalTestCode

End Sub

'--------------------------------------------
Public Property Get State() As String
'--------------------------------------------
' Get Study Defn state serialised as a string
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
' Set Study Defn state from serialised string
'--------------------------------------------
Dim pbBag As PropertyBag
Dim bites() As Byte

    bites = sByteArray
    Set pbBag = New PropertyBag
    pbBag.Contents = bites
    Call Class_ReadProperties(pbBag)
    Set pbBag = Nothing
    
End Property




