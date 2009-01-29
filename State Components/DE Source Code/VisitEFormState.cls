VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VisitEFormState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: VisitEFormState.cls
' Nicky Johns, InferMed, May 2001
' State object for VisitEForm object in MACRO 2.2
' This represents a cell in the Schedule grid in SD
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 23-29 May 01 - Initial development
' MLM 28/08/02: Include new EFormUse field in VisitEFormState
'----------------------------------------------------

Option Explicit

Public VisitId As Long
Public eFormId As Long
Public Repeating As Boolean
Public MaxRepeats As Integer
'MLM 28/08/02:
Public EFormUse As Long

'----------------------------------------------------
Public Property Get Key()
'----------------------------------------------------
' Get Key for this object in a collection
'----------------------------------------------------

    Key = VisitEFormKey(VisitId, eFormId)

End Property

'----------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------


End Sub

'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Read class properties from a Property Bag
'----------------------------------------------------

    VisitId = PropBag.ReadProperty("VisitId")
    eFormId = PropBag.ReadProperty("eFormId")
    Repeating = PropBag.ReadProperty("Repeating")
    MaxRepeats = PropBag.ReadProperty("MaxRepeats")
    'MLM 28/08/02:
    EFormUse = PropBag.ReadProperty("EFormUse")

End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Write class properties to a Property Bag
'----------------------------------------------------

    PropBag.WriteProperty "VisitId", VisitId
    PropBag.WriteProperty "eFormId", eFormId
    PropBag.WriteProperty "Repeating", Repeating
    PropBag.WriteProperty "MaxRepeats", MaxRepeats
    'MLM 28/08/02:
    PropBag.WriteProperty "EFormUse", EFormUse

End Sub

'--------------------------------------------
Public Property Get State() As String
'--------------------------------------------
' Get VisitEForm state serialised as a string
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
' Set VisitEForm state from serialised string
'--------------------------------------------
Dim pbBag As PropertyBag
Dim bites() As Byte

    bites = sByteArray
    Set pbBag = New PropertyBag
    pbBag.Contents = bites
    Call Class_ReadProperties(pbBag)
    Set pbBag = Nothing
    
End Property


