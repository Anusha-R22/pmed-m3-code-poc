VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabTestValidateState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: LabTestValidateState.cls
' Copyright InferMed Ltd. 2001 All Rights Reserved
' Author: Toby Aldridge, InferMed, Sept 2001
' Purpose: Lab Test Validation State class for MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 3 Oct 01  Corrected File Header
'           Set Persistable to TRUE
'----------------------------------------------------

Option Explicit

'lab test code
Public ClinicalTest As String
'laboratory code
Public Lab As String
'CTC Scheme code
Public CTCScheme As String

'collection of normal range states corresonding to Lab,ClinicalTest
Public NormalRangeStates As Collection

'collection of CTC states corresonding to CTC Scheme,ClinicalTest
Public CTCStates As Collection

'----------------------------------------------------
Public Property Get Key() As String
'----------------------------------------------------
' Return a key for use when added to collections
'----------------------------------------------------

    Key = ClinicalTest
    
End Property

'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Read class properties from a Property Bag
'----------------------------------------------------

    ClinicalTest = PropBag.ReadProperty("T")
    Lab = PropBag.ReadProperty("L")
    CTCScheme = PropBag.ReadProperty("C")
    Set NormalRangeStates = CollectionDeSerialise(PropBag.ReadProperty("NR"))
    Set CTCStates = CollectionDeSerialise(PropBag.ReadProperty("CTC"))
        
End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Write class properties to a Property Bag
'----------------------------------------------------

    PropBag.WriteProperty "T", ClinicalTest
    PropBag.WriteProperty "L", Lab
    PropBag.WriteProperty "C", CTCScheme
    PropBag.WriteProperty "NR", CollectionSerialise(NormalRangeStates)
    PropBag.WriteProperty "CTC", CollectionSerialise(CTCStates)
    
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

'--------------------------------------------
Private Sub Class_Initialize()
'--------------------------------------------
    Set NormalRangeStates = New Collection
    Set CTCStates = New Collection
End Sub

'--------------------------------------------
Private Sub Class_Terminate()
'--------------------------------------------
    Set NormalRangeStates = Nothing
    Set CTCStates = Nothing
End Sub
