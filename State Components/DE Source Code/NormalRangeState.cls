VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NormalRangeState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: NormalRangeState.cls
' Toby Aldridge, InferMed, Sept 2001
' Normal Range State class for MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------

'----------------------------------------------------

Option Explicit

'all properties
Public Id As Long
Public LabCode As String
Public AgeMin As Variant
Public AgeMax As Variant
Public EffectiveStart As Double
Public EffectiveEnd As Double
Public NormalMin As Variant
Public NormalMax As Variant
Public FeasibleMin As Variant
Public FeasibleMax As Variant
Public AbsoluteMin As Variant
Public Absolutemax As Variant
Public GenderCode As Integer
Public PercentCode As Integer
Public ClinicalTestCode As String


'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Read class properties from a Property Bag
'----------------------------------------------------

    Id = PropBag.ReadProperty("1")
    LabCode = PropBag.ReadProperty("2")
    AgeMin = PropBag.ReadProperty("3")
    AgeMax = PropBag.ReadProperty("4")
    EffectiveStart = PropBag.ReadProperty("5")
    EffectiveEnd = PropBag.ReadProperty("6")
    NormalMin = PropBag.ReadProperty("7")
    NormalMax = PropBag.ReadProperty("8")
    FeasibleMin = PropBag.ReadProperty("9")
    FeasibleMax = PropBag.ReadProperty("10")
    AbsoluteMin = PropBag.ReadProperty("11")
    Absolutemax = PropBag.ReadProperty("12")
    GenderCode = PropBag.ReadProperty("13")
    PercentCode = PropBag.ReadProperty("14")
    ClinicalTestCode = PropBag.ReadProperty("15")

    
End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Write class properties to a Property Bag
'----------------------------------------------------

    PropBag.WriteProperty "1", Id
    PropBag.WriteProperty "2", LabCode
    PropBag.WriteProperty "3", AgeMin
    PropBag.WriteProperty "4", AgeMax
    PropBag.WriteProperty "5", EffectiveStart
    PropBag.WriteProperty "6", EffectiveEnd
    PropBag.WriteProperty "7", NormalMin
    PropBag.WriteProperty "8", NormalMax
    PropBag.WriteProperty "9", FeasibleMin
    PropBag.WriteProperty "10", FeasibleMax
    PropBag.WriteProperty "11", AbsoluteMin
    PropBag.WriteProperty "12", Absolutemax
    PropBag.WriteProperty "13", GenderCode
    PropBag.WriteProperty "14", PercentCode
    PropBag.WriteProperty "15", ClinicalTestCode

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



