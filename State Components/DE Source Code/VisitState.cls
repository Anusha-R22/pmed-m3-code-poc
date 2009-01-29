VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VisitState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: VisitState.cls
' Nicky Johns, InferMed, May 2001
' State object for Visit object in MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 21-29 May 01 - Initial development
'----------------------------------------------------

Option Explicit

Public VisitId As Long
Public Code As String
Public Name As String
Public VisitOrder As Integer
Public DateLabelExpr As String
Public PromptForDate As Boolean
Public BackgroundColour As Long
Public MaxRepeats As Integer
Public VisitEForms As Collection

'----------------------------------------------------
Public Property Get Key() As String
'----------------------------------------------------
' Get the key of this object (for DeSerialisation)
'----------------------------------------------------

    Key = Str(VisitId)

End Property

'----------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------
'----------------------------------------------------

    Set VisitEForms = New Collection
    
End Sub

'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Unwrap the class properties from the given property bag
' Ignore Key property
'----------------------------------------------------

    VisitId = PropBag.ReadProperty("VisitId")
    Code = PropBag.ReadProperty("Code")
    Name = PropBag.ReadProperty("Name")
    VisitOrder = PropBag.ReadProperty("VisitOrder")
    DateLabelExpr = PropBag.ReadProperty("DateLabelExpr")
    PromptForDate = PropBag.ReadProperty("PromptForDate")
    BackgroundColour = PropBag.ReadProperty("BackgroundColour")
    MaxRepeats = PropBag.ReadProperty("MaxRepeats")
    Set VisitEForms = CollectionDeSerialise(PropBag.ReadProperty("VisitEForms"), True)
    
End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Wrap up the class properties in the given property bag
'----------------------------------------------------

    PropBag.WriteProperty "VisitId", VisitId
    PropBag.WriteProperty "Code", Code
    PropBag.WriteProperty "Name", Name
    PropBag.WriteProperty "VisitOrder", VisitOrder
    PropBag.WriteProperty "DateLabelExpr", DateLabelExpr
    PropBag.WriteProperty "PromptForDate", PromptForDate
    PropBag.WriteProperty "BackgroundColour", BackgroundColour
    PropBag.WriteProperty "MaxRepeats", MaxRepeats
    PropBag.WriteProperty "VisitEForms", CollectionSerialise(VisitEForms)
    
End Sub

'--------------------------------------------
Public Property Get State() As String
'--------------------------------------------
' Get Visit Defn state serialised as a string
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
' Set Visit Defn state from serialised string
'--------------------------------------------
Dim pbBag As PropertyBag
Dim bites() As Byte

    bites = sByteArray
    Set pbBag = New PropertyBag
    pbBag.Contents = bites
    Call Class_ReadProperties(pbBag)
    Set pbBag = Nothing
    
End Property


