VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eFormState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: eFormState.cls
' Nicky Johns, InferMed, May 2001
' State object for eForms in MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 21-29 May 01 - Initial development
' NCJ 15 Jun 01 - PromptForDate
' ZA  07/08/01 - Added background property
' NCJ 20 Sep 01 - Added Sequential and LabelIsLocal
' NCJ 4 Nov 02 - New eFormWidth property
'----------------------------------------------------

Option Explicit

Public eFormId As Long
Public Code As String
Public Name As String
Public eFormOrder As Integer
Public LabelExpr As String
Public DateExpr As String
Public DisplayNumbers As Boolean
Public HideIfInactive As Boolean
Public PromptForDate As Boolean
Public BackgroundColour As Long
Public Sequential As Boolean
Public LabelIsLocal As Boolean
Public eFormWidth As Long

'----------------------------------------------------
Public Property Get Key() As String
'----------------------------------------------------
' Get the key of this object (for DeSerialisation)
'----------------------------------------------------

    Key = Str(eFormId)

End Property

'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Unwrap the class properties from the given property bag
' Ignore Key property
'----------------------------------------------------

    eFormId = PropBag.ReadProperty("eFormId")
    Code = PropBag.ReadProperty("Code")
    Name = PropBag.ReadProperty("Name")
    eFormOrder = PropBag.ReadProperty("eFormOrder")
    LabelExpr = PropBag.ReadProperty("LabelExpr")
    DateExpr = PropBag.ReadProperty("DateExpr")
    DisplayNumbers = PropBag.ReadProperty("DisplayNumbers")
    HideIfInactive = PropBag.ReadProperty("HideIfInactive")
    PromptForDate = PropBag.ReadProperty("PromptForDate")
    BackgroundColour = PropBag.ReadProperty("BackgroundColour")
    Sequential = PropBag.ReadProperty("Sequential")
    LabelIsLocal = PropBag.ReadProperty("LabelIsLocal")
    eFormWidth = PropBag.ReadProperty("eFormWidth")

End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Wrap up the class properties in the given property bag
'----------------------------------------------------

    PropBag.WriteProperty "eFormId", eFormId
    PropBag.WriteProperty "Code", Code
    PropBag.WriteProperty "Name", Name
    PropBag.WriteProperty "eFormOrder", eFormOrder
    PropBag.WriteProperty "LabelExpr", LabelExpr
    PropBag.WriteProperty "DateExpr", DateExpr
    PropBag.WriteProperty "DisplayNumbers", DisplayNumbers
    PropBag.WriteProperty "HideIfInactive", HideIfInactive
    PropBag.WriteProperty "PromptForDate", PromptForDate
    PropBag.WriteProperty "BackgroundColour", BackgroundColour
    PropBag.WriteProperty "Sequential", Sequential
    PropBag.WriteProperty "LabelIsLocal", LabelIsLocal
    PropBag.WriteProperty "eFormWidth", eFormWidth

End Sub

'--------------------------------------------
Public Property Get State() As String
'--------------------------------------------
' Get eForm Defn state serialised as a string
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
' Set eForm Defn state from serialised string
'--------------------------------------------
Dim pbBag As PropertyBag
Dim bites() As Byte

    bites = sByteArray
    Set pbBag = New PropertyBag
    pbBag.Contents = bites
    Call Class_ReadProperties(pbBag)
    Set pbBag = Nothing
    
End Property


