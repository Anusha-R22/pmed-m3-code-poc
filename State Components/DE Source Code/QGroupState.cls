VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QGroupState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   File:       QGroupState.cls
'   Author:     Nicky Johns, October 2001
'   Purpose:    Question Group State object to be used
'               for Repeating Question Groups
'               in MACRO DE 3.0 and later versions.
'-----------------------------------------------------------------------------------
' REVISIONS
'   NCJ 22 Nov 01 - Added InitialRows
'   NCJ 23 Nov 01 - Removed ElementId
'-----------------------------------------------------------------------------------

Option Explicit

Public QGroupID As Long
Public Code As String
Public Name As String
Public eFormId As Long
Public DisplayType As Integer
Public DisplayRows As Integer
Public InitialRows As Integer
Public MinRepeats As Integer
Public MaxRepeats As Integer
Public Border As Boolean

'----------------------------------------------------
Public Property Get Key() As String
'----------------------------------------------------
' Get the key of this object (for DeSerialisation)
'----------------------------------------------------

    Key = Str(QGroupID)

End Property


'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Unwrap the class properties from the given property bag
' Ignore Key property
'----------------------------------------------------

    QGroupID = PropBag.ReadProperty("QGroupID")
    Code = PropBag.ReadProperty("Code")
    Name = PropBag.ReadProperty("Name")
    eFormId = PropBag.ReadProperty("eFormId")
    DisplayType = PropBag.ReadProperty("DisplayType")
    DisplayRows = PropBag.ReadProperty("DisplayRows")
    InitialRows = PropBag.ReadProperty("InitialRows")
    MinRepeats = PropBag.ReadProperty("MinRepeats")
    MaxRepeats = PropBag.ReadProperty("MaxRepeats")
    Border = PropBag.ReadProperty("Border")

End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Wrap up the class properties in the given property bag
'----------------------------------------------------

    PropBag.WriteProperty "QGroupID", QGroupID
    PropBag.WriteProperty "Code", Code
    PropBag.WriteProperty "Name", Name
    PropBag.WriteProperty "eFormId", eFormId
    PropBag.WriteProperty "DisplayType", DisplayType
    PropBag.WriteProperty "DisplayRows", DisplayRows
    PropBag.WriteProperty "InitialRows", InitialRows
    PropBag.WriteProperty "MinRepeats", MinRepeats
    PropBag.WriteProperty "MaxRepeats", MaxRepeats
    PropBag.WriteProperty "Border", Border

End Sub

'--------------------------------------------
Public Property Get State() As String
'--------------------------------------------
' Get Group Defn state serialised as a string
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
' Set Group Defn state from serialised string
'--------------------------------------------
Dim pbBag As PropertyBag
Dim bites() As Byte

    bites = sByteArray
    Set pbBag = New PropertyBag
    pbBag.Contents = bites
    Call Class_ReadProperties(pbBag)
    Set pbBag = Nothing
    
End Property



