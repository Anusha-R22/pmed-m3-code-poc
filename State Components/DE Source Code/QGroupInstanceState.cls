VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QGroupInstanceState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: QGroupInstanceState.cls
' Copyright InferMed Ltd 2001 All Rights Reserved
' Author: Nicky Johns, InferMed, July 2001
' Purpose: Question Group Instance state class for MACRO DM 3.0
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 20 Nov 01 - Initial development
'----------------------------------------------------

Option Explicit

Public QGroupID As Long
Public eFormTaskId As Long
Public Rows As Integer
Public Status As Integer
Public LockStatus As Integer
Public IsNew As Boolean
Public IsChanged As Boolean

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
    eFormTaskId = PropBag.ReadProperty("eFormTaskId")
    Rows = PropBag.ReadProperty("Rows")
    Status = PropBag.ReadProperty("Status")
    LockStatus = PropBag.ReadProperty("LockStatus")
    IsNew = PropBag.ReadProperty("IsNew")
    IsChanged = PropBag.ReadProperty("IsChanged")
    
End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Write class properties to a Property Bag
'----------------------------------------------------

    PropBag.WriteProperty "QGroupID", QGroupID
    PropBag.WriteProperty "eFormTaskId", eFormTaskId
    PropBag.WriteProperty "Rows", Rows
    PropBag.WriteProperty "Status", Status
    PropBag.WriteProperty "LockStatus", LockStatus
    PropBag.WriteProperty "IsNew", IsNew
    PropBag.WriteProperty "IsChanged", IsChanged
    
End Sub


