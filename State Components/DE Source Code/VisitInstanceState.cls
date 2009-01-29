VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VisitInstanceState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: VisitInstanceState.cls
' Nicky Johns, InferMed, May 2001
' State class for Visit Instance in MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 21-29 May 01 - Initial development
' NCJ 3 Jul 01 - eFormInstances moved to StudySubject
' NCJ 10 Jul 01 - Added IsNew, IsChanged, LockStatus
' NCJ 15 Aug 01 - Changed Key to be VisitTaskId
' TA 14 Aug 02 - New statuses
'----------------------------------------------------

Option Explicit

Public VisitId As Long
Public VisitTaskId As Long
Public CycleNo As Integer
Public Status As Integer
Public LockStatus As Integer
Public VisitDate As Double
Public IsNew As Boolean
Public IsChanged As Boolean

' TA 14 Aug 02 - New statuses
Public DiscrepancyStatus As Integer
Public SDVStatus As Integer
Public NoteStatus As Integer

'----------------------------------------------------
Public Property Get Key() As String
'----------------------------------------------------
' Get the key of this object (for DeSerialisation)
'----------------------------------------------------

    Key = Str(VisitTaskId)

End Property

'----------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------
    
End Sub

'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Unwrap the class properties from the given property bag
' Ignore Key property
' NB Do not change the order!
'----------------------------------------------------

    VisitId = PropBag.ReadProperty("VisitID")
    VisitTaskId = PropBag.ReadProperty("VisitTaskID")
    CycleNo = PropBag.ReadProperty("CycleNo")
    Status = PropBag.ReadProperty("Status")
    LockStatus = PropBag.ReadProperty("LockStatus")
    VisitDate = PropBag.ReadProperty("VisitDate")
    IsNew = PropBag.ReadProperty("IsNew")
    IsChanged = PropBag.ReadProperty("IsChanged")

    DiscrepancyStatus = PropBag.ReadProperty("DiscrepancyStatus")
    SDVStatus = PropBag.ReadProperty("SDVStatus")
    NoteStatus = PropBag.ReadProperty("NoteStatus")
    
End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Write class properties to a Property Bag
' NB Do not change the order!
'----------------------------------------------------

    PropBag.WriteProperty "VisitID", VisitId
    PropBag.WriteProperty "VisitTaskID", VisitTaskId
    PropBag.WriteProperty "CycleNo", CycleNo
    PropBag.WriteProperty "Status", Status
    PropBag.WriteProperty "LockStatus", LockStatus
    PropBag.WriteProperty "VisitDate", VisitDate
    PropBag.WriteProperty "IsNew", IsNew
    PropBag.WriteProperty "IsChanged", IsChanged
    
    PropBag.WriteProperty "DiscrepancyStatus", DiscrepancyStatus
    PropBag.WriteProperty "SDVStatus", SDVStatus
    PropBag.WriteProperty "NoteStatus", NoteStatus

End Sub

'--------------------------------------------
Public Property Get State() As String
'--------------------------------------------
' Get Visit Instance state serialised as a string
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
' Set Visit Instance state from serialised string
'--------------------------------------------
Dim pbBag As PropertyBag
Dim bites() As Byte

    bites = sByteArray
    Set pbBag = New PropertyBag
    pbBag.Contents = bites
    Call Class_ReadProperties(pbBag)
    Set pbBag = Nothing
    
End Property


