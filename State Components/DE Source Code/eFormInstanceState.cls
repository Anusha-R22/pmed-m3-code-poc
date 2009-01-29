VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eFormInstanceState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: eFormInstanceState.cls
' Nicky Johns, InferMed, May 2001
' State class for eForm Instance in MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 21-29 May 01 - Initial development
' NCJ 10 Jul 01 - Added IsNew, IsChanged, LockStatus
' TA 14 Aug 02 - New statuses
'----------------------------------------------------

Option Explicit

Public eFormTaskId As Long
Public eFormId As Long
Public CycleNo As Integer
Public VisitId As Long
Public VisitCycleNo As Integer
Public eFormLabel As String
Public eFormDate As Double
Public Status As Integer
Public LockStatus As Integer
Public LabCode As String
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

    Key = Str(eFormTaskId)

End Property

'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Unwrap the class properties from the given property bag
' Ignore Key property
'----------------------------------------------------

    eFormId = PropBag.ReadProperty("eFormID")
    eFormTaskId = PropBag.ReadProperty("eFormTaskID")
    CycleNo = PropBag.ReadProperty("CycleNo")
    VisitId = PropBag.ReadProperty("VisitId")
    VisitCycleNo = PropBag.ReadProperty("VisitCycleNo")
    eFormLabel = PropBag.ReadProperty("eFormLabel")
    eFormDate = PropBag.ReadProperty("eFormDate")
    Status = PropBag.ReadProperty("Status")
    LockStatus = PropBag.ReadProperty("LockStatus")
    LabCode = PropBag.ReadProperty("LabCode")
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
'----------------------------------------------------

    PropBag.WriteProperty "eFormID", eFormId
    PropBag.WriteProperty "eFormTaskID", eFormTaskId
    PropBag.WriteProperty "CycleNo", CycleNo
    PropBag.WriteProperty "VisitId", VisitId
    PropBag.WriteProperty "VisitCycleNo", VisitCycleNo
    PropBag.WriteProperty "eFormLabel", eFormLabel
    PropBag.WriteProperty "eFormDate", eFormDate
    PropBag.WriteProperty "Status", Status
    PropBag.WriteProperty "LockStatus", LockStatus
    PropBag.WriteProperty "LabCode", LabCode
    PropBag.WriteProperty "IsNew", IsNew
    PropBag.WriteProperty "IsChanged", IsChanged
    
    PropBag.WriteProperty "DiscrepancyStatus", DiscrepancyStatus
    PropBag.WriteProperty "SDVStatus", SDVStatus
    PropBag.WriteProperty "NoteStatus", NoteStatus

End Sub

