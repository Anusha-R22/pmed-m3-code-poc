VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResponseState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: ResponseState.cls
' Author: Nicky Johns
' Copyright: InferMed, June 2001, All Rights Reserved
' Data Entry Response State object for MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 18-21 Jun 01 - Initial development
' NCJ 10 Jul 01 - Changed LabResult to NRStatus
' NCJ 12 Jul 01 - Added IsNew and IsChanged
' TA  25/07/01 - DataItemId added
' NCJ 15 Oct 01 - Changed NRStatus and CTCGrade from String to Integer
' TA 4/3/02 - Added HadValue property
' MACRO 3.0
' NCJ 5 Jul 01 - Added RepeatNumber
' NCJ 12 Aug 02 - Added ChangeCount
' TA 14 Aug 02 - Added extra status properties
' RS 16/09/2002 - Added TimeZone
' NCJ 21 Jan 03 - Added StandardValue (for date/time qus)
' NCJ 28 Jan 03 - Added UserNameFull
' NCJ 29 Apr 03 - Added IgnoreMe property
'----------------------------------------------------

Option Explicit

Public ResponseId As Long
Public RepeatNumber As Integer      ' NCJ 5/7/02
Public ElementId As Long
'TA 25/07/01: added for saving
Public DataItemId As Long

Public TimeStamp As Double
Public TimeSTampTZ As Integer       ' RS 16/9/2002. Added for Timezone handling

Public Status As Integer
Public LockStatus As Integer
Public SoftwareVersion As String

Public Value As String
Public ValueCode As String
Public UserCode As String
Public Comments As String
Public ReasonForChange As String
Public ValidationMessage As String
Public OverruleReason As String

Public NRStatus As Integer
Public CTCGrade As Integer
Public LabTestDate As Double
Public LabCode As String

Public IsNew As Boolean
Public IsChanged As Boolean

'new hadvalue property for GGB
Public HadValue As Boolean

Public ChangeCount As Integer

' TA 14 Aug 02 - New statuses
Public DiscrepancyStatus As Integer
Public SDVStatus As Integer
Public NoteStatus As Integer

' NCJ 21 Jan 03
Public StandardValue As Double
' NCJ 28 Jan 03
Public UserNameFull As String
' NCJ 29 Apr 03
Public IgnoreMe As Boolean

'----------------------------------------------------
Public Property Get Key() As String
'----------------------------------------------------
' Get the key of this object (for DeSerialisation)
' NCJ 2 Nov 01 - Include RepeatNumber
'----------------------------------------------------

    Key = Str(ResponseId) & "|" & RepeatNumber

End Property

'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Unwrap the class properties from the given property bag
' Ignore Key property
'----------------------------------------------------

    ResponseId = PropBag.ReadProperty("ResponseId")
    RepeatNumber = PropBag.ReadProperty("RepeatNumber")
    ElementId = PropBag.ReadProperty("ElementId")
    DataItemId = PropBag.ReadProperty("DataItemId")
    TimeStamp = PropBag.ReadProperty("TimeStamp")
    TimeSTampTZ = PropBag.ReadProperty("TimeStampTZ")               ' RS 16/09/2002
    Status = PropBag.ReadProperty("Status")
    LockStatus = PropBag.ReadProperty("LockStatus")
    SoftwareVersion = PropBag.ReadProperty("SoftwareVersion")
    
    Value = PropBag.ReadProperty("Value")
    ValueCode = PropBag.ReadProperty("ValueCode")
    UserCode = PropBag.ReadProperty("UserCode")
    Comments = PropBag.ReadProperty("Comments")
    ReasonForChange = PropBag.ReadProperty("ReasonForChange")
    ValidationMessage = PropBag.ReadProperty("ValidationMessage")
    OverruleReason = PropBag.ReadProperty("OverruleReason")
    
    NRStatus = PropBag.ReadProperty("NRStatus")
    CTCGrade = PropBag.ReadProperty("CTCGrade")
    LabTestDate = PropBag.ReadProperty("LabTestDate")
    LabCode = PropBag.ReadProperty("LabCode")
    
    IsNew = PropBag.ReadProperty("IsNew")
    IsChanged = PropBag.ReadProperty("IsChanged")
    
    HadValue = PropBag.ReadProperty("HadValue")
    ChangeCount = PropBag.ReadProperty("ChangeCount")
    
    DiscrepancyStatus = PropBag.ReadProperty("DiscrepancyStatus")
    SDVStatus = PropBag.ReadProperty("SDVStatus")
    NoteStatus = PropBag.ReadProperty("NoteStatus")
    
    StandardValue = PropBag.ReadProperty("StandardValue")
    
    UserNameFull = PropBag.ReadProperty("UserNameFull")
    
    IgnoreMe = PropBag.ReadProperty("IgnoreMe")
    
End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Wrap up the class properties in the given property bag
'----------------------------------------------------

    PropBag.WriteProperty "ResponseId", ResponseId
    PropBag.WriteProperty "RepeatNumber", RepeatNumber
    PropBag.WriteProperty "ElementId", ElementId
    PropBag.WriteProperty "DataItemId", DataItemId
    PropBag.WriteProperty "TimeStamp", TimeStamp
    PropBag.WriteProperty "TimeStampTZ", TimeSTampTZ            ' RS 16/09/2002
    PropBag.WriteProperty "Status", Status
    PropBag.WriteProperty "LockStatus", LockStatus
    PropBag.WriteProperty "SoftwareVersion", SoftwareVersion
    
    PropBag.WriteProperty "Value", Value
    PropBag.WriteProperty "ValueCode", ValueCode
    PropBag.WriteProperty "UserCode", UserCode
    PropBag.WriteProperty "Comments", Comments
    PropBag.WriteProperty "ReasonForChange", ReasonForChange
    PropBag.WriteProperty "ValidationMessage", ValidationMessage
    PropBag.WriteProperty "OverruleReason", OverruleReason

    PropBag.WriteProperty "NRStatus", NRStatus
    PropBag.WriteProperty "CTCGrade", CTCGrade
    PropBag.WriteProperty "LabTestDate", LabTestDate
    PropBag.WriteProperty "LabCode", LabCode

    PropBag.WriteProperty "IsNew", IsNew
    PropBag.WriteProperty "IsChanged", IsChanged
    
    PropBag.WriteProperty "HadValue", HadValue
    PropBag.WriteProperty "ChangeCount", ChangeCount

    PropBag.WriteProperty "DiscrepancyStatus", DiscrepancyStatus
    PropBag.WriteProperty "SDVStatus", SDVStatus
    PropBag.WriteProperty "NoteStatus", NoteStatus

    PropBag.WriteProperty "StandardValue", StandardValue

    PropBag.WriteProperty "UserNameFull", UserNameFull

    PropBag.WriteProperty "IgnoreMe", IgnoreMe

End Sub

'--------------------------------------------
Public Property Get State() As String
'--------------------------------------------
' Get Response state serialised as a string
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
' Set Response state from serialised string
'--------------------------------------------
Dim pbBag As PropertyBag
Dim bites() As Byte

    bites = sByteArray
    Set pbBag = New PropertyBag
    pbBag.Contents = bites
    Call Class_ReadProperties(pbBag)
    Set pbBag = Nothing
    
End Property



