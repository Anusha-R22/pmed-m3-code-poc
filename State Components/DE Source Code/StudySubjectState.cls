VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudySubjectState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'----------------------------------------------------
' File: StudySubjectState.cls
' Nicky Johns, InferMed, May 2001
' State class for Study Subject in MACRO 2.2
'----------------------------------------------------

' This class just defines the properties of the Subject
' and contains serialisation code

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 18-29 May 01 - Initial development
' NCJ 3 Jul - Added eFormInstances here
' NCJ 10 Jul 01 - Added IsNew, IsChanged
' TA 14 Aug 02 - New statuses
' NCJ 7 Mar 03 - Added Subject Timestamp fields
'       Removed unused SaveTime field
'----------------------------------------------------

Option Explicit


Public StudyId As Long
Public Site As String
Public PersonID As Long
Public Status As Integer
Public LockStatus As Integer
Public Label As String              ' Subject Label
Public DOB As Double
Public Gender As Integer
Public RegStatus As Integer         ' Registration Status
Public RegID As String              ' Registration ID
Public VisitInstances As Collection     ' Collection of VisitInstanceState objects
Public eFormInstances As Collection     ' Collection of eFormInstanceState objects
Public ArezzoState As String        ' Binary compressed state data
Public IsNew As Boolean
Public IsChanged As Boolean

' TA 14 Aug 02 - New statuses
Public DiscrepancyStatus As Integer
Public SDVStatus As Integer
Public NoteStatus As Integer

' NCJ 7 Mar 03
Public SubjectTimeStamp As Double
Public SubjectTimeStamp_TZ As Integer

'--------------------------------------------
Private Sub Class_Initialize()
'--------------------------------------------
    
    Set VisitInstances = New Collection

End Sub

'--------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'--------------------------------------------
' Read all the class properties from the property bag
'--------------------------------------------

    StudyId = PropBag.ReadProperty("StudyID")
    Site = PropBag.ReadProperty("Site")
    PersonID = PropBag.ReadProperty("PersonID")
    Status = PropBag.ReadProperty("Status")
    LockStatus = PropBag.ReadProperty("LockStatus")
    Label = PropBag.ReadProperty("Label")
    DOB = PropBag.ReadProperty("DOB")
    Gender = PropBag.ReadProperty("Gender")
    RegStatus = PropBag.ReadProperty("RegStatus")
    RegID = PropBag.ReadProperty("RegID")
    Set VisitInstances = CollectionDeSerialise(PropBag.ReadProperty("VisitInstances"), True)
    Set eFormInstances = CollectionDeSerialise(PropBag.ReadProperty("eFormInstances"), True)
    ArezzoState = PropBag.ReadProperty("ArezzoState")
    IsNew = PropBag.ReadProperty("IsNew")
    IsChanged = PropBag.ReadProperty("IsChanged")
    
    DiscrepancyStatus = PropBag.ReadProperty("DiscrepancyStatus")
    SDVStatus = PropBag.ReadProperty("SDVStatus")
    NoteStatus = PropBag.ReadProperty("NoteStatus")
    
    SubjectTimeStamp = PropBag.ReadProperty("SubjectTimeStamp")
    SubjectTimeStamp_TZ = PropBag.ReadProperty("SubjectTimeStamp_TZ")
    
End Sub

'--------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'--------------------------------------------
' Write all the class properties out to the property bag
'--------------------------------------------

    PropBag.WriteProperty "StudyID", StudyId
    PropBag.WriteProperty "Site", Site
    PropBag.WriteProperty "PersonID", PersonID
    PropBag.WriteProperty "Status", Status
    PropBag.WriteProperty "LockStatus", LockStatus
    PropBag.WriteProperty "Label", Label
    PropBag.WriteProperty "DOB", DOB
    PropBag.WriteProperty "Gender", Gender
    PropBag.WriteProperty "RegStatus", RegStatus
    PropBag.WriteProperty "RegID", RegID
    PropBag.WriteProperty "VisitInstances", CollectionSerialise(VisitInstances)
    PropBag.WriteProperty "eFormInstances", CollectionSerialise(eFormInstances)
    PropBag.WriteProperty "ArezzoState", ArezzoState
    PropBag.WriteProperty "IsNew", IsNew
    PropBag.WriteProperty "IsChanged", IsChanged
 
    PropBag.WriteProperty "DiscrepancyStatus", DiscrepancyStatus
    PropBag.WriteProperty "SDVStatus", SDVStatus
    PropBag.WriteProperty "NoteStatus", NoteStatus

    PropBag.WriteProperty "SubjectTimeStamp", SubjectTimeStamp
    PropBag.WriteProperty "SubjectTimeStamp_TZ", SubjectTimeStamp_TZ

End Sub

'--------------------------------------------
Public Property Get State() As String
'--------------------------------------------
' Get subject state serialised as a string
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
' Set subject state from serialised string
'--------------------------------------------
Dim pbBag As PropertyBag
Dim bites() As Byte

    bites = sByteArray
    Set pbBag = New PropertyBag
    pbBag.Contents = bites
    Call Class_ReadProperties(pbBag)
    Set pbBag = Nothing
    
End Property
