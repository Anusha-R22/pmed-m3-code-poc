VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIMsgState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: MIMsgState.cls
' Toby Aldridge, InferMed, Nov 2001
' State class for MIMessage in MACRO
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' TA Nov 01 - Initial development
' RS 30/09/2002 Added Timezone members
'----------------------------------------------------

Option Explicit

Public Id As Long
Public Source As Integer
Public Site As String
Public MIMessageType As Integer
Public Scope As Integer
Public ObjectId As Long
Public ObjectSource As Long
Public Priority As Integer
Public StudyName As String
Public SubjectId As Long
Public VisitId As Long
Public VisitCycle As Integer
Public EFormTaskId As Long
Public ResponseTaskId As Long
Public ResponseValue As String
Public ExternalId As Long
Public TimeCreated As Double
Public TimeSent As Double
Public TimeReceived As Double
Public History As Integer
Public Processed As Integer
Public Status As Integer
Public Text As String
Public UserName As String
Public UserNameFull As String
Public ResponseTimeStamp As Double
Public ResponseCycle As Integer


' RS 30/09/2002: Added TimezoneOffset members
Public TimeCreatedTimezoneOffset As Integer
Public TimeSentTimezoneOffset As Integer
Public TimeReceivedTimezoneOffset As Integer

'TA 18/11/2002: New fields
Public EFormId As Long
Public EFormCycle As Integer
Public QuestionId As Long
Public DataUsername As String


Public IsNew As Boolean
Public IsChanged As Boolean

'----------------------------------------------------
Public Property Get Key() As String
'----------------------------------------------------
' Get the key of this object (for DeSerialisation)
' is id|source|site
'----------------------------------------------------

    Key = Str(Id) & "|" & Str(Source) & "|" & Site

End Property

'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Unwrap the class properties from the given property bag
' Ignore Key property
'----------------------------------------------------


    Id = PropBag.ReadProperty("i")
    Source = PropBag.ReadProperty("s")
    Site = PropBag.ReadProperty("si")
    MIMessageType = PropBag.ReadProperty("m")
    Scope = PropBag.ReadProperty("sc")
    ObjectId = PropBag.ReadProperty("o")
    ObjectSource = PropBag.ReadProperty("os")
    Priority = PropBag.ReadProperty("p")
    StudyName = PropBag.ReadProperty("sn")
    SubjectId = PropBag.ReadProperty("sub")
    VisitId = PropBag.ReadProperty("v")
    VisitCycle = PropBag.ReadProperty("vc")
    EFormTaskId = PropBag.ReadProperty("efi")
    ResponseTaskId = PropBag.ReadProperty("resp")
    ResponseValue = PropBag.ReadProperty("respval")
    ExternalId = PropBag.ReadProperty("ext")
    TimeCreated = PropBag.ReadProperty("cr")
    TimeSent = PropBag.ReadProperty("se")
    TimeReceived = PropBag.ReadProperty("re")
    History = PropBag.ReadProperty("h")
    Processed = PropBag.ReadProperty("process")
    Status = PropBag.ReadProperty("stat")
    Text = PropBag.ReadProperty("text")
    UserName = PropBag.ReadProperty("u")
    UserNameFull = PropBag.ReadProperty("uf")
    ResponseTimeStamp = PropBag.ReadProperty("rts")
    ResponseCycle = PropBag.ReadProperty("rc")
    
    TimeCreatedTimezoneOffset = PropBag.ReadProperty("crtz")
    TimeSentTimezoneOffset = PropBag.ReadProperty("setz")
    TimeReceivedTimezoneOffset = PropBag.ReadProperty("retz")
    
    EFormId = PropBag.ReadProperty("efiid")
    EFormCycle = PropBag.ReadProperty("eficycle")
    QuestionId = PropBag.ReadProperty("q")
    DataUsername = PropBag.ReadProperty("du")
    
    
    IsNew = PropBag.ReadProperty("IsNew")
    IsChanged = PropBag.ReadProperty("IsChanged")
    
End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Write class properties to a Property Bag
'----------------------------------------------------

    PropBag.WriteProperty "i", Id
    PropBag.WriteProperty "s", Source
    PropBag.WriteProperty "si", Site
    PropBag.WriteProperty "m", MIMessageType
    PropBag.WriteProperty "sc", Scope
    PropBag.WriteProperty "o", ObjectId
    PropBag.WriteProperty "os", ObjectSource
    PropBag.WriteProperty "p", Priority
    PropBag.WriteProperty "sn", StudyName
    PropBag.WriteProperty "sub", SubjectId
    PropBag.WriteProperty "v", VisitId
    PropBag.WriteProperty "vc", VisitCycle
    PropBag.WriteProperty "efi", EFormTaskId
    PropBag.WriteProperty "resp", ResponseTaskId
    PropBag.WriteProperty "respval", ResponseValue
    PropBag.WriteProperty "ext", ExternalId
    PropBag.WriteProperty "cr", TimeCreated
    PropBag.WriteProperty "se", TimeSent
    PropBag.WriteProperty "re", TimeReceived
    PropBag.WriteProperty "h", History
    PropBag.WriteProperty "process", Processed
    PropBag.WriteProperty "stat", Status
    PropBag.WriteProperty "text", Text
    PropBag.WriteProperty "u", UserName
    PropBag.WriteProperty "uf", UserNameFull
    PropBag.WriteProperty "rts", ResponseTimeStamp
    PropBag.WriteProperty "rc", ResponseCycle

    PropBag.WriteProperty "crtz", TimeCreatedTimezoneOffset
    PropBag.WriteProperty "setz", TimeSentTimezoneOffset
    PropBag.WriteProperty "retz", TimeReceivedTimezoneOffset

    PropBag.WriteProperty "efiid", EFormId
    PropBag.WriteProperty "eficycle", EFormCycle
    PropBag.WriteProperty "q", QuestionId
    PropBag.WriteProperty "du", DataUsername

    PropBag.WriteProperty "IsNew", IsNew
    PropBag.WriteProperty "IsChanged", IsChanged
    
End Sub



'--------------------------------------------
Public Property Get State() As String
'--------------------------------------------
' Getstate serialised as a string
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
' Set  state from serialised string
'--------------------------------------------
Dim pbBag As PropertyBag
Dim bites() As Byte

    bites = sByteArray
    Set pbBag = New PropertyBag
    pbBag.Contents = bites
    Call Class_ReadProperties(pbBag)
    Set pbBag = Nothing
    
End Property
