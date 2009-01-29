VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyDefState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: StudyDefState.cls
' Nicky Johns, InferMed, May 2001
' Study Definition State class for MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 18-29 May 01 - Initial development
' NCJ 15 Jun 01 - Added ArezzoFile
' ZA 07/08/01 - Added font properties
' ZA 08/08/01 - Added reason for change (RFCs) property
' NCJ 20 Sept 01 - Added RR fields (registration stuff)
' NCJ 27 Sept 01 - Added Status, StudyType and Phase
' NCJ 8 Oct 02 - Added RFOs (Reasons For Overrule) and eFormDateElements
'----------------------------------------------------

Option Explicit

Public StudyId As Long
Public Version As Integer
Public Name As String
Public Description As String
Public GenderExpr As String
Public LabelExpr As String
Public LabelIsLocal As Boolean
Public DOBExpr As String
Public DateFormat As String
Public TimeFormat As String
Public CTCScheme As String
Public Visits As Collection
Public eForms As Collection
Public ArezzoFile As String
Public FontName As String
Public FontSize As Integer
Public FontBold As Boolean
Public FontItalic As Boolean
Public FontColour As Long
Public eFormColour As Long
Public RFCs As Collection
Public Status As Integer
Public StudyType As Integer
Public Phase As Integer

Public RRServerType As Integer
Public RRHTTPAddress As String
Public RRUserName As String
Public RRPassword As String
Public RRProxyServer As String

' NCJ 8 Oct 02
Public RFOs As Collection
Public eFormDateElements As Collection

'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Read class properties from a Property Bag
'----------------------------------------------------

    StudyId = PropBag.ReadProperty("StudyID")
    Version = PropBag.ReadProperty("Version")
    Name = PropBag.ReadProperty("Name")
    Description = PropBag.ReadProperty("Description")
    GenderExpr = PropBag.ReadProperty("GenderExpr")
    LabelExpr = PropBag.ReadProperty("LabelExpr")
    LabelIsLocal = PropBag.ReadProperty("LabelIsLocal")
    DOBExpr = PropBag.ReadProperty("DOBExpr")
    DateFormat = PropBag.ReadProperty("DateFormat")
    TimeFormat = PropBag.ReadProperty("TimeFormat")
    CTCScheme = PropBag.ReadProperty("CTCScheme")
    Set Visits = CollectionDeSerialise(PropBag.ReadProperty("Visits"), True)
    Set eForms = CollectionDeSerialise(PropBag.ReadProperty("eForms"), True)
    ArezzoFile = PropBag.ReadProperty("ArezzoFile")
    
    FontName = PropBag.ReadProperty("FontName")
    FontSize = PropBag.ReadProperty("FontSize")
    FontBold = PropBag.ReadProperty("FontBold")
    FontItalic = PropBag.ReadProperty("FontItalic")
    FontColour = PropBag.ReadProperty("FontColour")
    eFormColour = PropBag.ReadProperty("eFormColour")
    Set RFCs = CollectionDeSerialise(PropBag.ReadProperty("RFCs"), False)
    Set RFOs = CollectionDeSerialise(PropBag.ReadProperty("RFOs"), False)   ' NCJ 8 Oct 02
    Status = PropBag.ReadProperty("Status")
    StudyType = PropBag.ReadProperty("StudyType")
    Phase = PropBag.ReadProperty("Phase")
    
    RRServerType = PropBag.ReadProperty("RRServerType")
    RRHTTPAddress = PropBag.ReadProperty("RRHTTPAddress")
    RRUserName = PropBag.ReadProperty("RRUserName")
    RRPassword = PropBag.ReadProperty("RRPassword")
    RRProxyServer = PropBag.ReadProperty("RRProxyServer")
        
    Set eFormDateElements = CollectionDeSerialise(PropBag.ReadProperty("eFormDateElements"), True)
        
End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Write class properties to a Property Bag
'----------------------------------------------------

    PropBag.WriteProperty "StudyID", StudyId
    PropBag.WriteProperty "Version", Version
    PropBag.WriteProperty "Name", Name
    PropBag.WriteProperty "Description", Description
    PropBag.WriteProperty "GenderExpr", GenderExpr
    PropBag.WriteProperty "LabelExpr", LabelExpr
    PropBag.WriteProperty "LabelIsLocal", LabelIsLocal
    PropBag.WriteProperty "DOBExpr", DOBExpr
    PropBag.WriteProperty "DateFormat", DateFormat
    PropBag.WriteProperty "TimeFormat", TimeFormat
    PropBag.WriteProperty "CTCScheme", CTCScheme
    PropBag.WriteProperty "Visits", CollectionSerialise(Visits)
    PropBag.WriteProperty "eForms", CollectionSerialise(eForms)
    PropBag.WriteProperty "ArezzoFile", ArezzoFile
    
    PropBag.WriteProperty "Fontname", FontName
    PropBag.WriteProperty "FontSize", FontSize
    PropBag.WriteProperty "FontBold", FontBold
    PropBag.WriteProperty "FontItalic", FontItalic
    PropBag.WriteProperty "FontColour", FontColour
    PropBag.WriteProperty "eFormColour", eFormColour
    PropBag.WriteProperty "RFCs", CollectionSerialise(RFCs)
    PropBag.WriteProperty "RFOs", CollectionSerialise(RFOs)     ' NCJ 8 Oct 02
    PropBag.WriteProperty "Status", Status
    PropBag.WriteProperty "StudyType", StudyType
    PropBag.WriteProperty "Phase", Phase

    PropBag.WriteProperty "RRServerType", RRServerType
    PropBag.WriteProperty "RRHTTPAddress", RRHTTPAddress
    PropBag.WriteProperty "RRUserName", RRUserName
    PropBag.WriteProperty "RRPassword", RRPassword
    PropBag.WriteProperty "RRProxyServer", RRProxyServer

    PropBag.WriteProperty "eFormDateElements", CollectionSerialise(eFormDateElements)

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


