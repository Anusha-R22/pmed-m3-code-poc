VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EFElementState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: EFElementState.cls
' Author: Nicky Johns
' Copyright: InferMed Ltd 2002-2002, All Rights Reserved
' Purpose: eFormElement State class for MACRO 3.0
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 15-21 Jun 01 - Initial development
' NCJ 6 July 01 - Added Categories and Validations
' ZA  07/08/01 - Added Fonts, Captions and X,Y properties
' MACRO 3.0
' NCJ 5 Jul 02 - Added back in QGroupID, OwnerQGroupID, ShowStatusFlag, GroupElementOrder
' ZA 19/07/02  - Added font/colour properties for question caption/comments
' MLM 28/08/02:  Added ElementUse property
' ZA  20/09/02 - Added DisplayLength property
' NCJ 8 Oct 02 - Changed Key (because of eForm Date Element pre-loading)
' NCJ 8 Nov 02 - Added Hotlink property
' ic 14/07/2005  added clinical coding
'----------------------------------------------------

Option Explicit

Public eFormId As Long
Public ElementId As Long

Public QuestionId As Long
Public Code As String
Public Name As String
Public DataType As Integer
Public Format As String
Public QuestionLength As Integer
Public Unit As String
Public DerivationExpr As String
Public Helptext As String
Public TextCase As Integer
Public ClinicalTest As String

Public CollectIfCond As String
Public ControlType As Integer
Public Caption As String
Public ElementOrder As Integer
Public Hidden As Boolean
Public IsOptional As Boolean
Public IsMandatory As Boolean
Public IsLocal As Boolean
Public RequiresRFC As Boolean
Public Authorisation As String
Public ClinicalTestDateExpr As String
Public FontName As String
Public FontSize As Integer
Public FontBold As Boolean
Public FontItalic As Boolean
Public FontColour As Long
Public ElementX As Long
Public ElementY As Long
Public CaptionX As Long
Public CaptionY As Long

Public QGroupID As Long
Public OwnerQGroupID As Long
Public ShowStatusFlag As Boolean
Public GroupElementOrder As Integer

Public Categories As Collection
Public Validations As Collection

Public CaptionFontName As String
Public CaptionFontSize As Integer
Public CaptionFontBold As Boolean
Public CaptionFontItalic As Boolean
Public CaptionFontColour As Long

'MLM 28/08/02:
Public ElementUse As Long

'ZA 20/09/2002
Public DisplayLength As Integer

' NCJ 8 Nov 2002
Public Hotlink As String

'ic 14/07/2005 the code of the dictionary plugin to use
Public DictionaryId As Integer


'----------------------------------------------------
Public Property Get Key() As String
'----------------------------------------------------
' Get the key of this object (for DeSerialisation)
' NCJ 8 Oct 02 - Combine eFormId and ElementId (needed for loading date elements for a study)
'----------------------------------------------------

    Key = Str(eFormId) & "|" & Str(ElementId)

End Property

'----------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Unwrap the class properties from the given property bag
' Ignore Key property
' ic 14/07/2005 added clinical coding
'----------------------------------------------------

    eFormId = PropBag.ReadProperty("eFormId")
    ElementId = PropBag.ReadProperty("ElementId")
    QuestionId = PropBag.ReadProperty("QuestionId")
    Code = PropBag.ReadProperty("Code")
    Name = PropBag.ReadProperty("Name")
    DataType = PropBag.ReadProperty("DataType")
    Format = PropBag.ReadProperty("Format")
    QuestionLength = PropBag.ReadProperty("QuestionLength")
    Unit = PropBag.ReadProperty("Unit")
    DerivationExpr = PropBag.ReadProperty("DerivationExpr")
    Helptext = PropBag.ReadProperty("Helptext")
    TextCase = PropBag.ReadProperty("TextCase")
    ClinicalTest = PropBag.ReadProperty("ClinicalTest")

    CollectIfCond = PropBag.ReadProperty("CollectIfCond")
    ControlType = PropBag.ReadProperty("ControlType")
    Caption = PropBag.ReadProperty("Caption")
    ElementOrder = PropBag.ReadProperty("ElementOrder")
    Hidden = PropBag.ReadProperty("Hidden")
    IsOptional = PropBag.ReadProperty("IsOptional")
    IsMandatory = PropBag.ReadProperty("IsMandatory")
    IsLocal = PropBag.ReadProperty("IsLocal")
    RequiresRFC = PropBag.ReadProperty("RequiresRFC")
    Authorisation = PropBag.ReadProperty("Authorisation")
    ClinicalTestDateExpr = PropBag.ReadProperty("ClinicalTestDateExpr")
    
    FontName = PropBag.ReadProperty("FontName")
    FontSize = PropBag.ReadProperty("FontSize")
    FontBold = PropBag.ReadProperty("FontBold")
    FontItalic = PropBag.ReadProperty("FontItalic")
    FontColour = PropBag.ReadProperty("FontColour")
    ElementX = PropBag.ReadProperty("ElementX")
    ElementY = PropBag.ReadProperty("ElementY")
    CaptionX = PropBag.ReadProperty("CaptionX")
    CaptionY = PropBag.ReadProperty("CaptionY")
    
    Set Categories = CollectionDeSerialise(PropBag.ReadProperty("Categories"), True)
    Set Validations = CollectionDeSerialise(PropBag.ReadProperty("Validations"), True)
  
    ' NCJ 5 Nov 01
    QGroupID = PropBag.ReadProperty("QGroupID")
    OwnerQGroupID = PropBag.ReadProperty("OwnerQGroupID")
    ShowStatusFlag = PropBag.ReadProperty("ShowStatusFlag")
    GroupElementOrder = PropBag.ReadProperty("GroupElementOrder")
    
    ' ZA 19/07/2002
    CaptionFontName = PropBag.ReadProperty("CaptionFontName")
    CaptionFontBold = PropBag.ReadProperty("CaptionFontBold")
    CaptionFontItalic = PropBag.ReadProperty("CaptionFontItalic")
    CaptionFontSize = PropBag.ReadProperty("CaptionFontSize")
    CaptionFontColour = PropBag.ReadProperty("CaptionFontColour")
    
    'MLM 28/08/02:
    ElementUse = PropBag.ReadProperty("ElementUse")
    
    'ZA 20/09/2002
    DisplayLength = PropBag.ReadProperty("DisplayLength")
    Hotlink = PropBag.ReadProperty("Hotlink")   ' NCJ 8 Nov 02
    
    'ic 14/07/2005 read the dictionary code from the propertybag
    DictionaryId = PropBag.ReadProperty("DictionaryId")
    
End Sub

'----------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------
' Wrap up the class properties in the given property bag
' ic 14/07/2005 added clinical coding
'----------------------------------------------------

    PropBag.WriteProperty "eFormId", eFormId
    PropBag.WriteProperty "ElementId", ElementId
    PropBag.WriteProperty "QuestionId", QuestionId
    PropBag.WriteProperty "Code", Code
    PropBag.WriteProperty "Name", Name
    PropBag.WriteProperty "DataType", DataType
    PropBag.WriteProperty "Format", Format
    PropBag.WriteProperty "QuestionLength", QuestionLength
    PropBag.WriteProperty "Unit", Unit
    PropBag.WriteProperty "DerivationExpr", DerivationExpr
    PropBag.WriteProperty "Helptext", Helptext
    PropBag.WriteProperty "TextCase", TextCase
    PropBag.WriteProperty "ClinicalTest", ClinicalTest

    PropBag.WriteProperty "CollectIfCond", CollectIfCond
    PropBag.WriteProperty "ControlType", ControlType
    PropBag.WriteProperty "Caption", Caption
    PropBag.WriteProperty "ElementOrder", ElementOrder
    PropBag.WriteProperty "Hidden", Hidden
    PropBag.WriteProperty "IsOptional", IsOptional
    PropBag.WriteProperty "IsMandatory", IsMandatory
    PropBag.WriteProperty "IsLocal", IsLocal
    PropBag.WriteProperty "RequiresRFC", RequiresRFC
    PropBag.WriteProperty "Authorisation", Authorisation
    PropBag.WriteProperty "ClinicalTestDateExpr", ClinicalTestDateExpr
    
    PropBag.WriteProperty "Fontname", FontName
    PropBag.WriteProperty "FontSize", FontSize
    PropBag.WriteProperty "FontBold", FontBold
    PropBag.WriteProperty "FontItalic", FontItalic
    PropBag.WriteProperty "FontColour", FontColour
    PropBag.WriteProperty "ElementX", ElementX
    PropBag.WriteProperty "ElementY", ElementY
    PropBag.WriteProperty "CaptionX", CaptionX
    PropBag.WriteProperty "CaptionY", CaptionY
    
    PropBag.WriteProperty "Categories", CollectionSerialise(Categories)
    PropBag.WriteProperty "Validations", CollectionSerialise(Validations)

    ' NCJ 5 Nov 01
    PropBag.WriteProperty "QGroupID", QGroupID
    PropBag.WriteProperty "OwnerQGroupID", OwnerQGroupID
    PropBag.WriteProperty "ShowStatusFlag", ShowStatusFlag
    PropBag.WriteProperty "GroupElementOrder", GroupElementOrder
    
    ' ZA 17/07/2002
    PropBag.WriteProperty "CaptionFontName", CaptionFontName
    PropBag.WriteProperty "CaptionFontBold", CaptionFontBold
    PropBag.WriteProperty "CaptionFontItalic", CaptionFontItalic
    PropBag.WriteProperty "CaptionFontSize", CaptionFontSize
    PropBag.WriteProperty "CaptionFontCOlour", CaptionFontColour
    
    'MLM 28/08/02:
    PropBag.WriteProperty "ElementUse", ElementUse
    
    'ZA 20/09/2002
    PropBag.WriteProperty "DisplayLength", DisplayLength
    
    PropBag.WriteProperty "Hotlink", Hotlink
    
    'ic 14/07/2005 write the dictionary code to the propertybag
    PropBag.WriteProperty "DictionaryId", DictionaryId

End Sub

'--------------------------------------------
Public Property Get State() As String
'--------------------------------------------
' Get Question Defn state serialised as a string
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
' Set Question Defn state from serialised string
'--------------------------------------------
Dim pbBag As PropertyBag
Dim bites() As Byte

    bites = sByteArray
    Set pbBag = New PropertyBag
    pbBag.Contents = bites
    Call Class_ReadProperties(pbBag)
    Set pbBag = Nothing
    
End Property



