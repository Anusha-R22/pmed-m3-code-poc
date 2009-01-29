VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eFormElementRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: eFormElementRO.cls
' Copyright InferMed Ltd 2001-2005 All Rights Reserved
' Author: Nicky Johns, InferMed, June 2001
' eFormElement Read-only Object for MACRO 2.2
'----------------------------------------------------

' Includes formatting functions

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 15-21 Jun 01 - Initial development
' NCJ 27 Jun 01 - Include formatting functions
' NCJ 3 Jul 01 - Added Key property
' NCJ 9 Jul 01 - Added Categories and Validations
' ZA 07/08/01  - Added Fonts, Caption and ElementX,Y properties
' NCJ 8 Aug 01 - Changed Categories to use new CategoryItem object
' NCJ 11 Sep 01 - Changed Validations to use new Validation object
' NCJ 17 Sep 01 - Updated comments
' TA 02/10/2001 - New error handling
' NCJ 2 Oct 01 - Made DisplayFormat and VBFormat available as properties
' NCJ 3 Oct 01 - Make sure there's a Formatter for every question
' DPH 24/10/2001 - Added Category Active state to SetState
' MACRO 3.0
' NCJ 5 Jul 02 - Added Question Group property things
'               Set up default font details when element is created (SetUpFontDetails)
'               SetGroup now sets the element in the group too
' ZA 19/07/2002 - Added font properties for caption and comment
' MLM 29/08/02: Added ElementUse property and eElementUse enumeration.
' MLM 04/09/02: Added EFormId property.
' ZA 20/09/2002 - Added DisplayLength property
' MLM 26/09/02: Added WebId property to aid in generating the form's HTML.
'   To generate its value, it was necessary to pass an EFormState into the element's SetState.
' NCJ 31 Oct 02 - Changed the WebId property to use eFormId and QuestionId
' NCJ 8 Nov 02 - Added Hotlink property
' RS  21/1/2003 - Added Explicit DateTimeType property
' NCJ 1 Jul 03 - Added Expandable property
' NCJ 24 Mar 04 - Clear out collections and set to Nothing in Terminate (possible memory leak)
' ic 14/07/2005 added clinical coding
' NCJ 8 Dec 05 - New eDateTimeType enumeration
'----------------------------------------------------

Option Explicit

'MLM 29/08/02:
Public Enum eElementUse
    User = 0
    EFormVisitDate = 1
End Enum

' NCJ 21 Jan 03
' NCJ 8 Dec 05 - New enumeration (copied from basEnumerations.bas)
Public Enum eDateTimeType
    dttInvalid = 0
    dttDMY = 1
    dttMDY = 2
    dttYMD = 3
    dttDMYT = 4
    dttMDYT = 5
    dttYMDT = 6
    dttMY = 7
    dttYM = 8
    dttY = 9
    dttT = 10
End Enum

' RS 20/01/2003 Store the exact type of DateFormat
Private menDateTimeType As eDateTimeType

Private moEFElementState As EFElementState

' The object that will format stuff for us
Private moFormatter As DataFormatter

' NCJ 8/8/01 - The collection of category items
Private mcolCategoryItems As Collection

' NCJ 11/9/01 - The collection of Validations
Private mcolValidations As Collection
Private mbActiveCategories As Boolean

' NCJ 5 Nov 01 - The Group/OwnerQGroup defn.
Private moGroup As QGroupRO
Private moOwnerQGroup As QGroupRO

' These are the values for this element
Private mlFontColour As Long
Private msFontName As String
Private mnFontSize As Integer

' ZA 19/07/2002 - caption font values for this element
Private mlCaptionFontColour As Long
Private msCaptionFontName As String
Private mnCaptionFontSize As Integer

'MLM 04/09/02:
Private mlEFormId As Long
'MLM 26/09/02:
Private msWebId As String

'----------------------------------------------------
Property Get DateTimeType() As eDateTimeType
'----------------------------------------------------
' The DateTime type of this element.
' Will have been set when element is loaded
' NCJ 8 Dec 05 - Note new range of values
'----------------------------------------------------
    
    DateTimeType = menDateTimeType

End Property

'----------------------------------------------------
Friend Sub SetDateTimeType(enDateTimeType As eDateTimeType)
'----------------------------------------------------
' Set the date/time type of this question
'----------------------------------------------------
    
    menDateTimeType = enDateTimeType

End Sub

'----------------------------------------------------
Friend Sub SetState(oEFELState As EFElementState, oEFormState As eFormState, _
            sDecPtChar As String, sThouSepChar As String)
'----------------------------------------------------
' Set the State object containing properties
' and initialise ourselves
' NCJ 8/8/01 - Create CategoryItems collection
' NCJ 11/9/01 - Create Validations collection
' DPH 24/10/2001 - Added Cat Active state to SetState
' MLM 04/09/02: Added lEFormId argument.
' MLM 26/09/02: Changed lEFormId argument into oEFormState.
' NCJ 13 May 03 - Added sDecPtChar and sThouSepChar
'----------------------------------------------------
Dim oCatState As CategoryState
Dim oCat As CategoryItem
Dim oValState As ValidationState
Dim oVal As Validation

    On Error GoTo ErrLabel
    
    ' Store the State object
    Set moEFElementState = oEFELState
    
    Set mcolCategoryItems = New Collection
    Set mcolValidations = New Collection
   
    
    With moEFElementState
        ' Create a formatting class
        ' For category questions, create Categories class
        If .QuestionId > 0 Then
            Set moFormatter = New DataFormatter
            ' NCJ 13 May 03 - Include dec.pt. and thou.sep.
            Call moFormatter.SetFormatString(.Format, .DataType, sDecPtChar, sThouSepChar)
            
            If .DataType = eDataType.Text Then
                moFormatter.TextCase = .TextCase
            End If
            
            mbActiveCategories = False
            If .DataType = eDataType.Category Then
                ' Create collection of CategoryItems
                For Each oCatState In moEFElementState.Categories
                    Set oCat = New CategoryItem
                    ' DPH 24/10/2001 - Added Cat Active state
                    If oCatState.Active Then
                        mbActiveCategories = True
                    End If
                    Call oCat.SetItem(oCatState.Value, oCatState.Code, oCatState.Active)
                    mcolCategoryItems.Add oCat, oCat.Code
                Next
            End If
            
            ' Create collection of Validation objects
            For Each oValState In moEFElementState.Validations
                Set oVal = New Validation
                Call oVal.SetItem(oValState.ValidationId, oValState.ValidationType, _
                                    oValState.ValidationCond, _
                                    oValState.MessageExpr)
                mcolValidations.Add oVal, Str(oVal.ValidationId)
            Next
        End If
    End With
    
    'MLM 26/09/02:
    mlEFormId = oEFormState.EFormId
    ' NCJ 31 Oct 02 - Changed WebID to use eFormId and QuestionId
    ' NB This MUST correspond to what's in PfToJavaScript.pc
'    msWebId = LCase("f_" & oEFormState.Code & "_" & oEFElState.Code)
    msWebId = "f_" & mlEFormId & "_" & oEFELState.QuestionId
    
    Set oCat = Nothing
    Set oCatState = Nothing

Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "eFormElementRO.SetState"

End Sub

'-------------------------------------------------------------------
Friend Sub SetUpFontDetails(sFontName As String, nFontSize As Integer, lFontColour As Long)
'-------------------------------------------------------------------
' Set up the font details, using given values if none given here
'-------------------------------------------------------------------

    With moEFElementState
        If .FontName > "" Then
            msFontName = .FontName
        Else
            msFontName = sFontName
        End If
        If .FontSize > 0 Then
            mnFontSize = .FontSize
        Else
            mnFontSize = nFontSize
        End If
        If .FontColour <> 0 Then
            mlFontColour = .FontColour
        Else
            mlFontColour = lFontColour
        End If
    End With

End Sub

'-------------------------------------------------------------------
Friend Sub SetUpCaptionFontDetails(sCaptionFontName As String, nCaptionFontSize As Integer, lCaptionFontColour As Long)
'-------------------------------------------------------------------
'Set up the caption font details, using given values if none give here
'-------------------------------------------------------------------

    With moEFElementState
        If .CaptionFontName > "" Then
            msCaptionFontName = .CaptionFontName
        Else
            msCaptionFontName = sCaptionFontName
        End If
        If .CaptionFontSize > 0 Then
            mnCaptionFontSize = .CaptionFontSize
        Else
            mnCaptionFontSize = nCaptionFontSize
        End If
        If .CaptionFontColour <> 0 Then
            mlCaptionFontColour = .CaptionFontColour
        Else
            mlCaptionFontColour = lCaptionFontColour
        End If
    End With
End Sub
'----------------------------------------------------
Friend Sub SetGroup(oQGroup As QGroupRO)
'----------------------------------------------------
' Set a pointer to this element's QGroup,
' and point the group at ourselves
'----------------------------------------------------

    Set moGroup = oQGroup
    Call oQGroup.SetElement(Me)
    
End Sub

'----------------------------------------------------
Friend Sub SetOwnerQGroup(oQGroup As QGroupRO)
'----------------------------------------------------
' Set a pointer to the QGroup to which this element belongs
'----------------------------------------------------

    Set moOwnerQGroup = oQGroup

End Sub

'----------------------------------------------------
Friend Function LocaliseValue(sValue As String) As String
'----------------------------------------------------
' "Localise" a "standard" value
' sValue is what's been read in from the database (the standard value)
' Return the value as it should appear on the screen,
' i.e. correctly formatted, and taking into account regional settings
'----------------------------------------------------

    On Error GoTo ErrLabel
    
    Select Case moEFElementState.DataType
    Case eDataType.Text
        LocaliseValue = moFormatter.FormatText(sValue)
    Case eDataType.IntegerNumber, eDataType.Real, eDataType.LabTest
        LocaliseValue = moFormatter.StandardToLocalNumber(sValue)
    Case eDataType.DateTime
        ' Get User format if available
        LocaliseValue = sValue
    Case Else
        LocaliseValue = sValue
    End Select

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "eFormElementRO.LocaliseValue"
    
End Function

'-------------------------------------
Friend Property Get Key() As String
'-------------------------------------
' Get unique string key for this object
'-------------------------------------

    Key = moEFElementState.Key

End Property

'----------------------------------------------------
Public Property Get QGroup() As QGroupRO
'----------------------------------------------------
' The Question Group definition for this element.
' Will be Nothing if this element is not of type Group.
'----------------------------------------------------

    Set QGroup = moGroup

End Property

'----------------------------------------------------
Public Property Get OwnerQGroup() As QGroupRO
'----------------------------------------------------
' The Question Group to which this element belongs.
' Will be Nothing if this element does not belong to a Group.
'----------------------------------------------------

    Set OwnerQGroup = moOwnerQGroup

End Property

'----------------------------------------------------
Public Property Get ElementID() As Integer
'----------------------------------------------------
' Element ID of question.
' (See also QuestionID)
'----------------------------------------------------

    ElementID = moEFElementState.ElementID

End Property

'----------------------------------------------------
Public Property Get Name() As String
'----------------------------------------------------
' Name of question
'----------------------------------------------------

    Name = moEFElementState.Name

End Property

'----------------------------------------------------
Public Property Get Code() As String
'----------------------------------------------------
' Unique text code of question
'----------------------------------------------------

    Code = moEFElementState.Code

End Property

'----------------------------------------------------
Public Property Get QuestionId() As Long
'----------------------------------------------------
' QuestionId - will be 0 if this element is not a question.
' (See also ElementID)
'----------------------------------------------------

    QuestionId = moEFElementState.QuestionId

End Property

'----------------------------------------------------
Public Property Get DataType() As Integer
'----------------------------------------------------
' Data type of question.
' Undefined if element not a question.
' See also DataTypeString.
'----------------------------------------------------

    DataType = moEFElementState.DataType

End Property

'----------------------------------------------------
Public Property Get DataTypeString() As String
'----------------------------------------------------
' Data type as a string (for display purposes)
'----------------------------------------------------

    DataTypeString = GetDataTypeString(moEFElementState.DataType)
    
End Property

'----------------------------------------------------
Public Property Get ControlType() As Integer
'----------------------------------------------------
' Control type (i.e. user interface display type).
' See also ControlTypeString.
'----------------------------------------------------

    ControlType = moEFElementState.ControlType

End Property

'----------------------------------------------------
Public Property Get ControlTypeString() As String
'----------------------------------------------------
' Descriptive string for ControlType
'----------------------------------------------------

    ControlTypeString = GetControlTypeString(moEFElementState.ControlType)
    
End Property

'----------------------------------------------------
Public Property Get QuestionLength() As Integer
'----------------------------------------------------
' Maximum length of question response
'----------------------------------------------------

    QuestionLength = moEFElementState.QuestionLength

End Property

'----------------------------------------------------
Public Property Get Unit() As String
'----------------------------------------------------
' Unit of measurement for question.
' May be empty string.
'----------------------------------------------------

    Unit = moEFElementState.Unit

End Property

'----------------------------------------------------
Public Property Get Format() As String
'----------------------------------------------------
' Format string as entered in SD for numeric or text questions.
' Always use VBFormat for actual formatting,
' or DisplayFormat for displaying on screen.
'----------------------------------------------------

    Format = moEFElementState.Format

End Property

'----------------------------------------------------
Public Property Get VBFormat() As String
'----------------------------------------------------
' Format string for use in VB 'Format' command.
'----------------------------------------------------

    VBFormat = moFormatter.VBFormat

End Property

'----------------------------------------------------
Public Property Get DisplayFormat() As String
'----------------------------------------------------
' "Localised" VB Format string, for display only.
' See also VBFormat and Format properties.
'----------------------------------------------------

    DisplayFormat = moFormatter.LocalFormat

End Property

'----------------------------------------------------
Public Property Get Helptext() As String
'----------------------------------------------------
' Help text associated with question.
'----------------------------------------------------

    Helptext = moEFElementState.Helptext

End Property

'----------------------------------------------------
Public Property Get TextCase() As Integer
'----------------------------------------------------
' Whether Upper case, Lower case or Leave
'----------------------------------------------------

    TextCase = moEFElementState.TextCase

End Property

'----------------------------------------------------
Public Property Get ClinicalTest() As String
'----------------------------------------------------
' For LabTest questions only. The Clinical Test associated with the question.
'----------------------------------------------------

    ClinicalTest = moEFElementState.ClinicalTest

End Property

'----------------------------------------------------
Public Property Get DerivationExpr() As String
'----------------------------------------------------
' Arezzo expression defining the question's derivation
'----------------------------------------------------

    DerivationExpr = moEFElementState.DerivationExpr

End Property

'----------------------------------------------------
Public Property Get ClinicalTestDateExpr() As String
'----------------------------------------------------
' For LabTest questions only.
' Arezzo expression defining the question's test date.
'----------------------------------------------------

    ClinicalTestDateExpr = moEFElementState.ClinicalTestDateExpr

End Property

'----------------------------------------------------
Public Property Get ElementOrder() As Integer
'----------------------------------------------------
' Order of this element on the eForm.
'----------------------------------------------------

    ElementOrder = moEFElementState.ElementOrder

End Property

'----------------------------------------------------
Public Property Get CollectIfCond() As String
'----------------------------------------------------
' Arezzo condition defining whether question should be answered.
' If condition is true, collect data.
' If condition is false, do not collect data
'----------------------------------------------------

    CollectIfCond = moEFElementState.CollectIfCond

End Property

'----------------------------------------------------
Public Property Get Caption() As String
'----------------------------------------------------
' Caption to appear on eForm
'----------------------------------------------------

    Caption = moEFElementState.Caption

End Property

'----------------------------------------------------
Public Property Get Authorisation() As String
'----------------------------------------------------
' Role code of user who must authorise question response
'----------------------------------------------------

    Authorisation = moEFElementState.Authorisation

End Property

'----------------------------------------------------
Public Property Get Hidden() As Boolean
'----------------------------------------------------
' If Hidden = TRUE, question is not displayed on eForm
'----------------------------------------------------

    Hidden = moEFElementState.Hidden

End Property

'----------------------------------------------------
Public Property Get IsOptional() As Boolean
'----------------------------------------------------
' Whether question is optional or not
'----------------------------------------------------

    IsOptional = moEFElementState.IsOptional

End Property

'----------------------------------------------------
Public Property Get IsMandatory() As Boolean
'----------------------------------------------------
' Whether question is mandatory or not
' i.e. whether question must be answered before eForm is considered to be complete
'----------------------------------------------------

    IsMandatory = moEFElementState.IsMandatory

End Property

'----------------------------------------------------
Public Property Get IsLocal() As Boolean
'----------------------------------------------------
' Whether question is Local or not.
' If TRUE, question response should NOT ne transmitted to MACRO server.
'----------------------------------------------------

    IsLocal = moEFElementState.IsLocal

End Property

'----------------------------------------------------
Public Property Get RequiresRFC() As Boolean
'----------------------------------------------------
' Whether response needs Reason For Change when changed
'----------------------------------------------------

    RequiresRFC = moEFElementState.RequiresRFC

End Property

'----------------------------------------------------
Public Property Get Categories() As Collection
'----------------------------------------------------
' The collection of category values (if any)
' as CategoryItem objects, indexed by category code.
' Collection items are in correct order for display
'----------------------------------------------------

    Set Categories = mcolCategoryItems
    
End Property

'----------------------------------------------------
Public Property Get ActiveCategories() As Boolean
'----------------------------------------------------
' Returns whether active categories exist (if appropriate)
'----------------------------------------------------

    ActiveCategories = mbActiveCategories
    
End Property

'----------------------------------------------------
Public Property Get CategoryValue(sCatCode As String) As String
'----------------------------------------------------
' Get the category Value corresponding to the given category code.
' Returns empty string if code does not exist or is not active.
'----------------------------------------------------
' REVISIONS
' DPH 24/10/2001 - Only Return active category codes
'----------------------------------------------------
Dim sCat As String
Dim oCat As CategoryState

    sCat = ""
    
    If CollectionMember(moEFElementState.Categories, LCase(sCatCode), True) Then
        Set oCat = moEFElementState.Categories.Item(LCase(sCatCode))
        If oCat.Active Then
            sCat = oCat.Value
        End If
    End If
    
    CategoryValue = sCat
    
End Property

'----------------------------------------------------
Public Property Get CategoryCode(sCatValue As String) As String
'----------------------------------------------------
' Get the category Code corresponding to the given category value.
' Returns empty string if value does not exist or inactive.
' NB There MIGHT be more than one category with this value.
' We return the code of the first (and that's just tough!)
'----------------------------------------------------
' REVISIONS
' DPH 24/10/2001 - Only Return active category codes
'----------------------------------------------------
Dim oCat As CategoryState
Dim sCatCode As String

    On Error GoTo ErrLabel
    
    sCatCode = ""
    For Each oCat In moEFElementState.Categories
        If LCase(oCat.Value) = LCase(sCatValue) And oCat.Active Then
            sCatCode = oCat.Code
            Exit For
        End If
    Next
    CategoryCode = sCatCode

Exit Property

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "eFormElementRO.CategoryCode(Get)"
    
End Property

'----------------------------------------------------
Public Property Get Validations() As Collection
'----------------------------------------------------
' The collection of validation conditions (if any)
' as Validation objects, indexed by ValidationID
'----------------------------------------------------
' NCJ 11/9/01 - Changed from ValidationState objects to Validation objects
'----------------------------------------------------

    Set Validations = mcolValidations
    
End Property

'----------------------------------------------------
Public Property Get Validation(nValidationID As Integer) As Validation
'----------------------------------------------------
' Get particular validation.
' NB ValidationID corresponds to the "warning flag" as returned by Arezzo
'----------------------------------------------------
' NCJ 11/9/01 - Changed from ValidationState to Validation object
'----------------------------------------------------

    Set Validation = mcolValidations.Item(Str(nValidationID))

End Property

'----------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------

    Set moFormatter = Nothing
    Set moEFElementState = Nothing
    Set moGroup = Nothing
    Set moOwnerQGroup = Nothing
    
End Sub

'----------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------
' NCJ 24 Mar 04 - Clear out collections and set to Nothing
'----------------------------------------------------
    
    Call Terminate

End Sub

'----------------------------------------------
Friend Sub Terminate()
'----------------------------------------------
' NCJ 25 Mar 04 - Ensure everything is tidied up
'----------------------------------------------

    On Error GoTo ErrLabel
    
    Call CollectionRemoveAll(mcolCategoryItems)
    Set mcolCategoryItems = Nothing
    
    Call CollectionRemoveAll(mcolValidations)
    Set mcolValidations = Nothing
    
    Set moFormatter = Nothing
    Set moEFElementState = Nothing
    
    If Not moGroup Is Nothing Then
        Call moGroup.Terminate
    End If
    Set moGroup = Nothing
    
    If Not moOwnerQGroup Is Nothing Then
        Call moOwnerQGroup.Terminate
    End If
    Set moOwnerQGroup = Nothing
    
Exit Sub

ErrLabel:
'    Debug.Print "Terminating element"

End Sub

'----------------------------------------------------
Friend Property Get Formatter() As DataFormatter
'----------------------------------------------------
' Give the MACROSchedule project access to the formatter.
'----------------------------------------------------

    Set Formatter = moFormatter

End Property

'----------------------------------------------------
Public Property Get FontName() As String
'----------------------------------------------------
' Name of font to be used for displaying element.
' This returns the default for the study if element
' doesn't have its own.
'----------------------------------------------------

    FontName = msFontName

End Property

'----------------------------------------------------
Public Property Get CaptionFontName() As String
'----------------------------------------------------
' Name of font to be used for displaying caption/comment.
' This returns the default for the study if caption
' doesn't have its own.
'----------------------------------------------------
    
    CaptionFontName = msCaptionFontName
    
End Property

'----------------------------------------------------
Public Property Get FontSize() As Integer
'----------------------------------------------------
' Size of font to be used for displaying element.
' This returns the default for the study if element
' doesn't have its own.
'----------------------------------------------------

    FontSize = mnFontSize

End Property

'----------------------------------------------------
Public Property Get CaptionFontSize() As Integer
'----------------------------------------------------
' Size of font to be used for displaying  caption/comment.
' This returns the default for the study if caption
' doesn't have its own.
'----------------------------------------------------

    CaptionFontSize = mnCaptionFontSize

End Property

'----------------------------------------------------
Public Property Get FontBold() As Boolean
'----------------------------------------------------
' Whether element's font is Bold
'----------------------------------------------------

    FontBold = moEFElementState.FontBold

End Property

'----------------------------------------------------
Public Property Get CaptionFontBold() As Boolean
'----------------------------------------------------
' Whether caption's font is Bold
'----------------------------------------------------

    CaptionFontBold = moEFElementState.CaptionFontBold

End Property

'----------------------------------------------------
Public Property Get DisplayLength() As Integer
'----------------------------------------------------
'DisplayLength property of a question
'----------------------------------------------------

    DisplayLength = moEFElementState.DisplayLength
    
End Property

'----------------------------------------------------
Public Property Get FontItalic() As Boolean
'----------------------------------------------------
' Whether element's font is Italic
'----------------------------------------------------

    FontItalic = moEFElementState.FontItalic

End Property

'----------------------------------------------------
Public Property Get CaptionFontItalic() As Boolean
'----------------------------------------------------
' Whether caption's font is Italic
'----------------------------------------------------

    CaptionFontItalic = moEFElementState.CaptionFontItalic

End Property


'----------------------------------------------------
Public Property Get FontColour() As Long
'----------------------------------------------------
' Colour of font to be used for displaying element.
' This returns the default for the study if element
' doesn't have its own.
'----------------------------------------------------

    FontColour = mlFontColour

End Property

'----------------------------------------------------
Public Property Get CaptionFontColour() As Long
'----------------------------------------------------
' Colour of font to be used for displaying caption/comment.
' This returns the default for the study if caption
' doesn't have its own.
'----------------------------------------------------

    CaptionFontColour = mlCaptionFontColour

End Property


'----------------------------------------------------
Public Property Get ElementX() As Long
'----------------------------------------------------
' X coordinate of element's field
'----------------------------------------------------

    ElementX = moEFElementState.ElementX

End Property

'----------------------------------------------------
Public Property Get ElementY() As Long
'----------------------------------------------------
' Y coordinate of element's field
'----------------------------------------------------

    ElementY = moEFElementState.ElementY

End Property

'----------------------------------------------------
Public Property Get CaptionX() As Long
'----------------------------------------------------
' X coordinate of element's caption
'----------------------------------------------------

    CaptionX = moEFElementState.CaptionX

End Property

'----------------------------------------------------
Public Property Get CaptionY() As Long
'----------------------------------------------------
' Y coordinate of element's caption
'----------------------------------------------------

    CaptionY = moEFElementState.CaptionY

End Property

'----------------------------------------------------
Public Property Get ShowStatusFlag() As Boolean
'----------------------------------------------------
' Whether to show element's status flag
'----------------------------------------------------

    ShowStatusFlag = moEFElementState.ShowStatusFlag

End Property

'----------------------------------------------------
Public Property Get GroupElementOrder() As Integer
'----------------------------------------------------
' The order of the element within its group.
' Will be 0 if element does not belong to a group.
'----------------------------------------------------

    GroupElementOrder = moEFElementState.GroupElementOrder

End Property

'----------------------------------------------------
Public Property Get ElementUse() As eElementUse
'----------------------------------------------------
' MLM 29/08/02: Added. Describes whether the element is the eForm or visit date, or a user element.
'----------------------------------------------------
    
    ElementUse = moEFElementState.ElementUse
    
End Property

'----------------------------------------------------
Public Property Get EFormId() As Long
'----------------------------------------------------
' MLM 04/09/02: Added. Returns the Element's EFormId.
'----------------------------------------------------

    EFormId = mlEFormId
    
End Property

'----------------------------------------------------
Public Property Get WebId() As String
'----------------------------------------------------
' MLM 04/09/02: Added. Returns the Element's WebId, used in modHTML.
' (Only applicable to Questions)
' NB This MUST correspond to what's in PfToJavaScript.pc
' (NCJ 31 Oct - msWebId value changed)
'----------------------------------------------------

    WebId = msWebId
    
End Property

'----------------------------------------------------
Public Property Get Hotlink() As String
'----------------------------------------------------
' The AREZZO Hotlink string
'----------------------------------------------------

    Hotlink = moEFElementState.Hotlink

End Property

'----------------------------------------------------
Public Property Get Expandable() As Boolean
'----------------------------------------------------
' NCJ 1 Jul 03 - Whether this element should have an Expand box
' i.e. if it's a Text Box, and its DisplayLength > 0 and < QuestionLength
' ic 14/07/2005 added clinical coding
'----------------------------------------------------
Const nTEXT_BOX = 1     ' Copied from modMACRODM_UI

    'ic 14/07/2005 check that a question is not a thesaurus type
    Expandable = (Me.ControlType = nTEXT_BOX) _
                And (Me.DisplayLength > 0) _
                And (Me.DisplayLength < Me.QuestionLength) _
                And (Me.DataType <> eDataType.Thesaurus)

End Property

'----------------------------------------------------
Public Property Get DictionaryId() As Integer
'----------------------------------------------------
' ic 14/07/2005 the dictionary code
'----------------------------------------------------
    DictionaryId = moEFElementState.DictionaryId
End Property
