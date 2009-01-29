VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------
' File: DataFormatter.cls
' Author: Nicky Johns, June 2001
' Copyright: InferMed, 2001-2002, All Rights Reserved
' Takes care of Question Formatting for MACRO 2.2
'----------------------------------------------------

' Requires libLibrary for number format conversion routines

' Intention is that there should be one of these classes
' for each eFormElement that will need data formatting

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 27 Jun 01 - Initial development
' TA 02/10/2001 - New error handling
' NCJ 2 Oct 01 - Made LocalFormat and VBFormat available as properties
' NCJ 15 Jan 02 - Convert to text formats to Upper Case before converting to VB format
'           ( Buglist 2.2.6, bug 13)
' NCJ 19 Mar 02 - Added fix already done to 3.0 by ATO, 26/02/2002
' MLM 22/05/02: Changed ValidNumber.
' NCJ 17 Jun 02 - CBB 2.2.14/20 in ValidNumber
' NCJ 9 Dec 02 - Allow reals in integer fields, and let VB round to integers
' NCJ 13 May 03 - New properties for local dec.pt and thou.sep
'----------------------------------------------------

Option Explicit

' Case conversions for data items
Public Enum eTextCase
    Leave = 0
    Upper = 1
    Lower = 2
End Enum

' The MACRO SD format string
Private msMACRODataFormat As String

' The data type of this question
Private mnDataType As Integer

' The format string that we pass to the VB Format function
Private msVBDataFormat As String

' The "local" format string used for display purposes only
Private msLocalDataFormat As String

' The max number allowed in this format (numeric data only)
Private msMaxNumberAllowed As String

' Whether negative numbers are allowed
Private mbAllowNegative As Boolean

Private mnTextCase As eTextCase

' NCJ 13 May 03 - Local number formatting
Private msDecimalPtChar As String
Private msThousandSepChar As String

'----------------------------------------------------
Public Property Let TextCase(nCase As eTextCase)
'----------------------------------------------------
' Set text case for text items
'----------------------------------------------------

    mnTextCase = nCase

End Property

'----------------------------------------------------
Public Property Get LocalFormat() As String
Attribute LocalFormat.VB_Description = "Return the ""local"" format, i.e. the VB format localised with regional settings (for display purposes only)"
'----------------------------------------------------
' Return the "local" format, i.e. the VB format
' localised with regional settings
' (for display purposes only)
'----------------------------------------------------

    LocalFormat = msLocalDataFormat

End Property

'----------------------------------------------------
Public Property Get VBFormat() As String
Attribute VBFormat.VB_Description = "Return the VB format to be used for formatting values"
'----------------------------------------------------
' Return the VB format to be used for formatting values
'----------------------------------------------------

    VBFormat = msVBDataFormat

End Property

'----------------------------------------------------
Public Sub SetFormatString(ByVal sMacroFormat As String, _
                ByVal nMacroDataType As Integer, _
                ByVal sDecPtChar As String, ByVal sThouSepChar As String)
'----------------------------------------------------
' Store format string and
' translate MACRO format string to VB format string,
' and set up other formatting details
' NCJ 13 May 03 - Added sDecPtChar and sThouSepChar
'----------------------------------------------------
Dim nPosition As Integer
Dim n As Integer
Dim sChar As String
Dim nStart As Integer

    On Error GoTo ErrLabel
    
    ' Store format and data type
    msMACRODataFormat = sMacroFormat
    mnDataType = nMacroDataType
    
    ' NCJ 13 May 03 - Store number format chars (assume valid!)
    msDecimalPtChar = sDecPtChar
    msThousandSepChar = sThouSepChar
    
    msMaxNumberAllowed = ""
    
    Select Case nMacroDataType
    Case eDataType.Real, eDataType.IntegerNumber, eDataType.LabTest
        ' Numeric format strings
        ' Convert 9 to 0 and # to 9
        ' Calculate max number allowed on the way
        msVBDataFormat = ""
        ' Pick off a minus sign
        ' Only allow negative nos. if there's a minus
        If Left(sMacroFormat, 1) = "-" Then
            nStart = 2
            mbAllowNegative = True
        Else
            nStart = 1
            mbAllowNegative = False
        End If
        ' Convert 9's to 0's for VB format
        ' Convert # to 9 for MaxNumberAllowed
        For n = nStart To Len(sMacroFormat)
            sChar = Mid(sMacroFormat, n, 1)
            Select Case sChar
            Case "9"
                msVBDataFormat = msVBDataFormat & "0"
                msMaxNumberAllowed = msMaxNumberAllowed & sChar
            Case "#"
                msVBDataFormat = msVBDataFormat & sChar
                msMaxNumberAllowed = msMaxNumberAllowed & "9"
            Case Else
                msVBDataFormat = msVBDataFormat & sChar
                msMaxNumberAllowed = msMaxNumberAllowed & sChar
            End Select
        Next n
        ' Set the local versions
        msLocalDataFormat = StandardNumToLocal(msVBDataFormat, msDecimalPtChar, msThousandSepChar)
        ' NCJ 5/6/00 SR3521 Add minus sign if negative
        If mbAllowNegative Then
            msLocalDataFormat = "-" & msLocalDataFormat
        End If
        msMaxNumberAllowed = StandardNumToLocal(msMaxNumberAllowed, msDecimalPtChar, msThousandSepChar)
        
    Case eDataType.Text
        ' Note that text formats are only used for CHECKING values
        ' rather than formatting them
        ' NCJ 15 Jan 02 Convert to Upper Case before converting
        msVBDataFormat = UCase(sMacroFormat)
        msVBDataFormat = Replace(msVBDataFormat, "A", "[A-Z]")
        msVBDataFormat = Replace(msVBDataFormat, "0", "#")
        msVBDataFormat = Replace(msVBDataFormat, "9", "#")
        msLocalDataFormat = sMacroFormat
    
    Case eDataType.DateTime
    ' NCJ 9 Feb 00 - Temporary "backwards compatibility" fix
    ' Replace "mmm" with "mm" to avoid text months
        msVBDataFormat = LCase(sMacroFormat)
        msVBDataFormat = Replace(msVBDataFormat, "mmm", "mm")
        msLocalDataFormat = msVBDataFormat
        
    Case Else
        msVBDataFormat = sMacroFormat
        msLocalDataFormat = sMacroFormat
    End Select

Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "DataFormatter.SetFormatString"

End Sub

'------------------------------------------------------------
Public Function FormatLocalNumber(ByVal sNumValue As String) As String
Attribute FormatLocalNumber.VB_Description = "Format the ""local"" numeric value given in sNumValue (i.e. sNumValue is what the user typed in) according to current format string FormatLocalNumber returns what should be displayed on screen"
'------------------------------------------------------------
' Format the "local" numeric value given in sNumValue
' (i.e. sNumValue is what the user typed in)
' according to current format string
' FormatLocalNumber returns what should be displayed on screen
'------------------------------------------------------------
Dim sDisplayText As String

    On Error GoTo ErrLabel
    sDisplayText = sNumValue
    
    If msVBDataFormat > "" And sNumValue > "" Then
        Select Case mnDataType
        
        Case eDataType.Real, eDataType.IntegerNumber, eDataType.LabTest
            sDisplayText = Format(sNumValue, msVBDataFormat)
            ' Screen out empty string and single decimal point etc.
            ' (Can happen if value is 0 and format contains only optional chars)
            If Not IsNumeric(sDisplayText) Then
                ' Assume zero
                sDisplayText = "0"
            'MLM 22/05/02: Remove trailing decimal point, if any.
            ' NCJ 13 May 03 - Use stored dec. pt. char
'            ElseIf Right(sDisplayText, 1) = RegionalDecimalPointChar Then
            ElseIf Right(sDisplayText, 1) = msDecimalPtChar Then
                sDisplayText = Left(sDisplayText, Len(sDisplayText) - 1)
'            Else
'                ' Check the decimal point (in case digit before/after was optional)
'                sDisplayText = AddLeadingTrailingZeros(sDisplayText)
            End If
    
        Case Else
            ' Not a number - leave it as it is
        End Select
    End If
    
    FormatLocalNumber = sDisplayText

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "DataFormatter.FormatLocalNumber"
        
End Function

'-----------------------------------------------------------------------------
Private Function AddLeadingTrailingZeros(sNumber As String) As String
'-----------------------------------------------------------------------------
' If sNumber begins or ends with a decimal point, add a leading or trailing zero
'-----------------------------------------------------------------------------
Dim sNum As String

    On Error GoTo ErrLabel
    
    ' Add on a leading zero if necessary
    ' NCJ 13 May 03 - Use stored dec. pt. char
'    If Left$(sNumber, 1) = RegionalDecimalPointChar Then
    If Left$(sNumber, 1) = msDecimalPtChar Then
        sNum = "0" & sNumber
    Else
        sNum = sNumber
    End If
    
    ' Add on a trailing zero if necessary
    ' NCJ 13 May 03 - Use stored dec. pt. char
'    If Right$(sNum, 1) = RegionalDecimalPointChar Then
    If Right$(sNum, 1) = msDecimalPtChar Then
        sNum = sNum & "0"
    End If

    AddLeadingTrailingZeros = sNum

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "DataFormatter.AddLeadingTrailingZeros"
    
End Function

'-----------------------------------------------------------------------------
Public Function ValidNumber(ByVal sResponseValue As String, _
                        ByRef sDisplayText As String, _
                        ByRef sArezzoValue As String) As Long
'-----------------------------------------------------------------------------
' Validate that the given string is the right number type
' Returns Long value as in eResponseValidation enumeration
' Return formatted (local) value in sDisplayText if result is OK
' sResponseValue assumed to be for a number field
'-----------------------------------------------------------------------------
' NCJ 19 Mar 02 - Use Double format for checking integers
' MLM 22/05/02: Pass in sArezzoValue and set to same as sDisplayText but in standard format
'               with leading zeros so that it will be accepted as a number by Prolog.
' NCJ 17 Jun 02 - CBB 2.2.14/20 Compare length of real AFTER formatting
' NCJ 9 Dec 02 - Same for integers (to allow reals to be rounded to integers)
'-----------------------------------------------------------------------------
Dim sNum As String
Dim dblNum As Double
'Dim lNum As Long

    On Error GoTo ErrLabel
    
    ValidNumber = eResponseValidation.ValueOK
    
    'MLM 22/05/02:
    sNum = sResponseValue   'AddLeadingTrailingZeros(sResponseValue)
        
    '   Check that the value is a valid number
    If Not IsNumeric(sNum) Then
        If mnDataType = eDataType.IntegerNumber Then
            ValidNumber = eResponseValidation.NotAnInteger
        Else
            ' Including LabTest type
            ValidNumber = eResponseValidation.NotANumber
        End If
        Exit Function
    End If
    
    Select Case mnDataType
    Case eDataType.IntegerNumber
    
        ' Convert to double first
        dblNum = CDbl(sNum)
        
        ' NCJ 9 Dec 02 - Don't check for an integer (let VB round reals to integers)
'        ' Check it's really an integer
'        If dblNum <> sNum Then
'            ValidNumber = eResponseValidation.NotAnInteger
'            Exit Function
'        End If
        
        ' Check negativity
        If dblNum < 0 And Not mbAllowNegative Then
            ValidNumber = eResponseValidation.NumberNotPositive
            Exit Function
        End If
        
        ' NCJ 9 Dec 02 - Format the number FIRST to make sure that rounding gets done
        dblNum = CDbl(Me.FormatLocalNumber(sNum))
        
        ' Check not too big
        If dblNum > CDbl(msMaxNumberAllowed) Then
            ValidNumber = eResponseValidation.NumberTooBig
            Exit Function
        End If

        ' Check not too small
        If -(dblNum) > CDbl(msMaxNumberAllowed) Then
            ValidNumber = eResponseValidation.NumberTooSmall
            Exit Function
        End If
    
    Case eDataType.Real, eDataType.LabTest
    
        ' Convert to double
        dblNum = CDbl(sNum)
        
        ' Check for valid conversion
        If dblNum <> sNum Then
            ValidNumber = eResponseValidation.NotANumber
            Exit Function
        End If
        
        ' Check not negative
        If dblNum < 0 And Not mbAllowNegative Then
            ValidNumber = eResponseValidation.NumberNotPositive
            Exit Function
        End If
        
        ' NCJ 17 Jun 2002 - CBB 2.2.14/20
        ' We must format the number FIRST to make sure that rounding etc. gets done
        dblNum = CDbl(Me.FormatLocalNumber(sNum))
        
        ' Check not too big
        If dblNum > CDbl(msMaxNumberAllowed) Then
            ValidNumber = eResponseValidation.NumberTooBig
            Exit Function
        End If
        
        ' Check not too small
        If -(dblNum) > CDbl(msMaxNumberAllowed) Then
            ValidNumber = eResponseValidation.NumberTooSmall
            Exit Function
        End If

    End Select
    
    ' We're OK - so finally we format it correctly
    sDisplayText = Me.FormatLocalNumber(sNum)
    
    'MLM 22/05/02: Make the display value acceptable to Arezzo by converting to standard format
    'and adding a leading 0.
    'MLM 27/05/02: This number is displayed w/o formatting in audit trail, data browser so be a bit more conservative with the leading 0s
    ' NCJ 13 May 03 - Use our stored dec.pt and thou.sep.
    sArezzoValue = LocalNumToStandard(sDisplayText, , msDecimalPtChar, msThousandSepChar)
    If Left(sArezzoValue, 1) = "." Then
        sArezzoValue = "0" & sArezzoValue
    ElseIf Left(sArezzoValue, 2) = "-." Then
        sArezzoValue = "-0" & Mid(sArezzoValue, 2)
    End If

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "DataFormatter.ValidNumber"
    
End Function

'----------------------------------------------------------------------------------
Public Function ValidText(ByVal sResponseValue As String, _
                    ByRef sDisplayText As String) As Long
Attribute ValidText.VB_Description = "Validate text value against its format and return eResponseValidation result Return sDisplayText (in correct case) if result is OK"
'----------------------------------------------------------------------------------
' Validate text value against its format
' and return eResponseValidation result
' Return sDisplayText (in correct case) if result is OK
'----------------------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    ValidText = eResponseValidation.ValueOK
        
    'Check text format (case insensitive)
    If msVBDataFormat > "" Then
        If Not (UCase(sResponseValue) Like msVBDataFormat) Then
            ValidText = eResponseValidation.TextWrongFormat
            Exit Function
        End If
    End If

    ' Convert to correct case
    sDisplayText = FormatText(sResponseValue)

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "DataFormatter.ValidText"

End Function

'------------------------------------------------------------
Public Function StandardToLocalNumber(ByVal sStdNumber As String) As String
Attribute StandardToLocalNumber.VB_Description = "Convert the ""standard"" value given in sStdNumber (i.e. sStdNumber is as read in from the MACRO DB) to a formatted ""local"" value (using current format and regional settings)"
'------------------------------------------------------------
' Convert the "standard" value given in sStdNumber
' (i.e. sStdNumber is as read in from the MACRO DB)
' to a formatted "local" value (using current format and regional settings)
'------------------------------------------------------------
    
    On Error GoTo ErrLabel
    ' NCJ 13 May 03 - Use stored dec.pt. and thou.sep.
    StandardToLocalNumber = FormatLocalNumber(StandardNumToLocal(sStdNumber, msDecimalPtChar, msThousandSepChar))

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "DataFormatter.StandardToLocalNumber"

End Function

'------------------------------------------------------------
Public Function FormatText(sText As String) As String
Attribute FormatText.VB_Description = "Change case of text according to text case"
'------------------------------------------------------------
' Change case of text according to text case
'------------------------------------------------------------
    
    On Error GoTo ErrLabel
    
    Select Case mnTextCase
        Case eTextCase.Leave
            FormatText = sText   'Leave it as it is
        Case eTextCase.Upper  'Upper
            FormatText = UCase(sText)
        Case eTextCase.Lower  'Lower
            FormatText = LCase(sText)
    End Select

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "DataFormatter.FormatText"

End Function
 
 
 
 
 
