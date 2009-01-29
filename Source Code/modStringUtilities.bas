Attribute VB_Name = "modStringUtilities"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999-2000. All Rights Reserved
'   File:       modStringUtilities.bas
'   Author      Paul Norris, 16/09/99
'   Purpose:    Alll common string functions held within this module.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
' NCJ 28 Sept 99    StartsWithAlpha function
' WillC  11/10/99   Added error handlers
' NCJ   15 Nov 99   Added "regional settings" routines
' NCJ   17 Nov 99    Moved ConvertDateToProformaSyntax here from basProformaEngine
' NCJ   25 Nov 99   Added ConvertDateFromArezzo
' NCJ 9 Feb 00 - Moved ConvertDateFromArezzo from here to frmArezzo
' Mo    24/5/00     Added function URLCharToHexEncoding
' TA    16/10/2000: StringtoNumberVariant and VariantToString fuctions moved here from modLabDefinitions
' Mo    14/11/2000  Function ReplaceCharacters removed. It has been replaced by the VB Replace Function.
' NCJ 3 Jan 01 - ConvertStringToCollection commented out since no longer used
' Mo    15/4/2002   gblnValidString changed new validation criteria valDecimalPoint added
' NCJ 17 Jun 02 - New optional parameter in ConvertLocalNumToStandard (CBB 2.2.14/19)
' MLM 24/06/02 Added HexEncodeChars and HexDecodeChars
' NCJ 2 Oct 02 - Discontinue use of ConvertLocalNumToStandard and ConvertStandardToLocalNum
'       and instead use the libLibrary versions LocalNumToStandard and StandardNumToLocal
' Mo    4/2/2003    New function SplitCSV added (for use of MACRO Batch Data Entry)
'----------------------------------------------------------------------------------------'

Option Explicit
Option Compare Text

' NCJ 31/5/00 - string constant for displaying when invalid chars are used
Public Const gsCANNOT_CONTAIN_INVALID_CHARS = " may not contain double or backward quotes, tildes or the | character."

' Alphabet and numbers - NCJ 28/9/99
Private Const msALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Private Const msNUMBERS = "0123456789"
Private Const msSTANDARD_DOT = "."
Private Const msSTANDARD_COMMA = ","

'------------------------------------------------------------------------------------'
Public Function AfterStr(sInput As String, ByVal lPos As Long) As String
'------------------------------------------------------------------------------------'
' PN 05/08/99 - routine added
' function that Returns the portion of the string after lPos
'------------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler
    
    If lPos < Len(sInput) Then
        AfterStr = Right(sInput, Len(sInput) - lPos)
    End If
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "AfterStr", "modStringUtilities")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function
'------------------------------------------------------------------------------------'
Public Function BeforeStr(sInput As String, ByVal lPos As Long) As String
'------------------------------------------------------------------------------------'
' PN 05/08/99 - routine added
' function that Returns the portion of the string before lPos
'------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler

    If lPos <= Len(sInput) Then
        BeforeStr = Left$(sInput, Len(sInput) - lPos)
     End If
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "BeforeStr", "modStringUtilities")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function


'------------------------------------------------------------------------------------'
Public Function CountStrsInStr(sMain As String, sSub As String, _
        Optional iCompare As VbCompareMethod) As Long
'------------------------------------------------------------------------------------'
' PN 30/08/99 - routine added
' Counts the occurences of a substring in a main string.
'------------------------------------------------------------------------------------'
Dim lCount As Long
Dim lPos As Long

    On Error GoTo ErrHandler
    
    lPos = InStr(1, sMain, sSub, iCompare)
    
    Do While lPos > 0
        lCount = lCount + 1
        lPos = InStr(lPos + 1, sMain, sSub, iCompare)
        
    Loop
    CountStrsInStr = lCount
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CountStrsInStr", "modStringUtilities")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'---------------------------------------------------------------------
Public Function gblnValidString(ByVal sString As String, _
                                ByVal nScheme As Integer) As Boolean
'---------------------------------------------------------------------
' This fx validates strings
' pass in nScheme to indicate the type of validation and the function
' will return true for successful validation and false for failed validation'
'
' NOTE that when using the valOnlySingleQuotes as part of the nScheme
' parameter other validation will not be carried out
' NCJ 17/5/00 SR3438 Block tilde character as standard
' Mo Morris 15/4/2002 valDecimalPoint added as new criteria for validation
' Mo Morris 25/7/2007 space and decimal point and hyphen added to valDateSeperators
'---------------------------------------------------------------------
Dim sValidationChars As String
Dim nIndex As Integer

    On Error GoTo ErrHandler
    
    '   ATN 7/7/99
    '   Added check for pipes as well as quotes
    If nScheme And valOnlySingleQuotes Then
        'changed Mo Morris 14/4/2000
        'note that Chr(34)= "
        'note that Chr(124)= |
        'note that Chr(126)= ~
        'note that the ! character is not been checked for,
        'it is here to reverse the manner in which the Like function operates
        ' NCJ 17/5/00 SR3438 Added Chr(126) for tilde
        sValidationChars = "[!`" & Chr(34) & Chr(124) & Chr(126) & "]"
        For nIndex = 1 To Len(sString)
            If Not Mid(sString, nIndex, 1) Like sValidationChars Then
                gblnValidString = False
                Exit Function
            End If
        Next nIndex
    Else
    
        sValidationChars = "["
        
        If nScheme And valAlpha Then
            sValidationChars = sValidationChars & msALPHABET
        End If
        If nScheme And valNumeric Then
            sValidationChars = sValidationChars & msNUMBERS
        End If
        If nScheme And valSpace Then
            sValidationChars = sValidationChars & " "
        End If
        If nScheme And valComma Then
            sValidationChars = sValidationChars & ","
        End If
        If nScheme And valUnderscore Then
            sValidationChars = sValidationChars & "_"
        End If
        If nScheme And valDateSeperators Then
            sValidationChars = sValidationChars & ":/ .-"
        End If
        ' PN 02/09/99
        ' include mathematical operators in the check string
        If nScheme And valMathsOperators Then
            sValidationChars = sValidationChars & "+-/*"
        End If
        'Mo Morris 15/4/2002
        If nScheme And valDecimalPoint Then
            sValidationChars = sValidationChars & "."
        End If
        
        sValidationChars = sValidationChars & "]"
            
        For nIndex = 1 To Len(sString)
            If Not Mid(sString, nIndex, 1) Like sValidationChars Then
                gblnValidString = False
                Exit Function
            End If
        Next nIndex
    End If
    
    gblnValidString = True
 
    Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gblnValidString", "modStringUtilities")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'------------------------------------------------------------------------------------'
Public Function StartsWithAlpha(sString As String) As Boolean
'------------------------------------------------------------------------------------'
' Check if string starts with an alphabetic character
' Returns True if OK, False if not
' NCJ 28 Sept 99
'------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler

    StartsWithAlpha = (Left(sString, 1) Like "[" & msALPHABET & "]")
 
    Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "StartsWithAlpha", "modStringUtilities")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'---------------------------------------------------------------------
Public Function StripFileNameFromPath(ByRef sPathAndFileName As String) As String
'---------------------------------------------------------------------
' This function strips a filename from its full path & filename string
' and returns the filename
'---------------------------------------------------------------------
Dim lPos As Long

    On Error GoTo ErrHandler
    
    'Check that PathAndFileName contains "\" characters, if not just return it
    If InStr(sPathAndFileName, "\") = 0 Then
        StripFileNameFromPath = sPathAndFileName
        Exit Function
    End If
    
    'initialise string position
    lPos = 0
    
    Do
        lPos = InStr(lPos + 1, sPathAndFileName, "\")
    Loop Until InStr(lPos + 1, sPathAndFileName, "\") = 0
    
    StripFileNameFromPath = Mid(sPathAndFileName, lPos + 1)
 
    Exit Function
    
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "StripFileNameFromPath", "modStringUtilities")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'------------------------------------------------------------------------------------'
Public Function TrimNull(ByVal sInString As String) As String
'------------------------------------------------------------------------------------'
' PN 05/08/99 - routine added
' Function to strip trailing nulls from a string
'------------------------------------------------------------------------------------'
Dim sResult As String

    On Error GoTo ErrHandler
    
    sResult = sInString
    Do While Left$(sResult, 1) = vbCr Or Left$(sResult, 1) = vbLf
        sResult = AfterStr(sResult, 1)
    Loop
    Do While Right$(sResult, 1) = vbCr Or Right$(sResult, 1) = vbLf
        sResult = BeforeStr(sResult, 1)
    Loop
    TrimNull = sResult
 
    Exit Function
    
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "TrimNull", "modStringUtilities")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'-------------------------------------------------------------------------
Public Function ConvertLocalNumToStandard(sLocalNumStr As String, _
                        Optional bIncludeThousands As Boolean = False) As String
'-------------------------------------------------------------------------
' NCJ 2 Oct 02 - Discontinue use of this function
' Always use (identical) LocalNumToStandard (in libLibrary)
'-------------------------------------------------------------------------

    ConvertLocalNumToStandard = LocalNumToStandard(sLocalNumStr, bIncludeThousands)

End Function

'-------------------------------------------------------------------------
Public Function ConvertStandardToLocalNum(sStdNumStr As String) As String
'-------------------------------------------------------------------------
' NCJ 2 Oct 02 - Discontinue use of this function
' Always use (identical) StandardNumToLocal (in libLibrary)
'-------------------------------------------------------------------------

    ConvertStandardToLocalNum = StandardNumToLocal(sStdNumStr)
    
End Function

'---------------------------------------------------------------------
Public Function SQLStandardNow() As String
'---------------------------------------------------------------------
' NCJ 4 Feb 00 SR2851
' Returns Now as a double in STANDARD numeric format
' suitable for adding to SQL strings
' NB This deals with problems caused by using CDbl(Now) in SQL strings
' with non-English regional settings
' NCJ 2 Oct 02 - Use new IMedNow function
'---------------------------------------------------------------------

    SQLStandardNow = LocalNumToStandard(IMedNow)

End Function

'----------------------------------------------------------------------
Public Function StandardiseLocalValue(sLocValue As String, nDataType As Integer) As String
'----------------------------------------------------------------------
' Standardise a local value according to its data type
' If not a number, do nothing
' Otherwise convert regional separators into dots and commas
' NCJ 17 Nov 99
'----------------------------------------------------------------------
Dim dblNum As Double
Dim lNum As Long

    ' Default to keeping it the same
    StandardiseLocalValue = sLocValue
    If sLocValue > "" Then
        ' Convert to numeric variable first to get rid of superfluous formatting
        Select Case nDataType
        Case DataType.IntegerData
            lNum = CLng(sLocValue)
            StandardiseLocalValue = LocalNumToStandard(CStr(lNum))
        Case DataType.Real
            dblNum = CDbl(sLocValue)
            StandardiseLocalValue = LocalNumToStandard(CStr(dblNum))
        End Select
    End If
    
End Function

'----------------------------------------------------------------------
Public Function LocaliseStandardValue(sStdValue As String, nDataType As Integer) As String
'----------------------------------------------------------------------
' Localise a standard value according to its data type
' If not a number, do nothing
' Otherwise convert dots and commas into regional separators
' NCJ 17 Nov 99
'----------------------------------------------------------------------

    Select Case nDataType
    Case DataType.IntegerData, DataType.Real
        LocaliseStandardValue = StandardNumToLocal(sStdValue)
    Case Else
        LocaliseStandardValue = sStdValue
    End Select

End Function

'----------------------------------------------------------------------
Public Function ConvertDateToProformaSyntax(vDate As Variant) As String
'----------------------------------------------------------------------
' Convert date to Arezzo syntax i.e. string of the form
'   date(Y,M,D,H,M,S)
'   ATN 19/10/98    SPR 488
'   Modified function to cope with dates in different formats
' Moved here from basProformaEngine - NCJ 17 Nov 99
' NCJ 12/6/00 - Rewritten (faster version)
'----------------------------------------------------------------------
Dim mDate As Date

    On Error GoTo ErrHandler

    mDate = CDate(vDate)
    ' NCJ 12/6/00 - New efficient version
    ConvertDateToProformaSyntax = "date" & Format(mDate, "(yyyy,m,d,h,n,s)")
    
    Exit Function

ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "ConvertDateToProformaSyntax", "modStringUtilities")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'----------------------------------------------------------------
Public Function URLCharToHexEncoding(sBeforeEncoding As String) As String
'----------------------------------------------------------------
'This function is used for the purpose of encoding specific characters
'in strings that are passed to URLs (in our case Active Server Pages)
'The characters are encoded to a %. character followed by a Hex value.
'The encoded characters are automatically un-encoded when received by
'an active server page script.
'Note that for the encoding to work correctly the % character must always
'be encoded first
'
'MLM 28/03/01: Also encode carriage returns and line feeds
'----------------------------------------------------------------
Dim sTemp As String

    
    sTemp = Replace(sBeforeEncoding, "%", "%25")

    sTemp = Replace(sTemp, " ", "%20")
    sTemp = Replace(sTemp, "[", "%5B")
    sTemp = Replace(sTemp, "\", "%5C")
    sTemp = Replace(sTemp, "]", "%5D")
    sTemp = Replace(sTemp, "^", "%5E")
    sTemp = Replace(sTemp, "{", "%7B")
    sTemp = Replace(sTemp, "}", "%7D")
    sTemp = Replace(sTemp, "#", "%23")
    sTemp = Replace(sTemp, "&", "%26")
    sTemp = Replace(sTemp, "+", "%2B")
    sTemp = Replace(sTemp, ",", "%2C")
    sTemp = Replace(sTemp, "/", "%2F")
    sTemp = Replace(sTemp, ":", "%3A")
    sTemp = Replace(sTemp, "<", "%3C")
    sTemp = Replace(sTemp, "=", "%3D")
    sTemp = Replace(sTemp, ">", "%3E")
    sTemp = Replace(sTemp, "?", "%3F")
    sTemp = Replace(sTemp, "@", "%40")
    
    sTemp = Replace(sTemp, vbLf, "%0A")
    sTemp = Replace(sTemp, vbCr, "%0D")
    
    URLCharToHexEncoding = sTemp
    
End Function

'--------------------------------------------------------------------------------------------------------
Public Function HexEncodeChars(ByVal sString As String, ByVal sCharsToEncode As String) As String
'--------------------------------------------------------------------------------------------------------
' MLM 21/06/02: Returns sString but with any characters from sCharsToEncode replaced by a % and a
' 2-character ASCII code in hex.
' sCharsToEncode must NOT contain % or 0 - F; these would make the output undecodable.
' Example: HexEncodeChars("abc%","b") = "a%62c%25"
'--------------------------------------------------------------------------------------------------------

Dim nCount As Integer
Dim sChar As String * 1
    
    On Error GoTo ErrHandler
    
    '%s must always be encoded, and done 1st
    sCharsToEncode = "%" & sCharsToEncode
    
    For nCount = 1 To Len(sCharsToEncode)
        sChar = Mid(sCharsToEncode, nCount, 1)
        'we expect most input strings to only contain few characters to encode, so performance will be
        'improved by doing the Replace() conditionally
        If InStr(sString, sChar) > 0 Then
            sString = Replace(sString, sChar, "%" & IIf(Asc(sChar) < 16, "0", "") & Hex(Asc(sChar)))
        End If
    Next nCount
    
    HexEncodeChars = sString
    
    Exit Function

ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|ModStringUtilities.HexEncodeChars"

End Function

'--------------------------------------------------------------------------------------------------------
Public Function HexDecodeChars(ByRef sString As String) As String
'--------------------------------------------------------------------------------------------------------
' MLM 21/06/02: Do the opposite of HexEncodeChars ;)
' Don't call this with strings that aren't encoded properly, or it might fall over.
'--------------------------------------------------------------------------------------------------------

Dim lPosition As Long
    
    On Error GoTo ErrHandler
    
    HexDecodeChars = sString
    lPosition = InStrRev(sString, "%")
    Do While lPosition > 1
        HexDecodeChars = Mid(HexDecodeChars, 1, lPosition - 1) & _
            Chr(CInt("&h" & Mid(HexDecodeChars, lPosition + 1, 2))) & Mid(HexDecodeChars, lPosition + 3)
            'NB: This takes the 2 characters after the % and converts them from hex to a single character.
        lPosition = InStrRev(sString, "%", lPosition - 1)
    Loop
    If lPosition = 1 Then
        HexDecodeChars = Chr(CInt("&h" & Mid(HexDecodeChars, 2, 2))) & Mid(HexDecodeChars, 4)
    End If
    
    Exit Function

ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|ModStringUtilities.HexDecodeChars"
    
End Function

'--------------------------------------------------------------------------------------------------------
Public Function JSStringLiteral(ByVal sString As String) As String
'--------------------------------------------------------------------------------------------------------
' MLM 18/02/03: Created. Escape certain characters within the input string, so that it can be used as a
'               string literal in JavaScript
' MLM 19/02/03: Added < to encoded characters, because </script> inside a JS string screws things up.
'--------------------------------------------------------------------------------------------------------

'it's important that the \ is first in this string:
Const sCHARS_TO_REPLACE As String = "\""'<" & vbTab & vbCrLf & vbBack & vbNullChar

Dim sChar As String * 1
Dim lCount As Long
    
    For lCount = 1 To Len(sCHARS_TO_REPLACE)
        sChar = Mid(sCHARS_TO_REPLACE, lCount, 1)
        If InStr(sString, sChar) > 0 Then
            sString = Replace(sString, sChar, "\x" & IIf(Asc(sChar) < 16, "0", "") & Hex(Asc(sChar)))
        End If
    Next lCount
    
    JSStringLiteral = sString

End Function

'--------------------------------------------------------------------------------------------------------
Public Function StringtoNumberVariant(sString As String) As Variant
'----------------------------------------------------------------------------------------'
' convert a string into a number variant or null if empty string
'----------------------------------------------------------------------------------------'
    
    sString = Trim(sString)
    If sString = "" Then
        StringtoNumberVariant = Null
    Else
        'val can't cope with commas so use convertlocalnumtostandard
        StringtoNumberVariant = Val(LocalNumToStandard(sString))
    End If

End Function

'----------------------------------------------------------------------------------------'
Public Function StringToSQLString(sString As String)
'----------------------------------------------------------------------------------------'
' convert a variant into a string (if null, "null" or empty string is returned)
'----------------------------------------------------------------------------------------'

    If sString = "" Then
        StringToSQLString = "null"
    Else
        StringToSQLString = "'" & sString & "'"
    End If

End Function

'----------------------------------------------------------------------------------------'
Public Function VarianttoString(vVar As Variant, Optional bForSQL As Boolean = False) As String
'----------------------------------------------------------------------------------------'
' convert a variant into a string (if null, "null" or empty string is returned)
' NCJ 12/10/00 - Include single quotes if a string for SQL
'----------------------------------------------------------------------------------------'

    If VarType(vVar) = vbNull Then
        If bForSQL Then
            VarianttoString = "null"
        Else
            VarianttoString = ""
        End If
    ElseIf VarType(vVar) = vbString Then
        If bForSQL Then
            VarianttoString = "'" & vVar & "'"
        Else
            VarianttoString = vVar
        End If
    Else
        VarianttoString = Format(vVar)
    End If
    
End Function

'----------------------------------------------------------------------------------------'
Public Function SplitCSV(ByVal sInputString As String) As String()
'----------------------------------------------------------------------------------------'
'This function works like the Split() function and is geared to converting
'lines of CSV (comma Separated Values) data into an Array
'
'The rules of CSV are as follows:-
'Strings containing double quotes have the double quotes duplicated.
'Strings containing double quotes have double quotes put around them.
'
'   e.g.    use "MACRO" always      becomes     "use ""MACRO"" always"
'           "this" and "that"       becomes     """this"" and ""that"""
'
'Strings containing a comma have double quotes put around them.
'
'   e.g.    1,500,250               becomes     "1,500,250"
'
'putting both together
'
'   e.g.    "you, what"             becomes     """you, what"""
'           me, "MACRO" and you     becomes     "me, ""MACRO"" and you"
'
'This function reverses the above maniplulations.
'----------------------------------------------------------------------------------------'
Dim i As Integer
Dim sChar As String
Dim bInQuotes As Boolean
Dim sSingleField As String
Dim asArray() As String
Dim nFieldNumber As Integer

   On Error GoTo ErrHandler

    'To make this function work correctly a comma is added at the end of the input string
    sInputString = sInputString & ","

    'Replace occurrences of ,""" with ,"<DQuote>
    sInputString = Replace(sInputString, ",""""""", ",""<DQuote>")
    'Replace occurences of """, with <DQuote>",
    sInputString = Replace(sInputString, """"""",", "<DQuote>"",")
    'Replace occurences of "" with <DQuote>
    sInputString = Replace(sInputString, """""", "<DQuote>")
    
    'initialise the field number used to index the array
    nFieldNumber = 0
    'initialise the variable into which a field is built up
    sSingleField = ""
    
    bInQuotes = False
    For i = 1 To Len(sInputString)
        sChar = Mid(sInputString, i, 1)
        If sChar = """" Then
            'switch true/false, false/true for InQuotes boolean
            bInQuotes = Not bInQuotes
            'drop the double quotes, it is no longer required
            sChar = ""
        End If
        'When a comma occurs and its not Inquotes its the end of a field
        If sChar = "," And Not bInQuotes Then
            'Replace <DQuote> with " in completed field
            sSingleField = Replace(sSingleField, "<DQuote>", """")
            ReDim Preserve asArray(nFieldNumber)
            asArray(nFieldNumber) = sSingleField
            nFieldNumber = nFieldNumber + 1
            'initialise sField ready for the next field
            sSingleField = ""
        Else
            'add current character to current field
            sSingleField = sSingleField & sChar
        End If
    Next i
    
    SplitCSV = asArray
    
Exit Function

ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modStringUtilities.SplitCSV"
    
End Function
