Attribute VB_Name = "modLabDefinitions"
'----------------------------------------------------------------------------------------'
'   File:       modLabDefinitions.bas
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, September 2000
'   Purpose:    genral routines for administering Normal Ranges and CTC Criteria
'----------------------------------------------------------------------------------------'
'Revisions:
'   MLM 27/09/00:  Changed End to Call MACROEnd in error handling routines
'   NCJ 12/10/00 - Changed VariantToString to put quotes round strings for SQL
'   ASH 12/2/2003 - Added check for max values and invalid characters
'   NCJ 14 Feb 03 - Refix - Cannot use Val because it doesn't work with Regional Settings
'----------------------------------------------------------------------------------------'

Option Explicit

'date format used throughout Library Management (user cannot set this yet)
Public Const DEFAULT_DATE_FORMAT = "dd/mm/yyyy"

'in macro 0 is used to mean unspecified date
Public Const UNSPECIFIED_DATE = 0

Public Enum GenderCode
    gNone = 0
    gFemale = 1
    gMale = 2
End Enum

Public Const GENDER_MALE = "M"
Public Const GENDER_FEMALE = "F"
Public Const GENDER_NONE = ""

Public Const NO_CTC_GRADE = -1

Public Enum NRFactor
    nrfabsolute = 0
    nrfLower = 1
    nrfUpper = 2
End Enum

Public Enum NormalRangeLNorH
    nrNotfound = 0
    nrLow = 1
    nrNormal = 2
    nrHigh = 3
    'nrsImpossible = 4
End Enum


'reasons for a normal range not being saved
'used to display reasons
Public Enum ValidRangeStatus
    vreOK = 0
    vreFeasibleinNormal = 1
    vreAbsoluteinFeasible = 2
    vreAbsoluteinNormal = 4
End Enum

'text typed in a text box needs to be validated as this MACRO database column type
Public Enum TextType
    ttAge = 1
    ttDouble = 2
    ttCode = 3
    ttDesc = 4
    ttDate = 5
End Enum

'----------------------------------------------------------------------------------------'
Public Function ValueInRange(Optional vValue As Variant = Null, Optional vMin As Variant = Null, Optional vMax As Variant = Null, _
                                Optional bInclusive As Boolean = True, Optional bAllowNullValue As Boolean = False) As Boolean
'----------------------------------------------------------------------------------------'
' returns true if value is in range
' when Binclusive is true - if bAllowNullValue is true then null is allowed
'   if either end of range is null, otherwise both ends must be null
'----------------------------------------------------------------------------------------'
    
    If VarType(vValue) = vbNull Then
        If bAllowNullValue Then
            'min OR max must be null , inclusive
            ValueInRange = (VarType(vMin) = vbNull) Or (VarType(vMax) = vbNull) And bInclusive
        Else
            'min AND max must be null, inclusive
            ValueInRange = (VarType(vMin) = vbNull) And (VarType(vMax) = vbNull) And bInclusive
        End If
    Else
            'min and max are null                       OR
            'min is null and value less than max        OR
            'max is null and value greater than min     OR
            'value is between min and max
        If bInclusive Then
            'min and max inclusive
            ValueInRange = ((VarType(vMin) = vbNull) And (VarType(vMax) = vbNull)) _
                Or ((VarType(vMin) = vbNull) And (vValue <= vMax)) _
                Or ((VarType(vMax) = vbNull) And (vValue >= vMin) _
                Or ((vValue >= vMin) And (vValue <= vMax)))
        Else
            'min and max exclusive
            ValueInRange = ((VarType(vMin) = vbNull) And (VarType(vMax) = vbNull)) _
                Or ((VarType(vMin) = vbNull) And (vValue < vMax)) _
                Or ((VarType(vMax) = vbNull) And (vValue > vMin) _
                Or ((vValue > vMin) And (vValue < vMax)))
            
        End If
    End If


End Function

'----------------------------------------------------------------------------------------'
Public Function ZeroToNull(ByVal vValue As Variant) As Variant
'----------------------------------------------------------------------------------------'
' returns null if vValue is 0 or null, otherwise returns value
'----------------------------------------------------------------------------------------'
    
    If VarType(vValue) = vbNull Then
        ZeroToNull = Null
    Else
        If vValue = 0 Then
            ZeroToNull = Null
        Else
            ZeroToNull = vValue
        End If
    End If


End Function


''----------------------------------------------------------------------------------------'
'Public Function RangeInRange(Optional vRange1Min As Variant = Null, Optional vRange1Max As Variant = Null, _
'                                Optional vRange2Min As Variant = Null, Optional vRange2Max As Variant = Null, Optional bInclusive As Boolean = True) As Boolean
''----------------------------------------------------------------------------------------'
''returns true if range 1 inside range 2
''----------------------------------------------------------------------------------------'
'
'   'true if min1 in range2 or max1 in range2
'    If (VarType(vRange1Min) = vbNull And VarType(vRange1Max) = vbNull) And (VarType(vRange2Min) = vbNull And VarType(vRange2Max) = vbNull) Then
'        'range 1 or 2 isn't a range - don't bother checking
'        RangeInRange = bInclusive
'    Else
'        If (VarType(vRange1Min) = vbNull And VarType(vRange2Min) <> vbNull) Or (VarType(vRange1Max) = vbNull And VarType(vRange2Max) <> vbNull) Then
'            'range 1 not bound and range 2 is
'            RangeInRange = False
'        Else
'            RangeInRange = ValueInRange(vRange1Min, vRange2Min, vRange2Max, bInclusive, True) And ValueInRange(vRange1Max, vRange2Min, vRange2Max, bInclusive, True)
'        End If
'    End If
'
'End Function

''----------------------------------------------------------------------------------------'
'Public Function RangeOverlap(Optional vRange1Min As Variant = Null, Optional vRange1Max As Variant = Null, _
'                                Optional vRange2Min As Variant = Null, Optional vRange2Max As Variant = Null) As Boolean
''----------------------------------------------------------------------------------------'
''returns true if range 1 and range 2 overlap
''----------------------------------------------------------------------------------------'
'
''nb. min and max inclusive
'   RangeOverlap = ValueInRange(vRange1Min, vRange2Min, vRange2Max) Or ValueInRange(vRange1Max, vRange2Min, vRange2Max) _
'                    Or ValueInRange(vRange2Min, vRange1Min, vRange1Max)
'
'End Function

'----------------------------------------------------------------------------------------'
Public Function ValidRange(ByVal vMin As Variant, ByVal vMax As Variant) As Boolean
'----------------------------------------------------------------------------------------'
'returns true if range min is greater than max
'----------------------------------------------------------------------------------------'
Dim bValidRange As Boolean
    
    bValidRange = True
    If Not (VarType(vMin) = vbNull Or VarType(vMax) = vbNull) Then
        If vMin > vMax Then
            bValidRange = False
        End If
    End If
    
    ValidRange = bValidRange
            

End Function

'----------------------------------------------------------------------------------------'
Public Function CTCExpr(sMin As String, sMax As String, nMinType As NRFactor, nMaxType As NRFactor, sUnits As String) As String
'----------------------------------------------------------------------------------------'
' Return a CTC string expression based on min , max and type
'----------------------------------------------------------------------------------------'
Dim sMinExpr As String
Dim sMaxExpr As String

    If sMin = "" And sMax = "" Then
        'show nothing
        CTCExpr = ""
    Else
        If sMin <> "" Then
            Select Case nMinType
            Case NRFactor.nrfabsolute: sMinExpr = " >= " & sMin
            Case NRFactor.nrfLower: sMinExpr = " >= " & sMin & " x " & "LLN"
            Case NRFactor.nrfUpper: sMinExpr = " >= " & sMin & " x " & "ULN"
            End Select
            sMinExpr = Replace(sMinExpr, ">= 1 x ", "> ")
        End If
        
        
        If sMax <> "" Then
            Select Case nMaxType
            Case NRFactor.nrfabsolute: sMaxExpr = " <= " & sMax
            Case NRFactor.nrfLower: sMaxExpr = " <= " & sMax & " x " & "LLN"
            Case NRFactor.nrfUpper: sMaxExpr = " <= " & sMax & " x " & "ULN"
            End Select
            sMaxExpr = Replace(sMaxExpr, "<= 1 x ", "< ")
        End If
        
        
        CTCExpr = sMinExpr
        If sMinExpr <> "" And sMaxExpr <> "" Then
            CTCExpr = CTCExpr & " - "
        End If
        
        CTCExpr = CTCExpr & sMaxExpr & " " & sUnits
   
    End If

End Function

'----------------------------------------------------------------------------------------'
Public Function RangeExpr(sMin As String, sMax As String) As String
'----------------------------------------------------------------------------------------'
' Return a range expression string expression based on min and max
'----------------------------------------------------------------------------------------'
Dim sMinExpr As String
Dim sMaxExpr As String

    If sMin = "" And sMax = "" Then
        RangeExpr = ""
    Else
        If sMin <> "" And sMax <> "" Then
            RangeExpr = sMin & " to " & sMax
        Else
            If sMin <> "" Then
                RangeExpr = ">=" & sMin
            Else
                RangeExpr = "<=" & sMax
            End If
        End If
    End If
    
End Function


'----------------------------------------------------------------------------------------'
Public Function GetValidRangeStatusText(nError As ValidRangeStatus) As String
'----------------------------------------------------------------------------------------'
' TA 21/09/2000: retunr status text according to validrangestatus
'----------------------------------------------------------------------------------------'
Dim sMessage As String

    On Error GoTo ErrHandler
    
    If nError And ValidRangeStatus.vreFeasibleinNormal Then
        sMessage = sMessage & vbCrLf & "Please ensure that the normal range is within the feasible range."
    End If
    
    If nError And ValidRangeStatus.vreAbsoluteinNormal Then
        sMessage = sMessage & vbCrLf & "Please ensure that the normal range is within the absolute range."
    End If
    
    If nError And ValidRangeStatus.vreAbsoluteinFeasible Then
        sMessage = sMessage & vbCrLf & "Please ensure that the feasible range is within the absolute range."
    End If

    GetValidRangeStatusText = sMessage
    
Exit Function
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetValidRangeStatusText", "modLabDefinition")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Function

Public Function GetGenderText(nGender As GenderCode) As String
'----------------------------------------------------------------------------------------'
'return gender text from code
'----------------------------------------------------------------------------------------'

    GetGenderText = Switch(nGender = GenderCode.gMale, GENDER_MALE, nGender = GenderCode.gFemale, GENDER_FEMALE, nGender = GenderCode.gNone, GENDER_NONE)
End Function

Public Function GetGenderCode(sGender As String) As GenderCode
'----------------------------------------------------------------------------------------'
' return gender code from text expression
'----------------------------------------------------------------------------------------'

    GetGenderCode = Switch(sGender = GENDER_MALE, GenderCode.gMale, sGender = GENDER_FEMALE, GenderCode.gFemale, sGender = GENDER_NONE, GenderCode.gNone)
    
End Function


'----------------------------------------------------------------------------------------'
Public Function GetNRCTCText(nNormalRangeLNorH As Variant, nCTCGrade As Variant) As String
'----------------------------------------------------------------------------------------'
' return the test for putting in display label in frmCRFDataEntry
'----------------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler
    
    'convert null values to nrNotFound and NO_CTC_GRADE
    If VarType(nNormalRangeLNorH) = vbNull Then
        nNormalRangeLNorH = NormalRangeLNorH.nrNotfound
    End If
    
    ' NCJ 6/10/00 - Bug fix (changed nNormalRangeLNorH to nCTCGrade)
    If VarType(nCTCGrade) = vbNull Then
        nCTCGrade = NO_CTC_GRADE
    End If
    
    GetNRCTCText = ""
    
    Select Case nNormalRangeLNorH
    Case nrLow: GetNRCTCText = "L"
    Case nrNormal: GetNRCTCText = "N"
    Case nrHigh: GetNRCTCText = "H"
    End Select
    
    If nCTCGrade <> -1 Then
        GetNRCTCText = GetNRCTCText & Format(nCTCGrade)
    End If
    
Exit Function
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetNRCTCText", "modLabDefinition")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'----------------------------------------------------------------------------------------'
Public Function ValidText(ByVal sText As String, nTextType As TextType) As Boolean
'----------------------------------------------------------------------------------------'
'returns true if the text passed in is valid acording to text type
'----------------------------------------------------------------------------------------'
Dim i As Long
Dim dblNum As Double
Dim sStdText As String

'ASH 12/2/2003
Const dblMIN_MAX_VALUE As Double = 999999.9
    
    On Error GoTo ErrHandler
    
    Select Case nTextType
    Case TextType.ttAge: ValidText = gblnValidString(sText, valNumeric)
    Case TextType.ttCode: ValidText = gblnValidString(sText, valAlpha + valNumeric + valUnderscore) And (sText <> "") And StartsWithAlpha(sText)
    Case TextType.ttDesc: ValidText = gblnValidString(sText, valAlpha + valNumeric + valUnderscore + valSpace + valOnlySingleQuotes + valComma + valUnderscore) And (sText <> "")
    Case TextType.ttDouble
    
        ' Allow empty fields
        If Trim(sText) = "" Then
            ValidText = True
            Exit Function
        End If
        
        ValidText = False
    
        If Not IsNumeric(sText) Then Exit Function
        
        ' Screen out "regional" number separators
        sStdText = ConvertLocalNumToStandard(sText)
        For i = 1 To Len(sStdText)
            If Not (Mid(sStdText, i, 1) Like "[-0123456789.]") Then Exit Function
        Next
        
        ' NCJ 14 Feb 03 - Do not use Val because doesn't work with regional settings
        dblNum = CDbl(sText)
        
        'ASH 12/2/2003 Do not allow numbers greater than dblMIN_MAX_VALUE
        If Abs(dblNum) > dblMIN_MAX_VALUE Then Exit Function
        
        'ASH 12/2/2003 Do not allow "-" after or within numbers or "-" alone as number
        'also do not allow "-." or "." alone except when followed by a number
        ValidText = Not (CountStrsInStr(sStdText, ".") > 1 Or InStr(2, sStdText, "-") > 0 _
                    Or sStdText = "-" Or sStdText = "-." Or sStdText = ".")
                    
    Case TextType.ttDate: 'unused as different calls needed for SD and DM
    End Select
    
    
Exit Function

ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ValidText", "modLabDefinition")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

Public Function ValidTextBox(txtTextBox As TextBox, nTextType As TextType) As Boolean
'----------------------------------------------------------------------------------------'
'set text box to yellow if invalid
'----------------------------------------------------------------------------------------'
    If ValidText(txtTextBox.Text, nTextType) Then
        'valid
        txtTextBox.BackColor = vbWindowBackground
    Else
        txtTextBox.BackColor = vbYellow
    End If
    
End Function

Public Function ValidRangeTextBoxes(txtMin As TextBox, txtMax As TextBox) As Boolean
'----------------------------------------------------------------------------------------'
'set text box to yellow if invalid
'----------------------------------------------------------------------------------------'
    If Not ValidRange(StringtoNumberVariant(txtMin.Text), StringtoNumberVariant(txtMax.Text)) Then
        'invalid
        txtMin.BackColor = vbYellow
        txtMax.BackColor = vbYellow
    End If
    
End Function
