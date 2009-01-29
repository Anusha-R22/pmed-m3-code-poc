Attribute VB_Name = "modMACRODM"
'----------------------------------------------------
' File: modMACRODM.bas
' Copyright: InferMed, 2001 - 2006, All Rights Reserved
' Author: Nicky Johns, InferMed, June 2001
' Purpose: General routines for MACRO Data Management
'       (Including enumerations)
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 25 Jun 01 - Initial development
' NCJ 4 Jul 01 - Added eResponseValidation
' NCJ 15 Aug 01 - Changed to eLockStatus
' NCJ 27 Sep 01 - eStudyStatus
' NCJ 12 Oct 01 - GetNRText
' ic 14/07/2005  added clinical coding
' NCJ 23 Jan 06 - Added date formatting which does not use AREZZO (e.g. for Web form/visit dates)
'----------------------------------------------------

Option Explicit

' Question data types
Public Enum eDataType
    Text = 0
    Category = 1
    IntegerNumber = 2
    Real = 3
    DateTime = 4
    MultiMedia = 5
    LabTest = 6
    
    'ic 14/07/2005 added clinical coding
    Thesaurus = 8
    
End Enum

' Status of Responses, eFormInstances and VisitInstances
Public Enum eStatus
    CancelledByUser = -20
    Requested = -10
    NotApplicable = -8
    Unobtainable = -5
    Success = 0
    Missing = 10
    Inform = 20
    OKWarning = 25
    Warning = 30
    InvalidData = 40
End Enum

' Lock status
Public Enum eLockStatus
    lsUnlocked = 0
    lsPending = 3
    lsLocked = 5
    lsFrozen = 6
End Enum


Public Enum GenderCode
    gNone = 0
    gFemale = 1
    gMale = 2
End Enum


' In MACRO 0 is used to mean unspecified date
Public Enum eMACRODate
    mdUnspecified = 0
End Enum


' Possible results of validating a response value
' NB The Arezzo results MUST have the values 1,-1,-2 and -3
Public Enum eResponseValidation
    ValueOK = 0
    ' NB The Arezzo values must be exactly as follows
    ArezzoWarning = 1
    ArezzoTypeError = -1
    ArezzoRangeError = -2
    ArezzoValidationError = -3
    NotANumber = 2
    NotAnInteger = 3
    NumberTooBig = 4
    NumberTooSmall = 5
    NumberNotPositive = 6
    TextWrongFormat = 7
    NotADateTime = 8
    ValidationReject = 9
    ValueUnchanged = 11
End Enum

' Type of validation result for questions
' NCJ 9 Jul 01
Public Enum eValidationType
    Reject = 0
    Warn = 1
    Inform = 2
End Enum

' NCJ 27/9/01
Public Enum eStudyStatus
    InPreparation = 1
    TrialOpen = 2
    ClosedToRecruitment = 3
    ClosedToFollowUp = 4
    Suspended = 5
End Enum

'ic 18/07/2005 added clinical coding
Public Enum eCodingStatus
    csEmpty = 0
    csNotCoded = 1
    csCoded = 2
    csPendingNewCode = 3
    csAutoEncoded = 4
    csValidated = 5
    csDoNotCode = 6
End Enum

' NCJ 23 Jan 06 - These constants are COPIED from the ALM
Const mdblYEAR_ONLY As Double = 800000
Const mdblYEAR_MONTH As Double = 400000
Const mdblMAX_FULL_DATE As Double = 290429

'property gets to represent constants
Public Property Get GENDER_MALE()
    GENDER_MALE = "M"
End Property

Public Property Get GENDER_FEMALE()
    GENDER_FEMALE = "F"
End Property

Public Property Get GENDER_NONE()
    GENDER_NONE = ""
End Property






'----------------------------------------------------
Public Function GetDataTypeString(nDataType As Integer) As String
'----------------------------------------------------
' Data type as a string (for display purposes)
' ic 14/07/2005 added clinical coding
'----------------------------------------------------

    Select Case nDataType
    Case eDataType.Text
        GetDataTypeString = "Text"
    Case eDataType.Category
        GetDataTypeString = "Category"
    Case eDataType.IntegerNumber
        GetDataTypeString = "Integer number"
    Case eDataType.Real
        GetDataTypeString = "Real number"
    Case eDataType.DateTime
        GetDataTypeString = "Date/Time"
    Case eDataType.MultiMedia
        GetDataTypeString = "Multimedia"
    Case eDataType.LabTest
        GetDataTypeString = "Lab Test"
    Case eDataType.Thesaurus
        GetDataTypeString = "Thesaurus"

    Case Else
        GetDataTypeString = "UNKNOWN"
    End Select
    
End Function

'----------------------------------------------------
Public Function GetStatusString(nStatus As Integer) As String
'----------------------------------------------------
' Status of a response, eform instance, visit instance
'----------------------------------------------------

    Select Case nStatus
    Case eStatus.CancelledByUser
        GetStatusString = "Cancelled"
    Case eStatus.Requested
        GetStatusString = "Requested"
    Case eStatus.NotApplicable
        GetStatusString = "Not Applicable"
    Case eStatus.Unobtainable
        GetStatusString = "Unobtainable"
    Case eStatus.Success
        GetStatusString = "Success"
    Case eStatus.Missing
        GetStatusString = "Missing"
    Case eStatus.Inform
        GetStatusString = "Inform"
    Case eStatus.OKWarning
        GetStatusString = "OK Warning"
    Case eStatus.Warning
        GetStatusString = "Warning"
    Case eStatus.InvalidData
        GetStatusString = "Invalid"
    Case Else
        GetStatusString = "UNKNOWN"
    End Select
    
End Function

'----------------------------------------------------
Public Function GetLockStatusString(nLockStatus As Integer) As String
'----------------------------------------------------
' Lock status as string
'----------------------------------------------------

    Select Case nLockStatus
    Case eLockStatus.lsFrozen
        GetLockStatusString = "Frozen"
    Case eLockStatus.lsLocked
        GetLockStatusString = "Locked"
    Case eLockStatus.lsPending
        GetLockStatusString = "Pending"
    Case eLockStatus.lsUnlocked
        GetLockStatusString = "Unlocked"
    Case Else
        GetLockStatusString = "UNKNOWN"
    End Select
    
End Function

'----------------------------------------------------
Public Function GetControlTypeString(nControlType As Integer) As String
'----------------------------------------------------
' EForm Element control type as a string (for display purposes)
'----------------------------------------------------

    Select Case nControlType
    Case 1
        GetControlTypeString = "Text box"
    Case 2
        GetControlTypeString = "Option buttons"
    Case 4
        GetControlTypeString = "Popup list"
    Case 8
        GetControlTypeString = "Calendar"
    Case 16385
        GetControlTypeString = "Line"
    Case 16386
        GetControlTypeString = "Text comment"
    Case 16388
        GetControlTypeString = "Picture"
    Case Else
        GetControlTypeString = CStr(nControlType)
    End Select

End Function

'----------------------------------------------------
Public Function GetResponseErrorString(lResult As Long) As String
'----------------------------------------------------
' Translation of eResponseValidation value into string
'----------------------------------------------------
Dim sErrorMessage As String

    Select Case lResult
    Case eResponseValidation.NotADateTime
        sErrorMessage = "This is not a valid date/time value"
    Case eResponseValidation.NotAnInteger
        sErrorMessage = "This is not a valid integer value"
    Case eResponseValidation.NotANumber
        sErrorMessage = "This is not a valid number value"
    Case eResponseValidation.TextWrongFormat, _
            eResponseValidation.ArezzoTypeError
        sErrorMessage = "This is not in the correct format"
    Case eResponseValidation.NumberNotPositive
        sErrorMessage = "This is not a positive number"
    Case eResponseValidation.NumberTooBig
        sErrorMessage = "This is bigger than the value allowed for this question"
    Case eResponseValidation.NumberTooSmall
        sErrorMessage = "This is smaller than the value allowed for this question"
    Case eResponseValidation.ArezzoRangeError
        sErrorMessage = "This is not one of the allowed values for this question"
    Case Else
        sErrorMessage = ""
    End Select
    
    GetResponseErrorString = sErrorMessage
    
End Function

'----------------------------------------------------
Public Function GetNRText(lNRStatus As Integer) As String
'----------------------------------------------------
' NCJ 12 Oct 01 - Convert an integer NRStatus to a letter
'----------------------------------------------------
    
    Select Case lNRStatus
    Case eNormalRangeLNorH.nrLow
        GetNRText = "L"
    Case eNormalRangeLNorH.nrNormal
        GetNRText = "N"
    Case eNormalRangeLNorH.nrHigh
        GetNRText = "H"
    Case Else
        GetNRText = ""
    End Select

End Function

'--------------------------------------------------------------
Public Function VBFormatPartialDate(dblDate As Double, sFormat As String) As String
'--------------------------------------------------------------
' NCJ 23 Jan 06 - Format a possibly partial date using VB only, i.e. without using AREZZO
' dblDate may represent full date, year-only, or year-month-only
' Technique for munging a partial double copied from ALM
'--------------------------------------------------------------
Dim sDate As String

    On Error GoTo ErrLabel
    
    If dblDate > mdblYEAR_ONLY Then
        ' Year only
        sDate = CStr((dblDate - mdblYEAR_ONLY))
    ElseIf dblDate > mdblMAX_FULL_DATE Then
        ' Year/month
        dblDate = dblDate - mdblYEAR_MONTH
        ' Remove the d (or dd) from the format
        sDate = Format(CDate(dblDate), MakeFormatMY(sFormat))
    Else
        ' Ordinary full date
        sDate = Format(CDate(dblDate), sFormat)
    End If

    VBFormatPartialDate = sDate

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modMACRODM.VBFormatPartialDate"

End Function

'--------------------------------------------------------------
Private Function MakeFormatMY(sFormat As String) As String
'--------------------------------------------------------------
' Remove the dd part of a date format to make it month-year,
' preserving order, separators etc.
'--------------------------------------------------------------
Dim sNewFormat As String
Dim nPos As Integer

    On Error GoTo ErrLabel
    
    nPos = 1
    sNewFormat = ""
    Select Case Left(sFormat, 1)
    Case "y"
        ' Assume "yyyy/mx"
        ' Take the yyyy/m
        sNewFormat = sNewFormat & "yyyy" & Mid(sFormat, nPos + 4, 2)
        If Len(sFormat) >= nPos + 6 Then
            If Mid(sFormat, nPos + 6, 1) = "m" Then
                ' Add on the extra m
                sNewFormat = sNewFormat & "m"
            End If
        End If
    Case "m"
        ' Assume m/d/y or m/y, becomes m/y
        If Mid(sFormat, nPos + 1, 1) = "m" Then
            sNewFormat = sNewFormat & "mm"
            nPos = nPos + 2
        Else
            sNewFormat = sNewFormat & "m"
            nPos = nPos + 1
        End If
        ' Insert separator and year
        sNewFormat = sNewFormat & Mid(sFormat, nPos, 1) & "yyyy"
    Case "d"
        ' Ignore the d and its separator
        If Mid(sFormat, nPos + 1, 1) = "d" Then
            nPos = nPos + 2
        Else
            nPos = nPos + 1
        End If
        ' Take what's left after ignoring the d's
        sNewFormat = Right(sFormat, Len(sFormat) - nPos)
    End Select
        
    MakeFormatMY = sNewFormat

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modMACRODM.MakeFormatMY"

End Function

