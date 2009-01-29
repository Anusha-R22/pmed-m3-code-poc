VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabtestValidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Validation class for Normal Ranges and CTC"
Attribute VB_Description = "Validation class for Normal Ranges and CTC"
'----------------------------------------------------------------------------------------'
'   File:       clsLabTestValidate.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, August 2000
'   Purpose:    Validation class for Normal Ranges and CTC
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 27/9/00 - Convert dates to NULL if 0
'   TA 05/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
'   TA 16/10/2000: CTCs based on normal range that fall in normal range are no longer returned
'   TA 25/10/2000: Previous revision revised - doesn't return a CTC grade if value equals 1*LLN or 1*ULN
'   TA 24/11/2000: Null values for CTC Max, CTC min, ULN and LLN are now checked for
'   TA 24/11/2000: New Age function to replace datediff
'   TA 02/10/2001 - New error handling
'   NCJ 3/10/2001 - Changed SetState to take State object
'----------------------------------------------------------------------------------------'

Option Explicit


Private moState As LabTestValidateState

'----------------------------------------------------------------------------------------'
Friend Sub SetState(oLTVState As LabTestValidateState)
'----------------------------------------------------------------------------------------'
' Set the state of this object
'----------------------------------------------------------------------------------------'

    Set moState = oLTVState

End Sub

'----------------------------------------------------------------------------------------'
Private Function GetNormalRange(dblTestDate As Double, dblDOB As Double, nGender As Integer) As NormalRangeRO
'----------------------------------------------------------------------------------------'
' Assumes collection of NRStates can only contain a 'none' gender range set if no male of female record exist.
' Returns nothing if no corresponding NR exists.
'----------------------------------------------------------------------------------------'
Dim vTestDate As Variant
Dim vDOB As Variant
Dim vAge As Variant
Dim oNormalRange As NormalRangeRO
Dim oNRState As NormalRangeState

    On Error GoTo ErrLabel
    
    'convert macro dates of 0 to null
    vTestDate = ConvertToNull(dblTestDate, DATATYPE_DATE)
    vDOB = ConvertToNull(dblDOB, DATATYPE_DATE)
    ' Calculate the age based on the Test Date
    If VarType(vTestDate) = vbNull Or VarType(vDOB) = vbNull Then
        vAge = Null
    Else
        vAge = Age(vDOB, vTestDate)
    End If
    

    'loop through NR states
    For Each oNRState In moState.NormalRangeStates
        With oNRState
            If RangeIncludesValue(vTestDate, _
                            ConvertToNull(.EffectiveStart, DATATYPE_DATE), _
                            ConvertToNull(.EffectiveEnd, DATATYPE_DATE)) Then
                'correct date
                If RangeIncludesValue(vAge, .AgeMin, .AgeMax) Then
                    'correct age
                    If (nGender = .GenderCode) Or (.GenderCode = GenderCode.gNone) Then
                        'correct gender or returned gender is 'none'
                        Set oNormalRange = New NormalRangeRO
                        Call oNormalRange.SetState(oNRState)
                        Exit For
                    End If
    
                End If
            End If
        End With

    Next
    
    Set GetNormalRange = oNormalRange

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "LabTestValidate.GetNormalRange"
    
End Function

'----------------------------------------------------------------------------------------'
Private Function GetCTC() As Collection
'----------------------------------------------------------------------------------------'
' Collection of CTC (can have 0 count)
'----------------------------------------------------------------------------------------'
Dim oCTCState As CTCriterionState
Dim oCTC As CTCriterionRO
Dim colCTC As Collection

    On Error GoTo ErrLabel
    
    Set colCTC = New Collection
    
    'loop through CTC states
    For Each oCTCState In moState.CTCStates
        Set oCTC = New CTCriterionRO
        oCTC.SetState oCTCState
        'add to collection
        colCTC.Add oCTC
    Next
    
    Set GetCTC = colCTC

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "LabTestValidate.GetCTC"
    
End Function


'----------------------------------------------------------------------------------------'
Private Function GetCTCKey(lId As Long) As String
'----------------------------------------------------------------------------------------'
' The key for an item by its CTC id
'----------------------------------------------------------------------------------------'

    GetCTCKey = "K" & Format(lId)

End Function

'----------------------------------------------------------------------------------------'
Public Function NormalRangeStatus(dblResult As Double, dblTestDate As Double, dblDOB As Double, nGender As Integer, sMessage As String) As Integer
Attribute NormalRangeStatus.VB_Description = "Returns a MACRO status and relevenat error message in sMessage. dblResult is the user's reponse."
'----------------------------------------------------------------------------------------'
' Returns a MACRO status and relevenat error message in sMessage.
' dblResult is the user's reponse.
'----------------------------------------------------------------------------------------'
' NCJ 26/9/00 - Corrected message wording
' TA 06/10/2000: lab code now in message
'----------------------------------------------------------------------------------------'
Dim nStatus As eStatus
Dim oNormalRange As NormalRangeRO

    On Error GoTo ErrLabel
    
    Set oNormalRange = GetNormalRange(dblTestDate, dblDOB, nGender)
    If oNormalRange Is Nothing Then
        'no range validation as there's no range to check against
        nStatus = eStatus.Success
    Else
        With oNormalRange
            If RangeIncludesValue(dblResult, .NormalMin, .NormalMax) Then
                'inside normal
                nStatus = eStatus.Success
            Else
                'outside normal
                If .HasAbsoluteRange And Not RangeIncludesValue(dblResult, .AbsoluteMinValue, .AbsoluteMaxValue) Then
                    'outside absolute
                    sMessage = "The value is outside the Laboratory's Absolute Range (" & .LabCode & ")" & vbCrLf & .AbsoluteRangeText
                    nStatus = eStatus.InvalidData
                Else
                    'outside normal and inside absolute or no abolute
                    If .HasFeasibleRange And (RangeIncludesValue(dblResult, .FeasibleMinValue, .FeasibleMaxValue)) Then
                            'has a feasible range and is in it
                            nStatus = eStatus.Success
                    Else
                        'outside feasible or feasible does not exist
                        If .HasFeasibleRange Then
                            'outside feasible and inside absolute or absolute does not exsit
                             sMessage = "The value is outside the Laboratory's Feasible Range (" & .LabCode & ")" & vbCrLf & .FeasibleRangeText
                             nStatus = eStatus.Warning
                        Else
                            'must be outside normal and inside absolute and feasible does not exsit
                            sMessage = "The value is outside the Laboratory's Normal Range (" & .LabCode & ")" & vbCrLf & .NormalRangeText
                            nStatus = eStatus.Warning
                        End If
                    End If
                End If
            End If
        End With
    End If

    NormalRangeStatus = nStatus

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "LabTestValidate.NormalRangeStatus"

End Function

'----------------------------------------------------------------------------------------'
Public Function NormalRangeResult(dblResult As Double, dblTestDate As Double, dblDOB As Double, nGender As Integer) As eNormalRangeLNorH
Attribute NormalRangeResult.VB_Description = "Returns a Normal Range Status of 'not found', 'low', 'normal', or 'high'"
'----------------------------------------------------------------------------------------'
' Returns a Normal Range Status of 'not found', 'low', 'normal', or 'high'
'----------------------------------------------------------------------------------------'
Dim oNormalRange As NormalRangeRO

    On Error GoTo ErrLabel
    Set oNormalRange = GetNormalRange(dblTestDate, dblDOB, nGender)
    
    If oNormalRange Is Nothing Then
        'no range validation as there's no range to check against
       NormalRangeResult = eNormalRangeLNorH.nrNotfound
    Else
        If RangeIncludesValue(dblResult, oNormalRange.NormalMin, oNormalRange.NormalMax) Then
            'inside normal
            NormalRangeResult = eNormalRangeLNorH.nrNormal
        Else
            If dblResult > oNormalRange.NormalMax Then
                'above normal max
                NormalRangeResult = eNormalRangeLNorH.nrHigh
            Else
                'must be below normal min (or normal max is null)
                NormalRangeResult = eNormalRangeLNorH.nrLow
            End If
        End If
    End If

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "LabTestValidate.NormalRangeResult"

End Function

'----------------------------------------------------------------------------------------'
Public Function CTCGrade(dblResult As Double, dblTestDate As Double, dblDOB As Double, nGender As Integer) As Integer
Attribute CTCGrade.VB_Description = "Returns the first matching CTC Grade (grades are in grade order) or eCTCGrade.ctcNone for no match. (doesn't return a grade if value equals 1*LLN or 1*ULN)"
'----------------------------------------------------------------------------------------'
' Returns the first matching CTC Grade (grades are in grade order)
' or eCTCGrade.ctcNone for no match.
'(doesn't return a grade if value equals 1*LLN or 1*ULN)
'----------------------------------------------------------------------------------------'
'TA 22/11/2000: additional checks for unspecifed LLN, ULN, CTC Min and CTC max
'----------------------------------------------------------------------------------------'
Dim oCTCriterion As CTCriterionRO
Dim vMin As Variant
Dim vMax As Variant
Dim bCalculate As Boolean
Dim colCTC As Collection
Dim oNormalRange As NormalRangeRO
    
    On Error GoTo ErrLabel
    
    'set up initial value
    CTCGrade = eCTCGrade.ctcNone
    
    Set oNormalRange = GetNormalRange(dblTestDate, dblDOB, nGender)
    
    Set colCTC = GetCTC
    For Each oCTCriterion In colCTC
        With oCTCriterion
            'TA 25/10/2000: If CTC not absolute then there must be a normal range
            If (NormalRangeResult(dblResult, dblTestDate, dblDOB, nGender) <> nrNotfound) Or _
                    ((.MinType = eNRFactor.nrfAbsolute) And (.MaxType = eNRFactor.nrfAbsolute)) Then
                'normal range exists or min and max absolute values
                bCalculate = True
                'establish min value
                Select Case .MinType
                Case eNRFactor.nrfAbsolute
                    vMin = .Min
                Case eNRFactor.nrfLower
                    vMin = .Min * oNormalRange.NormalMin
                    If IsNull(oNormalRange.NormalMin) Then
                        'related to LLN and there is no LLN so we are within NR
                        bCalculate = False
                    Else
                        If Not IsNull(.Min) Then
                            'there is a CTC min
                            'calculate if dblResult <> 1 * LLN
                             bCalculate = Not (dblResult = vMin And .Min = 1)
                        End If
                    End If
                Case eNRFactor.nrfUpper
                    vMin = .Min * oNormalRange.NormalMax
                    If IsNull(oNormalRange.NormalMax) Then
                         'related to ULN and there is no ULN so we are within NR
                        bCalculate = False
                    Else
                        If Not IsNull(.Min) Then
                            'there is a CTC min
                            'calculate if dblResult <> 1 * ULN
                            bCalculate = Not (dblResult = vMin And .Min = 1)
                        End If
                    End If
                End Select

                If bCalculate Then
                    'if we haven't yet ruled out this criterion
                    'establish max value
                    Select Case .MaxType
                    Case eNRFactor.nrfAbsolute: vMax = .Max
                    Case eNRFactor.nrfLower
                        vMax = .Max * oNormalRange.NormalMin
                        If IsNull(oNormalRange.NormalMin) Then
                            'related to LLN and ther is no LLN so we are within NR
                        Else
                            If Not IsNull(.Max) Then
                                'there is a CTC max
                                'calculate if dblResult <> 1 * LLN
                                 bCalculate = Not (dblResult = vMax And .Max = 1)
                            End If
                        End If
                    Case eNRFactor.nrfUpper
                        vMax = .Max * oNormalRange.NormalMax
                        If IsNull(oNormalRange.NormalMax) Then
                            'related to ULN and there is no ULN so we are within NR
                            bCalculate = False
                        Else
                            If Not IsNull(.Max) Then
                                'there is a CTC Max
                                'calculate if dblResult <> 1 * ULN
                                 bCalculate = Not (dblResult = vMax And .Max = 1)
                            End If
                        End If
                    End Select
                End If
                If bCalculate Then
                    If RangeIncludesValue(dblResult, vMin, vMax) Then
                        CTCGrade = .Grade
                        Exit Function
                    End If
                End If
            End If
        End With
    Next

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "LabTestValidate.CTCGrade"

End Function

'----------------------------------------------------------------------------------------'
Private Function Age(vDOB As Variant, vDate As Variant) As Long
'----------------------------------------------------------------------------------------'
' return age in years given two dates
'----------------------------------------------------------------------------------------'
Dim lAge As Long

    On Error GoTo ErrLabel
    
    lAge = DateDiff("yyyy", vDOB, vDate)

    If DatePart("m", vDate) < DatePart("m", vDOB) Then
        'current month is before birthday
        lAge = lAge - 1
    Else
        If DatePart("m", vDate) = DatePart("m", vDOB) Then
            'current month is birth month
            If DatePart("d", vDate) < DatePart("d", vDOB) Then
                'current day before birthday
                lAge = lAge - 1
            End If
        End If
    End If

    Age = lAge

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "LabTestValidate.Age"

End Function

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'
    
    Set moState = New LabTestValidateState

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'
    
    Set moState = Nothing

End Sub
 
 
