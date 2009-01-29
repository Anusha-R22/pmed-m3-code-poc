VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NormalRangeRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: NormalRangeRO.cls
' Author: Toby Aldridge
' Copyright: InferMed, September 2001, All Rights Reserved
' Normal Range Object for MACRO 2.2 Data Entry
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'   NCJ 24 Mar 04 - Added Terminate event
'--------------------------------------

Option Explicit

Public Enum eNormalRangeLNorH
    nrNotfound = 0
    nrLow = 1
    nrNormal = 2
    nrHigh = 3
    'nrsImpossible = 4
End Enum

Private moState As NormalRangeState

'----------------------------------------------------------------------------------------'
Friend Sub SetState(oState As NormalRangeState)
Attribute SetState.VB_Description = "set up reference to state object"
'----------------------------------------------------------------------------------------'
' set up reference to state object
'----------------------------------------------------------------------------------------'

    Set moState = oState
    
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get AbsoluteMinValue()
Attribute AbsoluteMinValue.VB_Description = "The absolute (calculated) value for the absolute min"
'----------------------------------------------------------------------------------------'
' The absolute (calculated) value for the absolute min
'----------------------------------------------------------------------------------------'

    With moState
        If AbsolutePercent Then
            AbsoluteMinValue = NormalMin - ((AbsoluteMin / 100) * NormalMin)
        Else
            AbsoluteMinValue = AbsoluteMin
        End If
    End With
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get AbsoluteMaxValue()
Attribute AbsoluteMaxValue.VB_Description = "The absolute (calculated) value for the absolute max."
'----------------------------------------------------------------------------------------'
' The absolute (calculated) value for the absolute max.
'----------------------------------------------------------------------------------------'
'TA 22/11/2000: absolute max added to the percentage
    If AbsolutePercent Then
        AbsoluteMaxValue = NormalMax + (Absolutemax / 100) * NormalMax
    Else
        AbsoluteMaxValue = Absolutemax
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get FeasibleMinValue()
Attribute FeasibleMinValue.VB_Description = "The absolute value for the feasible min"
'----------------------------------------------------------------------------------------'
' The absolute value for the feasible min
'----------------------------------------------------------------------------------------'
'TA 15/12/2000: subtract percentage from normal min.
    If FeasiblePercent Then
        FeasibleMinValue = NormalMin - ((FeasibleMin / 100) * NormalMin)
    Else
        FeasibleMinValue = FeasibleMin
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get FeasibleMaxValue()
Attribute FeasibleMaxValue.VB_Description = "The absolute value for the feasible max."
'----------------------------------------------------------------------------------------'
' The absolute value for the feasible max.
'----------------------------------------------------------------------------------------'
'TA 22/11/2000: feasible max added to the percentage
    If FeasiblePercent Then
        FeasibleMaxValue = NormalMax + (FeasibleMax / 100) * NormalMax
    Else
        FeasibleMaxValue = FeasibleMax
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get FeasiblePercent() As Boolean
Attribute FeasiblePercent.VB_Description = "Is feasible range based on a percentage?"
'----------------------------------------------------------------------------------------'
'Is feasible range based on a percentage?
'----------------------------------------------------------------------------------------'

    FeasiblePercent = PercentCode And 1

End Property

'----------------------------------------------------------------------------------------'
Public Property Get AbsolutePercent() As Boolean
Attribute AbsolutePercent.VB_Description = "Is absolute range based on a percentage?"
'----------------------------------------------------------------------------------------'
'Is absolute range based on a percentage?
'----------------------------------------------------------------------------------------'

    AbsolutePercent = PercentCode And 2

End Property

'----------------------------------------------------------------------------------------'
Public Property Get HasFeasibleRange() As Boolean
Attribute HasFeasibleRange.VB_Description = "True if a feasible range has been defined."
'----------------------------------------------------------------------------------------'
'True if a feasible range has been defined.
'----------------------------------------------------------------------------------------'
    
    HasFeasibleRange = Not (VarType(FeasibleMin) = vbNull And VarType(FeasibleMax) = vbNull)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get HasAbsoluteRange() As Boolean
Attribute HasAbsoluteRange.VB_Description = "True if a absolute range has been defined."
'----------------------------------------------------------------------------------------'
' True if a absolute range has been defined.
'----------------------------------------------------------------------------------------'
    
    HasAbsoluteRange = Not (VarType(AbsoluteMin) = vbNull And VarType(Absolutemax) = vbNull)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get NormalRangeText() As String
Attribute NormalRangeText.VB_Description = "The calculated Normal range as a text expression with units."
'----------------------------------------------------------------------------------------'
' The calculated Normal range as a text expression with units.
'----------------------------------------------------------------------------------------'
Dim sMinText As String
Dim sMaxText As String

    If VarType(NormalMin) <> vbNull Then
        sMinText = NormalMin
    End If
    
    If VarType(NormalMax) <> vbNull Then
        sMaxText = NormalMax
    End If
    
    NormalRangeText = RangeText(sMinText, sMaxText)
        
End Property

'----------------------------------------------------------------------------------------'
Public Property Get FeasibleRangeText() As String
Attribute FeasibleRangeText.VB_Description = "The calculated feasible range as a text expression with units."
'----------------------------------------------------------------------------------------'
' The calculated feasible range as a text expression with units.
'----------------------------------------------------------------------------------------'
Dim sMinText As String
Dim sMaxText As String

    If VarType(FeasibleMin) <> vbNull Then
        sMinText = FeasibleMinValue
    End If
    
    If VarType(FeasibleMax) <> vbNull Then
        sMaxText = FeasibleMaxValue
    End If
    
    FeasibleRangeText = RangeText(sMinText, sMaxText)

        
End Property

'----------------------------------------------------------------------------------------'
Public Property Get AbsoluteRangeText() As String
Attribute AbsoluteRangeText.VB_Description = "The calculated Absolute range as a text expression with units."
'----------------------------------------------------------------------------------------'
' The calculated Absolute range as a text expression with units.
'----------------------------------------------------------------------------------------'
Dim sMinText As String
Dim sMaxText As String

    If VarType(AbsoluteMin) <> vbNull Then
        sMinText = AbsoluteMinValue
    End If
    
    If VarType(Absolutemax) <> vbNull Then
        sMaxText = AbsoluteMaxValue
    End If
    
    AbsoluteRangeText = RangeText(sMinText, sMaxText)
    
End Property

Private Function RangeText(sMinText As String, sMaxText As String) As String
'----------------------------------------------------------------------------------------'
' The calculated range as a text expression with units.
'----------------------------------------------------------------------------------------'
    
    If sMinText = "" And sMaxText = "" Then
        'neither has a value
        RangeText = ""
    Else
        If sMinText <> "" And sMaxText <> "" Then
            'both have a value
            RangeText = sMinText & " to " & sMaxText & " " '& moClinicalTest.Unit
        Else
            If sMinText <> "" Then
                'min has value, max hasn't
                RangeText = "Less than " & sMinText & " " '& moClinicalTest.Unit
            Else
                'max has value, min hasn't
                'TA 11/01/01: ensure sMaxText displayed
                RangeText = "Greater than " & sMaxText & " " '& moClinicalTest.Unit
            End If
        End If
    End If


End Function


'all properties
'----------------------------------------------------------------------------------------'
Public Property Get Id() As Long
Attribute Id.VB_Description = "Normal Range ID."
'---------------------------------------------------------
' Normal Range ID.
'---------------------------------------------------------

    Id = moState.Id
    
End Property

'---------------------------------------------------------
Public Property Get LabCode() As String
Attribute LabCode.VB_Description = "Normal Range's lab code."
'---------------------------------------------------------
'Normal Range's lab code.
'---------------------------------------------------------

    LabCode = moState.LabCode
    
End Property

'---------------------------------------------------------
Public Property Get AgeMin() As Variant
Attribute AgeMin.VB_Description = "Range's min age."
'---------------------------------------------------------
'Range's min age.
'---------------------------------------------------------

    AgeMin = moState.AgeMin
    
End Property

'---------------------------------------------------------
Public Property Get AgeMax() As Variant
Attribute AgeMax.VB_Description = "Range's max age."
'---------------------------------------------------------
'Range's max age.
'---------------------------------------------------------

    AgeMax = moState.AgeMax
    
End Property

'---------------------------------------------------------
Public Property Get EffectiveStart() As Double
Attribute EffectiveStart.VB_Description = "Range's effective start date."
'---------------------------------------------------------
'Range's effective start date.
'---------------------------------------------------------

    EffectiveStart = moState.EffectiveStart
    
End Property

'---------------------------------------------------------
Public Property Get EffectiveEnd() As Double
Attribute EffectiveEnd.VB_Description = "Range's effective date end."
'---------------------------------------------------------
'Range's effective date end.
'---------------------------------------------------------

    EffectiveEnd = moState.EffectiveEnd
    
End Property

'---------------------------------------------------------
Public Property Get NormalMin() As Variant
Attribute NormalMin.VB_Description = "Range's normal min."
'---------------------------------------------------------
'Range's normal min.
'---------------------------------------------------------
    NormalMin = moState.NormalMin
    
End Property

'---------------------------------------------------------
Public Property Get NormalMax() As Variant
Attribute NormalMax.VB_Description = "Range's normal max."
'---------------------------------------------------------
'Range's normal max.
'---------------------------------------------------------

    NormalMax = moState.NormalMax
    
End Property

'---------------------------------------------------------
Public Property Get FeasibleMin() As Variant
Attribute FeasibleMin.VB_Description = "Range's feasible min."
'---------------------------------------------------------
'Range's feasible min.
'---------------------------------------------------------

    FeasibleMin = moState.FeasibleMin
    
End Property

'---------------------------------------------------------
Public Property Get FeasibleMax() As Variant
Attribute FeasibleMax.VB_Description = "Range's feasible max."
'---------------------------------------------------------
'Range's feasible max.
'---------------------------------------------------------
    FeasibleMax = moState.FeasibleMax
    
End Property

'---------------------------------------------------------
Public Property Get AbsoluteMin() As Variant
Attribute AbsoluteMin.VB_Description = "Range's absolute min."
'---------------------------------------------------------
'Range's absolute min.
'---------------------------------------------------------

    AbsoluteMin = moState.AbsoluteMin
    
End Property

'---------------------------------------------------------
Public Property Get Absolutemax() As Variant
Attribute Absolutemax.VB_Description = "Range's absolute max."
'---------------------------------------------------------
'Range's absolute max.
'---------------------------------------------------------

    Absolutemax = moState.Absolutemax
    
End Property

'---------------------------------------------------------
Public Property Get ClinicalTestCode() As String
Attribute ClinicalTestCode.VB_Description = "Range's clinical test code."
'---------------------------------------------------------
'Range's clinical test code.
'---------------------------------------------------------

    ClinicalTestCode = moState.ClinicalTestCode
    
End Property
 
'---------------------------------------------------------
Public Property Get GenderCode() As Integer
Attribute GenderCode.VB_Description = "Range's gender code. 0 for unspecified 1 for female. 2 for male."
'---------------------------------------------------------
'Range's gender code.
' 0 for unspecified
' 1 for female.
' 2 for male.
'---------------------------------------------------------

    GenderCode = moState.GenderCode
    
End Property

'---------------------------------------------------------
Public Property Get PercentCode() As Integer
Attribute PercentCode.VB_Description = "Ranges percent code. 0 means feasible and absolute are NOT based on the normal. 1 means feasible is beased on the normal. 2 means absolute is based on the normal. 3 means feasible and absolute are based on the normal."
'---------------------------------------------------------
'Ranges percent code.
' 0 means feasible and absolute are NOT based on the normal.
' 1 means feasible is beased on the normal.
' 2 means absolute is based on the normal.
' 3 means feasible and absolute are based on the normal.
'---------------------------------------------------------

    PercentCode = moState.PercentCode
    
End Property
 
'---------------------------------------------------------
Private Sub Class_Terminate()
'---------------------------------------------------------
' NCJ 24 Mar 04
'---------------------------------------------------------

    Set moState = Nothing

End Sub
