VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNormalRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsNormalRange.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, September 2000
'   Purpose:    Normal Range object for MACRO
'----------------------------------------------------------------------------------------'
' Revisions:
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
'   TA 17/10/2000: Validation for a range altered so that normal must be inside feasible
'   TA 18/10/2000: New read only properties -  HasFeasibleRange? and HasAbsoluteRange?
'   TA 15/12/2000: Feasible and Absoulte as percentages of normal corrected
'----------------------------------------------------------------------------------------'

Option Explicit

Const m_NO_VALUE = 0

'all properties
Public Id As Long
Public LabCode As String
Public AgeMin As Variant
Public AgeMax As Variant
Public EffectiveStart As Double
Public EffectiveEnd As Double
Public NormalMin As Variant
Public NormalMax As Variant
Public FeasibleMin As Variant
Public FeasibleMax As Variant
Public AbsoluteMin As Variant
Public Absolutemax As Variant
Public GenderCode As GenderCode
Public PercentCode As Integer

Private msClinicalTestCode As String
Private moClinicalTest As clsClinicalTest


'functions to retrieve corresponding tests

'----------------------------------------------------------------------------------------'
Public Property Get ClinicalTest() As clsClinicalTest
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Set ClinicalTest = moClinicalTest

End Property

'----------------------------------------------------------------------------------------'
Public Property Let ClinicalTestCode(sClinicalTestCode As String)
'----------------------------------------------------------------------------------------'
' set up linked Clinical Test
'----------------------------------------------------------------------------------------'

    msClinicalTestCode = sClinicalTestCode
    Set moClinicalTest = New clsClinicalTest
    moClinicalTest.Load msClinicalTestCode

End Property

'----------------------------------------------------------------------------------------'
Public Property Get ClinicalTestCode() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
    
    ClinicalTestCode = msClinicalTestCode

End Property


'functions to return text values of properties

'----------------------------------------------------------------------------------------'
Public Property Get AgeMinText()
'----------------------------------------------------------------------------------------'
    AgeMinText = VarianttoString(AgeMin)
End Property

'----------------------------------------------------------------------------------------'
Public Property Get AgeMaxText()
'----------------------------------------------------------------------------------------'
    
    AgeMaxText = VarianttoString(AgeMax)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get EffectiveStartText()
'----------------------------------------------------------------------------------------'
    
    If EffectiveStart = 0 Then
        EffectiveStartText = ""
    Else
        EffectiveStartText = Format(CDate(EffectiveStart), DEFAULT_DATE_FORMAT)
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get EffectiveEndText()
'----------------------------------------------------------------------------------------'
    
    If EffectiveEnd = 0 Then
        EffectiveEndText = ""
    Else
        EffectiveEndText = Format(CDate(EffectiveEnd), DEFAULT_DATE_FORMAT)
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get NormalMinText()
'----------------------------------------------------------------------------------------'
    
    NormalMinText = VarianttoString(NormalMin)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get NormalMaxText()
'----------------------------------------------------------------------------------------'
    
    NormalMaxText = VarianttoString(NormalMax)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get FeasibleMinText()
'----------------------------------------------------------------------------------------'
    
    FeasibleMinText = VarianttoString(FeasibleMin)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get FeasibleMaxText()
'----------------------------------------------------------------------------------------'
    
    FeasibleMaxText = VarianttoString(FeasibleMax)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get AbsoluteMinText()
'----------------------------------------------------------------------------------------'
    
    AbsoluteMinText = VarianttoString(AbsoluteMin)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get AbsolutemaxText()
'----------------------------------------------------------------------------------------'

    AbsolutemaxText = VarianttoString(Absolutemax)

End Property

'----------------------------------------------------------------------------------------'
Public Property Get FeasiblePercent() As Boolean
'----------------------------------------------------------------------------------------'

    FeasiblePercent = PercentCode And 1

End Property

'----------------------------------------------------------------------------------------'
Public Property Get AbsolutePercent() As Boolean
'----------------------------------------------------------------------------------------'

    AbsolutePercent = PercentCode And 2

End Property

'----------------------------------------------------------------------------------------'
Public Property Get FeasibleMinValue()
'----------------------------------------------------------------------------------------'
' return the absolute value for the feasible min
'----------------------------------------------------------------------------------------'
'TA 15/12/2000: subtract percentage from normal min
    If FeasiblePercent Then
        FeasibleMinValue = NormalMin - ((FeasibleMin / 100) * NormalMin)
    Else
        FeasibleMinValue = FeasibleMin
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get FeasibleMaxValue()
'----------------------------------------------------------------------------------------'
' return the absolute value for the feasible max
'----------------------------------------------------------------------------------------'
'TA 22/11/2000: feasible max added to the percentage
    If FeasiblePercent Then
        FeasibleMaxValue = NormalMax + (FeasibleMax / 100) * NormalMax
    Else
        FeasibleMaxValue = FeasibleMax
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get AbsoluteMinValue()
'----------------------------------------------------------------------------------------'
' return the absolute (calculated) value for the absolute min
'----------------------------------------------------------------------------------------'

    If AbsolutePercent Then
        AbsoluteMinValue = NormalMin - ((AbsoluteMin / 100) * NormalMin)
    Else
        AbsoluteMinValue = AbsoluteMin
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get AbsoluteMaxValue()
'----------------------------------------------------------------------------------------'
' return the absolute (calculated) value for the absolute max
'----------------------------------------------------------------------------------------'
'TA 22/11/2000: absolute max added to the percentage
    If AbsolutePercent Then
        AbsoluteMaxValue = NormalMax + (Absolutemax / 100) * NormalMax
    Else
        AbsoluteMaxValue = Absolutemax
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let FeasiblePercent(bPercent As Boolean)
'----------------------------------------------------------------------------------------'
' is feasible expressed as percentage?
'----------------------------------------------------------------------------------------'

    If bPercent Then
        PercentCode = PercentCode Or 1
    Else
        PercentCode = PercentCode And 2
    End If

End Property

'----------------------------------------------------------------------------------------'
Public Property Let AbsolutePercent(bPercent As Boolean)
'----------------------------------------------------------------------------------------'
' is absolute expressed as percentage?
'----------------------------------------------------------------------------------------'

    If bPercent Then
        PercentCode = PercentCode Or 2
    Else
        PercentCode = PercentCode And 1
    End If

End Property

'----------------------------------------------------------------------------------------'
Public Property Get NormalRangeText() As String
'----------------------------------------------------------------------------------------'
' reutrn the calcualted Normal range as a text expression with units
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
'----------------------------------------------------------------------------------------'
' reutrn the calcualted feasible range as a text expression with units
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
'----------------------------------------------------------------------------------------'
' reutrn the calcualted Absolute range as a text expression with units
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
' reutrn the calcualted range as a text expression with units
'----------------------------------------------------------------------------------------'
    
    If sMinText = "" And sMaxText = "" Then
        'neither has a value
        RangeText = ""
    Else
        If sMinText <> "" And sMaxText <> "" Then
            'both have a value
            RangeText = sMinText & " to " & sMaxText & " " & moClinicalTest.Unit
        Else
            If sMinText <> "" Then
                'min has value, max hasn't
                RangeText = "Less than " & sMinText & " " & moClinicalTest.Unit
            Else
                'max has value, min hasn't
                'TA 11/01/01: ensure sMaxText displayed
                RangeText = "Greater than " & sMaxText & " " & moClinicalTest.Unit
            End If
        End If
    End If


End Function

'----------------------------------------------------------------------------------------'
Public Property Get HasFeasibleRange() As Boolean
'----------------------------------------------------------------------------------------'
'reutrns true if a feasible range has been defined
'----------------------------------------------------------------------------------------'
    
    HasFeasibleRange = Not (VarType(FeasibleMin) = vbNull And VarType(FeasibleMax) = vbNull)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get HasAbsoluteRange() As Boolean
'----------------------------------------------------------------------------------------'
'reutrns true if a absolute range has been defined
'----------------------------------------------------------------------------------------'
    
    HasAbsoluteRange = Not (VarType(AbsoluteMin) = vbNull And VarType(Absolutemax) = vbNull)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get ListViewSubItems() As Collection
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Set ListViewSubItems = New Collection
    With ListViewSubItems
        .Add moClinicalTest.Code
        .Add GetGenderText(GenderCode)
        .Add RangeExpr(AgeMinText, AgeMaxText)
        .Add RangeExpr(NormalMinText, NormalMaxText)
        .Add RangeExpr(FeasibleMinText & Switch(FeasiblePercent, "%", Not (FeasiblePercent), ""), _
                        FeasibleMaxText & Switch(FeasiblePercent, "%", Not (FeasiblePercent), ""))
        .Add RangeExpr(AbsoluteMinText & Switch(AbsolutePercent, "%", Not (AbsolutePercent), ""), _
                        AbsolutemaxText & Switch(AbsolutePercent, "%", Not (AbsolutePercent), ""))
        .Add RangeExpr(EffectiveStartText, EffectiveEndText)
    End With
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get Valid() As Long
'----------------------------------------------------------------------------------------'
' validate as range
' returns a Valid Range Status
' TA 27/09/2000: currently should only ever be vreFeasibleinNormal,
'                    vreAbsoluteinNormal or AbsoluteinFeasible as
'                    otheres are checked for in form
'----------------------------------------------------------------------------------------'

Dim lValid As Long
      
    lValid = ValidRangeStatus.vreOK
    
    ' there will always be a valid normal range when execution reaches here
    
    If HasFeasibleRange And Not (RangeInRange(NormalMin, NormalMax, FeasibleMin, FeasibleMax) Or FeasiblePercent) Then
        'normal not inside feasible AND feasible not a percentage
        lValid = lValid Or ValidRangeStatus.vreFeasibleinNormal
    End If
    
    If HasAbsoluteRange And Not (RangeInRange(NormalMin, NormalMax, AbsoluteMin, Absolutemax) Or AbsolutePercent) Then
        'normal not inside absolute AND absolute not a percentage
        lValid = lValid Or ValidRangeStatus.vreAbsoluteinNormal
    End If
    
    If HasAbsoluteRange And Not (RangeInRange(FeasibleMin, FeasibleMax, AbsoluteMin, Absolutemax) Or (FeasiblePercent And AbsolutePercent)) Then
        'feasible not inside absolute AND feasible and absolute not a percentage
        lValid = lValid Or ValidRangeStatus.vreAbsoluteinFeasible
    End If

    Valid = lValid
    
End Property

'----------------------------------------------------------------------------------------'
Public Sub CopyInto(oNewNormalRange As clsNormalRange)
'----------------------------------------------------------------------------------------'
'copy all this objects data into oNewNormalRange's data
'----------------------------------------------------------------------------------------'

    With oNewNormalRange
        .Id = Id
        .LabCode = LabCode
        .ClinicalTestCode = ClinicalTestCode
        .AgeMin = AgeMin
        .AgeMax = AgeMax
        .EffectiveStart = EffectiveStart
        .EffectiveEnd = EffectiveEnd
        .NormalMin = NormalMin
        .NormalMax = NormalMax
        .FeasibleMin = FeasibleMin
        .FeasibleMax = FeasibleMax
        .AbsoluteMin = AbsoluteMin
        .Absolutemax = Absolutemax
        .GenderCode = GenderCode
        .PercentCode = PercentCode
    End With

End Sub
