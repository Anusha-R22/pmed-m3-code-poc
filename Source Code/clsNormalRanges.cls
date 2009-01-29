VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNormalRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsNormalRanges.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, September 2000
'   Purpose:    Collection of Normal Ranges for MACRO
'----------------------------------------------------------------------------------------'
' Revisions:
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
'   TA 09/11/2000: updating or inserting a normal range will now set the lab's changed property to changed
'----------------------------------------------------------------------------------------'


Option Explicit

'internal collection of normal ranges
Private mcolNormalRanges As Collection

'Lab Code for all these normal ranges
Private msLabCode As String

'----------------------------------------------------------------------------------------'
Public Sub Load(sLabCode As String)
'----------------------------------------------------------------------------------------'
' load the collection of normal ranges for a particular lab
'----------------------------------------------------------------------------------------'

Dim sSQL As String
Dim rsNR As ADODB.Recordset
Dim oNormalRange As clsNormalRange

    Call HourglassOn
    
    msLabCode = sLabCode
        
    Set mcolNormalRanges = New Collection
       
    sSQL = "SELECT * FROM NormalRange WHERE LaboratoryCode = '" & sLabCode & "'"
   
    Set rsNR = New ADODB.Recordset
    rsNR.Open sSQL, MacroADODBConnection

    Do While Not rsNR.EOF
        Set oNormalRange = New clsNormalRange
        With oNormalRange
            .Id = rsNR.Fields!NormalRangeId
            .LabCode = sLabCode
            .ClinicalTestCode = rsNR.Fields!ClinicalTestCode
            .GenderCode = rsNR.Fields!NormalRangeGender
            .AgeMin = rsNR.Fields!NormalRangeAgeMin
            .AgeMax = rsNR.Fields!NormalRangeAgeMax
            .EffectiveStart = rsNR.Fields!NormalRangeEffectiveStart
            .EffectiveEnd = rsNR.Fields!NormalRangeEffectiveEnd
            .NormalMin = rsNR.Fields!NormalRangeNormalMin
            .NormalMax = rsNR.Fields!NormalRangeNormalMax
            .FeasibleMin = rsNR.Fields!NormalRangeFeasibleMin
            .FeasibleMax = rsNR.Fields!NormalRangeFeasibleMAx
            .AbsoluteMin = rsNR.Fields!NormalRangeAbsoluteMin
            .Absolutemax = rsNR.Fields!NormalRangeAbsoluteMax
            .PercentCode = rsNR.Fields!NormalRangePercent
            'add to collection
            mcolNormalRanges.Add oNormalRange, GetKey(.Id)
        End With

        rsNR.MoveNext
    Loop
    
    rsNR.Close
    Set rsNR = Nothing
    
    Call HourglassOff
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function Item(lId As Long) As clsNormalRange
'----------------------------------------------------------------------------------------'
' return an item by its normal range id
'----------------------------------------------------------------------------------------'

    Set Item = mcolNormalRanges(GetKey(lId))

End Function

'----------------------------------------------------------------------------------------'
Private Function GetKey(lId As Long) As String
'----------------------------------------------------------------------------------------'
' return an the key for an item by its range id
'----------------------------------------------------------------------------------------'

    GetKey = "K" & Format(lId)

End Function

'----------------------------------------------------------------------------------------'
Public Function PopulateListView(lvwlistview As MSComctlLib.ListView) As Long
'----------------------------------------------------------------------------------------'
' populate a listview with the normal ranges collection
' returns the number of rows in the listview
'----------------------------------------------------------------------------------------'

Dim itmListItem As MSComctlLib.ListItem
Dim oNormalRange As clsNormalRange
Dim sHeadings As String
Dim vHeadings As Variant
Dim sValue As Variant
Dim i As Long
Dim bVisible As Boolean

    bVisible = lvwlistview.Visible
    lvwlistview.Visible = False
    
    lvwlistview.ListItems.Clear
    lvwlistview.ColumnHeaders.Clear

    'do column headers
    For Each sValue In Headings
        lvwlistview.ColumnHeaders.Add , , sValue, lvwlistview.Parent.TextWidth(sValue) + 12 * Screen.TwipsPerPixelX
    Next


     For Each oNormalRange In mcolNormalRanges
        With oNormalRange
            sValue = oNormalRange.ClinicalTest.ClinicalTestGroup.Code
            Set itmListItem = lvwlistview.ListItems.Add(, , sValue)
            'use the is as the tag
            itmListItem.Tag = Format(.Id)
            'check width
            If lvwlistview.ColumnHeaders(1).Width < (lvwlistview.Parent.TextWidth(sValue) + 6 * Screen.TwipsPerPixelX) Then
                lvwlistview.ColumnHeaders(1).Width = (lvwlistview.Parent.TextWidth(sValue) + 6 * Screen.TwipsPerPixelX)
            End If
            i = 1
            For Each sValue In .ListViewSubItems
                itmListItem.SubItems(i) = sValue
                'check width
                If lvwlistview.ColumnHeaders(i + 1).Width < (lvwlistview.Parent.TextWidth(sValue) + 12 * Screen.TwipsPerPixelX) Then
                  lvwlistview.ColumnHeaders(i + 1).Width = (lvwlistview.Parent.TextWidth(sValue) + 12 * Screen.TwipsPerPixelX)
                End If
                i = i + 1
            Next
        End With
     Next

    lvwlistview.Visible = bVisible
    
    PopulateListView = mcolNormalRanges.Count

End Function

'----------------------------------------------------------------------------------------'
Private Property Get Headings() As Collection
'----------------------------------------------------------------------------------------'
' return a collection of headings to be used in a listview
'----------------------------------------------------------------------------------------'
    
    Set Headings = New Collection
    With Headings
        .Add "Test Group"
        .Add "Test"
        .Add "Gender"
        .Add "Age Range"
        .Add "Normal Range"
        .Add "Feasible Range"
        .Add "Absolute Range"
        .Add "Effective Dates"
    End With
    
End Property

'----------------------------------------------------------------------------------------'
Public Sub Delete(lId As Long)
'----------------------------------------------------------------------------------------'
' delete a normal range from the collection and the database
'----------------------------------------------------------------------------------------'
    MacroADODBConnection.Execute "DELETE FROM NormalRange WHERE LaboratoryCode = '" & msLabCode & "'" _
                                    & " AND NormalRangeId = " & lId
    mcolNormalRanges.Remove GetKey(lId)
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function Update(ByVal lId As Long, sClinicalTestCode As String, nGender As GenderCode, _
                        sAgeMin As String, sAgeMax As String, _
                        dblEffectiveStart As Double, dblEffectiveEnd As Double, _
                        sNormalMin As String, sNormalMax As String, sFeasibleMin As String, sFeasibleMax As String, _
                        sAbsoluteMin As String, sAbsoluteMax As String, _
                        bFeasiblePercent As Boolean, bAbsolutePercent As Boolean) As Long
'----------------------------------------------------------------------------------------'
' add a single NormalRange to the collection and the database
' function returns ValidRange Status or a normal range id if a conflicting one is found
'----------------------------------------------------------------------------------------'
Dim oNormalRange As clsNormalRange
Dim lUpdate As Long
    
    'create new range
    Set oNormalRange = New clsNormalRange

    With oNormalRange
        .Id = lId
        .ClinicalTestCode = sClinicalTestCode
        .LabCode = msLabCode
        .GenderCode = nGender
        .AgeMin = StringtoNumberVariant(sAgeMin)
        .AgeMax = StringtoNumberVariant(sAgeMax)
        .EffectiveStart = dblEffectiveStart
        .EffectiveEnd = dblEffectiveEnd
        .NormalMin = StringtoNumberVariant(sNormalMin)
        .NormalMax = StringtoNumberVariant(sNormalMax)
        .FeasibleMin = StringtoNumberVariant(sFeasibleMin)
        .FeasibleMax = StringtoNumberVariant(sFeasibleMax)
        .AbsoluteMin = StringtoNumberVariant(sAbsoluteMin)
        .Absolutemax = StringtoNumberVariant(sAbsoluteMax)
        .FeasiblePercent = bFeasiblePercent
        .AbsolutePercent = bAbsolutePercent

        'get valid range status
        lUpdate = ValidRange(oNormalRange)
        
        'TA 01/03/2001: change local numbers to standard format
        If lUpdate = ValidRangeStatus.vreOK Then
            'valid range - save to database and update collection
            MacroADODBConnection.Execute "UPDATE NormalRange SET ClinicalTestCode='" & sClinicalTestCode & "'" _
                                        & ", NormalRangeGender=" & .GenderCode _
                                        & ", NormalRangeAgeMin=" & ConvertLocalNumToStandard(VarianttoString(.AgeMin, True)) _
                                        & ", NormalRangeAgeMax=" & ConvertLocalNumToStandard(VarianttoString(.AgeMax, True)) _
                                        & ", NormalRangeEffectiveStart=" & ConvertLocalNumToStandard(CStr(.EffectiveStart)) _
                                        & ", NormalRangeEffectiveEnd=" & ConvertLocalNumToStandard(CStr(.EffectiveEnd)) _
                                        & ", NormalRangeNormalMin=" & ConvertLocalNumToStandard(VarianttoString(.NormalMin, True)) _
                                        & ", NormalRangeNormalMax=" & ConvertLocalNumToStandard(VarianttoString(.NormalMax, True)) _
                                        & ", NormalRangeFeasibleMin=" & ConvertLocalNumToStandard(VarianttoString(.FeasibleMin, True)) _
                                        & ", NormalRangeFeasibleMax=" & ConvertLocalNumToStandard(VarianttoString(.FeasibleMax, True)) _
                                        & ", NormalRangeAbsoluteMin=" & ConvertLocalNumToStandard(VarianttoString(.AbsoluteMin, True)) _
                                        & ", NormalRangeAbsoluteMax=" & ConvertLocalNumToStandard(VarianttoString(.Absolutemax, True)) _
                                        & ", NormalRangePercent=" & .PercentCode _
                                        & " WHERE NormalRangeId=" & lId _
                                        & " AND LaboratoryCode = '" & msLabCode & "'"
            'update lab as changed
            MacroADODBConnection.Execute "UPDATE Laboratory SET Changed = " & Changed.Changed _
                                            & " WHERE LaboratoryCode = '" & msLabCode & "'"
                                            
            'copy our temporary normal range object into the collection
           .CopyInto mcolNormalRanges.Item(GetKey(lId))
        End If
        
    End With
    
    Update = lUpdate

End Function


'----------------------------------------------------------------------------------------'
Public Function Insert(sClinicalTestCode As String, nGender As GenderCode, _
                        sAgeMin As String, sAgeMax As String, _
                        dblEffectiveStart As Double, dblEffectiveEnd As Double, _
                        sNormalMin As String, sNormalMax As String, sFeasibleMin As String, sFeasibleMax As String, _
                        sAbsoluteMin As String, sAbsoluteMax As String, _
                        bFeasiblePercent As Boolean, bAbsolutePercent As Boolean, lNewId As Long) As Long
'----------------------------------------------------------------------------------------'
' add a single NormalRange to the collection and the database
' function returns ValidRange Status or a normal range id if a conflicting one is found
' lNewId passed in as a variable to receive new id
'----------------------------------------------------------------------------------------'
Dim oNormalRange As clsNormalRange
Dim lInsert As Long

    lNewId = GetNewId
    'create new NormalRange object
    Set oNormalRange = New clsNormalRange
    With oNormalRange
        .Id = lNewId
        .LabCode = msLabCode
        .ClinicalTestCode = sClinicalTestCode
        .GenderCode = nGender
        .AgeMin = StringtoNumberVariant(sAgeMin)
        .AgeMax = StringtoNumberVariant(sAgeMax)
        .EffectiveStart = dblEffectiveStart
        .EffectiveEnd = dblEffectiveEnd
        .NormalMin = StringtoNumberVariant(sNormalMin)
        .NormalMax = StringtoNumberVariant(sNormalMax)
        .FeasibleMin = StringtoNumberVariant(sFeasibleMin)
        .FeasibleMax = StringtoNumberVariant(sFeasibleMax)
        .AbsoluteMin = StringtoNumberVariant(sAbsoluteMin)
        .Absolutemax = StringtoNumberVariant(sAbsoluteMax)
        .FeasiblePercent = bFeasiblePercent
        .AbsolutePercent = bAbsolutePercent
        
        'TA 01/03/2001: change local numbers to standard format
        'get valid range status
         lInsert = ValidRange(oNormalRange)
         If lInsert = ValidRangeStatus.vreOK Then
            'valid range - save to database and add to collection
             MacroADODBConnection.Execute "INSERT INTO NormalRange values (" _
                                         & lNewId _
                                         & ", '" & msLabCode & "'" _
                                         & ", '" & sClinicalTestCode & "'" _
                                         & ", " & .GenderCode _
                                         & ", " & ConvertLocalNumToStandard(VarianttoString(.AgeMin, True)) _
                                         & ", " & ConvertLocalNumToStandard(VarianttoString(.AgeMax, True)) _
                                         & ", " & ConvertLocalNumToStandard(CStr(.EffectiveStart)) _
                                         & ", " & ConvertLocalNumToStandard(CStr(.EffectiveEnd)) _
                                         & ", " & ConvertLocalNumToStandard(VarianttoString(.NormalMin, True)) _
                                         & ", " & ConvertLocalNumToStandard(VarianttoString(.NormalMax, True)) _
                                         & ", " & ConvertLocalNumToStandard(VarianttoString(.FeasibleMin, True)) _
                                         & ", " & ConvertLocalNumToStandard(VarianttoString(.FeasibleMax, True)) _
                                         & ", " & ConvertLocalNumToStandard(VarianttoString(.AbsoluteMin, True)) _
                                         & ", " & ConvertLocalNumToStandard(VarianttoString(.Absolutemax, True)) _
                                         & ", " & .PercentCode & ")"
            'update lab as changed
            MacroADODBConnection.Execute "UPDATE Laboratory SET Changed = " & Changed.Changed _
                                            & " WHERE LaboratoryCode = '" & msLabCode & "'"
                                            
            'add to collection
            mcolNormalRanges.Add oNormalRange, GetKey(.Id)
        End If



    End With
    
    Insert = lInsert


End Function

'----------------------------------------------------------------------------------------'
Private Function ValidRange(oNormalRange As clsNormalRange) As Long
'----------------------------------------------------------------------------------------'
' compares a normal range with the others in the collection
' returns ValidRangeStatus.vreOK if valid range, else the conflicting id
'TA 22/11/2000: zeroto null put round dates as range matching assumes zero to mean zero
'----------------------------------------------------------------------------------------'
Dim oNRCheck As clsNormalRange
Dim lConflictingId As Long

    lConflictingId = ValidRangeStatus.vreOK
    
    With oNormalRange

        'TA have to change sign of result 'cos erros must be returned as negative
        lConflictingId = -.Valid
        
        If lConflictingId = ValidRangeStatus.vreOK Then
            For Each oNRCheck In mcolNormalRanges
                If (.LabCode = oNRCheck.LabCode) _
                    And (.ClinicalTestCode = oNRCheck.ClinicalTestCode) _
                    And (RangeOverlap(.AgeMin, .AgeMax, oNRCheck.AgeMin, oNRCheck.AgeMax)) _
                    And (RangeOverlap(ZeroToNull(.EffectiveStart), ZeroToNull(.EffectiveEnd), ZeroToNull(oNRCheck.EffectiveStart), ZeroToNull(oNRCheck.EffectiveEnd))) Then
                        'overlaps on test, age and effective date
                        If ((.GenderCode = oNRCheck.GenderCode) Or (.GenderCode = gNone) Or (oNRCheck.GenderCode = gNone)) Then
                            'overlaps on gender
                            If .Id <> oNRCheck.Id Then
                                'not the same id - this conflicts
                                lConflictingId = oNRCheck.Id
                            End If
                        End If
                End If
            Next
        End If
    End With
    
    ValidRange = lConflictingId
    
End Function

'----------------------------------------------------------------------------------------'
Public Function NewEnum() As IUnknown
'----------------------------------------------------------------------------------------'

    Set NewEnum = mcolNormalRanges.[_NewEnum]

End Function

'----------------------------------------------------------------------------------------'
Public Function GetNewId() As Long
'----------------------------------------------------------------------------------------'
' returns the next id in sequence
'----------------------------------------------------------------------------------------'

Dim rsTemp As ADODB.Recordset
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "SELECT Max(NormalRangeId) FROM NormalRange WHERE LaboratoryCode = '" & msLabCode & "'", MacroADODBConnection
    If IsNull(rsTemp.Fields(0).Value) Then
        'first ever normal range
        GetNewId = 1
    Else
        GetNewId = rsTemp.Fields(0).Value + 1
    End If
    
End Function

Private Sub Class_Initialize()
    msLabCode = "0"
End Sub
