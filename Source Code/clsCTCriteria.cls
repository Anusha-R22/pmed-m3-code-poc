VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCTCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsCTCrieria.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, September 2000
'   Purpose:    Collection of CTC for MACRO
'----------------------------------------------------------------------------------------'
' Revisions:
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
'   TA 06/10/2000: Now checks for duplicate criteria
'----------------------------------------------------------------------------------------'


Option Explicit

Private Const m_ERROR_DUPLICATECTC = "A CTC with an identical range already exists for this test"

'internal collection of normal ranges
Private mColCTC As Collection

'Scheme Code for all these CTC
Private msSchemeCode As String

'----------------------------------------------------------------------------------------'
Public Sub Load(sSchemeCode As String)
'----------------------------------------------------------------------------------------'
' load the collection of normal ranges for a particular lab
'----------------------------------------------------------------------------------------'

Dim sSQL As String
Dim rsCTC As ADODB.Recordset
Dim oCTCriterion As clsCTCriterion

    Call HourglassOn
    
    msSchemeCode = sSchemeCode
        
    Set mColCTC = New Collection
       
       
'TA 03/04/2003: Order by group,test,grade
    sSQL = "select CTC.* from CTC, CLINICALTEST"
    sSQL = sSQL & " where CTC.CLINICALTESTCODE = CLINICALTEST.CLINICALTESTCODE"
    sSQL = sSQL & " and CTCSCHEMECODE = '" & sSchemeCode & "'"
    sSQL = sSQL & " order by CLINICALTEST.CLINICALTESTGROUPCODE,CTC.CLINICALTESTCODE,CTC.CTCGRADE"
   
    Set rsCTC = New ADODB.Recordset
    rsCTC.Open sSQL, MacroADODBConnection

    Do While Not rsCTC.EOF
        Set oCTCriterion = New clsCTCriterion
        With oCTCriterion
            .Id = rsCTC.Fields!CTCId
            .SchemeCode = rsCTC.Fields!CTCSchemeCode
            .ClinicalTestCode = rsCTC.Fields!ClinicalTestCode
            .Grade = rsCTC.Fields!CTCGrade
            .Min = rsCTC.Fields!CTCMin
            .Max = rsCTC.Fields!CTCMax
            .MinType = rsCTC.Fields!CTCMinType
            .MaxType = rsCTC.Fields!CTCMaxType

            'add to collection
            mColCTC.Add oCTCriterion, GetKey(.Id)
        End With

        rsCTC.MoveNext
    Loop
    
    rsCTC.Close
    Set rsCTC = Nothing
    
    Call HourglassOff
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function Item(lId As Long) As clsCTCriterion
'----------------------------------------------------------------------------------------'
' return an item by its normal range id
'----------------------------------------------------------------------------------------'

    Set Item = mColCTC(GetKey(lId))

End Function

'----------------------------------------------------------------------------------------'
Private Function GetKey(lId As Long) As String
'----------------------------------------------------------------------------------------'
' return an the key for an item by its CTC id
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
Dim oCTCriterion As clsCTCriterion
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


     For Each oCTCriterion In mColCTC
        With oCTCriterion
            sValue = .ClinicalTest.ClinicalTestGroup.Code
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
    
    PopulateListView = mColCTC.Count

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
        .Add "Grade"
        .Add "Expression"
    End With
    
End Property

'----------------------------------------------------------------------------------------'
Public Sub Delete(lId As Long)
'----------------------------------------------------------------------------------------'
' delete a normal range from the collection and the database
'----------------------------------------------------------------------------------------'
'TA 20/02/2001: Removed mistyped space in SQL so that delete works on Oracle
    MacroADODBConnection.Execute "DELETE FROM CTC WHERE CTCId = " & lId _
                                    & " AND CTCSchemeCode = '" & msSchemeCode & "'"
    mColCTC.Remove GetKey(lId)
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function Update(ByVal lId As Long, sClinicalTestCode As String, nGrade As Integer, _
                        sMin As String, sMax As String, _
                        nMinType As NRFactor, nMaxType As NRFactor, sMessage As String) As Boolean
'----------------------------------------------------------------------------------------'
' add a single CTC to the collection and the database
'----------------------------------------------------------------------------------------'

Dim oCTCriterion As New clsCTCriterion
Dim lNewId As Long
        
    Set oCTCriterion = New clsCTCriterion
    
    With oCTCriterion
        .Id = lId
        .ClinicalTestCode = sClinicalTestCode
        .Grade = nGrade
        .Min = StringtoNumberVariant(sMin)
        .Max = StringtoNumberVariant(sMax)
        .MinType = nMinType
        .MaxType = nMaxType
        'TA 01/03/2001: change local numbers to standard format
        If ValidCTC(oCTCriterion) = -1 Then
            MacroADODBConnection.Execute "UPDATE CTC SET ClinicalTestCode = '" & sClinicalTestCode & "'" _
                                        & ", CTCGrade = " & nGrade _
                                        & ", CTCMin = " & ConvertLocalNumToStandard(VarianttoString(.Min, True)) _
                                        & ", CTCMax = " & ConvertLocalNumToStandard(VarianttoString(.Max, True)) _
                                        & ", CTCMintype = " & nMinType _
                                        & ", CTCMaxType = " & nMaxType _
                                        & " WHERE CTCId = " & lId _
                                        & " AND CTCSchemeCode = '" & msSchemeCode & "'"
                                        
            .CopyInto mColCTC.Item(GetKey(lId))
            Update = True
        Else
            'can't insert
            sMessage = m_ERROR_DUPLICATECTC
            Update = False
        End If
            
    End With

End Function


'----------------------------------------------------------------------------------------'
Public Function Insert(sClinicalTestCode As String, nGrade As Integer, _
                        sMin As String, sMax As String, _
                        nMinType As NRFactor, nMaxType As NRFactor, sMessage As String) As Long
'----------------------------------------------------------------------------------------'
' add a single NormalRange to the collection and the database
' function returns new id or -1 if not added
'----------------------------------------------------------------------------------------'
Dim oCTCriterion As New clsCTCriterion
Dim lNewId As Long
    Set oCTCriterion = New clsCTCriterion
    
    lNewId = GetNewId
    With oCTCriterion
        .Id = lNewId
        .SchemeCode = msSchemeCode
        .ClinicalTestCode = sClinicalTestCode
        .Grade = nGrade
        .Min = StringtoNumberVariant(sMin)
        .Max = StringtoNumberVariant(sMax)
        .MinType = nMinType
        .MaxType = nMaxType
        'TA 01/03/2001: change local numbers to standard format
        If ValidCTC(oCTCriterion) = -1 Then
            MacroADODBConnection.Execute "INSERT INTO CTC values (" _
                                        & lNewId _
                                        & ", '" & msSchemeCode & "'" _
                                        & ", '" & sClinicalTestCode & "'" _
                                        & ", " & nGrade _
                                        & ", " & ConvertLocalNumToStandard(VarianttoString(.Min, True)) _
                                        & ", " & ConvertLocalNumToStandard(VarianttoString(.Max, True)) _
                                        & ", " & nMinType _
                                        & ", " & nMaxType & ")"
            
            'add to collection
            mColCTC.Add oCTCriterion, GetKey(.Id)
            Insert = lNewId
        Else
            'can 't insert
            sMessage = m_ERROR_DUPLICATECTC
            Insert = -1
        End If

    End With
    
    
    
End Function

'----------------------------------------------------------------------------------------'
Private Function ValidCTC(oCTCriterion As clsCTCriterion) As Long
'----------------------------------------------------------------------------------------'
' compares a CTC with the others in the collection
' returns -1 if none found or the conflicting id
'----------------------------------------------------------------------------------------'
Dim oCTCCheck As clsCTCriterion
Dim lConflictingId As Long

    lConflictingId = -1
    
    With oCTCriterion
'no need to test scheme becuase all in same scheme for this collection
         For Each oCTCCheck In mColCTC
            If (.Id <> oCTCCheck.Id) And (.ClinicalTestCode = oCTCCheck.ClinicalTestCode) And (.Expr = oCTCCheck.Expr) Then
                    lConflictingId = oCTCCheck.Id
                    Exit For
            End If
        Next
    End With
    
    ValidCTC = lConflictingId
    
End Function

'----------------------------------------------------------------------------------------'
Public Function NewEnum() As IUnknown
'----------------------------------------------------------------------------------------'

    Set NewEnum = mColCTC.[_NewEnum]

End Function

'----------------------------------------------------------------------------------------'
Public Function GetNewId() As Long
'----------------------------------------------------------------------------------------'
' returns the next id in sequence
'----------------------------------------------------------------------------------------'

Dim rsTemp As ADODB.Recordset
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "SELECT Max(CTCId) FROM CTC WHERE CTCSchemeCode = '" & msSchemeCode & "'", MacroADODBConnection
    If IsNull(rsTemp.Fields(0).Value) Then
        'first ever normal range
        GetNewId = 1
    Else
        GetNewId = rsTemp.Fields(0).Value + 1
    End If
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'

    msSchemeCode = "0"

End Sub

