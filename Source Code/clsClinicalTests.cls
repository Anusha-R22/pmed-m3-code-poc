VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClinicalTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'----------------------------------------------------------------------------------------'
'   File:       clsClinicalTest.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, August 2000
'   Purpose:    Collection of ClinicalTests for MACRO
'----------------------------------------------------------------------------------------'
' Revisions:
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
'----------------------------------------------------------------------------------------'

Option Explicit

Private Const m_DELETE_ERROR = "Unable to delete test "
Private Const m_REFERENCES_NORMALRANGE = "There are associated normal ranges."
Private Const m_REFERENCES_CTC = "There are associated CTC."
Private Const m_REFERENCES_DATAITEM = "There are associated questions."

Private mcolClinicalTests As Collection

'----------------------------------------------------------------------------------------'
Private Sub Load()
'----------------------------------------------------------------------------------------'
' load ClinicalTests collection
'----------------------------------------------------------------------------------------'

Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim oClinicalTest As clsClinicalTest

    Call HourglassOn
    
    Set mcolClinicalTests = New Collection
        
    sSQL = "SELECT * FROM ClinicalTest ORDER BY ClinicalTestCode"
    
    Set rsTemp = New ADODB.Recordset
    
    rsTemp.Open sSQL, MacroADODBConnection
    
    Do While Not rsTemp.EOF
        Set oClinicalTest = New clsClinicalTest
        With oClinicalTest
            .Code = RemoveNull(rsTemp.Fields!ClinicalTestCode)
            .Description = rsTemp.Fields!ClinicalTestDescription
            .ClinicalTestGroupCode = rsTemp.Fields!ClinicalTestGroupCode
            'removenull becasue Oracle treats empty string as null
            .Unit = RemoveNull(rsTemp.Fields!Unit)
            mcolClinicalTests.Add oClinicalTest, GetKey(.Code)
        End With
        
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    
    Call HourglassOff
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function PopulateCombo(cboCombo As ComboBox, sClinicalTestGroupCode As String) As Long
'----------------------------------------------------------------------------------------'
'populate a combo with clinical test filtered by test group
'returns number of items added
'----------------------------------------------------------------------------------------'
Dim oClinicalTest As clsClinicalTest
    cboCombo.Clear

    For Each oClinicalTest In mcolClinicalTests
        If sClinicalTestGroupCode = oClinicalTest.ClinicalTestGroupCode Then
            cboCombo.AddItem oClinicalTest.Code
        End If
    Next
    
    PopulateCombo = cboCombo.ListCount
    
End Function

'----------------------------------------------------------------------------------------'
Public Function Item(sCode As String) As clsClinicalTest
'----------------------------------------------------------------------------------------'
        
    Set Item = mcolClinicalTests(GetKey(sCode))

End Function
    
'----------------------------------------------------------------------------------------'
Public Sub Refresh()
'----------------------------------------------------------------------------------------'
    
    Load

End Sub

'----------------------------------------------------------------------------------------'
Public Function Count() As Long
'----------------------------------------------------------------------------------------'

    Count = mcolClinicalTests.Count

End Function

'----------------------------------------------------------------------------------------'
Public Function Delete(sCode As String, sMessage As String) As Boolean
'----------------------------------------------------------------------------------------'
' Deletes a ClinicalTestGroup
'   Input:  sCode - ClinicalGroupCode
'   Output: function - delete successful?
'           sMessage - reason not successful
'----------------------------------------------------------------------------------------'
Dim rsCount As ADODB.Recordset
Dim bDelete As Boolean
    
    bDelete = True
    
    sMessage = ""
    'check for linked normal ranges
    Set rsCount = New ADODB.Recordset
    rsCount.Open "SELECT COUNT(*) FROM NormalRange WHERE ClinicalTestCode = '" & sCode & "'", MacroADODBConnection
    If rsCount.Fields(0).Value <> 0 Then
        bDelete = False
        sMessage = sMessage & vbCrLf & m_REFERENCES_NORMALRANGE
    End If
    rsCount.Close
    Set rsCount = Nothing
    
    'check for linked CTC
    Set rsCount = New ADODB.Recordset
    rsCount.Open "SELECT COUNT(*) FROM CTC WHERE ClinicalTestCode = '" & sCode & "'", MacroADODBConnection
    If rsCount.Fields(0).Value <> 0 Then
        bDelete = False
        sMessage = sMessage & vbCrLf & m_REFERENCES_CTC
    End If
    rsCount.Close
    Set rsCount = Nothing
    
    'check for linked dataitems
    Set rsCount = New ADODB.Recordset
    rsCount.Open "SELECT COUNT(*) FROM DataItem WHERE ClinicalTestCode = '" & sCode & "'", MacroADODBConnection
    If rsCount.Fields(0).Value <> 0 Then
        bDelete = False
        sMessage = sMessage & vbCrLf & m_REFERENCES_DATAITEM
    End If
    rsCount.Close
    Set rsCount = Nothing
    
    'TA 26/10/2000: sCode now in quotes
    If bDelete Then
        MacroADODBConnection.Execute "DELETE FROM ClinicalTest WHERE ClinicalTestCode = '" & sCode & "'"
        mcolClinicalTests.Remove GetKey(sCode)
    Else
        sMessage = m_DELETE_ERROR & Item(sCode).Code & "." & vbCrLf & sMessage
    End If
    
    Delete = bDelete
    
End Function

'----------------------------------------------------------------------------------------'
Public Function Update(sCode As String, sDescription As String, sClinicalTestGroupCode As String, sUnit As String) As Boolean
'----------------------------------------------------------------------------------------'
' add a single ClinicalTest to the collection
' returns false if a ClinicalTest with this code already exists
'----------------------------------------------------------------------------------------'
Dim oClinicalTest As clsClinicalTest
        MacroADODBConnection.Execute "UPDATE ClinicalTest SET " _
                        & "ClinicalTestDescription = '" & ReplaceQuotes(sDescription) & "', " _
                        & "ClinicalTestGroupCode = '" & sClinicalTestGroupCode & "', Unit = '" & sUnit & "'" _
                        & " WHERE ClinicalTestCode = '" & sCode & "'"
    Set oClinicalTest = Item(sCode)
    With oClinicalTest
        .Description = sDescription
        'TA 3/11/2000: update clinicaltestgroupcode
        .ClinicalTestGroupCode = sClinicalTestGroupCode
        .Unit = sUnit
    End With
    
    Update = True
Exit Function

Errorlabel:
    Update = False
Exit Function

End Function


'----------------------------------------------------------------------------------------'
Public Function Insert(sCode As String, sDescription As String, sClinicalTestGroupCode As String, sUnit As String) As Boolean
'----------------------------------------------------------------------------------------'
' add a single ClinicalTest to the collection
' returns false if a ClinicalTest with this code already exists
'----------------------------------------------------------------------------------------'
Dim oClinicalTest As clsClinicalTest
    For Each oClinicalTest In mcolClinicalTests
        If lCase(oClinicalTest.Code) = lCase(sCode) Then
            'code already exists - do not add
            GoTo Errorlabel
        End If
    Next
    
    MacroADODBConnection.Execute "INSERT INTO ClinicalTest VALUES (" _
                                & "'" & sCode & "','" & ReplaceQuotes(sDescription) & "'" _
                                & ", '" & sClinicalTestGroupCode & "', '" & sUnit & "')"

    'create new ClinicalTest object
    Set oClinicalTest = New clsClinicalTest
    With oClinicalTest
        .Code = sCode
        .Description = sDescription
        .ClinicalTestGroupCode = sClinicalTestGroupCode
        .Unit = sUnit
        mcolClinicalTests.Add oClinicalTest, GetKey(.Code)
        
    End With
    
    Insert = True
Exit Function

Errorlabel:
    Insert = False
Exit Function

End Function


'----------------------------------------------------------------------------------------'
Public Function NewEnum() As IUnknown
'----------------------------------------------------------------------------------------'

    Set NewEnum = mcolClinicalTests.[_NewEnum]

End Function

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'
    
    'load collectiom
    Call Load

End Sub

'----------------------------------------------------------------------------------------'
Private Function GetKey(sCode As String) As String
'----------------------------------------------------------------------------------------'

    GetKey = "K" & Format(sCode)

End Function


Public Function PopulateListView(lvwlistview As MSComctlLib.ListView) As Long
'----------------------------------------------------------------------------------------'
'poupulate a listview with Clinical Tests
'----------------------------------------------------------------------------------------'
Dim itmListItem As MSComctlLib.ListItem
Dim oClinicalTest As clsClinicalTest
Dim olistItem As MSComctlLib.ListItem
Dim bVisible As Boolean

    bVisible = lvwlistview.Visible
    lvwlistview.Visible = False
    
    'clear
    lvwlistview.ListItems.Clear
    lvwlistview.ColumnHeaders.Clear
    'add headings
    lvwlistview.ColumnHeaders.Add , , "Code", lvwlistview.Parent.TextWidth("Code") + 12 * Screen.TwipsPerPixelX
    lvwlistview.ColumnHeaders.Add , , "Description", lvwlistview.Parent.TextWidth("Description") + 12 * Screen.TwipsPerPixelX
    lvwlistview.ColumnHeaders.Add , , "Test Group", lvwlistview.Parent.TextWidth("Test Group") + 12 * Screen.TwipsPerPixelX
    lvwlistview.ColumnHeaders.Add , , "Unit", lvwlistview.Parent.TextWidth("Description") + 12 * Screen.TwipsPerPixelX
     For Each oClinicalTest In mcolClinicalTests
        With oClinicalTest
            Set itmListItem = lvwlistview.ListItems.Add(, , .Code)
            itmListItem.Tag = Format(.Code)
            If lvwlistview.ColumnHeaders(1).Width < (lvwlistview.Parent.TextWidth(.Code) + 6 * Screen.TwipsPerPixelX) Then
                lvwlistview.ColumnHeaders(1).Width = (lvwlistview.Parent.TextWidth(.Code) + 6 * Screen.TwipsPerPixelX)
            End If
            'make description one line
            If InStr(.Description, vbCrLf) = 0 Then
                itmListItem.SubItems(1) = .Description
            Else
                itmListItem.SubItems(1) = Left(.Description, InStr(.Description, vbCrLf) - 1) & " ..."
            End If
            If lvwlistview.ColumnHeaders(2).Width < (lvwlistview.Parent.TextWidth(itmListItem.SubItems(1)) + 12 * Screen.TwipsPerPixelX) Then
              lvwlistview.ColumnHeaders(2).Width = (lvwlistview.Parent.TextWidth(itmListItem.SubItems(1)) + 12 * Screen.TwipsPerPixelX)
            End If
            itmListItem.SubItems(2) = .ClinicalTestGroup.Code
            If lvwlistview.ColumnHeaders(3).Width < (lvwlistview.Parent.TextWidth(.ClinicalTestGroup.Code) + 12 * Screen.TwipsPerPixelX) Then
              lvwlistview.ColumnHeaders(3).Width = (lvwlistview.Parent.TextWidth(.ClinicalTestGroup.Code) + 12 * Screen.TwipsPerPixelX)
            End If
            itmListItem.SubItems(3) = .Unit
            If lvwlistview.ColumnHeaders(4).Width < (lvwlistview.Parent.TextWidth(.Unit) + 12 * Screen.TwipsPerPixelX) Then
              lvwlistview.ColumnHeaders(4).Width = (lvwlistview.Parent.TextWidth(.Unit) + 12 * Screen.TwipsPerPixelX)
            End If
        End With
     Next

    lvwlistview.Visible = bVisible

    PopulateListView = mcolClinicalTests.Count

End Function

