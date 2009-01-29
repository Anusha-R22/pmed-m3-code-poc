VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClinTestGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsClinTestGroups.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, August 2000
'   Purpose:    Collection of ClinicalTestGroups for MACRO
'----------------------------------------------------------------------------------------'
' Revisions:
'TA 27/09/2000: shortened class name from clsClinicalTestGroups to avoid compilation error
'----------------------------------------------------------------------------------------'

Option Explicit

Private Const m_DELETE_ERROR = "Unable to delete test group "
Private Const m_REFERENCES_CLINICALTEST = "There are associated tests."

Private mcolClinicalTestGroups As Collection


'----------------------------------------------------------------------------------------'
Private Sub Load()
'----------------------------------------------------------------------------------------'
' load collection of Clinical Test groups from database
'----------------------------------------------------------------------------------------'

Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim oClinicalTestGroup As clsClinTestGroup

    Call HourglassOn
    
    Set mcolClinicalTestGroups = New Collection
        
    sSQL = "SELECT * FROM ClinicalTestGroup ORDER BY ClinicalTestGroupCode"
    
    Set rsTemp = New ADODB.Recordset
    
    rsTemp.Open sSQL, MacroADODBConnection
    
    Do While Not rsTemp.EOF
        Set oClinicalTestGroup = New clsClinTestGroup
        With oClinicalTestGroup
            .Code = rsTemp.Fields!ClinicalTestGroupCode
            .Description = rsTemp.Fields!ClinicalTestGroupDescription
            mcolClinicalTestGroups.Add oClinicalTestGroup, GetKey(.Code)
        End With
        
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    Call HourglassOff
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function PopulateCombo(cboCombo As ComboBox)
'----------------------------------------------------------------------------------------'
'populate combo with clinical test groups
'----------------------------------------------------------------------------------------'

Dim oClinicalTestGroup As clsClinTestGroup
    cboCombo.Clear
    For Each oClinicalTestGroup In mcolClinicalTestGroups
        cboCombo.AddItem oClinicalTestGroup.Code
    Next

End Function

'----------------------------------------------------------------------------------------'
Public Function Item(sCode As String) As clsClinTestGroup
'----------------------------------------------------------------------------------------'
    
    Set Item = mcolClinicalTestGroups(GetKey(sCode))

End Function
    
'----------------------------------------------------------------------------------------'
Public Sub Refresh()
'----------------------------------------------------------------------------------------'
    
    Load

End Sub

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
    
    Set rsCount = New ADODB.Recordset
    rsCount.Open "SELECT COUNT(*) FROM ClinicalTest WHERE ClinicalTestGroupCode = '" & sCode & "'", MacroADODBConnection
    
    If rsCount.Fields(0).Value <> 0 Then
        bDelete = False
        sMessage = sMessage & vbCrLf & m_REFERENCES_CLINICALTEST
    End If
    
    rsCount.Close
    Set rsCount = Nothing
    
    If bDelete Then
        MacroADODBConnection.Execute "DELETE FROM ClinicalTestGroup WHERE ClinicalTestGroupCode = '" & sCode & "'"
        mcolClinicalTestGroups.Remove GetKey(sCode)
    Else
        sMessage = m_DELETE_ERROR & Item(sCode).Code & "." & vbCrLf & sMessage
    End If

    Delete = bDelete
    
End Function

'----------------------------------------------------------------------------------------'
Public Function Update(sCode As String, sDescription As String) As Boolean
'----------------------------------------------------------------------------------------'
' add a single ClinicalTestGroup to the collection
' returns false if a ClinicalTestGroup with this code already exists
'----------------------------------------------------------------------------------------'
Dim oClinicalTestGroup As clsClinTestGroup
        MacroADODBConnection.Execute "UPDATE ClinicalTestGroup SET " _
                        & " ClinicalTestGroupDescription = '" & ReplaceQuotes(sDescription) & "'" _
                        & " WHERE ClinicalTestGroupCode = '" & sCode & "'"
    Set oClinicalTestGroup = Item(sCode)
    With oClinicalTestGroup
        .Description = sDescription
    End With
    
    Update = True
Exit Function

Errorlabel:
    Update = False
Exit Function

End Function


'----------------------------------------------------------------------------------------'
Public Function Insert(sCode As String, sDescription As String) As Boolean
'----------------------------------------------------------------------------------------'
' add a single ClinicalTestGroup to the collection
' returns false if a ClinicalTestGroup with this code already exists
'----------------------------------------------------------------------------------------'
Dim oClinicalTestGroup As clsClinTestGroup
    For Each oClinicalTestGroup In mcolClinicalTestGroups
        If lCase(oClinicalTestGroup.Code) = lCase(sCode) Then
            'code already exists - do not add
            GoTo Errorlabel
        End If
    Next
    

    MacroADODBConnection.Execute "INSERT INTO ClinicalTestGroup VALUES ('" & sCode & "','" & ReplaceQuotes(sDescription) & "')"
    'create new ClinicalTestGroup object
    Set oClinicalTestGroup = New clsClinTestGroup
    With oClinicalTestGroup
        .Code = sCode
        .Description = sDescription
        mcolClinicalTestGroups.Add oClinicalTestGroup, GetKey(.Code)
        
    End With
    
    Insert = True
Exit Function

Errorlabel:
    Insert = False
Exit Function

End Function

'----------------------------------------------------------------------------------------'
Public Function Count() As Long
'----------------------------------------------------------------------------------------'

    Count = mcolClinicalTestGroups.Count

End Function

'----------------------------------------------------------------------------------------'
Public Function NewEnum() As IUnknown
'----------------------------------------------------------------------------------------'

    Set NewEnum = mcolClinicalTestGroups.[_NewEnum]

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
' populate a listview with ClinicalTestGroup collection
'----------------------------------------------------------------------------------------'
Dim itmListItem As MSComctlLib.ListItem
Dim oClinicalTestGroup As clsClinTestGroup
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

     For Each oClinicalTestGroup In mcolClinicalTestGroups
        With oClinicalTestGroup
            Set itmListItem = lvwlistview.ListItems.Add(, , .Code)
            itmListItem.Tag = Format(.Code)
            'use code as tag
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

        End With
     Next

    lvwlistview.Visible = bVisible
    
    PopulateListView = mcolClinicalTestGroups.Count

End Function


