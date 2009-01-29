VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCTCSchemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsCTCSchemes.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, August 2000
'   Purpose:    Collection of CTC Schemes for MACRO
'----------------------------------------------------------------------------------------'
' Revisions:
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
'----------------------------------------------------------------------------------------'

Option Explicit

Private Const m_DELETE_ERROR = "Unable to delete CTC scheme "
Private Const m_REFERENCES_CTC = "There are associated CTC."
Private Const m_REFERENCES_STUDYDEF = "There are associated study definitions."

Private mcolSchemes As Collection

'----------------------------------------------------------------------------------------'
Private Sub Load()
'----------------------------------------------------------------------------------------'
' load the collection of CTCSchemes
'----------------------------------------------------------------------------------------'

Dim sSQL As String
Dim rsTemp As adodb.Recordset
Dim oScheme As clsCTCScheme

    Call HourglassOn
    
    Set mcolSchemes = New Collection
        
    sSQL = "SELECT * FROM CTCScheme ORDER BY CTCSchemeCode"
    
    Set rsTemp = New adodb.Recordset
    
    rsTemp.Open sSQL, MacroADODBConnection
    
    Do While Not rsTemp.EOF
        Set oScheme = New clsCTCScheme
        With oScheme

            .Code = rsTemp.Fields!CTCSchemeCode
            .Description = rsTemp.Fields!CTCSchemeDescription
            mcolSchemes.Add oScheme, GetKey(.Code)
        End With
        
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    
    Call HourglassOff
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function Item(sCode As String) As clsCTCScheme
'----------------------------------------------------------------------------------------'
    
    Set Item = mcolSchemes(GetKey(sCode))

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
Dim rsCount As adodb.Recordset
Dim bDelete As Boolean
    
    bDelete = True
    
    sMessage = ""
    'check for linked CTC
    Set rsCount = New adodb.Recordset
    rsCount.Open "SELECT COUNT(*) FROM CTC WHERE CTCSchemeCode = '" & sCode & "'", MacroADODBConnection
    If rsCount.Fields(0).Value <> 0 Then
        bDelete = False
        sMessage = sMessage & vbCrLf & m_REFERENCES_CTC
    End If
    rsCount.Close
    Set rsCount = Nothing
    
    'check for linked StudyDefs
    Set rsCount = New adodb.Recordset
    rsCount.Open "SELECT COUNT(*) FROM StudyDefinition WHERE CTCSchemeCode = '" & sCode & "'", MacroADODBConnection
    If rsCount.Fields(0).Value <> 0 Then
        bDelete = False
        sMessage = sMessage & vbCrLf & m_REFERENCES_STUDYDEF
    End If
    rsCount.Close
    Set rsCount = Nothing
    
    If bDelete Then
        MacroADODBConnection.Execute "DELETE FROM CTCScheme WHERE CTCSchemeCode = '" & sCode & "'"
        mcolSchemes.Remove GetKey(sCode)
    Else
        sMessage = m_DELETE_ERROR & Item(sCode).Code & "." & vbCrLf & sMessage
    End If
        
    Delete = bDelete
    
End Function

'----------------------------------------------------------------------------------------'
Public Function Update(sCode As String, sDescription As String) As Boolean
'----------------------------------------------------------------------------------------'
' add a single Scheme to the collection
' returns false if a Scheme with this code already exists
'----------------------------------------------------------------------------------------'
Dim oScheme As clsCTCScheme

    MacroADODBConnection.Execute "UPDATE CTCScheme SET " _
                            & "CTCSchemeDescription = '" & ReplaceQuotes(sDescription) & "'" _
                            & " WHERE CTCSchemeCode = '" & sCode & "'"
    Set oScheme = Item(sCode)
    With oScheme
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
' add a single Scheme to the collection
' returns false if a Scheme with this code already exists
'----------------------------------------------------------------------------------------'
Dim oScheme As clsCTCScheme
    For Each oScheme In mcolSchemes
        If lCase(oScheme.Code) = lCase(sCode) Then
            'code already exists - do not add
            GoTo Errorlabel
        End If
    Next
    

    MacroADODBConnection.Execute "INSERT INTO CTCScheme VALUES ('" & sCode & "','" & ReplaceQuotes(sDescription) & "')"
    'create new Scheme object
    Set oScheme = New clsCTCScheme
    With oScheme
        .Code = sCode
        .Description = sDescription
        mcolSchemes.Add oScheme, GetKey(.Code)
        
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

    Set NewEnum = mcolSchemes.[_NewEnum]

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

'----------------------------------------------------------------------------------------'
Public Function PopulateListView(lvwlistview As MSComctlLib.ListView) As Long
'----------------------------------------------------------------------------------------'
' populate listview with CTC Schemes
'----------------------------------------------------------------------------------------'
Dim itmListItem As MSComctlLib.ListItem
Dim oScheme As clsCTCScheme
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

     For Each oScheme In mcolSchemes
        With oScheme
            Set itmListItem = lvwlistview.ListItems.Add(, , .Code)
            'use code as tag
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

        End With
     Next

    lvwlistview.Visible = bVisible
    
    PopulateListView = mcolSchemes.Count

End Function

