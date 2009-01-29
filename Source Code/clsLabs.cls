VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsLabs.cls
'   Copyright:  InferMed Ltd. 2000-2004. All Rights Reserved
'   Author:     Toby Aldridge, August 2000
'   Purpose:    Collection of Labs for MACRO
'----------------------------------------------------------------------------------------'
' Revisions:
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
' NCJ 9 Oct 00 - Added property Count
' NCJ 12 Oct 00 - Deal with extra Site and Changed fields
' TA 06/12/2000 - When creating lab, add lab and site to sitelab table where site is known
' NCJ 29 Jun 04 - Allow Item to take an index
'----------------------------------------------------------------------------------------'

Option Explicit


Private Const m_DELETE_ERROR = "Unable to delete laboratory "
Private Const m_REFERENCES_NORMALRANGE = "There are associated normal ranges."
Private Const m_REFERENCES_CRFPAGEINSTANCE = "There are associated eForm instances."
Private Const m_REFERENCES_SITE = "There are associated sites."

Private mcolLabs As Collection

'----------------------------------------------------------------------------------------'
Public Sub Load(Optional sSite As String = "")
'----------------------------------------------------------------------------------------'
' Load up the labs collection
' A value of Site means we check the SiteLaboratory table
'----------------------------------------------------------------------------------------'

Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim oLab As clsLab

    Call HourglassOn
    
    Set mcolLabs = New Collection
        
    If sSite = "" Then
        'get all labs
        sSQL = "SELECT * FROM Laboratory ORDER BY LaboratoryCode"
    Else
        'get labs that have this test
        sSQL = "SELECT Laboratory.LaboratoryCode, Laboratory.LaboratoryDescription, " _
                & " Laboratory.Site, Laboratory.Changed " _
                & " FROM Laboratory, SiteLaboratory " _
                & " WHERE Laboratory.LaboratoryCode = SiteLaboratory.LaboratoryCode and SiteLaboratory.Site = '" & sSite & "'" _
                & " ORDER BY Laboratory.LaboratoryCode"
    End If
    
    Set rsTemp = New ADODB.Recordset
    
    rsTemp.Open sSQL, MacroADODBConnection
    
    Do While Not rsTemp.EOF
        Set oLab = New clsLab
        With oLab
            .Code = rsTemp.Fields!LaboratoryCode
            .Description = rsTemp.Fields!LaboratoryDescription
            .Site = rsTemp.Fields!Site
            .Changed = rsTemp.Fields!Changed
            mcolLabs.Add oLab, GetKey(.Code)
        End With
        
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    Call HourglassOff
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function GetKey(sCode As String) As String
'----------------------------------------------------------------------------------------'
' return an the key for an item by its id
'----------------------------------------------------------------------------------------'
    
    GetKey = "K" & Format(sCode)

End Function

'----------------------------------------------------------------------------------------'
Public Function Delete(sCode As String, sMessage As String) As Boolean
'----------------------------------------------------------------------------------------'
' Deletes a Lab
'   Input:  sCode - LabCode
'   Output: function - delete successful?
'           sMessage - reason not successful
'----------------------------------------------------------------------------------------'
Dim rsCount As ADODB.Recordset
Dim bDelete As Boolean
    
    bDelete = True
    
    sMessage = ""
    'check for linked Sites
    Set rsCount = New ADODB.Recordset
    rsCount.Open "SELECT COUNT(*) FROM SiteLaboratory WHERE LaboratoryCode = '" & sCode & "'", MacroADODBConnection
    If rsCount.Fields(0).Value <> 0 Then
        bDelete = False
        sMessage = sMessage & vbCrLf & m_REFERENCES_SITE
    End If
    rsCount.Close
    Set rsCount = Nothing
    
    'check for linked CRFPageInstances
    Set rsCount = New ADODB.Recordset
    rsCount.Open "SELECT COUNT(*) FROM CRFPageInstance WHERE LaboratoryCode = '" & sCode & "'", MacroADODBConnection
    If rsCount.Fields(0).Value <> 0 Then
        bDelete = False
        sMessage = sMessage & vbCrLf & m_REFERENCES_CRFPAGEINSTANCE
    End If
    rsCount.Close
    Set rsCount = Nothing
    
    'check for linked Normal Ranges
    Set rsCount = New ADODB.Recordset
    rsCount.Open "SELECT COUNT(*) FROM NormalRange WHERE LaboratoryCode = '" & sCode & "'", MacroADODBConnection
    If rsCount.Fields(0).Value <> 0 Then
        bDelete = False
        sMessage = sMessage & vbCrLf & m_REFERENCES_NORMALRANGE
    End If
    rsCount.Close
    Set rsCount = Nothing

    If bDelete Then
        MacroADODBConnection.Execute "DELETE FROM Laboratory WHERE LaboratoryCode = '" & sCode & "'"
        mcolLabs.Remove GetKey(sCode)
    Else
        sMessage = m_DELETE_ERROR & Item(sCode).Code & "." & vbCrLf & sMessage
    End If
    
    Delete = bDelete
    
End Function

'----------------------------------------------------------------------------------------'
Public Function Insert(sCode As String, sDescription As String, vSite As Variant) As Boolean
'----------------------------------------------------------------------------------------'
' add a single Lab to the collection
' returns false if a lab with this code already exists
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim oLab As clsLab

    For Each oLab In mcolLabs
        If LCase(oLab.Code) = LCase(sCode) Then
            'code already exists - do not add
            GoTo Errorlabel
        End If
    Next
    
    'create new Lab object
    sSQL = "INSERT INTO Laboratory VALUES ('" _
            & sCode & "','" & ReplaceQuotes(sDescription) & "', " _
            & VarianttoString(vSite, True) & ", " _
            & Changed.Changed & ")"
    MacroADODBConnection.Execute sSQL
    
    If Not IsNull(vSite) Then
        'TA 06/12/2000: have site - must add it to sitelaboratory table
        sSQL = "INSERT INTO SiteLaboratory VALUES ('" & CStr(vSite) & "', '" & sCode & "')"
        MacroADODBConnection.Execute sSQL
    End If
    
    Set oLab = New clsLab
    With oLab
        .Code = sCode
        .Description = sDescription
        .Site = vSite
        .Changed = Changed.Changed
        mcolLabs.Add oLab, GetKey(.Code)
    End With
    
    Insert = True
Exit Function

Errorlabel:
    Insert = False
Exit Function

End Function

'----------------------------------------------------------------------------------------'
Public Property Get Item(sCode As String) As clsLab
'----------------------------------------------------------------------------------------'
' Return an item by its lab code
' NCJ 29 Jun 04 - Allow indexes too!
'----------------------------------------------------------------------------------------'
    
    If IsNumeric(sCode) Then
        ' Interpret as Index
        Set Item = mcolLabs.Item(CInt(sCode))
    Else
        Set Item = mcolLabs(GetKey(sCode))
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
'----------------------------------------------------------------------------------------'
    
    Set NewEnum = mcolLabs.[_NewEnum]

End Function

#If WebRDE <> -1 Then
'----------------------------------------------------------------------------------------'
Public Function PopulateListView(lvwlistview As MSComctlLib.IListView, Optional bShowSitesChanged = True) As Long
'----------------------------------------------------------------------------------------'
' populate a listview with laboratories
' returns the no. of rows in the listview
'----------------------------------------------------------------------------------------'
Dim itmListItem As MSComctlLib.ListItem
Dim oLab As clsLab
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
    If bShowSitesChanged Then
        lvwlistview.ColumnHeaders.Add , , "Site", lvwlistview.Parent.TextWidth("Site") + 12 * Screen.TwipsPerPixelX
        'lvwlistview.ColumnHeaders.Add , , "Changed", lvwlistview.Parent.TextWidth("Site") + 12 * Screen.TwipsPerPixelX
    End If
     For Each oLab In mcolLabs
        With oLab
            Set itmListItem = lvwlistview.ListItems.Add(, , .Code)
            'use the code as the tag
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
            If bShowSitesChanged Then
                itmListItem.SubItems(2) = .SiteText
                If lvwlistview.ColumnHeaders(3).Width < (lvwlistview.Parent.TextWidth(itmListItem.SubItems(2)) + 12 * Screen.TwipsPerPixelX) Then
                  lvwlistview.ColumnHeaders(3).Width = (lvwlistview.Parent.TextWidth(itmListItem.SubItems(2)) + 12 * Screen.TwipsPerPixelX)
                End If
    '            itmListItem.SubItems(3) = .ChangedText
    '            If lvwlistview.ColumnHeaders(4).Width < (lvwlistview.Parent.TextWidth(itmListItem.SubItems(3)) + 12 * Screen.TwipsPerPixelX) Then
    '              lvwlistview.ColumnHeaders(4).Width = (lvwlistview.Parent.TextWidth(itmListItem.SubItems(3)) + 12 * Screen.TwipsPerPixelX)
    '            End If
            End If
        End With
     Next

    lvwlistview.Visible = bVisible
    
    PopulateListView = mcolLabs.Count
    
End Function
#End If

'----------------------------------------------------------------------------------------'
Public Function Update(sCode As String, sDescription As String) As Boolean
'----------------------------------------------------------------------------------------'
' add a single Lab to the collection
' returns false if a lab with this code already exists
'----------------------------------------------------------------------------------------'
Dim oLab As clsLab
Dim sSQL As String

    'current record - update
    sSQL = "UPDATE Laboratory SET " _
            & " LaboratoryDescription = '" & ReplaceQuotes(sDescription) & "'," _
            & " Changed = " & Changed.Changed _
            & " WHERE LaboratoryCode='" & sCode & "'"
    MacroADODBConnection.Execute sSQL

    Set oLab = Item(sCode)
    With oLab
        .Description = sDescription
        .Changed = Changed.Changed
    End With
    
    Update = True
Exit Function

Errorlabel:
    Update = False
Exit Function
End Function

'----------------------------------------------------------------------------------------'
Public Property Get Count() As Integer
'----------------------------------------------------------------------------------------'
' NCJ 9 Oct 00 - Count of labs in collection
'----------------------------------------------------------------------------------------'

    Count = mcolLabs.Count
    
End Property

