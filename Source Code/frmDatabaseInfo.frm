VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDatabaseInfo 
   BorderStyle     =   0  'None
   Caption         =   "Databases"
   ClientHeight    =   4800
   ClientLeft      =   4530
   ClientTop       =   3345
   ClientWidth     =   8340
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   10680
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   345
      Left            =   9480
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   1125
   End
   Begin MSComctlLib.ListView lvwDatabaseInfo 
      Height          =   4275
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7541
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblFormName 
      Caption         =   "Database Information "
      Height          =   195
      Left            =   70
      TabIndex        =   3
      Top             =   0
      Width           =   1995
   End
End
Attribute VB_Name = "frmDatabaseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       frmDataBaseInfo.frm
'   Author:     Ashitei Trebi-Ollennu, September 2002
'   Purpose:    Replaces frmNewDataBase.frm which allowed the user to create a new blank database in either SQL
'               or Access for the Macro database and the security database
'-------------------------------------------------------------------------------------

Option Explicit

Public Event SelectedItem(enNodeTag As eSMNodeTag, sDatabaseCode As String, StudyId As Long, sStudyName As String, sSiteCode As String, sUsername As String, sRoleCode As String)

Private msDatabaseCode As String

'----------------------------------------------------------------------------------------------
Public Sub RefreshDatabases()
'----------------------------------------------------------------------------------------------
' Add the databases to their listview
'----------------------------------------------------------------------------------------------

 'Create a variable to add ListItem objects and receive the list of databases.
Dim itmX As MSComctlLib.ListItem
Dim sSQL As String
Dim rsDatabaseList As ADODB.Recordset
Dim vkey As Variant
Dim colDBConInfo As Collection
Dim sDatabaseCode As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM Databases"
    Set rsDatabaseList = New ADODB.Recordset
    rsDatabaseList.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, , adCmdText

    lvwDatabaseInfo.ListItems.Clear
    
    'get a collection of database for the tag
    Set colDBConInfo = frmMenu.DatabaseTagInfo
    
    ' While the record is not the last record, add a ListItem object.
    With rsDatabaseList
        Do Until .EOF = True
            'key used when adding to listview
            sDatabaseCode = rsDatabaseList!DatabaseCode
            vkey = colDBConInfo.Item(sDatabaseCode)
            'this places the databasenames in the listview
            Set itmX = lvwDatabaseInfo.ListItems.Add(, vkey, sDatabaseCode)
            Select Case !DatabaseType
            Case MACRODatabaseType.sqlserver
                itmX.SubItems(1) = "SQL Server/MSDE"
                itmX.SubItems(2) = "Server=" & !ServerName & ";Database=" & !NameOfDatabase
                itmX.SubItems(3) = RemoveNull(!HTMLLocation)
                itmX.SubItems(4) = RemoveNull(!SecureHTMLLocation)
            Case MACRODatabaseType.Oracle80
                itmX.SubItems(1) = "Oracle"
                itmX.SubItems(2) = "Database=" & !NameOfDatabase
                itmX.SubItems(3) = RemoveNull(!HTMLLocation)
                itmX.SubItems(4) = RemoveNull(!SecureHTMLLocation)
            End Select
            .MoveNext   ' Move to next record.
            
        Loop
    End With
   
    rsDatabaseList.Close
    Set rsDatabaseList = Nothing
    
    'Resize listview to text length
    Call lvw_SetAllColWidths(lvwDatabaseInfo, LVSCW_AUTOSIZE_USEHEADER)

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatabaseInfo.RefreshDatabases"
End Sub

'----------------------------------------------------------------------
Public Sub Display()
'----------------------------------------------------------------------
'
'----------------------------------------------------------------------
Dim colmX As MSComctlLib.ColumnHeader
Dim n As Integer
    
    On Error GoTo ErrHandler
    
    'clear listview of protocols
    lvwDatabaseInfo.ListItems.Clear
    
    'add column headers with widths
    Set colmX = lvwDatabaseInfo.ColumnHeaders.Add(, , "Name", 900)
    Set colmX = lvwDatabaseInfo.ColumnHeaders.Add(, , "Type", 900)
    Set colmX = lvwDatabaseInfo.ColumnHeaders.Add(, , "Parameters", 1000)
    Set colmX = lvwDatabaseInfo.ColumnHeaders.Add(, , "HTML Location", 2000)
    Set colmX = lvwDatabaseInfo.ColumnHeaders.Add(, , "Secure HTML Location", 2500)
 
    'set view type
    lvwDatabaseInfo.View = lvwReport
    'set initial sort to ascending on column 0 (Name)
    lvwDatabaseInfo.SortKey = 0
    lvwDatabaseInfo.SortOrder = lvwAscending
    
    RefreshDatabases
    
    'checks if the a row of the listview has been selected
    'if it has then the click event is fired to enable or
    'disable the menus
    For n = 1 To lvwDatabaseInfo.ListItems.Count
        If lvwDatabaseInfo.ListItems(n).Selected = True Then
            Call lvwDatabaseInfo_Click
        End If
    Next
    
    Me.Show

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatabaseInfo.Display"
End Sub

'----------------------------------------------------------------------
Private Sub Form_Resize()
'----------------------------------------------------------------------
' MLM 17/10/02: Created. Make the db list fill the whole form.
'----------------------------------------------------------------------

    lvwDatabaseInfo.Height = Me.Height
    lvwDatabaseInfo.Width = Me.Width

End Sub

'-------------------------------------------------------------------------
Private Sub DatabaseInfoParameters()
'-------------------------------------------------------------------------
    
    If lvwDatabaseInfo.ListItems.Count <> 0 Then
        msDatabaseCode = lvwDatabaseInfo.SelectedItem.Text
    
        RaiseEvent SelectedItem(eSMNodeTag.DatabaseTag, msDatabaseCode, 0, "", "", "", "")
        Call frmMenu.SelectedItemParameters(SelectedDatabaseTag(lvwDatabaseInfo.SelectedItem.Key), msDatabaseCode, 0, "", "", "", "")
    End If
    
End Sub

'-------------------------------------------------------------------------
Private Sub lvwDatabaseInfo_Click()
'-------------------------------------------------------------------------
'ASH 3/12/2002 Added call frmMenu.SelectedItemParameters
'REM 24/04/03 - moved all code from here into DatabaseInfoParameters routine
'-------------------------------------------------------------------------
    
    Call DatabaseInfoParameters
    
End Sub

'------------------------------------------------------------------------------------------
Private Sub lvwDatabaseInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'------------------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------------------
On Error GoTo ErrHandler
    
    Call lvw_Sort(lvwDatabaseInfo, ColumnHeader)

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|lvwDatabaseInfo_ColumnClick"
End Sub

'----------------------------------------------------------------------------------------------------
Private Sub lvwDatabaseInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'----------------------------------------------------------------------------------------------------
'pops up the set database password menu
'-----------------------------------------------------------------------------------------------------
Dim n As Integer
Dim Item As MSComctlLib.ListItem

    Set Item = lvwDatabaseInfo.HitTest(X, Y)

    If Not Item Is Nothing Then
        'check if right button pressed
        If Button = vbRightButton Then
            Item.Selected = True
            Call DatabaseInfoParameters
            'call routine to display form
            Call DisplayPopUpMenu
        End If
    End If

'    For n = 1 To lvwDatabaseInfo.ListItems.Count
'        If lvwDatabaseInfo.ListItems.Item(n).Selected = True Then
'            'check if right button pressed
'            If Button = vbRightButton Then
'                'call routine to display form
'                DisplayPopUpMenu
'            End If
'        End If
'    Next

End Sub

'-------------------------------------------------------------------------------
Private Sub DisplayPopUpMenu()
'-------------------------------------------------------------------------------
'Display the pop up menu with the correct menu items displayed and enabled
'MLM 11/06/03: Added in more menu items including new database timezone
'-------------------------------------------------------------------------------
Dim sMenuItemSelected As String
Dim oMenuItems As clsMenuItems
Dim sTag As String

    On Error GoTo ErrHandler
    
    Set oMenuItems = New clsMenuItems
    
    sTag = lvwDatabaseInfo.SelectedItem.Key
    
    Select Case SelectedDatabaseTag(sTag)
    Case eSMNodeTag.DatabaseTag
        If (goUser.Database.DatabaseCode <> msDatabaseCode) Then
            Call oMenuItems.Add("UNREG", "&Unregister Database", goUser.CheckPermission(gsFnRegisterDB))
        End If
        Call oMenuItems.Add("SETDBPSWD", "Change Database &Password...", goUser.CheckPermission(gsFnChangePassword))
        Call oMenuItems.Add("LOCKADMIN", "&Lock Administration...", goUser.CheckPermission(gsFnRemoveOwnLocks) Or goUser.CheckPermission(gsFnRemoveAllLocks))
        Call oMenuItems.Add("DBTZ", "Database &Timezone...", goUser.CheckPermission(gsFnRegisterDB))

    Case eSMNodeTag.DisconnectedDB
        Call oMenuItems.Add("SETDBPSWD", "Change Database &Password...", goUser.CheckPermission(gsFnChangePassword))
    Case eSMNodeTag.Upgrade
        Call oMenuItems.Add("UPGRADEDB", "Upgrade Database", True)
    End Select
    
    sMenuItemSelected = frmMenu.ShowPopUpMenu(oMenuItems)
    
    Select Case sMenuItemSelected
    Case "UNREG"
        Call frmMenu.UnRegisterDatabase(msDatabaseCode)
    Case "SETDBPSWD"
        Call frmMenu.SetDatabasePasswordForm(msDatabaseCode)
    Case "LOCKADMIN"
        Call frmMenu.LockAdministrationForm
    Case "DBTZ"
        frmMenu.mnuDTimezone_Click
    Case "UPGRADEDB"
        Call frmMenu.UpgradeDatabase
    End Select
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatabaseInfo.DisplayPopUpMenu"
End Sub

'---------------------------------------------------------------------
Private Function SelectedDatabaseTag(sNodeTag As String) As eSMNodeTag
'---------------------------------------------------------------------
' REM 20/06/03
' Returns the Tag for a selected database in list view.
'---------------------------------------------------------------------
Dim vTag As Variant
Dim sTag As String

    'split the tag as it also contains the tool tip
    vTag = Split(sNodeTag, "|")
    'return the first part as this is the tag
    sTag = vTag(0)
    
    Select Case sTag
    Case "D" 'the User node tag
        SelectedDatabaseTag = eSMNodeTag.DatabaseTag
    Case "DD" ' the Users node tag
        SelectedDatabaseTag = eSMNodeTag.DisconnectedDB
    Case "UG" ' Upgrade database
        SelectedDatabaseTag = eSMNodeTag.Upgrade
    End Select
    
End Function

'-------------------------------------------------------------------------------
Public Sub RefreshIfFormVisible()
'-------------------------------------------------------------------------------
    If Me.Visible Then
        Call RefreshDatabases
    End If

End Sub
