VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSiteAdmin 
   Caption         =   "Site Administration"
   ClientHeight    =   4785
   ClientLeft      =   8955
   ClientTop       =   7395
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7950
   Begin VB.CommandButton CmdTrialSiteAdmin 
      Caption         =   "&Study Site Administration..."
      Height          =   375
      Left            =   5460
      TabIndex        =   4
      Top             =   60
      Width           =   2430
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6660
      TabIndex        =   6
      Tag             =   "KeepBottomRight"
      Top             =   4380
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditSite 
      Caption         =   "&Edit Site..."
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdInactive 
      Caption         =   "&Inactive"
      Height          =   375
      Left            =   4020
      TabIndex        =   3
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdActive 
      Caption         =   "A&ctive"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddSite 
      Caption         =   "&Create Site..."
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwSites 
      Height          =   3315
      Left            =   180
      TabIndex        =   5
      Tag             =   "Resize"
      Top             =   780
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5847
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame fraSites 
      Caption         =   "Sites"
      Height          =   3735
      Left            =   60
      TabIndex        =   7
      Top             =   540
      Width           =   7815
   End
End
Attribute VB_Name = "FrmSiteAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       FrmSiteAdmin.frm
'   Author:     Andrew Newbigging, April 1998
'   Purpose:    User can create new RDE sites, make existing sites active/inactive
'   Creating a new site automatically creates an MSMQ queue for that site.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   1    Joanne Lau        22/04/98
'   2    Joanne Lau        24/04/98
'   3    Joanne Lau        30/04/98
'   4    Joanne Lau        30/04/98
'   5    Joanne Lau        8/05/98
'   6    Andrew Newbigging 17/07/98
'   7   Andrew Newbigging   18/2/99
'       Removed calls to MSMQ
'   8   Paul Norris         21/07/99    CmdAddSite_Click now loads a form to manage creating new sites.
'                                       New button to view sites by trials or vice versa
'   9   Paul Norris         21/07/99    Replaced resizing code with resizing object.
'   7   Paul Norris         07/09/99    Upgrade database access code from DAO to ADO
'   8   PN                  15/09/99    Changed call to ADODBConnection() to MacroADODBConnection()
'   PN  21/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   WillC 10/11/99  Added the Error handlers
'   NCJ 14/12/99    Check user's access rights
'                   Added OK button
'   NCJ 21/12/99    Changed "trial" to "study" in captions etc.
'   NCJ 15/1/00     SR2125 Disable Active buttons if user doesn't have correct rights
'   TA 28/04/2000 SR 3327  OK button tag set so that it moves when form resized
'   TA 08/05/2000   sublclassing removed and resizing dome manually
'   WillC 1/9/00   SR3345   Changed captions on buttons to make more sense
'   TA 06/10/2000   Tidied up buttons
'--------------------------------------------------------------------------------

Option Explicit

'TA 16/10/2000: API needed for full row select on version 5 listview
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const LVS_EX_FULLROWSELECT = &H20
Const LVM_FIRST = &H1000
Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + &H37
Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + &H36

Option Base 0
Option Compare Binary

Private Const mcInactive = "Inactive"
Private Const mcActive = "Active"

'form height and width when first shown
Private mlHeight As Long
Private mlWidth As Long
'ASH 11/12/2002
Private mconMACRO As ADODB.Connection
Private msDatabase As String
Private msStudyCode As String
Private msSiteCode As String

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
' OK button added - NCJ 14/12/99
'---------------------------------------------------------------------

    Unload Me

End Sub

'---------------------------------------------------------------------
Private Sub cmdActive_Click()
'---------------------------------------------------------------------
'**  updates the database and refreshes the listbox.
' This causes the selected site to be made active, if it is inactive.
'---------------------------------------------------------------------

  On Error GoTo ErrHandler

    If lvwSites.SelectedItem.SubItems(1) = mcInactive Then
        Call UpdateSiteStatus(lvwSites.SelectedItem, 0)
        Call EnableButtons
    
    End If
 
     
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdActive_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdAddSite_Click()
'---------------------------------------------------------------------
'Allows a new site to be added into the database.
' PN change 8
'---------------------------------------------------------------------

  On Error GoTo ErrHandler

    With frmNewSite
        .IsNewSite = True
        .Database = msDatabase
        .Show vbModal
    End With
    Call RefreshSiteList
    Call EnableButtons
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdAddSite_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdEditSite_Click()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

  On Error GoTo ErrHandler

    With frmNewSite
        .IsNewSite = False
        .Database = msDatabase
        .SiteCode = lvwSites.SelectedItem.Text
        .Show vbModal
    End With
    Call RefreshSiteList
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdEditSite_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
    
End Sub

'---------------------------------------------------------------------
Private Sub UpdateSiteStatus(sSiteName As String, iStatus As Integer)
'---------------------------------------------------------------------
' Make a site active or inactive
'---------------------------------------------------------------------
Dim sSQL As String
Dim oDatabase As MACROUserBS30.Database
Dim bLoad As Boolean
Dim sConnectionString As String

    On Error GoTo ErrHandler

    ' PN 07/09/99
    ' upgrade to ado from dao
    sSQL = "update Site SET SiteStatus = " & iStatus
    sSQL = sSQL & " Where Site = '" & sSiteName & "'"
    mconMACRO.Execute sSQL

    Call RefreshSiteList
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "UpdateSiteStatus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdInactive_Click()
'---------------------------------------------------------------------
' Make a site inactive
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
 
    If lvwSites.SelectedItem.SubItems(1) = mcActive Then
        Call UpdateSiteStatus(lvwSites.SelectedItem, 1)
        Call EnableButtons
    
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdInactive_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
    
End Sub

'---------------------------------------------------------------------
Private Sub CmdTrialSiteAdmin_Click()
'---------------------------------------------------------------------
' DPH 20/08/2002 - Use frmTrialSiteAdminVersioning form
'---------------------------------------------------------------------
Dim sName As String


    On Error GoTo ErrHandler
    sName = ""
    If Not lvwSites.SelectedItem Is Nothing Then 'SDM 01/02/00 SR2861
        sName = lvwSites.SelectedItem.Text
    End If
    ' DPH 20/08/2002 - Use frmTrialSiteAdminVersion form
'    Call frmTrialSiteAdmin.Display(DisplayTrialsBySite, sName)
    Call frmTrialSiteAdminVersioning.Display(msDatabase, DisplayTrialsBySite, sName)
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "CmdTrialSiteAdmin_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
' REVISIONS
' DPH 21/08/2002 - Added Site Location Column
'---------------------------------------------------------------------
    On Error GoTo ErrHandler

    Me.Icon = frmMenu.Icon
    FormCentre Me

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "Form_Load")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub RefreshSiteList()
'---------------------------------------------------------------------
' REVISIONS
' DPH 21/08/2002 - Added Site Location Column
'---------------------------------------------------------------------
Dim sSQL As String
Dim oDatabase As MACROUserBS30.Database
Dim bLoad As Boolean
Dim sConnectionString As String
Dim sMessage As String
Dim oSites As ADODB.Recordset
Dim itmSite As ListItem
Dim lSeletedItemIndex As Long
  
    On Error GoTo ErrHandler

    If Not lvwSites.SelectedItem Is Nothing Then
        lSeletedItemIndex = lvwSites.SelectedItem.Index
    End If
    
    ' lock the window to prevent updates
    Call LockWindow(lvwSites)
    
    'remove existing items
    lvwSites.ListItems.Clear
    
    'Get the list of sites
    Set oDatabase = New MACROUserBS30.Database
    bLoad = oDatabase.Load(SecurityADODBConnection, goUser.UserName, msDatabase, "", False, sMessage)
    sConnectionString = oDatabase.ConnectionString
    Set mconMACRO = New ADODB.Connection
    mconMACRO.Open sConnectionString
    mconMACRO.CursorLocation = adUseClient

    
    ' PN 07/09/99
    ' upgrade to ado from dao
    Set oSites = New ADODB.Recordset
    With oSites
        .Open "select * from Site", mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        'While not the last record, add a list item object.
        
        While Not .EOF
            'SDM 01/02/00 SR2861
            lvwSites.ToolTipText = "Double click to go to the site view for the selected site"
            'add the list item
            Set itmSite = lvwSites.ListItems.Add(, , .Fields("Site"))
            'add text description to SiteStatus
            If .Fields("SiteStatus") = 0 Then
                itmSite.SubItems(1) = mcActive
            Else
                itmSite.SubItems(1) = mcInactive
            End If
            
            If IsNull(.Fields("SiteDescription")) Then
                itmSite.SubItems(2) = vbNullString
            Else
                itmSite.SubItems(2) = .Fields("SiteDescription")
            End If
            
            ' DPH 21/08/2002 - Added Site Location Column
            If IsNull(.Fields("SiteLocation")) Then
                itmSite.SubItems(3) = vbNullString
            Else
                Select Case .Fields("SiteLocation")
                    Case SiteLocation.ESiteServer
                    itmSite.SubItems(3) = "Server"
                    Case SiteLocation.ESiteRemote
                    itmSite.SubItems(3) = "Remote"
                    Case SiteLocation.ESiteServer
                    itmSite.SubItems(3) = vbNullString
                End Select
            End If
            
            .MoveNext
        Wend
        
        .Close
        
    End With
    
    ' unlock the window for updates
    Call UnlockWindow
    On Error GoTo InvalidItemIndex
    lvwSites.ListItems(lSeletedItemIndex).Selected = True
    
    'REM 21/05/03 - reload study site permissions for user
    goUser.ReloadStudySitePermissions
    
InvalidItemIndex:
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "RefreshSiteList")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
    
End Sub


Private Sub Form_Resize()

On Error GoTo ErrHandler

    If Me.WindowState <> vbMinimized Then
        If Me.Height >= mlHeight Then
            fraSites.Height = Me.ScaleHeight - cmdOK.Height - cmdAddSite.Height - 360
            lvwSites.Height = fraSites.Height - 360
            cmdOK.Top = fraSites.Top + fraSites.Height + 120
        End If
        If Me.Width >= mlWidth Then
            fraSites.Width = Me.ScaleWidth - 120
            lvwSites.Width = fraSites.Width - 240
            cmdOK.Left = fraSites.Left + fraSites.Width - cmdOK.Width
        End If
    End If
    
ErrHandler:

End Sub

'---------------------------------------------------------
Private Sub lvwSites_Click()
'---------------------------------------------------------
' Enable buttons according to selected site
'---------------------------------------------------------
       
   Call EnableButtons
    
End Sub

'---------------------------------------------------------
Private Sub EnableUsersButtons()
'---------------------------------------------------------
' NCJ 14/12/99
' Enable/disable buttons according to current user's access rights
'---------------------------------------------------------
    
    If goUser.CheckPermission(gsFnCreateSite) Then
        cmdAddSite.Enabled = True
    Else
        cmdAddSite.Enabled = False
    End If

End Sub

'------------------------------
Private Sub EnableButtons()
'------------------------------

    On Error GoTo ErrHandler
 
    cmdEditSite.Enabled = False
    cmdActive.Enabled = False
    cmdInactive.Enabled = False
    
    If lvwSites.SelectedItem Is Nothing Then
        ' Leave them disabled
    Else
        ' enable the edit site button
        cmdEditSite.Enabled = True
        
        ' enable the active, inactive buttons
        ' NCJ 15/1/00 - Only if user has rights to assign sites to trials etc.
        If goUser.CheckPermission(gsFnAddSiteToTrialOrTrialToSite) Then
            If lvwSites.SelectedItem.SubItems(1) = mcInactive Then
                cmdActive.Enabled = True
                'TA 12/12/2000: disable study/site administration if inactive site
                CmdTrialSiteAdmin.Enabled = False
            Else
                cmdInactive.Enabled = True
                'TA 12/12/2000: enable study/site administration if active site
                CmdTrialSiteAdmin.Enabled = True
            End If
        End If
        
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "EnableButtons")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
    
End Sub

'------------------------------------------------------------------------------
Private Sub lvwSites_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'------------------------------------------------------------------------------
    'When a Coloumn Header is clicked,the listview control is sorted by the
    'items of that column
    'Set the sort key to the index of the column header -1
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
  
    lvwSites.SortKey = ColumnHeader.Index - 1
    
    If lvwSites.SortOrder = lvwAscending Then
        lvwSites.SortOrder = lvwDescending
    Else
        lvwSites.SortOrder = lvwAscending
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "lvwSites_ColumnClick")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
    
End Sub
'-----------------------------------------------------------------
Private Sub lvwSites_DblClick()
'-----------------------------------------------------------------
' DPH 20/08/2002 - Use frmTrialSiteAdminVersioning form
'-----------------------------------------------------------------

   On Error GoTo ErrHandler
   
    ' PN change 8
    If Not lvwSites.SelectedItem Is Nothing Then 'SDM 01/02/00 SR2861
        'TA 9/1/2001: following two checks added
        If goUser.CheckPermission(gsFnAddSiteToTrialOrTrialToSite) Then
            If lvwSites.SelectedItem.SubItems(1) = mcActive Then
                ' DPH 20/08/2002 - Use frmTrialSiteAdminVersioning form
'                frmTrialSiteAdmin.Display DisplayTrialsBySite, lvwSites.SelectedItem.Text
                frmTrialSiteAdminVersioning.Display msDatabase, DisplayTrialsBySite, lvwSites.SelectedItem.Text
            End If
        End If
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lvwSites_DblClick")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
    
End Sub

'-----------------------------------------------------------------
Private Sub lvwSites_ItemClick(ByVal Item As MSComctlLib.ListItem)
'-----------------------------------------------------------------
' Enable buttons when they click on a Site
'-----------------------------------------------------------------

    Call EnableButtons
    
End Sub

'--------------------------------------------------------------------
Public Sub Display(ByVal sDatabase As String, _
                    ByVal sStudy As String, _
                    Optional ByVal sSite As String = "")
'--------------------------------------------------------------------
'
'--------------------------------------------------------------------
Dim ColHdr As ColumnHeader
    
    'On Error GoTo ErrHandler
    
    msDatabase = sDatabase
    msStudyCode = sStudy
    msSiteCode = sSite
    
    mlHeight = 4000 'Me.Height
    mlWidth = 4000 'Me.Width

    'lvwColumnLeft
    Set ColHdr = lvwSites.ColumnHeaders.Add(, , "Site", 1700, lvwColumnLeft)
    Set ColHdr = lvwSites.ColumnHeaders.Add(, , "Site Status", 1700, lvwColumnLeft)
    
    ' PN change 8 - new column for the name of a site
    Set ColHdr = lvwSites.ColumnHeaders.Add(, , "Site Description", 2700, lvwColumnLeft)
    
    ' DPH 21/08/2002 - Added Site Location column
    Set ColHdr = lvwSites.ColumnHeaders.Add(, , "Site Location", 1400, lvwColumnLeft)
    
    With lvwSites
        .View = lvwReport
        .Arrange = lvwAutoLeft
    
        'sort on first column
        .SortKey = 0    ' 0 = Sort using the ListItem object's Text property.
        .SortOrder = lvwAscending
        .Sorted = True
    End With
    
    'Populate the list
    Call RefreshSiteList
    Call EnableUsersButtons
    Call EnableButtons

    'TA 16/10/2000: full row select on version 5 listview
    Call SendMessage(lvwSites.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal LVS_EX_FULLROWSELECT)
    
    Me.Caption = "Site Administration " & "[" & goUser.DatabaseCode & "]"
    
    Me.Show vbModal
     
End Sub
