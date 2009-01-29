VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserRolesInfo 
   BorderStyle     =   0  'None
   Caption         =   "User Roles Information"
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   1020
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserRolesInfo.frx":0000
            Key             =   "activeUser"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserRolesInfo.frx":0452
            Key             =   "inactiveUser"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwUserRolesInfo 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   7011
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblFormName 
      BackColor       =   &H8000000B&
      Caption         =   "User Role Information "
      Height          =   195
      Left            =   70
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmUserRolesInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       frmUserRolesInfo.frm
'   Author:     Ashitei Trebi-Ollennu, September 2002
'   Purpose:    Displays roles assigned to  a user on studies and sites
'------------------------------------------------------------------------------
'revisions
'------------------------------------------------------------------------------
'TA 20/01/2006: Check roles, sites, users against in memory list instead of in database for performance
'------------------------------------------------------------------------------


Option Explicit

Public Event SelectedItem(enNodeTag As eSMNodeTag, sDatabaseCode As String, StudyId As Long, sStudyName As String, sSiteCode As String, sUsername As String, sRoleCode As String)
Private mconMACRO As ADODB.Connection
Private msDatabase As String
Private msUserName As String
Private msRoleCode As String
Private msStudyCode As String
Private msSiteCode As String
Private Const msDB_DISCONNECTED_ICON = "databasedisconnected"
'----------------------------------------------------------------------------------
Private Sub Form_Load()
'----------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------
Dim colmX As MSComctlLib.ColumnHeader

    'Me.Icon = frmMenu.Icon
    
    'clear listview
    lvwUserRolesInfo.ListItems.Clear
    
    'add column headers with widths
    Set colmX = lvwUserRolesInfo.ColumnHeaders.Add(, , "User", 1000)
    Set colmX = lvwUserRolesInfo.ColumnHeaders.Add(, , "Full Name", 1500)
    Set colmX = lvwUserRolesInfo.ColumnHeaders.Add(, , "Study", 1000)
    Set colmX = lvwUserRolesInfo.ColumnHeaders.Add(, , "Site", 1000)
    Set colmX = lvwUserRolesInfo.ColumnHeaders.Add(, , "Role", 1000)
 
    'set view type
    lvwUserRolesInfo.View = lvwReport
    'set initial sort to ascending on column 0 (name)
    lvwUserRolesInfo.SortKey = 0
    lvwUserRolesInfo.SortOrder = lvwAscending
    
    FormCentre Me

End Sub

'------------------------------------------------------------------------------------
Public Sub Display(ByVal sDatabase As String, _
                    ByVal sStudy As String, _
                    ByVal sSite As String, eNodeTag As eSMNodeTag)
'------------------------------------------------------------------------------------
'TA 20/01/2006: Check roles, sites, users against in memory list instead of in database for performance
'------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsUserDetails As ADODB.Recordset
Dim rsUserNameFull As ADODB.Recordset
Dim rsDatabase As ADODB.Recordset
Dim itmX As MSComctlLib.ListItem
Dim sMsg As String
Dim sMessage As String
Dim skey As String
Dim n As Integer
Dim oDatabase As MACROUserBS30.Database
Dim bLoad As Boolean
Dim sConnectionString As String

    On Error GoTo ErrHandler
    
    If sDatabase = "" Then Exit Sub
    
    If eNodeTag = eSMNodeTag.DisconnectedDB Then
        
        lvwUserRolesInfo.ListItems.Clear
        Call lvw_SetAllColWidths(lvwUserRolesInfo, LVSCW_AUTOSIZE_USEHEADER)
    
        Me.Show
        Exit Sub
    End If
    
    Set oDatabase = New MACROUserBS30.Database
    
    bLoad = oDatabase.Load(SecurityADODBConnection, goUser.UserName, sDatabase, "", False, sMessage)
    
    sConnectionString = oDatabase.ConnectionString
    Set mconMACRO = New ADODB.Connection
    sMsg = "Database " & sDatabase & " has been unregistered." & vbCrLf & _
    "You need to re-register it before using."
    If sConnectionString = "" Then
        DialogInformation sMsg
        lvwUserRolesInfo.ListItems.Clear
        frmMenu.RefereshTreeView
        Exit Sub
    End If
    mconMACRO.Open sConnectionString
    mconMACRO.CursorLocation = adUseClient
    
    If sStudy = "" And sSite = "" Then
        sSQL = "SELECT * FROM UserRole"
    Else
        sSQL = "SELECT * FROM UserRole WHERE"
        If sStudy <> "" Then
            sSQL = sSQL & " AND (StudyCode = '" & ALL_STUDIES & "' OR StudyCode = '" & sStudy & "')"
        End If
        If sSite <> "" Then
            sSQL = sSQL & " AND (SiteCode = '" & ALL_SITES & "' OR SiteCode = '" & sSite & "')"
        End If
        sSQL = Replace(sSQL, "WHERE AND", "WHERE")
    End If
    
    Set rsUserDetails = New ADODB.Recordset
    rsUserDetails.Open sSQL, mconMACRO, adOpenKeyset, adLockOptimistic, adCmdText
    
    'if recordcount = 0 then there are no user roles assigned for the study site combonation
    If rsUserDetails.RecordCount <= 0 Then
        lvwUserRolesInfo.ListItems.Clear
        Exit Sub
    End If
    
    lvwUserRolesInfo.ListItems.Clear
    rsUserDetails.MoveFirst
 
 
    'TA 20/1/2006: Store user details in memory
    Dim rsUser As ADODB.Recordset
    Set rsUser = GetDisconnectedRecordset(SecurityADODBConnection, "SELECT Username,UserNameFull,Enabled,FailedAttempts FROM MacroUser")
'    'TA 20/1/2006: Store usernamefull in memory
'    Dim colUsers As Collection
'    Set colUsers = GetCollectionFromSQL(SecurityADODBConnection, "SELECT Username,UserNameFull FROM MacroUser")
     'TA 20/1/2006: Store site active/inactive in memory
    Dim colSites As Collection
    Set colSites = GetCollectionFromSQL(mconMACRO, "SELECT Site,SiteStatus FROM site")
     'TA 20/1/2006: Store site role list in memory
    Dim colRoles As Collection
    Set colRoles = GetCollectionFromSQL(SecurityADODBConnection, "SELECT RoleCode,Enabled FROM Role")
     'TA 20/1/2006: Store studyname/studyid in memory
    Dim colStudies As Collection
    Set colStudies = GetCollectionFromSQL(mconMACRO, "SELECT ClinicalTrialName,ClinicalTrialId FROM clinicaltrial")
     'TA 20/1/2006: Store trialsite information in memory
    Dim rsTrialSite As ADODB.Recordset
    Set rsTrialSite = GetDisconnectedRecordset(mconMACRO, "SELECT * FROM TrialSite")
     'TA 20/1/2006: Store max password retires in memory
    Dim lMaxPasswordRetries As Long
    lMaxPasswordRetries = PasswordRetries
    
    Do Until rsUserDetails.EOF
        rsUser.Filter = adFilterNone
        rsUser.Filter = "Username = '" & rsUserDetails.Fields("UserName").Value & "'"
        If rsUser.RecordCount > 0 Then
            rsUser.MoveLast
        'If CollectionMember(colUsers, rsUserDetails.Fields("UserName").Value, False) Then
            'username exists in security database
            sMessage = ""
            skey = sDatabase & "|" & rsUserDetails!UserName & "|" & rsUserDetails!RoleCode & "|" & rsUserDetails!SiteCode _
                & "|" & rsUserDetails!StudyCode & "|" & rsUser.Fields("USERNAMEFULL")
            Set itmX = lvwUserRolesInfo.ListItems.Add(, skey, rsUserDetails!UserName)
            'check for user status first because users are top of the tree since there must be users
            'for trials and sites to be assigned to them
            sMessage = CheckForUserStatus(rsUserDetails!UserName, rsUserDetails!RoleCode, lMaxPasswordRetries, colRoles, rsUser)
            If sMessage = "" Then
                'finally check for sites attached to studies
                sMessage = CheckForTrialSite(rsUserDetails.Fields("STUDYCODE").Value, rsUserDetails.Fields("SITECODE").Value, colStudies, rsTrialSite)
                'now check for sites since if the sites are not set up, there is no possibilty
                'of them being assigned to a trial or study
                If sMessage = "" Then
                    sMessage = CheckForInActiveSite(rsUserDetails!SiteCode, colSites)
                End If
            End If
            
            'sets up icon if message returned
            If sMessage = "" Then
                itmX.SmallIcon = "activeUser"
            Else
                itmX.SmallIcon = "inactiveUser"
                itmX.ToolTipText = sMessage
            End If
            
            itmX.SubItems(1) = rsUser!UserName
            itmX.SubItems(2) = rsUserDetails!StudyCode
            itmX.SubItems(3) = rsUserDetails!SiteCode
            itmX.SubItems(4) = rsUserDetails!RoleCode
        End If
        rsUserDetails.MoveNext
    Loop
    
    'TA 20/1/2006: clean up
    rsUser.Close
    rsTrialSite.Close
    
    
    Call lvw_SetAllColWidths(lvwUserRolesInfo, LVSCW_AUTOSIZE_USEHEADER)
    
    'checks if the a row of the listview has been selected
    'if it has then the click event is fired to enable or
    'disable the menus
    For n = 1 To lvwUserRolesInfo.ListItems.Count
        If lvwUserRolesInfo.ListItems(n).Selected = True Then
            Call lvwUserRolesInfo_Click
        End If
    Next
    
    Set rsUserDetails = Nothing
    
    Me.Show

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserRolesInfo.Display"
End Sub

'------------------------------------------------------------------------------------
Private Function CheckForTrialSite(ByVal sStudy As String, _
                                    ByVal sSite As String, _
                                    colStudies As Collection, rsTrialSite As ADODB.Recordset) As String
'------------------------------------------------------------------------------------
'checks for sites and returns message to be displayed in the tooltip
'------------------------------------------------------------------------------------
Dim sSQL As String
'Dim rsTrialSite As ADODB.Recordset
Dim lTrialId As Long

    On Error GoTo ErrHandler
    
    CheckForTrialSite = ""
    If Not CollectionMember(colStudies, sStudy, False) Then
        'simulate  GetClinicalTrialID(sStudy) which returns -1 when no study match is found
        lTrialId = -1
        'lTrialId = GetClinicalTrialID(sStudy)
    Else
        lTrialId = CLng(colStudies(sStudy))
    End If
    
    'RecordSet should contain All Field Data
    
    
    If sStudy <> ALL_STUDIES And sSite <> ALL_SITES Then
        rsTrialSite.Filter = "ClinicalTrialID = " & lTrialId & " AND TrialSite = '" & sSite & "'"
        'sSQL = "SELECT * FROM TrialSite WHERE ClinicalTrialID = " & lTrialId
        'sSQL = sSQL & " AND TrialSite = '" & sSite & "'"
    ElseIf sStudy = ALL_STUDIES And sSite <> ALL_SITES Then
        rsTrialSite.Filter = "TrialSite = '" & sSite & "'"
        'sSQL = "SELECT * FROM TrialSite WHERE TrialSite = '" & sSite & "'"
    ElseIf sStudy <> ALL_STUDIES And sSite = ALL_SITES Then
        rsTrialSite.Filter = "TrialSite = '" & sSite & "'"
        'sSQL = "SELECT * FROM TrialSite WHERE ClinicalTrialID = " & lTrialId
    Else
        'no filter
        'sSQL = "SELECT * FROM TrialSite"
    End If

    'Set rsTrialSite = New ADODB.Recordset
    'rsTrialSite.Open sSQL, mconMACRO, adOpenKeyset, adLockOptimistic, adCmdText
    
    If rsTrialSite.RecordCount <= 0 Then
        'specific study and specific site
        If sStudy <> ALL_STUDIES And sSite <> ALL_SITES Then
             CheckForTrialSite = "Site " & sSite & " is not participating in study " & sStudy & "."
        'allsites allstudies
        ElseIf sStudy = ALL_STUDIES And sSite = ALL_SITES Then
             CheckForTrialSite = "No sites are participating in any studies."
        'allsites and specific study
        ElseIf sStudy <> ALL_STUDIES And sSite = ALL_SITES Then
             CheckForTrialSite = "No sites are participating study " & sStudy & "."
        'allstudies and specific site
        ElseIf sStudy = ALL_STUDIES And sSite <> ALL_SITES Then
             CheckForTrialSite = "Site " & sSite & " is not participating in any study."
        End If
    End If

    'Set rsTrialSite = Nothing

    'Remove Filter
    rsTrialSite.Filter = adFilterNone
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserRolesInfo.CheckForInActiveSite"
End Function

'-------------------------------------------------------------------------------------
Private Function CheckForInActiveSite(ByVal sSite As String, colSites As Collection) As String
'-------------------------------------------------------------------------------------
'checks if sites are inactive and returns message to be displayed in the tooltip
'colSites is a collection of site stauses keyed by site
'TA 20/1/2006: use in memory colection rather than database
'-------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsSites As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    CheckForInActiveSite = ""
    
'    sSQL = "SELECT SiteStatus FROM Site WHERE Site = '" & sSite & "'"
'    Set rsSites = New ADODB.Recordset
'    rsSites.Open sSQL, mconMACRO, adOpenKeyset, adLockOptimistic, adCmdText
'    If rsSites.RecordCount <= 0 Then
    If Not CollectionMember(colSites, sSite, False) Then
        If sSite <> "AllSites" Then
            CheckForInActiveSite = "Site " & sSite & " does not exist"
            Exit Function
        End If
        Exit Function
    End If
    
    'rsSites.MoveFirst
    'If rsSites!SiteStatus = 1 Then
    If CLng(colSites(sSite)) = 1 Then
        CheckForInActiveSite = "Site " & sSite & " is inactive"
    End If
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserRolesInfo.CheckForInActiveSite"
End Function

'-------------------------------------------------------------------------------------
Private Function CheckForUserStatus(ByVal sUser As String, _
                                    ByVal sRoleCode As String, _
                                    lMaxPasswordRetries As Long, _
                                    colRoles As Collection, rsUser As ADODB.Recordset) As String
'-------------------------------------------------------------------------------------
'checks user statuses and returns message to be displayed in the tooltip
'-------------------------------------------------------------------------------------
Dim sSQL As String
'Dim rsUser As ADODB.Recordset
'Dim rsPasswords As ADODB.Recordset
Dim rsRole As ADODB.Recordset

    On Error GoTo ErrHandler
    
    CheckForUserStatus = ""
    
'    sSQL = " Select * FROM MACROUser WHERE UserName = '" & sUser & "'"
'    Set rsUSer = New ADODB.Recordset
'    rsUSer.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
    If rsUser.RecordCount > 0 Then
        If rsUser!Enabled = 0 Then
            CheckForUserStatus = "User " & rsUser!UserName & " is disabled"
        End If
    End If
    
'    sSQL = " Select * FROM Role WHERE RoleCode = '" & sRoleCode & "'"
'    Set rsRole = New ADODB.Recordset
'    rsRole.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
'    If rsRole.RecordCount <= 0 Then
    If Not CollectionMember(colRoles, sRoleCode, False) Then
        CheckForUserStatus = "Role definition not found"
    Else
        'If rsRole!Enabled = 0 Then
        If CLng(colRoles(sRoleCode)) = 0 Then
            CheckForUserStatus = "Role " & sRoleCode & " is disabled"
        End If
    End If

'    sSQL = " Select PasswordRetries FROM MACROPassword"
'    Set rsPasswords = New ADODB.Recordset
'    rsPasswords.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
'    If rsPasswords.RecordCount > 0 Then
        If rsUser!FailedAttempts = lMaxPasswordRetries Then
            CheckForUserStatus = "User " & rsUser!UserName & " is locked out"
        End If
'    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserRolesInfo.CheckForUserStatus"
End Function

'---------------------------------------------------------------------------------
Private Function GetClinicalTrialID(ByVal sClinicalTrialName As String, _
                                    Optional ByVal sDescription As String) As Long
'---------------------------------------------------------------------------------
'receives study name and or description and returns clinicaltrial ID
'---------------------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT ClinicalTrialId FROM ClinicalTrial " _
        & " WHERE ClinicalTrialName = '" & sClinicalTrialName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, mconMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        GetClinicalTrialID = -1
    Else
        GetClinicalTrialID = rsTemp!ClinicalTrialId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
        
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserRolesInfo.GetClinicalTrialID"
End Function

'---------------------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------------------

'---------------------------------------------------------------------------------

    lvwUserRolesInfo.Height = Me.Height
    lvwUserRolesInfo.Width = Me.Width

End Sub

'----------------------------------------------------------------------------------------
Private Sub UserRoleParamaters()
'----------------------------------------------------------------------------------------
Dim var As Variant

    If lvwUserRolesInfo.ListItems.Count <> 0 Then
        'split key to get parameters to pass to event
        var = Split(lvwUserRolesInfo.SelectedItem.Key, "|")
        msDatabase = var(0)
        msUserName = var(1)
        msRoleCode = var(2)
        msSiteCode = var(3)
        msStudyCode = var(4)
    
        RaiseEvent SelectedItem(eSMNodeTag.DatabaseTag, msDatabase, 0, msStudyCode, msSiteCode, msUserName, msRoleCode)
    End If
End Sub

'----------------------------------------------------------------------------------------
Private Sub lvwUserRolesInfo_Click()
'----------------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------

    Call UserRoleParamaters
    Call frmMenu.SelectedItemParameters(eSMNodeTag.DatabaseTag, msDatabase, 0, msStudyCode, msSiteCode, msUserName, msRoleCode)

End Sub

'------------------------------------------------------------------------------------------
Private Sub lvwUserRolesInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'------------------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------------------
On Error GoTo ErrHandler
    
    Call lvw_Sort(lvwUserRolesInfo, ColumnHeader)

Exit Sub:
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|lvwUserRolesInfo_ColumnClick"
End Sub

'---------------------------------------------------------------------------------------
Private Sub lvwUserRolesInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
Dim i As Integer
Dim Item As MSComctlLib.ListItem
    
    Set Item = lvwUserRolesInfo.HitTest(X, Y)
    
    If Not Item Is Nothing Then
        'check if right button pressed
        If Button = vbRightButton Then
            Item.Selected = True
            Call UserRoleParamaters
            'call routine to display form
            Call DisplayPopUpMenu
        End If
    
    End If

End Sub

'-------------------------------------------------------------------------------
Private Sub DisplayPopUpMenu()
'-------------------------------------------------------------------------------
'Display the pop up menu with the correct menu items displayed and enabled
'-------------------------------------------------------------------------------
Dim sMenuItemSelected As String
Dim oMenuItems As clsMenuItems
Dim sRoleToDelete As String
Dim sMsg As String

    On Error GoTo ErrHandler
    
    Set oMenuItems = New clsMenuItems
    Call oMenuItems.Add("DELUSRROLE", "Delete User Role...", goUser.CheckPermission(gsFnMaintRole))
    Call oMenuItems.Add("EDITUSRROLE", "Edit User Role...", goUser.CheckPermission(gsFnMaintRole))
    Call oMenuItems.Add("GOTOROLE", "Go To Role...", goUser.CheckPermission(gsFnMaintRole))
    Call oMenuItems.Add("GOTOUSR", "Go To User...", goUser.CheckPermission(gsFnMaintRole))
    sMenuItemSelected = frmMenu.ShowPopUpMenu(oMenuItems)
    
    Select Case sMenuItemSelected
        Case "DELUSRROLE"
            sRoleToDelete = msStudyCode & "|" & msSiteCode & "|" & msRoleCode & "|" & msUserName
            sMsg = "Are you sure you want to delete the selected user role?"
            If DialogQuestion(sMsg) = vbYes Then
                Call frmNewUserRole.DeleteRoles(sRoleToDelete, msDatabase, True)
                Call frmMenu.RefereshUserRoleInfoForm
            End If
            
        Case "EDITUSRROLE"
            Call frmMenu.UserRoleForm(False, msUserName, msDatabase, msRoleCode, msStudyCode, msSiteCode)
            Call frmMenu.RefereshUserRoleInfoForm
        
        Case "GOTOROLE"
            Call frmMenu.RoleManagementForm(msRoleCode)
            Call frmMenu.RefereshDatabaseInfoForm
            
        Case "GOTOUSR"
            Call frmMenu.NewUserForm(False, msUserName)
    End Select
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserRolesInfo.DisplayPopUpMenu"
End Sub

'-------------------------------------------------------------------------------
Private Function GetCollectionFromSQL(conDB As ADODB.Connection, sSQL As String) As Collection
'-------------------------------------------------------------------------------
'return collection of strings from the sql
'first column is the key, second is the value
'-------------------------------------------------------------------------------
Dim col As Collection
Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Set col = New Collection
    Set rs = New ADODB.Recordset
    rs.Open sSQL, conDB, adOpenKeyset, adLockOptimistic, adCmdText
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            col.Add CStr(rs.Fields(1).Value), CStr(rs.Fields(0).Value)
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Set GetCollectionFromSQL = col
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserRolesInfo.GetCollectionFromSQL"
End Function

'-------------------------------------------------------------------------------
Private Function GetDisconnectedRecordset(conDB As ADODB.Connection, sSQL As String) As ADODB.Recordset
'-------------------------------------------------------------------------------
'returns a recordset that can be navigated even after conncetion is closed
'-------------------------------------------------------------------------------
Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler

    Set rs = New ADODB.Recordset
    rs.Open sSQL, conDB, adOpenKeyset, adLockOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GetDisconnectedRecordset = rs
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserRolesInfo.GetDisconnectedRecordset"
End Function

