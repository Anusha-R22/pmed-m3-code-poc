VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewUserRole 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Role"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraRoles 
      Caption         =   "User R&oles"
      Height          =   2820
      Left            =   45
      TabIndex        =   13
      Top             =   1800
      Width           =   8490
      Begin MSComctlLib.ListView lvwUserRoles 
         Height          =   2460
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   4339
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "Remove A&ll"
         Height          =   345
         Left            =   7260
         TabIndex        =   16
         Top             =   660
         Width           =   1125
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "R&emove"
         Height          =   345
         Left            =   7260
         TabIndex        =   15
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame fraUserRoles 
      Caption         =   "&New User Role"
      Height          =   915
      Left            =   60
      TabIndex        =   5
      Top             =   780
      Width           =   8490
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   345
         Left            =   7260
         TabIndex        =   12
         Top             =   480
         Width           =   1125
      End
      Begin VB.ComboBox cboStudy 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   1755
      End
      Begin VB.ComboBox cboSite 
         Height          =   315
         Left            =   2535
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   1755
      End
      Begin VB.ComboBox cboUserRole 
         Height          =   315
         Left            =   4920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblSites 
         Caption         =   "S&ite"
         Height          =   195
         Left            =   2535
         TabIndex        =   8
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblStudy 
         Caption         =   "&Study"
         Height          =   200
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblUserRole 
         Caption         =   "&Role"
         Height          =   255
         Left            =   4950
         TabIndex        =   10
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "User and Database"
      Height          =   615
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   8490
      Begin VB.ComboBox cboUserDatabase 
         Height          =   315
         Left            =   4920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   210
         Width           =   1755
      End
      Begin VB.ComboBox cboUserName 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   1755
      End
      Begin VB.Label lblUserDatabase 
         Caption         =   "&Database"
         Height          =   240
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblUserName 
         Caption         =   "&User Name"
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   7320
      TabIndex        =   17
      Top             =   4680
      Width           =   1125
   End
   Begin VB.Label lblDist 
      Caption         =   "Creating messages for site distribution.  Please wait........"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4740
      Visible         =   0   'False
      Width           =   6075
   End
End
Attribute VB_Name = "frmNewUserRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2002-2004. All Rights Reserved
'   File:       frmNewUserRole.frm
'   Author:     Ashitei Trebi-Ollennu, September 2002
'   Purpose:    Updates user roles
'------------------------------------------------------------------------------
' REVISIONS
' NCJ 29 Apr 04 - Bug 2254 - Ensure timestamps are converted to standard in UpdateUserRole
'------------------------------------------------------------------------------

Option Explicit

Private mcolCode As Collection
Private mcolSite As Collection
Private mcolRoles As Collection
Private mbExistingUserRoles As Boolean 'used to check for existence of  user roles
Private mconMACRO As ADODB.Connection
Private mbChanged As Boolean
Private mbNew As Boolean
Private msUserName As String
Private msDatabaseCode As String
Private msRoleCode As String
Private msStudyCode As String
Private msSiteCode As String
Private mbDatabaseError As Boolean

'------------------------------------------------------------------------------------
Public Sub Display(bNew As Boolean, sUsername As String, sDatabaseCode As String, _
                   sRoleCode As String, sStudyCode As String, sSiteCode As String)
'------------------------------------------------------------------------------------

    Me.Icon = frmMenu.Icon
    
    mbNew = bNew
    msUserName = sUsername
    msDatabaseCode = sDatabaseCode
    msRoleCode = sRoleCode
    msStudyCode = sStudyCode
    msSiteCode = sSiteCode
    
    mbChanged = False
    
    ShowColumnHeaders
    
    RefreshUsers
    RefreshDatabases
    
    cmdAdd.Enabled = False
    
    FormCentre Me
    
    If (cboUserDatabase.Text <> "") And (cboUserName.Text <> "") Then
        EnableAddButton
    End If
        
    Me.Show vbModal

End Sub

'-------------------------------------------------------
Private Sub cboSite_Change()
'-------------------------------------------------------
'
'-------------------------------------------------------
    
    EnableAddButton

End Sub

'--------------------------------------------------------
Private Sub cboSite_Click()
'--------------------------------------------------------
'
'--------------------------------------------------------
    
    EnableAddButton

End Sub

'--------------------------------------------------------
Private Sub cboStudy_Change()
'--------------------------------------------------------
'
'--------------------------------------------------------
    
    EnableAddButton

End Sub

'--------------------------------------------------------
Private Sub cboStudy_Click()
'--------------------------------------------------------
'
'--------------------------------------------------------
    
    EnableAddButton

End Sub

'-------------------------------------------------------
Private Sub cboUserDatabase_Change()
'-------------------------------------------------------
'
'-------------------------------------------------------

    EnableAddButton

End Sub

'-------------------------------------------------------
Private Sub cboUserDatabase_Click()
'-------------------------------------------------------
'
'-------------------------------------------------------
        
       
    cboSite.Clear
    cboStudy.Clear
    cboUserRole.Clear
    lvwUserRoles.ListItems.Clear
    ShowColumnHeaders
    OpenSelectedDatabase (cboUserDatabase.Text)
    If cboUserDatabase.Text <> "All Databases" Then
        LoadExistingRoles
    End If
    EnableAddButton

End Sub

'-----------------------------------------------------------
Private Sub cboUserName_Change()
'-----------------------------------------------------------
'
'-----------------------------------------------------------
    msUserName = cboUserName.Text
    EnableAddButton

End Sub

'----------------------------------------------------------
Private Sub cboUserName_Click()
'----------------------------------------------------------
'
'----------------------------------------------------------
    
    If cboUserDatabase.Text <> "All Databases" Then
        msUserName = cboUserName.Text
        LoadExistingRoles
    Else
        DisplayUserRolesForAllDatabases
    End If
    
    EnableAddButton
    
End Sub

'---------------------------------------------------------
Private Sub cboUserRole_Change()
'---------------------------------------------------------
'
'---------------------------------------------------------
    
    EnableAddButton

End Sub

'-------------------------------------------------------------
Private Sub cboUserRole_Click()
'-------------------------------------------------------------
'
'-------------------------------------------------------------
    
    EnableAddButton

End Sub

'--------------------------------------------------------------
Private Sub cmdADD_Click()
'--------------------------------------------------------------
'
'--------------------------------------------------------------
    
    On Error GoTo Errlabel
    
    mbChanged = True
    
    Screen.MousePointer = vbHourglass
    

    
    If CheckForConflicts(cboUserName.Text, cboUserDatabase.Text, _
                    cboStudy.Text, cboSite.Text, cboUserRole.Text) Then
        
        lblDist.Visible = True
        Me.Refresh
        Call UpdateUserRole(cboUserName.Text, _
            cboStudy.Text, cboSite.Text, cboUserRole.Text, cboUserDatabase.Text)
        
        Call UpdateUserDatabases(cboUserName.Text, cboUserDatabase.Text)
        
        lblDist.Visible = False
    End If
  

    
    Screen.MousePointer = vbDefault
    
Exit Sub
Errlabel:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdADD_Click|frmNewUserRole")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
End Sub

'-----------------------------------------------------------
Private Sub cmdClose_Click()
'-----------------------------------------------------------

'-----------------------------------------------------------
    If mbChanged Then
        Call frmMenu.RefereshTreeView
    End If
    
    Unload Me

End Sub

'-----------------------------------------------------------
Private Sub cmdRemove_Click()
'-----------------------------------------------------------
'delete selected items from the database and listview
'-----------------------------------------------------------
Dim i As Integer

    On Error GoTo Errlabel
    
    mbChanged = True
    
    Screen.MousePointer = vbHourglass
    
    lblDist.Visible = True
    Me.Refresh
    
    For i = lvwUserRoles.ListItems.Count To 1 Step -1
        If lvwUserRoles.ListItems(i).Selected Then
            Call DeleteRoles(lvwUserRoles.ListItems(i).Key, cboUserDatabase.Text)
            lvwUserRoles.ListItems.Remove (i)
        End If
    Next
    
    'if there are no users role then remove from UserDatabase table
    If lvwUserRoles.ListItems.Count = 0 Then
        Call DeleteUserDatabase(cboUserName.Text, cboUserDatabase.Text)
    End If
    
    lblDist.Visible = False
    
    Screen.MousePointer = vbDefault
    
Exit Sub
Errlabel:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdRemove_Click|frmNewUserRole")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
End Sub

'--------------------------------------------------------------
Private Sub cmdRemoveAll_Click()
'--------------------------------------------------------------
'delete all items from the database and listview
'--------------------------------------------------------------
Dim i As Integer
    
    On Error GoTo Errlabel
    
    mbChanged = True
    
    Screen.MousePointer = vbHourglass
    
    lblDist.Visible = True
    Me.Refresh
    
    For i = lvwUserRoles.ListItems.Count To 1 Step -1
        Call DeleteRoles(lvwUserRoles.ListItems(i).Key, cboUserDatabase.Text)
        lvwUserRoles.ListItems.Remove i
    Next
    
    'removing all user roles for selected user and database therefore remove from UserDatabase table
    Call DeleteUserDatabase(cboUserName.Text, cboUserDatabase.Text)
    
    lblDist.Visible = False
    
    Screen.MousePointer = vbDefault
    

Exit Sub
Errlabel:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdRemoveAll_Click|frmNewUserRole")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
End Sub

'----------------------------------------------------------------
Private Sub RefreshUsers()
'----------------------------------------------------------------
' Add the list of Users to the User Name combo box
'----------------------------------------------------------------
Dim rsUsers As ADODB.Recordset
Dim i As Integer
Dim sUsername As String

    On Error GoTo ErrHandler
    
    cboUserName.Clear
    Set mcolCode = New Collection
    
    Set rsUsers = gdsUserList
    With rsUsers
        Do Until .EOF = True
            ' once the collection is instantiated add the members to it
            ' collection .ADD fields(1) are the names from the recordset and
            ' fields(0) is the usercode which is used as a key to write the data
            ' to the relevant textbox .
            cboUserName.AddItem .Fields(0).Value
            mcolCode.Add .Fields(1).Value, .Fields(0).Value
            .MoveNext
        Loop
    End With
    
    If (mbNew = False) And (msUserName <> "") Then
        For i = 0 To cboUserName.ListCount - 1
            If cboUserName.List(i) = msUserName Then
                cboUserName.ListIndex = i
                Exit For
            End If
        Next
    ElseIf (mbNew = False) And (msDatabaseCode <> "") Then
   
        cboUserName.ListIndex = 0
    End If
   
   rsUsers.Close
   Set rsUsers = Nothing
       
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmNewUserRole.RefreshDatabases"
End Sub

'--------------------------------------------------------------
Private Sub RefreshDatabases()
'--------------------------------------------------------------
' Add the databases to their listview
'--------------------------------------------------------------
Dim rsDatabases As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrHandler
    
    cboUserDatabase.Clear
    cboUserDatabase.AddItem "All Databases"
    Set rsDatabases = gdsDatabaseList
    With rsDatabases
        Do Until .EOF = True
            cboUserDatabase.AddItem .Fields(0).Value
            .MoveNext
        Loop
    End With
    
    If (mbNew = False) And (msDatabaseCode <> "") Then
        For i = 0 To cboUserDatabase.ListCount - 1
            If cboUserDatabase.List(i) = msDatabaseCode Then
                cboUserDatabase.ListIndex = i
                Exit For
            End If
        Next
    ElseIf (mbNew = False) And (msUserName <> "") Then
   
        cboUserDatabase.ListIndex = 0
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmNewUserRole.RefreshDatabases"
End Sub
  
'----------------------------------------------------------------
Private Sub RefreshRoles()
'----------------------------------------------------------------
 ' Add the roles to their combo box
'----------------------------------------------------------------
Dim rsRoles As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrHandler

    cboUserRole.Clear
    Set rsRoles = gdsRoleList
    
    'only do if there are user roles to add, may be none if all Roles are sys admin roles and user is not a sysadmin
    If rsRoles.RecordCount > 0 Then
        With rsRoles
           Do Until .EOF = True
             cboUserRole.AddItem .Fields(0).Value
            .MoveNext
           Loop
        End With
    
        If (mbNew = False) And (msRoleCode <> "") Then
            For i = 0 To cboUserRole.ListCount - 1
                If cboUserRole.List(i) = msRoleCode Then
                    cboUserRole.ListIndex = i
                    Exit For
                End If
            Next
        Else
            cboUserRole.ListIndex = 0
        End If
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmNewUserRole.RefreshRoles"
End Sub

'---------------------------------------------------------------
Public Sub LoadAllStudies()
'---------------------------------------------------------------
'Loads all studies for a selected database
'---------------------------------------------------------------
Dim rsTrialList As ADODB.Recordset
Dim sSQL As String
Dim i As Integer

    On Error GoTo ErrHandler

    sSQL = "SELECT ClinicalTrialName FROM ClinicalTrial WHERE ClinicalTrialId > 0 ORDER BY ClinicalTrialName"
    Set rsTrialList = New ADODB.Recordset
    rsTrialList.Open sSQL, mconMACRO, adOpenForwardOnly, , adCmdText
    
    cboStudy.Clear
    'empty database so add ALLSITES
    If rsTrialList.RecordCount <= 0 Then
        cboStudy.AddItem ALL_STUDIES
    Else
        cboStudy.AddItem ALL_STUDIES
        rsTrialList.MoveFirst
        Do Until rsTrialList.EOF
            cboStudy.AddItem rsTrialList!ClinicalTrialName
            rsTrialList.MoveNext
        Loop
    End If

    If (mbNew = False) And (msStudyCode <> "") Then
        For i = 0 To cboStudy.ListCount - 1
            If cboStudy.List(i) = msStudyCode Then
                cboStudy.ListIndex = i
                Exit For
            End If
        Next
    Else
        cboStudy.ListIndex = 0
        'cboUserRole.ListIndex = 0
    End If


    Set rsTrialList = Nothing
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmNewUserRole.LoadAllStudies"
End Sub

'-----------------------------------------------------------
Public Sub LoadAllSites()
'-----------------------------------------------------------
'Loads all sites for a selected study
'-----------------------------------------------------------
Dim rsSiteList As ADODB.Recordset
Dim sSQL As String
Dim i As Integer
Dim bServer As Boolean

    On Error GoTo ErrHandler
    
    Set mcolSite = New Collection
    'get TrialID from TrialName
    sSQL = "SELECT DISTINCT Site FROM Site ORDER BY Site"
    Set rsSiteList = New ADODB.Recordset
    rsSiteList.Open sSQL, mconMACRO, adOpenForwardOnly, , adCmdText

    cboSite.Clear
    
    bServer = (GetMacroDBSetting("datatransfer", "dbtype", mconMACRO, gsSERVER) = gsSERVER)
    
    'empty database so add ALLSITES
    If (rsSiteList.RecordCount <= 0) Then
        If (bServer = True) Then
            cboSite.AddItem ALL_SITES
        End If
    Else
        If bServer = True Then
            'put all sites into collection
            cboSite.AddItem ALL_SITES
        End If
        'add all other sites
        rsSiteList.MoveFirst
        Do Until rsSiteList.EOF = True
            cboSite.AddItem rsSiteList!Site
            rsSiteList.MoveNext
        Loop
    End If
    
    If (mbNew = False) And (msSiteCode <> "") Then
        For i = 0 To cboSite.ListCount - 1
            If cboSite.List(i) = msSiteCode Then
                cboSite.ListIndex = i
                Exit For
            End If
        Next
    Else
        If cboSite.ListCount <> 0 Then
            cboSite.ListIndex = 0
            'cboUserRole.ListIndex = 0
        End If
    End If
    
    Set rsSiteList = Nothing
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmNewUserRole.LoadAllSites"
End Sub

'----------------------------------------------------------------------------------
Private Sub OpenSelectedDatabase(sDatabase As String)
'----------------------------------------------------------------------------------
'Opens the selected database and loads all the studies and sites in the database
'----------------------------------------------------------------------------------
Dim sSQL As String
Dim rsVersion As ADODB.Recordset
Dim sVersion As String
Dim sSubVersion As String
Dim sMessage As String


    On Error GoTo ErrHandler
    
    If sDatabase = "All Databases" Then
        Call DisplayUserRolesForAllDatabases
        Exit Sub
    End If
    
    Set mconMACRO = CreateDBConnection(sDatabase)
    
    'checks if connection to the selected database failed
    If mbDatabaseError Then Exit Sub
    
    sSQL = "SELECT * FROM MACROControl"
    Set rsVersion = New ADODB.Recordset
    rsVersion.Open sSQL, mconMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    sVersion = rsVersion![MACROVersion]
    sSubVersion = rsVersion![BuildSubVersion]
    
    If (sVersion <> goUser.SecurityDBVersion) Or (sSubVersion <> goUser.SecurityDBSubVersion) Then
        sMessage = sDatabase & " database version is " & sVersion & "." & sSubVersion & " and must be upgraded to " & goUser.SecurityDBVersion & "." & goUser.SecurityDBSubVersion
        mbDatabaseError = True
        Call DialogError(sMessage)
        Exit Sub
    Else
        'load roles
        Call RefreshRoles
        'load all studies in the database
        Call LoadAllStudies
        'load all sites
        Call LoadAllSites
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmNewUserRole.OpenSelectedDatabase"
End Sub

'--------------------------------------------------------------------------------
Private Function CreateDBConnection(sDatabase As String, Optional bDisplayConError As Boolean = True) As ADODB.Connection
'--------------------------------------------------------------------------------
'REM 05/11/02
'create a coonection to a selected database
'--------------------------------------------------------------------------------
Dim oDatabase As MACROUserBS30.Database
Dim conMACRO As ADODB.Connection
Dim bLoad As Boolean
Dim sMessage As String
Dim sConnectionString As String

    On Error GoTo ErrHandler
    
    'initialise boolean variable
    mbDatabaseError = False

    Set oDatabase = New MACROUserBS30.Database
    
    bLoad = oDatabase.Load(SecurityADODBConnection, goUser.UserName, sDatabase, "", False, sMessage)
    
    sConnectionString = oDatabase.ConnectionString
    Set conMACRO = New ADODB.Connection
    conMACRO.Open sConnectionString
    conMACRO.CursorLocation = adUseClient
    
    Set CreateDBConnection = conMACRO

Exit Function
ErrHandler:
    'if the connection fails then return the error message
    sMessage = "Could not create connection to " & sDatabase & " becasue " & Err.Description
    If bDisplayConError Then
        Call DialogError(sMessage)
    End If
    'set boolean variable to show connection to database failed
    mbDatabaseError = True
End Function

'--------------------------------------------------------------------------------
Private Sub UpdateUserDatabases(sUsername As String, sDatabaseCode As String)
'--------------------------------------------------------------------------------
'REM 23/10/02
'Checks to see if the User/Database combination selected exists in the UserDatabase table
' if not create a new row
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsUserDB As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM UserDatabase" _
        & " WHERE UserName = '" & sUsername & "'" _
        & " AND DatabaseCode = '" & sDatabaseCode & "'"
    Set rsUserDB = New ADODB.Recordset
    rsUserDB.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    If rsUserDB.RecordCount = 0 Then
        sSQL = "INSERT INTO UserDatabase " _
            & " VALUES ('" & sUsername & "','" & sDatabaseCode & "')"
            SecurityADODBConnection.Execute sSQL, , adCmdText
    End If

Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    Err.Raise Err.Number, , Err.Description & "|frmNewUserRole.UpdateUserDatabases"
End Sub


'--------------------------------------------------------------------------------
Public Sub UpdateUserRole(ByVal sUsername As String, _
                                ByVal sStudyCode As String, _
                                ByVal sSiteCode As String, _
                                ByVal sRoleCode As String, _
                                ByVal sDatabaseCode As String)
'--------------------------------------------------------------------------------
' NCJ 29 Apr 04 - Bug 2254 - Ensure timestamps are converted to standard in MessageParameters
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim itmX As MSComctlLib.ListItem
Dim skey As String
Dim oSystemMessage As SysMessages
Dim sMessageParameters As String
Dim sUserNameFull As String
Dim sEncryptedPassword As String
Dim nEnabled As Integer
Dim sPasswordCreated As String
Dim lFailedAttempts As Long
Dim sLastLogin As String
Dim sFirstLogin As String
Dim vUserDetails As Variant
Dim nSysAdmin As Integer

    On Error GoTo ErrHandler
            
    If sRoleCode = "" Or sUsername = "" Then Exit Sub
    
    TransBegin
        If Not DoesUserRoleExist(sUsername, sStudyCode, sSiteCode, sRoleCode) Then
            
            sSQL = "INSERT INTO UserRole " _
            & " VALUES ('" & sUsername & "','" & sRoleCode & "','" & sStudyCode & "','" & sSiteCode & "',1)"
            mconMACRO.Execute sSQL, , adCmdText
            
            'log new user role
            Call goUser.gLog(goUser.UserName, gsNEW_USER_ROLE, "User " & sUsername & " was given the following user role: " & sStudyCode & ", " & sSiteCode & ", " & sRoleCode & " in database " & sDatabaseCode)
            
            'get the user details
            vUserDetails = GetUserDetails(sUsername)

            ' NCJ 29 Apr 04 - Bug 2254 - Ensure timestamps are converted to standard
            sUserNameFull = vUserDetails(1, 0)
            sEncryptedPassword = vUserDetails(2, 0)
            nEnabled = CInt(vUserDetails(3, 0))
            sLastLogin = LocalNumToStandard(vUserDetails(4, 0))
            sFirstLogin = LocalNumToStandard(vUserDetails(5, 0))
            lFailedAttempts = CLng(vUserDetails(7, 0))
            sPasswordCreated = LocalNumToStandard(vUserDetails(8, 0))
            nSysAdmin = vUserDetails(9, 0)
            
            'Write the UserDetails for the UserRole, treated as new user so send all parameters.
            If Not ExcludeUserRDE(sUsername) Then 'don't distribute rde's details
                
                Set oSystemMessage = New SysMessages
                sMessageParameters = sUsername & gsPARAMSEPARATOR & sUserNameFull & gsPARAMSEPARATOR & sEncryptedPassword & gsPARAMSEPARATOR & nEnabled & gsPARAMSEPARATOR & sLastLogin & gsPARAMSEPARATOR & sFirstLogin & gsPARAMSEPARATOR & lFailedAttempts & gsPARAMSEPARATOR & sPasswordCreated & gsPARAMSEPARATOR & nSysAdmin & gsPARAMSEPARATOR & eUserDetails.udEditUser
                Call oSystemMessage.AddNewSystemMessage(mconMACRO, ExchangeMessageType.User, goUser.UserName, sUsername, "User Details", sMessageParameters, "", sRoleCode)

                'write new UserRole message to Message table
                Set oSystemMessage = New SysMessages
                sMessageParameters = sUsername & gsPARAMSEPARATOR & sRoleCode & gsPARAMSEPARATOR & sStudyCode & gsPARAMSEPARATOR & sSiteCode & gsPARAMSEPARATOR & 1 & gsPARAMSEPARATOR & eUserRole.urAdd
                Call oSystemMessage.AddNewSystemMessage(mconMACRO, ExchangeMessageType.UserRole, goUser.UserName, sUsername, "New User Role", sMessageParameters, sSiteCode)
                Set oSystemMessage = Nothing
                
            End If
            
            'add to listview
            skey = sStudyCode & "|" & sSiteCode & "|" & sRoleCode & "|" & sUsername
            If Not AlreadyInListView(lvwUserRoles, skey) Then
                Set itmX = lvwUserRoles.ListItems.Add(, skey, sStudyCode)
                itmX.SubItems(1) = sSiteCode
                itmX.SubItems(2) = sRoleCode
                lvwUserRoles.Tag = skey
            End If
        End If
    TransCommit
    'Call lvw_SetAllColWidths(lvwUserRoles, LVSCW_AUTOSIZE_USEHEADER)

Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    Err.Raise Err.Number, , Err.Description & "|frmNewUserRole.UpdateUserRole"
End Sub

'----------------------------------------------------------------------------------
Private Function DoesUserRoleExist(ByVal sUsername As String, _
                                ByVal sStudy As String, _
                                ByVal sSite As String, _
                                ByVal sRoleCode As String) As Boolean
'----------------------------------------------------------------------------------
'checks to see if roles already exist in database to avoid dupicates
'----------------------------------------------------------------------------------
Dim sSQL As String
Dim rsUserRole As ADODB.Recordset

    On Error GoTo ErrHandler

    DoesUserRoleExist = False
    
    sSQL = "SELECT * FROM UserRole "
    sSQL = sSQL & " WHERE  UserName = '" & sUsername & "'"
    sSQL = sSQL & " AND  RoleCode = '" & sRoleCode & "'"
    sSQL = sSQL & " AND StudyCode = '" & sStudy & "'"
    sSQL = sSQL & " AND SiteCode = '" & sSite & "'"
    
    Set rsUserRole = New ADODB.Recordset
    rsUserRole.Open sSQL, mconMACRO, adOpenKeyset, adLockOptimistic, adCmdText
    
    If rsUserRole.RecordCount > 0 Then
        DoesUserRoleExist = True
    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmNewUserRole.DoesUserRoleExist"
End Function

'--------------------------------------------------------------------------------------
Private Function ShowExistingUserRoles(ByVal sUsername As String, _
                                ByVal sRoleCode As String, _
                                ByVal sStudyCode As String, _
                                ByVal sSiteCode As String) As Boolean
'--------------------------------------------------------------------------------------
'loads the existing roles for selected users.
'--------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsUserDetails As ADODB.Recordset
Dim itmX As MSComctlLib.ListItem
Dim skey As String
Dim i As Integer
    
    On Error GoTo ErrHandler
    
    'if all parameters not supplied then do nothing
    If sUsername = "" Or sRoleCode = "" Then Exit Function
    
    'initialise variable and function
    ShowExistingUserRoles = False
    mbExistingUserRoles = False
    
    sSQL = "SELECT * FROM UserRole WHERE UserName = '" & sUsername & "'"
    sSQL = sSQL & " AND RoleCode = '" & sRoleCode & "'"
    sSQL = sSQL & " AND StudyCode = '" & sStudyCode & "'"
    sSQL = sSQL & " AND SiteCode = '" & sSiteCode & "'"
    Set rsUserDetails = New ADODB.Recordset
    rsUserDetails.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'if no record exist then exit
    If rsUserDetails.RecordCount <= 0 Then Exit Function
    rsUserDetails.MoveFirst
        Do Until rsUserDetails.EOF
            skey = rsUserDetails!StudyCode & "|" & rsUserDetails!SiteCode & "|" & rsUserDetails!RoleCode & "|" & sUsername
            If Not AlreadyInListView(lvwUserRoles, skey) Then
                Set itmX = lvwUserRoles.ListItems.Add(, skey, rsUserDetails!StudyCode)
                itmX.SubItems(1) = rsUserDetails!SiteCode
                itmX.SubItems(2) = rsUserDetails!RoleCode
                lvwUserRoles.Tag = skey
            End If
            rsUserDetails.MoveNext
        Loop
    
    ShowExistingUserRoles = True
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmNewUserRole.ShowExistingUserRoles"
End Function

'----------------------------------------------------------------------------------
Private Function AlreadyInListView(oList As ListView, _
                                ByVal skey As String) As Boolean
'----------------------------------------------------------------------------------
'checks if the role to be added to the list box is already exists in it
'----------------------------------------------------------------------------------
Dim n As Integer

    On Error GoTo ErrHandler
    
    AlreadyInListView = False
    
    For n = 1 To oList.ListItems.Count
        If oList.ListItems(n).Key = skey Then
            AlreadyInListView = True
            Exit Function
        End If
    Next

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.AlreadyInListBox"
End Function

'------------------------------------------------------------------------------------
Private Sub ShowColumnHeaders(Optional bAddColumn As Boolean = False)
'------------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------------
Dim colmX As MSComctlLib.ColumnHeader
Dim newcolmX As MSComctlLib.ColumnHeader
Dim n As Integer
    
    If lvwUserRoles.ColumnHeaders.Count > 0 Then
        For n = lvwUserRoles.ColumnHeaders.Count To 1 Step -1
            lvwUserRoles.ColumnHeaders.Remove n
        Next
    End If
        
    If Not bAddColumn Then
        'add column headers with widths
        Set colmX = lvwUserRoles.ColumnHeaders.Add(, , "Study", 1780)
        Set colmX = lvwUserRoles.ColumnHeaders.Add(, , "Site", 1780)
        Set colmX = lvwUserRoles.ColumnHeaders.Add(, , "Role", 1780)
    Else
        'add column headers with widths
        Set colmX = lvwUserRoles.ColumnHeaders.Add(, , "Study", 1780)
        Set colmX = lvwUserRoles.ColumnHeaders.Add(, , "Site", 1780)
        Set colmX = lvwUserRoles.ColumnHeaders.Add(, , "Role", 1780)
        Set colmX = lvwUserRoles.ColumnHeaders.Add(, , "Database", 1780)
    End If

    'set view type
    lvwUserRoles.View = lvwReport
    'set initial sort to ascending on column 0 (Study)
    lvwUserRoles.SortKey = 0
    lvwUserRoles.SortOrder = lvwAscending

End Sub

'--------------------------------------------------------------------------------------
Private Function CheckForConflicts(ByVal sUsername As String, _
                                ByVal sDatabaseCode As String, _
                                ByVal sStudyCode As String, _
                                ByVal sSiteCode As String, _
                                ByVal sRoleCode As String) As Boolean
'--------------------------------------------------------------------------------------
'check for constraints in the TrialSite and User role tables
'--------------------------------------------------------------------------------------
Dim sMsg As String
Dim rsTrialSite As ADODB.Recordset
Dim rsTrialID As ADODB.Recordset
Dim sSQL As String
Dim lTrialId As Long
Dim i As Integer
Dim skey As String
Dim sCompare As String
Dim v As Variant
Dim scStudy As String
Dim scSite As String
Dim scRole As String
Dim sDeleteString As String
Dim sString As String
Dim nCount As Integer
Dim var As Variant
Dim sUser As String
Dim sUserStudy As String
Dim sUserSite As String
Dim sTrialString As String


    On Error GoTo ErrHandler
    
    CheckForConflicts = False
    
    sUser = Trim(cboUserRole.Text)
    sUserStudy = Trim(cboStudy.Text)
    sUserSite = Trim(cboSite.Text)
    sUsername = Trim(cboUserName.Text)
    
    'if any empty parameters passed then do nothing
    If sUsername = "" Or sStudyCode = "" Or sSiteCode = "" Or sRoleCode = "" Then Exit Function
    
    'get trialid
    If sStudyCode <> ALL_STUDIES Then
        sSQL = "SELECT ClinicalTrialID FROM ClinicalTrial WHERE ClinicalTrialName ='" & sStudyCode & "'"
        Set rsTrialID = New ADODB.Recordset
        rsTrialID.Open sSQL, mconMACRO, adOpenKeyset, adLockOptimistic, adCmdText
        'to take care of balnk databases
        If rsTrialID.RecordCount > 0 Then
            lTrialId = rsTrialID!ClinicalTrialId
        End If
    End If
    
    'use trial id(if available) to get trial sites
    'deals with Allsites and  Allstudies exceptions
    If sStudyCode <> ALL_STUDIES And sSiteCode <> ALL_SITES Then
        sSQL = "SELECT * FROM TrialSite WHERE ClinicalTrialID = " & lTrialId
        sSQL = sSQL & " AND TrialSite = '" & sSiteCode & "'"
    ElseIf sStudyCode = ALL_STUDIES And sSiteCode <> ALL_SITES Then
        sSQL = "SELECT * FROM TrialSite WHERE TrialSite = '" & sSiteCode & "'"
    ElseIf sStudyCode <> ALL_STUDIES And sSiteCode = ALL_SITES Then
        sSQL = "SELECT * FROM TrialSite WHERE ClinicalTrialID = " & lTrialId
    Else
        sSQL = "SELECT * FROM TrialSite"
    End If
    
    Set rsTrialSite = New ADODB.Recordset
    rsTrialSite.Open sSQL, mconMACRO, adOpenKeyset, adLockOptimistic, adCmdText
    
    'start checking for conflicts in trialsite table.
    'first conflict: specific site/study selection
    If rsTrialSite.RecordCount = 1 Then
        Rem do nothing
    End If
    
    'second conflict:if trial table is empty
    'third conflict:if allsites or allstudies selected warn if there are no rows
    If rsTrialSite.RecordCount <= 0 Then
        'specific study and specific site
        If sStudyCode <> ALL_STUDIES And sSiteCode <> ALL_SITES Then
             sTrialString = "site " & sSiteCode & " is not participating in study " & sStudyCode & "."
        'allsites allstudies
        ElseIf sStudyCode = ALL_STUDIES And sSiteCode = ALL_SITES Then
             sTrialString = "no sites are participating in any studies."
        'allsites and specific study
        ElseIf sStudyCode <> ALL_STUDIES And sSiteCode = ALL_SITES Then
             sTrialString = "no sites are participating study " & sStudyCode & "."
        'allstudies and specific site
        ElseIf sStudyCode = ALL_STUDIES And sSiteCode <> ALL_SITES Then
             sTrialString = "site " & sSiteCode & " is not participating in any study."
        End If
    End If
    
    'start checks for conflicts in the user role table
    For i = lvwUserRoles.ListItems.Count To 1 Step -1
        sCompare = sUserStudy & "|" & sUserSite & "|" & sUser
        v = Split(lvwUserRoles.ListItems(i).Key, "|")
        scStudy = v(0)
        scSite = v(1)
        scRole = v(2)
        skey = scStudy & "|" & scSite & "|" & scRole
        
            'check for specific site and specific study
        If sCompare = skey Then
            Rem do nothing
        
        'check for specific site and all studies
        'from combo to listview
        ElseIf sUserSite = scSite And scStudy = ALL_STUDIES And scRole = sUser Then
            sDeleteString = sDeleteString & vbCrLf & skey
        
        'check for specific study and all sites
        'from combo to listview
        ElseIf sUserStudy = scStudy And scSite = ALL_SITES And scRole = sUser Then
             sDeleteString = sDeleteString & vbCrLf & skey
        
        'check for allstudies and allsites.
        'from combo to listview
        ElseIf sUserStudy = ALL_STUDIES And sUserSite = ALL_SITES And scRole = sUser Then
             sDeleteString = sDeleteString & vbCrLf & skey
        
        'check for specific site and all studies
        'from listview to combo
        ElseIf scSite = sUserSite And scStudy = ALL_STUDIES And scRole = sUser Then
            sDeleteString = sDeleteString & vbCrLf & skey
        
        'check for specific study and all sites
        'from listview to combo
        ElseIf scStudy = sUserStudy And scSite = ALL_SITES And scRole = sUser Then
            sDeleteString = sDeleteString & vbCrLf & skey
        
        'check for AllSites and AllStudies
        'from listview to combo
        ElseIf scStudy = ALL_STUDIES And scSite = ALL_SITES And scRole = sUser Then
            sDeleteString = sDeleteString & vbCrLf & skey
        
        'check for allstudies and specific site
        'from combo to listview
        ElseIf sUserStudy = ALL_STUDIES And scSite = sUserSite And scRole = sUser Then
             sDeleteString = sDeleteString & vbCrLf & skey
        
        'check for allstudies and specific site
        'from listview to combo
        ElseIf scStudy = ALL_STUDIES And scSite = sUserSite And scRole = sUser Then
            sDeleteString = sDeleteString & vbCrLf & skey
        
        'check for specific study and allsites
        'from listview to combo
        ElseIf scStudy = sUserStudy And sUserSite = ALL_SITES And scRole = sUser Then
            sDeleteString = sDeleteString & vbCrLf & skey
        End If
    Next
    
    If sDeleteString <> "" Or sTrialString <> "" Then
        If frmUserRoleErrors.Display(sDeleteString, sTrialString) = True Then
            If sDeleteString <> "" Then
                sDeleteString = Right(sDeleteString, Len(sDeleteString) - Len(vbCrLf))
                var = Split(sDeleteString, vbCrLf)
                sDeleteString = Replace(sDeleteString, "|", "  ")
                For nCount = 0 To UBound(var)
                    'additional parameters needed for DeleteRoles function
                    sString = var(nCount) & "|" & sUsername
                    Call DeleteRoles(sString, sDatabaseCode)
                    lvwUserRoles.ListItems.Remove (sString)
                Next
            End If
        Else
            CheckForConflicts = False
            Exit Function
        End If
    End If
    CheckForConflicts = True
    Set rsTrialID = Nothing
    Set rsTrialSite = Nothing
    
Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.CheckForConflicts"
End Function

'---------------------------------------------------------------------------
Private Sub DeleteUserDatabase(sUsername As String, sDatabaseCode As String)
'---------------------------------------------------------------------------
'REM 23/10/02
'If all user roles are removed for a specific database then need to
'remove the user and database from the UserDatabase table
'---------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "DELETE FROM UserDatabase" _
        & " WHERE UserName = '" & sUsername & "'" _
        & " AND DatabaseCode = '" & sDatabaseCode & "'"
    SecurityADODBConnection.Execute sSQL, , adCmdText

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.DeleteUserDatabase"
End Sub


'---------------------------------------------------------------------------
Public Sub DeleteRoles(ByVal sRoleToDelete As String, sDatabaseCode As String, Optional bCreateCon As Boolean = False)
'---------------------------------------------------------------------------
'deletes roles from the database
'---------------------------------------------------------------------------
Dim sSQL As String
Dim sUsername As String
Dim sStudyCode As String
Dim sSiteCode As String
Dim sRoleCode As String
Dim v As Variant
Dim conLocalMACRO As ADODB.Connection
Dim oSystemMessage As SysMessages
Dim sMessageParameters As String
    
    On Error GoTo ErrHandler
    
    v = Split(sRoleToDelete, "|")
    sStudyCode = v(0)
    sSiteCode = v(1)
    sRoleCode = v(2)
    sUsername = v(3)
    
    sSQL = "DELETE UserRole WHERE UserName = '" & sUsername & "'"
    sSQL = sSQL & " AND RoleCode = '" & sRoleCode & "'"
    sSQL = sSQL & " AND StudyCode = '" & sStudyCode & "'"
    sSQL = sSQL & " AND SiteCode = '" & sSiteCode & "'"
    
    'use modular level connection when in form
    If Not bCreateCon Then
        Set conLocalMACRO = mconMACRO
    Else 'but use newly created connection if delete is called from the UserRolesInfo form
        Set conLocalMACRO = New ADODB.Connection
        Set conLocalMACRO = CreateDBConnection(sDatabaseCode)
    End If
    
    'add system message before do delete
    If Not ExcludeUserRDE(sUsername) Then 'don't distribute rde messages
        Set oSystemMessage = New SysMessages
        sMessageParameters = sUsername & gsPARAMSEPARATOR & sRoleCode & gsPARAMSEPARATOR & sStudyCode & gsPARAMSEPARATOR & sSiteCode & gsPARAMSEPARATOR & 1 & gsPARAMSEPARATOR & eUserRole.urDelete
        Call oSystemMessage.AddNewSystemMessage(conLocalMACRO, ExchangeMessageType.UserRole, goUser.UserName, sUsername, "Delete User Role", sMessageParameters, sSiteCode)
        Set oSystemMessage = Nothing
    End If
    
    'Delete UserRole
    conLocalMACRO.Execute sSQL, , adCmdText
    
    'log deleted user role
    Call goUser.gLog(goUser.UserName, gsDEL_USER_ROLE, "User " & sUsername & " had the following user role removed: " & sStudyCode & ", " & sSiteCode & ", " & sRoleCode & " in database " & sDatabaseCode)
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.DeleteRoles"
End Sub

'----------------------------------------------------------------
Private Sub LoadExistingRoles()
'----------------------------------------------------------------
'
'----------------------------------------------------------------
Dim sSQL As String
Dim sUsername As String
Dim sDatabaseCode As String
Dim rsRoles As ADODB.Recordset
Dim skey As String
Dim sStudyCode As String
Dim sSiteCode As String
Dim sRoleCode As String
Dim itmX As MSComctlLib.ListItem


    On Error GoTo ErrHandler
    
    'checks if connection to the selected database failed
    If mbDatabaseError Then Exit Sub
    
    sUsername = Trim(cboUserName.Text)
    sDatabaseCode = Trim(cboUserDatabase.Text)
    sRoleCode = Trim(cboUserRole.Text)
    
    If sDatabaseCode = "" Then Exit Sub 'Or sDatabaseCode = "All Databases" Then Exit Sub
    
    sSQL = "SELECT * From UserRole WHERE UserName = '" & sUsername & "'"
    Set rsRoles = New ADODB.Recordset
    rsRoles.Open sSQL, mconMACRO, adOpenKeyset, adLockOptimistic, adCmdText
    
    lvwUserRoles.ListItems.Clear
    
    If rsRoles.RecordCount > 0 Then
    

         rsRoles.MoveFirst
        
         Do Until rsRoles.EOF
             'add to listview
             skey = rsRoles!StudyCode & "|" & rsRoles!SiteCode & "|" & rsRoles!RoleCode & "|" & rsRoles!UserName
             If Not AlreadyInListView(lvwUserRoles, skey) Then
                 Set itmX = lvwUserRoles.ListItems.Add(, skey, rsRoles!StudyCode)
                 itmX.SubItems(1) = rsRoles!SiteCode
                 itmX.SubItems(2) = rsRoles!RoleCode
                 lvwUserRoles.Tag = skey
             End If
             rsRoles.MoveNext
         Loop
    End If
    
    cmdRemove.Enabled = SysAdminEnableAddRemove(sUsername) 'True
    cmdRemoveAll.Enabled = SysAdminEnableAddRemove(sUsername) 'True
    
    cboSite.Enabled = True
    cboStudy.Enabled = True
    cboUserRole.Enabled = True
    
    'Call lvw_SetAllColWidths(lvwUserRoles, LVSCW_AUTOSIZE_USEHEADER)

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.LoadExistingRoles"
End Sub

'---------------------------------------------------------------
Private Sub EnableAddButton()
'---------------------------------------------------------------
'
'---------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If (cboUserDatabase.Text = "") Or (cboStudy.Text = "") Or (cboUserName.Text = "") Or (cboSite.Text = "") _
        Or (cboUserRole.Text = "") Or (SysAdminEnableAddRemove(msUserName) = False) Then
        cmdAdd.Enabled = False
    Else
        cmdAdd.Enabled = True
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.EnableAddButton"
End Sub

'---------------------------------------------------------------
Private Function SysAdminEnableAddRemove(sUsername As String) As Boolean
'---------------------------------------------------------------
'REM 24/01/03
'Checks to see if a user has sys admin property
'---------------------------------------------------------------
Dim rsUser As ADODB.Recordset
Dim sSQL As String
Dim bSysAdmin As Boolean

    On Error GoTo Errlabel
    
    sSQL = "SELECT SysAdmin FROM MACROUser "
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.sqlserver
    sSQL = sSQL & " WHERE UserName = '" & sUsername & "'"
    Case MACRODatabaseType.Oracle80
    sSQL = sSQL & "WHERE upper(UserName) = upper('" & sUsername & "')"
    End Select
    
    Set rsUser = New ADODB.Recordset
    rsUser.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    If rsUser.RecordCount <> 0 Then
        bSysAdmin = (rsUser!SysAdmin = 1)
    End If

'    If (bSysAdmin = True) And (goUser.SysAdmin = True) Then
'        SysAdminEnableAddRemove = True
'    ElseIf (bSysAdmin = flase) And (goUser.SysAdmin = True) Then
'        SysAdminEnableAddRemove = True
'    ElseIf (bSysAdmin = False) And (goUser.SysAdmin = False) Then
    
    If (bSysAdmin = True) And (goUser.SysAdmin = False) Then
        SysAdminEnableAddRemove = False
    Else
        SysAdminEnableAddRemove = True
    End If
    
Exit Function
Errlabel:

End Function

'---------------------------------------------------------------
Private Function GetUserDetails(sUsername As String) As Variant
'---------------------------------------------------------------
'REM 20/11/02
'Returns all the users details
' NB Numbers returned here will be in LOCAL format!! (e.g. FirstLogin, LastLogin timestamps)
'---------------------------------------------------------------
Dim sSQL As String
Dim rsUser As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM MACROUser" _
        & " WHERE UserName = '" & sUsername & "'"
    Set rsUser = New ADODB.Recordset
    rsUser.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText

    If rsUser.RecordCount > 0 Then
        GetUserDetails = rsUser.GetRows
    Else
        GetUserDetails = Null
    End If
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.GetUserDetails"
End Function

'------------------------------------------------------------------------------------
Private Sub DisplayUserRolesForAllDatabases()
'------------------------------------------------------------------------------------
'displays all/current user roles on all databases
'------------------------------------------------------------------------------------
Dim rsRoles As ADODB.Recordset
Dim sSQL As String
Dim n As Integer
Dim sUsername As String
Dim skey As String
Dim sDatabaseCode As String
Dim itmX As MSComctlLib.ListItem
Dim colmX As MSComctlLib.ColumnHeader
    
    On Error GoTo ErrHandler
    
    sUsername = Trim(cboUserName.Text)
    sDatabaseCode = Trim(cboUserDatabase.Text)
    
    If sUsername = "" Or sDatabaseCode = "" Then Exit Sub
    
    ShowColumnHeaders (True)
    lvwUserRoles.ListItems.Clear
    For n = 0 To cboUserDatabase.ListCount - 1
        If cboUserDatabase.List(n) <> "All Databases" Then
            Set mconMACRO = CreateDBConnection(Trim(cboUserDatabase.List(n)), False)
            'checks if connection to the selected database failed
                If Not mbDatabaseError Then
                    sSQL = "SELECT * From UserRole WHERE UserName = '" & sUsername & "'"
                    Set rsRoles = New ADODB.Recordset
                    rsRoles.Open sSQL, mconMACRO, adOpenKeyset, adLockOptimistic, adCmdText
    
                    If rsRoles.RecordCount > 0 Then
                        rsRoles.MoveFirst
                        Do Until rsRoles.EOF
                            'add to listview
                            skey = rsRoles!StudyCode & "|" & rsRoles!SiteCode & "|" & rsRoles!RoleCode & "|" & rsRoles!UserName & "|" & cboUserDatabase.List(n)
                            Set itmX = lvwUserRoles.ListItems.Add(, skey, rsRoles!StudyCode)
                            itmX.SubItems(1) = rsRoles!SiteCode
                            itmX.SubItems(2) = rsRoles!RoleCode
                            itmX.SubItems(3) = cboUserDatabase.List(n)
                            lvwUserRoles.Tag = skey
                            rsRoles.MoveNext
                        Loop
                    End If
                End If
        End If
    Next
    
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    cmdRemoveAll.Enabled = False
    cboStudy.Enabled = False
    cboSite.Enabled = False
    cboUserRole.Enabled = False
    
    'Call lvw_SetAllColWidths(lvwUserRoles, LVSCW_AUTOSIZE_USEHEADER)
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.DisplayUserRolesForAllDatabases"
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call RestartSystemIdleTimer
End Sub

