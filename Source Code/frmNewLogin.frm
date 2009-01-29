VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   5325
   ClientLeft      =   6030
   ClientTop       =   6015
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkForgot 
      Caption         =   "Connect   to server"
      Height          =   375
      Left            =   100
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame fraRole 
      Caption         =   "Please select a role"
      Height          =   1575
      Left            =   60
      TabIndex        =   9
      Top             =   3240
      Width           =   3575
      Begin MSComctlLib.ListView lvwRole 
         Height          =   1155
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Role"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fraUser 
      Height          =   1095
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   3575
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   1125
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   2345
      End
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1125
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   630
         Width           =   2345
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&User name"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   645
         Width           =   900
      End
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "Please select a database"
      Height          =   2055
      Left            =   60
      TabIndex        =   5
      Top             =   1140
      Width           =   3575
      Begin MSComctlLib.ListView lvwDatabase 
         Height          =   1695
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Database"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1260
      TabIndex        =   2
      Top             =   4920
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2520
      TabIndex        =   3
      Top             =   4920
      Width           =   1125
   End
End
Attribute VB_Name = "frmNewLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmLogin.frm
'   Author:     Richard Meinesz
'   Purpose:    Allows user to enter user name and password.  If conditional compilation
'   constant NTSecurity  = 1 then the user's NT login name is automatically used
'   as the MACRO user name.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   TA 26/11/2002: We store the last selected db in the settings file
'   TA 03/12/2002: Adjusted resizing so that the form looks OK in XP (base height of form on position of OK button)
'   REM 05/02/04 - Added message to into LoginLog saying user does not have permission to enter module
'   ic 07/12/2005 added active directory login
'   Mo 5/2/2008    gsEnteredPassword setup in cmdOK_Click for use of DoubleDataEntry module
'------------------------------------------------------------------------------------'
Option Explicit
Option Base 0
Option Compare Binary

Private mbLoginSucceeded As Boolean

Private msUserName As String
Private msSecurityCon As String
Private moUser As MACROUser

#If ASE = 1 Then
  Private AseResult As Long
  Private hAseReader As Long
#End If

Private mbCheckPasswordOnly As Boolean
Private Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)

Private msDB As String

'------------------------------------------------------------------------------------'
Public Function Display(sSecurityCon As String, _
                        Optional sDB As String = "", _
                        Optional bChangeCaption As Boolean = False) As MACROUser
'------------------------------------------------------------------------------------'
'REM 09/10/02
'TA 07/01/2003: If sDB if passed in this is the selected db
'ic 07/12/2005 added active directory login
'------------------------------------------------------------------------------------'
Dim sMessage As String


    msDB = sDB
    msSecurityCon = sSecurityCon
    
    Set moUser = New MACROUser

    FormCentre Me, frmMenu
    Me.Icon = frmMenu.Icon
    
    'Default to false
    mbLoginSucceeded = False
    
    ' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True
    Me.BackColor = glFormColour
    
    'ASH 3/2/2003 Change caption when using User Switch in DE
    If bChangeCaption Then
        cmdCancel.Caption = "E&xit"
    Else
        cmdCancel.Caption = "Cancel"
    End If
    
    mbCheckPasswordOnly = False
    
    
    If (gbIsActiveDirectoryLogin) Then
    
        Set Display = Nothing

        txtUserName.Enabled = False
        txtPassword.Enabled = False
        lblLabels(0).Enabled = False
        lblLabels(1).Enabled = False
    
    
        If (moUser.LoginAD(msSecurityCon, gsWindowsCurrentUser, "", DefaultHTMLLocation, GetApplicationTitle, _
            sMessage) = LoginResult.Success) Then
            
            msUserName = gsWindowsCurrentUser
            txtUserName.Text = msUserName
        
            ' only do the full login if required so pass new param mbCheckPasswordOnly
            If mbCheckPasswordOnly Then
                mbLoginSucceeded = True
                HourglassOff
                Unload Me
            Else
               ' Delay setting of mbLoginSucceeded until they've chosen a database and role
                Call RefreshUserDatabase
                
                If (lvwDatabase.ListItems.Count = 1) Or msDB <> "" Then
                    'if there is one db or a db was passed in
                    If (lvwRole.ListItems.Count = 1) Then
                        ' There's only one database and role code so log user in
                        'set-up database and role
                        If Not SetUpDatabaseAndRole Then
                            Call TidyUpAndHideWindow(False)
                            Exit Function
                        End If
                        mbLoginSucceeded = True
                        Unload Me
                    Else
                        Me.Show vbModal
                    End If
                Else
                    Me.Show vbModal
                End If
            End If
        Else
            Call DialogError("Active Directory login failed: " & vbCrLf & sMessage)
            Call TidyUpAndHideWindow(False)
        End If
    Else
    
        txtUserName.Text = vbNullString
        txtPassword.Text = vbNullString
        
        Call Resize(False)

        Me.Show vbModal
    End If
    
    'prevent ghosting
    Me.Refresh
    DoEvents
    
    If mbLoginSucceeded Then
        Set Display = moUser
    End If
End Function

'------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'------------------------------------------------------------------------------------'
'REM 05/02/04 - Added message to log into LoginLog saying user does not have permission to enter module
'ic 07/12/2005 added active directory login
'------------------------------------------------------------------------------------'
Dim nButtons As Integer
Dim nLogin As Integer
Dim sUserName As String
Dim sPassword As String
Dim sMessage As String
Dim sNewPassword As String
Dim colDatabases As Collection
Dim vDatabase As Variant
Dim bLoadDBAndRole As Boolean
Dim sErrMsg As String

    On Error GoTo Errorlabel
    
    gsEnteredPassword = txtPassword.Text

    If (gbIsActiveDirectoryLogin) Then
        mbLoginSucceeded = True
        'user is selecting a database and role
        If Not SetUpDatabaseAndRole Then
            Call TidyUpAndHideWindow(False)
            Exit Sub
        End If
        Unload Me
        Exit Sub
    End If
    
    HourglassOn
    
    bLoadDBAndRole = False
    
    If lvwDatabase.Visible Then
    
        'user is selecting a database and role
        If Not SetUpDatabaseAndRole Then
            Call TidyUpAndHideWindow(False)
            Exit Sub
        End If
        
       'check the user has access to the module
        If moUser.UserHasAccessToModule(GetApplicationTitle, sMessage) Then
            mbLoginSucceeded = True
            'TA 26/11/2002: store selected db
            SetMACROSetting MACRO_SETTING_LAST_USED_DATABASE, lvwDatabase.SelectedItem
            SetMACROSetting MACRO_SETTING_LAST_USED_ROLE, lvwRole.SelectedItem
            Unload Me
        Else
            Call DialogInformation(sMessage, gsDIALOG_TITLE)
            'REM 05/02/04 - Added message to into LoginLog saying user does not have permission to enter module
            Call moUser.gLog(moUser.UserName, "Login", "User did not have permission to access " & GetApplicationTitle)
            ' We must boot them out
            Call TidyUpAndHideWindow(False)
        End If

    Else
        'user is entering a username and password
        nButtons = vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal
        If txtUserName.Text = vbNullString Then
            MsgBox "Please enter a user name", nButtons, gsDIALOG_TITLE
            Call SetTextFocus
            
        ElseIf txtPassword.Text = vbNullString Then
            MsgBox "Please enter a password", nButtons, gsDIALOG_TITLE
            txtPassword.SetFocus
            
        Else
            'check its data management as cannot do this through any other module
            If App.Title = "MACRO_DM" Then
                'forgotten password on a site
                If chkForgot.Value = 1 Then
                    If DialogWarning("Please ensure you are connected to the server.", gsDIALOG_TITLE, True) = vbOK Then
                        sUserName = txtUserName.Text
                        
                        If Not SiteUserExists(msSecurityCon, sUserName) Then
                            'user does not exist
                            Call DialogError("User name " & sUserName & "does not exist", gsDIALOG_TITLE)
                            'set focus back to password text box and selected password
                            txtPassword.SetFocus
                            txtPassword.SelStart = 0
                            txtPassword.SelLength = Len(txtPassword.Text)
                            Exit Sub
                            
                        Else
                            Select Case frmMenu.ForgottenPassword(msSecurityCon, sUserName, txtPassword.Text, sErrMsg)
                            Case eDTForgottenPassword.pSuccess
                                'just continue with login, new password has been downloaded and written to the security database
                            Case eDTForgottenPassword.pIncorrectPassword
                                'user login failed
                                Call DialogError(sErrMsg, gsDIALOG_TITLE)
                                'set focus back to password text box and selected password
                                txtPassword.SetFocus
                                txtPassword.SelStart = 0
                                txtPassword.SelLength = Len(txtPassword.Text)
                                HourglassOff
                                Exit Sub
                            Case eDTForgottenPassword.pNoPassword
                                Call DialogInformation("There was nothing to download!  Please contact your system administrator.", gsDIALOG_TITLE)
                                'set focus back to password text box and selected password
                                txtPassword.SetFocus
                                txtPassword.SelStart = 0
                                txtPassword.SelLength = Len(txtPassword.Text)
                                HourglassOff
                                Exit Sub
                            Case eDTForgottenPassword.pNoDatabases
                                Call DialogInformation(sErrMsg, gsDIALOG_TITLE)
                                ' We must boot them out
                                Call TidyUpAndHideWindow(False)
                                MACROEnd
                            Case eDTForgottenPassword.pError
                                Call DialogError("Could not connect to server!" & vbCrLf & sErrMsg, gsDIALOG_TITLE)
                                'set focus back to password text box and selected password
                                txtPassword.SetFocus
                                txtPassword.SelStart = 0
                                txtPassword.SelLength = Len(txtPassword.Text)
                                HourglassOff
                                Exit Sub
                            End Select
                        End If
                    Else
                        HourglassOff
                        Exit Sub
                    End If
                End If
            End If
            
            sMessage = ""
            msUserName = txtUserName.Text
            Select Case moUser.Login(msSecurityCon, msUserName, txtPassword.Text, DefaultHTMLLocation, GetApplicationTitle, sMessage)
        
            Case LoginResult.Success

                    bLoadDBAndRole = True
                    
            Case LoginResult.ChangePassword
                'display message box asking user if they want to change their password,
                'if so call update password routine else
                'continue with login process
                If DialogQuestion(sMessage, gsDIALOG_TITLE) = vbYes Then
                    If frmPasswordChange.Display(moUser, msSecurityCon, sNewPassword, txtPassword.Text) Then
                        'write to LoginLog table
                        Call moUser.gLog(msUserName, gsLOGIN, "User changed password")
                        txtPassword.Text = sNewPassword
                        bLoadDBAndRole = True
                    Else
                        bLoadDBAndRole = False
                    End If
                Else
                    bLoadDBAndRole = True
                End If
                            
            Case LoginResult.PasswordExpired
                
                'call the update password routine
                Call DialogInformation(sMessage, gsDIALOG_TITLE)
                If frmPasswordChange.Display(moUser, msSecurityCon, sNewPassword, txtPassword.Text) Then
                    'Write to LoginLog table
                    Call moUser.gLog(msUserName, gsLOGIN, "User changed password")
                    txtPassword.Text = sNewPassword
                    bLoadDBAndRole = True
                Else
                    bLoadDBAndRole = False
                End If
            Case LoginResult.AccountDisabled
            
                'user account has been disabled
                Call DialogInformation(sMessage, gsDIALOG_TITLE)
                
                '   Replaced conditional compilation with global variable
'                If gnSecurityMode = SecurityMode.NTSeparatePassword Then
'                    txtPassword.SetFocus
'                ElseIf gnSecurityMode = SecurityMode.UsernamePassword Then
'                    Call SetTextFocus
'                End If
                bLoadDBAndRole = False

            Case LoginResult.Failed
            
                'user login failed
                Call DialogError(sMessage, gsDIALOG_TITLE)
                'set focus back to password text box and selected password
                txtPassword.SetFocus
                txtPassword.SelStart = 0
                txtPassword.SelLength = Len(txtPassword.Text)
                
                '   Replaced conditional compilation with global variable
'                If gnSecurityMode = SecurityMode.NTSeparatePassword Then
'                    txtPassword.SetFocus
'                ElseIf gnSecurityMode = SecurityMode.UsernamePassword Then
'                    Call SetTextFocus
'                End If
                bLoadDBAndRole = False
                
            End Select
            
        End If
        
        ' this should be set to false since normal login
        ' shall be a full login
        mbCheckPasswordOnly = False
    End If
    
    
    'if login doesn't fail then load up database and role code list boxes
    If bLoadDBAndRole Then
    
        ' only do the full login if required so pass new param mbCheckPasswordOnly
        If mbCheckPasswordOnly Then
            mbLoginSucceeded = True
            HourglassOff
            Unload Me
            
        Else

           ' Delay setting of mbLoginSucceeded until they've chosen a database and role
            Call RefreshUserDatabase
     
            
            If (lvwDatabase.ListItems.Count = 1) Or msDB <> "" Then
                'if there is one db or a db was passed in
                If (lvwRole.ListItems.Count = 1) Then
                    ' There's only one database and role code so log user in
                    'set-up database and role
                    If Not SetUpDatabaseAndRole Then
                        Call TidyUpAndHideWindow(False)
                        Exit Sub
                    End If
                    'check user has access to the module
                    If moUser.UserHasAccessToModule(GetApplicationTitle, sMessage) Then
                        mbLoginSucceeded = True
                    Else
                        Call DialogInformation(sMessage, gsDIALOG_TITLE)
                        ' We must boot them out
                        Call TidyUpAndHideWindow(False)
                    End If
                    Unload Me
                Else
                
                    Call Resize(True) 'and wait until they choose a database and role
                    
                End If
                
            Else

                Call Resize(True) 'and wait until they choose a database and role

            End If
        
        End If
    End If
    
    HourglassOff
        
Exit Sub
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdOK_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'------------------------------------------------------------------------------------'
Private Function SiteUserExists(sSecurity As String, ByRef sUserName As String) As Boolean
'------------------------------------------------------------------------------------'
'REM 20/01/03
'Check to see if a user exists, only used on a site (only used with Forgotten password function)
'------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsUser As ADODB.Recordset
Dim conSecurity As ADODB.Connection

    On Error GoTo ErrLabel

    Set conSecurity = New ADODB.Connection
    
    conSecurity.Open sSecurity
    conSecurity.CursorLocation = adUseClient

    'get UserName from the MacroUser table
    'check user name in Uppercase in Oracle as it is case sensitive
    Select Case Connection_Property(CONNECTION_PROVIDER, conSecurity)
    Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
        sSQL = "SELECT * FROM MACROUser WHERE upper(UserName) = upper('" & sUserName & "')"
    Case Else
        sSQL = "SELECT * FROM MACROUser WHERE UserName = '" & sUserName & "'"
    End Select

    'set up class level recordset for this user
    Set rsUser = New ADODB.Recordset
    rsUser.Open sSQL, conSecurity, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsUser.RecordCount = 0 Then
        SiteUserExists = False
    Else
        SiteUserExists = True
    End If
    
    Set conSecurity = Nothing
    
    rsUser.Close
    Set rsUser = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmLogin.SiteUserExists"
End Function

'------------------------------------------------------------------------------------'
Private Sub SetTextFocus()
'------------------------------------------------------------------------------------'
     
    On Error GoTo Errorlabel
    
    If txtUserName.Enabled Then
        txtUserName.SetFocus
    Else
        txtPassword.SetFocus
    End If
    
Exit Sub
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmLogin.SetTextFocus"
End Sub

'------------------------------------------------------------------------------------'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'------------------------------------------------------------------------------------'
'
'------------------------------------------------------------------------------------'

'    If KeyCode = vbKeyF1 Then               ' Show user guide
'
'        'REM 07/12/01 - New Call for MACRO Help
'        Call MACROHelp(Me.hWnd, App.Title)
'    End If

End Sub

'------------------------------------------------------------------------------------'
Private Sub Form_Unload(Cancel As Integer)
'------------------------------------------------------------------------------------'
'helps prevent ghosting
'------------------------------------------------------------------------------------'


End Sub

'------------------------------------------------------------------------------------'
Private Sub lvwDatabase_ItemClick(ByVal Item As MSComctlLib.ListItem)
'------------------------------------------------------------------------------------'
'Updates the User Role list view everytime a new database is selected, but if can't connect to a DB returns the error
'message to the list view
'------------------------------------------------------------------------------------'
Dim sMessage As String

    If Not moUser.SetCurrentDatabase(msUserName, Item, DefaultHTMLLocation, False, True, sMessage) Then
        lvwRole.ListItems.Clear
        lvwRole.ListItems.Add , , "Could not connect to the database because " & sMessage
        lvw_SetAllColWidths lvwRole, LVSCW_AUTOSIZE_USEHEADER
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
        Call RefreshUserRole
    End If
End Sub

'------------------------------------------------------------------------------------'
Private Sub txtPassword_GotFocus()
'------------------------------------------------------------------------------------'
'
'------------------------------------------------------------------------------------'

    On Error GoTo Errorlabel
    
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
       
Exit Sub
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmLogin.txtPassword_GotFocus"
End Sub

'------------------------------------------------------------------------------------'
Private Sub txtUserName_GotFocus()
'------------------------------------------------------------------------------------'
 
'------------------------------------------------------------------------------------'
    
    On Error GoTo Errorlabel
    
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
       
Exit Sub
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmLogin.txtPassword_GotFocus"
End Sub

'------------------------------------------------------------------------------------'
Private Sub RefreshUserDatabase()
'------------------------------------------------------------------------------------'
'REM 20/09/02
'Loads the list box with users databases, if only one database selects that one
'------------------------------------------------------------------------------------'
Dim vUserDatabase As Variant
Dim olistItem As ListItem

    On Error GoTo Errorlabel
    
    lvwDatabase.ListItems.Clear
    
    For Each vUserDatabase In moUser.UserDatabases
    
        lvwDatabase.ListItems.Add , vUserDatabase, vUserDatabase
    
    Next
    
    lvw_SetAllColWidths lvwDatabase, LVSCW_AUTOSIZE_USEHEADER
    
    If lvwDatabase.ListItems.Count = 0 Then
        DialogInformation "You do not have permission to access any MACRO databases"
        MACROEnd
    Else
        If msDB = "" Then
            'if no db passed in then use last logged in one
            Set olistItem = lvw_ListItembyText(lvwDatabase, GetMACROSetting(MACRO_SETTING_LAST_USED_DATABASE, ""), 0)
        Else
            Set olistItem = lvw_ListItembyText(lvwDatabase, msDB, 0)
        End If
        
        If olistItem Is Nothing Then
            lvwDatabase.ListItems(1).Selected = True
            lvwDatabase_ItemClick lvwDatabase.ListItems(1)
        Else
            olistItem.Selected = True
            lvwDatabase_ItemClick olistItem
        End If
        cmdOK.Enabled = True
    End If
       
Exit Sub
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmLogin.RefreshUserDatabase"
End Sub

'------------------------------------------------------------------------------------'
Private Sub RefreshUserRole()
'------------------------------------------------------------------------------------'
'REM 20/09/02
'Loads the list box with users Roles,
'------------------------------------------------------------------------------------'
Dim vUserRole As Variant
Dim olistItem As ListItem

    On Error GoTo Errorlabel

    lvwRole.ListItems.Clear
    
    For Each vUserRole In moUser.UserRoles
    
        lvwRole.ListItems.Add , vUserRole, vUserRole
    
    Next
    
    lvw_SetAllColWidths lvwRole, LVSCW_AUTOSIZE_USEHEADER
    
    If lvwRole.ListItems.Count = 0 Then
        DialogInformation "You do not have any MACRO User Roles"
        cmdOK.Enabled = False
    Else
        If CollectionMember(moUser.UserRoles, "Error", False) Then
            cmdOK.Enabled = False
        Else
            Set olistItem = lvw_ListItembyText(lvwRole, GetMACROSetting(MACRO_SETTING_LAST_USED_ROLE, ""), 0)
            If olistItem Is Nothing Then
                lvwRole.ListItems(1).Selected = True
            Else
                olistItem.Selected = True
            End If
            cmdOK.Enabled = True
        End If
    End If

Exit Sub
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmLogin.RefreshUserRole"
End Sub

'------------------------------------------------------------------------------------'
Private Function SetUpDatabaseAndRole() As Boolean
'------------------------------------------------------------------------------------'
'REM 20/09/02
'Loads the selected database properties, user role and associated permissions
'------------------------------------------------------------------------------------'
Dim sMessage As String

    If moUser.SetCurrentDatabase(msUserName, lvwDatabase.SelectedItem, DefaultHTMLLocation, True, True, sMessage) Then
        Call moUser.SetUserRole(lvwRole.SelectedItem)
        SetUpDatabaseAndRole = True
    Else
        Call DialogError("Unable to log into the MACRO database!" & vbCrLf & sMessage, GetApplicationTitle)
        SetUpDatabaseAndRole = False
    End If

End Function

'------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'------------------------------------------------------------------------------------'

    
    Call TidyUpAndHideWindow(False)

End Sub

'------------------------------------------------------------------------------------'
Public Sub TidyUpAndHideWindow(bLoginSucceeded As Boolean)
'------------------------------------------------------------------------------------'
    ' this should be set to false since normal login
    ' shall be a full login
    mbCheckPasswordOnly = False

    mbLoginSucceeded = bLoginSucceeded
    Unload Me

End Sub


''------------------------------------------------------------------------------------'
'Public Property Let CheckPasswordOnly(bCheckPasswordOnly As Boolean)
''------------------------------------------------------------------------------------'
'
'    On Error GoTo Errorlabel
'
'    mbCheckPasswordOnly = bCheckPasswordOnly
'    txtUserName.Enabled = Not mbCheckPasswordOnly
'    lblLabels(0).Enabled = Not mbCheckPasswordOnly
'
'Exit Property
'Errorlabel:
'    Err.Raise Err.Number, , Err.Description & "|" & "frmLogin.CheckPasswordOnly"
'End Property

'----------------------------------------------------------------------------------------'
Private Sub Resize(ByVal bShowDatabases As Boolean)
'----------------------------------------------------------------------------------------'
' TA 20/04/2000 - resize form and enable controls according to whether
'                   the user is entering password or choosing a database
'----------------------------------------------------------------------------------------'
Dim lTitleHeight As Long
Dim sCheck As String

    'calculate title bat height
    lTitleHeight = Me.Height - Me.ScaleHeight
    
    fraDatabase.Visible = bShowDatabases
    fraRole.Visible = bShowDatabases
    fraUser.Enabled = Not bShowDatabases

    If App.Title = "MACRO_DM" Then
        chkForgot.Visible = (GetMACROSetting("ForgottenPasswordCheckBox", "on") = "on")
    End If
    
    If bShowDatabases Then
        'Me.Height = 5820
        cmdCancel.Top = fraRole.Top + fraRole.Height + 60
        lvwDatabase.SetFocus
        chkForgot.Visible = False
    Else
        'Me.Height = 2040
        cmdCancel.Top = fraUser.Top + fraUser.Height + 60
        
    End If
    
    cmdOK.Top = cmdCancel.Top
    chkForgot.Top = cmdCancel.Top
    
    Me.Height = cmdOK.Top + cmdOK.Height + lTitleHeight + 60
    
End Sub




#If ASE = 1 Then
Private Function GetASEUserId() As String

 Const WRITE_READ_DATA_SIZE = 8
 Const KEY_SIZE = 8
 Const FILE_ID = 6
 Const FILE_SIZE = 8
 Const OFFSET = 0
 
Dim dwActiveProtocol As Long
Dim hAseCard As Long
ReDim chWriteBuffer(WRITE_READ_DATA_SIZE) As Byte
ReDim chreadbuffer(WRITE_READ_DATA_SIZE) As Byte
Dim MainKey(KEY_SIZE) As Byte
Dim chWriteKey(KEY_SIZE) As Byte
Dim WriteBufferTmp As String
Dim ReadBufferTmp As String
Dim wActualDataRead As Integer
Dim IO As ASEIO_T0
Dim i As Integer
Dim CardCaps As HLCARDCAPS
Dim FileProperties  As FileProperties
Dim no As Variant

Dim msASEUserId As String

     On Error GoTo ErrHandler
    
   '-----------------------------------------------------------------------------------------
    ' Initialize Main key to HFF's
    '-----------------------------------------------------------------------------------------
    For i = 0 To KEY_SIZE
        MainKey(i) = &HFF
    Next i
    
    '-----------------------------------------------------------------------------------------
    ' Initializing file properties (ID: 6, 8 bytes, no acces rules)
    '-----------------------------------------------------------------------------------------
    FileProperties.wID = FILE_ID
    FileProperties.wBytesAllocated = FILE_SIZE
    FileProperties.wWriteConditions = AC_NONE
    FileProperties.wReadConditions = AC_NONE
    
    
    '-----------------------------------------------------------------------------------------
    '   Open the first reader that is in the ASE database
    '-----------------------------------------------------------------------------------------
    'Debug.Print "Open the ASEDrive."
    DoEvents
    AseResult = ASEReaderOpenByNameNull(0, hAseReader)
    If (ReportResult(AseResult) = 0) Then
        End
    End If
    
    
    '-----------------------------------------------------------------------------------------
    '   Power the ISO7816 T=0 card
    '-----------------------------------------------------------------------------------------
    'Debug.Print "Power the ISO7816 T=0 card."
    DoEvents
    
    AseResult = ASECardPowerOn( _
                                    hAseReader, _
                                    MAIN_SOCKET, _
                                    CARD_POWER_UP, _
                                    PROTOCOL_CPU7816_T0, _
                                    dwActiveProtocol, _
                                    hAseCard)
    
    Do Until ReportResult(AseResult) <> 0
        AseResult = ASECardPowerOn( _
                                    hAseReader, _
                                    MAIN_SOCKET, _
                                    CARD_POWER_UP, _
                                    PROTOCOL_CPU7816_T0, _
                                    dwActiveProtocol, _
                                    hAseCard)
    Loop
        
    
    '-----------------------------------------------------------------------------------------
    '   Check if the card in the socket is a T=0 card
    '-----------------------------------------------------------------------------------------
    If (dwActiveProtocol <> PROTOCOL_CPU7816_T0) Then
        MsgBox ("Error: The card in the socket is not a T=0 card !!!")
        Call CloseReader
        End
    End If


    '-----------------------------------------------------------------------------------------
    '   Selecting Card Level
    '-----------------------------------------------------------------------------------------
    'Debug.Print "Selecting Card Level"
    DoEvents
    
    AseResult = ASEHLSelectCardLevel(hAseCard)
    If (ReportResult(AseResult) = 0) Then
       Call CloseReader
       End
    End If
    

    '-----------------------------------------------------------------------------------------
    '   Get card capabilities in order to retrieve information for further use.
    '-----------------------------------------------------------------------------------------
    'Debug.Print "Get card capabilities "
    DoEvents
    AseResult = ASEHLGetCardCaps(hAseCard, CardCaps)
    If (ReportResult(AseResult) = 0) Then
       Call CloseReader
       End
    End If
    
    '-----------------------------------------------------------------------------------------
    '   Creation of file 6
    '-----------------------------------------------------------------------------------------
    
'    Debug.Print "Creation of file 6"
'    DoEvents
'    AseResult = ASEHLCreateFile(hAseCard, MainKey(0), FileProperties)
'    If (ReportResult(AseResult) = 0) Then
'       Call CloseReader
'       End
'    End If
   
    '-----------------------------------------------------------------------------------------
    '   Open file number 6
    '-----------------------------------------------------------------------------------------
    'Debug.Print "Open file number 6."
    DoEvents
    AseResult = ASEHLOpenFile(hAseCard, FILE_ID)
    If (ReportResult(AseResult) = 0) Then
        Call CloseReader
        End
    End If
    
    '-----------------------------------------------------------------------------------------
    '   Write to file number 6
    '-----------------------------------------------------------------------------------------
'    Debug.Print "Write to file number 6: 01234567"
'    DoEvents
    
    '-----------------------------------------------------------------------------------------
    ' Set the buffer
    '-----------------------------------------------------------------------------------------
'    For I = 0 To WRITE_READ_DATA_SIZE - 1
'        chWriteBuffer(I) = I
'    Next I
'        chWriteBuffer(0) = Asc("a")
'        chWriteBuffer(1) = Asc("n")
'        chWriteBuffer(2) = Asc("d")
'        chWriteBuffer(3) = Asc("r")
'        chWriteBuffer(4) = Asc("e")
'        chWriteBuffer(5) = Asc("w")
'        chWriteBuffer(6) = Asc("n")
'        chWriteBuffer(7) = Asc(" ")
'
'    If (CardCaps.dwSecuredWriting = 1) Then
'        For I = 0 To KEY_SIZE
'            chWriteKey(I) = &HFF
'        Next I
''        For I = 0 To WRITE_READ_DATA_SIZE - 1
'            AseResult = ASEHLWrite(hAseCard, WRITE_READ_DATA_SIZE, OFFSET, chWriteBuffer(0), chWriteKey(0))
'            If (ReportResult(AseResult) = 0) Then
'                Call CloseReader
'                End
'            End If
''        Next
'    Else
'        no = Null
'        AseResult = ASEHLWriteUnprotect(hAseCard, WRITE_READ_DATA_SIZE, OFFSET, chWriteBuffer(0), 0)
'        If (ReportResult(AseResult) = 0) Then
'            Call CloseReader
'            End
'        End If
'    End If
'
    
    '-----------------------------------------------------------------------------------------
    '   Initialize the read buffer - set it to 0
    '-----------------------------------------------------------------------------------------
    For i = 0 To WRITE_READ_DATA_SIZE - 1
        chreadbuffer(i) = 0
    Next i

    msASEUserId = ""

    '-----------------------------------------------------------------------------------------
    '   Read from file number 6
    '-----------------------------------------------------------------------------------------
    Debug.Print "Read from file number 6."
    DoEvents
    
'    For I = 0 To WRITE_READ_DATA_SIZE - 1
        AseResult = ASEHLRead(hAseCard, WRITE_READ_DATA_SIZE, OFFSET, chreadbuffer(0))
        If (ReportResult(AseResult) = 0) Then
            Call CloseReader
            End
        End If
 '   Next

    For i = 0 To WRITE_READ_DATA_SIZE - 1
        msASEUserId = msASEUserId & Chr(chreadbuffer(i))
    Next

    GetASEUserId = RTrim(msASEUserId)
    
    '-----------------------------------------------------------------------------------------
    '   Check if chWriteBuffer and chReadBuffer are identical
    '-----------------------------------------------------------------------------------------
'    Debug.Print "Check that the buffer are identical."
'    DoEvents
'    WriteBufferTmp = chWriteBuffer
'    ReadBufferTmp = chReadBuffer
'    If (WriteBufferTmp <> ReadBufferTmp) Then
'        MsgBox "Error: The buffers are not identical."
'        Call CloseReader
'        End
'    End If
'    Debug.Print "The buffers are identicale."
    
    Call CloseReader
       
    Exit Function
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GetASEUserID")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Function

'=============================================================================================
' @NAME: ReportResult
' @DESC: This fuction displays the error that has occured
' @ARGS: AseResult As Long - result of previous call to API
' @RTRN: True or False( depends on wether the result is positive or not )
'=============================================================================================
Private Function ReportResult(AseResult As Long) As Boolean

Dim bResult As Boolean

    ' Let errors be handed by calling routine - NCJ 10/11/99
    
    Select Case AseResult
        Case ASEERR_SUCCESS
            Debug.Print (" The function succeeded")
            
        Case ASEERR_FAIL
            MsgBox ("The function failed")
    
        Case ASEERR_READER_ALREADY_OPEN
            MsgBox ("The specified reader is already opened")
    
        Case ASEERR_TIMEOUT
            MsgBox ("The function returned after timeout")
    
        Case ASEERR_WRONG_READER_NAME
            MsgBox ("No reader with the specified name exists")
    
        Case ASEERR_READER_OPEN_ERROR
            MsgBox ("The specified reader could not be opened")
    
        Case ASEERR_READER_COMM_ERROR
            MsgBox ("Reader communication error")
    
        Case ASEERR_MAX_READERS_ALREADY_OPEN
            MsgBox ("The maximum number of readers is already opened")
    
        Case ASEERR_INVALID_READER_HANDLE
            MsgBox ("The specified reader handle is invalid")
    
        Case ASEERR_SYSTEM_ERROR
            MsgBox ("General system error has occurred")
        
        Case ASEERR_INVALID_SOCKET
            MsgBox ("The specified socket is invalid")
    
        Case ASEERR_OPERATION_TIMEOUT
            MsgBox ("Blocking operation has been canceled after timeout")
    
        Case ASEERR_OPERATION_CANCELED
            MsgBox ("Blocking operation has been canceled by the user")
    
        Case ASEERR_INVALID_PARAMETERS
            MsgBox ("One or more of the specified parameters is invalid")
    
        Case ASEERR_PROTOCOL_NOT_SUPPORTED
            MsgBox ("The specified protocol is not supported")
    
        Case ASEERR_CARD_COMM_ERROR
            MsgBox ("Card communication error")
    
        Case ASEERR_CARD_NOT_PRESENT
            MsgBox ("Please insert your SmartCard into the card reader.")
    
        Case ASEERR_CARD_NOT_POWERED
            MsgBox ("The card is not powered on")
    
        Case ASEERR_IFSD_OVERFLOW
            MsgBox ("The command's data length is too big")
    
        Case ASEERR_CARD_INVALID_PARAMETER
            MsgBox ("One or more of the card parameters is invalid")
    
        Case ASEERR_INVALID_CARD_HANDLE
            MsgBox ("The specified card handle is invalid")
    
        Case ASEERR_NOT_INSTALLED
            MsgBox ("Ase is not installed in your system")
    
        Case ASEERR_COMMAND_NOT_SUPPORTED
            MsgBox ("This command is not supported. Read manual for details")
    
        Case ASEERR_MEMORY_CARD_ERROR
            MsgBox ("A memory card error has occurred. Read manual for details")
    
        Case ASEERR_NO_RTC
            MsgBox ("There is no RTC on this reader")
    
        Case ASEERR_WRONG_ACTIVE_PROTOCOL
            MsgBox ("This command can not work with the current protocol")
    
        Case ASEERR_NO_READER_AT_PORT
            MsgBox ("There is no ASE reader in the specified port")
    
        Case ASEERR_CARD_ALREADY_POWERED
            MsgBox ("Card is already powered")
    
        Case ASEERR_NO_HL_CARD_SUPPORT
            MsgBox ("The card has no ASE high level API support")
    
        Case ASEERR_CANT_LOAD_CARD_DLL
            MsgBox ("Can not load the high level API of the current card")
    
        Case ASEERR_WRONG_PASSWORD
            MsgBox ("Wrong password")
    
        Case ASEWRN_SERIAL_NUMBER_MISMATCH
            MsgBox ("The reader serial number does not match the registered one")
            
        'High Level errors
        Case ASEHLERR_UNSUPPORTED_CARD
            MsgBox ("ASEHLERR_UNSUPPORTED_CARD")
            
        Case ASEHLERR_KEY_EXISTS
            MsgBox ("ASEHLERR_KEY_EXISTS")
            
        Case ASEHLERR_INVALID_ID
            MsgBox ("ASEHLERR_INVALID_ID")
            
        Case ASEHLERR_INVALID_OFFSET
            MsgBox ("ASEHLERR_INVALID_OFFSET")

        Case ASEHLERR_UNFULFILLED_CONDITIONS
            MsgBox ("ASEHLERR_UNFULFILLED_CONDITIONS")

        Case ASEHLERR_INVALID_LENGTH
            MsgBox ("ASEHLERR_INVALID_LENGTH")

        Case ASEHLERR_WRONG_KEY
            MsgBox ("ASEHLERR_WRONG_KEY")

        Case ASEHLERR_BLOCKED
            MsgBox ("ASEHLERR_BLOCKED")
            
        Case ASEHLERR_SECURE_WRITE_UNSUPPORTED
            MsgBox ("ASEHLERR_SECURE_WRITE_UNSUPPORTED")

        Case ASEHLERR_CARD_MEMORY_PROBLEM
            MsgBox ("ASEHLERR_CARD_MEMORY_PROBLEM")

            
        Case ASEHLERR_INVALID_KEYREF
            MsgBox ("ASEHLERR_INVALID_KEYREF")

        Case ASEHLERR_UNSUPPORTED_FUNCTION
            MsgBox ("ASEHLERR_UNSUPPORTED_FUNCTION")

        Case ASEHLERR_KEY_NOT_EXIST
            MsgBox ("ASEHLERR_KEY_NOT_EXIST")

        Case ASEHLERR_CARD_INSUFFICIENT_MEMORY
            MsgBox ("ASEHLERR_CARD_INSUFFICIENT_MEMORY")
                                                  
        Case ASEHLERR_ID_ALREADY_EXISTS
            MsgBox ("ASEHLERR_ID_ALREADY_EXISTS")
                                                  
        Case ASEHLERR_API_FATAL_ERROR
            MsgBox ("ASEHLERR_API_FATAL_ERROR")
            
        Case ASEHLERR_API
            MsgBox ("ASEHLERR_API")
            
        Case ASEHLERR_INCORRECT_PARAMETER
            MsgBox ("ASEHLERR_INCORRECT_PARAMETER")
            
        Case ASEHLERR_INVALID_FILE
            MsgBox ("ASEHLERR_INVALID_FILE")
            
        Case ASEHLERR_FILE_NOT_OPEN
            MsgBox ("ASEHLERR_FILE_NOT_OPEN")
            
        Case ASEHLERR_NO_MORE_CHANGES
            MsgBox ("ASEHLERR_NO_MORE_CHANGES")
            
        Case ASEHLERR_FAILURE
            MsgBox ("ASEHLERR_FAILURE")
            
        Case ASEHLERR_CARD_FATAL_ERROR
            MsgBox ("ASEHLERR_CARD_FATAL_ERROR")
            
        Case ASEHLERR_CARD_ERROR
            MsgBox ("ASEHLERR_CARD_ERROR")
            
        Case Else
            MsgBox ("Unknown ASE error " & AseResult)
    
    End Select

    If AseResult = ASEERR_SUCCESS Or AseResult = ASEERR_CARD_ALREADY_POWERED Then
        bResult = True
    Else
        bResult = False
    End If


    ReportResult = bResult

End Function

'=============================================================================================
' @NAME: CloseReader
' @DESC: Closes reader handler
' @ARGS: NONE
' @RTRN: NONE
'=============================================================================================
Private Sub CloseReader()
    Debug.Print "Close the reader."
    AseResult = ASEReaderClose(hAseReader)
    ReportResult (AseResult)
End Sub
#End If




