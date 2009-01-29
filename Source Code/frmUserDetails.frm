VERSION 5.00
Begin VB.Form frmUserDetails 
   BorderStyle     =   0  'None
   Caption         =   "User Details"
   ClientHeight    =   4485
   ClientLeft      =   11250
   ClientTop       =   4935
   ClientWidth     =   3900
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   345
      Left            =   2640
      TabIndex        =   13
      Top             =   4020
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Details"
      Height          =   2235
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   3795
      Begin VB.TextBox txtUserCode 
         Height          =   375
         Left            =   1500
         MaxLength       =   20
         TabIndex        =   1
         Top             =   240
         Width           =   2235
      End
      Begin VB.TextBox txtUserName 
         Height          =   375
         Left            =   1500
         TabIndex        =   2
         Top             =   720
         Width           =   2235
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1500
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1200
         Width           =   2235
      End
      Begin VB.TextBox txtConfirm 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1500
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1680
         Width           =   2235
      End
      Begin VB.Label lblUser 
         Alignment       =   1  'Right Justify
         Caption         =   "&User"
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label lblFullName 
         Alignment       =   1  'Right Justify
         Caption         =   "Full &Name"
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   780
         Width           =   1305
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "&Password"
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   1260
         Width           =   1305
      End
      Begin VB.Label lblCPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Con&firm Password"
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Top             =   1740
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Reset"
      Height          =   345
      Left            =   1350
      TabIndex        =   7
      Top             =   4020
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Apply"
      Height          =   345
      Left            =   60
      TabIndex        =   6
      Top             =   4020
      Width           =   1215
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Status"
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   2340
      Width           =   3795
      Begin VB.CheckBox chkActiveDirectory 
         Caption         =   "Active Directory User"
         Height          =   195
         Left            =   780
         TabIndex        =   16
         Top             =   1260
         Width           =   2055
      End
      Begin VB.CheckBox chkLocked 
         Caption         =   "Locked Out"
         Height          =   315
         Left            =   780
         TabIndex        =   15
         Top             =   560
         Width           =   1335
      End
      Begin VB.CheckBox chkSysAdmin 
         Caption         =   "System Administrator"
         Height          =   315
         Left            =   780
         TabIndex        =   14
         Top             =   880
         Width           =   1875
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
         Height          =   315
         Left            =   780
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmUserDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       frmUserDetails.frm
'   Author:     Ashitei Trebi-Ollennu, September 2002
'   Purpose:    Adds / Edits a user.
'------------------------------------------------------------------------------
'REVISIONS:
'    ASH 14/01/2003 - Minor change in txtusername_change event. Also added cmdNew.
'   ic  16/12/2005  added active directory login
'   NCJ 27 Feb 08 - Disable "New" button if user does not have "Create User" permission;
'               Disable password fields if user does not have ResetPassword permission
'------------------------------------------------------------------------------

Option Explicit

Private Enum eUserStatus
    usDisabled = 0
    usEnabled = 1
End Enum

Private Enum eUserLock
    ulUnlocked = 0
    ulLockout = 1
End Enum

Private Enum eUserDetails
    udNewUser = 0
    udEditUser = 1
    udDisableUser = 2
End Enum

Private mbNewUser As Boolean
Private msUserName As String
Private msUserNameFull As String
Private mbEnabled As Boolean
Private mbSysAdmin As Boolean
Private mnLocked As Integer
Private mbChanged As Boolean
Private mbIsLoading As Boolean
Private mbAllowPasswordChange As Boolean
Private mbAllowConfirmChange As Boolean
Private mbCreateNew As Boolean

'ic 07/12/2005 active directory checkbox
Private mbActiveDirectory As Boolean


'------------------------------------------------------------------------------
Public Sub Display(ByVal bNewUser As Boolean, _
                    Optional sUserName As String = "")
'------------------------------------------------------------------------------
'Decides whether EDIT USER or ADD USER button clicked
'revisions
' ic 07/12/2005 added active directory checkbox
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    mbIsLoading = True

    Me.Icon = frmMenu.Icon
    FormCentre Me
    
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    mbChanged = False
    
    'only enable sys admin check box if user is a system administrator
    chkSysAdmin.Enabled = goUser.SysAdmin
    
    'only enable active directory user button if active directory is enabled
    chkActiveDirectory.Enabled = gbActiveDirectory
    
    msUserName = sUserName
    ' NCJ 27 Feb 08 - Only allow New if user has correct permission
    mbNewUser = bNewUser And goUser.CheckPermission(gsFnCreateNewUser)
    If mbNewUser = True Then
        txtUserCode.Text = ""
        txtUserName.Text = ""
        txtUserCode.Enabled = True
        txtUserCode.SetFocus
        txtPassword.Text = ""
        txtConfirm.Text = ""
        chkEnabled.Value = 0
        chkSysAdmin.Value = 0
        chkLocked.Value = 0
        chkLocked.Enabled = False
        cmdNew.Enabled = True
        chkEnabled.Enabled = True
        txtConfirm.Enabled = True
        txtPassword.Enabled = True
        txtUserName.Enabled = True
        
        'ic 07/12/2005 initialise active directory checkbox
        chkActiveDirectory.Value = 0
        
        Me.Show
    Else
        mbAllowPasswordChange = False
        mbAllowConfirmChange = False
        Call UserToEdit
        Me.Show
    End If
    
    mbIsLoading = False
    mbCreateNew = False
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserDetails.Display"
End Sub

'------------------------------------------------------------------------------------'
Private Sub chkActiveDirectory_Click()
'------------------------------------------------------------------------------------'
' ic 07/12/2005
' active directory checkbox
'------------------------------------------------------------------------------------'
    If Not mbIsLoading Then
        mbChanged = True
        EnableOKButton
    End If
End Sub

'------------------------------------------------------------------------------------'
Private Sub chkEnabled_Click()
'------------------------------------------------------------------------------------'
    
    If Not mbIsLoading Then
        mbChanged = True
        EnableOKButton
    End If

End Sub

'------------------------------------------------------------------------------------'
Private Sub chkLocked_Click()
'------------------------------------------------------------------------------------'

    If Not mbIsLoading Then
        mbChanged = True
        EnableOKButton
    End If
End Sub

'------------------------------------------------------------------------------------'
Private Sub chkSysAdmin_Click()
'------------------------------------------------------------------------------------'
    
    If Not mbIsLoading Then
        mbChanged = True
        EnableOKButton
    End If
    
End Sub

'-----------------------------------------------------------------------------------
Private Sub cmdNew_Click()
'-----------------------------------------------------------------------------------
'clears controls for new users to be added
'-----------------------------------------------------------------------------------
    
    mbCreateNew = True
    mbNewUser = True
    txtUserCode.Text = ""
    txtUserName.Text = ""
    ' NCJ 28 Feb 08 - Enable password fields in case previously disabled
    txtPassword.Enabled = True
    txtPassword.Text = ""
    txtConfirm.Enabled = True
    txtConfirm.Text = ""
    chkEnabled.Value = 0
    chkSysAdmin.Value = 0
    chkLocked.Value = 0
    chkLocked.Enabled = False
    txtUserCode.Enabled = True
    
    'ic 07/12/2005 active directory checkbox
    chkActiveDirectory.Value = 0
    
    txtUserCode.SetFocus
    
End Sub

'------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'------------------------------------------------------------------------------------'
'  If creating a  new user, check to see if the  user exists already in the database using gblnUserExists
'  Else updating an existing user
'  ic 16/12/2005    added active directory login
'------------------------------------------------------------------------------------'
Dim sUserName As String
Dim sUserNameFull As String
Dim sPassword As String
Dim sConfirmPassword As String
Dim bEnabled As Boolean
Dim bSysAdmin As Boolean
Dim sMessage As String
Dim bRefreshTree As Boolean
Dim bActiveDirectory As Boolean

    On Error GoTo ErrHandler
    sUserName = Trim(txtUserCode.Text)
    sUserNameFull = Trim(txtUserName.Text)
    sPassword = txtPassword.Text
    sConfirmPassword = txtConfirm.Text
    bEnabled = (chkEnabled.Value = 1)
    bSysAdmin = (chkSysAdmin.Value = 1)
    bActiveDirectory = (chkActiveDirectory.Value = 1)
    
    If sPassword <> sConfirmPassword Then
        Call DialogError("The password does not match the confirm password.", gsDIALOG_TITLE)
        Exit Sub
    End If
    
    If mbNewUser Then 'if new user then creat one
        ' Check to see if the  user name already exists using gblnUserExists
        If gblnUserExists(sUserName) Then
            sMessage = "Sorry, a user with the name '" & sUserName & "' already exists. Each user name must be unique."
            Call DialogError(sMessage)
            txtUserCode.Text = ""
            txtUserCode.SetFocus
            Exit Sub
        Else
            If Not InsertNewUser(sUserName, sUserNameFull, sPassword, bEnabled, bSysAdmin, bActiveDirectory, sMessage) Then
                txtPassword.SetFocus
                
            Else
                'REM 10/02/03 - set to false as no longer new user once they have been created
                mbNewUser = False
                bRefreshTree = True
            End If
        End If
    Else
        Call UpdateUser(sUserName, sUserNameFull, sPassword, bEnabled, bSysAdmin, bActiveDirectory, sMessage, bRefreshTree)
    End If
    
    If bRefreshTree Then
        frmMenu.RefereshTreeView
    End If
    
    txtPassword.Text = "aaaaaa"
    txtConfirm.Text = "aaaaaa"
    
    mbChanged = False
    mbAllowPasswordChange = False
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    txtUserCode.Enabled = False
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "frmUserDetails.cmdOK_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'------------------------------------------------------------------------------------'
Private Sub txtConfirm_KeyDown(KeyCode As Integer, Shift As Integer)
'------------------------------------------------------------------------------------'
Dim nSelLength As Integer
Dim nTextLength As Integer
    
    nSelLength = txtConfirm.SelLength
    nTextLength = Len(txtConfirm.Text)
    
    If nSelLength = nTextLength Then
        mbAllowConfirmChange = True
    Else
        'check to see if mbAllowPasswordChange is already true from previous key down event
        If mbAllowConfirmChange Then
            'continue, as already true so user can change password
            
        Else 'set it to false as user is trying to chaneg password without selecting all
            mbAllowConfirmChange = False
        End If
    End If

End Sub


'------------------------------------------------------------------------------------'
Private Sub txtConfirm_Change()
'------------------------------------------------------------------------------------'
' Validate password
'------------------------------------------------------------------------------------

    On Error GoTo ErrHandler

    If Not mbIsLoading Then
        If mbAllowConfirmChange Then
            Call TextBoxChange(txtConfirm)
            txtConfirm.ToolTipText = ""
        Else
            If Not mbCreateNew Then
                txtConfirm.Text = "aaaaaa"
                txtConfirm.ToolTipText = "Select all to change"
            End If
        End If
    End If
    
    txtConfirm.Tag = txtConfirm.Text
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserDetails.txtConfirm_Change"
End Sub

'------------------------------------------------------------------------------------'
Private Sub txtConfirm_GotFocus()
'------------------------------------------------------------------------------------'
'
'------------------------------------------------------------------------------------'

    txtConfirm.SelStart = 0
    txtConfirm.SelLength = Len(txtConfirm.Text)

End Sub

'------------------------------------------------------------------------------------'
Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
'------------------------------------------------------------------------------------'
Dim nSelLength As Integer
Dim nTextLength As Integer
    
    nSelLength = txtPassword.SelLength
    nTextLength = Len(txtPassword.Text)
    
    If nSelLength = nTextLength Then
        mbAllowPasswordChange = True
    Else
        'check to see if mbAllowPasswordChange is already true from previous key down event
        If mbAllowPasswordChange Then
            'continue, as already true so user can change password
            
        Else 'set it to false as user is trying to chaneg password without selecting all
            mbAllowPasswordChange = False
        End If
    End If
    
End Sub
'------------------------------------------------------------------------------------'
Private Sub txtPassword_Change()
'------------------------------------------------------------------------------------'
' Validate password
'------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    If Not mbIsLoading Then
        
        If mbAllowPasswordChange Then
            Call TextBoxChange(txtPassword)
            txtPassword.ToolTipText = ""
        Else
            If Not mbCreateNew Then
                txtPassword.Text = "aaaaaa"
                txtPassword.ToolTipText = "Select all to change"
            End If
        End If
    End If
    
    txtPassword.Tag = txtPassword.Text
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserDetails.txtPassword_Change"
End Sub

'------------------------------------------------------------------------------------'
Private Sub txtPassword_GotFocus()
'------------------------------------------------------------------------------------'
'
'------------------------------------------------------------------------------------'

    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)

End Sub

'------------------------------------------------------------------------------------'
Private Sub EnableOKButton()
'------------------------------------------------------------------------------------'
' Enables/disables the OK button
'------------------------------------------------------------------------------------'

    If txtConfirm = "" Or _
        txtUserCode.Text = "" Or _
        txtUserName.Text = "" Or _
        txtPassword.Text = "" Then
        cmdOK.Enabled = False
        cmdCancel.Enabled = False
    Else
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
    End If

End Sub

'------------------------------------------------------------------------------------'
Private Sub TextBoxChange(txtTextBox As TextBox)
'------------------------------------------------------------------------------------'
' NCJ 26/10/00
' Validate a text field on this form
'------------------------------------------------------------------------------------'
Dim sText As String
    
    sText = Trim(txtTextBox.Text)
    
    If sText > "" Then
        If IsValidString(sText) = False Then
            ' Replace field with previous contents
            txtTextBox.Text = txtTextBox.Tag
            ' Put cursor at the end
            txtTextBox.SelStart = Len(txtTextBox.Text)
        Else
            ' Store contents for next time
            txtTextBox.Tag = sText
        End If
    Else
        ' Screen out superfluous spaces
        txtTextBox.Text = ""
        txtTextBox.Tag = ""
    End If
    
    ' Enable the OK button as appropriate
    Call EnableOKButton

End Sub

'------------------------------------------------------------------------------------'
Private Sub txtUserCode_Change()
'------------------------------------------------------------------------------------'
' Validate user code
'------------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler
    
    If Not mbIsLoading Then
        mbChanged = True
        Call TextBoxChange(txtUserCode)
    End If
    
    txtUserCode.Tag = txtUserCode.Text
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserDetails.txtUserCode_Change"
End Sub

'------------------------------------------------------------------------------------'
Private Sub txtUserName_Change()
'------------------------------------------------------------------------------------'
' Validate user name
' MLM 19/06/03: 3.0 buglist 1153: Limit input to 255 characters.
'------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    'ASH 14/01/2003 disallow certain characters
    If Not mbIsLoading Then
        With txtUserName
            mbChanged = True
            If Not gblnValidString(txtUserName.Text, valOnlySingleQuotes) Then
                Call DialogInformation("User name contains invalid characters")
            ElseIf InStr(.Text, "%") > 0 Then
                Call DialogInformation("The character " & "%" & " is not allowed as part of a user name.")
                .Text = .Tag
                Exit Sub
            ElseIf InStr(.Text, "*") > 0 Then
                Call DialogInformation("The character " & "*" & " is not allowed as part of a user name.")
                .Text = .Tag
                Exit Sub
            ElseIf InStr(.Text, ">") > 0 Then
                Call DialogInformation("The character " & ">" & " is not allowed as part of a user name.")
                .Text = .Tag
                Exit Sub
            ElseIf InStr(.Text, "<") > 0 Then
                Call DialogInformation("The character " & "<" & " is not allowed as part of a user name.")
                .Text = .Tag
                Exit Sub
            ElseIf InStr(.Text, "=") > 0 Then
                Call DialogInformation("The character " & "=" & " is not allowed as part of a user name.")
                .Text = .Tag
                Exit Sub
            ElseIf Len(.Text) > 255 Then
                Call DialogInformation("The user name must not contain more than 100 characters.")
                .Text = .Tag
                Exit Sub
            End If
            
            .Tag = .Text
            Call EnableOKButton
        End With
    End If
    
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserDetails.txtUserName_Change"
End Sub

'----------------------------------------------------------------------------------------'
Private Function IsValidString(sDescription As String) As Boolean
'----------------------------------------------------------------------------------------'
' Return TRUE if text is valid
' Displays any necessary messages
'----------------------------------------------------------------------------------------'
Dim sMSG As String


    On Error GoTo ErrHandler
    
    IsValidString = False
    
    If sDescription > "" Then
        sMSG = "A password or user name"
        If Not gblnValidString(sDescription, valOnlySingleQuotes) Then
            sMSG = sMSG & gsCANNOT_CONTAIN_INVALID_CHARS
            Call DialogError(sMSG)
        'after the single quote check remove it so the next check won't fail, as it will fail on a single quote
        ElseIf Not gblnValidString(sDescription, valAlpha + valNumeric + valSpace) Then
            sMSG = sMSG & " may only contain alphanumeric characters"
            Call DialogError(sMSG)
        ElseIf Len(sDescription) > 255 Then
            sMSG = sMSG & " may not be more than 255 characters"
            Call DialogError(sMSG)
        Else
            IsValidString = True
        End If
    End If
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserDetails.IsValidString"
End Function

'----------------------------------------------------------------------------------------'
Private Sub UserToEdit()
'----------------------------------------------------------------------------------------'
'this returns all the details of the user to edit
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsUser As ADODB.Recordset
Dim rsPasswords As ADODB.Recordset

    On Error GoTo ErrHandler
    
    sSQL = " Select * FROM MACROUser WHERE UserName ='" & msUserName & "'"
    Set rsUser = New ADODB.Recordset
    rsUser.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
    
    sSQL = " Select PasswordRetries FROM MACROPassword"
    Set rsPasswords = New ADODB.Recordset
    rsPasswords.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
    
    msUserNameFull = rsUser!UserNameFull
    
    'update controls with records from database
    txtUserCode.Enabled = False
    txtUserCode.Text = msUserName
    txtUserName.Text = msUserNameFull
    txtPassword.Text = "aaaaaa"
    txtConfirm.Text = "aaaaaa"
    chkSysAdmin.Value = rsUser!SysAdmin
    chkEnabled.Value = rsUser!Enabled
    
    'ic 07/12/2005 active directory checkbox
    If (gbActiveDirectory) Then
        If (IsNull(rsUser!Authentication)) Then
            chkActiveDirectory.Value = 0
            mbActiveDirectory = False
        Else
            chkActiveDirectory.Value = rsUser!Authentication
            mbActiveDirectory = (rsUser!Authentication = 1)
        End If
    End If
    
    mbEnabled = (rsUser!Enabled = 1)
    
    mbSysAdmin = (rsUser!SysAdmin = 1)
    
    If UserSysAdmin Then 'if user is not a system admin then can't edit other system admin users
    
        If (rsPasswords!PasswordRetries <> 0) And (rsUser!FailedAttempts >= rsPasswords!PasswordRetries) Then
            chkLocked.Enabled = True
            mnLocked = eUserLock.ulLockout
            chkLocked.Value = mnLocked
        Else
            chkLocked.Enabled = False
            mnLocked = eUserLock.ulUnlocked
            chkLocked.Value = mnLocked
        End If
        chkSysAdmin.Enabled = goUser.SysAdmin
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserDetails.UserToEdit"
End Sub

'-------------------------------------------------------------------------------------
Private Function UserSysAdmin() As Boolean
'-------------------------------------------------------------------------------------
'REM 24/01/03
'Checks to see if user being edited is a sys admin, if so then user logged in has to be sys admin to edit user
'-------------------------------------------------------------------------------------
    
    If Not goUser.SysAdmin Then
        If mbSysAdmin Then
            UserSysAdmin = False
        Else
            UserSysAdmin = True
        End If
    Else
        UserSysAdmin = True
    End If

    ' NCJ 27 Feb 08 - Only enable New if user has correct permission
    cmdNew.Enabled = UserSysAdmin And goUser.CheckPermission(gsFnCreateNewUser)
    chkSysAdmin.Enabled = UserSysAdmin
    chkLocked.Enabled = UserSysAdmin
    chkEnabled.Enabled = UserSysAdmin
    ' NCJ 28 Feb 08 - Also check Reset Password permission for existing users
    txtPassword.Enabled = UserSysAdmin And (mbNewUser Or goUser.CheckPermission(gsFnResetPassword))
    txtConfirm.Enabled = txtPassword.Enabled
    
    txtUserName.Enabled = UserSysAdmin
    
    'ic 07/12/2005 active directory checkbox
    chkActiveDirectory.Enabled = UserSysAdmin And gbActiveDirectory

End Function

'-------------------------------------------------------------------------------------
Private Function UpdateUser(ByVal sUserName As String, sUserNameFull As String, sPassword As String, _
                            bEnabled As Boolean, bSysAdmin As Boolean, bActiveDirectory As Boolean, _
                            ByVal sMessage As String, ByRef bRefreshTree As Boolean) As Boolean
'-------------------------------------------------------------------------------------
'This updates the user being edited.
'REVISIONS:
'REM 08/12/03 - change password now resets user password to expire if MACRO Setting "expirepassword" = true
'    To tell the user object to expire the password I have passed in "sysadminreset".
'-------------------------------------------------------------------------------------
Dim rsFailedAtt As ADODB.Recordset
Dim bChangeUserPswd As Boolean
Dim sSQL As String
Dim nFailedAttempts As Integer
Dim oSystemMessage As SysMessages
Dim sMessageParameters As String
Dim sHashedPassword As String
Dim sPswdCreateDate As String
Dim sFirstLogin As String
Dim sLastLogin As String

    On Error GoTo ErrHandler
    
    bRefreshTree = False
    
    'if mbAllowPasswordChange is true means the password has been changed and should therefore be updated
    If mbAllowPasswordChange Then
        'REM 08/12/03 - Pass in "sysadminreset" to tell user object that this is a sys admin resetting a users password
        sMessage = "sysadminreset"
        'if returns false then password change was unsuccessful
        bChangeUserPswd = goUser.ChangeUserPassword(sUserName, sPassword, sMessage, sHashedPassword, sPswdCreateDate)
        
        If (GetMACROSetting("expirepassword", "true") = "true") Then
            sFirstLogin = "36000" 'expire new password by setting old date
        Else
            sFirstLogin = SQLStandardNow
        End If
    
        sLastLogin = SQLStandardNow
                
        If Not bChangeUserPswd Then
            Call DialogError(sMessage, gsDIALOG_TITLE)
            Call goUser.gLog(goUser.UserName, gsCHANGE_PSWD, "Change password for user " & sUserName & " failed. " & sMessage)
            UpdateUser = False
        Else
            'when resetting a user password always ensure they are also unlocked
            If chkLocked.Value = 1 Then
                chkLocked.Value = 0
                chkLocked.Enabled = False
            End If
            
            Call goUser.gLog(goUser.UserName, gsCHANGE_PSWD, "The password for user " & sUserName & " was changed using System Management")
            
            'write New Password to the Message table
            If Not ExcludeUserRDE(msUserName) Then 'don't write message if its user rde
                Set oSystemMessage = New SysMessages
                sMessageParameters = msUserName & gsPARAMSEPARATOR & sHashedPassword & gsPARAMSEPARATOR & sLastLogin & gsPARAMSEPARATOR & sFirstLogin & gsPARAMSEPARATOR & sPswdCreateDate
                Call oSystemMessage.AddNewSystemMessage(MacroADODBConnection, ExchangeMessageType.PasswordChange, goUser.UserName, sUserName, "Change Password", sMessageParameters)
                Set oSystemMessage = Nothing
            End If
        End If
    End If
    
    'if mbChanged is true then other user info has been changed and should be saved
    If mbChanged Then
        
        'check to see if enabled status has changed if so log the change
        If mbEnabled <> bEnabled Then
            If bEnabled Then
                Call goUser.gLog(goUser.UserName, gsUSER_ENABLED, "User " & sUserName & " status changed to enabled")
            Else
                Call goUser.gLog(goUser.UserName, gsUSER_DISABLED, "User " & sUserName & " status changed to disabled")
            End If
            mbEnabled = bEnabled
            bRefreshTree = True
        End If
        
        
        'if the Locked out check box is unticked then set failed attempts to zero and log the change
        If mnLocked <> chkLocked.Value Then
            If chkLocked = eUserLock.ulUnlocked Then
                nFailedAttempts = 0
                sSQL = "UPDATE MACROUser SET FailedAttempts = 0" _
                    & " WHERE UserName = '" & sUserName & "'"
                SecurityADODBConnection.Execute sSQL, , adCmdText
                
                chkLocked.Value = 0
                chkLocked.Enabled = False
                Call goUser.gLog(goUser.UserName, gsUSER_UNLOCKED, "User " & sUserName & " was unlocked")
            End If
            bRefreshTree = True
        Else 'just return what it currently is
            sSQL = "SELECT FailedAttempts FROM MACROUser" _
                & " WHERE UserName = '" & sUserName & "'"
            Set rsFailedAtt = New ADODB.Recordset
            rsFailedAtt.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
            nFailedAttempts = rsFailedAtt!FailedAttempts
            rsFailedAtt.Close
            Set rsFailedAtt = Nothing
        End If
        
        'if the UserNameFull was changed then log it
        If msUserNameFull <> sUserNameFull Then
            Call goUser.gLog(goUser.UserName, gsCHANGE_USERNAME_FULL, "The Full UserName for User " & sUserName & " was changed from " & msUserNameFull & " to " & sUserNameFull)
            msUserNameFull = sUserNameFull
        End If
        
        'if the system administrator status has changed then log it
        If mbSysAdmin <> bSysAdmin Then
            Call goUser.gLog(goUser.UserName, gsCHANGE_SYSADMIN_STATUS, "The System Admin status of user " & sUserName & " was changed")
            mbSysAdmin = bSysAdmin
            bRefreshTree = True
        End If
        
        'update UserNameFull, enabled, lockout and SysAdmin status
        sSQL = "UPDATE MACROUser SET " _
        & " UserNameFull = '" & ReplaceQuotes(sUserNameFull) & "'," _
        & " Enabled = " & -CInt(bEnabled) & "," _
        & " SysAdmin = " & -CInt(bSysAdmin)
        If (gbActiveDirectory) Then
            sSQL = sSQL & ", Authentication = " & -CInt(bActiveDirectory)
            
            If (sUserName = goUser.UserName) Then
                If (bActiveDirectory) Then
                    frmMenu.mnuLogOff.Enabled = False
                Else
                    frmMenu.mnuLogOff.Enabled = True
                End If
            End If
        End If
        sSQL = sSQL & " WHERE UserName = '" & sUserName & "'"
        SecurityADODBConnection.Execute sSQL, , adCmdText
        
        'Write edited Users Details to Message table
        If Not ExcludeUserRDE(sUserName) Then 'don't write message if its rde
            Set oSystemMessage = New SysMessages
            sMessageParameters = sUserName & gsPARAMSEPARATOR & sUserNameFull & gsPARAMSEPARATOR & "" & gsPARAMSEPARATOR & -CInt(bEnabled) & gsPARAMSEPARATOR & 0 & gsPARAMSEPARATOR & 0 & gsPARAMSEPARATOR & nFailedAttempts & gsPARAMSEPARATOR & 0 & gsPARAMSEPARATOR & -CInt(bSysAdmin) & gsPARAMSEPARATOR & eUserDetails.udEditUser
            Call oSystemMessage.AddNewSystemMessage(MacroADODBConnection, ExchangeMessageType.User, goUser.UserName, sUserName, "Edit User Details", sMessageParameters)
            Set oSystemMessage = Nothing
        End If
        
    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserDetails.UpdateUser"
End Function

'------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'------------------------------------------------------------------------------------'

    If (mbChanged = True) Or (mbAllowPasswordChange = True) Or (mbAllowConfirmChange = True) Then
        If DialogQuestion("Are you sure you want to lose the changes?", gsDIALOG_TITLE) = vbYes Then

            If mbNewUser Then
                txtUserCode.Text = ""
                txtUserName.Text = ""
                txtPassword.Text = ""
                txtConfirm.Text = ""
                chkEnabled.Value = 0
                chkLocked.Value = 0
                chkLocked.Enabled = True
                
                'ic 07/12/2005 active directory checkbox
                chkActiveDirectory.Value = 0
            Else
                mbAllowPasswordChange = False
                mbAllowConfirmChange = False
                mbChanged = False
                Call UserToEdit
                cmdOK.Enabled = False
                cmdCancel.Enabled = False
            End If
        
        End If
    End If
         
End Sub

'--------------------------------------------------------------------------------
Private Function InsertNewUser(sUserName As String, sUserNameFull As String, sPassword As String, _
                                bEnabled As Boolean, bSysAdmin As Boolean, bActiveDirectory As Boolean, _
                                ByRef sMessage As String) As Boolean
'--------------------------------------------------------------------------------
'REM 21/10/02
'Create a new user and do all the checks on the password
'--------------------------------------------------------------------------------
Dim bSucceed As Boolean
Dim sEncryptedPassword As String
Dim sPasswordCreated As String
Dim oSystemMessage As SysMessages
Dim sMessageParameters As String
Dim sFirstLogin As String
Dim sLastLogin As String
Dim sSite As String

    On Error GoTo ErrHandler
    
    bSucceed = goUser.CreateNewUser(sUserName, sUserNameFull, sPassword, bEnabled, bSysAdmin, sEncryptedPassword, sPasswordCreated, sMessage, bActiveDirectory)
    
    If Not bSucceed Then
        Call DialogError(sMessage, gsDIALOG_TITLE)
    Else
        'ZA 18/09/2002 - update List_user.js now
        Call CreateUsersList(goUser)
        
        'log creation of new user
        Call goUser.gLog(goUser.UserName, gsCREATE_NEW_USER, "A new user was created with UserName(" & sUserName & ") and UserNameFull(" & sUserNameFull & ")")
        
        'If MACRO setting is true then expire all new passwords
        If (GetMACROSetting("expirepassword", "true") = "true") Then
            sFirstLogin = "36000" 'expire new password by setting old date
        Else
            sFirstLogin = SQLStandardNow
        End If
        
        sLastLogin = SQLStandardNow
        
        'add the new user to the MACRO message table for Data Transfer
        Set oSystemMessage = New SysMessages
        'create message parameters
        sMessageParameters = sUserName & gsPARAMSEPARATOR & sUserNameFull & gsPARAMSEPARATOR & sEncryptedPassword & gsPARAMSEPARATOR & -CInt(bEnabled) & gsPARAMSEPARATOR & sLastLogin & gsPARAMSEPARATOR & sFirstLogin & gsPARAMSEPARATOR & 0 & gsPARAMSEPARATOR & sPasswordCreated & gsPARAMSEPARATOR & -CInt(bSysAdmin) & gsPARAMSEPARATOR & eUserDetails.udNewUser
        
        Call oSystemMessage.AddNewSystemMessage(MacroADODBConnection, ExchangeMessageType.User, goUser.UserName, sUserName, "Create New User", sMessageParameters)
        
        mbEnabled = bEnabled
    End If
    
    InsertNewUser = bSucceed
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserDetails.InsertNewUser"
End Function

'-------------------------------------------------------------
Private Sub txtUserName_GotFocus()
'-------------------------------------------------------------
'
'-------------------------------------------------------------

    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)

End Sub
