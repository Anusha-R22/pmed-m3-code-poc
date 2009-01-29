VERSION 5.00
Begin VB.Form frmTempLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Authorisation"
   ClientHeight    =   2925
   ClientLeft      =   11730
   ClientTop       =   2085
   ClientWidth     =   3465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   60
      TabIndex        =   6
      Top             =   1260
      Width           =   3255
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1125
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   630
         Width           =   2025
      End
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   1125
         TabIndex        =   0
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   645
         Width           =   900
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Username"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   255
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txtPrompt 
         Alignment       =   2  'Center
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   780
      TabIndex        =   2
      Top             =   2460
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2100
      TabIndex        =   3
      Top             =   2460
      Width           =   1215
   End
End
Attribute VB_Name = "frmTempLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999-2003. All Rights Reserved
'   File:       frmTempLogin.frm
'   Author:     Steve Morris, December 1999
'   Purpose:    Temporary login form when authorisation is required
'               during data entry
'
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
' TA 04/10/2001: Completely redone for 2.2
' NCJ 27 Nov 02 - Changed to use correct Security Database
' NCJ 8 Jan 03 - Return Full User Name too
' NCJ 3 Dec 03 - Changed message when role does not exist at all
' TA 20/06/2005: clinicaltialid and site are checked and user is also prompted to change password if required - cbd 2583
'----------------------------------------------------------------------------------------'

Option Explicit


'TA 20/06/2005: added checking clinicaltialid and site  - cbd 2583

Private msDatabaseName As String
Private mbLoginSucceeded As Boolean
Private msUserName As String
Private msUserNameFull As String
Private msRole As String
Private mlClinicalTrialId As Long
Private msSite As String

'-----------------------------------------------------------------------
Public Function Display(sRole As String, sDatabaseName As String, lClinicalTrialId As Long, sSite As String, _
                        ByRef sUserNameFull As String) As String
'-----------------------------------------------------------------------
' Only way to show this form
' Returns UserName if successfully authorised, or "" if not
' Also fills in sUserNameFull
'-----------------------------------------------------------------------
Dim sPrompt As String

    Load Me

    msDatabaseName = sDatabaseName
    mbLoginSucceeded = False
    msUserName = ""
    msUserNameFull = ""
    msRole = sRole
    
    
    'TA 20/06/2005: added clinicaltialid and site to call - cbd 2583
    mlClinicalTrialId = lClinicalTrialId
    msSite = sSite
    
    If LCase(goUser.UserRole) = LCase(sRole) Then
        sPrompt = "Please re-enter your username and password to authorise this response"
    Else
        sPrompt = "This question needs to be authorised by a user with the following role: " & vbCrLf & sRole
    End If
    
    txtPrompt.Text = sPrompt
    txtPassword.Text = ""
    txtUserName.Text = ""
    
    On Error Resume Next
    txtUserName.SetFocus
    
    EnableOKButton
    
    FormCentre Me
    Me.Show vbModal
    
   
    If mbLoginSucceeded Then
        sUserNameFull = msUserNameFull
        Display = msUserName
    Else
        Display = ""
    End If
    
End Function


'-----------------------------------------------------------------------
Private Sub cmdCancel_Click()
'-----------------------------------------------------------------------
' They've cancelled the login
'-----------------------------------------------------------------------

    mbLoginSucceeded = False
    msUserName = ""
    msUserNameFull = ""
    Unload Me

End Sub
Private Function CheckAuthAfterLogin(oUser As MACROUser, lClinicalTrialId As Long) As Boolean
Dim colSites As Collection
Dim sSite As Site

    CheckAuthAfterLogin = False
    
    'TA 20/06/2005: added clinicaltialid and site to call - cbd 2583
    Set colSites = oUser.GetAllSites(lClinicalTrialId)
    For Each sSite In colSites
        If msSite = sSite.Site Then
            CheckAuthAfterLogin = True
            Exit Function
        End If
    Next

End Function
'-----------------------------------------------------------------------
Private Sub cmdOK_Click()
'-----------------------------------------------------------------------
' Store the result of their log in
' then hide window (so form variables are still accessible)
'-----------------------------------------------------------------------
Dim oMACROUser As MACROUser
Dim sMessage As String
Dim sMessageIfError As String
Dim sNewPassword As String
Dim sWrongRoleMessage As String

    msUserName = Trim(txtUserName.Text)
    Set oMACROUser = New MACROUser
    mbLoginSucceeded = False
    
    sWrongRoleMessage = "You are not authorised to answer this question.  Only users with role " & msRole & " can answer this question."
    
    Select Case oMACROUser.Login(SecurityDatabasePath, msUserName, Trim(txtPassword.Text), DefaultHTMLLocation, GetApplicationTitle, sMessage, True, msDatabaseName, msRole)
    Case LoginResult.Success
        If oMACROUser.Login(SecurityDatabasePath, msUserName, Trim(txtPassword.Text), DefaultHTMLLocation, GetApplicationTitle, sMessage, False, msDatabaseName, msRole) = LoginResult.Success Then
            sMessageIfError = sWrongRoleMessage
            mbLoginSucceeded = CheckAuthAfterLogin(oMACROUser, mlClinicalTrialId)
        End If
    Case LoginResult.ChangePassword
        sMessageIfError = ""
        'display message box asking user if they want to change their password,
        'if so call update password routine else
        If DialogQuestion(sMessage, gsDIALOG_TITLE) = vbYes Then
            If frmPasswordChange.Display(oMACROUser, SecurityDatabasePath, sNewPassword, txtPassword.Text) Then
                'write to LoginLog table
                Call oMACROUser.gLog(oMACROUser.UserName, gsLOGIN, "User changed password")
                If oMACROUser.Login(SecurityDatabasePath, msUserName, sNewPassword, DefaultHTMLLocation, GetApplicationTitle, sMessage, False, msDatabaseName, msRole) = LoginResult.Success Then
                    sMessageIfError = sWrongRoleMessage
                    mbLoginSucceeded = CheckAuthAfterLogin(oMACROUser, mlClinicalTrialId)
                End If
            End If
        Else
            If oMACROUser.Login(SecurityDatabasePath, msUserName, Trim(txtPassword.Text), DefaultHTMLLocation, GetApplicationTitle, sMessage, False, msDatabaseName, msRole) = LoginResult.Success Then
                sMessageIfError = sWrongRoleMessage
                mbLoginSucceeded = CheckAuthAfterLogin(oMACROUser, mlClinicalTrialId)
            End If
        End If
    Case LoginResult.PasswordExpired
        'call the update password routine
        Call DialogInformation(sMessage, gsDIALOG_TITLE)
        If frmPasswordChange.Display(oMACROUser, SecurityDatabasePath, sNewPassword, txtPassword.Text) Then
            'Write to LoginLog table
            Call oMACROUser.gLog(oMACROUser.UserName, gsLOGIN, "User changed password")
            If oMACROUser.Login(SecurityDatabasePath, msUserName, sNewPassword, DefaultHTMLLocation, GetApplicationTitle, sMessage, False, msDatabaseName, msRole) = LoginResult.Success Then
                sMessageIfError = sWrongRoleMessage
                mbLoginSucceeded = CheckAuthAfterLogin(oMACROUser, mlClinicalTrialId)
            End If
        Else
            sMessageIfError = ""
        End If
    Case LoginResult.AccountDisabled
        'user account has been disabled
        sMessageIfError = sMessage
    Case Else
        ' NCJ 3 Dec 03 - It is no longer appropriate to append the login message!
        sMessageIfError = "Your authorisation login failed." & vbCrLf & "Please check your details and try again."
    End Select
    
    If mbLoginSucceeded Then
        ' NCJ 8 Jan 03 - Return Full User Name too
        msUserNameFull = oMACROUser.UserNameFull
        Unload Me
    Else
        If sMessageIfError <> "" Then
            DialogError sMessageIfError
        End If
        ResetUsername
    End If
   
    Set oMACROUser = Nothing

End Sub

'-----------------------------------------------------------------------
Private Sub ResetUsername()
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
        txtPassword.Text = ""
        txtUserName.SetFocus
        txtUserName.SelStart = 0
        txtUserName.SelLength = Len(txtUserName.Text)
End Sub

'-----------------------------------------------------------------------
Private Sub Form_Load()
    Me.Icon = frmMenu.Icon
End Sub

'-----------------------------------------------------------------------
Private Sub txtPassword_Change()
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------

    EnableOKButton
    
End Sub

'-----------------------------------------------------------------------
Private Sub txtUserName_Change()
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------

    EnableOKButton
    
End Sub

'-----------------------------------------------------------------------
Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
'-----------------------------------------------------------------------
' Intercept Return key and move to Password field
'-----------------------------------------------------------------------
    
    If KeyCode = vbKeyReturn Then
        txtPassword.SetFocus
    End If

End Sub

'-----------------------------------------------------------------------
Private Sub EnableOKButton()
'-----------------------------------------------------------------------
' Enable OK button only if user name and password are non-empty
'-----------------------------------------------------------------------

    If Trim(txtUserName.Text) = "" Or Trim(txtPassword.Text) = "" Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
    
End Sub
