VERSION 5.00
Begin VB.Form frmTimeOutLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1665
   ClientLeft      =   7830
   ClientTop       =   6795
   ClientWidth     =   3405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3405
   Begin VB.Frame fraUser 
      Height          =   1095
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1125
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   630
         Width           =   2025
      End
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   1125
         TabIndex        =   4
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   645
         Width           =   900
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Username"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   255
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdExitMAcro 
      Caption         =   "E&xit MACRO"
      Height          =   375
      Left            =   2100
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   780
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "frmTimeOutLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       frmTimeOutLogin.frm
'   Author:     Will Casey, March 30 2000
'   Purpose:    Allows user login after timeout.
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------
'   Revisions:
'   NCJ 1 Apr 00 Tidying up
'   NCJ 27 Nov 02 - Use correct security database connection
'   TA 28/05/2003: If requested , propmpt user when eXit MACRO clicked
'---------------------------------------------------------------------

Option Explicit

Private mbLoginSucceeded As Boolean

'do we prompt user if exit clocked
Private mbPromptIfExitClicked As Boolean



Private Sub cmdExitMAcro_Click()

    If mbPromptIfExitClicked Then
        If DialogQuestion("Any unsaved data will be lost, are you sure you wish to exit?") = vbNo Then
        'they have chosen not to exit
'EXIT SUB
            Exit Sub
        End If
    End If

    Unload Me
End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
'check for correct password
'REM 07/04/03 - Added Account Disabled and Login Failed to timeout login.
'---------------------------------------------------------------------
Dim sPassword As String
Dim sMessage As String
     
'    Select Case goUser.Login(Connection_String(CONNECTION_MSJET_OLEDB_40, SecurityDatabasePath, , , gsSecurityDatabasePassword), txtUserName.Text, txtPassword.Text, "", "", "", True)
    Select Case goUser.Login(SecurityDatabasePath, _
                            txtUserName.Text, txtPassword.Text, DefaultHTMLLocation, "", sMessage, True)

    Case LoginResult.AccountDisabled
        'user account has been disabled
        Call DialogInformation(sMessage, gsDIALOG_TITLE)
        cmdExitMAcro.SetFocus
    
    Case LoginResult.Failed
        'user login failed
        Call DialogError(sMessage, gsDIALOG_TITLE)
        'set focus back to password text box and selected password
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
    
    Case Else
        mbLoginSucceeded = True
        Unload Me
    
    End Select
    
End Sub


'---------------------------------------------------------------------
Public Function Display(bPromptIfExitClicked As Boolean) As Boolean
'---------------------------------------------------------------------
' Display login form
'---------------------------------------------------------------------
    
    mbLoginSucceeded = False
    
    txtPassword.Text = vbNullString
    
    txtUserName.Text = goUser.UserName
    ' They can't change the user name
    txtUserName.Enabled = False
    cmdOK.Enabled = False
    Me.Icon = frmMenu.Icon
    
    FormCentre Me
    
    mbPromptIfExitClicked = bPromptIfExitClicked
    
    Me.Show vbModal
    
    Display = mbLoginSucceeded
    
End Function

Private Sub Form_Load()
    Me.Icon = frmMenu.Icon
End Sub

Private Sub txtPassword_Change()
    If Trim(txtPassword.Text) = "" Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
End Sub


