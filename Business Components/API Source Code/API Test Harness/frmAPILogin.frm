VERSION 5.00
Begin VB.Form frmAPILogin 
   Caption         =   "API Log in"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSecurity 
      Caption         =   "Use specified security DB"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtSecDB 
      Height          =   375
      Left            =   1980
      TabIndex        =   8
      Top             =   1920
      Width           =   5055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   180
      TabIndex        =   7
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Log in"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txtDB 
      Height          =   345
      Left            =   1980
      TabIndex        =   2
      Text            =   "MACRODPH3"
      Top             =   1380
      Width           =   3195
   End
   Begin VB.TextBox txtPwd 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1980
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "macrotm"
      Top             =   840
      Width           =   3195
   End
   Begin VB.TextBox txtUser 
      Height          =   345
      Left            =   1980
      TabIndex        =   0
      Text            =   "rde"
      Top             =   300
      Width           =   3195
   End
   Begin VB.Label Label4 
      Caption         =   "Security database connection string:"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Database:"
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   900
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "User name:"
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   300
      Width           =   975
   End
End
Attribute VB_Name = "frmAPILogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------------------'
' File:         frmAPILogin.frm
' Copyright:    InferMed Ltd. 2004-2006. All Rights Reserved
' Purpose:      Login for MACRO 3.0 API Test Harness
' Author:       NCJ 15 Dec 04 - Copied from LISA Login test
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 9 Aug 06 - We don't use SOAP any more; updated for Nicky's APITEST database
'   NCJ 23 Nov 07 - Added LoginSecurity and options to specify Security DB
'----------------------------------------------------------------------------------------'

Option Explicit

' Nicky's Settings
' MSDE Database on Nicky's PC for testing API
Private Const msMY_DB = "APITEST"

' Security Database on Nicky's PC for testing API LoginSecurity
Private Const msMY_SECDB = "PROVIDER=SQLOLEDB;DATA SOURCE=localhost;DATABASE=APISEC;USER ID=sa;PASSWORD=macrotm;"

' Nicky's Oracle database
'Private Const msMY_SECDB = "PROVIDER=MSDAORA;DATA SOURCE=HARDB_IMEDDB;USER ID=NICKYSEC;PASSWORD=NICKYSEC;"
'Private Const msMY_DB = "NICKYSEC"

Private Const msMY_NAME = "n"
Private Const msMY_PWD = "n"

' David's Settings
'Private Const msMY_DB = "MACRODPH3"
'Private Const msMY_NAME = "rde"
'Private Const msMY_PWD = "macrotm"
'Private Const msMY_SECDB = "PROVIDER=SQLOLEDB;DATA SOURCE=localhost;DATABASE=MainSecurity;USER ID=sa;PASSWORD=macrotm;"

Private msMsg As String
Private msUserFull As String
Private msSerialUser As String
Private mlLoginResult  As Long
Private msSecurityDB As String

'----------------------------------------------------------------
Public Function Display(ByRef sMsg As String, _
                        ByRef sUserFull As String, _
                        ByRef sUserState As String, _
                        ByRef sSecurityDb As String) As Long
'----------------------------------------------------------------

    mlLoginResult = -1
    
    Me.Top = (Screen.Height - Me.Height) \ 2
    Me.Left = (Screen.Width - Me.Width) \ 2
    
    txtDB.Text = msMY_DB
    txtSecDB.Text = msMY_SECDB
    txtUser.Text = msMY_NAME
    txtPwd.Text = msMY_PWD
    
    msSecurityDB = ""
    
    Me.Show vbModal
    Display = mlLoginResult
    If mlLoginResult = -1 Then
        sMsg = "Login cancelled"
    Else
        sMsg = msMsg
    End If
    sUserFull = msUserFull
    sUserState = msSerialUser
    sSecurityDb = msSecurityDB
    
End Function

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdLogin_Click()
Dim oAPI As MACROAPI
Dim sSecCon As String

    ' Do MACRO login
'         Go directly to MACRO DLL
    Set oAPI = New MACROAPI
    sSecCon = Trim(txtSecDB.Text)
    If chkSecurity.Value = vbChecked And sSecCon <> "" Then
        ' Remember the specified security DB
        msSecurityDB = sSecCon
        mlLoginResult = oAPI.LoginSecurity(txtUser.Text, txtPwd.Text, txtDB.Text, _
                        "MACROUser", sSecCon, msMsg, msUserFull, msSerialUser)
    Else
        mlLoginResult = oAPI.Login(txtUser.Text, txtPwd.Text, txtDB.Text, _
                        "MACROUser", msMsg, msUserFull, msSerialUser)
    End If
    Set oAPI = Nothing
    
    Unload Me

End Sub

