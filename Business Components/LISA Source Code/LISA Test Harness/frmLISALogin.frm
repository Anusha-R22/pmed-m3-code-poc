VERSION 5.00
Begin VB.Form frmLISALogin 
   Caption         =   "LISA Log in"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Soapiness"
      Height          =   615
      Left            =   180
      TabIndex        =   8
      Top             =   2340
      Width           =   4275
      Begin VB.OptionButton optDLL 
         Caption         =   "Use DLL directly"
         Height          =   315
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   2115
      End
      Begin VB.OptionButton optSoap 
         Caption         =   "Use SOAP"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   180
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Log in"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtDB 
      Height          =   345
      Left            =   1260
      TabIndex        =   2
      Text            =   "CRUKTEST"
      Top             =   1380
      Width           =   3195
   End
   Begin VB.TextBox txtPwd 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "n"
      Top             =   840
      Width           =   3195
   End
   Begin VB.TextBox txtUser 
      Height          =   345
      Left            =   1260
      TabIndex        =   0
      Text            =   "n"
      Top             =   300
      Width           =   3195
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
Attribute VB_Name = "frmLISALogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'   frmLISALogin.frm
'   NCJ 22 Jan 07 - Testing against MACRO 3.0.76

Option Explicit
' MSDE Database on Nicky's PC
' Private Const msMY_DB = "CRUKTEST"
' Oracle database on IMED3 (accessed from DEV1 machine)
'Private Const msMY_DB = "ora_lisa1"
Private Const msMY_DB = "cruk"      ' 22 Jan 07
'Private Const msMY_DB = "NORA1"      ' 23 Jan 07 - David's machine

Private msMsg As String
Private msUserFull As String
Private msSerialUser As String
Private mlLoginResult  As Long
Private mbSoapy As Boolean

Public Function Display(ByRef sMsg As String, _
                        ByRef sUserFull As String, _
                        ByRef sUserState As String, _
                        ByRef bSoapy As Boolean) As Long

    mlLoginResult = -1
    
    Me.Top = (Screen.Height - Me.Height) \ 2
    Me.Left = (Screen.Width - Me.Width) \ 2
    
    txtDB.Text = msMY_DB
    optSoap.Value = False
    optDLL.Value = True
    mbSoapy = optSoap.Value
    
    Me.Show vbModal
    Display = mlLoginResult
    If mlLoginResult = -1 Then
        sMsg = "Login cancelled"
    Else
        sMsg = msMsg
    End If
    sUserFull = msUserFull
    sUserState = msSerialUser
    bSoapy = mbSoapy
    
End Function

Private Sub cmdCancel_Click()

    mbSoapy = optSoap.Value
    Unload Me

End Sub

Private Sub cmdLogin_Click()
Dim oSoapClient As SoapClient30
Dim oLISA As MACROLISA

    ' Do LISA login
    If optSoap.Value = True Then
        ' Use SOAP
        Set oSoapClient = New SoapClient30
        Call oSoapClient.MSSoapInit(frmMenu.gsSOAP)
        mlLoginResult = oSoapClient.Login(txtUser.Text, txtPwd.Text, txtDB.Text, "MACROUser", msMsg, msUserFull, msSerialUser)
        Set oSoapClient = Nothing
    Else
        ' Go directly to LISA DLL
        Set oLISA = New MACROLISA
        mlLoginResult = oLISA.Login(txtUser.Text, txtPwd.Text, txtDB.Text, "MACROUser", msMsg, msUserFull, msSerialUser)
        Set oLISA = Nothing
    End If
    
    mbSoapy = optSoap.Value
    Unload Me

End Sub

