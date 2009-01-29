VERSION 5.00
Begin VB.Form frmPasswords 
   Caption         =   "Password Testing"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      Height          =   1095
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1920
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset Password"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtNewPwd 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   3975
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Result:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "New password:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "User name:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmPasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' File: frmPasswords.frm
' Nicky Johns, InferMed, Feb 2008
' Testing API Routines

Option Explicit

Private msSecCon As String
Private msSerialUser As String

'----------------------------------------------
Public Sub Display(sSerialUser As String, sSecCon As String)
'----------------------------------------------
' Display password testing form,
' using given Security connection (may be "", to use default sec DB)
' and serialised user
'----------------------------------------------

    msSerialUser = sSerialUser
    msSecCon = sSecCon
    Me.Show vbModal

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

'----------------------------------------------
Private Sub cmdReset_Click()
'----------------------------------------------
' Reset another user's password
'----------------------------------------------
Dim sUser As String
Dim sPwd As String
Dim sMsg As String
Dim oAPI As MACROAPI
Dim lResult As Long

    sUser = Trim(txtUserName.Text)
    sPwd = Trim(txtNewPwd.Text)
    
    If (sUser <> "" And sPwd <> "") Then
        Set oAPI = New MACROAPI
        lResult = oAPI.ResetPassword(msSerialUser, sUser, sPwd, sMsg)
        Set oAPI = Nothing
        txtResult.Text = "Result = " & lResult & vbCrLf _
                & sMsg
    End If
    
End Sub

