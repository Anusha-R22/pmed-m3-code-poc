VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAccessConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Access Database Connection"
   ClientHeight    =   1230
   ClientLeft      =   3990
   ClientTop       =   3645
   ClientWidth     =   7425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   7425
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   4440
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   780
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   780
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   300
      Width           =   1215
   End
   Begin VB.TextBox txtDatabaseLocation 
      Height          =   345
      Left            =   100
      TabIndex        =   0
      Top             =   300
      Width           =   5930
   End
   Begin VB.Label Label2 
      Caption         =   "Database Password          (If required)"
      Height          =   435
      Left            =   2940
      TabIndex        =   5
      Top             =   780
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Database location"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   1395
   End
End
Attribute VB_Name = "frmAccessConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       frmAccessConnect.frm
'   Author:     Mo Morris, Jan 5 2000
'   Purpose:    Only called from frmDataDefinition for locating an
'               Access database that is to be used for extracting
'               category codes and values
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------
'Revisions:
'   TA 10/10/01: Changed provider to 4.0
'---------------------------------------------------------------------

'---------------------------------------------------------------------
Private Sub cmdBrowse_Click()
'---------------------------------------------------------------------

    On Error Resume Next
    CommonDialog1.DialogTitle = "Access DataBase Location"
    CommonDialog1.DefaultExt = "mdb"
    CommonDialog1.Filter = "Access database (*.mdb)|*.mdb"
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.FileName = txtDatabaseLocation.Text
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen
    
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        On Error GoTo 0
    End If
    
    txtDatabaseLocation.Text = CommonDialog1.FileName

End Sub

'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------

    Me.Hide

End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------

    Me.Hide
    frmDataDefinition.txtConnect = "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
        "DATA SOURCE=" & txtDatabaseLocation.Text & ";" & _
        "USER ID=admin;JET OLEDB:DATABASE PASSWORD=" & txtPassword.Text & ";"
    
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenu.Icon
End Sub
