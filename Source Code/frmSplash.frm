VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   ClientHeight    =   4800
   ClientLeft      =   6420
   ClientTop       =   4485
   ClientWidth     =   7725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00700000&
      Height          =   675
      Left            =   4155
      TabIndex        =   1
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label lblModule 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00700000&
      Height          =   675
      Left            =   4140
      TabIndex        =   0
      Top             =   900
      Width           =   3045
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   4800
      Left            =   0
      Picture         =   "frmSplash.frx":0442
      Top             =   0
      Width           =   7725
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmSplash.frm
'   Author:     Andrew Newbigging, June 1997, Updated Jan 2003
'   Purpose:    Splash form used throughout Macro.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:

'------------------------------------------------------------------------------------'
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------

    FormCentre Me

    FormatLabels

End Sub

'---------------------------------------------------------------------
Private Sub Frame1_Click()
'---------------------------------------------------------------------

    Unload Me

End Sub

Private Sub FormatLabels()
'-------------------------------------------------------------------
'-------------------------------------------------------------------


    On Error GoTo ErrorHandler
    
    lblModule.Caption = Replace(GetApplicationTitle, "MACRO ", "")
    lblVersion.Caption = "Version" & " " & App.Major & "." & App.Minor

Exit Sub
ErrorHandler:
Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "frmSplash.FormatLabels", Err.Source)
   
    Case OnErrorAction.Retry
        Resume
    Case OnErrorAction.QuitMACRO
        Call ExitMACRO
    End Select

End Sub

