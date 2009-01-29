VERSION 5.00
Begin VB.Form frmMDAC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MDAC Installation for MACRO"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDCOM95 
      Caption         =   "Run &DCOM95.EXE"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdMDAC 
      Caption         =   "Run &MDAC_TYP.EXE"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmMDAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       frmMDAC.frm
'   Author:     Stephen Morris, Jan 2000
'   Purpose:    Form used to trap ADO installation error in MACRO 2.0.
'               Displays an error report offering the user
'               the option of running MDAC_TYP.EXE
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'Revisions:
' NCJ 17 Jan 00 - Added button and info about DCOM95 for Win95 users
'
'----------------------------------------------------------------------------------------'

Option Explicit

'----------------------------------------------------------------------------------------'
Private Sub cmdDCOM95_Click()
'----------------------------------------------------------------------------------------'
' Run DCOM95.exe (expecting it to be found in the application directory)
' Use /Q switch for "quiet mode"
'----------------------------------------------------------------------------------------'
Dim nReturn As Long

    nReturn = Shell(App.Path & "\DCOM95 /Q", vbNormalFocus)
    Call ExitMACRO
    Call MACROEnd
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdMDAC_Click()
'----------------------------------------------------------------------------------------'
' Run MDAC_TYP.exe (expecting it to be found in the application directory)
' Use /Q switch for "quiet mode"
'----------------------------------------------------------------------------------------'
Dim nReturn As Long

    nReturn = Shell(App.Path & "\MDAC_TYP.exe /Q", vbNormalFocus)
    Call ExitMACRO
    Call MACROEnd
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'----------------------------------------------------------------------------------------'
' Just quit
'----------------------------------------------------------------------------------------'
    
    Call ExitMACRO
    Call MACROEnd

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Load()
'----------------------------------------------------------------------------------------'
' Show the message about MDAC
' NCJ 17/1/00 - Added info about DCOM95 for Windows95 users
'----------------------------------------------------------------------------------------'
Dim sMsg As String

    Me.Icon = frmMenu.Icon
    sMsg = "An error has occurred while trying to start MACRO, "
    sMsg = sMsg & "which may mean that you need some extra system files on your computer."
    sMsg = sMsg & vbCrLf & "Please try running the program 'MDAC_TYP.EXE' in your MACRO installation directory."
    sMsg = sMsg & vbCrLf & "(You are then advised to restart your computer before trying to run MACRO again)."
    
    ' NCJ 17/1/00 Deal with Windows95
    If IsWin95 Then
        sMsg = sMsg & vbCrLf & vbCrLf
        sMsg = sMsg & "Note that if you are using Windows 95, you may need to first run 'DCOM95.EXE' before continuing. "
        cmdDCOM95.Visible = True
    Else
        cmdDCOM95.Visible = False
    End If
    
    lblWarning.Caption = sMsg
    
End Sub
