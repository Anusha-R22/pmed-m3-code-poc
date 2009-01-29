VERSION 5.00
Begin VB.Form frmTimeOutSplash 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5820
   ClientLeft      =   1155
   ClientTop       =   3990
   ClientWidth     =   8655
   FillColor       =   &H00800000&
   ForeColor       =   &H00800000&
   Icon            =   "FrmTimeOutSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5788.626
   ScaleMode       =   0  'User
   ScaleWidth      =   8623.924
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Label lblText 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Please re-enter your password to resume working with MACRO."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00700000&
      Height          =   1575
      Left            =   1980
      TabIndex        =   0
      Top             =   1020
      Width           =   4515
   End
   Begin VB.Image imgBG 
      Height          =   5580
      Left            =   0
      Picture         =   "FrmTimeOutSplash.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8640
   End
End
Attribute VB_Name = "frmTimeOutSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       frmTimeOutSplash.frm
'   Author:     Will Casey, March 30 2000
'   Purpose:    Hides Patient information after timeout.
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------
'   Revisions:
'    NCJ 1 Apr 2000 - Changed calls to frmTimeOutLogin
'    ZA 14/06/2002 - Fixed bug 14 in build 2.2.13
'   TA 20/03/2003 - changed to new icon
'   TA 28/05/2003: If requested , propmpt user when eXit MACRO clicked
'---------------------------------------------------------------------
Option Explicit

Option Base 0
Option Compare Binary

'do we prompt user if exit clocked
Private mbPromptIfExitClicked As Boolean

Private mbContinueApp As Boolean


Private Sub Form_Activate()

    If (gbIsActiveDirectoryLogin) Then
        mbContinueApp = True
    Else
        mbContinueApp = frmTimeOutLogin.Display(mbPromptIfExitClicked)
    End If
    Unload Me
End Sub

'---------------------------------------------------------------------
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Show login form when user moves the mouse
'---------------------------------------------------------------------
    
 '   mbContinueApp = frmTimeOutLogin.Display
    
  '  Unload Me
    
End Sub


Public Function Display(bPromptIfExitClicked As Boolean) As Boolean
'---------------------------------------------------------------------
' display splash screen
'   Output:
'           function - continue/shutdown app
'---------------------------------------------------------------------
    
    'Me.Icon = frmMenu.Icon
    
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    mbContinueApp = False
    
    mbPromptIfExitClicked = bPromptIfExitClicked
    
    'ZA 14/06/2002 - restore window state to avoid focus problem
    Me.WindowState = vbMaximized
    Me.Show vbModal
    
    Display = mbContinueApp

End Function

Private Sub Form_Resize()

    imgBG.Width = Me.Width
    imgBG.Height = Me.Height
    lblText.Top = 1000 'Me.ScaleHeight / 2 - lblText.Height
    lblText.Left = Me.ScaleWidth / 2 - lblText.Width / 2
    
    
End Sub
