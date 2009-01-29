VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHourglass 
   BackColor       =   &H00DFDFDF&
   BorderStyle     =   0  'None
   ClientHeight    =   1470
   ClientLeft      =   1380
   ClientTop       =   2235
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBackground 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   0
      ScaleHeight     =   2205
      ScaleWidth      =   5505
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin MSComCtl2.Animation Animation1 
         Height          =   315
         Left            =   300
         TabIndex        =   3
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483643
         FullWidth       =   17
         FullHeight      =   21
      End
      Begin VB.Timer Timer1 
         Left            =   3960
         Top             =   180
      End
      Begin VB.Label lblMsg 
         Caption         =   "lblmsg"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label lblWait 
         BackColor       =   &H80000008&
         BackStyle       =   0  'Transparent
         Caption         =   "lblWait"
         Height          =   315
         Left            =   300
         TabIndex        =   1
         Top             =   480
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmHourglass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       frmHourglass.frm
'   Copyright:  InferMed Ltd. 2002-2003. All Rights Reserved
'   Author:     Toby Aldridge, September 2002
'   Purpose:    Hourglass form
'----------------------------------------------------------------------------------------'
' REVISIONS
' NCJ 17 Feb 03 - Call Refresh rather than DoEvents in Display
'----------------------------------------------------------------------------------------'

Option Explicit


'----------------------------------------------------------------------------------------'
Private Sub Form_Load()
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
Dim sFileName As String

    WebFormsEnabled = False

    With picBackground
        .Top = 0
        .Left = 0
        .Width = Me.Width - 20
        .Height = Me.Height - 20
        .BackColor = &HDFDFDF
    End With
    
    lblWait.Font = "verdana"
    lblWait.Font.SIZE = 9
    lblWait.Caption = "please wait"
    lblWait.BackColor = &HDFDFDF
    
    lblMsg.Font = "verdana"
    lblMsg.Font.SIZE = 9
    lblMsg.Caption = "message"
    lblMsg.BackColor = &HDFDFDF
        
    Timer1.Interval = 50
    
    sFileName = App.Path & "\Clock.avi"
    Animation1.Open sFileName
    Animation1.Play
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function Display(sMessage As String, bTop As Boolean)
'----------------------------------------------------------------------------------------'
Dim ofrm As frmBorder

    
    Load Me
    Me.BorderStyle = vbBSNone
    If bTop Then
        Set ofrm = frmMenu.mofrmBorderTop
    Else
        Set ofrm = frmMenu.moFrmBorderBottom
    End If
    Me.Left = frmMenu.Left + ofrm.Left + (ofrm.Width / 2) - Me.Width / 2
    Me.Top = frmMenu.Top + ofrm.Top + (ofrm.Height / 2) - Me.Height / 2
    
    lblMsg.Caption = sMessage & "..."
    Me.Show vbModeless
    ' NCJ 17 Feb 03 - Call Refresh rather than DoEvents
    ' to prevent spurious events being generated in frmEFormDataEntry
    Me.Refresh
'    DoEvents
    Timer1.Enabled = True

End Function

'----------------------------------------------------------------------------------------'
Private Sub Form_Unload(Cancel As Integer)
'----------------------------------------------------------------------------------------'

    Timer1.Enabled = False
    
    WebFormsEnabled = True

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Timer1_Timer()
'----------------------------------------------------------------------------------------'
'keep on top
'----------------------------------------------------------------------------------------'

    Me.ZOrder

End Sub
