VERSION 5.00
Begin VB.Form frmBorder 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3465
   ClientLeft      =   3615
   ClientTop       =   6075
   ClientWidth     =   4155
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3465
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBackGround 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Height          =   2835
      Left            =   480
      ScaleHeight     =   2835
      ScaleWidth      =   3255
      TabIndex        =   0
      Top             =   435
      Width           =   3255
      Begin VB.PictureBox picCircle 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   180
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   3
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox picInnerBackground 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   1875
         Left            =   480
         ScaleHeight     =   1875
         ScaleWidth      =   2355
         TabIndex        =   1
         Top             =   480
         Width           =   2355
      End
      Begin VB.Label lblTitle 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   690
         TabIndex        =   2
         Top             =   60
         Width           =   1365
      End
   End
   Begin VB.PictureBox picLilacCurve 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   60
      ScaleHeight     =   375
      ScaleWidth      =   345
      TabIndex        =   4
      Top             =   -15
      Width           =   345
   End
End
Attribute VB_Name = "frmBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999-2000. All Rights Reserved
'   File:       frmBorder.frm
'   Author:     Toby Aldridge, October 2002
'   Purpose:    Border for the bottom right hand forms
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'Revisions:

Option Explicit

'----------------------------------------------------------------------------------------'
Public Function Display(bShowLilacCurve As Boolean)
'----------------------------------------------------------------------------------------'

    Load Me
    
        
    picLilacCurve.Top = 0
    picLilacCurve.Left = 0
    picLilacCurve.Width = WEB_BORDER_TITLE_HEIGHT '+ WEB_BORDER_HEIGHT
    picLilacCurve.Height = WEB_BORDER_TITLE_HEIGHT '+ WEB_BORDER_WIDTH
    If bShowLilacCurve Then
        picLilacCurve.BackColor = eMACROColour.emcNonWhiteBackGround
    Else
        picLilacCurve.BackColor = eMACROColour.emcBackGround
    End If
    
    picLilacCurve.FillStyle = 0 'solid
    picLilacCurve.FillColor = eMACROColour.emcBackGround
    picLilacCurve.ForeColor = eMACROColour.emcBackGround
    picLilacCurve.AutoRedraw = True
    
    picBackGround.Top = WEB_BORDER_HEIGHT
    picBackGround.Left = WEB_BORDER_WIDTH
    picBackGround.BackColor = eMACROColour.emcTitlebar
    picBackGround.AutoRedraw = True
    
    picCircle.Left = 0
    picCircle.Top = 0
    picCircle.Width = WEB_BORDER_TITLE_HEIGHT
    picCircle.Height = WEB_BORDER_TITLE_HEIGHT
    picCircle.BackColor = eMACROColour.emcBackGround
    picCircle.FillStyle = 0 'solid
    picCircle.FillColor = eMACROColour.emcTitlebar
    picLilacCurve.AutoRedraw = True
    
    lblTitle.Left = picCircle.Width
    lblTitle.Top = 0 'WEB_INNER_BORDER
    lblTitle.Height = WEB_BORDER_TITLE_HEIGHT - WEB_INNER_BORDER
    lblTitle.BackColor = eMACROColour.emcTitlebar
    
    lblTitle.Font = "Verdana"
    lblTitle.FontBold = True
    lblTitle.FontSize = 8
    
    picInnerBackground.Top = WEB_BORDER_TITLE_HEIGHT
    picInnerBackground.Left = WEB_INNER_BORDER
    picInnerBackground.BackColor = eMACROColour.emcBackGround
    
    DrawCurves

    Me.Show vbModeless

End Function

'----------------------------------------------------------------------------------------'
Private Sub Form_Resize()
'----------------------------------------------------------------------------------------'
On Error Resume Next

    picBackGround.Height = Me.Height - WEB_BORDER_HEIGHT
    picBackGround.Width = Me.Width - (2 * WEB_BORDER_WIDTH)
    lblTitle.Width = Me.Width - lblTitle.Left
    
    picInnerBackground.Height = picBackGround.Height - WEB_BORDER_TITLE_HEIGHT - WEB_INNER_BORDER - (2 * WEB_INNER_BORDER)
    picInnerBackground.Width = picBackGround.Width - (2 * WEB_INNER_BORDER)
    
    DrawCurves
    

    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub DrawCurves()
'----------------------------------------------------------------------------------------'
'Draw the two curves
'----------------------------------------------------------------------------------------'

        picCircle.Circle (picCircle.Width, picCircle.Width), picCircle.Width
        picLilacCurve.Circle (picLilacCurve.Height, picLilacCurve.Height), picLilacCurve.Height

End Sub

'----------------------------------------------------------------------------------------'
Public Sub SetBorderCaption(enWinForm As eWinForms)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
Dim sText As String

    Select Case enWinForm
    Case wfHome: sText = "Home"
    Case wfNewSubject: sText = "New subject"
    Case wfOpenSubject: sText = "Open subject"
    Case wfSubject: sText = "Subject"
    Case wfDiscepancies: sText = "Discrepancies"
    Case wfNotes: sText = "Notes"
    Case wfSDV: sText = "SDV"
    Case wfDataBrowser: sText = "Data browser"
    End Select

    lblTitle.Caption = sText

End Sub

