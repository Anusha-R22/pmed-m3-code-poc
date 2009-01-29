VERSION 5.00
Begin VB.Form frmExpandResponse 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Response"
   ClientHeight    =   2040
   ClientLeft      =   7995
   ClientTop       =   9420
   ClientWidth     =   4200
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtResponse 
      Height          =   1515
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmExpandResponse.frx":0000
      Top             =   420
      Width           =   3975
   End
   Begin VB.Label lcmdCancel 
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3420
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   60
      Width           =   675
   End
   Begin VB.Label lcmdOK 
      BackColor       =   &H80000005&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   60
      Width           =   315
   End
End
Attribute VB_Name = "frmExpandResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000-2003. All Rights Reserved
'   File:       frmExpandResponse.frm
'   Author:     Ronald Schravendeel, 21 Febrary 2003
'   Purpose:    Allow user to view/edit long strings when displaylength property is used
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'TA 05/11/2004: BD 2425 CRM 990 ADjustments to size of textbox so long text easily displayed
'----------------------------------------------------------------------------------------'
Option Explicit

Private msOriginalResponse As String     ' Store Original Response in case user cancels out
Private msNewResponse As String


'----------------------------------------------------------------------------------------'
' Display
'   Displays the form when user has clicked the Expand button. The properties of the
'   associated txtCRFElement are used to setup the form
'----------------------------------------------------------------------------------------'
Public Function Display(oTextBox As TextBox, sCaption As String) As String
'----------------------------------------------------------------------------------------'

    txtResponse.MaxLength = oTextBox.MaxLength
    
    msOriginalResponse = oTextBox.Text
    txtResponse.Text = msOriginalResponse
    
    Me.BackColor = eMACROColour.emcBackground
    
    With lcmdOK
        .BackColor = eMACROColour.emcBackground
        .Font = MACRO_DE_FONT
        .FontSize = 8
        .ForeColor = eMACROColour.emcLinkText
        .MouseIcon = frmImages.CursorHandPoint.MouseIcon
    End With
    With lcmdCancel
        .BackColor = eMACROColour.emcBackground
        .Font = MACRO_DE_FONT
        .FontSize = 8
        .ForeColor = eMACROColour.emcLinkText
        .MouseIcon = frmImages.CursorHandPoint.MouseIcon
    End With
    
    If oTextBox.Locked Then
        lcmdCancel.Visible = False
        lcmdOK.Caption = "Close"
        lcmdOK.Left = lcmdOK.Left - 200
        lcmdOK.Width = lcmdOK.Width + 200
        txtResponse.Locked = True
        txtResponse.BackColor = oTextBox.BackColor
    End If

    
    'TODO: Expand width to fit full string
    Me.Font = txtResponse.Font
    Me.FontBold = txtResponse.FontBold
    Me.FontItalic = txtResponse.FontItalic
    Me.FontName = txtResponse.FontName
    Me.FontSize = txtResponse.FontSize
    Me.Icon = frmMenu.Icon
    Me.Caption = sCaption

    
    Resize

    FormCentre Me, frmMenu
    
    Me.Show vbModal
    
    Display = msNewResponse
    
End Function
'----------------------------------------------------------------------------------------'
Private Sub Resize()
'----------------------------------------------------------------------------------------'
' Resize
'   Resizes the form and controls so that the entire string is visible
'----------------------------------------------------------------------------------------'
Dim lDiff As Long

    lDiff = Me.TextWidth(txtResponse.Text) - txtResponse.Width + 100
    'make the maximum width 25 characters
    lDiff = Min(lDiff, TextWidth(String(14, "W")))
    If lDiff > 0 Then
        Me.Width = Me.Width + lDiff
        txtResponse.Width = txtResponse.Width + lDiff
        lcmdCancel.Left = txtResponse.Left + txtResponse.Width - lcmdCancel.Width - 60
        lcmdOK.Left = lcmdCancel.Left - lcmdOK.Width - 60
    End If
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------------------'
'unload form if OK or cancel pressed
'----------------------------------------------------------------------------------------'

    Select Case KeyCode
    Case vbKeyEscape
        lcmdCancel_Click
    Case vbKeyReturn
        lcmdOK_Click
    End Select
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub lcmdCancel_Click()
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    msNewResponse = msOriginalResponse
    Unload Me
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub lcmdOK_Click()
'----------------------------------------------------------------------------------------'
'set internal variable that is used for return value to the value of the text boc
'----------------------------------------------------------------------------------------'

    msNewResponse = txtResponse.Text
    Unload Me
    
End Sub

Private Sub txtResponse_Change()
    RemoveCRLF
End Sub

Private Sub txtResponse_KeyDown(KeyCode As Integer, Shift As Integer)
    RemoveCRLF
End Sub

Private Sub txtResponse_KeyPress(KeyAscii As Integer)
    RemoveCRLF
End Sub

Private Sub RemoveCRLF()
    txtResponse.Text = Replace(txtResponse.Text, vbCrLf, " ")
End Sub
