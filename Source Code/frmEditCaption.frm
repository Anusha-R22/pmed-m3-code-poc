VERSION 5.00
Begin VB.Form frmEditCaption 
   BorderStyle     =   0  'None
   Caption         =   "Edit caption"
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   290
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   290
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   1125
   End
   Begin VB.TextBox txtEditCaption 
      Height          =   1095
      Left            =   16
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   16
      Width           =   3400
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000D&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   1455
      Left            =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmEditCaption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2003. All Rights Reserved
'   File:       frmEditCaption.frm
'   Author:     Matthew Martin, April 2003
'   Purpose:    Allow user to edit captions and comments from frmCRFDesign
'-----------------------------------------------------------------------------------
'Revisions
'   TA  08/08/2005: setting of Font.Charset = 1 to allow eastern european characters. CBD2591.
'-----------------------------------------------------------------------------------

Option Explicit

Dim msCaption As String
Dim mbCancel As Boolean

'-----------------------------------------------------------------------------------
Public Function Display(lblCaption As Label, sCaption As String, sglXOffset As Single, sglYOffset As Single) As Boolean
'-----------------------------------------------------------------------------------
' Display the caption editing form over the label to be edited, and size it
' according to the properties of the text.
'-----------------------------------------------------------------------------------
    
Dim sglLineHeight As Single
Dim sglHeight As Single
Dim sglWidth As Single
    
    mbCancel = False
    
    With lblCaption
        'copy oElement's text and font attributes to txtEditCaption
        txtEditCaption.Text = sCaption
    
        txtEditCaption.FontSize = .FontSize
        txtEditCaption.FontName = .FontName
        txtEditCaption.FontBold = .FontBold
        txtEditCaption.FontItalic = .FontItalic
'   TA  08/08/2005: setting of Font.Charset = 1 to allow eastern european characters. CBD2591.
        txtEditCaption.Font.Charset = 1
        'txtEditCaption.ForeColor = .ForeColor
        
        'copy oElement's font attributes to this form for the purpose of calculating textheight
        Me.FontSize = .FontSize
        Me.FontName = .FontName
        Me.FontBold = .FontBold
        Me.FontItalic = .FontItalic
        sglLineHeight = frmCRFDesign.TextHeight("_")
        
        'ensure a minimum width for txtEditCaption
        'note that the extra 400 twips is to allow space for scroll bar
        If .Width < 3000 Then
            sglWidth = 3400
        'WillC 24/2/2000 SR 2619
        'MLM 03/06/03: 3.0 buglist 1823: Don't allow the edit window to be so wide it goes off the screen.
        ElseIf .Width > Screen.Width - Me.Left Then
            sglWidth = Screen.Width - Me.Left
'        ElseIf .Width > 35000 Then
'            DialogInformation "Your caption is too long. Please choose a smaller font size if you wish " & vbCrLf _
'                  & " to edit it. You may then return the font size to its original setting."
'            Display = False
'            Exit Function
        Else
            sglWidth = .Width + 400
        End If
        
        'check that sglWidth is not beyond the right hand border of the form
    '    If sglLeft + sglWidth > picCRFPage.Width Then
    '        sglWidth = picCRFPage.Width - sglLeft
    '    End If
        
        'ensure a minimum height for txtEditCaption
        If .Height < 100 Then
            'set Height to 2 line + 400
            sglHeight = (sglLineHeight * 2) + 400
        Else
            'increase height by an extra line + 400
            sglHeight = .Height + sglLineHeight + 400
        End If
        
        'in addition to the text box, frmEditCaption needs space for its buttons
        sglHeight = sglHeight + cmdOK.Height
        
        'move the form slightly if doing so will keep it on top of frmCRFDesign
        If sglXOffset - (frmMenu.Left + frmCRFDesign.Left + frmCRFDesign.Width - sglWidth) > 0 Then
            sglXOffset = frmMenu.Left + frmCRFDesign.Left + frmCRFDesign.Width - sglWidth
        End If
        If sglXOffset < frmMenu.Left + frmCRFDesign.Left Then
            sglXOffset = frmMenu.Left + frmCRFDesign.Left
        End If
        If sglYOffset - (frmMenu.Top + frmMenu.Height - frmMenu.ScaleHeight + frmCRFDesign.Top + frmCRFDesign.Height - sglHeight) > 0 Then
            sglYOffset = frmMenu.Top + frmMenu.Height - frmMenu.ScaleHeight + frmCRFDesign.Top + frmCRFDesign.Height - sglHeight
        End If
        If sglYOffset < frmMenu.Top + frmMenu.Height - frmMenu.ScaleHeight + frmCRFDesign.Top Then
            sglYOffset = frmMenu.Top + frmMenu.Height - frmMenu.ScaleHeight + frmCRFDesign.Top
        End If
        
        Me.Move sglXOffset, sglYOffset, sglWidth, sglHeight
        
    End With
    
    Form_Resize
    Me.Show vbModal
    
    If mbCancel Then
        Display = False
    Else
        lblCaption.Caption = msCaption
        Display = True
    End If
    
End Function

'-----------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------

    mbCancel = True
    Me.Hide

End Sub

'-----------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------

    'WillC SR3688 & SR3689 Cutting and pasting erroneously allowed quotes in the caption.
    If Not gblnValidString(txtEditCaption.Text, valOnlySingleQuotes) Then
        DialogInformation "A caption" & gsCANNOT_CONTAIN_INVALID_CHARS
'        ' Reset text to its previous value
'        txtEditCaption.Text = txtEditCaption.Tag
        txtEditCaption.SetFocus
    Else
        msCaption = Trim(txtEditCaption.Text)
        Me.Hide
    End If

End Sub

'-----------------------------------------------------------------------------------
Private Sub Form_Activate()
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------

        txtEditCaption.SelStart = 0
        txtEditCaption.SelLength = Len(txtEditCaption.Text)
'        txtEditCaption.ZOrder
        txtEditCaption.SetFocus

End Sub

'-----------------------------------------------------------------------------------
Private Sub Form_Resize()
'-----------------------------------------------------------------------------------
' txtEditCaption should fill the entire form, except for the 2 buttons in the bottom right corner.
'-----------------------------------------------------------------------------------

Const nBorder As Integer = 30

    shpBorder.Height = Me.ScaleHeight
    shpBorder.Width = Me.ScaleWidth
    
    txtEditCaption.Left = nBorder
    txtEditCaption.Top = nBorder
    txtEditCaption.Height = Me.ScaleHeight - cmdOK.Height - 2 * nBorder
    txtEditCaption.Width = Me.ScaleWidth - 2 * nBorder
    cmdCancel.Top = txtEditCaption.Height + nBorder
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - nBorder
    cmdOK.Top = cmdCancel.Top
    cmdOK.Left = cmdCancel.Left - cmdOK.Width

End Sub

'-----------------------------------------------------------------------------------
Private Sub txtEditCaption_KeyPress(KeyAscii As Integer)
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------

    If Not gblnValidString(Chr(KeyAscii), valOnlySingleQuotes) Then
        DialogInformation "A caption" & gsCANNOT_CONTAIN_INVALID_CHARS
        '  Cancel the keystroke
        KeyAscii = 0
        txtEditCaption.SetFocus
    End If

End Sub
