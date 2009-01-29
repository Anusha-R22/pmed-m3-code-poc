VERSION 5.00
Begin VB.Form frmInputBox 
   ClientHeight    =   1395
   ClientLeft      =   9240
   ClientTop       =   5850
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   5340
      TabIndex        =   3
      Top             =   960
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   4080
      TabIndex        =   2
      Top             =   960
      Width           =   1140
   End
   Begin VB.Frame fraInput 
      Height          =   795
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6435
      Begin VB.TextBox txtAnswer 
         Height          =   375
         Left            =   120
         MaxLength       =   2000
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   6195
      End
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       frmInputBox.frm
'   Author:     Toby Aldridge April 2000
'   Purpose:    Generic Input Box
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'ASH 10/07/2002 - CBB 2.2.19 no.20 do not allow enter key in singleline mode
'TA 28/04/2003: disallow pasted in invalid chars - (`~ and pipe)
'-----------------------------------------------------------------------------------------

Option Explicit

Private Const mlMULTI_HEIGHT_INC = 1000

Private mbOK  As Boolean
Private msAnswer As String
Private mnValScheme As Integer
'ASH 10/07/2002 CBB 2.2.19 no.20
Private mbMultiline  As Boolean

'---------------------------------------------------------------------
Public Function Display(sTitle As String, sCaption As String, sAnswer As String, Optional bAllowNull As Boolean = False, Optional bMultiline = False, Optional bOKDefault As Boolean = True, Optional nValScheme As Integer = -1, Optional nMaxLength As Integer = 0) As Boolean
'---------------------------------------------------------------------
'   display form
'Input:
'       sTitle - form title
'       sCaption = frame caption
'       bAllowNull - enable OK button when nothing empty?
'Output:
'       function - OK pressed?
'---------------------------------------------------------------------
    mbOK = False
    
    Me.Caption = sTitle
    fraInput.Caption = sCaption
    txtAnswer.Text = sAnswer
    txtAnswer.SelStart = Len(sAnswer)
    txtAnswer.MaxLength = nMaxLength
    txtAnswer_Change
    mnValScheme = nValScheme
    
    If bMultiline Then
        Me.Height = Me.Height + mlMULTI_HEIGHT_INC
        fraInput.Height = fraInput.Height + mlMULTI_HEIGHT_INC
        txtAnswer.Height = txtAnswer.Height + mlMULTI_HEIGHT_INC
        cmdOK.Top = cmdOK.Top + mlMULTI_HEIGHT_INC
        cmdCancel.Top = cmdOK.Top
        
    End If
    
    cmdOK.Default = bOKDefault
    mbMultiline = bMultiline
   
    FormCentre Me
    Me.Show vbModal
    If mbOK Then
        sAnswer = msAnswer
    End If
    Display = mbOK
    
End Function

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------

    If ValidateTextField(txtAnswer.Text) Then
        msAnswer = txtAnswer.Text
        
        mbOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenu.Icon
End Sub

'---------------------------------------------------------------------
Private Sub txtAnswer_Change()
'---------------------------------------------------------------------

    cmdOK.Enabled = Trim(txtAnswer.Text) <> ""
End Sub

'---------------------------------------------------------------------
Private Sub txtAnswer_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
'Validate character entered
'---------------------------------------------------------------------
    If mnValScheme <> -1 Then
        If Not gblnValidString(Chr(KeyAscii), mnValScheme) Then
            KeyAscii = 0
        End If
    End If
    'ASH 10/07/2002 CBB 2.2.19 no.20 do not allow enter key in singleline mode
    If Not mbMultiline Then
        If KeyAscii = 13 Then
            KeyAscii = 0
        End If
    End If
End Sub


'------------------------------------------------------------------------------------------
Private Function ValidateTextField(ByVal sString As String) As Boolean
'------------------------------------------------------------------------------------------
'ASH 10/07/2002 CBB 2.2.15 no.R24
'Checks if text string does not contain invalid characters such as |,",` etc.
'------------------------------------------------------------------------------------------
On Error GoTo Errlabel

    ValidateTextField = True

    If Not gblnValidString(sString, valOnlySingleQuotes) Then
        DialogInformation " Text" & gsCANNOT_CONTAIN_INVALID_CHARS
        ValidateTextField = False
    End If

Exit Function
Errlabel:
  Err.Raise Err.Number, , Err.Description & "|" & "frmNewMIMessage.ValidateTextField"
End Function
