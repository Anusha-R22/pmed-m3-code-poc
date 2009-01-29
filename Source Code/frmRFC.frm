VERSION 5.00
Begin VB.Form frmRFC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reason for change"
   ClientHeight    =   1725
   ClientLeft      =   7170
   ClientTop       =   8880
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5280
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cboRFC 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label lblRFC 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       frmRFC.frm
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Zulfiqar Ahmed, August 2001
'   Purpose:    Reason for change (RFC) window  for Data Management in MACRO 2.2.
'----------------------------------------------------------------------------------------'
'Revisions:
' ASH 09/05/2002 Added code to display macro Logo
' MLM 20/05/02  Clear RFC when Cancel is clicked.
'               Disable OK button when no RFC specified.
' MLM 05/06/02: Added Form_Activate()
'-----------------------------------------------------------------------------------------
Option Explicit
Private msRFC As String

'---------------------------------------------------------------------
Public Sub DisplayRFC(colRFC As Collection, oCoord As clsCoords)
'---------------------------------------------------------------------
'Synopsis: This function iterates through the RFC collection and add
'          them into the combo box before displaying the form
'Input   : Reason for change (RFC) collection
'Output  : None
'---------------------------------------------------------------------
Dim iCount As Integer

    cboRFC.Clear
    

     Me.Top = oCoord.Row
     Me.Left = oCoord.Col
    
    For iCount = 1 To colRFC.Count
        cboRFC.AddItem colRFC.Item(iCount)
    Next iCount
    
    cboRFC.ListIndex = -1
    
    'MLM 20/05/02: Disable OK button until RFC is provided
    cmdOK.Enabled = False
        
    'ensure we appear completely on the screen
    If Me.Top < 100 Then
        Me.Top = 100
    End If
    If Me.Top + Me.Height > Screen.Height - 500 Then
        Me.Top = Screen.Height - Me.Height - 500
    End If
    If Me.Left < 100 Then
        Me.Left = 100
    End If
    If Me.Left + Me.Width > Screen.Width Then
        Me.Left = Screen.Width - Me.Width
    End If
    
    Me.Show vbModal
    
End Sub

'-------------------------------------------------------------------------------------------------------
Private Sub cboRFC_Change()
'-------------------------------------------------------------------------------------------------------
' MLM 20/05/02: Added. Enable the OK button based on whether an RFC has been provided.
'-------------------------------------------------------------------------------------------------------

    cmdOK.Enabled = (Trim(cboRFC.Text) <> "")

End Sub

'-------------------------------------------------------------------------------------------------------
Private Sub cboRFC_Click()
'-------------------------------------------------------------------------------------------------------
' MLM 20/05/02: Added. Enable the OK button based on whether an RFC has been provided.
'-------------------------------------------------------------------------------------------------------

    cmdOK.Enabled = (Trim(cboRFC.Text) <> "")

End Sub

'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------
' MLM 20/05/02: Current build buglist, 2.2.10 no. 4: Remove RFC when Cancel is pressed.
'---------------------------------------------------------------------

    cboRFC.Text = ""
    Me.Hide

End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
'---------------------------------------------------------------------
    
    Me.Hide

End Sub

'---------------------------------------------------------------------
Public Function GetRFC() As String
'---------------------------------------------------------------------
'Synopsis: This function stores the value of selected by the user from
'          the reason for change (RFC) combo box
'Input   : None
'Output  : Selected value of reason for change (RFC)
'---------------------------------------------------------------------
        
        GetRFC = Trim(cboRFC.Text)
    
End Function

'---------------------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------------------
' MLM 05/06/02: When form is shown, set focus to cboRFC so that user can type immediately.
'---------------------------------------------------------------------

    cboRFC.SetFocus

End Sub

'----------------------------------------------------------------------
Private Sub Form_Load()
'----------------------------------------------------------------------
' ASH Added 9/05/2002
'----------------------------------------------------------------------
    
    Me.Icon = frmMenu.Icon

End Sub
