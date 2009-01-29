VERSION 5.00
Begin VB.Form frmImageViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image viewer"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2835
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmImageViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmImageViewer.frm
'   Author:     Andrew Newbigging, October 1997
'   Purpose:    Simple form to display gif, jpeg, bmp, wmf images
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
 '  1       Andrew Newbigging       29/10/97
'   2       Andrew Newbigging       21/01/98
'   3       Andrew Newbigging       02/04/98
'   4       Andrew Newbigging       8/5/99
'           Modified to set a minimum size for the image
'   Mo Morris   2/3/00  ShowImage re-written SR3140
'------------------------------------------------------------------------------------'

Option Explicit

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
    Me.BackColor = glFormColour
    
    'TA 2/11/01 Use menu icon
    Me.Icon = frmMenu.Icon
    
End Sub


'---------------------------------------------------------------------
Public Sub ShowImage(ByVal vImageFile As String)
'---------------------------------------------------------------------

    'MLM 18/06/03: If there is an error loading the image, give a friendly dialog rather than an RTE.
    On Error Resume Next
    Set picImage.Picture = LoadPicture(vImageFile)
    If Err.Number <> 0 Then
        DialogError "The attachment cannot be viewed due to the following error:" & vbCrLf & _
            Err.Number & ": " & Err.Description & "." & vbCrLf & _
            "Please contact your MACRO systems administrator."
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    picImage.Top = 50
    picImage.Left = 50

    'Changed Mo Morris 2/3/00
    Me.Height = picImage.Height + 100 + (Me.Height - Me.ScaleHeight)
    Me.Width = picImage.Width + 100 + (Me.Width - Me.ScaleWidth)
    
    'force the window to be a minimum of 2000 * 2000 twips
    If Me.Height < 2000 Then
        Me.Height = 2000
    End If
    If Me.Width < 2000 Then
        Me.Width = 2000
    End If
    
    'force the window to be a maximum of 4/5th screen height & width
    If Me.Height > (Screen.Height * 0.8) Then
        Me.Height = (Screen.Height * 0.8)
    End If
    If Me.Width > (Screen.Width * 0.8) Then
        Me.Width = (Screen.Width * 0.8)
    End If
    
    Me.Show
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ShowImage")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub
