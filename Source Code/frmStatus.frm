VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1275
   ClientLeft      =   8565
   ClientTop       =   6735
   ClientWidth     =   5445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Wingdings"
      Size            =   8.25
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   5445
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   60
      Top             =   780
   End
   Begin MSComctlLib.ProgressBar prgProgress 
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   480
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblPercent 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblMove 
      Caption         =   "Ç"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   4
      Top             =   420
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMove 
      Caption         =   "Ã"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5100
      TabIndex        =   3
      Top             =   420
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMove 
      Caption         =   "Ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5100
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMove 
      Caption         =   "Æ"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMessage 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4635
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998-2000. All Rights Reserved
'   File:       frmStatus.frm
'   Author:     Toby Aldridge: October 2000
'   Purpose:    Form to 'reassure' user that something is happening
'
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------

'has user pressed cancel?
Private mbCancel As Boolean

Option Explicit

Public Sub Start(sCaption As String, sMessage As String, Optional bAnimate = True, _
                         Optional nMax As Integer = 0, Optional bCancel = False)
'--------------------------------------------------------------------------------
'Display the form with the Caption and message
' optionally display a naff animation and progress bar
'   optionally display a canel button though this does nothing yet
'--------------------------------------------------------------------------------


    mbCancel = False
    
    tmrMove.Interval = 200
    Me.Icon = frmMenu.Icon
    FormCentre Me
    
    'set form title and message
    Me.Caption = sCaption
    Message = sMessage
    
    'do they want animation
    If Not bAnimate Then
        Me.Width = Me.Width - (2 * lblMove(0).Width) - 220
    End If
    
    'show cancel button if required
    If bCancel Then
        'display in middle of form
        cmdCancel.Left = (Me.ScaleWidth - cmdCancel.Width) / 2
        'show arrowhourglass
        MousePointerChange vbArrowHourglass
    Else
        cmdCancel.Visible = False
        Me.Height = Me.Height - cmdCancel.Height - 120
        'show hourglass
        MousePointerChange vbHourglass
    End If
    
    'show progress bar if required
    If nMax = 0 Then
        'move message down
        lblMessage.Top = lblMessage.Top + (prgProgress.Height + 60) / 2
        prgProgress.Visible = False
        lblPercent.Visible = False
    Else
        'set max and min
        prgProgress.Min = 0
        prgProgress.Max = nMax
        lblPercent.Caption = ""
    End If
            
    
    Me.Show vbModeless
    'Me.Show vbModal
    
    'finally start animation if wanted
    tmrMove.Enabled = bAnimate

End Sub

'--------------------------------------------------------------------------------
Public Property Get Cancel() As Boolean
'--------------------------------------------------------------------------------
' read only property to return whether user has pressed cancel
'--------------------------------------------------------------------------------

    Cancel = mbCancel

End Property


'--------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------------------

    Message = "Cancelling at user's request..."
    mbCancel = True
    
End Sub

Private Sub tmrMove_Timer()
'--------------------------------------------------------------------------------
' keep the form display changing - rotation effect
'--------------------------------------------------------------------------------
Static nCounter As Integer
Dim i As Integer
    tmrMove.Enabled = False
    nCounter = nCounter + 1
    If nCounter = 4 Then
        nCounter = 0
    End If
    
    For i = 0 To 3
        lblMove(i).Visible = (nCounter = i)
    Next
    
    Me.Refresh
    
    tmrMove.Enabled = True

End Sub

Public Sub Finish()
'--------------------------------------------------------------------------------
' unload form
'--------------------------------------------------------------------------------
    
    tmrMove.Enabled = False
    'lose hourglass or arrowhourglass pointer
    MousePointerRestore
    Unload Me
    
End Sub

'--------------------------------------------------------------------------------
Public Property Let Message(sMessage As String)
'--------------------------------------------------------------------------------
' change the message
'--------------------------------------------------------------------------------

    lblMessage.Caption = "  " & sMessage
    Me.Refresh
    
End Property

Public Property Let Progress(nValue As Integer)
'--------------------------------------------------------------------------------
' change the progress bar if valid number and proress bar is visible
'--------------------------------------------------------------------------------
    
    With prgProgress
        If .Visible And (nValue >= .Min) And (nValue <= .Max) Then
            .Value = nValue
            lblPercent.Caption = Int(100 * (nValue / .Max)) & "%"
            Me.Refresh
        End If
    End With
    
End Property

'--------------------------------------------------------------------------------
Public Sub Status(sMessage As String, nValue As Integer)
'--------------------------------------------------------------------------------
' change the message and progress bar
'--------------------------------------------------------------------------------

    Message = sMessage
    Progress = nValue
    
End Sub
