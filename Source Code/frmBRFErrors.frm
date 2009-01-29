VERSION 5.00
Begin VB.Form frmBRFErrors 
   Caption         =   "Batch Response File Errors"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5000
   ScaleMode       =   0  'User
   ScaleWidth      =   8000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   6600
      TabIndex        =   1
      Top             =   4140
      Width           =   1200
   End
   Begin VB.TextBox txtErrors 
      Height          =   3825
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   7620
   End
End
Attribute VB_Name = "frmBRFErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         frmBRFErrors.frm
' Copyright:    InferMed Ltd. 2000. All Rights Reserved
' Author:       Mo Morris, September 2002
' Purpose:      USED to display Batch Respone File (BRF) errors
'----------------------------------------------------------------------------------------'
'   Revisions:
'----------------------------------------------------------------------------------------'

Option Explicit

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Unload Me

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdOK_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon
    
    'set initial size of form
    Me.Width = gnMINFORMWIDTHERRORSFORM
    Me.Height = gnMINFORMHEIGHTERRORSFORM

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Form_Load")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------
Dim lFormWidth As Long
Dim lFormHeight As Long

    On Error GoTo ErrHandler

    'check that form has not been minimised
    If Me.WindowState = vbMinimized Then Exit Sub
     
    If Me.Width < gnMINFORMWIDTHERRORSFORM Then
        Me.Width = gnMINFORMWIDTHERRORSFORM
    End If

    If Me.Height < gnMINFORMHEIGHTERRORSFORM Then
        Me.Height = gnMINFORMHEIGHTERRORSFORM
    End If
    
    lFormWidth = Me.ScaleWidth
    lFormHeight = Me.ScaleHeight

    txtErrors.Left = 100
    txtErrors.Top = 100
    txtErrors.Width = lFormWidth - 200
    txtErrors.Height = lFormHeight - cmdOK.Height - 300
    
    cmdOK.Left = lFormWidth - cmdOK.Width - 100
    cmdOK.Top = lFormHeight - cmdOK.Height - 100
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Form_Resize")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub
