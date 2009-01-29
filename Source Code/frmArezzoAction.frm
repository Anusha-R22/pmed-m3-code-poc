VERSION 5.00
Begin VB.Form frmArezzoAction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Action"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   ControlBox      =   0   'False
   Icon            =   "frmArezzoAction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H80000004&
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Text            =   "Task Description"
      Top             =   120
      Width           =   6135
   End
   Begin VB.TextBox txtProcedure 
      Height          =   1935
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmArezzoAction.frx":030A
      Top             =   1800
      Width           =   6135
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "The following action is now suggested:"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   6015
   End
End
Attribute VB_Name = "frmArezzoAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------
' File: frmArezzoAction.frm
' Copyright InferMed Ltd 1999-2003 All Rights Reserved
' Author: Nicky Johns, InferMed
' Purpose: Deal with Arezzo actions in MACRO Data Management
'-----------------------------------------
' REVISIONS
'   NCJ 28-29 Sep 99 - Initial Development
'   willc 10/11/99      Added error handlers
' MACRO 2.2
' NCJ 1 Oct 01 - Changed RefreshMe to Display
' NCJ 17 Jan 02 - Added Hourglass suspend/resume (Buglist 2.2.3, Bug 25)
'
'   NCJ 31 Jan 03 - Use oArezzo passed in
'   NCJ 10 Feb 03 - Use new DEBS routine to confirm action
'-----------------------------------------

Option Explicit

Private moAction As TaskInstance
Private mbOKClicked As Boolean

' NCJ 31 Jan 03
Private moArezzo As Arezzo_DM

'-----------------------------------------
Private Sub cmdCancel_Click()
'-----------------------------------------
' They don't want to confirm this action
'-----------------------------------------
        
    Unload Me

End Sub

'-----------------------------------------
Private Sub cmdOK_Click()
'-----------------------------------------
' Click on OK button
' Confirm the action (if it's still requested)
' NCJ 10 Feb 03 - Use new ConfirmAction call in DEBS
'-----------------------------------------

    On Error GoTo ErrHandler

    If moAction.TaskState = "requested" Then
'        moAction.Confirm
        Call moArezzo.ConfirmAction(moAction.TaskKey)
        mbOKClicked = True
    End If
    
    Unload Me
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdOK_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'-----------------------------------------
Public Function Display(oTask As TaskInstance, oArezzo As Arezzo_DM) As Boolean
'-----------------------------------------
' Refresh and display for the given action task
'-----------------------------------------

    On Error GoTo ErrHandler

    Set moArezzo = oArezzo
    
    mbOKClicked = False
    
    Set moAction = oTask
    Me.Caption = "Action - " & moAction.Name
    txtDesc.Text = moAction.Description
    txtProcedure.Text = moAction.Procedure & vbCrLf & moAction.Context
    
    ' NCJ 17 Jan 02 - Suspend/resume hourglass
    HourglassSuspend
    Me.Show vbModal
    HourglassResume
    
    Display = mbOKClicked
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Display")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Function

'------------------------------------------
Private Sub Form_Load()
'------------------------------------------
' Place the form centrally
'------------------------------------------

    On Error GoTo ErrHandler

    Me.BackColor = glFormColour
    Me.Top = (Screen.Height - Me.Height) \ 2
    Me.Left = (Screen.Width - Me.Width) \ 2
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshMe")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------
' Tidy up
'---------------------------------------------------------
    
    Set moAction = Nothing
    Set moArezzo = Nothing

End Sub
