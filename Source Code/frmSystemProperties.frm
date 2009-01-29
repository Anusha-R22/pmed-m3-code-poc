VERSION 5.00
Begin VB.Form frmSystemProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Properties"
   ClientHeight    =   1980
   ClientLeft      =   4830
   ClientTop       =   5955
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraProperties 
      Caption         =   "System idle timeout (minutes)"
      Height          =   1335
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   2835
      Begin VB.TextBox txtIdleTimeout 
         Height          =   285
         Left            =   900
         TabIndex        =   0
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lblTimeoutLimits 
         Alignment       =   2  'Center
         Caption         =   "This caption displays the limits of the timeout at run-time"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1500
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1500
      Width           =   1215
   End
End
Attribute VB_Name = "frmSystemProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       frmSystemProperties.frm
'   Author:     Andrew Newbiggin, 17/09/99
'   Purpose:    Maintain users and the databases which a user can access.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   PN  04/10/99    Amended txtIdleTimeout_Change() to disallow null from being saved
'   WILLC 11/10/99  Added the error handlers
'   TA 26/04/2000   Subclassingremoved
'--------------------------------------------------------------------------------
Option Explicit

Private msIdleTimeout As String
Private mbHasChanges As Boolean
Private mbIsValid As Boolean
Private mbIsLoading As Boolean

' limit timeout to 30 minutes max and 1 mintue min
Private Const mnMAX_TIMEOUT = "300"
Private Const mnMIN_TIMEOUT = "1"

'--------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------------------
    Unload Me
    
End Sub
'--------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'--------------------------------------------------------------------------------
    If mbHasChanges Then
        Call SaveProperties(txtIdleTimeout.Text)
        
    End If
    Unload Me
    
    Call gLog(gsSYS_TIMEOUT, " The system timeout is now set to " & txtIdleTimeout.Text & ".")
    

End Sub
'--------------------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------------------
' load system properties
'--------------------------------------------------------------------------------
Dim sIdleTimeout As String
    
       On Error GoTo ErrHandler
 
    Me.Icon = frmMenu.Icon
    
    mbIsLoading = True
    Call ReadProperties(sIdleTimeout)
    txtIdleTimeout.Text = sIdleTimeout
    lblTimeoutLimits.Caption = "Minimum of " & mnMIN_TIMEOUT & " and Maximum of "
    lblTimeoutLimits.Caption = lblTimeoutLimits.Caption & mnMAX_TIMEOUT & " minutes"
    msIdleTimeout = sIdleTimeout
    mbIsLoading = False
    
    FormCentre Me
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Form_Load")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub
'--------------------------------------------------------------------------------
Private Sub EnableOK()
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
    If mbIsValid Then
        cmdOK.Enabled = mbHasChanges
    
    Else
        cmdOK.Enabled = False
    
    End If
End Sub
'--------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'--------------------------------------------------------------------------------
Dim sMsg As String

    If mbHasChanges Then
        ' show save changes prompt
        sMsg = "Do you wish to save changes to system properties?"
        Select Case MsgBox(sMsg, vbQuestion + vbYesNoCancel, gsDIALOG_TITLE)
        Case vbYes
            ' save the data as requested then continue to unload
            Call SaveProperties(txtIdleTimeout.Text)
        Case vbNo
            ' just unload
            ' so do nothing here
        Case vbCancel
            ' cancel unload
            Cancel = True
        
        End Select
        
    End If
    
End Sub
'--------------------------------------------------------------------------------
Private Sub txtIdleTimeout_Change()
'--------------------------------------------------------------------------------
' allow only entry of numeric characters and
' limit timeout to 15 minutes (30 * 60 seconds)
'--------------------------------------------------------------------------------
Dim lPos As Long
    
    On Error GoTo InputErr
    If Not mbIsLoading Then
        mbIsValid = True
        
        ' PN 04/10/99 disallow null from being saved by flagging it as invalid
        If txtIdleTimeout.Text = vbNullString Then
            mbIsValid = False
            
        Else
            If Not gblnValidString(txtIdleTimeout.Text, valNumeric) Then
                ' the text entered was non-numeric
                ' raise the invalid property error
                Err.Raise 380
            
            ElseIf Val(txtIdleTimeout.Text) > mnMAX_TIMEOUT Then
                ' the timeout can not be more than mnMAX_TIMEOUT seconds
                ' raise the invalid property error
                Err.Raise 380
                
            ElseIf Val(txtIdleTimeout.Text) < mnMIN_TIMEOUT Then
                ' the timeout can not be less than mnMIN_TIMEOUT seconds
                ' allow the user to enter less than 60 but do not allow them to save
                ' by maintaining the mbIsValid flag
                mbIsValid = False
                
            Else
                If msIdleTimeout <> txtIdleTimeout.Text Then
                    mbHasChanges = True
                
                End If
                
            End If
        
        End If
        Call EnableOK
        msIdleTimeout = txtIdleTimeout.Text
        
    End If
    Exit Sub
    
InputErr:
    ' put the original text back into the control
    ' since the input failed validation
    Beep
    lPos = txtIdleTimeout.SelStart
    txtIdleTimeout = msIdleTimeout
    If lPos > 0 Then lPos = lPos - 1
    txtIdleTimeout.SelStart = lPos

End Sub
'--------------------------------------------------------------------------------
Private Sub ReadProperties(sIdleTimeout As String)
'--------------------------------------------------------------------------------
' read properties from macro db
'--------------------------------------------------------------------------------
Dim rsProperties As ADODB.Recordset
Dim sSQL As String
    On Error GoTo ErrHandler
    ' set-up the ado recordset, then open it
    ' no where clause because we are only interested in the first row
    Set rsProperties = New ADODB.Recordset
    sSQL = "Select IdleTimeout From MacroControl"
    rsProperties.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    ' read the data
    rsProperties.MoveFirst
    sIdleTimeout = rsProperties!IdleTimeout
    
    ' clean up
    rsProperties.Close
    Set rsProperties = Nothing
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ReadProperties")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub
'--------------------------------------------------------------------------------
Private Sub SaveProperties(sIdleTimeout As String)
'--------------------------------------------------------------------------------
' read properties from macro db
'--------------------------------------------------------------------------------
Dim rsProperties As ADODB.Recordset
Dim sSQL As String
    
    On Error GoTo ErrHandler
    ' set-up the ado recordset, then open it
    ' no where clause because we are only interested in the first row
    Set rsProperties = New ADODB.Recordset
    sSQL = "Select IdleTimeout From MacroControl"
    rsProperties.Open sSQL, MacroADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText

    ' save the data
    With rsProperties
        rsProperties!IdleTimeout = sIdleTimeout
        .Update
        
        ' clean up
        rsProperties.Close
        Set rsProperties = Nothing
    End With
    mbHasChanges = False
  
  Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SaveProperties")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
  
End Sub
