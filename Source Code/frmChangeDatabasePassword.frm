VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChangeDatabasePassword 
   Caption         =   "Change Database Password"
   ClientHeight    =   2295
   ClientLeft      =   10035
   ClientTop       =   6690
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   3735
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   3615
      Begin VB.TextBox txtOldPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtNewPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtConfirm 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Old password:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&New password:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Con&firm new password:"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2460
      TabIndex        =   4
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1140
      TabIndex        =   3
      Top             =   1860
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   0
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmChangeDatabasePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       frmChangeDatabasePassword.frm
'   Author:     Andrew Newbigging, 1999
'   Purpose:    To allow the user to change the password protecting an Access database
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'Revisions:
' REM 23/04/03 - reorganised form to cope with new MACRO 3.0 functionality
'----------------------------------------------------------------------------------------'

Option Explicit

Private msDatabaseCode As String
Private msNewPassword As String
Private msConfirmPassword As String
Private msOldPassword As String

'---------------------------------------------------------------------
Public Sub Display(sDatabaseCode As String)
'---------------------------------------------------------------------
'REM 23/04/03
'Display routine for form
' Pass in database code
'---------------------------------------------------------------------

    Me.Icon = frmMenu.Icon
    
    msDatabaseCode = sDatabaseCode
    
    Me.BackColor = glFormColour
    
    Me.Caption = "Change " & sDatabaseCode & " Database Password"

    Screen.MousePointer = vbDefault
    
    EnableOKButton

    FormCentre Me
    
    Me.Show vbModal


End Sub

'---------------------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------------------
' Make sure the OK button is enabled appropriately
'---------------------------------------------------------------------

    EnableOKButton
    
End Sub


'----------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'----------------------------------------------------------------------------------------'
' Unloads the form
' NCJ 28/1/00 - Only hide the form, because we want to get at its values later
'----------------------------------------------------------------------------------------'
    
    Unload Me
   
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'----------------------------------------------------------------------------------------'

   Call ChangeDBPassword
   EnableOKButton
   
End Sub

'----------------------------------------------------------------------------------------'
Private Sub ChangeDBPassword()
'----------------------------------------------------------------------------------------'
' Check that the new password and its confirmation match
' If so, set our local variables and hide the form
' Otherwise show message and set focus to new password field
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler

    If msNewPassword <> msConfirmPassword Then 'check that the password and confirm password match
        Call DialogError("The new and confirmed passwords do not match. Please type them again.")
        txtNewPassword.Text = ""
        txtConfirm.Text = ""
        txtNewPassword.SetFocus
    ElseIf Not CheckOldDBPassword Then 'check old password is correct
        Call DialogError("The old database password is incorrect.")
        txtOldPassword = ""
        txtOldPassword.SetFocus
    Else 'change database password
        Call UpDateDBPassword
        Call DialogInformation(msDatabaseCode & " database password has been changed.")
        Unload Me
    End If

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ChangePassword")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub

'----------------------------------------------------------------------------------------'
Private Function CheckOldDBPassword() As Boolean
'----------------------------------------------------------------------------------------'
'REM 23/4/03
'Check to see that the old password entered matches the one currently in the secuirty database
'----------------------------------------------------------------------------------------'
Dim sPswdOld As String
Dim sEncryptedOldPswd As String
Dim sSQL As String
Dim rsDatabase As ADODB.Recordset

    sSQL = "SELECT DatabasePassword FROM Databases" _
        & " WHERE DatabaseCode = '" & msDatabaseCode & "'"
    Set rsDatabase = New ADODB.Recordset
    rsDatabase.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    sEncryptedOldPswd = rsDatabase.Fields(0).Value
    
    sPswdOld = DecryptString(sEncryptedOldPswd)
    
    If LCase(sPswdOld) = LCase(msOldPassword) Then
        CheckOldDBPassword = True
    Else
        CheckOldDBPassword = False
    End If


End Function

'----------------------------------------------------------------------------------------'
Private Sub UpDateDBPassword()
'----------------------------------------------------------------------------------------'
'REM 23/04/03
'Update a database password in the security database
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim sEncryptDBPswd As String

    sEncryptDBPswd = EncryptString(msNewPassword)
    
    sSQL = "UPDATE Databases " _
    & " SET DatabasePassword = '" & sEncryptDBPswd & "'" _
    & " WHERE DatabaseCode = '" & msDatabaseCode & "'"
    SecurityADODBConnection.Execute sSQL, adCmdText
    
End Sub

''----------------------------------------------------------------------------------------'
'Public Property Get OldPassword() As Variant
''----------------------------------------------------------------------------------------'
'
'    OldPassword = msOldPassword
'
'End Property
'
''----------------------------------------------------------------------------------------'
'Public Property Let OldPassword(ByVal vNewValue As Variant)
''----------------------------------------------------------------------------------------'
'
'    txtOldPassword.Text = Trim(vNewValue)
'
'End Property
'
''----------------------------------------------------------------------------------------'
'Public Property Get NewPassword() As Variant
''----------------------------------------------------------------------------------------'
'
'    NewPassword = msNewPassword
'
'End Property
'
''----------------------------------------------------------------------------------------'
'Public Property Let NewPassword(ByVal vNewValue As Variant)
''----------------------------------------------------------------------------------------'
'
'    txtNewPassword.Text = Trim(vNewValue)
'    txtConfirm.Text = Trim(vNewValue)
'
'End Property

'----------------------------------------------------------------------------------------'
Private Sub EnableOKButton()
'----------------------------------------------------------------------------------------'
' NCJ 28/1/00
' At least one of the OldPassword and NewPassword fields must be non-empty
' (i.e. don't allow change from empty to empty!)
'----------------------------------------------------------------------------------------'

    If (msNewPassword > "") Or (msOldPassword > "") Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If

End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtNewPassword_Change()
'----------------------------------------------------------------------------------------'
' Set our local variable and enable OK button
'----------------------------------------------------------------------------------------'

    msNewPassword = Trim(txtNewPassword.Text)
    EnableOKButton
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtOldPassword_Change()
'----------------------------------------------------------------------------------------'
' Set our local variable and enable OK button
'----------------------------------------------------------------------------------------'

    msOldPassword = Trim(txtOldPassword.Text)
    EnableOKButton

End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtConfirm_Change()
'----------------------------------------------------------------------------------------'
' Set our local variable and enable OK button
'----------------------------------------------------------------------------------------'

    msConfirmPassword = Trim(txtConfirm.Text)
    EnableOKButton

End Sub
