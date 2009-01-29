VERSION 5.00
Begin VB.Form frmEditUserName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit User Name"
   ClientHeight    =   1575
   ClientLeft      =   5355
   ClientTop       =   5085
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   4395
      Begin VB.TextBox txtUserCode 
         Height          =   345
         Left            =   1260
         TabIndex        =   0
         Top             =   180
         Width           =   2895
      End
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   1260
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "User name"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Full real name"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1140
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1140
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditUserName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       frmEditUserName.frm
'   Author:     Will Casey, September 1999
'   Purpose:    To allow the user to edit the user name not user code
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'Revisions:
'   WillC 7/8/00 SR2956   Changed labels
'   ZA 10/06/2002 - Copied IsValidString function and TextBoxChanged functions from frmNewUser
'                   to fix bug 13 in build 2.2.14
'   ZA 18/09/02 updates List_users.js when a user name is edited
'----------------------------------------------------------------------------------------'
Option Explicit
Option Base 0


'----------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'----------------------------------------------------------------------------------------'
' unload the form
'----------------------------------------------------------------------------------------'
 
    Unload Me
   
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'----------------------------------------------------------------------------------------'
' Edit the users name and refresh the database
'----------------------------------------------------------------------------------------'
On Error GoTo ErrHandler

Dim sUserCode As String
Dim sUsername As String

    sUserCode = txtUserCode.Text
    sUsername = txtUserName.Text

    Call EditUserName(sUserCode, sUsername)
       
    RefreshUsers
    
    Unload Me
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, " cmdOK_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Form_Load()
'----------------------------------------------------------------------------------------'
'Disable the Ok button and usercode textbox
'----------------------------------------------------------------------------------------'
On Error GoTo ErrHandler

Dim sUserCode As String
Dim sUsername As String

    Me.Icon = frmMenu.Icon
    
    txtUserCode.Enabled = False
    cmdOK.Enabled = False
    
    sUserCode = frmUsers.cboUsers.Text
    sUsername = frmUsers.txtUserName.Text
    
    txtUserCode.Text = sUserCode
    txtUserName.Text = sUsername
    
    FormCentre Me
    
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

'----------------------------------------------------------------------------------------'
Private Sub txtUserCode_Change()
'----------------------------------------------------------------------------------------'
' only enable the Ok button so long as there is something in the box
'----------------------------------------------------------------------------------------'
On Error GoTo ErrHandler
    If txtUserName.Text = "" Or txtUserCode = "" Then
         cmdOK.Enabled = False
    Else
         cmdOK.Enabled = True
         
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtUserCode_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtUserName_Change()
'----------------------------------------------------------------------------------------'
' only enable the Ok button so long as there is something in the box
'----------------------------------------------------------------------------------------'
On Error GoTo ErrHandler
    
    
    If txtUserName.Text = "" Or txtUserCode = "" Then
       cmdOK.Enabled = False
    Else
        TextBoxChange txtUserName
        cmdOK.Enabled = True
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, " txtUserName_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EditUserName(sUserCode As String, sUsername As String)
'----------------------------------------------------------------------------------------'
'do the update
'----------------------------------------------------------------------------------------'
On Error GoTo ErrHandler

Dim sSQL As String

    'Mo Morris 20/9/01 Db Audit (UserCode to UserName, UserName to UserNameFull)
    sSQL = " Update MACROUser SET UserNameFull = '" & sUsername & "'" _
        & " WHERE UserName ='" & sUserCode & "'"
        
    SecurityADODBConnection.Execute sSQL
    
    'ZA 18/09/2002 - update List_users.js
    CreateUsersList
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "EditUserName")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'----------------------------------------------------------------------------------------------
Private Sub RefreshUsers()
'----------------------------------------------------------------------------------------------
' Add the list of Users to the User combobox
'----------------------------------------------------------------------------------------------
On Error GoTo ErrHandler

   frmUsers.cboUsers.Text = txtUserCode.Text
   frmUsers.txtUserName.Text = txtUserName.Text
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshUsers")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'----------------------------------------------------------------------------------------'
Private Function IsValidString(sDescription As String) As Boolean
'----------------------------------------------------------------------------------------'
' Return TRUE if text is valid
' Displays any necessary messages
'----------------------------------------------------------------------------------------'
Dim sMsg As String

    On Error GoTo ErrHandler
    
    IsValidString = False
    
    If sDescription > "" Then
        sMsg = "A user name"
        If Not gblnValidString(sDescription, valOnlySingleQuotes) Then
            sMsg = sMsg & gsCANNOT_CONTAIN_INVALID_CHARS
            Call DialogError(sMsg)
        ElseIf Not gblnValidString(sDescription, valAlpha + valNumeric + valSpace) Then
            sMsg = sMsg & " may only contain alphanumeric characters"
            Call DialogError(sMsg)
        ElseIf Len(sDescription) > 255 Then
            sMsg = sMsg & " may not be more than 255 characters"
            Call DialogError(sMsg)
        Else
            IsValidString = True
        End If
    End If
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "IsValidString")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function


'------------------------------------------------------------------------------------'
Private Sub TextBoxChange(txtTextBox As TextBox)
'------------------------------------------------------------------------------------'
' NCJ 26/10/00
' Validate a text field on this form
'------------------------------------------------------------------------------------'
Dim sText As String
    
    sText = Trim(txtTextBox.Text)
    
    If sText > "" Then
        If IsValidString(sText) = False Then
            ' Replace field with previous contents
            txtTextBox.Text = txtTextBox.Tag
            ' Put cursor at the end
            txtTextBox.SelStart = Len(txtTextBox.Text)
        Else
            ' Store contents for next time
            txtTextBox.Tag = sText
        End If
    Else
        ' Screen out superfluous spaces
        txtTextBox.Text = ""
        txtTextBox.Tag = ""
    End If
   

End Sub

