VERSION 5.00
Begin VB.Form frmNewUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New User"
   ClientHeight    =   2505
   ClientLeft      =   5265
   ClientTop       =   4215
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtConfirm 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtUserCode 
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Confirm password:"
      Height          =   225
      Left            =   60
      TabIndex        =   9
      Top             =   1600
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   225
      Left            =   600
      TabIndex        =   8
      Top             =   1125
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Full real name:"
      Height          =   225
      Left            =   300
      TabIndex        =   7
      Top             =   645
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "User name:"
      Height          =   225
      Left            =   540
      TabIndex        =   1
      Top             =   160
      Width           =   855
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       frmNewDataBase.frm
'   Author:     Will Casey, September 1999
'   Purpose:    Add a new user to the the database
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   WillC 11/10/99  Added the gLog call to log the creation of new users.
'   TA 08/05/2000   removed subclassing
'   WillC 7/8/00 SR2956   Changed labels
'   NCJ 26/10/00 - Tidied up layout; corrected validation of text fields
'   ASH 26/06/2002 - Disallow trailing spaces as part of password (Routine cmd_Ok)
'------------------------------------------------------------------------------------'

Option Explicit

Dim suCodeCol As Collection
Dim mnMaxPwdLength As Integer
Dim mnMinPwdLength As Integer

'------------------------------------------------------------------------------------'
Private Sub GetPasswordLengths()
'------------------------------------------------------------------------------------'
' Read the max. and min. password lengths allowed from DB
' Sets up module variables mnMaxPwdLength and mnMinPwdLength
'------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsPasswords As ADODB.Recordset

    sSQL = "SELECT * FROM MacroPassword"
    Set rsPasswords = New ADODB.Recordset
    rsPasswords.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
    
    mnMinPwdLength = rsPasswords!MinLength
    mnMaxPwdLength = rsPasswords!MaxLength
    
    rsPasswords.Close
    Set rsPasswords = Nothing

End Sub

'------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'------------------------------------------------------------------------------------'

   Unload Me
         
End Sub

'------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'------------------------------------------------------------------------------------'
'  Check to see if the  user exists already in the database using gblnUserExists
'  function in the Security module
'------------------------------------------------------------------------------------'
Dim sUserCode As String
Dim sMsg As String

    On Error GoTo ErrHandler
   
    sUserCode = Trim(txtUserCode.Text)
    ' Check to see if the  user exists already using gblnUserExists in the
    ' Security module
    If sUserCode <> "" Then
        If gblnUserExists(sUserCode) Then
            sMsg = "Sorry, a user with the name '" & sUserCode & "' already exists. Each user name must be unique."
            Call DialogError(sMsg)
            txtUserCode.Text = ""
            txtUserCode.SetFocus
            Exit Sub
        'ASH 26/06/2002 Do not allow passwords with leading or trailing spaces
        'since passwords are trimmed before stored in database.
        ElseIf txtPassword.Text <> Trim(txtPassword.Text) Then
            sMsg = "Sorry, passwords cannot contain leading or trailing spaces."
            Call DialogError(sMsg)
            txtPassword.Text = ""
            txtConfirm.Text = ""
            txtPassword.SetFocus
            Exit Sub
        Else
            Call NewPassword
        End If
    End If
    
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

'------------------------------------------------------------------------------------'
Private Sub NewPassword()
'------------------------------------------------------------------------------------'
' compare the two passwords to make sure they're the same. If so allow the
' insert then log the creation of a new user..
'------------------------------------------------------------------------------------'
Dim sPassword1 As String
Dim sPassword2 As String
Dim sUserCode As String
Dim sUserName As String
Dim sFirstLogin As String
Dim sMsg As String

    On Error GoTo ErrHandler

    sUserName = Trim(txtUserName.Text)
    sUserCode = Trim(txtUserCode.Text)
    sPassword1 = Trim(txtPassword.Text)
    sPassword2 = Trim(txtConfirm.Text)
    sFirstLogin = SQLStandardNow

   If sPassword1 = sPassword2 Then
        ' WillC 4/2/00 changed sFirstLogin to SQLStandardNow to cope with local settings
        Call InsertNewUser(sUserCode, sUserName, sFirstLogin, sPassword1)
        sMsg = "The new user '" & sUserName & "' has been successfully created"
        Call DialogInformation(sMsg)
        txtUserCode.Text = ""
        txtUserName.Text = ""
        txtPassword.Text = ""
        txtConfirm.Text = ""
        ' Log the creation of a new  user.
        Call gLog("CreateNewUser", "A new user '" & sUserName & "' was created.")
   ElseIf sPassword1 <> sPassword2 Then
        sMsg = "The passwords you have entered are not the same." & vbCrLf _
               & "Please enter the password again."
        Call DialogError(sMsg)
        txtConfirm.Text = ""
        txtPassword.Text = ""
        txtPassword.SetFocus
   End If

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "NewPassword")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'------------------------------------------------------------------------------------'
Private Sub Form_Load()
'------------------------------------------------------------------------------------'
' Open the database using the function in the Security module
'------------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler
 
    Me.Icon = frmMenu.Icon
    
    ' NCJ 26/10/00 - Read password lengths
    Call GetPasswordLengths
    
    'Clear the text boxes of their values
    txtUserCode.Text = ""
    txtUserName.Text = ""
    txtPassword.Text = ""
    txtConfirm.Text = ""
    txtUserCode.Tag = ""
    txtUserName.Tag = ""
    txtPassword.Tag = ""
    txtConfirm.Tag = ""
    
    cmdOK.Enabled = False
    
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

'------------------------------------------------------------------------------------'
Private Sub RefreshUsers()
'------------------------------------------------------------------------------------'
' Refresh the list of Users so that the one you have just added is visible to the
' User..
'------------------------------------------------------------------------------------'
Dim rsUsers As ADODB.Recordset

    On Error GoTo ErrHandler

    frmUserRoles.cboUsers.Clear
    Set suCodeCol = New Collection
    Set rsUsers = gdsUserList
    With rsUsers
             ' once the collection is instantiated add the members to the
             ' collection .ADD fields(1) are the names from the recordset and
             ' fields(0) is the usercode which is used as a key to write the data
             ' to the relevant textbox .
         Do Until .EOF = True
         frmUserRoles.cboUsers.AddItem .Fields(0).Value
             suCodeCol.Add .Fields(1).Value, .Fields(0).Value
                  .MoveNext
         Loop
    End With
    'changed Mo Morris 4/2/00
    rsUsers.Close
    Set rsUsers = Nothing
    
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

'------------------------------------------------------------------------------------'
Private Sub txtConfirm_Change()
'------------------------------------------------------------------------------------'
' Validate password
'------------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' NCJ 26/10/00
    Call TextBoxChange(txtConfirm)
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtConfirm_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'------------------------------------------------------------------------------------'
Private Sub txtPassword_Change()
'------------------------------------------------------------------------------------'
' Validate password
'------------------------------------------------------------------------------------'
 
    On Error GoTo ErrHandler
    
    ' NCJ 26/10/00
    Call TextBoxChange(txtPassword)
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtPassword_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'------------------------------------------------------------------------------------'
Private Sub EnableOKButton()
'------------------------------------------------------------------------------------'
' Enable/disable the OK button
'------------------------------------------------------------------------------------'

    If txtConfirm = "" Or _
        txtUserCode.Text = "" Or _
        txtUserName.Text = "" Or _
        txtPassword.Text = "" Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If

End Sub

'------------------------------------------------------------------------------------'
Private Sub txtPassword_LostFocus()
'------------------------------------------------------------------------------------'
' Check that the Password is within the boundaries set up for it.
'------------------------------------------------------------------------------------'
Dim nPwdLength As Integer
Dim sPrompt As String

    On Error GoTo ErrHandler
    
    nPwdLength = Len(txtPassword.Text)
    
    If nPwdLength > 0 Then
        If nPwdLength < mnMinPwdLength Then
            sPrompt = "The password is too short. The minimum number" & vbCr _
                        & " of characters for a password is " & mnMinPwdLength & "."
            Call DialogError(sPrompt)
            txtPassword.Text = ""
            txtPassword.SetFocus
        ElseIf nPwdLength > mnMaxPwdLength Then
            sPrompt = "The password may not be longer than " & mnMaxPwdLength & " characters."
            Call DialogError(sPrompt)
            txtPassword.Text = ""
            txtPassword.SetFocus
        End If
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtPassword_LostFocus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

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
    
    ' Enable the OK button as appropriate
    Call EnableOKButton


End Sub

'------------------------------------------------------------------------------------'
Private Sub txtUserCode_Change()
'------------------------------------------------------------------------------------'
' Validate user code
'------------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler
    
    ' NCJ 26/10/00
    Call TextBoxChange(txtUserCode)
    
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

'------------------------------------------------------------------------------------'
Private Sub txtUserName_Change()
'------------------------------------------------------------------------------------'
' Validate user name
'------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
   
    ' NCJ 26/10/00
    Call TextBoxChange(txtUserName)
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdApply_Click")
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
        sMsg = "A password or user name"
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
                                    "IsValidName")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function


