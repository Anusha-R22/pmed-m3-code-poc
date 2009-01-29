VERSION 5.00
Begin VB.Form frmPasswordPolicy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Policy Editor"
   ClientHeight    =   6600
   ClientLeft      =   3480
   ClientTop       =   3015
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4860
      TabIndex        =   13
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Password Restrictions"
      Height          =   1875
      Left            =   60
      TabIndex        =   17
      Top             =   4080
      Width           =   7400
      Begin VB.CheckBox chkLockout 
         Alignment       =   1  'Right Justify
         Caption         =   "Require account lockout"
         Height          =   375
         Left            =   200
         TabIndex        =   11
         Top             =   1200
         Width           =   2800
      End
      Begin VB.CheckBox chkExpiry 
         Alignment       =   1  'Right Justify
         Caption         =   "Require password expiry period"
         Height          =   375
         Left            =   180
         TabIndex        =   9
         Top             =   750
         Width           =   2800
      End
      Begin VB.CheckBox chkPassHistory 
         Alignment       =   1  'Right Justify
         Caption         =   "Check against previous passwords"
         Height          =   375
         Left            =   200
         TabIndex        =   7
         Top             =   300
         Width           =   2800
      End
      Begin VB.TextBox txtLockout 
         Height          =   315
         Left            =   6400
         MaxLength       =   3
         TabIndex        =   12
         Top             =   1200
         Width           =   800
      End
      Begin VB.TextBox txtExpiry 
         Height          =   315
         Left            =   6400
         MaxLength       =   3
         TabIndex        =   10
         Top             =   750
         Width           =   800
      End
      Begin VB.TextBox txtPassHistory 
         Height          =   315
         Left            =   6400
         MaxLength       =   3
         TabIndex        =   8
         Top             =   300
         Width           =   800
      End
      Begin VB.Label Label5 
         Caption         =   "Failed login attempts before account lockout"
         Height          =   375
         Left            =   3195
         TabIndex        =   22
         Top             =   1285
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Expiry period (days)"
         Height          =   375
         Left            =   3200
         TabIndex        =   21
         Top             =   820
         Width           =   2475
      End
      Begin VB.Label Label3 
         Caption         =   "Passwords remembered"
         Height          =   375
         Left            =   3200
         TabIndex        =   20
         Top             =   400
         Width           =   2400
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Repeating Characters and Username Usage (max 4)"
      Height          =   1250
      Left            =   60
      TabIndex        =   16
      Top             =   2760
      Width           =   7400
      Begin VB.CheckBox chkUserName 
         Alignment       =   1  'Right Justify
         Caption         =   "Allow portion of username"
         Height          =   375
         Left            =   200
         TabIndex        =   6
         Top             =   700
         Width           =   2800
      End
      Begin VB.CheckBox chkRepeat 
         Alignment       =   1  'Right Justify
         Caption         =   "Allow repeating of characters"
         Height          =   375
         Left            =   200
         TabIndex        =   5
         Top             =   300
         Width           =   2800
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Case and Numerical Settings"
      Height          =   1250
      Left            =   60
      TabIndex        =   15
      Top             =   1440
      Width           =   7400
      Begin VB.CheckBox chkDigit 
         Alignment       =   1  'Right Justify
         Caption         =   "At least one numerical digit"
         Height          =   375
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   2800
      End
      Begin VB.CheckBox chkMixedCase 
         Alignment       =   1  'Right Justify
         Caption         =   "Mixed case"
         Height          =   375
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Width           =   2800
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Password Length"
      Height          =   1300
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7400
      Begin VB.TextBox txtMaxLength 
         Height          =   315
         Left            =   2800
         MaxLength       =   3
         TabIndex        =   2
         Top             =   770
         Width           =   800
      End
      Begin VB.TextBox txtMinLength 
         Height          =   315
         Left            =   2800
         MaxLength       =   3
         TabIndex        =   1
         Top             =   300
         Width           =   800
      End
      Begin VB.Label Label2 
         Caption         =   "Maximum password length"
         Height          =   375
         Left            =   180
         TabIndex        =   19
         Top             =   780
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Minimum password length"
         Height          =   375
         Left            =   200
         TabIndex        =   18
         Top             =   340
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmPasswordPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2002 - 2003. All Rights Reserved
'   File:       frmPasswordPolicy.frm
'   Author:     Richard Meinesz, October 2002
'   Purpose:    Allow user to view and edit password policy
'--------------------------------------------------------------------
' Revisions:
' MLM 19/06/03 - Make yellow background work for Password history, expiry period and number of lockout attempts.
' REM 16/02/04 - In the Display routine we reload the password policy every time to ensure we have the correct values
'--------------------------------------------------------------------


Option Explicit

Private msMinLength As String
Private msMaxLength As String
Private msExpiry As String
Private mnMixCase As Integer
Private mnDigit As Integer
Private mnRepChars As Integer
Private mnUserName As Integer
Private msHistory As String
Private msRetries As String

Private mbIsLoading As Boolean
Private mbOKClicked As Boolean
Private mbIsChanged As Boolean

'--------------------------------------------------------------------
Public Function Display() As Boolean
'--------------------------------------------------------------------
'REM 18/09/02
'Display the password policy form
'ASH 13/1/20023  Added optional parameter to GetMacroDBSetting
'REM 16/02/04 - Reload the password policy every time to ensure we have the correct values
'--------------------------------------------------------------------
Dim sSecCon As String
Dim oSystemMessage As SysMessages
Dim sMessageParameters As String

    mbIsLoading = True
    
    mbOKClicked = False
    mbIsChanged = False
    
    cmdOK.Enabled = False
    
    'REM 16/02/04 - Reload the password policy every time to ensure we have the correct values
    Call goUser.PasswordPolicy.Load(SecurityADODBConnection)
    
    msMinLength = goUser.PasswordPolicy.MinPswdLength
    msMaxLength = goUser.PasswordPolicy.MaxPswdLength
    msExpiry = goUser.PasswordPolicy.ExpiryPeriod
    mnMixCase = -CInt(goUser.PasswordPolicy.EnforceMixedCase)
    mnDigit = -CInt(goUser.PasswordPolicy.EnforceDigit)
    mnRepChars = -CInt(goUser.PasswordPolicy.AllowRepeatChars)
    mnUserName = -CInt(goUser.PasswordPolicy.AllowUserName)
    msHistory = -CInt(goUser.PasswordPolicy.CheckPrevPswd)
    msRetries = goUser.PasswordPolicy.PasswordHistory

    txtMinLength.Text = msMinLength
    txtMaxLength.Text = msMaxLength
    
    chkMixedCase.Value = mnMixCase
    chkDigit.Value = mnDigit
    
    chkRepeat.Value = mnRepChars
    chkUserName.Value = mnUserName
    
    chkPassHistory.Value = msHistory
    txtPassHistory.Text = msRetries
    
    If goUser.PasswordPolicy.PasswordHistory = 0 Then
        txtPassHistory.Enabled = False
    End If
    
    chkExpiry.Value = -CInt(goUser.PasswordPolicy.RequirePswdExpiry)
    txtExpiry.Text = msExpiry
    
    If goUser.PasswordPolicy.ExpiryPeriod = 0 Then
        txtExpiry.Enabled = False
    End If
    
    chkLockout.Value = -CInt(goUser.PasswordPolicy.RequireAccountLockout)
    txtLockout.Text = goUser.PasswordPolicy.PasswordRetries
    
    If goUser.PasswordPolicy.PasswordRetries = 0 Then
        txtLockout.Enabled = False
    End If

    Me.Icon = frmMenu.Icon
    FormCentre Me
    
    mbIsLoading = False
    
    Me.Show vbModal
    
    If mbOKClicked Then
        If mbIsChanged Then
            'save password policy if changed
            Call goUser.PasswordPolicy.Save(SecurityADODBConnection)
            
            'send newe password policy to the message table (will only write a message it the database is a server DB
            Set oSystemMessage = New SysMessages
            sMessageParameters = msMinLength & gsPARAMSEPARATOR & msMaxLength & gsPARAMSEPARATOR & msExpiry & gsPARAMSEPARATOR & mnMixCase & gsPARAMSEPARATOR & mnDigit & gsPARAMSEPARATOR & mnRepChars & gsPARAMSEPARATOR & mnUserName & gsPARAMSEPARATOR & msHistory & gsPARAMSEPARATOR & msRetries
            Call oSystemMessage.AddNewSystemMessage(MacroADODBConnection, ExchangeMessageType.PasswordPolicy, goUser.UserName, goUser.UserName, "Password Policy", sMessageParameters)
            Set oSystemMessage = Nothing
            
        End If
    End If
    
    Display = mbOKClicked

End Function

'--------------------------------------------------------------------
Private Sub EnableOK(bOKEnabled As Boolean)
'--------------------------------------------------------------------
'Enable OK button
'--------------------------------------------------------------------

    If GetMacroDBSetting("datatransfer", "dbtype", , gsSERVER) = gsSERVER Then
        cmdOK.Enabled = bOKEnabled
    Else 'if its a site then never enable the OK button
        cmdOK.Enabled = False
    End If
    
End Sub

'--------------------------------------------------------------------
Private Sub chkDigit_Click()
'--------------------------------------------------------------------
'REM 18/09/02
'
'--------------------------------------------------------------------
Dim sMinLength As String
Dim sMaxLength As String

    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then

        goUser.PasswordPolicy.EnforceDigit = (chkDigit.Value = 1)
                
        'need to validate min password length if user chooses numeric digit, as it has to be at least 2
        ' and if user also chooses mixed case then min length has to be 3
        sMinLength = Trim(txtMinLength.Text)
        sMaxLength = Trim(txtMaxLength.Text)
        
        Call ValidPasswordLength(sMinLength, sMaxLength)
        
        mnDigit = chkDigit.Value
        
        mbIsChanged = True
    End If
    
End Sub

'--------------------------------------------------------------------
Private Sub chkMixedCase_Click()
'--------------------------------------------------------------------
'REM 18/09/02
'
'--------------------------------------------------------------------
Dim sMinLength As String
Dim sMaxLength As String

    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then

        goUser.PasswordPolicy.EnforceMixedCase = chkMixedCase.Value
        
        'need to validate min password length if user chooses mixed case, as it has to be at least 2
        ' and if user also chooses numeric digit then min length has to be 3
        sMinLength = Trim(txtMinLength.Text)
        sMaxLength = Trim(txtMaxLength.Text)
        
        Call ValidPasswordLength(sMinLength, sMaxLength)
        
        mnMixCase = chkMixedCase.Value
        
        mbIsChanged = True
    End If
    
End Sub

'--------------------------------------------------------------------
Private Sub chkLockout_Click()
'--------------------------------------------------------------------
'REM 18/09/02
'
'--------------------------------------------------------------------

    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
        'set the object property
        goUser.PasswordPolicy.RequireAccountLockout = (chkLockout.Value <> 0)
        
        'if the box is being unchecked need to set the lockout value to 0
        If Not goUser.PasswordPolicy.RequireAccountLockout Then
            goUser.PasswordPolicy.PasswordRetries = 0
            txtLockout.Enabled = False
            txtLockout.BackColor = vbWindowBackground
            Call EnableOK(True) 'cmdOK.Enabled = True
        Else
            txtLockout.Enabled = True
            If txtLockout.Text = "0" Then
                Call Valid(False, txtLockout)
                Call EnableOK(False) 'cmdOK.Enabled = False
            End If
        End If
        mbIsChanged = True
    End If

End Sub

'--------------------------------------------------------------------
Private Sub txtLockout_Change()
'--------------------------------------------------------------------
'REM 18/09/02
'
'--------------------------------------------------------------------
Dim sLockoutAttempts As String

    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
        
        sLockoutAttempts = Trim(txtLockout.Text)
        
        If Not IsNumeric(sLockoutAttempts) Then
            txtLockout.Text = txtLockout.Tag
            txtLockout.SelStart = 0
            txtLockout.SelLength = Len(txtLockout.Text)
        Else
            goUser.PasswordPolicy.PasswordRetries = CInt(sLockoutAttempts)
            msRetries = sLockoutAttempts
            If CLng(txtLockout.Text) = 0 Then
                Call Valid(False, txtLockout)
                Call EnableOK(False) 'cmdOK.Enabled = False
            Else
                Call Valid(True, txtLockout)
                Call EnableOK(False) '.Enabled = True
            End If
        End If
        
        mbIsChanged = True
    End If

    txtLockout.Tag = Trim(txtLockout.Text)
    
End Sub

'--------------------------------------------------------------------
Private Sub chkExpiry_Click()
'--------------------------------------------------------------------
'REM 18/09/02
'
'--------------------------------------------------------------------
    
    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
        'set the object property
        goUser.PasswordPolicy.RequirePswdExpiry = (chkExpiry.Value <> 0)
        
        'if the box is being unchecked need to set the expiry period value to 0
        If Not goUser.PasswordPolicy.RequirePswdExpiry Then
            goUser.PasswordPolicy.ExpiryPeriod = 0
            txtExpiry.Enabled = False
            txtExpiry.BackColor = vbWindowBackground
            Call EnableOK(True) 'cmdOK.Enabled = True
        Else
            txtExpiry.Enabled = True
            If txtExpiry.Text = "0" Then
                Call Valid(False, txtExpiry)
                Call EnableOK(False) 'cmdOK.Enabled = False
            End If
        End If
        mbIsChanged = True
    End If
    
End Sub

'--------------------------------------------------------------------
Private Sub txtExpiry_Change()
'--------------------------------------------------------------------
'REM 18/09/02
'
'--------------------------------------------------------------------
Dim sExpiryPeriod As String

    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
        
        sExpiryPeriod = Trim(txtExpiry.Text)
        
        If Not IsNumeric(sExpiryPeriod) Then
            txtExpiry.Text = txtExpiry.Tag
            txtExpiry.SelStart = 0
            txtExpiry.SelLength = Len(txtExpiry.Text)
        Else
            goUser.PasswordPolicy.ExpiryPeriod = CInt(sExpiryPeriod)
            msExpiry = sExpiryPeriod
            If CLng(txtExpiry.Text) = 0 Then
                Call Valid(False, txtExpiry)
                Call EnableOK(False) 'cmdOK.Enabled = False
            Else
                Call Valid(True, txtExpiry)
                Call EnableOK(True) ' cmdOK.Enabled = True
            End If

        End If
        
        mbIsChanged = True
    End If
    
    txtExpiry.Tag = Trim(txtExpiry.Text)
    
End Sub

'--------------------------------------------------------------------
Private Sub chkPassHistory_Click()
'--------------------------------------------------------------------
'REM 18/09/02
'
'--------------------------------------------------------------------
    
    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
        'set the object property
        goUser.PasswordPolicy.CheckPrevPswd = (chkPassHistory.Value <> 0)
        
        'if the box is being unchecked need to set the password history value to 0
        If Not goUser.PasswordPolicy.CheckPrevPswd Then
            goUser.PasswordPolicy.PasswordHistory = 0
            txtPassHistory.Enabled = False
            txtPassHistory.BackColor = vbWindowBackground
            Call EnableOK(True) ' cmdOK.Enabled = True
        Else
            txtPassHistory.Enabled = True
            If txtPassHistory.Text = "0" Then
                Call Valid(False, txtPassHistory)
                Call EnableOK(False) ' cmdOK.Enabled = False
            End If
        End If
        mbIsChanged = True
    End If


End Sub

'--------------------------------------------------------------------
Private Sub txtPassHistory_Change()
'--------------------------------------------------------------------
'REM 18/09/02
'
'--------------------------------------------------------------------
Dim sPasswordHistory As String
Dim bValid As Boolean

    
    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
        
        sPasswordHistory = Trim(txtPassHistory.Text)
        
        If Not IsNumeric(sPasswordHistory) Then
            txtPassHistory.Text = txtPassHistory.Tag
            txtPassHistory.SelStart = 0
            txtPassHistory.SelLength = Len(txtPassHistory.Text)
        Else
            If CLng(txtPassHistory.Text) = 0 Then
                Call Valid(False, txtPassHistory)
                Call EnableOK(False) ' cmdOK.Enabled = False
            Else
                Call Valid(True, txtPassHistory)
                Call EnableOK(True) ' cmdOK.Enabled = True
                goUser.PasswordPolicy.PasswordHistory = CInt(sPasswordHistory)
                msHistory = sPasswordHistory
            End If
        End If
        
        mbIsChanged = True
    End If
    
    txtPassHistory.Tag = Trim(txtPassHistory.Text)

End Sub

'--------------------------------------------------------------------
Private Sub chkRepeat_Click()
'--------------------------------------------------------------------
'REM 18/09/02
'
'--------------------------------------------------------------------

    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
    
        goUser.PasswordPolicy.AllowRepeatChars = (chkRepeat.Value <> 0)
        
        mnRepChars = chkRepeat.Value
        
        mbIsChanged = True
    End If
    
End Sub

'--------------------------------------------------------------------
Private Sub chkUserName_Click()
'--------------------------------------------------------------------
'REM 18/09/02
'
'--------------------------------------------------------------------

    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
    
        goUser.PasswordPolicy.AllowUserName = (chkUserName.Value <> 0)
        
        mnUserName = chkUserName.Value
        
        mbIsChanged = True
    End If
    
End Sub

'--------------------------------------------------------------------
Private Sub txtMaxLength_Change()
'--------------------------------------------------------------------
'REM 18/09/02
'
'--------------------------------------------------------------------

    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
    
        msMaxLength = Trim(txtMaxLength.Text)
        
        If Not IsNumeric(msMaxLength) Then
            txtMaxLength.Text = txtMinLength.Tag
            txtMaxLength.SelStart = 0
            txtMaxLength.SelLength = Len(txtMaxLength.Text)
        Else
        
            Call ValidPasswordLength(msMinLength, msMaxLength)
            goUser.PasswordPolicy.MaxPswdLength = msMaxLength
            mbIsChanged = True
        End If
    End If
    
    txtMaxLength.Tag = Trim(txtMaxLength.Text)
    
End Sub

'--------------------------------------------------------------------
Private Sub txtMinLength_Change()
'--------------------------------------------------------------------
'REM 18/09/02
'
'--------------------------------------------------------------------

    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
    
        msMinLength = Trim(txtMinLength.Text)
        
        If Not IsNumeric(msMinLength) Then
            txtMinLength.Text = txtMinLength.Tag
            txtMinLength.SelStart = 0
            txtMinLength.SelLength = Len(txtMinLength.Text)
        Else
            Call ValidPasswordLength(msMinLength, msMaxLength)
            goUser.PasswordPolicy.MinPswdLength = msMinLength
            mbIsChanged = True
        End If
    End If

    txtMinLength.Tag = Trim(txtMinLength.Text)
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub Valid(bIsValid As Boolean, txtTextBox As TextBox)
'--------------------------------------------------------------------------------------------------
' REM 18/09/02
' Changes the text box background colour if invalid data is entered
'--------------------------------------------------------------------------------------------------

    'If bValid is true then text box background white else yellow
    If bIsValid Then
        txtTextBox.BackColor = vbWindowBackground
    Else
        txtTextBox.BackColor = g_INVALID_BACKCOLOUR
    End If

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------------------------------------

    mbOKClicked = False
    Unload Me

End Sub

'--------------------------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'--------------------------------------------------------------------------------------------------

    mbOKClicked = True
    Unload Me
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'--------------------------------------------------------------------------------------------------
    'REM 16/02/04 - Only show message if its a server machine as you can't save changes on a Site
    If GetMacroDBSetting("datatransfer", "dbtype", , gsSERVER) = gsSERVER Then
        If (mbOKClicked = False) And (mbIsChanged = True) Then
            If DialogQuestion("Are you sure you want to cancel without saving?") = vbNo Then
                Cancel = 1
            End If
        End If
    End If
End Sub

'--------------------------------------------------------------------------------------------------
Private Function ValidInteger(sText As String) As Boolean
'--------------------------------------------------------------------------------------------------
' NCJ 28 Feb 02
' Validates sText - must be an integer and greater than 0
' Assume already trimmed
' Returns FALSE if not a valid integer
'--------------------------------------------------------------------------------------------------
        
    On Error GoTo NotAnInteger
    
    ValidInteger = False
    
    If sText = "" Then Exit Function
    
    If Not IsNumeric(sText) Then Exit Function
    
    If Val(sText) <> CInt(sText) Then Exit Function
    
    ValidInteger = True
        
NotAnInteger:

End Function

'--------------------------------------------------------------------------------------------------
Private Sub txtMaxLength_GotFocus()
'--------------------------------------------------------------------------------------------------

    txtMaxLength.SelStart = 0
    txtMaxLength.SelLength = Len(txtMaxLength.Text)
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub txtMinLength_GotFocus()
'--------------------------------------------------------------------------------------------------

    txtMinLength.SelStart = 0
    txtMinLength.SelLength = Len(txtMinLength.Text)
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub txtPassHistory_GotFocus()
'--------------------------------------------------------------------------------------------------

    txtPassHistory.SelStart = 0
    txtPassHistory.SelLength = Len(txtPassHistory.Text)
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub txtExpiry_GotFocus()
'--------------------------------------------------------------------------------------------------

    txtExpiry.SelStart = 0
    txtExpiry.SelLength = Len(txtExpiry.Text)
    
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub txtLockout_GotFocus()
'--------------------------------------------------------------------------------------------------

    txtLockout.SelStart = 0
    txtLockout.SelLength = Len(txtLockout.Text)
    
End Sub

'--------------------------------------------------------------------------------------------------
Public Sub ValidPasswordLength(sMinLength As String, sMaxLength As String)
'--------------------------------------------------------------------------------------------------
'REM 18/09/02
'Validate the password length fields
'--------------------------------------------------------------------------------------------------
Dim bValidMin As Boolean
Dim bValidMax As Boolean

        If sMinLength = "" Then
            sMinLength = "0"
        End If
        
        If sMaxLength = "" Then
            sMaxLength = "0"
        End If
        
        bValidMax = True
        
        'if not valid interger
        If ValidInteger(sMaxLength) = False Then
            bValidMax = False
        End If
        
        'if max length is less then min length
        If CInt(sMaxLength) < CInt(sMinLength) Then
            bValidMax = False
        End If
        
        'if max length is 0
        If CInt(sMaxLength) < 1 Then
            bValidMax = False
        End If
        
        ' else valid assign the value to the object
        If bValidMax Then
            goUser.PasswordPolicy.MaxPswdLength = CInt(sMaxLength)
        End If
        
        'change background colour of max length field according to valid status
        Call Valid(bValidMax, txtMaxLength)
        
'************************************
        
        bValidMin = True
        
        'if not valid integer
        If ValidInteger(sMinLength) = False Then
            bValidMin = False
        End If
        
        'if min length is greater than max length
        If CInt(sMinLength) > CInt(sMaxLength) Then
            bValidMin = False
        End If
        
        'if min length is 0
        If CInt(sMinLength) < 1 Then
            bValidMin = False
        End If
        
        'if the mixed case or numeric digit boxes are checked then min massword length must be 2
        If (chkMixedCase.Value = 1) Or (chkDigit.Value = 1) Then
            If CInt(sMinLength) < 2 Then
                bValidMin = False
            End If
        End If
        
        'if the mixed case and numeric digit boxes are checked then min massword length must be 3
        If (chkMixedCase.Value = 1) And (chkDigit.Value = 1) Then
            If CInt(sMinLength) < 3 Then
                bValidMin = False
            End If
        End If
        
        ' else valid so assign the value to the object
        If bValidMin = True Then
            goUser.PasswordPolicy.MinPswdLength = sMinLength
        End If
        
        'change background colour of max length field according to valid status
        Call Valid(bValidMin, txtMinLength)
        
        'set the cmdOK button enabled to false unless both the above fields are valid
        If bValidMax And bValidMin Then
            Call EnableOK(True) ' cmdOK.Enabled = True
        Else
            Call EnableOK(False) ' cmdOK.Enabled = False
        End If
        
End Sub

