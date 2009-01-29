VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2325
   ClientLeft      =   5895
   ClientTop       =   4605
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   3795
   Begin VB.Frame fraPassword 
      Height          =   1695
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   3675
      Begin VB.TextBox txtOldPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1860
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtNewPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1860
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtConfirm 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1860
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Old Password"
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&New Password"
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Con&firm New Password"
         Height          =   375
         Left            =   60
         TabIndex        =   6
         Top             =   1260
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1860
      Width           =   1215
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       frmChangePassword.frm
'   Author:     Will Casey, September 1999
'   Purpose:    To allow the user to create a new password for themselves when
'               their old one has run out.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'Revisions:
'WillC      11/10/99    Added gLog call to the Ok button to log the password being changed
'WillC      29/11/99    If the users password has expired disable the cancel button to
'                       stop the user doing anything until they change their password.
''  WillC 11 / 12 / 99
'          Changed the Following where present from Integer to Long  ClinicalTrialId
'           CRFPageId,VisitId,CRFElementID
'   TA 25/04/2000   subclassing removed
'   TA 15/04/2000   display function now displays the form and returns whether the password was succesfully changed
'   TA 19/05/2000   changed borders style to fixed dialog
'   NCJ 31/5/2000   Minor changes for coding standards compliance (for SR3376)
'                   Use gsCANNOT_CONTAIN_INVALID_CHARS
'   TA 24/10/2000:  There is now a system log when changing a password
'   Mo 18/7/2001    Changes stemming from field Password in table MacroUser being changed to
'                   UserPassword (stems from the swith to Jet 4.0)
'   DPH 25/10/2001  Need to return new password to user class
'   ASH 26/06/2002 - Disallow leading or trailing spaces as part of password (Routine ChangePassword)
'----------------------------------------------------------------------------------------'

Option Explicit
Private mbSuccess As Boolean
Private msUser As String
Private msNewPassword As String

'----------------------------------------------------------------------------------------'
Public Function Display(sUser As String, sNewPass As String) As Boolean
'----------------------------------------------------------------------------------------'
' Display Change Password form
' Input:
'   sUser - usercode
' Output:
'   function - successful change?
'----------------------------------------------------------------------------------------'
    mbSuccess = False
    msUser = sUser
    ' DPH 25/10/2001 - New password
    msNewPassword = ""
    
    cmdOK.Enabled = False
    Me.Icon = frmMenu.Icon
    fraPassword.Caption = "User - " & sUser
    
    FormCentre Me
    Me.Show vbModal
    
    ' DPH 25/10/2001 - New password (to return)
    sNewPass = msNewPassword
    Display = mbSuccess
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'----------------------------------------------------------------------------------------'
' Unloads the form
'----------------------------------------------------------------------------------------'
    
    Unload Me
   
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'----------------------------------------------------------------------------------------'
' Check to see if the old password is correct then do the boundary checks if everything is
' ok then do the change
'----------------------------------------------------------------------------------------'

   Call ChangePassword
   
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Initialize()
'----------------------------------------------------------------------------------------'
' Initialization
'----------------------------------------------------------------------------------------'

    msNewPassword = ""
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Load()
'----------------------------------------------------------------------------------------'
'Disable the ok button
'----------------------------------------------------------------------------------------'
    On Error GoTo ErrHandler
              
    Me.BackColor = glFormColour
   
    Me.Icon = frmMenu.Icon
    HourglassSuspend    'must turn off in form_unload
    
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

Private Sub Form_Unload(Cancel As Integer)

    HourglassResume ' to match hourglasssuspend in form_load
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtConfirm_Change()
'----------------------------------------------------------------------------------------'
' Disable the ok button until all three textboxes have entries
'----------------------------------------------------------------------------------------'
Dim sDescription As String

    On Error GoTo ErrHandler

    sDescription = txtConfirm.Text
    
    If IsValidString(sDescription) = False Then
    'If the user Types an invalid char this gets rid of it
        txtConfirm.Text = Replace(txtConfirm.Text, Right(txtConfirm.Text, 1), "")
    End If

    If txtOldPassword.Text = "" Or txtNewPassword.Text = "" Or txtConfirm.Text = "" Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
    
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

'----------------------------------------------------------------------------------------'
Private Sub txtNewPassword_Change()
'----------------------------------------------------------------------------------------'
' Disable the ok button until all three textboxes have entries
'----------------------------------------------------------------------------------------'
Dim sDescription As String

    On Error GoTo ErrHandler
 
    sDescription = txtNewPassword.Text
    
    If IsValidString(sDescription) = False Then
    'If the user Types an invalid char this gets rid of it
        txtNewPassword.Text = Replace(txtNewPassword.Text, Right(txtNewPassword.Text, 1), "")
    End If


    If txtOldPassword.Text = "" Or txtNewPassword.Text = "" Or txtConfirm.Text = "" Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtNewPassword_Change")
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
Private Sub txtOldPassword_Change()
'----------------------------------------------------------------------------------------'
' Disable the ok button until all three textboxes have entries
'----------------------------------------------------------------------------------------'
Dim sDescription As String
 
    On Error GoTo ErrHandler
    
    sDescription = txtOldPassword.Text
    
    If IsValidString(sDescription) = False Then
    'If the user Types an invalid char this gets rid of it
        txtOldPassword.Text = Replace(txtOldPassword.Text, Right(txtOldPassword.Text, 1), "")
    End If
    
    If txtOldPassword.Text = "" Or txtNewPassword.Text = "" Or txtConfirm.Text = "" Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtOldPassword_Change")
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
Public Sub ChangePassword()
'----------------------------------------------------------------------------------------'
' Check to see if the old password is correct then do the boundary checks if everything is
' ok then do the change and log the successful change.
'----------------------------------------------------------------------------------------'
Dim rsUser As ADODB.Recordset
Dim rsPassword As ADODB.Recordset
Dim sSQL As String
Dim sPassword As String
Dim sUserCode As String
Dim sNewPassword As String
Dim sOldPassword As String
Dim nMinLength As Integer
Dim nMaxLength As Integer
Dim sMsg As String

    On Error GoTo ErrHandler

    sUserCode = msUser
    sOldPassword = txtOldPassword.Text
    
    'Mo Morris 20/9/01 Db Audit (UserCode to UserName)
    sSQL = "SELECT * FROM MacroUser WHERE UserName = '" & sUserCode & "'"
    Set rsUser = New ADODB.Recordset
    rsUser.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, , adCmdText
    'Changed Mo 18/7/01
    'sPassword = rsUser!Password
    sPassword = rsUser!UserPassword
    Set rsUser = Nothing
    
    sSQL = "SELECT * FROM MacroPassWord "
    Set rsPassword = New ADODB.Recordset
    rsPassword.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, , adCmdText
    nMinLength = rsPassword!MinLength
    nMaxLength = rsPassword!MaxLength
    Set rsPassword = Nothing
   
   'ASH 26/06/2002 Do not allow passwords with leading or trailing spaces
   'since passwords are trimmed before stored in database.
   If txtNewPassword.Text <> Trim(txtNewPassword.Text) Then
        sMsg = "Sorry, passwords cannot contain leading or trailing spaces."
        Call DialogError(sMsg)
        txtNewPassword.Text = ""
        txtConfirm.Text = ""
        txtNewPassword.SetFocus
        Exit Sub
    End If
    
    If txtOldPassword.Text <> sPassword Then
        MsgBox "The Macro password that you have typed is not correct." & vbCr _
             & "Please enter it again.", vbCritical, "MACRO"
        txtOldPassword.Text = ""
        txtOldPassword.SetFocus
        Exit Sub
    End If
    
    If txtNewPassword.Text <> txtConfirm.Text Then
        MsgBox " The new and confirmed passwords do not match, please type them again.", vbCritical, "MACRO"
        txtNewPassword.Text = ""
        txtConfirm.Text = ""
        txtNewPassword.SetFocus
        Exit Sub
    End If
        
    If txtNewPassword.Text = txtConfirm.Text Then
        sNewPassword = txtNewPassword.Text
    End If
    
    If Len(sNewPassword) > nMaxLength Then
        MsgBox " The new password is too long, the maximum number" & vbCr _
             & " of characters for a password is " & nMaxLength & ".", vbInformation, "MACRO"
        txtNewPassword.Text = ""
        txtConfirm.Text = ""
        txtNewPassword.SetFocus
    ElseIf Len(sNewPassword) < nMinLength Then
        MsgBox " The new password is too short, the minimum number" & vbCr _
             & " of characters for a password is " & nMinLength & ".", vbInformation, "MACRO"
        txtNewPassword.Text = ""
        txtConfirm.Text = ""
        txtNewPassword.SetFocus
    End If
    
    If LCase(txtOldPassword.Text) = LCase(sNewPassword) Then
        MsgBox "You may not set your new password to be the " & vbCrLf _
            & "same as your old password.Please choose again.", vbInformation, "MACRO"
            txtNewPassword.Text = ""
            txtConfirm.Text = ""
            txtNewPassword.SetFocus
        Exit Sub
    End If

    If Len(sNewPassword) <= nMaxLength And Len(sNewPassword) >= nMinLength Then
        Call UpdateUserPassword(sUserCode, sNewPassword)
        mbSuccess = True
        MsgBox "Your password has been successfully changed.", vbInformation, "MACRO"
        
'        'WillC 21/10/99 Dont log the change if its the default user
'        If sUserCode = "rde" And sOldPassword = "macrotm" Then
'         Else
'            'WillC took out gLog call to log the change 29/11/99
'            'Call gLog("sChangePassword", " The password for user " & sUserCode & " was changed")
'        End If
      
        Unload Me
        
    Else
        txtOldPassword.Text = ""
        txtNewPassword.Text = ""
        txtConfirm.Text = ""
        txtOldPassword.SetFocus
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ChangePassword")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'--------------------------------------------------------------------------------
Private Sub UpdateUserPassword(sUserCode As String, sPassword As String)
'--------------------------------------------------------------------------------
' update the user password
'4/2/00 Will changed date to SQLStandardNow to cope with regional settings
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim sFirstLogin As String

    On Error GoTo ErrHandler
    
    'dtFirstLogin = Format(Now, "dd/mm/yyyy hh:mm:ss")
    sFirstLogin = SQLStandardNow
    ' Will 7/10/99 added firstlogin
    ' PN 15/09/99 -  change tablename User to MacroUser
    'changed Mo 18/7/01, field changed from Password to UserPassword
    'Mo Morris 20/9/01 Db Audit (UserCode to UserName)
    sSQL = "UPDATE MacroUser SET " _
        & " UserPassword = '" & sPassword & "'," _
        & " FirstLogin = " & sFirstLogin _
        & " WHERE UserName = '" & sUserCode & "'"
        
    SecurityADODBConnection.Execute sSQL, , adCmdText
    
    ' DPH 25/10/2001 - Store new password away to member
    msNewPassword = sPassword
    
    'TA 26/10/2000: log that the password has changed if we are logged into the database
    If Not (MacroADODBConnection Is Nothing) Then
        Call gLog("Change Password", "The password for user " & sUserCode & " was changed")
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "UpdateUserPassword")
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
' Return TRUE if text is valid name for Role description
' Displays any necessary messages
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    IsValidString = False
    
    
    If sDescription > "" Then
        If Not gblnValidString(sDescription, valOnlySingleQuotes) Then
            MsgBox "A password" & gsCANNOT_CONTAIN_INVALID_CHARS, _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        ElseIf Not gblnValidString(sDescription, valAlpha + valNumeric + valSpace) Then
            MsgBox " A password may only contain alphanumeric characters", _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        ElseIf Len(sDescription) > 255 Then
            MsgBox " A password may not be more than 255 characters", _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
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

'----------------------------------------------------------------------------------------'
Public Function NewPassword() As String
'----------------------------------------------------------------------------------------'
' Need to return new password to user class
'----------------------------------------------------------------------------------------'
    NewPassword = msNewPassword
End Function
