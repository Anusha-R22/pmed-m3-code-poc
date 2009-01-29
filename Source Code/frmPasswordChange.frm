VERSION 5.00
Begin VB.Form frmPasswordChange 
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
Attribute VB_Name = "frmPasswordChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999-2003. All Rights Reserved
'   File:       frmPasswordChange.frm
'   Author:     Will Casey, September 1999
'   Purpose:    To allow the user to create a new password for themselves when
'               their old one has run out.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' Revisions:
' NCJ 12 Jun 03 - Added Hourglass around change password (MACRO 3.0 Bug 1710)
'
'----------------------------------------------------------------------------------------'

Option Explicit
Private mbSuccess As Boolean
Private mconSecurity As ADODB.Connection
Private msSecurityCon As String
Private msUserName As String
Private msNewPassword As String
Private moMACROUser As MACROUser

'----------------------------------------------------------------------------------------'
Public Function Display(oMACROUser As MACROUser, sSecurityCon As String, Optional ByRef sNewPassword As String = "", Optional sOldPassword As String = "") As Boolean
'----------------------------------------------------------------------------------------'
' Display Change Password form
' Input:
'   sUser - usercode
' Output:
'   function - successful change?
'----------------------------------------------------------------------------------------'
    
    mbSuccess = False
    msUserName = oMACROUser.UserName
    msSecurityCon = sSecurityCon

    Set moMACROUser = oMACROUser
    
    ' Place old password in text box and disable it
    txtOldPassword.Text = sOldPassword
    If sOldPassword <> "" Then
        txtOldPassword.Enabled = False
    End If
    
    cmdOK.Enabled = False
    Me.Icon = frmMenu.Icon
    fraPassword.Caption = "User - " & msUserName
    
    FormCentre Me
    Me.Show vbModal
    
    sNewPassword = msNewPassword
    ' Set return for display
    Display = mbSuccess

End Function

'----------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'----------------------------------------------------------------------------------------'
' Unloads the form
'----------------------------------------------------------------------------------------'
    
    mbSuccess = False
    Unload Me
   
End Sub

'----------------------------------------------------------------------------------------'
Private Sub ChangeThePassword()
'----------------------------------------------------------------------------------------'
' Try to change the password based on what they've entered
'----------------------------------------------------------------------------------------'
Dim sNewPassword As String
Dim sOldPassword As String
Dim sConfirm As String
Dim sHashedPassword As String
Dim sPasswordCreated As String
Dim oSystemMessage As SysMessages
Dim sMessageParameters As String
Dim sFirstLogin As String
Dim nCount As Integer
Dim conMACRO As ADODB.Connection
Dim oDatabase As MACROUserBS30.Database
Dim vDatabases As Variant
Dim i As Integer
Dim sDatabaseCode As String
Dim sErrorMsg As String
'ASH 7/2/03
Dim sMessage As String
Dim oUser As MACROUser

    On Error GoTo ErrLabel

    'Check that the new password and confirm password match
    sNewPassword = txtNewPassword.Text
    sConfirm = txtConfirm.Text
    sOldPassword = txtOldPassword.Text
    
    'ASH 7/2/03 Do not do similar checks since all done when user logged in
    'only check if user password entered is correct
    Set oUser = New MACROUser
    If oUser.Login(msSecurityCon, msUserName, sOldPassword, DefaultHTMLLocation, GetApplicationTitle, sMessage, True) = LoginResult.Failed Then
        Call DialogError(sMessage)
        Exit Sub
    End If
    
    If sNewPassword <> sConfirm Then
        Call DialogInformation("The new password does not match the confirmed password")
        txtNewPassword.SetFocus
        txtNewPassword.SelStart = 0
        txtNewPassword.SelLength = Len(txtNewPassword.Text)
    Else
        
        'does all password policy checks and returns (ByRef) the new hashed password and its create date
        mbSuccess = moMACROUser.ChangeUserPassword(msUserName, sNewPassword, sMessage, _
                                                sHashedPassword, sPasswordCreated)
        
        'if change password returns false then display return message and set focus back to new password field
        If Not mbSuccess Then
            Call DialogInformation(sMessage)
            txtNewPassword.SetFocus
            txtNewPassword.SelStart = 0
            txtNewPassword.SelLength = Len(txtNewPassword.Text)
            Call moMACROUser.gLog(msUserName, gsCHANGE_PSWD, _
                                "Change password for user " & msUserName & " failed. " & sMessage)
        Else
            msNewPassword = sNewPassword
            sFirstLogin = SQLStandardNow    ' LastLogin is the same
        
            Call moMACROUser.gLog(msUserName, gsCHANGE_PSWD, _
                                msUserName & " password successfully changed")
            'add the new password to the message table
            sMessageParameters = msUserName & gsPARAMSEPARATOR & sHashedPassword & gsPARAMSEPARATOR & sFirstLogin & gsPARAMSEPARATOR & sFirstLogin & gsPARAMSEPARATOR & sPasswordCreated

            Set mconSecurity = New ADODB.Connection
            'setup security connection
            mconSecurity.Open msSecurityCon
            mconSecurity.CursorLocation = adUseClient
            
            vDatabases = UserDatabasesList(nCount)
            
            Select Case nCount
            Case 0 'no user databases
            
            Case Else 'enter message into all user databases
                For i = 0 To UBound(vDatabases, 2)
                    'database code
                    sDatabaseCode = vDatabases(0, i)
                    Set oDatabase = New MACROUserBS30.Database
                    If oDatabase.Load(mconSecurity, "", sDatabaseCode, "", False, sErrorMsg) Then
                        'create DB connection from DB connection string
                        Set conMACRO = CreateMACROCon(oDatabase.ConnectionString)
                        'if connection fails don't enter message
                        If Not conMACRO Is Nothing Then
                            If Not ExcludeUserRDE(msUserName) Then 'don't write message if its rde
                                Set oSystemMessage = New SysMessages
                                Call oSystemMessage.AddNewSystemMessage(conMACRO, ExchangeMessageType.PasswordChange, msUserName, msUserName, "Change Password", sMessageParameters)
                                Set oSystemMessage = Nothing
                            End If
                        End If
                        
                    End If
                Next
            
            End Select
            'TA 08/04/2003: inform user that change is successful
            DialogInformation "Password successfully changed"
        End If
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmChangePassword.ChangeThePassword"

End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'----------------------------------------------------------------------------------------'
' does the boundary checks, if everything is ok then change the password
' NCJ 12 Jun 03 - Farmed out code to ChangeThePassword, and added Hourglass around it (Bug 1710)
'----------------------------------------------------------------------------------------'

    On Error GoTo Errhandler
    
    HourglassOn
    Call ChangeThePassword
    HourglassOff
    
    ' Unload if successful
    If mbSuccess Then Unload Me
    
Exit Sub
Errhandler:
    If MACROErrorHandler("frmPasswordChange", Err.Number, Err.Description, "cmdOK_Click", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Function CreateMACROCon(sConnection As String) As ADODB.Connection
'---------------------------------------------------------------------
'REM 13/01/03
'Create a MACRO database connection, if it fails will return nothing
'---------------------------------------------------------------------
Dim conMACRO As ADODB.Connection

    On Error GoTo ErrLabel
        Set conMACRO = New ADODB.Connection
        conMACRO.Open sConnection
        conMACRO.CursorLocation = adUseClient
        
        Set CreateMACROCon = conMACRO

Exit Function
ErrLabel:
    Set CreateMACROCon = Nothing
End Function

'---------------------------------------------------------------------
Private Function UserDatabasesList(ByRef nCount As Integer) As Variant
'---------------------------------------------------------------------
'REM 06/12/02
'Loads the database list box with all the users databases
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsDBs As ADODB.Recordset
Dim vDatabases As Variant
Dim i As Integer
Dim sDatabaseCode As String

    
    'get list of user databases
    sSQL = "SELECT DatabaseCode FROM UserDatabase" _
        & " WHERE UserName = '" & msUserName & "'"
    Set rsDBs = New ADODB.Recordset
    rsDBs.Open sSQL, mconSecurity, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsDBs.RecordCount <> 0 Then
        nCount = rsDBs.RecordCount
        vDatabases = rsDBs.GetRows
    Else
        nCount = 0
        vDatabases = Null
    End If
    
    UserDatabasesList = vDatabases
    
    rsDBs.Close
    Set rsDBs = Nothing
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub Form_Initialize()
'----------------------------------------------------------------------------------------'
' Initialization
'----------------------------------------------------------------------------------------'

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Load()
'----------------------------------------------------------------------------------------'
'Disable the ok button
'----------------------------------------------------------------------------------------'
    
    On Error GoTo Errorlabel
              
    Me.BackColor = glFormColour
   
    Me.Icon = frmMenu.Icon
    HourglassSuspend    'must turn off in form_unload
    
Exit Sub
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmChangePassword.Form_Load"
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

    On Error GoTo Errorlabel

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
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmChangePassword.txtConfirm_Change"
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtConfirm_GotFocus()
'----------------------------------------------------------------------------------------'

        txtConfirm.SelStart = 0
        txtConfirm.SelLength = Len(txtConfirm.Text)
        
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtNewPassword_Change()
'----------------------------------------------------------------------------------------'
' Disable the ok button until all three textboxes have entries
'----------------------------------------------------------------------------------------'
Dim sDescription As String

    On Error GoTo Errorlabel
 
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
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmChangePassword.txtNewPassword_Change"
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtOldPassword_Change()
'----------------------------------------------------------------------------------------'
' Disable the ok button until all three textboxes have entries
'----------------------------------------------------------------------------------------'
Dim sDescription As String
 
    On Error GoTo Errorlabel
    
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
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmChangePassword.txtOldPassword_Change"
End Sub

'----------------------------------------------------------------------------------------'
Private Function IsValidString(sDescription As String) As Boolean
'----------------------------------------------------------------------------------------'
' Return TRUE if text is valid name for Role description
' Displays any necessary messages
'----------------------------------------------------------------------------------------'

    On Error GoTo Errorlabel
    
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
'
Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmChangePassword.IsValidString"
End Function

