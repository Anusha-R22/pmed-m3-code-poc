VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2003-2003, All Rights Reserved
'   File:       Password.cls
'   Author:     Richard Meinesz, September 2002
'   Purpose:    Used to check user passwords and returns success or a message about failure.
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
' Revisions:
' NCJ 14 Feb 03 - Change to PasswordExpiryCheck to work with Regional Settings
' NCJ 27 Nov 03 - Added class Terminate
' REM 05/12/03 - Added MACROSettings function to ChangePassword so setting "expirepassword" can be used to detemine
'   if a new user's password or a password reset by a sys admin should be automatically expired.
' TA cdb2641 02/11/2005: sql is eloborate in GetPasswordHistory to prevent duplicate password created, passwordhistory rows crashing MACRO
'----------------------------------------------------------------------------------------'

Option Explicit

Public Enum ePasswordCheck
    pcPass = 1
    pcAccountDisabled = 2
    pcFailed = 3
    pcAccountLockout = 4
End Enum

Private msUserName As String
Private msUserNameFull As String
Private msPassword As String
Private mnUserEnabled As Integer
Private mdLastLogin As Double
Private mdFirstLogin As Double
Private msDefaultUserRole As String
Private mnFailedLoginAttempts As Integer
Private mdPasswordCreated As Double
Private mnSysAdmin As Integer

Private moPswdPolicy As PasswordPolicy

'----------------------------------------------------------------------------------------'
Friend Property Get UserName() As String
'----------------------------------------------------------------------------------------'
' NB This is actually the User Code
'----------------------------------------------------------------------------------------'
'Changed Mo Morris 2/5/00   When Trial Offices /AI Auto import or /TR are being used, there
'needs to be a UserName available when DataItemResponse tables are changed/updated. These will
'now get a UserName of "AutoImport". (Note: it was thought that /TR would not get used much
'and did not need to be reflected in the choice of UserName)
'----------------------------------------------------------------------------------------'

    If RTrim(msUserName) = "" Then
        UserName = "AutoImport"
    Else
        UserName = msUserName
    End If

End Property

'----------------------------------------------------------------------------------------'
Friend Property Let UserName(sUserName As String)
'----------------------------------------------------------------------------------------'

    If sUserName <> msUserName Then
        msUserName = sUserName
    End If

End Property

'----------------------------------------------------------------------------------------'
Friend Property Get UserNameFull() As String
'----------------------------------------------------------------------------------------'

    UserNameFull = msUserNameFull

End Property

'----------------------------------------------------------------------------------------'
Friend Property Let UserNameFull(sUserNameFull As String)
'----------------------------------------------------------------------------------------'

    If sUserNameFull <> msUserNameFull Then
        msUserNameFull = sUserNameFull
    End If

End Property

'----------------------------------------------------------------------------------------'
Friend Property Get Password() As String
'----------------------------------------------------------------------------------------'

    Password = msPassword

End Property

'----------------------------------------------------------------------------------------'
Property Let Password(sPassword As String)
'----------------------------------------------------------------------------------------'

    If sPassword <> msPassword Then
        msPassword = sPassword
    End If

End Property

'----------------------------------------------------------------------------------------'
Friend Property Get UserEnabled() As Integer
'----------------------------------------------------------------------------------------'

    UserEnabled = mnUserEnabled

End Property

'----------------------------------------------------------------------------------------'
Friend Property Let UserEnabled(nUserEnabled As Integer)
'----------------------------------------------------------------------------------------'

    If nUserEnabled <> mnUserEnabled Then
        mnUserEnabled = nUserEnabled
    End If

End Property

'----------------------------------------------------------------------------------------'
Friend Property Get SysAdmin() As Integer
'----------------------------------------------------------------------------------------'

    SysAdmin = mnSysAdmin
    
End Property

'----------------------------------------------------------------------------------------'
Friend Property Get LastLogin() As Double
'----------------------------------------------------------------------------------------'

    LastLogin = mdLastLogin

End Property

'----------------------------------------------------------------------------------------'
Friend Property Let LastLogin(dLastLogin As Double)
'----------------------------------------------------------------------------------------'

    If dLastLogin <> mdLastLogin Then
        mdLastLogin = dLastLogin
    End If

End Property

'----------------------------------------------------------------------------------------'
Friend Property Get FirstLogin() As Double
'----------------------------------------------------------------------------------------'

    FirstLogin = mdFirstLogin

End Property

'----------------------------------------------------------------------------------------'
Friend Property Let FirstLogin(dFirstLogin As Double)
'----------------------------------------------------------------------------------------'

    If dFirstLogin <> mdFirstLogin Then
        mdFirstLogin = dFirstLogin
    End If

End Property

'----------------------------------------------------------------------------------------'
Friend Property Get DefaultUserRole() As String
'----------------------------------------------------------------------------------------'

    DefaultUserRole = msDefaultUserRole

End Property

'----------------------------------------------------------------------------------------'
Friend Property Let DefaultUserRole(sDefaultUserRole As String)
'----------------------------------------------------------------------------------------'

    If sDefaultUserRole <> msDefaultUserRole Then
        msDefaultUserRole = sDefaultUserRole
    End If

End Property

'----------------------------------------------------------------------------------------'
Friend Property Get FailedLoginAttempts() As Integer
'----------------------------------------------------------------------------------------'

    FailedLoginAttempts = mnFailedLoginAttempts

End Property

'----------------------------------------------------------------------------------------'
Friend Property Let FailedLoginAttempts(nFailedLoginAttempts As Integer)
'----------------------------------------------------------------------------------------'

    If nFailedLoginAttempts <> mnFailedLoginAttempts Then
        mnFailedLoginAttempts = nFailedLoginAttempts
    End If

End Property

'----------------------------------------------------------------------------------------'
Friend Property Get PasswordCreated() As Double
'----------------------------------------------------------------------------------------'

    PasswordCreated = mdPasswordCreated

End Property

'----------------------------------------------------------------------------------------'
Friend Property Let PasswordCreated(dPasswordCreated As Double)
'----------------------------------------------------------------------------------------'

    If dPasswordCreated <> mdPasswordCreated Then
        mdPasswordCreated = dPasswordCreated
    End If

End Property

'----------------------------------------------------------------------------------------'
Friend Function UserExists(oSecCon As Connection, sSecCon As String, ByRef sUserName As String, ByRef sMessage As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 14/09/02
'Check the existance of a UserName, if exists set up all user properties
'REM 14/01/03 - check user name in Uppercase in Oracle as it is case sensitive
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsUser As ADODB.Recordset

    On Error GoTo Errorlabel

    'get UserName from the MacroUser table
    'REM 14/01/03 - check user name in Uppercase in Oracle as it is case sensitive
    Select Case Connection_Property(CONNECTION_PROVIDER, sSecCon)
    Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
        sSQL = "SELECT * FROM MACROUser WHERE upper(UserName) = upper('" & sUserName & "')"
    Case Else
        sSQL = "SELECT * FROM MACROUser WHERE UserName = '" & sUserName & "'"
    End Select
    
    
    'set up class level recordset for this user
    Set rsUser = New ADODB.Recordset
    rsUser.Open sSQL, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsUser.RecordCount > 0 Then
        'User does exist
        UserExists = True
        sMessage = ""
        
        'set up all user properties
        sUserName = rsUser!UserName
        msUserName = sUserName
        msUserNameFull = rsUser!UserNameFull
        msPassword = rsUser!UserPassword
        mnUserEnabled = rsUser!Enabled
        mdLastLogin = rsUser!LastLogin
        mdFirstLogin = rsUser!FirstLogin
        msDefaultUserRole = RemoveNull(rsUser!DefaultUserRoleCode)
        mnFailedLoginAttempts = rsUser!FailedAttempts
        mdPasswordCreated = rsUser!PasswordCreated
        mnSysAdmin = rsUser!SysAdmin
        
        'set up the password policy
        Set moPswdPolicy = New PasswordPolicy
        moPswdPolicy.Load oSecCon
        
    Else
        sMessage = "User name does not exist."
        UserExists = False

    End If
    
    rsUser.Close
    Set rsUser = Nothing
    
Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Password.UserNameExits"
End Function

'----------------------------------------------------------------------------------------'
Friend Function PasswordCheck(oSecCon As Connection, ByVal sPassword As String, ByRef sMessage As String) As ePasswordCheck
'----------------------------------------------------------------------------------------'
'REM 16/09/02
'Check to see - that the user password is correct
'               whether account is disabled or locked out
'----------------------------------------------------------------------------------------'
Dim nFailedAttempts As Integer
    
    On Error GoTo Errorlabel
    
    sMessage = ""
    
    'if password does not match
    If msPassword <> sPassword Then
        'return message
        sMessage = "Incorrect Password"
        
        'password incorrect
        PasswordCheck = ePasswordCheck.pcFailed
        
        If mnSysAdmin = 0 Then 'if the user is a system adminstrator (1) then don't do lockout check
        
            'see if policy requires account lockout check
            If moPswdPolicy.RequireAccountLockout Then
            
                'get the current number of failed attempts
                nFailedAttempts = mnFailedLoginAttempts
                
                'add 1 to the number of failed attempts
                nFailedAttempts = nFailedAttempts + 1
                
                'if the number of failed attempts is greater than or equal to that set then disable account
                If nFailedAttempts >= moPswdPolicy.PasswordRetries Then
                    sMessage = "User account has been locked. User attempted to login with incorrect password more than " & moPswdPolicy.PasswordRetries & " times."
                    PasswordCheck = ePasswordCheck.pcAccountLockout
                End If
                mnFailedLoginAttempts = nFailedAttempts
            End If
        
            'set property to new value
            mnFailedLoginAttempts = nFailedAttempts
            
            'update the users failed attempts and enabled status
            Call UpdateFailedAttempts(oSecCon)
        End If
        
    Else
        'check to see if user account is disabled
        If mnUserEnabled = 0 Then
            sMessage = "User account is disabled"
            PasswordCheck = ePasswordCheck.pcAccountDisabled
            Exit Function
        End If
        
        'see if policy requires account lockout check, if the user is a system adminstrator (1) then don't do lockout check
        If (moPswdPolicy.RequireAccountLockout) And (mnSysAdmin = 0) Then
            'check to see if user account is locked out
            If mnFailedLoginAttempts >= moPswdPolicy.PasswordRetries Then
                sMessage = "User account is locked out"
                PasswordCheck = ePasswordCheck.pcAccountLockout
                Exit Function
            End If
        End If
        
        'set the failed attempts property to 0 when successful login
        mnFailedLoginAttempts = 0
        'update the user login; LastLogin, Failed Attempts fields
        Call UpdateUserLogin(oSecCon)
        PasswordCheck = ePasswordCheck.pcPass
    End If
    
Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Password.PasswordExpiryCheck"
End Function

'----------------------------------------------------------------------------------------'
Friend Function PasswordExpiryCheck(ByVal sUserName As String, ByRef lValidDaysLeft As Long) As Boolean
'----------------------------------------------------------------------------------------'
'REM 16/09/02
'Check to see if a password expiry check required, if so returns true
' and returns the number of days to expiry (nValidDaysLeft)
'----------------------------------------------------------------------------------------'
Dim lExpiryPeriod As Long
Dim lDifference As Long
'Dim sNow As String

    On Error GoTo Errorlabel
    
    'see if require password expiry check, and that user is not a system administrator
    If (moPswdPolicy.RequirePswdExpiry) And (mnSysAdmin = 0) Then
    
        lExpiryPeriod = moPswdPolicy.ExpiryPeriod
        
'        sNow = LocalNumToStandard(IMedNow)
        ' NCJ 14 Feb 03 - Do not convert IMedNow to "standard" string!
        ' Find no. of days from first login with this password to today's date
        lDifference = DateDiff("d", mdFirstLogin, IMedNow)
        
        ' Find no. of valid days left with this password
        ' (may be negative if password has expired)
        lValidDaysLeft = lExpiryPeriod - lDifference
    
        PasswordExpiryCheck = True
    
    Else 'don't require check
        lValidDaysLeft = 0
        PasswordExpiryCheck = False
    End If

    
Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Password.PasswordExpiryCheck"
End Function

'----------------------------------------------------------------------------------------'
Friend Function ChangePassword(oSecCon As Connection, ByVal sUserName As String, sNewPassword, _
                               ByRef sMessage As String, Optional ByRef sHashedPassword As String = "", _
                               Optional ByRef sCreateDate As String = "", Optional bUpdatePassword As Boolean = True) As Boolean
'----------------------------------------------------------------------------------------'
'REM 16/09/02
'Does checks to see if the new password conforms to the password policy.
'Checks - Trailing and leading spaces, MinLength, MaxLength, EnforcedMixedCase,
'         EnforcedDigit, AllowRepeatChars, AllowUserName, PasswordHistory
'If bUpdatepassword is set to false will only check password against policy
'REVISIONS:
' REM 05/12/03 - Added bExpirePassword, which is set if bUpdatePassword is true.  If sMessage is passed in as "resetpassword" then
'   bExpirePassword = true, else false
'----------------------------------------------------------------------------------------'
Dim nMinLength As Integer
Dim nMaxLength As Integer
Dim bEnforceMixedCase As Boolean
Dim sLetter As String
Dim nAsc As Integer
Dim bEnforceDigit As Boolean
Dim nNumeric As Integer
Dim nUppercase As Integer
Dim nLowerCase As Integer
Dim bAllowRepeatChars As Boolean
Dim sLetterComp As String
Dim bAllowUserName As Boolean
Dim bCheckPrevPswd As Boolean
Dim nPasswordHistory As Integer
Dim colOldPasswords As Collection
Dim nCount As Integer
Dim nFound As Integer
Dim sPasswordAndDate As String
Dim sHashedNewPasswordAndDate As String
Dim i As Integer
Dim colCreateDates As Collection
Dim oPswdPolicy As PasswordPolicy
Dim bExpirePassword As Boolean

    On Error GoTo ErrLabel
    
    Set oPswdPolicy = New PasswordPolicy
    Call oPswdPolicy.Load(oSecCon)
    
    'set all the variables for the password policy
    nMinLength = oPswdPolicy.MinPswdLength
    nMaxLength = oPswdPolicy.MaxPswdLength
    bEnforceMixedCase = oPswdPolicy.EnforceMixedCase
    bEnforceDigit = oPswdPolicy.EnforceDigit
    bAllowRepeatChars = oPswdPolicy.AllowRepeatChars
    bAllowUserName = oPswdPolicy.AllowUserName
    bCheckPrevPswd = oPswdPolicy.CheckPrevPswd
    nPasswordHistory = oPswdPolicy.PasswordHistory
   
   'Do not allow passwords with leading or trailing spaces
   'since passwords are trimmed before stored in database.
   If sNewPassword <> Trim(sNewPassword) Then
        sMessage = "Sorry, passwords cannot contain leading or trailing spaces."
        ChangePassword = False
        Exit Function
    End If
    
    '**PASSWORD LENGTH**
    'Check Max and Min length passwod settings
    If Len(sNewPassword) > nMaxLength Then
        sMessage = " The new password is too long, the maximum number" & vbCr _
             & " of characters for a password is " & nMaxLength & "."
        ChangePassword = False
        Exit Function
    ElseIf Len(sNewPassword) < nMinLength Then
        sMessage = "The new password is too short, the minimum number" & vbCr _
             & " of characters for a password is " & nMinLength & "."
        ChangePassword = False
        Exit Function
    End If
    
    '**MIXED CASE**
    'Enforce Mixed Case Chars in Password
    'Password must contain at least one upper case and at least one lower case character
    If bEnforceMixedCase Then
        nUppercase = 0
        nLowerCase = 0
        'loop through the new password checking for upper and lower case characters
        For i = 1 To Len(sNewPassword)
            sLetter = Mid(sNewPassword, i, 1)
            nAsc = Asc(sLetter)
            'ascii values for "A" to "Z"
            If (nAsc > 64) And (nAsc < 91) Then
                nUppercase = nUppercase + 1
            'ascii values for "z" to "z"
            ElseIf (nAsc > 96) And (nAsc < 123) Then
                nLowerCase = nLowerCase + 1
            End If
        Next
        'check to see if have both cases, if not exit function
        If (nUppercase = 0) Or (nLowerCase = 0) Then
            ChangePassword = False
            sMessage = "The new password does not contain at least one lower case and at least one upper case character."
            Exit Function
        End If
    End If
    
    '**NUMERIC DIGIT**
    'Enforce use of at least one numeric digit in new password
    If bEnforceDigit Then
        nNumeric = 0
        For i = 1 To Len(sNewPassword)
            'if found one numeric digit then exit for
            If nNumeric > 0 Then Exit For
            sLetter = Mid(sNewPassword, i, 1)
            nAsc = Asc(sLetter)
            'ascii values for 0 to 9
            If (nAsc > 47) And (nAsc < 58) Then
                nNumeric = nNumeric + 1
            End If
        Next
        If nNumeric = 0 Then
            ChangePassword = False
            sMessage = "The new password does not contain at least one numerical digit."
            Exit Function
        End If

    End If
    
    '**REPEAT CHARACTERS**
    'if do not allow more than 4 repeat characters
    If Not bAllowRepeatChars Then
        'get first letter from password
        sLetterComp = Mid(sNewPassword, 1, 1)
        nCount = 1
        'loop through rest of password
        For i = 2 To Len(sNewPassword)
            'get next letter
            sLetter = Mid(sNewPassword, i, 1)
            'check if its the same as last letter
            If sLetterComp = sLetter Then 'if it is then add one to nCount
                nCount = nCount + 1
            Else 'else set sLetterComp to sLetter
                sLetterComp = sLetter
                nCount = 1
            End If
            
            If nCount > 4 Then
                sMessage = "The password contains more than four repeats of the same character."
                ChangePassword = False
                Exit Function
            End If
        Next
    End If
    
    
    '**PORTION OF USERNAME**
    'No more than 4 character portion of user name
    'if do not allow 4 portion part of user name then
    If Not bAllowUserName Then
        'loop through new password
        For i = 1 To Len(sNewPassword)
            'get 4 character portion, starting at beginning
            sLetter = Mid(sUserName, i, 5)
            'if the remaing length is less than 4 letter then do not compare
            If Len(sLetter) < 5 Then Exit For
            'if 5 character portion is found in the new password
            nFound = InStr(1, sNewPassword, sLetter, vbTextCompare)
            If nFound > 0 Then
                sMessage = "The password contains more than a 4 letter portion of the user name."
                ChangePassword = False
                Exit Function
            End If
        Next
    
    End If
    
    '**PASSWORD HISTORY**
    'See if need to check against password history
    If bCheckPrevPswd Then
        
        Set colCreateDates = New Collection
        Set colOldPasswords = New Collection
    
        'get collection of old passwords up to the number required (nPasswordHistory)
        Set colOldPasswords = GetPasswordHistory(oSecCon, sUserName, nPasswordHistory, colCreateDates)
        
        For i = 1 To colCreateDates.Count
            
            'add create date to password
            sPasswordAndDate = sNewPassword & colCreateDates(i)
            
            'hash the password and old create date to check it against the password history table
            sHashedNewPasswordAndDate = HashHexEncodeString(sPasswordAndDate)
            
            'if the New Password exists in collection of old ones then exit function
            If CollectionMember(colOldPasswords, sHashedNewPasswordAndDate, False) Then
                
                sMessage = "You may not set your new password to be the " & vbCrLf _
                    & "same as any of your last " & nPasswordHistory & " passwords."
                
                ChangePassword = False
                ' NCJ 27 Nov 03 - Added tidying up
                Set colCreateDates = Nothing
                Set colOldPasswords = Nothing
                Exit Function
            End If
        Next
    End If

    'updates password if doing a password change
    If bUpdatePassword Then
        'REM 05/12/03 - If sMessage has been passsed in as "sysadminreset" then Sys Admin has reset a user's password
        bExpirePassword = (sMessage = "sysadminreset")
            
        Call UpdateUserPassword(oSecCon, sUserName, sNewPassword, sHashedPassword, sCreateDate, bExpirePassword)
        sMessage = "Your password has been successfully changed."
    End If
    
    ChangePassword = True
    
    ' NCJ 27 Nov 03 - Added tidying up
    Set colCreateDates = Nothing
    Set colOldPasswords = Nothing
  
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Password.ChangePassword"
End Function

'--------------------------------------------------------------------------------
Private Sub UpdateUserPassword(oSecCon As Connection, ByVal sUserName As String, ByVal sNewPassword As String, _
                               ByRef sHashedNewPasswordAndDate As String, ByRef sPasswordCreated As String, _
                               ByVal bExpirePassword As Boolean)
'--------------------------------------------------------------------------------
'REM 17/09/02
' update the user password after all checks have been completed
' and insert password into password history table
'REVISIONS:
' REM 05/12/03 - Added MACROSetting "expirepassword", if its true then expire new user passwords and users passowrds reset via sys admin
' REM 05/12/03 - Added new parameter, bResetPassword, to indicate if the UpdateUserPassword is coming from
' a Sys Admin restting a User's password (true) or if a user changed their own password (false)
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim sFirstLogin As String
Dim nHistoryNumber As Integer
Dim sNewPasswordAndDate As String
Dim rsMaxHistoryNo As ADODB.Recordset
Dim dImedNow As Double

    On Error GoTo ErrLabel
    
    dImedNow = IMedNow
    
    'REM 05/12/03 - Added MACROSetting
    'Initialise the Settings File
    InitialiseSettingsFile True
    'If MACRO Setting is true then expire passwords reset by sys admin
    If (GetMACROSetting("expirepassword", "true") = "true") Then
        If bExpirePassword Then
            sFirstLogin = "36000" 'set to an old date to force password to expire
            
        Else
            sFirstLogin = LocalNumToStandard(dImedNow) 'if user is changing their password then don't expire it
        End If
    Else 'don't expire reset password
        sFirstLogin = LocalNumToStandard(dImedNow)
    End If
    
    sPasswordCreated = LocalNumToStandard(dImedNow)
    
    sNewPasswordAndDate = sNewPassword & sPasswordCreated
    
    'hash the new password and create date
    sHashedNewPasswordAndDate = HashHexEncodeString(sNewPasswordAndDate)
    
    'Update password, first login date and PasswordCreated date in MACROUser table
    sSQL = "UPDATE MACROUser SET " _
        & " UserPassword = '" & sHashedNewPasswordAndDate & "'," _
        & " FirstLogin = " & sFirstLogin & "," _
        & " PasswordCreated = " & sPasswordCreated _
        & " WHERE UserName = '" & sUserName & "'"
    oSecCon.Execute sSQL
    
    'get max passord history number for user
    sSQL = "SELECT MAX (HistoryNumber) as MaxHistoryNumber FROM PasswordHistory" _
        & " WHERE UserName = '" & sUserName & "'"
    Set rsMaxHistoryNo = New ADODB.Recordset
    rsMaxHistoryNo.Open sSQL, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    If IsNull(rsMaxHistoryNo!MaxHistoryNumber) Then
        nHistoryNumber = 1
    Else
        nHistoryNumber = rsMaxHistoryNo!MaxHistoryNumber + 1
    End If
    
    ' NCJ 27 Nov 03 - Close down recordset
    Call rsMaxHistoryNo.Close
    Set rsMaxHistoryNo = Nothing
    
    'insert new password into passwordhistory table
    sSQL = "INSERT INTO PasswordHistory VALUES ('" & sUserName & "'," & nHistoryNumber & "," & sPasswordCreated & ", '" & sHashedNewPasswordAndDate & "')"
    oSecCon.Execute sSQL
    
    'make sure properties get the new values
    msPassword = sHashedNewPasswordAndDate
    mdFirstLogin = dImedNow
    mdPasswordCreated = dImedNow
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Password.UpdateUserPassword"
End Sub

'--------------------------------------------------------------------------------
Friend Sub UpdateUserLogin(oSecCon)
'--------------------------------------------------------------------------------
'REM 17/09/02
'Update the last login time stamp when a user has successfully logged on
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim sLastlogin As String

    On Error GoTo ErrLabel
    
'****TODO - replace with below after added to MACRO****
    'sLastLogin = SQLStandardNow
    sLastlogin = LocalNumToStandard(IMedNow)

    sSQL = "UPDATE MacroUser " _
    & " SET LastLogin = " & sLastlogin & "," _
    & " FailedAttempts = " & mnFailedLoginAttempts _
    & " WHERE UserName = '" & msUserName & "'"
        
    oSecCon.Execute sSQL, adOpenKeyset, adCmdText
 
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Password.UpdateUserLogin"
End Sub

'----------------------------------------------------------------------------------------'
Friend Sub UpdateFailedAttempts(oSecCon As Connection)
'----------------------------------------------------------------------------------------'
'REM 17/09/02
'Updates the Failed Attempts field in the MACROUser table after each failed login attempt
'----------------------------------------------------------------------------------------'
Dim sSQL As String

    On Error GoTo Errorlabel
    
    sSQL = "UPDATE MACROUser SET " _
        & " Enabled = " & mnUserEnabled & "," _
        & " FailedAttempts = " & mnFailedLoginAttempts _
        & " WHERE UserName = '" & msUserName & "'"

    oSecCon.Execute sSQL
    
Exit Sub
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Password.UpdateFailedAttempts"
End Sub

''----------------------------------------------------------------------------------------'
'Private Sub IsPasswordDefault(sUser As String, sPassword As String)
''----------------------------------------------------------------------------------------'
'' check to see if the user who has logged in is the default User
'' UserCode is rde and the password is macrotm
''----------------------------------------------------------------------------------------'
'Dim sMsg As String
'Dim sNewPassword As String
'
'    On Error GoTo ErrLabel
'
'    'reset new password
'    sNewPassword = ""
'
'    If LCase(sUser) = "rde" And LCase(sPassword) = "macrotm" Then
'        sMsg = " You have logged on using the default username and password." & vbCr _
'             & " Would you like to change the default password now?"
'        Select Case MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE)
'                Case vbYes
'                    Call frmChangePassword.Display(sUser, sNewPassword)
'                    'storing any password change in class member
'                    If sNewPassword <> "" Then
'                        Me.Password = sNewPassword
'                    End If
'
'                Case vbNo
'        End Select
'    End If
'
'Exit Sub
'ErrLabel:
'    Err.Raise Err.Number, , Err.Description & "|" & "Password.IsPasswordDefault"
'End Sub

'----------------------------------------------------------------------------------------'
Private Function GetPasswordHistory(oSecCon As Connection, ByVal sUserName As String, _
                                    ByVal nPasswordHistory As Integer, ByRef colCreateDates As Collection) As Collection
'----------------------------------------------------------------------------------------'
'REM 16/09/02
'returns the old passwords from the password history table
'will return passwords from latest old password to oldest up to number specified by nPasswordHistory
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsPswdHistory As ADODB.Recordset
Dim colPasswordHistory As Collection
Dim nCount As Integer
Dim sPassword As String
Dim sLastPassword As String

    Set colPasswordHistory = New Collection
    Set rsPswdHistory = New ADODB.Recordset
    
    'TA cdb2641 02/11/2005: sql is eloborate to prevent duplicate password created, passwordhistory rows crashing MACRO
    ' these should not exist in MACRO but can result if data xfer times out while xferring password changes
    sSQL = "select username,max(historynumber) historynumber, passwordcreated,userpassword" & vbNewLine & _
            " from passwordhistory where username = '" & sUserName & "' group by  username,passwordcreated,userpassword" & vbNewLine & _
            "order by historynumber desc"

    Set rsPswdHistory = New ADODB.Recordset
    rsPswdHistory.Open sSQL, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText

    nCount = 0
    'loop through the recordset of old passwords.  Add all old passwords from latest to oldest
    Do While Not rsPswdHistory.EOF
        'if there are more passwords than required for check then exit loop
        If nCount > nPasswordHistory Then Exit Do
        
        sPassword = rsPswdHistory.Fields(3).Value
        'check to see if passwords match, if so skip it as can't add it to the collection
        If sPassword <> sLastPassword Then
            'Add old passwords to collection, Fields(3) is the password field
            colPasswordHistory.Add sPassword, sPassword
            'Add the create dates of each password to the ByRef collection passed in.  Fields(2) is the Create date field
            colCreateDates.Add rsPswdHistory.Fields(2).Value, CStr(rsPswdHistory.Fields(2).Value)
        End If
        
        sLastPassword = sPassword
        nCount = nCount + 1
        rsPswdHistory.MoveNext
    Loop

    Set GetPasswordHistory = colPasswordHistory
    
    ' NCJ 27 Nov 03 - Tidy up on exit
    Call rsPswdHistory.Close
    Set rsPswdHistory = Nothing
    Set colPasswordHistory = Nothing

End Function

'----------------------------------------------------------------------
Private Function RemoveNull(vVariable As Variant) As String
'----------------------------------------------------------------------

    RemoveNull = ConvertFromNull(vVariable, vbString)

End Function

'----------------------------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------------------------
' NCJ 27 Nov 03 - Tidy up on termination
'----------------------------------------------------------------------

    Set moPswdPolicy = Nothing
    
End Sub
