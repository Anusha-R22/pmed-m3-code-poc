VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PasswordPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       PasswordPolicy.cls
'   Author:     Richard Meinesz, September 2002
'   Purpose:    Used to get and set the password policy parameters.
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
'   Revisions:
'----------------------------------------------------------------------------------------'

Option Explicit

Private mnMinPswdLength As Integer
Private mnMaxPswdLength  As Integer
Private mnExpiryPeriod As Integer
Private mbEnforceMixedCase As Boolean
Private mbEnforceDigit As Boolean
Private mbAllowRepeatChars As Boolean
Private mbAllowUserName As Boolean
Private mnPasswordHistory As Integer
Private mnPasswordRetries As Integer

'--------------------------------------------------------------------
Public Property Get MinPswdLength() As Integer
'--------------------------------------------------------------------
    
    MinPswdLength = mnMinPswdLength

End Property

'--------------------------------------------------------------------
Public Property Let MinPswdLength(nMinPswdLength As Integer)
'--------------------------------------------------------------------
    
    If nMinPswdLength <> mnMinPswdLength Then
        mnMinPswdLength = nMinPswdLength
    End If
    
End Property

'--------------------------------------------------------------------
Public Property Get MaxPswdLength() As Integer
'--------------------------------------------------------------------

    MaxPswdLength = mnMaxPswdLength

End Property

'--------------------------------------------------------------------
Public Property Let MaxPswdLength(nMaxPswdLength As Integer)
'--------------------------------------------------------------------

    If nMaxPswdLength <> mnMaxPswdLength Then
        mnMaxPswdLength = nMaxPswdLength
    End If

End Property

'--------------------------------------------------------------------
Public Property Get EnforceMixedCase() As Boolean
'--------------------------------------------------------------------

    EnforceMixedCase = mbEnforceMixedCase

End Property

'--------------------------------------------------------------------
Public Property Let EnforceMixedCase(bEnforceMixedCase As Boolean)
'--------------------------------------------------------------------
    
    If bEnforceMixedCase <> mbEnforceMixedCase Then
        mbEnforceMixedCase = bEnforceMixedCase
    End If
    
End Property

'--------------------------------------------------------------------
Public Property Get EnforceDigit() As Boolean
'--------------------------------------------------------------------

    EnforceDigit = mbEnforceDigit

End Property

'--------------------------------------------------------------------
Public Property Let EnforceDigit(bEnforceDigit As Boolean)
'--------------------------------------------------------------------

    If bEnforceDigit <> mbEnforceDigit Then
        mbEnforceDigit = bEnforceDigit
    End If

End Property

'--------------------------------------------------------------------
Public Property Get AllowRepeatChars() As Boolean
'--------------------------------------------------------------------

    AllowRepeatChars = mbAllowRepeatChars

End Property

'--------------------------------------------------------------------
Public Property Let AllowRepeatChars(bAllowRepeatChars As Boolean)
'--------------------------------------------------------------------

        mbAllowRepeatChars = bAllowRepeatChars

End Property

'--------------------------------------------------------------------
Public Property Get AllowUserName() As Boolean
'--------------------------------------------------------------------

    AllowUserName = mbAllowUserName

End Property

'--------------------------------------------------------------------
Public Property Let AllowUserName(bAllowUserName As Boolean)
'--------------------------------------------------------------------

    mbAllowUserName = bAllowUserName

End Property

'--------------------------------------------------------------------
Public Property Get CheckPrevPswd() As Boolean
'--------------------------------------------------------------------

    CheckPrevPswd = (mnPasswordHistory <> 0)

End Property

'--------------------------------------------------------------------
Public Property Let CheckPrevPswd(bCheckPrevPswd As Boolean)
'--------------------------------------------------------------------

    mnPasswordHistory = -CInt(bCheckPrevPswd)

End Property

'--------------------------------------------------------------------
Public Property Get PasswordHistory() As Integer
'--------------------------------------------------------------------

    PasswordHistory = mnPasswordHistory

End Property

'--------------------------------------------------------------------
Public Property Let PasswordHistory(nPasswordHistory As Integer)
'--------------------------------------------------------------------

    If nPasswordHistory <> mnPasswordHistory Then
        mnPasswordHistory = nPasswordHistory
    End If

End Property

'--------------------------------------------------------------------
Public Property Get RequirePswdExpiry() As Boolean
'--------------------------------------------------------------------

    RequirePswdExpiry = (mnExpiryPeriod <> 0)

End Property

'--------------------------------------------------------------------
Public Property Let RequirePswdExpiry(bRequirePswdExpiry As Boolean)
'--------------------------------------------------------------------

    mnExpiryPeriod = -CInt(bRequirePswdExpiry)

End Property

'--------------------------------------------------------------------
Public Property Get ExpiryPeriod() As Integer
'--------------------------------------------------------------------

    ExpiryPeriod = mnExpiryPeriod

End Property

'--------------------------------------------------------------------
Public Property Let ExpiryPeriod(nExpiryPeriod As Integer)
'--------------------------------------------------------------------

    If nExpiryPeriod <> mnExpiryPeriod Then
        mnExpiryPeriod = nExpiryPeriod
    End If
    
End Property

'--------------------------------------------------------------------
Public Property Get RequireAccountLockout() As Boolean
'--------------------------------------------------------------------

    RequireAccountLockout = (mnPasswordRetries <> 0)

End Property

'--------------------------------------------------------------------
Public Property Let RequireAccountLockout(bRequireAccountLockout As Boolean)
'--------------------------------------------------------------------

    mnPasswordRetries = -CInt(bRequireAccountLockout)

End Property

'--------------------------------------------------------------------
Public Property Get PasswordRetries() As Integer
'--------------------------------------------------------------------

    PasswordRetries = mnPasswordRetries

End Property

'--------------------------------------------------------------------
Public Property Let PasswordRetries(nPasswordRetries As Integer)
'--------------------------------------------------------------------

    If nPasswordRetries <> mnPasswordRetries Then
        mnPasswordRetries = nPasswordRetries
    End If

End Property

'--------------------------------------------------------------------
Public Sub Load(oSecCon As Connection)
'--------------------------------------------------------------------
'REM 16/09/02
'Load all the password policy paramaters from the MACROPassword table
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsPswdPolicy As ADODB.Recordset

    sSQL = "SELECT * FROM MACROPassword"
    Set rsPswdPolicy = New ADODB.Recordset
    rsPswdPolicy.Open sSQL, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText

    mnMinPswdLength = rsPswdPolicy!MinLength
    mnMaxPswdLength = rsPswdPolicy!MaxLength
    mnExpiryPeriod = rsPswdPolicy!ExpiryPeriod
    mbEnforceMixedCase = (rsPswdPolicy!EnforceMixedCase <> 0)
    mbEnforceDigit = (rsPswdPolicy!EnforceDigit <> 0)
    mbAllowRepeatChars = (rsPswdPolicy!AllowRepeatChars <> 0)
    mbAllowUserName = (rsPswdPolicy!AllowUserName <> 0)
    mnPasswordHistory = rsPswdPolicy!PasswordHistory
    mnPasswordRetries = rsPswdPolicy!PasswordRetries
    
    rsPswdPolicy.Close
    Set rsPswdPolicy = Nothing
    
    
End Sub

'--------------------------------------------------------------------
Public Sub Save(ByVal oSecCon As Connection)
'--------------------------------------------------------------------
'REM 16/09/02
'Update the Password Policy
'--------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo Errorlabel
    
    'user -CInt to convert a boolean into an integer
    sSQL = "UPDATE MACROPassword SET" _
        & " MinLength = " & mnMinPswdLength & "," _
        & " MaxLength = " & mnMaxPswdLength & "," _
        & " ExpiryPeriod = " & mnExpiryPeriod & "," _
        & " EnforceMixedCase = " & -CInt(mbEnforceMixedCase) & "," _
        & " EnforceDigit = " & -CInt(mbEnforceDigit) & "," _
        & " AllowRepeatChars = " & -CInt(mbAllowRepeatChars) & "," _
        & " AllowUserName = " & -CInt(mbAllowUserName) & "," _
        & " PasswordHistory = " & mnPasswordHistory & "," _
        & " PasswordRetries = " & mnPasswordRetries

    oSecCon.Execute sSQL, adOpenKeyset, adCmdText
    
Exit Sub
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "PasswordPolicy.Save"
End Sub




