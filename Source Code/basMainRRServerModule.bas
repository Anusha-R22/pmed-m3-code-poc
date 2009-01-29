Attribute VB_Name = "basMainRRServerModule"

'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       MainRRServerModule.bas
'   Author:     Nicky Johns November 2000
'   Purpose:    Main module for MACRO Registration/Randomisation Server
'------------------------------------------------------------------------------------------------'
'
'------------------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 22 Nov - Initial development
' MACRO 2.2
'   TA 1/10/2001 - GetApplicationTitle added to MainRRServerModule
' NCJ 10 Oct 01 - Changed Security Path to 2.2
'               Change default error handlers to pass errors upwards
'------------------------------------------------------------------------------------------------'


Option Explicit

' These things are the minimum required to make this DLL compile properly
Public goUser As MACROUser
Public SecurityADODBConnection As ADODB.Connection
Public gnTransactionControlOn As Integer
Public gsHTML_FORMS_LOCATION As String
Public gsSECURE_HTML_LOCATION As String
Public gsAppPath As String
'REM 30/08/02 - added goNewUser
Public goNewUser As MACROUser

' Pretend we have a frmChangePassword just to make the DLL compile
Public frmChangePassword As Control

Public Const gblnRemoteSite = True
Public Const gsDIALOG_TITLE As String = "MACRO"

Public Const valAlpha                   As Integer = 1
Public Const valNumeric                 As Integer = 2
Public Const valSpace                   As Integer = 4
Public Const valOnlySingleQuotes        As Integer = 8
Public Const valComma                   As Integer = 16
Public Const valUnderscore              As Integer = 32
Public Const valDateSeperators          As Integer = 64
Public Const valMathsOperators          As Integer = 128
Public Const valDecimalPoint            As Integer = 256

'---------------------------------------------------------------------
Public Sub InitialisationSettings()
'---------------------------------------------------------------------
    
    'set-up a Application Path variable
    gsAppPath = App.Path
                                
    AddDirSep gsAppPath
    
    'ZA/ASH 10/09/2002 Initialise IMEDSettings component
    InitialiseSettingsFile
    
End Sub

'---------------------------------------------------------------------
Public Function MACROCodeErrorHandler(nTrappedErrNum As Long, sTrappedErrDesc As String, _
                            sProcName As String, sModuleName As String) As OnErrorAction
'---------------------------------------------------------------------
' NCJ 10 Oct 01 - Simply re-raise the error for MACRO 2.2
'---------------------------------------------------------------------
    
    Err.Raise nTrappedErrNum, , sTrappedErrDesc & "|" & sModuleName & "." & sProcName

End Function

'---------------------------------------------------------------------
Public Function MACROFormErrorHandler(oForm As Form, nTrappedErrNum As Long, _
        sTrappedErrDesc As String, sProcName As String) As OnErrorAction
'---------------------------------------------------------------------
' NCJ 10 Oct 01 - Simply re-raise the error for MACRO 2.2
'---------------------------------------------------------------------
    
    Err.Raise nTrappedErrNum, , sTrappedErrDesc & "|" & oForm.Name & "." & sProcName

End Function

Public Function MACROErrorHandler(sObjectName As String, nTrappedErrNum As Long, _
            sTrappedErrDesc As String, sProcName As String, sSource As String) As OnErrorAction
            
            
    MACROErrorHandler = MACROCodeErrorHandler(nTrappedErrNum, sTrappedErrDesc, sProcName, sProcName)
    
End Function

Public Sub ExitMACRO()

    'Do nothing
    
End Sub


Public Sub MACROEnd()

    'Do nothing

End Sub

'----------------------------------------------------------------------------------------'
Public Function GetMacroRegistryKey() As String
'----------------------------------------------------------------------------------------'
' Get the name of the Registry key for the Security database path
' NCJ 10 Oct 01 - Changed to 2.2
'REM 30/08/02 - Changed for MACRO 3.0
'----------------------------------------------------------------------------------------'

    GetMacroRegistryKey = "Software\InferMed Limited\MACRO\3.0"

End Function

'----------------------------------------------------------------------------------------'
Public Function GetApplicationTitle() As String
'----------------------------------------------------------------------------------------'
'Return the default title of an app
'----------------------------------------------------------------------------------------'

    Select Case App.Title
    Case "MACRO_SD"
        If LCase$(Command) = "library" Then
            GetApplicationTitle = "MACRO Library Management"
        Else
            GetApplicationTitle = "MACRO Study Definition"
        End If
    Case "MACRO_DM"
         If LCase$(Command) = "review" Then
            GetApplicationTitle = "MACRO Data Review"
        Else
            'TA 23/10/2000 changend from Data Management"
            GetApplicationTitle = "MACRO Data Entry"
        End If
    Case "MACRO_EX"
         GetApplicationTitle = "MACRO Exchange"
    Case "MACRO_SM"
        GetApplicationTitle = "MACRO System Management"
    Case Else
        GetApplicationTitle = "MACRO"
    End Select
    
End Function


'----------------------------------------------------------------------------------------'
Public Property Get SecurityDatabasePath() As String
'----------------------------------------------------------------------------------------'


    SecurityDatabasePath = GetMACROSetting("SecurityPath", DefaultSecurityDatabasePath)
    If SecurityDatabasePath <> "" Then
        SecurityDatabasePath = DecryptString(SecurityDatabasePath)
    End If

    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let SecurityDatabasePath(sNewSecurityPath As String)
'----------------------------------------------------------------------------------------'
' NCJ 24/1/00 - Allow setting of default path (e.g. from System Management)
' Assume sNewSecurityPath is valid
' Note this does NOT change the local value
' ASH 12/9/2002 - Registry keys replaced with calls to new Settings file
'----------------------------------------------------------------------------------------'
Dim sRegPath As String
    
    If sNewSecurityPath <> "" Then
        'ASH 11/9/2002 Now using new IMEDSettings component to add to settings file
        Call SetMACROSetting("SecurityPath", EncryptString(sNewSecurityPath))
    End If
    
End Property


'----------------------------------------------------------------------------------------'
Public Property Get DefaultSecurityDatabasePath() As String
'----------------------------------------------------------------------------------------'
' Get MACRO's default security path (i.e. the one set on installation)
'----------------------------------------------------------------------------------------'

'TA 06/02/2002: VTRACK Changes for default security path
#If VTRACK = 1 Then
    DefaultSecurityDatabasePath = App.Path & "\databases\VTRACKSecurity.mdb"
#Else
    DefaultSecurityDatabasePath = ""
#End If
    
End Property
