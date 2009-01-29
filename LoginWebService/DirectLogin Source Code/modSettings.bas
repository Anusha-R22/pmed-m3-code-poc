Attribute VB_Name = "modSettings"
'----------------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2002- 2006 All Rights Reserved
'   File:       modSettings.bas
'   Author:     Ashitei Trebi-Ollennu, September 2002
'   Purpose:    Used to get and set registry keys and values. Makes use of the
'               IMEDSettings.dll to achieve this.
'----------------------------------------------------------------------------------------
'Revisions:
'----------------------------------------------------------------------------------------
'ASH 10/11/2002 - Added ByVal to GetMacroSetting and SetMacroSetting
'ASH 12/9/2002 - Check for readonly file added.
' ic 24/04/2003 added trace constant
' ic 18/08/2003 added passthread constant
' ic 09/10/2003 added compression constant
' ic 07/07/2004 added illegal parameters constant
' NCJ 19 Jan 05 (Issue 2502) Added "Enable list filtering" (for Windows DE)
' DPH 02/02/2005 (PDU2300) Added pdu.exe location setting & "use pdu" setting (for Win DE)
' MLM 04/07/05: bug 2464: Added showschedulesdvmenu setting to control schedule menu items display.
' NCJ 7 Dec 05 - Added Partial Dates setting
' NCJ 12 Jul 06 - Added MultiUserSD setting
'----------------------------------------------------------------------------------------
Option Explicit
Private msFileName As String
Private moSetting As IMEDSettings
Private Const msFIXEDFILENAME = "MACROSettings30.txt"

'constant to store 'last used db' setting
Public Const MACRO_SETTING_LAST_USED_DATABASE = "lastuseddatabase"
Public Const MACRO_SETTING_LAST_USED_ROLE = "lastusedrole"

Public Const MACRO_SETTING_SECPATH As String = "securitypath"
Public Const MACRO_SETTING_WEBPATH As String = "web html"
Public Const MACRO_SETTING_USESSL As String = "usessl"
Public Const MACRO_SETTING_USESCI As String = "usesci"
Public Const MACRO_SETTING_WEBHELPURL As String = "webhelpurl"
Public Const MACRO_SETTING_MAXAREZZO As String = "maxarezzo"
Public Const MACRO_SETTING_TRACE As String = "trace"
' DPH 17/06/2003 - web version setting
Public Const MACRO_SETTING_WEBVERSION As String = "webversion"
'ic 18/08/2003 - pass thread setting
Public Const MACRO_SETTING_PASSTHREAD As String = "passthread"
'ic 09/10/2003 - use compression, originally for schedule
Public Const MACRO_SETTING_USECOMPRESSION As String = "compression"
'ic 07/07/2004 - log errors raised by illegal parameters
Public Const MACRO_SETTING_LOG_ILLEGAL_PARAMETERS As String = "logillegalparameters"

Public Const MACRO_SETTING_REMOTETIMESYNCSERVER As String = "remotetimesyncserver"
'offset in hours of remote time server timezone to GMT
Public Const MACRO_SETTING_REMOTETIMESYNCOFFSET As String = "remotetimesyncserveroffset"

'TA 18/11/2004: flag to show OC Ids isssue 2448
'whether use OC id in discrepancies
Public Const MACRO_SETTING_USE_OC_ID As String = "useocid"

' NCJ 19 Jan 05 - Enabling of list filtering in drop-downs (Issue 2502)
' Must be set to "true" to enable functionality
Public Const MACRO_SETTING_LIST_FILTER = "listfiltering"

' DPH 02/02/2005 (PDU2300) Added pdu.exe location setting & "use pdu" setting
Public Const MACRO_SETTING_PDUEXE_LOCATION = "pduexelocation"
Public Const MACRO_SETTING_USE_PDU = "usepdu"

' TA
Public Const MACRO_SETTING_CDV_BIGINT = "cdvbigint"

' MLM 29/06/05:
Public Const MACRO_SETTING_SHOW_SDV_SCHEDULE_MENU = "showsdvschedulemenu"

'TA
Public Const MACRO_SETTING_DATATRANSFER_TIMEOUT = "datatransfertimeout"

' NCJ 7 Dec 05 - Switch on partial dates?
Public Const MACRO_SETTING_PARTIAL_DATES = "partialdates"
' NCJ 12 Jul 06 - Switch on Multi User SD?
Public Const MACRO_SETTING_MUSD = "musd"

Public Enum eMACRO_PC_Setting
    mpcUserSettingsFile = 0
    mpcAuthorisedUser = 1
    mpcOrganisation = 2
End Enum

'----------------------------------------------------------------------------------------
Public Function GetMACROPCSetting(ByVal enMACRO_PC_Setting As eMACRO_PC_Setting, ByVal sDefault As String, _
                                    Optional bUsingWWWDLL As Boolean = False) As String
'----------------------------------------------------------------------------------------
'Returns the item for a key if it exists or returns the default if passed
'----------------------------------------------------------------------------------------
Dim sValue As String
Dim oSettings As IMEDSettings
Dim sAppPath As String

    On Error GoTo ErrHandler
    
    'TA 20/11/2002: web needs to look up one folder higher
    If bUsingWWWDLL Then
        sAppPath = App.Path & "\..\"
    Else
        sAppPath = App.Path & "\"
    End If
    
    Set oSettings = New IMEDSettings
    oSettings.Init sAppPath & msFIXEDFILENAME
    
    GetMACROPCSetting = sDefault
   

    Select Case enMACRO_PC_Setting
    Case eMACRO_PC_Setting.mpcOrganisation
        GetMACROPCSetting = oSettings.GetKeyValue("organisation", sDefault)
    Case eMACRO_PC_Setting.mpcAuthorisedUser
        GetMACROPCSetting = oSettings.GetKeyValue("authoriseduser", sDefault)
    Case eMACRO_PC_Setting.mpcUserSettingsFile
        GetMACROPCSetting = oSettings.GetKeyValue("usersettingsfile", sDefault)
    End Select
    
    Set oSettings = Nothing
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSettings.GetPCMACROSetting"
End Function


'----------------------------------------------------------------------------------------
Public Function GetMACROSetting(ByVal sKey As String, ByVal sDefault As String) As String
'----------------------------------------------------------------------------------------
'Returns the item for a key if it exists or returns the default if passed
'----------------------------------------------------------------------------------------
Dim sValue As String

    On Error GoTo ErrHandler
    
    GetMACROSetting = sDefault
    
    'Exit if nothing passed
    If sKey = "" Then Exit Function
    
    'Exit if file does not exist
    If msFileName = "" Then Exit Function
    

    GetMACROSetting = moSetting.GetKeyValue(sKey, sDefault)

    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSettings.GetMACROSetting"
End Function

'----------------------------------------------------------------------------------------
Public Function SetMACROSetting(ByVal sKey As String, ByVal sValue As String) As Boolean
'----------------------------------------------------------------------------------------
'Adds/Updates/Deletes registry keys
'----------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    SetMACROSetting = False
    
    'Exit if nothing passed
    If sKey = "" Then Exit Function
    
    'Exit if file does not exist
    If msFileName = "" Then Exit Function
    
    SetMACROSetting = moSetting.SetKeyValue(sKey, sValue)

    
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSettings.SetMACROSetting"
End Function

'----------------------------------------------------------------------------------------
Public Function InitialiseSettingsFile(Optional bUsingWWWDLL As Boolean = False) As String
'----------------------------------------------------------------------------------------
'creates and initialises IMEDSettings DLL
'----------------------------------------------------------------------------------------
' REVISIONS
' DPH 22/03/2004 - tidy up of objects
'----------------------------------------------------------------------------------------
Dim sSettingsFileName As String
Dim objFSO As New FileSystemObject
Dim objFile As File
Dim sAppPath As String

    On Error GoTo ErrHandler
    
    InitialiseSettingsFile = ""
    
    'TA 20/11/2002: web needs to look up one folder higher
    If bUsingWWWDLL Then
        sAppPath = App.Path & "\..\"
    Else
        sAppPath = App.Path & "\"
    End If
    
    Set moSetting = New IMEDSettings
    'check if file exists then get the file path from MACROSettings30.txt
    If FileExists(sAppPath & msFIXEDFILENAME) Then
        sSettingsFileName = GetMACROPCSetting(mpcUserSettingsFile, "", bUsingWWWDLL)
    Else
        InitialiseSettingsFile = "Your MACROSettings30.txt file is missing"
    End If

    'only continue if file exists
    If sSettingsFileName <> "" Then
        'assume filename if there is no \
        If InStr(1, sSettingsFileName, "\") = 0 Then
            msFileName = sAppPath & sSettingsFileName
        Else
            msFileName = sSettingsFileName
        End If
        
        If FileExists(msFileName) Then
            Set objFile = objFSO.GetFile(msFileName)
            'Check if it is a readonly file
            If Not objFile.Attributes And ReadOnly Then
                Call moSetting.Init(msFileName)
            Else
                msFileName = ""
            End If
        Else
            'TA 10/12/2002: if it doesn't exist - create it
            On Error GoTo Errhandler2
            StringToFile msFileName, ""
            On Error GoTo ErrHandler
            Call moSetting.Init(msFileName)
'            msFileName = ""
        End If
    End If

    Set objFile = Nothing
    Set objFSO = Nothing
    
Exit Function
ErrHandler:
    InitialiseSettingsFile = "Please check your settings file is valid" & vbCrLf & "Error: " & Err.Number & " - " & Err.Description & "|modSettings.InitialiseSettingsFile"
    
    Exit Function
    
Errhandler2:
    InitialiseSettingsFile = "Cannot create file " & msFileName
    msFileName = ""
End Function

'---------------------------------------------------------------------
Function FileExists(ByVal strPathName As String) As Integer
'---------------------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'---------------------------------------------------------------------
Dim intFileNum As Integer

    On Error Resume Next

    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err, False, True)

    Close intFileNum

    Err = 0

End Function

'----------------------------------------------------------------------------------------'
Public Sub StringToFile(sFileName As String, sText As String)
'----------------------------------------------------------------------------------------'
' Write string to given file
'----------------------------------------------------------------------------------------'
Dim n As Integer

    n = FreeFile
    Open sFileName For Output As n
    
    Print #n, sText
    
    Close n

End Sub
