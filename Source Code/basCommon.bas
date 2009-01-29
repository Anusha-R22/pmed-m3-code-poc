Attribute VB_Name = "basCommon"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998-2006. All Rights Reserved
'   File:       basCommon.bas
'   Author:     Andrew Newbigging, June 1997
'   Purpose:    Assorted Windows functions used throughout MACRO
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'  1-17    Andrew Newbigging, Mo Morris      4/06/97 - 3/3/99
'   18      Paul Norris             28/07/99
'           Added the ConvertStringToCollection(), LockWindow(), UnLockWindow(), TextChange(),
'           TextLostFocus() routines
'   19      Paul Norris             05/08/99
'           Added routines TrimNull(),AfterStr(),BeforeStr(),RemovePreceedingChars()
'           Mo Morris               9/8/99
'           Removed Type PROTOCOL
'                   Declarions gProtocol, gcProtocols
'                   Function ReadProtocols
'   20      Paul Norris             18/08/99    Added ValidateDateFormat(), DisplayCRFPageDefinition()
'           AttemptDeleteCRFPage(),SelectNextCell(),PromptDeleteGridRow(),AttemptDeleteCRFPage()
'           and LoadCombo()
'   21      Paul Norris             20/08/99    Added DoesNameExistInRecordset()
'           GetUniqueName() and GenerateNextUniqueExportName()
'   22      Paul Norris             30/08/99    Amended ValidateDateFormat(),CountStrsInStr()
'           IsFormDisplayOrderAlphabetic()
'   23      PN  16/09/99    Moved string handling routines AfterStr(), BeforeStr(),
'                           ConvertStringToCollection(), CountStrsInStr(), Extension(), TrimNull(),
'                           RemovePreceedingChars(), ReplaceCharacters() and StripFileNameFromPath()
'                           to modStringUtilities.bas
'   24      PN  16/09/99    Moved StudyDefinition specific code to modStudyDefinition module
'                           AttemptDeleteCRFPage() and DisplayCRFPageDefinition()
'   PN      20/09/99        Added RestartSystemIdleTimer()
'   PN      24/09/99        Removed routines that are not called:
'                           RemoveTitleBar()
'   PN      26/09/99        Added FormatToValidDate() for mtm1.6 changes
'   NCJ 30/9/99 Commented out debug timer messages
'   WillC   11/10/99        Added Global variables for the logging of important actions ie
'                           the creation of a new User or the editing of a role etc etc.
'   Mo Morris   19/10/99    API declarations as required by frmAbout.GetFromRegistry and
'                           frmAbout.QueryValueEx added (From Knowledge base article Q145679)
'   Mo Morris   29/10/99    DAO to ADO conversion
'   NCJ     10 Nov 99       Added error handlers
'   Mo Morris   1/12/99     Changes made to gLog
'   WillC       9/12/99     Added the functions to find out what OS we are running on.
'   WillC    Changed the Following where present from Integer to Long  ClinicalTrialId
'           CRFPageId,VisitId,CRFElementID
'   ATN         13/12/99    Moved RemoveNull function into this module
'   NCJ 22 Dec 99   SR 2045 Wrote separate ValidateDateOnlyFormat and ValidateTimeOnlyFormat functions
'   NCJ 12 Jan 00   PromptDeleteGridRow moved to frmDataDefinition
'   ATN 14/1/2000   GetFromRegistry moved here from frmAbout
'   NCJ 18/1/00     Changed some "End" statements to "MACROEnd"
'   WillC 25/1/2000  Added function IsPrinterInstalled to see if a printer has been
'                   installed on the machine.
'   NCJ 4 Feb 00 SR2851 - gLog now uses "standard" SQL datestamp
'   NCJ 10 Feb 00 - Removed FormatToValidDate
'   NCJ 26/9/00 - Added in new NR/CTC fields to SQLToUpdateDataItemResponseHistory
'   NCJ 30/10/00 - Moved GetSQLStringLike and GetSQLStringEquals here from TrialData
'   NCJ 20/11/00 - Made arguments to GetSQLStringLike and GetSQLStringEquals ByVal
'   NCJ 21/11/00 - Added in new LaboratoryCode to SQLToUpdateDataItemResponseHistory
'   DPH 17/10/2001 Added FolderExistence routine to create missing folders return if file exists
'   TA 18/1/03  New encryption/decryption function for DCBB2.2.7.7
'   DPH 08/04/2002 - HexDecodeFile Turned into function & added Check to see if HexDecode Produces any output
'                   ExecCmd Passes back return value from GetProcess call
'                   GetFileLength function added
'   DPH 10/04/2002 - Added GetSiteSubjectLabFromFileName function & Added securehtmlfolder into InitialisationSettings
'   MLM 22/01/2002: Moved hex en/decoding functions to libXCeed.
'   DPH 22/01/2002 - Added RepeatNumber in SQLToUpdateDataItemResponseHistory
'   REM 11/07/02 - Added HadValue in SQLToUpdateDataItemResponseHistory
'   DPH 22/08/2002 - GetMacroDBSetting, SetMacroDBSetting added
'   REM 30/08/02 - Changed GetMacroDBSetting and SetMacroDBSetting Error handling to use new type error handler
'   RS 25/10/2002 - Added new timezone columns to routine SQLToUpdateDataItemResponseHistory
'   REM 31/10/02 - added timezone offset and Location to gLog
'   ic 05/11/2002   moved FileExists() to libLibrary
'   ASH 13/1/20023  Added optional parameter in sub GetMacroDBSetting
'   NCJ 16 Jun 03 - Tidied up GenerateNextUniqueExportName
'   REM/NCJ 1 Sept 03 - Added InsertMessage (for patient data transfer logging)
'   REM 0 Sept 03 - added DisplayGMTTime routine
'   DPH 02/02/2005 - Added new functions ExecCmdNoWait, ExecCmdExitCode for requirement PDU2300
'   NCJ 21 Jun 06 - Issue 2745 - Added new LogDetail constants for opening/closing trials
'------------------------------------------------------------------------------------'
Option Explicit
Option Compare Text

Public Enum EDateFormatResult
    EOK
    EInvalidFormat
    ETwoCharYear
End Enum


'REM 03/11/02 - Public constants for all LogDetail TaskIds
'NB: Remeber to add any new constants to frmLogDetails routine PopulateDropdowns
Public Const gsAUTOIMPORT = "AutoImport"
Public Const gsAUTO_IMPORT_LLD = "AutoImportLDD"
Public Const gsDEL_TRIAL_PRD = "DeleteTrialPRD"
Public Const gsDEL_TRIAL_SD = "DeleteTrialSD"
Public Const gsIMPORT_SDD = "ImportSDD"
Public Const gsEXPORT_SDD = "ExportSDD"
Public Const gsIMPORT_PRD = "ImportPRD"
Public Const gsEXPORT_PRD = "ExportPRD"
Public Const gsEXPORT_PAT_CAB = "ExportPatCAB"
Public Const gsIMPORT_UPGRADE = "ImportUpgrade"
Public Const gsEXPORT_STUDY_CAB = "ExportStudyCAB"
Public Const gsIMPORT_STUDY_CAB = "ImportStudyCAB"
Public Const gsIMPORT_PAT_CAB = "ImportPatCAB"
Public Const gsAUTO_EXPORT_PRD = "AutoExportPRD"
Public Const gsAUTO_IMPORT_PRD = "AutoImportPRD"
Public Const gsCLEAR_CABEXTR_FOLDER = "ClearCabExtractFolder"
Public Const gsIMPORT_DOC = "ImportDoc"
Public Const gsIMPORT_DOC_AND_GRAPHICS = "ImportDocAndGraphics"
Public Const gsEXPORT_LDD = "ExportLDD"
Public Const gsEXPORT_LDD_CAB = "ExportLDDCAB"
Public Const gsIMPORT_LDD = "ImportLDD"
Public Const gsCLEANUP_PRD = "CleanUpPRD"
Public Const gsIMPORT_PAT_ZIP = "ImportPatZIP"
Public Const gsIMPORT_STUDY_ZIP = "ImportStudyZIP"
Public Const gsIMPORT_LDD_ZIP = "ImportLDDZIP"
Public Const gsEXPORT_PAT_ZIP = "ExportPatZIP"
Public Const gsEXPORT_STUDY_ZIP = "ExportStudyZIP"
Public Const gsEXPORT_LDD_ZIP = "ExportLDDZIP"
Public Const gsHEX_ENCODE = "HexEncode"
Public Const gsHEX_DECODE = "HexDecode"
Public Const gsVALIDATE_ZIP = "ValidateZIP"
Public Const gsDOWNLOAD_MESG = "DownloadMessages"
Public Const gsDOWNLOAD_MIMESG = "DownloadMIMessages"
Public Const gsDATA_INTEG_COMMS = "DataIntegrityComms"
Public Const gsDATA_INTEG = "DataIntegrity"
Public Const gsSYS_TIMEOUT = "SystemTimeout"
Public Const gsPATDATA_SEND = "PatientDataSend"
Public Const gsSYSMSG_SEND_ERR = "SysMessageSendError"
Public Const gsSYSMSG_DOWNLOAD_ERR = "SysMesssageDownloadError"
Public Const gsREPORT_XFER = "ReportXfer"
Public Const gsREPORT_XFER_SITE = "ReportXferSite"
Public Const gsREPORT_XFER_SERVER = "ReportXferServer"
Public Const gsREPORT_XFER_ERR = "ReportXferErr"
Public Const gsDOWNLOAD_LFMESG = "DownloadLFMessages"
Public Const gsCANCEL_TRANSFER = "TransferCancelled"
Public Const gsCONNECT_FAIL = "ConnectToServerFailed"
Public Const gsPDUMSG_DOWNLOAD_ERR = "PduMesssageDownloadError"
' NCJ 21 Jun 06 - Added new constants
Public Const gsOPEN_TRIAL_SD = "OpenTrialSD"
Public Const gsNEW_TRIAL_SD = "NewTrialSD"
Public Const gsCLOSE_TRIAL_SD = "CloseTrialSD"
Public Const gsCOPY_TRIAL_SD = "CopyTrialSD"


'REM 03/11/02 - Public constants for User Log TaskId's
'NB: Remember to add any new constants to frmLogDetails routine PopulateDropdowns
Public Const gsCREATE_DB = "CreateDatabase"
Public Const gsNEW_ROLE = "NewRole"
Public Const gsEDIT_ROLE = "EditRole"
Public Const gsCHANGE_PSWD = "ChangePassword"
Public Const gsNEW_USER_ROLE = "NewUserRole"
Public Const gsDEL_USER_ROLE = "DeleteUserRole"
Public Const gsUSER_ENABLED = "UserEnabled"
Public Const gsUSER_DISABLED = "UserDisabled"
Public Const gsUSER_UNLOCKED = "UserUnLocked"
Public Const gsCHANGE_USERNAME_FULL = "ChangeUserNameFull"
Public Const gsCREATE_NEW_USER = "CreateNewUser"
Public Const gsLOGIN = "Login"
Public Const gsLOGOFF = "LogOff"
Public Const gsCHANGE_SYSADMIN_STATUS = "ChangeUserSysAdminStatus"
Public Const gsUSERNAME_CONFLICT = "UserNameConflict"

'
'Global Constants
'
Global Const gstrSEP_DIR$ = "\"                         'Directory separator character
Global Const gstrSEP_DIRALT$ = "/"                      'Alternate directory separator character
Global Const gstrSEP_EXT$ = "."                         'Filename extension separator character
Global Const gstrDECIMAL$ = "."

Global Const gintNOVERINFO% = 32767                     'flag indicating no version info
'
'Type Definitions
'
Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    nReserved1 As Integer
    nReserved2 As Integer
    szPathName As String * 256
End Type

Type VERINFO                                            'Version FIXEDFILEINFO
    strPad1 As Long                                     'Pad out struct version
    strPad2 As Long                                     'Pad out struct signature
    nMSLo As Integer                                    'Low word of ver # MS DWord
    nMSHi As Integer                                    'High word of ver # MS DWord
    nLSLo As Integer                                    'Low word of ver # LS DWord
    nLSHi As Integer                                    'High word of ver # LS DWord
    strPad3(1 To 36) As Byte                            'Pad out rest of VERINFO struct (36 bytes)
End Type

Public Type SelectedGridCell
    Row As Integer
    Col As Integer
End Type

'TA 06/12/2000: min max values for vb datatypes
Public Const INTEGER_MIN As Integer = -32768
Public Const INTEGER_MAX As Integer = 32767
Public Const LONG_MIN As Long = -2147483648#
Public Const LONG_MAX As Long = 2147483647

'
'Global Variables
'
Global LF$                                              'single line break
Global LS$                                              'double line break



Public sCreateNewUser As String                          'global variables for logging
Public sCreateNewRole As String
Public sAssignRoleToUser As String
Public sEditUserRole  As String
Public sAssignTrialToUser As String
Public sAssignSiteToUser As String
Public sAssignASiteATrialToUser As String
Public sChangePassword As String
Public sChangePasswordSettings As String
Public sCreateNewDatabase As String
Public sPasswordDisabled As String

'
'API/DLL Declarations for 32 bit SetupToolkit
''unused ones goes here
'Declare Function DiskSpaceFree Lib "STKIT432.DLL" Alias "DISKSPACEFREE" () As Long
'
'Declare Function AllocUnit Lib "STKIT432.DLL" () As Long
'Declare Function GetWinPlatform Lib "STKIT432.DLL" () As Long
'Declare Function fNTWithShell Lib "STKIT432.DLL" () As Boolean
'Declare Function FSyncShell Lib "STKIT432.DLL" Alias "SyncShell" (ByVal strCmdLine As String, ByVal intCmdShow As Long) As Long
'Declare Function DLLSelfRegister Lib "STKIT432.DLL" (ByVal lpDllName As String) As Integer
'Declare Sub lmemcpy Lib "STKIT432.DLL" (strDest As Any, ByVal strSrc As Any, ByVal lBytes As Long)
'Declare Function OSfCreateShellGroup Lib "STKIT432.DLL" Alias "fCreateShellFolder" (ByVal lpstrDirName As String) As Long
'Declare Function OSfCreateShellLink Lib "STKIT432.DLL" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
'Declare Function OSfRemoveShellLink Lib "STKIT432.DLL" Alias "fRemoveShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String) As Long
'Private Declare Function OSGetLongPathName Lib "STKIT432.DLL" Alias "GetLongPathName" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'
'
'Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFileName As String) As Long
'Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long
'Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Declare Function GetDriveType32 Lib "kernel32" Alias "GetDriveTypeA" (ByVal strWhichDrive As String) As Long
'Declare Function GetTempFileName32 Lib "kernel32" Alias "GetTempFileNameA" (ByVal strWhichDrive As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Long
'
'Declare Function VerInstallFile Lib "VERSION.DLL" Alias "VerInstallFileA" (ByVal Flags&, ByVal SrcName$, ByVal DestName$, ByVal SrcDir$, ByVal DestDir$, ByVal CurrDir As Any, ByVal tmpName$, lpTmpFileLen&) As Long
'Declare Function GetFileVersionInfoSize Lib "VERSION.DLL" Alias "GetFileVersionInfoSizeA" (ByVal strFileName As String, lVerHandle As Long) As Long
'Declare Function GetFileVersionInfo Lib "VERSION.DLL" Alias "GetFileVersionInfoA" (ByVal strFileName As String, ByVal lVerHandle As Long, ByVal lcbSize As Long, lpvData As Byte) As Long
'Declare Function VerQueryValue Lib "VERSION.DLL" Alias "VerQueryValueA" (lpvVerData As Byte, ByVal lpszSubBlock As String, lplpBuf As Long, lpcb As Long) As Long
'
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Long) As Long
'Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long



'used here
Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function OSGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function SetTime Lib "STKIT432.DLL" (ByVal strFileGetTime As String, ByVal strFileSetTime As String) As Integer
Const GWL_STYLE = (-16)
Const WS_DLGFRAME = &H400000
Const WS_SYSMENU = &H80000
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000

   Private Type STARTUPINFO
      cb As Long
      lpReserved As String
      lpDesktop As String
      lpTitle As String
      dwX As Long
      dwY As Long
      dwXSize As Long
      dwYSize As Long
      dwXCountChars As Long
      dwYCountChars As Long
      dwFillAttribute As Long
      dwFlags As Long
      wShowWindow As Integer
      cbReserved2 As Integer
      lpReserved2 As Long
      hStdInput As Long
      hStdOutput As Long
      hStdError As Long
   End Type

   Private Type PROCESS_INFORMATION
      hProcess As Long
      hThread As Long
      dwProcessID As Long
      dwThreadID As Long
   End Type

   Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
      hHandle As Long, ByVal dwMilliseconds As Long) As Long

   Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
      lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Long

   Private Declare Function CloseHandle Lib "kernel32" (ByVal _
      hObject As Long) As Long

   Private Const NORMAL_PRIORITY_CLASS = &H20&
    Private Const INFINITE = -1&
   
   ' DPH 17/10/2001 - API call to get focus
    Public Declare Function GetFocus Lib "user32" () As Long

'---------------------------------------------------------------------
'Following lines added by Mo Morris 19/10/99
'From Knowledge base article Q145679
'as required by frmAbout.GetFromRegistry and frmAbout.QueryValueEx
'---------------------------------------------------------------------
Public Const REG_SZ As Long = 1
Public Const REG_DWORD As Long = 4
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_ARENA_TRASHED = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259
Public Const KEY_ALL_ACCESS = &H3F
Public Const REG_OPTION_NON_VOLATILE = 0
Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
"RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
"RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, lpData As _
Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
ByVal cbData As Long) As Long
'---------------------------------------------------------------------
'End of lines added by Mo Morris 19/10/99
'---------------------------------------------------------------------

'SDM 04/01/00 SR1950    Find window last up

Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3


'SDM 05/01/00 SR918
Public Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Public Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long

Public Type SCROLLINFO
    cbSize As Long     ' Size of structure
    fMask As Long     ' Which value(s) you are changing
    nMin As Long     ' Minimum value of the scroll bar
    nMax As Long     ' Maximum value of the scroll bar
    nPage As Long     ' Large-change amount
    nPos As Long     ' Current value
    nTrackPos As Long     ' Current scroll position
End Type

' SCROLLINFO fMask constants:
Public Const SIF_RANGE = &H1
Public Const SIF_PAGE = &H2
Public Const SIF_POS = &H4
Public Const SIF_DISABLENOSCROLL = &H8
Public Const SIF_TRACKPOS = &H10
Public Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Public Const SB_VERT = 1
Public Const SB_THUMBPOSITION = 4

Public Const WM_VSCROLL = &H115

'REM 21/10/02 - constants for All studies and sites
Public Const ALL_STUDIES = "AllStudies"
Public Const ALL_SITES = "AllSites"

'---------------------------------------------------------------------
'Start of lines added by WillC 9/12/99
'---------------------------------------------------------------------

Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2
'windows-defined type OSVERSIONINFO

Public Type OSVERSIONINFO
  OSVSize         As Long         'size, in bytes, of this data structure
  dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
  dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
  dwBuildNumber   As Long         'NT: build number of the OS
                                  'Win9x: build number of the OS in low-order word.
                                  '       High-order word contains major & minor ver nos.
  PlatformID      As Long         'Identifies the operating system platform.
  szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3"
                                  'Win9x: 'arbitrary additional information'
End Type
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
  (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal _
    hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
    
' DPH 02/02/2005 - Added new API call GetExitCodeProcess used by ExecCmdExitCode
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Sub SetTopmostWindow(ByVal hWnd As Long, Optional topmost As Boolean = True)
    Const HWND_NOTOPMOST = -2
    Const HWND_TOPMOST = -1
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    SetWindowPos hWnd, IIf(topmost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, _
        SWP_NOMOVE + SWP_NOSIZE
End Sub



'SDM Q189170
Public Function MakeDWord(LoWord As Integer, HiWord As Integer) As Long
   MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function
Public Function LoWord(DWord As Long) As Integer
   If DWord And &H8000& Then ' &H8000& = &H00008000
      LoWord = DWord Or &HFFFF0000
   Else
      LoWord = DWord And &HFFFF&
   End If
End Function
Public Function HiWord(DWord As Long) As Integer
   HiWord = (DWord And &HFFFF0000) \ &H10000
End Function
  
  
  
'---------------------------------------------------------------------
Public Function IsWin95() As Boolean
'---------------------------------------------------------------------
'returns True if running Win95
'---------------------------------------------------------------------
   #If Win32 Then
  
      Dim OSV As OSVERSIONINFO
   
      OSV.OSVSize = Len(OSV)
   
      If GetVersionEx(OSV) = 1 Then
   
        'PlatformId contains a value representing the OS.
        'If VER_PLATFORM_WIN32_WINDOWS and
        'dwVerMajor = 4, and dwVerMinor = 0,
        'return true
         IsWin95 = (OSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                   (OSV.dwVerMajor = 4 And OSV.dwVerMinor = 0)
                            
      End If

   #End If

End Function

'---------------------------------------------------------------------
Public Function IsWin98() As Boolean
'---------------------------------------------------------------------
'returns True if running Win98
'---------------------------------------------------------------------
   #If Win32 Then
  
      Dim OSV As OSVERSIONINFO
   
      OSV.OSVSize = Len(OSV)
   
      If GetVersionEx(OSV) = 1 Then
   
        'PlatformId contains a value representing the OS.
        'If VER_PLATFORM_WIN32_WINDOWS and
        'dwVerMajor => 4, or dwVerMajor = 4 and
        'dwVerMinor > 0, return true
         IsWin98 = (OSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                   (OSV.dwVerMajor > 4) Or _
                   (OSV.dwVerMajor = 4 And OSV.dwVerMinor > 0)
                            
      End If

   #End If

End Function

'---------------------------------------------------------------------
Public Function IsWinNT4() As Boolean
'---------------------------------------------------------------------
'returns True if running WinNT4
'---------------------------------------------------------------------
   #If Win32 Then
  
      Dim OSV As OSVERSIONINFO
   
      OSV.OSVSize = Len(OSV)
   
      If GetVersionEx(OSV) = 1 Then
   
        'PlatformId contains a value representing the OS.
        'If VER_PLATFORM_WIN32_NT and dwVerMajor is 4, return true
         IsWinNT4 = (OSV.PlatformID = VER_PLATFORM_WIN32_NT) And _
                    (OSV.dwVerMajor = 4)
      End If

   #End If

End Function

'---------------------------------------------------------------------
Public Function IsWinNT5() As Boolean
'---------------------------------------------------------------------
'returns True if running WinNT2000 (NT5)
'---------------------------------------------------------------------
   #If Win32 Then
  
      Dim OSV As OSVERSIONINFO
   
      OSV.OSVSize = Len(OSV)
   
      If GetVersionEx(OSV) = 1 Then
   
        'PlatformId contains a value representing the OS.
        'If VER_PLATFORM_WIN32_NT and dwVerMajor is 5, return true
         IsWinNT5 = (OSV.PlatformID = VER_PLATFORM_WIN32_NT) And _
                    (OSV.dwVerMajor = 5)
      End If

   #End If

End Function

' NCJ 10 Feb 00 - Removed FormatToValidDate because no longer used
''---------------------------------------------------------------------
'Public Function FormatToValidDate(sDate As String) As String
''---------------------------------------------------------------------
'' convert format of the date passed in to enabel system, to cope with
'' different regional settings
''---------------------------------------------------------------------
'Dim sSafeDate As String
'
'    sSafeDate = ReplaceCharacters(Format$(sDate, "dd/mmm/yyyy"), ".", "/")
'    FormatToValidDate = sSafeDate
'
'End Function

'---------------------------------------------------------------------
Public Function ExecCmd(cmdline$) As Long
'---------------------------------------------------------------------
' REVISIONS
' DPH 08/04/2002 - ExecCmd Passes back return value from GetProcess call
'---------------------------------------------------------------------
Dim proc As PROCESS_INFORMATION
Dim Start As STARTUPINFO
Dim lRet As Long

    On Error GoTo ErrHandler
    
    ' Initialize the STARTUPINFO structure:
    Start.cb = Len(Start)
    ' Start the shelled application:
    lRet = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
       NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc)

    ' DPH 08/04/2002 - Set function return value
    ExecCmd = lRet

    ' Wait for the shelled application to finish:
    lRet = WaitForSingleObject(proc.hProcess, INFINITE)
    lRet = CloseHandle(proc.hProcess)
    
    Exit Function
       
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "ExecCmd", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'---------------------------------------------------------------------
Public Function ExecCmdNoWait(cmdline$) As Long
'---------------------------------------------------------------------
' REVISIONS
' DPH 25/01/2005 - ExecCmdNoWait launches process with no waiting
'---------------------------------------------------------------------
Dim proc As PROCESS_INFORMATION
Dim Start As STARTUPINFO
Dim lRet As Long

    On Error GoTo ErrHandler
    
    ' Initialize the STARTUPINFO structure:
    Start.cb = Len(Start)
    ' Start the shelled application:
    lRet = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
       NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc)

    ' Set function return value
    ExecCmdNoWait = lRet

    lRet = CloseHandle(proc.hProcess)
    
    Exit Function
       
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "ExecCmdNoWait", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'---------------------------------------------------------------------
Public Function ExecCmdExitCode(cmdline$) As Long
'---------------------------------------------------------------------
' REVISIONS
' DPH 02/02/2005 - ExecCmdExitCode Passes back exit code process value
'---------------------------------------------------------------------
Dim proc As PROCESS_INFORMATION
Dim Start As STARTUPINFO
Dim lRet As Long
Dim lExitCode As Long

    On Error GoTo ErrHandler
    
    ' Initialize the STARTUPINFO structure:
    Start.cb = Len(Start)
    ' Start the shelled application:
    lRet = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
       NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc)

    ' check exit code if launched properly
    If lRet <> 0 Then
        ' Wait for the shelled application to finish:
        lRet = WaitForSingleObject(proc.hProcess, INFINITE)
        ' get exit code
        lRet = GetExitCodeProcess(proc.hProcess, lExitCode)
        ' close handle of process
        lRet = CloseHandle(proc.hProcess)
    Else
        ' set exit code to -1
        lExitCode = -1
    End If
    
    ' Set function return value
    ExecCmdExitCode = lExitCode
        
    Exit Function
       
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "ExecCmdExitCode", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'---------------------------------------------------------------------
Sub AddDirSep(strPathName As String)
'---------------------------------------------------------------------
' SUB: AddDirSep
' Add a trailing directory path separator (back slash) to the
' end of a pathname unless one already exists
'
' IN/OUT: [strPathName] - path to add separator to
'---------------------------------------------------------------------

    If Right$(RTrim$(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
    
End Sub

'---------------------------------------------------------------------
Function gsAddDirSep(strPathName As String) As String
'---------------------------------------------------------------------
' FUNCTION: gsAddDirSep
' Add a trailing directory path separator (back slash) to the
' end of a pathname unless one already exists
'---------------------------------------------------------------------

    If Right$(RTrim$(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        gsAddDirSep = RTrim$(strPathName) & gstrSEP_DIR
    Else
        gsAddDirSep = RTrim$(strPathName)
    End If
    
End Function

'---------------------------------------------------------------------
Function ResolveResString(ByVal resID As Integer, ParamArray varReplacements() As Variant) As String
'---------------------------------------------------------------------
' FUNCTION: ResolveResString
' Reads resource and replaces given macros with given values
'
' Example, given a resource number 14:
'    "Could not read '|1' in drive |2"
'   The call
'     ResolveResString(14, "|1", "TXTFILE.TXT", "|2", "A:")
'   would return the string
'     "Could not read 'TXTFILE.TXT' in drive A:"
'
' IN: [resID] - resource identifier
'     [varReplacements] - pairs of macro/replacement value
'---------------------------------------------------------------------
Dim intMacro As Integer
Dim strResString As String
    
    strResString = LoadResString(resID)
    
    ' For each macro/value pair passed in...
    For intMacro = LBound(varReplacements) To UBound(varReplacements) Step 2
        Dim strMacro As String
        Dim strValue As String
        
        strMacro = varReplacements(intMacro)
        On Error GoTo MismatchedPairs
        strValue = varReplacements(intMacro + 1)
        On Error GoTo 0
        
        ' Replace all occurrences of strMacro with strValue
        Dim intPos As Integer
        Do
            intPos = InStr(strResString, strMacro)
            If intPos > 0 Then
                strResString = Left$(strResString, intPos - 1) & strValue & Right$(strResString, Len(strResString) - Len(strMacro) - intPos + 1)
            End If
        Loop Until intPos = 0
        
        
    Next intMacro
    
    ResolveResString = strResString
    
    Exit Function
    
MismatchedPairs:
    Resume Next
End Function

 
 '-----------------------------------------------------------
 ' FUNCTION GetShortPathName
 '
 ' Retrieve the short pathname version of a path possibly
 '   containing long subdirectory and/or file names
 '-----------------------------------------------------------
 '
 #If Win32 Then
 Function GetShortPathName(ByVal strLongPath As String) As String
     Const cchBuffer = 300
     Dim strShortPath As String * cchBuffer
     Dim lResult As Long

     On Error GoTo 0
     lResult = OSGetShortPathName(strLongPath, strShortPath, cchBuffer)
     If lResult = 0 Then
         Error 53 ' File not found
     Else
         GetShortPathName = StripTerminator(strShortPath)
     End If
 End Function
 #End If
 
'---------------------------------------------------------------------
Function StripTerminator(ByVal strString As String) As String
'---------------------------------------------------------------------
' FUNCTION: StripTerminator
'
' Returns a string without any zero terminator.  Typically,
' this was a string returned by a Windows API call.
'
' IN: [strString] - String to remove terminator from
'
' Returns: The value of the string passed in minus any
'          terminating zero.
'---------------------------------------------------------------------
Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

'---------------------------------------------------------------------
Function GetRemoteSupportFileVerStruct(ByVal strFileName As String, sVerInfo As VERINFO, Optional ByVal fIsRemoteServerSupportFile) As Boolean
'---------------------------------------------------------------------
' FUNCTION: GetRemoteSupportFileVerStruct
'
' Gets the file version information of a remote OLE server
' support file into a VERINFO TYPE variable (Enterprise
' Edition only).  Such files do not have a Windows version
' stamp, but they do have an internal version stamp that
' we can look for.
'
' IN: [strFileName] - name of file to get version info for
' OUT: [sVerInfo] - VERINFO Type to fill with version info
'
' Returns: True if version info found, False otherwise
'---------------------------------------------------------------------
Const strVersionKey = "Version="
Dim cchVersionKey As Integer
Dim iFile As Integer

    cchVersionKey = Len(strVersionKey)
    sVerInfo.nMSHi = gintNOVERINFO
    
    On Error GoTo Failed
    
    iFile = FreeFile

    Open strFileName For Input Access Read Lock Read Write As #iFile
    
    ' Loop through each line, looking for the key
    While (Not EOF(iFile))
        Dim strLine As String

        Line Input #iFile, strLine
        If Left$(strLine, cchVersionKey) = strVersionKey Then
            ' We've found the version key.  Copy everything after the equals sign
            Dim strVersion As String
            
            strVersion = Mid$(strLine, cchVersionKey + 1)
            
            'Parse and store the version information
            PackVerInfo strVersion, sVerInfo

            'Convert the format 1.2.3 from the .VBR into
            '1.2.0.3, which is really want we want
            sVerInfo.nLSLo = sVerInfo.nLSHi
            sVerInfo.nLSHi = 0
            
            GetRemoteSupportFileVerStruct = True
            Close iFile
            Exit Function
        End If
    Wend
    
    Close iFile
    Exit Function

Failed:
    GetRemoteSupportFileVerStruct = False
End Function

'---------------------------------------------------------------------
Sub PackVerInfo(ByVal strVersion As String, sVerInfo As VERINFO)
'---------------------------------------------------------------------
' SUB: PackVerInfo
'
' Parses a file version number string of the form
' x[.x[.x[.x]]] and assigns the extracted numbers to the
' appropriate elements of a VERINFO type variable.
' Examples of valid version strings are '3.11.0.102',
' '3.11', '3', etc.
'
' IN: [strVersion] - version number string
'
' OUT: [sVerInfo] - VERINFO type variable whose elements
'                   are assigned the appropriate numbers
'                   from the version number string
'---------------------------------------------------------------------
Dim intOffset As Integer
Dim intAnchor As Integer

    On Error GoTo PVIError

    intOffset = InStr(strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nMSHi = Val(strVersion)
        GoTo PVIMSLo
    Else
        sVerInfo.nMSHi = Val(Left$(strVersion, intOffset - 1))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor))
        GoTo PVILSHi
    Else
        sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor))
        GoTo PVILSLo
    Else
        sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nLSLo = Val(Mid$(strVersion, intAnchor))
    Else
        sVerInfo.nLSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
    End If

    Exit Sub

PVIError:
    sVerInfo.nMSHi = 0
PVIMSLo:
    sVerInfo.nMSLo = 0
PVILSHi:
    sVerInfo.nLSHi = 0
PVILSLo:
    sVerInfo.nLSLo = 0
End Sub

'---------------------------------------------------------------------
Public Sub gLog(ByVal sTaskId As String, ByVal sMessage As String, _
                Optional sConMACRO As ADODB.Connection = Nothing)
'---------------------------------------------------------------------
' WillC 11/10/99 Changed the variable prefixes to Infermed standards
'Changed by Mo Morris 17/11/99
'call to Get TimeStamp removed because for Access databases it was returning a
'Now date in format dd/mm/yyyy, which if the dd part was 12 or less would get
'switched to a mm/dd/yyyy format by the software that executes SQL statement
'REM 31/10/02 - added timezone offset and Location
'---------------------------------------------------------------------
Dim sSQL As String
Dim sSQLNow As String
Dim nLogNumber As Long
Dim rsLogDetails As ADODB.Recordset
Dim nTimeZone As Integer
Dim sLocation As String
Dim oTimeZone As TimeZone
'Ash 12/12/2002
Dim conNewMACRO As ADODB.Connection

    On Error GoTo ErrHandler
    'Ash 12/12/2002
    If Not sConMACRO Is Nothing Then
        Set conNewMACRO = sConMACRO
    Else
        Set conNewMACRO = MacroADODBConnection
    
    End If
    
    Set oTimeZone = New TimeZone
    
    ' NCJ 4 Feb 00 SR2851 - Use standard SQL datestamp
    sSQLNow = SQLStandardNow
    
    'Log messages have a combined key of LogDateTime and LogNumber. The first log
    'messages for a particular time will have a LogNumber of 0, the next 1 and so on
    'until the LogDateTime moves on a second.
    
    ' NCJ 4/2/00 - Changed dNow to sSQLNow
    sSQL = " SELECT LogNumber From LogDetails WHERE LogDateTime = " & sSQLNow

    'assess the number of records and set the LogNumber for this entry (nLogNumber)
    Set rsLogDetails = New ADODB.Recordset
    'Note use of adOpenKeyset cusor. Recordcount does not work with a adOpenDynamic cursor
    rsLogDetails.Open sSQL, conNewMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    nLogNumber = rsLogDetails.RecordCount
    rsLogDetails.Close
    Set rsLogDetails = Nothing
    
    'Changed Mo Morris 6/4/01, truncate large messages that might have ben created by an error
    'to 255 characters so that it fits into field LogMessage in table LogDetails
    If Len(sMessage) > 255 Then
        sMessage = Left(sMessage, 255)
    End If
    
    'changed Mo Morris 6/1/00
    'check log message for single quotes
    sMessage = ReplaceQuotes(sMessage)
    
    'REM 31/10/02 - added timezone offset and Location
    nTimeZone = oTimeZone.TimezoneOffset
    'Location will always be Local, will only be chnaged when transfered back to the server if a site
    sLocation = "Local"
    
    ' NCJ 4/2/00 - Changed dNow to sSQLNow
    'Mo Morris 30/8/01 Db Audit (UserId to UserName)
    'REM 31/10/02 - added TimeZone and Location
    sSQL = " INSERT INTO LogDetails " _
        & "(LogDateTime,LogNumber,TaskId,LogMessage,UserName,LogDateTime_TZ,Location,Status)" _
        & " Values (" & sSQLNow & "," & nLogNumber & ",'" & sTaskId _
        & "','" & sMessage & "','" & goUser.UserName & "'," & nTimeZone & ",'" & sLocation & "'," & 0 & ")"

    conNewMACRO.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "gLog", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Function ExtractFirstItemFromList( _
        ByRef rExtractFrom As String, _
        ByVal vSeparator As String) As String
'---------------------------------------------------------------------
' Extract first item from string rExtractFrom up to vSeparator
' Return what's left of rExtractFrom after separator
' NCJ 24/11/00 - Changed SeparatorPosition to Long; tidied up code
'---------------------------------------------------------------------
Dim lSeparatorPosition As Long

    On Error GoTo ErrHandler
    
    lSeparatorPosition = InStr(rExtractFrom, vSeparator)
    
    If lSeparatorPosition = 0 Then
        ' No separator found - return whole string
        
'        lSeparatorPosition = Len(rExtractFrom)
'        '  Extract item
'        ExtractFirstItemFromList = _
'            Left(rExtractFrom, lSeparatorPosition)

        ExtractFirstItemFromList = rExtractFrom
        rExtractFrom = ""
    Else
        '  Extract item
        ExtractFirstItemFromList = _
            Left(rExtractFrom, lSeparatorPosition - 1)
        ' Return what remains after separator
        rExtractFrom = Mid(rExtractFrom, lSeparatorPosition + Len(vSeparator))
    End If

    Exit Function
 
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "ExtractFirstItemFromList", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Function

'---------------------------------------------------------------------
Public Sub CreateBinaryFileFromString(ByVal vFilename As String, _
                                      ByVal vString As Variant)
'---------------------------------------------------------------------
Dim nFileNumber As Integer

    On Error GoTo ErrHandler

    nFileNumber = FreeFile
    
    Open vFilename For Output As nFileNumber
    
    Print #nFileNumber, vString
    
    Close nFileNumber

    Exit Sub
 
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "CreateBinaryFileFromString", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------------------
Public Function ReadStringFromBinaryFile(ByVal vFilename As String) As String
'---------------------------------------------------------------------
Dim nFileNumber As Integer

    On Error GoTo ErrHandler
    
    nFileNumber = FreeFile
    
    Open vFilename For Input As nFileNumber
    
    ReadStringFromBinaryFile = Input(FileLen(vFilename), #nFileNumber)
    
    Close nFileNumber

    Exit Function
 
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "ReadStringFromBinaryFile", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'---------------------------------------------------------------------
Public Function GetTimeStamp() As String
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    '   ATN 26/2/99
    '   Use databasetype property of the user class to determine which time format
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.sqlserver
        GetTimeStamp = Format(Now, "mm/dd/yyyy hh:mm:ss")
    Case MACRODatabaseType.Access
        GetTimeStamp = Format(Now, "dd/mm/yyyy hh:mm:ss")
    Case Else
        GetTimeStamp = Format(Now, "dd/mm/yyyy hh:mm:ss")
    End Select

    Exit Function
 
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "GetTimeStamp", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function


'---------------------------------------------------------------------
Public Sub SetReportDatabase(ByVal vReportControl As Control)
'---------------------------------------------------------------------
'   SPR 429 ATN 8/10/98
'   New subroutine which receives a Crystal Report control and sets
'   its connection to the database which the user has selected
'---------------------------------------------------------------------

Dim nNumberOfTables As Integer
Dim cnn1 As ADODB.Connection

    On Error GoTo ErrHandler
   

    '   Connect the report to the database to which the user has logged on
    '   Userid and password are left blank so that NT integrated security will be used
    vReportControl.Connect = "DSN=" & goUser.Database.NameOfDatabase & ";UID=;PWD=;DSQ=" & goUser.Database.NameOfDatabase
    
    
    '   Get the number of tables used in the report and then, for each table,
    '   strip out the tablename (the datafile is saved in the format server.owner.tablename
    '   so its the final segment that we need) and use this to set the correct table
    For nNumberOfTables = 0 To vReportControl.RetrieveDataFiles
        vReportControl.DataFiles(nNumberOfTables) = Mid(vReportControl.DataFiles(nNumberOfTables), InStr(vReportControl.DataFiles(nNumberOfTables), ".") + 1)
        vReportControl.DataFiles(nNumberOfTables) = Mid(vReportControl.DataFiles(nNumberOfTables), InStr(vReportControl.DataFiles(nNumberOfTables), ".") + 1)
    Next
    
    vReportControl.PrinterStartPage = 0
    vReportControl.PrinterStopPage = -1
    
    vReportControl.WindowLeft = 50
    vReportControl.WindowTop = 50
    vReportControl.WindowWidth = Screen.Width / 18
    vReportControl.WindowHeight = Screen.Height / 18

    Exit Sub
 
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "SetReportDatabase", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub LockWindow(oWindow As Object)
'---------------------------------------------------------------------
' PN 22/07/99 - routine added
' this is an api call to prevent updates to a window
' use it when populating grids, listviews or treeviews with large amounts of data
' don't forget to call the UnlockWindow() function when the control is populated,
' otherwise the window will never update
'---------------------------------------------------------------------
    
    Call LockWindowUpdate(oWindow.hWnd)

End Sub

'---------------------------------------------------------------------
Public Sub UnlockWindow()
'---------------------------------------------------------------------
' PN 22/07/99 - routine added
' this api call will allow updates to a previously locked window
'---------------------------------------------------------------------
    
    Call LockWindowUpdate(0&)

End Sub

'---------------------------------------------------------------------
Public Sub TextChange(ctl As Control, oObject As Object, sProp As String)
'---------------------------------------------------------------------
' PN 28/07/99 - routine added
' this function is used by forms in control change events
' it will assign the oObject's sProp property with the ctl.Text string. if an error
' occurs it will reset the ctl.Text to oObject's sProp property.
' it uses the CallByName function to keep the code generic
'---------------------------------------------------------------------
Dim lPos As Long
    
    On Error GoTo InputErr
    
    If Not oObject Is Nothing Then
        ' assign the property
        Call CallByName(oObject, sProp, VbLet, ctl.Text)
    End If
    
    Exit Sub
    
InputErr:
    ' put the original text back into the control
    ' since the input failed validation
    Beep
    lPos = ctl.SelStart
    ctl = CallByName(oObject, sProp, VbGet)
    If lPos > 0 Then lPos = lPos - 1
    ctl.SelStart = lPos

End Sub

'---------------------------------------------------------------------
Public Function TextLostFocus(ctl As Control, oObject As Object, sProp As String) As String
'---------------------------------------------------------------------
' PN 28/07/99 - routine added
' this function is used by forms in control lostfocus events
' this fx works on the same principle as TextChange. It handles reading oObjects's
' sProp property
' it uses the CallByName function to keep the code generic
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If Not oObject Is Nothing Then
        TextLostFocus = CallByName(oObject, sProp, VbGet)
    End If

    Exit Function
 
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "TextLostFocus", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Function

'---------------------------------------------------------------------
Public Sub LoadCombo(cbo As ComboBox, oList As clsTextList)
'---------------------------------------------------------------------
' PN 18/08/99 - routine added
' load a combo box from a clstextlist object'
' used by forms that have combos to be populated from a clsTextList object
'---------------------------------------------------------------------
Dim vItem As Variant

    On Error GoTo ErrHandler

    With cbo
        .Clear
        For Each vItem In oList
            .AddItem vItem
        Next vItem
        If .ListCount > 0 Then .ListIndex = 0
    End With

    Exit Sub
 
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "LoadCombo", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Function ValidateTimeOnlyFormat(sFormat As String) As EDateFormatResult
'---------------------------------------------------------------------
' Validate sFormat as a time (not date) format
' Based on PN's original ValidateDateFormat
' NCJ 22/12/99, SR 2045
' NCJ 3 Feb 00 - Not used any more
'   (see instead ValidateDateFormatString in ProformaEditor)
'---------------------------------------------------------------------

    ValidateTimeOnlyFormat = EInvalidFormat
    
End Function

'---------------------------------------------------------------------
Public Function ValidateDateOnlyFormat(sFormat As String) As EDateFormatResult
'---------------------------------------------------------------------
' Validate sFormat as a date (not time) format
' NCJ 22/12/99, SR 2045
' NCJ 3 Feb 00 - Not used any more
'   (see instead ValidateDateFormatString in ProformaEditor)
'---------------------------------------------------------------------
    
    ValidateDateOnlyFormat = EInvalidFormat

End Function

'---------------------------------------------------------------------
Public Function ValidateDateFormat(sFormat As String) As EDateFormatResult
'---------------------------------------------------------------------
' PN 18/08/99 - routine added
' will validate a format for dates
' if sFormat is a valid date format then returns True else False
' NCJ 3 Feb 00 - Not used any more
'   (see instead ValidateDateFormatString in ProformaEditor)
'---------------------------------------------------------------------

    ValidateDateFormat = EInvalidFormat

End Function

'---------------------------------------------------------------------
Public Sub SelectNextCell(oGrid As MSFlexGrid)
'---------------------------------------------------------------------
' PN 18/08/99 - routine added
' select the next cell in a MSFlexGrid grid
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    With oGrid
        If .Col < .Cols - 1 Then
            .Col = .Col + 1
            If .ColWidth(.Col) = 0 And .AllowUserResizing = flexResizeNone Then
                ' this is an invisible column so it must be skipped
                Call SelectNextCell(oGrid)
                
            End If
            
        ElseIf .Row < .Rows - 1 Then
            .Row = .Row + 1
            .Col = 1
        Else
            .Row = 1
            .Col = 1
        End If
    End With

    Exit Sub
 
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "SelectNextCell", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Public Function GenerateNextUniqueExportName(sDataItemCode As String, _
                            lClinicalTrialId As Long, _
                            nVersionId As Integer) As String
'---------------------------------------------------------------------
' this function will generate the next available unique export name for a data item.
' this has to be used because the export codes can be edited by the user and uniqueness
' is not guaranteed
' NCJ 16 Jun 03 - Tidied up routine and debugged
'---------------------------------------------------------------------
Dim sUniqueName As String       'unique export code
Dim rsExistingExportNames As ADODB.Recordset
Dim nCounter As Integer
Dim sSQL As String

    On Error GoTo ErrHandler

    ' Default to data item code
    sUniqueName = sDataItemCode
    
    ' first query for all other export names
    sSQL = "SELECT ExportName FROM DataItem WHERE ClinicalTrialID=" & lClinicalTrialId _
            & " AND VersionID=" & nVersionId
    Set rsExistingExportNames = New ADODB.Recordset
    rsExistingExportNames.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    With rsExistingExportNames
    
        If .RecordCount > 0 Then
            ' There are existing export codes
            nCounter = 0
            Do While DoesNameExistInRecordset(rsExistingExportNames, sUniqueName)
                ' That one already exists - try another
                nCounter = nCounter + 1
                sUniqueName = sDataItemCode & nCounter
            Loop
        End If
        
    End With
    
    rsExistingExportNames.Close
    Set rsExistingExportNames = Nothing
    
    GenerateNextUniqueExportName = sUniqueName

    Exit Function
 
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basCommon.GenerateNextUniqueExportName"
    
End Function

'---------------------------------------------------------------------
Public Function DoesNameExistInRecordset(oArray As ADODB.Recordset, sName As String) As Boolean
'---------------------------------------------------------------------
' PN 20/08/99 - routine added
' search for sName in a recordset
' return true if found and false if not found
'---------------------------------------------------------------------
Dim lIndex As Long  ' loop counter

    On Error GoTo ErrHandler

    With oArray
        lIndex = 1
        Do While Not .EOF
            If .Fields(0) = sName Then
                DoesNameExistInRecordset = True
                Exit Do
            End If
            .MoveNext
        Loop
    End With

    Exit Function
 
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|basCommon.DoesNameExistInRecordset"
    
End Function

'---------------------------------------------------------------------
Public Function IsFormDisplayOrderAlphabetic() As Boolean
'---------------------------------------------------------------------
' PN 30/08/99 - routine added
' determine if the form display order in the data list window is alphabetic or not
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    If GetSetting(App.Title, "Settings", "DataListFormOrderAlphabetic") = "-1" Then
        IsFormDisplayOrderAlphabetic = True
    Else
        IsFormDisplayOrderAlphabetic = False
    End If

    Exit Function
 
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "IsFormDisplayOrderAlphabetic", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'---------------------------------------------------------------------
Public Sub RestartSystemIdleTimer()
'---------------------------------------------------------------------
' PN 20/09/99 - routine added
' restart the timer because the user has interacted with the system
'
' the timer raises its timer event every minute
' it increments a counter glSystemIdleTimeoutCount to indicate how many minutes
' have passed since the last user interaction
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    #If WebRDE <> -1 Then
            
        ' here the counter must be set to 0 to effectively restart the timer
        glSystemIdleTimeoutCount = 0
        frmMenu.tmrSystemIdleTimeout.Enabled = False
            
        'enable timer if DevMode is zero
        #If DevMode = 0 Then
                frmMenu.tmrSystemIdleTimeout.Enabled = True
        #End If
    
    #End If
    
    ' Debug.Print "Now reset timer " & Format(Now, "hh:mm:ss")
 
Exit Sub

ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "RestartSystemIdleTimer", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
   
End Sub

'---------------------------------------------------------------------
Public Function RemoveNull(ByVal Text As Variant)
'---------------------------------------------------------------------
'   SDM 23/11/99
'   To take a string that may contain null and return it
'   NCJ 1 Dec 99 - RTrim the string before returning
'---------------------------------------------------------------------
    
    If IsNull(Text) Then
        RemoveNull = ""
    Else
        RemoveNull = RTrim(Text)
    End If

End Function

'---------------------------------------------------------------------
Public Function GetFromRegistry(sKeyName As String, sValueName As String) As String
'---------------------------------------------------------------------
'Used to extract information from the HKEY_LOCAL_MACHINE section of the registry
'as opposed to GetSetting that only extracts from HKey_CURRENT_USER
'From Knowledge base article Q145679
'---------------------------------------------------------------------
Dim lRetVal As Long      'result of the API functions
Dim hKey As Long         'handle of opened key
Dim vValue As Variant    'setting of queried value

    On Error GoTo ErrHandler
    
    lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    RegCloseKey (hKey)
    GetFromRegistry = vValue
    
    Exit Function
    
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "GetFromRegistry", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'---------------------------------------------------------------------
Public Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
'---------------------------------------------------------------------
    
Dim lRetVal As Long         'result of the SetValueEx function
Dim hKey As Long         'handle of open key

    On Error GoTo ErrHandler
       'open the specified key
'       lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, _
'                              KEY_ALL_ACCESS, hKey)
       lRetVal = RegCreateKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0&, _
                 vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                 0&, hKey, lRetVal)
       lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
       RegCloseKey (hKey)
       
       Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "SetKeyValue", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Public Function SetValueEx(ByVal hKey As Long, sValueName As String, _
   lType As Long, vValue As Variant) As Long
'---------------------------------------------------------------------
Dim lValue As Long
Dim sValue As String

       Select Case lType
       Case REG_SZ
               sValue = vValue & Chr$(0)
               SetValueEx = RegSetValueExString(hKey, sValueName, 0&, _
                                              lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        End Select
        
End Function

'---------------------------------------------------------------------
Private Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
'---------------------------------------------------------------------
Dim hNewKey As Long         'handle to the new key
Dim lRetVal As Long         'result of the RegCreateKeyEx function

    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
                 vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                 0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
    
End Sub

'---------------------------------------------------------------------
Public Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
   String, vValue As Variant) As Long
'---------------------------------------------------------------------
'called from GetFRomRegistry
'From Knowledge base article Q145679
'---------------------------------------------------------------------
Dim cch As Long
Dim lrc As Long
Dim lType As Long
Dim lValue As Long
Dim sValue As String

    On Error GoTo QueryValueExError
    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5
    
    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If
        ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
        'all other data types not supported
        lrc = -1
    End Select
    
QueryValueExExit:
    QueryValueEx = lrc
    Exit Function
       
QueryValueExError:
    Resume QueryValueExExit
    End Function

'---------------------------------------------------------------------
Public Function IsPrinterInstalled() As Boolean
'---------------------------------------------------------------------
' Check to see if a printer has been installed on the machine.
'---------------------------------------------------------------------

    Dim sdummy As String

    On Error Resume Next
    
    sdummy = Printer.DeviceName
    
    If Err.Number Then
        MsgBox "No default printer installed." & vbCrLf & _
            "To install and select a default printer, select the " & _
            "Setting / Printers command in the Start menu, and then " & _
            "double-click on the Add Printer icon.", vbExclamation, _
            "Printer Error"
        IsPrinterInstalled = False
    Else
        IsPrinterInstalled = True
    End If
    
End Function

'---------------------------------------------------------------------
Public Function GetSQLStringLike(ByVal sFieldName As String, _
                                    ByVal sValue As String, Optional bSecurityDB As Boolean = False) As String
'---------------------------------------------------------------------
' NCJ 20/10/00 - Returns the SQL which compares sFieldName with sValue
' using "Like". Assumes that sValue > "" (otherwise returns empty string)
' Returns correct SQL for current database (Oracle or Access/SQL)
' See also GetSQLStringEquals
' NCJ 20/11/00 - Made arguments ByVal
' REM 09/05/03 - Added optional prarmeter for security database
'---------------------------------------------------------------------
Dim sSQL As String
Dim nDBType As eMACRODatabaseType

    If bSecurityDB Then
        nDBType = goUser.SecurityDatabaseType
    Else
        nDBType = goUser.Database.DatabaseType
    End If
    
    sSQL = ""
    If sValue > "" Then
        If nDBType = MACRODatabaseType.Oracle80 Then
            sSQL = " (NLS_LOWER(" & sFieldName & ") LIKE '%" _
                            & ReplaceQuotes(LCase(sValue)) & "%') "
        Else
            ' Access or SQL Server
            sSQL = " " & sFieldName & " Like '%" & ReplaceQuotes(sValue) & "%' "
        End If
    End If
    
    GetSQLStringLike = sSQL
    
End Function

'-------------------------------------------------------------------
Public Function SQLToUpdateDataItemResponseHistory() As String
'-------------------------------------------------------------------
' NCJ 17/2/00 - This same SQL string was used in several places in this module
' so extracted it into a common function
' It copies a complete record from DataItemReponse into DataItemResponseHistory
' and only requires a WHERE clause to be added at the end
' NCJ 26/4/00 - Added in new validation/overrule fields
' NCJ 26/9/00 - Added in new NR/CTC fields
' NCJ 21/11/00 - Added in LaboratoryCode
' DPH 22/01/2002 - Added RepeatNumber
' REM 11/07/02 - Added HadValue
' RS 25/10/2002 - Added Timezone columns
'-------------------------------------------------------------------
Dim sSQL As String

    'SDM SR2611
    'Mo Morris 30/8/01 Db Audit (ReviewComment removed, UserId to UserName)
    sSQL = "INSERT INTO DataItemResponseHistory (" & _
           "ClinicalTrialId, " & _
           "TrialSite, " & _
           "PersonId, " & _
           "ResponseTaskId, " & _
           "ResponseTimestamp, " & _
           "VisitId, " & _
           "CRFPageId, " & _
           "CRFElementId, " & _
           "DataItemId, " & _
           "VisitCycleNumber, " & _
           "CRFPageCycleNumber, " & _
           "CRFPageTaskId, "
    sSQL = sSQL & _
           "ResponseValue, " & _
           "ResponseStatus, " & _
           "ValueCode, " & _
           "UserName, " & _
           "UnitOfMeasurement, " & _
           "Comments, " & _
           "Changed, " & _
           "SoftwareVersion, " & _
           "ReasonForChange, " & _
           "LockStatus, "
    ' NCJ 26/4/00
    sSQL = sSQL & _
            "ValidationId, ValidationMessage, OverruleReason, "
    ' NCJ 26/9/00
    ' NCJ 21/11/00 - Added LaboratoryCode
    ' DPH 22/01/2002 - Added RepeatNumber
    ' REM 11/07/02 - Added HadValue
    ' RS 25/10/2002 - Added ResponseTimestamp_TZ, ImportTimestamp_TZ, DatabaseTimestamp, DatabaseTimestamp_TZ
    sSQL = sSQL & _
            "LabResult, CTCGrade, ClinicalTestDate, LaboratoryCode, HadValue, RepeatNumber, "
    sSQL = sSQL & _
            "ResponseTimestamp_TZ, ImportTimestamp_TZ, DatabaseTimestamp, DatabaseTimestamp_TZ )"
    sSQL = sSQL & _
           "SELECT " & _
           "ClinicalTrialId, " & _
           "TrialSite, " & _
           "PersonId, " & _
           "ResponseTaskId, " & _
           "ResponseTimestamp, " & _
           "VisitId, " & _
           "CRFPageId, " & _
           "CRFElementId, " & _
           "DataItemId, " & _
           "VisitCycleNumber, " & _
           "CRFPageCycleNumber, " & _
           "CRFPageTaskId, "
    sSQL = sSQL & _
           "ResponseValue, " & _
           "ResponseStatus, " & _
           "ValueCode, " & _
           "UserName, " & _
           "UnitOfMeasurement, " & _
           "Comments, " & _
           "Changed, " & _
           "SoftwareVersion, " & _
           "ReasonForChange, " & _
           "LockStatus, "
    ' NCJ 26/4/00
    sSQL = sSQL & _
            "ValidationId, ValidationMessage, OverruleReason, "
    ' NCJ 26/9/00
    ' NCJ 21/11/00 - Added LaboratoryCode
    ' DPH 22/01/2002 - Added RepeatNumber
    'REM 11/07/02 - Added HadValue
    ' RS 25/10/2002 - Added ResponseTimestamp_TZ, ImportTimestamp_TZ, DatabaseTimestamp, DatabaseTimestamp_TZ
    sSQL = sSQL & _
            "LabResult, CTCGrade, ClinicalTestDate, LaboratoryCode, HadValue, RepeatNumber, "
    sSQL = sSQL & _
            "ResponseTimestamp_TZ, ImportTimestamp_TZ, DatabaseTimestamp, DatabaseTimestamp_TZ "
    sSQL = sSQL & _
           "FROM " & _
           "DataItemResponse "
    ' WHERE clause to be appended as required

    SQLToUpdateDataItemResponseHistory = sSQL

End Function


'---------------------------------------------------------------------
Public Function GetSQLStringEquals(ByVal sFieldName As String, _
                                    ByVal sValue As String) As String
'---------------------------------------------------------------------
' NCJ 27/10/00 - Returns the SQL which compares sFieldName with sValue using "=".
' Returns correct SQL for current database (Oracle or Access/SQL)
' See also GetSQLStringLike
' NCJ 20/11/00 - Made arguments ByVal
'---------------------------------------------------------------------
Dim sSQL As String

    If goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        sSQL = "NLS_LOWER(" & sFieldName & ") = '" & LCase(ReplaceQuotes(sValue)) & "' "
    Else
        sSQL = sFieldName & " = '" & ReplaceQuotes(sValue) & "' "
    End If
    
    GetSQLStringEquals = sSQL

End Function

'---------------------------------------------------------------------
Public Function TableExists(ByVal sTableName As String) As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error Resume Next
    
    sSQL = "SELECT * FROM " & sTableName
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    If Err.Number <> 0 Then
        TableExists = False
        Err.Clear
    Else
        TableExists = True
    End If

End Function

'---------------------------------------------------------------------------------
Public Function FolderExistence(sFilePath As String, Optional bCheckFile As Boolean) As Boolean
'---------------------------------------------------------------------------------
' DPH 17/10/2001 - To avoid crashes when trying to open a folder not in existence
' Work Way Through entered Folder list to make sure that all folders exist
' if they do not then create them as you go
' NB Assume have been given full FILE path - checks if file exists & uses scrrun.dll
'---------------------------------------------------------------------------------
Dim oFileSystem As New FileSystemObject
Dim oFileFolder As Folder
Dim oFile As Object
Dim nPos As Integer, nCount As Integer
Dim sFolderCheck As String, sFileCheck As String
    
    On Error GoTo ErrLabel

    FolderExistence = True
    
    If IsMissing(bCheckFile) Then
        bCheckFile = False
    End If
    
    ' Initialise
    nPos = 1
    sFolderCheck = ""
    sFileCheck = ""
    Set oFileFolder = Nothing
    
    ' Loop through \ to a maximum arbitary depth of 50
    For nCount = 0 To 50
        nPos = InStr(nPos, sFilePath, "\", vbBinaryCompare)
        If nPos > 0 Then
            sFolderCheck = Left(sFilePath, nPos - 1)
            ' Check Not empty foldername or mapped network drive
            If (sFolderCheck <> "") And (sFolderCheck <> "\") And Not (nCount = 2 And Left(sFolderCheck, 2) = "\\") Then
                If Not oFileSystem.FolderExists(sFolderCheck) Then
                    If oFileFolder.Attributes Mod 2 = 1 Then
                        ' Cannot create file so exit
                        FolderExistence = False
                        Exit Function
                    End If
                    ' Create Missing Folder
                    oFileSystem.CreateFolder sFolderCheck
                End If
                Set oFileFolder = oFileSystem.GetFolder(sFolderCheck)
            End If
            nPos = nPos + 1
        Else
            ' Exit loop as last "\" (folder) found
            If bCheckFile Then
                If Not oFileSystem.FileExists(sFilePath) Then
                    FolderExistence = False
                End If
            Else
                If Not (oFileFolder Is Nothing) Then
                    If oFileFolder.Attributes Mod 2 = 1 Then
                        ' Cannot create file so exit
                        FolderExistence = False
                        Exit Function
                    End If
                    ' Check can create files in this directory
                    Set oFile = oFileFolder.CreateTextFile("dummy.txt", True)
                    oFile.Close
                    oFileSystem.DeleteFile (oFileFolder & "\dummy.txt")
                End If
            End If
            Exit For
        End If
    Next
    
    ' Deallocate FileSystemObject Memory
    Set oFileSystem = Nothing
    Set oFileFolder = Nothing
    Set oFile = Nothing
    
    Exit Function
ErrLabel:
    ' If error is within file error range
    If Err.Number > 50 And Err.Number < 77 Then
        FolderExistence = False
        Exit Function
    End If
    Err.Raise Err.Number, , Err.Description & "|basCommon.FolderExistence"
End Function

'----------------------------------------------------------------------
Public Function Crypt(ByVal sText As String) As String
'----------------------------------------------------------------------
'TA 18/01/2002: Encrypt a string
' The string used to encrypt is hardcoded
'----------------------------------------------------------------------
Dim i As Long
Dim vKeyChar As Variant

    vKeyChar = Array(12, 145, 98, 229, 50)
    For i = 1 To Len(sText)
        Mid(sText, i, 1) = Chr(Asc(Mid(sText, i, 1)) Xor vKeyChar(i Mod 5))
    Next
    
    Crypt = sText

End Function

'----------------------------------------------------------------------------------------'
Public Function GetFileLength(sFile As String) As Long
'----------------------------------------------------------------------------------------'
' DPH 08/04/2002
' Returns File length - -1 for missing file
' FileLen throws error for missing files
'----------------------------------------------------------------------------------------'
Dim lFileLen As Long
  
    On Error GoTo ErrHandler

    ' initialise file length
    lFileLen = -1
    
    If FileExists(sFile) Then
        lFileLen = FileLen(sFile)
    End If
    
    GetFileLength = lFileLen
    
    Exit Function
    
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "GetFileLength", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function GetSiteStudySubjectLabFromFileName(ByVal sFileName As String) As String
'---------------------------------------------------------------------
'REM 04/06/03
'Returns the site, study and personid from the file name, or just the lab for LABS
'---------------------------------------------------------------------
Dim sSite As String
Dim sStudy As String
Dim sSubject As String
Dim sLab As String
Dim nPos As Integer
Dim vSiteStudySubject As Variant
Dim i As Integer

    On Error GoTo ErrorHandler

    ' Is it a lab or subject file
    If InStr(1, sFileName, "_LAB.", vbTextCompare) > 0 Then
        ' LAB File
        nPos = InStr(1, sFileName, "_LAB", vbTextCompare)
        sLab = Left(sFileName, nPos - 1)
        
        ' Return LAB
        GetSiteStudySubjectLabFromFileName = sLab
    Else
        sFileName = Replace(sFileName, ".zip", "")
        vSiteStudySubject = Split(sFileName, "_")
        
        Select Case UBound(vSiteStudySubject)
        Case 3 ' is old format
            sSite = vSiteStudySubject(1)
            sSubject = vSiteStudySubject(3)
            GetSiteStudySubjectLabFromFileName = sSite & "_" & sSubject
        Case Else ' new format
            sSite = vSiteStudySubject(1)
            sSubject = vSiteStudySubject(UBound(vSiteStudySubject))
            For i = 2 To UBound(vSiteStudySubject) - 2
                sStudy = sStudy & "_" & vSiteStudySubject(i)
            Next
            GetSiteStudySubjectLabFromFileName = sSite & sStudy & "_" & sSubject
        End Select
        
    End If

Exit Function
ErrorHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "GetSiteSubjectLabFromFileName", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'---------------------------------------------------------------------
Public Function GetSiteSubjectLabFromFileName(sFileName As String) As String
'---------------------------------------------------------------------
' Given a file return the site & subject as site_subject
'---------------------------------------------------------------------
Dim sSubject As String
Dim sSite As String
Dim sLab As String
Dim nPos As Integer
Dim nPosDot As Integer
    
    On Error GoTo ErrorHandler
    
    ' Is it a lab or subject file
    If InStr(1, sFileName, "_LAB.", vbTextCompare) > 0 Then
        ' LAB File
        nPos = InStr(1, sFileName, "_LAB", vbTextCompare)
        sLab = Left(sFileName, nPos - 1)
        
        ' Return LAB
        GetSiteSubjectLabFromFileName = sLab
    Else
        ' Subject file
        ' Get Site
        nPos = InStr(5, sFileName, "_", vbTextCompare)
        sSite = Mid(sFileName, 5, nPos - 5)
        
        ' Get Subject
        ' Pass over date
        nPos = InStr(nPos + 1, sFileName, "_", vbTextCompare)
        ' set nPos to beggining of subject string
        nPos = nPos + 1
        nPosDot = InStr(nPos, sFileName, ".", vbTextCompare)
        sSubject = Mid(sFileName, nPos, nPosDot - nPos)
        
        
        
        ' Return Subject
        GetSiteSubjectLabFromFileName = sSite & "_" & sSubject
    End If
    
    Exit Function

ErrorHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "GetSiteSubjectLabFromFileName", "basCommon")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'--------------------------------------------------------------------------------
Public Function GetMacroDBSetting(sSection As String, sKey As String, _
                                Optional MACROCon As ADODB.Connection = Nothing, _
                                Optional sDefault = "") As String
'--------------------------------------------------------------------------------
' Get required value from MACRODBSetting table
' Returns Empty string if not found
' REM 09/12/02 - added optional parameter for a MACRO DB connection
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsSetting As ADODB.Recordset
Dim sValue As String
Dim conMACRO As ADODB.Connection

    On Error GoTo ErrLabel
    
    'REM 06/12/02 - so can pass in a connection
    If Not MACROCon Is Nothing Then
        Set conMACRO = MACROCon
    Else
        Set conMACRO = MacroADODBConnection
    End If
    
    Set rsSetting = New ADODB.Recordset
    
    sValue = ""
    
    ' Set SQL
    sSQL = "SELECT SettingValue FROM MACRODBSetting WHERE SettingSection = '" & LCase(sSection) _
        & "' AND SettingKey = '" & LCase(sKey) & "'"
    rsSetting.Open sSQL, conMACRO, adOpenForwardOnly, adLockReadOnly
    ' get setting value
    If Not rsSetting.EOF Then
        sValue = LCase(rsSetting("SettingValue"))
    Else
        sValue = sDefault
    End If
    rsSetting.Close
    Set rsSetting = Nothing
    
    GetMacroDBSetting = sValue
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|basCommon.GetMacroDBSetting"
End Function

'--------------------------------------------------------------------------------
Public Sub SetMacroDBSetting(sSection As String, sKey As String, sValue As String)
'--------------------------------------------------------------------------------
' Set value in MACRODBSetting table
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim lRowsUpdated As Long
    
    On Error GoTo ErrLabel
    
    ' try to update firstly
    sSQL = "UPDATE MACRODBSetting SET SettingValue = '" & LCase(sValue) & "' WHERE SettingSection = '" _
            & LCase(sSection) & "' AND SettingKey = '" & LCase(sKey) & "'"
    MacroADODBConnection.Execute sSQL, lRowsUpdated, adCmdText
    
    ' if no update then must insert
    If lRowsUpdated = 0 Then
        sSQL = "INSERT INTO MACRODBSetting (SettingSection,SettingKey,SettingValue) VALUES ('" _
            & LCase(sSection) & "','" & LCase(sKey) & "','" & LCase(sValue) & "')"
        MacroADODBConnection.Execute sSQL, -1, adCmdText
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|basCommon.SetMacroDBSetting"
End Sub

'----------------------------------------------------------------------------
Public Function DoesTableExist(conMACRO As ADODB.Connection, sTableName As String) As Boolean
'----------------------------------------------------------------------------
'REM 12/02/03
'Check to see if a table exists in a database
'----------------------------------------------------------------------------
Dim sSQL As String
Dim rsTable As ADODB.Recordset

    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM " & sTableName
    Set rsTable = New ADODB.Recordset
    rsTable.Open sSQL, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    DoesTableExist = True
Exit Function
ErrHandler:
    DoesTableExist = False
End Function

'------------------------------------------------------------------------------
Public Function CheckForMacroTables(ByVal oConnection As ADODB.Connection, ByRef sVersion As String) As Boolean
'------------------------------------------------------------------------------
'Check there is not already a database
'
'Mo Morris  17/8/01 This function no longer displays a message.
'                   It now updates SVersion as well as return a Boolean (success/fail)
'REM 14/02/03 - moved here from frmConnectionString
'------------------------------------------------------------------------------
Dim rs As ADODB.Recordset
Dim sMessage As String

    On Error GoTo ErrDBExists
    Set rs = New ADODB.Recordset
    'next line causes an error and jumps to the error handler if there is no MACROControl table
    rs.Open "SELECT MACROVersion, BuildSubVersion FROM MACROCONTROL", oConnection
    sVersion = rs!MACROVersion & "." & rs!BuildSubVersion
    
    rs.Close
    Set rs = Nothing
    
    CheckForMacroTables = True
    Exit Function
    
ErrDBExists:
    Set rs = Nothing
    CheckForMacroTables = False
    Exit Function

End Function

'---------------------------------------------------------------------
Public Function DisplayGMTTime(ByVal dblTime As Double, ByVal sDateFormat As String, ByVal vTimezoneOffset As Variant) As String
'---------------------------------------------------------------------
'TA 22/05/2003: Function to convert a time and timezone offset into GMT
'should be static method in TimeZone class
'---------------------------------------------------------------------

    DisplayGMTTime = Format(CDate(dblTime), sDateFormat) & " " & DisplayGMTTimeZoneOffset(vTimezoneOffset)

End Function

'---------------------------------------------------------------------
Private Function DisplayGMTTimeZoneOffset(ByVal vTimezoneOffset As Variant) As String
'---------------------------------------------------------------------
'TA 22/05/2003: Function to convert a timezoneoffset as returned by TimeZone class into GMT offset
'should be static method in TimeZone class
'---------------------------------------------------------------------
Dim sText As String

    If IsNull(vTimezoneOffset) Then
        sText = ""
    Else
        sText = "(GMT"
        If vTimezoneOffset <> 0 Then
            If vTimezoneOffset < 0 Then
                sText = sText & "+"
            End If
            sText = sText & -vTimezoneOffset \ 60 & ":" & Format(Abs(vTimezoneOffset) Mod 60, "00")
        End If
                                                
        sText = sText & ")"
    End If

     DisplayGMTTimeZoneOffset = sText
 
End Function
