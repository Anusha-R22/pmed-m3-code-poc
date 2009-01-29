Attribute VB_Name = "basMainMACROModuleSCT"
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998-2001. All Rights Reserved
'   File:       basMainMACROModule.bas
'   Author:     Andrew Newbigging  June 1997
'   Purpose:    Main Module used in all MACRO modules
'------------------------------------------------------------------------------------------------'
'
'------------------------------------------------------------------------------------------------'
'   Revisions:
'   WillC Removed Comments 13/7/00
'
'   Mo 5 / 10 / 99
'       Added global variable gnTransactionControlLevel for indicating whether or not transaction
'       control is on. The reason for this being that the ADO/ODBC database connection does not
'       support nested Transactions.
'   NCJ 19 Oct 99   No longer need to check if other modules are running
'   Willc 20/10/99  Added check to ReferenceData to see if dsthe user has chosen a Security db
'                   and not a MacroDB.
'   Mo  25/10/99    MacroADODBConnection added as a global variable
'                   SecurityADODBConnection added as a global variable
'   Mo Morris   29/10/99
'                   DAO to ADO conversion.
'   SDM 23/11/99    Added RemoveNull function
'   NCJ 1/12/99     Added gblnRemoteSite
'   ATN 2/12/99     Added gsREPORT_LABEL
'   ATN 11/12/99    Modified Main to process command-line switches
'                   /ai     do an autoimport of messages not yet processed
'                   /tr     do a send/receive communication with MACRO Exchange server
'   ATN         13/12/99    Moved RemoveNull function into basCommon
'   NCJ 13/12/99    Added routine MACROEnd
'                   Use MACROEnd instead of End in error handlers
'   NCJ 21/12/99    SRs 2516,2419 Amended gIsDate to recognise 3 or 4 digit time values
'   NCJ 4/1/00      Bug fix to gIsDate
'   NCJ 6/1/00      Added "date" as a reserved word
'   NCJ 12/1/00     Added Lock check routine on studies/subjects for use in SD and DM
'   NCJ 13/1/00     Do not set UpdateMode in Lock check routines in this module
'   Mo  14/1/00     GetTimeStamp call removed from gIsDate
'   NCJ 18/1/00     Moved gsSecurityDatabasePassword to modADODBConnection
'   NCJ 26/1/00     TrialLockFileName and RecordLockFileName
'                   Also routines to delete the lock files
'   NCJ 27/1/00     Removed unused Report global variables
'   NCJ 21/2/00     Added gsGLOBAL_TEMP_PATH for DB lock files
'   TA 17/04/2000   Additional hourglass functions added
'   TA 10/05/2000   GetLockStatusText function added
'   WillC 13/7/00   SR3685 Trap for out of memory in the error handlers, raised by a crash when printing all forms.
'   NCJ 31/8/00     Changed Registry Key for version 2.1
'   NCJ 24/10/00    Added gsWEB_HTML_LOCATION
'   NCJ 8 Feb 01    Added VISIT, FORM and CYCLE to reserved words list
'   NCJ 27 Apr 01   Changed registry key location for 2.2
'   TA 04/07/2001:  2 new global variables to store the token for locks, so that we can unlock
'   DPH 18/10/2001  Added Check in InitialisationSettings to create Temp folder for Apps using Arezzo
'   DPH 18/10/2001  Removed References to TrialLockFileName, RecordLockFileName, CheckLockOnTrial
'                   TakeLockOnTrial, RemoveLockOnRecord, RemoveLockOnTrial, TakeLockOnTrial
'   NCJ 18/10/2001  Added GetPrologSwitches for setting Prolog memory
'   DPH 16/1/2002 - Corrected Web forms default location in InitialisationSettings
'   TA 17/1/02: modified GetApplicationTitle as part of VTRACK buglist build 1.0.3 Bug 6
'   DPH 10/04/2002 - Added gsSECURE_HTML_LOCATION global setting
'   Mo 15/4/2002    constant valDecimalPoint added
'   NCJ 20 Jun 02 - CBB 2.2.15/14&15 GetPrologSwitches - read all Prolog switches from Registry
'   ZA 01/07/2002 - Added global AREZZO memory constants
'   REM 03/07/02 - Module use to be called MainModule.bas
'   REM 03/07/02 - Added RQG code, new Public Const,
'   ZA  22/08/2002 - added form visit date picture
'   ZA/ASH 10/9/2002 - Added call to initialise IMED Settings
'   ASH 12/9/2002 - Changed all registry keys to be read from settings file.
'   ZA 18/09/2002 - added script path under that exists under www folder
'   TA 26/09/2002: New enumeration for User Interface colours
'   ASH 4/11/2002 - Added 2 new constants for PORTRAIT and LANDSCAPE eforms
'   NCJ 7 Nov 02 - Added gn_HOTLINK control type
'   NCJ 3 Jan 03 - Removed unused parameter from SilentUserLogin and corrected variable typo
'   TA 07/01/2003 - Added error handler to Main
'   Mo 13/1/2003    Changes made in MACRO 2.2 for Batch Data Entry and the Query module added.
'                   When the MACRO Batch Data Entry module is launched and calls Main a MACRO_BD
'                   specific call is made to new sub AddBatchDataEntryFunction.
'   Mo 13/1/2003    BDCommandLineLoginOK and CreateBDCLLogFile moved here from modBatchDataEntry
'   Mo 27/1/2003    AddBatchDataEntryFunction removed (now part of standard database)
'   NCJ 29 Jan 03 - Moved the Prolog Memory constants from here to basEnumerations
'   NCJ 23 Apr 03 - Temporary change to MACROHelp to display draft MACRO 3.0 Help File
'   NCJ 28 Aug 03 - Changed back to proper method of context-sensitive Help
'   ic 16/12/2005   added active directory login
'------------------------------------------------------------------------------------------'
Option Explicit
Option Base 0
Option Compare Binary

Public Const gsTRIAL_LABEL As String = "Trial"
Public Const gsLIBRARY_LABEL As String = "Library"
Public Const gsCRF_PAGE_LABEL As String = "CRFpage"
Public Const gsLARGE_CRF_PAGE_LABEL As String = "CRFpageLarge"
Public Const gsCRF_ELEMENT_LABEL As String = "CRFelement"
Public Const gsDATA_ITEM_LABEL As String = "Dataitem"
Public Const gsQGROUP_LABEL As String = "QGroup"
Public Const gsPAGE_DATA_LABEL As String = "CRFpageswithdataitems"
Public Const gsLINE_LABEL As String = "Line"
Public Const gsDOCUMENT_LABEL As String = "Document"
Public Const gsCOMMENT_LABEL As String = "CRFcomment"
Public Const gsLINK_LABEL As String = "Hotlink"     ' NCJ 6 Nov 02
Public Const gsVISIT_LABEL As String = "Visit"
Public Const gsPROFORMA_LABEL As String = "Proforma"
Public Const gsCOPY_LABEL As String = "Copy"
Public Const gsTRIAL_LIST_LABEL As String = "Listoftrials"
Public Const gsREPEATING_CRF_PAGE_LABEL As String = "RepeatingCRFpage"
Public Const gsINACTIVE_CRF_PAGE_LABEL As String = "InactiveCRFpage"
Public Const gsERROR_CRF_PAGE_LABEL As String = "ErrorCRFpage"
Public Const gsWARNING_CRF_PAGE_LABEL As String = "WarningCRFpage"
Public Const gsOKWARNING_CRF_PAGE_LABEL As String = "OKWarningCRFpage"
Public Const gsINFORM_CRF_PAGE_LABEL As String = "InformCRFpage"
Public Const gsTICK_CRF_PAGE_LABEL As String = "TickCRFpage"
Public Const gsDATA_COMMENT_LABEL As String = "Comment"
Public Const gsPICTURE_LABEL As String = "Picture"
Public Const gsBLANK_CRF_PAGE_LABEL As String = "BlankCRFpage"
Public Const gsTICK_LABEL As String = "Tick"
Public Const gsNEW_CRF_PAGE_LABEL As String = "NewCRFpage"
Public Const gsREPORT_LABEL As String = "Report"
'WillC SR3589 3/8/00
Public Const gsNOTAPPLICABLE_CRF_PAGE_LABEL As String = "NOTAPPLICABLECRFPAGE"
Public Const gsUNOBTAINABLECRFPAGE_CRF_PAGE_LABEL As String = "UNOBTAINABLECRFPAGE"

'ZA 22/08/2002
Public Const gsVISIT_EFORM = "Visit_eForm"
'WillC SR2673 29/8/00
Public Const gsMISSINGCRFPAGE_CRF_PAGE_LABEL As String = "MISSINGCRFPAGE"

' NCJ 27/4/01 - Removed unused Country & Subject labels

' NCJ 26/4/00 - Added gsVALIDATION_OKWARNING_LABEL
Public Const gsVALIDATION_MANDATORY_LABEL As String = "Mandatoryvalidation"


'TA 1/11/2002: These ahve been replaced
Public Const gsVALIDATION_WARNING_LABEL As String = "Warningvalidation"
Public Const gsVALIDATION_OKWARNING_LABEL As String = "OKWarning"
Public Const gsVALIDATION_DATA_MANAGER_LABEL As String = "Datamanagervalidation"
Public Const gsVALIDATION_HELP_TEXT_LABEL As String = "Helptext"
Public Const gsVALIDATION_OK_LABEL As String = "Tick"
Public Const gsVALIDATION_MISSING_LABEL As String = "Missing"
Public Const gsVALIDATION_UNOBTAINABLE_LABEL As String = "UnObtainable"

Public Const gsVALIDATION_NOT_APPLICABLE_LABEL As String = "NOTAPPLICABLE"
Public Const gsVALIDATION_LOCK_LABEL As String = "LOCK"
Public Const gsVALIDATION_FREEZE_LABEL As String = "FREEZE"


'TA 02/09/2002: New icnos for MACRO 3.0 UI Design
Public Const DM30_ICON_CHANGE_COUNT1 = "DM30_ChangeCount1"
Public Const DM30_ICON_CHANGE_COUNT2 = "DM30_ChangeCount2"
Public Const DM30_ICON_CHANGE_COUNT3 = "DM30_ChangeCount3"
'show all together - for status bar in DM
Public Const DM30_ICON_CHANGE_COUNTALL = "DM30_ChangeCountAll"
Public Const DM30_ICON_NOTE = "DM30_Note"
Public Const DM30_ICON_COMMENT = "DM30_Comment"
Public Const DM30_ICON_NOTE_COMMENT = "DM30_NoteComment"
Public Const DM30_ICON_RAISED_DISC = "DM30_RaisedDisc"
Public Const DM30_ICON_RESPONDED_DISC = "DM30_RespondedDisc"
Public Const DM30_ICON_FROZEN = "DM30_Frozen"
Public Const DM30_ICON_INFORM = "DM30_Inform"
Public Const DM30_ICON_LOCKED = "DM30_Locked"
Public Const DM30_ICON_MISSING = "DM30_Missing"
Public Const DM30_ICON_NA = "DM30_NA"
Public Const DM30_ICON_OK = "DM30_OK"
Public Const DM30_ICON_OK_WARNING = "DM30_OKWarning"
Public Const DM30_ICON_UNOBTAINABLE = "DM30_Unobtainable"
Public Const DM30_ICON_WARNING = "DM30_Warning"
Public Const DM30_ICON_INVALID = "DM30_Invalid"
'TA 21/10/200 New SDV icons
Public Const DM30_ICON_QUERIED_SDV = "DM30_QueriedSDV"
Public Const DM30_ICON_PLANNED_SDV = "DM30_PlannedSDV"
Public Const DM30_ICON_DONE_SDV = "DM30_DoneSDV"

Public Const DM30_ICON_NEW_FORM = "DM30_NewForm"
Public Const DM30_ICON_INACTIVE_FORM = "DM30_InactiveForm"
    
'ic 27/10/2005 added clinical coding
Public Const DM30_ICON_DICTIONARY = "DM30_DICTIONARY"
Public Const DM30_ICON_DICTIONARY_VALIDATED = "DM30_VDICTIONARY"
Public Const DM30_ICON_DICTIONARY_PENDING = "DM30_PDICTIONARY"
Public Const DM30_ICON_DICTIONARY_DONOT = "DM30_XDICTIONARY"
Public Const DM30_ICON_DICTIONARY_CODED = "DM30_CDICTIONARY"


Public Const gsCREATE As String = "Create"
Public Const gsREAD As String = "Read"
Public Const gsUPDATE As String = "Update"
Public Const gsDELETE As String = "Delete"
Public Const gsCopy As String = "Copy"

'changed by Mo Morris 15/12/99
'gsEXCHANGE_MODE and gsSYSTEM_MANAGER_MODE added
'gsSTUDY_DEFINITION_MODE changed to "SD" and gsTRIAL_SUBJECT_MODE changed to "DM"
Public Const gsSTUDY_DEFINITION_MODE = "SD"
Public Const gsTRIAL_SUBJECT_MODE = "DM"
Public Const gsEXCHANGE_MODE = "EX"
Public Const gsSYSTEM_MANAGER_MODE = "SM"

Public Const gsDATA_NOT_VALID As String = "This data is not valid"
Public Const gsDIALOG_TITLE As String = "MACRO"

Public Const gnFIRST_ID = 1
Public Const gnID_INCREMENT = 1

' Public nl As String

Public Const gnRECOMMENDED_MAXIMUM_NUMBER_OF_OPTION_BUTTONS = 3

' REM 07/12/01 - used to open the new .chm Help file
Public Const HH_HELP_CONTEXT = &HF

Public gsAppPath As String
Public gsMACROUserGuidePath As String

' NCJ 21/2/00 SR2129 Added gsGLOBAL_TEMP_PATH
'Public gsGLOBAL_TEMP_PATH As String
Public gsTEMP_PATH As String
Public gsTEMP_EXTENSION As String
Public gsTEMPLATE_PATH As String
Public gsTEMPLATE_EXTENSION As String
Public gsIN_FOLDER_LOCATION As String
Public gsOUT_FOLDER_LOCATION As String
'Public gsHTML_FORMS_LOCATION As String
Public gsDOCUMENTS_PATH As String   ' NCJ 1 Oct 99
Public gsCAB_EXTRACT_LOCATION As String 'SDM 26/01/00 SR2794
Public gsWEB_HTML_LOCATION As String    ' NCJ 24/10/00
'Public gsSECURE_HTML_LOCATION As String ' DPH 10/04/2002
Public gsSCRIPT_FOLDER_LOCATION As String      'ZA 18/09/2002

Public Const glPORTRAIT_WIDTH = 8515    'ASH 4/11/2002
Public Const glLANDSCAPE_WIDTH = 14500   'ASH 4/11/2002

Public Const gDefaultFontColour = -2147483630
Public Const gDefaultCRFPageColour = 12632256
Public Const gDefaultFontName = "Arial"
Public Const gDefaultFontBold = 0
Public Const gDefaultFontItalic = 0
Public Const gDefaultFontSize = 10

Public Const valAlpha                   As Integer = 1
Public Const valNumeric                 As Integer = 2
Public Const valSpace                   As Integer = 4
Public Const valOnlySingleQuotes        As Integer = 8
Public Const valComma                   As Integer = 16
Public Const valUnderscore              As Integer = 32
Public Const valDateSeperators          As Integer = 64
' PN 02/09/99
' allow includusion of mathematical operators in the check string
Public Const valMathsOperators          As Integer = 128
'Mo Morris 15/4/2002 valDecimalPoint added
Public Const valDecimalPoint            As Integer = 256

'Removed by Mo Morris 6/8/99
'Public gnSponsorId As Integer

Public gsTrialPhase() As String
Public gsTrialStatus() As String

'   SPR 427 ATN 8/10/98
'   Reference to Separate ImedSecurity module removed.

' NCJ 1/12/99 - Copied from MACRO 1.6
Public gblnRemoteSite As Boolean


'Public gUser As New IMedSecurity.User

Public goUser As MACROUser
' NCJ 13/1/00 - Removed unused function access constants (not used in 2.0)


Public Const gsIconArrowData = "ArrowData"
Public Const gsIconBox = "BOX"
Public Const gsIconBoxTicked = "BOXTICK"
Public Const gsIconCopy = "Copy"
Public Const gsIconNoDrop = "NoDrop"
Public Const gsIconMandatoryValidation = "MandatoryValidation"
Public Const gsIconWarningValidation = "WarningValidation"
Public Const gsIconDataManagerValidation = "DATAMANAGERVALIDATION"
Public Const gsIConMissing = "MISSING"
Public Const gsIConUnobtainable = "UNOBTAINABLE"
Public Const gsIConNotApplicable = "NOTAPPLICABLE"

'   ATN 1/3/99
'   New global variable to hold the security mode from the security database
Public gnSecurityMode As Integer
Public gsMacroVersion As String

' PN change 34 30/08/99
' new global for Single Use Data Items in study definition
Public gbSingleUseDataItems As Boolean

' PN 20/09/99
' the idle timeout for the application
Public glSystemIdleTimeout As Long
Public glSystemIdleTimeoutCount As Long

'   ATN 11/12/99
'   New command line switches
Public Const gsAUTO_IMPORT = "/AI"
Public Const gsTRANSFER_DATA = "/TR"

'Mo 5/10/99
'Macro's ADO/ODBC database connection does not support nested Transaction control.
'Macro contains code elements that contain/require transaction control. In some
'situations these code elements call each other and a nested BeginTrans call would
'cause an error. To prevent this every BeginTrans call is preceded by checking the
'value of gnTransactionControlOn:-
'   if its '0', BeginTrans will be called and gnTransactionControlOn incremented to '1'
'   if its > '0', BeginTrans will not be called, but gnTransactionControlOn is still incremented
'Every CommitTrans call is also preceded by checking the value of gnTransactionControlOn:-
'   If its '1', CommitTrans will be called and gnTransactionControlOn decremented
'   If its > '1', CommitTrans will not be called, but gnTransactionControlOn is still decremented
'
'Note that this is all handled by the Subs TransBegin, TRansCommit and TransRollback
Public gnTransactionControlOn As Integer

'Mo Morris 27/3/00
'MacroADODBConnection changed to a Property of ModADODBConnection
'Public MacroADODBConnection As ADODB.Connection
Public SecurityADODBConnection As ADODB.Connection


'WillC 8/11/99
Public gnTrappedErrNum As Long
Public gsTrappedErrDesc As String

Public Const glFormColour As Long = vbButtonFace

'added by Mo Morris 23/11/99
Public gnPrintAllFormsPageNumber As Integer

'Mo Morris 21/3/00, SR 3191 - Now unused (NCJ 24 Jan 03)
'Public gbDataBeingTransfered As Boolean

'TA 04/07/2001: store the token for locks, so that we can unlock
Public gsStudyToken As String
Public gsSubjectToken As String

'REM Global Security connection string
Public gsSecCon As String

' RJCW 26/11/01: stores Keywords
Private gsdicKeywords As Scripting.Dictionary

'Mo 13/1/2003
Public gsBDCLPathAndNameOfFile As String

'ic 27/10/2005 clinical coding on/off
Public gbClinicalCoding As Boolean
Private Const mCCSwitch As String = "CLINICALCODING"

'ic 07/12/2005 active directory on/off
Public gbActiveDirectory As Boolean
'is this user an active directory login user
Public gbIsActiveDirectoryLogin As Boolean
'windows username
Public gsWindowsCurrentUser As String
Private Const mActiveDirectorySwitch As String = "ACTIVEDIRECTORY"

' REM 07/12/01: MACRO help API call
Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
    (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
    
'ic 15/12/2005 get windows username
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    
    
'---------------------------------------------------------------------
Public Sub Main()
'---------------------------------------------------------------------
' The main startup routine for MACRO
' This now calls frmMenu.InitialiseMe for app-specific initialisations (NCJ 16/9/99)
' ic 16/12/2005 added active directory login
'---------------------------------------------------------------------

'   ATN 11/12/99
'   Check for command line switches:
'   /AI - do an 'AutoImport'
'   /TR - transfer data
'   Only applicable to data management at present
'Dim sSecCon As String
Dim blnShowUserInterface As Boolean
Dim blnLoginSucceeded As Boolean
Dim sLoginMsg As String
Dim sMessage As String
Dim sEncryptedDBPwd As String
Dim bMACRODesktop As Boolean
Dim oVersion As MACROVERSION.Checker


On Error GoTo Errlabel

    blnShowUserInterface = True
    
    'Mo 13/1/2003, Batch Data Entry Command Line facilities
    If App.Title = "MACRO_BD" Then
        If UCase(Left(Command, 3)) = "/BI" _
        Or UCase(Left(Command, 3)) = "/BU" Then
            blnShowUserInterface = False
        End If
    End If
    
    If App.Title = "MACRO_DM" Then
        If UCase(Left(Command, Len(gsAUTO_IMPORT))) = UCase(gsAUTO_IMPORT) _
        Or UCase(Left(Command, Len(gsTRANSFER_DATA))) = UCase(gsTRANSFER_DATA) Then
            blnShowUserInterface = False
        End If
    End If

    If blnShowUserInterface Then
        MousePointerChange vbArrowHourglass
        
        ' Show the splash form immediately to minimise the apparent wait
        frmSplash.Show
        'DoEvents to give splash screen time to paint
        DoEvents
    
    End If
    
    'Mo 5/10/99
    gnTransactionControlOn = 0
    
    'Mo Morris 21/3/00
'    gbDataBeingTransfered = False
    
    ' NCJ 1/12/99 - Initialise to remote site
    ' (Certain things are blocked for central site)
    gblnRemoteSite = True
    
    'ZA/ASH 10/09/2002 Initialise IMEDSettings component
    InitialiseSettingsFile
    
    ' Set up default paths etc. from settings file
    InitialisationSettings
    
#If DESKTOP = 1 Then
    'REM 26/03/04 - Check to see if its MACRO Desktop being run for the first time, if there is
    ' the keyword Desktop in the settings file with an encrypted string then means it is first time
    sEncryptedDBPwd = GetMACROSetting("Desktop", "")
    
    If sEncryptedDBPwd <> "" Then
        bMACRODesktop = True
        Call AttachMACRODesktopDB(sEncryptedDBPwd)
    Else
        bMACRODesktop = False
    End If
#End If
    
    'Changed by Mo Morris 28/10/99 - ADO InitializeSecurityADODBConnection call added
    gsSecCon = InitializeSecurityADODBConnection
    
#If DESKTOP = 1 Then
    'REM 26/03/04 - If it is MACRO Desktop being run for the first time then update DB password and remove
        'the Desktop reference from the MACROUserSettings30.txt file
    If bMACRODesktop Then
        Call UpdateDesktopDatabase(sEncryptedDBPwd)
        'then delete all the MSDE 2000 folders and files
        Call DeleteMSDEFiles
    End If
#End If

    'Mo 13/1/2003, Batch Data Entry Command Line facilities
    If App.Title = "MACRO_BD" Then
        If UCase(Left(Command, 3)) = "/BI" _
        Or UCase(Left(Command, 3)) = "/BU" Then
            'Note that if the command line paramaters are wrong then BDCommandLineLogin will return
            'False and normal login will take place because blnShowUserInterface will be set to True
            If BDCommandLineLoginOK(Command) Then
                blnShowUserInterface = False
            Else
                blnShowUserInterface = True
            End If
        End If
    End If
    
    If blnShowUserInterface Then
        ' Show main form
        'frmMenu.Show
        ' Get rid of splash
        Unload frmSplash
#If DM = 1 Then
        frmMenu.DisplayMDIBackGround
#End If
        MousePointerRestore
        
        'check for active directory login
        Call InitActiveDirectoryParams
        
        ' Show login dialog
        Set goUser = frmNewLogin.Display(gsSecCon)
    
        blnLoginSucceeded = Not goUser Is Nothing
        
    Else
        'Mo 13/1/2003, Batch Data Entry Command Line facilities
        'Only call SilentUserLogin when not in Batch Data Entry mode
        If App.Title <> "MACRO_BD" Then
            Set goUser = SilentUserLogin(gsSecCon, Mid(Command, Len(gsAUTO_IMPORT) + 2), gsAppPath & "HTML")
            blnLoginSucceeded = Not goUser Is Nothing
        End If
        
    End If
    
    If blnLoginSucceeded Or blnShowUserInterface = False Then
    
        If goUser.CheckPermission(gsFnStudyDefinition) Then
            Dim SCT As StudyCopy.StudyCopy
            Set SCT = New StudyCopy.StudyCopy
            Call SCT.Init(gsSecCon, goUser.CurrentDBConString, goUser.DatabaseCode, goUser.UserName, goUser.UserNameFull)
            Set SCT = Nothing
        Else
            Call DialogInformation("User does not have permission to access module", "Permission Denied")
        End If
    
        End
    
       
'        'TA 17/04/2000
'        HourglassOn "Connecting to database..."
'
'        'Changed by Mo Morris 27/10/99 - ADO InitializeMacroADODBConnection call added
'        InitializeMacroADODBConnection
'
'        ' Retrieve reference data
'        Call ReferenceData
'
'        'Changed Mo Morris 6/2/01, "MacroCreateDataViews" added plus spaces to each panels text
'        ' NCJ 13 Feb 01 - Include MACRO_DM here
'        'Mo Morris 12/12/01, "MACRO_QM" , Query Module added
'        'Mo Morris 13/1/2003, "MACRO_BD" , Batch Data Entry Module added
'        ' NCJ 26 Feb 03 - Added MACRO_BV
'        Select Case App.Title
'        Case "MACRO_SD", "MacroCreateDataViews", "MACRO_QM", "MACRO_SM", _
'            "MACRO_BD", "MACRO_BV", "MACRO_UT"
'            With frmMenu.Controls("sbrMenu").Panels
'                'TA 18/10/2002: reference the statusbar by name as that MACRO Data Entry compiles
'                .Item("UserKey").Text = "User: " & goUser.UserNameFull & " (" & goUser.UserName & ") "
'                .Item("RoleKey").Text = "Role: " & goUser.UserRole & " "
'                .Item("UserDatabase").Text = "Database: " & goUser.DatabaseCode & "          "
'            End With
'        End Select
'
'        'check for clinical coding version
'        Set oVersion = New MACROVERSION.Checker
'        gbClinicalCoding = oVersion.HasUpgrade(goUser.CurrentDBConString, mCCSwitch)
'        Set oVersion = Nothing
'
'        ' NCJ 16/9/99
'        Call frmMenu.InitialiseMe
'
'        'TA 17/04/2000
'        HourglassOff
    
    Else

        MACROEnd

    End If      ' If login succeeded
    
    
    If blnShowUserInterface Then
        'enable timer if DevMode is zero
        #If DEVMODE = 0 Then
                frmMenu.tmrSystemIdleTimeout.Enabled = True
        #End If
    Else
        End
    End If
Exit Sub


Errlabel:

    If MACROErrorHandler("basMainMACROModule", Err.Number, Err.Description, "Main", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function GetWindowsCurrentUser() As String
'----------------------------------------------------------------------------------------'
' ic 15/12/2005
' get current windows username
'----------------------------------------------------------------------------------------'
Dim lpBuff As String * 25
Dim ret As Long

    ret = GetUserName(lpBuff, 25)
    GetWindowsCurrentUser = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
End Function

'----------------------------------------------------------------------------------------'
Private Sub InitActiveDirectoryParams()
'----------------------------------------------------------------------------------------'
' ic 07/12/2005
' initialise active directory login variables
'----------------------------------------------------------------------------------------'
Dim oVersion As MACROVERSION.Checker
Dim sWinUserName As String
Dim rsADUser As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler

    'check for clinical coding version
    Set oVersion = New MACROVERSION.Checker
    gbActiveDirectory = oVersion.HasUpgrade(gsSecCon, mActiveDirectorySwitch)
    Set oVersion = Nothing
    
    If (gbActiveDirectory) Then
        'get the current windows username
        gsWindowsCurrentUser = GetWindowsCurrentUser
        
        'see if this windows username matches a macro username, and if so
        'if that user is flagged as AD login
        Set rsADUser = New ADODB.Recordset
        sSQL = "SELECT AUTHENTICATION FROM MACROUSER WHERE USERNAME = '" & gsWindowsCurrentUser & "'"
        rsADUser.Open sSQL, gsSecCon, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not (rsADUser.BOF And rsADUser.EOF) Then
            gbIsActiveDirectoryLogin = (ConvertFromNull(rsADUser!Authentication, vbInteger) = 1)
        End If
        
        rsADUser.Close
        Set rsADUser = Nothing
    End If
        
ErrHandler:
    If MACROErrorHandler("basMainMACROModule", Err.Number, Err.Description, "InitActiveDirectoryParams", Err.Source) = Retry Then
        Resume
    End If
End Sub

'----------------------------------------------------------------------------------------'
Private Sub AttachMACRODesktopDB(ByVal sEncryptedDBPwd As String)
'----------------------------------------------------------------------------------------'
'REM 26/03/04
'Attach the MACRO30 MSDE database to the MSDE instance installed with MACRO desktop
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim sDBPassword As String
Dim conDB As ADODB.Connection

    On Error GoTo Errlabel

    sDBPassword = DecryptString(sEncryptedDBPwd)

    'Attach MACRO 3.0 database file
    sSQL = "EXEC sp_attach_db @dbname = 'MACRO30', @filename1 = '" & App.Path & "\Database\MACRO30.mdf'"
    Set conDB = New ADODB.Connection
    conDB.Open "PROVIDER=SQLOLEDB;DATA SOURCE=localhost;USER ID=sa;PASSWORD=" & sDBPassword & ";"
    conDB.Execute sSQL

    conDB.Close
    Set conDB = Nothing
'
Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "basMainMACROModule.AttachMACRODesktopDB"
End Sub

'----------------------------------------------------------------------------------------'
Private Sub UpdateDesktopDatabase(sEncryptedPassword As String)
'----------------------------------------------------------------------------------------'
'REM 26/03/04
'Routine to update the new MACRO Desktop security databases Database password field in the Databases table
'and create new MACROUserSetting30.txt that only contains the security path
'----------------------------------------------------------------------------------------'
Dim sSQL As String
    
    On Error GoTo Errlabel

    'update the database password
    sSQL = "UPDATE Databases SET DatabasePassword = '" & sEncryptedPassword & "'" _
         & "WHERE DatabaseCode = 'MACRO30'"
    SecurityADODBConnection.Execute sSQL
    
    'Get rid of the MACROUserSettings30.txt
    Kill App.Path & "\MACROUserSettings30.txt"
    
    'Create new MACROUserSettings30.txt to hold the securitypath
    StringToFile App.Path & "\" & "MACROUserSettings30.txt", "securitypath=" & EncryptString(gsSecCon)

Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "basMainMACROModule.UpdateDesktopDatabase"
End Sub

'----------------------------------------------------------------------------------------'
Private Sub DeleteMSDEFiles()
'----------------------------------------------------------------------------------------'
'REM 26/03/04
'Delete the MSDE 2000 folders and files from the MACRO 3.0 folder
'----------------------------------------------------------------------------------------'
Dim sFolders As String
Dim vFolders As Variant
Dim i As Integer
Dim sNextFile As String
Dim sInstallPath As String
    
    On Error GoTo Errlabel
    
    'Delete MSDE 2000 setup files
    'List of all MSDE folders
    sFolders = "MSDE_2000\Setup;MSDE_2000\MSM\1033;MSDE_2000\MSM;MSDE_2000\Msi;MSDE_2000"
    vFolders = Split(sFolders, ";")
    sInstallPath = App.Path & "\"
    'loop through all folders, delete contents then delete folder
    For i = 0 To UBound(vFolders)
        'get each file in folder
        sNextFile = Dir(sInstallPath & vFolders(i) & "\" & "*.*")
        Do While sNextFile <> ""
            'delete file
            Kill sInstallPath & vFolders(i) & "\" & sNextFile
            Do Until Not FileExists(sNextFile)
                DoEvents
            Loop
            'get next file via the DIR command
            sNextFile = Dir(sInstallPath & vFolders(i) & "\" & "*.*")
        Loop
        'delete empty folder
        RmDir (sInstallPath & vFolders(i))
    Next

Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "basMainMACROModule.AttachMACRODesktopDB"
End Sub

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

'---------------------------------------------------------------------
Public Sub MACROHelp(ByVal lhWnd As Long, ByVal AppTitle As String)
'---------------------------------------------------------------------
' REM 07/12/01
' Opens MACRO help based on the context ID of the different MACRO modules
'REVISIONS:
' REM 17/01/02 - added Case MacroCreateDataViews as it did not exist
' Mo 10/7/2002, CBB 2.2.18.6, New context Id call to MACRO_QM added
' NCJ 28 Aug 03 - Go back to using Help module IDs
'---------------------------------------------------------------------
Dim hwndHelp
Dim sHelpFile As String

'    ' NCJ 23 Apr 03 - Temporary call to show draft MACRO 3.0 Help File
'    Call ShowDocument(lhWnd, gsMACROUserGuidePath & "MACRO3Help.chm")
'    Exit Sub

'Context ID's for each module in MACRO Help
' NB These have been defined in RoboHelp so DON'T change them here!
Const lDATAENTRY As Long = 1
Const lDATAREVIEW As Long = 2
Const lWEBDEDR As Long = 3
Const lLIBRARYMANAGEMENT = 4
Const lSTUDYDEFINITION = 5
'Const lEXCHANGE = 6
Const lSYSTEMMANAGEMENT = 7
Const lCREATEDATAVIEWS = 8
Const lMACROWELCOME = 9
Const lQUERYMODULE = 10
Const lBATCHDATAENTRY = 11
Const lBATCHVALIDATION = 12
    
    ' Calls the help which is contained in the MACRO.chm file that
    ' requires a context ID to open the specific help for each module
    
    sHelpFile = gsMACROUserGuidePath & "MACRO3Help.chm"
    
    Select Case AppTitle
    Case "MACRO_SD"
        If LCase$(Command) = "library" Then
            hwndHelp = HtmlHelp(lhWnd, sHelpFile, HH_HELP_CONTEXT, lLIBRARYMANAGEMENT)
        Else
            hwndHelp = HtmlHelp(lhWnd, sHelpFile, HH_HELP_CONTEXT, lSTUDYDEFINITION)
        End If
    Case "MACRO_DM"
         If LCase$(Command) = "review" Then
            hwndHelp = HtmlHelp(lhWnd, sHelpFile, HH_HELP_CONTEXT, lDATAREVIEW)
        Else
            hwndHelp = HtmlHelp(lhWnd, sHelpFile, HH_HELP_CONTEXT, lDATAENTRY)
        End If
    Case "MACRO_SM"
        hwndHelp = HtmlHelp(lhWnd, sHelpFile, HH_HELP_CONTEXT, lSYSTEMMANAGEMENT)
    'REM 17/01/02 - Added MacroCteateDataViews to Help
    Case "MacroCreateDataViews"
        hwndHelp = HtmlHelp(lhWnd, sHelpFile, HH_HELP_CONTEXT, lCREATEDATAVIEWS)
    'Mo 10/7/2002, CBB 2.2.18.6, New context Id call to MACRO_QM added
    Case "MACRO_QM"
        hwndHelp = HtmlHelp(lhWnd, sHelpFile, HH_HELP_CONTEXT, lQUERYMODULE)
    Case "MACRO_BD"
        hwndHelp = HtmlHelp(lhWnd, sHelpFile, HH_HELP_CONTEXT, lBATCHDATAENTRY)
    Case "MACRO_BV"
        hwndHelp = HtmlHelp(lhWnd, sHelpFile, HH_HELP_CONTEXT, lBATCHVALIDATION)
    Case Else
        hwndHelp = HtmlHelp(lhWnd, sHelpFile, HH_HELP_CONTEXT, lMACROWELCOME)
    End Select

    
End Sub

'---------------------------------------------------------------------
Public Sub MACROEnd()
'---------------------------------------------------------------------
' NCJ 13 Dec 99
' End statement for MACRO
' (This should not be done if compiling a DLL)
'---------------------------------------------------------------------

    End
    
End Sub

'---------------------------------------------------------------------
Public Sub UnloadAllChildForms(bUnloadfrmMenu As Boolean)
'---------------------------------------------------------------------
' unload all forms in the project including frmmenu if bUnloadfrmMenu = True
'---------------------------------------------------------------------
Dim oForm As Form

    For Each oForm In Forms
        If oForm.Name <> "frmMenu" Or bUnloadfrmMenu Then
            ' call the AutoClose method to allow forms to close
            ' in a controlled way
            ' Ignore error where a form does not support this method
            On Error Resume Next
            oForm.AutoClose
            On Error GoTo 0
            
            Unload oForm
        End If
    Next oForm

End Sub

'---------------------------------------------------------------------
Public Sub UnloadAllForms()
'---------------------------------------------------------------------
' TA 25/04/2000
' unload all forms in the project active forms first, frmmenu last
' the quit application
' NCJ 20 Mar 03 - Make sure we deal with frmEFormDataEntry FIRST
' NCJ 24 Mar 03 - Only do this for DM!!
'---------------------------------------------------------------------
Dim oForm As Form
    
    HourglassOn "Closing down..."
    
    'first do modal forms
    Set oForm = Screen.ActiveForm
    Do While Not oForm Is Nothing
        If oForm.Name <> "frmMenu" Then
            Unload oForm
            Set oForm = Screen.ActiveForm
        Else
            Exit Do
        End If
    
    Loop

#If DM = 1 Then
    ' NCJ 20 Mar 03 - Deal with data entry form first
    ' Give it a chance to tidy up
    If frmMenu.IsDataEntryFormLoaded(False) Then
        Call frmEFormDataEntry.eFormAction(eaCancel)
    End If
#End If

    'then the rest
    For Each oForm In Forms
        If oForm.Name <> "frmMenu" Then
            ' call the AutoClose method to allow forms to close
            ' in a controlled way
            ' catch error where a form does not support this method
            On Error Resume Next
            oForm.AutoClose
            On Error GoTo 0
            Unload oForm
        End If
    Next oForm
    
    'finally menu form
    Unload frmMenu

    HourglassOff
    
    'this will cause a crash in debug mode
    End

End Sub
'
''---------------------------------------------------------------------
'Public Function DoTimeoutLogin() As Boolean
''---------------------------------------------------------------------
'' attempt to login after the system idle timeout has been reached
''---------------------------------------------------------------------
'Dim bResult As Boolean
'
'    With frmNewLogin
'
'        ' setup the form for password confirmation
'        .LoginSucceeded = False
'        .CheckPasswordOnly = True
'        .txtUserName.Text = goUser.UserName
'        .txtPassword.Text = vbNullString
'
'        ' show and then act upon result
'        .Show vbModal
'        If .LoginSucceeded Then
'            bResult = True
'            Call RestartSystemIdleTimer
'
'        Else
'            bResult = False
'
'        End If
'
'    End With
'    DoTimeoutLogin = bResult
'
'End Function

'---------------------------------------------------------------------
Public Sub ExitMACRO()
'---------------------------------------------------------------------
Dim mDatabase As Database
        
    '#If Exchange = 1 Then
    '    ExportData
    '#End If
    
    ' PN 24/09/99 call routine to unload forms
    ' Unload all forms except mdi parent
    Call UnloadAllChildForms(False)

    ' Unload mdi parent, if it exists
    Call UnloadAllChildForms(True)
    
    'Close all open files
    Close
    
    'Close database connections
    For Each mDatabase In DBEngine(0).Databases
        mDatabase.Close
    Next
    ' NCJ 15 Sept 99
    ' Close down CLM and PSS in frmMenu rather than here
    
    ' PN 22/09/99 - the end statement is not needed here because this proc
    ' is called when the main form is unloaded
    ' it also causes the app to intermittently crash on shutdown
    'Terminate the application
    'End
    
    ' PN 23/09/99
    ' ensure that the ado connection is terminated properly
    Call TerminateAllADODBConnections
    
End Sub

'---------------------------------------------------------------------
Private Sub ReferenceData()
'----------------------------------------------------------------------------------------------
'Added the Error Handling Routine to cope with a user choosing  a Security database
'instead of a MAcro database by trapping for the error raised by a missing table.
'---------------------------------------------------------------------------------------------
Dim rsReferenceData As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    ' read the trial phase data
    sSQL = "SELECT * FROM TrialPhase ORDER BY PhaseId"
    Set rsReferenceData = New ADODB.Recordset
    rsReferenceData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    While Not rsReferenceData.EOF
        ReDim Preserve gsTrialPhase(rsReferenceData!PhaseId)
        gsTrialPhase(rsReferenceData!PhaseId) = rsReferenceData!PhaseName
        rsReferenceData.MoveNext
    Wend
    rsReferenceData.Close
    
    ' read the trial status data
    sSQL = "SELECT *  FROM TrialStatus ORDER BY StatusId"
    Set rsReferenceData = New ADODB.Recordset
    rsReferenceData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    While Not rsReferenceData.EOF
        ReDim Preserve gsTrialStatus(rsReferenceData!statusId)
        gsTrialStatus(rsReferenceData!statusId) = rsReferenceData!StatusName
        rsReferenceData.MoveNext
    Wend
    rsReferenceData.Close
    
    ' nl = Chr(13)

    ' PN 20/09/99
    ' read the system idle timeout parameter and set the timer global
    ' globals are used because the timer interval max = 1 minute
    sSQL = "Select IdleTimeout From MacroControl"
    Set rsReferenceData = New ADODB.Recordset
    rsReferenceData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    glSystemIdleTimeout = rsReferenceData!IdleTimeout
    glSystemIdleTimeoutCount = 0
    
    ' temp value for testing purposes only
    With frmMenu.tmrSystemIdleTimeout
        .Enabled = False
        ' set the timer to 1 minute(since 65000 is the limit)
        ' the global counter will keep track of how many minutes have passed
        .Interval = 60000
        
'enable timer if DevMode is zero
#If DEVMODE = 0 Then
        .Enabled = True
#End If

    End With
    rsReferenceData.Close

    Set rsReferenceData = Nothing
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "basMainMACROModule.ReferenceData"
'    Select Case Err.Number
'        Case 3078   ' Special error to trap here
'         MsgBox "This is not a MACRO database. Please choose a valid" + vbCr _
'              & " MACRO database.", vbInformation, "MACRO"
'         frmLogin.Show
'         Exit Sub
'        Case -2147217865 'SQL error
'         MsgBox "This is not a MACRO database. Please choose a valid" + vbCr _
'              & " MACRO database.", vbInformation, "MACRO"
'         frmLogin.Show
'         Exit Sub
'        Case Else   ' Otherwise do general error handler
'            Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
'                                    "ReferenceData", "MainMACROModule")
'                  Case OnErrorAction.Ignore
'                      Resume Next
'                  Case OnErrorAction.Retry
'                      Resume
'                  Case OnErrorAction.QuitMACRO
'                      Call ExitMACRO
'                      Call MACROEnd
'             End Select
'    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub InitialisationSettings()
'---------------------------------------------------------------------
' Set up application folders and paths
' REVISIONS
' DPH 16/1/2002 - Corrected Web forms default location
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'set-up a Application Path variable
    gsAppPath = App.Path
                                
    AddDirSep gsAppPath
    
    'set-up specific Path variables for application folders
    ' NB every PATH or LOCATION variable ends with a backslash
    
    ' ASH 11/9/2002 Using new IMEDSettings component to get named settings
    ' or return the default passed
     gsTEMP_PATH = GetMACROSetting("Temp", gsAppPath & "Temp\")
    
    ' DPH 18/10/2001 - Need to check for TEMP folders existence - vital to AREZZO
    ' Only Applicable on MACRO_DM, MACRO_SD - Where Arezzo is used
    ' Quit if no Temp Folder
    If App.Title = "MACRO_SD" Or App.Title = "MACRO_DM" Then
        If Not (FolderExistence(gsTEMP_PATH & "dummy.txt")) Then
            ' Show User error Message
            Call DialogError("Temporary Path " & gsTEMP_PATH & vbCrLf & "could not be created or has no write permissions." & vbCrLf & "MACRO cannot work without this folder")
            ' Quit
            End
        End If
    End If
    
    ' ASH 11/9/2002 Using new IMEDSettings component to get named settings
    ' or return the default passed
    gsIN_FOLDER_LOCATION = GetMACROSetting("In Folder", gsAppPath & "In Folder\")
    gsOUT_FOLDER_LOCATION = GetMACROSetting("Out Folder", gsAppPath & "Out Folder\")
    gsDOCUMENTS_PATH = GetMACROSetting("Documents", gsAppPath & "Documents\")
    gsCAB_EXTRACT_LOCATION = GetMACROSetting("CabExtract", gsAppPath & "CabExtract\")
    gsWEB_HTML_LOCATION = GetMACROSetting("Web HTML", gsAppPath & "www\")
    gsMACROUserGuidePath = GetMACROSetting("Help", gsAppPath & "Help\")
    gsSCRIPT_FOLDER_LOCATION = GetMACROSetting("Script Folder", gsAppPath & "www\script\")
     
     
    ' NCJ 21/2/00 SR2129 Added gsGLOBAL_TEMP_PATH
    ' DPH 18/10/2001 - Removed Global Temp Path as not used
'    gsGLOBAL_TEMP_PATH = GetFromRegistry(GetMacroRegistryKey, "GlobalTemp")
'    If Len(gsGLOBAL_TEMP_PATH) = 0 Then
'        ' Default to local temp folder
'        gsGLOBAL_TEMP_PATH = gsTEMP_PATH
'    End If
    
    'gsIN_FOLDER_LOCATION = GetFromRegistry(GetMacroRegistryKey, "In Folder")
    'If Len(gsIN_FOLDER_LOCATION) = 0 Then
        'gsIN_FOLDER_LOCATION = gsAppPath & "In Folder\"
    'End If
    
'    gsOUT_FOLDER_LOCATION = GetFromRegistry(GetMacroRegistryKey, "Out Folder")
'    If Len(gsOUT_FOLDER_LOCATION) = 0 Then
'        gsOUT_FOLDER_LOCATION = gsAppPath & "Out Folder\"
'    End If
    
'    gsDOCUMENTS_PATH = GetFromRegistry(GetMacroRegistryKey, "Documents")
'    If Len(gsDOCUMENTS_PATH) = 0 Then
'         gsDOCUMENTS_PATH = gsAppPath & "Documents\"
'    End If
   
'    gsCAB_EXTRACT_LOCATION = GetFromRegistry(GetMacroRegistryKey, "CabExtract")
'    If Len(gsCAB_EXTRACT_LOCATION) = 0 Then
'         gsCAB_EXTRACT_LOCATION = gsAppPath & "CabExtract\"
'    End If
    
'    ' NCJ 24/10/00 - Web HTML folder location
'    gsWEB_HTML_LOCATION = GetFromRegistry(GetMacroRegistryKey, "Web HTML")
'    If Len(gsWEB_HTML_LOCATION) = 0 Then
'        ' DPH 16/1/2002 - Corrected Web forms default location
'         gsWEB_HTML_LOCATION = gsAppPath & "WWW\"
'    End If
    
    ' NCJ 13/1/00 - UserGuidePath is now just generic Help directory
'    gsMACROUserGuidePath = gsAppPath & "Help\"
    
    'set-up file extention variables
    gsTEMP_EXTENSION = ".tmp"
    
    Exit Sub
    
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "InitialisationSettings", "MainMACROModule")
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
Public Function GetTransferStatusText(nTransferStatus As Changed)
'---------------------------------------------------------------------
' Get text to display for given "Changed" status
' Used in the Data Browser
'---------------------------------------------------------------------

    Select Case nTransferStatus
        Case Changed.Changed
            GetTransferStatusText = "Not exported"
        Case Changed.Imported
            GetTransferStatusText = "New"
        Case Changed.NoChange
            GetTransferStatusText = "Exported"
        Case Else
            GetTransferStatusText = ""
    End Select

End Function

'---------------------------------------------------------------------
Public Function GetLockStatusText(nStatus As LockStatus) As String
'---------------------------------------------------------------------
' Get the text that represents a question lock status
'---------------------------------------------------------------------
Dim sLockStatus As String

    Select Case nStatus
    Case LockStatus.lsFrozen: sLockStatus = "Frozen"
    Case LockStatus.lsLocked: sLockStatus = "Locked"
    Case LockStatus.lsPending: sLockStatus = "Pending"
    Case LockStatus.lsUnlocked: sLockStatus = ""
    End Select
    
    GetLockStatusText = sLockStatus

End Function

'---------------------------------------------------------------------
Public Function GetStudyStatusText(nStatus As eTrialStatus) As String
'---------------------------------------------------------------------
' Get the text that represents a question's lock status
'---------------------------------------------------------------------

Dim sStudyStatus As String

    Select Case nStatus
    Case eTrialStatus.ClosedToFollowUp: sStudyStatus = "Closed to Follow Up"
    Case eTrialStatus.ClosedToRecruitment: sStudyStatus = "Closed to Recruitment"
    Case eTrialStatus.InPreparation: sStudyStatus = "In Preparation"
    Case eTrialStatus.Suspended: sStudyStatus = "Suspended"
    Case eTrialStatus.TrialOpen: sStudyStatus = "Open"
    End Select
    
    GetStudyStatusText = sStudyStatus

End Function

'---------------------------------------------------------------------
Public Function GetStatusText(vStatus As Integer) As String
'---------------------------------------------------------------------
' Get the text that represents a question status
' NCJ 16/5/00 SR3452 Show Inform as OK for Investigators
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
'   ATN 26/4/99
'   New values added from Locked and Frozen data
' NCJ 26/4/00 Added OK Warning; removed Locked, Frozen and Pending
 'Mo 24/5/00 SR 3501 Deleted removed
    Select Case vStatus
    Case Status.Requested
        GetStatusText = "Blank"
    Case Status.CancelledByUser
        GetStatusText = "Cancelled by user"
    Case Status.Success
        GetStatusText = "OK"
    Case Status.Missing
        GetStatusText = "Missing"
    Case Status.Inform
        If goUser.CheckPermission(gsFnMonitorDataReviewData) Then
            GetStatusText = "Inform"
        Else
            GetStatusText = "OK"
        End If
    Case Status.Warning
        GetStatusText = "Warning"
    Case Status.OKWarning
        GetStatusText = "OK Warning"
    Case Status.InvalidData
        GetStatusText = "Error"
    Case Status.Unobtainable
        GetStatusText = "Unobtainable"
    Case Status.NotApplicable
        GetStatusText = "Not applicable"
    End Select
Exit Function
    
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "GetStatusText", "MainMACROModule")
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
Public Function gblnNotAReservedWord(ByVal vString As String) As Boolean
'---------------------------------------------------------------------
' Returns TRUE if vString is not a "reserved word" of MACRO
' added by Mo Morris 6/1/99 SR 668
' NCJ 17 Sept 99 - Added "Registration" and "Randomisation"
'       Rewrote as Case statement
' NCJ 6 Jan 00 - Added "Date" (because of form dates)
' NCJ 21/3/00 SR 2811 Added STATUS here (not in Arezzo because of
'           backwards compatibility with Roche Demo. This stops new
'           questions being created called "status" but it stops Arezzo
'           complaining about existing ones).
' NCJ 8 Feb 01 - Added VISIT, FORM and CYCLE
' Mo Morris 27/2/01 The following field names from Data View tables added:-
'           "CLINICALTRIALID", "SITE", "PERSONID", "VISITID",
'           "VISITCYCLENUMBER", "CRFPAGEID", "CRFPAGECYCLENUMBER"
' Mo Morris 6/3/01 "YESNO" added
' Mo Morris 16/3/01 "TEXT", "MEMO", "NUMBER", "TIME", "CURRENCY", "CATEGORY", "AUTONUMBER", "YES", "NO",
'           "HYPERLINK", "VARCHAR", "SMALLINT", "INTEGER", "REAL", "LONG", "TINYINT", "BYTE" added
' Mo Morris 2/4/01 "TEMP" added
' Mo Morris 17/4/01 "VALIDATE" added
' Mo Morris 7/9/01 "LOCAL","PASSWORD","NAME" added
' RJCW 26/11/01 Changed this procedure to search a new dictionary object,
'               rather than using the previous select case method
'               also reviewed and extended the number of keywords to 309
'---------------------------------------------------------------------------------------------------

    gblnNotAReservedWord = Not gsdicKeywords.Exists(UCase(vString))
    
End Function

'------------------------------------------------------------------------------------
Private Sub ReplaceColumn(sTableName As String, sOldFieldName As String, _
                          sNewFieldName As String, sFieldSize As String, _
                          oMacroDatabase As Database)
'------------------------------------------------------------------------------------
' This routine will rename the column passed in by copying the data to a temp column
'------------------------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrHandler
    ' PN 08/09/99
    ' change to sTableName table
    sSQL = "Alter table " & sTableName & " Add Column TempColumn " & sFieldSize
    oMacroDatabase.Execute sSQL, dbFailOnError
    ' copy the contents in original into temp column
    sSQL = "UPDATE " & sTableName & " SET TempColumn = " & sOldFieldName
    oMacroDatabase.Execute sSQL, dbFailOnError
    ' drop original column
    sSQL = "ALTER Table " & sTableName & " DROP COLUMN [" & sOldFieldName & "]"
    oMacroDatabase.Execute sSQL, dbFailOnError
    ' recreate new column with new name
    sSQL = "Alter table " & sTableName & " Add Column " & sNewFieldName & " " & sFieldSize
    oMacroDatabase.Execute sSQL, dbFailOnError
    ' copy the contents in temp into new column
    sSQL = "UPDATE " & sTableName & " SET " & sNewFieldName & " = TempColumn"
    oMacroDatabase.Execute sSQL, dbFailOnError
    sSQL = "ALTER Table " & sTableName & " DROP COLUMN TempColumn"
    oMacroDatabase.Execute sSQL, dbFailOnError
 
    Exit Sub
    
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "ReplaceColumn", "MainMACROModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'------------------------------------------------------------
Public Function gIsDate(ByRef rDate As String) As Boolean
'------------------------------------------------------------
' Determine if rDate represents a valid date/time
' Return "standard" date/time representation in rDate
' NCJ 21/12/99 - Check for 3 or 4 digit numbers representing times
' NCJ 4/1/00 - Bug fix to time checking! And check year >= 1600
'------------------------------------------------------------

Dim msDate As String
Dim msTempDateString1 As String
Dim msTempDateString2 As String
Dim msTempDateString3 As String
Dim msTempDateString4 As String
Dim msTempDate As String

    On Error GoTo ErrHandler
    
    ' Initialise result
    gIsDate = False
    
    msDate = rDate
    
    If LCase$(msDate) = "t" Then
        'changed Mo Morris 14/1/00, SR2603, call to GetTimeStamp removed
        msDate = Format(Now, "dd/mm/yyyy hh:mm:ss")
        
    ElseIf Len(msDate) = 6 And IsNumeric(msDate) Then
        msDate = Left$(msDate, 2) & "/" & Mid$(msDate, 3, 2) & "/" & Right$(msDate, 2)
        
    ElseIf Len(msDate) = 8 And IsNumeric(msDate) Then
        msDate = Left$(msDate, 2) & "/" & Mid$(msDate, 3, 2) & "/" & Right$(msDate, 4)
        
    ' Try to interpret times...
    ElseIf Len(msDate) <= 4 And IsNumeric(msDate) Then
        ' Try to interpret times of 0000 to 2359
        If Len(msDate) = 4 And CInt(msDate) < 2400 Then
            msDate = Left$(msDate, 2) & ":" & Mid$(msDate, 3, 2)
        ' Try to interpret times of 000 to 959
        ElseIf Len(msDate) = 3 And CInt(msDate) < 960 Then
            msDate = Left$(msDate, 1) & ":" & Mid$(msDate, 2, 2)
        End If
        
    Else
        msDate = Replace(msDate, ".", "/")
        msDate = Replace(msDate, " ", "/")
        
    End If
    
    '   Check for four slashes - maybe dd/mm/yyyy/hh:mm:ss after previous step
    msTempDate = msDate
    msTempDateString1 = ExtractFirstItemFromList(msTempDate, "/")
    msTempDateString2 = ExtractFirstItemFromList(msTempDate, "/")
    msTempDateString3 = ExtractFirstItemFromList(msTempDate, "/")
    msTempDateString4 = ExtractFirstItemFromList(msTempDate, "/")
    
    If msTempDateString4 > "" Then
        msDate = msTempDateString1 & "/" & msTempDateString2 & "/" & msTempDateString3 & " " & msTempDateString4
    End If
    
    If IsDate(msDate) Then
        ' NCJ 4/1/00, SR 2017 - Check the year is 1600 or later
        If DatePart("yyyy", msDate) >= 1600 Then
            rDate = msDate
            gIsDate = True
        End If
    End If
     
    Exit Function
    
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "gIsDate", "MainMACROModule")
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
Public Function gblnIsAlphaNumericWithDelete(ByVal lAsciiKeyCode As Long) As Boolean
'---------------------------------------------------------------------
' this function will determine if a key code is an alpha numeric character
' or is a delete or backspace keystroke
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If Not lAsciiKeyCode = vbKeyDelete And Not lAsciiKeyCode = vbKeyBack Then
        If Not gblnValidString(Chr(lAsciiKeyCode), valAlpha + valNumeric) Then
            gblnIsAlphaNumericWithDelete = False
        Else
            gblnIsAlphaNumericWithDelete = True
        End If
    Else
        gblnIsAlphaNumericWithDelete = True
    End If
 
Exit Function
    
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "gblnIsAlphaNumericWithDelete", "MainMACROModule")
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
Public Function gblnIsNumericWithDelete(ByVal lAsciiKeyCode As Long) As Boolean
'---------------------------------------------------------------------
' this function will determine if a key code is an alpha numeric character
' or is a delete or backspace keystroke
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If Not lAsciiKeyCode = vbKeyDelete And Not lAsciiKeyCode = vbKeyBack Then
        If Not gblnValidString(Chr(lAsciiKeyCode), valNumeric) Then
            gblnIsNumericWithDelete = False
        Else
            gblnIsNumericWithDelete = True
        End If
    Else
        gblnIsNumericWithDelete = True
    End If
 
Exit Function
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "gblnIsNumericWithDelete", "MainMACROModule")
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
Public Function MACROFormErrorHandler(oForm As Form, nTrappedErrNum As Long, _
        sTrappedErrDesc As String, sProcName As String) As OnErrorAction
'---------------------------------------------------------------------
' Call the error handling form pass it the err no. the err desc and the
' form that raised the error.
'---------------------------------------------------------------------
    On Error GoTo MemoryErr
    
    'za 28/09/01
    'here we call the new error handler routine
    MACROFormErrorHandler = MACROErrorHandler(oForm.Name, nTrappedErrNum, sTrappedErrDesc, sProcName, GetApplicationTitle)
    
'    If nTrappedErrNum = 0 Then
'        MACROFormErrorHandler = OnErrorAction.Ignore
'        Exit Function
'    Else
'        Call frmErrHandler.FormRefreshMe(oForm, nTrappedErrNum, sTrappedErrDesc, sProcName)
'        frmErrHandler.Show vbModal
'        MACROFormErrorHandler = frmErrHandler.gOnErrorAction
'        If frmErrHandler.gOnErrorAction = OnErrorAction.QuitMACRO Then
'            oForm.Hide
'        End If
'    End If
   
' SR3685 Trap for out of memory
    
Exit Function
MemoryErr:
    Select Case Err.Number
        Case 7 ' Out of Memory error
            MsgBox "The application has run out of memory and will now be shut down.", vbCritical, "MACRO"
            Call ExitMACRO
            Call MACROEnd
    End Select
    

End Function

'---------------------------------------------------------------------
Public Function MACROCodeErrorHandler(nTrappedErrNum As Long, sTrappedErrDesc As String, _
                            sProcName As String, sModuleName As String) As OnErrorAction
'---------------------------------------------------------------------
' Call the error handling form pass it the err no. the err desc
' there is no form to pass as this one is used in the modules.
'---------------------------------------------------------------------
    On Error GoTo MemoryErr
    
    'za 28/09/01
    'here we call the new error handler routine
    MACROCodeErrorHandler = MACROErrorHandler(sModuleName, nTrappedErrNum, sTrappedErrDesc, sProcName, GetApplicationTitle)
'
'    If nTrappedErrNum = 0 Then
'        MACROCodeErrorHandler = OnErrorAction.Ignore
'        Exit Function
'    Else
'        Call frmErrHandler.CodeRefreshMe(nTrappedErrNum, sTrappedErrDesc, sProcName, sModuleName)
'        frmErrHandler.Show vbModal
'        MACROCodeErrorHandler = frmErrHandler.gOnErrorAction
'        If frmErrHandler.gOnErrorAction = OnErrorAction.QuitMACRO Then
'            Screen.ActiveForm.Hide
'        End If
'    End If
    
' SR3685 Trap for out of memory
Exit Function
MemoryErr:
    Select Case Err.Number
        Case 7 ' Out of Memory error
            MsgBox "The application has run out of memory and will now be shut down.", vbCritical, "MACRO"
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'---------------------------------------------------------------------
Private Function SilentUserLogin(sSecCon As String, ByVal sDatabase As String, _
                                sDefaultHTMLLocation As String) As MACROUser
'---------------------------------------------------------------------
' NCJ 3 Jan 03 - Is this used anywhere in MACRO DM now (version 3.0)????
'---------------------------------------------------------------------
Dim oUser As MACROUser
Dim bLoginSucceeded As Boolean
Dim sMessage As String

    Set oUser = New MACROUser
    bLoginSucceeded = oUser.SilentLogin(sSecCon, sDatabase, sDefaultHTMLLocation, sMessage)
    
    If bLoginSucceeded Then
        Set SilentUserLogin = oUser
    Else
        Set SilentUserLogin = Nothing
    End If

End Function

'----------------------------------------------------------------------------------------'
Public Function GetMacroRegistryKey() As String
'----------------------------------------------------------------------------------------'
' NCJ 24/1/00
' Get the name of the Registry key for the Security database path and other things
' Specify the text for the key elements (ProductName, CompanyName)
' in case Project Properties are inadvertently changed
'   SDM 26/01/00 SR2794 Moved from class User
' NCJ 31/8/00 - Changed registry key for 2.1 (use 2.1 explicitly)
' NCJ 27/4/00 - Changed registry key for 2.2
' TA 31/10/2001 - changes for VTRACK
' REM 01/07/02 - Changed registry key for 3.0
'----------------------------------------------------------------------------------------'

'31/10/2001: VTRACK Changes for separate registry settings
#If VTRACK = 1 Then
    GetMacroRegistryKey = "Software\InferMed Limited\VTRACK\1.0"
#Else
    GetMacroRegistryKey = "Software\InferMed Limited\MACRO\3.0"
#End If

End Function

'----------------------------------------------------------------------------------------'
Public Function GetApplicationTitle() As String
'----------------------------------------------------------------------------------------'
'Return the default title of an app
'ic 17/11/2005 added clinical coding
'----------------------------------------------------------------------------------------'

    Select Case App.Title
    Case "MACRO_SD"
        If LCase$(Command) = "library" Then
            GetApplicationTitle = APPTITLE_LM
        Else
            GetApplicationTitle = APPTITLE_SD
        End If
    Case "MACRO_DM"
         If LCase$(Command) = "review" Then
            GetApplicationTitle = APPTITLE_DR
        Else
            'TA 23/10/2000 changend from Data Management"
            GetApplicationTitle = APPTITLE_DE
        End If
    'REM 14/02/03 - Commented out as there is no longer MACRO Exchange
'    Case "MACRO_EX"
'         GetApplicationTitle = "MACRO Exchange"
    Case "MACRO_SM"
        GetApplicationTitle = APPTITLE_SM
    Case "MacroCreateDataViews"
        GetApplicationTitle = APPTITLE_DV
    Case "MACRO_BD"
        GetApplicationTitle = APPTITLE_BD
    Case "MACRO_QM"
        GetApplicationTitle = APPTITLE_QM
    Case "MACRO_BV"     ' Added NCJ 26 Feb 03
        GetApplicationTitle = APPTITLE_BV
    Case "MACRO_UT"
        GetApplicationTitle = "MACRO Utilities"
    Case "MACRO_CC"
        GetApplicationTitle = APPTITLE_CC
    Case "MACRO_SC"
        GetApplicationTitle = APPTITLE_SC
    Case Else
        'TA 17/1/02: part of VTRACK buglist build 1.0.3 Bug 6
        'else case just returns the real app.title
        GetApplicationTitle = App.Title
    End Select
    
End Function

' NCJ 29 Jan 03 - No longer user this GetPrologSwitches routine
''----------------------------------------------------------------------------------------'
'Public Function GetPrologSwitches(Optional ByVal lTrialId As Long = 0) As String
''----------------------------------------------------------------------------------------'
'' NCJ 18 October 2001
'' Get the Prolog memory settings for Arezzo (in both SD and DE)
'' Read Program Space and Text Space from the registry, and use defaults for the others
'' NCJ 19 Jun 02 - CBB 2.2.15/14 Read ALL prolog switches from registry
''----------------------------------------------------------------------------------------'
'Dim nProgramSpace As Integer
'Dim nTextSpace As Integer
'Dim sPrologSwitches As String
'
'    sPrologSwitches = ""
'    sPrologSwitches = sPrologSwitches & "/P" & GetPrologSwitch(gsPROGRAM_SPACE, glPROGRAM_SPACE)
'    sPrologSwitches = sPrologSwitches & " /T" & GetPrologSwitch(gsTEXT_SPACE, glTEXT_SPACE)
'    sPrologSwitches = sPrologSwitches & " /L" & GetPrologSwitch(gsLOCAL_SPACE, glLOCAL_SPACE)
'    sPrologSwitches = sPrologSwitches & " /B" & GetPrologSwitch(gsBACKTRACK_SPACE, glBACKTRACK_SPACE)
'    sPrologSwitches = sPrologSwitches & " /H" & GetPrologSwitch(gsHEAP_SPACE, glHEAP_SPACE)
'    sPrologSwitches = sPrologSwitches & " /I" & GetPrologSwitch(gsINPUT_SPACE, glINPUT_SPACE)
'    sPrologSwitches = sPrologSwitches & " /O" & GetPrologSwitch(gsOUTPUT_SPACE, glOUTPUT_SPACE)
'
'    GetPrologSwitches = sPrologSwitches
'
'End Function
'
''----------------------------------------------------------------------------------------'
'Private Function GetPrologSwitch(sKeyName As String, lDefaultValue As Long) As Long
''----------------------------------------------------------------------------------------'
'' NCJ 19 Jun 02 - CBB 2.2.15/14
'' Get a Prolog switch value from the registry
'' sKeyName is switch name (as stored in registry)
'' and nDefaultValue is default (minimum) value
''ZA 01/07/2002 - changed all values from integer to long
'' ASH 12/9/2002 - Call now made to settings file for keys
''----------------------------------------------------------------------------------------'
'Dim lValue As Long
'
'    ' Using CInt and Val deals with non-integer/text or other strange values
'    'lValue = CLng(Val(GetFromRegistry(GetMacroRegistryKey, sKeyName)))
'
'    'ASH 11/9/2002 Use new IMEDSettings to get return value/default
'    lValue = CLng(Val(GetMACROSetting(sKeyName, lDefaultValue)))
'
'    ' Returns 0 if no key value exists (or if invalid)
'    ' Also screen out values less than recommended minimum (in case of user error!)
'    If lValue < lDefaultValue Then
'        lValue = lDefaultValue
'    End If
'
'    GetPrologSwitch = lValue
'
'End Function

'----------------------------------------------------------------------------------------'
Public Sub FetchKeywords()
'----------------------------------------------------------------------------------------'
' RJCW 26/11/2001
' Initialises the system Keywords from the static table Keyword
' into a public dictionary object
'----------------------------------------------------------------------------------------'
Dim rsKeywords As ADODB.Recordset
Dim sSQL As String
Dim sKeywordCode As String

    On Error GoTo ErrHandler

    sSQL = "SELECT Keyword FROM Keyword"
        
    Set rsKeywords = New ADODB.Recordset
    ' recordset is FORWARD ONLY & READ ONLY
    rsKeywords.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set gsdicKeywords = New Scripting.Dictionary

    With rsKeywords
        Do Until .EOF = True
            sKeywordCode = UCase(rsKeywords!Keyword)
            gsdicKeywords.Add sKeywordCode, ""
            .MoveNext
        Loop
    End With

    rsKeywords.Close
    Set rsKeywords = Nothing
        
    Exit Sub

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "FetchKeywords", "MainMACROModule")
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
Public Sub UserLogOff(bExit As Boolean)
'---------------------------------------------------------------------
' REM 18/10/02
'---------------------------------------------------------------------
Dim sMsg As String

        
    If Not bExit Then
        sMsg = "Are you sure you wish to finish the current MACRO session" & vbNewLine
        sMsg = sMsg & "and log in as a different user?"
        If DialogQuestion(sMsg) = vbYes Then
            'Call CloseSubject(goStudyDef, gsSubjectToken, True)
            goUser.LogOff
            Call UnloadAllChildForms(False)
            Call Main
        End If
    Else
        goUser.LogOff
        Call UnloadAllChildForms(False)
        MACROEnd
    End If
    
End Sub

'---------------------------------------------------------------------
Public Function DefaultHTMLLocation() As String
'---------------------------------------------------------------------
'return the default HTML folder to pass to user object
'---------------------------------------------------------------------

    DefaultHTMLLocation = App.Path & "\HTML\"
    
End Function

'---------------------------------------------------------------------
Public Function BDCommandLineLoginOK(ByVal sCommandLine As String) As Boolean
'---------------------------------------------------------------------
Dim sBDCLCommandCode As String
Dim sBDCLUserName As String
Dim sBDCLPassword As String
Dim sBDCLDatabaseCode As String
Dim asBDCLCommandParts() As String
Dim vBDCLUserDatabase As Variant
Dim bBDCLDBValid As Boolean
Dim sMessage As String

    On Error GoTo ErrHandler
    
    Set goUser = New MACROUser

    asBDCLCommandParts = Split(sCommandLine, "/")
    sBDCLCommandCode = UCase(asBDCLCommandParts(1))
    Select Case sBDCLCommandCode
    Case "BI"
        If UBound(asBDCLCommandParts) = 5 Then
            sBDCLUserName = asBDCLCommandParts(2)
            sBDCLPassword = asBDCLCommandParts(3)
            sBDCLDatabaseCode = asBDCLCommandParts(4)
            gsBDCLPathAndNameOfFile = asBDCLCommandParts(5)
        Else
            Call CreateBDCLLogFile(sCommandLine, "Wrong number of Batch Import parameters.")
            BDCommandLineLoginOK = False
            Exit Function
        End If
    Case "BU"
        If UBound(asBDCLCommandParts) = 4 Then
            sBDCLUserName = asBDCLCommandParts(2)
            sBDCLPassword = asBDCLCommandParts(3)
            sBDCLDatabaseCode = asBDCLCommandParts(4)
        Else
            Call CreateBDCLLogFile(sCommandLine, "Wrong number of Batch Upload parameters.")
            BDCommandLineLoginOK = False
            Exit Function
        End If
    End Select
    
    'Check that the none of the commandline arguments are empty
    If sBDCLUserName = "" Or sBDCLPassword = "" Or sBDCLDatabaseCode = "" Then
        Call CreateBDCLLogFile(sCommandLine, "Command line contains empty parameters.")
        BDCommandLineLoginOK = False
        Exit Function
    End If
    
    'Validate the UserName and Password
    'Mo***
    If goUser.Login(gsSecCon, sBDCLUserName, sBDCLPassword, DefaultHTMLLocation, GetApplicationTitle, sMessage, False) <> LoginResult.Success Then
        Call CreateBDCLLogFile(sCommandLine, "Could not Login using supplied UserName and Password.")
        BDCommandLineLoginOK = False
        Exit Function
    End If
    
    'Check that sBDCLDatabaseCode is a valid database for sBDCLUserName
    bBDCLDBValid = False
    For Each vBDCLUserDatabase In goUser.UserDatabases
        If vBDCLUserDatabase = sBDCLDatabaseCode Then
            bBDCLDBValid = True
        End If
    Next
    If Not bBDCLDBValid Then
        Call CreateBDCLLogFile(sCommandLine, "Supplied Database is not valid for the supplied UserName.")
        BDCommandLineLoginOK = False
        Exit Function
    End If
    
    'If its a Batch Import call Check that the file to be imported exists
    If sBDCLCommandCode = "BI" Then
        If Not FileExists(gsBDCLPathAndNameOfFile) Then
            Call CreateBDCLLogFile(sCommandLine, "Supplied Batch Import File does not exist.")
            BDCommandLineLoginOK = False
            Exit Function
        End If
    End If
    
    'Perform DB Initialisation activities
    'Mo***
    Call goUser.SetCurrentDatabase(sBDCLUserName, sBDCLDatabaseCode, DefaultHTMLLocation, True, True, sMessage)
    
    'Everything is ok with the command arguments
    BDCommandLineLoginOK = True
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "BDCommandLineLoginOK", "MainMACROModule")
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
Public Sub CreateBDCLLogFile(ByVal sCommandLine As String, ByVal sLogMessage As String)
'---------------------------------------------------------------------
Dim sBDCLLogFile As String
Dim nBDCLFileNumber As Integer

    On Error GoTo ErrHandler

    sBDCLLogFile = App.Path & "\BDCL" & Format(Now, "yyyymmddhhmmss") & ".log"
    
     'Open the Log file and write to it
    nBDCLFileNumber = FreeFile
    Open sBDCLLogFile For Output As #nBDCLFileNumber
    Print #nBDCLFileNumber, "MACRO Batch Data Entry, Command Line Login Failure Log"
    Print #nBDCLFileNumber, "------------------------------------------------------"
    Print #nBDCLFileNumber, "The syntax for a Batch Import Commamd Line call is:-"
    Print #nBDCLFileNumber, "     /BI/UserName/Password/DatabaseCode/PathAndNameOfFile"
    Print #nBDCLFileNumber, "The syntax for a Batch Upload Commamd Line call is:-"
    Print #nBDCLFileNumber, "     /BU/UserName/Password/DatabaseCode"
    Print #nBDCLFileNumber, "------------------------------------------------------"
    Print #nBDCLFileNumber, "Date/Time: " & Now
    Print #nBDCLFileNumber, "Entered Command Line: " & sCommandLine
    Print #nBDCLFileNumber, "Error: " & sLogMessage
    Print #nBDCLFileNumber, "------------------------------------------------------"
    'Close the Log file
    Close #nBDCLFileNumber

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CreateBDCLLogFile", "MainMACROModule")
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
Public Function ExcludeUserRDE(sUserName As String) As Boolean
'---------------------------------------------------------------------
'REM 24/01/03
'Used to check if user 'rde' details should not be written to message table
'---------------------------------------------------------------------
    
    ExcludeUserRDE = (sUserName = "rde")

End Function

'---------------------------------------------------------------------
Public Function ShowAbout() As Boolean
'---------------------------------------------------------------------

'---------------------------------------------------------------------
    
    

End Function


