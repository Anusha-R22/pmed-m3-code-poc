Attribute VB_Name = "basEnumerations"
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998-2006. All Rights Reserved
'   File:       basEnumerations.bas
'   Author:     Andrew Newbigging, April 1998
'   Purpose:    Common enumerations used throughout MACRO
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
' 1998-1999 - Various revisions
'   NCJ 3 Feb 00 - eDateFormatType for use with new CLM/PLM date string routines
'   Mo  25/4/00 Additions made to the ExchangeMessageType enumeration
'               changes made to MessageReceived enumeration
'               changes made to Status enumeration OKWarning added, Frozen, Locked, Pending removed
'   Mo 17/5/00  New enumerations for the MIMessage table added:-
'               TypeOfInstallation,MIMType,MIMScope,MIMHistoryStatus,DiscrepancyStatus,SDVStatus
'   TA June 00  New enums for Direction and Scedule Cell for CRFDataEntry collections
'   Mo Morris   LabDefinitionServerToSite and LabDefinitionSiteToServer added to Enum ExchangeMessageType
'   Mo Morris   new enumeration DataViewState created
' DPH 08/04/2002 - New additions to the MessageReceived ,
' NCJ 16 Oct 02 - MIMsg enumerations no longer used
' ic 30/10/2002 added 'PopUp' to DataType enum *since removed
' ic 03/12/2002 added 2 setting consts
' TA 13/12/2002 copied in enums from modDMEnum
' NCJ 29 Jan 03 - Moved the Prolog Memory constants here from basMainMACROModule
' ic 13/03/2003 added menu id constants
' NCJ 14 May 03 - Changed max. AREZZO expression length from 2000 to 4000 (Bug 1305)
' Mo 24/11/2004 - Bug 2413, new enumeration for Create Data Views
' DPH 21/01/2005 - Added Pdu messages to ExchangeMessageType enumeration (PDU2300)
' ic 10/06/2005 added clinical coding
' NCJ 7 Dec 05 - eDateFormatType enumeration changed
' NCJ 10 May 06 - New eSDAccessMode for Multi User SD
' NCJ 25 Sept 06 - Corrected GetAccessModeString for non-MUSD
' Mo 2/7/2007   APPTITLE_DD added
'               APPTITLE_UT added
'               APPTITLE_DU added
'--------------------------------------------------------------------------------

Option Explicit

' NCJ 26/8/99
' The prefixes used for Arezzo tasks
' Changed NCJ 7 Sept 99
Public Const gsVisitPlanPrefix = "Visit "
Public Const gsCRFPlanPrefix = "Form "
Public Const gsCRFActionPrefix = "Data_Entry_"

'TA
Public Const SETTING_VIEW_FUNCTION_KEYS = "viewfunctionkeys"
Public Const SETTING_VIEW_SYMBOLS = "viewsymbols"
Public Const SETTING_SPLIT_SCREEN = "splitscreen"
Public Const SETTING_LOCAL_FORMAT = "localformat"
Public Const SETTING_SERVER_TIME = "servertime"
Public Const SETTING_LAST_USED_STUDY = "lastusedstudy"
Public Const SETTING_LAST_USED_EFORM = "lastusedeform"
'ic
Public Const SETTING_SAME_EFORM = "sameeform"
Public Const SETTING_PAGE_LENGTH = "pagelength"
' RS
Public Const SETTING_LOCAL_DATE_FORMAT = "localdateformat"
Public Const SETTING_LOCAL_TIME_FORMAT = "localtimeformat"
' MLM 23/04/05
Public Const SETTING_BANDWIDTH As String = "bandwidth"

' NCJ 29 Jan 03 - Moved the Prolog Memory constants here from basMainMACROModule
'ZA 26/06/2002 - added Prolog Memory keys and default minimum and max values
'Prolog memory keys
Public Const gsPROGRAM_SPACE = "PrologProgramSpace"
Public Const gsTEXT_SPACE = "PrologTextSpace"
Public Const gsLOCAL_SPACE = "PrologLocalSpace"
Public Const gsBACKTRACK_SPACE = "PrologBackTrackSpace"
Public Const gsHEAP_SPACE = "PrologHeapSpace"
Public Const gsINPUT_SPACE = "PrologInputSpace"
Public Const gsOUTPUT_SPACE = "PrologOutputSpace"

'default (minimum) values for Prolog memory settings
Public Const glPROGRAM_SPACE = 8000
Public Const glTEXT_SPACE = 2000
Public Const glLOCAL_SPACE = 256
Public Const glBACKTRACK_SPACE = 256
Public Const glHEAP_SPACE = 512
Public Const glINPUT_SPACE = 400
Public Const glOUTPUT_SPACE = 400

'maximum values for Prolog memory settings
Public Const glMAX_PROGRAM_SPACE = 200000
Public Const glMAX_TEXT_SPACE = 100000
Public Const glMAX_LOCAL_SPACE = 10000
Public Const glMAX_BACKTRACK_SPACE = 10000
Public Const glMAX_HEAP_SPACE = 20000
Public Const glMAX_INPUT_SPACE = 10000
Public Const glMAX_OUTPUT_SPACE = 10000

'average/mid range values for Prolog memory settings
Public Const glMID_PROGRAM_SPACE = 32000
Public Const glMID_TEXT_SPACE = 5000
Public Const glMID_LOCAL_SPACE = 1000
Public Const glMID_BACKTRACK_SPACE = 1000
Public Const glMID_HEAP_SPACE = 2000
Public Const glMID_INPUT_SPACE = 1000
Public Const glMID_OUTPUT_SPACE = 1000

'REM 14/02/03 - Application Titles
' NCJ 26 Feb 03 - Added Batch Validation
' ic 17/11/2005 added clinical coding
Public Const APPTITLE_DE = "MACRO Data Entry"
Public Const APPTITLE_DR = "MACRO Data Review"
Public Const APPTITLE_SM = "MACRO System Management"
Public Const APPTITLE_SD = "MACRO Study Definition"
Public Const APPTITLE_BD = "MACRO Batch Data Entry"
Public Const APPTITLE_QM = "MACRO Query Module"
Public Const APPTITLE_LM = "MACRO Library Management"
Public Const APPTITLE_DV = "MACRO Create Data Views"
Public Const APPTITLE_BV = "MACRO Batch Validation"
Public Const APPTITLE_CC = "MACRO Clinical Coding Console"
Public Const APPTITLE_SC = "MACRO Study Copy Tool"
'Mo 2/7/2007
Public Const APPTITLE_DD = "MACRO Double Data Entry"
Public Const APPTITLE_UT = "MACRO Utilities"
Public Const APPTITLE_DU = "MACRO Diagnostic Utility"

' NCJ 23 Jan 03 - Max length of an AREZZO expression
' NCJ 14 May 03 - Changed from 2000 to 4000 (Bug 1305)
Public Const glMAX_AREZZO_EXPR_LEN As Long = 4000

'ic 13/03/2003
Public Const gsCREATE_NEW_SUBJECT_MENUID = "CNS"
Public Const gsVIEW_SUBJECT_LIST_MENUID = "VSL"
Public Const gsVIEW_RAISED_DISCREPANCIES_MENUID = "VRAD"
Public Const gsVIEW_RESPONDED_DISCREPANCIES_MENUID = "VRED"
Public Const gsVIEW_OC_DISCREPANCIES_MENUID = "VOCD"
Public Const gsVIEW_PLANNED_SDV_MARKS_MENUID = "VPSM"
Public Const gsLAB_AND_NR_MENUID = "LANR"
Public Const gsVIEW_CHANGES_SINCE_LAST_SESSION_MENUID = "VCSLS"
Public Const gsTEMPLATES_MENUID = "T"
Public Const gsREGISTER_SUBJECT_MENUID = "RS"
Public Const gsVIEW_LOCK_FREEZE_HISTORY_MENUID = "VLFH"
Public Const gsDB_LOCK_ADMIN_MENUID = "DLA"
Public Const gsREMOTE_TIME_SYNCH_MENUID = "RTS"
Public Const gsCHANGE_PASSWORD_MENUID = "CP"
'TA 18/03/2003
Public Const gsTRANSFER_DATA_MENUID = "TD"

'ic 13/09/2005 maximum length allotted in the db for a coding path (codingdetails column)
Public Const gnMAX_CODE_LENGTH As Integer = 4000

'NCJ 10 May 06 - New enumeration for Multi-user SD
Public Enum eSDAccessMode
    sdReadOnly = 0
    sdLayoutOnly = 1
    sdReadWrite = 2
    sdFullControl = 3
End Enum
    

' NCJ 23/11/00 Registration status
Public Enum eRegStatus
    NotReady = 0
    Ready = 1
    Ineligible = 2
    Failed = 3
    Registered = 4
End Enum

' NCJ 22/11/00
' Type of Registration/Randomisation server
Public Enum eRRServerType
    RRNone = 0
    RRLocal = 1
    RRTrialOffice = 2
    RRRemote = 3
End Enum

' NCJ 22/11/00
' Type of Registration result
Public Enum eRegResult
    RegOK = 0
    RegAlreadyRegistered = 1
    RegNotUnique = 2
    RegMissingInfo = 3
    RegError = 4
End Enum

' NCJ 9 Mar 00
' Types of MACRO window
' (used for keeping track of user's coming and goings)
Public Enum eMACROWindow
    None = 0
    Schedule = 1
    SubjectBrowser = 2
    MonitorBrowser = 3
'extra ones for modeless mimessage form
    SubjectMIMEssage = 4
    MonitorMIMessage = 5
End Enum

' NCJ 16 Nov 99
' Case conversions for data items
Public Enum eTextCase
    Leave = 0
    Upper = 1
    Lower = 2
End Enum

' NCJ 3 Feb 00
' NCJ 7 Dec 05 - We have more types now (NB same value also appear in eFormElementRO.cls)
Public Enum eDateFormatType
    InvalidFormat = 0
    dftDMY = 1
    dftMDY = 2
    dftYMD = 3
    dftDMYT = 4
    dftMDYT = 5
    dftYMDT = 6
    dftMY = 7
    dftYM = 8
    dftY = 9
    dftT = 10
End Enum

' NCJ 3/9/99
' Arezzo data error results
Public Enum eArezzoDataError
    OK = 0
    Warning = 1
    TypeError = -1
    RangeError = -2
    ValidationError = -3
End Enum

' NCJ 3 Sep 1999
' What to do with a data error
Public Enum ValidationAction
    Reject = 0
    Warn = 1
    Inform = 2
End Enum

'   ATN 26/2/99
'   New enumeration for the security to be used at a particular site
Public Enum SecurityMode
    UsernamePassword = 0
    NTIntegrated = 1
    NTSeparatePassword = 2
End Enum

'   ATN 18/2/99
'   New enumeration for the direction of a message
Public Enum MessageDirection
    MessageOut = 0
    MessageIn = 1
End Enum


'   ATN 18/2/99
'   New enumeration for the status of a trial
Public Enum eTrialStatus
    InPreparation = 1
    TrialOpen = 2
    ClosedToRecruitment = 3
    ClosedToFollowUp = 4
    Suspended = 5
End Enum

'   ATN 18/2/99
'   New enumeration for the type of database
Public Enum MACRODatabaseType
    Access = 0
    sqlserver = 1
    SQLServer70 = 2
    Oracle80 = 3
End Enum

'   New enumeration for message.  Indicates the type of message
Public Enum ExchangeMessageType
    NewTrial = 0
    InPreparation = 1
    TrialOpen = 2
    ClosedRecruitment = 3
    ClosedFollowUp = 4
    TrialSuspended = 5
    NewVersion = 8
    PatientData = 10
    Mail = 11
    TrialSubjectLockStatus = 16
    VisitInstanceLockStatus = 17
    CRFPageInstanceLockStatus = 18
    DataItemLockStatus = 19
    TrialSubjectUnLock = 20
    VisitInstanceUnLock = 21
    CRFPageInstanceUnLock = 22
    LabDefinitionServerToSite = 30
    LabDefinitionSiteToServer = 31
    User = 32
    UserRole = 33
    PasswordChange = 34
    Role = 35
    SystemLog = 36
    UserLog = 37
    RestoreUserRole = 38
    PasswordPolicy = 40
    PatientDataSent = 41
    PduInstruction = 50
    PduPackage = 51
End Enum

Public Enum Status
    Requested = -10
    CancelledByUser = -20
    NotApplicable = -8
    Unobtainable = -5
    Success = 0
    Missing = 10
    Inform = 20
    OKWarning = 25
    Warning = 30
    InvalidData = 40
End Enum

Public Enum Changed
    NoChange = 0
    Changed = 1
    Imported = 2
End Enum

Public Enum DataType
    Text = 0
    Category = 1
    IntegerData = 2
    Real = 3
    Date = 4
    Multimedia = 5
    LabTest = 6
    'ic 10/06/2005 clinical coding: added thesaurus type. see also eDataType below
    Thesaurus = 8
End Enum

Public Enum ValidationType
    Mandatory = 0
    Warning = 1
    Inform = 2
    LabNormal = 3
    LabFeasible = 4
    LabAbsolute = 5
    Skip = 9
End Enum

Public Enum SDDImportType
    MACRO = 0
    COMPACT = 1
End Enum

' PN 26/07/99 message received enumerators
' Changed Mo Morris 26/4/00
' DPH 08/04/2002 - Added Error / Skipped
' DPH 04/09/2002 - Added Superceeded
' REM 03/12/02 - added PendingOverRule - for system data transfer
Public Enum MessageReceived
    NotYetReceived = 0
    Received = 1
    Error = 2
    Locked = 3
    Skipped = 4
    Superceeded = 5
    PendingOverRule = 6
End Enum

Public Enum OnErrorAction
    QuitMACRO = 0
    Retry = 1
    Ignore = 2
End Enum

Public Enum ReportType
'    Crystal = 0        ATN 16/12/99    Not used.
    TabFile = 1
    CSVFile = 2
    Excel = 3
    HTML = 4            'ATN 16/12/99    Formatted web report
    HTMLPre = 5         'ATN 16/12/99   Unformatted web report
    SPSS = 6            'ATN 16/12/99   SPSS
    SAS = 7             'ATN 16/12/99   SAS
End Enum

'SDM 07/12/99
Public Enum LockStatus
    lsUnlocked = 0
    lsPending = 3
    lsLocked = 5
    lsFrozen = 6
End Enum

Public Enum TypeOfInstallation
    Server = 0
    RemoteSite = 1
End Enum

' NCJ 16 Oct 02 - MIMsg enumerations no longer used
' (We now use the ones in the MACROMIMsgBS30 component instead)

Public Enum ScheduleCell
    scEmpty = 0
    scSingle = 1
    scRepeating = 2
End Enum

Public Enum Direction
    dirNext = 1
    dirPrevious = 2
End Enum

Public Enum DataViewState
    NotCreated = 0
    Created = 1
End Enum

Public Enum DataViewOption
    NotRequired = 0
    Required = 1
End Enum

'Mo 24/11/2004 - Bug 2413, new enumeration for Create Data Views
Public Enum DataViewCategoryOptions
    Codes = 0
    Values = 1
    TypedCodes = 2
End Enum

Public Const gnZERO = 0
Public Const glMINUS_ONE = -1
Public Const gsEMPTY_STRING = ""


'REM 11/12/01
Public Enum EditQGroup
    Cancel = 0
    Edit = 1
    Delete = 2
End Enum

'REM 12/07/02
Public Enum eNodeTag
    CRFpageTag = 0
    QGroupTag = 1
    QuestionTag = 2
    UnUsedQuestandRQGTag = 3
End Enum

'ZA 19/08/2002
Public Enum eMACROOnly
    NotMACROOnly = 0
    MACROOnly = 1
End Enum

'ZA 20/08/2002 - Default RFC for question item on or off
Public Enum eRFCDefault
    RFCDefaultOff = 0
    RFCDefaultOn = 1
End Enum

'ZA 20/08/2002
Public Enum eArezzoUpdateStatus
    auNotRequired = 0
    auRequired = 1
End Enum

'ZA 21/8/2002 - Form and Visit date validation
Public Enum eElementUse
    User = 0
    EFormVisitDate = 1
End Enum

Public Enum eEFormUse
    User = 0
    VisitEForm = 1
End Enum

'ZA 06/09/2002 - hide/show the status icon for a question
Public Enum eStatusFlag
    Hide = 0
    Show = 1
End Enum
'ZA 06/09/2002 - auto numbering on/off
Public Enum eAutoNumbering
    NumberingOff = 0
    NumberingOn = 1
End Enum

'ZA 30/09/2002 - Reason types
Public Enum eReasonType
    ReasonForChange = 0
    ReasonForOverrule = 1
End Enum

'REM 20/11/02 - Type of UserDetails for System Data transfer
Public Enum eUserDetails
    udNewUser = 0
    udEditUser = 1
    udDisableUser = 2
End Enum

'REM 20/11/02 - User Status for System Data transfer
Public Enum eUserStatus
    usDisabled = 0
    usEnabled = 1
End Enum

'REM 20/11/02 - User Lock status for System Data transfer
Public Enum eUserLock
    ulUnlocked = 0
    ulLockout = 1
End Enum

'REM 22/11/02 - UserRole Message, indicates whether its a delete or add
Public Enum eUserRole
    urDelete = 0
    urAdd = 1
End Enum

'REM 22/11/02 - System data transfer Message parameters (reading them out of string)
Public Enum eSDTMessage
    mMessageId = 0
    mTrialSite = 1
    mClinicalTrialId = 2
    mMessageType = 3
    mUserName = 4
    mMessageTimeStamp = 5
    mMessageTimeStamp_TZ = 6
    mMessageBody = 7
    mMessageParameters = 8
    mMessageDirection = 9
End Enum

'REM 09/12/02 - System Data transfer, message table fields
Public Enum eSDTMsgFields
    fTrialSite = 0
    fClinicalTrialId = 1
    fMessageType = 2
    fMessageTimeStamp = 3
    fUserName = 4
    fMessageBody = 5
    fMessageParameters = 6
    fMessageDirection = 8
    fMessageId = 9
    fMessageTimeStamp_TZ = 11
End Enum

'REM 06/12/02 - for forgotten password data transfer
Public Enum eDTForgottenPassword
    pSuccess = 0
    pError = 1
    pNoPassword = 2
    pNoDatabases = 3
    pIncorrectPassword = 4
End Enum


'The message types of all system messages
Public Const gsSYSTEM_MESSAGE_TYPES As String = "32,33,34,35,36,37,40"

Public Const gsCORRUPT_DATA As String = "System Management data corrupted during transfer"

'constants for System DataTransfer
Public Const gsERROR As String = "Error"
Public Const gsMSGSEPARATOR As String = "<p>"
Public Const gsFIELDSEPARATOR As String = "<br>"
Public Const gsPARAMSEPARATOR As String = "*"
Public Const gsSEPARATOR As String = "|"
Public Const gsCHECKSUM_SEPARATOR = "<chk>"
Public Const gsERRMSG_SEPARATOR = "<msg>"

Public Const gsSERVER As String = "server"

Public Const gsEND_OF_MESSAGES = "."

'ash 11/12/2002 Moved from Exchange
' display mode enums
Public Enum eDisplayType
    DisplaySitesByTrial = 1
    DisplayTrialsBySite = 2
    DisplaySitesByLab = 3
    DisplayLabsBySite = 4
    DisplaySitesByUser = 5
    DisplayUsersBySite = 6
End Enum



'----------------------------------------------------
' File: modDMEnum.bas
' Author: Toby Aldridge
' Copyright: InferMed, June 2001, All Rights Reserved
' Enumerations etc. for MACRO 2.2
'----------------------------------------------------
' NB This is mostly a COPY of what's in modMACRODM (not in MACRO_DM project)
' and it NEEDS SORTING OUT (NCJ 2 Oct 01)
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 2 Oct 01 - Added file header
'               GetValidationTypeString
'----------------------------------------------------



' Question data types
Public Enum eDataType
    Text = 0
    Category = 1
    IntegerNumber = 2
    Real = 3
    DateTime = 4
    Multimedia = 5
    LabTest = 6
    'ic 10/06/2005 clinical coding: added thesaurus type. see also DataType above
    Thesaurus = 8
End Enum

' Status of Responses, eFormInstances and VisitInstances
Public Enum eStatus
    CancelledByUser = -20
    Requested = -10
    NotApplicable = -8
    Unobtainable = -5
    Success = 0
    Missing = 10
    Inform = 20
    OKWarning = 25
    Warning = 30
    InvalidData = 40
End Enum

' Lock status
Public Enum eLockStatus
    lsUnlocked = 0
    lsPending = 3
    lsLocked = 5
    lsFrozen = 6
End Enum

' In MACRO 0 is used to mean unspecified date
Public Enum eMACRODate
    mdUnspecified = 0
End Enum

' Possible results of validating a response value
' NB The Arezzo results MUST have the values 1,-1,-2 and -3
Public Enum eResponseValidation
    ValueOK = 0
    ' NB The Arezzo values must be exactly as follows
    ArezzoWarning = 1
    ArezzoTypeError = -1
    ArezzoRangeError = -2
    ArezzoValidationError = -3
    NotANumber = 2
    NotAnInteger = 3
    NumberTooBig = 4
    NumberTooSmall = 5
    NumberNotPositive = 6
    TextWrongFormat = 7
    NotADateTime = 8
    ValidationReject = 9
    ValueUnchanged = 11
End Enum

' Type of validation result for questions
' NCJ 9 Jul 01
Public Enum eValidationType
    Reject = 0
    Warn = 1
    Inform = 2
End Enum

'ic 18/07/2005 added clinical coding
Public Enum eCodingStatus
    csEmpty = 0
    csNotCoded = 1
    csCoded = 2
    csPendingNewCode = 3
    csAutoEncoded = 4
    csValidated = 5
    csDoNotCode = 6
End Enum

'propery gets to represent constants
Public Property Get GENDER_MALE()
    GENDER_MALE = "M"
End Property

Public Property Get GENDER_FEMALE()
    GENDER_FEMALE = "F"
End Property

Public Property Get GENDER_NONE()
    GENDER_NONE = ""
End Property


'----------------------------------------------------
Public Function GetCodingStatusString(nStatus As Integer)
'----------------------------------------------------
' ic 26/07/2005
' function returns the coding status string for display purposes
'----------------------------------------------------

    Select Case nStatus
    Case eCodingStatus.csAutoEncoded: GetCodingStatusString = "Auto encoded"
    Case eCodingStatus.csCoded: GetCodingStatusString = "Coded"
    Case eCodingStatus.csDoNotCode: GetCodingStatusString = "Do not code"
    Case eCodingStatus.csNotCoded: GetCodingStatusString = "Not coded"
    Case eCodingStatus.csPendingNewCode: GetCodingStatusString = "Pending new code"
    Case eCodingStatus.csValidated: GetCodingStatusString = "Validated"
    Case Else: GetCodingStatusString = ""
    End Select
    
End Function

'----------------------------------------------------
Public Function GetDataTypeString(nDataType As Integer) As String
'----------------------------------------------------
' Data type as a string (for display purposes)
' ic 10/06/2005 added clinical coding
'----------------------------------------------------

    Select Case nDataType
    Case eDataType.Text
        GetDataTypeString = "Text"
    Case eDataType.Category
        GetDataTypeString = "Category"
    Case eDataType.IntegerNumber
        GetDataTypeString = "Integer number"
    Case eDataType.Real
        GetDataTypeString = "Real number"
    Case eDataType.DateTime
        GetDataTypeString = "Date/Time"
    Case eDataType.Multimedia
        GetDataTypeString = "Multimedia"
    Case eDataType.LabTest
        GetDataTypeString = "Lab Test"
    'ic 10/06/2005 clinical coding: added check for thesaurus datatype
    Case eDataType.Thesaurus
        GetDataTypeString = "Thesaurus"
    Case Else
        GetDataTypeString = "UNKNOWN"
    End Select
    
End Function

'----------------------------------------------------
Public Function GetStatusString(nStatus As Integer) As String
'----------------------------------------------------
' Status of a response, eform instance, visit instance
'----------------------------------------------------

    Select Case nStatus
    Case eStatus.CancelledByUser
        GetStatusString = "Cancelled"
    Case eStatus.Requested
        GetStatusString = "Requested"
    Case eStatus.NotApplicable
        GetStatusString = "Not Applicable"
    Case eStatus.Unobtainable
        GetStatusString = "Unobtainable"
    Case eStatus.Success
        GetStatusString = "Success"
    Case eStatus.Missing
        GetStatusString = "Missing"
    Case eStatus.Inform
        GetStatusString = "Inform"
    Case eStatus.OKWarning
        GetStatusString = "OK Warning"
    Case eStatus.Warning
        GetStatusString = "Warning"
    Case eStatus.InvalidData
        GetStatusString = "Invalid"
    Case Else
        GetStatusString = "UNKNOWN"
    End Select
    
End Function

'----------------------------------------------------
Public Function GetLockStatusString(ByVal nLockStatus As Integer) As String
'----------------------------------------------------
' Lock status as string
'----------------------------------------------------

    Select Case nLockStatus
    Case eLockStatus.lsFrozen
        GetLockStatusString = "Frozen"
    Case eLockStatus.lsLocked
        GetLockStatusString = "Locked"
    Case eLockStatus.lsPending
        GetLockStatusString = "Pending"
    Case eLockStatus.lsUnlocked
        GetLockStatusString = "Unlocked"
    Case Else
        GetLockStatusString = "UNKNOWN"
    End Select
    
End Function

'----------------------------------------------------
Public Function GetControlTypeString(nControlType As Integer) As String
'----------------------------------------------------
' EForm Element control type as a string (for display purposes)
'----------------------------------------------------

    Select Case nControlType
    Case 1
        GetControlTypeString = "Text box"
    Case 2
        GetControlTypeString = "Option buttons"
    Case 4
        GetControlTypeString = "Popup list"
    Case 8
        GetControlTypeString = "Calendar"
    Case 16385
        GetControlTypeString = "Line"
    Case 16386
        GetControlTypeString = "Text comment"
    Case 16388
        GetControlTypeString = "Picture"
    Case Else
        GetControlTypeString = CStr(nControlType)
    End Select

End Function

'----------------------------------------------------
Public Function GetResponseErrorString(lResult As Long) As String
'----------------------------------------------------
' Translation of eResponseValidation value into string
'----------------------------------------------------
Dim sErrorMessage As String

    Select Case lResult
    Case eResponseValidation.NotADateTime
        sErrorMessage = "This is not a valid date/time value"
    Case eResponseValidation.NotAnInteger
        sErrorMessage = "This is not a valid integer value"
    Case eResponseValidation.NotANumber
        sErrorMessage = "This is not a valid number value"
    Case eResponseValidation.TextWrongFormat, _
            eResponseValidation.ArezzoTypeError
        sErrorMessage = "This is not in the correct format"
    Case eResponseValidation.NumberNotPositive
        sErrorMessage = "This is not a positive number"
    Case eResponseValidation.NumberTooBig
        sErrorMessage = "This is bigger than the value allowed for this question"
    Case eResponseValidation.NumberTooSmall
        sErrorMessage = "This is smaller than the value allowed for this question"
    Case eResponseValidation.ArezzoRangeError
        sErrorMessage = "This is not one of the allowed values for this question"
    Case Else
        sErrorMessage = ""
    End Select
    
    GetResponseErrorString = sErrorMessage
    
End Function

'----------------------------------------------------------------------
Public Function GetValidationTypeString(nValidationType As Integer) As String
'----------------------------------------------------------------------
' Convert validation type to a string
'----------------------------------------------------------------------

    Select Case nValidationType
    Case eValidationType.Reject
        GetValidationTypeString = "Reject if"
    Case eValidationType.Warn
        GetValidationTypeString = "Warn if"
    Case eValidationType.Inform
        GetValidationTypeString = "Inform if"
    Case Else
        ' Catchall - should never happen?
        GetValidationTypeString = "Validation"
    End Select

End Function

'----------------------------------------------------------------------
Public Function GetAccessModeString(enSDAccessMode As eSDAccessMode, Optional bMUSD As Boolean = False)
'----------------------------------------------------------------------
' NCJ 10 May 06 - Convert access mode into a string
' NCJ 14 Sept 06 - Allow "Update" or "Read write"; added bMUSD (for Multi User SD)
' NCJ 25 Sept 06 - Corrected for non-MUSD
'----------------------------------------------------------------------

    Select Case enSDAccessMode
    Case eSDAccessMode.sdReadOnly
        GetAccessModeString = "Read only"
    Case eSDAccessMode.sdLayoutOnly
        GetAccessModeString = "Layout only"
    Case eSDAccessMode.sdReadWrite
        If Not bMUSD Then
            ' Stick to plain old "Update" for non-MUSD
            GetAccessModeString = "Update"
        Else
            GetAccessModeString = "Read write"
        End If
    Case eSDAccessMode.sdFullControl
        If Not bMUSD Then
            ' Stick to plain old "Update" for non-MUSD
            GetAccessModeString = "Update"
        Else
            GetAccessModeString = "Full control"
        End If
    End Select


End Function
