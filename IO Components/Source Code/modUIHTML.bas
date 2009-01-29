Attribute VB_Name = "modUIHTML"
'----------------------------------------------------------------------------------------'
'   File:       modUIHTML.bas
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     i curtis 10/2002
'   Purpose:    functions returning html versions of MACRO pages (GENERAL)
'----------------------------------------------------------------------------------------'
'   revisions:
'   ic 01/10/2002 - initial development
'   DPH Oct/Nov 2002 - Added RQG functionality
'   DPH 18/02/2003 - Added GetLocalFormatDate functionality
'   ic 20/06/2003 added vbLF to replaced chars
'   ic 27/08/2003 moved code from modUIHTMLApplication, clsWWW
'   ic 16/09/2003 added ServeriseValue() function
'   ic 30/09/2003 dont write a blank gif to the page
'   ic 29/06/2004 added error handling
'   ic 15/07/2004 added LockSubjectIfNeeded() function
'   ic 21/06/2005 issue 2592, create system messages in ChangePassword()
'   ic 04/07/2005 issue 2464, added RtnShowSDVScheduleMenuFlag() function
'   NCJ 8 Dec 05 - New eDateTimeType enumeration values
'   DPH 28/02/2007 - bug 2882. Pass User object into StudiesSitesWhereSQL in RtnMIMsgStatusCount
'----------------------------------------------------------------------------------------'

Public Enum eInterface
    iwww = 0
    iWindows = 1
End Enum

Public Enum eWWWErrorType
    ePermission = 0
    eConfiguration = 1
    eInternal = 2
    eEForm = 3
End Enum

'ic 30/10/2002 this is like the enum in basEnumerations, but has a popup defined
Public Enum WWWDataType
    Text = 0
    Category = 1
    IntegerData = 2
    Real = 3
    Date = 4
    Multimedia = 5
    LabTest = 6
    PopUp = 7
End Enum

Public Enum ControlType
    TextBox = 1
    OptionButtons = 2
    PopUp = 4
    Calendar = 8
    RichTextBox = 16
    Attachment = 32
    PushButtons = 258
    Line = 16385
    Comment = 16386
    Picture = 16388
    Hotlink = 16390
End Enum

Public Const gsDELIMITER1 As String = "`"
Public Const gsDELIMITER2 As String = "|"
Public Const gsDELIMITER3 As String = "~"

Public Const NO_CTC_GRADE = -1

Public Enum NormalRangeLNorH
    nrNotfound = 0
    nrLow = 1
    nrNormal = 2
    nrHigh = 3
    'nrsImpossible = 4
End Enum

Public Enum eWWWIdentifier
    idStudyId = 0
    idSite = 1
    idSubject = 2
    idVisitId = 3
    idVisitCycle = 4
    idVisitTaskId = 5
    idEformId = 6
    idEformCycle = 7
    idEformTaskId = 8
    idResponseId = 9
    idResponseCycle = 10
    idResponseTaskId = 11
End Enum

Public Enum eOpenSubjectError
    UnspecifiedError = 0
    NoPermission = -1
    SubjectLocked = -2
    StudyLocked = -3
End Enum

'greatest number of records that can be retrieved from db
Public Const gnMAXWWWRECORDS As Integer = 5000
Public Const gnMAXWINRECORDS As Integer = 10000

'greatest number of records that can be displayed on a page
Public Const gnMAXWWWRECORDSPERPAGE = 1000
Public Const gnMAXWINRECORDSPERPAGE = 10000

' Alphabet and numbers - ic 18/06/2004
Private Const msALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Private Const msNUMBERS = "0123456789"
Private Const msSTANDARD_DOT = "."
Private Const msSTANDARD_COMMA = ","
Private Const msSPACE = " "
Private Const msUNDERSCORE = "_"
Private Const msFORBIDDEN_CHARS = "`|~"""

Public Const msPROLOG_SWITCHES = "/P8000 /T2000 /L256 /B256 /H512 /I400 /O400"
Public Const MACRO_SETTING_WEBPATHDEF As String = "C:\Program Files\InferMed\MACRO 3.0\www\"

Option Explicit

'---------------------------------------------------------------------
Public Function LockSubjectIfNeeded(ByRef oUser As MACROUser, ByVal lStudyId As Long, ByVal sSite As String, _
    ByVal lSubjectId As Long, ByVal lEformTaskId As Long, ByVal sASPVToken As String, ByVal sASPEToken As String, _
    ByRef sToken As String, ByRef vErrors As Variant) As Boolean
'---------------------------------------------------------------------
'   ic 15/07/2004
'   function compares user asp lock token details with the details of
'   the eform requiring a lock. if one of the lock tokens match, the
'   user already has a lock on the eform. if neither token matches,
'   try to get a subject lock
'   this function should be superceded when the locking business
'   objects are amended to allow locking of subjects when the locking
'   user has an eform of the subject open
'---------------------------------------------------------------------
Dim sTokenDatabase As String
Dim sTokenStudy As String
Dim sTokenSite As String
Dim sTokenSubject As String
Dim sTokenTaskId As String
Dim bLockOK As Boolean

    On Error GoTo CatchAllError

    bLockOK = False
    
    Call RtnLockTokenSplit(sASPVToken, sTokenDatabase, sTokenStudy, sTokenSite, sTokenSubject, sTokenTaskId, "")
    If (sTokenDatabase = oUser.DatabaseCode) And (sTokenStudy = CStr(lStudyId)) And (sTokenSite = sSite) _
    And (sTokenSubject = CStr(lSubjectId)) And (sTokenTaskId = CStr(lEformTaskId)) Then
        'visit token matches this discrepancy, no subject lock is required
        bLockOK = True
    Else
        Call RtnLockTokenSplit(sASPEToken, sTokenDatabase, sTokenStudy, sTokenSite, sTokenSubject, sTokenTaskId, "")
        If (sTokenDatabase = oUser.DatabaseCode) And (sTokenStudy = CStr(lStudyId)) And (sTokenSite = sSite) _
        And (sTokenSubject = CStr(lSubjectId)) And (sTokenTaskId = CStr(lEformTaskId)) Then
            'eform token matches this discrepancy, no subject lock is required
            bLockOK = True
        Else
            'neither user token match this discrepancy - must get a subject lock
            sToken = LockSubjectA(oUser, lStudyId, sSite, lSubjectId, vErrors)
            If (sToken <> "") Then
                bLockOK = True
            End If
        End If
    End If
    
CatchAllError:
    LockSubjectIfNeeded = bLockOK
End Function

'---------------------------------------------------------------------
Public Function LockSubjectA(ByRef oUser As MACROUser, lStudyId As Long, sSite As String, lSubjectId As Long, _
                            ByRef vErrors As Variant) As String
'---------------------------------------------------------------------
' Lock a subject.
' Returns if token if lock successful or empty string if not
' ic 22/06/2004 added parameter checking, error handling
'---------------------------------------------------------------------
Dim sLockDetails As String
Dim sMsg As String
Dim sToken As String

    On Error GoTo CatchAllError
    
    'TA 04.07.2001: use new locking
    sToken = MACROLOCKBS30.LockSubject(oUser.CurrentDBConString, oUser.UserName, lStudyId, sSite, lSubjectId)
    Select Case sToken
    Case MACROLOCKBS30.DBLocked.dblStudy
        sLockDetails = MACROLOCKBS30.LockDetailsStudy(oUser.CurrentDBConString, lStudyId)
        If sLockDetails = "" Then
            vErrors = AddToArray(vErrors, lStudyId, "This study is currently being edited by another user.")
        Else
            vErrors = AddToArray(vErrors, lStudyId, "This study is currently being edited by " & Split(sLockDetails, "|")(0) & ".")
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblSubject

        sLockDetails = MACROLOCKBS30.LockDetailsSubject(oUser.CurrentDBConString, lStudyId, sSite, lSubjectId)
        If sLockDetails = "" Then
            vErrors = AddToArray(vErrors, lSubjectId, "This subject is currently being edited by another user.")
        Else
            vErrors = AddToArray(vErrors, lSubjectId, "This subject is currently being edited by " & Split(sLockDetails, "|")(0) & ".")
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblEFormInstance
        ' NCJ 14 Jan 03 - Bug fix to lock message
        ' An eForm is in use, but we don't know which one, so give a generic message
        vErrors = AddToArray(vErrors, lSubjectId, "This subject is currently being edited by another user.")
    
        sToken = ""
    Case Else
        'hurrah, we have a lock

    End Select
    LockSubjectA = sToken
    Exit Function
    
CatchAllError:
        vErrors = AddToArray(vErrors, Err.Number, Err.Description & "|" & "clsWWW.LockSubjectA")
End Function


'---------------------------------------------------------------------
Public Sub UnlockSubjectA(ByRef oUser As MACROUser, lStudyId As Long, sSite As String, lSubjectId As Long, sToken As String, ByRef vErrors As Variant)
'---------------------------------------------------------------------
' Unlock the subject
' ic 22/06/2004 added parameter checking, error handling
'---------------------------------------------------------------------

    On Error GoTo CatchAllError
    
    'TA 04.07.2001: use new locking model
    If sToken <> "" Then
        'if no gsStudyToken then UnlockSubject is being called without a corresponding LockSubject being called first
        MACROLOCKBS30.UnlockSubject oUser.CurrentDBConString, sToken, lStudyId, sSite, lSubjectId
        'always set this to empty string for same reason as above
        sToken = ""
    End If
    Exit Sub
    
CatchAllError:
    vErrors = AddToArray(vErrors, Err.Number, Err.Description & "|" & "modUIHTML.UnlockSubjectA")
End Sub

'------------------------------------------------------------------------------'
Public Function RtnLogIllegalParametersFlagA()
'------------------------------------------------------------------------------'
'   ic 29/06/2004
'   revisions
'------------------------------------------------------------------------------'
Dim sLog As String

    On Error GoTo CatchAllError
    
    InitialiseSettingsFile True
    sLog = GetMACROSetting(MACRO_SETTING_LOG_ILLEGAL_PARAMETERS, "false")
    RtnLogIllegalParametersFlagA = (LCase(sLog) = "true")
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnLogIllegalParametersFlagA")
End Function

'------------------------------------------------------------------------------'
Public Function RtnTraceFlagA()
'------------------------------------------------------------------------------'
'   ic 29/06/2004
'   revisions
'------------------------------------------------------------------------------'
Dim sTrace As String

    On Error GoTo CatchAllError
    
    InitialiseSettingsFile True
    sTrace = GetMACROSetting(MACRO_SETTING_TRACE, "false")
    RtnTraceFlagA = (LCase(sTrace) = "true")
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnTraceFlagA")
End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnTimezoneOffset(ByVal sTimezoneOffset As String, _
                                   ByRef bOK As Boolean) As Integer
'--------------------------------------------------------------------------------------------------
'   ic 11/12/2002
'   function checks a supplied timezoneoffset string is ok, returns integer version
'   revisions
'   ic 29/06/2004 added error handling, moved from clswww
'--------------------------------------------------------------------------------------------------
Dim nTimezoneOffset As Integer

    On Error GoTo CatchAllError

    bOK = False
    nTimezoneOffset = 0
    If IsNumeric(sTimezoneOffset) Then
        bOK = ((sTimezoneOffset >= -780) And (sTimezoneOffset <= 720))
        If bOK Then nTimezoneOffset = CInt(sTimezoneOffset)
    End If
    RtnTimezoneOffset = nTimezoneOffset
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnTimezoneOffset")
End Function

'------------------------------------------------------------------------------'
Public Function RtnUseOCIdFlag() As Boolean
'------------------------------------------------------------------------------'
'   TA 18/11/2004: issue 2448 get Use OC ID flag
'   revisions
'------------------------------------------------------------------------------'
Dim sUseOC As String

    On Error GoTo CatchAllError
    
    InitialiseSettingsFile True
    sUseOC = GetMACROSetting(MACRO_SETTING_USE_OC_ID, "true")
    RtnUseOCIdFlag = (LCase(sUseOC) = "true")
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnUseOCIdFlag")
End Function


'--------------------------------------------------------------
Public Function ToggleEFIStatus(ByRef oSubject As StudySubject, _
                                 ByRef oEFI As EFormInstance, _
                                 ByRef oUser As MACROUser, _
                                 ByVal nToStatus As Integer, _
                                 ByVal nOldStatus As Integer, _
                                 ByVal bOldStatusMissingRequested As Boolean, _
                                 ByRef vErrors As Variant, _
                                 ByVal bDoVisitEForm As Boolean) As Boolean
'--------------------------------------------------------------
'Make an eForm unobtainable by changing all its 'Missing' responses to 'Unobtainable'
' or vice versa
'   revisions
'   ic 29/06/2004 added error handling, moved from clswww
'   DPH 29/09/2006 save visit eForm responses (if necessary)
'--------------------------------------------------------------
Dim oResponse As Response
Dim sLockErrMsg As String
Dim sEFILockToken As String
Dim sVEFILockToken As String
Dim vItm As Variant
Dim bUseSCI As Boolean
Dim bLoadFail As Boolean
Dim nUpdate As eUIUpdateMode
Dim vSubject As Variant

    On Error GoTo CatchAllError
    
    ToggleEFIStatus = False
        
    'load efi's responses
    'we don't need to hold onto the EFILock Token or VEFILockToken
    If oSubject.LoadResponses(oEFI, sLockErrMsg, sEFILockToken, sVEFILockToken) <> lrrReadWrite Then
        vErrors = AddToArray(vErrors, oEFI.Name, sLockErrMsg)
    Else
    
        'loop through each response
        For Each oResponse In oEFI.Responses
            If oResponse.LockStatus = eLockStatus.lsUnlocked Then
                If ((bOldStatusMissingRequested) And ((oResponse.Status = eStatus.Missing) Or (oResponse.Status = eStatus.Requested))) _
                    Or ((Not bOldStatusMissingRequested) And (oResponse.Status = eStatus.Unobtainable)) Then
                    'If oResponse.Status = nOldStatus And oResponse.LockStatus = eLockStatus.lsUnlocked Then
                    'this response is unlocked so toggle this response's status. nb we change derived questions
                    Call oResponse.SetStatusFromSchedule(nToStatus)
                End If
            End If
        Next
        
        ' DPH 29/06/2006 Check if need to save visit eform responses
        If bDoVisitEForm Then
            If Not oEFI.VisitInstance.VisitEFormInstance Is Nothing Then
                'loop through each response on associated visit eForm (will be loaded with eForm)
                For Each oResponse In oEFI.VisitInstance.VisitEFormInstance.Responses
                    If oResponse.LockStatus = eLockStatus.lsUnlocked Then
                        If ((bOldStatusMissingRequested) And ((oResponse.Status = eStatus.Missing) Or (oResponse.Status = eStatus.Requested))) _
                            Or ((Not bOldStatusMissingRequested) And (oResponse.Status = eStatus.Unobtainable)) Then
                            'If oResponse.Status = nOldStatus And oResponse.LockStatus = eLockStatus.lsUnlocked Then
                            'this response is unlocked so toggle this response's status. nb we change derived questions
                            Call oResponse.SetStatusFromSchedule(nToStatus)
                        End If
                    End If
                Next
            End If
        End If
        
        'save the responses - this will save the subject
        Select Case oSubject.SaveResponses(oEFI, sLockErrMsg)
        Case srrNoLockForSaving
            vErrors = AddToArray(vErrors, oEFI.Name, sLockErrMsg)
        Case srrSubjectReloaded
            ' we'll just try again...
            If oSubject.SaveResponses(oEFI, sLockErrMsg) <> srrSuccess Then
                vErrors = AddToArray(vErrors, oEFI.Name, "Unable to save changes because another user is editing this subject")
            End If
        Case srrSuccess
            ' OK
            ToggleEFIStatus = True
        End Select

        'remove the response from memory
        Call oSubject.RemoveResponses(oEFI, True)
    End If
    
    Exit Function
    
CatchAllError:
    On Error Resume Next
    vErrors = AddToArray(vErrors, "clsWWW.ToggleEFIStatus", Err.Description)
End Function

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
Private Function ExcludeUserRDE(sUserName As String) As Boolean
'---------------------------------------------------------------------
'REM 24/01/03
'Used to check if user 'rde' details should not be written to message table
'---------------------------------------------------------------------
    
    ExcludeUserRDE = (sUserName = "rde")

End Function

'---------------------------------------------------------------------
Private Function SQLStandardNow() As String
'---------------------------------------------------------------------
' NCJ 4 Feb 00 SR2851
' Returns Now as a double in STANDARD numeric format
' suitable for adding to SQL strings
' NB This deals with problems caused by using CDbl(Now) in SQL strings
' with non-English regional settings
' NCJ 2 Oct 02 - Use new IMedNow function
'---------------------------------------------------------------------

    SQLStandardNow = LocalNumToStandard(IMedNow)

End Function

'--------------------------------------------------------------------------------------------------
Public Function ChangePassword(ByRef oUser As MACROUser, _
                                ByVal sOldPassword As String, _
                                ByVal sNewPassword As String, _
                                ByVal bCheckPassword As Boolean, _
                                ByRef vResult As Variant, _
                                ByRef vErrors As Variant) As MACROUser
'--------------------------------------------------------------------------------------------------
'   ic 21/11/2002
'   function changes a user password (well duh!)
'--------------------------------------------------------------------------------------------------
'   revisions
'   ic 10/06/2003 changed LoginResult enum
'   ic 29/06/2004 added error handling, moved from clswww
'   ic 21/06/2005 issue 2592, create system messages when changing password
'--------------------------------------------------------------------------------------------------
Dim sMessage As String
Dim vLogin As LoginResult
Dim oSystemMessage As SysMessages
Dim conMACRO As ADODB.Connection
Dim sMessageParameters As String
Dim sHashedPassword As String
Dim sCreateDate As String
Dim sFirstLogin As String

    On Error GoTo CatchAllError

    'if requested, check old password
    If bCheckPassword Then
        vLogin = oUser.Login(GetSecurityConx(), oUser.UserName, sOldPassword, "", "MACRO Web Data Entry", "", True, "", "", False)
    Else
        vLogin = LoginResult.Success
    End If
    
    If (vLogin <> LoginResult.Failed) Then
        'attempt update password
        vResult = oUser.ChangeUserPassword(oUser.UserName, sNewPassword, sMessage, sHashedPassword, sCreateDate)
        If Not vResult Then
            vErrors = AddToArray(vErrors, "Password update", ReplaceWithJSChars(sMessage))
        Else
            'create DB connection from DB connection string
            Set conMACRO = CreateMACROCon(oUser.CurrentDBConString)
            sFirstLogin = SQLStandardNow
            
            'if connection fails don't enter message
            If Not conMACRO Is Nothing Then
                If Not ExcludeUserRDE(oUser.UserName) Then 'don't write message if its rde
                    sMessageParameters = oUser.UserName & gsPARAMSEPARATOR _
                        & sHashedPassword & gsPARAMSEPARATOR _
                        & sFirstLogin & gsPARAMSEPARATOR _
                        & sFirstLogin & gsPARAMSEPARATOR _
                        & sCreateDate
                    Set oSystemMessage = New SysMessages
                    Call oSystemMessage.AddNewSystemMessage(conMACRO, ExchangeMessageType.PasswordChange, oUser.UserName, oUser.UserName, "Change Password", sMessageParameters)
                    Set oSystemMessage = Nothing
                End If
                
                conMACRO.Close
                Set conMACRO = Nothing
            End If
            
        End If
    Else
        vErrors = AddToArray(vErrors, "Password update", "Old password incorrect")
        vResult = False
    End If
    Set ChangePassword = oUser
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.ChangePassword")
End Function

'--------------------------------------------------------------------------------------------------
Public Sub UnlockInstance(ByRef oUser As MACROUser, _
                           ByRef sASPLockToken As String)
'--------------------------------------------------------------------------------------------------
'   ic 09/01/2003
'   function unlocks a locked visit/eform instance
'   revisions
'   ic 29/06/2004 added error handling, moved from clswww
'--------------------------------------------------------------------------------------------------

Dim sDatabase As String
Dim sStudy As String
Dim sSite As String
Dim sSubject As String
Dim sTaskID As String
Dim sLockToken As String

    On Error GoTo CatchAllError
    
    If (sASPLockToken <> "") Then
        Call RtnLockTokenSplit(sASPLockToken, sDatabase, sStudy, sSite, sSubject, sTaskID, sLockToken)
        If (sLockToken <> "") Then
            Call MACROLOCKBS30.UnlockEFormInstance(oUser.CurrentDBConString, sLockToken, CLng(sStudy), sSite, CLng(sSubject), CLng(sTaskID))
        End If
        sASPLockToken = ""
    End If
    
CatchAllError:
End Sub

'--------------------------------------------------------------------------------------------------
Public Function RtnLockTokenCreate(ByVal sDatabase As String, _
                                    ByVal sStudy As String, _
                                    ByVal sSite As String, _
                                    ByVal sSubject As String, _
                                    ByVal sTaskID As String, _
                                    ByVal sLockToken As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 15/10/2002
'   function returns an asp lock token for holding in an asp session var
'   revisions
'   ic 29/06/2004 added error handling, moved from clswww
'--------------------------------------------------------------------------------------------------
Dim sRtn As String
    
    On Error GoTo CatchAllError

    If (sLockToken <> "") Then
        sRtn = sDatabase & gsDELIMITER1 _
             & sStudy & gsDELIMITER1 _
             & sSite & gsDELIMITER1 _
             & sSubject & gsDELIMITER1 _
             & sTaskID & gsDELIMITER1 _
             & sLockToken
    End If
    
    RtnLockTokenCreate = sRtn
    
CatchAllError:
End Function


'--------------------------------------------------------------------------------------------------
Public Sub RtnLockTokenSplit(ByVal sASPLockToken As String, _
                              ByRef sDatabase As String, _
                              ByRef sStudy As String, _
                              ByRef sSite As String, _
                              ByRef sSubject As String, _
                              ByRef sTaskID As String, _
                              ByRef sLockToken As String)
'--------------------------------------------------------------------------------------------------
'   ic 15/10/2002
'   function splits an asp lock token
'   revisions
'   ic 29/06/2004 added error handling, moved from clswww
'--------------------------------------------------------------------------------------------------
Dim sItems As Variant
    
    On Error GoTo CatchAllError
    
    If (sASPLockToken <> "") Then
        sItems = Split(sASPLockToken, gsDELIMITER1)
        sDatabase = sItems(0)
        sStudy = sItems(1)
        sSite = sItems(2)
        sSubject = sItems(3)
        sTaskID = sItems(4)
        sLockToken = sItems(5)
    End If
    
CatchAllError:
End Sub

'--------------------------------------------------------------------------------------------------
Public Function RtnVisitEformTaskId(ByRef oSubject As StudySubject, _
                                ByVal lEformTaskId As Long) As Long
'--------------------------------------------------------------------------------------------------
' ic 10/01/2003
' function returns a visittaskid when passed an eformtaskid
'   revisions
'   ic 29/06/2004 added error handling, moved from clswww
'--------------------------------------------------------------------------------------------------
Dim oEFI As EFormInstance
Dim lRtn As Long

    On Error GoTo CatchAllError
    
    Set oEFI = oSubject.eFIByTaskId(lEformTaskId)
    lRtn = oEFI.VisitInstance.VisitEFormInstance.EFormTaskId
    
CatchAllError:
    On Error Resume Next
    Set oEFI = Nothing
    RtnVisitEformTaskId = lRtn
End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnTaskIdNextOrPrevious(ByRef oUser As MACROUser, _
                                         ByRef oSubject As StudySubject, _
                                         ByVal lEformPageTaskId As Long, _
                                         ByVal bForward As Boolean) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 08/10/2002
'   function returns the eformtaskid of the next or previous eform as a variant array
'   (0) non-cycling taskid
'   (1) cycling taskid
' revisions
'   (3) first eform, next visit. ic 13/02/2003
'   ic 16/04/2003 fixed bug 1609: GetEformInNextVisit bChangeData value
'   ic 17/04/2003 GetNextForm bchangedata value
'   ic 03/09/2003 added bChangeData variable, add taskid and efi name for alert
'   ic 29/06/2004 added error handling, moved from clswww
'--------------------------------------------------------------------------------------------------

Dim oEFI As EFormInstance
Dim oNextEfi As EFormInstance
Dim vRtn(3) As Variant
Dim bChangeData As Boolean

    On Error GoTo CatchAllError
    
    'load current eform
    Set oEFI = oSubject.eFIByTaskId(lEformPageTaskId)

     'ic 03/09/2003 dont allow form movement to new form if no write access / visit locked or frozen
    bChangeData = CanChangeData(oUser, oSubject.Site) And Not oSubject.ReadOnly And oEFI.VisitInstance.LockStatus = eLockStatus.lsUnlocked

    'get next non-cycling eform
    Set oNextEfi = oSubject.GetNextForm(oEFI, bForward, False, bChangeData)
   
    If Not oNextEfi Is Nothing Then
        'there is a non-cycling eform
        vRtn(0) = oNextEfi.EFormTaskId
        
    End If

    'get next cycling eform
    Set oNextEfi = oSubject.GetNextForm(oEFI, bForward, True, bChangeData)
            
    If Not oNextEfi Is Nothing Then
        'there is a cycling eform
        'ic 03/09/2003 add taskid and efi name for alert
        vRtn(1) = oNextEfi.EFormTaskId
        vRtn(3) = oNextEfi.Name
    End If

    'get first eform, next visit
    Set oNextEfi = oSubject.GetFirstFormInNextVisit(oEFI, bChangeData)

    If Not oNextEfi Is Nothing Then
        'there is a next visit with an eform
        vRtn(2) = oNextEfi.EFormTaskId

    End If
    
    Set oEFI = Nothing
    Set oNextEfi = Nothing
    RtnTaskIdNextOrPrevious = vRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnTaskIdNextOrPrevious")
End Function

'--------------------------------------------------------------------------------------------------
Public Function CheckEformAvailability(ByRef oSubject As StudySubject, ByVal lEformTaskId As Long, _
    ByVal bChangeData As Boolean) As Long
'--------------------------------------------------------------------------------------------------
'   ic 11/11/2003
'   function checks whether an eform can be loaded based on its status and the users read/write permissions
'   revisions
'   ic 29/06/2004 added error handling, moved from clswww
'--------------------------------------------------------------------------------------------------
Dim oEFI As EFormInstance
    
    On Error GoTo CatchAllError
    
    Set oEFI = oSubject.eFIByTaskId(lEformTaskId)
    If ((oEFI.Status = eStatus.Requested) And (Not bChangeData)) Then
        CheckEformAvailability = 0
    Else
        CheckEformAvailability = lEformTaskId
    End If
    Set oEFI = Nothing
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.CheckEformAvailability")
End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnTaskIdInVisit(ByRef oSubject As StudySubject, ByVal lEformTaskId As Long, ByVal lVisitTaskId As Long, _
    ByVal bChangeData As Boolean) As Long
'--------------------------------------------------------------------------------------------------
'   ic 08/10/2002
'   function returns the taskid of an eform in another visit
'   if passed eform is not found in passed visit, returns first active eform taskid in the visit
'   if no active eforms are found in the visit, returns 0
'--------------------------------------------------------------------------------------------------
'   revisions
'   ic 27/05/2003 changed to use debs method
'   ic 11/11/2003 added bChangeData argument
'   ic 29/06/2004 added error handling, moved from clswww
'--------------------------------------------------------------------------------------------------
Dim oEFI As EFormInstance
Dim oNewEFI As EFormInstance

    On Error GoTo CatchAllError

    RtnTaskIdInVisit = 0
    
    ' Pick up the current eForm instance
    Set oEFI = oSubject.eFIByTaskId(lEformTaskId)
    
    'get eform in passed visit
    Set oNewEFI = oSubject.GeteFormInOtherVisit(oEFI, lVisitTaskId, True, bChangeData)
    
    'if one exists, get taskid
    If Not oNewEFI Is Nothing Then
        RtnTaskIdInVisit = oNewEFI.EFormTaskId
    End If
                        
    Set oEFI = Nothing
    Set oNewEFI = Nothing
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnTaskIdInVisit")
End Function


'--------------------------------------------------------------------------------------------------
Public Function RtnNonVisitTaskId(ByRef oSubject As StudySubject, ByVal sEformTaskId As String) As Long
'--------------------------------------------------------------------------------------------------
'   ic 18/03/2003
'   function checks whether a passed eform taskid is for a visit eform. if it is, it returns
'   the eform task id for the first eform in the visit
'   revisions
'   ic 17/04/2003 pass correct 'changedata' parameter
'   ic 29/06/2004 added error handling, moved from clswww
'--------------------------------------------------------------------------------------------------
Dim oEFI As EFormInstance
Dim oVEFI As EFormInstance
Dim lRtn As Long

    On Error GoTo CatchAllError

    lRtn = CLng(sEformTaskId)
    Set oEFI = oSubject.eFIByTaskId(CLng(sEformTaskId))
    Set oVEFI = oEFI.VisitInstance.VisitEFormInstance
    If Not oVEFI Is Nothing Then
        If oEFI.EFormTaskId = oVEFI.EFormTaskId Then
            'the user asked for the visit eform
            lRtn = oSubject.GetFirstVisitForm(oEFI.VisitInstance, (Not oSubject.ReadOnly)).EFormTaskId
        End If
    End If
    
    Set oEFI = Nothing
    Set oVEFI = Nothing
    RtnNonVisitTaskId = lRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnNonVisitTaskId")
End Function

'--------------------------------------------------------------------------------------------------
Public Function IsCreateSubjectOK(ByRef oUser As MACROUser, _
                                   ByVal sSite As String, _
                                   ByVal lStudy As Long) As Boolean
'--------------------------------------------------------------------------------------------------
'   ic 21/11/2002
'   function checks a user has rights to create a subject for the passed site/study
'   revisions
'   ic 29/06/2004 moved from clswww, added error handling
'--------------------------------------------------------------------------------------------------
Dim colStudy As Collection
Dim colSite As Collection
Dim oStudy As Study
Dim oSite As Site
Dim bIsOK As Boolean
    
    On Error GoTo CatchAllError
    
    bIsOK = False
    Set colStudy = oUser.GetNewSubjectStudies
    For Each oStudy In colStudy
        If oStudy.StudyId = lStudy Then
            Set colSite = oUser.GetNewSubjectSites(lStudy)
            For Each oSite In colSite
                If oSite.Site = sSite Then
                    bIsOK = True
                    Exit For
                End If
            Next
        End If
        If bIsOK Then Exit For
    Next
                                  
    Set colStudy = Nothing
    Set colSite = Nothing
    Set oStudy = Nothing
    Set oSite = Nothing
    IsCreateSubjectOK = bIsOK
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.IsCreateSubjectOK")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetSiteCountry(ByRef oUser As MACROUser, ByVal sSiteCode As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 20/02/2003
'   function returns a site country
'   revisions
'   ic 29/06/2004 moved from clswww, added error handling
'--------------------------------------------------------------------------------------------------
Dim colSite As Collection
Dim oSite As Site
Dim sRtn As String

    On Error GoTo CatchAllError

    Set colSite = oUser.GetAllSites
    For Each oSite In colSite
        If oSite.Site = sSiteCode Then
            sRtn = oSite.CountryName
            Exit For
        End If
    Next
    Set oSite = Nothing
    Set colSite = Nothing
    GetSiteCountry = sRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.GetSiteCountry")
End Function


'--------------------------------------------------------------------------------------------------
Public Function RtnStudyName(ByRef oUser As MACROUser, ByVal lStudy As Long) As String
'--------------------------------------------------------------------------------------------------
'   ic 25/11/2002
'   function returns a study name, passed a user object and study id
'   revisions
'   ic 29/06/2004 moved from clswww, added error handling
'--------------------------------------------------------------------------------------------------
Dim colStudy
Dim oStudy As Study
Dim sRtn As String

    On Error GoTo CatchAllError

    Set colStudy = oUser.GetAllStudies
    For Each oStudy In colStudy
        If (oStudy.StudyId = lStudy) Then
            sRtn = oStudy.StudyName
            Exit For
        End If
    Next
    
    Set oStudy = Nothing
    Set colStudy = Nothing
    RtnStudyName = sRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnStudyName")
End Function

'------------------------------------------------------------------------------'
Public Function RtnCompressionFlag() As Boolean
'------------------------------------------------------------------------------'
' ic 14/10/2003
' function returns compression on/off boolean. default is 'true'
' revisions
' ic 16/10/2003 made private to maintain compatibility
'   ic 29/06/2004 moved from clswww, added error handling
'------------------------------------------------------------------------------'
Dim sCompression As String

    On Error GoTo CatchAllError
    
    InitialiseSettingsFile True
    sCompression = GetMACROSetting(MACRO_SETTING_USECOMPRESSION, "true")
    RtnCompressionFlag = (LCase(sCompression) <> "false")
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnCompressionFlag")
End Function

'------------------------------------------------------------------------------'
Public Function RtnShowSDVScheduleMenuFlag() As Boolean
'------------------------------------------------------------------------------'
' ic 04/07/2005
' function returns compression on/off boolean. default is 'true'
'------------------------------------------------------------------------------'
Dim sShow As String

    On Error GoTo CatchAllError
    
    InitialiseSettingsFile True
    sShow = GetMACROSetting(MACRO_SETTING_SHOW_SDV_SCHEDULE_MENU, "true")
    RtnShowSDVScheduleMenuFlag = (LCase(sShow) <> "false")
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnShowSDVScheduleMenuFlag")
End Function


'--------------------------------------------------------------------------------------------------
Public Function CanChangeData(ByRef oUser As MACROUser, ByVal sSite As String) As Boolean
'--------------------------------------------------------------------------------------------------
'   ic 02/09/2003
'   function checks if user can change data based on their permissions and whether is site
'--------------------------------------------------------------------------------------------------
Dim bCan As Boolean

    On Error GoTo handler
    bCan = False
    'for www, always assume db is server
    bCan = (oUser.CheckPermission(gsFnChangeData) And Not (oUser.GetAllSites.Item(sSite).SiteLocation = 1))
handler:
    CanChangeData = bCan
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetErrorHTML(ByVal eErrorType As eWWWErrorType, _
                             ByVal sErrorDescription As String, _
                    Optional ByVal lErrorCode As Long = 0, _
                    Optional ByVal sErrorSource As String = "None specified", _
                    Optional ByVal enInterface As eInterface = iwww, _
                    Optional ByVal bHideLoader As Boolean = False) As String
'--------------------------------------------------------------------------------------------------
'   ic 12/11/2002
'   function returns an html error page
'--------------------------------------------------------------------------------------------------
Dim sHTML As String

        If enInterface = iwww Then
            sHTML = sHTML & "<html>" & vbCrLf _
                          & "<head>" & vbCrLf _
                          & "<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>" & vbCrLf _
                          & "<body onload='fnPageLoaded();'>" & vbCrLf
            
            sHTML = sHTML & "<script language='javascript'>" & vbCrLf _
                          & "function fnPageLoaded(){" & vbCrLf _
                          & "window.sWinState='';}" & vbCrLf
                          
            If (bHideLoader) Then
                sHTML = sHTML & "fnHideLoader();" & vbCrLf
            End If
            
            sHTML = sHTML & "</script>" & vbCrLf
            
            sHTML = sHTML & "<table border='0' width='95%' class='clsLabelText'>" _
                            & "<tr height='15'>" _
                              & "<td></td>" _
                            & "</tr>" _
                            & "<tr>" _
                              & "<td height='15' width='15'></td><td class='clsTableHeaderText' colspan='2'></td>" _
                            & "</tr>" _
                            & "<tr height='5'><td></td></tr><tr><td></td>" & vbCrLf
               
            Select Case eErrorType
            Case eWWWErrorType.ePermission:
                sHTML = sHTML & "<td colspan='2' class='clsMessageText'>&nbsp;<img src='../img/ico_error_perm.gif'>&nbsp;Permission denied : This user does not have adequate MACRO permissions to view this page.</td>"
            Case eWWWErrorType.eConfiguration:
                sHTML = sHTML & "<td colspan='2' class='clsMessageText'>&nbsp;<img src='../img/ico_error_conf.gif'>&nbsp;A configuration error occurred : The MACRO study has not been configured to allow this operation</td>"
            Case eWWWErrorType.eInternal, eWWWErrorType.eEForm:
                sHTML = sHTML & "<td colspan='2' class='clsMessageText'>&nbsp;<img src='../img/ico_error_int.gif'>&nbsp;An internal error occurred : The request could not be completed</td>"
            End Select
            
            sHTML = sHTML & "</tr><tr height='5'><td></td></tr>"
            
            If (lErrorCode > 0) Then
                sHTML = sHTML & "<tr>" _
                                & "<td></td><td width='100' valign='top'>Code</td><td valign='top'>" & lErrorCode & "</td>" _
                              & "</tr>"
            End If
            sHTML = sHTML & "<tr>" _
                            & "<td></td><td valign='top'>Source</td><td valign='top'>" & sErrorSource & "</td>" _
                          & "</tr>" _
                          & "<tr>" _
                            & "<td></td><td valign='top'>Description</td><td valign='top'>" & sErrorDescription & "</td>" _
                          & "</tr>" _
                          & "</table>" _
                          & "</body>" _
                          & "</html>"
            
            GetErrorHTML = sHTML
        
        Else
            Err.Raise lErrorCode, sErrorSource, sErrorDescription & "|" & sErrorSource
        End If
        
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetSecurityConx() As String
'--------------------------------------------------------------------------------------------------
' ic 12/07/2002
' function returns the connection string for the macro security db specified in the settings file
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
    On Error GoTo CatchAllError
    
    InitialiseSettingsFile True
    GetSecurityConx = GetMACROSetting(MACRO_SETTING_SECPATH, "")
    If GetSecurityConx <> "" Then
        GetSecurityConx = DecryptString(GetSecurityConx)
    End If
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.GetSecurityConx")
End Function

'------------------------------------------------------------------------------'
Public Function ServeriseValue(ByVal vValue As Variant, ByVal nDataType As Integer, ByVal sDecimalPoint As String, _
    ByVal sThousandSeparator As String) As String
'------------------------------------------------------------------------------'
'   ic 26/06/2003
'   function converts a value in client locale format (browser) to a number in server locale
'   format
'   revisions
'   ic 29/06/2004 added error handling
'------------------------------------------------------------------------------'
    On Error GoTo CatchAllError
    
    Select Case nDataType
    Case DataType.IntegerData, DataType.Real, DataType.LabTest
        ServeriseValue = StandardNumToLocal(LocalNumToStandard(vValue, False, sDecimalPoint, sThousandSeparator))
    Case Else
        ServeriseValue = vValue
    End Select
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.ServeriseValue")
End Function

'------------------------------------------------------------------------------'
Public Function StandardiseValue(ByVal vValue As Variant, ByVal nDataType As Integer) As String
'------------------------------------------------------------------------------'
'   ic 26/06/2003
'   function converts a value in local format (server locale) to a standard numeric value
'   using '.' as decimal point and ',' as thousand delimiter
'   revisions
'   ic 29/06/2004 added error handling
'------------------------------------------------------------------------------'
    On Error GoTo CatchAllError

    Select Case nDataType
    Case DataType.IntegerData, DataType.Real, DataType.LabTest
        StandardiseValue = LocalNumToStandard(vValue)
    Case Else
        StandardiseValue = vValue
    End Select
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.StandardiseValue")
End Function

'------------------------------------------------------------------------------'
Public Function LocaliseValue(ByVal vValue As Variant, ByVal nDataType As Integer, ByVal sDecimalPoint As String, _
    ByVal sThousandSeparator As String) As String
'------------------------------------------------------------------------------'
'   ic 26/06/2003
'   function converts a standard numeric value using '.' as decimal point and
'   ',' as thousand delimiter, to locale specific. preferred delimiters are always
'   passed in as client is www and may not be in same locale as server
'   revisions
'   ic 29/06/2004 added error handling
'------------------------------------------------------------------------------'
    On Error GoTo CatchAllError
    
    Select Case nDataType
    Case DataType.IntegerData, DataType.Real, DataType.LabTest
        LocaliseValue = StandardNumToLocal(vValue, sDecimalPoint, sThousandSeparator)
    Case Else
        LocaliseValue = vValue
    End Select
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.LocaliseValue")
End Function

'------------------------------------------------------------------------------'
Public Sub RtnMIMsgStatusCount(ByRef oUser As MACROUser, ByRef lRaised As Long, ByRef lResponded As Long, _
    ByRef lPlanned As Long)
'------------------------------------------------------------------------------'
'   revisions
'   ic 29/06/2004 added error handling
'   DPH 28/02/2007 - bug 2882. Pass User object into StudiesSitesWhereSQL
'------------------------------------------------------------------------------'
Dim oMDL As MIDataLists
    
    On Error GoTo CatchAllError

    Set oMDL = New MIDataLists
    'TA 30/11/2003: use MIMESSAGETRIALNAME instead of CLINICALTRIAL.CLINICALTRIALID
    ' DPH 28/02/2007 - bug 2882. Pass User into StudiesSitesWhereSQL
    Call oMDL.GetMIMsgStatusCount(oUser.CurrentDBConString, oUser.DataLists.StudiesSitesWhereSQL("MIMESSAGETRIALNAME", "MIMESSAGESITE", oUser), lRaised, lResponded, 0, lPlanned, 0)
    Set oMDL = Nothing
    Exit Sub

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnMIMsgStatusCount")
End Sub

'------------------------------------------------------------------------------'
Public Function RtnJSBoolean(ByVal bVal As Boolean) As String
'------------------------------------------------------------------------------'
'   ic 28/04/2003
'   function returns a javascript value representing the passed vb boolean
'   revisions
'   ic 29/06/2004 added error handling
'------------------------------------------------------------------------------'
    RtnJSBoolean = IIf(bVal, "1", "0")
End Function

'------------------------------------------------------------------------------'
Public Sub WriteLog(ByVal bTrace As Boolean, ByVal sLog As String)
'------------------------------------------------------------------------------'
' Write to the log file. create file if it doesnt exist
'------------------------------------------------------------------------------'
Dim n As Integer
Dim sFileName As String

    If Not bTrace Then Exit Sub
    On Error GoTo IgnoreErrors
    
    n = FreeFile
    sFileName = App.Path & "\Temp\IOLog.dat"
    Open sFileName For Append As n
    Print #n, Format(Now, "hh:mm:ss") & " " & sLog
    Close n
    
IgnoreErrors:
End Sub

'------------------------------------------------------------------------------'
Public Sub WriteErrorLog(ByVal sLocation As String, ByVal sErrorCode As String, ByVal sErrorMessage As String, _
    ByVal vParams As Variant)
'------------------------------------------------------------------------------'
' Write errors to an error log file. create file if it doesnt exist
'------------------------------------------------------------------------------'
Dim n As Integer
Dim sFileName As String

    On Error GoTo IgnoreErrors
    
    n = FreeFile
    sFileName = App.Path & "\Temp\IOErrorLog.dat"
    Open sFileName For Append As n
    Print #n, Format(Now, "dd/mm/yyyy hh:mm:ss") & " " & sLocation & gsDELIMITER1 & sErrorCode & gsDELIMITER1 _
        & sErrorMessage & "`" & Join(vParams, gsDELIMITER3)
    Close n
    
IgnoreErrors:
End Sub

'--------------------------------------------------------------------------------------------------
Public Function ReplaceControlChars(ByVal vString As Variant, ByVal sReplace As String) As String
'--------------------------------------------------------------------------------------------------
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sRtn As String
Dim nChar As Integer

    On Error GoTo CatchAllError

    For nChar = 1 To Len(vString)
        If (Asc(Mid(vString, nChar, 1)) > 32) And (Asc(Mid(vString, nChar, 1)) < 127) Then
            sRtn = sRtn & Chr(Asc(Mid(vString, nChar, 1)))
        Else
            sRtn = sRtn & sReplace
        End If
    Next
    ReplaceControlChars = sRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.ReplaceControlChars")
End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnDifferenceFromGMT(ByVal vDiff As Variant) As String
'--------------------------------------------------------------------------------------------------
'   ic 14/02/2003
'   function returns a GMT string (eg GMT +3:30)
'   revisions
'   ic 01/04/2003 moved from modUIHTMLDatabrowser, made public
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sRtn As String

    On Error GoTo CatchAllError

    sRtn = "(GMT"
    If Not IsNull(vDiff) Then
        If (vDiff <> 0) Then
            sRtn = sRtn & IIf(vDiff < 0, "+", "") & -vDiff \ 60 & ":" & Format(Abs(vDiff) Mod 60, "00")
        End If
    End If
    sRtn = sRtn & ")"
    RtnDifferenceFromGMT = sRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnDifferenceFromGMT")
End Function

'--------------------------------------------------------------------------------------------------
Public Function ReplaceWithHTMLCodes(ByVal sValue As String) As String
'--------------------------------------------------------------------------------------------------
' revisions
' ic 20/06/2003 added vbLF
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
    On Error GoTo CatchAllError

    If Not IsNull(sValue) Then
        'first replace '&' to encode possible html codes
        sValue = Replace(sValue, "&", "&#38;")
        
        'replace html tag chars
        sValue = Replace(sValue, "<", "&#60;")
        sValue = Replace(sValue, ">", "&#62;")
        
        'replace control chars
        sValue = Replace(sValue, vbCrLf, "<br>")
        sValue = Replace(sValue, vbCr, "<br>")
        sValue = Replace(sValue, vbLf, "<br>")
    End If
    ReplaceWithHTMLCodes = sValue
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.ReplaceWithHTMLCodes")
End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnSubjectText(ByVal sSubjectId As String, ByVal vSubjectLabel As Variant) As String
'--------------------------------------------------------------------------------------------------
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sRtn As String

    On Error GoTo CatchAllError

    If Not IsNull(vSubjectLabel) Then
        If (vSubjectLabel <> "") Then sRtn = vSubjectLabel
    End If
    If (sRtn = "") Then
        sRtn = "(" & sSubjectId & ")"
    End If
    RtnSubjectText = sRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnSubjectText")
End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnRecordDblDate(ByVal sDate As String, Optional ByRef bDateOK As Boolean) As Double
'--------------------------------------------------------------------------------------------------
'   ic 08/08/01
'   function takes a macro date and converts it into a regular date
'   revisions
'   ic 23/01/2003 moved from clswww
'   MLM 26/03/03: Use VB to check for date validity
'--------------------------------------------------------------------------------------------------
'Dim nDay As Integer
'Dim nMonth As Integer
'Dim nYear As Integer
'Dim vDate As Variant
    
    On Error GoTo handler
    
    If Not IsMissing(bDateOK) Then bDateOK = True
    If (sDate = "") Then
        RtnRecordDblDate = 0
    Else
        RtnRecordDblDate = CDbl(CDate(sDate))
'        sDate = Format(sDate, "dd/mm/yyyy")
'
'        nDay = CInt(Mid(sDate, 1, 2))
'        nMonth = CInt(Mid(sDate, 4, 2))
'        nYear = CInt(Mid(sDate, 7))
'
'        If (nDay > 31) Or (nMonth > 12) Then
'            GoTo handler
'        Else
'            vDate = DateSerial(nYear, nMonth, nDay)
'            RtnRecordDblDate = CDbl(vDate)
'        End If
    End If
    Exit Function
    
handler:
    If Not IsMissing(bDateOK) Then bDateOK = False
    RtnRecordDblDate = 0
End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnStatusImages(ByVal nStatus As Integer, _
                                ByVal bViewInformIcon As Boolean, _
                       Optional ByVal nLockStatus As eLockStatus = eLockStatus.lsUnlocked, _
                       Optional ByVal bUseMIEnum As Boolean = False, _
                       Optional ByVal nSDVStatus As Integer = 0, _
                       Optional ByVal nDiscStatus As Integer = 0, _
                       Optional ByVal bNote As Boolean = False, _
                       Optional ByVal bComment As Boolean = False, _
                       Optional ByVal nChanges As Integer = 0, _
                       Optional ByRef sStatusLabel As String, _
                       Optional ByRef sLockLabel As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 30/01/2003
'   function returns an html status image cluster
'   revisions
'   ic 16/04/2003 added responded discrepancy icon
'   ic 30/09/2003 dont write a blank gif to the page
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim nSDV As Integer
Dim nDisc As Integer
Dim sHTML As String
Dim sChangesHTML As String
Dim sCommentHTML As String
Dim sNoteHTML As String
Dim sStatusHTML As String
Dim sSDVHTML As String
Dim sSLabel As String
Dim sLLabel As String
Dim sSDVLabel As String
Dim sDLabel As String
Dim sToolTip As String


    On Error GoTo CatchAllError
     
    If bUseMIEnum Then
        'convert to debs enumeration
        Select Case nSDVStatus
        Case eSDVMIMStatus.ssCancelled: nSDV = eSDVStatus.ssCancelled
        Case eSDVMIMStatus.ssDone: nSDV = eSDVStatus.ssComplete
        Case eSDVMIMStatus.ssPlanned: nSDV = eSDVStatus.ssPlanned
        Case eSDVMIMStatus.ssQueried: nSDV = eSDVStatus.ssQueried
        Case Else: nSDV = eSDVStatus.ssNone
        End Select
        
        Select Case nDiscStatus
        Case eDiscrepancyMIMStatus.dsClosed: nDisc = eDiscrepancyStatus.dsClosed
        Case eDiscrepancyMIMStatus.dsRaised: nDisc = eDiscrepancyStatus.dsRaised
        Case eDiscrepancyMIMStatus.dsResponded: nDisc = eDiscrepancyStatus.dsResponded
        Case Else: nDisc = eDiscrepancyStatus.dsNone
        End Select
    Else
        nSDV = nSDVStatus
        nDisc = nDiscStatus
    End If
    
    
    'changes icon
    If (nChanges > 1) Then
        If (nChanges > 4) Then nChanges = 4
        sChangesHTML = "<img src='../img/ico_change" & (nChanges - 1) & ".gif'>"
    End If
    
    'note icon
    If bNote Then
        sNoteHTML = "<img src='../img/ico_note.gif'>"
    End If
    
    'comment icon
    If bComment Then
        sNoteHTML = "<img src='../img/ico_comment.gif'>"
    End If
    
    'lock icon
    Select Case nLockStatus
    Case eLockStatus.lsFrozen:
        sStatusHTML = "ico_frozen"
        sLLabel = "Frozen"
    Case eLockStatus.lsLocked:
        sStatusHTML = "ico_locked"
        sLLabel = "Locked"
    Case Else:
    End Select
    
    'discrepancy icon
    Select Case nDisc
    Case eDiscrepancyStatus.dsRaised:
        If sStatusHTML = "" Then sStatusHTML = "ico_disc_raise"
        sDLabel = "Raised Discrepancy"
    Case eDiscrepancyStatus.dsResponded:
        If sStatusHTML = "" Then sStatusHTML = "ico_disc_resp"
        sDLabel = "Responded Discrepancy"
    Case Else:
    End Select

    'status icon
    Select Case nStatus
        Case eStatus.Warning:
            If sStatusHTML = "" Then sStatusHTML = "ico_warn"
            sSLabel = "Warning" & sSLabel
        Case eStatus.OKWarning:
            If sStatusHTML = "" Then sStatusHTML = "ico_ok_warn"
            sSLabel = "OK Warning" & sSLabel
        Case eStatus.Inform:
            If bViewInformIcon Then
                If sStatusHTML = "" Then sStatusHTML = "ico_inform"
                sSLabel = "Inform" & sSLabel
            Else
                If sStatusHTML = "" Then sStatusHTML = "ico_ok"
                sSLabel = "OK" & sSLabel
            End If
        Case eStatus.Missing:
            If sStatusHTML = "" Then sStatusHTML = "ico_missing"
            sSLabel = "Missing" & sSLabel
        Case eStatus.Unobtainable:
            If sStatusHTML = "" Then sStatusHTML = "ico_uo"
            sSLabel = "Unobtainable" & sSLabel
        Case eStatus.NotApplicable:
            If sStatusHTML = "" Then sStatusHTML = "ico_na"
            sSLabel = "Not Applicable" & sSLabel
        Case eStatus.Success:
            If sStatusHTML = "" Then sStatusHTML = "ico_ok"
            sSLabel = "OK" & sSLabel
        Case Else:
            '-10...
            'ic 30/09/2003 dont write a blank gif to the page
'            If sStatusHTML = "" Then sStatusHTML = "blank"
    End Select

    'sdv icon
    Select Case nSDV
    Case eSDVStatus.ssQueried:
        sSDVHTML = "<img src='../img/icof_sdv_query.gif'>"
        sSDVLabel = "Queried SDV"
    Case eSDVStatus.ssPlanned:
        sSDVHTML = "<img src='../img/icof_sdv_plan.gif'>"
        sSDVLabel = "Planned SDV"
    Case eSDVStatus.ssComplete:
        sSDVHTML = "<img src='../img/icof_sdv_done.gif'>"
        sSDVLabel = "Done SDV"
    End Select

    If Not IsMissing(sStatusLabel) Then sStatusLabel = sSLabel
    If Not IsMissing(sLockLabel) Then sLockLabel = sLLabel

    'build tooltip
    If (sLLabel <> "") Then sToolTip = sLLabel & ", "
    sToolTip = sToolTip & sSLabel
    If (sSDVLabel <> "") Then sToolTip = sToolTip & ", " & sSDVLabel
    If (sDLabel <> "") Then sToolTip = sToolTip & ", " & sDLabel
    
    'ic 30/09/2003 dont write a blank gif to the page
    If sStatusHTML <> "" Then
        sStatusHTML = "<img alt='" & sToolTip & "' src='../img/" & sStatusHTML & ".gif'>"
    End If

    'build image structure
    If (sChangesHTML = "") And (sNoteHTML = "") And (sCommentHTML = "") And (sSDVHTML = "") Then
        sHTML = sStatusHTML
    Else
        sHTML = "<table cellpadding='0' cellspacing='0'><tr>"
        If (sChangesHTML <> "") Then
            sHTML = sHTML & "<td>" & sChangesHTML & "</td>"
        End If
        
        If (sCommentHTML <> "") Or (sNoteHTML <> "") Then
            sHTML = sHTML & "<td>"
            If (sCommentHTML <> "" And sNoteHTML <> "") Then
                sHTML = sHTML & "<table cellpadding='0' cellspacing='0'>" _
                              & "<tr><td>" & sCommentHTML & "</td></tr>" _
                              & "<tr><td>" & sNoteHTML & "</td></tr>" _
                              & "</table>"
            Else
                sHTML = sHTML & sCommentHTML & sNoteHTML
            End If
            sHTML = sHTML & "</td>"
        End If
        
        sHTML = sHTML & "<td>"
        If (sSDVHTML <> "") Then
            sHTML = sHTML & "<table cellpadding='0' cellspacing='0'>" _
                          & "<tr><td>" & sStatusHTML & "</td></tr>" _
                          & "<tr><td>" & sSDVHTML & "</td></tr>" _
                          & "</table>"
        Else
            sHTML = sHTML & sStatusHTML
        End If
        sHTML = sHTML & "</td>"
        sHTML = sHTML & "</tr></table>"
    End If
    
    RtnStatusImages = sHTML
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnStatusImages")
End Function

'--------------------------------------------------------------------------------------------------
Public Function ReplaceWithJSChars(ByVal sStr As String) As String
'--------------------------------------------------------------------------------------------------
' ic 10/05/2001
' function accepts a string and replaces characters in the string that interrupt javascript with
' the js equivelent escape sequence
' revisions
' ic 20/06/2003 added vbLF
' ic 21/06/2004 added / and "
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sRtn As String
    
    On Error GoTo CatchAllError
    
    sRtn = Replace(sStr, "\", "\\")
    sRtn = Replace(sRtn, "/", "\/")
    sRtn = Replace(sRtn, vbCrLf, "\n")
    sRtn = Replace(sRtn, vbCr, "\n")
    sRtn = Replace(sRtn, vbLf, "\n")
    sRtn = Replace(sRtn, "'", "\'")
    sRtn = Replace(sRtn, Chr(34), "\" & Chr(34))

    ReplaceWithJSChars = sRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.ReplaceWithJSChars")
End Function

''--------------------------------------------------------------------------------------------------
'Public Function ReplaceWithHTMLChars(ByVal sStr As String) As String
''--------------------------------------------------------------------------------------------------
'' ic 10/05/2001
'' function accepts a string and replaces characters in the string that interrupt html with
'' the html equivelent string
''--------------------------------------------------------------------------------------------------
'Dim sRtn As String
'
'    sRtn = Replace(sStr, vbCrLf, "<br>")
'    sRtn = Replace(sRtn, vbCr, "<br>")
'
'    ReplaceWithHTMLChars = sRtn
'End Function

'--------------------------------------------------------------------------------------------------
Public Function ReplaceLfWithDelimiter(ByVal sStr As String, ByVal sReplace As String) As String
'--------------------------------------------------------------------------------------------------
' ic 12/02/2003
' function accepts a string and replaces linefeeds with passed replacement string
' revisions
' ic 20/06/2003 added vbLF
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sRtn As String
    
    On Error GoTo CatchAllError
    
    sRtn = Replace(sStr, vbCrLf, sReplace)
    sRtn = Replace(sRtn, vbCr, sReplace)
    sRtn = Replace(sRtn, vbLf, sReplace)

    ReplaceLfWithDelimiter = sRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.ReplaceLfWithDelimiter")
End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnHTMLCol(lCol As Long) As String
'--------------------------------------------------------------------------------------------------
'   ic 11/07/01
'   function converts a vb long colour into a html hex colour and returns it as a string
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sRGB As String
Dim sR As String
Dim sG As String
Dim sB As String
    
    On Error GoTo CatchAllError
    
    If lCol = 0 Then
        RtnHTMLCol = "ffffff"
    Else
        sRGB = Hex(2 ^ 24 + lCol)
    
        sB = Mid(sRGB, 2, 2)
        sG = Mid(sRGB, 4, 2)
        sR = Mid(sRGB, 6, 2)
    
        RtnHTMLCol = sR + sG + sB
    End If
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnHTMLCol")
End Function

'--------------------------------------------------------------------------------------------------
Public Function ReplaceHTMLCodes(ByVal sString As String) As String
'--------------------------------------------------------------------------------------------------
'   REM 12/09/01
'
'--------------------------------------------------------------------------------------------------
' REVISIONS
' DPH 05/11/2001 - Convert all writable HTML char codes
' DPH 07/01/2003 - Skip certain characters as need not replace
' MLM 19/02/03: Changed to use HexDecodeChars, as this copes with all hexed values
' ic 16/04/2003 changed name to ReplaceHTMLCodes()
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
    On Error GoTo CatchAllError

    ' Convert spaces firstly
    sString = Replace(sString, "+", " ")
    ReplaceHTMLCodes = HexDecodeChars(sString)
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.ReplaceHTMLCodes")
End Function

'--------------------------------------------------------------------------------------------------
Public Sub AddStringToVarArr(ByRef vArr() As String, ByVal sData As String)
'--------------------------------------------------------------------------------------------------
' Add string to variant array
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
    Dim lArray As Long
    Dim bAddNew As Boolean
    
    On Error GoTo CatchAllError
    
    lArray = UBound(vArr)
    bAddNew = True
    
    If lArray = 0 And (IsEmpty(vArr(0)) Or (vArr(0) = "")) Then
        vArr(0) = sData
        bAddNew = False
    End If
    
    If bAddNew Then
        ReDim Preserve vArr(lArray + 1)
        vArr(lArray + 1) = sData
    End If
    Exit Sub

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.AddStringToVarArray")
End Sub

'--------------------------------------------------------------------------------------------------
Public Function AddToArray(ByVal vArray As Variant, _
                              ByVal sCode As String, _
                              ByVal sText As String) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 30/01/2002
'   accepts a variant - empty if items have not been added yet, otherwise a 2d array
'   returns a 2d array with the passed code and text added
'--------------------------------------------------------------------------------------------------
' revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------

    On Error GoTo CatchAllError

    If IsEmpty(vArray) Then
        ReDim vArray(1, 0)
    Else
        ReDim Preserve vArray(1, UBound(vArray, 2) + 1)
    End If
    vArray(0, UBound(vArray, 2)) = sCode
    vArray(1, UBound(vArray, 2)) = sText
    
    AddToArray = vArray
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.AddToArray")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetLocalFormatDate(ByRef oUser As MACROUser, ByVal dateToFormat As Date, ByVal enDateTimeType As eDateTimeType) As String
'--------------------------------------------------------------------------------------------------
' DPH 18/02/2003 - Get date in local format
'--------------------------------------------------------------------------------------------------
' revisions
'   ic 29/06/2004 added error handling
'   NCJ 8 Dec 05 - New eDateTimeType enumeration values
'--------------------------------------------------------------------------------------------------
Dim sDate As String
Dim sLocalFormat As String

    On Error GoTo CatchAllError

    If oUser.UserSettings.GetSetting(SETTING_LOCAL_FORMAT, False) Then
        Select Case enDateTimeType
            Case eDateTimeType.dttDMY, eDateTimeType.dttMDY, eDateTimeType.dttYMD ' Date type
                sLocalFormat = oUser.UserSettings.GetSetting(SETTING_LOCAL_DATE_FORMAT, "dd/mm/yyyy")
            Case eDateTimeType.dttT ' time only
            Case eDateTimeType.dttDMYT, eDateTimeType.dttMDYT, eDateTimeType.dttYMDT ' Date/Time
                sLocalFormat = oUser.UserSettings.GetSetting(SETTING_LOCAL_DATE_FORMAT, "dd/mm/yyyy") & " hh:mm:ss"
        End Select
    Else
        Select Case enDateTimeType
            Case eDateTimeType.dttDMY, eDateTimeType.dttMDY, eDateTimeType.dttYMD ' Date type
                sLocalFormat = "dd/mm/yyyy"
            Case eDateTimeType.dttT ' time only
                sLocalFormat = "hh:mm:ss"
            Case eDateTimeType.dttDMYT, eDateTimeType.dttMDYT, eDateTimeType.dttYMDT ' Date/Time
                sLocalFormat = "dd/mm/yyyy hh:mm:ss"
        End Select
    End If
    
    If CDbl(dateToFormat) > 0 Then
        sDate = Format(dateToFormat, sLocalFormat)
    Else
        sDate = ""
    End If
    
    GetLocalFormatDate = sDate
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.GetLocalFormatDate")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetNRCTCHTML(oElement As eFormElementRO) As String
'--------------------------------------------------------------------------------------------------
' dph 24/02/2003 - Return Normal Range / CTC blank cell
'--------------------------------------------------------------------------------------------------
' REVISIONS
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sHTML As String

    On Error GoTo CatchAllError
    
    sHTML = "<td valign='top' id='" & oElement.WebId & "_tdCTC'></td>"
    
    GetNRCTCHTML = sHTML
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.GetNRCTCHTML")
End Function

''--------------------------------------------------------------------------------------------------
'Public Function RtnNRCTCText(oElement As eFormElementRO, oResponse As Response) As String
''--------------------------------------------------------------------------------------------------
'' dph 24/02/2003 - Return Normal Range / CTC grade html
''--------------------------------------------------------------------------------------------------
'' REVISIONS
''--------------------------------------------------------------------------------------------------
'Dim sText As String
'Dim nNormalRangeLNorH As Variant
'Dim nCTCGrade As Variant
'
'    sText = ""
'
'    ' Do nothing for non-labtest questions
'    If oElement.DataType = eDataType.LabTest Then
'        If (oResponse.Status <> eStatus.InvalidData) Then
'            nNormalRangeLNorH = oResponse.NRStatus
'            nCTCGrade = oResponse.CTCGrade
'
'            'convert null values to nrNotFound and NO_CTC_GRADE
'            If VarType(nNormalRangeLNorH) = vbNull Then
'                nNormalRangeLNorH = NormalRangeLNorH.nrNotfound
'            End If
'
'            ' NCJ 6/10/00 - Bug fix (changed nNormalRangeLNorH to nCTCGrade)
'            If VarType(nCTCGrade) = vbNull Then
'                nCTCGrade = NO_CTC_GRADE
'            End If
'
'            Select Case nNormalRangeLNorH
'            Case nrLow: sText = "L"
'            Case nrNormal: sText = "N"
'            Case nrHigh: sText = "H"
'            End Select
'
'            If nCTCGrade <> -1 Then
'                sText = sText & Format(nCTCGrade)
'            End If
'        End If
'    End If
'
'    RtnNRCTCText = sText
'End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnNRCTC(ByVal nStatus As Integer, ByVal nNR As Integer, ByVal nCTC As Integer) As String
'--------------------------------------------------------------------------------------------------
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sRtn As String

    On Error GoTo CatchAllError
    
    sRtn = RtnNRCTCText(nStatus, nNR, nCTC)
    'RtnNRCTC = IIf(sRtn <> "", "<table class='clsNRCTC'><tr><td>" & sRtn & "</td></tr></table>", "")
    RtnNRCTC = IIf(sRtn <> "", "[" & sRtn & "]", "")
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnNRCTC")
End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnNRCTCText(ByVal nStatus As Integer, ByVal nNR As Integer, ByVal nCTC As Integer) As String
'--------------------------------------------------------------------------------------------------
' dph 24/02/2003 - Return Normal Range / CTC grade html
'--------------------------------------------------------------------------------------------------
' REVISIONS
' ic commented out (above) modified for use also for audit
' DPH 02/12/2003 - Added checks for null_integer
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sText As String

    On Error GoTo CatchAllError

    sText = ""

    If (nStatus <> eStatus.InvalidData) Then
    
        ' DPH 02/12/2003 - Added checks for null_integer
        'convert null values to nrNotFound and NO_CTC_GRADE
        If VarType(nNR) = vbNull Or nNR = NULL_INTEGER Then
            nNR = NormalRangeLNorH.nrNotfound
        End If
        
        ' DPH 02/12/2003 - Added checks for null_integer
        ' NCJ 6/10/00 - Bug fix (changed nNormalRangeLNorH to nCTCGrade)
        If VarType(nCTC) = vbNull Or nCTC = NULL_INTEGER Then
            nCTC = NO_CTC_GRADE
        End If
                
        Select Case nNR
        Case nrLow: sText = "L"
        Case nrNormal: sText = "N"
        Case nrHigh: sText = "H"
        End Select
        
        If nCTC <> -1 Then
            sText = sText & Format(nCTC)
        End If
    End If

    RtnNRCTCText = sText
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnNRCTCText")
End Function

'--------------------------------------------------------------------------------------------------
Public Function JSCompress(ByVal sJSString As String) As String
'--------------------------------------------------------------------------------------------------
' MLM 08/04/03: Replace a JavaScript block with a single call to a JS library function
' that will do the same as the original block by uncompressing and executing its argument.
' The purpose of this is to the reduce the download time for large, repetitive blocks of code.
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------

Const sSEP As String = "`"
Const nMIN_LENGTH As Integer = 2 '10
Const nMIN_COUNT As Integer = 1 '2
Const nSTEP As Integer = 3

Dim nLength As Integer
Dim lPos As Long
Dim sSubStr As String
Dim dicShort As Scripting.Dictionary
Dim dicLong As Scripting.Dictionary
Dim colStrings As Collection
'Dim colScores As Collection
Dim colOffsets As Collection
Dim colShortOffsets As Collection
Dim vKeys As Variant
Dim lCount As Long
Dim lCount2 As Long
Dim lCount3 As Long
Dim bEndOfRepeat As Boolean
Dim lMaxScore As Long
Dim lScore As Long
Dim nReplacements As Integer
Dim alScores() As Long
Dim aString() As Byte
Dim nMIN_SCORE As Integer
'debugging
Dim t As Single
Dim lStartLen As Long
    
    On Error GoTo CatchAllError
    
    'debugging: start time and length
    lStartLen = Len(sJSString)
    nMIN_SCORE = lStartLen / 350
    Debug.Print "Start: " & lStartLen
    t = Timer
    
    ReDim alScores(Len(sJSString))
    
    'start by doing some simple find and replaces
    sJSString = Replace(sJSString, vbCr, "")
    sJSString = Replace(sJSString, vbLf, "")
    sJSString = Replace(sJSString, sSEP, "\x" & IIf(Asc(sSEP) < 16, "0", "") & Hex(Asc(sSEP)))
    aString = sJSString
    
    Set dicLong = New Scripting.Dictionary
    Set colStrings = New Collection
'    Set colScores = New Collection
    
    'seed dictionary with short strings
    For lPos = 1 To Len(sJSString) - nMIN_LENGTH + 1
        sSubStr = Mid(sJSString, lPos, nMIN_LENGTH)
        If dicLong.Exists(sSubStr) Then
            Set colOffsets = dicLong.Item(sSubStr)
            colOffsets.Add lPos ' = dicLong.Item(sSubStr) + 1
'            For lCount = 1 To colOffsets.Count
'                alScores(colOffsets.Item(lCount)) = nMIN_LENGTH * colOffsets.Count
'            Next lCount
        Else
            Set colOffsets = New Collection
            colOffsets.Add lPos
            dicLong.Add sSubStr, colOffsets
            lPos = lPos + nMIN_LENGTH
'            alScores(lPos) = nMIN_LENGTH
        End If
    Next lPos
    nLength = nMIN_LENGTH
    
    Do
        Set dicShort = dicLong
        Set dicLong = New Scripting.Dictionary
        
        'Debug.Print dicShort.Count & "*" & nLength & ";";
        
        vKeys = dicShort.Keys
        For lCount = 0 To UBound(vKeys)
            Set colShortOffsets = dicShort.Item(vKeys(lCount))
            If colShortOffsets.Count > nMIN_COUNT Then
                bEndOfRepeat = False
                For lCount2 = 1 To colShortOffsets.Count
                    If colShortOffsets.Item(lCount2) + nLength + nSTEP - 1 > lStartLen Then
                        bEndOfRepeat = True
                        Exit For
                    Else
                        For lPos = 0 To nSTEP - 1
                            If alScores(colShortOffsets.Item(lCount2) + nLength + lPos) >= colShortOffsets.Count * nLength Then
                                bEndOfRepeat = True
                                Exit For
                            End If
                        Next lPos
                    End If
                Next lCount2
                If Not bEndOfRepeat Then
                    bEndOfRepeat = True
                    For lCount2 = 1 To colShortOffsets.Count
                        sSubStr = Mid(sJSString, colShortOffsets.Item(lCount2), nLength)
                        If dicLong.Exists(sSubStr) Then
                            Set colOffsets = dicLong.Item(sSubStr)
                            colOffsets.Add colShortOffsets.Item(lCount2)
                            For lCount3 = 1 To colOffsets.Count
                                alScores(colOffsets.Item(lCount3)) = nLength * colOffsets.Count
                            Next lCount3
                            bEndOfRepeat = False
                        Else
                            Set colOffsets = New Collection
                            colOffsets.Add colShortOffsets.Item(lCount2)
                            dicLong.Add sSubStr, colOffsets
                            alScores(lPos) = nLength
                        End If
                    Next lCount2
                End If
                'we've come to the end of a repeated string; if it scores highly enough, remember it
                If bEndOfRepeat And nLength * colShortOffsets.Count >= nMIN_SCORE Then
                    colStrings.Add vKeys(lCount)
                End If
            End If
        Next lCount
        nLength = nLength + nSTEP
    Loop Until dicLong.Count = 0
    Debug.Print
    Debug.Print "Finding " & colStrings.Count & " replacements took " & (Timer - t)
    
    For lCount = colStrings.Count To 1 Step -1
        If InStr(sJSString, colStrings(lCount)) > 0 Then
            nReplacements = nReplacements + 1
            sJSString = Replace(sJSString, colStrings(lCount), sSEP & nReplacements & sSEP) & sSEP & colStrings(lCount)
        End If
    Next lCount
                            
    Debug.Print "Making " & nReplacements & " replacements took " & (Timer - t)
    Debug.Print "End: " & Len(sJSString)
    Debug.Print "Bandwidth: " & (lStartLen - Len(sJSString)) / 1024 * 8 / (Timer - t)

    'add function header and make string arguments
    JSCompress = "fnExecCompressed('" & Replace(sJSString, "'", "\'") & sSEP & nReplacements & "');"
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.JSCompress")
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateAppState(ByVal sAppState As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sIllegal As String

    sIllegal = Chr(34)
    ValidateAppState = Not ContainsIllegalChars(sAppState, sIllegal)
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateUsername(ByVal sUserName As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET & msNUMBERS & msSPACE
    ValidateUsername = (ContainslegalChars(sUserName, sLegal) And LengthIsBetween(sUserName, 0, 20))
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidatePassword(ByVal sPassword As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET & msNUMBERS & msSPACE
    ValidatePassword = (ContainslegalChars(sPassword, sLegal) And LengthIsBetween(sPassword, 0, 100))
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateDatabase(ByVal sDatabase As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET & msNUMBERS & msSTANDARD_DOT & msSPACE & msUNDERSCORE
    ValidateDatabase = (ContainslegalChars(sDatabase, sLegal) And LengthIsBetween(sDatabase, 0, 15))
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateRole(ByVal sRole As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET & msNUMBERS & msSPACE
    ValidateRole = (ContainslegalChars(sRole, sLegal) And LengthIsBetween(sRole, 0, 15))
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateSite(ByVal sSite As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    ValidateSite = (IsAlphanumeric(sSite) And LengthIsBetween(sSite, 0, 8))
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateStudyName(ByVal sStudy As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET & msNUMBERS & msUNDERSCORE
    ValidateStudyName = (ContainslegalChars(sStudy, sLegal) And LengthIsBetween(sStudy, 0, 15))
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateLabCode(ByVal sLabCode As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET & msNUMBERS & msUNDERSCORE & msSPACE
    ValidateLabCode = (ContainslegalChars(sLabCode, sLegal) And LengthIsBetween(sLabCode, 0, 15))
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateDateTime(ByVal sDateTime As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    ValidateDateTime = Not ContainsIllegalChars(sDateTime, msFORBIDDEN_CHARS)
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateLabel(ByVal sLabel As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    ValidateLabel = Not ContainsIllegalChars(sLabel, msFORBIDDEN_CHARS)
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateText(ByVal sText As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    ValidateText = Not ContainsIllegalChars(sText, msFORBIDDEN_CHARS)
End Function

'--------------------------------------------------------------------------------------------------
'Public Function ValidateIdentifier(ByVal sIdentifier As String) As Boolean
''--------------------------------------------------------------------------------------------------
'   ic 01/07/2004
'   cant use this function to validate until the identifier is sorted out - is being created with
'   parameters in the wrong order in some places
''--------------------------------------------------------------------------------------------------
'Dim vItm As Variant
'
'    ValidateIdentifier = False
'    vItm = Split(sIdentifier, gsDELIMITER1)
'    If Not IsNumeric(vItm(eWWWIdentifier.idStudyId)) Then Exit Function
'    If Not ValidateSite(vItm(eWWWIdentifier.idSite)) Then Exit Function
'    If Not IsNumeric(vItm(eWWWIdentifier.idSubject)) Then Exit Function
'    If Not IsNumeric(vItm(eWWWIdentifier.idVisitId)) Or vItm(eWWWIdentifier.idVisitId) = "" Then Exit Function
'    If Not IsNumeric(vItm(eWWWIdentifier.idVisitCycle)) Or vItm(eWWWIdentifier.idVisitCycle) = "" Then Exit Function
'    If Not IsNumeric(vItm(eWWWIdentifier.idVisitTaskId)) Or vItm(eWWWIdentifier.idVisitTaskId) = "" Then Exit Function
'    If Not IsNumeric(vItm(eWWWIdentifier.idEformId)) Or vItm(eWWWIdentifier.idEformId) = "" Then Exit Function
'    If Not IsNumeric(vItm(eWWWIdentifier.idEformCycle)) Or vItm(eWWWIdentifier.idEformCycle) = "" Then Exit Function
'    If Not IsNumeric(vItm(eWWWIdentifier.idEformTaskId)) Or vItm(eWWWIdentifier.idEformTaskId) = "" Then Exit Function
'    If Not IsNumeric(vItm(eWWWIdentifier.idResponseId)) Or vItm(eWWWIdentifier.idResponseId) = "" Then Exit Function
'    If Not IsNumeric(vItm(eWWWIdentifier.idResponseCycle)) Or vItm(eWWWIdentifier.idResponseCycle) = "" Then Exit Function
'    If Not IsNumeric(vItm(eWWWIdentifier.idResponseTaskId)) Or vItm(eWWWIdentifier.idResponseTaskId) = "" Then Exit Function
'
'    ValidateIdentifier = True
'End Function

'--------------------------------------------------------------------------------------------------
Public Function IsAlphanumeric(ByVal sString As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET & msNUMBERS
    IsAlphanumeric = ContainslegalChars(sString, sLegal)
End Function

'--------------------------------------------------------------------------------------------------
Public Function IsAlphabetic(ByVal sString As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET
    IsAlphabetic = ContainslegalChars(sString, sLegal)
End Function

'--------------------------------------------------------------------------------------------------
Public Function LengthIsBetween(ByVal sString As String, ByVal n1 As Integer, ByVal n2 As Integer) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim nLen As Integer

    nLen = Len(sString)
    LengthIsBetween = (nLen >= n1 And nLen <= n2)
End Function

'--------------------------------------------------------------------------------------------------
Public Function ContainsIllegalChars(ByVal sString As String, ByVal sIllegal As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim n As Integer
Dim b As Boolean
    
    b = False
    For n = 1 To Len(sIllegal)
        If InStr(sString, Mid(sIllegal, n, 1)) > 0 Then
            b = True
            Exit For
        End If
    Next
    ContainsIllegalChars = b
End Function

'--------------------------------------------------------------------------------------------------
Public Function ContainslegalChars(ByVal sString As String, ByVal sLegal As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim n As Integer
Dim b As Boolean
    
    b = True
    For n = 1 To Len(sString)
        If InStr(sLegal, Mid(sString, n, 1)) = 0 Then
            b = False
            Exit For
        End If
    Next
    ContainslegalChars = b
End Function

'--------------------------------------------------------------------------------------------------
Private Function GetErrorRedirect(Optional ByVal sMsg As String = "") As String
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    GetErrorRedirect = "<HTML>" _
        & "<body onload='window.navigate(" & Chr(34) _
        & "Error.asp?msg=" & URLEncodeString(sMsg) _
        & Chr(34) & ")'>" _
        & "</body>" _
        & "</html>"
End Function

'--------------------------------------------------------------------------------------------------
Public Function URLEncodeString(ByVal sString As String) As String
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim n As Integer
Dim nAsc As Integer
Dim sEncString As String


    For n = 1 To Len(sString)
        nAsc = Asc(Mid(sString, n, 1))
        
        If (nAsc < 65) Or (nAsc > 90 And nAsc < 97) Or (nAsc > 122) Then
            sEncString = sEncString & "%" & Hex(nAsc)
        Else
            sEncString = sEncString & Mid(sString, n, 1)
        End If
    Next
    URLEncodeString = sEncString
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetErrorRedirectAndLogError(ByVal sLocation As String, ByVal sErrorCode As String, _
    ByVal sErrorMessage As String, ByVal vParams As Variant) As String
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    
    Select Case CLng(sErrorCode)
    Case (vbObjectError + 2)
        'parameter validation failed
        If RtnLogIllegalParametersFlagA() Then
            Call WriteErrorLog(sLocation, sErrorCode, sErrorMessage, vParams)
        End If
        GetErrorRedirectAndLogError = GetErrorRedirect(sErrorMessage)
    Case Else
        'unexpected error
        Call WriteErrorLog(sLocation, sErrorCode, sErrorMessage, vParams)
        GetErrorRedirectAndLogError = GetErrorRedirect()
    End Select
End Function

