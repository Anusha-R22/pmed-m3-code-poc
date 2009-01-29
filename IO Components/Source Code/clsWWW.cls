VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWWW"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------------------------
' File:         clsWWW.cls
' Copyright:    InferMed Ltd. 1998 All Rights Reserved
' Author:       Richard Meinesz, May 2001
' Purpose:      Web Data Entry/Review, MACRO I/O Layer
'--------------------------------------------------------------------------------------------------
' Revisions:
' rem 30/10/01 added path to Temp folder for oALM.Init
' MACRO 3.0
' NCJ 1 Nov 01 - Changed MACROLOCKBS22 to MACROLOCKBS30
' DPH 05/11/2001 - Convert all writable HTML char codes
' ZA 03/04/02    - Implemented Arezzo Cache Object
' NCJ 27 May 03 - Looking at better error handling...
' ic 27/05/2003   changed to call debs method in RtnTaskIdInVisit()
' ic 27/05/2003 in SaveAndLoadDataBrowser() allow empty label for changes since last session. bug 1679
' ic 28/05/2003 prevent users without monitor rights from seeing multiple subjects in SaveAndLoadDataBrowser()
' DPH 29/05/2003 - Don't check derived fields for RFCs
' ic 10/06/2003 removed eUpdatePassword enum
'               changed LoginResult enum
' RS 10/06/2003 - Added public method RtnStatusImagesHTML (defined in modHTML) for use by ASP reports
' dph 17/06/2003 added rfc in SaveForm()
' ic 19/06/2003 added registration
' ic 20/06/2003 added GetDatabaseChoiceHTML function
' ic 23/06/2003 added log parameter to Login(), bug 1706
' ic 23/06/2003 added vAlerts argument in SaveAndLoadEform(), bug 1869
' ic 26/06/2003 added decimalpoint, thousandseparator arguments to GetQuestionAuditHtml &
'               SaveAndLoadDatabrowser(), bug 1873
' ic 05/08/2003 always open subject as read/write if possible in SaveAndLoadEform(), bug 1938
' ic 18/08/2003 added passthread setting in SaveAndLoadEform() to improve performance on certain machines
' ic 19/08/2003 bug 1962, fixed incorrect split of form string in ChangePasswordForce()
'               added parameter to all oUser.Login calls to stop logging except on login action in Login()
' ic 26/08/2003 dont set optional fields to missing, set them straight to ok in SaveForm()
' ic 27/08/2003 moved eform code to modUIHTMLEform, mimessage code to modUIHTMLMIMEssage
' ic 02/09/2003 added new CanChangeData function to various functions
' ic 03/09/2003 added bChangeData variable, add taskid and efi name for alert in RtnTaskIdNextOrPrevious()
' DPH 09/09/2003 - check if the eform being dealt with is a decision form or an empty eform
' Mo 09/09/2003 Changes for the ORAMA project.
'               A call to CheckArezzoEvents added to SaveAndLoadEform.
'               The CheckArezzoEvents call is enlosed by conditional compilation (ORAMA = 1)
'               CheckArezzoEvents is in the newly added module modArezzoEvents
' ic 16/09/2003 added arezzo event handling
' ic 10/10/2003 added alternative non-compression schedule handling
' ic 16/10/2003 made RtnCompressionFlag() private to maintain compatibility
' ic 11/11/2003 added bChangeData to RtnTaskIdInVisit() call, CheckEformAvailability() call to SaveAndLoadEform()
' ic 26/02/2004 SaveAndLoadEform(): change to 'sNext' variable format: id~jsfn~saveflag.
'               modified 'no responses on eform' code
' ic 25/03/2004 added oStudydef.Terminate call to free up memory
' ic 10/05/2004 invalidate cache entry so that new mimessage status is picked up in RaiseMIMessageCall()
' ic 30/06/2004 added error handling, parameter checking, moved private functions to modUIHTML
' ic 12/07/2004 reinstated sdv fix
' ic 16/07/2004 fixed locking during sdv update
' ic 21/07/2004 changed Failed to LoginResult.Failed in ChangePasswordForce() bug 2348
' ic 24/08/2004 added setting planned sdvs to done in SaveAndLoadSchedule()
' ic 20/10/2004 bug 2440, added ValidateStudyName() check in GetSubjectListHTML()
' ic 01/11/2004 bug 2435, allow sVEformId and sVTEformId to be empty strings in GetEformPrintHTML()
' ic 20/12/2004 moved PlannedSDVsToDone() from modUIHTMLMIMessage to avoid a cycling reference
'               bug 2395 - pass the sSerialisedUser byref to LoadEform (SaveAndLoadEform) when using
'               SCI not passing thread, so that lastused eform gets set and passed back
' ic 01/04/2005 issue 2541 added show existing sdv code and arguments to RaiseMIMessageCall()
'               and SaveAndLoadSchedule()
' ic 18/05/2005 issue 2560, pass database connection for locating cached subjects
' ic 04/07/2005 issue 2464, modify sdv workflow
' NCJ 1 Nov 05 - Bug 2635 - Append R/W flag to UserNameFull in LoadSubject
' ic 11/07/2006 issue 2719 must invalidate cache after lock/freeze or changes are not
'               shown until the subject is reloaded in the cache
' NCJ 29 Dec 06 - Bug 2861 - Check response exists before raising an MIMessage
' ic 26/02/2007 bug 2669 added logoff function
' dph 28/02/2007 bug 2882. allow for tracing in SaveAndLoadAppMenuLh
' ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
'--------------------------------------------------------------------------------------------------

Option Explicit

'-----------------------------------------------------------------------------
Public Function GetAboutHTML(Optional sModuleName As String = "", Optional sVersionNumber As String = "") As String
'-----------------------------------------------------------------------------
'get html for about screen
'   revisions
'   ic 30/06/2004 added error handling
'-----------------------------------------------------------------------------
    On Error GoTo CatchAllError

    If sModuleName = "" Then
        sModuleName = "Data Entry"
    End If
    
    If sVersionNumber = "" Then
        sVersionNumber = App.Major & "." & App.Minor & "." & App.Revision
    End If

    GetAboutHTML = GetAbout(sModuleName, sVersionNumber, iwww)
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetAboutHTML")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetSecurityCon() As String
'--------------------------------------------------------------------------------------------------
' ic 12/07/2002
' function returns the connection string for the macro security db specified in the settings file
'   revisions
'   ic 30/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
    On Error GoTo CatchAllError
    
    GetSecurityCon = GetSecurityConx
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetSecurityCon")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetDatabaseCon(ByVal sSerialisedUser As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 02/12/2002
'   function returns the connection string for the chosen database of the passed serialised user
'   revisions
'   ic 30/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser

    On Error GoTo CatchAllError

    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)
    GetDatabaseCon = oUser.CurrentDBConString
    Set oUser = Nothing
    Exit Function

CatchAllError:
    Set oUser = Nothing
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetAboutHTML")
End Function

'--------------------------------------------------------------------------------------------------
Public Function Login(ByVal sUserName As String, _
                      ByVal sPassword As String, _
                      ByVal sDatabase As String, _
                      ByVal sUserRole As String, _
                      ByVal sSerialisedUser As String, _
                      ByRef sUserFullName As Variant, _
                      ByRef vResult As Variant) As Variant
'--------------------------------------------------------------------------------------------------
'   REM 24/05/01
'   revisions
'   ic 19/10/2002   changed to using business objects
'   TA 6/11/2002: return the logged in user object if succssful
'   ic 06/11/2002 return
'   DPH 12/11/2002 Corrected login parameters
'   'ic 23/06/2003 added log parameter, bug 1706
'   dph 24/03/2004 - kill user object
'   ic 18/06/2004 added parameter checking, error handling
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser
Dim vRtn As Variant

    On Error GoTo CatchAllError
    vResult = LoginResult.Failed
    
    
    If ((sDatabase <> "") And (sUserRole <> "")) Then
        'login 2nd stage - set database and role

        'validate parameters for illegal input
        If Not (ValidateDatabase(sDatabase)) Then
            Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal database parameter:" & sDatabase)
        End If
        If Not (ValidateRole(sUserRole)) Then
            Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal role parameter:" & sUserRole)
        End If
    
        'parameters are legal, continue with login
        Set oUser = New MACROUser
        vResult = oUser.Login(GetSecurityCon(), sSerialisedUser, "", "", "MACRO Web Data Entry", "", False, sDatabase, sUserRole, True)
        vRtn = oUser.GetState(False)
        Set oUser = Nothing
        Login = vRtn
    Else
        'login first stage - check username and password
        
        'validate parameters for illegal input
        If Not (ValidateUsername(sUserName)) Then
            Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal username parameter:" & sUserName)
        End If
        If Not (ValidatePassword(sPassword)) Then
            Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal password parameter: ********")
        End If
    
        'parameters are legal, continue with login
        Set oUser = New MACROUser
        vResult = oUser.Login(GetSecurityCon(), sUserName, sPassword, "", "MACRO Web Data Entry", "", True, "", "", False)
        sUserFullName = oUser.UserNameFull
        vRtn = oUser.GetState(True)
        Set oUser = Nothing
        Login = vRtn
    End If
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    If Err.Number = 3709 Then
        'handle this error. it can happen if incorrect database and/or role is used in the
        'second login stage. ignore the error, the asp will display the selectdatabase page
    Else
        'raise any other error and let the asp handle it for now
        Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.Login")
    End If
End Function

'--------------------------------------------------------------------------------------------------
Public Sub LogOff(ByRef vSerialisedUser As Variant)
'--------------------------------------------------------------------------------------------------
'   ic 26/02/2007
'   logs the user off - bug 2669
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser
Dim sSerialisedUser As String
Dim conSecurity As ADODB.Connection

    On Error GoTo CatchAllError

    'create user object
    sSerialisedUser = CStr(vSerialisedUser)
    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)
    
    'create connection
    Set conSecurity = New Connection

    'open connection to security database
    conSecurity.Open GetSecurityCon()
    conSecurity.CursorLocation = adUseClient
    
    'log to the database
    Call oUser.gLog(oUser.UserName, "LogOff", "User logged off: MACRO Web Data Entry", conSecurity)
    
    
CatchAllError:
    On Error Resume Next
    
    conSecurity.Close
    Set conSecurity = Nothing
    Set oUser = Nothing
End Sub

'--------------------------------------------------------------------------------------------------
Public Function GetSubjectListHTML(ByRef vSerialisedUser As Variant, ByVal sSite As String, _
    ByVal sStudy As String, ByVal sLabel As String, ByVal sId As String, ByVal sOrderBy As String, _
    ByVal sAscend As String, ByVal sBookmark As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 01/11/2002
'   function wrapper for moduihtml function
' DPH 08/11/2002 Changed to use Serialised User object
' DPH 20/12/2002 Byref serialiseduser
' DPH 06/01/2003 Corrected problem of sStudy being blanked out if text
' ic 21/06/2004 added parameter checking, error handling
' ic 20/10/2004 bug 2440, added ValidateStudyName() check to allow either studyid or name to be passed in
'--------------------------------------------------------------------------------------------------
Dim lId As Long
Dim oUser As MACROUser
Dim bAscend As Boolean
Dim sSerialisedUser As String

    On Error GoTo CatchAllError

    sSerialisedUser = CStr(vSerialisedUser)
    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)

    If (oUser.CheckPermission(gsFnDataEntry)) Or (oUser.CheckPermission(gsFnDataReview)) Then
        If (Not ValidateSite(sSite)) Then
            sSite = ""
        End If
        If (IsNumeric(sStudy)) Then
            sStudy = RtnStudyName(oUser, CLng(sStudy))
        Else
            If (Not ValidateStudyName(sStudy)) Then
                sStudy = ""
            End If
        End If
        If (Not ValidateLabel(sLabel)) Then
            sLabel = ""
        End If
        If IsNumeric(sId) Then
            lId = Val(sId)
        Else
            lId = -1
        End If
        
        
        If sOrderBy = "" Then sOrderBy = "-1"
        bAscend = (sAscend = "true")
        GetSubjectListHTML = GetSubjectList(oUser, sSite, sStudy, sLabel, lId, True, sOrderBy, bAscend, iwww, CInt(sBookmark))
    
        'DPH 20/12/2002 - get state to return to web page
        vSerialisedUser = oUser.GetState(False)
    Else
        Call Err.Raise(vbObjectError + 1, , "User " & oUser.UserName & " does not have data-entry or data-review permission")
    End If
    
    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    GetSubjectListHTML = GetErrorRedirectAndLogError("clsWWW.GetSubjectListHTML", Err.Number, _
        Err.Description, Array(sSite, sStudy, sLabel, lId, sOrderBy))
End Function

'--------------------------------------------------------------------------------------------------
Public Function CreateNewSubject(ByVal sSerialisedUser As String, _
                                 ByVal sSiteCode As String, _
                                 ByVal sStudyCode As String, _
                                 ByVal sUseSCI As String) As Long
'--------------------------------------------------------------------------------------------------
'   ic
'   function creates a new subject and returns the new subject's PersonId
'
'   amendments
'   ic  24/09/2002  removed locking code, added code to handle business object locking
'   DPH 08/11/2002 Changed to use Serialised User object
'   ic 21/11/2002 changed to use IsCreateSubjectOK() function
'   ic 21/02/2003 added site country code
'   ic 08/04/2003 pass in usesci flag
'   ic 29/05/2003 added tracing
'   ic 25/03/2004 added oStudydef.Terminate call to free up memory
'   ic 30/06/2004 added error handling
'   ic 18/05/2005 issue 2560, pass database connection for locating cached subjects
'--------------------------------------------------------------------------------------------------
Dim oStudyDef As StudyDefRO
Dim oSubject As StudySubject
Dim oALM As Arezzo_DM
Dim oSCI As MACROSCI30.clsSubjectCacheInterface
Dim oUser As MACROUser
Dim sDummyLockToken As String
Dim bUseSCI As Boolean
Dim vSubject As Variant
Dim lngRtn As Long

Dim bTrace As Boolean
Dim lErrNumber As Long
Dim sErrDescription As String

   
    On Error GoTo CatchAllError
    
    bTrace = RtnTraceFlag()
    bUseSCI = (LCase(sUseSCI) = "true")
    
    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)
    
    lngRtn = eOpenSubjectError.UnspecifiedError

    If IsCreateSubjectOK(oUser, sSiteCode, CLng(sStudyCode)) Then
       If bUseSCI Then
            'use the cache
            Set oSCI = New MACROSCI30.clsSubjectCacheInterface
            Call WriteLog(bTrace, "clsWWW.CreateNewSubject ->oSCI.CreateNewSubject(" & oUser.DatabaseCode & "," & sStudyCode & "," & sSiteCode & "," & oUser.UserName & "," & GetSiteCountry(oUser, sSiteCode) & "," & oUser.UserNameFull & "," & oUser.UserRole & ")")
            vSubject = oSCI.CreateNewSubject(oUser.CurrentDBConString, CLng(sStudyCode), sSiteCode, oUser.UserName, GetSiteCountry(oUser, sSiteCode), oUser.UserNameFull, oUser.UserRole)
            Call WriteLog(bTrace, "clsWWW.CreateNewSubject <-oSCI.CreateNewSubject returned " & vSubject(0))
            'get new subject id if no error code returned
            Select Case vSubject(0)
            Case 2
                lngRtn = eOpenSubjectError.StudyLocked
            Case 0
                lngRtn = vSubject(1).PersonId
            Case Else
            End Select
            
       Else
            'create a subject the old fashioned way
            'Create and initialise a new Arezzo instance
            Set oALM = New Arezzo_DM
            Call oALM.Init(App.Path & "\Temp\", msPROLOG_SWITCHES)
           
            'Create a new StudyDef class and load study def
            Set oStudyDef = New StudyDefRO
            Call oStudyDef.Load(oUser.CurrentDBConString, CLng(sStudyCode), 1, oALM)
           
            'create the subject
            Set oSubject = oStudyDef.NewSubject(sSiteCode, oUser.UserName, GetSiteCountry(oUser, sSiteCode), oUser.UserNameFull, oUser.UserRole)
            If Not oSubject.CouldNotLoad Then
                'return new subject id
                lngRtn = oSubject.PersonId
           End If
       End If
       
    Else
        Call Err.Raise(vbObjectError + 1, , "User does not have permission to create a subject")
           
    End If
    
    Set oSubject = Nothing
    
    If bUseSCI Then
        Call WriteLog(bTrace, "clsWWW.CreateNewSubject ->oSCI.ReleaseSubject(" & sStudyCode & "," & sSiteCode & "," & lngRtn & "," & vSubject(2) & ")" & vbCrLf)
        If (vSubject(0) = 0) Then oSCI.ReleaseSubject oUser.CurrentDBConString, CLng(sStudyCode), sSiteCode, lngRtn, vSubject(2)
        Set oSCI = Nothing
    Else
        oALM.Finish
        Set oALM = Nothing
        'ic 25/03/2004 MUST call the terminate method to free up memory
        Call oStudyDef.Terminate
        Set oStudyDef = Nothing
    End If
    CreateNewSubject = lngRtn
    Set oUser = Nothing
    Exit Function


CatchAllError:
    lErrNumber = Err.Number
    sErrDescription = Err.Description
    On Error Resume Next
    
    Call WriteLog(bTrace, "clsWWW.CreateNewSubject ->CatchAllError(" & Err.Number & "," & Err.Description & ")")
    Set oSubject = Nothing

    If bUseSCI Then
        If Not IsEmpty(vSubject) Then
            Call WriteLog(bTrace, "clsWWW.CreateNewSubject ->oSCI.ReleaseSubject(" & sStudyCode & "," & sSiteCode & "," & lngRtn & "," & vSubject(2) & ")" & vbCrLf)
            If (vSubject(0) = 0) Then oSCI.ReleaseSubject oUser.CurrentDBConString, CLng(sStudyCode), sSiteCode, lngRtn, vSubject(2)
        Else
            Call WriteLog(bTrace, "clsWWW.CreateNewSubject ->XoSCI.ReleaseSubject(" & sStudyCode & "," & sSiteCode & ",0,EMPTY)" & vbCrLf)
        End If
        Set oSCI = Nothing
    Else
        If IsObject(oALM) Then
            oALM.Finish
        End If
        Set oALM = Nothing
        If IsObject(oStudyDef) Then
            'ic 25/03/2004 MUST call the terminate method to free up memory
            Call oStudyDef.Terminate
        End If
        Set oStudyDef = Nothing
    End If
    Set oUser = Nothing
    Call Err.Raise(lErrNumber, , sErrDescription & "|clsWWW.CreateNewSubject")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetWebHelpURL() As String
'--------------------------------------------------------------------------------------------------
' ic 29/07/2002
' function to return connection string to the web help url from the registry
' revisions
' ic 29/01/2003 changed to use settings file
'--------------------------------------------------------------------------------------------------
Dim sURL As String
    
    On Error GoTo CatchAllError

    InitialiseSettingsFile True
    sURL = GetMACROSetting(MACRO_SETTING_WEBHELPURL, "")
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetWebHelpURL")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetMIMessageAuditHTML(ByVal sSerialisedUser As String, _
                                      ByVal sType As String, _
                                      ByVal sStudy As String, _
                                      ByVal sSite As String, _
                                      ByVal sId As String, _
                                      ByVal sSrc As String) As String
'--------------------------------------------------------------------------------------------------
' ic 22/06/2004 added parameter checking, error handling
'--------------------------------------------------------------------------------------------------
Dim nType As MIMsgType
Dim oUser As MACROUser

    On Error GoTo CatchAllError

    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)

    Select Case sType
    Case "0": nType = mimtDiscrepancy
    Case "2": nType = mimtNote
    Case Else: nType = mimtSDVMark
    End Select
    If Not IsNumeric(sStudy) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal study parameter:" & sStudy)
    End If
    If Not ValidateSite(sSite) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal site parameter:" & sSite)
    End If
    If Not IsNumeric(sId) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal id parameter:" & sId)
    End If
    If Not IsNumeric(sSrc) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal src parameter:" & sSrc)
    End If
    
    GetMIMessageAuditHTML = GetMIMessageAudit(oUser, nType, CLng(sStudy), sSite, CLng(sId), CInt(sSrc))
    
    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    GetMIMessageAuditHTML = GetErrorRedirectAndLogError("clsWWW.GetMIMessageAuditHTML", Err.Number, _
        Err.Description, Array(nType, sStudy, sSite, sId, sSrc))
End Function

'---------------------------------------------------------------------
Public Function LockSubject(ByRef oUser As MACROUser, lStudyId As Long, sSite As String, lSubjectId As Long, _
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
    
    If Not ValidateSite(sSite) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal site parameter:" & sSite)
    End If
    
    LockSubject = LockSubjectA(oUser, lStudyId, sSite, lSubjectId, vErrors)
    Exit Function
    
CatchAllError:
        vErrors = AddToArray(vErrors, Err.Number, Err.Description & "|" & "clsWWW.LockSubject")
End Function

'---------------------------------------------------------------------
Public Sub UnlockSubject(ByRef oUser As MACROUser, lStudyId As Long, sSite As String, lSubjectId As Long, sToken As String, ByRef vErrors As Variant)
'---------------------------------------------------------------------
' Unlock the subject
' ic 22/06/2004 added parameter checking, error handling
'---------------------------------------------------------------------

    On Error GoTo CatchAllError
    
    If Not ValidateSite(sSite) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal site parameter:" & sSite)
    End If
    If Not IsAlphanumeric(sToken) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal token parameter:" & sSite)
    End If
    
    Call UnlockSubjectA(oUser, lStudyId, sSite, lSubjectId, sToken, vErrors)
    Exit Sub
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.UnlockSubject")
End Sub

'------------------------------------------------------------------------------'
Public Function GetVersionInfo() As Variant
'------------------------------------------------------------------------------'
' ic 30/04/2002
' returns version info
'   revisions
'   ic 30/06/2004 added error handling
'------------------------------------------------------------------------------'
Dim vRtn As Variant

    On Error GoTo CatchAllError

    vRtn = App.FileDescription
    GetVersionInfo = vRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetVersionInfo")
End Function

'------------------------------------------------------------------------------'
Public Function GetSSLInfo() As Boolean
'------------------------------------------------------------------------------'
' ic 11/07/2002
' function returns boolean representing UseSSL flag in registry
' revisions
' ic 28/01/2003 changed to use settings file
'   ic 30/06/2004 added error handling
'------------------------------------------------------------------------------'
Dim sUseSSL As String

    On Error GoTo CatchAllError

    InitialiseSettingsFile True
    sUseSSL = GetMACROSetting(MACRO_SETTING_USESSL, "false")
    GetSSLInfo = (LCase(sUseSSL) = "true")
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetSSLInfo")
End Function

'------------------------------------------------------------------------------'
Public Function RtnUseSCIFlag() As Boolean
'------------------------------------------------------------------------------'
' ic 11/07/2002
' function returns boolean representing UseSCI flag in registry
' revisions
' ic 28/01/2003 changed to use settings file
' ic 09/04/2003 changed default to 'true'
'   ic 30/06/2004 added error handling
'------------------------------------------------------------------------------'
Dim sUseSCI As String
    
    On Error GoTo CatchAllError

    InitialiseSettingsFile True
    sUseSCI = GetMACROSetting(MACRO_SETTING_USESCI, "true")
    RtnUseSCIFlag = (LCase(sUseSCI) <> "false")
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.RtnUseSCIFlag")
End Function


'--------------------------------------------------------------------------------------------------
Public Sub ShutDownCacheManager()
'--------------------------------------------------------------------------------------------------
'   ic 22/07 2002
'   closes the cache manager down
'--------------------------------------------------------------------------------------------------
Dim oSCI As MACROSCI30.clsSubjectCacheInterface
    
    On Error Resume Next
    
    Set oSCI = New MACROSCI30.clsSubjectCacheInterface
    oSCI.ShutDownCacheManager
    Set oSCI = Nothing
End Sub


'--------------------------------------------------------------------------------------------------
Public Function GetCacheReport() As Variant
'--------------------------------------------------------------------------------------------------
'   ic 22/07/2002
'   returns cache manager status report
'   revisions
'   ic 30/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oSCI As MACROSCI30.clsSubjectCacheInterface
Dim vRtn As Variant

    On Error GoTo CatchAllError
    
    Set oSCI = New MACROSCI30.clsSubjectCacheInterface
    vRtn = oSCI.GetCacheReport
    Set oSCI = Nothing
    GetCacheReport = vRtn
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetCacheReport")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetUserSetting(ByVal sSerialisedUser As String, _
                               ByVal sOptionCode As String) As Variant
'--------------------------------------------------------------------------------------------------
' ic 19/09/2002
' function returns a user setting
'--------------------------------------------------------------------------------------------------
' revisions
'   ic 30/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sUserName As String
Dim oUser As MACROUser
Dim vRtn As Variant

    On Error GoTo CatchAllError

    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)

    With oUser.UserSettings
        Select Case sOptionCode
        Case "optFunctionKeys":
            vRtn = .GetSetting(SETTING_VIEW_FUNCTION_KEYS, False)
        Case "optSymbols":
            vRtn = .GetSetting(SETTING_VIEW_SYMBOLS, False)
        Case "optDateFormat":
            vRtn = .GetSetting(SETTING_LOCAL_FORMAT, False)
        Case "optServerDate":
            vRtn = .GetSetting(SETTING_SERVER_TIME, False)
        Case "optSplitScreen":
            vRtn = .GetSetting(SETTING_SPLIT_SCREEN, False)
        Case "optSameForm":
            vRtn = .GetSetting(SETTING_SAME_EFORM, False)
        Case "optPageLength":
            vRtn = .GetSetting(SETTING_PAGE_LENGTH, 50)
        End Select
    End With
    
    Set oUser = Nothing
    GetUserSetting = vRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetUserSetting")
End Function

'--------------------------------------------------------------------------------------------------
Private Function SaveUserSettings(ByRef oUser As MACROUser, _
                                  ByVal sForm As String) As Boolean
'--------------------------------------------------------------------------------------------------
' ic 19/09/2002
' function saves user settings, updates user object, returns boolean: page refresh required
'--------------------------------------------------------------------------------------------------
' revisions
' ic 10/01/2003 made private
' ic 09/04/2003 dont refresh after updating function keys, symbols
'   ic 30/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim vForm As Variant
Dim sValue As String
Dim sSavedValue As String
Dim bRefresh As Boolean

    On Error GoTo CatchAllError

    bRefresh = False
    vForm = Split(sForm, "&")
    With oUser.UserSettings
        'function keys
        sValue = LCase(GetSettingFromForm(vForm, "optFunctionKeys"))
        sSavedValue = LCase(oUser.UserSettings.GetSetting(SETTING_VIEW_FUNCTION_KEYS, "false"))
        If sValue = "on" Then
            sValue = "true"
        Else
            sValue = "false"
        End If
        If (sValue <> sSavedValue) Then
            Call .SetSetting(SETTING_VIEW_FUNCTION_KEYS, sValue)
            'bRefresh = True
        End If
        
        'symbols
        sValue = LCase(GetSettingFromForm(vForm, "optSymbols"))
        sSavedValue = LCase(oUser.UserSettings.GetSetting(SETTING_VIEW_SYMBOLS, "false"))
        If sValue = "on" Then
            sValue = "true"
        Else
            sValue = "false"
        End If
        If (sValue <> sSavedValue) Then
            Call .SetSetting(SETTING_VIEW_SYMBOLS, sValue)
            'bRefresh = True
        End If
        
        'local format
        sValue = LCase(GetSettingFromForm(vForm, "optDateFormat"))
        sSavedValue = LCase(oUser.UserSettings.GetSetting(SETTING_LOCAL_FORMAT, "false"))
        If sValue = "on" Then
            sValue = "true"
        Else
            sValue = "false"
        End If
        If (sValue <> sSavedValue) Then
            Call .SetSetting(SETTING_LOCAL_FORMAT, sValue)
        End If
        
        'server time
        sValue = LCase(GetSettingFromForm(vForm, "optServerDate"))
        sSavedValue = LCase(oUser.UserSettings.GetSetting(SETTING_SERVER_TIME, "false"))
        If sValue = "on" Then
            sValue = "true"
        Else
            sValue = "false"
        End If
        If (sValue <> sSavedValue) Then
            Call .SetSetting(SETTING_SERVER_TIME, sValue)
        End If
        
        'split screen
        sValue = LCase(GetSettingFromForm(vForm, "optSplitScreen"))
        sSavedValue = LCase(oUser.UserSettings.GetSetting(SETTING_SPLIT_SCREEN, "false"))
        If sValue = "on" Then
            sValue = "true"
        Else
            sValue = "false"
        End If
        If (sValue <> sSavedValue) Then
            Call .SetSetting(SETTING_SPLIT_SCREEN, sValue)
            'bRefresh = True
        End If
        
        'same eform
        sValue = LCase(GetSettingFromForm(vForm, "optSameForm"))
        sSavedValue = LCase(oUser.UserSettings.GetSetting(SETTING_SAME_EFORM, "false"))
        If sValue = "on" Then
            sValue = "true"
        Else
            sValue = "false"
        End If
        If (sValue <> sSavedValue) Then
            Call .SetSetting(SETTING_SAME_EFORM, sValue)
        End If
        
        'page length
        sValue = GetSettingFromForm(vForm, "optPageLength")
        sSavedValue = oUser.UserSettings.GetSetting(SETTING_PAGE_LENGTH, "false")
        If ((sValue <> sSavedValue) And IsNumeric(sValue)) Then
            'only save value if it is less than maximum allowed
            If (sValue <= gnMAXWWWRECORDSPERPAGE) Then
                Call .SetSetting(SETTING_PAGE_LENGTH, sValue)
            End If
        End If
    End With

    SaveUserSettings = bRefresh
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.SaveUserSettings")
End Function

'--------------------------------------------------------------------------------------------------
Private Function GetSettingFromForm(ByVal vForm As Variant, _
                                    ByVal sSetting As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 04/12/2002
'   function takes an html form string as array, returns the value of the search setting, if found
'   revisions
'   ic 30/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim nLoop As Integer
Dim vOption As Variant
Dim sRtn As String

    On Error GoTo CatchAllError

    For nLoop = LBound(vForm) To UBound(vForm)
        vOption = Split(vForm(nLoop), "=")
        If (vOption(0) = sSetting) Then
            sRtn = vOption(1)
            Exit For
        End If
    Next
    GetSettingFromForm = sRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetSettingFromForm")
End Function

'--------------------------------------------------------------------------------------------------
Public Function SaveAndLoadEform(ByRef vSerialisedUser As Variant, ByVal sSiteCode As String, ByVal sStudyCode As String, _
                                 ByVal sEformPageTaskId As String, ByVal sAllSubjectsPageId As String, _
                                 ByVal sSubjectId As String, ByVal sForm As String, ByVal sASPEFILockToken As String, _
                                 ByVal sASPVILockToken As String, ByVal sLabCode As String, ByVal sReadOnly As String, _
                                 ByVal sNext As String, ByVal sTimezoneOffset As String, ByVal sUseSCI As String, _
                                 ByVal sDecimalPoint As String, ByVal sThousandSeparator As String, _
                                 ByVal sLocalDate As String) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 01/10/02
'   accepts crfpage instance identifiers. saves page data if any. loads requested page
'   vRtn(0) = html body code
'   vRtn(1) = asp eform lock token
'   vRtn(2) = asp visit lock token
'   vRtn(3) = error description
' revisions
' DPH 08/11/2002 Changed to use Serialised User object
' ic 09/01/2003 changed to handle x subject identifier, to handle cancel&return (unlock eforms),
'               to handle passed js function call
' ic 16/01/2003 make sure that if locks arent removed, the token is returned to the asp
' ic 21/01/2003 only open subject writeable if user has change data permissions
' ic 18/03/2003 added check for visit eform task id being passed in
' ic 04/04/2003 added labcode saving
' ic 08/04/2003 pass in usesci flag
' DPH 14/05/2003 cancel & prev/next functionality ("m6"/"m7")
' ic 19/06/2003 added registration
' ic 23/06/2003 added vAlerts argument, bug 1869
' DPH 26/06/2003 added localdate, bug 1723
' ic 05/08/2003 always open subject as read/write if possible, bug 1938
' ic 18/08/2003 added passthread setting to improve performance on certain machines
' DPH 09/09/2003 - check if the eform being dealt with is a decision form or an empty eform
' ic 16/09/2003 added arezzo event handling
' ic 11/11/2003 added bChangeData to RtnTaskIdInVisit() call, CheckEformAvailability() call
' ic 26/02/2004 change to 'sNext' variable format: id~jsfn~saveflag. modified no responses
'               on eform code
' ic 28/06/2004 added parameter checking, error handling
' ic 20/12/2004 bug 2395 - pass the sSerialisedUser byref to LoadEform when using SCI not passing thread,
'               so that lastused eform gets set and passed back
' ic 18/05/2005 issue 2560, pass database connection for locating cached subjects
' NCJ 31 Oct 05 - Bug 2635 - Pass R/W flag through on end of sUserNameFull
' ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
' ic 09/10/2007 use arezzo token string variable instead of vSubject(2), introduce bCanChangeData variable
'--------------------------------------------------------------------------------------------------
Dim oSCI As MACROSCI30.clsSubjectCacheInterface
Dim oArezzo As Arezzo_DM
Dim oStudyDef As StudyDefRO
Dim oSubject As StudySubject
Dim colE As Collection
Dim oUser As MACROUser

Dim bUseSCI As Boolean
Dim vSubject As Variant
Dim vRtn(3) As Variant
Dim bLoadFail As Boolean
Dim sLoadFailMessage As String
Dim lLoadEformPageTaskId As Long
Dim vEformTaskId As Variant
Dim vAlerts As Variant
Dim vErrors As Variant
Dim sEFILockToken As String
Dim sVILockToken As String
Dim nUpdate As eUIUpdateMode
Dim bTimezoneOffsetOK As Boolean
Dim bEFIUnavailable As Boolean
Dim bVReadOnly As Boolean
Dim bEReadOnly As Boolean
Dim sJsFn As String
Dim bTrace As Boolean
Dim sRegister As String
Dim bPassThread As Boolean
'Dim bDecisionEform As Boolean
'Dim bNoResponsesEform As Boolean
Dim bArezzoPage As Boolean
Dim bArezzoEvents As Boolean
Dim bSave As Boolean
Dim sMenuOption As String
Dim lErrNumber As Long
Dim sErrDescription As String
Dim sSerialisedUser As String

'ic 09/10/2007
Dim sToken As String
Dim bCanChangeData As Boolean

' NCJ 31 Oct 05
Dim sUserNameFull As String


    On Error GoTo CatchAllError
    bArezzoEvents = False
    bTrace = RtnTraceFlag()
    bUseSCI = (LCase(sUseSCI) = "true")
    bPassThread = RtnPassThreadFlag() 'do we directly manipulate threaded objects passed from the cache
                                      'on certain machines this can lead to severe performance degrading
    
    'validate passed parameters
    If Not ValidateSite(sSiteCode) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal site parameter:" & sSiteCode)
    End If
    If Not IsNumeric(sStudyCode) And sStudyCode <> "" Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal study parameter:" & sStudyCode)
    End If
    If Not IsNumeric(sEformPageTaskId) And sEformPageTaskId <> "" Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal eformpagetaskid parameter:" & sEformPageTaskId)
    End If
    If Not IsNumeric(sAllSubjectsPageId) And sAllSubjectsPageId <> "" Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal allsubjectspageid parameter:" & sAllSubjectsPageId)
    End If
    If Not IsNumeric(sSubjectId) And sSubjectId <> "" Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal subjectid parameter:" & sSubjectId)
    End If
    If Not ValidateLabCode(sLabCode) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal labcode parameter:" & sLabCode)
    End If
    
    
    'get menu option, next function, save flag from sNext parameter
    If (sNext <> "") Then
        sMenuOption = Split(sNext, gsDELIMITER3)(0)
        sJsFn = Split(sNext, gsDELIMITER3)(1)
        bSave = (Split(sNext, gsDELIMITER3)(2) = "1")
    End If

    Set oUser = New MACROUser
    sSerialisedUser = CStr(vSerialisedUser)
    Call oUser.SetState(sSerialisedUser)
   
    'm4 menu function (cancel&goto), remove any locks, return js function caller page
    If (sMenuOption = "m4") Then
        Call UnlockInstance(oUser, sASPEFILockToken)
        Call UnlockInstance(oUser, sASPVILockToken)
        vRtn(0) = GetLoadBody(sJsFn)
        vRtn(1) = sASPEFILockToken
        vRtn(2) = sASPVILockToken
        SaveAndLoadEform = vRtn
        Set oUser = Nothing
        Exit Function
    End If
    
    'ic 05/08/2003 always open as read/write if possible. this is because user may want to add notes which
    'require no specific permission
    nUpdate = eUIUpdateMode.Read_Write
    
    'ic 09/10/2007 introduce variable to avoid many checks
    bCanChangeData = CanChangeData(oUser, sSiteCode)
    
    ' NCJ 31 Oct 05
    If bCanChangeData Then
        sUserNameFull = oUser.UserNameFull
    Else
        ' Disallow AREZZO events
        sUserNameFull = oUser.UserNameFull & "|0"
    End If
    
    
    'ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
    'ic 09/10/2007 set arezzo token to empty
    sToken = ""
    If bCanChangeData Then
        If bUseSCI Then
            Set oSCI = New MACROSCI30.clsSubjectCacheInterface
            Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform ->oSCI.LoadSubject(" & oUser.DatabaseCode & "," & sStudyCode & "," & sSiteCode & "," & sSubjectId & "," & oUser.UserName & "," & nUpdate & "," & oUser.UserNameFull & "," & oUser.UserRole & ")")
            vSubject = oSCI.LoadSubject(oUser.CurrentDBConString, CLng(sStudyCode), sSiteCode, CLng(sSubjectId), _
                            oUser.UserName, nUpdate, sUserNameFull, oUser.UserRole)
            If Not IsEmpty(vSubject) Then
                Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform <-oSCI.LoadSubject " & vSubject(0))
                If (vSubject(0) = 0) Then
                    Set oSubject = vSubject(1)
                    sToken = vSubject(2)
                    Set oStudyDef = oSubject.StudyDef
                Else
                    bLoadFail = True
                    sLoadFailMessage = vSubject(2)
                End If
            Else
                'error, cache returned nothing
                Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform <-oSCI.LoadSubject EMPTY")
                bLoadFail = True
                sLoadFailMessage = "Subject cache returned empty"
            End If
        Else
            Set oArezzo = New Arezzo_DM
            Call oArezzo.Init(App.Path & "\Temp\", msPROLOG_SWITCHES)
            Set oStudyDef = New StudyDefRO
            Call oStudyDef.Load(oUser.CurrentDBConString, CLng(sStudyCode), 1, oArezzo)
            Set oSubject = oStudyDef.LoadSubject(sSiteCode, CLng(sSubjectId), oUser.UserName, nUpdate, _
                            sUserNameFull, oUser.UserRole, False)
            
            If oSubject.CouldNotLoad Then
                bLoadFail = True
                sLoadFailMessage = oSubject.CouldNotLoadReason
                Call oSubject.StudyDef.RemoveSubject
                Set oSubject = Nothing
            End If
        End If
    Else
        'ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
        'create a subject without passing an arezzo, saving valuable seconds and processing power!!
        Set oStudyDef = New StudyDefRO
        Call oStudyDef.Load(oUser.CurrentDBConString, CLng(sStudyCode), 1) ' oArezzo)
        Set oSubject = oStudyDef.LoadSubject(sSiteCode, CLng(sSubjectId), oUser.UserName, nUpdate, _
                            sUserNameFull, oUser.UserRole, False)
    End If


    If Not bLoadFail Then
        'state that eform was opened in - so that we dont try to save eforms
        'that were opened as readonly, but are now available as writeable
        If (sReadOnly = "") Then sReadOnly = "1" & gsDELIMITER1 & "1"
        bVReadOnly = CBool(Split(sReadOnly, gsDELIMITER1)(0))
        bEReadOnly = CBool(Split(sReadOnly, gsDELIMITER1)(1))
    
        'get passed lock tokens
        Call RtnLockTokenSplit(sASPEFILockToken, "", "", "", "", "", sEFILockToken)
        Call RtnLockTokenSplit(sASPVILockToken, "", "", "", "", "", sVILockToken)
    
        'ic 10/09/2003 added arezzo decision check
        ' DPH 09/09/2003 - check if the form being dealt with is a decision form or an empty eform
        If (Left(sForm, 4) = "next") _
        And (InStr(sForm, "labcode") = 0) _
        And (InStr(sForm, "arezzo") <> 0) Then
            bArezzoPage = True
        Else
            bArezzoPage = False
        End If
        
        If (bArezzoPage) Then
        
            #If ORAMA = 1 Then
            'ic 11/09/2003 save arezzo decision/action outcome
            If (oSubject.ReadOnly) Then
                'want to save arezzo page but cant get a writeable subject: tell user
                vErrors = AddToArray(vErrors, "Unable to save new Arezzo state", oSubject.ReadOnlyReason)
            Else
                If oUser.CheckPermission(gsFnChangeData) Then
                    'ic 09/10/2007 use new bCanChangeData variable
                    If Not bPassThread And bUseSCI And CanChangeData Then
                        'ic 09/10/2007 pass in new token
                        Call oSCI.SaveArezzoEvents(oUser.CurrentDBConString, (sStudyCode), sSiteCode, CLng(sSubjectId), sToken, sForm)
                    Else
                        Call SaveArezzoEvents(oSubject, sForm)
                    End If
                End If
            End If
            #End If
            
        Else
        
            'ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
            'work out if we need to save - need to check for special case of dr users who will not
            'have an arezzo, making the subject readonly, but still want to save their mimessages
            Dim bERO As Boolean
            
            If (oSubject.Arezzo Is Nothing) Then
                bERO = False
            Else
                bERO = oSubject.ReadOnly
            End If
            
            
            'save eform - must have permission, a form string, eform lock token
            ' DPH 14/05/2003 - & not be a cancel & move (no save required)
            If (Not bERO) And (sForm <> "") And (sEFILockToken <> "") And bSave Then
                'check to screen for 'next=e#####' form contents (submitted from decision pages)
                ' DPH 26/06/2003 - added localdate
                ' DPH 09/09/2003 - Added check to tell if is decision page (previously incorrectly screened out empty eform)
                'If (Left(sForm, 4) <> "next") Then
                'ic 18/08/2003 do we manipulate passed threaded objects
                If Not bPassThread And bUseSCI And bCanChangeData Then
                    'no, allow the cache to manipulate these objects
                    Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform ->oSCI.SaveEform(...)")
                    'ic 09/10/2007 pass in new token
                    vErrors = oSCI.SaveEform(sSerialisedUser, CLng(sStudyCode), sSiteCode, CLng(sSubjectId), sToken, _
                        sEformPageTaskId, sForm, sEFILockToken, sVILockToken, bVReadOnly, bEReadOnly, sLabCode, _
                        sDecimalPoint, sThousandSeparator, sRegister, sLocalDate, RtnTimezoneOffset(sTimezoneOffset, bTimezoneOffsetOK))
                    Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform <-oSCI.SaveEform(...)")
                Else
                    Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform ->SaveForm(...)")
                    vErrors = SaveForm(oUser, oSubject, sEformPageTaskId, sForm, sEFILockToken, sVILockToken, _
                        bVReadOnly, bEReadOnly, sLabCode, sDecimalPoint, sThousandSeparator, sRegister, _
                        sLocalDate, RtnTimezoneOffset(sTimezoneOffset, bTimezoneOffsetOK))
                    Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform <-SaveForm(...)")
                End If
                
                'if registration was attempted, return message
                If (sRegister <> "") Then
                    vAlerts = AddToArray(vAlerts, "", sRegister)
                End If
                
                'if saveform successfully called 'SaveResponses()' method (removed locks), clear passed asp lock token strings
                If (sEFILockToken = "") Then sASPEFILockToken = ""
                If (sVILockToken = "") Then sASPVILockToken = ""

            End If
            
            'fail safe - remove any locks here, though there should be none
            Call UnlockInstance(oUser, sASPEFILockToken)
            Call UnlockInstance(oUser, sASPVILockToken)
        End If
        
    
        lLoadEformPageTaskId = 0
        If IsEmpty(vErrors) Then
            'next routine
            
            'Mo 9/9/2003 new call to CheckArezzoEvents added for the ORAMA project
            #If ORAMA = 1 Then
                If (Not oSubject.ReadOnly) And oUser.CheckPermission(gsFnChangeData) Then
                    If Not bPassThread And bUseSCI Then
                        'ic 09/10/2007 pass in new token
                        vRtn(0) = oSCI.CheckArezzoEvents(oUser.CurrentDBConString, (sStudyCode), sSiteCode, CLng(sSubjectId), sToken, _
                            oUser.DatabaseCode, sEformPageTaskId, sMenuOption, bArezzoEvents)
                    Else
                        vRtn(0) = CheckArezzoEvents(oSubject, oUser.DatabaseCode, sEformPageTaskId, sMenuOption, bArezzoEvents)
                    End If
                End If
            #End If
            
            If (Not bArezzoEvents) Then
                Select Case sMenuOption
                    Case "":
                        'no next value - load using passed identifiers
                        If (sAllSubjectsPageId <> "") Then
                            'cross subject page identifier supplied
                            lLoadEformPageTaskId = oSubject.StudyDef.Subject.GetEFIbyAllSubjectsKey(sAllSubjectsPageId).EFormTaskId
                        ElseIf (sEformPageTaskId <> "") Then
                            'eformpagetaskid identifier supplied
                            'check its not a visit eform task id, if it is, get first eform task id of visit
                            lLoadEformPageTaskId = RtnNonVisitTaskId(oSubject, sEformPageTaskId)
                        Else
                            'assume 'same eform' option
                            lLoadEformPageTaskId = oSubject.GetEFIbyAllSubjectsKey(oUser.UserSettings.GetSetting(SETTING_LAST_USED_EFORM, "1|1|1|1")).EFormTaskId
                        End If
                    Case "m0":
                        'm0=save&same
                        lLoadEformPageTaskId = CLng(sEformPageTaskId)
                    Case "m1":
                        'm1=save&prev
                        ' DPH 14/05/2003 m6=cancel&prev
                        'ic 03/09/2003 pass user object
                        vEformTaskId = RtnTaskIdNextOrPrevious(oUser, oSubject, CLng(sEformPageTaskId), False)
                        If (vEformTaskId(0) = vEformTaskId(1)) And vEformTaskId(0) <> "" Then
                            'previous eform is non-cycling
                            'load previous eform
                            lLoadEformPageTaskId = vEformTaskId(0)
                        ElseIf (vEformTaskId(0) = vEformTaskId(1)) And vEformTaskId(0) = "" Then
                            'no more eforms in this visit
                            'reload same eform
                            lLoadEformPageTaskId = CLng(sEformPageTaskId)
                            vAlerts = AddToArray(vAlerts, "", "This is the first eform in the current visit")
                        End If
                    Case "m2", "m5":
                        'm2=save&next
                        ' DPH 14/05/2003 m7=cancel&next
                        'ic 03/09/2003 pass user object
                        vEformTaskId = RtnTaskIdNextOrPrevious(oUser, oSubject, CLng(sEformPageTaskId), True)
                        If (vEformTaskId(0) = vEformTaskId(1)) And vEformTaskId(0) <> "" Then
                            'next eform is non-cycling
                            'load next eform
                            lLoadEformPageTaskId = vEformTaskId(0)
                        ElseIf (vEformTaskId(0) <> vEformTaskId(1)) And (vEformTaskId(0) <> "") And (vEformTaskId(1) <> "") Then
                            'next eform is cycling but there is a non-cycling eform in this visit after that
                            'choose between cycling and non-cycling eform
                        ElseIf (vEformTaskId(0) = "") And (vEformTaskId(1) <> "") And (vEformTaskId(2) <> "") Then
                            'next eform is cycling but there is a first eform in the next visit
                            'choose between cycling eform and next visit eform
                        ElseIf (vEformTaskId(1) <> "") And (vEformTaskId(2) = "") Then
                            'next eform is cycling and there are no more eforms in following visits
                            'choose between cycling eform and reload same eform
                        ElseIf (vEformTaskId(0) & vEformTaskId(1) = "") And (vEformTaskId(2) <> "") Then
                            'no more eforms in this visit but there is a first eform in the next visit
                            'choose between next visit eform and reload same eform
                        ElseIf (vEformTaskId(0) & vEformTaskId(1) & vEformTaskId(2) = "") Then
                            'no more eforms in this visit or following visits
                            'reload same eform
                            lLoadEformPageTaskId = CLng(sEformPageTaskId)
                            vAlerts = AddToArray(vAlerts, "", "This is the last eForm in the study")
                        End If
                    Case "m3":
                        'm3=save&goto
                        'this call will have saved the eform, goto is handled by load routine below
                    Case "m4":
                        'm4=cancel&goto
                        'this call will have unlocked eform/visit, will never get here as this is handled at
                        'the top of the function
                    Case Else:
                        Select Case Left(sMenuOption, 1)
                        Case "e"
                            'goto different eform, same visit
                            lLoadEformPageTaskId = CheckEformAvailability(oSubject, CLng(Mid(sMenuOption, 2)), _
                                bCanChangeData)
                            
                            If lLoadEformPageTaskId = 0 Then
                                lLoadEformPageTaskId = CLng(sEformPageTaskId)
                                vAlerts = AddToArray(vAlerts, "", "You may not enter new data for this subject")
                            End If
                        Case "v"
                            'goto same eform, different visit
                            lLoadEformPageTaskId = RtnTaskIdInVisit(oSubject, CLng(sEformPageTaskId), CLng(Mid(sMenuOption, 2)), _
                                bCanChangeData)
                            
                            If lLoadEformPageTaskId = 0 Then
                                lLoadEformPageTaskId = CLng(sEformPageTaskId)
                                vAlerts = AddToArray(vAlerts, "", "There are no eForms in the selected visit that you may access")
                            End If
                        Case "h"
                            lLoadEformPageTaskId = oSubject.eFIByTaskId(CLng(sEformPageTaskId)).GetHotlinkTarget(Mid(sMenuOption, 2))
                            If lLoadEformPageTaskId = 0 Then
                                'Hotlink could not be followed; stay on same form
                                lLoadEformPageTaskId = CLng(sEformPageTaskId)
                                vAlerts = AddToArray(vAlerts, "", "The eForm associated with this link is not available.")
                            End If
                        End Select
                End Select
            End If
        Else
            'if any errors, return to the same page
            lLoadEformPageTaskId = CLng(sEformPageTaskId)
        End If
        
        
        If (Not bArezzoEvents) Then
            'load routine
            If (lLoadEformPageTaskId > 0) Then
                'we have calculated a pagetaskid to load
                'ic 18/08/2003 do we manipulate passed threaded objects
                'ic 09/10/2007 use new bCanChangeData variable
                If Not bPassThread And bUseSCI And bCanChangeData Then
                    'no allow the cache to manipulate these objects
                    Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform ->oSCI.LoadEform(...)")
                    'ic 09/10/2007 pass in new token
                    vRtn(0) = oSCI.LoadEform(sSerialisedUser, CLng(sStudyCode), sSiteCode, CLng(sSubjectId), sToken, lLoadEformPageTaskId, sEFILockToken, _
                    sVILockToken, bEFIUnavailable, sDecimalPoint, sThousandSeparator, vAlerts, vErrors, (sMenuOption = "m5") And IsEmpty(vAlerts))
                    'update the state of the user object with any changes made during load (last loaded eform)
                    Call oUser.SetState(sSerialisedUser)
                    Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform <-oSCI.LoadEform(...)")
                Else
                    Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform ->GetEformBody(...)")
                    vRtn(0) = GetEformBody(oUser, oSubject, oUser.DatabaseCode, sSiteCode, lLoadEformPageTaskId, sEFILockToken, _
                    sVILockToken, bEFIUnavailable, sDecimalPoint, sThousandSeparator, iwww, vAlerts, vErrors, (sMenuOption = "m5") And IsEmpty(vAlerts), bTrace)
                    Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform <-GetEformBody(...)")
                End If
                
                If Not bEFIUnavailable Then
                    vRtn(1) = RtnLockTokenCreate(oUser.DatabaseCode, sStudyCode, sSiteCode, sSubjectId, CStr(lLoadEformPageTaskId), sEFILockToken)
                    If (sVILockToken <> "") Then
                        vRtn(2) = RtnLockTokenCreate(oUser.DatabaseCode, sStudyCode, sSiteCode, sSubjectId, CStr(RtnVisitEformTaskId(oSubject, lLoadEformPageTaskId)), sVILockToken)
                    End If
                Else
                    vAlerts = AddToArray(vAlerts, "", "Another user is entering the visit date and new eForms cannot be opened in this visit")
                    vRtn(0) = GetLoadBody("fnScheduleUrl(" & Chr(34) & sStudyCode & Chr(34) & "," & Chr(34) & sSiteCode & Chr(34) & "," & Chr(34) & sSubjectId & Chr(34) & ")", vAlerts)
                    vRtn(1) = sASPEFILockToken
                    vRtn(2) = sASPVILockToken
                End If
            Else
                'no pagetaskid to load. either we have a choice of next/previous eforms (cycling/non)
                'or we are saving and goto
                Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform ->GetDecisionBody(...)")
                If (sMenuOption = "m1") Or (sMenuOption = "m2") Or sMenuOption = "m5" Then
                    'choice page
                    'ic 23/06/2003 added vAlerts argument, bug 1869
                    vRtn(0) = GetDecisionBody(vEformTaskId(0), vEformTaskId(1), vEformTaskId(2), CLng(sEformPageTaskId), vEformTaskId(3), oUser.DatabaseCode, sStudyCode, sSiteCode, sSubjectId, vAlerts, iwww)
                    vRtn(1) = RtnLockTokenCreate(oUser.DatabaseCode, sStudyCode, sSiteCode, sSubjectId, CStr(lLoadEformPageTaskId), sEFILockToken)
                    If (sVILockToken <> "") Then
                        vRtn(2) = RtnLockTokenCreate(oUser.DatabaseCode, sStudyCode, sSiteCode, sSubjectId, CStr(RtnVisitEformTaskId(oSubject, lLoadEformPageTaskId)), sASPVILockToken)
                    End If
                Else
                    'goto page
                    vRtn(0) = GetLoadBody(sJsFn, vAlerts)
                End If
                Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform <-GetDecisionBody(...)")
            End If
        End If
    Else
        Call Err.Raise(vbObjectError + 1, , sLoadFailMessage)
    End If
    vSerialisedUser = oUser.GetState(False)
    
    
    'clear up
    Set oSubject = Nothing
    
    'ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
    If bCanChangeData Then
        If bUseSCI Then
            'ic 09/10/2007 pass in new token
            Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform ->oSCI.ReleaseSubject(" & sStudyCode & "," & sSiteCode & "," & sSubjectId & "," & sToken & ")" & vbCrLf)
            If (vSubject(0) = 0) Then oSCI.ReleaseSubject oUser.CurrentDBConString, CLng(sStudyCode), sSiteCode, CLng(sSubjectId), sToken
            Set oSCI = Nothing
        Else
            oArezzo.Finish
            Set oArezzo = Nothing
            'ic 25/03/2004 MUST call the terminate method to free up memory
            oStudyDef.Terminate
            Set oStudyDef = Nothing
        End If
    Else
        oStudyDef.Terminate
        Set oStudyDef = Nothing
    End If
    Set oUser = Nothing
    SaveAndLoadEform = vRtn
    Exit Function
   
   
CatchAllError:
    lErrNumber = Err.Number
    sErrDescription = Err.Description
    On Error Resume Next
    
    Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform ->CatchAllError(" & lErrNumber & "," & sErrDescription & ")")
    Set oSubject = Nothing
    Call UnlockInstance(oUser, sASPEFILockToken)
    Call UnlockInstance(oUser, sASPVILockToken)

    'ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
    If bCanChangeData Then
        If bUseSCI Then
            If Not IsEmpty(vSubject) Then
                'ic 09/10/2007 pass in new token
                Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform ->oSCI.ReleaseSubject(" & sStudyCode & "," & sSiteCode & "," & sSubjectId & "," & sToken & ")" & vbCrLf)
                If (vSubject(0) = 0) Then oSCI.ReleaseSubject oUser.CurrentDBConString, CLng(sStudyCode), sSiteCode, CLng(sSubjectId), sToken
            Else
                Call WriteLog(bTrace, "clsWWW.SaveAndLoadEform ->XoSCI.ReleaseSubject(" & sStudyCode & "," & sSiteCode & "," & sSubjectId & ",EMPTY)" & vbCrLf)
            End If
            Set oSCI = Nothing
        Else
            If IsObject(oArezzo) Then
                oArezzo.Finish
            End If
            Set oArezzo = Nothing
            If IsObject(oStudyDef) Then
                'ic 25/03/2004 MUST call the terminate method to free up memory
                oStudyDef.Terminate
            End If
            Set oStudyDef = Nothing
        End If
    Else
        If IsObject(oStudyDef) Then
            'ic 25/03/2004 MUST call the terminate method to free up memory
            oStudyDef.Terminate
        End If
        Set oStudyDef = Nothing
    End If
    Set oUser = Nothing
    Call Err.Raise(lErrNumber, , sErrDescription & "|clsWWW.SaveAndLoadEform")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetQuestionAuditHtml(ByVal sSerialisedUser As String, ByVal sStudy As String, ByVal sSite As String, _
                                     ByVal sPersonID As String, ByVal sEformPageTaskId As String, _
                                     ByVal sEformElementId As String, ByVal sElementCode As String, _
                                     ByVal sDecimalPoint As String, ByVal sThousandSeparator As String, _
                            Optional ByVal sRepeat As String = "1") As String
'--------------------------------------------------------------------------------------------------
'   ic 10/10/2002
'   function returns the html for a response audit trail
' revisions
' DPH 08/11/2002 Changed to use Serialised User object
' ic 26/06/2003 added decimal point, thousandseparator, bug 1873
' ic 22/06/2004 added parameter checking, error handling
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser

    On Error GoTo CatchAllError
    
    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)
    
    'check passed parameters for correct data types
    If Not IsNumeric(sStudy) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal study parameter:" & sStudy)
    End If
    If Not ValidateSite(sSite) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal site parameter:" & sSite)
    End If
    If Not IsNumeric(sPersonID) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal personid parameter:" & sPersonID)
    End If
    If Not IsNumeric(sEformPageTaskId) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal eformpagetaskid parameter:" & sEformPageTaskId)
    End If
    If Not IsNumeric(sEformElementId) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal eformelementid parameter:" & sEformElementId)
    End If
    If Not IsNumeric(sRepeat) Then sRepeat = "1"
    
    GetQuestionAuditHtml = GetQuestionAudit(oUser, CLng(sStudy), sSite, CLng(sPersonID), _
        CLng(sEformPageTaskId), CLng(sEformElementId), sElementCode, sDecimalPoint, _
        sThousandSeparator, CInt(sRepeat), iwww)
    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    'cant return redirect script as modaldialog will open a new ie window
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetQuestionAuditHTML")
End Function

'--------------------------------------------------------------------------------------------------
Public Sub UnlockASPInstance(ByVal sSerialisedUser As String, _
                             ByVal sASPLockToken As String)
'--------------------------------------------------------------------------------------------------
'   ic 14/10/2002
'   function unlocks a locked visit/eform instance
'   revisions
'   ic 17/04/2003 set ouser=nothing
'   ic 30/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser
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
            Set oUser = New MACROUser
            Call oUser.SetState(sSerialisedUser)
            Call MACROLOCKBS30.UnlockEFormInstance(oUser.CurrentDBConString, sLockToken, CLng(sStudy), sSite, CLng(sSubject), CLng(sTaskID))
            Set oUser = Nothing
        End If
    End If
    
CatchAllError:
End Sub

'--------------------------------------------------------------------------------------------------
Public Function GetRFCHTML(ByRef oStudyDef As StudyDefRO) As String
'--------------------------------------------------------------------------------------------------
'   ic 30/10/2002
'   wrapper for modUIHTML.GetRFC
'--------------------------------------------------------------------------------------------------
    
    On Error GoTo CatchAllError
    
    GetRFCHTML = GetRFC(oStudyDef, iwww)
    
CatchAllError:
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetRFOHTML(ByRef oStudyDef As StudyDefRO) As String
'--------------------------------------------------------------------------------------------------
'   ic 30/10/2002
'   wrapper for modUIHTML.GetRFO
'--------------------------------------------------------------------------------------------------

    On Error GoTo CatchAllError

    GetRFOHTML = GetRFO(oStudyDef, iwww)

CatchAllError:
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetQuestionDefinitionHTML(ByRef oEformElement As eFormElementRO) As String
'--------------------------------------------------------------------------------------------------
'   ic 29/10/2002
'   wrapper for modUIHTML.GetQUestionDefinition
'--------------------------------------------------------------------------------------------------

    On Error GoTo CatchAllError

    GetQuestionDefinitionHTML = GetQuestionDefinition(oEformElement, iwww)

CatchAllError:
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetAppMenuHeaderLhHTML(ByVal sSerialisedUser As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 05/11/2002
'   wrapper for modUIHTML.GetAppMenuHeader
' revisions
' DPH 08/11/2002 Changed to use Serialised User object
' ic 28/01/2003 changed to use settings file
' ic 30/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sMACROWebPath As String
Dim sLogoPath As String
Dim oUser As MACROUser

    On Error GoTo CatchAllError
    
    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)
                         
    InitialiseSettingsFile True
    sMACROWebPath = GetMACROSetting(MACRO_SETTING_WEBPATH, MACRO_SETTING_WEBPATHDEF)
    If (FileExists(sMACROWebPath & "img\_logo.gif")) Then sLogoPath = "../img/_logo.gif"
    GetAppMenuHeaderLhHTML = GetAppHeaderLh(oUser, sLogoPath, iwww)
    
    Set oUser = Nothing
    Exit Function

CatchAllError:
    Set oUser = Nothing
    GetAppMenuHeaderLhHTML = GetErrorRedirectAndLogError("clsWWW.GetAppMenuHeaderLhHTML", _
        Err.Number, Err.Description, Array())
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetAppMenuTopHTML(ByVal sSerialisedUser As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 06/03/2003
'   wrapper for modUIHTMLApplication.GetAppMenuTop
'   revisions
'   ic 30/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser

    On Error GoTo CatchAllError

    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)
                         
    GetAppMenuTopHTML = GetAppMenuTop(oUser, iwww)
    
    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    GetAppMenuTopHTML = GetErrorRedirectAndLogError("clsWWW.GetAppMenuTopHTML", Err.Number, Err.Description, _
    Array())
End Function

'--------------------------------------------------------------------------------------------------
Public Function SaveAndLoadAppMenuLh(ByRef vSerialisedUser As Variant, _
                                     ByRef vSplit As Variant, _
                                     ByVal sForm As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 05/11/2002
'   wrapper for modUIHTML.GetAppMenuHeader
'   revisions
'   ic 22/05/2003 fixed password bug 1796
'   ic 23/05/2003 removed update password
'   ic 30/06/2004 added error handling
'   dph 28/02/2007 bug 2882. allow for tracing in call to GetAppMenuLh
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser
Dim bRefresh As Boolean
Dim bTrace As Boolean

    On Error GoTo CatchAllError

    Set oUser = New MACROUser
    Call oUser.SetState(CStr(vSerialisedUser))
                
    '   dph 28/02/2007 bug 2882. allow for tracing in call to GetAppMenuLh
    bTrace = RtnTraceFlag()
                      
    If (sForm <> "") Then
        bRefresh = SaveUserSettings(oUser, sForm)
    End If
    SaveAndLoadAppMenuLh = GetAppMenuLh(oUser, iwww, bRefresh, , , bTrace)
    
    vSplit = CBool(oUser.UserSettings.GetSetting(SETTING_SPLIT_SCREEN, "false"))
    vSerialisedUser = oUser.GetState(False)
    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    SaveAndLoadAppMenuLh = GetErrorRedirectAndLogError("clsWWW.SaveAndLoadAppMenuLh", Err.Number, _
        Err.Description, Array(vSplit, sForm))
End Function

'--------------------------------------------------------------------------------------------------
Public Function CheckPermission(ByVal sSerialisedUser As String, _
                                ByVal sFunctionCode As String) As Boolean
'--------------------------------------------------------------------------------------------------
'   ic 11/11/2002
'   function checks permission held in passed serial user object
'   revisions
'   ic 30/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser

    On Error GoTo CatchAllError
    
    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)
                               
    CheckPermission = oUser.CheckPermission(sFunctionCode)
    
    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.CheckPermission")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetNewSubjectHTML(ByVal sSerialisedUser As String) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 11/11/2002
'   wrapper for ModUIHTML.GetNewSubject
'   revisions
'   ic 30/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser

    On Error GoTo CatchAllError

    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)

    If oUser.CheckPermission(gsFnDataEntry) Then
        If oUser.CheckPermission(gsFnCreateNewSubject) Then
            GetNewSubjectHTML = GetNewSubject(oUser)
        Else
            Call Err.Raise(vbObjectError + 1, , "User does not have create new subject permission")
        End If
    Else
        Call Err.Raise(vbObjectError + 1, , "User does not have data-entry permission")
    End If

    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    GetNewSubjectHTML = GetErrorRedirectAndLogError("clsWWW.GetNewSubjectHTML", Err.Number, _
        Err.Description, Array())
End Function

'--------------------------------------------------------------------------------------------------
Public Function ChangePasswordRequest(ByRef vSerialisedUser As Variant, ByVal sForm As String, _
    ByRef vErrors As Variant) As Boolean
'--------------------------------------------------------------------------------------------------
'   ic 23/05/2003
'   function changes a user password, called from asp when user requests change
'   ic 22/06/2004 added parameter checking, error handling
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser
Dim vForm As Variant
Dim sOldPassword As String
Dim sNewPassword As String
Dim vResult As Variant
Dim bOK As Boolean
    
    On Error GoTo CatchAllError

    bOK = False
    If (sForm <> "") Then
        Set oUser = New MACROUser
        Call oUser.SetState(CStr(vSerialisedUser))
        vForm = Split(sForm, "&")
        sOldPassword = ReplaceHTMLCodes(Split(vForm(0), "=")(1))
        sNewPassword = ReplaceHTMLCodes(Split(vForm(1), "=")(1))
        
        'validate input before use
        If Not ValidatePassword(sOldPassword) Then sOldPassword = ""
        If Not ValidatePassword(sNewPassword) Then sNewPassword = ""
        
        Set oUser = ChangePassword(oUser, sOldPassword, sNewPassword, True, vResult, vErrors)
        If (vResult) Then
            bOK = True
            vSerialisedUser = oUser.GetState(False)
        End If
        Set oUser = Nothing
    End If
    
    ChangePasswordRequest = bOK
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.ChangePasswordRequest")
End Function

'--------------------------------------------------------------------------------------------------
Public Function ChangePasswordForce(ByVal sUserName As String, _
                                    ByVal sForm As String, _
                                    ByRef vAcceptedNewPassword As Variant, _
                                    ByRef vResult As Variant, _
                                    ByRef vErrors As Variant) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 21/11/2002
'   function changes a user password. called from asp when user password expires
'   revisions
'   ic 19/08/2003 bug 1962, fixed incorrect split of form string
'   ic 22/06/2004 added parameter checking, error handling
'   ic 21/07/2004 changed Failed to LoginResult.Failed, bug 2348
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser
Dim sTemp As String
Dim sOld As String
Dim sNew As String

    On Error GoTo CatchAllError
    Set oUser = New MACROUser
    
    'extract old, new password from form
    sTemp = ReplaceHTMLCodes(Split(sForm, "=")(1))
    sOld = Split(sTemp, gsDELIMITER1)(0)
    sNew = Split(sTemp, gsDELIMITER1)(1)
    
    'validate input before use
    If Not ValidateUsername(sUserName) Then sUserName = ""
    If Not ValidatePassword(sOld) Then sOld = ""
    If Not ValidatePassword(sNew) Then sNew = ""
    
    'check old password, create user state in the process
    If (oUser.Login(GetSecurityCon(), sUserName, sOld, "", "MACRO Web Data Entry", "", False, "", "", False) <> LoginResult.Failed) Then
        'attempt change password
        Set oUser = ChangePassword(oUser, sOld, sNew, False, vResult, vErrors)
        If vResult Then vAcceptedNewPassword = sNew
    Else
        vErrors = AddToArray(vErrors, "Password update", "Old password incorrect")
    End If
    
    'return updated user state string
    ChangePasswordForce = oUser.GetState(True)
    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.ChangePasswordForce")
End Function


'--------------------------------------------------------------------------------------------------
Public Function SaveAndLoadSchedule(ByVal sSerialisedUser As String, ByVal sSiteCode As String, _
    ByVal sStudyId As String, ByVal sSubjectId As String, ByVal sNew As String, ByVal sUpdate As String, _
    ByVal sIdentifier As String, ByVal sTimezoneOffset As String, ByVal sUseSCI As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 09/12/2002
'   function saves any schedule updates (eg mimessages,status) and loads the schedule
'--------------------------------------------------------------------------------------------------
' REVISIONS
' DPH 27/05/2003 - collect any errors that may occur
' ic 30/06/2004 added parameter checking, error handling
' ic 24/08/2004 added setting planned sdvs to done
' ic 01/04/2005 issue 2541 added show existing sdv code and arguments
'--------------------------------------------------------------------------------------------------

Dim oUser As MACROUser
Dim vUpdate As Variant
Dim eScope As MIMsgScope
Dim bTimezoneOffsetOK As Boolean
Dim vErrors As Variant
Dim vAlerts As Variant
Dim bUseSCI As Boolean
Dim bNew As Boolean
Dim nPlannedToDone As Integer
Dim sMsg As String
Dim sSDVCall As String

    On Error GoTo CatchAllError

    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)
    
    If Not ValidateSite(sSiteCode) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal site parameter:" & sSiteCode)
    End If
    If Not IsNumeric(sStudyId) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal study parameter:" & sStudyId)
    End If
    If Not IsNumeric(sSubjectId) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal subjectid parameter:" & sSubjectId)
    End If
    
    If (sUpdate <> "") Then
        Select Case Left(sUpdate, 1)
        Case "s":
            'mimessage
            vUpdate = Split(sUpdate, gsDELIMITER1)
            Select Case vUpdate(1)
            Case "e": eScope = mimscEForm
            Case "v": eScope = mimscVisit
            Case "s": eScope = mimscSubject
            End Select
            Call RaiseMIMessageCall(oUser, eScope, vUpdate(0), sIdentifier, RtnTimezoneOffset(sTimezoneOffset, _
            bTimezoneOffsetOK), vErrors, True, sSDVCall)
        Case "u", "m":
            'unobtainable,missing
            bUseSCI = (LCase(sUseSCI) = "true")
            Call ToggleStatus(oUser, sIdentifier, Left(sUpdate, 1), bUseSCI, vErrors)
            'Call ToggleEFIStatus(oUser, sIdentifier, Left(sUpdate, 1), vErrors)
        Case "l"
            'lock/freeze data
            ' DPH 27/05/2003 - collect any errors that may occur
            Call ChangeLockUnlockFreeze(oUser, sIdentifier, vErrors)
        Case "d"
            bUseSCI = (LCase(sUseSCI) = "true")
            nPlannedToDone = PlannedSDVsToDone(oUser, bUseSCI, sIdentifier, vErrors)
            
            Select Case nPlannedToDone
            Case 0
                sMsg = "There were no Planned question SDVs on this eForm."
            Case 1
                sMsg = "One SDV was changed to Done."
            Case Else
                sMsg = nPlannedToDone & " SDVs were changed to Done."
            End Select
            
            vAlerts = AddToArray(vAlerts, "", sMsg)
        End Select
    End If
    
    bNew = (sNew = "1")
    SaveAndLoadSchedule = GetScheduleHTML(oUser, sSiteCode, CLng(sStudyId), CLng(sSubjectId), vErrors, vAlerts, bNew, sSDVCall)
    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    SaveAndLoadSchedule = GetErrorRedirectAndLogError("clsWWW.SaveAndLoadSchedule", Err.Number, _
        Err.Description, Array(sSiteCode, sStudyId, sSubjectId, sNew, sUpdate, sIdentifier, _
        sTimezoneOffset, sUseSCI))
End Function



'--------------------------------------------------------------------------------------------------
Public Function SaveAndLoadDataBrowser(ByVal sSerialisedUser As String, ByVal sStudyCode As String, _
                               ByVal sSiteCode As String, ByVal sVisitCode As String, ByVal sCRFPageId As String, _
                               ByVal sQuestion As String, ByVal sSrchUserName As String, ByVal sSubjectId As String, _
                               ByVal sSubjectLabel As String, ByVal sStatus As String, ByVal sLockStatus As String, _
                               ByVal sTime As String, ByVal sBefore As String, ByVal sComment As String, _
                               ByVal sDiscrepancy As String, ByVal sSDV As String, ByVal sNote As String, _
                               ByVal sType As String, sBookmark As String, ByVal sForm As String, _
                               ByVal sTimezoneOffset As String, ByVal sDecimalPoint As String, _
                               ByVal sThousandSeparator As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 11/12/2002
'   function saves databrowser updates (mimessages,locking etc) and loads data
'   revisions
'   ic 16/04/2003 fixed bug 1539: check correct permissions
'   DPH 27/05/2003 - collect any errors that may occur during locking
'   ic 28/05/2003 prevent users without monitor rights from seeing multiple subjects
'   ic 26/06/2003 added decimalpoint, thousandseparator to arguments, bug 1873
'   ic 28/06/2004 added parameter checking
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser
Dim bTimezoneOffsetOK As Boolean
Dim vErrors As Variant
Dim lSubjectId As Long
Dim sIdentifier As String
Dim sUType As String
Dim sAdd As String
Dim sScope As String
Dim vData As Variant
Dim sUSiteCode As String

    On Error GoTo CatchAllError

    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)
    
    'validate passed argument values
    If (Not IsNumeric(sStudyCode)) Then sStudyCode = "0"
    If Not ValidateSite(sSiteCode) Then sSiteCode = ""
    If Trim(sSiteCode) = "" Then sSiteCode = "ALL"
    If Trim(sSiteCode) = "ALL" Then
        sUSiteCode = ""
    Else
        sUSiteCode = Trim(sSiteCode)
    End If
    If (Not IsNumeric(sVisitCode)) Then sVisitCode = "0"
    If (Not IsNumeric(sCRFPageId)) Then sCRFPageId = "0"
    If (Not IsNumeric(sQuestion)) Then sQuestion = "0"
    If Not ValidateUsername(sSrchUserName) Then sSrchUserName = ""
    If Trim(sSrchUserName) = "" Then sSrchUserName = "ALL"
    If IsNumeric(sSubjectId) Then
        If (CLng(sSubjectId) > 0) Then
            lSubjectId = CLng(sSubjectId)
            sSubjectLabel = ""
        Else
            lSubjectId = 0
        End If
    Else
        lSubjectId = 0
    End If
    If Not ValidateLabel(sSubjectLabel) Then sSubjectLabel = ""
    If Not ValidateDateTime(sTime) Then sTime = ""
    If (sBefore = "1") Or (sBefore = "true") Then
        sBefore = "1"
    Else
        sBefore = "0"
    End If
    If Not IsNumeric(sComment) Then sComment = "-1"
    If Not IsNumeric(sDiscrepancy) Then sDiscrepancy = "-1"
    If Not IsNumeric(sSDV) Then sSDV = "-1"
    If Not IsNumeric(sNote) Then sNote = "-1"
    If (sType <> "0") And (sType <> "1") And (sType <> "2") And (sType <> "3") Then sType = "1"
    If Not IsNumeric(sBookmark) Then sBookmark = "0"
    
    
    If (sForm <> "") Then
        'sForm = ReplaceHTMLCodes(sForm)
        sIdentifier = ReplaceHTMLCodes(Split(Split(sForm, "&")(0), "=")(1))
        sUType = ReplaceHTMLCodes(Split(Split(sForm, "&")(1), "=")(1))
        sAdd = ReplaceHTMLCodes(Split(Split(sForm, "&")(2), "=")(1))
        sScope = ReplaceHTMLCodes(Split(Split(sForm, "&")(3), "=")(1))
        
        Select Case (sUType)
        Case "3", "4", "5", "6":
            ' DPH 27/05/2003 - collect any errors that may occur during locking
            Call UpdateLockFreeze(oUser, sIdentifier, sUType, sScope, vErrors)
        Case "0", "1", "2":
            Call RaiseMIMessageCall(oUser, CInt(sScope), sAdd, sIdentifier, RtnTimezoneOffset(sTimezoneOffset, bTimezoneOffsetOK), vErrors)
        End Select
    End If
    
    If (oUser.CheckPermission(gsFnViewData)) Then
        Select Case sType
        Case "0":
            'forms, currently unsupported
        Case "1", "2":
            'current data, audit trail
            If Not oUser.CheckPermission(gsFnMonitorDataReviewData) Then
                If (sSubjectLabel = "") Then
                    'no passed subject label and no monitor/review permission
                    Call Err.Raise(vbObjectError + 1, , "User does not have monitor data review permission and no label passed")
                Else
                    'passed subject label and no monitor/review permission
                    'check for multiple subjects
                    vData = oUser.DataLists.GetSubjectList(sSubjectLabel, oUser.Studies.StudyById(sStudyCode).StudyName, sUSiteCode)
                    If Not IsNull(vData) Then
                        If (UBound(vData, 2) > 0) Then
                            'more than 1 subject matches the search label
                            SaveAndLoadDataBrowser = GetMessageHTML("There is more than 1 subject that matches that subject label.  Pease refine your search", True, iwww)
                        Else
                            'only a single subject matches the passed search label
                            SaveAndLoadDataBrowser = GetDataBrowser(oUser, CInt(sStudyCode), sSiteCode, CLng(sVisitCode), CLng(sCRFPageId), _
                                CLng(sQuestion), sSrchUserName, lSubjectId, sSubjectLabel, sStatus, sLockStatus, sTime, CBool(sBefore), _
                                CLng(sComment), CLng(sDiscrepancy), CLng(sSDV), CLng(sNote), sType, sDecimalPoint, sThousandSeparator, iwww, _
                                CLng(sBookmark), vErrors)
                        End If
                    Else
                        'no subjects match the search label
                        SaveAndLoadDataBrowser = GetMessageHTML("Your query returned no records", True, iwww)
                    End If
                
                End If
            Else
                If (sType = "2") And (Not oUser.CheckPermission(gsFnViewAuditTrail)) Then
                    'trying to view audit trail without permission
                    Call Err.Raise(vbObjectError + 1, , "User does not have view audit trail permission")
                Else
                    SaveAndLoadDataBrowser = GetDataBrowser(oUser, CInt(sStudyCode), sSiteCode, CLng(sVisitCode), CLng(sCRFPageId), _
                        CLng(sQuestion), sSrchUserName, lSubjectId, sSubjectLabel, sStatus, sLockStatus, sTime, CBool(sBefore), _
                        CLng(sComment), CLng(sDiscrepancy), CLng(sSDV), CLng(sNote), sType, sDecimalPoint, sThousandSeparator, _
                        iwww, CLng(sBookmark), vErrors)
                End If
            End If
        Case "3":
            'changes since last session
            SaveAndLoadDataBrowser = GetDataBrowser(oUser, CInt(sStudyCode), sSiteCode, CLng(sVisitCode), CLng(sCRFPageId), _
                CLng(sQuestion), sSrchUserName, lSubjectId, sSubjectLabel, sStatus, sLockStatus, sTime, CBool(sBefore), _
                CLng(sComment), CLng(sDiscrepancy), CLng(sSDV), CLng(sNote), sType, sDecimalPoint, sThousandSeparator, _
                iwww, CLng(sBookmark), vErrors)
        End Select
        
    Else
        'no view data permission
        Call Err.Raise(vbObjectError + 1, , "User does not have view data permission")
    End If
    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    SaveAndLoadDataBrowser = GetErrorRedirectAndLogError("clsWWW.SaveAndLoadDataBrowser", Err.Number, _
        Err.Description, Array(sStudyCode, sSiteCode, sVisitCode, sCRFPageId, sQuestion, sSrchUserName, _
        lSubjectId, sSubjectLabel, sStatus, sLockStatus, sTime, sBefore, sComment, sDiscrepancy, _
        sSDV, sNote, sType, sDecimalPoint, sThousandSeparator, sBookmark))
End Function

'--------------------------------------------------------------------------------------------------
Public Function SaveAndLoadMIMessageList(ByVal sSerialisedUser As String, ByVal sType As String, _
    ByVal sStudyCode As String, ByVal sSiteCode As String, ByVal sVisitCode As String, ByVal sVisitCycle As String, _
    ByVal sCRFPageId As String, ByVal sCRFPageCycle As String, ByVal sQuestion As String, ByVal sQuestionCycle As String, _
    ByVal sSrchUserName As String, ByVal sSubjectId As String, ByVal sSubjectLabel As String, ByVal sStatus As String, _
    ByVal sTime As String, ByVal sBefore As String, ByVal sScope As String, ByVal sBookmark As String, _
    ByVal sForm As String, ByVal sTimezoneOffset As String, ByVal sNewWindow As String, ByVal sASPVToken As String, _
    ByVal sASPEToken As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 11/12/2002
'   function saves mimessage list, updates mimessages
'   revisions
'   ic 05/03/2004 added subjectid to UpdateMIMessage() call for subject locking
'   ic 01/07/2004 added parameter checking, error handling
'   ic 16/07/2004 fixed locking during update
'   ic 04/07/2005 issue 2464, added visit, eform, question cycle
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser
Dim bTimezoneOffsetOK As Boolean
Dim vErrors As Variant
Dim lSubjectId As Long
Dim sIdentifier As String
Dim vIdentifier As Variant
Dim sUType As String
Dim nType As String
Dim sAdd As String
Dim sAction As String
Dim bNewWindow As Boolean
Dim sUpdate As String

    On Error GoTo CatchAllError
    
    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)
    
    'validate parameters
    If (Not IsNumeric(sStudyCode)) Then sStudyCode = ""
    If Not ValidateSite(sSiteCode) Then sSiteCode = ""
    If (Not IsNumeric(sVisitCode)) And (sVisitCode <> "") Then sVisitCode = ""
    If (Not IsNumeric(sCRFPageId)) And (sCRFPageId <> "") Then sCRFPageId = ""
    If (Not IsNumeric(sQuestion)) And (sQuestion <> "") Then sQuestion = ""
    If Not ValidateUsername(sSrchUserName) Then sSrchUserName = ""
    If (Not IsNumeric(sSubjectId)) And (sSubjectId <> "") Then sSubjectId = ""
    If Not ValidateLabel(sSubjectLabel) Then sSubjectLabel = ""
    If Not IsNumeric(sStatus) Then sStatus = ""
    If Not ValidateDateTime(sTime) Then sTime = ""
    If LCase(sBefore) <> "true" Then sBefore = "false"
    If Not IsNumeric(sScope) Then sScope = ""
    If Not IsNumeric(sBookmark) Then sBookmark = "0"
    bNewWindow = (sNewWindow = "1")
    If Not IsNumeric(sVisitCycle) Then sVisitCycle = ""
    If Not IsNumeric(sCRFPageCycle) Then sCRFPageCycle = ""
    If Not IsNumeric(sQuestionCycle) Then sQuestionCycle = ""
    

    If (sForm <> "") Then
        'note: split on form delimiters ('&','=') before replacing codes
        sIdentifier = ReplaceHTMLCodes(Split(Split(sForm, "&")(0), "=")(1))
        sUType = ReplaceHTMLCodes(Split(Split(sForm, "&")(1), "=")(1))
        Select Case sUType
            Case "0": nType = mimtDiscrepancy
            Case "2": nType = mimtNote
            Case Else: nType = mimtSDVMark
        End Select
        sAdd = ReplaceHTMLCodes(Split(Split(sForm, "&")(2), "=")(1))
        vIdentifier = Split(sIdentifier, gsDELIMITER1)
        sAction = Split(sAdd, gsDELIMITER1)(0)
        sAdd = Split(sAdd, gsDELIMITER1)(1)
        
        'check parameters
        If Not ValidateStudyName(vIdentifier(0)) Then
            Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal study parameter:" & vIdentifier(0))
        End If
        If Not IsNumeric(vIdentifier(1)) Then
            Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal studyid parameter:" & vIdentifier(1))
        End If
        If Not IsNumeric(vIdentifier(16)) Then
            Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal subjectid parameter:" & vIdentifier(16))
        End If
        If (Not IsNumeric(vIdentifier(13))) Then
            Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal responsetaskid parameter:" & vIdentifier(13))
        End If
        If (Not IsNumeric(vIdentifier(14))) Then
            Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal responsecycle parameter:" & vIdentifier(14))
        End If
        If (Not IsNumeric(vIdentifier(10))) Then
            Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal id parameter:" & vIdentifier(10))
        End If
        If (Not IsNumeric(vIdentifier(11))) Then
            Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal src parameter:" & vIdentifier(11))
        End If
        If Not ValidateSite(vIdentifier(2)) Then
            Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal site parameter:" & vIdentifier(2))
        End If
        If Not ValidateText(sAdd) Then
            Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal text parameter:" & sAdd)
        End If
        
        'update mimessage
        Call UpdateMIMessage(oUser, nType, sAction, vIdentifier(0), CLng(vIdentifier(1)), CLng(vIdentifier(16)), _
            CLng(vIdentifier(13)), CInt(vIdentifier(14)), CLng(vIdentifier(10)), CInt(vIdentifier(11)), _
            vIdentifier(2), sAdd, RtnTimezoneOffset(sTimezoneOffset, bTimezoneOffsetOK), sASPVToken, sASPEToken, _
            sUpdate, vErrors)
    End If
    
    SaveAndLoadMIMessageList = GetMIMessageList(oUser, sType, sStudyCode, sSiteCode, sVisitCode, sVisitCycle, _
        sCRFPageId, sCRFPageCycle, sQuestion, sQuestionCycle, sSrchUserName, sSubjectId, sSubjectLabel, sStatus, _
        sTime, sBefore, sScope, bNewWindow, sUpdate, iwww, sBookmark, vErrors)
    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    SaveAndLoadMIMessageList = GetErrorRedirectAndLogError("clsWWW.SaveAndLoadMIMessageList", Err.Number, _
        Err.Description, Array(sType, sStudyCode, sSiteCode, sVisitCode, sVisitCycle, sCRFPageId, sCRFPageCycle, _
        sQuestion, sQuestionCycle, sSrchUserName, sSubjectId, sSubjectLabel, sStatus, sTime, sBefore, sScope, _
        bNewWindow, iwww, sBookmark))
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetUserDB(ByVal sSerialisedUser As String) As String
'--------------------------------------------------------------------------------------------------
'   revisions
'   ic 01/07/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser

    On Error GoTo CatchAllError

    Set oUser = New MACROUser
    oUser.SetState (sSerialisedUser)
    GetUserDB = oUser.DatabaseCode
    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetUserDB")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetUserRole(ByVal sSerialisedUser As String) As String
'--------------------------------------------------------------------------------------------------
'   revisions
'   ic 01/07/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser

    On Error GoTo CatchAllError

    Set oUser = New MACROUser
    oUser.SetState (sSerialisedUser)
    GetUserRole = oUser.UserRole
    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetUserRole")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetEformLabChoiceHTML(ByVal sSerialisedUser As String, ByVal sSite As String) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 03/04/2003
'   wrapper for ModUIHTMLEForm.GetEformLabChoice
'   revisions
'   ic 01/07/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser

    On Error GoTo CatchAllError

    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)

    GetEformLabChoiceHTML = GetEformLabChoice(oUser, sSite)

    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    GetEformLabChoiceHTML = GetErrorRedirectAndLogError("clsWWW.GetEformLabChoiceHTML", Err.Number, _
        Err.Description, Array(sSite))
End Function

'------------------------------------------------------------------------------'
Public Function RtnTraceFlag() As Boolean
'------------------------------------------------------------------------------'
' ic 24/04/2003
' function returns trace on/off boolean. default is 'false'
'   revisions
'   ic 01/07/2004 added error handling
'------------------------------------------------------------------------------'
    On Error GoTo CatchAllError
    
    RtnTraceFlag = RtnTraceFlagA()
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.RtnTraceFlag")
End Function

'------------------------------------------------------------------------------'
Public Function RtnPassThreadFlag() As Boolean
'------------------------------------------------------------------------------'
' ic 18/08/2003
' function returns passthread on/off boolean. default is 'false'
'   revisions
'   ic 01/07/2004 added error handling
'------------------------------------------------------------------------------'
Dim sTrace As String

    On Error GoTo CatchAllError

    InitialiseSettingsFile True
    sTrace = GetMACROSetting(MACRO_SETTING_PASSTHREAD, "false")
    RtnPassThreadFlag = (LCase(sTrace) = "true")
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.RtnPassThreadFlag")
End Function

'------------------------------------------------------------------------------'
Public Function GetDBLockAdminHTML(ByVal sSerialisedUser As String) As String
'------------------------------------------------------------------------------'
'   ic 21/05/2003
'   function gets db lock administration html
'   revisions
'   ic 01/07/2004 added error handling
'------------------------------------------------------------------------------'
Dim oUser As MACROUser

    On Error GoTo CatchAllError

    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)
    GetDBLockAdminHTML = GetDBLockAdmin(oUser, iwww)
    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    'cant return redirect script as modaldialog will open a new ie window
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetDBLockAdminHTML")
End Function

'----------------------------------------------------------------------------------------
Public Sub DeleteDBLocks(ByVal sSerialisedUser As String, ByVal sForm As String)
'----------------------------------------------------------------------------------------
'   ic 22/05/2003
'   function releases database locks
'   revisions
'   ic 01/07/2004 added error handling, parameter checking
'----------------------------------------------------------------------------------------
Dim oUser As MACROUser
Dim oDBLock As DBLock
Dim nLoop As Integer
Dim vForm As Variant
Dim vItem As Variant
Dim sIdentifier As String
Dim vIdentifier As Variant

    On Error GoTo CatchAllError
    
    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)
    
    If (sForm <> "") Then
        Set oDBLock = New DBLock
        
        vForm = Split(sForm, "&")
        For nLoop = LBound(vForm) To UBound(vForm)
            sIdentifier = ReplaceHTMLCodes(Split(vForm(nLoop), "=")(1))
            vIdentifier = Split(sIdentifier, gsDELIMITER1)
            
            If Not IsAlphanumeric(vIdentifier(0)) Then
                Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal token parameter:" & vIdentifier(0))
            End If
            If Not IsNumeric(vIdentifier(1)) Then
                Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal studyid parameter:" & vIdentifier(1))
            End If
            If Not ValidateSite(vIdentifier(2)) Then
                Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal site parameter:" & vIdentifier(2))
            End If
            If Not IsNumeric(vIdentifier(3)) Then
                Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal subjectid parameter:" & vIdentifier(3))
            End If
            If (Not IsNumeric(vIdentifier(4))) And vIdentifier(4) <> "" Then
                Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal eforminstanceid parameter:" & vIdentifier(3))
            End If
            
            If (vIdentifier(4) <> "") Then
                'user is deleting an eform lock
                oDBLock.UnlockEFormInstance oUser.CurrentDBConString, _
                        vIdentifier(0), vIdentifier(1), vIdentifier(2), vIdentifier(3), vIdentifier(4)
            ElseIf (vIdentifier(3) <> "") Then
                'user is deleting a subject lock
                oDBLock.UnlockSubject oUser.CurrentDBConString, _
                        vIdentifier(0), vIdentifier(1), vIdentifier(2), vIdentifier(3)
            Else
                'user is deleting a study lock
                oDBLock.UnlockStudy oUser.CurrentDBConString, _
                        vIdentifier(0), vIdentifier(1)
            End If
        Next
    
        Set oDBLock = Nothing
    End If
    Set oUser = Nothing
    Exit Sub
    
CatchAllError:
    Set oUser = Nothing
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.DeleteDBLocks")
End Sub

'--------------------------------------------------------------------------------------------------
Public Function RtnStatusImagesHTML(ByVal nStatus As Integer, _
                                ByVal bViewInformIcon As Boolean, _
                       Optional ByVal nLockStatus As Integer, _
                       Optional ByVal bUseMIEnum As Boolean = False, _
                       Optional ByVal nSDVStatus As Integer = 0, _
                       Optional ByVal nDiscStatus As Integer = 0, _
                       Optional ByVal bNote As Boolean = False, _
                       Optional ByVal bComment As Boolean = False, _
                       Optional ByVal nChanges As Integer = 0, _
                       Optional ByRef sStatusLabel As String, _
                       Optional ByRef sLockLabel As String) As String
'--------------------------------------------------------------------------------------------------
' RS 10/06/2003: Expose RtnStatusImages from modUIHTML
'   revisions
'   ic 01/07/2004 added error handling
'--------------------------------------------------------------------------------------------------
    On Error GoTo CatchAllError
                           
    RtnStatusImagesHTML = RtnStatusImages(nStatus, bViewInformIcon, nLockStatus, bUseMIEnum, nSDVStatus, nDiscStatus, bNote, bComment, nChanges, sStatusLabel, sLockLabel)
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.RtnStatusImagesHTML")
End Function

'--------------------------------------------------------------------------------------------------
Public Function RegisterWWWSubject(ByVal sSerialisedUser As String, ByVal sState As String, ByVal sUseSCI As String, _
    ByRef vMessage As Variant) As String
'--------------------------------------------------------------------------------------------------
'   ic 13/06/2003
'   function registers a subject using passed subject identifier state
'   revisions
'   ic 01/07/2004 added parameter checking, error handling
'   ic 18/05/2005 issue 2560, pass database connection for locating cached subjects
'--------------------------------------------------------------------------------------------------
Dim bUseSCI As Boolean
Dim oSCI As MACROSCI30.clsSubjectCacheInterface
Dim vSubject As Variant
Dim oUser As MACROUser
Dim oArezzo As Arezzo_DM
Dim vState As Variant
Dim bRegister As Boolean
Dim bLoadFail As Boolean
Dim sLoadFailMessage As String
Dim oSubject As StudySubject
Dim oStudyDef As StudyDefRO
Dim sMessage As String
Dim lErrNumber As Long
Dim sErrDescription As String
Dim bTrace As Boolean

    On Error GoTo CatchAllError

    bRegister = False
    bTrace = RtnTraceFlag()
    bUseSCI = (LCase(sUseSCI) = "true")
    
    Set oUser = New MACROUser
    Call oUser.SetState(sSerialisedUser)
    
    vState = Split(sState, gsDELIMITER2)
    
    'check parameters
    If Not IsNumeric(vState(1)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal study parameter:" & vState(1))
    End If
    If Not ValidateSite(vState(2)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal site parameter:" & vState(2))
    End If
    If Not IsNumeric(vState(3)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal subjectid parameter:" & vState(3))
    End If
    
    'ic 02/09/2003 use new CanChangeData function
    If (CanChangeData(oUser, vState(2)) And oUser.CheckPermission(gsFnRegisterSubject)) Then
        
        If bUseSCI Then
            Set oSCI = New MACROSCI30.clsSubjectCacheInterface
            Call WriteLog(bTrace, "clsWWW.RegisterWWWSubject * calling oSCI.LoadSubject(" & oUser.CurrentDBConString & "," & vState(1) & "," & vState(2) & "," & vState(3) & "," & oUser.UserName & "," & oUser.UserNameFull & "," & oUser.UserRole & ")")
            vSubject = oSCI.LoadSubject(oUser.CurrentDBConString, CLng(vState(1)), vState(2), CLng(vState(3)), oUser.UserName, eUIUpdateMode.Read_Write, oUser.UserNameFull, oUser.UserRole)
            If Not IsEmpty(vSubject) Then
                Call WriteLog(bTrace, "clsWWW.RegisterWWWSubject * oSCI.LoadSubject returned " & vSubject(0))
                If (vSubject(0) = 0) Then
                    Set oSubject = vSubject(1)
                    Set oStudyDef = oSubject.StudyDef
                Else
                    bLoadFail = True
                    sLoadFailMessage = vSubject(2)
                End If
            Else
                'error, cache returned nothing
                Call WriteLog(bTrace, "clsWWW.RegisterWWWSubject * oSCI.LoadSubject returned empty")
                bLoadFail = True
                sLoadFailMessage = "Subject cache returned empty"
            End If
        Else
            Set oArezzo = New Arezzo_DM
            Call oArezzo.Init(App.Path & "\Temp\", msPROLOG_SWITCHES)
            Set oStudyDef = New StudyDefRO
            Call oStudyDef.Load(oUser.CurrentDBConString, CLng(vState(1)), 1, oArezzo)
            Set oSubject = oStudyDef.LoadSubject(vState(2), CLng(vState(3)), oUser.UserName, eUIUpdateMode.Read_Write, oUser.UserNameFull, oUser.UserRole, False)
    
            If oSubject.CouldNotLoad Then
                bLoadFail = True
                sLoadFailMessage = oSubject.CouldNotLoadReason
                Call oSubject.StudyDef.RemoveSubject
                Set oSubject = Nothing
            End If
        End If
    
        If Not bLoadFail Then
            If (Not oSubject.ReadOnly) Then
                bRegister = DoRegistration(oSubject, 0, oUser.CurrentDBConString, oUser.DatabaseCode, sMessage)
                If (sMessage <> "") Then
                    sMessage = ReplaceWithHTMLCodes(sMessage)
                Else
                    sMessage = "This subject is not eligible for registration"
                End If
            Else
                sMessage = "This subject is read only. " & oSubject.ReadOnlyReason
            End If
        Else
            Call Err.Raise(vbObjectError + 1, , sLoadFailMessage)
        End If
        
        Set oSubject = Nothing
        If bUseSCI Then
            Call WriteLog(bTrace, "clsWWW.RegisterWWWSubject * calling oSCI.ReleaseSubject(" & vState(1) & "," & vState(2) & "," & vState(3) & "," & vSubject(2) & ")")
            If (vSubject(0) = 0) Then oSCI.ReleaseSubject oUser.CurrentDBConString, CLng(vState(1)), vState(2), CLng(vState(3)), vSubject(2)
            Set oSCI = Nothing
        Else
            oArezzo.Finish
            Set oArezzo = Nothing
            'ic 25/03/2004 MUST call the terminate method to free up memory
            oStudyDef.Terminate
            Set oStudyDef = Nothing
        End If
        vMessage = sMessage
        RegisterWWWSubject = bRegister
    Else
        Call Err.Raise(vbObjectError + 1, , "User does not have permission to register subject")
    End If
    Set oUser = Nothing
    Exit Function
    

CatchAllError:
    lErrNumber = Err.Number
    sErrDescription = Err.Description
    On Error Resume Next
    
    Call WriteLog(bTrace, "clsWWW.RegisterWWWSubject ->CatchAllError(" & Err.Number & "," & Err.Description & ")")
    Set oSubject = Nothing
    
    If bUseSCI Then
        If Not IsEmpty(vSubject) Then
            If (vSubject(0) = 0) Then oSCI.ReleaseSubject oUser.CurrentDBConString, CLng(vState(1)), vState(2), CLng(vState(3)), vSubject(2)
        Else
            Call WriteLog(bTrace, "clsWWW.RegisterWWWSubject ->XoSCI.ReleaseSubject(" & vState(1) & "," & vState(2) & "," & vState(3) & ",EMPTY)" & vbCrLf)
        End If
        Set oSCI = Nothing
    Else
        If IsObject(oArezzo) Then
            oArezzo.Finish
        End If
        Set oArezzo = Nothing
        If IsObject(oStudyDef) Then
            'ic 25/03/2004 MUST call the terminate method to free up memory
            oStudyDef.Terminate
        End If
        Set oStudyDef = Nothing
    End If
    Set oUser = Nothing
    Call Err.Raise(lErrNumber, , sErrDescription & "|clsWWW.RegisterWWWSubject")
End Function

'------------------------------------------------------------------------------'
Public Function GetWebVersionInfo() As Long
'------------------------------------------------------------------------------'
' dph 17/06/2003
' returns the web version from the settings file
'   revisions
'   ic 01/07/2004 added error handling
'------------------------------------------------------------------------------'
Dim sWebVersion As String

    On Error GoTo CatchAllError

    InitialiseSettingsFile True
    sWebVersion = GetMACROSetting(MACRO_SETTING_WEBVERSION, "0")
    GetWebVersionInfo = CLng(sWebVersion)
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetWebVersionInfo")
End Function

'----------------------------------------------------------------------------------------'
Public Function GetDatabaseChoiceHTML(ByVal sUser As String, ByVal sPassword As String, ByRef sDatabase As Variant, _
    ByRef sRole As Variant, ByVal sAppState As String) As String
'----------------------------------------------------------------------------------------'
'   ic 20/06/2003
'   function wrapper for GetDatabaseChoice()
'   revisions
'   ic 01/07/2004 added error handling
'----------------------------------------------------------------------------------------'
    On Error GoTo CatchAllError
    
    'validate parameters, ignore appstate, this will be validated
    'when it reaches appfrm.asp
    If Not ValidatePassword(sPassword) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal password parameter: ********")
    End If
    If Not ValidateDatabase(sDatabase) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal database parameter:" & sDatabase)
    End If
    If Not ValidateRole(sRole) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal role parameter:" & sRole)
    End If
    If Not ValidateAppState(sAppState) Then sAppState = ""
    

    GetDatabaseChoiceHTML = GetDatabaseChoice(GetSecurityCon(), sUser, sPassword, sDatabase, sRole, sAppState)
    Exit Function

CatchAllError:
    GetDatabaseChoiceHTML = GetErrorRedirectAndLogError("clsWWW.GetDatabaseChoiceHTML", _
        Err.Number, Err.Description, Array(sPassword, sDatabase, sRole, sAppState))
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetEformPrintHTML(ByRef sSerialisedUser As Variant, ByVal sSiteCode As String, ByVal sStudyCode As String, _
    ByVal sSubjectId As String, ByVal sVisitCode As String, ByVal sVEformId As String, ByVal sVTEformId As String, _
    ByVal sEformId As String, ByVal sTEformId As String, ByVal sDecimalPoint As String, ByVal sThousandSeparator As String) _
    As String
'--------------------------------------------------------------------------------------------------
'   ic 16/12/2003
'   function returns an eform html page suitable for printing through internet explorer
'   revisions
'   ic 01/07/2004 added error handling
'   ic 01/11/2004 bug 2435, allow sVEformId and sVTEformId to be empty strings
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser

    On Error GoTo CatchAllError

    Set oUser = New MACROUser
    Call oUser.SetState(CStr(sSerialisedUser))

    If Not ValidateSite(sSiteCode) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal site parameter:" & sSiteCode)
    End If
    If (Not IsNumeric(sStudyCode) And sStudyCode <> "") Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal studyid parameter:" & sStudyCode)
    End If
    If Not IsNumeric(sSubjectId) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal subjectid parameter:" & sSubjectId)
    End If
    If Not IsNumeric(sVisitCode) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal visitcode parameter:" & sVisitCode)
    End If
    If Not IsNumeric(sVEformId) And sVEformId <> "" Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal veformid parameter:" & sVEformId)
    End If
    If Not IsNumeric(sVTEformId) And sVTEformId <> "" Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal vteformid parameter:" & sVTEformId)
    End If
    If Not IsNumeric(sEformId) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal eformid parameter:" & sEformId)
    End If
    If Not IsNumeric(sTEformId) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal teformid parameter:" & sTEformId)
    End If

    GetEformPrintHTML = GetEformPrint(oUser, sSiteCode, sStudyCode, sSubjectId, sVisitCode, sVEformId, _
        sVTEformId, sEformId, sTEformId, sDecimalPoint, sThousandSeparator)

    Set oUser = Nothing
    Exit Function
    
CatchAllError:
    Set oUser = Nothing
    'cant return redirect script as modaldialog will open a new ie window
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.GetEformPrintHTML")
End Function

'--------------------------------------------------------------------------------------------------
Public Sub LogError(ByVal sLocation As String, ByVal sErrorCode As String, ByVal sErrorMessage As String, _
    ByVal vParams As Variant)
'--------------------------------------------------------------------------------------------------
'   ic 17/06/2004
'   function calls WriteErrorLog function
'--------------------------------------------------------------------------------------------------
    Call WriteErrorLog(sLocation, sErrorCode, sErrorMessage, vParams)
End Sub

'------------------------------------------------------------------------------'
Public Function RtnLogIllegalParametersFlag()
'------------------------------------------------------------------------------'
'   ic 29/06/2004
'   revisions
'------------------------------------------------------------------------------'
Dim sLog As String

    On Error GoTo CatchAllError
    
    RtnLogIllegalParametersFlag = RtnLogIllegalParametersFlagA()
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.RtnLogIllegalParametersFlag")
End Function

'--------------------------------------------------------------------------------------------------
Private Sub RaiseMIMessageCall(ByRef oUser As MACROUser, ByVal eScope As MIMsgScope, ByVal sMIMessageAdd As String, _
    ByVal sIdentifier As String, ByVal nTimezoneOffset As Integer, ByRef vMIErrors As Variant, _
    Optional ByVal bShowExistingSDV As Boolean = False, Optional ByRef sSDVCall As String = "")
'--------------------------------------------------------------------------------------------------
' ic 25/02/02
' function raises a mimessage by building business objects required before calling RaiseMIMessage()
' function. called from databrowser
' revisions
' ic 21/11/2002 updated with new business object functions
' DPH 27/11/2002 MIMessage saving for RQG questions (Response cycle incorporated)
' ic 25/03/2004 added oStudydef.Terminate call to free up memory
' ic 10/05/2004 invalidate cache entry so that new mimessage status is picked up
' ic 29/06/2004 added parameter checking, error handling. this function should be moved to modUIHTMLMIMessage
'               commented out previous mimessage fix - this will need to be reinstated
' ic 12/07/2004 reinstated fix
' ic 01/04/2005 issue 2541 added show existing sdv code and arguments
' NCJ 29 Dec 06 - Bug 2861 - Check response exists before creating MIMessage
'--------------------------------------------------------------------------------------------------
Dim oStudyDef As StudyDefRO
Dim oSubject As StudySubject
Dim oResponse As Response
Dim oEFI As EFormInstance
Dim oMDL As MIDataLists
Dim vDetails As Variant
Dim vItm As Variant
Dim eLoadOK As eLoadResponsesResult
Dim eSaveOK As eSaveResponsesResult
Dim nRepeatNo As Integer
Dim sToken As String
Dim lErrNumber As Long
Dim sErrDescription As String
    
    
    On Error GoTo CatchAllError
    
    vItm = Split(sIdentifier, gsDELIMITER1)

    If Not IsNumeric(vItm(eWWWIdentifier.idStudyId)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal studyid parameter:" & vItm(eWWWIdentifier.idStudyId))
    End If
    If Not ValidateSite(vItm(eWWWIdentifier.idSite)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal site parameter:" & vItm(eWWWIdentifier.idSite))
    End If
    If Not IsNumeric(vItm(eWWWIdentifier.idSubject)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal subjectid parameter:" & vItm(eWWWIdentifier.idSubject))
    End If

    'Create a new StudyDef class and load study def
    Set oStudyDef = New StudyDefRO
    Call oStudyDef.Load(oUser.CurrentDBConString, CLng(vItm(eWWWIdentifier.idStudyId)), 1)
    
    'Create a new moSubject class and load subject def
    Set oSubject = oStudyDef.LoadSubject(vItm(eWWWIdentifier.idSite), CLng(vItm(eWWWIdentifier.idSubject)), oUser.UserName, Read_Only, oUser.UserNameFull, oUser.UserRole, False)
    
    
    If oSubject.CouldNotLoad Then
        vMIErrors = AddToArray(vMIErrors, vItm(eWWWIdentifier.idStudyId) & "/" & vItm(eWWWIdentifier.idSite) & "/" & vItm(eWWWIdentifier.idSubject), "Could not load subject: " & oSubject.CouldNotLoadReason)
    Else
        sToken = LockSubject(oUser, oSubject.StudyId, oSubject.Site, oSubject.PersonId, vMIErrors)
        
        If (sToken <> "") Then
            If (eScope = mimscSubject) Then
                Call RaiseMIMessage(oUser, vMIErrors, sMIMessageAdd, eScope, nTimezoneOffset, oSubject.StudyDef.Name, oSubject.StudyId, _
                                    oSubject.Site, oSubject.PersonId, , , , , , , , , , , , , bShowExistingSDV, sSDVCall)
            Else
                If (eScope = mimscVisit) Then
                    If IsNumeric(vItm(eWWWIdentifier.idVisitId)) And IsNumeric(vItm(eWWWIdentifier.idVisitCycle)) Then
                        Call RaiseMIMessage(oUser, vMIErrors, sMIMessageAdd, eScope, nTimezoneOffset, oSubject.StudyDef.Name, oSubject.StudyId, _
                        oSubject.Site, oSubject.PersonId, CLng(vItm(eWWWIdentifier.idVisitId)), CInt(vItm(eWWWIdentifier.idVisitCycle)), , , , , , , , , , , bShowExistingSDV, sSDVCall)
                    End If
                ElseIf (eScope = mimscEForm) Then
                    If IsNumeric(vItm(eWWWIdentifier.idEformTaskId)) Then
                        'load this crfpagetaskid into oefi variable
                        Set oEFI = oSubject.eFIByTaskId(CLng(vItm(eWWWIdentifier.idEformTaskId)))
                            
                        Call RaiseMIMessage(oUser, vMIErrors, sMIMessageAdd, eScope, nTimezoneOffset, oSubject.StudyDef.Name, oSubject.StudyId, _
                            oSubject.Site, oSubject.PersonId, oEFI.VisitInstance.VisitId, oEFI.VisitInstance.CycleNo, _
                            oEFI.EFormTaskId, oEFI.eForm.EFormId, oEFI.CycleNo, , , , , , , , bShowExistingSDV, sSDVCall)
                                                
                        Set oEFI = Nothing
                    End If
                Else
                    If IsNumeric(vItm(eWWWIdentifier.idEformTaskId)) And IsNumeric(vItm(eWWWIdentifier.idResponseTaskId)) _
                    And IsNumeric(vItm(eWWWIdentifier.idResponseTaskId)) And IsNumeric(vItm(eWWWIdentifier.idResponseCycle)) Then
                        'load this crfpagetaskid into oefi variable
                        Set oEFI = oSubject.eFIByTaskId(CLng(vItm(eWWWIdentifier.idEformTaskId)))
                       
                        Set oMDL = New MIDataLists
                        vDetails = oMDL.GetResponseDetails(oUser.CurrentDBConString, oSubject.StudyCode, oSubject.Site, _
                           oSubject.PersonId, CLng(vItm(eWWWIdentifier.idResponseTaskId)), CLng(vItm(eWWWIdentifier.idResponseCycle)))
                        ' NCJ 29 Dec 06 - Bug 2861 - Check response exists
                        If Not IsNull(vDetails) Then
                            Call RaiseMIMessage(oUser, vMIErrors, sMIMessageAdd, eScope, nTimezoneOffset, oSubject.StudyDef.Name, _
                               oSubject.StudyId, oSubject.Site, oSubject.PersonId, oEFI.VisitInstance.VisitId, _
                               oEFI.VisitInstance.CycleNo, oEFI.EFormTaskId, oEFI.eForm.EFormId, oEFI.CycleNo, _
                               vItm(eWWWIdentifier.idResponseTaskId), CLng(vItm(eWWWIdentifier.idResponseCycle)), vDetails(eResponseDetails.rdResponseTimeStamp, 0), _
                               ConvertFromNull(vDetails(eResponseDetails.rdResponseValue, 0), vbString), vDetails(eResponseDetails.rdQuestionId, 0), _
                               vDetails(eResponseDetails.rdUserName, 0), , bShowExistingSDV, sSDVCall)
                        End If
                        Set oEFI = Nothing
                        Set oMDL = Nothing
                    End If
                End If
                
                
            End If
            
            'ic 10/05/2004 invalidate cache entry so that new mimessage status is picked up
            Call MACROLOCKBS30.CacheInvalidate(oUser.CurrentDBConString, CLng(vItm(eWWWIdentifier.idStudyId)), _
                CStr(vItm(eWWWIdentifier.idSite)), CLng(vItm(eWWWIdentifier.idSubject)))
            
            Call UnlockSubject(oUser, oSubject.StudyId, oSubject.Site, oSubject.PersonId, sToken, vMIErrors)
        Else
            'wasnt able to lock subject
        End If
    End If
    
    Set oSubject = Nothing
    'ic 25/03/2004 added oStudydef.Terminate call to free up memory
    Call oStudyDef.Terminate
    Set oStudyDef = Nothing
    Exit Sub
    
    
CatchAllError:
    lErrNumber = Err.Number
    sErrDescription = Err.Description
    
    On Error Resume Next
    Set oSubject = Nothing
    If IsObject(oStudyDef) Then
        'ic 25/03/2004 added oStudydef.Terminate call to free up memory
        Call oStudyDef.Terminate
    End If
    Set oStudyDef = Nothing
    Err.Raise lErrNumber, , sErrDescription & "|clsWWW.RaiseMIMessageCall"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub ChangeLockUnlockFreeze(ByRef oUser As MACROUser, _
                                    ByVal sIdentifier As String, _
                                    ByRef vErrors As Variant)
'--------------------------------------------------------------------------------------------------
' DPH 17/01/2003
' Change the lock/freeze status in the Browser/schedule for web
'--------------------------------------------------------------------------------------------------
' REVISIONS
' DPH 27/05/2003 - Removed possible problem of falling out of procedure
'               without giving an error message
'   ic 01/07/2004 added parameter checking, error handling
'   ic 11/07/2006 issue 2719 must invalidate cache after lock/freeze or changes are not
'                 shown until the subject is reloaded in the cache
'--------------------------------------------------------------------------------------------------
Dim sToken As String
Dim oLFObj As LFObject
Dim oFlocker As LockFreeze
Dim oLF As LockFreeze
Dim nSource As Integer
Dim enStatus As LockStatus

Dim vItem As Variant
Dim enAction As LFAction
Dim enScope As LFScope
Dim lStudyId As Long
Dim sStudyName As String
Dim sSite As String
Dim lSubjectId As Long
Dim lVisitId As Long
Dim nVisitCycleNumber As Integer
Dim lCRFPageTaskId As Long
Dim lCRFPageId As Long
Dim nCRFPageCycleNumber As Integer
Dim lResponseId As Long
Dim nResponseCycle As Integer
Dim lQuestionId As Long

    On Error GoTo CatchAllError

    ' Get basic variables
    vItem = Split(sIdentifier, gsDELIMITER1)
    
    If Not IsNumeric(vItem(0)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal studyid parameter:" & vItem(0))
    End If
    If Not ValidateSite(vItem(1)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal site parameter:" & vItem(1))
    End If
    If Not IsNumeric(vItem(2)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal subjectid parameter:" & vItem(2))
    End If
    
    lStudyId = vItem(0)
    sStudyName = RtnStudyName(oUser, lStudyId)
    sSite = vItem(1)
    lSubjectId = vItem(2)
    enScope = vItem(7)
    enAction = vItem(8)
    
    ' check if can Lock/Freeze
    Set oLF = New LockFreeze
    If Not oLF.CanLockFreezeOnServerWWW(oUser.CurrentDBConString, _
                                    sStudyName, sSite, lSubjectId) Then
        ' DPH 27/05/2003 - store error to display to user
        Call AddToArray(vErrors, lSubjectId, "Lock/freeze operations may not be carried out on this subject because there is unimported site data")
'        DialogInformation "Lock/freeze operations may not be carried out on this subject because there is unimported site data"
        Exit Sub
    End If
    Set oLF = Nothing

    'nb The subject could be locked by by the current user having the schedule open
    sToken = LockSubject(oUser, lStudyId, sSite, lSubjectId, vErrors)
    If sToken = "" Then
        ' this subject is currently locked (Message already given to User in LockSubject)
        Exit Sub
    End If

    Set oLFObj = New LFObject
    Select Case enScope
    Case LFScope.lfscSubject
        ' Apply to the Trial Subject
        Call oLFObj.Init(LFScope.lfscSubject, lStudyId, sSite, lSubjectId)
    Case LFScope.lfscVisit
        ' Apply to the Visit
        lVisitId = vItem(3)
        nVisitCycleNumber = vItem(4)
        Call oLFObj.Init(LFScope.lfscVisit, lStudyId, sSite, lSubjectId, _
                        lVisitId, nVisitCycleNumber)
    Case LFScope.lfscEForm
        ' Apply to the eForm
        lVisitId = vItem(3)
        nVisitCycleNumber = vItem(4)
        lCRFPageId = vItem(5)
        nCRFPageCycleNumber = vItem(6)
        Call oLFObj.Init(LFScope.lfscEForm, lStudyId, sSite, lSubjectId, _
                        lVisitId, nVisitCycleNumber, _
                        lCRFPageId, nCRFPageCycleNumber)
    Case LFScope.lfscQuestion
        ' Apply to the Question
        lVisitId = vItem(3)
        nVisitCycleNumber = vItem(4)
        lCRFPageId = vItem(5)
        nCRFPageCycleNumber = vItem(6)
        lResponseId = vItem(9)
        nResponseCycle = vItem(10)
        lQuestionId = vItem(11)
        Call oLFObj.Init(LFScope.lfscQuestion, lStudyId, sSite, lSubjectId, _
                        lVisitId, nVisitCycleNumber, _
                        lCRFPageId, nCRFPageCycleNumber, _
                        lResponseId, nResponseCycle, lQuestionId)
    End Select

    ' Get the LockFreeze object to do all the work
    Set oFlocker = New LockFreeze
    nSource = TypeOfInstallation.Server
    Call oFlocker.DoLockFreezeWWW(oUser.CurrentDBConString, oLFObj, enAction, nSource, _
                                oUser.UserName, oUser.UserNameFull)

    'ic 11/07/2006 issue 2719 must invalidate cache after lock/freeze or changes are not
    'shown until the subject is reloaded in the cache
    Call MACROLOCKBS30.CacheInvalidate(oUser.CurrentDBConString, lStudyId, sSite, lSubjectId)

    ' unlock the subject (this is the database lock!)
    Call UnlockSubject(oUser, lStudyId, sSite, lSubjectId, sToken, vErrors)

    Set oLFObj = Nothing
    Set oFlocker = Nothing
    Exit Sub

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.ChangeLockUnlockFreeze")
End Sub

'--------------------------------------------------------------------------------------------------
Private Function UpdateLockFreeze(ByRef oUser As MACROUser, _
                                  ByVal sItem As String, _
                                  ByVal sType As String, _
                                  ByVal sScope As String, _
                                  ByRef vErrors As Variant) As Integer
'--------------------------------------------------------------------------------------------------
' New Lock Freeze Handler for DataBrowser
'
' revisions
' ic 24/01/2003 made private, changed to pass user object, use identifier enum
' DPH 27/05/2003 added error array as parameter
'   ic 01/07/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sIdentifier As String
Dim enAction As LFAction
Dim vData As Variant

    On Error GoTo CatchAllError

    Select Case sType
        Case "3":
            enAction = lfaLock
        Case "4":
            enAction = lfaUnlock
        Case "5":
            enAction = lfaFreeze
        Case "6":
            enAction = lfaUnfreeze
        Case Else:
    End Select
    
    ' Order data for ChangeLockUnlockFreeze function
    vData = Split(sItem, gsDELIMITER1)
    
    'ic quick fix to ensure subscript wont be out of range
    ReDim Preserve vData(12)
    
    ' lStudyId / sSite / lSubjectId / lVisitId / nVisitCycleNumber / EformId
    sIdentifier = vData(eWWWIdentifier.idStudyId) & gsDELIMITER1 & vData(eWWWIdentifier.idSite) _
    & gsDELIMITER1 & vData(eWWWIdentifier.idSubject) & gsDELIMITER1 & vData(eWWWIdentifier.idVisitId) _
    & gsDELIMITER1 & vData(eWWWIdentifier.idVisitCycle) & gsDELIMITER1 & vData(eWWWIdentifier.idEformId) _
    & gsDELIMITER1
    
    ' Eform Cycle / LFScope / LFAction
    sIdentifier = sIdentifier & vData(eWWWIdentifier.idEformCycle) & gsDELIMITER1 & sScope & gsDELIMITER1 _
    & enAction & gsDELIMITER1
    ' lResponseId / nResponseCycle / lQuestionId
    sIdentifier = sIdentifier & vData(eWWWIdentifier.idResponseTaskId) & gsDELIMITER1 & vData(eWWWIdentifier.idResponseCycle) _
    & gsDELIMITER1 & vData(eWWWIdentifier.idResponseId)

    Call ChangeLockUnlockFreeze(oUser, sIdentifier, vErrors)
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.UpdateLockFreeze")
End Function

'--------------------------------------------------------------
Private Function ToggleStatus(ByRef oUser As MACROUser, _
                                 ByVal sIdentifier As String, _
                                 ByVal sToStatus As String, _
                                 ByVal bUseSCI As Boolean, _
                                 ByRef vErrors As Variant) As Boolean
'--------------------------------------------------------------
'Make an eForm unobtainable by changing all its 'Missing' responses to 'Unobtainable'
' or vice versa
'   revisions
'   ic 29/06/2004 added error handling, moved from clswww
'   ic 18/05/2005 issue 2560, pass database connection for locating cached subjects
'   dph 02/10/2006 bug 2794 - visit eform not set to unobtainable when toggled from schedule
'--------------------------------------------------------------
Dim oSCI As MACROSCI30.clsSubjectCacheInterface
Dim oArezzo As Arezzo_DM
Dim oStudyDef As StudyDefRO
Dim oSubject As StudySubject
Dim oEFI As EFormInstance
Dim oResponse As Response
Dim nToStatus As Integer
Dim nOldStatus As Integer
Dim sLockErrMsg As String
Dim vItm As Variant
Dim bLoadFail As Boolean
Dim nUpdate As eUIUpdateMode
Dim vSubject As Variant
Dim bSaveResponses As Boolean
Dim bOldStatusMissingRequested As Boolean
Dim sToggleLevel As String
Dim oVI As VisitInstance
Dim colVI As Collection
Dim nI As Integer
Dim bTrace As Boolean
Dim sLoadFailMessage As String
Dim lErrNumber As Long
Dim sErrDescription As String
Dim lVEFITaskId As Long
Dim bDoVisitEForm As Boolean
    
    On Error GoTo CatchAllError
    bTrace = RtnTraceFlagA()
    
    vItm = Split(sIdentifier, gsDELIMITER1)
    
    If Not IsNumeric(vItm(0)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal studyid parameter:" & vItm(0))
    End If
    If Not ValidateSite(vItm(1)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal site parameter:" & vItm(1))
    End If
    If Not IsNumeric(vItm(2)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal subjectid parameter:" & vItm(2))
    End If
    
    sToggleLevel = vItm(7)
    ToggleStatus = True
    
    If bUseSCI Then
        Set oSCI = New MACROSCI30.clsSubjectCacheInterface
        Call WriteLog(bTrace, "clsWWW.ToggleStatus ->oSCI.LoadSubject(" & oUser.DatabaseCode & "," & vItm(0) & "," & vItm(1) & "," & vItm(2) & "," & oUser.UserName & "," & nUpdate & "," & oUser.UserNameFull & "," & oUser.UserRole & ")")
        vSubject = oSCI.LoadSubject(oUser.CurrentDBConString, CLng(vItm(0)), vItm(1), CLng(vItm(2)), oUser.UserName, nUpdate, oUser.UserNameFull, oUser.UserRole)
        If Not IsEmpty(vSubject) Then
            Call WriteLog(bTrace, "clsWWW.ToggleStatus <-oSCI.LoadSubject returned " & vSubject(0))
            If (vSubject(0) = 0) Then
                Set oSubject = vSubject(1)
                Set oStudyDef = oSubject.StudyDef
            Else
                bLoadFail = True
                sLoadFailMessage = vSubject(2)
            End If
        Else
            'error, cache returned nothing
            Call WriteLog(bTrace, "clsWWW.ToggleStatus <-oSCI.LoadSubject EMPTY")
            bLoadFail = True
            sLoadFailMessage = "Subject cache returned empty"
        End If
    Else
        Set oArezzo = New Arezzo_DM
        Call oArezzo.Init(App.Path & "\Temp\", msPROLOG_SWITCHES)
        Set oStudyDef = New StudyDefRO
        Call oStudyDef.Load(oUser.CurrentDBConString, CLng(vItm(0)), 1, oArezzo)
        Set oSubject = oStudyDef.LoadSubject(vItm(1), CLng(vItm(2)), oUser.UserName, Read_Write, oUser.UserNameFull, oUser.UserRole, False)
        
        If oSubject.CouldNotLoad Then
            bLoadFail = True
            sLoadFailMessage = oSubject.CouldNotLoadReason
            Call oSubject.StudyDef.RemoveSubject
            Set oSubject = Nothing
        End If
    End If
    
    If Not bLoadFail Then
        
        'choose the statuses to look for
        If sToStatus = "m" Then
            nToStatus = eStatus.Missing
            nOldStatus = eStatus.Unobtainable
            bOldStatusMissingRequested = False
        Else
            nToStatus = eStatus.Unobtainable
            nOldStatus = eStatus.Missing
            bOldStatusMissingRequested = True
        End If
        
        Select Case sToggleLevel
            Case "s":
                ' subject
                ' loop through visits
                For Each oVI In oSubject.VisitInstances
                   If oVI.LockStatus = eLockStatus.lsUnlocked Then
                       If ((bOldStatusMissingRequested) And ((oVI.Status = eStatus.Missing) Or (oVI.Status = eStatus.Requested))) _
                           Or ((Not bOldStatusMissingRequested) And (oVI.Status = eStatus.Unobtainable)) Then
                            ' DPH 02/10/2006 - Bug 2794, Need to process visit eform with first eform processed (not as separate entity)
                            ' collect eformtaskid of visit eform and do not process as part of loop
                            lVEFITaskId = 0
                            If Not oVI.VisitEFormInstance Is Nothing Then
                                lVEFITaskId = oVI.VisitEFormInstance.EFormTaskId
                            End If
                            ' only process visiteform on first cycle
                            bDoVisitEForm = True
                            ' loop through eForms
                            For Each oEFI In oVI.eFormInstances
                                If oEFI.LockStatus = eLockStatus.lsUnlocked Then
                                    If ((bOldStatusMissingRequested) And ((oEFI.Status = eStatus.Missing) Or (oEFI.Status = eStatus.Requested))) _
                                        Or ((Not bOldStatusMissingRequested) And (oEFI.Status = eStatus.Unobtainable)) Then
                                            If Not ToggleEFIStatus(oSubject, oEFI, oUser, nToStatus, nOldStatus, _
                                                                    bOldStatusMissingRequested, vErrors, bDoVisitEForm) Then
                                                ToggleStatus = False
                                            End If
                                    End If
                                End If
                                bDoVisitEForm = False
                            Next
                       End If
                   End If
                Next
           Case "v":
                ' visit
                If Not IsNumeric(vItm(3)) Then
                    Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal visitid parameter:" & vItm(3))
                End If
                If Not IsNumeric(vItm(4)) Then
                    Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal visitcycle parameter:" & vItm(4))
                End If
                
                Set colVI = oSubject.VisitInstancesById(CLng(vItm(3)))
                'Set oVI = oSubject.VisitInstanceByTaskId(CLng(vItm(3)))
                For nI = 1 To colVI.Count
                    Set oVI = colVI(nI)
                    If oVI.CycleNo = CInt(vItm(4)) Then
                        ' DPH 02/10/2006 - Bug 2794, Need to process visit eform with first eform processed (not as separate entity)
                        ' collect eformtaskid of visit eform and do not process as part of loop
                        lVEFITaskId = 0
                        If Not oVI.VisitEFormInstance Is Nothing Then
                            lVEFITaskId = oVI.VisitEFormInstance.EFormTaskId
                        End If
                        ' only process visiteform on first cycle
                        bDoVisitEForm = True
                        ' loop through eForms
                        For Each oEFI In oVI.eFormInstances
                            If oEFI.LockStatus = eLockStatus.lsUnlocked And (oEFI.EFormTaskId <> lVEFITaskId) Then
                                If ((bOldStatusMissingRequested) And ((oEFI.Status = eStatus.Missing) Or (oEFI.Status = eStatus.Requested))) _
                                    Or ((Not bOldStatusMissingRequested) And (oEFI.Status = eStatus.Unobtainable)) Then
                                        If Not ToggleEFIStatus(oSubject, oEFI, oUser, nToStatus, nOldStatus, _
                                                                bOldStatusMissingRequested, vErrors, bDoVisitEForm) Then
                                            ToggleStatus = False
                                        End If
                                End If
                            End If
                            bDoVisitEForm = False
                        Next
                    End If
                Next
                
            Case "e":
                ' eform
                If Not IsNumeric(vItm(6)) Then
                    Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal eformtaskid parameter:" & vItm(6))
                End If
                
                'load this crfpagetaskid into oefi variable
                Set oEFI = oSubject.eFIByTaskId(CLng(vItm(6)))
                If Not ToggleEFIStatus(oSubject, oEFI, oUser, nToStatus, nOldStatus, _
                                        bOldStatusMissingRequested, vErrors, True) Then
                    ToggleStatus = False
                End If
                
          Case Else:
        End Select
    Else
        Call Err.Raise(vbObjectError + 1, , sLoadFailMessage)
    End If
    
    
    Set oResponse = Nothing
    Set oEFI = Nothing
    Set oVI = Nothing
    Set oSubject = Nothing
    If bUseSCI Then
        Call WriteLog(bTrace, "clsWWW.ToggleStatus ->oSCI.ReleaseSubject(" & vItm(0) & "," & vItm(1) & "," & vItm(2) & "," & vSubject(2) & ")" & vbCrLf)
        If (vSubject(0) = 0) Then oSCI.ReleaseSubject oUser.CurrentDBConString, CLng(vItm(0)), vItm(1), CLng(vItm(2)), vSubject(2)
        Set oSCI = Nothing
    Else
        oArezzo.Finish
        Set oArezzo = Nothing
        'ic 25/03/2004 MUST call the terminate method to free up memory
        oStudyDef.Terminate
        Set oStudyDef = Nothing
    End If
    Exit Function
    
CatchAllError:
    lErrNumber = Err.Number
    sErrDescription = Err.Description
    On Error Resume Next
    
    Call WriteLog(bTrace, "clsWWW.ToggleStatus ->CatchAllError(" & Err.Number & "," & Err.Description & ")")
    Set oResponse = Nothing
    Set oEFI = Nothing
    Set oSubject = Nothing
    If bUseSCI Then
        If Not IsEmpty(vSubject) Then
            Call WriteLog(bTrace, "clsWWW.ToggleStatus ->oSCI.ReleaseSubject(" & vItm(0) & "," & vItm(1) & "," & vItm(2) & "," & vSubject(2) & ")" & vbCrLf)
            If (vSubject(0) = 0) Then oSCI.ReleaseSubject oUser.CurrentDBConString, CLng(vItm(0)), vItm(1), CLng(vItm(2)), vSubject(2)
        Else
            Call WriteLog(bTrace, "clsWWW.ToggleStatus ->XoSCI.ReleaseSubject(" & vItm(0) & "," & vItm(1) & "," & vItm(2) & ",EMPTY)" & vbCrLf)
        End If
        Set oSCI = Nothing
    Else
        If IsObject(oArezzo) Then
            oArezzo.Finish
        End If
        Set oArezzo = Nothing
        If IsObject(oStudyDef) Then
            'ic 25/03/2004 MUST call the terminate method to free up memory
            oStudyDef.Terminate
        End If
        Set oStudyDef = Nothing
    End If
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.ToggleStatus")
End Function


'--------------------------------------------------------------------------------------------------
Private Function PlannedSDVsToDone(ByRef oUser As MACROUser, ByVal bUseSCI As Boolean, sIdentifier As String, _
    ByRef vErrors As Variant) As Integer
'--------------------------------------------------------------------------------------------------
' ic 24/08/2004
' function sets all planned sdvs to done on a specified eform
' revisions
' ic 20/12/2004 moved from modUIHTMLMIMessage to avoid a cycling reference
' ic 18/05/2005 issue 2560, pass database connection for locating cached subjects
' NCJ 1 Nov 05 - Issue 2635 - Add R/W flag to UserNameFull
'--------------------------------------------------------------------------------------------------
Dim oSCI As MACROSCI30.clsSubjectCacheInterface
Dim oArezzo As Arezzo_DM
Dim oStudyDef As StudyDefRO
Dim oSubject As StudySubject
Dim oEFI As EFormInstance

Dim vItm As Variant
Dim bLoadFail As Boolean
Dim nUpdate As eUIUpdateMode
Dim vSubject As Variant

Dim bTrace As Boolean
Dim sLoadFailMessage As String
Dim lErrNumber As Long
Dim sErrDescription As String

' NCJ 1 Nov 05
Dim sUserNameFull As String

    
    On Error GoTo CatchAllError
    bTrace = RtnTraceFlagA()
    
    vItm = Split(sIdentifier, gsDELIMITER1)
    
    If Not IsNumeric(vItm(0)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal studyid parameter:" & vItm(0))
    End If
    If Not ValidateSite(vItm(1)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal site parameter:" & vItm(1))
    End If
    If Not IsNumeric(vItm(2)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal subjectid parameter:" & vItm(2))
    End If
        If Not IsNumeric(vItm(8)) Then
        Call Err.Raise(vbObjectError + 2, , "Output terminated. Illegal eformtaskid parameter:" & vItm(8))
    End If
    
    ' NCJ 1 Nov 05
    ' Set up sUserNameFull
    If CanChangeData(oUser, vItm(1)) Then
        sUserNameFull = oUser.UserNameFull
    Else
        ' Disallow AREZZO events
        sUserNameFull = oUser.UserNameFull & "|0"
    End If
    
    If bUseSCI Then
        Set oSCI = New MACROSCI30.clsSubjectCacheInterface
        Call WriteLog(bTrace, "clsWWW.PlannedSDVsToDone ->oSCI.LoadSubject(" & oUser.DatabaseCode & "," _
                        & vItm(0) & "," & vItm(1) & "," & vItm(2) & "," & oUser.UserName & "," _
                        & nUpdate & "," & sUserNameFull & "," & oUser.UserRole & ")")
        vSubject = oSCI.LoadSubject(oUser.CurrentDBConString, CLng(vItm(0)), vItm(1), CLng(vItm(2)), _
                    oUser.UserName, nUpdate, sUserNameFull, oUser.UserRole)
        If Not IsEmpty(vSubject) Then
            Call WriteLog(bTrace, "clsWWW.PlannedSDVsToDone <-oSCI.LoadSubject returned " & vSubject(0))
            If (vSubject(0) = 0) Then
                Set oSubject = vSubject(1)
                Set oStudyDef = oSubject.StudyDef
            Else
                bLoadFail = True
                sLoadFailMessage = vSubject(2)
            End If
        Else
            'error, cache returned nothing
            Call WriteLog(bTrace, "clsWWW.PlannedSDVsToDone <-oSCI.LoadSubject EMPTY")
            bLoadFail = True
            sLoadFailMessage = "Subject cache returned empty"
        End If
    Else
        Set oArezzo = New Arezzo_DM
        Call oArezzo.Init(App.Path & "\Temp\", msPROLOG_SWITCHES)
        Set oStudyDef = New StudyDefRO
        Call oStudyDef.Load(oUser.CurrentDBConString, CLng(vItm(0)), 1, oArezzo)
        Set oSubject = oStudyDef.LoadSubject(vItm(1), CLng(vItm(2)), oUser.UserName, Read_Write, _
                        sUserNameFull, oUser.UserRole, False)
        
        If oSubject.CouldNotLoad Then
            bLoadFail = True
            sLoadFailMessage = oSubject.CouldNotLoadReason
            Call oSubject.StudyDef.RemoveSubject
            Set oSubject = Nothing
        End If
    End If
    
    If Not bLoadFail Then
        Set oEFI = oSubject.eFIByTaskId(CLng(vItm(8)))
        PlannedSDVsToDone = ChangePlannedQSDVsToDone(oUser.CurrentDBConString, oUser.UserName, oUser.UserNameFull, _
            oSubject, oEFI, mimsServer)
        Set oEFI = Nothing
        
    Else
        Call Err.Raise(vbObjectError + 1, , sLoadFailMessage)
    End If
    
    
    
    Set oSubject = Nothing
    If bUseSCI Then
        Call WriteLog(bTrace, "clsWWW.PlannedSDVsToDone ->oSCI.ReleaseSubject(" & vItm(0) & "," & vItm(1) & "," & vItm(2) & "," & vSubject(2) & ")" & vbCrLf)
        If (vSubject(0) = 0) Then oSCI.ReleaseSubject oUser.CurrentDBConString, CLng(vItm(0)), vItm(1), CLng(vItm(2)), vSubject(2)
        Set oSCI = Nothing
    Else
        oArezzo.Finish
        Set oArezzo = Nothing
        'ic 25/03/2004 MUST call the terminate method to free up memory
        oStudyDef.Terminate
        Set oStudyDef = Nothing
    End If
    Exit Function
    
CatchAllError:
    lErrNumber = Err.Number
    sErrDescription = Err.Description
    On Error Resume Next
    
    Call WriteLog(bTrace, "clsWWW.PlannedSDVsToDone ->CatchAllError(" & Err.Number & "," & Err.Description & ")")
    Set oEFI = Nothing
    Set oSubject = Nothing
    If bUseSCI Then
        If Not IsEmpty(vSubject) Then
            Call WriteLog(bTrace, "clsWWW.PlannedSDVsToDone ->oSCI.ReleaseSubject(" & vItm(0) & "," & vItm(1) & "," & vItm(2) & "," & vSubject(2) & ")" & vbCrLf)
            If (vSubject(0) = 0) Then oSCI.ReleaseSubject oUser.CurrentDBConString, CLng(vItm(0)), vItm(1), CLng(vItm(2)), vSubject(2)
        Else
            Call WriteLog(bTrace, "clsWWW.PlannedSDVsToDone ->XoSCI.ReleaseSubject(" & vItm(0) & "," & vItm(1) & "," & vItm(2) & ",EMPTY)" & vbCrLf)
        End If
        Set oSCI = Nothing
    Else
        If IsObject(oArezzo) Then
            oArezzo.Finish
        End If
        Set oArezzo = Nothing
        If IsObject(oStudyDef) Then
            'ic 25/03/2004 MUST call the terminate method to free up memory
            oStudyDef.Terminate
        End If
        Set oStudyDef = Nothing
    End If
    Call Err.Raise(Err.Number, , Err.Description & "|clsWWW.PlannedSDVsToDone")
End Function


