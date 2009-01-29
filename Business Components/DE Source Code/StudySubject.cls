VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudySubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'----------------------------------------------------
' File: StudySubject.cls
' Copyright InferMed Ltd 2001-2004
' Author: Nicky Johns, InferMed, May 2001
' Purpose: Study Subject class for MACRO 2.2 (now for 3.0)
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 18-29 May 01 - Initial development
' NCJ 15 Jun 01 - Added Arezzo & patient state
' NCJ 25 Jun 01 - Changed LoadResponses to accept string from DS
' TA  07 Jul 01 - Saving routines
' NCJ 30 Aug 01 - UpdateStatus; other work
' NCJ 3 Sep 01 - Include final Save in NewSubject
' TA 5/9/01 - Connection string now stored on load and new and is then used for every subsequent operation
' NCJ 14 Sep 01 - Added some comments; standardised RegistrationID and RegistrationStatus
' NCJ 17 Sep - More commenting; made Registration properties Get/Let; SaveRegistration
' NCJ 20 Sep 01 - Ensure StudySubject Changed flag is set when saving responses
' NCJ 24 Sep 01 - Added UserName property to store user name of current MACRO user
'               Added Terminate method
' NCJ 26 Sep 01 - Update Subject Label when setting RegistrationID
'               CancelResponses
' NCJ 27 Sep 01 - Added ReadOnly and ReadOnlyReason
' NCJ 28 Sep 01 - Added check in SaveResponses for new empty form
' NCJ 10 Oct 01 - Added oefi.Cancel in CancelResponses
' TA  12 Oct 01 - Proforma state is only retrieved from dataservices if we have an Arezzo
' NCJ 13 Feb 02 - Need to look for new eForms in Load
' NCJ 20 Mar 02 - Ensure correct behaviour of study updates for Locked/Frozen subjects
' NCJ 21 Mar 02 - Added UpdateMode as argument to Load
' NCJ 27 Mar 02 - Only look for new forms in Load if we have Arezzo
' TA 25/04/2002 - Only load EFIs that are still in the study definition
' NCJ 8 May 02 - Added support for Multi User Mode MUSaveAllLoadedResponses
' NCJ 13-15 May 02 - Further conditionally compiled stuff for Multi User Mode
' NCJ 11 Jun 02 - CBB 2.2.14/17 Make sure we run Arezzo whenever data's saved
' TA 20/08/2002: added new MIMessage statuses
' MLM 29/08/02: Changed CreateNewEFI for new form and visit dates
' MLM 04/09/02: Load and remove the visit eform's responses along with the main eform's
' MLM 10/09/02: Also save the visit eform's responses.
' NCJ 17 Sept 02 - Changes for new locking and MultiUser stuff
' NCJ 20 Sept 02 - Ensure eForm and Question contexts correctly reset for Subject's AREZZO activities
' NCJ 26 Sept 02 - Bug fixing Multi User; added new Reload function
' NCJ 26 Nov 02 - Removed calls to RunEngine (we now only need GenerateArezzoEvents)
' NCJ 16 Jan 03 - Added GetFirsteFormInNextVisit
' NCJ 17 Jan 03 - Added GetFirstVisitForm
'           Fixed ScheduleVisits to have eForms in correct order
' RS/NCJ 21 Jan 03 - Added properties LocalDateFormat & LocalTimeFormat
' NCJ 21 Jan 03 - We won't use TimeFormats
' NCJ 23 Jan 03 - Amended GeteFormInOtherVisit to match on eForm cycle if poss.
' NCJ 24 Jan 03 - Added sCountry (the site's country) to NewSubject; added SetUserProperties
' NCJ 27 Jan 03 - Added EFIState synchronisation for multi-user in SetUpCollections
' NCJ 28 Jan 03 - Added UserNameFull and UserRole properties
' NCJ 11 Feb 03 - Handle situation where VisitEForm is loaded on its own (as in Batch Data Entry)
'               And handle situation when first eForm in a visit is inactive
' NCJ 13 Feb 03 - Make sure we don't try and display the Visit eForm when moving between visits
' NCJ 28 Feb 03 - Same as 13 Feb, but did it properly this time!
' NCJ 5 Mar 03 - Do not remove responses of EFIs we want to keep in Me.Terminate
' NCJ 7 Mar 03 - Added Timestamp properties
' NCJ 26-28 Mar 03 - New Save routine to save subject data with appropriate locking
' NCJ 1 Apr 03 - No longer need to refresh visit date when saving
' NCJ 2 Apr 03 - Make sure Registration Identifiers are not munged by Prolog
' NCJ 28 Apr 03 - Bugs 1597 & 1598 : New eForms should be added for Frozen visits (because in MACRO 3.0 you can Unfreeze)
' NCJ 6 May 03 - Added new parameters sUserNameFull and sUserRole to Load and NewSubject
' NCJ 21 May 03 - Ignore requested cycling visits in UpdateStatus (BUG 1595)
' NCJ 27 May 03 - Added extra LockForSave in NewSubject; improved error handling in LoadResponses & SaveResponses
' NCJ 11 Jun 03 - Bug 1840 - Tidied up reload in Save routine
' NCJ 25 Mar 04 - Better Terminate routine which explicitly terminates relevant objects
' NCJ 26 Mar 04 - Further tweak to Termination (don't kill moStudy until the last minute)
' NCJ 27 May 04 - Bug 2282 - Have more respect for colKeepEFIs in Terminate
' ic 09/09/2005   added clinical coding
' NCJ 24 Nov 05 - Patched in changes from 3.0.72a (Bug 2635 - R/0 patients in Web DE)
' NCJ 8 Dec 05 - Partial dates stuff: deal with new set of Date/Time types
' NCJ 25 Jan 06 - Set Changed flag after adding AREZZO data in NewSubject (Bug 2677)
' NCJ 28 Jun 06 - Added bRemoveLock to CancelResponses (as in RemoveResponses) for Web DE
' NCJ 19 Mar 07 - Issue 2872 - Added HiddenLabWarnings, read from Settings file
'----------------------------------------------------

'----------------------------------------------------
' NOTES
' When a new eForm instance or Visit Instance is created, we must
'   - Set mcolScheduleVisits and moScheduleGrid to Nothing (to force a refresh)
'   - For a new VisitInstance, add it to mcolVisitInstances
'   - For a new eFormInstance, add it to mcolEFormInsts AND to appropriate VisitInstance
'   - For a new eFormInstance, add an item to mcolIsAnEFI
'
'----------------------------------------------------
'

Option Explicit

' NCJ 21 Mar 02 - The update mode for the subject
' determines whether data can be changed
Public Enum eUIUpdateMode
    Read_Write = 0
    Read_Only = 1
End Enum

'enum to store constants for reasons for load/save failure
Public Enum eLoadSaveErrors
    lseNoLockForSaving = -1
End Enum

' NCJ 17 Sept 02 Result of trying to save responses
Public Enum eSaveResponsesResult
    srrSuccess = 0
    srrNoLockForSaving = 1
    srrSubjectReloaded = 2
End Enum

' NCJ 17 Sept 02 Result of trying to save responses
Public Enum eLoadResponsesResult
    lrrReadWrite = 0
    lrrReadOnly = 1
    lrrCouldNotLockForSave = 2
End Enum

' The subject state class
Private moSubjectState As StudySubjectState

' The Study Definition
Private moStudy As StudyDefRO

' The Arezzo instance we're using
Private WithEvents moArezzo As Arezzo_DM
Attribute moArezzo.VB_VarHelpID = -1

' The collection of VisitInstance objects
Private mcolVisitInstances As Collection

' The collection of ScheduleVisit objects
' (Set to Nothing when refresh is required)
Private mcolScheduleVisits As Collection
' The Schedule Grid object
' (Set to Nothing when refresh is required)
Private moScheduleGrid As ScheduleGrid

' The collection of eFormInstances for this subject
Private mcolEFormInsts As Collection
' The collection indicating which eFormId/Cycle exist as instances
Private mcolIsAnEFI As Collection

'store connection string used for load
Private msConnection As String

' The name of the user who opened this subject this session
Private msUserName As String
Private msUserNameFull As String
Private msUserRole As String

Private mbArezzoStateErr As Boolean
Private msArezzoErr As String

'store could not load reason
Private mbCouldNotLoad As Boolean
Private msCouldNotLoadReason As String

' NCJ 21 Mar 02
Private mnUpdateMode As Integer

Private msCacheToken As String

' RS 22/9/2002
Private moTimezone As Timezone

' RS 20/01/2003
Private msLocalDateFormat As String
'Private msLocalTimeFormat As String

' NCJ 19 Mar 07 - Issue 2872 - Do they want warnings on hidden lab questions?
Private mbHiddenLabWarnings As Boolean

'notify subject label has changed
Public Event SubjectLabelChanged(sOldLabel As String, sNewLabel As String)


'---------------------------------------------------------------------
Friend Property Get HiddenLabWarnings() As Boolean
'---------------------------------------------------------------------
' NCJ 19 Mar 07 - Issue 2872 - Do they want warnings on hidden lab questions?
'---------------------------------------------------------------------

    HiddenLabWarnings = mbHiddenLabWarnings

End Property

'---------------------------------------------------------------------
Public Property Get LocalDateFormat() As String
'---------------------------------------------------------------------
' The user's preferred date format when doing Data Entry
' Empty if we are to use the Question's predefined format
'---------------------------------------------------------------------
    
    LocalDateFormat = msLocalDateFormat

End Property

'---------------------------------------------------------------------
Public Property Let LocalDateFormat(sLocalDateFormat As String)
'---------------------------------------------------------------------
    
    msLocalDateFormat = sLocalDateFormat

End Property

'----------------------------------------------------------------------------------
Friend Function LocalDateTypeFormat(enDateTimeType As eDateTimeType)
'----------------------------------------------------------------------------------
' NCJ 21 Jan 03
' Return the local format string for the given DateTimeType
' Returns empty string if no relevant local format
' NCJ 8 Dec 05 - Deal with new set of Date/Time types
'----------------------------------------------------------------------------------

    Select Case enDateTimeType
    
'    Case eDateTimeType.dttDateOnly: ' Date type
    Case eDateTimeType.dttYMD, eDateTimeType.dttMDY, eDateTimeType.dttDMY   ' Date type
        LocalDateTypeFormat = msLocalDateFormat
        
'    Case eDateTimeType.dttTimeOnly: ' Time type
    Case eDateTimeType.dttT         ' Time type
        ' We don't do time formats
'        LocalDateTypeFormat = msLocalTimeFormat
        
'    Case eDateTimeType.dttDateAndTime: ' Combined Date/Time
    Case eDateTimeType.dttYMDT, eDateTimeType.dttMDYT, eDateTimeType.dttDMYT   ' Combined Date/Time
        If msLocalDateFormat <> "" Then
            ' We add on a "standard" time format
            LocalDateTypeFormat = msLocalDateFormat & " hh:mm:ss"
        End If
        
    End Select

End Function

'---------------------------------------------------------------------
Public Sub SetUserProperties(sUserNameFull As String, sRole As String)
'---------------------------------------------------------------------
' NCJ 24 Jan 03 - Set the details of the current user
' NCJ 24 Nov 05 - Bug 2635 - Check there's nothing on the end of sUserNameFull
'---------------------------------------------------------------------
Dim vRW As Variant

    ' NCJ 31 Oct 05 - Check the sUserNameFull
    vRW = Split(sUserNameFull, "|")
    If UBound(vRW) > 0 Then
        'We found something on the end - assume it's 0 or 1
        sUserNameFull = Left(sUserNameFull, Len(sUserNameFull) - 2)
    End If

    msUserNameFull = sUserNameFull
    msUserRole = sRole
    Call moArezzo.SetUserProperties(sUserNameFull, sRole)

End Sub

'---------------------------------------------------------------------
Public Property Get CouldNotLoad() As Boolean
'---------------------------------------------------------------------

    CouldNotLoad = mbCouldNotLoad

End Property

'---------------------------------------------------------------------
Public Property Get CouldNotLoadReason() As String
'---------------------------------------------------------------------

    CouldNotLoadReason = msCouldNotLoadReason
    
End Property

'---------------------------------------------------------------------
Private Function LockForSave(sCon As String, ByRef sLockErrMsg As String, ByRef sToken As String, ByRef bNeedToReloadSubject As Boolean, _
                                sUser As String, lStudyId As Long, sSite As String, lSubjectId As Long) As Boolean
'---------------------------------------------------------------------
' Lock a subject.
' In:
'   sUser: username to lock subject with
'   lStudyId, sSite, lSubjectId: key of subject to lock
' Out:
'   function: success/failure
'   sToken: token if lock successful or "" if not
'   sLockErrMsg: the reason the lock failed or "" if successful
'   nNeedToReloadSubject: if the lock is successful this returns whehther the subject needs reloading
'---------------------------------------------------------------------
Dim sLockDetails As String
Const sSTUDY_BEING = "This study is currently being "
Const sSUBJECT_BEING = "This subject is currently being "
Const sEDITED = "edited by "
Const sSAVED = "saved by "
Const sANOTHER_USER = "another user"

    On Error GoTo ErrorLabel
    
    'set initial output variables to failure (change it later if we have success)
    sToken = ""
    sLockErrMsg = ""
    'set initial function output to failure (change it later if we have success)
    LockForSave = False
    
    'TA 04.07.2001: use new locking
    sToken = MACROLOCKBS30.LockSubjectForSaving(sCon, sUser, lStudyId, sSite, lSubjectId)
    Select Case sToken
    Case MACROLOCKBS30.DBLocked.dblStudy
        sLockDetails = MACROLOCKBS30.LockDetailsStudy(sCon, lStudyId)
        sLockErrMsg = sSTUDY_BEING & sEDITED
        If sLockDetails = "" Then
            sLockErrMsg = sLockErrMsg & sANOTHER_USER
        Else
            sLockErrMsg = sLockErrMsg & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblSubject
        sLockDetails = MACROLOCKBS30.LockDetailsSubject(sCon, lStudyId, sSite, lSubjectId)
        sLockErrMsg = sSUBJECT_BEING & sEDITED
        If sLockDetails = "" Then
            sLockErrMsg = sLockErrMsg & sANOTHER_USER
        Else
            sLockErrMsg = sLockErrMsg & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblEFormInstance
        sLockDetails = MACROLOCKBS30.LockDetailsSubjectSave(sCon, lStudyId, sSite, lSubjectId)
        sLockErrMsg = sSUBJECT_BEING & sSAVED
        If sLockDetails = "" Then
            sLockErrMsg = sLockErrMsg & sANOTHER_USER
        Else
            sLockErrMsg = sLockErrMsg & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case Else
        'we have a lock we need to check wether it is invalid to decide whether to reload or not
        bNeedToReloadSubject = Not MACROLOCKBS30.CacheEntryStillValid(msConnection, msCacheToken)
        'function to return success
        LockForSave = True
    End Select
    Exit Function
    
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.LockForSave"

End Function

'---------------------------------------------------------------------
Private Sub UnlockForSave(sCon As String, sToken As String, _
                            lStudyId As Long, sSite As String, lSubjectId As Long)
'---------------------------------------------------------------------
' Unlock the subject
'---------------------------------------------------------------------

    On Error GoTo ErrorLabel
    
    'TA 04.07.2001: use new locking model
    If sToken <> "" Then
        'if no gsStudyToken then UnlockSubject is being called without a corresponding LockSubject being called first
        MACROLOCKBS30.UnlockSubjectForSaving sCon, sToken, lStudyId, sSite, lSubjectId
        'always set this to empty string for same reason as above
        sToken = ""
    End If
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.UnlockForSave"

End Sub

'-------------------------------------------------------------------
Friend Sub Load(sConnection As String, oStudy As StudyDefRO, _
                    ByVal sSite As String, ByVal lPersonID As Long, _
                    ByVal sUserName As String, _
                    ByVal nUpdateMode As eUIUpdateMode, _
                    ByVal bUpdateCachetable As Boolean, _
                    ByVal sUserNameFull As String, _
                    ByVal sUserRole As String, _
                    Optional oArezzo As Arezzo_DM = Nothing)
'-------------------------------------------------------------------
' This is only called by the StudyDefRO class
' Load an existing subject specified by given study, site and personid
' sConnection is the DB connection string
' oArezzo is the (initialised) Arezzo object we can use for this subject
' NCJ 22 Mar 02 - Added nUpdateMode argument
' TA 09/04/2002: Parameter to update cache table added
' NCJ 6 May 03 - Added new parameters sUserNameFull and sUserRole
' NCJ 24 Nov 05 - sUserNameFull tells us if subject is REALLY R/W (bug 2635)
' NCJ 19 Mar 07 - Added call to ReadSettingsFile
'-------------------------------------------------------------------
Dim oStudySubjDS As New StudySubjectDS
Dim bLoadArezzo As Boolean
Dim sToken As String
Dim bNeedToReloadSubject As Boolean
Dim sLockErrMsg As String
Dim bMustNotLoadSubject As Boolean
Dim enLockResult As DBLocked

Dim bWritable As Boolean
Dim vRW As Variant

    On Error GoTo ErrorLabel
    
    'TODO do we want a terminate subject in here? - it is currently in studydef.loadsubject
    
    'store connection string
    msConnection = sConnection
    
    Me.UserName = sUserName
    
    ' Store the study definition
    Set moStudy = oStudy
    ' Store the Arezzo instance
    Set moArezzo = oArezzo
    
    ' NCJ 22 Mar 02 - Include the UpdateMode
    mnUpdateMode = nUpdateMode
    
    'override the provided update mode if oArezzo not supplied (maybe we need not do this)
    If oArezzo Is Nothing Then
        mnUpdateMode = eUIUpdateMode.Read_Only
    End If
    
    'get lock - the reason for lock failure will be passed back in msCouldNotLoadReason to the calling function
    mbCouldNotLoad = Not LockForSave(msConnection, msCouldNotLoadReason, sToken, bNeedToReloadSubject, sUserName, _
                                            oStudy.StudyId, sSite, lPersonID)
    
    'we do not need to check bNeedToReloadSubject as we are loading the subject
        
    If mbCouldNotLoad Then
'EXIT SUB HERE
        Exit Sub
    End If
    
    'only load AREZZO if we are read/write
    bLoadArezzo = (mnUpdateMode = eUIUpdateMode.Read_Write)
    
    ' Set the subject state
    moSubjectState.State = oStudySubjDS.Load(sConnection, _
                        oStudy.StudyId, sSite, lPersonID, bLoadArezzo, bUpdateCachetable)
    Set oStudySubjDS = Nothing

    ' NCJ 24 Nov 05 - See if there's a R/W flag appended to end of sUserNameFull
    ' Default to writable = true
    bWritable = True
    vRW = Split(sUserNameFull, "|")
    If UBound(vRW) > 0 Then
        'We found something on the end - assume it's 0 or 1
        bWritable = (vRW(0) = 1)
        sUserNameFull = Left(sUserNameFull, Len(sUserNameFull) - 2)
    End If
    
    mbArezzoStateErr = False
    msArezzoErr = ""
    If bLoadArezzo Then
        ' Load the patient state into Arezzo
        On Error Resume Next
        moArezzo.PatientState = moSubjectState.ArezzoState
        If Err.Number <> 0 Then
            mbArezzoStateErr = True
            msArezzoErr = "There was an error in loading this subject's data." _
                            & vbCrLf & Err.Description
        Else
            ' NCJ 6 May 03 - Set these properties here
            Call SetUserProperties(sUserNameFull, sUserRole)
        End If
        On Error GoTo 0
    End If
    
    ' NCJ 18 Sept 02 - Don't need to keep any existing EFIs
    Call SetUpCollections(Nothing)
    
    ' NCJ 13 Feb 02 - Need to look for new eForms (if not read-only)
    ' NCJ 27 Mar 02 - But only if we have Arezzo!
    ' NCJ 24 Nov 05 - Check our "Writable" flag too
'    If Not Me.ReadOnly Then
    If (Not Me.ReadOnly) And bWritable Then
    
'        Call moArezzo.RunEngine
        Call moArezzo.GenerateArezzoEvents
        If Me.Changed Then
            ' Save any new eForms which might have arrived
            Call SaveData
        End If
    
    End If
    
    ' NCJ 19 Mar 07 - Issue 2872 - Read Lab Questions setting
    Call ReadSettingsFile
    
    'get the cache token to return to the caller - this must be done for a save
    msCacheToken = MACROLOCKBS30.CacheAddSubjectRow(msConnection, Me.StudyDef.StudyId, Me.Site, Me.PersonID)
    
    'we must have a token if we have got this far
    Call UnlockForSave(msConnection, sToken, Me.StudyDef.StudyId, Me.Site, Me.PersonID)
    
Exit Sub

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.Load"
    Exit Sub
    
End Sub

'-------------------------------------------------------------------
Private Sub ReadSettingsFile()
'-------------------------------------------------------------------
' NCJ 19 Mar 07 - Read MACRO settings file and set things accordingly
'-------------------------------------------------------------------

    ' bUsingWWWDll = TRUE because we're a DLL!
    Call InitialiseSettingsFile(True)
    mbHiddenLabWarnings = (LCase(GetMACROSetting(MACRO_SETTING_HIDDEN_LAB_WARNINGS, "true")) = "true")

End Sub

'-------------------------------------------------------------------
Friend Sub NewSubject(sConnection As String, _
                oStudy As StudyDefRO, ByVal sSite As String, sUserName As String, _
                sCountry As String, _
                sUserNameFull As String, _
                sUserRole As String, _
                Optional oArezzo As Arezzo_DM = Nothing)
'-------------------------------------------------------------------
' Set ourselves up as a new subject for given Study, Site.
' PersonID will be automatically generated
' nb We watch for a lock error but this should never happen
' NCJ 24 Jan 03 - Pass in the site country as well
' NCJ 6 May 03 - Added new parameters sUserNameFull and sUserRole
' NCJ 27 May 03 - Added LockForSave around creation of SubjectID
' NCJ 25 Jan 06 - Set Changed flag after adding AREZZO data
' NCJ 19 Mar 07 - Added call to ReadSettingsFile
'-------------------------------------------------------------------
Dim sLockErrMsg As String
Dim sToken As String
Dim bNeedToReloadSubject As Boolean
Dim bMustNotLoadSubject As Boolean
Dim sDummyToken As String

    On Error GoTo ErrorLabel
    
    'store connection string
    msConnection = sConnection
    
    'set this up first - needed for save
    Me.UserName = sUserName
    
    ' Store the study definition
    Set moStudy = oStudy
    ' Store the Arezzo instance
    Set moArezzo = oArezzo
    
    ' Create new subject state
    Set moSubjectState = New StudySubjectState
    ' Initialise stuff
    With moSubjectState
        .StudyId = oStudy.StudyId
        .Site = sSite
        .PersonID = NULL_LONG   ' Indicate this is a new subject
        .Label = ""
        .LockStatus = eLockStatus.lsUnlocked
        .Status = eStatus.Requested
        .IsNew = True
        .IsChanged = True
        .Gender = GenderCode.gNone
        .DOB = eMACRODate.mdUnspecified
        .RegStatus = 0  ' ???
        .RegID = ""
        
        'TA 20/08/2002: added new MIMessage statuses
        .DiscrepancyStatus = 0
        .SDVStatus = 0
        .NoteStatus = 0
        
        Set .VisitInstances = New Collection
        Set .eFormInstances = New Collection
    
    End With

    ' Start the subject off in the Arezzo engine
    Call moArezzo.StartNewSubject
    
    'NCJ 27 May 03 - Added LockForSave around subject creation (using dummy PersonID = -1)
    mbCouldNotLoad = Not LockForSave(msConnection, msCouldNotLoadReason, sDummyToken, bNeedToReloadSubject, sUserName, _
                                        Me.StudyDef.StudyId, sSite, -1)
    If mbCouldNotLoad Then
        'this must be a study lock so can't create new subject
'EXIT SUB HERE
        Exit Sub
    End If
    
    'TODO use a sequence to generate the id
    moSubjectState.PersonID = SaveData
    
    ' Now re-lock the subject for saving with its known ID before releasing the -1 lock
    ' (This shouldn't fail!)
    mbCouldNotLoad = Not LockForSave(msConnection, msCouldNotLoadReason, sToken, bNeedToReloadSubject, sUserName, _
                                        Me.StudyDef.StudyId, sSite, Me.PersonID)
    
    'we do not need to check bNeedToReloadSubject as we are creating this subject
    
    ' Now release the lock with dummy ID
    Call UnlockForSave(msConnection, sDummyToken, Me.StudyDef.StudyId, Me.Site, -1)

    If mbCouldNotLoad Then
        ' this must be a study lock - this shouldn't happen here...
'EXIT SUB HERE
        Exit Sub
    End If
    
    ' Add person's data to Arezzo
    ' NCJ 24 Jan 03 - Added country
    Call moArezzo.AddData("person:clinicaltrialname", oStudy.Name)
    Call moArezzo.AddData("person:trialsite", sSite)
    Call moArezzo.AddData("person:personid", CStr(moSubjectState.PersonID))
    Call moArezzo.AddData("person:country", sCountry)
    ' NCJ 6 May 03
    Call SetUserProperties(sUserNameFull, sUserRole)
    
    ' NCJ 25 Jan 06 - Must set Changed flag (Bug 2677)
    moSubjectState.IsChanged = True
    
    ' Generate all the events which result from new eForm/Visit instances
    Call moArezzo.GenerateArezzoEvents
    
    ' Now save all the eForm/visit instances that will have been created
    ' (This assumes that all the Arezzo events have been processed
    ' but experience suggests that they always will have been)

    Call SaveData
    
    ' NCJ 19 Mar 07 - Issue 2872 - Read Lab Questions setting
    Call ReadSettingsFile
    
    'get the cache token
    msCacheToken = MACROLOCKBS30.CacheAddSubjectRow(msConnection, Me.StudyDef.StudyId, Me.Site, Me.PersonID)
    
    'we must have a token if we have got this far
    Call UnlockForSave(msConnection, sToken, Me.StudyDef.StudyId, Me.Site, Me.PersonID)
    
Exit Sub

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.NewSubject"
    Exit Sub
    
End Sub

'--------------------------------------------
Public Sub SaveCodedResponses(oEFI As EFormInstance)
'--------------------------------------------
' ic 13/09/2005
' saves coded terms
'--------------------------------------------
Dim sErr As String
Dim sToken As String
     
    If (LockForSave(msConnection, sErr, sToken, False, Me.UserName, Me.StudyDef.StudyId, Me.Site, Me.PersonID)) Then
        Call oEFI.SaveCodedTermHistories(msConnection)
        Call UnlockForSave(msConnection, sToken, Me.StudyDef.StudyId, Me.Site, Me.PersonID)
        sToken = ""
    End If
End Sub

'--------------------------------------------
Public Function SaveResponses(oEFI As EFormInstance, ByRef sLockErrMsg As String, _
                        Optional nTimezoneOffset As Integer = 0) As eSaveResponsesResult
'--------------------------------------------
' Save the responses for the particular eForm instance
' AND the subject with all its eForm and Visit instances.
' Only saves if there are changes to responses, or it's a new form with no responses.
' This automatically updates all labels, dates and statuses as necessary.
'this returns true if we save or did not need to save and false if failed because of locking
'- the lock reason is returned in sLockErrMsg
'--------------------------------------------
' i.e.
'   Update eForm instance Date & Label
'   Update Visit Date
'   Update Subject Label, DOB and Gender
'   Update status of eForm instance (based on Response statuses)
'   Update status of Visit instance (based on eForm instance statuses)
'   Update status of Subject (based on Visit instance statuses)
' Then the entire Subject is saved
'--------------------------------------------
' NCJ 11 Jun 02 - CBB 2.2.14/17 Make sure we run Arezzo whenever data's saved
' MLM 10/09/02: Save the Visit EForm as well, if required.
' MLM 13/09/02: Tidied up; use new SaveEFIResponses()
' NCJ 17 Sept 02 - Return result as eSaveResponsesResult
' NCJ 18 Sept 02 - If reloading, keep a copy of this eForm
' RS 22/9/2002 -    Added optional nTimezoneOffset parameter (to be used by web client)
' NCJ 21 May 03 - Ensure eForm label is saved if necessary
' NCJ 27 May 03 - Undo lock if necessary if error occurs
'--------------------------------------------

Dim oVisitEFI As EFormInstance
Dim bSaveResponses As Boolean
Dim sToken As String
Dim bNeedToReloadSubject As Boolean
Dim bCouldLockSubjectForSave As Boolean
Dim lEFITaskId As Long
Dim colKeepEFIs As Collection
Dim bReloadVisitEFIResponses As Boolean
Dim lErrNo As Long
Dim sErrDesc As String

    On Error GoTo ErrorLabel
    
    sToken = ""
    
    'we want to default to success , we will always set this to something else when we find a problem
    SaveResponses = srrSuccess
    
    ' NCJ 20 Mar 02 - Extra check to ensure we don't save if a read-only subject
    If Me.ReadOnly Then Exit Function
    
    'Main EForm
    bSaveResponses = DoEFIResponsesNeedSaving(oEFI)
        
    'Get the visit EForm instance (if any)
    Set oVisitEFI = GetVisitEFI(oEFI)
    If Not oVisitEFI Is Nothing Then
        If DoEFIResponsesNeedSaving(oVisitEFI) Then
            bSaveResponses = True
        End If
    End If
    
    If bSaveResponses Then
    
        'get lock - if the lock fails it must be "locked for saving" as the study and subject could not be locked at this stage
         bCouldLockSubjectForSave = LockForSave(msConnection, sLockErrMsg, sToken, bNeedToReloadSubject, Me.UserName, _
                                Me.StudyDef.StudyId, Me.Site, Me.PersonID)
        
        If bCouldLockSubjectForSave And bNeedToReloadSubject Then
        
            ' Reload the subject, keeping hold of the current EFI and its VisitEFI
            Set colKeepEFIs = New Collection
            colKeepEFIs.Add oEFI
            bReloadVisitEFIResponses = False
            If Not oVisitEFI Is Nothing Then
                ' Hang on to it if it's not ReadOnly
                If Not oVisitEFI.ReadOnly Then
                    colKeepEFIs.Add oVisitEFI
                Else
                    ' Remember we need to reload its responses
                    bReloadVisitEFIResponses = True
                End If
            End If
            Call ReloadMe(colKeepEFIs)
            
            Set colKeepEFIs = Nothing
            
            ' NCJ 26 Sept 02
            ' Must reload the VisitEFI's responses
            If bReloadVisitEFIResponses Then
                ' Get a new reference to the VisitEFI
                Set oVisitEFI = oEFI.VisitInstance.VisitEFormInstance
                ' Assume LoadResponses works OK (we're not bothered about the token - it's locked by someone else anyway!)
                Call oVisitEFI.LoadResponses(msConnection, "", "")
            End If
    
            'remove the save lock
            Call UnlockForSave(msConnection, sToken, Me.StudyDef.StudyId, Me.Site, Me.PersonID)
            SaveResponses = srrSubjectReloaded
            ' and we don't continue with the save
            
        ElseIf Not bCouldLockSubjectForSave Then
            ' We didn't get the lock
            SaveResponses = srrNoLockForSaving
            'this must mean some one else has 'save lock' as study definition cannot be locked at this point
        Else
        
            'Main EForm
            Call SaveEFIResponses(oEFI, nTimezoneOffset)
                
            'Visit EForm
            If Not oVisitEFI Is Nothing Then
                Call SaveEFIResponses(oVisitEFI, nTimezoneOffset)
            End If
            
            ' Update subject label etc.
            Call Me.RefreshProperties
            
            Call oEFI.VisitInstance.UpdateStatus
            Call Me.UpdateStatus
    
            ' NCJ 11 Jun 02 - CBB 2.2.14/17 Call RunEngine whenever data's been saved
            Call moArezzo.GenerateArezzoEvents
            
            ' If things have changed, we need to save the Arezzo state
            ' so ensure the StudySubject's "Changed" flag is set
            moSubjectState.IsChanged = True
            
            ' This saves the eForm & Visit instances etc.
            Call SaveData
            
            'remove the lock
            Call UnlockForSave(msConnection, sToken, Me.StudyDef.StudyId, Me.Site, Me.PersonID)
            sToken = ""

        End If
    
    End If
    
Exit Function

ErrorLabel:
    ' Remember the error no. and desc.
    lErrNo = Err.Number
    sErrDesc = Err.Description & "|StudySubject.SaveResponses"
    On Error Resume Next
    ' NCJ 27 May 03 - See if we need to unlock
    If sToken > "" Then
        'remove the lock if we can
        Call UnlockForSave(msConnection, sToken, Me.StudyDef.StudyId, Me.Site, Me.PersonID)
    End If
    ' Raise the original error
    Err.Raise lErrNo, , sErrDesc
    
End Function

'-------------------------------------------------------------------
Public Property Get ReadOnly() As Boolean
Attribute ReadOnly.VB_Description = "TRUE if subject is read-only and cannot be edited. See also the ReadOnlyReason property"
'-------------------------------------------------------------------
' TRUE if subject is read-only and cannot be edited.
' See also the ReadOnlyReason property
'-------------------------------------------------------------------
' NCJ 20 Mar 02 - Added Frozen and Locked (needed to handle Study Updates properly)
' NCJ 22 Mar 02 - Added mnUpdateMode ( ditto )
'-------------------------------------------------------------------

    ReadOnly = mbArezzoStateErr Or (moStudy.Status = eStudyStatus.ClosedToFollowUp) _
                Or (moStudy.Status = eStudyStatus.Suspended) _
                Or (Me.LockStatus = eLockStatus.lsFrozen) _
                Or (Me.LockStatus = eLockStatus.lsLocked) _
                Or mnUpdateMode = eUIUpdateMode.Read_Only

End Property

'-------------------------------------------------------------------
Public Property Get ReadOnlyReason() As String
Attribute ReadOnlyReason.VB_Description = "The reason why a subject is read-only. Only applicable if ReadOnly property is TRUE."
'-------------------------------------------------------------------
' The reason why a subject is read-only.
' Only applicable if ReadOnly property is TRUE.
'-------------------------------------------------------------------
' NCJ 20 Mar 02 - Include locked/frozen status
' (don't include a message for Read-Only update mode)
'-------------------------------------------------------------------

    If mbArezzoStateErr Then
        ReadOnlyReason = msArezzoErr
    ElseIf mnUpdateMode = eUIUpdateMode.Read_Only Then
            ' No message if user doesn't have Change Data rights or ui rquested readonly
            ReadOnlyReason = ""
    ElseIf Me.LockStatus = eLockStatus.lsFrozen Then
        ReadOnlyReason = "This subject's data has been frozen."
    ElseIf Me.LockStatus = eLockStatus.lsLocked Then
        ReadOnlyReason = "This subject's data has been locked."
    Else
        Select Case moStudy.Status
        Case eStudyStatus.ClosedToFollowUp
            ReadOnlyReason = "This study is closed to followup."
        Case eStudyStatus.Suspended
            ReadOnlyReason = "This study is suspended."
        Case Else
            ' Nothing wrong!
            ReadOnlyReason = ""
        End Select
    End If
    
End Property


'-------------------------------------------------------------------
Public Property Get UserName() As String
Attribute UserName.VB_Description = "The name of the user who opened this subject. This user name will be the default for all response data entered this session"
'-------------------------------------------------------------------
' The name of the user who opened this subject.
' This user name will be the default for all response data
' entered this session
'-------------------------------------------------------------------

    UserName = msUserName
    
End Property

'-------------------------------------------------------------------
Public Property Let UserName(sUserName As String)
'-------------------------------------------------------------------
' The name of the user who opened this subject.
' This user name will be the default for all response data
' entered this session
'-------------------------------------------------------------------

    msUserName = sUserName

End Property

'-------------------------------------------------------------------
Public Property Get UserNameFull() As String
'-------------------------------------------------------------------
' The full name of the user who opened this subject.
' To set this property, use SetUserProperties
'-------------------------------------------------------------------

    ' Return UserName if we don't know the full user name
    If msUserNameFull > "" Then
        UserNameFull = msUserNameFull
    Else
        UserNameFull = msUserName
    End If
    
End Property

'-------------------------------------------------------------------
Public Property Get UserRole() As String
'-------------------------------------------------------------------
' The role of the user who opened this subject.
' To set this property, use SetUserProperties
'-------------------------------------------------------------------

    UserRole = msUserRole
    
End Property

'-------------------------------------------------------------------
Public Property Get ConnectionString() As String
Attribute ConnectionString.VB_Description = "Return connection string used for load"
'-------------------------------------------------------------------
' Return connection string used for load
'-------------------------------------------------------------------

    ConnectionString = msConnection
    
End Property

'-------------------------------------------------------------------
Public Property Get Arezzo() As Arezzo_DM
'-------------------------------------------------------------------
' Get the Arezzo object being used for this study subject
'-------------------------------------------------------------------

    Set Arezzo = moArezzo

End Property

'-------------------------------------------------------------------
Private Sub SetUpCollections(colKeepEFIs As Collection)
'-------------------------------------------------------------------
' Set up our collections of business objects from the State objects we've been given.
' NCJ 18 Sept 02 - Ignore the new EFIs for the ones in the list, but keep the ones given
' (so we can avoid overwriting loaded EFIs that we want to keep)
' colKeepEFIs may be Nothing
' NCJ 27 Jan 03 - Must synchronise the SubjectState collection of EFIStates for the colKeepEFIs
'-------------------------------------------------------------------
Dim oVisitInstance As VisitInstance
Dim oVIState As VisitInstanceState
Dim oEFIState As eFormInstanceState
Dim oKeepEFI As EFormInstance

    On Error GoTo ErrorLabel
    
    ' Set up the Visit Instances collection first
    Set mcolVisitInstances = New Collection
    For Each oVIState In moSubjectState.VisitInstances
        Set oVisitInstance = New VisitInstance
        Call oVisitInstance.SetState(moStudy, Me, oVIState, moArezzo)
        mcolVisitInstances.Add oVisitInstance, oVIState.Key
    Next
    
    ' Set up the eForm Instances collection
    Set mcolEFormInsts = New Collection
    Set mcolIsAnEFI = New Collection
    For Each oEFIState In moSubjectState.eFormInstances
        'TA 25/04/2002
        'basically an eform may have been removed from a visit after a data has been entered
        ' we don't want efis that are no longer in the study to appear
        If moStudy.VisitById(oEFIState.VisitId).IsInVisitById(oEFIState.EFormId) Then
            'only add efi if eform is in visit
            Set oVisitInstance = mcolVisitInstances.Item(VisitInstanceKey(oEFIState.VisitId, oEFIState.VisitCycleNo))
            ' This won't set the state for any EFI in colKeepEFIs
            Call CreateNewEFI(oEFIState, oVisitInstance, colKeepEFIs)
        End If
    Next

    ' NCJ 27 Jan 03 - Now we must replace the EFIStates in the SubjectState for the eFIs in colKeepEFIs
    ' (because the load will have overwritten them)
    If Not colKeepEFIs Is Nothing Then
        For Each oKeepEFI In colKeepEFIs
            ' Remove the one the SubjectState has got
            moSubjectState.eFormInstances.Remove oKeepEFI.EFIState.Key
            ' Now replace it with the one we've kept hold of
            moSubjectState.eFormInstances.Add oKeepEFI.EFIState, oKeepEFI.EFIState.Key
        Next
    End If
    
    ' Tidy up
    Set oVisitInstance = Nothing
    Set oVIState = Nothing
    Set oEFIState = Nothing
    Set oKeepEFI = Nothing
    
Exit Sub

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.SetUpCollections"
    Exit Sub
    
End Sub

'--------------------------------------------------------------------
Friend Function VisitInstanceKey(ByVal lVisitId As Long, ByVal nCycleNo As Integer) As String
'--------------------------------------------------------------------
' temp work around
' Create key comprising VisitId and CycleNo
' this should be replaced when eFI has VI task id
'--------------------------------------------------------------------
Dim oVIState As VisitInstanceState
    For Each oVIState In moSubjectState.VisitInstances
        If oVIState.VisitId = lVisitId And oVIState.CycleNo = nCycleNo Then
             VisitInstanceKey = Str(oVIState.VisitTaskId)
             Exit Function
        End If
    Next
    Err.Raise 100000, "No VisitTaskId for this Visit/Cycle No in StudySubject.SetUpCollections"
    
End Function

'-------------------------------------------------------------------
Private Function eFormCycleKey(oEForm As eFormRO, nCycle As Integer)
'-------------------------------------------------------------------
' Get a key based on eForm definition and a cycle number
'-------------------------------------------------------------------

    eFormCycleKey = oEForm.EFormId & "|" & nCycle
    
End Function

'-------------------------------------------------------------------
Private Function GetHiddenEFormIds() As Collection
'-------------------------------------------------------------------
' Get the eForm Ids that should be hidden in the schedule
' by seeing if there's an instance of the HideIfInactive eForms
'-------------------------------------------------------------------
Dim colToHide As Collection
Dim oEForm As eFormRO

    Set colToHide = New Collection
    
    For Each oEForm In moStudy.eForms
        ' Is it Hide If Inactive?
        If oEForm.HideIfInactive Then
            If Not Me.IsAnInstance(oEForm) Then
                colToHide.Add oEForm.EFormId, oEForm.Key
            End If
        End If
    Next
    
    Set GetHiddenEFormIds = colToHide
    
    ' Tidying up
    Set colToHide = Nothing
    Set oEForm = Nothing
    
End Function

'-------------------------------------------------------------------
Private Sub RefreshSchedule()
'-------------------------------------------------------------------
' Refresh our Schedule Visit collection
' This populates our private mcolScheduleVisits collection
'-------------------------------------------------------------------
Dim oVisit As VisitRO
Dim oVisitInstance As VisitInstance
Dim colVIs As Collection      ' of Visit instances
Dim oSchedVisit As ScheduleVisit
Dim i As Long
Dim colEFormIdsToHide As Collection

    ' Initialise collection of Schedule Visits
    Set mcolScheduleVisits = New Collection
    Set colEFormIdsToHide = GetHiddenEFormIds
    
    ' Walk through the Study Def schedule
    For Each oVisit In moStudy.Visits
        ' Collect the existing Visit Instances
        Set colVIs = Me.VisitInstancesById(oVisit.VisitId)
        If colVIs.Count > 0 Then
            ' There is at least one visit instance
            ' so create a ScheduleVisit object for each
            For i = 1 To colVIs.Count
                ' Collection is indexed by Cycle no
                ' so we do them in order
                Set oVisitInstance = colVIs.Item(Str(i))
                Set oSchedVisit = New ScheduleVisit
                ' NCJ 17 Jan 03 - Added moStudy.eForms parameter
                Call oSchedVisit.Load(oVisit, oVisitInstance, moStudy.eForms, colEFormIdsToHide)
                mcolScheduleVisits.Add oSchedVisit, oSchedVisit.Key
            Next i
        Else
            ' There are no visit instances for this visit
            ' so create just one ScheduleVisit
            Set oSchedVisit = New ScheduleVisit
            ' NCJ 17 Jan 03 - Added moStudy.eForms parameter
            Call oSchedVisit.Load(oVisit, Nothing, moStudy.eForms, colEFormIdsToHide)
            mcolScheduleVisits.Add oSchedVisit, oSchedVisit.Key
        End If
    Next
    
    ' Tidy up
    Set oVisit = Nothing
    Set oVisitInstance = Nothing
    Set colVIs = Nothing
    Set oSchedVisit = Nothing

End Sub

'--------------------------------------------
Public Function LoadResponses(oEFI As EFormInstance, sLockErrMsg As String, _
            ByRef sEFILockToken As String, ByRef sVEFILockToken As String) As eLoadResponsesResult
'--------------------------------------------
' Load the responses for the given eForm instance.
' This also loads the eFormElement definitions for the form.
'' MLM 04/09/02: Load the responses for the corresponding visit eform as well, if there is one.
' now receives and returns the LockToken in sEFILockTOken
'       - this is can be stored and passed back in to ignore EFI locks with this token - for WWW
' NCJ 27 May 03 - Improved error handler (remove subject locks if we can)
'--------------------------------------------
Dim bNeedToReloadSubject As Boolean
Dim sToken As String
Dim lEFormTaskId As Long
Dim bCouldNotLockSubjectForSave As Boolean
Dim oVEFI As EFormInstance
Dim nErrNo As Integer
Dim sErrDesc As String

    On Error GoTo ErrorLabel
    
    sToken = ""
    
    'get lock - the reason for lock failure will be passed back in sLockErrMSg to the calling function
    bCouldNotLockSubjectForSave = Not LockForSave(msConnection, sLockErrMsg, sToken, bNeedToReloadSubject, Me.UserName, _
                                Me.StudyDef.StudyId, Me.Site, Me.PersonID)
        
    If bCouldNotLockSubjectForSave Then
        'study, subbject, subject save lock
        LoadResponses = eLoadResponsesResult.lrrCouldNotLockForSave
'EXIT FUNCTION here
        Exit Function
    End If
        
    'we have a lock we need to check wether it is invalid to decide whether to reload or not
    
    If bNeedToReloadSubject Then
        ' Remember the EFI
        lEFormTaskId = oEFI.eFormTaskId
        ' Reload without keeping copies of anything
        Call ReloadMe(Nothing)
        ' Reset the EFI after the reload
        Set oEFI = eFIByTaskId(lEFormTaskId)
    End If
    
    If oEFI.LoadResponses(msConnection, sLockErrMsg, sEFILockToken) Then
        LoadResponses = eLoadResponsesResult.lrrReadWrite
    Else
        LoadResponses = eLoadResponsesResult.lrrReadOnly
    End If
    
    ' NCJ 11 Feb 03 - Check that it's not the VEFI that we're loading here
    ' (this can happen in Batch Data Entry)
    Set oVEFI = GetVisitEFI(oEFI)       ' NCJ 11 Feb 03
    If Not oVEFI Is Nothing Then
        oVEFI.LoadResponses msConnection, "", sVEFILockToken
        Set oVEFI = Nothing
    End If
    
    'we will only reasch here if we managed to lock
    Call UnlockForSave(msConnection, sToken, Me.StudyDef.StudyId, Me.Site, Me.PersonID)
    sToken = ""
    
Exit Function

ErrorLabel:
    ' Remember the error no. and desc.
    nErrNo = Err.Number
    sErrDesc = Err.Description & "|StudySubject.LoadResponses"
    On Error Resume Next
    ' NCJ 27 May 03 - See if we need to unlock
    If sToken > "" Then
        'remove the lock if we can
        Call UnlockForSave(msConnection, sToken, Me.StudyDef.StudyId, Me.Site, Me.PersonID)
    End If
    ' Raise the original error
    Err.Raise nErrNo, , sErrDesc
    
End Function

'---------------------------------------------------------------------
Private Function GetVisitEFI(oEFI As EFormInstance) As EFormInstance
'---------------------------------------------------------------------
' NCJ 11 Feb 03
' Get the Visit eForminstance if there is one
' AND if it's different from the given EFI (necessary for BDE to know)
' Returns NOTHING if either fails
'---------------------------------------------------------------------
Dim oVEFI As EFormInstance

    Set oVEFI = oEFI.VisitInstance.VisitEFormInstance
    If Not oVEFI Is Nothing Then
        ' Only return it if it's different
        If oVEFI.eFormTaskId = oEFI.eFormTaskId Then
            Set oVEFI = Nothing
        End If
    End If

    Set GetVisitEFI = oVEFI
    
    Set oVEFI = Nothing

End Function

'--------------------------------------------
Public Sub RemoveResponses(oEFI As EFormInstance, ByVal bRemoveLock As Boolean)
Attribute RemoveResponses.VB_Description = "Remove from memory the responses for the eForm instance. Also removes the eFormElements from the corresponding eForm. This should be called when an EFI is finished with."
'--------------------------------------------
' Remove from memory the responses for the eForm instance.
' Also removes the eFormElements from the corresponding eForm.
' This should be called when an EFI is finished with.
'
' MLM 04/09/02: Remove the visit's visit date eform's instance's responses too.
' NCJ 20 Sept 02 - Only remove eForm lock if bRemoveLock = true
' NCJ 11 Feb 03 - Check to see if it's the VEFI we're dealing with (fix for BDE)
'--------------------------------------------
Dim oVEFI As EFormInstance

    On Error GoTo ErrorLabel
    
    Call oEFI.RemoveResponses(bRemoveLock)
    
    Set oVEFI = GetVisitEFI(oEFI)       ' NCJ 11 Feb 03
    If Not oVEFI Is Nothing Then
        Call oVEFI.RemoveResponses(bRemoveLock)
        Set oVEFI = Nothing
    End If
    
Exit Sub

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.RemoveResponses"
    Exit Sub
    
End Sub

'--------------------------------------------
Public Sub CancelResponses(oEFI As EFormInstance, Optional ByVal bRemoveLock As Boolean = True)
Attribute CancelResponses.VB_Description = "Cancel an eForm - discard responses and reset the Arezzo state back to what it was at the last Save. This automatically calls RemoveResponses."
'--------------------------------------------
' Cancel an eForm - discard responses and reset the Arezzo state
' back to what it was at the last Save.
' This automatically calls RemoveResponses (AND removes the eForm lock).
'--------------------------------------------
' NCJ 10 Oct 01 - Cancel any Lab code selections too
' NCJ 28 Jun 06 - Only remove lock if bRemoveLock = TRUE (pass through to RemoveResponses) for Web DE
'--------------------------------------------
Dim bNeedToReloadSubject As Boolean
Dim sToken As String
Dim bCouldNotLockSubjectForSave As Boolean
Dim sLockErrMsg As String
Dim bReloaded As Boolean
Dim lEFormTaskId As Long
Dim oVEFI As EFormInstance

    On Error GoTo ErrorLabel
    
    bReloaded = False
    
    'get lock - the reason for lock failure will be passed back in sLockErrMSg to the calling function
    bCouldNotLockSubjectForSave = Not LockForSave(msConnection, sLockErrMsg, sToken, bNeedToReloadSubject, Me.UserName, _
                                Me.StudyDef.StudyId, Me.Site, Me.PersonID)
        
    If bCouldNotLockSubjectForSave Then
        ' study, subject, subject save lock
        ' We just forget it and act as if we didn't need a reload
    Else
        
        'we have a lock, so we need to check whether subject is invalid to decide whether to reload or not
        If bNeedToReloadSubject Then
            ' NCJ 28 Jun 06 - Use bRemoveLock
            Call RemoveResponses(oEFI, bRemoveLock) 'also removes visitEFI responses.
            ' Remember the EFI
            lEFormTaskId = oEFI.eFormTaskId
            ' Reload without keeping copies of anything
            Call ReloadMe(Nothing)
            ' Reset the EFI after the reload
            Set oEFI = eFIByTaskId(lEFormTaskId)
            bReloaded = True
        End If
    
        'we will only reach here if we managed to lock
        Call UnlockForSave(msConnection, sToken, Me.StudyDef.StudyId, Me.Site, Me.PersonID)
    
    End If
    
    ' If we didn't reload, just carry on as normal
    If Not bReloaded Then
        ' Reload the last stored patient state back into Arezzo
        moArezzo.PatientState = moSubjectState.ArezzoState
        Call oEFI.Cancel        ' NCJ 10/10/01
        
        'MLM 12/09/02: Also cancel the visit eform instance, if there is one
        Set oVEFI = GetVisitEFI(oEFI)       ' NCJ 11 Feb 03
        If Not oVEFI Is Nothing Then
            Call oVEFI.Cancel
            Set oVEFI = Nothing
        End If
        
        ' NCJ 28 Jun 06 - Use bRemoveLock
        Call RemoveResponses(oEFI, bRemoveLock) 'also removes visitEFI responses.
    End If
    
Exit Sub

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.CancelResponses"
    Exit Sub
    
End Sub

'--------------------------------------------
Private Function SaveEFIResponses(ByRef oEFI As EFormInstance, Optional nTimezoneOffset As Integer = 0) As Boolean
'--------------------------------------------
' MLM 13/09/02: Added. Saves the responses, etc for one EFormInstance, if required.
'               Uses code formerly in SaveResponses.
'               Returns TRUE if the form needed to be saved.
' RS 22/9/2002: Added optional nTimezoneOffset parameter, this parameter will be used by the
'               web client to pass the local nTimezoneOffset
'--------------------------------------------
    
    On Error GoTo ErrorLabel
    
    ' If responses changed or it's a new empty form
    If oEFI.Responses.Changed Or _
     (oEFI.Responses.Count = 0 And oEFI.Status = eStatus.Requested) Then
        
        ' NCJ 28/9/01 - Ensure Requested Responses are changed to Missing
        Call oEFI.Responses.DealWithRequestedResponses(nTimezoneOffset)
        Call oEFI.RefreshLabels
        Call oEFI.UpdateStatus
        
        ' This saves JUST the responses
        ' (Does nothing if no responses to save)
        Call oEFI.SaveResponses(msConnection)
        
        ' Confirm form if it's complete
        ' NCJ 1 May 03 - Bug 1280 - Do not confirm eForm if it's Unobtainable
        If oEFI.Complete And oEFI.Status <> eStatus.Unobtainable Then
            Call moArezzo.ConfirmRequestedForm(oEFI.eFormTaskId)
        End If
        
        SaveEFIResponses = True
    Else
        SaveEFIResponses = False
    End If

Exit Function

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.SaveResponses"

End Function


'--------------------------------------------
Private Function DoEFIResponsesNeedSaving(ByRef oEFI As EFormInstance) As Boolean
'--------------------------------------------
' TA Sept 02
' Code copied from SaveEFIResponses
' Returns TRUE if the form needs to be saved.
' Always return FALSE for a read-only eForm
' NCJ 21 May 03 - Return TRUE if eForm label has changed (even if responses haven't)
'--------------------------------------------
    
    On Error GoTo ErrorLabel
    
    DoEFIResponsesNeedSaving = False
    
    If oEFI.ReadOnly Then Exit Function
    
    ' NCj 21 May 03 - Refresh label
    Call oEFI.RefreshLabels
    
    ' If responses changed or it's a new empty form
    ' or its label has changed
    If oEFI.Responses.Changed Or _
     (oEFI.Responses.Count = 0 And oEFI.Status = eStatus.Requested) _
     Or oEFI.Changed Then
        DoEFIResponsesNeedSaving = True
    Else
        DoEFIResponsesNeedSaving = False
    End If

Exit Function

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.DoEFIResponsesNeedSaving"

End Function

'--------------------------------------------
Public Function SaveRegistration(Optional bStatusOnly As Boolean = False)
Attribute SaveRegistration.VB_Description = "This saves JUST the registration details for the subject. It does not save visit or eForm instances etc. If bStatusOnly is TRUE, then only the RegStatus will be saved"
'--------------------------------------------
' This saves JUST the registration details for the subject.
' It does not save visit or eForm instances etc.
' If bStatusOnly is TRUE, then only the RegStatus will be saved
'--------------------------------------------
Dim oStudySubjDS As StudySubjectDS

    On Error GoTo ErrorLabel
    
    If Not moSubjectState.IsChanged Then Exit Function
    
    ' NCJ 26 Mar 03 - Just use the general subject Save routine
    ' (the sledgehammer approach)
    ' because original code doesn't contain necessary locking
    Call Me.Save
    
'    Set oStudySubjDS = New StudySubjectDS
'
'    If bStatusOnly Then
'        ' Only want to save RegistrationStatus
'        ' (NCJ 7 Mar 03 - We don't set the SaveTimeStamp here because no subject data has changed)
'        Call oStudySubjDS.SaveRegistrationStatus(msConnection, _
'                            moSubjectState.RegStatus, _
'                            moSubjectState.StudyId, _
'                            moSubjectState.Site, _
'                            moSubjectState.PersonID)
'    Else
'        ' Retrieve the latest Arezzo state
'        ' (contains the Registration ID)
'        moSubjectState.ArezzoState = moArezzo.PatientState
'        ' NCJ 7 Mar 03 - Set the save time stamp
'        Call SetSaveTimeStamp
'        Call oStudySubjDS.SaveRegistration(msConnection, moSubjectState.State)
'    End If
'
'    moSubjectState.IsNew = False
'    moSubjectState.IsChanged = False
'
'    Set oStudySubjDS = Nothing
    
Exit Function

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.SaveRegistration"
    Exit Function

End Function

'--------------------------------------------
Public Function Save() As Long
'--------------------------------------------
' "Standalone" save for just the subject details, including Label, DOB, Gender and AREZZO State
' (Does not save any response data)
' Gets appropriate save locks etc.
' Use for stuff outside eForms, e.g. changes to AREZZO tasks, Registration etc.
' Returns an eSaveResponsesResult
' NCJ 11 Jun 03 - Must preserve all loaded eForm responses through a reload
'--------------------------------------------
Dim sToken As String
Dim sLockErrMsg As String
Dim bNeedToReloadSubject As Boolean
Dim bCouldLockSubjectForSave As Boolean
Dim colKeepEFIs As Collection
Dim colReloadEFIs As Collection
Dim oEFI As EFormInstance
Dim vEFormTaskId As Variant

    ' Extra check to ensure we don't save a read-only subject
    If Me.ReadOnly Then Exit Function
    
    ' Get lock for save
     bCouldLockSubjectForSave = LockForSave(msConnection, sLockErrMsg, sToken, bNeedToReloadSubject, Me.UserName, _
                            Me.StudyDef.StudyId, Me.Site, Me.PersonID)
    
    ' Check for a reload
    If bCouldLockSubjectForSave And bNeedToReloadSubject Then
    
        ' Reload the subject (keeping any eForms which currently have responses loaded)
        Set colKeepEFIs = New Collection
        Set colReloadEFIs = New Collection
        For Each oEFI In Me.eFormInsts
            If oEFI.ResponsesLoaded Then
                If oEFI.ReadOnly Then
                    ' We have to reload this one's responses after the subject relaod
                    colReloadEFIs.Add oEFI.eFormTaskId
                Else
                    ' Hang on to this one
                    colKeepEFIs.Add oEFI
                End If
            End If
        Next
        Call ReloadMe(colKeepEFIs)
        
        ' Were there any read-only eForms to reload?
        For Each vEFormTaskId In colReloadEFIs
            Set oEFI = Me.eFIByTaskId(CLng(vEFormTaskId))
            Call oEFI.LoadResponses(msConnection, "", "")
        Next
        
        Set colKeepEFIs = Nothing
        Set colReloadEFIs = Nothing
        Set oEFI = Nothing
    End If
    
    If Not bCouldLockSubjectForSave Then
        ' We didn't get the lock
        Save = eSaveResponsesResult.srrNoLockForSaving
    Else
    
        ' NCJ 28 Mar 03 - Must make sure we get all the latest eForm and Visit instances
        Call moArezzo.GenerateArezzoEvents
    
        ' We want to save the Arezzo state
        ' so ensure the StudySubject's "Changed" flag is set
        moSubjectState.IsChanged = True
        
        ' This saves the whole caboodle
        Call SaveData
        
        'remove the lock
        Call UnlockForSave(msConnection, sToken, Me.StudyDef.StudyId, Me.Site, Me.PersonID)

        Save = eSaveResponsesResult.srrSuccess
    
    End If

Exit Function
ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.Save"

End Function

'--------------------------------------------
Private Function SaveData() As Long
Attribute SaveData.VB_Description = "Save changes to subject and to Visit and eForm instances. Returns SubjectID. (See SaveResponses for saving Response data)."
'--------------------------------------------
' Save changes to subject and to Visit and eForm instances.
' Returns SubjectID.
' (See SaveResponses for saving Response data).
' Assumes lock for saving already in place
'--------------------------------------------
Dim oStudySubjDS As StudySubjectDS

    On Error GoTo ErrorLabel
    
    ' NCJ 20 Mar 02 - Extra check to ensure we don't save a read-only subject
    If Me.ReadOnly Then Exit Function
    
    ' Retrieve the latest Arezzo state
    moSubjectState.ArezzoState = moArezzo.PatientState
    
    Set oStudySubjDS = New StudySubjectDS
    
    'invalidate the cache (but not our own entry
    Call MACROLOCKBS30.CacheInvalidate(msConnection, Me.StudyDef.StudyId, Me.Site, Me.PersonID, msCacheToken)
    
    ' NCJ 7 Mar 03 - Set the save time stamp
    Call SetSaveTimeStamp
    
    SaveData = oStudySubjDS.Save(msConnection, moSubjectState.State)
   
    Call ChangeStatetoSaved
    Set oStudySubjDS = Nothing

Exit Function

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.SaveData"

End Function

'--------------------------------------------
Private Sub SetSaveTimeStamp()
'--------------------------------------------
' NCJ 7 Mar 03 - Set the save time stamp in the Subject State
'--------------------------------------------

    moSubjectState.SubjectTimeStamp = IMedNow
    moSubjectState.SubjectTimeStamp_TZ = moTimezone.TimezoneOffset

End Sub

'--------------------------------------------
Public Property Get ScheduleVisits() As Collection
Attribute ScheduleVisits.VB_Description = "The collection of Schedule Visit objects. Created on demand and automatically refreshed as required."
'--------------------------------------------
' The collection of Schedule Visit objects.
' Created on demand and automatically refreshed as required.
'--------------------------------------------

    ' Is it OK from last time?
    If mcolScheduleVisits Is Nothing Then
        Call RefreshSchedule
    End If
    Set ScheduleVisits = mcolScheduleVisits

End Property

'--------------------------------------------
Public Property Get ScheduleGrid() As ScheduleGrid
Attribute ScheduleGrid.VB_Description = "The Schedule Grid structure. Created on demand and automatically refreshed as required."
'--------------------------------------------
' The Schedule Grid structure.
' Created on demand and automatically refreshed as required.
'--------------------------------------------
    
    ' Is it OK from last time?
    If moScheduleGrid Is Nothing Then
        Set moScheduleGrid = New ScheduleGrid
        Call moScheduleGrid.Build(Me)
    End If
    Set ScheduleGrid = moScheduleGrid
    
End Property

'--------------------------------------------
Private Sub Class_Initialize()
'--------------------------------------------
' Initialise ourselves
'--------------------------------------------

    Set moSubjectState = New StudySubjectState
    Set moArezzo = Nothing
    Set mcolEFormInsts = New Collection
    Set mcolVisitInstances = New Collection
    Set mcolIsAnEFI = New Collection
    Set mcolScheduleVisits = Nothing
    Set moScheduleGrid = Nothing

    ' RS 22/9/2002
    Set moTimezone = New Timezone

End Sub

'--------------------------------------------
Friend Sub Terminate(Optional colKeepEFIs As Collection = Nothing)
'--------------------------------------------
' Tidy up our objects before closing down
' NCJ 14 Feb 03 - Ensure all EFI responses are cleared up
' NCJ 5 Mar 03 - Added colKeepEFIs so that we do not remove responses
'       for EFIs we're trying to hang on to during a subject reload
' NCJ 24 Mar 04 - Ensure we clear out object collections
' NCJ 27 May 04 - Bug 2282 - Don't terminate things in colKeepEFIs
'--------------------------------------------
Dim oEFI As EFormInstance
Dim oVI As VisitInstance

    Call RemoveAllResponses(colKeepEFIs)
    
    'remove our token form the AREZZO token table
    Call MACROLOCKBS30.CacheRemoveSubjectRow(msConnection, msCacheToken)
    'clear the token
    msCacheToken = ""
    
    Set moSubjectState = Nothing
    Set moArezzo = Nothing
    Set mcolIsAnEFI = Nothing
    ' NCJ 26 Mar 04 - Do NOT set Study to Nothing here (only do it in the class terminate)
    ' Set moStudy = Nothing
    
    ' NCJ 24 Mar 04 - These are object collections
    If Not mcolEFormInsts Is Nothing Then
        For Each oEFI In mcolEFormInsts
            ' NCJ 27 May 04 - Check it's not one we want to keep
            If EFIToBeKept(oEFI.eFormTaskId, colKeepEFIs) Is Nothing Then
                Call oEFI.Terminate
            End If
        Next
        Call CollectionRemoveAll(mcolEFormInsts)
    End If
    Set mcolEFormInsts = Nothing
    Set oEFI = Nothing
    
    If Not mcolVisitInstances Is Nothing Then
        For Each oVI In mcolVisitInstances
            Call oVI.Terminate
        Next
        Call CollectionRemoveAll(mcolVisitInstances)
    End If
    Set mcolVisitInstances = Nothing
    Set oVI = Nothing
    
    Call EmptyTheSchedule       ' New routine - NCJ 24 Mar 04

    ' NCJ 26 Sept 02
    Set moTimezone = Nothing

End Sub

'--------------------------------------------
Private Sub Class_Terminate()
'--------------------------------------------
' Tidy up when we close
'--------------------------------------------
    
    Call Terminate
    ' It's only safe to blitz the Study when the Subject is REALLY going away
    Set moStudy = Nothing

End Sub

'--------------------------------------------
Private Sub RemoveAllResponses(colKeepEFIs As Collection)
'--------------------------------------------
' Remove all responses of loaded EFIs
' (to be done during class termination)
' If colKeepEFIs is not nothing, do NOT throw away responses of EFIs in the collection
'--------------------------------------------
Dim oEFI As EFormInstance

    ' Forget about errors because if we're exiting out of a crash
    ' we don't want more errors to deal with!
    On Error Resume Next
    
    If mcolEFormInsts Is Nothing Then Exit Sub
    
    For Each oEFI In mcolEFormInsts
        If oEFI.ResponsesLoaded Then
            ' Check it's not one we want to keep
            If EFIToBeKept(oEFI.eFormTaskId, colKeepEFIs) Is Nothing Then
                Call RemoveResponses(oEFI, True)
            End If
        End If
    Next

End Sub

'--------------------------------------------
Public Property Get StudyDef() As StudyDefRO
'--------------------------------------------
' The study definition object for the study
'--------------------------------------------

    Set StudyDef = moStudy
    
End Property

'--------------------------------------------
Public Property Get StudyId() As Long
'--------------------------------------------
' The numeric ID of the Study definition
'--------------------------------------------

    StudyId = moStudy.StudyId
    
End Property

'--------------------------------------------
Public Property Get Version() As Integer
'--------------------------------------------
' Version of the study
'--------------------------------------------

    Version = moStudy.Version
    
End Property

'--------------------------------------------
Public Property Get StudyCode() As String
'--------------------------------------------
' Unique text code of the study definition
'--------------------------------------------

    StudyCode = moStudy.Name
    
End Property

'--------------------------------------------
Public Property Get Site() As String
Attribute Site.VB_Description = "Site of subject"
'--------------------------------------------
' Site of subject
'--------------------------------------------

    Site = moSubjectState.Site
    
End Property

'--------------------------------------------
Public Property Get PersonID() As Long
Attribute PersonID.VB_Description = "The subject ID for this particular study and site"
'--------------------------------------------
' The subject ID for this particular study and site
'--------------------------------------------

    PersonID = moSubjectState.PersonID
    
End Property

'--------------------------------------------
Public Property Get Status() As Integer
Attribute Status.VB_Description = "The Status of the subject, calculated from the status of all the visit and eform instances"
'--------------------------------------------
' The Status of the subject, calculated from the status
' of all the visit and eform instances
'--------------------------------------------

    Status = moSubjectState.Status
    
End Property

'--------------------------------------------
Public Property Get LockStatus() As Integer
Attribute LockStatus.VB_Description = "The Lock Status of the subject"
'--------------------------------------------
' The Lock Status of the subject
'--------------------------------------------

    LockStatus = moSubjectState.LockStatus
    
End Property

'--------------------------------------------
Public Property Get DOB() As Double
Attribute DOB.VB_Description = "The calculated Date of Birth for the subject. Will be 0 if undefined"
'--------------------------------------------
' The calculated Date of Birth for the subject.
' Will be 0 if undefined
'--------------------------------------------

    DOB = moSubjectState.DOB
    
End Property

'--------------------------------------------
Public Property Get Label() As String
Attribute Label.VB_Description = "The calculated label for the subject"
'--------------------------------------------
' The calculated label for the subject
'--------------------------------------------

    Label = moSubjectState.Label
    
End Property

'--------------------------------------------
Public Property Get RegistrationID() As String
Attribute RegistrationID.VB_Description = "The calculated registration identifier for the subject. Setting this property automatically sets the Arezzo data too. Need to call SaveRegistration after setting RegistrationID"
'--------------------------------------------
' The calculated registration identifier for the subject.
' Setting this property automatically sets the Arezzo data too.
' Need to call SaveRegistration after setting RegistrationID
'--------------------------------------------

    RegistrationID = moSubjectState.RegID
    
End Property

'--------------------------------------------
Public Property Let RegistrationID(ByVal sRegId As String)
'--------------------------------------------
' The calculated registration identifier for the subject.
' Setting this property automatically sets the Arezzo data too.
' Need to call SaveRegistration after setting RegistrationID
'--------------------------------------------
' NCJ 26 Sep 01 - Update Subject Label here too
'--------------------------------------------
Dim sArezzoID As String

    If moSubjectState.RegID <> sRegId Then
        moSubjectState.RegID = sRegId
        ' Add subject identifier to Arezzo and save it
        ' NCJ 4/12/00 - Ensure numbers are added as a string to stop Arezzo stripping off leading zeros
        ' NCJ 2 Apr 03 - Add single quotes anyway to stop AREZZO parsing as a Prolog term (e.g. Prefix-Number)
        sArezzoID = "'" & ReplaceQuotes(sRegId) & "'"
'        If IsNumeric(sRegId) Then
'            sArezzoID = "'" & sRegId & "'"
'        Else
'            ' Don't do it for other ones or we might get into trouble with single quotes
'            sArezzoID = sRegId
'        End If
        Call moArezzo.AddData("person:subjectidentifier", sArezzoID)
        Call RefreshProperties      ' NCJ 26 Sep 01
        moSubjectState.IsChanged = True
    End If

End Property

'--------------------------------------------
Public Property Get RegistrationStatus() As Integer
Attribute RegistrationStatus.VB_Description = "The registration status of the subject. After setting this property need to call SaveRegistration."
'--------------------------------------------
' The registration status of the subject.
' After setting this property need to call SaveRegistration.
'--------------------------------------------

    RegistrationStatus = moSubjectState.RegStatus
    
End Property

'--------------------------------------------
Public Property Let RegistrationStatus(ByVal nStatus As Integer)
'--------------------------------------------
' The registration status of the subject.
' After setting this property need to call SaveRegistration.
'--------------------------------------------

    If moSubjectState.RegStatus <> nStatus Then
        moSubjectState.RegStatus = nStatus
        moSubjectState.IsChanged = True
    End If

End Property

'--------------------------------------------
Public Property Get Gender() As Integer
Attribute Gender.VB_Description = "The calculated gender of the subject"
'--------------------------------------------
' The calculated gender of the subject
'--------------------------------------------

    Gender = moSubjectState.Gender
    
End Property

'-------------------------------------
Public Property Get Changed() As Boolean
Attribute Changed.VB_Description = "Whether Subject has changed (or is new)"
'-------------------------------------
' Whether Subject has changed (or is new)
'-------------------------------------

    Changed = (moSubjectState.IsChanged Or moSubjectState.IsNew)

End Property

'--------------------------------------------
Public Sub RefreshProperties()
Attribute RefreshProperties.VB_Description = "Recalculate the subject label, DOB and Gender."
'--------------------------------------------
' Recalculate the subject label, DOB and Gender.
'--------------------------------------------
Dim sNewLabel As String
Dim dblDOB As Double
Dim nNewGender As Integer

    On Error GoTo ErrorLabel
    
    ' Reset context for AREZZO (no eForm or Visit at this point)
    Call moArezzo.SetMACROTaskContext(Nothing, Nothing)
    
    ' Call Arezzo to evaluate Label expression
    sNewLabel = moArezzo.EvaluateExpression(moStudy.LabelExpr)
    If Not moArezzo.ResultOK(sNewLabel) Then
        sNewLabel = ""
    End If
    ' Store it if it has changed
    If sNewLabel <> moSubjectState.Label Then
        'inform listeners that subjectlabel has changed
        RaiseEvent SubjectLabelChanged(moSubjectState.Label, sNewLabel)
        moSubjectState.Label = sNewLabel
        moSubjectState.IsChanged = True
'        Call moArezzo.AddData("person:", sNewLabel)
    End If
    
    ' Call Arezzo to evaluate Date of Birth expression
    dblDOB = moArezzo.EvaluateDate(moStudy.DOBExpr)
    If dblDOB <> moSubjectState.DOB Then
        moSubjectState.DOB = dblDOB
        moSubjectState.IsChanged = True
    End If
    
    ' Call Arezzo to evaluate Gender expression
    nNewGender = GenderCode.gNone
    If moStudy.GenderExpr > "" Then
        sNewLabel = moArezzo.EvaluateExpression(moStudy.GenderExpr)
        If IsNumeric(sNewLabel) Then
            nNewGender = Val(sNewLabel)
            ' Check it's a valid Gender value
            Select Case nNewGender
            Case GenderCode.gFemale, GenderCode.gMale
                ' These are OK
            Case Else
                nNewGender = GenderCode.gNone
            End Select
        End If
    End If
    If nNewGender <> moSubjectState.Gender Then
        moSubjectState.Gender = nNewGender
        moSubjectState.IsChanged = True
    End If
    
Exit Sub

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.RefreshProperties"
    Exit Sub
    
End Sub

'----------------------------------------------------------------------
Public Function eFIByTaskId(lEFITaskId As Long) As EFormInstance
Attribute eFIByTaskId.VB_Description = "Return the EFI corresponding to this eFormTaskId"
'----------------------------------------------------------------------
' Return the EFI corresponding to this eFormTaskId
'----------------------------------------------------------------------

    Set eFIByTaskId = mcolEFormInsts.Item(Str(lEFITaskId))

End Function

'----------------------------------------------------------------------
Public Function VisitInstanceByTaskId(lVisitTaskId As Long) As VisitInstance
Attribute VisitInstanceByTaskId.VB_Description = "Return the visit instance correponding to this VisitTaskId"
'----------------------------------------------------------------------
' Return the visit instance correponding to this VisitTaskId
'----------------------------------------------------------------------

    Set VisitInstanceByTaskId = mcolVisitInstances.Item(Str(lVisitTaskId))

End Function

'------------------------------------------------
Public Property Get VisitInstances() As Collection
Attribute VisitInstances.VB_Description = "The visit instances collection (as VisitInstance objects)"
'------------------------------------------------
' The visit instances collection (as VisitInstance objects)
'------------------------------------------------

    Set VisitInstances = mcolVisitInstances

End Property

'-----------------------------------------------------------------
Public Function VisitInstancesById(lVisitId As Long) As Collection
Attribute VisitInstancesById.VB_Description = "Return all visit instances which have this VisitId (may be none). Collection is indexed by Cycle number"
'-----------------------------------------------------------------
' Return all visit instances which have this VisitId (may be none).
' Collection is indexed by Cycle number
'-----------------------------------------------------------------
Dim ocolVIs As Collection
Dim oVI As VisitInstance

    On Error GoTo ErrorLabel
    
    Set ocolVIs = New Collection
    
    For Each oVI In mcolVisitInstances
        If oVI.Visit.VisitId = lVisitId Then
            ' Use Cycle no. as key
            ocolVIs.Add oVI, Str(oVI.CycleNo)
        End If
    Next
    
    Set VisitInstancesById = ocolVIs
    
    Set oVI = Nothing
    Set ocolVIs = Nothing
    
Exit Function

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.VisitInstancesById"
    Exit Function
    
End Function

'------------------------------------------------------------------------------
Public Function IsAnInstance(oEForm As eFormRO, Optional nCycle As Integer = 1) As Boolean
Attribute IsAnInstance.VB_Description = "Returns TRUE if there is an instance of the given eForm/Cycle no."
'------------------------------------------------------------------------------
' Returns TRUE if there is an instance of the given eForm/Cycle no.
' MLM 29/08/02: mcolIsAnEFI has been changed to not contain instances of Visit eForms,
'   so IsAnInstance returns FALSE if there are no user instances.
'------------------------------------------------------------------------------

    ' Test for membership
    IsAnInstance = CollectionMember(mcolIsAnEFI, eFormCycleKey(oEForm, nCycle), False)
    
End Function

'------------------------------------------------
Public Property Get eFormInsts() As Collection
Attribute eFormInsts.VB_Description = "The collection of eFormInstances for this subject (indexed by eFormTaskID)"
'------------------------------------------------
' The collection of eFormInstances for this subject
' (indexed by eFormTaskID)
'------------------------------------------------

    Set eFormInsts = mcolEFormInsts

End Property

'------------------------------------------------
Public Property Get eForms() As Collection
Attribute eForms.VB_Description = "Get collection of eForm definitions in the study (as eFormRO objects), indexed by eFormId."
'------------------------------------------------
' Get collection of eForm definitions in the study (as eFormRO objects),
' indexed by eFormId.
'------------------------------------------------

    Set eForms = moStudy.eForms

End Property

'------------------------------------------------
Private Sub ChangeStatetoSaved()
'------------------------------------------------
' set all the state objects IsNew and IsChanged properties to false
' (this is required after a save)
'------------------------------------------------
Dim i As Long

    With moSubjectState
        .IsChanged = False
        .IsNew = False
        For i = 1 To .VisitInstances.Count
            .VisitInstances(i).IsChanged = False
            .VisitInstances(i).IsNew = False
        Next
        For i = 1 To .eFormInstances.Count
            .eFormInstances(i).IsChanged = False
            .eFormInstances(i).IsNew = False
        Next
    End With
    
    ' NCJ 31 Jan 03 - Also clear Arezzo activities
    Call moArezzo.SetAsSaved
    
End Sub

'------------------------------------------------
Private Sub AddNewEFI(ByVal lEFormId As Long, ByVal lEFormTaskId As Long, _
                    ByVal nCycleNo As Integer, ByVal lVisitTaskId As Long)
'------------------------------------------------
' Add a new EForm Instance object
' and add it in to all relevant collections
' NB We assume that its owning VisitInstance object already exists,
' which it will if Arezzo passes through events in the right order,
' which it normally does :-)
' NCJ 20 Mar 02 - Handle locked/frozen visits, assuming Subject is not Read-Only
' NCJ 28 Apr 03 - Do add new EFIs to frozen visits
'------------------------------------------------
Dim oEFIState As eFormInstanceState
Dim oEForm As eFormRO
Dim oVI As VisitInstance

    On Error GoTo ErrorLabel
        
    Set oVI = VisitInstanceByTaskId(lVisitTaskId)   ' Assume this exists!
    
    ' NCJ 20 Mar 02 - Do not add EFIs to frozen visits
    ' (but we do add them to Locked visits)
    ' NCJ 28 Apr 03 - We DO add EFIs to frozen visits (because in MACRO 3.0 they can unfreeze)
'    If oVI.LockStatus = eLockStatus.lsFrozen Then Exit Sub
    
    ' Pick up the associated objects
    Set oEForm = moStudy.eFormById(lEFormId)
    
    ' Create and initialise a new eForm Instance State object
    Set oEFIState = New eFormInstanceState
    Call InitialiseEFIState(oEFIState, oEForm, lEFormTaskId, nCycleNo, oVI)
    
    ' Create the eFormInstance itself
    Call CreateNewEFI(oEFIState, oVI)
    ' Add it to subject state
    moSubjectState.eFormInstances.Add oEFIState, oEFIState.Key
    
    ' Reset the Schedule to force a refresh
'    Set mcolScheduleVisits = Nothing
'    Set moScheduleGrid = Nothing
    Call EmptyTheSchedule       ' New routine - NCJ 24 Mar 04
    
    ' Tidy up
    Set oEFIState = Nothing
    Set oEForm = Nothing
    Set oVI = Nothing
    
Exit Sub

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.AddNewEFI"
    Exit Sub
    
End Sub

'------------------------------------------------
Private Function CreateNewEFI(oEFIState As eFormInstanceState, _
                                oVI As VisitInstance, _
                                Optional colKeepEFIs As Collection = Nothing) As EFormInstance
'------------------------------------------------
' Create a new eFormInstance belonging to given Visit instance
' and add to relevant collections
' oEFIState is initialised state object
'
' MLM 29/08/02: Only record the EFI in mcolIsAnEFI if the EFI is a User EFI,
'   i.e. if it should be displayed on the schedule.
'   Also, if creating a Visit eForm instance, tell its visit instance about it
' NCJ 18 Sept 02 - Do not create new EFIs for those in colKeepEFIs but use the instances given instead
'------------------------------------------------

Dim oEFI As EFormInstance

    On Error GoTo ErrorLabel
    
    ' Is this one we want to keep the old version of?
    Set oEFI = EFIToBeKept(oEFIState.eFormTaskId, colKeepEFIs)
    If oEFI Is Nothing Then
        ' Create a new one
        Set oEFI = New EFormInstance
        Call oEFI.SetState(moStudy, oVI, oEFIState)
    End If
    
    ' Add to our collection
    mcolEFormInsts.Add oEFI, oEFIState.Key
    
    ' Give a copy to the Visit Instance
    oVI.AddEFormInstance oEFI
    
    'MLM 29/08/02:
    Select Case oVI.Visit.VisitEFormByEForm(oEFI.eForm).EFormUse
    Case eEFormUse.User
        ' Record an eFormInstance by eFormId and Cycle
        Call CollectionAddAnyway(mcolIsAnEFI, True, eFormCycleKey(oEFI.eForm, oEFI.CycleNo))
    Case eEFormUse.VisitEForm
        Set oVI.VisitEFormInstance = oEFI
    End Select
    
    Set CreateNewEFI = oEFI
    
    Set oEFI = Nothing
    
Exit Function

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.CreateNewEFI"
    Exit Function
    
End Function

'------------------------------------------------
Private Function EFIToBeKept(lEFormTaskId As Long, colKeepEFIs As Collection) As EFormInstance
'------------------------------------------------
' Returns EFI to be kept if the lEFormTaskId matches an EFI in the colKeepEFIs collection,
' or Nothing otherwise
' colKeepEFIs may be Nothing
'------------------------------------------------
Dim oKeepEFI As EFormInstance

    Set EFIToBeKept = Nothing
    
    If Not colKeepEFIs Is Nothing Then
        ' See if any EFIs in the collection match this eFormTaskID
        For Each oKeepEFI In colKeepEFIs
            If oKeepEFI.eFormTaskId = lEFormTaskId Then
                ' It matches
                Set EFIToBeKept = oKeepEFI
                Exit For
            End If
        Next
    End If

    Set oKeepEFI = Nothing

End Function

'------------------------------------------------
Private Sub AddNewVI(ByVal lVisitId As Long, ByVal lVisitTaskId As Long, _
                    ByVal nCycleNo As Integer)
'------------------------------------------------
' Add a new VisitInstance object
' and add it in to all relevant collections
'------------------------------------------------
Dim oVIState As VisitInstanceState
Dim oVisit As VisitRO
Dim oVI As VisitInstance

    On Error GoTo ErrorLabel
        
    ' Pick up the Visit definition
    Set oVisit = moStudy.VisitById(lVisitId)
    ' Create and initialise a new VisitInstance State object
    Set oVIState = New VisitInstanceState
    Call InitialiseVIState(oVIState, oVisit, lVisitTaskId, nCycleNo)
    
    ' Add State object to StudySubject state
    moSubjectState.VisitInstances.Add oVIState, oVIState.Key
    
    ' Create a new Visit Instance and add it to our collection
    Set oVI = New VisitInstance
    Call oVI.SetState(moStudy, Me, oVIState, moArezzo)
    mcolVisitInstances.Add oVI, oVIState.Key
    
    ' Reset the Schedule to force a refresh
    Call EmptyTheSchedule       ' New routine - NCJ 24 Mar 04
    
    ' Tidy up after ourselves
    Set oVIState = Nothing
    Set oVisit = Nothing
    Set oVI = Nothing
    
Exit Sub

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.AddNewVI"
    Exit Sub
    
End Sub

'------------------------------------------------
Private Sub EmptyTheSchedule()
'------------------------------------------------
' Reset the Schedule to force a refresh
' NCJ 25 Mar 04 - Terminate each ScheduleVisit before discarding
'------------------------------------------------
Dim oSchedV As ScheduleVisit

    If Not mcolScheduleVisits Is Nothing Then
        For Each oSchedV In mcolScheduleVisits
            Call oSchedV.Terminate
        Next
        Call CollectionRemoveAll(mcolScheduleVisits)
    End If
    Set mcolScheduleVisits = Nothing
    Set oSchedV = Nothing
    
    If Not moScheduleGrid Is Nothing Then
        Call moScheduleGrid.Terminate
    End If
    Set moScheduleGrid = Nothing

End Sub

'------------------------------------------------
Private Sub InitialiseVIState(oVIState As VisitInstanceState, oVisit As VisitRO, _
                    ByVal lVisitTaskId As Long, ByVal nCycleNo As Integer)
'------------------------------------------------
' Initialise a new VisitInstance State object
'------------------------------------------------

    With oVIState
        .VisitId = oVisit.VisitId
        .VisitTaskId = lVisitTaskId
        .CycleNo = nCycleNo
        .IsChanged = True
        .IsNew = True
        .LockStatus = eLockStatus.lsUnlocked
        .Status = eStatus.Requested
        .VisitDate = NULL_DATE
        'TA 20/08/2002: added new MIMessage statuses
        .DiscrepancyStatus = 0
        .SDVStatus = 0
        .NoteStatus = 0
    End With

End Sub

'------------------------------------------------
Private Sub InitialiseEFIState(oEFIState As eFormInstanceState, oEForm As eFormRO, _
                    ByVal lEFormTaskId As Long, ByVal nCycleNo As Integer, _
                    oVI As VisitInstance)
'------------------------------------------------
' Initialise a new EFormInstance State object
'------------------------------------------------

    With oEFIState
        .eFormTaskId = lEFormTaskId
        .EFormId = oEForm.EFormId
        .CycleNo = nCycleNo
        .VisitId = oVI.Visit.VisitId
        .VisitCycleNo = oVI.CycleNo
        .IsChanged = True
        .IsNew = True
        .eFormLabel = ""
        .eFormDate = NULL_DATE
        ' NCJ 20 Mar 02 - If visit is locked, must also lock the eForm
        ' NCJ 28 Apr 03 - We also do this if the visit is frozen
        .LockStatus = oVI.LockStatus
        .Status = eStatus.Requested
        'TA 20/08/2002: added new MIMessage statuses
        .DiscrepancyStatus = 0
        .SDVStatus = 0
        .NoteStatus = 0
    End With

End Sub

'------------------------------------------------------------------------------------
Private Sub moArezzo_Idle(bPatientDirty As Boolean)
'------------------------------------------------------------------------------------
' The Arezzo engine has finished its current spate of processing
' bPatientDirty is whether any changes were made to the Arezzo state
'------------------------------------------------------------------------------------

'Debug.Print Timer & " Arezzo Idle event"
    
    If bPatientDirty Then
        moSubjectState.IsChanged = True
    End If

End Sub

'------------------------------------------------------------------------------------
Private Sub moArezzo_NewEFormInstance(lEFormId As Long, lEFormTaskId As Long, _
                                    lVisitTaskId As Long, nCycleNo As Integer)
'------------------------------------------------------------------------------------
' Create a new EForm instance
' NCJ 20 Mar 02 - Do not do this if subject is Read-Only
'------------------------------------------------------------------------------------

    If Not Me.ReadOnly Then
        Call AddNewEFI(lEFormId, lEFormTaskId, nCycleNo, lVisitTaskId)
        'Debug.Print Timer & " NewEFormInstance " & lEFormTaskId
    End If

End Sub

'------------------------------------------------------------------------------------
Private Sub moArezzo_NewVisitInstance(lVisitId As Long, lVisitTaskId As Long, nCycleNo As Integer)
'------------------------------------------------------------------------------------
' Create a new Visit instance
' NCJ 20 Mar 02 - Do not do this if subject is Read-Only
'------------------------------------------------------------------------------------

    If Not Me.ReadOnly Then
        Call AddNewVI(lVisitId, lVisitTaskId, nCycleNo)
        'Debug.Print Timer & " NewVisitInstance " & lVisitTaskId
    End If

End Sub

'------------------------------------------------------------------------------------
Friend Function UpdateStatus() As Integer
'------------------------------------------------------------------------------------
' Update the trial subject's status
' based on the statuses of its Visit instances
' NB Do this only AFTER updating all the changed VisitInstance statuses
' NCJ 21 May 03 - Ignore requested visits if they're cycling (BUG 1595)
'------------------------------------------------------------------------------------
Dim nStatus As Integer
Dim oVI As VisitInstance
Dim bRequestedVisits As Boolean

    On Error GoTo ErrorLabel
    
    ' Store whether there are any non-cycling Requested visits
    bRequestedVisits = False
    
    If mcolVisitInstances.Count > 0 Then
        nStatus = eStatus.Requested
        For Each oVI In mcolVisitInstances
            nStatus = Max(nStatus, oVI.Status)
            ' NCJ 21 May 03 - Only consider non-cycling requested visits
            If oVI.Status = eStatus.Requested And oVI.Visit.MaxRepeats = 1 Then
                bRequestedVisits = True
            End If
        Next

        ' If nStatus = Success, check for requested Visits
        ' If there are any requested visits, set status to "missing"
        If nStatus = eStatus.Success Then
            If bRequestedVisits Then
                nStatus = eStatus.Missing
            End If
        End If

    Else
        ' There are no visit instances
        nStatus = eStatus.Success
    End If
    
    If nStatus <> moSubjectState.Status Then
        moSubjectState.Status = nStatus
        moSubjectState.IsChanged = True
    End If
    
    UpdateStatus = nStatus
    
Exit Function

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.UpdateStatus"
    Exit Function

End Function

'--------------------------------------------------------------------------------------------------
Public Function GeteFormInOtherVisit(oEFI As EFormInstance, lVisitTaskId As Long, _
                                        bAllowCyclingRequested As Boolean, bChangeData As Boolean) As EFormInstance
'--------------------------------------------------------------------------------------------------
'   ta 15/10/2002: adapted from clsWWW.RtnTaskIdInVisit
'   function returns  an eform in another visit
'   if passed eform is not found in passed visit, returns first active allowable eform in the visit
'   if no active eforms are found in the visit, returns nothing
'   NCJ 23 Jan 03 - Try and match the eForm's cycle
'   NCJ 27 Jan 03 - Do not assume all SchedVisits have visit intances!
'   NCJ 11 Feb 03 - Nor can we assume that all VisitEFormInstances are active
'--------------------------------------------------------------------------------------------------
Dim oVI As VisitInstance
Dim oNewEFI As EFormInstance
Dim oSchedVisit As ScheduleVisit

    Set oNewEFI = Nothing
    Set oVI = Me.VisitInstanceByTaskId(lVisitTaskId)
    
    ' NCJ 27 Jan 03 - Check the visit exists...
    If oVI Is Nothing Then Exit Function
    
    ' Does this eForm exist in this visit? NCJ 23 Jan 03 - Try same cycle first
    If oEFI.CycleNo > 1 Then
        Set oNewEFI = oVI.eFormInstanceByEFormCycle(oEFI.eForm, oEFI.CycleNo)
        If Not CanDisplayEForm(oNewEFI, bChangeData, True, bAllowCyclingRequested) Then
            ' Give up and start again
            Set oNewEFI = Nothing
        End If
    End If
    
    If oNewEFI Is Nothing Then
        ' We didn't match for cycle > 1, so try cycle 1
        Set oNewEFI = oVI.eFormInstanceByEFormCycle(oEFI.eForm, 1)
        If Not CanDisplayEForm(oNewEFI, bChangeData, True, bAllowCyclingRequested) Then
            ' Give up and start again
            Set oNewEFI = Nothing
        End If
        If oNewEFI Is Nothing Then
            ' We have not found our target eForm
            ' Look for the first available eForm in the target visit
            For Each oSchedVisit In Me.ScheduleVisits
                ' Check it's an active visit
                If Not oSchedVisit.VisitInstance Is Nothing Then
                    If oSchedVisit.VisitInstance.VisitTaskId = lVisitTaskId Then
                        ' We've got the right ScheduleVisit - get first EFI (if any)
                        
                        If oSchedVisit.VisitEFormInstances.Count > 0 Then
'                            Set oNewEFI = oSchedVisit.VisitEFormInstances(1).EFormInstance
                            ' NCJ 11 Feb 03 - Cannot assume first one is active!
                            ' NCJ 28 Feb 03 - GetFirstSchedVisitForm now returns the first available eform that can be opened
                            Set oNewEFI = GetFirstSchedVisitForm(oSchedVisit, bChangeData, bAllowCyclingRequested)
'                            If Not oNewEFI Is Nothing Then
'                                If Not CanDisplayEForm(oNewEFI, bChangeData, True, bAllowCyclingRequested) Then
'                                    Set oNewEFI = GetNextForm(oNewEFI, True, bAllowCyclingRequested, bChangeData)
'                                End If
'                            End If
                        End If
                        Exit For    ' Because we found our target visit
                        
                    End If
                End If
            Next
            
        End If
    End If
                        
    Set GeteFormInOtherVisit = oNewEFI  ' May be Nothing
    Set oSchedVisit = Nothing
    Set oNewEFI = Nothing
    Set oVI = Nothing

End Function

'--------------------------------------------------------------------------------------------------
Public Function GetFirstFormInNextVisit(oEFI As EFormInstance, bChangeData As Boolean) As EFormInstance
'--------------------------------------------------------------------------------------------------
' Get the first eForm in the "next" visit, depending on whether user can change data
' Returns nothing if there isn't a next visit/eForm
' (or if all the forms are Requested and bChangeData = FALSE)
'--------------------------------------------------------------------------------------------------
Dim oSchedVisit As ScheduleVisit
Dim oNextSchedVisit As ScheduleVisit
Dim i As Integer
Dim j As Integer
Dim bVisitFound As Boolean
Dim lVisitTaskId As Long
Dim oNextEFI As EFormInstance

    On Error GoTo ErrorLabel
    
    Set oNextEFI = Nothing
    
    ' Get the current visit task ID
    lVisitTaskId = oEFI.VisitInstance.VisitTaskId
    
    i = 1
    bVisitFound = False
    
    Do While i <= ScheduleVisits.Count And Not bVisitFound
        Set oSchedVisit = ScheduleVisits(i)
        If Not oSchedVisit.VisitInstance Is Nothing Then
            ' It is an active visit
            If oSchedVisit.VisitInstance.VisitTaskId = lVisitTaskId Then
                ' We've found the current visit - get the next if we can
                If i < ScheduleVisits.Count Then
                    ' Search for next suitable visit after this one
                    For j = i + 1 To ScheduleVisits.Count
                        Set oNextSchedVisit = ScheduleVisits(j)
                        ' Get the first form in this next visit (if any)
                        Set oNextEFI = GetFirstSchedVisitForm(oNextSchedVisit, bChangeData)
                        If Not oNextEFI Is Nothing Then
                            bVisitFound = True
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        i = i + 1
    Loop
    
    Set GetFirstFormInNextVisit = oNextEFI
    
    Set oNextEFI = Nothing
    Set oSchedVisit = Nothing
    Set oNextSchedVisit = Nothing
    
Exit Function
    
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.GetFirstFormInNextVisit"
    
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetFirstVisitForm(oVI As VisitInstance, bChangeData As Boolean) As EFormInstance
'--------------------------------------------------------------------------------------------------
' Get the first eForminstance in the given visit instance, if any
' If bChangeData = False, exclude Requested eForms
' Returns Nothing if no available eForms
'--------------------------------------------------------------------------------------------------
Dim oSchedVisit As ScheduleVisit

    ' First of all we need to find the matching ScheduleVisit
    For Each oSchedVisit In ScheduleVisits
        If Not oSchedVisit.VisitInstance Is Nothing Then
            ' It is an active visit
            If oSchedVisit.VisitInstance.VisitTaskId = oVI.VisitTaskId Then
                ' Hurrah we've got it
                Set GetFirstVisitForm = GetFirstSchedVisitForm(oSchedVisit, bChangeData)
                Exit For
            End If
        End If
    Next
    
    Set oSchedVisit = Nothing
    
End Function

'--------------------------------------------------------------------------------------------------
Private Function GetFirstSchedVisitForm(oSchedVisit As ScheduleVisit, bChangeData As Boolean, _
                    Optional bShowCyclingRequested As Boolean = True) As EFormInstance
'--------------------------------------------------------------------------------------------------
' Get the first eForminstance in the given schedule visit, if any
' If bChangeData = False, exclude Requested eForms
' Returns Nothing if no available eForms
' NCJ 28 Feb 03 - Make sure it's not the VisitEForm
'--------------------------------------------------------------------------------------------------
Dim oEFI As EFormInstance
Dim oVEFI As VEFInstance

    Set GetFirstSchedVisitForm = Nothing
    
    If Not oSchedVisit.VisitInstance Is Nothing Then
        ' It's an active visit - look for first suitable eForm
        For Each oVEFI In oSchedVisit.VisitEFormInstances
            Set oEFI = oVEFI.EFormInstance
            ' NCJ 28 Feb 03 - Use CanDisplayEForm to handle all eventualities
            If CanDisplayEForm(oEFI, bChangeData, True, bShowCyclingRequested) Then
                Set GetFirstSchedVisitForm = oEFI
                Exit For
            End If
            
'            ' NCJ 11 Feb 03 - Check that it's an active eForm
'            If Not oEFI Is Nothing Then
'                If Not bChangeData Then
'                    ' They can't change data so make sure it's not Requested
'                    If oEFI.Status <> eStatus.Requested Then
'                        Set GetFirstSchedVisitForm = oEFI
'                        Exit For
'                    End If
'                Else
'                    ' They can have the form whatever its status
'                    Set GetFirstSchedVisitForm = oEFI
'                    Exit For
'                End If
'            End If
    
        Next
        Set oEFI = Nothing
        Set oVEFI = Nothing
    End If
    
End Function

'-----------------------------------------------------------
Public Function GetNextForm(oEFI As EFormInstance, _
                          bForward As Boolean, bAllowCyclingRequested As Boolean, bChangeData As Boolean) As EFormInstance
Attribute GetNextForm.VB_Description = "Return the next EFI according to input parameters. Used for MoveToForm operations in MACRO. oEFI-eForm instance we are moving from, bForward-TRUE for next, FALSE for previous, bAllowCyclingRequested-return requested cycling form?, bChangeData-can user cha"
'-----------------------------------------------------------
' Return the next EFI according to input parameters.
' Used for MoveToForm operations in MACRO.
'   oEFI-eForm instance we are moving from,
'   bForward-TRUE for next, FALSE for previous,
'   bAllowCyclingRequested-return requested cycling form?,
'   bChangeData-can user change data ie. return locked requested forms?.
'   Outputs the next EFI or nothing if none found
'-----------------------------------------------------------
Dim oCurrentCell As GridCell
Dim nCheckRow As Integer
Dim oCell As GridCell
Dim nRow As Integer
Dim nCol As Integer

    On Error GoTo ErrorLabel
    
    With ScheduleGrid
        Set oCurrentCell = .CellByEFITaskId(oEFI.eFormTaskId)
        nRow = oCurrentCell.Row
        nCol = oCurrentCell.Col
    
        If bForward Then
            'go down from one below
            For nCheckRow = nRow + 1 To .RowMax
                Set oCell = .Cells(nCheckRow, nCol)
                If oCell.CellType = Active Then
                    'we're active therefore will always have an eFormInst
                    If CanDisplayEForm(oCell.eFormInst, bChangeData, bForward, bAllowCyclingRequested) Then
                        Set GetNextForm = oCell.eFormInst
                        Exit Function
                    End If
                End If
            Next
        Else
            'go up from one above
            For nCheckRow = nRow - 1 To 1 Step -1
                Set oCell = .Cells(nCheckRow, nCol)
                If oCell.CellType = Active And Not (oCell.eFormInst Is Nothing) Then
                    'we're active and have an eFormInst (if active should there always be an instance?)
                    If CanDisplayEForm(oCell.eFormInst, bChangeData, bForward, bAllowCyclingRequested) Then
                        Set GetNextForm = oCell.eFormInst
                        Exit Function
                    End If
                End If
            Next
            
        End If
    End With

    Set GetNextForm = Nothing
    
Exit Function

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.GetNextForm"
    Exit Function
    
End Function

'-----------------------------------------------------------
Private Function CanDisplayEForm(oEFI As EFormInstance, _
                                    bChangeData As Boolean, bForward As Boolean, bShowCyclingRequested As Boolean) As Boolean
'-----------------------------------------------------------
' TA 4/9/01: returns whether a form should be displayed according to the booleans passed in
' Used for MoveToForm operations in MACRO
' Input -
'   eForm Instance
'   user attributes - bChangeData, bForward, bShowCyclingRequested
' Output -
'   function - satisfies criteria for displaying?
' NCJ 23 Jan 03 - Pass in the EFI instead of its attributes (EFI can be Nothing)
' NCJ 13 Feb 03 - Return FALSE if this is the Visit eForm
'-----------------------------------------------------------
Dim bRequested As Boolean
Dim bCycling As Boolean
Dim bFirstCycle As Boolean
Dim oVEFI As EFormInstance

    CanDisplayEForm = False
    
    ' First check for Nothing
    If oEFI Is Nothing Then Exit Function
    
    ' NCJ 13 Feb 03 - Check for Visit eForm
    Set oVEFI = oEFI.VisitInstance.VisitEFormInstance
    If Not oVEFI Is Nothing Then
        ' Is it the same as this one?
        If oVEFI.eFormTaskId = oEFI.eFormTaskId Then Exit Function
    End If
    
    ' Pick up the EFI attributes
    With oEFI
        bRequested = (.Status = eStatus.Requested)
        bCycling = .VisitInstance.Visit.VisitEFormByEForm(.eForm).Repeating
        bFirstCycle = (.CycleNo = 1)
    End With

    If bRequested Then
        If bChangeData Then
            If bCycling Then
                If bForward And (bShowCyclingRequested Or bFirstCycle) Then
                    CanDisplayEForm = True
                Else 'going backwards or not first cycle or not show requested
                    CanDisplayEForm = False
                End If
            Else ' bCycling is false
                CanDisplayEForm = True
            End If 'bCycling
        Else 'bChangeData is false
            CanDisplayEForm = False
        End If 'bChangeData
    Else 'bRequested is false
        CanDisplayEForm = True
    End If 'bRequested
                                
    Set oVEFI = Nothing
                                
End Function
 
'-------------------------------------------------------------------
Public Function Reload(ByRef sLockErrMsg As String) As Boolean
'-------------------------------------------------------------------
' NCJ 26 Sept 02
' Sees if subject need reloading and returns TRUE if succesful reload,
' or FALSE if no reload necessary OR we couldn't get Lock For Save
' If FALSE, then if we couldn't get lock, sLockErrMsg is error message (otherwise empty)
' This should ONLY be called from the Subject Schedule (it splats all the Subject's object references
'-------------------------------------------------------------------
Dim bCouldLockSubjectForSave As Boolean
Dim bNeedToReloadSubject As Boolean
Dim sToken As String

    On Error GoTo ErrorLabel
    
    Reload = False
    
    'get lock - if the lock fails it must be "locked for saving" as the study and subject could not be locked at this stage
    bCouldLockSubjectForSave = LockForSave(msConnection, sLockErrMsg, sToken, bNeedToReloadSubject, _
                                            Me.UserName, Me.StudyDef.StudyId, Me.Site, Me.PersonID)
        
    If bCouldLockSubjectForSave Then
        sLockErrMsg = ""
        If bNeedToReloadSubject Then
            ' Reload the subject
            Call ReloadMe(Nothing)
            Reload = True
        End If
        'remove the save lock
        Call UnlockForSave(msConnection, sToken, Me.StudyDef.StudyId, Me.Site, Me.PersonID)
    End If

Exit Function
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.Reload"

End Function

'-------------------------------------------------------------------
Private Function ReloadMe(colKeepEFIs As Collection) As Boolean
'-------------------------------------------------------------------
' NCJ 16 Sept 02
' This is called in MultiUser mode to refresh the current user from the database,
' with the exception of the EFIs in the colKeepEFIs collection (they won't be reloaded).
' ASSUME we are inside a subject DB lock-for-saving
' Returns TRUE if reload successful (it should be!)
' NOTE All objects will be different after this function (except for what's in colKeepEFIs)
' so don't hang on to any of your references!
' NCJ 26 Sept 02 - Preserve moTimeZone through reload
'-------------------------------------------------------------------
Dim oStudySubjDS As New StudySubjectDS
Dim bLoadArezzo As Boolean
Dim oTempArezzo As Arezzo_DM
Dim sSite As String
Dim lPersonID As Long
Dim oKeptEFI As EFormInstance
Dim oTempTimeZone As Timezone

    On Error GoTo ErrorLabel

    ' Only deal with AREZZO in r/w mode
    bLoadArezzo = (mnUpdateMode = eUIUpdateMode.Read_Write)

    ' Store the references we need before tidying up
    Set oTempArezzo = moArezzo
    Set oTempTimeZone = moTimezone  ' NCJ 26 Sept 02
    sSite = Me.Site
    lPersonID = Me.PersonID
    ' Tidy up before reloading (making sure we hang on to colKeepEFIs)
    Call Me.Terminate(colKeepEFIs)
    ' Restore the Arezzo and TimeZone references
    Set moArezzo = oTempArezzo
    Set moTimezone = oTempTimeZone
    
    ' Reload the subject state
    Set moSubjectState = New StudySubjectState
    moSubjectState.State = oStudySubjDS.Load(msConnection, _
                        moStudy.StudyId, sSite, lPersonID, bLoadArezzo, True)
    Set oStudySubjDS = Nothing

    mbArezzoStateErr = False
    msArezzoErr = ""
    If bLoadArezzo Then
        ' Load the patient state into Arezzo
        On Error Resume Next
        moArezzo.PatientState = moSubjectState.ArezzoState
        If Err.Number <> 0 Then
            mbArezzoStateErr = True
            msArezzoErr = "There was an error in loading this subject's data." _
                            & vbCrLf & Err.Description
        End If
        On Error GoTo 0
    End If

    ' Redo all the collections
    Call SetUpCollections(colKeepEFIs)
    
    If Not colKeepEFIs Is Nothing Then
        ' Need to re-add Arezzo values to kept EFIs
        ' (Assume they've got Responses loaded)
        For Each oKeptEFI In colKeepEFIs
            Call oKeptEFI.Responses.ReapplyArezzo
        Next
        Set oKeptEFI = Nothing
    End If

    ' NCJ 31 Jan 03 - Also reapply any other Arezzo activities
    Call moArezzo.ReDoArezzoActivities
    
    'get the cache token and store it - this must be done so that Save knows whether to reload
    msCacheToken = MACROLOCKBS30.CacheAddSubjectRow(msConnection, Me.StudyDef.StudyId, Me.Site, Me.PersonID)
    
    ReloadMe = Not mbArezzoStateErr

    Set oTempArezzo = Nothing
    Set oTempTimeZone = Nothing

Exit Function
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.ReloadMe"

End Function

'----------------------------------------------------
Public Property Get SDVStatus() As eSDVStatus
'----------------------------------------------------
' TA 14 Aug 02
' The current "SDV status" of the subject
'----------------------------------------------------

    SDVStatus = moSubjectState.SDVStatus

End Property

'----------------------------------------------------
Public Property Let SDVStatus(ByVal enSDVStatus As eSDVStatus)
'----------------------------------------------------
' TA 14 Aug 02
' The current "SDV status" of the subject
' NB This doesn't update the "Changed" flag
'----------------------------------------------------

    moSubjectState.SDVStatus = enSDVStatus

End Property

'----------------------------------------------------
Public Property Get NoteStatus() As eNoteStatus
'----------------------------------------------------
' TA 14 Aug 02
' The current "Note status" of the subject
'----------------------------------------------------

    NoteStatus = moSubjectState.NoteStatus

End Property

'----------------------------------------------------
Public Property Let NoteStatus(ByVal enNoteStatus As eNoteStatus)
'----------------------------------------------------
' TA 14 Aug 02
' The current "Note status" of the subject
' NB This doesn't update the "Changed" flag
'----------------------------------------------------

    moSubjectState.NoteStatus = enNoteStatus

End Property

'----------------------------------------------------
Public Property Get DiscrepancyStatus() As eDiscrepancyStatus
'----------------------------------------------------
' TA 14 Aug 02
' The current "Discrepancy status" of the subject
'----------------------------------------------------

    DiscrepancyStatus = moSubjectState.DiscrepancyStatus

End Property

'----------------------------------------------------
Public Property Let DiscrepancyStatus(ByVal enDiscrepancyStatus As eDiscrepancyStatus)
'----------------------------------------------------
' TA 14 Aug 02
' The current "Discrepancy status" of the subject
' NB This doesn't update the "Changed" flag
'----------------------------------------------------

    moSubjectState.DiscrepancyStatus = enDiscrepancyStatus

End Property
 
'----------------------------------------------------
Public Property Get CacheToken() As String
'----------------------------------------------------
' The token that we got when loading this subject
'----------------------------------------------------

    CacheToken = msCacheToken

End Property

'----------------------------------------------------
Public Property Get Timezone() As Timezone
'----------------------------------------------------
' The Timezone object
'----------------------------------------------------
    
    Set Timezone = moTimezone

End Property

'----------------------------------------------------
Public Function GetEFIbyAllSubjectsKey(sKey As String) As EFormInstance

'----------------------------------------------------
' return an efi by VisitId|VisitCycle|EformId|EFormCycle key string
'or nothing if visit and eform do not exist
'----------------------------------------------------
Dim sVisitCode As String
Dim nVisitCycle As Integer
Dim sEFormCode As String
Dim neFormCycle As Integer
Dim vValues As Variant
Dim oVI As VisitInstance
Dim oEFI As EFormInstance

    vValues = Split(sKey, "|")
    
    
    On Error Resume Next
    
    sVisitCode = vValues(0)
    nVisitCycle = vValues(1)
    sEFormCode = vValues(2)
    neFormCycle = vValues(3)
        
    On Error GoTo ErrLabel
    'first catch exactly right
    For Each oVI In VisitInstances
        If LCase(oVI.Code) = LCase(sVisitCode) Then
            If oVI.CycleNo = nVisitCycle Then
                Exit For
            End If
        End If
    Next
    
    If oVI Is Nothing Then
        'no exact match - try different cycle
        For Each oVI In VisitInstances
            If LCase(oVI.Code) = LCase(sVisitCode) Then
                Exit For
            End If
        Next
    End If
            
    If oVI Is Nothing Then
        If VisitInstances.Count > 0 Then
            Set oVI = VisitInstances(1)
        Else
            GetEFIbyAllSubjectsKey = Nothing
'exit function here
            Exit Function
        End If
    End If
    
    'if we have got this far we have av oVI
        
    'first catch exactly right
    For Each oEFI In oVI.eFormInstances
        If LCase(oEFI.Code) = LCase(sEFormCode) Then
            If oEFI.CycleNo = neFormCycle Then
                Exit For
            End If
        End If
    Next
    
    If oEFI Is Nothing Then
        'no exact match - try different cycle
        For Each oEFI In oVI.eFormInstances
            If LCase(oEFI.Code) = LCase(sEFormCode) Then
                Exit For
            End If
        Next
    End If
            
    If oEFI Is Nothing Then
        If oVI.eFormInstances.Count > 0 Then
            Set oEFI = oVI.eFormInstances(1)
        End If
    End If
    
    'oefi could be nothing if there are none in this visit
    
    Set GetEFIbyAllSubjectsKey = oEFI
    
    Exit Function
    
ErrLabel:
    
    Err.Raise Err.Number, , Err.Descritpion & "|StudySubject.GetEFIbyAllSubjectsKey"
    
End Function


