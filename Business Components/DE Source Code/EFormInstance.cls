VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EFormInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: EFormInstance.cls
' Copyright InferMed Ltd 2001-2006 All Rights Reserved
' Author: Nicky Johns, InferMed, May 2001
' Purpose: EFormInstance object in MACRO 2.2 (and now 3.0)
' This represents an actual eForm in the Schedule in DE
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 24-29 May 01 - Initial development
' NCJ 18-21 Jun 01 - Further development
' NCJ 3 Jul 01 - Added Key property
' NCJ 10 Jul 01 - Use IsChanged property of State object
' NCJ 11 Jul 01 - Recognise 0 for undefined dates
' TA  27/07/01 - now has a link to Responses class rather than containing collection
' NCJ 31/7/01 - Use new Replace Data for eForm date
' TA 13/08/01 - Load and removing from memory of EFormElements and Response
'               called publicly through this class
' NCJ 16/8/01 - Added Complete property and UpdateStatus function
' NCJ 31/8/01 - RefreshSkipsAndDerivations
' NCJ 17 Sep 01 - Updated all comments
' NCJ/TA 26 Sep 01 - LoadLabTestData, called from LabCode Let
' NCJ 26 Sep 01 - Added LockStatus (previously forgotten!)
' TA 02/10/2001 - New Error Handling
' NCJ 3 Oct 01 - Added LabDescription and LabExists, and sorted out loading of lab test data
' NCJ 10 Oct 01 - Remember saved LabCode so we can reset it when cancelling the eForm
' TA 20/11/01 = Replaced moVisitInstance reference with mlVisitTaskId to prevent
'                   circular referece memory leak
' NCJ 11 Jan 02 - Must localise derived numeric values
' NCJ 18 Mar 02 - Return correct result from SetSkipsAndDerivations
' MACRO 3.0
' NCJ 5 Jul 2002 - Merged RQG stuff
' NCJ 8 Jul 2002 - Added extra error handlers in SetSkipDerivation and CheckCollectIf
' NCJ 12 Jul 02 - Part of CBB 2.2.19/12: Do not touch unobtainable derived fields if new value is empty
' NCJ 9-13 Aug 02 - Handle automatic RFCs
' TA 14 Aug 02 - New statuses
' NCJ 27 Aug 02 - New LabResultChanged in eAutoRFCChange
' MLM 02/09/02: Set form date in RefreshLabels based on new eform date element
' MLM 16/09/02: Only refresh current question's group if it's on our EFI.
' MLM 16/09/02: eFormDateString uses format from study def.
' NCJ 19/20 Sept 02 - Must be careful to set AREZZO context for each evaluation of Arezzo expressions
'       (because potentially we're flipping between an eForm and its VisitEForm)
' NCJ 24 Sept 02 - Made ReadOnly property Public
' NCJ 31 Oct 02 - Keep track of warnings and rejects during RefreshSkipsAndDerivations
'               Delete derived values which are rejected
' NCJ 8 Nov 02 - Added GetHotlinkTarget function
' NCJ 14 Nov 02 - Removed loading/saving of QGroupInstances
' NCJ 22 Nov 02 - No longer keep Group statuses always updated
' NCJ 27 Nov 02 - Create collection of skips and derivations for responses & groups (see InitAREZZOExprCollection)
'               Improved efficiency of Skips&Derivs.
' NCJ 27 Jan 03 - Added property EFIState (necessary for MultiUser bug fix)
' NCJ 11 Feb 03 - Do not get lock token for read-only EFI responses
'   Care when flipping optional questions from N/A
' NCJ 13 Feb 03 - Ignore Unobtainable mandatory questions in EFI completeness
' NCJ 14 Feb 03 - Ensure moResponses removed when terminating class
' NCJ 25 Feb 03 - Bug 861 - Must take account of local date settings for derived dates
' NCJ 1 Apr 03 - Changed RefreshEFormDate
' NCJ 30 Apr 03 - Changes to creation of new responses in QGroups
' NCJ 20 May 03 - Ensure we cancel the eFormdate if the eForm isn't saved
' NCJ 30 Jul 03 - Check for ReadOnly in RemoveResponses when doing CheckSaveOfLastRow
' NCJ 4 Dec 03 - Set eForm context only once before evaluating Skips & Derivs.
' NCJ 24/25 Mar 04 - Ensure collections cleared out in Terminate (to prevent memory leaks)
' NCJ 22 Jun 04 - Don't allow hotlinking to a visit eForm (Bug 2300)
' NCJ 30 Jun 04 - If LabCode is set to "", must remove existing lab
' NCJ 8 Jul 04 - Consider N/A responses which might have had CollectIfs removed
' NCJ 6 Sept 04 - Bug 2390 - Deal with eForm status when all qus. are "N/A" with one "Inform"
' ic 09/09/2005 added clinical coding
' NCJ 5 Jan 06 - eFormDateString now copes with Partial Dates
' NCJ 17 Jan 06 - Temp fix for Web DE - ignore partial form dates!
' NCJ 4 Dec 06 - Bug 2848 - Integers must be "localised" too in CheckDerivation
' NCJ 18 Dec 06 - Do not attempt to use AREZZO if read-only (because there may be no AREZZO)
'-------------------------------------------------------------------------------------

Option Explicit

' NCJ 9 Aug 02 - The context for which we are generating an Automatic RFC
Public Enum eAutoRFCContext
    OpeningEForm = 0
    ChangingResponse = 1
    Revalidation = 2
    ChangingLab = 3
    SettingEFormStatus = 4
    AddingGroupRow = 5
    ReloadingSubjectData = 6
End Enum

' NCJ 9 Aug 02 - The nature of a change for which we are generating an Automatic RFC
' NCJ 27 Aug 02 - Added LabResultChanged
Public Enum eAutoRFCChange
    NewDerivedValue = 0
    FieldDisabled = 1
    FieldEnabled = 2
    ValueRejected = 3
    StatusChanged = 4
    LabResultChanged = 5
End Enum

' This instance's collection of responses
Private moResponses As Responses

' The study def to which this instance belongs
Private moStudy As StudyDefRO
' The Visit Instance to which this instance belongs
Private mlVisitTaskId As Long

' The eForm defn. for this instance
Private moEForm As eFormRO

' The Arezzo instance we're using
Private moArezzo As Arezzo_DM

' The state class which holds all the properties
Private moEFIState As eFormInstanceState

' NCJ 22 Nov 01 - The Group Instances
Private mcolQGroupInstances As Collection

' Lab Description (if any), set up only when LoadResponses is called
Private msLabDescription As String

' The Lab code as it was when last saved
' (we use this when cancelling an eForm)
Private msSavedLabCode As String

'TA 15/09/2002
'Store token of lock for this EFI
'NB if there is no token then we don't have this efi read/write
Private msLockToken As String

' NCJ 31 Oct 02 - Need to collect up warnings and rejects from Refresh Skips and Derivations
Private mcolWarnDerivations As Collection
Private mcolRejectDerivations As Collection
Private mcolWarnRejectMessages As Collection

' The max no of iterations we should try in RefreshSkipsAndDerivations
' before an emergency exit
Private Const mn_MAX_ITERATIONS = 10

' NCJ 27 Nov 02 - The collection of Derivation & Collect If expressions for AREZZO
Private mcolAREZZOExprs As Collection
Private mcolGroupSkipExprs As Collection

' NCJ 20 May 03 - Save eForm and visit dates to restore if eForm cancelled
Private mdblSavedEFormDate As Double
Private mdblSavedVisitDate As Double

'ic 09/09/2005 the object holding all coded response terms
Private moCodedTermHistories As Object 'MACROCCBS30.CodedTermHistories
Private Const mCCSwitch As String = "CLINICALCODING"


'---------------------------------------------------------------------
Private Function DBLock(sCon As String, ByRef sLockErrMsg As String, ByRef sEFILockToken As String, sUser As String, _
                                lStudyId As Long, sSite As String, lSubjectId As Long, _
                                lEFITaskId As Long) As String
'---------------------------------------------------------------------
' Lock an efi.
' Returns if token if lock successful or empty string if not
' reurns the reason if the lock fails in sLockErrMsg
'---------------------------------------------------------------------
Dim sLockDetails As String
Dim sToken As String

    On Error GoTo ErrorLabel
    
    'TA 04.07.2001: use new locking
    sToken = MACROLOCKBS30.LockEFormInstance(sCon, sUser, lStudyId, sSite, lSubjectId, lEFITaskId)
    Select Case sToken
    Case MACROLOCKBS30.DBLocked.dblStudy
        sLockDetails = MACROLOCKBS30.LockDetailsStudy(sCon, lStudyId)
        If sLockDetails = "" Then
            sLockErrMsg = "This study is currently being edited by another user."
        Else
            sLockErrMsg = "This study is currently being edited by " & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblSubject
        sLockDetails = MACROLOCKBS30.LockDetailsSubject(sCon, lStudyId, sSite, lSubjectId)
        If sLockDetails = "" Then
            sLockErrMsg = "This subject is currently being edited by another user."
        Else
            sLockErrMsg = "This subject is currently being edited by " & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblEFormInstance
        sLockDetails = MACROLOCKBS30.LockDetailseFormInstance(sCon, lStudyId, sSite, lSubjectId, lEFITaskId)
        If sLockDetails = "" Then
            sLockErrMsg = "This eForm is currently being used by another user."
             sToken = ""
        Else
            If sEFILockToken = Split(sLockDetails, "|")(2) Then
                'we already have the lock - still return it though
                sToken = sEFILockToken
            Else
                sLockErrMsg = "This eForm is currently being used by " & Split(sLockDetails, "|")(0) & "."
                sToken = ""
            End If
        End If
        
    Case Else
        'hurrah, we have a lock

    End Select
    DBLock = sToken
    Exit Function
    
ErrorLabel:
        Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.LockForSave"
    Exit Function

End Function

'---------------------------------------------------------------------
Private Sub DBUnlock(sCon As String, sToken As String, _
                            lStudyId As Long, sSite As String, lSubjectId As Long, _
                                lEFITaskId As Long)
'---------------------------------------------------------------------
' Unlock the efi
'---------------------------------------------------------------------

On Error GoTo ErrorLabel
    'TA 04.07.2001: use new locking model
    If sToken <> "" Then
        'if no gsStudyToken then UnlockSubject is being called without a corresponding LockSubject being called first
        MACROLOCKBS30.UnlockEFormInstance sCon, sToken, lStudyId, sSite, lSubjectId, lEFITaskId
        'always set this to empty string for same reason as above
        sToken = ""
    End If
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.UnlockForSave"

End Sub

'------------------------------------------------------------------------
Public Property Get eForm() As eFormRO
'------------------------------------------------------------------------
' The eform object defining this eFormInstance
'------------------------------------------------------------------------

    Set eForm = moEForm
    
End Property

'------------------------------------------------------------------------
Friend Sub SetState(oStudy As StudyDefRO, _
                    oVI As VisitInstance, oEFIState As eFormInstanceState)
'------------------------------------------------------------------------
' Set the state containing our properties and initialise ourselves
'------------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    Set moStudy = oStudy
    Set moEFIState = oEFIState
    mlVisitTaskId = oVI.VisitTaskId
    Set moEForm = moStudy.eFormById(moEFIState.EFormId)
    Set moArezzo = moStudy.Arezzo_DM
    msSavedLabCode = moEFIState.LabCode     ' NCJ 10/10/01
    
Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.SetState"

End Sub

'------------------------------------------------
Public Property Get ReadOnly() As Boolean
'------------------------------------------------
' Is this eForm read-only?
' if there is no lock token we are read only
' NCJ 24 Sept 02 - Made Public
' NCJ 11 Feb 03 - Only check LockToken if Responses are loaded
'------------------------------------------------
    
    ReadOnly = moStudy.Subject.ReadOnly Or (Me.LockStatus <> eLockStatus.lsUnlocked) _
                    Or ((Not moResponses Is Nothing) And msLockToken = "")

End Property

'------------------------------------------------
Public Property Get ReadOnlyReason() As String
'------------------------------------------------
' If this form is read only then return a reason - otherwise return ""
'should not call this if the form not read only
' if there is no lock token we are read only
' NCJ 11 Feb 03 - Only check LockToken if Responses are loaded
'------------------------------------------------
Dim sMsg As String

    sMsg = ""
    If Me.ReadOnly Then
        If moStudy.Subject.ReadOnly Then
            sMsg = moStudy.Subject.ReadOnlyReason
            If sMsg = "" Then
                'user does not have data rights
                'this is a site subject being opened on the server
                '(we shoudn't have got this far if the study or subject are locked)'
                sMsg = "You do not have permission to change this subject's data."
            End If
        Else
            'first check lock status
            If Me.LockStatus <> eLockStatus.lsUnlocked Then
                sMsg = "eForm is " & LCase(Me.LockStatusString) & "."
            Else
                If msLockToken = "" And (Not moResponses Is Nothing) Then
                    'lock token empty and we have attempted to load responses
                    'someone else has this eForm locked
                    sMsg = "Another user is editing this eForm."
                End If
            End If
        End If
        
    End If
                
    ReadOnlyReason = sMsg

End Property

'------------------------------------------------
Friend Sub Cancel()
'------------------------------------------------
' This is called when user cancels an eForm.
' Just need to reset the LabCode to the saved value,
' and reset the form date if necessary
'
' MLM 12/09/02: Cancelling of form date no longer required.
' NCJ 20 May 03 - Oh yes it is!
'------------------------------------------------

    moEFIState.LabCode = msSavedLabCode
    msLabDescription = ""
    ' NCJ 18 Dec 06 - Only do this if not RO (because AREZZO might not be around)
    If Not Me.ReadOnly Then
        ' NCJ 20 May 03 - Must reset form date
        Call RefreshEFormDate(mdblSavedEFormDate, moArezzo.GetArezzoDate(mdblSavedEFormDate))
        ' And the visit date if we're the Visit eForm
        If Not Me.VisitInstance.VisitEFormInstance Is Nothing Then
            If Me.eFormTaskId = Me.VisitInstance.VisitEFormInstance.eFormTaskId Then
                Call Me.VisitInstance.RefreshVisitDate(mdblSavedVisitDate, moArezzo.GetArezzoDate(mdblSavedVisitDate))
            End If
        End If
    End If
    
End Sub

'------------------------------------------------
Public Property Get Arezzo() As Arezzo_DM
'------------------------------------------------
' The Arezzo_DM instance we're using
'------------------------------------------------

    Set Arezzo = moArezzo

End Property

'------------------------------------------------
Friend Sub SetContext()
'------------------------------------------------
' Set the context to this eForm instance for all future Arezzo evaluations.
' Should be done before performing any validations or derivations
' or before calling RefreshLabels.
'------------------------------------------------

    On Error GoTo ErrLabel
    Call moArezzo.SetMACROTaskContext(VisitInstance, Me)

Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.SetContext"
    
End Sub

'-------------------------------------
Public Property Get Changed() As Boolean
'-------------------------------------
' TRUE if eFormInstance has changed (or is new).
'-------------------------------------
' NCJ 6/8/01 - Check Responses too
'-------------------------------------

    Changed = (moEFIState.IsChanged Or moEFIState.IsNew)

End Property

'-------------------------------------
Friend Property Get Key() As String
'-------------------------------------
' Get unique string key for object.
'-------------------------------------

    Key = moEFIState.Key
    
End Property

'-------------------------------------
Public Property Get Responses() As Responses
'-------------------------------------
' The collection of Response objects
' belonging to this eForm instance.
' Not available until after a call to StudySubject.LoadResponses
'-------------------------------------

    Set Responses = moResponses
    
End Property

'-------------------------------------
Public Property Get QGroupInstances() As Collection
'-------------------------------------
' The collection of QGroupInstance objects belonging to this EFI
'-------------------------------------

    Set QGroupInstances = mcolQGroupInstances

End Property

'-------------------------------------
Public Property Get Code() As String
'-------------------------------------
' The unique text code of the eForm for this instance.
'-------------------------------------

    Code = moEForm.Code

End Property

'-------------------------------------
Public Property Get Name() As String
'-------------------------------------
' The name of the eForm for this instance.
'-------------------------------------

    Name = moEForm.Name

End Property

'-------------------------------------
Public Property Get eFormTaskId() As Long
'-------------------------------------
' The unique numeric ID for this eForm instance
'-------------------------------------

    eFormTaskId = moEFIState.eFormTaskId

End Property

'------------------------------------------------------
Public Property Get VisitInstance() As VisitInstance
'------------------------------------------------------
' The visit instance to which this eForm instance belongs
'-------------------------------------

    Set VisitInstance = moStudy.Subject.VisitInstanceByTaskId(mlVisitTaskId)

End Property

'-------------------------------------
Public Property Get CycleNo() As Integer
'-------------------------------------
' The cycle number of this instance.
' Will be 1 for non-cycling eForms.
'-------------------------------------
    
    CycleNo = moEFIState.CycleNo

End Property

'-------------------------------------
Public Property Get eFormLabel() As String
'-------------------------------------
' The calculated eForm label for this instance.
'-------------------------------------

    eFormLabel = moEFIState.eFormLabel

End Property

'-------------------------------------
Public Property Get eFormDate() As Double
'-------------------------------------
' The calculated date for this instance.
' Will be 0 if undefined
'-------------------------------------

    eFormDate = moEFIState.eFormDate

End Property

'-----------------------------------------------------------
Public Property Get eFormDateString() As String
'-----------------------------------------------------------
' Prepare the form date for display
' Will be empty string if date is undefined
'
' MLM 16/09/02: Use the format defined in the study def.
' NCJ 21 Jan 03 - Use local format string if available
' NCJ 5 Jan 06 - Be prepared for partial dates!
' NCJ 17 Jan 06 - Temp fix for Web DE - ignore partial dates!
' NCJ 23 Jan 06 - Partial dates back in
'-----------------------------------------------------------
Dim oDateElement As eFormElementRO
Dim sFormat As String

' NCJ 17 Jan 05 - Temporarily look for partial dates
Const MAX_YEAR = 2695
    
    'start by guessing that a form date should not be displayed
    eFormDateString = ""
    
    ' If no form date, nothing to do
    If EFIState.eFormDate = 0 Then Exit Property
    
    ' A form date has been collected
    Set oDateElement = moEForm.EFormDateElement
    If Not oDateElement Is Nothing Then
        ' The EFI also has a question defining the format to use
        ' but local format overrides question format
        sFormat = moStudy.Subject.LocalDateTypeFormat(oDateElement.DateTimeType)
        If sFormat = "" Then
            ' Use the element's own format
            sFormat = oDateElement.VBFormat
        End If
        eFormDateString = VBFormatPartialDate(EFIState.eFormDate, sFormat)
        
'        ' NCJ 17 Jan 06 - Temporary check for partial dates
'        If DatePart("yyyy", CDate(EFIState.eFormDate)) < MAX_YEAR Then
'            ' Use VB formatting (can handle partial sFormat)
'            eFormDateString = Format(EFIState.eFormDate, sFormat)
'        Else
'            ' NCJ 5 Jan 06 - Go through AREZZO in case of partial dates
'            ' But ignore errors if no AREZZO (e.g. in Web Schedule)
'            On Error Resume Next
'            eFormDateString = moArezzo.FormatDoubleDate(EFIState.eFormDate, sFormat)
'        End If
        Set oDateElement = Nothing
    End If


End Property

'-----------------------------------------------------------
Public Property Get LabCode() As String
'-----------------------------------------------------------
' The LabCode for this eForm instance.
' Setting the LabCode causes the lab test validation information
' to be loaded for all loaded responses.
'-----------------------------------------------------------

    LabCode = moEFIState.LabCode
    
End Property

'-----------------------------------------------------------
Public Property Let LabCode(sLabCode As String)
'-----------------------------------------------------------
' The LabCode for this eForm instance.
' Setting the LabCode causes the lab test validation information
' to be loaded for all loaded responses (if any).
' Assume this eForm's elements and responses are loaded.
' NCJ 30 Jun 04 - LoadLabTestData removes current lab if sLabCode = ""
'-----------------------------------------------------------

    On Error GoTo ErrLabel

    ' NCJ 20 Nov 01 - Do nothing if subject is read-only
    If Me.ReadOnly Then Exit Property
    
    If sLabCode <> moEFIState.LabCode Then
        ' See if it's a valid lab and load it if OK (deals with empty labcode)
        If LoadLabTestData(moStudy.Subject.ConnectionString, sLabCode) Then
            moEFIState.LabCode = sLabCode
            moEFIState.IsChanged = True
        End If
    End If
    
Exit Property

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.LabCode(Let)"

End Property

'-----------------------------------------------------------
Public Property Get LabDescription() As String
'-----------------------------------------------------------
' The Lab description for this eForm instance.
' May be empty string if no lab is associated with this eForm,
' or if the lab definition hasn't been loaded yet
'-----------------------------------------------------------

    LabDescription = msLabDescription
    
End Property

'-----------------------------------------------------------
Public Property Get LabExists() As Boolean
'-----------------------------------------------------------
' Whether a valid Lab exists for this eForm instance.
' NB A Lab definition is only loaded when LoadResponses is called.
'-----------------------------------------------------------

    LabExists = (msLabDescription > "")

End Property

'-------------------------------------
Public Property Get ResponsesLoaded() As Boolean
'-------------------------------------
' Is this EFI's collection of Response objects loaded?
'-------------------------------------

    ResponsesLoaded = Not (moResponses Is Nothing)
    
End Property

'------------------------------------------------
Friend Function LoadResponses(sConnection As String, ByRef sLockErrMsg As String, ByRef sEFILockToken As String) As Boolean
'------------------------------------------------
' Get collection of question responses
' for this eform instance.
' This automatically loads the eForm's elements too,
' and the lab test data if appropriate.
' And the QGroupInstances too (NCJ 21 Nov 01)
'true means read/write
' false if can't get lock (or subject already readonly) and his loaded readonly
'- reason is returned in sLockErrMsg
' NCJ 27 Nov 02 - Create collection of skips and derivations for responses
' NCJ 11 Feb 03 - Do not create lock token if EFI is read-only
' ic 09/09/2005 added clinical coding
'------------------------------------------------
Dim oResponse As Response
Dim oElement As eFormElementRO
Dim oVersion As MACROVersion.Checker
Dim bCC As Boolean


    On Error GoTo ErrLabel
    
    'check for clinical coding version
    Set oVersion = New MACROVersion.Checker
    bCC = oVersion.HasUpgrade(sConnection, mCCSwitch)
    
    'load elements first
    moEForm.LoadElements sConnection
    
    ' NCJ 11 Feb 03 - Check for EFI's read-onlyness
'    If Not moStudy.Subject.ReadOnly Then
    If Not Me.ReadOnly Then
        'try to lock if we are not read only
        With moStudy.Subject
            msLockToken = DBLock(sConnection, sLockErrMsg, sEFILockToken, .UserName, _
                            .StudyDef.StudyId, .Site, .PersonID, Me.eFormTaskId)
        End With
    End If
    
    Set moResponses = New Responses
    moResponses.Load sConnection, Me
    
    ' Load the Group Instances
    ' NCJ 14 Nov 02 - Actually we create them now rather than load them
    Call CreateGroupInstances
    
    ' NCJ 27 Nov 02 - Generate the collection of AREZZO Expressions
    ' NB It's important to do it here BEFORE CreateNewResponses
    ' (because the new ones will get added to the collection afterwards)
    Call InitAREZZOExprCollection
    
    ' NCJ 20 Nov 01 - Create all the necessary new Response objects here
    Call CreateNewResponses

    ' Load the lab test data (if any)
    If moEFIState.LabCode > "" Then
        LoadLabTestData sConnection, moEFIState.LabCode
    End If
    
    ' NCJ 20 May 03 - Remember the current eForm/visit dates in case we need to restore them later
    mdblSavedEFormDate = moEFIState.eFormDate
    mdblSavedVisitDate = Me.VisitInstance.VisitDate
    
    'return token for use later on
    sEFILockToken = msLockToken
    'this will be blank if we are readonly (ie lock failed)
    
    If (bCC) Then
        'ic 09/09/2005 load the coded response object
        Call LoadCodedTermHistories(sConnection)
    End If
    
    'retuns if read only ie subject locked/eform locked/couldn't get lock
    LoadResponses = Not Me.ReadOnly

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.LoadResponses"
    
End Function

'------------------------------------------------
Private Sub CreateGroupInstances()
'------------------------------------------------
' NCJ 21 Nov 01 - Load the group instances
' NCJ 14 Nov 02 - Do not load from the database but just create in memory
'------------------------------------------------
Dim oQGroup As QGroupRO
Dim oQGI As QGroupInstance

    On Error GoTo ErrLabel
    
    ' NCJ 14 Nov 02 - Generate the QGroup instances for the eForm's QGroups
    ' (this is instead of loading QGroupInstances from the DB)
    Set mcolQGroupInstances = New Collection
    Set mcolGroupSkipExprs = New Collection     ' The collection of AREZZO skip expressions
    For Each oQGroup In Me.eForm.QGroups
        Set oQGI = New QGroupInstance
        Call oQGI.Init(oQGroup, Me)
        mcolQGroupInstances.Add oQGI, Str(oQGroup.QGroupID)
        ' NCJ 27 Nov 02 - Add any skip condition
        If oQGroup.Element.CollectIfCond > "" Then
            mcolGroupSkipExprs.Add oQGroup.Element.CollectIfCond
        End If
    Next
    
    Set oQGI = Nothing
    Set oQGroup = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.CreateGroupInstances"

End Sub

'------------------------------------------------
Private Sub CreateNewResponses()
'------------------------------------------------
' NCJ 20 Nov 01 - Create new responses for all eform elements that need them,
' including repeating group rows
'------------------------------------------------
Dim oElement As eFormElementRO
Dim oQGI As QGroupInstance
'Dim nRow As Integer

    On Error GoTo ErrLabel
    
    ' Can't do it if the eForm is ReadOnly
    If Me.ReadOnly Then Exit Sub
    
    For Each oElement In moEForm.EFormElements
        ' Ignore non-questions and group questions
        If oElement.QuestionId > 0 Then
            If oElement.OwnerQGroup Is Nothing Then
                ' For non-group elements, row = 1
                Call GetResponse(oElement, 1)
            End If
'        ElseIf Not oElement.QGroup Is Nothing Then
'            ' Process elements within a group
'            ' Get the GroupInstance
'            Set oQGI = GetQGInstance(oElement)
'            ' Ensure we have responses for each row
'            For nRow = 1 To oQGI.Rows
'                Call GetRowResponses(oQGI.QGroup, nRow)
'            Next
        End If
    Next
    
    ' NCJ 30 Apr 03 - Do the QGroup responses separately here
    For Each oQGI In mcolQGroupInstances
        Call oQGI.CreateNewResponses
    Next
    
    Set oElement = Nothing
    Set oQGI = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.CreateNewResponses"

End Sub

'-------------------------------------------------------------------
Friend Sub GetRowResponses(oQG As QGroupRO, nRow As Integer)
'-------------------------------------------------------------------
' Make sure there are responses for every row in a QGroup
'-------------------------------------------------------------------
Dim oElement As eFormElementRO

    For Each oElement In oQG.Elements
        If oElement.QuestionId > 0 Then
            Call GetResponse(oElement, nRow)
        End If
    Next

    Set oElement = Nothing

End Sub

'------------------------------------------------
Private Sub GetResponse(oElement As eFormElementRO, nRow As Integer)
'------------------------------------------------
' Make sure there's a Response for this element and row
' Assume we're allowed to do it
' NCJ 27 Nov 02 - Maintain collection of AREZZO expressions for each response
' ic 28/10/2005 added clinical coding
'------------------------------------------------
Dim oResponse As Response

    Set oResponse = moResponses.ResponseByElement(oElement, nRow)
    If oResponse Is Nothing Then
        ' We need to create a new one
        Set oResponse = moResponses.AddNew(oElement, nRow)
        ' NCJ 27 Nov 02 - Add its AREZZO expressions to our collection
        Call AddResponseAREZZOExprs(oResponse)
        
        If (oElement.DataType = eDataType.Thesaurus) Then
            If (Not (moCodedTermHistories Is Nothing)) Then
                If (Not moCodedTermHistories.Exists(CLng(oResponse.ResponseId), CInt(oResponse.RepeatNumber))) Then
                    Dim oCodedTermHistory As MACROCCBS30.CodedTermHistory
                    
                    'the response has not been saved before, create a blank CodedTermHistory object
                    Set oCodedTermHistory = New MACROCCBS30.CodedTermHistory
            
                    Call oCodedTermHistory.InitEmpty(CLng(moStudy.StudyId), CStr(moStudy.Subject.Site), CLng(moStudy.Subject.PersonID), _
                        CLng(oResponse.ResponseId), CInt(oResponse.RepeatNumber))
                    Call moCodedTermHistories.AddCodedTermHistory(oCodedTermHistory)
                End If
            End If
        End If
    End If

    Set oResponse = Nothing

End Sub

'------------------------------------------------
Public Function SaveCodedTermHistoriesNeeded() As Boolean
'------------------------------------------------
' ic 09/09/2005
' function returns true if any coded term has changed and needs saving
'------------------------------------------------
    SaveCodedTermHistoriesNeeded = moCodedTermHistories.SaveNeeded
End Function

'------------------------------------------------
Public Function CodedTermHistoryFromTaskId(lResponseTaskId As Long, nRepeat As Integer) As Object 'MACROCCBS30.CodedTermHistory
'------------------------------------------------
' ic 09/09/2005
' function returns a specified CodedTermHistory object from the collection
'------------------------------------------------
    Set CodedTermHistoryFromTaskId = moCodedTermHistories.CodedTermHistoryFromTaskId(CLng(lResponseTaskId), CInt(nRepeat))
End Function

'------------------------------------------------
Private Sub LoadCodedTermHistories(sConnection As String)
'------------------------------------------------
' ic 09/09/2005
' sub loads all saved CodedTermHistory objects from the database and creates empty
' ones for those that have not yet been saved
'------------------------------------------------
    Dim oResponseArray As Object
    Dim oCodedTermHistory As MACROCCBS30.CodedTermHistory
    Dim oResponse As Response
    Dim n As Integer

    On Error GoTo ErrLabel
    Set moCodedTermHistories = New MACROCCBS30.CodedTermHistories
    Call moCodedTermHistories.InitAuto(sConnection, CLng(moStudy.Subject.StudyId), moStudy.Subject.Site, _
        CLng(moStudy.Subject.PersonID), CLng(Me.eForm.EFormId), CLng(Me.CycleNo))

    For Each oResponse In moResponses
        If oResponse.Element.DataType = eDataType.Thesaurus Then
            If (Not moCodedTermHistories.Exists(CLng(oResponse.ResponseId), CInt(oResponse.RepeatNumber))) Then
                'the response has not been saved before, create a blank CodedTermHistory object
                Set oCodedTermHistory = New MACROCCBS30.CodedTermHistory

                Call oCodedTermHistory.InitEmpty(CLng(moStudy.StudyId), CStr(moStudy.Subject.Site), CLng(moStudy.Subject.PersonID), _
                    CLng(oResponse.ResponseId), CInt(oResponse.RepeatNumber))
                Call moCodedTermHistories.AddCodedTermHistory(oCodedTermHistory)
            End If
        End If
    Next
    Exit Sub
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.LoadCodedTermHistories"
End Sub

'------------------------------------------------
Friend Sub SaveCodedTermHistories(ByVal sConnection As String)
'------------------------------------------------
'  ic 09/09/2005
' sub saves any changed CodedTermHistory objects
'------------------------------------------------
    On Error GoTo ErrLabel

    Call moCodedTermHistories.Save(sConnection, CLng(Me.VisitInstance.VisitTaskId), CInt(Me.VisitInstance.CycleNo), _
        CLng(Me.eForm.EFormId), CInt(Me.CycleNo))
    Exit Sub

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.SaveResponses"
End Sub

'------------------------------------------------
Friend Sub SaveResponses(sConnection As String)
'------------------------------------------------
' Save responses (if any)
' NCJ 6th Aug 01 - The Responses class knows all the values it needs
' (because it remembers them from the Load).
' Must also confirm form to Arezzo (to trigger next cycles)
' TA 5/9/01: only saves if there are changes
' NCJ 29 Apr 03 - Added check for last rows of QGroups
'------------------------------------------------
Dim oQGI As QGroupInstance

    On Error GoTo ErrLabel
    
    ' NCJ 20 Nov 01 - Do nothing if subject is read-only
    If Me.ReadOnly Then Exit Sub
    
    ' NCJ 29 Apr 03 - Check whether we need to save last row of each of the QGroup instances
    ' This sets the "Ignore me" flag and removes values from AREZZO as required
    For Each oQGI In mcolQGroupInstances
        oQGI.CheckSaveOfLastRow
    Next
    
    If moResponses.Changed Then
        moResponses.Save sConnection
    End If

    ' NCJ 10/10/01 - Reset our SavedLabCode
    ' NCJ 20 May 03 - And remember the eForm/visit dates
    msSavedLabCode = moEFIState.LabCode
    mdblSavedEFormDate = moEFIState.eFormDate
    mdblSavedVisitDate = Me.VisitInstance.VisitDate

Exit Sub

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.SaveResponses"

End Sub

'------------------------------------------------
Friend Sub RemoveResponses(ByVal bRemoveLock As Boolean)
'------------------------------------------------
' Remove responses from memory.
' This also removes the eFormElements from the eForm.
' Don't remove the eForm lock if bRemoveLock is FALSE
'------------------------------------------------
Dim oQGI As QGroupInstance

    On Error GoTo ErrLabel
    
    ' NCJ 7 May 03 - Make sure we remove unsaved AREZZO responses from last row of QGroup instances
    ' NCJ 30 Jul 03 - Must not do this for read-only eForms (because last row of responses won't exist)
    If Not Me.ReadOnly Then
        For Each oQGI In mcolQGroupInstances
            oQGI.CheckSaveOfLastRow
        Next
    End If
    
    Call moEForm.RemoveElements
    Set moResponses = Nothing
    
    Call CollectionRemoveAll(mcolQGroupInstances)
    Set mcolQGroupInstances = Nothing   ' Added NCJ 14 Nov 02
    
    Set mcolAREZZOExprs = Nothing       ' Added NCJ 27 Nov 02
    Set mcolGroupSkipExprs = Nothing
    
    If msLockToken <> "" And bRemoveLock Then
        'remove the lock on the form
        With moStudy.Subject
            Call DBUnlock(moStudy.ConnectionString, msLockToken, _
                            .StudyDef.StudyId, .Site, .PersonID, Me.eFormTaskId)
        End With
        'reset the lock token
        msLockToken = ""
        
    End If

    Set oQGI = Nothing

Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.RemoveResponses"

End Sub

'-------------------------------------
Friend Sub RefreshEFormDate(ByVal dblDate As Double, ByVal sArezzoDate As String)
'-------------------------------------
' NCJ 1 Apr 03
' Update an eForm date to this new value, e.g. when the user enters it on an eForm
' dblDate is the new value (may be 0), and sArezzoDate is the Arezzo date term (may be "")
' Assume we're not Read-only!
'-------------------------------------
Dim sUserTag As String
Dim sFormDateName As String

    ' Only reset if it's changed
    If dblDate <> moEFIState.eFormDate Then
        moEFIState.eFormDate = dblDate
        moEFIState.IsChanged = True
        ' Add it to Arezzo
        sUserTag = CStr(Me.eFormTaskId * 100000)
        If dblDate <> 0 Then
            sFormDateName = VisitInstance.Code & ":" & moEForm.Code & ":date"
            Call moArezzo.ReplaceQData(sFormDateName, sArezzoDate, sUserTag)
        Else
            ' Just remove previous value
            Call moArezzo.RemoveDataByRef(sUserTag)
        End If
    End If

End Sub

'-------------------------------------
Public Sub RefreshLabels()
'-------------------------------------
' Recalculate the eForm's label and date values.
'-------------------------------------
' MLM 02/09/02: Set form date based on new form date question instead
' NCJ 1 Apr 03 - Don't do the eForm date here any more
' NCJ 29 May 03 - Set changed flag if label changed to empty
'-------------------------------------
Dim sLabel As String
'Dim dblDate As Double

    On Error GoTo ErrLabel
    
    ' NCJ 20 Nov 01 - Do nothing if subject is read-only
    If Me.ReadOnly Then Exit Sub
    
    If moEForm.LabelExpr > "" Then
        ' Recalculate label
        ' Set the context for Arezzo
        Call SetContext
        sLabel = moArezzo.EvaluateExpression(moEForm.LabelExpr)
        ' Check it was calculated OK
        If Not moArezzo.ResultOK(sLabel) Then
            sLabel = ""
        End If
    Else
        ' No eForm label
        sLabel = ""
    End If
    
    ' Only reset if it's changed
    If sLabel <> moEFIState.eFormLabel Then
        moEFIState.eFormLabel = sLabel
        moEFIState.IsChanged = True
    End If
    
    ' Reset the context for Arezzo
    Call moArezzo.SetMACROTaskContext(Nothing, Nothing)

Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.RefreshLabels"
    
End Sub

'-------------------------------------------------------
Public Property Get Status() As Integer
'-------------------------------------------------------
' The calculated status of this instance, as an integer.
' See also the StatusString property.
'-------------------------------------------------------

    Status = moEFIState.Status

End Property

'----------------------------------------------------
Public Property Get StatusString() As String
'----------------------------------------------------
' The status of this eForm instance, as a string (for display to the user).
'----------------------------------------------------

    StatusString = GetStatusString(moEFIState.Status)

End Property

'-------------------------------------------------------
Public Property Get LockStatus() As Integer
'-------------------------------------------------------
' The lock status of this instance.
' See also the LockStatusString property.
'-------------------------------------------------------

    LockStatus = moEFIState.LockStatus

End Property

'----------------------------------------------------
Public Property Get LockStatusString() As String
'----------------------------------------------------
' The lock status of this eForm instance, as a string (for display to the user).
'----------------------------------------------------

    LockStatusString = GetLockStatusString(moEFIState.LockStatus)

End Property


'-----------------------------------------------------------
Private Sub Class_Initialize()
'-----------------------------------------------------------

    Call Terminate
    Set mcolWarnDerivations = New Collection
    Set mcolRejectDerivations = New Collection
    Set mcolWarnRejectMessages = New Collection
    Set mcolAREZZOExprs = New Collection      ' Added NCJ 27 Nov 02
    Set mcolGroupSkipExprs = New Collection
    
End Sub

'-----------------------------------------------------------
Private Sub Class_Terminate()
'-----------------------------------------------------------
       
    Call Terminate

End Sub

'-----------------------------------------------------------
Friend Sub Terminate()
'-----------------------------------------------------------
' Set all module vars to nothing
' NCJ 24/25 Mar 04 - Explicitly clear out all collections (prevent possible memory leaks)
'-----------------------------------------------------------
Dim oQGI As QGroupInstance

    Set moStudy = Nothing
    Set moEForm = Nothing
    Set moArezzo = Nothing
    Set moEFIState = Nothing

    If Not mcolQGroupInstances Is Nothing Then
        For Each oQGI In mcolQGroupInstances
            Call oQGI.Terminate
        Next
        Call CollectionRemoveAll(mcolQGroupInstances)
    End If
    Set mcolQGroupInstances = Nothing
    Set oQGI = Nothing
    
    Call CollectionRemoveAll(mcolWarnDerivations)
    Set mcolWarnDerivations = Nothing
    
    Call CollectionRemoveAll(mcolRejectDerivations)
    Set mcolRejectDerivations = Nothing
    
    If Not moResponses Is Nothing Then
        Call moResponses.Terminate
    End If
    Set moResponses = Nothing

    ' These collections don't contain objects
    Set mcolWarnRejectMessages = Nothing
    Set mcolAREZZOExprs = Nothing       ' Added NCJ 27 Nov 02
    Set mcolGroupSkipExprs = Nothing
    
    'ic 14/09/2005 clear clinical coding object
    Set moCodedTermHistories = Nothing
End Sub

'------------------------------------------------
Public Property Get Complete() As Boolean
'------------------------------------------------
' TRUE if this eForm's collection of responses is "complete",
' i.e. all mandatory questions answered.
' Returns FALSE if no responses are currently loaded
' NCJ 7 Jan 02 - Changed to correctly take account of Question Groups
' NCJ 13 Feb 03 - Ignore Unobtainable mandatory questions
'------------------------------------------------
Dim oResponse As Response
Dim bComplete As Boolean
Dim oElement As eFormElementRO
Dim oQGI As QGroupInstance

    On Error GoTo ErrLabel
    
    If moResponses Is Nothing Then
        Complete = False
        Exit Property
    End If
    
    bComplete = True
    ' Look for a mandatory item that's missing
    For Each oResponse In moResponses
        ' Only look at non-group questions here
        If oResponse.Element.OwnerQGroup Is Nothing Then
            Select Case oResponse.Status
            ' Look at Missing, Requested
            ' (but don't consider NotApplicable - or Unobtainable)
            Case eStatus.Missing, eStatus.Requested
                ' NCJ 7 May 03 - Use Question ID rather than ElementId
                Set oElement = moEForm.eFormElementByQuestionId(oResponse.Element.QuestionId)
                If oElement.IsMandatory Then
                    ' We found a mandatory question with no answer
                    bComplete = False
                    Exit For
                End If
            Case Else
            End Select
        End If
    Next
    
    Set oResponse = Nothing
    Set oElement = Nothing
    
    ' If all mandatory questions have been answered, check the mandatory group instances
    If bComplete Then
        For Each oQGI In mcolQGroupInstances
            If oQGI.QGroup.Element.IsMandatory Then
                If Not oQGI.Complete Then
                    ' We found a mandatory group that's not complete
                    bComplete = False
                    Exit For
                End If
            End If
        Next
        Set oQGI = Nothing
    End If
    
    Complete = bComplete

Exit Property

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.Complete(Get)"
    
End Property

'------------------------------------------------
Friend Function UpdateStatus() As Integer
'------------------------------------------------
' Update this eForm instance's status
' based on the statuses of its responses
' and its Group Instances
' NB Do this BEFORE calculating its Visit Instance's status
' NCJ 6 Sept 04 - Bug 2390 - Deal with all "N/A" with one "Inform"
'------------------------------------------------
Dim nStatus As Integer
Dim oResponse As Response
Dim oElement As eFormElementRO
Dim nMaxOptionalStatus As Integer
Dim nMaxNonOptionalStatus As Integer
Dim bInform As Boolean
Dim oQGI As QGroupInstance

    On Error GoTo ErrLabel
    
    ' NCJ 20 Nov 01 - Do nothing if subject is read-only
    If Me.ReadOnly Then
        UpdateStatus = Me.Status
        Exit Function
    End If
    
    nStatus = eStatus.Success
    bInform = False     ' Initialise as no "Inform" status responses
    
    If moResponses.Count > 0 Then
        ' Initialise to "requested" (lowest value)
        nMaxNonOptionalStatus = eStatus.Requested
        nMaxOptionalStatus = eStatus.Requested
        For Each oResponse In moResponses
            ' We deal with group elements later...
            If oResponse.Element.OwnerQGroup Is Nothing Then
                ' Exclude Inform status on this first pass
                If oResponse.Status <> eStatus.Inform Then
                    ' NCJ 7 May 03 - Index by QuestionId (more reliable than ElementId)
                    Set oElement = moEForm.eFormElementByQuestionId(oResponse.Element.QuestionId)
                    ' Deal with optional questions separately
                    If oElement.IsOptional Then
                        nMaxOptionalStatus = Max(nMaxOptionalStatus, oResponse.Status)
                    Else
                        nMaxNonOptionalStatus = Max(nMaxNonOptionalStatus, oResponse.Status)
                    End If
                Else
                    bInform = True
                End If
            End If
        Next
    
        ' NCJ 21 Nov 01 - Now include the GroupInstance statuses
        ' (these won't have Inform status)
        For Each oQGI In mcolQGroupInstances
            ' Update its status
            Call oQGI.UpdateStatus
            Set oElement = moEForm.eFormElementById(oQGI.QGroup.ElementID)
            If oElement.IsOptional Then
                nMaxOptionalStatus = Max(nMaxOptionalStatus, oQGI.Status)
            Else
                nMaxNonOptionalStatus = Max(nMaxNonOptionalStatus, oQGI.Status)
            End If
        Next
        
        ' Now sort out which value to use
        If nMaxNonOptionalStatus > nMaxOptionalStatus Then
            nStatus = nMaxNonOptionalStatus
        ElseIf (nMaxNonOptionalStatus > eStatus.Requested And nMaxOptionalStatus = eStatus.Missing) Then
            nStatus = nMaxNonOptionalStatus
        Else
            nStatus = nMaxOptionalStatus
        End If
         
        ' If status is Requested it means
        ' we didn't find any questions that weren't "Inform" on the first pass
        ' so check if there are any "Inform" questions
        ' NCJ 6 Sept 04 - Bug 2390 - Also consider Not Applicable, Unobtainable
        If bInform Then
            Select Case nStatus
            Case eStatus.Requested, eStatus.NotApplicable, eStatus.Unobtainable
                nStatus = eStatus.Success
            End Select
        End If
'        If nStatus = eStatus.Requested And bInform Then
'            nStatus = eStatus.Success
'        End If
        
    End If
 
    ' Update the state object
    If moEFIState.Status <> nStatus Then
        moEFIState.Status = nStatus
        moEFIState.IsChanged = True
    End If
    ' Return the result
    UpdateStatus = nStatus

Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.UpdateStatus"
    
End Function

'-------------------------------------------------------------------------------
Public Sub DerivationWarnRejects(ByRef colWarnResponses As Collection, _
                                ByRef colRejectResponses As Collection, _
                                ByRef colErrMessages As Collection)
'-------------------------------------------------------------------------------
' NCJ 31 Oct 02 - The collection of responses that had warnings or were rejected
' in the most recent call to RefreshSkipsAndDerivations
' colErrMessages will be the corresponding collection of error/warning messages
'-------------------------------------------------------------------------------

    Set colWarnResponses = mcolWarnDerivations
    Set colRejectResponses = mcolRejectDerivations
    Set colErrMessages = mcolWarnRejectMessages

End Sub

'-------------------------------------------------------------------------------
Public Function RefreshSkipsAndDerivations(ByVal enContext As eAutoRFCContext, _
                                ByVal sUserName As String, _
                                Optional oChangedResponse As Response = Nothing) As Boolean
'-------------------------------------------------------------------------------
' This recalculates all derived values and Collect If conditions
' and sets status of Skipped elements to "Not Applicable".
' Nothing is done if the the subject is read-only.
' oResponse is the response that's just changed (if any):
' if oResponse is given, we update the status of its group if appropriate.
' Returns TRUE if anything changed as a result of these calculations.
'-------------------------------------------------------------------------------
' NCJ 7 Jan 02 - Added optional oResponse parameter; return Boolean result
' NCJ 13 Aug 02 - Added sUserName and enContext parameters
' MLM 16/09/02: Only refresh current question's group if it's on our EFI.
'-------------------------------------------------------------------------------
Dim i As Integer
Dim oQGI As QGroupInstance
    
    On Error GoTo ErrLabel
    
    RefreshSkipsAndDerivations = False
    
    ' NCJ 31 Oct 02 - We have to keep a record of derived warnings and rejects
    ' so they can be popped up to the user
    Set mcolWarnDerivations = New Collection
    Set mcolRejectDerivations = New Collection
    Set mcolWarnRejectMessages = New Collection
    
    ' NCJ 20 Nov 01 - Do nothing if subject is read-only or if eForm is locked
    If Me.ReadOnly Then Exit Function
    
    ' Se MACRO Visit/eForm context
    ' NCJ 20 Sept 02 - Don't need this here now because it's done
    ' in every SetQuestionContext
    ' Call moArezzo.SetMACROTaskContext(VisitInstance, Me)
    
    i = 0
    Do While SetAllSkipsAndDerivations(enContext, sUserName, oChangedResponse) _
     And i < mn_MAX_ITERATIONS
        i = i + 1
    Loop

'    If i >= 8 Then
'        Debug.Print Timer & " Excessive looping in SetSkips!"
'    End If
    
    ' Finally remove the last Question context
    Call moArezzo.RemoveQuestionContext
    
    ' Return TRUE if anything happened
    RefreshSkipsAndDerivations = (i > 0)
    
    Set oQGI = Nothing
    
Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.RefreshSkipsAndDerivations"

End Function

'--------------------------------------------------------------------
Private Function SetSkipDerivation(oResponse As Response, _
                                ByVal sDerivResult As String, _
                                ByVal bCollectIfResult As Boolean, _
                                ByVal enContext As eAutoRFCContext, _
                                ByVal sUserName As String, _
                                oChangedResponse As Response) As Boolean
'--------------------------------------------------------------------
' NCJ 2 Jan 02
' Deal with the "collect if" and derivation for this response
' If something changed, return TRUE, otherwise return FALSE
' Assume its owner group status (if any) is already set
' NCJ 27 Nov 02 - The Derivation and CollectIf results are now passed in
' NCJ 8 Jul 04 - Consider NA responses which might have had CollectIfs removed
'--------------------------------------------------------------------
Dim oElement As eFormElementRO
Dim oQGI As QGroupInstance

    On Error GoTo ErrLabel
    
    SetSkipDerivation = False
    
    If oResponse.LockStatus <> eLockStatus.lsUnlocked Then
        ' Can't change a locked response
        Exit Function
    End If
    
    Set oElement = oResponse.Element
    
    If Not oElement.OwnerQGroup Is Nothing Then
        ' Check to see if its owner group is NotApplicable
        Set oQGI = Me.QGroupInstanceById(oElement.OwnerQGroup.QGroupID)
        If oQGI.Status = eStatus.NotApplicable Then
            ' Must set this response to NotApplicable too
            SetSkipDerivation = CheckCollectIf(oResponse, False, _
                                    enContext, sUserName, oChangedResponse)
            ' That's all we do if the Question Group is NotApplicable
            Exit Function
        Else
            ' Reset response only if it has no Collect If of its own
            If oElement.CollectIfCond = "" Then
                SetSkipDerivation = CheckCollectIf(oResponse, True, _
                                    enContext, sUserName, oChangedResponse)
            End If
        End If
    End If
    
    ' If we get here, the response does not belong to a disabled group
    
    If (oElement.CollectIfCond = "" And oResponse.Status <> eStatus.NotApplicable) _
     And oElement.DerivationExpr = "" Then
        ' Nothing to do
        Exit Function
    End If
    
    ' At least one of Collect If and Derivation is non-empty, or its status is NA
    
    ' First see if we should disable the field
    ' NCJ 8 Jul 04 - Consider NA responses in case the CollectIf has gone
    If oElement.CollectIfCond > "" Or oResponse.Status = eStatus.NotApplicable Then
        If CheckCollectIf(oResponse, bCollectIfResult, _
                                enContext, sUserName, oChangedResponse) Then
            ' Something changed
            SetSkipDerivation = True
        End If
    End If
        
    ' Now see if it's derived (but only if field is enabled)
    If oElement.DerivationExpr > "" And oResponse.Status <> eStatus.NotApplicable Then
    
        If CheckDerivation(oResponse, oElement, oElement.DerivationExpr, sDerivResult, _
                                enContext, sUserName, oChangedResponse) Then
            SetSkipDerivation = True
        End If
    
    End If
    
'    Set colExprs = Nothing
'    Set colResults = Nothing
    Set oElement = Nothing
   
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" _
                    & "EFormInstance.SetSkipDerivation(" & oResponse.Element.Code & ")"

End Function

'---------------------------------------------------------------------------------------------
Private Function CheckCollectIf(oResponse As Response, _
                                ByVal bCollectIfResult As Boolean, _
                                ByVal enContext As eAutoRFCContext, _
                                ByVal sUserName As String, _
                                oChangedResponse As Response) As Boolean
'---------------------------------------------------------------------------------------------
' Set the status of this response according to whether bCollectIfResult is true
' Assume the response does have a collect-if condition
' Returns TRUE if the response status changed
' NCJ 11 Feb 03 - Care when flipping optional questions from N/A
'---------------------------------------------------------------------------------------------
Dim sMsg As String
Dim bChanged As Boolean
Dim sRFC As String
Dim enNewStatus As eStatus

    On Error GoTo ErrLabel
    
    sRFC = ""
    
    CheckCollectIf = False
    If bCollectIfResult Then
        ' Collect If cond is TRUE so switch from NA to Missing
        ' NCJ 11 Feb 03 - Cannot assume Missing because question might be optional...
        ' Call ValidateValue instead to do it properly
        If oResponse.Status = eStatus.NotApplicable Then
            ' Let ValidateValue work out the new status
            enNewStatus = oResponse.ValidateValue("", sMsg, bChanged)
            If oResponse.RequiresStatusRFC(enNewStatus) Then
                sRFC = AutomaticRFC(enContext, FieldEnabled, oChangedResponse)
            End If
            Call oResponse.ConfirmValue("", sRFC, sUserName)
'            Call oResponse.SetStatus(eStatus.Missing, sRFC, sUserName)
            ' Something changed
            CheckCollectIf = True
        End If
    Else
        ' Collect If cond is FALSE
        ' We must disable the field if not already Not Applicable
        If oResponse.Status <> eStatus.NotApplicable Then
            If oResponse.RequiresStatusRFC(eStatus.NotApplicable) Then
                sRFC = AutomaticRFC(enContext, FieldDisabled, oChangedResponse)
            End If
            ' Clear the field and set it to NA
            ' (assume no validation errors/warnings)
            Call oResponse.ValidateValue("", sMsg, bChanged)
            Call oResponse.ConfirmValue("", "", "")
            Call oResponse.SetStatus(eStatus.NotApplicable, sRFC, sUserName)
            ' Something changed
            CheckCollectIf = True
        End If
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.CheckCollectIf"

End Function

'--------------------------------------------------------
Private Function SetAllSkipsAndDerivations(ByVal enContext As eAutoRFCContext, _
                                ByVal sUserName As String, _
                                oChangedResponse As Response) As Boolean
'--------------------------------------------------------
' NCJ 2 Jan 02 - New for Repeating Question Groups
' Iterate through all the skips and derivations
' and return TRUE if anything was changed
' enContext is context, sUserName is name of user doing refresh,
' and oresponse is response (if any) being changed
' NCJ 27 Nov 02 - We evaluate all the derivs. & skips from mcolAREZZOExprs
' NCJ 4 Dec 03 - Set eForm context once before evaluating
'---------------------------------------------------------
Dim oResponse As Response
Dim bSomethingChanged As Boolean
Dim colResults As Collection
Dim i As Long
Dim sDerivResult As String
Dim bCollectIfResult As Boolean

'Debug.Print "  " & Timer & " BEGIN SetAllSkipsAndDerivations "
    ' First see if any Groups should be disabled
    bSomethingChanged = SetGroupSkips
    
    ' NCJ 27 Nov 02 - Evaluate the whole lot together
    ' NCJ 4 Dec 03 - Set context before doing this (it's not now in mcolAREZZOExprs)
    Call SetContext
    Set colResults = moArezzo.EvaluateExpressions(mcolAREZZOExprs)
    
    ' Now go through the responses
    ' Assume the collection is in sync. with the skips & derivs results collection
    i = 1
    For Each oResponse In moResponses
        If oResponse.Element.DerivationExpr > "" Or oResponse.Element.CollectIfCond > "" Then
            ' Result i is setting of context which we ignore
            sDerivResult = colResults(i + 1)
            bCollectIfResult = (CStr(colResults(i + 2)) = "true")
            i = i + 3
        Else
            ' sDerivResult will actually be ignored by SetSkipDerivation
            sDerivResult = ""
            bCollectIfResult = True
        End If
        If SetSkipDerivation(oResponse, sDerivResult, bCollectIfResult, _
                                enContext, sUserName, oChangedResponse) Then
            bSomethingChanged = True
        End If
    Next
    
    SetAllSkipsAndDerivations = bSomethingChanged
    Set oResponse = Nothing
    
'Debug.Print "  " & Timer & " END SetAllSkipsAndDerivations "

End Function

'------------------------------------------------
Private Function SetGroupSkips() As Boolean
'------------------------------------------------
' NCJ 2 Jan 02
' Assess the status of all Group instances that have Collect If conditions
' Returns TRUE if anything changed
' NCj 27 Nov 02 - Use previously created mcolGroupSkipExprs collection
' NCJ 8 Jul 04 - Take care of removed skips too
'------------------------------------------------
Dim oQGI As QGroupInstance
Dim colResults As Collection
Dim i As Long

    On Error GoTo ErrLabel
    
    SetGroupSkips = False
    
    ' NCJ 8 Jul 04 - Need to consider all groups anyway
    ' in case some skips have been removed
'    If mcolGroupSkipExprs.Count = 0 Then Exit Function   ' Nothing to do if no skips
    
    ' Set MACRO eForm context for AREZZO
    Call SetContext
    
    ' Evaluate all group skips together
    If mcolGroupSkipExprs.Count > 0 Then
        Set colResults = moArezzo.EvaluateExpressions(mcolGroupSkipExprs)
    End If
    ' Loop through colResults
    i = 1
    For Each oQGI In mcolQGroupInstances
        If oQGI.QGroup.Element.CollectIfCond > "" Then
            If colResults(i) = "true" Then
                ' We make sure the group is enabled
                If oQGI.Status = eStatus.NotApplicable Then
                    ' It's not NotApplicable any more
                    Call oQGI.SetNotApplicableStatus(False)
                    SetGroupSkips = True
                End If
            Else
                ' We set the group status to NotApplicable
                If oQGI.Status <> eStatus.NotApplicable Then
                    Call oQGI.SetNotApplicableStatus(True)
                    SetGroupSkips = True
                End If
            End If
            ' Move to next result
            i = i + 1
        Else
            ' NCJ 8 Jul 04 - No collect if - but did it have one before?
            If oQGI.Status = eStatus.NotApplicable Then
                ' It's not NotApplicable any more
                Call oQGI.SetNotApplicableStatus(False)
                SetGroupSkips = True
            End If
        End If
    Next
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.SetGroupSkips"
    
End Function

'------------------------------------------------
Private Function CheckDerivation(oResponse As Response, _
                        oElement As eFormElementRO, _
                        ByVal sDerivExpr As String, _
                        ByVal sDerivResult As String, _
                        ByVal enContext As eAutoRFCContext, _
                        ByVal sUserName As String, _
                        oChangedResponse As Response) As Boolean
'------------------------------------------------
' NCJ 31 Aug 01
' Set the derived value (if any) for this Response object
' Return TRUE if something changed
' or FALSE otherwise
' NCJ 12 Jul 02 - CBB 2.2.19/12 Do not touch unobtainable derived fields if new value is empty
' NCJ 13 Aug 02 - Added enContext, sUserName, oChangedResponse for automatic RFCs
' NCJ 31 Oct 02 - Keep track of any warnings and rejects that happen
' NCJ 25 Feb 03 - Bug 861 - Must take account of local date settings for derived dates
' NCJ 4 Dec 06 - Bug 2848 - Integers must be "localised" too (otherwise regional settings problems)
'------------------------------------------------
Dim sMsg As String
Dim bChanged As Boolean
Dim sNewValue As String
Dim sRFC As String
Dim enNewStatus As eStatus

    On Error GoTo ErrLabel
    
    CheckDerivation = False
    If sDerivExpr = "" Then Exit Function
    
    sNewValue = ""
    If moArezzo.ResultOK(sDerivResult) Then
        sNewValue = sDerivResult
    End If
    
    If sNewValue > "" Then
        Select Case oElement.DataType
        Case eDataType.DateTime
            ' Convert from Arezzo date
            ' NCJ 25 Feb 03 - Take account of local date settings using new oResponse.LocalDateFormat
            sNewValue = moArezzo.FormatDate(sNewValue, oResponse.LocalDateFormat)
        Case eDataType.Real, eDataType.LabTest, eDataType.IntegerNumber
        ' NCJ 4 Dec 06 - Bug 2848 - Added IntegerNumber to this case
            ' Localise real numbers
            ' NCJ 13 May 03 - Use element's formatter (which knows about local dec.pts. etc.)
'            sNewValue = StandardNumToLocal(sNewValue)
            ' NCJ 4 Dec 06 - Bug 2848 - Localise integers too
            sNewValue = oElement.LocaliseValue(sNewValue)
        Case Else
            ' Leave as it is
        End Select
    End If
    
    ' NCJ 12 Jul 02 - If value is empty and question is Unobtainable, don't do anything
    If sNewValue = "" And oResponse.Status = eStatus.Unobtainable Then
        ' Leave it alone
    Else
        ' NCJ 13 Mar 03 - Let ValidateValue decide if it's changed (deals with category values/codes correctly)
        enNewStatus = oResponse.ValidateValue(sNewValue, sMsg, bChanged)
        If bChanged Then
            ' NCJ 13 Aug 02 - Added automatic RFCs
            sRFC = ""
            ' We keep a note of warnings and invalid data
            ' (NB we might come through here more than once for each response)
            If enNewStatus <> eStatus.InvalidData Then
                ' NCJ 31 Oct 02 - If it's a warning add it to our collection
                If enNewStatus = eStatus.Warning Then
                    Call CollectionAddAnyway(mcolWarnDerivations, oResponse, Str(oResponse.ResponseId))
                    Call CollectionAddAnyway(mcolWarnRejectMessages, sMsg, Str(oResponse.ResponseId))
                Else
                    ' There's no warning or rejection
                    Call CollectionRemoveAnyway(mcolWarnDerivations, Str(oResponse.ResponseId))
                    Call CollectionRemoveAnyway(mcolRejectDerivations, Str(oResponse.ResponseId))
                End If
                If oResponse.RequiresValueRFC() Then
                    sRFC = AutomaticRFC(enContext, NewDerivedValue, oChangedResponse)
                End If
                Call oResponse.ConfirmValue("", sRFC, sUserName)
                CheckDerivation = bChanged
            Else
                ' NCJ 31 Oct 02 - If it's rejected add it to our collection
                Call CollectionAddAnyway(mcolRejectDerivations, oResponse, Str(oResponse.ResponseId))
                Call CollectionAddAnyway(mcolWarnRejectMessages, sMsg, Str(oResponse.ResponseId))
                Call oResponse.RejectValue
                ' NCJ 31 Oct 02 - And here we must delete the value
                Call oResponse.ValidateValue("", sMsg, bChanged)
                If oResponse.RequiresValueRFC() Then
                    sRFC = AutomaticRFC(enContext, ValueRejected, oChangedResponse)
                End If
                Call oResponse.ConfirmValue("", sRFC, sUserName)
                CheckDerivation = True
            End If
        End If
    End If
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.CheckDerivation"
    
End Function
 
'------------------------------------------------
Private Function LoadLabTestData(ByVal sCon As String, ByVal sLabCode As String) As Boolean
'------------------------------------------------
' Load in lab test data for this EFI for given lab code.
' This sets up a LabTestValidate object for each Response (of type LabTest)
' and sets up the LabDescription.
' Returns TRUE if Lab was successfully loaded,
' or FALSE if the lab was not found.
' ASSUME Elements and Responses are loaded
' NCJ 30 Jun 04 - If empty LabCode, remove lab test data
'------------------------------------------------
Dim sTest As String
Dim sScheme As String
Dim oElement As eFormElementRO
Dim oResponse As Response
Dim colTests As Collection
Dim oDS As StudySubjectDS
Dim colValidations As Collection
Dim i As Long

    On Error GoTo ErrLabel
    
    ' NCJ 30 Jun 04 - If empty lab code, remove lab defns
    If sLabCode = "" Then
        Call RemoveLabtestData
        msLabDescription = ""
        ' Treat as successful
        LoadLabTestData = True
        Exit Function
    End If
    
    Set oDS = New StudySubjectDS
    
    ' First see if the Lab definition exists
    msLabDescription = oDS.LoadLabDescription(sCon, sLabCode)
    If msLabDescription = "" Then
        ' No lab definition for this lab code
        Set oDS = Nothing
        LoadLabTestData = False
        Exit Function
    Else
        ' There is a lab definition for this lab code
        LoadLabTestData = True
    End If
    
    sScheme = moStudy.CTCScheme
    
    ' Create a collection of tests that appear on this eForm
    Set colTests = New Collection
    For Each oElement In Me.eForm.EFormElements
        If oElement.DataType = eDataType.LabTest Then
            'use a collection so duplicate tests aren't added
            CollectionAddAnyway colTests, oElement.ClinicalTest, oElement.ClinicalTest
        End If
    Next
    
    ' Create comma delimited string of tests that appear on this eForm
    If colTests.Count > 0 Then
        sTest = ""
        ' There are some lab test questions on this form
        For i = 1 To colTests.Count
            sTest = sTest & colTests(i) & ","
        Next
        ' Take off last comma
        sTest = Left(sTest, Len(sTest) - 1)
   
        ' Use Data Services to load lab test stuff
        Set colValidations = CollectionDeSerialise(oDS.LoadNRandCTC(sCon, sTest, sLabCode, sScheme), True)

        ' NCJ 4 Mar 02 - We will set the LabTestValidate object
        ' only for the Responses with repeat number of 1
        ' (Those with higher repeat nos. automatically inherit this first one)
        For Each oElement In Me.eForm.EFormElements
            If oElement.DataType = eDataType.LabTest Then
                ' This is a lab test question
                With Me.Responses.ResponseByElement(oElement, 1)
                    Call .LabTestValidate.SetState(colValidations(oElement.ClinicalTest))
                End With
            End If
        Next
    End If
    
    Set oDS = Nothing
    
    Set colTests = Nothing
    Set colValidations = Nothing
    Set oElement = Nothing


Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.LoadLabTestData"
    
End Function

'--------------------------------------------------------------------
Private Sub RemoveLabtestData()
'--------------------------------------------------------------------
' NCJ 30 Jun 04 - Remove the lab from this eForm
'--------------------------------------------------------------------
Dim oLabTestState As LabTestValidateState
Dim oElement As eFormElementRO

    On Error GoTo ErrLabel
    
    msLabDescription = ""
    
    Set oLabTestState = New LabTestValidateState
    
    ' We will set an empty LabTestValidateState for each LabTestValidate object
    ' only for the Responses with repeat number of 1
    ' (Those with higher repeat nos. automatically inherit this first one)
    For Each oElement In Me.eForm.EFormElements
        If oElement.DataType = eDataType.LabTest Then
            ' This is a lab test question
            With Me.Responses.ResponseByElement(oElement, 1)
                Call .LabTestValidate.SetState(oLabTestState)
            End With
        End If
    Next
    
    Set oLabTestState = Nothing
    Set oElement = Nothing

Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.RemoveLabTestData"

End Sub

'--------------------------------------------------------------------
Public Function QGroupInstanceById(lQGroupID As Long) As QGroupInstance
'--------------------------------------------------------------------
' Get the group instance corresponding to this QGroupID, if any.
' Returns Nothing if an instance doesn't exist
'--------------------------------------------------------------------

    On Error Resume Next
    Set QGroupInstanceById = mcolQGroupInstances.Item(Str(lQGroupID))

End Function

'--------------------------------------------------------------------
Private Function AddNewQGI(oQG As QGroupRO) As QGroupInstance
'--------------------------------------------------------------------
' Create and add a new Group Instance for this question group
' NCJ 14 Nov 02 - Generate in memory only (without a State object)
'--------------------------------------------------------------------
Dim oQGI As QGroupInstance

    On Error GoTo ErrLabel
    
    Set oQGI = New QGroupInstance
    
    Call oQGI.Init(oQG, Me)
    mcolQGroupInstances.Add oQGI, Str(oQG.QGroupID)
    Set AddNewQGI = oQGI
    
    Set oQGI = Nothing

Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "EFormInstance.AddNewQGI"

End Function

'--------------------------------------------------------------------
Friend Function AutomaticRFC(ByVal enContext As eAutoRFCContext, _
                            ByVal enChange As eAutoRFCChange, _
                            Optional oChangedResponse As Response = Nothing) As String
'--------------------------------------------------------------------
' NCJ 13 Aug 02
' Get the automatic RFC text for a response change
' based on context of change,
' nature of change,
' and perhaps the response being changed
' NB If enContext is "Changing Response" then oChangedResponse MUST be valid
' NCJ 27 Aug 02 - Added LabResultChanged
'--------------------------------------------------------------------
Dim sRFC As String

    ' Auto RFCs always start with three asterisks
    sRFC = "*** "
    
    Select Case enChange
    Case eAutoRFCChange.FieldDisabled
        sRFC = sRFC & "Response deleted and disabled "
        
    Case eAutoRFCChange.FieldEnabled
        sRFC = sRFC & "Response field enabled "
        
    Case eAutoRFCChange.NewDerivedValue
        sRFC = sRFC & "Derived value recalculated "
        
    Case eAutoRFCChange.StatusChanged
        sRFC = sRFC & "Status changed "
    
    Case eAutoRFCChange.ValueRejected
        sRFC = sRFC & "Value rejected "
    
    Case eAutoRFCChange.LabResultChanged
        sRFC = sRFC & "NR Status or CTC Grade changed "
    
    End Select
    
    Select Case enContext
    Case eAutoRFCContext.OpeningEForm
        sRFC = sRFC & "on reopening eForm"
        
    Case eAutoRFCContext.ChangingResponse
        sRFC = sRFC & "after change to question " & oChangedResponse.Element.Code
        ' Include repeat number for group questions
        If Not oChangedResponse.Element.OwnerQGroup Is Nothing Then
            sRFC = sRFC & " [" & oChangedResponse.RepeatNumber & "]"
        End If
        
    Case eAutoRFCContext.Revalidation
        sRFC = sRFC & "as a result of revalidation"
        
    Case eAutoRFCContext.ChangingLab
        sRFC = sRFC & "as a result of changing eForm laboratory"
        
    Case eAutoRFCContext.SettingEFormStatus
        sRFC = sRFC & "as a result of changing eForm status"
        
    Case eAutoRFCContext.AddingGroupRow
        sRFC = sRFC & "as a result of adding a group row"
    
    Case eAutoRFCContext.ReloadingSubjectData
        sRFC = sRFC & "after reloading subject data"
    
    End Select
    
    AutomaticRFC = sRFC

End Function

'----------------------------------------------------
Public Property Get SDVStatus() As eSDVStatus
'----------------------------------------------------
' TA 14 Aug 02
' The current "SDV status" of the eForm
'----------------------------------------------------

    SDVStatus = moEFIState.SDVStatus

End Property

'----------------------------------------------------
Public Property Let SDVStatus(ByVal enSDVStatus As eSDVStatus)
'----------------------------------------------------
' TA 14 Aug 02
' The current "SDV status" of the eForm
' NB This doesn't update the "Changed" flag
'----------------------------------------------------

    moEFIState.SDVStatus = enSDVStatus

End Property

'----------------------------------------------------
Public Property Get NoteStatus() As eNoteStatus
'----------------------------------------------------
' TA 14 Aug 02
' The current "Note status" of the eForm
'----------------------------------------------------

    NoteStatus = moEFIState.NoteStatus

End Property

'----------------------------------------------------
Public Property Let NoteStatus(ByVal enNoteStatus As eNoteStatus)
'----------------------------------------------------
' TA 14 Aug 02
' The current "Note status" of the eForm
' NB This doesn't update the "Changed" flag
'----------------------------------------------------

    moEFIState.NoteStatus = enNoteStatus

End Property

'----------------------------------------------------
Public Property Get DiscrepancyStatus() As eDiscrepancyStatus
'----------------------------------------------------
' TA 14 Aug 02
' The current "Discrepancy status" of the eForm
'----------------------------------------------------

    DiscrepancyStatus = moEFIState.DiscrepancyStatus

End Property

'----------------------------------------------------
Public Property Let DiscrepancyStatus(ByVal enDiscrepancyStatus As eDiscrepancyStatus)
'----------------------------------------------------
' TA 14 Aug 02
' The current "Discrepancy status" of the eForm
' NB This doesn't update the "Changed" flag
'----------------------------------------------------

    moEFIState.DiscrepancyStatus = enDiscrepancyStatus

End Property

'----------------------------------------------------
Public Function GetHotlinkTarget(sHotlink As String) As Long
'----------------------------------------------------
' Return the EFormTaskId corresponding to the given Hotlink string
' Returns 0 if Hotlink cannot be calculated
' NCJ 22 Jun 04 - Don't allow a visit eForm! (Bug 2300)
'----------------------------------------------------
Dim lEFITaskId As Long
Dim oVEFI As EFormInstance
Dim oEFI As EFormInstance

    ' NCJ 18 Dec 06 - Beware of AREZZO not being instantiated
    If Not moArezzo Is Nothing Then
        lEFITaskId = moArezzo.GetHotlinkTarget(sHotlink, moEFIState.eFormTaskId)
    Else
        lEFITaskId = 0
    End If
    If lEFITaskId > 0 Then
        ' NCJ 22 Jun 04 - Check for Visit eForm
        Set oEFI = moStudy.Subject.eFIByTaskId(lEFITaskId)
        Set oVEFI = oEFI.VisitInstance.VisitEFormInstance
        If Not oVEFI Is Nothing Then
            ' Is it the same as this one?
            If oVEFI.eFormTaskId = lEFITaskId Then
                ' Disallow it
                lEFITaskId = 0
            End If
            Set oVEFI = Nothing
        End If
        Set oEFI = Nothing
    End If
    
    GetHotlinkTarget = lEFITaskId

End Function

'----------------------------------------------------
Private Sub InitAREZZOExprCollection()
'----------------------------------------------------
' Generate the collection of derivations and collect if conditions for existing responses
'----------------------------------------------------
Dim oResponse As Response

    Set mcolAREZZOExprs = New Collection
    ' NCJ 18 Dec 06 - Only do this for non-RO subjects
    If Not Me.ReadOnly Then
        For Each oResponse In Responses
            Call AddResponseAREZZOExprs(oResponse)
        Next
        Set oResponse = Nothing
    End If
    
End Sub

'----------------------------------------------------
Private Sub AddResponseAREZZOExprs(oResponse As Response)
'----------------------------------------------------
' Add the derivation/collect if to our collection
' Only adds an entry if there is a derivation or a collect if
'----------------------------------------------------

    With oResponse.Element
        If .DerivationExpr > "" Or .CollectIfCond > "" Then
            ' Add the Prolog goal to set the context
            ' NCJ 4 Dec 03 - Pass Nothing as the EFI (we now do the eForm context elsewhere)
            mcolAREZZOExprs.Add moArezzo.SetQuestionContextString(.Code, oResponse.RepeatNumber, Nothing)
            ' Then add the Derivation and the Collect If
            mcolAREZZOExprs.Add .DerivationExpr
            mcolAREZZOExprs.Add .CollectIfCond
        End If
    End With

End Sub

'----------------------------------------------------
Public Function GetAllSubjectsKey() As String
'----------------------------------------------------
'return a key which will unqiely ideintfy this form for any subject
'ie visitcode|visitcycle|eformcode|eformcycle
'----------------------------------------------------

    With Me
        GetAllSubjectsKey = .VisitInstance.Code & "|" & .VisitInstance.CycleNo _
                                    & "|" & .Code & "|" & .CycleNo
       
    End With
    
End Function

'----------------------------------------------------
Friend Property Get EFIState() As eFormInstanceState
'----------------------------------------------------
' NCJ 27 Jan 03
' This is used when tidying up after a reload
'----------------------------------------------------

    Set EFIState = moEFIState

End Property
