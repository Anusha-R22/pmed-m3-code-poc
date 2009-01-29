VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Revalidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         Revalidator.cls
' Copyright:    InferMed Ltd. 2003-2007. All Rights Reserved
' Author:       Nicky Johns, February 2003
' Purpose:      Class to handle MACRO 3.0 Batch Validation
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 21-28 Feb 03 - Initial development
'   NCJ 3 Mar 03 - Re-use StudyDef object if possible
'   NCJ 7 Mar 03 - Check each object revalidatable before revalidating
'   NCJ 25 Mar 03 - Include refreshing of subject label etc.
'   NCJ 7 May 03 - Don't log changes to responses that won't be saved (in last blank rows of RQGs)
'   NCJ 20 May 03 - Bugs 1763-6, extra changes to be logged
'   NCJ 2 Mar 04 - Must check for changes on eForm AND visit eForm in EFormRevalidate
'   NCJ 8 Jul 04 - Consider removed Collect Ifs in WhatsChanged
'   NCJ 10 Nov 04 - Issue 2443 - New property Verbose; do not output non-changes if Verbose = false
'   NCJ 29 Nov 04 - Issue 2468 - Include NRStatus/CTCGrade in log file for lab questions
'   NCJ 28 Feb 07 - Issues 2102/2871 - All Studies and Single EForm
'   NCJ 31 Aug 07 - Issue 2931 - Allow "report only" mode, i.e. do not save any changes
'   NCJ 31 Aug 07 - Issue 2944 - Ensure log file is complete for "changes only" on unchanged subjects
'   NCJ 3 Sep 07 - Issue 2211 - Don't report "No changes" if subj. label was changed
'   NCJ 27 Sept 07 - Bug 2935 - Don't pass user name to RefreshSkips
'----------------------------------------------------------------------------------------'

Option Explicit

' The name of the log file
Private msLogFile As String

Private moSubject As StudySubject
Private moStudyDef As StudyDefRO
Private moUser As MACROUser

Private mlStudyId As Long
Private msSite As String
Private mlSubjectId As Long

Private msLockToken As String

Private msSubjSpec As String

Private mcolChangedEForms As Collection

' NCJ 10 Nov 04 - Collection of lines for log file
Private mcolLogLines As Collection

' To what does the log message refer?
Private Const MSG_TOP_LEVEL = 0
Private Const MSG_SUBJECT = 1
Private Const MSG_VISIT = 2
Private Const MSG_EFORM = 3
Private Const MSG_QUESTION = 4

' Do they want "verbose" output?
Public Verbose As Boolean
' NCJ 31 Aug 07 - 2931 - Do they want to save changes or just see what happens?
Private mbSaveChanges As Boolean

'---------------------------------------------------------------------
Public Sub Revalidate(oUser As MACROUser, _
                    ByVal lStudyId As Long, ByVal sSite As String, ByVal lSubjectId As Long, _
                    ByVal sStudyName As String, ByVal sSubjLabel As String, oArezzo As Arezzo_DM, _
                    ByVal lEFormId As Long)
'---------------------------------------------------------------------
' Revalidate the given subject
' NCJ 6 May 03 - Added UserName and UserRole to LoadSubject
' NCJ 28 Feb 07 - Added lEFormID to allow revalidation of single eForm; lEFormId = 0 means do all forms
'---------------------------------------------------------------------
Dim sErrMsg As String
Dim sSubj As String
Dim bChanged As Boolean

    On Error GoTo ErrLabel
    
    ' Store the info
    Set moUser = oUser
    
    mlStudyId = lStudyId
    msSite = sSite
    mlSubjectId = lSubjectId
    
    ' Start off empty
    Set moSubject = Nothing
    
    ' Initialise changed eforms
    Set mcolChangedEForms = New Collection
    
    ' Initialise log entries for this subject
    Set mcolLogLines = New Collection
    
    ' Use subject label or ID if not available
    If sSubjLabel > "" Then
        sSubj = sSubjLabel
    Else
        sSubj = lSubjectId
    End If
    
    msSubjSpec = sStudyName & "/" & sSite & "/" & sSubj
   
    Call LogToFile("Subject: " & msSubjSpec, MSG_SUBJECT)
    Call LogToFile("Revalidation started at " & CurTime, MSG_SUBJECT)

    ' NCJ 28 Feb 07 - Added oUser to LoadStudyDef
    If LoadStudyDef(lStudyId, oUser, oArezzo) Then
        Call moStudyDef.LoadSubject(sSite, lSubjectId, moUser.UserName, eUIUpdateMode.Read_Write, _
                                    moUser.UserNameFull, moUser.UserRole)
        If moStudyDef.Subject.CouldNotLoad Then
            ' Give up
            Call LogToFile("Unable to open subject: " & moStudyDef.Subject.CouldNotLoadReason, MSG_SUBJECT)
        Else
            ' Successfully loaded
            Set moSubject = moStudyDef.Subject
        End If
    End If

    ' See if we got through all that OK
    If Not moSubject Is Nothing Then
        ' Check subject isn't locked etc.
        If SubjectRevalidatable Then
            ' Proceed with revalidation!
            ' NCJ 28 Feb 07 - Pass in lEFormId (may be 0)
            bChanged = SubjectRevalidate(lEFormId)
            ' Clear up the subject
            Call moStudyDef.RemoveSubject
            Set moSubject = Nothing
            ' Do the Changes Summary
            Call ReportChanges(bChanged)
        End If
    End If
    
    Call LogToFile("Revalidation completed at " & CurTime & vbCrLf, MSG_SUBJECT)
    
    ' Leave the StudyDef in case we want to re-use it for the next subject
'    ' Clear up the study def
'    Call moStudyDef.Terminate
'    Set moStudyDef = Nothing

    ' Finally flush the log messages to file if necessary
    ' NCJ 31 Aug 07 - Bug 2944 - Do this all the time (not just if bChanged)
'    If bChanged Then
        FlushLogLines
'    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Revalidator.Revalidate"

End Sub

'---------------------------------------------------------------------
Public Property Get SaveChanges() As Boolean
'---------------------------------------------------------------------
' Do they want changes to be automatically saved?
'---------------------------------------------------------------------

    SaveChanges = mbSaveChanges
    
End Property

'---------------------------------------------------------------------
Public Property Let SaveChanges(bSave As Boolean)
'---------------------------------------------------------------------

    mbSaveChanges = bSave
    
End Property

'---------------------------------------------------------------------
Private Function LoadStudyDef(ByVal lStudyId As Long, oUser As MACROUser, oArezzo As Arezzo_DM) As Boolean
'---------------------------------------------------------------------
' Load the study def into moStudyDef
' Only reload if it's different from what we've already got
' Returns TRUE if study successfully loaded
' NCJ 28 Feb 07 - Added oUser
'---------------------------------------------------------------------
Dim sErrMsg As String
Dim bNeedToLoad As Boolean
Dim bLoadedOK As Boolean

    On Error GoTo ErrLabel
    
    bLoadedOK = True
    bNeedToLoad = True
    
    If Not moStudyDef Is Nothing Then
        ' Load if it's a different study
        bNeedToLoad = (moStudyDef.StudyId <> lStudyId)
        If bNeedToLoad Then
            ' Tidy up the old one
            Call moStudyDef.Terminate
            Set moStudyDef = Nothing
        End If
    End If
    
    If bNeedToLoad Then
        Set moStudyDef = New StudyDefRO
        sErrMsg = moStudyDef.Load(oUser.CurrentDBConString, lStudyId, 1, oArezzo)
        If sErrMsg > "" Then
            ' Give up
            Call LogToFile("Unable to load study: " & sErrMsg, MSG_SUBJECT)
            bLoadedOK = False
        End If
    End If
    
    LoadStudyDef = bLoadedOK

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Revalidator.LoadStudyDef"

End Function

'---------------------------------------------------------------------
Public Function StudyEForms(lStudyId As Long, oUser As MACROUser, oArezzo As Arezzo_DM) As Collection
'---------------------------------------------------------------------
' NCJ 28 Feb 07 - Issue 2871 - Allow selection of a single eform
' Return eForms for a given study
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    If LoadStudyDef(lStudyId, oUser, oArezzo) Then
        Set StudyEForms = moStudyDef.eForms
    End If
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Revalidator.StudyEForms"

End Function

'---------------------------------------------------------------------
Public Sub InitRevalidation(ByVal sFileName As String, oUser As MACROUser, seForm As String)
'---------------------------------------------------------------------
' Initialise the revalidation log file
' NCJ 10 Nov 04 - Write header info directly to file here, including "verbosity" mode
' NCJ 6 Mar 07 - Added sEForm for single eForm (may be "")
' NCJ 31 Aug 07 - Issue 2931 - Added "SaveChanges" mode
'---------------------------------------------------------------------
Dim n As Integer
Dim sText As String

    ' Store the log file name
    msLogFile = sFileName
    
    ' Create initial text
    sText = "MACRO 3.0 Subject Revalidation Report "
    sText = sText & CurTime & vbCrLf
    sText = sText & "User: " & oUser.UserNameFull & vbCrLf
    sText = sText & "Database: " & oUser.DatabaseCode & vbCrLf
    sText = sText & "Report mode: "
    If Me.Verbose Then
        sText = sText & "Show all"
    Else
        sText = sText & "Show changes only"
    End If
    sText = sText & vbCrLf & "Save mode: "
    If mbSaveChanges Then
        sText = sText & "Save all changes"
    Else
        sText = sText & "Do not save any changes"
    End If
    If seForm <> "" Then
        sText = sText & vbCrLf & "Single eForm: " & seForm
    End If
    
    ' Create a new log file
    n = FreeFile
    Open msLogFile For Output As n
    Print #n, sText & vbCrLf
    Close n
        
'    Call LogToFile(sText)
    
    Set moStudyDef = Nothing

End Sub

'---------------------------------------------------------------------
Public Sub EndRevalidation()
'---------------------------------------------------------------------
' Write out final line to log file and
' clear up our objects at the end
'---------------------------------------------------------------------
Dim n As Integer

    ' NCJ 11 Nov 04 - Say what we did
    If msLogFile > "" Then
        n = FreeFile
        Open msLogFile For Append As n
        Print #n, vbCrLf & "End of revalidation " & CurTime
        Close n
    End If
    
    Call Terminate
    
End Sub

'---------------------------------------------------------------------
Private Sub Terminate()
'---------------------------------------------------------------------
' Tidy up all our objects
'---------------------------------------------------------------------
    
    Set moUser = Nothing
    Set moSubject = Nothing
    If Not moStudyDef Is Nothing Then
        Call moStudyDef.RemoveSubject
        ' Clear up the study def
        Call moStudyDef.Terminate
        Set moStudyDef = Nothing
    End If

End Sub

'---------------------------------------------------------------------
Private Function CurTime() As String
'---------------------------------------------------------------------
' Current timestamp as a string
'---------------------------------------------------------------------

    CurTime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    
End Function

'---------------------------------------------------------------------
Private Sub LogToFile(ByVal sText As String, _
                    Optional nMsgType As Integer)
'---------------------------------------------------------------------
' Add text to the revalidation log file (assume initialised)
' Insert no. of tabs according to msg type (Subj, Visit, eForm etc. - see MSG_xxx constants)
'---------------------------------------------------------------------
Dim n As Integer
Dim sLog As String

    ' Add on appropriate no. of spaces
    sLog = Space(nMsgType * 2) & sText
    
    If Me.Verbose Then
        ' Ordinary logging
        If msLogFile = "" Then
            Debug.Print "No log file: " & sText
        Else
            n = FreeFile
            Open msLogFile For Append As n
            Print #n, sLog
            Close n
        End If
    Else
        ' Store for later
        mcolLogLines.Add sLog
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub Class_Initialize()
'---------------------------------------------------------------------

    msLogFile = ""
    Set moStudyDef = Nothing
    ' Default to "verbose" output
    Me.Verbose = True
    ' Default to saving changes
    mbSaveChanges = True
    Set mcolLogLines = New Collection

End Sub

'---------------------------------------------------------------------
Private Sub Class_Terminate()
'---------------------------------------------------------------------

    Call Terminate

End Sub

'---------------------------------------------------------------------
Private Function GetEFormVisits(ByVal lEFormId As Long) As Collection
'---------------------------------------------------------------------
' Find all the visits in which this eForm occurs in the current subject
'---------------------------------------------------------------------


End Function

'---------------------------------------------------------------------
Private Function SubjectRevalidate(ByVal lEFormId As Long) As Boolean
'---------------------------------------------------------------------
' Revalidate all eForms in the current subject
' NCJ 28 Feb 07 - If lEFormId > 0, revalidate that form only
'---------------------------------------------------------------------
Dim oEFI As EFormInstance
Dim oSchedVisit As ScheduleVisit
Dim oVEFI As VEFInstance
Dim oVI As VisitInstance
Dim lVEFTaskId As Long
Dim sEFISpec As String
Dim sErrMsg As String
Dim sTok1 As String
Dim sTok2 As String
Dim bNeedToSave As Boolean
Dim bVisitChanged As Boolean
Dim bFormChanged As Boolean
Dim bSubjChanged As Boolean
Dim nFLogLine As Integer
Dim nVLogLine As Integer
Dim oEForm As eFormRO
Dim neSaveResponsesResult As eSaveResponsesResult

    On Error GoTo ErrLabel
      
    ' Have any of the subject properties changed?
    ' (We'll save at the end if no eForms have been saved)
    bNeedToSave = RefreshSubjectProperties
    bSubjChanged = bNeedToSave

    If lEFormId > 0 Then
        ' Find the eForm we're looking for
        Set oEForm = moStudyDef.eFormById(lEFormId)
    End If
    
    ' Work through the schedule visits
    For Each oSchedVisit In moSubject.ScheduleVisits
    
        ' Remember our place in the log file for this visit
        nVLogLine = mcolLogLines.Count
        bVisitChanged = False
        
        ' Check it's an active visit
        Set oVI = oSchedVisit.VisitInstance
        ' NCJ 6 Mar 07 - See if it's got our eForm (if any)
        If VisitRevalidatable(oVI, oEForm) Then
            ' Find out if it's got a Visit eForm
            lVEFTaskId = 0
            If Not oVI.VisitEFormInstance Is Nothing Then
                lVEFTaskId = oVI.VisitEFormInstance.EFormTaskId
            End If
            ' Work through the visit eform instances
            For Each oVEFI In oSchedVisit.VisitEFormInstances
            
                ' Remember our place in the log file for this eForm
                nFLogLine = mcolLogLines.Count
                bFormChanged = False
                
                Set oEFI = oVEFI.EFormInstance
                If EFIRevalidatable(oEFI, lVEFTaskId, lEFormId, sEFISpec) Then
                    If moSubject.LoadResponses(oEFI, sErrMsg, sTok1, sTok2) = lrrReadWrite Then
                        ' Revalidate eForm
                        If EFormRevalidate(oEFI) Then
                            ' Something changed so we must save the responses if we can
                            If OKToSaveEForm(oEFI) Then
                                ' Report changes to eForm label
                                Call UpdateEFormLabel(oEFI)
                                ' NCJ 31 Aug 07 - Only save if mbSaveChanges is true
                                If mbSaveChanges Then
                                    Call moSubject.SaveResponses(oEFI, sErrMsg)
                                    ' Add to our collection of changed eForms
                                    Call LogToFile("Saved changes to eForm " & sEFISpec, MSG_QUESTION)
'                                Else
'                                    Call LogToFile("Changes not saved to eForm " & sEFISpec, MSG_QUESTION)
                                End If
                                mcolChangedEForms.Add EFormSpec(oEFI)
                                bFormChanged = True
                                ' We've done a subject save now
                                bNeedToSave = False
                            End If
                        Else
                            ' No responses changed but check the eForm label
                            If UpdateEFormLabel(oEFI) Then
                                ' Save the subject
                                ' NCJ 31 Aug 07 - Only save if mbSaveChanges is true
                                If mbSaveChanges Then
                                    neSaveResponsesResult = moSubject.Save
                                Else
                                    neSaveResponsesResult = srrSuccess
                                End If
                                If neSaveResponsesResult <> eSaveResponsesResult.srrSuccess Then
                                    Call LogToFile("Unable to save changed eForm label", MSG_EFORM)
                                Else
                                    ' Say that we saved successfully
                                    If mbSaveChanges Then
                                        Call LogToFile("Saved changes to eForm " & sEFISpec, MSG_QUESTION)
'                                    Else
'                                        Call LogToFile("Changes not saved to eForm " & sEFISpec, MSG_QUESTION)
                                    End If
                                    mcolChangedEForms.Add EFormSpec(oEFI)
                                    bFormChanged = True
                                    ' We've done a subject save now
                                    bNeedToSave = False
                                End If
                            End If
                        End If
                        ' Remove responses
                        Call moSubject.RemoveResponses(oEFI, True)
                    Else
                        ' Remove responses (in case loaded but not read-write)
                        Call moSubject.RemoveResponses(oEFI, True)  ' NB Does no harm if not loaded
                        ' Failed to load responses read-write
                        Call LogToFile("Unable to revalidate " & sEFISpec & ": " & sErrMsg, MSG_QUESTION)
                    End If
                End If
                ' Scrap the log messages for an unchanged eForm
                If bFormChanged Then
                    bVisitChanged = True
                Else
                    Call RemoveUnwantedLogs(nFLogLine)
                End If

            Next    'oVEFI
        End If
        ' Scrap the log entries for visit if nothing changed
        If bVisitChanged Then
            bSubjChanged = True
        Else
            Call RemoveUnwantedLogs(nVLogLine)
        End If
        
    Next    ' SchedVisit

    ' If the subject properties changed but we didn't save any eForms, need to save now
    ' NCJ 31 Aug 07 - 2931 - Also check whether we're saving
    If bNeedToSave And mbSaveChanges Then
        If moSubject.Save <> eSaveResponsesResult.srrSuccess Then
            Call LogToFile("Unable to save changed subject details", MSG_SUBJECT)
        Else
            Call LogToFile("Saved changed subject details", MSG_SUBJECT)
        End If
    End If
    
    Set oSchedVisit = Nothing
    Set oVI = Nothing
    Set oEFI = Nothing
    Set oVEFI = Nothing

    SubjectRevalidate = bSubjChanged

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Revalidator.SubjectRevalidate"

End Function

'---------------------------------------------------------------------
Private Function SubjectRevalidatable() As Boolean
'---------------------------------------------------------------------
' See if we can revalidate the current subject
' Assume moSubject not Nothing and not read-only
' Log suitable messages to log file if not
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    SubjectRevalidatable = False

    If moSubject.LockStatus <> eLockStatus.lsUnlocked Then
        ' Visit is locked/frozen
        Call LogToFile("Subject is " & GetLockStatusString(moSubject.LockStatus), MSG_SUBJECT)
        Exit Function
    End If
    
    ' See if subject is read-only for some other reason
    If moSubject.ReadOnly Then
        Call LogToFile("Unable to change data for subject: " & moStudyDef.Subject.ReadOnlyReason, MSG_SUBJECT)
        Exit Function
    End If
    
    Select Case moSubject.Status
    Case eStatus.Requested
        ' Ignore subject but log it
        Call LogToFile("No data has been entered for this subject", MSG_SUBJECT)
        Exit Function
    Case eStatus.Unobtainable
        ' Ignore subject but log it
        Call LogToFile("Subject is " & GetStatusString(moSubject.Status), MSG_SUBJECT)
        Exit Function
    Case Else
        ' We'll do this subject
    End Select
    
    SubjectRevalidatable = True

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Revalidator.SubjectRevalidatable"

End Function

'---------------------------------------------------------------------
Private Function VisitRevalidatable(oVI As VisitInstance, oEForm As eFormRO) As Boolean
'---------------------------------------------------------------------
' See if we can revalidate this visit
' Log suitable messages to log file if not
' NCJ 6 Mar 07 - Issue 2871 - Look for target eForm too (may be Nothing)
'---------------------------------------------------------------------
Dim sVISpec As String

    On Error GoTo ErrLabel
    
    VisitRevalidatable = False
    
    ' Check that it's an active visit
    If oVI Is Nothing Then Exit Function
    
    If Not oEForm Is Nothing Then
        ' We want a single eForm only - check it's in the visit
        If Not oVI.IsAnInstance(oEForm) Then Exit Function
    End If
    
    ' Set up visit description string for future messages
    sVISpec = AddCycleNo(oVI.Code, oVI.CycleNo)
    
    If oVI.LockStatus <> eLockStatus.lsUnlocked Then
        ' Visit is locked/frozen
        Call LogToFile("Visit " & sVISpec & " is " & GetLockStatusString(oVI.LockStatus), MSG_VISIT)
        Exit Function
    End If
    
    Select Case oVI.Status
    Case eStatus.Requested
        ' Ignore visit
        Exit Function
    Case eStatus.Unobtainable
        ' Ignore visit but log it
        Call LogToFile("Visit " & sVISpec & " is " & oVI.StatusString, MSG_VISIT)
        Exit Function
    Case Else
        ' We'll do this visit
        Call LogToFile("Visit " & sVISpec, MSG_VISIT)
    End Select

    ' If we get to here, it's OK!
    VisitRevalidatable = True

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Revalidator.VisitRevalidatable"

End Function

'---------------------------------------------------------------------
Private Function EFIRevalidatable(oEFI As EFormInstance, ByVal lVEFTaskId As Long, _
            ByVal lEFormId As Long, ByRef sEFISpec As String) As Boolean
'---------------------------------------------------------------------
' See if we can revalidate the eForm
' lVEFTaskId is the task id of the visit eForm (which we want to ignore)
' Log suitable messages to log file if not
' Returns sEFISpec as the description of this eForm for the log file
' NCJ 28 Feb 07 - Added lEFormId to match on (for single form revalidation)
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    EFIRevalidatable = False
    
    ' Check that it's an active eForm
    If oEFI Is Nothing Then Exit Function
    
    ' NCJ 28 Feb 07 - Check to see if we're looking for a specific form,
    ' i.e. if lEFormId > 0 and it's a match for this form
    If lEFormId > 0 Then
        ' We want a specific eForm
        If lEFormId <> oEFI.eForm.EFormId Then
            ' But not this one!
            Exit Function
        End If
    End If
    
    ' Ignore it if it's the Visit eForm
    If oEFI.EFormTaskId = lVEFTaskId Then Exit Function
    
    ' Set up EFI description string for future messages
    sEFISpec = AddCycleNo(oEFI.Code, oEFI.CycleNo)
    
    If oEFI.LockStatus <> eLockStatus.lsUnlocked Then
        ' EForm is locked/frozen
        Call LogToFile("EForm " & sEFISpec & " is " & oEFI.LockStatusString, MSG_EFORM)
        Exit Function
    End If
    
    If oEFI.ReadOnly Then
        ' For some reason EForm is read-only
        Call LogToFile("EForm " & sEFISpec & " is read-only", MSG_EFORM)
        Exit Function
    End If
    
    Select Case oEFI.Status
    Case eStatus.Requested
        ' Ignore eForm
        Exit Function
    Case eStatus.Unobtainable
        ' Ignore eForm but log it
        Call LogToFile("EForm " & sEFISpec & " is " & oEFI.StatusString, MSG_EFORM)
        Exit Function
    Case Else
        Call LogToFile("EForm " & sEFISpec, MSG_EFORM)
    End Select

    ' If we get to here, it's OK!
    EFIRevalidatable = True

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Revalidator.EFIRevalidatable"

End Function

'---------------------------------------------------------------------
Private Function EFormSpec(oEFI As EFormInstance) As String
'---------------------------------------------------------------------
' Get the "spec" of an eForm instance as a string
' (for output to the log file)
'---------------------------------------------------------------------
Dim sSpec As String

    sSpec = oEFI.VisitInstance.Code
    sSpec = AddCycleNo(sSpec, oEFI.VisitInstance.CycleNo)
    sSpec = sSpec & ":" & oEFI.Code
    sSpec = AddCycleNo(sSpec, oEFI.CycleNo)
    
    EFormSpec = sSpec
    
End Function

'---------------------------------------------------------------------
Private Function AddCycleNo(ByVal sStr As String, ByVal nCycleNo As Integer, _
                        Optional ByVal bAddAnyway As Boolean = False)
'---------------------------------------------------------------------
' Add cycle no. to text string if nCycleNo > 1 or if bAddAnyway = TRUE
'---------------------------------------------------------------------

    If (nCycleNo > 1) Or bAddAnyway Then
        AddCycleNo = sStr & "[" & nCycleNo & "]"
    Else
        AddCycleNo = sStr
    End If

End Function

'---------------------------------------------------------------------
Private Function EFormRevalidate(oEFI As EFormInstance) As Boolean
'---------------------------------------------------------------------
' Revalidate all nonempty fields on the eForm (checking non-labtest questions only)
' Returns TRUE if something changed,
' or FALSE if nothing happened.
' Revalidate the Visit EForm, if there is one.
' NCJ 29 Nov 04 - Added bLabQuestionChanged to RevalidateEFI
' NCJ 27 Sept 07 - Bug 2935 - Don't pass user name to RefreshSkips
'---------------------------------------------------------------------
Dim bSomethingChanged As Boolean
Dim oRejectedResponse As Response
Dim oVisitEFI As EFormInstance
Dim bLabQuestionChanged As Boolean

    On Error GoTo ErrLabel
    
    Set oRejectedResponse = Nothing
    bLabQuestionChanged = False
    
    ' First of all check for any newly created responses
    bSomethingChanged = AnyNewResponses(oEFI)
    
    ' Refresh skips and validations
    ' NCJ 27 Sept 07 - Bug 2935 - Don't pass in user name! (It's only for authorisation qus.)
'    bSomethingChanged = bSomethingChanged Or oEFI.RefreshSkipsAndDerivations(eAutoRFCContext.Revalidation, moUser.UserName)
    bSomethingChanged = bSomethingChanged Or oEFI.RefreshSkipsAndDerivations(eAutoRFCContext.Revalidation, "")

    'Revalidate the Visit EForm, if there is one.
    Set oVisitEFI = oEFI.VisitInstance.VisitEFormInstance
    If Not oVisitEFI Is Nothing Then
        bSomethingChanged = bSomethingChanged Or RevalidateEFI(oVisitEFI, oRejectedResponse, bLabQuestionChanged)
    End If
    
    ' Revalidate the main eform
    bSomethingChanged = bSomethingChanged Or RevalidateEFI(oEFI, oRejectedResponse, bLabQuestionChanged)

    ' NCJ 29 Nov 04 - Check for changed lab questions too (in case of derived NR/CTC values)
    If bLabQuestionChanged Or (Not oRejectedResponse Is Nothing) Then
        ' Something was rejected, or a lab question changed status, so we re-run skips and derivations
        ' Set the context for auto RFC generation
        ' NCJ 27 Sept 07 - Bug 2935 - Don't pass in user name here! (It's only for authorisation qus.)
'        bSomethingChanged = bSomethingChanged _
'            Or oEFI.RefreshSkipsAndDerivations(eAutoRFCContext.Revalidation, moUser.UserName, oRejectedResponse)
        bSomethingChanged = bSomethingChanged _
            Or oEFI.RefreshSkipsAndDerivations(eAutoRFCContext.Revalidation, "", oRejectedResponse)
        Set oRejectedResponse = Nothing
    End If
    
    If bSomethingChanged Then
        ' Report the skips and derivs.
        Call WhatsChanged(oEFI)
    End If
    
    ' Return TRUE if something changed in a to-be-saved response
'    EFormRevalidate = bSomethingChanged
    ' NCJ 2 Mar 04 - Must check the main eForm AND the visit eForm
    bSomethingChanged = EFIHasChanged(oEFI)
    If (Not bSomethingChanged) And (Not oVisitEFI Is Nothing) Then
        ' Nothing changed on the main eform - check the Visit eForm
        bSomethingChanged = EFIHasChanged(oVisitEFI)
    End If
    
    EFormRevalidate = bSomethingChanged
    
    Set oRejectedResponse = Nothing
    Set oVisitEFI = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Revalidator.EFormRevalidate"

End Function

'---------------------------------------------------------------------
Private Function RevalidateEFI(ByRef oEFI As EFormInstance, _
                                ByRef oRejectedResponse As Response, _
                                ByRef bLabQuestionChanged As Boolean) As Boolean
'---------------------------------------------------------------------
' Returns TRUE if revalidating the EFI made something change.
' oRejectedResponse may also be set.
' NCJ 29 Nov 04 - bLabQuestionChanged will be set to TRUE if a lab question changed status
'---------------------------------------------------------------------
Dim oResponse As Response
Dim bChangedStatus As Boolean
Dim sErrMsg As String
Dim lStatus As Long
Dim bSomethingChanged As Boolean

    On Error GoTo ErrLabel
    
    bSomethingChanged = False
    
    For Each oResponse In oEFI.Responses
        ' Ignore responses which won't be saved
        If oResponse.WillBeSaved Then
            ' Only deal with "Revalidatable" responses
            If oResponse.Revalidatable(False) Then
            
                lStatus = oResponse.RevalidateValue(sErrMsg, bChangedStatus)
                If bChangedStatus Then
                    bSomethingChanged = True
                    ' NCJ 29 Nov 04 - Remember if a lab question changed
                    If oResponse.Element.DataType = eDataType.LabTest Then
                        bLabQuestionChanged = True
                    End If
                    ' Confirm the change (no RFO)
                    ' NCJ 29 Nov 04 - Must do this BEFORE logging validation message (for NR/CTC stuff)
                    Call oResponse.ConfirmRevalidation("")
                    ' Log the validation message (whatever the new status)
                    Call LogValidationMessage(oResponse, lStatus, sErrMsg)
                    If lStatus = eStatus.InvalidData And oRejectedResponse Is Nothing Then
                        ' Remember the first rejected response
                        Set oRejectedResponse = oResponse
                    End If
    
                Else
                    ' No change so ignore the revalidation
                    Call oResponse.RejectValue
                End If      ' if changed
                
            End If      ' if we revalidated
        End If      ' if response will be saved
    Next    ' oResponse
    
    Set oResponse = Nothing

    RevalidateEFI = bSomethingChanged
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Revalidator.RevalidateEFI"

End Function

'---------------------------------------------------------------------
Private Sub LogValidationMessage(oResponse As Response, _
                                ByVal nStatus As Integer, _
                                ByVal sValMsg As String)
'---------------------------------------------------------------------
' Log a revalidation message for a response
'---------------------------------------------------------------------
Dim sText As String

    On Error GoTo ErrLabel
    
    sText = QuestionSpec(oResponse)
    
    Select Case nStatus
    Case eStatus.InvalidData
        sText = sText & " value rejected"
    Case Else
        sText = sText & " new status is " & GetStatusString(nStatus)
        ' NCJ 25 Nov 04 - Check for NRStatus/CTC
        If oResponse.Element.DataType = eDataType.LabTest Then
            If oResponse.NRStatusString > "" Or oResponse.CTCGrade > 0 Then
                sText = sText & " [" & oResponse.NRStatusString
                If oResponse.CTCGrade > 0 Then
                    sText = sText & oResponse.CTCGrade
                End If
                sText = sText & "]"
            End If
        End If
    End Select
    If sValMsg > "" Then
        sText = sText & ": " & sValMsg
    End If
    
    Call LogToFile(sText, MSG_QUESTION)

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Revalidator.LogValidationMessage"
    
End Sub

'---------------------------------------------------------------------
Private Function QuestionSpec(oResponse As Response) As String
'---------------------------------------------------------------------
' Get the description for a question, preceded by "Question "
'---------------------------------------------------------------------
    
    QuestionSpec = "Question " & AddCycleNo(oResponse.Element.Code, oResponse.RepeatNumber, _
                        (Not oResponse.Element.OwnerQGroup Is Nothing))

End Function

'---------------------------------------------------------------------
Private Sub ReportChanges(bSubjChanged As Boolean)
'---------------------------------------------------------------------
' Report the eForm changes that have occurred to this subject
' NCJ 3 Sep 07 - Issue 2211 - Also take into account bSubjChanged (for subject label)
'---------------------------------------------------------------------
Dim vEForm As Variant
Dim nChanged As Integer
Dim sMSG As String
Dim sComma As String

    On Error GoTo ErrLabel
    
    nChanged = mcolChangedEForms.Count
    If nChanged = 0 Then
        If bSubjChanged Then
            ' No eForms changed, but something else did
            Call LogToFile("No eForm changes occurred in this subject", MSG_SUBJECT)
        Else
            ' No changes
            Call LogToFile("No changes occurred in this subject", MSG_SUBJECT)
        End If
    Else
        ' Say how many were changed
        ' NCJ 31 Aug 07 - Issue 2931 - Say "affected" if we're not saving changes
        sMSG = nChanged & " eForm"
        If nChanged > 1 Then
'            sMsg = sMsg & "s were changed :"
            sMSG = sMSG & "s were "
        Else
'            sMsg = sMsg & " was changed :"
            sMSG = sMSG & " was "
        End If
        If mbSaveChanges Then
            sMSG = sMSG & "changed :"
        Else
            sMSG = sMSG & "affected :"
        End If
        Call LogToFile(sMSG, MSG_SUBJECT)
        ' Now give a list of the eForms
        sMSG = "   "
        sComma = ""
        For Each vEForm In mcolChangedEForms
            sMSG = sMSG & sComma & CStr(vEForm)
            sComma = ", "
        Next
        Call LogToFile(sMSG, MSG_SUBJECT)
    End If
    ' NCJ 31 Aug 07 - Issue 2931 - Remind them if something changed but they chose not to save
    If (nChanged > 0 Or bSubjChanged) And Not mbSaveChanges Then
        Call LogToFile("No changes were saved for this subject", MSG_SUBJECT)
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Revalidator.ReportChanges"
    
End Sub

'---------------------------------------------------------------------
Private Function OKToSaveEForm(oEFI As EFormInstance) As Boolean
'---------------------------------------------------------------------
' Is it OK to save this eForm?
' Returns TRUE if eForm/Visit dates OK, or False otherwise
' Logs appropriate messages to the Log File
' (Based on EFormAndVisitDatesBlank function in Data Entry)
'---------------------------------------------------------------------
Dim oVisitEFI As EFormInstance
Dim oResponse As Response
Dim bBlank As Boolean
Dim sMSG As String
Dim oElement As eFormElementRO

    On Error GoTo ErrLabel

    bBlank = False 'assume we're ok to begin with
    
    ' Check visit date first
    Set oVisitEFI = oEFI.VisitInstance.VisitEFormInstance
    If Not oVisitEFI Is Nothing Then
        ' Ignore if read-only
        If (Not oVisitEFI.ReadOnly) And (Not oVisitEFI.eForm.EFormDateElement Is Nothing) Then
            Set oElement = oVisitEFI.eForm.EFormDateElement
            Set oResponse = oVisitEFI.Responses.ResponseByElement(oElement)
            bBlank = (oResponse.Value = "") And oResponse.Enterable
            sMSG = "visit"
        End If
    End If
        
    'check eForm date
    'if the visit date isn't missing, and the form has a form date
    If Not bBlank And (Not oEFI.eForm.EFormDateElement Is Nothing) Then
        Set oElement = oEFI.eForm.EFormDateElement
        Set oResponse = oEFI.Responses.ResponseByElement(oElement)
        bBlank = (oResponse.Value = "") And oResponse.Enterable
        sMSG = "eForm"
    End If
    
    If bBlank Then
        'either the form or visit date is missing; decide what to do about it
        sMSG = "This eForm cannot be saved because the " & sMSG & " date is missing."
        Call LogToFile(sMSG, MSG_QUESTION)
    End If
    
    ' OK to save if dates not blank
    OKToSaveEForm = Not bBlank
    
    Set oVisitEFI = Nothing
    Set oResponse = Nothing
    Set oElement = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Revalidator.OKToSaveEForm"
    
End Function

'---------------------------------------------------------------------
Private Sub WhatsChanged(oEFI As EFormInstance)
'---------------------------------------------------------------------
' Report what skips and derivations have changed on the eForm
' NCJ 8 Jul 04 - Consider removed Collect Ifs
'---------------------------------------------------------------------
Dim oResponse As Response

    On Error GoTo ErrLabel
    
    For Each oResponse In oEFI.Responses
        ' Ignore those that won't be saved (last blank rows of RQGs)
        If oResponse.WillBeSaved Then
            ' Only deal with skippies and deriveds
            ' NCJ 8 Jul 04 - Consider removed CollectIfs
            If oResponse.Element.CollectIfCond > "" Or oResponse.SavedStatus = eStatus.NotApplicable Then
                If oResponse.Status <> oResponse.SavedStatus Then
                    If oResponse.Status = eStatus.NotApplicable Then
                        ' It's changed to NotApplicable
                        Call LogToFile(QuestionSpec(oResponse) & " has been disabled ", MSG_QUESTION)
                    End If
                    If oResponse.SavedStatus = eStatus.NotApplicable Then
                        ' It's changed from NotApplicable
                        Call LogToFile(QuestionSpec(oResponse) _
                                & " has been enabled, new status is " & oResponse.StatusString, MSG_QUESTION)
                    End If
                End If
            End If
            If oResponse.Element.DerivationExpr > "" Then
                ' Only report for non-N/A
                If oResponse.Value <> oResponse.SavedValue And oResponse.Status <> eStatus.NotApplicable Then
                    ' It's changed value
                    If oResponse.Value > "" Then
                        Call LogToFile(QuestionSpec(oResponse) & " value has changed to " & oResponse.Value, MSG_QUESTION)
                        ' NCJ 20 May 03 - Also report if its status has changed (Bug 1763)
                        If oResponse.Status <> oResponse.SavedStatus Then
                            Call LogValidationMessage(oResponse, oResponse.Status, oResponse.ValidationMessage)
                        End If
                    Else
                        Call LogToFile(QuestionSpec(oResponse) & " has become blank", MSG_QUESTION)
                    End If
                End If
            End If
            
        End If
    Next

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Revalidator.WhatsChanged"
    
End Sub

'---------------------------------------------------------------------
Private Function AnyNewResponses(oEFI As EFormInstance) As Boolean
'---------------------------------------------------------------------
' Returns TRUE if the eForm has any new responses, i.e. with status Requested
' Logs new responses to log file
' NCJ 7 May 03 - Note that new responses in last row of an RQG will NOT be saved
' so test for new WillBeSaved property here
'---------------------------------------------------------------------
Dim bNewResponses As Boolean
Dim oResponse As Response

    bNewResponses = False
    
    For Each oResponse In oEFI.Responses
        If oResponse.Status = eStatus.Requested And oResponse.WillBeSaved Then
            Call LogToFile("New response created for question: " & oResponse.Element.Code, MSG_QUESTION)
            bNewResponses = True
        End If
    Next
    
    Set oResponse = Nothing
    
    AnyNewResponses = bNewResponses

End Function

'---------------------------------------------------------------------
Private Function EFIHasChanged(oEFI As EFormInstance) As Boolean
'---------------------------------------------------------------------
' NCJ 7 May 03
' Has anything on the EFI changed?
' This takes into account the won't-be-saved responses in last rows of RQGs
'---------------------------------------------------------------------
Dim oResponse As Response
Dim bChanged As Boolean

    bChanged = False
    For Each oResponse In oEFI.Responses
        ' Only consider genuine responses
        If oResponse.WillBeSaved Then
            ' Is it new or changed?
            If oResponse.Changed Or oResponse.Status = eStatus.Requested Then
                ' We've found something that's changed
                bChanged = True
                Exit For
            End If
        End If
    Next
    
    EFIHasChanged = bChanged
    Set oResponse = Nothing
    
End Function

'---------------------------------------------------------------------
Private Function RefreshSubjectProperties() As Boolean
'---------------------------------------------------------------------
' Refresh the subject's properties (DOB, Gender, Label)
' in case the evaluations have changed
' or the SD definitions have changed
' Returns TRUE if anything changed or FALSE otherwise
'---------------------------------------------------------------------
Dim dblDOB As Double
Dim sLabel As String
Dim nGender As Integer
Dim sMSG As String
Dim bChanged As Boolean

    bChanged = False
    
    ' Remember the existing ones
    dblDOB = moSubject.DOB
    sLabel = moSubject.label
    nGender = moSubject.Gender
    
    Call moSubject.RefreshProperties
    
    ' See if any have changed
    If dblDOB <> moSubject.DOB Then
        sMSG = "Subject date of birth changed to "
        If moSubject.DOB <> 0 Then
            Call LogToFile(sMSG & Format(CDate(moSubject.DOB), "yyyy/mm/dd"), MSG_SUBJECT)
        Else
            Call LogToFile(sMSG & "undefined", MSG_SUBJECT)
        End If
        bChanged = True
    End If
    
    If sLabel <> moSubject.label Then
        sMSG = "Subject label changed to " & StringToShow(moSubject.label)
        Call LogToFile("Subject label changed to " & StringToShow(moSubject.label), MSG_SUBJECT)
        bChanged = True
    End If
    
    If nGender <> moSubject.Gender Then
        Call LogToFile("Subject gender changed to " & moSubject.Gender, MSG_SUBJECT)
        bChanged = True
    End If
    
    RefreshSubjectProperties = bChanged

End Function

'---------------------------------------------------------------------
Private Function UpdateEFormLabel(oEFI As EFormInstance) As Boolean
'---------------------------------------------------------------------
' Update the eFormLabel and return TRUE if anything changed
'---------------------------------------------------------------------
Dim sEFormLabel As String

    ' Remember the current eForm label
    sEFormLabel = oEFI.eFormLabel
    Call oEFI.RefreshLabels
    If (oEFI.eFormLabel <> sEFormLabel) Then
        Call LogToFile("EForm label changed to: " & StringToShow(oEFI.eFormLabel), MSG_QUESTION)
        UpdateEFormLabel = True
    Else
        ' No change
        UpdateEFormLabel = False
    End If

End Function

'---------------------------------------------------------------------
Private Function StringToShow(sText As String) As String
'---------------------------------------------------------------------
' Return pair of double quotes for empty string,
' otherwise return sText
'---------------------------------------------------------------------
        
    If sText > "" Then
        StringToShow = sText
    Else
        ' Show empty string as pair of double quotes
        StringToShow = """"""
    End If

End Function

'---------------------------------------------------------------------
Private Sub RemoveUnwantedLogs(ByVal nLineCount As Integer)
'---------------------------------------------------------------------
' Unwind the log messages back to the given line count in non-Verbose mode
'---------------------------------------------------------------------

    If Me.Verbose Then Exit Sub
    
    Do While mcolLogLines.Count > nLineCount
        ' Remove the "last" entry
        mcolLogLines.Remove mcolLogLines.Count
    Loop

End Sub

'---------------------------------------------------------------------
Private Sub FlushLogLines()
'---------------------------------------------------------------------
' Flush the log messages in non-Verbose mode
'---------------------------------------------------------------------
Dim n As Integer
Dim i As Integer

    If Me.Verbose Then Exit Sub
    
    If mcolLogLines.Count > 0 Then
        If msLogFile = "" Then
            Debug.Print "No log file!"
        Else
            n = FreeFile
            Open msLogFile For Append As n
            For i = 1 To mcolLogLines.Count
                Print #n, CStr(mcolLogLines(i))
            Next
            Close n
        End If
    End If

End Sub
