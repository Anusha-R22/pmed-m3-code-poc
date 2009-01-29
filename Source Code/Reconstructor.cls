VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Reconstructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         Reconstructor.cls
' Copyright:    InferMed Ltd. 2006-2007. All Rights Reserved
' Author:       Nicky Johns, November 2006
' Purpose:      Class to handle reconstruction of AREZZO Patient State in MACRO 3.0,
'               based on contents of MACRO DB tables
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 15 Nov 06 - Initial development, based on Revalidator.cls from MACRO BV
'   NCJ 26 Jun 07 - Revisited, continued development
'----------------------------------------------------------------------------------------'

Option Explicit

'#If Reconstruct Then

' The name of the log file
Private msLogFile As String

Private moSubject As StudySubject
Private moStudyDef As StudyDefRO
Private moUser As MACROUser
Private moArezzo As Arezzo_DM

Private mlStudyId As Long
Private msSite As String
Private mlSubjectId As Long

Private msLockToken As String

Private msSubjSpec As String

' Collection of lines for log file
Private mcolLogLines As Collection

' Collection of Visit eForms we've done
Private mcolVisitEForms As Collection

' Count how many things we've done
Private mlDataValues As Long
Private mlEFIs As Long

' To what does the log message refer?
Private Const MSG_TOP_LEVEL = 0
Private Const MSG_SUBJECT = 1
Private Const MSG_VISIT = 2
Private Const MSG_EFORM = 3
Private Const MSG_QUESTION = 4

' Do they want "verbose" output?
Public Verbose As Boolean

'---------------------------------------------------------------------
Public Sub Reconstruct(ByVal lStudyId As Long, ByVal sSite As String, ByVal lSubjectId As Long, _
                    ByVal sStudyName As String, ByVal sSubjLabel As String, oArezzo As Arezzo_DM)
'---------------------------------------------------------------------
' Reconstruct AREZZO for the given subject
'---------------------------------------------------------------------
Dim sErrMsg As String
Dim sSubj As String
Dim bChanged As Boolean

    On Error GoTo ErrLabel
    
    ' Store the info
    mlStudyId = lStudyId
    msSite = sSite
    mlSubjectId = lSubjectId
    
    ' Store the AREZZO
    Set moArezzo = oArezzo
    
    ' Start off empty
    Set moSubject = Nothing
    
    ' Initialise log entries for this subject
    Set mcolLogLines = New Collection
    
    ' Use subject label or ID if not available
    If sSubjLabel > "" Then
        sSubj = sSubjLabel
    Else
        sSubj = lSubjectId
    End If
    
    msSubjSpec = sStudyName & "/" & sSite & "/" & sSubj
    
    Call LogToFile("Subject: " & msSubjSpec)
    Call LogToFile("Reconstruction started at " & CurTime)

    ' Here we have to load the study but don't tell MACRO that we have an AREZZO
    If LoadStudyDef(lStudyId, oArezzo) Then
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
        If SubjectReconstructible Then
            ' Proceed with reconstruction!
            Call SubjectReconstruct
            ' Clear up the subject
            Call moStudyDef.RemoveSubject
            Set moSubject = Nothing
            Call LogToFile(mlDataValues & " non-empty response values in " _
                                & mlEFIs & " eForm instances")
            Call LogToFile("Reconstruction of subject " & msSubjSpec & " completed at " & CurTime & vbCrLf)
        End If
    End If
    
    
    ' Leave the StudyDef in case we want to re-use it for the next subject
'    ' Clear up the study def
'    Call moStudyDef.Terminate
'    Set moStudyDef = Nothing

    ' Finally flush the log messages to file if necessary
'    If bChanged Then
        FlushLogLines
'    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Reconstructor.Reconstruct"

End Sub

'---------------------------------------------------------------------
Private Function LoadStudyDef(ByVal lStudyId As Long, oArezzo As Arezzo_DM) As Boolean
'---------------------------------------------------------------------
' Load the study def into moStudyDef
' Only reload if it's different from what we've already got
' Returns TRUE if study successfully loaded
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
        sErrMsg = moStudyDef.Load(moUser.CurrentDBConString, lStudyId, 1, oArezzo)
        If sErrMsg > "" Then
            ' Give up
            Call LogToFile("Unable to load study: " & sErrMsg, MSG_SUBJECT)
            bLoadedOK = False
        End If
    End If
    
    LoadStudyDef = bLoadedOK

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Reconstructor.LoadStudyDef"

End Function

'---------------------------------------------------------------------
Public Sub InitReconstruction(ByVal sFileName As String, oUser As MACROUser)
'---------------------------------------------------------------------
' Initialise the Reconstruction log file
'---------------------------------------------------------------------
Dim n As Integer
Dim sText As String
Dim oChecker As MACROVersion.Checker

    ' Store the log file name
    msLogFile = sFileName
    Set moUser = oUser
    
    Set oChecker = New Checker
    
    ' Create initial text
    sText = "MACRO " & oChecker.PatchVersion & " Subject Reconstruction Report "
    sText = sText & CurTime & vbCrLf
    sText = sText & "User: " & oUser.UserNameFull & vbCrLf
    sText = sText & "Database: " & oUser.DatabaseCode & vbCrLf
'    sText = sText & "Report mode: "
'    If Me.Verbose Then
'        sText = sText & "Show all"
'    Else
'        sText = sText & "Show changes only"
'    End If
    
    Set oChecker = Nothing
    
    ' Create a new log file
    n = FreeFile
    Open msLogFile For Output As n
    Print #n, sText & vbCrLf
    Close n
    
    Set moStudyDef = Nothing

End Sub

'---------------------------------------------------------------------
Public Sub EndReconstruction()
'---------------------------------------------------------------------
' Write out final line to log file and
' clear up our objects at the end
'---------------------------------------------------------------------
    
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
' Add text to the Reconstruction log file (assume initialised)
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
    Set mcolLogLines = New Collection

End Sub

'---------------------------------------------------------------------
Private Sub Class_Terminate()
'---------------------------------------------------------------------

    Call Terminate

End Sub

'---------------------------------------------------------------------
Private Sub AddRegistrationDetails()
'---------------------------------------------------------------------
' Add person:subjectidentifier if necessary
'---------------------------------------------------------------------

End Sub

'---------------------------------------------------------------------
Private Sub ResetArezzo()
'---------------------------------------------------------------------
' Reset AREZZO by restarting the guideline
' and adding the preliminary data
'---------------------------------------------------------------------
Dim sCountry As String
Dim sAzValues As String

    sAzValues = ""
    
    ' Retrieve existing "country" value
    sCountry = moArezzo.EvaluateExpression("person:country")
    
    ' Restart guideline
    Call moArezzo.StartNewSubject
    
    ' Add person's data to Arezzo - these calls copied from StudySubject.cls
    Call moArezzo.AddData("person:clinicaltrialname", moStudyDef.Name)
    sAzValues = sAzValues & " person:clinicaltrialname = " & moStudyDef.Name & vbCrLf
    
    Call moArezzo.AddData("person:trialsite", msSite)
    sAzValues = sAzValues & " person:trialsite = " & msSite & vbCrLf
    
    Call moArezzo.AddData("person:personid", mlSubjectId)
    sAzValues = sAzValues & " person:personid = " & mlSubjectId & vbCrLf
    
    Call moArezzo.AddData("person:country", sCountry)
    sAzValues = sAzValues & " person:country = " & sCountry & vbCrLf
    
    If moSubject.RegistrationID <> "" Then
        ' Only if reg ID exists - NB: DEBS does this thing with the quotes so we do it here too
        Call moArezzo.AddData("person:subjectidentifier", "'" & ReplaceQuotes(moSubject.RegistrationID) & "'")
        sAzValues = sAzValues & " person:subjectidentifier = " & moSubject.RegistrationID & vbCrLf
    End If
    
    Call moSubject.SetUserProperties(moUser.UserNameFull, moUser.UserRole)
    
    ' Generate events to set the ball rolling
    Call moArezzo.ALM.GuidelineInstance.RunEngine
    
    Call LogToFile(sAzValues)
    
End Sub

'---------------------------------------------------------------------
Private Function SubjectReconstruct() As Boolean
'---------------------------------------------------------------------
' Reconstruct all AREZZO for the current subject
'---------------------------------------------------------------------
Dim oEFI As EFormInstance
Dim sEFISpec As String
Dim sErrMsg As String
Dim sTok1 As String
Dim sTok2 As String
Dim sStateFileName As String
Dim sTaskIDsFileName As String
Dim lEFITaskId As Long
Dim colConfirmEForms As Collection


    On Error GoTo ErrLabel
      
    sStateFileName = gsTEMP_PATH _
                & moSubject.StudyCode & "_" & moSubject.Site & "_" & moSubject.PersonID & ".plm"
    ' Delete file if it already exists (ignore errors if not)
    On Error Resume Next
    Call Kill(sStateFileName)
    On Error GoTo ErrLabel

    ' Must assert the TaskIds
    sTaskIDsFileName = gsTEMP_PATH _
                & moSubject.StudyCode & "_" & moSubject.Site & "_" & moSubject.PersonID & "_TASKIDS.pl"
    ' Delete file if it already exists (ignore errors if not)
    On Error Resume Next
    Call Kill(sTaskIDsFileName)
    On Error GoTo ErrLabel
    
    ' Now do the asserts
    Call AssertTaskIDs(sTaskIDsFileName)
    
    ' Now clear out AREZZO and start from scratch as if a new subject
    Call ResetArezzo

    Set colConfirmEForms = New Collection
    Set mcolVisitEForms = New Collection
    ' Initialise counts
    mlDataValues = 0
    mlEFIs = 0
    
    ' Troll through all the eForm instances
    For Each oEFI In moSubject.eFormInsts
        ' Ignore blank eForms and visit eForms
        If EFIReconstructible(oEFI) Then
            ' Set up EFI description string for future messages
            sEFISpec = EFormSpec(oEFI)
            If moSubject.LoadResponses(oEFI, sErrMsg, sTok1, sTok2) <> lrrCouldNotLockForSave Then
                ' Re-add all the data for this eForm to AREZZO
                Call ReconstructEFI(oEFI)
                ' And do the visit eForm (if any)
                Call ReconstructVEFI(oEFI)
                Call LogToFile("Created AREZZO for eForm " & sEFISpec, MSG_QUESTION)
                ' Increment EFI count
                mlEFIs = mlEFIs + 1
                
                ' Confirm the eForm if possible
                If oEFI.Complete Then
                    If NotDormant(oEFI.eFormTaskId) Then
                        ' Needs confirming
                        Call moArezzo.ConfirmRequestedForm(oEFI.eFormTaskId)
                    Else
                        ' Try the Confirm again later
                        colConfirmEForms.Add oEFI.eFormTaskId
        Debug.Print "eForm " & oEFI.eFormTaskId & " is unexpectedly dormant"
                    End If
                End If
                
                ' Run the engine
                Call moArezzo.ALM.GuidelineInstance.RunEngine
                
                ' Save the interim AREZZO patient state
                Call moArezzo.ALM.SaveState(sStateFileName)
                
                ' Remove responses
                ' NB this doesn't actually remove anything from the current AREZZO!
                Call moSubject.RemoveResponses(oEFI, True)
            Else
                ' Report failure to load
                Call LogToFile("Unable to load responses for eForm " & sEFISpec)
            End If
        End If
    Next
    
    ' Now we need to do Confirms for all the eForms we missed on the way through
    ' e.g. if they were inactive at the time
    If DoTheConfirms(colConfirmEForms) > 0 Then
        ' Some eForms are complete but seem to be still dormant
        Call ReportDormantTasks(colConfirmEForms)
    End If
        
    Set oEFI = Nothing
    Set colConfirmEForms = Nothing
    Set mcolVisitEForms = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Reconstructor.SubjectReconstruct"

End Function

'---------------------------------------------------------------------
Private Function DoTheConfirms(colToConfirm As Collection) As Long
'---------------------------------------------------------------------
' Confirm all the eForms that weren't active first time round
' colToConfirm is a collection of EFITaskIDs that are Complete
' Returns the number of tasks that could not be confirmed (still dormant) - should be 0!
'---------------------------------------------------------------------
Dim vEFITaskId As Variant
Dim lEFITaskId As Long
Dim colTryAgain As Collection
Dim lTasksLeft As Long

    ' Collection to try again if they fail this time
    Set colTryAgain = New Collection
    
    For Each vEFITaskId In colToConfirm
        lEFITaskId = CLng(vEFITaskId)
        If NotDormant(lEFITaskId) Then
            ' Do the confirm
            Call moArezzo.ConfirmRequestedForm(lEFITaskId)
        Else
            ' Save it to try again
            colTryAgain.Add lEFITaskId
        End If
    Next

    ' Any to try again?
    ' Make sure there are fewer than we started with, otherwise we'll loop for ever!
    lTasksLeft = colTryAgain.Count
    If lTasksLeft > 0 And lTasksLeft < colToConfirm.Count Then
        lTasksLeft = DoTheConfirms(colTryAgain)
    End If
    
    Set colTryAgain = Nothing
    DoTheConfirms = lTasksLeft
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Reconstructor.DoTheConfirms"
End Function

'---------------------------------------------------------------------
Private Function NotDormant(lEFITaskId As Long) As Boolean
'---------------------------------------------------------------------
' Returns True if this eForm is active, i.e. not "dormant"
'---------------------------------------------------------------------

    NotDormant = (moArezzo.ALM.GuidelineInstance.colTaskInstances.Item(Str(lEFITaskId)).TaskState <> "dormant")

End Function

'---------------------------------------------------------------------
Private Sub ReportDormantTasks(colTaskIds As Collection)
'---------------------------------------------------------------------
' Report all the tasks that are still dormant (but shouldn't be!)
'---------------------------------------------------------------------
Dim vEFITaskId As Variant
Dim lEFITaskId As Long
Dim oEFI As EFormInstance
Dim sList As String
Dim sComma As String
Dim lCount As Long

    lCount = 0
    sList = ""
    sComma = ""         ' No comma first time round
    
    For Each vEFITaskId In colTaskIds
        lEFITaskId = CLng(vEFITaskId)
        If NotDormant(lEFITaskId) Then
            ' This one's not dormant so it's OK
        Else
            Set oEFI = moSubject.eFIByTaskId(lEFITaskId)
            sList = sList & sComma & oEFI.eFormTaskId & " (" & oEFI.code & ")"
            sComma = ", "       ' Comma before the next one
            lCount = lCount + 1
        End If
    Next
    
    If lCount > 0 Then
        ' There were some to report
        If lCount = 1 Then
            ' Only one
            sList = "Unable to complete this AREZZO eForm task: " & vbCrLf & sList
        Else
            ' More than one
            sList = "Unable to complete these " & lCount & " AREZZO eForm tasks: " & vbCrLf & sList
        End If
        Call LogToFile(sList)
    End If
    
    Set oEFI = Nothing
    
End Sub

'---------------------------------------------------------------------
Private Function SubjectReconstructible() As Boolean
'---------------------------------------------------------------------
' See if we can Reconstruct the current subject
' Assume moSubject not Nothing and not read-only
' Log suitable messages to log file if not
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    ' Always true for now - ******* CHANGE THIS LATER!!! *******
    SubjectReconstructible = True
    
    Exit Function
    
    SubjectReconstructible = False

    If moSubject.LockStatus <> eLockStatus.lsUnlocked Then
        ' Subject is locked/frozen
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
    
    SubjectReconstructible = True

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Reconstructor.SubjectReconstructible"

End Function

'---------------------------------------------------------------------
Private Function EFIReconstructible(oEFI As EFormInstance) As Boolean
'---------------------------------------------------------------------
' See if we can Reconstruct the eForm
' lVEFTaskId is the task id of the visit eForm (which we want to ignore)
' Log suitable messages to log file if not
' Returns sEFISpec as the description of this eForm for the log file
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    EFIReconstructible = False
    
    ' Check that it's an active eForm
    If oEFI Is Nothing Then Exit Function
    
    ' Ignore blank eForms
    If oEFI.Status = eStatus.Requested Then Exit Function
    
    ' Ignore it if it's the Visit eForm
    If IsVisitEForm(oEFI) Then Exit Function

    ' If we get to here, it's OK!
    EFIReconstructible = True

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Reconstructor.EFIReconstructible"

End Function

'---------------------------------------------------------------------
Private Function IsVisitEForm(oEFI As EFormInstance) As Boolean
'---------------------------------------------------------------------
' Is this a visit eForm?
'---------------------------------------------------------------------

    If Not oEFI.VisitInstance.VisitEFormInstance Is Nothing Then
        IsVisitEForm = (oEFI.eFormTaskId = oEFI.VisitInstance.VisitEFormInstance.eFormTaskId)
    Else
        IsVisitEForm = False
    End If

End Function

'---------------------------------------------------------------------
Private Function EFormSpec(oEFI As EFormInstance) As String
'---------------------------------------------------------------------
' Get the "spec" of an eForm instance as a string
' (for output to the log file)
'---------------------------------------------------------------------
Dim sSpec As String

    sSpec = oEFI.VisitInstance.code
    sSpec = AddCycleNo(sSpec, oEFI.VisitInstance.CycleNo)
    sSpec = sSpec & ":" & oEFI.code
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
        AddCycleNo = sStr & "(" & nCycleNo & ")"
    Else
        AddCycleNo = sStr
    End If

End Function

'---------------------------------------------------------------------
Private Sub ReconstructVEFI(ByRef oEFI As EFormInstance)
'---------------------------------------------------------------------
' Do the visit eForm, if there is one, and if we haven't done it before
'---------------------------------------------------------------------
Dim oVEFI As EFormInstance

    On Error GoTo ErrLabel
    
    Set oVEFI = oEFI.VisitInstance.VisitEFormInstance
    If Not oVEFI Is Nothing Then
        ' Ignore it if we have done it before
        If Not CollectionMember(mcolVisitEForms, "K" & oVEFI.eFormTaskId, False) Then
            Call ReconstructEFI(oVEFI)
            ' Remember that we've done it
            Call CollectionAddAnyway(mcolVisitEForms, oVEFI.eFormTaskId, "K" & oVEFI.eFormTaskId)
        End If
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Reconstructor.ReconstructVEFI"

End Sub

'---------------------------------------------------------------------
Private Sub ReconstructEFI(ByRef oEFI As EFormInstance)
'---------------------------------------------------------------------
' Force all the response values on this eForm through to AREZZO
' by calling Validate and Confirm
'---------------------------------------------------------------------
Dim oResponse As Response
Dim sErrMsg As String
Dim lStatus As Long
Dim bSomethingChanged As Boolean
Dim sValue As String

    On Error GoTo ErrLabel
    
    For Each oResponse In oEFI.Responses
        ' Ignore blanks and responses which won't be saved
        sValue = oResponse.Value
        If oResponse.WillBeSaved And sValue <> "" Then
            ' Just force the value through to AREZZO
            ' (We don't care what happens in oResponse because it won't get saved anyway)
            ' Delete first (add "") and then add the real one
            lStatus = oResponse.ValidateValue("", sErrMsg, bSomethingChanged, oResponse.TimeStamp)
            Call oResponse.ConfirmValue("", "", oResponse.UserName, oResponse.UserNameFull)
            ' Now add "real" value
            lStatus = oResponse.ValidateValue(sValue, sErrMsg, bSomethingChanged, oResponse.TimeStamp)
            Call oResponse.ConfirmValue("", "", oResponse.UserName, oResponse.UserNameFull)
            ' Increment data values count
            mlDataValues = mlDataValues + 1
        End If
    Next    ' oResponse
        
    Set oResponse = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Reconstructor.ReconstructEFI"

End Sub

'---------------------------------------------------------------------
Private Sub LogValidationMessage(oResponse As Response, _
                                ByVal nStatus As Integer, _
                                ByVal sValMsg As String)
'---------------------------------------------------------------------
' Log a Reconstruction message for a response
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
    Err.Raise Err.Number, , Err.Description & "|" & "Reconstructor.LogValidationMessage"
    
End Sub

'---------------------------------------------------------------------
Private Function QuestionSpec(oResponse As Response) As String
'---------------------------------------------------------------------
' Get the description for a question, preceded by "Question "
'---------------------------------------------------------------------
    
    QuestionSpec = "Question " & AddCycleNo(oResponse.Element.code, oResponse.RepeatNumber, _
                        (Not oResponse.Element.OwnerQGroup Is Nothing))

End Function


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
    Err.Raise Err.Number, , Err.Description & "|" & "Reconstructor.OKToSaveEForm"
    
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

'#End If

'---------------------------------------------------------------------
Private Sub AssertTaskIDs(sFileName As String)
'---------------------------------------------------------------------
' Assert the predefined task IDs from the MACRO subject
' as macro_predefinedID/4 clauses
' macro_predefinedID( TaskName, ParentTaskID, CycleNo, TaskID )
'---------------------------------------------------------------------
Dim sAsserts As String
Dim oVI As VisitInstance
Dim oEFI As EFormInstance
Dim sQuery As String
Dim sR As String

Const sPREDEFINED = "macro_predefinedID("
Const nSTUDY_TASKID As Integer = 10000

    On Error GoTo ErrLabel
    
    sAsserts = ""
    
    For Each oVI In moSubject.VisitInstances
        sAsserts = sAsserts & sPREDEFINED _
                    & LCase(oVI.code) & "," & nSTUDY_TASKID & "," & oVI.CycleNo & "," & oVI.VisitTaskId _
                    & "). " & vbCrLf
        For Each oEFI In oVI.eFormInstances
            sAsserts = sAsserts & sPREDEFINED _
                        & LCase(oEFI.code) & "," & oVI.VisitTaskId & "," & oEFI.CycleNo & "," & oEFI.eFormTaskId _
                        & "). " & vbCrLf
        Next
    Next
    
    ' Remove any previous IDs
    sQuery = "retractall(macro_predefinedID(_,_,_,_)), write( '0000' ). "
    Call moSubject.Arezzo.ALM.GetPrologResult(sQuery, sR)
    
    If sAsserts <> "" Then
        ' Create a file from the asserts
        Call StringToFile(sFileName, sAsserts)
        ' Load the file into Prolog
        sQuery = "ensure_loaded( '" & sFileName & "' ), write( '0000' ). "
        Call moSubject.Arezzo.ALM.GetPrologResult(sQuery, sR)
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Reconstructor.AssertTaskIDs"
    
End Sub
