VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Arezzo_DM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: Arezzo_DM.cls
' Copyright InferMed 2001 - 2004, All Rights Reserved
' Author: Nicky Johns, InferMed, June 2001
' Arezzo class for MACRO DM 2.2
' Provides Data Entry functions and Arezzo support for MACRO
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'-------------------------------------------------------------------
' NCJ 12-21 Jun 01 - Initial development
' NCJ 11 Jul 01 - Changed text space from 1000 to 2000
' NCJ 30 Jul 01 - SetQuestionStatus
' NCJ 31 Jul 01 - AddQData -> ReplaceQData (so no need to delete prev. value first)
' NCJ 14 Aug 01 - Events, starting new protocol
' NCJ 31 Aug 01 - Load Macro_Arezzo.pc
' NCJ 10 Sep 01 - Made ALM object public
' NCJ 17 Sep 01 - Pass sTempDir to Init
' NCJ 26 Sep 01 - Pass Prolog memory settings to Init
' NCJ 28 Sep 01 - Added CheckStudy to check for infinite cycles
' NCJ 28 Sep 01 - Bizarre bug: changed moStudyInstance to moALM.GuidelineInstance
'               (but still need it for collecting events)
' NCJ 1 Oct 01 - GetArezzoTasks to return non_MACRO Arezzo tasks
'               No longer process them in the event loop
' NCJ 2 Oct 01 - Changed AddMeValue and DeleteMeValue
'               We now do automatic Add me:value in ValidateResponse,
'               and DeleteMeValue is now called DeleteLocalValues
'           Added Error Handlers
' NCJ 18 Oct 01 - Pass Prolog switches to Init (rather than individual Program/Text space values)
' NCJ 13 Feb 02 - Added new RunEngine sub (required for new eForms Arezzo patch)
' MACRO 3.0
' NCJ 5 July 02 - Merged 2.2 with 3.0
' NCJ 8 July 02 - Changed Macro_Arezzo3.pc to Macro3_Arezzo.pc
' NCJ 29 Aug 02 - Changed ALM4 to ALM5
' NCJ 20 Sept 02 - Ensure eForm and Question contexts correctly set for AREZZO activities
' NCJ 8 Nov 02 - Added GetHotlinkTarget function
' NCJ 25-27 Nov 02 - Various improvements using new "mini" ALM
' NCJ 9 Dec 02 - Removed some debugging and some unused code
' NCJ 24 Jan 03 - Added SetUserProperties
' NCJ 31 Jan 03 - New routines to handle non-MACRO Arezzo activities
' NCJ 5 Mar 03 - Removed SetQuestionStatus because it doesn't work...
' NCJ 4 Dec 03 - Changed collection key in GetArezzoTasks
'               Changed QuestionContext call to new shorter version
' NCJ 24 Mar 04 - Explicitly clear collections when terminating (cure for memory leaks)
' NCJ 6 May 04 - Don't do GetNewAREZZOTasks in ConfirmRequestedForm (to improve performance)
' NCJ 3 Mar 05 - Bug 2537 - Handle file of results in GetNewAREZZOTasks
' NCJ 8 Dec 05 - GetArezzoDate now handles PDs
' NCJ 5-11 Jan 06 - Added FormatDoubleDate (for eForm and Visit dates)
'--------------------------------------------------------------------

Option Explicit

' These are the events that Arezzo will pass through to MACRO
' - New eForm Instance
' - New Visit Instance
' - Idle event (engine has finished current phase of processing)
' -
Public Event NewEFormInstance(lEFormId As Long, lEFormTaskId As Long, lVisitTaskId As Long, nCycleNo As Integer)
Public Event NewVisitInstance(lVisitId As Long, lVisitTaskId As Long, nCycleNo As Integer)
Public Event Idle(bPatientDirty As Boolean)

' The ALM
Private moALM As ALM5
' The Guideline Instance
Private WithEvents moStudyInstance As GuidelineInstance
Attribute moStudyInstance.VB_VarHelpID = -1

Private Const msME_VALUE As String = "me:value"

' Store whether we need to re-generate Arezzo events
Private mbNeedToGenerateEvents As Boolean
' Has there been an Arezzo state change?
Private mbPatientDirty As Boolean

' Store fixed lengths of task name prefixes here for greater efficiency
Private mnLenEFormPlanPrefix As Integer
Private mnLenVisitPlanPrefix As Integer
Private mnLenEFormActionPrefix As Integer

' The predefined prefixes used for Arezzo tasks
Private Const ms_VISIT_PLAN_PREFIX As String = "Visit "
Private Const ms_EFORM_PLAN_PREFIX As String = "Form "
Private Const ms_EFORM_ACTION_PREFIX As String = "Data_Entry_"

Private Const msAREZZO_SEPARATOR As String = "|"

' NCJ 31 Jan 03
Private mcolArezzoActivities As Collection

'----------------------------------------------------
Public Sub Init(Optional sTempDir As String = "", _
                Optional sMemorySwitches As String = "")
'----------------------------------------------------
' Initialise the Arezzo engine for use with MACRO.
' sTempDir is the writeable directory for Arezzo's temp files (may be "").
' sMemorySwitches is (optional) set of switch values for Prolog's memory.
'----------------------------------------------------
' NCJ 17/9/01 - Added sTempDir
' NCJ 26/9/01 - Added nProgramSpace and nTextSpace
' NCJ 18/10/01 - Replaced nProgramSpace and nTextSpace with sMemorySwitches
'----------------------------------------------------
Dim sFile As String
Dim sGoal As String
Dim sR As String
Dim sPrologErr As String

    On Error GoTo ErrLabel

    mbPatientDirty = False
    mbNeedToGenerateEvents = False
    
    Set moALM = New ALM5
    ' NCJ 26 Nov 02 - Run ALM as "mini version"
    Call moALM.StartALM(sMemorySwitches, True)
    
    Set moStudyInstance = moALM.GuidelineInstance
    Call SetUpPLMEvents
    
    ' Set TempDirectory
    If sTempDir > "" Then
        moALM.TempDirectory = sTempDir
    End If
    
    ' Load the MACRO cycling data add-on (ignore result)
    ' NCJ 8 July 02 - Changed Macro_Arezzo3.pc to Macro3_Arezzo.pc
    sFile = App.Path & "\Macro3_Arezzo.pc"
    sGoal = "ensure_loaded('" & sFile & "'), write('0000'). "
    sPrologErr = moALM.GetPrologResult(sGoal, sR)

    ' NCJ 31 Jan 03
    Set mcolArezzoActivities = New Collection
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.Init"

End Sub

'----------------------------------------------------
Public Sub Finish()
'----------------------------------------------------
' Finish with the Arezzo_DM object
'----------------------------------------------------

    Call CloseMe

End Sub

'---------------------------------------------------------------------
Public Property Get ALM() As ALM5
'---------------------------------------------------------------------
' Return the Arezzo ALM object (required by Web for PfToJavaScript)
'---------------------------------------------------------------------

    Set ALM = moALM

End Property

'---------------------------------------------------------------------
Friend Function CheckStudy() As String
'---------------------------------------------------------------------
' Returns empty string if study is OK.
' Returns appropriate error message string if study contains an infinite cycling visit.
'---------------------------------------------------------------------
Dim sBadVisit As String
Dim sR As String
Dim sMsg As String

    On Error GoTo ErrLabel

    sMsg = ""
    sBadVisit = moALM.GetPrologResult("macro_is_infinite. ", sR)
    If sR = "0000" And sBadVisit > "" Then
        ' Yes, there were infinite cycles
        sMsg = "This study cannot be opened because it contains " _
                & vbCrLf & "an infinitely cycling visit '" _
                & sBadVisit & "'."
        sMsg = sMsg & vbCrLf & vbCrLf _
                & "The study definition must be changed before it can be used for subject data."
    End If
    CheckStudy = sMsg

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.CheckStudy"

End Function

'---------------------------------------------------------------------
Friend Sub SetMACROTaskContext(oVisitInst As VisitInstance, oEFormInst As EFormInstance)
'---------------------------------------------------------------------
' Set the context for Arezzo expression and condition evaluation.
' Either oVisitInst or oEFormInst may be Nothing.
'---------------------------------------------------------------------
Dim sGoal As String
Dim sR As String

    On Error GoTo ErrLabel

    sGoal = SetMACROEFormContextString(oVisitInst, oEFormInst)
    
    sGoal = sGoal & " write('0000'). "
    ' Run this Prolog goal but ignore the result
    Call moALM.GetPrologResult(sGoal, sR)

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.SetMACROTaskContext"

End Sub

'---------------------------------------------------------------------
Private Function SetMACROEFormContextString(oVisitInst As VisitInstance, oEFormInst As EFormInstance) As String
'---------------------------------------------------------------------
' Returns string representing retracts/asserts for setting MACRO context
' String terminates with a comma
' NCJ 4 Dec 03 - New calls which are more efficient! Use macro_setv(VTaskId) and macro_setf(EFormTaskId)
'---------------------------------------------------------------------
Dim sGoal As String

'Const sRETRACTALL = "retractall(macro_context(_,_)), retractall(macro_context(_)), "

    On Error GoTo ErrLabel

'    sGoal = sRETRACTALL
    If Not oEFormInst Is Nothing Then
        ' Assert the eForm Task ID
'        sGoal = sGoal & "assert(macro_context(" & oEFormInst.eFormTaskId & ")), "
        sGoal = "macro_setf(" & oEFormInst.eFormTaskId & "),"
    Else
        ' We only have a visitinstance
        If Not oVisitInst Is Nothing Then
'            sGoal = sGoal & "assert(macro_context(" & oVisitInst.VisitTaskId & ", 0)), "
            sGoal = "macro_setv(" & oVisitInst.VisitTaskId & "),"
        Else
            sGoal = "macro_setv(0),"
        End If
    End If

    SetMACROEFormContextString = sGoal

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.SetMACROEFormContextString"

End Function

'---------------------------------------------------------------------
Friend Function GetHotlinkTarget(sHotlink As String, lCurrentEformTaskId As Long) As Long
'---------------------------------------------------------------------
' Return the EFormTaskId corresponding to the given Hotlink string
' lCurrentEformTaskId is the current eForm
' Returns 0 if Hotlink cannot be calculated
'---------------------------------------------------------------------
Dim sGoal As String
Dim sResult As String

    On Error GoTo ErrLabel
    
    sGoal = "macro_hotlink_ID( `" & sHotlink & "`," & lCurrentEformTaskId & "). "
    GetHotlinkTarget = CLng(moALM.GetPrologResult(sGoal, sResult))
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.GetHotlinkTarget(" & sHotlink & ")"

End Function

'---------------------------------------------------------------------
Friend Sub SetUserProperties(sUserNameFull As String, sRole As String)
'---------------------------------------------------------------------
' NCJ 24 Jan 03 - Store the details of the current user in AREZZO
'---------------------------------------------------------------------
Dim sGoal As String
Dim sR As String

    On Error GoTo ErrLabel

    sGoal = "macro_set_user_properties('" & sUserNameFull & "','" & sRole & "'),"
    sGoal = sGoal & " write('0000'). "
    ' Run this Prolog goal but ignore the result
    Call moALM.GetPrologResult(sGoal, sR)

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.SetUserProperties"

End Sub


'---------------------------------------------------------------------
Public Function ValidExpression(ByVal sExpr As String) As Boolean
'---------------------------------------------------------------------
' Returns TRUE if sExpr represents a syntactically valid Arezzo expression
'---------------------------------------------------------------------

    ValidExpression = moALM.IsValidExpression(sExpr)
    
End Function

'---------------------------------------------------------------------
Public Function ValidCondition(ByVal sExpr As String) As Boolean
'---------------------------------------------------------------------
' Returns TRUE if sExpr represents a syntactically valid Arezzo condition
'---------------------------------------------------------------------

    ValidCondition = moALM.IsValidCondition(sExpr)

End Function

'---------------------------------------------------------------------
Public Function ConditionTrue(ByVal sCondition As String) As Boolean
'---------------------------------------------------------------------
' Returns TRUE if sCondition is currently True in Arezzo
' or FALSE otherwise
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    ConditionTrue = (moALM.GuidelineInstance.Evaluate(sCondition) = "true")

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.ConditionTrue"

End Function

'----------------------------------------------------
Public Function EvaluateDate(ByVal sExpr As String) As Double
'----------------------------------------------------
' Evaluate an Arezzo expression with an expected date result
' and return a VB double.
' Returns 0 if expression cannot be evaluated as a date
'----------------------------------------------------
    
    On Error GoTo ErrLabel
    
    EvaluateDate = ArezzoDateToDouble(moALM.GuidelineInstance.Evaluate(sExpr))

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.EvaluateDate"

End Function

'----------------------------------------------------
Public Function ResultOK(sResult As String) As Boolean
'----------------------------------------------------
' Test whether the given result represents a valid Arezzo result.
' Returns FALSE if it represents an error or other failure.
' sResult is as returned from EvaluateExpression or EvaluateExpressions.
'----------------------------------------------------

    ResultOK = (sResult <> "Syntax error")
    
End Function

'----------------------------------------------------
Public Function EvaluateExpression(ByVal sExpr As String) As String
'----------------------------------------------------
' Evaluate an Arezzo expression.
' Should be followed by call to ResultOK to check if any error occurred.
'----------------------------------------------------

    On Error GoTo ErrLabel
    
    EvaluateExpression = moALM.GuidelineInstance.Evaluate(sExpr)

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.EvaluateExpression"

End Function

'----------------------------------------------------
Public Function EvaluateExpressions(colExprs As Collection) As Collection
'----------------------------------------------------
' Evaluate each expression string in the given collection
' and return a corresponding collection of result strings.
' Call ResultOK for each result in collection.
'----------------------------------------------------
    
    On Error GoTo ErrLabel
    
    Set EvaluateExpressions = moALM.GuidelineInstance.EvaluateCollection(colExprs)

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.EvaluateExpressions"

End Function

'----------------------------------------------------
Friend Property Get ArezzoDef() As String
'----------------------------------------------------
' The Arezzo guideline definition.
' Load a guideline by setting this property
'----------------------------------------------------

    ArezzoDef = moALM.ArezzoFile

End Property

'----------------------------------------------------
Friend Property Let ArezzoDef(ByVal sArezzo As String)
'----------------------------------------------------
' The Arezzo guideline definition.
' Load a guideline by setting this property
' NCJ 10 May 04 - Set the CLM mode to "run" for greater efficiency
'----------------------------------------------------

    On Error GoTo ErrLabel

    Call SetCLMRunMode      ' We're going to be running the guideline
    moALM.ArezzoFile = sArezzo
    
Exit Property
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.ArezzoDef"

End Property

'----------------------------------------------------
Private Sub SetCLMRunMode()
'----------------------------------------------------
' NCJ 10 May 04
' Improved efficiency on loading AREZZO by setting CLM mode
'----------------------------------------------------
Dim sGoal As String
Dim sR As String

    ' Ignore errors so there's no problem with earlier versions of Arezzo5.pc
    On Error Resume Next
    
    sGoal = "clm_set_mode(run), write('0000'). "
    ' Run this Prolog goal but ignore the result
    Call moALM.GetPrologResult(sGoal, sR)

End Sub

'----------------------------------------------------
Public Sub StartNewSubject()
'----------------------------------------------------
' Start a new MACRO subject
'----------------------------------------------------

    Call moALM.GuidelineInstance.Start

End Sub


'----------------------------------------------------
Friend Property Get PatientState() As String
'----------------------------------------------------
' Get the current Arezzo patient state
'----------------------------------------------------

    PatientState = moALM.GetState

End Property

'----------------------------------------------------
Friend Property Let PatientState(ByVal sState As String)
'----------------------------------------------------
' The Arezzo patient state. Set this property to load an existing subject's data.
' NB The ArezzoDef must have been set first
'----------------------------------------------------

    On Error GoTo ErrLabel

    moALM.GuidelineInstance.Clear   ' Clear out the previous one
    mbPatientDirty = False
    mbNeedToGenerateEvents = False
    
    moALM.SetState sState

Exit Property
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.PatientState (Let)"

End Property

'----------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------
'----------------------------------------------------

    Set moALM = Nothing
    Set moStudyInstance = Nothing
    
    ' Store fixed string lengths here
    mnLenEFormPlanPrefix = Len(ms_EFORM_PLAN_PREFIX)
    mnLenVisitPlanPrefix = Len(ms_VISIT_PLAN_PREFIX)
    mnLenEFormActionPrefix = Len(ms_EFORM_ACTION_PREFIX)

    mbPatientDirty = False
    mbNeedToGenerateEvents = False
    
End Sub

'----------------------------------------------------
Private Sub CloseMe()
'----------------------------------------------------
' Tidy up and close ourselves down
'----------------------------------------------------
    
    Set moStudyInstance = Nothing
    Call CollectionRemoveAll(mcolArezzoActivities)  ' NCJ 24 Mar 04
    Set mcolArezzoActivities = Nothing
    If Not moALM Is Nothing Then
        moALM.CloseALM
        Set moALM = Nothing
    End If

End Sub

'----------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------

    Call CloseMe

End Sub

'---------------------------------------------------------
Private Sub SetUpPLMEvents()
'---------------------------------------------------------
' Set up the Arezzo's PLM events
' Initialise event filter
'---------------------------------------------------------
Dim colEvents As Collection

    ' NCJ 26 Nov 02 - New faster method
    Set colEvents = New Collection
    colEvents.Add "idle"
    colEvents.Add "confirm_cands"
    colEvents.Add "state_change(plan,in_progress)"
    
    Call moALM.GuidelineInstance.EventFilter.SetFilters(colEvents)
    
    Set colEvents = Nothing
    
End Sub

'----------------------------------------------------------------
Public Function ValidateDateFormat(ByVal sFormatString As String) As Integer
'----------------------------------------------------------------
' NCJ 20 Jan 03 - Added for Internationlisation stuff
' Parse a date format as entered by the user in sFormatString.
' Returns 0 if NOT a valid date/time format, otherwise
' returns integer 1 - 10 (NCJ 8 Dec 05 - new range of date formats)
'----------------------------------------------------------------

    ValidateDateFormat = moALM.ValidateDateFormat(sFormatString)

End Function

'----------------------------------------------------------------
Public Function ReadValidDate(ByVal sDateString As String, _
                                ByVal sFormatString As String, _
                                ByRef sArezzoDate As String) As String
'----------------------------------------------------------------
' Parse a date as entered by the user in sDateString according to the format in sFormatString.
' Returns correctly formatted date string and Arezzo equivalent in sArezzoDate.
'----------------------------------------------------------------
' (see also ConvertDateFromArezzo)
'----------------------------------------------------------------

    ReadValidDate = moALM.ReadFormattedDate(sFormatString, sDateString, sArezzoDate)

End Function

'----------------------------------------------------------------
Public Function FormatDate(ByVal sArezzoDate As String, _
                            ByVal sFormat As String) As String
'----------------------------------------------------------------
' Format the Arezzo date (as date(Y,M,D)) in the format given in sFormat (e.g. "dd/mm/yyyy").
' Returns empty string if date or format isn't valid.
'----------------------------------------------------------------

    FormatDate = moALM.FormatDate(sArezzoDate, sFormat)
    
End Function

'----------------------------------------------------------------
Public Function FormatDoubleDate(ByVal dblDate As Double, ByVal sFormat As String) As String
'----------------------------------------------------------------
' NCJ 5 Jan 06 - Format a Double date (needed for eForm and Visit dates)
' Have to go via an AREZZO date in case it's partial
'----------------------------------------------------------------

    FormatDoubleDate = moALM.FormatDoubleDate(dblDate, sFormat)

End Function

'----------------------------------------------------------------
Public Function ArezzoDateToDouble(ByVal sArezzoDate As String) As Double
'----------------------------------------------------------------
' Convert from Arezzo date string
' to VB's internal format as double.
' Returns 0 if sArezzoDate not a valid date
'----------------------------------------------------------------
' Assume sArezzoDate is one of
'   "date(Y,M,D)"
'   "time(H,Mn,S)"
'   "date(Y,M,D,H,Mn,S)"
'----------------------------------------------------------------

    ArezzoDateToDouble = moALM.ArezzoDateToDouble(sArezzoDate)

End Function

'------------------------------------------------
Public Function IsPartialDate(ByVal sArezzoDate As String) As Boolean
'------------------------------------------------
' NCJ 8 Dec 05 - Returns TRUE if sArezzoDate represents a date with Month and/or Day = 0
' sArezzoDate should be date/[3,6] or time/3 term as a string
'------------------------------------------------

    IsPartialDate = moALM.IsPartialDate(sArezzoDate)
    
End Function

'----------------------------------------------------------------
Private Sub moStudyInstance_ConfirmCandidates(DecisionKey As String, Candidates As Collection)
'----------------------------------------------------------------
' Confirm the commitment of non-recommended candidates.
' Simply by-pass the user and re-commit
'---------------------------------------------------------
Dim oDecision As TaskInstance

    Set oDecision = moALM.GuidelineInstance.colTaskInstances.Item(DecisionKey)
    oDecision.CommitCandidates Candidates
    ' Do a check on other events after we've done it
    mbNeedToGenerateEvents = True

End Sub

'----------------------------------------------------------------
Private Sub moStudyInstance_Idle()
'----------------------------------------------------------------
' The Engine has finished its current phase of processing.
'----------------------------------------------------------------

    ' See if something happened during last cycle of processing
    If mbNeedToGenerateEvents Then
        mbNeedToGenerateEvents = False
        moALM.GuidelineInstance.RunEngine     ' Ensure we get another Idle event
        GenerateArezzoEvents
    Else
        ' Only notify MACRO on the "final" Idle event
        RaiseEvent Idle(mbPatientDirty)
        mbPatientDirty = False
    End If
    
End Sub

'----------------------------------------------------------------------
' NCJ 17 Feb 03 - Removed unused GetArezzoValue
'----------------------------------------------------------------------

'----------------------------------------------------------------------
Public Function GetArezzoDate(dblDate As Double) As String
'----------------------------------------------------------------------
' For the given VB date, return the equivalent Arezzo format date,
' as date(Y,M,D) or time(H,M,S) etc.
' NCJ 8 Dec 05 - Use new ALM DoubleToArezzoDate (handles partial dates)
'----------------------------------------------------------------------

    GetArezzoDate = moALM.DoubleToArezzoDate(dblDate)
    
End Function

'----------------------------------------------------------------------
Friend Function ValidateResponse(ByVal sDName As String, _
                        ByVal nCycle As Integer, _
                        oEFI As EFormInstance, _
                        ByVal sValue As String, _
                        ByRef sWFlag As String) As Long
'----------------------------------------------------------------------
' Check the response value given by validating it in Arezzo.
' sDName is the data item name (i.e. MACRO question code),
' and nCycle is its repeat number
' sValue is the (Arezzo-ised) value.
' NB This temporarily adds the value and validates it.
' It should be removed (see DeleteLocalValues) afterwards.
' Returns a MACRO eResponseValidation result.
' NCJ 20 Dec 01 - Set the question context here
' NCJ 19 Sept 02 - Must also set Visit/Eform context because it might have changed since last time
' (because of Visit eForms) Pass in oEFI too
'----------------------------------------------------------------------
Dim sV As String

    On Error GoTo ErrLabel

    ' Set the question context
    Call SetQuestionContext(sDName, nCycle, oEFI)
    ' Add the "me:value" data item
    Call AddData(msME_VALUE, sValue)
    ' Add the local value
'    Call AddData(sDName, sValue)
'    ' Note that the DataError routine returns 0,1,2,3 or 4 as in eResponseValidation
'    ValidateResponse = moALM.GuidelineInstance.colDataValues.DataError(sDName, sV, sWFlag)
    
'    ' NCJ 26 Nov 02 - more efficient version
    ValidateResponse = moALM.GuidelineInstance.colDataValues.AddCheckData(sDName, sValue, "", sWFlag)
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.ValidateResponse(" & sDName & ")"

End Function

'----------------------------------------------------------------------
Friend Sub SetQuestionContext(ByVal sDName As String, ByVal nCycle As Integer, _
                            oEFI As EFormInstance)
'----------------------------------------------------------------------
' Set the question context for Arezzo
' NCJ 19 Sept 02 - Must include EFI here too
'----------------------------------------------------------------------
Dim sGoal As String
Dim sR As String

    ' Get the part than sets the visit/eform context
    sGoal = SetQuestionContextString(sDName, nCycle, oEFI)
    sGoal = sGoal & ", write('0000'). "
    ' Run this Prolog goal but ignore the result
    Call moALM.GetPrologResult(sGoal, sR)

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.SetQuestionContext(" & sDName & ")"
    
End Sub

'----------------------------------------------------------------------
Friend Function SetQuestionContextString(ByVal sDName As String, ByVal nCycle As Integer, _
                                        oEFI As EFormInstance) As String
'----------------------------------------------------------------------
' Get the Prolog goal to set the question context for Arezzo
' Does NOT have a comma or a full stop at the end
' NCJ 4 Dec 03 - If oEFI is Nothing, ONLY do the question context, and use new macro_setq/2
'----------------------------------------------------------------------
Dim sGoal As String

    sGoal = ""
    ' Get the part than sets the visit/eform context
    If Not oEFI Is Nothing Then
        sGoal = SetMACROEFormContextString(oEFI.VisitInstance, oEFI)    ' Ends with comma
    End If
'    sGoal = sGoal & "retractall(macro_question(_,_)), "
'    sGoal = sGoal & "assert(macro_question( " & LCase(sDName) & "," & nCycle & ")) "
    sGoal = sGoal & "macro_setq(" & LCase(sDName) & "," & nCycle & ")"

    SetQuestionContextString = sGoal
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.SetQuestionContextString"
    
End Function

'----------------------------------------------------------------------
Public Sub RemoveQuestionContext()
'----------------------------------------------------------------------
' Remove the question context for Arezzo
' (but we don't remove the visit/eForm context)
'----------------------------------------------------------------------
Dim sGoal As String
Dim sR As String

    sGoal = "retractall(macro_question(_,_)), write('0000'). "
    ' Run this Prolog goal but ignore the result
    Call moALM.GetPrologResult(sGoal, sR)

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.RemoveQuestionContext"
    
End Sub

'----------------------------------------------------------------------
Friend Sub DeleteLocalValues(ByVal sDName As String)
'----------------------------------------------------------------------
' Delete me:value added by ValidateResponse
' and delete the "local" value too
'----------------------------------------------------------------------
Dim sGoal As String
Dim sR As String
Dim colDNames As Collection

    On Error GoTo ErrLabel

    Set colDNames = New Collection
    colDNames.Add msME_VALUE
    colDNames.Add sDName
    Call moALM.GuidelineInstance.colDataValues.RemoveAll(colDNames)
    Set colDNames = Nothing
    
    ' Remove the question context, too
    Call RemoveQuestionContext
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.DeleteLocalValues(" & sDName & ")"

End Sub

'----------------------------------------------------------------------
Public Sub AddData(ByVal sDName As String, ByVal sValue As String)
'----------------------------------------------------------------------
' Add data value to Arezzo WITHOUT a user tag.
' If sValue = "", data item is deleted.
'----------------------------------------------------------------------
' Use this for unvalidated MACRO-specific things like person:trialid etc.
'----------------------------------------------------------------------
    
    On Error GoTo ErrLabel

    If sValue > "" Then
        Call moALM.GuidelineInstance.colDataValues.Add(sDName, sValue, False)
    Else
        Call moALM.GuidelineInstance.colDataValues.Delete(sDName)
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.AddData(" & sDName & ")"

End Sub

'---------------------------------------------------------
Public Sub DeleteData(ByVal sDName As String)
'---------------------------------------------------------
' Delete all data values for given data name
'---------------------------------------------------------

    Call moALM.GuidelineInstance.colDataValues.Delete(sDName)

End Sub

'---------------------------------------------------------
Friend Sub ReplaceQData(ByVal sDName As String, _
                            ByVal sValue As String, _
                            ByVal sUserRef As String, _
                            Optional sRefPrefix As String = "")
'---------------------------------------------------------
' Add data value WITH a user reference.
' This automatically deletes any previous value with this user reference
' (But only if sUserRef > "").
' Optional sRefPrefix to add to front of UserRef if non-zero.
' If sValue = "", delete sDName
'---------------------------------------------------------
' NCJ 31/7/01 - Use new Replace function
' NCJ 20 Dec 01 - Changed lUserRef to sUserRef
' NCJ 14 Feb 02 - If value is empty, delete data value
'---------------------------------------------------------
Dim sAzUserRef As String

    On Error GoTo ErrLabel

    If sUserRef > "" Then
        sAzUserRef = sRefPrefix & sUserRef
    Else
        sAzUserRef = ""
    End If
    
    If sValue = "" Then
        ' Empty value - delete data item
        Call RemoveDataByRef(sAzUserRef)
    Else
        ' Value is non-empty so replace existing
        ' NCJ 30 Apr 03 - Do not run engine when adding data!
        Call moALM.GuidelineInstance.colDataValues.Replace(sDName, sValue, sAzUserRef, False)

    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.ReplaceQData(" & sDName & ")"

End Sub

'---------------------------------------------------------
Friend Sub RemoveDataByRef(ByVal sUserTag As String)
'---------------------------------------------------------
' Delete the data value with this UserTag.
'---------------------------------------------------------

    On Error GoTo ErrLabel
    
    ' NCJ 26 Nov 02 - New ALM supports RemoveByTag
    Call moALM.GuidelineInstance.colDataValues.RemoveByTag(sUserTag)
  
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.RemoveDataByRef(" & sUserTag & ")"

End Sub

'---------------------------------------------------------
Friend Sub SetQuestionStatus(sDataItemCode As String, nStatus As Integer)
'---------------------------------------------------------
' Set question status to Missing, Unobtainable or Not Applicable.
' NCJ 5 Mar 03 - THIS DOESN'T WORK FOR MACRO!!
' Reinstate when we have proper support for status(Qu) in AREZZO
'---------------------------------------------------------

    On Error GoTo ErrLabel

'    Select Case nStatus
'    Case eStatus.Missing
'        moALM.GuidelineInstance.colDataValues.Status(sDataItemCode) = "unknown"
'    Case eStatus.Unobtainable
'        moALM.GuidelineInstance.colDataValues.Status(sDataItemCode) = "u_o"
'    Case eStatus.NotApplicable
'        moALM.GuidelineInstance.colDataValues.Status(sDataItemCode) = "n_a"
'    Case Else
'        ' Ignore others for Arezzo
'    End Select

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.SetQuestionStatus(" & sDataItemCode & ")"

End Sub

'---------------------------------------------------------
Private Sub DealWithInProgressPlan(oPlan As TaskInstance)
'---------------------------------------------------------
' Deal with an in_progress Arezzo "plan" task.
' oPlan assumed to be of type "plan".
'---------------------------------------------------------
            
    On Error GoTo ErrLabel

    If IsMacroVisitPlan(oPlan.Caption) Then
        ' It's a Visit plan being activated
        ' Raise a "New Visit Instance" event
        RaiseEvent NewVisitInstance(oPlan.CLMTaskKey, oPlan.TaskKey, oPlan.CycleNumber)
        
    ElseIf IsMacroEFormPlan(oPlan.Caption) Then
        ' It's an eForm plan being activated
        ' Raise a "New eForm Instance" event
        RaiseEvent NewEFormInstance(oPlan.CLMTaskKey, oPlan.TaskKey, _
                                    oPlan.ParentPlan, oPlan.CycleNumber)

    Else
        ' Ignore other plans

    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.DealWithInProgressPlan"

End Sub

'---------------------------------------------------------
Private Function IsMacroEFormPlan(ByVal sCaption As String) As Boolean
'---------------------------------------------------------
' Is this task a MACRO eForm plan?
' Its caption starts with predefined prefix
'---------------------------------------------------------

    IsMacroEFormPlan = False      ' Default value
    
    If Len(sCaption) > mnLenEFormPlanPrefix Then   ' Is caption long enough to consider?
        ' Compare the caption prefix with our predefined prefix
        If Left(sCaption, mnLenEFormPlanPrefix) = ms_EFORM_PLAN_PREFIX Then
            IsMacroEFormPlan = True
        End If
    End If

End Function

'---------------------------------------------------------
Private Function IsMacroEFormAction(ByVal sTaskName As String) As Boolean
'---------------------------------------------------------
' Is this task a MACRO eForm action?
' Its name starts with predefined prefix
'---------------------------------------------------------

    IsMacroEFormAction = False      ' Default value
    
    If Len(sTaskName) > mnLenEFormActionPrefix Then   ' Is name long enough to consider?
        ' Compare the name prefix with our predefined prefix
        If Left(sTaskName, mnLenEFormActionPrefix) = LCase(ms_EFORM_ACTION_PREFIX) Then
            IsMacroEFormAction = True
        End If
    End If
    
End Function

'---------------------------------------------------------
Private Function IsMacroVisitPlan(ByVal sCaption As String) As Boolean
'---------------------------------------------------------
' Is this task a MACRO Visit plan?
' Its caption starts with predefined prefix
' Check also for lock
'---------------------------------------------------------

    IsMacroVisitPlan = False        ' Default value
    If Len(sCaption) > mnLenVisitPlanPrefix Then   ' Is caption long enough to consider?
        ' Compare the caption prefix with our predefined prefix
        If Left(sCaption, mnLenVisitPlanPrefix) = ms_VISIT_PLAN_PREFIX Then
            IsMacroVisitPlan = True
        End If
    End If

End Function

'---------------------------------------------------------
Private Sub moStudyInstance_StateChange(TaskKey As String, _
                            OldState As String, NewState As String)
'---------------------------------------------------------
' Arezzo task state change
' We're only interested in certain state changes
'---------------------------------------------------------
Dim oplmTask As TaskInstance    ' The Arezzo task object

    On Error GoTo ErrLabel

    ' Patient needs saving
    mbPatientDirty = True
        
    ' Branch on the new state
    Select Case NewState
    
    Case "in_progress"  ' Only interested in plans
        Set oplmTask = moALM.GuidelineInstance.colTaskInstances.Item(TaskKey)
        If oplmTask.TaskType = "plan" Then
            Call DealWithInProgressPlan(oplmTask)
        End If
        Set oplmTask = Nothing
    Case Else
        ' Ignore other state changes
        
    End Select  ' Task state

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.StateChange"

End Sub

''---------------------------------------------------------
'Friend Sub RunEngine()
''---------------------------------------------------------
'' NCJ 14 Feb 2002 - Run the engine
'' (e.g. to nudge new eForms into existence when loading a subject)
'' NCJ 26 Nov 02 - Use new specialised routine to collect new tasks
''---------------------------------------------------------
'
'    Call moALM.GuidelineInstance.RunEngine
'
'End Sub

'---------------------------------------------------------
Private Sub GetNewAREZZOTasks()
'---------------------------------------------------------
' NCJ 26 Nov 02 - Get the new AREZZO tasks as a result of running the engine
' This REPLACES RunEngine and GenerateArezzoEvents
' NCJ 3 Mar 05 - Bug 2537 - Prolog might return results in a FILE!
'---------------------------------------------------------
Dim sR As String
Dim sTasksList As String
Dim vTasks As Variant
Dim i As Long
Dim lParentInstanceId As Long
Dim lTaskInstanceId As Long
Dim lTaskId As Long
Dim nCycleNo As Integer
Dim lStudyPlanId As Long
Dim sFileName As String
Const sRESULTS_IN_FILE = "0001"

    On Error GoTo ErrLabel
    
    sTasksList = moALM.GetPrologResult("macro_get_new_instances. ", sR)
    ' NCJ 3 Mar 05 - Bug 2537 - Check for results in a file
    If sR = sRESULTS_IN_FILE Then
        ' Pick up the output from temporary file
        sFileName = sTasksList
        sTasksList = StringFromFile(sFileName)
        Kill sFileName
    End If

    ' Each new instance is returned in form
    '   TaskID|CycleNo|ParentID|CLMTaskId|
    If sTasksList > "" Then
        vTasks = Split(sTasksList, msAREZZO_SEPARATOR)
        i = 0
        lStudyPlanId = Val(moALM.GuidelineInstance.TopLevelPlanKey)
        Do While i <= UBound(vTasks)
            If vTasks(i) = "" Then Exit Do     ' Gone past final separator
            lTaskInstanceId = Val(vTasks(i))
            nCycleNo = Val(vTasks(i + 1))
            lParentInstanceId = Val(vTasks(i + 2))
            lTaskId = Val(vTasks(i + 3))
            i = i + 4
            If lParentInstanceId = lStudyPlanId Then
                ' Parent is Study so it's a visit
'    Debug.Print "NewVisitInstance - " & lTaskId & ", " & lTaskInstanceId
                RaiseEvent NewVisitInstance(lTaskId, lTaskInstanceId, nCycleNo)
            Else
                ' Parent is a visit so it's an eForm
'    Debug.Print "NewEFormInstance - " & lTaskId & ", " & lTaskInstanceId & ", " & lParentInstanceId
                RaiseEvent NewEFormInstance(lTaskId, lTaskInstanceId, lParentInstanceId, nCycleNo)
            End If
        Loop
        ' Say that things have changed
        RaiseEvent Idle(True)
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.GetNewAREZZOTasks"

End Sub

'---------------------------------------------------------
Public Sub GenerateArezzoEvents()
'---------------------------------------------------------
' Get the latest events from the Arezzo Engine
' NCJ 26 Nov 02 - Use new specialised call to just get new eForm/Visit instances
'---------------------------------------------------------

'    moALM.GuidelineInstance.RaisePLMEvents
    Call GetNewAREZZOTasks

End Sub

'---------------------------------------------------------
Public Sub ConfirmRequestedForm(ByVal lEFormTaskId As Long)
'---------------------------------------------------------
' Confirm the completion of the specified eFormInstance
' (when all mandatory questions have been answered).
' Does nothing if eForm instance already complete.
'---------------------------------------------------------
' This confirms the action task in an eForm plan
' The lEFormTaskId is actually the task id of the plan
' Must extricate the action task from it
' and generate events after doing this
'---------------------------------------------------------
Dim sGoal As String
Dim sR As String

    On Error GoTo ErrLabel

    ' NCJ 26 Nov 02 - New faster version
    sGoal = "macro_confirm_eForm( " & lEFormTaskId & " ). "
    Call moALM.GetPrologResult(sGoal, sR)
    ' NCJ 6 May 04 - Don't call GetNewAREZZOTasks here because it takes too long
    ' (and it gets done later anyway)
'    Call GetNewAREZZOTasks

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.ConfirmRequestedForm(" & lEFormTaskId & ")"

End Sub

'---------------------------------------------------------
Public Function GetArezzoTasks() As Collection
'---------------------------------------------------------
' Get all the non-MACRO Arezzo tasks that need attention.
' Collection includes requested actions and enquiries,
' in-progress decisions and permitted decisions.
' NCJ 4 Dec 03 - Must key on TaskId, not name!
'---------------------------------------------------------
Dim colTasks As Collection
Dim oTask As TaskInstance
Dim vTaskKey As Variant

    On Error GoTo ErrLabel

    Set colTasks = New Collection
    With moALM.GuidelineInstance.colTaskInstances
    
        ' Pick up requested actions
        For Each vTaskKey In .CurrentTasks("action", "requested")
            Set oTask = .Item(CStr(vTaskKey))
            ' Is it a MACRO action?
            If IsMacroEFormAction(oTask.Name) Then
                ' Ignore it - it's dealt with elsewhere
            Else
                colTasks.Add oTask, "K" & oTask.TaskKey
            End If
        Next
        
        ' Get all the requested enquiries
        For Each vTaskKey In .CurrentTasks("enquiry", "requested")
            Set oTask = .Item(CStr(vTaskKey))
            If oTask.RequestedData.Count > 0 Then
                colTasks.Add oTask, "K" & oTask.TaskKey
            End If
        Next
        
        ' Now add in all the in_progress decisions which have requested data
        For Each vTaskKey In .CurrentTasks("decision", "in_progress")
            Set oTask = .Item(CStr(vTaskKey))
            If oTask.RequestedData.Count > 0 Then
                colTasks.Add oTask, "K" & oTask.TaskKey
            End If
        Next
        
        ' And finally add in the permitted decisions
        For Each vTaskKey In .CurrentTasks("decision", "permitted")
            Set oTask = .Item(CStr(vTaskKey))
            colTasks.Add oTask, "K" & oTask.TaskKey
        Next
    
    End With
    
    Set GetArezzoTasks = colTasks
    Set oTask = Nothing
    Set colTasks = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Arezzo_DM.GetArezzoTasks"

End Function

' NCJ 31 Jan 03 - These are routines that will carry out an AREZZO activity
' and store the activity in case we need to redo it later (after a multi-user reload)

'---------------------------------------------------------
Public Sub ConfirmAction(sTaskKey As String)
'---------------------------------------------------------
'---------------------------------------------------------
Dim oAA As ArezzoActivity

    Set oAA = New ArezzoActivity
    
    ' Set up the activity class
    oAA.ActivityType = eArezzoActivity.aaConfirmAction
    oAA.TaskKey = sTaskKey
    Call oAA.DoActivity(moALM)
    ' Add it to our collection of "done" activities
    mcolArezzoActivities.Add oAA
    
    Set oAA = Nothing

End Sub

'---------------------------------------------------------
Public Sub CommitDecision(sTaskKey As String, sCand As String)
'---------------------------------------------------------
'---------------------------------------------------------
Dim oAA As ArezzoActivity

    Set oAA = New ArezzoActivity
    
    ' Set up the activity class
    oAA.ActivityType = eArezzoActivity.aaCommitDecision
    oAA.TaskKey = sTaskKey
    oAA.ItemName = sCand
    Call oAA.DoActivity(moALM)
    ' Add it to our collection of "done" activities
    mcolArezzoActivities.Add oAA
    
    Set oAA = Nothing

End Sub

'---------------------------------------------------------
Public Sub AddNonMacroData(sDName As String, sValue As String)
'---------------------------------------------------------
'---------------------------------------------------------
Dim oAA As ArezzoActivity

    Set oAA = New ArezzoActivity
    
    ' Set up the activity class
    oAA.ActivityType = eArezzoActivity.aaAddData
    oAA.ItemName = sDName
    oAA.Value = sValue
    Call oAA.DoActivity(moALM)
    ' Add it to our collection of "done" activities
    mcolArezzoActivities.Add oAA
    
    Set oAA = Nothing

End Sub

'---------------------------------------------------------
Friend Sub SetAsSaved()
'---------------------------------------------------------
' Set Arezzo as saved, i.e. clear out our Arezzo Activities collection
'---------------------------------------------------------

    Set mcolArezzoActivities = Nothing
    Set mcolArezzoActivities = New Collection

End Sub

'---------------------------------------------------------
Friend Sub ReDoArezzoActivities()
'---------------------------------------------------------
' Redo all the Arezzo activities that happened since the last save
' This is necessary after a patient state reload
'---------------------------------------------------------
Dim oAA As ArezzoActivity

    For Each oAA In mcolArezzoActivities
        Call oAA.DoActivity(moALM)
    Next

End Sub
