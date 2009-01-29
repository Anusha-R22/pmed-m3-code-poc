Attribute VB_Name = "modArezzoReBuild"
'------------------------------------------------------------------------------'
' File:         modArezzoReBuild.bas
' Copyright:    InferMed Ltd. 2000-2006. All Rights Reserved
' Author:       Mo Morris, February 2000
' Purpose:      Contains  routines for (re)building the whole Arrezo file for
'               a study and placing it in the Protocols table.
'------------------------------------------------------------------------------'
'   Revisions:
'   Mo Morris   16/2/00     CheckProteditForTrial added.
'   Mo Morris   18/2/00     Adding facilities for a user to provide an alternative Code
'                           when a duplicated CRFPageCode, VisitCode or DataItemCode is
'                           discovered during the re-build.
'   'TA         28/03/2000  generic function for visit/question/eform validation
'                           generic function for users' changes to codes
'   TA 16/05/2000 SR3459: new trials now allowed underscore characters
'   NCJ 17-26/May/00 SR3437 Routine to update Arezzo for changed MACRO database
'           (to support direct DB updates from the GGB)
'   NCJ 17 Apr 01 - Added check for Arezzo validity for trial names
'   NCJ 17 May 01 - Changed to use new ALM4 (Arezzo Logic Module)
'   MLM 13/06/02: Rebuild protocols collection in CheckProteditForTrial.
'   MLM 14/06/02: Fixed error handler in CreateAllArezzo.
'   ZA 21/06/02 : Fixed bug 13 in build 2.2.16
' MACRO 3.0
'   NCJ 3 Jan 02 - Added Question Group code checks in ValidateItemCode
'   ZA 20/08/2002 - changed UpdateAllArezzo sub to a function, created new function UpdateArezzoStatus
'   ZA 10/09/2002 - Use enum values for Arezzo update rather than 1 or 0
'   ZA 23/09/2002 - Removed PSS references
'   NCJ 25 Sept 02 - Use SQL call to check for Group codes in ValidateItemCode
'                   Changed "Arezzo" in user messages to "AREZZO"
'   NCJ 29 Nov 02 - Make sure Protocols table is accessed by FileName; set Visit repeats as appropriate
'   NCJ 12 Mar 03 - When updating AREZZO, add in any new category codes for existing data items
'   NCJ 3 Apr 03 - Added SynchroniseVisitCycles (for Bug 870)
'   NCJ 12 Aug 03 - Do not reset cycle numbers during an AREZZO update!
'   NCJ 2 May 06 - During an update, re-save all validations for existing questions too
'   NCJ 19 Jun 06 - Added bCanEdit to CheckProtEditForTrial
'   NCJ 6 Sept 07 - Issue 2904 - Correction of typo in message for new category values
'------------------------------------------------------------------------------'

Option Explicit

'constansts for DataItemValidation
Public Const gsITEM_TYPE_VISIT = "Visit"
Public Const gsITEM_TYPE_QUESTION = "Question"
Public Const gsITEM_TYPE_EFORM = "eForm"

Private mbCodeChangesExist As Boolean
Private moCollectionOfCodesToChange As clsCodesForChanging
Private mnFileNumber As Integer
Private msFileName As String
' NCJ 17/5/00 - Store the new things we've created
Private mcolNewIds As Collection

' NCJ 12 Mar 03 - Store how many new categories were created
Private mnNewCats As Integer

Private mlClinicalTrialId As Long

' NCj 25 Sept 02 - Made assumption of VersionID = 1 explicit
' We always assume a version ID of 1 in this module!
Private Const mnVersionId As Integer = 1

' NCJ 26/5/00 - New enumeration for Update and Create
Private Enum ArezzoUpdateMode
    Create
    Update
End Enum
' Store the current mode
Private mUpdateMode As ArezzoUpdateMode

'---------------------------------------------------------------------
Public Function CheckProteditForTrial(ByVal sClinicalTrialName As String, _
            Optional bCanEdit As Boolean = True) As Boolean
'---------------------------------------------------------------------
' Check the Protocols collection for a trial with this name
' NCJ 19 Jun 06 - Added bCanEdit
'---------------------------------------------------------------------
Dim sSQL As String
Dim oRS As ADODB.Recordset

    On Error GoTo ErrHandler
        
    Set oRS = New ADODB.Recordset
    sSQL = "SELECT * FROM Protocols WHERE FileName = '" & sClinicalTrialName & "'"
        
    oRS.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic
    
    If oRS.EOF Or oRS.BOF And bCanEdit Then
        ' If it doesn't exist, try creating the Arezzo
        ' NCJ 19 Jun 06 - Only build AREZZO if they can edit
        If bCanEdit Then
            CheckProteditForTrial = CreateAllArezzo(sClinicalTrialName)
        Else
            ' No AREZZO but they don't have edit permission
            DialogInformation "This study has no AREZZO definition and cannot be opened."
            CheckProteditForTrial = False
        End If
    Else
        CheckProteditForTrial = True
    End If
    
    Set oRS = Nothing
   
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.CheckProteditForTrial"

End Function

'---------------------------------------------------------------------
Public Function UpdateArezzoStatus(ByVal lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
'ZA 20/08/2002 - calls the UpdateAllArezzo function, if ArezzoUpdateStatus
'column is set to 1 in StudyDefintion table.
'ZA 10/09/2002 - used the enumeration instead of 1 or 0
'---------------------------------------------------------------------
Dim sQuery As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
    
    UpdateArezzoStatus = False
    
    sQuery = "Select ArezzoUpdateStatus from StudyDefinition where ClinicalTrialId = " & lClinicalTrialId
    Set rsTemp = New ADODB.Recordset
    
    rsTemp.Open sQuery, MacroADODBConnection, adOpenKeyset, adLockOptimistic
    
    'ZA 10/09/2002 - replaced 1, 0 with enumeration
    If rsTemp.Fields("ArezzoUpdateStatus").Value = eArezzoUpdateStatus.auRequired Then
        If UpdateAllArezzo(lClinicalTrialId) Then
            rsTemp.Fields("ArezzoUpdateStatus").Value = eArezzoUpdateStatus.auNotRequired
            rsTemp.Update
            UpdateArezzoStatus = True
        End If
    Else
        UpdateArezzoStatus = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    ' NCJ 3 April 03 - Synchronise visit cycles from AREZZO
    Call SynchroniseVisitCycles(lClinicalTrialId)
    
    Exit Function
    
ErrHandler:
    UpdateArezzoStatus = False
    Select Case MACROErrorHandler("modArezzoRebuild", Err.Number, Err.Description, "UpdateArezzoStatus", Err.Source)
    
        Case OnErrorAction.Retry
            Resume
    End Select

End Function

'---------------------------------------------------------------------
Public Function UpdateAllArezzo(ByVal lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
' NCJ 25/5/00 - Update Arezzo with the NEW things that aren't in Arezzo already
' i.e. if new rows have arrived in the database from CG
' ZA 20/08/2002 - changed this sub to a function
'---------------------------------------------------------------------
Dim sOrigCLMFile As String
Dim sMSG As String
Dim sMsgBoxTitle As String
Dim nNewItems As Integer

    On Error GoTo ErrHandler
    
    UpdateAllArezzo = False
    
    sMsgBoxTitle = "AREZZO Study Update"
    
    'REM 10/06/02 - Changed GGB to CG
    sMSG = "This will update the AREZZO definition of this study based on "
    sMSG = sMSG & "recently imported changes from the CG." & vbNewLine
    sMSG = sMSG & "Are you sure you wish to continue?"
    If DialogQuestion(sMSG, sMsgBoxTitle) = vbNo Then
        Exit Function
    End If

    Call SetUpThingsAtTheBeginning(TrialNameFromId(lClinicalTrialId))
    
    mlClinicalTrialId = lClinicalTrialId
    
    ' We're going to update rather than create
    mUpdateMode = Update
    
    ' Save the existing CLMFile
    ' NCJ 17/5/01 - Use new ALM
    sOrigCLMFile = goALM.ArezzoFile
    
    If Not CreateAllCRFPages(lClinicalTrialId) Then
        Err.Raise vbObjectError + 10102, "Update All CRFPages Failure"
    End If
    
    If Not CreateAllVisits(lClinicalTrialId) Then
        Err.Raise vbObjectError + 10103, "Update All Visits Failure"
    End If
    
    If Not CreateAllPageVisits(lClinicalTrialId) Then
        Err.Raise vbObjectError + 10104, "Update All Page/Visits Failure"
    End If
    
    If Not CreateAllDataItems(lClinicalTrialId) Then
        Err.Raise vbObjectError + 10105, "Update All DataItems Failure"
    End If
    
    If Not CreateAllDataItemCategories(lClinicalTrialId) Then
        Err.Raise vbObjectError + 10106, "Update All DataItem Categories Failure"
    End If
    
    If Not CreateAllDataItemValidations(lClinicalTrialId) Then
        Err.Raise vbObjectError + 10107, "Update All DataItem Validations Failure"
    End If
    
    'Inserting Pages generates a DataEntry Task
    'A DataEntry Task generates a new TaskId
    'The taskId generation could prevent non-generated Id's from being inserted
    'For this reason InsertAllCRFPages is called last
    If Not InsertAllCRFPages(lClinicalTrialId) Then
        Err.Raise vbObjectError + 10108, "Update All CRFPage Insertions Failure"
    End If
    
    nNewItems = mcolNewIds.Count
    sMSG = "The AREZZO definition for this study has been successfully updated." & vbNewLine
    If nNewItems = 1 Then
        sMSG = sMSG & "1 new item was created."
    ElseIf nNewItems = 0 Then
        sMSG = sMSG & "No new visits, eForms or questions were created."
    Else
        sMSG = sMSG & nNewItems & " new items were created."
    End If
    ' NCJ 12 Mar 03 - Add any new category items added
    If mnNewCats > 0 Then
        sMSG = sMSG & vbNewLine & mnNewCats & " new category value"
        If mnNewCats = 1 Then
            sMSG = sMSG & " was added."
        Else
            ' NCJ 6 Sept 07 - Issue 2904 - Correction of typo in message
            sMSG = sMSG & "s were added."
'            sMSG = sMSG & mnNewCats & "s were added."
        End If
    End If
    Call DialogInformation(sMSG, sMsgBoxTitle)
    
    Call TidyUpAtTheEnd
    
    UpdateAllArezzo = True
    
    Exit Function
ErrHandler:
    Call TidyUpAfterError
    ' Check for one of our specially created errors
    If ((Err.Number - vbObjectError)) > 10099 And ((Err.Number - vbObjectError) < 10109) Then
        sMSG = "The update of the AREZZO definition for this study is being aborted." & vbNewLine
        sMSG = sMSG & "No changes have been made." & vbNewLine
        Call DialogInformation(sMSG, sMsgBoxTitle)
        ' Restore the original CLM file to memory
        gclmGuideline.Clear
        ' NCJ 17/5/01 - Use new ALM
        goALM.ArezzoFile = sOrigCLMFile
    Else
        Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "UpdateAllArezzo", "modArezzoReBuild")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
        End Select
    End If
    
End Function

'---------------------------------------------------------------------
Public Function CreateAllArezzo(ByVal sClinicalTrialName As String) As Boolean
'---------------------------------------------------------------------
'This function creates an Arezzo file in the Arezzo engine based on the
'contents of the database. This is achieved by mimicing the CLM calls that
'would have been made if/when the trial was initially created.
'Eventually this creates a record in Macro's Protocols table (accessed via PSS.DLL)
'Note that the raised errors (10100 to 10108) are never displayed and only
'exist as a means to exit and go down to ErrHandler

'MLM 14/06/02: Fixed error handler.
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sMSG As String
Dim sMsgBoxTitle As String

    On Error GoTo ErrHandler
    
    sMsgBoxTitle = "Create AREZZO Study Definition"
    
    sMSG = "The AREZZO definition needs to be created for study " & sClinicalTrialName & "."
    sMSG = sMSG & vbNewLine & "Click OK to continue."
    Call DialogInformation(sMSG, sMsgBoxTitle)
        
    Call SetUpThingsAtTheBeginning(sClinicalTrialName)
    
    ' We're going to create from scratch
    mUpdateMode = Create
    
    'get the ClinicaltrialId for the specified trial
    sSQL = "Select ClinicalTrialId FROM ClinicalTrial " _
        & " WHERE ClinicalTrialName = '" & sClinicalTrialName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTemp.RecordCount <> 1 Then
        'Something is wrong so exit
        Err.Raise vbObjectError + 10100, "Study Id/Name error"
    End If
    mlClinicalTrialId = rsTemp!ClinicalTrialId
    rsTemp.Close
    Set rsTemp = Nothing
    
    'check that the ClinicalTrialName is valid and unique
    If Not ValidateTrialName(sClinicalTrialName, False) Then
        Err.Raise vbObjectError + 10101, "Invalid Study Name"
    End If
    
    'call CreateProformaTrial which creates gpssProtocol in Arezzo using CLM calls and saves it
    CreateProformaTrial mlClinicalTrialId, sClinicalTrialName
    
    'create Pages by calling CreateAllCRFPages
    If Not CreateAllCRFPages(mlClinicalTrialId) Then
        Err.Raise vbObjectError + 10102, "Create All CRFPages Failure"
    End If
    
    'create Visits by calling CreateAllVisits
    If Not CreateAllVisits(mlClinicalTrialId) Then
        Err.Raise vbObjectError + 10103, "Create All Visits Failure"
    End If
    
    'create Page/Visits by calling CreateAllPageVisits
    If Not CreateAllPageVisits(mlClinicalTrialId) Then
        Err.Raise vbObjectError + 10104, "Create All Page/Visits Failure"
    End If
    
    'create DataItems by calling CreateAllDataItems
    If Not CreateAllDataItems(mlClinicalTrialId) Then
        Err.Raise vbObjectError + 10105, "Create All DataItems Failure"
    End If
    
    'create DataItem category codes by calling CreateAllDataItemCategories
    If Not CreateAllDataItemCategories(mlClinicalTrialId) Then
        Err.Raise vbObjectError + 10106, "Create All DataItem Categories Failure"
    End If
    
    'create DataItem Validation conditions by calling CreateAllDataItemValidations
    If Not CreateAllDataItemValidations(mlClinicalTrialId) Then
        Err.Raise vbObjectError + 10107, "Create All DataItem Validations Failure"
    End If
    
    'Inserting Pages generates a DataEntry Task
    'A DataEntry Task generates a new TaskId
    'The taskId generation could prevent non-generated Id's from being inserted
    'For this reason InsertAllCRFPages is called last
    If Not InsertAllCRFPages(mlClinicalTrialId) Then
        Err.Raise vbObjectError + 10108, "Insert All CRFPages Failure"
    End If

    Call TidyUpAtTheEnd
        
    'Arezzo file created successfully
    CreateAllArezzo = True

Exit Function
ErrHandler:
    
    'MLM 14/06/02 CBB 2.2.15/10: Don't call TidyUpAfterError until after testing error number, otherwise it is reset to 0.
    'Call TidyUpAfterError
    If ((Err.Number - vbObjectError)) > 10099 And ((Err.Number - vbObjectError) < 10109) Then
        Call TidyUpAfterError
        sMSG = "The creation of the AREZZO definition for study " & sClinicalTrialName
        sMSG = sMSG & " is being aborted." & vbNewLine
        sMSG = sMSG & "The study cannot be opened." & vbNewLine & "Click OK to continue."
        Call DialogInformation(sMSG, sMsgBoxTitle)
        CreateAllArezzo = False
        
        'Remove the partly created Arezzo file by calling the PSS.DLL
        'Call DeleteProformaTrial(sClinicalTrialName)
        
        gclmGuideline.Clear
    Else
        Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CreateAllArezzo", "modArezzoReBuild")
            Case OnErrorAction.Ignore
                Resume Next
            Case OnErrorAction.Retry
                Resume
            Case OnErrorAction.QuitMACRO
                Call TidyUpAfterError
                Call ExitMACRO
                Call MACROEnd
        End Select
    End If
End Function

'---------------------------------------------------------------------
Private Sub SetUpThingsAtTheBeginning(sClinicalTrialName)
'---------------------------------------------------------------------
' Set up things at the beginning before starting a Create or an Update
'---------------------------------------------------------------------
    
    Call HourglassOn
    
    'Turn the Do CLM Save flag off until the whole of the Arezzo file has been rebuilt
    gbDoCLMSave = False
    
    'initialise the flag that indicates that code changes have taken place
    mbCodeChangesExist = False
    'set up the name of the Changed Code File
    msFileName = gsTEMP_PATH & sClinicalTrialName & "_ChangedCodes.txt"
    
    ' Set up module-level collections
    Set mcolNewIds = New Collection
    Set moCollectionOfCodesToChange = New clsCodesForChanging
    mnNewCats = 0
    
End Sub

'---------------------------------------------------------------------
Private Sub TidyUpAtTheEnd()
'---------------------------------------------------------------------
' Tidy up after a successful finish (either Create or Update)
'---------------------------------------------------------------------

    If mbCodeChangesExist Then
        'close the changed codes log file
        Print #mnFileNumber,
        Print #mnFileNumber, "Check all 'Derivations' and 'Collect if' expressions for"
        Print #mnFileNumber, "any of the above OldCodes and change them to the NewCodes."
        Close #mnFileNumber
        'call ProcessCodeChangesInDatabase which will perform the code changes in the database
        ProcessCodeChangesInDatabase (mlClinicalTrialId)
        DialogInformation ("A file containing the codes that you changed has been placed at" & vbNewLine & msFileName)
    End If
    
    'Turn the Do CLM Save flag on
    gbDoCLMSave = True
    
    'save the newly created Arezzo definition file
    SaveCLMGuideline
    
    ' Clear down module-level collections
    Set mcolNewIds = Nothing
    Set moCollectionOfCodesToChange = Nothing

    Call HourglassOff

End Sub

'---------------------------------------------------------------------
Private Sub TidyUpAfterError()
'---------------------------------------------------------------------
' Tidy up after an error during either Create or Update
'---------------------------------------------------------------------

    If mbCodeChangesExist Then
        Close #mnFileNumber
        'delete the changed codes log file
        Kill msFileName
    End If
    
    ' Clear down module-level collections
    Set mcolNewIds = Nothing
    Set moCollectionOfCodesToChange = Nothing

    Call HourglassOff

End Sub

'---------------------------------------------------------------------
Public Function ValidateTrialName(sClinicalTrialName As String, _
                                Optional bCheckNameUnique As Boolean = True) As Boolean
'---------------------------------------------------------------------
'Note that checking for unique name is optional:-
'   when called from frmMenu.NewTrial a name check is required
'   when called from CreateAllArezzo the name will already exist on file
'   and a name check is not required
'---------------------------------------------------------------------
Dim sMSG As String

    On Error GoTo ErrHandler

    ValidateTrialName = False
    sMSG = ""
    
    If sClinicalTrialName = "" Then
        Exit Function
    ElseIf Not gblnValidString(sClinicalTrialName, valAlpha + valNumeric + valUnderscore) Then
        'TA 16/05/2000 SR3459: now allows underscore characters
        sMSG = "Study names can only contain alphanumeric characters."
    ElseIf Not StartsWithAlpha(sClinicalTrialName) Then
        sMSG = "Study names must start with a letter."
    ElseIf Len(sClinicalTrialName) > 15 Then
        sMSG = "Study names cannot be more than 15 characters long."
    ' NCJ 17/4/01
    ElseIf Not gblnNotAReservedWord(sClinicalTrialName) Then
        sMSG = "Study names cannot be reserved words of MACRO or AREZZO."
    ' NCJ 17/4/01
    'ZA - 21/06/2002 - commented out the following line of code, as it
    'is not needed. fix of bug 13 in build 2.2.16
    'ElseIf Not IsValidTrialName(sClinicalTrialName, sMsg) Then
        ' sMsg already set up in IsValidTrialName
    ElseIf bCheckNameUnique Then
        If gblnTrialExists(sClinicalTrialName) Then
            sMSG = "A study with this name already exists."
        Else
            ValidateTrialName = True
        End If
    Else
        ValidateTrialName = True
    End If

    If sMSG > "" Then
        Call DialogWarning(sMSG)
    End If
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ValidateTrialName", "modArezzoReBuild")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Public Function ValidateItemCode(ByVal sCode As String, ByVal sType As String, _
                            ByRef sMSG As String, bMessage As Boolean) As Boolean
'---------------------------------------------------------------------
'TA 28/03/2000: One function to replcae code checking functions
'   Input:
'       sCode - vist/question/eform code
'       sType - visit/question/eform?
'       bMessage - show validation message?
'   Output:
'       sMsg - validation message
'       function - valid code?
' NCJ 3 Jan 02 - Also check Question Group codes (assuming QuestionGroups loaded in frmMenu)
'   NCJ 25 Sept 02 - Use SQL call to check for Group codes
'---------------------------------------------------------------------

Dim bValid As Boolean

    On Error GoTo ErrHandler

    bValid = False
    sMSG = ""
    
    If sCode = "" Then
        sMSG = sType & " codes cannot be blank."
    ElseIf Not gblnValidString(sCode, valAlpha + valNumeric + valUnderscore) Then
        sMSG = sType & " codes can only contain alphanumeric characters."
    ElseIf Not gblnValidString(Left$(sCode, 1), valAlpha) Then
        sMSG = sType & " codes must start with an alphabetic character."
    ElseIf Not gblnValidString(Right$(sCode, 1), valAlpha + valNumeric) Then
        sMSG = sType & " codes must end with an alphanumeric character."
    ElseIf Not gblnNotAReservedWord(sCode) Then
        sMSG = sType & " codes cannot be reserved words of MACRO or AREZZO."
    ElseIf Len(sCode) > 15 Then
        sMSG = sType & " codes cannot be more than 15 characters long."
    ElseIf Not UniqueVisitFormDataitemCode(sCode, sMSG) Then
        ' sMsg is already set up
    ElseIf QGroupCodeExists(sCode) Then
        sMSG = "A Question Group with this code already exists"
    Else
        bValid = True
    End If

    ' Show the message if there's an error and they want to see it
    If (Not bValid) And bMessage Then
        Call DialogInformation(sMSG)
    End If
    
    ValidateItemCode = bValid

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ValidateItemCode", "modArezzoReBuild")
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
Private Function QGroupCodeExists(ByVal sCode As String) As Boolean
'---------------------------------------------------------------------
' NCJ 25 Sept 02
' Search loaded QGroups if they exist,
' otherwise assume we've got a mlClinicalTrialID and go to the database
'---------------------------------------------------------------------

    ' Look for loaded QGroups first (i.e. only if study is open)
    If Not frmMenu.QuestionGroups Is Nothing Then
        QGroupCodeExists = frmMenu.QuestionGroups.CodeExists(sCode)
    Else
        ' Go to database to check for this study (e.g. during RebuildArezzo)
        QGroupCodeExists = (QGroupExists(mlClinicalTrialId, mnVersionId, sCode) <> -1)
    End If
    
End Function

'---------------------------------------------------------------------
Public Function GetItemCode(sType As String, sPrompt As String, Optional sCode As String = "") As String
'---------------------------------------------------------------------
' TA 28/03/2000
'Input:
'   sType - visit/eForm/question/question group
'   sPrompt - Prompt for user
'Output:
'   function - Item code (empty string for cancel)
'---------------------------------------------------------------------
Dim bValid As Boolean
Dim sMSG As String

    bValid = False
    Do Until bValid
        sCode = InputBox(sPrompt, gsDIALOG_TITLE, sCode)
        If sCode = "" Then    ' if cancel, then return control to user
            bValid = True
        Else
            bValid = ValidateItemCode(sCode, sType, sMSG, True)
        End If
    Loop
    GetItemCode = sCode
    Exit Function
    
End Function

'---------------------------------------------------------------------
Private Sub NewCLMPlan(ByVal sPlanName As String, ByVal lPlanId As Long)
'---------------------------------------------------------------------
' Create a new CLM "plan" task with specified ID.
' Lock it to stop it being deleted when removed from a Visit plan
'---------------------------------------------------------------------
Dim clmtask As Task

    On Error GoTo ErrHandler
    
'    Debug.Print "NEWCLMPlan " & sPlanName & " " & lPlanId
    Set clmtask = gclmGuideline.colTasks.AddWithId(sPlanName, "plan", lPlanId)
    clmtask.Locked = True       ' Lock to prevent deletion
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.NewCLMPlan"

End Sub

'---------------------------------------------------------------------
Private Function CreateAllCRFPages(lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bNamesAllValid As Boolean
Dim sCRFPageCode As String
Dim sMSG As String
Dim lCRFPageId As Long

    On Error GoTo ErrHandler

    CreateAllCRFPages = False
    
    'create a recordset of CRFPages within the trial and create the required Arezzo code using CLM calls
    sSQL = "Select CRFPageId, CRFPageCode FROM CRFPage " _
        & " WHERE ClinicalTrialID = " & lClinicalTrialId _
        & " ORDER BY CRFPageOrder"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    bNamesAllValid = True
    Do Until rsTemp.EOF And bNamesAllValid
        sCRFPageCode = rsTemp!CRFPageCode
        ' NCJ 17/5/00 Check to see if we're doing new ones only
        If ArezzoTaskDoesNotExist(sCRFPageCode) Then
        'TA 28/03/2000  now calls generic code validate function
            If Not ValidateItemCode(sCRFPageCode, gsITEM_TYPE_EFORM, sMSG, True) Then
                If Not UserChangesCode(gsITEM_TYPE_EFORM, sCRFPageCode, rsTemp!CRFPageId) Then
                    bNamesAllValid = False
                    Exit Do
                End If
            End If
            lCRFPageId = rsTemp!CRFPageId
            Call NewCLMPlan(gsCLMCRFName(sCRFPageCode), lCRFPageId)
            ' If we're doing an update, store that this CRFPage is new
            If mUpdateMode = Update Then
                mcolNewIds.Add lCRFPageId, "K" & lCRFPageId
            End If
            'Note that InsertProformaCRFPage is called from within InsertCRFPages
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    
    If bNamesAllValid Then
        CreateAllCRFPages = True
    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.CreateAllCRFPages"

End Function

'---------------------------------------------------------------------
Private Function CreateAllVisits(lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bNamesAllValid As Boolean
Dim sVisitCode As String
Dim sMSG As String
Dim lVisitId As Long

    On Error GoTo ErrHandler

    CreateAllVisits = False
    
    ' Create a recordset of Visits within the trial and create the required Arezzo code using CLM calls
    ' NCJ 29 Nov 02 - Include Repeating
    sSQL = "Select VisitId, VisitCode, VisitOrder, Repeating FROM StudyVisit " _
        & " WHERE ClinicalTrialID = " & lClinicalTrialId _
        & " ORDER BY VisitOrder"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    bNamesAllValid = True
    Do Until rsTemp.EOF And bNamesAllValid
        sVisitCode = rsTemp!VisitCode
        lVisitId = rsTemp!VisitId
        ' NCJ 17/5/00 Check to see if we're doing new ones only
        If ArezzoTaskDoesNotExist(sVisitCode) Then
            'TA 28/03/2000  now calls generic code validate function
            If Not ValidateItemCode(sVisitCode, gsITEM_TYPE_VISIT, sMSG, True) Then
                If Not UserChangesCode(gsITEM_TYPE_VISIT, sVisitCode, lVisitId) Then
                    bNamesAllValid = False
                    Exit Do
                End If
            End If
            Call NewCLMPlan(gsCLMVisitName(sVisitCode), lVisitId)
            Call InsertProformaVisit(lClinicalTrialId, lVisitId, rsTemp!VisitOrder)
            ' NCJ 29 Nov 02 - Set visit repeats if appropriate
            If Not IsNull(rsTemp!Repeating) Then
                Call SetVisitRepeats(lVisitId, rsTemp!Repeating)
            End If
            ' If we're doing an update, store that this Visit is new
            If mUpdateMode = Update Then
                mcolNewIds.Add lVisitId, "K" & lVisitId
            End If
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    
    If bNamesAllValid Then
        CreateAllVisits = True
    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.CreateAllVisits"
    
End Function

'---------------------------------------------------------------------
Private Function CreateAllPageVisits(lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
' Create all the eForms within the visits
' NCJ 12 Aug 03 - Do not reset cycle numbers during an update!
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim lVisitId As Long
Dim lCRFPageId As Long

    On Error GoTo ErrHandler

    CreateAllPageVisits = False
    
    'create a recordset of Page/Visits within the trial and create the required Arezzo code using CLM calls
    sSQL = "Select StudyVisitCRFPage.VisitId, StudyVisitCRFPage.CRFPageId, " _
        & " StudyVisitCRFPage.Repeating, CRFPage.CRFPageOrder FROM StudyVisitCRFPage, CRFPAge" _
        & " WHERE StudyVisitCRFPage.ClinicalTrialId = CRFPAge.ClinicalTrialId " _
        & " AND StudyVisitCRFPage.CRFPageId = CRFPAge.CRFPageId" _
        & " AND StudyVisitCRFPage.ClinicalTrialID = " & lClinicalTrialId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    Do Until rsTemp.EOF
        lVisitId = rsTemp!VisitId
        lCRFPageId = rsTemp!CRFPageId
        If ArezzoPageVisitAlreadyExists(lClinicalTrialId, lCRFPageId, lVisitId) Then
            ' Do nothing if Arezzo already has this one
        Else
            Call InsertProformaStudyVisitCRFPage(lClinicalTrialId, lVisitId, lCRFPageId, rsTemp!CRFPageOrder)
            ' Reset the cycling parameter
            ' NCJ 12 Aug 03 - ONLY do this for newly inserted eForms
            Call SetCyclingTask(lVisitId, lCRFPageId, (rsTemp!Repeating = 1))
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    
    CreateAllPageVisits = True

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.CreateAllPageVisits"
    
End Function

'---------------------------------------------------------------------
Private Function ArezzoPageVisitAlreadyExists(lClinicalTrialId As Long, _
                            lCRFPageId As Long, _
                            lVisitId As Long) As Boolean
'---------------------------------------------------------------------
' NCJ 25/5/00
' Returns TRUE if CRF page already exists in the Visit in Arezzo
' (so we don't need to add it again)
'---------------------------------------------------------------------
Dim bExists As Boolean
Dim oVisitTask As Task
Dim vTaskKey As Variant

    On Error GoTo ErrHandler
    
    ' Set to FALSE as default value
    ArezzoPageVisitAlreadyExists = False
    
    ' Assume it doesn't exist if we're creating from new
    If mUpdateMode = Create Then Exit Function
    
    ' If either CRFPageId or VisitId is new, assume page visit doesn't exist
    If ArezzoIsNewThing(lCRFPageId) Then Exit Function
    If ArezzoIsNewThing(lVisitId) Then Exit Function
    
    ' Neither Page nor Visit is new to Arezzo,
    ' but does the Page already exist as component of Visit?
    ' Get the visit plan
    Set oVisitTask = gclmGuideline.colTasks.Item(CStr(lVisitId))
    For Each vTaskKey In oVisitTask.Components
        If CLng(vTaskKey) = lCRFPageId Then
            ' It is one of the components
            ArezzoPageVisitAlreadyExists = True
            Exit For
        End If
    Next
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.CreateAllPageVisits"
    
End Function

'---------------------------------------------------------------------
Private Function CreateAllDataItems(lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bNamesAllValid As Boolean
Dim sMSG As String
'Note that sMsg is only used as an argument when calling ValidateDataItemCode
Dim sDataItemCode As String
Dim lDataItemId As Long

    On Error GoTo ErrHandler

    CreateAllDataItems = False
    
    'create a recordset of DataItems within the trial and create the required Arezzo code using CLM calls
    sSQL = "Select DataItemId, DataItemCode, DataItemName, DataType, Derivation,UnitOfMeasurement  " _
        & " FROM DataItem " _
        & " WHERE ClinicalTrialID = " & lClinicalTrialId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    bNamesAllValid = True
    Do Until rsTemp.EOF And bNamesAllValid
        sDataItemCode = rsTemp!DataItemCode
        ' NCJ 17/5/00 Check to see if we're doing new ones only
        If ArezzoDataItemDoesNotExist(sDataItemCode) Then
            'TA 28/03/2000  now calls generic code validate function
            If Not ValidateItemCode(sDataItemCode, gsITEM_TYPE_QUESTION, sMSG, False) Then
                If Not UserChangesCode(gsITEM_TYPE_QUESTION, sDataItemCode, rsTemp!DataItemId) Then
                    bNamesAllValid = False
                    Exit Do
                End If
            End If
            lDataItemId = rsTemp!DataItemId
            Call NewCLMDataItem(sDataItemCode, lDataItemId)
            Call UpdateProformaDataItem(lDataItemId, sDataItemCode, rsTemp!DataItemName, _
                    rsTemp!DataType, RemoveNull(rsTemp!Derivation), RemoveNull(rsTemp!UnitOfMeasurement))
            ' NCJ 26/5/00 If we're doing an update, store that this Data Item is new
            If mUpdateMode = Update Then
                mcolNewIds.Add lDataItemId, "K" & lDataItemId
            End If
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    
    If bNamesAllValid Then
        CreateAllDataItems = True
    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.CreateAllDataItems"

End Function

'---------------------------------------------------------------------
Private Function CreateAllDataItemCategories(lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
' Create all the categories for new data items
' NCJ 12 Mar 03 - Also add in new category items for existing data items
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsCatQuestions As ADODB.Recordset
Dim rsValueData As ADODB.Recordset
Dim colCodes As Collection
Dim lDataItemId As Long
Dim bNewDataItem As Boolean

    On Error GoTo ErrHandler

    CreateAllDataItemCategories = False
    
    'create a recordset of DataItems within the trial that are of type category
    'note that datatype 1 is category
    sSQL = "Select DataItemId FROM DataItem " _
        & " WHERE ClinicalTrialID = " & lClinicalTrialId _
        & " AND DataType = " & DataType.Category
    Set rsCatQuestions = New ADODB.Recordset
    rsCatQuestions.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    'Loop through the dataitems that are of type category
    Do Until rsCatQuestions.EOF
        lDataItemId = rsCatQuestions!DataItemId
        ' Is it a new one created this time round?
        bNewDataItem = ArezzoIsNewThing(lDataItemId)
        If Not bNewDataItem Then
            ' Get the codes that AREZZO already knows about for this existing data item
            Set colCodes = CategoryCodes(lDataItemId)
        End If
        'create a recordset of the DB ValueData/Category codes for current dataitem
        sSQL = "Select ValueCode FROM ValueData " _
            & " WHERE ClinicalTrialId = " & lClinicalTrialId _
            & " AND DataItemId = " & lDataItemId
        Set rsValueData = New ADODB.Recordset
        rsValueData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        Do Until rsValueData.EOF
            'CONSIDER VALIDATION CHECKS ON THE VALUECODES OR LET THE ENGINE REJECT THEM
            ' NCJ 26/5/00 Check to see if we're doing new ones only
            If bNewDataItem Then
                Call SaveProformaRangeValue(lDataItemId, rsValueData!ValueCode, "", True)
            Else
                ' NCJ 12 Mar 03 - Add it if it's a new category for an existing data item
                If Not CollectionMember(colCodes, LCase(rsValueData!ValueCode), False) Then
                    Call SaveProformaRangeValue(lDataItemId, rsValueData!ValueCode, "", True)
                    ' Add to count of new categories
                    mnNewCats = mnNewCats + 1
                End If
            End If
            rsValueData.MoveNext
        Loop
        rsValueData.Close
        Set rsValueData = Nothing
        rsCatQuestions.MoveNext
    Loop
    rsCatQuestions.Close
    Set rsCatQuestions = Nothing
    
    Set colCodes = Nothing
    
    CreateAllDataItemCategories = True

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.CreateAllDataItemCategories"

End Function

'---------------------------------------------------------------------
Private Function CategoryCodes(lDataItemId As Long) As Collection
'---------------------------------------------------------------------
' NCJ 12 Mar 03
' Return a keyed collection of category codes for this data item
' (so we can check for new ones in CreateAllDataItemCategories)
'---------------------------------------------------------------------
Dim oDItem As DataItem
Dim colCodes As Collection
Dim vCode As Variant

    Set colCodes = New Collection
    Set oDItem = gclmGuideline.colDataItems.Item(CStr(lDataItemId))
    For Each vCode In oDItem.RangeValues
        colCodes.Add CStr(vCode), LCase(CStr(vCode))
    Next
    
    Set CategoryCodes = colCodes
    Set colCodes = Nothing
    Set oDItem = Nothing

End Function

'---------------------------------------------------------------------
Private Function CreateAllDataItemValidations(lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
' Create all data item validations
' NCJ 2 May 06 - Re-save all validations for existing questions too
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim lDataItemId As Long
Dim colFlags As Collection
Dim colWarns As Collection

    On Error GoTo ErrHandler

    CreateAllDataItemValidations = False
    
    ' Create a recordset of DataItems with Validation criteria and create the required Arezzo code using CLM calls
    sSQL = "Select DataItemId, ValidationID,DataItemValidation FROM DataItemValidation " _
        & " WHERE ClinicalTrialID = " & lClinicalTrialId _
        & " ORDER BY DataItemId "
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    lDataItemId = 0
    
    Do Until rsTemp.EOF
'        If ArezzoIsNewThing(rsTemp!DataItemId) Then
'            Call SaveProformaWarningCondition(rsTemp!DataItemId, rsTemp!ValidationID, rsTemp!DataItemValidation)
'        End If
        ' NCJ 2 May 06 - For each data item, remove all its Warning conditions first
        ' and then just add them all again
        If rsTemp!DataItemId <> lDataItemId Then
            ' We're on to the next data item - save the previous Warnings
            If lDataItemId > 0 Then
                ' Add the new ones
                Call SaveProformaWarningConditions(lDataItemId, colFlags, colWarns)
            End If
            ' Change to the next data item
            lDataItemId = rsTemp!DataItemId
            ' Zap the old ones for this data item
            Call DeleteProformaWarningConditions(lDataItemId)
            Set colFlags = New Collection
            Set colWarns = New Collection
        End If
        ' Add the flag and condition to our collections
        colFlags.Add CStr(rsTemp!ValidationID)
        colWarns.Add CStr(rsTemp!DataItemValidation)
        
        rsTemp.MoveNext
    Loop
    ' Save the final set of validations
    If lDataItemId > 0 Then
        Call SaveProformaWarningConditions(lDataItemId, colFlags, colWarns)
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    Set colFlags = Nothing
    Set colWarns = Nothing
    
    CreateAllDataItemValidations = True

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.CreateAllDataItemValidations"

End Function

'---------------------------------------------------------------------
Private Function ArezzoIsNewThing(lThingID As Long) As Boolean
'---------------------------------------------------------------------
' See if this ID represents a new thing we've just created
' lThingID may be lVisitId, lCRFPageId or lDataItemId
' Returns TRUE if thing is newly created in this run
'---------------------------------------------------------------------
Dim lTempID As Long

    ' Assume new if in Create mode
    If mUpdateMode = Create Then
        ArezzoIsNewThing = True
    Else
        ' See if it's in the NewIDs collection
        On Error Resume Next
        lTempID = mcolNewIds.Item("K" & lThingID)
        ' Error = 0 means it was in the "NewIds" collection so it's new
        If Err.Number = 0 Then
            ArezzoIsNewThing = True
        Else
            ArezzoIsNewThing = False
        End If
    End If
    
End Function

'---------------------------------------------------------------------
Private Function InsertAllCRFPages(lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
' Create the data entry tasks for all the CRF Pages just created
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    InsertAllCRFPages = False
    
    'create a recordset of CRFPages within the trial and create the required Arezzo code using CLM calls
    sSQL = "Select CRFPageId, CRFPageCode FROM CRFPage " _
        & " WHERE ClinicalTrialID = " & lClinicalTrialId _
        & " ORDER BY CRFPageOrder"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    Do Until rsTemp.EOF
        ' NCJ 25/5/00 Check whether we're in incremental update mode
        If ArezzoIsNewThing(rsTemp!CRFPageId) Then
            If mbCodeChangesExist Then
                'If code changes have taken place check the changed code collection
                Call InsertProformaCRFPage(lClinicalTrialId, rsTemp!CRFPageId, _
                    moCollectionOfCodesToChange.ChangedCode(rsTemp!CRFPageCode, rsTemp!CRFPageId))
            Else
                Call InsertProformaCRFPage(lClinicalTrialId, rsTemp!CRFPageId, rsTemp!CRFPageCode)
            End If
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    
    InsertAllCRFPages = True

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.InsertAllCRFPages"

End Function


'---------------------------------------------------------------------
Private Sub NewCLMDataItem(ByVal sDItemCode As String, ByVal lDataItemId As Long)
'---------------------------------------------------------------------
' Create a new Arezzo data item with a given name and ID
' Default to type "text"
'---------------------------------------------------------------------
Dim clmDItem As DataItem

    On Error GoTo ErrHandler
    
'    Debug.Print "NEWDATITEM " & sDItemCode & " " & lDataItemId
    Set clmDItem = gclmGuideline.colDataItems.AddWithId(sDItemCode, "text", lDataItemId)
    clmDItem.Locked = True
    
    SaveCLMGuideline
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.NewCLMDataItem"

End Sub

'---------------------------------------------------------------------
Private Function UserChangesCode(sType As String, sCode As String, ByVal lId As Long) As Boolean
'---------------------------------------------------------------------
' TA 28/03/2000 combined version for visit/eform/question
'Input:
'   sType - question/eform/vist
'   sCode - item code
'   lId - item id
'Output:
'   function - changed?
'---------------------------------------------------------------------
Dim sOldCode As String
Dim sTypeOld As String

    On Error GoTo ErrHandler
    
    Select Case sType
    Case gsITEM_TYPE_VISIT: sTypeOld = "Visit"
    Case gsITEM_TYPE_QUESTION: sTypeOld = "DataItem"
    Case gsITEM_TYPE_EFORM: sTypeOld = "Form"
    End Select
    sOldCode = sCode
    sCode = GetItemCode(sType, "Change " & sType & " code or click CANCEL to abort: ", sOldCode)
    If sCode = "" Then
        UserChangesCode = False
        Exit Function
    End If
    Call StoreCodesForChanging(sOldCode, sCode, sTypeOld, lId)
    UserChangesCode = True
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.UserChangesCode"
   
End Function

'---------------------------------------------------------------------
Private Sub StoreCodesForChanging(sOldCode As String, sNewCode As String, _
                            sCodeType As String, lCodeId As Long)
'---------------------------------------------------------------------
Dim oCodeToChange As clsCodeForChanging

    On Error GoTo ErrHandler

    'If this is the first call to StoreCodesForChanging
    If Not mbCodeChangesExist Then
        mbCodeChangesExist = True
        'Open the file that logs the codes that have been changed
        mnFileNumber = FreeFile
        Open msFileName For Output As #mnFileNumber
        Print #mnFileNumber, "This file logs the codes that were changed during a session of AREZZO re-build"
        Print #mnFileNumber,    'blank line
    End If
    
    'add to the collection of changed codes
    moCollectionOfCodesToChange.Add sOldCode, sNewCode, sCodeType, lCodeId
    'write to the changed codes log file
    Print #mnFileNumber, "CodeType: " & sCodeType & vbTab & "OldCode: " & sOldCode _
        & vbTab & "NewCode: " & sNewCode & vbTab & "CodeId: " & lCodeId
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.StoreCodesForChanging"

End Sub

'---------------------------------------------------------------------
Private Sub ProcessCodeChangesInDatabase(lClinicalTrialId As Long)
'---------------------------------------------------------------------
Dim oCodeToChange As clsCodeForChanging
Dim sSQL As String

    On Error GoTo ErrHandler

    For Each oCodeToChange In moCollectionOfCodesToChange
        Select Case oCodeToChange.CodeType
        Case "DataItem"
            sSQL = "UPDATE DataItem SET" _
                & " DataItemCode = '" & oCodeToChange.NewCode & "'" _
                & " WHERE ClinicalTrialId = " & lClinicalTrialId _
                & " AND VersionId = " & mnVersionId _
                & " AND DataItemId = " & oCodeToChange.CodeId
        Case "Form"
            sSQL = "UPDATE CRFPage SET" _
                & " CRFPageCode = '" & oCodeToChange.NewCode & "'" _
                & " WHERE ClinicalTrialId = " & lClinicalTrialId _
                & " AND VersionId = " & mnVersionId _
                & " AND CRFPageId = " & oCodeToChange.CodeId
        Case "Visit"
            sSQL = "UPDATE StudyVisit SET" _
                & " VisitCode = '" & oCodeToChange.NewCode & "'" _
                & " WHERE ClinicalTrialId = " & lClinicalTrialId _
                & " AND VersionId = " & mnVersionId _
                & " AND VisitId = " & oCodeToChange.CodeId
        End Select
        MacroADODBConnection.Execute sSQL
    Next

    'clear down the collection of codes to change
    Set moCollectionOfCodesToChange = Nothing

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.ProcessCodeChangesInDatabase"

End Sub

'---------------------------------------------------------------------
Private Function ArezzoDataItemDoesNotExist(sCode As String) As Boolean
'---------------------------------------------------------------------
' NCJ 17/5/00
' Returns TRUE if Arezzo data item does not exist with given code
'---------------------------------------------------------------------
Dim sKey As String

    On Error GoTo ErrHandler
    
    ' Always assume it doesn't exist if Creating
    If mUpdateMode = Create Then
        ArezzoDataItemDoesNotExist = True
    Else
        ' Get the key for this data item code
        ' Key will be empty if data item does not exist
        sKey = gclmGuideline.colDataItems.GetDataItemKey(sCode)
        ArezzoDataItemDoesNotExist = (sKey = "")
    End If
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.ArezzoDataItemDoesNotExist"

End Function

'---------------------------------------------------------------------
Private Function ArezzoTaskDoesNotExist(sCode As String) As Boolean
'---------------------------------------------------------------------
' NCJ 17/5/00
' Returns TRUE if Arezzo task does not exist with given code
'---------------------------------------------------------------------
Dim sKey As String

    On Error GoTo ErrHandler
    
    ' Always assume it doesn't exist if Creating
    If mUpdateMode = Create Then
        ArezzoTaskDoesNotExist = True
    Else
        ' Get the key for this task code
        ' Key will be empty if task does not exist
        sKey = gclmGuideline.colTasks.GetTaskKey(sCode)
        ArezzoTaskDoesNotExist = (sKey = "")
    End If
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.ArezzoTaskDoesNotExist"

End Function

'---------------------------------------------------------------------
Public Function SynchroniseVisitCycles(lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
' NCJ 3 Apr 03
' Synchronise the MACRO visit cycles with those stored in AREZZO
' NB The AREZZO values are the "real" ones
' Returns TRUE if anything was changed, or FALSE if no changes made
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsVisits As ADODB.Recordset
Dim bUpdate As Boolean
Dim nCycles As Integer
Dim bChanged As Boolean

    bChanged = False
    
    sSQL = "SELECT VisitId, Repeating FROM StudyVisit WHERE ClinicalTrialId = " & lClinicalTrialId
    Set rsVisits = New ADODB.Recordset
    rsVisits.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    Do Until rsVisits.EOF
        ' Get this visit's AREZZO visit cycles
        nCycles = GetVisitRepeats(rsVisits.Fields("VisitId").Value)
        If Not IsNull(rsVisits.Fields("Repeating")) Then
            ' See if it's the same as stored in AREZZO
            bUpdate = (nCycles <> rsVisits.Fields("Repeating").Value)
        Else
            ' Repeating is NULL - only change it if cycles is not the default
            bUpdate = (nCycles <> 1)
        End If
        ' Update if necessary
        If bUpdate Then
            rsVisits.Fields("Repeating").Value = nCycles
            rsVisits.Update
            bChanged = True
        End If
        rsVisits.MoveNext
    Loop
    
    Call rsVisits.Close
    Set rsVisits = Nothing

    ' Return if any changes were made
    SynchroniseVisitCycles = bChanged
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modArezzoRebuild.SynchroniseVisitCycles"

End Function

