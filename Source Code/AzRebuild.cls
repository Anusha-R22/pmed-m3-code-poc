VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AzRebuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'----------------------------------------------------------------------------------------
'   File:       AzRebuild.cls
'   Copyright:  InferMed Ltd. 2006-2007. All Rights Reserved
'   Author:     Nicky Johns, May 2006
'   Purpose:    AREZZO Rebuild as a DLL
'----------------------------------------------------------------------------------------
'Revisions:
'   NCJ 8/9 May 06 - Initial development based on AzRebuild.bas from SD
'   NCJ 26 Nov 07 - Added DoCategoryUpdates for MACRO 3.0 API
'----------------------------------------------------------------------------------------

Option Explicit

'constants for DataItemValidation
Private Const gsITEM_TYPE_VISIT = "Visit"
Private Const gsITEM_TYPE_QUESTION = "Question"
Private Const gsITEM_TYPE_EFORM = "eForm"

Private mbCodeChangesExist As Boolean
Private moCollectionOfCodesToChange As clsCodesForChanging
'Private mnFileNumber As Integer
'Private msFileName As String
' NCJ 17/5/00 - Store the new things we've created
Private mcolNewIds As Collection

' NCJ 12 Mar 03 - Store how many new categories were created
Private mnNewCats As Integer

Private mlClinicalTrialId As Long
Private msClinicalTrialName As String

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

Private mlErrNo As Long
Private msErrDesc As String

'---------------------------------------------------------------------
Public Function DoCategoryUpdates(ByVal sDBCon As String, ByVal lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
' NCJ 26 Nov 07 - Added for MACRO 3.0 API
' Update all the categories for this trial
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
        
    mlClinicalTrialId = lClinicalTrialId
    ' Initialise the database connection
    Call InitializeMacroADODBConnection(sDBCon)
    Call InitialiseSettingsFile
    
    msClinicalTrialName = TrialNameFromId(lClinicalTrialId)
    
    ' Start the CLM
    Call StartUpCLM(lClinicalTrialId)
    
    Call LoadProformaTrial(msClinicalTrialName)
    DoCategoryUpdates = UpdateAllCategories(lClinicalTrialId)

    ' Tidy up before we leave
    Call CloseProformaTrial
    Call ShutDownCLM
    
    Call TerminateMacroADODBConnection
    
Exit Function

ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.DoCategoryUpdates"

End Function

'---------------------------------------------------------------------
Public Function DoAREZZOUpdates(ByVal sDBCon As String, ByVal lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
' Do all the necessary AREZZO updates for this trial
' If there's no CLM file, create one from scratch
' If there is a CLM file, check for updates
' sDBCon is the database connection string
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
        
    mlClinicalTrialId = lClinicalTrialId
    ' Initialise the database connection
    Call InitializeMacroADODBConnection(sDBCon)
    Call InitialiseSettingsFile
    
    msClinicalTrialName = TrialNameFromId(lClinicalTrialId)
    
    ' Start the CLM
    Call StartUpCLM(lClinicalTrialId)
    
    If CLMExists Then
        ' CLM file exists - check for updates
        Call LoadProformaTrial(msClinicalTrialName)
        DoAREZZOUpdates = UpdateAllArezzo(lClinicalTrialId)
    Else
        ' CLM file does not exist - create from scratch
        DoAREZZOUpdates = CreateAllArezzo
    End If

    ' Reset flag to say that updates have been done
    Call ResetArezzoStatus(lClinicalTrialId)

    ' Tidy up before we leave
    Call CloseProformaTrial
    Call ShutDownCLM
    
    Call TerminateMacroADODBConnection
    
Exit Function

ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.DoAREZZOUpdates"

End Function

'---------------------------------------------------------------------
Private Function CLMExists() As Boolean
'---------------------------------------------------------------------
' Check the Protocols collection for a trial with this name
' TRUE if CLM file exists, FALSE if not
'---------------------------------------------------------------------
Dim sSQL As String
Dim oRS As ADODB.Recordset

    On Error GoTo ErrHandler
        
    Set oRS = New ADODB.Recordset
    sSQL = "SELECT * FROM Protocols WHERE FileName = '" & msClinicalTrialName & "'"
        
    oRS.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic
    
    CLMExists = (oRS.RecordCount = 1)
    
    oRS.Close
    Set oRS = Nothing
   
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.CLMExists"

End Function

'---------------------------------------------------------------------
Private Sub ResetArezzoStatus(ByVal lClinicalTrialId As Long)
'---------------------------------------------------------------------
' Reset AREZZO status to "update not required"
'---------------------------------------------------------------------
Dim sQuery As String

    On Error GoTo ErrHandler
    
    sQuery = "UPDATE STUDYDEFINITION SET ArezzoUpdateStatus = 0 " _
            & " WHERE ClinicalTrialId = " & lClinicalTrialId
    
    MacroADODBConnection.Execute sQuery
            
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.ResetArezzoStatus"

End Sub

'---------------------------------------------------------------------
Private Function UpdateAllCategories(lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
' NCJ 26 Nov 07 - Added for MACRO 3.0 API
' Update only the data item categories in this file
'---------------------------------------------------------------------
Dim sOrigCLMFile As String

    On Error GoTo ErrHandler
    
    UpdateAllCategories = False
    
    Call SetUpThingsAtTheBeginning
    
    ' We're going to update rather than create
    mUpdateMode = Update
    
    ' Save the existing CLMFile
    ' NCJ 17/5/01 - Use new ALM
    sOrigCLMFile = goALM.ArezzoFile
    
    If Not CreateAllDataItemCategories(lClinicalTrialId) Then
        Err.Raise vbObjectError + 10106, "Update All DataItem Categories Failure"
    End If
    
    Call TidyUpAtTheEnd
    
    UpdateAllCategories = True
    
    Exit Function
ErrHandler:
    Call TidyUpAfterError
    ' Check for one of our specially created errors
    If ((Err.Number - vbObjectError)) > 10099 And ((Err.Number - vbObjectError) < 10109) Then
        ' Restore the original CLM file to memory
        gclmGuideline.Clear
        ' NCJ 17/5/01 - Use new ALM
        goALM.ArezzoFile = sOrigCLMFile
    Else
        Err.Raise Err.Number, , Err.Description & "|AzRebuild.UpdateAllCategories"
    End If
    
End Function

'---------------------------------------------------------------------
Private Function UpdateAllArezzo(lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
' NCJ 25/5/00 - Update Arezzo with the NEW things that aren't in Arezzo already
' i.e. if new rows have arrived in the database from CG
' ZA 20/08/2002 - changed this sub to a function
'---------------------------------------------------------------------
Dim sOrigCLMFile As String
Dim sMsg As String
Dim sMsgBoxTitle As String
Dim nNewItems As Integer

    On Error GoTo ErrHandler
    
    UpdateAllArezzo = False
    
    Call SetUpThingsAtTheBeginning
    
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
'    sMSG = "The AREZZO definition for this study has been successfully updated." & vbNewLine
'    If nNewItems = 1 Then
'        sMSG = sMSG & "1 new item was created."
'    ElseIf nNewItems = 0 Then
'        sMSG = sMSG & "No new visits, eForms or questions were created."
'    Else
'        sMSG = sMSG & nNewItems & " new items were created."
'    End If
'    ' NCJ 12 Mar 03 - Add any new category items added
'    If mnNewCats > 0 Then
'        sMSG = sMSG & vbNewLine & mnNewCats & " new category value"
'        If mnNewCats = 1 Then
'            sMSG = sMSG & " was added."
'        Else
'            sMSG = sMSG & mnNewCats & "s were added."
'        End If
'    End If
'    Call DialogInformation(sMSG, sMsgBoxTitle)
    
    Call TidyUpAtTheEnd
    
    UpdateAllArezzo = True
    
    Exit Function
ErrHandler:
    Call TidyUpAfterError
    ' Check for one of our specially created errors
    If ((Err.Number - vbObjectError)) > 10099 And ((Err.Number - vbObjectError) < 10109) Then
        ' Restore the original CLM file to memory
        gclmGuideline.Clear
        ' NCJ 17/5/01 - Use new ALM
        goALM.ArezzoFile = sOrigCLMFile
    Else
        Err.Raise Err.Number, , Err.Description & "|AzRebuild.NewCLMPlan"
    End If
    
End Function

'---------------------------------------------------------------------
Private Function CreateAllArezzo() As Boolean
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
Dim sMsg As String
Dim sMsgBoxTitle As String

    On Error GoTo ErrHandler
            
    Call SetUpThingsAtTheBeginning
    
    ' We're going to create from scratch
    mUpdateMode = Create
            
    ' Create Proforma Trial
    Call CreateProformaTrial(msClinicalTrialName)
    
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
    mlErrNo = Err.Number
    msErrDesc = Err.Description
    'MLM 14/06/02 CBB 2.2.15/10: Don't call TidyUpAfterError until after testing error number, otherwise it is reset to 0.
    'Call TidyUpAfterError
    If ((Err.Number - vbObjectError)) > 10099 And ((Err.Number - vbObjectError) < 10109) Then
        Call TidyUpAfterError
        CreateAllArezzo = False
        'Remove the partly created Arezzo file
        gclmGuideline.Clear
    Else
        Err.Raise mlErrNo, , msErrDesc & "|AzRebuild.CreateAllArezzo"
    End If
End Function

'---------------------------------------------------------------------
Private Sub SetUpThingsAtTheBeginning()
'---------------------------------------------------------------------
' Set up things at the beginning before starting a Create or an Update
'---------------------------------------------------------------------
        
    'Turn the Do CLM Save flag off until the whole of the Arezzo file has been rebuilt
    gbDoCLMSave = False
    
    'initialise the flag that indicates that code changes have taken place
    mbCodeChangesExist = False
'    'set up the name of the Changed Code File
'    msFileName = goALM.TempDirectory & msClinicalTrialName & "_ChangedCodes.txt"
    
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
        ProcessCodeChangesInDatabase (mlClinicalTrialId)
    End If
    
    'Turn the Do CLM Save flag on
    gbDoCLMSave = True
    
    'save the newly created Arezzo definition file
    SaveCLMGuideline
    
    ' Clear down module-level collections
    Set mcolNewIds = Nothing
    Set moCollectionOfCodesToChange = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub TidyUpAfterError()
'---------------------------------------------------------------------
' Tidy up after an error during either Create or Update
'---------------------------------------------------------------------

    ' Clear down module-level collections
    Set mcolNewIds = Nothing
    Set moCollectionOfCodesToChange = Nothing

End Sub

'---------------------------------------------------------------------
Public Function ValidateItemCode(ByVal sCode As String, ByVal sType As String, _
                            ByRef sMsg As String, bMessage As Boolean) As Boolean
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
    sMsg = ""
    
    If sCode = "" Then
        sMsg = sType & " codes cannot be blank."
    ElseIf Len(sCode) > 15 Then
        sMsg = sType & " codes cannot be more than 15 characters long."
    ElseIf Not UniqueVisitFormDataitemCode(sCode, sMsg) Then
        ' sMsg is already set up
'    ElseIf QGroupCodeExists(sCode) Then
'        sMsg = "A Question Group with this code already exists"
    Else
        bValid = True
    End If
    
    ValidateItemCode = bValid

Exit Function
ErrHandler:
      Err.Raise Err.Number, , Err.Description & "|AzRebuild.ValidateItemCode"

End Function

'---------------------------------------------------------------------
Private Function QGroupCodeExists(ByVal sCode As String) As Boolean
'---------------------------------------------------------------------
' NCJ 25 Sept 02
' Search loaded QGroups if they exist,
' otherwise assume we've got a mlClinicalTrialID and go to the database
'---------------------------------------------------------------------

    ' Go to database to check for this study (e.g. during RebuildArezzo)
'    QGroupCodeExists = (QGroupExists(mlClinicalTrialId, mnVersionId, sCode) <> -1)

    
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
Dim sMsg As String

    GetItemCode = ""
'    bValid = False
'    Do Until bValid
'        sCode = InputBox(sPrompt, gsDIALOG_TITLE, sCode)
'        If sCode = "" Then    ' if cancel, then return control to user
'            bValid = True
'        Else
'            bValid = ValidateItemCode(sCode, sType, sMsg, True)
'        End If
'    Loop
'    GetItemCode = sCode
'    Exit Function
    
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
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.NewCLMPlan"

End Sub

'---------------------------------------------------------------------
Private Function CreateAllCRFPages(lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bNamesAllValid As Boolean
Dim sCRFPageCode As String
Dim sMsg As String
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
            If Not ValidateItemCode(sCRFPageCode, gsITEM_TYPE_EFORM, sMsg, True) Then
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
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.CreateAllCRFPages"

End Function

'---------------------------------------------------------------------
Private Function CreateAllVisits(lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bNamesAllValid As Boolean
Dim sVisitCode As String
Dim sMsg As String
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
            If Not ValidateItemCode(sVisitCode, gsITEM_TYPE_VISIT, sMsg, True) Then
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
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.CreateAllVisits"
    
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
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.CreateAllPageVisits"
    
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
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.CreateAllPageVisits"
    
End Function

'---------------------------------------------------------------------
Private Function CreateAllDataItems(lClinicalTrialId As Long) As Boolean
'---------------------------------------------------------------------
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bNamesAllValid As Boolean
Dim sMsg As String
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
            If Not ValidateItemCode(sDataItemCode, gsITEM_TYPE_QUESTION, sMsg, False) Then
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
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.CreateAllDataItems"

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
        & " AND DataType = 1 "          ' NCJ 8 May 06 - DataType = 1 is Category
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
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.CreateAllDataItemCategories"

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
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.CreateAllDataItemValidations"

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
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.InsertAllCRFPages"

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
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.NewCLMDataItem"

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
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.UserChangesCode"
   
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
'        'Open the file that logs the codes that have been changed
'        mnFileNumber = FreeFile
'        Open msFileName For Output As #mnFileNumber
'        Print #mnFileNumber, "This file logs the codes that were changed during a session of AREZZO re-build"
'        Print #mnFileNumber,    'blank line
    End If
    
    'add to the collection of changed codes
    moCollectionOfCodesToChange.Add sOldCode, sNewCode, sCodeType, lCodeId
    'write to the changed codes log file
'    Print #mnFileNumber, "CodeType: " & sCodeType & vbTab & "OldCode: " & sOldCode _
'        & vbTab & "NewCode: " & sNewCode & vbTab & "CodeId: " & lCodeId
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.StoreCodesForChanging"

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
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.ProcessCodeChangesInDatabase"

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
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.ArezzoDataItemDoesNotExist"

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
    Err.Raise Err.Number, , Err.Description & "|AzRebuild.ArezzoTaskDoesNotExist"

End Function

'---------------------------------------------------------------------
Private Function TrialNameFromId(ByVal lClinicalTrialId As Long) As String
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT ClinicalTrialName FROM ClinicalTrial " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        TrialNameFromId = ""
    Else
        TrialNameFromId = rsTemp!ClinicalTrialName
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
      Err.Raise Err.Number, , Err.Description & "|AzRebuild.TrialNameFromId"

End Function

'---------------------------------------------------------------------
Private Function RemoveNull(ByVal Text As Variant)
'---------------------------------------------------------------------
'   To take a string that may contain null and return it
'   NCJ 1 Dec 99 - RTrim the string before returning
'---------------------------------------------------------------------
    
    If IsNull(Text) Then
        RemoveNull = ""
    Else
        RemoveNull = RTrim(Text)
    End If

End Function
