Attribute VB_Name = "modSDTrialData"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999 - 2006. All Rights Reserved
'   File:       modSDTrialData.bas
'   Author      Paul Norris, 23/09/99
'   Purpose:    All common TrialData functions for the StudyDefintion project are in this module.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   PN 24/09/99 Moved DeleteTrialPRD() and DeleteTrialSD() back to TrialData module
'   PN  24/09/99    Changed class names
'                   clsDataDefinitionValidations to clsDDValidations
'                   clsDataDefinitionValidation to clsDDValidation
'                   clsDataDefinitionCategories to clsDDCategories
'                   clsDataDefinitionCategory to clsDDCategory
'                   because prog id is too long with original name
'  NCJ 27/9/99      Removed conditional compilation
'  PN  28/09/99     Amended CopyDataItem() and CopyCategories() and CopyValidations()
'                   copy the categories and validations to the new trial
'  NCJ 26 Oct 99 - Added default SingleUseDataItems = 0 in InsertTrial
'   WillC 10/11/99 Added the Error handlers
'   Mo Morris   11/11/99    DAO to ADO conversion
'   NCJ 30 Nov 99   Timestamps as Doubles
'   NCJ 3 Dec 99 - Use single quotes in all SQL strings
'   NCJ 13 Dec 99 - Ids to Long
'   NCJ 16 Dec 99 - CopyDataItem returns 0 if user cancels DItemCode dialog
'   NCJ 15 Jan 00 - SR 2733, Need to include Active, ValueOrder and DefaultCat when overwriting categories
'   NCJ 18/1/00 - SR2746, generate new CRFPageOrder for copied forms
'   NCJ 14/2/00 - Added RoleCode and Mandatory fields when copying CRFElement
'   Mo Morris   5/4/00  SR 3313 Case sensitivity changes made to gblnTrialExists
'   TA 26/09/2000: code copy clinical test question data when copying
'   NCJ 29/12/00, SR 4091 - Need to copy validations in CopyDataItemOverwrite
'   NCJ 02/01/01, SRs 4091, 4092 - Copy validations & categories to Arezzo in Data Item Copy
'   NCJ 30/01/01 - Fixed bug in CopyDataItem (copy data name correctly)
'   NCJ 19/2/01 SR4189 - Only save CLM guideline at the end of CopyCRFPage
'   ATO 12 July 01 - Added NewclinicalTrialID
'   ATO  1/08/2001 -  Change to give option of overwriting all when copying study questions
'   TA 16/04/02: Dummy lock now used when inserting trial so that 2 users cannot so it at the same time
'   REM 17/05/02 - check deleted items before copying in CopyCRFPage and CopyDataItem
'   REM 14/06/02 - CBB 2.2.15 No. 16 changed message returned if DataItem or CRFPage code is a reserved word
'   REM 03/7/02  Roll forward RQG's - Added new Question Group fields to mnCopyCRFElement
'                - Added copying of Question Groups to CopyCRFPage
'   ZA 19/07/2002 - Added font & colour properties of a caption
'   ZA 19/08/2002 - Added MACRO only and description properties for a question
'   ZA 20/08/2002 - Added RFCDefault and ArezzoUpdateStatus fields in InsertTrial routine
'   REM 09/09/02 - Added PasteCRFElement routine for th enew Copy/Paste functionality
'   ZA 11/09/2002 - Remove RFCDefault field from InsertTrialRoutine
'   NCJ 11 Nov 02 - Added DisplayLength and Hotlink fields to mnCopyCRFElement
'   NCJ 17 Jan 03 - Removed call to CloseTrial from InsertTrial
'   ASH 12 FEB 03 - Added CLINICALTESTDATEEXPR field to mnCopyCRFElement
'   NCJ 27 Mar 03 - Added EFormWidth to CopyCRFPage
'   NCJ 30 Apr 03 - Added Timezone offset to TrialStatusHistory insert
'   NCJ 15 Jan 04 - Added extra FieldOrder parameter to mnCopyCRFElement
' MLM 28/06/05: bug 2544: tidied up error handling/transactions/gbDoCLMSave in CopyCRFPage.
'   ic 03/01/2006 copy clinical coding dictionary id
'   NCJ 12 Jun 06 - Added CRFPagesForDataItem to get all eForms where a data item is used
'   NCJ 20 Jun 06 - Mark study as changed at various points
'   NCJ 21 Sept 06 - Added CRFPagesForQGroup to get all eForms where RQG is used
' ic 15/01/2008 issue 2972 check for null value before copying dictionary id
'----------------------------------------------------------------------------------------'

Option Explicit
'added 1/08/2001 ATO
Dim mblnOverwriteAllQuestions As Boolean

'---------------------------------------------------------------------
Public Sub DeleteCRFPage(ByVal lClinicalTrialId As Long, _
                         ByVal nVersionId As Integer, _
                         ByVal lCRFPageId As Long)
'---------------------------------------------------------------------
'Revisions:
'REM 28/02/02 - added QGroup delete
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel
    
    'Begin transaction
    TransBegin
    
    sSQL = "DELETE FROM CRFPage " _
        & "  WHERE  ClinicalTrialId          = " & lClinicalTrialId _
        & "  AND    VersionId                = " & nVersionId _
        & "  AND    CRFPageId                = " & lCRFPageId
                        
    MacroADODBConnection.Execute sSQL
    
    sSQL = "DELETE FROM CRFElement " _
        & "  WHERE  ClinicalTrialId          = " & lClinicalTrialId _
        & "  AND    VersionId                = " & nVersionId _
        & "  AND    CRFPageId                = " & lCRFPageId
                        
    MacroADODBConnection.Execute sSQL
    
    sSQL = "DELETE FROM StudyVisitCRFPage " _
        & "  WHERE  ClinicalTrialId          = " & lClinicalTrialId _
        & "  AND    VersionId                = " & nVersionId _
        & "  AND    CRFPageId                = " & lCRFPageId
    
    MacroADODBConnection.Execute sSQL
    
    'REM 28/02/02 - added QGroup delete
    'Delete QGroup from EFormQGroup table
    sSQL = "DELETE FROM EFormQGroup" _
        & " WHERE   ClinicalTrialID         = " & lClinicalTrialId _
        & " AND     VersionID               = " & nVersionId _
        & " AND     CRFPageID               = " & lCRFPageId
        
    MacroADODBConnection.Execute sSQL
    
    ' PN commented this out because it causes system to hang
    ' Reinstated by NCJ, 23/8/99
    DeleteProformaCRFPage lCRFPageId, lClinicalTrialId
    
    'End transaction
    TransCommit
       
Exit Sub
ErrLabel:
    'RollBack transaction
    TransRollBack
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.DeleteCRFPage"
End Sub

'---------------------------------------------------------------------
Public Function gblnCRFPageExists(ByVal vClinicalTrialId As Long, _
                                  ByVal vVersionId As Integer, _
                                  ByVal vCRFPageCode As String) As Boolean
'---------------------------------------------------------------------
On Error GoTo ErrHandler

Dim rsTemp As ADODB.Recordset
Dim sSQL As String

    sSQL = "SELECT CRFPageId FROM CRFPage " _
        & " WHERE CRFPageCode = '" & vCRFPageCode _
        & "' AND ClinicalTrialId = " & vClinicalTrialId _
        & " AND VersionId = " & vVersionId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        gblnCRFPageExists = False
    Else
        gblnCRFPageExists = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
       
Exit Function
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gblnCRFPageExists", "modSDTrialData")
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
Public Function gblnTrialExists(ByVal vClinicalTrialName As String) As Boolean
'---------------------------------------------------------------------
'Changed Mo Morris 5//4/00  SR 3313
'SQL statemnt made non-case sensitive for Oracle databases using the NLS_LOWER function.
'Access and SQLServer databases are already non-case sensitive
'---------------------------------------------------------------------

On Error GoTo ErrHandler

Dim rsTemp As ADODB.Recordset
Dim sSQL As String

    If goUser.Database.DatabaseType = MACRODatabaseType.oracle80 Then
        sSQL = "SELECT ClinicalTrialId FROM ClinicalTrial " _
            & " WHERE NLS_LOWER(ClinicalTrialName) = '" & LCase(vClinicalTrialName) & "'"
    Else
        sSQL = "SELECT ClinicalTrialId FROM ClinicalTrial " _
            & " WHERE ClinicalTrialName = '" & vClinicalTrialName & "'"
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
'   ATN 21/12/99
'   Oracle recordcount will return -1
    If rsTemp.RecordCount < 1 Then
        gblnTrialExists = False
    Else
        gblnTrialExists = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
       
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gblnTrialExists", "modSDTrialData")
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
Private Function CRFPageElementList(ByVal lClinicalTrialId As Long, _
                                    ByVal nVersionId As Integer, _
                                    ByVal lCRFPageId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
'Retrieve details of CRF Elements (data and visual) on a CRF page
'Creates a ReadOnly recordset
' REM 25/02/02 - added QGroup and OwnerQGroup = 0 so that recordset only returns
' DataItems not in a question group
'---------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrLabel
    
    sSQL = "SELECT * FROM CRFElement " _
        & "  WHERE  ClinicalTrialId             = " & lClinicalTrialId _
        & "  AND    VersionId                   = " & nVersionId _
        & "  AND    CRFElement.CRFPageId        = " & lCRFPageId _
        & "  AND    CRFElement.QGroupId         = 0" _
        & "  AND    CRFElement.OwnerQgroupId    = 0"
    
    Set CRFPageElementList = New ADODB.Recordset
    'changed by Mo Morris 21/12/99, adOpenForwardOnly to adOpenStatic
    CRFPageElementList.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
       
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.CRFPageElementList"
End Function

'---------------------------------------------------------------------
Private Function QGroupList(ByVal lClinicalTrialId As Long, _
                           ByVal nVersionId As Integer, _
                           ByVal lCRFPageId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
' REM 25/02/02
' Retrieves a recordset of all the Question Groups on an EForm
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel
    
    sSQL = "SELECT * FROM CRFElement " _
        & "  WHERE  CRFElement.ClinicalTrialId  = " & lClinicalTrialId _
        & "  AND    CRFElement.VersionId        = " & nVersionId _
        & "  AND    CRFElement.CRFPageId        = " & lCRFPageId _
        & "  AND    CRFElement.QGroupId         > 0" _
        & "  ORDER BY QGroupId"
    
    Set QGroupList = New ADODB.Recordset
    QGroupList.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
       
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.QGroupList"
End Function

'---------------------------------------------------------------------
Public Function CRFElementGroupQuestionList(ByVal lClinicalTrialId As Long, _
                                  ByVal nVersionId As Integer, _
                                  ByVal lCRFPageId As Long, _
                                  ByVal lQGroupId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
' REM 25/02/02
' Retrieves a recordset of a specific Groups Questions from the CRFElement table
'---------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrLabel
    
    sSQL = "SELECT * FROM CRFElement " _
        & " WHERE   CRFElement.ClinicalTrialId  = " & lClinicalTrialId _
        & " AND     CRFElement.VersionId        = " & nVersionId _
        & " AND     CRFElement.CRFPageId        = " & lCRFPageId _
        & " AND     CRFElement.OwnerQGroupId    = " & lQGroupId _
        & " ORDER BY QGroupFieldOrder "
    
    Set CRFElementGroupQuestionList = New ADODB.Recordset
    CRFElementGroupQuestionList.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
       
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.CRFElementGroupQuestionList"

End Function

'---------------------------------------------------------------------
Public Function QGroupQuestionList(ByVal lClinicalTrialId As Long, _
                                  ByVal nVersionId As Integer, _
                                  ByVal lQGroupId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
' REM 28/02/02
' Returns a recordeset of a specific QGroups questions from the QGroupQuestion table
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel

    sSQL = "SELECT * FROM QGroupQuestion " _
        & " WHERE   QGroupQuestion.ClinicalTrialId  = " & lClinicalTrialId _
        & " AND     QGroupQuestion.VersionId        = " & nVersionId _
        & " AND     QGroupQuestion.QGroupId    = " & lQGroupId _
        & " ORDER BY QOrder "
    
    Set QGroupQuestionList = New ADODB.Recordset
    QGroupQuestionList.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.QGroupQuestionList"
End Function

'---------------------------------------------------------------------
Public Function gdsDataItem(ClinicalTrialId As Long, VersionId As Integer, DataItemId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
'Creates a ReadOnly recordset
'---------------------------------------------------------------------
On Error GoTo ErrHandler

Dim sSQL As String
    
    sSQL = "SELECT * FROM DataItem " _
        & " WHERE ClinicalTrialId = " & ClinicalTrialId _
        & " AND VersionId = " & VersionId _
        & " AND DataItemId = " & Str(DataItemId)
    Set gdsDataItem = New ADODB.Recordset
    gdsDataItem.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsDataItem", "modSDTrialData")
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
Public Function EFormQGroup(ByVal lFromClinicalTrialId As Long, ByVal nFromVersionId As Integer, _
                            ByVal lFromQGroupId As Long, ByVal lFromCRFPageId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
' REM 27/02/02
' Returns a recordset containing all the fields for a given EFormQGroup
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel
    
    sSQL = "SELECT * FROM EFormQGroup " _
        & "  WHERE  EFormQGroup.ClinicalTrialId  = " & lFromClinicalTrialId _
        & "  AND    EFormQGroup.VersionId        = " & nFromVersionId _
        & "  AND    EFormQGroup.CRFPageId        = " & lFromCRFPageId _
        & "  AND    EFormQGroup.QGroupId         = " & lFromQGroupId
    
    Set EFormQGroup = New ADODB.Recordset
    EFormQGroup.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.EFormQGroup"
End Function

'---------------------------------------------------------------------
Public Sub EFormQGroupUpdate(ByVal lFromClinicalTrialId As Long, ByVal nFromVersionId As Integer, _
                             ByVal lFromCRFPageId As Long, ByVal lFromQGroupId, _
                             ByVal lToClinicalTrialId As Long, ByVal nToVersionId As Integer, _
                             ByVal lNewCRFPageId As Long, ByVal lNewQGroupId As Long)
'---------------------------------------------------------------------
' REM 22/02/02
' Updates the EFormQGroup object with the copied EFormQGroup
'---------------------------------------------------------------------
Dim rsEFG As ADODB.Recordset
Dim oEFGs As EFormGroupsSD
Dim oEFG As EFormGroupSD

    On Error GoTo ErrLabel
    
    Set rsEFG = EFormQGroup(lFromClinicalTrialId, nFromVersionId, lFromQGroupId, lFromCRFPageId)
    
    ' Get the eFormGroups
    Set oEFGs = New EFormGroupsSD
    Call oEFGs.Load(lToClinicalTrialId, nToVersionId, lNewCRFPageId)
    
    'Create a new EFormQGroup
    Set oEFG = oEFGs.NewEFormGroup(lNewQGroupId, lNewCRFPageId)
    oEFG.Store
    
    With oEFG
        .Border = rsEFG!Border
        .DisplayRows = rsEFG!DisplayRows
        .InitialRows = rsEFG!InitialRows
        .MinRepeats = rsEFG!MinRepeats
        .MaxRepeats = rsEFG!MaxRepeats
    End With
    
    oEFG.Save
    
    Set oEFG = Nothing
    Set oEFGs = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.EFormQGroupUpdate"

End Sub

'---------------------------------------------------------------------
Public Function CopyQGroup(ByVal lFromClinicalTrialId As Long, ByVal nFromVersionId As Integer, _
                                  ByVal lFromQGroupId As Integer, ByVal lToClinicalTrialId As Long, _
                                  ByVal nToVersionId As Integer, Optional sToQGroupCode As String) As QuestionGroup
'---------------------------------------------------------------------
' REM 20/02/02
' Checks to see if QGroup code already exists in study, is so ask if user wants to
' Rename or skip, if not skipped, inserts Question Group into QGroup Table .
' If code doesn't exist then adds Question Group to QGroup
' NB does not copy any of the groups questions
'---------------------------------------------------------------------

Dim sQGroupCode As String
Dim rsQG As ADODB.Recordset
Dim sOption As String
Dim nOption As Integer
Dim sMSG As String

    On Error GoTo ErrLabel

    If sToQGroupCode = "" Then
        Set rsQG = QGroupFromID(lFromClinicalTrialId, nFromVersionId, lFromQGroupId)
        'QGroup code not supplied - retrieve from database
        sQGroupCode = ReplaceQuotes(rsQG!QGroupCode)
    Else
        sQGroupCode = sToQGroupCode 'optional parameter
    End If

    If frmMenu.QuestionGroups.CodeExists(sQGroupCode) Then 'then QGroup code already exists
        
        'ask user if want to rename or skip question group
        sOption = "Rename question group|Skip question group"
        nOption = frmOptionMsgBox.Display("Copying question group - " & sQGroupCode, "", _
                                          "A question group with this code '" & sQGroupCode & "' already exists in this study definition.", sOption, , "Cancel")
                                          
        DoEvents
        
        Select Case nOption
        Case 1 'Rename
            sQGroupCode = GetItemCode("Question Group", "New " & "Question Group" & " code:", sQGroupCode)
            If sQGroupCode = "" Then 'user hasn't entered a new code
                Set CopyQGroup = Nothing
            Else 'update database
                Set CopyQGroup = CopyQGroupAppend(sQGroupCode, rsQG!QGroupName, rsQG!DisplayType)
            End If
        Case 3, glMINUS_ONE 'Skip, or Cancel
            Set CopyQGroup = Nothing
        End Select
        
    Else ' QGroup does not already exist in study

        If Not ValidateItemCode(sQGroupCode, gsITEM_TYPE_QUESTION, sMSG, False) Then
            'not valid there must be an eForm, Visit or Study with the same code
            sOption = "Rename question group|Skip question group"
            nOption = frmOptionMsgBox.Display("Copying question group - " & sQGroupCode, "", _
                                              "A Study, Visit or eForm with the code '" & sQGroupCode & "' already exists in this study definition.", sOption, , "Skip")
                                              
            DoEvents
        
            Select Case nOption
            Case 1 ' Rename
                sQGroupCode = GetItemCode("Question Group", "New " & "Question Group" & " code:", sQGroupCode)
                If sQGroupCode = "" Then 'user hasn't entered a new code
                    Set CopyQGroup = Nothing
                Else 'update database
                    Set CopyQGroup = CopyQGroupAppend(sQGroupCode, rsQG!QGroupName, rsQG!DisplayType)
                End If
            Case 3, glMINUS_ONE 'Skip or Cancel
                Set CopyQGroup = Nothing
            End Select
        
        Else
            'valid so update database
            Set CopyQGroup = CopyQGroupAppend(sQGroupCode, rsQG!QGroupName, rsQG!DisplayType)
        End If

    End If
    
    rsQG.Close
    Set rsQG = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.CopyQGroup"

End Function

'---------------------------------------------------------------------
Public Function CopyDataItem(ByVal lFromClinicalTrialId As Long, ByVal nFromVersionId As Integer, ByVal lFromDataItemId As Long, _
                                ByVal lToClinicalTrialId As Long, ByVal nToVersionId As Integer, Optional sToDataItemCode As String, Optional bSingleQuestion As Boolean = False) As Long
'---------------------------------------------------------------------
' Returns Long instead of Integer - NCJ 13 Dec 99
' Returns 0 if user cancels supplementary dialog asking for new name - NCJ 16 Dec 99
' TA  29/03/2000  Prompt for new dataitem code when existing code conflicts with visit/study/eform
' REM 16/01/02 - added optional parameter called bSingleQuestion, it is set to true when CopyDataItem is used to
'   drag a single question from one study to another
' REM 14/06/02 - CBB 2.2.15 No. 16 changed message returned if DataItem code is a reserved word
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim lDataItemId As Long
Dim sDataItemCode As String
Dim sMSG As String
Dim sPrompt As String
Dim sOption As String
Dim nOption As Long
    On Error GoTo ErrHandler

    If sToDataItemCode = "" Then
        'code not supplied - retrieve from database
        sDataItemCode = DataItemCodeFromId(lFromClinicalTrialId, lFromDataItemId)
'        Set rsTemp = gdsDataItem(lFromClinicalTrialId, nFromVersionId, lFromDataItemId)
'        sDataItemCode = rsTemp!DataItemCode
'        rsTemp.Close
'        Set rsTemp = Nothing

        'REM 13/05/02 - check that a DataItemCode exists in the study being copied from (in case it has been deleted in another instance of SD after question list was opened)
        If sDataItemCode = "" Then
            Call DialogWarning("The question being copied has been deleted from the database.")
            Call frmMenu.RefreshQuestionLists(lFromClinicalTrialId)
            Exit Function
        End If

    Else
        sDataItemCode = sToDataItemCode
    End If
    
    'checks to see if the dataitem exists in the study being copied to
    lDataItemId = DataItemExists(lToClinicalTrialId, nToVersionId, sDataItemCode)
    If lDataItemId <> glMINUS_ONE Then
        'question code already exists
        'prompt user for overwrite
        
        'Added 1/08/2001 ASH
        'New module level boolean, If overwriteall is selected in option then
        'we skip dialog message
        If mblnOverwriteAllQuestions = True Then
            nOption = 1
        Else
        
            'REM 16/01/02 - If bSingle is true then dragging a single qiestion from one study to another
            ' so msg box only display options relevant to single question
            If bSingleQuestion = True Then
                sOption = "Overwrite existing question|Rename question"
                nOption = frmOptionMsgBox.Display("Copying Question - " & sDataItemCode, "", _
                        "A question with the code '" & sDataItemCode & "' already exists in this study definition.", sOption, , "Cancel")
            Else
                sOption = "Overwrite existing question|Rename question|Skip question|Overwrite all questions"
                nOption = frmOptionMsgBox.Display("Copying Question - " & sDataItemCode, "", _
                        "A question with the code '" & sDataItemCode & "' already exists in this study definition.", sOption, , "Skip")
            End If
            
        End If
        'TA 5/5/2000: to stop outline of form lingering
        DoEvents
        
        Select Case nOption
        Case 1  ' Overwrite
            'update question data
            CopyDataItem = CopyDataItemOverwrite(lDataItemId, lFromClinicalTrialId, nFromVersionId, lFromDataItemId, lToClinicalTrialId, nToVersionId, sDataItemCode)
        Case 2  'Rename
            ' Here we ask them for an alternative data item code
            sDataItemCode = GetItemCode(gsITEM_TYPE_QUESTION, "New " & gsITEM_TYPE_QUESTION & " code:", sDataItemCode)
            If sDataItemCode = "" Then
                'NCJ 16/12/99 - Treat this as cancelling the copy
                CopyDataItem = glMINUS_ONE
            Else
                'call copy routine again with new data item code
                CopyDataItem = CopyDataItemAppend(lFromClinicalTrialId, nFromVersionId, lFromDataItemId, lToClinicalTrialId, nToVersionId, sDataItemCode)
            End If
    
        Case 3, glMINUS_ONE 'Skip or Cancel
            ' NCJ 16/12/99 - Cancel selected - cancel the copy
            CopyDataItem = glMINUS_ONE
        
        'Added 1/08/2001 ASH to deal with overwriting all questions during
        'the copying of studies
        Case 4  ' OverwriteAll
             'update question data
             mblnOverwriteAllQuestions = True
            CopyDataItem = CopyDataItemOverwrite(lDataItemId, lFromClinicalTrialId, nFromVersionId, lFromDataItemId, lToClinicalTrialId, nToVersionId, sDataItemCode)
        
        End Select
       
       DoEvents
       
    Else
        ' Data item doesn't already exist
        If Not ValidateItemCode(sDataItemCode, gsITEM_TYPE_QUESTION, sMSG, False) Then
            'not valid must be an eForm, Visit or Study with the same code

            sOption = "Rename question|Skip question"
            'REM 14/06/02 - changed message to return the sMsg from ValidateItemCode
            nOption = frmOptionMsgBox.Display("Copying Question - " & sDataItemCode, _
                "", sMSG, sOption, , "Skip")
            
            'TA 5/5/2000: to stop outline of form lingering
            DoEvents
            
            Select Case nOption
            Case 1
                ' Rename question
                sDataItemCode = GetItemCode(gsITEM_TYPE_QUESTION, "New " & gsITEM_TYPE_QUESTION & " code:", sDataItemCode)
                If sDataItemCode = "" Then
                    'user hasn't entered new code
                    CopyDataItem = glMINUS_ONE
                Else
                    'update database
                    CopyDataItem = CopyDataItemAppend(lFromClinicalTrialId, nFromVersionId, lFromDataItemId, lToClinicalTrialId, nToVersionId, sDataItemCode)
                End If
            Case glMINUS_ONE, 2
                ' Cancel or Skip question
                CopyDataItem = glMINUS_ONE
            End Select
        Else
            'valid so update database
            CopyDataItem = CopyDataItemAppend(lFromClinicalTrialId, nFromVersionId, lFromDataItemId, lToClinicalTrialId, nToVersionId, sDataItemCode)
        End If
    End If
    
    Exit Function
ErrHandler:

    
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CopyDataItem", "modSDTrialData")
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
Private Function CopyDataItemOverwrite(ByVal lDataItemId As Long, _
                                    ByVal lFromClinicalTrialId As Long, _
                                    ByVal nFromVersionId As Integer, _
                                    ByVal lFromDataItemId As Long, _
                                    ByVal lToClinicalTrialId As Long, _
                                    ByVal nToVersionId As Integer, _
                                    ByVal sDataItemCode As String) As Long
'---------------------------------------------------------------------
' TA  29/03/2000
' Database updating for CopyDataItem when overwriting previous Data Item
' Returns data item ID of copied item
' NCJ 9 Jan 02 - MACRO 2.2.6 Bug 19, Need to delete all Proforma Data Item properties
'                   before updating with new ones
'                   Also ensure empty properties are cleared in target question
' ic 03/01/2006 copy clinical coding dictionary id
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sDataItemName As String
Dim sSQL As String
Dim sUnit As String
Dim sDerivation As String
Dim nDataType As Integer
Dim sDescription As String

    On Error GoTo ErrHandler
    
     ' SDM 03/02/00 SR2843
     ' If we're "overwriting" but in the same study there's nothing to do
     ' (This can happen when we're copying a form within the same study)
     If Not ((lFromClinicalTrialId = lToClinicalTrialId) And (nFromVersionId = nToVersionId)) Then
        'Begin transaction
        TransBegin
        
        Set rsTemp = gdsDataItem(lFromClinicalTrialId, nFromVersionId, lFromDataItemId)
         sDataItemName = rsTemp!DataItemName
         nDataType = rsTemp!DataType
         
         sSQL = "UPDATE DataItem SET " _
                 & " DataItemName = '" & ReplaceQuotes(sDataItemName) & "'," _
                 & " DataItemLength = " & rsTemp!DataItemLength _
                 & " ,DataType = " & rsTemp!DataType
         
         If rsTemp!DataItemFormat > "" Then
             sSQL = sSQL & ", DataItemFormat = '" & ReplaceQuotes(rsTemp!DataItemFormat) & "'"
         Else
             sSQL = sSQL & ", DataItemFormat = ''"
         End If
         
         If rsTemp!UnitOfMeasurement > "" Then
             sUnit = rsTemp!UnitOfMeasurement
             sSQL = sSQL & ", UnitOfMeasurement = '" & sUnit & "'"
         Else
             sUnit = ""
              sSQL = sSQL & ", UnitOfMeasurement = ''"
        End If
    
         If rsTemp!Derivation > "" Then
             sDerivation = rsTemp!Derivation
             sSQL = sSQL & ", Derivation = '" & ReplaceQuotes(sDerivation) & "'"
         Else
             sDerivation = ""
             sSQL = sSQL & ", Derivation = ''"
         End If
         
         If rsTemp!DataItemHelpText > "" Then
             sSQL = sSQL & ", DataItemHelpText = '" & ReplaceQuotes(rsTemp!DataItemHelpText) & "'"
         Else
            sSQL = sSQL & ", DataItemHelpText = ''"
         End If
    
        'TA 26/9/00: append clinicaltest id
        If IsNull(rsTemp!ClinicalTestCode) Then
            sSQL = sSQL & ", ClinicalTestCode = null"
        Else
            sSQL = sSQL & ", ClinicalTestCode = '" & rsTemp!ClinicalTestCode & "'"
        End If
   
    sDescription = "'" & ReplaceQuotes(RemoveNull(rsTemp!Description)) & "'"
    If sDescription = "''" Then
        sDescription = "null"
    End If
    
    'ic 03/01/2006 copy clinical coding dictionary id
    If (gbClinicalCoding) Then
        If (IsNull(rsTemp!DictionaryId)) Then
            sSQL = sSQL & ", DictionaryId = null"
        Else
            sSQL = sSQL & ", DictionaryId = " & rsTemp!DictionaryId
        End If
    End If
    
    
        sSQL = sSQL _
                 & " , CopiedFromClinicalTrialId = " & lFromClinicalTrialId _
                 & " , CopiedFromVersionId = " & nFromVersionId _
                 & " , CopiedFromDataItemId = " & lFromDataItemId _
                 & " , MACROOnly = " & rsTemp!MACROOnly _
                 & " , Description = " & sDescription _
                 & " WHERE ClinicalTrialId = " & lToClinicalTrialId _
                 & " AND VersionId = " & nToVersionId _
                 & " AND DataItemId = " & lDataItemId
                 
         MacroADODBConnection.Execute sSQL
         
         ' NCJ 9 Jan 02 - Do not do UpdateProformaDataItem until range values are deleted
'         UpdateProformaDataItem lDataItemId, sDataItemCode, sDataItemName, rsTemp!DataType, sDerivation, sUnit
        
         rsTemp.Close
         Set rsTemp = Nothing
     
        ' Delete category values from ValueData
         sSQL = "DELETE FROM ValueData " _
                 & "WHERE ClinicalTrialId = " & lToClinicalTrialId _
                 & " AND VersionId = " & nToVersionId _
                 & " AND DataItemId = " & lDataItemId
                             
         MacroADODBConnection.Execute sSQL
         
        ' NCJ 02/01/01 Delete category values from Arezzo
        Call DeleteProformaRangeValues(lDataItemId)
         
         ' NCJ 29/12/00 - SR 4091 Need to copy validations too
         ' Delete existing ones first
         sSQL = "DELETE FROM DataItemValidation " _
                 & "WHERE ClinicalTrialId = " & lToClinicalTrialId _
                 & " AND VersionId = " & nToVersionId _
                 & " AND DataItemId = " & lDataItemId
                             
         MacroADODBConnection.Execute sSQL
         
        ' NCJ 02/01/01 Delete existing validations from Arezzo
        Call DeleteProformaWarningConditions(lDataItemId)

        ' NCJ 9 Jan 02 - Do UpdateProformaDataItem here before adding new vals. & cats.
         UpdateProformaDataItem lDataItemId, sDataItemCode, sDataItemName, _
                                nDataType, sDerivation, sUnit
        
        ' NCJ 02/01/01 - Call new Routine to copy validations and categories
        Call CopyValidationsAndCategories(lDataItemId, _
                                    lFromClinicalTrialId, _
                                    nFromVersionId, _
                                    lFromDataItemId, _
                                    lToClinicalTrialId, _
                                    nToVersionId)
                 
     End If
    
    'End transaction
    TransCommit

     CopyDataItemOverwrite = lDataItemId
     
Exit Function
     
ErrHandler:
    'RollBack transaction
    TransRollBack
    
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CopyDataItemOverwrite", "modSDTrialData")
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
Private Sub CopyValidationsAndCategories(ByVal lToDataItemID As Long, _
                                    ByVal lFromClinicalTrialId As Long, _
                                    ByVal nFromVersionId As Integer, _
                                    ByVal lFromDataItemId As Long, _
                                    ByVal lToClinicalTrialId As Long, _
                                    ByVal nToVersionId As Integer)
'---------------------------------------------------------------------
' NCJ 2 Jan 2001 SR3825 (efficiency issues)
' Copy all the Validations and Category values across from lFromDataItemId to lToDataItemId
' Includes Arezzo handling (assuming lToDataItemId is in currently open study)
'---------------------------------------------------------------------
Dim sSQL As String

    ' Categories stored in ValueData table
    'Mo Morris 30/8/01 Db Audit (DefaultCat removed)
    sSQL = "INSERT INTO ValueData (  ClinicalTrialId, VersionId, DataItemId, " _
            & " ValueId, ValueCode, ItemValue, Active, ValueOrder ) " _
            & " SELECT " & lToClinicalTrialId & "," & nToVersionId & "," & lToDataItemID & "," _
            & " ValueId, ValueCode, ItemValue, Active, ValueOrder " _
            & " FROM ValueData " _
            & " WHERE ClinicalTrialId = " & lFromClinicalTrialId _
            & " AND VersionId = " & Str(nFromVersionId) _
            & " AND DataItemId = " & lFromDataItemId
            
    MacroADODBConnection.Execute sSQL
    
    ' NCJ 2/01/01 SR4092 - Copy categories to Arezzo
    Call ReplaceArezzoCategories(lToDataItemID, lFromClinicalTrialId, _
                                   nFromVersionId, lFromDataItemId)
    
    sSQL = "INSERT INTO DataItemValidation ( ClinicalTrialId, VersionId, DataItemId, " _
            & " ValidationId, ValidationTypeID, DataItemValidation, ValidationMessage ) " _
            & " SELECT " & lToClinicalTrialId & "," & nToVersionId & "," & lToDataItemID & "," _
            & " ValidationId, ValidationTypeID, DataItemValidation, ValidationMessage " _
            & " FROM DataItemValidation " _
            & " WHERE ClinicalTrialId = " & lFromClinicalTrialId _
            & " AND VersionId = " & Str(nFromVersionId) _
            & " AND DataItemId = " & lFromDataItemId
            
    MacroADODBConnection.Execute sSQL
    
    ' NCJ 02/01/01 - SR 4091 Need to copy validations to Arezzo
    Call ReplaceArezzoValidations(lToDataItemID, lFromClinicalTrialId, _
                                   nFromVersionId, lFromDataItemId)
     
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "CopyValidationsAndCategories", "modSDTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub ReplaceArezzoCategories(ByVal lToDataItemID As Long, _
                                    ByVal lFromClinicalTrialId As Long, _
                                    ByVal nFromVersionId As Integer, _
                                    ByVal lFromDataItemId As Long)
'---------------------------------------------------------------------
' NCJ 2/1/01 - Replace Data Item's Arezzo categories
' by those from the specified data item (does not delete first)
' Assume lToDataItemID is in the currently open study (if not, we can't do it!)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsCats As ADODB.Recordset
Dim colValues As Collection

    On Error GoTo ErrHandler
    
    ' Select all the category codes (only the code goes to Arezzo)
    sSQL = "SELECT ValueCode FROM ValueData " _
            & " WHERE ClinicalTrialId = " & lFromClinicalTrialId _
            & " AND VersionId = " & Str(nFromVersionId) _
            & " AND DataItemId = " & lFromDataItemId
    Set rsCats = New ADODB.Recordset
    rsCats.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' Add each category code to Values collection
    Set colValues = New Collection
    Do While Not rsCats.EOF
        colValues.Add CStr(RemoveNull(rsCats!ValueCode))
        rsCats.MoveNext
    Loop
    
    rsCats.Close
    
    ' Save new ones to Arezzo
    Call SaveProformaRangeValues(lToDataItemID, colValues)
    
    Set rsCats = Nothing
    Set colValues = Nothing
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "ReplaceArezzoCategories", "modSDTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub ReplaceArezzoValidations(ByVal lToDataItemID As Long, _
                                    ByVal lFromClinicalTrialId As Long, _
                                    ByVal nFromVersionId As Integer, _
                                    ByVal lFromDataItemId As Long)
'---------------------------------------------------------------------
' NCJ 2/1/01 - Replace Data Item's Arezzo Warning Conditions
' by those from the specified data item
' Assume lToDataItemID is in the currently open study (if not, we can't do it!)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsWConds As ADODB.Recordset
Dim colConds As Collection
Dim colFlags As Collection

    On Error GoTo ErrHandler
    
    ' Select all the "flags" and "conditions" for Arezzo
    sSQL = "SELECT ValidationId, DataItemValidation FROM DataItemValidation " _
            & " WHERE ClinicalTrialId = " & lFromClinicalTrialId _
            & " AND VersionId = " & Str(nFromVersionId) _
            & " AND DataItemId = " & lFromDataItemId
    Set rsWConds = New ADODB.Recordset
    rsWConds.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' Add each flag and condition to collections
    Set colConds = New Collection
    Set colFlags = New Collection
    Do While Not rsWConds.EOF
        colFlags.Add CStr(RemoveNull(rsWConds!ValidationID))
        colConds.Add CStr(RemoveNull(rsWConds!DataItemValidation))
        rsWConds.MoveNext
    Loop
    
    rsWConds.Close
    
    ' Save new ones to Arezzo
    Call SaveProformaWarningConditions(lToDataItemID, colFlags, colConds)
    
    Set rsWConds = Nothing
    Set colConds = Nothing
    Set colFlags = Nothing
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "ReplaceArezzoValidations", "modSDTrialData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Public Function CopyQGroupAppend(ByVal sQGroupCode As String, ByVal sQGroupName As String, _
                                ByVal nDisplayType As Integer) As QuestionGroup
'---------------------------------------------------------------------
' REM 20/02/02
' Database updating for CopyQGroup when appending (i.e. adding new) QGroup to the QGroup Table
' Returns the newly created QGroup
'---------------------------------------------------------------------
Dim oNewQG As QuestionGroup

    On Error GoTo ErrLabel
    
    Set oNewQG = frmMenu.QuestionGroups.NewGroup(sQGroupCode)
    
    With oNewQG
        .QGroupName = sQGroupName
        .DisplayType = nDisplayType
    End With
    
    ' NCJ 20 Jun 06 - Mark study as changed
    Call frmMenu.MarkStudyAsChanged

    Set CopyQGroupAppend = oNewQG
    Set oNewQG = Nothing
    
Exit Function
ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.CopyQGroupAppend"

End Function

'---------------------------------------------------------------------
Private Function CopyDataItemAppend(ByVal lFromClinicalTrialId As Long, _
                                ByVal nFromVersionId As Integer, _
                                ByVal lFromDataItemId As Long, _
                                ByVal lToClinicalTrialId As Long, _
                                ByVal nToVersionId As Integer, _
                                ByVal sDataItemCode As String) As Long
'---------------------------------------------------------------------
' TA 29/03/2000
' Database updating for CopyDataItem when appending (i.e. adding new) Data Item
' Returns ID of newly created data item
' NCJ 2 Jan 01 - Tidied code, changed copying of validations & categories
' ic 03/01/2006 copy clinical coding dictionary id
' ic 15/01/2008 issue 2972 check for null value before copying dictionary id
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim lNewDataItemId As Long
Dim sDataItemName As String
Dim nDataItemType As Integer
Dim sUniqueExportName As String
Dim sSQL As String
Dim sUnit As String
Dim sDerivation As String

    On Error GoTo ErrHandler
    
    'Begin transaction
    TransBegin
    
    ' Get new Arezzo data item ID - NCJ 12/8/99
    lNewDataItemId = gnNewCLMDataItem(sDataItemCode)
    
    Set rsTemp = gdsDataItem(lFromClinicalTrialId, nFromVersionId, lFromDataItemId)
    
    ' NCJ 12/8/99
    sDataItemName = ReplaceQuotes(rsTemp!DataItemName)
    nDataItemType = rsTemp!DataType

    ' PN 17/09/99
    ' changed parameters lFromClinicalTrialId to lToClinicalTrialId and
    ' nFromVersionId to nToVersionId
    sUniqueExportName = GenerateNextUniqueExportName(sDataItemCode, lToClinicalTrialId, nToVersionId)
    
    'changed by Mo Morris 4/12/99 single quotes removed from iDataItemType
    'TA 26/9/00: append clinicaltest id
    ' NCJ 30 Jan 01 - Corrected second sDataItemCode to sDataItemName
    'Mo Morris 30/8/01 Db Audit (Required & RequiredTrialTypeId removed)
    sSQL = "INSERT INTO DataItem (  ClinicalTrialId, VersionId, DataItemId, " _
        & "DataItemCode, DataItemName, DataType, ExportName, DataItemCase," _
        & "DataItemFormat, UnitOfMeasurement, DataItemLength,  " _
        & "Derivation, DataItemHelpText, ClinicalTestCode," _
        & "CopiedFromClinicalTrialId, CopiedFromVersionId, CopiedFromDataItemId, " _
        & "MACROOnly, Description"
        
    'ic 03/01/2006 copy clinical coding dictionary id
    If (gbClinicalCoding) Then
        sSQL = sSQL & ", DictionaryId"
    End If
        
    sSQL = sSQL & ") VALUES (" & lToClinicalTrialId & "," & nToVersionId & "," & lNewDataItemId & "," _
        & "'" & sDataItemCode & "','" & sDataItemName & "'," & nDataItemType

    sSQL = sSQL & ", '" & sUniqueExportName & "'"
    sSQL = sSQL & ", " & rsTemp!DataItemCase
    sSQL = sSQL & ", '" & ReplaceQuotes(RemoveNull(rsTemp!DataItemFormat)) & "'"
    
    sUnit = RemoveNull(rsTemp!UnitOfMeasurement)
    sSQL = sSQL & ", '" & ReplaceQuotes(sUnit) & "'"
    sSQL = sSQL & "," & rsTemp!DataItemLength
    
    sDerivation = RemoveNull(rsTemp!Derivation)
    sSQL = sSQL & ",  '" & ReplaceQuotes(sDerivation) & "'"
    sSQL = sSQL & ", '" & ReplaceQuotes(RemoveNull(rsTemp!DataItemHelpText)) & "'"

    'TA 26/9/00: append clinicaltest id
    If IsNull(rsTemp!ClinicalTestCode) Then
        sSQL = sSQL & ", null"
    Else
        sSQL = sSQL & ", '" & rsTemp!ClinicalTestCode & "'"
    End If

    sSQL = sSQL _
            & " , " & lFromClinicalTrialId _
            & " , " & nFromVersionId _
            & " ,  " & lFromDataItemId _
    '        & " )"
    
    'ZA add MACROONLY, description
    sSQL = sSQL & " , " & rsTemp!MACROOnly
    
    If IsNull(rsTemp!Description) Then
        sSQL = sSQL & ", null"
    Else
        sSQL = sSQL & ", '" & ReplaceQuotes(rsTemp!Description) & "'"
    End If
    
    'ic 03/01/2006 copy clinical coding dictionary id
    If (gbClinicalCoding) Then
        ' ic 15/01/2008 issue 2972 check for null value before copying dictionary id
        If IsNull(rsTemp!DictionaryId) Then
            sSQL = sSQL & ", null"
        Else
            sSQL = sSQL & ", " & rsTemp!DictionaryId
        End If
    End If
    
    sSQL = sSQL & " )"
            
    MacroADODBConnection.Execute sSQL
    
    ' Changed NCJ 12/8/99
    ' InsertProformaDataItem lToClinicalTrialId, nToVersionId, lNewDataItemId, sDataItemCode, rsTemp!DataType, sDerivation, msMandatoryValidation, msWarningValidation, msDataManagerValidation
    UpdateProformaDataItem lNewDataItemId, sDataItemCode, sDataItemName, nDataItemType, sDerivation, sUnit
    
    rsTemp.Close
    Set rsTemp = Nothing
    
'    ' PN change 20
'    ' PN 28/09/99 copy the categories and validations to the new trial
'    Call CopyValidations(lFromClinicalTrialId, nFromVersionId, lToClinicalTrialId, _
'                        nToVersionId, lFromDataItemId, lNewDataItemId)
'    Call CopyCategories(lFromClinicalTrialId, nFromVersionId, lToClinicalTrialId, _
'                        nToVersionId, lFromDataItemId, lNewDataItemId)
                        
    ' NCJ 02/01/01 - Call new (more efficient) routine to copy validations and categories
    Call CopyValidationsAndCategories(lNewDataItemId, _
                                    lFromClinicalTrialId, _
                                    nFromVersionId, _
                                    lFromDataItemId, _
                                    lToClinicalTrialId, _
                                    nToVersionId)
        
    'End transaction
    TransCommit
                        
    CopyDataItemAppend = lNewDataItemId
        
Exit Function
     
ErrHandler:
    'RollBack transaction
    TransRollBack
    
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CopyDataItemAppend", "modSDTrialData")
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
Public Function InsertTrial(ByRef rClinicalTrialId As Long, _
                            ByRef rVersionId As Integer, _
                            ByRef rClinicalTrialName As String, _
                            Optional ByVal vCopyFromClinicalTrialId As Variant) As Boolean
'---------------------------------------------------------------------
' NCJ 30 Nov 99 - Store timestamps as doubles
'WillC 4/2/00 changed the dates to SQLStandardNow to avoid regional setting problems
'TA 16/04/2002: Return whether insert was successful
'ZA 20/08/2002 - Insert 0,0 for RFCDefault and ArezzoUpdateStatus column values
'ZA 11/09/2002 - Removed RFCDefault field
' NCJ 17 Jan 03 - Removed call to CloseTrial
'               (assume there is NO trial open when this routine is called)
' NCJ 30 Apr 03 - Added time zone info
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTrial As ADODB.Recordset
Dim sDate As String
Dim sStudyToken As String
Dim oTimezone As TimeZone

    On Error GoTo ErrHandler
    
    'TA 16/04/2002: lock db from creating new trials
    sStudyToken = MACROLOCKBS30.LockStudy(gsADOConnectString, goUser.UserName, -1) '-1 denotes lock new studies
    Select Case sStudyToken
    Case MACROLOCKBS30.DBLocked.dblStudy
        DialogInformation "A study definition is currently being created by another user. Please try again later."
        InsertTrial = False
        Exit Function
    Case Else
        'lock successful   'Case MACROLOCKBS30.DBLocked.dblSubject should never happen
    End Select
    
    'Begin transaction
    TransBegin
    
    'ATO 12/07/01 replaced with new function
     rClinicalTrialId = NewClinicalTrialID
    
    ' Timestamp as double - NCJ 30 Nov 99
    ' WillC 4/2/00 changed the dates to SQLStandardNow
    sDate = SQLStandardNow
    
    Set oTimezone = New TimeZone
    
    If IsMissing(vCopyFromClinicalTrialId) Then         'not being copied
    
        rVersionId = 1
        
        'Changed by Mo Morris 6/8/99, Sponsor removed from SQL and TrialTypeId added
        sSQL = "INSERT INTO ClinicalTrial (ClinicalTrialId,ClinicalTrialName," _
                & "ClinicalTrialDescription, PhaseId, StatusId, Keywords, ExpectedRecruitment, TrialTypeId) " _
                & "VALUES (" & rClinicalTrialId & ",'" & rClinicalTrialName _
                & "','" & rClinicalTrialName & "',1,1,'', 0, 0)"
                       
        MacroADODBConnection.Execute sSQL
        
        ' PN change 21 to add defaults for stand. time and date
        ' NCJ 26 Oct 99 - Added default SingleUseDataItems = 0
        ' NCJ 30 Nov 99 - Use dblDate instead of GetTimeSTamp
        ' Mo Morris 30/8/01 Db Audit (UserId to UserName)
        ' ZA 29/08/2002 - Reason for change is set to 1 for a new study
        sSQL = "INSERT INTO StudyDefinition (ClinicalTrialId,VersionId," _
                & "DefaultFontColour, DefaultCRFPageColour, DefaultFontName, " _
                & "DefaultFontBold,DefaultFontItalic,DefaultFontSize, " _
                & "UserName,StudyDefinitionTimeStamp, " _
                & "StandardDateFormat, StandardTimeFormat, " _
                & "SingleUseDataItems, TrialSubjectLabel,LocalTrialSubjectLabel, " _
                & " ArezzoUpdateStatus) " _
                & "VALUES (" & rClinicalTrialId & "," & rVersionId & "," _
                & gDefaultFontColour & "," & gDefaultCRFPageColour & ",'" & gDefaultFontName & "'," _
                & gDefaultFontBold & "," & gDefaultFontItalic & "," & gDefaultFontSize & ",'" _
                & goUser.UserName & "'," & sDate & ", 'dd/mm/yyyy', 'hh:mm:ss', 0,'',0," _
                & eArezzoUpdateStatus.auNotRequired & " )"
    
        MacroADODBConnection.Execute sSQL
        
        'Mo Morris 30/8/01 Db Audit (UserId to UserName)
        ' NCJ 30 Apr 03 - Added StatusChangedTimestamp_TZ
        sSQL = "INSERT INTO TrialStatusHistory ( ClinicalTrialId, VersionId, " _
            & " TrialStatusChangeId, StatusId, UserName, " _
            & " StatusChangedTimestamp, StatusChangedTimestamp_TZ)" _
            & " VALUES (" & rClinicalTrialId & "," & rVersionId & "" _
            & ",1,1,'" & goUser.UserName & "'," _
            & sDate & ", " & oTimezone.TimezoneOffset & ")"
            
        MacroADODBConnection.Execute sSQL
        
        ' NCJ 17 Jan 03 - Removed call to CloseTrial
'        frmMenu.CloseTrial
            
        CreateProformaTrial rClinicalTrialId, rClinicalTrialName
                        
    End If
    
    'End transaction
    TransCommit
    
    Set oTimezone = Nothing
    
    'TA 04/07/2001: release 'creating new trial lock'
    MACROLOCKBS30.UnlockStudy gsADOConnectString, sStudyToken, -1
       
    InsertTrial = True
     
Exit Function

ErrHandler:
    'RollBack transaction
    TransRollBack
    
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "InsertTrial", "modSDTrialData")
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
Private Function CRFElements(lClinicalTrialId As Long, nVersionId As Integer, lCRFPageId As Long, sCRFElementIdlist As String) As ADODB.Recordset
'---------------------------------------------------------------------
'REM 09/09/02
'Returns a recordset of all selected CRFElements (as defined by the string sCRFElementIdlist) from the CRFElement table
'---------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrLabel

    'get all the CRFElements from the CRFPage being copied from
    sSQL = "SELECT * FROM CRFElement" _
         & " WHERE ClinicalTrialID = " & lClinicalTrialId & "" _
         & " AND VersionID = " & nVersionId _
         & " AND CRFPageID = " & lCRFPageId _
         & " AND CRFElementID IN (" & sCRFElementIdlist & ")" _
         & " ORDER BY CRFElementID"
         
    Set CRFElements = New ADODB.Recordset
    CRFElements.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.CRFElements"
End Function

'---------------------------------------------------------------------
Public Function PasteCRFElements(ByVal lClinicalTrialId As Long, ByVal nVersionId As Integer, _
                                 ByVal lFromCRFPageId As Long, ByVal lToCRFPageId As Long, _
                                 ByVal colCopiedCRFElementIds As Collection) As Boolean
'---------------------------------------------------------------------
'REM 06/09/02
'Function to paste copied CRFElements, checks to see if the CRFElements being copied still exist on the eForm
'being copied from and that no elements being copied already exist on the form being copied to
' NCJ 29 May 03 - We do NOT check the colCopiedCRFElementIds exist here - assume done already!
' NCJ 15 Jan 04 - Generate new field orders for pasted elements (to avoid duplicates, which cause problems elsewhere)
'---------------------------------------------------------------------
Dim rsPasteCRFElements As ADODB.Recordset
Dim rsGroupElements As ADODB.Recordset
Dim sSQL As String
Dim vCRFElement As Variant
Dim sCRFElementIdlist As String
Dim nCRFElementID As Integer
Dim colCopiedDataItemIds As Collection
Dim colCopiedQGroupIds As Collection
Dim nFieldOrder As Integer

    On Error GoTo ErrLabel
    
    ' NCJ 29 May 03 - BUG 1813 - Removed check that colCopiedCRFElementIds exist - assume this has been done already!
    'Check to see if the CRFElements still exist on the form being copied from
    '(can only check questions and QGroups but not Group Questions as the ElementIds of Group Questions are not passed in)
'    If CRFElementsExist(lClinicalTrialId, nVersionId, lFromCRFPageId, colCopiedCRFElementIds) = False Then
'        PasteCRFElements = False
'        Exit Function
'    End If
    
    sCRFElementIdlist = ""
     
    'loop through the collection of copied element ids and create a string of them
    For Each vCRFElement In colCopiedCRFElementIds
        If sCRFElementIdlist <> "" Then sCRFElementIdlist = sCRFElementIdlist & ","
        sCRFElementIdlist = sCRFElementIdlist & vCRFElement
    Next
    
    'get a recordset of the all the copied CRFElements from the CRFPage being copied from (this excludes the Group Questions)
    Set rsPasteCRFElements = CRFElements(lClinicalTrialId, nVersionId, lFromCRFPageId, sCRFElementIdlist)
    
    'Set up new collections
    Set colCopiedDataItemIds = New Collection
    Set colCopiedQGroupIds = New Collection
    
    'Returns collections of all DataItemIds and QGroupIds being copied (includes group question DataItemIds)
    Call CollectionsOfDataItemAndQGroupIds(rsPasteCRFElements, colCopiedDataItemIds, colCopiedQGroupIds)

    'Check if the copied elements can be pasted onto the selected form,
    'i.e do any of the questions, QGroups or Group Questions exist on the form being copied to
    If CRFElementsCanBePasted(lClinicalTrialId, nVersionId, lToCRFPageId, colCopiedDataItemIds, colCopiedQGroupIds) Then
        
        'move to first record
        rsPasteCRFElements.MoveFirst
        
        'loop through all the copied elements and add them to the CRFElement table
        'and add the QGroups the to EFormQGroup table
        Do While Not rsPasteCRFElements.EOF
             'copy all the CRFElements to the CRFElement table
       
            ' NCJ 15 Jan 04 - We have to generate new field orders here for questions/groups on the target eForm
            ' to prevent duplicate field orders (which cause crashes elsewhere)
            If rsPasteCRFElements!DataItemId > 0 Or rsPasteCRFElements!QGroupID > 0 Then
                nFieldOrder = mnNextFieldOrder(lClinicalTrialId, nVersionId, lToCRFPageId)
            Else
                ' Not a question/group
                nFieldOrder = 0
            End If
            
            nCRFElementID = mnCopyCRFElement(lClinicalTrialId, nVersionId, lFromCRFPageId, _
                                    rsPasteCRFElements!CRFelementID, lClinicalTrialId, nVersionId, lToCRFPageId, _
                                    rsPasteCRFElements!DataItemId, _
                                    rsPasteCRFElements!QGroupID, rsPasteCRFElements!OwnerQGroupID, _
                                    nFieldOrder)
        
            'If an element is a QGroup then needs to be added to the EFormQGroup table and the
            'Group Questions need to be added to the CRFElement table
            If rsPasteCRFElements!QGroupID > 0 Then
            
                ' Get a recordset of the Group Questions for the specific QGroup from the CRFElement table
                Set rsGroupElements = GetQGroupQuestions(lClinicalTrialId, nVersionId, lFromCRFPageId, _
                                                    rsPasteCRFElements!QGroupID)
                 
                ' Add the Group Questions to the CRFElement table
                ' NCJ 15 Jan 04 - Use previously generated FieldOrder
                ' (All the questions in a group share the same field order)
                Do While Not rsGroupElements.EOF
                    nCRFElementID = mnCopyCRFElement(lClinicalTrialId, nVersionId, lFromCRFPageId, _
                                        rsGroupElements!CRFelementID, lClinicalTrialId, nVersionId, _
                                        lToCRFPageId, rsGroupElements!DataItemId, _
                                        rsGroupElements!QGroupID, rsGroupElements!OwnerQGroupID, nFieldOrder)
                    rsGroupElements.MoveNext
                Loop
                'close recordset
                rsGroupElements.Close
                
                'Insert copied EFormGroup into the EFormQGroup table with the new CRFPageId
                sSQL = "INSERT INTO EFormQGroup" & _
                    "(ClinicalTrialID, VersionID,CRFPageID,QGroupID,Border,DisplayRows,InitialRows,MinRepeats,MaxRepeats)" & _
                    " SELECT " & lClinicalTrialId & "," & nVersionId _
                    & "," & lToCRFPageId & ", EFormQGroup.QGroupID" _
                    & ", EFormQGroup.Border, EFormQGroup.DisplayRows" _
                    & ", EFormQGroup.InitialRows, EFormQGroup.MinRepeats" _
                    & ", EFormQGroup.MaxRepeats" _
                    & " FROM EFormQGroup" _
                    & " WHERE ClinicalTrialId = " & lClinicalTrialId _
                    & " AND CRFPageId = " & lFromCRFPageId _
                    & " AND VersionId = " & nVersionId _
                    & " AND QGroupId = " & rsPasteCRFElements!QGroupID
                MacroADODBConnection.Execute sSQL
                  
            End If
            rsPasteCRFElements.MoveNext
        Loop
        
        PasteCRFElements = True
    Else
        
        PasteCRFElements = False
        
    End If

    'set the recordsets and collections to nothing
    rsPasteCRFElements.Close
    Set rsPasteCRFElements = Nothing

    Set rsGroupElements = Nothing
    
    Set colCopiedDataItemIds = Nothing
    Set colCopiedQGroupIds = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.PasteCRFElements"
End Function

'---------------------------------------------------------------------
Private Function CRFElementsCanBePasted(lClinicalTrialId As Long, nVersionId As Integer, _
                                        lToCRFPageId As Long, _
                                        colCopiedDataItemIds As Collection, _
                                        colCopiedQGroupIds As Collection) As Boolean
'---------------------------------------------------------------------
'REM 06/09/02
'Checks to see if the copied CRFElements can be pasted
'i.e. do any of the dataitems or QGroups already exist on the CRFPage being copied to
'---------------------------------------------------------------------
Dim rsCRFElements As ADODB.Recordset
Dim sSQL As String
Dim i As Integer
Dim colDataItemIDs As Collection
Dim colQGroupIds As Collection

    On Error GoTo ErrLabel
        
     'get all the CRFElements on the CRFPage being copied to
     sSQL = "SELECT * FROM CRFElement " _
         & " WHERE CRFElement.ClinicalTrialId = " & lClinicalTrialId _
         & " AND CRFElement.VersionId = " & nVersionId _
         & " AND CRFElement.CRFPageId = " & lToCRFPageId
     
     Set rsCRFElements = New ADODB.Recordset
     rsCRFElements.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    'set up the collections
    Set colDataItemIDs = New Collection
    Set colQGroupIds = New Collection
    
    'reurns collections of DataItemIds and QGroupIds from the CRFPage being copied to
    Call CollectionsOfDataItemAndQGroupIds(rsCRFElements, colDataItemIDs, colQGroupIds)
    
    'check if any of the dataitems exist on the form being copied to
    For i = 1 To colCopiedDataItemIds.Count
        If CollectionMember(colDataItemIDs, Str(colCopiedDataItemIds.Item(i)), False) Then
            'at least one of the copied dataitems exist on the form being copied to
            CRFElementsCanBePasted = False
            Exit Function
        Else
            ' Now check for single use dataitems and see if it's been used already
            If gbSingleUseDataItems Then
            'TODO - DataItemCodeUsedInStudy needs to be changed so that part of it returns DataItemIdUsedInStudy
'                If DataItemIdUsedInStudy(lClinicalTrialId, colCopiedDataItemIds.Item(i)) Then
'                    CRFElementsCanBePasted = False
'                    Exit Function
'                End If
            End If
            
        End If
    Next

    'check if any of the QGroups already exist on the form being copied to
    For i = 1 To colCopiedQGroupIds.Count
        If CollectionMember(colQGroupIds, Str(colCopiedQGroupIds.Item(i)), False) Then
            'at least one of the copied QGroups exist on the form being copied to
            CRFElementsCanBePasted = False
            Exit Function
        End If
    Next
    
    CRFElementsCanBePasted = True
    
    'set the recordsets and collection to nothing
    rsCRFElements.Close
    Set rsCRFElements = Nothing
    
    Set colDataItemIDs = Nothing
    Set colQGroupIds = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.CRFElementsCanBePasted"
End Function

'---------------------------------------------------------------------
Public Function CRFPagesForDataItem(lClinicalTrialId As Long, nVersionId As Integer, lDataItemId As Long)
'---------------------------------------------------------------------
' NCJ 12 Jun 06 - Get all the eForms on which this question is used
'---------------------------------------------------------------------
Dim rsCRFElements As ADODB.Recordset
Dim sSQL As String
Dim colCRFPages As Collection

    On Error GoTo ErrLabel

    Set colCRFPages = New Collection
    
     ' Get all the CRFPages for this data item
     sSQL = "SELECT CRFPageID FROM CRFElement " _
         & " WHERE CRFElement.ClinicalTrialId = " & lClinicalTrialId _
         & " AND CRFElement.VersionId = " & nVersionId _
         & " AND CRFElement.DataItemId = " & lDataItemId
     
     Set rsCRFElements = New ADODB.Recordset
     rsCRFElements.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    ' Loop through recordset of eForms
    Do While Not rsCRFElements.EOF
        colCRFPages.Add CLng(rsCRFElements!CRFPageId)
        rsCRFElements.MoveNext
    Loop

    rsCRFElements.Close
    Set rsCRFElements = Nothing
    
    Set CRFPagesForDataItem = colCRFPages
    Set colCRFPages = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.CRFPagesForDataItem"
End Function

'---------------------------------------------------------------------
Public Function CRFPagesForQGroup(lClinicalTrialId As Long, nVersionId As Integer, lQGroupId As Long)
'---------------------------------------------------------------------
' NCJ 21 Sept 06 - Get all the eForms on which this RQG is used
'---------------------------------------------------------------------
Dim rsCRFElements As ADODB.Recordset
Dim sSQL As String
Dim colCRFPages As Collection

    On Error GoTo ErrLabel

    Set colCRFPages = New Collection
    
     ' Get all the CRFPages for this data item
     sSQL = "SELECT CRFPageID FROM CRFElement " _
         & " WHERE CRFElement.ClinicalTrialId = " & lClinicalTrialId _
         & " AND CRFElement.VersionId = " & nVersionId _
         & " AND CRFElement.QGroupId = " & lQGroupId
     
     Set rsCRFElements = New ADODB.Recordset
     rsCRFElements.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    ' Loop through recordset of eForms
    Do While Not rsCRFElements.EOF
        colCRFPages.Add CLng(rsCRFElements!CRFPageId)
        rsCRFElements.MoveNext
    Loop

    rsCRFElements.Close
    Set rsCRFElements = Nothing
    
    Set CRFPagesForQGroup = colCRFPages
    Set colCRFPages = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.CRFPagesForQGroup"
End Function

'---------------------------------------------------------------------
Public Function CRFElementsExist(lClinicalTrialId As Long, nVersionId As Integer, lFromCRFPageId As Long, colCopiedCRFElementIds As Collection) As Boolean
'---------------------------------------------------------------------
'REM 06/09/02
'Checks for the existence of given CRFElements
' NCJ 29 May 03 - BUG 1813 - Made public so we can give correct error messages when pasting
'---------------------------------------------------------------------
Dim rsCRFElemnts As ADODB.Recordset
Dim colCRFelementIds As Collection
Dim vCRFElementId As Variant
Dim sSQL As String
Dim i As Integer

    On Error GoTo ErrLabel
    
    CRFElementsExist = True
    
    If colCopiedCRFElementIds.Count = 0 Then Exit Function
    
    'Get all the CRFElements from the copied page
    sSQL = "SELECT CRFElementId FROM CRFElement " _
        & " WHERE CRFElement.ClinicalTrialId = " & lClinicalTrialId _
        & " AND CRFElement.VersionId = " & nVersionId _
        & " AND CRFElement.CRFPageId = " & lFromCRFPageId
    
    Set rsCRFElemnts = New ADODB.Recordset
    rsCRFElemnts.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    'check there are some
    If rsCRFElemnts.RecordCount > 0 Then
        Set colCRFelementIds = New Collection
        'loop through them and add to a collection
        Do While Not rsCRFElemnts.EOF
            colCRFelementIds.Add rsCRFElemnts.Fields(0).Value, Str(rsCRFElemnts.Fields(0).Value)
            rsCRFElemnts.MoveNext
        Loop
        
        'loop through the collection of copied CRFElements and sees if the Elements still exist
        'on the CRFPage being copied from
        For i = 1 To colCopiedCRFElementIds.Count
            If Not CollectionMember(colCRFelementIds, Str(colCopiedCRFElementIds.Item(i)), False) Then
                CRFElementsExist = False
                Exit For
            End If
        Next
    Else
        'there are no CRFElements on the from CRFPage
        CRFElementsExist = False
    End If
    
    rsCRFElemnts.Close
    Set rsCRFElemnts = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.CRFElementsExist"
End Function

'---------------------------------------------------------------------
Private Sub CollectionsOfDataItemAndQGroupIds(ByVal rsCRFElements As ADODB.Recordset, ByRef colDataItemIDs As Collection, ByRef colQGroupIds As Collection)
'---------------------------------------------------------------------
'REM 09/09/02
'returns collection of DataItemIds, QGroupIds and CRFelementIds from a recordset of CRFElements from one eForm
'---------------------------------------------------------------------
Dim rsQGroupQuestions As ADODB.Recordset
Dim lClinicalTrialId As Long
Dim nVersionId As Integer
Dim lCRFPageId As Long
Dim lDataItemId As Long
Dim lQGroupId As Long
Dim lOwnerQGroupId As Long
Dim sSQL As String

    On Error GoTo ErrLabel
    
    'loop through recordset of CRFElements
    Do While Not rsCRFElements.EOF
        
        lClinicalTrialId = rsCRFElements!ClinicalTrialId
        nVersionId = rsCRFElements!VersionId
        lCRFPageId = rsCRFElements!CRFPageId
        lDataItemId = rsCRFElements!DataItemId
        lQGroupId = rsCRFElements!QGroupID
        lOwnerQGroupId = rsCRFElements!OwnerQGroupID

        
        'if has a DataItemId but not OwnerQGroupId (QGroup Questions will be handled in the ElseIf) then add to dataitem collection
        If (lDataItemId > 0) And (lOwnerQGroupId = 0) Then
            colDataItemIDs.Add lDataItemId, Str(lDataItemId)
        ' else if has no DataItemId but does have a QGroupId then add to QGroup collection
        ElseIf (lDataItemId = 0) And (lQGroupId > 0) Then
            'add QGroupId to QGroup collection
            colQGroupIds.Add lQGroupId, Str(lQGroupId)
            
            'recordset of the group questions
            Set rsQGroupQuestions = GetQGroupQuestions(lClinicalTrialId, nVersionId, lCRFPageId, lQGroupId)
             
            'Then add all the QGroup Questions to the DataItem collection
            Do While Not rsQGroupQuestions.EOF
                'field 4 is the DataItemId
                colDataItemIDs.Add rsQGroupQuestions.Fields(4).Value, Str(rsQGroupQuestions.Fields(4).Value)
                rsQGroupQuestions.MoveNext
            Loop
            
            rsQGroupQuestions.Close
            Set rsQGroupQuestions = Nothing
            
        End If
        
        rsCRFElements.MoveNext
    Loop
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.CollectionsOfDataItemAndQGroupIds"
End Sub

'---------------------------------------------------------------------
Private Function GetQGroupQuestions(lClinicalTrialId As Long, nVersionId As Integer, lCRFPageId As Long, lQGroupId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
'REM 09/09/02
'returns recordset of Group Questions for a specific QGroup
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel

    'get the Group Questions from the CRFElement table
    sSQL = "SELECT * FROM CRFElement" _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND VersionId = " & nVersionId _
        & " AND CRFPageId = " & lCRFPageId _
        & " AND OwnerQGroupId = " & lQGroupId
    Set GetQGroupQuestions = New ADODB.Recordset
    GetQGroupQuestions.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.GetQGroupQuestions"
End Function

'---------------------------------------------------------------------
Private Function mnCopyCRFElement(ByVal lFromClinicalTrialId As Long, _
                                  ByVal nFromVersionId As Integer, _
                                  ByVal lFromCRFPageId As Long, _
                                  ByVal nFromCRFElementId As Integer, _
                                  ByVal lToClinicalTrialId As Long, _
                                  ByVal nToVersionId As Integer, _
                                  ByVal lToCRFPageId As Long, _
                                  ByVal lNewDataItemId As Long, _
                                  ByVal lQGroupId As Long, _
                                  ByVal lOwnerQGroupId As Long, _
                                  ByVal nFieldOrder As Integer) As Integer
'---------------------------------------------------------------------
'   Copies a CRF element
'   Note: only used when copying a CRF page (therefore, don't need to re-check field order)
'   TA  27/07/2001: CRFElement.Local changed to CRFElement/LocalFlag (JET4 prob)
'   NCJ 4 Feb 2002 - Must include new QGroup fields when copying
'   REM 22/02/02 - Added lQGroup, lOwnerQGroup variables
'   ZA 17/07/2002 - Added CaptionFontName, CaptionFontBold, CaptionFontItalic,
'                 CaptionFontSize, CaptionFontColour
'   REM 10/09/02 - Added ElementUse field
'   NCJ 11 Nov 02 - Added DisplayLength and Hotlink
'   NCJ 15 Jan 04 - Made Private, and added nFieldOrder
'---------------------------------------------------------------------
Dim nNextCRFElementId As Integer
Dim sSQL As String

    On Error GoTo ErrLabel

    nNextCRFElementId = mnNextCRFElementId(lToClinicalTrialId, nToVersionId, lToCRFPageId)
    
    ' NCJ 14 Feb 00 (related to SR2954?) Added Mandatory and RoleCode values
    'Mo Morris 30/8/01 Db Audit (RequireComment added)
    ' NCJ 4 Feb 02 - Added QGroupID, OwnerQGroupId, QGroupFieldOrder, ShowStatusFlag
    ' REM 22/02/02 - Added lQGroup, lOwnerQGroup variables
    ' ZA 17/07/2002 - Added CaptionFontName, CaptionFontBold, CaptionFontItalic,
    '                 CaptionFontSize, CaptionFontColour
    'REM 10/09/02 - Added ElementUse field
    'NCJ 11 Nov 02 - Added DisplayLength and Hotlink fields
    'ASH 12 FEB 03 - Added CLINICALTESTDATEEXPR field
    sSQL = "INSERT INTO CRFElement ( ClinicalTrialId, VersionId," _
        & " CRFPageId, CRFElementId, DataItemId, ControlType," _
        & " X, Y, CaptionX, CaptionY," _
        & " FontColour, Caption, FontName," _
        & " FontBold, FontItalic, FontSize," _
        & " FieldOrder, SkipCondition," _
        & " Optional, Hidden, LocalFlag," _
        & " Mandatory, RoleCode, RequireComment," _
        & " OwnerQGroupID, QGroupId, QGroupFieldOrder, ShowStatusFlag," _
        & " CaptionFontName, CaptionFontBold, CaptionFontItalic," _
        & " CaptionFontSize, CaptionFontColour, ElementUse, " _
        & " DisplayLength, Hotlink,CLINICALTESTDATEEXPR,DESCRIPTION) "
    ' NCJ 15 Jan 04 - Replaced CRFElement.FieldOrder with nFieldOrder
    sSQL = sSQL & " SELECT " & lToClinicalTrialId & ", " _
        & nToVersionId & ", " & lToCRFPageId & ", " & nNextCRFElementId _
        & ", " & lNewDataItemId & ", CRFElement.ControlType," _
        & " CRFElement.X, CRFElement.Y,CRFElement.CaptionX, CRFElement.CaptionY," _
        & " CRFElement.FontColour, CRFElement.Caption, CRFElement.FontName," _
        & " CRFElement.FontBold, CRFElement.FontItalic, CRFElement.FontSize, " _
        & nFieldOrder & ", CRFElement.SkipCondition," _
        & " CRFElement.Optional, CRFElement.Hidden, CRFElement.LocalFlag," _
        & " CRFElement.Mandatory, CRFElement.RoleCode, CRFElement.RequireComment, " _
        & lOwnerQGroupId & ", " & lQGroupId & ", CRFElement.QGroupFieldOrder, CRFElement.ShowStatusFlag, " _
        & " CRFElement.CaptionFontName, CRFElement.CaptionFontBold, " _
        & " CRFElement.CaptionFontItalic, CRFElement.CaptionFontSize, CRFElement.CaptionFontColour, CRFElement.ElementUse, " _
        & " CRFElement.DisplayLength, CRFElement.Hotlink, CRFElement.CLINICALTESTDATEEXPR,CRFELEMENT.DESCRIPTION" _
        & " FROM CRFElement "
    sSQL = sSQL _
        & " WHERE CRFElement.ClinicalTrialId = " & lFromClinicalTrialId _
        & " AND CRFElement.VersionId = " & nFromVersionId _
        & " AND CRFElement.CRFPageId = " & lFromCRFPageId _
        & " AND CRFElement.CRFElementId = " & nFromCRFElementId
    
    MacroADODBConnection.Execute sSQL
        
    mnCopyCRFElement = nNextCRFElementId
       
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.mnCopyCRFElement"
    
End Function

'---------------------------------------------------------------------
Public Function CopyCRFPage(ByVal lCopyFromClinicalTrialId As Long, _
                            ByVal nCopyFromVersionId As Integer, _
                            ByVal lCopyToClinicalTrialId As Long, _
                            ByVal nCopyToVersionId As Integer, _
                            ByVal lCopyFromCRFPageId As Long) As Long
'---------------------------------------------------------------------
' Returns Long instead of Integer - NCJ 13 Dec 99
' NCJ 18/1/00 - SR2746 fix
' TA 28/03/2000 - check for visit,question, study with same code
' NCJ 19/2/01 - SR4189 Only save CLM guideline at the end of the whole shebang
' REM 14/06/02 - CBB 2.2.15 No. 16 changed message that appears if CRFPage code is a reserved word
' NCJ 27 Mar 03 - Added EFormWidth to copied fields
' MLM 28/06/05: bug 2544: tidied up error handling/transactions/gbDoCLMSave
'---------------------------------------------------------------------
Dim rsCRFElement As ADODB.Recordset
Dim rsCRFPage As ADODB.Recordset
Dim rsQGroups As ADODB.Recordset
Dim rsGroupQuestions As ADODB.Recordset
Dim lNewDataItemId As Long
Dim lCRFPageId As Long
Dim nNewCRFElementId As Integer
Dim sCRFPageCode As String
Dim sSQL As String
Dim nNewCRFPageOrder As Integer
Dim sMSG As String
Dim sPrompt As String
Dim sMsgBoxTitle As String
Dim nErr As Integer
Dim sErrDesc As String

    On Error GoTo ErrHandler

    Set rsCRFPage = gdsCRFPage(lCopyFromClinicalTrialId, nCopyFromVersionId, lCopyFromCRFPageId)
    
    'REM 13/05/02 - check that recordset contains data
    If rsCRFPage.RecordCount = 0 Then 'if not then pop up message box to inform user that the eForm no longer exists
        Call DialogWarning("The eForm being copied has been deleted from the database.")
        'call a refresh
        Call frmMenu.RefreshQuestionLists(lCopyFromClinicalTrialId)
        'then exit function
        Exit Function
    Else
        sCRFPageCode = rsCRFPage!CRFPageCode
        rsCRFPage.Close
        Set rsCRFPage = Nothing
    End If

    sMsgBoxTitle = "Copying eForm - " & sCRFPageCode
        
    ' NCJ 19/2/01 - Switch off intermediate guideline saving
    gbDoCLMSave = False
    
    'Begin transaction
    TransBegin
    'REM 13/05/02 - added new error handler for rollback in a transaction
    On Error GoTo ErrRollback
    
    lCRFPageId = DuplicateCRFPage(lCopyFromClinicalTrialId, nCopyFromVersionId, _
        lCopyFromCRFPageId, lCopyToClinicalTrialId, nCopyToVersionId, sMsgBoxTitle, sCRFPageCode)
    If lCRFPageId = 0 Then
        'cancel
        CopyCRFPage = lCopyFromCRFPageId
        TransRollBack
        gbDoCLMSave = True
        Exit Function
    End If
    
    Set rsCRFElement = CRFPageElementList(lCopyFromClinicalTrialId, nCopyFromVersionId, lCopyFromCRFPageId)
    
    'initially set to false(it will be set in routine CopyDataItem)
    mblnOverwriteAllQuestions = False
    Do While Not rsCRFElement.EOF
   
        If RemoveNull(rsCRFElement!DataItemId) <> 0 Then
            'is a data item
            lNewDataItemId = CopyDataItem(lCopyFromClinicalTrialId, nCopyFromVersionId, _
                                                rsCRFElement!DataItemId, lCopyToClinicalTrialId, nCopyToVersionId)
        Else
            'is a graphical item
            lNewDataItemId = gnZERO
        End If
        
        'Data Item Id would be -1 if they cancelled the copy
        If lNewDataItemId >= gnZERO Then
            'add Questions and other elements to the CRFElement table
            ' NCJ 15 Jan 04 - Added in FieldOrder (OK to use existing FieldOrders when copying entire eForm)
            nNewCRFElementId = mnCopyCRFElement(lCopyFromClinicalTrialId, nCopyFromVersionId, _
                                            lCopyFromCRFPageId, rsCRFElement!CRFelementID, _
                                            lCopyToClinicalTrialId, nCopyToVersionId, _
                                            lCRFPageId, lNewDataItemId, _
                                            rsCRFElement!QGroupID, rsCRFElement!OwnerQGroupID, _
                                            rsCRFElement!FieldOrder)
        End If
        
       ' Removed dummy calls to InsertProformaCRFElement - NCJ 8 Sep 9
        rsCRFElement.MoveNext
    Loop
     'set to false ready for next
     mblnOverwriteAllQuestions = False
     
    'REM 20/02/02
    'Copy the Question Groups on an EForm
    Call CopyEFormQGroups(lCopyFromClinicalTrialId, nCopyFromVersionId, lCopyFromCRFPageId, _
                          lCopyToClinicalTrialId, nCopyToVersionId, lCRFPageId)
     
    ' NCJ 19/2/01 - Save guideline and switch guideline saving back on
    gbDoCLMSave = True
    Call SaveCLMGuideline
    
    'End transaction
    TransCommit
    
    On Error GoTo ErrHandler
    
    rsCRFElement.Close
    Set rsCRFElement = Nothing
    
    CopyCRFPage = lCRFPageId
       
Exit Function
    
ErrRollback:
    nErr = Err.Number
    sErrDesc = Err.Description & "|modSDTrialData.CopyCRFPage"
    'RollBack transaction
    TransRollBack
    Err.Raise nErr, , sErrDesc
Exit Function

ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.CopyCRFPage"
End Function

'---------------------------------------------------------------------
Public Function DuplicateCRFPage(ByRef lFromStudyId As Long, _
                            ByRef nFromVersionId As Integer, _
                            ByRef lFromEFormId As Long, _
                            ByRef lToStudyId As Long, _
                            ByRef nToVersionId As Integer, _
                            Optional ByRef sTitle As String = "", _
                            Optional ByRef sToEFormCode As String = "") As Long
'---------------------------------------------------------------------
' MLM 31/03/03: Created based on CopyCRFPage. Copy just the row in CRFPage.
' Returns the CRFPageId of the new form, unless the user cancels or the destination
' study already contains 255 forms, in which case 0 is returned.
'---------------------------------------------------------------------

Dim sMSG As String
Dim lNextCRFPageId As Long
Dim nNewCRFPageOrder As Integer
Dim sSQL As String
    
    DuplicateCRFPage = 0
    
    If EFormCount(lToStudyId, nToVersionId) = 255 Then
        DialogInformation "You can not create more than 255 eForms in a study"
        Exit Function
    End If
    If sToEFormCode <> "" And Not ValidateItemCode(sToEFormCode, gsITEM_TYPE_EFORM, sMSG, False) Then
        If vbNo = DialogQuestion(sMSG & vbCrLf & "Would you like to rename the eForm?", sTitle) Then
            Exit Function
        Else
            sToEFormCode = GetItemCode(gsITEM_TYPE_EFORM, "New " & gsITEM_TYPE_EFORM & " code:", sToEFormCode)
        End If
    End If
    If sToEFormCode = "" Then
        sToEFormCode = GetItemCode(gsITEM_TYPE_EFORM, "New " & gsITEM_TYPE_EFORM & " code:", sToEFormCode)
    End If
    If sToEFormCode = "" Then
        Exit Function
    End If

    ' lNextCRFPageId = gnNewDataTag
    ' Create new CLM plan for this CRFPage - NCJ 11/8/99
    lNextCRFPageId = gnNewCLMPlan(gsCLMCRFName(sToEFormCode))
    InsertProformaCRFPage lToStudyId, lNextCRFPageId, sToEFormCode

    '   ATN 3/5/99  SR 886
    '   Use new CRF page code (may be different from the copied one)
    'Changed Mo Morris 10/9/99
    'LocalCRFPageLabel, SequentialEntry, CRFPageDateLabel, DisplayNumbers added
    ' NCJ 18/1/00 SR2746 - Generate new CRFPageOrder instead of copying old one
    'Mo Morris 30/8/01 Db Audit (HideIfInactive and eFormDatePrompt added)
    ' NCJ 27 Mar 03 - Added eFormWidth
    nNewCRFPageOrder = mnNextCRFPageOrder(lToStudyId, nToVersionId)
    sSQL = "INSERT INTO CRFPage (  ClinicalTrialId, VersionId, " _
            & "CRFPageId, CRFTitle,  BackgroundColour, CRFPageOrder, CRFPageCode, " _
            & "CopiedFromClinicalTrialId, CopiedFromVersionId, CopiedFromCRFPageId, CRFPageLabel, " _
            & "LocalCRFPageLabel, SequentialEntry, CRFPageDateLabel, DisplayNumbers, " _
            & "HideIfInactive, eFormDatePrompt, eFormWidth ) " _
            & " SELECT " & lToStudyId & "," _
            & nToVersionId & "," & lNextCRFPageId & ",CRFPage.CRFTitle, " _
            & " CRFPage.BackgroundColour, " & nNewCRFPageOrder & ", '" & sToEFormCode & "', " _
            & lFromStudyId & "," & nFromVersionId & "," & lFromEFormId _
            & ",CRFPageLabel, LocalCRFPageLabel, SequentialEntry, CRFPageDateLabel, DisplayNumbers, " _
            & " HideIfInactive, eFormDatePrompt, eFormWidth " _
            & " FROM    CRFPage " _
            & " WHERE   ClinicalTrialId             = " & lFromStudyId _
            & " AND     VersionId                   = " & nFromVersionId _
            & " AND     CRFPageId                   = " & lFromEFormId
            
    MacroADODBConnection.Execute sSQL
    
    DuplicateCRFPage = lNextCRFPageId

End Function

'-------------------------------------------------------------------------------
Private Function EFormCount(lStudyId As Long, nVersionId As Integer) As Integer
'-------------------------------------------------------------------------------
' MLM 31/03/03: Created. Return the number of eForms in the specified study.
'-------------------------------------------------------------------------------

Dim sSQL As String
Dim rsEForms As ADODB.Recordset

    sSQL = "SELECT COUNT(*) FROM CRFPage WHERE ClinicalTrialId = " & lStudyId & _
        " AND VersionId = " & nVersionId
    Set rsEForms = MacroADODBConnection.Execute(sSQL)
    If rsEForms.EOF Then
        EFormCount = 0
    Else
        EFormCount = rsEForms.Fields(0).Value
    End If

End Function

'-------------------------------------------------------------------------------
Private Sub CopyEFormQGroups(ByVal lCopyFromClinicalTrialId As Long, ByVal nCopyFromVersionId As Integer, _
                            ByVal lCopyFromCRFPageId As Long, ByVal lCopyToClinicalTrialId As Long, _
                            ByVal nCopyToVersionId As Integer, ByVal lNextCRFPageId As Long)
'-------------------------------------------------------------------------------
' REM 28/02/02
' Copying the questions groups during the copying of an EForm
' NCJ 15 Jan 04 - Made Private; added FieldOrders to element copy
'-------------------------------------------------------------------------------
Dim rsQGroups As ADODB.Recordset
Dim rsGroupQuestions As ADODB.Recordset
Dim lNewDataItemId As Long
Dim lNewQGroupId As Long
Dim lNewOwnerQGroupId As Long
Dim nNewCRFElementId As Integer
Dim oQG As QuestionGroup
Dim nGroupFieldOrder As Integer

    On Error GoTo ErrLabel
    
    'Returns a recordset of all the question groups on the EForm being copied
    Set rsQGroups = QGroupList(lCopyFromClinicalTrialId, nCopyFromVersionId, lCopyFromCRFPageId)
    
    'Loop through the QGroups
    Do While Not rsQGroups.EOF

        'Copys the Question Group into QGroup Table, but first checks to see if the
        'QGroupcode already exists, and if so then asks; Rename or Skip
        Set oQG = CopyQGroup(lCopyFromClinicalTrialId, nCopyFromVersionId, _
                                   rsQGroups!QGroupID, lCopyToClinicalTrialId, nCopyToVersionId)
 
        If Not oQG Is Nothing Then 'if oQG is nothing then skipped QGroup
            oQG.Store

            'Returns a recordset of all the questions in a specific QGroup on a specific CRFPage
            Set rsGroupQuestions = CRFElementGroupQuestionList(lCopyFromClinicalTrialId, nCopyFromVersionId, _
                                                               lCopyFromCRFPageId, rsQGroups!QGroupID)
                                                     
            'Loop through the QGroup Questions
            Do While Not rsGroupQuestions.EOF
                
                'adds the group questions to the DataItem table
                lNewDataItemId = CopyDataItem(lCopyFromClinicalTrialId, nCopyFromVersionId, _
                                                rsGroupQuestions!DataItemId, lCopyToClinicalTrialId, _
                                                nCopyToVersionId, , False)
                
                If lNewDataItemId > -1 Then 'If -1 then skipped question, don't add it, else
                    'add the question to the QGroup object
                    oQG.AddQuestion (lNewDataItemId)
                    
                    'QGroup questions have an OwnerQgroupId but no QGroupId in the CRFElement table
                    lNewOwnerQGroupId = oQG.QGroupID
                    lNewQGroupId = 0

                    'add Group Question to the CRFElement table
                    ' NCJ 15 Jan 04 - Added Fieldorder
                    nNewCRFElementId = mnCopyCRFElement(lCopyFromClinicalTrialId, nCopyFromVersionId, _
                                            lCopyFromCRFPageId, rsGroupQuestions!CRFelementID, _
                                            lCopyToClinicalTrialId, nCopyToVersionId, _
                                            lNextCRFPageId, lNewDataItemId, _
                                            lNewQGroupId, lNewOwnerQGroupId, rsGroupQuestions!FieldOrder)
                End If
                rsGroupQuestions.MoveNext
            Loop
            'added all questions so save the group
            oQG.Save
            
            'A QGroup has a QGroupId and no OwnerQGroupId in the CRFElement table
            lNewQGroupId = oQG.QGroupID
            lNewOwnerQGroupId = 0
                                   
            'QGroups do not have DataItem Id's
            lNewDataItemId = 0

            'add Question Group to the CRFElement table
            'NB This must be done AFTER copying the group questions
            ' NCJ 15 Jan 04 - Added Fieldorder
            nNewCRFElementId = mnCopyCRFElement(lCopyFromClinicalTrialId, nCopyFromVersionId, _
                                    lCopyFromCRFPageId, rsQGroups!CRFelementID, _
                                    lCopyToClinicalTrialId, nCopyToVersionId, _
                                    lNextCRFPageId, lNewDataItemId, lNewQGroupId, _
                                    lNewOwnerQGroupId, rsQGroups!FieldOrder)

            'add the QGroup to the EFormQGroup table
            Call EFormQGroupUpdate(lCopyFromClinicalTrialId, nCopyFromVersionId, lCopyFromCRFPageId, _
                                   rsQGroups!QGroupID, lCopyToClinicalTrialId, nCopyToVersionId, _
                                   lNextCRFPageId, lNewQGroupId)
        End If
        rsQGroups.MoveNext
        
    Loop
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.CopyEFormQGroups"
End Sub

'-------------------------------------------------------------------------------
Public Function CopyingQGroup(ByVal lCopyFromClinicalTrialId As Long, ByVal nCopyFromVersionId As Integer, _
                         ByVal lCopyFromQGroupId As Long, ByVal lCopyToClinicalTrialId As Long, _
                         ByVal nCopyToVersionId As Integer) As Long
'-------------------------------------------------------------------------------
' REM 28/02/02
' Copying just the Question Group from one study to another. Question group can be on
' an EForm or in the Unused question group list
'-------------------------------------------------------------------------------
Dim oQG As QuestionGroup
Dim rsQGroupQuestions As ADODB.Recordset
Dim lNewDataItemId As Long

    On Error GoTo ErrLabel
    
    TransBegin
    
    'Copys the Question Group into QGroupTable, but first checks to see if the
    'QGroupcode already exists, and if so then asks; Rename or Skip
    Set oQG = CopyQGroup(lCopyFromClinicalTrialId, nCopyFromVersionId, _
                               lCopyFromQGroupId, lCopyToClinicalTrialId, nCopyToVersionId)

    If Not oQG Is Nothing Then 'if oQG is nothing then skipped QGroup
        'store the group questions as they currently are
        oQG.Store
        
        'Switch guideline saving off until we have finished
        gbDoCLMSave = False
        
        'Returns a recordset of all the questions in a specific QGroup
        Set rsQGroupQuestions = QGroupQuestionList(lCopyFromClinicalTrialId, nCopyFromVersionId, _
                                                   lCopyFromQGroupId)
                                                 
        'Loop through the QGroup Questions
        Do While Not rsQGroupQuestions.EOF
            'adds the group questions to the DataItem table
            lNewDataItemId = CopyDataItem(lCopyFromClinicalTrialId, nCopyFromVersionId, _
                                            rsQGroupQuestions!DataItemId, lCopyToClinicalTrialId, _
                                            nCopyToVersionId, , False)
            
            If lNewDataItemId > -1 Then 'If -1 then skipped question, don't add it, else
                'add the question to the QGroup object
                oQG.AddQuestion (lNewDataItemId)
                
            End If
            rsQGroupQuestions.MoveNext
        Loop
        'added all questions so save the group
        oQG.Save
        'Save guideline and switch guideline saving back on
        gbDoCLMSave = True
        Call SaveCLMGuideline
    End If
    
    TransCommit
    
    If oQG Is Nothing Then
        'Canceled copy
        CopyingQGroup = -1
    Else
        CopyingQGroup = oQG.QGroupID
    End If
    
Exit Function
ErrLabel:
    gbDoCLMSave = True

    Err.Raise Err.Number, , Err.Description & "|modSDTrialData.CopyQGroup"
End Function


'-------------------------------------------------------------------------------
Public Function NewClinicalTrialID() As Long
'-------------------------------------------------------------------------------
'Creates new clinical trial ID
'-------------------------------------------------------------------------------
Dim sSQL As String
Dim rsTrial As ADODB.Recordset

    sSQL = "SELECT max(ClinicalTrialId) as MaxClinicalTrialId FROM ClinicalTrial "
    Set rsTrial = New ADODB.Recordset
    rsTrial.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                        
    If IsNull(rsTrial!MaxClinicalTrialId) Then
        NewClinicalTrialID = 1
    Else
        NewClinicalTrialID = rsTrial!MaxClinicalTrialId + 1
    End If
    
    rsTrial.Close
    Set rsTrial = Nothing

End Function
