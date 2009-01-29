Attribute VB_Name = "basDataDefinition"
'----------------------------------------------------------------------------------------'
'   File:       basDataDefintion.bas
'   Module:     basDataDefinition (used to be modDataITemSQL)
'   Copyright:  InferMed Ltd. 1998-2001. All Rights Reserved
'   Author:     Andrew Newbigging
'   Purpose:    Module
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   1-8    Andrew Newbigging     4/06/97 - 27/08/98
'   9    Joanne Lau          07/09/98           Details:
'                                               Changes to gdsUpdateDataItem.
'                                               Replaced the Update table statements
'                                               used to update table DataItem by using
'                                               the edit recordset method. This prevents
'                                               SQLServer crashing when updating a column
'                                               with a long text string. Bug no 424.
'   10  Mo Morris           24/9/98             SPR 433
'                                               Additional changes to gdsUpdateDataItem. All
'                                               checking of variables not being null before their update
'                                               have been removed, because it prevented the saving
'                                               of variables that had been edited to a null.
'   11   Andrew Newbigging       11/11/98
'       Following routines moved form this module to TrialData module:
'           CopyDataItem,gdsDataItem,DataItemExists,gnNextDataItemId
'   19      Paul Norris             05/08/99
'           Amended gdsUpdateDataItem() routine for MTM v2.0
'   20      NCJ 12/8/99
'           Removed call to DeleteProformaDataItem
'   21      Paul Norris             17/08/99    SR 1501, 1540
'           Added GetUniqueDataItemCode(), ValidateDataItemCode() and ValidateDataItemExportName()
'   22      Paul Norris          23/-8/99   SR 1712
'           Amended gdsCRFPageList(),gdsDataList() to include bFormAlphabeticOrder parameter
'   23      Paul Norris          13/09/99    Field name DataItem.Case changed to DataItem.DataItemCase
'   PN  17/09/99    Amended gdsUpdateDataItem() to include library specific fields
'                   Required and TrialTypeId
'   PN  24/09/99    Moved gdsDataValues() to TrialData module
'  WillC    10/11/99 Added the Error handlers
'   Mo Morris   12/11/99    DAO to ADO conversion
'   NCJ 13 Dec 99   Ids to Long
'   NCJ 27 Jan 00 - Tidying validation of question codes
'   Mo Morris   15/2/00     lClinicalTrialId and lVersionId removed as calling arguments
'                           from ValidateDataItemCode
' MACRO 3.0
'   NCJ 5 Dec 01 - Added QuestionGroup table to gdsDeleteDataItem
'                  Changed to new-style error handlers
'   NCJ 11 Dec 01 - Added DataItemIdsNotAllowedInQGroup
'   REM 12/12/01 - Added new routines for RQG's: QuestionList, IsQuestionOnCRFPageOrQGroup, AllQuestionList, CRFPagesList, QuestionGroupList
'       CRFPageQGroupList, UnusedQuestionList, QGroupNotOnEForm, QGroups, QGroupQuestionNotOnEForm, BuildCRFPages, QuestionNames,
'       GroupNames
'   ic 15/06/2005 added clinical coding
'----------------------------------------------------------------------------------------'

Option Explicit

'---------------------------------------------------------------------
Public Function DataItemIdsNotAllowedInQGroup(ByVal lClinicalTrialId As Long, _
                                   ByVal nVersionId As Integer, _
                                   ByVal lGroupId As Long) As Collection
'---------------------------------------------------------------------
' Return a collection of DataItemIds which aren't allowed in the given group
' because they're already on an eForm with the group
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsCRFPages As ADODB.Recordset
Dim rsDataOnPage As ADODB.Recordset
Dim nErr As Integer
Dim sErrDesc As String
Dim colDataItemIDs As Collection

    On Error GoTo ErrLabel
    
    Set colDataItemIDs = New Collection
    
    ' Find all the CRFPages which use this group
    sSQL = "SELECT CRFPageID from EFormQGroup " _
            & " WHERE ClinicalTrialId = " & lClinicalTrialId _
            & " AND  VersionId    = " & nVersionId _
            & " AND QGroupId = " & lGroupId
    Set rsCRFPages = New ADODB.Recordset
    rsCRFPages.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set rsDataOnPage = New ADODB.Recordset
    Do While Not rsCRFPages.EOF
        ' Get the data item IDs on this page
        sSQL = "SELECT DataItemId FROM CRFElement " _
            & " WHERE ClinicalTrialId = " & lClinicalTrialId _
            & " AND VersionId = " & nVersionId _
            & " AND CRFPageId = " & rsCRFPages!CRFPageId _
            & " AND DataItemId > 0 "
        rsDataOnPage.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsDataOnPage.EOF
            On Error Resume Next
            ' Add each DataItemId to the collection (ignore duplicates)
            colDataItemIDs.Add CLng(rsDataOnPage!DataItemId), Str(rsDataOnPage!DataItemId)
            ' Next Data item id
            rsDataOnPage.MoveNext
        Loop
        On Error GoTo ErrLabel
        rsDataOnPage.Close
        ' Next CRF Page
        rsCRFPages.MoveNext
    Loop
    
    Set rsDataOnPage = Nothing
    
    rsCRFPages.Close
    Set rsCRFPages = Nothing
            
    Set DataItemIdsNotAllowedInQGroup = colDataItemIDs
    Set colDataItemIDs = Nothing
    
Exit Function
ErrLabel:
    nErr = Err.Number
    sErrDesc = Err.Description & "|DataDefinition.DataItemIdsNotAllowedInQGroup"
    'RollBack transaction
    TransRollBack
    Err.Raise nErr, , sErrDesc

End Function

'---------------------------------------------------------------------
Public Function QuestionList(ClinicalTrialId As Long, VersionId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
' REM 19/12/01
' Returns a list of all the Questions in a Study, excluding the ones that are exclusively in Question Groups
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel
    
    sSQL = "SELECT DataItemId, DataItemName,DataItem.DataItemCode" _
        & " FROM DataItem " _
        & " WHERE DataItemID NOT IN (SELECT DataItemID" _
        & " FROM QGroupQuestion WHERE QGroupQuestion.ClinicalTrialID = " & ClinicalTrialId _
        & " AND QGroupQuestion.VersionID = " & VersionId & ")" _
        & " AND ClinicalTrialId = " & ClinicalTrialId _
        & " AND VersionId = " & VersionId _
        & " ORDER BY DataItemName"
        
    Set QuestionList = New ADODB.Recordset
    QuestionList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|DataDefinition.QuestionList"
    
End Function

'---------------------------------------------------------------------
Public Function IsQuestionOnCRFPageOrQGroup(lClinicalTrialId As Long, nVersionId As Integer, lDataItemId As Long) As Boolean
'---------------------------------------------------------------------
' REM 20/12/01
' Checks to see if a specific question exist on any CRFPages or in any Question Groups
'REVISIONS:
'REM 30/01/02 - Changed SQL for checking specific question exits on anyEForms or in any question groups
'---------------------------------------------------------------------
Dim sSQL As String
Dim sSQL2 As String
Dim rsQuestion As ADODB.Recordset
Dim nCRFelement As Integer
Dim nQGroupQuestion As Integer

    On Error GoTo ErrLabel
    
    'Check to see if the question is on any other EForms
    sSQL = "SELECT COUNT(*)" & _
            " FROM CRFElement" & _
            " WHERE CRFElement.ClinicalTrialId = " & lClinicalTrialId & _
            " AND CRFElement.VersionId = " & nVersionId & _
            " AND CRFElement.DataItemId = " & lDataItemId
    Set rsQuestion = New ADODB.Recordset
    rsQuestion.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

    nCRFelement = rsQuestion.Fields(0).Value

    'Checks to see if the question is in any Question Groups
    sSQL2 = "SELECT COUNT(*)" & _
            " FROM QGroupQuestion" & _
            " WHERE QGroupQuestion.ClinicalTrialId = " & lClinicalTrialId & _
            " AND QGroupQuestion.VersionId = " & nVersionId & _
            " AND QGroupQuestion.DataItemId = " & lDataItemId
    Set rsQuestion = New ADODB.Recordset
    rsQuestion.Open sSQL2, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

    nQGroupQuestion = rsQuestion.Fields(0).Value

    'If the question is either on an eform or in a question group it returns true
    If (nCRFelement = 0) And (nQGroupQuestion = 0) Then
        IsQuestionOnCRFPageOrQGroup = False
    Else
        IsQuestionOnCRFPageOrQGroup = True
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|DataDefinition.IsQuestionOnCRFPageOrQGroup"
    
End Function

'---------------------------------------------------------------------
Public Function AllQuestionList(ClinicalTrialId As Long, VersionId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
' REM 19/12/01
' Gets all questions in a study
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel

        sSQL = "SELECT Distinct DataItemId, DataItemName, DataItem.DataItemCode" _
        & " FROM DataItem " _
        & " WHERE ClinicalTrialId = " & ClinicalTrialId _
        & " AND VersionId = " & VersionId _
        & " ORDER BY DataItemName"
        
    Set AllQuestionList = New ADODB.Recordset
    AllQuestionList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|DataDefinition.AllQuestionList"
End Function

'---------------------------------------------------------------------
Public Function CRFPagesList(ClinicalTrialId As Long, VersionId As Integer, _
                                  bFormAlphabeticOrder As Boolean) As ADODB.Recordset
'---------------------------------------------------------------------
' REM 18/12/01
' Get a list of all CRFPages in a study
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel

    sSQL = "SELECT CRFPage.CRFPageId, CRFPage.CRFTitle, CRFPage.CRFPageCode, CRFElement.CRFElementId, "
    sSQL = sSQL & "CRFElement.FieldOrder, DataItem.DataItemId, DataItem.DataItemName,DataItem.DataItemCode "
    sSQL = sSQL & "FROM CRFPage, DataItem, CRFElement "
    sSQL = sSQL & "WHERE CRFPage.ClinicalTrialId = CRFElement.ClinicalTrialId "
    sSQL = sSQL & "AND CRFElement.ClinicalTrialId  = DataItem.ClinicalTrialId "
    sSQL = sSQL & "AND CRFPage.VersionId = CRFElement.VersionId "
    sSQL = sSQL & "AND CRFElement.VersionId = DataItem.VersionId "
    sSQL = sSQL & "AND CRFElement.DataItemId =  DataItem.DataItemId "
    sSQL = sSQL & "AND CRFPage.CRFPageId = CRFElement.CRFPageId "
    sSQL = sSQL & "AND CRFElement.OwnerQGroupID = 0"
    sSQL = sSQL & "AND CRFPage.ClinicalTrialId = " & ClinicalTrialId
    sSQL = sSQL & "AND CRFPage.VersionId = " & VersionId

    If bFormAlphabeticOrder Then
        sSQL = sSQL & " ORDER BY DataItem.DataItemName,DataItem.DataItemId, CRFPage.CRFTitle, CRFElement.FieldOrder"
    Else
        sSQL = sSQL & " ORDER BY DataItem.DataItemName,DataItem.DataItemId, CRFPage.CRFPageOrder, CRFElement.FieldOrder"
    End If

    Set CRFPagesList = New ADODB.Recordset
    CRFPagesList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|DataDefinition.CRFPagesList"

End Function

'---------------------------------------------------------------------
Public Function QuestionGroupList(ClinicalTrialId As Long, VersionId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
' REM 19/12/01
' Returns a list of all question groups in the study
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel

    sSQL = "SELECT QGroup.QGroupCode, QGroup.QGroupName, QgroupQuestion.QGroupID, QGroupQuestion.DataItemID" & _
            " FROM QGroup, QGroupQuestion" & _
            " WHERE QGroup.ClinicalTrialId = QGroupQuestion.ClinicalTrialId" & _
            " AND QGroup.VersionId = QGroupQuestion.VersionId" & _
            " AND QGroup.QGroupID = QGroupQuestion.QGroupID" & _
            " AND QGroup.ClinicalTrialId = " & ClinicalTrialId & _
            " AND QGroup.VersionId = " & VersionId

    Set QuestionGroupList = New ADODB.Recordset
    QuestionGroupList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|DataDefinition.QuestionGroupList"
End Function

'---------------------------------------------------------------------
Public Function CRFPageQGroupList(ClinicalTrialId As Long, VersionId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
' REM 19/12/01
' Returns a recordset of all question groups and their associated questions
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel

    sSQL = "SELECT QgroupQuestion.QGroupID, QGroupQuestion.DataItemID, EFormQGroup.CRFPageID, CRFPAge.CRFPageCode, CRFPage.CRFTitle" & _
            " FROM EFormQGroup, QGroupQuestion, CRFPage" & _
            " WHERE EFormQGroup.ClinicalTrialId = QGroupQuestion.ClinicalTrialId" & _
            " AND QGroupQuestion.ClinicalTrialId = CRFPage.ClinicalTrialId" & _
            " AND EFormQGroup.VersionId = QGroupQuestion.VersionId" & _
            " AND QGroupQuestion.VersionId = CRFPage.VersionId" & _
            " AND EFormQGroup.CRFPageId = CRFPage.CRFPageId" & _
            " AND EFormQGroup.QGroupID = QGroupQuestion.QGroupID" & _
            " AND EFormQGroup.ClinicalTrialId = " & ClinicalTrialId & _
            " AND EFormQGroup.VersionId = " & VersionId

    Set CRFPageQGroupList = New ADODB.Recordset
    CRFPageQGroupList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|DataDefinition.CRFPageQGroupList"
End Function

'---------------------------------------------------------------------
Public Function UnusedQuestionList(ClinicalTrialId As Long, VersionId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
' REM 19/12/01
' Returns a recordset of all questions which are not on an EForm and not in a Question Group
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel
    
    sSQL = "SELECT DataItem.DataItemID, DataItem.DataItemName, DataItem.DataItemCode" & _
           " FROM DataItem" & _
           " WHERE DataItemID NOT IN (SELECT DataItemID" & _
           " FROM CRFElement WHERE CRFElement.ClinicalTrialID = " & ClinicalTrialId & _
           " AND CRFElement.VersionID = " & VersionId & ")" & _
           " AND DataItemID NOT IN (SELECT DataItemID" & _
           " FROM QGroupQuestion WHERE QGroupQuestion.ClinicalTrialID = " & ClinicalTrialId & _
           " AND QGroupQuestion.VersionID = " & VersionId & ")" & _
           " AND DataItem.ClinicalTrialID = " & ClinicalTrialId & _
           " AND DataItem.VersionID = " & VersionId & _
           " ORDER BY DataItem.DataItemName"
           
    Set UnusedQuestionList = New ADODB.Recordset
    UnusedQuestionList.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|DataDefinition.UnusedQuestionList"
    
End Function

'---------------------------------------------------------------------
Public Function QGroupNotOnEForm(ClinicalTrialId As Long, VersionId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
' REM 17/12/01
' Recordset of Question Groups in a study that are not currently on an EForm
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel
    
    sSQL = "SELECT QGroup.QGroupID, QGroup.QGroupCode, QGroup.QGroupName" & _
            " FROM QGroup" & _
            " WHERE QGroupID NOT IN (SELECT QGroupID" & _
            " FROM EformQGroup WHERE EformQGroup.ClinicalTrialID = " & ClinicalTrialId & _
            " AND EformQGroup.VersionID = " & VersionId & ")" & _
            " AND QGroup.ClinicalTrialID = " & ClinicalTrialId & _
            " AND QGroup.VersionID = " & VersionId & _
            " ORDER BY QGroup.QGroupName"
    
    Set QGroupNotOnEForm = New ADODB.Recordset
    QGroupNotOnEForm.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|DataDefinition.QGroupNotOnEForm"
    
End Function

'---------------------------------------------------------------------
Public Function QGroups(ClinicalTrialId As Long, VersionId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
' REM 17/12/01
' Recordset of all Question Groups in a study
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel
    
    sSQL = "SELECT QGroup.QGroupID, QGroup.QGroupCode, QGroup.QGroupName" & _
           " FROM QGroup" & _
           " WHERE QGroup.ClinicalTrialID = " & ClinicalTrialId & _
           " AND QGroup.VersionID = " & VersionId & _
           " ORDER BY QGroup.QGroupName"
            
    Set QGroups = New ADODB.Recordset
    QGroups.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|DataDefinition.QGroups"
End Function

'---------------------------------------------------------------------
Public Function QGroupQuestionNotOnEForm(ClinicalTrialId As Long, VersionId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
' REM 17/12/01
' Recordset of the groups questions for groups on on an EForm
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel

    sSQL = "SELECT QGroupQuestion.QGroupID, QGroupQuestion.DataItemID, QGroupQuestion.QOrder," & _
            " DataItem.DataItemCode, DataItem.DataItemName " & _
            " FROM QGroupQuestion, DataItem" & _
            " WHERE QGroupQuestion.QGroupID NOT IN" & _
            " (SELECT EFormQGroup.QGroupID FROM EformQGroup" & _
            " WHERE EformQGroup.ClinicalTrialID = " & ClinicalTrialId & _
            " AND EformQGroup.VersionID = " & VersionId & ")" & _
            " AND QGroupQuestion.DataItemId = DataItem.DataItemId" & _
            " AND QGroupQuestion.ClinicalTrialId = DataItem.ClinicalTrialId" & _
            " AND QGroupQuestion.VersionId = DataItem.VersionId" & _
            " AND QGroupQuestion.ClinicalTrialID = " & ClinicalTrialId & _
            " AND QGroupQuestion.VersionID = " & VersionId & _
            " ORDER BY QGroupQuestion.QOrder"

    Set QGroupQuestionNotOnEForm = New ADODB.Recordset
    QGroupQuestionNotOnEForm.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|DataDefinition.QGroupQuestionNotOnEForm"
End Function

'---------------------------------------------------------------------
Public Function BuildCRFPages(lClinicalTrialId As Long, nVersionId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
'REM 09/07/02
'Returns a recordset of all Questions, Qgroups and QGroupQuestions on a specific CRFPage
'---------------------------------------------------------------------
Dim sSQL As String

    sSQL = "SELECT CRFPageId, CRFElementId, DataItemId, QGroupId, OwnerQGroupId, FieldOrder" _
        & " FROM CRFElement" _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND VersionId = " & nVersionId _
        & " AND (DataItemId > 0 OR QGroupId > 0)" _
        & " ORDER BY CRFPageId, FieldOrder, QGroupFieldOrder"
        
    Set BuildCRFPages = New ADODB.Recordset
    BuildCRFPages.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.BuildCRFPages"
End Function

'---------------------------------------------------------------------
Public Function QuestionNames(lClinicalTrialId As Long, nVersionId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
'09/07/02
'returns a recordset of all the question names in a specific study
'---------------------------------------------------------------------
Dim sSQL As String

    sSQL = "SELECT DataItemId, DataItemCode, DataItemName" _
        & " FROM DataItem" _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND VersionId = " & nVersionId

    Set QuestionNames = New ADODB.Recordset
    QuestionNames.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.QuestionNames"
End Function

'---------------------------------------------------------------------
Public Function GroupNames(lClinicalTrialId As Long, nVersionId As Integer) As ADODB.Recordset
'---------------------------------------------------------------------
'REM 09/07/02
'Returns a recordset of all teh question group names
'---------------------------------------------------------------------
Dim sSQL As String

    sSQL = "SELECT QGroupId, QGroupCode, QGroupName" _
        & " FROM QGroup" _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND VersionId = " & nVersionId

    Set GroupNames = New ADODB.Recordset
    GroupNames.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataList.GroupNames"
End Function

'----------------------------------------------------------------------------------------'
Public Sub gdsUpdateDataItem(nClinicalTrialId As Long, _
                            nVersionId As Integer, nDataItemId As Long, _
                            sDataItemCode As String, sDataItemName As String, nDataType As Integer, _
                            nDataItemLength As Integer, sDataItemFormat As String, _
                            sUnitOfMeasurement As String, sDerivation As String, _
                            sDataItemHelpText As String, sCase As String, sExportName As String, _
                            nRequired As Integer, nTrialTypeId As Integer, bLibraryMode As Boolean, sClinicalTestCode As String, _
                            enMACROOnly As eMACROOnly, sDescription As String, nDictionaryId As Integer)
'----------------------------------------------------------------------------------------'
' Update a data definition with the values passed in
' PN 17/09/99 - updated procedure to conform to vb coding standards v1.0
'
'JL 07/09/98. Replaced the Update table statements used to update table DataItem by using the edit recordset
'method. This prevents SQLServer crashing when updating a column with a long text string.
'Bug 424.
' ZA 19/08/2002 - added MACROOnly and Description parameters
' ic 14/06/2005 added clinical coding
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsDataItem As ADODB.Recordset

    On Error GoTo ErrLabel
    
    sSQL = " Select * From DataItem " _
            & " WHERE ClinicalTrialId = " & nClinicalTrialId _
            & " AND VersionId = " & nVersionId _
            & " AND DataItemId = " & nDataItemId
    
    Set rsDataItem = New ADODB.Recordset
    rsDataItem.Open sSQL, MacroADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    
    'With rsDataItem
        '.Edit
        
        rsDataItem!DataItemName = sDataItemName
        rsDataItem!ExportName = sExportName
        rsDataItem!DataItemLength = nDataItemLength
        rsDataItem!DataType = nDataType
        If sDataItemFormat = "" Then    'See Q239781 before coding here
            rsDataItem!DataItemFormat = Null
        Else
            rsDataItem!DataItemFormat = sDataItemFormat
        End If
        If sDerivation = "" Then
            rsDataItem!Derivation = Null
        Else
            rsDataItem!Derivation = sDerivation
        End If
        If sDataItemHelpText = "" Then
            rsDataItem!DataItemHelpText = Null
        Else
            rsDataItem!DataItemHelpText = sDataItemHelpText
        End If
        If sUnitOfMeasurement = "" Then
            rsDataItem!UnitOfMeasurement = Null
        Else
            rsDataItem!UnitOfMeasurement = sUnitOfMeasurement
        End If
        
        'ZA 19/08/2002 - update MACROOnly and Description fields
        
        rsDataItem!MACROOnly = enMACROOnly
        rsDataItem!Description = ConvertToNull(sDescription, vbString)
        
        'TA 15/09/2000: update LabTestCode
        If sClinicalTestCode <> "" And nDataType = DataType.LabTest Then
            rsDataItem!ClinicalTestCode = sClinicalTestCode
        Else
            rsDataItem!ClinicalTestCode = Null
        End If
'        ' PN 17/09/99
'        ' added required and trialtypeid fields
'        If bLibraryMode Then
'            ' only save in library mode
'            rsDataItem!Required = nRequired
'            rsDataItem!RequiredTrialTypeID = nTrialTypeId
'
'        End If
'
        ' PN 13/09/99 field name Case changed to DataItemCase
        rsDataItem!DataItemCase = sCase
        
        If gbClinicalCoding Then
            'ic 15/06/2005 clinical coding: update chosen dictionary
            If (nDictionaryId > -1) Then
                rsDataItem!DictionaryId = nDictionaryId
            Else
                rsDataItem!DictionaryId = Null
            End If
        End If
        
        rsDataItem.Update
        
        rsDataItem.Close
    
    'End With
    Set rsDataItem = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|DataDefinition.gdsUpdateDataItem"
    
End Sub

'---------------------------------------------------------------------
Public Sub gdsDeleteDataItem(ClinicalTrialId As Long, _
                            VersionId As Integer, _
                            DataItemId As Long)
'---------------------------------------------------------------------
' Removed call to DeleteProformaDataItem (since done in frmDataList) NCJ 12/8/99
' NCJ 19/5/00 SR3487  Make sure entries are also removed from DataItemValidation
'---------------------------------------------------------------------
Dim sSQL As String
Dim sSQLWhere As String
Dim nErr As Integer
Dim sErrDesc As String

    On Error GoTo ErrLabel
    
    'Begin transaction
    TransBegin
    
    sSQLWhere = "WHERE ClinicalTrialId = " & ClinicalTrialId _
            & " AND VersionId = " & VersionId _
            & " AND DataItemId = " & DataItemId
    
    sSQL = "DELETE FROM DataItem " & sSQLWhere
    MacroADODBConnection.Execute sSQL
    
    sSQL = "DELETE FROM ValueData " & sSQLWhere
    MacroADODBConnection.Execute sSQL
    
    sSQL = "DELETE FROM CRFElement " & sSQLWhere
    MacroADODBConnection.Execute sSQL

    'SDM 16/12/99
    sSQL = "DELETE FROM RequiredData WHERE " & _
           "DataItemId = " & DataItemId

    MacroADODBConnection.Execute sSQL

    ' NCJ 19/5/00 SR3487
    ' Delete data from DataItemValidation table
    sSQL = "DELETE FROM DataItemValidation " & sSQLWhere
    MacroADODBConnection.Execute sSQL

    ' NCJ 5 Dec 01 - Delete from Question Group table too
    sSQL = "DELETE FROM QGroupQuestion " & sSQLWhere
    MacroADODBConnection.Execute sSQL
    
    'End transaction
    TransCommit

Exit Sub
ErrLabel:
    nErr = Err.Number
    sErrDesc = Err.Description
    
    'RollBack transaction
    TransRollBack
    
    Err.Raise nErr, , sErrDesc & "|DataDefinition.gdsDeleteDataItem"
    
End Sub

'---------------------------------------------------------------------
Public Function ValidateDataItemExportName(sName As String, _
                                        lClinicalTrialId As Long, _
                                        lVersionId As Integer, _
                                        sDataItemCode As String) As String
'---------------------------------------------------------------------
' Validate Export Code
' Returns error message if not valid, otherwise returns empty string
' NCJ 6 Mar 01 - SR 3471 (Revisit), Changed validation to allow underscore chars
'---------------------------------------------------------------------
Dim sMsg As String
Dim sExportCodes As String

    On Error GoTo ErrLabel

    'WillC SR3741 16/8/00
    sName = Trim(sName)
    
    sExportCodes = "Export codes "
    sMsg = ""
    
    ' NCJ 6/3/01 - New validation based on ValidateItemCode in modArezzoRebuild
    If sName = "" Then
        sMsg = sExportCodes & "cannot be blank."
    ElseIf Not gblnValidString(sName, valAlpha + valNumeric + valUnderscore) Then
        sMsg = sExportCodes & "can only contain alphanumeric characters."
    ElseIf Not gblnValidString(Left$(sName, 1), valAlpha) Then
        sMsg = sExportCodes & "must start with an alphabetic character."
    ElseIf Not gblnValidString(Right$(sName, 1), valAlpha + valNumeric) Then
        sMsg = sExportCodes & "must end with an alphanumeric character."
    ElseIf Len(sName) > 50 Then
        sMsg = sExportCodes & "cannot be more than 50 characters long."
    ElseIf gblnDataItemExportExists(lClinicalTrialId, lVersionId, sName, sDataItemCode) Then
        sMsg = "This export name is already in use. Please select another export name."
    End If

    ValidateDataItemExportName = sMsg

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|DataDefinition.ValidateDataItemExportName(" & sName & ")"
    
End Function



