VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   File:       DataBrowser.cls
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Toby Aldridge, May 2001
'   Purpose:    Data Browser class
'----------------------------------------------------------------------------------------'
' Revisions:
'   TA 02/10/2001: Filer on SiteUser (IT WOULD BE BETTER IF A LIST OF ALLOWED SITES WERE
'                    PASSED IN SO I WOULDN'T NEED TO DO THE EXTRA JOIN ON THE SITE USER TABLE.
'                    THIS SHOULD BE DONE WHEN WE HAVE DECIDED WHERE PERMSISSIONS AND PERMITTED
'                    SITES WILL BE STORED (MACRO OR SECURITY))
'   TA 11/01/02:    order by study name - bug list 2.2.7 bug 3
'   TA 24/06/2002   SR 4662: Order audit data by ResponseTimeStamp descending
'   RS 08/10/2002   Add Timezone to selected columns
'   TA 06/11/2002: Redesigned and StudySiteSQL and LockStatus parameters added
'   TA 31/03/2003: If no matches on subjectlabel then subject label is used as a person id
'   TA 27/05/2003: Added eFormLabel
'   DPH 08/10/2003 - Performance enhancements in GetQuestionQueryDef, GetAuditQueryDef
'   ic 28/07/2005   added clinical coding
'----------------------------------------------------------------------------------------'

Option Explicit

Private Const mALIASES As String = "ClinicalTrialId,ClinicalTrialName,TrialStatus,TrialLockStatus,SubjectDiscStatus,SubjectSDVStatus,SubjectNoteStatus,TrialSite,PersonId,PatientLabel,VisitName,VisitCycleNumber,VisitStatus,VisitLockStatus,VisitDiscStatus,VisitSDVStatus,VisitNoteStatus,CRFPageID,CRFTitle,CRFPageCycleNumber,FormStatus,FormLockStatus,FormDiscStatus,FormSDVStatus,FormNoteStatus,ResponseTimestamp,DataItemName,DataType,ResponseValue,ResponseStatus,UserName,Comments,DataItemLockStatus,DataItemDiscStatus,DataItemSDVStatus,DataItemNoteStatus,ReasonForChange,Changed,OverruleReason,LabResult,CTCGrade,VisitId,CRFPageTaskId,ResponseTaskId,CRFElementId,ValMessage,ResponseCycleNumber,ResponseTimestamp_TZ,QuestionId,ChangeCount,OwnerQGroupId,FullUserName,DatabaseTimeStamp,DatabaseTimeStmap_TZ,eFormLabel"

'ic 28/07/2005 added additional columns
Private Const mCCALIASES As String = ",DictionaryName,DictionaryVersion,CodingStatus,CodingDetails"
Private Const mCCSwitch As String = "CLINICALCODING"


Public Enum eDataBrowserType
    dbtDataItemResponse = 1
    dbtDataItemResponseHistory = 2
    dbteForms = 3
End Enum

'column number enumeration
Public Enum DataBrowserCol
    dbcStudyId = 0
    dbcStudyName = 1
    dbcsubjectStatus = 2
    dbcSubjectLockStatus = 3
    'new statuses
    dbcSubjectDiscStatus = 4
    dbcSubjectSDVStatus = 5
    dbcSubjectNoteStatus = 6
    
    dbcSite = 7
    dbcSubjectId = 8
    dbcSubjectLabel = 9
    dbcVisitName = 10
    dbcVisitCycleNumber = 11
    dbcVisitStatus = 12
    dbcVisitLockStatus = 13
    'new statuses
    dbcVisitDiscStatus = 14
    dbcVisitSDVStatus = 15
    dbcVisitNoteStatus = 16
    
    dbcEFormId = 17
    dbcEFormTitle = 18
    dbcEFormCycleNumber = 19
    dbcEFormStatus = 20
    dbcEFormLockStatus = 21
    
    'new statuses
    dbcEFormDiscStatus = 22
    dbcEFormSDVStatus = 23
    dbcEFormNoteStatus = 24
    
    dbcResponseTimeStamp = 25
'ones below here are not used in the eForm version
    dbcDataItemName = 26
    dbcDataType = 27
    dbcResponseValue = 28
    dbcResponseStatus = 29
    dbcUserName = 30
    dbcComments = 31
    dbcDataItemLockStatus = 32
    'new statuses
    dbcDataItemDiscStatus = 33
    dbcDataItemSDVStatus = 34
    dbcDataItemNoteStatus = 35
    
    dbcReasonForChange = 36
    dbcChanged = 37
    dbcOverruleReason = 38
    dbcLabResult = 39
    dbcCTCGrade = 40
    dbcVisitId = 41
    dbcEFormTaskID = 42
    dbcResponseTaskId = 43
    dbcEFormElementId = 44
    dbcValMessage = 45
    dbcResponseCycleNumber = 46
    dbcResponseTimestamp_TZ = 47
    dbcQuestionId = 48
    dbcChangeCount = 49
    dbcOwnerQGroupId = 50
    dbcFullUserName = 51
    
    'TA 26/03/2003: added database timestamps
    dbcDatabaseTimeStamp = 52
    dbcDatabaseTimestamp_TZ = 53
    
    'TA 27/05/2003: eFormLabel
    dbcEFormLabel = 54
    
    'ic 28/07/2005 added clinical coding fields
    dbcDictionaryName = 55
    dbcDictionaryVersion = 56
    dbcCodingStatus = 57
    dbcCodingDetails = 58
    
End Enum


'ic 18/07/2005 added clinical coding enumeration
Public Enum eCodingStatus
    csEmpty = 0
    csNotCoded = 1
    csCoded = 2
    csPendingNewCode = 3
    csAutoEncoded = 4
    csValidated = 5
    csDoNotCode = 6
    csNeverCoded = 7
End Enum

'-----------------------------------------------------
Public Function GetData(sConnect As String, bCountOnly As Boolean, _
                    enType As eDataBrowserType, sStudySiteSQL As String, bSingleSubject As Boolean, _
                    Optional nClinicalTrialId As Integer = -1, Optional sSite As String = "ALL", _
                    Optional lVisitId As Long = 0, Optional lCRFPageId As Long = 0, _
                    Optional sSubjectLabel As String = "", Optional lPersonId As Long = 0, _
                    Optional vStatus As Variant = "", Optional vLockStatus As Variant = "", _
                    Optional bBefore As Boolean = False, Optional dblDate As Double = 0, _
                    Optional lDataItemId As Long = 0, Optional sUserCode As String = "", _
                    Optional lComments As Long = -1, _
                    Optional lDiscStatus As Long = -1, Optional lSDVStatus As Long = -1, Optional lNotes As Long = -1, _
                    Optional lCodingStatus As Long = -1, Optional sDictionaryName As String = "", Optional sDictionaryVersion As String = "", _
                    Optional bGetSerialisedVersion As Boolean = False, Optional ByRef sSerialisedVersion As String = "") As Variant
'-----------------------------------------------------
'Return the data as an array.
'If bGetSerialisedVersion is true then sSerialiseVersion returns the serialised
'  recordset too for use by the ADO_DeSerialiseRecordset function in libADO.
'revisions
'ic 28/07/2005 added clinical coding
'-----------------------------------------------------
Dim rs As Recordset
Dim sSerialisedData As String
Dim oQDs As QueryDefs
Dim oQD As QueryDef
Dim oVersion As MACROVersion.Checker
Dim bCC As Boolean

    
    'check for clinical coding version
    Set oVersion = New MACROVersion.Checker
    bCC = oVersion.HasUpgrade(sConnect, mCCSwitch)


    'change lock status not passed in to null
    If Not IsNull(vLockStatus) Then
        If VarType(vLockStatus) = vbString Then
            If vLockStatus = "" Then
                vLockStatus = Null
            End If
        End If
    End If
    
    'change status not passed in to null
    If Not IsNull(vStatus) Then
        If VarType(vStatus) = vbString Then
            If vStatus = "" Then
                vStatus = Null
            End If
        End If
    End If
    
    Select Case enType
    Case eDataBrowserType.dbtDataItemResponse
        Set oQD = GetQuestionQueryDef(bCountOnly, bCC)
    Case eDataBrowserType.dbtDataItemResponseHistory
        Set oQD = GetAuditQueryDef(bCountOnly, bCC)
'    Case eDataBrowserType.dbteForms
'        Set oQD = GeteFormQueryDef(bCountOnly)
    End Select

    'ic 28/07/2005 added clinical coding
    'use 0 as personid to start with
    AddFilter enType, oQD, sStudySiteSQL, bSingleSubject, _
                    nClinicalTrialId, sSite, _
                    lVisitId, lCRFPageId, _
                    sSubjectLabel, lPersonId, _
                    vStatus, vLockStatus, _
                    bBefore, dblDate, _
                    lDataItemId, sUserCode, _
                     lComments, lDiscStatus, lSDVStatus, lNotes, _
                    lCodingStatus, sDictionaryName, sDictionaryVersion, bCC

    Set oQDs = New QueryDefs
    oQDs.Init sConnect, ""
    
    
    oQDs.Add oQD
    sSerialisedData = oQDs.GetDataString("temp")
    
    Set rs = ADO_DeSerialiseRecordset(sSerialisedData)
'   TA 31/03/2003: If no matches on subjectlabel then subject label is used as a person id
    If rs.RecordCount = 0 Then
    
        If lPersonId = 0 And IsNumeric(sSubjectLabel) Then
            Set oQD = New QueryDef
            'if personid supplied then search on thatperson id if no subject label matches were found
            Select Case enType
            Case eDataBrowserType.dbtDataItemResponse
                Set oQD = GetQuestionQueryDef(bCountOnly, bCC)
            Case eDataBrowserType.dbtDataItemResponseHistory
                Set oQD = GetAuditQueryDef(bCountOnly, bCC)
            End Select
        
            'ic 28/07/2005 added clinical coding
            'same as above with "" for subject label and subjectlabel as person id
            AddFilter enType, oQD, sStudySiteSQL, bSingleSubject, _
                            nClinicalTrialId, sSite, _
                            lVisitId, lCRFPageId, _
                            "", CLng(sSubjectLabel), _
                            vStatus, vLockStatus, _
                            bBefore, dblDate, _
                            lDataItemId, sUserCode, _
                             lComments, lDiscStatus, lSDVStatus, lNotes, _
                            lCodingStatus, sDictionaryName, sDictionaryVersion, bCC

                             
            oQD.Code = "personid"
            oQDs.Add oQD
            sSerialisedData = oQDs.GetDataString("personid")
            Set rs = ADO_DeSerialiseRecordset(sSerialisedData)
        End If
    End If
    
    If bGetSerialisedVersion Then
        'they want the serialised version too
        sSerialisedVersion = sSerialisedData
    End If
    sSerialisedData = ""
    
    If rs.RecordCount = 0 Then
        GetData = Null
    Else
        GetData = rs.GetRows
    End If
    
    Set oQD = Nothing
    Set oQDs = Nothing

    rs.Close
    Set rs = Nothing

End Function

'-----------------------------------------------------
Private Sub AddFilter(enType As eDataBrowserType, oQueryDef As QueryDef, _
                    sStudySiteSQL As String, bSingleSubject As Boolean, _
                    Optional nClinicalTrialId As Integer = -1, Optional sSite As String = "ALL", _
                    Optional lVisitId As Long = 0, Optional lCRFPageId As Long = 0, _
                    Optional sSubjectLabel As String = "", Optional lPersonId As Long = 0, _
                    Optional vStatus As Variant = Null, Optional vLockStatus As Variant = Null, _
                    Optional bBefore As Boolean = False, Optional dblDate As Double = 0, _
                    Optional lDataItemId As Long = 0, Optional sUserCode As String = "", _
                    Optional lComments As Long = -1, _
                    Optional lDiscStatus As Long = -1, Optional lSDVStatus As Long = -1, Optional lNotes As Long = -1, _
                    Optional lCodingStatus As Long = -1, Optional sDictionaryName As String = "", _
                    Optional sDictionaryVersion As String = "", Optional bCC As Boolean = False)
'-----------------------------------------------------
'Add a filter to a DB QueryDef
' vStatus is an array of statuses or null if none
'revisions
'ic 28/07/2005 added clinical coding
'-----------------------------------------------------
Dim sSQL As String
Dim colWheres As Collection
Dim sDateValue As String
Dim i As Long
Dim sTableName As String
Dim sStatusField As String
Dim sTimestampField As String
Dim sVisitField As String
Dim sCRFField As String
Dim sTrialField As String
Dim sSiteField As String
Dim sPersonField As String
Dim vHeaders As Variant


    If (bCC) Then
        vHeaders = Split(mALIASES & mCCALIASES, ",")
    Else
        vHeaders = Split(mALIASES, ",")
    End If
    
    With oQueryDef
        .UserFiltersRemove
        If nClinicalTrialId <> -1 Then
            .UserFilters.Add (vHeaders(dbcStudyId)), "=", (nClinicalTrialId)
        End If

        'filter on site
        If sSite <> "ALL" Then
            .UserFilters.Add (vHeaders(dbcSite)), "=", sSite
        End If

        'filter on visit
        If lVisitId <> 0 Then
            'vHeaders not used as VisitId not in eFroms query
            .UserFilters.Add "VisitInstance.VisitId", "=", (lVisitId)
        End If

        'filter on CRFPage
        If lCRFPageId <> 0 Then
            .UserFilters.Add vHeaders(dbcEFormId), "=", (lCRFPageId)
        End If

    'filter on subject label
        If sSubjectLabel <> "" Then
            If bSingleSubject Then
                'tilde means to case insensitive search (required for oracle)
                .UserFilters.Add (vHeaders(dbcSubjectLabel)), "=~", sSubjectLabel
            Else
                .UserFilters.Add (vHeaders(dbcSubjectLabel)), "like", "%" & sSubjectLabel & "%"
            End If
        End If

        'if personid passed through then use it
        If lPersonId > 0 Then
             .UserFilters.Add (vHeaders(dbcSubjectId)), "=", (lPersonId)
        End If



        'filter on status
        If Not IsNull(vStatus) Then
            If enType = eDataBrowserType.dbteForms Then
                .UserFilters.Add (vHeaders(dbcEFormStatus)), "IN", vStatus
            Else
                .UserFilters.Add (vHeaders(dbcResponseStatus)), "IN", vStatus
            End If
        End If
        
        'ic 28/07/2005 added clinical coding filters
        If (bCC) Then
            If (enType <> eDataBrowserType.dbteForms) Then
                If (lCodingStatus <> -1) Then
                    If (lCodingStatus = eCodingStatus.csNotCoded) Then
                        'if looking for not coded questions, also look for never coded questions
                       .UserFilters.Add (vHeaders(dbcCodingStatus)), "IN", Array(CInt(eCodingStatus.csNotCoded), _
                            CInt(eCodingStatus.csNeverCoded))
                    Else
                        .UserFilters.Add (vHeaders(dbcCodingStatus)), "=", lCodingStatus
                    End If
                End If
                If (sDictionaryName & sDictionaryVersion > "") Then
                    .UserFilters.Add (vHeaders(dbcDictionaryName)), "=", sDictionaryName
                    .UserFilters.Add (vHeaders(dbcDictionaryVersion)), "=", sDictionaryVersion
                End If
            End If
        End If

        'filter on lock status
        If Not IsNull(vLockStatus) Then
            If enType = eDataBrowserType.dbteForms Then
                .UserFilters.Add (vHeaders(dbcEFormLockStatus)), "IN", vLockStatus
            Else
                .UserFilters.Add (vHeaders(dbcDataItemLockStatus)), "IN", vLockStatus
            End If
        End If

        'filter on date
        If dblDate <> 0 Then
            If bBefore Then
                .UserFilters.Add (vHeaders(dbcResponseTimeStamp)), "<", dblDate
            Else
                .UserFilters.Add (vHeaders(dbcResponseTimeStamp)), ">", dblDate + 1
            End If
        End If


        If enType <> eDataBrowserType.dbteForms Then
            'no question, user, comment or new filter on eForms

            'filter on question
            If lDataItemId <> 0 Then
                If enType = dbtDataItemResponse Then
                    .UserFilters.Add "DataItemResponse.DataItemId", "=", (lDataItemId)
                Else
                    .UserFilters.Add "DataItemResponseHistory.DataItemId", "=", (lDataItemId)
                End If
            End If

            'filter on user
            If sUserCode <> "ALL" Then
                If sUserCode <> "" Then
                    .UserFilters.Add (vHeaders(dbcUserName)), "=~", sUserCode
                End If
            End If

            'filter if there are comments
            Select Case lComments
            Case -1 ' no sleection
                'do nothing
            Case -2 'all with comments
                .UserFilters.Add (vHeaders(dbcComments)), "is not null"
            Case 0 ' without notes
                .UserFilters.Add (vHeaders(dbcComments)), "is null"
            End Select


            If enType = dbtDataItemResponse Then
                Select Case lNotes
                Case -1 ' no selection
                    'do nothing
                Case -2 'all
                    .UserFilters.Add "DataItemResponse.NoteStatus", "=", 1 'have notes
                Case 0 'without notes
                    .UserFilters.Add "DataItemResponse.NoteStatus", "=", 0 'no notes
                End Select
                
                Select Case lDiscStatus
                Case -1 'no selection
                        'do nothing
                Case -2  'all with
                    .UserFilters.Add "DataItemResponse.DiscrepancyStatus", "<>", 0  'eDiscrepancyStatus.none
                Case Else ' status passed in
                    .UserFilters.Add "DataItemResponse.DiscrepancyStatus", "=", lDiscStatus  'eDiscrepancyStatus
                End Select
                
                Select Case lSDVStatus
                Case -1 'no selection
                        'do nothing
                Case -2  'all with
                    .UserFilters.Add "DataItemResponse.SDVStatus", "<>", 0  'eSDVStatus.none
                Case Else ' status passed in
                    .UserFilters.Add "DataItemResponse.SDVStatus", "=", lSDVStatus  'eSDVStatus
                End Select
                
            End If

        End If

        If (nClinicalTrialId = -1) Or (sSite = "ALL") Then
              'omly do thsi if a study and site are not selected
            'filter on studies/sites they have a role for
            If sStudySiteSQL <> "" Then
                .UserFilters.Add sStudySiteSQL, "user defined"
            End If
        End If

    End With

End Sub

'-----------------------------------------------------
Private Function GetQuestionQueryDef(bCount As Boolean, bCC As Boolean) As QueryDef
'-----------------------------------------------------
'Create a QueryDef corresponsding to Questions being chosen in the databrowser.
'-----------------------------------------------------
' REVISIONS
' DPH 08/10/2003 - Performance change - retrieve data from DataItemResponse
'           table whereever possible
'-----------------------------------------------------
Dim oSD As QueryDef
Dim vFields As Variant
Dim vAliases As Variant

    Set oSD = New QueryDef
    oSD.Code = "temp"

    oSD.Count = bCount

    ' RS 08/10/2002 Add Timezone column
    ' DPH 08/10/2003 - Changed ClinicalTrial.ClinicalTrialId, TrialSubject.TrialSite, TrialSubject.PersonId
    '           VisitInstance.VisitCycleNumber, CRFPage.CRFPageId, CRFPageInstance.CRFPageCycleNumber
    '           to be retrieved from DataItemResponse table
    'ic 28/07/2005 added clinical coding
    If (bCC) Then
        vFields = Array("DataItemResponse.ClinicalTrialId", "ClinicalTrial.ClinicalTrialName", "TrialSubject.TrialSubjectStatus", "TrialSubject.LockStatus", _
                    "TrialSubject.DiscrepancyStatus", "TrialSubject.SDVStatus", "TrialSubject.NoteStatus", _
                    "DataItemResponse.TrialSite", "DataItemResponse.PersonId", "TrialSubject.LocalIdentifier1", "StudyVisit.VisitName", _
                    "DataItemResponse.VisitCycleNumber", "VisitInstance.VisitStatus", "VisitInstance.LockStatus", _
                    "VisitInstance.DiscrepancyStatus", "VisitInstance.SDVStatus", "VisitInstance.NoteStatus", _
                    "DataItemResponse.CRFPageId", "CRFPage.CRFTitle", _
                    "DataItemResponse.CRFPageCycleNumber", "CRFPageInstance.CRFPageStatus", "CRFPageInstance.LockStatus", _
                    "CRFPageInstance.DiscrepancyStatus", "CRFPageInstance.SDVStatus", "CRFPageInstance.NoteStatus", _
                    "DataItemResponse.ResponseTimestamp", "DataItem.DataItemName", "DataItem.DataType", "DataItemResponse.ResponseValue", _
                    "DataItemResponse.ResponseStatus", "DataItemResponse.UserName", "DataItemResponse.Comments", _
                    "DataItemResponse.LockStatus", _
                    "DataItemResponse.DiscrepancyStatus", "DataItemResponse.SDVStatus", "DataItemResponse.NoteStatus", _
                    "DataItemResponse.ReasonForChange", "DataItemResponse.Changed", _
                    "DataItemResponse.OverruleReason", "DataItemResponse.LabResult", "DataItemResponse.CTCGrade", _
                    "DataItemResponse.VisitId", "DataItemResponse.CRFPageTaskId", "DataItemResponse.ResponseTaskId", _
                    "DataItemResponse.CRFElementId", "substr(DataItemResponse.ValidationMessage,1,255)", "DataItemResponse.RepeatNumber", _
                    "DataItemResponse.ResponseTimestamp_TZ", _
                    "DataItemResponse.DataItemId", "DataItemResponse.ChangeCount", "CRFELEMENT.OWNERQGROUPID", "DATAITEMRESPONSE.USERNAMEFULL", _
                    "DATAITEMRESPONSE.DATABASETIMESTAMP", "DATAITEMRESPONSE.DATABASETIMESTAMP_TZ", _
                    "CRFPAGEINSTANCE.CRFPAGEINSTANCELABEL", "DataItemResponse.DictionaryName", "DataItemResponse.DictionaryVersion", _
                    "DataItemResponse.CodingStatus", "DataItemResponse.CodingDetails")
                    'nb substr must be in lowercase
                    
         vAliases = Split(mALIASES & mCCALIASES, ",")
    Else
        vFields = Array("DataItemResponse.ClinicalTrialId", "ClinicalTrial.ClinicalTrialName", "TrialSubject.TrialSubjectStatus", "TrialSubject.LockStatus", _
                    "TrialSubject.DiscrepancyStatus", "TrialSubject.SDVStatus", "TrialSubject.NoteStatus", _
                    "DataItemResponse.TrialSite", "DataItemResponse.PersonId", "TrialSubject.LocalIdentifier1", "StudyVisit.VisitName", _
                    "DataItemResponse.VisitCycleNumber", "VisitInstance.VisitStatus", "VisitInstance.LockStatus", _
                    "VisitInstance.DiscrepancyStatus", "VisitInstance.SDVStatus", "VisitInstance.NoteStatus", _
                    "DataItemResponse.CRFPageId", "CRFPage.CRFTitle", _
                    "DataItemResponse.CRFPageCycleNumber", "CRFPageInstance.CRFPageStatus", "CRFPageInstance.LockStatus", _
                    "CRFPageInstance.DiscrepancyStatus", "CRFPageInstance.SDVStatus", "CRFPageInstance.NoteStatus", _
                    "DataItemResponse.ResponseTimestamp", "DataItem.DataItemName", "DataItem.DataType", "DataItemResponse.ResponseValue", _
                    "DataItemResponse.ResponseStatus", "DataItemResponse.UserName", "DataItemResponse.Comments", _
                    "DataItemResponse.LockStatus", _
                    "DataItemResponse.DiscrepancyStatus", "DataItemResponse.SDVStatus", "DataItemResponse.NoteStatus", _
                    "DataItemResponse.ReasonForChange", "DataItemResponse.Changed", _
                    "DataItemResponse.OverruleReason", "DataItemResponse.LabResult", "DataItemResponse.CTCGrade", _
                    "DataItemResponse.VisitId", "DataItemResponse.CRFPageTaskId", "DataItemResponse.ResponseTaskId", _
                    "DataItemResponse.CRFElementId", "substr(DataItemResponse.ValidationMessage,1,255)", "DataItemResponse.RepeatNumber", _
                    "DataItemResponse.ResponseTimestamp_TZ", _
                    "DataItemResponse.DataItemId", "DataItemResponse.ChangeCount", "CRFELEMENT.OWNERQGROUPID", "DATAITEMRESPONSE.USERNAMEFULL", _
                    "DATAITEMRESPONSE.DATABASETIMESTAMP", "DATAITEMRESPONSE.DATABASETIMESTAMP_TZ", _
                    "CRFPAGEINSTANCE.CRFPAGEINSTANCELABEL")
                    'nb substr must be in lowercase
                    
        vAliases = Split(mALIASES, ",")
    End If
    

    oSD.QueryFields.Add vFields, vAliases


    oSD.QueryTables.Add "DataItemResponse"
    oSD.QueryTables.Add "ClinicalTrial", , qdjtInner, "ClinicalTrial.ClinicalTrialId", "DataItemResponse.ClinicalTrialId"
    oSD.QueryTables.Add "StudyVisit", , qdjtInner, Array("StudyVisit.ClinicalTrialId", "StudyVisit.VisitId"), _
                            Array("DataItemResponse.ClinicalTrialId", "DataItemResponse.VisitId")
    oSD.QueryTables.Add "CRFPage", , qdjtInner, Array("CRFPage.CRFPageId", "CRFPage.ClinicalTrialId"), _
                            Array("DataItemResponse.CRFPageId", "DataItemResponse.ClinicalTrialId")
    oSD.QueryTables.Add "DataItem", , qdjtInner, Array("DataItem.DataItemId", "DataItem.ClinicalTrialId"), _
                            Array("DataItemResponse.DataItemId", "DataItemResponse.ClinicalTrialId")
    oSD.QueryTables.Add "VisitInstance", , qdjtInner, Array("VisitInstance.ClinicalTrialId", "VisitInstance.TrialSite", "VisitInstance.PersonId", "VisitInstance.VisitId", "VisitInstance.VisitCycleNumber"), _
                            Array("DataItemResponse.ClinicalTrialId", "DataItemResponse.TrialSite", "DataItemResponse.PersonId", "DataItemResponse.VisitId", "DataItemResponse.VisitCycleNumber")
    oSD.QueryTables.Add "CRFPageInstance", , qdjtInner, Array("CRFPageInstance.ClinicalTrialId", "CRFPageInstance.TrialSite", "CRFPageInstance.PersonId", "CRFPageInstance.VisitId", "CRFPageInstance.VisitCycleNumber", "CRFPageInstance.CRFPageTaskId"), _
                            Array("DataItemResponse.ClinicalTrialId", "DataItemResponse.TrialSite", "DataItemResponse.PersonId", "DataItemResponse.VisitId", "DataItemResponse.VisitCycleNumber", "DataItemResponse.CRFPageTaskId")
    oSD.QueryTables.Add "TrialSubject", , qdjtInner, Array("TrialSubject.ClinicalTrialId", "TrialSubject.TrialSite", "TrialSubject.PersonId"), _
                            Array("DataItemResponse.ClinicalTrialId", "DataItemResponse.TrialSite", "DataItemResponse.PersonId")
    oSD.QueryTables.Add "CRFElement", , qdjtInner, Array("CRFElement.ClinicalTrialId", "CRFElement.CRFPageId", "CRFElement.DataItemId"), _
                            Array("DataItemResponse.ClinicalTrialId", "DataItemResponse.CRFPageId", "DataItemResponse.DataItemId")
    If Not bCount Then
        oSD.QueryOrders.Add Array("ClinicalTrial.ClinicalTrialName", "DataItemResponse.TrialSite", "DataItemResponse.PersonId", _
                                "StudyVisit.VisitOrder", "DataItemResponse.VisitCycleNumber,CRFPage.CRFPageOrder", _
                                "DataItemResponse.CRFPageCycleNumber", "CRFElement.FieldOrder", "DataItemResponse.RepeatNumber", "CRFElement.QGroupFieldOrder")
    End If

    Set GetQuestionQueryDef = oSD

    Set oSD = Nothing
    
End Function

'-----------------------------------------------------
Private Function GetAuditQueryDef(bCount As Boolean, bCC As Boolean) As QueryDef
'-----------------------------------------------------
'Create a QueryDef corresponsding to Audit being chosen in the databrowser.
'-----------------------------------------------------
' REVISIONS
' DPH 08/10/2003 - Performance change - retrieve data from DataItemResponseHistory
'           table whereever possible
'-----------------------------------------------------
Dim oSD As QueryDef
Dim vFields As Variant
Dim vAliases As Variant

    Set oSD = New QueryDef
    oSD.Code = "temp"

    oSD.Count = bCount

    ' DPH 08/10/2003 - Changed ClinicalTrial.ClinicalTrialId, TrialSubject.TrialSite, TrialSubject.PersonId
    '           VisitInstance.VisitCycleNumber, CRFPage.CRFPageId, CRFPageInstance.CRFPageCycleNumber
    '           to be retrieved from DataItemResponseHistory table
    'ic 28/07/2005 added clinical coding
    If (bCC) Then
        vFields = Array("DataItemResponseHistory.ClinicalTrialId", "ClinicalTrial.ClinicalTrialName", "TrialSubject.TrialSubjectStatus", "TrialSubject.LockStatus", _
                    "TrialSubject.DiscrepancyStatus", "TrialSubject.SDVStatus", "TrialSubject.NoteStatus", _
                    "DataItemResponseHistory.TrialSite", "DataItemResponseHistory.PersonId", "TrialSubject.LocalIdentifier1", "StudyVisit.VisitName", _
                    "DataItemResponseHistory.VisitCycleNumber", "VisitInstance.VisitStatus", "VisitInstance.LockStatus", _
                    "VisitInstance.DiscrepancyStatus", "VisitInstance.SDVStatus", "VisitInstance.NoteStatus", _
                    "DataItemResponseHistory.CRFPageId", "CRFPage.CRFTitle", _
                    "DataItemResponseHistory.CRFPageCycleNumber", "CRFPageInstance.CRFPageStatus", "CRFPageInstance.LockStatus", _
                    "CRFPageInstance.DiscrepancyStatus", "CRFPageInstance.SDVStatus", "CRFPageInstance.NoteStatus", _
                    "DataItemResponseHistory.ResponseTimestamp", "DataItem.DataItemName", "DataItem.DataType", "DataItemResponseHistory.ResponseValue", _
                    "DataItemResponseHistory.ResponseStatus", "DataItemResponseHistory.UserName", "DataItemResponseHistory.Comments", _
                    "DataItemResponseHistory.LockStatus", _
                    "0", "0", "0", _
                    "DataItemResponseHistory.ReasonForChange", "DataItemResponseHistory.Changed", _
                    "DataItemResponseHistory.OverruleReason", "DataItemResponseHistory.LabResult", "DataItemResponseHistory.CTCGrade", _
                    "DataItemResponseHistory.VisitId", "DataItemResponseHistory.CRFPageTaskId", "DataItemResponseHistory.ResponseTaskId", _
                    "DataItemResponseHistory.CRFElementId", "substr(DataItemResponseHistory.ValidationMessage,1,255)", "DataItemResponseHistory.RepeatNumber", _
                    "DataItemResponseHistory.ResponseTimestamp_TZ", _
                    "DataItemResponseHistory.DataItemId", 0, "CRFELEMENT.OWNERQGROUPID", "DATAITEMRESPONSEHISTORY.USERNAMEFULL", _
                    "DATAITEMRESPONSEHISTORY.DATABASETIMESTAMP", "DATAITEMRESPONSEHISTORY.DATABASETIMESTAMP_TZ", _
                    "CRFPAGEINSTANCE.CRFPAGEINSTANCELABEL", _
                    "0", "0", "0", "0")
                    'ic 28/07/2005 nb the 4 0s are the dictionaryname/dictionaryversion/codingstatus/codingdetails that dont
                    'exist in dataitemresponsehistory
        
        vAliases = Split(mALIASES & mCCALIASES, ",")
    Else
        vFields = Array("DataItemResponseHistory.ClinicalTrialId", "ClinicalTrial.ClinicalTrialName", "TrialSubject.TrialSubjectStatus", "TrialSubject.LockStatus", _
                    "TrialSubject.DiscrepancyStatus", "TrialSubject.SDVStatus", "TrialSubject.NoteStatus", _
                    "DataItemResponseHistory.TrialSite", "DataItemResponseHistory.PersonId", "TrialSubject.LocalIdentifier1", "StudyVisit.VisitName", _
                    "DataItemResponseHistory.VisitCycleNumber", "VisitInstance.VisitStatus", "VisitInstance.LockStatus", _
                    "VisitInstance.DiscrepancyStatus", "VisitInstance.SDVStatus", "VisitInstance.NoteStatus", _
                    "DataItemResponseHistory.CRFPageId", "CRFPage.CRFTitle", _
                    "DataItemResponseHistory.CRFPageCycleNumber", "CRFPageInstance.CRFPageStatus", "CRFPageInstance.LockStatus", _
                    "CRFPageInstance.DiscrepancyStatus", "CRFPageInstance.SDVStatus", "CRFPageInstance.NoteStatus", _
                    "DataItemResponseHistory.ResponseTimestamp", "DataItem.DataItemName", "DataItem.DataType", "DataItemResponseHistory.ResponseValue", _
                    "DataItemResponseHistory.ResponseStatus", "DataItemResponseHistory.UserName", "DataItemResponseHistory.Comments", _
                    "DataItemResponseHistory.LockStatus", _
                    "0", "0", "0", _
                    "DataItemResponseHistory.ReasonForChange", "DataItemResponseHistory.Changed", _
                    "DataItemResponseHistory.OverruleReason", "DataItemResponseHistory.LabResult", "DataItemResponseHistory.CTCGrade", _
                    "DataItemResponseHistory.VisitId", "DataItemResponseHistory.CRFPageTaskId", "DataItemResponseHistory.ResponseTaskId", _
                    "DataItemResponseHistory.CRFElementId", "substr(DataItemResponseHistory.ValidationMessage,1,255)", "DataItemResponseHistory.RepeatNumber", _
                    "DataItemResponseHistory.ResponseTimestamp_TZ", _
                    "DataItemResponseHistory.DataItemId", 0, "CRFELEMENT.OWNERQGROUPID", "DATAITEMRESPONSEHISTORY.USERNAMEFULL", _
                    "DATAITEMRESPONSEHISTORY.DATABASETIMESTAMP", "DATAITEMRESPONSEHISTORY.DATABASETIMESTAMP_TZ", _
                    "CRFPAGEINSTANCE.CRFPAGEINSTANCELABEL")
                    ' nb 0 for change count in DIRH to show no change count
                    'nb substr must be in lowercase
                    'nb the 3 0s are disc/sdv/note status that don't exist in DataItemResponseHistory
                    
        vAliases = Split(mALIASES, ",")
    End If
    

    oSD.QueryFields.Add vFields, vAliases


    oSD.QueryTables.Add "DataItemResponseHistory"
    oSD.QueryTables.Add "ClinicalTrial", , qdjtInner, "ClinicalTrial.ClinicalTrialId", "DataItemResponseHistory.ClinicalTrialId"
    oSD.QueryTables.Add "StudyVisit", , qdjtInner, Array("StudyVisit.ClinicalTrialId", "StudyVisit.VisitId"), _
                            Array("DataItemResponseHistory.ClinicalTrialId", "DataItemResponseHistory.VisitId")
    oSD.QueryTables.Add "CRFPage", , qdjtInner, Array("CRFPage.CRFPageId", "CRFPage.ClinicalTrialId"), _
                            Array("DataItemResponseHistory.CRFPageId", "DataItemResponseHistory.ClinicalTrialId")
    oSD.QueryTables.Add "DataItem", , qdjtInner, Array("DataItem.DataItemId", "DataItem.ClinicalTrialId"), _
                            Array("DataItemResponseHistory.DataItemId", "DataItemResponseHistory.ClinicalTrialId")
    oSD.QueryTables.Add "VisitInstance", , qdjtInner, Array("VisitInstance.ClinicalTrialId", "VisitInstance.TrialSite", "VisitInstance.PersonId", "VisitInstance.VisitId", "VisitInstance.VisitCycleNumber"), _
                            Array("DataItemResponseHistory.ClinicalTrialId", "DataItemResponseHistory.TrialSite", "DataItemResponseHistory.PersonId", "DataItemResponseHistory.VisitId", "DataItemResponseHistory.VisitCycleNumber")
    oSD.QueryTables.Add "CRFPageInstance", , qdjtInner, Array("CRFPageInstance.ClinicalTrialId", "CRFPageInstance.TrialSite", "CRFPageInstance.PersonId", "CRFPageInstance.VisitId", "CRFPageInstance.VisitCycleNumber", "CRFPageInstance.CRFPageTaskId"), _
                            Array("DataItemResponseHistory.ClinicalTrialId", "DataItemResponseHistory.TrialSite", "DataItemResponseHistory.PersonId", "DataItemResponseHistory.VisitId", "DataItemResponseHistory.VisitCycleNumber", "DataItemResponseHistory.CRFPageTaskId")
    oSD.QueryTables.Add "TrialSubject", , qdjtInner, Array("TrialSubject.ClinicalTrialId", "TrialSubject.TrialSite", "TrialSubject.PersonId"), _
                            Array("DataItemResponseHistory.ClinicalTrialId", "DataItemResponseHistory.TrialSite", "DataItemResponseHistory.PersonId")
    oSD.QueryTables.Add "CRFElement", , qdjtInner, Array("CRFElement.ClinicalTrialId", "CRFElement.CRFPageId", "CRFElement.DataItemId"), _
                            Array("DataItemResponseHistory.ClinicalTrialId", "DataItemResponseHistory.CRFPageId", "DataItemResponseHistory.DataItemId")

    If Not bCount Then
        oSD.QueryOrders.Add Array("ClinicalTrial.ClinicalTrialName", "DataItemResponseHistory.TrialSite", "DataItemResponseHistory.PersonId", _
                                "StudyVisit.VisitOrder", "DataItemResponseHistory.VisitCycleNumber,CRFPage.CRFPageOrder", _
                                "DataItemResponseHistory.CRFPageCycleNumber", "CRFElement.FieldOrder", "DataItemResponseHistory.RepeatNumber", "CRFElement.QGroupFieldOrder")

        'TA 24/06/2002 SR 4662: Order audit data by ResponseTimeStamp descending
        oSD.QueryOrders.Add "DataItemResponseHistory.ResponseTimeStamp", True
    End If

    Set GetAuditQueryDef = oSD

    Set oSD = Nothing

End Function

'TA 20/01/03: OCMMENTED OUT AS NO LONGER KEPT UPTODATE WITH CHANGES - IE SDVSTATUS ETC
''-----------------------------------------------------
'Private Function GeteFormQueryDef(bCount As Boolean) As QueryDef
''-----------------------------------------------------
''Create a QueryDef corresponsding to Audit being chosen in the databrowser.
''-----------------------------------------------------
'Dim oSD As QueryDef
'Dim vFields As Variant
'Dim vAliases As Variant
'
'    Set oSD = New QueryDef
'    oSD.Code = "temp"
'
'    oSD.Count = bCount
'
'    vFields = Array("ClinicalTrial.ClinicalTrialId", "ClinicalTrial.ClinicalTrialName", "TrialSubject.TrialSubjectStatus", "TrialSubject.LockStatus", _
'                    "TrialSubject.TrialSite", "TrialSubject.PersonId", "TrialSubject.LocalIdentifier1", "StudyVisit.VisitName", _
'                    "VisitInstance.VisitCycleNumber", "VisitInstance.VisitStatus", "VisitInstance.LockStatus", "CRFPage.CRFPageId", "CRFPage.CRFTitle", _
'                    "CRFPageInstance.CRFPageCycleNumber", "CRFPageInstance.CRFPageStatus", "CRFPageInstance.LockStatus", "CRFPageInstance.CRFPageDate")
'
'                    'nb substr must be in lowercase
'    vAliases = Split(mALIASES, ",")
'
'    oSD.QueryFields.Add vFields, vAliases
'
'
'    oSD.QueryTables.Add "CRFPageInstance"
'    oSD.QueryTables.Add "ClinicalTrial", , qdjtInner, "ClinicalTrial.ClinicalTrialId", "CRFPageInstance.ClinicalTrialId"
'    oSD.QueryTables.Add "StudyVisit", , qdjtInner, Array("StudyVisit.ClinicalTrialId", "StudyVisit.VisitId"), _
'                            Array("CRFPageInstance.ClinicalTrialId", "CRFPageInstance.VisitId")
'    oSD.QueryTables.Add "CRFPage", , qdjtInner, Array("CRFPage.CRFPageId", "CRFPage.ClinicalTrialId"), _
'                            Array("CRFPageInstance.CRFPageId", "CRFPageInstance.ClinicalTrialId")
'    oSD.QueryTables.Add "VisitInstance", , qdjtInner, Array("VisitInstance.ClinicalTrialId", "VisitInstance.TrialSite", "VisitInstance.PersonId", "VisitInstance.VisitId", "VisitInstance.VisitCycleNumber"), _
'                            Array("CRFPageInstance.ClinicalTrialId", "CRFPageInstance.TrialSite", "CRFPageInstance.PersonId", "CRFPageInstance.VisitId", "CRFPageInstance.VisitCycleNumber")
'    oSD.QueryTables.Add "TrialSubject", , qdjtInner, Array("TrialSubject.ClinicalTrialId", "TrialSubject.TrialSite", "TrialSubject.PersonId"), _
'                            Array("CRFPageInstance.ClinicalTrialId", "CRFPageInstance.TrialSite", "CRFPageInstance.PersonId")
'
'
'    If Not bCount Then
'        oSD.QueryOrders.Add Array("ClinicalTrial.ClinicalTrialName", "CRFPageInstance.TrialSite", "CRFPageInstance.PersonId", _
'                                    "StudyVisit.VisitOrder", "CRFPageInstance.VisitCycleNumber", "CRFPage.CRFPageOrder", _
'                                    "CRFPageInstance.CRFPageCycleNumber")
'    End If
'
'
'    Set GeteFormQueryDef = oSD
'
'End Function

