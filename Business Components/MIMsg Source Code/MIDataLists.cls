VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIDataLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "MIMessages Data Lists"
'----------------------------------------------------------------------------------------'
'   File:       MIMessageDataList.cls
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Toby Aldridge, Nov 2001
'   Purpose:    MIMessages Data Lists
'----------------------------------------------------------------------------------------'
' Revisions:
'   TA 02/10/2001: Filer on SiteUser (IT WOULD BE BETTER IF A LIST OF ALLOWED SITES WERE
'                    PASSED IN SO I WOULDN'T NEED TO DO THE EXTRA JOIN ON THE SITE USER TABLE.
'                    THIS SHOULD BE DONE WHEN WE HAVE DECIDED WHERE PERMSISSIONS AND PERMITTED
'                    SITES WILL BE STORED (MACRO OR SECURITY))
'   ic 02/04/2002  added CRFPageInstance.CRFPageId to fields returned in data list
' NCJ 15 Oct 02 - Removed apparently unused Headers property, mALIASES and eDropDownCol
' NCJ 18 Oct 02 - Added ResponseValue to MIMessageList columns
' NCJ 5 Nov 02 - New MIMessageExists function
' NCJ 6 Nov 02 - Added Scope filters argument to GetMIMessageList
' TA 20/11/2002: Get ResponseValue and Username from DataItemResponseHistory
' TA 13/03/2003: Added GetDiscrepancyStatusCount top count no of disc at any time
' DPH 13/10/2003 - performance enhancements to SQL in GetMIMsgStatusCount
' ic 09/11/2004 bug 2400, add responsecycle argument to MIMessageExists()
' TA 18/11/2004: flag is passed in to GetMIMessageList whether show OC Ids - issue 2448
' ic 27/02/2007 issue 2114, added GMT to timestamps
'----------------------------------------------------------------------------------------'

Option Explicit
                                    
' column number enumeration
' NB This MUST be in sync. with the SQL in GetMIMessageList
' NCJ 15 Oct 02 - Added MIMessageScope
Public Enum MIMsgCol
    mmcId = 0
    mmcSource = 1
    mmcObjectId = 2
    mmcObjectSource = 3
    mmcScope = 4        ' NCJ 15 Oct 02 (and all following values incremented)
    mmcSite = 5
    mmcPrioirty = 6
    mmcCreated = 7
    mmcStatus = 8
    mmcText = 9
    mmcUserNameFull = 10
    mmcUserName = 11
    mmcExternalId = 12
    mmcStudyName = 13
    mmcSubjectId = 14
    mmcVisitId = 15
    mmcVisitCycle = 16
    mmcEFormTaskId = 17
    mmcResponseTaskId = 18
    mmcResponseCycle = 19
    mmcResponseValue = 20   ' Added NCJ 18 Oct 02
    'ta 18/04/2002: added timesent for web
    mmcTimeSent = 21
    mmcStudyid = 22
    mmcSubjectLabel = 23
    
    'TA changed visitname, questions name , eform title to new fields
    mmcEFormCycle = 24
    mmcQuestionId = 25
    mmcDataUsername = 26
    'ic 02/04/2002
    'added mmcEFormId to column number enumeration for web I/O module
    mmcEFormId = 27
    'ic 27/02/2007 issue 2114, added GMT to timestamps
    mmcCreated_TZ = 28

End Enum

'----------------------------------------------------------------------------------------'
Public Enum eResponseDetails
'----------------------------------------------------------------------------------------'

    rdResponseValue = 0
    rdResponseTimeStamp = 1
    rdQuestionId = 2
    rdUserName = 3

End Enum

'----------------------------------------------------------------------------------------'
Public Function GetMIMessageList(sCon As String, bUseOCID As Boolean, sCurrentUserName As String, sStudySiteSQL As String, nType As MIMsgType, _
                                    vScope As Variant, _
                                    Optional sStudyName As String = "", _
                                    Optional sSite As String = "", _
                                    Optional sSubjectlabel As String = "", _
                                    Optional lSubjectId As Long = -1, _
                                    Optional lVisitId As Long = -1, _
                                    Optional lEFormId As Long = -1, _
                                    Optional lQuId As Long = -1, _
                                    Optional sUserName As String = "", _
                                    Optional vStatus As Variant = "", _
                                    Optional bBefore As Boolean = True, Optional dblDate As Double = 0, _
                                    Optional sDataUsername As String = "", _
                                    Optional bGetSerialisedVersion As Boolean = False, _
                                    Optional ByRef sSerialisedVersion As String = "", _
                                    Optional lVisitCycle As Long = -1, _
                                    Optional lEFormCycle As Long = -1, _
                                    Optional lQuestionCycle As Long = -1) As Variant
'----------------------------------------------------------------------------------------'
' Get the list of all MIMessages, depending on optional filters
' NCJ/TA 15 Oct 02 - Changed joins from Inner to Left (because for SDVs there may not be a Question, eForm or Visit)
' NCJ 6 Nov 02 - Added vScope argument (array of scopes to filter on, or empty string for no scope filters)
' TA 20/11/2002: Get ResponseValue and Username from DataItemResponseHistory
' DPH 26/03/2004 Close recordsets & terminate objects
' MLM 30/06/05: Added the ability to search on visit, eForm and question cycle number.
' ic 27/02/2007 issue 2114, added GMT to timestamps
'----------------------------------------------------------------------------------------'
Dim oQDs As QueryDefs
Dim oQD As QueryDef
Dim sFields As String
Dim sTables As String
Dim rs As Recordset
Dim sSerialisedData As String
Dim sOCIDField As String

    On Error GoTo ErrLabel
    
    
' TA 18/11/2004: flag is passed in to GetMIMessageList whether show OC Ids - issue 2448
    If bUseOCID Then
        sOCIDField = "MIMessageOCDiscrepancyID"
    Else
        'use concatenation to make unique id
        If Connection_AsType(sCon).Provider = CONNECTION_MSDAORA Then
            sOCIDField = "MIMESSAGESITE||'-'||MIMESSAGEOBJECTSOURCE||'-'||MIMESSAGEOBJECTID"
        Else
            sOCIDField = "MIMESSAGESITE + '-' + cast(MIMESSAGEOBJECTSOURCE as varchar(10)) + '-' + cast(MIMESSAGEOBJECTID as varchar(10))"
        End If
    End If
        
    
    Set oQD = New QueryDef

    

    ' !!!! NB These cols. MUST be in sync. with the MIMsgCol enumeration above
    ' NCJ 15 Oct 02 - Added MIMessageScope
    sFields = "MIMessageId,MIMessageSource,MIMessageObjectId,MIMessageObjectSource,MIMessageScope, " _
                & "MIMessageSite,MIMessagePriority,MIMessageCreated,MIMessageStatus, " _
                & "MIMessageText,MIMessageUserNameFull,MIMessageUserName," & sOCIDField
    'following fields for the scope
    'ic 02/04/2002 - added CRFPageInstance.CRFPageId to sFields for web I/O module
    'TA 2/4/2002: response cycle changed to person id as not in 2.2
    'NCJ 18 Oct 02 - ResponseValue added
    
'    sFields = sFields & ",MIMessageTrialName, MIMessagePersonId, " _
'                & "MIMessageVisitId,MIMessageVisitCycle,MIMessageCRFPageTaskId, " _
'                & "MIMessageResponseTaskId, MIMessageResponseCycle, DATAITEMRESPONSEHISTORY.RESPONSEVALUE, " _
'                & "MIMessageSent"
'    sFields = sFields & ",ClinicalTrial.ClinicalTrialID,LocalIdentifier1, " _
'                & "MIMESSAGECRFPAGECYCLE, MIMESSAGEDATAITEMID,DATAITEMRESPONSEHISTORY.USERNAME,MIMESSAGECRFPAGEID"
    
    'ic 27/02/2007 issue 2114, added GMT to timestamps
    sFields = sFields & ",MIMessageTrialName, MIMessagePersonId, " _
                & "MIMessageVisitId,MIMessageVisitCycle,MIMessageCRFPageTaskId, " _
                & "MIMessageResponseTaskId, MIMessageResponseCycle, MIMESSAGERESPONSEVALUE, " _
                & "MIMessageSent"
    sFields = sFields & ",ClinicalTrial.ClinicalTrialID,LocalIdentifier1, " _
                & "MIMESSAGECRFPAGECYCLE, MIMESSAGEDATAITEMID,'',MIMESSAGECRFPAGEID, MIMESSAGECREATED_TZ"
    
    
    oQD.InitSelect "MIMessage", Split(sFields, ",")
    
    oQD.QueryTables.Add "ClinicalTrial", , qdjtInner, "ClinicalTrial.ClinicalTrialName", "MIMessage.MIMessageTrialName"

    'TA 07/08/2000 SR3763: additional WHERE clause to get subject label from trial subject
    oQD.QueryTables.Add "TrialSubject", , qdjtInner, _
                        Array("TrialSubject.ClinicalTrialId", "TrialSubject.TrialSite", "TrialSubject.PersonId"), _
                        Array("ClinicalTrial.ClinicalTrialId", "MIMessage.MIMessageSite", "MIMessage.MIMessagePersonId")
                           
                           
 '   oQD.QueryTables.Add "DATAITEMRESPONSEHISTORY", , qdjtLeft, Array( _
        "CLINICALTRIAL.CLINICALTRIALID", "MIMESSAGE.MIMESSAGESITE", "MIMESSAGE.MIMESSAGEPERSONID", "MIMESSAGE.MIMESSAGERESPONSETASKID", _
                                        "MIMESSAGE.MIMESSAGERESPONSECYCLE", "MIMESSAGE.MIMESSAGERESPONSETIMESTAMP"), _
                                        Array( _
        "DATAITEMRESPONSEHISTORY.CLINICALTRIALID", "DATAITEMRESPONSEHISTORY.TRIALSITE", "DATAITEMRESPONSEHISTORY.PERSONID", "DATAITEMRESPONSEHISTORY.RESPONSETASKID", _
                                        "DATAITEMRESPONSEHISTORY.REPEATNUMBER", "DATAITEMRESPONSEHISTORY.RESPONSETIMESTAMP")
    
                           
    oQD.QueryOrders.Add "MIMessagePriority"
    oQD.QueryOrders.Add "MIMessageCreated", True
    
    'filter section
    
   
    If (sStudyName = "") Or (sSite = "") Then
        'only do thsi if a study and site are not selected
        'filter on studies/sites they have a role for
        If sStudySiteSQL <> "" Then
            oQD.QueryFilters.Add sStudySiteSQL, "user defined"
        End If
    End If
    
    oQD.QueryFilters.Add "MIMessageHistory", "=", MIMsgHistory.mimhCurrent
    
    oQD.QueryFilters.Add "MIMessageType", "=", nType
    

    If sStudyName <> "" Then
        oQD.QueryFilters.Add "MIMessageTrialName", "=", sStudyName
    End If
    
    If sSite <> "" Then
        oQD.QueryFilters.Add "MIMessageSite", "=", sSite
    End If
    
    If lVisitId <> -1 Then
        oQD.QueryFilters.Add "MIMessageVisitId", "=", lVisitId
    End If
    
    ' MLM 30/06/05:
    If lVisitCycle <> -1 Then
        oQD.QueryFilters.Add "MIMessageVisitCycle", "=", lVisitCycle
    End If
    
    'filter on eform and dataitem
    If lEFormId <> -1 Then
        oQD.QueryFilters.Add "MIMESSAGECRFPAGEID", "=", lEFormId
    End If
    
    ' MLM 30/06/05:
    If lEFormCycle <> -1 Then
        oQD.QueryFilters.Add "MIMessageCRFPageCycle", "=", lEFormCycle
    End If
    
    If lQuId <> -1 Then
        oQD.QueryFilters.Add "MIMESSAGEDATAITEMID", "=", lQuId
    End If
        
    ' MLM 30/06/05:
    If lQuestionCycle <> -1 Then
        oQD.QueryFilters.Add "MIMessageResponseCycle", "=", lQuestionCycle
    End If
    
    'use user filter so that we can remove it later
    'Get subject label search string from textbox
    If sSubjectlabel <> "" Then
        oQD.UserFilters.Add "LocalIdentifier1", SQL_LIKE, "%" & sSubjectlabel & "%"
    End If
    
    'TA 30/10/2000: if we come though the view menu we need to restrict on user id
    If lSubjectId <> -1 Then
         oQD.QueryFilters.Add "MIMessagePersonId", "=", lSubjectId
    End If
    
    If sUserName <> "" Then
        'use tilde to make caase insensitive search
        oQD.QueryFilters.Add "MIMessageUserName", "=~", sUserName
    End If
    

    If VarType(vStatus) <> vbString Then
        oQD.QueryFilters.Add "MIMessageStatus", SQL_IN, vStatus
    End If
    
    ' NCJ 6 Nov 02 - Filter on Scope too
    If VarType(vScope) <> vbString Then
        oQD.QueryFilters.Add "MIMessageScope", SQL_IN, vScope
    End If
    
    'filter on date
    If dblDate <> 0 Then
        If bBefore Then
            oQD.UserFilters.Add "MIMessageCreated", "<", dblDate
        Else
            oQD.UserFilters.Add "MIMessageCreated", ">", dblDate + 1
        End If
    End If

    oQD.Code = "MIMessage"
    
    Set oQDs = New QueryDefs
    oQDs.Init sCon, ""
    oQDs.Add oQD

    sSerialisedData = oQDs.GetDataString("MIMessage")
    Set rs = ADO_DeSerialiseRecordset(sSerialisedData)
    
    If nType = mimtNote Then
        'if it is a note wew need to filter out private notes that are not for this user
        rs.Filter = "MIMessageStatus = " & eNoteMIMStatus.nsPublic & " or MIMessageUserName = '" & sCurrentUserName & "'"
        If bGetSerialisedVersion Then
            sSerialisedData = ADO_SerialiseRecordset(rs)
        End If
    End If
    
    If rs.RecordCount = 0 Then
        'remove subjecft label filter
        oQDs.Item("MIMessage").UserFiltersRemove
        If IsNumeric(sSubjectlabel) Then
            'if a long use subjectlabel as personid filter
            On Error Resume Next
             oQD.QueryFilters.Add "MIMessagePersonId", "=", CLng(sSubjectlabel)
             On Error GoTo ErrLabel
             If Err.Number = 0 Then
                sSerialisedData = oQDs.GetDataString("MIMessage")
                Set rs = ADO_DeSerialiseRecordset(sSerialisedData)
                If nType = mimtNote Then
                    'if it is a note wew need to filter out private notes that are not for this user
                    rs.Filter = "MIMessageStatus = " & eNoteMIMStatus.nsPublic & " or MIMessageUserName = '" & sCurrentUserName & "'"
                    If bGetSerialisedVersion Then
                        sSerialisedData = ADO_SerialiseRecordset(rs)
                    End If
                End If
             End If
        End If
    End If
    
    If bGetSerialisedVersion Then
        'they want the serialised version too
        sSerialisedVersion = sSerialisedData
    End If
    sSerialisedData = ""
    
    If rs.RecordCount = 0 Then
        GetMIMessageList = Null
    Else
        GetMIMessageList = rs.GetRows
    End If
    
    rs.Close
    Set rs = Nothing
    Set oQDs = Nothing
    Set oQD = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIDataLists.GetMIMessageList"

End Function
    

'---------------------------------------------------------------------
Public Function GetResponseDetails(sCon As String, _
                                      sStudyName As String, sSite As String, lSubjectId As Long, _
                                      lResponseTaskId As Long, nResponseCycle As Integer) As Variant
'---------------------------------------------------------------------
' Retrieves dataitemresponse details for a given responsetaskid.
' cols as in enum eResponseDetials
' This is used to fill in the new values for each new mimessage row.
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs

    Set oQD = New QueryDef
    
    oQD.InitSelect "DataItemResponse", Array("RESPONSEVALUE", "RESPONSETIMESTAMP", "DATAITEMID", "USERNAME"), _
                        Array("TrialSite", "PersonId", "ResponseTaskId", "RepeatNumber"), _
                        Array(sSite, lSubjectId, lResponseTaskId, nResponseCycle)
    oQD.QueryTables.Add "ClinicalTrial", , qdjtInner, _
                            "ClinicalTrial.ClinicalTrialId", "DataItemResponse.ClinicalTrialId"
    oQD.QueryFilters.Add "ClinicalTrial.ClinicalTrialName", "=", sStudyName
         
    oQD.Code = "temp"
    Set oQDs = New QueryDefs
    
    oQDs.Init sCon, ""
    oQDs.Add oQD
    
    GetResponseDetails = oQDs.GetDataArray("temp")
    
    Set oQD = Nothing
    Set oQDs = Nothing
    
End Function


'functions to return 2 column arrays to populate drop down lists
    

'---------------------------------------------------------------------
Public Function GetStudyList(sCon As String, nMsgType As Integer) As Variant
Attribute GetStudyList.VB_Description = "retrieves list of studies as a 2d array for a combo. First col is the text(name) the second is the id."
'---------------------------------------------------------------------
' retrieves list of studies as a 2d array for a combo.
' First col is the text(name) the second is the id.
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs

    Set oQD = New QueryDef
    
    'data browser way
'    With oQD
'        .InitSelect "ClinicalTrial", Array("ClinicalTrialName", "ClinicalTrialId")
'        .QueryFilters.Add "ClinicalTrialId", "<>", 0
'    End With

    oQD.InitSelect "MIMessage", "MIMessageTrialName", "MIMessageType", nMsgType, "MIMessageTrialName"
    oQD.Distinct = True
    
    oQD.Code = "temp"
    Set oQDs = New QueryDefs
    
    oQDs.Init sCon, ""
    oQDs.Add oQD
    
    GetStudyList = oQDs.GetDataArray("temp")
    
    Set oQD = Nothing
    Set oQDs = Nothing
    
End Function


'---------------------------------------------------------------------
Public Function GetSiteList(sCon As String, nMsgType As Integer, sCurrentUserName As String, _
                            Optional sStudyName As String = "") As Variant
Attribute GetSiteList.VB_Description = "retrieves list of sites filtered by studyid and current username as a 2d array for a combo. First col is the text(code) the second is the description. lStudyId is not currently used."
'---------------------------------------------------------------------
' retrieves list of sites filtered by studyid and current username as a 2d array for a combo.
' First col is the text(code) the second is the description.
' lStudyId is not currently used.
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs

    Set oQD = New QueryDef
    
    'data browser way - uses a study id
'    With oQD
'        .InitSelect "Site", Array("Site.SiteDescription", "Site.Site")
'        .QueryTables.Add "TrialSite", , qdjtInner, "TrialSite.TrialSite", "Site.Site"
'        'add site user table
'        .QueryTables.Add "SiteUser", , qdjtInner, "TrialSite.TrialSite", "SiteUser.Site"
'        If lStudyId <> -1 Then
'        .QueryFilters.Add "ClinicalTrialId", "=", lStudyId
'        end if
'        'restrict by user
'        .QueryFilters.Add "UserName", "=", sCurrentUserName
'    End With
    
    With oQD
        .InitSelect "SiteUser", "SiteUser.Site", "SiteUSer.UserName", sCurrentUserName, "SiteUser.Site"
        .QueryTables.Add "MIMessage", , qdjtInner, "MIMEssage.MIMessageSite", "SiteUser.Site"
        .QueryFilters.Add "MIMessageType", "=", nMsgType
        If sStudyName <> "" Then
            .QueryFilters.Add "MIMessageTrialName", "=", sStudyName
        End If
        .Distinct = True
    End With
    
    oQD.Code = "temp"
    Set oQDs = New QueryDefs
    
    oQDs.Init sCon, ""
    oQDs.Add oQD
    
    GetSiteList = oQDs.GetDataArray("temp")
    
    Set oQD = Nothing
    Set oQDs = Nothing
    
    
End Function

'---------------------------------------------------------------------
Public Function GetVisitList(sCon As String, nMsgType As Integer, _
                                Optional sStudyName As String = "", Optional sSite As String = "") As Variant
Attribute GetVisitList.VB_Description = "retrieves list of Visits as a 2d array for a combo. First col is the text(name) the second is the id."
'---------------------------------------------------------------------
' retrieves list of Visits as a 2d array for a combo.
' First col is the text(name) the second is the id.
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs

    Set oQD = New QueryDef
    
'    'data browser way - uses a study id
'    With oQD
'        .InitSelect "StudyVisit", Array("StudyVisit.VisitName", "StudyVisit.VisitId"), "StudyVisit.ClinicalTrialId", lStudyId
'    End With
    
    With oQD
        .InitSelect "StudyVisit", Array("VisitName", "VisitId"), , , "StudyVisit.VisitName"
        .QueryTables.Add "ClinicalTrial", , qdjtInner, "StudyVisit.ClinicalTrialId", "ClinicalTrial.ClinicalTrialId"
        .QueryTables.Add "MIMessage", , qdjtInner, "ClinicalTrial.ClinicalTrialName", "MIMessage.MIMessageTrialName"
        .QueryFilters.Add "MIMessageType", "=", nMsgType
        If sStudyName <> "" Then
            .QueryFilters.Add "MIMessageTrialName", "=", sStudyName
        End If
        If sSite <> "" Then
            .QueryFilters.Add "MIMessageSite", "=", sSite
        End If
        
        .Distinct = True
    End With
        
    oQD.Code = "temp"
    Set oQDs = New QueryDefs
    
    oQDs.Init sCon, ""
    oQDs.Add oQD
    
    GetVisitList = oQDs.GetDataArray("temp")
    
    Set oQD = Nothing
    Set oQDs = Nothing
    
    
End Function

'---------------------------------------------------------------------
Public Function GetEFormList(sCon As String, nMsgType As Integer, _
                                Optional sStudyName As String = "", Optional sSite As String = "") As Variant
Attribute GetEFormList.VB_Description = "retrieves list of eForms as a 2d array for a combo. First col is the text(CRFTitle) the second is the id."
'---------------------------------------------------------------------
' retrieves list of eForms as a 2d array for a combo.
' First col is the text(CRFTitle) the second is the id.
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs

    Set oQD = New QueryDef
    
'    '    'data browser way - uses a study id and visit id
'    With oQD
'        .InitSelect "CRFPage", Array("CRFPage.CRFTitle", "CRFPage.CRFPageId"), "CRFPage.ClinicalTrialId", lStudyId
'        If lVisitId <> 0 Then
'            .QueryTables.Add "StudyVisitCRFPage", , qdjtInner, _
'            Array("StudyVisitCRFPage.ClinicalTrialId", "StudyVisitCRFPage.CRFPageId"), _
'            Array("CRFPage.ClinicalTrialId", "CRFPage.CRFPageId")
'            .QueryFilters.Add "StudyVisitCRFPage.VisitId", "=", lVisitId
'        End If
'    End With
    
    
    With oQD
        .InitSelect "CRFPage", "CRFTitle", , , "CRFPage.CRFTitle"
        .QueryTables.Add "ClinicalTrial", , qdjtInner, "CRFPage.ClinicalTrialId", "ClinicalTrial.ClinicalTrialId"
        .QueryTables.Add "MIMessage", , qdjtInner, "ClinicalTrial.ClinicalTrialName", "MIMessage.MIMessageTrialName"
        .QueryFilters.Add "MIMessageType", "=", nMsgType
        If sStudyName <> "" Then
            .QueryFilters.Add "MIMessageTrialName", "=", sStudyName
        End If
        If sSite <> "" Then
            .QueryFilters.Add "MIMessageSite", "=", sSite
        End If
        
        .Distinct = True
    End With
    
    oQD.Code = "temp"
    Set oQDs = New QueryDefs
    
    oQDs.Init sCon, ""
    oQDs.Add oQD
    
    GetEFormList = oQDs.GetDataArray("temp")
    
    Set oQD = Nothing
    Set oQDs = Nothing
    
    
End Function

'---------------------------------------------------------------------
Public Function GetQuestionList(sCon As String, nMsgType As Integer, _
                                Optional sStudyName As String = "", Optional sSite As String = "") As Variant
Attribute GetQuestionList.VB_Description = "retrieves list of questions as a 2d array for a combo. First col is the text (question name) the second is the id."
'---------------------------------------------------------------------
' retrieves list of questions as a 2d array for a combo.
' First col is the text (question name) the second is the id.
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs

    Set oQD = New QueryDef
    
'    '    'data browser way - uses a study id and visit id
'    With oQD
'        .InitSelect "DataItem", Array("DataItem.DataItemName", "DataItem.DataITemID"), "DataItem.ClinicalTrialId", lStudyId
'        If leFormId <> 0 Then
'            .QueryTables.Add "CRFElement", , qdjtInner, _
'            Array("DataItem.ClinicalTrialId", "DataITem.DataItemId"), _
'            Array("CRFElement.ClinicalTrialId", "CRFElement.DataItemId")
'            .QueryFilters.Add "CRFElement.CRFPageId", "=", leFormId
'        End If
'    End With
    
    
    With oQD
        .InitSelect "DataItem", "DataItemName", , , "DataItem.DataItemName"
        .QueryTables.Add "ClinicalTrial", , qdjtInner, "DataItem.ClinicalTrialId", "ClinicalTrial.ClinicalTrialId"
        .QueryTables.Add "MIMessage", , qdjtInner, "ClinicalTrial.ClinicalTrialName", "MIMessage.MIMessageTrialName"
        .QueryFilters.Add "MIMessageType", "=", nMsgType
        If sStudyName <> "" Then
            .QueryFilters.Add "MIMessageTrialName", "=", sStudyName
        End If
        If sSite <> "" Then
            .QueryFilters.Add "MIMessageSite", "=", sSite
        End If
        
        .Distinct = True
    End With
        
    oQD.Code = "temp"
    Set oQDs = New QueryDefs
    
    oQDs.Init sCon, ""
    oQDs.Add oQD
    
    GetQuestionList = oQDs.GetDataArray("temp")
    
    Set oQD = Nothing
    Set oQDs = Nothing
    
    
End Function

'---------------------------------------------------------------------
Public Function GetUserList(sCon As String, nMsgType As Integer, _
                                Optional sStudyName As String = "", Optional sSite As String = "") As Variant
Attribute GetUserList.VB_Description = "retrieves list of users for a dropdown as a 2d array for a combo. First col is the text (UserName) the second is also the username."
'---------------------------------------------------------------------
' retrieves list of users for a dropdown as a 2d array for a combo.
' First col is the text (UserName) the second is also the username.
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs

    Set oQD = New QueryDef
    
'    '    'data browser way - uses a study id
'    With oQD
'        .InitSelect "DataItemResponse", Array("UserName", "UserName"), "ClinicalTrialId", lStudyId
'        .Distinct = True
'    End With
    
    With oQD
        .InitSelect "MIMessage", "MIMessageUserName", "MIMessageType", nMsgType, "MIMessageUserName"
        If sStudyName <> "" Then
            .QueryFilters.Add "MIMessageTrialName", "=", sStudyName
        End If
        If sSite <> "" Then
            .QueryFilters.Add "MIMessageSite", "=", sSite
        End If
        
        .Distinct = True
    End With
        
    oQD.Code = "temp"
    Set oQDs = New QueryDefs
    
    oQDs.Init sCon, ""
    oQDs.Add oQD
    
    GetUserList = oQDs.GetDataArray("temp")
    
    Set oQD = Nothing
    Set oQDs = Nothing
    
    
End Function

'---------------------------------------------------------------------
Public Function MIMessageExists(sCon As String, _
                                nType As MIMsgType, nScope As MIMsgScope, _
                                sStudyName As String, sSite As String, lSubjectId As Long, _
                                ByRef lObjectId As Long, ByRef nObjectSource As Integer, _
                                Optional lVisitId As Long = 0, Optional nVisitCycle As Integer = 0, _
                                Optional lEFormTaskId As Long = 0, _
                                Optional lResponseId As Long = 0, Optional nResponseCycle As Integer = 0) As Boolean
'---------------------------------------------------------------------
' NCJ 5 Nov 02 - Does an MIMessage already exist for this object?
' If so, its ObjectId and ObjectSource will be returned in lObjectId and nObjectSource
' Note that the optional parameters should be supplied according to the Scope
' Revisions
' ic 09/11/2004 bug 2400, add responsecycle argument
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs
Dim vData As Variant
Const sCODE = "mimsg"   ' Arbitrary!

    On Error GoTo ErrLabel
    
    MIMessageExists = False
    lObjectId = 0
    
    Set oQD = New QueryDef
    oQD.Code = sCODE
    
    ' Select just the ObjectId and the ObjectSource, for this MessageType and Scope
    oQD.InitSelect "MIMessage", _
                    Array("MIMessageObjectId", "MIMessageObjectSource"), _
                    Array("MIMessageType", "MIMessageScope", "MIMessageHistory"), _
                    Array(nType, nScope, MIMsgHistory.mimhCurrent)
    ' Filter on Study, Site, Subject
    oQD.QueryFilters.Add "MIMessageTrialName", "=", sStudyName
    oQD.QueryFilters.Add "MIMessageSite", "=", sSite
    oQD.QueryFilters.Add "MIMessagePersonId", "=", lSubjectId
    ' Assume correct filtering parameters are given for the scope
    If lVisitId > 0 Then
        oQD.QueryFilters.Add "MIMessageVisitId", "=", lVisitId
        oQD.QueryFilters.Add "MIMessageVisitCycle", "=", nVisitCycle
    End If
    If lEFormTaskId > 0 Then
        oQD.QueryFilters.Add "MIMessageCRFPageTaskId", "=", lEFormTaskId
    End If
    If lResponseId > 0 Then
        oQD.QueryFilters.Add "MIMessageResponseTaskId", "=", lResponseId
        oQD.QueryFilters.Add "MIMessageResponseCycle", "=", nResponseCycle
    End If
    
    Set oQDs = New QueryDefs
    oQDs.Init sCon, ""
    oQDs.Add oQD
    
    vData = oQDs.GetDataArray(sCODE)
    
    Set oQD = Nothing
    Set oQDs = Nothing
    
    If Not IsNull(vData) Then
        MIMessageExists = True
        lObjectId = vData(0, 0)
        nObjectSource = vData(1, 0)
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIDataLists.MIMessageExists"

End Function
 
'---------------------------------------------------------------------
Public Sub GetMIMsgStatusCount(sCon As String, sStudySiteWhereSQL As String, _
                                            ByRef lRaised As Long, ByRef lResponded As Long, ByRef lClosed As Long, _
                                            ByRef lPlanned As Long, lQueried As Long)
'---------------------------------------------------------------------
'return the no of discrepancies for each status in byref params:
'    lRaised disc
'    lResponded disc
'    lClosed disc
'    lQueried SDV
'    lPlanned SDV
'in a db for a the studysite combo sql passed in
' we join on the trialsubject table - as we do this when returning the full list
'---------------------------------------------------------------------
' REVISIONS
' DPH 13/10/2003 - performance enhancements to SQL
'---------------------------------------------------------------------
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim oCon As ADODB.Connection
Dim vStatus As Variant
Dim i As Long
    
    sSQL = "select MIMESSAGETYPE, MIMESSAGESTATUS, COUNT(*) from MIMESSAGE where"
    'TA 30/11/2003: no longer needed as using mimessagetrialname
    'sSQL = "select MIMESSAGETYPE, MIMESSAGESTATUS, COUNT(*) from MIMESSAGE, CLINICALTRIAL where"
    'sSQL = sSQL & " CLINICALTRIAL.CLINICALTRIALNAME = MIMESSAGE.MIMESSAGETRIALNAME and "
            
    sSQL = sSQL & " MIMESSAGETYPE in (" & MIMsgType.mimtDiscrepancy & "," & MIMsgType.mimtSDVMark & ")"
    
    sSQL = sSQL & " and MIMESSAGEHISTORY=0"
    
    If sStudySiteWhereSQL <> "" Then
        'add studysite sql if it exists
         sSQL = sSQL & " and " & sStudySiteWhereSQL
    End If
    
    sSQL = sSQL & " group by MIMESSAGETYPE,MIMESSAGESTATUS"

    Set oCon = New Connection
    Set rs = New Recordset
    
    oCon.Open sCon
    oCon.CursorLocation = adUseClient
    rs.Open sSQL, oCon, adOpenKeyset, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
        vStatus = rs.GetRows
    Else
        vStatus = Null
    End If
    rs.Close
    Set rs = Nothing
    oCon.Close
    Set oCon = Nothing
    
    'init to 0 - rows aren't returned if there are none for a certain status
    lRaised = 0
    lResponded = 0
    lClosed = 0
    lPlanned = 0
    lQueried = 0
    
    'match the first col wiht the status and set variable
    If Not IsNull(vStatus) Then
        For i = 0 To UBound(vStatus, 2)
            Select Case Val(vStatus(0, i))
            Case MIMsgType.mimtDiscrepancy
                Select Case Val(vStatus(1, i))
                Case eDiscrepancyMIMStatus.dsClosed
                    lClosed = CLng(vStatus(2, i))
                Case eDiscrepancyMIMStatus.dsRaised
                    lRaised = CLng(vStatus(2, i))
                Case eDiscrepancyMIMStatus.dsResponded
                    lResponded = CLng(vStatus(2, i))
                End Select
            Case MIMsgType.mimtSDVMark
                Select Case Val(vStatus(1, i))
                Case eSDVMIMStatus.ssPlanned
                    lPlanned = CLng(vStatus(2, i))
                Case eSDVMIMStatus.ssQueried
                    lQueried = CLng(vStatus(2, i))
                End Select
            End Select
        Next
    
    End If
    
    
End Sub
