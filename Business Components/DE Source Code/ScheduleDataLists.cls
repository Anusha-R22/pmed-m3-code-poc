VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScheduleDataLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'!! note that this is a class in work and is currently in the wrong layer
'----------------------------------------------------
' File: SubjectList.cls
' Author: Toby Aldridge
' Copyright: InferMed, Aug 2001, All Rights Reserved
' Retrieves subject list from database
'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
'
'   TA 03/10/2001: Restrict sites by UserName (SiteUser table)
'                    and restrict studies that don''t have sites
'   TA 15/10/2001: Do not allow return inactive sites for new subject datalists
'----------------------------------------------------

Option Explicit

Private Const SUBJECT_LIST As String = "SubjectList"
Private Const STUDY_LIST As String = "StudyList"
Private Const SITE_LIST As String = "SiteList"

'enumeration of the column indexes in the subject data array that is returned
Public Enum eSubjectListCol
    SubjectId = 0
    Site = 1
    StudyId = 2
    StudyName = 3
    SubjectLabel = 4
    SubjectStatus = 5
    Changed = 6
End Enum

'enumeration of the column indexes in the study data array that is returned
Public Enum eStudyListCol
    slStudyId = 0
    slStudyName = 1
End Enum

'enumeration of the column indexes in the site data array that is returned
Public Enum eSiteListCol
    Code = 0
    Description = 1
End Enum


Public Function GetStudyHeaders() As Variant
Attribute GetStudyHeaders.VB_Description = "Array of study col headers for use in UI (eg listview)."
'----------------------------------------------------
'Array of study col headers for use in UI (eg listview).
'----------------------------------------------------

    GetSubjectHeaders = Array("Id", "Name")
    
End Function

'---------------------------------------------------------------------
Public Function GetStudyList(sCon As String, sCurrentUserName As String, bExcludeClosedForRecruitment As Boolean) As Variant
Attribute GetStudyList.VB_Description = "2d array of studies. Studies without sites for this user are not returned."
'---------------------------------------------------------------------
' 2d array of studies.
' Studies without sites for this user are not returned.
' Studies with no active sties are not returned.
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs

    Set oQD = New QueryDef
    
    With oQD
        .InitSelect "ClinicalTrial", Array("ClinicalTrial.ClinicalTrialId", "ClinicalTrial.ClinicalTrialName")
        .Distinct = True
        'add trial site table
        .QueryTables.Add "TrialSite", , qdjtInner, "TrialSite.ClinicalTrialId", "ClinicalTrial.ClinicalTrialId"
        'add site user table
        .QueryTables.Add "SiteUser", , qdjtInner, "TrialSite.TrialSite", "SiteUser.Site"
        'don't want the library
        .QueryFilters.Add "ClinicalTrial.ClinicalTrialId", "<>", 0
        'restrict by user
        .QueryFilters.Add "SiteUser.UserName", "=", sCurrentUserName
        
        If bExcludeClosedForRecruitment Then
            .QueryFilters.Add "ClinicalTrial.StatusId", "<>", 3 '3 is closed for recruitment
        End If
        
        'TA 15/10/2001: Add site table to restrict on site status
        .QueryTables.Add "Site", , qdjtInner, "Site.Site", "TrialSite.TrialSite"
        .QueryFilters.Add "Site.SiteStatus", "=", 0
    End With
    
    oQD.Code = STUDY_LIST
    Set oQDs = New QueryDefs
    
    oQDs.Init sCon, ""
    oQDs.Add oQD
    
    GetStudyList = oQDs.GetDataArray(STUDY_LIST)
    
    Set oQD = Nothing
    Set oQDs = Nothing
    
    
End Function

Public Function GetSiteHeaders() As Variant
Attribute GetSiteHeaders.VB_Description = "Array of site col headers for use in UI (eg listview)."
'----------------------------------------------------
'Array of site col headers for use in UI (eg listview).
'----------------------------------------------------

    GetSubjectHeaders = Array("Site", "Description")
    
End Function

'---------------------------------------------------------------------
Public Function GetSiteList(sCon As String, lStudyId As Long, sCurrentUserName As String) As Variant
Attribute GetSiteList.VB_Description = "2d array of sites."
'---------------------------------------------------------------------
' 2d array of sites.
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs

    Set oQD = New QueryDef
    
    With oQD
        .InitSelect "Site", Array("Site.Site", "Site.SiteDescription")
        .QueryTables.Add "TrialSite", , qdjtInner, "TrialSite.TrialSite", "Site.Site"
        'add site user table
        .QueryTables.Add "SiteUser", , qdjtInner, "TrialSite.TrialSite", "SiteUser.Site"
        .QueryFilters.Add "ClinicalTrialId", "=", lStudyId
        'restrict by user
        .QueryFilters.Add "SiteUser.UserName", "=", sCurrentUserName
        'TA 15/10/2001: Add site table to restrict on site status
        .QueryFilters.Add "Site.SiteStatus", "=", 0
    End With
    
    oQD.Code = SITE_LIST
    Set oQDs = New QueryDefs
    
    oQDs.Init sCon, ""
    oQDs.Add oQD
    
    GetSiteList = oQDs.GetDataArray(SITE_LIST)
    
    Set oQD = Nothing
    Set oQDs = Nothing
    
    
End Function

Public Function GetSubjectHeaders() As Variant
Attribute GetSubjectHeaders.VB_Description = "Array of subject col headers for use in UI (eg listview)."
'----------------------------------------------------
'Array of subject col headers for use in UI (eg listview).
'----------------------------------------------------

    GetSubjectHeaders = Array("Id", "Site", "Study Id", "Study", "Label", "Status", "New")
    
End Function
'---------------------------------------------------------------------
Public Function GetSubjectList(sCon As String, sLabel As String, sStudy As String, sSite As String, lSubjectId As Long, sCurrentUserName As String) As Variant
Attribute GetSubjectList.VB_Description = "2d array of subjects with a like filter on sLabel, sStudy and sSite and an equals filter on lSubjectId. Empty strings or a SubjectId of null integer will result in no filter on that parameter."
'---------------------------------------------------------------------
' 2d array of subjects with a like filter on sLabel, sStudy and sSite
' and an equals filter on lSubjectId.
' Empty strings or a SubjectId of null integer will result in no filter on that parameter.
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs

    Set oQD = New QueryDef
    
    With oQD
        .QueryTables.Add "ClinicalTrial"
        .QueryTables.Add "TrialSubject", , qdjtInner, "TrialSubject.ClinicalTrialId", "ClinicalTrial.ClinicalTrialId"
        .QueryFields.Add Array("PersonId", "TrialSite", "ClinicalTrial.ClinicalTrialId", "ClinicalTrial.ClinicalTrialName", _
                                    "LocalIdentifier1", "TrialSubject.TrialSubjectStatus", "TrialSubject.Changed")
        .QueryOrders.Add Array("LocalIdentifier1", "ClinicalTrialName", "TrialSite", "PersonId")
    
        'extra code to restrict on site user
        If sCurrentUserName <> "" Then
            'add site user table (using TrialSubject)
            oQD.QueryTables.Add "SiteUser", , qdjtInner, "TrialSubject.TrialSite", "SiteUser.Site"
            'restrict by user
            oQD.QueryFilters.Add "SiteUser.UserName", "=", sCurrentUserName
        End If
    
    
        'study
        If sStudy <> "" Then
            .UserFilters.Add "ClinicalTrial.ClinicalTrialName", "like", "%" & sStudy & "%"
        End If
    
        'site
        If sSite <> "" Then
            .UserFilters.Add "TrialSubject.TrialSite", "like", "%" & sSite & "%"
        End If
    
        'subject label
        If sLabel <> "" Then
            .UserFilters.Add "TrialSubject.LocalIdentifier1", "like", "%" & sLabel & "%"
        End If
    
        'subject id
        If lSubjectId <> NULL_LONG Then
            'TA 20/6/2000 SR 3634: if not blank use criteria
            .UserFilters.Add "TrialSubject.PersonId", "=", lSubjectId
        End If

    End With
    
    oQD.Code = SUBJECT_LIST
    Set oQDs = New QueryDefs
    
    oQDs.Init sCon, ""
    oQDs.Add oQD
    
    GetSubjectList = oQDs.GetDataArray(SUBJECT_LIST)
    
    Set oQD = Nothing
    Set oQDs = Nothing
    
    
End Function

 
 
 
 
