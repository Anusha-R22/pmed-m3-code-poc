VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserDataLists"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002-2003. All Rights Reserved
'   File:       UserDataLists.cls
'   Author:     Toby Aldridge, November 2002
'   Purpose:    Class for reurning lists of data to fill combos etc
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 7 Mar 03 - Added subject timestamp columns
'   DPH 08/10/2003 - Performance enhancements in StudiesSitesWhereSQL
' NCJ 26 Nov 03 - Store user as string rather than as object (to avoid non-termination problems)
' NCJ 27 Nov 03 - Tidy up on termination
' DPH 20/01/2004 - If no sites exist for a specific study return 1=1 sql in StudiesSitesWhereSQL
' DPH 24/03/2004 - Clean up collections in class_terminate
' DPH 02/03/2007 - bug 2882. Changes to GetSubjectList & StudiesSitesWhereSQL for performance reasons
'----------------------------------------------------------------------------------------'

Option Explicit

'enumeration of the column indexes in the subject data array that is returned
Public Enum eSubjectListCols
    SubjectId = 0
    Site = 1
    StudyId = 2
    StudyName = 3
    SubjectLabel = 4
    SubjectStatus = 5
    Changed = 6
    SubjectTimeStamp = 7
    SubjectTimeStampTZ = 8
    DiscStatus = 9
    SDVStatus = 10
    LockStatus = 11
End Enum

'enumeration of the column indexes in the subject data array that is returned from GetLabsForSiteList
Public Enum eLabCols
    lcCode = 0
    lcDescription = 1
End Enum

'enum for drop lists cols
'col 0 is name and col 1 is id
Public Enum eDropDownCol
    Text = 0
    Id = 1
End Enum


'store reference to parent object
' NCJ 26 Nov 03 - Store user as string rather than as object (to avoid non-termination problems)
' Private moUser As MACROUser
Private msUser As String

Private mcolStudies As Collection

'store the sort orders used in GetSubjectList
Private mcolSubjectSortOrder As Collection

'----------------------------------------------------------------------------------------'
Private Property Get TheUser() As MACROUser
'----------------------------------------------------------------------------------------'
' NCJ 26 Nov 03 - Property which creates user object "on the fly"
'----------------------------------------------------------------------------------------'
Dim oUser As MACROUser

    Set oUser = New MACROUser
    Call oUser.SetState(msUser)
    Set TheUser = oUser
    Set oUser = Nothing

End Property

'----------------------------------------------------------------------------------------'
Public Function GetStudyItem(enObject As eStudyObject, lStudyId As Long, Optional lObjectId As Long = 0) As StudyItem
'----------------------------------------------------------------------------------------'
'lobjectid is optional when getting a study
'----------------------------------------------------------------------------------------'
Dim oStudyItem As New StudyItem

    If Not CollectionMember(mcolStudies, "k" & lStudyId) Then
        Call LoadStudy(lStudyId)
    End If
    
   'Set GetStudyItem = mcolStudies("k" & lStudyId).Item("k" & lStudyId & "|" & enObject & "|" & lObjectId)
   Set GetStudyItem = mcolStudies("k" & lStudyId).Item(oStudyItem.GetKey(lStudyId, enObject, lObjectId))
   
   Set oStudyItem = Nothing

End Function

'----------------------------------------------------------------------------------------'
Public Function GetStudyItemCode(enObject As eStudyObject, lStudyId As Long, Optional lObjectId As Long = 0) As String
'----------------------------------------------------------------------------------------'
'lobjectid is optional when getting a study
'----------------------------------------------------------------------------------------'

    GetStudyItemCode = GetStudyItem(enObject, lStudyId, lObjectId).Code

End Function

'----------------------------------------------------------------------------------------'
Public Function GetStudyItemName(enObject As eStudyObject, lStudyId As Long, Optional lObjectId As Long = 0) As String
'----------------------------------------------------------------------------------------'
'lobjectid is optional when getting a study
'----------------------------------------------------------------------------------------'

    GetStudyItemName = GetStudyItem(enObject, lStudyId, lObjectId).Name

End Function

'----------------------------------------------------------------------------------------'
Private Sub LoadStudy(lStudyId As Long)
'----------------------------------------------------------------------------------------'
'load all the information in for a specific company
'----------------------------------------------------------------------------------------'
' REVISIONS
' DPH 22/03/2004 - Use oUser object can destroy rather than TheUser function
'                   suspect memory leak
'---------------------------------------------------------------------
Dim sSQL As String
Dim sWhere As String
Dim oCon As Connection
Dim rs As Recordset
Dim vData As Variant
Dim colItems As Collection
Dim oStudyItem As StudyItem
Dim i As Long
Dim oUser As MACROUser

    Set oUser = TheUser

    sWhere = " WHERE CLINICALTRIALID = " & lStudyId
         
    sSQL = sSQL & "SELECT " & eStudyObject.soVisit & ", CLINICALTRIALID, VISITCODE, VISITNAME, VISITID FROM STUDYVISIT" & sWhere
    sSQL = sSQL & "UNION SELECT " & eStudyObject.soeform & ", CLINICALTRIALID, CRFPAGECODE, CRFTITLE, CRFPAGEID FROM CRFPAGE" & sWhere
    sSQL = sSQL & "UNION SELECT " & eStudyObject.soQuestion & ", CLINICALTRIALID, DATAITEMCODE, DATAITEMNAME, DATAITEMID FROM DATAITEM" & sWhere
    
    Set oCon = New Connection
    oCon.Open oUser.CurrentDBConString
    oCon.CursorLocation = adUseClient
    Set rs = New Recordset
    rs.Open sSQL, oCon
    vData = rs.GetRows
    rs.Close
    Set rs = Nothing
    oCon.Close
    Set oCon = Nothing
    
    Set colItems = New Collection
    
    'add id of 0 showing blank name and code
    Set oStudyItem = New StudyItem
    With oStudyItem
        .StudyId = lStudyId
        .Code = ""
        .Name = ""
        .Id = 0
        
        .StudyObject = eStudyObject.soVisit
        colItems.Add oStudyItem, oStudyItem.Key
        
        .StudyObject = eStudyObject.soeform
        colItems.Add oStudyItem, oStudyItem.Key
        
        .StudyObject = eStudyObject.soQuestion
        colItems.Add oStudyItem, oStudyItem.Key
        
    End With

            

    For i = 0 To UBound(vData, 2)
        Set oStudyItem = New StudyItem
        With oStudyItem
            .StudyObject = vData(0, i)
            .StudyId = vData(1, i)
            .Code = vData(2, i)
            .Name = vData(3, i)
            .Id = vData(4, i)
        End With
        colItems.Add oStudyItem, oStudyItem.Key
    Next
        
    mcolStudies.Add colItems, "k" & lStudyId
    
    ' NCJ 27 Nov 03 - Tidy up on exit
    Set colItems = Nothing
    Set oStudyItem = Nothing
    Set oUser = Nothing
    
End Sub

'---------------------------------------------------------------------
Friend Sub Init(oUser As MACROUser)
'---------------------------------------------------------------------

'---------------------------------------------------------------------

'    Set moUser = oUser
    msUser = oUser.GetState(False)
    Set mcolStudies = New Collection
    
End Sub


'---------------------------------------------------------------------
Public Function GetSubjectList(Optional sLabel As String = "", Optional sStudyName As String = "", _
                                Optional sSite As String = "", Optional lSubjectId As Long = -1, _
                                Optional enSortOrder As eSubjectListCols = -1, Optional bAscend As Boolean = True) As Variant
'---------------------------------------------------------------------
' 2d array of subjects with a like filter on sLabel, sStudy and sSite
' and an equals filter on lSubjectId.
' Empty strings or a SubjectId of null integer will result in no filter on that parameter.
' NCJ 7 Mar 03 - Added SubjectTimeStamp and SubjectTimeStamp_TZ
' DPH 22/03/2004 - Use oUser object can destroy rather than TheUser function
'                   suspect memory leak
' DPH 27/02/2007 - Function calls StudiesSitesWhereSQL multiple times, once only required
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs
Dim sStudySiteSQl As String
Dim sOrderSQl As String
Dim i As Long
Dim oUser As MACROUser

    Set oUser = TheUser
    Set oQD = New QueryDef
    
    With oQD
        .QueryTables.Add "ClinicalTrial"
        .QueryTables.Add "TrialSubject", , qdjtInner, "TrialSubject.ClinicalTrialId", "ClinicalTrial.ClinicalTrialId"
        .QueryFields.Add Array("PersonId", "TrialSite", "ClinicalTrial.ClinicalTrialId", "ClinicalTrial.ClinicalTrialName", _
                                    "LocalIdentifier1", "TrialSubject.TrialSubjectStatus", "TrialSubject.Changed", _
                                    "SubjectTimeStamp", "SubjectTimeStamp_TZ", "DISCREPANCYSTATUS", "SDVSTATUS", "LOCKSTATUS")
        'ordering now done dynamically
        '.QueryOrders.Add Array("LocalIdentifier1", "ClinicalTrialName", "TrialSite", "PersonId")

        'study
        If sStudyName <> "" Then
            .QueryFilters.Add "ClinicalTrial.ClinicalTrialName", "=", sStudyName
        End If
    
        'site
        If sSite <> "" Then
            .QueryFilters.Add "TrialSubject.TrialSite", "=", sSite
        End If
    
        'subject label
        If sLabel <> "" Then
            .QueryFilters.Add "TrialSubject.LocalIdentifier1", "like", "%" & sLabel & "%"
        End If
    
        'subject id
        If lSubjectId <> -1 Then
            'TA 20/6/2000 SR 3634: if not blank use criteria
            .QueryFilters.Add "TrialSubject.PersonId", "=", lSubjectId
        End If
        
        sStudySiteSQl = StudiesSitesWhereSQL("TrialSubject.ClinicalTrialId", "TrialSubject.TrialSite", oUser)
        
        If sStudySiteSQl <> "" Then
            'filter on the allowed study site combinations
            ' DPH 27/02/2007 - Function calls StudiesSitesWhereSQL multiple times, once only required - use variable
            .QueryFilters.Add sStudySiteSQl, "user defined"
        End If
        
        If enSortOrder <> -1 Then
            'they have added a sort order
            Select Case enSortOrder

            Case SubjectId: sOrderSQl = "PERSONID"
            Case Site: sOrderSQl = "TRIALSITE"
            Case StudyId: sOrderSQl = "CLINICALTRIALID"
            Case StudyName: sOrderSQl = "CLINICALTRIALNAME"
            Case SubjectLabel: sOrderSQl = "LOCALIDENTIFIER1"
            Case SubjectStatus: sOrderSQl = "STATUS"
            Case Changed: sOrderSQl = "CHANGED"
            Case SubjectTimeStamp: sOrderSQl = "SUBJECTTIMESTAMP"
            Case SubjectTimeStampTZ: sOrderSQl = "SUBJECTTIMESTAMP_TZ"
            Case DiscStatus: sOrderSQl = "DISCREPANCYSTATUS"
            Case SDVStatus: sOrderSQl = "SDVSTATUS"
            Case LockStatus: sOrderSQl = "LOCKSTATUS"
            End Select
            bAscend = True
            'firstly remove previous references to this column if there
            For i = 1 To mcolSubjectSortOrder.Count
                If InStr(CStr(mcolSubjectSortOrder(i)), sOrderSQl) > 0 Then
                    If i = mcolSubjectSortOrder.Count Then
                        'if this is the last one then calculate bascend
                        bAscend = (InStr(CStr(mcolSubjectSortOrder(i)), "DESC") > 0)
                    End If
                    mcolSubjectSortOrder.Remove (i)
                    Exit For
                End If
            Next
            
            mcolSubjectSortOrder.Add sOrderSQl & IIf(bAscend, "", " DESC") 'add desc if not ascending
        End If
        
    
        'add sort orders
        For i = mcolSubjectSortOrder.Count To 1 Step -1
            .QueryOrders.Add mcolSubjectSortOrder(i)
        Next
        
    End With
    
    
    oQD.Code = "SUBJECT_LIST"
    Set oQDs = New QueryDefs
    
    oQDs.Init oUser.CurrentDBConString, ""
    oQDs.Add oQD
    
    GetSubjectList = oQDs.GetDataArray("SUBJECT_LIST")
    
    Set oQD = Nothing
    Set oQDs = Nothing
    Set oUser = Nothing
    
End Function


'---------------------------------------------------------------------
Public Function GetVisitList(Optional lStudyId As Long = -1) As Variant
'---------------------------------------------------------------------
' retrieves list of Visits as a 2d array for a combo.
' First col is the text(name) the second is the id.
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs
Dim oUser As MACROUser

    Set oQD = New QueryDef
    Set oUser = TheUser
    
    With oQD
        .InitSelect "StudyVisit", Array("StudyVisit.VisitName", "StudyVisit.VisitId")
        
        If lStudyId <> -1 Then
              .QueryFilters.Add "StudyVisit.ClinicalTrialId", "=", lStudyId
        End If
        
        'filter on the allowed study site combinations
        .QueryFilters.Add "StudyVisit.ClinicalTrialId", "IN", StudyIdsArray(oUser.GetAllStudies)
        .Distinct = True
    End With
    
    oQD.Code = "temp"
    Set oQDs = New QueryDefs
    
    oQDs.Init oUser.CurrentDBConString, ""
    oQDs.Add oQD
    
    GetVisitList = oQDs.GetDataArray("temp")
    
    Set oQD = Nothing
    Set oQDs = Nothing
    Set oUser = Nothing
    
End Function

'---------------------------------------------------------------------
Public Function GetEFormList(Optional lStudyId As Long = -1, Optional lVisitId As Long = 0) As Variant
'---------------------------------------------------------------------
' retrieves list of eForms as a 2d array for a combo.
' First col is the text(CRFTitle) the second is the id.
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs
Dim oUser As MACROUser

    Set oQD = New QueryDef
    Set oUser = TheUser
    
    With oQD
        .InitSelect "CRFPage", Array("CRFPage.CRFTitle", "CRFPage.CRFPageId")
        
        If lStudyId <> -1 Then
              .QueryFilters.Add "CRFPage.ClinicalTrialId", "=", lStudyId
        End If
        
        If lVisitId <> 0 Then
            .QueryTables.Add "StudyVisitCRFPage", , qdjtInner, _
            Array("StudyVisitCRFPage.ClinicalTrialId", "StudyVisitCRFPage.CRFPageId"), _
            Array("CRFPage.ClinicalTrialId", "CRFPage.CRFPageId")
            .QueryFilters.Add "StudyVisitCRFPage.VisitId", "=", lVisitId
        End If
        
        'filter on the allowed study site combinations
        .QueryFilters.Add "CRFPage.ClinicalTrialId", "IN", StudyIdsArray(oUser.GetAllStudies)
        
        .Distinct = True
    End With
    
    oQD.Code = "temp"
    Set oQDs = New QueryDefs
    
    oQDs.Init oUser.CurrentDBConString, ""
    oQDs.Add oQD
    
    GetEFormList = oQDs.GetDataArray("temp")
    
    Set oQD = Nothing
    Set oQDs = Nothing
    Set oUser = Nothing
    
End Function

'---------------------------------------------------------------------
Public Function GetQuestionList(Optional lStudyId As Long = -1, Optional lEFormId As Long = 0) As Variant
'---------------------------------------------------------------------
' retrieves list of questions as a 2d array for a combo.
' First col is the text (question name) the second is the id.
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs
Dim oUser As MACROUser

    Set oQD = New QueryDef
    Set oUser = TheUser
    
    With oQD
        .InitSelect "DataItem", Array("DataItem.DataItemName", "DataItem.DataITemID")
        
        If lStudyId <> -1 Then
              .QueryFilters.Add "DataItem.ClinicalTrialId", "=", lStudyId
        End If
        
        If lEFormId <> 0 Then
            .QueryTables.Add "CRFElement", , qdjtInner, _
            Array("DataItem.ClinicalTrialId", "DataITem.DataItemId"), _
            Array("CRFElement.ClinicalTrialId", "CRFElement.DataItemId")
            .QueryFilters.Add "CRFElement.CRFPageId", "=", lEFormId
        
            'filter on the allowed study site combinations
            .QueryFilters.Add "DataItem.ClinicalTrialId", "IN", StudyIdsArray(oUser.GetAllStudies)
            
    
        End If
        
        .Distinct = True
    End With
    
    oQD.Code = "temp"
    Set oQDs = New QueryDefs
    
    oQDs.Init oUser.CurrentDBConString, ""
    oQDs.Add oQD
    
    GetQuestionList = oQDs.GetDataArray("temp")
    
    Set oQD = Nothing
    Set oQDs = Nothing
    Set oUser = Nothing
    
End Function

'---------------------------------------------------------------------
Public Function GetUserList(Optional lStudyId As Long = -1, Optional sSite As String = "") As Variant
'---------------------------------------------------------------------
' retrieves list of users for a dropdown as a 2d array for a combo.
' First col is the text (UserName) the second is also the username.
' clinicaltrial id not currently used
'---------------------------------------------------------------------
' REVISIONS
' DPH 22/03/2004 - Use oUser object can destroy rather than TheUser function
'                   suspect memory leak
' MLM 02/06/2008: Issue 3004: Order the list of users
'---------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQDs As QueryDefs
Dim oUser As MACROUser

    Set oUser = TheUser
    Set oQD = New QueryDef
    
    With oQD
        .InitSelect "UserRole", Array("UserName", "UserName UName"), , , Array("UserName")
        
        .Distinct = True
    End With
    
    
    
    oQD.Code = "temp"
    Set oQDs = New QueryDefs
    
    oQDs.Init oUser.CurrentDBConString, ""
    oQDs.Add oQD
    
    GetUserList = oQDs.GetDataArray("temp")
    
    Set oQD = Nothing
    Set oQDs = Nothing
    Set oUser = Nothing
    
End Function

'---------------------------------------------------------------------
Public Function StudiesSitesWhereSQL(sStudyCol As String, sSiteCol As String, _
                        Optional ByRef oUser As MACROUser = Nothing) As String
'---------------------------------------------------------------------
'Return SQL to filter on study dsite combinations
'---------------------------------------------------------------------
' REVISIONS
' DPH 08/10/2003 - Performance enhancements - use IN instead of multiple =
'       If a numeric is passed into sStudyCol then just produce a list of
'       sites for that study
'       If studycol is a numeric then are only dealing with that particular
'       study so just want to produce a site clause
'       TODO - might want as a sepearate function for MACRO 3.1
'       done here to avoid breaking compatibility
' DPH 20/01/2004 - If no sites exist for a specific study return 1=1 sql
' DPH 22/03/2004 - Use oUser object can destroy rather than TheUser function
'                   suspect memory leak
' DPH 27/02/2007 - Bug 2882: StudiesSitesWhereSQL is slow for databases with many
'                   Studies and sites. Replace concatenation with array string
'                   building. Changes to all SQL collating code.
'                   Pass in optional user object to avoid creating from serialized.
'---------------------------------------------------------------------
Dim oStudySite As StudySite
Dim sSQL As String
Dim i As Long
Dim colStudySite As Collection
Dim nStudyId As Long
Dim nOnlyStudyId As Long
Dim bProcessedSQL As Boolean
Dim bAnySQL As Boolean
'Dim oUser As MACROUser
Dim vWhereSQL() As String
Dim bUserPassed As Boolean

    If IsMissing(oUser) Or (oUser Is Nothing) Then
        Set oUser = TheUser
        bUserPassed = False
    Else
        bUserPassed = True
    End If
    
    Set colStudySite = oUser.GetStudiesSites
    
    ' DPH 20/01/2004 - initialise anysql variable
    bAnySQL = False
    
    ReDim vWhereSQL(0)

    If colStudySite.Count = 0 Then
        'they have no access to studies and sites
        Call AddStringToVarArr(vWhereSQL, "1=0")
        'sSQL = "1=0"
        bAnySQL = True
    Else
        Call AddStringToVarArr(vWhereSQL, "(")
        'sSQL = "("
        nStudyId = -1
        bProcessedSQL = False
        
        ' get only study (if numeric)
        If IsNumeric(sStudyCol) Then
            nOnlyStudyId = CLng(sStudyCol)
        Else
            nOnlyStudyId = -1
        End If
        
        For i = 1 To colStudySite.Count
            Set oStudySite = colStudySite(i)
            
            ' only perform this code if getting all studies or a particular study
            If nOnlyStudyId = -1 Or nOnlyStudyId = oStudySite.StudyId Then
                ' DPH 08/10/2003 - performance enhancement changes
                If nStudyId <> oStudySite.StudyId Then
                    ' SQL for study and site (using IN)
                    If bProcessedSQL Then
                        ' close previous clause
                        Call AddStringToVarArr(vWhereSQL, "))) OR ")
                        'sSQL = sSQL & "))) OR "
                    End If
                    If nOnlyStudyId = -1 Then
                        ' need study clause
                        'TA 30/11/2003: if using mimessage table then use study name
                        If UCase(sStudyCol) = "MIMESSAGETRIALNAME" Then
                            'sSQL = sSQL & "((" & sStudyCol & "='" & oStudySite.StudyName & "') AND "
                            Call AddStringToVarArr(vWhereSQL, "((")
                            Call AddStringToVarArr(vWhereSQL, sStudyCol)
                            Call AddStringToVarArr(vWhereSQL, "='")
                            Call AddStringToVarArr(vWhereSQL, oStudySite.StudyName)
                            Call AddStringToVarArr(vWhereSQL, "') AND ")
                        Else
                            'original code
                            'sSQL = sSQL & "((" & sStudyCol & "=" & oStudySite.StudyId & ") AND "
                            Call AddStringToVarArr(vWhereSQL, "((")
                            Call AddStringToVarArr(vWhereSQL, sStudyCol)
                            Call AddStringToVarArr(vWhereSQL, "=")
                            Call AddStringToVarArr(vWhereSQL, oStudySite.StudyId)
                            Call AddStringToVarArr(vWhereSQL, ") AND ")
                        End If
                    End If
                    'sSQL = sSQL & "(" & sSiteCol & " IN ('" & oStudySite.Site & "'"
                    Call AddStringToVarArr(vWhereSQL, "(")
                    Call AddStringToVarArr(vWhereSQL, sSiteCol)
                    Call AddStringToVarArr(vWhereSQL, " IN ('")
                    Call AddStringToVarArr(vWhereSQL, oStudySite.Site)
                    Call AddStringToVarArr(vWhereSQL, "'")
                    nStudyId = oStudySite.StudyId
                    bProcessedSQL = True
                    bAnySQL = True
                Else
                    ' continue site IN clause as same study
                    'sSQL = sSQL & ",'" & oStudySite.Site & "'"
                    Call AddStringToVarArr(vWhereSQL, ",'")
                    Call AddStringToVarArr(vWhereSQL, oStudySite.Site)
                    Call AddStringToVarArr(vWhereSQL, "'")
                    bAnySQL = True
                End If
            End If
        Next
        
        ' close last clause
        If nOnlyStudyId = -1 Then
            Call AddStringToVarArr(vWhereSQL, ")))")
            'sSQL = sSQL & ")))"
        Else
            Call AddStringToVarArr(vWhereSQL, "))")
            'sSQL = sSQL & "))"
        End If
        
        ' close complete clause
        Call AddStringToVarArr(vWhereSQL, ")")
        'sSQL = sSQL & ")"
    End If
    
    ' DPH 20/01/2004 - If no sites exist for a specific study return 1=1 sql
    If Not bAnySQL Then
        Call AddStringToVarArr(vWhereSQL, "(1=1)")
        'sSQL = "(1=1)"
    End If
    
    StudiesSitesWhereSQL = Join(vWhereSQL, "")
    'StudiesSitesWhereSQL = sSQL
    
    ' NCJ 27 Nov 03 - Tidy up on exit
    Set oStudySite = Nothing
    Set colStudySite = Nothing
    If Not bUserPassed Then
        Set oUser = Nothing
    End If
    
End Function

'--------------------------------------------------------------------------------------------------
Private Sub AddStringToVarArr(ByRef vArr() As String, ByVal sData As String)
'--------------------------------------------------------------------------------------------------
' Add string to variant array
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
    Dim lArray As Long
    Dim bAddNew As Boolean
    
    On Error GoTo CatchAllError
    
    lArray = UBound(vArr)
    bAddNew = True
    
    If lArray = 0 And (IsEmpty(vArr(0)) Or (vArr(0) = "")) Then
        vArr(0) = sData
        bAddNew = False
    End If
    
    If bAddNew Then
        ReDim Preserve vArr(lArray + 1)
        vArr(lArray + 1) = sData
    End If
    Exit Sub

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|UserDataLists.AddStringToVarArray")
End Sub

'---------------------------------------------------------------------
Public Function GetLabsForSiteList(sSite As String) As Variant
'---------------------------------------------------------------------
'return list of labs for specific site
'---------------------------------------------------------------------
' DPH 22/03/2004 - Use oUser object can destroy rather than TheUser function
'                   suspect memory leak
'---------------------------------------------------------------------
Dim sSQL As String
Dim oCon As Connection
Dim rs As Recordset
Dim vData As Variant
Dim oUser As MACROUser
         
    Set oUser = TheUser
    
    sSQL = "SELECT Laboratory.LaboratoryCode, Laboratory.LaboratoryDescription" _
                & " FROM Laboratory, SiteLaboratory " _
                & " WHERE Laboratory.LaboratoryCode = SiteLaboratory.LaboratoryCode and SiteLaboratory.Site = '" & sSite & "'" _
                & " ORDER BY Laboratory.LaboratoryCode"

    Set oCon = New Connection
    oCon.Open oUser.CurrentDBConString
    oCon.CursorLocation = adUseClient
    
    Set rs = New Recordset
    rs.Open sSQL, oCon
    If rs.RecordCount > 0 Then
        vData = rs.GetRows
    Else
        vData = Null
    End If
    
    rs.Close
    Set rs = Nothing
    oCon.Close
    Set oCon = Nothing
    Set oUser = Nothing
    
    GetLabsForSiteList = vData
    
End Function

'---------------------------------------------------------------------
Private Function StudyIdsArray(colStudies As Collection) As Variant
'---------------------------------------------------------------------
'Return array of study ids
'---------------------------------------------------------------------
Dim vStudies As Variant
Dim i As Long

    ReDim vStudies(colStudies.Count - 1) As Long
    
    For i = 1 To colStudies.Count
        vStudies(i - 1) = colStudies(i).StudyId
    Next
    
    StudyIdsArray = vStudies
    
End Function

'----------------------------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------------------------
    
        'default order for subject list (study,site,subject)
        Set mcolSubjectSortOrder = New Collection
        mcolSubjectSortOrder.Add "PERSONID"
        mcolSubjectSortOrder.Add "LOCALIDENTIFIER1"
        mcolSubjectSortOrder.Add "TRIALSITE"
        mcolSubjectSortOrder.Add "CLINICALTRIALNAME"
    
End Sub

'----------------------------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------------------------
' NCJ 27 Nov 03 - Tidy up on termination
' DPH 24/03/2004 - terminate object
'----------------------------------------------------------------------

    Call Terminate
    
End Sub


'---------------------------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'---------------------------------------------------------------------
'these are not currently used TA 17/11/2002
'could be used to serialise for web
'---------------------------------------------------------------------

    Set mcolSubjectSortOrder = CollectionDeSerialise(PropBag.ReadProperty("sso"))
'Dim lCount As Long
'Dim oCol As Collection
'Dim i As Long
'
'    msCon = PropBag.ReadProperty("s")
'    lCount = PropBag.ReadProperty("c")
'
'    Set mcolStudies = New Collection
'    For i = 0 To lCount
'        Set oCol = CollectionDeSerialise(PropBag.ReadProperty(CStr(i)), True)
'        mcolStudies.Add oCol, "k" & oCol(1).StudyId
'    Next
'
'
End Sub
'
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'Dim lCount As Long
'Dim i As Long
'

    PropBag.WriteProperty "sso", CollectionSerialise(mcolSubjectSortOrder)

'    PropBag.WriteProperty "s", msCon
'
'    lCount = mcolStudies.Count
'    PropBag.WriteProperty "c", lCount
'
'
'    For i = 0 To mcolStudies.Count
'        PropBag.WriteProperty CStr(i), CollectionSerialise(mcolStudies(i))
'
'
'    Next
'
End Sub

'----------------------------------------------------------------------------------------'
Friend Sub Terminate()
'----------------------------------------------------------------------------------------'
' DPH 24/03/2004 - terminate object
'----------------------------------------------------------------------------------------'

    Call CollectionRemoveAll(mcolStudies)
    Set mcolStudies = Nothing
    Call CollectionRemoveAll(mcolSubjectSortOrder)
    Set mcolSubjectSortOrder = Nothing

End Sub
