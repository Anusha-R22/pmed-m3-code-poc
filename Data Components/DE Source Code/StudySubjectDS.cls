VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudySubjectDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'----------------------------------------------------
' File: StudySubjectDS.cls
' Copyright InferMed Ltd. 2001-2002 All Rights Reserved
' Author: Nicky Johns, InferMed, May 2001
' Purpose: Data Services for Study Subject for MACRO 3.0
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 18-29 May 01 - Initial development
' NCJ 27 Jun 01 - Use same Query Server to load Visit and eForm instances
' NCJ 3 Jul 01 - Load eFormInstances here (not in VisitInstances)
' TA 11/07/2001 - Code for saving
' NCJ 17 Sep 01 - SaveRegistration and SaveRegistrationStatus
' TA 21 Sep 01 - All field and table names now are constants - no schema name within subs any more
'   TA 01/10/2001 - Error handling added/classes combined
' NCJ/TA 3 Oct 01 - Fixes to LoadNRandCTC
'                   Added LabCode to SaveEFormInstances
' NCJ 15 Oct 01 - Make sure NRStatus is read in as an integer
' MLM 6/12/01: Include Changed flag when saving subject.
' TA 05/04/2002: Include HadValue flag when saving
' TA 05/04/2002: Changes to update AREZZOToken table
' TA 09/05/2002: Make updating cahce table optional
' MACRO 3.0
' NCJ 5 Jul 2002 - Brought in RQG changes from prev. 3.0
' TA 23/08/2002: Add new MIMessage statuses when loading only
' TA 23/08/2002: Add new MIMessage statuses when saving for the first time
' TA 28/08/2002: Corrected new MIMessage statuses
' RS 25/10/2002: Pass 0 as DatabaseTimestamp in SaveResponses will cause new trigger to add correct databasetimestamp
' NCJ 14 Nov 02 - Removed loading/saving of QGroupInstances
' NCJ 22 Jan 03 - Added StandardValue to Response fields
' NCJ 28 Jan 03 - Added UserNameFull to Response fields
' NCJ 7 Mar 03 - Added Subject time stamp fields
' NCJ 29 Apr 03 - Do not save Responses with IgnoreMe = TRUE
'----------------------------------------------------

Option Explicit

Private Const msDELIMITER = "|"


' The Subject table name
Private Const msSTUDY_SUBJECT = "TrialSubject"
' The Visit Instance table
Private Const msVI_TABLE = "VisitInstance"
' The eForm Instance table
Private Const msEFORMI_TABLE = "CRFPageInstance"

'changed column - occurs in VisitInstance and CRFPageInstace tables
Private Const msCHANGED_COL = "Changed"

' The columns of VisitInstance information
Private Const msVI_COLS = "VisitTaskID|VisitID|VisitCycleNumber|" _
                        & "VisitStatus|VisitDate|LockStatus"

' The subset of columns of VisitInstance information to update
Private Const msVI_UPDATE_COLS = "VisitStatus|VisitDate|Changed"

'constants for cols than are used duirng the load but are never saved
Private Const msVI_LOADONLY_COLS = "DiscrepancyStatus|SDVStatus|NoteStatus"

Private Const msVI_TASKIDCOL = "VisitTaskID"

' The columns of subject information
Private Const msSUBJECT_COLS = "LocalIdentifier1|LocalIdentifier2|DateOfBirth|" _
                            & "TrialSubjectStatus|LockStatus|SubjectGender|" _
                            & "RegistrationStatus|SubjectTimeStamp|SubjectTimeStamp_TZ"
                            
'constants for cols than are used duirng the load but are never saved
Private Const msSUBJECT_LOADONLY_COLS = "DiscrepancyStatus|SDVStatus|NoteStatus"
                            
Private Const msSUBJECT_PROFORMACOL = "ProformaState"

'cols reuired to calculate new subject id
 Private Const msSUBJECT_IDCOL = "PersonID"
 Private Const msSUBJECT_NEWIDFILTERCOLS = "ClinicalTrialId|TrialSite"


' The columns of eFormInstance information
Private Const msEFORMI_COLS = "CRFPageTaskID|CRFPageID|CRFPageCycleNumber|" _
                & "VisitID|VisitCycleNumber|CRFPageInstanceLabel|" _
                & "CRFPageDate|CRFPageStatus|LockStatus|LaboratoryCode"

' The subset of columns of eFormInstance information to update
Private Const msEFORMI_UPDATE_COLS = "CRFPageInstanceLabel|CRFPageDate|CRFPageStatus|LaboratoryCode|Changed"

'constants for cols than are used duirng the load but are never saved
Private Const msEFORMI_LOADONLY_COLS = "DiscrepancyStatus|SDVStatus|NoteStatus"

Private Const msEFORMI_TASKID_COL = "CRFPageTaskId"

' The "key" columns, StudyID, Site, PersonID
Private Const msSTUDY_SITE_PERSON = "ClinicalTrialId|TrialSite|PersonID"

' The Response table
Private Const msRESPONSE_TABLE = "DataItemResponse"

' The Response history table
Private Const msRESPONSEHISTORY_TABLE = "DataItemResponseHistory"

' The columns we want
' NCJ 2 Nov 01 - Added RepeatNumber
' RS 22/9/2002 - Added |ResponseTimeStamp_TZ|DatabaseTimeStamp|DatabaseTimestamp_TZ
' NCJ 22 Jan 03 - Added StandardValue
' NCJ 28 Jan 03 - Added UserNameFull
Private Const msRESPONSE_COLS = "ResponseTaskId|CRFElementId|DataItemId|" _
                & "ResponseValue|ValueCode|" _
                & "ResponseTimeStamp|UserName|" _
                & "Comments|ResponseStatus|SoftwareVersion|" _
                & "ReasonForChange|LockStatus|ValidationMessage|OverruleReason|" _
                & "LabResult|CTCGrade|ClinicalTestDate|LaboratoryCode|Changed|" _
                & "HadValue|RepeatNumber|ResponseTimeStamp_TZ|DatabaseTimeStamp|DatabaseTimestamp_TZ|" _
                & "StandardValue|UserNameFull"
                
' NCJ 13 Aug 02 - Added ChangeCount for DataItemResponse but NOT DataItemResponseHistory
Private Const msNON_HISTORY_COLS = "ChangeCount"

'constants for cols than are used duirng the load but are never saved
Private Const msRESPONSE_LOADONLY_COLS = "DiscrepancyStatus|SDVStatus|NoteStatus"

' The "key" columns
Private Const msRESPONSE_KEY_COLS = "ClinicalTrialId|TrialSite|PersonId|CRFPageTaskId"

'cols conting the extra information required for an INSERT
Private Const msRESPONSE_EXTRA_COLS = "VisitId|VisitCycleNumber|CRFPageId|CRFPageCycleNumber"

' NCJ 5 Mar 02 - Order cols for Response table
Private Const msRESPONSE_ORDERCOLS = "ResponseTaskId|RepeatNumber"

'taken from MACRO enumeration for changed data
Private Const m_CHANGED_CHANGED = 1


' NCJ 17/9/01 - The columns used when saving Registration details
Private Const msREG_STATUS = "RegistrationStatus"
Private Const msREG_COLS = "RegistrationStatus|LocalIdentifier2|SubjectTimeStamp|SubjectTimeStamp_TZ|ProformaState"

Private Const msNORMAL_RANGE = "NormalRange"
Private Const msNR_KEY_COL = "LaboratoryCode"
Private Const msCTC = "CTC"
Private Const msCTC_KEY_COL = "CTCSchemeCode"
Private Const msCLINICAL_TEST_COL = "ClinicalTestCode"
Private Const msLABORATORY = "Laboratory"
Private Const msLAB_COLS = "LaboratoryCode|LaboratoryDescription"

'TA 5/4/02: SubjectCache table constants
Private Const msAREZZO_TABLE = "ArezzoToken"
Private Const msCOL_AREZZOID = "ArezzoID"
Private Const msCOL_DBTOKEN = "DBToken"

'--------------------------------------------------------------------
Public Function Load(ByVal sConnection As String, _
            ByVal lStudyId As Long, _
            ByVal sSite As String, _
            ByVal lPersonID As Long, ByVal bIncludeProformaState As Boolean, _
            ByVal bUpdateCacheTable As Boolean) As String
'--------------------------------------------------------------------
' Load the given study subject and return info as (property bag) string
' sConnection is the DB connection string.
' If BIncludeProformaState is false then an empty string is retuned for the proformastate
'--------------------------------------------------------------------
Dim oSubjState As StudySubjectState
Dim oQS As QueryServer
Dim oQD As QueryDef
Dim vSubject As Variant

    On Error GoTo ErrLabel
    
    ' Create the Query Server and Query Defn objects
    Set oQS = New QueryServer
    Set oQD = New QueryDef
    
    ' We'll SELECT from TrialSubject table, filtering on Trial, Site, Person
    oQD.InitSelect msSTUDY_SUBJECT, _
                    Split(msSUBJECT_COLS & msDELIMITER & msSUBJECT_LOADONLY_COLS, msDELIMITER), _
                    Split(msSTUDY_SITE_PERSON, msDELIMITER), _
                    Array(lStudyId, sSite, lPersonID)
    
    If bIncludeProformaState Then
        'include this col if wanted
        oQD.QueryFields.Add msSUBJECT_PROFORMACOL
    End If
     
    oQS.Init sConnection
    oQS.ConnectionOpen
    oQS.BeginTrans
    
    On Error GoTo ErrLabelTrans
    
    'TA 09/04/2002: Update chache table if requested
    If bUpdateCacheTable Then
        'notify subject cache that that this subject is out of date for other users
        Call UpdateCacheTable(oQS, lPersonID, sSite, lStudyId)
    End If
    
    ' Get the array of values from the database
    vSubject = oQS.SelectArray(oQD)
    
    ' Keep oQS open for loading VisitInstances
    Set oQD = Nothing
    
    Set oSubjState = New StudySubjectState

    Set oSubjState.VisitInstances = LoadVisitInstances(oQS, _
                                    lStudyId, sSite, lPersonID)
    Set oSubjState.eFormInstances = LoadEFormInstances(oQS, _
                                    lStudyId, sSite, lPersonID)
                                    
    'commit
    oQS.Commit
    On Error GoTo ErrLabel
    
    Set oQS = Nothing       ' This closes connection
        
    ' Fill in values from Row 0 of our array
    With oSubjState
        .StudyId = lStudyId
        .PersonID = lPersonID
        .Site = sSite
        .Label = ConvertFromNull(vSubject(0, 0), vbString)
        .RegID = ConvertFromNull(vSubject(1, 0), vbString)
        .DOB = ConvertFromNull(vSubject(2, 0), vbLong)
        .Status = vSubject(3, 0)
        .LockStatus = vSubject(4, 0)
        .Gender = ConvertFromNull(vSubject(5, 0), vbInteger)
        .RegStatus = ConvertFromNull(vSubject(6, 0), vbInteger)
        ' NCJ 7 Mar 03 - New time stamp values
        .SubjectTimeStamp = ConvertFromNull(vSubject(7, 0), DATATYPE_DATE)
        .SubjectTimeStamp_TZ = IIf(IsNull(vSubject(8, 0)), 0, vSubject(8, 0))
        ' MIMsg Status cols
        .DiscrepancyStatus = vSubject(9, 0)
        .SDVStatus = vSubject(10, 0)
        .NoteStatus = vSubject(11, 0)
        If bIncludeProformaState Then
            .ArezzoState = ConvertFromNull(vSubject(12, 0), vbString)
        Else
            .ArezzoState = ""
        End If
        
        
        Load = .State
    End With

    Set oSubjState = Nothing
    
    Exit Function
    
ErrLabel:
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.Load"
    Exit Function
    
ErrLabelTrans:

    oQS.Rollback
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.Load"
    Exit Function

End Function

'--------------------------------------------------------------------
Private Function LoadEFormInstances(ByRef oQS As QueryServer, _
            ByVal lStudyId As Long, _
            ByVal sSite As String, _
            ByVal lPersonID As Long) As Collection
'--------------------------------------------------------------------
' Load ALL the eForm instances for this study subject
' and return info as a collection of eFormInstanceStates
' oQS is the open Query Server
'--------------------------------------------------------------------
Dim oQD As QueryDef
Dim vEFI As Variant
Dim oEFIState As eFormInstanceState
Dim i As Long
Dim colEFIs As Collection

    On Error GoTo ErrLabel
    
    ' Create the Query Server and Query Defn objects
    Set oQD = New QueryDef

    ' Select on Study, Site, Person, VisitID, VicitCycleNo
    oQD.InitSelect msEFORMI_TABLE, _
                    Split(msEFORMI_COLS & msDELIMITER & msEFORMI_LOADONLY_COLS, msDELIMITER), _
                    Split(msSTUDY_SITE_PERSON, msDELIMITER), _
                    Array(lStudyId, sSite, lPersonID)
    
    ' Read info as array from DB, using oQS
    vEFI = oQS.SelectArray(oQD)
    
    Set oQD = Nothing
    
    ' Create new collection for collections of eForm Instances
    Set colEFIs = New Collection
    
    ' Check we got some data returned (NCJ 28 Jun 01)
    If Not IsNull(vEFI) Then
        ' Walk through the array
        For i = 0 To UBound(vEFI, 2)     ' Rows is second dimension
            Set oEFIState = New eFormInstanceState
            With oEFIState
                .eFormTaskId = vEFI(0, i)
                .eFormId = vEFI(1, i)
                .CycleNo = vEFI(2, i)
                .VisitId = vEFI(3, i)
                .VisitCycleNo = vEFI(4, i)
                .eFormLabel = ConvertFromNull(vEFI(5, i), vbString)
                .eFormDate = ConvertFromNull(vEFI(6, i), DATATYPE_DATE)
                .Status = ConvertFromNull(vEFI(7, i), vbInteger)
                .LockStatus = ConvertFromNull(vEFI(8, i), vbInteger)
                .LabCode = ConvertFromNull(vEFI(9, i), vbString)
                'new status cols
                .DiscrepancyStatus = vEFI(10, i)
                .SDVStatus = vEFI(11, i)
                .NoteStatus = vEFI(12, i)
                
                
                .IsChanged = False
                .IsNew = False
                ' Add to the collection
                colEFIs.Add oEFIState, oEFIState.Key
            End With
        Next
    End If
    
    ' Return the collection of eForm instances
    Set LoadEFormInstances = colEFIs
    
    Set colEFIs = Nothing
    Set oEFIState = Nothing
    
Exit Function

ErrLabel:
    
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.LoadEFormInstances"
    Exit Function

End Function

'--------------------------------------------------------------------
Private Function LoadVisitInstances(ByRef oQS As QueryServer, _
            ByVal lStudyId As Long, _
            ByVal sSite As String, _
            ByVal lPersonID As Long) As Collection
'--------------------------------------------------------------------
' Load the visit instances for this study subject
' and return info as a collection of VisitInstanceState objects
' oQS is the open Query Server
'--------------------------------------------------------------------
Dim oQD As QueryDef
Dim vVI As Variant
Dim oVIState As VisitInstanceState
Dim i As Long
Dim oVIs As Collection


    On Error GoTo ErrLabel
    
    ' Create the Query Server and Query Defn objects
    Set oQD = New QueryDef

    oQD.InitSelect msVI_TABLE, _
                    Split(msVI_COLS & msDELIMITER & msVI_LOADONLY_COLS, msDELIMITER), _
                    Split(msSTUDY_SITE_PERSON, msDELIMITER), _
                    Array(lStudyId, sSite, lPersonID)
    
    ' Read info as array from DB
    vVI = oQS.SelectArray(oQD)
    
    Set oQD = Nothing
    
    ' Create new collection for Visit Instances
    Set oVIs = New Collection

    ' Check we got some data returned
    If Not IsNull(vVI) Then
        ' Walk through the array
        For i = 0 To UBound(vVI, 2)     ' Rows is second dimension
            Set oVIState = New VisitInstanceState
            With oVIState
                .VisitTaskId = vVI(0, i)
                .VisitId = vVI(1, i)
                .CycleNo = vVI(2, i)
                .Status = vVI(3, i)
                .VisitDate = ConvertFromNull(vVI(4, i), DATATYPE_DATE)
                .LockStatus = ConvertFromNull(vVI(5, i), vbInteger)
                'new status cols
                .DiscrepancyStatus = vVI(6, i)
                .SDVStatus = vVI(7, i)
                .NoteStatus = vVI(8, i)
                
                .IsNew = False
                .IsChanged = False
                oVIs.Add oVIState, Str(.VisitTaskId)
            End With
        Next
    End If
    
    ' Return the collection of visit instances
    Set LoadVisitInstances = oVIs
    
    Set oVIs = Nothing
    Set oVIState = Nothing
    
Exit Function

ErrLabel:
    
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.LoadVisitInstances"
    Exit Function
    
End Function

'--------------------------------------------------------------------
Public Function Save(ByVal sConnection As String, _
                ByVal sState As String) As Long
'--------------------------------------------------------------------
' Save the given study subject represented as (property bag) string sState.
' sConnection is the DB connection string.
' Returns Person ID (newly generated if a new subject)
'--------------------------------------------------------------------
Dim oQS As QueryServer
Dim lSubjectId As Long

    On Error GoTo ErrLabel
    
    ' Create the Query Server and Query Defn objects
    Set oQS = New QueryServer
    ' Open the DB connection
    oQS.Init sConnection
    oQS.ConnectionOpen

    'start transaction
    oQS.BeginTrans
    On Error GoTo ErrLabelTrans
    
    lSubjectId = SaveSubject(oQS, sState)

    'end transaction
    oQS.Commit
    Set oQS = Nothing       ' This closes connection
    
    Save = lSubjectId
    
    Exit Function
    
ErrLabel:
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.Save"
    Exit Function
    
ErrLabelTrans:

    oQS.Rollback
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.Save"
    Exit Function

End Function

'--------------------------------------------------------------------
Private Function SaveSubject(oQS As QueryServer, sState As String) As Long
'--------------------------------------------------------------------
'SAves the subject and returns an id if a new subject.
' This is called by Save and SaveResponses.

'MLM 6/12/01: Flag subject record as "Changed" when saving.
'--------------------------------------------------------------------
Dim oSubjState As StudySubjectState
Dim oQD As QueryDef
Dim vValues As Variant
Dim stSaveType As QSSaveType

    ' Create the Query Defn object

    Set oQD = New QueryDef
    
    ' Set the subject state we've been passed
    Set oSubjState = New StudySubjectState
    oSubjState.State = sState

    With oSubjState
        ' Check for new subject
        If .IsNew Then
            'get person id from sequnce
'TA this way is not working            .PersonID = oQS.GetNextSequenceNo("SEQ_PERSONID")
'            ' New subject - create a New ID for them (within this Trial and Site)
            oQD.InitNewId msSTUDY_SUBJECT, msSUBJECT_IDCOL, _
                        Split(msSUBJECT_NEWIDFILTERCOLS, msDELIMITER), _
                        Array(.StudyId, .Site)
            .PersonID = oQS.QueryDefNewId(oQD)
            stSaveType = stInsert
        Else
            stSaveType = stUpdate
        End If
        
        If .IsNew Or .IsChanged Then
           ' Create the array of values to save
           'MLM 6/12/01: Include Changed flag
           ' NCJ 7 Mar 03 - Include timestamp values
           vValues = Array(.StudyId, .Site, .PersonID, .Label, .RegID, _
                           .DOB, .Status, .LockStatus, .Gender, .RegStatus, _
                           .SubjectTimeStamp, .SubjectTimeStamp_TZ, _
                           .ArezzoState, m_CHANGED_CHANGED)
           ' Save to the TrialSubject table
           oQD.InitSave msSTUDY_SUBJECT, _
                           Split(msSTUDY_SITE_PERSON & msDELIMITER & msSUBJECT_COLS & msDELIMITER & msSUBJECT_PROFORMACOL & msDELIMITER & msCHANGED_COL, msDELIMITER), _
                           vValues, _
                           Split(msSTUDY_SITE_PERSON, msDELIMITER), _
                           Array(.StudyId, .Site, .PersonID)
            
            If .IsNew Then
                oQD.QueryFields.Add Split(msSUBJECT_LOADONLY_COLS, msDELIMITER)(0), , .DiscrepancyStatus
                oQD.QueryFields.Add Split(msSUBJECT_LOADONLY_COLS, msDELIMITER)(1), , .SDVStatus
                oQD.QueryFields.Add Split(msSUBJECT_LOADONLY_COLS, msDELIMITER)(2), , .NoteStatus
            End If
            
           oQS.SelectSave stSaveType, oQD
        Else
            'nothing to save in trial subject
        End If
    
        ' Now save the Visit and eForm instances
        Call SaveVisitInstances(oQS, .VisitInstances, .StudyId, .Site, .PersonID)
        Call SaveEFormInstances(oQS, .eFormInstances, .StudyId, .Site, .PersonID)
        
    End With

    Set oQD = Nothing
    SaveSubject = oSubjState.PersonID
    Set oSubjState = Nothing
    
Exit Function

ErrLabel:
    
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.SaveSubject"
    Exit Function
    
End Function

'--------------------------------------------------------------------
Private Sub SaveVisitInstances(oQS As QueryServer, colVIs As Collection, _
            ByVal lStudyId As Long, _
            ByVal sSite As String, _
            ByVal lPersonID As Long)
'--------------------------------------------------------------------
' Save the new or changed visit instances
' colVIs is collection of VisitInstanceState objects
'--------------------------------------------------------------------
Dim oVIState As VisitInstanceState
Dim oQD As QueryDef
Dim vValues As Variant

    On Error GoTo ErrLabel
    
    Set oQD = New QueryDef
    
    For Each oVIState In colVIs
        With oVIState
            If .IsNew Then
                ' Create the array of values to save
                vValues = Array(lStudyId, sSite, lPersonID, _
                                .VisitTaskId, .VisitId, .CycleNo, _
                                .Status, .VisitDate, .LockStatus, 1, .DiscrepancyStatus, .SDVStatus, .NoteStatus)
                ' Save to the visit instance table
                oQD.InitSave msVI_TABLE, _
                                Split(msSTUDY_SITE_PERSON & msDELIMITER _
                                        & msVI_COLS & msDELIMITER & msCHANGED_COL & msDELIMITER & msVI_LOADONLY_COLS, msDELIMITER), _
                                vValues
                ' Do an insert
                oQS.QueryInsert oQD
            Else
                If .IsChanged Then
                    ' Create the array of values to save
                    vValues = Array(.Status, .VisitDate, 1)
                    ' Save to the visit instance table
                    oQD.InitSave msVI_TABLE, _
                                    Split(msVI_UPDATE_COLS, msDELIMITER), _
                                    vValues, _
                                    Split(msSTUDY_SITE_PERSON & msDELIMITER & msVI_TASKIDCOL, msDELIMITER), _
                                    Array(lStudyId, sSite, lPersonID, .VisitTaskId)
                    ' Do an update
                    oQS.QueryUpdate oQD
                Else
                    ' Do nothing if not changed or new
                End If
            End If
        End With
    Next
    
    Set oVIState = Nothing
    
Exit Sub

ErrLabel:
    
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.SaveVisitInstances"
    Exit Sub
    
End Sub

'--------------------------------------------------------------------
Private Sub SaveEFormInstances(oQS As QueryServer, colEFIs As Collection, _
            ByVal lStudyId As Long, _
            ByVal sSite As String, _
            ByVal lPersonID As Long)
'--------------------------------------------------------------------
' Save the new or changed eForm instances
' colEFIs is collection of eFormInstanceState objects
'--------------------------------------------------------------------
Dim oEFIState As eFormInstanceState
Dim oQD As QueryDef
Dim vValues As Variant

    On Error GoTo ErrLabel
    
    Set oQD = New QueryDef

    For Each oEFIState In colEFIs
        With oEFIState
            If .IsNew Then
                ' Create the array of values to save
                vValues = Array(lStudyId, sSite, lPersonID, _
                                .eFormTaskId, .eFormId, .CycleNo, _
                                .VisitId, .VisitCycleNo, _
                                .eFormLabel, .eFormDate, .Status, .LockStatus, .LabCode, m_CHANGED_CHANGED, .DiscrepancyStatus, .SDVStatus, .NoteStatus)
                ' Save to the eForm instance table
                oQD.InitSave msEFORMI_TABLE, _
                                Split(msSTUDY_SITE_PERSON & msDELIMITER _
                                        & msEFORMI_COLS & msDELIMITER & msCHANGED_COL & msDELIMITER & msEFORMI_LOADONLY_COLS, msDELIMITER), _
                                vValues
                ' Do an insert
                oQS.QueryInsert oQD
            Else
                If .IsChanged Then
                    ' Create the array of (changed) values to save
                    vValues = Array(.eFormLabel, .eFormDate, .Status, .LabCode, 1)
                    ' Save to the eForm instance table
                    oQD.InitSave msEFORMI_TABLE, _
                                    Split(msEFORMI_UPDATE_COLS, msDELIMITER), _
                                    vValues, _
                                    Split(msSTUDY_SITE_PERSON & msDELIMITER & msEFORMI_TASKID_COL, msDELIMITER), _
                                    Array(lStudyId, sSite, lPersonID, .eFormTaskId)
                    ' Do an update
                    oQS.QueryUpdate oQD
                Else
                    ' If it's not new or changed then we do nothing
                End If
            End If
        End With
    Next
    
    Set oEFIState = Nothing
    
Exit Sub

ErrLabel:
    
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.SaveEFormInstances"
    Exit Sub
    
End Sub


'----------------------------------------------------
Public Function LoadResponses(ByVal sConnection As String, _
            ByVal lStudyId As Long, _
            ByVal sSite As String, _
            ByVal lPersonID As Long, _
            ByVal lEFormTaskId As Long) As String
'----------------------------------------------------
' Load all responses for this subject
' and return as string, ordered by ResponseTaskId and RepeatNumber
' (representing serialised collection of ResponseState objects)
'----------------------------------------------------
' NCJ 2 Nov 01 - Include RepeatNumber for MACRO 3.0
' NCJ 5 Mar 02 - Order by ReponseTaskId and Repeat Number (necessary for lab test handling)
' NCJ 12 Aug 02 - Added ChangeCount
'----------------------------------------------------
Dim oQS As QueryServer
Dim oQD As QueryDef
Dim oResponseState As ResponseState
Dim colResponses As Collection
Dim vResponseData As Variant
Dim i As Long

    On Error GoTo ErrLabel
    
    Set oQS = New QueryServer
    ' Create the Query Defn object
    Set oQD = New QueryDef

    ' Select on Study, Site, Person, eFormTaskId
    ' NCJ 5 Mar 02 - Order on responseTaskId and Repeat Number
    ' NCJ 13 Aug 02 - Added msNON_HISTORY_COLS
    oQD.InitSelect msRESPONSE_TABLE, _
                    Split(msRESPONSE_COLS & msDELIMITER & msNON_HISTORY_COLS & msDELIMITER & msRESPONSE_LOADONLY_COLS, msDELIMITER), _
                    Split(msRESPONSE_KEY_COLS, msDELIMITER), _
                    Array(lStudyId, sSite, lPersonID, lEFormTaskId), _
                    Split(msRESPONSE_ORDERCOLS, msDELIMITER)
   
    oQS.ConnectionOpen sConnection
    oQS.BeginTrans
    On Error GoTo ErrLabelTrans
    
    ' Read info as array from DB
    vResponseData = oQS.SelectArray(oQD)
    
    oQS.Commit
    On Error GoTo ErrLabel
    
    Set oQS = Nothing
    Set oQD = Nothing
    
    Set colResponses = New Collection
    
    ' Check we got some data returned
    If Not IsNull(vResponseData) Then
    
        ' Walk through the array
        For i = 0 To UBound(vResponseData, 2)     ' Rows is second dimension
            Set oResponseState = New ResponseState
            With oResponseState
                .ResponseId = vResponseData(0, i)
                .ElementId = vResponseData(1, i)
                .DataItemId = vResponseData(2, i)
                .Value = ConvertFromNull(vResponseData(3, i), vbString)
                .ValueCode = ConvertFromNull(vResponseData(4, i), vbString)
                .TimeStamp = vResponseData(5, i)
                .UserCode = ConvertFromNull(vResponseData(6, i), vbString)
                .Comments = ConvertFromNull(vResponseData(7, i), vbString)
                .Status = vResponseData(8, i)
                .SoftwareVersion = ConvertFromNull(vResponseData(9, i), vbString)
                .ReasonForChange = ConvertFromNull(vResponseData(10, i), vbString)
                .LockStatus = vResponseData(11, i)
                .ValidationMessage = ConvertFromNull(vResponseData(12, i), vbString)
                .OverruleReason = ConvertFromNull(vResponseData(13, i), vbString)
                ' NRStatus is an integer even though it's stored as text in the DB!
                .NRStatus = ConvertFromNull(vResponseData(14, i), vbInteger)
                .CTCGrade = ConvertFromNull(vResponseData(15, i), vbInteger)
                .LabTestDate = ConvertFromNull(vResponseData(16, i), vbDouble)
                .LabCode = ConvertFromNull(vResponseData(17, i), vbString)
                ' vResponseData(18, i) is the Changed field but we set it to False
                .IsChanged = False
                .HadValue = (vResponseData(19, i) = 1)
                ' NCJ 2 Nov 01 - Added RepeatNumber
                .RepeatNumber = ConvertFromNull(vResponseData(20, i), vbInteger)
                ' Deal with old data with possibly no repeat number
                If .RepeatNumber < 1 Then
                    .RepeatNumber = 1
                End If
                
                ' RS 22/9/2002
                .TimeSTampTZ = IIf(IsNull(vResponseData(21, i)), 0, vResponseData(21, i))
                
                ' Ignore DatabaseTimeStamp & Zone for load:  22, 23
                
                ' NCJ 22 Jan 03
                .StandardValue = ConvertFromNull(vResponseData(24, i), DATATYPE_DATE)
                ' NCJ 28 Jan 03 - Default to UserCode if no UserNameFull
                .UserNameFull = IIf(IsNull(vResponseData(25, i)), .UserCode, vResponseData(25, i))
                
                ' NCJ 13 Aug 02 - Added ChangeCount
                .ChangeCount = vResponseData(26, i)
                
                'new status cols
                .DiscrepancyStatus = vResponseData(27, i)
                .SDVStatus = vResponseData(28, i)
                .NoteStatus = vResponseData(29, i)
                
                
                .IsNew = False
                ' Add to the collection
                colResponses.Add oResponseState, oResponseState.Key
            End With
        Next
    
    End If

    ' Serialise the collection of states
    LoadResponses = CollectionSerialise(colResponses)
    
    Set colResponses = Nothing
    Set oResponseState = Nothing
    
    Exit Function
    
ErrLabel:
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.LoadResponses"
    Exit Function
    
ErrLabelTrans:

    oQS.Rollback
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.LoadResponses"
    Exit Function

End Function

'--------------------------------------------------------------------
Public Function SaveResponses(ByVal sConnection As String, _
                ByVal sResponseStates As String, ByVal sSubjectState As String, _
                ByVal lStudyId As Long, ByVal sSite As String, ByVal lPersonID As Long, _
                ByVal lEFormTaskId As Long, _
                ByVal lVisitId As Long, _
                ByVal nVisitCycleNo As Integer, ByVal lEFormId As Long, _
                ByVal nEFormCycleNo As Integer) As Long
'--------------------------------------------------------------------
' Save the new or changed responses - save subject state as well if this is passed in
' NCJ 2 Nov 01 - Include RepeatNumber for MACRO 3.0
' NCJ 22 Nov 01 - Include GroupInstanceStates too
' NCJ 13 Aug 02 - Include ChangeCount
' RS 25/10/2002 - DatabaseTimestamp value passed as 0 (in vValue array), so that trigger will put correct database timestamp
' NCJ 14 Nov 02 - Removed QGroupsState argument
' NCJ 29 Apr 03 - If ResponseState.IgnoreMe = TRUE, do not save the response
'--------------------------------------------------------------------
Dim oQS As QueryServer
Dim oQD As QueryDef
Dim oResponseState As ResponseState
Dim colResponses As Collection
Dim vValues As Variant
'Dim colGroupStates As Collection

    On Error GoTo ErrLabel
    Set oQD = New QueryDef
    
    Set oQS = New QueryServer
    oQS.ConnectionOpen sConnection
    oQS.BeginTrans
    On Error GoTo ErrLabelTrans
    
    ' Deal with the Responses first
    Set colResponses = CollectionDeSerialise(sResponseStates)
    
    For Each oResponseState In colResponses
        With oResponseState
            ' NCJ 29 Apr 03 - Make sure we're not ignoring it
            If Not .IgnoreMe Then
                'DataItemResponse table
                ' NCJ 2 Nov 01 - Added RepeatNumber
                ' NCJ 13 Aug 02 - Added ChangeCount
                ' RS 22/9/2002 - Added  .TimeSTampTZ, CDbl(Now), 0  (for the three new Timestamp&TimeZone columns
                If .IsNew Then
                    ' NCJ 22 Jan 03 - Added StandardValue
                    ' NCJ 28 Jan 03 - Added UserNameFull
                     vValues = Array(lVisitId, nVisitCycleNo, lEFormId, nEFormCycleNo, lStudyId, sSite, lPersonID, lEFormTaskId, .ResponseId, .ElementId, .DataItemId, ConvertToNull(.Value), _
                        ConvertToNull(.ValueCode), ConvertToNull(.TimeStamp), ConvertToNull(.UserCode), ConvertToNull(.Comments), _
                        ConvertToNull(.Status), ConvertToNull(.SoftwareVersion), ConvertToNull(.ReasonForChange), ConvertToNull(.LockStatus), _
                        ConvertToNull(.ValidationMessage), ConvertToNull(.OverruleReason), ConvertToNull(.NRStatus), _
                        ConvertToNull(.CTCGrade), ConvertToNull(.LabTestDate), ConvertToNull(.LabCode), _
                        m_CHANGED_CHANGED, ConvertHadValue(.HadValue), .RepeatNumber, _
                        .TimeSTampTZ, 0, 0, .StandardValue, .UserNameFull, .ChangeCount, .DiscrepancyStatus, .SDVStatus, .NoteStatus)
                        
                    ' NCJ 13 Aug 02 - Added non-history cols
                    oQD.InitSave msRESPONSE_TABLE, _
                        Split(msRESPONSE_EXTRA_COLS & msDELIMITER _
                            & msRESPONSE_KEY_COLS & msDELIMITER _
                            & msRESPONSE_COLS & msDELIMITER _
                            & msNON_HISTORY_COLS & msDELIMITER & msRESPONSE_LOADONLY_COLS, msDELIMITER), vValues
                    
                    ' Do an insert
                    oQS.QueryInsert oQD
                Else
                    If .IsChanged Then
                        ' NCJ 22 Jan 03 - Added StandardValue
                        ' NCJ 28 Jan 03 - Added UserNameFull
                        vValues = Array(.ResponseId, .ElementId, .DataItemId, ConvertToNull(.Value), _
                            ConvertToNull(.ValueCode), ConvertToNull(.TimeStamp), ConvertToNull(.UserCode), ConvertToNull(.Comments), _
                            ConvertToNull(.Status), ConvertToNull(.SoftwareVersion), ConvertToNull(.ReasonForChange), ConvertToNull(.LockStatus), _
                            ConvertToNull(.ValidationMessage), ConvertToNull(.OverruleReason), ConvertToNull(.NRStatus), _
                            ConvertToNull(.CTCGrade), ConvertToNull(.LabTestDate), ConvertToNull(.LabCode), _
                            m_CHANGED_CHANGED, ConvertHadValue(.HadValue), .RepeatNumber, _
                            .TimeSTampTZ, 0, 0, .StandardValue, .UserNameFull, .ChangeCount)
                            
                          '(response task id is part of key)
                          ' NCJ 2 Nov 01 - RepeatNumber is now part of the key too
                          ' NCJ 13 Aug 02 - Added non-history cols
                        oQD.InitSave msRESPONSE_TABLE, _
                            Split(msRESPONSE_COLS & msDELIMITER & msNON_HISTORY_COLS, msDELIMITER), _
                            vValues, _
                            Split(msRESPONSE_KEY_COLS _
                                    & msDELIMITER & "ResponseTaskId" _
                                    & msDELIMITER & "RepeatNumber", msDELIMITER), _
                            Array(lStudyId, sSite, lPersonID, lEFormTaskId, .ResponseId, .RepeatNumber)
    
                            ' Do an update
                            oQS.QueryUpdate oQD
                    Else
                        ' Do nothing if not changed or new
                    End If
                End If
                
                'DataItemResponseHistory
                ' NCJ 2 Nov 01 - Added RepeatNumber
                ' RS 22/9/2002 - Added  .TimeSTampTZ, CDbl(Now), 0  (for the three new Timestamp&TimeZone columns
                ' NCJ 22 Jan 03 - Added StandardValue
                ' NCJ 28 Jan 03 - Added UserNameFull
                If .IsChanged Or .IsNew Then
                     vValues = Array(lVisitId, nVisitCycleNo, lEFormId, nEFormCycleNo, lStudyId, sSite, lPersonID, lEFormTaskId, .ResponseId, .ElementId, .DataItemId, ConvertToNull(.Value), _
                        ConvertToNull(.ValueCode), ConvertToNull(.TimeStamp), ConvertToNull(.UserCode), ConvertToNull(.Comments), _
                        ConvertToNull(.Status), ConvertToNull(.SoftwareVersion), ConvertToNull(.ReasonForChange), ConvertToNull(.LockStatus), _
                        ConvertToNull(.ValidationMessage), ConvertToNull(.OverruleReason), ConvertToNull(.NRStatus), _
                        ConvertToNull(.CTCGrade), ConvertToNull(.LabTestDate), ConvertToNull(.LabCode), _
                        m_CHANGED_CHANGED, ConvertHadValue(.HadValue), .RepeatNumber, _
                        .TimeSTampTZ, 0, 0, .StandardValue, .UserNameFull)
                        
                    oQD.InitSave msRESPONSEHISTORY_TABLE, _
                        Split(msRESPONSE_EXTRA_COLS & msDELIMITER & msRESPONSE_KEY_COLS & msDELIMITER & msRESPONSE_COLS, msDELIMITER), vValues
        
                    ' Do an insert
                    oQS.QueryInsert oQD
                End If
            End If
        End With
    Next
    
    'save the subject state
    If sSubjectState <> "" Then
        Call SaveSubject(oQS, sSubjectState)
    End If
    
    oQS.Commit
    On Error GoTo ErrLabel
    Set oQS = Nothing
    Set oQD = Nothing
    
    Set oResponseState = Nothing
    Set colResponses = Nothing
'    Set colGroupStates = Nothing
    
Exit Function
    
ErrLabel:
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.SaveResponses"
    Exit Function
    
ErrLabelTrans:

    oQS.Rollback
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.SaveResponses"
    Exit Function
    
End Function

'--------------------------------------------------------------------
Public Sub SaveRegistration(ByVal sConnection As String, _
                ByVal sState As String)
'--------------------------------------------------------------------
' Save Registration details ONLY for subject.
' This includes registration Status and ID.
' sState is serialised (property bag) Subject state string.
'--------------------------------------------------------------------
Dim oSubjState As StudySubjectState
Dim oQS As QueryServer
Dim oQD As QueryDef
Dim vValues As Variant

    On Error GoTo ErrLabel
    
    ' Create the Query Server and Query Defn objects
    Set oQS = New QueryServer
    Set oQD = New QueryDef
    
    ' Open the DB connection
    oQS.Init sConnection
    oQS.ConnectionOpen

    'start transaction
    oQS.BeginTrans
    On Error GoTo ErrLabelTrans
    
    ' Set the subject state we've been passed
    Set oSubjState = New StudySubjectState
    oSubjState.State = sState

    With oSubjState
            ' NCJ 7 Mar 03 - Added times stamp fields before ArezzoState
            vValues = Array(.RegStatus, .RegID, .SubjectTimeStamp, .SubjectTimeStamp_TZ, .ArezzoState)
            ' Save to the TrialSubject table
            oQD.InitSave msSTUDY_SUBJECT, _
                            Split(msREG_COLS, msDELIMITER), _
                            vValues, _
                            Split(msSTUDY_SITE_PERSON, msDELIMITER), _
                            Array(.StudyId, .Site, .PersonID)
            ' Use SelectSave because we're saving ProformaState
            oQS.SelectSave stUpdate, oQD
    End With

    ' End transaction
    oQS.Commit
    On Error GoTo ErrLabel
    
    Set oQS = Nothing       ' This closes connection
    Set oQD = Nothing
    Set oSubjState = Nothing
    
    Exit Sub
    
ErrLabel:
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.SaveRegistration"
    Exit Sub
    
ErrLabelTrans:

    oQS.Rollback
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.SaveRegistration"
    Exit Sub
    

End Sub

'--------------------------------------------------------------------
Public Sub SaveRegistrationStatus(ByVal sConnection As String, _
                ByVal nStatus As Integer, _
                ByVal lStudyId As Long, _
                ByVal sSite As String, _
                ByVal lPersonID As Long)
'--------------------------------------------------------------------
' Save Registration Status only for this subject.
'--------------------------------------------------------------------
Dim oQS As QueryServer
Dim oQD As QueryDef

    On Error GoTo ErrLabel
    ' Create the Query Server and Query Defn objects
    Set oQS = New QueryServer
    Set oQD = New QueryDef
    
    ' Open the DB connection
    oQS.Init sConnection
    oQS.ConnectionOpen

    ' Start transaction
    oQS.BeginTrans
    On Error GoTo ErrLabelTrans
    
    ' Save RegStatus to the TrialSubject table
    ' using given Study/Site/PersonID
    oQD.InitSave msSTUDY_SUBJECT, _
                    msREG_STATUS, _
                    nStatus, _
                    Split(msSTUDY_SITE_PERSON, msDELIMITER), _
                    Array(lStudyId, sSite, lPersonID)
    ' Can use QueryUpdate because we're not saving ProformaState
    oQS.QueryUpdate oQD
    
    ' End transaction
    oQS.Commit
    On Error GoTo ErrLabel
    Set oQS = Nothing       ' This closes connection
    Set oQD = Nothing
    
    Exit Sub
    
ErrLabel:
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.SaveRegistrationStatus"
    Exit Sub
    
ErrLabelTrans:

    oQS.Rollback
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.SaveRegistrationStatus"
    Exit Sub
    

End Sub

'----------------------------------------------------------------------------------------'
Public Function LoadLabDescription(ByVal sConnection As String, ByVal sLab As String) As String
'----------------------------------------------------------------------------------------'
' Load the Lab description for the given Lab.
' If return string is empty, it means the Lab definition wasn't found.
'----------------------------------------------------------------------------------------'
Dim vLab As Variant
Dim oQS As QueryServer
Dim oQD As QueryDef

    On Error GoTo ErrLabel
    
    ' Create the Query Server and Query Defn objects
    Set oQD = New QueryDef
    Set oQS = New QueryServer
    
    oQS.Init sConnection
    oQS.ConnectionOpen
    
    oQD.InitSelect msLABORATORY, Split(msLAB_COLS, msDELIMITER), msNR_KEY_COL, sLab
    vLab = oQS.SelectArray(oQD)
    oQS.ConnectionClose
    
    Set oQS = Nothing
    Set oQD = Nothing
  
    If IsNull(vLab) Then
        ' Lab not found
        ' Return an empty Lab description
        LoadLabDescription = ""
    Else
        ' We found the lab
        LoadLabDescription = vLab(1, 0)
    End If
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.LoadLabDescription"

End Function

'----------------------------------------------------------------------------------------'
Public Function LoadNRandCTC(ByVal sConnection As String, ByVal sClinicalTests As String, _
                        ByVal sLab As String, ByVal sScheme As String) As String
'----------------------------------------------------------------------------------------'
' Load NRs and CTCs from db
' according to Lab Tests (separated by commas), Lab and CTC Scheme
' Returns serialised collection of LabTestValidateStates
' Assume Lab exists
'----------------------------------------------------------------------------------------'
Dim oQS As QueryServer
Dim oQD As QueryDef
Dim vNRs As Variant
Dim vCTC As Variant
Dim oNRState As NormalRangeState
Dim oCTCState As CTCriterionState
Dim i As Long
Dim colValidateStates As Collection
Dim oValidateState As LabTestValidateState
Dim vTests As Variant

    On Error GoTo ErrLabel
    
    vTests = Split(sClinicalTests, ",")
    
    Set colValidateStates = New Collection
    
    ' Create the Query Server and Query Defn objects
    Set oQD = New QueryDef

    Set oQS = New QueryServer
    oQS.Init sConnection
    oQS.ConnectionOpen
    oQS.BeginTrans
    On Error GoTo ErrLabelTrans
    
    ' See if there are any tests
    oQD.InitSelect msNORMAL_RANGE, , msNR_KEY_COL, sLab
    oQD.QueryFilters.Add msCLINICAL_TEST_COL, "IN", vTests
    'oQD.QueryOrders.Add msCLINICAL_TEST_COL
    
    ' Get the array of Normal Range values from the database
    ' Read info as array from DB, using oQS
    vNRs = oQS.SelectArray(oQD)

    oQD.InitSelect msCTC, , msCTC_KEY_COL, sScheme
    oQD.QueryFilters.Add msCLINICAL_TEST_COL, "IN", vTests
    'oQD.QueryOrders.Add msCLINICAL_TEST_COL
    ' Get the array of CTC values from the database
    vCTC = oQS.SelectArray(oQD)
    
    oQS.Commit
    On Error GoTo ErrLabel
    Set oQS = Nothing
    Set oQD = Nothing
    
    ' Create a LabTestValidateState for each clinical test
    For i = 0 To UBound(vTests)
        Set oValidateState = New LabTestValidateState
        oValidateState.ClinicalTest = vTests(i)
        oValidateState.Lab = sLab
        oValidateState.CTCScheme = sScheme
        ' Add to collection, indexed by ClinicalTestCode
        colValidateStates.Add oValidateState, vTests(i)
    Next
    
    ' Check we got some Normal Range data returned
    If Not IsNull(vNRs) Then
        ' Walk through the array
        For i = 0 To UBound(vNRs, 2)     ' Rows is second dimension
            Set oNRState = New NormalRangeState
            With oNRState
                .Id = vNRs(0, i)
                .LabCode = vNRs(1, i)
                .ClinicalTestCode = vNRs(2, i)
                .GenderCode = vNRs(3, i)
                .AgeMin = vNRs(4, i)
                .AgeMax = vNRs(5, i)
                .EffectiveStart = vNRs(6, i)
                .EffectiveEnd = vNRs(7, i)
                .NormalMin = vNRs(8, i)
                .NormalMax = vNRs(9, i)
                .FeasibleMin = vNRs(10, i)
                .FeasibleMax = vNRs(11, i)
                .AbsoluteMin = vNRs(12, i)
                .Absolutemax = vNRs(13, i)
                .PercentCode = vNRs(14, i)
                ' Add NRState to the collection
                colValidateStates(.ClinicalTestCode).NormalRangeStates.Add oNRState
            End With
        Next
    End If
    
    Set oNRState = Nothing
    
    ' Check we got some CTC data returned
    If Not IsNull(vCTC) Then
        ' Walk through the array
        For i = 0 To UBound(vCTC, 2)     ' Rows is second dimension
            Set oCTCState = New CTCriterionState
            With oCTCState
                .Id = vCTC(0, i)
                .SchemeCode = vCTC(1, i)
                .ClinicalTestCode = vCTC(2, i)
                .Grade = vCTC(3, i)
                .Min = vCTC(4, i)
                .Max = vCTC(5, i)
                .MinType = vCTC(6, i)
                .MaxType = vCTC(7, i)
                ' Add CTCState to the collection
                colValidateStates(.ClinicalTestCode).CTCStates.Add oCTCState
            End With
        Next
    End If
    
    Set oCTCState = Nothing

    ' Return the serialised LabTestValidateState collection
     LoadNRandCTC = CollectionSerialise(colValidateStates)
    
Exit Function
    
ErrLabelTrans:

    oQS.Rollback

ErrLabel:
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.LoadNRandCTC"
    Exit Function
    
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub UpdateCacheTable(oQS As QueryServer, lSubjectId As Long, sSite As String, lStudyId As Long)
'----------------------------------------------------------------------------------------'
' Notify subject cache that that this subject is out of date for other users
'----------------------------------------------------------------------------------------'

Dim sAREZZOId As String
Dim oQD As QueryDef

    On Error GoTo ErrLabel
    
    sAREZZOId = lSubjectId & "," & sSite & "," & lStudyId
    Set oQD = New QueryDef
    oQD.InitSave msAREZZO_TABLE, msCOL_DBTOKEN, 0, msCOL_AREZZOID, sAREZZOId
    oQS.QueryUpdate oQD
    Set oQD = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubjectDS.UpdateCacheTable"
    
    
End Sub
 
'----------------------------------------------------------------------------------------'
Private Function ConvertHadValue(bHadValue As Boolean) As Integer
'----------------------------------------------------------------------------------------'
' Convert HadValue boolean to the database column type
'----------------------------------------------------------------------------------------'

    If bHadValue Then
        ConvertHadValue = 1
    Else
        ConvertHadValue = 0
    End If
    
 End Function



