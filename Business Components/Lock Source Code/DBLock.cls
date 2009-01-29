VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBLock"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   File:       MACRODBLock.cls
'   Copyright:  InferMed Ltd. 2001-2007. All Rights Reserved
'   Author:     Toby Aldridge, June 2001
'   Purpose:    High level db object locking class.
'----------------------------------------------------------------------------------------'
' Revisions:
' MLM 10/09/01: Added function AllLockDetails, to be used by admin tool for removing expired locks/
'               Added public enumeration of columns in the array returned from this function.
' MLM 21/05/02: Modified AllLockDetails to cope with new lock type for creation of new studies.
' TA 12/09/2002: Added support for locking the subject for saving
' NCJ 24 May 06 - Added stuff for MultiUser SD (eForm locks)
' NCJ 1-6 Jun 06 - Added more MUSD stuff
' NCJ 27 Jun 06 - Added OpenStudyLO
' NCJ 25 Oct 06 - Handle eForm Locks in SD in AllLockDetails
' NCJ 30 Nov 06 - Changed AllLockDetails to use explicit SQL to get round buglet in Query Services
' NCJ 26 Feb 07 - Bug 2879 - Added user filtering back in to AllLockDetails!
'----------------------------------------------------------------------------------------

Option Explicit

'enumeration of what is locked
' NCJ 31 May 06 - Added dblEForm
Public Enum DBLocked
    dblStudy
    dblSubject
    dblEFormInstance
    dblEForm
End Enum

' NCJ 25 Oct 06 - Added ldcEFormSDTitle to handle eForm locks in MUSD
Public Enum LockDetailColumn
    ldcStudyName = 0
    ldcSite = 1
    ldcSubjectId = 2
    ldcEFormTitle = 3
    ldcEFormCycleNumber = 4
    ldcUser = 5
    ldcLockTimeStamp = 6
    ldcStudyId = 7
    ldcEFormInstanceId = 8
    ldcToken = 9
    ldcEFormSDTitle = 10
End Enum

'TA 12/09/2002: mock efi id to signify locking the subject for saving
Private Const m_EFI_LOCK_SUBJECT_FOR_SAVE = -1

'----------------------------------------------------------------------
Public Function LockStudyRW(ByVal sCon As String, _
                            ByVal sUser As String, _
                            ByVal lStudyId As Long, _
                            Optional ByVal sngWait As Single = 0) As String
'----------------------------------------------------------------------
' NCJ 30 May 06 - Lock a study for editing in SD
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    LockStudyRW = oLockDBDS.LockStudyRW(sCon, sUser, lStudyId)
    Set oLockDBDS = Nothing
    
End Function

'----------------------------------------------------------------------
Public Function LockStudy(ByVal sCon As String, _
                            ByVal sUser As String, _
                            ByVal lStudyId As Long, _
                            Optional ByVal sngWait As Single = 0) As String
'----------------------------------------------------------------------
' Lock a study (a "global" lock)
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    LockStudy = oLockDBDS.LockStudy(sCon, sUser, lStudyId)
    Set oLockDBDS = Nothing
    
End Function

'----------------------------------------------------------------------
Public Function LockEForm(ByVal sCon As String, _
                                ByVal sUser As String, _
                                ByVal lStudyId As Long, _
                                ByVal lEFormID As Long) As String
'----------------------------------------------------------------------
' NCJ 1 Jun 06 - Lock an eForm for editing in SD.
' Returns token for unlocking if successful.
' Returns empty string if unsuccesful.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    LockEForm = oLockDBDS.LockEForm(sCon, sUser, lStudyId, lEFormID)
    Set oLockDBDS = Nothing

End Function

'----------------------------------------------------------------------
Public Function LockedEForms(ByVal sCon As String, _
                                ByVal sUser As String, _
                                ByVal lStudyId As Long) As Collection
'----------------------------------------------------------------------
' NCJ 6 Jun 06 - Return collection of eFormIDs currently locked by other users in this study
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    Set LockedEForms = oLockDBDS.LockedEForms(sCon, sUser, lStudyId)
    Set oLockDBDS = Nothing

End Function

'----------------------------------------------------------------------
Public Function LockSubject(ByVal sCon As String, _
                                ByVal sUser As String, _
                                ByVal lStudyId As Long, _
                                ByVal sSite As String, _
                                ByVal lSubjectId As Long, _
                                Optional ByVal sngWait As Single = 0) As String
'----------------------------------------------------------------------
' Lock a subject.
' Returns token for unlocking if successful.
' Returns empty string if unsuccesful.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    LockSubject = oLockDBDS.LockSubject(sCon, sUser, lStudyId, sSite, lSubjectId)
    Set oLockDBDS = Nothing

End Function

'----------------------------------------------------------------------
Public Function LockEFormInstance(ByVal sCon As String, _
                                    ByVal sUser As String, _
                                    ByVal lStudyId As Long, _
                                    ByVal sSite As String, _
                                    ByVal lSubjectId As Long, _
                                    ByVal lEFormInstanceId As Long, _
                                    Optional ByVal sngWait As Single = 0) As String
'----------------------------------------------------------------------
' Lock an eForm instance.
' Returns token for unlocking if successful.
' Returns empty string if unsuccesful.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    LockEFormInstance = oLockDBDS.LockEFormInstance(sCon, sUser, lStudyId, sSite, lSubjectId, lEFormInstanceId)
    Set oLockDBDS = Nothing
    
End Function

'----------------------------------------------------------------------
Public Function LockSubjectForSaving(ByVal sCon As String, _
                                    ByVal sUser As String, _
                                    ByVal lStudyId As Long, _
                                    ByVal sSite As String, _
                                    ByVal lSubjectId As Long) As String
'----------------------------------------------------------------------
' Lock an subject for saving
' Returns token for unlocking if successful.
' Returns empty string if unsuccesful.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    LockSubjectForSaving = oLockDBDS.LockSubjectForSaving(sCon, sUser, lStudyId, sSite, lSubjectId)
    Set oLockDBDS = Nothing
    
End Function

'----------------------------------------------------------------------
Public Function OpenStudyLO(ByVal sCon As String, ByVal sUser As String, _
                                ByVal lStudyId As Long) As String
'----------------------------------------------------------------------
' NCJ 27 Jun 06 - Open study LO
' No locks placed, but returns empty string if no FC or subject locks,
' otherwise returns type of lock that prevents LO opening
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    OpenStudyLO = oLockDBDS.OpenStudyLO(sCon, sUser, lStudyId)
    Set oLockDBDS = Nothing

End Function

'----------------------------------------------------------------------
Public Sub UnlockSubjectForSaving(ByVal sCon As String, _
                                        ByVal sToken As String, _
                                        ByVal lStudyId As Long, _
                                        ByVal sSite As String, _
                                        ByVal lSubjectId As Long)
'----------------------------------------------------------------------
' Unlock an subject for saving
' Errors if the token is invalid.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS
    
    Set oLockDBDS = New LockDBDS
    Call oLockDBDS.UnlockEFormInstance(sCon, sToken, lStudyId, sSite, lSubjectId, m_EFI_LOCK_SUBJECT_FOR_SAVE)
    Set oLockDBDS = Nothing
    
End Sub

'----------------------------------------------------------------------
Public Sub UnlockStudy(ByVal sCon As String, _
                        ByVal sToken As String, _
                        ByVal lStudyId As Long)
'----------------------------------------------------------------------
' Unlock a study ("global" lock)
' Errors if the token is invalid.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    Call oLockDBDS.UnlockStudy(sCon, sToken, lStudyId)
    Set oLockDBDS = Nothing

End Sub

'----------------------------------------------------------------------
Public Sub UnlockStudyRW(ByVal sCon As String, _
                        ByVal sToken As String, _
                        ByVal lStudyId As Long)
'----------------------------------------------------------------------
' NCJ 1 Jun 06 - Unlock a study as RW.
' Errors if the token is invalid.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    Call oLockDBDS.UnlockStudyRW(sCon, sToken, lStudyId)
    Set oLockDBDS = Nothing

End Sub
    
'----------------------------------------------------------------------
Public Sub UnlockSubject(ByVal sCon As String, _
                                ByVal sToken As String, _
                                ByVal lStudyId As Long, _
                                ByVal sSite As String, _
                                ByVal lSubjectId As Long)
'----------------------------------------------------------------------
' Unlock a subject.
' Errors if the token is invalid.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    Call oLockDBDS.UnlockSubject(sCon, sToken, lStudyId, sSite, lSubjectId)
    Set oLockDBDS = Nothing
        
End Sub

'----------------------------------------------------------------------
Public Sub UnlockEFormInstance(ByVal sCon As String, _
                                        ByVal sToken As String, _
                                        ByVal lStudyId As Long, _
                                        ByVal sSite As String, _
                                        ByVal lSubjectId As Long, _
                                        ByVal lEFormInstanceId As Long)
'----------------------------------------------------------------------
' Unlock an eForm instance.
' Errors if the token is invalid.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS
    
    Set oLockDBDS = New LockDBDS
    Call oLockDBDS.UnlockEFormInstance(sCon, sToken, lStudyId, sSite, lSubjectId, lEFormInstanceId)
    Set oLockDBDS = Nothing
    
End Sub

'----------------------------------------------------------------------
Public Sub UnlockEForm(ByVal sCon As String, _
                                        ByVal sToken As String, _
                                        ByVal lStudyId As Long, _
                                        ByVal lEFormID As Long)
'----------------------------------------------------------------------
' NCJ 1 Jun 06 - Unlock an SD edit eForm lock.
' Errors if the token is invalid.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS
    
    Set oLockDBDS = New LockDBDS
    Call oLockDBDS.UnlockEForm(sCon, sToken, lStudyId, lEFormID)
    Set oLockDBDS = Nothing
    
End Sub

'----------------------------------------------------------------------
Public Function LockDetailsStudy(ByVal sCon As String, _
                                ByVal lStudyId As Long) As String
'----------------------------------------------------------------------
' Returns User and LockTimeStamp separated by | for a specific lock.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    LockDetailsStudy = oLockDBDS.LockDetails(sCon, lStudyId)
    Set oLockDBDS = Nothing

End Function

'----------------------------------------------------------------------
Public Function LockDetailsEForm(ByVal sCon As String, _
                                ByVal lStudyId As Long, ByVal lEFormID As Long) As String
'----------------------------------------------------------------------
' Returns User and LockTimeStamp separated by | for a specific lock.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    LockDetailsEForm = oLockDBDS.LockDetails(sCon, lStudyId, , , lEFormID)
    Set oLockDBDS = Nothing

End Function

'----------------------------------------------------------------------
Public Function LockDetailsSubject(ByVal sCon As String, _
                                ByVal lStudyId As Long, _
                                ByVal sSite As String, _
                                ByVal lSubjectId As Long) As String
'----------------------------------------------------------------------
' Returns User and LockTimeStamp separated by | for a specific lock.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    LockDetailsSubject = oLockDBDS.LockDetails(sCon, lStudyId, sSite, lSubjectId)
    Set oLockDBDS = Nothing

End Function

'----------------------------------------------------------------------
Public Function LockDetailseFormInstance(ByVal sCon As String, _
                                ByVal lStudyId As Long, _
                                ByVal sSite As String, _
                                ByVal lSubjectId As Long, _
                                ByVal lEFormInstanceId As Long) As String
'----------------------------------------------------------------------
' Returns User and LockTimeStamp separated by | for a specific lock.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    LockDetailseFormInstance = oLockDBDS.LockDetails(sCon, lStudyId, sSite, lSubjectId, lEFormInstanceId)
    Set oLockDBDS = Nothing

End Function

'----------------------------------------------------------------------
Public Function LockDetailsSubjectSave(ByVal sCon As String, _
                                ByVal lStudyId As Long, _
                                ByVal sSite As String, _
                                ByVal lSubjectId As Long) As String
'----------------------------------------------------------------------
' Returns User and LockTimeStamp separated by | for a specific lock.
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    LockDetailsSubjectSave = oLockDBDS.LockDetails(sCon, lStudyId, sSite, lSubjectId, m_EFI_LOCK_SUBJECT_FOR_SAVE)
    Set oLockDBDS = Nothing

End Function

'-------------------------------------------------------------------------------------------------
Public Function AllLockDetails(ByRef sConnection As String, Optional sUserName As String = "") As Variant
'-------------------------------------------------------------------------------------------------
' Returns a array of details about current locks.
' This is to allow users to select expired locks to remove.
' If SUSerNAme is passed in then the locks returned will be filtered on that UserName
'-------------------------------------------------------------------------------------------------
'TA 02/10/2001:PERMISSIONS CHEcK REMOVED - CHECK IN USERINTERFACE OR IO LAYER
'MLM 21/05/02:  Modified to cope with new lock type for creation of new studies.
' NCJ 25 Oct 06 - Distinguish EformIds for DE and MUSD (use CRFPageDE and CRFPageSD)
' NCJ 30 Nov 06 - Bypass Query services here because there's a bug. Use explicit SQL instead
' NCJ 26 Feb 07 - Bug 2879 - Include user filtering in explicit SQL
'-------------------------------------------------------------------------------------------------
'Dim oQueryDef As QueryDef
'Dim oQueryDefs As QueryDefs
'
'    Set oQueryDef = New QueryDef
'    With oQueryDef
'        .QueryTables.Add "MACROLock"
'        'MLM 21/05/02: Use left join so that locks are included in results even if not in ClinicalTrial table
'        .QueryTables.Add "ClinicalTrial", , qdjtLeft, "MACROLock.StudyId", "ClinicalTrial.ClinicalTrialId"
'        .QueryTables.Add "CRFPageInstance", , qdjtLeft, _
'            Array("MACROLock.StudyId", "MACROLock.Site", "MACROLock.SubjectId", "MACROLock.EFormInstanceId"), _
'            Array("CRFPageInstance.ClinicalTrialId", "CRFPageInstance.TrialSite", "CRFPageInstance.PersonId", "CRFPageInstance.CRFPageTaskId")
'        ' This picks up the CRFTitle for Subject locks (alias CRFPageDE)
'        .QueryTables.Add "CRFPage CRFPageDE", "CRFPageDE", qdjtLeft, _
'            Array("CRFPageInstance.ClinicalTrialId", "CRFPageInstance.CRFPageId"), _
'            Array("CRFPageDE.ClinicalTrialId", "CRFPageDE.CRFPageId")
'        ' NCJ 25 Oct 06 - Interpret eForm locks in MUSD (alias CRFPageSD)
'        .QueryTables.Add "CRFPage CRFPageSD", "CRFPageSD", qdjtLeft, _
'            Array("MACROLock.StudyId", "MACROLock.EFormInstanceId"), _
'            Array("CRFPageSD.ClinicalTrialId", "CRFPageSD.CRFPageId")
'        .QueryFields.Add Array("ClinicalTrialName", "Site", "SubjectId", "CRFPageDE.CRFTitle", "CRFPageCycleNumber", _
'            "UserName", "LockTimeStamp", "StudyId", "EFormInstanceId", "Token", "CRFPageSD.CRFTitle")
'        If sUserName <> "" Then
'            'username passed through so filter on it
'            .QueryFilters.Add "UserName", "=", sUserName
'        End If
'        .QueryOrders.Add Array("ClinicalTrialName", "Site", "SubjectId", "EFormInstanceId")
'        .Code = "Locks"
'    End With
'
'    Set oQueryDefs = New QueryDefs
'    oQueryDefs.Init sConnection, ""
'    oQueryDefs.Add oQueryDef
'
'    AllLockDetails = oQueryDefs.GetDataArray("Locks")

Dim sSQL As String
Dim oDBCon As ADODB.Connection
Dim vLocks As Variant
Dim bSQLServer As Boolean
Dim rsLocks As ADODB.Recordset

    On Error GoTo ErrLabel
    
    ' Create database connection
    Set oDBCon = New ADODB.Connection
    Call oDBCon.Open(sConnection)
    oDBCon.CursorLocation = adUseClient
    
    ' NCJ 30 Nov 06 - These SQL strings were obtained by stepping through the code and
    ' scraping off oQueryDef.SQL, and then correcting them! (With Matthew's help)
    ' Decide whether we're SQLServer or Oracle
    bSQLServer = (InStr(Connection_Property(CONNECTION_PROVIDER, oDBCon.Connectionstring), CONNECTION_SQLOLEDB) > 0)
    
    ' The selections
    sSQL = "SELECT ClinicalTrialName, Site, SubjectId, CRFPageDE.CRFTitle, CRFPageCycleNumber, " _
            & "UserName, LockTimeStamp, StudyId, EFormInstanceId, Token, CRFPageSD.CRFTitle FROM "
    
    If bSQLServer Then
        sSQL = sSQL & "((((MACROLock LEFT JOIN ClinicalTrial ON MACROLock.StudyId = ClinicalTrial.ClinicalTrialId) " _
                & "LEFT JOIN CRFPageInstance ON MACROLock.StudyId = CRFPageInstance.ClinicalTrialId " _
                & "AND MACROLock.Site = CRFPageInstance.TrialSite AND MACROLock.SubjectId = CRFPageInstance.PersonId " _
                & "AND MACROLock.EFormInstanceId = CRFPageInstance.CRFPageTaskId) " _
                & "LEFT JOIN CRFPage CRFPageDE ON CRFPageInstance.ClinicalTrialId = CRFPageDE.ClinicalTrialId " _
                & "AND CRFPageInstance.CRFPageId = CRFPageDE.CRFPageId) " _
                & "LEFT JOIN CRFPage CRFPageSD ON MACROLock.StudyId = CRFPageSD.ClinicalTrialId " _
                & "AND MACROLock.EFormInstanceId = CRFPageSD.CRFPageId) "
        ' NCJ 26 Feb 07 - Bug 2879 - include user filtering if necessary
        If sUserName <> "" Then
            ' Username passed through so filter on it
            sSQL = sSQL & " WHERE MACROLock.UserName='" & sUserName & "'"
        End If
    Else
        ' Assume Oracle
        sSQL = sSQL & "MACROLock, ClinicalTrial, CRFPageInstance, CRFPage CRFPageDE, CRFPage CRFPageSD "
        sSQL = sSQL & " WHERE MACROLock.StudyId = ClinicalTrial.ClinicalTrialId(+) " _
                & "AND MACROLock.StudyId = CRFPageInstance.ClinicalTrialId(+) " _
                & "AND MACROLock.Site = CRFPageInstance.TrialSite(+) " _
                & "AND MACROLock.SubjectId = CRFPageInstance.PersonId(+) " _
                & "AND MACROLock.EFormInstanceId = CRFPageInstance.CRFPageTaskId(+) " _
                & "AND CRFPageInstance.ClinicalTrialId = CRFPageDE.ClinicalTrialId(+) " _
                & "AND CRFPageInstance.CRFPageId = CRFPageDE.CRFPageId(+) " _
                & "AND MACROLock.StudyId = CRFPageSD.ClinicalTrialId(+) " _
                & "AND MACROLock.EFormInstanceId = CRFPageSD.CRFPageId(+) "
        ' NCJ 26 Feb 07 - Bug 2879 - include user filtering if necessary
        If sUserName <> "" Then
            'username passed through so filter on it
            sSQL = sSQL & " AND MACROLock.UserName='" & sUserName & "'"
        End If
    End If
    
    
    ' Do the ordering
    sSQL = sSQL & " ORDER BY ClinicalTrialName, Site, SubjectId, EformInstanceId"
    
    Set rsLocks = New ADODB.Recordset
    rsLocks.Open sSQL, oDBCon
    If rsLocks.RecordCount > 0 Then
        vLocks = rsLocks.GetRows
    Else
        vLocks = Null
    End If
    Call rsLocks.Close
    Set rsLocks = Nothing
    Call oDBCon.Close
    Set oDBCon = Nothing
    
    AllLockDetails = vLocks
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|DBLock.AllLockDetails"
    
End Function

'----------------------------------------------------------------------
Public Function CacheAddStudyRow(sCon As String, lStudyId As Long) As String
'----------------------------------------------------------------------
' NCj 19 Jun 06 - Add a row in the cache for a study
' Returns token for cache entry
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    CacheAddStudyRow = oLockDBDS.CacheAddStudyRow(sCon, lStudyId)
    Set oLockDBDS = Nothing

End Function

'----------------------------------------------------------------------
Public Function CacheAddSubjectRow(sCon As String, _
                                lStudyId As Long, _
                                sSite As String, _
                                lSubjectId As Long) As String
'----------------------------------------------------------------------
' add a row in the cache for a subject
' Returns token for cache - to dlete it later on and used to exclude when invalidating cache
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    CacheAddSubjectRow = oLockDBDS.CacheAddSubjectRow(sCon, lStudyId, sSite, lSubjectId)
    Set oLockDBDS = Nothing

End Function

'----------------------------------------------------------------------
Public Sub CacheRemoveSubjectRow(sCon As String, sTokenToRemove As String)
'----------------------------------------------------------------------
'remove the row trhat corresponds to the given token
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    Call oLockDBDS.CacheRemoveSubjectRow(sCon, sTokenToRemove)
    Set oLockDBDS = Nothing
End Sub

'----------------------------------------------------------------------
Public Function CacheEntryStillValid(sCon As String, sCacheToken As String) As Boolean
'----------------------------------------------------------------------
'Return whether a cache entry is still valid
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    CacheEntryStillValid = oLockDBDS.CacheEntryStillValid(sCon, sCacheToken)
    Set oLockDBDS = Nothing
    
End Function

'----------------------------------------------------------------------
Public Sub CacheInvalidateStudy(sCon As String, _
                                lStudyId As Long, _
                                Optional sTokenToLeave As String = "")
'----------------------------------------------------------------------
' NCJ 19 Jun 06 - Invalidate everyone's cache tokens for this study
' but keep sTokenToLeave (if a row with that token exists, it will not be deleted)
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    Call oLockDBDS.CacheInvalidateStudy(sCon, lStudyId, sTokenToLeave)
    Set oLockDBDS = Nothing
    
End Sub

'----------------------------------------------------------------------
Public Sub CacheInvalidate(sCon As String, _
                                Optional lStudyId As Long, _
                                Optional sSite As String, _
                                Optional lSubjectId As Long, _
                                Optional sTokenToLeave As String)
'----------------------------------------------------------------------
' called with just connection string wipes the whole AREZZOToken table
' inlcude a study id and all subject entries for that study are wiped
' include a site and subject id and just that subject entry is wiped
'inlcude sTokenToLeave and if a row with that token exists, it will not be deleted
'----------------------------------------------------------------------
Dim oLockDBDS As LockDBDS

    Set oLockDBDS = New LockDBDS
    Call oLockDBDS.CacheInvalidate(sCon, lStudyId, sSite, lSubjectId, sTokenToLeave)
    Set oLockDBDS = Nothing
End Sub

