VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LockDBDS"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   File:       LockDBDS.cls
'   Copyright:  InferMed Ltd. 2001-2006. All Rights Reserved
'   Author:     Toby Aldridge, June 2001
'   Purpose:    Low level object locking class.
'----------------------------------------------------------------------------------------'
' Revisions:
'TA 02/10/2001: Randomize called on class creation to ensure unique tokens
'TA 18/07/2002 CBB2.2.20.2: Try to get lock several times when locking subject
' TA 12/09/2002: Added support for locking the subject for saving
'TA 19/03/2003: Try to get lock several times in LockeFormInstance as in LockSubject
' NCJ 24 May 06 - Added stuff for MultiUser SD (eForm locks)
' NCJ 1 Jun 06 - Checked all lock routines for MUSD
' NCJ 19 Jun 06 - Added CacheAddStudyRow
' NCJ 27 Jun 06 - Added OpenStudyLO
' NCJ 24 Oct 06 - Try to get a user name in LockDetails, even if only for the study; check for LO locks in LockStudy
'----------------------------------------------------------------------------------------

Option Explicit

Private Const m_LOCK_TABLE = "MACROLock"
'table MACROLock has columns StudyId - long, Site - text(8), SubjectId - long,
'           EFormInstanceId - long, UserName - text, LockTimeStamp - double, Token - text(10)
' key is study,site,subject,eforminstance

Private Const m_CACHE_TABLE = "AREZZOToken"

'enumeration of what is locked
' NCJ 24 may 06 - Added dblEForm
Public Enum DBDSLock
    dblStudy
    dblSubject
    dblEFormInstance
    dblEForm
End Enum


Private Const m_NO_SITE = "none"
Private Const m_NO_SUBJECT = 0
Private Const m_NO_EFORMINSTANCE = 0


Private Const m_COL_STUDY = "StudyId"
Private Const m_COL_SITE = "Site"
Private Const m_COL_SUBJECT = "SubjectId"
Private Const m_COL_EFORMINSTANCE = "EformInstanceId"
Private Const m_COL_USER = "UserName"
Private Const m_COL_LOCKTIMESTAMP = "LockTimeStamp"
Private Const m_COL_TOKEN = "Token"

Private Const m_ERR_NO_INVALIDTOKEN = 1010
Private Const m_ERR_DESC_INVALIDTOKEN = "Token does not correspond to this lock"

'how many atttmepts made and how long between each attempt to lock a subject
Private Const m_RETRY_SUBJECT_LOCK_ATTEMPTS As Long = 8
Private Const m_RETRY_SUBJECT_LOCK_PAUSE As Long = 250   'milliseconds

'TA 12/09/2002: mock efi id to signify locking the subject for saving
Private Const m_EFI_LOCK_SUBJECT_FOR_SAVE = -1
Private Const m_RETRY_SAVE_SUBJECT_LOCK_ATTEMPTS As Long = 8
Private Const m_RETRY_SAVE_SUBJECT_LOCK_PAUSE As Long = 250   'milliseconds

'----------------------------------------------------------------------
Public Function LockStudyRW(ByVal sCon As String, _
                            ByVal sUser As String, _
                            ByVal lStudyId As Long) As String
'----------------------------------------------------------------------
' NCJ 1 Jun 06 - Lock a study for RW editing.
' Returns token for unlocking if successful,
' or returns enum of the level of object locked.
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer
Dim sToken As String
Dim sLocked As String
Dim lRetry As Long

    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sCon
    
    ' Try to get "global" study lock a few times
    For lRetry = 1 To m_RETRY_SUBJECT_LOCK_ATTEMPTS
        'try to lock study
        sToken = InsertLock(oQueryServer, sUser, lStudyId)
        'if we have a lock then don't bother retrying
        If sToken <> "" Then Exit For
        'pause before trying again
        Call Sleep(m_RETRY_SUBJECT_LOCK_PAUSE)
    Next
    
    If sToken = "" Then
        ' Global lock failed - study already locked - exit
        sLocked = dblStudy
        GoTo ErrLocked
    End If
    
    ' Now try to lock study with RW lock
    sToken = InsertLock(oQueryServer, sUser, lStudyId, m_NO_SITE, m_NO_SUBJECT, m_EFI_LOCK_SUBJECT_FOR_SAVE)
    If sToken = "" Then
        ' RW study lock failed - study already RW locked - remove global lock and exit
        RemoveLock oQueryServer, lStudyId
        sLocked = dblStudy
        GoTo ErrLocked
    End If
    
    If IsStudySubjectLock(oQueryServer, lStudyId) Then
        ' There's already a subject lock, so remove our other attempted locks and fail
        Call RemoveLock(oQueryServer, lStudyId, m_NO_SITE, m_NO_SUBJECT, m_EFI_LOCK_SUBJECT_FOR_SAVE)
        Call RemoveLock(oQueryServer, lStudyId)
        sLocked = dblSubject
        GoTo ErrLocked
    End If
    
    ' Looks like we're OK - remove initial global lock and exit
    RemoveLock oQueryServer, lStudyId
    
    Set oQueryServer = Nothing
    
    'study now locked - return RW lock token
    LockStudyRW = sToken
    
    Exit Function
    
ErrLocked:
    Set oQueryServer = Nothing
    LockStudyRW = sLocked
End Function

'----------------------------------------------------------------------
Public Function OpenStudyLO(ByVal sCon As String, _
                            ByVal sUser As String, _
                            ByVal lStudyId As Long) As String
'----------------------------------------------------------------------
' NCJ 27 Jun 06 - Can we open the study LO?
' Not if there's a full study lock OR any subject locks
' If so, returns type of lock preventing LO, otherwise returns empty string
' Doesn't actually place any locks
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer
Dim sToken As String
Dim sLocked As String

    sLocked = ""
    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sCon
    
    'try to lock study
    sToken = InsertLock(oQueryServer, sUser, lStudyId)
    If sToken = "" Then
        'insert study key failed - study already locked - exit
        sLocked = dblStudy
    Else
        ' See if there are any subject locks
        If IsStudySubjectLock(oQueryServer, lStudyId) Then
            ' Subject lock
            sLocked = dblSubject
        End If
        ' Remove our study lock
        RemoveLock oQueryServer, lStudyId
    End If
    
    Set oQueryServer = Nothing
    
    ' Return type of lock, or empty string if none
    OpenStudyLO = sLocked
    
End Function

'----------------------------------------------------------------------
Public Function LockStudy(ByVal sCon As String, _
                            ByVal sUser As String, _
                            ByVal lStudyId As Long) As String
'----------------------------------------------------------------------
' Lock a study "globally", i.e. to allow no other access.
' Returns token for unlocking if successful.
' REturns enum of the level of object locked.
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer
Dim sToken As String
Dim sLocked As String

    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sCon
    
    'try to lock study
    sToken = InsertLock(oQueryServer, sUser, lStudyId)
    If sToken = "" Then
        'insert study key failed - study already locked - exit
        sLocked = dblStudy
        GoTo ErrLocked
    End If
    
    ' NCJ 24 Oct 06 - Check for Layout Only lock on study
    If IsStudyEditLock(oQueryServer, lStudyId) Then
        RemoveLock oQueryServer, lStudyId
        sLocked = dblStudy
        GoTo ErrLocked
    End If
    
    If SubObjectLocked(oQueryServer, lStudyId) Then
        'sub objects locked - remove study lock and exit
        RemoveLock oQueryServer, lStudyId
        sLocked = dblSubject
        GoTo ErrLocked
    End If
    
    Set oQueryServer = Nothing
    
    'study now locked - return token
    LockStudy = sToken
    
    Exit Function
    
ErrLocked:
    Set oQueryServer = Nothing
    LockStudy = sLocked
End Function

'----------------------------------------------------------------------
Public Function LockSubject(ByVal sCon As String, _
                                ByVal sUser As String, _
                                ByVal lStudyId As Long, _
                                ByVal sSite As String, _
                                ByVal lSubjectId As Long) As String
'----------------------------------------------------------------------
' Lock a subject.
' Returns token for unlocking if successful,
' or returns enum of the level of object locked.
' TA 18/07/2002 CBB2.2.20.2: Try to get lock several times
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer
Dim sToken As String
Dim sLocked As String
Dim lRetry As Long

    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sCon
    
    'TA 18/07/2002 CBB2.2.20.2: Try to get lock several times
    For lRetry = 1 To m_RETRY_SUBJECT_LOCK_ATTEMPTS
        'try to lock study
        sToken = InsertLock(oQueryServer, sUser, lStudyId)
        'if we have a lock then don't bother retrying
        If sToken <> "" Then Exit For
        'pause before trying again
        Call Sleep(m_RETRY_SUBJECT_LOCK_PAUSE)
    Next
    
    If sToken = "" Then
        'insert study key failed - study already locked - exit
        sLocked = dblStudy
        GoTo ErrLocked
    End If
    
    'Study now locked - check no study edit locks in place
    If IsStudyEditLock(oQueryServer, lStudyId) Then
        ' Someone's editing the study - remove global lock and fail
        RemoveLock oQueryServer, lStudyId
        sLocked = dblStudy
        GoTo ErrLocked
    End If
    
    ' Now try to lock subject
    sToken = InsertLock(oQueryServer, sUser, lStudyId, sSite, lSubjectId)
    If sToken = "" Then
        'insert subject key failed - subject already locked - unlock study and exit
        RemoveLock oQueryServer, lStudyId
        sLocked = dblSubject
        GoTo ErrLocked
    End If

    If SubObjectLocked(oQueryServer, lStudyId, sSite, lSubjectId) Then
        'sub objects locked - remove subject and study lock and exit
        RemoveLock oQueryServer, lStudyId, sSite, lSubjectId
        RemoveLock oQueryServer, lStudyId
        sLocked = dblEFormInstance
        GoTo ErrLocked
    End If
    
    'subject now locked so remove study lock
    RemoveLock oQueryServer, lStudyId
    
    Set oQueryServer = Nothing
    
    'return token
    LockSubject = sToken

    Exit Function
    
ErrLocked:
    Set oQueryServer = Nothing
    LockSubject = sLocked
End Function
    
'----------------------------------------------------------------------
Public Function LockSubjectForSaving(ByVal sCon As String, _
                                    ByVal sUser As String, _
                                    ByVal lStudyId As Long, _
                                    ByVal sSite As String, _
                                    ByVal lSubjectId As Long) As String
'----------------------------------------------------------------------
' Lock the subject for saving by doing an eform lock with a dummy efi id
'----------------------------------------------------------------------
                                    
    LockSubjectForSaving = LockEFormInstance(sCon, sUser, lStudyId, sSite, lSubjectId, m_EFI_LOCK_SUBJECT_FOR_SAVE)
    
End Function

 '----------------------------------------------------------------------
Public Function LockEForm(ByVal sCon As String, _
                                    ByVal sUser As String, _
                                    ByVal lStudyId As Long, _
                                    ByVal lEFormId As Long) As String
'----------------------------------------------------------------------
' NCJ 24 May 06 - Lock an eForm for editing in Study Definition
' Returns token for unlocking if successful.
' Returns enum of the level of object locked.
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer
Dim sToken As String
Dim sLocked As String
Dim lRetry As Long

    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sCon
    
    ' Check the global lock - try a few times
    For lRetry = 1 To m_RETRY_SUBJECT_LOCK_ATTEMPTS
        'try to lock study
        sToken = InsertLock(oQueryServer, sUser, lStudyId)
        'if we have a lock then don't bother retrying
        If sToken <> "" Then Exit For
        'pause before trying again
        Call Sleep(m_RETRY_SUBJECT_LOCK_PAUSE)
    Next
    
    If sToken = "" Then
       ' Study already globally locked
       sLocked = dblStudy
       GoTo ErrLocked
    End If
    
    ' Lock eForm, using given User, StudyId and eFormId
    sToken = InsertLock(oQueryServer, sUser, lStudyId, m_NO_SITE, m_NO_SUBJECT, lEFormId)
    If sToken = "" Then
        'insert eform key failed - eform already locked
        sLocked = dblEForm
        RemoveLock oQueryServer, lStudyId
        GoTo ErrLocked
    End If

    ' Now check there aren't any subject locks here
    If IsStudySubjectLock(oQueryServer, lStudyId) Then
        Call RemoveLock(oQueryServer, lStudyId, m_NO_SITE, m_NO_SUBJECT, lEFormId)
        Call RemoveLock(oQueryServer, lStudyId)
        sLocked = dblEFormInstance
        GoTo ErrLocked
    End If
    
    ' Remove the global study lock
    RemoveLock oQueryServer, lStudyId
    
    Set oQueryServer = Nothing
    
    'return token
    LockEForm = sToken

    Exit Function
    
ErrLocked:
    Set oQueryServer = Nothing
    LockEForm = sLocked

End Function
  
'----------------------------------------------------------------------
Public Function LockEFormInstance(ByVal sCon As String, _
                                    ByVal sUser As String, _
                                    ByVal lStudyId As Long, _
                                    ByVal sSite As String, _
                                    ByVal lSubjectId As Long, _
                                    ByVal lEFormInstanceId As Long) As String
'----------------------------------------------------------------------
' Lock an eForm instance.
' Returns token for unlocking if successful.
' REturns enum of the level of object locked.
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer
Dim sToken As String
Dim sLocked As String
Dim lRetry As Long

    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sCon
    
    'try to lock study
    'TA 19/03/2003: Try to get lock several times as in LockSubject
    For lRetry = 1 To m_RETRY_SUBJECT_LOCK_ATTEMPTS
        'try to lock study
        sToken = InsertLock(oQueryServer, sUser, lStudyId)
        'if we have a lock then don't bother retrying
        If sToken <> "" Then Exit For
        'pause before trying again
        Call Sleep(m_RETRY_SUBJECT_LOCK_PAUSE)
    Next
    
    If sToken = "" Then
        'insert study key failed - study already locked - exit
        sLocked = dblStudy
        GoTo ErrLocked
    End If
    
    'Study now locked - check no RW or LO study edit locks in place
    If IsStudyEditLock(oQueryServer, lStudyId) Then
        ' Someone's editing the study - remove global lock and fail
        RemoveLock oQueryServer, lStudyId
        sLocked = dblStudy
        GoTo ErrLocked
    End If
    
    'Study now locked - try to lock subject
    sToken = InsertLock(oQueryServer, sUser, lStudyId, sSite, lSubjectId)
    If sToken = "" Then
        'insert subject key failed - subject already locked - unlock study and exit
        RemoveLock oQueryServer, lStudyId
        sLocked = dblSubject
        GoTo ErrLocked
    End If
    
    ' If we want a "save" lock, try a few times for the subject lock
    If lEFormInstanceId = m_EFI_LOCK_SUBJECT_FOR_SAVE Then
        'TA 18/07/2002 locking subject for save - Try to get lock several times
        For lRetry = 1 To m_RETRY_SAVE_SUBJECT_LOCK_ATTEMPTS
            'try to lock with dummy id
            sToken = InsertLock(oQueryServer, sUser, lStudyId, sSite, lSubjectId, m_EFI_LOCK_SUBJECT_FOR_SAVE)
            'if we have a lock then don't bother retrying
            If sToken <> "" Then Exit For
            'pause before trying again
            Call Sleep(m_RETRY_SAVE_SUBJECT_LOCK_PAUSE)
        Next
    Else
        'subject now locked - try to lock eForm Instance
         sToken = InsertLock(oQueryServer, sUser, lStudyId, sSite, lSubjectId, lEFormInstanceId)
    End If
         
    ' No longer need subject and study locks
    RemoveLock oQueryServer, lStudyId, sSite, lSubjectId
    RemoveLock oQueryServer, lStudyId
    
    If sToken = "" Then
        'insert EFI lock failed - eform already locked
        sLocked = dblEFormInstance
        GoTo ErrLocked
    End If

    Set oQueryServer = Nothing
    
    'return token
    LockEFormInstance = sToken

    Exit Function
    
ErrLocked:
    Set oQueryServer = Nothing
    LockEFormInstance = sLocked

End Function

'----------------------------------------------------------------------
Public Sub UnlockStudy(ByVal sCon As String, _
                        ByVal sToken As String, _
                        ByVal lStudyId As Long)
'----------------------------------------------------------------------
' Unlock a study.
' Errors if the token is invalid.
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer

    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sCon
    If Not ValidToken(oQueryServer, sToken, lStudyId) Then
        'invalid token - raise error
        Err.Raise vbObjectError + m_ERR_NO_INVALIDTOKEN, , m_ERR_DESC_INVALIDTOKEN
    End If
    'valid token so remove lock
    RemoveLock oQueryServer, lStudyId
    Set oQueryServer = Nothing
    
End Sub
   
'----------------------------------------------------------------------
Public Sub UnlockStudyRW(ByVal sCon As String, _
                        ByVal sToken As String, _
                        ByVal lStudyId As Long)
'----------------------------------------------------------------------
' NCJ 1 Jun 06 - Unlock a study as RW.
' Errors if the token is invalid.
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer

    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sCon
    If Not ValidToken(oQueryServer, sToken, lStudyId, m_NO_SITE, m_NO_SUBJECT, m_EFI_LOCK_SUBJECT_FOR_SAVE) Then
        'invalid token - raise error
        Err.Raise vbObjectError + m_ERR_NO_INVALIDTOKEN, , m_ERR_DESC_INVALIDTOKEN
    End If
    ' valid token so remove lock
    Call RemoveLock(oQueryServer, lStudyId, m_NO_SITE, m_NO_SUBJECT, m_EFI_LOCK_SUBJECT_FOR_SAVE)
    Set oQueryServer = Nothing
    
End Sub
  
'----------------------------------------------------------------------
Public Function UnlockSubject(ByVal sCon As String, _
                                ByVal sToken As String, _
                                ByVal lStudyId As Long, _
                                ByVal sSite As String, _
                                ByVal lSubjectId As Long)
'----------------------------------------------------------------------
' Unlock a subject.
' Errors if the token is invalid.
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer

    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sCon
    If Not ValidToken(oQueryServer, sToken, lStudyId, sSite, lSubjectId) Then
        'invalid token - raise error
        Err.Raise vbObjectError + m_ERR_NO_INVALIDTOKEN, , m_ERR_DESC_INVALIDTOKEN
    End If
    'valid token so remove lock
    RemoveLock oQueryServer, lStudyId, sSite, lSubjectId
    Set oQueryServer = Nothing
        
End Function

'----------------------------------------------------------------------
Public Function UnlockEFormInstance(ByVal sCon As String, _
                                        ByVal sToken As String, _
                                        ByVal lStudyId As Long, _
                                        ByVal sSite As String, _
                                        ByVal lSubjectId As Long, _
                                        lEFormInstanceId As Long)
'----------------------------------------------------------------------
' Unlock an eForm instance.
' Errors if the token is invalid.
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer

    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sCon
    
    If Not ValidToken(oQueryServer, sToken, lStudyId, sSite, lSubjectId, lEFormInstanceId) Then
        'invalid token - raise error
        Err.Raise vbObjectError + m_ERR_NO_INVALIDTOKEN, , m_ERR_DESC_INVALIDTOKEN
    End If
    'valid token so remove lock
    
    RemoveLock oQueryServer, lStudyId, sSite, lSubjectId, lEFormInstanceId

    Set oQueryServer = Nothing
    
End Function

'----------------------------------------------------------------------
Public Function UnlockEForm(ByVal sCon As String, _
                                        ByVal sToken As String, _
                                        ByVal lStudyId As Long, _
                                        lEFormId As Long)
'----------------------------------------------------------------------
' NCJ 1 Jun 06 - Unlock an SD edit eForm.
' Errors if the token is invalid.
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer

    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sCon
    
    If Not ValidToken(oQueryServer, sToken, lStudyId, m_NO_SITE, m_NO_SUBJECT, lEFormId) Then
        'invalid token - raise error
        Err.Raise vbObjectError + m_ERR_NO_INVALIDTOKEN, , m_ERR_DESC_INVALIDTOKEN
    End If
    'valid token so remove lock
    
    RemoveLock oQueryServer, lStudyId, m_NO_SITE, m_NO_SUBJECT, lEFormId

    Set oQueryServer = Nothing
    
End Function

'----------------------------------------------------------------------
Public Function LockDetails(ByVal sCon As String, _
                                ByVal lStudyId As Long, _
                                Optional ByVal sSite As String = m_NO_SITE, _
                                Optional ByVal lSubjectId As Long = m_NO_SUBJECT, _
                                Optional ByVal lEFormInstanceId As Long = m_NO_EFORMINSTANCE) As String
'----------------------------------------------------------------------
' Returns User and LockTimeStamp separated by | for a specific lock.
'TA 19/09/2002: added token to the return
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer
Dim oQD As QueryDef
Dim vRow As Variant

    Set oQD = New QueryDef
    oQD.InitSelect m_LOCK_TABLE, Array(m_COL_USER, m_COL_LOCKTIMESTAMP, m_COL_TOKEN), _
            Array(m_COL_STUDY, m_COL_SITE, m_COL_SUBJECT, m_COL_EFORMINSTANCE), _
            Array(lStudyId, sSite, lSubjectId, lEFormInstanceId)
    
    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sCon
    vRow = oQueryServer.SelectArray(oQD)
    If IsNull(vRow) Then
        ' NCJ 24 Oct 06 - Try again with the study only
        oQD.InitSelect m_LOCK_TABLE, Array(m_COL_USER, m_COL_LOCKTIMESTAMP, m_COL_TOKEN), _
            Array(m_COL_STUDY), Array(lStudyId)
            vRow = oQueryServer.SelectArray(oQD)
    End If
    
    If IsNull(vRow) Then
        LockDetails = ""
    Else
        LockDetails = vRow(0, 0) & "|" & vRow(1, 0) & "|" & vRow(2, 0)
    End If

    Set oQD = Nothing
    Set oQueryServer = Nothing
    
End Function

'----------------------------------------------------------------------
Private Sub RemoveLock(oQueryServer As QueryServer, _
                                lStudyId As Long, _
                                Optional sSite As String = m_NO_SITE, _
                                Optional lSubjectId As Long = m_NO_SUBJECT, _
                                Optional lEFormInstanceId As Long = m_NO_EFORMINSTANCE)
'----------------------------------------------------------------------
' Remove a row from the lock table
'----------------------------------------------------------------------

Dim oQD As QueryDef

    Set oQD = New QueryDef
    oQD.InitDelete m_LOCK_TABLE, _
                    Array(m_COL_STUDY, m_COL_SITE, m_COL_SUBJECT, m_COL_EFORMINSTANCE), _
                    Array(lStudyId, sSite, lSubjectId, lEFormInstanceId)
       
    oQueryServer.SelectDelete oQD
    

End Sub

'----------------------------------------------------------------------
Private Function ValidToken(oQueryServer As QueryServer, _
                                sToken As String, _
                                lStudyId As Long, _
                                Optional sSite As String = m_NO_SITE, _
                                Optional lSubjectId As Long = m_NO_SUBJECT, _
                                Optional lEFormInstanceId As Long = m_NO_EFORMINSTANCE) As Boolean
'----------------------------------------------------------------------
' Return whether a token is valid for a lock table row.
'----------------------------------------------------------------------
Dim oQD As QueryDef
Dim vCount As Variant

    Set oQD = New QueryDef
    oQD.InitSelect m_LOCK_TABLE, , _
                    Array(m_COL_STUDY, m_COL_SITE, m_COL_SUBJECT, m_COL_EFORMINSTANCE, m_COL_TOKEN), _
                    Array(lStudyId, sSite, lSubjectId, lEFormInstanceId, sToken)
       
    oQD.Count = True
    
    vCount = oQueryServer.SelectArray(oQD)

    Set oQD = Nothing
    
    ValidToken = (vCount(0, 0) = 1)

End Function

'----------------------------------------------------------------------
Private Function SubObjectLocked(oQueryServer As QueryServer, _
                                lStudyId As Long, _
                                Optional sSite As String = m_NO_SITE, _
                                Optional lSubjectId As Long = m_NO_SUBJECT, _
                                Optional lEFormInstanceId As Long = m_NO_EFORMINSTANCE) As Boolean
'----------------------------------------------------------------------
' Returns whether an object's subobjects are locked.
'----------------------------------------------------------------------
Dim oQD As QueryDef
Dim vCount As Variant

    Set oQD = New QueryDef
    
    oQD.InitSelect m_LOCK_TABLE, , m_COL_STUDY, lStudyId
    oQD.Count = True
    
    If lSubjectId <> m_NO_SUBJECT Then
        oQD.QueryFilters.Add m_COL_SITE, "=", sSite
        oQD.QueryFilters.Add m_COL_SUBJECT, "=", lSubjectId
    End If

    vCount = oQueryServer.SelectArray(oQD)

    Set oQD = Nothing
    SubObjectLocked = vCount(0, 0) > 1


End Function

'----------------------------------------------------------------------
Private Function IsStudySubjectLock(oQueryServer As QueryServer, _
                                lStudyId As Long) As Boolean
'----------------------------------------------------------------------
' NCJ 24 May 06 - Is there a subject lock for this study?
' Look for a row that contains this Study ID and a "real" site
'----------------------------------------------------------------------
Dim oQD As QueryDef
Dim vCount As Variant

    Set oQD = New QueryDef
    
    oQD.InitSelect m_LOCK_TABLE, , m_COL_STUDY, lStudyId
    oQD.Count = True
    
    ' Look for a "real" site
    oQD.QueryFilters.Add m_COL_SITE, "<>", m_NO_SITE

    vCount = oQueryServer.SelectArray(oQD)

    Set oQD = Nothing
    IsStudySubjectLock = vCount(0, 0) > 0


End Function

'----------------------------------------------------------------------
Private Function IsStudyEditLock(oQueryServer As QueryServer, _
                                lStudyId As Long) As Boolean
'----------------------------------------------------------------------
' NCJ 1 Jun 06 - Is there a study edit lock for this study?
' Look for a row that contains this Study ID, no site and non-zero EFI
'----------------------------------------------------------------------
Dim oQD As QueryDef
Dim vCount As Variant

    Set oQD = New QueryDef
    
    oQD.InitSelect m_LOCK_TABLE, , m_COL_STUDY, lStudyId
    oQD.Count = True
    
    ' Look for no site and non-zero EFI
    oQD.QueryFilters.Add m_COL_SITE, "=", m_NO_SITE
    oQD.QueryFilters.Add m_COL_EFORMINSTANCE, "<>", m_NO_EFORMINSTANCE

    vCount = oQueryServer.SelectArray(oQD)

    Set oQD = Nothing
    IsStudyEditLock = vCount(0, 0) > 0


End Function

'----------------------------------------------------------------------
Public Function LockedEForms(ByVal sCon As String, ByVal sUser As String, ByVal lStudyId As String) As Collection
'----------------------------------------------------------------------
' Returns a collection of eFormIds that are currently locked
' in this study by anyone other than this user
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer
Dim oQD As QueryDef
Dim vRows As Variant
Dim colIDs As Collection
Dim i As Long

    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sCon
    
    Set oQD = New QueryDef
    
    ' Select eFormIds > 0 for other users (this study, no site)
    oQD.InitSelect m_LOCK_TABLE, m_COL_EFORMINSTANCE, Array(m_COL_STUDY, m_COL_SITE), Array(lStudyId, m_NO_SITE)
    oQD.QueryFilters.Add m_COL_USER, "<>", sUser
    oQD.QueryFilters.Add m_COL_EFORMINSTANCE, ">", 0
    
    vRows = oQueryServer.SelectArray(oQD)
    Set colIDs = New Collection
    If Not IsNull(vRows) Then
        For i = 0 To UBound(vRows, 2)
            colIDs.Add vRows(0, i)
        Next i
    End If
    
    Set LockedEForms = colIDs
    
    Set oQD = Nothing
    Set oQueryServer = Nothing
    Set colIDs = Nothing
    
End Function

'----------------------------------------------------------------------
Private Function InsertLock(oQueryServer As QueryServer, _
                                sUser As String, _
                                lStudyId As Long, _
                                Optional sSite As String = m_NO_SITE, _
                                Optional lSubjectId As Long = m_NO_SUBJECT, _
                                Optional lEFormInstanceId As Long = m_NO_EFORMINSTANCE) As String
'----------------------------------------------------------------------
' Insert a row into the lock table.
'----------------------------------------------------------------------
Dim oQD As QueryDef
Dim sToken As String

    Set oQD = New QueryDef
    
    sToken = Token
    oQD.InitSave m_LOCK_TABLE, _
                    Array(m_COL_STUDY, m_COL_SITE, m_COL_SUBJECT, m_COL_EFORMINSTANCE, m_COL_USER, m_COL_LOCKTIMESTAMP, m_COL_TOKEN), _
                    Array(lStudyId, sSite, lSubjectId, lEFormInstanceId, sUser, CDbl(Now), sToken)
                    
    On Error GoTo ErrLabel
    oQueryServer.SelectSave stInsert, oQD
    On Error GoTo 0
    
    InsertLock = sToken
    Exit Function
    
ErrLabel:
    InsertLock = ""
    If Err.Number <> QS_ERR_PRIMARY_KEY_VIOLATED Then
        Err.Raise Err.Number
    End If
                                

End Function

Private Function Token() As String
'----------------------------------------------------------------------
' Calculate a new token
'----------------------------------------------------------------------
Dim nAscStart As Integer
Dim i As Long
Dim sToken As String

    For i = 1 To 10
        sToken = sToken & Chr$(65 + RndLong(25, True))
    Next
    
    Token = sToken
    
End Function

'----------------------------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------------------------
' Prevent identical tokens being created when running as a DLL
'----------------------------------------------------------------------

    Randomize
    
End Sub

'----------------------------------------------------------------------
Public Function CacheAddStudyRow(sCon As String, lStudyId As Long) As String
'----------------------------------------------------------------------
' NCJ 19 June 06 - Add a row in the cache for a study
' Returns token for cache entry
'----------------------------------------------------------------------
Dim oQS As QueryServer
Dim oQD As QueryDef
Dim sToken As String

    sToken = Token
    
    Set oQD = New QueryDef
    oQD.InitSave m_CACHE_TABLE, Array("ArezzoId", "ClinicalTrialId", "TrialSite", "Personid"), _
                    Array(sToken, lStudyId, m_NO_SITE, m_NO_SUBJECT)

    Set oQS = New QueryServer
    oQS.ConnectionOpen sCon
    oQS.QueryInsert oQD
    Set oQS = Nothing
    Set oQD = Nothing
    
    CacheAddStudyRow = sToken
    
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
Dim oQS As QueryServer
Dim oQD As QueryDef
Dim sToken As String

    sToken = Token
    
    Set oQD = New QueryDef
    oQD.InitSave m_CACHE_TABLE, Array("ArezzoId", "ClinicalTrialId", "TrialSite", "Personid"), _
                    Array(sToken, lStudyId, sSite, lSubjectId)

    Set oQS = New QueryServer
    oQS.ConnectionOpen sCon
    oQS.QueryInsert oQD
    Set oQS = Nothing
    Set oQD = Nothing
    
    CacheAddSubjectRow = sToken
    
End Function

Public Sub CacheRemoveSubjectRow(sCon As String, sTokenToRemove As String)
'----------------------------------------------------------------------
'remove the row trhat corresponds to the given token
'----------------------------------------------------------------------
Dim oQS As QueryServer
Dim oQD As QueryDef

    
    Set oQD = New QueryDef
    oQD.InitDelete m_CACHE_TABLE, "ArezzoId", sTokenToRemove

    Set oQS = New QueryServer
    oQS.ConnectionOpen sCon
    oQS.SelectDelete oQD
    Set oQS = Nothing
    Set oQD = Nothing
    
End Sub

Public Function CacheEntryStillValid(sCon As String, sSubjectToken As String) As Boolean
'----------------------------------------------------------------------
'Return whether a cache entry is still valid
'----------------------------------------------------------------------
 Dim oQS As QueryServer
Dim oQD As QueryDef
Dim vCount As Variant
    
    Set oQD = New QueryDef
    oQD.InitSelect m_CACHE_TABLE, "Count(*)", "ArezzoId", sSubjectToken

    Set oQS = New QueryServer
    oQS.ConnectionOpen sCon
    vCount = oQS.SelectArray(oQD)
    Set oQS = Nothing
    Set oQD = Nothing
    
    'our count(*) query should always return the value 0 or 1
    CacheEntryStillValid = (vCount(0, 0) = 1)
    
    
End Function

'----------------------------------------------------------------------
Public Sub CacheInvalidateStudy(sCon As String, _
                                lStudyId As Long, _
                                Optional sTokenToLeave As String = "")
'----------------------------------------------------------------------
' NCJ 19 Jun 06 - Invalidate everyone's cache tokens for this study
' but keep sTokenToLeave (if a row with that token exists, it will not be deleted)
'----------------------------------------------------------------------

    Call CacheInvalidate(sCon, lStudyId, m_NO_SITE, m_NO_SUBJECT, sTokenToLeave)
    
End Sub

'----------------------------------------------------------------------
Public Sub CacheInvalidate(sCon As String, _
                                Optional lStudyId As Long = -1, _
                                Optional sSite As String = "", _
                                Optional lSubjectId As Long = -1, _
                                Optional sTokenToLeave As String = "")
'----------------------------------------------------------------------
' called with just connection string wipes the whole AREZZOToken table
' inlcude a study id and all subject entries for that study are wiped
' include a site and subject id and just that subject entry is wiped
' include sTokenToLeave and if a row with that token exists, it will not be deleted
'----------------------------------------------------------------------

Dim oQS As QueryServer
Dim oQD As QueryDef

    
    Set oQD = New QueryDef
    oQD.InitDelete m_CACHE_TABLE

    If lStudyId <> -1 Then
        'have a study id - just delete for this study
        oQD.QueryFilters.Add "ClinicalTrialId", "=", lStudyId
        If sSite <> "" Then
            'have a site (therefore a subject -  delete for this subject
            oQD.QueryFilters.Add Array("TrialSite", "PersonId"), Array("=", "="), Array(sSite, lSubjectId)
        End If
    End If
    
    If sTokenToLeave <> "" Then
        'with been given a row identifier to ignore
        oQD.QueryFilters.Add "AREZZOId", "<>", sTokenToLeave
    End If

    Set oQS = New QueryServer
    oQS.ConnectionOpen sCon
    
    oQS.SelectDelete oQD
    Set oQS = Nothing
    
    Set oQD = Nothing
    
End Sub




