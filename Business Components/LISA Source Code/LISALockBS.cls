VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LISALockBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       LISALockBS.cls
'   Copyright:  InferMed Ltd. 2004. All Rights Reserved
'   Author:     Nicky Johns, July 2004
'   Purpose:    Low level object locking class for use with LISA only
'----------------------------------------------------------------------------------------'
' Revisions:
' NCJ 15 July 04 - Created as copy of LockDBDS.cls from MACROLockDS30
'               Requires the LISALockedForm object class
' NCJ 19 Jul 04 - Removed all routines not used by LISA
'----------------------------------------------------------------------------------------'

Option Explicit

Private Const m_LOCK_TABLE = "MACROLock"
'table MACROLock has columns StudyId - long, Site - text(8), SubjectId - long,
'           EFormInstanceId - long, UserName - text, LockTimeStamp - double, Token - text(10)
' key is study,site,subject,eforminstance

'enumeration of what is locked
Public Enum LISALockType
    lkStudy
    lkSubject
    lkEFormInstance
End Enum


Private Const QS_ERR_UNROCOGNISED_PROVIDER = 1000 Or vbObjectError
Private Const QS_ERR_PRIMARY_KEY_VIOLATED = 1001 Or vbObjectError

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

'how many atttmepts made and how long between each attemp to lock a subject
Private Const m_RETRY_SUBJECT_LOCK_ATTEMPTS As Long = 8
Private Const m_RETRY_SUBJECT_LOCK_PAUSE As Long = 250   'milliseconds

'----------------------------------------------------------------------
Public Function LockEFormInstances(ByVal sCon As String, _
                                    ByVal sUser As String, _
                                    ByVal lStudyId As Long, _
                                    ByVal sSite As String, _
                                    ByVal lSubjectId As Long, _
                                    ByRef colEFormInstanceIds As Collection, _
                                    ByRef colLockedForms As Collection, _
                                    ByRef sMsg As String) As Boolean
'----------------------------------------------------------------------
' Lock a collection of eForm instances
' If successful, adds relevant LISALockedForm objects to colLockedForms collection
' If unsuccessful, the entire LockedForms collection is unlocked and emptied
' If unsuccessful, sMsg returns what happened
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer
Dim sToken As String
Dim lRetry As Long
Dim vEFITaskId As Variant
Dim oLockedForm As LISALockedForm

    On Error GoTo ErrLabel
    
    LockEFormInstances = False
    sMsg = ""
    
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
        'insert study key failed because study already locked - exit
        sMsg = GetLockDetails(oQueryServer, LISALockType.lkStudy, lStudyId)
        GoTo ErrLocked
    End If
    
    'Study now locked - try to lock subject
    sToken = InsertLock(oQueryServer, sUser, lStudyId, sSite, lSubjectId)
    If sToken = "" Then
        ' insert subject key failed because subject already locked - unlock study and exit
        RemoveLock oQueryServer, lStudyId
        sMsg = GetLockDetails(oQueryServer, LISALockType.lkSubject, lStudyId, sSite, lSubjectId)
        GoTo ErrLocked
    End If
    
    'subject now locked so unlock study
    RemoveLock oQueryServer, lStudyId
    
    'subject now locked - try to lock the eForm Instances
    For Each vEFITaskId In colEFormInstanceIds
        sToken = InsertLock(oQueryServer, sUser, lStudyId, sSite, lSubjectId, CLng(vEFITaskId))
             
        If sToken = "" Then
           'insert eform key failed because eform already locked
            sMsg = GetLockDetails(oQueryServer, LISALockType.lkEFormInstance, _
                                lStudyId, sSite, lSubjectId, CLng(vEFITaskId))
           ' Unlock what we've locked so far
            For Each oLockedForm In colLockedForms
                ' Assume the tokens are valid!
                Call RemoveLock(oQueryServer, lStudyId, sSite, lSubjectId, oLockedForm.EFormTaskId)
            Next
            ' Reset collection to empty
            Call CollectionRemoveAll(colLockedForms)
            Set colLockedForms = New Collection
            ' unlock subject and exit
            RemoveLock oQueryServer, lStudyId, sSite, lSubjectId
            GoTo ErrLocked
        Else
            ' Create locked form object and add to collection
            Set oLockedForm = New LISALockedForm
            oLockedForm.EFormTaskId = CLng(vEFITaskId)
            oLockedForm.LockToken = sToken
            colLockedForms.Add oLockedForm, FormKey(oLockedForm.EFormTaskId)
        End If

    Next
    
    Set oLockedForm = Nothing
    
    ' eform instances are now locked - remove subject key
    RemoveLock oQueryServer, lStudyId, sSite, lSubjectId
    
    Set oQueryServer = Nothing
    
    ' Finally report success
    LockEFormInstances = True

Exit Function
    
ErrLocked:
    Set oQueryServer = Nothing
    
Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|LISALockBS.LockEFormInstances"

End Function

'----------------------------------------------------------------------
Public Function UnlockEFormInstances(ByVal sCon As String, _
                                        ByVal lStudyId As Long, _
                                        ByVal sSite As String, _
                                        ByVal lSubjectId As Long, _
                                        colEFormLocks As Collection) As Boolean
'----------------------------------------------------------------------
' NCJ 19 Jul 04 - Created based on UnlockEFormInstance
' Unlock a load of eForm instances,
' as specified by collection of LISALockedForm objects
' Returns FALSE if any token is invalid (but unlocks what it can).
'----------------------------------------------------------------------
Dim oQueryServer As QueryServer
Dim oLockedForm As LISALockedForm
Dim bUnlockedOK As Boolean

    bUnlockedOK = True
    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sCon
    
    For Each oLockedForm In colEFormLocks
        With oLockedForm
            If Not ValidToken(oQueryServer, .LockToken, lStudyId, sSite, lSubjectId, .EFormTaskId) Then
                'invalid token
                bUnlockedOK = False
            Else
                'valid token so remove lock
                RemoveLock oQueryServer, lStudyId, sSite, lSubjectId, .EFormTaskId
            End If
        End With
    Next
    
    Set oLockedForm = Nothing
    Set oQueryServer = Nothing
    UnlockEFormInstances = bUnlockedOK

End Function

'----------------------------------------------------------------------
Private Function GetLockDetails(oQueryServer As QueryServer, _
                                ByVal enWhatsLocked As LISALockType, _
                                ByVal lStudyId As Long, _
                                Optional ByVal sSite As String = m_NO_SITE, _
                                Optional ByVal lSubjectId As Long = m_NO_SUBJECT, _
                                Optional ByVal lEFormInstanceId As Long = m_NO_EFORMINSTANCE) As String
'----------------------------------------------------------------------
' Returns suitable message for a specific lock.
'----------------------------------------------------------------------
Dim oQD As QueryDef
Dim vRow As Variant
Dim sUser As String

    Set oQD = New QueryDef
    oQD.InitSelect m_LOCK_TABLE, Array(m_COL_USER, m_COL_LOCKTIMESTAMP, m_COL_TOKEN), _
            Array(m_COL_STUDY, m_COL_SITE, m_COL_SUBJECT, m_COL_EFORMINSTANCE), _
            Array(lStudyId, sSite, lSubjectId, lEFormInstanceId)
    
    vRow = oQueryServer.SelectArray(oQD)
    If IsNull(vRow) Then
        sUser = "another user"
    Else
        sUser = vRow(0, 0)
    End If

    Select Case enWhatsLocked
    Case LISALockType.lkStudy
        GetLockDetails = "The study is currently being edited by " & sUser & "."
    Case LISALockType.lkSubject
        GetLockDetails = "The subject is currently being edited by " & sUser & "."
    Case LISALockType.lkEFormInstance
        GetLockDetails = "eForm is currently being edited by " & sUser & "."
    End Select
    
    Set oQD = Nothing
    
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

'----------------------------------------------------------------------
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

