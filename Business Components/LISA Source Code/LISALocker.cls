VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LISALocker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
' File: LISALocker.cls
' Copyright: InferMed Ltd 2004 All Rights Reserved
' Author: Nicky Johns, February 2004
' Purpose: EForm and Subject locking for LISA application
'------------------------------------------------------------------
' REVISIONS
' NCJ 24 Feb 04 - Initial development
' NCJ 18 Mar 04 - Further major development
' NCJ 26-30 Mar 04 - Testing and debugging
' NCJ 13 Apr 04 - Handle case where there is NO data in PH6_VS visit
' NCJ 27 Apr 04 - Changed GetLockTokenString to not use a StudySubject
' NCJ 24 May 04 - Handle invalid lock tokens in SetLockTokensString
' NCJ 15 July 04 - Made locking more efficient
' NCJ 19 July 04 - Use new LISALockBS class
'------------------------------------------------------------------

Option Explicit

Private mlStudyId As Long
Private msSite As String
Private mlPersonId As Long

Private msUserName As String

Private msConnectionString As String

Private msStudySiteSubject As String

' Collection of LISALockedForm objects, indexed by "FormKey"
Private mcolLockedForms As Collection

Private Const msPIPE = "|"

'------------------------------------------------------------------
Public Function GetLockTokenString(oUser As MACROUser, _
                                ByVal lStudyId As Long, _
                                ByVal sSite As String, _
                                ByVal lSubjId As Long, _
                                ByRef sMsg As String) As String
'------------------------------------------------------------------
' Lock all the eForms
' and return the LockTokens string
' NCJ 19 July 04 - Major code surgery
'------------------------------------------------------------------
Dim sLISAVisit As String
Dim sLockForms As String
Dim sCurForms As String
Dim vLockForms As Variant
Dim vLockForm As Variant
Dim i As Integer
Dim sVisitCode As String
Dim sFormCode As String
Dim bLockedOK As Boolean
Dim oVsAndFs As VisitsAndEForms

    On Error GoTo Errorlabel
    
    bLockedOK = True
    
    ' Store the subject details
    mlStudyId = lStudyId
    msSite = sSite
    mlPersonId = lSubjId
    
    msConnectionString = oUser.CurrentDBConString
    msUserName = oUser.UserName
    
    msStudySiteSubject = lStudyId & msPIPE _
                    & sSite & msPIPE _
                    & lSubjId
    
    ' Create new Visits and EForms info class
    Set oVsAndFs = New VisitsAndEForms
    Call oVsAndFs.Init(msConnectionString, lStudyId, sSite, lSubjId)
    
    ' Get the lists of eforms etc. from the Settings file
    Call GetLISALockList(sLISAVisit, sLockForms, sCurForms)
    
    ' Initialise EFI/token collection
    Set mcolLockedForms = New Collection
    
    ' Start with the LockForms list
    ' (comma separated items of form VisitCode-EFormCode)
    ' In this iteration we consider ALL visit cycles
    vLockForms = Split(sLockForms, ",")
    For i = 0 To UBound(vLockForms)
        vLockForm = Split(vLockForms(i), "-")
        sVisitCode = Trim(vLockForm(0))
        sFormCode = Trim(vLockForm(1))
        If Not LockVisitEForms(oVsAndFs, sVisitCode, sFormCode, sMsg) Then
            bLockedOK = False
            Exit For
        End If
    Next

    If bLockedOK Then
        ' Now do the Forms in the current LISA visit
        sVisitCode = sLISAVisit     ' Fixed visit this time
        bLockedOK = LockCurrentVisitEForms(oVsAndFs, sVisitCode, sCurForms, sMsg)
    End If
    
    If bLockedOK Then
        GetLockTokenString = msStudySiteSubject & LockTokensString
    Else
        GetLockTokenString = ""
    End If
    
    Call oVsAndFs.Terminate
    Set oVsAndFs = Nothing
    
Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|LISALocker.GetLockTokenString"

End Function

'----------------------------------------------------------------------
Public Function SetLockTokensString(ByVal sLockTokens As String, _
                            ByVal sDBCon As String, _
                            Optional ByRef oSubject As StudySubject = Nothing) As Boolean
'----------------------------------------------------------------------
' Ingest the previously created lock tokens for this subject
' sDBCon is the DB connection string
' If oSubject is given, returns FALSE if the lock tokens string doesn't match our subject,
' otherwise we don't check the subj. details
'----------------------------------------------------------------------
Dim vLocks As Variant
Dim i As Integer
Dim bOK As Boolean

    On Error GoTo Errorlabel
    
    ' Remember the DB connection
    msConnectionString = sDBCon
    
    Set mcolLockedForms = New Collection
    bOK = True
    If sLockTokens > "" Then
        ' sLockTokens is a pipe-separated string
        vLocks = Split(sLockTokens, msPIPE)
        ' The first three entries are the StudyId, Site and SubjId
        ' NCJ 26 Apr 04 - Check there are 3 entries
        If UBound(vLocks) >= 2 Then
            ' NCJ 24 May 04 - Trap funny lock token string...
            On Error GoTo LocksErr1
            mlStudyId = CLng(vLocks(0))
            msSite = vLocks(1)
            mlPersonId = CLng(vLocks(2))
            
            ' Check it matches our subject (only if known)
            If Not oSubject Is Nothing Then
                bOK = (mlStudyId = oSubject.StudyId) _
                    And (msSite = oSubject.Site) _
                    And (mlPersonId = oSubject.PersonID)
            End If
            
            If bOK Then
                i = 3
                On Error GoTo LocksErr2
                Do While i <= UBound(vLocks)
                    ' They come in EFITaskid|Token pairs
                    Call AddEFormLock(CLng(vLocks(i)), vLocks(i + 1))
                    i = i + 2
                Loop
            End If
        Else
            ' Wierd lock tokens string!
            bOK = False
        End If
    Else
        ' NCJ 26 Apr 04 - Flag empty lock tokens string
        bOK = False
    End If
    
    SetLockTokensString = bOK

Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|LISALocker.SetLockTokenString"

Exit Function
LocksErr1:
    ' There was some garbage in the first part of the lock tokens string
    SetLockTokensString = False

Exit Function
LocksErr2:
    ' There was some garbage in the actual lock tokens
    SetLockTokensString = False

End Function

'----------------------------------------------------------------------
Public Function UnlockSubject(ByVal sLockTokens As String, _
                            ByVal sDBCon As String, ByRef sMsg As String) As Long
'----------------------------------------------------------------------
' Unlock everything for this subject
' after setting up the relevant lock tokens and DB connection
' NCJ 24 May 04 - Added sMsg argument
'----------------------------------------------------------------------

    On Error GoTo Errorlabel
    
    If SetLockTokensString(sLockTokens, sDBCon) Then
        If UnlockAllEForms Then
            ' Return Success
            UnlockSubject = 0
        Else
            UnlockSubject = glINVALID_LOCKS
            sMsg = "One or more eForms could not be unlocked"
        End If
    Else
        ' Couldn't cope with lock tokens
        UnlockSubject = glINVALID_LOCKS
        sMsg = "Invalid lock tokens string"
    End If

Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|LISALocker.UnlockSubject"

End Function

'----------------------------------------------------------------------
Private Function LockCurrentVisitEForms(oVsAndFs As VisitsAndEForms, _
                    ByVal sVisitCode As String, ByVal sFormCodes As String, _
                    ByRef sErrMsg As String) As Boolean
'----------------------------------------------------------------------
' Lock all the eForms in the LAST (current) cycle of this VisitCode
' specified by sFormCodes (comma-separated list of quoted form codes)
' and store the tokens in our module-level collection.
' Returns FALSE if any locking failed.
'----------------------------------------------------------------------
Dim lVisitId As Long
Dim colEFITaskIds As Collection
Dim bLockedOK As Boolean

    On Error GoTo Errorlabel
    
    bLockedOK = True
    
    ' Get the visit Id
    lVisitId = oVsAndFs.VisitIdByCode(sVisitCode)
    If lVisitId > 0 Then
        Set colEFITaskIds = oVsAndFs.AllEFIsInLastVisit(lVisitId, sFormCodes)
        bLockedOK = LockEForms(colEFITaskIds, sErrMsg)
        If Not bLockedOK Then
            sErrMsg = "Unable to lock eForm in visit " & sVisitCode & ". " & sErrMsg
        End If
    End If
    
    LockCurrentVisitEForms = bLockedOK
    
    Call CollectionRemoveAll(colEFITaskIds)
    Set colEFITaskIds = Nothing
    
    Exit Function
    
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|LISALocker.LockCurrentVisitEForms"

End Function

'----------------------------------------------------------------------
Private Function LockVisitEForms(oVsAndFs As VisitsAndEForms, _
                    ByVal sVisitCode As String, ByVal sFormCode As String, _
                    ByRef sErrMsg As String) As Boolean
'----------------------------------------------------------------------
' Lock all the eForms matching this VisitCode:FormCode
' and store the tokens in our module-level collection
' If bLastVisitCycleOnly = TRUE, we ONLY consider the LATEST visit cycle (otherwise we do all)
' If any locking fails, we return FALSE
'----------------------------------------------------------------------
Dim lVisitId As Long
Dim lEformId As Long
Dim bLockedOK As Boolean
Dim colEFITaskIds As Collection

    On Error GoTo Errorlabel
    
    bLockedOK = True
    
    ' Get the visit/form Ids
    lVisitId = oVsAndFs.VisitIdByCode(sVisitCode)
    lEformId = oVsAndFs.EFormIdByCode(sFormCode)
    
    If (lVisitId > 0) And (lEformId > 0) Then
        Set colEFITaskIds = oVsAndFs.AllEFIsByVisitEForm(lVisitId, lEformId)
        bLockedOK = LockEForms(colEFITaskIds, sErrMsg)
        If Not bLockedOK Then
            sErrMsg = "Unable to lock eForm " & sVisitCode & ":" & sFormCode & ". " & sErrMsg
        End If
    Else
        ' Visit or eForm code doesn't exist! (Shall we ignore this?)
    End If
    
    LockVisitEForms = bLockedOK
    
    Call CollectionRemoveAll(colEFITaskIds)
    Set colEFITaskIds = Nothing
    
    Exit Function
    
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|LISALocker.LockVisitEForms"

End Function

'----------------------------------------------------------------------
Private Function LockEForms(colEFITaskIds As Collection, ByRef sErrMsg As String) As Boolean
'----------------------------------------------------------------------
' Lock a collection of EFITaskIDs
' Returns TRUE if all OK
' otherwise FALSE with sErrMsg an appropriate error message
'----------------------------------------------------------------------
Dim oLockBS As LISALockBS

    On Error GoTo Errorlabel
    
    Set oLockBS = New LISALockBS
    
    LockEForms = oLockBS.LockEFormInstances(msConnectionString, msUserName, _
                            mlStudyId, msSite, mlPersonId, colEFITaskIds, mcolLockedForms, sErrMsg)
    Set oLockBS = Nothing
    
Exit Function
    
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|LISALocker.LockEForms"

End Function

'---------------------------------------------------------------------
Private Sub AddEFormLock(ByVal lEFITaskId As Long, ByVal sToken As String)
'---------------------------------------------------------------------
' Add an eForm lock object to our module collection for this EFI and token
'---------------------------------------------------------------------
Dim oLockedForm As LISALockedForm
    
    Set oLockedForm = New LISALockedForm
    
    oLockedForm.EFormTaskId = lEFITaskId
    oLockedForm.LockToken = sToken
    mcolLockedForms.Add oLockedForm, FormKey(lEFITaskId)
    
    Set oLockedForm = Nothing

End Sub

'----------------------------------------------------------------------
Public Property Get EFILockToken(oEFI As EFormInstance) As String
'----------------------------------------------------------------------
' Retrieve the stored lock token for this eForm, if any
' Will be empty if lock token not found
'----------------------------------------------------------------------
Dim oLockedForm As LISALockedForm

    ' Allow item not to be found in collection
    On Error Resume Next
    Set oLockedForm = mcolLockedForms.Item(FormKey(oEFI.EFormTaskId))
    If Not oLockedForm Is Nothing Then
        EFILockToken = oLockedForm.LockToken
        Set oLockedForm = Nothing
    Else
        EFILockToken = ""
    End If
    
End Property

'----------------------------------------------------------------------
Private Function UnlockAllEForms() As Boolean
'----------------------------------------------------------------------
' Unlock all the eForms in our collection of locks
' and empty the collection
' ASSUME the subject details, DB connection string and locks collection already set up
' NCJ 24 May 04 - Added Boolean return value
' NCJ 19 July 04 - Use new LISALockBS class
'----------------------------------------------------------------------
Dim oLockBS As LISALockBS

    Set oLockBS = New LISALockBS
    UnlockAllEForms = oLockBS.UnlockEFormInstances(msConnectionString, _
                                                mlStudyId, msSite, mlPersonId, mcolLockedForms)
    Set oLockBS = Nothing
    
End Function

'----------------------------------------------------------------------
Private Function LockTokensString() As String
'----------------------------------------------------------------------
' Concatenate the EFITaskIds and Tokens in our collection of locks
' Returned string starts with a separator if there are any,
' but empty string if collection is empty
'----------------------------------------------------------------------
Dim oLockedForm As LISALockedForm
Dim sTokens As String

    sTokens = ""
    For Each oLockedForm In mcolLockedForms
        sTokens = sTokens & msPIPE _
                    & oLockedForm.EFormTaskId & msPIPE & oLockedForm.LockToken
    Next
    Set oLockedForm = Nothing
    LockTokensString = sTokens
    
End Function

'----------------------------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------------------------
    
    Set mcolLockedForms = New Collection

End Sub

'----------------------------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------------------------

    Call Terminate
    
End Sub

'----------------------------------------------------------------------
Public Sub Terminate()
'----------------------------------------------------------------------

    Call CollectionRemoveAll(mcolLockedForms)
    Set mcolLockedForms = Nothing
    
End Sub




