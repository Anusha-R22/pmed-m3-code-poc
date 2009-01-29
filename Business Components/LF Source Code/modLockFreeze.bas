Attribute VB_Name = "modLockFreeze"
'----------------------------------------------------------------------------------------'
'   File:       modLockFreeze.bas
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Nicky Johns, December 2002
'   Purpose:    Carries out Lock/Freeze operations for a subject in MACRO 3.0
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 4-19 Dec 02 - Initial development
'   NCJ 24 Dec 02 - Default is now to NOT set the Changed flag when locking/freezing
'           Added GetUnprocessedSiteMessages
'   NCJ 6 Jan 03 - Corrections to CanPerformLFAction (during code review with MLM)
'   NCJ 7 Jan 03 - Added GetSubjectRollbacks
'   NCJ 9 Jan 03 - Re-implemented Unfreeze (as a rollback of Freeze)
'   NCJ 14 Jan 03 - Changes to getting sets of messages
'----------------------------------------------------------------------------------------'

Option Explicit

'--------------------------------------------------------------
Public Function GetAllRollbacks(oDBCon As ADODB.Connection) As Collection
'--------------------------------------------------------------
' Get all unprocessed Rollback messages
' Returns collection of LFMessage objects ordered by Trial, Site, Subject
' and in reverse order of MessageID
' (See also GetSubjectRollBacks)
'--------------------------------------------------------------
Dim sSQL As String
Dim rsLFMsgs As ADODB.Recordset

    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM LFMessage WHERE " _
            & " MsgType = " & LFAction.lfaRollback _
            & " AND ProcessedStatus = " & LFProcessStatus.lfpUnProcessed _
            & " ORDER BY ClinicalTrialName, TrialSite, PersonId, RollbackMessageID DESC "
    Set rsLFMsgs = New ADODB.Recordset
    rsLFMsgs.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    ' Create collection of messages from the recordset
    Set GetAllRollbacks = MessagesFromRS(rsLFMsgs)

    rsLFMsgs.Close
    Set rsLFMsgs = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.GetAllRollbacks"
        
End Function

'--------------------------------------------------------------
Public Function GetSubjectRollbacks(oDBCon As ADODB.Connection, _
                    sStudyName As String, sSite As String, lSubjectId As Long) As Collection
'--------------------------------------------------------------
' Get all unprocessed Rollback messages for a particular subject
' Returns collection of LFMessage objects in reverse order of MessageID
' (See also GetAllRollBacks)
'--------------------------------------------------------------
Dim sSQL As String
Dim rsLFMsgs As ADODB.Recordset

    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM LFMessage WHERE " _
            & " ClinicalTrialName = '" & sStudyName & "'" _
            & " AND TrialSite = '" & sSite & "'" _
            & " AND PersonId = " & lSubjectId _
            & " AND MsgType = " & LFAction.lfaRollback _
            & " AND ProcessedStatus = " & LFProcessStatus.lfpUnProcessed _
            & " ORDER BY RollbackMessageID DESC "
    Set rsLFMsgs = New ADODB.Recordset
    rsLFMsgs.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    ' Create collection of messages from the recordset
    Set GetSubjectRollbacks = MessagesFromRS(rsLFMsgs)

    rsLFMsgs.Close
    Set rsLFMsgs = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.GetSubjectRollbacks"
        
End Function

'--------------------------------------------------------------
Private Function MessagesFromRS(rsLFMsgs As ADODB.Recordset) As Collection
'--------------------------------------------------------------
' Create a collection of LFMessages from the given recordset of messages
' Use string Sequence number as key (unique per LFMessage table)
'--------------------------------------------------------------
Dim colMsgs As Collection
Dim oLFMsg As LFMessage
    
    On Error GoTo ErrHandler
    
    Set colMsgs = New Collection
    
    If rsLFMsgs.RecordCount > 0 Then
        rsLFMsgs.MoveFirst
        Do While Not rsLFMsgs.EOF
            Set oLFMsg = New LFMessage
            Call oLFMsg.LoadFromRS(rsLFMsgs)
            colMsgs.Add oLFMsg, Str(oLFMsg.SequenceNo)
            rsLFMsgs.MoveNext
        Loop
    End If
    
    Set MessagesFromRS = colMsgs
    
    Set colMsgs = Nothing
    Set oLFMsg = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.MessagesFromRS"
        
End Function

'--------------------------------------------------------------
Public Function GetRollbackObjectStatus(oDBCon As ADODB.Connection, _
                                oLFObj As LFObject, lMsgSeqNo As Long) As Integer
'--------------------------------------------------------------
' Get the lock status of an object as it was before the LFMessage with the given Sequence No. was executed
' If lMsgSeqNo = 0, consider ALL messages
'--------------------------------------------------------------
Dim sSQL As String
Dim rsLFMsgs As ADODB.Recordset
Dim colLockUnlockMsgs As Collection
Dim oLFPrevMsg As LFMessage
Dim bFoundStatus As Boolean
Dim i As Long
Dim enRequiredStatus As LockStatus

    On Error GoTo ErrHandler
    
    ' Select "Lock" and "Unlock" processed LF messages for this subject
    ' with sequence no. before the message being rolled back
    With oLFObj
        sSQL = "SELECT * FROM LFMessage WHERE " _
            & " ClinicalTrialId = " & .StudyId _
            & " AND TrialSite = '" & .Site & "'" _
            & " AND PersonId = " & .SubjectId _
            & " AND ProcessedStatus = " & LFProcessStatus.lfpProcessed _
            & " AND (MsgType = " & LFAction.lfaLock & " OR MsgType = " & LFAction.lfaUnlock & ")"
        If lMsgSeqNo > 0 Then
            sSQL = sSQL & " AND SequenceNo < " & lMsgSeqNo
        End If
        sSQL = sSQL & " ORDER BY SequenceNo DESC "
    End With
    
    Set rsLFMsgs = New ADODB.Recordset
    rsLFMsgs.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    ' Create collection of messages from the recordset
    Set colLockUnlockMsgs = MessagesFromRS(rsLFMsgs)

    rsLFMsgs.Close
    Set rsLFMsgs = Nothing
    
    ' Now search for the "latest" one that's either
    '   Lock and it refers to object or a parent, OR
    '   Unlock and it refers to object or parent or descendant
    
    enRequiredStatus = LockStatus.lsUnlocked
    bFoundStatus = False
    i = 1
    Do While (bFoundStatus = False) And (i <= colLockUnlockMsgs.Count)
        Set oLFPrevMsg = colLockUnlockMsgs(i)
        If MsgRefersToObject(oLFPrevMsg, oLFObj) Then
            ' We've got the required status - it's either locked or unlocked
            If oLFPrevMsg.ActionType = lfaLock Then
                enRequiredStatus = LockStatus.lsLocked
            End If
            bFoundStatus = True
        Else
            ' Look at the next message
            i = i + 1
        End If
    Loop

    GetRollbackObjectStatus = enRequiredStatus

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.GetRollbackObjectStatus"
        
End Function

'--------------------------------------------------------------
Private Function MsgRefersToObject(oLFMsg As LFMessage, oLFObj As LFObject)
'--------------------------------------------------------------
' Does the given LF message affect the given object?
' Consider downwards effects only unless message is Unlock (which can have upwards effects too)
' Assume that Study, Site and Subject already match
' The message may be of type Rollback, and in this case we only consider downwards effects
'--------------------------------------------------------------
Dim bAffectsObject As Boolean

    On Error GoTo ErrHandler
    
    bAffectsObject = False
    
    ' Look at scope of message and see if it covers object downwards
    ' i.e. Subject messages affect everything
    '   Visit messages affect visit, eForms and questions
    '   eForm messages affect eForm and questions
    Select Case oLFMsg.Scope
    Case LFScope.lfscSubject
        ' If message scope is Subject, then it refers to every object!
        bAffectsObject = True
        
    Case LFScope.lfscVisit
        If SameVisit(oLFMsg, oLFObj) Then
            ' Object scope must be Visit, EForm or Question
            bAffectsObject = (oLFObj.Scope >= lfscVisit)
        End If
        
    Case LFScope.lfscEForm
        If SameEForm(oLFMsg, oLFObj) Then
            ' Object scope must be EForm or Question
            bAffectsObject = (oLFObj.Scope >= lfscEForm)
        End If
        
    Case LFScope.lfscQuestion
        If SameQuestion(oLFMsg, oLFObj) Then
            ' Object scope must be Question
            bAffectsObject = (oLFObj.Scope = lfscQuestion)
        End If
    
    End Select
        
    If Not bAffectsObject And oLFMsg.ActionType = lfaUnlock Then
        ' For unlock messages we need to consider upward effects
        ' i.e. question messages affect eform, visit and subject
        ' eForm messages affect visit and subject
        ' visit messages affect subject
        
        ' When considering upward effects, every message affects a subject
        bAffectsObject = (oLFObj.Scope = lfscSubject)
        
        If Not bAffectsObject Then
            ' The object is not a Subject
            Select Case oLFMsg.Scope
            Case LFScope.lfscQuestion
                ' Message is Unlock Question
                ' Is the object the question's eForm or visit?
                bAffectsObject = (oLFObj.Scope = lfscVisit) And SameVisit(oLFMsg, oLFObj) _
                              Or (oLFObj.Scope = lfscEForm) And SameEForm(oLFMsg, oLFObj)
                              
            Case LFScope.lfscEForm
                ' Message is Unlock EForm
                ' Is the object the eForm's visit?
                bAffectsObject = (oLFObj.Scope = lfscVisit) And SameVisit(oLFMsg, oLFObj)
                
            End Select
        End If
    End If
    
    MsgRefersToObject = bAffectsObject
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.MsgRefersToObject"
        
End Function

'--------------------------------------------------------------
Private Function SameVisit(oLFMsg As LFMessage, oLFObj As LFObject) As Boolean
'--------------------------------------------------------------
' Returns TRUE if message and object contain same Visit spec
'--------------------------------------------------------------

    SameVisit = (oLFMsg.VisitId = oLFObj.VisitId) And (oLFMsg.VisitCycle = oLFObj.VisitCycle)

End Function

'--------------------------------------------------------------
Private Function SameEForm(oLFMsg As LFMessage, oLFObj As LFObject) As Boolean
'--------------------------------------------------------------
' Returns TRUE if message and object contain same EForm spec (and same visit)
'--------------------------------------------------------------

    SameEForm = (oLFMsg.EFormId = oLFObj.EFormId) And (oLFMsg.EFormCycle = oLFObj.EFormCycle) _
                And SameVisit(oLFMsg, oLFObj)

End Function

'--------------------------------------------------------------
Private Function SameQuestion(oLFMsg As LFMessage, oLFObj As LFObject) As Boolean
'--------------------------------------------------------------
' Returns TRUE if message and object contain same Question spec (and same eForm and visit)
'--------------------------------------------------------------

    SameQuestion = (oLFMsg.ResponseId = oLFObj.ResponseId) And (oLFMsg.ResponseCycle = oLFObj.ResponseCycle) _
                And SameEForm(oLFMsg, oLFObj)

End Function

'--------------------------------------------------------------
Public Sub ExecuteRollback(oDBConn As ADODB.Connection, oLFMsg As LFMessage)
'--------------------------------------------------------------
' Rollback the given LFMessage by setting relevant lock statuses
' to what they were before this message was done
'--------------------------------------------------------------
Dim oLFObj As LFObject

    On Error GoTo ErrHandler
        
    ' Get the object to which the rollback refers
    Set oLFObj = oLFMsg.MessageObject

    Select Case oLFMsg.ActionType
    Case LFAction.lfaLock
        ' Set the top level object to Unlocked
        Call SetObjectLockStatus(oDBConn, oLFObj, LockStatus.lsUnlocked)
        ' Reset the lock status for each non-frozen child (and grandchildren etc.)
        Call ResetLFStatusOfNonFrozenChildren(oDBConn, oLFObj, oLFMsg)
        
    Case LFAction.lfaUnlock
        ' Set the top level object to Locked
        Call SetObjectLockStatus(oDBConn, oLFObj, LockStatus.lsLocked)
        ' Reset the lock status for each non-frozen child (and grandchildren etc.)
        Call ResetLFStatusOfNonFrozenChildren(oDBConn, oLFObj, oLFMsg)
        ' Reset the lock status of its parent(s)
        Call ResetLFStatusOfParents(oDBConn, oLFObj, oLFMsg)
        
    Case LFAction.lfaFreeze
        ' Set the top level object to what it was before it was Frozen
        Call SetObjectLockStatus(oDBConn, oLFObj, GetRollbackObjectStatus(oDBConn, oLFObj, oLFMsg.SequenceNo))
        ' Reset the freeze status for each child (and grandchildren etc.)
        Call ResetFreezeStatusOfChildren(oDBConn, oLFObj, oLFMsg.SequenceNo)
        
    Case LFAction.lfaUnfreeze
        ' Rolling back an unfreeze means setting everything back to frozen
        Call ReFreeze(oDBConn, oLFObj)
    
    End Select
    
    Set oLFObj = Nothing
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.ExecuteRollback"
        
End Sub

'-------------------------------------------------------------------------------------------
Public Sub UnfreezeObject(oDBConn As ADODB.Connection, oLFObj As LFObject)
'-------------------------------------------------------------------------------------------
' NCJ 9 Jan 03
' Unfreeze the object, by setting its status back to what it would have been had it not been frozen,
' and reset all its children too
'-------------------------------------------------------------------------------------------
        
    On Error GoTo ErrHandler
    
    ' Set the top level object to what it would be had it not been Frozen
    Call SetObjectLockStatus(oDBConn, oLFObj, GetRollbackObjectStatus(oDBConn, oLFObj, 0))
    ' Now reset the freeze status for each child (and grandchildren etc.)
    Call ResetFreezeStatusOfChildren(oDBConn, oLFObj, 0)

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.UnfreezeObject"
        
End Sub

'------------------------------------------------------------
Private Sub ResetFreezeStatusOfChildren(oDBConn As ADODB.Connection, _
                                        oLFObj As LFObject, lMsgSeqNo As Long)
'------------------------------------------------------------
' Reset the lock statuses for the children of the given object
' to what they were before the LF message with specified Seq No, as long as they weren't frozen,
' and recursively reset the lock statuses for all their children
' If lMsgSeqNo = 0, consider all messages
'------------------------------------------------------------
Dim oLFChildObject As LFObject

    On Error GoTo ErrHandler
    
    ' Reset the freeze status for each child
    For Each oLFChildObject In GetLFChildren(oDBConn, oLFObj, False)
        If ObjectWasFrozen(oDBConn, oLFChildObject, lMsgSeqNo) Then
            ' Leave it (and its children) as Frozen and do nothing more
        Else
            Call SetObjectLockStatus(oDBConn, oLFChildObject, _
                                    GetRollbackObjectStatus(oDBConn, oLFChildObject, lMsgSeqNo))
            ' Do its own children as well
            Call ResetFreezeStatusOfChildren(oDBConn, oLFChildObject, lMsgSeqNo)
        End If
    Next

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.ResetFreezeStatusOfChildren"
        
End Sub

'------------------------------------------------------------
Private Sub ResetLFStatusOfNonFrozenChildren(oDBConn As ADODB.Connection, _
                            oLFObj As LFObject, oLFMsg As LFMessage)
'------------------------------------------------------------
' Reset the lock statuses for the non-frozen children of the given object
' to what they were before the specified LF message,
' and recursively reset the lock statuses for all its non-frozen children
'------------------------------------------------------------
Dim oLFChildObject As LFObject
        
    On Error GoTo ErrHandler
    
    ' Reset the lock status for each non-frozen child
    For Each oLFChildObject In GetLFChildren(oDBConn, oLFObj, True)
        Call SetObjectLockStatus(oDBConn, oLFChildObject, GetRollbackObjectStatus(oDBConn, oLFChildObject, oLFMsg.SequenceNo))
        ' Do its own children as well
        Call ResetLFStatusOfNonFrozenChildren(oDBConn, oLFChildObject, oLFMsg)
    Next

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.ResetLFStatusOfNonFrozenChildren"
        
End Sub

'------------------------------------------------------------
Private Sub ResetLFStatusOfParents(oDBConn As ADODB.Connection, _
                            oLFObj As LFObject, oLFMsg As LFMessage)
'------------------------------------------------------------
' Reset the lock status for the parent of the given object
' to what it was before the specified LF message,
' and recursively reset the lock statuses for its parent
'------------------------------------------------------------
Dim oLFParentObject As LFObject
        
    ' Reset the lock status for parent
    Set oLFParentObject = GetLFParent(oDBConn, oLFObj)
    ' If no parent, we stop
    If oLFParentObject Is Nothing Then Exit Sub
    
    ' Set its correct LockStatus
    Call SetObjectLockStatus(oDBConn, oLFParentObject, GetRollbackObjectStatus(oDBConn, oLFParentObject, oLFMsg.SequenceNo))

    ' Now do its parent
    Call ResetLFStatusOfParents(oDBConn, oLFParentObject, oLFMsg)
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.ResetLFStatusOfParents"
        
End Sub

'------------------------------------------------------------
Private Sub SetObjectLockStatus(oDBConn As ADODB.Connection, _
                            oLFObj As LFObject, ByVal nLockSetting As LockStatus)
'------------------------------------------------------------
' Apply lock setting directly to given object ONLY
' (without setting any of its related objects)
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrHandler

    ' Set the lock status and specify the Subject
    sSQLSetLockWhere = SQLToSetLockStatus(nLockSetting, False) & oLFObj.SQLSubjectWhere

    Select Case oLFObj.Scope
    Case LFScope.lfscSubject
        sSQL = "UPDATE TrialSubject " & sSQLSetLockWhere
        
    Case LFScope.lfscVisit
        sSQL = "UPDATE VisitInstance " & sSQLSetLockWhere _
                & oLFObj.SQLAndVisitWhere
                
    Case LFScope.lfscEForm
        sSQL = "UPDATE CRFPageInstance " & sSQLSetLockWhere _
                & oLFObj.SQLAndVisitWhere _
                & oLFObj.SQLAndEFormWhere
                
    Case LFScope.lfscQuestion
        sSQL = "UPDATE DataItemResponse " & sSQLSetLockWhere _
                & oLFObj.SQLAndVisitWhere _
                & oLFObj.SQLAndEFormWhere _
                & oLFObj.SQLAndQuestionWhere
    End Select
    
    oDBConn.Execute sSQL
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modLockFreeze.SetObjectLockStatus"

End Sub

'------------------------------------------------------------
Private Function GetLFChildren(oDBConn As ADODB.Connection, _
                                oLFObj As LFObject, bNonFrozen As Boolean) As Collection
'------------------------------------------------------------
' Return collection of LFObjects representing the immediate children of this object
' If bNonFrozen = TRUE, only look for non-frozen children
' NCJ 2 Jan 03 - Select specific fields for greater efficiency
'------------------------------------------------------------
Dim sSQL As String
Dim colLFObjs As Collection
Dim sSQLSetLockWhere As String
Dim rsLFObjs As ADODB.Recordset
Dim enScope As LFScope

    On Error GoTo ErrHandler

    Set colLFObjs = New Collection
    
    ' Questions don't have children
    If oLFObj.Scope < LFScope.lfscQuestion Then
    
        sSQLSetLockWhere = oLFObj.SQLSubjectWhere
        If bNonFrozen Then
            ' Make sure we exclude Frozen ones
            sSQLSetLockWhere = sSQLSetLockWhere & " AND LockStatus <> " & LockStatus.lsFrozen
        End If
        
        Select Case oLFObj.Scope
        Case LFScope.lfscSubject
            ' Get the Visits for the Subject
            sSQL = "SELECT ClinicalTrialId, TrialSite, PersonId, " _
                    & " VisitId, VisitCycleNumber FROM VisitInstance " & sSQLSetLockWhere
            enScope = LFScope.lfscVisit
            
        Case LFScope.lfscVisit
            ' Get the EForms in the Visit
            sSQL = "SELECT ClinicalTrialId, TrialSite, PersonId, " _
                    & " VisitId, VisitCycleNumber, " _
                    & " CRFPageId, CRFPageCycleNumber FROM CRFPageInstance " & sSQLSetLockWhere _
                    & oLFObj.SQLAndVisitWhere
            enScope = LFScope.lfscEForm
                    
        Case LFScope.lfscEForm
            ' Get the Questions on the EForm
            sSQL = "SELECT ClinicalTrialId, TrialSite, PersonId, " _
                    & " VisitId, VisitCycleNumber, " _
                    & " CRFPageId, CRFPageCycleNumber, " _
                    & " ResponseTaskId, RepeatNumber, DataItemId FROM DataItemResponse " & sSQLSetLockWhere _
                    & oLFObj.SQLAndVisitWhere _
                    & oLFObj.SQLAndEFormWhere
            enScope = LFScope.lfscQuestion
        
        End Select
        
        ' Get the recordset and unwrap into a collection of objects
        Set rsLFObjs = New ADODB.Recordset
        rsLFObjs.Open sSQL, oDBConn, adOpenKeyset, adLockPessimistic, adCmdText
        Set colLFObjs = ObjectsFromRS(rsLFObjs, enScope)
        rsLFObjs.Close
        Set rsLFObjs = Nothing
    
    End If
    
    Set GetLFChildren = colLFObjs
    Set colLFObjs = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modLockFreeze.GetLFChildren"

End Function

'------------------------------------------------------------
Private Function GetLFParent(oDBConn As ADODB.Connection, _
                                        oLFObj As LFObject) As LFObject
'------------------------------------------------------------
' Return immediate parent of this object,
' or Nothing if no parent
'------------------------------------------------------------
Dim oLFParentObj As LFObject

    On Error GoTo ErrHandler

    Set oLFParentObj = Nothing
    
    ' Subjects don't have parents
    If oLFObj.Scope > LFScope.lfscSubject Then
    
        Set oLFParentObj = New LFObject
        
        ' Now see what our original object was
        With oLFObj
            Select Case .Scope
            Case LFScope.lfscVisit
                ' Get the Subject
                Call oLFParentObj.Init(LFScope.lfscSubject, _
                                .StudyId, .Site, .SubjectId)
            
            Case LFScope.lfscEForm
                ' Get the eForm's Visit
                Call oLFParentObj.Init(LFScope.lfscVisit, _
                                .StudyId, .Site, .SubjectId, _
                                .VisitId, .VisitCycle)
            
            Case LFScope.lfscQuestion
                ' Get the Question's eForm
                Call oLFParentObj.Init(LFScope.lfscEForm, _
                                .StudyId, .Site, .SubjectId, _
                                .VisitId, .VisitCycle, _
                                .EFormId, .EFormCycle)
                
            End Select
        End With
    End If
    
    Set GetLFParent = oLFParentObj     ' May be Nothing
    Set oLFParentObj = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modLockFreeze.GetLFParent"

End Function

'--------------------------------------------------------------
Private Function ObjectsFromRS(rsLFObjs As ADODB.Recordset, enScope As LFScope) As Collection
'--------------------------------------------------------------
' Create a collection of LFObjects with given scope (i.e. Subject, Visit, eForm or Question)
' from the given recordset of object details
'--------------------------------------------------------------
Dim colObjs As Collection
Dim oLFObj As LFObject
    
    On Error GoTo ErrHandler
    
    Set colObjs = New Collection
    
    If rsLFObjs.RecordCount > 0 Then
        rsLFObjs.MoveFirst
        Do While Not rsLFObjs.EOF
            Set oLFObj = New LFObject
            Call oLFObj.LoadFromRS(rsLFObjs, enScope)
            colObjs.Add oLFObj
            rsLFObjs.MoveNext
        Loop
    End If
    
    Set ObjectsFromRS = colObjs
    
    Set colObjs = Nothing
    Set oLFObj = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.ObjectsFromRS"
        
End Function

'--------------------------------------------------------------
Private Function ObjectWasFrozen(oDBConn As ADODB.Connection, oLFObj As LFObject, lMsgSeqNo As Long) As Boolean
'--------------------------------------------------------------
' Returns TRUE if given object was Frozen before the LFMessage with given SeqNo was carried out
' If lMsgSeqNo = 0, check ALL messages for the object
' Only look for messages for this exact object
'--------------------------------------------------------------
Dim sSQL As String
Dim rsLFMsgs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    With oLFObj
        ' Select "Freeze" and "Unfreeze" processed LF messages applied to this object
        ' with sequence no. before the message being rolled back
        ' NB Only consider messages on this exact object (not its parents)
        sSQL = "SELECT MsgType FROM LFMessage " _
            & oLFObj.SQLObjectWhere _
            & " AND Scope = " & .Scope _
            & " AND ProcessedStatus = " & LFProcessStatus.lfpProcessed _
            & " AND (MsgType = " & LFAction.lfaFreeze & " OR MsgType = " & LFAction.lfaUnfreeze & ")"
        If lMsgSeqNo > 0 Then
            sSQL = sSQL & " AND SequenceNo < " & lMsgSeqNo
        End If
        sSQL = sSQL & " ORDER BY SequenceNo DESC "
    End With
    
    Set rsLFMsgs = New ADODB.Recordset
    rsLFMsgs.Open sSQL, oDBConn, adOpenKeyset, adLockPessimistic, adCmdText
    If rsLFMsgs.RecordCount > 0 Then
        ' It was Frozen if the most recent message was of type Freeze
        ObjectWasFrozen = (rsLFMsgs!MsgType = LFAction.lfaFreeze)
    Else
        ' No such messages so it can't have been frozen
        ObjectWasFrozen = False
    End If
    
    rsLFMsgs.Close
    Set rsLFMsgs = Nothing
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.ObjectWasFrozen"
        
End Function

'------------------------------------------------------------
Private Function SQLToCheckLockStatus(ByVal nLockSetting As LockStatus) As String
'------------------------------------------------------------
' Returns SQL to check that item isn't frozen and doesn't have this LockSetting
'------------------------------------------------------------
    
    ' Make sure we only look at non-frozen items
    SQLToCheckLockStatus = " AND LockStatus <> " & nLockSetting _
                        & " AND LockStatus <> " & LockStatus.lsFrozen

End Function

'------------------------------------------------------------
Private Function SQLToSetLockStatus(ByVal nLockSetting As LockStatus, _
                    Optional bSetChanged As Boolean = False) As String
'------------------------------------------------------------
' SQL to set Lockstatus to given value
' and set Changed flag if bSetChanged = TRUE
'------------------------------------------------------------
Dim sSQL As String

    sSQL = " SET LockStatus = " & nLockSetting
    If bSetChanged Then
        sSQL = sSQL & ", Changed = " & Changed.Changed
    End If
    
    SQLToSetLockStatus = sSQL

End Function

'------------------------------------------------------------
Public Sub SetTrialSubjectLockStatus(oDBConn As ADODB.Connection, _
                    ByVal nLockSetting As LockStatus, oLFObj As LFObject, _
                    Optional bSetChanged As Boolean = False)
'------------------------------------------------------------
' Lock, Unlock or Freeze a trial subject
' i.e. apply setting to subject and to all its visits, forms and data items that are not frozen
' Don't change items that already have this status
' Set Changed flag if bSetChanged = TRUE
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrHandler

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus and the Changed flag
    sSQLSetLockWhere = SQLToSetLockStatus(nLockSetting, bSetChanged) & _
                    oLFObj.SQLSubjectWhere & _
                    SQLToCheckLockStatus(nLockSetting)

    ' Set LockStatus on Trial Subject
    sSQL = "UPDATE TrialSubject"
    sSQL = sSQL & sSQLSetLockWhere
    oDBConn.Execute sSQL
    
    ' Set LockStatus on all the visit instances
    sSQL = "UPDATE VisitInstance" & sSQLSetLockWhere
    oDBConn.Execute sSQL
    
    ' Set LockStatus on all the CRF Pages
    sSQL = "UPDATE CRFPageInstance" & sSQLSetLockWhere
    oDBConn.Execute sSQL
    
    ' Set LockStatus on all the Data Items
    sSQL = "UPDATE DataItemResponse" & sSQLSetLockWhere
    oDBConn.Execute sSQL

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modLockFreeze.SetTrialSubjectLockStatus"

End Sub

'------------------------------------------------------------
Public Sub SetVisitInstanceLockStatus(oDBConn As ADODB.Connection, _
                    ByVal nLockSetting As LockStatus, oLFObj As LFObject, _
                    Optional bSetChanged As Boolean = False)
'------------------------------------------------------------
' Lock, Unlock or Freeze a visit instance
' Apply setting to visit and to all its forms and data items that are not frozen
' Don't change items that already have this status
' Set Changed flag if bSetChanged = TRUE
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrHandler

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus and the Changed flag
    sSQLSetLockWhere = SQLToSetLockStatus(nLockSetting, bSetChanged) & _
            oLFObj.SQLSubjectWhere & _
            oLFObj.SQLAndVisitWhere & _
            SQLToCheckLockStatus(nLockSetting)
            
    ' The visit instance
    sSQL = "UPDATE VisitInstance" & sSQLSetLockWhere
    oDBConn.Execute sSQL
    
    ' All the CRF Pages
    sSQL = "UPDATE CRFPageInstance" & sSQLSetLockWhere
    oDBConn.Execute sSQL
    
    ' All the Data Items
    sSQL = "UPDATE DataItemResponse" & sSQLSetLockWhere
    oDBConn.Execute sSQL
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modLockFreeze.SetVisitInstanceLockStatus"

End Sub

'------------------------------------------------------------
Public Sub SetCRFPageInstanceLockStatus(oDBConn As ADODB.Connection, _
                    ByVal nLockSetting As LockStatus, oLFObj As LFObject, _
                    Optional bSetChanged As Boolean = False)
'------------------------------------------------------------
' Lock, Unlock or Freeze a CRF page
' Apply setting to page and to all its data items that are not frozen
' Don't change items that already have this status
' Set Changed flag if bSetChanged = TRUE
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrHandler

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus and the Changed flag
    sSQLSetLockWhere = SQLToSetLockStatus(nLockSetting, bSetChanged) & _
            oLFObj.SQLSubjectWhere & _
            oLFObj.SQLAndVisitWhere & _
            oLFObj.SQLAndEFormWhere & _
            SQLToCheckLockStatus(nLockSetting)

    ' Lock the CRF Page
    sSQL = "UPDATE CRFPageInstance" & sSQLSetLockWhere
    oDBConn.Execute sSQL
    
    ' Lock all the Data Items
    sSQL = "UPDATE DataItemResponse" & sSQLSetLockWhere
    oDBConn.Execute sSQL

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modLockFreeze.SetCRFPageInstanceLockStatus"

End Sub

'------------------------------------------------------------
Public Sub SetDataItemLockStatus(oDBConn As ADODB.Connection, _
                    ByVal nLockSetting As LockStatus, oLFObj As LFObject, _
                    Optional bSetChanged As Boolean = False)
'------------------------------------------------------------
' Lock, Freeze or Unlock a data item
' Don't change items that already have this status
' Set Changed flag if bSetChanged = TRUE
'------------------------------------------------------------
Dim sSQL  As String

    On Error GoTo ErrHandler

    ' Lock the Data Item - set the LockStatus, and the Changed flag if required
    ' Don't need VisitWhere and EFormWhere because ResponseTaskId uniquely defines response
    sSQL = "UPDATE DataItemResponse " & _
            SQLToSetLockStatus(nLockSetting, bSetChanged) & _
            oLFObj.SQLSubjectWhere & _
            oLFObj.SQLAndQuestionWhere & _
            SQLToCheckLockStatus(nLockSetting)

    oDBConn.Execute sSQL
            
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modLockFreeze.SetDataItemLockStatus"

End Sub

'------------------------------------------------------------
Private Function ValidStatusTransition(nOldLockSetting As LockStatus, _
                                    enAction As LFAction) As Boolean
'------------------------------------------------------------
' Is it valid to apply action to primary object with Old LockStatus?
' (We assume that if not, action is dependent on previously refused LFmessage)
'------------------------------------------------------------

    Select Case nOldLockSetting
    Case LockStatus.lsFrozen
        ' Can only unfreeze
        ValidStatusTransition = (enAction = lfaUnfreeze)
    Case LockStatus.lsLocked
        ' Can freeze or unlock
        ValidStatusTransition = ((enAction = lfaFreeze) Or (enAction = lfaUnlock))
    Case LockStatus.lsUnlocked
        ' Can freeze or lock
        ValidStatusTransition = ((enAction = lfaFreeze) Or (enAction = lfaLock))
    End Select

End Function

'----------------------------------------------------------------------------------------'
Public Function CanPerformLFAction(oDBConn As ADODB.Connection, oLFServerMsg As LFMessage) As Boolean
'----------------------------------------------------------------------------------------'
' Returns TRUE if message's action can be done on message's object, i.e. data hasn't changed
' and object is in a suitable state
' NCJ 2 Jan 02 - Added consideration of existing site LF messages
'----------------------------------------------------------------------------------------'
Dim sSQL  As String
Dim rsTemp As ADODB.Recordset
Dim bCanDoIt As Boolean
Dim colLFMsgs As Collection
Dim oLFSiteMsg As LFMessage
Dim oLFObj As LFObject

    On Error GoTo ErrHandler

    ' We don't consider rollback messages here
    If oLFServerMsg.ActionType = LFAction.lfaRollback Then
        CanPerformLFAction = False
        Exit Function
    End If
    
    bCanDoIt = True
    
    Set oLFObj = oLFServerMsg.MessageObject
    
    ' Pick up any Changed flags from DataItemResponse
    sSQL = "SELECT COUNT(*) FROM DataItemResponse " & _
                oLFObj.SQLObjectWhere _
                & " AND Changed = " & Changed.Changed
     
    ' See if there were any Changed flags
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, oDBConn, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTemp.Fields(0).Value > 0 Then
        ' Something changed so not OK
        bCanDoIt = False
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    
    If bCanDoIt Then
        ' Finally make sure there wasn't a Site LFMessage or Rollback message
        ' which affects this object
        ' i.e. we'll refuse Server messages within the scope of Site messages or other already refused Server messages
        Set colLFMsgs = GetSubjectSiteLFMsgs(oDBConn, oLFServerMsg.StudyId, oLFServerMsg.Site, oLFServerMsg.SubjectId)
        If colLFMsgs.Count > 0 Then
            ' If the incoming message refers to the Subject, we can't do it
            If oLFServerMsg.Scope = LFScope.lfscSubject Then
                bCanDoIt = False
            Else
                ' See if any site messages overlap the server's message
                For Each oLFSiteMsg In colLFMsgs
                    ' See if either message affects the other
                    If MessagesInteract(oLFSiteMsg, oLFServerMsg) Then
                        bCanDoIt = False
                        Exit For
                    End If
                Next
            End If
        End If
    End If
    
    CanPerformLFAction = bCanDoIt
    
    Set oLFObj = Nothing
            
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modLockFreeze.CanPerformLFAction"

End Function

'----------------------------------------------------------------------------------------'
Private Function MessagesInteract(oLFMsg1 As LFMessage, oLFMsg2 As LFMessage) As Boolean
'----------------------------------------------------------------------------------------'
' Do these messages "interact", i.e. do their scopes overlap?
'----------------------------------------------------------------------------------------'
Dim bScopeOverlaps As Boolean

    ' See if the first message affects the object of the second
    bScopeOverlaps = MsgRefersToObject(oLFMsg1, oLFMsg2.MessageObject)
    
    ' If not, see if the second message affects the object of the first
    If Not bScopeOverlaps Then
        bScopeOverlaps = MsgRefersToObject(oLFMsg2, oLFMsg1.MessageObject)
    End If

    MessagesInteract = bScopeOverlaps

End Function

'----------------------------------------------------------------------------------------'
Private Sub ReFreeze(oDBConn As ADODB.Connection, oLFObj As LFObject)
'----------------------------------------------------------------------------------------'
' Refreeze the given object (as part of a Rollback Unfreeze)
' Set everything to Frozen but don't set the Changed flag
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    Select Case oLFObj.Scope
    Case LFScope.lfscQuestion
        Call SetDataItemLockStatus(oDBConn, LockStatus.lsFrozen, oLFObj, False)
        
    Case LFScope.lfscEForm
        Call SetCRFPageInstanceLockStatus(oDBConn, LockStatus.lsFrozen, oLFObj, False)
        
    Case LFScope.lfscVisit
        Call SetVisitInstanceLockStatus(oDBConn, LockStatus.lsFrozen, oLFObj, False)
        
    Case LFScope.lfscSubject
        Call SetTrialSubjectLockStatus(oDBConn, LockStatus.lsFrozen, oLFObj, False)

    End Select
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.ReFreeze"

End Sub

'----------------------------------------------------------------------------------------'
Public Function GetLFMessagesToTransfer(oDBCon As ADODB.Connection, _
                            ByVal nSource As Integer, ByVal sSite As String) As Collection
'----------------------------------------------------------------------------------------'
' Returns all unsent messages (SentTimeStamp = 0) for Site from given Source,
' ordered by Study, SubjectId, MessageId,
' as a collection of LFMessage objects
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsLFMsgs As ADODB.Recordset
Dim colLFMsgs As Collection

    On Error GoTo ErrHandler
    
    ' Select unsent LF messages for this Source
    sSQL = "SELECT * FROM LFMessage WHERE " _
        & " Source = " & nSource _
        & " AND TrialSite = '" & sSite & "'" _
        & " AND SentTimeStamp = 0 " _
        & " ORDER BY ClinicalTrialId, PersonId, MessageId "
    
    Set rsLFMsgs = New ADODB.Recordset
    rsLFMsgs.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    ' Create collection of messages from the recordset
    Set colLFMsgs = MessagesFromRS(rsLFMsgs)

    rsLFMsgs.Close
    Set rsLFMsgs = Nothing
    
    Set GetLFMessagesToTransfer = colLFMsgs
    
    Set colLFMsgs = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.GetLFMessagesToTransfer"

End Function

'----------------------------------------------------------------------------------------'
Public Function GetScopeText(enScope As LFScope) As String
'----------------------------------------------------------------------------------------'
' Scope as a string
'----------------------------------------------------------------------------------------'

    Select Case enScope
    Case LFScope.lfscQuestion
        GetScopeText = "Question"
    Case LFScope.lfscEForm
        GetScopeText = "EForm"
    Case LFScope.lfscVisit
        GetScopeText = "Visit"
    Case LFScope.lfscSubject
        GetScopeText = "Subject"
    End Select

End Function

'----------------------------------------------------------------------------------------'
Public Function GetActionText(enAction As LFAction) As String
'----------------------------------------------------------------------------------------'
' LF Action as a string
'----------------------------------------------------------------------------------------'

    Select Case enAction
    Case LFAction.lfaFreeze
        GetActionText = "Freeze"
    Case LFAction.lfaLock
        GetActionText = "Lock"
    Case LFAction.lfaUnfreeze
        GetActionText = "Unfreeze"
    Case LFAction.lfaUnlock
        GetActionText = "Unlock"
    Case LFAction.lfaRollback
        GetActionText = "Rollback"
    End Select

End Function

'----------------------------------------------------------------------------------------'
Public Function GetProcessStatusText(enStatus As LFProcessStatus) As String
'----------------------------------------------------------------------------------------'
' LF Processed Status as a string
'----------------------------------------------------------------------------------------'

    Select Case enStatus
    Case LFProcessStatus.lfpProcessed
        GetProcessStatusText = "Processed"
    Case LFProcessStatus.lfpRefused
        GetProcessStatusText = "Refused"
    Case LFProcessStatus.lfpUnProcessed
        GetProcessStatusText = "Unprocessed"
    End Select

End Function

'------------------------------------------------------------
Public Sub UnlockTrialSubject(oDBConn As ADODB.Connection, oLFObj As LFObject)
'------------------------------------------------------------
' Unlock the subject
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrHandler

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus, but only if it is locked (not frozen or unlocked)
    sSQLSetLockWhere = SQLToSetLockStatus(LockStatus.lsUnlocked) & _
            oLFObj.SQLSubjectWhere & _
            " AND LockStatus = " & LockStatus.lsLocked

    sSQL = "UPDATE TrialSubject" & sSQLSetLockWhere
    
    oDBConn.Execute sSQL
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modLockFreeze.UnlockTrialSubject"

End Sub

'------------------------------------------------------------
Public Sub UnlockVisitInstance(oDBConn As ADODB.Connection, oLFObj As LFObject)
'------------------------------------------------------------
' Unlock a Visit Instance if not frozen
' and ALSO unlock the Trial Subject
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrHandler

    ' Set the LockStatus and the Changed flag, but only if it is locked (not frozen or unlocked)
    sSQLSetLockWhere = SQLToSetLockStatus(LockStatus.lsUnlocked) & _
            oLFObj.SQLSubjectWhere & _
            oLFObj.SQLAndVisitWhere & _
            " AND LockStatus = " & LockStatus.lsLocked

    ' Lock the visit instance
    sSQL = "UPDATE VisitInstance" & sSQLSetLockWhere
    
    oDBConn.Execute sSQL

    ' Make sure the Trial Subject is also unlocked
    Call UnlockTrialSubject(oDBConn, oLFObj)
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modLockFreeze.UnlockVisitInstance"

End Sub

'------------------------------------------------------------
Public Sub UnlockCRFPageInstance(oDBConn As ADODB.Connection, oLFObj As LFObject)
'------------------------------------------------------------
' Unlock a CRF Page Instance if not frozen
' and ALSO unlock its Visit and Trial Subject
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrHandler

    ' Set the LockStatus and the Changed flag, but only if it is locked (not frozen or unlocked)
    sSQLSetLockWhere = SQLToSetLockStatus(LockStatus.lsUnlocked) & _
            oLFObj.SQLSubjectWhere & _
            oLFObj.SQLAndVisitWhere & _
            oLFObj.SQLAndEFormWhere & _
            " AND LockStatus = " & LockStatus.lsLocked

    ' Lock the CRF Page
    sSQL = "UPDATE CRFPageInstance" & sSQLSetLockWhere
    oDBConn.Execute sSQL

    ' Make sure the Visit and Trial subject are also unlocked
    Call UnlockVisitInstance(oDBConn, oLFObj)

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modLockFreeze.UnlockCRFPageInstance"

End Sub

'--------------------------------------------------------------
Public Function GetUnprocessedSiteMessages(oDBCon As ADODB.Connection, _
            sStudyName As String, sSite As String, lSubjectId As Long, lLastLFMessageId As Long) As Collection
'--------------------------------------------------------------
' Get all unprocessed site messages for the given Subject (excluding Rollbacks)
' with MessageId <= lLastMessageId
' Returns collection of LFMessage objects, ordered by MessageId
'--------------------------------------------------------------
Dim sSQL As String
Dim rsLFMsgs As ADODB.Recordset

    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM LFMessage WHERE " _
        & " ClinicalTrialName = '" & sStudyName & "'" _
        & " AND TrialSite = '" & sSite & "'" _
        & " AND PersonId = " & lSubjectId _
        & " AND Source = " & TypeOfInstallation.RemoteSite _
        & " AND MessageId <= " & lLastLFMessageId _
        & " AND ProcessedStatus = " & LFProcessStatus.lfpUnProcessed _
        & " AND MsgType <> " & LFAction.lfaRollback _
        & " ORDER BY MessageID "
    Set rsLFMsgs = New ADODB.Recordset
    rsLFMsgs.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    ' Create collection of messages from the recordset
    Set GetUnprocessedSiteMessages = MessagesFromRS(rsLFMsgs)

    rsLFMsgs.Close
    Set rsLFMsgs = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.GetUnprocessedSiteMessages"
        
End Function

'--------------------------------------------------------------
Private Function GetSubjectSiteLFMsgs(oDBCon As ADODB.Connection, _
                lStudyId As Long, sSite As String, lSubjectId As Long) As Collection
'--------------------------------------------------------------
' Get the unsent Site LFAs, including Rollbacks, for a subject
' as a collection of LFMessages
'--------------------------------------------------------------
Dim sSQL As String
Dim rsLFMsgs As ADODB.Recordset

    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM LFMessage WHERE " _
        & " ClinicalTrialId = " & lStudyId _
        & " AND TrialSite = '" & sSite & "'" _
        & " AND PersonId = " & lSubjectId _
        & " AND Source = " & TypeOfInstallation.RemoteSite _
        & " AND SentTimeStamp = 0 " _
        & " ORDER BY MessageID "
    Set rsLFMsgs = New ADODB.Recordset
    rsLFMsgs.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    ' Create collection of messages from the recordset
    Set GetSubjectSiteLFMsgs = MessagesFromRS(rsLFMsgs)

    rsLFMsgs.Close
    Set rsLFMsgs = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.GetSubjectSiteLFMsgs"

End Function

'--------------------------------------------------------------
Public Function GetProcessableLFMsgs(oDBCon As ADODB.Connection, colIgnoreSubjects As Collection) As Collection
'--------------------------------------------------------------
' Get the unprocessed Site LFAs, excluding Rollbacks, as a collection of LFMessages
' for any subjects that aren't in colIgnoreSubjects
' colIgnoreSubjects is a collection with items keyed by "ClinicalTrialName|TrialSite|PersonId"
'--------------------------------------------------------------
Dim sSQL As String
Dim rsLFMsgs As ADODB.Recordset
Dim colLFMsgs As Collection
Dim oLFMsg As LFMessage
Dim sKey As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM LFMessage WHERE " _
        & " Source = " & TypeOfInstallation.RemoteSite _
        & " AND ProcessedStatus = " & LFProcessStatus.lfpUnProcessed _
        & " AND MsgType <> " & LFAction.lfaRollback _
        & " ORDER BY ClinicalTrialName, TrialSite, PersonId, MessageID "
    Set rsLFMsgs = New ADODB.Recordset
    rsLFMsgs.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    Set colLFMsgs = New Collection
    ' Create collection of messages from the recordset
    If rsLFMsgs.RecordCount > 0 Then
        rsLFMsgs.MoveFirst
        Do While Not rsLFMsgs.EOF
            sKey = rsLFMsgs!ClinicalTrialName & "|" & rsLFMsgs!TrialSite & "|" & rsLFMsgs!PersonId
            If CollectionMember(colIgnoreSubjects, sKey, False) Then
                ' We ignore this one
            Else
                Set oLFMsg = New LFMessage
                Call oLFMsg.LoadFromRS(rsLFMsgs)
                colLFMsgs.Add oLFMsg
            End If
            rsLFMsgs.MoveNext
        Loop
    End If
    
    rsLFMsgs.Close
    Set rsLFMsgs = Nothing
    
    Set GetProcessableLFMsgs = colLFMsgs

    Set oLFMsg = Nothing
    Set colLFMsgs = Nothing
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modLockFreeze.GetProcessableLFMsgs"

End Function


