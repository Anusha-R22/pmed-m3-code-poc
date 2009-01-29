VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIFreezer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: MIFreezer.cls
' Copyright InferMed 2002, All Rights Reserved
' Author: Nicky Johns, InferMed, Nov 2002
' Purpose: Freezer class for MACRO 3.0, to perform Locking and Freezing operations
'----------------------------------------------------
' REVISIONS
'   NCJ 20-21 Nov 02 - Initial Development
'
'----------------------------------------------------

Option Explicit

' Freezer statuses
Public Enum eFreezerMIMStatus
    fsUnlocked = 0
    fsLocked = 5
    fsFrozen = 6
End Enum

Private mcolMIMsgStates As Collection
Private mcolMIMessages As Collection

' connection string
Private msCon As String

' reference to current message
Private moCurrentMsg As MIMsg

' Whether freezing overrides locking or not
Private mbFreezeOverridesLock As Boolean

' The connection for the Subject data tables update
Private moADODBConnection As ADODB.Connection

Private mlStudyId As Long

'----------------------------------------------------------------------------------------'
Public Sub Load(sCon As String, lObjectId As Long, nObjectSource As Integer, sSite As String)
'----------------------------------------------------------------------------------------'
'Load a set of Freezer MIMessages according to site, object source and object id.
'----------------------------------------------------------------------------------------'
Dim oMsg As MIMsg
Dim oDS As MIMsgDS
Dim colSerStates As Collection
Dim oState As MIMsgState
Dim i As Long
   
    On Error GoTo ErrLabel
    
    Set oDS = New MIMsgDS
    msCon = sCon
    Set mcolMIMsgStates = CollectionDeSerialise(oDS.Load(sCon, lObjectId, nObjectSource, sSite, False))
    Set oDS = Nothing
    
    Set mcolMIMessages = New Collection
    For i = 1 To mcolMIMsgStates.Count
        Set oState = mcolMIMsgStates(i)
        Set oMsg = New MIMsg
        oMsg.SetState oState
        mcolMIMessages.Add oMsg
        If oMsg.History = mimhCurrent Then
            Set moCurrentMsg = oMsg
        End If
    Next
  
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIFreezer.Load"
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Save()
'----------------------------------------------------------------------------------------'
'Save the current set of Freezer MIMessages
'----------------------------------------------------------------------------------------'
Dim oDS As MIMsgDS
Dim colSerStates As Collection
Dim oState As MIMsgState
Dim i As Long
Dim sUpdated
Dim oMsg As MIMsg

    On Error GoTo ErrLabel
    
    Set oDS = New MIMsgDS
    sUpdated = oDS.Save(msCon, CollectionSerialise(mcolMIMsgStates))
    Set mcolMIMsgStates = CollectionDeSerialise(sUpdated)
    Set oDS = Nothing
    
    Set mcolMIMessages = New Collection
    For i = 1 To mcolMIMsgStates.Count
        Set oState = mcolMIMsgStates(i)
        Set oMsg = New MIMsg
        oMsg.SetState oState
        mcolMIMessages.Add oMsg
        If oMsg.History = mimhCurrent Then
            Set moCurrentMsg = oMsg
        End If
    Next

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIFreezer.Save"

End Sub

'----------------------------------------------------------------------------------------'
Public Function Raise(sCon As String, _
                    sText As String, nStatus As eFreezerMIMStatus, _
                    sUserName As String, sUserNameFull As String, nSource As MIMsgSource, _
                    dblTimestamp As Double, nTimezoneOffset As Integer, _
                    nScope As MIMsgScope, _
                    sStudyName As String, sSite As String, lSubjectId As Long, _
                    lVisitId As Long, nVisitCycle As Integer, _
                    lEFormTaskId As Long, _
                    lResponseTaskId As Long, nResponseCycle As Integer, _
                    lEFormId As Long, nEFormCycle As Integer, _
                    lQuestionId As Long, sDataUsername As String, _
                    dblResponseTimeStamp As Double, sResponseValue As String) As Integer
'----------------------------------------------------------------------------------------'
'Raise a Freezer (create first MIMessage row).
'Raises mimeInvalidForThisStatus error if already a current row.
'Function returns index in messages collection of create message row.
' Depending on Scope, some arguments may be 0
'----------------------------------------------------------------------------------------'
Dim oMsg As MIMsg
Dim oMsgState As MIMsgState
Dim colSerStates As Collection
Dim oState As MIMsgState

    On Error GoTo ErrLabel
    
    If Not mcolMIMessages Is Nothing Then     'already raised
        Err.Raise MIMsgErrors.mimeInvalidForThisStatus, , "Operation invalid for current status"
    End If

    msCon = sCon
    Set mcolMIMsgStates = New Collection
    Set mcolMIMessages = New Collection
    Raise = Add(sText, nStatus, _
                sUserName, sUserNameFull, nSource, _
                dblTimestamp, nTimezoneOffset, _
                nScope, sStudyName, sSite, lSubjectId, _
                lVisitId, nVisitCycle, lEFormTaskId, _
                lResponseTaskId, nResponseCycle, dblResponseTimeStamp, sResponseValue, _
                lEFormId, nEFormCycle, lQuestionId, sDataUsername)
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIFreezer.Raise"

End Function

'----------------------------------------------------------------------------------------'
Public Function UnFreeze(sText As String, sUserName As String, sUserNameFull As String, _
                        nSource As MIMsgSource, _
                        dblTimestamp As Double, nTimezoneOffset As Integer, _
                        Optional dblResponseTimeStamp As Double = 0, _
                        Optional sResponseValue As String = "") As Integer
'----------------------------------------------------------------------------------------'
' Unfreeze a frozen Freezer
' Need to determine the target status (based on the object's parent status)
' Returns index in Messages collection of newly created message row.
'----------------------------------------------------------------------------------------'
Dim enStatus As eFreezerMIMStatus

    On Error GoTo ErrLabel
    
    ' We can't unfreeze if we're not frozen
    If Me.CurrentStatus <> fsFrozen Then
        Err.Raise MIMsgErrors.mimeInvalidForThisStatus, , "Operation invalid for current status"
    End If
    
    ' Get the target status
    enStatus = GetUnfrozenLockStatus(Scope, StudyId, Site, SubjectId, VisitId, VisitCycle, EFormTaskId)
    
    UnFreeze = ChangeStatus(enStatus, sText, sUserName, sUserNameFull, nSource, _
                        dblTimestamp, nTimezoneOffset, dblResponseTimeStamp, sResponseValue)
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIFreezer.UnFreeze"
    
End Function

'----------------------------------------------------------------------------------------'
Public Function ChangeStatus(enNewStatus As eFreezerMIMStatus, _
                        sText As String, sUserName As String, sUserNameFull As String, _
                        nSource As MIMsgSource, _
                        dblTimestamp As Double, nTimezoneOffset As Integer, _
                        Optional dblResponseTimeStamp As Double = 0, _
                        Optional sResponseValue As String = "") As Integer
'----------------------------------------------------------------------------------------'
' Change the status of this Freezer
' Raises mimeInvalidForThisStatus error if enNewStatus not allowed for current status.
' Returns index in Messages collection of newly created message row.
'----------------------------------------------------------------------------------------'
    
    On Error GoTo ErrLabel
    
    If Not ValidStatusTransition(enNewStatus) Then
        Err.Raise MIMsgErrors.mimeInvalidForThisStatus, , "Operation invalid for current status"
    End If
    
    ChangeStatus = Add(sText, enNewStatus, _
                sUserName, sUserNameFull, nSource, _
                dblTimestamp, nTimezoneOffset, _
                Scope, StudyName, Site, SubjectId, VisitId, VisitCycle, EFormTaskId, _
                ResponseTaskId, ResponseCycle, dblResponseTimeStamp, sResponseValue, _
                EFormId, EFormCycle, QuestionId, DataUserName, _
                FreezerID, FreezerSource)

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIFreezer.ChangeStatus"

End Function

'----------------------------------------------------------------------------------------'
Private Function Add(sText As String, nStatus As eFreezerMIMStatus, _
                    sUserName As String, sUserNameFull As String, nSource As MIMsgSource, _
                    dblTimestamp As Double, nTimezoneOffset As Integer, _
                    nScope As MIMsgScope, _
                    sStudyName As String, sSite As String, lSubjectId As Long, _
                    lVisitId As Long, nVisitCycle As Integer, _
                    lEFormTaskId As Long, _
                    lResponseTaskId As Long, nResponseCycle As Integer, _
                    dblResponseTimeStamp As Double, sResponseValue As String, _
                    lEFormId As Long, nEFormCycle As Integer, _
                    lQuestionId As Long, sDataUsername As String, _
                    Optional lObjectId As Long = -1, Optional nObjectSource As Integer = -1) As Integer
'----------------------------------------------------------------------------------------'
' General purpose routine to add a MIMessage row to a Freezer.
' (Copied from SDV class)
'----------------------------------------------------------------------------------------'
Dim oMsg As MIMsg
Dim oMsgState As MIMsgState

    On Error GoTo ErrLabel
    
    Set oMsgState = New MIMsgState
    mcolMIMsgStates.Add oMsgState
    Set oMsg = New MIMsg
    oMsg.SetState oMsgState
    Set oMsgState = Nothing
    
    oMsg.Init sUserName, sUserNameFull, nSource, MIMsgType.mimtFreezer, nStatus, _
                    nScope, sStudyName, sSite, lSubjectId, _
                    lVisitId, nVisitCycle, lEFormTaskId, _
                    lResponseTaskId, nResponseCycle, _
                    dblResponseTimeStamp, sResponseValue, _
                    dblTimestamp, nTimezoneOffset, _
                    lEFormId, nEFormCycle, lQuestionId, sDataUsername, _
                    lObjectId, nObjectSource
                            
    oMsg.Text = sText
                            
    'uncurrent the current message
    If Not moCurrentMsg Is Nothing Then
        moCurrentMsg.Processed = mimpProcessed
        moCurrentMsg.History = MIMsgHistory.mimhNotCurrent
    End If
    
    mcolMIMessages.Add oMsg
    Set moCurrentMsg = oMsg
    
    Set oMsg = Nothing
    'return the new index
    Add = mcolMIMessages.Count

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIFreezer.Add"

End Function

'----------------------------------------------------------------------------------------'
Private Function ValidStatusTransition(enNewStatus As eFreezerMIMStatus) As Boolean
'----------------------------------------------------------------------------------------'
' Returns TRUE if NewStatus is a valid status transition for this Freezer
'----------------------------------------------------------------------------------------'

    ValidStatusTransition = False

    ' Mustn't be the same status
    If CurrentStatus = enNewStatus Then Exit Function

    Select Case CurrentStatus
    Case eFreezerMIMStatus.fsUnlocked
        ' Can go to anything (except the same, already screened out)
        ValidStatusTransition = True

    Case eFreezerMIMStatus.fsLocked
        ' Can be frozen or unlocked
        If enNewStatus = eFreezerMIMStatus.fsFrozen Then
            ValidStatusTransition = True
        End If

    Case eFreezerMIMStatus.fsFrozen
        ' No status transitions allowed

    End Select
 
End Function


'----------------------------------------------------------------------------------------'
Public Property Get CurrentMessage() As MIMsg
'----------------------------------------------------------------------------------------'
'Return the current message row for the Freezer.
'----------------------------------------------------------------------------------------'
    
    If mcolMIMessages.Count > 0 Then
        Set CurrentMessage = mcolMIMessages(mcolMIMessages.Count)
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get Messages() As Collection
'----------------------------------------------------------------------------------------'
'Collection of all the message rows.
'----------------------------------------------------------------------------------------'
    
    Set Messages = mcolMIMessages

End Property

'properties common to the Freezer (all MIMessage rows)

Public Property Get MIMessageType() As MIMsgType
    MIMessageType = MIMsgType.mimtFreezer
End Property

Public Property Get Scope() As MIMsgScope
    Scope = moCurrentMsg.Scope
End Property

Public Property Get StudyName() As String
    StudyName = moCurrentMsg.StudyName
End Property

Public Property Get Site() As String
    Site = moCurrentMsg.Site
End Property

Public Property Get SubjectId() As Long
    SubjectId = moCurrentMsg.SubjectId
End Property

Public Property Get VisitId() As Long
    VisitId = moCurrentMsg.VisitId
End Property

Public Property Get VisitCycle() As Integer
    VisitCycle = moCurrentMsg.VisitCycle
End Property

Public Property Get EFormTaskId() As Long
    EFormTaskId = moCurrentMsg.EFormTaskId
End Property

Public Property Get ResponseTaskId() As Long
    ResponseTaskId = moCurrentMsg.ResponseTaskId
End Property

Public Property Get ResponseCycle() As Integer
    ResponseCycle = moCurrentMsg.ResponseCycle
End Property

Public Property Get FreezerID() As Long
    FreezerID = moCurrentMsg.MIMessageId
End Property

Public Property Get FreezerSource() As MIMsgSource
    FreezerSource = moCurrentMsg.MIMessagesSource
End Property

Public Property Get CurrentStatus() As eFreezerMIMStatus
    CurrentStatus = moCurrentMsg.Status
End Property

Public Property Get EFormId() As Long
    EFormId = moCurrentMsg.EFormId
End Property
 
Public Property Get EFormCycle() As Integer
    EFormCycle = moCurrentMsg.EFormCycle
End Property

Public Property Get QuestionId() As Long
    QuestionId = moCurrentMsg.QuestionId
End Property

Public Property Get DataUserName() As String
    DataUserName = moCurrentMsg.DataUserName
End Property
 
'--------------------------------------------
Public Property Get StudyId() As Long
'--------------------------------------------
' NCJ 21 Nov 02
' Get the StudyId the first time we use it
' Make a quick dive into the database
'--------------------------------------------
Dim sSQL  As String
Dim rsID As ADODB.Recordset

    If mlStudyId = 0 Then
        sSQL = "SELECT ClinicalTrialId FROM ClinicalTrial " & _
                " WHERE ClinicalTrialName = '" & Me.StudyName & "'"
        Set rsID = New ADODB.Recordset
        rsID.Open sSQL, moADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        mlStudyId = rsID!ClinicalTrialId
        
        rsID.Close
        Set rsID = Nothing
    End If
    
    StudyId = mlStudyId

End Property
 
'--------------------------------------------
Private Sub Class_Initialize()
'--------------------------------------------

    ' NCJ 20 Nov 02 - For now we predefine the hierarchy to be Frozen overrides Locked
    mbFreezeOverridesLock = True
    mlStudyId = 0

End Sub

'------------------------------------------------------------
Private Sub SetTrialSubjectLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal nLockSetting As eFreezerMIMStatus, _
                            Optional bCheckForFrozen As Boolean = True)
'------------------------------------------------------------
' Lock, Unlock or Freeze a trial subject
' i.e. apply setting to subject and to all its visits, forms and data items that are not frozen
' NCJ 25/4/00 - Don't change items that already have this status
' Mo Morris 28/4/00, optional override of Frozen state added
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrLabel

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus
    sSQLSetLockWhere = " LockStatus = " & nLockSetting & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND LockStatus <> " & nLockSetting
    
    If bCheckForFrozen Then
        sSQLSetLockWhere = sSQLSetLockWhere & _
            " AND LockStatus <> " & eFreezerMIMStatus.fsFrozen
    End If

    ' Set lock on Trial Subject
    sSQL = "UPDATE TrialSubject SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    moADODBConnection.Execute sSQL
    
    ' Set lock on all the visit instances
    sSQL = "UPDATE VisitInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    moADODBConnection.Execute sSQL
    
    ' Set lock on all the CRF Pages
    sSQL = "UPDATE CRFPageInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    moADODBConnection.Execute sSQL
    
    ' Set lock on all the Data Items
    sSQL = "UPDATE DataItemResponse SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    moADODBConnection.Execute sSQL
            
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIFreezer.SetTrialSubjectLockStatus"

End Sub

'------------------------------------------------------------
Private Sub SetVisitInstanceLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lVisitId As Long, _
                            ByVal nVisitCycleNumber As Integer, _
                            ByVal nLockSetting As eFreezerMIMStatus, _
                            Optional bCheckForFrozen As Boolean = True)
'------------------------------------------------------------
' NCJ 29/2/00
' Lock, Unlock or Freeze a visit instance
' Apply setting to visit and to all its forms and data items that are not frozen
' NCJ 25/4/00 - Don't change items that already have this status
' Mo Morris 28/4/00, optional override of Frozen state added
' MLM 10/07/02: Be more careful with DIRH insert
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrLabel

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus and the Changed flag
    sSQLSetLockWhere = " LockStatus = " & nLockSetting & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND VisitId = " & lVisitId & _
            " AND VisitCycleNumber = " & nVisitCycleNumber & _
            " AND LockStatus <> " & nLockSetting
    
    If bCheckForFrozen Then
        sSQLSetLockWhere = sSQLSetLockWhere & _
            " AND LockStatus <> " & eFreezerMIMStatus.fsFrozen
    End If

    ' Lock the visit instance
    sSQL = "UPDATE VisitInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    moADODBConnection.Execute sSQL
    
    ' Lock all the CRF Pages
    sSQL = "UPDATE CRFPageInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    moADODBConnection.Execute sSQL
    
    ' Lock all the Data Items
    sSQL = "UPDATE DataItemResponse SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    moADODBConnection.Execute sSQL

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIFreezer.SetVisitInstanceLockStatus"

End Sub

'------------------------------------------------------------
Private Sub SetCRFPageInstanceLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lCRFPageTaskId As Long, _
                            ByVal nLockSetting As eFreezerMIMStatus, _
                            Optional bCheckForFrozen As Boolean = True)
'------------------------------------------------------------
' NCJ 29/2/00
' Lock, Unlock or Freeze a CRF page
' Apply setting to page and to all its data items that are not frozen
' NCJ 25/4/00 - Don't change items that already have this status
' Mo Morris 28/4/00, optional override of Frozen state added
' MLM 10/07/02: Be more careful with DIRH insert.
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrLabel

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus and the Changed flag
    sSQLSetLockWhere = " LockStatus = " & nLockSetting & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND CRFPageTaskId = " & lCRFPageTaskId & _
            " AND LockStatus <> " & nLockSetting

    If bCheckForFrozen Then
        sSQLSetLockWhere = sSQLSetLockWhere & _
            " AND LockStatus <> " & eFreezerMIMStatus.fsFrozen
    End If

    ' Lock the CRF Page
    sSQL = "UPDATE CRFPageInstance SET"
    sSQL = sSQL & sSQLSetLockWhere
    
    moADODBConnection.Execute sSQL
    
    ' Lock all the Data Items
    sSQL = "UPDATE DataItemResponse SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    moADODBConnection.Execute sSQL

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIFreezer.SetCRFPageInstanceLockStatus"

End Sub

'------------------------------------------------------------
Private Sub SetDataItemLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lResponseTaskId As Long, _
                            ByVal nRepeatNumber As Integer, _
                            ByVal nLockSetting As eFreezerMIMStatus, _
                            Optional bCheckForFrozen As Boolean = True)
'------------------------------------------------------------
' NCJ 29/2/00
' Lock, Freeze or Unlock a data item
' NCJ 25/4/00 - Don't change items that already have this status
' Mo Morris 28/4/00, optional override of Frozen state added
' ATO 20/08/2002 Added RepeatNumber
'------------------------------------------------------------
Dim sSQL  As String

    On Error GoTo ErrLabel

    ' Lock the Data Item - set the LockStatus
    sSQL = "UPDATE DataItemResponse " & _
            " SET LockStatus = " & nLockSetting & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND ResponseTaskId = " & lResponseTaskId & _
            " AND RepeatNumber = " & nRepeatNumber & _
            " AND LockStatus <> " & nLockSetting
            
    If bCheckForFrozen Then
        sSQL = sSQL & _
            " AND LockStatus <> " & eFreezerMIMStatus.fsFrozen
    End If

    moADODBConnection.Execute sSQL

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIFreezer.SetDataItemLockStatus"

End Sub

'------------------------------------------------------------
Private Sub UnlockTrialSubject(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer)
'------------------------------------------------------------
' Unlock a trial subject
' Apply setting only to subject if not frozen
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrLabel

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus and the Changed flag, but only if it is locked (not frozen or unlocked)
    sSQLSetLockWhere = " LockStatus = " & eFreezerMIMStatus.fsUnlocked & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND LockStatus = " & eFreezerMIMStatus.fsLocked

    sSQL = "UPDATE TrialSubject SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    moADODBConnection.Execute sSQL
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIFreezer.UnlockTrialSubject"

End Sub

'------------------------------------------------------------
Private Sub UnlockVisitInstance(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lVisitId As Long, _
                            ByVal nVisitCycleNumber As Integer)
'------------------------------------------------------------
' Unlock a Visit Instance if not frozen
' and ALSO unlock the Trial Subject
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrLabel

    ' Set the LockStatus and the Changed flag, but only if it is locked (not frozen or unlocked)
    sSQLSetLockWhere = " LockStatus = " & eFreezerMIMStatus.fsUnlocked & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND VisitId = " & lVisitId & _
            " AND VisitCycleNumber = " & nVisitCycleNumber & _
            " AND LockStatus = " & eFreezerMIMStatus.fsLocked

    ' Lock the visit instance
    sSQL = "UPDATE VisitInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    moADODBConnection.Execute sSQL

    ' Make sure the Trial Subject is also unlocked
    Call UnlockTrialSubject(lClinicalTrialId, sTrialSite, nPersonId)
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIFreezer.UnlockVisitInstance"

End Sub

'------------------------------------------------------------
Private Sub UnlockCRFPageInstance(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lCRFPageTaskId As Long, _
                            ByVal lVisitId As Long, _
                            ByVal nVisitCycleNumber As Integer)
'------------------------------------------------------------
' Unlock a CRF Page Instance if not frozen
' and ALSO unlock its Visit and Trial Subject
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrLabel

    ' Set the LockStatus and the Changed flag, but only if it is locked (not frozen or unlocked)
    sSQLSetLockWhere = " LockStatus = " & eFreezerMIMStatus.fsUnlocked & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND CRFPageTaskId = " & lCRFPageTaskId & _
            " AND LockStatus = " & eFreezerMIMStatus.fsLocked

    ' Lock the CRF Page
    sSQL = "UPDATE CRFPageInstance SET"
    sSQL = sSQL & sSQLSetLockWhere
    
    moADODBConnection.Execute sSQL

    ' Make sure the Visit and Trial subject are also unlocked
    Call UnlockVisitInstance(lClinicalTrialId, sTrialSite, nPersonId, _
                            lVisitId, nVisitCycleNumber)

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIFreezer.UnlockCRFPageInstance"

End Sub

'------------------------------------------------------------
Private Function GetUnfrozenLockStatus(ByVal enScope As MIMsgScope, _
                            ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lVisitId As Long, _
                            ByVal nVisitCycleNumber As Integer, _
                            ByVal lCRFPageTaskId As Long) As eFreezerMIMStatus
'------------------------------------------------------------
' When unfreezing an object, determine the new lock status,
' which is the lockstatus of the first non-frozen parent
'------------------------------------------------------------
Dim enLockStatus As eFreezerMIMStatus

    On Error GoTo ErrLabel
    
    Select Case enScope
    Case MIMsgScope.mimscQuestion
        ' For a question, try upwards from the CRFPageInstance table
        enLockStatus = GetEFormInstanceLockStatus(lClinicalTrialId, sTrialSite, nPersonId, lCRFPageTaskId)
        If enLockStatus = fsFrozen Then
            ' Try the VisitInstance
            enLockStatus = GetVisitInstanceLockStatus(lClinicalTrialId, sTrialSite, nPersonId, _
                                            lVisitId, nVisitCycleNumber)
            If enLockStatus = fsFrozen Then
                ' Try the subject
                enLockStatus = GetSubjectLockStatus(lClinicalTrialId, sTrialSite, nPersonId)
            End If
        End If
        
    Case MIMsgScope.mimscEForm
        ' For a form, try upwards from the VisitInstance table
        enLockStatus = GetVisitInstanceLockStatus(lClinicalTrialId, sTrialSite, nPersonId, _
                                        lVisitId, nVisitCycleNumber)
        If enLockStatus = fsFrozen Then
            ' Try the subject
            enLockStatus = GetSubjectLockStatus(lClinicalTrialId, sTrialSite, nPersonId)
        End If
        
    Case MIMsgScope.mimscVisit
        ' For a visit, try the TrialSubject table
        enLockStatus = GetSubjectLockStatus(lClinicalTrialId, sTrialSite, nPersonId)
        
    Case MIMsgScope.mimscSubject
        ' To unfreeze a subject, set it to Unlocked
        enLockStatus = fsUnlocked
        
    End Select
    
    ' If we only found frozen stuff, we set to Unlocked
    If enLockStatus = fsFrozen Then
        enLockStatus = fsUnlocked
    End If
    
    GetUnfrozenLockStatus = enLockStatus
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|MIFreezer.GetUnfrozenLockStatus"

End Function

'------------------------------------------------------------
Private Function GetVisitInstanceLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lVisitId As Long, _
                            ByVal nVisitCycleNumber As Integer) As eFreezerMIMStatus
'------------------------------------------------------------
' Get the lockstatus of a visit instance
'------------------------------------------------------------
Dim sSQL  As String
Dim rsLocks As ADODB.Recordset

    sSQL = "SELECT LockStatus FROM VisitInstance " & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND VisitId = " & lVisitId & _
            " AND VisitCycleNumber = " & nVisitCycleNumber
    Set rsLocks = New ADODB.Recordset
    rsLocks.Open sSQL, moADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    GetVisitInstanceLockStatus = rsLocks!LockStatus
    
    rsLocks.Close
    Set rsLocks = Nothing

End Function

'------------------------------------------------------------
Private Function GetEFormInstanceLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lCRFPageTaskId As Long) As eFreezerMIMStatus
'------------------------------------------------------------
' Get the lockstatus of a visit instance
'------------------------------------------------------------
Dim sSQL  As String
Dim rsLocks As ADODB.Recordset

    sSQL = "SELECT LockStatus FROM CRFPageInstance " & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND CRFPageTaskId = " & lCRFPageTaskId
    Set rsLocks = New ADODB.Recordset
    rsLocks.Open sSQL, moADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    GetEFormInstanceLockStatus = rsLocks!LockStatus
    
    rsLocks.Close
    Set rsLocks = Nothing
    
End Function


'------------------------------------------------------------
Private Function GetSubjectLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer) As eFreezerMIMStatus
'------------------------------------------------------------
' Get the lockstatus of a visit instance
'------------------------------------------------------------
Dim sSQL  As String
Dim rsLocks As ADODB.Recordset

    sSQL = "SELECT LockStatus FROM TrialSubject " & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId
    Set rsLocks = New ADODB.Recordset
    rsLocks.Open sSQL, moADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    GetSubjectLockStatus = rsLocks!LockStatus
    
    rsLocks.Close
    Set rsLocks = Nothing
    
End Function

