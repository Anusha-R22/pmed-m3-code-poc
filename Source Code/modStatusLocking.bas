Attribute VB_Name = "modStatusLocking"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       modStatusLocking.bas
'   Author:     Mo Morris, April 2000
'   Purpose:    Patient data Locking/UnLocking/Freezing Subroutines
'
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   ASH 18/06/2002 Added routine SetTrialSubjectChanged to set the
'   changed flag in the TrialSubject table
'   MLM 10/07/02: CBB 2.2.19/28: Be more careful with DataItemResponseHistory inserts.
'   ATO 20/08/2002 Added RepeatNumber to SetDataItemLockStatus to allow locking of RQGs
'   NCJ 2 Oct 02 - Replaced CDbl(Now) with new more accurate IMedNow function
'----------------------------------------------------------------------------------------'

Option Explicit
Option Base 0
Option Compare Binary

'------------------------------------------------------------
Public Sub SetTrialSubjectLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal nLockSetting As LockStatus, _
                            Optional bCheckForFrozen As Boolean = True)
'------------------------------------------------------------
' Lock, Unlock or Freeze a trial subject
' i.e. apply setting to subject and to all its visits, forms and data items that are not frozen
' NCJ 25/4/00 - Don't change items that already have this status
' Mo Morris 28/4/00, optional override of Frozen state added
'------------------------------------------------------------
Dim sSQL  As String
Dim sTimestamp As String
Dim sSQLSetLockWhere As String
Dim oTimezone As Timezone

    On Error GoTo ErrHandler

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus and the Changed flag
    sSQLSetLockWhere = " LockStatus = " & nLockSetting & ", " & _
            " Changed = " & Changed.Changed & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND LockStatus <> " & nLockSetting
    
    If bCheckForFrozen Then
        sSQLSetLockWhere = sSQLSetLockWhere & _
            " AND LockStatus <> " & LockStatus.lsFrozen
    End If

    ' Set lock on Trial Subject
    sSQL = "UPDATE TrialSubject SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL
    
    ' Set lock on all the visit instances
    sSQL = "UPDATE VisitInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL
    
    ' Set lock on all the CRF Pages
    sSQL = "UPDATE CRFPageInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL
    
    ' Get the string timestamp
    sTimestamp = LocalNumToStandard(CStr(IMedNow))
    
    ' Get TimeZone
    Set oTimezone = New Timezone

    ' Set lock on all the Data Items
    'Mo Morris 30/8/01 Db Audit (UserId to UserName)
    ' RS 25/10/2002 Add Timezone, and Set DatabaseTimestamp to 0 to force trigger to update value
    sSQL = "UPDATE DataItemResponse SET " & _
           " ResponseTimestamp = " & sTimestamp & ", " & _
           " ResponseTimestamp_TZ = " & oTimezone.TimezoneOffset & ", " & _
           " DatabaseTimestamp = " & "0" & ", " & _
           " UserName = '" & goUser.UserName & "', "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL

    ' Now transfer the stuff into DataItemResponseHistory
    sSQL = SQLToUpdateDataItemResponseHistory
    ' Pick off the items we've just created, matching on the TimeStamp
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND ResponseTimeStamp = " & sTimestamp
    
    MacroADODBConnection.Execute sSQL
    
    'Changed Mo Morris 26/4/00
    'Create a message to be sent to the remote site
    Call CreateTrialSubjectLockStatusMessage(lClinicalTrialId, sTrialSite, nPersonId, nLockSetting, sTimestamp)
            
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, _
                                      Err.Description, _
                                      "SetTrialSubjectLockStatus", _
                                      "modStatusLocking")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------------------------
Public Sub SetVisitInstanceLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lVisitId As Long, _
                            ByVal nVisitCycleNumber As Integer, _
                            ByVal nLockSetting As LockStatus, _
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
Dim sTimestamp As String
Dim sSQLSetLockWhere As String
Dim oTimezone As Timezone

    On Error GoTo ErrHandler

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus and the Changed flag
    sSQLSetLockWhere = " LockStatus = " & nLockSetting & ", " & _
            " Changed = " & Changed.Changed & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND VisitId = " & lVisitId & _
            " AND VisitCycleNumber = " & nVisitCycleNumber & _
            " AND LockStatus <> " & nLockSetting
    
    If bCheckForFrozen Then
        sSQLSetLockWhere = sSQLSetLockWhere & _
            " AND LockStatus <> " & LockStatus.lsFrozen
    End If

    ' Lock the visit instance
    sSQL = "UPDATE VisitInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL
    
    ' Lock all the CRF Pages
    sSQL = "UPDATE CRFPageInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL
    
    ' Get the string timestamp
    sTimestamp = LocalNumToStandard(CStr(IMedNow))
    Set oTimezone = New Timezone
    
    ' Lock all the Data Items
    'Mo Morris 30/8/01 Db Audit (UserId to UserName)
    sSQL = "UPDATE DataItemResponse SET " & _
           " ResponseTimestamp = " & sTimestamp & ", " & _
           " ResponseTimestamp_TZ = " & oTimezone.TimezoneOffset & ", " & _
           " DatabaseTimestamp = " & "0" & ", " & _
           " UserName = '" & goUser.UserName & "', "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL

    ' Now transfer the stuff into DataItemResponseHistory
    sSQL = SQLToUpdateDataItemResponseHistory
    ' Pick off the items we've just created, matching on the TimeStamp
    'MLM 10/07/02: Added VisitId and VisitCycleNumber to where clause
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND VisitId = " & lVisitId & _
            " AND VisitCycleNumber = " & nVisitCycleNumber & _
            " AND ResponseTimeStamp = " & sTimestamp
    
    MacroADODBConnection.Execute sSQL
    
    'Changed Mo Morris 26/4/00
    'Create a message to be sent to the remote site
    Call CreateVisitInstanceLockStatusMessage(lClinicalTrialId, sTrialSite, nPersonId, lVisitId, nVisitCycleNumber, nLockSetting, sTimestamp)
    
    'ASH 18/06/2002
    'Bug 2.2.14 no.7
    Call SetTrialSubjectChanged(lClinicalTrialId, sTrialSite, nPersonId)
            
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, _
                                      Err.Description, _
                                      "SetVisitInstanceLockStatus", _
                                      "modStatusLocking")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------------------------
Public Sub SetCRFPageInstanceLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lCRFPageTaskId As Long, _
                            ByVal nLockSetting As LockStatus, _
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
Dim sTimestamp As String
Dim sSQLSetLockWhere As String
Dim oTimezone As Timezone

    On Error GoTo ErrHandler

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus and the Changed flag
    sSQLSetLockWhere = " LockStatus = " & nLockSetting & ", " & _
            " Changed = " & Changed.Changed & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND CRFPageTaskId = " & lCRFPageTaskId & _
            " AND LockStatus <> " & nLockSetting

    If bCheckForFrozen Then
        sSQLSetLockWhere = sSQLSetLockWhere & _
            " AND LockStatus <> " & LockStatus.lsFrozen
    End If

    ' Lock the CRF Page
    sSQL = "UPDATE CRFPageInstance SET"
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL
    
    ' Get the string timestamp
    sTimestamp = LocalNumToStandard(CStr(IMedNow))
    Set oTimezone = New Timezone
    
    ' Lock all the Data Items
    'Mo Morris 30/8/01 Db Audit (UserId to UserName)
    sSQL = "UPDATE DataItemResponse SET " & _
           " ResponseTimestamp = " & sTimestamp & ", " & _
           " ResponseTimestamp_TZ = " & oTimezone.TimezoneOffset & ", " & _
           " DatabaseTimestamp = " & "0" & ", " & _
           " UserName = '" & goUser.UserName & "', "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL

    ' Now transfer the stuff into DataItemResponseHistory
    sSQL = SQLToUpdateDataItemResponseHistory
    ' Pick off the items we've just created, matching on the TimeStamp
    'MLM 10/07/02: Added CRFPageTaskId to where clause.
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND CRFPageTaskId = " & lCRFPageTaskId & _
            " AND ResponseTimeStamp = " & sTimestamp
    
    MacroADODBConnection.Execute sSQL
    
    'Changed Mo Morris 26/4/00
    'Create a message to be sent to the remote site
    Call CreateCRFPageInstanceLockStatusMessage(lClinicalTrialId, sTrialSite, nPersonId, lCRFPageTaskId, nLockSetting, sTimestamp)
    
    'ash 18/06/2002
    'Bug 2.2.14 no.7
    Call SetTrialSubjectChanged(lClinicalTrialId, sTrialSite, nPersonId)
            
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, _
                                      Err.Description, _
                                      "SetCRFPageInstanceLockStatus", _
                                      "modStatusLocking")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------------------------
Public Sub SetDataItemLockStatus(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer, _
                            ByVal lResponseTaskId As Long, _
                            ByVal nRepeatNumber As Integer, _
                            ByVal nLockSetting As LockStatus, _
                            Optional bCheckForFrozen As Boolean = True)
'------------------------------------------------------------
' NCJ 29/2/00
' Lock, Freeze or Unlock a data item
' NCJ 25/4/00 - Don't change items that already have this status
' Mo Morris 28/4/00, optional override of Frozen state added
' ATO 20/08/2002 Added RepeatNumber
'------------------------------------------------------------
Dim sSQL  As String
Dim sTimestamp As String
Dim oTimezone As Timezone

    On Error GoTo ErrHandler

    ' Get the string timestamp
    sTimestamp = LocalNumToStandard(CStr(IMedNow))
    Set oTimezone = New Timezone
    
    ' Lock the Data Item - set the LockStatus and the Changed flag
    'Mo Morris 30/8/01 Db Audit (UserId to UserName)
    sSQL = "UPDATE DataItemResponse " & _
           " SET ResponseTimestamp = " & sTimestamp & ", " & _
           " ResponseTimestamp_TZ = " & oTimezone.TimezoneOffset & ", " & _
           " DatabaseTimestamp = " & "0" & ", " & _
           " UserName = '" & goUser.UserName & "', " & _
           " LockStatus = " & nLockSetting & ", " & _
           " Changed = " & Changed.Changed
    sSQL = sSQL & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND ResponseTaskId = " & lResponseTaskId & _
            " AND RepeatNumber = " & nRepeatNumber & _
            " AND LockStatus <> " & nLockSetting
            
    If bCheckForFrozen Then
        sSQL = sSQL & _
            " AND LockStatus <> " & LockStatus.lsFrozen
    End If

    MacroADODBConnection.Execute sSQL

    ' Now transfer the stuff into DataItemResponseHistory
    sSQL = SQLToUpdateDataItemResponseHistory
    ' Pick off the item we've just created, matching on the TimeStamp
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND ResponseTaskId = " & lResponseTaskId & _
            " AND RepeatNumber = " & nRepeatNumber & _
            " AND ResponseTimeStamp = " & sTimestamp
    
    MacroADODBConnection.Execute sSQL
    
    'Changed Mo Morris 26/4/00
    'Create a message to be sent to the remote site
    Call CreateDataItemLockStatusMessage(lClinicalTrialId, sTrialSite, nPersonId, lResponseTaskId, nRepeatNumber, nLockSetting, sTimestamp)
    'ASH 18/06/2002
    'Bug 2.2.14 no.7
    Call SetTrialSubjectChanged(lClinicalTrialId, sTrialSite, nPersonId)
            
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, _
                                      Err.Description, _
                                      "SetDataItemLockStatus", _
                                      "modStatusLocking")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------------------------
Public Sub UnlockTrialSubject(ByVal lClinicalTrialId As Long, _
                            ByVal sTrialSite As String, _
                            ByVal nPersonId As Integer)
'------------------------------------------------------------
' Unlock a trial subject
' Apply setting only to subject if not frozen
'------------------------------------------------------------
Dim sSQL  As String
Dim sSQLSetLockWhere As String

    On Error GoTo ErrHandler

    ' The WHERE clause is the same for all the tables we update
    ' Set the LockStatus and the Changed flag, but only if it is locked (not frozen or unlocked)
    sSQLSetLockWhere = " LockStatus = " & LockStatus.lsUnlocked & ", " & _
            " Changed = " & Changed.Changed & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND LockStatus = " & LockStatus.lsLocked

    sSQL = "UPDATE TrialSubject SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, _
                                      Err.Description, _
                                      "UnlockTrialSubject", _
                                      "modStatusLocking")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------------------------
Public Sub UnlockVisitInstance(ByVal lClinicalTrialId As Long, _
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

    On Error GoTo ErrHandler

    ' Set the LockStatus and the Changed flag, but only if it is locked (not frozen or unlocked)
    sSQLSetLockWhere = " LockStatus = " & LockStatus.lsUnlocked & ", " & _
            " Changed = " & Changed.Changed & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND VisitId = " & lVisitId & _
            " AND VisitCycleNumber = " & nVisitCycleNumber & _
            " AND LockStatus = " & LockStatus.lsLocked

    ' Lock the visit instance
    sSQL = "UPDATE VisitInstance SET "
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL

    ' Make sure the Trial Subject is also unlocked
    Call UnlockTrialSubject(lClinicalTrialId, sTrialSite, nPersonId)
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, _
                                      Err.Description, _
                                      "UnlockVisitInstance", _
                                      "modStatusLocking")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------------------------
Public Sub UnlockCRFPageInstance(ByVal lClinicalTrialId As Long, _
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

    On Error GoTo ErrHandler

    ' Set the LockStatus and the Changed flag, but only if it is locked (not frozen or unlocked)
    sSQLSetLockWhere = " LockStatus = " & LockStatus.lsUnlocked & ", " & _
            " Changed = " & Changed.Changed & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId & _
            " AND CRFPageTaskId = " & lCRFPageTaskId & _
            " AND LockStatus = " & LockStatus.lsLocked

    ' Lock the CRF Page
    sSQL = "UPDATE CRFPageInstance SET"
    sSQL = sSQL & sSQLSetLockWhere
    
    MacroADODBConnection.Execute sSQL

    ' Make sure the Visit and Trial subject are also unlocked
    Call UnlockVisitInstance(lClinicalTrialId, sTrialSite, nPersonId, _
                            lVisitId, nVisitCycleNumber)

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, _
                                      Err.Description, _
                                      "UnlockCRFPageInstance", _
                                      "modStatusLocking")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub
'-----------------------------------------------------------------------------
Public Sub SetTrialSubjectChanged(ByVal lClinicalTrialId As Long, _
                                ByVal sTrialSite As String, _
                                ByVal nPersonId As Integer)
'-----------------------------------------------------------------------------
'ASH 18/06/2002 Bug 2.2.14 no.7
'Sets the changed flag in the TrialSubject table
'-----------------------------------------------------------------------------
Dim sSQL  As String

On Error GoTo ErrHandler

    sSQL = "Update TrialSubject Set Changed = " & Changed.Changed & _
            " WHERE ClinicalTrialId = " & lClinicalTrialId & _
            " AND TrialSite = '" & sTrialSite & "'" & _
            " AND PersonId = " & nPersonId
            
    MacroADODBConnection.Execute sSQL


Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modStatusLocking.SetTrialSubjectChanged"

End Sub
