Attribute VB_Name = "modStatusLockingWWW"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       StatusLockingWW.bas
'   Author:     REM August 2001
'   Purpose:    Patient data Locking/UnLocking/Freezing Subroutines
'
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
' REM 03/08/01 This is a copy of the original StatusLocking.bas from the windows version that has been modified
' for the web application.  Also the entire module has been convereted across into QueryDef from ADO.
' revisions
' ic 22/11/2002 pasted LockStatusResult enum
'----------------------------------------------------------------------------------------'

Public Enum LockStatusResult
    LockStatusFailed = 0
    LockStatusPassed = 1
End Enum

Option Explicit
Option Base 0
Option Compare Binary

'--------------------------------------------------------------------------------------------------
Public Function SetTrialSubjectLockStatus(ByVal sDatabaseCnn As String, _
                                          ByVal sUserName As String, _
                                          ByVal sDatabase As String, _
                                          ByVal lStudyCode As Long, _
                                          ByVal sSiteCode As String, _
                                          ByVal nSubjectId As Integer, _
                                          ByVal nLockSetting As LockStatus, _
                                       Optional bCheckForFrozen As Boolean = True) As Integer
'--------------------------------------------------------------------------------------------------
' REM 03/08/01
' Lock, Unlock or Freeze a trial subject
'--------------------------------------------------------------------------------------------------
Dim vFilterFields As Variant
Dim vFilterComps As Variant
Dim vFilterExpr As Variant
Dim vCols As Variant
Dim vVals As Variant
Dim dblTimeStamp As Double
Dim sSQL As String
Dim oQueryDef As QueryDef
Dim oQueryServer As QueryServer

    vCols = Array("LockStatus", "Changed")
    vVals = Array(nLockSetting, Changed.Changed)

    If bCheckForFrozen Then
        vFilterFields = Array("ClinicalTrialId", "TrialSite", "PersonId", "LockStatus", "LockStatus")
        vFilterComps = Array("=", "=", "=", "<>", "<>")
        vFilterExpr = Array(lStudyCode, sSiteCode, nSubjectId, nLockSetting, LockStatus.lsFrozen)
    Else
        vFilterFields = Array("ClinicalTrialId", "TrialSite", "PersonId", "LockStatus")
        vFilterComps = Array("=", "=", "=", "<>")
        vFilterExpr = Array(lStudyCode, sSiteCode, nSubjectId, nLockSetting)
    End If

    Set oQueryDef = New QueryDef
    Set oQueryServer = New QueryServer
    
    oQueryServer.ConnectionOpen sDatabaseCnn
    
    oQueryServer.BeginTrans
   ' On Error GoTo ErrHandlerWhenInTrans

    ' Set lock on Trial Subject
    oQueryDef.InitSave "TrialSubject", vCols, vVals
    oQueryDef.QueryFilters.Add vFilterFields, vFilterComps, vFilterExpr
    oQueryServer.SelectSave stUpdate, oQueryDef
    
   ' Set lock on all the visit instances
    oQueryDef.InitSave "VisitInstance", vCols, vVals
    oQueryDef.QueryFilters.Add vFilterFields, vFilterComps, vFilterExpr
    oQueryServer.SelectSave stUpdate, oQueryDef
    
    ' Set lock on all the CRF Pages
    oQueryDef.InitSave "CRFPageInstance", vCols, vVals
    oQueryDef.QueryFilters.Add vFilterFields, vFilterComps, vFilterExpr
    oQueryServer.SelectSave stUpdate, oQueryDef

    ' Get the string timestamp,
    ' this will not deal with regional settings!!
    dblTimeStamp = CDbl(Now)

    vCols = Array("LockStatus", "Changed", "ResponseTimeStamp", "UserName")
    vVals = Array(nLockSetting, Changed.Changed, dblTimeStamp, sUserName)

    ' Set lock on all the Data Items
    oQueryDef.InitSave "DataItemResponse", vCols, vVals
    oQueryDef.QueryFilters.Add vFilterFields, vFilterComps, vFilterExpr
   ' oQueryDef.QueryFields.Add Array("ResponseTimeStamp", "UserName"), , Array(dblTimeStamp, sUserName)
    oQueryServer.SelectSave stUpdate, oQueryDef
    
    
    ' Now transfer the stuff into DataItemResponseHistory
    sSQL = SQLToUpdateDataItemResponseHistory
    ' Pick off the items we've just created, matching on the TimeStamp
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lStudyCode & _
            " AND TrialSite = '" & sSiteCode & "'" & _
            " AND PersonId = " & nSubjectId & _
            " AND ResponseTimeStamp = " & dblTimeStamp
    oQueryServer.SQLExecute sSQL
    
    oQueryServer.Commit
    
    
    Set oQueryServer = Nothing
    Set oQueryDef = Nothing
    
    On Error GoTo ErrHandler
            
    SetTrialSubjectLockStatus = LockStatusResult.LockStatusPassed
            
Exit Function

ErrHandler:

    SetTrialSubjectLockStatus = LockStatusResult.LockStatusFailed
    Debug.Print "Ran Error"
    
Exit Function

ErrHandlerWhenInTrans:
    oQueryServer.Rollback
    Err.Raise Err.Number
    SetTrialSubjectLockStatus = LockStatusResult.LockStatusFailed
Exit Function

End Function


'--------------------------------------------------------------------------------------------------
Public Function SetVisitInstanceLockStatus(ByVal sDatabaseCnn As String, _
                                           ByVal sUserName As String, _
                                           ByVal sDatabase As String, _
                                           ByVal lStudyCode As Long, _
                                           ByVal sSiteCode As String, _
                                           ByVal nSubjectId As Integer, _
                                           ByVal sVisitId As String, _
                                           ByVal sVisitCycleNumber As String, _
                                           ByVal nLockSetting As LockStatus, _
                                        Optional bCheckForFrozen As Boolean = True) As Integer
'--------------------------------------------------------------------------------------------------
' REM 03/08/01
' Lock, Unlock or Freeze a Visit Instance
'--------------------------------------------------------------------------------------------------
Dim vFilterFields As Variant
Dim vFilterComps As Variant
Dim vFilterExpr As Variant
Dim vCols As Variant
Dim vVals As Variant
Dim dblTimeStamp As Double
Dim sSQL As String
Dim oQueryDef As QueryDef
Dim oQueryServer As QueryServer

    vCols = Array("LockStatus", "Changed")
    vVals = Array(nLockSetting, Changed.Changed)

    If bCheckForFrozen Then
        vFilterFields = Array("ClinicalTrialId", "TrialSite", "PersonId", "VisitId", "VisitCycleNumber", "LockStatus", "LockStatus")
        vFilterComps = Array("=", "=", "=", "=", "=", "<>", "<>")
        vFilterExpr = Array(lStudyCode, sSiteCode, nSubjectId, CLng(sVisitId), CInt(sVisitCycleNumber), nLockSetting, LockStatus.lsFrozen)
    Else
        vFilterFields = Array("ClinicalTrialId", "TrialSite", "PersonId", "VisitId", "VisitCycleNumber", "LockStatus")
        vFilterComps = Array("=", "=", "=", "=", "=", "<>")
        vFilterExpr = Array(lStudyCode, sSiteCode, nSubjectId, CLng(sVisitId), CInt(sVisitCycleNumber), nLockSetting)
    End If


    Set oQueryDef = New QueryDef
    Set oQueryServer = New QueryServer
    
    oQueryServer.ConnectionOpen sDatabaseCnn
    
    oQueryServer.BeginTrans
    On Error GoTo ErrHandlerWhenInTrans
        
   ' Set lock on all the visit instances
    oQueryDef.InitSave "VisitInstance", vCols, vVals
    oQueryDef.QueryFilters.Add vFilterFields, vFilterComps, vFilterExpr
    oQueryServer.SelectSave stUpdate, oQueryDef
    
    ' Set lock on all the CRF Pages
    oQueryDef.InitSave "CRFPageInstance", vCols, vVals
    oQueryDef.QueryFilters.Add vFilterFields, vFilterComps, vFilterExpr
    oQueryServer.SelectSave stUpdate, oQueryDef

    ' Get the string timestamp,
    ' this will not deal with regional settings!!
    dblTimeStamp = CDbl(Now)

    vCols = Array("LockStatus", "Changed", "ResponseTimeStamp", "UserName")
    vVals = Array(nLockSetting, Changed.Changed, dblTimeStamp, sUserName)

    ' Set lock on all the Data Items
    oQueryDef.InitSave "DataItemResponse", vCols, vVals
    oQueryDef.QueryFilters.Add vFilterFields, vFilterComps, vFilterExpr
'    oQueryDef.QueryFields.Add Array("ResponseTimeStamp", "UserName"), , Array(sTimeStamp, sUserName)
    oQueryServer.SelectSave stUpdate, oQueryDef

    ' Now transfer the stuff into DataItemResponseHistory
    sSQL = SQLToUpdateDataItemResponseHistory
    ' Pick off the items we've just created, matching on the TimeStamp
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lStudyCode & _
            " AND TrialSite = '" & sSiteCode & "'" & _
            " AND PersonId = " & nSubjectId & _
            " AND ResponseTimeStamp = " & dblTimeStamp
    oQueryServer.SQLExecute sSQL
    
    oQueryServer.Commit
    
    Set oQueryServer = Nothing
    Set oQueryDef = Nothing
    
    On Error GoTo ErrHandler
            
    SetVisitInstanceLockStatus = LockStatusResult.LockStatusPassed
            
Exit Function

ErrHandler:

    SetVisitInstanceLockStatus = LockStatusResult.LockStatusFailed
Exit Function

ErrHandlerWhenInTrans:
    oQueryServer.Rollback
    SetVisitInstanceLockStatus = LockStatusResult.LockStatusFailed
Exit Function

End Function


'--------------------------------------------------------------------------------------------------
Public Function SetCRFPageInstanceLockStatus(ByVal sDatabaseCnn As String, _
                                             ByVal sUserName As String, _
                                             ByVal sDatabase As String, _
                                             ByVal lStudyCode As Long, _
                                             ByVal sSiteCode As String, _
                                             ByVal nSubjectId As Integer, _
                                             ByVal sCRFPageTaskId As String, _
                                             ByVal nLockSetting As LockStatus, _
                                          Optional bCheckForFrozen As Boolean = True) As Integer
'--------------------------------------------------------------------------------------------------
' REM 03/08/01
' Lock, Unlock or Freeze a CRF page
'--------------------------------------------------------------------------------------------------
Dim vFilterFields As Variant
Dim vFilterComps As Variant
Dim vFilterExpr As Variant
Dim vCols As Variant
Dim vVals As Variant
Dim dblTimeStamp As Double
Dim sSQL As String
Dim oQueryDef As QueryDef
Dim oQueryServer As QueryServer

    vCols = Array("LockStatus", "Changed")
    vVals = Array(nLockSetting, Changed.Changed)

    If bCheckForFrozen Then
        vFilterFields = Array("ClinicalTrialId", "TrialSite", "PersonId", "CRFPageTaskId", "LockStatus", "LockStatus")
        vFilterComps = Array("=", "=", "=", "=", "<>", "<>")
        vFilterExpr = Array(lStudyCode, sSiteCode, nSubjectId, CLng(sCRFPageTaskId), nLockSetting, LockStatus.lsFrozen)
    Else
        vFilterFields = Array("ClinicalTrialId", "TrialSite", "PersonId", "CRFPageTaskId", "LockStatus")
        vFilterComps = Array("=", "=", "=", "=", "<>")
        vFilterExpr = Array(lStudyCode, sSiteCode, nSubjectId, CLng(sCRFPageTaskId), nLockSetting)
    End If

    Set oQueryDef = New QueryDef
    Set oQueryServer = New QueryServer
    
    oQueryServer.ConnectionOpen sDatabaseCnn
    
    oQueryServer.BeginTrans
    On Error GoTo ErrHandlerWhenInTrans
    
    ' Set lock on all the CRF Pages
    oQueryDef.InitSave "CRFPageInstance", vCols, vVals
    oQueryDef.QueryFilters.Add vFilterFields, vFilterComps, vFilterExpr
    oQueryServer.SelectSave stUpdate, oQueryDef

    ' Get the string timestamp,
    ' this will not deal with regional settings!!
    dblTimeStamp = CDbl(Now)

    vCols = Array("LockStatus", "Changed", "ResponseTimeStamp", "UserName")
    vVals = Array(nLockSetting, Changed.Changed, dblTimeStamp, sUserName)

    ' Set lock on all the Data Items
    oQueryDef.InitSave "DataItemResponse", vCols, vVals
    oQueryDef.QueryFilters.Add vFilterFields, vFilterComps, vFilterExpr
'    oQueryDef.QueryFields.Add Array("ResponseTimeStamp", "UserName"), , Array(sTimeStamp, sUserName)
    oQueryServer.SelectSave stUpdate, oQueryDef

    ' Now transfer the stuff into DataItemResponseHistory
    sSQL = SQLToUpdateDataItemResponseHistory
    ' Pick off the items we've just created, matching on the TimeStamp
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lStudyCode & _
            " AND TrialSite = '" & sSiteCode & "'" & _
            " AND PersonId = " & nSubjectId & _
            " AND ResponseTimeStamp = " & dblTimeStamp
    oQueryServer.SQLExecute sSQL
    
    oQueryServer.Commit
    
    Set oQueryServer = Nothing
    Set oQueryDef = Nothing
    
    On Error GoTo ErrHandler
            
    SetCRFPageInstanceLockStatus = LockStatusResult.LockStatusPassed
            
Exit Function

ErrHandler:

    SetCRFPageInstanceLockStatus = LockStatusResult.LockStatusFailed
Exit Function

ErrHandlerWhenInTrans:
    oQueryServer.Rollback
    SetCRFPageInstanceLockStatus = LockStatusResult.LockStatusFailed
Exit Function

End Function


'--------------------------------------------------------------------------------------------------
Public Function SetDataItemLockStatus(ByVal sDatabaseCnn As String, _
                                      ByVal sUserName As String, _
                                      ByVal sDatabase As String, _
                                      ByVal lStudyCode As Long, _
                                      ByVal sSiteCode As String, _
                                      ByVal nSubjectId As Integer, _
                                      ByVal sResponseTaskId As String, _
                                      ByVal nLockSetting As LockStatus, _
                                   Optional nRepeatNumber As Integer = 1, _
                                   Optional bCheckForFrozen As Boolean = True) As Integer
'--------------------------------------------------------------------------------------------------
' REM 03/08/01
' Lock, Freeze or Unlock a data item
'--------------------------------------------------------------------------------------------------
Dim vFilterFields As Variant
Dim vFilterComps As Variant
Dim vFilterExpr As Variant
Dim vCols As Variant
Dim vVals As Variant
Dim dblTimeStamp As Double
Dim sSQL As String
Dim oQueryDef As QueryDef
Dim oQueryServer As QueryServer

    ' Get the string timestamp,
    ' this will not deal with regional settings!!
    dblTimeStamp = CDbl(Now) ' REM 08/05/02 - moved setting of time stamp to beginning of routine

    vCols = Array("LockStatus", "Changed", "ResponseTimeStamp", "UserName")
    vVals = Array(nLockSetting, Changed.Changed, dblTimeStamp, sUserName)

    If bCheckForFrozen Then
        vFilterFields = Array("ClinicalTrialId", "TrialSite", "PersonId", "ResponseTaskId", "LockStatus", "LockStatus", "RepeatNumber")
        vFilterComps = Array("=", "=", "=", "=", "<>", "<>", "=")
        vFilterExpr = Array(lStudyCode, sSiteCode, nSubjectId, CLng(sResponseTaskId), nLockSetting, LockStatus.lsFrozen, nRepeatNumber)
    Else
        vFilterFields = Array("ClinicalTrialId", "TrialSite", "PersonId", "ResponseTaskId", "LockStatus", "RepeatNumber")
        vFilterComps = Array("=", "=", "=", "=", "<>", "=")
        vFilterExpr = Array(lStudyCode, sSiteCode, nSubjectId, CLng(sResponseTaskId), nLockSetting, nRepeatNumber)
    End If

    Set oQueryDef = New QueryDef
    Set oQueryServer = New QueryServer
    
    oQueryServer.ConnectionOpen sDatabaseCnn
    
    oQueryServer.BeginTrans
    On Error GoTo ErrHandlerWhenInTrans

    ' Set lock on all the Data Items
    oQueryDef.InitSave "DataItemResponse", vCols, vVals
    oQueryDef.QueryFilters.Add vFilterFields, vFilterComps, vFilterExpr
'    oQueryDef.QueryFields.Add Array("ResponseTimeStamp", "UserName"), , Array(sTimeStamp, sUserName)
    oQueryServer.SelectSave stUpdate, oQueryDef

    ' Now transfer the stuff into DataItemResponseHistory
    sSQL = SQLToUpdateDataItemResponseHistory
    ' Pick off the items we've just created, matching on the TimeStamp
    sSQL = sSQL & " WHERE ClinicalTrialId = " & lStudyCode & _
            " AND TrialSite = '" & sSiteCode & "'" & _
            " AND PersonId = " & nSubjectId & _
            " AND ResponseTimeStamp = " & dblTimeStamp
    oQueryServer.SQLExecute sSQL
    
    oQueryServer.Commit
    
    Set oQueryServer = Nothing
    Set oQueryDef = Nothing
    
    On Error GoTo ErrHandler
            
    SetDataItemLockStatus = LockStatusResult.LockStatusPassed
            
Exit Function

ErrHandler:

    SetDataItemLockStatus = LockStatusResult.LockStatusFailed
Exit Function

ErrHandlerWhenInTrans:
    oQueryServer.Rollback
    SetDataItemLockStatus = LockStatusResult.LockStatusFailed
Exit Function

End Function


'--------------------------------------------------------------------------------------------------
Public Function UnlockTrialSubject(ByVal sDatabaseCnn As String, _
                                   ByVal sDatabase As String, _
                                   ByVal lStudyCode As Long, _
                                   ByVal sSiteCode As String, _
                                   ByVal nSubjectId As Integer) As Integer
'--------------------------------------------------------------------------------------------------
' REM 03/08/01
' Unlock a trial subject
' Applies setting only to subject if not frozen
'--------------------------------------------------------------------------------------------------
Dim oQueryDef As QueryDef
Dim oQueryServer As QueryServer

    On Error GoTo ErrHandler
    
    Set oQueryDef = New QueryDef
    Set oQueryServer = New QueryServer

    oQueryServer.ConnectionOpen sDatabaseCnn
    
    oQueryDef.InitSave "TrialSubject", Array("LockStatus", "Changed"), _
                                       Array(LockStatus.lsUnlocked, Changed.Changed)
    oQueryDef.QueryFilters.Add Array("ClinicalTrialId", "TrialSite", "PersonId", "LockStatus"), _
                               Array("=", "=", "=", "="), _
                               Array(lStudyCode, sSiteCode, nSubjectId, LockStatus.lsLocked)
    oQueryServer.SelectSave stUpdate, oQueryDef
    
    Set oQueryServer = Nothing
    Set oQueryDef = Nothing
    
    UnlockTrialSubject = LockStatusResult.LockStatusPassed

Exit Function

ErrHandler:
    UnlockTrialSubject = LockStatusResult.LockStatusFailed
Exit Function

End Function


'--------------------------------------------------------------------------------------------------
Public Function UnlockVisitInstance(ByVal sDatabaseCnn As String, _
                                    ByVal sDatabase As String, _
                                    ByVal lStudyCode As Long, _
                                    ByVal sSiteCode As String, _
                                    ByVal nSubjectId As Integer, _
                                    ByVal sVisitId As String, _
                                    ByVal sVisitCycleNumber As String) As Integer
'--------------------------------------------------------------------------------------------------
' REM 03/08/01
' Unlock a Visit Instance if not frozen
' and ALSO unlock the Trial Subject
'--------------------------------------------------------------------------------------------------
Dim oQueryDef As QueryDef
Dim oQueryServer As QueryServer

    On Error GoTo ErrHandler
    
    Set oQueryDef = New QueryDef
    Set oQueryServer = New QueryServer

    oQueryServer.ConnectionOpen sDatabaseCnn
    
    oQueryDef.InitSave "VisitInstance", Array("LockStatus", "Changed"), _
                                        Array(LockStatus.lsUnlocked, Changed.Changed)
    oQueryDef.QueryFilters.Add Array("ClinicalTrialId", "TrialSite", "PersonId", "VisitId", "VisitCycleNumber", "LockStatus"), _
                               Array("=", "=", "=", "=", "=", "="), _
                               Array(lStudyCode, sSiteCode, nSubjectId, CLng(sVisitId), CInt(sVisitCycleNumber), LockStatus.lsLocked)
    oQueryServer.SelectSave stUpdate, oQueryDef
    
    Set oQueryServer = Nothing
    Set oQueryDef = Nothing
    
    ' Make sure the Trial Subject is also unlocked
    Call UnlockTrialSubject(sDatabaseCnn, sDatabase, lStudyCode, sSiteCode, nSubjectId)

    UnlockVisitInstance = LockStatusResult.LockStatusPassed

Exit Function

ErrHandler:
    UnlockVisitInstance = LockStatusResult.LockStatusFailed
Exit Function


End Function

'--------------------------------------------------------------------------------------------------
Public Function UnlockCRFPageInstance(ByVal sDatabaseCnn As String, _
                                      ByVal sDatabase As String, _
                                      ByVal lStudyCode As Long, _
                                      ByVal sSiteCode As String, _
                                      ByVal nSubjectId As Integer, _
                                      ByVal sCRFPageTaskId As String, _
                                      ByVal sVisitId As String, _
                                      ByVal sVisitCycleNumber As String) As Integer
'--------------------------------------------------------------------------------------------------
' REM 03/08/01
' Unlock a CRF Page Instance if not frozen
' and ALSO unlock its Visit and Trial Subject
'--------------------------------------------------------------------------------------------------
Dim oQueryDef As QueryDef
Dim oQueryServer As QueryServer

    On Error GoTo ErrHandler
    
    Set oQueryDef = New QueryDef
    Set oQueryServer = New QueryServer

    oQueryServer.ConnectionOpen sDatabaseCnn
    
    oQueryDef.InitSave "CRFPageInstance", Array("LockStatus", "Changed"), _
                                          Array(LockStatus.lsUnlocked, Changed.Changed)
    oQueryDef.QueryFilters.Add Array("ClinicalTrialId", "TrialSite", "PersonId", "CRFPageTaskId", "LockStatus"), _
                               Array("=", "=", "=", "=", "="), _
                               Array(lStudyCode, sSiteCode, nSubjectId, CLng(sCRFPageTaskId), LockStatus.lsLocked)
    oQueryServer.SelectSave stUpdate, oQueryDef
    
    Set oQueryServer = Nothing
    Set oQueryDef = Nothing
    
    ' Make sure the Visit and Trial subject are also unlocked
    Call UnlockVisitInstance(sDatabaseCnn, sDatabase, lStudyCode, sSiteCode, nSubjectId, _
                            sVisitId, sVisitCycleNumber)
                            
    UnlockCRFPageInstance = LockStatusResult.LockStatusPassed

Exit Function

ErrHandler:
    UnlockCRFPageInstance = LockStatusResult.LockStatusFailed
Exit Function

End Function


'--------------------------------------------------------------------------------------------------
Public Function SQLToUpdateDataItemResponseHistory() As String
'--------------------------------------------------------------------------------------------------
' It copies a complete record from DataItemReponse into DataItemResponseHistory
' and only requires a WHERE clause to be added at the end
' NCJ 26/4/00 - Added in new validation/overrule fields
' NCJ 26/9/00 - Added in new NR/CTC fields
' NCJ 21/11/00 - Added in LaboratoryCode
' DPH 05/11/2002 - Brought into line with windows version
'--------------------------------------------------------------------------------------------------
Dim sSQL As String

    'SDM SR2611
    'Mo Morris 30/8/01 Db Audit (ReviewComment removed, UserId to UserName)
    sSQL = "INSERT INTO DataItemResponseHistory (" & _
           "ClinicalTrialId, " & _
           "TrialSite, " & _
           "PersonId, " & _
           "ResponseTaskId, " & _
           "ResponseTimestamp, " & _
           "VisitId, " & _
           "CRFPageId, " & _
           "CRFElementId, " & _
           "DataItemId, " & _
           "VisitCycleNumber, " & _
           "CRFPageCycleNumber, " & _
           "CRFPageTaskId, "
    sSQL = sSQL & _
           "ResponseValue, " & _
           "ResponseStatus, " & _
           "ValueCode, " & _
           "UserName, " & _
           "UnitOfMeasurement, " & _
           "Comments, " & _
           "Changed, " & _
           "SoftwareVersion, " & _
           "ReasonForChange, " & _
           "LockStatus, "
    ' NCJ 26/4/00
    sSQL = sSQL & _
            "ValidationId, ValidationMessage, OverruleReason, "
    ' NCJ 26/9/00
    ' NCJ 21/11/00 - Added LaboratoryCode
    ' DPH 22/01/2002 - Added RepeatNumber
    ' REM 11/07/02 - Added HadValue
    ' RS 25/10/2002 - Added ResponseTimestamp_TZ, ImportTimestamp_TZ, DatabaseTimestamp, DatabaseTimestamp_TZ
    sSQL = sSQL & _
            "LabResult, CTCGrade, ClinicalTestDate, LaboratoryCode, HadValue, RepeatNumber, "
    sSQL = sSQL & _
            "ResponseTimestamp_TZ, ImportTimestamp_TZ, DatabaseTimestamp, DatabaseTimestamp_TZ )"
    sSQL = sSQL & _
           "SELECT " & _
           "ClinicalTrialId, " & _
           "TrialSite, " & _
           "PersonId, " & _
           "ResponseTaskId, " & _
           "ResponseTimestamp, " & _
           "VisitId, " & _
           "CRFPageId, " & _
           "CRFElementId, " & _
           "DataItemId, " & _
           "VisitCycleNumber, " & _
           "CRFPageCycleNumber, " & _
           "CRFPageTaskId, "
    sSQL = sSQL & _
           "ResponseValue, " & _
           "ResponseStatus, " & _
           "ValueCode, " & _
           "UserName, " & _
           "UnitOfMeasurement, " & _
           "Comments, " & _
           "Changed, " & _
           "SoftwareVersion, " & _
           "ReasonForChange, " & _
           "LockStatus, "
    ' NCJ 26/4/00
    sSQL = sSQL & _
            "ValidationId, ValidationMessage, OverruleReason, "
    ' NCJ 26/9/00
    ' NCJ 21/11/00 - Added LaboratoryCode
    ' DPH 22/01/2002 - Added RepeatNumber
    'REM 11/07/02 - Added HadValue
    ' RS 25/10/2002 - Added ResponseTimestamp_TZ, ImportTimestamp_TZ, DatabaseTimestamp, DatabaseTimestamp_TZ
    sSQL = sSQL & _
            "LabResult, CTCGrade, ClinicalTestDate, LaboratoryCode, HadValue, RepeatNumber, "
    sSQL = sSQL & _
            "ResponseTimestamp_TZ, ImportTimestamp_TZ, DatabaseTimestamp, DatabaseTimestamp_TZ "
    sSQL = sSQL & _
           "FROM " & _
           "DataItemResponse "
    ' WHERE clause to be appended as required

    SQLToUpdateDataItemResponseHistory = sSQL

End Function


