VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSiteTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       clsSiteTransfer.cls
'   Author:     Richard Meinesz, September 2003
'   Purpose:    Class used to transfer a patient from one server site to another server site
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'   Revisions:

Option Explicit

Private Const msUPDATE_TABLES = "TrialSubject,VisitInstance,CRFPageInstance,DataItemResponse,DataItemResponseHistory,MIMessage,LFMessage,RSSubjectIdentifier,RSUniquenessCheck,Treatments"

'--------------------------------------------------------------------------------
Public Function PatientSiteTransfer(sConnection As String, sUsername As String, lClinicalTrialId As Long, sOriginalSite As String, _
                                    sNewSite As String, lPersonId As Long, ByRef sCouldNotLoadReason As String) As Boolean
'--------------------------------------------------------------------------------
'REM 17/09/03
'Transfer a Subject from one site to another.
'Routine locks selected subject, then create dummy lock while creating new subject Id
'then locks new subject and transfers subject to new site and updates AREZZO data
'--------------------------------------------------------------------------------
Dim sOriginalToken As String
Dim sDummyToken As String
Dim sNewToken As String
Dim bNeedToReloadSubject As Boolean
Dim bCouldLock As Boolean
Dim lNewPersonId As Long
Dim bTransferSubject As Boolean
    
    On Error GoTo ErrorLabel

    'Lock the selected subject
    'bCouldLock = LockForSave(sConnection, sCouldNotLoadReason, sOriginalToken, bNeedToReloadSubject, sUsername, _
                                        lClinicalTrialId, sOriginalSite, lPersonId)
    
    
    sOriginalToken = GetSubjectLock(sConnection, sUsername, lClinicalTrialId, sOriginalSite, lPersonId, sCouldNotLoadReason)
    
    If sOriginalToken = "" Then
        'there must be a study or subject lock, so can't transfer Subject
        PatientSiteTransfer = False
        Exit Function
    End If


    'Added LockForSave around subject being created (using dummy PersonID = -1)
    bCouldLock = LockForSave(sConnection, sCouldNotLoadReason, sDummyToken, bNeedToReloadSubject, sUsername, _
                                        lClinicalTrialId, sNewSite, -1)
                                                                   
    If Not bCouldLock Then
        'there must be a study lock for creating a new subject, so can't transfer Subject
        PatientSiteTransfer = False
        'Release the original subject lock
        Call RemoveSubjectLock(sConnection, lClinicalTrialId, sOriginalSite, lPersonId, sOriginalToken)
        Exit Function
    End If
                                        
    'Now that we have locked both the Patient that is to be transfered and a dummy patient so that no one else
    'can create a subject while we are doing site transfer, we can do the transfer
    
    'Get a new personId for the particulat study and new site
    lNewPersonId = GetNewPersonId(lClinicalTrialId, sNewSite)
    
'    'Added LockForSave around new subject
'    bCouldLock = LockForSave(sConnection, sCouldNotLoadReason, sNewToken, bNeedToReloadSubject, sUsername, _
'                                        lClinicalTrialId, sNewSite, lNewPersonId)
    
    sNewToken = GetSubjectLock(sConnection, sUsername, lClinicalTrialId, sNewSite, lNewPersonId, sCouldNotLoadReason)
    
    'release study/site lock for dummy personid = -1
    'we must have a token if we have got this far
    Call UnlockForSave(sConnection, sDummyToken, lClinicalTrialId, sNewSite, -1)

    If sNewToken = "" Then
        'there must be a study lock for the new subject
        PatientSiteTransfer = False
        'Release the original subject lock
        Call RemoveSubjectLock(sConnection, lClinicalTrialId, sOriginalSite, lPersonId, sOriginalToken)

        Exit Function
    End If
    
    'Transfer the Subject from the old site to the new one
    bTransferSubject = SiteTransfer(lClinicalTrialId, sOriginalSite, lPersonId, sNewSite, lNewPersonId, sCouldNotLoadReason)
    
    'Release the original subject lock
    Call RemoveSubjectLock(sConnection, lClinicalTrialId, sOriginalSite, lPersonId, sOriginalToken)
    
    'release the new subject lock
    Call RemoveSubjectLock(sConnection, lClinicalTrialId, sNewSite, lNewPersonId, sNewToken)
    
    'if transfer was success then update AREZZO data
    If bTransferSubject Then
        'Set AREZZO Data once subject has been transfered
        Call SetArezzoData(lClinicalTrialId, sNewSite, lNewPersonId, sCouldNotLoadReason)
    End If
    
    'set return to equal whether the transfer was successful
    PatientSiteTransfer = bTransferSubject

Exit Function
ErrorLabel:
     Err.Raise Err.Number, , Err.Description & "|clsSiteTransfer.PatientSiteTransfer"
End Function

'---------------------------------------------------------------------
Private Function SiteTransfer(lClinicalTrialId As Long, sOriginalSite As String, lOriginalPersonId As Long, _
                              sNewSite As String, lNewPersonId As Long, ByRef sCouldNotLoadReason As String) As Boolean
'---------------------------------------------------------------------
'REM 29/09/03
'Transfer a subject from one server site to another
'---------------------------------------------------------------------
Dim vTableName As Variant
Dim sTableName As String
Dim i As Integer
Dim bSiteTransfer As Boolean


    On Error GoTo ErrLabel
    
    'Routine must be inside a transaction
    'Begin transaction
    TransBegin
    
    vTableName = Split(msUPDATE_TABLES, ",")
    
    'loop through all the tables that must be updated and change the Site and SubjectId in each table
    For i = 0 To UBound(vTableName)
        sTableName = vTableName(i)
        'Update all the database table fields to new Site and PersonId
        bSiteTransfer = UpdateSubjectSite(sTableName, lClinicalTrialId, sOriginalSite, lOriginalPersonId, sNewSite, lNewPersonId, sCouldNotLoadReason)
        If Not bSiteTransfer Then
            GoTo ErrLabel
        End If
    Next
    
    'End transaction
    TransCommit
    
    'Successful transfer
    SiteTransfer = bSiteTransfer
    
Exit Function
ErrLabel:
    SiteTransfer = False
    If sCouldNotLoadReason = "" Then
        sCouldNotLoadReason = "An error occurred while trying to transfer subject! Error Description: " & Err.Description & " Error Number: " & Err.Number
    End If
    
    'Rollback transaction on error
    TransRollBack
    
End Function

'---------------------------------------------------------------------
Private Function UpdateSubjectSite(sTableName As String, lClinicalTrialId As Long, sOriginalSite As String, _
                                   lOriginalPersonId As Long, sNewSite As String, lNewPersonId As Long, _
                                   ByRef sCouldNotLoadReason As String) As Boolean
'---------------------------------------------------------------------
'REM 01/10/03
'Updates to the new PersonId and Site codes
'---------------------------------------------------------------------
Dim sSQL As String
Dim sClinicalTrialName As String
Dim rsTable As ADODB.Recordset
Dim lErrNo As Long

    On Error GoTo ErrLabel
    
    'check to see if the Treatments table exists before updating it
    If sTableName = "Treatments" Then
        On Error Resume Next
        sSQL = "SELECT ClinicalTrialId FROM Treatments"
        Set rsTable = New ADODB.Recordset
        rsTable.Open sSQL, MacroADODBConnection
        
        lErrNo = Err.Number
        Err.Clear
        
        On Error GoTo ErrLabel
        
        'then table does not exist so do not try and update it
        If lErrNo <> 0 Then
            UpdateSubjectSite = True
            Exit Function
        End If
        
    End If
    
    'get trial name from the Id
    sClinicalTrialName = TrialNameFromId(lClinicalTrialId)
    
    Select Case sTableName
    Case "MIMessage"
                
        If Not UpdateMIMessageTable(sClinicalTrialName, sOriginalSite, lOriginalPersonId, sNewSite, lNewPersonId) Then
            GoTo ErrLabel
        End If
                
    Case "RSSubjectIdentifier", "RSUniquenessCheck"
        
        sSQL = "UPDATE " & sTableName & " SET TrialSite = '" & sNewSite & "', " _
                     & " PersonId = " & lNewPersonId _
                     & " WHERE ClinicalTrialName = '" & sClinicalTrialName & "'" _
                     & " AND TrialSite = '" & sOriginalSite & "' " _
                     & " AND PersonId = " & lOriginalPersonId
        MacroADODBConnection.Execute sSQL, adOpenKeyset, adCmdText
    Case Else
    
        sSQL = "UPDATE " & sTableName & " SET TrialSite = '" & sNewSite & "', " _
                     & " PersonId = " & lNewPersonId _
                     & " WHERE ClinicalTrialId = " & lClinicalTrialId _
                     & " AND TrialSite = '" & sOriginalSite & "' " _
                     & " AND PersonId = " & lOriginalPersonId
        MacroADODBConnection.Execute sSQL, adOpenKeyset, adCmdText
    End Select
    
    
    sCouldNotLoadReason = ""
    UpdateSubjectSite = True
         
Exit Function
ErrLabel:
    UpdateSubjectSite = False
    sCouldNotLoadReason = "Error updating table " & sTableName & ", Error Description: " & Err.Description & ", Error Number: " & Err.Number
End Function

'---------------------------------------------------------------------
Private Function UpdateMIMessageTable(sClinicalTrialName As String, sOriginalSite As String, lOriginalPersonId As Long, _
                                      sNewSite As String, lNewPersonId As Long) As Boolean
'---------------------------------------------------------------------
'REM 08/10/03
'Update the MIMessage table with new Site
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsMIMsg As ADODB.Recordset
Dim vMIMsg As Variant
Dim nMIMsgSource As Integer
Dim lMaxMIMessageId As Long
Dim lMIMessageId As Long
Dim nMaxMIMessageObjectId As Integer
Dim nMIMessageObjectId As Integer
Dim nNextMIMessageObjectId As Integer
Dim nMIMsgObjectSource As Integer
Dim i As Long
        
    On Error GoTo ErrLabel
    
    'get all MIMessages that match the study, site and subject
    sSQL = "SELECT * FROM MIMessage " _
         & " WHERE MIMessageTrialName = '" & sClinicalTrialName & "' " _
         & " AND MIMessageSite = '" & sOriginalSite & "' " _
         & " AND MIMessagePersonId = " & lOriginalPersonId _
         & " ORDER BY MIMESSAGEOBJECTID"
    Set rsMIMsg = New ADODB.Recordset
    rsMIMsg.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsMIMsg.RecordCount > 0 Then
        vMIMsg = rsMIMsg.GetRows
        
        'close recordset
        rsMIMsg.Close
        Set rsMIMsg = Nothing
        
        'loop through all MIMessage rows
        For i = 0 To UBound(vMIMsg, 2)
        
            nMIMsgSource = vMIMsg(2, i)
            lMIMessageId = vMIMsg(0, i)
            nNextMIMessageObjectId = vMIMsg(5, i)
            nMIMsgObjectSource = vMIMsg(6, i)
            
            'get new MIMessageId for every update
            lMaxMIMessageId = GetNewMIMessageId(sNewSite, nMIMsgSource)
            
            'Only get a new MIMessageObjectId if its is not the same as the last one in the loop
            If nNextMIMessageObjectId <> nMIMessageObjectId Then
                nMIMessageObjectId = GetNewMIMessageObjectId(sNewSite, nMIMsgObjectSource)
            Else
                nMIMessageObjectId = nNextMIMessageObjectId
            End If
            
            'Update the MIMessage table
            sSQL = "UPDATE MIMESSAGE SET MIMessageId = " & lMaxMIMessageId & "," _
                & " MIMessageSite = '" & sNewSite & "'," _
                & " MIMessageObjectId = " & nMIMessageObjectId & "," _
                & " MIMessagePersonId = " & lNewPersonId _
                & " WHERE MIMessageTrialName = '" & sClinicalTrialName & "' " _
                & " AND MIMessageId = " & lMIMessageId _
                & " AND MIMessageSite = '" & sOriginalSite & "' " _
                & " AND MIMessageSource = " & nMIMsgSource
            MacroADODBConnection.Execute sSQL, adOpenKeyset, adCmdText
        
        Next
        UpdateMIMessageTable = True
    Else
        UpdateMIMessageTable = True
        rsMIMsg.Close
        Set rsMIMsg = Nothing
    End If
    
Exit Function
ErrLabel:
    UpdateMIMessageTable = False

End Function

'---------------------------------------------------------------------
Private Function GetNewMIMessageId(sNewSite As String, nMIMsgSource As Integer) As Long
'---------------------------------------------------------------------
'REM 09/10/03
'Returns the max MIMessageId + 1
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsMaxId As ADODB.Recordset

    sSQL = "SELECT MAX(MIMessageId) as MaxMIMessageId FROM MIMessage " _
        & " WHERE MIMessageSite = '" & sNewSite & "'" _
        & " AND MIMessageSource = " & nMIMsgSource
    Set rsMaxId = New ADODB.Recordset
    rsMaxId.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

    If IsNull(rsMaxId!MaxMIMessageId) Then
        GetNewMIMessageId = 1
    Else
        GetNewMIMessageId = rsMaxId!MaxMIMessageId + 1
    End If
    
    rsMaxId.Close
    Set rsMaxId = Nothing
    
End Function

'---------------------------------------------------------------------
Private Function GetNewMIMessageObjectId(sNewSite As String, nMIMsgObjectSource As Integer) As Integer
'---------------------------------------------------------------------
'REM 09/10/03
'Returns max MIMessageObjectId + 1
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsMaxObjectId As ADODB.Recordset

    sSQL = "SELECT MAX(MIMessageObjectId) as MaxMIMessageObjectId FROM MIMessage " _
        & " WHERE MIMessageSite = '" & sNewSite & "'" _
        & " AND MIMessageObjectSource = " & nMIMsgObjectSource
    Set rsMaxObjectId = New ADODB.Recordset
    rsMaxObjectId.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

    If IsNull(rsMaxObjectId!MaxMIMessageObjectId) Then
        GetNewMIMessageObjectId = 1
    Else
        GetNewMIMessageObjectId = rsMaxObjectId!MaxMIMessageObjectId + 1
    End If

    rsMaxObjectId.Close
    Set rsMaxObjectId = Nothing

End Function


'---------------------------------------------------------------------
Private Function GetNewPersonId(lClinicalTrialId As Long, sNewSite As String) As Long
'---------------------------------------------------------------------
'REM 29/09/03
'Returns a new person id for a given study and site combination
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsMaxPersonId As ADODB.Recordset
    
    On Error GoTo ErrorLabel
    
    sSQL = "SELECT MAX(PersonId) as MaxPersonId FROM TrialSubject " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND TrialSite = '" & sNewSite & "'"
    Set rsMaxPersonId = New ADODB.Recordset
    rsMaxPersonId.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If IsNull(rsMaxPersonId!MaxPersonId) Then
        GetNewPersonId = 1
    Else
        GetNewPersonId = rsMaxPersonId!MaxPersonId + 1
    End If
    
Exit Function
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "clsSiteTransfer.GetNewPersonId"
End Function

'---------------------------------------------------------------------
Private Function LockForSave(sCon As String, ByRef sLockErrMsg As String, ByRef sToken As String, ByRef bNeedToReloadSubject As Boolean, _
                                sUser As String, lClinicalTrialId As Long, sSite As String, lPersonId As Long) As Boolean
'---------------------------------------------------------------------
' Lock a subject.
' In:
'   sUser: username to lock subject with
'   lClinicalTrialId, sSite, lPersonID: key of subject to lock
' Out:
'   function: success/failure
'   sToken: token if lock successful or "" if not
'   sLockErrMsg: the reason the lock failed or "" if successful
'   nNeedToReloadSubject: if the lock is successful this returns whehther the subject needs reloading
'---------------------------------------------------------------------
Dim sLockDetails As String
Const sSTUDY_BEING = "This study is currently being "
Const sSUBJECT_BEING = "This subject is currently being "
Const sEDITED = "edited by "
Const sSAVED = "saved by "
Const sANOTHER_USER = "another user"

    On Error GoTo ErrorLabel
    
    'set initial output variables to failure (change it later if we have success)
    sToken = ""
    sLockErrMsg = ""
    'set initial function output to failure (change it later if we have success)
    LockForSave = False
    
    'TA 04.07.2001: use new locking
    sToken = MACROLOCKBS30.LockSubjectForSaving(sCon, sUser, lClinicalTrialId, sSite, lPersonId)
    Select Case sToken
    Case MACROLOCKBS30.DBLocked.dblStudy
        sLockDetails = MACROLOCKBS30.LockDetailsStudy(sCon, lClinicalTrialId)
        sLockErrMsg = sSTUDY_BEING & sEDITED
        If sLockDetails = "" Then
            sLockErrMsg = sLockErrMsg & sANOTHER_USER
        Else
            sLockErrMsg = sLockErrMsg & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblSubject
        sLockDetails = MACROLOCKBS30.LockDetailsSubject(sCon, lClinicalTrialId, sSite, lPersonId)
        sLockErrMsg = sSUBJECT_BEING & sEDITED
        If sLockDetails = "" Then
            sLockErrMsg = sLockErrMsg & sANOTHER_USER
        Else
            sLockErrMsg = sLockErrMsg & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblEFormInstance
        sLockDetails = MACROLOCKBS30.LockDetailsSubjectSave(sCon, lClinicalTrialId, sSite, lPersonId)
        sLockErrMsg = sSUBJECT_BEING & sSAVED
        If sLockDetails = "" Then
            sLockErrMsg = sLockErrMsg & sANOTHER_USER
        Else
            sLockErrMsg = sLockErrMsg & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case Else
        'we have a lock we need to check wether it is invalid to decide whether to reload or not
        bNeedToReloadSubject = Not MACROLOCKBS30.CacheEntryStillValid(sCon, sToken)
        'function to return success
        LockForSave = True
    End Select
    
Exit Function
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "clsSiteTransfer.LockForSave"
End Function

'---------------------------------------------------------------------
Private Sub UnlockForSave(sCon As String, sToken As String, _
                            lClinicalTrialId As Long, sSite As String, lPersonId As Long)
'---------------------------------------------------------------------
' Unlock the subject
'---------------------------------------------------------------------

    On Error GoTo ErrorLabel
    
    'TA 04.07.2001: use new locking model
    If sToken <> "" Then
        'if no gsStudyToken then UnlockSubject is being called without a corresponding LockSubject being called first
        MACROLOCKBS30.UnlockSubjectForSaving sCon, sToken, lClinicalTrialId, sSite, lPersonId
        'always set this to empty string for same reason as above
        sToken = ""
    End If

Exit Sub
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "clsSiteTransfer..UnlockForSave"
End Sub

'---------------------------------------------------------------------
Private Function LoadStudyDef(ByVal lStudyId As Long, oArezzo As Arezzo_DM, ByRef sErrMsg As String) As StudyDefRO
'---------------------------------------------------------------------
' Load the study def
' Returns StudyDef if study successfully loaded, else returns Nothing
'---------------------------------------------------------------------
Dim oStudyDef As StudyDefRO
    
    On Error GoTo ErrLabel

    Set oStudyDef = New StudyDefRO
    sErrMsg = oStudyDef.Load(goUser.CurrentDBConString, lStudyId, 1, oArezzo)
    If sErrMsg > "" Then
        ' Give up
        
        LoadStudyDef = Nothing
    End If

    Set LoadStudyDef = oStudyDef
    
Exit Function
ErrLabel:
    LoadStudyDef = Nothing
    sErrMsg = "Error loading study definition. Error Description: " & Err.Description & ", Error Number: " & Err.Number
End Function

'---------------------------------------------------------------------
Private Function SetArezzoData(lStudyId As Long, sNewSite As String, lNewSubjectId As Long, _
                               ByRef sErrMsg As String) As Boolean
'---------------------------------------------------------------------
'REM 30/09/03
'Update the AREZZO state so it has the new Site:Subject combination
'---------------------------------------------------------------------
Dim oStudyDef As StudyDefRO
Dim oArezzo As Arezzo_DM
Dim oArezzoMemory As clsAREZZOMemory
Dim sCountry As String
    
    On Error GoTo ErrLabel

    'Create and initialise a new Arezzo instance
    Set oArezzo = New Arezzo_DM
    
    ' NCJ 29 Jan 03 - Get Prolog switches from new ArezzoMemory class
    Set oArezzoMemory = New clsAREZZOMemory
    Call oArezzoMemory.Load(0, goUser.CurrentDBConString)
    Call oArezzo.Init(gsTEMP_PATH, oArezzoMemory.AREZZOSwitches)
    Set oArezzoMemory = Nothing
    
    Set oStudyDef = LoadStudyDef(lStudyId, oArezzo, sErrMsg)
    
    If Not oStudyDef Is Nothing Then
        Call oStudyDef.LoadSubject(sNewSite, lNewSubjectId, goUser.UserName, eUIUpdateMode.Read_Write, _
                                    goUser.UserNameFull, goUser.UserRole)
        If oStudyDef.Subject.CouldNotLoad Then
            ' return reason it could not load
            sErrMsg = oStudyDef.Subject.CouldNotLoadReason
            'Return false
            SetArezzoData = False
        Else
            ' Successfully loaded
            Call oStudyDef.Subject.Arezzo.AddData("person:personid", "")
            Call oStudyDef.Subject.Arezzo.AddData("person:personid", CStr(lNewSubjectId))
            Call oStudyDef.Subject.Arezzo.AddData("person:trialsite", "")
            Call oStudyDef.Subject.Arezzo.AddData("person:trialsite", sNewSite)
            'Change the country to the new sites one
            sCountry = goUser.GetAllSites.Item(sNewSite).CountryName
            Call oStudyDef.Subject.Arezzo.AddData("person:country", "")
            Call oStudyDef.Subject.Arezzo.AddData("person:country", sCountry)
            
            Call oStudyDef.Subject.Save
            ' Clear up the subject
            Call oStudyDef.RemoveSubject
            Call oStudyDef.Terminate
            
            SetArezzoData = True
        End If
    Else 'error while loading study
        SetArezzoData = False
        
    End If
    
    Set oStudyDef = Nothing
    
    ' Only shut down the ALM if it has been started
    If Not oArezzo Is Nothing Then
        oArezzo.Finish
        Set oArezzo = Nothing
    End If
    
Exit Function
ErrLabel:
    SetArezzoData = False
    sErrMsg = "Error while setting the AREZZO Data. Error Description: " & Err.Description & ", Error Number: " & Err.Number
End Function

'---------------------------------------------------------------------
Private Function GetSubjectLock(sConnnection As String, sUser As String, lStudyId As Long, sSite As String, lSubjectId As Long, ByRef sMessage As String) As String
'---------------------------------------------------------------------
Dim sMsg As String
Dim sToken As String
    
    On Error GoTo LockFailed

    sToken = MACROLOCKBS30.LockSubject(sConnnection, sUser, lStudyId, sSite, lSubjectId)
    Select Case sToken
    Case MACROLOCKBS30.DBLocked.dblStudy
        sToken = ""
        sMessage = "Another user is currently editing the study definition."
        
    Case MACROLOCKBS30.DBLocked.dblSubject, MACROLOCKBS30.DBLocked.dblEFormInstance
        sToken = ""
        sMessage = "Another user currently has this subject open for data entry."
        
    Case Else
        'hurrah, we have a lock
    End Select
    
    GetSubjectLock = sToken
    
Exit Function
LockFailed:
    Err.Raise Err.Number, , Err.Description & "|" & "clsSiteTransfer.GetSubjectLock"
End Function

'---------------------------------------------------------------------
Private Sub RemoveSubjectLock(sConnnection As String, lStudyId As Long, sSite As String, lSubjectId As Long, sToken As String)
'---------------------------------------------------------------------
On Error GoTo LockFailed

    Call MACROLOCKBS30.UnlockSubject(sConnnection, sToken, lStudyId, sSite, lSubjectId)
    
Exit Sub
LockFailed:
    ' Ignore Errors
End Sub
