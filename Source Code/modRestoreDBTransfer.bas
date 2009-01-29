Attribute VB_Name = "modRestoreDBTransfer"
'------------------------------------------------------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       modRestoreDBTransfer.bas
'   Author:     Will Casey    27  March 2000
'   Purpose:    For restoring databases in the event that a site laptop is lost as
'               requested in SR2727
'   Revisions:
'
'   WillC 18/5/00   Added DoMIMessage to the patient data as there is a new table
'                   in the database.
'   WillC 14/6/00  SR3590 Having memory problems while load testing on a database with 3/4 of a million records in the DataItemResponse
'                   and DataItemResponseHistory tables so refined the search to bring back records person by person...
'   WillC 19/6/00   SR3590 Added a small cache size to DoDataItemResponse, DoDataItemResponseHistory and DoCRFPageInstance and changed the cursor location
'                   to adUseServer to free up memory on the client. Have retested this and it ran through the 3/4 million record database properly.
'   TA 17/10/2000: New Version 2.1.4 fields and tables added
'                   DoClinicalTestClincalTestGroup, DoCTCSchemeCTC, DoLaboratoryNRSiteLaboratory created for new tables
'                   DoStudyDefintion, DoCRFElement, DoCRFPageInstance, DoTrialSubject, DoDataItemResponse, DoDataItemResponseHistory altered for new fields
'   TA 18/10/2000: SiteUser, Units and UnitConversionFactors tables now transfered
'   TA 06/12/2000: New columns for registration added (new tables to be done)
'   TA 12/12/2000: Begrudgingly added new tables for registration - Eligibility, UniquenessCheck and SubjectNumbering
'   Ash 30/08/2001: Changed Routines to function Generically
'   Ash 07/09/2001: Deleted most of the old routines.Left a couple for cross reference
'   Ash 06/02/2002:  Added IsDBVersionValid to check databases versions match
'   Ash 17/04/2002:  Commented out rsToTable.CursorLocation = adUseClient in all routines
'   DPH 06/12/2002: Added missing MIMessage functionality to DBRestore
'   REM 03/12/2003: In the CopySTYDEFFromMacroTable routine added RestoreProtocolsTable as this table requires being handles differently
'   REM 03/12/2003: Tidied up some old commented out routines
'---------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Private CopyDataADODBConnection As ADODB.Connection
Private mfldMemoFrom As ADODB.Field
Private mfldMemoTo As ADODB.Field
Private mrsTrFrom As ADODB.Recordset
Private mrsTrTo As ADODB.Recordset
Private msSQLfrom As String
Private msSQLto As String
Private mlTotalRecordCount As Long
Private mlRecordCounter  As Long
Private mlTargetTrialId As Long
Private l As Long
Private vFields As Variant
Private vValues As Variant
'Ash 30/08/2001
Private msTrialName As String
Private msSite As String
Private mlTrialID As Long
Private msDatabaseName As String

'--------------------------------------------------------------------------------
Private Sub InitializeCopyDataADODBConnection(sConnection As String)
'--------------------------------------------------------------------------------
'This will initialize the ADO connection to the 2.0x  database
'--------------------------------------------------------------------------------
' DPH 06/12/2002 - Use passed in db password
'--------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Set CopyDataADODBConnection = New ADODB.Connection
        
    CopyDataADODBConnection.Open sConnection
    CopyDataADODBConnection.CursorLocation = adUseClient
      
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "InitializeCopyDataAdodbConnection", "modRestoreDBTransfer")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
   
End Sub

'---------------------------------------------------------------------
Public Sub DoDataTransfer(sConnection As String, ByRef bRestoreSecurityData As Boolean)
'---------------------------------------------------------------------
'check to see If a trial has been imported already if so use the existing ClinicalTrialId
'and just do the DoPatientDataTransfer if not get a new mlTargetTrialId
'then DoStudyDefinitionTransfer , then DoPatientDataTransfer
'---------------------------------------------------------------------
' REVISIONS
' DPH 06/12/2002 - Added in collecting DB password
'---------------------------------------------------------------------
Dim sMsg As String
Dim rsDoesStudyExist As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    msTrialName = frmRestoreSiteDatabase.RestoreTrialName
    msSite = frmRestoreSiteDatabase.RestoreTrialSite
    mlTrialID = frmRestoreSiteDatabase.RestoreTrialId
    msDatabaseName = frmRestoreSiteDatabase.RestoreDBName
    If msDatabaseName = "" Then
        msDatabaseName = frmRestoreSiteDatabase.RestoreDataSource
    End If
    
    ' DPH 06/12/2002 - Use database password collected on form
    Call InitializeCopyDataADODBConnection(sConnection)
    
    If Not IsDBVersionValid(CopyDataADODBConnection, MacroADODBConnection) Then
        Exit Sub
    End If
    ' Check to see if the Study/Site combination has been done already if so disallow.
    If HasStudySiteBeenRestored = True Then
        Exit Sub
    End If

    sSQL = "Select * from ClinicalTrial where ClinicalTrialName = '" & msTrialName & "'"
    Set rsDoesStudyExist = New ADODB.Recordset
    rsDoesStudyExist.Open sSQL, CopyDataADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    'WillC 29/4/00 Changed to HourglassOn
    HourglassOn
   
   'Ash 31/08/2001 transaction processing to maintain data intergrity
   
   CopyDataADODBConnection.BeginTrans
    
        'If the studydef already exists in target db then just
        'copy patient data using the existing TrialId
        If rsDoesStudyExist.RecordCount = 1 Then
            mlTargetTrialId = rsDoesStudyExist!ClinicalTrialId
            Call CopyPATRSPFromMacroTable 'ash
            Call RestoreExtraMacroTableSiteData 'ash
        Else
            'The study doesn't exist get a new TargetTrialId and copy over
            'the studydef and patient data
            Call GetRestoreTrialId
            Call CopySTYDEFFromMacroTable 'ash
            Call CopyPATRSPFromMacroTable 'ash
            Call RestoreExtraMacroTableSiteData 'ash
        End If
        
        If Not bRestoreSecurityData Then
            'REM 10/02/03 - restore the security database
            Call RestoreSecurityDatabase
            bRestoreSecurityData = True
        End If
        
        Set rsDoesStudyExist = Nothing
       
        
        'Inform the user the transfer is done
        'WillC 29/4/00 Changed to HourglassOff
        HourglassOff
        
        sMsg = "The restoration of study " & frmRestoreSiteDatabase.RestoreTrialName & vbCrLf
        sMsg = sMsg & "for the site " & frmRestoreSiteDatabase.RestoreTrialSite & vbCrLf
        sMsg = sMsg & "has completed successfully." & vbCrLf & vbCrLf
        sMsg = sMsg & Format(Now, "dd mmm yyyy hh:mm:ss")
        Call MsgBox(sMsg, vbInformation, "MACRO Restore Database")
 
 CopyDataADODBConnection.CommitTrans
 
 Set CopyDataADODBConnection = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "DoDataTransfer", "modRestoreDBTransfer")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub GetRestoreTrialId()
'---------------------------------------------------------------------
'Find the max TrialId in the database we are sending the data to and increment it.
'---------------------------------------------------------------------
Dim rsNextTrialId As ADODB.Recordset
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    sSQL = "Select Max(ClinicalTrialId) as MaxClinicalTrialId from ClinicalTrial"
    Set rsNextTrialId = New ADODB.Recordset
    rsNextTrialId.Open sSQL, CopyDataADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    mlTargetTrialId = rsNextTrialId!MaxClinicalTrialId + 1
    Set rsNextTrialId = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "GetRestoreTrialId", "modRestoreDBTransfer")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
   
End Sub

'---------------------------------------------------------------------
Private Function HasStudySiteBeenRestored() As Boolean
'---------------------------------------------------------------------
' check to see if a study at a site has been previously imported
' if so disallow it.
'---------------------------------------------------------------------
Dim sClinicalTrialName As String
Dim lClinicalTrialId As Long
Dim sTrialSite As String
Dim sSQL As String
Dim rsHasStudySiteBeenRestored As ADODB.Recordset
Dim rsTrialName As ADODB.Recordset

    On Error GoTo ErrHandler
    
    HasStudySiteBeenRestored = False
    
    sSQL = " Select * " _
    & " FROM  TrialSubject WHERE TrialSite = '" & msSite & "'"
  
    'Find out if a study at the site has been restored already and get the trialId
    Set rsHasStudySiteBeenRestored = New ADODB.Recordset
    rsHasStudySiteBeenRestored.Open sSQL, CopyDataADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    'Make sure we have a valid record to do the comparison on
    If rsHasStudySiteBeenRestored.RecordCount > 0 Then
  
        'Get the site and Id for the chosen study
        sTrialSite = rsHasStudySiteBeenRestored!TrialSite
        lClinicalTrialId = rsHasStudySiteBeenRestored!ClinicalTrialId
        Set rsTrialName = New ADODB.Recordset
        
        'Get the study name using the trial id fom ClinicalTrial
        sSQL = " Select ClinicalTrialName FROM ClinicalTrial WHERE ClinicalTrialId = " & lClinicalTrialId
        rsTrialName.Open sSQL, CopyDataADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        sClinicalTrialName = rsTrialName!ClinicalTrialName
        
        'Compare the Site and study name (as we cant use the trialId) to those chosen in the form and do a comparison on if they match
        'if so set HasStudySiteBeenRestored = True
        If sTrialSite = msSite And sClinicalTrialName = msTrialName Then
           MsgBox "This study at this site has already been restored.", vbInformation, "MACRO"
           HasStudySiteBeenRestored = True
        Else
           HasStudySiteBeenRestored = False
        End If
      
    End If
  
    Set rsTrialName = Nothing
    Set rsHasStudySiteBeenRestored = Nothing

Exit Function
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "HasStudySiteBeenRestored", "modRestoreDBTransfer")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select


End Function

'--------------------------------------------------------------------------------------------------------
Private Sub CopySTYDEFFromMacroTable()
'--------------------------------------------------------------------------------------------------------
'ASH  30/08/2001 To make current routine generic
'REVISIONS:
' REM 03/12/03 - Added RestoreProtocolsTable as this table requires being handles differently
'--------------------------------------------------------------------------------------------------------
Dim sSQL As String
Dim sTableName As String
Dim sSegmentId As String
Dim nSegmentId As Integer
Dim rsTotalTables As ADODB.Recordset

    On Error GoTo ErrHandler
    
    sSQL = "SELECT Tablename, Segmentid From MacroTable Where STYDEF = 1"
    Set rsTotalTables = New ADODB.Recordset
    rsTotalTables.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
      rsTotalTables.MoveFirst
   
    Do Until rsTotalTables.EOF
        sTableName = rsTotalTables![TableName]
        sSegmentId = rsTotalTables![SegmentId]
        nSegmentId = CInt(sSegmentId)
        
        Select Case nSegmentId
        Case Is < 300
            'copy all trial specific info
            Call CopyTrialIDRecordsToRestore(sTableName, mlTrialID)
        Case Is = 300
            'protocols table handled separatly
            Call RestoreProtocolsTable
        Case Is > 300
            'tables from SegmentId 510 - 640.Data not based on a particular studyID
            Call RestoreAllDataInTable(sTableName)
        End Select
        rsTotalTables.MoveNext
    Loop

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CopySTYDEFFromMacroTable", "modRestoreDBTransfer")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub
'----------------------------------------------------------------------------------
Private Sub CopyPATRSPFromMacroTable()
'----------------------------------------------------------------------------------
'ASH 30/08/2001 Routine that restores the Patient Response Data
'----------------------------------------------------------------------------------
Dim sSQL As String
Dim msTablename As String
Dim msSegmentId As String
Dim rsTotalTables As ADODB.Recordset

    On Error GoTo ErrHandler
    
    sSQL = "SELECT Tablename, segmentid From macrotable Where PATRSP = 1"
    Set rsTotalTables = New ADODB.Recordset
    rsTotalTables.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
        
    Do Until rsTotalTables.EOF
        msTablename = rsTotalTables![TableName]
        msSegmentId = rsTotalTables![SegmentId]
            
        Call CopyTrialSiteRecordsToRestore(msTablename)
        
        rsTotalTables.MoveNext
          
    Loop
        
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CopyPATRSPFromMacroTable", "modRestoreDBTransfer")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'------------------------------------------------------------------------------------------------------
Private Sub RestoreProtocolsTable()
'------------------------------------------------------------------------------------------------------
' REM 03/12/03
' Restores Protocols table
'------------------------------------------------------------------------------------------------------
Dim sFromSQL As String
Dim sToSQL As String
Dim i As Long
Dim j As Long
Dim rsFromTable As New ADODB.Recordset
Dim rsToTable As New ADODB.Recordset
     
     On Error GoTo ErrHandler
     
    'creates recordset to contain records to be restored
    sFromSQL = "SELECT * FROM PROTOCOLS" _
    & " WHERE FileName = '" & msTrialName & "'"

    sToSQL = "SELECT * FROM Protocols" _
    & " WHERE 1 = 0"

    Set rsFromTable = New ADODB.Recordset
    rsFromTable.Open sFromSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    Set rsToTable = New ADODB.Recordset
    rsToTable.Open sToSQL, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    If rsFromTable.RecordCount <= 0 Then
        Exit Sub
    End If

    'move to first record in recordset
    rsFromTable.MoveFirst
    'loop to copy records into new trialID
        For j = 1 To rsFromTable.RecordCount
            rsToTable.AddNew
            rsToTable.Fields(0) = msTrialName
                For i = 1 To rsFromTable.Fields.Count - 1
                    rsToTable.Fields(i).Value = rsFromTable.Fields(i).Value
                 Next
        rsToTable.Update
        rsFromTable.MoveNext
        Next j

    rsFromTable.Close
    Set rsFromTable = Nothing
    rsToTable.Close
    Set rsToTable = Nothing
        
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreProtocolsTable", "modRestoreDBTransfer")
                                    
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------------------------------------------------------------------
Private Sub CopyTrialIDRecordsToRestore(ByVal msTablename As String, _
                                    ByVal mlTrialID As Long)
'------------------------------------------------------------------------------------------------------
'Restores records to be restored based on ClinicalTrialIDs
'------------------------------------------------------------------------------------------------------
Dim sFromSQL As String
Dim sToSQL As String
Dim i As Long
Dim j As Long
Dim rsFromTable As New ADODB.Recordset
Dim rsToTable As New ADODB.Recordset
     
     On Error GoTo ErrHandler
     
    'creates recordset to contain records to be restored
    sFromSQL = "SELECT * FROM " & msTablename _
    & " WHERE ClinicalTrialID = " & mlTrialID

    'REM 31/01/03 - changed true = false to 1 = 0
    sToSQL = "SELECT * FROM " & msTablename _
    & " WHERE 1 = 0"

'    sToSQL = "Select * from " & msTablename _
'    & " Where ClinicalTrialID = -1"

    Set rsFromTable = New ADODB.Recordset
    rsFromTable.Open sFromSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    Set rsToTable = New ADODB.Recordset
    'ASH 16/04/2002
    'commented out to fix SR 4271 in Macro 2.2/3.0
    'rsToTable.CursorLocation = adUseClient
    rsToTable.Open sToSQL, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    If rsFromTable.RecordCount <= 0 Then
        Exit Sub
    End If

    'move to first record in recordset
    rsFromTable.MoveFirst
    'loop to copy records into new trialID
        For j = 1 To rsFromTable.RecordCount
            rsToTable.AddNew
            rsToTable.Fields(0) = mlTargetTrialId
                For i = 1 To rsFromTable.Fields.Count - 1
                    rsToTable.Fields(i).Value = rsFromTable.Fields(i).Value
                 Next
        rsToTable.Update
        rsFromTable.MoveNext
        Next j

    rsFromTable.Close
    Set rsFromTable = Nothing
    rsToTable.Close
    Set rsToTable = Nothing
        
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CopyTrialIDRecordsToRestore", "modRestoreDBTransfer")
                                    
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub
'----------------------------------------------------------------------------------------------
Private Sub CopyTrialSiteRecordsToRestore(ByVal msTablename As String)
'----------------------------------------------------------------------------------------------
'Restores data based on Trialsite and StudyId
'----------------------------------------------------------------------------------------------
Dim sFromSQL As String
Dim sToSQL As String
Dim i As Long
Dim j As Long
Dim rsFromTable As New ADODB.Recordset
Dim rsToTable As New ADODB.Recordset
     
    On Error GoTo ErrHandler
  
    'creates recordset to contain records to be copied
    sFromSQL = "Select * from " & msTablename _
    & " Where ClinicalTrialID = " & mlTrialID _
    & " AND TrialSite = '" & msSite & "'"

    sToSQL = "Select * from " & msTablename _
    & " where 0 = 1"

    Set rsFromTable = New ADODB.Recordset
    rsFromTable.Open sFromSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    Set rsToTable = New ADODB.Recordset
    'ASH 16/04/2002
    'commented out to fix SR 4271 in Macro 2.2/3.0
    'rsToTable.CursorLocation = adUseClient
    rsToTable.Open sToSQL, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    If rsFromTable.RecordCount <= 0 Then
        Exit Sub
    End If

    'move to first record inrecordset
    rsFromTable.MoveFirst
    'loop to copy records into new trialID
    For j = 1 To rsFromTable.RecordCount
        rsToTable.AddNew
        rsToTable.Fields(0) = mlTargetTrialId
            For i = 1 To rsFromTable.Fields.Count - 1
                rsToTable.Fields(i).Value = rsFromTable.Fields(i).Value
            Next
        rsToTable.Update
        rsFromTable.MoveNext
    Next j

    rsFromTable.Close
    Set rsFromTable = Nothing
    rsToTable.Close
    Set rsToTable = Nothing
        
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CopyTrialSiteRecordsToRestore", "modRestoreDBTransfer")
                                    
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub
'------------------------------------------------------------------
Private Sub RestoreExtraMacroTableSiteData()
'------------------------------------------------------------------
'Calls routines that deal with site information.These routines
'are treated outside the main 2 routines because they do not
'have identifiers in the MacroTable or because they do not warrant
'a generic routine since they are only single tables
'-------------------------------------------------------------------
    
    Call RestoreTrialSites
    Call RestoreSiteDatas
    'REM 10/02/03 - no longer a SiteUser Table has been replaced by the UserRole table
    Call RestoreUserRoleData
    
    Call RestoreSiteLaboratoryData
    Call RestoreLaboratoryData
    Call RestoreNormalRangeData
    ' DPH 12/12/2002 - Restore MIMessage data SR4980
    Call RestoreMIMessages

End Sub

'----------------------------------------------------------------------------------
Private Sub RestoreSecurityDatabase()
'----------------------------------------------------------------------------------
'REM 10/02/03
'Restore the security database
'----------------------------------------------------------------------------------

    Call RestoreMACROUsers
    Call RestoreUserDatabases
    Call RestoreMACROPassword
    Call RestoreRoles
    Call RestoreRoleFunctions

End Sub

'----------------------------------------------------------------------------------
Private Sub RestoreRoleFunctions()
'----------------------------------------------------------------------------------
'REM 11/02/03
'Restore the Role Functions on a site
'----------------------------------------------------------------------------------
Dim rsRoleFunctions As ADODB.Recordset
Dim rsCopyRoleFunctions As ADODB.Recordset
Dim i As Long
Dim j As Long
Dim sSQL As String
Dim sSQL1 As String

    On Error GoTo ErrHandler
     
    CopyDataADODBConnection.Execute "DELETE FROM RoleFunction"

    'creates recordset to contain records to be copied
    sSQL = "SELECT * FROM RoleFunction "
    Set rsRoleFunctions = New ADODB.Recordset
    rsRoleFunctions.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'creates recordset to contain records to be copied
    sSQL1 = "SELECT * FROM RoleFunction " _
    & " WHERE 0 = 1"
    Set rsCopyRoleFunctions = New ADODB.Recordset
    rsCopyRoleFunctions.Open sSQL1, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'checks if records exist
    If rsRoleFunctions.RecordCount <= 0 Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    'move to first record inrecordset
    rsRoleFunctions.MoveFirst

    'begin record insertion
    For j = 1 To rsRoleFunctions.RecordCount
        rsCopyRoleFunctions.AddNew
            For i = 0 To rsRoleFunctions.Fields.Count - 1
                rsCopyRoleFunctions.Fields(i).Value = rsRoleFunctions.Fields(i).Value
            Next
        rsCopyRoleFunctions.Update
        rsRoleFunctions.MoveNext
    Next
    
    rsRoleFunctions.Close
    Set rsRoleFunctions = Nothing
    rsCopyRoleFunctions.Close
    Set rsCopyRoleFunctions = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreSiteDatas", "modRestoreDBTransfer")

        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub


'----------------------------------------------------------------------------------
Private Sub RestoreRoles()
'----------------------------------------------------------------------------------
'REM 11/02/03
'Restore the Roles on the site
'----------------------------------------------------------------------------------
Dim rsRoles As ADODB.Recordset
Dim rsCopyRoles As ADODB.Recordset
Dim i As Long
Dim j As Long
Dim sSQL As String
Dim sSQL1 As String

    On Error GoTo ErrHandler
     
    CopyDataADODBConnection.Execute "DELETE FROM Role"

    'creates recordset to contain records to be copied
    sSQL = "SELECT * FROM Role "
    Set rsRoles = New ADODB.Recordset
    rsRoles.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'creates recordset to contain records to be copied
    sSQL1 = "SELECT * FROM Role " _
    & " WHERE 0 = 1"
    Set rsCopyRoles = New ADODB.Recordset
    rsCopyRoles.Open sSQL1, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'checks if records exist
    If rsRoles.RecordCount <= 0 Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    'move to first record inrecordset
    rsRoles.MoveFirst

    'begin record insertion
    For j = 1 To rsRoles.RecordCount
        rsCopyRoles.AddNew
            For i = 0 To rsRoles.Fields.Count - 1
                rsCopyRoles.Fields(i).Value = rsRoles.Fields(i).Value
            Next
        rsCopyRoles.Update
        rsRoles.MoveNext
    Next
    
    rsRoles.Close
    Set rsRoles = Nothing
    rsCopyRoles.Close
    Set rsCopyRoles = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreSiteDatas", "modRestoreDBTransfer")

        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select


End Sub

'----------------------------------------------------------------------------------
Private Sub RestoreMACROPassword()
'----------------------------------------------------------------------------------
'REM 11/02/03
'Restore the site Password policy
'----------------------------------------------------------------------------------
Dim rsPswdPolicy As ADODB.Recordset
Dim rsCopyPswdPolicy As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Long
    
    On Error GoTo ErrHandler
     
    CopyDataADODBConnection.Execute "DELETE FROM MACROPassword"
    
    'creates recordset to contain records to be restored
    sSQL = "SELECT * FROM MACROPassword "
    Set rsPswdPolicy = New ADODB.Recordset
    rsPswdPolicy.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'creates receiving recordset
    sSQL1 = "SELECT * FROM MACROPassword " _
    & " WHERE 0 = 1"
    Set rsCopyPswdPolicy = New ADODB.Recordset
    rsCopyPswdPolicy.Open sSQL1, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checks if records exist
    If rsPswdPolicy.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record in recordset
    rsPswdPolicy.MoveFirst
    rsCopyPswdPolicy.AddNew

        For i = 0 To rsPswdPolicy.Fields.Count - 1
            rsCopyPswdPolicy.Fields(i).Value = rsPswdPolicy.Fields(i).Value
        Next
    rsCopyPswdPolicy.Update

    rsPswdPolicy.Close
    Set rsPswdPolicy = Nothing
    rsCopyPswdPolicy.Close
    Set rsCopyPswdPolicy = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreMACRPPassword", "modRestoreDBTransfer")

        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub


'----------------------------------------------------------------------------------
Private Sub RestoreUserDatabases()
'----------------------------------------------------------------------------------
'REM 10/02/03
'Restore the User Database table
'----------------------------------------------------------------------------------
Dim sSQL As String
Dim sSQL1 As String
Dim rsUsers As ADODB.Recordset
Dim rsCopyUserDatabases As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrHandler

    CopyDataADODBConnection.Execute "DELETE FROM UserDatabase"

    'get recordset of Users for a specific site from the UserRole table
    sSQL = "SELECT DISTINCT UserName FROM UserRole " _
        & " WHERE (UserRole.SiteCode = '" & msSite & "' OR UserRole.SiteCode = 'AllSites')"
    Set rsUsers = New ADODB.Recordset
    rsUsers.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText

    If rsUsers.RecordCount > 0 Then

        'creates site recordset to contain records to be restored
        sSQL1 = "Select * from UserDatabase " _
        & " Where 0 = 1"
        Set rsCopyUserDatabases = New ADODB.Recordset
        rsCopyUserDatabases.Open sSQL1, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

        'checks if records exist
        If rsUsers.RecordCount <= 0 Then
            Screen.MousePointer = vbNormal
            Exit Sub
        End If

        'move to first record in recordset
        rsUsers.MoveFirst

        'begin record restoration
        For i = 1 To rsUsers.RecordCount
            rsCopyUserDatabases.AddNew
  
            rsCopyUserDatabases.Fields(0).Value = rsUsers.Fields(0).Value
            rsCopyUserDatabases.Fields(1).Value = msDatabaseName

            rsCopyUserDatabases.Update
            rsUsers.MoveNext
        Next

        rsUsers.Close
        Set rsUsers = Nothing
        rsCopyUserDatabases.Close
        Set rsCopyUserDatabases = Nothing
    End If

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreUserDatabases", "modRestoreDBTransfer")

        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'----------------------------------------------------------------------------------
Private Sub RestoreMACROUsers()
'----------------------------------------------------------------------------------
'REM 10/02/03
'Restore the site's users
'----------------------------------------------------------------------------------
Dim sSQL As String
Dim sSQL1 As String
Dim rsUserRole As ADODB.Recordset
Dim rsUsers As ADODB.Recordset
Dim rsCopyUsers As ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim sUserNames As String
Dim rsPswdHistory As ADODB.Recordset
Dim rsCopyPswdHistory As ADODB.Recordset
    
    On Error GoTo ErrHandler

    CopyDataADODBConnection.Execute "DELETE FROM MACROUser"

    'get recordset of Users for a specific site from the UserRole table
    sSQL = "SELECT DISTINCT UserName FROM UserRole " _
        & " WHERE (UserRole.SiteCode = '" & msSite & "' OR UserRole.SiteCode = 'AllSites')"
    Set rsUserRole = New ADODB.Recordset
    rsUserRole.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    If rsUserRole.RecordCount > 0 Then
    
        'get list of user names
        Do While Not rsUserRole.EOF
            sUserNames = sUserNames & rsUserRole!UserName & "','"
            rsUserRole.MoveNext
        Loop
        'remove last delimiter
        sUserNames = Left(sUserNames, Len(sUserNames) - 3)
        
        'get recordset of all user details from MACROUser table
        sSQL = "SELECT * FROM MACROUser WHERE UserName IN ( '" & sUserNames & "')"
        Set rsUsers = New ADODB.Recordset
        rsUsers.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
            
        'creates site recordset to contain records to be restored
        sSQL1 = "Select * from MACROUser " _
        & " Where 0 = 1"
        Set rsCopyUsers = New ADODB.Recordset
        rsCopyUsers.Open sSQL1, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            
        'checks if records exist
        If rsUsers.RecordCount <= 0 Then
            'added 3/08/2001
            Screen.MousePointer = vbNormal
            Exit Sub
        End If

        'move to first record in recordset
        rsUsers.MoveFirst
    
        'begin record restoration
        For j = 1 To rsUsers.RecordCount
            rsCopyUsers.AddNew
                For i = 0 To rsUsers.Fields.Count - 1
                    rsCopyUsers.Fields(i).Value = rsUsers.Fields(i).Value
                Next
            rsCopyUsers.Update
            rsUsers.MoveNext
        Next
        
        rsUsers.Close
        Set rsUsers = Nothing
        rsCopyUsers.Close
        Set rsCopyUsers = Nothing

    End If
    
'************Restore Password History*****************

    CopyDataADODBConnection.Execute "DELETE FROM PasswordHistory"

    sSQL = "SELECT * FROM PasswordHistory WHERE UserName IN ('" & sUserNames & "')"
    Set rsPswdHistory = New ADODB.Recordset
    rsPswdHistory.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'creates recordset to contain records to be copied
    sSQL1 = "SELECT * FROM PasswordHistory " _
    & " WHERE 0 = 1"
    Set rsCopyPswdHistory = New ADODB.Recordset
    rsCopyPswdHistory.Open sSQL1, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'checks if records exist
    If rsPswdHistory.RecordCount <= 0 Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    'move to first record inrecordset
    rsPswdHistory.MoveFirst

    'begin record insertion
    For j = 1 To rsPswdHistory.RecordCount
        rsCopyPswdHistory.AddNew
            For i = 0 To rsPswdHistory.Fields.Count - 1
                rsCopyPswdHistory.Fields(i).Value = rsPswdHistory.Fields(i).Value
            Next
        rsCopyPswdHistory.Update
        rsPswdHistory.MoveNext
    Next
    
    rsPswdHistory.Close
    Set rsPswdHistory = Nothing
    rsCopyPswdHistory.Close
    Set rsCopyPswdHistory = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreMACROUsers", "modRestoreDBTransfer")
                                    
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'----------------------------------------------------------------------------------
Private Sub RestoreTrialSites()
'----------------------------------------------------------------------------------
'added 08/08/2001 Ash restores existing TrialSites rows
'----------------------------------------------------------------------------------
Dim rsTrialSites As ADODB.Recordset
Dim rsCopyTrialSites As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim i As Long
Dim j As Long
    
    On Error GoTo ErrHandler
     
    CopyDataADODBConnection.Execute "DELETE FROM TrialSite" _
        & " Where ClinicalTrialID = " & mlTargetTrialId _
        & " AND TrialSite = '" & msSite & "'"
   
    'creates recordset to contain records to be restored
    sSQL = "Select * from TrialSite " _
    & " Where ClinicalTrialID = " & mlTrialID _
    & " AND TrialSite = '" & msSite & "'"
    
    'creates receiving recordset
    sSQL1 = "Select * from TrialSite " _
    & " Where 0 = 1"
    
    'setting and initialising recordset
    Set rsTrialSites = New ADODB.Recordset
    rsTrialSites.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'setting and initialising recordset
    Set rsCopyTrialSites = New ADODB.Recordset
    'ASH 16/04/2002
    'commented out to fix SR 4271 in Macro 2.2/3.0
    'rsCopyTrialSites.CursorLocation = adUseClient
    rsCopyTrialSites.Open sSQL1, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    'checks if records exist
    If rsTrialSites.RecordCount <= 0 Then
        Exit Sub
    End If
    
    'move to first record in recordset
    rsTrialSites.MoveFirst
     
    'begin record restoration
     For j = 1 To rsTrialSites.RecordCount
        rsCopyTrialSites.AddNew
        rsCopyTrialSites.Fields(0) = mlTargetTrialId
            For i = 1 To rsTrialSites.Fields.Count - 1
                rsCopyTrialSites.Fields(i).Value = rsTrialSites.Fields(i).Value
            Next
        rsCopyTrialSites.Update
        rsTrialSites.MoveNext
    Next j

    rsTrialSites.Close
    Set rsTrialSites = Nothing
    rsCopyTrialSites.Close
    Set rsCopyTrialSites = Nothing
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreTrialSites", "modRestoreDBTransfer")
                                   
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'-------------------------------------------------------------------------------------
Private Sub RestoreSiteDatas()
'-------------------------------------------------------------------------------------
'restores site data rows
'-------------------------------------------------------------------------------------
Dim rsSiteDatas As ADODB.Recordset
Dim rsCopySiteDatas As ADODB.Recordset
Dim i As Long
Dim j As Long
Dim sSQL As String
Dim sSQL1 As String

    On Error GoTo ErrHandler
     
    CopyDataADODBConnection.Execute "DELETE FROM Site where Site ='" & msSite & "'"

    'creates recordset to contain records to be copied
    sSQL = "Select * from Site " _
    & " Where Site = '" & msSite & "'"

    'creates recordset to contain records to be copied
    sSQL1 = "Select * from Site " _
    & " Where 0 = 1"

    'setting and initialising recordset
    Set rsSiteDatas = New ADODB.Recordset
   rsSiteDatas.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'setting and initialising recordset
    Set rsCopySiteDatas = New ADODB.Recordset
    'ASH 16/04/2002
    'commented out to fix SR 4271 in Macro 2.2/3.0
    'rsCopySiteDatas.CursorLocation = adUseClient
    rsCopySiteDatas.Open sSQL1, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'checks if records exist
    If rsSiteDatas.RecordCount <= 0 Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    'move to first record inrecordset
    rsSiteDatas.MoveFirst

    'begin record insertion
     For j = 1 To rsSiteDatas.RecordCount
        rsCopySiteDatas.AddNew
            For i = 0 To rsSiteDatas.Fields.Count - 1
                rsCopySiteDatas.Fields(i).Value = rsSiteDatas.Fields(i).Value
            Next
        rsCopySiteDatas.Update
        rsSiteDatas.MoveNext
    Next j
    rsSiteDatas.Close
    Set rsSiteDatas = Nothing
    rsCopySiteDatas.Close
    Set rsCopySiteDatas = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreSiteDatas", "modRestoreDBTransfer")

        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'-----------------------------------------------------------------
Private Sub RestoreUserRoleData()
'-----------------------------------------------------------------
'REM 10/02/03
'Restore the UserRole data
'-----------------------------------------------------------------
Dim sSQL As String
Dim sSQL1 As String
Dim rsUserRole As ADODB.Recordset
Dim rsCopyUserRole As ADODB.Recordset
Dim i As Single
Dim j As Integer

    On Error GoTo ErrHandler

    CopyDataADODBConnection.Execute "DELETE FROM UserRole"

    'get recordset of UserRole for specific site
    sSQL = "SELECT * FROM UserRole " _
        & " WHERE (UserRole.SiteCode = '" & msSite & "' OR UserRole.SiteCode = 'AllSites')"
    Set rsUserRole = New ADODB.Recordset
    rsUserRole.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText

    'creates recordset to contain records to be restored
    sSQL1 = "SELECT * FROM UserRole " _
    & " WHERE 0 = 1"
    Set rsCopyUserRole = New ADODB.Recordset
    rsCopyUserRole.Open sSQL1, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'checks if records exist
    If rsUserRole.RecordCount <= 0 Then
        'added 3/08/2001
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    'move to first record in recordset
    rsUserRole.MoveFirst

    'begin record restoration
    For j = 1 To rsUserRole.RecordCount
        rsCopyUserRole.AddNew
            For i = 0 To rsUserRole.Fields.Count - 1
                rsCopyUserRole.Fields(i).Value = rsUserRole.Fields(i).Value
            Next
        rsCopyUserRole.Update
        rsUserRole.MoveNext
    Next
    rsUserRole.Close
    Set rsUserRole = Nothing
    rsCopyUserRole.Close
    Set rsCopyUserRole = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreSiteUserData", "modRestoreDBTransfer")

        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'-----------------------------------------------------------------
Private Sub RestoreSiteUserData()
'-----------------------------------------------------------------
'restores SiteUser data as part of retore process
'-----------------------------------------------------------------
Dim rsSiteUser As ADODB.Recordset
Dim rsCopySiteUser As ADODB.Recordset
Dim i As Long
Dim j As Long
Dim sSQL As String
Dim sSQL1 As String

    On Error GoTo ErrHandler
     
    CopyDataADODBConnection.Execute "DELETE FROM SiteUser WHERE Site = '" & msSite & "'"

    'creates recordset to contain records to be restored
    sSQL = "Select * from SiteUser " _
    & " Where Site = '" & msSite & "'"

    'creates recordset to contain records to be restored
    sSQL1 = "Select * from SiteUser " _
    & " Where 0 = 1"

    'setting and initialising recordset
    Set rsSiteUser = New ADODB.Recordset
   rsSiteUser.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'setting and initialising recordset
    Set rsCopySiteUser = New ADODB.Recordset
    'ASH 16/04/2002
    'commented out to fix SR 4271 in Macro 2.2/3.0
    'rsCopySiteUser.CursorLocation = adUseClient
    rsCopySiteUser.Open sSQL1, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'checks if records exist
    If rsSiteUser.RecordCount <= 0 Then
        'added 3/08/2001
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    'move to first record in recordset
   rsSiteUser.MoveFirst

    'begin record restoration
     For j = 1 To rsSiteUser.RecordCount
        rsCopySiteUser.AddNew
            For i = 0 To rsSiteUser.Fields.Count - 1
                rsCopySiteUser.Fields(i).Value = rsSiteUser.Fields(i).Value
            Next
        rsCopySiteUser.Update
        rsSiteUser.MoveNext
    Next j
    rsSiteUser.Close
    Set rsSiteUser = Nothing
    rsCopySiteUser.Close
    Set rsCopySiteUser = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreSiteUserData", "modRestoreDBTransfer")

        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub
'----------------------------------------------------------------------------
Private Sub RestoreAllDataInTable(ByVal msTablename As String)
'----------------------------------------------------------------------------
'This routine restores all the data in the approprite tables
'without any SQL limiting factor i.e no limitation based on Clinicaltrialid
'or trialsite
'----------------------------------------------------------------------------
Dim sSQL As String
Dim sFromSQL As String
Dim sToSQL As String
Dim i As Long
Dim j As Long
Dim msSegmentId As String
Dim rsTotalTables As ADODB.Recordset
Dim rsFromTable As New ADODB.Recordset
Dim rsToTable As New ADODB.Recordset
     
     On Error GoTo ErrHandler
  
     CopyDataADODBConnection.Execute "DELETE FROM " & msTablename
     
     'creates recordset to contain records to be copied
     sFromSQL = "Select * from " & msTablename
    
     sToSQL = "Select * from " & msTablename _
     & " where 0 = 1"

     Set rsFromTable = New ADODB.Recordset
     rsFromTable.Open sFromSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

     Set rsToTable = New ADODB.Recordset
     'ASH 16/04/2002
    'commented out to fix SR 4271 in Macro 2.2/3.0
    'rsToTable.CursorLocation = adUseClient
     rsToTable.Open sToSQL, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

     If rsFromTable.RecordCount <= 0 Then
         Exit Sub
     End If

     'move to first record in recordset
     rsFromTable.MoveFirst
     'loop to copy records into new trialID
      For j = 1 To rsFromTable.RecordCount
         rsToTable.AddNew
             For i = 0 To rsFromTable.Fields.Count - 1
                 rsToTable.Fields(i).Value = rsFromTable.Fields(i).Value
              Next
     rsToTable.Update
     rsFromTable.MoveNext
     Next j
     
    rsFromTable.Close
    Set rsFromTable = Nothing
    rsToTable.Close
    Set rsToTable = Nothing
        
       
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreAllDataInTable", "modRestoreDBTransfer")
                                    
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub
'--------------------------------------------------------------------
Private Sub RestoreNormalRangeData()
'--------------------------------------------------------------------
'Restores NormalRange Data as part of retore process
'--------------------------------------------------------------------
Dim rsNormalRange As ADODB.Recordset
Dim rsCopyNormalRange As ADODB.Recordset
Dim i As Long
Dim j As Long
Dim sSQL As String
Dim sSQL1 As String

    On Error GoTo ErrHandler
     
    CopyDataADODBConnection.Execute "DELETE FROM NormalRange"

    'creates recordset to contain records to be restored
    sSQL = "Select * from NormalRange "
    
    'creates recordset to contain records to be restored
    sSQL1 = "Select * from NormalRange " _
    & " Where 0 = 1"

    'setting and initialising recordset
    Set rsNormalRange = New ADODB.Recordset
    rsNormalRange.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'setting and initialising recordset
    Set rsCopyNormalRange = New ADODB.Recordset
    'ASH 16/04/2002
    'commented out to fix SR 4271 in Macro 2.2/3.0
    'rsCopyNormalRange.CursorLocation = adUseClient
    rsCopyNormalRange.Open sSQL1, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'checks if records exist
    If rsNormalRange.RecordCount <= 0 Then
        'added 3/08/2001
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    'move to first record in recordset
   rsNormalRange.MoveFirst

    'begin record restoration
     For j = 1 To rsNormalRange.RecordCount
        rsCopyNormalRange.AddNew
            For i = 0 To rsNormalRange.Fields.Count - 1
                rsCopyNormalRange.Fields(i).Value = rsNormalRange.Fields(i).Value
            Next
        rsCopyNormalRange.Update
        rsNormalRange.MoveNext
    Next j
    
    rsNormalRange.Close
    Set rsNormalRange = Nothing
    rsCopyNormalRange.Close
    Set rsCopyNormalRange = Nothing

    

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreNormalRangeData", "modRestoreDBTransfer")
                                    
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub
'--------------------------------------------------------------------
Private Sub RestoreLaboratoryData()
'--------------------------------------------------------------------
'Restores LaboratoryData as part of retore process
'--------------------------------------------------------------------
Dim rsLaboratory As ADODB.Recordset
Dim rsCopyLaboratory As ADODB.Recordset
Dim i As Long
Dim j As Long
Dim sSQL As String
Dim sSQL1 As String

    On Error GoTo ErrHandler
     
    CopyDataADODBConnection.Execute "DELETE FROM Laboratory WHERE Site= '" & msSite & "'"

    'creates recordset to contain records to be restored
    sSQL = "Select * from Laboratory " _
    & " Where Site = '" & msSite & "'"

    'creates recordset to contain records to be restored
    sSQL1 = "Select * from Laboratory " _
    & " Where 0 = 1"

    'setting and initialising recordset
    Set rsLaboratory = New ADODB.Recordset
   rsLaboratory.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'setting and initialising recordset
    Set rsCopyLaboratory = New ADODB.Recordset
    'ASH 16/04/2002
    'commented out to fix SR 4271 in Macro 2.2/3.0
    'rsCopyLaboratory.CursorLocation = adUseClient
    rsCopyLaboratory.Open sSQL1, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'checks if records exist
    If rsLaboratory.RecordCount <= 0 Then
        'added 3/08/2001
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    'move to first record in recordset
    rsLaboratory.MoveFirst

    'begin record insertion
     For j = 1 To rsLaboratory.RecordCount
        rsCopyLaboratory.AddNew
        rsCopyLaboratory.Fields(0) = mlTargetTrialId
            For i = 1 To rsLaboratory.Fields.Count - 1
                rsCopyLaboratory.Fields(i).Value = rsLaboratory.Fields(i).Value
            Next
        rsCopyLaboratory.Update
        rsLaboratory.MoveNext
    Next j
    rsLaboratory.Close
    Set rsLaboratory = Nothing
    rsCopyLaboratory.Close
    Set rsCopyLaboratory = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreLaboratoryData", "modRestoreDBTransfer")
                                    
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub
'---------------------------------------------------------------------
Private Sub RestoreSiteLaboratoryData()
'---------------------------------------------------------------------
'Restores SiteLaboratory Data as part of retore process
'---------------------------------------------------------------------
Dim rsSiteLaboratory As ADODB.Recordset
Dim rsCopySiteLaboratory As ADODB.Recordset
Dim i As Long
Dim j As Long
Dim sSQL As String
Dim sSQL1 As String

    On Error GoTo ErrHandler
     
    CopyDataADODBConnection.Execute "DELETE FROM SiteLaboratory WHERE Site= '" & msSite & "'"

   'creates recordset to contain records to be restored
    sSQL = "Select * from SiteLaboratory " _
    & " Where Site = '" & msSite & "'"

    'creates recordset to contain records to be restored
    sSQL1 = "Select * from SiteLaboratory " _
    & " Where 0 = 1"

    'setting and initialising recordset
    Set rsSiteLaboratory = New ADODB.Recordset
    rsSiteLaboratory.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    'setting and initialising recordset
    Set rsCopySiteLaboratory = New ADODB.Recordset
    'ASH 16/04/2002
    'commented out to fix SR 4271 in Macro 2.2/3.0
    'rsCopySiteLaboratory.CursorLocation = adUseClient
    rsCopySiteLaboratory.Open sSQL1, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText

    ' checks if records exist
    If rsSiteLaboratory.RecordCount <= 0 Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    ' move to first record in recordset
   rsSiteLaboratory.MoveFirst

    ' begin record insertion
     For j = 1 To rsSiteLaboratory.RecordCount
        rsCopySiteLaboratory.AddNew
            For i = 0 To rsSiteLaboratory.Fields.Count - 1
                rsCopySiteLaboratory.Fields(i).Value = rsSiteLaboratory.Fields(i).Value
            Next
        rsCopySiteLaboratory.Update
        rsSiteLaboratory.MoveNext
    Next j
    rsSiteLaboratory.Close
    Set rsSiteLaboratory = Nothing
    rsCopySiteLaboratory.Close
    Set rsCopySiteLaboratory = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreSiteLaboratoryData", "modRestoreDBTransfer")
                                    
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub
'----------------------------------------------------------------------------------------
Private Function IsDBVersionValid(oLoginCon As ADODB.Connection, oCopyCon As ADODB.Connection) As Boolean
'----------------------------------------------------------------------------------------
'Added 06/02/2002 ASH
'Checks if database versions match before restore carried out
'----------------------------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim rsTemp1 As ADODB.Recordset
Dim sSQL As String
Dim sSQL1 As String
Dim sMacroVersion As String
Dim sMacroVersion1 As String
Dim sBuildVersion As String
Dim sBuildVersion1 As String
    
    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM MACROControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, oLoginCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    sMacroVersion = rsTemp![MACROVersion]
    sBuildVersion = rsTemp![BuildSubVersion]
    
    sSQL1 = "SELECT * FROM MACROControl"
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.Open sSQL1, oCopyCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    sMacroVersion1 = rsTemp1![MACROVersion]
    sBuildVersion1 = rsTemp1![BuildSubVersion]
    
    rsTemp.Close
    rsTemp1.Close
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    
    If sMacroVersion <> sMacroVersion1 And sBuildVersion1 <> sBuildVersion Then
        Call DialogError("Your login database is of version " & sMacroVersion1 & "." & sBuildVersion1 & vbCrLf _
            & "and your restore database is of version " & sMacroVersion & "." & sBuildVersion & "." & vbCrLf _
            & "You may need to upgrade your restore database.", "Restore Failed")
        IsDBVersionValid = False
    Else
        IsDBVersionValid = True
    End If
    
    

Exit Function
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "IsDBVersionValid", "modRestoreDBTransfer")
                                    
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'----------------------------------------------------------------------------------------
Private Sub RestoreMIMessages()
'---------------------------------------------------------------------
' DPH 06/12/2002 - SR 4980 Restores MIMessages Data as part
'    of restore process - Code modified from MACRO 2.1 version
'---------------------------------------------------------------------
Dim dblMIMessageReceived As Double
Dim rsMIMessage As ADODB.Recordset
Dim rsCopyMIMessage As ADODB.Recordset
Dim sSQLFrom As String
Dim sSQLTo As String

    On Error GoTo ErrHandler
    
    Set rsMIMessage = New ADODB.Recordset
    Set rsCopyMIMessage = New ADODB.Recordset

    ' clear destination MIMessage table
    CopyDataADODBConnection.Execute "DELETE FROM MIMessage where MIMessageTrialName  = '" & msTrialName & "' AND MIMessageSite = '" & msSite & "'"
    
    sSQLFrom = "SELECT * FROM MIMessage where MIMessageTrialName  = '" & msTrialName & "' AND MIMessageSite = '" & msSite & "'"
    sSQLTo = "SELECT * FROM MIMessage"
    'Get the appopriate records and open the recordset we are sending the data to
    rsMIMessage.Open sSQLFrom, MacroADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    
    rsCopyMIMessage.Open sSQLTo, CopyDataADODBConnection, adOpenKeyset, adLockOptimistic

    With rsMIMessage
        If rsMIMessage.RecordCount > 0 Then
            rsMIMessage.MoveLast
            mlTotalRecordCount = rsMIMessage.RecordCount
            rsMIMessage.MoveFirst
            ' The mlTotalRecordCount -1 is because you have the count but have already moved to the first record
            For mlRecordCounter = l To mlTotalRecordCount - 1
            
                vFields = Array("MIMESSAGEID", "MIMESSAGESITE", "MIMESSAGESOURCE", "MIMESSAGETYPE", "MIMESSAGESCOPE", "MIMESSAGEOBJECTID", _
                                "MIMESSAGEOBJECTSOURCE", "MIMESSAGEPRIORITY", "MIMESSAGETRIALNAME", "MIMESSAGEPERSONID", "MIMESSAGEVISITID", _
                                "MIMESSAGEVISITCYCLE", "MIMESSAGECRFPAGETASKID", "MIMESSAGERESPONSETASKID", "MIMESSAGERESPONSEVALUE", _
                                "MIMESSAGEOCDISCREPANCYID", "MIMESSAGECREATED", "MIMESSAGESENT", "MIMESSAGERECEIVED", "MIMESSAGEHISTORY", _
                                "MIMESSAGEPROCESSED", "MIMESSAGESTATUS", "MIMESSAGETEXT", "MIMESSAGEUSERNAME", "MIMESSAGEUSERNAMEFULL", _
                                "MIMESSAGERESPONSETIMESTAMP", "MIMESSAGERESPONSECYCLE", "MIMESSAGECREATED_TZ", "MIMESSAGESENT_TZ", _
                                "MIMESSAGERECEIVED_TZ", "SEQUENCEID", "MIMESSAGECRFPAGEID", "MIMESSAGECRFPAGECYCLE", "MIMESSAGEDATAITEMID")
            
                               'if the message originated on the site then we dont have the MIMessageReceived value so make it 0
                               If rsMIMessage!MIMessageSource = TypeOfInstallation.RemoteSite Then
                                   dblMIMessageReceived = 0
                               'if the message originated on the server then set the MIMessageReceived value to when it was sent from the server its the best we've got
                               ElseIf rsMIMessage!MIMessageSource = TypeOfInstallation.Server Then
                                    dblMIMessageReceived = rsMIMessage!MIMessageSent
                               End If
                               
                ' Note that the MIMessageProcessed field is hard coded to 1 to signify that this record had been responded to by the user
                vValues = Array(rsMIMessage!MIMESSAGEID, rsMIMessage!MIMessageSite, rsMIMessage!MIMessageSource, _
                                rsMIMessage!MIMessageType, rsMIMessage!MIMESSAGESCOPE, rsMIMessage!MIMESSAGEOBJECTID, _
                                rsMIMessage!MIMESSAGEOBJECTSOURCE, rsMIMessage!MIMESSAGEPRIORITY, _
                                rsMIMessage!MIMessageTrialName, rsMIMessage!MIMessagePersonId, rsMIMessage!MIMessageVisitId, _
                                rsMIMessage!MIMESSAGEVISITCYCLE, rsMIMessage!MIMessageCRFPageTaskId, rsMIMessage!MIMessageResponseTaskId, _
                                rsMIMessage!MIMESSAGERESPONSEVALUE, rsMIMessage!MIMESSAGEOCDISCREPANCYID, rsMIMessage!MIMessageCreated, _
                                rsMIMessage!MIMessageSent, rsMIMessage!MIMessageReceived, rsMIMessage!MIMESSAGEHISTORY, 1, _
                                rsMIMessage!MIMESSAGESTATUS, rsMIMessage!MIMessageText, rsMIMessage!MIMessageUserName, _
                                rsMIMessage!MIMESSAGEUSERNAMEFULL, rsMIMessage!MIMESSAGERESPONSETIMESTAMP, rsMIMessage!MIMESSAGERESPONSECYCLE, _
                                rsMIMessage!MIMESSAGECREATED_TZ, rsMIMessage!MIMESSAGESENT_TZ, rsMIMessage!MIMESSAGERECEIVED_TZ, _
                                rsMIMessage!SEQUENCEID, rsMIMessage!MIMESSAGECRFPAGEID, rsMIMessage!MIMESSAGECRFPAGECYCLE, rsMIMessage!MIMESSAGEDATAITEMID)
                                   
                'Add the new row
                 rsCopyMIMessage.AddNew vFields, vValues
                .MoveNext
                'Increment the counter
            Next mlRecordCounter
        
        End If
    End With
        
    Set rsMIMessage = Nothing
    Set rsCopyMIMessage = Nothing
       
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RestoreMIMessages", "modRestoreDBTransfer")
                                    
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub



