VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpMACRO20To22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2003. All Rights Reserved
'   File:       UpMACRO20To22.cls
'   Author:     Richard Meinesz 18/06/2003
'   Purpose:    Class to upgrade a MACRO database to the latest version of MACRO 2.2 from any previous version
'----------------------------------------------------------------------------------------'
'REVISIONS:
'
'----------------------------------------------------------------------------------------'

Option Explicit

Private Const msMACRO_DATA_DB_DEFINITION As String = "MacroDataDb.txt"
Private Const msMACRO_SECURITY_DB_DEFINITION As String = "MacroSecDb.txt"
Private Const msMACRO_DATA_TYPES_DEFINITION As String = "MacroDataTypes.txt"
Private Const msMACRO_VTRACK_DB_DEFINITION As String = "VtrackTables.txt"

Private moMACROcon As ADODB.Connection

Private mColDataTypeLookUp As Collection

Private meDatabaseType As MACRODatabaseType
Private msDatabaseLocation As String
Private msDatabasePswd As String

'---------------------------------------------------------------------
Public Function Init(ByRef sErrMsg As String, sConnection As String, eDatabaseType As MACRODatabaseType, Optional sDatabaseLocation As String = "", Optional sDatabasePswd As String = "") As Boolean
'---------------------------------------------------------------------
'REM 18/06/03
'Initialse the class, create the database connection
'---------------------------------------------------------------------
    On Error GoTo Errlabel
    
    Set moMACROcon = New ADODB.Connection
    moMACROcon.Open sConnection
    moMACROcon.CursorLocation = adUseClient
    
    meDatabaseType = eDatabaseType
    msDatabaseLocation = sDatabaseLocation
    msDatabasePswd = sDatabasePswd
    
    sErrMsg = ""
    Init = True
    
Exit Function
Errlabel:
    sErrMsg = "Error Description: " & Err.Description & " Error Number: " & Err.Number
    Init = False
End Function

'---------------------------------------------------------------------
Public Sub UpgradeDataDatabase20To22(nDatabaseType As Integer)
'---------------------------------------------------------------------
' REM 18/06/03
' Upgrade a database to latest version of MACRO 2.2
'---------------------------------------------------------------------
Dim oMACRODatabase As Database

    ' Upgrade 2.0 Data database
    Call UpgradeDataDatabase2_0
    ' Upgrade 2.1 Data database
    Call UpgradeDataDatabase2_1
    ' Upgrade 2.2 Data database
    Call UpgradeDataDatabase2_2
    
    If nDatabaseType = MACRODatabaseType.Access Then
        'REM 05/08/03 - Refresh the Access database tables to ensure there are no timimng issues
        Set oMACRODatabase = DBEngine.Workspaces(0).OpenDatabase(msDatabaseLocation, False, False, "MS Access;PWD=" & msDatabasePswd)
        
        oMACRODatabase.TableDefs.Refresh
        
        Set oMACRODatabase = Nothing
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub UpgradeDataDatabase2_0()
'---------------------------------------------------------------------
' Upgrade a 2.0 database to the last 2.0 database
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sMacroVersion As String
Dim sBuildSubVersion As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM MACROControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, moMACROcon, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        MsgBox ("Your Macro database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    sMacroVersion = rsTemp![MACROVersion]
    sBuildSubVersion = rsTemp![BuildSubVersion]
    rsTemp.Close
    Set rsTemp = Nothing
    
    If sMacroVersion <> "2.0" Then
        ' Not a 2.0 database
        Exit Sub
    End If
    
    'Check for need to upgrade from BuildSubVersion [15] to [16]
    If sBuildSubVersion = "15" Then
        sBuildSubVersion = "16"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [16] to [17]
    If sBuildSubVersion = "16" Then
        sBuildSubVersion = "17"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [17] to [18]
    If sBuildSubVersion = "17" Then
        UpGradeData20_17to18       ' This involves database changes
        sBuildSubVersion = "18"
    End If
    
    'Check for need to upgrade from BuildSubVersion [18] to [19]
    If sBuildSubVersion = "18" Then
        sBuildSubVersion = "19"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [19] to [20]
    If sBuildSubVersion = "19" Then
        sBuildSubVersion = "20"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'Check for need to upgrade from BuildSubVersion [20] to [21]
    If sBuildSubVersion = "20" Then
        sBuildSubVersion = "21"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'Check for need to upgrade from BuildSubVersion [21] to [22]
    ' NCJ 8/3/00
    If sBuildSubVersion = "21" Then
        UpGradeData20_21to22       ' This involves database changes
        sBuildSubVersion = "22"
    End If

    'Check for need to upgrade from BuildSubVersion [22] to [23]
    ' WillC 14/3/00
    If sBuildSubVersion = "22" Then
        UpGradeData20_22to23       ' This involves database changes
        sBuildSubVersion = "23"
    End If

    'Check for need to upgrade from BuildSubVersion [23] to [27]
    ' NCJ 3/4/00
    If sBuildSubVersion = "23" Then
        sBuildSubVersion = "27"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [27] to [29]
    ' WillC 26/4/00
    If sBuildSubVersion = "27" Then
        UpGradeData20_27to29       ' This involves database changes
        sBuildSubVersion = "29"
    End If

    ' NCJ
    If sBuildSubVersion = "29" Then
        sBuildSubVersion = "31"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    ' NCJ
    If sBuildSubVersion = "31" Then
        sBuildSubVersion = "32"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'WillC 16/5/00
    'Check for need to upgrade from BuildSubVersion [32] to [33]
    If sBuildSubVersion = "32" Then
        UpGradeData20_32to33       ' This involves database changes
        sBuildSubVersion = "33"
    End If
    
    'WillC 26/5/00
    'Check for need to upgrade from BuildSubVersion [33] to [34]
    If sBuildSubVersion = "33" Then
        sBuildSubVersion = "34"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'WillC 30/5/00
    'Check for need to upgrade from BuildSubVersion [34] to [35]
    If sBuildSubVersion = "34" Then
        sBuildSubVersion = "35"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'NCJ 2/6/00
    If sBuildSubVersion = "35" Then
        sBuildSubVersion = "36"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
        'WillC 9/6/00
    If sBuildSubVersion = "36" Then
        sBuildSubVersion = "37"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

        'Nicky 15/6/00
    If sBuildSubVersion = "37" Then
        sBuildSubVersion = "39"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

        'Nicky 16/6/00
    If sBuildSubVersion = "39" Then
        sBuildSubVersion = "40"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpgradeDataDatabase2_0", "modUpgradeDatabases.bas")
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
Private Sub UpGradeDataToSubVersion(sSubVersion As String)
'---------------------------------------------------------------------
' Nicky 1/9/00
' Generic routine to update the build subversion in MACROControl in MACRO.mdb
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
  
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '" & sSubVersion & "'"
    moMACROcon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeDataToSubVersion" & sSubVersion, "modUpgradeDatabases")
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
Private Sub UpGradeData20_17to18()
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'Correct an entry in table StandardDataFormat
    sSQL = "UPDATE StandardDataFormat Set DataTypeId = 0 " _
       & " WHERE StandardDataFormatId = 2"
    moMACROcon.Execute sSQL
    
    'Change an entry in table StandardDataFormat
    sSQL = "UPDATE StandardDataFormat Set DataFormat = 'mm/dd/yyyy'" _
       & " WHERE StandardDataFormatId = 11"
    moMACROcon.Execute sSQL
    
    'Add ImportTimeStamp to the 5 Patient data files
    If meDatabaseType = MACRODatabaseType.Access Then
        'add to TrialSubject
        sSQL = "ALTER Table TrialSubject ADD COLUMN ImportTimeStamp DOUBLE"
        moMACROcon.Execute sSQL
        'add to VisitInstance
        sSQL = "ALTER Table VisitInstance ADD COLUMN ImportTimeStamp DOUBLE"
        moMACROcon.Execute sSQL
        'add to CRFPageInstance
        sSQL = "ALTER Table CRFPageInstance ADD COLUMN ImportTimeStamp DOUBLE"
        moMACROcon.Execute sSQL
        'add to DataItemResponse
        sSQL = "ALTER Table DataItemResponse ADD COLUMN ImportTimeStamp DOUBLE"
        moMACROcon.Execute sSQL
        'add to DataItemResponseHistory
        sSQL = "ALTER Table DataItemResponseHistory ADD COLUMN ImportTimeStamp DOUBLE"
        moMACROcon.Execute sSQL
    Else    'SQLServer or Oracle
        'add to TrialSubject
        sSQL = "ALTER Table TrialSubject ADD ImportTimeStamp DECIMAL(16,10) NULL"
        moMACROcon.Execute sSQL
        'add to VisitInstance
        sSQL = "ALTER Table VisitInstance ADD ImportTimeStamp DECIMAL(16,10) NULL"
        moMACROcon.Execute sSQL
        'add to CRFPageInstance
        sSQL = "ALTER Table CRFPageInstance ADD ImportTimeStamp DECIMAL(16,10) NULL"
        moMACROcon.Execute sSQL
        'add to DataItemResponse
        sSQL = "ALTER Table DataItemResponse ADD ImportTimeStamp DECIMAL(16,10) NULL"
        moMACROcon.Execute sSQL
        'add to DataItemResponseHistory
        sSQL = "ALTER Table DataItemResponseHistory ADD ImportTimeStamp DECIMAL(16,10) NULL"
        moMACROcon.Execute sSQL
    End If
  
    'Upgrade BuildSubVersion from [17] to [18]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '18'"
    moMACROcon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData20_17to18", "modUpgradeDatabases.bas")
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
Private Sub UpGradeData20_21to22()
'---------------------------------------------------------------------

Dim sSQL As String

    On Error GoTo ErrHandler
    
    'drop the no duplcates index on TrialName in table TrialType
    If meDatabaseType = MACRODatabaseType.Access Then
        sSQL = "DROP INDEX TrialTypeName on TrialType"
    Else
        sSQL = "ALTER TABLE TrialType DROP CONSTRAINT TrialTypeName"
    End If
    moMACROcon.Execute sSQL
    
    'drop the no duplicates index on PhaseName in table TrialPhase
    If meDatabaseType = MACRODatabaseType.Access Then
        sSQL = "DROP INDEX idx_PhaseName on TrialPhase"
    Else
        sSQL = "ALTER TABLE TrialPhase DROP CONSTRAINT PhaseName"
    End If
    moMACROcon.Execute sSQL
  
    'Upgrade BuildSubVersion from [21] to [22]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '22'"
    moMACROcon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData20_21to22", "modUpgradeDatabases.bas")
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
Private Sub UpGradeData20_22to23()
'---------------------------------------------------------------------
'WillC 14/3/00
'---------------------------------------------------------------------

Dim sSQL As String

    On Error GoTo ErrHandler
    
    ' Add the column to the StudyVisit table
    If meDatabaseType = MACRODatabaseType.Access Then
        sSQL = "ALTER TABLE  StudyVisit  ADD Column VisitDatePrompt BYTE "
    ElseIf meDatabaseType = MACRODatabaseType.sqlserver Then
        sSQL = sSQL & " ALTER TABLE  StudyVisit  ADD  VisitDatePrompt TINYINT  NULL DEFAULT 0"  'Alter Table can only allow nullable columns to be added
    ElseIf meDatabaseType = MACRODatabaseType.Oracle80 Then
        ' NOTE that you dont need to put COLUMN in the PLSQL statement
        sSQL = sSQL & " ALTER TABLE  StudyVisit  ADD (VisitDatePrompt NUMBER(1) DEFAULT 0)"
    End If

    moMACROcon.Execute sSQL
    
  
    'Upgrade BuildSubVersion from [22] to [23]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '23'"
    moMACROcon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData20_22to23", "modUpgradeDatabases.bas")
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
Private Sub UpGradeData20_27to29()
'---------------------------------------------------------------------
'WillC 26/4/00
'---------------------------------------------------------------------
Dim sSQL As String
Dim oTableDef As TableDef
Dim oMACRODatabase As Database

  If meDatabaseType = MACRODatabaseType.Access Then      ',NonExcusive,ReadWrite,Password
     Set oMACRODatabase = DBEngine.Workspaces(0).OpenDatabase(msDatabaseLocation, False, False, "MS Access;PWD=" & msDatabasePswd)
  End If

    ' Add the columns to the DataItemResponse table
    If meDatabaseType = MACRODatabaseType.Access Then
        sSQL = "ALTER TABLE  DataItemResponse  ADD Column ValidationId SMALLINT  "   ' Cant Add a default with the Jet ALTER TABLE Statement
    ElseIf meDatabaseType = MACRODatabaseType.sqlserver Then
        sSQL = "ALTER TABLE  DataItemResponse  ADD  ValidationId INTEGER  NULL DEFAULT 0"   'Alter Table can only allow nullable columns to be added
    ElseIf meDatabaseType = MACRODatabaseType.Oracle80 Then
        ' NOTE that you dont need to put COLUMN in the PLSQL statement
        sSQL = "ALTER TABLE  DataItemResponse  ADD (ValidationId NUMBER(8) DEFAULT 0)"
    End If

    moMACROcon.Execute sSQL

    If meDatabaseType = MACRODatabaseType.Access Then
        sSQL = "ALTER TABLE  DataItemResponse  ADD Column ValidationMessage MEMO "
    ElseIf meDatabaseType = MACRODatabaseType.sqlserver Then
        sSQL = "ALTER TABLE  DataItemResponse  ADD  ValidationMessage TEXT "  'Alter Table can only allow nullable columns to be added
    ElseIf meDatabaseType = MACRODatabaseType.Oracle80 Then
        ' NOTE that you dont need to put COLUMN in the PLSQL statement
        'WillC 29/4/2000 changed ValidationMessage to VARCHAR(2000)
        sSQL = "ALTER TABLE  DataItemResponse  ADD (ValidationMessage VARCHAR(2000))"
    End If

    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.Access Then
        sSQL = "ALTER TABLE  DataItemResponse  ADD Column OverruleReason TEXT(255) "
    ElseIf meDatabaseType = MACRODatabaseType.sqlserver Then
        sSQL = "ALTER TABLE  DataItemResponse  ADD  OverruleReason VARCHAR(255) "  'Alter Table can only allow nullable columns to be added
    ElseIf meDatabaseType = MACRODatabaseType.Oracle80 Then
        ' NOTE that you dont need to put COLUMN in the PLSQL statement
        sSQL = "ALTER TABLE  DataItemResponse  ADD (OverruleReason VARCHAR(255))"
    End If
    
    moMACROcon.Execute sSQL
        
    
    If meDatabaseType = MACRODatabaseType.Access Then
    
        ' TA 28/04/2000: code to get round refresh problem
        Set oMACRODatabase = Nothing
        Set oMACRODatabase = DBEngine.Workspaces(0).OpenDatabase(msDatabaseLocation, False, False, "MS Access;PWD=" & msDatabasePswd)
    
        oMACRODatabase.TableDefs.Refresh
        Set oTableDef = oMACRODatabase.TableDefs("DataItemResponse")
        With oTableDef
              .Fields("ValidationId").DefaultValue = 0
              .Fields("ValidationId").Required = False
              .Fields("ValidationMessage").AllowZeroLength = True
              .Fields("ValidationMessage").Required = False
              .Fields("OverruleReason").AllowZeroLength = True
              .Fields("OverruleReason").Required = False
        End With
    End If
    
    
'------ ' Add the columns to the DataItemResponseHistory table----------------------------
    
    
    If meDatabaseType = MACRODatabaseType.Access Then
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD Column ValidationId SMALLINT  "
    ElseIf meDatabaseType = MACRODatabaseType.sqlserver Then
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD  ValidationId INTEGER  NULL DEFAULT 0"  'Alter Table can only allow nullable columns to be added
    ElseIf meDatabaseType = MACRODatabaseType.Oracle80 Then
        ' NOTE that you dont need to put COLUMN in the PLSQL statement
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD (ValidationId NUMBER(8) DEFAULT 0)"
    End If

    moMACROcon.Execute sSQL

    If meDatabaseType = MACRODatabaseType.Access Then
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD Column ValidationMessage MEMO "
    ElseIf meDatabaseType = MACRODatabaseType.sqlserver Then
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD  ValidationMessage TEXT "  'Alter Table can only allow nullable columns to be added
    ElseIf meDatabaseType = MACRODatabaseType.Oracle80 Then
        ' NOTE that you dont need to put COLUMN in the PLSQL statement
        'WillC 29/4/2000 changed ValidationMessage to VARCHAR(2000)
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD (ValidationMessage VARCHAR(2000))"
    End If

    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.Access Then
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD Column OverruleReason TEXT(255) "
    ElseIf meDatabaseType = MACRODatabaseType.sqlserver Then
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD  OverruleReason VARCHAR(255) " 'Alter Table can only allow nullable columns to be added
    ElseIf meDatabaseType = MACRODatabaseType.Oracle80 Then
        ' NOTE that you dont need to put COLUMN in the PLSQL statement
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD (OverruleReason VARCHAR(255))"
    End If
    
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.Access Then
    
        ' TA 28/04/2000: code to get round refresh problem
        Set oMACRODatabase = Nothing
        Set oMACRODatabase = DBEngine.Workspaces(0).OpenDatabase(msDatabaseLocation, False, False, "MS Access;PWD=" & msDatabasePswd)
    
        oMACRODatabase.TableDefs.Refresh
        Set oTableDef = oMACRODatabase.TableDefs("DataItemResponseHistory")
        With oTableDef
              .Fields("ValidationId").DefaultValue = 0
              .Fields("ValidationId").Required = False
              .Fields("ValidationMessage").AllowZeroLength = True
              .Fields("ValidationMessage").Required = False
              .Fields("OverruleReason").AllowZeroLength = True
              .Fields("OverruleReason").Required = False
        End With
    End If
    Set oMACRODatabase = Nothing
    

    'Upgrade BuildSubVersion from [27] to [29]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '29'"
    moMACROcon.Execute sSQL
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "UpGradeData20_27to29", "modUpgradeDatabases")
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
Private Sub UpGradeData20_32to33()
'---------------------------------------------------------------------
' WillC 16/5/00
' If Any changes to the MACRO database make sure they are reflected
' in the RestoreSite module to keep it up to date
'---------------------------------------------------------------------
Dim sSQL As String
Dim oTableDef As TableDef
Dim oMACRODatabase As Database

    On Error GoTo ErrHandler
  
    'REM 07/06/02 - Leaving this as DAO as we never insert values into this table, only users do
    'Create the new Table MIMessage.
    If meDatabaseType = MACRODatabaseType.Access Then
        Set oMACRODatabase = Nothing
        Set oMACRODatabase = DBEngine.Workspaces(0).OpenDatabase(msDatabaseLocation, False, False, "MS Access;PWD=" & msDatabasePswd)
        sSQL = "CREATE TABLE MIMessage(MIMessageID INTEGER ,"
        sSQL = sSQL & " MIMessageSite TEXT(8),MIMessageSource SMALLINT,"
        sSQL = sSQL & " MIMessageType SMALLINT, MIMessageScope SMALLINT,"
        sSQL = sSQL & " MIMessageObjectID INTEGER, MIMessageObjectSource SMALLINT, MIMessagePriority SMALLINT,"
        sSQL = sSQL & " MIMessageTrialName TEXT(15), MIMessagePersonId INTEGER,"
        sSQL = sSQL & " MIMessageVisitId INTEGER, MIMessageVisitCycle SMALLINT,"
        sSQL = sSQL & " MIMessageCRFPageTaskID INTEGER, MIMessageResponseTaskId INTEGER,"
        sSQL = sSQL & " MIMessageResponseValue TEXT(255),MIMessageOCDiscrepancyID INTEGER, MIMessageCreated DOUBLE,"
        sSQL = sSQL & " MIMessageSent DOUBLE, MIMessageReceived DOUBLE,"
        sSQL = sSQL & " MIMessageHistory SMALLINT, MIMessageProcessed SMALLINT,"
        sSQL = sSQL & " MIMessageStatus SMALLINT, MIMessageText TEXT(255),"
        sSQL = sSQL & " MIMessageUserCode TEXT(50), MIMessageUserName TEXT(255),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (MIMessageID,MIMessageSite,MIMessageSource))"
        oMACRODatabase.Execute sSQL, dbFailOnError
        
        oMACRODatabase.TableDefs.Refresh
        Set oTableDef = oMACRODatabase.TableDefs("MIMessage")
        With oTableDef
            .Fields("MIMessageID").DefaultValue = 0
            .Fields("MIMessageID").Required = True
            .Fields("MIMessagesite").AllowZeroLength = False
            .Fields("MIMessagesite").Required = True
            .Fields("MIMessageSource").DefaultValue = 0
            .Fields("MIMessageSource").Required = True
            .Fields("MIMessageType").DefaultValue = 0
            .Fields("MIMessageType").Required = True
            .Fields("MIMessageScope").DefaultValue = 4
            .Fields("MIMessageScope").Required = True
            .Fields("MIMessageObjectID").DefaultValue = 0
            .Fields("MIMessageObjectID").Required = True
            .Fields("MIMessagePriority").DefaultValue = 5
            .Fields("MIMessagePriority").Required = False
            .Fields("MIMessageTrialName").AllowZeroLength = True
            .Fields("MIMessageTrialName").Required = True
            .Fields("MIMessagePersonId").DefaultValue = 0
            .Fields("MIMessagePersonId").Required = False
            .Fields("MIMessageVisitId").DefaultValue = 0
            .Fields("MIMessageVisitId").Required = False
            .Fields("MIMessageVisitCycle").DefaultValue = 0
            .Fields("MIMessageVisitCycle").Required = False
            .Fields("MIMessageCRFPageTaskID").DefaultValue = 0
            .Fields("MIMessageCRFPageTaskID").Required = False
            .Fields("MIMessageResponseTaskId").DefaultValue = 0
            .Fields("MIMessageResponseTaskId").Required = False
            .Fields("MIMessageResponseValue").AllowZeroLength = True
            .Fields("MIMessageResponseValue").Required = False
            .Fields("MIMessageOCDiscrepancyID").Required = False
            .Fields("MIMessageOCDiscrepancyID").DefaultValue = 0
            .Fields("MIMessageCreated").Required = True
            .Fields("MIMessageCreated").DefaultValue = 0
            .Fields("MIMessageSent").Required = True
            .Fields("MIMessageSent").DefaultValue = 0
            .Fields("MIMessageReceived").Required = True
            .Fields("MIMessageReceived").DefaultValue = 0
            .Fields("MIMessageHistory").Required = True
            .Fields("MIMessageHistory").DefaultValue = 0
            .Fields("MIMessageProcessed").Required = True
            .Fields("MIMessageProcessed").DefaultValue = 0
            .Fields("MIMessageStatus").Required = False
            .Fields("MIMessageStatus").DefaultValue = 0
            .Fields("MIMessageText").AllowZeroLength = True
            .Fields("MIMessageText").Required = False
            .Fields("MIMessageUserCode").AllowZeroLength = False
            .Fields("MIMessageUserCode").Required = True
            .Fields("MIMessageUserName").AllowZeroLength = True
            .Fields("MIMessageUserName").Required = False
        End With
        Set oTableDef = Nothing
        
        sSQL = "CREATE INDEX idx_MIMessageID "
        sSQL = sSQL & " ON MIMessage ( MIMessageID)"
        oMACRODatabase.Execute sSQL, dbFailOnError
        
        sSQL = " CREATE INDEX idx_MIMessageUserCode "
        sSQL = sSQL & " ON MIMessage ( MIMessageUserCode )"
        oMACRODatabase.Execute sSQL, dbFailOnError
          
        sSQL = " CREATE INDEX idx_MIMessageObjectID "
        sSQL = sSQL & " ON MIMessage ( MIMessageObjectID )"
        oMACRODatabase.Execute sSQL, dbFailOnError
        
        Set oMACRODatabase = Nothing
    ElseIf meDatabaseType = MACRODatabaseType.sqlserver Or MACRODatabaseType.Oracle80 Then
        sSQL = "CREATE TABLE MIMessage(MIMessageID INTEGER DEFAULT 0 ,"
        sSQL = sSQL & " MIMessageSite VARCHAR(8) ,MIMessageSource SMALLINT DEFAULT 0 ,"
        sSQL = sSQL & " MIMessageType SMALLINT DEFAULT 0 NOT NULL, MIMessageScope SMALLINT DEFAULT 4 NOT NULL,"
        sSQL = sSQL & " MIMessageObjectID INTEGER DEFAULT 0 NOT NULL, MIMessageObjectSource SMALLINT DEFAULT 0, MIMessagePriority SMALLINT DEFAULT 5 ,"
        sSQL = sSQL & " MIMessageTrialName VARCHAR(15) NULL, MIMessagePersonId INTEGER DEFAULT 0 NULL,"
        sSQL = sSQL & " MIMessageVisitId INTEGER DEFAULT 0 NULL, MIMessageVisitCycle SMALLINT DEFAULT 0 NULL,"
        sSQL = sSQL & " MIMessageCRFPageTaskID INTEGER DEFAULT 0 NULL, MIMessageResponseTaskId INTEGER DEFAULT 0 NULL,"
        sSQL = sSQL & " MIMessageResponseValue VARCHAR(255) NULL,MIMessageOCDiscrepancyID INTEGER DEFAULT 0, MIMessageCreated DECIMAL(16,10) DEFAULT 0  NOT NULL,"
        sSQL = sSQL & " MIMessageSent DECIMAL(16,10) DEFAULT 0 NOT NULL, MIMessageReceived DECIMAL(16,10) DEFAULT 0 NOT NULL,"
        sSQL = sSQL & " MIMessageHistory SMALLINT DEFAULT 0 NOT NULL, MIMessageProcessed SMALLINT DEFAULT 0 NOT NULL,"
        sSQL = sSQL & " MIMessageStatus SMALLINT DEFAULT 0 NULL, MIMessageText VARCHAR(255) NULL,"
        sSQL = sSQL & " MIMessageUserCode VARCHAR(50) NOT NULL, MIMessageUserName VARCHAR(255) NULL,"
        sSQL = sSQL & " CONSTRAINT PKMIMessage PRIMARY KEY "
        sSQL = sSQL & " (MIMessageID,MIMessageSite,MIMessageSource))"
        moMACROcon.Execute sSQL
                
        sSQL = "CREATE INDEX idx_MIMessageID "
        sSQL = sSQL & " ON MIMessage (MIMessageID)"
        moMACROcon.Execute sSQL
        
        sSQL = " CREATE INDEX idx_MIMessageUserCode "
        sSQL = sSQL & " ON MIMessage ( MIMessageUserCode )"
        moMACROcon.Execute sSQL
          
        sSQL = " CREATE INDEX idx_MIMessageObjectID "
        sSQL = sSQL & " ON MIMessage ( MIMessageObjectID )"
        moMACROcon.Execute sSQL
    End If

    'Upgrade BuildSubVersion from [32] to [33]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '33'"
    moMACROcon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeData20_32to33", "modUpgradeDatabases")
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
Private Sub UpgradeDataDatabase2_1()
'---------------------------------------------------------------------
' Upgrade a 2.1 MACRO.mdb, checking first for 2.0 -> 2.1
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sMacroVersion As String
Dim sBuildSubVersion As String
Dim sMSG As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM MACROControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, moMACROcon, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        MsgBox ("Your Macro database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    sBuildSubVersion = rsTemp![BuildSubVersion]
    sMacroVersion = rsTemp![MACROVersion]
    rsTemp.Close
    Set rsTemp = Nothing

    'don't do anything if futre version
    If Val(sMacroVersion) >= 2.1 Then
        Exit Sub
    End If
    ' Check for version 2.0
    If sMacroVersion = "2.0" Then
        'sMsg = "You are about to upgrade your MACRO database from 2.0 to 2.1. Do you wish to continue?"
        'Select Case MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE)
            'Case vbYes
                Call UpGradeData2_0to2_1_1
                sMacroVersion = "2.1"
                sBuildSubVersion = "1"
            'Case vbNo
                'Call ExitMACRO
                'Call MACROEnd
        'End Select
    End If
    
    If sMacroVersion <> "2.1" And sMacroVersion <> "2.2" Then
        MsgBox ("Your Macro database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    'ADD NEW 2.1 UPGRADES HERE
    'Check for "01" as well as "1" only necessary for Sub Version 1
    ' NCj 5/10/00 Go straight from 1 to 3 (version 2 is defunct)
    If sBuildSubVersion = "1" Or sBuildSubVersion = "01" Then
        UpGradeData2_1from1to4
        sBuildSubVersion = "4"
    End If
              
    ' NCJ 5/10/00 - Version 2 must be thrown away
    ' TA 12/10/00 - Version 3 must be thrown away aswell
    If sBuildSubVersion = "2" Or sBuildSubVersion = "3" Then
        sMSG = "Your MACRO database is version 2.1." & sBuildSubVersion & ", which is no longer supported. "
        sMSG = sMSG & vbCrLf & "We cannot upgrade this database, so you'll have to throw it away."
        sMSG = sMSG & vbCrLf & "Sorry about that!"
        Call MsgBox(sMSG, vbApplicationModal, gsDIALOG_TITLE)
        ' Throw them out unless it's System Management (where DB version doesn't matter)
        If App.Title <> "MACRO_SM" Then
            Call ExitMACRO
            Call MACROEnd
        End If
    End If

    If sBuildSubVersion = "4" Or sBuildSubVersion = "5" Then
        UpGradeData2_1from4to6
        sBuildSubVersion = "6"
    End If
   
   ' NCJ 1/11/00 - Upgrade from 6 to 11
   ' NCJ 28/11/00 - Include 7,8,9,10 as well!
    Select Case sBuildSubVersion
    Case "6", "7", "8", "9", "10"
        sBuildSubVersion = "11"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    Case Else
        ' Not relevant here
    End Select
    
   ' NCJ 8/11/00 - Upgrade from 11 to 12
    If sBuildSubVersion = "11" Then
        Call UpGradeData2_1from11to12
        sBuildSubVersion = "12"
    End If
    
    'Mo Morris 20/11/00 - Upgrade from 12 to 13
    If sBuildSubVersion = "12" Then
        Call UpGradeData2_1from12to13
        sBuildSubVersion = "13"
    End If
    
    'Mo Morris 19/12/00 - Upgrade from 13,14,15,16 to 17
    Select Case sBuildSubVersion
    Case sBuildSubVersion = "13", "14", "15", "16"
        sBuildSubVersion = "17"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    Case Else
        ' Not relevant here
    End Select
   
    'TA 30/1/2001: Upgrade from 17 to 30 ' leaving room for Ronald builds
    If Val(sBuildSubVersion) >= 17 And Val(sBuildSubVersion) < 30 Then
        Call UpGradeData2_1from17to30
        sBuildSubVersion = "30"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    ' NCJ 8 Feb 01 - Upgrade to 32
    ' NCJ 19 Feb 01 - Upgrade to 34
    If Val(sBuildSubVersion) >= 30 And Val(sBuildSubVersion) <= 33 Then
        sBuildSubVersion = "34"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
   
   'TA 01/03/2001
    If sBuildSubVersion = "34" Then
        Call UpGradeData2_1from34to35
        sBuildSubVersion = "35"
    End If
   
   ' TA 10/03/2001 - version 35-36
   ' NCJ 23/3/01 - version 36-37
   ' NCJ 2/4/01 - version 38
   ' NCJ 10/4/01 - version 39
   ' TA 18/04/2001 - version 40
    If Val(sBuildSubVersion) >= 35 And Val(sBuildSubVersion) <= 39 Then
        sBuildSubVersion = "40"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

   
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpgradeDataDatabase2_1", "modUpgradeDatabases.bas")
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
Private Sub UpGradeData2_0to2_1_1()
'---------------------------------------------------------------------
' Upgrade 2.0.x (latest) MACRO.mdb to 2.1.1
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
  
    'Upgrade MacroVersion to [2.1]
    sSQL = "UPDATE MACROControl Set MacroVersion = '2.1'"
    moMACROcon.Execute sSQL
    
    'Upgrade BuildSubVersion to [1]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '1'"
    moMACROcon.Execute sSQL
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData2_0to2_1_1", "modUpgradeDatabases.bas")
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
Private Sub UpGradeData2_1from1to4()
'---------------------------------------------------------------------
' NCJ 5/10/00 Renamed from UpGradeData2_1from1to2 since version 2 no longer supported
' Removed long integer identifiers and use codes instead
' TA 12/10/2000 Renamed from UpGradeData2_1from1to3 since version 3 no longer supported
'---------------------------------------------------------------------
Dim sSQL As String
Dim oTableDef As TableDef
Dim oMACRODatabase As Database

    On Error GoTo ErrHandler
    
    'Remove the existing Laboratory table, that existed prior to version 2.1.2, but was not being used
    sSQL = "DROP Table Laboratory"
    moMACROcon.Execute sSQL
    
    'Create the new Laboratory table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table Laboratory (LaboratoryCode TEXT(15),"
        sSQL = sSQL & " LaboratoryDescription TEXT(255),"
        'TA 11/10/2000: two new columns for laboratory
        sSQL = sSQL & " Site TEXT(8), Changed SMALLINT DEFAULT 0,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (LaboratoryCode))"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        oMacroDatabase.TableDefs.Refresh
'        With oMacroDatabase.TableDefs("Laboratory")
'            .Fields("Changed").DefaultValue = 0
'        End With
'
'        Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table Laboratory (LaboratoryCode VARCHAR(15),"
        sSQL = sSQL & " LaboratoryDescription VARCHAR(255),"
        'TA 11/10/2000: two new columns for laboratory
        sSQL = sSQL & " Site VARCHAR(8), Changed INTEGER DEFAULT 0,"
        sSQL = sSQL & " CONSTRAINT PKLaboratory PRIMARY KEY"
        sSQL = sSQL & " (LaboratoryCode))"
        moMACROcon.Execute sSQL
    End If
    
    'Remove the existing NormalRange table, that existed prior to version 2.1.2, but was not being used
    sSQL = "DROP Table NormalRange"
    moMACROcon.Execute sSQL
    
    'Create the new NormalRange table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table NormalRange (NormalRangeId INTEGER,"
        sSQL = sSQL & " LaboratoryCode TEXT(15) , ClinicalTestCode TEXT(15),"
        sSQL = sSQL & " NormalRangeGender SMALLINT,"
        sSQL = sSQL & " NormalRangeAgeMin SMALLINT, NormalRangeAgeMax SMALLINT,"
        sSQL = sSQL & " NormalRangeEffectiveStart DOUBLE, NormalRangeEffectiveEnd DOUBLE,"
        sSQL = sSQL & " NormalRangeNormalMin DOUBLE, NormalRangeNormalMax DOUBLE,"
        sSQL = sSQL & " NormalRangeFeasibleMin DOUBLE, NormalRangeFeasibleMax DOUBLE,"
        sSQL = sSQL & " NormalRangeAbsoluteMin DOUBLE, NormalRangeAbsoluteMax DOUBLE,"
        sSQL = sSQL & " NormalRangePercent SMALLINT,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & "(NormalRangeId,LaboratoryCode))"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        oMacroDatabase.TableDefs.Refresh
        
        sSQL = " CREATE INDEX idx_LaboratoryCode "
        sSQL = sSQL & "ON NormalRange ( LaboratoryCode )"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError

        sSQL = " CREATE INDEX idx_ClinicalTestCode "
        sSQL = sSQL & "ON NormalRange ( ClinicalTestCode )"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
        
'        Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table NormalRange (NormalRangeId INTEGER,"
        sSQL = sSQL & " LaboratoryCode VARCHAR(15), ClinicalTestCode VARCHAR(15),"
        sSQL = sSQL & " NormalRangeGender INTEGER,"
        sSQL = sSQL & " NormalRangeAgeMin INTEGER, NormalRangeAgeMax INTEGER,"
        sSQL = sSQL & " NormalRangeEffectiveStart DECIMAL(16,10), NormalRangeEffectiveEnd DECIMAL(16,10),"
        sSQL = sSQL & " NormalRangeNormalMin DECIMAL(16,10), NormalRangeNormalMax DECIMAL(16,10),"
        sSQL = sSQL & " NormalRangeFeasibleMin DECIMAL(16,10), NormalRangeFeasibleMax DECIMAL(16,10),"
        sSQL = sSQL & " NormalRangeAbsoluteMin DECIMAL(16,10), NormalRangeAbsoluteMax DECIMAL(16,10),"
        sSQL = sSQL & " NormalRangePercent INTEGER,"
        sSQL = sSQL & " CONSTRAINT PKNormalRange PRIMARY KEY "
        sSQL = sSQL & " (NormalRangeId, LaboratoryCode))"
        moMACROcon.Execute sSQL
        
        sSQL = "CREATE INDEX idx_NR_LaboratoryCode "
        sSQL = sSQL & "ON NormalRange ( LaboratoryCode )"
        moMACROcon.Execute sSQL

        sSQL = "CREATE INDEX idx_NR_ClinicalTestCode "
        sSQL = sSQL & "ON NormalRange ( ClinicalTestCode )"
        moMACROcon.Execute sSQL
    End If
    
    
    'Create the new SiteLaboratory table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table SiteLaboratory (Site Text(8),"
        sSQL = sSQL & " LaboratoryCode Text(15),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (Site,LaboratoryCode))"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        oMacroDatabase.TableDefs.Refresh
'
'        Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table SiteLaboratory (Site VARCHAR(8),"
        sSQL = sSQL & " LaboratoryCode VARCHAR(15),"
        sSQL = sSQL & " CONSTRAINT PKSiteLaboratory PRIMARY KEY "
        sSQL = sSQL & " (Site,LaboratoryCode))"
        moMACROcon.Execute sSQL
    End If
    
    
    'Create the new ClinicalTestGroup table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table ClinicalTestGroup (ClinicalTestGroupCode TEXT(15),"
        sSQL = sSQL & " ClinicalTestGroupDescription TEXT(255),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTestGroupCode))"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        oMacroDatabase.TableDefs.Refresh
'
'        Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table ClinicalTestGroup (ClinicalTestGroupCode VARCHAR(15),"
        sSQL = sSQL & " ClinicalTestGroupDescription VARCHAR(255),"
        sSQL = sSQL & " CONSTRAINT PKClinicalTestGroup PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTestGroupCode))"
        moMACROcon.Execute sSQL
    End If
    
    'Create the new ClinicalTest table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table ClinicalTest (ClinicalTestCode TEXT(15),"
        sSQL = sSQL & " ClinicalTestDescription TEXT(255),"
        sSQL = sSQL & " ClinicalTestGroupCode TEXT(15),"
        sSQL = sSQL & " Unit TEXT(15),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTestCode))"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        oMacroDatabase.TableDefs.Refresh
    
        sSQL = " CREATE INDEX idx_ClinicalTestGroupCode "
        sSQL = sSQL & "ON ClinicalTest ( ClinicalTestGroupCode )"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table ClinicalTest (ClinicalTestCode VARCHAR(15),"
        sSQL = sSQL & " ClinicalTestDescription VARCHAR(255),"
        sSQL = sSQL & " ClinicalTestGroupCode VARCHAR(15),"
        sSQL = sSQL & " Unit VARCHAR(15),"
        sSQL = sSQL & " CONSTRAINT PKClinicalTest PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTestCode))"
        moMACROcon.Execute sSQL
    
        sSQL = " CREATE INDEX idx_CT_ClinicalTestGroupCode "
        sSQL = sSQL & "ON ClinicalTest ( ClinicalTestGroupCode )"
        moMACROcon.Execute sSQL
    End If
    
    'Create the new CTCScheme table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table CTCScheme (CTCSchemeCode TEXT(15),"
        sSQL = sSQL & " CTCSchemeDescription TEXT(255),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (CTCSchemeCode))"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        oMacroDatabase.TableDefs.Refresh
'
'        Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table CTCScheme (CTCSchemeCode VARCHAR(15),"
        sSQL = sSQL & " CTCSchemeDescription VARCHAR(255),"
        sSQL = sSQL & " CONSTRAINT PKCTCScheme PRIMARY KEY "
        sSQL = sSQL & " (CTCSchemeCode))"
        moMACROcon.Execute sSQL
    End If
    
    
    'Create the new CTC table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table CTC (CTCId INTEGER ,"
        sSQL = sSQL & " CTCSchemeCode TEXT(15),"
        sSQL = sSQL & " ClinicalTestCode TEXT(15),"
        sSQL = sSQL & " CTCGrade SMALLINT,"
        sSQL = sSQL & " CTCMin DOUBLE,"
        sSQL = sSQL & " CTCMax DOUBLE,"
        sSQL = sSQL & " CTCMinType SMALLINT,"
        sSQL = sSQL & " CTCMaxType SMALLINT,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (CTCId,CTCSchemeCode))"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        oMacroDatabase.TableDefs.Refresh
    
        sSQL = " CREATE INDEX idx_CTCSchemeCode "
        sSQL = sSQL & "ON CTC ( CTCSchemeCode )"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        sSQL = " CREATE INDEX idx_ClinicalTestCode "
        sSQL = sSQL & "ON CTC ( ClinicalTestCode )"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
                
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table CTC (CTCId INTEGER,"
        sSQL = sSQL & " CTCSchemeCode VARCHAR(15),"
        sSQL = sSQL & " ClinicalTestCode VARCHAR(15),"
        sSQL = sSQL & " CTCGrade INTEGER,"
        sSQL = sSQL & " CTCMin DECIMAL(16,10),"
        sSQL = sSQL & " CTCMax DECIMAL(16,10),"
        sSQL = sSQL & " CTCMinType INTEGER,"
        sSQL = sSQL & " CTCMaxType INTEGER,"
        sSQL = sSQL & " CONSTRAINT PKCTC PRIMARY KEY "
        sSQL = sSQL & " (CTCId,CTCSchemeCode))"
        moMACROcon.Execute sSQL
    
        sSQL = " CREATE INDEX idx_CTC_CTCSchemeCode "
        sSQL = sSQL & "ON CTC ( CTCSchemeCode )"
        moMACROcon.Execute sSQL
        
        sSQL = " CREATE INDEX idx_CTC_ClinicalTestCode "
        sSQL = sSQL & "ON CTC ( ClinicalTestCode )"
        moMACROcon.Execute sSQL
    End If
    
    
    'Amend the StudyDefinition table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "ALTER Table StudyDefinition ADD COLUMN CTCSchemeCode TEXT(15)"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        sSQL = "ALTER Table StudyDefinition ADD COLUMN DOBExpr MEMO"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        sSQL = "ALTER Table StudyDefinition ADD COLUMN GenderExpr MEMO"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        
        'Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table StudyDefinition ADD CTCSchemeCode VARCHAR(15)"
        moMACROcon.Execute sSQL
        'Oracle / SQL server specifics
        If meDatabaseType = MACRODatabaseType.sqlserver Then
            sSQL = "ALTER Table StudyDefinition ADD DOBExpr TEXT"
            moMACROcon.Execute sSQL
            sSQL = "ALTER Table StudyDefinition ADD GenderExpr TEXT"
            moMACROcon.Execute sSQL
        ElseIf meDatabaseType = MACRODatabaseType.Oracle80 Then
            sSQL = "ALTER Table StudyDefinition ADD DOBExpr VARCHAR(2000)"
            moMACROcon.Execute sSQL
            sSQL = "ALTER Table StudyDefinition ADD GenderExpr VARCHAR(2000)"
            moMACROcon.Execute sSQL
        End If
    
    End If
    
    
    'Amend the DataItem table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "ALTER Table DataItem ADD COLUMN ClinicalTestCode TEXT(15)"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        
        'Set oMacroDatabase = Nothing
    
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table DataItem ADD ClinicalTestCode VARCHAR(15)"
        moMACROcon.Execute sSQL
        
    End If
    
    
    'Amend the CRFElement table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table CRFElement ADD COLUMN ClinicalTestDateExpr Memo"
        moMACROcon.Execute sSQL
    Else    'SQL or Oracle, using ADO
        'Oracle / SQL server specifics
        If meDatabaseType = MACRODatabaseType.sqlserver Then
            sSQL = "ALTER Table CRFElement ADD ClinicalTestDateExpr TEXT"
        ElseIf meDatabaseType = MACRODatabaseType.Oracle80 Then
            sSQL = "ALTER Table CRFElement ADD ClinicalTestDateExpr VARCHAR(2000)"
        End If
        moMACROcon.Execute sSQL
    End If
    
    
    'Amend the CRFPageInstance table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "ALTER Table CRFPageInstance ADD COLUMN LaboratoryCode TEXT(15)"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
            
        'Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table CRFPageInstance ADD LaboratoryCode VARCHAR(15)"
        moMACROcon.Execute sSQL
            
    End If
    
    'Amend the TrialSubject table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        ' TA 12/10/2000: no longer deleting the current Gender and DateofBirth columns - they should not be used though
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "DROP INDEX idx_Gender on TrialSubject"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        sSQL = "Alter Table TrialSubject ADD COLUMN SubjectGender SMALLINT"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "DROP INDEX idx_Gender"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table TrialSubject ADD SubjectGender INTEGER"
        moMACROcon.Execute sSQL
    End If
    
    
    'Amend the DataItemResponse table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table DataItemResponse ADD COLUMN LabResult TEXT(1)"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table DataItemResponse ADD COLUMN CTCGrade SMALLINT"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table DataItemResponse ADD COLUMN ClinicalTestDate DOUBLE"
        moMACROcon.Execute sSQL
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table DataItemResponse ADD LabResult VARCHAR(1)"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table DataItemResponse ADD CTCGrade INTEGER"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table DataItemResponse ADD ClinicalTestDate DECIMAL(16,10)"
        moMACROcon.Execute sSQL
    End If
    
    
    'Amend the DataItemResponseHistory table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table DataItemResponseHistory ADD COLUMN LabResult TEXT(1)"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD COLUMN CTCGrade SMALLINT"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD COLUMN ClinicalTestDate DOUBLE"
        moMACROcon.Execute sSQL
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table DataItemResponseHistory ADD LabResult VARCHAR(1)"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD CTCGrade INTEGER"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD ClinicalTestDate DECIMAL(16,10)"
        moMACROcon.Execute sSQL
    End If
    
    'TA 12/10/2000: Create the new NewDbColumn table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'REM 07/06/02 - Changed to ADO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        Call CreateNewDBColumn2_1_4(moMACROcon, True)
        'Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        Call CreateNewDBColumn2_1_4(moMACROcon, False)
    End If

    
    'TA 12/10/2000: Create the new SiteUser table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table SiteUser (Site Text(8),"
        sSQL = sSQL & " UserCode Text(15),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (Site,UserCode))"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        
        'Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table SiteUser (Site VARCHAR(8),"
        sSQL = sSQL & " UserCode VARCHAR(15),"
        sSQL = sSQL & " CONSTRAINT PKSiteUser PRIMARY KEY "
        sSQL = sSQL & " (Site,UserCode))"
        moMACROcon.Execute sSQL
    End If
    
    'TA 12/10/2000: MACROTAble Table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'REM 07/06/02 - Changed to ADO
        Call CreateMacroTablev2_1_4(moMACROcon, True)
    Else    'SQL or Oracle, using ADO
        Call CreateMacroTablev2_1_4(moMACROcon, False)
    End If

    'TA 13/10/2000: drop old import/export tables
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        Set oMACRODatabase = DBEngine.Workspaces(0).OpenDatabase(msDatabaseLocation, False, False, "MS Access;PWD=" & msDatabasePswd)

        oMACRODatabase.Execute "DROP TABLE SDDExportImport", dbFailOnError
        oMACRODatabase.Execute "DROP TABLE PRDExportImport", dbFailOnError

        oMACRODatabase.TableDefs.Refresh

        Set oMACRODatabase = Nothing

    Else    'SQL or Oracle, using ADO
        moMACROcon.Execute "DROP TABLE SDDExportImport"
        moMACROcon.Execute "DROP TABLE PRDExportImport"
    End If

    'Upgrade BuildSubVersion to [4]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '4'"
    moMACROcon.Execute sSQL


Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeData2_1from1to4", "modUpgradeDatabases.bas")
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
Private Sub UpGradeData2_1from4to6()
'---------------------------------------------------------------------
' NCJ 25/10/00 Upgrade 2.1. from 4 to 6
' Add SegmentIds for Units and UnitConversionFactors
'---------------------------------------------------------------------
Dim sSQL As String

    sSQL = "UPDATE MACROTable Set SegmentId = '550', STYDEF = 1 WHERE TableName = 'Units'"
    moMACROcon.Execute sSQL

    sSQL = "UPDATE MACROTable Set SegmentId = '560', STYDEF = 1 WHERE TableName = 'UnitConversionFactors'"
    moMACROcon.Execute sSQL
    
    'Upgrade BuildSubVersion to [6]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '6'"
    moMACROcon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeData2_1from4to6", "modUpgradeDatabases")
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
Private Sub UpGradeData2_1from11to12()
'---------------------------------------------------------------------
' NCJ 25/10/00 Upgrade 2.1. from 11 to 12
' Add ClinicalTest, ClinicalTestGroup and Units to LDD export
'---------------------------------------------------------------------
Dim sSQL As String

    sSQL = "UPDATE MACROTable Set LABDEF = 1 WHERE TableName = 'Units'"
    moMACROcon.Execute sSQL

    sSQL = "UPDATE MACROTable Set LABDEF = 1 WHERE TableName = 'ClinicalTest'"
    moMACROcon.Execute sSQL
    
    sSQL = "UPDATE MACROTable Set LABDEF = 1 WHERE TableName = 'ClinicalTestGroup'"
    moMACROcon.Execute sSQL
    
    'Upgrade BuildSubVersion to [12]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '12'"
    moMACROcon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeData2_1from11to12", "modUpgradeDatabases")
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
Private Sub UpGradeData2_1from12to13()
'---------------------------------------------------------------------
Dim sSQL As String
Dim oMACRODatabase As Database

    On Error GoTo ErrHandler
    
    'Create the new Eligibility table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table Eligibility ("
        sSQL = sSQL & " ClinicalTrialId INTEGER,"
        sSQL = sSQL & " VersionId SMALLINT,"
        sSQL = sSQL & " EligibilityCode TEXT(15),"
        sSQL = sSQL & " RandomisationCode TEXT(15),"
        sSQL = sSQL & " Flag SMALLINT,"
        sSQL = sSQL & " Condition MEMO,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialId,VersionId,EligibilityCode))"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table Eligibility ("
        sSQL = sSQL & " ClinicalTrialId INTEGER,"
        sSQL = sSQL & " VersionId INTEGER,"
        sSQL = sSQL & " EligibilityCode VARCHAR(15),"
        sSQL = sSQL & " RandomisationCode VARCHAR(15),"
        sSQL = sSQL & " Flag INTEGER,"
        If meDatabaseType = MACRODatabaseType.sqlserver Then
            sSQL = sSQL & " Condition  TEXT,"
        ElseIf meDatabaseType = MACRODatabaseType.Oracle80 Then
            sSQL = sSQL & " Condition  VARCHAR(2000),"
        End If
        sSQL = sSQL & " CONSTRAINT PKEligibility PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialId,VersionId,EligibilityCode))"
        moMACROcon.Execute sSQL
    End If
    
    
    'Create the new UniquenessCheck table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table UniquenessCheck ("
        sSQL = sSQL & " ClinicalTrialId INTEGER,"
        sSQL = sSQL & " VersionId SMALLINT,"
        sSQL = sSQL & " CheckCode TEXT(15),"
        sSQL = sSQL & " Expression MEMO,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialId,VersionId,CheckCode))"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table UniquenessCheck ("
        sSQL = sSQL & " ClinicalTrialId INTEGER,"
        sSQL = sSQL & " VersionId INTEGER,"
        sSQL = sSQL & " CheckCode VARCHAR(15),"
        If meDatabaseType = MACRODatabaseType.sqlserver Then
            sSQL = sSQL & " Expression  TEXT,"
        ElseIf meDatabaseType = MACRODatabaseType.Oracle80 Then
            sSQL = sSQL & " Expression  VARCHAR(2000),"
        End If
        sSQL = sSQL & " CONSTRAINT PKUniquenessCheck PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialId,VersionId,CheckCode))"
        moMACROcon.Execute sSQL
    End If
    
    
    'Create the new SubjectNumbering table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table SubjectNumbering ("
        sSQL = sSQL & " ClinicalTrialId INTEGER,"
        sSQL = sSQL & " VersionId SMALLINT,"
        sSQL = sSQL & " StartNumber INTEGER,"
        sSQL = sSQL & " NumberWidth SMALLINT,"
        sSQL = sSQL & " Prefix MEMO,"
        sSQL = sSQL & " UsePrefix SMALLINT,"
        sSQL = sSQL & " Suffix MEMO,"
        sSQL = sSQL & " UseSuffix SMALLINT,"
        sSQL = sSQL & " TriggerVisitId INTEGER,"
        sSQL = sSQL & " TriggerFormId INTEGER,"
        sSQL = sSQL & " UseRegistration SMALLINT,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialId,VersionId))"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table SubjectNumbering ("
        sSQL = sSQL & " ClinicalTrialId INTEGER,"
        sSQL = sSQL & " VersionId INTEGER,"
        sSQL = sSQL & " StartNumber INTEGER,"
        sSQL = sSQL & " NumberWidth INTEGER,"
        If meDatabaseType = MACRODatabaseType.sqlserver Then
            sSQL = sSQL & " Prefix  TEXT,"
        ElseIf meDatabaseType = MACRODatabaseType.Oracle80 Then
            sSQL = sSQL & " Prefix  VARCHAR(2000),"
        End If
        sSQL = sSQL & " UsePrefix INTEGER,"
        If meDatabaseType = MACRODatabaseType.sqlserver Then
            sSQL = sSQL & " Suffix  TEXT,"
        ElseIf meDatabaseType = MACRODatabaseType.Oracle80 Then
            sSQL = sSQL & " Suffix  VARCHAR(2000),"
        End If
        sSQL = sSQL & " UseSuffix INTEGER,"
        sSQL = sSQL & " TriggerVisitId INTEGER,"
        sSQL = sSQL & " TriggerFormId INTEGER,"
        sSQL = sSQL & " UseRegistration INTEGER,"
        sSQL = sSQL & " CONSTRAINT PKSubjectNumbering PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialId,VersionId))"
        moMACROcon.Execute sSQL
    End If
    
    
    'Create the new RSSubjectIdentifier table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table RSSubjectIdentifier ("
        sSQL = sSQL & " ClinicalTrialName TEXT(15),"
        sSQL = sSQL & " TrialSite TEXT(8),"
        sSQL = sSQL & " PersonId INTEGER,"
        sSQL = sSQL & " SubjectIdentifier TEXT(255),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialName,TrialSite,PersonId))"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table RSSubjectIdentifier ("
        sSQL = sSQL & " ClinicalTrialName VARCHAR(15),"
        sSQL = sSQL & " TrialSite VARCHAR(8),"
        sSQL = sSQL & " PersonId INTEGER,"
        sSQL = sSQL & " SubjectIdentifier VARCHAR(255),"
        sSQL = sSQL & " CONSTRAINT PKRSSubjectIdentifier PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialName,TrialSite,PersonId))"
        moMACROcon.Execute sSQL
    End If
    
    
    'Create the new RSNextNumber table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table RSNextNumber ("
        sSQL = sSQL & " ClinicalTrialName TEXT(15),"
        sSQL = sSQL & " NextNumberId INTEGER,"
        sSQL = sSQL & " Prefix TEXT(255),"
        sSQL = sSQL & " Suffix TEXT(255),"
        sSQL = sSQL & " NextNumber INTEGER,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialName,NextNumberId))"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table RSNextNumber ("
        sSQL = sSQL & " ClinicalTrialName VARCHAR(15),"
        sSQL = sSQL & " NextNumberId INTEGER,"
        sSQL = sSQL & " Prefix VARCHAR(255),"
        sSQL = sSQL & " Suffix VARCHAR(255),"
        sSQL = sSQL & " NextNumber INTEGER,"
        sSQL = sSQL & " CONSTRAINT PKRSNextNumber PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialName,NextNumberId))"
        moMACROcon.Execute sSQL
    End If
    
    
    'Create the new RSUniquenessCheck table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table RSUniquenessCheck ("
        sSQL = sSQL & " ClinicalTrialName TEXT(15),"
        sSQL = sSQL & " TrialSite TEXT(8),"
        sSQL = sSQL & " PersonId INTEGER,"
        sSQL = sSQL & " CheckCode TEXT(15),"
        sSQL = sSQL & " CheckValue TEXT(255),"
        sSQL = sSQL & " CheckTimeStamp DOUBLE,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialName,TrialSite,PersonId,CheckCode))"
        moMACROcon.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table RSUniquenessCheck ("
        sSQL = sSQL & " ClinicalTrialName VARCHAR(15),"
        sSQL = sSQL & " TrialSite VARCHAR(8),"
        sSQL = sSQL & " PersonId INTEGER,"
        sSQL = sSQL & " CheckCode VARCHAR(15),"
        sSQL = sSQL & " CheckValue VARCHAR(255),"
        sSQL = sSQL & " CheckTimeStamp DECIMAL(16,10),"
        sSQL = sSQL & " CONSTRAINT PKRSUniquenessCheck PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialName,TrialSite,PersonId,CheckCode))"
        moMACROcon.Execute sSQL
    End If
    
    'Amend the TrialSubject table, add field RegistrationStatus
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table TrialSubject ADD COLUMN RegistrationStatus SMALLINT"
        moMACROcon.Execute sSQL
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table TrialSubject ADD RegistrationStatus INTEGER"
        moMACROcon.Execute sSQL
    End If
    
    
    'Amend the MIMessage table, add field MIMessageResponseTimeStamp
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table MIMessage ADD COLUMN MIMessageResponseTimeStamp DOUBLE"
        moMACROcon.Execute sSQL
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table MIMessage ADD MIMessageResponseTimeStamp DECIMAL(16,10)"
        moMACROcon.Execute sSQL
    End If
    
    
    'Amend the DataItemResponse table, add field LaboratoryCode
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table DataItemResponse ADD COLUMN LaboratoryCode TEXT(15)"
        moMACROcon.Execute sSQL
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table DataItemResponse ADD LaboratoryCode VARCHAR(15)"
        moMACROcon.Execute sSQL
    End If
    
    'Amend the DataItemResponseHistory table, add field LaboratoryCode
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table DataItemResponseHistory ADD COLUMN LaboratoryCode TEXT(15)"
        moMACROcon.Execute sSQL
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table DataItemResponseHistory ADD LaboratoryCode VARCHAR(15)"
        moMACROcon.Execute sSQL
    End If
    
    'Amend the StudyDfinition table
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table StudyDefinition ADD COLUMN RRServerType SMALLINT"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD COLUMN RRHTTPAddress TEXT(255)"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD COLUMN RRUserName TEXT(50)"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD COLUMN RRPassword TEXT(50)"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD COLUMN RRProxyServer TEXT(255)"
        moMACROcon.Execute sSQL
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table StudyDefinition ADD RRServerType INTEGER"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD RRHTTPAddress VARCHAR(255)"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD RRUserName VARCHAR(50)"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD RRPassword VARCHAR(50)"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD RRProxyServer VARCHAR(255)"
        moMACROcon.Execute sSQL
    End If
    
    'add new tables to table MACROTable
    sSQL = "INSERT INTO MACROTable VALUES ('Eligibility', '250', 1, 0, 0)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('UniquenessCheck', '260', 1, 0, 0)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('SubjectNumbering', '270', 1, 0, 0)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('RSSubjectIdentifier', '', 0, 0, 0)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('RSNextNumber', '', 0, 0, 0)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('RSUniquenessCheck', '', 0, 0, 0)"
    moMACROcon.Execute sSQL
    
    
    'add tables changes to table NewDBColumn
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'TrialSubject','RegistrationStatus',null,'#NULL#')"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'MIMessage','MIMessageResponseTimeStamp',null,'#NULL#')"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'DataItemResponse','LaboratoryCode',null,'#NULL#')"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'DataItemResponseHistory','LaboratoryCode',null,'#NULL#')"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'StudyDefinition','RRServerType',1,'#NULL#')"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'StudyDefinition','RRHTTPAddress',2,'#NULL#')"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'StudyDefinition','RRUserName',3,'#NULL#')"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'StudyDefinition','RRPassword',4,'#NULL#')"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'StudyDefinition','RRProxyServer',5,'#NULL#')"
    moMACROcon.Execute sSQL
    
    
    'Upgrade BuildSubVersion to [13]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '13'"
    moMACROcon.Execute sSQL
    

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeData2_1from12to13", "modUpgradeDatabases.bas")
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
Private Sub UpGradeData2_1from17to30()
'---------------------------------------------------------------------
Dim sSQL As String
Dim oMACRODatabase As Database

    On Error GoTo ErrHandler
    
    'TA 30/1/2001 SR4131: allow zero length on TrialSubject.Gender in Access
    If meDatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        Set oMACRODatabase = DBEngine.Workspaces(0).OpenDatabase(msDatabaseLocation, False, False, "MS Access;PWD=" & msDatabasePswd)
        oMACRODatabase.TableDefs.Refresh
        oMACRODatabase.TableDefs("TrialSubject").Fields("Gender").AllowZeroLength = True
        Set oMACRODatabase = Nothing
    End If

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData2_1from17to30", "modUpgradeDatabases.bas")
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
Private Sub UpGradeData2_1from34to35()
'---------------------------------------------------------------------
Dim sSQL As String


    On Error GoTo ErrHandler

    'TA 01/03/2001: Lab data type for reporting use
    sSQL = "Insert Into DataType (DataTypeId, DataTypeName) values (6,'Laboratory Test')"
    moMACROcon.Execute sSQL
    
    'update MACROControl
    Call UpGradeDataToSubVersion("35")

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData2_1from34to35", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'------------------------------------------------------------------------------'
Private Function CreateNewDBColumn2_1_4(oExecuter As Object, bAccess As Boolean)
'------------------------------------------------------------------------------'
'TA 12/10/2000: New CreateNewDBColumn table - stores added columns and the version it happened
'------------------------------------------------------------------------------'
Dim sSQLStart As String
Dim sSQL As String
    On Error GoTo ErrHandler
    
    'TA 12/10/2000: New MACROTable table - stores all tables
    If bAccess Then
        'Using DAO
        sSQL = "CREATE Table NewDBColumn (VersionMajor SMALLINT, VersionMinor SMALLINT, VersionRevision SMALLINT,"
        sSQL = sSQL & " TableName TEXT(30), ColumnName TEXT(30), ColumnOrder SMALLINT, DefaultValue TEXT(255),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (VersionMajor, VersionMinor, VersionRevision, TableName, ColumnName))"
        'REM 07/06/02 - changed to ADO connection
'        With oExecuter
'            .Execute sSQL, dbFailOnError
'            .TableDefs.Refresh
'        End With
        
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table NewDBColumn (VersionMajor INTEGER, VersionMinor INTEGER, VersionRevision INTEGER,"
        sSQL = sSQL & " TableName VARCHAR(30), ColumnName VARCHAR(30), ColumnOrder INTEGER, DefaultValue VARCHAR(255),"
        sSQL = sSQL & " CONSTRAINT PKNewDBColumn PRIMARY KEY"
        sSQL = sSQL & " (VersionMajor, VersionMinor, VersionRevision, TableName, ColumnName))"

    End If
    oExecuter.Execute sSQL
    
    sSQLStart = "INSERT INTO NewDBColumn VALUES ("
    
    With oExecuter
        .Execute sSQLStart & "2,1,4,'CRFElement','ClinicalTestDate',null,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'DataItem','ClinicalTestDate',null,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'StudyDefinition','CTCSchemeCode',1,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'StudyDefinition','DOBExpr',2,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'StudyDefinition','GenderExpr',3,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'Trialsubject','SubjectGender',null,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'CRFPageInstance','LaboratoryCode',null,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'DataItemResponse','LabResult',1,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'DataItemResponse','CTCGrade',2,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'DataItemResponse','ClinicalTestDate',3,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'DataItemResponseHistory','LabResult',1,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'DataItemResponseHistory','CTCGrade',2,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'DataItemResponseHistory','ClinicalTestDate',3,'#NULL#')"
    End With

Exit Function
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CreateNewDBColumn2_1_4", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'------------------------------------------------------------------------------'
Private Function CreateMacroTablev2_1_4(oExecuter As Object, bAccess As Boolean)
'------------------------------------------------------------------------------'
'TA 12/10/2000: New MACROTable table - stores all tables
' NCJ 8/11/00 - Include Units, ClinicalTest and ClinicalTestGroup in LDD
'------------------------------------------------------------------------------'
Dim sSQLStart As String
Dim sSQL As String
    On Error GoTo ErrHandler
    
    'TA 12/10/2000: New MACROTable table - stores all tables
    If bAccess Then
        'Using DAO
        sSQL = "CREATE Table MACROTable (TableName TEXT(30),"
        sSQL = sSQL & " SegmentId TEXT(3),"
        sSQL = sSQL & " STYDEF SMALLINT, PATRSP SMALLINT, LABDEF SMALLINT,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (TableName))"
        'REM 07/06/02 - changed to use ADO connection
'        With oExecuter
'            .Execute sSQL, dbFailOnError
'            .TableDefs.Refresh
'        End With
        
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table MACROTable (TableName VARCHAR(30),"
        sSQL = sSQL & " SegmentId VARCHAR(3),"
        sSQL = sSQL & " STYDEF INTEGER, PATRSP INTEGER, LABDEF INTEGER,"
        sSQL = sSQL & " CONSTRAINT PKMACROTable PRIMARY KEY "
        sSQL = sSQL & " (TableName))"
        
    End If
    oExecuter.Execute sSQL
    
    sSQLStart = "INSERT INTO MACROTable VALUES ("
    
'3rd column for SDD, 4th for PRD, 5th for LDD
' NCJ 25/10/00 - Added new values for Units and UnitConversionFactors
' NCJ 8/11/00 - Include Units, ClinicalTest and ClinicalTestGroup in LDD
With oExecuter
        .Execute sSQLStart & "'ClinicalTest', '620', 1, 0, 1)"
        .Execute sSQLStart & "'ClinicalTestGroup', '610', 1, 0, 1)"
        .Execute sSQLStart & "'ClinicalTrial', '001', 1, 0, 0)"
        .Execute sSQLStart & "'CRFElement', '140', 1, 0, 0)"
        .Execute sSQLStart & "'CRFPage', '130', 1, 0, 0)"
        .Execute sSQLStart & "'CRFPageInstance', '070', 0, 1, 0)"
        .Execute sSQLStart & "'CTC', '640', 1, 0, 0)"
        .Execute sSQLStart & "'CTCScheme', '630', 1, 0, 0)"
        .Execute sSQLStart & "'DataItem', '120', 1, 0, 0)"
        .Execute sSQLStart & "'DataItemResponse', '080', 0, 1, 0)"
        .Execute sSQLStart & "'DataItemResponseHistory', '090', 0, 1, 0)"
        .Execute sSQLStart & "'DataItemValidation', '125', 1, 0, 0)"
        .Execute sSQLStart & "'DataType', '', 0, 0, 0)"
        .Execute sSQLStart & "'ExternalDataMapping', '', 0, 0, 0)"
        .Execute sSQLStart & "'ExternalDataSource', '', 0, 0, 0)"
        .Execute sSQLStart & "'Laboratory', '001', 0, 0, 1)"
        .Execute sSQLStart & "'LogDetails', '', 0, 0, 0)"
        .Execute sSQLStart & "'MACROControl', '', 0, 0, 0)"
        .Execute sSQLStart & "'MACROTable', '', 0, 0, 0)"
        .Execute sSQLStart & "'Message', '', 0, 0, 0)"
        .Execute sSQLStart & "'MIMessage', '', 0, 0, 0)"
        .Execute sSQLStart & "'NewDBColumn', '', 0, 0, 0)"
        .Execute sSQLStart & "'NormalRange', '030', 0, 0, 1)"
'        .Execute sSQLStart & "'PRDExportImport', '', 0, 0, 0)"
        .Execute sSQLStart & "'Protocols', '300', 1, 0, 0)"
        .Execute sSQLStart & "'RandomisationStep', '170', 1, 0, 0)"
        .Execute sSQLStart & "'ReasonForChange', '060', 1, 0, 0)"
        .Execute sSQLStart & "'ReportType', '', 0, 0, 0)"
        .Execute sSQLStart & "'RequiredData', '', 0, 0, 0)"
'        .Execute sSQLStart & "'SDDExportImport', '', 0, 0, 0)"
        .Execute sSQLStart & "'Site', '', 0, 0, 0)"
        .Execute sSQLStart & "'SiteLaboratory', '', 0, 0, 0)"
        .Execute sSQLStart & "'SiteUser', '', 0, 0, 0)"
        .Execute sSQLStart & "'StandardDataFormat', '', 0, 0, 0)"
        .Execute sSQLStart & "'StratificationFactor', '180', 1, 0, 0)"
        .Execute sSQLStart & "'StudyDefinition', '030', 1, 0, 0)"
        .Execute sSQLStart & "'StudyDocument', '040', 1, 0, 0)"
        .Execute sSQLStart & "'StudyReport', '190', 1, 0, 0)"
        .Execute sSQLStart & "'StudyReportData', '200', 1, 0, 0)"
        .Execute sSQLStart & "'StudyVisit', '150', 1, 0, 0)"
        .Execute sSQLStart & "'StudyVisitCRFPage', '160', 1, 0, 0)"
        .Execute sSQLStart & "'TrialOffice', '', 0, 0, 0)"
        .Execute sSQLStart & "'TrialPhase', '510', 1, 0, 0)"
        .Execute sSQLStart & "'TrialSite', '', 0, 0, 0)"
        .Execute sSQLStart & "'TrialStatus', '', 0, 0, 0)"
        .Execute sSQLStart & "'TrialStatusHistory', '050', 1, 0, 0)"
        .Execute sSQLStart & "'TrialSubject', '040', 0, 1, 0)"
        .Execute sSQLStart & "'TrialType', '520', 1, 0, 0)"
        .Execute sSQLStart & "'UnitClasses', '', 0, 0, 0)"
        .Execute sSQLStart & "'UnitConversionFactors', '560', 1, 0, 0)"
        .Execute sSQLStart & "'Units', '550', 1, 0, 1)"
        .Execute sSQLStart & "'ValidationAction', '530', 1, 0, 0)"
        .Execute sSQLStart & "'ValidationType', '540', 1, 0, 0)"
        .Execute sSQLStart & "'ValueData', '090', 1, 0, 0)"
        .Execute sSQLStart & "'VisitInstance', '060', 0, 1, 0)"
    End With

Exit Function
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CreateMacroTablev2_1_4", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Private Sub UpgradeDataDatabase2_2()
'---------------------------------------------------------------------
' Upgrade a 2.2 MACRO.mdb, checking first for 2.1 -> 2.2
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sMacroVersion As String
Dim sBuildSubVersion As String
Dim sMSG As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM MACROControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, moMACROcon, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        MsgBox ("Your Macro database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    sBuildSubVersion = rsTemp![BuildSubVersion]
    sMacroVersion = rsTemp![MACROVersion]
    rsTemp.Close
    Set rsTemp = Nothing

    'don't do anything if future version
    If Val(sMacroVersion) > Val(2.2) Then
        Exit Sub
    End If

    ' Check for version 2.1
    If sMacroVersion = "2.1" Then
        'sMsg = "You are about to upgrade your MACRO database from 2.1 to 2.2. Do you wish to continue?"
        'Select Case MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE)
            'Case vbYes
                Call UpGradeData2_1to2_2_1
                sMacroVersion = "2.2"
                sBuildSubVersion = "1"
            'Case vbNo
                'Call ExitMACRO
                'Call MACROEnd
        'End Select
    End If
    
    If sMacroVersion <> "2.2" Then
        MsgBox ("Your MACRO database is not valid. MACRO is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    'ADD NEW 2.2 UPGRADES HERE
    ' NCJ 19 Oct 01 - 1 to 2
    If sBuildSubVersion = "1" Then
        sBuildSubVersion = "2"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'REM 26/10/01 - 2 to 3
    If sBuildSubVersion = "2" Then
        sBuildSubVersion = "3"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'REM 20/11/01 - 3 to 4
    If sBuildSubVersion = "3" Then
        sBuildSubVersion = "4"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'RJCW 26/11/01 - 4 to 5
    If sBuildSubVersion = "4" Then
        sBuildSubVersion = "5"
        Call UpGradeData2_2from1to5
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'MLM 20/12/01 - 5 to 6
    If sBuildSubVersion = "5" Then
        sBuildSubVersion = "6"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'REM 11/01/02 - 6 to 7
    If sBuildSubVersion = "6" Then
        sBuildSubVersion = "7"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'REM 18/01/02 - 7 to 8
    If sBuildSubVersion = "7" Then
        sBuildSubVersion = "8"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'REM 31/01/02 - 8 to 9
    If sBuildSubVersion = "8" Then
        sBuildSubVersion = "9"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'REM 15/04/02 - 9 to 10
    If sBuildSubVersion = "9" Then
        sBuildSubVersion = "10"
        Call UpGradeData2_2from5to10
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'MLM 28/04/02 - 10 to 11
    If sBuildSubVersion = "10" Then
        sBuildSubVersion = "11"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'MLM 30/04/02 - 11 to 12
    If sBuildSubVersion = "11" Then
        sBuildSubVersion = "12"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'MLM 24/05/02 - 12 to 13
    If sBuildSubVersion = "12" Then
        sBuildSubVersion = "13"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'MLM 24/05/02 - 13 to 14 (whoops, forgot 13 upgrade)
    If sBuildSubVersion = "13" Then
        sBuildSubVersion = "14"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'REM 10/06/02 - 14 to 15
    If sBuildSubVersion = "14" Then
        sBuildSubVersion = "15"
        Call UpGradeData2_2from14to15
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'MLM 25/06/02 - 15 to 18 (16 & 17 were used for patches - no db upgrade)
    If sBuildSubVersion = "15" Then
        sBuildSubVersion = "18"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'MLM 28/06/02 - 18 to 19
    If sBuildSubVersion = "18" Then
        sBuildSubVersion = "19"
        Call UpgradeData2_2from18to19
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'REM 31/07/03 - 19 to 21
    If sBuildSubVersion = "19" Then
        sBuildSubVersion = "21"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpgradeDataDatabase2_2", "modUpgradeDatabases.bas")
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
Private Sub UpGradeData2_1to2_2_1()
'---------------------------------------------------------------------
' Upgrade 2.1.x (latest) to 2.2.1 (Includes all of the 2.1 to 2.2 database audit changes
'---------------------------------------------------------------------
Dim oCat As Catalog
Dim sSQL As String
Dim sDefaultConstraintName As String

    On Error GoTo ErrHandler
    
    ' *** Drop Tables ***
    sSQL = "DROP Table ExternalDataMapping"
    moMACROcon.Execute sSQL
    sSQL = "DROP Table ExternalDataSource"
    moMACROcon.Execute sSQL
    sSQL = "DROP Table RandomisationStep"
    moMACROcon.Execute sSQL
    sSQL = "DROP Table ReportType"
    moMACROcon.Execute sSQL
    sSQL = "DROP Table StratificationFactor"
    moMACROcon.Execute sSQL
    
    ' *** Remove Dropped tables from table MACROTable ***
    sSQL = "DELETE  FROM MACROTable WHERE TableName = 'ExternalDataMapping'"
    moMACROcon.Execute sSQL
    sSQL = "DELETE  FROM MACROTable WHERE TableName = 'ExternalDataSource'"
    moMACROcon.Execute sSQL
    sSQL = "DELETE  FROM MACROTable WHERE TableName = 'RandomisationStep'"
    moMACROcon.Execute sSQL
    sSQL = "DELETE  FROM MACROTable WHERE TableName = 'ReportType'"
    moMACROcon.Execute sSQL
    sSQL = "DELETE  FROM MACROTable WHERE TableName = 'StratificationFactor'"
    moMACROcon.Execute sSQL
    
    ' *** Add New Tables ***
    'Mo Morris 1/10/01, optional field to prevent CreateDB displaying messages added
    Call CreateDB(eMACRODBFunction.Data, meDatabaseType, moMACROcon, False, False, "MACROLock", "2.2.1")
    
    ' *** Add New Tables to table MACROTable ***
    sSQL = "INSERT INTO MACROTable VALUES ('MACROLock', '', 0, 0, 0)"
    moMACROcon.Execute sSQL
    
    ' *** New Columns ***
    'Add eFormDatePrompt to table CRFPage
    Select Case meDatabaseType
    Case MACRODatabaseType.Access
        sSQL = "ALTER Table CRFPage ADD COLUMN eFormDatePrompt BYTE"
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sSQL = "ALTER Table CRFPage ADD eFormDatePrompt TINYINT"
    Case MACRODatabaseType.Oracle80
        sSQL = "ALTER Table CRFPage ADD eFormDatePrompt NUMBER(3)"
    End Select
    moMACROcon.Execute sSQL
    
    'Add ChangeType and ColumnNumber to table NewDBColumn
    Select Case meDatabaseType
    Case MACRODatabaseType.Access
        sSQL = "ALTER Table NewDBColumn ADD COLUMN ChangeType TEXT(15)"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table NewDBColumn ADD COLUMN ColumnNumber SMALLINT"
        moMACROcon.Execute sSQL
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sSQL = "ALTER Table NewDBColumn ADD ChangeType VARCHAR(15)"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table NewDBColumn ADD ColumnNumber SMALLINT"
        moMACROcon.Execute sSQL
    Case MACRODatabaseType.Oracle80
        sSQL = "ALTER Table NewDBColumn ADD ChangeType VARCHAR2(15)"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table NewDBColumn ADD ColumnNumber NUMBER(6)"
        moMACROcon.Execute sSQL
    End Select
    
    'As a one off set all the existing entries in table NewDBColumn to have
    'their new field "ChangeType" set to "NEWCOLUMN"
    sSQL = "UPDATE NewDBColumn SET ChangeType = 'NEWCOLUMN'"
    moMACROcon.Execute sSQL
    
    ' *** Add new columns to NewDBColumn ***
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'CRFPage','eFormDatePrompt',null,'#NULL#','NEWCOLUMN',null)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'NewDBColumn','ChangeType',1,'#NULL#','NEWCOLUMN',null)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'NewDBColumn','ColumnNumber',2,'#NULL#','NEWCOLUMN',null)"
    moMACROcon.Execute sSQL
    
    ' *** Drop Columns ***
    sSQL = "ALTER Table CRFElement DROP Column DefaultValue"
    moMACROcon.Execute sSQL
    
    'Drop the Default value constraint before dropping the column in SQL Server databases
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("CRFPageInstance", "SDVStatus")
    End If
    sSQL = "ALTER Table CRFPageInstance DROP Column SDVStatus"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("CRFPageInstance", "ReviewStatus")
    End If
    sSQL = "ALTER Table CRFPageInstance DROP Column ReviewStatus"
    moMACROcon.Execute sSQL
    
    sSQL = "ALTER Table DataItem DROP Column ExternalDataSource"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("DataItem", "RequiredTrialTypeId")
    End If
    sSQL = "ALTER Table DataItem DROP Column RequiredTrialTypeId"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("DataItem", "Required")
    End If
    sSQL = "ALTER Table DataItem DROP Column Required"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("DataItemResponse", "SDVStatus")
    End If
    sSQL = "ALTER Table DataItemResponse DROP Column SDVStatus"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("DataItemResponse", "ReviewStatus")
    End If
    sSQL = "ALTER Table DataItemResponse DROP Column ReviewStatus"
    moMACROcon.Execute sSQL
    
    sSQL = "ALTER Table DataItemResponse DROP Column ReviewComment"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("DataItemResponseHistory", "SDVStatus")
    End If
    sSQL = "ALTER Table DataItemResponseHistory DROP Column SDVStatus"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("DataItemResponseHistory", "ReviewStatus")
    End If
    sSQL = "ALTER Table DataItemResponseHistory DROP Column ReviewStatus"
    moMACROcon.Execute sSQL
    
    sSQL = "ALTER Table DataItemResponseHistory DROP Column ReviewComment"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("MACROControl", "DateReference")
    End If
    sSQL = "ALTER Table MACROControl DROP Column DateReference"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("StudyReport", "HideBlankRows")
    End If
    sSQL = "ALTER Table StudyReport DROP Column HideBlankRows"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("StudyReportData", "HideBlankRows")
    End If
    sSQL = "ALTER Table StudyReportData DROP Column HideBlankRows"
    moMACROcon.Execute sSQL
    
    sSQL = "ALTER Table TrialOffice DROP Column TransferSchedule"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("TrialSubject", "SDVStatus")
    End If
    sSQL = "ALTER Table TrialSubject DROP Column SDVStatus"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("TrialSubject", "ReviewStatus")
    End If
    sSQL = "ALTER Table TrialSubject DROP Column ReviewStatus"
    moMACROcon.Execute sSQL
    
    sSQL = "ALTER Table TrialSubject DROP Column Checksum"
    moMACROcon.Execute sSQL
    
    sSQL = "ALTER Table TrialSubject DROP Column Initials"
    moMACROcon.Execute sSQL
    
    sSQL = "ALTER Table TrialSubject DROP Column Surname"
    moMACROcon.Execute sSQL
    
    sSQL = "ALTER Table TrialSubject DROP Column Forename"
    moMACROcon.Execute sSQL
    
    sSQL = "ALTER Table TrialSubject DROP Column SubjectIdentificationCode"
    moMACROcon.Execute sSQL
    
    sSQL = "ALTER Table UnitConversionFactors DROP Column ConversionExpression"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("ValueData", "DefaultCat")
    End If
    sSQL = "ALTER Table ValueData DROP Column DefaultCat"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("VisitInstance", "SDVStatus")
    End If
    sSQL = "ALTER Table VisitInstance DROP Column SDVStatus"
    moMACROcon.Execute sSQL
    
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("VisitInstance", "ReviewStatus")
    End If
    sSQL = "ALTER Table VisitInstance DROP Column ReviewStatus"
    moMACROcon.Execute sSQL
    
    ' *** Add dropped columns to NewDBColumn ***
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'CRFElement','DefaultValue',null,'#NULL#','DROPCOLUMN',28)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'CRFPageInstance','SDVStatus',1,'#NULL#','DROPCOLUMN',15)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'CRFPageInstance','ReviewStatus',2,'#NULL#','DROPCOLUMN',14)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItem','ExternalDataSource',1,'#NULL#','DROPCOLUMN',19)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItem','RequiredTrialTypeId',2,'#NULL#','DROPCOLUMN',16)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItem','Required',3,'#NULL#','DROPCOLUMN',15)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItemResponse','SDVStatus',1,'#NULL#','DROPCOLUMN',25)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItemResponse','ReviewStatus',2,'#NULL#','DROPCOLUMN',24)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItemResponse','ReviewComment',3,'#NULL#','DROPCOLUMN',18)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItemResponseHistory','SDVStatus',1,'#NULL#','DROPCOLUMN',25)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItemResponseHistory','ReviewStatus',2,'#NULL#','DROPCOLUMN',24)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItemResponseHistory','ReviewComment',3,'#NULL#','DROPCOLUMN',18)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'MACROControl','DateReference',null,'#NULL#','DROPCOLUMN',2)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'StudyReport','HideBlankRows',null,'#NULL#','DROPCOLUMN',4)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'StudyReportData','HideBlankRows',null,'#NULL#','DROPCOLUMN',4)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialOffice','TransferSchedule',null,'#NULL#','DROPCOLUMN',7)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialSubject','SDVStatus',1,'#NULL#','DROPCOLUMN',18)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialSubject','ReviewStatus',2,'#NULL#','DROPCOLUMN',17)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialSubject','Checksum',3,'#NULL#','DROPCOLUMN',15)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialSubject','Initials',4,'#NULL#','DROPCOLUMN',7)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialSubject','Surname',5,'#NULL#','DROPCOLUMN',6)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialSubject','Forename',6,'#NULL#','DROPCOLUMN',5)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialSubject','SubjectIdentificationCode',7,'#NULL#','DROPCOLUMN',4)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'UnitConversionFactors','ConversionExpression',null,'#NULL#','DROPCOLUMN',5)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'ValueData','DefaultCat',null,'#NULL#','DROPCOLUMN',9)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'VisitInstance','SDVStatus',1,'#NULL#','DROPCOLUMN',12)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'VisitInstance','ReviewStatus',2,'#NULL#','DROPCOLUMN',11)"
    moMACROcon.Execute sSQL

    ' *** Columns with Name and/or Type Changes ***
    'Table CRFElement, field Local (Long) changed to LocalFlag (integer)
    'Table CRFElement, field Hidden (Long) changed to Hidden (integer)
    'Table CRFElement, field Mandatory (Long) changed to Mandatory (integer)
    'Table CRFElement, field Optional (Long) changed to Optional (integer)
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, meDatabaseType, moMACROcon, "CRFElement", "2.2.1")
    
    'Table MIMessage, field MIMessageUserCode (Text(50)) changed to MIMessageUserName (Text(20))
    'Table MIMessage, field MIMessageUserName (Text(255)) changed to MIMessageUserNameFull (Text(100))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, meDatabaseType, moMACROcon, "MIMessage", "2.2.1")
    
    'Table DataItemResponse, field UserId (Text(15)) changed to UserName (Text(20))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, meDatabaseType, moMACROcon, "DataItemResponse", "2.2.1")
    
    'Table DataItemResponseHistory, field UserId (Text(15)) changed to UserName (Text(20))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, meDatabaseType, moMACROcon, "DataItemResponseHistory", "2.2.1")
    
    'Table LogDetails, field UserId (Text(15)) changed to UserName (Text(20))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, meDatabaseType, moMACROcon, "LogDetails", "2.2.1")
    
    'Table Message, field UserId (Text(15)) changed to UserName (Text(20))
    'Table Message, field MessageId is no longer an autonumber
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, meDatabaseType, moMACROcon, "Message", "2.2.1")
    
    'Table SiteUser, field UserCode (Text(15)) changed to UserName (Text(20))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, meDatabaseType, moMACROcon, "SiteUser", "2.2.1")
    
    'Table StudyDefinition, field UserId (Text(15)) changed to UserName (Text(20))
    'Table StudyDefinition, field RRUserName (Text(50) changed to RRUserName (Text(20)
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, meDatabaseType, moMACROcon, "StudyDefinition", "2.2.1")
    
    'Table TrialStatusHistory, field UserId (Text(15)) changed to UserName (Text(20))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, meDatabaseType, moMACROcon, "TrialStatusHistory", "2.2.1")
    
    'Table TrialOffice, field Site (Text(50)) changed to Site (Text(8))
    'Table TrialOffice, field Password (Text(50)) changed to UserPassword (Text(50))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, meDatabaseType, moMACROcon, "TrialOffice", "2.2.1")
    
    'Upgrade MacroVersion to [2.2]
    sSQL = "UPDATE MACROControl Set MacroVersion = '2.2'"
    moMACROcon.Execute sSQL
    
    'Upgrade BuildSubVersion to [1]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '1'"
    moMACROcon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData2_1to2_2_1", "modUpgradeDatabases.bas")
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
Private Sub UpGradeData2_2from1to5()
'---------------------------------------------------------------------
' RJCW 26/11/01 Upgrade to database structure to include the new
' keyword table
'---------------------------------------------------------------------

Dim sSQL As String
'Dim oTableDef As TableDef
Dim oMACRODatabase As Database

On Error GoTo ErrHandler

'Create the new Keyword table
    
    
' *** Add New Tables to the Database ***
        'TA 6/12/2001: Pass through connection object rather than string so that moMACROcon is up to date
        Call CreateDB(eMACRODBFunction.Data, meDatabaseType, moMACROcon, False, False, "Keyword", "2.2.5")
        
'Insert data into new Keyword table
    
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ABSOLUTE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ACCESS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ACTION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ADA')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ADD')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ALL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ALLOCATE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ALTER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AND')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ANY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ARE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ASC')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ASSERTION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AUDIT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AUTHORIZATION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AUTONUMBER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AVG')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BEGIN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BETWEEN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BIT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BIT_LENGTH')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BOTH')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BYTE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CASCADE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CASCADED')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CASE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CAST')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CATALOG')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CATEGORY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CHAR')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CHAR_LENGTH')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CHARACTER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CHARACTER_LENGTH')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CHECK')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CLINICALTRIALID')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CLOSE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CLUSTER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COALESCE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COLLATE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COLLATION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COLUMN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COMMENT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COMMIT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COMPRESS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CONNECT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CONNECTION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CONSTRAINT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CONSTRAINTS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CONTINUE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CONVERT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CORRESPONDING')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COUNT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CREATE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CRFPAGECYCLENUMBER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CRFPAGEID')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CROSS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CURRENCY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CURRENT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CURRENT_DATE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CURRENT_TIME')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CURRENT_TIMESTAMP')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CURRENT_USER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CURSOR')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CYCLE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DAY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DEALLOCATE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DEC')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DECIMAL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DECLARE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DEFAULT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DEFERRABLE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DEFERRED')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DELETE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DESC')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DESCRIBE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DESCRIPTOR')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DIAGNOSTICS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DISCONNECT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DISTINCT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DOMAIN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DOUBLE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DROP')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DYNAMIC')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ELSE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('END')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('END-EXEC')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ESCAPE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXCEPT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXCEPTION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXCLUSIVE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXEC')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXECUTE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXISTS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXTERNAL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXTRACT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FALSE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FETCH')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FILE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FIRST')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FLOAT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FOR')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FOREIGN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FORM')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FORTRAN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FOUND')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FROM')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FULL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('GET')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('GLOBAL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('GO')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('GOTO')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('GRANT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('GROUP')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('HAVING')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('HOUR')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('HYPERLINK')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('IDENTIFIED')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('IDENTITY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('IMMEDIATE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('IN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INCLUDE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INCREMENT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INDEX')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INDICATOR')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INITIAL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INITIALIZATION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INITIALLY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INNER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INPUT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INSENSITIVE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INSERT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INTEGER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INTERSECT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INTERVAL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INTO')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('IS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ISOLATION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('JOIN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('KEY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LANGUAGE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LAST')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LEADING')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LEFT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LEVEL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LIKE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LOCAL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LOCK')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LONG')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LOWER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MATCH')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MAX')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MAXEXTENTS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MEMO')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('META_PREDICATE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MIN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MINUS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MINUTE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MOD')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MODE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MODIFY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MODULE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MONTH')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MULTIFILE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NAME')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NAMES')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NATIONAL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NATURAL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NCHAR')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NEXT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NO')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NOAUDIT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NOCOMPRESS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NONE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NOSPY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NOT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NOWAIT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NULL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NULLIF')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NUMBER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NUMERIC')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OCTET_LENGTH')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OF')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OFFLINE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ON')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ONE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ONLINE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ONLY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OPEN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OPTION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OR')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ORDER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OUTER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OUTPUT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OVERLAPS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PAD')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PARTIAL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PASCAL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PASSWORD')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PCTFREE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PERSONID')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('POSITION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PRECISION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PREPARE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PRESERVE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PRIMARY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PRIOR')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PRIVILEGES')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PROCEDURE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PUBLIC')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('QUESTION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RANDOMISATION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RAW')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('READ')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('REAL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('REFERENCES')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('REGISTRATION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RELATIVE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RENAME')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RESOURCE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RESTRICT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('REVOKE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RIGHT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ROLLBACK')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ROW')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ROWID')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ROWLABEL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ROWNUM')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ROWS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SCHEMA')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SCROLL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SECOND')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SECTION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SELECT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SESSION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SESSION_USER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SET')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SHARE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SITE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SIZE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SMALLINT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SOME')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SPACE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SPY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SQL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SQLCA')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SQLCODE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SQLERROR')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SQLSTATE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SQLWARNING')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('START')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('STATUS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SUBSTRING')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SUCCESSFUL')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SUM')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SYNONYM')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SYSDATE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SYSTEM_USER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TABLE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TEMP')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TEMPORARY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TEXT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('THEN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TIME')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TIMESTAMP')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TIMEZONE_HOUR')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TIMEZONE_MINUTE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TINYINT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TO')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TRAILING')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TRANSACTION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TRANSLATE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TRANSLATION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TRIGGER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TRIM')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TRUE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('UID')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('UNION')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('UNIQUE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('UNKNOWN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('UPDATE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('UPPER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('USAGE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('USER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('USING')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VALIDATE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VALUE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VALUES')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VARCHAR')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VARCHAR2')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VARYING')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VIEW')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VISIT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VISITCYCLENUMBER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VISITID')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VOLATILE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WHEN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WHENEVER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WHERE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WITH')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WORK')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WRITE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('YEAR')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('YES')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('YESNO')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ZONE')"
        moMACROcon.Execute sSQL
        'REM 10/05/02 - added new keywords
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ABS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AFTER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AVERAGE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BEFORE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATE_AND_TIME')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATE_DIFF')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATEDIFF')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATENOW')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATEOF')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATEPART')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATETIME')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DAYS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DURING')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EVERY')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FORMAT_DATE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('HOURS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('IF')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INCLUDES')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('IS_KNOWN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ISAFTER')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ISBEFORE')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ISDURING')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ISKNOWN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LEN')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MINUTES')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MONTHS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NETSUPPORT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NOT_ONEOF')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NOW')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ONEOF')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RESULT_OF')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SECONDS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SQRT')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TIME_DIFF')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TIMENOW')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WEEK')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WEEKS')"
        moMACROcon.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('YEARS')"
        moMACROcon.Execute sSQL

        
        'insert table information into MACROTABLE for Keyword
        sSQL = "INSERT INTO MACROTable (TableName, SegmentId, STYDEF, PATRSP, LABDEF) "
        sSQL = sSQL & " VALUES ('Keyword', '', 0, 0, 0)"
        moMACROcon.Execute sSQL
    
    Exit Sub

ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData2_2from1to5", "modUpgradeDatabases.bas")
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
Private Sub UpGradeData2_2from5to10()
'---------------------------------------------------------------------
' REM 15/04/02 Upgrade to MACRO 2.2 database structure to include the new
' ArezzoToken and AutoImport tables and a new column called HadValue to the
' DataItemResponse and DataItemResponseHistory tables and sql to insert the correct
' HadValue into the column
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    ' *** Add New Tables to the Database ***
    Call CreateDB(eMACRODBFunction.Data, meDatabaseType, moMACROcon, False, False, "ArezzoToken", "2.2.10")

    Call CreateDB(eMACRODBFunction.Data, meDatabaseType, moMACROcon, False, False, "AutoImportControl", "2.2.10")

    ' *** Add New Tables to the table MACROTable ***
    sSQL = "INSERT INTO MACROTable VALUES ('ArezzoToken','',0,0,0)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('AutoImportControl','',0,0,0)"
    moMACROcon.Execute sSQL

    ' *** Add New Columns ***
    'Added 1 column to DataItemResponse: HadValue
    'Added 1 column to DataItemResponseHistory: HadValue
    Select Case meDatabaseType
    Case MACRODatabaseType.Access
        sSQL = "ALTER Table DataItemResponse ADD COLUMN HadValue SMALLINT"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD COLUMN HadValue SMALLINT"
        moMACROcon.Execute sSQL
    
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sSQL = "ALTER Table DataItemResponse ADD HadValue SMALLINT"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD HadValue SMALLINT"
        moMACROcon.Execute sSQL
        
    Case MACRODatabaseType.Oracle80
        sSQL = "ALTER Table DataItemResponse ADD HadValue NUMBER(6)"
        moMACROcon.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD HadValue NUMBER(6)"
        moMACROcon.Execute sSQL
        
    End Select

    ' *** Drop Columns ***
    'First drop the default constraints on the columns for SQL server
    If meDatabaseType = MACRODatabaseType.sqlserver Or meDatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("ClinicalTrial", "ActualRecruitment")
        Call DropDefaultConstraint("Trialsite", "TrialSiteActualRecruitment")
    End If
    
    'Drop ActualRecruitment column from ClinicalTrail table
    'Drop TrialSiteActualRecruitment column from TrialSite table
    sSQL = "ALTER Table ClinicalTrial DROP Column ActualRecruitment"
    moMACROcon.Execute sSQL
    sSQL = "ALTER Table TrialSite DROP Column TrialSiteActualRecruitment"
    moMACROcon.Execute sSQL

    ' *** Add values to the new HadValue columns ***
    'Update DataItemResponseHistory table
    Select Case meDatabaseType
    Case MACRODatabaseType.Access
    sSQL = "Update DataItemResponseHistory a" _
        & " Set HadValue = 0" _
        & " WHERE   (a.ResponseValue IS NULL AND NOT EXISTS (SELECT  * FROM DataItemResponseHistory b" _
        & "                                                  WHERE b.ResponseValue IS NOT Null" _
        & "                                                  AND b.ClinicalTrialId = a.ClinicalTrialId" _
        & "                                                  AND b.TrialSite = a.TrialSite" _
        & "                                                  AND b.PersonId = a.PersonId" _
        & "                                                  AND b.ResponseTaskId = a.ResponseTaskId" _
        & "                                                  AND b.ResponseTimeStamp < a.ResponseTimeStamp))"
        moMACROcon.Execute sSQL
    
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
    sSQL = "Update DataItemResponseHistory" _
        & " Set HadValue = 0" _
        & " FROM    DataItemResponseHistory a" _
        & " WHERE   (a.ResponseValue IS NULL AND NOT EXISTS (SELECT  * FROM DataItemResponseHistory b" _
        & "                                                  WHERE b.ResponseValue IS NOT Null" _
        & "                                                  AND b.ClinicalTrialId = a.ClinicalTrialId" _
        & "                                                  AND b.TrialSite = a.TrialSite" _
        & "                                                  AND b.PersonId = a.PersonId" _
        & "                                                  AND b.ResponseTaskId = a.ResponseTaskId" _
        & "                                                  AND b.ResponseTimeStamp < a.ResponseTimeStamp))"
        moMACROcon.Execute sSQL
        
    Case MACRODatabaseType.Oracle80
    sSQL = "Update DataItemResponseHistory a" _
        & " Set HadValue = 0" _
        & " WHERE   (a.ResponseValue IS NULL AND NOT EXISTS (SELECT  * FROM DataItemResponseHistory b" _
        & "                                                  WHERE b.ResponseValue IS NOT Null" _
        & "                                                  AND b.ClinicalTrialId = a.ClinicalTrialId" _
        & "                                                  AND b.TrialSite = a.TrialSite" _
        & "                                                  AND b.PersonId = a.PersonId" _
        & "                                                  AND b.ResponseTaskId = a.ResponseTaskId" _
        & "                                                  AND b.ResponseTimeStamp < a.ResponseTimeStamp))"
        moMACROcon.Execute sSQL
        
    End Select

    sSQL = "Update DataItemResponseHistory" _
        & " SET HadValue = 1 WHERE HadValue IS NULL"
        moMACROcon.Execute sSQL


    'Update DataItemResponse table by copying value from DataItemResponseHistory table
    Select Case meDatabaseType
    Case MACRODatabaseType.Access
    sSQL = "Update DataItemResponse, DataItemResponseHistory" _
        & " Set DataItemResponse.HadValue = DataItemResponseHistory.HadValue" _
        & " WHERE DataItemResponse.ClinicalTrialId = DataItemResponseHistory.ClinicalTrialId" _
        & " AND DataItemResponse.TrialSite = DataItemResponseHistory.TrialSite" _
        & " AND DataItemResponse.PersonId = DataItemResponseHistory.PersonId" _
        & " AND DataItemResponse.ResponseTaskId = DataItemResponseHistory.ResponseTaskId" _
        & " AND DataItemResponse.ResponseTimeStamp = DataItemResponseHistory.ResponseTimeStamp"
        moMACROcon.Execute sSQL

    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
    sSQL = "Update DataItemResponse" _
        & " Set DataItemResponse.HadValue = DataItemResponseHistory.HadValue" _
        & " FROM DataItemResponse, DataItemResponseHistory" _
        & " WHERE DataItemResponse.ClinicalTrialId = DataItemResponseHistory.ClinicalTrialId" _
        & " AND DataItemResponse.TrialSite = DataItemResponseHistory.TrialSite" _
        & " AND DataItemResponse.PersonId = DataItemResponseHistory.PersonId" _
        & " AND DataItemResponse.ResponseTaskId = DataItemResponseHistory.ResponseTaskId" _
        & " AND DataItemResponse.ResponseTimeStamp = DataItemResponseHistory.ResponseTimeStamp"
        moMACROcon.Execute sSQL

    Case MACRODatabaseType.Oracle80
     sSQL = "UPDATE DataItemResponse SET DataItemResponse.Hadvalue = (Select DataItemResponseHistory.HadValue" _
        & " From DataItemResponseHistory" _
        & " Where DataItemResponse.ClinicalTrialId = DataItemResponseHistory.ClinicalTrialId" _
        & " AND DataItemResponse.TrialSite = DataItemResponseHistory.TrialSite" _
        & " AND DataItemResponse.PersonId = DataItemResponseHistory.PersonId" _
        & " AND DataItemResponse.ResponseTaskId = DataItemResponseHistory.ResponseTaskId" _
        & " AND DataItemResponse.ResponseTimeStamp = DataItemResponseHistory.ResponseTimeStamp)"
        moMACROcon.Execute sSQL
    
    End Select
    
    ' *** Add the new column names to the NewDBColumn table ***
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,10,'DataItemResponse','HadValue',null,'1','NEWCOLUMN',null)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,10,'DataItemResponseHistory','HadValue',null,'1','NEWCOLUMN',null)"
    moMACROcon.Execute sSQL
    
    ' *** Add the dropped columns to the NewDBColumn table ***
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,10,'ClinicalTrial','ActualRecruitment',null,'#NULL#','DROPCOLUMN',8)"
    moMACROcon.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,10,'TrialSite','TrialSiteActualRecruitment',null,'#NULL#','DROPCOLUMN',3)"
    moMACROcon.Execute sSQL
    
    '*** Add new Index ***
    sSQL = "CREATE INDEX IDX_VI_SECONDARYKEY "
    sSQL = sSQL & "ON VisitInstance ( ClinicalTrialId, TrialSite, PersonId, VisitId, VisitCycleNumber )"
    moMACROcon.Execute sSQL
    
    'Update BuildSubVersion to [10]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '10'"
    moMACROcon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData2_2from5to10", "modUpgradeDatabases.bas")
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
Private Sub UpGradeData2_2from14to15()
'---------------------------------------------------------------------
' REM 10/06/02
' Insert Default values into the AutoImportControl table
'---------------------------------------------------------------------
Dim sSQL As String

    'if there is already a row in the table then
    On Error Resume Next
    sSQL = "INSERT INTO AutoImportControl VALUES (1,'START',60)"
    moMACROcon.Execute sSQL

End Sub

'---------------------------------------------------------------------
Private Sub UpgradeData2_2from18to19()
'---------------------------------------------------------------------
' MLM 28/06/02: Added. Add indexes to TrialSubject to improve performance.
'---------------------------------------------------------------------

Dim sSQL As String

    sSQL = "CREATE INDEX IDX_TS_TRIALSITE "
    sSQL = sSQL & "ON TrialSubject ( TrialSite )"
    moMACROcon.Execute sSQL

    sSQL = "CREATE INDEX IDX_TS_PERSONID "
    sSQL = sSQL & "ON TrialSubject ( PersonId )"
    moMACROcon.Execute sSQL


End Sub

'---------------------------------------------------------------------
Private Sub DropDefaultConstraint(ByVal sTableName As String, _
                                    ByVal sColumnName As String)
'---------------------------------------------------------------------
'The SQL within this sub was created by Richard Weare
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sDefaultConstraintName As String

    sSQL = "SELECT name FROM sysobjects " _
        & "WHERE id =( SELECT constid " _
                    & "FROM sysconstraints " _
                    & "WHERE constid in (  SELECT id " _
                                        & "FROM sysobjects " _
                                        & "WHERE xtype = 'D' " _
                                        & "AND parent_obj = (  SELECT id " _
                                                            & "FROM sysobjects " _
                                                            & "WHERE name = '" & sTableName & "')) " _
        & "AND colid in (  SELECT colid " _
                        & "FROM syscolumns " _
                        & "WHERE name = '" & sColumnName & "' " _
                        & "AND id = (  SELECT id " _
                                    & "FROM sysobjects " _
                                    & "WHERE name = '" & sTableName & " ' " _
                                    & "AND type = 'U')) " _
                    & ")"
                    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, moMACROcon, adOpenKeyset, adLockReadOnly, adCmdText
    
    'If a default constraint exists the record count should be 1
    If rsTemp.RecordCount = 1 Then
        sDefaultConstraintName = rsTemp!Name
        sSQL = "ALTER Table " & sTableName & " DROP Constraint " & sDefaultConstraintName
        moMACROcon.Execute sSQL
    End If

End Sub

'---------------------------------------------------------------------
Private Sub FieldNameTypeUpgrade(ByVal nDBFunction As eMACRODBFunction, _
                                ByVal nDBType As MACRODatabaseType, _
                                ByVal oConnection As ADODB.Connection, _
                                ByVal sTable As String, _
                                ByVal sVersion As String)
'---------------------------------------------------------------------
'Utility used to upgrade a table with one or more fields that have had
'a Name and/or Type change.
'
'sTable contains the table to be upgraded.
'sVersion contains the version that the table is being updated to.
'
'The upgrade happens in 4 stages:-
'   Read the contents of the specified table out into a recordset.
'   Drop the old version of the table.
'   Create a new version of the table using CreateDB.
'   Read the contents of the recordset back into the newly created version of the table.
'
'Note that if a table has had new fields added, or old fields removed, at the same time
'as other fields have had name/type changes then the adding and/or dropping should take place
'prior to this sub being called. (i.e. the old and new versions of a table need to have the
'same number of fields for this sub to work)
'
'THIS UTILITY WILL HANDLE TYPE CHANGES BETWEEN INTEGER AND LONG
'THIS UTILITY WILL HANDLE TYPE CHANGES BETWEEN AND STRINGS OF DIFFERENT LENGTHS
'THIS UTILITY WILL NOT HANDLE TYPE CHANGES THAT REQUIRE A FUNDAMENTAL TYPE CHANGE (I.E. INTEGER TO STRING)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsFromTable As ADODB.Recordset
Dim rsToTable As ADODB.Recordset
Dim j As Long
Dim i As Integer
Dim lNumberOfRecords As Long
Dim nNumberOfFields As Integer

    On Error GoTo ErrHandler

    'Place the contents of the table into recordset rsFromTable
    sSQL = "SELECT * FROM " & sTable
    Set rsFromTable = New ADODB.Recordset
    rsFromTable.CursorLocation = adUseClient
    rsFromTable.Open sSQL, oConnection, adOpenKeyset, adLockReadOnly, adCmdText
    rsFromTable.ActiveConnection = Nothing
    
    'Drop the original version of the table
    sSQL = "DROP Table " & sTable
    oConnection.Execute sSQL
    
    'Create a new version of the table
    'Mo Morris 1/10/01, optional field to prevent CreateDB displaying messages added
    Call CreateDB(nDBFunction, nDBType, moMACROcon, False, False, sTable, sVersion)
    
    'Prepare recordset rsToTable to receive the contents of rsFromTable
    sSQL = "SELECT * FROM " & sTable
    Set rsToTable = New ADODB.Recordset
    rsToTable.Open sSQL, oConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    'copy the contents of rsTemp into the newly created table
    lNumberOfRecords = rsFromTable.RecordCount
    nNumberOfFields = rsFromTable.Fields.Count
    For j = 1 To lNumberOfRecords
        rsToTable.AddNew
        For i = 0 To nNumberOfFields - 1
            rsToTable.Fields(i).Value = rsFromTable.Fields(i).Value
        Next i
        rsToTable.Update
        rsFromTable.MoveNext
    Next j
    
    rsToTable.Close
    Set rsToTable = Nothing
    rsFromTable.Close
    Set rsFromTable = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "FieldNameTypeUpgrade", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

