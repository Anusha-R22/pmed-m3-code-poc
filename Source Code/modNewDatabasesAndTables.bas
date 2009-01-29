Attribute VB_Name = "modNewDatabasesAndTables"
'----------------------------------------------------------------------------------------'
'   Copyright:  Inferfrmformd Ltd. 2000. All Rights Reserved
'   File:       modNewDatabasesAndTables.bas
'   Author:     Mo Morris, August 2001
'   Purpose:    Utilities for creating a new database
'               With facilities for creating a single table for a specified version
'
'----------------------------------------------------------------------------------------'
'   Revisions:
' MLM 6/12/01: Changed CreateEmptyAccessDb to allow overwrite of existing files
' TA 8/1/02: VTRACK code options put in
' TA 18/01/2002: Special Insert 4 is used to put our support site password in the database
'               so that the script file does not contain the support site password. DCBB2.2.7.7
'TA 09/10/2002: CreateDB now has new error handling
'TA 03/11/2002: Use scripts to create new MACRO db
' NCJ 15 Jan 04 - Corrected "SecuirtyDB" to "SecurityDB" in WriteDBScriptToFile
'----------------------------------------------------------------------------------------'
Option Explicit

'Mo Morris 4/7/01
' RS 08JAN2002:  Added VTRACK constant
Public Enum eMACRODBFunction
    Data = 0
    Security = 1
    VTRACK = 2
End Enum

Private Const msMACRO_DATA_DB_DEFINITION As String = "MacroDataDb.txt"
Private Const msMACRO_SECURITY_DB_DEFINITION As String = "MacroSecDb.txt"
Private Const msMACRO_DATA_TYPES_DEFINITION As String = "MacroDataTypes.txt"
Private Const msMACRO_VTRACK_DB_DEFINITION As String = "VtrackTables.txt"

Private mColDataTypeLookUp As Collection

'--------------------------------------------------------------------
Public Sub CreateDB(ByVal nDBFunction As eMACRODBFunction, _
                    ByVal nDBType As MACRODatabaseType, _
                    ByVal sConnectionOrPath As Variant, _
                    ByVal bWriteDBScriptToFile As Boolean, _
                    Optional bDisplayMessages As Boolean = True, _
                    Optional sRequiredTable As String, _
                    Optional sVersion As String)
'--------------------------------------------------------------------
'Note that this sub runs under 3 modes:-
'
'The first will create all of MACRO's tables in the database pointed
'to by the connection string sConnectionOrPath.
'[bWriteDBScriptToFile=False, sRequiredTable=Blank, sVersion=Blank]
'
'The second will write the table creating SQL statements as a script
'into the file pointed at by sConnectionOrPath.
'[bWriteDBScriptToFile=True, sRequiredTable=Blank, sVersion=Blank]
'
'The third uses the optional paramaters sRequiredTable/sVersion, and
'will create the single table specified by sRequiredTable/sVersion in the
'database pointed to by the connection string sConnectionOrPath.
'[bWriteDBScriptToFile=False, sRequiredTable=RequiredTable, sVersion=RequiredVersion]
'Not that inserts will not be performed when this sub is called in SingleTableMode.
'
'--------------------------------------------------------------------
'Macro's upgrade database programs needed a facility for creating a single
'table as it would have been at a particular version.
'To allow this to happen the database definition files (MacroDataDb.txt &
'MacroSecurityDb.txt will now contain different versions of a table definition.
'
'The most recent version of a table definition will start with
'   NAME|TableName
'Older versions of a table definition will start with
'   NAME|TableName|Major.Minor.Sub
'where Major.Minor.Sub represents the version number at the time the database
'changed.
'
'For example (if the table CRFElement changed in versions 2.2.5, 2.2.12 and 3.0.2
'then the following table name tags would exist
'   NAME|CRFElement|2.2.5   (followed by the table definition used upto version 2.2.5)
'   NAME|CRFElement|2.2.12  (followed by the table definition used between 2.2.5 & 2.2.12)
'   NAME|CRFElement|3.0.2   (followed by the table definition used between 2.2.12 & 3.0.2)
'   NAME|CRFElement         (Followed by the latest table definition used from version 3.0.2
'                            onwards when creating new databases)
'Note:- The order of the different versions is important
'--------------------------------------------------------------------
'TA 18/01/2002: Special Insert 4 is used to put our support site password in the database
'               so that the script file does not contain the support site password.
'TA 09/10/2002: Now has new error handling
'--------------------------------------------------------------------
Dim nDbDefinitionFileNumber As Integer
Dim sRecord As String
Dim asRecordContents() As String
Dim sTag As String
Dim sTableTagVersion As String
Dim sTableName As String
Dim sFieldName As String
Dim sFieldType As String
Dim sUCONName As String
Dim sUCONFields As String
Dim sFKName As String
Dim sFKFields As String
Dim sFKRefTable As String
Dim sFKRefFields As String
Dim sPKName As String
Dim sPKFields As String
Dim sUIndexName As String
Dim sUIndexFields As String
Dim sIndexName As String
Dim sIndexFields As String
Dim sInsertFields As String
Dim sInsertValues As String
Dim sSQL As String
Dim oCreateDatabaseConnection As ADODB.Connection
Dim nDbScriptFileNumber As Integer
Dim bErrorsOccured As Boolean
Dim sDBFunctionText As String
Dim sDBTypeText As String
Dim sToScriptText As String
Dim sMessageLine As String

Dim bWholeDatabaseMode As Boolean
Dim bTableRequired As Boolean

Dim bSingleTableMode As Boolean
Dim bSingleTableReached As Boolean

Dim rsOnlineSupport As ADODB.Recordset ' recordset to update online support password
    On Error GoTo ErrHandler

    HourglassOn
    bErrorsOccured = False
    
    'Assess whether sub is being called in SingleTableMode or WholeDatabaseMode
    If RTrim(sRequiredTable & sVersion) = "" Then
        bWholeDatabaseMode = True
        bSingleTableMode = False
    Else
        bSingleTableMode = True
        bWholeDatabaseMode = False
    End If
    
    bSingleTableReached = False
    bTableRequired = False

    nDbDefinitionFileNumber = FreeFile
    'Open the relevant database definition file (data or security)
    Select Case nDBFunction
    Case eMACRODBFunction.Data
        Open gsAppPath & "Database Scripts\MACRO22Upgrade\" & msMACRO_DATA_DB_DEFINITION For Input As #nDbDefinitionFileNumber
        sDBFunctionText = "Data"
    Case eMACRODBFunction.Security
        Open gsAppPath & "Database Scripts\MACRO22Upgrade\" & msMACRO_SECURITY_DB_DEFINITION For Input As #nDbDefinitionFileNumber
        sDBFunctionText = "Security"
    Case eMACRODBFunction.VTRACK
        Open gsAppPath & "Database Scripts\MACRO22Upgrade\" & msMACRO_VTRACK_DB_DEFINITION For Input As #nDbDefinitionFileNumber
        sDBFunctionText = "Vtrack"
    End Select
    
    'Give sDBTypeText a value (for use in messages)
    Select Case nDBType
    Case MACRODatabaseType.Access
        sDBTypeText = "Access"
    Case MACRODatabaseType.sqlserver
        sDBTypeText = "SQLServer"
    Case MACRODatabaseType.SQLServer70
        sDBTypeText = "SQLServer"
    Case MACRODatabaseType.Oracle80
        sDBTypeText = "Oracle"
    End Select
    
    'Give sToScriptText a value (for use in messages)
    If bWriteDBScriptToFile Then
        sToScriptText = " script"
    Else
        sToScriptText = ""
    End If
    
    'Load the contents of the Macro Data Types file into a structure for retieving
    'the specific database data type
    Call CreateMacroDataTypesLookUp(nDBType)
    
    If bWriteDBScriptToFile Then
        'Open the file into which the DB creating script will be written
        nDbScriptFileNumber = FreeFile
        Open sConnectionOrPath For Output As #nDbScriptFileNumber
    Else
        If Not IsObject(sConnectionOrPath) Then
            'set up an ADO connection to the about to be created database
            Set oCreateDatabaseConnection = New ADODB.Connection
            oCreateDatabaseConnection.Open sConnectionOrPath
        Else
            Set oCreateDatabaseConnection = sConnectionOrPath
        End If
    End If
    
    'The following loop controls the reading of a database definition file.
    Do While Not EOF(nDbDefinitionFileNumber)
        'The file is read one line at a time.
        '"|" is the deliminator
        'Each line either:-
        '   starts with a Tag (NAME,FIELD,UCON,PK,UINDEX,INDEX,INSERT)
        '   starts with a "|" and is treated as a comment line, which is just read without any action
        '   or is a blank line, spacing between table definitions, which is just read without any action
        Line Input #nDbDefinitionFileNumber, sRecord
        asRecordContents = Split(sRecord, "|")
        If UBound(asRecordContents) > 0 Then
            sTag = asRecordContents(0)
        Else
            sTag = ""
        End If
        
        'If its a table name tag extract the table (asRecordContents(1)) and version (asRecordContents(2), if one exists.
        If sTag = "NAME" Then
            sTableName = asRecordContents(1)
            If UBound(asRecordContents) = 2 Then
                sTableTagVersion = RTrim(asRecordContents(2))
            Else
                sTableTagVersion = ""
            End If
        End If
        
        'When in WholeDatabaseMode (Create New Database), table definitions with a version should
        'be skipped, because table definitions without a version will be the most recent and hence
        'the one that should be used in a new database.
        If bWholeDatabaseMode And sTag = "NAME" Then
            If sTableTagVersion = "" Then
                bTableRequired = True
            Else
                bTableRequired = False
            End If
        End If
        
        'Check for the end of a single table having been created. This can be deemed
        'to have finished when the next "NAME" tag is reached
        If bSingleTableMode And (sTag = "NAME") And bSingleTableReached Then
            Exit Do
        End If
        
        'When in SingleTableMode the table definition's version needs to be checked for the correct version.
        'The calling version (sVersion) is compaired to the TableTagVersion until one is found that it less than.
        'When or if a table definition with a blank version is reached then that definition is the one that should be used.
        If bSingleTableMode And (sTag = "NAME") And (sTableName = sRequiredTable) Then
            If sTableTagVersion = "" Then
                bSingleTableReached = True
            Else
                'Check that the table definition is the correct version
                If CallVersionLessThanTagVersion(sVersion, sTableTagVersion) Then
                    bSingleTableReached = True
                End If
            End If
        End If
        
        'If required Process the record based on the Tag
        If (bWholeDatabaseMode And bTableRequired) Or (bSingleTableMode And bSingleTableReached) Then
            Select Case sTag
            Case "NAME"
                sSQL = "CREATE TABLE " & sTableName & " ("
            Case "FIELD"
                sFieldName = asRecordContents(1)
                sFieldType = asRecordContents(2)
                sSQL = sSQL & " " & sFieldName & " " & DbSpecificDataType(sFieldType) & ","
            Case "UCON"
                sUCONName = asRecordContents(1)
                sUCONFields = asRecordContents(2)
                'note comma at the end of the Unique Constraint statement
                sSQL = sSQL & " CONSTRAINT " & sUCONName & " UNIQUE " & sUCONFields & ","
            Case "FK"
                sFKName = asRecordContents(1)
                sFKFields = asRecordContents(2)
                sFKRefTable = asRecordContents(3)
                sFKRefFields = asRecordContents(4)
                'note comma at the end of the Foreign key statement
                sSQL = sSQL & " CONSTRAINT " & sFKName & " FOREIGN KEY " & sFKFields & " REFERENCES " & sFKRefTable & " " & sFKRefFields & ","
            Case "PK"
                sPKName = asRecordContents(1)
                sPKFields = asRecordContents(2)
                'note closing bracket at the end of the Primary key statement
                sSQL = sSQL & " CONSTRAINT " & sPKName & " PRIMARY KEY " & sPKFields & ")"
                'Create the Table
                If bWriteDBScriptToFile Then
                    Print #nDbScriptFileNumber, sSQL & vbNewLine
                Else
                    oCreateDatabaseConnection.Execute sSQL
                End If
            Case "BUILD"
                'strip off the trailing comma and add a closing bracket
                sSQL = Mid(sSQL, 1, Len(sSQL) - 1) & ")"
                'Create the table
                If bWriteDBScriptToFile Then
                    Print #nDbScriptFileNumber, sSQL & vbNewLine
                Else
                    oCreateDatabaseConnection.Execute sSQL
                End If
            Case "UINDEX"
                sUIndexName = asRecordContents(1)
                sUIndexFields = asRecordContents(2)
                sSQL = "CREATE UNIQUE INDEX " & sUIndexName & " ON " & sTableName & " " & sUIndexFields
                'Create the Unique Index
                If bWriteDBScriptToFile Then
                    Print #nDbScriptFileNumber, sSQL & vbNewLine
                Else
                    oCreateDatabaseConnection.Execute sSQL
                End If
            Case "INDEX"
                sIndexName = asRecordContents(1)
                sIndexFields = asRecordContents(2)
                sSQL = "CREATE INDEX " & sIndexName & " ON " & sTableName & " " & sIndexFields
                'Create the Index
                If bWriteDBScriptToFile Then
                    Print #nDbScriptFileNumber, sSQL & vbNewLine
                Else
                    oCreateDatabaseConnection.Execute sSQL
                End If
            Case "INSERT"
                'Inserts not performed in Single Table Mode
                If bWholeDatabaseMode Then
                    sInsertFields = asRecordContents(1)
                    sInsertValues = asRecordContents(2)
                    sSQL = "INSERT INTO " & sTableName & " " & sInsertFields & " VALUES " & sInsertValues
                    'Perform the insert
                    If bWriteDBScriptToFile Then
                        Print #nDbScriptFileNumber, sSQL & vbNewLine
                    Else
                        oCreateDatabaseConnection.Execute sSQL
                    End If
                End If
            Case "SPECIALINSERT1"
                'Inserts not performed in Single Table Mode
                If bWholeDatabaseMode Then
                    'Special insert into security database table Databases
                    sInsertFields = "(DatabaseCode,DatabaseLocation,DatabaseType,DatabaseUser,DatabasePassword)"
                    sInsertValues = "('Access','" & App.Path & "\Databases\Macro.mdb',0,'rde','macrotm')"
                    sSQL = "INSERT INTO " & sTableName & " " & sInsertFields & " VALUES " & sInsertValues
                    'Perform the insert
                    If bWriteDBScriptToFile Then
                        Print #nDbScriptFileNumber, sSQL & vbNewLine
                    Else
                        oCreateDatabaseConnection.Execute sSQL
                    End If
                End If
            Case "SPECIALINSERT2"
                'Inserts not performed in Single Table Mode
                If bWholeDatabaseMode Then
                'Special insert into security database table SecurityControl
                    sInsertFields = "(SecurityMode,MACROVersion,BuildSubVersion)"
                    sInsertValues = "(0,'" & App.Major & "." & App.Minor & "','" & App.Revision & "')"
                    sSQL = "INSERT INTO " & sTableName & " " & sInsertFields & " VALUES " & sInsertValues
                    'Perform the insert
                    If bWriteDBScriptToFile Then
                        Print #nDbScriptFileNumber, sSQL & vbNewLine
                    Else
                        oCreateDatabaseConnection.Execute sSQL
                    End If
                End If
            Case "SPECIALINSERT3"
                'Inserts not performed in Single Table Mode
                If bWholeDatabaseMode Then
                    'Special insert into data database table MACROControl
                    sInsertFields = "(MACROVersion,BuildSubVersion,IdleTimeOut)"
                    sInsertValues = "('" & App.Major & "." & App.Minor & "','" & App.Revision & "',300)"
                    sSQL = "INSERT INTO " & sTableName & " " & sInsertFields & " VALUES " & sInsertValues
                    'Perform the insert
                    If bWriteDBScriptToFile Then
                        Print #nDbScriptFileNumber, sSQL & vbNewLine
                    Else
                        oCreateDatabaseConnection.Execute sSQL
                    End If
                End If
            Case "SPECIALINSERT4"
                'TA 18/1/2002: We must do this in code so the script file does not contain a password: DCBB2.2.7.7
                'Inserts not performed in Single Table Mode
                If bWholeDatabaseMode Then
                    'Special insert into data database table MACROControl
                    sInsertFields = "(SupportUserName,SupportUserPassWord,SupportURL)"
                    sInsertValues = "('INFERMED','unspecified','www.infermed.com/support/insertproblem.asp')"
                    sSQL = "INSERT INTO " & sTableName & " " & sInsertFields & " VALUES " & sInsertValues
                    'Perform the insert
                    If bWriteDBScriptToFile Then
                        'TA 18/1/2002: We must not write this line to a file as it contains a password
                        'Print #nDbScriptFileNumber, sSQL & vbNewLine
                    Else
                        oCreateDatabaseConnection.Execute sSQL
                        'TA 18/1/2002: I am using a recordset for this to cope with the any
                        '               dodgy characters that might me in the encrypted sting
                        Set rsOnlineSupport = New ADODB.Recordset
                        rsOnlineSupport.Open "Select SupportUserPassword from OnlineSupport where SupportUserName = 'INFERMED'" _
                                    , oCreateDatabaseConnection, adOpenForwardOnly, adLockPessimistic, adCmdText
                        rsOnlineSupport.Fields(0).Value = Crypt("guido")
                        rsOnlineSupport.Update
                        rsOnlineSupport.Close
                        Set rsOnlineSupport = Nothing
                        
                    End If
                End If
            Case ""
                'Do nothing, because its a comment line or a blank line
            Case Else
                'An erroneous tag has been read
                MsgBox "An error has occurred during the creation of a database" & sToScriptText & "." _
                    & vbNewLine & "The error occurred during the execution of the following DB script line." _
                    & vbNewLine & "[" & sRecord & "]", vbInformation, "MACRO"
                bErrorsOccured = True
                Exit Do
            End Select
        End If
    Loop
    
    'Close the database definition file
    Close #nDbDefinitionFileNumber
    
    If bWriteDBScriptToFile Then
        Close #nDbScriptFileNumber
    End If
    
    'Destroy the Data Type LookUp collection
    Set mColDataTypeLookUp = Nothing
    
    HourglassOff
    
    'If in SingleTableMode or CreatDB has been called from Restore site database then
    'exit now without displaying any messages
    If Not bDisplayMessages Then
        Exit Sub
    End If
    
    'Generate the relevant glog entries and messages. (dependent on errors having occured)
    If bErrorsOccured Then
        sMessageLine = "Failed to create " & sDBFunctionText & " database" & sToScriptText & " (" & sDBTypeText & ")"
        goUser.gLog goUser.UserName, gsCREATE_DB, sMessageLine
        'Remind the user to delete the erroroneous database or DB creation script
        If bWriteDBScriptToFile Then
            MsgBox sMessageLine & vbNewLine _
                & "Delete the database creation script as soon as possible.", vbInformation, "MACRO"
        Else
            MsgBox sMessageLine & vbNewLine _
                & "Delete the erroneous database as soon as possible.", vbInformation, "MACRO"
        End If
    Else
        sMessageLine = sDBFunctionText & " database" & sToScriptText & " (" & sDBTypeText & ") created."
        goUser.gLog goUser.UserName, gsCREATE_DB, sMessageLine
        'If its a Data database then remind user for the need to register the database
        If (nDBFunction = eMACRODBFunction.Data) And (Not bWriteDBScriptToFile) Then
            MsgBox sMessageLine & vbNewLine _
                & "You must register the database before you can use it for data entry.", vbInformation, "MACRO"
        Else
            MsgBox sMessageLine, vbInformation, "MACRO"
        End If
    End If

Exit Sub
ErrHandler:
    HourglassOff
    Err.Raise Err.Number, , Err.Description & "|" & "modNewDatabasesAndTables.CreateDB"

End Sub

'--------------------------------------------------------------------
Public Function CreateEmptyAccessDb(ByVal sPath As String, _
                                    Optional bDefaultPasswordRequired As Boolean) As String
'--------------------------------------------------------------------
'Changed Mo Morris 21/9/01
'Optional parameter added to control setting of default Data database
'password of 'macrotm'.
'MLM 6/12/01: Make sure file doesn't already exist before trying to create it.
'--------------------------------------------------------------------
Dim oMacroDatabase As dao.Database
Dim sConnectionString As String

    On Error GoTo ErrHandler
    
    'if the file exists already, delete it
    'the user will have already confirmed the overwrite via the common dialog
    If FileExists(sPath) Then
        Kill sPath
    End If
    
    'Create an Access database using DAO
    Set oMacroDatabase = DBEngine.Workspaces(0).CreateDatabase(sPath, dbLangGeneral, dbEncrypt)
    
    If bDefaultPasswordRequired Then
        'set the default password of 'macrotm'
        oMacroDatabase.NewPassword "", "macrotm"
    End If
    
    oMacroDatabase.Close
    Set oMacroDatabase = Nothing
    
    'set up a connection string to the new Access database and return it to the calling Sub
    sConnectionString = Connection_String(CONNECTION_MSJET_OLEDB_40, sPath, , , "macrotm")
    CreateEmptyAccessDb = sConnectionString

Exit Function

ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CreateEmptyAccessDb", "modNewDatabasesAndTables")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'--------------------------------------------------------------------
Private Sub CreateMacroDataTypesLookUp(ByVal nDBType As MACRODatabaseType)
'--------------------------------------------------------------------
Dim nMacroDataTypesFileNumnber As Integer
Dim sRecord As String
Dim asRecordContents() As String
Dim sMacroDataType As String
Dim sDbDataType As String

    On Error GoTo ErrHandler
    
    Set mColDataTypeLookUp = New Collection
    
    nMacroDataTypesFileNumnber = FreeFile
    Open gsAppPath & "Database Scripts\MACRO22Upgrade\" & msMACRO_DATA_TYPES_DEFINITION For Input As #nMacroDataTypesFileNumnber
    
    Do While Not EOF(nMacroDataTypesFileNumnber)
        'Read a line from the Macro DataTypes Definition file, each line contains
        'DataTypeCode,AccessDataType,SQLServerDataType,OracleDataType
        Line Input #nMacroDataTypesFileNumnber, sRecord
        asRecordContents = Split(sRecord, "|")
        sMacroDataType = asRecordContents(0)
        'extract the required database specific data type definition code
        Select Case nDBType
        Case MACRODatabaseType.Access
            sDbDataType = asRecordContents(1)
        Case MACRODatabaseType.sqlserver
            sDbDataType = asRecordContents(2)
        Case MACRODatabaseType.Oracle80
            sDbDataType = asRecordContents(3)
        End Select
        'Add the database specific data type code (sDbDataType) into the mColDataTypeLookUp
        'collection under the key of the Macro Data Type code (sMacroDataType)
        mColDataTypeLookUp.Add sDbDataType, sMacroDataType
    Loop
    
    Close #nMacroDataTypesFileNumnber

Exit Sub

ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CreateMacroDataTypesLookUp", "modNewDatabasesAndTables")
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
Private Function DbSpecificDataType(ByVal sMacroDataType As String) As String
'--------------------------------------------------------------------
'Call with a Macro Data Type code
'Returns a Database Specific Data Type Code
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    DbSpecificDataType = mColDataTypeLookUp(sMacroDataType)
    
Exit Function

ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DbSpecificDataType", "modNewDatabasesAndTables")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'--------------------------------------------------------------------------------
Private Function CallVersionLessThanTagVersion(ByVal sCallingVersion As String, _
                                                ByVal sTagVersion As String) As Boolean
'--------------------------------------------------------------------------------
Dim asCallingVersion() As String
Dim asTagVersion() As String
Dim lCallingVersion As Long
Dim lTagVersion As Long

    On Error GoTo ErrHandler

    asCallingVersion = Split(sCallingVersion, ".")
    asTagVersion = Split(sTagVersion, ".")
    lCallingVersion = Val(Format(asCallingVersion(0), "00") & Format(asCallingVersion(1), "00") & Format(asCallingVersion(2), "0000"))
    lTagVersion = Val(Format(asTagVersion(0), "00") & Format(asTagVersion(1), "00") & Format(asTagVersion(2), "0000"))
    
    If lCallingVersion < lTagVersion Then
        CallVersionLessThanTagVersion = True
    Else
        CallVersionLessThanTagVersion = False
    End If
    
Exit Function

ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CallVersionLessThanTagVersion", "modNewDatabasesAndTables")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function


'--------------------------------------------------------------------------------
Public Sub CreateNewMACRODB(sConnection As String, enDBType As MACRODatabaseType, _
                                bScriptToFile As Boolean, Optional sScriptFileName As String = "", _
                                Optional bDisplayMessages As Boolean = True, Optional sSiteCode As String = "")

'--------------------------------------------------------------------------------
'TA 03/11/2002
'create  a new macro database
'or write to sScriptFileName if bScriptToFile is true
'--------------------------------------------------------------------------------
Dim oCon As ADODB.Connection
Dim vSQL As Variant
Dim sPrefix As String
Dim sSQL As String
Dim sPath As String
Dim sSiteSQL As String

On Error GoTo ErrLabel

    HourglassOn
    
    sPath = App.Path & "\Database Scripts\New Database\"
    
    If enDBType = Oracle80 Then
        sPrefix = "ORA"
    Else
        sPrefix = "MSSQL"
    End If
    
    If sSiteCode = "" Then
        sSiteSQL = ""
    Else
        sSiteSQL = "INSERT INTO MACRODBSETTING VALUES ('datatransfer', 'dbsitename' , '" & LCase(sSiteCode) & "');" & vbCrLf
        sSiteSQL = sSiteSQL & "UPDATE MACRODBSETTING SET SettingValue = 'site' WHERE SettingSection = 'datatransfer' AND SettingKey = 'dbtype';"
    End If
    
    
    If bScriptToFile Then
        If enDBType <> Oracle80 Then
            'add sequence support if sql server
            sSQL = StringFromFile(sPath & "MSSQL_SP_SequenceSupport.sql")
        End If
    'tables
        sSQL = StringFromFile(sPath & sPrefix & "_Tables.sql")
    'indexes
        sSQL = sSQL & StringFromFile(sPath & sPrefix & "_Indexes.sql")
    'sequences
        sSQL = sSQL & StringFromFile(sPath & sPrefix & "_Sequences.sql")
    'dbtimestamp trigger
        sSQL = sSQL & StringFromFile(sPath & sPrefix & "_DBTS_Triggers.SQL")
    'stored procedures
        sSQL = sSQL & StringFromFile(sPath & sPrefix & "_Procs.SQL")
    'inserts
        sSQL = sSQL & StringFromFile(sPath & "Insert.sql")
        StringToFile sScriptFileName, sSQL & vbCrLf & sSiteSQL
    Else
        Set oCon = New ADODB.Connection
        oCon.Open sConnection
        If enDBType <> Oracle80 Then
            'add sequence support if sql server
            ExecuteMultiLineSQL oCon, StringFromFile(sPath & "MSSQL_SP_SequenceSupport.sql")
        End If
    'tables
        ExecuteMultiLineSQL oCon, StringFromFile(sPath & sPrefix & "_Tables.sql")
    'indexes
        ExecuteMultiLineSQL oCon, StringFromFile(sPath & sPrefix & "_Indexes.sql")
    'sequences
        ExecuteMultiLineSQL oCon, StringFromFile(sPath & sPrefix & "_Sequences.sql")
    'dbtimestamp trigger
        ExecuteMultiLineSQL oCon, StringFromFile(sPath & sPrefix & "_DBTS_Triggers.SQL")
    'stored procdedures
        ExecuteMultiLineSQL oCon, StringFromFile(sPath & sPrefix & "_Procs.SQL")
    'inserts
        ExecuteMultiLineSQL oCon, StringFromFile(sPath & "Insert.sql") & vbCrLf & sSiteSQL
        
        oCon.Close
    End If
    



    
    HourglassOff
    'CreatDB has been called from Restore site database then
    'exit now without displaying any messages
    If bDisplayMessages Then

        'If its a Data database then remind user for the need to register the database
        If bScriptToFile Then
            DialogInformation "Script file " & sScriptFileName & " created"
        Else
            'Generate the relevant glog entries and messages.
            goUser.gLog goUser.UserName, gsCREATE_DB, "MACRO database created."
            DialogInformation "MACRO database created."
        End If
    End If

Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|modNewDatabaseandTables.CreateNewMACODB"

End Sub


'--------------------------------------------------------------------------------
Public Sub CreateNewSecurityDB(sConnection As String, enDBType As MACRODatabaseType, Optional sScriptFileName As String = "", _
                                Optional bScriptToFile As Boolean = False, Optional bDisplayMessages As Boolean = False)

'--------------------------------------------------------------------------------
'TA 03/11/2002
'create  a new macro database
'or write to sScriptFileName if bScriptToFile is true
'--------------------------------------------------------------------------------
Dim oCon As ADODB.Connection
Dim vSQL As Variant
Dim sPrefix As String
Dim sSQL As String
Dim sPath As String

On Error GoTo ErrLabel

    HourglassOn
    
    sPath = App.Path & "\Database Scripts\New Database\"
    
    If enDBType = Oracle80 Then
        sPrefix = "ORA"
    Else
        sPrefix = "MSSQL"
    End If

    If bScriptToFile Then

    'tables
        sSQL = StringFromFile(sPath & sPrefix & "_Security_Tables.sql")
    'inserts
        sSQL = sSQL & StringFromFile(sPath & "Security_Insert.sql")

        StringToFile sScriptFileName, sSQL
    Else
        
        Set oCon = New ADODB.Connection
        oCon.Open sConnection
        'tables
        ExecuteMultiLineSQL oCon, StringFromFile(sPath & sPrefix & "_Security_Tables.sql")
        'inserts
        ExecuteMultiLineSQL oCon, StringFromFile(sPath & "Security_Insert.sql")
        oCon.Close
    End If
    


    HourglassOff
    
    If bDisplayMessages Then
        If bScriptToFile Then
            DialogInformation "Script file " & sScriptFileName & " created"
        End If
    End If

Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|modNewDatabaseandTables.CreateNewMACODB"

End Sub

'--------------------------------------------------------------------------------
Public Function CreateOrRegisterSecurityDB(Optional SysAdmin As Boolean = False, Optional bRegister As Boolean = True, _
                                           Optional bRestoreSite As Boolean = False) As String
'--------------------------------------------------------------------------------
'REM 23/01/03
'Create or register a new security database
'Returns the security connection string
'--------------------------------------------------------------------------------
Dim lCreateRegister As Long
Dim lDBType As Long
Dim sScriptType As String

    On Error GoTo ErrLabel
    
If Not SysAdmin Then
#If (SM = 1) Or (SD = 1) Or (DM = 1) Then
    lCreateRegister = frmOptionMsgBox.Display(GetApplicationTitle, "Security database does not exist", "Please select one of the following:", "Create and register new security database|Register security database|Write DB script to file|Exit MACRO", "&OK", "", True, False)
#End If

Else
    If bRegister Then
        lCreateRegister = 2
    Else
        lCreateRegister = 1
    End If
End If


#If (SM = 1) Or (SD = 1) Or (DM = 1) Then
    Select Case lCreateRegister
    Case 1 'create database
        'Create DB containing both security and MACRO DB in same schema
        CreateOrRegisterSecurityDB = CreateNewSecurityDatabase
        
        If bRegister Then
            'Register DB
            SecurityDatabasePath = CreateOrRegisterSecurityDB
        End If
        
    Case 2 'register database
    
        CreateOrRegisterSecurityDB = RegisterNewSecurityDatabase(bRestoreSite)
    
    Case 3 ' script to file
        lDBType = frmOptionMsgBox.Display(GetApplicationTitle, "Database Type", "Please select a database type", "SQL Server/MSDE|Oracle", "&OK", "", True, False)
        If lDBType = 1 Then 'SQL SERVER
            sScriptType = CONNECTION_SQLOLEDB
        Else 'Oracle
            sScriptType = CONNECTION_MSDAORA
        End If
        CreateOrRegisterSecurityDB = CreateNewSecurityDatabase(, True, sScriptType)
    
    Case 4 'quit
    
        CreateOrRegisterSecurityDB = ""
        
    End Select
#End If
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modNewDatabaseandTables.CreateOrRegisterSecurityDB"
End Function

'--------------------------------------------------------------------------------
Public Function CreateNewSecurityDatabase(Optional bRestoreSiteDB As Boolean = False, _
                                          Optional bWriteDBScriptToFile As Boolean = False, _
                                          Optional sScriptType As String = "") As String
'--------------------------------------------------------------------------------
'REM 23/01/03
'Routine to create a new security database and MACRO database in one schema
'--------------------------------------------------------------------------------
Dim tCon As udtConnection
Dim sSecCon As String
Dim sDBAlias As String
Dim sSiteCode As String
'Dim bWriteDBScriptToFile As Boolean

    On Error GoTo ErrLabel

If Not bWriteDBScriptToFile Then
#If (SM = 1) Or (SD = 1) Or (DM = 1) Then
    sSecCon = frmConnectionString.Display(False, True, True, sDBAlias, sSiteCode)
#End If
Else
    sSecCon = "PROVIDER=" & sScriptType & ";"
End If

    If sSecCon = "" Then
        If bRestoreSiteDB Then
            CreateNewSecurityDatabase = sSecCon
        Else
            
            'MACROEnd
            
        End If
    Else

        tCon = Connection_AsType(sSecCon)
        
        Select Case Connection_Property(CONNECTION_PROVIDER, sSecCon)
        
        Case CONNECTION_SQLOLEDB 'SQL server
            If bWriteDBScriptToFile Then
                Call WriteDBScriptToFile(sSecCon)
                sSecCon = ""
            Else
                Call CreateNewSecurityDB(sSecCon, sqlserver)
                Call CreateAndRegFirstDB(tCon.Datasource, tCon.Database, "", sDBAlias, tCon.UserId, tCon.Password, sSiteCode)
            End If
        Case Else 'Oracle
            If bWriteDBScriptToFile Then
                Call WriteDBScriptToFile(sSecCon)
                sSecCon = ""
            Else
                Call CreateNewSecurityDB(sSecCon, Oracle80)
                Call CreateAndRegFirstDB("", tCon.Database, tCon.Datasource, sDBAlias, tCon.UserId, tCon.Password, sSiteCode)
            End If
        End Select


    End If

    CreateNewSecurityDatabase = sSecCon
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modNewDatabaseandTables.CreateNewSecurityDatabase"
End Function

'--------------------------------------------------------------------------------
Private Sub WriteDBScriptToFile(sSecCon As String)
'--------------------------------------------------------------------------------
'REM 12/02/03
'Writes the database script to file
' NCJ 15 Jan 04 - Corrected "SecuirtyDB" to "SecurityDB"
'--------------------------------------------------------------------------------
Dim enDBType As MACRODatabaseType
Dim sMACROFileName As String
Dim sSecurityFileName As String
Dim sSiteCode As String

    On Error GoTo ErrLabel

        Select Case Connection_Property(CONNECTION_PROVIDER, sSecCon)
            
        Case CONNECTION_SQLOLEDB 'SQL server
            enDBType = sqlserver
        Case CONNECTION_MSDAORA
            enDBType = Oracle80
        End Select

        sMACROFileName = InputBox("Please enter the name of the MACRO database script file", "MACRO Script File", "MACRODB")
        If sMACROFileName = "" Then Exit Sub
        sSiteCode = InputBox("Please enter a site name if this is a site database", "MACRO Script File")
        Call CreateNewMACRODB(sSecCon, enDBType, True, sMACROFileName, , sSiteCode)
        sSecurityFileName = InputBox("Please enter the name of the MACRO security database script file", "MACRO Script File", "SecurityDB")
        If sSecurityFileName = "" Then sSecurityFileName = "SecurityDB"
        Call CreateNewSecurityDB(sSecCon, enDBType, sSecurityFileName, True, True)
        
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modNewDatabaseandTables.WriteDBScriptToFile"
End Sub

'--------------------------------------------------------------------------------
Public Function RegisterNewSecurityDatabase(bRestoreSite As Boolean) As String
'--------------------------------------------------------------------------------
'REM 23/01/03
'Routine for registering a new security database
'--------------------------------------------------------------------------------
Dim sSecCon As String
    
    On Error GoTo ErrLabel
    
    sSecCon = ""
#If (SM = 1) Or (SD = 1) Or (DM = 1) Then
    sSecCon = frmConnectionString.Display(True, False, False)
#End If

    'if not  restoring a site then register security database
    If Not bRestoreSite Then
        SecurityDatabasePath = sSecCon
    End If

    RegisterNewSecurityDatabase = sSecCon

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modNewDatabaseandTables.RegisterNewSecurityDatabase"
End Function



Public Function CreateOrRegisterSecdb(bCreate As Boolean, bRegister As Boolean, bFirstUse As Boolean) As String
'create or register new security database
Dim sSecCon As String
Dim nResult As Integer
Dim sServer As String
Dim sDBName As String
Dim sTNSName As String
Dim sUserId As String
Dim sPwd As String
Dim tCon As udtConnection


        If bFirstUse Then
            DialogInformation "You have no security database, one will be created"
        End If

#If SM = 1 Then
        sSecCon = frmConnectionString.Display(True, True, True, "", "")
#End If

        If sSecCon = "" Then
            MACROEnd
        Else
            'we need to create as well
            If bCreate Then
                tCon = Connection_AsType(sSecCon)
                Select Case Connection_Property(CONNECTION_PROVIDER, sSecCon)
                Case CONNECTION_SQLOLEDB
                    CreateNewSecurityDB sSecCon, sqlserver

                    If bFirstUse Then
                        CreateAndRegFirstDB tCon.Datasource, tCon.Database, "", "", tCon.UserId, tCon.Password, ""
                    End If
                Case Else
                    CreateNewSecurityDB sSecCon, Oracle80
                    If bFirstUse Then
                        CreateAndRegFirstDB "", tCon.Database, tCon.Datasource, "", tCon.UserId, tCon.Password, ""

                    End If
                End Select
            End If

        End If

        If bRegister Then
            'put this in the settings file
            SecurityDatabasePath = sSecCon
        End If

        CreateOrRegisterSecdb = sSecCon

End Function

'--------------------------------------------------------------------------------
Private Sub CreateAndRegFirstDB(sServerName As String, sDBName As String, sTNSName As String, sDBAlias As String, _
                                sUserId As String, sPwd As String, sSiteCode As String)
'--------------------------------------------------------------------------------
'
'Routine to register first created database
'--------------------------------------------------------------------------------
Dim oCon As ADODB.Connection
Dim sSQL As String
Dim sEncryptedUserId As String
Dim sEncryptedPwd As String
Dim sCon As String
'Dim sSiteCode As String
    
    On Error GoTo ErrLabel

    If sUserId = "" Then
        sEncryptedUserId = "null"
    Else
        
        sEncryptedUserId = "'" & EncryptString(sUserId) & "'"
    End If
    
    If sPwd = "" Then
        sEncryptedPwd = "null"
    Else
        sEncryptedPwd = "'" & EncryptString(sPwd) & "'"
    End If
    
    Set oCon = New ADODB.Connection
    'sSiteCode = InputBox("MACRO is creating a defualt database, please enter a site code if this is a remote site database)")
    
    If sTNSName = "" Then
        If sDBAlias = "" Then
            'Only use first 15 characters as database table can only hold 15 chars for code
            sDBAlias = Left(sDBName, 15)
        End If
        sSQL = " INSERT INTO Databases " _
            & "(DataBaseCode, DatabaseType, Servername, Nameofdatabase, DatabaseUser,databasepassword)" _
            & " VALUES ('" & sDBAlias & "'," & MACRODatabaseType.sqlserver & ",'" & sServerName & "','" & sDBName & "'," & sEncryptedUserId & "," & sEncryptedPwd & ")"
        sCon = Connection_String(CONNECTION_SQLOLEDB, sServerName, sDBName, sUserId, sPwd)
    CreateNewMACRODB sCon, sqlserver, False, , False, sSiteCode
    Else
        If sDBAlias = "" Then
            sDBAlias = Left(sTNSName, 15)
        End If
            sSQL = " INSERT INTO Databases " _
            & "(DataBaseCode, DatabaseType, NameofDatabase, DatabaseUser, DatabasePassword)" _
            & " VALUES ('" & sDBAlias & "'," & MACRODatabaseType.Oracle80 & ",'" & sTNSName & "'," & sEncryptedUserId & "," & sEncryptedPwd & ")"
    
     sCon = Connection_String(CONNECTION_MSDAORA, sTNSName, "", sUserId, sPwd)
        CreateNewMACRODB sCon, Oracle80, False, , False, sSiteCode
    End If
    oCon.Open sCon
    oCon.Execute sSQL
    
    'if its a new site DB then make default user rde not a system admin
    If sSiteCode <> "" Then
        oCon.Execute "UPDATE MACROUser SET SysAdmin = 0 WHERE UserName = 'rde'"
    End If
    
    sSQL = "INSERT into UserDatabase VALUES ('rde','" & sDBAlias & "')"
    oCon.Execute sSQL
    oCon.Execute "INSERT INTO UserRole  VALUES ('rde','MACROUser','AllStudies','AllSites',1)"
    oCon.Close

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modNewDatabaseandTables.CreateAndRegFirstDB"
End Sub

