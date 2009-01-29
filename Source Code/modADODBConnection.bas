Attribute VB_Name = "modADODBConnection"
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998-2006. All Rights Reserved
'   File:       MADODBConnection.cls
'   Author:     Paul Norris, 07/09/99
'   Purpose:    ADODB connection initialization and storage.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   1   PN  14/09/99    Updated code to conform to VB standards doc version 1.0
'   1   PN  15/09/99    Added security database access using ADO
'   2   WillC 11/10/99  Added the error handlers
'   ATN 21/12/99        Added connection strings for Oracle
'   4   Willc 5/1/00    Added  global String to hold connection strings from frmADOConnect
'   NCJ 18/1/00     Moved gsSecurityDatabasePassword here from MainMACROMOdule
'   NCJ 25/1/00     Removed private ADODB connection variables and made sure
'                   the global ones are terminated when MACRO exits
'   Mo Morris       28/1/00
'                   InitializeSecurityADODBConnection now calls UpgradeSecurityDatabase
'                   InitializeMacroADODBConnection now calls UpgradeDataDatabase
'   Mo Morris       27/3/00
'                   changes stemming from changing MacroADODBConnection from a global variable
'                   in MainMACROModule to to a Property Get in modADODBConnection, for the
'                   purpose of logging data base accesses.
'   TA  27/03/2000  restart timeout timer every time database is accessed
'   Mo 30/8/01      DoesDatabaseExist moved here from (mod) Security)
'   ATO 14/01/2002  Added routine IsDbSizeEnough
'   REM 28/06/02 - CBB 2.2.18 No. 5 - In InitializeMacroADODBConnection() Set Time out to 0 so won't crash after 30 seconds
'   TA 24/08/2002: We now use a client side sursoe for booth Oracle and SQL Server
'   ASH 12/9/2002 Changed gUser.Database to gUser.NameOfDatabase in InitializeMacroADODBConnection
'   TA 06/11/2002: Added ExecuteMultiLineSQL
'ASH 16/01/2003: Stripped off user id and password from SecurityDatabasePath in InitializeSecurityADODBConnection
' NCJ 17 Jan 03 - Only rollback if there's something to roll back!
' NCJ 24 May 06 - Watch out for moMacroADODBConnection being nothing (like after app error)
'--------------------------------------------------------------------------------
Option Explicit

Public gsADOConnectString As String

Public SecurityDatabaseType As MACRODatabaseType

'   ATN 14/1/2000
Public Const gsSecurityDatabasePassword = "M58NP75BJA"

'Changed Mo Morris 27/3/00
'Local form of MacroADODBConnection
Private moMacroADODBConnection As ADODB.Connection
Private moNewMacroCon As ADODB.Connection


' NCJ 25/1/00 - Global ADODB connections now used (see MAINMacroModule)
'' this is the adodb.connection object that persists throughout the life
'' of the application
'' all Macro database access shall use this connection
'Private moMacroADODBConnection As ADODB.Connection
'
'' all security access should use this connection
'Private moSecurityADODBConnection As ADODB.Connection

'--------------------------------------------------------------------------------
Public Sub TerminateAllADODBConnections()
'--------------------------------------------------------------------------------
' PN 15/09/99 terminate all connections properly
'--------------------------------------------------------------------------------
    Call TerminateMacroADODBConnection
    Call TerminateSecurityADODBConnection

End Sub

'--------------------------------------------------------------------------------
Public Sub TerminateMacroADODBConnection()
'--------------------------------------------------------------------------------
' PN 15/09/99 close the macro connection if it is open
' NCJ 25/1/00 - Changed to use Global variables
'--------------------------------------------------------------------------------
    If Not MacroADODBConnection Is Nothing Then
        ' close and terminate the macro connection
        'changed Mo Morris 27/3/00, from MacroADODBConnection to moMacroADODBConnection
        moMacroADODBConnection.Close
        Set moMacroADODBConnection = Nothing
    End If

End Sub

'--------------------------------------------------------------------------------
Public Sub TerminateSecurityADODBConnection()
'--------------------------------------------------------------------------------
' PN 15/09/99 close the security connection if it is open
' NCJ 25/1/00 - Changed to use Global variables
'--------------------------------------------------------------------------------
    If Not SecurityADODBConnection Is Nothing Then
        ' close and terminate the security connection
        SecurityADODBConnection.Close
        Set SecurityADODBConnection = Nothing
    End If

End Sub

'--------------------------------------------------------------------------------
Public Sub InitializeMacroADODBConnection(Optional bUpgrade As Boolean = True)
'--------------------------------------------------------------------------------
'This will initialize the ado connection to the database
'Re-written by Mo Morris 27/10/99
'Upgrade are only run if bUpgrade is true
'--------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    'changed Mo Morris 27/3/00, from MacroADODBConnection to moMacroADODBConnection
    Set moMacroADODBConnection = New ADODB.Connection
    
    Select Case goUser.Database.DatabaseType
'    Case MACRODatabaseType.Access
'            'JET OLE DB NATIVE PROVIDER
'            'Note that an Access database connection requires the full database path
'            '   ATN 14/1/2000
'            '   Use database password
'            gsADOConnectString = Connection_String(CONNECTION_MSJET_OLEDB_40, gUser.DatabasePath, , , gUser.DatabasePassword)
'            moMacroADODBConnection.Open gsADOConnectString
'
'            'ATO 14/01/2002 checks the size of current ACCESS DATABASE
'             Call IsDbSizeEnough(gUser.DatabasePath)
    
    Case MACRODatabaseType.sqlserver
            'SQL SERVER OLE DB NATIVE PROVIDER

            gsADOConnectString = Connection_String(CONNECTION_SQLOLEDB, goUser.Database.ServerName, goUser.Database.NameOfDatabase, _
                          goUser.Database.DatabaseUser, goUser.Database.DatabasePassword)
            moMacroADODBConnection.Open gsADOConnectString

            'TA 24/08/2002: USe client side cursor for ssql server (apart from patient export)
             moMacroADODBConnection.CursorLocation = adUseClient
            'REM 28/06/02 - CBB 2.2.18 No. 5 - Set Time out to 0 so won't crash after 30 seconds
            moMacroADODBConnection.CommandTimeout = 0
            
'   ATN 21/12/99
'   Added connection to Oracle
    Case MACRODatabaseType.Oracle80
            'Oracle OLE DB NATIVE PROVIDER

            gsADOConnectString = Connection_String(CONNECTION_MSDAORA, goUser.Database.NameOfDatabase, , _
               goUser.Database.DatabaseUser, goUser.Database.DatabasePassword)
            moMacroADODBConnection.Open gsADOConnectString

            '   Need to use client side cursors for Oracle
            moMacroADODBConnection.CursorLocation = adUseClient
    End Select
    
    If bUpgrade Then
        'REM 19/06/03 - Check to see if  the database is a version of MACRO 3.0,
        'if not tell user to upgrade in SM and exit MACRO
        If MACRO30Version(moMacroADODBConnection) Then
            UpgradeDataDatabase
        Else
            Call DialogInformation("This is not a MACRO 3.0 database and must be upgraded using System Management.")
            MACROEnd
        End If
    End If
 
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "InitializeMacroADODBConnection", "modADODBConnection.bas ")
 
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'--------------------------------------------------------------------------------
Private Function MACRO30Version(oMACROCon As ADODB.Connection) As Boolean
'--------------------------------------------------------------------------------
'REM 19/06/03
'Check to see if a database is a MACRO 3.0 version
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsDatabase As ADODB.Recordset
Dim sDatabaseVersion As String
    
    On Error GoTo ErrLabel

    sSQL = "SELECT MACROVERSION FROM MACROCONTROL"
    Set rsDatabase = New ADODB.Recordset
    rsDatabase.Open sSQL, oMACROCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsDatabase.RecordCount <> 0 Then
        sDatabaseVersion = rsDatabase!MACROVersion
        If sDatabaseVersion = "3.0" Then
            MACRO30Version = True
        Else
            MACRO30Version = False
        End If
    Else
        MACRO30Version = False
    End If
    
Exit Function
ErrLabel:
    MACRO30Version = False
End Function

'--------------------------------------------------------------------------------
Public Function InitializeSecurityADODBConnection() As String
'--------------------------------------------------------------------------------
'This will initialize the ado connection to the database
'Re-written by Mo Morris 27/10/99
'TA 18/11/2002: reutrns connection string
'ASH 16/01/2003: Stripped off user id and password from SecurityDatabasePath
'--------------------------------------------------------------------------------
Dim sSecCon As String
Dim vSecString As Variant
Dim sNewSecurityString As String
Dim rs As ADODB.Recordset
Dim lDBCount As Long

    On Error GoTo ErrHandler
    
    ' NCJ 25/1/00 - Close if already in existence
    If Not SecurityADODBConnection Is Nothing Then
        SecurityADODBConnection.Close
    End If
    
    sSecCon = SecurityDatabasePath
    
    'TA 18/11/2002: create or register secutiry db if needed
    If sSecCon = "" Then
    
#If SM = 1 Then
        'sSecCon = CreateOrRegisterSecdb(True, True, True)
        sSecCon = CreateOrRegisterSecurityDB
#End If

#If SD = 1 Then
        sSecCon = CreateOrRegisterSecurityDB
#End If

#If DM = 1 Then
        sSecCon = CreateOrRegisterSecurityDB
#End If

        If sSecCon = "" Then MACROEnd
        
    End If
    
    Select Case Connection_Property(CONNECTION_PROVIDER, sSecCon)
    Case "MSDAORA": SecurityDatabaseType = Oracle80
    Case "SQLOLEDB": SecurityDatabaseType = sqlserver
    End Select
    
    
    Set SecurityADODBConnection = New ADODB.Connection
    SecurityADODBConnection.Open sSecCon
    SecurityADODBConnection.CursorLocation = adUseClient
        
    'Changed Mo 28/1/00, check for need to upgrade Security database
    UpgradeSecurityDatabase
    
   InitializeSecurityADODBConnection = sSecCon
 
Exit Function
ErrHandler:
    If Err.Number = 3706 Then
        '   Don't show form in WEBRDE dll
        #If WebRDE <> -1 Then
            frmMDAC.Show vbModal
        #Else
            Call ExitMACRO
            Call MACROEnd
        #End If
        'WillC 23/3/00 SR2959 trap if macro can't find its security database
    ElseIf Err.Number = -2147467259 Then
            'ASH 16/01/2003: Stripped off user id and password from SecurityDatabasePath
                vSecString = Split(SecurityDatabasePath, ";")
                Select Case Connection_Property(CONNECTION_PROVIDER, sSecCon)
                    Case "MSDAORA": SecurityDatabaseType = Oracle80
                        sNewSecurityString = vSecString(0) & ";" & vSecString(1) & ";"
                    Case "SQLOLEDB": SecurityDatabaseType = sqlserver
                        sNewSecurityString = vSecString(0) & ";" & vSecString(1) & ";" & vSecString(2) & ";"
                End Select
            MsgBox "MACRO could not find the security database:" & vbCrLf _
                & vbCrLf _
                & sNewSecurityString & vbCrLf _
                & vbCrLf _
                & "Please place a valid security database in this location.", vbInformation, "MACRO"
               Call MACROEnd
    Else
        Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "InitializeSecurityADODBConnection", "modADODBConnection.bas ")
              Case OnErrorAction.Ignore
                  Resume Next
              Case OnErrorAction.Retry
                  Resume
              Case OnErrorAction.QuitMACRO
                  Call ExitMACRO
                  Call MACROEnd
         End Select
    End If
End Function


'--------------------------------------------------------------------------------
Public Sub TransBegin()
'--------------------------------------------------------------------------------
'Begin Transaction or increment the nested level variable gnTransactionControlOn
'--------------------------------------------------------------------------------

    If gnTransactionControlOn = 0 Then
        MacroADODBConnection.BeginTrans
       gnTransactionControlOn = 1
    Else
        gnTransactionControlOn = gnTransactionControlOn + 1
    End If
    
    'Debug.Print "BeginTrans gnTransactionControlOn = " & gnTransactionControlOn

End Sub

'--------------------------------------------------------------------------------
Public Sub TransCommit()
'--------------------------------------------------------------------------------
'Commit Transaction if outer nested level has been reached
'--------------------------------------------------------------------------------

    If gnTransactionControlOn = 1 Then
        MacroADODBConnection.CommitTrans
        gnTransactionControlOn = 0
    Else
        gnTransactionControlOn = gnTransactionControlOn - 1
    End If
    
    'Debug.Print "TransCommit gnTransactionControlOn = " & gnTransactionControlOn

End Sub

'--------------------------------------------------------------------------------
Public Sub TransRollBack()
'--------------------------------------------------------------------------------
'RollBack transaction
' NCJ 17 Jan 03 - Only rollback if there's something to roll back!
'--------------------------------------------------------------------------------
    
    If gnTransactionControlOn > 0 Then      ' NCJ 17 Jan 03
        MacroADODBConnection.RollbackTrans
    End If
    
    gnTransactionControlOn = 0
    
End Sub

'--------------------------------------------------------------------------------
Public Property Get MacroADODBConnection() As ADODB.Connection
'--------------------------------------------------------------------------------

    'TA 10/04/2003: if we are not in a transaction, check to see if connection is open
        'if not then reinitialise it
        'this is needed after windows hibernation mode
        ' NCJ 24 May 06 - Check that moMacroADODBConnection not Nothing
    If gnTransactionControlOn = 0 And Not moMacroADODBConnection Is Nothing Then
       If moMacroADODBConnection.State = adStateClosed Then
            InitializeMacroADODBConnection False
        End If
    End If
    
    Set MacroADODBConnection = moMacroADODBConnection
    'TA 27/03/2000 restart timeout timer every time database is accessed
    Call RestartSystemIdleTimer
    'Debug.Print "get MacroADODBConnection called " & Now

End Property

'--------------------------------------------------------------------------------
Public Function DoesDatabaseExist(sDescription As String) As Boolean
'--------------------------------------------------------------------------------
'Check to see if the database already exists in the Security Databases table
'--------------------------------------------------------------------------------
'Mo Morris  17/8/01 The database path is no longer used as a check parameter.
'                   With the DatabaseDescription being the tables key, no other
'                   field needs to be checked.
'--------------------------------------------------------------------------------
Dim rsDatabase As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler

    'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode)
    sSQL = "SELECT * FROM Databases " _
    & " WHERE DatabaseCode = '" & sDescription & "'"
    
    Set rsDatabase = New ADODB.Recordset
    rsDatabase.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
    
    If rsDatabase.RecordCount > 0 Then
        DoesDatabaseExist = True
    Else
        DoesDatabaseExist = False
    End If
        
    Set rsDatabase = Nothing
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DoesDatabaseExist", "Security.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
       
End Function

'-------------------------------------------------------------------------------------------------------
Public Sub IsDbSizeEnough(sDatabase As String)
'-------------------------------------------------------------------------------------------------------
'This checks the size of the current ACCESS database. Needed since access databases bigger than
'1GB can cause Macro to crash. Gives 2 messages,1 when database over 700mb and the other when database
'over 900mb
'-------------------------------------------------------------------------------------------------------
Dim sMsg As String
Dim objFSO As FileSystemObject
Dim objFile As File
Const WARN_SIZE As Long = 786432000 '750Mb (750 * 1024 * 1024)
Const DANGEROUS_SIZE As Long = 943718400 '900Mb (900 * 1024 * 1024)
    
    On Error GoTo ErrHandler
    
    Set objFSO = New FileSystemObject
    Set objFile = objFSO.GetFile(sDatabase)
    
    Select Case objFile.SIZE
    Case WARN_SIZE To DANGEROUS_SIZE
         sMsg = "Your " & goUser.Database.NameOfDatabase & " database is approaching its 1GB maximum size. " & vbCrLf
        sMsg = sMsg & "Please contact your system administator."
        Call DialogWarning(sMsg)
    Case Is > DANGEROUS_SIZE
        sMsg = "Your " & goUser.Database.NameOfDatabase & " database is dangerously close to its 1GB maximum limit. " & vbCrLf
        sMsg = sMsg & "Please contact your system administator."
        Call DialogWarning(sMsg)
    End Select
    
    Set objFile = Nothing
    Set objFSO = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "IsDbSizeEnough", "modADODBConnection.bas ")
              Case OnErrorAction.Ignore
                  Resume Next
              Case OnErrorAction.Retry
                  Resume
              Case OnErrorAction.QuitMACRO
                  Call ExitMACRO
                  Call MACROEnd
         End Select
End Sub

'--------------------------------------------------------------------------------
Public Sub ExecuteMultiLineSQLRowAtATime(oCon As ADODB.Connection, sMultiLineSQL As String)
'--------------------------------------------------------------------------------
'TA 03/11/2002
'Execute each sql line in an array in turn
'If vSQL is a string an array is created based on splitting on VBCRLF
'--------------------------------------------------------------------------------
Dim i As Long
Dim sSQL As String
Dim vSQL As Variant

On Error GoTo ErrLabel

    If VarType(sMultiLineSQL) = vbString Then
        'string passed in lets split it
        vSQL = Split(sMultiLineSQL, vbCrLf)
    End If

    For i = 0 To UBound(vSQL)
        sSQL = Trim(vSQL(i))
        If sSQL <> "" And UCase(sSQL) <> "GO" Then
            'ensure we don't have blank row or a GO line in SQL server
            If Right(sSQL, 1) = ";" Then
                'strip off ";"
                sSQL = Left(sSQL, Len(sSQL) - 1)
            End If
            oCon.Execute sSQL
        End If
    Next
    
Exit Sub

ErrLabel:

    If MACROErrorHandler("modNewDatabaseandTables", Err.Number, Err.Description & " - " & sSQL, "ExecuteMultiLineSQLRowAtATime", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'--------------------------------------------------------------------------------
Public Sub ExecuteMultiLineSQL(oCon As ADODB.Connection, sMultiLineSQL As String, Optional sSeparator As String = "--//--")
'--------------------------------------------------------------------------------
'TA 03/11/2002
'Execute each sql line in an array in turn
'If vSQL is a string an array is created based on splitting on VBCRLF
'RS 15/01/2003
' Based on the original routine by TA. this routine will execute multiline
' statements, separated by a given separator line
'--------------------------------------------------------------------------------
Dim i As Long
Dim sSQL As String
Dim vSQL As Variant

On Error GoTo ErrLabel

    If VarType(sMultiLineSQL) = vbString Then
        'string passed in lets split it
        vSQL = Split(sMultiLineSQL, vbCrLf)
    End If

    ' RS 27/01/2003: If the first line dos not contain 'SEPARATOR', call original function
    If InStr(1, vSQL(0), "SEPARATOR") = 0 Then
        ' Call original function
        vSQL = ""
        Call ExecuteMultiLineSQLRowAtATime(oCon, sMultiLineSQL)
        Exit Sub
    End If

    ' Loop through the array, and build SQL statements, Line starting with <sSeparator> is considered separator
    sSQL = ""
    For i = 0 To UBound(vSQL)
        If Left(vSQL(i), 6) = sSeparator Or UCase(Left(vSQL(i), 2)) = "GO" Then
            ' Separator found, execute current SQL statement, if any
            If sSQL <> "" And UCase(sSQL) <> "GO" Then
                ' Trim off leading linefeeds & spaces
                Do While Left(sSQL, 1) = vbCrLf Or Left(sSQL, 1) = " " Or Left(sSQL, 1) = Chr(10) Or Left(sSQL, 1) = Chr(13)
                    sSQL = Mid(sSQL, 2, 9999)
                Loop
                ' Trim off trailing linefeeds & spaces
                Do While Right(sSQL, 1) = vbCrLf Or Right(sSQL, 1) = " " Or Right(sSQL, 1) = Chr(10) Or Right(sSQL, 1) = Chr(13)
                    sSQL = Left(sSQL, Len(sSQL) - 1)
                Loop
                ' Trim off trailing semicolon
                If Right(sSQL, 1) = ";" Or Right(sSQL, 1) = "/" And Left(oCon.Provider, 7) = "MSDAORA" Then
                    'strip off ";"
                    sSQL = Left(sSQL, Len(sSQL) - 1)
                End If
                ' Execute the statement
                'Debug.Print sSQL
                If sSQL <> "" Then
                    oCon.Execute sSQL
                End If
                ' Reset statement
                sSQL = ""
            End If
        Else
            ' Not a separator, add to current SQL statement
            sSQL = sSQL & vSQL(i) & vbCrLf
        End If
    Next
    
    ' Process final SQL statement
    If sSQL <> "" And UCase(sSQL) <> "GO" Then
        ' Trim off leading linefeeds & spaces
        Do While Left(sSQL, 1) = vbCrLf Or Left(sSQL, 1) = " " Or Left(sSQL, 1) = Chr(10) Or Left(sSQL, 1) = Chr(13)
            sSQL = Mid(sSQL, 2, 9999)
        Loop
        ' Trim off trailing linefeeds & spaces
        Do While Right(sSQL, 1) = vbCrLf Or Right(sSQL, 1) = " " Or Right(sSQL, 1) = Chr(10) Or Right(sSQL, 1) = Chr(13)
            sSQL = Left(sSQL, Len(sSQL) - 1)
        Loop
        ' Trim off trailing semicolon (NOT FOR ORACLE)
        If Right(sSQL, 1) = ";" Or Right(sSQL, 1) = "/" And Left(oCon.Provider, 7) = "MSDAORA" Then
            'strip off ";"
            sSQL = Left(sSQL, Len(sSQL) - 1)
        End If
        ' Execute the statement
        If sSQL <> "" Then
            oCon.Execute sSQL
        End If
        ' Reset statement
        sSQL = ""
    End If
   
Exit Sub

ErrLabel:
    If MACROErrorHandler("modNewDatabaseandTables", Err.Number, Err.Description & " - " & sSQL, "ExecuteMultiLineSQL", Err.Source) = Retry Then
        Resume
    End If

    
End Sub

