Attribute VB_Name = "modADOConnection"
Option Explicit

''--------------------------------------------------------------------------------------------------
'Private Function OpenSecurityDb() As ADODB.Connection
''--------------------------------------------------------------------------------------------------
'' REM 31/05/01
'' Function to open Security Database Connection
''--------------------------------------------------------------------------------------------------
'Dim sSecurityDatabasePath As String
'Dim sDatabasePswd As String
'Dim sConnectionString As String
'Dim conSecurityCnn As ADODB.Connection
'
'    On Error GoTo Errhandler
'
'    ' variables assigned the Database path and database password
'    sSecurityDatabasePath = MACROGetFromRegistry("2.1", "SecurityPath")
'    sDatabasePswd = "M58NP75BJA"
'
'    Debug.Print "DB Connect = " & sSecurityDatabasePath
'
'    ' variable assigned ADO connection string to the security database
'    ' and the database password
'    sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " & _
'                         "Data Source=" & sSecurityDatabasePath & "; " & _
'                         "Jet OLEDB:Database Password=" & sDatabasePswd
'
'
'    'Creat new connection object and opens the connection
'    Set conSecurityCnn = New ADODB.Connection
'    conSecurityCnn.Open sConnectionString
'    conSecurityCnn.CursorLocation = adUseClient
'
'
'    Set OpenSecurityDb = conSecurityCnn
'    Exit Function
'
'Errhandler:
'    Debug.Print "Connection Error to Security Databse"
'
'    ' Must check what to do on error in this function???
'
'End Function
'
''--------------------------------------------------------------------------------------------------
'Public Function cnnMACRO(ByVal sDatabase As String) As ADODB.Connection
''--------------------------------------------------------------------------------------------------
'' REM 06/06/01
'' Opens one of three Database Connections depending on Database Description passed.
'' This is done by using the Database Description to extract a recordset from the Security Database
'' containing the properties of the required Database, which are then used to select and open the
'' specific Database connection.
''--------------------------------------------------------------------------------------------------
'Dim sDatabasePswd As String
'Dim nDatabaseType As Integer
'Dim sDatabaseName As String
'Dim sServer As String
'Dim sDatabaseUser As String
'Dim sConnectionString As String
'Dim conSecurityCnn As ADODB.Connection
'Dim conDatabaseCnn As ADODB.Connection
'Dim rsSecurity As ADODB.Recordset
'Dim sSQL As String
'Dim sProvider As String
'
'    On Error GoTo Errhandler
'
'    ' Open connection to Security Database via OpenSecurityDb() Function
'    Set conSecurityCnn = OpenSecurityDb()
'
'    ' Retrieves the Database Location, Database Type, Server Name, Name of Database, Database User
'    ' and Database Password for a specific Database Description from the Security Database.
'    sSQL = "SELECT Databases.DatabaseDescription, Databases.DatabaseLocation,Databases.DatabaseType," & _
'            "Databases.ServerName,Databases.NameOfDatabase,Databases.DatabaseUser," & _
'            "Databases.DatabasePassword" & _
'            " FROM Databases" & _
'            " WHERE Databases.DatabaseDescription='" & sDatabase & "'"
'    Debug.Print sSQL
'
'    Set rsSecurity = New ADODB.Recordset
'    rsSecurity.Open sSQL, conSecurityCnn
'
'
'    ' Assign fields from the recordset to variables used to open Database connection
'    ' Check for Null values in the recordset fields and if so change them to empty strings
'    nDatabaseType = removenull(rsSecurity.Fields(2).Value)
'    If nDatabaseType = 0 Then
'        sServer = removenull(rsSecurity.Fields(1).Value)
'    Else
'        sServer = removenull(rsSecurity.Fields(3).Value)
'    End If
'    sDatabaseName = removenull(rsSecurity.Fields(4).Value)
'    sDatabaseUser = removenull(rsSecurity.Fields(5).Value)
'    sDatabasePswd = removenull(rsSecurity.Fields(6).Value)
'
'    Debug.Print "Recordset Returned:" & nDatabaseType, sServer, sDatabaseName, sDatabaseUser, sDatabasePswd
'
'
'    Select Case nDatabaseType
'        Case 0 ' Access Database
'            sProvider = CONNECTION_MSJET_OLEDB_351
'
'        Case 1 ' SQL Server
'            sProvider = CONNECTION_SQLOLEDB
'
'        Case 3 ' ORACLE
'            sProvider = CONNECTION_MSDAORA
'
'        Case Else ' Exits Function as Database Type has not been defined in the Security Database
'            Exit Function
'
'    End Select
'
'    ' Build different connection string for each type of database
'    sConnectionString = Connection_String(sProvider, sServer, sDatabaseName, sDatabaseUser, sDatabasePswd)
'
'    'Creat new connection object and open the connection
'    Set conDatabaseCnn = New ADODB.Connection
'    conDatabaseCnn.Open sConnectionString
'    conDatabaseCnn.CursorLocation = adUseClient
'
'    Set cnnMACRO = conDatabaseCnn
'    Exit Function
'
'Errhandler:
'    Debug.Print "Error in cnnMACRO Function"
'
'    ' Check what to do on error
'
'End Function

Public Function removenull(vValue As Variant) As String

    If IsNull(vValue) Then
        removenull = ""
    Else
        removenull = Format(vValue)
    End If

    

End Function

