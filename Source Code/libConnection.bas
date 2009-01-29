Attribute VB_Name = "libConnection"
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   File:       libConnection.bas
'   Author:     Toby Aldridge May 2001
'   Purpose:    ADODB connection string manipulation
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------

Option Explicit

'these must be in capitals
Public Const CONNECTION_MSJET_OLEDB_351 = "MICROSOFT.JET.OLEDB.3.51"
Public Const CONNECTION_MSDAORA = "MSDAORA"
Public Const CONNECTION_MSJET_OLEDB_40 = "MICROSOFT.JET.OLEDB.4.0"
Public Const CONNECTION_ORAOLEDB_ORACLE = "ORAOLEDB.ORACLE" '"ORAOLEDB.ORACLE.1"
Public Const CONNECTION_SQLOLEDB = "SQLOLEDB"

'oledb provider - all
Public Const CONNECTION_PROVIDER = "PROVIDER"
'db file and path for access, tnsname for oracle, server for SQL Server
Public Const CONNECTION_DATASOURCE = "DATA SOURCE"
'database for SQL server only
Public Const CONNECTION_DATABASE = "DATABASE"
'user id for oracle and SQL server
Public Const CONNECTION_USERID = "USER ID"
'password for SQL server and Oracle
Public Const CONNECTION_PASSWORD = "PASSWORD"
'db password for access
Public Const CONNECTION_JETOLEDB_DATABASEPASSWORD = "JET OLEDB:DATABASE PASSWORD"

'connectiontype
Public Type udtConnection
    Provider As String
    Datasource As String
    Database As String
    UserId As String
    Password As String
End Type
    
'----------------------------------------------------------------------------------------'
Public Function Connection_AsType(sCon As String) As udtConnection
'----------------------------------------------------------------------------------------'

        Connection_AsType.Provider = Connection_Property(CONNECTION_PROVIDER, sCon)
        Connection_AsType.Datasource = Connection_Property(CONNECTION_DATASOURCE, sCon)
        Connection_AsType.Database = Connection_Property(CONNECTION_DATABASE, sCon)
        Connection_AsType.UserId = Connection_Property(CONNECTION_USERID, sCon)
        Connection_AsType.Password = Connection_Property(CONNECTION_PASSWORD, sCon)
        
End Function


Public Function Connection_Property(ByVal sProperty As String, ByVal sConnection As String) As String
'----------------------------------------------------------------------------------------'
' Return connection property value.
' Returns "" if no matching property.
'----------------------------------------------------------------------------------------'
Dim vProperties As Variant
Dim i As Long
Dim lStart As Long
Dim sCurrentProperty As String
Dim sResult As String
    
    sResult = ""
    vProperties = Split(sConnection, ";")
    For i = 0 To UBound(vProperties) - 1
        'current property is trimmed capitalised string in front of equals sign
        sCurrentProperty = UCase(Trim(Split(vProperties(i), "=")(0)))
        If sCurrentProperty = sProperty Then
            sResult = Trim(Split(vProperties(i), "=")(1))
            Exit For
        End If
    Next
    
    Connection_Property = sResult
        
End Function


Public Function Connection_String(sProvider As String, sDataSource As String, Optional sDatabase As String, Optional sUserId As String, Optional sPassword As String) As String
Dim sCon As String

    sCon = CONNECTION_PROVIDER & "=" & sProvider & ";" & CONNECTION_DATASOURCE & "=" & sDataSource & ";"
    Select Case UCase(sProvider)
    Case CONNECTION_MSJET_OLEDB_351, CONNECTION_MSJET_OLEDB_40
        sCon = sCon & CONNECTION_JETOLEDB_DATABASEPASSWORD & "=" & sPassword & ";"
    Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
        sCon = sCon & CONNECTION_USERID & "=" & sUserId & ";"
        sCon = sCon & CONNECTION_PASSWORD & "=" & sPassword & ";"
    Case CONNECTION_SQLOLEDB
        sCon = sCon & CONNECTION_DATABASE & "=" & sDatabase & ";"
        sCon = sCon & CONNECTION_USERID & "=" & sUserId & ";"
        sCon = sCon & CONNECTION_PASSWORD & "=" & sPassword & ";"
    Case Else
        'provider not found - raise error
        Err.Raise 1000 Or vbObjectError, "QueryServer.Init", "Unrecognised provider - " & sProvider
    End Select

    Connection_String = sCon
    
End Function
