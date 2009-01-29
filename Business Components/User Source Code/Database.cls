VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       Database.cls
'   Author:     Richard Meinesz, September 2002
'   Purpose:
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
' Revisions:
' NCJ 27 Nov 03 - Added class Terminate
' DPH 24/03/2004 - Clean up collections in class_terminate
' DPH 06/03/07 - bug 2656. Making database selection case-insensitive in Load()
'----------------------------------------------------------------------------------------'

Option Explicit


Public Enum eMACRODatabaseType
    mdtAccess = 0
    mdtSQLServer = 1
'    mdtSQLServer70 = 2
    mdtOracle80 = 3
End Enum

Private msDatabaseCode As String
Private msDatabaseLocation As String
Private mnDatabaseType As Integer
Private msServerName As String
Private msNameOfDatabase As String
Private msDatabaseUser As String
Private msDatabasePassword As String
Private msConnection As String
Private mcolUserRoles As Collection
'
Private msHTMLLocation As String
Private msSecureHTMLLocation As String

'----------------------------------------------------------------------------------------'
Public Property Get DatabaseCode() As String
'----------------------------------------------------------------------------------------'

    DatabaseCode = msDatabaseCode

End Property

'----------------------------------------------------------------------------------------'
Public Property Get HTMLLocation() As String
'----------------------------------------------------------------------------------------'

    HTMLLocation = msHTMLLocation

End Property

'----------------------------------------------------------------------------------------'
Public Property Get SecureHTMLLocation() As String
'----------------------------------------------------------------------------------------'

    SecureHTMLLocation = msSecureHTMLLocation

End Property

'----------------------------------------------------------------------------------------'
Public Property Get DatabaseLocation() As String
'----------------------------------------------------------------------------------------'

    DatabaseLocation = msDatabaseLocation

End Property

'----------------------------------------------------------------------------------------'
Public Property Get DatabaseType() As eMACRODatabaseType
'----------------------------------------------------------------------------------------'

    DatabaseType = mnDatabaseType

End Property

'----------------------------------------------------------------------------------------'
Public Property Get ServerName() As String
'----------------------------------------------------------------------------------------'

    ServerName = msServerName

End Property

'----------------------------------------------------------------------------------------'
Public Property Get NameOfDatabase() As String
'----------------------------------------------------------------------------------------'

    NameOfDatabase = msNameOfDatabase

End Property

'----------------------------------------------------------------------------------------'
Public Property Get DatabaseUser() As String
'----------------------------------------------------------------------------------------'

    DatabaseUser = msDatabaseUser

End Property

'----------------------------------------------------------------------------------------'
Public Property Get DatabasePassword() As String
'----------------------------------------------------------------------------------------'

    DatabasePassword = msDatabasePassword

End Property

'----------------------------------------------------------------------------------------'
Public Property Get ConnectionString() As String
'----------------------------------------------------------------------------------------'

    ConnectionString = msConnection

End Property

'----------------------------------------------------------------------------------------'
Public Property Get UserRoles() As Collection
'----------------------------------------------------------------------------------------'

    Set UserRoles = mcolUserRoles

End Property

'----------------------------------------------------------------------------------------'
Public Function UpdateHTMLLoaction(oSecCon As ADODB.Connection, sDatabaseCode As String, sHTMLLocation As String, sSecureHTMLLocation As String) As Boolean
'----------------------------------------------------------------------------------------'
'Update the HTML location of the current database
'----------------------------------------------------------------------------------------'
Dim sSQL As String
    
    On Error GoTo ErrHandler

    sSQL = "UPDATE DATABASES SET HTMLLOCATION = '" & sHTMLLocation & "'," _
        & " SECUREHTMLLOCATION = '" & sSecureHTMLLocation & "'" _
        & " WHERE DATABASECODE = '" & sDatabaseCode & "'"
    oSecCon.Execute sSQL, adOpenKeyset, adCmdText

    msHTMLLocation = sHTMLLocation

    msSecureHTMLLocation = sSecureHTMLLocation

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "Database.UpdateHTMLLoaction"
End Function

'----------------------------------------------------------------------------------------'
Public Function Load(oSecCon As Connection, sUserName As String, sDatabase As String, _
                     sDefaultHTMLLocation As String, bLoadUserRoles As Boolean, _
                     ByRef sMessage As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 19/09/02
'Loads all the database properties
' DPH 06/03/07 - bug 2656. Making database selection case-insensitive
'----------------------------------------------------------------------------------------'
Dim sPath As String
Dim rsDatabase As ADODB.Recordset
Dim sSQL As String
Dim sString As String
Dim bLoad As Boolean
    
    On Error GoTo ErrHandler

    sSQL = "SELECT HTMLLocation, DatabaseLocation, DatabaseType, ServerName, NameOfDatabase, " _
         & " DatabaseUser, DatabasePassword, SecureHTMLLocation" _
         & " FROM Databases "
         
    ' DPH 06/03/07 - bug 2656. Making database selection case-insensitive
    If (InStr(1, Connection_Property(CONNECTION_PROVIDER, oSecCon.ConnectionString), CONNECTION_MSDAORA, vbTextCompare) > 0) _
        Or (InStr(1, Connection_Property(CONNECTION_PROVIDER, oSecCon.ConnectionString), CONNECTION_ORAOLEDB_ORACLE, vbTextCompare) > 0) Then
        ' oracle - in lowercase
        sSQL = sSQL & " WHERE NLS_LOWER(DatabaseCode) = '" & LCase(sDatabase) & "'"
    Else
        ' sql server - remains the same
        sSQL = sSQL & " WHERE DatabaseCode = '" & sDatabase & "'"
    End If
    
    Set rsDatabase = New ADODB.Recordset
    rsDatabase.Open sSQL, oSecCon, adOpenForwardOnly, adLockReadOnly, adCmdText

    If rsDatabase.RecordCount = 0 Then
        sMessage = "Database not found"
        Load = False
    Else

        msDatabaseLocation = RemoveNull(rsDatabase!DatabaseLocation)
        mnDatabaseType = rsDatabase!DatabaseType
        msDatabaseCode = sDatabase
        
        If RemoveNull(rsDatabase!DatabaseUser) = "" Then
            msDatabaseUser = ""
        Else
            msDatabaseUser = DecryptString(rsDatabase!DatabaseUser)
        End If
        
        If RemoveNull(rsDatabase!DatabasePassword) = "" Then
            msDatabasePassword = ""
        Else
            msDatabasePassword = DecryptString(rsDatabase!DatabasePassword)
        End If
        
        msServerName = RemoveNull(rsDatabase!ServerName)
        msNameOfDatabase = RemoveNull(rsDatabase!NameOfDatabase)
    
        If IsNull(rsDatabase!HTMLLocation) Then
            msHTMLLocation = sDefaultHTMLLocation
        Else
            msHTMLLocation = rsDatabase!HTMLLocation
        End If
    
        'TA add back slash if one does not exist
        If msHTMLLocation <> "" Then
            If Right(msHTMLLocation, 1) <> "\" Then
                msHTMLLocation = msHTMLLocation & "\"
            End If
        End If
    
        If IsNull(rsDatabase!SecureHTMLLocation) Then
            ' default to HTML location
            msSecureHTMLLocation = msHTMLLocation
        Else
            msSecureHTMLLocation = rsDatabase!SecureHTMLLocation
        End If
    
        rsDatabase.Close
        Set rsDatabase = Nothing
    
        'create connection string for selected database
        Select Case DatabaseType
        Case eMACRODatabaseType.mdtSQLServer
                'SQL SERVER OLE DB NATIVE PROVIDER
            msConnection = Connection_String(CONNECTION_SQLOLEDB, ServerName, NameOfDatabase, _
                    DatabaseUser, DatabasePassword)
        Case eMACRODatabaseType.mdtOracle80
                'Oracle OLE DB NATIVE PROVIDER
            msConnection = Connection_String(CONNECTION_MSDAORA, NameOfDatabase, , _
                    DatabaseUser, DatabasePassword)
        End Select
        
        If bLoadUserRoles Then
            'Load all the user roles for the specific database
            Load = LoadDatabaseUserRoles(oSecCon, msConnection, sUserName, sDatabase, sMessage)
        Else
            sMessage = ""
            Load = True
        End If

    End If
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "Database.Load"
End Function

'----------------------------------------------------------------------------------------'
Friend Function LoadDatabaseUserRoles(oSecCon As Connection, sMACROCon As String, ByVal sUserName As String, _
                                ByVal sDatabaseCode As String, ByRef sMessage As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 19/09/02
'Get a collection of RoleCodes associated with a specific user and database from the
' UserRole table in the selected MACRO DB
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsRoleCodes As ADODB.Recordset
Dim colRoleCodes As Collection
Dim sRoleCode As String
Dim conMACROADODBConnection As ADODB.Connection
Dim oDatabase As Database

    On Error GoTo Errorlabel
    
    Set oDatabase = New Database
    'create connection for selected database
    Set conMACROADODBConnection = New Connection
    conMACROADODBConnection.Open sMACROCon
    conMACROADODBConnection.CursorLocation = adUseClient
    
    Call UpgradeData3_0from13to14_SPECIAL(oSecCon, conMACROADODBConnection)
    
    'get all role codes for user in the database
    sSQL = "SELECT DISTINCT RoleCode" _
         & " FROM UserRole" _
         & " WHERE UserName = '" & sUserName & "'"
    Set rsRoleCodes = New ADODB.Recordset
    rsRoleCodes.Open sSQL, conMACROADODBConnection ', adOpenKeyset, adLockReadOnly, adCmdText
    
    Set colRoleCodes = New Collection
    
    'place roles into a collection
    Do While Not rsRoleCodes.EOF
        sRoleCode = rsRoleCodes!RoleCode
        colRoleCodes.Add sRoleCode, sRoleCode
    rsRoleCodes.MoveNext
    Loop
    
    'return collection to the property
    Set mcolUserRoles = colRoleCodes
    
    sMessage = ""
    LoadDatabaseUserRoles = True
    
    Set colRoleCodes = Nothing
    rsRoleCodes.Close
    Set rsRoleCodes = Nothing
    conMACROADODBConnection.Close
    
Exit Function
Errorlabel:
    'Return the error message
    sMessage = Err.Description
    'also return the error message to the collection
    Set mcolUserRoles = New Collection
    mcolUserRoles.Add "Could not connect to database because " & sMessage, "Error"
    'set load to false
    LoadDatabaseUserRoles = False
End Function

'---------------------------------------------------------------------
Private Sub UpgradeData3_0from13to14_SPECIAL(oSecCon As Connection, conMACRO As Connection)
'---------------------------------------------------------------------
'REM 08/10/02
' Add UserRole table to the MACRO database from the Security Db
'this routine is called from basMainMacroModule in routine Main
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsMACROControl As ADODB.Recordset
Dim rsUserRole As ADODB.Recordset
Dim sUserNameDT As String
Dim sRoleCodeDT As String
Dim sStudySiteDT As String
Dim sInstallDT As String

    On Error GoTo ErrLabel
    
    Select Case DatabaseType
    Case mdtSQLServer
        sUserNameDT = "VARCHAR(20)"
        sRoleCodeDT = "VARCHAR(15)"
        sStudySiteDT = "VARCHAR(50)"
        sInstallDT = "SMALLINT"
    Case mdtOracle80
        sUserNameDT = "VARCHAR2(20)"
        sRoleCodeDT = "VARCHAR2(15)"
        sStudySiteDT = "VARCHAR2(50)"
        sInstallDT = "NUMBER(6)"
    End Select
    
    '***Add New UserRole Table***
    sSQL = "CREATE Table UserRole(UserName " & sUserNameDT & "," _
        & " RoleCode " & sRoleCodeDT & "," _
        & " StudyCode " & sStudySiteDT & "," _
        & " SiteCode " & sStudySiteDT & "," _
        & " TypeOfInstallation " & sInstallDT & "," _
        & " CONSTRAINT PKUserRole PRIMARY KEY (UserName,RoleCode,StudyCode,SiteCode,TypeOfInstallation))"
    conMACRO.Execute sSQL


    '*** Add New Table to the table MACROTable ***
    sSQL = "INSERT INTO MACROTable VALUES ('UserRole','',0,0,0)"
    conMACRO.Execute sSQL

    'get the user roles from the old security UserRole table that are for the current MACRO database
    sSQL = "SELECT UserName, RoleCode FROM UserRole" _
        & " WHERE DatabaseCode = '" & msNameOfDatabase & "'"
    Set rsUserRole = New ADODB.Recordset
    rsUserRole.Open sSQL, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText

    'add the roles to the new MACRO UserRole table
    Do While Not rsUserRole.EOF
        sSQL = "INSERT INTO UserRole VALUES ('" & rsUserRole!UserName & "','" & rsUserRole!RoleCode & "', 'AllStudies', 'AllSites', 0)"
        conMACRO.Execute sSQL

        rsUserRole.MoveNext
    Loop
    
Exit Sub
ErrLabel:
Exit Sub
'    Err.Raise Err.Number, , Err.Description & "|modUpgradeDatabases.UpgradeData3_0from13to14_SPECIAL"
End Sub


'----------------------------------------------------------------------
Private Function RemoveNull(vVariable As Variant) As String
'----------------------------------------------------------------------

    RemoveNull = ConvertFromNull(vVariable, vbString)

End Function

'----------------------------------------------------------------------
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------
'deserialisation code
'currently only db type is persisted
'----------------------------------------------------------------------
    mnDatabaseType = PropBag.ReadProperty("T")
    
End Sub

'----------------------------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------------------------
' NCJ 27 Nov 03 - Tidy up on terminate
' DPH 24/03/2004 - remove all from collection
'----------------------------------------------------------------------

    Call CollectionRemoveAll(mcolUserRoles)
    Set mcolUserRoles = Nothing
    
End Sub

'----------------------------------------------------------------------
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------
'serialisation code
'currently only db type is persisted
'----------------------------------------------------------------------

    PropBag.WriteProperty "T", mnDatabaseType
    
End Sub
