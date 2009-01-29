Attribute VB_Name = "basSecurity"
  '--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       basSecurity.bas
'   Author:     Andrew Newbigging, June 1997
'   Purpose:    Checks user security authorisation.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   1    Andrewn    4/06/97
'   2    Andrewn    27/11/97
'   3    Andrewn    2/12/97
'   4    PN         14/09/99 Upgrade from DAO to ADO and updated code to conform
'                            to VB standards doc version 1.0
'   PN  04/10/99    Amended gblnUserExists(), gblnRoleExists(), gblnUserRoleExists()
'                   code reading RecordCount() property
'   Mo 13/12/99     Id's from integer to Long
'   WillC 21/2/00   Added gblnUserRoleOnDatabaseExists SR 2853 One role per user per database.
'   WillC 26/6/00   Took out dead code at start of version 2.1
'   Mo 18/7/2001  Changes stemming from field Password in table MacroUser being changed to
'               UserPassword (stems from the swith to Jet 4.0)
'   Mo 30/8/01  Moved DoesDatabaseExist to modADODBConnections
'   ZA 18/09/02 updates List_users.js when a new user is added into the database
'   ic 25/11/2002 changed call to CreateUsersList() in InsertNewUser() function
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------

Option Explicit
Option Base 0
Option Compare Binary
'Public nErrNum As Long

' NCJ 2 May 00
' Result of trying to open a MACRO Access database
Public Enum eIsMacroDB
    OK = 0
    InvalidPassword = 1
    NotMacro = 2
End Enum

Public Enum RestoreSite
    SecurityAndMACRO = 1
    MACROOnly = 2
    ExitRestore = 3
End Enum

'--------------------------------------------------------------------------------
Public Function gblnUserExists(sUserName As String) As Boolean
'--------------------------------------------------------------------------------
' determine if a user exists in the system
' REM 21/01/03 - check user name in Uppercase in Oracle as it is case sensitive
'--------------------------------------------------------------------------------
Dim rsUserRecord As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler

    ' PN 15/09/99 -  change tablename User to MacroUser
    'Mo Morris 20/9/01 Db Audit (UserCode to UserName)
'    sSQL = "SELECT * FROM MacroUser WHERE UserName = '" & sUserCode & "'"
'    Set rsUserRecord = New ADODB.Recordset

    'REM 21/01/03 - check user name in Uppercase in Oracle as it is case sensitive
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.oracle80
        sSQL = "SELECT * FROM MACROUser WHERE upper(UserName) = upper('" & sUserName & "')"
    Case Else
        sSQL = "SELECT * FROM MACROUser WHERE UserName = '" & sUserName & "'"
    End Select
    
    ' PN 04/10/99
    ' use a keyset cursor to get the correct count of records
    Set rsUserRecord = New ADODB.Recordset
    rsUserRecord.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
    
    If rsUserRecord.RecordCount > 0 Then
        gblnUserExists = True
    Else
        gblnUserExists = False
    End If
    rsUserRecord.Close
    Set rsUserRecord = Nothing
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gblnUserExists", "Security.bas")
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
Public Function gdsUserDetails(sUserName As String, sPassword As String) As ADODB.Recordset
'--------------------------------------------------------------------------------
' read user details
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsUserDetails As ADODB.Recordset

    On Error GoTo ErrHandler

    ' PN 15/09/99 -  change tablename User to MacroUser
    'changed Mo 18/7/01 field changed from Password to UserPassword
    'Mo Morris 20/9/01 Db Audit (UserName to UserNameFull)
    sSQL = "SELECT * FROM MacroUser WHERE UserNameFull = '" & sUserName _
           & "' AND UserPassword = '" & sPassword & "'"
    Set rsUserDetails = New ADODB.Recordset
    rsUserDetails.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, , adCmdText
    Set gdsUserDetails = rsUserDetails
    Set rsUserDetails = Nothing
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsUserDetails", "Security.bas")
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
Public Function gdsUserFunction(sUserName As String) As ADODB.Recordset
'--------------------------------------------------------------------------------
' return the user functions
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsUserFunction As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM UserFunction WHERE UserName = '" & sUserName & "'"
    Set rsUserFunction = New ADODB.Recordset
    rsUserFunction.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, , adCmdText
    Set gdsUserFunction = rsUserFunction
    Set rsUserFunction = Nothing
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsUserFunction", "Security.bas")
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
Public Sub gdsAddFunctionToUser(sUserName As String, sFunctionName As String)
'--------------------------------------------------------------------------------
' insert the user function
'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "INSERT INTO UserFunction " _
            & "VALUES ('" & sUserName & "','" _
            & sFunctionName & "')"

    SecurityADODBConnection.Execute sSQL
 
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsAddFunctionToUser", "Security.bas")
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
Public Sub gdsRemoveFunctionFromUser(sUserName As String, sFunctionName As String)
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
Dim SecurityADODBConnection As ADODB.Connection
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "DELETE FROM UserFunction " _
            & "WHERE UserName = '" & sUserName & "' " _
            & "AND Function = '" & sFunctionName & "'"
    
    SecurityADODBConnection.Execute sSQL, , adCmdText
 
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsRemoveFunctionFromUser", "Security.bas")
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
Public Function gdsUserList() As ADODB.Recordset
'--------------------------------------------------------------------------------
' return a list of users
'Mo Morris 4/2/00
'Local recordset removed. Now the recordset that is built is the one that is returned

'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'WillC 3/2/00 - Removed Order by Could be causing an error and not crucial.
    ' PN 15/09/99 -  change tablename User to MacroUser
    sSQL = "SELECT *  FROM MacroUser" ' ORDER BY UserName"
    'changed Mo Morris 4/2/00
    Set gdsUserList = New ADODB.Recordset
    gdsUserList.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, , adCmdText
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsUserList", "Security.bas")
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
Public Function gdsFunctionList() As ADODB.Recordset
'--------------------------------------------------------------------------------
' return a list of functions
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsFunctionList As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT *  FROM Function"
    Set rsFunctionList = New ADODB.Recordset
    rsFunctionList.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, , adCmdText
    Set gdsFunctionList = rsFunctionList
    Set rsFunctionList = Nothing
 
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsFunctionList", "Security.bas")
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
Public Function gdsUser(sUserCode As String) As ADODB.Recordset
'--------------------------------------------------------------------------------
' return the specified user details
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsUser As ADODB.Recordset

    On Error GoTo ErrHandler

    ' PN 15/09/99 -  change tablename User to MacroUser
    'Mo Morris 20/9/01 Db Audit (UserCode to UserName)
    sSQL = "SELECT *  FROM MacroUser WHERE UserName = '" & sUserCode & "'"
    Set rsUser = New ADODB.Recordset
    rsUser.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, , adCmdText
    Set gdsUser = rsUser
    Set rsUser = Nothing
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsUser", "Security.bas")
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
Public Function gdsDatabaseList() As ADODB.Recordset
'--------------------------------------------------------------------------------
' return a list of databases
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsDatabaseList As ADODB.Recordset

    On Error GoTo ErrHandler

    'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode)
    sSQL = "SELECT Distinct DatabaseCode FROM Databases"
    Set rsDatabaseList = New ADODB.Recordset
    rsDatabaseList.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, , adCmdText
    Set gdsDatabaseList = rsDatabaseList
    Set rsDatabaseList = Nothing
 
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsDatabaseList", "Security.bas")
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
Public Sub gdsUserEnabled(sUserCode As String, nEnabled As Integer)
'--------------------------------------------------------------------------------
' set the user enabled flag
'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    ' PN 15/09/99 -  change tablename User to MacroUser
    'Mo Morris 20/9/01 Db Audit (UserCode to UserName)
    sSQL = "UPDATE MacroUser " _
    & " SET Enabled = " & nEnabled _
    & " WHERE UserName = '" & sUserCode & "'"
        
    SecurityADODBConnection.Execute sSQL, , adCmdText
 
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsUserEnabled", "Security.bas")
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
Public Sub AddDatabaseToUser(sUserCode As String, sDatabaseDescription As String)
'--------------------------------------------------------------------------------
' add the user to a database
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsUserDatabase As ADODB.Recordset

    On Error GoTo ErrHandler
    
    'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode, UserCode to UserName)
    sSQL = "Select * from Userdatabase Where UserName = '" & sUserCode & "'" _
    & " AND DatabaseCode = '" & sDatabaseDescription & "'"
    
    Set rsUserDatabase = New ADODB.Recordset
    rsUserDatabase.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
    If rsUserDatabase.RecordCount > 0 Then
        Exit Sub
    Else
        sSQL = "INSERT INTO UserDatabase " _
                & "VALUES ('" & sUserCode & "','" _
                & sDatabaseDescription & "')"
        SecurityADODBConnection.Execute sSQL, , adCmdText
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "AddDatabaseToUser", "Security.bas")
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
Public Sub RemoveDatabaseFromUser(sUserCode As String, sDatabaseDescription As String)
'--------------------------------------------------------------------------------
' remove the user from a database
'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    'Changed by Mo Morris 18/1/00. '*' removed from Delete SQL statement
    'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode, UserCode to UserName)
    sSQL = "DELETE FROM UserDatabase " _
            & "WHERE UserName = '" & sUserCode & "' " _
            & "AND DatabaseCode = '" & sDatabaseDescription & "'"
    
    SecurityADODBConnection.Execute sSQL, adCmdText
 
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RemoveDatabaseFromUser", "Security.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'Mo Morris 20/9/01 Db Audit
'The sub is not called from anywhere so I have commented it out
''--------------------------------------------------------------------------------
'Public Function gdsUserDatabase(sUsername As String) As ADODB.Recordset
''--------------------------------------------------------------------------------
'' return a list of all databases that the user belongs to
''--------------------------------------------------------------------------------
'Dim sSQL As String
'Dim rsUserDatabases As ADODB.Recordset
'
'    On Error GoTo ErrHandler
'
'    sSQL = "SELECT * FROM UserDatabase WHERE UserName = '" & sUsername & "'"
'    Set rsUserDatabases = New ADODB.Recordset
'    rsUserDatabases.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, , adCmdText
'    Set gdsUserDatabase = rsUserDatabases
'    Set rsUserDatabases = Nothing
'
'Exit Function
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsUserDatabase", "Security.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Function

'--------------------------------------------------------------------------------
Public Sub gdsUserLogin(sUserName As String)
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    ' PN 15/09/99 -  change tablename User to MacroUser
    'WillC 4/2/00 - Changed to SQLStandardNow tocope with local settings
    'Mo Morris 20/9/01 Db Audit (UserName to UserNameFull)
    sSQL = "UPDATE MacroUser " _
    & " SET LastLogin = " & SQLStandardNow _
    & " WHERE UserNameFull = '" & sUserName & "'"
        
    SecurityADODBConnection.Execute sSQL, adOpenKeyset, adCmdText
 
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsUserLogin", "Security.bas")
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
Public Sub InsertNewUser(sUserCode As String, sUserName As String, sFirstLogin As String, _
                         sPassword As String)
'--------------------------------------------------------------------------------
' WillC 4/2/00 Changed to cope with local settings
'revisions
'ic 25/11/2002  changed call to CreateUserList()
'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    ' PN 15/09/99 -  change tablename User to MacroUser
    'Will 6/10/99 - Added firstlogin and lastlogin for password expiration
    'changed Mo Morris 5/1/00, single quotes removed from either side of CDbl(dtFirstLogin)
    'changed Mo 18/7/01 field changed from Password to UserPassword
    'Mo Morris 20/9/01 Db Audit (UserCode to UserName, UserName to UserNameFull)
    sSQL = "INSERT INTO MacroUser " _
        & "(UserName,UserNameFull,FirstLogin,UserPassword,Enabled)" _
        & " VALUES ('" & sUserCode & "','" & sUserName & "'," & sFirstLogin & ",'" _
        & sPassword & "',1)"

    SecurityADODBConnection.Execute sSQL, , adCmdText
    
    'ZA 18/09/2002 - update List_user.js now
    Call CreateUsersList(goUser)
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "InsertNewUser", "Security.bas")
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
Public Function gblnHasRoleFunction(tmpRoleCode As String, tmpFunctionCode As String) As Boolean
'--------------------------------------------------------------------------------
' check to see if a Role has a function already if it doesnt do the insert if it has
' then do nothing.
'--------------------------------------------------------------------------------
Dim rsRoleFunction As ADODB.Recordset
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM RoleFunction WHERE RoleCODe = '" & tmpRoleCode & "'"
    sSQL = sSQL & " AND FunctionCode = '" & tmpFunctionCode & "'"
    
    Set rsRoleFunction = New ADODB.Recordset
    rsRoleFunction.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
    
    If rsRoleFunction.RecordCount > 0 Then
        gblnHasRoleFunction = True
    Else
        gblnHasRoleFunction = False
         AddFunctionToRole tmpRoleCode, tmpFunctionCode
    End If
    
    Set rsRoleFunction = Nothing
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gblnRoleExists", "Security.bas")
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
Public Function gblnRoleExists(sRoleCode As String) As Boolean
'--------------------------------------------------------------------------------
' check to see if a Role exists in the database already
'--------------------------------------------------------------------------------
Dim rsRoleCode As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
     
   sSQL = "SELECT RoleCode FROM Role  WHERE RoleCode = '" & sRoleCode & "'"
     
     
    ' PN 04/10/99
    ' use a keyset cursor to get the correct count of records
    Set rsRoleCode = New ADODB.Recordset
    rsRoleCode.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
    
    If rsRoleCode.RecordCount > 0 Then
        gblnRoleExists = True
    Else
        gblnRoleExists = False
    End If
    Set rsRoleCode = Nothing
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gblnRoleExists", "Security.bas")
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
Public Function gblnUserRoleExists(sUserCode As String, sRoleCode As String, _
                                   sDatabaseDescription As String) As Boolean ', nTrialId As Long, _
                                   sTrialSite As String) As Boolean
  
'--------------------------------------------------------------------------------
' See if a User has this Role already
'--------------------------------------------------------------------------------
Dim rsUserRole As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode, UserCode to UserName)
    sSQL = "SELECT * FROM UserRole" _
        & " WHERE UserName = '" & sUserCode & "'" _
        & " AND RoleCode ='" & sRoleCode & "'" _
        & " AND DatabaseCode ='" & sDatabaseDescription & "'" '_
'     & " AND ClinicalTrialId = " & nTrialId _
'     & " AND TrialSite ='" & sTrialSite & "'"
                       
    ' PN 04/10/99
    ' use a keyset cursor to get the correct count of records
    Set rsUserRole = New ADODB.Recordset
    rsUserRole.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
        
    If rsUserRole.RecordCount > 0 Then
        gblnUserRoleExists = True
    Else
        gblnUserRoleExists = False
    End If
    
  Set rsUserRole = Nothing
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gblnUserRoleExists", "Security.bas")
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
Public Function gblnUserRoleOnDatabaseExists(sUserCode As String, sDatabaseDescription As String) As Boolean
'--------------------------------------------------------------------------------
' See if a User has a Role On a database already
' WillC 21/2/00 SR 2853 One role per user per database.
'--------------------------------------------------------------------------------
Dim rsUserRole As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode, UserCode to UserName)
    sSQL = "SELECT * FROM UserRole"
    sSQL = sSQL & " WHERE UserName = '" & sUserCode & "'"
    sSQL = sSQL & " AND DatabaseCode ='" & sDatabaseDescription & "'"
                    
    ' PN 04/10/99
    ' use a keyset cursor to get the correct count of records
    Set rsUserRole = New ADODB.Recordset
    rsUserRole.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
        
    If rsUserRole.RecordCount > 0 Then
        gblnUserRoleOnDatabaseExists = True
    Else
        gblnUserRoleOnDatabaseExists = False
    End If
    
  Set rsUserRole = Nothing
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gblnUserRoleOnDatabaseExists", "Security.bas")
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
Public Sub AddFunctionToRole(sRoleCode As String, tmpFunctionCode As String)
'--------------------------------------------------------------------------------
' Allows the insertion of new roles into the RoleFunction table
'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "INSERT INTO RoleFunction " _
            & " VALUES ('" & sRoleCode & " ','" & tmpFunctionCode & "')"
            
    SecurityADODBConnection.Execute sSQL, , adCmdText
    
Exit Sub
ErrHandler:
    Select Case Err.Number
        Case 3320
             MsgBox "You must select a UserCode", vbInformation
        Case Else
            Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "AddFunctionToRole", "Security.bas")
                  Case OnErrorAction.Ignore
                      Resume Next
                  Case OnErrorAction.Retry
                      Resume
                  Case OnErrorAction.QuitMACRO
                      Call ExitMACRO
                      Call MACROEnd
             End Select
    End Select


End Sub

'--------------------------------------------------------------------------------
Public Sub gdsUpdateAllTrialsAllSites(sUserCode As String, sRoleCode As String, _
                                        sDatabaseDescription As String)
'--------------------------------------------------------------------------------
' WillC 21/2/00 Added to update a users single role per database. SR2853.
'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
        'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode, UserCode to UserName)
        sSQL = "UPDATE UserRole SET " _
        & " UserName = '" & sUserCode & "'" _
        & " ,RoleCode = '" & sRoleCode & "'" _
        & " ,DatabaseCode = '" & sDatabaseDescription & "'" _
        & " WHERE UserName = '" & sUserCode & "'" _
        & " AND DatabaseCode= '" & sDatabaseDescription & "'"
        
    SecurityADODBConnection.Execute sSQL, , adCmdText
    
Exit Sub

ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsUpdateAllTrialsAllSites", "Security.bas")
          Case OnErrorAction.Ignore
              Resume Next
          Case OnErrorAction.Retry
              Resume
          Case OnErrorAction.QuitMACRO
              Call ExitMACRO
              Call MACROEnd
     End Select

End Sub

'Mo Morris 20/9/01 Db Audit
'The sub is not called from anywhere so I have commented it out
''--------------------------------------------------------------------------------
'Public Sub gdsAddFunctionsToRole(tmpRoleCode As String, tmpFunctionCode As String, _
'                                sDatabaseDescription As String)
''--------------------------------------------------------------------------------
'' Update the functions that an existing role comprises of
''--------------------------------------------------------------------------------
'Dim sSQL As String
'Dim sUserCode As String
'
'    On Error GoTo ErrHandler
'
'    sSQL = "UPDATE  RoleFunction SET " _
'    & " RoleCode = " & "'" & tmpRoleCode & "'," _
'    & " FunctionCode = " & "'" & tmpFunctionCode & "'," _
'    & "AND DatabaseDescription = '" & sDatabaseDescription & "'" _
'    & " WHERE UserCode = '" & sUserCode & "'"
'
'    SecurityADODBConnection.Execute sSQL, , adCmdText
'
'Exit Sub
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsAddFunctionsToRole", "Security.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Sub

'--------------------------------------------------------------------------------
Public Function GetSiteList() As ADODB.Recordset
'--------------------------------------------------------------------------------
'Select all the available sites
'--------------------------------------------------------------------------------
Dim rsSiteList As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "SELECT DISTINCT TrialSite FROM TrialSite WHERE ClinicalTrialId > 0 "
    Set rsSiteList = New ADODB.Recordset
    rsSiteList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, , adCmdText
    Set GetSiteList = rsSiteList
    Set rsSiteList = Nothing
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetSiteList", "Security.bas")
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
Public Function GetTrialList() As ADODB.Recordset
'--------------------------------------------------------------------------------
'Select all the available Trials
'--------------------------------------------------------------------------------
Dim rsTrialList As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "SELECT DISTINCT ClinicalTrialId FROM ClinicalTrial WHERE ClinicalTrialId > 0"
    Set rsTrialList = New ADODB.Recordset
    rsTrialList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, , adCmdText
    Set GetTrialList = rsTrialList
    Set rsTrialList = Nothing
                                       
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetTrialList", "Security.bas")
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
Public Function gdsGetFnCode(sRoleCode As String) As ADODB.Recordset
'--------------------------------------------------------------------------------
' When you Select a Role this gives the corresponding
' Functions for that Role
'--------------------------------------------------------------------------------
Dim rsGetFnCode As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler

    'TA 18/11/2002: name changed from function to MACROFunction
    sSQL = "SELECT DISTINCT MACROFunction FROM MACROFunction, RoleFunction" _
         & " WHERE MACROFunction.FunctionCode = RoleFunction.FunctionCode" _
         & " AND RoleFunction.RoleCode = '" & sRoleCode & "'"
    
    Set rsGetFnCode = New ADODB.Recordset
    rsGetFnCode.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, , adCmdText
    Set gdsGetFnCode = rsGetFnCode
    Set rsGetFnCode = Nothing

 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsGetFnCode", "Security.bas")
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
Public Function gdsRoleList() As ADODB.Recordset
'--------------------------------------------------------------------------------
'Get a list of all the rolecodes
'--------------------------------------------------------------------------------
Dim rsRoleList As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler

    'return all roles if user is sysadmin, else only return non-SysAdmin roles
    If goUser.SysAdmin Then
        sSQL = "SELECT DISTINCT RoleCode FROM Role "
    Else
        sSQL = "SELECT DISTINCT RoleCode FROM Role " _
            & " WHERE SysAdmin = 0"
    End If
    
    Set rsRoleList = New ADODB.Recordset
    rsRoleList.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adCmdText
    Set gdsRoleList = rsRoleList
    Set rsRoleList = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsRoleList", "Security.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'Mo Morris 20/9/01 Db Audit
'The sub is not called from anywhere so I have commented it out
''--------------------------------------------------------------------------------
'Public Sub gdsNewRoleForUser(sUserCode As String, sRoleCode As String, sDatabaseDescription As String)
''--------------------------------------------------------------------------------
''Add a new role for an Existing User
''--------------------------------------------------------------------------------
'Dim sSQL As String
'
'    On Error GoTo ErrHandler
'
'    ''Mo Morris 20/9/01 Db Audit (ClinicalTrialId and TrialSite removed)
'    sSQL = "INSERT INTO UserRole " _
'           & "(UserCode,RoleCode,DatabaseDescription,AllTrials,AllSites)" _
'           & " VALUES ('" & sUserCode & "','" & sRoleCode & "','" _
'           & sDatabaseDescription & "'," & "1,1)"
'
'    SecurityADODBConnection.Execute sSQL, , adCmdText
'
'Exit Sub
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsNewRoleForUser", "Security.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Sub
 

'--------------------------------------------------------------------------------
Public Function gdsGetTempFunctionCode(tmpRoleFunction As String)
'--------------------------------------------------------------------------------
' Get the FunctionCode from the FunctionName selected so we can then use the
' FunctionCode so we can complete the Insert in AddFunctionToRole
'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "SELECT FunctionCode FROM Function" _
            & " WHERE FunctionName = '" & tmpRoleFunction & "'"
            
    SecurityADODBConnection.Execute sSQL, , adCmdText
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsGetTempFunctionCode", "Security.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'Mo Morris 20/9/01 Db Audit
'The following sub contained an erroneous SQL statement.
'The sub is not called from anywhere so I have commented it out
''--------------------------------------------------------------------------------
'Public Sub gdsAddRoleToUser(sUserCode As String, sDatabaseDescription As String)
''--------------------------------------------------------------------------------
'' Add another role to a user
''--------------------------------------------------------------------------------
'Dim sSQL As String
'
'    On Error GoTo ErrHandler
'
'    sSQL = " INSERT INTO UserRole " & _
'            "VALUES (' & sUserCode & " & "','" _
'            & sDatabaseDescription & "')"
'
'    SecurityADODBConnection.Execute sSQL, , adCmdText
'
'Exit Sub
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsAddRoleToUser", "Security.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Sub

'--------------------------------------------------------------------------------
Public Sub gdsRemoveRoleFromUser(sUserCode As String, sDatabaseDescription As String)
'--------------------------------------------------------------------------------
'Remove a role from a user
'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode, UserCode to UserName)
    sSQL = " DELETE FROM UserRole " _
           & " WHERE UserName = '" & sUserCode & "'" _
           & " AND DatabaseCode = '" & sDatabaseDescription & "'"
        
    SecurityADODBConnection.Execute sSQL, , adCmdText
 
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsRemoveRoleFromUser", "Security.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'Mo Morris 20/9/01 Db Audit
'The following sub contained an erroneous SQL statement.
'The sub is not called from anywhere so I have commented it out
''--------------------------------------------------------------------------------
'Public Sub gdsAddSiteToUser(sUserCode As String, sDatabaseDescription As String)
''--------------------------------------------------------------------------------
''Adds another Role to a user when a site is bing added
''--------------------------------------------------------------------------------
'Dim sSQL As String
'
'    On Error GoTo ErrHandler
'
'    sSQL = "INSERT INTO UserRole " _
'        & "VALUES ('" & sUserCode & "','" _
'        & sDatabaseDescription & "')"
'
'    SecurityADODBConnection.Execute sSQL, , adCmdText
'
'Exit Sub
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsAddSiteToUser", "Security.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Sub

'Mo Morris 20/9/01 Db Audit
'The sub is not called from anywhere so I have commented it out
''--------------------------------------------------------------------------------
'Public Sub gdsRemoveSiteFromUser(sUserCode As String, sDatabaseDescription As String)
''--------------------------------------------------------------------------------
'' removes the userRole when a Site is being taken from a user
''--------------------------------------------------------------------------------
'Dim sSQL As String
'
'    On Error GoTo ErrHandler
'
'    sSQL = "DELETE FROM UserRole " _
'            & " WHERE UserCode = '" & sUserCode & "'" _
'            & " AND DatabaseDescription = ' " & sDatabaseDescription & "'"
'
'    SecurityADODBConnection.Execute sSQL, , adCmdText
'
'Exit Sub
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsRemoveSiteFromUser", "Security.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Sub

'Mo Morris 20/9/01 Db Audit
'The following sub contained an erroneous SQL statement.
'The sub is not called from anywhere so I have commented it out
''--------------------------------------------------------------------------------
'Public Sub gdsAddTrialToUser(sUserCode As String, sDatabaseDescription As String)
''--------------------------------------------------------------------------------
'''Adds another Role to a user when a trial is bing added
''--------------------------------------------------------------------------------
'Dim sSQL As String
'
'    On Error GoTo ErrHandler
'
'    sSQL = "INSERT INTO UserRole " _
'            & "VALUES ('" & sUserCode & "','" _
'            & sDatabaseDescription & " ')"
'
'    SecurityADODBConnection.Execute sSQL, , adCmdText
'
'Exit Sub
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsAddTrialToUser", "Security.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Sub

'Mo Morris 20/9/01 Db Audit
'The sub is not called from anywhere so I have commented it out
''--------------------------------------------------------------------------------
'Public Sub gdsRemoveTrialFromUser(sUserCode As String, sDatabaseDescription As String)
''--------------------------------------------------------------------------------
'' removes the userRole when a trial is being taken from a user
''--------------------------------------------------------------------------------
'Dim sSQL As String
'
'    On Error GoTo ErrHandler
'
'        sSQL = "DELETE FROm UserRole " _
'                & " WHERE UserCode = '" & sUserCode & "'" _
'                & "AND DatabaseDescription = '" & sDatabaseDescription & "'"
'
'    SecurityADODBConnection.Execute sSQL, , adCmdText
'
'Exit Sub
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsRemoveTrialFromUser", "Security.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Sub


'--------------------------------------------------------------------------------
Public Sub gdsInsertAllTrialsAllSites(sUserCode As String, sRoleCode As String, _
                                                sDatabaseDescription As String)
'--------------------------------------------------------------------------------
'Set the AllTrials & AllSites to 1 when thats the option chosen by the user..
'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'Mo Morris 20/9/01 Db Audit (ClinicalTrialId and TrialSite removed,
    'DatabaseDescription to DatabaseCode, UserCode to UserName)
    sSQL = "INSERT INTO UserRole " _
           & "(UserName,RoleCode,DatabaseCode,AllTrials,AllSites)" _
           & " VALUES ('" & sUserCode & "','" & sRoleCode & "','" _
           & sDatabaseDescription & "'," & " 1,1 )"
 
    SecurityADODBConnection.Execute sSQL, , adCmdText
     
Exit Sub
ErrHandler:
    Select Case Err.Number
        Case 3315
            MsgBox "You must select a UserCode", vbInformation
        Case Else
            Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsInsertAllTrialsAllSites", "Security.bas")
                  Case OnErrorAction.Ignore
                      Resume Next
                  Case OnErrorAction.Retry
                      Resume
                  Case OnErrorAction.QuitMACRO
                      Call ExitMACRO
                      Call MACROEnd
             End Select
    End Select

End Sub

'Mo Morris 20/9/01 Db Audit
'The sub is not called from anywhere so I have commented it out
''--------------------------------------------------------------------------------
'Public Sub gdsInsertAllTrialsASite(sUserCode As String, sRoleCode As String, _
'                                sTrialSite As String, sDatabaseDescription As String)
''--------------------------------------------------------------------------------
''allow a user permissions to all trials at a site
''--------------------------------------------------------------------------------
'Dim sSQL As String
'
'    On Error GoTo ErrHandler
'
'    sSQL = "INSERT INTO UserRole " _
'           & "(UserCode,RoleCode,ClinicalTrialId,TrialSite,DatabaseDescription,AllTrials,AllSites)" _
'           & " VALUES ('" & sUserCode & "','" & sRoleCode & "'," & "'0'" & "" _
'           & ",'" & sTrialSite & "','" & sDatabaseDescription & "'," & " 1,0 )"
'
'    SecurityADODBConnection.Execute sSQL, , adCmdText
'
'Exit Sub
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsInsertAllTrialsASite", "Security.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Sub

'Mo Morris 20/9/01 Db Audit
'The sub is not called from anywhere so I have commented it out
''--------------------------------------------------------------------------------
'Public Sub gdsInsertAllSitesATrial(sUserCode As String, sRoleCode As String, _
'                                    nTrialId As Long, sDatabaseDescription As String)
''--------------------------------------------------------------------------------
''Set  AllSites to 1 so that this UserRole is set for all sites for a particular trial
''--------------------------------------------------------------------------------
'Dim sSQL As String
'
'    On Error GoTo ErrHandler
'
'   sSQL = "INSERT INTO UserRole " _
'           & "(UserCode,RoleCode,ClinicalTrialId,TrialSite,DatabaseDescription,AllTrials,AllSites)" _
'           & " VALUES ('" & sUserCode & "','" & sRoleCode & "','" & nTrialId & "" _
'           & "','" & "AllSites" & "','" & sDatabaseDescription & "'," & " 0,1 )"
'
'    SecurityADODBConnection.Execute sSQL, , adCmdText
'
'Exit Sub
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsInsertAllSitesATrial", "Security.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Sub

'Mo Morris 20/9/01 Db Audit
'The sub is not called from anywhere so I have commented it out
''--------------------------------------------------------------------------------
'Public Sub gdsInsertASiteATrial(sUserCode As String, sRoleCode As String, _
'                        nTrialId As Long, sTrialSite As String, sDatabaseDescription As String)
''--------------------------------------------------------------------------------
''Set the Users Trial/Sites options to those chosen by the user
''--------------------------------------------------------------------------------
'Dim sSQL As String
'
'    On Error GoTo ErrHandler
'
'   sSQL = "INSERT INTO UserRole " _
'           & "(UserCode,RoleCode,ClinicalTrialId,TrialSite,DatabaseDescription,AllTrials,AllSites)" _
'           & " VALUES ('" & sUserCode & "','" & sRoleCode & "','" & nTrialId & "" _
'           & "','" & sTrialSite & "','" & sDatabaseDescription & "'," & " 0,0 )"
'
'   SecurityADODBConnection.Execute sSQL, , adCmdText
'
'Exit Sub
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "gdsInsertASiteATrial", "Security.bas")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Sub

'--------------------------------------------------------------------------------
Public Sub InsertRole(sRoleCode As String, sRoleDescription As String)
                               
'--------------------------------------------------------------------------------
' Add a new role into the database
'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "INSERT INTO Role " _
        & "(RoleCode,RoleDescription,Enabled)" _
        & " VALUES ('" & sRoleCode & "','" & sRoleDescription & "',1)"
 
    SecurityADODBConnection.Execute sSQL, , adCmdText
 
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "InsertRole", "Security.bas")
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
Public Sub RemoveFunctionFromRole(tmpRoleCode As String, tmpRoleFunction As String)
'--------------------------------------------------------------------------------
'Called from the EditUserRole to remove functions from a role when the role
'has been changed
'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
     
    sSQL = " DELETE FROM RoleFunction " _
        & " WHERE RoleCode = '" & tmpRoleCode & "'" _
        & " AND FunctionCode = '" & tmpRoleFunction & "'"

    SecurityADODBConnection.Execute sSQL, , adCmdText
 
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RemoveFunctionFromRole", "Security.bas")
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
Public Function UserProfile(sUserCode As String) As ADODB.Recordset
'--------------------------------------------------------------------------------
'  Get all of a specific Users Role details
'  this Rs is OpenKeyset as we want to do a recordcount which means going forward
'  and back through the recordset, The lock is optimistic to allow updating
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsUserProfile As ADODB.Recordset

    On Error GoTo ErrHandler

    'Mo Morris 20/9/01 Db Audit (UserCode to UserName)
    sSQL = "SELECT Distinct *  FROM UserDatabase WHERE UserName = '" & sUserCode & "'"
    Set rsUserProfile = New ADODB.Recordset
    rsUserProfile.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    Set UserProfile = rsUserProfile
    Set rsUserProfile = Nothing
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UserProfile", "Security")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'----------------------------------------------------------------------------
Public Function IsMACRODatabase(sConnection As String, lSecAndMACRO As Long) As Boolean
'----------------------------------------------------------------------------
' REM 12/02/03
' Had to re-write this routine as no longer use Access
' Checks to see if a selected SQL Server or Oracle database is a MACRO DB
'----------------------------------------------------------------------------
Dim sSQL As String
Dim conMACRO As ADODB.Connection
Dim rsTest As ADODB.Recordset
Dim lMACRORecordCount As Long

    On Error GoTo ErrHandler
    
    IsMACRODatabase = False
        
    'connect to the database
    Set conMACRO = New ADODB.Connection
    conMACRO.Open sConnection
    conMACRO.CursorLocation = adUseClient
    
    ' Test for MACRO DB by looking for ClinicalTrialID = 0
    sSQL = "SELECT ClinicalTrialId FROM ClinicalTrial WHERE ClinicalTrialId = 0"
    Set rsTest = New ADODB.Recordset
    rsTest.Open sSQL, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    ' Store the record count before closing down the connection
    lMACRORecordCount = rsTest.RecordCount
    
    'security and MACRO database
    If lSecAndMACRO = RestoreSite.SecurityAndMACRO Then
        If (lMACRORecordCount > 0) And (DoesTableExist(conMACRO, "MACROUser") = True) Then
            IsMACRODatabase = True
        Else
            IsMACRODatabase = False
        End If
    'just a MACRO database
    ElseIf lSecAndMACRO = RestoreSite.MACROOnly Then
        ' Did we find a ClinicalTrial record?
        If lMACRORecordCount > 0 Then
            IsMACRODatabase = True
        Else
            IsMACRODatabase = False
        End If
    End If
    
    rsTest.Close
    Set rsTest = Nothing
    conMACRO.Close
    Set conMACRO = Nothing

Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "IsMACRODatabase", "Security.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

''----------------------------------------------------------------------------
'Public Function GetMACRODBPassword(ByVal sPath As String, _
'                                    ByRef sPassword As String) As Boolean
''----------------------------------------------------------------------------
'' NCJ 2 May 00
'' Get the password for the specified MACRO (Access) DB
'' Return TRUE if password is OK and MACRO DB is OK (and sPassword is correct password)
'' Otherwise return FALSE
'' Relevant error messages will have already been given
''----------------------------------------------------------------------------
'Dim bPasswordOK As Boolean
'Dim sPwd As String
'Dim nAttempts As Integer
'
'    On Error GoTo ErrHandler
'
'    bPasswordOK = False
'    sPwd = Trim(sPassword)
'    nAttempts = 0
'
'    Do While (Not bPasswordOK) And nAttempts < 2
'        Select Case IsAMACROAccessDatabase(sPath, sPwd)
'        Case eIsMacroDB.InvalidPassword
'            ' Ask them for a new password
'            If nAttempts = 0 Then
'                sPwd = Trim(InputBox("Please enter the password for this database"))
'            End If
'            nAttempts = nAttempts + 1
'            ' Unfortunately we can't tell the difference between
'            ' them entering an empty string and clicking Cancel - but never mind
'        Case eIsMacroDB.NotMacro
'            Exit Do
'        Case eIsMacroDB.OK
'            bPasswordOK = True
'            sPassword = sPwd
'            Exit Do
'        End Select
'    Loop
'    GetMACRODBPassword = bPasswordOK
'
'Exit Function
'ErrHandler:
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
'                                        "GetMACRODBPassword", "Security")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Function
'
