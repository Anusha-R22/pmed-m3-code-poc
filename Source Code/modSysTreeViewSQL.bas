Attribute VB_Name = "modSysTreeViewSQL"
'----------------------------------------------------------------------------------------'
'   Copyright:  Inferfrmformd Ltd. 2003-2004. All Rights Reserved
'   File:       modSysTreeViewSQL.bas
'   Author:     Richard Meinesz, October 2002
'   Purpose:    Utilities for creating TreevIew in System Management
'----------------------------------------------------------------------------------------'
'REVISIONS:
'REM 10/02/04 - Changed routine NewDBAlias
'TA 19/1/2006 - routine to return max password retries value
' Mo 21/9/2007  Bug 2927, Overflow error when creating treeview of database studies & sites.
'               DISTINCT added to StudySiteWithRoles
'----------------------------------------------------------------------------------------'

Option Explicit

'---------------------------------------------------------------------
Public Function Databases() As Variant
'---------------------------------------------------------------------
'REM 03/10/02
'Returns an array of all the databases that the Systems administarator has access to
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsDatabases As ADODB.Recordset
Dim vDatabases As Variant

    On Error GoTo ErrLabel

    'get all databases
    sSQL = "SELECT DatabaseCode, ServerName, DatabaseType" _
         & " FROM Databases"
    Set rsDatabases = New ADODB.Recordset
    rsDatabases.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsDatabases.RecordCount > 0 Then
        vDatabases = rsDatabases.GetRows
    Else
        vDatabases = Null
    End If
    
    Databases = vDatabases
    
    rsDatabases.Close
    Set rsDatabases = Nothing
    
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "SQL.Databases"
End Function

'---------------------------------------------------------------------
Public Function StudySiteWithRoles(conMACRO As ADODB.Connection) As Variant
'---------------------------------------------------------------------
'REM 14/10/02
'Returns an array of studies or sites that appear in the UserRole and Trial Site table
'i.e. studies and sites that have assigned UserRoles
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsStudiesSites As ADODB.Recordset
Dim vStudiesSites As Variant

    On Error GoTo ErrLabel
    
    'Mo 21/9/2007  Bug 2927, DISTINCT added to following SQL statement
    sSQL = "SELECT DISTINCT Clinicaltrial.ClinicalTrialId, ClinicalTrial.ClinicalTrialName, ClinicalTrial.StatusId, Site.Site, Site.SiteStatus" _
         & " From UserRole, TrialSite, ClinicalTrial, Site" _
         & " WHERE (TrialSite.TrialSite = UserRole.SiteCode OR UserRole.SiteCode = '" & ALL_SITES & "')" _
         & " AND (ClinicalTrial.ClinicalTrialName = UserRole.StudyCode OR UserRole.StudyCode = '" & ALL_STUDIES & "')" _
         & " AND ClinicalTrial.ClinicalTrialId = TrialSite.ClinicalTrialId" _
         & " AND Site.Site = Trialsite.trialSite"

    Set rsStudiesSites = New ADODB.Recordset
    rsStudiesSites.Open sSQL, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsStudiesSites.RecordCount > 0 Then
        vStudiesSites = rsStudiesSites.GetRows
    Else
        vStudiesSites = Null
    End If
    
    StudySiteWithRoles = vStudiesSites
    
    rsStudiesSites.Close
    Set rsStudiesSites = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "SQL.StudyOrSitesWithRoles"
End Function

'---------------------------------------------------------------------
Public Function StudySiteWithoutRoles(conMACRO As ADODB.Connection, nDatabaseType As Integer) As Variant
'---------------------------------------------------------------------
'REM 15/10/02
'Returns an array of Study/Site combinations that appear in the TrialSite table but not the UserRole table
'i.e. Study/site combinations that have been distributed but have no roles assigned
'---------------------------------------------------------------------
'REVISIONS:
'REM 08/10/03 - Changed SQL to improve performace
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsStSiUserRole As ADODB.Recordset
Dim rsStSiTrialSite As ADODB.Recordset
Dim vStudySite As Variant
Dim i As Integer
Dim j As Integer
Dim sTsStudy As String
Dim sTsSite As String
Dim sUrStudy As String
Dim sUrSite As String

    On Error GoTo ErrLabel
    
    'get all Study/Site combinations that appear in the TrialSite table but not the UserRole table
    Select Case nDatabaseType
    Case MACRODatabaseType.Oracle80
        sSQL = " SELECT Clinicaltrial.ClinicalTrialId, ClinicalTrial.ClinicalTrialName, ClinicalTrial.StatusId, " _
             & " Site.Site , Site.SiteStatus " _
             & " FROM TrialSite, Site, ClinicalTrial " _
             & " WHERE TrialSite.TrialSite = Site.Site " _
             & " AND TrialSite.ClinicalTrialId = ClinicalTrial.ClinicalTrialId " _
             & " MINUS " _
             & " SELECT DISTINCT Clinicaltrial.ClinicalTrialId, ClinicalTrial.ClinicalTrialName, ClinicalTrial.StatusId, " _
             & " Site.Site , Site.SiteStatus " _
             & " FROM UserRole, Site, ClinicalTrial " _
             & " WHERE (Site.Site = UserRole.SiteCode OR UserRole.SiteCode = 'AllSites') " _
             & " AND (ClinicalTrial.ClinicalTrialName = UserRole.StudyCode OR UserRole.StudyCode = 'AllStudies') " _
             & " AND ClinicalTrial.ClinicalTrialId <> 0 "

    Case MACRODatabaseType.sqlserver
        sSQL = "SELECT DISTINCT Clinicaltrial.ClinicalTrialId, ClinicalTrial.ClinicalTrialName, " _
            & " ClinicalTrial.statusId , Site.Site, Site.SiteStatus " _
            & " FROM TrialSite, Site, ClinicalTrial " _
            & " WHERE TrialSite.TrialSite = Site.Site " _
            & " AND TrialSite.ClinicalTrialId = ClinicalTrial.ClinicalTrialId " _
            & " AND ClinicalTrial.ClinicalTrialName + '|' + Site.Site " _
            & " NOT IN (SELECT ClinicalTrial.ClinicalTrialName + '|' + Site.Site " _
            & " FROM UserRole, Site, ClinicalTrial " _
            & " WHERE  (Site.Site = UserRole.SiteCode " _
            & " OR UserRole.SiteCode = 'AllSites') " _
            & " AND (ClinicalTrial.ClinicalTrialName = UserRole.StudyCode " _
            & " OR UserRole.StudyCode = 'AllStudies') " _
            & " AND ClinicalTrial.ClinicalTrialId <> 0) "
    End Select
    
    Set rsStSiTrialSite = New ADODB.Recordset
    rsStSiTrialSite.Open sSQL, conMACRO, adOpenKeyset, adLockPessimistic, adCmdText
    
    'return an array
    If rsStSiTrialSite.RecordCount > 0 Then
        vStudySite = rsStSiTrialSite.GetRows
    Else
        vStudySite = Null
    End If

    StudySiteWithoutRoles = vStudySite

    rsStSiTrialSite.Close
    Set rsStSiTrialSite = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "SQL.StudySiteWithoutRoles"
End Function

'---------------------------------------------------------------------
Public Function StudySiteNotDist(conMACRO As ADODB.Connection, nDatabaseType As Integer) As Variant
'---------------------------------------------------------------------
'REM 16/10/02
'Returns an array of Study/Site combinations that appear in the UserRole table but not in the trialSite table
'i.e. Study/Site combinations that have user roles assigned but have not been distributed
'---------------------------------------------------------------------
'REVISIONS:
'REM 08/10/03 - Changed SQL to improve performace
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsStSiUserRole As ADODB.Recordset
Dim rsStSiTrialSite As ADODB.Recordset
Dim vStudySite As Variant
Dim i As Integer
Dim j As Integer
Dim sTsStudy As String
Dim sTsSite As String
Dim sUrStudy As String
Dim sUrSite As String

    On Error GoTo ErrLabel
    
    'get all Study/Site combinations that appear in the UserRole table but not in the trialSite table
    Select Case nDatabaseType
    Case MACRODatabaseType.Oracle80
    
        sSQL = "SELECT DISTINCT Clinicaltrial.ClinicalTrialId, ClinicalTrial.ClinicalTrialName, ClinicalTrial.StatusId, " _
            & " Site.Site , Site.SiteStatus " _
            & " FROM UserRole, Site, ClinicalTrial " _
            & " WHERE (Site.Site = UserRole.SiteCode OR UserRole.SiteCode = 'AllSites') " _
            & " AND (ClinicalTrial.ClinicalTrialName = UserRole.StudyCode OR UserRole.StudyCode = 'AllStudies') " _
            & " AND ClinicalTrial.ClinicalTrialId <> 0 " _
            & " MINUS " _
            & " SELECT Clinicaltrial.ClinicalTrialId, ClinicalTrial.ClinicalTrialName, ClinicalTrial.StatusId, " _
            & " Site.Site , Site.SiteStatus " _
            & " FROM TrialSite, Site, ClinicalTrial " _
            & " WHERE TrialSite.TrialSite = Site.Site " _
            & " AND TrialSite.ClinicalTrialId = ClinicalTrial.ClinicalTrialId "
    
    
    Case MACRODatabaseType.sqlserver
    
        sSQL = "SELECT DISTINCT Clinicaltrial.ClinicalTrialId, ClinicalTrial.ClinicalTrialName, " _
            & " ClinicalTrial.statusId , Site.Site, Site.SiteStatus " _
            & " FROM UserRole, Site, ClinicalTrial " _
            & " WHERE  (Site.Site = UserRole.SiteCode " _
            & " OR UserRole.SiteCode = 'AllSites') " _
            & " AND (ClinicalTrial.ClinicalTrialName = UserRole.StudyCode " _
            & " OR UserRole.StudyCode = 'AllStudies') " _
            & " AND ClinicalTrial.ClinicalTrialId <> 0 " _
            & " AND ClinicalTrial.ClinicalTrialName + '|' + Site.Site " _
            & " NOT IN (SELECT ClinicalTrial.ClinicalTrialName + '|' + Site.Site " _
            & " FROM TrialSite, Site, ClinicalTrial " _
            & " WHERE TrialSite.TrialSite = Site.Site " _
            & " AND TrialSite.ClinicalTrialId = ClinicalTrial.ClinicalTrialId) "
        
    End Select
    
    Set rsStSiUserRole = New ADODB.Recordset
    rsStSiUserRole.Open sSQL, conMACRO, adOpenKeyset, adLockPessimistic, adCmdText
    
    'return as an array
    If rsStSiUserRole.RecordCount > 0 Then
        vStudySite = rsStSiUserRole.GetRows
    Else
        vStudySite = Null
    End If
    
    StudySiteNotDist = vStudySite
    
    rsStSiUserRole.Close
    Set rsStSiUserRole = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysTreeViewSQL.StudySiteNotDist"
End Function


'---------------------------------------------------------------------
Public Function Users() As Variant
'---------------------------------------------------------------------
'REM 07/10/02
'Returns an array of all the users in the Security Database
'TA 19/01/2006 - include failedattempts
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsUsers As ADODB.Recordset
Dim vUsers As Variant

    On Error GoTo ErrLabel

    sSQL = "SELECT UserName, UserNameFull, Enabled, SysAdmin, failedattempts" _
        & " FROM MACROUser" _
        & " ORDER BY UserName"
    Set rsUsers = New ADODB.Recordset
    rsUsers.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsUsers.RecordCount > 0 Then
        vUsers = rsUsers.GetRows
    Else
        vUsers = Null
    End If
    
    Users = vUsers
    
    rsUsers.Close
    Set rsUsers = Nothing
    
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "SQL.Users"
End Function

'---------------------------------------------------------------------
Public Function Roles() As Variant
'---------------------------------------------------------------------
'REM 07/10/02
'Returns all the roles in the security database
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsRoles As ADODB.Recordset
Dim vRoles As Variant

    On Error GoTo ErrLabel
    
    sSQL = "SELECT * FROM Role"
    Set rsRoles = New ADODB.Recordset
    rsRoles.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsRoles.RecordCount > 0 Then
        vRoles = rsRoles.GetRows
    Else
        vRoles = Null
    End If
    
    Roles = vRoles
    
    rsRoles.Close
    Set rsRoles = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "SQL.Roles"
End Function

'---------------------------------------------------------------------
Public Function Studies(conMACRO As ADODB.Connection) As Variant
'---------------------------------------------------------------------
'REM 08/11/02
'Returns all the studies from the ClinicalTrial table
'Added extra return parameters as routine that uses this needs 5 return parameters
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsStudies As ADODB.Recordset
Dim vStudies As Variant

    On Error GoTo ErrLabel
    
    
    sSQL = "SELECT Clinicaltrial.ClinicalTrialId, ClinicalTrial.ClinicalTrialName, ClinicalTrial.StatusId,Clinicaltrial.ClinicalTrialId, ClinicalTrial.ClinicalTrialName " _
        & " FROM ClinicalTrial" _
        & " WHERE ClinicalTrialId <> 0"
    Set rsStudies = New ADODB.Recordset
    rsStudies.Open sSQL, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsStudies.RecordCount > 0 Then
        vStudies = rsStudies.GetRows
    Else
        vStudies = Null
    End If
    
    Studies = vStudies
    
    rsStudies.Close
    Set rsStudies = Nothing
    
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "SQL.Studies"
End Function

'---------------------------------------------------------------------
Public Function Sites(conMACRO As ADODB.Connection) As Variant
'---------------------------------------------------------------------
'REM 08/11/02
'Returns all sites from the site table
'Added extra return parameters as routine that uses this needs 5 return parameters
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsSites As ADODB.Recordset
Dim vSites As Variant

    On Error GoTo ErrLabel
    
    sSQL = "SELECT Site.SiteStatus, Site.SiteStatus,Site.SiteStatus,Site.Site, Site.SiteStatus" _
        & " FROM Site"
    Set rsSites = New ADODB.Recordset
    rsSites.Open sSQL, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsSites.RecordCount > 0 Then
        vSites = rsSites.GetRows
    Else
        vSites = Null
    End If
    
    Sites = vSites
    
    rsSites.Close
    Set rsSites = Nothing
    
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "SQL.Sites"
End Function

'--------------------------------------------------------------------------------
Public Sub RegisterMACRODatabase(sCon As String, sDatabaseAlias As String, nDatabaseType As Integer, sServer As String, _
                                 sNameofDatabase As String, sDatabasePswd As String, sDatabaseUser As String)
'--------------------------------------------------------------------------------
'REM 01/04/03
'Register a MACRO database in the security database
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim sHTMLPath As String
Dim sSHTMLPATH As String
Dim sEncryptedDatabasePswd As String
Dim sEncryptedDatabaseUser As String

    On Error GoTo ErrLabel

   'ASH 15/04/2002 Added new parameter to frmSetHTMLFolder
        Call frmSetHTMLFolder.HTMLPath(sHTMLPath, sSHTMLPATH)
        
        'if the user canceled out of setting HTML path, just insert the defaults
        If sHTMLPath = "" Then
            sHTMLPath = gsAppPath & "HTML\"
            sSHTMLPATH = gsAppPath & "HTML\"
        End If
        
        'encrypt the database password and user
        If sDatabasePswd = "" Then
            sEncryptedDatabasePswd = "null"
        Else
            
            sEncryptedDatabasePswd = "'" & EncryptString(sDatabasePswd) & "'"
        End If
        
        If sDatabaseUser = "" Then
            sEncryptedDatabaseUser = "null"
        Else
            sEncryptedDatabaseUser = "'" & EncryptString(sDatabaseUser) & "'"
        End If
        
        'ASH 15/04/2002 to get value for new field for SQL string
        sSQL = "INSERT INTO Databases (DatabaseCode,DatabaseType,ServerName,NameOfDatabase,DatabaseUser,DatabasePassword,HTMLLocation,SecureHTMLLocation)"
        'ASH 12/9/2002 Added Alias to get new database name
        sSQL = sSQL & "VALUES ('" & sDatabaseAlias & "', " & nDatabaseType & ",'" & sServer & "','" & sNameofDatabase & "'," & sEncryptedDatabaseUser & "," & sEncryptedDatabasePswd & ", '" & sHTMLPath & "','" & sSHTMLPATH & "'" & ")"
        SecurityADODBConnection.Execute sSQL, adCmdText
        
        'REM 11/02/03 - restore User Database links
        Call RestoreUserDatabaseLinks(sCon, sDatabaseAlias)
        
        DialogInformation ("The database has been registered successfully")

        'Refresh the tree view first
        Call frmMenu.RefereshTreeView
        
        'Refresh the list of databases on frmDatabases
        Call frmMenu.RefereshDatabaseInfoForm

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modNewDatabaseandTables.RegisterMACRODatabase"
End Sub

'----------------------------------------------------------------------------------------------
Public Function NewDBAlias(sDatabase As String) As String
'----------------------------------------------------------------------------------------------
'REM 20/05/03
'Routine returns new DB alias and checks to see if it is unique
'REM 10/02/04 - Changed to use frmInputBox so could limit DBAlias to 15 chars
'----------------------------------------------------------------------------------------------
Dim bValid As Boolean
Dim sDBAlias As String
Dim bDialog As Boolean

    bValid = False
    bDialog = False
    Do Until bValid
        If bDialog Then DialogWarning ("A database with the alias " & sDBAlias & " has already been registered!")
        'get the Database Alias from the database name, limit it to 15 chars
        sDBAlias = Left(sDatabase, 15)
        'REM 10/02/04 - Changed to use frmInputBox so could limit DBAlias to 15 chars
        Call frmInputBox.Display("Database DB Alias", "MACRO Alias", sDBAlias, False, False, True, , 15)
        'sDBAlias = InputBox("MACRO DB Alias", "MACRO Alias", sDatabase)
        If sDBAlias = "" Then    ' if cancel, then return control to user
            bValid = True
        Else
            'if DB does not exist then true
            bValid = Not DoesDatabaseExist(sDBAlias)
            'Display the warning dialog to tell user they have entered an existing DBAlias
            bDialog = Not bValid
        End If
    Loop
    
    NewDBAlias = sDBAlias

End Function

'----------------------------------------------------------------------------------------------
Private Sub RestoreUserDatabaseLinks(sCon As String, sDatabaseAlias As String)
'----------------------------------------------------------------------------------------------
'REM 11/02/03
'Restores the UserDatabase links using the databases UserRole table
' RS 13/06/2003: BUG 1822, prevent insertion of duplicate records (PK violation)
'----------------------------------------------------------------------------------------------
Dim conMACRO As ADODB.Connection
Dim sSQL As String
Dim sSQL1 As String
Dim rsVersion As ADODB.Recordset
Dim sVersion As String
Dim rsUsers As ADODB.Recordset
Dim rsCopyUserDatabases As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrHandler

    Set conMACRO = New ADODB.Connection
    conMACRO.Open sCon
    conMACRO.CursorLocation = adUseClient

    'check to see the user is not registering a pre-MACRO 3.0 database, if so exit the routine and don't restore User Database links
    sSQL = "SELECT * FROM MACROControl"
    Set rsVersion = New ADODB.Recordset
    rsVersion.Open sSQL, conMACRO, adOpenKeyset, adLockPessimistic, adCmdText
    sVersion = rsVersion!MACROVersion
    rsVersion.Close
    Set rsVersion = Nothing
    If sVersion <> "3.0" Then Exit Sub

    'get recordset of Users for a specific site from the UserRole table
    sSQL = "SELECT DISTINCT UserName FROM UserRole "
    Set rsUsers = New ADODB.Recordset
    rsUsers.Open sSQL, conMACRO, adOpenKeyset, adLockPessimistic, adCmdText

    If rsUsers.RecordCount > 0 Then

' RS 13/06/2003: Do separate selection for each user found (below)
'        'creates site recordset to contain records to be restored
'        sSQL1 = "SELECT * FROM UserDatabase "
'        sSQL1 = sSQL1 & " WHERE 0 = 1"
        Set rsCopyUserDatabases = New ADODB.Recordset

' RS 13/06/2003: This never happens as already in a >0 condition
'        'checks if records exist
'        If rsUsers.RecordCount <= 0 Then
'            Screen.MousePointer = vbNormal
'            Exit Sub
'        End If

        'move to first record in recordset
        rsUsers.MoveFirst

        'begin record restoration
        For i = 1 To rsUsers.RecordCount
        
            ' Open UserDatabase: Select for current user/databasealias
            sSQL1 = "SELECT * FROM UserDatabase "
            sSQL1 = sSQL1 & " WHERE username = '" & rsUsers("username") & "' AND databasecode = '" & sDatabaseAlias & "'"
            rsCopyUserDatabases.Open sSQL1, SecurityADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            
            ' Only insert if combination does not yet exists
            If rsCopyUserDatabases.EOF Then
                ' Can insert the user/database record
                rsCopyUserDatabases.AddNew
                rsCopyUserDatabases.Fields(0).Value = rsUsers.Fields(0).Value
                rsCopyUserDatabases.Fields(1).Value = sDatabaseAlias
                rsCopyUserDatabases.Update
            Else
                ' User/Database already present: No action
            End If
            rsCopyUserDatabases.Close
            rsUsers.MoveNext
        Next

        rsUsers.Close
        Set rsUsers = Nothing
        Set rsCopyUserDatabases = Nothing
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modNewDatabaseandTables.RestoreUserDatabaseLinks"
End Sub

'---------------------------------------------------------------------
Public Function PasswordRetries() As Long
'---------------------------------------------------------------------
'TA 19/01/2006
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsPasswords As ADODB.Recordset
    
On Error GoTo ErrHandler

    sSQL = " Select PasswordRetries FROM MACROPassword"
    Set rsPasswords = New ADODB.Recordset
    rsPasswords.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
    PasswordRetries = rsPasswords.Fields(0).Value
    rsPasswords.Close
Exit Function

ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSysTreeViewSQL.PasswordRetries"

End Function
