Attribute VB_Name = "modUIHTMLApplication"
'----------------------------------------------------------------------------------------'
'   File:       modUIHTML.bas
'   Copyright:  InferMed Ltd. 2003-2007. All Rights Reserved
'   Author:     i curtis 02/2003
'   Purpose:    functions returning html versions of MACRO pages (MENU)
'----------------------------------------------------------------------------------------'
' ic 21/05/2003 replaced html select control in GetAppMenuLh()
' ic 28/05/2003 added GetMessageHTML() function
' ic 28/05/2003 quick fix for z-order bug on loading menus in GetAppMenuLh()
' ic 29/05/2003 added extra call to resize event in GetAppMenuLh(). bug 1806
' ic 18/06/2003 changed string format for regular expression in GetAppMenuLh(), added registration
' ic 20/06/2003 added GetDatabaseChoice() and related functions
' ic 23/06/2003 added 'false' log parameter to GetDatabaseChoice(), bug 1706
' ic 23/06/2003 added bInit parameter for select lists to GetAppMenuLh()
' NCJ 25 Jun 03 - Bug 1830 - Changed permission on View Responded Discrepancies
' ic 26/06/2003 display values in local format, bug 1803
' DPH 02/07/2003 - Stop null values upsetting display in GetQuestionAudit
' ic 27/08/2003 moved GetErrorHTML to modUIHTML
' ic 02/09/2003 bug 1934, permission now required to change password in GetAppMenuLhTaskList()
' ic 22/09/2003 removed include for rfo.js - out-of-date reference causing iis 404 logging
' ic 14/10/2003 handle no quicklist permission in GetAppMenuLh()
' REM 05/12/03 - Added MACRO setting to GetAppMenuLhTaskList to enable and disable the use of the reset password permission
' ic 12/12/2003 added database check in GetDatabaseChoice() function
' NCJ 6 Jan 04 - Check for NULL CTC/NR values in GetQuestionAudit (CRM 390)
' ic 21/01/2004 added GMT difference to timestamp and database timestamp in GetQuestionAudit()
' ic 29/01/2004 changed copyright to 2004 in about screen
' ic 29/06/2004 added parameter checking, error handling
' NCJ 22 July 04 - In GetQuestionAudit deal with NULL time zones (Bug 2349)
' ic 09/11/2004 disable submit button after submitting in GetAppMenuLh()
' ic 05/05/2005 changed copyright to 2005 in GetAbout()
' ic 27/05/2005 changed warning to "W" and reject to "R" in GetRFO()
' ic 28/07/2005 added clinical coding
' NCJ 23 Jan 06 - Changed copyright to 2006 in GetAbout
' NCJ 25 Oct 06 - Changed GetDBLockAdmin to show eForm names and handle MUSD eForm locks
' dph 28/02/2007 bug 2882. added tracing to GetAppMenuLH
' ic 08/03/2007 issue 2602 full user names with apostrophes cause the side panel to crash - js encode them
' ic 16/04/2007 issue 2759, added the derivation expression
' NCJ 18 Apr 07 - Changed copyright to 2007 in GetAbout
' Mo 10/10/2007, Bug 2923, overflow error in GetAppMenuLh when user has access to a greater than integer number of subjects
' NCJ 18 Mar 08 - Changed copyright to 2008 in GetAbout
'----------------------------------------------------------------------------------------'

Option Explicit
Private Const mCCSwitch As String = "CLINICALCODING"

'----------------------------------------------------------------------------------------'
Public Function GetDatabaseChoice(ByVal sSDBCon As String, ByVal sUser As String, ByVal sPassword As String, _
ByRef sDatabase As Variant, ByRef sRole As Variant, ByVal sAppState As String) As String
'----------------------------------------------------------------------------------------'
'   ic 20/06/2003
'   function returns database choice page html if a choice exists, or a database code
'   and role code if there is a single option, or an error page if no options
'   revisions
'   ic 23/06/2003 added 'false' log parameter, bug 1706
'   ic 12/12/2003 added 'var sDatabase' line to javascript to check requested and chosen
'                 database match - so we can wipe appstate variable if necessary
'   ic 29/06/2004 added parameter checking, error handling
'----------------------------------------------------------------------------------------'
Dim oUser As MACROUser
Dim colDatabases As Collection
Dim colRoles As Collection
Dim oDatabase As Database
Dim sMessage As String
Dim vErrors As Variant
Dim vJSComm() As String
Dim sDatabases As String
Dim sRoles As String
Dim n As Integer
Dim bFoundDB As Boolean

    On Error GoTo CatchAllError
    
    Set oUser = New MACROUser
    'ic 23/06/2003 added 'false' log parameter, bug 1706
    Call oUser.Login(sSDBCon, sUser, sPassword, "", "MACRO Web Data Entry", "", True, "", "", False)
    Set colDatabases = oUser.AllUserDatabases(sSDBCon, sUser, "")
    
    If (colDatabases.Count = 0) Then
        'no databases
        Call Err.Raise(vbObjectError + 1, , "User does not have permission to access any MACRO databases")
    
    ElseIf (colDatabases.Count = 1) Then
        'single database
        If (IsDatabaseOK(sSDBCon, colDatabases.Item(1).DatabaseCode, sMessage)) Then
            'single database, connection verified
            Set colRoles = colDatabases.Item(1).UserRoles
            If (colRoles.Count = 0) Then
                'single database, no roles
                Call Err.Raise(vbObjectError + 1, , "User does not have roles on any MACRO databases")
                
            ElseIf (colRoles.Count = 1) Then
                'single database, one role
                sDatabase = colDatabases.Item(1).DatabaseCode
                sRole = colRoles.Item(1)
                Exit Function
                
            Else
                'single database, multiple roles: continue
            End If
        Else
            'single database, connection failed
            Call Err.Raise(vbObjectError + 1, , "Unable to connect to MACRO database [" _
            & colDatabases.Item(1).DatabaseCode & "]. " & ReplaceWithJSChars(sMessage))
    
        End If
    End If
    
    
    'if we get here there is a user choice to be made
    bFoundDB = False
    sDatabases = ""
    sRoles = ""
    For Each oDatabase In colDatabases
        If (IsDatabaseOK(sSDBCon, oDatabase.DatabaseCode, sMessage)) Then
            'connection verified
            'add database to database list
            sDatabases = sDatabases & oDatabase.DatabaseCode & gsDELIMITER1
            'compare requested db to verify requested db exists
            If sDatabase = oDatabase.DatabaseCode Then bFoundDB = True
            'add roles to role list
            Set colRoles = oDatabase.UserRoles
            sRoles = sRoles & gsDELIMITER2
            For n = 1 To colRoles.Count
                sRoles = sRoles & colRoles.Item(n) & gsDELIMITER1
            Next
            If (colRoles.Count > 0) Then
                'remove trailing delimiter
                sRoles = Left(sRoles, Len(sRoles) - 1)
            End If
        Else
            'connection failed. dont list database, notify user
            vErrors = AddToArray(vErrors, oDatabase.DatabaseCode, ReplaceWithJSChars(sMessage))
        End If
    Next
    'remove trailing delimiter
    sDatabases = Left(sDatabases, Len(sDatabases) - 1)
    'remove starting delimiter
    sRoles = Mid(sRoles, 2)
    'if requested db was not found, clear it
    If Not bFoundDB Then sDatabase = ""
    
    ReDim vJSComm(0)
    Call AddStringToVarArr(vJSComm, "<body onload='fnPageLoaded();'>" _
        & "<script language='javascript'>" & vbCrLf _
        & "function fnPageLoaded(){")
    
    If (Not IsEmpty(vErrors)) Then
        Call AddStringToVarArr(vJSComm, "alert('MACRO was unable to connect to one or more registered databases." _
            & "\nThese are listed below\n\n")
            
        For n = LBound(vErrors, 2) To UBound(vErrors, 2)
                Call AddStringToVarArr(vJSComm, vErrors(0, n) & " - " & vErrors(1, n) & "\n")
        Next
        
        Call AddStringToVarArr(vJSComm, "');" & vbCrLf)
    End If
    'ic 12/12/2003 added var sDatabase line
    Call AddStringToVarArr(vJSComm, "fnLoadDb(" & Chr(34) & sDatabase & Chr(34) & ");" & vbCrLf _
        & "document.all.db.focus();")
    Call AddStringToVarArr(vJSComm, "}")
    Call AddStringToVarArr(vJSComm, "var sDatabases=" & Chr(34) & sDatabases & Chr(34) & ";" & vbCrLf _
        & "var sRoles=" & Chr(34) & sRoles & Chr(34) & ";" & vbCrLf _
        & "var sDatabase=" & Chr(34) & sDatabase & Chr(34) & ";" & vbCrLf)
    Call AddStringToVarArr(vJSComm, "</script>")
    
    
    Call AddStringToVarArr(vJSComm, "<div><img height='100%' width='100%' src='../img/bg.jpg'></div>" _
        & "<div style='position:absolute; left:0; top:30%;'>" _
        & "<table width='100%' height='100%'><tr><td align='center' valign='middle'>" _
        & "<form name='Form1' method='get' action='SelectDatabase.asp'>" _
        & "<input type='hidden' name='app' value=" & Chr(34) & sAppState & Chr(34) & ">" _
        & "<table width='300' align='center' border='0'>" _
        & "<tr height='20'><td class='clsTableHeaderText'>&nbsp;Database</td></tr>" _
        & "<tr><td>" _
        & "<select style='width:140px;' class='clsSelectList' size='4' name='db' onchange='fnDbClick();'></select>" _
        & "</td></tr>" _
        & "<tr height='20'><td class='clsTableHeaderText'>&nbsp;Role</td></tr>" _
        & "<tr><td>" _
        & "<select style='width:140px;' class='clsSelectList' size='4' name='rl'></select>" _
        & "</td></tr>" _
        & "<tr height='10'><td>&nbsp;</td></tr>" _
        & "<tr><td align='right'>" _
        & "<input style='width:100px;' class='clsButton'  type='button' value='OK' name='btnSubmit' onmouseup='fnSubmit();'></td></tr>" _
        & "<tr height='10'><td>&nbsp;</td></tr>" _
        & "</table>" _
        & "</form>" _
        & "</td></tr></table>" _
        & "</div>")

    Call AddStringToVarArr(vJSComm, "</body>")
    Set oUser = Nothing
    Set colDatabases = Nothing
    Set colRoles = Nothing
    Set oDatabase = Nothing
    GetDatabaseChoice = Join(vJSComm, "")
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTMLApplication.GetDatabaseChoice")
End Function

'----------------------------------------------------------------------------------------'
Private Function IsDatabaseOK(ByVal sSDBCon As String, ByVal sDatabase As String, ByRef sMessage As String) As Boolean
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
Dim conMACROADODBConnection As ADODB.Connection
Dim sDbCon As String

    IsDatabaseOK = False
    On Error GoTo ConFailed
    
    If (Not CreateDatabaseCon(sSDBCon, sDatabase, sDbCon)) Then
        'couldnt create a database connection string, return error
        sMessage = sDbCon
        Exit Function
    End If
    
    'create connection to database
    Set conMACROADODBConnection = New Connection
    conMACROADODBConnection.Open sDbCon
    conMACROADODBConnection.Close
    Set conMACROADODBConnection = Nothing
    IsDatabaseOK = True
    Exit Function

ConFailed:
    'Return the error message
    sMessage = Err.Description
    'ensure that no connections are left behind
    On Error Resume Next
    Set conMACROADODBConnection = Nothing
End Function

'----------------------------------------------------------------------------------------'
Private Function CreateDatabaseCon(ByVal sSDBCon As String, ByVal sDatabase As String, _
    ByRef sDbCon As String) As Boolean
'----------------------------------------------------------------------------------------'
'   ic 20/06/2003
'   returns a database connection string
'----------------------------------------------------------------------------------------'
Dim conMACROADODBConnection As ADODB.Connection
Dim rsDatabase As ADODB.Recordset
Dim sSQL As String
Dim sServerName As String
Dim sDatabaseUser As String
Dim sNameOfDatabase As String
Dim sDatabasePassword As String
Dim nDatabaseType As Integer
    
    CreateDatabaseCon = False
    On Error GoTo CatchAllError

    'create connection to security database
    Set conMACROADODBConnection = New Connection
    conMACROADODBConnection.Open sSDBCon

    sSQL = "SELECT HTMLLocation, DatabaseLocation, DatabaseType, ServerName, NameOfDatabase, " _
         & " DatabaseUser, DatabasePassword, SecureHTMLLocation" _
         & " FROM Databases " _
         & " WHERE DatabaseCode = '" & sDatabase & "'"
    Set rsDatabase = New ADODB.Recordset
    rsDatabase.Open sSQL, conMACROADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    If rsDatabase.RecordCount = 0 Then
        sDbCon = "Database not found"

    Else

        sServerName = ConvertFromNull(rsDatabase!ServerName, vbString)
        If ConvertFromNull(rsDatabase!DatabaseUser, vbString) = "" Then
            sDatabaseUser = ""
        Else
            sDatabaseUser = DecryptString(rsDatabase!DatabaseUser)
        End If
        sNameOfDatabase = ConvertFromNull(rsDatabase!NameOfDatabase, vbString)
        If ConvertFromNull(rsDatabase!DatabasePassword, vbString) = "" Then
            sDatabasePassword = ""
        Else
            sDatabasePassword = DecryptString(rsDatabase!DatabasePassword)
        End If
        nDatabaseType = rsDatabase!DatabaseType
    
        'create connection string for selected database
        Select Case nDatabaseType
        Case eMACRODatabaseType.mdtSQLServer
                'SQL SERVER OLE DB NATIVE PROVIDER
            sDbCon = Connection_String(CONNECTION_SQLOLEDB, sServerName, sNameOfDatabase, _
                    sDatabaseUser, sDatabasePassword)
        Case eMACRODatabaseType.mdtOracle80
                'Oracle OLE DB NATIVE PROVIDER
            sDbCon = Connection_String(CONNECTION_MSDAORA, sNameOfDatabase, , _
                    sDatabaseUser, sDatabasePassword)
        End Select
        CreateDatabaseCon = True
    End If
    
    rsDatabase.Close
    Set rsDatabase = Nothing
    conMACROADODBConnection.Close
    Set conMACROADODBConnection = Nothing
    Exit Function

CatchAllError:
    sDbCon = Err.Description
    'ensure that no connections are left behind
    On Error Resume Next
    Set rsDatabase = Nothing
    Set conMACROADODBConnection = Nothing
End Function

'----------------------------------------------------------------------------------------
Public Function GetMessageHTML(ByVal sMessage As String, Optional bLoader As Boolean = False, _
    Optional ByVal enInterface As eInterface = iwww) As String
'----------------------------------------------------------------------------------------
'   ic 28/05/2003
'   function returns a formatted html message page
'   revisions
'   ic 29/06/2004 added error handling
'----------------------------------------------------------------------------------------
Dim vJSComm() As String
    
    On Error GoTo CatchAllError
    ReDim vJSComm(0)

    Call AddStringToVarArr(vJSComm, "<html><link rel='stylesheet' HREF='../style/MACRO1.css' type='text/css'><head>")
    If bLoader Then
        Call AddStringToVarArr(vJSComm, "<script language='javascript'>" & vbCrLf _
                & "function fnPageLoaded(){" _
                & "fnHideLoader();" & vbCrLf _
                & "}</script>")
    End If
    Call AddStringToVarArr(vJSComm, "</head><body onload='fnPageLoaded();'>")
    Call AddStringToVarArr(vJSComm, "<div class='clsMessageText'>" & sMessage & "</div>")
    Call AddStringToVarArr(vJSComm, "</body></html>")
    
    GetMessageHTML = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLApplication.GetMIMessageHTML"
End Function

'----------------------------------------------------------------------------------------
Public Function GetDBLockAdmin(ByRef oUser As MACROUser, ByVal enInterface As eInterface) As String
'----------------------------------------------------------------------------------------
'   ic 21/05/2003
'   function returns html for database lock admin page
'   revisions
'   ic 29/06/2004 added error handling
'   NCJ 25 Oct 06 - Handle MUSD locks; show eForm names
'----------------------------------------------------------------------------------------
Dim oDBLock As DBLock
Dim lCount As Long
Dim vLocks As Variant
Dim bRemoveAll As Boolean
Dim bRemoveOwn As Boolean
Dim vJSComm() As String


    On Error GoTo CatchAllError
    ReDim vJSComm(0)
    
    bRemoveAll = oUser.CheckPermission(gsFnRemoveAllLocks)
    bRemoveOwn = oUser.CheckPermission(gsFnRemoveOwnLocks)
    Set oDBLock = New DBLock
    
    Call AddStringToVarArr(vJSComm, "<html>" & vbCrLf _
                  & "<head>" & vbCrLf _
                  & "<title>Database Lock Administration</title>" & vbCrLf)
    
    If (enInterface = iwww) Then
        Call AddStringToVarArr(vJSComm, "<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>")
    End If
    
    Call AddStringToVarArr(vJSComm, "<script language='javascript'>" & vbCrLf _
        & "window.name='WinLock';" & vbCrLf _
        & "function fnDelete(){" & vbCrLf _
        & "if(confirm('Are you sure you wish to delete the checked lock(s)?'))document.FORM1.submit();}" & vbCrLf _
        & "</script>")
    
    Call AddStringToVarArr(vJSComm, "</head>" _
                  & "<body>")
    
    If bRemoveAll Or bRemoveOwn Then
        If bRemoveAll Then
            vLocks = oDBLock.AllLockDetails(oUser.CurrentDBConString)
        Else
            vLocks = oDBLock.AllLockDetails(oUser.CurrentDBConString, oUser.UserName)
        End If
    
        Call AddStringToVarArr(vJSComm, "<form name='FORM1' method='post' action='DialogLockAdmin.asp' target='WinLock'>")
        
        Call AddStringToVarArr(vJSComm, "<table cellpadding='0' cellspacing='0' align='center' border='0' width='600'>" _
            & "<tr height='10'><td colspan='3'></td></tr>" _
            & "<tr height='20' class='clsLabelText'>" _
            & "<td width='400'><b>Current locks</b></td>" _
            & "<td width='100'><a style='cursor:hand;' onclick='javascript:window.close();'<u>Close</u></a></td>" _
            & "<td width='100'><a style='cursor:hand;' onclick='javascript:fnDelete();'<u>Delete checked</u></a></td>" _
            & "</tr>" _
            & "<tr height='10'><td colspan='3'></td></tr>" _
            & "<tr><td colspan'3'>")
            
        Call AddStringToVarArr(vJSComm, "</td></tr></table>")
        
        'build lock admin html
        Call AddStringToVarArr(vJSComm, "<table cellpadding='0' cellspacing='0' align='center' border='0' width='800'>" _
                      & "<tr height='15' class='clsTableHeaderText'>" & vbCrLf _
                      & "<td></td>" & vbCrLf _
                      & "<td>Study</td>" & vbCrLf _
                      & "<td>Site</td>" & vbCrLf _
                      & "<td>Subject</td>" & vbCrLf _
                      & "<td>eForm</td>" & vbCrLf _
                      & "<td>eForm Cycle</td>" & vbCrLf _
                      & "<td>User</td>" & vbCrLf _
                      & "<td>Timestamp</td>" & vbCrLf _
                    & "</tr>")
        
        If Not IsNull(vLocks) Then
        
            For lCount = 0 To UBound(vLocks, 2)
                If ((lCount Mod 2) = 0) Then
                    Call AddStringToVarArr(vJSComm, "<tr height='15' class='clsTableText'>")
                Else
                    Call AddStringToVarArr(vJSComm, "<tr height='15' class='clsTableTextS'>")
                End If
                ' NCJ 25 Oct 06 - Handle MUSD eForm locks
                If vLocks(LockDetailColumn.ldcSubjectId, lCount) = 0 And vLocks(LockDetailColumn.ldcEFormInstanceId, lCount) > 0 Then
                    vLocks(LockDetailColumn.ldcEFormTitle, lCount) = vLocks(LockDetailColumn.ldcEFormSDTitle, lCount)
                End If
                ' NCJ 26 Oct 06 - Changed ldcEFormInstanceId to ldcEFormTitle to show eForm names
                Call AddStringToVarArr(vJSComm, "<td><input type='checkbox' name='chkLock' value='" & vLocks(LockDetailColumn.ldcToken, lCount) & gsDELIMITER1 & vLocks(LockDetailColumn.ldcStudyId, lCount) & gsDELIMITER1 & ConvertFromNull(vLocks(LockDetailColumn.ldcSite, lCount), vbString) & gsDELIMITER1 & ConvertFromNull(vLocks(LockDetailColumn.ldcSubjectId, lCount), vbString) & gsDELIMITER1 & ConvertFromNull(vLocks(LockDetailColumn.ldcEFormInstanceId, lCount), vbString) & "'></td>" & vbCrLf _
                    & "<td>" & IIf((vLocks(LockDetailColumn.ldcStudyId, lCount) = -1), "All studies", vLocks(LockDetailColumn.ldcStudyName, lCount)) & "</td>" _
                    & "<td>" & ConvertFromNull(vLocks(LockDetailColumn.ldcSite, lCount), vbString) & "</td>" _
                    & "<td>" & RtnSubjectText(ConvertFromNull(vLocks(LockDetailColumn.ldcSubjectId, lCount), vbString), "") & "</td>" _
                    & "<td>" & ConvertFromNull(vLocks(LockDetailColumn.ldcEFormTitle, lCount), vbString) & "&nbsp;</td>" _
                    & "<td>" & ConvertFromNull(vLocks(LockDetailColumn.ldcEFormCycleNumber, lCount), vbString) & "&nbsp;</td>" _
                    & "<td>" & ConvertFromNull(vLocks(LockDetailColumn.ldcUser, lCount), vbString) & "&nbsp;</td>" _
                    & "<td>" & Format(CDate(vLocks(LockDetailColumn.ldcLockTimeStamp, lCount)), "yyyy/mm/dd hh:mm:ss") & "</td>" _
                    & "</tr>")
            Next lCount
        End If
        
        Call AddStringToVarArr(vJSComm, "</table>")
        
        Call AddStringToVarArr(vJSComm, "</form>")
    Else
    
        'build message html
        Call AddStringToVarArr(vJSComm, "<div class='clsMessageText'>You do not have permission to remove locks</div>")
    End If
    
    Call AddStringToVarArr(vJSComm, "</body></html>")
    Set oDBLock = Nothing
    GetDBLockAdmin = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTMLApplication.GetDBLockAdmin")
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnDictionaryList(ByVal sSecDbCon) As String
'--------------------------------------------------------------------------------------------------
' ic 28/07/2005
' function returns a delimited list of the dictionaries installed in the format required by the
' interface javascript select controls
'--------------------------------------------------------------------------------------------------
Dim oDictionaries As MACROCCBS30.Dictionaries
Dim oDictionary As MACROCCBS30.Dictionary
Dim sDictionaries As String
Dim oDics As Object
Dim n As Integer

    On Error GoTo CatchAllError
    
    Set oDictionaries = New MACROCCBS30.Dictionaries
    Call oDictionaries.Init(sSecDbCon)
    
    sDictionaries = gsDELIMITER2 & gsDELIMITER3 & gsDELIMITER1 & "All dictionaries"
    Set oDics = oDictionaries.DictionaryList
    
    For n = 0 To oDictionaries.Count - 1
        Set oDictionary = oDics(n)
        sDictionaries = sDictionaries & gsDELIMITER2 & oDictionary.Name & gsDELIMITER3 _
        & oDictionary.Version & gsDELIMITER1 & oDictionary.Name & " " & oDictionary.Version
    Next
    
    Set oDics = Nothing
    Set oDictionary = Nothing
    Set oDictionaries = Nothing
    RtnDictionaryList = sDictionaries
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTMLApplication.RtnDictionaryList")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetAppMenuLh(ByRef oUser As MACROUser, _
                    Optional ByVal enInterface As eInterface = iwww, _
                    Optional ByVal bRefresh As Boolean = False, _
                    Optional ByVal vErrors As Variant, _
                    Optional bSiteDB As Boolean = True, _
                    Optional ByVal bTrace As Boolean = False) As String
'--------------------------------------------------------------------------------------------------
'   ic 07/11/2002
'   function returns application lh menu as html string
'   REVISIONS
'   DPH 19/12/2002 - Show EITHER subject label OR (subjectid) not both
'   ic 15/01/2003   fixed windows fnSubmitOptions() call - changed ',,' to ',false,false'
'   TA 11/02/2003: Added bSiteDB boolean - if false then communication menu items not shown
'   RS 18/02/2003: Added Communication Panel (moved from main Menu)
'   ic 19/02/2003  disable 'inform' checkbox if no datamonitor permission
'   TA 06/03/2002:  allow toggling of left hand menu
'   MLM 26/03/03: Display the format required for date searches
'   ic 26/03/2003 fixed search pane positional bug 774&382
'   ic 31/03/2003 hide disc,sdv,note select list when audit trail
'   ic 01/04/2003 set sdv scope statuses to selected by default
'   ic 08/04/2003 amended GetRFO()
'   ic 09/04/2003 added fnShowLegend() js call to fnPageLoaded()
'   ic 16/04/2003 code changed to exclude 'inform' checkbox if no monitor permission
'   ic 16/05/2003 replaced html select control, unaltered function copied and commented out below
'   ic 28/05/2003 quick fix for z-order bug on loading menus
'   ic 29/05/2003 added extra call to resize event. bug 1806
'   ic 18/06/2003 changed string format for regular expression, added registration
'   ic 23/06/2003 added bInit parameter for select lists
'   ic 22/09/2003 removed include for rfo.js - out-of-date reference causing iis 404 logging
'   ic 14/10/2003 handle no quicklist permission
'   ic 29/06/2004 added error handling
'   ic 09/11/2004 disable submit button after submitting
'   ic 28/07/2005 added clinical coding
'   dph 28/02/2007 bug 2882. added tracing
'   ic 08/03/2007 issue 2602 full user names with apostrophes cause the side panel to crash - js encode them
'   Mo 10/10/2007 Bug 2923, overflow error in GetAppMenuLh when user has access to a greater than integer number of subjects
'                 Loop variable changed from "nLoop as Integer" to lLoop as Long"
'--------------------------------------------------------------------------------------------------
Dim vData As Variant
Dim lLoop As Long
Dim colGeneral As Collection
Dim oStudy As Study
Dim oSite As Site
Dim sArgList As String
Dim bFunctionKeys As Boolean
Dim bSymbols As Boolean
Dim bLocalFormat As Boolean
Dim bServerTime As Boolean
Dim bSplitScreen As Boolean
Dim bSameEform As Boolean
Dim sPageLength As Variant
Dim sLabelOrId As String
Dim bChkRaised As Boolean
Dim bChkResponded As Boolean
Dim bChkClosed As Boolean
Dim vJSComm() As String
Dim oTimezone As TimeZone
Dim bMonitorReview As Boolean
Dim sTempJS As String
Dim bQuickList As Boolean
Dim oVersion As MACROVersion.Checker
Dim bCC As Boolean

    On Error GoTo CatchAllError
    
    Call WriteLog(bTrace, "modUIHTMLApplication.GetAppMenuLh start time-->" & CDate(Now))
    
    'check for clinical coding version
    Set oVersion = New MACROVersion.Checker
    bCC = oVersion.HasUpgrade("", mCCSwitch)
    ReDim vJSComm(0)
       
    'get user settings
    With oUser.UserSettings
        bFunctionKeys = CBool(.GetSetting(SETTING_VIEW_FUNCTION_KEYS, "false"))
        bSymbols = CBool(.GetSetting(SETTING_VIEW_SYMBOLS, "false"))
        bLocalFormat = CBool(.GetSetting(SETTING_LOCAL_FORMAT, "false"))
        bServerTime = CBool(.GetSetting(SETTING_SERVER_TIME, "false"))
        bSplitScreen = CBool(.GetSetting(SETTING_SPLIT_SCREEN, "false"))
        bSameEform = CBool(.GetSetting(SETTING_SAME_EFORM, "false"))
        sPageLength = .GetSetting(SETTING_PAGE_LENGTH, "50")
    End With
    If Not IsNumeric(sPageLength) Then
        sPageLength = "50"
    End If
    If oUser.CheckPermission(gsFnCreateDiscrepancy) Then
        bChkRaised = False
        bChkResponded = True
        bChkClosed = False
    ElseIf oUser.CheckPermission(gsFnChangeData) Then
        bChkRaised = True
        bChkResponded = False
        bChkClosed = False
    Else
        bChkRaised = False
        bChkResponded = False
        bChkClosed = False
    End If
    bMonitorReview = oUser.CheckPermission(gsFnMonitorDataReviewData)
    bQuickList = oUser.CheckPermission(gsFnViewQuickList)
    
    Call AddStringToVarArr(vJSComm, "<html>" & vbCrLf _
                  & "<head>" & vbCrLf)

    If (enInterface = iwww) Then
        Call AddStringToVarArr(vJSComm, "<script language='javascript'>" & vbCrLf _
                        & "var lstVisits=new Array();" & vbCrLf _
                        & "var lstEForms=new Array();" & vbCrLf _
                        & "var lstQuestions=new Array();" & vbCrLf _
                        & "var lstUsers=new Array();" & vbCrLf)
        
        
        If (bQuickList) Then
            Call WriteLog(bTrace, "modUIHTMLApplication.GetAppMenuLh-->RtnDelimitedSubjectList start-->" & CDate(Now))
            Call AddStringToVarArr(vJSComm, "var lstSubjects='" & ReplaceWithJSChars(RtnDelimitedSubjectList(oUser)) & "';" & vbCrLf)
            Call WriteLog(bTrace, "modUIHTMLApplication.GetAppMenuLh-->RtnDelimitedSubjectList end-->" & CDate(Now))
        End If
                                        
        'ic 18/06/2003 changed string format for regular expression
        'add js variables containing sites and studies
        Set colGeneral = oUser.GetAllSites
        sTempJS = sTempJS & gsDELIMITER2 & gsDELIMITER1 & "All Sites"
        For Each oSite In colGeneral
            sTempJS = sTempJS & gsDELIMITER2 & oSite.Site & gsDELIMITER1 & oSite.Site
        Next
        Set oSite = Nothing
        'sTempJS = Left(sTempJS, (Len(sTempJS) - 1))
        Call AddStringToVarArr(vJSComm, "var lstSites='" & sTempJS & "';" & vbCrLf)
        
        sTempJS = ""
        Set colGeneral = oUser.GetAllStudies
        For Each oStudy In colGeneral
            sTempJS = sTempJS & gsDELIMITER2 & oStudy.StudyId & gsDELIMITER1 & oStudy.StudyName
        Next
        'If (sTempJS <> "") Then sTempJS = Left(sTempJS, (Len(sTempJS) - 1))
        Call AddStringToVarArr(vJSComm, "var lstStudies='" & sTempJS & "';" & vbCrLf)
                
        Call AddStringToVarArr(vJSComm, "</script>")
    
        Call AddStringToVarArr(vJSComm, "<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>" & vbCrLf _
                      & "<script language='javascript' src='../script/HoverButton1.js'></script>" & vbCrLf _
                      & "<script language='javascript' src='../script/SlideMenu1.js'></script>" & vbCrLf _
                      & "<script language='javascript' src='../script/MenuLh.js'></script>" & vbCrLf _
                      & "<script language='javascript' src='../script/lUsers.js'></script>" & vbCrLf _
                      & "<script language='javascript' src='../script/SelectControl.js'></script>" & vbCrLf _
                      & "<script language='javascript' src='../script/RowOver.js'></script>" & vbCrLf)
                      
        'ic 22/09/2003 removed include for rfo.js - out-of-date reference causing iis 404 logging
        'include js for all studies user has access to
        For Each oStudy In colGeneral
            Call AddStringToVarArr(vJSComm, "<script language='javascript' src='../sites/" & oUser.DatabaseCode & "/" & oStudy.StudyId & "/lVisits.js'></script>" & vbCrLf _
                          & "<script language='javascript' src='../sites/" & oUser.DatabaseCode & "/" & oStudy.StudyId & "/lEForms.js'></script>" & vbCrLf _
                          & "<script language='javascript' src='../sites/" & oUser.DatabaseCode & "/" & oStudy.StudyId & "/lQuestions.js'></script>" & vbCrLf)
        Next

        'ic 08/03/2007 issue 2602 full user names with apostrophes cause the side panel to crash - js encode them
        Call AddStringToVarArr(vJSComm, "<script language='javascript'>" & vbCrLf _
                      & "function fnDBLockUrl(){fnReleaseLocks();}" & vbCrLf _
                      & "function fnNewSubjectUrl(){oMain.fnSaveDataFirst('fnNewSubjectUrl()');}" & vbCrLf _
                      & "function fnSubjectListUrl(sSt,sSi,sSjLb){oMain.fnSaveDataFirst('fnSubjectListUrl(" & Chr(34) & "'+sSt+'" & Chr(34) & "," & Chr(34) & "'+sSi+'" & Chr(34) & "," & Chr(34) & "'+sSjLb+'" & Chr(34) & ")');}" & vbCrLf _
                      & "function fnScheduleUrl(sSt,sSi,sSj){oMain.fnSaveDataFirst('fnScheduleUrl(" & Chr(34) & "'+sSt+'" & Chr(34) & "," & Chr(34) & "'+sSi+'" & Chr(34) & "," & Chr(34) & "'+sSj+'" & Chr(34) & "," & RtnJSBoolean(bSameEform) & ")');}" & vbCrLf _
                      & "function fnLNRUrl(){alert('not implemented');}" & vbCrLf _
                      & "function fnChangePasswordUrl(){fnChangePassword('" & ReplaceWithJSChars(oUser.UserNameFull) & "');}" & vbCrLf _
                      & "function fnRegister(){fnRegisterSubject();}" & vbCrLf _
                      & "function fnSelectChange(sLst,sVa){if(sLst=='fltSt')fnLoadStudy(sVa)}" & vbCrLf _
                      & "function fnNextDiscUrl(){alert('not implemented')}" & vbCrLf _
                      & "function fnSubmitOptions(){" & vbCrLf _
                        & "if (fnOnlyNumeric(document.Form1.optPageLength.value)){" & vbCrLf _
                          & "alert('Records per page value must be positive numerical value');" & vbCrLf _
                          & "document.Form1.optPageLength.value=" & sPageLength & ";}" & vbCrLf _
                        & "else{" & vbCrLf _
                          & "if ((document.Form1.optPageLength.value<1)||(document.Form1.optPageLength.value>" & gnMAXWWWRECORDSPERPAGE & ")){" & vbCrLf _
                            & "alert('Records per page value must be between 1 and " & gnMAXWWWRECORDSPERPAGE & "')" & vbCrLf _
                            & "document.Form1.optPageLength.value=" & sPageLength & ";}" & vbCrLf _
                          & "else{" & vbCrLf _
                            & "Form1.btnSubmit.disabled=true;" & vbCrLf _
                            & "Form1.submit();" & vbCrLf _
                        & "}}" & vbCrLf _
                      & "}")
                      
                      
        If (bSplitScreen) Then
            'without 'is eform loaded?' check
            Call AddStringToVarArr(vJSComm, "function fnRaisedDiscUrl(){oMain.fnMIMessageUrl('0','','','','','','','','','','100','');}" & vbCrLf _
                          & "function fnRespondedDiscUrl(){oMain.fnMIMessageUrl('0','','','','','','','','','','010','');}" & vbCrLf _
                          & "function fnDiscUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSs){oMain.fnMIMessageUrl('0',sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSs);}" & vbCrLf _
                          & "function fnPlannedSDVUrl(){oMain.fnMIMessageUrl('1','','','','','','','','','','1000','1111');}" & vbCrLf _
                          & "function fnSDVUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSs,sObj){oMain.fnMIMessageUrl('1',sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSs,sObj);}" & vbCrLf _
                          & "function fnNoteUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSs){oMain.fnMIMessageUrl('2',sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSs);}" & vbCrLf _
                          & "function fnBrowserUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sSs,sLk,sB4,sTm,sCm,sDi,sSd,sNo,sGet){oMain.fnBrowserUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sSs,sLk,sB4,sTm,sCm,sDi,sSd,sNo,sGet);}" & vbCrLf _
                          & "function fnViewChangesUrl(){oMain.fnBrowserUrl('','','','','','','','','','','','','','','','3');}" & vbCrLf)

        Else
            'with 'is eform loaded?' check
            Call AddStringToVarArr(vJSComm, "function fnRaisedDiscUrl(){oMain.fnSaveDataFirst('fnMIMessageUrl(" & Chr(34) & "0" & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & "100" & Chr(34) & "," & Chr(34) & Chr(34) & ")');}" & vbCrLf _
                          & "function fnRespondedDiscUrl(){oMain.fnSaveDataFirst('fnMIMessageUrl(" & Chr(34) & "0" & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & "010" & Chr(34) & "," & Chr(34) & Chr(34) & ")');}" & vbCrLf _
                          & "function fnDiscUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSs){oMain.fnSaveDataFirst('fnMIMessageUrl(" & Chr(34) & "0" & Chr(34) & "," & Chr(34) & "'+sSt+'" & Chr(34) & "," & Chr(34) & "'+sSi+'" & Chr(34) & "," & Chr(34) & "'+sVi+'" & Chr(34) & "," & Chr(34) & "'+sEf+'" & Chr(34) & "," & Chr(34) & "'+sQu+'" & Chr(34) & "," & Chr(34) & "'+sSj+'" & Chr(34) & "," & Chr(34) & "'+sUs+'" & Chr(34) & "," & Chr(34) & "'+sB4+'" & Chr(34) & "," & Chr(34) & "'+sTm+'" & Chr(34) & "," & Chr(34) & "'+sSs+'" & Chr(34) & ")');}" & vbCrLf _
                          & "function fnPlannedSDVUrl(){oMain.fnSaveDataFirst('fnMIMessageUrl(" & Chr(34) & "1" & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & "1000" & Chr(34) & "," & Chr(34) & "1111" & Chr(34) & ")');}" & vbCrLf _
                          & "function fnSDVUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSs,sObj){oMain.fnSaveDataFirst('fnMIMessageUrl(" & Chr(34) & "1" & Chr(34) & "," & Chr(34) & "'+sSt+'" & Chr(34) & "," & Chr(34) & "'+sSi+'" & Chr(34) & "," & Chr(34) & "'+sVi+'" & Chr(34) & "," & Chr(34) & "'+sEf+'" & Chr(34) & "," & Chr(34) & "'+sQu+'" & Chr(34) & "," & Chr(34) & "'+sSj+'" & Chr(34) & "," & Chr(34) & "'+sUs+'" & Chr(34) & "," & Chr(34) & "'+sB4+'" & Chr(34) & "," & Chr(34) & "'+sTm+'" & Chr(34) & "," & Chr(34) & "'+sSs+'" & Chr(34) & "," & Chr(34) & "'+sObj+'" & Chr(34) & ")');}" & vbCrLf _
                          & "function fnNoteUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSs){oMain.fnSaveDataFirst('fnMIMessageUrl(" & Chr(34) & "2" & Chr(34) & "," & Chr(34) & "'+sSt+'" & Chr(34) & "," & Chr(34) & "'+sSi+'" & Chr(34) & "," & Chr(34) & "'+sVi+'" & Chr(34) & "," & Chr(34) & "'+sEf+'" & Chr(34) & "," & Chr(34) & "'+sQu+'" & Chr(34) & "," & Chr(34) & "'+sSj+'" & Chr(34) & "," & Chr(34) & "'+sUs+'" & Chr(34) & "," & Chr(34) & "'+sB4+'" & Chr(34) & "," & Chr(34) & "'+sTm+'" & Chr(34) & "," & Chr(34) & "'+sSs+'" & Chr(34) & ")');}" & vbCrLf _
                          & "function fnBrowserUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sSs,sLk,sB4,sTm,sCm,sDi,sSd,sNo,sGet){oMain.fnSaveDataFirst('fnBrowserUrl(" & Chr(34) & "'+sSt+'" & Chr(34) & "," & Chr(34) & "'+sSi+'" & Chr(34) & "," & Chr(34) & "'+sVi+'" & Chr(34) & "," & Chr(34) & "'+sEf+'" & Chr(34) & "," & Chr(34) & "'+sQu+'" & Chr(34) & "," & Chr(34) & "'+sSj+'" & Chr(34) & "," & Chr(34) & "'+sUs+'" & Chr(34) & "," & Chr(34) & "'+sSs+'" & Chr(34) & "," & Chr(34) & "'+sLk+'" & Chr(34) & "," & Chr(34) & "'+sB4+'" & Chr(34) & "," & Chr(34) & "'+sTm+'" & Chr(34) & "," & Chr(34) & "'+sCm+'" & Chr(34) & "," & Chr(34) & "'+sDi+'" & Chr(34) & "," & Chr(34) & "'+sSd+'" & Chr(34) & "," & Chr(34) & "'+sNo+'" & Chr(34) & "," & Chr(34) & "'+sGet+'" & Chr(34) & ")');}" & vbCrLf _
                          & "function fnViewChangesUrl(){oMain.fnSaveDataFirst('fnBrowserUrl(" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & "3" & Chr(34) & ")');}" & vbCrLf)
        End If
        
        'ic 16/04/2003 code changed to exclude 'inform' checkbox if no monitor permission
        Call AddStringToVarArr(vJSComm, "function fnSelect(bChk){" & vbCrLf _
                        & "for (var n=0;n<fltSs.length;n++){")
        If (bMonitorReview) Then
            Call AddStringToVarArr(vJSComm, "fltSs[n].checked=bChk;")
        Else
            Call AddStringToVarArr(vJSComm, "if (n!=4) fltSs[n].checked=bChk;")
        End If
        Call AddStringToVarArr(vJSComm, "}}" & vbCrLf)
                     
        'ic 10/01/2003 moved pageloaded() function from js include to add page refresh functionality
        Call AddStringToVarArr(vJSComm, "function fnPageLoaded(){" & vbCrLf)
        
        'select list initialise
       Call AddStringToVarArr(vJSComm, "var ofltSt=fnSelectCreate('window.document','fltSt',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltSi',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltVi',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltEf',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltQu',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltUs',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltDi',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltSD',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltNo',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltCm',0,140,17,0);" & vbCrLf)
        
        'ic 23/06/2003 added bInit parameter=1 for last select list only
        'select list populate
        Call AddStringToVarArr(vJSComm, "fnLoadSelect('fltSt',lstStudies,0);" & vbCrLf _
                                      & "fnLoadSelect('fltSi',lstSites,0);" & vbCrLf _
                                      & "var nIndex=ofltSt.getValue();" & vbCrLf _
                                      & "if (nIndex!=null) fnLoadSelect('fltVi',lstVisits[nIndex],0);" & vbCrLf _
                                      & "if (nIndex!=null) fnLoadSelect('fltEf',lstEForms[nIndex],0);" & vbCrLf _
                                      & "if (nIndex!=null) fnLoadSelect('fltQu',lstQuestions[nIndex],0);" & vbCrLf _
                                      & "fnLoadSelect('fltUs',lstUsers,0);" & vbCrLf _
                                      & "fnLoadSelect('fltDi','|-1`|-2`All|0`None|30`Raised|20`Responded|10`Closed',0);" & vbCrLf _
                                      & "fnLoadSelect('fltSD','|-1`|-2`All|0`None|30`Planned|40`Queried|20`Done|10`Cancelled',0);" & vbCrLf _
                                      & "fnLoadSelect('fltNo','|-1`|-2`All|0`None',0);" & vbCrLf _
                                      & "fnLoadSelect('fltCm','|-1`|-2`All|0`None',1);" & vbCrLf)
                                      
        
        Call AddStringToVarArr(vJSComm, "fnShowLegend(" & RtnJSBoolean(bSymbols) & "," & RtnJSBoolean(bFunctionKeys) & ");" & vbCrLf _
                        & "fnInitialiseMenu('divMenu');" & vbCrLf _
                        & "fnInitialiseButton('divSearchBtn','clsHoverButton clsHoverButtonActive','clsHoverButton clsHoverButtonInactive','clsHoverButton clsHoverButtonSelected',true);" & vbCrLf _
                        & "fnSpaceMenus();" & vbCrLf)
                                
        If (bQuickList) Then
            Call AddStringToVarArr(vJSComm, "fnReloadQuickList(lstSubjects);" & vbCrLf)
        End If

        'errors encountered during save
        If Not IsMissing(vErrors) Then
            If Not IsEmpty(vErrors) Then
                Call AddStringToVarArr(vJSComm, "alert('MACRO encountered problems while updating. Some updates could not be completed." _
                    & "\nIncomplete updates are listed below\n\n")

                For lLoop = LBound(vErrors, 2) To UBound(vErrors, 2)
                    Call AddStringToVarArr(vJSComm, vErrors(0, lLoop) & " - " & vErrors(1, lLoop) & "\n")
                Next

                Call AddStringToVarArr(vJSComm, "');" & vbCrLf)
            End If
        End If
        
        
        
        'ic 23/04/2003
'        If bRefresh Then
'            Call AddStringToVarArr(vJSComm, "oMain.fnRefresh();" & vbCrLf)
'        End If
              
        'select list onclick handler
        Call AddStringToVarArr(vJSComm, "document.onclick=fnDocumentOnClick;")
                      
        'ic 19/06/2003 disable register menu item
        Call AddStringToVarArr(vJSComm, "fnEnableTaskListItem('RS',false);" & vbCrLf)
                      
        'ic 16/10/2003 call fnInitMain() instead of individual functions
        'ic 29/05/2003 added extra call to resize event. bug 1806
        Call AddStringToVarArr(vJSComm, "fnInitMain(" & RtnJSBoolean(bSplitScreen) & ");" & vbCrLf)
                           
        Call AddStringToVarArr(vJSComm, "}" & vbCrLf _
                        & "</script>" & vbCrLf)
                        
        Call AddStringToVarArr(vJSComm, "</head>" & vbCrLf _
                  & "<body bgcolor='#e0e0ff' onload='fnPageLoaded();'>" & vbCrLf)
    Else
    
        Call AddStringToVarArr(vJSComm, "  <script language='javascript'>" & vbCrLf)
        
      
        
        Call AddStringToVarArr(vJSComm, "function fnNewSubjectUrl(){window.navigate('VBfnNewSubjectUrl');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnSubjectListUrl(sSt,sSi,sSjLb){window.navigate('VBfnSubjectListUrl|'+sSt+'|'+sSi+'|'+sSjLb+'|'+'-1');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnRaisedDiscUrl(){window.navigate('VBfnRaisedDiscUrl');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnRespondedDiscUrl(){window.navigate('VBfnRespondedDiscUrl');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnNextDiscUrl(){window.navigate('VBfnNextDiscUrl');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnDiscUrl(sSt, sSi, sVi, sEf, sQu, sSj, sUs, sB4, sTm, sSs){window.navigate('VBfnDiscUrl|'+sSt+'|'+sSi+'|'+sVi+'|'+sEf+'|'+sQu+'|'+sSj+'|'+sUs+'|'+sB4+'|'+sTm+'|'+sSs);}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnPlannedSDVUrl(){window.navigate('VBfnPlannedSDVUrl');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnSDVUrl(sSt, sSi, sVi, sEf, sQu, sSj, sUs, sB4, sTm, sSs,sObj){window.navigate('VBfnSDVUrl|'+sSt+'|'+sSi+'|'+sVi+'|'+sEf+'|'+sQu+'|'+sSj+'|'+sUs+'|'+sB4+'|'+sTm+'|'+sSs+'|'+sObj);}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnNoteUrl(sSt, sSi, sVi, sEf, sQu, sSj, sUs, sB4, sTm, sSs){window.navigate('VBfnNoteUrl|'+sSt+'|'+sSi+'|'+sVi+'|'+sEf+'|'+sQu+'|'+sSj+'|'+sUs+'|'+sB4+'|'+sTm+'|'+sSs);}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnScheduleUrl(sSt,sSi,sSj){window.navigate('VBfnScheduleUrl|'+sSt+'|'+sSi+'|'+sSj);}" & vbCrLf)
        'fnBrowserUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sSs,sLk,sB4,sTm,sCm,sDi,sSd,sNo,"2");
        
        'ic 28/07/2005 added clinical coding
        Call AddStringToVarArr(vJSComm, "function fnBrowserUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sSs,sLk,sB4,sTm,sCm,sDi,sSd,sNo,sCS,sDc,sGet)" _
                                & "{window.navigate('VBfnBrowserUrl|'+sSt+'|'+sSi+'|'+sVi+'|'+sEf+'|'+sQu+'|'+sSj+'|'+sUs+'|'+sSs+'|'+sLk+'|'+sB4+'|'+sTm+'|'+sCm+'|'+sDi+'|'+sSd+'|'+sNo+'|'+sCS+'|'+sDc+'|'+sGet);}" & vbCrLf)

        Call AddStringToVarArr(vJSComm, "function fnSelectChange(sLst,sVa){window.navigate('VBfnSelectChange|'+sLst+'|'+sVa);}" & vbCrLf)
        
        Call AddStringToVarArr(vJSComm, "function fnSubmitOptions(bFnKey,bSymb,bDtFormat,bServerDt,bSplit,bSameForm){window.navigate('VBfnSubmitOptions|'+bFnKey+'|'+bSymb+'|'+bDtFormat+'|'+bServerDt+'|'+bSplit+'|'+bSameForm);}" & vbCrLf)
  
  'new function
        Call AddStringToVarArr(vJSComm, "function fnLNRUrl(){window.navigate('VBfnLNRUrl');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnViewChangesUrl(){window.navigate('VBfnViewChangesUrl');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnChangePasswordUrl(){window.navigate('VBfnChangePasswordUrl');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnDBLockUrl(){window.navigate('VBfnDBLockUrl');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnXferUrl(){window.navigate('VBfnXferUrl');}" & vbCrLf)
        
        Call AddStringToVarArr(vJSComm, "function fnTemplates(){window.navigate('VBfnTemplates');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnRegister(){window.navigate('VBfnRegister');}" & vbCrLf)
        
        Call AddStringToVarArr(vJSComm, "function fnViewLFUrl(){window.navigate('VBfnViewLFUrl');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnViewComUrl(){window.navigate('VBfnViewComUrl');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnChangeComUrl(){window.navigate('VBfnChangeComUrl');}" & vbCrLf)
        
        ' RS 18/03/2003
        Call AddStringToVarArr(vJSComm, "function fnResetTransferStatusUrl(){window.navigate('VBfnResetTransferStatusUrl');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnCommunicationHistoryUrl(){window.navigate('VBfnCommunicationHistoryUrl');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnCommunicationStatusReportUrl(){window.navigate('VBfnCommunicationStatusReportUrl');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnRemoteTimeSynchronisationUrl(){window.navigate('VBfnRemoteTimeSynchronisationUrl');}" & vbCrLf)
        'TA 10/03/03
        Call AddStringToVarArr(vJSComm, "function fnOCDiscUrl(){window.navigate('VBfnOCDiscUrl');}" & vbCrLf)
        'TA 28/03/03
        Call AddStringToVarArr(vJSComm, "function fnTimeSynchUrl(){window.navigate('VBfnTimeSynchUrl');}" & vbCrLf)
  
        Call AddStringToVarArr(vJSComm, " <!--r--><!--r-->")
        
        'ic 16/04/2003 code changed to exclude 'inform' checkbox if no monitor permission
        Call AddStringToVarArr(vJSComm, "function fnSelect(bChk){" & vbCrLf _
                        & "for (var n=0;n<fltSs.length;n++){")
        If (bMonitorReview) Then
            Call AddStringToVarArr(vJSComm, "fltSs[n].checked=bChk;")
        Else
            Call AddStringToVarArr(vJSComm, "if (n!=4) fltSs[n].checked=bChk;")
        End If
        Call AddStringToVarArr(vJSComm, "}}" & vbCrLf)
        
        'pageloaded stuff
        Call AddStringToVarArr(vJSComm, "function fnPageLoaded(){" & _
            "fnInitialiseMenu('divMenu');" & _
            "fnInitialiseButton('divSearchBtn','clsHoverButton clsHoverButtonActive','clsHoverButton clsHoverButtonInactive','clsHoverButton clsHoverButtonSelected',true);" & _
            "fnSpaceMenus();")
            
       
        'select list initialise
        Call AddStringToVarArr(vJSComm, "fnSelectCreate('window.document','fltSt',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltSi',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltVi',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltEf',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltQu',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltUs',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltDi',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltSD',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltNo',0,140,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltCm',0,140,17,0);" & vbCrLf)
                                      
        'ic 28/07/2005 added clinical coding
        If (bCC) Then
            Call AddStringToVarArr(vJSComm, "fnSelectCreate('window.document','fltCS',0,116,17,0);" & vbCrLf _
                                      & "fnSelectCreate('window.document','fltDc',0,116,17,0);" & vbCrLf)
                                                       
            Call AddStringToVarArr(vJSComm, "fnLoadSelect('fltCS','|-1`All statuses|1`Not coded|2`Coded|3`Pending new code|4`Auto encoded|5`Validated|6`Do not code',0);" & vbCrLf _
                                      & "fnLoadSelect('fltDc','" & RtnDictionaryList(GetSecurityConx()) & "',0);" & vbCrLf)
        End If
               
        'select list populate
        Call AddStringToVarArr(vJSComm, "fnLoadSelect('fltDi','|-1`|-2`All|0`None|30`Raised|20`Responded|10`Closed',0);" & vbCrLf _
                                      & "fnLoadSelect('fltSD','|-1`|-2`All|0`None|30`Planned|40`Queried|20`Done|10`Cancelled',0);" & vbCrLf _
                                      & "fnLoadSelect('fltNo','|-1`|-2`All|0`None',0);" & vbCrLf _
                                      & "fnLoadSelect('fltCm','|-1`|-2`All|0`None',1);" & vbCrLf)
                                      

        'ic 28/05/2003 quick fix for z-order bug on loading menus
        Dim sMenu As String
        sMenu = IIf(oUser.CheckPermission(gsFnViewQuickList), "2", "1")
        Call AddStringToVarArr(vJSComm, "fnToggleMenu(document.all['divMenu'][" & sMenu & "]);" & vbCrLf _
            & "fnToggleSearch(document.all['divMenu'][" & sMenu & "]);" & vbCrLf _
            & "fnToggleMenu(document.all['divMenu'][" & sMenu & "]);" & vbCrLf _
            & "fnToggleSearch(document.all['divMenu'][" & sMenu & "]);" & vbCrLf _
            & "fnSetMenuUnHover(document.all['divMenu'][" & sMenu & "]);}")
        
        'select list onclick handler
        Call AddStringToVarArr(vJSComm, "document.onclick=fnDocumentOnClick;" & vbCrLf)
        
        Call AddStringToVarArr(vJSComm, "</script>" & vbCrLf)
 
        Call AddStringToVarArr(vJSComm, "</head>" & vbCrLf _
                  & "<body bgcolor='#e0e0ff'>" & vbCrLf)
    End If
    
    'select panes must go here for now
    Call AddStringToVarArr(vJSComm, "<table id='sel_table_fltSt'><tr><td>" _
        & "<div id='fltSt_p' style='position:absolute; left:0; top:0; visibility:hidden; width:140px; overflow:auto; height:100; z-index:10' class='clsMSelectPane'>" _
        & "</div>" _
        & "</td></tr></table>")
    Call AddStringToVarArr(vJSComm, "<table id='sel_table_fltSi'><tr><td>" _
        & "<div id='fltSi_p' style='position:absolute; left:0; top:0; visibility:hidden; width:140px; overflow:auto; height:100; z-index:10' class='clsMSelectPane'>" _
        & "</div>" _
        & "</td></tr></table>")
    Call AddStringToVarArr(vJSComm, "<table id='sel_table_fltVi'><tr><td>" _
        & "<div id='fltVi_p' style='position:absolute; left:0; top:0; visibility:hidden; width:140px; overflow:auto; height:100; z-index:10' class='clsMSelectPane'>" _
        & "</div>" _
        & "</td></tr></table>")
    Call AddStringToVarArr(vJSComm, "<table id='sel_table_fltEf'><tr><td>" _
        & "<div id='fltEf_p' style='position:absolute; left:0; top:0; visibility:hidden; width:140px; overflow:auto; height:100; z-index:10' class='clsMSelectPane'>" _
        & "</div>" _
        & "</td></tr></table>")
    Call AddStringToVarArr(vJSComm, "<table id='sel_table_fltQu'><tr><td>" _
        & "<div id='fltQu_p' style='position:absolute; left:0; top:0; visibility:hidden; width:140px; overflow:auto; height:100; z-index:10' class='clsMSelectPane'>" _
        & "</div>" _
        & "</td></tr></table>")
    Call AddStringToVarArr(vJSComm, "<table id='sel_table_fltUs'><tr><td>" _
        & "<div id='fltUs_p' style='position:absolute; left:0; top:0; visibility:hidden; width:140px; overflow:auto; height:100; z-index:10' class='clsMSelectPane'>" _
        & "</div>" _
        & "</td></tr></table>")
    Call AddStringToVarArr(vJSComm, "<table id='sel_table_fltDi'><tr><td>" _
        & "<div id='fltDi_p' style='position:absolute; left:0; top:0; visibility:hidden; width:140px; overflow:auto; height:100; z-index:10' class='clsMSelectPane'>" _
        & "</div>" _
        & "</td></tr></table>")
    Call AddStringToVarArr(vJSComm, "<table id='sel_table_fltSD'><tr><td>" _
        & "<div id='fltSD_p' style='position:absolute; left:0; top:0; visibility:hidden; width:140px; overflow:auto; height:100; z-index:10' class='clsMSelectPane'>" _
        & "</div>" _
        & "</td></tr></table>")
    Call AddStringToVarArr(vJSComm, "<table id='sel_table_fltNo'><tr><td>" _
        & "<div id='fltNo_p' style='position:absolute; left:0; top:0; visibility:hidden; width:140px; overflow:auto; height:100; z-index:10' class='clsMSelectPane'>" _
        & "</div>" _
        & "</td></tr></table>")
    Call AddStringToVarArr(vJSComm, "<table id='sel_table_fltCm'><tr><td>" _
        & "<div id='fltCm_p' style='position:absolute; left:0; top:0; visibility:hidden; width:140px; overflow:auto; height:100; z-index:10' class='clsMSelectPane'>" _
        & "</div>" _
        & "</td></tr></table>")
    
    'ic 27/07/2005 added clinical coding
    If (bCC) Then
        Call AddStringToVarArr(vJSComm, "<table id='sel_table_fltCS'><tr><td>" _
            & "<div id='fltCS_p' style='position:absolute; left:0; top:0; visibility:hidden; width:116px; overflow:auto; height:100; z-index:10' class='clsMSelectPane'>" _
            & "</div>" _
            & "</td></tr></table>")
        Call AddStringToVarArr(vJSComm, "<table id='sel_table_fltDc'><tr><td>" _
            & "<div id='fltDc_p' style='position:absolute; left:0; top:0; visibility:hidden; width:116px; overflow:auto; height:100; z-index:10' class='clsMSelectPane'>" _
            & "</div>" _
            & "</td></tr></table>")
    End If
      
    '<!-- task list subject menu -->
    Call WriteLog(bTrace, "modUIHTMLApplication.GetAppMenuLh-->GetAppMenuLhTaskList start-->" & CDate(Now))
    Call AddStringToVarArr(vJSComm, GetAppMenuLhTaskList(oUser, enInterface, bRefresh, vErrors, bSiteDB))
    Call WriteLog(bTrace, "modUIHTMLApplication.GetAppMenuLh-->GetAppMenuLhTaskList end-->" & CDate(Now))

    '<!-- quicklist subject menu -->
    If (bQuickList) Then

        Call AddStringToVarArr(vJSComm, "<table>" _
                    & "<tr>" _
                      & "<td align='center'>" _
                        & "<div style='top:10; z-index:2;' name='subject' id='divMenu' class='clsMenuHeader clsMenuHeaderInactive' onclick='javascript:fnToggleMenu(this);' onmouseover='javascript:fnSetMenuHover(this);' onmouseout='javascript:fnSetMenuUnHover(this);'>" _
                          & "Subjects" _
                            & "<div style='position:absolute; left:0;'><img id='divMenuImg' align='right' src='../img/exp_inactive.gif'>" _
                            & "</div>" _
                          & "</div>" _
                        & "</td>" _
                      & "</tr>" _
                    & "</table>" & vbCrLf)


        Call AddStringToVarArr(vJSComm, "<div id='divMenuPane' style='overflow:auto; height:150; z-index:0' class='clsMenuPane'>" _
                        & "<div id='divQL' style='z-index:0'>" & vbCrLf)
                        
        If (enInterface = iWindows) Then
            vData = oUser.DataLists.GetSubjectList()
            
            Call AddStringToVarArr(vJSComm, "<table onmouseover='fnOnMouseOver(this,0);' onmouseout='fnOnMouseOut(this);' id='tsubject' style='cursor:hand;' width='100%' class='clsTableText'>" & vbCrLf)

            If Not IsNull(vData) Then
                For lLoop = LBound(vData, 2) To UBound(vData, 2)
                    'DPH 19/12/2002 - subject label or subjectid not both
                    If vData(eSubjectListCols.SubjectLabel, lLoop) <> "" Then
                        sLabelOrId = vData(eSubjectListCols.SubjectLabel, lLoop)
                    Else
                        sLabelOrId = "(" & vData(eSubjectListCols.SubjectId, lLoop) & ")"
                    End If
                    'striping
                    If ((lLoop Mod 2) = 0) Then
                        Call AddStringToVarArr(vJSComm, "<tr height='10'>")
                    Else
                        Call AddStringToVarArr(vJSComm, "<tr class='clsTableTextS' height='10'>")
                    End If
                    Call AddStringToVarArr(vJSComm, "<a style='text-decoration:none;' href='javascript:fnScheduleUrl(" & Chr(34) & vData(eSubjectListCols.StudyId, lLoop) & Chr(34) & "," & Chr(34) & vData(eSubjectListCols.Site, lLoop) & Chr(34) & "," & Chr(34) & vData(eSubjectListCols.SubjectId, lLoop) & Chr(34) & ");'>" _
                                    & "<td>" _
                                      & vData(eSubjectListCols.StudyName, lLoop) & "/" & vData(eSubjectListCols.Site, lLoop) & "/" & sLabelOrId _
                                    & "</td>" _
                                    & "</a>" _
                                  & "</tr>" & vbCrLf)
                Next
            End If

            Call AddStringToVarArr(vJSComm, "</table>" & vbCrLf)
        End If
        
        Call AddStringToVarArr(vJSComm, "</div>" _
                      & "</div>" & vbCrLf)
        
    End If


    '<!-- search menu -->
    Call AddStringToVarArr(vJSComm, "<table>" _
                    & "<tr>" _
                      & "<td align='center'>" _
                        & "<div style='top:10; z-index:2;' name='search' id='divMenu' class='clsMenuHeader clsMenuHeaderInactive' onclick='javascript:fnToggleMenu(this);fnToggleSearch(this);' onmouseover='javascript:fnSetMenuHover(this);' onmouseout='javascript:fnSetMenuUnHover(this);'>" _
                          & "Search" _
                            & "<div style='position:absolute; left:0;'><img id='divMenuImg' align='right' src='../img/exp_inactive.gif'>" _
                            & "</div>" _
                          & "</div>" _
                        & "</td>" _
                      & "</tr>" _
                    & "</table>" & vbCrLf)


    Call AddStringToVarArr(vJSComm, "<div id='divMenuPane' class='clsMenuPane'>" _
                    & "<table border='0'>" _
                      & "<tr align='center'>" _
                        & "<td>" & vbCrLf)


    'search mode select hover buttons
    If (oUser.CheckPermission(gsFnViewDiscrepancies)) Then
        Call AddStringToVarArr(vJSComm, "<div id='divSearchBtn' class='clsHoverButton clsHoverButtonInactive' onclick='fnSetButtonSelected(this);' onmouseover='fnSetButtonHover(this);' onmouseout='fnSetButtonUnHover(this);'>" _
                        & "Disc" _
                      & "</div>" & vbCrLf)
    Else
        Call AddStringToVarArr(vJSComm, "<div id='divSearchBtn' class='clsHoverButton clsHoverButtonDisabled'>Disc</div>" & vbCrLf)
    End If

    Call AddStringToVarArr(vJSComm, "</td>" _
                  & "<td>" & vbCrLf)

    If (oUser.CheckPermission(gsFnViewSDV)) Then
        Call AddStringToVarArr(vJSComm, "<div id='divSearchBtn' class='clsHoverButton clsHoverButtonInactive' onclick='fnSetButtonSelected(this);' onmouseover='fnSetButtonHover(this);' onmouseout='fnSetButtonUnHover(this);'>" _
                        & "SDV" _
                      & "</div>" & vbCrLf)
    Else
        Call AddStringToVarArr(vJSComm, "<div id='divSearchBtn' class='clsHoverButton clsHoverButtonDisabled'>SDV</div>" & vbCrLf)
    End If

    Call AddStringToVarArr(vJSComm, "</td>" _
                  & "<td>" & vbCrLf)

    Call AddStringToVarArr(vJSComm, "<div id='divSearchBtn' class='clsHoverButton clsHoverButtonInactive' onclick='fnSetButtonSelected(this);' onmouseover='fnSetButtonHover(this);' onmouseout='fnSetButtonUnHover(this);'>" _
                    & "Notes" _
                  & "</div>" & vbCrLf)

    Call AddStringToVarArr(vJSComm, "</td>" _
                  & "</tr>" _
                  & "<tr align='center'>" _
                  & "<td>" & vbCrLf)

    If (oUser.CheckPermission(gsFnViewData)) Then
        Call AddStringToVarArr(vJSComm, "<div id='divSearchBtn' class='clsHoverButton clsHoverButtonInactive' onclick='fnSetButtonSelected(this);' onmouseover='fnSetButtonHover(this);' onmouseout='fnSetButtonUnHover(this);'>" _
                        & "Subject" _
                      & "</div>" & vbCrLf)
    Else
        Call AddStringToVarArr(vJSComm, "<div id='divSearchBtn' class='clsHoverButton clsHoverButtonDisabled'>Subject</div>" & vbCrLf)
    End If

    Call AddStringToVarArr(vJSComm, "</td>" _
                  & "<td>" & vbCrLf)

    If (oUser.CheckPermission(gsFnViewData)) Then
        Call AddStringToVarArr(vJSComm, "<div id='divSearchBtn' class='clsHoverButton clsHoverButtonInactive' onclick='fnSetButtonSelected(this);' onmouseover='fnSetButtonHover(this);' onmouseout='fnSetButtonUnHover(this);'>" _
                        & "Data" _
                      & "</div>" & vbCrLf)
    Else
        Call AddStringToVarArr(vJSComm, "<div id='divSearchBtn' class='clsHoverButton clsHoverButtonDisabled'>Data</div>")
    End If

    Call AddStringToVarArr(vJSComm, "</td>" _
                  & "<td>" & vbCrLf)

    If (oUser.CheckPermission(gsFnViewData) And oUser.CheckPermission(gsFnViewAuditTrail)) Then
        Call AddStringToVarArr(vJSComm, "<div id='divSearchBtn' class='clsHoverButton clsHoverButtonInactive' onclick='fnSetButtonSelected(this);' onmouseover='fnSetButtonHover(this);' onmouseout='fnSetButtonUnHover(this);'>" _
                        & "Audit" _
                      & "</div>" & vbCrLf)
    Else
        Call AddStringToVarArr(vJSComm, "<div id='divSearchBtn' class='clsHoverButton clsHoverButtonDisabled'>" _
                        & "Audit" _
                      & "</div>" & vbCrLf)
    End If

    Call AddStringToVarArr(vJSComm, "</td>" _
                  & "</tr>" _
                  & "</table>" & vbCrLf)


    'search panes
    Call AddStringToVarArr(vJSComm, "<div style='position:absolute; visibility:hidden; top:40;' id='searchDiv0'>" _
                    & "<table width='100%' height='100%' class='clsLabelText'>" _
                      & "<tr height='20' align='center'>" _
                        & "<td colspan='2'>" _
                          & "<input style='width:100;' class='clsButton' type='button' value='Refresh' onclick='javascript:fnSearch(" & RtnJSBoolean(oUser.CheckPermission(gsFnMonitorDataReviewData)) & ");'>" _
                        & "</td>" _
                      & "</tr>" _
                      & "<tr height='0'></tr>" _
                      & "<tr height='20' align='left'>" _
                        & "<td width='55'>Study</td>" _
                        & "<td>" _
                        & "<div id='fltSt_c' style='position:absolute; top:30; left:62;'></div>" & vbCrLf)
    
    
    Call AddStringToVarArr(vJSComm, "</td>" _
                      & "</tr>" _
                      & "<tr height='20' align='left'>" _
                        & "<td>Site</td>" _
                        & "<td><div id='fltSi_c' style='position:absolute; top:52; left:62;'></div>" & vbCrLf)


    Call AddStringToVarArr(vJSComm, "</td>" _
                      & "</tr>" _
                      & "<tr height='20' align='left'>" _
                        & "<td>Subject</td>" _
                        & "<td><input style='width:140px;' class='clsTextbox' type='text' name='fltLb'></td>" _
                      & "</tr>" _
                    & "</table>" _
                  & "</div>" & vbCrLf)

    Call AddStringToVarArr(vJSComm, "<div style='position:absolute; left:0; top:0; visibility:hidden;' id='searchDiv2'>" _
                    & "<table width='100%' height='100%' class='clsLabelText'>" _
                      & "<tr height='1px'><td colspan='2' bgcolor='#cccccc'></td></tr>" _
                      & "<tr><td><input type='checkbox' name='fltSs'>&nbsp;OK</td><td><input type='checkbox' name='fltSs'>&nbsp;Missing</td></tr>" _
                      & "<tr><td><input type='checkbox' name='fltSs'>&nbsp;N/A</td><td><input type='checkbox' name='fltSs'>&nbsp;Warning</td></tr>" _
                      & "<tr><td><input type='checkbox'")
                      
    If Not (bMonitorReview) Then
        Call AddStringToVarArr(vJSComm, " disabled name='fltSs'>&nbsp;<font color='#a9a9a9'>Inform</font>")
    Else
        Call AddStringToVarArr(vJSComm, " name='fltSs'>&nbsp;Inform")
    End If
    
    
    Call AddStringToVarArr(vJSComm, "</td><td><input type='checkbox' name='fltSs'>&nbsp;OK Warning</td></tr>" _
                      & "<tr><td><input type='checkbox' name='fltSs'>&nbsp;U/O</td></tr>" _
                      & "<tr><td><input type='button' style='width:90;' class='clsButton' name='btn1' value='Select all' onclick='javascript:fnSelect(1);'></td>" _
                      & "<td><input type='button' style='width:90;' class='clsButton' name='btn1' value='Clear' onclick='javascript:fnSelect(0);'></td></tr>" _
                    & "</table>" _
                  & "</div>" & vbCrLf)

                                        
    If (bCC) Then
        Call AddStringToVarArr(vJSComm, "<div style='position:absolute; left:0; top:0; visibility:hidden;' id='searchDiv12'>" _
                    & "<table width='100%' height='100%' class='clsLabelText'>" _
                    & "<tr height='1px'><td colspan='2' bgcolor='#cccccc'></td></tr>" _
                      & "<tr height='20' align='left'><td colspan='2'>Coding Status</td><td>" _
                        & "<div id='fltCS_c' style='position:absolute; top:8; left:87;'></div>" _
                      & "</td></tr>" _
                      & "<tr height='20' align='left'><td>Dictionary</td><td>" _
                        & "<div id='fltDc_c' style='position:absolute; top:29; left:87;'></div>" _
                      & "</td></tr>" _
                    & "</table>" _
                  & "</div>" & vbCrLf)
    End If

                    
    Call AddStringToVarArr(vJSComm, "<div style='position:absolute; visibility:hidden;' id='searchDiv9'>" _
                    & "<table width='100%' height='100%' class='clsLabelText'>" _
                      & "<tr height='1px'><td colspan='2' bgcolor='#cccccc'></td></tr>" _
                      & "<tr height='20' align='left'><td width='55'>Disc</td><td>" _
                        & "<div id='fltDi_c' style='position:absolute; top:8; left:62;'></div>" _
                      & "</td></tr>" _
                      & "<tr height='20' align='left'><td>SDV</td><td>" _
                        & "<div id='fltSD_c' style='position:absolute; top:29; left:62;'></div>" _
                      & "</td></tr>" _
                      & "<tr height='20' align='left'><td>Notes</td><td>" _
                        & "<div id='fltNo_c' style='position:absolute; top:50; left:62;'></div>" _
                      & "</td></tr>" _
                    & "</table>" _
                  & "</div>" & vbCrLf)
    
    Call AddStringToVarArr(vJSComm, "<div style='position:absolute; visibility:hidden;' id='searchDiv11'>" _
                    & "<table width='100%' height='100%' class='clsLabelText'>" _
                      & "<tr height='1px'><td colspan='2' bgcolor='#cccccc'></td></tr>" _
                      & "<tr height='20' align='left'><td width='55'>Cmnts</td><td>" _
                        & "<div id='fltCm_c' style='position:absolute; top:8; left:62;'></div>" _
                      & "</td></tr>" _
                    & "</table>" _
                  & "</div>" & vbCrLf)
                                    
                                    
                                    
' ATN 17/1/2003
        Call AddStringToVarArr(vJSComm, "<div style='position:absolute; left:0; top:0; visibility:hidden;' id='searchDiv10'>" _
                    & "<table width='100%' height='100%' class='clsLabelText'>" _
                      & "<tr height='1px'><td colspan='2' bgcolor='#cccccc'></td></tr>" _
                      & "<tr><td><input type='checkbox' name='fltLk'>&nbsp;Locked</td><td><input type='checkbox' name='fltLk'>&nbsp;Frozen</td></tr>" _
                    & "</table>" _
                  & "</div>" & vbCrLf)
                                                                

    Call AddStringToVarArr(vJSComm, "<div style='position:absolute; visibility:hidden;' id='searchDiv1'>" _
                    & "<table width='100%' height='100%' class='clsLabelText'>" _
                      & "<tr height='1px'><td colspan='2' bgcolor='#cccccc'></td></tr>" _
                      & "<tr height='20' align='left'><td width='55'>Visit</td><td>" _
                        & "<div id='fltVi_c' style='position:absolute; top:8; left:62;'></div>" _
                      & "</td></tr>" _
                      & "<tr height='20' align='left'><td>eForm</td><td>" _
                        & "<div id='fltEf_c' style='position:absolute; top:29; left:62;'></div>" _
                      & "</td></tr>" _
                      & "<tr height='20' align='left'><td>Question</td><td>" _
                        & "<div id='fltQu_c' style='position:absolute; top:50; left:62;'></div>" _
                      & "</td></tr>" _
                      & "<tr height='1px'><td colspan='2' bgcolor='#cccccc'></td></tr>" _
                      & "<tr height='20' align='left'><td>User</td><td>" _
                        & "<div id='fltUs_c'style='position:absolute; top:78; left:62;'></div>" _
                      & "</td></tr>" _
                    & "</table>" _
                  & "</div>" & vbCrLf)
                        
                        
    'MLM 26/03/03:
    Set oTimezone = New TimeZone
    Call AddStringToVarArr(vJSComm, "<div style='position:absolute; left:0; top:0; visibility:hidden;' id='searchDiv3'>" _
                    & "<table width='100%' height='100%' class='clsLabelText'>" _
                      & "<tr height='1px'><td colspan='2' bgcolor='#cccccc'></td></tr>" _
                      & "<tr><td><input type='radio' name='fltB4'>&nbsp;Before</td><td><input checked type='radio' name='fltB4'>&nbsp;After</td></tr>" _
                      & "<tr><td colspan='2'><input class='clsTextbox' type='text' name='fltTm' size=13>&nbsp;(" & LCase(oTimezone.LocalDateFormat) & ")</td></tr>" _
                    & "</table>" _
                  & "</div>" & vbCrLf)
                  

    Call AddStringToVarArr(vJSComm, "<div style='position:absolute; left:0; top:0; visibility:hidden;' id='searchDiv4'>" _
                    & "<table width='100%' height='100%' class='clsLabelText'>" _
                      & "<tr height='1px'><td colspan='2' bgcolor='#cccccc'></td></tr>" _
                      & "<tr><td><input type='checkbox' name='fltDSs'")
    If bChkRaised Then Call AddStringToVarArr(vJSComm, " checked")
    Call AddStringToVarArr(vJSComm, ">Raised</td><td><input type='checkbox' name='fltDSs'")
    If bChkResponded Then Call AddStringToVarArr(vJSComm, " checked")
    Call AddStringToVarArr(vJSComm, ">Responded</td></tr>" _
                      & "<tr><td><input type='checkbox' name='fltDSs'")
    If bChkClosed Then Call AddStringToVarArr(vJSComm, " checked")
    Call AddStringToVarArr(vJSComm, ">Closed</td></tr>" _
                    & "</table>" _
                  & "</div>" & vbCrLf)

    Call AddStringToVarArr(vJSComm, "<div style='position:absolute; left:0; top:0; visibility:hidden;' id='searchDiv5'>" _
                    & "<table width='100%' height='100%' class='clsLabelText'>" _
                      & "<tr height='1px'><td colspan='2' bgcolor='#cccccc'></td></tr>" _
                      & "<tr><td><input type='checkbox' name='fltSSs'>&nbsp;Planned</td><td><input type='checkbox' name='fltSSs'>&nbsp;Done</td></tr>" _
                      & "<tr><td><input type='checkbox' name='fltSSs'>&nbsp;Queried</td><td><input type='checkbox' name='fltSSs'>&nbsp;Cancelled</td></tr>" _
                    & "</table>" _
                  & "</div>" & vbCrLf)

    Call AddStringToVarArr(vJSComm, "<div style='position:absolute; left:0; top:0; visibility:hidden;' id='searchDiv6'>" _
                    & "<table width='100%' height='100%' class='clsLabelText'>" _
                      & "<tr height='1px'><td colspan='2' bgcolor='#cccccc'></td></tr>" _
                      & "<tr><td><input type='checkbox' name='fltNSs'>&nbsp;Public</td><td><input type='checkbox' name='fltNSs'>&nbsp;Private</td></tr>" _
                    & "</table>" _
                  & "</div>" & vbCrLf)

    Call AddStringToVarArr(vJSComm, "<div style='position:absolute; left:0; top:0; visibility:hidden;' id='searchDiv8'>" _
                    & "<table width='100%' height='100%' class='clsLabelText'>" _
                      & "<tr height='1px'><td colspan='2' bgcolor='#cccccc'></td></tr>" _
                      & "<tr><td><input type='checkbox' name='fltObj' checked>&nbsp;Subject</td><td><input type='checkbox' name='fltObj' checked>&nbsp;Visit</td></tr>" _
                      & "<tr><td><input type='checkbox' name='fltObj' checked>&nbsp;eForm</td><td><input type='checkbox' name='fltObj' checked>&nbsp;Question</td></tr>" _
                    & "</table>" _
                  & "</div>" _
                & "</div>" & vbCrLf)

     '<!-- options menu -->
    Call AddStringToVarArr(vJSComm, "<table>" _
                    & "<tr>" _
                      & "<td align='center'>" _
                        & "<div style='top:10; z-index:2;' name='option' id='divMenu' class='clsMenuHeader clsMenuHeaderInactive' onclick='javascript:fnToggleMenu(this);' onmouseover='javascript:fnSetMenuHover(this);' onmouseout='javascript:fnSetMenuUnHover(this);'>" _
                          & "Options" _
                            & "<div style='position:absolute; left:0;'><img id='divMenuImg' align='right' src='../img/exp_inactive.gif'>" _
                            & "</div>" _
                          & "</div>" _
                        & "</td>" _
                      & "</tr>" _
                    & "</table>" & vbCrLf)


    Call AddStringToVarArr(vJSComm, "<form name='Form1' action='AppMenuLh.asp' method='post'>" _
                  & "<input type='hidden' name='optSubmit' value='true'>" _
                  & "<div id='divMenuPane' class='clsMenuPane'>" _
                    & "<table width='100%' class='clsLabelText'>" _
                      & "<tr height='20'><td><input type='checkbox' name='optFunctionKeys' ")
    If (bFunctionKeys) Then Call AddStringToVarArr(vJSComm, "checked")
    Call AddStringToVarArr(vJSComm, "></td><td>View function keys</td></tr>" _
                  & "<tr height='20'><td><input type='checkbox' name='optSymbols' ")
    If (bSymbols) Then Call AddStringToVarArr(vJSComm, "checked")
    Call AddStringToVarArr(vJSComm, "></td><td>View symbols</td></tr>")
    If (oUser.CheckPermission(gsFnChangeDateDisplay)) Then
        Call AddStringToVarArr(vJSComm, "<tr height='20'><td><input type='checkbox' name='optDateFormat' ")
        If (bLocalFormat) Then Call AddStringToVarArr(vJSComm, "checked")
        Call AddStringToVarArr(vJSComm, "></td><td>Display dates in local format<div id='divDateLabel'>")
        If (bLocalFormat) Then
            Call AddStringToVarArr(vJSComm, "(" & oUser.UserSettings.GetSetting(SETTING_LOCAL_DATE_FORMAT, "dd/mm/yyyy") & ")")
        End If
        Call AddStringToVarArr(vJSComm, "</td></tr>")
        
        'ic 21/01/2003 not to be implemented yet
'        sHTML = sHTML & "<tr height='20'><td><input type='checkbox' name='optServerDate' "
'        If (bServerTime) Then sHTML = sHTML & "checked"
'        sHTML = sHTML & "></td><td>Display server dates and times</td></tr>"
        sArgList = sArgList & ",optDateFormat.checked,false"
    Else
        sArgList = sArgList & ",false,false"
    End If
    If (oUser.CheckPermission(gsFnSplitScreen)) Then
        Call AddStringToVarArr(vJSComm, "<tr height='20'><td><input type='checkbox' name='optSplitScreen' ")
        If (bSplitScreen) Then Call AddStringToVarArr(vJSComm, "checked")
        Call AddStringToVarArr(vJSComm, "></td><td>Split screen</td></tr>")
        sArgList = sArgList & ",optSplitScreen.checked"
    Else
        sArgList = sArgList & ",false"
    End If
    Call AddStringToVarArr(vJSComm, "<tr height='20'><td><input type='checkbox' name='optSameForm' ")
    If (bSameEform) Then Call AddStringToVarArr(vJSComm, "checked")
    Call AddStringToVarArr(vJSComm, "></td><td>Open subject record at same eForm</td></tr>")
    If (enInterface = iwww) Then
        Call AddStringToVarArr(vJSComm, "<tr height='20'><td><input class='clsTextbox' type='text' size='3' name='optPageLength' value='" & sPageLength & "'></td><td>Records per page</td></tr>")
    End If
    Call AddStringToVarArr(vJSComm, "<tr height='20'><td colspan='2' align='center'>")

    If (enInterface = iwww) Then
        Call AddStringToVarArr(vJSComm, "<input name='btnSubmit' class='clsButton' type='button' value='Apply' onclick='javascript:fnSubmitOptions();'>")
    Else
        'fnSubmitOptions(functionkeys,symbols,dateformat,serverdate,splitscreen,sameform)
        'if the user doesnt have access to one of these, the argument will be 'undefined'
        Call AddStringToVarArr(vJSComm, "<input name='btnSubmit' class='clsButton' type='button' value='Apply' " _
                             & "onclick='javascript:fnSubmitOptions(optFunctionKeys.checked" _
                                                                 & ",optSymbols.checked" _
                                                                 & sArgList _
                                                                 & ",optSameForm.checked);'>")
    End If

    Call AddStringToVarArr(vJSComm, "</td></tr>" _
                  & "</table>" _
                  & "</div>" _
                  & "</form>" & vbCrLf)
                 
    
    If enInterface = iWindows Then
        'add communication menu
        Call AddStringToVarArr(vJSComm, GetAppMenuLhCommunication(oUser, bSiteDB))
    End If
    
        
    
    Call AddStringToVarArr(vJSComm, "<form name='Form2' action='AppMenuLh.asp' method='post'>" _
                  & "<input type='hidden' name='upassword'>" _
                  & "</form>")

    Call AddStringToVarArr(vJSComm, "</body>" _
                  & "</html>")

    Set oStudy = Nothing
    Set oSite = Nothing
    Set colGeneral = Nothing
    GetAppMenuLh = Join(vJSComm, "")
    ' trace
    Call WriteLog(bTrace, "modUIHMTLApplication.GetAppMenuLh End Time-->" & CDate(Now))

    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTMLApplication.GetAppMenuLh")
End Function

Public Function GetAppMenuLhTaskList(ByRef oUser As MACROUser, _
                    Optional ByVal enInterface As eInterface = iwww, _
                    Optional ByVal bRefresh As Boolean = False, _
                    Optional ByVal vErrors As Variant, Optional bSiteDB As Boolean = True) As String
'--------------------------------------------------------------------------------------------------
'return task list HTML
'
'   revisions
'   ic 13/03/2003 added <td> id tags. use tagname, less 'td' prefix for enabling/disabling/set count
'   ic 17/03/2003 added mimessage counter
'   ic 25/04/2003 added if around nr for www
'   ic 18/06/2003 uncommented registration for www
'   ic 02/09/2003 permission now required to change password, bug 1934
'   REM 05/12/03 - Added MACRO setting to enable and disable the use of the reset password permission
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim vJSComm() As String
Dim lRaised As Long
Dim lResponded As Long
Dim lPlanned As Long

    On Error GoTo CatchAllError
    ReDim vJSComm(0)
    Call RtnMIMsgStatusCount(oUser, lRaised, lResponded, lPlanned)
    
    '<!-- task list menu -->
    Call AddStringToVarArr(vJSComm, "<table>" _
                    & "<tr>" _
                      & "<td align='center'>" _
                        & "<div style='top:10; z-index:2;' name='task' id='divMenu' class='clsMenuHeader clsMenuHeaderInactive' onclick='javascript:fnToggleMenu(this);' onmouseover='javascript:fnSetMenuHover(this);' onmouseout='javascript:fnSetMenuUnHover(this);'>" _
                          & "Task List" _
                            & "<div style='position:absolute; left:0;'><img id='divMenuImg' align='right' src='../img/exp_inactive.gif'>" _
                            & "</div>" _
                          & "</div>" _
                        & "</td>" _
                      & "</tr>" _
                    & "</table>" & vbCrLf)

    'task list options
    Call AddStringToVarArr(vJSComm, "<div style='top:25; z-index:1;' id='divMenuPane' class='clsMenuPane'>" _
                    & "<table width='100%' height='100%' border='0' class='clsMenuLinkText'>" & vbCrLf)

    If (oUser.CheckPermission(gsFnCreateNewSubject)) Then
        Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                        & "<td id='td" & gsCREATE_NEW_SUBJECT_MENUID & "'><a href='javascript:fnNewSubjectUrl();'>Create new subject</a></td>" _
                      & "</tr>" & vbCrLf)
    End If

    If (oUser.CheckPermission(gsFnViewData)) Then
        Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                        & "<td id='td" & gsVIEW_SUBJECT_LIST_MENUID & "'><a href='javascript:fnSubjectListUrl(" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & ");'>View subject list</a></td>" _
                      & "</tr>" & vbCrLf)
    End If

    If (oUser.CheckPermission(gsFnViewDiscrepancies)) Then
        Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                        & "<td id='td" & gsVIEW_RAISED_DISCREPANCIES_MENUID & "'><a href='javascript:fnRaisedDiscUrl();'>View raised discrepancies (" & CStr(lRaised) & ")</a></td>" _
                      & "</tr>" & vbCrLf)
    End If

    ' NCJ 25 Jun 03 - Bug 1830 - Changed from Create permission to View
'    If (oUser.CheckPermission(gsFnCreateDiscrepancy)) Then
    If (oUser.CheckPermission(gsFnViewDiscrepancies)) Then
        Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                        & "<td id='td" & gsVIEW_RESPONDED_DISCREPANCIES_MENUID & "'><a href='javascript:fnRespondedDiscUrl();'>View responded discrepancies (" & CStr(lResponded) & ")</a></td>" _
                      & "</tr>" & vbCrLf)
    End If
    
    If (enInterface = iWindows) Then
        If oUser.ShowOCDiscMenu Then
            Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                            & "<td id='td" & gsVIEW_OC_DISCREPANCIES_MENUID & "'><a href='javascript:fnOCDiscUrl();'>View Oracle Clinical discrepancies</a></td>" _
                          & "</tr>" & vbCrLf)
        End If
    End If

'ic 16/12/2002 temporarily removed
'    If (oUser.CheckPermission(gsFnViewDiscrepancies)) Then
'        sHTML = sHTML & "<tr height='15'>" _
'                        & "<td><a href='javascript:fnNextDiscUrl();'>View next discrepancy</a></td>" _
'                      & "</tr>" & vbCrLf
'    End If

    If (oUser.CheckPermission(gsFnViewSDV)) Then
        Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                        & "<td id='td" & gsVIEW_PLANNED_SDV_MARKS_MENUID & "'><a href='javascript:fnPlannedSDVUrl();'>View planned SDV marks (" & CStr(lPlanned) & ")</a></td>" _
                      & "</tr>" & vbCrLf)
    End If

    If enInterface = iWindows Then
        If (oUser.CheckPermission(gsFnMaintainLaboratories) Or oUser.CheckPermission(gsFnMaintainNormalRanges)) Then
            Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                            & "<td id='td" & gsLAB_AND_NR_MENUID & "'><a href='javascript:fnLNRUrl();'>Laboratories and normal ranges</a></td>" _
                          & "</tr>" & vbCrLf)
        End If
    End If

    If (oUser.CheckPermission(gsFnViewChangesSinceLast)) Then
        Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                        & "<td id='td" & gsVIEW_CHANGES_SINCE_LAST_SESSION_MENUID & "'><a href='javascript:fnViewChangesUrl();'>View changes since last session</a></td>" _
                      & "</tr>" & vbCrLf)
    End If

    'TA 7/1/03: xfer data and db lock for windows only
    If enInterface = iWindows Then
         Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                        & "<td id='td" & gsTEMPLATES_MENUID & "'><a href='javascript:fnTemplates();'>Templates</a></td>" _
                      & "</tr>" & vbCrLf)
    End If
    
    'ic 13/06/2003 registration added to web
    Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                        & "<td id='td" & gsREGISTER_SUBJECT_MENUID & "'><a href='javascript:fnRegister();'>Register subject</a></td>" _
                      & "</tr>" & vbCrLf)
    
    If enInterface = iWindows Then
        If oUser.CheckPermission(gsFnViewLFHistory) Then
            Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                            & "<td id='td" & gsVIEW_LOCK_FREEZE_HISTORY_MENUID & "'><a href='javascript:fnViewLFUrl();'>View lock/freeze history</a></td>" _
                          & "</tr>" & vbCrLf)
        End If

    End If
    'ic 22/05/2003 db lock added to web
    If (oUser.CheckPermission(gsFnRemoveOwnLocks)) Or (oUser.CheckPermission(gsFnRemoveAllLocks)) Then
        Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                        & "<td id='td" & gsDB_LOCK_ADMIN_MENUID & "'><a href='javascript:fnDBLockUrl();'>Database lock administration</a></td>" _
                      & "</tr>" & vbCrLf)
    End If
    
    'REM 05/12/03 - Added MACRO setting to enable and disable the use of the reset password permission
    InitialiseSettingsFile True
    'if there is a setting to check for reset password permission then,
    If (GetMACROSetting("resetpassword", "false") = "true") Then
        'ic 02/09/2003 permission now required to change password, bug 1934
        If (oUser.CheckPermission(gsFnResetPassword)) Then
        '   TA 11/03/2003: no permission required to change own password
            Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                                & "<td id='td" & gsCHANGE_PASSWORD_MENUID & "'><a href='javascript:fnChangePasswordUrl();'>Change password</a></td>" _
                              & "</tr>" & vbCrLf)
        End If
    Else 'don't check for permission
        '   TA 11/03/2003: no permission required to change own password
            Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                                & "<td id='td" & gsCHANGE_PASSWORD_MENUID & "'><a href='javascript:fnChangePasswordUrl();'>Change password</a></td>" _
                              & "</tr>" & vbCrLf)
    End If
    
    Call AddStringToVarArr(vJSComm, "</table>" _
                  & "</div>" & vbCrLf)
                  
    GetAppMenuLhTaskList = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLApplication.GetAppMenuLhTaskList"
End Function


'--------------------------------------------------------------------------------------------------
Public Function GetAppMenuLhCommunication(ByRef oUser As MACROUser, bSiteDB As Boolean) As String
'--------------------------------------------------------------------------------------------------
'Get Communication task list - windows only
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim vJSComm() As String
    
    On Error GoTo CatchAllError
    ReDim vJSComm(0)
    
    ' RS 18/02/2003: Added Communication Menu Pane
    If bSiteDB Then
        
        '<!-- Communication Menu Start -->
        Call AddStringToVarArr(vJSComm, "<table>" _
                        & "<tr>" _
                          & "<td align='center'>" _
                            & "<div style='top:10; z-index:2;' name='comm' id='divMenu' class='clsMenuHeader clsMenuHeaderInactive' onclick='javascript:fnToggleMenu(this);' onmouseover='javascript:fnSetMenuHover(this);' onmouseout='javascript:fnSetMenuUnHover(this);'>" _
                              & "Communication" _
                                & "<div style='position:absolute; left:0;'><img id='divMenuImg' align='right' src='../img/exp_inactive.gif'>" _
                                & "</div>" _
                              & "</div>" _
                            & "</td>" _
                          & "</tr>" _
                        & "</table>" & vbCrLf)
            
        ' START OF TABLE CONTAINING MENU ITEMS
        Call AddStringToVarArr(vJSComm, "<div style='top:25;' id='divMenuPane' class='clsMenuPane'>" _
                        & "<table width='100%' height='100%' border='0' class='clsMenuLinkText'>" & vbCrLf)

        If (oUser.CheckPermission(gsFnTransferData)) Then
            Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                            & "<td id='td" & gsTRANSFER_DATA_MENUID & "'><a href='javascript:fnXferUrl();'>Transfer data</a></td>" _
                          & "</tr>" & vbCrLf)
        End If
            
        If oUser.CheckPermission(gsFnViewSiteServerCommunication) Then
            Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                          & "<td><a href='javascript:fnResetTransferStatusUrl();'>Reset transfer status</a></td>" _
                        & "</tr>" & vbCrLf)
            Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                          & "<td><a href='javascript:fnCommunicationStatusReportUrl();'>Communications status report</a></td>" _
                        & "</tr>" & vbCrLf)
        End If
        
        Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                          & "<td><a href='javascript:fnCommunicationHistoryUrl();'>History</a></td>" _
                        & "</tr>" & vbCrLf)
          
              
        'TA 11/02/03: only show data transfer optiuons if not on server
        If (oUser.CheckPermission(gsFnChangeCommSettings)) Then
            Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                            & "<td><a href='javascript:fnChangeComUrl();'>Change communication settings</a></td>" _
                          & "</tr>" & vbCrLf)
        Else
            If (oUser.CheckPermission(gsFnViewCommSettings)) Then
                Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                                & "<td><a href='javascript:fnViewComUrl();'>View communication settings</a></td>" _
                              & "</tr>" & vbCrLf)
            End If
        End If
        
        'TA 28/03/2003: add remtoe time synch option
        Call AddStringToVarArr(vJSComm, "<tr height='15'>" _
                        & "<td><a href='javascript:fnTimeSynchUrl();'>Remote time synchronisation</a></td>" _
                      & "</tr id='td" & gsREMOTE_TIME_SYNCH_MENUID & "'>" & vbCrLf)
        
        'END OF TABLE CONTAINING MENU ITEMS
        Call AddStringToVarArr(vJSComm, "</table>" _
                      & "</div>" & vbCrLf)
        
        '<!-- Communication Menu End -->
    End If  ' Communication menu only availabe in windows
    
    
    GetAppMenuLhCommunication = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLApplication.GetAppMenuLhCommunication"
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetAppMenuTop(ByRef oUser As MACROUser, Optional ByVal enInterface As eInterface = iwww) As String
'--------------------------------------------------------------------------------------------------
'   ic 06/03/2003
'   function returns application top menu as html string
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim vJSComm() As String
Dim sURL As String

    On Error GoTo CatchAllError
    ReDim vJSComm(0)
    
    InitialiseSettingsFile True
    sURL = GetMACROSetting(MACRO_SETTING_WEBHELPURL, "")
    
    Call AddStringToVarArr(vJSComm, "<html>" _
        & "<head>" & vbCrLf)
    
    If (enInterface = iwww) Then
        Call AddStringToVarArr(vJSComm, "<link rel='stylesheet' HREF='../style/MACRO1.css' type='text/css'>" & vbCrLf _
            & "<script language='javascript' src='../script/HoverButton1.js'></script>" & vbCrLf _
            & "<script language='javascript' src='../script/MenuTop.js'></script>" & vbCrLf)
    
        Call AddStringToVarArr(vJSComm, "<script language='javascript'>" & vbCrLf _
            & "function fnAbout(){" & vbCrLf _
            & "var sRtn=window.showModalDialog('AboutMACRO.asp','ModuleName=InferMed MACRO&Version=" & App.FileDescription & "','dialogHeight:349px; dialogWidth:522px; center:yes; status:0; dependent,scrollbars');}" & vbCrLf)
        
        Call AddStringToVarArr(vJSComm, "function fnHelp(){" _
            & "window.open('" & sURL & "','help');}" & vbCrLf _
            & "</script>" & vbCrLf)
            
    Else
    
    
        Call AddStringToVarArr(vJSComm, "<script language='javascript'>" & vbCrLf _
                                            & "function fnAbout(){window.navigate('VBfnAbout')};" & vbCrLf _
                                            & "function fnHelp(){window.navigate('VBfnHelp');}" & vbCrLf _
                                            & "function fnLogOutUrl(){window.navigate('VBfnLogOutUrl');}" & vbCrLf _
                                            & "function fnHoldUrl(){window.navigate('VBfnHoldUrl');}" & vbCrLf _
                                            & "function fnSwitchUrl(){window.navigate('VBfnSwitchUrl');}" _
                                            & "</script>" & vbCrLf)
    
    End If
    
    Call AddStringToVarArr(vJSComm, "</head>" _
        & "<body bgcolor='#e0e0ff' onload='fnPageLoaded();'>")
    
    '<!-- menu table start -->
    Call AddStringToVarArr(vJSComm, "<table width='700' cellpadding='0' cellspacing='0' height='100%'>" _
        & "<tr align='center'>" _
        & "<td width='4'></td>")
        
    'show/hide icon
    Call AddStringToVarArr(vJSComm, "<td width='15'>" _
        & "<img style='cursor:hand;' alt='Hide menu' id='colImg' src='../img/col_horiz_inactive.gif' onclick='fnOnClick(this);' onmouseover='fnOnMouseOver(this);' onmouseout='fnOnMouseOut(this);'>" _
        & "</td>")
    
    'user full name
    Call AddStringToVarArr(vJSComm, "<td width='170' align='center' title='Database: " & oUser.DatabaseCode _
        & "    Role: " & oUser.UserRole & "' class='clsUserNameText'>" & oUser.UserNameFull & "</td>")
      
    'hover menu
    Call AddStringToVarArr(vJSComm, "<td width='60'></td>" _
        & "<td width='70'><div id='topMenu' class='clsHoverButton clsHoverButtonInactive' onclick='javascript:fnLogOutUrl();' onmouseover='javascript:fnSetButtonHover(this);' onmouseout='javascript:fnSetButtonUnHover(this);'>Logout</div></td>" & vbCrLf _
        & "<td width='15'></td>" _
        & "<td width='70'><div id='topMenu' class='clsHoverButton clsHoverButtonInactive' onclick='javascript:fnHoldUrl();' onmouseover='javascript:fnSetButtonHover(this);' onmouseout='javascript:fnSetButtonUnHover(this);'>Stand by</div></td>" & vbCrLf _
        & "<td width='15'></td>" _
        & "<td width='70'><div id='topMenu' class='clsHoverButton clsHoverButtonInactive' onclick='javascript:fnSwitchUrl();' onmouseover='javascript:fnSetButtonHover(this);' onmouseout='javascript:fnSetButtonUnHover(this);'>Switch</div></td>" & vbCrLf _
        & "<td width='15'></td>" _
        & "<td width='70'><div id='topMenu' class='clsHoverButton clsHoverButtonInactive' onclick='javascript:fnAbout();' onmouseover='javascript:fnSetButtonHover(this);' onmouseout='javascript:fnSetButtonUnHover(this);'>About</div></td>" & vbCrLf _
        & "<td width='15'></td>" _
        & "<td width='70'><div id='topMenu' class='clsHoverButton clsHoverButtonInactive' onclick='javascript:fnHelp();' onmouseover='javascript:fnSetButtonHover(this);' onmouseout='javascript:fnSetButtonUnHover(this);'>Help</div></td>" & vbCrLf _
        & "</tr>" _
        & "</table>")
    '<!-- menu table end -->

    Call AddStringToVarArr(vJSComm, "</body>" _
        & "</html>")
    
    GetAppMenuTop = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTMLApplication.GetAppMenuTop")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetAppHeaderLh(ByRef oUser As MACROUser, _
                      Optional ByVal sLogoPath As String = "", _
                      Optional ByVal enInterface As eInterface = iwww) As String
'--------------------------------------------------------------------------------------------------
'   ic 05/11/2002
'   function returns application menu header as html string
'   TA 06/03/2002:  allow toggling of left hand menu in windows
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sHTML As String

    On Error GoTo CatchAllError

    sHTML = sHTML & "<html>" & vbCrLf _
                  & "<head>" & vbCrLf
                  
    If (enInterface = iwww) Then
        sHTML = sHTML & "<link rel='stylesheet' HREF='../style/MACRO1.css' type='text/css'>"
    End If
    
    sHTML = sHTML & "<title></title>" & vbCrLf _
                  & "<script language='javascript'>" & vbCrLf _
                    & "function fnSetButtonHover(oDiv){oDiv.className='clsMenuHeader clsMenuHeaderActive';}" & vbCrLf _
                    & "function fnSetButtonUnHover(oDiv){oDiv.className='clsMenuHeader clsMenuHeaderInactive';}" & vbCrLf
                    
    If (enInterface = iwww) Then
        sHTML = sHTML & "function fnHomeUrl(){window.parent.window.frames[4].fnSaveDataFirst('fnHomeUrl()');}" & vbCrLf
    Else
        sHTML = sHTML & "function fnHomeUrl(){window.navigate('VBfnHomeUrl');}" & vbCrLf
    End If
    
    sHTML = sHTML & "</script>" & vbCrLf _
                  & "</head>" & vbCrLf _
                  & "<body bgcolor='#e0e0ff'>"
                  
        'draw the rest as normal - logo - home buttton - username
        sHTML = sHTML & "<div style='position:absolute; left:4; top:2; width:210'>" _
                          & "<table align='center' cellpadding='0' cellspacing='0'>" _
                            & "<tr>"
                                                   
                             
        If (sLogoPath <> "") Then
            sHTML = sHTML & "<td><img src='" & sLogoPath & "'></td>"
        Else
            sHTML = sHTML & "<td class='clsMacroText'>MACRO</td>"
        End If
                              
                              
        sHTML = sHTML & "</tr>" _
                      & "</table>" _
                      & "</div>"
                      
        
        sHTML = sHTML & "<table>" _
                        & "<tr>" _
                          & "<td align='center'>" _
                            & "<div style='top:60;' " _
                                 & "id='homeDiv' class='clsMenuHeader clsMenuHeaderInactive' " _
                                 & "onclick='javascript:fnHomeUrl();' " _
                                 & "onmouseover='javascript:fnSetButtonHover(this);' " _
                                 & "onmouseout='javascript:fnSetButtonUnHover(this);'>Home" _
                            & "</div>" _
                          & "</td>" _
                        & "</tr>" _
                      & "</table>"
                      
'        sHTML = sHTML & "<table>" _
'                        & "<tr>" _
'                          & "<td align='center'>" _
'                            & "<div title='Database: " & oUser.DatabaseCode & "    Role: " & oUser.UserRole & "' style='top:75; color:blue; cursor:default; BORDER-RIGHT: blue 1px solid; BORDER-LEFT: blue 1px solid; BORDER-BOTTOM: blue 1px solid;' " _
'                                 & "id='idDiv' style='background-color:#e0e0ff;' class='clsMenuHeader'>" _
'                                 & oUser.UserNameFull _
'                            & "</div>" _
'                          & "</td>" _
'                        & "</tr>" _
'                      & "</table>"
'    End If
                  
    sHTML = sHTML & "</body>" _
                  & "</html>"

    GetAppHeaderLh = sHTML
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTMLApplication.GetAppHeaderLh")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetRFO(ByRef oStudyDef As StudyDefRO, _
              Optional ByVal enInterface As eInterface = iwww) As String
'--------------------------------------------------------------------------------------------------
'   ic 08/04/2003
'   builds and returns a js string representing rfo options
'--------------------------------------------------------------------------------------------------
' REVISIONS
' DPH 22/05/2003 - Changed to create complete Reject/Warning/Inform dialog
' ic 27/05/2005 changed warning to "W" and reject to "R"
'--------------------------------------------------------------------------------------------------
Dim vDialogHTML() As String
Dim sJS As String
Dim sOverrule As String
Dim sHTML As String
Dim nLoop As Integer

    ' initialise vDialogHTML
    ReDim vDialogHTML(0)
    
    'add rfcs to select list
    With oStudyDef.RFOs
        For nLoop = 1 To .Count
            sJS = sJS & .Item(nLoop) & gsDELIMITER1
        Next
    End With
    If (Len(sJS) > 0) Then sJS = Left(sJS, Len(sJS) - 1)
    sOverrule = "var sOverrule=" & Chr(34) & sJS & Chr(34) & ";"
    
    ' Create Form
    Call AddStringToVarArr(vDialogHTML, "<html>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<head>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<title>Reject/Warn/Inform</title>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<link rel='stylesheet' href='../../../style/MACRO1.css' type='text/css'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<script language='javascript' src='../../../script/Dialog.js'></script>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<script id='rfoScript' language='javascript' src=''></script>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<script language='javascript'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sDel1=""`"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sType;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var bCollectRFO=true;" & vbCrLf)
    ' sOverrule
    Call AddStringToVarArr(vDialogHTML, sOverrule & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnPageLoaded()" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var olArg=window.dialogArguments;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sName=olArg[0];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sType=olArg[1];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sMessage=olArg[2];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sExpression=olArg[3];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var bMenu=olArg[4];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sRFO=olArg[5];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sDatabase=olArg[6];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sStudy=olArg[7];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var bOverrule=olArg[8];" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sTitle;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sHeader;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "switch (sType)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""W"":" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sTitle=""Warning"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHeader = ""The warning for question "" + sName + "" has been generated because:"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnWarning();" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "if (sRFO=="""")" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnHideRFO();" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "this.returnValue=""R"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "else" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "txtInput.value=sRFO;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "txtInput.focus();" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "this.returnValue=""O""+fnReplaceWithJSChars(sRFO);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "if (!bOverrule) fnDisableOverrule();" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""R"":" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sTitle=""Reject"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHeader = ""The value for question "" + sName + "" has been rejected because:"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "this.returnValue=""R"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""I"":" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sTitle=""Inform"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHeader = ""The flag for question "" + sName + "" has been generated because:"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "this.returnValue=""I"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnSetName(sName);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnSetTitle(sTitle);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnSetIcon(sType,(sRFO!=""""));" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnSetHeader(sHeader);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnSetMessage(sMessage,sExpression);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnSetName(sName)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all[""divName""].innerHTML=""Question: ""+sName;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnSetTitle(sTitle)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all[""divTitle""].innerHTML=sTitle;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnSetIcon(sType,bOverruled)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sIcon="""";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "switch(sType)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""W"":sIcon=(bOverruled)?""ico_ok_warn.gif"":""ico_warn.gif"";break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""R"":sIcon=""ico_invalid.gif"";break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""I"":sIcon=""ico_inform.gif"";break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all[""divIcon""].innerHTML=""<img border='0' src='../../../img/""+sIcon+""'>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnSetHeader(sHeader)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all[""divHeader""].innerHTML=sHeader;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnSetMessage(sMessage, sExpression)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all[""divMessage""].innerHTML=sMessage;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnWarning()" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sHtm=""<table border='0'><tr height='15' class='clsLabelText'><td width='25%'>&nbsp;</td>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""<td width='75%'><div id='reasononoff'><a style='cursor:hand;' onclick='fnHideRFO();'><u>Remove overrule Reason</u></a></div></td></tr>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""<tr height='5'><td></td></tr>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""<tr id='RowRFOMess' height='30'><td colspan='4' class='clsLabelText'>""" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""Please enter or choose the reason for overruling this warning"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""</td></tr><tr id='RowRFO'>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""<td colspan='2'><input style='width:400px;' class='clsTextbox' name='txtInput' type='text'></td>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""</tr><tr id='RowRFOList'>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""<td colspan='2'><select style='width:400px;' name='selRFO' class='clsSelectList' size='3' onchange='javascript:SelectChange(\""selRFO\"",\""txtInput\"")'>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "//add rfo" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var aOverrule=sOverrule.split(sDel1);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "if (aOverrule!=undefined)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "for (var n=0;n<aOverrule.length;n++)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=(aOverrule[n]!="""")?""<option value='""+aOverrule[n]+""'>""+aOverrule[n]+""</option>"":"""";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sHtm+=""</select></td></tr></table>""" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all[""divWarning""].innerHTML=sHtm;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnClose()" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "switch (sType)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""W"":fnCheckRFO();break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""R"":fnReturn(""R"");break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "case ""I"":fnReturn(""I"");break;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnReturn(sVal)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "this.returnValue=fnReplaceWithJSChars(sVal);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "this.close();" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnDisableOverrule()" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['reasononoff'].innerHTML='';" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "selRFO.disabled=true;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "txtInput.disabled=true;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnHideRFO()" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['RowRFO'].style.visibility=""hidden"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['RowRFOList'].style.visibility=""hidden"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['RowRFOMess'].style.visibility=""hidden"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['reasononoff'].innerHTML=""<a style='cursor:hand;' onclick='fnShowRFO();'><u>Overrule this warning</u></a>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "txtInput.value="""";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "bCollectRFO=false;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnShowRFO()" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['RowRFO'].style.visibility=""visible"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['RowRFOList'].style.visibility=""visible"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['RowRFOMess'].style.visibility=""visible"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "document.all['reasononoff'].innerHTML=""<a style='cursor:hand;' onclick='fnHideRFO();'><u>Remove overrule Reason</u></a>"";" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "bCollectRFO=true;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnCheckRFO()" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "if(bCollectRFO)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "if(txtInput.value=="""")" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "alert('You must enter an overrule reason');" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "txtInput.focus();" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "else" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnReturn('O' + txtInput.value);" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "else" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "fnReturn('W');" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "function fnReplaceWithJSChars(sStr)" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "{" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "var sVal=sStr.replace(/\""/g, '\""');" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sVal=sVal.replace(/\'/g, ""\'"");" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sVal=sVal.replace(/\</g, ""&lt;"");" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "sVal=sVal.replace(/\>/g, ""&gt;"");" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "return sVal;" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "}" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</script>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</head>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<body onload='fnPageLoaded();'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<table align='center' width='95%' border='0'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<tr height='10'><td></td></tr>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<tr height='15' class='clsLabelText'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<td colspan='2'><b><div id='divName'></div></b></div></td>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<td><a style='cursor:hand;' onclick='fnClose();'><u>Close</u></a></td>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</tr>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<tr height='15' class='clsLabelText'><td bgcolor='#6699CC' colspan='4'><div id='divTitle'></div></td></tr>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<tr height='15' class='clsLabelText'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<td width='20%'><div id='divIcon'></div></td>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<td width='80%'><div id='divHeader'></div></td>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</tr>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<tr height='15' class='clsLabelText'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<td width='20%'>&nbsp;</td>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<td width='80%'><div id='divMessage'></div></td>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</tr>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<tr><td colspan='2'>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "<div id='divWarning'></div>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</td></tr>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</table>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</body>" & vbCrLf)
    Call AddStringToVarArr(vDialogHTML, "</html>" & vbCrLf)

    GetRFO = Join(vDialogHTML, "")
End Function

'ic 13/02/2003 no longer used
''--------------------------------------------------------------------------------------------------
'Public Function GetRFO(ByRef oStudyDef As StudyDefRO, _
'              Optional ByVal enInterface As eInterface = iwww) As String
''--------------------------------------------------------------------------------------------------
''   ic 26/11/02
''   builds and returns an html table representing a rfo dialog
''--------------------------------------------------------------------------------------------------
'Dim sHTML As String
'Dim nLoop As Integer
'
'
'    sHTML = sHTML & "<html>" & vbCrLf _
'                  & "<head>" & vbCrLf _
'                  & "<title>Reason For Unobtainable</title>" & vbCrLf
'
'    If (enInterface = iwww) Then
'        sHTML = sHTML & "<link rel='stylesheet' href='../../../style/MACRO1.css' type='text/css'>" & vbCrLf _
'                      & "<script language='javascript' src='../../../script/Dialog.js'></script>" & vbCrLf
'
'        sHTML = sHTML & "<script language='javascript'>" & vbCrLf _
'                      & "function fnPageLoaded()" & vbCrLf _
'                        & "{" & vbCrLf _
'                          & "var aQS=fnSplitQS(window.location.search);" & vbCrLf _
'                          & "document.all['divName'].innerHTML='<b>Question: '+aQS['name']+'</b>';" & vbCrLf _
'                          & "txtInput.focus();" & vbCrLf _
'                        & "}" & vbCrLf _
'                      & "</script>" & vbCrLf
'    End If
'
'    sHTML = sHTML & "</head>" _
'                  & "<body onload='fnPageLoaded();'>"
'
'    sHTML = sHTML & "<table align='center' width='95%' border='0'>" & vbCrLf _
'                    & "<tr height='10'><td></td></tr>" & vbCrLf _
'                    & "<tr height='15' class='clsLabelText'>" _
'                      & "<td colspan='2'>" _
'                        & "<div id='divName'></div>" _
'                      & "</td>" _
'                      & "<td><a style='cursor:hand;' onclick='javascript:fnReturn(" & Chr(34) & Chr(34) & ");'><u>Cancel</u></a></td>" _
'                      & "<td><a style='cursor:hand;' onclick='javascript:fnReturn1(txtInput.value);'><u>Close</u></a></td>" _
'                    & "</tr>" _
'                    & "<tr height='15'><td></td></tr>" & vbCrLf
'
'    sHTML = sHTML & "<tr height='5'>" _
'                      & "<td></td>" _
'                    & "</tr>" _
'                    & "<tr height='30'>" _
'                      & "<td colspan='4' class='clsLabelText'>" _
'                        & "Please enter or choose the reason for overruling this warning" _
'                      & "</td>" _
'                    & "</tr>" _
'                    & "<tr>" _
'                      & "<td width='100'></td>" _
'                      & "<td><input style='width:180px;' class='clsTextbox' name='txtInput' type='text'></td>" _
'                    & "</tr>" _
'                    & "<tr>" _
'                      & "<td width='100'></td>" _
'                      & "<td><select style='width:180px;' name='selRFO' class='clsSelectList' size='3' onchange='javascript:SelectChange(" & Chr(34) & "selRFO" & Chr(34) & "," & Chr(34) & "txtInput" & Chr(34) & ")'>"
'
'    'add rfcs to select list
'    With oStudyDef.RFOs
'        For nLoop = 1 To .Count
'            sHTML = sHTML & "<option value='" & .Item(nLoop) & "'>" _
'                          & .Item(nLoop) _
'                          & "</option>" & vbCrLf
'        Next
'    End With
'
'
'    sHTML = sHTML & "</select></td>" _
'                  & "</tr>" _
'                  & "</table>" _
'                  & "</body>" _
'                  & "</html>"
'
'
'    GetRFO = sHTML
'End Function

'--------------------------------------------------------------------------------------------------
Public Function GetRFC(ByRef oStudyDef As StudyDefRO, _
              Optional ByVal enInterface As eInterface = iwww) As String
'--------------------------------------------------------------------------------------------------
'   ic 30/10/02
'   builds and returns an html table representing a rfc dialog
'--------------------------------------------------------------------------------------------------
' REVISIONS
' DPH 21/05/2003 - OK & Cancel
'--------------------------------------------------------------------------------------------------
Dim sHTML As String
Dim nLoop As Integer
    
    
    sHTML = sHTML & "<html>" & vbCrLf _
                  & "<head>" & vbCrLf _
                  & "<title>Reason For Change</title>" & vbCrLf
    
    If (enInterface = iwww) Then
        sHTML = sHTML & "<link rel='stylesheet' href='../../../style/MACRO1.css' type='text/css'>" & vbCrLf _
                      & "<script language='javascript' src='../../../script/Dialog.js'></script>" & vbCrLf
                      
        sHTML = sHTML & "<script language='javascript'>" & vbCrLf _
                      & "function fnPageLoaded()" & vbCrLf _
                        & "{" & vbCrLf _
                          & "var aQS=fnSplitQS(window.location.search);" & vbCrLf _
                          & "document.all['divName'].innerHTML='<b>Question: '+aQS['name']+'</b>';" & vbCrLf _
                          & "document.all['divLabel'].innerHTML='Please enter or choose the reason for changing '+aQS['name'];" & vbCrLf _
                          & "txtInput.focus();" & vbCrLf _
                        & "}" & vbCrLf _
                      & "</script>" & vbCrLf
    End If

    sHTML = sHTML & "</head>" _
                  & "<body onload='fnPageLoaded();'>"

    ' DPH 21/05/2003 - OK & Cancel
    sHTML = sHTML & "<table align='center' width='95%' border='0'>" & vbCrLf _
                    & "<tr height='10'><td></td></tr>" & vbCrLf _
                    & "<tr height='15' class='clsLabelText'>" _
                      & "<td colspan='2'>" _
                        & "<div id='divName'></div>" _
                      & "</td>" _
                      & "<td><a style='cursor:hand;' onclick='javascript:fnReturn1(txtInput.value);'><u>OK</u></a></td>" _
                      & "<td><a style='cursor:hand;' onclick='javascript:fnReturn(" & Chr(34) & Chr(34) & ");'><u>Cancel</u></a></td>" _
                    & "</tr>" _
                    & "<tr height='15'><td></td></tr>" & vbCrLf
    
    sHTML = sHTML & "<tr height='5'>" _
                      & "<td></td>" _
                    & "</tr>" _
                    & "<tr height='30'>" _
                      & "<td colspan='4' class='clsLabelText'>" _
                        & "<div id='divLabel'></div>" _
                      & "</td>" _
                    & "</tr>" _
                    & "<tr>" _
                      & "<td width='100'></td>" _
                      & "<td><input style='width:180px;' class='clsTextbox' name='txtInput' type='text'></td>" _
                    & "</tr>" _
                    & "<tr>" _
                      & "<td width='100'></td>" _
                      & "<td><select style='width:180px;' name='selRFC' class='clsSelectList' size='3' onchange='javascript:SelectChange(" & Chr(34) & "selRFC" & Chr(34) & "," & Chr(34) & "txtInput" & Chr(34) & ")'>"
                      
    'add rfcs to select list
    With oStudyDef.RFCs
        For nLoop = 1 To .Count
            sHTML = sHTML & "<option value='" & .Item(nLoop) & "'>" _
                          & .Item(nLoop) _
                          & "</option>" & vbCrLf
        Next
    End With
                      
    
    sHTML = sHTML & "</select></td>" _
                  & "</tr>" _
                  & "</table>" _
                  & "</body>" _
                  & "</html>"
    
    
    GetRFC = sHTML
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetQuestionDefinition(ByRef oEformElement As eFormElementRO, _
                             Optional ByVal enInterface As eInterface = iwww) As String
'--------------------------------------------------------------------------------------------------
'   ic 14/10/02
'   builds and returns an html table representing a question audit trail
'   revisions
'   ic 31/03/2003 adjusted validation table width
'   ic 16/04/2007 issue 2759, added the derivation expression
'--------------------------------------------------------------------------------------------------

Dim oValidation As Validation
Dim sHTML As String
    
    
    sHTML = sHTML & "<html>" & vbCrLf _
                  & "<head>" & vbCrLf _
                  & "<title>Question Definition</title>" & vbCrLf
                  
    If (enInterface = iwww) Then
        sHTML = sHTML & "<link rel='stylesheet' href='../../../style/MACRO1.css' type='text/css'>" & vbCrLf
    End If
    
    sHTML = sHTML & "</head>" & vbCrLf _
                  & "<body>"
    
    'outer table, code,age,help,format rows
    sHTML = sHTML & "<table align='center' border='0' width='95%'>" & vbCrLf _
                    & "<tr height='10'><td></td></tr>" & vbCrLf _
                    & "<tr height='15' class='clsLabelText'>" _
                      & "<td colspan='2'><b>Question: " & oEformElement.Code & "</b></td>" _
                      & "<td width='10%'><a style='cursor:hand;' "
                      
    If (enInterface = iwww) Then
        sHTML = sHTML & "onclick='javascript:window.close();'"
    Else
        sHTML = sHTML & "onclick='javascript:window.navigate(" & """" & "VBfnClose" & """" & ");'"
    End If
    
    sHTML = sHTML & "><u>Close</u></a></td>" _
                    & "</tr>" _
                    & "<tr height='15'><td></td></tr>" & vbCrLf _
                    & "<tr height='15' class='clsTableText'>" _
                      & "<td width='20%'>Code:</td><td width='70%' colspan='2'>" & oEformElement.Code & "</td>" _
                    & "</tr>" & vbCrLf _
                    & "<tr height='15' class='clsTableText'>" _
                      & "<td>Name:</td><td colspan='2'>" & oEformElement.Name & "</td>" _
                    & "</tr>" & vbCrLf _
                    & "<tr height='15' class='clsTableText'>" _
                      & "<td>User help text:</td><td colspan='2'>" & oEformElement.Helptext & "</td>" _
                    & "</tr>" & vbCrLf _
                    & "<tr height='15' class='clsTableText'>" _
                      & "<td>Format:</td><td colspan='2'>" & oEformElement.Format & "</td>" _
                    & "</tr>" & vbCrLf _
                    & "<tr height='15' class='clsTableText'>" _
                      & "<td>Collect data if:</td><td colspan='2'>" & oEformElement.CollectIfCond & "</td>" _
                    & "</tr>" & vbCrLf _
                    & "<tr height='15' class='clsTableText'>" _
                      & "<td>Derivation:</td><td colspan='2'>" & oEformElement.DerivationExpr & "</td>" _
                    & "</tr>" & vbCrLf
                   
    If (oEformElement.Validations.Count > 0) Then
        sHTML = sHTML & "<tr height='10'><td></td></tr>" _
                      & "<tr height='20' valign='middle'>" _
                        & "<td class='clsTableText' colspan='3'>Validation:</td>" _
                      & "</tr>" _
                      & "<tr>" _
                        & "<td colspan='3'>"
               
        'inner validation table
        sHTML = sHTML & "<table border='1' width='100%'>" _
                        & "<tr height='15' class='clsTableHeaderText'>" _
                          & "<td width='10%'>Type</td>" _
                          & "<td width='45%'>Validation</td>" _
                          & "<td width='45%'>Message</td>" _
                        & "</tr>"
    
        'validation loop
        For Each oValidation In oEformElement.Validations
            sHTML = sHTML & "<tr height='15' class='clsTableText'>" _
                            & "<td>" & GetValidationTypeString(oValidation.ValidationType) & "</td>" _
                            & "<td>" & oValidation.ValidationCond & "</td>" _
                            & "<td>" & oValidation.MessageExpr & "</td>" _
                          & "</tr>" & vbCrLf
    
        Next

        'end of validation table
        sHTML = sHTML & "</table>"
        sHTML = sHTML & "</td></tr>"
    End If
    
    'end of outer table
    sHTML = sHTML & "</table>"

    sHTML = sHTML & "</body>" & vbCrLf _
                  & "</html>"

    Set oValidation = Nothing
    GetQuestionDefinition = sHTML
End Function


'--------------------------------------------------------------------------------------------------
Public Function GetQuestionAudit(ByRef oUser As MACROUser, ByVal lStudy As Long, ByVal sSite As String, _
                                 ByVal lPersonID As Long, ByVal lEformPageTaskId As Long, ByVal lDataItemId As Long, _
                                 ByVal sElementCode As String, ByVal sDecimalPoint As String, _
                                 ByVal sThousandSeparator As String, Optional ByVal nRepeat As Integer = 1, _
                                 Optional ByVal enInterface As eInterface = iwww) As String
'--------------------------------------------------------------------------------------------------
'   ic 11/10/02
'   builds and returns an html table representing a question audit trail
'--------------------------------------------------------------------------------------------------
' REVISIONS
' DPH 04/02/2003 - Dates displayed in yyyy/mm/dd hh:mm:ss Bug 656
' MLM 09/06/03: Only show comments if the user is entitled to see them.
' ic 26/06/2003 added decimalpoint, thousandseparator, bug 1873
' DPH 02/07/2003 - Stop null values upsetting display
' NCJ 6 Jan 04 - Stop null CTCGrade values upsetting display
' ic 21/01/2004 added GMT difference to timestamp and database timestamp
' ic 29/06/2004 added error handling
' NCJ 22 July 04 - (Bug 2349) Don't do ConvertFromNull on time zones because it gives silly results
'--------------------------------------------------------------------------------------------------
Dim vAudit As Variant
Dim sHTML As String
Dim nLoop As Integer
Dim sStatusImage As String
Dim sStatusLabel As String
Dim sLockLabel As String
Dim sTime As String
Dim sDbTime As String
Dim bViewInformIcon As Boolean
Dim nRowSpan As Integer
Dim sRowSpan As String
Dim sNRCTCString As String
Dim sStatusString As String
Dim nType As Integer
Dim sGMT As String
Dim sDbGMT As String


    On Error GoTo CatchAllError


    vAudit = RtnQuestionAudit(oUser.CurrentDBConString, lStudy, sSite, lPersonID, lEformPageTaskId, lDataItemId, nType, nRepeat)


    sHTML = sHTML & "<html>" & vbCrLf _
                  & "<head>" & vbCrLf _
                  & "<title>Question Audit</title>" & vbCrLf
                  
    If (enInterface = iwww) Then
        sHTML = sHTML & "<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>"
    End If
    
    sHTML = sHTML & "</head>" _
                  & "<body>"
    
    sHTML = sHTML & "<table align='center' border='0' width='800'>" & vbCrLf _
                    & "<tr height='10'><td></td></tr>" & vbCrLf _
                    & "<tr height='15' class='clsLabelText'>" _
                      & "<td width='650'><b>Question: " & sElementCode & "</b></td>" _
                      & "<td width='50'><a style='cursor:hand;' "
    If (enInterface = iwww) Then
        sHTML = sHTML & "onclick='javascript:window.close();'"
    Else
        sHTML = sHTML & "onclick='javascript:window.navigate(" & """" & "VBfnClose" & """" & ");'"
    End If
                      
    sHTML = sHTML & "><u>Close</u></a></td>" _
                      & "<td width='550'></td>" _
                    & "</tr>" _
                    & "<tr height='15'><td></td></tr>" & vbCrLf _
                    & "<tr><td colspan='3'>"
    
    'table header
    sHTML = sHTML & "<table border='1' width='800'>" & vbCrLf _
                    & "<tr height='15' class='clsTableHeaderText'>" & vbCrLf _
                      & "<td>Value</td>" & vbCrLf _
                      & "<td colspan='2'>Status</td>" & vbCrLf _
                      & "<td>UserName</td>" & vbCrLf _
                      & "<td>Timestamp</td>" & vbCrLf _
                      & "<td>Database timestamp</td>" & vbCrLf _
                      & "<td colspan='2'>Comments and messages</td>" & vbCrLf _
                    & "</tr>"
                    
                    
    If Not IsNull(vAudit) Then
        bViewInformIcon = oUser.CheckPermission(gsFnMonitorDataReviewData)
        
        For nLoop = 0 To UBound(vAudit, 2)
            'work out rowspan depending on number of comments/messages present
            nRowSpan = 0
            sRowSpan = ""
            If Not IsNull(vAudit(4, nLoop)) And oUser.CheckPermission(gsFnViewIComments) Then nRowSpan = nRowSpan + 1
            If Not IsNull(vAudit(5, nLoop)) Then nRowSpan = nRowSpan + 1
            If Not IsNull(vAudit(6, nLoop)) Then nRowSpan = nRowSpan + 1
            If Not IsNull(vAudit(7, nLoop)) Then nRowSpan = nRowSpan + 1
            If (nRowSpan > 1) Then sRowSpan = " rowspan='" & nRowSpan & "'"
            
            sStatusImage = RtnStatusImages(vAudit(1, nLoop), bViewInformIcon, vAudit(8, nLoop), , , , , , , sStatusLabel, sLockLabel)

            ' DPH 04/02/2003 - Dates displayed in yyyy/mm/dd hh:mm:ss Bug 656
            If vAudit(3, nLoop) <> 0 Then
                sTime = Format(CDate(vAudit(3, nLoop)), "yyyy/MM/dd HH:nn:ss")
                'sTime = CStr(CDate(vAudit(3, nLoop)))
            Else
                sTime = ""
            End If
            'ic 21/01/2004 added GMT difference to timestamp
            ' NCJ 22 July 04 - (Bug 2349) - Don't do ConvertFromNull
            ' (RtnDifferenceFromGMT already handles NULL values)
'            sGMT = RtnDifferenceFromGMT(ConvertFromNull(vAudit(13, nLoop), vbInteger))
            sGMT = RtnDifferenceFromGMT(vAudit(13, nLoop))
            
            If vAudit(12, nLoop) <> 0 Then
                sDbTime = Format(CDate(vAudit(12, nLoop)), "yyyy/MM/dd HH:nn:ss")
                'sDbTime = CStr(CDate(vAudit(12, nLoop)))
            Else
                sDbTime = ""
            End If
            'ic 21/01/2004 added GMT difference to database timestamp
            ' NCJ 22 July 04 - (Bug 2349) - Don't do ConvertFromNull
'            sDbGMT = RtnDifferenceFromGMT(ConvertFromNull(vAudit(14, nLoop), vbInteger))
            sDbGMT = RtnDifferenceFromGMT(vAudit(14, nLoop))

            ' NCJ 6 Jan 04 - Check for NULL values here (can occur in a DB upgraded from 2.1)
            If Not IsNull(vAudit(9, nLoop)) And Not IsNull(vAudit(10, nLoop)) Then
                sNRCTCString = RtnNRCTC(vAudit(1, nLoop), vAudit(9, nLoop), vAudit(10, nLoop))
            Else
                sNRCTCString = ""
            End If
            sStatusString = IIf(sNRCTCString <> "", "<table cellpadding='0' cellspacing='0'><tr class='clsTableText'><td>" & sStatusImage & sStatusLabel & "&nbsp;</td><td>" & sNRCTCString & "</td></tr></table>", sStatusImage & sStatusLabel & "&nbsp;")
            ' DPH 02/07/2003 - Stop null values upsetting display
            sHTML = sHTML & "<tr valign='top' height='15' class='clsTableText'>" & vbCrLf _
                            & "<td" & sRowSpan & ">" & LocaliseValue(ConvertFromNull(vAudit(0, nLoop), vbString), nType, sDecimalPoint, sThousandSeparator) & "&nbsp;</td>" & vbCrLf _
                            & "<td" & sRowSpan & ">" & sStatusString & "</td>" & vbCrLf _
                            & "<td width='40'" & sRowSpan & ">" & sLockLabel & "&nbsp;</td>" & vbCrLf _
                            & "<td" & sRowSpan & ">" & vAudit(2, nLoop) & "&nbsp;</td>" & vbCrLf _
                            & "<td" & sRowSpan & ">" & sTime & "&nbsp;" & sGMT & "&nbsp;</td>" & vbCrLf _
                            & "<td" & sRowSpan & ">" & sDbTime & "&nbsp;" & sDbGMT & "&nbsp;</td>" & vbCrLf
                           
            If Not IsNull(vAudit(4, nLoop)) And oUser.CheckPermission(gsFnViewIComments) Then
                sHTML = sHTML & "<td>Comments</td>" & vbCrLf _
                              & "<td>" & vAudit(4, nLoop) & "&nbsp;</td>" & vbCrLf _
                              & "</tr>"
            End If
            If Not IsNull(vAudit(5, nLoop)) Then
                If (Right(sHTML, 5) = "</tr>") Then sHTML = sHTML & "<tr valign='top' height='15' class='clsTableText'>"
                sHTML = sHTML & "<td>Reason for change</td>" & vbCrLf _
                              & "<td>" & vAudit(5, nLoop) & "&nbsp;</td>" & vbCrLf _
                              & "</tr>"
            End If
            If Not IsNull(vAudit(6, nLoop)) Then
                If (Right(sHTML, 5) = "</tr>") Then sHTML = sHTML & "<tr valign='top' height='15' class='clsTableText'>"
                sHTML = sHTML & "<td>Message</td>" & vbCrLf _
                              & "<td>" & vAudit(6, nLoop) & "&nbsp;</td>" & vbCrLf _
                              & "</tr>"
            End If
            If Not IsNull(vAudit(7, nLoop)) Then
                If (Right(sHTML, 5) = "</tr>") Then sHTML = sHTML & "<tr valign='top' height='15' class='clsTableText'>"
                sHTML = sHTML & "<td>Reason for overrule</td>" & vbCrLf _
                              & "<td>" & vAudit(7, nLoop) & "&nbsp;</td>" & vbCrLf _
                              & "</tr>"
            End If
            If (nRowSpan = 0) Then sHTML = sHTML & "<td>&nbsp;</td><td>&nbsp;</td></tr>"
        Next
    End If
    
    sHTML = sHTML & "</table>" _
                  & "</td></tr></table>" _
                  & "</body>" _
                  & "</html>"

    GetQuestionAudit = sHTML
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTMLApplication.GetQuestionAudit")
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetQuestionCodingAudit(ByRef oUser As MACROUser, ByVal lStudy As Long, ByVal sSite As String, ByVal lPersonID As Long, _
    ByVal lResponseTaskId As Long, ByVal nRepeat As Integer, sElementCode As String, Optional ByVal enInterface As eInterface = iwww) As String
'--------------------------------------------------------------------------------------------------
'   ic 12/09/2005
'   builds and returns an html table representing a question coding audit trail
'--------------------------------------------------------------------------------------------------
Dim vAudit As Variant
Dim sHTML As String
Dim nLoop As Integer
Dim sRTime As String
Dim sCTime As String
Dim sRGMT As String
Dim sCGMT As String
Dim sCodingDetails As String
Dim nCodingStatus As Integer
Dim oDictionaries As MACROCCBS30.Dictionaries
Dim oDictionary As MACROCCBS30.Dictionary
Dim sCode As String
Dim sText As String
Dim sError As String



    On Error GoTo CatchAllError

    vAudit = RtnQuestionCodingAudit(oUser.CurrentDBConString, lStudy, sSite, lPersonID, lResponseTaskId, nRepeat)


    Set oDictionaries = New MACROCCBS30.Dictionaries
    Call oDictionaries.Init(GetSecurityConx())

    sHTML = sHTML & "<html>" & vbCrLf _
                  & "<head>" & vbCrLf _
                  & "<title>Question Coding Audit</title>" & vbCrLf
                  
    If (enInterface = iwww) Then
        sHTML = sHTML & "<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>"
    End If
    
    sHTML = sHTML & "</head>" _
                  & "<body>"
    
    sHTML = sHTML & "<table align='center' border='0' width='800'>" & vbCrLf _
                    & "<tr height='10'><td></td></tr>" & vbCrLf _
                    & "<tr height='15' class='clsLabelText'>" _
                      & "<td width='650'><b>Question: " & sElementCode & "</b></td>" _
                      & "<td width='50'><a style='cursor:hand;' "
    If (enInterface = iwww) Then
        sHTML = sHTML & "onclick='javascript:window.close();'"
    Else
        sHTML = sHTML & "onclick='javascript:window.navigate(" & """" & "VBfnClose" & """" & ");'"
    End If
                      
    sHTML = sHTML & "><u>Close</u></a></td>" _
                      & "<td width='550'></td>" _
                    & "</tr>" _
                    & "<tr height='15'><td></td></tr>" & vbCrLf _
                    & "<tr><td colspan='3'>"
    
    'table header
    sHTML = sHTML & "<table border='1' width='800'>" & vbCrLf _
                    & "<tr height='15' class='clsTableHeaderText'>" & vbCrLf _
                      & "<td>Value</td>" & vbCrLf _
                      & "<td>Status</td>" & vbCrLf _
                      & "<td>Dictionary</td>" & vbCrLf _
                      & "<td>Code</td>" & vbCrLf _
                      & "<td>UserName</td>" & vbCrLf _
                      & "<td>Response timestamp</td>" & vbCrLf _
                      & "<td>Coding timestamp</td>" & vbCrLf _
                      & "<td>Reason for change</td>" & vbCrLf _
                    & "</tr>"
          
                    
    If Not IsNull(vAudit) Then
        
        For nLoop = 0 To UBound(vAudit, 2)
            
            'dates displayed in yyyy/mm/dd hh:mm:ss
            If vAudit(7, nLoop) <> 0 Then
                sRTime = Format(CDate(vAudit(7, nLoop)), "yyyy/MM/dd HH:nn:ss")
            Else
                sRTime = ""
            End If
            sRGMT = RtnDifferenceFromGMT(vAudit(8, nLoop))
            
            If vAudit(9, nLoop) <> 0 Then
                sCTime = Format(CDate(vAudit(9, nLoop)), "yyyy/MM/dd HH:nn:ss")
            Else
                sCTime = ""
            End If
            sCGMT = RtnDifferenceFromGMT(vAudit(10, nLoop))
            
            sCode = ConvertFromNull(vAudit(3, nLoop), vbString)
            If (sCode > "") Then
                'get the dictionary for this element
                Set oDictionary = oDictionaries.DictionaryFromVersion(ConvertFromNull(vAudit(0, nLoop), vbString), ConvertFromNull(vAudit(1, nLoop), vbString))
            
                If Not (oDictionary Is Nothing) Then
                    'display the code
                    If oDictionary.ToText(sCode, sText, sError) Then
                        sCode = sText
                    Else
                        sCode = "MACRO plugin '" & ConvertFromNull(vAudit(0, nLoop), vbString) & " " & ConvertFromNull(vAudit(1, nLoop), vbString) _
                        & "' encountered errors :" & vbCrLf & sError
                    End If
                Else
                     sCode = "The dictionary specified for this question was not found."
                End If
            End If

            'stop null values upsetting display
            sHTML = sHTML & "<tr valign='top' height='15' class='clsTableText'>" & vbCrLf _
                            & "<td>" & ConvertFromNull(vAudit(6, nLoop), vbString) & "&nbsp;</td>" & vbCrLf _
                            & "<td>" & GetCodingStatusString(CInt(vAudit(2, nLoop))) & "</td>" & vbCrLf _
                            & "<td>" & ConvertFromNull(vAudit(0, nLoop), vbString) & " " & ConvertFromNull(vAudit(1, nLoop), vbString) & "&nbsp;</td>" & vbCrLf _
                            & "<td>" & sCode & "&nbsp;</td>" & vbCrLf _
                            & "<td>" & ConvertFromNull(vAudit(5, nLoop), vbString) & "&nbsp;</td>" & vbCrLf _
                            & "<td>" & sRTime & "&nbsp;" & sRGMT & "&nbsp;</td>" & vbCrLf _
                            & "<td>" & sCTime & "&nbsp;" & sCGMT & "&nbsp;</td>" & vbCrLf _
                            & "<td>" & ConvertFromNull(vAudit(11, nLoop), vbString) & "&nbsp;</td>" & vbCrLf _
                            & "</tr>" & vbCrLf
        Next
    End If
    
    sHTML = sHTML & "</table>" _
                  & "</td></tr></table>" _
                  & "</body>" _
                  & "</html>"


    Set oDictionary = Nothing
    Set oDictionaries = Nothing


    GetQuestionCodingAudit = sHTML
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTMLApplication.GetQuestionCodingAudit")
End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnQuestionCodingAudit(ByVal sDbCon As String, ByVal lStudy As Long, ByVal sSite As String, _
    ByVal lPersonID As Long, ByVal lResponseTaskId As Long, ByVal nRepeat As Integer) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 13/09/2005
'--------------------------------------------------------------------------------------------------
Dim sDatabaseCnn As String
Dim oQueryDef As QueryDef
Dim oQueryServer As QueryServer
Dim vRtn As Variant


    On Error GoTo CatchAllError

    Set oQueryDef = New QueryDef
    
    With oQueryDef
        .QueryTables.Add "CodingHistory"
        
        .QueryFields.Add "DictionaryName"
        .QueryFields.Add "DictionaryVersion"
        .QueryFields.Add "CodingStatus"
        .QueryFields.Add "CodingDetails"
        .QueryFields.Add "UserName"
        .QueryFields.Add "UserNameFull"
        .QueryFields.Add "ResponseValue"
        .QueryFields.Add "ResponseTimeStamp"
        .QueryFields.Add "ResponseTimeStamp_TZ"
        .QueryFields.Add "CodingTimeStamp"
        .QueryFields.Add "CodingTimeStamp_TZ"
        .QueryFields.Add "ReasonForChange"

        
        .QueryFilters.Add "ClinicalTrialId", "=", lStudy
        .QueryFilters.Add "TrialSite", "=", sSite
        .QueryFilters.Add "PersonId", "=", lPersonID
        .QueryFilters.Add "ResponseTaskId", "=", lResponseTaskId
        .QueryFilters.Add "RepeatNumber", "=", nRepeat
        
        .QueryOrders.Add "CodingTimeStamp", True 'DESC
    End With
    
    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sDbCon
    vRtn = oQueryServer.SelectArray(oQueryDef)
    
    
    Set oQueryServer = Nothing
    Set oQueryDef = Nothing
    
    RtnQuestionCodingAudit = vRtn
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLApplication.RtnQuestionCodingAudit"
End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnQuestionAudit(ByVal sDbCon As String, ByVal lStudy As Long, ByVal sSite As String, _
                                 ByVal lPersonID As Long, ByVal lEformPageTaskId As Long, ByVal lDataItemId As Long, _
                                 ByRef nType As Integer, Optional ByVal nRepeat As Integer = 1) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 21/03/02
'   function returns a specified response audit trail as a 2d array
'   revisions
'   ic 26/06/2003 get datatype so we can local format, bug 1873
'   ic 21/01/2004 added response timezone and database timezone
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sDatabaseCnn As String
Dim oQueryDef As QueryDef
Dim oQueryServer As QueryServer
Dim vRtn As Variant


    On Error GoTo CatchAllError

    Set oQueryDef = New QueryDef
    
    With oQueryDef
        .QueryTables.Add "DataItemResponseHistory"
        
        .QueryFields.Add "ResponseValue"
        .QueryFields.Add "ResponseStatus"
        .QueryFields.Add "UserName"
        .QueryFields.Add "ResponseTimeStamp"
        .QueryFields.Add "Comments"
        .QueryFields.Add "ReasonForChange"
        .QueryFields.Add "ValidationMessage"
        .QueryFields.Add "OverruleReason"
        .QueryFields.Add "LockStatus"
        .QueryFields.Add "LabResult"
        .QueryFields.Add "CTCGrade"
        .QueryFields.Add "LaboratoryCode"
        .QueryFields.Add "DatabaseTimestamp"
        .QueryFields.Add "ResponseTimestamp_TZ"
        .QueryFields.Add "DatabaseTimestamp_TZ"
        
        .QueryFilters.Add "ClinicalTrialId", "=", lStudy
        .QueryFilters.Add "TrialSite", "=", sSite
        .QueryFilters.Add "PersonId", "=", lPersonID
        .QueryFilters.Add "DataItemId", "=", lDataItemId
        .QueryFilters.Add "CRFPageTaskId", "=", lEformPageTaskId
        .QueryFilters.Add "RepeatNumber", "=", nRepeat
        
        .QueryOrders.Add "ResponseTimeStamp", True 'DESC
        
    
    End With
    
    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sDbCon
    vRtn = oQueryServer.SelectArray(oQueryDef)
    
    'ic 26/06/2003 get datatype so we can local format
    oQueryDef.Clear
    With oQueryDef
        .QueryTables.Add "DataItem"
        
        .QueryFields.Add "DataType"
        
        .QueryFilters.Add "ClinicalTrialId", "=", lStudy
        .QueryFilters.Add "DataItemId", "=", lDataItemId
    
    End With
    
    nType = oQueryServer.SelectArray(oQueryDef)(0, 0)
    
    Set oQueryServer = Nothing
    Set oQueryDef = Nothing
    
    RtnQuestionAudit = vRtn
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLApplication.RtnQuestionAudit"
End Function

'-----------------------------------------------------------------------------
Public Function GetAbout(sModuleName As String, sVersionNumber As String, _
                           Optional ByVal enInterface As eInterface = iwww) As String
'-----------------------------------------------------------------------------
' revisions
' ic 29/01/2004 changed copyright to 2004
' ic 05/05/2005 changed copyright to 2005
' NCJ 23 Jan 06 - Changed copyright to 2006
' NCJ 18 Apr 07 - Changed copyright to 2007
' NCJ 18 Mar 08 - Changed copyright to 2008
'-----------------------------------------------------------------------------

Dim sAuthorisedUser As String
Dim sOrganisation As String
    
    sAuthorisedUser = GetMACROPCSetting(mpcAuthorisedUser, "Unknown", True)
    sOrganisation = GetMACROPCSetting(mpcOrganisation, "Unknown", True)

Dim s As String
    s = ""
    s = s & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" & vbCrLf
    s = s & "<html>" & vbCrLf
    s = s & "  <head>" & vbCrLf
    s = s & "    <title>" & vbCrLf
    s = s & "      About MACRO" & vbCrLf
    s = s & "    </title>" & vbCrLf
    If enInterface = iwww Then
        s = s & "    <link rel=""stylesheet"" href=""../style/MACRO1.css"" type=" & vbCrLf
        s = s & "    ""text/css"">" & vbCrLf
    End If
    s = s & "  </head>" & vbCrLf
    s = s & "  <body>" & vbCrLf
    s = s & "    <div>" & vbCrLf
    s = s & "      <img height=""320px"" width=""516px"" src=""../img/bg.jpg"">" & vbCrLf
    s = s & "    </div>" & vbCrLf
    s = s & "    <div style=" & vbCrLf
    s = s & "    ""position:absolute; left:0PX; top:0px;height:320px; width:515px;"">" & vbCrLf
    s = s & "      <table border=""0"" width=""100%"" height=""100%"" class=" & vbCrLf
    s = s & "      ""clsLabelText"" style=""color:#ffffff; font-weight:bold;"">" & vbCrLf
    If enInterface = iwww Then
        s = s & "           <tr><td height=""32px""></td><td></td><td></td><td align=""left""><a style=""color:#ffffff; font-weight:normal;cursor: hand;"" onclick=""javascript:window.close();"">Close</a></td></tr>" & vbCrLf
    Else
        s = s & "           <tr><td height=""32px""></td><td></td><td></td><td align=""left""><a style=""color:#ffffff; font-weight:normal;cursor: hand;"" onclick=""javascript:window.navigate('VBfnClose');"">Close</a></td></tr>" & vbCrLf
    End If
    s = s & "        <tr>" & vbCrLf
    s = s & "          <td width=""32px""></td>" & vbCrLf
    s = s & "                 <td width=""160px"" align=""middle"" valign=""middle"" align=""center"">" & vbCrLf
    s = s & "            <img src=""../img/macrologo1.gif""> " & vbCrLf
    s = s & "          </td>" & vbCrLf
    s = s & "          <td align=""middle"">" & vbCrLf
    s = s & "            For further information about MACRO, please visit our" & vbCrLf
    s = s & "            website at <a  style=""color:#ffffff; font-weight:bold;"" href=""http://www.infermed.com"" target=" & vbCrLf
    s = s & "            ""_new"">www.infermed.com</a>" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "          </td>" & vbCrLf
    s = s & "          <td valign=""top"" width=""32px"">" & vbCrLf
    s = s & "          </td>" & vbCrLf
    s = s & "        </tr>" & vbCrLf
    s = s & "        <tr>" & vbCrLf
    s = s & "          <td width=""32px""></td>" & vbCrLf
    s = s & "          <td align=""middle"" >" & vbCrLf
    s = s & "                 <br>" & vbCrLf
    s = s & "            <b>" & sModuleName & "</b>" & vbCrLf
    s = s & "          <br>" & vbCrLf
    s = s & "            Version: " & sVersionNumber & vbCrLf
    s = s & "                 </td>" & vbCrLf
    s = s & "                           <td align=""middle"" valign=""bottom"" style=""color:#ffffff; font-weight:normal;"">" & vbCrLf
    s = s & "                  Warning: This program is protected by international" & vbCrLf
    s = s & "            copyright law." & vbCrLf
    s = s & "          </td>" & vbCrLf
    s = s & "                 </tr>" & vbCrLf
    s = s & "                 <tr>" & vbCrLf
    s = s & "          <td width=""32px""></td>" & vbCrLf
    s = s & "          <td align=""middle"" valign=""bottom"" style=""color:#ffffff; font-weight:normal;"">" & vbCrLf
    s = s & "            &copy; <i>Infer</i>Med Limited, 2008" & vbCrLf
    s = s & "          </td>" & vbCrLf
    s = s & "                 <td align=""middle""  valign=""bottom"" style=""color:#ffffff; font-weight:normal;"">" & vbCrLf
    s = s & "                 This product is licensed to:<br>" & vbCrLf
    s = s & "                 " & sAuthorisedUser & "<br>" & vbCrLf
    s = s & "                 " & sOrganisation
    s = s & "                 </td>" & vbCrLf
    s = s & "        </tr>" & vbCrLf
    s = s & "                 <tr><td height=""32px""></td></tr>" & vbCrLf
    s = s & "      </table>" & vbCrLf
    s = s & "    </div>" & vbCrLf
    s = s & "  </body>" & vbCrLf
    s = s & "</html>" & vbCrLf
    s = s & "" & vbCrLf
    s = s & ""
    GetAbout = s
End Function


