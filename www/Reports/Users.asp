<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Users"

sIncludeVML = 0 'Don't include VML styles

%>
<!--#include file="report_initialise.asp" -->
<!--#include file="report_open_security_database.asp" -->
<!--#include file="report_functions.asp" -->
<!--#include file="report_header.asp" -->
<%

'*************************
' Content block
'*************************
Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")

sQuery = "" & _
"SELECT MU.USERNAME," & vbCrLf & _
"       MU.USERNAMEFULL," & vbCrLf & _
"       MU.ENABLED," & vbCrLf & _
"       MU.FIRSTLOGIN," & vbCrLf & _
"       MU.LASTLOGIN," & vbCrLf & _
"       MU.SYSADMIN," & vbCrLf & _
"       MU.FAILEDATTEMPTS," & vbCrLf & _
"       UPPER(UB.DATABASECODE) DB" & vbCrLf & _
"  FROM MACROUSER    MU," & vbCrLf & _
"       USERDATABASE UB" & vbCrLf & _
" WHERE MU.USERNAME = UB.USERNAME AND" & vbCrLf & _
"       UPPER(UB.DATABASECODE) = UPPER('" & sDatabase & "')" & vbCrLf & _
" ORDER BY MU.USERNAME"

rsResult.open sQuery,Connect

sQuery = "Select  passwordretries "
sQuery = sQuery & "from MACROPassword  "

rsResult1.open sQuery,Connect

WriteTableStart
WriteTableRowStart
WriteHeaderCell "User"
WriteHeaderCell "Full user name"
WriteHeaderCell "Status"
WriteHeaderCell ""
WriteHeaderCell "First login"
WriteHeaderCell "Last login"
WriteHeaderLink ""
WriteHeaderLink ""
WriteTableRowEnd



do until rsResult.eof 
WriteTableRowStart
if cint(rsResult("sysadmin")) = 1 then
	 WriteCell rsResult("username") & "*"
else
	 WriteCell rsResult("username")
end if 
WriteCell rsResult("usernamefull") 
WriteCell fRoleEnabled( rsResult("enabled") )
if cint(rsResult("failedattempts")) >= cint(rsResult1("passwordretries")) then
	 WriteCell "Locked out"
else
	 WriteCell ""
end if 
WriteCell fConvertDate(rsResult("firstlogin") ) 
WriteCell fConvertDate(rsResult("lastlogin") ) 
WriteLink "Logins", "user_login.asp", "username=" & rsResult("username") 
WriteLink "Roles", "user_role.asp", "username=" & rsResult("username") 
WriteTableRowEnd
rsResult.movenext
loop

WriteTableEnd

WritePara "* = system administrator"

rsResult.Close
set RsResult = Nothing
rsResult1.Close
set RsResult1 = Nothing

'*************************
' Footer block
'*************************

%>
<!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->