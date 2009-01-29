<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Roles"

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


sQuery = "Select rolecode, roledescription,enabled,sysadmin "
sQuery = sQuery & "from role  "
sQuery = sQuery & "order by rolecode "

rsResult1.open sQuery,Connect



do until rsResult1.eof 


WriteGroupHeader "Role", rsResult1("roledescription") & " (" & fRoleEnabled( rsResult1("Enabled") ) & ")"

WriteTableStart
WriteTableRowStart
WriteCell "Role Code:"
WriteCell rsResult1("rolecode")
WriteTableRowEnd
WriteTableRowStart
WriteCell "Description:"
WriteCell rsResult1("roledescription")
WriteTableRowEnd
WriteTableEnd

sQuery = "Select  f.functioncode, f.macrofunction "
sQuery = sQuery & "from rolefunction r, macrofunction f  "
sQuery = sQuery & "where r.functioncode = f.functioncode "
sQuery = sQuery & " and r.rolecode = '" & rsResult1("rolecode") & "' "
sQuery = sQuery & "order by f.functioncode "

rsResult.open sQuery,Connect


WriteTableStart
WriteTableRowStart
WriteHeaderCell "Function Code"
WriteHeaderCell "Function"
WriteTableRowEnd



do until rsResult.eof 
WriteTableRowStart
WriteCell rsResult("functioncode") 
WriteCell rsResult("macrofunction") 

WriteTableRowEnd
rsResult.movenext
loop

WriteTableEnd

rsResult.Close


rsResult1.movenext
loop

rsResult1.Close
set RsResult1 = Nothing
set RsResult = Nothing

'*************************
' Footer block
'*************************

%>
<!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->