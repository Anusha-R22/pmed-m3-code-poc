<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "User roles"

sIncludeVML = 0 'Don't include VML styles

%>
<!--#include file="report_initialise.asp" -->
<!--#include file="report_open_macro_database.asp" -->
<!--#include file="report_functions.asp" -->
<!--#include file="report_header.asp" -->
<%

'*************************
' Content block
'*************************
Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")

sQuery = "Select  username,rolecode,studycode,sitecode "
sQuery = sQuery & "from userrole  "

if request.querystring("username") > "" then
sQuery = sQuery & " where username = '" & request.querystring("username") & "' " 
end if 

sQuery = sQuery & "order by username,rolecode "

rsResult.open sQuery,Connect


WriteTableStart
WriteTableRowStart
WriteHeaderCell "User"
WriteHeaderCell "Role"
WriteHeaderCell "Studies"
WriteHeaderCell "Sites"
WriteTableRowEnd



do until rsResult.eof 
WriteTableRowStart
WriteCell rsResult("username")
WriteCell rsResult("rolecode") 
WriteCell rsResult("studycode")
WriteCell rsResult("sitecode")
WriteTableRowEnd
rsResult.movenext
loop

WriteTableEnd

rsResult.Close
set RsResult = Nothing

'*************************
' Footer block
'*************************

%>
<!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->