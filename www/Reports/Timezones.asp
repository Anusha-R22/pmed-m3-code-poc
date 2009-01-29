<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Timezones"

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

sQuery = "Select description,offsetmins "
sQuery = sQuery & "from MACROTimezone  "
sQuery = sQuery & "order by offsetmins "

rsResult.open sQuery,Connect

WriteTableStart
WriteTableRowStart
WriteHeaderCell "Description"
WriteHeaderCell "Offset (minutes)"
WriteTableRowEnd



do until rsResult.eof 
WriteTableRowStart
WriteCell rsResult("description")
WriteCell rsResult("offsetmins")
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