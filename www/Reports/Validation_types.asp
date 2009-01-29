<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Validation types"

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

sQuery = "Select validationactionname,validationtypename "
sQuery = sQuery & "from validationaction,validationtype  "
sQuery = sQuery & " where validationaction.validationactionid = validationtype.validationactionid "
sQuery = sQuery & "order by validationactionname,validationtypename "

rsResult.open sQuery,Connect

WriteTableStart
WriteTableRowStart
WriteHeaderCell "Action"
WriteHeaderCell "Type"
WriteTableRowEnd



do until rsResult.eof 
WriteTableRowStart
WriteCell rsResult("validationactionname")
WriteCell rsResult("validationtypename")
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