<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Standard data formats"

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

sQuery = "Select datatypename, dataformat "
sQuery = sQuery & "from datatype, standarddataformat  "
sQuery = sQuery & " where datatype.datatypeid = standarddataformat.datatypeid "
sQuery = sQuery & "order by datatypename, dataformat "

rsResult.open sQuery,Connect

WriteTableStart
WriteTableRowStart
WriteHeaderCell "Data type"
WriteHeaderCell "Format"
WriteTableRowEnd



do until rsResult.eof 
WriteTableRowStart
WriteCell rsResult("datatypename")
WriteCell rsResult("dataformat")
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