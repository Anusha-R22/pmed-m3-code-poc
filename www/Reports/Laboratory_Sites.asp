<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Laboratory sites"

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

sQuery = "Select laboratorycode,laboratorydescription "
sQuery = sQuery & "from laboratory  "
sQuery = sQuery & "order by laboratorydescription "

rsResult1.open sQuery,Connect


do until rsResult1.eof 


WriteGroupHeader "Laboratory", rsResult1("laboratorydescription") 

sQuery = "Select site "
sQuery = sQuery & "from sitelaboratory  "
sQuery = sQuery & " where laboratorycode = '" & rsResult1("laboratorycode") & "' "
sQuery = sQuery & "order by  site "

rsResult.open sQuery,Connect


WriteTableStart
WriteTableRowStart
WriteHeaderCell "Site"
WriteTableRowEnd



do until rsResult.eof 
WriteTableRowStart
WriteCell rsResult("site") 


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