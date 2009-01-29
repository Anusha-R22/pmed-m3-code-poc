<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Clinical test groups"

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

sQuery = "Select ClinicalTestGroupcode, ClinicalTestGroupdescription "
sQuery = sQuery & "from ClinicalTestGroup  "
sQuery = sQuery & "order by ClinicalTestGroupcode "

rsResult.open sQuery,Connect

WriteTableStart
WriteTableRowStart
WriteHeaderCell "Code"
WriteHeaderCell "Description"
WriteTableRowEnd



do until rsResult.eof 
WriteTableRowStart
WriteCell rsResult("ClinicalTestGroupcode")
WriteCell rsResult("ClinicalTestGroupdescription") 
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