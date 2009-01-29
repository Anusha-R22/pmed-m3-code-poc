<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Clinical tests"

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

sQuery = "Select ClinicalTestGroupcode, ClinicalTestdescription, clinicaltestcode, unit "
sQuery = sQuery & "from ClinicalTest  "
sQuery = sQuery & "order by ClinicalTestGroupcode, clinicaltestcode "

rsResult.open sQuery,Connect

WriteTableStart
WriteTableRowStart
WriteHeaderCell "Group"
WriteHeaderCell "Code"
WriteHeaderCell "Description"
WriteHeaderCell "Unit"
WriteTableRowEnd



do until rsResult.eof 
WriteTableRowStart
WriteCell rsResult("ClinicalTestGroupcode")
WriteCell rsResult("ClinicalTestcode")
WriteCell rsResult("ClinicalTestdescription") 
WriteCell rsResult("Unit")
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