<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
' cdisc study list - allows user to run against choice of study
'*************************

sReportTitle = "CDISC Study list"

sIncludeVML = 0 'Don't include VML styles

%>
<!--#include file="report_initialise.asp" -->
<%
' override querystring - this report always displayed
sReportType = 0
%>
<!--#include file="report_open_macro_database.asp" -->
<!--#include file="report_functions.asp" -->
<!--#include file="report_header.asp" -->
<%

'*************************
' Content block
'*************************

Set rsResult = CreateObject("ADODB.Recordset")

sQuery = "Select clinicaltrialid,clinicaltrialname, clinicaltrialdescription,statusname "
sQuery = sQuery & "from clinicaltrial,trialstatus  "
sQuery = sQuery & " where clinicaltrial.statusid = trialstatus.statusid "
sQuery = sQuery & "   and clinicaltrial.clinicaltrialid > 0 "
sQuery = sQuery & " order by clinicaltrialname "

rsResult.open sQuery,Connect

WriteTableStart
WriteTableRowStart
WriteHeaderCell "Study"
WriteHeaderCell "Description"
WriteHeaderCell "Status"
WriteTableRowEnd



do until rsResult.eof 
WriteTableRowStart
WriteLink rsResult("clinicaltrialname"), "cdisc_output.asp", "PrintDatabase=1&clinicaltrialid=" & rsResult("clinicaltrialid")
WriteFixedWidthCell rsResult("clinicaltrialdescription") ,320
WriteCell rsResult("statusname")
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