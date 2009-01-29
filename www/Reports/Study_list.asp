<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Studies"

sIncludeVML = 0 'Don't include VML styles

%>
<!--#include file="report_initialise.asp" -->
<%
' DPH 15/03/2004 - Customers wish to print this report in csv
' override querystring - this report always displayed
'sReportType = 0
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
WriteLink rsResult("clinicaltrialname"), "study_details.asp", "clinicaltrialid=" & rsResult("clinicaltrialid")
WriteFixedWidthCell rsResult("clinicaltrialdescription") ,320
WriteCell rsResult("statusname")
WriteTableRowEnd
WriteTableRowStart
WriteCell ""
WriteLink "Details", "study_details.asp", "clinicaltrialid=" & rsResult("clinicaltrialid")
WriteTableRowEnd
WriteTableRowStart
WriteCell ""
WriteLink "Schedule", "study_schedule.asp", "clinicaltrialid=" & rsResult("clinicaltrialid")
WriteTableRowEnd
WriteTableRowStart
WriteCell ""
WriteLink "Questions", "study_questions.asp", "clinicaltrialid=" & rsResult("clinicaltrialid")
WriteTableRowEnd
WriteTableRowStart
WriteCell ""
WriteLink "Question details", "study_question_details.asp", "clinicaltrialid=" & rsResult("clinicaltrialid")
WriteTableRowEnd
WriteTableRowStart
WriteCell ""
WriteLink "Validations", "study_validations.asp", "clinicaltrialid=" & rsResult("clinicaltrialid")
WriteTableRowEnd
WriteTableRowStart
WriteCell ""
WriteLink "eForm details", "study_eForm_details.asp", "clinicaltrialid=" & rsResult("clinicaltrialid")
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