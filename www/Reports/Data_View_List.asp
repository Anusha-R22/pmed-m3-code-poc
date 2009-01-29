<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Data view list"

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

nClinicalTrialId = request.querystring("clinicaltrialid")

Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")

on error resume next

sQuery = "Select * from DataViewTables  "
' dph 12/02/2004 - only show relevant clinicaltrials
sQuery = sQuery & " where clinicaltrialid in (" & sPermittedStudyList & ") "
sQuery = sQuery & " order by clinicaltrialname, crfpagecode "

rsResult1.open sQuery,Connect

if err.number <> 0 then
	 Response.write "Data views have not been set up."
else

on error goto 0

response.write "<form action=""data_view.asp"" method=""get"" >"
WriteTableStart
WriteTableRowStart
WriteTableCellStart ""
WriteTableStart
WriteTableRowStart
WriteHeaderCell "Clinical trial"
WriteHeaderCell "eForm"
WriteTableRowEnd

do until rsResult1.eof 

	 WriteTableRowStart
	 WriteCell rsResult1("ClinicalTrialName") 
	 sCRFPageCode = rsResult1("crfpagecode")
	 if not isnull(rsResult1("qgroupcode")) then
		sCRFPageCode = sCRFPageCode & " (" & rsResult1("qgroupcode") & ")"
	 end if
	 WriteCell  "<input name=""tablename""  value=""" & rsResult1("dataviewname") & """ type=""checkbox"">" & sCRFPageCode & "</input>"
	 WriteTableRowEnd

	 rsResult1.movenext
loop


WriteTableEnd
WriteTableCellEnd
WriteTableCellStart "width=""50px"""
WriteTableCellEnd
WriteTableCellStart "valign=""top"""
WriteTableStart
WriteTableRowStart
WriteCell ""
WriteCell "<input type=""radio"" name=""ReportType"" value=""0"" >Display</input>"
WriteTableRowEnd
WriteTableRowStart
WriteCell ""
WriteCell "<input type=""radio"" name=""ReportType"" value=""1"" >Excel</input>"
WriteTableRowEnd
WriteTableRowStart
WriteCell ""
WriteCell "<input type=""radio"" name=""ReportType"" value=""2"" >CSV</input>"
WriteTableRowEnd
WriteTableRowStart
WriteCell "Number of records:"
WriteCell "<input type=""text"" name=""NumRows"">"
WriteTableRowEnd
WriteTableRowStart
WriteCell "Advanced:"
WriteCell "<textarea rows=""10"" cols=""30"" name=""SelectClause""></textarea>" 
WriteTableRowEnd
' DPH 17/03/2004 - 
WriteTableRowStart
WriteCell ""
WriteCell "Please select eForm view(s) by checking its associated box.<br>Note that selecting multiple views will only return data for rows with matching Studies, Sites and Subjects."
WriteTableRowEnd
WriteTableRowStart
WriteCell ""
WriteCell "<button type=""submit"">Submit</button>"
WriteTableRowEnd
response.write "</form>"
WriteTableEnd
WriteTableCellEnd
WriteTableRowEnd
WriteTableEnd




rsResult1.Close
set RsResult1 = Nothing

end if

'*************************
' Footer block
'*************************

%>
<!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->