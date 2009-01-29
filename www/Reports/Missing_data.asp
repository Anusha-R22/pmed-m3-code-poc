<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Missing data - Study Summary"

sIncludeVML = 0 'Don't include VML styles

%>
<!--#include file="report_initialise.asp" -->
<!--#include file="report_open_macro_database.asp" -->
<!--#include file="report_functions.asp" -->
<!--#include file="report_header.asp" -->
  <%


'******************************************
' RS 11AUG2003 Bail out of no studies exist
'******************************************
if sPermittedStudyList="" then
	WriteGroupHeader "Available Studies","You do not have access to any studies"
	%>
  <!--#include file="report_close.asp" -->
  <%
	Response.end()
end if


'*************************
' Content block
'*************************

Dim sVisitCycleNumber(999)
Dim sVisitId(999)


Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")

WriteGroupHeader "Missing Data", "All Studies"

WriteTableStart
WriteTableRowStart
WriteHeaderCell "Study"
WriteHeaderCell "Number of missing values"
WriteHeaderCell "Total number"
WriteHeaderCell "%"

WriteTableRowEnd

sQuery =  "Select clinicaltrialid,  count(responsetaskid) as NumberOfValues, 'TOTAL' as CountType "
sQuery = sQuery & "from dataitemresponse"
sQuery = sQuery & "   where dataitemresponse.clinicaltrialid in (" & sPermittedStudyList & ") "
' dph 12/02/2004 - study/site permissions
sQuery = sQuery & " and " & replace(replace(sStudySiteSQL, "clinicaltrial.", "dataitemresponse."), "trialsite.", "dataitemresponse." )
sQuery = sQuery & " group by clinicaltrialid "
sQuery = sQuery & " union "
sQuery = sQuery & "Select clinicaltrialid, count(responsetaskid) as NumberOfValues, 'MISS' as CountType "
sQuery = sQuery & "from dataitemresponse"
sQuery = sQuery & " where dataitemresponse.responsestatus = 10 "
sQuery = sQuery & "   and dataitemresponse.clinicaltrialid in (" & sPermittedStudyList & ") "
' dph 12/02/2004 - study/site permissions
sQuery = sQuery & " and " & replace(replace(sStudySiteSQL, "clinicaltrial.", "dataitemresponse."), "trialsite.", "dataitemresponse." )
sQuery = sQuery & " group by clinicaltrialid "

rsResult.open sQuery,Connect

do until rsResult.eof 
	' initialise data
	nMissing = -1
	nTotal = -1
	' get data
	nClinTrialId = cint(rsResult("clinicaltrialid"))
	'trialname
	sQuery = "select clinicaltrialname from clinicaltrial where clinicaltrialid = " & rsResult("clinicaltrialid")
	rsResult1.open sQuery, Connect
	nClinName = rsResult1("clinicaltrialname")
	rsResult1.close	
	' 1st value for search criteria 
	' total response / missing
	if rsResult("CountType") = "TOTAL" Then
		nTotal = clng(rsResult("NumberOfValues"))
	elseif rsResult("CountType") = "MISS" Then
		nMissing = clng(rsResult("NumberOfValues"))
	end if
	' Should be 2nd value (but possible not if no missing values exist)
	rsResult.MoveNext 
	if not rsResult.EOF then
		' if same clinical trial
		if nClinTrialId = cint(rsResult("clinicaltrialid")) then
			if rsResult("CountType") = "TOTAL" Then
				nTotal = clng(rsResult("NumberOfValues"))
			elseif rsResult("CountType") = "MISS" Then
				nMissing = clng(rsResult("NumberOfValues"))
			end if
			rsResult.MoveNext 
		end if
	end if
	' default values
	if nTotal = -1 then
		nTotal = 0
	end if
	if nMissing = -1 then
		nMissing = 0
	end if
	' DPH - 15/03/2004 - use WriteLink to avoid problms in CSV/Excel 
	sQueryString = "ClinicalTrialId=" & nClinTrialId & "&trialsite=" & "" & ""
	WriteLink nClinName, "Missing_data_subject.asp", sQueryString
	'sCell = "<a href=""Missing_data_subject.asp?ClinicalTrialId=" & nClinTrialId & "&trialsite=" & "" & """>"
	'sCell = sCell & nClinName
	'sCell = sCell & "</a>"
	'WriteCell sCell
	sQueryString = "status=10&ClinicalTrialId=" & nClinTrialId & "&trialsite=" & "" & ""
	WriteLink nMissing, "data.asp", sQueryString
	'sCell = "<a href=""data.asp?status=10&ClinicalTrialId=" & nClinTrialId & "&trialsite=" & "" & """>"
	'sCell = sCell & nMissing
	'sCell = sCell & "</a>"
	'WriteCentredCell sCell
	WriteCentredCell nTotal
	if nTotal > 0 then
	   WriteCell cInt( (clng(nMissing) / clng(nTotal)) * 100)
	else
	   WriteCell ""
	end if
	WriteTableRowEnd
loop

WriteTableEnd

rsResult.close

set RsResult = Nothing
set RsResult1 = Nothing

'*************************
' Footer block
'*************************

%>
  <!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->