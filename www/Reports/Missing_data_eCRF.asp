<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Missing data - eForm summary"

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
	%> <!--#include file="report_close.asp" -->
	<%
	Response.end()
end if

'*************************
' Content block
'*************************
Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")

sQuery = "Select distinct c.clinicaltrialid,clinicaltrialname "
sQuery = sQuery & "from clinicaltrial c  "
sQuery = sQuery & " where c.clinicaltrialid in (" & sPermittedStudyList & ") "
if request.querystring("ClinicalTrialId") > "" then
	sQuery = sQuery & "and c.clinicaltrialid = " & request.querystring("ClinicalTrialId")
end if
sQuery = sQuery & "order by clinicaltrialname "

rsResult1.open sQuery,Connect

do until rsResult1.eof 

WriteGroupHeader "Study", rsResult1("clinicaltrialname") 

WriteTableStart
WriteTableRowStart
WriteHeaderCell "Site"
WriteHeaderCell "Subject"
WriteHeaderCell "Visit"
WriteHeaderCell "eCRF"
WriteHeaderCell "Number of missing values"

WriteTableRowEnd

sQuery =  "Select trialsubject.trialsite,trialsubject.personid,localidentifier1, studyvisit.visitid,visitorder," _
		& "visitname,crfpage.crfpageid,crfpageorder,crftitle, count(ResponseTaskId) as NumberOfValues  "
sQuery = sQuery & "from trialsubject, studyvisit,crfpage,crfpageinstance,dataitemresponse "
sQuery = sQuery & " where trialsubject.clinicaltrialid = crfpageinstance.clinicaltrialid"
sQuery = sQuery & "   and trialsubject.trialsite = crfpageinstance.trialsite "
sQuery = sQuery & "   and trialsubject.personid = crfpageinstance.personid"
sQuery = sQuery & "   and dataitemresponse.clinicaltrialid = crfpageinstance.clinicaltrialid"
sQuery = sQuery & "   and dataitemresponse.trialsite = crfpageinstance.trialsite "
sQuery = sQuery & "   and dataitemresponse.personid = crfpageinstance.personid"
sQuery = sQuery & "   and dataitemresponse.crfpagetaskid = crfpageinstance.crfpagetaskid"
sQuery = sQuery & "   and trialsubject.clinicaltrialid = crfpage.clinicaltrialid"
sQuery = sQuery & "   and crfpageinstance.crfpageid = crfpage.crfpageid"
sQuery = sQuery & "   and trialsubject.clinicaltrialid = studyvisit.clinicaltrialid"
sQuery = sQuery & "   and crfpageinstance.visitid = studyvisit.visitid"
sQuery = sQuery & "   and dataitemresponse.responsestatus = 10 "
sQuery = sQuery & "   and trialsubject.clinicaltrialid =  " & rsResult1("clinicaltrialid")
sQuery = sQuery & " and " & replace(replace(sStudySiteSQL, "clinicaltrial.", "trialsubject."), "trialsite.", "trialsubject." )
if request.querystring("trialsite") > "" then
	 sQuery = sQuery & "   and trialsubject.trialsite = '" & request.querystring("trialsite") & "' "
end if
if request.querystring("personid") > "" then
	 sQuery = sQuery & "   and trialsubject.personid = " & request.querystring("personid")
end if

sQuery = sQuery & " group by trialsubject.trialsite,trialsubject.personid,localidentifier1,studyvisit.visitid,visitorder, visitname,crfpage.crfpageid,crfpageorder,crftitle "

rsResult.open sQuery,Connect

sTrialSite = ""
sPersonId = ""
do until rsResult.eof 

	' DPH - 15/03/2004 - use WriteLink to avoid problms in CSV/Excel 
	sQueryString = "status=10&ClinicalTrialId=" & rsResult1("clinicaltrialid") & "&trialsite=" & rsResult("trialsite")
	WriteLink rsResult("trialsite"), "data.asp", sQueryString
	'sCell = "<a href=""data.asp?status=10&ClinicalTrialId=" & rsResult1("clinicaltrialid") & "&trialsite=" & rsResult("trialsite") & """>"
	'sCell = sCell & rsResult("trialsite")
	'sCell = sCell & "</a>"
	'WriteCell sCell
	sQueryString = "status=10&ClinicalTrialId=" & rsResult1("clinicaltrialid") & "&trialsite=" & rsResult("trialsite") & "&personid=" & rsResult("personid")
	WriteLink fIdOrLabel(rsResult("personid"),rsResult("localidentifier1")), "data.asp", sQueryString
	'sCell = "<a href=""data.asp?status=10&ClinicalTrialId=" & rsResult1("clinicaltrialid") & "&trialsite=" & rsResult("trialsite") & "&personid=" & rsResult("personid") & """>"
	'sCell = sCell & fIdOrLabel(rsResult("personid"),rsResult("localidentifier1"))
	'sCell = sCell & "</a>"
	'WriteCell sCell
	sQueryString = "status=10&ClinicalTrialId=" & rsResult1("clinicaltrialid") & "&trialsite=" & rsResult("trialsite") & "&personid=" & rsResult("personid") & "&visitid=" & rsResult("visitid")
	WriteLink rsResult("visitname"), "data.asp", sQueryString
	'sCell = "<a href=""data.asp?status=10&ClinicalTrialId=" & rsResult1("clinicaltrialid") & "&trialsite=" & rsResult("trialsite") & "&personid=" & rsResult("personid") & "&visitid=" & rsResult("visitid") & """>"
	'sCell = sCell & rsResult("visitname")
	'sCell = sCell & "</a>"
	'WriteCell sCell
	sQueryString = "status=10&ClinicalTrialId=" & rsResult1("clinicaltrialid") & "&trialsite=" & rsResult("trialsite") & "&personid=" & rsResult("personid") & "&visitid=" & rsResult("visitid") & "&crfpageid=" & rsResult("crfpageid")
	WriteLink rsResult("crftitle"), "data.asp", sQueryString
	'sCell = "<a href=""data.asp?status=10&ClinicalTrialId=" & rsResult1("clinicaltrialid") & "&trialsite=" & rsResult("trialsite") & "&personid=" & rsResult("personid") & "&visitid=" & rsResult("visitid") & "&crfpageid=" & rsResult("crfpageid") & """>"
	'sCell = sCell & rsResult("crftitle")
	'sCell = sCell & "</a>"
	'WriteCell sCell
	WriteCentredCell rsResult("NumberOfValues")
	WriteTableRowEnd
	rsResult.movenext

loop


WriteTableEnd


rsResult.Close
rsResult1.movenext
loop

rsResult1.close
set RsResult = Nothing
set RsResult1 = Nothing

'*************************
' Footer block
'*************************

%>
  <!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->