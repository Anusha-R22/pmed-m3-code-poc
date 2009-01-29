<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Data"

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

sQuery = "Select distinct c.clinicaltrialid,clinicaltrialname "
sQuery = sQuery & "from clinicaltrial c  "
sQuery = sQuery & " where c.clinicaltrialid in (" & sPermittedStudyList & ") "
if request.querystring("clinicaltrialid") > "" then
	 sQuery = sQuery & "   and c.clinicaltrialid = " & request.querystring("clinicaltrialid")
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
WriteHeaderCell "Question"
WriteHeaderCell "Status"

WriteTableRowEnd

sQuery =  "Select trialsubject.trialsite,trialsubject.personid,localidentifier1, visitorder,visitname,crfpageorder,crftitle, crfpagecyclenumber,dataitemname,responsestatus  "
sQuery = sQuery & "from trialsubject, studyvisit,crfpage,dataitemresponse,dataitem "
sQuery = sQuery & " where dataitemresponse.clinicaltrialid = trialsubject.clinicaltrialid"
sQuery = sQuery & "   and dataitemresponse.trialsite = trialsubject.trialsite "
sQuery = sQuery & "   and dataitemresponse.personid = trialsubject.personid"
sQuery = sQuery & "   and dataitemresponse.clinicaltrialid = crfpage.clinicaltrialid"
sQuery = sQuery & "   and dataitemresponse.crfpageid = crfpage.crfpageid"
sQuery = sQuery & "   and dataitemresponse.clinicaltrialid = studyvisit.clinicaltrialid"
sQuery = sQuery & "   and dataitemresponse.visitid = studyvisit.visitid"
sQuery = sQuery & "   and dataitemresponse.clinicaltrialid = dataitem.clinicaltrialid"
sQuery = sQuery & "   and dataitemresponse.dataitemid = dataitem.dataitemid"
if request.querystring("status") > "" then
	 sQuery = sQuery & "   and dataitemresponse.responsestatus = " & request.querystring("status")
end if
sQuery = sQuery & "   and trialsubject.clinicaltrialid =  " & rsResult1("clinicaltrialid")
sQuery = sQuery & " and " & replace(replace(sStudySiteSQL, "clinicaltrial.", "trialsubject."), "trialsite.", "trialsubject." )
if request.querystring("trialsite") > "" then
	 sQuery = sQuery & "   and trialsubject.trialsite = '" & request.querystring("trialsite") & "' "
end if
if request.querystring("personid") > "" then
	 sQuery = sQuery & "   and trialsubject.personid = " & request.querystring("personid")
end if
if request.querystring("crfpagetaskid") > "" then
	 sQuery = sQuery & "   and dataitemresponse.crfpagetaskid = " & request.querystring("crfpagetaskid")
end if
if request.querystring("crfpageid") > "" then
	 sQuery = sQuery & "   and dataitemresponse.crfpageid = " & request.querystring("crfpageid")
end if
if request.querystring("visitid") > "" then
	 sQuery = sQuery & "   and dataitemresponse.visitid = " & request.querystring("visitid")
end if

sQuery = sQuery & " order by trialsubject.trialsite,trialsubject.personid,localidentifier1,visitorder, visitname,crfpageorder,crftitle,dataitemname "

rsResult.open sQuery,Connect

sTrialSite = ""
sPersonId = ""
do until rsResult.eof 

	 WriteCell rsResult("trialsite")
	 WriteCell fIdOrLabel(rsResult("personid"),rsResult("localidentifier1"))
	 WriteCell rsResult("visitname")
	 if cint(rsResult("crfpagecyclenumber")) = 1 then
		WriteCell rsResult("crftitle")
	 else
		WriteCell rsResult("crftitle") & " (" & rsResult("crfpagecyclenumber") & ")"
	 end if
	 WriteCell rsResult("dataitemname")
	 WriteCell fStatus(rsResult("responsestatus"))
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