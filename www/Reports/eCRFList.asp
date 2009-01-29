<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "eForm List"

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
Dim sVisitCycleNumber(999)
Dim sVisitId(999)

' RS 10/06/2003: Create WWW object
set oIo = server.CreateObject("MACROWWWIO30.clsWWW")


Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")
Set rsResult2 = CreateObject("ADODB.Recordset")

sQuery = "Select distinct c.clinicaltrialid,clinicaltrialname "
sQuery = sQuery & "from clinicaltrial c  "
sQuery = sQuery & " where c.clinicaltrialid in (" & sPermittedStudyList & ") "
if request.querystring("ClinicalTrialId") > "" then
	 sQuery = sQuery & "   and c.ClinicalTrialId = '" & request.querystring("ClinicalTrialId") & "' "
end if
sQuery = sQuery & "order by clinicaltrialname "

rsResult2.open sQuery,Connect

do until rsResult2.eof 

WriteGroupHeader "Study", rsResult2("clinicaltrialname") 

'*************************
' Now get all eForms for this trial
'*************************

sQuery = " select vi.personid, ct.clinicaltrialname,  sv.visitname, cp.crftitle,"
sQuery = sQuery & "        ci.sdvstatus, ci.notestatus, ci.discrepancystatus,"
sQuery = sQuery & "        ci.lockstatus, ci.crfpagestatus, vi.trialsite, ts.localidentifier1"
sQuery = sQuery & "   from crfpage cp, crfpageinstance ci, visitinstance vi, studyvisit sv, clinicaltrial ct, trialsubject ts"
sQuery = sQuery & "  where ct.clinicaltrialid = vi.clinicaltrialid"
sQuery = sQuery & "        and vi.clinicaltrialid = sv.clinicaltrialid"
sQuery = sQuery & "        and vi.visitid = sv.visitid"
sQuery = sQuery & "        and vi.clinicaltrialid = ci.clinicaltrialid"
sQuery = sQuery & "        and vi.trialsite = ci.trialsite"
sQuery = sQuery & "        and vi.personid = ci.personid"
sQuery = sQuery & "        and vi.visitid = ci.visitid"
sQuery = sQuery & "        and ci.clinicaltrialid = cp.clinicaltrialid"
sQuery = sQuery & "        and ci.crfpageid = cp.crfpageid"
sQuery = sQuery & "        and ts.personid = ci.personid"
sQuery = sQuery & "        and ts.clinicaltrialid = ci.clinicaltrialid"
sQuery = sQuery & "        and ts.trialsite = ci.trialsite"

sQuery = sQuery & "        and ct.clinicaltrialid = " & rsResult2("ClinicalTrialID")
if request.querystring("trialsite") > "" then
	 sQuery = sQuery & "   and vi.trialsite = '" & request.querystring("trialsite") & "' "
end if
if request.querystring("visitid") > "" then
	 sQuery = sQuery & "   and vi.visitid = '" & request.querystring("visitid") & "' "
end if
if request.querystring("crfpageid") > "" then
	 sQuery = sQuery & "   and ci.crfpageid = '" & request.querystring("crfpageid") & "' "
end if
' dph - adding missing status clause
if request.querystring("status") > "" then
	 sQuery = sQuery & "   and ci.crfpagestatus = " & request.querystring("status")
end if
' dph 12/02/2004 - study/site permissions
sQuery = sQuery & " and " & replace(replace(sStudySiteSQL, "clinicaltrial.", "ct."), "trialsite.", "vi." )

sQuery = sQuery & " order by ci.trialsite, vi.personid, sv.visitorder"

'response.write sQuery
'response.Flush()

rsResult1.open sQuery,Connect

WriteTableStart
WriteTableRowStart
WriteHeaderCell "Subject"
WriteHeaderCell "Site"
WriteHeaderCell "Visit"
WriteHeaderCell "eForm"
WriteHeaderCell "eForm Status"
WriteTableRowEnd
nCount = 1

do until rsResult1.eof 

	WriteTableRowStart

	sLabel = fIdOrLabel(rsResult1("PersonId"),rsResult1("LocalIdentifier1"))

	WriteCell sLabel
	WriteCell rsResult1("Trialsite")
	WriteCell rsResult1("Visitname")
	WriteCell rsResult1("CRFtitle")

	sImages = oIo.RtnStatusImagesHTML(cint(rsResult1("crfpagestatus")), false, cint(rsResult1("lockstatus")), false, cint(rsResult1("sdvstatus")), cint(rsResult1("discrepancystatus")), false, false, 0)
	WriteCentredCell sImages

	WriteTableRowEnd

	'sCell = "<a href=""data.asp?ClinicalTrialId=" & rsResult2("clinicaltrialid") & "&trialsite=" & rsResult("trialsite") & "&personid=" & rsResult("personid") & """>"
	'sImages = oIo.RtnStatusImagesHTML(cint(rsResult("visitstatus")), false, cint(rsResult("lockstatus")), false, cint(rsResult("sdvstatus")), cint(rsResult("discrepancystatus")), false, false, 0)
	'WriteCentredCell sImages

	rsResult1.MoveNext
loop


WriteTableEnd

rsResult1.Close
rsResult2.movenext
loop

rsResult2.close
set RsResult2 = Nothing
set RsResult1 = Nothing
set RsResult = Nothing
set oIo = Nothing

'*************************
' Footer block
'*************************

%>
  <!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->