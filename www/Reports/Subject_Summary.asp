<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Subject summary"

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
sQuery = sQuery & "order by clinicaltrialname "

rsResult2.open sQuery,Connect

do until rsResult2.eof 

WriteGroupHeader "Study", rsResult2("clinicaltrialname") 

sQuery =  "Select distinct studyvisit.visitorder,studyvisit.visitname,studyvisit.visitid, visitinstance.visitcyclenumber  "
sQuery = sQuery & "from studyvisit, visitinstance "
sQuery = sQuery & " where studyvisit.clinicaltrialid = visitinstance.clinicaltrialid"
sQuery = sQuery & "   and studyvisit.visitid= visitinstance.visitid "
sQuery = sQuery & "   and studyvisit.clinicaltrialid =  " & rsResult2("clinicaltrialid")
sQuery = sQuery & " and " & replace(replace(sStudySiteSQL, "clinicaltrial.", "visitinstance."), "trialsite.", "visitinstance." )

if request.querystring("trialsite") > "" then
	 sQuery = sQuery & "   and visitinstance.trialsite = '" & request.querystring("trialsite") & "' "
end if

sQuery = sQuery & " order by visitorder,studyvisit.visitid,visitcyclenumber "

rsResult1.open sQuery,Connect



WriteTableStart
WriteTableRowStart
WriteHeaderCell "Subject"
WriteHeaderCell "Site"
WriteHeaderLink "Link"

' RS 10/06/2003: Equivalent columns are empty for now
'WriteHeaderLink ""		' Link: disc column
'WriteHeaderLink ""		' audit trail column

nCount = 1

do until rsResult1.eof 

	 sVisitId(nCount) = rsResult1("VisitId")
	 sVisitCycleNumber(nCount) = rsResult1("VisitCycleNumber")

	 if cint(rsResult1("VisitCycleNumber")) > 1 then
	 	 WriteHeaderCell rsResult1("visitname") & "<br>(" & rsResult1("VisitCycleNumber")& ")"
		else
		 WriteHeaderCell rsResult1("visitname")
	end if

	rsResult1.movenext
	nCount = nCount + 1
loop

WriteTableRowEnd

sQuery =  "Select trialsubject.trialsite,trialsubject.personid, localidentifier1,"
sQuery = sQuery & "visitorder,studyvisit.visitid,visitcyclenumber,"
sQuery = sQuery & " visitstatus,vi.lockstatus,vi.discrepancystatus,vi.sdvstatus,vi.notestatus  "
sQuery = sQuery & "from trialsubject, visitinstance vi,studyvisit "
sQuery = sQuery & " where trialsubject.clinicaltrialid = vi.clinicaltrialid"
sQuery = sQuery & "   and trialsubject.trialsite= vi.trialsite "
sQuery = sQuery & "   and trialsubject.personid= vi.personid "
sQuery = sQuery & "   and trialsubject.clinicaltrialid =  " & rsResult2("clinicaltrialid")
sQuery = sQuery & "   and vi.clinicaltrialid = studyvisit.clinicaltrialid"
sQuery = sQuery & "   and vi.visitid= studyvisit.visitid "
sQuery = sQuery & " and " & replace(replace(sStudySiteSQL, "clinicaltrial.", "vi."), "trialsite.", "vi." )


if request.querystring("trialsite") > "" then
	 sQuery = sQuery & "   and vi.trialsite = '" & request.querystring("trialsite") & "' "
end if

sQuery = sQuery & " order by trialsubject.trialsite,trialsubject.personid,visitorder,studyvisit.visitid,visitcyclenumber "

rsResult.open sQuery,Connect

sPersonId = -1

do until rsResult.eof 
	'REM 14/09/04 - Added check for TrialSite in case person Id matched but data was from different site
	 if not ((cint(sPersonId) = cint(rsResult("PersonId"))) and (sTrialSite = rsResult("trialsite")))  then
		 ' this is different xsubject
	 	 WriteTableRowEnd
		 sPersonId = rsResult("PersonId")
		 sTrialSite = rsResult("trialsite") 
	   WriteTableRowStart
	 	 nCount = 1
		 WriteCell fIdOrLabel(rsResult("PersonId"),rsResult("LocalIdentifier1"))
		 WriteCell rsResult("trialsite")
		 
		 sCell = "<a href=""data.asp?ClinicalTrialId=" & rsResult2("clinicaltrialid") & "&trialsite=" & rsResult("trialsite") & "&personid=" & rsResult("personid") & """>"
	   sCell = sCell & "Data"
	   sCell = sCell & "</a>"
	   WriteCell sCell

		' WriteCell ""  'disc
		' WriteCell ""  'audit trail
	end if

	do while cint(sVisitId(nCount)) <> cint(rsResult("VisitId")) or cint(sVisitCycleNumber(nCount)) <> cint(rsResult("VisitCycleNumber")) 
	 	 nCount = nCount + 1
	 	 WriteCell ""
	'writecell cint(sVisitId(nCount))  & "," &  cint(rsResult("VisitId")) & "," &  cint(sVisitCycleNumber(nCount)) & "," & cint(rsResult("VisitCycleNumber")) 
if ncount > 900 then
exit do
end if
	loop
if ncount > 900 then
exit do
end if

	sImages = oIo.RtnStatusImagesHTML(cint(rsResult("visitstatus")), false, cint(rsResult("lockstatus")), false, cint(rsResult("sdvstatus")), cint(rsResult("discrepancystatus")), false, false, 0)
	WriteCentredCell sImages
	' WriteCentredCell fFormStatusImage(rsResult("visitstatus"))
	nCount = nCount + 1

	rsResult.movenext
loop

WriteTableRowEnd
WriteTableEnd

rsResult1.Close
rsResult.Close
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