<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "eForm summary"

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

Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")
Set rsResultTrial = CreateObject("ADODB.Recordset")

sSelection = "<option value="""">All sites</option>"
sQuery = "Select trialsite from trialsite where clinicaltrialid in (" & sPermittedStudyList & ")"
rsResult1.open sQuery,Connect

do until rsResult1.eof
	 if request.querystring("trialsite") = rsResult1("TrialSite") then
	 		sSelected = "selected"		
		else
			sSelected = ""
	end if
	sSelection = sSelection & "<option value=""" & rsResult1("TrialSite") & """ " & sSelected & " >" & rsResult1("TrialSite") & "</option>"
	rsResult1.movenext
loop

' DPH - 15/03/2004 - don't draw table unless a display report 
if sReportType = 0 then
	response.write "<table style=""font-family:verdana,arial;"" width=""100%""><tr>"
	response.write "<td width=""50px""></td><td>"
	response.write "<form action=""schedule_form_summary.asp"" method=""get"">"
	response.write "<input type=""hidden"" name=""clinicaltrialid"" value=""" & nClinicalTrialId & """>"
	response.write "<input type=""hidden"" name=""fltDb"" value=""" & request.querystring("fltDb") & """>"
	response.write "Site: <select style=""font-family:verdana,arial;""  name=""trialsite"" "
	response.write ">" & sSelection & "</select>"
	response.write "<select style=""font-family:verdana,arial;""  name=""status"" >"
	response.write "<option value=""""" 

	if request.querystring("status") = "" then
		 response.write " selected "
	end if

	response.write ">All eCRFs</option>"
	response.write "<option value=""0""" 

	if request.querystring("status") = "0" then
		 response.write " selected "
	end if

	response.write ">OK</option>"
	response.write "<option value=""10""" 

	if request.querystring("status") = "10" then
		 response.write " selected "
	end if

	response.write ">Missing data</option>"
	response.write "<option value=""30""" 

	if request.querystring("status") = "30" then
		 response.write " selected "
	end if

	response.write ">Warnings</option>"
	response.write "</select>"
	response.write "<button type=""submit"">Go</button>"
	response.write "</form>"
	response.write "<td width=""50px""></td></tr></table>"
end if

rsResult1.Close

' RS 12/06/2003: Display a separate table for each trial
sQuery = "Select clinicaltrialid,clinicaltrialname from clinicaltrial where clinicaltrialid in (" & sPermittedStudyList & ")"
rsResultTrial.open sQuery,Connect

do while not rsResultTrial.eof

	WriteGroupHeader "Study", rsResultTrial("clinicaltrialname") 
	
	nClinicaltrialID = rsResultTrial("clinicaltrialid")

	sQuery =  "Select distinct studyvisit.visitorder,studyvisit.visitname,studyvisit.visitid, visitinstance.visitcyclenumber  "
	sQuery = sQuery & "from studyvisit, visitinstance "
	sQuery = sQuery & " where studyvisit.clinicaltrialid = visitinstance.clinicaltrialid"
	sQuery = sQuery & "   and studyvisit.visitid= visitinstance.visitid "
	sQuery = sQuery & "   and studyvisit.clinicaltrialid = " & rsResultTrial("clinicaltrialid")
	
	'WAS: sQuery = sQuery & "   and studyvisit.clinicaltrialid in (" & sPermittedStudyList & ")"
	' dph 12/02/2004 - study/site permissions
	sQuery = sQuery & " and " & replace(replace(sStudySiteSQL, "clinicaltrial.", "studyvisit."), "trialsite.", "visitinstance." )

	if request.querystring("trialsite") > "" then
		 sQuery = sQuery & " and visitinstance.trialsite = '" & request.querystring("trialsite") & "' "
	end if
	
	sQuery = sQuery & " order by visitorder,studyvisit.visitid,visitcyclenumber "
	
	rsResult1.open sQuery,Connect
	
	if not rsResult1.eof then	

		WriteTableStart
		WriteTableRowStart
		WriteHeaderCell "eCRF"
		
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
		
		sQuery =  "Select distinct crfpage.crfpageorder,crfpage.crfpageid,crfpage.crfpagecode, crfpage.crftitle,"
		sQuery = sQuery & "visitorder,crfpageinstance.visitid,visitcyclenumber,"
		sQuery = sQuery & " count(crfpagetaskid) as NumberOfPages  "
		sQuery = sQuery & "from crfpage, crfpageinstance,studyvisit "
		sQuery = sQuery & " where crfpage.clinicaltrialid = crfpageinstance.clinicaltrialid"
		sQuery = sQuery & "   and crfpage.crfpageid= crfpageinstance.crfpageid "
		sQuery = sQuery & "   and crfpageinstance.clinicaltrialid = studyvisit.clinicaltrialid"
		sQuery = sQuery & "   and crfpageinstance.visitid= studyvisit.visitid "
		sQuery = sQuery & "   and crfpage.clinicaltrialid = " & rsResultTrial("clinicaltrialid") & " "
		'WAS: sQuery = sQuery & "   and crfpage.clinicaltrialid in (" & sPermittedStudyList & ")"
		sQuery = sQuery & "   and crfpageinstance.crfpagestatus >= 0 "
		
		' dph 12/02/2004 - study/site permissions
		sQuery = sQuery & " and " & replace(replace(sStudySiteSQL, "clinicaltrial.", "crfpage."), "trialsite.", "crfpageinstance." )

		if request.querystring("trialsite") > "" then
			 sQuery = sQuery & "   and crfpageinstance.trialsite = '" & request.querystring("trialsite") & "' "
		end if
		
		if request.querystring("status") > "" then
			 sQuery = sQuery & "   and crfpageinstance.crfpagestatus = " & request.querystring("status")
		end if
		
		sQuery = sQuery & " group by crfpageorder,crfpage.crfpageid,crfpagecode,crftitle,visitorder,crfpageinstance.visitid,visitcyclenumber "
		sQuery = sQuery & " order by crfpageorder,crfpage.crfpageid,crfpagecode,crftitle,visitorder,crfpageinstance.visitid,visitcyclenumber "
		
		rsResult.open sQuery,Connect
		
		sCRFCode = ""
		
		do until rsResult.eof 
			 if sCRFCode <> rsResult("CRFPageCode")  then
				 WriteTableRowEnd
				 sCRFCode = rsResult("CRFPageCode") 
			   WriteTableRowStart
				 nCount = 1
				 WriteCell rsResult("CRFTitle")
			end if
		
			do while cint(sVisitId(nCount)) <> cint(rsResult("VisitId")) or cint(sVisitCycleNumber(nCount)) <> cint(rsResult("VisitCycleNumber")) 
				 nCount = nCount + 1
				 WriteCell ""
			loop
		
			' DPH - 15/03/2004 - use WriteLink to avoid problms in CSV/Excel 
			sQueryString = "ClinicalTrialId=" & nClinicalTrialId 
			sQueryString = sQueryString & "&TrialSite=" & request.querystring("trialsite") 
			sQueryString = sQueryString & "&Status=" & request.querystring("status") 
			sQueryString = sQueryString & "&VisitId=" & rsResult("VisitId") 
			sQueryString = sQueryString & "&CRFPageId=" & rsResult("CRFPageId") 
			sQueryString = sQueryString & "&fltDB=" & request.querystring("fltDb")
			WriteLink rsResult("NumberofPages"), "eCRFList.asp", sQueryString
			'sCell = "<a href=""eCRFList.asp?ClinicalTrialId=" & nClinicalTrialId 
			'sCell = sCell & "&TrialSite=" & request.querystring("trialsite") 
			'sCell = sCell & "&Status=" & request.querystring("status") 
			'sCell = sCell & "&VisitId=" & rsResult("VisitId") 
			'sCell = sCell & "&CRFPageId=" & rsResult("CRFPageId") 
			'sCell = sCell & "&fltDB=" & request.querystring("fltDb")
			'sCell = sCell & """>" & rsResult("NumberofPages") & "</a>"
		
			'WriteCentredCell sCell
			nCount = nCount + 1
		
			rsResult.movenext
		loop
		
		WriteTableRowEnd
		WriteTableEnd
		
		rsResult.Close
	end if
	rsResult1.Close
	

	rsResultTrial.Movenext
loop

set RsResult1 = Nothing
set RsResult = Nothing
set rsResultTrial = Nothing


'*************************
' Footer block
'*************************

%>
  <!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->