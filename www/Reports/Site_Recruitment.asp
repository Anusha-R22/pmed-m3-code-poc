<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Site recruitment"

if request.querystring("chart") > "" then
		sIncludeVML = 1 'Include VML styles
else
		sIncludeVML = 0 'Don't include VML styles
end if
%>
<!--#include file="report_initialise.asp" -->
<!--#include file="report_open_macro_database.asp" -->
<!--#include file="report_functions.asp" -->
<!--#include file="report_header.asp" -->
<%
' dph 15/03/2004 - disallow VML for csv reports
if sReportType = 2 then
	sIncludeVML = 0 'Don't include VML styles
end if
'*************************
' Content block
'*************************

Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")

sQuery = "Select distinct c.clinicaltrialid,clinicaltrialname "
sQuery = sQuery & "from clinicaltrial c  "
sQuery = sQuery & " where c.clinicaltrialid in (" & sPermittedStudyList & ") "
sQuery = sQuery & "order by clinicaltrialname "

rsResult.open sQuery,Connect

do until rsResult.eof 

WriteGroupHeader "Study", rsResult("clinicaltrialname") 

sQuery = "Select trialsite,count(personid) as Recruitment  "
sQuery = sQuery & "from  trialsubject "
sQuery = sQuery & " where clinicaltrialid = '" & rsResult("clinicaltrialid") & "' "
sQuery = sQuery & " and " & replace(replace(sStudySiteSQL, "clinicaltrial.", "trialsubject."), "trialsite.", "trialsubject." )
sQuery = sQuery & " group by trialsite order by recruitment "

rsResult1.open sQuery,Connect

if sIncludeVML = 1 then

	 response.write "<table width=""400px"" height = ""200px"">"
	 nTotal = 0
	 nMax = 0
	 do until rsResult1.eof 
	 		if cint(rsResult1("Recruitment")) > nMax then
	 			 nMax = rsResult1("Recruitment")
			end if
			nTotal = nTotal + 1
			rsResult1.movenext
	loop
	response.write "<tr>"
	response.write "<td width= ""80px"">Number of subjects</td>"
	rsResult1.movefirst

	nCount = 0
	do until rsResult1.eof 
		 nCount = nCount + 1
		 response.write "<td height=""90%"" valign=""bottom"" align=""center"">"
		 if nTotal < 20 then
	 	 		nWidth = 20
		 else
		 		 nWidth = nCount/nTotal*400
		 end if
		 response.write rsResult1("Recruitment") & "<br>"
		 response.write "<v:rect type=""#Bar"" style=""width:" & nWidth & "px;height:" & cint(rsResult1("Recruitment")) / cint(nMax) * 150 & "px"" >"
		 response.write "</v:rect>"
		 response.write "</td>"
		 rsResult1.movenext
	loop
	response.write "</tr>"
	response.write "<tr><td></td>"
	rsResult1.movefirst
	do until rsResult1.eof 
		 response.write "<td height=""10%""  align=""center"">"
		 response.write rsResult1("TrialSite")
		 response.write "</td>"
		 rsResult1.movenext
	loop
	response.write "</tr>"
	response.write "</table>"

else

		if rsResult1.eof then
			 WritePara "No subjects recruited yet."
		else

		WriteTableStart
		WriteTableRowStart
		WriteHeaderCell "Site"
		WriteHeaderCell "Recruitment"
		WriteTableRowEnd

		do until rsResult1.eof 
			 WriteTableRowStart
			 WriteCell rsResult1("trialsite") 
			 WriteCell rsResult1("recruitment") 
			 WriteTableRowEnd
			 rsResult1.movenext
		loop

		WriteTableEnd

		end if
		
end if

rsResult1.Close



rsResult.movenext
loop



rsResult.Close
set RsResult = Nothing
set RsResult1 = Nothing


'*************************
' Footer block
'*************************

%>
<!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->