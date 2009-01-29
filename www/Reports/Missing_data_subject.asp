<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Missing data - subject summary"

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

' RS 26MAY2005
' Rewrite: Select all results in a single select and filter out sites that the user is not permitted to see
' Use the oUser object to get the list of permitted sites for each study

'	 Create objects
	set oUser = server.CreateObject("MACROUSERBS30.MACROUser")
	if session("ssUser") > "" then				 'WWW
		oUser.setstate(session("UserObject"))
	else																	 'Windows
	  oUser.setstatehex(session("UserObject"))
	end if
	sSiteList = oUser.DataLists.StudiesSitesWhereSQL("clinicaltrialid", "trialsite")
	set oUser = Nothing

newQuery = "select b.clinicaltrialname, a.trialsite TSITE, a.personid PID, trialsubject.localidentifier1 LID, a.* from" & vbNewLine & _
"(" & vbNewLine & _
"select clinicaltrialid, trialsite, personid, count(responsetaskid) as NUMBEROFVALUES, 'TOTAL' as CountType" & vbNewLine & _
"from dataitemresponse" & vbNewLine & _
"where XXX" & vbNewLine & _
"group by clinicaltrialid,trialsite, personid" & vbNewLine & _
"union all" & vbNewLine & _
"select clinicaltrialid, trialsite, personid, count(responsetaskid) as NUMBEROFVALUES, 'MISS' as CountType" & vbNewLine & _
"from dataitemresponse" & vbNewLine & _
"where XXX" & vbNewLine & _
"and responsestatus = 10" & vbNewLine & _
"group by clinicaltrialid,trialsite, personid" & vbNewLine & _
") a , trialsubject, clinicaltrial b" & vbNewLine & _
"where trialsubject.clinicaltrialid = a.clinicaltrialid" & vbNewLine & _
"and b.clinicaltrialid = a.clinicaltrialid" & vbNewLine & _
"and trialsubject.trialsite = a.trialsite" & vbNewLine & _
"and trialsubject.personid = a.personid" & vbNewLine & _
"order by a.clinicaltrialid, a.trialsite, a.personid"

newQuery = replace(newQuery,"XXX",sSiteList)


Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")

sQuery = "Select distinct c.clinicaltrialid,clinicaltrialname "
sQuery = sQuery & "from clinicaltrial c  "
sQuery = sQuery & " where c.clinicaltrialid in (" & sPermittedStudyList & ") "

if request.querystring("clinicaltrialid") > "" then
	 sQuery = sQuery & "   and c.clinicaltrialid = '" & request.querystring("clinicaltrialid") & "' "
end if

sQuery = sQuery & "order by clinicaltrialname "

'response.write "Permitted Studies: " & sPermittedStudyList

'response.write "Time: " & Timer & "<BR>" & sQuery
response.Flush()

rsResult1.open sQuery,Connect

'do until rsResult1.eof 

'WriteGroupHeader "Study", rsResult1("clinicaltrialname") 

'WriteTableStart
'WriteTableRowStart
'WriteHeaderCell "Site"
'WriteHeaderCell "Subject"
'WriteHeaderCell "Number of missing values"
'WriteHeaderCell "Total"
'WriteHeaderCell "%"

'WriteTableRowEnd

' This query returns two rows per subject. The first row contains the number of missing dataitems in the NumberofValues column
' The second row returns the total number of dataitems for the subject. The difference is used to calculate a percentage.

' dph 09/01/2004 : changed sql to work with both SQL server & Oracle
' RS 11/08/2003: Modified query to return data for one study only. Each study is treated separately, see outer loop)
sQuery = "Select trialsubject.trialsite ""TSITE"",  trialsubject.personid ""PID"", localidentifier1 ""LID"", count(responsetaskid) as NumberOfValues, 'TOTAL' as CountType "
sQuery = sQuery & "from trialsubject, dataitemresponse"
'sQuery = sQuery & " where " & replace(replace(sStudySiteSQL, "clinicaltrial.", "trialsubject."), "trialsite.", "trialsubject." )
sQuery = sQuery & " where trialsubject.clinicaltrialid = " & rsResult1("clinicaltrialid")
sQuery = sQuery & " and dataitemresponse.clinicaltrialid = trialsubject.clinicaltrialid "
sQuery = sQuery & " and dataitemresponse.trialsite = trialsubject.trialsite "
sQuery = sQuery & " and dataitemresponse.personid = trialsubject.personid "
if request.querystring("trialsite") > "" then
	 sQuery = sQuery & "   and trialsubject.trialsite = '" & request.querystring("trialsite") & "' "
end if
' dph 12/02/2004 - study/site permissions
sQuery = sQuery & " and " & replace(replace(sStudySiteSQL, "clinicaltrial.", "trialsubject."), "trialsite.", "trialsubject." )
sQuery = sQuery & " group by trialsubject.trialsite,  trialsubject.personid, localidentifier1 "
sQuery = sQuery & " union "
sQuery = sQuery & "Select trialsubject.trialsite ""TSITE"",  trialsubject.personid ""PID"", localidentifier1 ""LID"", count(responsetaskid) as NumberOfValues, 'MISS' as CountType "
sQuery = sQuery & "from trialsubject, dataitemresponse"
'sQuery = sQuery & " where " & replace(replace(sStudySiteSQL, "clinicaltrial.", "trialsubject."), "trialsite.", "trialsubject." )
sQuery = sQuery & " where trialsubject.clinicaltrialid = " & rsResult1("clinicaltrialid")
sQuery = sQuery & " and dataitemresponse.responsestatus = 10 "
sQuery = sQuery & " and dataitemresponse.clinicaltrialid = trialsubject.clinicaltrialid "
sQuery = sQuery & " and dataitemresponse.trialsite = trialsubject.trialsite "
sQuery = sQuery & " and dataitemresponse.personid = trialsubject.personid "
if request.querystring("trialsite") > "" then
	 sQuery = sQuery & "   and trialsubject.trialsite = '" & request.querystring("trialsite") & "' "
end if
' dph 12/02/2004 - study/site permissions
sQuery = sQuery & " and " & replace(replace(sStudySiteSQL, "clinicaltrial.", "trialsubject."), "trialsite.", "trialsubject." )
sQuery = sQuery & " group by trialsubject.trialsite,  trialsubject.personid, localidentifier1"
sQuery = sQuery & " order by TSITE, PID, LID "

sQuery = newQuery

'response.write "Time: " & Timer & "<BR>" & sQuery
'response.Flush()
rsResult.open sQuery,Connect
'response.write "Time: " & Timer
response.Flush()
sTrialSite = ""
nPersonId = -1
curStudy = ""
do while not rsResult.eof 

	response.Flush()
	
	if rsResult("clinicaltrialname") <> curStudy then
		if curStudy<>"" then
			WriteTableEnd
		end if
		WriteGroupHeader "Study", rsResult("clinicaltrialname") 
		
		WriteTableStart
		WriteTableRowStart
		WriteHeaderCell "Site"
		WriteHeaderCell "Subject"
		WriteHeaderCell "Number of missing values"
		WriteHeaderCell "Total"
		WriteHeaderCell "%"
	
		WriteTableRowEnd
		curStudy = rsResult("clinicaltrialname") 
	end if
			


	' initialise data
	nMissing = -1
	nTotal = -1

	sTrialSite = rsResult("TSITE")
	nPersonId = cint(rsResult("PID"))
	sLID = ReturnLID(rsResult("LID"))
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
		' if same trialsite and personid and localident
		if sTrialSite = rsResult("TSITE") and nPersonId = cint(rsResult("PID")) _
			and sLID = ReturnLID(rsResult("LID")) then
				' get value
				if rsResult("CountType") = "TOTAL" Then
					nTotal = clng(rsResult("NumberOfValues"))
				elseif rsResult("CountType") = "MISS" Then
					nMissing = clng(rsResult("NumberOfValues"))
				end if
				' movenext - (to next group)
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

	'show data
	' DPH - 15/03/2004 - use WriteLink to avoid problms in CSV/Excel 
	sQueryString = "ClinicalTrialId=" & rsResult1("clinicaltrialid") & "&trialsite=" & sTrialSite
	WriteLink sTrialSite, "Missing_data_eCRF.asp", sQueryString
	'sCell = "<a href=""Missing_data_eCRF.asp?ClinicalTrialId=" & rsResult1("clinicaltrialid") & "&trialsite=" & sTrialSite & """>"
	'sCell = sCell & sTrialSite
	'sCell = sCell & "</a>"
	'WriteCell sCell
	 
	sQueryString = "ClinicalTrialId=" & rsResult1("clinicaltrialid") & "&trialsite=" & sTrialSite & "&personid=" & nPersonId
	WriteLink fIdOrLabel(nPersonId,sLID), "Missing_data_eCRF.asp", sQueryString
	'sCell = "<a href=""Missing_data_eCRF.asp?ClinicalTrialId=" & rsResult1("clinicaltrialid") & "&trialsite=" & sTrialSite & "&personid=" & nPersonId & """>"
	'sCell = sCell & fIdOrLabel(nPersonId,sLID)
	'sCell = sCell & "</a>"
	'WriteCell sCell
	WriteCentredCell nMissing
	WriteCentredCell nTotal
	if nTotal > 0 then
		WriteCell cInt( cint(nMissing) / cint(nTotal) * 100)
	end if
	WriteTableRowEnd
loop

WriteTableEnd

rsResult.Close
rsResult1.movenext
'loop

rsResult1.close
set RsResult = Nothing
set RsResult1 = Nothing

'*************************
' Footer block
'*************************
%>
  <%
function ReturnLID(sVal)
	if isnull(sVal) then
		ReturnLID = ""
	else
		ReturnLID = sVal
	end if
end function
%>
  <!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->

