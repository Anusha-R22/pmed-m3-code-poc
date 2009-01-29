<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Schedule"

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

nClinicalTrialId = request.querystring("clinicaltrialid")

Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")

sQuery = "Select clinicaltrialname "
sQuery = sQuery & "from clinicaltrial "
sQuery = sQuery & "where clinicaltrial.clinicaltrialid = " & nClinicalTrialId

rsResult.open sQuery,Connect
WriteGroupHeader "Study" , rsResult("clinicaltrialname")
rsResult.close

sQuery = "Select  studyvisit.* "
sQuery = sQuery & "from studyvisit "
sQuery = sQuery & "where studyvisit.clinicaltrialid = " & nClinicalTrialId
sQuery = sQuery & " order by visitorder "

rsResult.open sQuery,Connect

do until rsResult.eof 

	 WriteGroupHeader "Visit",  rsResult("visitcode")
	 WriteTableStart
	 WriteTableRowStart
	 WriteCell "Code:"
	 WriteCell rsResult("visitcode")
	 WriteTableRowEnd 
	 WriteTableRowStart
	 WriteCell "Name:"
	 WriteCell rsResult("visitname")
	 WriteTableRowEnd 
	 WriteTableRowStart
	 WriteCell "Repeats:"
	 WriteCell fVisitRepeats (rsResult("Repeating") )
	 WriteTableRowEnd 

	 WriteTableEnd
	 
	 sQuery = "Select crfpagecode,crftitle,repeating,eFormUse "
	 sQuery = sQuery & "from crfpage,studyvisitcrfpage "
	 sQuery = sQuery & "where crfpage.clinicaltrialid = " & nClinicalTrialId
	 sQuery = sQuery & "  and studyvisitcrfpage.visitid = " & rsResult("visitid")
	 sQuery = sQuery & " and  crfpage.clinicaltrialid = studyvisitcrfpage.clinicaltrialid "
	 sQuery = sQuery & " and  crfpage.crfpageid = studyvisitcrfpage.crfpageid "
	 sQuery = sQuery & " order by crfpageorder "

	 rsResult1.open sQuery,Connect
	 if rsResult1.eof then
	 		WritePara "<B>This visit contains no eForms.</b>"
	 else
	 		sVisits =  "<b>eForms: "
	 	 do until rsResult1.eof 
			sVisits = sVisits &  rsResult1("crfpagecode")
			if cint(rsResult1("repeating")) = 1 then
				 sVisits = sVisits &  fFormRepeats(rsResult1("repeating"))
			end if
			if cint(rsResult1("eFormUse")) = 1 then
				 sVisits = sVisits &  fVisitEForm(rsResult1("eformuse"))
			end if
			sVisits = sVisits & ", "
				rsResult1.movenext
			loop
			WritePara left(sVisits,len(sVisits) - 2) & "</b>"
	 end if
	 rsResult1.close

 
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