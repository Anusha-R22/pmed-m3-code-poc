<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Questions"

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

sQuery = "Select datatypeid,datatypename "
sQuery = sQuery & "from datatype "
sQuery = sQuery & " order by datatypeid "

rsResult.open sQuery,Connect

do until rsResult.eof 
WriteGroupheader "Type", rsResult("datatypename")

sQuery = "Select  dataitemcode, dataitemname, derivation,MACROOnly  "
sQuery = sQuery & "from dataitem "
sQuery = sQuery & "where dataitem.clinicaltrialid = " & nClinicalTrialId
sQuery = sQuery & "  and dataitem.datatype = " & rsResult("datatypeid")
sQuery = sQuery & " order by dataitemcode "


rsResult1.open sQuery,Connect

if rsResult1.eof then
	 WritePara "<b>No questions.</b>"
else

	 WriteTableStart
	 WriteTableRowStart
	 WriteHeaderCell "Question code"
	 WriteHeaderCell "Name"
	 WriteTableRowEnd

	 do until rsResult1.eof 
	 		WriteTableRowStart
			WriteCell rsResult1("dataitemcode")
			WriteCell rsResult1("dataitemname")
			if rsResult1("derivation") > "" then
						WriteCell "(derived)"
			else
						WriteCell ""
			end if
			if cint(rsResult1("MACROOnly")) = 1 then
						WriteCell "(MACRO Only)"
			else
						WriteCell ""
			end if
	 		WriteTableRowEnd
	 		rsResult1.movenext
	 loop
	 WriteTableEnd		
end if
rsResult1.close
rsResult.movenext
loop


rsResult.Close
set RsResult = Nothing

'*************************
' Footer block
'*************************

%>
<!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->