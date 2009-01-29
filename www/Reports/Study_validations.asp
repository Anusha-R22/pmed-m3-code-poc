<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Validations"

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

sQuery = "Select clinicaltrialname "
sQuery = sQuery & "from clinicaltrial "
sQuery = sQuery & "where clinicaltrial.clinicaltrialid = " & nClinicalTrialId

rsResult.open sQuery,Connect
WriteGroupHeader "Study" , rsResult("clinicaltrialname")
rsResult.close

sQuery = "Select  dataitemcode, dataitemname, validationid, validationtypename , dataitemvalidation, validationmessage  "
sQuery = sQuery & "from dataitem, validationtype, dataitemvalidation "
sQuery = sQuery & "where dataitem.clinicaltrialid = " & nClinicalTrialId
sQuery = sQuery & "  and dataitem.clinicaltrialid = dataitemvalidation.clinicaltrialid "
sQuery = sQuery & "  and dataitem.dataitemid = dataitemvalidation.dataitemid "
sQuery = sQuery & "  and dataitemvalidation.validationtypeid = validationtype.validationtypeid "
sQuery = sQuery & " order by dataitemcode, validationid "

rsResult.open sQuery,Connect

	 WriteTableStart
	 WriteTableRowStart
	 WriteHeaderCell "Question"
	 WriteHeaderCell ""
	 WriteHeaderCell "Validation type"
	 WriteHeaderCell "Validation expression"
	 WriteHeaderCell "Validation message"
	 WriteTableRowEnd

do until rsResult.eof 

	 sDataItemCode = rsResult("dataitemcode")
	 nCount = 1
	 do until sDataItemCode <> rsResult("dataitemcode") 
	 		WriteTableRowStart
			if nCount = 1 then
			  		WriteCell sDataItemCode
			else
					WriteCell ""
			end if
			WriteCell rsResult("validationid")
			WriteCell rsResult("validationtypename")
			WriteCell rsResult("dataitemvalidation")
			WriteCell rsResult("validationmessage")
	 		WriteTableRowEnd
	 		rsResult.movenext
			nCount = nCount + 1
			if rsResult.eof then exit do
	 loop
loop

	 WriteTableEnd	

rsResult.Close
set RsResult = Nothing

'*************************
' Footer block
'*************************

%>
<!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->