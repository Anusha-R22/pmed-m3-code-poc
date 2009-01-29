<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Question details"

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

sQuery = "Select  dataitem.*,datatypename "
sQuery = sQuery & "from dataitem,datatype "
sQuery = sQuery & "where dataitem.clinicaltrialid = " & nClinicalTrialId
sQuery = sQuery & " and  dataitem.datatype = datatype.datatypeid "
sQuery = sQuery & " order by dataitemcode "

rsResult.open sQuery,Connect

do until rsResult.eof 

	 WriteGroupHeader "Question",  rsResult("dataitemcode")
	 WriteTableStart
	 WriteTableRowStart
	 WriteCell "Code:"
	 WriteCell rsResult("dataitemcode")
	 WriteTableRowEnd 
	 WriteTableRowStart
	 WriteCell "Name:"
	 WriteCell rsResult("dataitemname")
	 WriteTableRowEnd 
	 WriteTableRowStart
	 WriteCell "Export code:"
	 WriteCell rsResult("exportname")
	 WriteTableRowEnd 
	 WriteTableRowStart
	 WriteCell ""
	 WriteCell fMACROOnly (rsResult("MACROOnly") )
	 WriteTableRowEnd 
	 WriteTableRowStart
	 WriteCell "Type:"
	 WriteCell rsResult("datatypename")
	 WriteTableRowEnd 
	 if cint(rsResult("datatype")) = 0 then 'text
	 	 WriteTableRowStart
	 	 WriteCell "Case:"
	 	 WriteCell fCase (rsResult("dataitemcase") )
	 	 WriteTableRowEnd 
	 end if 
	 if cint(rsResult("datatype")) = 6 then 'lab test
	 	 WriteTableRowStart
	 	 WriteCell "Clinical test:"
	 	 WriteCell rsResult("clinicaltestcode") 
	 	 WriteTableRowEnd 
	 end if 
	 if rsResult("Unitofmeasurement") > "" then
	 	 WriteTableRowStart
		 WriteCell "Unit of measurement:"
		 WriteCell rsResult("Unitofmeasurement")
		 WriteTableRowEnd 
	 end if
	 if rsResult("dataitemformat") > "" then
	 		WriteTableRowStart
	 		WriteCell "Format:"
	 		WriteCell rsResult("dataitemformat")
	 		WriteTableRowEnd 
	 		if cint(rsResult("datatype")) = 4 and cint(rsResult("dataitemcase") ) = 1 then 'date/time
	 	 		WriteTableRowStart
	 	 		WriteCell ""
	 	 		WriteCell "(Allow partial dates)"
	 	 		WriteTableRowEnd 
	 		end if 
	 end if
	 if cint(rsResult("datatype")) = 0 or cint(rsResult("datatype")) = 1 then 'Text or category
	 		WriteTableRowStart
	 		WriteCell "Maximum length:"
	 		WriteCell rsResult("dataitemlength")
	 		WriteTableRowEnd 
	 end if
	 if rsResult("derivation") > "" then
	 		WriteTableRowStart
	 		WriteCell "Derivation:"
	 		WriteCell rsResult("derivation")
	 		WriteTableRowEnd 
		end if
	 if rsResult("dataitemhelptext") > "" then
	 		WriteTableRowStart
	 		WriteCell "User help text:"
	 		WriteCell rsResult("dataitemhelptext")
	 		WriteTableRowEnd 
	 end if
	 if rsResult("description") > "" then
	 		WriteTableRowStart
	 		WriteCell "Metadata description:"
	 		WriteCell rsResult("description")
	 		WriteTableRowEnd 
	 end if
	 WriteTableEnd
	 
	 sQuery = "Select validationtypename,dataitemvalidation.* "
	 sQuery = sQuery & "from dataitemvalidation,validationtype "
	 sQuery = sQuery & "where dataitemvalidation.clinicaltrialid = " & nClinicalTrialId
	 sQuery = sQuery & "  and dataitemvalidation.dataitemid = " & rsResult("dataitemid")
	 sQuery = sQuery & " and  dataitemvalidation.validationtypeid = validationtype.validationtypeid "
	 sQuery = sQuery & " order by validationid "

	 rsResult1.open sQuery,Connect
	 if rsResult1.eof then
	 		' Write nothing
	 else
	 WriteTableStart
	 WriteTableRowStart
	 WriteHeaderCell ""
	 WriteHeaderCell "Validation type"
	 WriteHeaderCell "Validation expression"
	 WriteHeaderCell "Validation message"
	 WriteTableRowEnd

	 	 do until rsResult1.eof 
		 		WriteTableRowStart
	 		WriteCell rsResult1("validationid")
			WriteCell rsResult1("validationtypename")
			WriteCell rsResult1("dataitemvalidation")
			WriteCell rsResult1("validationmessage")
	 		WriteTableRowEnd
				rsResult1.movenext
			loop
			WriteTableEnd
	 end if
	 rsResult1.close

	 sQuery = "Select valuedata.* "
	 sQuery = sQuery & "from valuedata "
	 sQuery = sQuery & "where valuedata.clinicaltrialid = " & nClinicalTrialId
	 sQuery = sQuery & "  and valuedata.dataitemid = " & rsResult("dataitemid")
	 sQuery = sQuery & " order by valueorder "

	 rsResult1.open sQuery,Connect
	 if rsResult1.eof then
	 		' Write nothing
	 else
	 WriteTableStart
	 WriteTableRowStart
	 WriteHeaderCell "Value code"
	 WriteHeaderCell "Description"
	 WriteTableRowEnd

	 	 do until rsResult1.eof 
		 		WriteTableRowStart
			WriteCell rsResult1("valuecode")
			WriteCell rsResult1("itemvalue")
			WriteCell fActive( rsResult1("active") )
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
set RsResult1 = Nothing

'*************************
' Footer block
'*************************

%>
<!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->