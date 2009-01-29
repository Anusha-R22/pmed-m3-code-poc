<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "CTC schemes"

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

Set rsResult1 = CreateObject("ADODB.Recordset")
Set rsResult2 = CreateObject("ADODB.Recordset")
Set rsResult3 = CreateObject("ADODB.Recordset")


sQuery = "Select * from CTCScheme  "
if request.querystring("CTCSchemeCode") > "" then
sQuery = sQuery & " where CTCSchemeCode = '" & request.querystring("CTCSchemeCode") & "' "
end if
sQuery = sQuery & "order by CTCSchemeCode "

rsResult1.open sQuery,Connect

do until rsResult1.eof 

WriteGroupHeader "CTC Scheme", rsResult1("CTCSchemeDescription")

sQuery =  "Select c.*, cg.ClinicalTestGroupDescription, ct.ClinicalTestDescription, ct.Unit from CTC c, ClinicalTest ct, ClinicalTestGroup cg  "
sQuery = sQuery & "where CTCSchemeCode ='" & rsResult1("CTCSchemeCode") & "' "
sQuery = sQuery & " and c.clinicaltestcode = ct.clinicaltestcode "
sQuery = sQuery & "  and ct.clinicaltestgroupcode = cg.clinicaltestgroupcode "
sQuery = sQuery & "order by ClinicalTestGroupDescription,ClinicalTestDescription, CTCGrade "

rsResult2.open sQuery,Connect


WriteTableStart
WriteTableRowStart
WriteHeaderCell "Test Group"
WriteHeaderCell "Clinical Test"
WriteHeaderCell "Grade"
WriteHeaderCell "Minimum"
WriteHeaderCell "Maximum"
WriteHeaderCell "Unit"
WriteTableRowEnd

do until rsResult2.eof 
WriteTableRowStart
WriteCell rsResult2("ClinicalTestGroupDescription")
WriteCell rsResult2("ClinicalTestDescription")
WriteCell rsResult2("CTCGrade")
sCTCMin = ">=" & rsResult2("CTCMin").Value 
if not isnull(rsResult2("CTCMinType")) then
	if cint(rsResult2("CTCMinType").Value) = 1 then
		 sCTCMin = sCTCMin &  " x LLN"
	end if
	if cint(rsResult2("CTCMinType").Value) = 2 then
		 sCTCMin = sCTCMin & " x ULN"
	end if
end if
WriteCell  sCTCMin
sCTCMax = "<=" & rsResult2("CTCMax").Value  
if not isnull(rsResult2("CTCMaxType")) then
	if cint(rsResult2("CTCMaxType").Value) = 1 then
		 sCTCMax = sCTCMax & " x LLN"
	end if
	if cint(rsResult2("CTCMaxType").Value) = 2 then
		 sCTCMax = sCTCMax & " x ULN"
	end if
end if
WriteCell  sCTCMax
WriteCell  rsResult2("Unit").Value 
WriteTableRowEnd
rsResult2.movenext
loop

rsResult2.Close
WriteTableEnd



rsResult1.movenext
loop



rsResult1.Close
set RsResult1 = Nothing
set RsResult2 = Nothing


'*************************
' Footer block
'*************************

%>
<!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->