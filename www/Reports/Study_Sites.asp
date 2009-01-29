<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Study sites"

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

Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")

sQuery = "Select clinicaltrialid,clinicaltrialname "
sQuery = sQuery & "from clinicaltrial  "
sQuery = sQuery & " where clinicaltrialid > 0 "
if request.querystring("clinicaltrialid") > "" then
sQuery = sQuery & "  and clinicaltrialid = " & request.querystring("clinicaltrialid")
end if
sQuery = sQuery & "order by clinicaltrialname "

rsResult1.open sQuery,Connect


do until rsResult1.eof 


WriteGroupHeader "Study", rsResult1("clinicaltrialname") 



sQuery = "Select trialsite,studyversion "
sQuery = sQuery & "from trialsite  "
sQuery = sQuery & " where clinicaltrialid = '" & rsResult1("clinicaltrialid") & "' "
sQuery = sQuery & "order by  trialsite "

rsResult.open sQuery,Connect


WriteTableStart
WriteTableRowStart
WriteHeaderCell "Site"
WriteHeaderCell "Study Version"
WriteTableRowEnd



do until rsResult.eof 
WriteTableRowStart
WriteCell rsResult("trialsite") 
WriteCell rsResult("studyversion") 

WriteTableRowEnd
rsResult.movenext
loop

WriteTableEnd

rsResult.Close


rsResult1.movenext
loop

rsResult1.Close
set RsResult1 = Nothing
set RsResult = Nothing

'*************************
' Footer block
'*************************

%>
<!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->