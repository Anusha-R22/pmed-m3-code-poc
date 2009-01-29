<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Sites"

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

' DPH 22/07/2004 - allow for a country having no site so it appears in the report
sQuery = "Select site,sitedescription,sitestatus,sitelocation, CountryDescription "
sQuery = sQuery & "from site, MACROCountry "
if sDatabaseType = 1 then
	' sql server
	sQuery = sQuery & " where site.sitecountry *= MACROCountry.CountryId "
else
	' oracle
	sQuery = sQuery & " where MACROCountry.CountryId (+)= site.sitecountry "
end if
sQuery = sQuery & "order by site "

rsResult.open sQuery,Connect

WriteTableStart
WriteTableRowStart
WriteHeaderCell "Site"
WriteHeaderCell "Description"
WriteHeaderCell "Status"
WriteHeaderCell "Location"
WriteHeaderCell "Country"
WriteTableRowEnd



do until rsResult.eof 
WriteTableRowStart
WriteCell rsResult("site")
WriteCell rsResult("sitedescription") 
WriteCell fEnabled( rsResult("sitestatus") )
if not isnull(rsResult("sitelocation")) then
	WriteCell fSiteLocation( rsResult("sitelocation") )
else
	WriteCell ""
end if
WriteCell rsResult("countrydescription") 
WriteTableRowEnd
rsResult.movenext
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