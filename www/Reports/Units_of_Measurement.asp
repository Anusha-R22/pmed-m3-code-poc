<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Units of measurement"

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


sQuery = "Select * from UnitClasses  "
sQuery = sQuery & " order by unitclass "

rsResult1.open sQuery,Connect


do until rsResult1.eof 

WriteGroupHeader "Class",  rsResult1("unitclass")

response.write "<table>"
response.write "<tr valign=top>"
response.write "<td width=""50%"">"

sQuery = "Select * from Units  "
sQuery = sQuery & " where unitclass = '" & rsResult1("unitclass") & "' "
sQuery = sQuery & " order by unit "

rsResult2.open sQuery,Connect

WriteTableStart

WriteTableRowStart
WriteHeaderCell "Unit"
WriteTableRowEnd

do until rsResult2.eof 
WriteTableRowStart
WriteCell rsResult2("unit") 
WriteTableRowEnd
rsResult2.movenext
loop

WriteTableEnd
rsResult2.close


response.write "</td>"
response.write "<td>"

sQuery = "Select * from UnitConversionFactors "
sQuery = sQuery & " where unitclass = '" & rsResult1("unitclass") & "' "
sQuery = sQuery & " order by fromunit,tounit "

rsResult2.open sQuery,Connect

WriteTableStart

WriteTableRowStart
WriteHeaderCell  "Convert from"
WriteHeaderCell  "To"
WriteHeaderCell  "Conversion factor"
WriteTableRowEnd

do until rsResult2.eof 
WriteTableRowStart
WriteCell rsResult2("fromunit")
WriteCell rsResult2("tounit")
WriteCell rsResult2("conversionfactor")
WriteTableRowEnd
rsResult2.movenext
loop

WriteTableEnd
rsResult2.close

response.write "</td>"
response.write "</tr>"
response.write "</table>"





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