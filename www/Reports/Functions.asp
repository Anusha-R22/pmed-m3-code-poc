<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Functions"

sIncludeVML = 0 'Don't include VML styles

%>
<!--#include file="report_initialise.asp" -->
<!--#include file="report_open_security_database.asp" -->
<!--#include file="report_functions.asp" -->
<!--#include file="report_header.asp" -->
<%

'*************************
' Content block
'*************************

Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")


sQuery = "Select distinct m.MACROModule "
sQuery = sQuery & "from functionmodule m  "
sQuery = sQuery & "order by MACROModule "

rsResult1.open sQuery,Connect



do until rsResult1.eof 


select case rsResult1("MACROModule") 
case "DE"
		 WriteGroupHeader "Module", "Data Entry"
case "DR"
		 WriteGroupHeader "Module", "Data Review"
case "DV"
		 WriteGroupHeader "Module", "Create Data Views"
case "LM"
		 WriteGroupHeader "Module", "Library Management"
case "QM"
		 WriteGroupHeader "Module", "Query Module"
case "SD"
		 WriteGroupHeader "Module", "Study Definition"
case "SM"
		 WriteGroupHeader "Module", "System Management"
case else
		 WriteGroupHeader "Module", rsResult1("MACROModule") 
end select 

sQuery = "Select  f.functioncode, f.macrofunction "
sQuery = sQuery & "from functionmodule m, macrofunction f  "
sQuery = sQuery & "where m.functioncode = f.functioncode "
sQuery = sQuery & " and m.macromodule = '" & rsResult1("MACROModule") & "' "
sQuery = sQuery & "order by f.functioncode "

rsResult.open sQuery,Connect


WriteTableStart
WriteTableRowStart
WriteHeaderCell "Function Code"
WriteHeaderCell "Function"
WriteTableRowEnd



do until rsResult.eof 
WriteTableRowStart
WriteCell rsResult("functioncode") 
WriteCell rsResult("macrofunction") 

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