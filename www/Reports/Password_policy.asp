<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Password Policy"

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

sQuery = "Select  * "
sQuery = sQuery & "from MACROPassword "

rsResult.open sQuery,Connect


WriteTableStart
WriteTableRowStart
WriteCell "Minimum length:"
WriteCell rsResult("minlength")
WriteTableRowEnd
WriteTableRowStart
WriteCell "Maximum length:"
WriteCell rsResult("maxlength")
WriteTableRowEnd
WriteTableRowStart
WriteCell "Expiry period:"
if cint(rsResult("expiryperiod")) = 0 then
	 WriteCell "Not enforced"
else
	 WriteCell rsResult("expiryperiod") & " days"
end if
WriteTableRowEnd
WriteTableRowStart
WriteCell "Check against previous passwords:"
if cint(rsResult("passwordhistory")) = 0 then
	 WriteCell "Not enforced"
else
	 WriteCell rsResult("passwordhistory") & " passwords remembered"
end if
WriteTableRowEnd
WriteTableRowStart
WriteCell "Account lockout:"
if cint(rsResult("passwordretries")) = 0 then
	 WriteCell "Not enforced"
else
	 WriteCell rsResult("passwordretries") & " retries allowed before account is locked out"
end if
WriteTableRowEnd
WriteTableRowStart
WriteCell "Enforce mixed case:"
if cint(rsResult("enforcemixedcase")) = 0 then
	 WriteCell "No"
else
	 WriteCell "Yes"
end if
WriteTableRowEnd
WriteTableRowStart
WriteCell "Enforce at least 1 numerical digit:"
if cint(rsResult("enforcedigit")) = 0 then
	 WriteCell "No"
else
	 WriteCell "Yes"
end if
WriteTableRowEnd
WriteTableRowStart
WriteCell "Allow repeating of characters:"
if cint(rsResult("allowrepeatchars")) = 0 then
	 WriteCell "No"
else
	 WriteCell "Yes"
end if
WriteTableRowEnd
WriteTableRowStart
WriteCell "Allow portion of username:"
if cint(rsResult("allowusername")) = 0 then
	 WriteCell "No"
else
	 WriteCell "Yes"
end if

WriteTableRowEnd


rsResult.Close
set RsResult = Nothing
'*************************
' Footer block
'*************************

%>
<!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->