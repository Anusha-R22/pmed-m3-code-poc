<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "User logins"

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


sQuery = "Select  l.logdatetime,l.logmessage,m.usernamefull,l.location "
sQuery = sQuery & "from MACROUser m, loginlog l "
sQuery = sQuery & "where m.username = l.username "

if request.querystring("username") > "" then
sQuery = sQuery & " and m.username = '" & request.querystring("username") & "' " 
end if 
if request.querystring("failed") > "" then
sQuery = sQuery & " and l.logmessage like 'Login Failed%'"  
end if 

sQuery = sQuery & "order by logdatetime,taskid "

rsResult.open sQuery,Connect


WriteTableStart
WriteTableRowStart
WriteHeaderCell "Date & time"
WriteHeaderCell "User"
WriteHeaderCell "Location"
WriteHeaderCell "Activity"
WriteTableRowEnd



do until rsResult.eof 
WriteTableRowStart
WriteFixedWidthCell cdate(rsResult("logdatetime")) , 150
WriteCell rsResult("usernamefull") 
WriteCell rsResult("location") 
WriteCell rsResult("logmessage")

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