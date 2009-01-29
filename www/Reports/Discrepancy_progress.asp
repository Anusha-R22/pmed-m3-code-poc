<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
sReportTitle = "Site recruitment"
sUserName = "Andrew Newbigging"
sReportDate = date() & " " & time()


%>
<!--#include file="report_header.asp" -->
<%

sDatabaseType = "SQL SERVER"
sSQLServer = "A_NEWBIGGING_LA"
sSQLDatabase = "msde_macro"
sSQLUserId = "sa"
sSQLPassword = "macrotm"

Set Connect = CreateObject("ADODB.Connection")

Set rsResult1 = CreateObject("ADODB.Recordset")
Set rsResult2 = CreateObject("ADODB.Recordset")
Set rsResult3 = CreateObject("ADODB.Recordset")
Connect.ConnectionString = "PROVIDER=SQLOLEDB;DATA SOURCE=" & sSQLServer & ";DATABASE=" & sSQLDatabase & ";USER ID=" & sSQLUserId & ";PASSWORD=" & sSQLPassword & ";" 

Connect.Open


Dim sQuery 

sQuery = sQuery & "Select MIMessageCreatedDate,count(MIMEssageId) as Recruitment  "
sQuery = sQuery & "from  V_MIMessage  group by MIMessageCreatedDate  "

rsResult1.open sQuery,Connect

response.write "<table width=""400px"" height = ""200px"">"


nTotal = 0
nMax = 0
do until rsResult1.eof 
if rsResult1("Recruitment") > nMax then
	 nMax = rsResult1("Recruitment")
end if
nTotal = nTotal + 1
rsResult1.movenext
loop


response.write "<tr>"
response.write "<td width= ""80px"">Number of subjects</td>"
response.write "<td height=""90%"" valign=""bottom"" align=""center""><v:polyline points="""
rsResult1.movefirst
nCount = 0
do until rsResult1.eof 
nCount = nCount + 1

if nTotal < 20 then
	 nWidth = 20
else
		nWidth = nCount/nTotal*400
end if
if ncount > 1 then
response.write ","
end if
response.write nCount*10 & "px," & rsResult1("Recruitment") * 10 & "px"
rsResult1.movenext

loop
response.write """></polyline>"
response.write "</td>"
response.write "</tr>"

response.write "<tr><td></td>"
rsResult1.movefirst
do until rsResult1.eof 
response.write "<td height=""10%""  align=""center"">"
response.write rsResult1("MIMessageCreatedDate")
response.write "</td>"
response.write "<td height=""10%""  align=""center"">"
response.write rsResult1("Recruitment")
response.write "</td>"
rsResult1.movenext
loop
response.write "</tr>"

response.write "</table>"

rsResult1.Close
set RsResult1 = Nothing

Connect.Close
Set Connect = Nothing



%>



<!--#include file="report_footer.asp" -->