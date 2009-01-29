<% 
if sReportType = 0 then
' HTML
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html xmlns:v="urn:schemas-microsoft-com:vml">
<head>
<title><%= sReportTitle %></title>
<link href="report.css" rel="stylesheet" type="text/css">
<% 
if sIncludeVML = 0 then
%>
</head>
<body >
<%
else
%>
<style>
v\:* { behavior: url(#default#VML); }
</style>
</head>
<body >
<v:shapetype id="Bar" fillcolor="blue" strokecolor="blue" strokeweight="0.5pt">
<v:fill type="gradient" />
</v:shapetype>
<%
end if
%>
<DIV style="display:auto;position:relative;LEFT: 5px; TOP: 0px" valign="top" width="95%">
<table width="100%" border="0" cellpadding="0" cellspacing="0" >
<tr  valign="middle">
<td align="right"><img id="HHomeFrameHLTab1" src="curve_mo.gif" alt="" border="0" /></td>
<td align="left" width="100%" height="10"><img  src="title_mo.gif" alt="" border="0" width="100%" height="14px" /></td>
</tr>
</table>
</div>
<DIV style="display:auto;position:relative;LEFT: 5px; TOP: 0px;width:100%;border-bottom:#3366CC 1px solid;border-left:#3366CC 1px solid;border-right:#3366CC 1px solid;" valign="top">
<table width="100%" class="ReportHeaderTable">
<tr>
<td width="320px" class="ReportHeaderTitle" ><%= sReportTitle %></td>
<td width="320px">
<table width="100%" >
<tr>
<td >Requested by:</td>
<td ><%= sUserName %></td>
<td  bgcolor="#ffffff" ><a media="screen" onclick="window.print()" style="cursor:hand;" >Print</a></td>
</tr>
<tr>
<td></td>
<td><%= sReportDate %></td>
<%
dim sReferringURL
dim nLastPage
' check if used 'back'
if Request.QueryString("back") <> "" then
	' use stored page
	nLastPage=session("LastPage").Count
	' as long as > 0
	if nLastPage > 0 then
		' have moved back so remove page from collection
		if session("LastPage").Exists(("k"&nLastPage)) then
			' if exists
			session("LastPage").Remove(("k"&nLastPage))
		'else
			' debug
			'Response.Write "<td>page not in collection - k" & nLastPage & ", Count - " & nLastPage & "</td>"
		end if
		' debug
		'Response.Write "<td>use stored page removed - k" & nLastPage & "</td>"
		' decrement page count
		nLastPage=session("LastPage").Count
		if session("LastPage").Exists(("k"&nLastPage)) then
			sReferringURL = session("LastPage").Item(("k"&nLastPage))
		else
			' lost place somehow in session variable - take back to home page & clear other items
			session("LastPage").RemoveAll
			sReferringURL = "home.asp"
		end if
	else
		sReferringURL = "home.asp"
	end if
	' debug
	'if session("LastPage").Exists("k1") then
	'	Response.Write "<td>k1=" & session("LastPage").Item(("k1")) & "</td>"
	'end if
	'if session("LastPage").Exists("k2") then
	'	Response.Write "<td>k2=" & session("LastPage").Item(("k2")) & "</td>"
	'end if
	' remove page from collection
	'session("LastPage").Remove(("k"&nLastPage))
else
	' use and store last referer page
	sReferringURL = Request.ServerVariables("HTTP_REFERER")
	if sReferringURL = "" then
		sReferringURL = "home.asp"
	end if
	' fix for web when thinks referrer is mainborder.asp
	if instr(1,sReferringURL,"MainBorder.asp",0) > 0 then
		sReferringURL="home.asp"
	end if
	' store - unless home.asp (as already have)
	nLastPage=session("LastPage").Count
	if session("LastPage").Exists(("k"&(nLastPage+1))) then
		session("LastPage").item(("k"&(nLastPage+1)))=sReferringURL
		'Response.Write "<td>overwrite - k" & (nLastPage+1) & "</td>"
	else
		session("LastPage").add ("k"&(nLastPage+1)),sReferringURL
		'Response.Write "<td>add - k" & (nLastPage+1) & "</td>"
		' debug
		'Response.Write "<td>add - k" & (nLastPage+1) & "</td>"
		'if session("LastPage").Exists("k1") then
		'	Response.Write "<td>k=" & (nLastPage+1) & session("LastPage").Item("k"&(nLastPage+1)) & "</td>"
		'end if
	end if
end if
' add "back" to querystring
if instr(1,sReferringURL,"?",0) > 0 then
	sReferringURL = sReferringURL & "&back=1"
else
	sReferringURL = sReferringURL & "?back=1"
end if
%>
<td bgcolor="#ffffff"><a  media="screen" href="<%= sReferringURL %>" style="cursor:hand;">Close</a></td>
</tr>
<%
if sPrintDatabase = 1 then
%>
<tr>
<td>Database:</td>
<td><%= sDatabase %></td>
</tr>
<%
end if
%>
</table>
</td>
<td>&nbsp;</td>
</tr>
</table>
<% 
elseif sReportType = 1 then
' Excel
Response.Buffer = false
Response.ContentType = "application/vnd.ms-excel"
%>

<% 
elseif sReportType = 2 then
' CSV
Response.ContentType = "text"

end if
%>
