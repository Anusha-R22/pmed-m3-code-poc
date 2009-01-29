<%@ Language=VBScript %>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>

<%
' dph 12/03/2004 store page in session variable
set session("LastPage")=server.CreateObject("Scripting.Dictionary")

'ic/mlm 05/11/2003 CRM 248
' Get serialised user object
if session("ssUser") > "" then
	'web
	session("UserObject") = session("ssUser")
else
	'windows
	if request.form > "" then
		session("UserObject") = request.form
	end if
end if	
%>
<html xmlns:v="urn:schemas-microsoft-com:vml">
	<HEAD>
	<title><%=sAPP_TITLE%></title>
	<script language="javascript">
					function OpenReport(sReport,bNewWin)
					{
					 if (reporttype[0].checked) { sURL = sReport + '&reporttype=0'};
					 if (reporttype[1].checked) { sURL = sReport + '&reporttype=1'};
					 if (reporttype[2].checked) { sURL = sReport + '&reporttype=2'};

					 if(bNewWin==true)
					 {
						var oWin = window.open(sURL,"newrepwin","location=no,menubar=yes,toolbar=no,scrollbars=yes")
						oWin.focus();
					 }
					 else
					 {
						window.navigate(sURL);
					 }
					}
	</script>
<link href="report.css" rel="stylesheet" type="text/css">
<style>
v\:* { behavior: url(#default#VML); }
</style>
</head>
<!--#include file="report_initialise.asp" -->
<!--#include file="report_open_macro_database.asp" -->
<!--#include file="home_functions.asp" -->
<body>
<div style="height:100%;width:100%;">


<%
if request.querystring("displayallreports") = "yes" then
%>
<!--#include file="home_role_macrouser.asp" -->
<%
else
select case lcase(sUserRole)
case "macrouser"
%>
<!--#include file="home_role_macrouser.asp" -->
<%
case "inv"
%>
<!--#include file="home_role_nurse.asp" -->
<%
case "nurse"
%>
<!--#include file="home_role_nurse.asp" -->
<%
case "datamgr"
%>
<!--#include file="home_role_datamgr.asp" -->
<%
case "monitor"
%>
<!--#include file="home_role_monitor.asp" -->
<%
case "siteadmin"
%>
<!--#include file="home_role_macrouser.asp" -->
<%
case else
%>
<!--#include file="home_role_macrouser.asp" -->
<%
end select
end if
%>

</div>
</body>
</html>
	
<!--#include file="report_close.asp" -->