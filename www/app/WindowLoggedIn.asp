<%
'==================================================================================================
'	copyright:		InferMed Ltd 2001. all rights reserved
'	file:			checkLoggedIn.asp
'	date:			17/05/2001
'	author:			ilc
'	purpose:		transmits a page to the browser that will load login.asp in the top
'					frame after a session timeout occurs
'	amendments:		
'==================================================================================================
%>
<%
if (session("ssUser") = "") then
%>	
	<script language="javascript">
	var sUrl="login.asp?exp=true";
	sUrl+=(top.sUserDb!=undefined)?"&"+top.fnGetAppUrl(false):"";
	top.location.href=sUrl;
	</script>
<%
	Response.End 
End If
%>
