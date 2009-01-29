<%
'==================================================================================================
'	copyright:		InferMed Ltd 2001. all rights reserved
'	file:			checkSSL.asp
'	date:			16/05/2001
'	author:			ilc
'	purpose:		redirects browser to application entry page if request didnt come via ssl
'	version:		0.1
'	amendments:		
'==================================================================================================
%>
<%
'if not ssl connection redirect to entry page
'if Request.ServerVariables("HTTPS") <> "ON" then Response.Redirect("http://www." & Request.ServerVariables("SERVER_NAME") & "/")
%>

