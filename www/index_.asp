<%@ Language=VBScript %>
<%
'==================================================================================================
'	copyright:		InferMed Ltd 2001. all rights reserved
'	file:			index.asp
'	date:			16/05/2001
'	author:			ilc
'	purpose:		redirects browser to ssl encrypted login page
'	version:		0.1
'	amendments:		
'==================================================================================================
%>
<%
dim sDomain
dim sUrl

'get domain
sDomain = "http"
if Application("asUSESSL") = true then sDomain = sDomain & "s"
sDomain = sDomain & "://" & Request.ServerVariables("SERVER_NAME")

'get this script path and replace the script name with the login page
sUrl = Request.ServerVariables("URL")
sUrl = replace(sUrl,"index_.asp","app/index.asp")

Response.Redirect(sDomain & sUrl) 
%>
