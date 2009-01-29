<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = false%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
'	copyright:		InferMed Ltd 2001. all rights reserved
'	file:			MIMessageBot.asp
'	date:			28/06/2001
'	author:			ilc
'	purpose:		receives a discrepancy id in the querystring. retrieves and displays the
'					audit information for this discrepancy
'	version:		0.1
'==================================================================================================
'	amendments:		
'	ic 29/07/2002	changed dll reference for 3.0
'	ic	22/11/2002	changed www directory structure
'	ic 28/06/2004 added error handling
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<%
dim oIo
dim nMiMessageType
dim sStudy
dim sId
dim sSrc
dim sSite

	on error resume next

	nMiMessageType = Request.QueryString("mimessagetype")
	sStudy = Request.QueryString("study")
	sId = Request.QueryString("id")
	sSrc = Request.QueryString("src")
	sSite = Request.QueryString("site")

%>
	<html>
	<head>
	<link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css">
	<!-- #include file="HandleBrowserEvents.asp" -->
	</head>
	<body>

<%

	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	Response.Write(oIo.GetMIMessageAuditHTML(session("ssUser"),nMiMessageType,sStudy,sSite,sId,sSrc))
	set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"MIMessageBot.asp oIo.GetMIMessageAuditHTML()",Array(nMiMessageType,sStudy,sSite,sId,sSrc),false)
	end if

%>
	</body>
	</html>