<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		appHeaderLh.asp
'	Author: 	I Curtis
'	Purpose: 	application header frame
'==================================================================================================
'	Revisions:
'	ic	22/11/2002	changed www directory structure
'	ic 28/06/2004 added error handling
'==================================================================================================
%>
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="HandleBrowserEvents.asp" -->
<!-- #include file="Global.asp" -->
<%
	dim oIo
	
	on error resume next
	
	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	Response.Write(oIo.GetAppMenuHeaderLhHTML(Session("ssUser")))
	set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"AppHeaderLh.asp oIo.GetAppMenuHeaderLhHTML()",Array(),true)
	end if
%>