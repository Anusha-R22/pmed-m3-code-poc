<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		appMenuLh.asp
'	Author: 	I Curtis
'	Purpose: 	application lh menu frame
'				querystring parameters:
'==================================================================================================
'	Revisions:
'	ic	22/11/2002	changed www directory structure
'	ic  10/01/2003	combined load and save call
'	ic 28/06/2004 added error handling
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="HandleBrowserEvents.asp" -->
<!-- #include file="Global.asp" -->
<%
	dim oIo
	dim sForm
	dim sUser
	dim bSplit
	
	on error resume next
	
	sUser = Session("ssUser")
	sForm = Request.Form()

	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	Response.Write(oIo.SaveAndLoadAppMenuLh(sUser,bSplit,sForm))	
	set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"AppMenuLh.asp oIo.SaveAndLoadAppMenuLh()",Array(bSplit,sForm),true)
	end if
	
	Session("ssUser") = sUser
	Session("sbSplit") = bSplit
%>
