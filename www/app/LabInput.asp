<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		LabInput.asp
'	Author: 	I Curtis
'	Purpose: 	
'==================================================================================================
'	revisions
'	ic 28/06/2004 added error handling
'==================================================================================================
%>
<!-- #include file="checkSSL.asp" -->
<!-- #include file="Global.asp" -->
<%
dim oIo
dim sSite
dim sUser

	on error resume next

	sSite = Request.QueryString("site")
	sUser = Session("ssUser")

	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	Response.Write(oIo.GetEformLabChoiceHTML(sUser,sSite))
	set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"LabInput.asp oIo.GetEformLabChoiceHTML()",Array(sSite),true)
	end if
%>