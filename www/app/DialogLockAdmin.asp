<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		LockAdmin.asp
'	Author: 	I Curtis
'	Purpose: 	
'==================================================================================================
'	Revisions:
'	ic 28/06/2004 added error handling
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="DialogLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<%
dim oIo
dim sForm
dim sRtn

	on error resume next

	sForm = Request.Form()
	
	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	if (sForm <> "") then
		call oIo.DeleteDBLocks(session("ssUser"),sForm)
		if err.number <> 0 then 
			set oIo = nothing
			call fnError(err.number,err.description,"DialogLockAdmin.asp oIo.DeleteDBLocks()",Array(sForm),true)
		end if
	end if
	sRtn = oIo.GetDBLockAdminHTML(session("ssUser"))
	set oIo = nothing
	if err.number <> 0 then 
		call fnError(err.number,err.description,"DialogLockAdmin.asp oIo.GetDBLockAdminHTML()",Array(),true)
	end if
	
	Response.Write(sRtn)
%>
	