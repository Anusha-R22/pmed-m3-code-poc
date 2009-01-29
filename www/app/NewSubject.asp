<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 2000. All Rights Reserved
'	File:		NewSubject.asp
'	Authors: 	i curtis
'	Purpose: 	Allow the user to create a new subject, originally contained in SubjectList.asp
'				querystring parameters:
'					fltDb: selected database
'==================================================================================================
'	Revisions:
' DPH 08/11/2002 Removed User object references
'	ic	22/11/2002	changed www directory structure
'	ic 28/06/2004 added error handling
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Unlock.asp" -->
<!-- #include file="HandleBrowserEvents.asp" -->
<!-- #include file="Global.asp" -->
<%
dim oIo

	on error resume next

	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	Response.Write(oIo.GetNewSubjectHTML(Session("ssUser")))
	set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"NewSubject.asp oIo.GetNewSubjectHTML()",Array(),true)
	end if
%>
