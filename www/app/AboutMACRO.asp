<%@ Language=VBScript %>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		AboutMACRO.asp
'	Author: 	
'==================================================================================================
'	Revisions:
'	ic	30/07/2004 added error handling
'==================================================================================================
%>
<!-- #include file="Global.asp" -->
<%
	dim oIO
	
	on error resume next
	
	'create i/o object instance
	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	response.write  oIO.GetAboutHTML(request.querystring("ModuleName"),request.querystring("Version")) 
	set oIO = Nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"AboutMACRO.asp oIo.GetAboutHTML()",Array(request.querystring("ModuleName"), _
		request.querystring("Version")),true)
	end if
%>