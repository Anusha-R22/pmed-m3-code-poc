<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Check_System_Messages.asp
'   Author:     Richard Meinesz, 2002
'   Purpose:    Used to check if there are any system messages to download from the server to a site
'
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'	ic 24/05/2004 added variable checking
'-----------------------------------------------------------------------------------------------'

dim oSysDataXfer 

	on error resume next

	if err.number <> 0 then

		response.write "ERROR1:" & err.number
		Response.End 
	
	else

		'validate site
		if (not fnValidateSite(Request.Form("site"))) then
			Response.Write("ERROR:The site '" & Request.Form("site") & "' does not exist")
			Response.End 
		end if
		'validate username
		if not fnValidateUsername(Request.Form("username")) then
			Response.Write("ERROR:The userid '" & Request.Form("username") & "' is not valid")
			Response.End
		end if

		set oSysDataXfer = CreateObject("MACROSysDataXfer30.SysDataXfer")
		
		if oSysDataXfer.CheckSystemMessages(Request.Form("username"),Request.Form("site"),session("DataBaseDesc"),vMessage) = false then
			Response.Write "No system messages to download"
		
		else
			Response.Write "There are system messages to download"
		
		end if
		
	end if

%>

<!--#include file=CloseDataConnection.txt-->