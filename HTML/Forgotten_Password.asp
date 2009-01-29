<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=StringUtilities.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Forgotten_Password.asp
'   Author:     Richard Meinesz, 2002
'   Purpose:    Used by Data Management Login to check and retrieve new passwords from the server
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'	ic 24/05/2004 added variable checking
'-----------------------------------------------------------------------------------------------'
dim oSysDataXfer 
dim vSysMessages
dim vMessages

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
			Response.Write("ERROR:The username '" & Request.Form("username") & "' is not valid")
			Response.End 
		end if
		'validate password
		'TA 10/09/2004: don't valid ti if it wasn't sent
		if Request.Form("password") <> "" then
			if not fnValidatePassword(Request.Form("password")) then
				Response.Write("ERROR:The password is not valid")
				Response.End 
			end if
		end if
		

		set oSysDataXfer = CreateObject("MACROSysDataXfer30.SysDataXfer")
		
		vSysMessages = oSysDataXfer.GetNewPassword(Request.Form("username"),Request.Form("password"),session("DataBaseDesc"),Request.Form("site"),vMessage,Request.Form("confirmationids"))
		
		Response.Write vSysMessages
		Response.Write "<msg>"
		Response.Write vMessage
		
	end if


%>


<!--#include file=CloseDataConnection.txt-->
