<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=StringUtilities.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Get_System_Messages.asp
'   Author:     Richard Meinesz, 2002
'   Purpose:    Used by TrialOffice for the purpose of retreving system messages from the server
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


		set oSysDataXfer = CreateObject("MACROSysDataXfer30.SysDataXfer")
		
		vSysMessages = oSysDataXfer.GetSystemMessages(Request.Form("site"),session("DataBaseDesc"),vMessage, ,Request.Form("confirmationids"))
		
		Response.Write vSysMessages
		Response.Write "<msg>"
		Response.Write vMessage
		
	end if


%>


<!--#include file=CloseDataConnection.txt-->