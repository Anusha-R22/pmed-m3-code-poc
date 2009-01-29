<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=StringUtilities.txt-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Write_System_Messages.asp
'   Author:     Richard Meinesz, 2002
'   Purpose:    Used by TrialOffice for the purpose of writing system messages to the server database
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'-----------------------------------------------------------------------------------------------'
dim oSysDataXfer 
dim vMessage
dim vConfirmation

	'on error resume next

	if err.number <> 0 then

		response.write "ERROR1:" & err.number
		Response.End 
	
	else

		set oSysDataXfer = CreateObject("MACROSysDataXfer30.SysDataXfer")

		vMessage = ""
		vConfirmation = oSysDataXfer.WriteSystemMessage(session("DataBaseDesc"),Request.Form("systemmessage"), vMessage)

		Response.Write vConfirmation
		Response.Write "<msg>"
		Response.Write vMessage
		
	end if

%>



<!--#include file=CloseDataConnection.txt-->