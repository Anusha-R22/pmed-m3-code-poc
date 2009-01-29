<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2005 All Rights Reserved
'   File:       Pdu_Next_Message.asp
'   Author:     David Hook, 2005
'   Purpose:    Used to get the next Pdu message
'
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'-----------------------------------------------------------------------------------------------'
dim oSysDataXfer 
dim sPduFileInfo
dim vErrors

	on error resume next

	if err.number <> 0 then

		response.write "ERROR1:" & err.number
		Response.End 
	
	else
		'validate previousmessageid
		if (not fnNumeric(request.querystring("PreviousMessageID"))) then
			Response.Write("ERROR:The PreviousMessageID '" & request.querystring("PreviousMessageID") & "' is not valid")
			Response.End 
		end if
		'validate site
		if (not fnValidateSite(request.querystring("Site"))) then
			Response.Write("ERROR:The site '" & request.querystring("Site") & "' does not exist")
			Response.End 
		end if

		set oSysDataXfer = CreateObject("MACROSysDataXfer30.SysDataXfer")
		
		' check for existence of pdu setting
		if gsSecurePdu <> "" then
			' get pdu folder
			sPduDirectory = gsSecurePdu
		else
			' default to published html folder
			sPduDirectory = gsAppPath
		end if

		' initialise error variable
		vErrors = ""
		
		' get next pdu message
		sPduFileInfo = oSysDataXfer.GetNextPduMessage(request.querystring("Site"), session("DataBaseDesc"), sPduDirectory, request.querystring("PreviousMessageID"), vErrors)
		
		Response.Write sPduFileInfo
		Response.Write "<msg>"
		Response.Write vErrors
		
		' kill object
		set oSysDataXfer = nothing
	end if

%>
<!--#include file=CloseDataConnection.txt-->