<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=StringUtilities.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2003 All Rights Reserved
'   File:       Get_Report_Files.asp
'   Author:     David Hook, 2003
'   Purpose:    Used by TrialOffice for the purpose of retrieving report files from the server
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'	ic 24/05/2004 added variable checking
'-----------------------------------------------------------------------------------------------'
dim oSysDataXfer 
dim sFilename
dim sResponse

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

		set oSysDataXfer = CreateObject("MACROSysDataXfer30.SysDataXfer")

		if Request.Form("confirm") = "" then
			'validate lasttrans
			if (not fnNumeric(Request.Form("lasttrans"))) then
				Response.Write("ERROR:Last transaction '" & Request.Form("lasttrans") & "' is not valid")
				Response.End 
			end if
		
		
			sFilename = oSysDataXfer.GetReportFiles(Request.Form("site"),session("DataBaseDesc"),Request.Form("username"),Request.Form("lasttrans"),vMessage)
		
			Response.Write sFilename
			Response.Write "<br>"

		else

			sResponse = oSysDataXfer.ConfirmReportFiles(Request.Form("site"),session("DataBaseDesc"),Request.Form("username"),vMessage)
		
			Response.Write sResponse
			Response.Write "<br>"

		end if

		Response.Write vMessage
		
	end if
%>

<!--#include file=CloseDataConnection.txt-->