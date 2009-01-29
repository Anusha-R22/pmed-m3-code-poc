<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=StringUtilities.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Exchange_receive_Laboratory.asp
'   Author:     Mo Morris
'   Purpose:    Used by TrialOffice for the purpose of receiving a Laboratory data definition file
'		and making an entry in the Server's Message table to record its sending
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'	Mo Morris 27//9/2001, Changes around field Message.MessageId no longer being an
'		autonumber. MessageId is no calculated as Max + 1 of MessageId's that already exist
'	Nicky Johns 25/10/01 - Added MLM's 2.1 changes (for regional settings bugs)
'	DPH	30/11/2001 - Changed names of recordset used to be consistent throughout + added lMessageID to Insert SQL statement
'	DPH 18/1/2002 - Changed handling of 'MAX' SQL return value to force to long (defaulted to string in oracle)
'	DPH 03/04/2002 Added Checksum check / changed WriteLine to Write
'	DPH 08/04/2002 - Changed write to a secure folder
'	DPH 01/05/2002 - Changed error handling to <> 0
'	ic 24/05/2004 added variable checking
'   TA  18/01/2006 - MessageId now calculated by a sequence to avoid duplicate id problem
'-----------------------------------------------------------------------------------------------'

dim sSQL
dim rsRecordSet
dim lMessageId

on error resume next

if err.number <> 0 then

	response.write "ERROR AT LAUNCH:" & err.number
	Response.End 
	
else

	'validate filename
	if (not fnValidateFilename(request.form("FileName"))) then
		Response.Write("ERROR:The filename '" & request.form("FileName") & "' is not valid")
		Response.End 
	end if
	'validate site
	if (not fnValidateSite(request.form("Site"))) then
		Response.Write("ERROR:The site '" & request.form("Site") & "' does not exist")
		Response.End 
	end if

	set objFSO = CreateObject("Scripting.FileSystemObject")

'	DPH 08/04/2002 - Changed write to a secure folder
	'set objFile = objFSO.OpenTextFile(gsAppPath & request.form("FileName") ,2,1)
	set objFile = objFSO.OpenTextFile(gsSecureFolder & request.form("FileName") ,2,1)

	objFile.Write(request.form("Data"))
	'objFile.WriteLine(request.form("Data"))

	objFile.close

	if err.number <> 0 then
	
		response.write "ERROR CREATING FILE:" & err.number 
		Response.End 

	else
	
		' DPH 03/04/2002 - Checksum validation
		' Perform Checksum validation (if checksum value has been sent)
		' No sent value will still allow older unpatched sites to send subject data
		If Request.Form("chksum") <> "" Then
			' Create Checksum object
			set objChkSum = CreateObject("IMEDCheckSum10.CheckSum")
			if err.number <> 0 then
				response.write "ERROR CREATING OBJECT:" & err.number 
				Response.End 
			end if
			
			If objChkSum.CheckFileCheckSum(gsSecureFolder & request.form("FileName") , Request.Form("chksum")) = False Then
				' CAB has failed checksum so exit
				Response.Write "Error CAB Failed Checksum"
				Response.End 
			End If

			if err.number <> 0 then
				response.write "ERROR USING OBJECT:" & err.number 
				Response.End 
			end if
			
		End If

        '   TA  18/01/2006 - MessageId now calculated by a sequence to avoid duplicate id problem
        dim xfer
        set xfer = Server.CreateObject("MACROSysDataXfer30.SysDataXfer")
        lMessageId = Clng(xfer.GetNextMessageId((MACROCnn)))

		'note that MessageType is set to 31 meaning LabDefinitionSiteToServer
		' NCJ 25/10/01 - Use ConvertLocalNumToStandard
		msSQL = "INSERT INTO Message (MessageId,TrialSite,MessageType,MessageReceived,MessageBody,MessageParameters,MessageDirection,MessageTimeStamp, MessageTimeStamp_TZ) " _
	    	& "  VALUES (" & lMessageId & ",'" & request.form("Site") & "',31,0,'','" & request.form("FileName") & "', 1," & ConvertLocalNumToStandard(CStr(CDbl(Now))) & "," & session("strTimeZone") & ")"

		MACROCnn.Execute msSQL

		if err.number <> 0 then
			response.write "ERROR WRITING TO MESSAGE TABLE:" & err.number

		else

			response.write "SUCCESS"

		end if
		
	end if
	
end if

%>

<!--#include file=CloseDataConnection.txt-->