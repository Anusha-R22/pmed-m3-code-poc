<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=StringUtilities.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Exchange_receive_data3.asp
'   Author:     Andrew Newbigging
'   Purpose:    Used by TrialOffice for the purpose of receiving part of a patient data file
'		and making an entry in the Server's Message table
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'	Mo Morris 7/6/2000 SR 3560, setting the Messsage table's MessageTimeStamp field
'	Mo Morris 27//9/2001, Changes around field Message.MessageId no longer being an
'		autonumber. MessageId is no calculated as Max + 1 of MessageId's that already exist
'	Nicky Johns 25/10/01 - Added MLM's 2.1 changes (for regional settings bugs)
'	DPH	30/11/2001 - Changed names of recordset used to be consistent throughout + added lMessageID to Insert SQL statement
'	DPH 18/1/2002 - Changed handling of 'MAX' SQL return value to force to long (defaulted to string in oracle)
'	DPH 03/04/2002 - Added Checksum check / changed WriteLine to Write
'	DPH 08/04/2002 - Changed write to a secure folder
'	DPH 01/05/2002 - Changed error handling to <> 0
'	NCJ 23 Dec 02 - Receive extra info with cab file and store in DataImport table
'	NCJ 16 Jan 03 - Corrected typo in SQL for DataImport table
'	ic 24/05/2004 added variable checking
'   TA  18/01/2006 - MessageId now calculated by a sequence to avoid duplicate id problem
'-----------------------------------------------------------------------------------------------'

dim sSQL
dim rsRecordSet
dim lMessageId
dim sTimestamp

on error resume next

if err.number <> 0 then

	response.write "ERROR1:" & err.number
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
	'validate study
	if (not fnValidateStudy(request.form("TrialName"))) then
		Response.Write("ERROR:The study '" & request.form("TrialName") & "' does not exist")
		Response.End 
	end if
	'validate subjectid
	if (not fnNumeric(request.form("SubjectId"))) then
		Response.Write("ERROR:The SubjectId '" & request.form("SubjectId") & "' is not valid")
		Response.End 
	end if
	'validate lastlfmessageid
	if (not fnNumeric(request.form("LastLFMessageId"))) then
		Response.Write("ERROR:The LastLFMessageId '" & request.form("LastLFMessageId") & "' is not valid")
		Response.End 
	end if
	

	set objFSO = CreateObject("Scripting.FileSystemObject")

	' DPH 08/04/2002 - Changed write to a secure folder
	'set objFile = objFSO.OpenTextFile(gsAppPath & request.form("FileName") ,8 ) '8 -> append
	set objFile = objFSO.OpenTextFile(gsSecureFolder & request.form("FileName") ,8 ) '8 -> append

	objFile.Write(request.form("hex"))
	'objFile.WriteLine(request.form("hex"))
	
	objFile.close

	if err.number <> 0 then
	
		response.write "ERROR2:" & err.number 
		Response.End 

	else
	
		' DPH 03/04/2002 - Checksum validation
		' Perform Checksum validation (if checksum value has been sent)
		' No sent value will still allow older unpatched sites to send subject data
		If Request.Form("chksum") <> "" Then
			' Create Checksum object
			set objChkSum = CreateObject("IMEDCheckSum10.CheckSum")
			if err.number <> 0 then
				response.write "ERROR4:" & err.number 
				Response.End 
			end if
			
			If objChkSum.CheckFileCheckSum(gsSecureFolder & request.form("FileName") , Request.Form("chksum")) = False Then
				' CAB has failed checksum so exit
				Response.Write "Error CAB Failed Checksum"
				Response.End 
			End If

			if err.number <> 0 then
				response.write "ERROR6:" & err.number 
				Response.End 
			end if
			
		End If
	
        '   TA  18/01/2006 - MessageId now calculated by a sequence to avoid duplicate id problem
        dim xfer
        set xfer = Server.CreateObject("MACROSysDataXfer30.SysDataXfer")
        lMessageId = Clng(xfer.GetNextMessageId((MACROCnn)))

		' NCJ 25/10/01 Use ConvertLocalNumToStandard
		' NCJ 16 Jan 03 - Get current timestamp for next two DB operations
		sTimestamp = ConvertLocalNumToStandard(CStr(CDbl(Now)))

		msSQL = "INSERT INTO Message (MessageId, TrialSite, MessageType, MessageReceived, " _
			& " MessageBody, MessageParameters, MessageDirection, MessageTimeStamp, MessageTimeStamp_TZ) " _
	    		& "  VALUES (" & lMessageId & ",'" & request.form("Site") & "',10,0,'','" _
			& request.form("FileName") & "', 1," & sTimestamp & "," & session("strTimeZone") & ")"


		MACROCnn.Execute msSQL

		if err.number <> 0 then
			response.write "ERROR6:" & err.number & " while writing to Message table"
			Response.End 
		end if

		' NCJ 23 Dec 02 - Insert a row in the DataImport table
		msSQL = "INSERT INTO DataImport "
		msSQL = msSQL & " (ClinicalTrialName,TrialSite,PersonId, " _
				& " DataFileName,LastLFMessageId,ReceivedTimeStamp, ReceivedTimeStamp_TZ) "
	    	msSQL = msSQL & " VALUES ('" & request.form("TrialName") & "','" & request.form("Site") & "', " _
				& request.form("SubjectId") & ", '" & request.form("FileName") & "', " _
				& request.form("LastLFMessageId") & ", " _
				& sTimestamp & "," & session("strTimeZone") & ")"

		MACROCnn.Execute msSQL

		if err.number <> 0 then
			response.write "ERROR3:" & err.number & " while writing to DataImport table"

		else

			response.write "SUCCESS"

		end if
	end if
end if

%>

<!--#include file=CloseDataConnection.txt-->
