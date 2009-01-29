<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'	DPH 08/04/2002 - Changed write to a secure folder
'	DPH 01/05/2002 - Changed error handling to <> 0
'	ic 24/05/2004 added variable checking
'-----------------------------------------------------------------------------------------------'

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

	set objFSO = CreateObject("Scripting.FileSystemObject")

	' DPH 08/04/2002 - Changed write to a secure folder
	'set objFile = objFSO.OpenTextFile(gsAppPath & request.form("FileName") ,8 )  '8-> append
	set objFile = objFSO.OpenTextFile(gsSecureFolder & request.form("FileName") ,8 )  '8-> append

	objFile.Write(request.form("hex"))

	objFile.close

	if err.number <> 0 then
		response.write "ERROR2:" & err.number 

	else
		response.write "SUCCESS"
	end if
end if

%>

<!--#include file=CloseDataConnection.txt-->
