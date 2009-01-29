<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  	InferMed Ltd. 2000 All Rights Reserved
'   File:       	Registration.asp
'   Author:     	Mo Morris, 23 November 2000
'   Purpose:    	Used by Macro to call the Registration server for the purpose of getting a
'			SubjectIdentifier and a ResultCode that reflects the outcome of the call
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'	NCJ 29 Nov 00 - oRegClass Initialise now takes extra parameters	
'-----------------------------------------------------------------------------------------------'

On Error Resume Next

Dim sSubjectIdentifier
Dim nResultCode

'Declare the Registration class
Set oRegClass = Server.CreateObject("MACRORR.clsRSSubjectNumbering")

'Tell the Registration Class the Database description that is declared in Global.asa
oRegClass.SetDatabase session("DataBaseDesc")

'Initialise the Registration Class
oRegClass.Initialise Request.Querystring("TrialName"), _
			Request.Querystring("Site"), _
			Request.Querystring("PersonId"), _
			Request.Querystring("Prefix"), _
			Request.Querystring("Suffix"), _
			Request.QueryString("UsePrefix"), _
			Request.QueryString("UseSuffix"), _
			Request.QueryString("StartNumber"), _
			Request.QueryString("NumberWidth")

'Pass the Uniqueness checks data to the Registration Class
oRegClass.AddUniquenessChecks Request.Querystring("UCheckString")

'Call the Registration Class twice
'The first call returns the SubjectIdentifier
sSubjectIdentifier = oRegClass.SubjectIdentifier

'The second call returns the ResultCode (Status) of the call
nResultCode = oRegClass.ResultCode

'Close down the Registration Class
Set oRegClass = Nothing

If Err.Number = 0 Then
	'Both are written back at the same time
	Response.Write sSubjectIdentifier
	Response.Write "<br>"
	Response.Write nResultCode
Else
	'Note that 4 is the Registration Class's code for an error has occurred
	Response.Write "<br>4"
End If


%>
