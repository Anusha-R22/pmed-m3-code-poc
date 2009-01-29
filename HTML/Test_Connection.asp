<%

'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       Test_Connection.asp
'   Author:     Richard Meinesz, 2002
'   Purpose:    Used by Config form to test the connection to the server
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'-----------------------------------------------------------------------------------------------'

dim oConn
	
	on error resume next
	
	Response.Write "Success<p>"
	
	
	set oConn = server.CreateObject("ADODB.Connection")
	
	if err.number <> 0 then
		Response.Write "Error creating server database connection object.  Error Description: " & err.Description & "<p>"
		Response.End
	else
		Response.Write "<p>"
	end if	
	
	
	oConn.Open (session("strConn"))
	
	if err.number <> 0 then
		Response.Write "Error opening connection to server database." & vbcrlf &  "Error Description: " & err.Description
		Response.End
	else
		Response.Write "<p>"
	end if
	
	oConn.Close
	set oConn = nothing

 %>

