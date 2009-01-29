<%
'==================================================================================================
'	copyright:		InferMed Ltd 2001. all rights reserved
'	file:			DialogLoggedIn.asp
'	date:			
'	amendments:		
'==================================================================================================
%>
<%
if (session("ssUser") = "") then
%>	
	<link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css">
	<div class='clsMessageText'>Your session has expired. Please log back into MACRO</div>
	
<%
	Response.End 
End If
%>
