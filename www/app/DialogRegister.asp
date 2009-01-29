<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		DialogRegister.asp
'	Author: 	I Curtis
'	Purpose: 	Registers a subject, defined in passed 'state' querystring param
'==================================================================================================
'	Revisions:
'	ic 28/06/2004 added error handling
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="DialogLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<%
dim oIo
dim sState
dim bOk
dim sMessage

	on error resume next

	sState = Request.QueryString("state")
%>
	<html>
	<head>
	<link rel="stylesheet" href="../style/MACRO1.css" type="text/css">
	
	<script language='javascript'>
	function fnHideLoader()
	{
		document.all.divMsgBox.style.visibility='hidden';
	}
	</script>
	
	<%if (Application("asDEV") <> "true") then%>
		<script language="javascript">
		document.oncontextmenu=fnContextMenu;
		function fnContextMenu(){return false};
		</script>
	<%end if%>
	</head>
	
	<div class="clsProcessingMessageBox" id="divMsgBox">
	<table height="100%" align="center" width="90%" ID="Table1">
	<tr><td valign="middle" class="clsMessageText">please wait<br><br><img src="../img/clock.gif">
	&nbsp;&nbsp;Processing registration...</td></tr></table></div>
<%	
	Response.Flush
	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	bOk = oIo.RegisterWWWSubject(session("ssUser"),sState,Application("abUSESCI"),sMessage)
	set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"DialogRegister.asp oIo.RegisterWWWSubject()",Array(sState,Application("abUSESCI")),false)
	end if
%>
	<body onload='fnHideLoader()'>
	<div class='clsMessageText'><%=sMessage%></div>
	</body>
	</html>