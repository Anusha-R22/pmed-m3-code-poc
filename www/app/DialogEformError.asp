<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' Copyright:InferMed Ltd. 1998. All Rights Reserved
'	File:			DialogEformError.asp
'	Author: 	I Curtis
'	Purpose: 	Reports eForm errors to the server
'==================================================================================================
'	Revisions:
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="DialogLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<%
dim sMsg
dim sUrl
dim sLine
dim sSource
dim sAgent
dim bReported
dim n
dim oIo

	on error resume next

	sMsg = Request.Form("Msg")
	sUrl = Request.Form("Url")
	sLine = Request.Form("Line")
	sAgent = Request.Form("Agent")
	
	
	bReported = false
	if (sMsg <> "") then
		'rebuild the html body source
		For n = 1 To Request.Form("SourceHTML").Count
			sSource = sSource & Request.Form("SourceHTML")(n)
		Next
	
		'page has submitted to itself, report the error
		set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
		call oIo.LogError(sUrl+":"+sLine+":"+sAgent,"",sMsg,Array(sSource))
		if err.number <> 0 then 
			set oIo = nothing
			call fnError(err.number,err.description,"DialogEformError.asp oIo.LogError()",Array(sMsg,sUrl,sLine,sSource,sAgent),true)
		end if
		set oIo = nothing
		bReported = true
	end if
	
	if (not bReported) then
%>
		<html>
		<head>
		<title>InferMed MACRO - Error</title>
		<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>
		<script language="javascript">
		function fnPageLoaded()
		{
			window.name="winError";
			var aArgs=window.dialogArguments;
			document.FormError.Msg.value=aArgs[0]+" ";
			document.FormError.Url.value=aArgs[1];
			document.FormError.Line.value=aArgs[2];
			fnSplitLargeString(aArgs[3]);
			document.FormError.Agent.value=aArgs[4];
			FormError.submit();
		}
		function fnSplitLargeString(s)
		{		
			var MAXLENGTH = 102300;
			var temps = s;
			
			temps = s.substr(0,MAXLENGTH);
			s = s.substr(MAXLENGTH);
			
			while (temps.length > 0)
			{
				var objTEXTAREA = document.createElement("TEXTAREA");
				
				objTEXTAREA.name = "SourceHTML";
				objTEXTAREA.value = temps;
				document.FormError.appendChild(objTEXTAREA);

				temps = s.substr(0,MAXLENGTH);
				s = s.substr(MAXLENGTH);
			}
		}
		</script>
		</head>
		
		<div class="clsProcessingMessageBox" id="divMsgBox">
		<table height="100%" align="center" width="90%" ID="Table1">
		<tr><td id="tdMsg" valign="middle" class="clsMessageText">please wait<br><br><img src="../img/clock.gif">
		&nbsp;&nbsp;Contacting server...</td></tr></table></div>
		
		<body onload='fnPageLoaded();'>
		<div style='visibility:hidden;'>
		<form name='FormError' action='DialogEformError.asp' method='post' target='winError'>
		<input type='hidden' name='Msg' value=''>
		<input type='hidden' name='Url' value=''>
		<input type='hidden' name='Line' value=''>
		<input type='hidden' name='Agent' value=''>
		</form>
		<div>
		</body>
		</html>
<%
	else
%>
		<head>
		<title>InferMed MACRO - Error</title>
		<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>
		</head>
		
		<body>
		<table width='95%' align='center' valign='middle' class="clsMessageText">
		<tr height='30'><td></td></tr>
		<tr><td><%=FormatDateTime(now,1) & " " & FormatDateTime(now,3)%><br><br>
		<img src='../img/ico_error_perm.gif'>&nbsp;<b><u>An unexpected error has occurred on this eForm</u></b> 
		<br><br><br>
				
		<div id='div1' class='clsLabelText'>
		This has been logged on the server. Close this dialog and exit the eForm by clicking an option from the left-hand menu, 
		then try to reopen it. If the problem persists contact your study administrator.		
		
		</td></tr></table>
		</body></html>
<%
	end if
%>