<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="Global.asp" -->
<%
	dim sMsg
	sMsg = Request.QueryString("msg")
%>
<html>
	<head>
		<title>&nbsp;<%=Application("asName")%></title>
		<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>
	</head>
	<body class='clsMessageText'>
		<img src='../img/ico_error_perm.gif'>&nbsp;<%=FormatDateTime(now,1) & " " & FormatDateTime(now,3)%><br><br>
		<%if sMsg = "" then%>
			MACRO has encountered a problem whilst processing your request. This has been 
			recorded on the server.<br>
			Please try the operation again and if the problem persists, contact your study 
			administrator.
		<%else
			Response.Write(fnReplaceWithHTMLCodes(sMsg))
		end if%>
	</body>
</html>
