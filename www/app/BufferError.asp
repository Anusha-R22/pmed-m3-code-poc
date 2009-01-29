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
		<!-- #include file="HandleBrowserEvents.asp" -->
	</head>
	<body class='clsMessageText'>
		<img src='../img/ico_error_perm.gif'>&nbsp;<%=FormatDateTime(now,1) & " " & FormatDateTime(now,3)%><br><br>
		<%if sMsg = "" then%>
			MACRO has encountered a problem whilst processing your request. This has been 
			recorded on the server.<br>
			Please try the operation again and if the problem persists, contact your study 
			administrator.
		<%else
			Response.Write(fnBufferReplaceWithHTMLCodes(sMsg))
		end if%>
	</body>
</html>
<%
'--------------------------------------------------------------------------------------------------
Function fnBufferReplaceWithHTMLCodes(sValue)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------

	'first replace '&' to encode possible html codes
    sValue = Replace(sValue, "&", "&#38;")
        
    'replace html tag chars
	sValue = Replace(sValue, "<", "&#60;")
	sValue = Replace(sValue, ">", "&#62;")
	
	' change back any <br>
	sValue = Replace(sValue, "&#60;br&#62;", "<br>")
	
    fnBufferReplaceWithHTMLCodes = sValue
End Function
%>