<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<!--r-->
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<!--r-->
<%
dim oBuffer
dim oUser
dim sUser
dim fltSi
dim fltSt
dim fltSj
dim sHTML
dim sMsg

	on error resume next

	sUser = session("ssUser")
	fltSi = Request.QueryString("fltSi")
	fltSt = Request.QueryString("fltSt")
	fltSj = Request.QueryString("fltSj")
%>
<html>
	<head>
		<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>
		<script language="javascript">
			function NavigateToPage(sURL)
			{
				window.navigate(sURL);
			}
		</script>
		<!-- #include file="HandleBrowserEvents.asp" -->
	</head>
<%
	Response.Flush
	' retrieve hex user from user object
	set oUser = server.CreateObject("MACROUSERBS30.MACROUser")
	call oUser.SetState(session("ssUser"))
	sUser = oUser.GetStateHex(false)
	set oUser = nothing
	' check numeric parameters are ok
	if fltSt = "" then
		fltSt = -1
	end if
	if fltSj = "" then
		fltSj = -1
	end if
	' create buffer browser object instance
	set oBuffer = server.CreateObject("InferMed.MACROBuffer.MACROBufferBrowser")
	sHTML = oBuffer.BufferSummaryPage(sUser,true,fltSt,fltSi,fltSj)
	Response.Write sHTML
	set oBuffer = nothing
	
	' if an error during processing display user message
	if sHTML = "" then
		sMsg = "MACRO Buffer Browser (Buffer Summary Page) has encountered a problem whilst processing your request. <br>" 
		sMsg = sMsg & "Please try the operation again. <br>" 
		sMsg = sMsg & "If the problem persists, please inform your study administrator that an error "
		sMsg = sMsg & "occurred in the MACRO Buffer Summary Page so he may check the server log file (MACROBufferBrowser.log)."	
		' redirect to error page
		Response.Write("<script>window.navigate('BufferError.asp?msg=" & Server.URLEncode(sMsg) & "')</script>")
		Response.End 
	end if

	' handle any error
	if err.number <> 0 then 
		call fnError(err.number,err.description,"BufferSummary.asp oBuffer.BufferSummaryPage()",Array(fltSt,fltSi,fltSj),false)
	end if
%>
</html>