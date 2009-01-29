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
dim sForm
dim sHTML

	on error resume next

	sUser = session("ssUser")
	' form info
	sForm = Request.Form
%>
<html>
	<head>
		<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>
		<script language="javascript" src="../script/BufferDataBrowser.js"></script>
		<script language="javascript">
			function NavigateToPage(sURL)
			{
				window.navigate(sURL);
			}
			// screen title
			window.parent.fnSetTitle("Buffer Data Browser Save Results");
		</script>
		<script language="javascript">
		    function fnPageLoaded()
		    {
		        document.all.divMsgBox.style.visibility='hidden';;
		    }
		</script>
		<!-- #include file="HandleBrowserEvents.asp" -->
	</head>

	<div class="clsProcessingMessageBox" id="divMsgBox">
	<table height="100%" align="center" width="90%">
	<tr><td id="tdMsg" valign="middle" class="clsMessageText">please wait<br><br><img src="../img/clock.gif">
	&nbsp;&nbsp;Saving data...</td></tr></table></div>

	<div class="clsPopMenu" id="divPopMenu" onclick="menu=this;this.tid=setTimeout('menu.style.visibility=\'hidden\'',20);" 
	onmouseout="menu=this;this.tid=setTimeout('menu.style.visibility=\'hidden\'',20);" onmouseover="clearTimeout(this.tid);">
	</div>
<%
	' * BufferDataSaveResults.asp - save data and show results
	Response.Flush
	' retrieve hex user from user object
	set oUser = server.CreateObject("MACROUSERBS30.MACROUser")
	call oUser.SetState(session("ssUser"))
	sUser = oUser.GetStateHex(false)
	set oUser = nothing
	' create buffer browser object instance
	set oBuffer = server.CreateObject("InferMed.MACROBuffer.MACROBufferBrowser")
	' save data / show next page
	'Response.Write sUser
	sHTML = oBuffer.GetBufferSaveResultsPage(sUser,true,sForm)
	Response.Write sHTML
	set oBuffer = nothing
	
	' if an error during processing display user message
	if sHTML = "" then
		sMsg = "MACRO Buffer Browser (Buffer Save Results) has encountered a problem whilst processing your request.<br>"
		sMsg = sMsg & "Please try the operation again.<br>"
		sMsg = sMsg & "If the problem persists, please inform your study administrator that an error "
		sMsg = sMsg & "occurred in the MACRO Buffer Save Results Page so he may check the server log file (MACROBufferBrowser.log)."		
		' redirect to error page
		Response.Write("<script>window.navigate('BufferError.asp?msg=" & Server.URLEncode(sMsg) & "')</script>")
		Response.End 
	end if

	' handle any error
	if err.number <> 0 then 
		call fnError(err.number,err.description,"BufferDataSaveResults.asp oBuffer.GetBufferSaveResultsPage()",Array(""),false)
	end if
%>
</html>