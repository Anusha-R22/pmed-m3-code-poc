<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = false%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       ChangePassword.asp
'   Author:     i curtis
'   Purpose:    prompts the user to change password
'==================================================================================================
'	Revisions:
'	ic 25/06/2004 added parameter checking, error handling
'==================================================================================================
'
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<%
dim sDatabase
dim sRole
dim sAppState
dim sForm
dim bPasswordOK
dim nLoop
dim oIo
dim vErrors
dim sUser
dim vPassword
dim sUserName

	on error resume next

	sForm = Request.Form 
	sUserName = session("ssUserName")
	sDatabase = Request.QueryString("db")
	sRole = Request.QueryString("rl")
	sAppState = Request.QueryString("app")
	
	'validate parameters that we will write into the page
	if not (fnValidateDatabase(sDatabase)) then sDatabase = ""
	if not (fnValidateRole(sRole)) then sRole = ""
	if not (fnValidateAppState(sAppState)) then sAppState = ""
	

	if (sForm <> "") then
		Set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
		sUser = oIo.ChangePasswordForce(sUserName,sForm,vPassword,bPasswordOK,vErrors)
		set oIo = nothing
		
		if err.number <> 0 then 
			call fnError(err.number,err.description,"ChangePassword.asp oIo.ChangePasswordForce()",Array(sUserName,sForm,vPassword),true)
		end if
		
		if bPasswordOK then
			Session("ssUserPassword") = vPassword
			Session("ssUser") = sUser
			Response.Redirect("SelectDatabase.asp?db=" & server.URLEncode(sDatabase) & "&rl=" & Server.URLEncode(sRole) & "&app=" & Server.URLEncode(sAppState))
		end if
	end if
	
%>
	<html>
	<head>
	<title><%=Application("asName")%></title>
	<link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css">
	<!-- #include file="HandleBrowserEvents.asp" -->
	<script language='javascript'>
	function fnPageLoaded()
	{
<%
	if not isempty(vErrors) then
		If Not IsEmpty(vErrors) Then
			response.write("alert('MACRO encountered problems while updating. Some updates could not be completed." _
                    & "\nIncomplete updates are listed below\n\n")

            For nLoop = LBound(vErrors,2) To UBound(vErrors,2)
                response.write(vErrors(0, nLoop) & " - " & vErrors(1, nLoop) & "\n")
            Next

            response.write("');" & vbCrLf)
        End If
	end if
%>
	fnGetPassword();
}
function fnGetPassword()
{
	var sRtn=window.showModalDialog('ChangePasswordInput.htm?name=<%=sUserName%>&exp=true','','dialogHeight:300px; dialogWidth:500px; center:yes; status:0; dependent,scrollbars');
	if ((sRtn!=undefined)&&(sRtn!=""))
	{
		document.Form1.all["upassword"].value=sRtn;
		document.Form1.submit();
	}
	else
	{
		alert("This account will not be accessible until the password has been changed");
	}
}
</script>
</head>
<body onload='fnPageLoaded();'>
<div><img height="100%" width="100%" src="../img/bg.jpg"></div>
<div style="position:absolute; left:0; top:0;">&nbsp;<a class="clsLabelText" style="color:#ffffff; font-weight:bold;" href="javascript:fnGetPassword();">Retry</a>
<form name="Form1" action="ChangePassword.asp?db=<%=Server.URLEncode(sDatabase)%>&rl=<%=Server.URLEncode(sRole)%>&app=<%=Server.URLEncode(sAppState)%>" method="post">
<input type="hidden" name="upassword">
</form>
</div>
</body>
</html>