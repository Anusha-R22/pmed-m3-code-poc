<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		DialogChangePassword.asp
'	Author: 	I Curtis
'	Purpose: 	gets password authentication input
'				querystring parameters:
'					name: response name
'					role: role of authorisor
'	Version:	1.0
'==================================================================================================
'	Revisions:
'==================================================================================================
%>
<!-- #include file="checkSSL.asp" -->
<!-- #include file="DialogLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<%
dim sExp
dim sName
dim sForm
dim vErrors
dim bOK
dim sUser
dim nLoop
dim oIo

	on error resume next

	sExp = Request.QueryString("exp")
	sName = Request.QueryString("name")
	sForm = Request.Form()
	sUser = session("ssUser")
	
	if (sForm <> "") then
		set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
		bOK = oIo.ChangePasswordRequest(sUser,sForm,vErrors)
		set oIo = nothing
		
		if err.number <> 0 then 
			call fnError(err.number,err.description,"DialogChangePassword.asp oIo.ChangePasswordRequest()",Array(),true)
		end if
		
		if (bOK) then
			session("ssUser") = sUser
			response.Write("<script language='javascript'>alert('Password updated successfully');window.close();</script>")
		end if
	end if

%>
<html>
<head>
<title>Change Password</title>
<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>
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
	document.FORM1.txtP1.focus();
	window.name="WinChangePassword"
}
function fnSubmit()
{
	var sOld = document.FORM1.txtP1.value;
	var sNew = document.FORM1.txtP2.value;
	var sConfirm = document.FORM1.txtP3.value;
	
	if (sOld=="")
	{
		alert("you have not entered your existing password");
		document.FORM1.txtP1.focus();
		return false;
	}
	if (sNew=="")
	{
		alert("you have not entered a new password");
		document.FORM1.txtP2.focus();
		return false;
	}
	if (sNew!=sConfirm)
	{
		alert("password and confirmation do not match");
		document.FORM1.txtP3.value="";
		document.FORM1.txtP3.focus();
		return false;
	}
	if (!fnAlphaNumeric(document.FORM1.txtP2.value))
	{
		alert("Password may only contain alphanumeric characters")
		return false;
	}
	document.FORM1.submit();
}
function fnClose()
{
	this.close();
}
function fnAlphaNumeric(vValue)
{
	//allow alphanumeric or space
	var	sCheck = /[^a-zA-Z0-9 ]/;
	return (sCheck.exec(vValue)==null);
}
</script>
</head><body onload='fnPageLoaded();'>
<form name='FORM1' action='DialogChangePassword.asp' method='post' target='WinChangePassword'>
<table align='center' width='95%' border='0'>
<tr height='10'><td></td></tr>
<tr height='15' class='clsLabelText'><td colspan='2'><div id='divName'><b>User: <%=sName%></b></div></td><td><a style='cursor:hand;' onclick='javascript:fnClose();'><u>Cancel</u></a></td><td><a style='cursor:hand;' onclick='javascript:fnSubmit();'><u>OK</u></a></td></tr><tr height='15'><td></td></tr>
<tr height='5'><td></td></tr><tr height='30'><td colspan='4' class='clsLabelText'>
<div id='divLabel'><%if (sExp <> "") then response.Write("<b>YOUR ACCOUNT PASSWORD HAS EXPIRED</b><br><br>")%>Please update login password for user <%=sName%></div></td></tr>
<tr height='5'><td></td></tr>
<tr><td width='150' class='clsLabelText'>Old Password</td>
<td><input style='width:180px;' class='clsTextbox' name='txtP1' type='password'></td>
</tr>
<tr><td class='clsLabelText'>New Password</td>
<td><input style='width:180px;' class='clsTextbox' name='txtP2' type='password'>
</td></tr>
<tr><td class='clsLabelText'>Confirm Password</td>
<td><input style='width:180px;' class='clsTextbox' name='txtP3' type='password'>
</td></tr>
</table>
</form>
</body></html>