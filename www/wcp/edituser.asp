<%@ Language=VBScript %>
<html xmlns:v="urn:schemas-microsoft-com:vml">
	<HEAD>
	<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>
	<script language=javascript>
		function fnValidate(){
			var password1 = document.userdetails.password1.value;
			var password2 = document.userdetails.password2.value;
			var bNowEnabled = document.userdetails.chkEnabled.checked;
			if (password1!=""||password2!="")
			{
				if (password1==password2)
				{
					//alert("changing passwords");
					return true;
				}
				else
				{
					alert("passwords do not match");
					return false;
				}
			}
			else 
			{		
				if (bEnabled^bNowEnabled)
				{
					return true;
				}
				else
				{
					alert("no changes to save");
					return false;
				}
			}
		}
	</script>
	<title>Update user</title>
	</head>
<body bgcolor=#E0E0FF scroll=no onload="myAlert();">
<%
set oDetail = CreateObject("MACROAPI30.UserDetail")
set oDetail = ShowUpdateUser(Request.QueryString("username"),Request.QueryString("doupdate"))
%>
    <div style="position:absolute; left:10%; top:10%;">
<%
'ic 07/03/2007 issue 2889
DisplayAppHeader()
if not (oDetail is nothing) then
%>
	<%
	if oDetail.Enabled then
	%>
		<script language='javascript'>var bEnabled = true;</script>
	<%
	else
	%>
		<script language='javascript'>var bEnabled = false;</script>
	<%
	end if
	%>

	<form name="userdetails" onSubmit="return fnValidate()" method=post action=edituser.asp?username=<%=oDetail.UserName%>&doupdate=true>
		<table cellpadding=3>
			<tr><td class='clsLabelText'>User</td><td class='clsLabelText'><%=oDetail.UserName%></td></tr>
			<tr><td class='clsLabelText'>Full Name</td><td class='clsLabelText'><%=oDetail.UserNameFull%></td></tr>
			<tr><td class='clsLabelText'>Password</td><td class='clsTextbox'><INPUT type=password id=password1 name=password1></td></tr>
			<tr><td class='clsLabelText'>Confirm Password</td><td class='clsTextbox'><INPUT type=password id=password2 name=password2></td></tr>
			<tr><td/><td><INPUT type=checkbox  class='clsLabelText' <%=TrueToChecked(oDetail.Enabled)%> id=chkEnabled name=chkEnabled>Enabled</td></tr>
			<tr><td/><td><INPUT type=checkbox class='clsLabelText' <%=TrueToChecked(oDetail.FailedAttempts>=3)%>  style='color:#555555;' disabled=true id=chkLocked name=chkLocked>Locked Out</td></tr>
			<tr><td/><td><INPUT type=checkbox class='clsLabelText' <%=TrueToChecked(oDetail.SysAdmin)%> disabled=true  style='color:#555555;' id=chkSysAdmin name=chkSysAdmin>Systems Administrator</td></td><td></td></tr>
			<tr align=right><td></td><td><INPUT type=submit class='clsButton' value=Apply name=submit></td></tr>
		</table>
	</form>
	<p>
	<a class='clsLabelText' style='cursor:hand;' href=showusers.asp>Make no changes and take me back to the list</a>
	<p>
<%
end if
%>
        <a class='clsLabelText' style='cursor:hand;' href=wcplogin.asp>Logout</a>
	</p>
	</div>
</body>
</html>
<!--#include file="wcplibrary.asp" -->
