<%@ Language=VBScript %>
<%
sUsername = Request.querystring("username")
if sUSername = "" then
	'clear user
	session("user") = ""
end if
%>
<html xmlns:v="urn:schemas-microsoft-com:vml">
<HEAD>
	<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>
	<title>Login</title>
	<script language=javascript>
		function fnValidate(){
			return true;
			var password = document.userdetails.password.value;
			var password1 = document.userdetails.newpassword1.value;
			var password2 = document.userdetails.newpassword2.value;
			return true;
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
					return false;
			
			}
		}
	</script>	
</head>
<body scroll=no>
<div><img height="100%" width="100%" src="../img/bg.jpg"></div>
<div style="position:absolute; left:30%; top:30%;">
	<form name="login" method=post  onSubmit="return fnValidate()" action=showusers.asp>
		<table cellpadding=3 border="0" align=center valign=middle>
			<tr>
				<td class='clsLabelText' style='color:#ffffff;font-weight:bold;'>Username</td>
				<td class='clsTextbox'><INPUT id=username name=username
				<%If sUsername <> "" then Response.Write " value=" & sUsername%>
				></td></tr>
			<tr>
				<td class='clsLabelText' style='color:#ffffff;font-weight:bold;'>Password</td>
				<td class='clsTextbox'><INPUT type=password id=password name=password
				></td>
			</tr>
			<%if sUsername <> "" then%>
			<tr>
				<td class='clsLabelText' style='color:#ffffff;font-weight:bold;'>New password</td>
				<td class='clsTextbox'><INPUT type=password id=newpassword1 name=newpassword1
				></td>
			</tr>
			<tr>
				<td class='clsLabelText' style='color:#ffffff;font-weight:bold;'>Confirm password</td>
				<td class='clsTextbox'><INPUT type=password id=newpassword2 name=newpassword2
				></td>
			</tr>
			<%end if%>
			<tr align=right><td></td><td><INPUT style='width:100px;' type=submit class='clsButton' value=OK name=submit></td></tr>
		</table>
	</form>
</div>
</body>
</html>