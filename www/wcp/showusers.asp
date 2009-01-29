<%@ Language=VBScript %>
<html xmlns:v="urn:schemas-microsoft-com:vml">
	<HEAD><title>User list</title>
	<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>
	</head>
<body bgcolor=#E0E0FF scroll=auto onload="myAlert();">
<%
dim vSerialisedUser: vSerialisedUser=""
dim colDetails: set colDetails = nothing
dim bPasswordChanged: bPasswordChanged=false


vSerialisedUser=session("user") 

if Request.Form("newpassword1") <> "" then
	'change password request
	bPasswordChanged = ChangeOwnPassword(vSerialisedUser,Request.Form("password"),Request.Form("newpassword1"),Request.Form("newpassword2"))
else
	call NoAlert()
	if vSerialisedUser = "" then
		set oAPI = Login(Request.Form("username"),Request.Form("password"))
		call GetUser(false)
	end if
end if

if vSerialisedUser <> "" then
	set colDetails = GetUsersDetails(vSerialisedUser)
	set oDetail = CreateObject("MACROAPI30.UserDetail")
end if


%>
    <div style="position:absolute; left:10%; top:10%;">
<%
'ic 07/03/2007 issue 2889
DisplayAppHeader()
if not (colDetails is nothing) then
%>
		<table border=1 cellpadding=3>
		<tr align=center>
			<td class='clsTableHeaderText'>Username</td><td class='clsTableHeaderText'>Full Username</td>
			<td class='clsTableHeaderText'>Enabled</td>
			<td class='clsTableHeaderText'>Locked</td>
			<td class='clsTableHeaderText'>Sys Admin</td>
			<td>-</td></tr>
		<%for each oDetail in colDetails%>
			<tr align=center>
				<td class='clsTableText'><%=oDetail.UserName%></td>
				<td class='clsTableText'><%=oDetail.UserNameFull%></td>
				<td class='clsTableText'><%=TrueToYes(oDetail.Enabled)%></td>
				<td class='clsTableText'><%=TrueToYes(oDetail.FailedAttempts>=3)%></td>
				<td class='clsTableText'><%=TrueToYes(oDetail.SysAdmin)%></td>
				<td ><a style='cursor:hand;' href=edituser.asp?username=<%=oDetail.UserName%>&doupdate=false>Edit</a></td>
			</tr>
		<%next%>
		</table>
		<p>
<%
End IF
%>
        <a class='clsLabelText' style='cursor:hand' href=wcplogin.asp>Logout</a>
	</div>
</body>
</html>
<!--#include file="wcplibrary.asp" -->