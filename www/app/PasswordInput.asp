<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		passwordInput.asp
'	Author: 	I Curtis
'	Purpose: 	gets password authentication input
'				querystring parameters:
'					name: response name
'					role: role of authorisor
'	Version:	1.0
'==================================================================================================
'	Revisions:
'	ic	22/11/2002	changed www directory structure
'==================================================================================================
%>
<!-- #include file="checkSSL.asp" -->
<%
dim sName
dim sRole

	sName = Request.QueryString("name")
	sRole = Request.QueryString("role")

%>
	<html>
	<head>
	<title>Password Authentication</title>
	<link rel="stylesheet" href="../style/MACRO1.css" type="text/css">
	<script language="javascript">
		//ic 12/12/2001
		//function returns a string value to calling function from modal dialog
		function retval(sVal)
		{
			this.returnValue=sVal;
			this.close();
		}
	
		//ic 12/12/2001
		//function checks a passed string [sStr] for illegal chars [sIllegal]
		function checkIllegalChars(sStr,sIllegal)
		{
			for (var n=0;n<sIllegal.length;n++)
			{
				for (var p=0;p<sStr.length;p++)
				{
					if (sStr.charAt(p)==sIllegal.charAt(n)) return false;
				}
			}
			return true;
		}
		
		//ic 12/12/2001
		//function checks a passed string [sStr] isnt longer than a passed length [nLength]
		function checkLength(sStr,nLength)
		{
			if (sStr.length>nLength) return false;
			return true;
		}
		
		//ic 29/04/2002
		//function submits page when user hits return key
		function keyPress()
		{
			if (window.event.keyCode==13)
			{
				btnReturn.click();
			}
		}
		
		window.document.onkeypress = keyPress;
	</script>
	</head>
	<body onload="txtId.focus();">



	<script language="javascript">
	//ic 12/12/2001
	//function builds a delimited Authorisation string from input, validates and returns if ok
	function buildAuthorise(sId,sPswd)
	{
		var sIllegal='`|~"'
		var sRtn="";
		
		if ((sId!="")&&(sPswd!=""))
		{
			if (checkIllegalChars(sId,sIllegal))
			{
				if (checkIllegalChars(sPswd,sIllegal))
				{
					sRtn=sId+"|"+sPswd;
				}
				else
				{
					alert("The supplied password contains invalid characters")
					txtPswd.select();
					return;
				}
			}
			else
			{
				alert("The supplied user name contains invalid characters");
				txtId.select();
				return;
			}
		}
		retval(sRtn);
	}
	</script>
	
	<table class="clsLabelText" align="center" width="95%" border="0">
	<tr height="30"><td></td></tr>
	<tr height="5">
	<td>This question needs to be authorised by a user with the following role: <%=sRole%></td>
	<td align="right"><input class='clsButton' style='width:60;' name="btnReturn" type="button" value="&nbsp;&nbsp;&nbsp;&nbsp;OK&nbsp;&nbsp;&nbsp;&nbsp;" onclick="javascript:buildAuthorise(txtId.value,txtPswd.value)"></td>
	</tr>
	<tr height="5">
	<td></td>
	<td align="right"><input class='clsButton' style='width:60;' type="button" value="Cancel" onclick="javascript:retval('');"></td>
	</tr>
	<tr height="20"><td></td></tr>
	<tr>
	<td colspan="2">
	
	<table width="100%" class="clsLabelText">
	<tr>
	<td>User name&nbsp;</td>
	<td><input name="txtId" type="text" size="20"></td>
	</tr>
	<tr>
	<td>Password&nbsp;</td>
	<td><input name="txtPswd" type="password" size="20"></td>
	</tr>
	</table>
	
	</td>
	</tr>

	</table>

	</body>
	</html>