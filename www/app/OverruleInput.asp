<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		overruleInput.asp
'	Author: 	I Curtis
'	Purpose: 	gets overrule warning text
'				querystring parameters:
'					name: response name
'	Version:	1.0
'==================================================================================================
'	Revisions:
'==================================================================================================
%>
<!-- #include file="checkSSL.asp" -->
<%

dim sName

	sName = Request.QueryString("name")

%>
	<html>
	<head>
	<title>Overrule Warning</title>
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
	<body onload="txtOverrule.focus();">

	<script language="javascript">
	//ic 12/12/2001
	//function builds a delimited overrule string from input, validates and returns if ok
	function fnBuildOverrule(sStr)
	{
		var sIllegal='`|~"'
		var sRtn="";
		
		if (sStr!="")
		{
			if (checkIllegalChars(sStr,sIllegal))
			{
				if (checkLength(sStr,255))
				{
					sRtn=sStr;
				}
				else
				{
					alert("The total length of warning overrules cannot exceed 255 characters")
					txtOverrule.select();
					return;
				}
			}
			else
			{
				alert("Warning overrules cannot contain double or backward quotes or the | character");
				txtOverrule.select();
				return;
			}
		}
		retval(sRtn);
	}
	</script>

	<table class="INPUT_LABEL" align="center" width="95%" border="0">
	<tr height="30"><td></td></tr>
	<tr height="5">
	<td>Please enter the reason overruling warning on <%=sName%></td>
	<td align="right"><input name="btnReturn" type="button" value="&nbsp;&nbsp;&nbsp;&nbsp;OK&nbsp;&nbsp;&nbsp;&nbsp;" onclick="javascript:fnBuildOverrule(txtOverrule.value)"></td>
	</tr>
	<tr height="5">
	<td></td>
	<td align="right"><input type="button" value="Cancel" onclick="javascript:retval('');"></td>
	</tr>
	<tr height="20"><td></td></tr>
	<tr>
	<td colspan="2"><input name="txtOverrule" type="text" size="60"></td>
	</tr>
	</table>

	</body>
	</html>