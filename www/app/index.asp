<%
'==================================================================================================
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       index.asp
'   Author:     i curtis
'   Purpose:    redirection page
'	Version:	1.0
'==================================================================================================
'	Revisions:
'	ic	22/11/2002	changed www directory structure
'==================================================================================================
'
%>
<html>
<head>
<title>InferMed MACRO</title>
<script language="javascript">
function fnOpenMacroWin()
{
	var nH=screen.availHeight-30;
	var nW=screen.availWidth-15;
	window.open("Login.asp","","toolbar=no,resizable=yes,status=no,scrollbars=auto,height="+nH+",width="+nW+",top=0,left=0");
}
</script>
<link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css">
</head>
<body onload="fnOpenMacroWin();">
<div class="MESSAGE">You can close this browser window.</div>
</body>
</html>

