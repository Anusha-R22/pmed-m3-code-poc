<%@ Language=VBScript %>
<%Option Explicit
'==================================================================================================
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       Login.asp
'   Author:     i curtis
'   Purpose:    MACRO web de/dr login page
'				querystring parameters:
'					url: redirects to supplied url after login
'					exp=true: displays session expired message on page
'					action=lock: locks application under current user awaiting unlock
'	Version:	1.0
'==================================================================================================
'	Revisions:
'	ic	29/06/01	amendments for 2.2
'	ic 29/07/2002	changed dll reference for 3.0
'	ic 01/10/2002	changed ie version check to check for ie5 or ie6
'	ic	22/11/2002	changed www directory structure
'	dph 09/06/2003	detects incorrect browser cookie setup
'	dph 17/06/2003	detects old web version
'	ic 23/06/2004	added parameter checking, error handling
'	ic 09/11/2004	disable submit button after submitting
'   ic 12/05/2005   improved error message wording
'   ic 17/11/2006   allow ie7
'   ic 02/11/2007   issue 2958 2 browsers with same session id bug
'==================================================================================================
'
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="Global.asp" -->
<%
Const nLOGIN_OK = 0
Const nACCOUNT_DISABLED = 1
Const nLOGIN_FAILED = 2
Const nCHANGE_PASSWORD = 3
Const nPASSWORD_EXPIRED = 4

Dim bBrowserChecked			
Dim bIEVersionOk				
Dim bCookiesEnabled		
Dim bJavaEnabled
Dim sDecimalPoint
Dim sThousandSeparator
Dim lWebVersion
			
Dim sUserName
Dim sUserFullName		
Dim sUserPassword	
Dim nRtn

Dim sAction
Dim sAppState
Dim sDatabase
Dim sRole

Dim sExpired
Dim oIo
dim sUser	

Dim bAttempt

	on error resume next

	'get querystring parameters
	sExpired = Request.QueryString("exp")
	sAction = Request.QueryString("act")
	bAttempt = cbool(Request.QueryString("att"))
	sAppState = Request.QueryString("app")
	sDatabase = Request.QueryString("db")
	sRole = Request.QueryString("rl")

    'ic 02/11/2007 issue 2958 2 browsers with same session id bug
	if ((session("ssUserName") <> "") and (sAction <> "switch")) then
        Response.Write("Only one window allowed per session")
        
    else
	
	'validate parameters that we will write into the page
	if not (fnValidateDatabase(sDatabase)) then sDatabase = ""
	if not (fnValidateRole(sRole)) then sRole = ""
	if not (fnValidateAppState(sAppState)) then sAppState = ""
	

	If (Session("sbBrowserAccept") <> TRUE) Then

		'get browser info from form
		bBrowserChecked = Request.Form("browserChecked")
		bIEVersionOk = Request.Form("IEVersionOk")
		bCookiesEnabled = Request.Form("cookiesEnabled")
		bJavaEnabled = Request.Form("javaEnabled")
		sDecimalPoint = Request.Form("decimalpoint")
		sThousandSeparator = Request.Form("thousandseparator")
		lWebVersion = clng(Request.Form("webversion"))
		
		'check for minimum spec: ie5+
		If (bBrowserChecked = "yes") Then
			If (bIEVersionOk) Then
				' dph 17/06/2003 - detect have correct web version
				If (Application("anWEBVERS") > 0) AND (Application("anWEBVERS") <= lWebVersion) Then
					'dph 09/06/2003 -  detect if cookies are disabled 
					' If Attempt set then Session Username should exist - else cookie session not set correctly
					If (bCookiesEnabled) And Not (bAttempt And (IsEmpty(Session("ssUserName")) or Session("ssUserName")="")) then
						session("sbBrowserAccept") = TRUE
						' RS 26/09/2002: Set TimezoneOffset for this session, to be passed to the server when saving data
						session("TimezoneOffset") = Request.Form("TimezoneOffset")
						'ic keep local separators too
						session("ssDecimalPoint") = sDecimalPoint
						session("ssThousandSeparator") = sThousandSeparator
					Else
%>					
						<html>
						<head><link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css"><title>&nbsp;<%=Application("asName")%></title>
						</head>
						<body>
						<table width='95%' align='center' valign='middle' class="clsMessageText" ID="Table1">
						<tr height='30'><td></td></tr>
						<tr><td><%=FormatDateTime(now,1) & " " & FormatDateTime(now,3)%><br><br>
						<img src='../img/ico_error_perm.gif'>&nbsp;<b><u>Browser configuration problem</u></b> 
						<br><br><br>
						Please try to log in to MACRO again. If you continue to experience this problem, check that your browser is 
						configured to accept Cookies. Your browser must be configured to accept cookies for MACRO to function correctly.
						</td></tr></table></body></html>
<%
					End If
				Else
					'newer version of pages detected error page
%>					
						<html>
						<head><link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css"><title>&nbsp;<%=Application("asName")%></title>
						<script language='javascript'>
						function fnHideDiv()
						{
							if(document.all.div2.style.visibility=='hidden')
							{
								document.all.div2.style.visibility='visible';
								document.all.div2.innerHTML=sHtm;						
							}
							else
							{
								document.all.div2.style.visibility='hidden';
								document.all.div2.innerHTML="";
							}
						}
						</script>
						</head>
						<body>
						<table width='95%' align='center' valign='middle' class="clsMessageText">
						<tr height='30'><td></td></tr>
						<tr><td><%=FormatDateTime(now,1) & " " & FormatDateTime(now,3)%>&nbsp;&nbsp;<a href='javascript:window.print();'>Print this page</a><br><br>
						<img src='../img/ico_error_perm.gif'>&nbsp;<b><u>A newer version of MACRO has been detected on the server</u></b> 
						<br><br><br>
						
						<div id='div1' class='clsLabelText'>
						These changes should be picked up by Internet Explorer automatically. To try to fix this problem:<br><br>
						<ul>
						<li>Close all open Internet Explorer windows and use a fresh Internet Explorer instance.</li>
						<li>Clear your browser cache. To do this:</li>
						<ul>
						<li>Select '<b>Internet Options...</b>' from the '<b>Tools</b>' menu in Internet Explorer. The '<b>Internet Options</b>' dialog opens.</li>
						<li>Click the '<b>Delete Files...</b>' button in the '<b>Temporary internet files</b>' section. The '<b>Delete Files</b>' dialog opens.</li>
						<li>Check the '<b>Delete all offline content</b>' checkbox.</li>
						<li>Click the '<b>OK</b>' button on the '<b>Delete Files</b>' dialog. The dialog closes.</li>
						<li>Click the '<b>OK</b>' button on the '<b>Internet Options</b>' dialog. The dialog closes.</li>
						</ul>
						<li>Check that your browser settings allow checking for newer versions 
						of stored pages. To do this: 
						<ul>
						<li>Select '<b>Internet Options...</b>' from the '<b>Tools</b>' menu in Internet Explorer. The '<b>Internet Options</b>' dialog opens.</li>
						<li>Click the '<b>Settings...</b>' button in the '<b>Temporary internet files</b>' section. The '<b>Settings</b>' dialog opens.</li>
						<li>The '<b>Automatically</b>' option should be selected in the '<b>Check for newer versions of stored pages</b>' section. Select it if it is not already selected.</li>
						<li>Click the '<b>OK</b>' button on the '<b>Settings</b>' dialog. The dialog closes.</li>
						<li>Click the '<b>OK</b>' button on the '<b>Internet Options</b>' dialog. The dialog closes.</li>
						</ul></li>
						<li>If you are unable to do this and/or are still experiencing this problem, contact your study administrator.</li></ul><br></div><br>
						
						
						<a href="javascript:fnHideDiv()">MACRO Administrators</a><br><br>
						<div id="div2" class='clsLabelText' style="background-color:lightgrey; visibility:hidden;"></div>
						
						<script language='javascript'>
							var sHtm="";
							
						<%if(Application("anWEBVERS") = 0) then%>
							sHtm+="IIS could not read the '<b>webversion</b>' parameter from the MACRO settings file. The parameter may be missing, or ";
							sHtm+="IIS may not have permission to access the file.<br>To fix the problem try the following in the order listed. " 
				
						<%elseif(lWebVersion=-1) then%>
							sHtm+="The client browser was unable to load the '<b>www/script/lWebVersion.js</b>' file. This file is created ";
							sHtm+="by MACRO when HTML is generated through the <b>MACRO System Management</b> console.<br>To fix the problem "; 
							sHtm+="try the following in the order listed. ";
							
						<%else%>
							sHtm+="The MACRO version expected by IIS (<b><%=Application("anWEBVERS")%></b>) differs from the version detected ";
							sHtm+="in the browser (<b><%=lWebVersion%></b>).<br>";
							sHtm+="If the browser checks listed above have failed to fix the problem, try the following in the order listed. ";
							
						<%end if%>
						sHtm+="Make sure you close all open Internet Explorer windows and open a fresh Internet Explorer instance between each one.";
						sHtm+="<ul>";
						sHtm+="<li>Restart the IIS service. To do this:</li>";
						sHtm+="<ul>";
						sHtm+="<li>Select '<b>Run...</b>' from the Windows start menu. The '<b>Run</b>' dialog opens.</li>";
						sHtm+="<li>Type '<b>iisreset</b>' into the textbox.</li>";
						sHtm+="<li>Click the '<b>OK</b>' button. A DOS dialog opens and closes.</li>";
						sHtm+="</ul>";
						sHtm+="<li>Generate HTML through the <b>MACRO System Management</b> console, Restart the IIS service.</li>";
						sHtm+="<li>Check that the '<b>webversion</b>' parameter in the '<b>MACROUserSettings30.txt</b>'<sup>*</sup> file matches the value found in ";
						sHtm+="the '<b>www/script/lWebVersion.js</b>' file.</li>";
						sHtm+="<li>Check that the IIS user has Windows permissions to read the '<b>MACROSettings30.txt</b>' file and the '<b>MACROUserSettings30.txt</b>'<sup>*</sup> file.</li>";
						sHtm+="<li>Check for the latest advice on the MACRO support site.</li>";
						sHtm+="<li>Contact InferMed support.</li>";
						sHtm+="</ul><br>";
						sHtm+="<sup>*</sup>Your settings file name may differ from the default. Check the '<b>MACROSettings30.txt</b>' file to confirm ";
						sHtm+="your settings file name.";
						sHtm+="<br><br>";
						</script>
						</td></tr></table></body></html>
				<%End If
			Else
				Response.Write("MACRO requires IE5.5 or IE6.")
			End If
		Else
		
			'this html page submits browser information back to the asp
			Response.Write("<!--" & Application("asCopyright") & "-->")
%>
			<html>
			<head>
			<title>&nbsp;<%=Application("asName")%></title>
			<link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css">
			<script language="javascript" src="../script/lWebVersion.js"></script>
			<script language="javascript">
				function fnCheckBrowser()
				{
					var agt=navigator.userAgent.toLowerCase();
					//check for ie5.5 or ie6 or 7.0
					var bIe55Up=((agt.indexOf("msie 5.5")!=-1)||(agt.indexOf("msie 6")!=-1)||(agt.indexOf("msie 7")!=-1))
					// Date object to used to get timezoneoffset
					var oToday = new Date();
					
					document.browserFrm.timezoneoffset.value=oToday.getTimezoneOffset();
					document.browserFrm.IEVersionOk.value=bIe55Up;
					document.browserFrm.cookiesEnabled.value=navigator.cookieEnabled;
					document.browserFrm.javaEnabled.value=navigator.javaEnabled();
					document.browserFrm.decimalpoint.value=fnDP();
					document.browserFrm.thousandseparator.value=fnTS();
					document.browserFrm.webversion.value=fnWebVersion();
					document.browserFrm.submit();
				}
				//decimal delimiter
				function fnDP()
				{
					var n=(1/2)
					n=n.toLocaleString();
					return n.substr(1,1);
				}
				//thousands delimiter
				function fnTS()
				{
					var n=1000;
					n=n.toLocaleString();
					return n.substr(1,1);
				}
				function fnError(sMsg,sUrl,sLine)
				{
					if(sLine==22)
					{
						document.browserFrm.webversion.value=-1;
						document.browserFrm.submit();
						return true;
					}
				}
				window.onerror=fnError;
			</script>
			<noscript>
				<html>
						<head><link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css"><title>&nbsp;<%=Application("asName")%></title>
						</head>
						<body>
						<table width='95%' align='center' valign='middle' class="clsMessageText" ID="Table2">
						<tr height='30'><td></td></tr>
						<tr><td><%=FormatDateTime(now,1) & " " & FormatDateTime(now,3)%><br><br>
						<img src='../img/ico_error_perm.gif'>&nbsp;<b><u>Browser configuration problem</u></b> 
						<br><br><br>
						MACRO requires your browser to be configured to run JavaScript. Please reconfigure your browser to 
						enable '<b>Active Scripting</b>'.
						</td></tr></table></body></html>
			</noscript>	 
			</head>
			<body onload="javascript:document.all.loadingDiv.style.visibility='visible';fnCheckBrowser();">
			<form name="browserFrm" action="login.asp?exp=<%=Server.URLEncode(sExpired)%>&db=<%=Server.URLEncode(sDatabase)%>&rl=<%=Server.URLEncode(sRole)%>&app=<%=Server.URLEncode(sAppState)%>&act=<%=Server.URLEncode(sAction)%>&att=<%=CInt(bAttempt)%>" method="post">
			<input type="hidden" name="browserChecked" value="yes">
			<input type="hidden" name="IEVersionOk">
			<input type="hidden" name="cookiesEnabled">
			<input type="hidden" name="javaEnabled">
			<input type="hidden" name="timezoneoffset">
			<input type="hidden" name="decimalpoint">
			<input type="hidden" name="thousandseparator">
			<input type="hidden" name="webversion">
			</form>
		
			<div style="visibility:hidden;" class="clsProcessingMessageBox" id="loadingDiv">
				<table height="100%" align="center" width="90%" class="clsMessageText">
				<tr>
				<td valign="middle">
				please wait<br><br>
				<img src="../img/clock.gif">
				&nbsp;&nbsp;checking browser...
				</td>
				</tr>
				</table>
			</div>
		
			</body>
			</html>
<%
		End If 
	End If


	'if browser hasnt been accepted then stop
	If (session("sbBrowserAccept") <> TRUE) Then Response.End 


	'handle possible page requests
	If (sAction = "standby") And (Session("ssUserName") <> "") Then
		'get user name from session variables
		sUserName = Session("ssUserName")
		'user is locked, clear 'user' session var
		Session("ssUser") = ""
	Elseif sAction = "switch" then
		'switch user, clear 'user' session var
		Session("ssUser") = ""
		sUserName = Request.Form("usrName")
	else
		sAction = ""
		sUserName = Request.Form("usrName")
	End If

	'get password from login form
	sUserPassword = Request.Form("usrPswd")


	'if both user name and password are supplied, attempt login
	If ((sUserName <> "") And (sUserPassword <> "")) Then
	 	
		'create instance of i/o object
		Set oIo = Server.CreateObject("MACROWWWIO30.clsWWW")
		
		'confirm password, get serialised user object
		nRtn = 0
		sUser = oIo.Login(sUserName,sUserPassword,"","","",sUserFullName,nRtn)
		set oIo = nothing
		
		if err.number <> 0 then
			call fnError(err.number,err.description,"Login.asp oIo.Login()",Array(sUserName,sUserPassword),true)
		end if

		If ((nRtn = nLOGIN_OK) Or (nRtn = nPASSWORD_EXPIRED) Or (nRtn = nCHANGE_PASSWORD)) Then
			
			'set username session variable
			session("ssUserName") = sUserName

			'set serialised user object session variable
			session("ssUser") = sUser
			
			If nRtn = nLOGIN_OK Then
				'set password session variable
				Session("ssUserPassword") = sUserPassword
				
				Response.Redirect("SelectDatabase.asp?db=" & Server.URLEncode(sDatabase) & "&rl=" & Server.URLEncode(sRole) & "&app=" & Server.URLEncode(sAppState))
			Else
				'password expired
				Response.Redirect("ChangePassword.asp?db=" & Server.URLEncode(sDatabase) & "&rl=" & Server.URLEncode(sRole) & "&app=" & Server.URLEncode(sAppState))
				
			End if
		End If
	End If

	
	'this html page submits login information back to the asp
	Response.Write("<!--" & Application("asCopyright") & "-->")
	%>
	<html>
	<head>
	<link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css">
	<title><%=Application("asName")%></title>

	<script language="javascript">
		function fnSubmit()
		{
			var sUserName = loginFrm.usrName.value;
			var sUserPassword = loginFrm.usrPswd.value;
			
			if (sUserName=="") 
			{
				alert("Please enter a valid User name");
				loginFrm.usrName.focus();
				return false;
			}
			if (sUserPassword=="") 
			{
				alert("Please enter a valid Password");
				loginFrm.usrPswd.focus()
				return false;
			}
			if (!fnAlphaNumeric(sUserName))
			{
				alert("User name may only contain alphanumeric characters")
				loginFrm.usrName.focus();
				return false;
			}
			if (!fnAlphaNumeric(sUserPassword))
			{
				alert("Password may only contain alphanumeric characters")
				loginFrm.usrPswd.focus()
				return false;
			}
			loginFrm.btnSubmit.disabled = true;
		}
		function fnAlphaNumeric(vValue)
		{
			//allow alphanumeric or space
			var	sCheck = /[^a-zA-Z0-9 ]/;
			return (sCheck.exec(vValue)==null);
		}
	</script>

	</head>
	
	<body onload="
<%
	If (sUserName = "") Then
		Response.Write("loginFrm.usrName.focus();")
	Else	
		Response.Write("loginFrm.usrPswd.focus();")
	End If
			 
	If (sUserPassword <> "") Then
		Response.Write("alert('Login Failed!');")
	End If
%>

	">
	<div><img height="100%" width="100%" src="../img/bg.jpg"></div>
	
	<div style="position:absolute; left:0; top:30%;">
	<form name="loginFrm" method="post" action="Login.asp?exp=<%=Server.URLEncode(sExpired)%>&act=<%=Server.URLEncode(sAction)%>&db=<%=Server.URLEncode(sDatabase)%>&rl=<%=Server.URLEncode(sRole)%>&app=<%=Server.URLEncode(sAppState)%>&att=1" onsubmit="return fnSubmit();">

	<table border="0" width="100%" height="100%">
	<tr><td align="center" valign="middle">
	<table border="0">
	<tr><td class="clsCopyrightText" valign="middle" align="center"><img src="../img/macrologo1.gif"> 

	</td>
	</tr>
	<tr height="15"><td>&nbsp;</td></tr>
	<tr>
	<td>
	<table border="0" cellpadding="0" cellspacing="0" class="clsLabelText" style="color:#ffffff; font-weight:bold;">
	<tr height="5"></tr>
	<%If (sAction = "standby") Then%>
		<tr height="50" valign="top"><td colspan="2">
		MACRO is in standby mode. Please re-enter your <br>password to resume working with MACRO</td></tr>
	<%
	ElseIf (sAction = "switch") Then
	%>
		<tr height="50" valign="top"><td colspan="2">
		Enter your user name and password <br>to log in to MACRO as a different user</td></tr>
	<%
	Else
		If (sExpired = "true") then
	%>
			<tr height="50" valign="top"><td colspan="2">
			A user name and password is required to access this page.<br>
			Either you followed an external link to this page,<br>
			or your MACRO session has timed out due to inactivity.<br>
			Please enter your login details to continue.<br><br></td></tr>
	<%	
		End If
	End If
	%>
	<tr><td>&nbsp;User name&nbsp;</TD>
	<td><input autocomplete="off" name="usrName" 
	<%
	if (sAction = "standby") then Response.Write(" disabled ")
	%>
	maxlength="20" SIZE="25" value="<%=sUserName%>">&nbsp</td></tr>
	<tr height="5"></tr>
	<tr><TD>&nbsp;Password&nbsp;</TD><TD>
	<input name="usrPswd" maxlength="50" size="25" type="password" autocomplete="off">&nbsp;</td></tr>
	<tr height="5"></tr>
	</table>
	</td>
	</tr>
	<tr height="15"></tr>
	<tr>
	<td align="center"><input style='width:100px;' class='clsButton' name='btnSubmit' type="submit" value="OK">
	<%If (sAction = "standby") Then%>
	<input style='width:100px;' class='clsButton'  type="button" value="Cancel" onclick="window.navigate('logout.asp')">
	<%End If%>
	</td>
	</table>
	</td></tr>
	</table>
	</form>
	</div>
	
	</body>
	
	</html>
	<%end if%>