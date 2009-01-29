<%@ Language=VBScript %>
<%Option Explicit
'==================================================================================================
'   Copyright:  InferMed Ltd. 2006. All Rights Reserved
'   File:       LoginDirect.asp
'   Author:     I Curtis
'   Purpose:    MACRO web de/dr direct login page
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

Dim oIo
dim sUser	

Dim bAttempt
dim lc

dim decryptedUsername
dim decryptedPassword

	on error resume next

	bAttempt = cbool(Request.QueryString("att"))
	
    'get posted username and password
	if (Request.Form("username") <> "") and (Request.Form("password") <> "") then
	    'decrypt username and password
        set lc = server.CreateObject("TrialsUsers.LifeCrypt")
    
        decryptedUsername = lc.Decrypt(Request.Form("username"))
        decryptedPassword = lc.Decrypt(Request.Form("password"))
        
        'Check for ERRORS
        If Left(decryptedUsername, "4") = "ERR_" Or Left(decryptedPassword, "4") = "ERR_" then
	        'You don't need to check the specific error code. We added the specific error codes in case
	        'we want to debug, then it's nice to see where the error comes from.
	        Response.Write "Sorry, we cannot process your request. Please contact the site administrator."
	        response.End
	    else
	        session("DLUsername") = decryptedUsername
            session("DLPassword") = decryptedPassword
        end if
	end if
	

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
					//check for ie5.5 or ie6 or ie7.0
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
			<form name="browserFrm" action="LoginDirect.asp?att=<%=CInt(bAttempt)%>"% method="post">
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


    sUserName = session("DLUsername")
    sUserPassword = session("DLPassword")
    session("DLUsername") = ""
    session("DLPassword") = ""
    
    
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
				
				Response.Redirect("SelectDatabase.asp")
			Else
				'password expired
				Response.Redirect("ChangePassword.asp")
				
			End if
		End If
	End If

%>

<html>
<head><link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css"><title>&nbsp;<%=Application("asName")%></title>
</head>
<body class="clsMessageText">
<br>
&nbsp;Username/password missing or incorrect. Please click <a href="Login.asp">here</a> to login
</body>
</html>
