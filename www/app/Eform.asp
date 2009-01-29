<%@ LANGUAGE=VBScript%>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1
'==================================================================================================
' 	Copyright:	InferMed Ltd. 2000. All Rights Reserved
'	File:		Eform.asp
'	Authors: 	i curtis
'	Purpose: 	eform loader 
'				
'==================================================================================================
'	Revisions:
'	ic	22/11/2002	changed www directory structure
'	ic 28/06/2004 added error handling
'==================================================================================================
%>
<!-- #include file="checkSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<%
dim sUser
dim oIo
dim fltSt
dim fltSi
dim fltSj
dim fltId
dim fltXId
dim sForm
dim sReadOnly
dim sLabCode
dim sNext
dim aRtn
dim sLocalDate

	on error resume next

	sUser = Session("ssUser")
	fltSt = Request.QueryString("fltSt")
	fltSi = Request.QueryString("fltSi")
	fltSj = Request.QueryString("fltSj")
	fltId = Request.QueryString("fltId")
	fltXId = Request.QueryString("fltXId")

	sForm = Request.Form()
	sLabCode = Request.Form("labcode")
	sReadOnly = Request.Form("readonly")
	sNext = Request.Form("next")
	sLocalDate = Request.Form("localdate")
	
%>
	<html>
	<head>
	<script language="javascript" src="../script/RadioControl.js"></script>
	<script language="javascript" src="../script/ValidationEngine.js"></script>
	<script language="javascript" src="../script/Eform.js"></script>
	<script language="javascript" src="../script/RQG.js"></script>
	<link rel="stylesheet" href="../style/MACRO1.css" type="text/css">
	
	<%if (Application("asDEV") <> "true") then%>
		<script language="javascript">
		document.oncontextmenu=fnContextMenu;
		function fnContextMenu(){return false};
		</script>
	<%end if%>
	</head>

	
	<div class="clsProcessingMessageBox" id="divMsgBox">
	<table height="100%" align="center" width="90%">
	<tr><td id="tdMsg" valign="middle" class="clsMessageText">please wait<br><br><img src="../img/clock.gif">
	&nbsp;&nbsp;Processing eForm...</td></tr></table></div>
	
	<div class="clsPopMenu" id="divPopMenu" onclick="menu=this;this.tid=setTimeout('fnHideSelects(0);menu.style.visibility=\'hidden\'',20);" 
	onmouseout="menu=this;this.tid=setTimeout('fnHideSelects(0);menu.style.visibility=\'hidden\'',20);" onmouseover="clearTimeout(this.tid);">
	</div>
	
<%
	Response.Flush
	set oIo = Server.CreateObject("MACROWWWIO30.clsWWW")
	aRtn = oIo.SaveAndLoadEform(sUser,fltSi,fltSt,fltId,fltXId,fltSj,sForm,cstr(Session("ssEFIToken")),cstr(session("ssVIToken")), _
	sLabCode,sReadOnly,sNext,session("TimezoneOffset"),Application("abUSESCI"),session("ssDecimalPoint"), _
	session("ssThousandSeparator"),sLocalDate)
	set oIo = nothing

	if err.number <> 0 then 
		call fnError(err.number,err.description,"Eform.asp oIo.SaveAndLoadEform()",Array(fltSi,fltSt,fltId,fltXId,fltSj,sForm,cstr(Session("ssEFIToken")), _
		cstr(session("ssVIToken")),sLabCode,sReadOnly,sNext,session("TimezoneOffset"),Application("abUSESCI"), _
		session("ssDecimalPoint"),session("ssThousandSeparator"),sLocalDate),false)
	end if

	Session("ssUser") = sUser
	Session("ssEFIToken") = aRtn(1)
	session("ssVIToken") = aRtn(2)

	Response.Write(aRtn(0))
%>
</html>