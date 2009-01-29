<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = false%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
'	copyright:		InferMed Ltd 2001. all rights reserved
'	file:			MIMessageTop.asp
'	date:			18/12/2001
'	author:			i curtis
'	purpose:		displays rows of filtered discrepancies. depending on user role and requested
'					module user can respond, close, edit, re-raise. selecting a rows radio button
'					reloads the audit frame below with the selected discrepancies history
'	version:		0.1
'	amendments:		
'	ic 19/07/2002	cbb 2.2.20 #1 added clause to fnShowInput() js function
'	ic 29/07/2002	changed dll reference for 3.0
'	dph 05/11/2002	Added repeat number to question titles
'	ic	22/11/2002	changed www directory structure
'	dph 12/02/2003	Reinstated subject id for search from eForm
' ic 05/07/2005		added visit, eform, question cycle
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<%
dim oIo
dim nMiMessageType
dim sUser
dim fltSt			
dim fltSi			
dim fltVi			
dim fltEf			
dim fltQu			
dim fltUs
dim fltSj			
dim fltSjLb			
dim fltB4			
dim fltTm			
dim fltSs
dim fltObj
dim bookmark
dim sForm
dim bNewWindow
dim fltVRpt
dim fltERpt
dim fltQRpt


	on error resume next

	nMiMessageType = Request.QueryString("type")
	fltSt = Request.QueryString("fltSt")
	fltSi = Request.QueryString("fltSi")
	fltVi = Request.QueryString("fltVi")
	fltEf = Request.QueryString("fltEf")
	fltQu = Request.QueryString("fltQu")
	fltUs = Request.QueryString("fltUs")
	fltSj = Request.QueryString("fltSj")
	fltSjLb = Request.QueryString("fltSjLb")
	fltB4 = Request.QueryString("fltB4")
	fltTm = Request.QueryString("fltTm")
	fltSs = Request.QueryString("fltSs")
	fltObj = Request.QueryString("fltObj")
	bookmark = Request.QueryString("bookmark")
	sUser = session("ssUser")
	sForm = Request.Form
	bNewWindow = Request.QueryString("newwin")
	fltVRpt = Request.QueryString("fltVRpt")
	fltERpt = Request.QueryString("fltERpt")
	fltQRpt = Request.QueryString("fltQRpt")

%>
	<html>
	<head>
	<link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css">
	<script language='javascript' src='../script/MIMessage.js'></script>
	<script language='javascript' src='../script/RowOver.js'></script>
	<!-- #include file="HandleBrowserEvents.asp" -->
	</head>
	
	<div class="clsProcessingMessageBox" id="divMsgBox">
	<table height="100%" align="center" width="90%">
	<tr><td valign="middle" class="clsMessageText">please wait<br><br><img src="../img/clock.gif">
	&nbsp;&nbsp;Processing MIMessage Browser...</td></tr></table></div>

	<div class="clsPopMenu" id="divPopMenu" onclick="menu=this;this.tid=setTimeout('menu.style.visibility=\'hidden\'',20);"
    onmouseout="menu=this;this.tid=setTimeout('menu.style.visibility=\'hidden\'',20);" onmouseover="clearTimeout(this.tid);">
    </div>
<%

	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	Response.Write(oIo.SaveAndLoadMIMessageList(sUser,nMiMessageType,fltSt,fltSi,fltVi,fltVRpt,fltEf,fltERpt,fltQu,fltQRpt,fltUs,fltSj,fltSjLb, _
	fltSs,fltTm,fltB4,fltObj,bookmark,sForm,session("TimezoneOffset"),bNewWindow,cstr(Session("ssVIToken")),cstr(Session("ssEFIToken"))))
	set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"MIMessageTop.asp oIo.SaveAndLoadMIMessageList()",Array(nMiMessageType,fltSt,fltSi,fltVi,fltVRpt,fltEf, _
		fltERpt,fltQu,fltQRpt,fltUs,fltSj,fltSjLb,fltSs,fltTm,fltB4,fltObj,bookmark,sForm,session("TimezoneOffset"),bNewWindow),false)
	end if
%>
	</html>	