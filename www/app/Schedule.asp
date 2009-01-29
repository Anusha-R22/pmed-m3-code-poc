<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		Forms.asp
'	Author: 	I Curtis
'	Purpose: 	displays the schedule for a requested subject
'				querystring parameters:
'					fltDb: selected database
'					fltSi: selected site
'					fltSt: selected study
'					fltSj: selected subject ID
'==================================================================================================
'	Revisions:
'	ic 29/07/2002	changed dll reference for 3.0
'	ic 27/08/2002	added extra icon handling
'	ic 03/09/2002	added form name, date, & line returns between repeating eform instances
'	ic	22/11/2002	changed www directory structure
'	dph 23/01/2003	show loading DIV
'	ic 28/06/2004 added error handling
'==================================================================================================
%>
<!--r-->
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Unlock.asp" -->
<!-- #include file="Global.asp" -->
<!--r-->
<%
dim oIo
dim sUser
dim fltSi
dim fltSt
dim fltSj
dim bNew
dim sUpdate
dim sIdentifier

	on error resume next

	sUser = session("ssUser")
	fltSi = Request.QueryString("fltSi")
	fltSt = Request.QueryString("fltSt")
	fltSj = Request.QueryString("fltSj")
	bNew = Request.QueryString("new")
	sUpdate = Request.Form("SchedUpdate")
	sIdentifier = Request.Form("SchedIdentifier")
%>
<html>
	<head>
	<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>
	<script language='javascript' src='../script/Schedule.js'></script>
	<script language='javascript' src='../script/DLAEnable.js'></script>
	<!-- #include file="HandleBrowserEvents.asp" -->
	</head>

	
	<div class="clsProcessingMessageBox" id="divMsgBox">
	<table height="100%" align="center" width="90%">
	<tr><td valign="middle" class="clsMessageText">please wait<br><br><img src="../img/clock.gif">
	&nbsp;&nbsp;Processing Schedule...</td></tr></table></div>	
<%
	Response.Flush
	'create i/o object instance
	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	Response.Write oIo.SaveAndLoadSchedule(sUser,fltSi,fltSt,fltSj,bNew,sUpdate,sIdentifier,session("TimezoneOffset"), _
	Application("abUSESCI"))
	set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"Schedule.asp oIo.SaveAndLoadSchedule()",Array(fltSi,fltSt,fltSj,bNew,sUpdate,sIdentifier, _
		session("TimezoneOffset"),Application("abUSESCI")),false)
	end if
%>
</html>