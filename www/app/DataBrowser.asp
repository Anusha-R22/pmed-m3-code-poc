<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		dataBrowser.asp
'	Author: 	I Curtis
'	Purpose: 	Displays data browser table. allows user to freeze,lock,unlock data. allows user to
'				add discrepancy,note,sdv
'				querystring parameters:
'					fltDb: selected database
'					fltSt: search study
'					fltSi: search site
'					fltVi: search visit
'					fltEf: search eform
'					fltQu: search question
'					fltUs: search user
'					fltSj: search subject label
'					fltSs: search status binary string
'					fltTm: search time
'					fltB4: search before time
'					fltAd: search additional parameters
'					recs: number of records to display per page
'					bookmark: record number to begin display at
'					current: show current data/audit trail
'==================================================================================================
'	Revisions:
'	ic 11/10/2001	bug 033 fixed type mismatch when comparing values from an oracle db
'	ic 29/07/2002	changed dll reference for 3.0
'	ic 09/09/2002	switched inline constants for constants include file
'	ic	22/11/2002	changed www directory structure
'	ic 28/06/2004 added parameter checking, error handling
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<%
	'only unlock any eform/visit if not in split screen mode
	if not session("sbSplit") then
%>
	<!-- #include file="Unlock.asp" -->
<%
	end if

dim oIo
dim fltDb
dim fltSt
dim fltSi
dim fltVi
dim fltEf
dim fltQu
dim fltUs
dim fltSj
dim fltSjLb
dim fltSs
dim fltLk
dim fltTm
dim fltB4
dim fltAd
dim fltDi
dim fltSd
dim fltNo
dim sForm
dim bookmark
dim sGet

	on error resume next

	fltSt = Request.QueryString("st")			
	fltSi = Request.QueryString("si")			
	fltVi = Request.QueryString("vi")			
	fltEf = Request.QueryString("ef")			
	fltQu = Request.QueryString("qu")			
	fltUs = Request.QueryString("us")			
	fltSj = Request.QueryString("sj")			
	fltSjLb = Request.QueryString("sjlb")		
	fltSs = Request.QueryString("ss")			
	fltLk = Request.QueryString("lk")
	fltTm = Request.QueryString("tm")			
	fltB4 = Request.QueryString("b4")			
	fltAd = Request.QueryString("cm")			
	fltDi = Request.QueryString("di")
	fltSd = Request.QueryString("sd")
	fltNo = Request.QueryString("no")
	bookmark = Request.QueryString("bookmark")		
	sGet = Request.QueryString("get")				
	sForm = Request.Form
	
%>
	<html>
	<head>
	<link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css">
	<script language='javascript' src='../script/RegistrationDisable.js'></script>
	<script language='javascript' src='../script/DLAEnable.js'></script>
    <script language="javascript" src="../script/DataBrowser.js"></script>
    <!-- #include file="HandleBrowserEvents.asp" -->
	</head>
	
	<div class="clsProcessingMessageBox" id="divMsgBox">
	<table height="100%" align="center" width="90%">
	<tr><td valign="middle" class="clsMessageText">please wait<br><br><img src="../img/clock.gif">
	&nbsp;&nbsp;Processing DataBrowser...</td></tr></table></div>
	
	<div class="clsPopMenu" id="divPopMenu" onclick="menu=this;this.tid=setTimeout('menu.style.visibility=\'hidden\'',20);" 
	onmouseout="menu=this;this.tid=setTimeout('menu.style.visibility=\'hidden\'',20);" onmouseover="clearTimeout(this.tid);">
	</div>
	
<%
	Response.Flush
	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	Response.Write(oIo.SaveAndLoadDataBrowser(cstr(session("ssUser")),cstr(fltSt),cstr(fltSi),cstr(fltVi),cstr(fltEf), _
	cstr(fltQu),cstr(fltUs),cstr(fltSj),cstr(fltSjLb),cstr(fltSs),cstr(fltLk),cstr(fltTm),cstr(fltB4),cstr(fltAd),cstr(fltDi), _
	cstr(fltSd),cstr(fltNo),cstr(sGet),cstr(bookmark),cstr(sForm),cstr(session("TimezoneOffset")), _
	cstr(session("ssDecimalPoint")),cstr(session("ssThousandSeparator"))))
	set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"DataBrowser.asp oIo.SaveAndLoadDataBrowser()",Array(cstr(fltSt),cstr(fltSi),cstr(fltVi),cstr(fltEf), _
		cstr(fltQu),cstr(fltUs),cstr(fltSj),cstr(fltSjLb),cstr(fltSs),cstr(fltLk),cstr(fltTm),cstr(fltB4),cstr(fltAd),cstr(fltDi), _
		cstr(fltSd),cstr(fltNo),cstr(sGet),cstr(bookmark),cstr(sForm),cstr(session("TimezoneOffset")), _
		cstr(session("ssDecimalPoint")),cstr(session("ssThousandSeparator"))),false)
	end if
%>

	</html>