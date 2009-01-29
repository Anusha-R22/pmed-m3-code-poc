<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = false%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		CreateSubject.asp
'	Author: 	I Curtis
'	Purpose: 	creates a new subject and redirects to schedule
'				querystring parameters:
'					fltDb: selected database
'					fltSi: selected site
'					fltSt: selected study
'==================================================================================================
'	Revisions:
'	ic	22/11/2002	changed www directory structure
'	ic 28/06/2004 added error handling
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Unlock.asp" -->
<!-- #include file="Global.asp" -->
<%
CONST UNSPECIFIED_ERROR = "0"
CONST NO_ACCESS = "-1"
CONST SUBJECT_LOCKED = "-2"
CONST STUDY_LOCKED = "-3"

dim fltSi
dim fltSt
dim fltSj
dim oIo
dim sMsg
dim sUrl

	on error resume next
 
	fltSi = Request.QueryString("fltSi")
	fltSt = Request.QueryString("fltSt")
	fltSj = Request.QueryString("fltSj")
%>
	<html>
	<head>
	<title></title>
	<link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css">
	<!-- #include file="HandleBrowserEvents.asp" -->
	</head>
	<body onload="javascript:fnPageLoaded();">
	
	<div style="visibility:visible;" class="clsProcessingMessageBox" id="divLoading">
	<table height="100%" align="center" width="90%" class="clsMessageText">
	<tr><td valign="middle">please wait<br><br><img src="../img/clock.gif">
	&nbsp;&nbsp;creating MACRO subject...</td></tr></table></div>

<%
	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	fltSj = oIo.CreateNewSubject(session("ssUser"),fltSi,fltSt,Application("abUSESCI"))
	set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"CreateSubject.asp oIo.CreateNewSubject()",Array(fltSi,fltSt,Application("abUSESCI")),false)
	end if
	

	select case cstr(fltSj)
	case UNSPECIFIED_ERROR:
		sMsg = "An unspecified error occurred, MACRO was unable to create a new subject"
	case NO_ACCESS:
		sMsg = "You do not have permission to create a subject at the requested site"
	case STUDY_LOCKED:
		sMsg = "The selected study is currently locked, MACRO was unable to create a new subject"
	case else:
		'success
	end select
	
%>
	<html>
	<HEAD>
	<link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css">
	<title></title>
	</HEAD>
	<body>
	<div class="clsErrorText"><%=sMsg%></div>
	</body>
	</html>		
	<script language="">
	  function fnPageLoaded()
	  {
		document.all.divLoading.style.visibility='hidden';
<%		if sMsg = "" then
%>			window.parent.fnScheduleUrl('<%=fltSt%>','<%=fltSi%>','<%=fltSj%>',undefined,undefined,"1");
			//window.parent.fnRefreshMenu();
<%		end if
%>	  }
	</script>
