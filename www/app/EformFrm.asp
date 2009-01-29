<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = true%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		eformFrm.asp
'	Author: 	I Curtis
'	Purpose: 	frame holder for data entry, contains eform menu, eform
'				querystring parameters:
'					fltDb: selected database
'					fltSi: selected site
'					fltSt: selected study
'					fltSj: selected subject ID
'					fltEf: selected eform id
'					fltId: selected eformtaskid
'					rtnUrl: url to return to
'	Version:	1.0
'==================================================================================================
'	Revisions:
'	ic	22/11/2002	changed www directory structure
'	ic 28/06/2004 added parameter checking
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<%
dim fltSi
dim fltSt
dim fltSj
dim fltEf
dim fltId
dim sCancelState

	fltSi = Request.QueryString("fltSi")
	fltSt = Request.QueryString("fltSt")
	fltSj = Request.QueryString("fltSj")
	fltEf = Request.QueryString("fltEf")
	fltId = Request.QueryString("fltId")
	sCancelState = Request.QueryString("cancelstate")

	'validate parameters that we will write into the page
	if not (fnValidateSite(fltSi)) then
		call fnError(cstr(lVBObjectError + 2),"Output terminated. Illegal site parameter:" & fltSi,"Eform.asp",Array(),true)
	end if
	if not isnumeric(fltSt) then
		call fnError(cstr(lVBObjectError + 2),"Output terminated. Illegal study parameter:" & fltSt,"Eform.asp",Array(),true)
	end if
	if not isnumeric(fltSj) then
		call fnError(cstr(lVBObjectError + 2),"Output terminated. Illegal subjectid parameter:" & fltSj,"Eform.asp",Array(),true)
	end if
	if not isnumeric(fltEf) and fltEf <> "" then
		call fnError(cstr(lVBObjectError + 2),"Output terminated. Illegal eformtaskid parameter:" & fltEf,"Eform.asp",Array(),true)
	end if
	if not isnumeric(fltId) and fltId <> "" then
		call fnError(cstr(lVBObjectError + 2),"Output terminated. Illegal id parameter:" & fltId,"Eform.asp",Array(),true)
	end if
	if not fnValidateAppState(sCancelState) then sCancelState = ""

%>
	<html>
	<head>
	<!-- #include file="HandleBrowserEvents.asp" -->
<%if sCancelState <> "" then%>	
	<script language='javascript'>
	//store cancel state in parent window
	window.parent.sCancelState="<%=fnReplaceWithJSChars(sCancelState)%>";
	window.parent.sBackState="<%=fnReplaceWithJSChars(sCancelState)%>";
	</script>
<%end if%>
	<script language='javascript'>
	var bLoading=true;
	</script>
	<title></title>
	</head>
	<frameset id='f0' rows="75,*" framespacing="0" frameborder="no" border="4">
	<frame src="EformTop.htm" name="FRAME0" scrolling="no" marginheight="0" marginwidth="0">
	<frameset id='f1' cols="170,*" framespacing="0" frameborder="no" border="0">
	  <frame src="EformLh.htm" name="FRAME1" scrolling="auto" marginheight="0" marginwidth="0">
	    <frame name="FRAME2" src="Eform.asp?fltSi=<%=fltSi%>&fltSt=<%=fltSt%>&fltEf=<%=fltEf%>&fltId=<%=fltId%>&fltSj=<%=fltSj%>" scrolling="auto" marginheight="0" marginwidth="0">
      </frameset>
    </frameset>
	
	</body>
	</html>