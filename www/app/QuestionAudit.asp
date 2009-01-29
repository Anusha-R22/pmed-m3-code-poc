<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		QuestionAudit.asp
'	Author: 	I Curtis
'	Purpose: 	
'==================================================================================================
'	Revisions:
'	DPH 06/11/2002 - Repeat No Added
'	ic	22/11/2002	changed www directory structure
'	ic 28/06/2004 added error handling
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="DialogLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<%
dim fltSt
dim fltSi
dim fltEf
dim fltId
dim fltSj
dim fltFd
dim fltRp
dim sCaption
dim sRtn
dim oIo

	on error resume next

	fltSt = Request.QueryString("fltSt")
	fltSi = Request.QueryString("fltSi")
	fltEf = Request.QueryString("fltEf")
	fltId = Request.QueryString("fltId")
	fltSj = Request.QueryString("fltSj")
	fltFd = Request.QueryString("fltFd")
	sCaption = Request.QueryString("caption")
	' Repeat zero based so add one
	fltRp = cstr(cint(Request.QueryString("fltRp"))+1)

	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	sRtn = oIo.GetQuestionAuditHtml(session("ssUser"),fltSt,fltSi,fltSj,fltId,fltFd,sCaption,session("ssDecimalPoint"), _
	session("ssThousandSeparator"),fltRp)
	set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"QuestionAudit.asp oIo.GetQuestionAuditHtml()",Array(fltSt,fltSi,fltSj,fltId,fltFd,sCaption, _
		session("ssDecimalPoint"),session("ssThousandSeparator"),fltRp),true)
	end if
	
	Response.Write(sRtn)
%>
	