<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		DialogPrintEform.asp
'	Author: 	I Curtis
'	Purpose: 	prints an eform, defined in passed 'state' querystring param
'==================================================================================================
'	Revisions:
'	ic 28/06/2004 added error handling
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="DialogLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<%
dim oIo
dim sUser
dim fltSt
dim fltSi
dim fltSj
dim fltVi
dim fltVId
dim fltVTId
dim fltId
dim fltTId
dim sRtn

	on error resume next

	sUser = Session("ssUser")
	fltSt = Request.QueryString("fltSt")
	fltSi = Request.QueryString("fltSi")
	fltSj = Request.QueryString("fltSj")
	fltVi = Request.QueryString("fltVi")
	fltVId = Request.QueryString("fltVId")
	fltVTId = Request.QueryString("fltVTId")
	fltId = Request.QueryString("fltId")
	fltTId = Request.QueryString("fltTId")
	
	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	Response.Write(oIo.GetEformPrintHTML(sUser,fltSi,fltSt,fltSj,fltVi,fltVId,fltVTId,fltId,fltTId,session("ssDecimalPoint"),_
	session("ssThousandSeparator")))
	set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"DialogPrintEform.asp oIo.GetEformPrintHTML()",Array(fltSi,fltSt,fltSj,fltVi,fltVId,fltVTId,fltId,_
		fltTId,session("ssDecimalPoint"),session("ssThousandSeparator")),true)
	end if
%>