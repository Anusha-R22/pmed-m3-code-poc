<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 2000. All Rights Reserved
'	File:		ListSubject.asp
'	Authors: 	i curtis
'	Purpose: 	Allow the user open an existing subject or create a new one 
'				querystring parameters:
'					fltDb: selected database
'					fltSi: site filter
'					fltSt: study filter
'					fltLb: label filter
'					fltId: subject ID filter
'					nextaction=refresh: refreshes subject list 
'==================================================================================================
'	Revisions:
'	ic 29/07/2002	changed dll reference for 3.0
'	ic 06/09/2002	added non-buffer header
'					added paging, row-highlighting to subject list
'	ic 12/09/2002	split new subject code into separate file
'	dph 06/11/2002	MACROUser Object
'	ic	22/11/2002	changed www directory structure
'	dph 20/12/2002	
'	ic 28/06/2004 added error handling
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Unlock.asp" -->
<!-- #include file="HandleBrowserEvents.asp" -->
<!-- #include file="Global.asp" -->
<%

dim oIo
dim sUser
dim sFltSi
dim sFltSt
Dim sFltLb
Dim sFltId
dim sOrderBy
dim nBookmark

	on error resume next

	sUser = session("ssUser")
	sFltSi = Request.Querystring("fltSi")
	sFltSt = Request.Querystring("fltSt")
	sFltLb = Request.Querystring("fltLb")
	sFltId = Request.Querystring("fltId")
	sOrderBy = Request.QueryString("orderby")
	nBookmark = Request.QueryString("bookmark")
	
	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	Response.Write(oIo.GetSubjectListHTML(sUser,sFltSi,sFltSt,sFltLb,sFltId,sOrderBy,"true",nBookmark))
	set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"ListSubject.asp oIo.GetSubjectListHTML()",Array(sFltSi,sFltSt,sFltLb,sFltId,sOrderBy,nBookmark),true)
	end if
	
	'DPH 20/12/2002 - Write Serialised user (with changed state) back to session var
	session("ssUser") = sUser
%>
