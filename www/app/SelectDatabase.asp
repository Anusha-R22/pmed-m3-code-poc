<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       SelectDatabase.asp
'   Author:     i curtis
'   Purpose:    Allows User to choose a database
'				querystring parameters:
'					module=[de][dr]: module selection from login page
'==================================================================================================
'	Revisions:
'	ic 29/06/01	amendments for 2.2
'	ic 29/07/2002	changed dll reference for 3.0
'	ic 10/10/2002	changed to list db and role choices
'	ic 22/11/2002	changed www directory structure
'	ic 12/12/2003	added comments, application state check
'	ic 23/06/2004	added parameter checking, error handling
'==================================================================================================
'
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<%
Const nLOGIN_OK = 0
Const nACCOUNT_DISABLED = 1
Const nLOGIN_FAILED = 2
Const nCHANGE_PASSWORD = 3
Const nPASSWORD_EXPIRED = 4

Dim oIo
dim sDatabase
dim sRole
dim sAppState
dim nRtn
dim sRtn
dim sUser
dim sReqDatabase
Dim nLoop
dim sCurrDb
Dim sDatabaseString
Dim sRoleString
dim aRtn

	on error resume next
	
	'get requested database, role and application state
	'these 3 items may have been passed in the querystring from
	'the login page, or from the page itself by a form 'get'
	sDatabase = Request.QueryString("db")
	sRole = Request.QueryString("rl")
	sAppState = Request.QueryString("app")

	'validate parameters that we will write into the page
	if not (fnValidateDatabase(sDatabase)) then sDatabase = ""
	if not (fnValidateRole(sRole)) then sRole = ""
	if not (fnValidateAppState(sAppState)) then sAppState = ""
	
	Set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	
	'a requested database and role were passed
	if ((sDatabase <> "") and (sRole <> "")) then
		nRtn = 0
		'try to log user in to requested database
		sUser = oIo.Login("","",sDatabase,sRole,session("ssUser"),"",nRtn)
		
		if err.number <> 0 then 
			call fnError(err.number,err.description,"SelectDatabase.asp oIo.Login()",Array(sDatabase,sRole),true)
		end if
		
		if (nRtn = nLOGIN_OK) then
			'if login was successful, user has the role on the database,
			'update the session user string and clear the password that
			'was passed from the login page
			session("ssUser") = sUser
			session("ssUserPassword") = ""
			
			set oIo = nothing
			'redirect to the application, passing any application state string
			Response.Redirect("AppFrm.asp?app=" & Server.URLEncode(sAppState))
		else
			'the database login failed. this may happen if a user has illegally
			'manipulated the form in some way
			'Session("ssUser") = ""
			'Session("ssUserPassword") = ""
			'Response.End
			sDatabase = ""
			sRole = ""
		end if
	end if
	
	'remember the database requested, if any
	'this so we can check that the actual database the user
	'logs in to matches. if it doesnt match they cant return
	'to an application state
	sReqDatabase = sDatabase
	
	'get database choice html, or single database/role choice, or error html
	sRtn = oIo.GetDatabaseChoiceHTML(Session("ssUserName"),session("ssUserPassword"),sDatabase,sRole,sAppState)
	Set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"SelectDatabase.asp oIo.GetDatabaseChoiceHTML()",Array(sDatabase,sRole),true)
	end if

	if ((sDatabase <> "") and (sRole <> "")) then
		'if the user wasnt logged into their requested database, dont
		'allow user to return to application state
		if (sReqDatabase <> sDatabase) then
			sAppState = ""
		end if
		'if we now have a database and role, there is a single database/role choice
		Response.Redirect "SelectDatabase.asp?db=" & Server.URLEncode(sDatabase) & "&rl=" & Server.URLEncode(sRole) & "&app=" & Server.URLEncode(sAppState)
	end if

	'else display html error/choice page
%>
	<html>
    <head>
    <title>InferMed MACRO</title>
    <link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css">
    <script language='javascript' src='../script/SelectDatabase.js'></script>
    <!-- #include file="HandleBrowserEvents.asp" -->
	</head>

<%	
	Response.Write(sRtn)
%>
	</html>	