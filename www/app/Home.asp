<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%'Response.redirect "../reports/home.asp"%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 2000. All Rights Reserved
'	File:		home.asp
'	Authors: 	i curtis
'==================================================================================================
'	Revisions:
'	ic	22/11/2002	changed www directory structure
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Unlock.asp" -->

	<HTML>
	<HEAD>
	<link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css">
	<script language='javascript' src='../script/RegistrationDisable.js'></script>
	<script language='javascript' src='../script/DLAEnable.js'></script>
	<!-- #include file="HandleBrowserEvents.asp" -->
	<title></title>
	</head>
	<body onload='window.navigate("../reports/home.asp")'>
	</body>
	</html>