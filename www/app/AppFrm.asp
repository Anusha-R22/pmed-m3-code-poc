<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 2000. All Rights Reserved
'	File:		AppFrm.asp
'	Authors: 	i curtis
'	Purpose: 	application frameset 
'				querystring parameters:
'					fltDb: selected database
'	Version:	1.0 
'==================================================================================================
'	Revisions:
'	ic	22/11/2002	changed www directory structure
'	dph 06/10/2003	added eForm image loading
'	ic 28/06/2004 added parameter checking, error handling
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Global.asp" -->
<%
dim oIo
dim sDatabase
dim sRole
dim sAppState

	on error resume next

	Set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	sDatabase = oIo.GetUserDB(session("ssUser"))
	if err.number <> 0 then 
		set oIo = nothing
		call fnError(err.number,err.description,"AppFrm.asp oIo.GetUserDB()",Array(),true)
	end if
	
	sRole = oIo.GetUserRole(session("ssUser"))
	set oIo = nothing
	if err.number <> 0 then 
		call fnError(err.number,err.description,"AppFrm.asp oIo.GetUserRole()",Array(),true)
	end if
	sAppState = Request.QueryString("app")
	if not (fnValidateAppState(sAppState)) then sAppState = ""
	
	Response.Write("<!-- " & Application("asCopyright") & " -->")	
%>
	<html>
	<head>
	<title><%=Application("asName")%></title>
	<!-- #include file="HandleBrowserEvents.asp" -->
	<script language='javascript'>
	function fnSetAppState(sApp)
	{
		window.sAppState=sApp;
	}
	//returns an application state string
	function fnGetAppState(nWin)
	{
		var oF=window.frames;
		var sApp="";
		if (nWin==undefined)
		{
			//sApp+="1`"+oF[1].sWinState;
			sApp+=(oF[4].frames["Win0"].sWinState!=undefined)?"5|"+oF[4].frames["Win0"].sWinState:"";
			if (oF[4].fnIsSplit())
			{
				sApp+=(oF[4].frames["Win1"].sWinState!=undefined)?"`6|"+oF[4].frames["Win1"].sWinState:"";
			}
		}
		else
		{	
		}
		return sApp;
	}
	function fnGetAppUrl(bSameUser)
	{
		//omit role from url if not bSameUser
		var sUrl="db="+sUserDb;
		sUrl+=(bSameUser)?"&rl="+sUserRl:"";
		sUrl+="&app="+fnGetAppState();
		return sUrl;
	}
	fnSetAppState("<%=fnReplaceWithJSChars(sAppState)%>");
	var sUserDb="<%=sDatabase%>";
	var sUserRl="<%=sRole%>";

	// Image preloading - moved from validation engine
	var oImages = new Array();
	var aImageNames = new Array();
	var nImgDone = 0;
	// get them now
	initialiseimages();

	function loadimages()
	{
		for(n=0;n<aImageNames.length;n++)
		{
			oImages[aImageNames[n]]=new Image();
			oImages[aImageNames[n]].src="../img/"+aImageNames[n]+".gif";
		}
		checkloadedimages();
	}

	function checkloadedimages()
	{
		var nCount=0;
		var nImg=aImageNames.length;
		while(nCount<nImg && oImages[aImageNames[nCount]].complete) nCount++;
		if(nCount<nImg)
		{
			setTimeout('checkloadedimages()',100);
		}
	}

	function initialiseimages()
	{
		// initialise names list
		aImageNames[0]="ico_ok";
		aImageNames[1]="blank";
		aImageNames[2]="ico_missing";
		aImageNames[3]="ico_locked";
		aImageNames[4]="ico_frozen";
		aImageNames[5]="ico_disc_raise";
		aImageNames[6]="ico_disc_resp";
		aImageNames[7]="ico_warn";
		aImageNames[8]="ico_ok_warn";
		aImageNames[9]="ico_inform";
		aImageNames[10]="ico_uo";
		aImageNames[11]="ico_na";
		aImageNames[12]="RadioSelHi";
		aImageNames[13]="RadioSel";
		aImageNames[14]="RadioUnselHi";
		aImageNames[15]="RadioUnsel";
		aImageNames[16]="ico_change1";
		aImageNames[17]="ico_change2";	
		aImageNames[18]="ico_change3";	
		aImageNames[19]="ico_changeuser";	
		aImageNames[20]="ico_note";	
		aImageNames[21]="icof_disc_raise";	
		aImageNames[22]="icof_disc_resp";	
		aImageNames[23]="icof_error";	
		aImageNames[24]="icof_frozen";	
		aImageNames[25]="icof_inactive";	
		aImageNames[26]="icof_inform";	
		aImageNames[27]="icof_locked";	
		aImageNames[28]="icof_missing";	
		aImageNames[29]="icof_new";	
		aImageNames[30]="icof_ok";	
		aImageNames[31]="icof_ok_warn";	
		aImageNames[32]="icof_uo";	
		aImageNames[33]="icof_warn";
		aImageNames[34]="ico_sdv_plan";	
		aImageNames[35]="ico_sdv_query";
		aImageNames[36]="blank_status";
		aImageNames[37]="blank_change";
		aImageNames[38]="ico_comment";
		aImageNames[39]="ico_sdv_done";
		aImageNames[40]="ico_note_comment";
		
		// now check the are loaded
		loadimages();
	}
	</script>
	</head>

	<frameset id="f1" rows="27,*,0,0" framespacing="0" frameborder="no" border="0">
	  <frame src="appMenuTop.asp" name="" scrolling="no" marginheight="0" marginwidth="0" noresize>
	  <frameset id="f2" cols="230,*" framespacing="0" frameborder="no" border="0">
	    <frameset id="f3" rows="75,*,70" framespacing="0" frameborder="no" border="0">
	      <frame src="appHeaderLh.asp" scrolling="no" marginheight="0" marginwidth="0" noresize>
	      <frame src="appMenuLh.asp" scrolling="auto" marginheight="0" marginwidth="0">
	      <frame src="appFooterLh.asp" scrolling="no" marginheight="0" marginwidth="0" noresize>
        </frameset>
	    <frame src="MainBorder.asp" scrolling="auto" marginheight="0" marginwidth="0">
      </frameset>
      <frame src="Symbols.htm" name="" scrolling="no" marginheight="0" marginwidth="0" noresize>
      <frame src="FunctionKeys.htm" name="" scrolling="no" marginheight="0" marginwidth="0" noresize>  
    </frameset>
    
	</body>
	</html>