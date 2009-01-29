<%@ LANGUAGE=VBScript%>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1
'==================================================================================================
' 	Copyright:	InferMed Ltd. 2000. All Rights Reserved
'	File:		Eform.asp
'	Authors: 	i curtis
'	Purpose: 	eform loader 
'				
'==================================================================================================
'	Revisions:
'	ic	22/11/2002	changed www directory structure
'==================================================================================================
%>
<!-- #include file="checkSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->

<HTML>
<HEAD>
<link rel="stylesheet" HREF="../style/MACRO1.css" type="text/css">
<script language="javascript" src="../script/LoadURL.js"></script>
<!-- #include file="HandleBrowserEvents.asp" -->
<script language='javascript'>
var bSplit=false;
function fnResize()
{
	var nSpacing=18; //top corner bluecurve height
	nSpacing+=15; //top frame header
	nSpacing+=(bSplit)?10:0; //central spacer
	nSpacing+=(bSplit)?15:0; //bottom frame header
	nSpacing+=22; //bottom spacer
	var nMinWidth=150; //minimum screen width before stop resizing
	var nMinHeight=100; //minimum screen height before stop resizing

	if (window.document.body.clientWidth>(36+nMinWidth))
	{
		window.document.all["Win0"].style.width=window.document.body.clientWidth-36;
		window.document.all["Win1"].style.width=window.document.body.clientWidth-36;
	}
	if (window.document.body.clientHeight>(nSpacing+nMinHeight))
	{
		if(bSplit)
		{
			window.document.all["Win0"].style.height=(window.document.body.clientHeight-nSpacing)/2;
		}
		else
		{
			window.document.all["Win0"].style.height=(window.document.body.clientHeight-nSpacing);
		}
	}
	if (bSplit)
	{
		if (window.document.body.clientHeight>(nSpacing+nMinHeight))
		{
			window.document.all["Win1"].style.height=(window.document.body.clientHeight-nSpacing)/2;
		}
	}
}
function fnSetTitle(sTitle,nWin)
{
	nWin=(nWin==undefined)? 0:nWin;
	document.all["divTitle"+nWin].innerHTML=sTitle;
}
function fnShowWin(bSplitWin)
{
	if (bSplitWin==bSplit) return;
	if (bSplitWin)
	{
		window.document.all["Win1"].style.visibility="visible";
		window.document.all["trSpace"].style.height=10;
		window.document.all["trHead"].style.visibility="visible";
		window.document.all["trHead"].style.height=15;
		window.document.all["imgCurve"].src="../img/bluecurve2.jpg"
	}
	else
	{
		window.document.all["Win1"].style.height=0;
		window.document.all["Win1"].style.visibility="hidden";
		window.document.all["trSpace"].style.height=0;
		window.document.all["imgCurve"].src="../img/blank.gif"
		window.document.all["trHead"].style.visibility="hidden";
		window.document.all["trHead"].style.height=0;
	}
	bSplit=bSplitWin;
	fnResize();
}
window.onresize=fnResize;
</script>
<TITLE></TITLE>
</HEAD>
<BODY onload='fnResize();'>

<table cellpadding='0' cellspacing='0' border='0'>

  <tr><td align="left" valign="top"><img src="../img/bluecurve.jpg"></td></tr>
  
  <tr height='15'>
    <td width='18' rowspan='2'></td>
    <td bgcolor='blue'>
      <table cellpadding='0' cellspacing='0'>
      <td width='20' align="left" valign="top" bgcolor="blue"><img src="../img/bluecurve2.jpg"></td>
      <td width='250' bgcolor="blue" align='left'><div id="divTitle0" class='clsMainBorderText'>Home</div></td>
	  </table>
	</td>
  </tr>
  <tr>
    <td>
	  <iframe id='Win0' frameborder='0' src='Home.asp' style='border-left:blue 1px solid; border-right:blue 1px solid; border-top:blue 1px solid; border-bottom:blue 1px solid;'></iframe>
    <td>
  </tr>

  <tr id='trSpace' height="0"><td></td></tr>

  <tr id='trHead' height='0' style='visibility:hidden;'>
    <td width='18' rowspan='2'></td>
    <td bgcolor='blue'>
      <table cellpadding='0' cellspacing='0'>
      <td width='20' align="left" valign="top" bgcolor="blue"><img id="imgCurve" src="../img/blank.gif"></td>
      <td width='250' bgcolor="blue" align='left'><div id="divTitle1" class='clsMainBorderText'></div></td>
	  </table>
	</td>
  </tr>
  <tr>
    <td>
	  <iframe id='Win1' frameborder='0' src='Blank.htm' style='height:0; visibility:hidden; border-left:blue 1px solid; border-right:blue 1px solid; border-top:blue 1px solid; border-bottom:blue 1px solid;'></iframe>
    <td>
  </tr>
</table>



</BODY>
</HTML>