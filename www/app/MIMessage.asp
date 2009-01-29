<%@ Language=VBScript %>
<%Option Explicit
'==================================================================================================
'	copyright:		InferMed Ltd 2001. all rights reserved
'	file:			MIMessage.asp
'	date:			28/06/2001
'	author:			ilc
'	purpose:		holds 2 frames. top frame is edit frame. bottom frame is audit frame 
'	version:		0.1
'	amendments:		
'	ic	22/11/2002	changed www directory structure
'	DPH 12/02/2003	Can open in new window so don't do unlock if this is the case / title
'	ic 20/07/2004		fixed locking on mimessage updates
' ic 05/07/2005		added visit, eform, question cycle
'==================================================================================================
%>
<!-- #include file="CheckSSL.asp" -->
<!-- #include file="WindowLoggedIn.asp" -->
<!-- #include file="Global.asp" -->

<%
dim sNewWin
dim nMiMessageType
dim fltSt			
dim fltSi			
dim fltVi			
dim fltEf			
dim fltQu			
dim fltUs
dim fltSj			
dim fltSjLb			
dim fltB4			
dim fltTm			
dim fltSs
dim fltObj
dim bookmark
dim bNewWindow
dim fltVRpt
dim fltERpt
dim fltQRpt


	'is this page opening in a separate window
	bNewWindow = Request.QueryString("newwindow")
	
	'only unlock any eform/visit if not in split screen mode
	'and not viewing from an eform
	if (not session("sbSplit")) and (not bNewWindow = "1") then
%>
	<!-- #include file="Unlock.asp" -->
<%
	end if
	
	nMiMessageType = Request.QueryString("type")
	fltSt = Request.QueryString("fltSt")
	fltSi = Request.QueryString("fltSi")
	fltVi = Request.QueryString("fltVi")
	fltEf = Request.QueryString("fltEf")
	fltQu = Request.QueryString("fltQu")
	fltUs = Request.QueryString("fltUs")
	fltSj = Request.QueryString("fltSj")
	fltSjLb = Request.QueryString("fltSjLb")
	fltB4 = Request.QueryString("fltB4")
	fltTm = Request.QueryString("fltTm")
	fltSs = Request.QueryString("fltSs")
	fltObj = Request.QueryString("fltObj")
	bookmark = Request.QueryString("bookmark")
	fltVRpt = Request.QueryString("fltVRpt")
	fltERpt = Request.QueryString("fltERpt")
	fltQRpt = Request.QueryString("fltQRpt")

	
	'validate parameters that will be written into the page
	if not isnumeric(nMiMessageType) then nMiMessageType = 0
	if not isnumeric(fltSt) then fltSt = ""
	if not fnValidateSite(fltSi) then fltSi = ""
	if not isnumeric(fltVi) then fltVi = ""
	if not isnumeric(fltEf) then fltEf = ""
	if not isnumeric(fltQu) then fltQu = ""
	if not fnValidateUserName(fltUs) then fltUs = ""
	if not isnumeric(fltSj) then fltSj = ""
	if not fnValidateLabel(fltSjLb) then fltSjLb = ""
	if lcase(fltB4) <> "true" then fltB4 = "false"
	if not fnValidateDateTime(fltTm) then fltTm = ""
	if not isnumeric(fltSs) then fltSs = ""
	if not isnumeric(fltObj) then fltObj = ""
	if bNewWindow <> "1" then bNewWindow = "0"
	if not isnumeric(fltVRpt) then fltVRpt = ""
	if not isnumeric(fltERpt) then fltERpt = ""
	if not isnumeric(fltQRpt) then fltQRpt = ""

%>
<html>
<head>
<!-- #include file="HandleBrowserEvents.asp" -->
<title>
<%
select case Request.QueryString("type")
	case 0: Response.Write "Discrepancy Window"
	case 2: Response.Write "Notes Window"
	case 3: Response.Write "SDV Mark Window"
end select
%>
</title>
<%
if (not bNewWindow = "1") then
	response.Write("<script language='javascript' src='../script/RegistrationDisable.js'></script>")
	response.Write("<script language='javascript' src='../script/DLAEnable.js'></script>")
else
%>
<script language="javascript">
function fnUpdateStatusOnEform(sSite,sStudy,sSubject,sPageTaskId,sFieldID,nRepeat,nType,sStatus)
{
	var aArgs=window.dialogArguments;
	oParent=aArgs[0];
	oParent.fnUpdateStatusChange(sSite,sStudy,sSubject,sPageTaskId,sFieldID,nRepeat,nType,sStatus,1);
}
</script>
<%
end if
%>
</head>
<frameset rows="70%,35%" framespacing="0" frameborder="yes" border="1">
<frame src="MIMessageTop.asp?type=<%=nMiMessageType%>&fltSt=<%=fltSt%>&fltSi=<%=fltSi%>&fltVi=<%=fltVi%>&fltEf=<%=fltEf%>&fltQu=<%=fltQu%>&fltUs=<%=fltUs%>&fltSj=<%=fltSj%>&fltSjLb=<%=fltSjLb%>&fltB4=<%=fltB4%>&fltTm=<%=fltTm%>&fltSs=<%=fltSs%>&fltObj=<%=fltObj%>&bookmark=<%=bookmark%>&newwin=<%=bNewWindow%>&fltVRpt=<%=fltVRpt%>&fltERpt=<%=fltERpt%>&fltQRpt=<%=fltQRpt%>">
<frame src="Blank.htm">
</frameset>
</html>