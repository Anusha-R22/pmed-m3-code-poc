<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'==================================================================================================
' 	Copyright:	InferMed Ltd. 1998. All Rights Reserved
'	File:		ServerProperties.asp
'	Author: 	I Curtis
'	Purpose: 	displays server properties report
'	Version:	1.0
'==================================================================================================
'	Revisions:
'	ic 29/07/2002	changed dll reference for 3.0
'	ic 04/06/2003	renamed to ServerProperties.asp
'	ic 25/11/2003	added user column
'	ic 01/07/2004	added error handling
'==================================================================================================
%>
<!-- #include file="Global.asp" -->
<%
dim oIo
dim aData

	on error resume next
	
	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
	
	if (Request.Form("resetcache") = "1") then
		oIo.ShutDownCacheManager
	end if
	
	aData = oIo.GetCacheReport	
	set oIo = nothing
	
	if err.number <> 0 then 
		call fnError(err.number,err.description,"ServerProperties.asp oIo.GetCacheReport()",Array(),true)
	end if
	
	function fnUpTime(sStart)
		dim nMins
		dim nHours
		dim nDays
		nMins = datediff("n",sStart,now)
		nDays = cint((nMins-720)/1440)
		nHours = cint((nMins-30)/60)
		if (nHours > 23) then
			nHours = nHours mod 24
		end if
		if (nMins > 59) then
			nMins = nMins mod 60
		end if
		
		fnUpTime = nDays & "d:" & nHours & "h:" & nMins & "m"
	end function
	
	function fnCacheReport(aCache)
		dim sHtm
		dim nMax
		dim s
		dim n
		
		sHtm = "<table width='95%' border='0' cellpadding='1' cellspacing='1'>"
		sHtm = sHtm & "<tr height='20' class='clsTableHeaderText'>" _
			& "<td>&nbsp;</td>" _
			& "<td>Database Id</td>" _
			& "<td>Study</td>" _
			& "<td>Site</td>" _
			& "<td>Subject</td>" _
			& "<td>Timestamp</td>" _
			& "<td>Cache Key</td>" _
			& "<td>Cache Token</td>" _
			& "<td>User</td>" _
			& "</tr>"
		
		'get max arezzo value
		nMax = split(aCache(0),"=")(1)	
	
		for n = 1 to nMax
			sHtm = sHtm & "<tr height='10' class='clsTableText'"
			if (n < ubound(aCache)) then
				s = split(aCache(n),"|")
				if (trim(s(1)) <> "Status = Not Busy") then
					sHtm = sHtm & " bgcolor='#ffcccc'>"
				else
					sHtm = sHtm & " bgcolor='#ccffcc'>"
				end if
				sHtm = sHtm & "<td>" & n & "</td>" _
					& "<td>" & trim(split(s(9),"=")(1)) & "</td>" _
					& "<td>" & trim(split(s(2),"=")(1)) & "</td>" _
					& "<td>" & trim(split(s(3),"=")(1)) & "</td>" _
					& "<td>" & trim(split(s(4),"=")(1)) & "</td>" _
					& "<td>" & trim(split(s(5),"=")(1)) & "</td>" _
					& "<td>" & trim(split(s(6),"=")(1)) & "</td>" _
					& "<td>" & trim(split(s(7),"=")(1)) & "</td>" _
					& "<td>" & trim(split(s(8),"=")(1)) & "</td>"
			else
				sHtm = sHtm & "bgcolor='lightgrey'><td>" & n & "</td>" _
					& "<td>&nbsp;</td>" _
					& "<td>&nbsp;</td>" _
					& "<td>&nbsp;</td>" _
					& "<td>&nbsp;</td>" _
					& "<td>&nbsp;</td>" _
					& "<td>&nbsp;</td>" _
					& "<td>&nbsp;</td>" _
					& "<td>&nbsp;</td>"
			end if
			sHtm = sHtm & "</tr>"
		next
		
		sHtm = sHtm & "<tr height='5'><td></td></tr>" _
			& "<tr><td colspan='9' align='right'>" _
			& "<input name='refreshbutton' type='button' value='Refresh Page' onclick='javascript:fnRefresh();'>&nbsp;" _
			& "<input name='resetbutton' type='button' value='Reset Cache' onclick='javascript:fnResetCache();'>" _
			& "</td></tr>"
		
		sHtm = sHtm & "</table>"
		fnCacheReport = sHtm
	end function
%>
<html>
<head>
<title><%=Application("asName")%></title>
<link rel="stylesheet" href="../style/MACRO1.css" type="text/css">
<script language='javascript'>
	function fnRefresh()
	{
		document.resetcache.resetbutton.disabled = true;
		document.resetcache.refreshbutton.disabled = true;
		document.resetcache.submit();
	}
	
	function fnResetCache()
	{
		if(confirm("Are you sure you want to reset the cache?"))
		{
			document.resetcache.resetbutton.disabled = true;
			document.resetcache.refreshbutton.disabled = true;
			document.resetcache.resetcache.value = '1';
			document.resetcache.submit();
		}
	}
</script>
</head>
<body class='clsLabelText'>
<form name='resetcache' action='ServerProperties.asp' method='post'>
<input type='hidden' name='resetcache' value=''>
<%	
	Response.Write("<br><b>Current Server Time</b><br>")
	Response.Write(FormatDateTime(now,1) & " " & FormatDateTime(now,3) & "<br>")
	
	Response.Write("<br><b>Server Started</b><br>")
	Response.Write(FormatDateTime(Application("asSTARTTIME"),1) & " " & FormatDateTime(Application("asSTARTTIME"),3) & " (time up " & fnUpTime(Application("asSTARTTIME")) & ")<br>")

	Response.Write("<br><b>Cache Report</b><br>")
	Response.Write(fnCacheReport(aData))
	
	Response.Write("<br><b>Application Variables</b><br>")
	
	Response.Write("USESSL=" & Application("asUSESSL") & "<br>" _
				 & "HELPURL=" & Application("asHELPURL") & "<br>" _
				 & "DLLINFO=" & Application("asDLLINFO") & "<br>" _
				 & "USESCI=" & Application("abUSESCI") & "<br>" _
				 & "WEBVERS=" & Application("anWEBVERS") & "<br>")
	
	Response.Write("<br><b>Session Variables</b><br>")
	
	Response.Write("SESSIONID=" & (Session.SessionID) & "<br>")
	Response.Write("SERIALISEDUSER=" & (Session("ssUser") <> "") & "<br>")
	Response.Write("BROWSERACCEPT=" & session("sbBrowserAccept") & "<br>")
	Response.Write("TIMEZONEOFFSET=" & session("TimezoneOffset") & "<br>")
	Response.Write("DECIMALPOINT=" & session("ssDecimalPoint") & "<br>")
	Response.Write("THOUSANDSEPARATOR=" & session("ssThousandSeparator") & "<br>")
	
%>
</form>
</body>
</html>