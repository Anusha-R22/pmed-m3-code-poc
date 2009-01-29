<%
sub WriteUserPanel

response.write "<table  border=1><tr><td><img src=""_photo_" & sUserName & ".jpg"" /></td>"
response.write "<td><table style=""font-family:verdana;font-size:8pt"">"
response.write "<tr><td>User:</td><td>" & sUserNameFull & "</td></tr>"
response.write "<tr><td>Logged in:</td><td>" & cdate(sLoginDate) & "</td></tr>"
response.write "<tr><td>User role:</td><td>" & sUserRole & "</td></tr>"
response.write "</table>"
response.write "</td></tr></table>"

end sub

sub WritePanelStart(sTitle)

%>
        <table width="180px" border="0" cellpadding="0" cellspacing="0" >
          <tr>
            <td>
              <table width="180px" border="0" cellpadding="0" cellspacing="0" style="font-family:verdana,arial,helvetica;font-size:8pt;">
                <tr bgcolor="#CCCCCC">
                  <td>
                    <img src="curve.gif" border="0">
                  </td>
                  <td>
                    <font color="#336699">
										<b>
<% 
	 response.write sTitle
%></b>
										</font> 
                  </td>
                </tr>
              </table>
              <div style= "width:100%;border-bottom:#cccccc 1px solid;border-left:#cccccc 1px solid;border-right:#cccccc 1px solid;background-color:#F0F0F0">
									 <table cellspacing="0" cellpadding="0" border="0" width=100% style="font-family:verdana,arial,helvetica;font-size:8pt;">
<%

end sub

sub WritePanelEnd

%>
								 </table>
              </div>
            </td>
          </tr>
        </table>


<%
end sub

sub WritePanel (sTitle, sContent)

WritePanelStart sTitle
response.write sContent
WritePanelEnd 

end sub

sub WritePanelLine (sContent)

response.write "<tr><td width=20px></td><td>" & sContent & "</td></tr>"

end sub

sub WriteReportLink (sURL, sTitle)

sContent = "<a onclick=""javascript:OpenReport('" & sURL & "');"" style=""cursor:hand;color:#0000ff"" >" & sTitle & "</a>"
WritePanelLine sContent

end sub

sub WriteReportLinkNewWin (sURL, sTitle)

sContent = "<a onclick=""javascript:OpenReport('" & sURL & "',true);"" style=""cursor:hand;color:#0000ff"" >" & sTitle & "</a>"
WritePanelLine sContent

end sub

%>
