<%
dim oIoUnlock

if ((Session("ssEFIToken") <> "") or (Session("ssVIToken") <> "")) then
	set oIoUnlock = Server.CreateObject("MACROWWWIO30.clsWWW")

	if (Session("ssEFIToken") <> "") then
		call oIoUnlock.UnlockASPInstance(Session("ssUser"),Session("ssEFIToken"))
		Session("ssEFIToken") = ""
	end if
	if (Session("ssVIToken") <> "") then
		call oIoUnlock.UnlockASPInstance(Session("ssUser"),Session("ssVIToken"))
		Session("ssVIToken") = ""
	end if

	set oIoUnlock = nothing
end if
%>