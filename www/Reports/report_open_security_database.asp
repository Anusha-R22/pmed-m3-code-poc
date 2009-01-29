<% 

' Get connection string
	set oIo = server.CreateObject("MACROWWWIO30.clsWWW")	
  sConnect = oIO.GetSecurityCon	
	
' Establish connection to the database
	Set Connect = CreateObject("ADODB.Connection")
  connect.connectionstring = sconnect
	Connect.Open

' Close object
  set oIO = Nothing
	
%>
