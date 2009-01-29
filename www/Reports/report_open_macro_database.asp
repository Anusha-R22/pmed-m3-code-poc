<% 
' Establish connection to the database

	Set Connect = CreateObject("ADODB.Connection")
	Connect.ConnectionString = sConnectionString
	Connect.Open

%>
