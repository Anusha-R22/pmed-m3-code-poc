<%@ Language=VBScript %>
<%Option Explicit%>
<%

dim username
dim password
dim eusername
dim epassword
dim lc

username = request.form("username")
password = request.form("password")

if (username <> "") then
    'encrypt username and password
    set lc = server.CreateObject("DAL.LifeCrypt")
    
    eusername = lc.Encrypt(username)
    epassword= lc.Encrypt(password)
end if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>Login Direct Test</title>
</head>
<body>

<form name="encryptform" action="TestLoginDirect.asp" method="post">
<input type=text name=username value="<%=username %>" />
<input type=text name=password value="<%=password %>" />
<input type=button value=encrypt onclick="javascript:encryptform.submit();" />
</form>

<form name="submitform" action="LoginDirect.asp" method="post">
<input type=text name=username value="<%=eusername %>" />
<input type=text name=password value="<%=epassword %>" />
<input type=button value=submit onclick="javascript:submitform.submit();" />
</form>

</body>
</html>
