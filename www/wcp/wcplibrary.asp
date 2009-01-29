<%
'--------------------------------------------------------------------------------------------------
function DisplayAppHeader()
'--------------------------------------------------------------------------------------------------
' ic 07/03/2007 issue 2889
'--------------------------------------------------------------------------------------------------
dim oIo
dim sVersion

    Set oIo = Server.CreateObject("MACROWWWIO30.clsWWW")
    sVersion = oIo.GetVersionInfo()
    set oIo = nothing

	Response.Write("<div class='clsTableHeaderText'>")
    Response.Write("&nbsp;User: " & session("userfullname") & "<br>")
    Response.Write("&nbsp;Version: " & sVersion & "<br><br><br>")
    Response.Write("</div>")
end function

'--------------------------------------------------------------------------------------------------
function TrueToChecked(bValue) 
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
	if bvalue then
		TrueToChecked = "checked"
	else
		TrueToChecked = ""
	end if
end function

'--------------------------------------------------------------------------------------------------
function TrueToYes(bValue) 
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
	if bvalue then
		TrueToYes = "Yes"
	else
		TrueToYes = "No"
	end if
end function

'--------------------------------------------------------------------------------------------------
function DisplayError(sError,sMessage)
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
		Response.Write "<div class='clsLabelText' style='position:absolute; left:10%;top:10%;'>"
		Response.Write sError
		if sMessage <> "" then
			Response.write ", the following error was returned:<p>"
			Response.Write sMessage
		end if
		Response.Write "<p><p><a href=wcplogin.asp>Return to login screen<a>"
		Response.Write "</div>"
end function

'--------------------------------------------------------------------------------------------------
function Alert(sError, sMessage)
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
		Response.Write "<script language='javascript'>function myAlert(){alert('"
		Response.Write sError
		if sMessage <> "" then
			Response.write ".\nThe following error was returned:\n"
			Response.Write ReplaceWithJSChars(sMessage)
		end if		
		Response.Write "');}</script>"
end function

'--------------------------------------------------------------------------------------------------
function NoAlert()
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------

		Response.Write "<script language='javascript'>function myAlert(){}</script>"
end function


'--------------------------------------------------------------------------------------------------
function ChangeOwnPassword(vSerialisedUser,sOldPassword,sNewPassword1,sNewPassword2)
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
dim sMessage ' get message back
	if sNewPassword1 <> sNewPassword2 then
		'passwords don't match - do something
		session("user")=""
		'do something with smessage
		displayerror "Password change failed", "passwords do not match"
		bPasswordChanged= false
	else
	
		set oAPI = createobject("MACROAPI30.MACROAPI")
		bPasswordChanged= oAPI.ChangeUserPasswordForASP(vSerialisedUser,sNewPassword1, sOldPassword,sMessage )
		if bPasswordChanged then
			session("user")=vSerialisedUser
			call Alert("Password changed","")
		else
			session("user")=""
			vSerialisedUser=""
			'do something with smessage
			displayerror "Password change failed", sMessage
			NoAlert
		end if
	end if
	
	ChangeOwnPassword = bPasswordChanged

end function
'--------------------------------------------------------------------------------------------------
function Login(vUserName,vPassword)
'--------------------------------------------------------------------------------------------------
'userrole and database defined in wcpconfig.asp
'--------------------------------------------------------------------------------------------------
%><!--#include file="wcpconfig.asp" --><%


	if vUsername = "" then
		set Login = nothing
		Response.Redirect "wcplogin.asp"
	else
		set oAPI = createobject("MACROAPI30.MACROAPI")
		lLoginCode = oAPI.LoginForASP(vUserName,vPassword,sDatabase,sUserRole,vMessage,vUserFullName,vSerialisedUser)
	
		select case lLoginCode
		case  0
			session("user") = vSerialisedUser
			session("userfullname") = vUserFullName
			set Login = oAPI
		case 4' change password required
			session("user") = vSerialisedUser
			session("userfullname") = vUserFullName
			Response.redirect "wcplogin.asp?username=" & vUserName
			set login = nothing
		case else		
			displayerror "Login failed", vMessage
			set Login=nothing
		end select
		
	end if

end function

'--------------------------------------------------------------------------------------------------
function GetUser(bRedirect)
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
	vSerialisedUser=session("user") 
	If vSerialisedUser="" then
		GetUSer=""
		if bRedirect then
			Response.Redirect "wcplogin.asp"
		end if
	else
		GetUser=vSerialisedUser
	end if	 
end function

'--------------------------------------------------------------------------------------------------
Function GetUsersDetails(vSerialisedUser)
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
	set oAPI = Createobject("MACROAPI30.MACROAPI")
	set colDetails = oAPI.GetUsersDetails(vSerialisedUser,"",vMessage)	
	If colDetails is nothing then
		'call alert("Retrieving user details failed", vMessage)
		call Displayerror("Retrieving user details failed", vMessage)
		set GetUsersDetails=nothing
	else
		'call noAlert()
		set GetUsersDetails = colDetails
	end if	

end function

'--------------------------------------------------------------------------------------------------
Function ShowUpdateUser(sEditedUserName,sDoUpdate)
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
	vSerialisedUser=GetUser(false)
	bDoUpdate = (SDoUpdate="true")
	set oAPI = createobject("MACROAPI30.MACROAPI")	
	set colDetails = oAPI.GetUsersDetails(vSerialisedUser,sEditedUserName,vMessage)	
	If colDetails is nothing then
		DisplayError "Retrieving individual user details failed", vMessage
	end if		
	set oDetail = CreateObject("MACROAPI30.UserDetail")
	set oDetail = colDetails(1)
	if bDoUpdate then
		oDetail.Enabled =(Request.Form("chkEnabled")="on")
		if Request.Form("password1") <> "" then
			oDetail.FailedAttempts=0
			oDetail.UnEncryptedPassword =Request.Form("password1")
		end if
		bSuccess=oAPI.ChangeUserDetails((vSerialisedUser),(oDetail),vMessage)
		if not bSuccess then
			call alert("Update failed",vmessage)
		else
			call alert("Details changed","")
		end if
	else
		call noalert()
	end if
	set showupdateuser=oDetail
end function
'--------------------------------------------------------------------------------------------------
Function ReplaceWithHTMLCodes(sValue) 
'--------------------------------------------------------------------------------------------------
' revisions
' ic 20/06/2003 added vbLF
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
    
    If Not IsNull(sValue) Then
        'first replace '&' to encode possible html codes
        sValue = Replace(sValue, "&", "&#38;")
        
        'replace html tag chars
        sValue = Replace(sValue, "<", "&#60;")
        sValue = Replace(sValue, ">", "&#62;")
        
        'replace control chars
        sValue = Replace(sValue, vbCrLf, "<br>")
        sValue = Replace(sValue, vbCr, "<br>")
        sValue = Replace(sValue, vbLf, "<br>")
    End If
    ReplaceWithHTMLCodes = sValue
    
End Function
'--------------------------------------------------------------------------------------------------
Function ReplaceWithJSChars(sStr) 
'--------------------------------------------------------------------------------------------------
' ic 10/05/2001
' function accepts a string and replaces characters in the string that interrupt javascript with
' the js equivelent escape sequence
' revisions
' ic 20/06/2003 added vbLF
' ic 21/06/2004 added / and "
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sRtn 
       
    sRtn = Replace(sStr, "\", "\\")
    sRtn = Replace(sRtn, "/", "\/")
    sRtn = Replace(sRtn, vbCrLf, "\n")
    sRtn = Replace(sRtn, vbCr, "\n")
    sRtn = Replace(sRtn, vbLf, "\n")
    sRtn = Replace(sRtn, "'", "\'")
    sRtn = Replace(sRtn, Chr(34), "\" & Chr(34))

    ReplaceWithJSChars = sRtn
    
End Function

%>