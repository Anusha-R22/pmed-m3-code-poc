<%
'==================================================================================================
'   Copyright:  InferMed Ltd. 2004 All Rights Reserved
'   File:       Global.asp
'   Author:     i curtis
'   Purpose:    Global code
'	Version:	1.0
'==================================================================================================
'	Revisions:
'==================================================================================================

Const sALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Const sNUMBERS = "0123456789"
Const sSTANDARD_DOT = "."
Const sSTANDARD_COMMA = ","
Const sSPACE = " "
Const sUNDERSCORE = "_"
Const sFORBIDDEN_CHARS = "`|~"""
Const lVBObjectError = -2147221504

'--------------------------------------------------------------------------------------------------
function fnError(sErrNumber,sErrDescription,sLocation,aParams,bASPRedirect)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim oIoError
Dim sMsg
	
	on error resume next
	
	Set oIo = Server.CreateObject("MACROWWWIO30.clsWWW")
	
	select case cstr(sErrNumber)
	case cstr(lVBObjectError + 2)
		'illegal parameter
		if oIo.RtnLogIllegalParametersFlag() then
			call oIo.LogError(sLocation,sErrNumber,sErrDescription,aParams)
		end if
		sMsg = sErrDescription
		
	case else
		'unexpected error
		call oIo.LogError(sLocation,sErrNumber,sErrDescription,aParams)
		
	end select
	set oIo = nothing
	
	if err.number <> 0 then
		if (sMsg <> "") then sMsg = sMsg & "<br>"
		sMsg = sMsg & " Dll object error. Check server configuration"
	end if
	
	if bASPRedirect then
		Response.Redirect("Error.asp?msg=" & Server.URLEncode(sMsg))
	else
		Response.Write("<script>window.navigate('Error.asp?msg=" & Server.URLEncode(sMsg) & "')</script>")
		Response.End 
	end if
end function

'--------------------------------------------------------------------------------------------------
Function fnValidateUsername(sUsername)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal

	sLegal = sALPHABET & sNUMBERS & sSPACE
    fnValidateUsername = (fnContainsLegalChars(sUsername, sLegal) And fnLengthIsBetween(sUsername, 0, 20))
End Function

'--------------------------------------------------------------------------------------------------
Function fnValidateDateTime(sDateTime)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    fnValidateDateTime = Not (fnContainsIllegalChars(sDateTime, sFORBIDDEN_CHARS))
End Function

'--------------------------------------------------------------------------------------------------
Function fnValidateLabel(sLabel)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    fnValidateLabel = Not (fnContainsIllegalChars(sLabel, sFORBIDDEN_CHARS))
End Function

'--------------------------------------------------------------------------------------------------
Function fnValidateSite(sSite)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    fnValidateSite = (fnIsAlphanumeric(sSite) And fnLengthIsBetween(sSite, 0, 8))
End Function

'--------------------------------------------------------------------------------------------------
Function fnValidateAppState(sAppState)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sIllegal

    sIllegal = chr(34)
    fnValidateAppState = not (fnContainsIllegalChars(sAppState, sIllegal))
End Function

'--------------------------------------------------------------------------------------------------
Function fnValidateDatabase(sDatabase)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal

    sLegal = sALPHABET & sNUMBERS & sSTANDARD_DOT & sSPACE & sUNDERSCORE
    fnValidateDatabase = (fnContainslegalChars(sDatabase, sLegal) And fnLengthIsBetween(sDatabase, 0, 15))
End Function

'--------------------------------------------------------------------------------------------------
Function fnValidateRole(sRole)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal

    sLegal = sALPHABET & sNUMBERS & sSPACE
    fnValidateRole = (fnContainsLegalChars(sRole, sLegal) And fnLengthIsBetween(sRole, 0, 15))
End Function

'--------------------------------------------------------------------------------------------------
Function fnIsAlphanumeric(sString)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal

    sLegal = sALPHABET & sNUMBERS
    fnIsAlphanumeric = fnContainsLegalChars(sString, sLegal)
End Function

'--------------------------------------------------------------------------------------------------
Function fnIsAlphabetic(sString)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal

    sLegal = sALPHABET
    fnIsAlphabetic = fnContainsLegalChars(sString, sLegal)
End Function

'--------------------------------------------------------------------------------------------------
Function fnLengthIsBetween(sString,n1,n2)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim nLen

    nLen = Len(sString)
    fnLengthIsBetween = (nLen >= n1 And nLen <= n2)
End Function

'--------------------------------------------------------------------------------------------------
Function fnContainsIllegalChars(sString,sIllegal)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim n
Dim b
    
    b = False
    For n = 1 To Len(sIllegal)
        If InStr(sString, Mid(sIllegal, n, 1)) > 0 Then
            b = True
            Exit For
        End If
    Next
    fnContainsIllegalChars = b
End Function

'--------------------------------------------------------------------------------------------------
Function fnContainsLegalChars(sString,sLegal)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim n
Dim b
    
    b = True
    For n = 1 To Len(sString)
        If InStr(sLegal, Mid(sString, n, 1)) = 0 Then
            b = False
            Exit For
        End If
    Next
    fnContainsLegalChars = b
End Function

'--------------------------------------------------------------------------------------------------
function fnReplaceWithJSChars(sString)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
	dim sRtn
	
	sRtn = sString
	sRtn = Replace(sRtn, "\", "\\")
	sRtn = Replace(sRtn, "/", "\/")
	sRtn = Replace(sRtn, Chr(34), "\" & Chr(34))
	
	fnReplaceWithJSChars = sRtn
end function

'--------------------------------------------------------------------------------------------------
Function fnReplaceWithHTMLCodes(sValue)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------

	'first replace '&' to encode possible html codes
    sValue = Replace(sValue, "&", "&#38;")
        
    'replace html tag chars
    sValue = Replace(sValue, "<", "&#60;")
    sValue = Replace(sValue, ">", "&#62;")
   
    fnReplaceWithHTMLCodes = sValue
End Function
%>
