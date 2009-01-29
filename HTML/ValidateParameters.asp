<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       ValidateParameters.asp
'   Author:     I Curtis, 2004
'   Purpose:    validate querystring/form parameters before they are used
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
'	ic 19/07/2004 modified acceptable characters in study, username and password
'-----------------------------------------------------------------------------------------------'
Const sALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Const sNUMBERS = "0123456789"
Const sSTANDARD_DOT = "."
Const sSTANDARD_COMMA = ","
Const sSPACE = " "
Const sUNDERSCORE = "_"

function fnValidateSite(sSite)
dim bValidated
dim cmd
dim rs

	bValidated = false
	if (Session("ValidatedSite") = sSite) then
		'this site has already been validated
		bValidated = true
	else
		'site parameter must be alphanumeric and 1-8 chars long
		If (fnAlphaNumeric(sSite) And fnLengthBetween(sSite, 1, 8)) Then
			'validate this site
			Set oCmd = Server.CreateObject("ADODB.Command")
			Set oCmd.ActiveConnection = MACROCnn
			oCmd.CommandText = "Select Count(*) From Site Where Site = ?"

			Set oParam = Server.CreateObject("ADODB.Parameter")
			With oParam
				.Type = 200
				.Size = 8
				.Direction = 1
				.Value = sSite
			End With
			oCmd.Parameters.Append oParam
	        
			Set oRs = oCmd.Execute()
	        
			oRs.MoveFirst
			If (cint(oRs.Fields(0)) = 1) Then
				Session("ValidatedSite") = sSite
				bValidated = True
			End If
	        
			oRs.Close
			Set oRs = Nothing
			Set oParam = Nothing
			Set oCmd = Nothing
		end if
	end if
	fnValidateSite = bValidated
end function

function fnValidateStudy(sStudy)
Dim bValidated
Dim oParam
Dim oCmd
Dim oRs

	bValidated = false
	if (Session("ValidatedStudy") = sStudy) then
		'this study has already been validated
		bValidated = true
	else
		'study parameter must be alphanumeric or underscore and 1-15 chars long
		If (fnLegalChars(sStudy,sALPHABET & sNUMBERS & sUNDERSCORE) And fnLengthBetween(sStudy, 1, 15)) Then
    
			Set oCmd = Server.CreateObject("ADODB.Command")
			Set oCmd.ActiveConnection = MACROCnn
			oCmd.CommandText = "Select Count(*) From ClinicalTrial Where ClinicalTrialName = ?"

			Set oParam = Server.CreateObject("ADODB.Parameter")
			With oParam
				.Type = 200
				.Size = 15
				.Direction = 1
				.Value = sStudy
			End With
			oCmd.Parameters.Append oParam
	        
			Set oRs = oCmd.Execute()
	        
			oRs.MoveFirst
			If (cint(oRs.Fields(0)) = 1) Then
				Session("ValidatedStudy") = sStudy
				bValidated = True
			End If
	        
			oRs.Close
			Set oRs = Nothing
			Set oParam = Nothing
			Set oCmd = Nothing
		End If
	end if
	fnValidateStudy = bValidated
end function

function fnValidateUsername(sUsername)
	fnValidateUsername = (fnLegalChars(sUsername,sALPHABET & sNUMBERS & sSPACE) and fnLengthBetween(sUsername, 1, 20))
end function

function fnValidatePassword(sPassword)
	fnValidatePassword = (fnLegalChars(sPassword,sALPHABET & sNUMBERS & sSPACE) and fnLengthBetween(sPassword, 1, 100))
end function

function fnValidateFilename(sFilename)
	fnValidateFilename = not fnIllegalChars(sFilename,"\/'")
end function

function fnAlphanumeric(sString)
dim n
dim b

	b = true
	for n = 1 to len(sString)
		if (not fnNumeric(mid(sString,n,1)) and not fnAlphabetic(mid(sString,n,1))) then
			b = false
		end if
	next

	fnAlphaNumeric = b
end function

function fnNumeric(sString)
    fnNumeric = IsNumeric(sString)
end function

function fnAlphabetic(sString)
Dim n
Dim b
Dim nAsc
    
    b = True
    For n = 1 To Len(sString)
        nAsc = Asc(Mid(sString, n, 1))
        If (nAsc < 65) Or (nAsc > 90 And nAsc < 97) Or (nAsc > 122) Then
            b = False
            Exit For
        End If
    Next
    fnAlphabetic = b
end function

function fnLengthBetween(sString,n1,n2)
	nLen = Len(sString)
    fnLengthBetween = (nLen >= n1 And nLen <= n2)
end function

function fnIllegalChars(sString,sIllegalChars)
dim n
dim m
dim b
	
	b = false
	For n = 1 To Len(sString)
		For m = 1 To Len(sIllegalChars)
			if (Mid(sString, n, 1) = Mid(sIllegalChars, m, 1)) then
				b = true
				exit for
			end if
		next
		if (b = true) then exit for
	next
	fnIllegalChars = b
end function

'--------------------------------------------------------------------------------------------------
Function fnLegalChars(sString,sLegal)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim n
Dim b
    'response.Write(sString & " " & sLegal) 
    b = True
    For n = 1 To Len(sString)
        If InStr(sLegal, Mid(sString, n, 1)) = 0 Then        
            b = False
            Exit For
        End If
    Next
    fnLegalChars = b
End Function
%>
