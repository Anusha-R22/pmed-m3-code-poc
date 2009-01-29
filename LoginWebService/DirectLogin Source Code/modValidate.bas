Attribute VB_Name = "modValidate"
'----------------------------------------------------------------------------------------'
'   File:       modValidate.bas (routines copied from modUIHTML UNCHANGED)
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     i curtis / t aldridge 2004
'   Purpose:    functions validating MACRO fields
'----------------------------------------------------------------------------------------'
'   revisions:
'----------------------------------------------------------------------------------------'

Public Enum eInterface
    iwww = 0
    iWindows = 1
End Enum

Public Enum eWWWErrorType
    ePermission = 0
    eConfiguration = 1
    eInternal = 2
    eEForm = 3
End Enum


Public Const gsDELIMITER1 As String = "`"
Public Const gsDELIMITER2 As String = "|"
Public Const gsDELIMITER3 As String = "~"


' Alphabet and numbers - ic 18/06/2004
Private Const msALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Private Const msNUMBERS = "0123456789"
Private Const msSTANDARD_DOT = "."
Private Const msSTANDARD_COMMA = ","
Private Const msSPACE = " "
Private Const msUNDERSCORE = "_"
Private Const msFORBIDDEN_CHARS = "`|~"""

Option Explicit

'--------------------------------------------------------------------------------------------------
Public Function GetErrorHTML(ByVal eErrorType As eWWWErrorType, _
                             ByVal sErrorDescription As String, _
                    Optional ByVal lErrorCode As Long = 0, _
                    Optional ByVal sErrorSource As String = "None specified", _
                    Optional ByVal enInterface As eInterface = iwww, _
                    Optional ByVal bHideLoader As Boolean = False) As String
'--------------------------------------------------------------------------------------------------
'   ic 12/11/2002
'   function returns an html error page
'--------------------------------------------------------------------------------------------------
Dim sHTML As String

        If enInterface = iwww Then
            sHTML = sHTML & "<html>" & vbCrLf _
                          & "<head>" & vbCrLf _
                          & "<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>" & vbCrLf _
                          & "<body onload='fnPageLoaded();'>" & vbCrLf
            
            sHTML = sHTML & "<script language='javascript'>" & vbCrLf _
                          & "function fnPageLoaded(){" & vbCrLf _
                          & "window.sWinState='';}" & vbCrLf
                          
            If (bHideLoader) Then
                sHTML = sHTML & "fnHideLoader();" & vbCrLf
            End If
            
            sHTML = sHTML & "</script>" & vbCrLf
            
            sHTML = sHTML & "<table border='0' width='95%' class='clsLabelText'>" _
                            & "<tr height='15'>" _
                              & "<td></td>" _
                            & "</tr>" _
                            & "<tr>" _
                              & "<td height='15' width='15'></td><td class='clsTableHeaderText' colspan='2'></td>" _
                            & "</tr>" _
                            & "<tr height='5'><td></td></tr><tr><td></td>" & vbCrLf
               
            Select Case eErrorType
            Case eWWWErrorType.ePermission:
                sHTML = sHTML & "<td colspan='2' class='clsMessageText'>&nbsp;<img src='../img/ico_error_perm.gif'>&nbsp;Permission denied : This user does not have adequate MACRO permissions to view this page.</td>"
            Case eWWWErrorType.eConfiguration:
                sHTML = sHTML & "<td colspan='2' class='clsMessageText'>&nbsp;<img src='../img/ico_error_conf.gif'>&nbsp;A configuration error occurred : The MACRO study has not been configured to allow this operation</td>"
            Case eWWWErrorType.eInternal, eWWWErrorType.eEForm:
                sHTML = sHTML & "<td colspan='2' class='clsMessageText'>&nbsp;<img src='../img/ico_error_int.gif'>&nbsp;An internal error occurred : The request could not be completed</td>"
            End Select
            
            sHTML = sHTML & "</tr><tr height='5'><td></td></tr>"
            
            If (lErrorCode > 0) Then
                sHTML = sHTML & "<tr>" _
                                & "<td></td><td width='100' valign='top'>Code</td><td valign='top'>" & lErrorCode & "</td>" _
                              & "</tr>"
            End If
            sHTML = sHTML & "<tr>" _
                            & "<td></td><td valign='top'>Source</td><td valign='top'>" & sErrorSource & "</td>" _
                          & "</tr>" _
                          & "<tr>" _
                            & "<td></td><td valign='top'>Description</td><td valign='top'>" & sErrorDescription & "</td>" _
                          & "</tr>" _
                          & "</table>" _
                          & "</body>" _
                          & "</html>"
            
            GetErrorHTML = sHTML
        
        Else
            Err.Raise lErrorCode, sErrorSource, sErrorDescription & "|" & sErrorSource
        End If
        
End Function

'------------------------------------------------------------------------------'
Public Sub WriteErrorLog(ByVal sLocation As String, ByVal sErrorCode As String, ByVal sErrorMessage As String, _
    ByVal vParams As Variant)
'------------------------------------------------------------------------------'
' Write errors to an error log file. create file if it doesnt exist
'------------------------------------------------------------------------------'
Dim n As Integer
Dim sFileName As String

    On Error GoTo IgnoreErrors
    
    n = FreeFile
    sFileName = App.Path & "\Temp\IOErrorLog.dat"
    Open sFileName For Append As n
    Print #n, Format(Now, "dd/mm/yyyy hh:mm:ss") & " " & sLocation & gsDELIMITER1 & sErrorCode & gsDELIMITER1 _
        & sErrorMessage & "`" & Join(vParams, gsDELIMITER3)
    Close n
    
IgnoreErrors:
End Sub

'--------------------------------------------------------------------------------------------------
Public Function ReplaceControlChars(ByVal vString As Variant, ByVal sReplace As String) As String
'--------------------------------------------------------------------------------------------------
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sRtn As String
Dim nChar As Integer

    On Error GoTo CatchAllError

    For nChar = 1 To Len(vString)
        If (Asc(Mid(vString, nChar, 1)) > 32) And (Asc(Mid(vString, nChar, 1)) < 127) Then
            sRtn = sRtn & Chr(Asc(Mid(vString, nChar, 1)))
        Else
            sRtn = sRtn & sReplace
        End If
    Next
    ReplaceControlChars = sRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.ReplaceControlChars")
End Function

'--------------------------------------------------------------------------------------------------
Public Function ReplaceWithHTMLCodes(ByVal sValue As String) As String
'--------------------------------------------------------------------------------------------------
' revisions
' ic 20/06/2003 added vbLF
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
    On Error GoTo CatchAllError

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
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.ReplaceWithHTMLCodes")
End Function
'--------------------------------------------------------------------------------------------------
Public Function ReplaceWithJSChars(ByVal sStr As String) As String
'--------------------------------------------------------------------------------------------------
' ic 10/05/2001
' function accepts a string and replaces characters in the string that interrupt javascript with
' the js equivelent escape sequence
' revisions
' ic 20/06/2003 added vbLF
' ic 21/06/2004 added / and "
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sRtn As String
    
    On Error GoTo CatchAllError
    
    sRtn = Replace(sStr, "\", "\\")
    sRtn = Replace(sRtn, "/", "\/")
    sRtn = Replace(sRtn, vbCrLf, "\n")
    sRtn = Replace(sRtn, vbCr, "\n")
    sRtn = Replace(sRtn, vbLf, "\n")
    sRtn = Replace(sRtn, "'", "\'")
    sRtn = Replace(sRtn, Chr(34), "\" & Chr(34))

    ReplaceWithJSChars = sRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.ReplaceWithJSChars")
End Function

'--------------------------------------------------------------------------------------------------
Public Function ReplaceLfWithDelimiter(ByVal sStr As String, ByVal sReplace As String) As String
'--------------------------------------------------------------------------------------------------
' ic 12/02/2003
' function accepts a string and replaces linefeeds with passed replacement string
' revisions
' ic 20/06/2003 added vbLF
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sRtn As String
    
    On Error GoTo CatchAllError
    
    sRtn = Replace(sStr, vbCrLf, sReplace)
    sRtn = Replace(sRtn, vbCr, sReplace)
    sRtn = Replace(sRtn, vbLf, sReplace)

    ReplaceLfWithDelimiter = sRtn
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.ReplaceLfWithDelimiter")
End Function

''--------------------------------------------------------------------------------------------------
'Public Function ReplaceHTMLCodes(ByVal sString As String) As String
''--------------------------------------------------------------------------------------------------
''   REM 12/09/01
''
''--------------------------------------------------------------------------------------------------
'' REVISIONS
'' DPH 05/11/2001 - Convert all writable HTML char codes
'' DPH 07/01/2003 - Skip certain characters as need not replace
'' MLM 19/02/03: Changed to use HexDecodeChars, as this copes with all hexed values
'' ic 16/04/2003 changed name to ReplaceHTMLCodes()
''   ic 29/06/2004 added error handling
''--------------------------------------------------------------------------------------------------
'    On Error GoTo CatchAllError
'
'    ' Convert spaces firstly
'    sString = Replace(sString, "+", " ")
'    ReplaceHTMLCodes = HexDecodeChars(sString)
'    Exit Function
'
'CatchAllError:
'    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.ReplaceHTMLCodes")
'End Function

'--------------------------------------------------------------------------------------------------
Public Sub AddStringToVarArr(ByRef vArr() As String, ByVal sData As String)
'--------------------------------------------------------------------------------------------------
' Add string to variant array
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
    Dim lArray As Long
    Dim bAddNew As Boolean
    
    On Error GoTo CatchAllError
    
    lArray = UBound(vArr)
    bAddNew = True
    
    If lArray = 0 And (IsEmpty(vArr(0)) Or (vArr(0) = "")) Then
        vArr(0) = sData
        bAddNew = False
    End If
    
    If bAddNew Then
        ReDim Preserve vArr(lArray + 1)
        vArr(lArray + 1) = sData
    End If
    Exit Sub

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.AddStringToVarArray")
End Sub

'--------------------------------------------------------------------------------------------------
Public Function AddToArray(ByVal vArray As Variant, _
                              ByVal sCode As String, _
                              ByVal sText As String) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 30/01/2002
'   accepts a variant - empty if items have not been added yet, otherwise a 2d array
'   returns a 2d array with the passed code and text added
'--------------------------------------------------------------------------------------------------
' revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------

    On Error GoTo CatchAllError

    If IsEmpty(vArray) Then
        ReDim vArray(1, 0)
    Else
        ReDim Preserve vArray(1, UBound(vArray, 2) + 1)
    End If
    vArray(0, UBound(vArray, 2)) = sCode
    vArray(1, UBound(vArray, 2)) = sText
    
    AddToArray = vArray
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTML.AddToArray")
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateAppState(ByVal sAppState As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sIllegal As String

    sIllegal = Chr(34)
    ValidateAppState = Not ContainsIllegalChars(sAppState, sIllegal)
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateUsername(ByVal sUserName As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET & msNUMBERS & msSPACE
    ValidateUsername = (ContainslegalChars(sUserName, sLegal) And LengthIsBetween(sUserName, 0, 20))
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidatePassword(ByVal sPassword As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET & msNUMBERS & msSPACE
    ValidatePassword = (ContainslegalChars(sPassword, sLegal) And LengthIsBetween(sPassword, 0, 100))
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateDatabase(ByVal sDatabase As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET & msNUMBERS & msSTANDARD_DOT & msSPACE & msUNDERSCORE
    ValidateDatabase = (ContainslegalChars(sDatabase, sLegal) And LengthIsBetween(sDatabase, 0, 15))
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateRole(ByVal sRole As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET & msNUMBERS & msSPACE
    ValidateRole = (ContainslegalChars(sRole, sLegal) And LengthIsBetween(sRole, 0, 15))
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateSite(ByVal sSite As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    ValidateSite = (IsAlphanumeric(sSite) And LengthIsBetween(sSite, 0, 8))
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateStudyName(ByVal sStudy As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET & msNUMBERS & msUNDERSCORE
    ValidateStudyName = (ContainslegalChars(sStudy, sLegal) And LengthIsBetween(sStudy, 0, 15))
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateLabCode(ByVal sLabCode As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET & msNUMBERS & msUNDERSCORE & msSPACE
    ValidateLabCode = (ContainslegalChars(sLabCode, sLegal) And LengthIsBetween(sLabCode, 0, 15))
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateDateTime(ByVal sDateTime As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    ValidateDateTime = Not ContainsIllegalChars(sDateTime, msFORBIDDEN_CHARS)
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateLabel(ByVal sLabel As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    ValidateLabel = Not ContainsIllegalChars(sLabel, msFORBIDDEN_CHARS)
End Function

'--------------------------------------------------------------------------------------------------
Public Function ValidateText(ByVal sText As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    ValidateText = Not ContainsIllegalChars(sText, msFORBIDDEN_CHARS)
End Function

'--------------------------------------------------------------------------------------------------
Public Function IsAlphanumeric(ByVal sString As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET & msNUMBERS
    IsAlphanumeric = ContainslegalChars(sString, sLegal)
End Function

'--------------------------------------------------------------------------------------------------
Public Function IsAlphabetic(ByVal sString As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim sLegal As String

    sLegal = msALPHABET
    IsAlphabetic = ContainslegalChars(sString, sLegal)
End Function

'--------------------------------------------------------------------------------------------------
Public Function LengthIsBetween(ByVal sString As String, ByVal n1 As Integer, ByVal n2 As Integer) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim nLen As Integer

    nLen = Len(sString)
    LengthIsBetween = (nLen >= n1 And nLen <= n2)
End Function

'--------------------------------------------------------------------------------------------------
Public Function ContainsIllegalChars(ByVal sString As String, ByVal sIllegal As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim n As Integer
Dim b As Boolean
    
    b = False
    For n = 1 To Len(sIllegal)
        If InStr(sString, Mid(sIllegal, n, 1)) > 0 Then
            b = True
            Exit For
        End If
    Next
    ContainsIllegalChars = b
End Function

'--------------------------------------------------------------------------------------------------
Public Function ContainslegalChars(ByVal sString As String, ByVal sLegal As String) As Boolean
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim n As Integer
Dim b As Boolean
    
    b = True
    For n = 1 To Len(sString)
        If InStr(sLegal, Mid(sString, n, 1)) = 0 Then
            b = False
            Exit For
        End If
    Next
    ContainslegalChars = b
End Function

'--------------------------------------------------------------------------------------------------
Private Function GetErrorRedirect(Optional ByVal sMsg As String = "") As String
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    GetErrorRedirect = "<HTML>" _
        & "<body onload='window.navigate(" & Chr(34) _
        & "Error.asp?msg=" & URLEncodeString(sMsg) _
        & Chr(34) & ")'>" _
        & "</body>" _
        & "</html>"
End Function

'--------------------------------------------------------------------------------------------------
Public Function URLEncodeString(ByVal sString As String) As String
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Dim n As Integer
Dim nAsc As Integer
Dim sEncString As String


    For n = 1 To Len(sString)
        nAsc = Asc(Mid(sString, n, 1))
        
        If (nAsc < 65) Or (nAsc > 90 And nAsc < 97) Or (nAsc > 122) Then
            sEncString = sEncString & "%" & Hex(nAsc)
        Else
            sEncString = sEncString & Mid(sString, n, 1)
        End If
    Next
    URLEncodeString = sEncString
End Function
