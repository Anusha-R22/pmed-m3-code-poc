VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACROLISA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------
' File: MACROLISA.cls
' Copyright: InferMed Ltd 2003 All Rights Reserved
' Author: Nicky Johns/Richard Meinesz, August 2003
' Purpose: Interface between MACRO and LISA application
'------------------------------------------------------------------
' REVISIONS
' NCJ 7 Aug 2003 - Initial development
' NCJ 27 Aug 03 - Added generic error handling
' NCJ Mar 04 - LISA Phase II development work
' NCJ 27 Apr 04 - Tidied up error handling
'------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------
Public Function GetLISASubjectData(ByVal sSerialisedUser As String, _
                                    ByVal sStudyName As String, _
                                    ByVal sSubjLabel As String, _
                                    ByRef sReturnedData As String, _
                                    ByRef sLockTokens As String) As Long
'---------------------------------------------------------------------
' Given a user and an XML data request string (specifying subject and required data),
' return the requested data as an XML string
' Function result is an error code. If code = 0, sReturnedData is XML data,
' otherwise sReturnedData is an error message
'   0 - All OK
'   1 - XMLDatarequest not understood
'   2 - Subject could not be opened
'---------------------------------------------------------------------
Dim oXMLSubject As XMLSubject

    On Error GoTo ErrLabel
    
    Set oXMLSubject = New XMLSubject
    
    GetLISASubjectData = oXMLSubject.GetLISASubjectData(sSerialisedUser, _
                                    sStudyName, sSubjLabel, sReturnedData, sLockTokens)
    Set oXMLSubject = Nothing
    
Exit Function
ErrLabel:
    ' Return the error details
    sReturnedData = GetErrorString("GetLISASubjectData", Err.Number, Err.Description)
    GetLISASubjectData = glERROR_RESULT

End Function

'---------------------------------------------------------------------
Public Function InputLISASubjectData(ByVal sSerialisedUser As String, _
                                ByVal sXMLDataInput As String, _
                                ByVal sLockTokens As String, _
                                ByRef sXMLReport As String) As Long
'---------------------------------------------------------------------
' Given a user and an XML data input string,
' add all the specified data values
' Returns error code
'   0 - All OK
'   1 - XMLDataInput not understood
'   2 - Subject does not exist
'   2 - Subject could not be opened
'   3 - Some data not added
' If not all OK, sXMLReport contains error messages
'---------------------------------------------------------------------
Dim oXMLInputData As XMLInputData
    
    On Error GoTo ErrLabel
    
    Set oXMLInputData = New XMLInputData
    
    InputLISASubjectData = oXMLInputData.InputLISASubjectData(sSerialisedUser, _
                        sXMLDataInput, sLockTokens, sXMLReport)
    
    Set oXMLInputData = Nothing

Exit Function
ErrLabel:
    ' Return the error details
    sXMLReport = GetErrorString("InputLISASubjectData", Err.Number, Err.Description)
    InputLISASubjectData = glERROR_RESULT

End Function

'---------------------------------------------------------------------
Public Function UnlockSubject(ByVal sSerialisedUser As String, _
                            ByVal sLockTokens As String, _
                            ByRef sErrMsg As String) As Long
'---------------------------------------------------------------------
' Unlock a previously locked subject
'---------------------------------------------------------------------
Dim oLISASubject As XMLSubject

    On Error GoTo ErrLabel
    
    Set oLISASubject = New XMLSubject
    UnlockSubject = oLISASubject.UnlockSubject(sSerialisedUser, sLockTokens, sErrMsg)
    Set oLISASubject = Nothing
 
Exit Function
ErrLabel:
    ' Return the error details
    sErrMsg = GetErrorString("UnlockSubject", Err.Number, Err.Description)
    UnlockSubject = glERROR_RESULT

End Function

'---------------------------------------------------------------------
Public Function Login(ByVal sUserName As String, _
                      ByVal sPassword As String, _
                      ByVal sDatabaseCode As String, _
                      ByVal sUserRole As String, _
                      ByRef sMessage As String, _
                      ByRef sUserNameFull As String, _
                      ByRef sSerialisedUser As String) As Long
'---------------------------------------------------------------------
'REM 12/08/03
'Routine for LISA app to login into MACRO database
'---------------------------------------------------------------------
Dim oLogin As LISALogin

    On Error GoTo ErrLabel
    
    Set oLogin = New LISALogin
    
    Login = oLogin.LISALogin(sUserName, sPassword, sDatabaseCode, sUserRole, sMessage, sUserNameFull, sSerialisedUser)
    
    Set oLogin = Nothing
    
Exit Function
ErrLabel:
    ' Return the error details
    sMessage = GetErrorString("Login", Err.Number, Err.Description)
    Login = glERROR_RESULT

End Function

'---------------------------------------------------------------------
Public Function ChangeUserPassword(ByRef sSerialisedUser As String, _
                                   ByVal sNewPassword As String, _
                                   ByVal sOldPassword As String, _
                                   ByRef sMessage As String) As Boolean
'---------------------------------------------------------------------
'REM 12/08/03
'Routine to change a users MACRO password
'---------------------------------------------------------------------
Dim oPassword As LISALogin

    On Error GoTo ErrLabel
    
    Set oPassword = New LISALogin
    
    ChangeUserPassword = oPassword.ChangePasswordRequest(sSerialisedUser, sNewPassword, sOldPassword, sMessage)

    Set oPassword = Nothing

Exit Function
ErrLabel:
    ' Return the error details
    sMessage = GetErrorString("ChangeUserPassword", Err.Number, Err.Description)
    ChangeUserPassword = False

End Function

'---------------------------------------------------------------------
Private Function GetErrorString(sProcName As String, lErrNo As Long, sErrDesc As String) As String
'---------------------------------------------------------------------
' NCJ 27 Aug 03 - Create a suitable error message for an unexpected MACRO error
'---------------------------------------------------------------------

    ' Let's hope we don't get any errors here...
    On Error Resume Next
    GetErrorString = "ERROR in: " & sProcName & vbCrLf _
                    & "Error Number: " & lErrNo & vbCrLf _
                    & "Error Description: " & sErrDesc

End Function
