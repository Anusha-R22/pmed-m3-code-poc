VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACROAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------
' File: MACROAPI.cls
' Copyright: InferMed Ltd 2004-2008 All Rights Reserved
' Author: Nicky Johns, February 2004
' Purpose: MACRO interface, based on original interface between MACRO and LISA application
'------------------------------------------------------------------
' REVISIONS
' NCJ 2 Feb 04 - This file created from original MACROLISA.cls
' TA 24/06/2005 - added update question derivation - COMMENTED OUT
' NCJ 9 Aug 06 - CreateSubject returns a Long
' NCJ 15 Aug 06 - Added call to do Registration (Issue 2735)
' NCJ 23 Nov 07 - New LoginSecurity (and LoginSecurityForASP) for ARFG
' NCJ 23 Nov 07 - Added ImportCategories and ExportCategories (uses MACROCATBS30.dll) for ARFG
' NCJ 19 Feb 08 - Added ResetPassword for ARFG
' NCJ 6 Mar 08 - Added Export and Import Associations
' NCJ 17 Mar 08 - Removed Security DB parameter from ResetPassword
'------------------------------------------------------------------

Option Explicit

Private Const mlERROR_RESULT As Long = -1

' NCJ 23 Nov 07 - New error code for LoginSecurity
Private Const mlINVALID_SECURITYDB = 5

' The overall result for data requests
Public Enum eDataRequestResult
    Success = 0
    InvalidXML = 1
    SubjectNotExist = 2
    SubjectNotOpened = 3
End Enum

' The overall result for data input
Public Enum eDataInputResult
    Success = 0
    InvalidXML = 1
    SubjectNotExist = 2
    SubjectNotOpened = 3
    DataNotAdded = 4
End Enum

' The individual data error message types for data input
' NCJ 18 Aug 06 - Added InvalidLab
Public Enum eDataInputError
    InvalidXML = 1
    SubjectNotExist = 2
    SubjectNotOpened = 3
    VisitNotExist = 4
    EFormNotExist = 5
    QuestionNotExist = 6
    EFormInUse = 7
    VisitLockedFrozen = 8
    EFormLockedFrozen = 9
    QuestionNotEnterable = 10
    NoVisitEformDate = 11
    NoLockForSave = 12
    ValueRejected = 13
    ValueWarning = 14
    InvalidLab = 15
End Enum

' The result of API registration
Public Enum eAPIRegResult
    Success = 0
    AlreadyRegistered = 1
    NotReady = 2
    Ineligible = 3
    NotUnique = 4
    MissingInfo = 5
    UnknownError = 10
    SubjectNotOpened = 11
    SubjectReadOnly = 12
End Enum

' The result of Import Categories or Import Associations
Public Enum eImportResult
    Success = 0
    InvalidXML = 1
    NotAllDone = 2
    PermissionDenied = 3
    InvalidSecurityDB = 4
End Enum

'---------------------------------------------------------------------
Public Function GetXMLSubjectDataForASP(ByVal sSerialisedUser As String, _
                                    ByVal sXMLDataRequest As String, _
                                    ByRef vReturnedData As Variant) As eDataRequestResult
'---------------------------------------------------------------------
' Given a user and an XML data request string (specifying subject and required data),
' return the requested data as an XML string
' Function result is an error code. If code = 0, sReturnedData is XML data,
' otherwise sReturnedData is an error message
'   0 - All OK
'   1 - XMLDatarequest not understood
'   2 - Subject could not be opened
'---------------------------------------------------------------------
Dim oXMLSubject As APIGetData
Dim sReturnedData As String

    On Error GoTo ErrLabel
    
    Set oXMLSubject = New APIGetData
    
    sReturnedData = vReturnedData

    GetXMLSubjectDataForASP = oXMLSubject.GetXMLSubjectData(sSerialisedUser, _
                                    sXMLDataRequest, sReturnedData)
                                    
    vReturnedData = sReturnedData

    Set oXMLSubject = Nothing
    
Exit Function
ErrLabel:
    ' Return the error details
    vReturnedData = GetErrorString("GetXMLSubjectDataForASP", Err.Number, Err.Description)
    GetXMLSubjectDataForASP = mlERROR_RESULT

End Function

'---------------------------------------------------------------------
Public Function GetXMLSubjectData(ByVal sSerialisedUser As String, _
                                    ByVal sXMLDataRequest As String, _
                                    ByRef sReturnedData As String) As eDataRequestResult
'---------------------------------------------------------------------
' Given a user and an XML data request string (specifying subject and required data),
' return the requested data as an XML string
' Function result is an error code. If code = 0, sReturnedData is XML data,
' otherwise sReturnedData is an error message
'   0 - All OK
'   1 - XMLDatarequest not understood
'   2 - Subject could not be opened
'---------------------------------------------------------------------
Dim oXMLSubject As APIGetData

    On Error GoTo ErrLabel
    
    Set oXMLSubject = New APIGetData
    
    GetXMLSubjectData = oXMLSubject.GetXMLSubjectData(sSerialisedUser, _
                                    sXMLDataRequest, sReturnedData)
    Set oXMLSubject = Nothing
    
Exit Function
ErrLabel:
    ' Return the error details
    sReturnedData = GetErrorString("GetXMLSubjectData", Err.Number, Err.Description)
    GetXMLSubjectData = mlERROR_RESULT

End Function

'---------------------------------------------------------------------
Public Function InputXMLSubjectDataForASP(ByVal sSerialisedUser As String, _
                                ByVal sXMLDataInput As String, _
                                ByRef vXMLReport As Variant) As eDataInputResult
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
Dim oAPIInputData As APIInputData
Dim sXMLReport As String

    On Error GoTo ErrLabel
    
    Set oAPIInputData = New APIInputData
    
    sXMLReport = vXMLReport
    
    InputXMLSubjectDataForASP = oAPIInputData.InputXMLSubjectData(sSerialisedUser, sXMLDataInput, sXMLReport)
    
    vXMLReport = sXMLReport
    
    Set oAPIInputData = Nothing

Exit Function
ErrLabel:
    ' Return the error details
    vXMLReport = GetErrorString("InputXMLSubjectDataForASP", Err.Number, Err.Description)
    InputXMLSubjectDataForASP = mlERROR_RESULT

End Function

'---------------------------------------------------------------------
Public Function InputXMLSubjectData(ByVal sSerialisedUser As String, _
                                ByVal sXMLDataInput As String, _
                                ByRef sXMLReport As String) As eDataInputResult
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
Dim oAPIInputData As APIInputData
    
    On Error GoTo ErrLabel
    
    Set oAPIInputData = New APIInputData
    
    InputXMLSubjectData = oAPIInputData.InputXMLSubjectData(sSerialisedUser, sXMLDataInput, sXMLReport)
    
    Set oAPIInputData = Nothing

Exit Function
ErrLabel:
    ' Return the error details
    sXMLReport = GetErrorString("InputXMLSubjectData", Err.Number, Err.Description)
    InputXMLSubjectData = mlERROR_RESULT

End Function

'---------------------------------------------------------------------
Public Function LoginForASP(ByVal sUserName As String, _
                      ByVal sPassword As String, _
                      ByVal sDatabaseCode As String, _
                      ByVal sUserRole As String, _
                      ByRef vMessage As Variant, _
                      ByRef vUserNameFull As Variant, _
                      ByRef vSerialisedUser As Variant) As Long
'---------------------------------------------------------------------
'REM 12/08/03
'Routine for LISA app to login into MACRO database
'---------------------------------------------------------------------
Dim oLogin As APILogin
Dim sMessage As String
Dim sUserNameFull As String
Dim sSerialisedUser As String

    On Error GoTo ErrLabel
    
    sMessage = vMessage
    sUserNameFull = vUserNameFull
    sSerialisedUser = vSerialisedUser
    
    LoginForASP = Login(sUserName, sPassword, sDatabaseCode, sUserRole, sMessage, sUserNameFull, sSerialisedUser)
    
    vMessage = sMessage
    vUserNameFull = sUserNameFull
    vSerialisedUser = sSerialisedUser
        
    
Exit Function
ErrLabel:
    ' Return the error details
    vMessage = GetErrorString("LoginForASP", Err.Number, Err.Description)
    LoginForASP = mlERROR_RESULT

End Function

'---------------------------------------------------------------------
Public Function LoginSecurityForASP(ByVal sUserName As String, _
                      ByVal sPassword As String, _
                      ByVal sDatabaseCode As String, _
                      ByVal sUserRole As String, _
                      ByVal sSecurityCon As String, _
                      ByRef vMessage As Variant, _
                      ByRef vUserNameFull As Variant, _
                      ByRef vSerialisedUser As Variant) As Long
'---------------------------------------------------------------------
' NCJ 23 Nov 07 - Routine to login into MACRO database with Security DB details
' (ASP version)
'---------------------------------------------------------------------
Dim oLogin As APILogin
Dim sMessage As String
Dim sUserNameFull As String
Dim sSerialisedUser As String

    On Error GoTo ErrLabel
    
    sMessage = vMessage
    sUserNameFull = vUserNameFull
    sSerialisedUser = vSerialisedUser
    
    LoginSecurityForASP = LoginSecurity(sUserName, sPassword, sDatabaseCode, sUserRole, sSecurityCon, _
                        sMessage, sUserNameFull, sSerialisedUser)
    
    vMessage = sMessage
    vUserNameFull = sUserNameFull
    vSerialisedUser = sSerialisedUser
        
    
Exit Function
ErrLabel:
    ' Return the error details
    vMessage = GetErrorString("LoginSecurityForASP", Err.Number, Err.Description)
    LoginSecurityForASP = mlERROR_RESULT

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
'Routine to login into MACRO database
' NCJ 23 Nov 07 - Added bGoodSecurity and sSecCon
'---------------------------------------------------------------------
Dim oLogin As APILogin
Dim bGoodSecurity As Boolean

    On Error GoTo ErrLabel
    
    Set oLogin = New APILogin
    
    ' NCJ 23 Nov 07 - Added bGoodSecurity and empty sSecurityCon
    Login = oLogin.MACROLogin(sUserName, sPassword, sDatabaseCode, sUserRole, _
                    sMessage, sUserNameFull, sSerialisedUser, bGoodSecurity, "")
    
    Set oLogin = Nothing
    
Exit Function
ErrLabel:
    ' Return the error details
    sMessage = GetErrorString("Login", Err.Number, Err.Description)
    Login = mlERROR_RESULT

End Function

'---------------------------------------------------------------------
Public Function LoginSecurity(ByVal sUserName As String, _
                      ByVal sPassword As String, _
                      ByVal sDatabaseCode As String, _
                      ByVal sUserRole As String, _
                      ByVal sSecurityCon As String, _
                      ByRef sMessage As String, _
                      ByRef sUserNameFull As String, _
                      ByRef sSerialisedUser As String) As Long
'---------------------------------------------------------------------
' NCJ 23 Nov 07 - Routine to login into MACRO database with Security DB details
'---------------------------------------------------------------------
Dim oLogin As APILogin
Dim bGoodSecurity As Boolean
Dim lLoginResult As Long

    On Error GoTo ErrLabel
    
    Set oLogin = New APILogin
    
    lLoginResult = oLogin.MACROLogin(sUserName, sPassword, sDatabaseCode, sUserRole, _
                    sMessage, sUserNameFull, sSerialisedUser, bGoodSecurity, sSecurityCon)
    
    Set oLogin = Nothing
    
    ' Check for invalid Security DB
    If Not bGoodSecurity Then
        ' Convert result to "Invalid security DB"
        lLoginResult = mlINVALID_SECURITYDB
    End If
    
    LoginSecurity = lLoginResult
    
Exit Function
ErrLabel:
    ' Return the error details
    sMessage = GetErrorString("LoginSecurity", Err.Number, Err.Description)
    LoginSecurity = mlERROR_RESULT

End Function

Public Function ChangeUserDetails(ByVal sSerialisedUser As String, oNewDetails As UserDetail, ByRef sMessage As Variant) As Boolean
'---------------------------------------------------------------------
'TA 28/09/2004
'Routine to change users opassword, enabled status and failed attempts
'---------------------------------------------------------------------
Dim oChange As ChangeUserDetails

On Error GoTo ErrLabel

    Set oChange = New ChangeUserDetails
    ChangeUserDetails = oChange.UpdateUser(sSerialisedUser, oNewDetails, sMessage)
    
Exit Function
ErrLabel:
    ' Return the error details
    sMessage = GetErrorString("ChangeUserDetails", Err.Number, Err.Description)
    ChangeUserDetails = False


End Function


Public Function GetUsersDetails(ByVal sSerialisedUser As String, ByVal sUserName As String, ByRef sMessage As Variant) As Collection
'---------------------------------------------------------------------
'TA 28/09/2004
'Routine to retrieve users' username, usernamefull, failed attempts, enabled status and sysadmin status
' if sUserName is blank then all are returned
'---------------------------------------------------------------------
Dim oChange As ChangeUserDetails

On Error GoTo ErrLabel

    Set oChange = New ChangeUserDetails
    Set GetUsersDetails = oChange.GetUsersDetails(sSerialisedUser, sUserName, sMessage)
    
Exit Function
ErrLabel:
    ' Return the error details
    sMessage = GetErrorString("GetUsersDetails", Err.Number, Err.Description)
    Set GetUsersDetails = Nothing


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
Dim oPassword As APILogin

    On Error GoTo ErrLabel
    
    Set oPassword = New APILogin
    
    ChangeUserPassword = oPassword.ChangePasswordRequest(sSerialisedUser, sNewPassword, sOldPassword, sMessage)

    Set oPassword = Nothing

Exit Function
ErrLabel:
    ' Return the error details
    sMessage = GetErrorString("ChangeUserPassword", Err.Number, Err.Description)
    ChangeUserPassword = False

End Function

'---------------------------------------------------------------------
Public Function ChangeUserPasswordForASP(ByRef vSerialisedUser As Variant, _
                                   ByVal sNewPassword As String, _
                                   ByVal sOldPassword As String, _
                                   ByRef vMessage As Variant) As Long
'---------------------------------------------------------------------
'REM 12/08/03
' REF parameters are variants to work with ASP
'---------------------------------------------------------------------
Dim oLogin As APILogin
Dim sMessage As String
Dim sUserNameFull As String
Dim sSerialisedUser As String

    sMessage = vMessage
    sSerialisedUser = vSerialisedUser
    
    ChangeUserPasswordForASP = ChangeUserPassword(sSerialisedUser, sNewPassword, sOldPassword, sMessage)
    
    vMessage = sMessage
    vSerialisedUser = sSerialisedUser

End Function

'---------------------------------------------------------------------
Public Function ResetPassword(ByRef sSerialisedUser As String, _
                                    ByVal sUserName As String, _
                                    ByVal sNewPwd As String, _
                                    ByRef sMessage As String) As Long
'---------------------------------------------------------------------
' NCJ 19 Feb 08 - Reset another user's password
' NB sSecurity can be "" to use default Security DB
' NCJ 17 Mar 08 - No longer take Security DB parameter
'---------------------------------------------------------------------
Dim oLogin As APILogin

    On Error GoTo ErrLabel

    Set oLogin = New APILogin
    ResetPassword = oLogin.ResetAnotherPassword(sSerialisedUser, sUserName, sNewPwd, sMessage)
    Set oLogin = Nothing
        
    
Exit Function
ErrLabel:
    ResetPassword = -1
    
End Function

'---------------------------------------------------------------------
Public Function ResetPasswordForASP(ByRef vSerialisedUser As Variant, _
                                    ByVal sUserName As String, _
                                    ByVal sNewPwd As String, _
                                    ByRef vMessage As Variant) As Long
'---------------------------------------------------------------------
' NCJ 19 Feb 08 - Reset another user's password, ASP version
'---------------------------------------------------------------------
Dim oLogin As APILogin
Dim sMessage As String
Dim sSerialisedUser As String

    On Error GoTo ErrLabel

    sSerialisedUser = vSerialisedUser
    sMessage = vMessage
    
    ResetPasswordForASP = ResetPassword(sSerialisedUser, sUserName, sNewPwd, sMessage)
        
    vSerialisedUser = sSerialisedUser
    vMessage = sMessage
    
Exit Function
ErrLabel:
    ResetPasswordForASP = -1
    
End Function

'---------------------------------------------------------------------
Public Function CreateSubject(ByRef sSerialisedUser As String, _
                                    ByVal lStudyId As Long, _
                                    ByVal sSite As String, _
                                    ByRef sMessage As String) As Long
'---------------------------------------------------------------------
' NCJ 21 Jan 04 - Create a new subject
' Returns subject ID, or -1 if subject not created
' NCJ 9 Aug 06 - Changed return type to Long
'---------------------------------------------------------------------
                                    
    On Error GoTo ErrLabel
    
    CreateSubject = modAPIUtils.CreateSubject(sSerialisedUser, lStudyId, sSite, sMessage)
                                    
Exit Function
ErrLabel:
    ' Return the error details
    sMessage = GetErrorString("CreateSubject", Err.Number, Err.Description)
    CreateSubject = -1
End Function

'---------------------------------------------------------------------
Public Function CreateSubjectForASP(ByRef vSerialisedUser As Variant, _
                                    ByVal lStudyId As Long, _
                                    ByVal sSite As String, _
                                    ByRef vMessage As Variant) As Long
'---------------------------------------------------------------------
' NCJ 21 Jan 04 - Create a new subject
' Returns subject ID, or -1 if subject not created
'---------------------------------------------------------------------
Dim sMessage As String
Dim sSerialisedUser As String
                                    
    On Error GoTo ErrLabel
    
    sSerialisedUser = vSerialisedUser
    sMessage = vMessage
    
    CreateSubjectForASP = modAPIUtils.CreateSubject(sSerialisedUser, lStudyId, sSite, sMessage)
                                    
    vSerialisedUser = sSerialisedUser
    vMessage = sMessage
                                   
Exit Function
ErrLabel:
    ' Return the error details
    vMessage = GetErrorString("CreateSubjectForASP", Err.Number, Err.Description)
    CreateSubjectForASP = -1
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

'---------------------------------------------------------------------
Public Function RegisterSubject(ByVal sSerialisedUser As String, _
                                    ByVal sStudyName As String, _
                                    ByVal sSite As String, _
                                    ByVal sSubject As String, _
                                    ByRef sRegID As String) As Long
'---------------------------------------------------------------------
' NCJ 15 August 2006
' Register a subject and return the subject identifier
' sSubject may be either subj label or subjID
' sRegID is the subject identifier (empty if not registered),
' or an error message if subject not opened
'---------------------------------------------------------------------
Dim oUser As MACROUser
Dim oSubject As StudySubject
Dim oAPIRegister As APIRegistration
Dim sErrMsg As String
Dim nRegResult As eAPIRegResult

    sRegID = ""
    
    ' Create the MACRO User
    Set oUser = New MACROUser
    Call oUser.SetStateHex(sSerialisedUser)

    Set oSubject = LoadSubject(oUser, sStudyName, sSite, sSubject, sErrMsg)
    If Not oSubject Is Nothing Then
        ' Check we can update
        If oSubject.ReadOnly Then
            RegisterSubject = eAPIRegResult.SubjectReadOnly
        ElseIf oSubject.RegistrationStatus = eRegStatus.Registered Then
            ' Already registered
            RegisterSubject = eAPIRegResult.AlreadyRegistered
            sRegID = oSubject.RegistrationID
        Else
            ' Attempt registration
            Set oAPIRegister = New APIRegistration
            nRegResult = oAPIRegister.DoRegistration(oSubject, oUser.CurrentDBConString, oUser.DatabaseCode)
            If nRegResult = eAPIRegResult.Success Then
                sRegID = oSubject.RegistrationID
            End If
            RegisterSubject = nRegResult
            Set oAPIRegister = Nothing
        End If
        
        Call oSubject.StudyDef.RemoveSubject
        ' Clear up the study def
        Call oSubject.StudyDef.Terminate
        Set oSubject = Nothing
        
    Else
        ' Error in loading the subject
        RegisterSubject = eAPIRegResult.SubjectNotOpened
        sRegID = sErrMsg
    End If

    Set oUser = Nothing

End Function

'---------------------------------------------------------------------
Public Function RegisterSubjectForASP(ByVal sSerialisedUser As Variant, _
                                    ByVal sStudyName As String, _
                                    ByVal sSite As String, _
                                    ByVal sSubject As String, _
                                    ByRef vRegID As Variant) As Long
'---------------------------------------------------------------------
' Registration for ASP
'---------------------------------------------------------------------
Dim sRegID As String

    sRegID = vRegID
    RegisterSubjectForASP = RegisterSubject(sSerialisedUser, sStudyName, sSite, sSubject, sRegID)
    vRegID = sRegID

End Function

'------------------------------------------------------------------
Public Function ImportCategories(ByVal sSerialisedUser As String, _
                                ByVal sXMLCatsInput As String, _
                                ByRef sXMLReport As String) As Long
'------------------------------------------------------------------
' Import categories contained in sXMLCatsInput specification
' We use the TopCats class in the C# "MACROCATBS30" DLL
'------------------------------------------------------------------
Dim oAPICats As APICategories

    On Error GoTo ImportErr
    
    Set oAPICats = New APICategories
    
    ImportCategories = oAPICats.ImportCategories(sSerialisedUser, sXMLCatsInput, sXMLReport)

    Set oAPICats = Nothing
    
Exit Function

ImportErr:
    ' Return error details
    sXMLReport = GetErrorString("MACROAPI.ImportCategories", Err.Number, Err.Description)
    Set oAPICats = Nothing
    ImportCategories = mlERROR_RESULT
                
End Function

'------------------------------------------------------------------
Public Function ExportCategories(ByVal sSerialisedUser As String, _
                                ByVal sXMLCatRequest As String, _
                                ByRef sXMLReport As String) As Boolean
'------------------------------------------------------------------
' Import categories contained in sXMLCatRequest specification
' We use the TopCats class in the C# "MACROCATBS30" DLL
'------------------------------------------------------------------
Dim oAPICats As APICategories

    On Error GoTo ExportErr
    
    Set oAPICats = New APICategories
    
    ExportCategories = oAPICats.ExportCategories(sSerialisedUser, sXMLCatRequest, sXMLReport)

    Set oAPICats = Nothing
    
Exit Function

ExportErr:
    ' Return error details
    sXMLReport = GetErrorString("MACROAPI.ExportCategories", Err.Number, Err.Description)
    Set oAPICats = Nothing
    ExportCategories = False
                
End Function

'------------------------------------------------------------------
Public Function ImportCategoriesForASP(ByVal sSerialisedUser As String, _
                                ByVal sXMLCatsInput As String, _
                                ByRef vXMLReport As String) As Long
'------------------------------------------------------------------
' Import categories contained in sXMLCatsInput specification
' (ASP version - use variants for ByRef parameters)
'------------------------------------------------------------------
Dim oAPICats As APICategories
Dim sXMLReport As String

    On Error GoTo ImportErr
    
    Set oAPICats = New APICategories
    
    sXMLReport = vXMLReport
    ImportCategoriesForASP = oAPICats.ImportCategories(sSerialisedUser, sXMLCatsInput, sXMLReport)
    vXMLReport = sXMLReport
    
    Set oAPICats = Nothing
    
Exit Function

ImportErr:
    ' Return error details
    vXMLReport = GetErrorString("MACROAPI.ImportCategories", Err.Number, Err.Description)
    Set oAPICats = Nothing
    ImportCategoriesForASP = mlERROR_RESULT
                
End Function

'------------------------------------------------------------------
Public Function ExportCategoriesForASP(ByVal sSerialisedUser As String, _
                                ByVal sXMLCatRequest As String, _
                                ByRef vXMLReport As String) As Boolean
'------------------------------------------------------------------
' Import categories contained in sXMLCatRequest specification
' (ASP version - use variants for ByRef parameters)
'------------------------------------------------------------------
Dim oAPICats As APICategories
Dim sXMLReport As String

    On Error GoTo ExportErr
    
    Set oAPICats = New APICategories
    
    sXMLReport = vXMLReport
    ExportCategoriesForASP = oAPICats.ExportCategories(sSerialisedUser, sXMLCatRequest, sXMLReport)
    vXMLReport = sXMLReport

    Set oAPICats = Nothing
    
Exit Function

ExportErr:
    ' Return error details
    vXMLReport = GetErrorString("MACROAPI.ExportCategories", Err.Number, Err.Description)
    Set oAPICats = Nothing
    ExportCategoriesForASP = False
                
End Function

'------------------------------------------------------------------
Public Function ExportAssociations(ByVal sSerialisedUser As String, _
                                ByVal sXMLAssocRequest As String, _
                                ByRef sXMLAssociations As String) As Boolean
'------------------------------------------------------------------
' Export user role associations contained in sXMLAssocRequest specification
' We use the SSURIO class in the C# "MACROSSURBS30" DLL
'------------------------------------------------------------------
Dim oAPIAssoc As APIAssociations

    On Error GoTo ExportErr
    
    Set oAPIAssoc = New APIAssociations
    
    ExportAssociations = oAPIAssoc.ExportAssociations(sSerialisedUser, sXMLAssocRequest, _
                                        sXMLAssociations)

    Set oAPIAssoc = Nothing
    
Exit Function

ExportErr:
    ' Return error details
    sXMLAssociations = GetErrorString("MACROAPI.ExportAssociations", Err.Number, Err.Description)
    Set oAPIAssoc = Nothing
    ExportAssociations = False
                
End Function

'------------------------------------------------------------------
Public Function ImportAssociations(ByVal sSerialisedUser As String, _
                              ByVal sXMLAssociations As String, _
                              ByRef sMessage As String) As Long
'------------------------------------------------------------------
' Import user role associations contained in sXMLAssociations specification
' We use the SSURIO class in the C# "MACROSSURBS30" DLL
'------------------------------------------------------------------
Dim oAPIAssoc As APIAssociations

    On Error GoTo ImportErr
    
    Set oAPIAssoc = New APIAssociations
    ImportAssociations = oAPIAssoc.ImportAssociations(sSerialisedUser, sXMLAssociations, sMessage)
    Set oAPIAssoc = Nothing
    
Exit Function

ImportErr:
    ' Return error details
    sMessage = GetErrorString("MACROAPI.ImportAssociations", Err.Number, Err.Description)
    Set oAPIAssoc = Nothing
    ImportAssociations = mlERROR_RESULT
                
End Function

'------------------------------------------------------------------
Public Function ImportAssociationsForASP(ByRef vSerialisedUser As Variant, _
                              ByVal sXMLAssociations As String, _
                              ByRef vMessage As Variant) As Long
'------------------------------------------------------------------
' Import user role associations contained in sXMLAssociations specification (for ASP)
' We use the SSURIO class in the C# "MACROSSURBS30" DLL
'------------------------------------------------------------------
Dim oAPIAssoc As APIAssociations
Dim sSerialisedUser As String
Dim sMessage As String

    On Error GoTo ImportErr
    
    sSerialisedUser = vSerialisedUser
    sMessage = vMessage
    
    Set oAPIAssoc = New APIAssociations
    ImportAssociationsForASP = oAPIAssoc.ImportAssociations(sSerialisedUser, sXMLAssociations, sMessage)
    Set oAPIAssoc = Nothing
    
    vMessage = sMessage
    vSerialisedUser = sSerialisedUser
    
Exit Function

ImportErr:
    ' Return error details
    sMessage = GetErrorString("MACROAPI.ImportAssociationsForASP", Err.Number, Err.Description)
    Set oAPIAssoc = Nothing
    ImportAssociationsForASP = mlERROR_RESULT
                
End Function

'------------------------------------------------------------------
Public Function ExportAssociationsForASP(ByRef vSerialisedUser As Variant, _
                                ByVal sXMLAssocRequest As String, _
                                ByRef vXMLAssociations As Variant) As Boolean
'------------------------------------------------------------------
' Export user role associations contained in sXMLAssocRequest specification
' We use the SSURIO class in the C# "MACROSSURBS30" DLL
'------------------------------------------------------------------
Dim oAPIAssoc As APIAssociations
Dim sSerialisedUser As String
Dim sXMLAssociations As String

    On Error GoTo ExportErr
    
    sSerialisedUser = vSerialisedUser
    sXMLAssociations = vXMLAssociations
    
    Set oAPIAssoc = New APIAssociations
    
    ExportAssociationsForASP = oAPIAssoc.ExportAssociations(sSerialisedUser, sXMLAssocRequest, _
                                        sXMLAssociations)
    Set oAPIAssoc = Nothing
    
    vSerialisedUser = sSerialisedUser
    vXMLAssociations = sXMLAssociations
    
Exit Function

ExportErr:
    ' Return error details
    sXMLAssociations = GetErrorString("MACROAPI.ExportAssociationsForASP", Err.Number, Err.Description)
    Set oAPIAssoc = Nothing
    ExportAssociationsForASP = False
                
End Function


