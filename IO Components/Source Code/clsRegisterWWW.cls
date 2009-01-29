VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegisterWWW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsRegisterWWW.cls
'   Copyright:  InferMed Ltd. 2003 - 2006. All Rights Reserved
'   Author:     Nicky Johns, June 2003
'   Purpose:    Subject Registration class for MACRO WWW
'               (based on clsRegister in Windows Data Entry)
'
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 11 Jun 2003 - Initial development based on copy of clsRegister from Windows DM
'   TA 10/11/2005 cbd 2629 'check settings file for remote randomisation server
' NCJ 11 Jan 06 - Issue 2619 - Check eForm complete before registering
' NCJ 21 Jun 06 - Bug 2718 - Ignore errors arising from TENELEA fix cbd2629 (which breaks Batch Data Entry)
'----------------------------------------------------------------------------------------'

Option Explicit

' Columns in checks and conditions table
Private Const mnCOL_CODE = 1
Private Const mnCOL_AREZZO = 2
' Only for Uniqueness Checks
Private Const mnCOL_FLAG = 3

' Store person details
Private msTrialName As String
Private mlTrialID As Long
Private mnVersionId As Integer
Private msSite As String
Private mlPersonId As Long
Private moSubject As StudySubject
' NCJ 12 Jun 03 - We need the database code
Private msDatabaseCode As String

' Store checks  prefixes etc.
Private msCheckString As String
Private msPrefixValue As String
Private msSuffixValue As String
Private msSubjectId As String

Private mnUsePrefix As Integer
Private mnUseSuffix As Integer

' Store status and result
Private mnRegStatus As eRegStatus
Private mnRegResult As eRegResult

' Store SubjectNumbering details
Private moSubjectNumbering As clsSubjNumbersWWW

'----------------------------------------------------------------------------------------'
Public Sub Initialise(oSubject As StudySubject, _
                    ByVal sConnection As String, ByVal sDatabaseCode As String)
'----------------------------------------------------------------------------------------'
' Initialise registration class for given Trial, Site and Person
' Create SubjectNumbering and ServerConnection classes
'----------------------------------------------------------------------------------------'
                        
    On Error GoTo Errlabel
    
    Set moSubject = oSubject
    msDatabaseCode = sDatabaseCode
    
    ' Store person details from Subject object
    msTrialName = oSubject.StudyDef.Name
    mlTrialID = oSubject.StudyDef.StudyId
    mnVersionId = oSubject.StudyDef.Version
    msSite = oSubject.Site
    mlPersonId = oSubject.PersonId
    
    ' Initialise subject identifier
    msSubjectId = ""
    
    ' Find current registration status
    mnRegStatus = moSubject.RegistrationStatus
    
    ' Load up subject numbering details
    Set moSubjectNumbering = New clsSubjNumbersWWW
    Call moSubjectNumbering.Load(sConnection, mlTrialID, mnVersionId)
    ' Convert use prefix to 1/0
    If moSubjectNumbering.UsePrefix Then
        mnUsePrefix = 1
    Else
        mnUsePrefix = 0
    End If
    If moSubjectNumbering.UseSuffix Then
        mnUseSuffix = 1
    Else
        mnUseSuffix = 0
    End If

Exit Sub
Errlabel:
  Err.Raise Err.Number, , Err.Description & "|clsRegisterWWW.Initialise"

End Sub

'----------------------------------------------------------------------------------------'
Public Function RegistrationTrigger(oEFI As EFormInstance) As Boolean
'----------------------------------------------------------------------------------------'
' Returns TRUE if this is the form/visit which triggers registration
' If so, sets the subject's registration status
' from NotReady to Ready To Register
' NCJ 11 Jan 06 - Issue 2619 - Check eForm complete before registering
'----------------------------------------------------------------------------------------'
    
    RegistrationTrigger = False
    
    ' Are we on the right form and visit?
    If oEFI.eForm.EFormId <> moSubjectNumbering.TriggerFormId Then
        Exit Function
    End If
    If oEFI.VisitInstance.Visit.VisitId <> moSubjectNumbering.TriggerVisitId Then
        Exit Function
    End If
    
    ' NCJ 11 Jan 06 - Check eForm completeness
    If Not oEFI.Complete Then
        Exit Function
    End If
    
    ' We're ready to rock
    RegistrationTrigger = True
    ' Convert from "Not ready" to "Ready"
    If mnRegStatus = eRegStatus.NotReady Then
        mnRegStatus = eRegStatus.Ready
        Call SetRegistrationStatus(mnRegStatus)
    End If
    
End Function

'----------------------------------------------------------------------------------------'
Public Property Get RegistrationStatus() As Long
'----------------------------------------------------------------------------------------'
' The registration status of this subject (as eRegStatus)
' Can only be called after Initialise
'----------------------------------------------------------------------------------------'

    RegistrationStatus = mnRegStatus
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get TrialId() As Long
'----------------------------------------------------------------------------------------'
' Clinical Trial ID
'----------------------------------------------------------------------------------------'

    TrialId = mlTrialID
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get Site() As String
'----------------------------------------------------------------------------------------'
' Site
'----------------------------------------------------------------------------------------'

    Site = msSite
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get PersonId() As Long
'----------------------------------------------------------------------------------------'
' Person ID
'----------------------------------------------------------------------------------------'

    PersonId = mlPersonId
    
End Property

'----------------------------------------------------------------------------------------'
Public Function ShouldRegisterSubject() As Boolean
'----------------------------------------------------------------------------------------'
' Decide whether we should start registration
' Assume that trigger form/visit has completed
' We do registration if:
'   Study RRServerType <> None
'   UseRegistration = 1
'   Subject not already registered
'----------------------------------------------------------------------------------------'
    
    On Error GoTo Errlabel
    
    ShouldRegisterSubject = False
    
    ' Check whether subject is already registered
    ' or not ready to register
    Select Case mnRegStatus
    Case eRegStatus.Registered, eRegStatus.NotReady
        Exit Function
    End Select
    
    ' Are they using registration?
    If Not moSubjectNumbering.UseRegistration Then
        Exit Function
    End If

    ' If we got this far we'll forge ahead!
    ShouldRegisterSubject = True

Exit Function
Errlabel:
  Err.Raise Err.Number, , Err.Description & "|clsRegisterWWW.ShouldRegisterSubject"

End Function

'----------------------------------------------------------------------------------------'
Public Function IsEligible() As Boolean
'----------------------------------------------------------------------------------------'
' Returns TRUE if current subject is eligible for registration
'----------------------------------------------------------------------------------------'

    ' Check both eligibility and ineligibility
    IsEligible = CheckEligibilityConditions
    
End Function

'----------------------------------------------------------------------------------------'
Public Function DoRegistration() As Integer
'----------------------------------------------------------------------------------------'
' Actually perform the registration
' Assume everything is set up and all preliminary checks have been done
' Set the subject's registration status appropriately
' Returns result of attempting to register the patient
' NCJ 11 Jun 03 - Only do local registration for Web MACRO
'TA 10/11/2005 cbd 2629 'check settings file for remote randomisation server
' NCJ 21 Jun 06 - cbd2629 breaks Batch Data Entry so ignore errors here
'----------------------------------------------------------------------------------------'
Dim sRandUrl As String

    On Error GoTo Errlabel
    
    ' Initialise to "error" (always assume the worst...)
    mnRegResult = eRegResult.RegError
    
    ' NCJ 21 Jun 06 - Ignore errors if we can't find the settings file
    ' because this bit isn't necessary for registration in Batch Data Entry
    sRandUrl = ""
    On Error Resume Next
    'TA 10/11/2005 cbd 2629 'check settings file for remote randomisation server
    InitialiseSettingsFile True
    sRandUrl = GetMACROSetting(LCase(msTrialName) & "RandomisationUrl", "")
    On Error GoTo Errlabel
    
    If sRandUrl <> "" Then
        'use ths URL
        mnRegResult = RegisterRemotely(sRandUrl, msSubjectId)
    Else
        'Use local registration if RS type is "Local"
        ' or if it's "TrialOffice" and we're on the server
        mnRegResult = RegisterLocally(msSubjectId)
    End If

    Select Case mnRegResult
    Case eRegResult.RegOK
        mnRegStatus = eRegStatus.Registered
        ' Store identifier AND registration status
        Call SetSubjectIdentifier
    Case eRegResult.RegNotUnique
        mnRegStatus = eRegStatus.Failed
        Call SetRegistrationStatus(mnRegStatus)
    Case eRegResult.RegError, eRegResult.RegMissingInfo
        ' In case of errors, do not record "failure"
    End Select

    DoRegistration = mnRegResult

Exit Function
Errlabel:
  Err.Raise Err.Number, , Err.Description & "|clsRegisterWWW.DoRegistration"

End Function

'----------------------------------------------------------------------------------------'
Public Property Get SubjectIdentifier() As String
'----------------------------------------------------------------------------------------'
' The subject identifier returned from registration
' Empty string unless registration was successful
'----------------------------------------------------------------------------------------'
    
    SubjectIdentifier = msSubjectId

End Property

'----------------------------------------------------------------------------------------'
Private Function CheckEligibilityConditions() As Boolean
'----------------------------------------------------------------------------------------'
' Check all registration conditions for trial
'----------------------------------------------------------------------------------------'
Dim nEligibilityFlag As Integer
Dim sCond As String
Dim vResult As Variant
Dim bEligible As Boolean
Dim sResult As String
Dim colResults As Collection
Dim colArezzoExprs As Collection
Dim lRow As Long

    On Error GoTo Errlabel
    
    ' Put all the registration conditions into a collection
    ' ready to evaluate them all
    Set colArezzoExprs = New Collection
    With moSubjectNumbering.RegistrationConditions
        For lRow = 1 To .Rows
            ' Condition is field 2, flag is field 3 (yet to be implemented!)
            colArezzoExprs.Add .Field(lRow, mnCOL_AREZZO)    ' Assume non-empty condition
        Next
    End With
    
    bEligible = True
    If colArezzoExprs.Count > 0 Then
        ' Evaluate all the conditions
        Set colResults = New Collection
        Set colResults = moSubject.Arezzo.EvaluateExpressions(colArezzoExprs)
        ' See if any of them failed
        lRow = 1
        For Each vResult In colResults
            sResult = CStr(vResult)
            ' Flag is field 3
            ' For MACRO 3.0, the flag's value is always 1
            nEligibilityFlag = 1
            If (nEligibilityFlag = 1 And sResult <> "true") _
             Or (nEligibilityFlag = 0 And sResult <> "false") Then
                ' We have an untrue Eligibility condition
                ' or an unfalse Ineligibility condition
                bEligible = False
                Exit For
            End If
            lRow = lRow + 1
        Next
    End If
    
    ' Store registration status of Ineligible
    If Not bEligible Then
        mnRegStatus = eRegStatus.Ineligible
        Call SetRegistrationStatus(mnRegStatus)
    End If
    
    Set colArezzoExprs = Nothing
    Set colResults = Nothing
    
    CheckEligibilityConditions = bEligible
    
Exit Function
Errlabel:
  Err.Raise Err.Number, , Err.Description & "|clsRegisterWWW.CheckEligibilityConditions"

End Function

'----------------------------------------------------------------------------------------'
Public Function EvaluateUniquenessChecks() As Boolean
'----------------------------------------------------------------------------------------'
' Get uniqueness check values concatenated into module string msCheckString
' in form CheckCode1|CheckValue1|CheckCode2|CheckValue2| ... etc.
' Returns TRUE if all checks are OK
' or FALSE if any are missing or could not be calculated
' If FALSE, msCheckString will be empty
' If TRUE, msCheckString may be empty if there were no uniqueness checks defined
'----------------------------------------------------------------------------------------'
Dim bOK As Boolean
Dim sArezzoExpr As String
Dim sResult As String
Dim lRow As Long

Const sSeparator = "|"

    On Error GoTo Errlabel
    
    msCheckString = ""
    bOK = True
    
    ' Collect the uniqueness codes & expressions for this trial
    
    ' NB We evaluate each expression on its own because there are likely to be
    ' only a few (probably less than 5)
    ' For increased efficiency we could evaluate them as a collection if needed
    
    With moSubjectNumbering.UniquenessChecks
        For lRow = 1 To .Rows
            ' Expression is field 2
            sArezzoExpr = .Field(lRow, mnCOL_AREZZO)
            ' Evaluate the expression
            sResult = moSubject.Arezzo.EvaluateExpression(sArezzoExpr)
            ' Was it OK?
            If moSubject.Arezzo.ResultOK(sResult) And sResult > "" Then
                ' Add it to result string as Code|Result|
                msCheckString = msCheckString & .Field(lRow, mnCOL_CODE) _
                            & sSeparator & sResult & sSeparator
            Else
                ' Expression wasn't OK
                bOK = False
                Exit For
            End If
        Next
    End With
     
    ' Reset string and set status if not OK
    If Not bOK Then
        msCheckString = ""
        mnRegStatus = eRegStatus.Failed
        Call SetRegistrationStatus(mnRegStatus)
    End If
    
    EvaluateUniquenessChecks = bOK
    
Exit Function
Errlabel:
  Err.Raise Err.Number, , Err.Description & "|clsRegisterWWW.EvaluateUniquenessChecks"

End Function

'----------------------------------------------------------------------------------------'
Public Function EvaluatePrefixSuffixValues() As Boolean
'----------------------------------------------------------------------------------------'
' Get prefix and suffix values
' Returns TRUE if prefix/suffix OK (even if empty)
' or FALSE if either could not be calculated
' Sets registration status to Failed if not OK
'----------------------------------------------------------------------------------------'
Dim bOK As Boolean
Dim sArezzoExpr As String
Dim sResult As String

    On Error GoTo Errlabel
    
    msPrefixValue = ""
    msSuffixValue = ""
    bOK = True
    
    ' Get expression for Prefix
    sArezzoExpr = moSubjectNumbering.Prefix
    If sArezzoExpr > "" Then
        ' Evaluate the expression
        sResult = moSubject.Arezzo.EvaluateExpression(sArezzoExpr)
        ' Was it OK?
        If moSubject.Arezzo.ResultOK(sResult) Then
            msPrefixValue = sResult
        Else
            ' Expression wasn't OK
            bOK = False
        End If
    End If
    
    ' Don't bother with Suffix if Prefix wasn't OK
    If bOK Then
        ' Get expression for Suffix
        sArezzoExpr = moSubjectNumbering.Suffix
        If sArezzoExpr > "" Then
            ' Evaluate the expression
            sResult = moSubject.Arezzo.EvaluateExpression(sArezzoExpr)
            ' Was it OK?
            If moSubject.Arezzo.ResultOK(sResult) Then
                msSuffixValue = sResult
            Else
                ' Expression wasn't OK
                bOK = False
            End If
        End If
    End If
    
    If Not bOK Then
        mnRegStatus = eRegStatus.Failed
        Call SetRegistrationStatus(mnRegStatus)
    End If

    EvaluatePrefixSuffixValues = bOK
    
Exit Function
Errlabel:
  Err.Raise Err.Number, , Err.Description & "|clsRegisterWWW.EvaluatePrefixSuffixValues"

End Function

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'
' Tidy up after ourselves
'----------------------------------------------------------------------------------------'

    Set moSubjectNumbering = Nothing
    Set moSubject = Nothing

End Sub

'----------------------------------------------------------------------------------------'
Private Function RegisterLocally(ByRef sSubjId As String) As eRegResult
'----------------------------------------------------------------------------------------'
' Perform LOCAL registration, i.e. using local DLL
' Assume everything already set up and preliminary checks done
'----------------------------------------------------------------------------------------'
Dim oSubjNumbering As clsRSSubjectNumbering

    On Error GoTo Errlabel
    
    Set oSubjNumbering = New clsRSSubjectNumbering
    With oSubjNumbering
        Call .SetDatabase(msDatabaseCode)
        Call .Initialise(msTrialName, msSite, mlPersonId, _
                            msPrefixValue, msSuffixValue, _
                            mnUsePrefix, mnUseSuffix, _
                            moSubjectNumbering.StartNumber, _
                            moSubjectNumbering.NumberWidth)
        Call .AddUniquenessChecks(msCheckString)
        sSubjId = .SubjectIdentifier
        RegisterLocally = .ResultCode
    End With

    Set oSubjNumbering = Nothing
    
Exit Function
Errlabel:
  Err.Raise Err.Number, , Err.Description & "|clsRegisterWWW.RegisterLocally"

End Function

'----------------------------------------------------------------------------------------'
Private Function RegisterRemotely(sRemoteUrl As String, ByRef sSubjId As String) As eRegResult
'----------------------------------------------------------------------------------------'
' Perform REMOTE registration, i.e. using web page
' Assume everything already set up and preliminary checks done
'----------------------------------------------------------------------------------------'
Dim sURL As String
Dim sResultOfRegistrationCall As String
Dim oRequest As Object
Dim sError As String

    On Error GoTo Errlabel
    
    RegisterRemotely = eRegResult.RegError
    
    sURL = sRemoteUrl _
        & "?TrialName=" & msTrialName & "&Site=" & msSite & "&PersonId=" & mlPersonId _
        & "&Prefix=" & URLCharToHexEncoding(msPrefixValue) & "&Suffix=" & URLCharToHexEncoding(msSuffixValue) _
        & "&UsePrefix=" & mnUsePrefix & "&UseSuffix=" & mnUseSuffix _
        & "&StartNumber=" & moSubjectNumbering.StartNumber & "&NumberWidth=" & moSubjectNumbering.NumberWidth _
        & "&UCheckString=" & URLCharToHexEncoding(msCheckString)
                
      Set oRequest = CreateObject("InferMed.MACRO.GetHttp.Request")
      sResultOfRegistrationCall = oRequest.GetHTTPResponse(sURL, 60)

    'The Registration server returns a string in the format SUBJECTIDENTIFIER<br>RESULTCODE
    'The SubjectIdentifier is returned to the calling procedure via the ByRef variable sSubjectIdentifier
    'The result code is a general status code and is return as the result of this function
    If Left$(sResultOfRegistrationCall, 6) = "<html>" Then
        sSubjId = ""
        RegisterRemotely = eRegResult.RegError
        'raise error
        sError = "The randomisation server is not responding correctly. It returned the following error:-" _
                & vbNewLine & sResultOfRegistrationCall
        Err.Raise vbObjectError + 15, , Replace(sError, "|", "-") & "|clsRegisterWWW.RegisterRemotely"
    Else
        sSubjId = ExtractFirstItemFromList(sResultOfRegistrationCall, "<br>")
        ' What's left should be the numeric result code
        If IsNumeric(sResultOfRegistrationCall) Then
            RegisterRemotely = CInt(sResultOfRegistrationCall)
        Else
            RegisterRemotely = eRegResult.RegError
        End If
    End If

Exit Function

Errlabel:
  Err.Raise Err.Number, , Err.Description & "|clsRegisterWWW.RegisterRemotely"

End Function

'---------------------------------------------------------------------
Public Function ExtractFirstItemFromList( _
        ByRef rExtractFrom As String, _
        ByVal vSeparator As String) As String
'---------------------------------------------------------------------
'COPIED FROM basCommon
' Extract first item from string rExtractFrom up to vSeparator
' Return what's left of rExtractFrom after separator
' NCJ 24/11/00 - Changed SeparatorPosition to Long; tidied up code
'---------------------------------------------------------------------
Dim lSeparatorPosition As Long

 On Error GoTo ErrHandler
    
    lSeparatorPosition = InStr(rExtractFrom, vSeparator)
    
    If lSeparatorPosition = 0 Then

        ExtractFirstItemFromList = rExtractFrom
        rExtractFrom = ""
    Else
        '  Extract item
        ExtractFirstItemFromList = Left(rExtractFrom, lSeparatorPosition - 1)
        ' Return what remains after separator
        rExtractFrom = Mid(rExtractFrom, lSeparatorPosition + Len(vSeparator))
    End If

    Exit Function
 
ErrHandler:
  Err.Raise Err.Number, , Err.Description & "|clsRegisterWWW.ExtractFirstItemFromList"

End Function

'----------------------------------------------------------------
Private Function URLCharToHexEncoding(sBeforeEncoding As String) As String
'----------------------------------------------------------------
'COPIED HERE FROM modStringUtilities
'This function is used for the purpose of encoding specific characters
'in strings that are passed to URLs (in our case Active Server Pages)
'The characters are encoded to a %. character followed by a Hex value.
'The encoded characters are automatically un-encoded when received by
'an active server page script.
'Note that for the encoding to work correctly the % character must always
'be encoded first
'
'MLM 28/03/01: Also encode carriage returns and line feeds
'----------------------------------------------------------------
Dim sTemp As String
On Error GoTo ErrHandler
    
    sTemp = Replace(sBeforeEncoding, "%", "%25")

    sTemp = Replace(sTemp, " ", "%20")
    sTemp = Replace(sTemp, "[", "%5B")
    sTemp = Replace(sTemp, "\", "%5C")
    sTemp = Replace(sTemp, "]", "%5D")
    sTemp = Replace(sTemp, "^", "%5E")
    sTemp = Replace(sTemp, "{", "%7B")
    sTemp = Replace(sTemp, "}", "%7D")
    sTemp = Replace(sTemp, "#", "%23")
    sTemp = Replace(sTemp, "&", "%26")
    sTemp = Replace(sTemp, "+", "%2B")
    sTemp = Replace(sTemp, ",", "%2C")
    sTemp = Replace(sTemp, "/", "%2F")
    sTemp = Replace(sTemp, ":", "%3A")
    sTemp = Replace(sTemp, "<", "%3C")
    sTemp = Replace(sTemp, "=", "%3D")
    sTemp = Replace(sTemp, ">", "%3E")
    sTemp = Replace(sTemp, "?", "%3F")
    sTemp = Replace(sTemp, "@", "%40")
    
    sTemp = Replace(sTemp, vbLf, "%0A")
    sTemp = Replace(sTemp, vbCr, "%0D")
    
    URLCharToHexEncoding = sTemp

    Exit Function
 
ErrHandler:
  Err.Raise Err.Number, , Err.Description & "|clsRegisterWWW.URLCharToHexEncoding"
End Function

'----------------------------------------------------------------------------------------'
Private Sub SetSubjectIdentifier()
'----------------------------------------------------------------------------------------'
' For a successful registration,
' set the subject identifier AND registration status for this subject
' Also add value to Arezzo and save it
'----------------------------------------------------------------------------------------'

    On Error GoTo Errlabel
    
    moSubject.RegistrationStatus = eRegStatus.Registered
    moSubject.RegistrationID = msSubjectId
    Call moSubject.SaveRegistration(False)  ' Save both ID and Status

Exit Sub
Errlabel:
  Err.Raise Err.Number, , Err.Description & "|clsRegisterWWW.SetSubjectIdentifier"
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub SetRegistrationStatus(nStatus As eRegStatus)
'----------------------------------------------------------------------------------------'
' Set the current registration status of this subject
'----------------------------------------------------------------------------------------'

    On Error GoTo Errlabel
    
    moSubject.RegistrationStatus = nStatus
    Call moSubject.SaveRegistration(True)   ' Save ONLY the reg. status

Exit Sub
Errlabel:
  Err.Raise Err.Number, , Err.Description & "|clsRegisterWWW.SetRegistrationStatus"

End Sub


