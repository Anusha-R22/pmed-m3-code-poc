VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsRegister.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Nicky Johns, November 2000
'   Purpose:    Subject Registration class for MACRO
'               for use in MACRO Data Entry
'
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 21-24 Nov 2000 - Initial development
'   NCJ 29 Nov 00 - Use clsSubjectNumbering to read Reg. Conds. and Uniqueness checks
'   NCJ 4 Dec 00 - Sort out preservation of leading zeros in Arezzo
'
'   MACRO 2.2
'   NCJ 25 Sep 01 - Use new StudySubject object and goArezzo object
'   DPH 18/04/2002 - Registration using new data transfer form
'----------------------------------------------------------------------------------------'

Option Explicit

' Columns in checks and conditions table
Private Const mnCOL_CODE = 1
Private Const mnCOL_AREZZO = 2
' Only for Uniqueness Checks
Private Const mnCOL_FLAG = 3

' Store person details
Private msTrialName As String
Private mlTrialId As Long
Private mnVersionId As Integer
Private msSite As String
Private mlPersonId As Long
Private moSubject As StudySubject

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
Private moSubjectNumbering As clsSubjectNumbering

' Server connection class
Private moRRConnection As clsRRConnection

'----------------------------------------------------------------------------------------'
Public Sub Initialise(oSubject As StudySubject)
'----------------------------------------------------------------------------------------'
' Initialise registration class for given Trial, Site and Person
' Create SubjectNumbering and ServerConnection classes
'----------------------------------------------------------------------------------------'
                        
    On Error GoTo ErrHandler
    
    Set moSubject = oSubject
    
    ' Store person details from Subject object
    msTrialName = oSubject.StudyDef.Name
    mlTrialId = oSubject.StudyDef.StudyId
    mnVersionId = oSubject.StudyDef.Version
    msSite = oSubject.Site
    mlPersonId = oSubject.PersonId
    
    ' Initialise subject identifier
    msSubjectId = ""
    
    ' Find current registration status
    mnRegStatus = moSubject.RegistrationStatus
    
    ' Load up subject numbering details
    Set moSubjectNumbering = New clsSubjectNumbering
    Call moSubjectNumbering.Load(mlTrialId, mnVersionId)
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
    
    ' Load their RS connection
    Set moRRConnection = New clsRRConnection
    Call moRRConnection.Load(mlTrialId, mnVersionId)

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "Initialise", "clsRegister")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'----------------------------------------------------------------------------------------'
Public Function RegistrationTrigger(oEFI As EFormInstance) As Boolean
'----------------------------------------------------------------------------------------'
' Returns TRUE if this is the form/visit which triggers registration
' If so, sets the subject's registration status
' from NotReady to Ready To Register
'----------------------------------------------------------------------------------------'
    
    RegistrationTrigger = False
    
    ' Are we on the right form and visit?
    If oEFI.eForm.EFormId <> moSubjectNumbering.TriggerFormId Then
        Exit Function
    End If
    If oEFI.VisitInstance.Visit.VisitId <> moSubjectNumbering.TriggerVisitId Then
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
Public Property Get RegistrationStatus() As eRegStatus
'----------------------------------------------------------------------------------------'
' The registration status of this subject
' Can only be called after Initialise
'----------------------------------------------------------------------------------------'

    RegistrationStatus = mnRegStatus
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get TrialId() As Long
'----------------------------------------------------------------------------------------'
' Clinical Trial ID
'----------------------------------------------------------------------------------------'

    TrialId = mlTrialId
    
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
    
    On Error GoTo ErrHandler
    
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

    ' See if they've got a server connection set up
    If moRRConnection.ServerType = eRRServerType.RRNone Then
        Exit Function
    End If
    
    ' If we got this far we'll forge ahead!
    ShouldRegisterSubject = True

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "ShouldRegisterSubject", "clsRegister")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

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
' DPH 18/04/2002 - Registration using new data transfer form
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    ' Initialise to "error" (always assume the worst...)
    mnRegResult = eRegResult.RegError
    
    'Use local registration if RS type is "Local"
    ' or if it's "TrialOffice" and we're on the server
    If moRRConnection.ServerType = eRRServerType.RRLocal _
      Or (moRRConnection.ServerType = eRRServerType.RRTrialOffice And Not gblnRemoteSite) Then
        mnRegResult = RegisterLocally(msSubjectId)
    ElseIf moRRConnection.ServerType = eRRServerType.RRTrialOffice Then
        ' Execute on Trial Office
        ' DPH 18/04/2002 - Registration using new data transfer form (was frmExchangeStatus)
        mnRegResult = frmDataTransfer.TrialOfficeRegistration(msTrialName, _
                            msSite, _
                            mlPersonId, _
                            msPrefixValue, _
                            msSuffixValue, _
                            mnUsePrefix, _
                            mnUseSuffix, _
                            moSubjectNumbering.StartNumber, _
                            moSubjectNumbering.NumberWidth, _
                            msCheckString, _
                            msSubjectId)
    ElseIf moRRConnection.ServerType = eRRServerType.RRRemote Then
        ' Execute remotely
        ' *** TO BE DONE !! ***
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
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "DoRegistration", "clsRegister")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

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
' NCJ 25/9/01 - Use goArezzo instead of frmArezzo
'----------------------------------------------------------------------------------------'
Dim nEligibilityFlag As Integer
Dim sCond As String
Dim vResult As Variant
Dim bEligible As Boolean
Dim sResult As String
Dim colResults As Collection
Dim colArezzoExprs As Collection
Dim lRow As Long

    On Error GoTo ErrHandler
    
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
        Set colResults = goArezzo.EvaluateExpressions(colArezzoExprs)
        ' See if any of them failed
        lRow = 1
        For Each vResult In colResults
            sResult = CStr(vResult)
            ' Flag is field 3
'            nEligibilityFlag = moSubjectNumbering.RegistrationConditions.Field(lRow, mnCOL_FLAG)
            ' For MACRO 2.1, the flag's value is always 1
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
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "CheckEligibilityConditions", "clsRegister")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

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

    On Error GoTo ErrHandler
    
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
            sResult = goArezzo.EvaluateExpression(sArezzoExpr)
            ' Was it OK?
            If goArezzo.ResultOK(sResult) And sResult > "" Then
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
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "EvaluateUniquenessChecks", "clsRegister")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

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

    On Error GoTo ErrHandler
    
    msPrefixValue = ""
    msSuffixValue = ""
    bOK = True
    
    ' Get expression for Prefix
    sArezzoExpr = moSubjectNumbering.Prefix
    If sArezzoExpr > "" Then
        ' Evaluate the expression
        sResult = goArezzo.EvaluateExpression(sArezzoExpr)
        ' Was it OK?
        If goArezzo.ResultOK(sResult) Then
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
            sResult = goArezzo.EvaluateExpression(sArezzoExpr)
            ' Was it OK?
            If goArezzo.ResultOK(sResult) Then
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
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "EvaluatePrefixSuffixValues", "clsRegister")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'
' Tidy up after ourselves
'----------------------------------------------------------------------------------------'

    Set moSubjectNumbering = Nothing
    Set moRRConnection = Nothing
    Set moSubject = Nothing

End Sub

'----------------------------------------------------------------------------------------'
Private Function RegisterLocally(ByRef sSubjId As String) As eRegResult
'----------------------------------------------------------------------------------------'
' Perform LOCAL registration, i.e. using local DLL
' Assume everything already set up and preliminary checks done
'----------------------------------------------------------------------------------------'
Dim oSubjNumbering As clsRSSubjectNumbering

    On Error GoTo ErrHandler
    
    Set oSubjNumbering = New clsRSSubjectNumbering
    With oSubjNumbering
        Call .SetDatabase(goUser.Database.DatabaseCode)
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
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "RegisterLocally", "clsRegister")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'----------------------------------------------------------------------------------------'
Private Sub SetSubjectIdentifier()
'----------------------------------------------------------------------------------------'
' For a successful registration,
' set the subject identifier AND registration status for this subject
' Also add value to Arezzo and save it
'----------------------------------------------------------------------------------------'

    moSubject.RegistrationStatus = eRegStatus.Registered
    moSubject.RegistrationID = msSubjectId
    Call moSubject.SaveRegistration(False)  ' Save both ID and Status

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "SetSubjectIdentifier", "clsRegister")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub SetRegistrationStatus(nStatus As eRegStatus)
'----------------------------------------------------------------------------------------'
' Set the current registration status of this subject
'----------------------------------------------------------------------------------------'

    moSubject.RegistrationStatus = nStatus
    Call moSubject.SaveRegistration(True)   ' Save ONLY the reg. status

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "SetRegistrationStatus", "clsRegister")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

