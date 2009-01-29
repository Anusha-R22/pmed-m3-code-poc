Attribute VB_Name = "modRegistrationWWW"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2003 All Rights Reserved
'   File:       modRegistrationWWW.bas
'   Author:     Nicky Johns, June 2003
'   Purpose:    Handle subject registration in MACRO WWW
'               Based on copy of modRegistration from Windows DE
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 12 Jun 03 - Initial development, based on copy of modRegistration from Windows DE
'   ic 19/06/2003 dont check for readonly subject in ShouldEnableRegistrationMenu()
'   NCJ 1st Oct 03 - Added Randomisation (conditionally compiled if RANDOM = 1)
'   NCJ 6 Jan 04 - Removed conditional compilation for Randomisation (include as standard for 3.0.61)
'----------------------------------------------------------------------------------------'

Option Explicit

Private Const msCANNOT_REGISTER = "This subject cannot be registered because " & vbCrLf
Private Const msCONTACT_ADMINISTRATOR = vbCrLf & "Please contact your study administrator."

' Possible randomisation results
Private Enum RandomisationResult
    rrSuccess = 0
    rrCantStratify = 1
    rrNoTreatments = 2
    rrError = 3
End Enum

'----------------------------------------------------------------------------------------'
Public Function DoRegistration(oSubject As StudySubject, _
                                ByVal lEFITaskId As Long, _
                                ByVal sConnection As String, _
                                ByVal sDatabaseCode As String, _
                                ByRef sReturn As String) As Boolean
'----------------------------------------------------------------------------------------'
' Handle registration if appropriate
' If lEFITaskId > 0 , check that this is the eForm that triggers registration
' If lEFITaskId = 0, do not check for trigger eForm (e.g. when registering from menu item)
' Returns TRUE if registration happened, and sReturn is message containing the registration ID
' Returns FALSE if registration didn't happen, and
'   if sReturn = "" then it wasn't meant to happen (so ignore)
'   if sReturn > "" then it contains an error message for the user
'----------------------------------------------------------------------------------------'
Dim oRegister As clsRegisterWWW
Dim oEFI As EFormInstance
' NCJ 1st Oct 03 - for Randomisation
Dim colRResults As Collection

    On Error GoTo ErrHandler
    
    DoRegistration = False
    sReturn = ""
    
    If oSubject.ReadOnly Then Exit Function
    
    Set oRegister = New clsRegisterWWW
    ' Set up with current subject details
    Call oRegister.Initialise(oSubject, sConnection, sDatabaseCode)
    
    ' Get eForm instance if we have an eform task ID
    If lEFITaskId > 0 Then
        Set oEFI = oSubject.eFIByTaskId(lEFITaskId)
    End If
    
    DoRegistration = RegisterSubject(oRegister, oSubject, oEFI, sReturn)
    
    Set oEFI = Nothing
    Set oRegister = Nothing

    ' NCJ 1st Oct 03 - Add Randomisation in here for an eForm save
    If lEFITaskId > 0 Then
        ' Do randomisation but ignore results
        Call RandomiseSubject(sConnection, oSubject, colRResults)
        Set colRResults = Nothing
    End If
    
Exit Function
ErrHandler:
  Err.Raise Err.Number, , Err.Description & "|modRegistrationWWW.RegisterSubject"

End Function

'----------------------------------------------------------------------------------------'
Private Function RegisterSubject(oRegister As clsRegisterWWW, _
                                oSubject As StudySubject, _
                                oEFI As EFormInstance, _
                                ByRef sReturn As String) As Boolean
'----------------------------------------------------------------------------------------'
' Handle subject's registration
' Returns TRUE if registration successful, or FALSE otherwise, and sets sReturn as appropriate
'----------------------------------------------------------------------------------------'
Dim nResult As Integer

    On Error GoTo ErrHandler
    
    RegisterSubject = False
    
    ' Only check the trigger eForm if we have an eForm instance
    If Not oEFI Is Nothing Then
        ' If it's not the trigger eform, forget it
        If Not oRegister.RegistrationTrigger(oEFI) Then Exit Function
    End If
    
    ' See if registration is appropriate
    If Not oRegister.ShouldRegisterSubject Then Exit Function
    
    ' Here we just go ahead and try for registration (we don't ask the user if they want to!)
    
    ' Check the registration conditions
    If Not oRegister.IsEligible Then
        sReturn = msCANNOT_REGISTER _
                    & "the registration conditions for this study have not been met." _
                    & msCONTACT_ADMINISTRATOR
        Exit Function
    End If

    ' Get the identifier prefix and suffix
    If Not oRegister.EvaluatePrefixSuffixValues Then
        sReturn = msCANNOT_REGISTER _
                    & "some identifier information is missing." _
                    & msCONTACT_ADMINISTRATOR
        Exit Function
    End If
    
    ' Get the uniqueness check values
    If Not oRegister.EvaluateUniquenessChecks Then
        sReturn = msCANNOT_REGISTER _
                    & "some uniqueness check information is missing." _
                    & msCONTACT_ADMINISTRATOR
        Exit Function
    End If
    
    ' OK - we're all ready to go!
    nResult = oRegister.DoRegistration
    Select Case nResult
    Case eRegResult.RegOK
        sReturn = "This subject has been successfully registered" _
                & vbCrLf & "with unique identifier: " _
                & vbCrLf & oRegister.SubjectIdentifier
        RegisterSubject = True

    Case eRegResult.RegNotUnique
        sReturn = "Registration failed because this subject's details are not unique in this study." _
                & msCONTACT_ADMINISTRATOR
        
    Case Else
        ' Anything else represents an error
        sReturn = "An error occurred during the registration of this subject."
        
    End Select
    
Exit Function
ErrHandler:
  Err.Raise Err.Number, , Err.Description & "|modRegistrationWWW.RegisterSubject"

End Function

'----------------------------------------------------------------------------------------'
Public Function ShouldEnableRegistrationMenu(oSubject As StudySubject) As Boolean
'----------------------------------------------------------------------------------------'
' Decide whether to Enable/disable registration menu item for the Web
'
' revisions
' ic 19/06/2003 dont check for readonly subject
'----------------------------------------------------------------------------------------'
    
    ShouldEnableRegistrationMenu = False
    
    'ic 19/06/2003 subject may be readonly, ignore
    'If oSubject.ReadOnly Then Exit Function
    
    Select Case oSubject.RegistrationStatus
    Case eRegStatus.NotReady, eRegStatus.Registered
        ' Either not ready or already registered - leave menu item disabled
    Case eRegStatus.Failed, eRegStatus.Ineligible, eRegStatus.Ready
        ' Ready or previously failed - let them try again
        ShouldEnableRegistrationMenu = True
    End Select

End Function

' **** FROM HERE TO THE  END IS RANDOMISATION CODE ****
' **** NCJ 2nd Oct 03 ************************
' **** Include it unconditionally for build 3.0.61 - NCJ 6 Jan 04 ****

'----------------------------------------------------------------------------------------'
Private Function RandomiseSubject(ByVal sDBConnection As String, _
                oSubject As StudySubject, _
                ByRef colResults As Collection) As Boolean
'----------------------------------------------------------------------------------------'
' Randomise a subject and return results in colResults
'----------------------------------------------------------------------------------------'
Dim oDBCon As ADODB.Connection
Dim sSQL As String
Dim rsRands As ADODB.Recordset
Dim sRandCode As String
Dim sStratValue As String
Dim sTreatment As String
Dim bHaveRandomised As Boolean

    On Error GoTo ErrHandler
    
    ' Have we randomised?
    bHaveRandomised = False
    Set colResults = New Collection
    
    Set oDBCon = New ADODB.Connection
    Call oDBCon.Open(sDBConnection)
    oDBCon.CursorLocation = adUseClient
    
    ' Check we have the necessary randomisation tables
    If RandomisationAvailable(oDBCon) Then
        ' Get the possible randomisation codes
        sSQL = "SELECT RandomisationCode, RandomisationCond, StratificationExpr" _
            & " FROM Randomisation WHERE " _
            & " ClinicalTrialId = " & oSubject.StudyId
        Set rsRands = New ADODB.Recordset
        rsRands.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    
        ' See if any need processing
        If rsRands.RecordCount > 0 Then
            rsRands.MoveFirst
            Do While Not rsRands.EOF
                sRandCode = rsRands!RandomisationCode
                If Not IsSubjectRandomised(oSubject, oDBCon, sRandCode) Then
                    ' See if we need to process this randomisation
                    If ShouldRandomiseSubject(oSubject, rsRands!RandomisationCond) Then
                        ' We're going to attempt a randomisation
                        bHaveRandomised = True
                        sStratValue = EvaluateStratification(oSubject, rsRands!StratificationExpr)
                        If sStratValue > "" Then
                            ' Assign treatment
                            Call AssignTreatment(oSubject, oDBCon, colResults, sRandCode, sStratValue)
                        Else
                            ' Add a "Can't stratify" error for this randomisation
                            Call IncludeResult(colResults, rrCantStratify, sRandCode, "")
                        End If
                    End If
                End If
                rsRands.MoveNext
            Loop
        End If
        rsRands.Close
        Set rsRands = Nothing
    End If
    
    Set oDBCon = Nothing
    
    RandomiseSubject = bHaveRandomised
    
Exit Function
ErrHandler:
    ' Set the results to show an error
    Call IncludeResult(colResults, rrError, "", Err.Description)
    ' Signal that there's something to look at
    RandomiseSubject = True
    
End Function

'----------------------------------------------------------------------------------------'
Private Function IsSubjectRandomised(oSubject As StudySubject, _
                                    oDBCon As ADODB.Connection, _
                                    ByVal sRandCode As String) As Boolean
'----------------------------------------------------------------------------------------'
' Returns TRUE if subject is already randomised on this randomisation
'----------------------------------------------------------------------------------------'
Dim rsRands As ADODB.Recordset
Dim sSQL As String

    IsSubjectRandomised = False
    
    On Error GoTo ErrHandler
    
    ' See if this subject has been allocated a treatment on this randomisation
    sSQL = "SELECT Treatment FROM Treatments WHERE " _
        & " ClinicalTrialId = " & oSubject.StudyId _
        & " AND RandomisationCode = '" & sRandCode & "'" _
        & " AND TrialSite = '" & oSubject.Site & "'" _
        & " AND PersonId = " & oSubject.PersonId
    Set rsRands = New ADODB.Recordset
    rsRands.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    If rsRands.RecordCount > 0 Then
        ' They've been done
        IsSubjectRandomised = True
    End If
    rsRands.Close
    
    Set rsRands = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modRegistrationWWW.IsSubjectRandomised"

End Function

'----------------------------------------------------------------------------------------'
Private Function ShouldRandomiseSubject(oSubject As StudySubject, _
                                        ByVal sRandCond As String) As Boolean
'----------------------------------------------------------------------------------------'
' Returns TRUE if the randomisation condition is true
'----------------------------------------------------------------------------------------'
Dim sResult As String

    sResult = oSubject.Arezzo.EvaluateExpression(sRandCond)
    ShouldRandomiseSubject = (sResult = "true")

End Function

'----------------------------------------------------------------------------------------'
Private Function TreatmentDataItem(ByVal sRandCode As String) As String
'----------------------------------------------------------------------------------------'
' Returns the AREZZO data item name of the treatment for this randomisation
'----------------------------------------------------------------------------------------'

    TreatmentDataItem = "person:" & LCase(sRandCode) & ":treatment"
    
End Function

'----------------------------------------------------------------------------------------'
Private Function EvaluateStratification(oSubject As StudySubject, _
                                    ByVal sStratExpr As String) As String
'----------------------------------------------------------------------------------------'
' Evaluate stratification expression
' Returns empty string if no go
'----------------------------------------------------------------------------------------'
Dim sStrat As String

    sStrat = oSubject.Arezzo.EvaluateExpression(sStratExpr)
    If Not oSubject.Arezzo.ResultOK(sStrat) Then
        sStrat = ""
    End If
    EvaluateStratification = sStrat

End Function

'----------------------------------------------------------------------------------------'
Private Sub IncludeResult(colResults As Collection, _
                            ByVal enResultType As RandomisationResult, _
                            ByVal sRandCode As String, _
                            ByVal sTreatment As String)
'----------------------------------------------------------------------------------------'
' Add a result to the results collection
'----------------------------------------------------------------------------------------'
Dim sResult As String
Const sSEP = "|"

    sResult = enResultType & sSEP & sRandCode & sSEP & sTreatment
    colResults.Add sResult

End Sub

'----------------------------------------------------------------------------------------'
Private Sub AssignTreatment(oSubject As StudySubject, _
                                oDBCon As ADODB.Connection, _
                                colResults As Collection, _
                                ByVal sRandCode As String, _
                                ByVal sStratValue As String)
'----------------------------------------------------------------------------------------'
' Get the treatment for this subject
' and add an entry to the colResults collection accordingly
' NCJ 6 Oct 03 - Add AREZZO data first, before updating Treatments table
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim oTimezone As Timezone
Dim rsTreats As ADODB.Recordset
Dim sTreat As String

    On Error GoTo ErrHandler
    
    sTreat = ""
    ' Get the next unassigned treatment
    sSQL = "SELECT * FROM Treatments WHERE" _
        & " ClinicalTrialId = " & oSubject.StudyId _
        & " AND RandomisationCode = '" & sRandCode & "'" _
        & " AND StratificationValue = '" & sStratValue & "'" _
        & " AND TrialSite IS NULL" _
        & " ORDER BY TreatSeqNo "
    
    Set rsTreats = New ADODB.Recordset
    rsTreats.Open sSQL, oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    ' Hopefully there is a record
    If rsTreats.RecordCount > 0 Then
        rsTreats.MoveFirst
        ' Quickly mark this one as ours!
        rsTreats!TrialSite = oSubject.Site
        Call rsTreats.Update
        ' Haul out the treatment
        sTreat = rsTreats!Treatment
        ' Add the data to AREZZO first to make sure we can do it
        Call oSubject.Arezzo.AddData(TreatmentDataItem(sRandCode), sTreat)
        If oSubject.Save = eSaveResponsesResult.srrSuccess Then
            ' Fill in the rest of the subject details
            rsTreats!PersonId = oSubject.PersonId
            If oSubject.Label > "" Then
                rsTreats!SubjectLabel = oSubject.Label
            End If
            rsTreats!RandomiseDate = IMedNow
            Set oTimezone = New Timezone
            rsTreats!RandomiseDate_TZ = oTimezone.TimezoneOffset
            Set oTimezone = Nothing
            Call rsTreats.Update
            ' Add this successful result to the collection
            Call IncludeResult(colResults, rrSuccess, sRandCode, sTreat)
        Else
            ' Can't update patient state so give up
            ' Release this treatment
            rsTreats!TrialSite = Null
            Call rsTreats.Update
            Call IncludeResult(colResults, rrError, sRandCode, sTreat)
        End If
    Else
        ' If the recordset is empty, we either have an unknown stratification value
        ' or we've run out of treatments for this value
        Call IncludeResult(colResults, rrNoTreatments, sRandCode, "")
    End If

    rsTreats.Close
    Set rsTreats = Nothing

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "modRegistrationWWW.AssignTreatment"
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function RandomisationAvailable(oDBCon As ADODB.Connection) As Boolean
'----------------------------------------------------------------------------------------'
' Are the randomisation tables available in this database?
'----------------------------------------------------------------------------------------'
Dim rsRands As ADODB.Recordset

    RandomisationAvailable = True
    
    On Error GoTo NoTable
    
    Set rsRands = New ADODB.Recordset
    ' Just try and access the table with a dummy select
    rsRands.Open "SELECT RandomisationCode FROM Randomisation WHERE 1=2", oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    rsRands.Close
    rsRands.Open "SELECT RandomisationCode FROM Treatments WHERE 1=2", oDBCon, adOpenKeyset, adLockPessimistic, adCmdText
    rsRands.Close
    Set rsRands = Nothing

Exit Function
NoTable:
    ' Error reading from tables so return FALSE
    Set rsRands = Nothing
    RandomisationAvailable = False
End Function

' *************** END OF RANDOMISATION CODE **********************************************************
