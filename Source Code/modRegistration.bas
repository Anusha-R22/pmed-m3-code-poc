Attribute VB_Name = "modRegistration"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000-2006 All Rights Reserved
'   File:       modRegistration.bas
'   Author:     Nicky Johns, November 2000
'   Purpose:    Handle subject registration in MACRO Data Management
'               Most of the work is done by clsRegister
'               but this module makes the calls and gives user interface messages etc.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 23/11/00 - Initial development
'   NCJ 30/11/00 - Give error messages for ALL failures of registration
'               Update subject label after registration
'   NCJ 4/12/00 - Take account of Change Data rights and Update Mode
'               Changed some message wording
' MACRO 2.2
'   NCJ 25 Sep 01 - Changed InitialiseRegistration to take StudySubject
'               and TimeToRegister to take eFormInstance
' NCJ 1 Oct 01 - Removed redundant call to UpdateSubjectDetails
'
' NCJ 13 May 03 - RegisterSubject now returns Boolean result
'           Removed redundant calls to EnableRegistrationMenu
' NCJ 11 Dec 06 - Issue 2847 - Check EFI's completeness in TimeToRegister
'----------------------------------------------------------------------------------------'

Option Explicit
' Registration class
Private moRegister As clsRegister

Private Const msCANNOT_REGISTER = "This subject cannot be registered because " & vbCrLf
Private Const msCONTACT_ADMINISTRATOR = vbCrLf & "Please contact your study administrator."

'----------------------------------------------------------------------------------------'
Public Sub InitialiseRegistration(oSubject As StudySubject)
'----------------------------------------------------------------------------------------'
' Initialise registration for the current subject
'----------------------------------------------------------------------------------------'
    
    Set moRegister = New clsRegister
    ' Set up with current subject details
    Call moRegister.Initialise(oSubject)
    ' Call EnableRegistrationMenu
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub CloseRegistration()
'----------------------------------------------------------------------------------------'
' Tidy up and close down classes etc.
'----------------------------------------------------------------------------------------'

    Set moRegister = Nothing

End Sub

'----------------------------------------------------------------------------------------'
Public Function TimeToRegister(oEFI As EFormInstance) As Boolean
'----------------------------------------------------------------------------------------'
' Check whether we're on the right form and visit for registration
' Assumes InitialiseRegistration has already been done
' NCJ 11 Dec 06 - Issue 2847 - Check EFI's completeness first
'----------------------------------------------------------------------------------------'

    TimeToRegister = False
    If oEFI.Complete Then
        TimeToRegister = moRegister.RegistrationTrigger(oEFI)
    End If

End Function

'----------------------------------------------------------------------------------------'
Public Function RegisterSubject() As Boolean
'----------------------------------------------------------------------------------------'
' Handle current subject's registration
' Assume InitialiseRegistration has already been done
' Enable/Disable "Register Subject" menu item
' NCJ 13 May 03 - Returns TRUE if registration happened successfully, or FALSE otherwise
'----------------------------------------------------------------------------------------'
    
    RegisterSubject = SubjectRegistration
        
    ' The subject's registration status will now have been set
    ' Call EnableRegistrationMenu
    
End Function

'----------------------------------------------------------------------------------------'
Private Function SubjectRegistration() As Boolean
'----------------------------------------------------------------------------------------'
' Handle current subject's registration
' Assume InitialiseRegistration has already been done
' NCJ 13 May 03 - Returns TRUE if registration successful, or FALSE otherwise
'----------------------------------------------------------------------------------------'
Dim sPrompt As String
Dim nResult As Integer

    On Error GoTo ErrHandler
    
    SubjectRegistration = False
    
    ' Don't do registration if user can't Change Data
    If (Not goUser.CheckPermission(gsFnChangeData)) Or goStudyDef.Subject.ReadOnly Then
        Exit Function
    End If
    
    ' See if registration is appropriate
    If Not moRegister.ShouldRegisterSubject Then
        Exit Function
    End If
    
    ' Give an appropriate "start" message
    Select Case moRegister.RegistrationStatus
    Case eRegStatus.Failed, eRegStatus.Ineligible
        ' See if they want to re-do a failed attempt
        sPrompt = "This subject previously failed to register successfully." _
                & vbCrLf & "Would you like to try registration again?"
        If DialogQuestion(sPrompt) <> vbYes Then
            Exit Function
        End If
    Case eRegStatus.NotReady, eRegStatus.Ready
        ' See if they would like to register the subject now
        ' NB At this stage we have NOT yet checked whether they're eligible etc.
        sPrompt = "Would you like to proceed with registration for this subject?"
        If DialogQuestion(sPrompt) <> vbYes Then
            Exit Function
        End If
    End Select
    
    ' Check the registration conditions
    If Not moRegister.IsEligible Then
        sPrompt = msCANNOT_REGISTER _
                    & "the registration conditions for this study have not been met." _
                    & msCONTACT_ADMINISTRATOR
        Call DialogError(sPrompt)
        Exit Function
    End If

    ' Get the identifier prefix and suffix
    If Not moRegister.EvaluatePrefixSuffixValues Then
        sPrompt = msCANNOT_REGISTER _
                    & "some identifier information is missing." _
                    & msCONTACT_ADMINISTRATOR
        Call DialogError(sPrompt)
        Exit Function
    End If
    
    ' Get the uniqueness check values
    If Not moRegister.EvaluateUniquenessChecks Then
        sPrompt = msCANNOT_REGISTER _
                    & "some uniqueness check information is missing." _
                    & msCONTACT_ADMINISTRATOR
        Call DialogError(sPrompt)
        Exit Function
    End If
    
    ' OK - we're all ready to go!
    HourglassOn
    nResult = moRegister.DoRegistration
    HourglassOff
    Select Case nResult
    Case eRegResult.RegOK
        sPrompt = "This subject has been successfully registered" _
                & vbCrLf & "with unique identifier: " _
                & vbCrLf & moRegister.SubjectIdentifier
        Call DialogInformation(sPrompt)
        SubjectRegistration = True

    Case eRegResult.RegNotUnique
        sPrompt = "Registration failed because this subject's details are not unique in this study." _
                & msCONTACT_ADMINISTRATOR
        Call DialogError(sPrompt)
        
'    Case eRegResult.RegError, eRegResult.RegMissingInfo
    Case Else
        ' Anything else represents an error
        sPrompt = "An error occurred during the registration of this subject."
        Call DialogError(sPrompt)
        
    End Select
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "SubjectRegistration", "modRegistration")
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
Public Function EnableRegistrationMenu() As Boolean
'----------------------------------------------------------------------------------------'
' Decide whether to Enable/disable registration menu item
' NCJ 4/12/00 - Also check change data rights & update mode
' TA - This DOES NOT actually do the enabling!
'----------------------------------------------------------------------------------------'
    
    EnableRegistrationMenu = False
    
    ' NCJ 18 Jun 03 - Use the ShouldRegisterSubject routine (Bug 1003)
    ' Also check new Register Subject permission
    If (goUser.CheckPermission(gsFnChangeData) _
     And goUser.CheckPermission(gsFnRegisterSubject) _
     And (Not goStudyDef.Subject.ReadOnly)) Then
        EnableRegistrationMenu = moRegister.ShouldRegisterSubject
    End If
    
'    Select Case moRegister.RegistrationStatus
'    Case eRegStatus.NotReady, eRegStatus.Registered
'        ' Either not ready or already registered - leave menu item disabled
'    Case eRegStatus.Failed, eRegStatus.Ineligible, eRegStatus.Ready
'        If (goUser.CheckPermission(gsFnChangeData) And (Not goStudyDef.Subject.ReadOnly)) Then
'            ' Ready or previously failed - let them try again
''            frmMenu.mnuODRegistration.Enabled = True
'            EnableRegistrationMenu = True
'        End If
'    End Select

End Function
