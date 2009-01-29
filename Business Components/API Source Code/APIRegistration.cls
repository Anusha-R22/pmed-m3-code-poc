VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APIRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2006 All Rights Reserved
'   File:       APIRegistration.cls
'   Author:     Nicky Johns, August 2006
'   Purpose:    Handle subject registration in the MACRO API
'               Based on copy of modRegistrationWWW from Web DE
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 15 Aug 06 - Initial development, based on copy of modRegistrationWWW
'----------------------------------------------------------------------------------------'

Option Explicit

'----------------------------------------------------------------------------------------'
Public Function DoRegistration(oSubject As StudySubject, _
                                ByVal sConnection As String, _
                                ByVal sDatabaseCode As String) As eAPIRegResult
'----------------------------------------------------------------------------------------'
' Handle registration if appropriate
' Returns result code of what happened
'----------------------------------------------------------------------------------------'
Dim oRegister As clsRegisterWWW
Dim nResult As Integer

    On Error GoTo ErrHandler
    
    If oSubject.ReadOnly Then
        DoRegistration = eAPIRegResult.SubjectReadOnly
        Exit Function
    End If
    
    Set oRegister = New clsRegisterWWW
    ' Set up with current subject details
    Call oRegister.Initialise(oSubject, sConnection, sDatabaseCode)
    
    If Not oRegister.ShouldRegisterSubject Then
        ' Either already registered, or not ready, or no registration to be done
        DoRegistration = eAPIRegResult.NotReady
    ElseIf Not oRegister.IsEligible Then
        ' The registration conditions have not been met
        DoRegistration = eAPIRegResult.Ineligible
    ElseIf Not oRegister.EvaluatePrefixSuffixValues Then
        'Some identifier information is missing
        DoRegistration = eAPIRegResult.MissingInfo
    ElseIf Not oRegister.EvaluateUniquenessChecks Then
        ' Some uniqueness check information is missing
        DoRegistration = eAPIRegResult.MissingInfo
        
    Else
    
        ' OK - we're all ready to go!
        nResult = oRegister.DoRegistration
        
        Select Case nResult
        Case eRegResult.RegOK
            DoRegistration = eAPIRegResult.Success
        Case eRegResult.RegNotUnique
            DoRegistration = eAPIRegResult.NotUnique
        Case Else
            ' Anything else represents an error
            DoRegistration = eAPIRegResult.UnknownError
        End Select
        
    End If
    
    Set oRegister = Nothing
    
Exit Function
ErrHandler:
  Err.Raise Err.Number, , Err.Description & "|APIRegistration.DoRegistration"

End Function

