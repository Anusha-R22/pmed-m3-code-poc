VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLockForMIMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsLockForMIMsg.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Toby Aldridge, January 2003
'   Purpose:    class to handle locking when changing the status
'                   of a MIMEssage when use might have the eform open
'----------------------------------------------------------------------------------------'
' REVISIONS:
'
'----------------------------------------------------------------------------------------'

Option Explicit


Private msLockToken As String

Private msStudyName As String
Private msSite As String
Private mlSubjectId As Long


'----------------------------------------------------------------------------------------'
Public Function LockIfNeeded(sStudyName As String, sSite As String, lSubjectId As Long, nMIMType As MIMsgType, _
                            Optional oMIMsg As Object = Nothing, Optional oResponse As Response = Nothing) As Boolean
'----------------------------------------------------------------------------------------'
'lock subject if we don't already have it locked
'oMIMsg can be SDV or Discrepancy
'----------------------------------------------------------------------------------------'
Dim oSubject As StudySubject
Dim bLockSubject As Boolean
Dim oEFI As EFormInstance
Dim oDisc As MIDiscrepancy
Dim oSDV As MISDV

    msStudyName = sStudyName
    msSite = sSite
    mlSubjectId = lSubjectId
    
    If Not oResponse Is Nothing Then
        Set oEFI = oResponse.EFormInstance
    Else
        If FormIsLoaded(g_DATAENTRY_FORM_NAME) Then
            Set oEFI = frmEFormDataEntry.EFI
            If frmEFormDataEntry.SaveNeeded Then
                'they must save before we can change the mimessage status
                DialogInformation "Please save changes to the eForm first"
                LockIfNeeded = False
'EXIT SUB HERE
                Exit Function
            End If
        Else
            Set oEFI = Nothing
        End If
    End If
  
    If oEFI Is Nothing Then
        'we just need to do a simple lock
        bLockSubject = True
    Else
        'we do simple lock if eform readonly
        If (oEFI.ReadOnly) Then
            bLockSubject = True
        Else
           
           Set oSubject = oEFI.VisitInstance.Subject
           If sStudyName = oSubject.StudyCode _
                   And sSite = oSubject.Site _
                       And lSubjectId = oSubject.PersonId Then
               'same subject
               'no need to lock
               bLockSubject = False
'               If moMIMsgs.EFormTaskId = oEFI.EFormTaskId Then
'                   'same eform
'                   'no need to lock
'                   bLockSubject = False
'               Else
'                   'same subject, different eform
'                   'attemp subject lock - it will fail
'                   bLockSubject = True
'               End If
           Else
               'different subject
               'just do a simple lock
               bLockSubject = True
           End If
        End If
        
        
    End If
    
    If bLockSubject Then
        msLockToken = LockSubject(goUser.UserName, goUser.Studies.StudyByName(msStudyName).StudyId, msSite, mlSubjectId)
        'return success when mslocktoken is not nothing
        LockIfNeeded = (msLockToken <> "")
    Else
        'do nothing
        msLockToken = ""
        LockIfNeeded = True
    End If
    
    
    'TA 14/04/2003: reload now we have a lock just incase there have been any changes
    If LockIfNeeded Then
        If Not oMIMsg Is Nothing Then
            Select Case nMIMType
            Case MIMsgType.mimtDiscrepancy
                Set oDisc = oMIMsg
                oDisc.Reload
                
            Case MIMsgType.mimtSDVMark
                Set oSDV = oMIMsg
                oSDV.Reload
            End Select
        End If
    End If

End Function

Public Sub UnlockIfNeeded()
'unlck subject if we have a token

    If msLockToken <> "" Then
        Call UnlockSubject(goUser.Studies.StudyByName(msStudyName).StudyId, msSite, mlSubjectId, msLockToken)
    End If
End Sub

