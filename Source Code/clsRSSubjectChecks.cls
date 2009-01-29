VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRSSubjectChecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       clsRSSubjectChecks.cls
'   Author:     Nicky Johns November 2000
'   Purpose:    Collection of clsRSSubjectCheck objects class for use during registration
'------------------------------------------------------------------------------------------------'
'
'------------------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 5 Dec 00 - Tidy up as result of white box test by TA
'
'------------------------------------------------------------------------------------------------'
Option Explicit

Private mcolSubjectChecks As Collection
Private moSubjectCheck As clsRSSubjectCheck

'----------------------------------------------------------------------------------------'
Public Function Add(sCode As String, sValue As String) As clsRSSubjectCheck
'----------------------------------------------------------------------------------------'
' Add a new Patient Check - assume unique sCode
'----------------------------------------------------------------------------------------'

    Set moSubjectCheck = New clsRSSubjectCheck
    moSubjectCheck.CheckCode = sCode
    moSubjectCheck.CheckValue = sValue
    mcolSubjectChecks.Add moSubjectCheck, sCode
    Set Add = moSubjectCheck

End Function

'----------------------------------------------------------------------------------------'
Public Sub Remove(ByVal sCode As String)
'----------------------------------------------------------------------------------------'
' Remove SubjectCheck with code sCode
' Ignores errors if sCode does not exist
'----------------------------------------------------------------------------------------'

    On Error Resume Next
    Call mcolSubjectChecks.Remove(sCode)
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function Item(ByVal sCode As String) As clsRSSubjectCheck
'----------------------------------------------------------------------------------------'
' Get SubjectCheck with code sCode
' Returns Nothing if sCode does not exist
'----------------------------------------------------------------------------------------'

    Set moSubjectCheck = New clsRSSubjectCheck
    On Error Resume Next
    Set moSubjectCheck = mcolSubjectChecks.Item(sCode)
    Set Item = moSubjectCheck
    
End Function

'----------------------------------------------------------------------------------------'
Public Function Count() As Integer
'----------------------------------------------------------------------------------------'
' Count of items in collection
'----------------------------------------------------------------------------------------'

    Count = mcolSubjectChecks.Count

End Function

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'

    Set mcolSubjectChecks = New Collection
    Set moSubjectCheck = New clsRSSubjectCheck
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'

    Set mcolSubjectChecks = Nothing
    Set moSubjectCheck = Nothing

End Sub

'----------------------------------------------------------------------------------------'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
'----------------------------------------------------------------------------------------'
' The enumerator to allow For...Each iterations
'----------------------------------------------------------------------------------------'
    
    Set NewEnum = mcolSubjectChecks.[_NewEnum]

End Function

