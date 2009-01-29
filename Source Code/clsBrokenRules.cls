VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBrokenRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       clsBrokenRules.cls
'   Author:     Paul Norris 17/08/99
'   Purpose:    Handles class validation for classes that
'               have complex validation rules
'   Sample Use: When a rule is broken in a class add the rule name to this class.
'               To check if the object is valid simply read the Count property.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions: TA 23/11/2000: Check for broken rules also in error handler
'----------------------------------------------------------------------------------------'
Option Explicit

Public Event BrokenRule()
Public Event NoBrokenRules()

Private moBroken As Collection
Private Sub Class_Initialize()
    Set moBroken = New Collection
End Sub
Private Sub Class_Terminate()
    Set moBroken = Nothing
End Sub
' add or remove the rule
Public Sub RuleBroken(sRule As String, bIsBroken As Boolean)
    On Error GoTo ErrorHandler
    
    If bIsBroken Then
        moBroken.Add True, sRule
        RaiseEvent BrokenRule
    Else
        moBroken.Remove sRule
        If moBroken.Count = 0 Then
            RaiseEvent NoBrokenRules
        End If
    End If

Exit Sub

ErrorHandler:

    If moBroken.Count = 0 Then
        RaiseEvent NoBrokenRules
    End If

End Sub
Public Property Get Count() As Integer
    Count = moBroken.Count
End Property
