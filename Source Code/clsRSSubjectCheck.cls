VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRSSubjectCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       clsRSSubjectCheck.cls
'   Author:     Nicky Johns November 2000
'   Purpose:    SubjectCheck object class for use during registration
'------------------------------------------------------------------------------------------------'
'
'------------------------------------------------------------------------------------------------'
'   Revisions:
'
'------------------------------------------------------------------------------------------------'
Option Explicit

Public CheckCode As String
Public CheckValue As String


Private Sub Class_Initialize()
    CheckCode = ""
    CheckValue = ""
End Sub
