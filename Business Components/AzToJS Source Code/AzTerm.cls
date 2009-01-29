VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AzTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------
'
'   File: AzTerm.cls
'   Copyright InferMed 2001 All Rights Reserved
'   Author: Nicky Johns
'   Purpose: Object to store an Arezzo term for a quetsion in MACRO
'
'------------------------------------------------------------
' REVISIONS:
'   NCJ 10 Sep 01 - Initial development
'------------------------------------------------------------

Option Explicit

' Simply store the Term and its Question code
Public ArezzoTerm As String
Public QuestionCode As String


