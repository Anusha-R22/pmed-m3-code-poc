VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WWWField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------
' File:         WWWField.cls
' Copyright:    InferMed Ltd. 2002 All Rights Reserved
' Author:       I Curtis, 29/04/2003
' Purpose:      Store web response information prior to saving
'--------------------------------------------------------------------------------------------------
' Revisions:
' ic 16/09/2003 added sServerLocaleValue property
' DPH 29/09/2003 added bDone flag
' ic 05/03/2004 added EformUse property
' ic 21/04/2004 added 'bRejected' property, renamed 'bDone' property to 'bProcessed' for clarity
'--------------------------------------------------------------------------------------------------
Option Explicit

Public oResponse As Response
Public oEFormInstance As EFormInstance

Public sWebID As String
Public sEformId As String
Public sElementID As String
Public nRepeat As Integer
Public dblTimestamp As Double
Public sValue As String
Public sServerLocaleValue As String
Public bOK As Boolean
Public sRFO As String
Public bRFOPresent As Boolean
Public sRFC As String
Public sAuthPassword As String
Public sAuthUserName As String
Public sDiscrepancy As String
Public sSDV As String
Public sNote As String
Public sComment As String
Public sUnobtainable As String

Public sAInfo As String

Public bProcessed As Boolean
Public EformUse As eEFormUse
Public bRejected As Boolean

Private Sub Class_Initialize()
    bOK = True
    bRFOPresent = False
    bProcessed = False
    bRejected = False
End Sub
