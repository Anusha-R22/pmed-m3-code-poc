VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TempResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'----------------------------------------------------
' File: TempResponse.cls
' Author: Nicky Johns
' Copyright: InferMed, June 2001, All Rights Reserved
' Temporary store for Response values in MACRO 2.2
' This class is used to store the results of response validation
' until the user has gone through the RFC, Overrule and Authorisation hoops.
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 5-10 Jul 01 - Initial development
' TA  27/07/2001 - TimeStamp added
' NCJ 15 Oct 01 - Changed NRStatus from String to Integer
' RS 21/9/2002 - Added TimeStampTZ property
' RS 21/1/2003 - Added DisplayValue property
'----------------------------------------------------

Option Explicit

Public Value As String
Public ValueCode As String
Public ArezzoValue As String
Public ResponseStatus As Integer
Public ValidationMessage As String
Public LabTestDate As Double
Public NRStatus As Integer
Public CTCGrade As Integer
Public TimeStamp As Double
Public DisplayValue As String   ' RS 21/01/2003 Used for displaying local format (date/time)
 
 
 
