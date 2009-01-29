VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCoords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999-2000. All Rights Reserved
'   File:       clsLockedFrozenCoords.cls
'   Author:     Steven Morris, Nov 1999
'   Purpose:    Class to contain Row and Col coordinates in Data Browser
'               See also clsColCoords for collection of clsCoords objects
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
    ' NCJ 1/3/00 - Changed Row and Col from Integer to Long
'----------------------------------------------------------------------------------------'


Option Explicit

' Object has two properties, Row and Column

Public Row As Long
Public Col As Long
