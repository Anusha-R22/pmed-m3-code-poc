VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EFormItemSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   File:       EFormItemSD.cls
'   Author:     Nicky Johns, November 2001
'   Purpose:    EFormItem object to represent a data entry field
'               on an eForm in MACRO SD 3.0 and later versions.
'-----------------------------------------------------------------------------------
' REVISIONS
'   NCJ 27 Nov 01 - Copied from EFormItem.cls in MACRO DE
'-----------------------------------------------------------------------------------

Option Explicit

' Note that for non-group items, RowNo will be 0

Public Element As CRFElement
Public RowNo As Integer
Public ControlIndex As Integer



