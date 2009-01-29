VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsMenuItem.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Toby Aldridge, October 2002
'   Purpose:    Class to represent a single menuitem on a popup list
'----------------------------------------------------------------------------------------'


Option Explicit

'fields
Public Key As String
Public Caption As String
Public Enabled As Boolean
Public Checked As Boolean

'----------------------------------------------------------------------------------------'
Public Sub Init(sKey As String, sCaption As String, bEnabled As Boolean, _
                    Optional bChecked As Boolean = False)
'----------------------------------------------------------------------------------------'

    Key = sKey
    Caption = sCaption
    Enabled = bEnabled
    Checked = bChecked
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub MakeSeparator()
'----------------------------------------------------------------------------------------'

    Key = ""
    Caption = "-"
    Enabled = True
    Checked = False
    
End Sub


