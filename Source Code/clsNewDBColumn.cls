VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNewDBColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       clsNewDbColumn.cls
'   Author:     Toby Aldridge, October 2000
'   Purpose:    Class that contains information of a column changes between versions
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'   Revisions:
'
'   Mo Morris 25/9/01   Changes around  expanding the clsNewDBColumn functionality
'                       to handle the dropping of columns.
'                       New class fields ChangeType (containg 'NEWCOLUMN' or 'DROPCOLUMN')
'                       and ColumnNumber added corresponding to new table fields
'                       NewDBColumn.ChangeType and NewDBColumn.ColumnNumber.
'--------------------------------------------------------------------------------

Option Explicit


Public Version As Long
Public TableName As String
Public ColumnName As String
Public ColumnOrder As Variant
Public DefaultValue As String
Public Segment As Long
Public ChangeType As String
Public ColumnNumber As Variant
