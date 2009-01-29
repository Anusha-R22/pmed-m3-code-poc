VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClinTestGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsClinTestGroup.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, August 2000
'   Purpose:    ClinicalTestGroup Instance class for MACRO
'----------------------------------------------------------------------------------------'
' Revisions:
'TA 27/09/2000: shortened class name from clsClinicalTestGroup to avoid compilation error
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
'----------------------------------------------------------------------------------------'

Option Explicit


'fields
Public Code As String
Public Description As String

'----------------------------------------------------------------------------------------'
Public Function Load(sCode As String)
'----------------------------------------------------------------------------------------'
'load up a single ClinicaltestGroup by code
'----------------------------------------------------------------------------------------'

Dim rsTemp As adodb.Recordset

    Set rsTemp = New adodb.Recordset
    rsTemp.Open "SELECT * FROM ClinicalTestGroup WHERE ClinicalTestGroupCode = '" & sCode & "'", MacroADODBConnection
    Code = rsTemp!ClinicalTestGroupCode
    Description = rsTemp!ClinicalTestGroupDescription
    
End Function
