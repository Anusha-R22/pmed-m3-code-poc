VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClinicalTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsClinicalTest.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, August 2000
'   Purpose:    ClinicalTest Instance class for MACRO
'----------------------------------------------------------------------------------------'
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
Option Explicit


'fields
Public Code As String
Public Description As String
Public Unit As String
Private msClinicalTestGroupCode As String
Private moClinicalTestGroup As clsClinTestGroup

'----------------------------------------------------------------------------------------'
Public Property Get ClinicalTestGroup() As clsClinTestGroup
'----------------------------------------------------------------------------------------'

    Set ClinicalTestGroup = moClinicalTestGroup

End Property

'----------------------------------------------------------------------------------------'
Public Property Let ClinicalTestGroupCode(sClinicalTestGroupCode As String)
'----------------------------------------------------------------------------------------'

' set up linked Clinical Test gRoup
    msClinicalTestGroupCode = sClinicalTestGroupCode
    Set moClinicalTestGroup = New clsClinTestGroup
    moClinicalTestGroup.Load msClinicalTestGroupCode

End Property

'----------------------------------------------------------------------------------------'
Public Property Get ClinicalTestGroupCode() As String
'----------------------------------------------------------------------------------------'

    ClinicalTestGroupCode = msClinicalTestGroupCode

End Property

'----------------------------------------------------------------------------------------'
Public Function Load(sCode As String)
'----------------------------------------------------------------------------------------'
'load up a single ClinicaltestGroup by code
'----------------------------------------------------------------------------------------'

Dim rsTemp As ADODB.Recordset

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "SELECT * FROM ClinicalTest WHERE ClinicalTestCode = '" & sCode & "'", MacroADODBConnection
    Code = rsTemp!ClinicalTestCode
    Description = rsTemp!ClinicalTestDescription
    'TA 25/10/2000: remove null needed
    Unit = RemoveNull(rsTemp!Unit)
    ClinicalTestGroupCode = rsTemp!ClinicalTestGroupCode
    
End Function
