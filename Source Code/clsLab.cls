VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsLab.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, August 2000
'   Purpose:    Labt Instance class for MACRO
'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 27/9/00 - Added Load method
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
' NCJ 12/10/00 - Added Site and Changed
'
'----------------------------------------------------------------------------------------'

Option Explicit


'fields
Public Code As String
Public Description As String
' NCJ 12/10/00 Added Site and Changed
Public Site As Variant
Public Changed As Changed

'----------------------------------------------------------------------------------------'
Public Function Load(sCode As String)
'----------------------------------------------------------------------------------------'
' Load up a single Lab by Code
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    Set rsTemp = New ADODB.Recordset
    sSQL = "SELECT * FROM Laboratory WHERE LaboratoryCode = '" & sCode & "'"
    ' NCJ 23/3/01 - Added parameters on to .Open to allow a recordcount
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    ' NCJ 22/3/01 - Check the lab is defined
    If rsTemp.RecordCount > 0 Then
        Code = rsTemp!LaboratoryCode
        Description = rsTemp!LaboratoryDescription
        Site = rsTemp!Site
        Changed = rsTemp!Changed
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
End Function

Public Property Get SiteText() As String

    SiteText = RemoveNull(Site)
    
End Property

Public Property Get ChangedText() As String

    If Changed = 1 Then ' changed.changed
        ChangedText = "Yes"
    Else
        ChangedText = ""
    End If
    
End Property

