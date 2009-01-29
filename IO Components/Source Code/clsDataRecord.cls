VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsDataRecord.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, November 2000
'   Purpose:    Record class
'----------------------------------------------------------------------------------------'
' Revisions:

Option Explicit

Private mvRecord As Variant
'0 element of array is unused

'----------------------------------------------------------------------------------------'
Public Sub Init(lFields As Long)
'----------------------------------------------------------------------------------------'
'intialise arecord with specified number of fields
'----------------------------------------------------------------------------------------'

    ReDim mvRecord(lFields) As String

End Sub

'----------------------------------------------------------------------------------------'
Public Property Let Field(lField As Long, sValue As String)
'----------------------------------------------------------------------------------------'
' set the value of a field (default property)
'----------------------------------------------------------------------------------------'

     mvRecord(lField) = sValue

End Property

'----------------------------------------------------------------------------------------'
Public Property Get Field(lField As Long) As String
Attribute Field.VB_UserMemId = 0
'----------------------------------------------------------------------------------------'
' get value from a field (default property)
'----------------------------------------------------------------------------------------'

    Field = mvRecord(lField)

End Property

'----------------------------------------------------------------------------------------'
Public Sub Add(sValue As String)
'----------------------------------------------------------------------------------------'
' add a field
'----------------------------------------------------------------------------------------'

    ReDim Preserve mvRecord(UBound(mvRecord) + 1)
    mvRecord(UBound(mvRecord)) = sValue
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Insert(lBefore As Long, sValue As String)
'----------------------------------------------------------------------------------------'
' insert a field
'----------------------------------------------------------------------------------------'
Dim i As Long

    ReDim Preserve mvRecord(UBound(mvRecord) + 1)
    For i = UBound(mvRecord) To lBefore + 1 Step -1
        mvRecord(i) = mvRecord(i - 1)
    Next
    mvRecord(lBefore) = sValue
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Delete(lField As Long)
'----------------------------------------------------------------------------------------'
' delete a field
'----------------------------------------------------------------------------------------'
Dim i As Long

    
    For i = lField To UBound(mvRecord) - 1
        mvRecord(i) = mvRecord(i + 1)
    Next
    ReDim Preserve mvRecord(UBound(mvRecord) - 1)
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Build(ParamArray vFields() As Variant)
'----------------------------------------------------------------------------------------'
'build a record from a list of parameter strings
'----------------------------------------------------------------------------------------'

Dim i As Long
    ReDim mvRecord(UBound(vFields) + 1) As String
    For i = 0 To UBound(vFields)
        mvRecord(i + 1) = vFields(i)
    Next

End Sub

'----------------------------------------------------------------------------------------'
Public Property Get Fields() As Long
'----------------------------------------------------------------------------------------'
'return the numberof fields
'----------------------------------------------------------------------------------------'

    Fields = UBound(mvRecord)
    
End Property

'----------------------------------------------------------------------------------------'
Public Function Duplicate() As clsDataRecord
'----------------------------------------------------------------------------------------'
' make copy of record
'----------------------------------------------------------------------------------------'

Dim i As Long
    Set Duplicate = New clsDataRecord
    Duplicate.Init Fields
    For i = 1 To Fields
        Duplicate.Field(i) = Field(i)
    Next
    
End Function
