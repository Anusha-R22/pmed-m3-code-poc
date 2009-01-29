VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsDataTable.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, November 2000
'   Purpose:    Table class
'----------------------------------------------------------------------------------------'
' Revisions:

Option Explicit

Public Event FieldChanged(lRow As Long, lCol As Long)

'each row in the table is a record
Private mcolRecords As Collection

'collect the headings (optional)
Public Headings As clsDataRecord

'----------------------------------------------------------------------------------------'
Public Function GetHeadingColumn(sHeading As String) As Long
'----------------------------------------------------------------------------------------'
'return the column no from a heading
'----------------------------------------------------------------------------------------'

Dim i As Long
    GetHeadingColumn = -1
    For i = 1 To Headings.Fields
        If lCase(Headings(i)) = lCase(sHeading) Then
            GetHeadingColumn = i
            Exit For
        End If
    Next
End Function

'----------------------------------------------------------------------------------------'
Public Property Get Field(lRow As Long, lCol As Long) As String
Attribute Field.VB_UserMemId = 0
'----------------------------------------------------------------------------------------'
' return the value from a field (default property)
'----------------------------------------------------------------------------------------'

    Field = mcolRecords.Item(lRow)(lCol)

End Property

'----------------------------------------------------------------------------------------'
Public Property Let Field(lRow As Long, lCol As Long, sValue As String)
'----------------------------------------------------------------------------------------'
' set the value of a field (default property)
'----------------------------------------------------------------------------------------'

     mcolRecords.Item(lRow)(lCol) = sValue
     RaiseEvent FieldChanged(lRow, lCol)

End Property

'----------------------------------------------------------------------------------------'
Public Sub Build(ParamArray vRecords() As Variant)
'----------------------------------------------------------------------------------------'
' build a table from a parameter list of records
'----------------------------------------------------------------------------------------'

Dim i As Long

    Set Headings = New clsDataRecord
    Headings.Init UBound(vRecords) + 1
    For i = 0 To UBound(vRecords)
        Call mcolRecords.Add(vRecords(i))
    Next
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Add(recRecord As clsDataRecord)
'----------------------------------------------------------------------------------------'
' add a row(record) to a table
'----------------------------------------------------------------------------------------'

    Call mcolRecords.Add(recRecord)

End Sub

'----------------------------------------------------------------------------------------'
Public Property Get Row(lRow As Long) As clsDataRecord
'----------------------------------------------------------------------------------------'
' retrun a row of the table as a record
'----------------------------------------------------------------------------------------'

    Set Row = mcolRecords.Item(lRow)

End Property

'----------------------------------------------------------------------------------------'
Public Sub Insert(recRow As clsDataRecord, lBeforeRow)
'----------------------------------------------------------------------------------------'
' insert a record as a row to the table
'----------------------------------------------------------------------------------------'

    mcolRecords.Add recRow, , lBeforeRow

End Sub

'----------------------------------------------------------------------------------------'
Public Sub Delete(lRow As Long)
'----------------------------------------------------------------------------------------'
' delete a row from the table
'----------------------------------------------------------------------------------------'

    mcolRecords.Remove lRow

End Sub

'----------------------------------------------------------------------------------------'
Public Sub InsertColumn(lBefore As Long, Optional sHeading As String = "", Optional sDefault As String = "")
'----------------------------------------------------------------------------------------'
' Insert a column to the table (a field to each row record)
'----------------------------------------------------------------------------------------'

Dim recRecord As clsDataRecord

    Headings.Insert lBefore, sHeading
    For Each recRecord In mcolRecords
        recRecord.Insert lBefore, sDefault
    Next
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub DeleteColumn(lCol As Long)
'----------------------------------------------------------------------------------------'
' delete a column to the table (a field to each row record)
'----------------------------------------------------------------------------------------'

Dim recRecord As clsDataRecord

    Headings.Delete lCol
    For Each recRecord In mcolRecords
        recRecord.Delete lCol
    Next
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub AddColumn(Optional sHeading As String = "", Optional sDefault As String = "")
'----------------------------------------------------------------------------------------'
' add a column to the table (a field to each row record)
'----------------------------------------------------------------------------------------'

Dim recRecord As clsDataRecord

    Headings.Add sHeading
    For Each recRecord In mcolRecords
        recRecord.Add (sDefault)
    Next
    
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get Rows()
'----------------------------------------------------------------------------------------'
' return the number fo rows
'----------------------------------------------------------------------------------------'

    Rows = mcolRecords.Count
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get Cols()
'----------------------------------------------------------------------------------------'
' return the number of cols in the table (uses first row fields count)
'----------------------------------------------------------------------------------------'

    Cols = Headings.Fields
    
End Property

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Set mcolRecords = New Collection
    Set Headings = New clsDataRecord
    Headings.Init 0
    
End Sub

Public Function Duplicate() As clsDataTable
Dim i As Long
Dim tblTemp As clsDataTable

    Set tblTemp = New clsDataTable
    
    Set tblTemp.Headings = Headings.Duplicate
    For i = 1 To Rows
        tblTemp.Add Row(i).Duplicate
    Next
    
    Set Duplicate = tblTemp
    
End Function
