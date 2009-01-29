Attribute VB_Name = "modDataUtils"
'----------------------------------------------------------------------------------------'
'   File:       modDataUtils.bas.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, November 2000
'   Purpose:    Utilitiy routines for the DataTable and DataRecord classes
'----------------------------------------------------------------------------------------'
' Revisions:

Option Explicit

'----------------------------------------------------------------------------------------'
Public Function RecordBuild(ParamArray vFields() As Variant) As clsDataRecord
'----------------------------------------------------------------------------------------'
'return a record from a list of parameter strings
'----------------------------------------------------------------------------------------'
Dim recRecord As clsDataRecord
Dim i As Long
    Set recRecord = New clsDataRecord
    
    recRecord.Init UBound(vFields) + 1
    
    For i = 0 To UBound(vFields)
        recRecord(i + 1) = vFields(i)
    Next

    Set RecordBuild = recRecord
    Set recRecord = Nothing
    
End Function


'----------------------------------------------------------------------------------------'
Public Function RecordBuildFromString(sFieldString As String, sSep As String) As clsDataRecord
'----------------------------------------------------------------------------------------'
' return a record based of the field string passed in
'----------------------------------------------------------------------------------------'
Dim vFields As Variant
Dim recRecord As clsDataRecord
Dim i As Long
    
    Set recRecord = New clsDataRecord
    vFields = Split(sFieldString, sSep)
    recRecord.Build Format(vFields(0))
    For i = 1 To UBound(vFields)
        recRecord.Add Format(vFields(i))
    Next
    
    Set RecordBuildFromString = recRecord
    Set recRecord = Nothing
    
End Function

'----------------------------------------------------------------------------------------'
Public Function RecordInit(lFields As Long) As clsDataRecord
'----------------------------------------------------------------------------------------'
' return a record of specified number of fields
'----------------------------------------------------------------------------------------'
 Dim recRecord As clsDataRecord
 
    Set recRecord = New clsDataRecord
    recRecord.Init lFields
    Set RecordInit = recRecord
    Set recRecord = Nothing
    
End Function

'----------------------------------------------------------------------------------------'
Public Function TableFromSQL(sSQL As String, Optional recHeadings As Variant = Empty, Optional bRemoveNull As Boolean = True) As clsDataTable
'----------------------------------------------------------------------------------------'
'Return a data table from SQL
'Input: optional RecHeadings - a data record of headings for each column,
'               if nothing is passed through the db column names will be used
'       optional RemoveNull - if true bull will be replaced by empty strings,
'               otherwise they will be replaced by "#NULL#"
'----------------------------------------------------------------------------------------'
Dim rs As ADODB.Recordset
Dim tblTable As clsDataTable
Dim recRow As clsDataRecord
Dim lCols As Long
Dim i As Long

    Set rs = New ADODB.Recordset
    rs.Open sSQL, MacroADODBConnection
    lCols = rs.Fields.Count
    Set tblTable = New clsDataTable
    
    
    If IsEmpty(recHeadings) Then
        Set tblTable.Headings = New clsDataRecord
        tblTable.Headings.Init lCols
        For i = 1 To lCols
            tblTable.Headings.Field(i) = rs.Fields(i - 1).Name
        Next
    Else
        Set tblTable.Headings = recHeadings
    End If
    
        
    
    Do While Not rs.EOF
        Set recRow = New clsDataRecord
        recRow.Init lCols
        For i = 1 To lCols
            recRow(i) = FieldToString(rs.Fields(i - 1).Value, bRemoveNull)
        Next
        tblTable.Add recRow
        Set recRow = Nothing
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing

    Set TableFromSQL = tblTable
End Function

'----------------------------------------------------------------------------------------'
Public Function FieldToString(vValue As Variant, Optional bRemoveNull As Boolean = True)
'----------------------------------------------------------------------------------------'
'converts a variant to "" or "#NULL#" according to bRemoveNull
'----------------------------------------------------------------------------------------'
    If VarType(vValue) = vbNull Then
        If bRemoveNull Then
            FieldToString = ""
        Else
            FieldToString = "#NULL#"
        End If
    Else
        FieldToString = Format(vValue)
    End If
    

End Function
