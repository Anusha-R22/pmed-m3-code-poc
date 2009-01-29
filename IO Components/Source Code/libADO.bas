Attribute VB_Name = "libADO"
'----------------------------------------------------------------------------------------'
'   File:       modADO.bas
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Toby Aldridge, April 2001
'   Purpose:    General ADO functions
'----------------------------------------------------------------------------------------'
' Revisions:
'
'----------------------------------------------------------------------------------------'

Option Explicit

Public Const ADO_PB_RECORDSET As String = "RS"

'----------------------------------------------------------------------------------------'
Public Sub ADO_CloseConnection(oCon As ADODB.Connection)
'----------------------------------------------------------------------------------------'
'close connection if one exists and set to nothing
'----------------------------------------------------------------------------------------'

    If Not oCon Is Nothing Then
        If oCon.State = adStateOpen Then
            oCon.Close
        End If
        Set oCon = Nothing
    End If

End Sub
    

Public Function ADO_FieldType(ByVal intType As Integer) As String

   Select Case intType
      Case adVarChar: ADO_FieldType = "string"
      Case adSmallInt: ADO_FieldType = "integer"
      Case adInteger: ADO_FieldType = "long"
      Case adNumeric: ADO_FieldType = "long"
      Case adDouble: ADO_FieldType = "double"
      Case adUnsignedTinyInt: ADO_FieldType = "integer"
      Case adLongVarChar: ADO_FieldType = "String"
      Case Else: ADO_FieldType = intType
         
   End Select

End Function

Public Function ADO_FieldPrefix(ByVal intType As Integer) As String
      'DataTypeEnum
   Select Case intType
      Case adVarChar: ADO_FieldPrefix = "s"
      Case adSmallInt: ADO_FieldPrefix = "n"
      Case adInteger: ADO_FieldPrefix = "l"
      Case adDouble: ADO_FieldPrefix = "dbl"
      Case adUnsignedTinyInt: ADO_FieldPrefix = "n"
      Case adLongVarChar: ADO_FieldPrefix = "s"

      Case adNumeric: ADO_FieldPrefix = "l"
      Case Else: ADO_FieldPrefix = intType
         
   End Select

End Function

Public Function ADO_SerialiseRecordset(rs As Recordset) As String
'----------------------------------------------------------------------------------------'
' Serialise a collection
'----------------------------------------------------------------------------------------'
Dim pb As PropertyBag
    
    Set pb = New PropertyBag
    pb.WriteProperty ADO_PB_RECORDSET, rs
    ADO_SerialiseRecordset = pb.Contents
    Set pb = Nothing

End Function

Public Function ADO_DeSerialiseRecordset(ByVal sByteArray As String) As Recordset
'----------------------------------------------------------------------------------------'
' DeSerialise a previously serialised recordset
'----------------------------------------------------------------------------------------'
Dim pb As PropertyBag
Dim ByteArray() As Byte
 
    ByteArray = sByteArray
    Set pb = New PropertyBag
    pb.Contents = ByteArray
    Set ADO_DeSerialiseRecordset = pb.ReadProperty(ADO_PB_RECORDSET)
    Set pb = Nothing

End Function


Public Function ADO_HeadersFromRecordset(ByRef rs As Recordset) As String()
Dim v As Variant
Dim i As Long

    ReDim v(rs.Fields.Count - 1) As String
    For i = 0 To rs.Fields.Count - 1
        v(i) = rs.Fields(i).Name
    Next
    
    ADO_HeadersFromRecordset = v
    
    
End Function
