Attribute VB_Name = "libSQL"
'----------------------------------------------------------------------------------------'
'   File:       libSQL.bas
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Toby Aldridge, April 2001
'   Purpose:    SQL Library
'----------------------------------------------------------------------------------------'
' Revisions:

'TA 27/9/2001: format removed from sql_valuetostringvalue
'TA 13/10/2001: if it is a number it needs to be converted to database standard
'----------------------------------------------------------------------------------------'

Option Explicit


Public Const SQL_SELECT = "select"
Public Const SQL_FROM = "from"
Public Const SQL_WHERE = "where"
Public Const SQL_LIKE = "like"
Public Const SQL_SUBSTR = "substr"
Public Const SQL_IS_NOT_NULL = "is not null"
Public Const SQL_IN = "in"
Public Const SQL_NULL = "null"

'----------------------------------------------------------------------------------------'
Public Function SQL_ValueToStringValue(ByRef vVariable As Variant) As Variant
'----------------------------------------------------------------------------------------'
' Return a string expression of a vlue for use in a SQL statement
'----------------------------------------------------------------------------------------'

    If IsNull(vVariable) Then
        SQL_ValueToStringValue = SQL_NULL
    Else
        Select Case VarType(vVariable)
        Case vbString: SQL_ValueToStringValue = SQL_StringToSQLString((vVariable))
        Case vbInteger, vbLong, vbDouble, vbSingle, vbCurrency, vbDecimal
            'TA 13/10/2001: if it is a number it needs to be converted to database standard
            SQL_ValueToStringValue = LocalNumToStandard(vVariable)
        Case Else
            SQL_ValueToStringValue = vVariable
        End Select
    End If
    
    
End Function


Public Function SQL_StringToSQLString(ByVal sValue As String) As String
'----------------------------------------------------------------------------------------'
' Put quotes around and double up single quotes for a string value
'----------------------------------------------------------------------------------------'

    SQL_StringToSQLString = "'" & Replace(sValue, "'", "''") & "'"

End Function
