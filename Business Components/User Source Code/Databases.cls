VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Databases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       Databases.cls
'   Author:     Richard Meinesz, September 2002
'   Purpose:
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
'   Revisions:
'----------------------------------------------------------------------------------------'

Option Explicit

'----------------------------------------------------------------------------------------'
Friend Function UserDatabases(oSecCon As Connection, ByVal sUserName As String) As Collection
'----------------------------------------------------------------------------------------'
'REM 19/09/02
'Returns a collection of user databases from the UserDatabase table
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsDatabases As ADODB.Recordset
Dim colDatabases As Collection
Dim sDatabaseCode As String

    On Error GoTo Errorlabel

    'get all User databases
    sSQL = "SELECT DatabaseCode" _
         & " FROM UserDatabase" _
         & " WHERE UserName = '" & sUserName & "'"
    Set rsDatabases = New ADODB.Recordset
    rsDatabases.Open sSQL, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText

    Set colDatabases = New Collection
    
    'load the into a collection
    Do While Not rsDatabases.EOF
        sDatabaseCode = rsDatabases!DatabaseCode
        colDatabases.Add sDatabaseCode, sDatabaseCode
        rsDatabases.MoveNext
    Loop
    
    Set UserDatabases = colDatabases
    
    Set colDatabases = Nothing
    rsDatabases.Close
    Set rsDatabases = Nothing

Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Permissions.UserDatabases"
End Function

