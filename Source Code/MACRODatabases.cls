VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACRODatabases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       clsMACRODatabases.cls for MTM 2.0
'   Author:
'   Purpose:    MACRO Database object
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:

Option Explicit
Option Base 0
Option Compare Binary

Private mcolDatabase As Collection

Private Sub Class_Initialize()
    Set mcolDatabase = New Collection
End Sub

Public Function Count() As Long
    Count = mcolDatabase.Count
End Function

' The enumerator to allow For...Each iterations
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolDatabase.[_NewEnum]
End Function


Friend Sub Clear()

    Set mcolDatabase = Nothing     ' Throw away old collection
    Set mcolDatabase = New Collection

End Sub

Friend Sub AddUserDatabase(vUserDatabase As String)

    mcolDatabase.Add vUserDatabase

End Sub
