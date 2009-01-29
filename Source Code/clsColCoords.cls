VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColCoords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       clsColCoords.cls
'   Author:     Nicky Johns, Feb 2000
'   Purpose:    Class to contain collection of clsCoords objects
'               (Used by Data Browser to store Locked and Frozen cells)
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 29/2/00 - Code extracted and generalised from frmDataItemResponse
'   NCJ 1/3/00 - Row, Col and FixedInteger are now Long
'----------------------------------------------------------------------------------------'

Option Explicit

Private mlFixedInteger As Long
Private mcolCoords As Collection
' For temporary use
Private moTempCoord As clsCoords

'----------------------------------------------------------------------------------------'
Public Property Get FixedInteger() As Long
'----------------------------------------------------------------------------------------'
' This is the fixed integer used when calculating items' keys
'----------------------------------------------------------------------------------------'
    
    FixedInteger = mlFixedInteger
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let FixedInteger(lFixed As Long)
'----------------------------------------------------------------------------------------'
' This is the fixed integer used when calculating items' keys
' (Assume FixedInteger is greater than any Col value)
'----------------------------------------------------------------------------------------'
    
    mlFixedInteger = lFixed
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get Item(lRow As Long, lCol As Long) As clsCoords
'----------------------------------------------------------------------------------------'
' Get item from collection
'----------------------------------------------------------------------------------------'

    Set Item = mcolCoords.Item(GetKey(lRow, lCol))

End Property

'----------------------------------------------------------------------------------------'
Public Function IsItem(lRow As Long, lCol As Long) As Boolean
'----------------------------------------------------------------------------------------'
' Return TRUE if item exists in collection
' Return FALSE is item does not exist
'----------------------------------------------------------------------------------------'

    On Error Resume Next
    Set moTempCoord = mcolCoords.Item(GetKey(lRow, lCol))
    IsItem = (Err.Number = 0)
    
End Function

'----------------------------------------------------------------------------------------'
Public Function AddItem(lRow As Long, lCol As Long) As clsCoords
'----------------------------------------------------------------------------------------'
' Add item to collection
'----------------------------------------------------------------------------------------'

    ' Create new Coordinates object
    Set moTempCoord = New clsCoords
    ' Set the row and column
    moTempCoord.Row = lRow
    moTempCoord.Col = lCol
    ' Add it to our collection
    mcolCoords.Add moTempCoord, GetKey(lRow, lCol)
    
    Set AddItem = moTempCoord
    
End Function

'----------------------------------------------------------------------------------------'
Public Sub Remove(lRow As Long, lCol As Long)
'----------------------------------------------------------------------------------------'
' Remove item from collection
'----------------------------------------------------------------------------------------'

    mcolCoords.Remove GetKey(lRow, lCol)

End Sub

'----------------------------------------------------------------------------------------'
Public Function Count() As Integer
'----------------------------------------------------------------------------------------'
' Count of items in collection
'----------------------------------------------------------------------------------------'

    Count = mcolCoords.Count
    
End Function

'----------------------------------------------------------------------------------------'
Private Function GetKey(lRow As Long, lCol As Long) As String
'----------------------------------------------------------------------------------------'
' Return key for item with this Row and Column
' Generate key using FixedInteger (assume lCol always less than FixedInteger)
'----------------------------------------------------------------------------------------'

    GetKey = lRow * mlFixedInteger + lCol & "K"
    
End Function

'----------------------------------------------------------------------------------------'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
' The enumerator to allow For...Each iterations
'----------------------------------------------------------------------------------------'
    Set NewEnum = mcolCoords.[_NewEnum]
End Function

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'

    Set mcolCoords = New Collection
    mlFixedInteger = 1
    Set moTempCoord = New clsCoords
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'

    Set mcolCoords = Nothing
    Set moTempCoord = Nothing
    
End Sub
