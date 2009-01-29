VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsMenuItems.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Toby Aldridge, October 2002
'   Purpose:    Class to represent a collection of clsMenuItems
'----------------------------------------------------------------------------------------'
' REVISIONS
'   NCJ 10 Oct 02 - Added KeyedItem and made Add return a clsMenuItem
'----------------------------------------------------------------------------------------'

Option Explicit

Private mcolMenuItems As Collection
Public DefaultItemIndex As Long

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Set mcolMenuItems = New Collection
    DefaultItemIndex = -1
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Set mcolMenuItems = Nothing
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function Count() As Long
'----------------------------------------------------------------------------------------'

    Count = mcolMenuItems.Count
    
End Function

'----------------------------------------------------------------------------------------'
Public Function KeyedItem(sKey As String) As clsMenuItem
'----------------------------------------------------------------------------------------'
'make our collection class 0 based
'----------------------------------------------------------------------------------------'

    Set KeyedItem = mcolMenuItems(sKey)

End Function

'----------------------------------------------------------------------------------------'
Public Function Item(lIndex As Long) As clsMenuItem
'----------------------------------------------------------------------------------------'
'make our collection class 0 based
'----------------------------------------------------------------------------------------'

    Set Item = mcolMenuItems(lIndex + 1)

End Function

'----------------------------------------------------------------------------------------'
Public Function Add(sKey As String, sCaption As String, Optional bEnabled As Boolean = False, _
                    Optional bChecked As Boolean = False, Optional bDefault As Boolean = False) As clsMenuItem
Attribute Add.VB_UserMemId = -4
'----------------------------------------------------------------------------------------'
' Create a new menu item
'----------------------------------------------------------------------------------------'
Dim oMI As clsMenuItem

    Set oMI = New clsMenuItem
    Call oMI.Init(sKey, sCaption, bEnabled, bChecked)
    mcolMenuItems.Add oMI, sKey
    If bDefault Then
        DefaultItemIndex = mcolMenuItems.Count - 1 'we are zero based
    End If
    Set Add = oMI
    Set oMI = Nothing

End Function

'----------------------------------------------------------------------------------------'
Public Function AddSeparator()
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
Dim oMI As clsMenuItem

    Set oMI = New clsMenuItem
    Call oMI.MakeSeparator
    mcolMenuItems.Add oMI
    Set oMI = Nothing


End Function
