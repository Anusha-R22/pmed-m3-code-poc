VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OramaDec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       OramaDec.cls
'   Copyright:  InferMed Ltd. 2004. All Rights Reserved
'   Author:     Nicky Johns, October 2004
'   Purpose:    Class to store decision info for frmORAMA in Windows DE
'----------------------------------------------------------------------------------------'
' NB This class is only ever used if MACRO_DM is compiled with ORAMA = 1
'----------------------------------------------------------------------------------------'
' Revisions:
'
' NCJ 27 Oct 04 - Initial version
' NCJ 2 Nov 04 - Added "Committed" proprerty
'----------------------------------------------------------------------------------------'

Option Explicit

' The decision
Public DecisionTask As TaskInstance
Public IsCommitted As Boolean


' The list of option/check indexes
Private mcolIndexes As Collection

'----------------------------------------------------------------------------------------'
Public Property Get IndexCollection() As Collection
'----------------------------------------------------------------------------------------'
' The indexes of the options/checks belonging to this decision
'----------------------------------------------------------------------------------------'
    
    Set IndexCollection = mcolIndexes

End Property

'----------------------------------------------------------------------------------------'
Public Sub AddIndex(ByVal nIndex As Integer)
'----------------------------------------------------------------------------------------'
' Add an index to our collection
'----------------------------------------------------------------------------------------'

    mcolIndexes.Add nIndex

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'

    Set mcolIndexes = New Collection
    Me.IsCommitted = False
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'

    Call Terminate
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Terminate()
'----------------------------------------------------------------------------------------'

    Set mcolIndexes = Nothing
    Set DecisionTask = Nothing
    
End Sub

