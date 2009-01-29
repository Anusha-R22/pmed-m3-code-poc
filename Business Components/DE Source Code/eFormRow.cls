VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eFormRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Eform Row object for Macro 2.2 ScheduleGrid"
Attribute VB_Description = "Eform Row object for Macro 2.2 ScheduleGrid"
Attribute VB_Description = "Eform Row object for Macro 2.2 ScheduleGrid"
Attribute VB_Description = "Eform Row object for Macro 2.2 ScheduleGrid"
'----------------------------------------------------------------------------------------'
'   File:       eFormRow.cls
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Ash Trebi-Ollennu, June 2001
'   Purpose:    Eform Row object for Macro 2.2 ScheduleGrid
'----------------------------------------------------------------------------------------'
' Revisions:
'   AT-O 29 June 01 - Initial Development
'   NCJ 3 Jul 01 - Further work
'----------------------------------------------------------------------------------------'

Option Explicit

' The eForm definition
Private moEForm As eFormRO
' The no. of grid rows for this eForm
Private mnDepth As Integer

'-------------------------------------------------
Public Property Get eForm() As eFormRO
Attribute eForm.VB_Description = ""
'-------------------------------------------------
    
    Set eForm = moEForm

End Property

'---------------------------------------------------
Public Property Set eForm(ByVal oNewEForm As eFormRO)
'---------------------------------------------------
    
    Set moEForm = oNewEForm

End Property

'----------------------------------------------------
Public Property Let Depth(ByVal nNewDepth As Integer)
'----------------------------------------------------
     
     mnDepth = nNewDepth

End Property

'----------------------------------------------------
Public Property Get Depth() As Integer
Attribute Depth.VB_Description = ""
'----------------------------------------------------
     
     Depth = mnDepth

End Property


 
 
 
 
