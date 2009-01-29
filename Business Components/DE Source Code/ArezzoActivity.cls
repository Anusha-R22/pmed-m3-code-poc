VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArezzoActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------
' File: ArezzoActivity.cls
' Author: Nicky Johns
' Copyright: InferMed, January 2003, All Rights Reserved
' Purpose: Stores activity on non-MACRO AREZZO tasks for MACRO 3.0
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
' NCJ 31 Jan 03 - Initial development
'----------------------------------------------------

Option Explicit

Public Enum eArezzoActivity
    aaConfirmAction = 1
    aaAddData = 2
    aaCommitDecision = 3
End Enum

' Type of activity
Public ActivityType As eArezzoActivity

' Only relevant for ConfirmAction and CommitDecision
Public TaskKey As String

' DataName for AddData or CandName for CommitDecision
Public ItemName As String

' Data value for AddData
Public Value As String

'----------------------------------------------------
Public Sub DoActivity(oALM As ALM5)
'----------------------------------------------------
' Do this activity in the given ALM
' NB Check we can still do it, in case another user got in first (in Multi-User mode)
'----------------------------------------------------
Dim oTask As TaskInstance

    ' We only have a task key for Actions and Decisions (not for adding data)
    If TaskKey > "" Then
        Set oTask = oALM.GuidelineInstance.colTaskInstances.Item(TaskKey)
    End If
    
    Select Case ActivityType
    Case eArezzoActivity.aaAddData
        ' Add data (check it's still requested)
        If oALM.GuidelineInstance.IsRequested(Me.ItemName) Then
            Call oALM.GuidelineInstance.colDataValues.Add(Me.ItemName, Me.Value, True)
        End If
        
    Case eArezzoActivity.aaConfirmAction
        ' Confirm an action (check it's still requested)
        If oTask.TaskState = "requested" Then
            Call oTask.Confirm
        End If
    
    Case eArezzoActivity.aaCommitDecision
        ' Commit a decision (check it's permitted)
        If oTask.TaskState = "permitted" Then
            Call oTask.Commit(Me.ItemName)
        End If
    End Select

End Sub

'----------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------
'----------------------------------------------------

    TaskKey = ""
    ItemName = ""
    Value = ""

End Sub
