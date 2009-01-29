VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VisitEFormRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: VisitEFormRO.cls
' Copyright InferMed Ltd 2001 - 2004
' Author: Nicky Johns, InferMed, May 2001
' Visit eForm Read-only Object for MACRO 3.0
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 24-29 May 01 - Initial development
' MLM 28/08/02: Added new EFormUse property and eEFormUse enumeration
' MLM 29/08/02: In SetState, set the parent eForm's new AppearsInSchedule property
' NCJ 24 Mar 04 - Ensure all objects are tidied up on Termination
'----------------------------------------------------

Option Explicit

'MLM 28/08/02:
Public Enum eEFormUse
    User = 0
    VisitEForm = 1
End Enum

' The state object containing its properties
Private moVisitEFormState As VisitEFormState

' The study to which this VisitEForm belongs
' (Do we need this?) - No, we don't! (NCJ 24 Mar 04)
'Private moStudy As StudyDefRO

Private moVisit As VisitRO
Private moEForm As eFormRO

'----------------------------------------------------
Friend Sub SetState(oStudy As StudyDefRO, oVisit As VisitRO, _
                        oVState As VisitEFormState)
'----------------------------------------------------
' Set the state and initialise ourselves
'
' MLM 29/08/02: Initialising the VisitEForm notifies its EForm if it needs to be shown on the schedule.
'----------------------------------------------------

    Set moVisitEFormState = oVState
    ' NCJ 24 Mar 04 - Don't need moStudy
'    Set moStudy = oStudy
    Set moVisit = oVisit
    Set moEForm = oStudy.eFormById(moVisitEFormState.EFormId)
    
    If EFormUse = User Then
        moEForm.AppearsInSchedule = True
    End If
    
End Sub

'----------------------------------------------------
Public Property Get eForm() As eFormRO
Attribute eForm.VB_Description = "This VisitEForm 's eForm object."
'----------------------------------------------------
'This VisitEForm 's eForm object.
'----------------------------------------------------

    Set eForm = moEForm
    
End Property

'----------------------------------------------------
Public Property Get Visit() As VisitRO
Attribute Visit.VB_Description = "This VisitEform's Visit object."
'----------------------------------------------------
' This VisitEform's Visit object.
'----------------------------------------------------

    Set Visit = moVisit
    
End Property

'----------------------------------------------------
Public Property Get MaxRepeats() As Integer
Attribute MaxRepeats.VB_Description = "Not yet used."
'----------------------------------------------------
'Not yet used.
'----------------------------------------------------

    MaxRepeats = moVisitEFormState.MaxRepeats

End Property

'----------------------------------------------------
Public Property Get Repeating() As Boolean
Attribute Repeating.VB_Description = "Is this VisitEForm repeating?"
'----------------------------------------------------
'Is this VisitEForm repeating?
'----------------------------------------------------

    Repeating = moVisitEFormState.Repeating

End Property

'----------------------------------------------------
Public Property Get EFormUse() As eEFormUse
'----------------------------------------------------
'MLM 28/08/02: Added. Describes whether the eForm is a User or VisitEForm form.
'----------------------------------------------------

    EFormUse = moVisitEFormState.EFormUse

End Property

'----------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------

'    Set moStudy = New StudyDefRO
    Set moVisitEFormState = New VisitEFormState
    
End Sub

'----------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------

    Call Terminate

End Sub
 
'----------------------------------------------
Friend Sub Terminate()
'----------------------------------------------
' NCJ 25 Mar 04 - Ensure everything is tidied up
'----------------------------------------------

    Set moVisitEFormState = Nothing
    Set moVisit = Nothing
    Set moEForm = Nothing

End Sub
 
 
 
