VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VEFInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: VEFInstance.cls
' Nicky Johns, InferMed, May 2001
' VisitEFormInstance object in MACRO 2.2
' This represents a cell on the Schedule grid in DE
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 23-29 May 01 - Initial development
' NCJ 3 Jul 01 - Added Key property; added private versions of variables
'               and Property Get/Lets
' NCJ 24/25 Mar 04 - Added class Terminate
'----------------------------------------------------

Option Explicit

' The object's Visit eForm definition
Private moVEForm As VisitEFormRO

' The eForm Instance it represents
Private moEFI As EFormInstance

'----------------------------------------------------
Public Property Get VisitEForm() As VisitEFormRO
Attribute VisitEForm.VB_Description = "The VisitEForm definition for this instance"
'----------------------------------------------------
' The VisitEForm definition for this instance
'----------------------------------------------------

    Set VisitEForm = moVEForm
    
End Property

'----------------------------------------------------
Friend Property Set VisitEForm(oVEF As VisitEFormRO)
'----------------------------------------------------
' The VisitEForm definition for this instance
'----------------------------------------------------

    Set moVEForm = oVEF

End Property

'----------------------------------------------------
Public Property Get EFormInstance() As EFormInstance
Attribute EFormInstance.VB_Description = "The eFormInstance object"
'----------------------------------------------------
' The eFormInstance object
'----------------------------------------------------

    Set EFormInstance = moEFI
    
End Property

'----------------------------------------------------
Friend Property Set EFormInstance(oEFI As EFormInstance)
'----------------------------------------------------
' The eFormInstance object
'----------------------------------------------------

    Set moEFI = oEFI

End Property

'----------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------

    Set moVEForm = New VisitEFormRO
    Set moEFI = Nothing     ' Must initialise to Nothing

End Sub

'-------------------------------------
Friend Property Get Key() As String
'-------------------------------------
' Get unique string key for this object
'-------------------------------------

    If Not moEFI Is Nothing Then
        Key = VisitEFormInstanceKey(moVEForm.eForm.EFormId, moEFI.eFormTaskId)
    Else
        Key = VisitEFormInstanceKey(moVEForm.eForm.EFormId, 0)
    End If
    
End Property
 
'----------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------
' NCJ 24 Mar 04
'----------------------------------------------------
    
    Call Terminate

End Sub

'----------------------------------------------------
Friend Sub Terminate()
'----------------------------------------------------
' NCJ 25 Mar 04
'----------------------------------------------------
    
    Set moVEForm = Nothing
    Set moEFI = Nothing

End Sub
