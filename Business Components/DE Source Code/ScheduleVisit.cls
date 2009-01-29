VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScheduleVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: ScheduleVisit.cls
' Copyright InferMed Ltd 2001-2004 All Rights Reserved
' Author: Nicky Johns, InferMed, May 2001
' ScheduleVisit object in MACRO 2.2
' This represents an actual Visit in the Schedule in DE
' It is created by the StudySubject object (not loaded from the DB)
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 23-29 May 01 - Initial development
' NCJ 17 Sep 01 - Updated comments; changed some props to Friend
' NCJ 17 Jan 03 - Added colEForms argument to Load because VisitEForms are not in the correct order
' NCJ 24/25 Mar 04 - Clear out object collections in Terminate
'----------------------------------------------------

Option Explicit

Private moVisit As VisitRO
Private moVisitInstance As VisitInstance
Private mcolVEFormInstances As Collection

'----------------------------------------------------
Friend Sub Load(oVisit As VisitRO, oVisitInstance As VisitInstance, _
                colEForms As Collection, colEFormIdsToHide As Collection)
Attribute Load.VB_Description = "Create ourselves by reading the Study Def schedule and building the VisitEFormInstance objects oVisitInstance may be Nothing colEFormsToHide is collection of eForms which should not appear on the schedule"
'----------------------------------------------------
' Create ourselves by reading the Study Def schedule
' and building the VisitEFormInstance objects
' oVisitInstance may be Nothing
' colEFormsToHide is collection of eForms which
' should not appear on the schedule
' NCJ 17 Jan 03 - Added colEForms argument (all eForms) because VisitEForms are not in the correct order
'----------------------------------------------------
Dim oEFI As EFormInstance
Dim oVEForm As VisitEFormRO
Dim oVEFInstance As VEFInstance
Dim oEFIs As Collection     ' of eForm instances
Dim i As Long
Dim nCountEFIs As Long
Dim oEForm As eFormRO

    Set moVisit = oVisit
    Set moVisitInstance = oVisitInstance
    
    ' Walk through the Study Def schedule for this visit
    ' (to preserve the ordering)
    ' NCJ 17 Jan 03 - oVisit.VisitEForms are NOT in the correct order, so go by eForms collection
'    For Each oVEForm In oVisit.VisitEForms
    For Each oEForm In colEForms
        ' See if it's in the visit (ignore it if not)
        If oVisit.IsInVisit(oEForm) Then
            Set oVEForm = oVisit.VisitEFormByEForm(oEForm)
            
            ' Get all the instances of this eForm in the visit (if any)
            If Not oVisitInstance Is Nothing Then
                Set oEFIs = oVisitInstance.eFormInstancesByEForm(oEForm)
                nCountEFIs = oEFIs.Count
            Else
                ' We don't have a Visit Instance
                nCountEFIs = 0
            End If
            If nCountEFIs > 0 Then
                ' There is at least one eForm instance for this eform
                ' The collection is indexed by cycle number
                For i = 1 To nCountEFIs
                    Set oEFI = oEFIs.Item(Str(i))
                    ' For each eForm instance create a Visit eForm instance
                    Set oVEFInstance = New VEFInstance
                    Set oVEFInstance.VisitEForm = oVEForm   ' Set the Visit eForm
                    Set oVEFInstance.EFormInstance = oEFI   ' Set the eForm instance
                    mcolVEFormInstances.Add oVEFInstance, oVEFInstance.Key
                Next i
            ElseIf CollectionMember(colEFormIdsToHide, oEForm.EFormId) Then
                ' Do nothing because we want to hide this one
            Else
                ' There are no eForm instances for this eform
                ' so create just one VisitEFormInstance
                Set oVEFInstance = New VEFInstance
                Set oVEFInstance.VisitEForm = oVEForm   ' Set the Visit eForm
                mcolVEFormInstances.Add oVEFInstance, oVEFInstance.Key
            End If
        End If
    Next
    
    Set oEFI = Nothing
    Set oVEForm = Nothing
    Set oVEFInstance = Nothing
    Set oEFIs = Nothing

End Sub

'----------------------------------------------------
Public Property Get VisitInstance() As VisitInstance
Attribute VisitInstance.VB_Description = "The Visit instance object"
'----------------------------------------------------
' The Visit instance object
'----------------------------------------------------

    Set VisitInstance = moVisitInstance
    
End Property

'----------------------------------------------------
Public Property Get Visit() As VisitRO
Attribute Visit.VB_Description = "The Visit definition object"
'----------------------------------------------------
' The Visit definition object
'----------------------------------------------------

    Set Visit = moVisit
    
End Property

'----------------------------------------------------
Friend Property Get Key() As String
'----------------------------------------------------
' Get the key of this objct (for unique identification)
'----------------------------------------------------
Dim sInstKey As String

    If moVisitInstance Is Nothing Then
        sInstKey = 0
    Else
        sInstKey = moVisitInstance.VisitTaskId
    End If
    Key = ScheduleVisitKey(moVisit.VisitId, sInstKey)
    
End Property

'----------------------------------------------------
Public Property Get VisitEFormInstances() As Collection
Attribute VisitEFormInstances.VB_Description = "The collection of VisitEForm instance objects"
'----------------------------------------------------
' The collection of VisitEForm instance objects
'----------------------------------------------------

    Set VisitEFormInstances = mcolVEFormInstances
    
End Property

'----------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------
' Initialise all our objects
'----------------------------------------------------

    Set moVisit = New VisitRO
    Set moVisitInstance = Nothing   ' Must initialise to Nothing
    Set mcolVEFormInstances = New Collection
    
End Sub

'----------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------
' Set all our objects to Nothing
' NCJ 24 Mar 04 - And clear object collection too
'----------------------------------------------------

    Call Terminate

End Sub

'----------------------------------------------------
Friend Sub Terminate()
'----------------------------------------------------
' NCJ 25 Mar 04
'----------------------------------------------------
Dim oVEFI As VEFInstance

    Set moVisit = Nothing
    Set moVisitInstance = Nothing
    
    If Not mcolVEFormInstances Is Nothing Then
        For Each oVEFI In mcolVEFormInstances
            Call oVEFI.Terminate
        Next
        Call CollectionRemoveAll(mcolVEFormInstances)   ' NCJ 24 Mar 04
    End If
    Set mcolVEFormInstances = Nothing
    Set oVEFI = Nothing
    
End Sub
 
 
