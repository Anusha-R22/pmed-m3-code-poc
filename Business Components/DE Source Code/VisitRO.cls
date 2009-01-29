VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VisitRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: VisitRO.cls
' Copyright InferMed Ltd 2001-2004 All Rights Reserved
' Author: Nicky Johns, InferMed, May 2001
' Purpose: Visit Read-only Object for MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 24-29 May 01 - Initial development
' NCJ 2 Jul 01 - IsInVisit now takes eForm rather than eFormID
' NCJ 3 Jul - 01 - Added Key property
' NCJ 9 Jul 01 - Fixed bug in getting VisitEForm
' NCJ 16 Aug 01 - VisitEFormByEForm
' NCJ 17 Sep 01 - Updated comments
' TA 25/04/2002 - New function to determine whether an eform is in a visit by the eFormTaskId
' MLM 04/09/02: Added new VisitDateVEF property so that it's easier to tell if a visit has a visit date
' NCJ 21 May 03 - Return correct value of MaxRepeats
' NCJ 24 Mar 04 - Ensure object collections are cleared on exit
'----------------------------------------------------

Option Explicit

' The state object containing its properties
Private moVisitState As VisitState

' The study to which this Visit belongs
Private moStudy As StudyDefRO

' Collection of VisitEForm objects
Private mcolVisitEForms As Collection

'MLM 04/09/02: A single VisitEForm to store the visit's visit eform
Private moVisitDateVEF As VisitEFormRO

'----------------------------------------------------
Friend Sub SetState(oStudy As StudyDefRO, oVState As VisitState)
'----------------------------------------------------
' Set the state and initialise ourselves
'----------------------------------------------------
Dim oVEFState As VisitEFormState
Dim oVEF As VisitEFormRO

    Set moVisitState = oVState
    Set moStudy = oStudy
    ' Now create the VisitEForm objects
    Set mcolVisitEForms = New Collection
    For Each oVEFState In moVisitState.VisitEForms
        Set oVEF = New VisitEFormRO
        Call oVEF.SetState(oStudy, Me, oVEFState)
        ' Add it to our collection, indexed by eFormID (NCJ 9/7/01)
        mcolVisitEForms.Add oVEF, Str(oVEFState.EFormId)
        'MLM 04/09/02: Keep a reference to the VEF if it's the visit date eform
        If oVEF.EFormUse = VisitEForm Then
            Set moVisitDateVEF = oVEF
        End If
    Next
    
    Set oVEFState = Nothing
    Set oVEF = Nothing
    
End Sub

'----------------------------------------------------
Public Property Get VisitEForms() As Collection
Attribute VisitEForms.VB_Description = "The Visit EForms (as collection of VisitEFormRO objects)"
'----------------------------------------------------
' The Visit EForms (as collection of VisitEFormRO objects)
'----------------------------------------------------

    Set VisitEForms = mcolVisitEForms

End Property

'----------------------------------------------------
Public Property Get VisitDateVEF() As VisitEFormRO
'----------------------------------------------------
' MLM 04/09/02: Added. Returns the VisitEForm of the visit's visit date eform.
'----------------------------------------------------

    Set VisitDateVEF = moVisitDateVEF
    
End Property

'-------------------------------------
Friend Property Get Key() As String
'-------------------------------------
' Get unique string key for this object
'-------------------------------------

    Key = moVisitState.Key

End Property

'-------------------------------------
Public Property Get Name() As String
Attribute Name.VB_Description = "The name of the visit"
'-------------------------------------
' The name of the visit
'-------------------------------------

    Name = moVisitState.Name
    
End Property

'-------------------------------------
Public Property Get Code() As String
Attribute Code.VB_Description = "The unique text code of the visit"
'-------------------------------------
' The unique text code of the visit
'-------------------------------------
    
    Code = moVisitState.Code

End Property

'-------------------------------------
Public Property Get VisitId() As Long
Attribute VisitId.VB_Description = "The unique numeric ID of the visit"
'-------------------------------------
' The unique numeric ID of the visit
'-------------------------------------

    VisitId = moVisitState.VisitId
    
End Property

'-------------------------------------
Public Property Get PromptForDate() As Boolean
Attribute PromptForDate.VB_Description = "Whether the user should be prompted for a date when the first eForm in the Visit is opened"
'-------------------------------------
' Whether the user should be prompted for a date
' when the first eForm in the Visit is opened
'-------------------------------------

    PromptForDate = moVisitState.PromptForDate
    
End Property

'-------------------------------------
Public Property Get DateLabelExpr() As String
Attribute DateLabelExpr.VB_Description = "Arezzo expression defining the visit date"
'-------------------------------------
' Arezzo expression defining the visit date
'-------------------------------------

    DateLabelExpr = moVisitState.DateLabelExpr
    
End Property

'-------------------------------------
Public Property Get MaxRepeats() As Integer
Attribute MaxRepeats.VB_Description = "The maximum number of repeats for this visit"
'-------------------------------------
' The maximum number of repeats for this visit
' May be -1 for indefinite cycling
' NCJ 21 May 03 - If it's NULL_INTEGER, convert to 1
'-------------------------------------

    If moVisitState.MaxRepeats = NULL_INTEGER Then
        MaxRepeats = 1
    Else
        MaxRepeats = moVisitState.MaxRepeats
    End If
    
End Property

'-------------------------------------
Public Property Get VisitOrder() As Integer
Attribute VisitOrder.VB_Description = "The position of this visit in the ordering of visits i.e. as on the top row of the schedule"
'-------------------------------------
' The position of this visit in the ordering of visits
' i.e. as on the top row of the schedule
'-------------------------------------

    VisitOrder = moVisitState.VisitOrder
    
End Property

'----------------------------------------------
Public Property Get BackgroundColour() As Long
Attribute BackgroundColour.VB_Description = "The background colour for the visit"
'----------------------------------------------
' The background colour for the visit
'----------------------------------------------

    BackgroundColour = moVisitState.BackgroundColour
    
End Property

'----------------------------------------------
Public Property Get IsInVisit(oEForm As eFormRO) As Boolean
Attribute IsInVisit.VB_Description = "Returns TRUE if this eForm appears in this visit, and FALSE otherwise"
'----------------------------------------------
' Returns TRUE if this eForm appears in this visit,
' and FALSE otherwise
'----------------------------------------------
    
    IsInVisit = CollectionMember(mcolVisitEForms, oEForm.Key)
    
End Property

'----------------------------------------------
Friend Property Get IsInVisitById(lEFormTaskId As Long) As Boolean
'----------------------------------------------
' 'TA 25/04/2002
' Returns TRUE if this eForm appears in this visit,
' and FALSE otherwise
'----------------------------------------------
 Dim oEForm As eFormRO
 
    Err.Clear
    On Error Resume Next
    Set oEForm = moStudy.eFormById(lEFormTaskId)
    If Err.Number = 5 Then
        Err.Clear
        On Error GoTo 0
        'eform does not exist
        IsInVisitById = False
    Else
        'eform exists in study - is it in this visit
        'reset error handling
        Err.Clear
        On Error GoTo 0
        IsInVisitById = IsInVisit(oEForm)
    End If
    
    
End Property
'-----------------------------------------------------------------
Public Function VisitEFormByEForm(oEForm As eFormRO) As VisitEFormRO
Attribute VisitEFormByEForm.VB_Description = "Get the VisitEForm object corresponding to this eForm"
'-----------------------------------------------------------------
' Get the VisitEForm object corresponding to this eForm
'-----------------------------------------------------------------

    Set VisitEFormByEForm = mcolVisitEForms.Item(oEForm.Key)

End Function

'----------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------

    Set moStudy = New StudyDefRO
    Set moVisitState = New VisitState
    Set mcolVisitEForms = New Collection
    
End Sub

'----------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------
' NCJ 24 Mar 04 - Ensure everything is tidied up
'----------------------------------------------

    Call Terminate

End Sub

'----------------------------------------------
Friend Sub Terminate()
'----------------------------------------------
' NCJ 25 Mar 04 - Ensure everything is tidied up
'----------------------------------------------
Dim oVEFRO As VisitEFormRO

    On Error GoTo ErrLabel
    
    Set moStudy = Nothing
    Set moVisitState = Nothing
    
    If Not moVisitDateVEF Is Nothing Then
        Call moVisitDateVEF.Terminate
    End If
    Set moVisitDateVEF = Nothing
    
    If Not mcolVisitEForms Is Nothing Then
        For Each oVEFRO In mcolVisitEForms
            Call oVEFRO.Terminate
        Next
        Call CollectionRemoveAll(mcolVisitEForms)
    End If
    Set mcolVisitEForms = Nothing
    Set oVEFRO = Nothing

    
Exit Sub
ErrLabel:
'    Debug.Print "Error in Terminating Visit"

End Sub

 
 
