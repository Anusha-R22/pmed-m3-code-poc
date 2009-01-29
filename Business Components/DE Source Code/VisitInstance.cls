VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VisitInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: VisitInstance.cls
' Copyright InferMed Ltd 2001-2004 All Rights Reserved
' Author: Nicky Johns, InferMed, May 2001
' Purpose: VisitInstance object in MACRO 2.2
'       This represents an actual visit in the Schedule in DE
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 24-29 May 01 - Initial development
' NCJ 2 Jul 01 - Changed eFormID input parameters to eForm
' NCJ 3 Jul 01 - Added Key property
' NCJ 10 Jul 01 - Use IsChanged property of State object
' NCJ 11 Jul 01 - Recognise 0 for undefined dates
' NCJ 31/7/01 - Use new Replace Data for Visit date
' NCJ 16/8/01 - Added UpdateStatus
' NCJ 17 Sep 01 - Updated comments; made some properties Friendly
' TA 14 Aug 02 - New statuses
' MLM 29/08/02: Added a new VisitEFormInstance property, intended to replace some of the visit date stuff.
' MLM 10/09/02: New style visit date handling in RefreshVisitDate.
' MLM 16/09/02: VisitDateString formats visit date according to the date q in SD.
' NCJ 19 Sept 02 - Fixed bug in VisitDateString
' NCJ 20 Sept 02 - Removed unused context setting for AREZZO activities
' NCJ 1 Apr 03 - Changed RefreshVisitDate
' NCJ 24/25 Mar 04 - Ensure all objects dealt with in Terminate
' NCJ 5 Jan 06 - VisitDateString now copes with Partial Dates
' NCJ 17 Jan 06 - Temp fix for Web DE - ignore partial visit dates!
' NCJ 23 Jan 06 - Partial dates back in VisitDateString
' NCJ 15 Aug 06 - Do NOT terminate the VisitEForm in Terminate (Bug 2778)
'----------------------------------------------------

Option Explicit

' The study def to which this instance belongs
Private moStudy As StudyDefRO

' The subject to which this visit instance belongs
Private moSubject As StudySubject

' The state class which holds all the properties
Private moVIState As VisitInstanceState

' The Visit definition class
Private moVisit As VisitRO

' The collection of eForm Instances belonging to this visit
Private mcolEFIs As Collection

' The collection of eFormIds which have instances in this VisitInstance
Private mcolEFormIds As Collection

'MLM 29/08/02:
Private moVisitEFormInstance As EFormInstance

' The Arezzo instance we're using
Private moArezzo As Arezzo_DM

'-------------------------------------
Public Property Get Subject() As StudySubject
Attribute Subject.VB_Description = "The Study SUbject to which this visit instance belongs"
'-------------------------------------
' The Study SUbject to which this visit instance belongs
'-------------------------------------

    Set Subject = moSubject

End Property

'-------------------------------------
Friend Sub SetState(oStudy As StudyDefRO, oSubject As StudySubject, oVIState As VisitInstanceState, _
                    oArezzo As Arezzo_DM)
'-------------------------------------
' Initialise ourselves with the given state
'-------------------------------------

    ' Store what's been passed in to us
    Set moStudy = oStudy
    Set moSubject = oSubject
    Set moVIState = oVIState
    Set moArezzo = oArezzo
    
    Set moVisit = moStudy.VisitById(moVIState.VisitId)
    ' Create our collection of eForm instances
    ' (these get filled in later)
    Set mcolEFIs = New Collection
    
End Sub

'-------------------------------------
Private Sub AddEFormId(lEFormId As Long)
'-------------------------------------
' Add to our "eForm Ids we've got" collection
' Ignore errors if we've already got this eFormID
'-------------------------------------
        
    CollectionAddAnyway mcolEFormIds, True, Str(lEFormId)

End Sub


' NCJ 20 Sept 02 - We don;t do any AREZZO calcuations within the VisitInstance any more
' so we don't need this SetContext routine
''------------------------------------------------
'Friend Sub SetContext()
''------------------------------------------------
'' Set the Arezzo context for future calculations
''------------------------------------------------
'
'    Call moArezzo.SetMACROContext(Me, Nothing)
'
'End Sub

'-------------------------------------------------
Friend Sub AddEFormInstance(oEFI As EFormInstance)
'-------------------------------------------------
' Add a new eFormInstance to this Visit Instance
'-------------------------------------------------

    mcolEFIs.Add oEFI, oEFI.Key
    ' Include it in our "eForms we've got" collection
    AddEFormId oEFI.eForm.EFormId
    
End Sub

'-------------------------------------------------
Public Property Get Arezzo() As Arezzo_DM
Attribute Arezzo.VB_Description = "The Arezzo instance being used for this subject"
'-------------------------------------------------
' The Arezzo instance being used for this subject
'-------------------------------------------------

    Set Arezzo = moArezzo
    
End Property

'-------------------------------------
Friend Property Get Key() As String
'-------------------------------------
' Get unique string key for this object
'-------------------------------------

    Key = moVIState.Key

End Property

'-------------------------------------
Public Property Get Visit() As VisitRO
Attribute Visit.VB_Description = "The visit definition object for this instance"
'-------------------------------------
' The visit definition object for this instance
'-------------------------------------

    Set Visit = moVisit

End Property

'---------------------------------------------------
Public Property Get eFormInstances() As Collection
Attribute eFormInstances.VB_Description = "The collection of eForm Instance objects belonging to this visit"
'---------------------------------------------------
' The collection of eForm Instance objects belonging to this visit
'---------------------------------------------------

    Set eFormInstances = mcolEFIs

End Property

'---------------------------------------------------
Public Property Get eFormInstancesByEForm(oEForm As eFormRO) As Collection
Attribute eFormInstancesByEForm.VB_Description = "The eFormInstances corresponding to this specific eForm. Collection Key is eform's cycle no. Returns empty collection if none exist."
'---------------------------------------------------
' The eFormInstances corresponding to this specific eForm.
' Collection Key is eform's cycle no.
' Returns empty collection if none exist.
'---------------------------------------------------
Dim colEFIs As Collection
Dim i As Long

    Set colEFIs = New Collection
    ' Loop through our collection of eForm instances
    For i = 1 To mcolEFIs.Count
        If mcolEFIs(i).eForm.EFormId = oEForm.EFormId Then
            ' Use cycle no. as key
            colEFIs.Add mcolEFIs(i), Str(mcolEFIs(i).CycleNo)
        End If
    Next
    Set eFormInstancesByEForm = colEFIs
    
    Set colEFIs = Nothing

End Property

'---------------------------------------------------
Public Property Get eFormInstanceByEFormCycle(oEForm As eFormRO, _
                        nCycle As Integer) As EFormInstance
Attribute eFormInstanceByEFormCycle.VB_Description = "Get eFormInstance for specific eForm and Cycle no. Returns Nothing if none exists"
'---------------------------------------------------
' Get eFormInstance for specific eForm and Cycle no.
' Returns Nothing if none exists
'---------------------------------------------------
Dim oEFI As EFormInstance

    For Each oEFI In mcolEFIs
        If oEFI.eForm.EFormId = oEForm.EFormId And oEFI.CycleNo = nCycle Then
            Set eFormInstanceByEFormCycle = oEFI
            Exit For
        End If
    Next
    Set oEFI = Nothing

End Property

'---------------------------------------------------
Public Property Get eFormInstanceByTaskId(lEFITaskId As Long) As EFormInstance
Attribute eFormInstanceByTaskId.VB_Description = "Get eFormInstance with specific form instance task Id. Returns Nothing if none exists"
'---------------------------------------------------
' Get eFormInstance with specific form instance task Id.
' Returns Nothing if none exists
'---------------------------------------------------
Dim oEFI As EFormInstance
Dim i As Long

    Set eFormInstanceByTaskId = mcolEFIs(Str(lEFITaskId))

End Property

'-------------------------------------
Public Property Get Code() As String
Attribute Code.VB_Description = "The unique text code for the visit"
'-------------------------------------
' The unique text code for the visit
'-------------------------------------

    Code = moVisit.Code

End Property

'-------------------------------------
Public Property Get Name() As String
Attribute Name.VB_Description = "The name of the visit"
'-------------------------------------
' The name of the visit
'-------------------------------------

    Name = moVisit.Name

End Property

'-------------------------------------
Public Property Get VisitId() As Long
Attribute VisitId.VB_Description = "The unique numeric ID for the visit definition"
'-------------------------------------
' The unique numeric ID for the visit definition
'-------------------------------------

    VisitId = moVisit.VisitId

End Property

'-------------------------------------
Public Property Get VisitTaskId() As Long
Attribute VisitTaskId.VB_Description = "The unique numeric ID for this visit instance"
'-------------------------------------
' The unique numeric ID for this visit instance
'-------------------------------------

    VisitTaskId = moVIState.VisitTaskId

End Property

'-------------------------------------
Public Property Get CycleNo() As Integer
Attribute CycleNo.VB_Description = "The cycle number of this visit instance. Will be 1 for non-cycling visits."
'-------------------------------------
' The cycle number of this visit instance.
' Will be 1 for non-cycling visits.
'-------------------------------------

    CycleNo = moVIState.CycleNo

End Property

'-------------------------------------
Public Property Get Status() As Integer
Attribute Status.VB_Description = "The calculated status of the instance (based on the statuses of its eForm instances). See also StatusString."
'-------------------------------------
' The calculated status of the instance
' (based on the statuses of its eForm instances).
' See also StatusString.
'-------------------------------------

    Status = moVIState.Status

End Property

'-------------------------------------
Public Property Get LockStatus() As Integer
Attribute LockStatus.VB_Description = "The lock status of the instance"
'-------------------------------------
' The lock status of the instance
'-------------------------------------

    LockStatus = moVIState.LockStatus

End Property

'----------------------------------------------------
Public Property Get StatusString() As String
Attribute StatusString.VB_Description = "Get instance status as a string (for display to user)"
'----------------------------------------------------
' Get instance status as a string (for display to user)
'----------------------------------------------------

    StatusString = GetStatusString(moVIState.Status)

End Property

'-------------------------------------
Public Property Get VisitDate() As Double
Attribute VisitDate.VB_Description = "The visit date. May be 0 for ""unspecified"". May only be set if ""prompt for date"" is true."
'-------------------------------------
' The visit date. May be 0 for "unspecified".
' This is set from the Visit eForm date.
'-------------------------------------

    VisitDate = moVIState.VisitDate

End Property

'-----------------------------------------------------------
Friend Sub RefreshVisitDate(ByVal dblDate As Double, ByVal sArezzoDate As String)
Attribute RefreshVisitDate.VB_Description = "Recalculate the visit date and return result. May be 0 if unspecified"
'-----------------------------------------------------------
' Store the visit date in AREZZO as the special Visit:date term
' May be 0 if unspecified
'-----------------------------------------------------------
' MLM 10/09/02: Use new form and visit date system
' NCJ 1 Apr 03 - Changed to be called whenever the visit date changes (from Response.cls)
' Note: this assumes there IS a visit eform instance with a visit date
'-----------------------------------------------------------
Dim sVisitDateName As String
Dim sUserTag As String

    If dblDate <> moVIState.VisitDate Then
        ' It's changed
        moVIState.VisitDate = dblDate
        moVIState.IsChanged = True
        sUserTag = CStr(Me.VisitTaskId * 100000)
        If dblDate <> 0 Then
            sVisitDateName = moVisit.Code & ":date"
            Call moArezzo.ReplaceQData(sVisitDateName, sArezzoDate, sUserTag)
        Else
            ' Just remove previous value
            Call moArezzo.RemoveDataByRef(sUserTag)
        End If
    End If
    
End Sub

'-----------------------------------------------------------
Public Property Get VisitDateString() As String
Attribute VisitDateString.VB_Description = "The date formatted according to study's default date format. Will be empty string if date is undefined"
'-----------------------------------------------------------
' Will be empty string if date is undefined
'
' MLM 16/09/02: Use the visit date's question format.
' NCJ/MLM 19 Sept 02 - Load VisitEForm's elements, and made sure we check for moVisit.VisitDateVEF being Nothing
' NCJ 21 Jan 03 - Use local date format if set
' NCJ 5 Jan 06 - Handle Partial Dates
' NCJ 17 Jan 06 - Temp fix for Web DE - ignore partial dates!
' NCJ 23 Jan 06 - Partial dates back in
'-----------------------------------------------------------

Dim oVisitEForm As VisitEFormRO
Dim oVisitDateElement As eFormElementRO
Dim sFormat As String

' NCJ 17 Jan 05 - Temporarily look for partial dates
' Const MAX_YEAR = 2695

    ' Start by guessing that no visit date need be displayed
    VisitDateString = ""
    
    ' If no visit date, nothing to do
    If moVIState.VisitDate = 0 Then Exit Property
        
    ' A visit date has been collected
    Set oVisitEForm = moVisit.VisitDateVEF
    If Not oVisitEForm Is Nothing Then
        ' There is a visit eform
        ' Check its elements are loaded
        Call moStudy.LoadElements(oVisitEForm.eForm)
        Set oVisitDateElement = oVisitEForm.eForm.EFormDateElement
        If Not oVisitDateElement Is Nothing Then
            'the visit eform has a date q, so use its format
            ' NCJ 21 Jan 03 - Use local date format if any
            sFormat = moStudy.Subject.LocalDateTypeFormat(oVisitDateElement.DateTimeType)
            If sFormat = "" Then
                ' Use element's own format
                sFormat = oVisitDateElement.VBFormat
            End If
            VisitDateString = VBFormatPartialDate(moVIState.VisitDate, sFormat)
'            ' NCJ 17 Jan 06 - Temporary check for partial dates
'            If DatePart("yyyy", CDate(moVIState.VisitDate)) < MAX_YEAR Then
'                ' Use VB formatting (can handle partial sFormat)
'                VisitDateString = Format(moVIState.VisitDate, sFormat)
'            Else
'                ' NCJ 5 Jan 06 - Go through AREZZO in case of partial dates
'                ' But ignore errors if no AREZZO (e.g. in Web Schedule)
'                On Error Resume Next
'                VisitDateString = moArezzo.FormatDoubleDate(moVIState.VisitDate, sFormat)
'            End If
        End If
    End If
    ' tidy up, in case these were set
    Set oVisitDateElement = Nothing
    Set oVisitEForm = Nothing

End Property

'-------------------------------------
Public Function IsAnInstance(oEForm As eFormRO) As Boolean
Attribute IsAnInstance.VB_Description = "Returns TRUE if this eForm appears as an instance in this visit, and FALSE otherwise"
'-------------------------------------
' Returns TRUE if this eForm appears as an instance in this visit,
' and FALSE otherwise
'----------------------------------------------

    IsAnInstance = CollectionMember(mcolEFormIds, oEForm.Key, False)

End Function

'-------------------------------------
Public Property Get Changed() As Boolean
Attribute Changed.VB_Description = "Whether VisitInstance has changed (or is new)"
'-------------------------------------
' Whether VisitInstance has changed (or is new)
'-------------------------------------

    Changed = (moVIState.IsChanged Or moVIState.IsNew)

End Property

'-------------------------------------
Private Sub Class_Initialize()
'-------------------------------------

    Set moStudy = New StudyDefRO
    Set moVisit = New VisitRO
    Set mcolEFIs = New Collection
    Set mcolEFormIds = New Collection
    
End Sub

'-------------------------------------
Private Sub Class_Terminate()
'-------------------------------------

    Call Terminate
    
End Sub

'------------------------------------------------
Friend Sub Terminate()
'------------------------------------------------
' Tidy up everything
'------------------------------------------------

    Set moSubject = Nothing
    Set moStudy = Nothing
    Set moVisit = Nothing
    Set moVIState = Nothing
    
    ' Don't terminate the EFIs here - it gets done by the StudySubject
    Call CollectionRemoveAll(mcolEFIs)     ' NCJ 24 Mar 04
    Set mcolEFIs = Nothing
    
    Set mcolEFormIds = Nothing
    
    ' NCJ 15 Aug 06 - Do NOT terminate the VisitEForm here
    ' because it wreaks havoc elsewhere! (Bug 2778)
    'MLM 29/08/02:
'    If Not moVisitEFormInstance Is Nothing Then
'        Call moVisitEFormInstance.Terminate
'    End If
    Set moVisitEFormInstance = Nothing
    
    ' NCJ 24 Mar 04
    Set moArezzo = Nothing

End Sub

'------------------------------------------------
Friend Function UpdateStatus() As Integer
'------------------------------------------------
' Update this visit instance's status
' based on the statuses of its eForm instances
' NB Do this BEFORE calculating the StudySubject's status
'------------------------------------------------
Dim nStatus As Integer
Dim bRequestedNonRepeating As Boolean
Dim oEFI As EFormInstance
Dim oVEForm As VisitEFormRO

    bRequestedNonRepeating = False
    
    If mcolEFIs.Count > 0 Then
        nStatus = eStatus.Requested
        For Each oEFI In mcolEFIs
            If oEFI.Status = eStatus.Requested Then
                ' Is it repeating?
                Set oVEForm = oEFI.VisitInstance.Visit.VisitEFormByEForm(oEFI.eForm)
                If Not oVEForm.Repeating Then
                    bRequestedNonRepeating = True
                End If
            Else
                nStatus = Max(nStatus, oEFI.Status)
            End If
        Next
        ' If status is Success must do additional check for any requested forms,
        ' and if there are any non-repeating requested forms, set status to "missing"
        If nStatus = eStatus.Success And bRequestedNonRepeating Then
            nStatus = eStatus.Missing
        End If
    Else
        nStatus = eStatus.Success
    End If

    ' Update the state object
    If moVIState.Status <> nStatus Then
        moVIState.Status = nStatus
        moVIState.IsChanged = True
    End If
    
    ' Return the result
    UpdateStatus = nStatus

End Function
 
'----------------------------------------------------
Public Property Get SDVStatus() As eSDVStatus
'----------------------------------------------------
' TA 14 Aug 02
' The current "SDV status" of the visit
'----------------------------------------------------

    SDVStatus = moVIState.SDVStatus

End Property

'----------------------------------------------------
Public Property Let SDVStatus(ByVal enSDVStatus As eSDVStatus)
'----------------------------------------------------
' TA 14 Aug 02
' The current "SDV status" of the visit
' NB This doesn't update the "Changed" flag
'----------------------------------------------------

    moVIState.SDVStatus = enSDVStatus

End Property

'----------------------------------------------------
Public Property Get NoteStatus() As eNoteStatus
'----------------------------------------------------
' TA 14 Aug 02
' The current "Note status" of the visit
'----------------------------------------------------

    NoteStatus = moVIState.NoteStatus

End Property

'----------------------------------------------------
Public Property Let NoteStatus(ByVal enNoteStatus As eNoteStatus)
'----------------------------------------------------
' TA 14 Aug 02
' The current "Note status" of the visit
' NB This doesn't update the "Changed" flag
'----------------------------------------------------

    moVIState.NoteStatus = enNoteStatus

End Property

'----------------------------------------------------
Public Property Get DiscrepancyStatus() As eDiscrepancyStatus
'----------------------------------------------------
' TA 14 Aug 02
' The current "Discrepancy status" of the visit
'----------------------------------------------------

    DiscrepancyStatus = moVIState.DiscrepancyStatus

End Property

'----------------------------------------------------
Public Property Let DiscrepancyStatus(ByVal enDiscrepancyStatus As eDiscrepancyStatus)
'----------------------------------------------------
' TA 14 Aug 02
' The current "Discrepancy status" of the visit
' NB This doesn't update the "Changed" flag
'----------------------------------------------------

    moVIState.DiscrepancyStatus = enDiscrepancyStatus

End Property

'----------------------------------------------------
Public Property Get VisitEFormInstance() As EFormInstance
'----------------------------------------------------
' MLM 29/08/02: Added.
' Return the Visit Instance's Visit EForm EFormInstance
'----------------------------------------------------

    Set VisitEFormInstance = moVisitEFormInstance

End Property

'----------------------------------------------------
Friend Property Set VisitEFormInstance(oVisitEFormInstance As EFormInstance)
'----------------------------------------------------
' MLM 29/08/02: Added.
' Used by StudySubject.CreateNewEFI to give the visit instance its visit eform.
'----------------------------------------------------

    Set moVisitEFormInstance = oVisitEFormInstance

End Property
