Attribute VB_Name = "modMIMsgStatus"
'----------------------------------------------------------------------------------------'
'   File:       modMIMsgStatus.bas
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, Aug 2002
'   Purpose:    Routines for updating MIMessageStatus
'----------------------------------------------------------------------------------------'
' Revisions:
'   TA 28/08/2002 - Initial development
'   NCJ 14 Oct 02 - Extra statuses for SDVs
'   NCJ 18 Oct 02 - Added ChangeDoneSDVsToPlanned (for use in Windows/Web)
'   NCJ 21 Oct 02 - Debugging UpdateMIMsgStatus; improving ChangeDoneSDVsToPlanned
'   NCJ 12 Aug 04 - Added new routine ChangePlannedQSDVsToDone (for use in Windows/Web)
'   NCJ 8 Sept 04 - Fix for Null values in ChangePlannedQSDVsToDone
'   NCJ 18 Nov 04 - Issue 2451 - Must check all levels of SDV in ChangeDoneSDVsToPlanned
'   NCJ 22 Nov 04 - Added extra param to GetMIMessageList (after Toby's change)
'----------------------------------------------------------------------------------------'

Option Explicit


'----------------------------------------------------------------------------------------'
Public Sub ChangeDoneSDVsToPlanned(sCon As String, sUserName As String, sUserNameFull As String, _
                                oSubject As StudySubject, oEFI As EFormInstance, enMIMsgSource As MIMsgSource)
'----------------------------------------------------------------------------------------'
' NCJ 18 Oct 02
' Change all Done SDVs to Planned for any changed Responses on this eForm
' and for this eForm, Visit and Subject
' Assume there is AT LEAST ONE CHANGED RESPONSE on the eForm (and that the Responses are already loaded).
' NCJ 18 Nov 04 - Issue 2451 - Check here that there is at least one SDV with status of "Done"
'----------------------------------------------------------------------------------------'
Dim oMIMList As New MIDataLists
Dim vData As Variant
Dim colStatus As Collection
Dim i As Long
Dim bSetToPlanned As Boolean
Dim oResponse As Response
Dim oSDV As MISDV
Dim dblResponseTimeStamp As Double
Dim sResponseValue As String
Dim oTimezone As Timezone
Dim colChangedResponses As Collection
Dim sStudyName As String
Dim bSomethingChanged As Boolean
Dim bThereIsDoneSDV As Boolean

' The default text we use when changing to Planned
Const sCHANGE_TO_PLANNED = "*** Reset to Planned due to change in response data"

    On Error GoTo Errlabel
    
    ' NCJ 18 Nov 04 - Check for Done SDV at Subject, Visit, eForm or Qu level
    bThereIsDoneSDV = (oSubject.SDVStatus = eSDVStatus.ssComplete)
    If Not bThereIsDoneSDV Then
        bThereIsDoneSDV = (oEFI.VisitInstance.SDVStatus = eSDVStatus.ssComplete)
        If Not bThereIsDoneSDV Then
            bThereIsDoneSDV = (oEFI.SDVStatus = eSDVStatus.ssComplete)
            If Not bThereIsDoneSDV Then
                For Each oResponse In oEFI.Responses
                    If (oResponse.SDVStatus = eSDVStatus.ssComplete) Then
                        bThereIsDoneSDV = True
                        Exit For
                    End If
                Next
            End If
        End If
    End If
    ' If there are no SDVs to bother about then forget it
    If Not bThereIsDoneSDV Then Exit Sub
    
    Set oMIMList = New MIDataLists
    
    sStudyName = oSubject.StudyDef.Name
    
    ' Collect all the "Done" SDVs for this subject
    Set colStatus = New Collection
    colStatus.Add eSDVMIMStatus.ssDone
    ' NCJ 6 Nov 02 - Added Scope (empty string for all scopes)
    ' NCJ 22 Nov 04 - Added extra False param (after Toby's change to GetMIMessageList)
    vData = oMIMList.GetMIMessageList(sCon, False, "", "1=1", MIMsgType.mimtSDVMark, "", _
                sStudyName, oSubject.Site, oSubject.label, oSubject.PersonId, _
                -1, -1, -1, "", _
                CollectionToArray(colStatus), _
                False, 0)

    Set colStatus = Nothing
    Set oMIMList = Nothing
    
    If Not IsNull(vData) Then       ' We have some SDVs for this subject
    
        bSomethingChanged = False
        Set colChangedResponses = New Collection
        Set oTimezone = New Timezone
        ' Iterate through the retrieved MIMsgs
        For i = 0 To UBound(vData, 2)
        
            bSetToPlanned = False
            dblResponseTimeStamp = 0
            sResponseValue = ""
            
            Select Case vData(MIMsgCol.mmcScope, i)
            Case MIMsgScope.mimscQuestion
                If vData(MIMsgCol.mmcEFormTaskId, i) = oEFI.EFormTaskId Then
                    Set oResponse = oEFI.Responses.ResponseByResponseId(CLng(vData(MIMsgCol.mmcResponseTaskId, i)), CInt(vData(MIMsgCol.mmcResponseCycle, i)))
                    ' Assume response exists
                    If oResponse.Value <> vData(MIMsgCol.mmcResponseValue, i) Then
                        bSetToPlanned = True
                        ' Remember it's changed
                        colChangedResponses.Add oResponse
                        dblResponseTimeStamp = oResponse.TimeStamp
                        sResponseValue = oResponse.Value
                    End If
                End If
                
            Case MIMsgScope.mimscEForm
                bSetToPlanned = (CLng(vData(MIMsgCol.mmcEFormTaskId, i)) = oEFI.EFormTaskId)
                
            Case MIMsgScope.mimscVisit
                bSetToPlanned = (CLng(vData(MIMsgCol.mmcVisitId, i)) = oEFI.VisitInstance.VisitId) _
                            And (CInt(vData(MIMsgCol.mmcVisitCycle, i)) = oEFI.VisitInstance.CycleNo)
                
            Case MIMsgScope.mimscSubject
                bSetToPlanned = True
    
            End Select
            
            If bSetToPlanned Then
                ' Change this SDV to Planned
                Set oSDV = New MISDV
                Call oSDV.Load(sCon, _
                                CLng(vData(MIMsgCol.mmcObjectId, i)), _
                                CInt(vData(MIMsgCol.mmcObjectSource, i)), _
                                CStr(vData(MIMsgCol.mmcSite, i)))
                Call oSDV.ChangeStatus(eSDVMIMStatus.ssPlanned, sCHANGE_TO_PLANNED, _
                                sUserName, sUserNameFull, _
                                enMIMsgSource, IMedNow, oTimezone.TimezoneOffset, _
                                dblResponseTimeStamp, sResponseValue)
                Call oSDV.Save
                ' Remember something changed for non-response SDVs
                If dblResponseTimeStamp = 0 Then
                    bSomethingChanged = True
                End If
            End If
        Next i
        
        ' Now update the MIMessage Status for any question SDVs that we changed
        ' (this automatically does the eForm, Visit and Subject statuses too)
        If colChangedResponses.Count > 0 Then
            For Each oResponse In colChangedResponses
                Call UpdateMIMsgStatus(sCon, mimtSDVMark, sStudyName, oSubject.StudyId, _
                                oSubject.Site, oSubject.PersonId, _
                                oEFI.VisitInstance.VisitId, oEFI.VisitInstance.CycleNo, oEFI.EFormTaskId, _
                                oResponse.ResponseId, oResponse.RepeatNumber, oSubject)
            Next
        ElseIf bSomethingChanged Then
            ' No Question SDVs changed but an eForm, Visit or Subject did,
            ' so just recalculate from the eForm upwards
            Call UpdateMIMsgStatus(sCon, mimtSDVMark, sStudyName, oSubject.StudyId, _
                                oSubject.Site, oSubject.PersonId, _
                                oEFI.VisitInstance.VisitId, oEFI.VisitInstance.CycleNo, oEFI.EFormTaskId, _
                                0, 0, oSubject)
        End If
        
        Set oTimezone = Nothing
        Set colChangedResponses = Nothing
        Set oSDV = Nothing
        Set oResponse = Nothing

    End If
    
Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modMIMsgStatus.ChangeDoneSDVsToPlanned"
        
End Sub

'----------------------------------------------------------------------------------------'
Public Function ChangePlannedQSDVsToDone(sCon As String, sUserName As String, sUserNameFull As String, _
                                oSubject As StudySubject, oEFI As EFormInstance, _
                                enMIMsgSource As MIMsgSource) As Integer
'----------------------------------------------------------------------------------------'
' NCJ 12 Aug 04 / 23 Aug 04
' Change all Planned question SDVs to Done for this eForm instance
' Returns how many SDVs were changed
'----------------------------------------------------------------------------------------'
Dim oMIMList As New MIDataLists
Dim vSDVData As Variant
Dim colStatus As Collection
Dim colScope As Collection
Dim i As Long
Dim oSDV As MISDV
Dim oTimezone As Timezone
Dim colChangedSDVs As Collection
Dim sStudyName As String
Dim vIndex As Variant
Dim vResponseData As Variant

' The default text we use when changing to DONE
Const sCHANGE_TO_DONE = "*** Set as Done at eForm level"

    On Error GoTo Errlabel
    
    Set oMIMList = New MIDataLists
    
    sStudyName = oSubject.StudyDef.Name
    
    ' Collect all the Planned Question SDVs for this EFI
    Set colStatus = New Collection
    colStatus.Add eSDVMIMStatus.ssPlanned
    Set colScope = New Collection
    colScope.Add MIMsgScope.mimscQuestion
    
    ' NCJ 22 Nov 04 - Added extra False param (after Toby's change to GetMIMessageList)
    vSDVData = oMIMList.GetMIMessageList(sCon, False, "", "1=1", _
                MIMsgType.mimtSDVMark, CollectionToArray(colScope), _
                sStudyName, oSubject.Site, oSubject.label, oSubject.PersonId, _
                oEFI.VisitInstance.VisitId, oEFI.eForm.EFormId, -1, "", _
                CollectionToArray(colStatus), _
                False, 0)

    Set colStatus = Nothing
    Set colScope = Nothing
    Set oMIMList = Nothing
    
    ' This stores the indexes of the SDVS we change
    Set colChangedSDVs = New Collection
    
    If Not IsNull(vSDVData) Then       ' We have some Planned SDVs for this EFI
    
        Set oTimezone = New Timezone
        ' Iterate through the retrieved MIMsgs
        For i = 0 To UBound(vSDVData, 2)
            ' Check it's definitely this EFI
            If vSDVData(MIMsgCol.mmcEFormTaskId, i) = oEFI.EFormTaskId Then
                ' Change this SDV to Done
                ' First need to get Response details (it's not in the vSDVData)
                vResponseData = oMIMList.GetResponseDetails(sCon, _
                                    CStr(vSDVData(MIMsgCol.mmcStudyName, i)), _
                                    CStr(vSDVData(MIMsgCol.mmcSite, i)), _
                                    CLng(vSDVData(MIMsgCol.mmcSubjectId, i)), _
                                    CLng(vSDVData(MIMsgCol.mmcResponseTaskId, i)), _
                                    CInt(vSDVData(MIMsgCol.mmcResponseCycle, i)))
                Set oSDV = New MISDV
                Call oSDV.Load(sCon, _
                                CLng(vSDVData(MIMsgCol.mmcObjectId, i)), _
                                CInt(vSDVData(MIMsgCol.mmcObjectSource, i)), _
                                CStr(vSDVData(MIMsgCol.mmcSite, i)))
                ' NCJ 8 Sept 04 - Beware of empty response values! Added RemoveNull
                ' NCJ 9 Sept 04 - Use ConvertFromNull (RemoveNull not in EForm DLL)
                Call oSDV.ChangeStatus(eSDVMIMStatus.ssDone, sCHANGE_TO_DONE, _
                                sUserName, sUserNameFull, _
                                enMIMsgSource, IMedNow, oTimezone.TimezoneOffset, _
                                CDbl(vResponseData(eResponseDetails.rdResponseTimeStamp, 0)), _
                                ConvertFromNull(vResponseData(eResponseDetails.rdResponseValue, 0), vbString))
                Call oSDV.Save
                ' Keep count of which ones we've done
                colChangedSDVs.Add i
            End If
        Next i
        
        ' Now update the MIMessage Status for any question SDVs that we changed
        ' (this automatically does the eForm, Visit and Subject statuses too)
        If colChangedSDVs.Count > 0 Then
            For Each vIndex In colChangedSDVs
                i = CLng(vIndex)
                Call UpdateMIMsgStatus(sCon, mimtSDVMark, sStudyName, oSubject.StudyId, _
                                oSubject.Site, oSubject.PersonId, _
                                oEFI.VisitInstance.VisitId, oEFI.VisitInstance.CycleNo, oEFI.EFormTaskId, _
                                CLng(vSDVData(MIMsgCol.mmcResponseTaskId, i)), _
                                CInt(vSDVData(MIMsgCol.mmcResponseCycle, i)), oSubject)
            Next
        End If
        
        ' Return how many we did
        ChangePlannedQSDVsToDone = colChangedSDVs.Count
        
        Set oTimezone = Nothing
        Set oSDV = Nothing
        Set colChangedSDVs = Nothing

    End If
    
Exit Function
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modMIMsgStatus.ChangePlannedQSDVsToDone"

End Function


'----------------------------------------------------------------------------------------'
Public Sub UpdateNoteStatus(sCon As String, enScope As MIMsgScope, _
                                sStudyName As String, lClinicalTrialId As Long, sSite As String, lSubjectId As Long, _
                                Optional lVisitId As Long = 0, Optional nVisitCycle As Integer = 0, _
                                Optional lCRFPageTaskId As Long = 0, _
                                Optional lResponseTaskId As Long = 0, Optional nResponseCycle As Integer = 0, _
                                Optional oSubject As StudySubject = Nothing)
'----------------------------------------------------------------------------------------'
'Update the MIMessageStatus in the subject data tables according the contents of the MIMessage table
'
'----------------------------------------------------------------------------------------'
Dim oVI As VisitInstance
    
    'update the db
    Call UpdateNoteStatusInDB(sCon, enScope, _
                                sStudyName, lClinicalTrialId, sSite, lSubjectId, _
                                 lVisitId, nVisitCycle, _
                                 lCRFPageTaskId, _
                                 lResponseTaskId, nResponseCycle)
    
    If Not oSubject Is Nothing Then
        If oSubject.StudyId = lClinicalTrialId And oSubject.Site = sSite And oSubject.PersonId = lSubjectId Then
            'if they've given us the StudySubject and it's the roight one then update it
            Select Case enScope
            Case MIMsgScope.mimscSubject
                oSubject.NoteStatus = 1
            Case MIMsgScope.mimscVisit
                For Each oVI In oSubject.VisitInstancesById(lVisitId)
                    If oVI.CycleNo = nVisitCycle Then
                        oVI.NoteStatus = 1
                    End If
                Next
            Case MIMsgScope.mimscEForm
                oSubject.eFIByTaskId(lCRFPageTaskId).NoteStatus = 1
            Case MIMsgScope.mimscQuestion
                With oSubject.eFIByTaskId(lCRFPageTaskId)
                    If .ResponsesLoaded Then
                        'only do responses if loaded
                        .Responses.ResponseByResponseId(lResponseTaskId, nResponseCycle).NoteStatus = 1
                    End If
                End With
            End Select
        End If
    End If
    
Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|modMIMsgStatus.UpdateNoteStatus"
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub UpdateMIMsgStatus(sCon As String, enMIMsgType As MIMsgType, _
                                sStudyName As String, lClinicalTrialId As Long, sSite As String, lSubjectId As Long, _
                                Optional lVisitId As Long = 0, Optional nVisitCycle As Integer = 0, _
                                Optional lCRFPageTaskId As Long = 0, _
                                Optional lResponseTaskId As Long = 0, Optional nResponseCycle As Integer = 0, _
                                Optional oSubject As StudySubject = Nothing)
'----------------------------------------------------------------------------------------'
' Update the MIMessageStatus in the subject data tables according the contents of the MIMessage table
' and then update the subject object if passed in
'----------------------------------------------------------------------------------------'
Dim enSubjectStatus As Integer
Dim enVisitStatus As Integer
Dim eneFormStatus As Integer
Dim enResponseStatus As Integer
Dim oVI As VisitInstance

    On Error GoTo Errlabel

    Call UpdateMIMsgStatusInDB(enSubjectStatus, enVisitStatus, eneFormStatus, enResponseStatus, _
                                sCon, enMIMsgType, sStudyName, lClinicalTrialId, sSite, lSubjectId, _
                                 lVisitId, nVisitCycle, _
                                 lCRFPageTaskId, _
                                 lResponseTaskId, nResponseCycle)
    
    If Not oSubject Is Nothing Then
        If oSubject.StudyId = lClinicalTrialId And oSubject.Site = sSite And oSubject.PersonId = lSubjectId Then
            'if they've given us the StudySubject and it's the right one then update it
            Select Case enMIMsgType
            Case MIMsgType.mimtDiscrepancy
                oSubject.DiscrepancyStatus = enSubjectStatus
                For Each oVI In oSubject.VisitInstancesById(lVisitId)
                    If oVI.CycleNo = nVisitCycle Then
                        oVI.DiscrepancyStatus = enVisitStatus
                    End If
                Next
                With oSubject.eFIByTaskId(lCRFPageTaskId)
                    .DiscrepancyStatus = eneFormStatus
                    If .ResponsesLoaded Then
                        'only do responses if loaded
                        .Responses.ResponseByResponseId(lResponseTaskId, nResponseCycle).DiscrepancyStatus = enResponseStatus
                    End If
                End With
                
            Case MIMsgType.mimtSDVMark
                oSubject.SDVStatus = enSubjectStatus
                ' Check whether we have a Visit Id
                If lVisitId > 0 Then
                    For Each oVI In oSubject.VisitInstancesById(lVisitId)
                        If oVI.CycleNo = nVisitCycle Then
                            oVI.SDVStatus = enVisitStatus
                        End If
                    Next
                    If lCRFPageTaskId > 0 Then
                        With oSubject.eFIByTaskId(lCRFPageTaskId)
                            .SDVStatus = eneFormStatus
                            ' Only do responses if loaded and we have a valid ResponseTaskId
                            If .ResponsesLoaded And lResponseTaskId > 0 Then
                                .Responses.ResponseByResponseId(lResponseTaskId, nResponseCycle).SDVStatus = enResponseStatus
                            End If
                        End With
                    End If
                End If
                
            End Select
        End If
    End If
    
Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|modMIMsgStatus.UpdateMIMsgStatus"
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function GetHeirachicalMIMsgText(enMIMsgType As MIMsgType, lStatus As Long) As String
'----------------------------------------------------------------------------------------'
'Return the text according to a heirachical MIMSgStatus
'----------------------------------------------------------------------------------------'

    ' Convert the MIMessage table statuses to the DEBS30 hierarchical statuses
    Select Case enMIMsgType
    Case MIMsgType.mimtDiscrepancy
        ' raised, responded, closed
        Select Case lStatus
        Case eDiscrepancyStatus.dsRaised: GetHeirachicalMIMsgText = "Raised"
        Case eDiscrepancyStatus.dsResponded: GetHeirachicalMIMsgText = "Responded"
        Case eDiscrepancyStatus.dsClosed: GetHeirachicalMIMsgText = "Closed"
        Case Else: GetHeirachicalMIMsgText = "None"
        End Select
    Case MIMsgType.mimtSDVMark
        ' NCJ 14 Oct 02 - Planned, Queried, Done, Cancelled
        Select Case lStatus
        Case eSDVStatus.ssPlanned: GetHeirachicalMIMsgText = "Planned" '30
        Case eSDVStatus.ssQueried: GetHeirachicalMIMsgText = "Queried"
        Case eSDVStatus.ssComplete: GetHeirachicalMIMsgText = "Done"
        Case eSDVStatus.ssCancelled: GetHeirachicalMIMsgText = "Cancelled"
        Case Else: GetHeirachicalMIMsgText = "None"
        End Select
    End Select

End Function

