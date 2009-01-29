VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Responses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: Responses.cls
' Copyright InferMed Ltd 2001-2004 All Rights Reserved
' Author: Toby Aldridge, InferMed, July 2001
' Purpose: Responses object in MACRO 2.2
' This represents the responses in an eForm in the Schedule in DE
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
' NCJ 6 Aug 01 - Store which EFormInstance these responses belong to
'               Added Changed property
' NCJ 10 Aug 01 - AddNew method to add a new Response
' NCJ 16 Aug 01 - Added enumerator to allow iteration through Responses collection
' NCJ 31 Aug 01 - Before saving, mark all 'requested' responses as 'missing'
' NCJ 27 Sep 01 - Changed definition of "Changed" property
' NCJ 3 Oct 01 - Added ResponseByResponseId
' TA 04/04/2002: Changes for HadValue property of a response
' TA 25/4/2002: Retrieve elements by dataitem id rather than crfelement id
'               This id changes every time the dataitems on an eform added/removed
' MACRO 3.0
' NCJ 5 July 02 - AddNew is now Friend instead of Public
'               In AddNewResponse, copy LabTestValidate object from first qu. repeat
' NCJ 8 Jul 02 - Needed to re-add -4 proc. attr. in NewEnum
' NCJ 12 Aug 02 - Update ChangeCount in DealWithRequestedResponses
'TA 20/08/2002: added new MIMessage statuses
' RS 22/9/2002: TimezoneOffset handling
' NCJ 2 Oct 02 - Use new IMedNow function instead of CDbl(Now)
' NCJ 18 Oct 02 - New ValuesChanged property
' NCJ 29 Oct 02 - Blank optional questions should have status "Success" (brought forward from 2.2 Patch)
' NCJ 14 Nov 02 - Removed loading/saving of QGroupInstances
' NCJ 28 Jan 03 - Added UserNameFull to ResponseState
' NCJ 29 Apr 03 - Added IgnoreMe to InitialiseResponseState
' NCJ 24/25 Mar 04 - Explicitly clear out object collections in Terminate
' NCJ 23 Aug 04 - In Changed must also check the WillBeSaved property (Bug 2374)
'----------------------------------------------------

Option Explicit

' the collection of response objects for a certain form
Private mcolResponses As Collection
' Index into responses by ElementId & Repeat No
Private mcolElementResponses As Collection

'hang on to states for saving
Private mcolResponseStates As Collection    ' of ResponseState objects

' NCJ 18 Sept 02 - Remember the Subject and the eFormTaskId
Private mlEFormTaskID As Long
Private moSubject As StudySubject

'-----------------------------------------------------------
Public Function ResponseByElement(oEFElement As eFormElementRO, Optional nRepeat As Integer = 1) As Response
'-----------------------------------------------------------
' Return the Response for this eFormElement & repeat no. (if any).
' Assume a repeat no. of 1 if nRepeat not given.
' Will be Nothing if none exists.
'-----------------------------------------------------------

    On Error Resume Next    ' May be no response
    Set ResponseByElement = mcolElementResponses(ResponseElementKey(oEFElement, nRepeat))
    
End Function

'-----------------------------------------------------------
Private Function ResponseElementKey(oElement As eFormElementRO, nRepeat As Integer) As String
'-----------------------------------------------------------
' Get key formed of ElementID and repeat number,
' used for items in mcolElementResponses collection
'-----------------------------------------------------------

    ResponseElementKey = "K" & oElement.ElementID & "|" & nRepeat

End Function

'-----------------------------------------------------------
Public Function ResponseByResponseId(lResponseId As Long, _
                        Optional nRepeat As Integer = 1) As Response
'-----------------------------------------------------------
' Return the Response corresponding to this ResponseTaskId (if any).
' Will be Nothing if none exists (but it should!).
'-----------------------------------------------------------

    On Error Resume Next    ' May be no response
    Set ResponseByResponseId = mcolResponses.Item(Str(lResponseId) & "|" & nRepeat)
    
End Function

'------------------------------------------------
Friend Sub Load(ByVal sConnection As String, oEFI As EFormInstance)
'------------------------------------------------
' NCJ 20 June 01
' Get collection of question responses
' for the given eform instance
' NB Assume that the form's eFormElements collection ALREADY loaded
' TA 25/07/01: moved here from CRFPageInstance class
' Currently friend as calls always come through CRFPageInstace class
'------------------------------------------------
Dim oResponsesDS As StudySubjectDS
Dim oRState As ResponseState
Dim oEFElement As eFormElementRO
Dim lStudyId As Long
Dim sSite As String
Dim lSubjectId As Long

    ' Remember where the responses belong
'    Set moEFI = oEFI
    mlEFormTaskID = oEFI.eFormTaskId
    Set moSubject = oEFI.VisitInstance.Subject
    
    ' Remember these for saving
    lStudyId = moSubject.StudyId
    sSite = moSubject.Site
    lSubjectId = moSubject.PersonID
    
    ' Use the Study Subject data services
    Set oResponsesDS = New StudySubjectDS
    ' Get the Response state objects
    Set mcolResponseStates = CollectionDeSerialise(oResponsesDS.LoadResponses(sConnection, _
                                    lStudyId, sSite, _
                                    lSubjectId, mlEFormTaskID), True)
    Set oResponsesDS = Nothing
    
    ' Set up new collections
    Set mcolResponses = New Collection
    Set mcolElementResponses = New Collection
    
    ' Populate the collection of Response objects
    ' corresponding to the ResponseState objects

    ' TA 25/4/2002
    For Each oRState In mcolResponseStates
        ' Get the eFormElement for this response
        If oEFI.eForm.IsQuestionInEForm(oRState.DataItemId) Then
            Set oEFElement = oEFI.eForm.eFormElementByQuestionId(oRState.DataItemId)
            'basically a question may have been removed from a form after a response has been entered
            'Create and add a new Response object (if element still exitsts)
            Call AddNewResponse(oRState, oEFElement)
        End If
    Next
    
    Set oRState = Nothing
    Set oEFElement = Nothing
    
End Sub

'------------------------------------------------
Public Property Get Changed() As Boolean
'------------------------------------------------
' Returns TRUE if any responses have changed.
' NB Returns FALSE if all Responses are new but untouched.
' NCJ 23 Aug 04 - Must also check the WillBeSaved property
'------------------------------------------------
Dim bChanged As Boolean
Dim oResponse As Response

    bChanged = False
    For Each oResponse In mcolResponses
        ' NCJ 23 Aug 04 - Consider WillBeSaved too (Bug 2374)
        If oResponse.Changed And oResponse.WillBeSaved Then
                bChanged = True
                Exit For        ' Stop at first changed response
        End If
    Next
    Changed = bChanged

End Property

'------------------------------------------------
Public Property Get ValuesChanged() As Boolean
'------------------------------------------------
' Returns TRUE if any response values have changed
' i.e. not counting status changes, validation message changes etc.
'------------------------------------------------
Dim bChanged As Boolean
Dim oResponse As Response

    bChanged = False
    For Each oResponse In mcolResponses
        If oResponse.ValueChanged Then
                bChanged = True
                Exit For        ' Stop at first changed response
        End If
    Next
    ValuesChanged = bChanged

End Property

'------------------------------------------------
Friend Sub Save(ByVal sConnection As String)
'------------------------------------------------
' TA 27 July 01
' Save Responses for a given eform instance
' TA 25/07/01: moved here from eFormInstance class
' Currently friend as calls always come through eFormInstance class
' NCJ 6/8/01 - Use same values that we had on Loading
' NCJ 28/9/01 - Assume Requested responses have already been dealt with
' NCJ 21/11/01 - Save Group Instances here too
' NCJ 14 Nov 02 - No longer save QGroupInstances
'------------------------------------------------
Dim oResponsesDS As StudySubjectDS
Dim oResponse As Response
Dim oVFI As VisitInstance
Dim lStudyId As Long
Dim sSite As String
Dim lSubjectId As Long
'Dim sGIStates As String
Dim oEFI As EFormInstance

    Set oEFI = Me.EFI
    Set oVFI = oEFI.VisitInstance
    ' Remember these for saving
    lStudyId = moSubject.StudyId
    sSite = moSubject.Site
    lSubjectId = moSubject.PersonID

    ' Use the Response data services object
    Set oResponsesDS = New StudySubjectDS
    ' Get the Response state objects
    ' NCJ 6/8/01 - Use stored values for parameters
    
'    If colGroupInstanceStates.Count > 0 Then
'        sGIStates = CollectionSerialise(colGroupInstanceStates)
'    Else
'        sGIStates = ""
'    End If
    
    '3rd parameter is serialised subject state (empty string means no attempt is made to save it)
    ' NCJ 14 Nov 02 - Removed sGIStates parameter
    oResponsesDS.SaveResponses sConnection, CollectionSerialise(mcolResponseStates), _
                        "", _
                        lStudyId, sSite, lSubjectId, _
                        mlEFormTaskID, _
                        oVFI.VisitId, oVFI.CycleNo, _
                        oEFI.eForm.EFormId, oEFI.CycleNo
        
    Set oResponsesDS = Nothing

    ' Now mark all the responses as saved and not new
    For Each oResponse In mcolResponses
       oResponse.ChangeStatetoSaved
    Next

    Set oResponse = Nothing
    Set oEFI = Nothing
    Set oVFI = Nothing
    
End Sub

'---------------------------------------------------------------
Friend Sub DealWithRequestedResponses(Optional nTimezoneOffset As Integer = 0)
'---------------------------------------------------------------
' Before saving, mark all 'requested' responses as 'missing'
' NCJ 12 Aug 02 - Also update their ChangedCount
' RS 22/9/2002 -    Added optional TimezoneOffset parameter
' NCJ 29 Oct 02 - Optional questions should be Success
'---------------------------------------------------------------
Dim oResponse As Response

    For Each oResponse In mcolResponses
       If oResponse.Status = eStatus.Requested Then
            ' NCJ 29 Oct 02 - Optional questions should be Success
            If oResponse.Element.IsOptional Then
                oResponse.Status = eStatus.Success
            Else
                oResponse.Status = eStatus.Missing
            End If
            ' RS 22/9/2002: At this point the Response object contains the server's Timestamp & TimezoneOffset
            '               Update the values to the corresponding local values (which can be different for web client)
            If nTimezoneOffset <> 0 Then
                ' Convert the server Timestamp & TimezoneOffset to corresponding client values
                oResponse.UpdateTimeStamp (nTimezoneOffset)
            End If
            
       End If
       ' TA 04/04/2002: Changes for HadValue property of a response
       'MLM 04/09/03: Make HadValue True for Unobtainable to help Clinical Gateway.
        If oResponse.Value > "" Or oResponse.Status = eStatus.Unobtainable Then
           oResponse.HadValue = True
        End If
        ' NCJ 12 Aug 02 - Update ChangeCount
        Call oResponse.SetChangeCount
    Next

End Sub

'---------------------------------------------------------------
Private Function AddNewResponse(oResponseState As ResponseState, _
                        oElement As eFormElementRO) As Response
'---------------------------------------------------------------
' Create a new Response object, set its ResponseState,
' and add it to the Responses collection.
' NCJ 4 Mar 02 - If it's a LabTest question with Repeat > 1,
' must set up correct LabtestValidate object
'---------------------------------------------------------------
Dim oResponse As Response

    Set oResponse = New Response
    Call oResponse.SetState(oResponseState, Me.EFI, oElement)
    mcolResponses.Add oResponse, oResponse.Key
    ' Also index each Response by its eFormElementId AND Repeat no.
    mcolElementResponses.Add oResponse, ResponseElementKey(oElement, oResponse.RepeatNumber)
    
    ' NCJ 4 Mar 02 - Point to correct LabtestValidate object
    If oResponseState.RepeatNumber > 1 And oElement.DataType = eDataType.LabTest Then
        ' Copy the LabTestValidate object from the first repeat of the question
        ' Assume first repeat exists but if LabTestValidate is nothing it doesn't matter
        Set oResponse.LabTestValidate = ResponseByElement(oElement, 1).LabTestValidate
    End If
    
    Set AddNewResponse = oResponse
    
    Set oResponse = Nothing
    
End Function

'---------------------------------------------------------------
Friend Function AddNew(oElement As eFormElementRO, Optional nRepeat As Integer = 1) As Response
'---------------------------------------------------------------
' Create a new, initialised Response object for this eForm Element
' with the specified Repeat number
' and add it to the Responses collection
'---------------------------------------------------------------
Dim oResponseState As ResponseState

    Set oResponseState = New ResponseState
    ' Initialise values in ResponseState object
    Call InitialiseResponseState(oResponseState, oElement, nRepeat)
    mcolResponseStates.Add oResponseState, oResponseState.Key
    
    ' Create Response and return the newly created Response object
    ' using this ResponseState
    Set AddNew = AddNewResponse(oResponseState, oElement)
        
    Set oResponseState = Nothing
    
End Function

'---------------------------------------------------------------
Private Sub InitialiseResponseState(oResponseState As ResponseState, _
                        oElement As eFormElementRO, Optional nRepeat As Integer = 1)
'---------------------------------------------------------------
' Initialise all the values in a ResponseState object
' NCJ 5 Nov 01 - Added RepeatNumber = 1
' NCJ 12 Aug 02 - Added HadValue and ChangeCount
' RS 22/9/2002 - Added TimeStampTZ initialization
' NCJ 28 Jan 03 - Added User name
' NCJ 29 Apr 03 - Added IgnoreMe = false
'---------------------------------------------------------------

    With oResponseState
    
        .IsChanged = True
        .IsNew = True
        ' ResponseID (as in MACRO 2.1) is (CRFPageTaskId*10000 + CRFElementID)
        .ResponseId = mlEFormTaskID * 10000 + oElement.ElementID
        .RepeatNumber = nRepeat
        .Status = eStatus.Requested
        .LockStatus = eLockStatus.lsUnlocked
        .DataItemId = oElement.QuestionId
        .ElementID = oElement.ElementID
        .TimeStamp = IMedNow                                ' RS 22/9/2002: Initial values, will be overriden
        .TimeSTampTZ = moSubject.Timezone.TimezoneOffset      ' RS 22/9/2002
        
        .Comments = ""
        .CTCGrade = eCTCGrade.ctcNone
        .LabCode = ""
        .LabTestDate = eMACRODate.mdUnspecified
        .NRStatus = eNormalRangeLNorH.nrNotfound
        .OverruleReason = ""
        .ReasonForChange = ""
        .ValidationMessage = ""
        .Value = ""
        .ValueCode = ""
        .HadValue = False
        .ChangeCount = 0
        
        .UserCode = moSubject.UserName
        .UserNameFull = moSubject.UserNameFull
        .SoftwareVersion = App.Major & "." & App.Minor & "." & App.Revision

        'TA 20/08/2002: added new MIMessage statuses
        .DiscrepancyStatus = 0
        .SDVStatus = 0
        .NoteStatus = 0
        
        ' NCJ 29 Apr 03 - default is not to ignore
        .IgnoreMe = False

    End With
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'
' Tidy up
' NCJ 24 Mar 04 - Explicitly clear out object collections
'----------------------------------------------------------------------------------------'

    Call Terminate
    
End Sub

'----------------------------------------------------------------------------------------'
Friend Sub Terminate()
'----------------------------------------------------------------------------------------'
' Tidy up
' NCJ 24/25 Mar 04 - Explicitly clear out object collections
'----------------------------------------------------------------------------------------'
Dim oResponse As Response

    ' Responses terminated later
    Call CollectionRemoveAll(mcolElementResponses)
    Set mcolElementResponses = Nothing
    
    If Not mcolResponses Is Nothing Then
        For Each oResponse In mcolResponses
            Call oResponse.Terminate
        Next
        Call CollectionRemoveAll(mcolResponses)
    End If
    Set mcolResponses = Nothing
    Set oResponse = Nothing
    
    Call CollectionRemoveAll(mcolResponseStates)
    Set mcolResponseStates = Nothing

    Set moSubject = Nothing     ' NCJ 24 Mar 04

End Sub

'----------------------------------------------------------------------------------------'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
'----------------------------------------------------------------------------------------'
' The enumerator to allow For...Each iterations.
'----------------------------------------------------------------------------------------'
    
    Set NewEnum = mcolResponses.[_NewEnum]

End Function

'----------------------------------------------------------------------------------------'
Public Property Get Count() As Long
'----------------------------------------------------------------------------------------'
' Count of responses
'----------------------------------------------------------------------------------------'

    Count = mcolResponses.Count

End Property

'----------------------------------------------------------------------------------------'
Public Property Get Item(vIndex As Variant) As Response
'----------------------------------------------------------------------------------------'
' Get a Response from the collection
'----------------------------------------------------------------------------------------'

    Set Item = mcolResponses.Item(vIndex)

End Property
 
'----------------------------------------------------------------------------------------'
Public Property Get EFI() As EFormInstance
'----------------------------------------------------------------------------------------'
' NCJ 17 Sept 02
' Get the eFormInstance to which we belong
'----------------------------------------------------------------------------------------'

    Set EFI = moSubject.eFIByTaskId(mlEFormTaskID)

End Property
 
'----------------------------------------------------
Friend Sub ReapplyArezzo()
'----------------------------------------------------
' NCJ 18 Sept 02
' Reapply current response values to Arezzo if necessary
'----------------------------------------------------
Dim oResponse As Response

    For Each oResponse In mcolResponses
        Call oResponse.ReapplyArezzo
    Next

End Sub

