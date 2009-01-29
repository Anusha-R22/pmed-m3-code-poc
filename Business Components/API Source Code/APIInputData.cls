VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APIInputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
' File: APIInputData.cls
' Copyright InferMed Ltd 2004 All Rights Reserved
' Author: Nicky Johns, Feb 2004
' Purpose: Routines for inputting MACRO subject data from XML
'------------------------------------------------------------------
' REVISIONS
' NCJ 7-13 Aug 2003 - Initial development
' NCJ 26-27 Aug 03 - Do not allow enty into non-enterable fields, locked/frozen visits etc.
' NCJ 3 Sept 03 - Added error message for read-only subject
' NCJ 9 Sept 03 - Added error message for eForm/Visit date missing
' NCJ 2 Feb 04 - Allow data entry into hidden fields
' NCJ 2 Feb 04 - This file created from original XMLInputData.cls
' TA 29/04/2005 - allowed timestamp and timezone to be specified when loading data
' TA 24/05/2005 - allow trialid, visitid,eformid, and crflementid to be used
' NCJ 17 Aug 06 - Check question authorisation
' NCJ 18 Aug 06 - Moved GetCodeAndCycle to modAPIUtils
' NCJ 23 Oct 06 - Bug 2827 - Check registration status in SaveEFIResponses
'------------------------------------------------------------------

Option Explicit


Private moXMLInputDoc As MSXML2.DOMDocument

Private moUser As MACROUser
Private moSubject As StudySubject

' Collection of error strings
Private msXMLErrorReport As String
' Count any errors
Private mnErrors As Integer

Private msCurSubject As String

' The current visit, eForm and question
Private msCurVisit As String
Private msCurEForm As String
Private msCurQuestion As String

'type to return timestamp and timezone info
Private Type TimeAttributes
    Timezone As Integer
    Timestamp As Double
End Type

'---------------------------------------------------------------------
Public Function InputXMLSubjectData(ByVal sSerialisedUser As String, _
                                ByVal sXMLDataInput As String, _
                                ByRef sXMLReport As String) As eDataInputResult
'---------------------------------------------------------------------
' Given a user and an XML data input string,
' add all the specified data values
' Returns error code
'   0 - All OK
'   1 - XMLDataInput not understood
'   2 - Subject does not exist
'   2 - Subject could not be opened
'   3 - Some data not added
' If not all OK, sXMLReport contains error messages
'---------------------------------------------------------------------
Dim sErrMsg As String
Dim sSite As String
Dim sStudyName As String
Dim sSubjLabel As String

    On Error GoTo Errlabel
    
    mnErrors = 0
    msXMLErrorReport = ""
    
    ' Create the MACRO User
    Set moUser = New MACROUser
    Call moUser.SetStateHex(sSerialisedUser)

    'Load the XML and retrieve the subject details
    If SetXMLRequest(sXMLDataInput, moXMLInputDoc, sStudyName, sSite, sSubjLabel) Then
        Set moSubject = LoadSubject(moUser, sStudyName, sSite, sSubjLabel, sErrMsg)
        If Not moSubject Is Nothing Then
            msCurSubject = sStudyName & "/" & sSite & "/" & sSubjLabel
            Call ProcessDataInput(moXMLInputDoc, moSubject)
            ' Tidy everything up before we leave
            Call Terminate
            If mnErrors = 0 Then
                ' Assume all OK
                InputXMLSubjectData = eDataInputResult.Success
            Else
                ' We hit some problemettes
                InputXMLSubjectData = eDataInputResult.DataNotAdded
            End If
        Else
            ' Return sErrMsg
            AddErrorMsg eDataInputError.SubjectNotOpened, sErrMsg
            InputXMLSubjectData = eDataInputResult.SubjectNotOpened
        End If
    Else
        ' Return error in XML
        AddErrorMsg eDataInputError.InvalidXML, "Error in XML data input string"
        InputXMLSubjectData = eDataInputResult.InvalidXML
    End If
    
    sXMLReport = CollectErrors
'    Set mcolErrors = Nothing
    
Exit Function
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIInputData.InputXMLSubjectData"

End Function

'---------------------------------------------------
Private Sub ProcessDataInput(oXMLDoc As MSXML2.DOMDocument, oSubject As StudySubject)
'---------------------------------------------------
' Input all the data values specified in oXMLDoc
'---------------------------------------------------
Dim oVINodeList As MSXML2.IXMLDOMNodeList
Dim i As Integer

    If oSubject.ReadOnly Then
        AddErrorMsg eDataInputError.SubjectNotOpened, "Subject " & msCurSubject & " is read-only: " & oSubject.ReadOnlyReason
    Else
    
        ' Get the Visit nodes
        Set oVINodeList = oXMLDoc.getElementsByTagName(gsTAG_VISIT)
        If oVINodeList.length > 0 Then
            For i = 0 To oVINodeList.length - 1
                Call GetVisitNodeData(oVINodeList(i), oSubject)
            Next
        Else
            ' Nothing to do!
            ' (Return error message?)
        End If
    
        Set oVINodeList = Nothing
    
    End If
    
End Sub

'---------------------------------------------------
Private Sub Class_Terminate()
'---------------------------------------------------

    Call Terminate
    
End Sub

'---------------------------------------------------
Private Sub Terminate()
'---------------------------------------------------

    Dim oArezzo As Arezzo_DM
    
    Set moXMLInputDoc = Nothing
    Set moUser = Nothing
    
    If Not moSubject Is Nothing Then
        Set oArezzo = moSubject.Arezzo
        Call moSubject.StudyDef.RemoveSubject
        ' Clear up the study def
        Call moSubject.StudyDef.Terminate
         ' clear up arezzo
        If Not oArezzo Is Nothing Then
            oArezzo.Finish
            Set oArezzo = Nothing
        End If
        Set moSubject = Nothing
    End If
    
End Sub

'---------------------------------------------------
Private Sub GetVisitNodeData(oElReqVINode As MSXML2.IXMLDOMElement, oSubject As StudySubject)
'---------------------------------------------------
' Process this visit's data
' oElReqVINode is a visit node from the Data Input XML
'---------------------------------------------------
Dim sCode As String
Dim nCycle As Integer
Dim oVI As VisitInstance
Dim oVisit As VisitRO
Dim oEFINodeList As MSXML2.IXMLDOMNodeList
Dim i As Integer
Dim bVisitFound As Boolean

    msCurVisit = ""
    msCurEForm = ""
    msCurQuestion = ""
    
    Call GetCodeAndCycle(oElReqVINode, sCode, nCycle)
    msCurVisit = sCode & "[" & nCycle & "]"
    
    'TA 24/05/2005
    'assume that if visit code is numeric then it is an id
    If IsNumeric(sCode) Then
        Set oVisit = oSubject.StudyDef.VisitById(CLng(sCode))
    Else
        Set oVisit = oSubject.StudyDef.VisitByCode(sCode)    ' Returns Nothing if not a Visit
    End If
    
    ' Get out now if visit doesn't exist
    If oVisit Is Nothing Then
        ' DPH 17/10/2005 - Change order of error messages so Visit[cycle]:eForm[cycle]:Question[cycle] is first
        AddErrorMsg eDataInputError.VisitNotExist, msCurVisit & " - visit does not exist"
        Exit Sub
    End If
    
    bVisitFound = False
    ' Search for the right visit cycle
    For Each oVI In oSubject.VisitInstancesById(oVisit.VisitId)
        If oVI.CycleNo = nCycle Then
            ' We want this one
            bVisitFound = True
            ' Check that it's unlocked
            If oVI.LockStatus <> eLockStatus.lsUnlocked Then
                ' DPH 17/10/2005 - Change order of error messages so Visit[cycle]:eForm[cycle]:Question[cycle] is first
                AddErrorMsg eDataInputError.VisitLockedFrozen, msCurVisit & " - visit is locked or frozen"
            Else
                ' Process the list of its eForms
                Set oEFINodeList = oElReqVINode.getElementsByTagName(gsTAG_EFORM)
                If oEFINodeList.length > 0 Then
                    For i = 0 To oEFINodeList.length - 1
                        Call GetEformNodeData(oEFINodeList(i), oVI)
                    Next
                Else
                    ' No eForms so nothing to do (return error?)
                    ' DPH 17/10/2005 - Change order of error messages so Visit[cycle]:eForm[cycle]:Question[cycle] is first
                    AddErrorMsg eDataInputError.InvalidXML, msCurVisit & " - No eForms/questions specified for this visit."
                End If
                Exit For
            End If
        End If
    Next
    
    If Not bVisitFound Then
        ' DPH 17/10/2005 - Change order of error messages so Visit[cycle]:eForm[cycle]:Question[cycle] is first
        AddErrorMsg eDataInputError.VisitNotExist, msCurVisit & " - visit does not exist"
    End If
    
    Set oVisit = Nothing
    Set oVI = Nothing
    Set oEFINodeList = Nothing
    
Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIInputData.GetVisitNodeData"

End Sub

'---------------------------------------------------
Private Sub GetEformNodeData(oElReqEFINode As MSXML2.IXMLDOMElement, oVI As VisitInstance)
'---------------------------------------------------
' Process this eForm node
' oElReqEFINode is an eForm node from the Data Input XML
' and oVI is the current visit instance
'---------------------------------------------------
Dim sCode As String
Dim nCycle As Integer
Dim oEFI As EFormInstance
Dim oEform As eFormRO
Dim i As Integer

    On Error GoTo Errlabel
    
    ' Get the code and cycle of the required eForm
    Call GetCodeAndCycle(oElReqEFINode, sCode, nCycle)
    msCurEForm = sCode & "[" & nCycle & "]"
    
    'TA 24/05/2005
    'assume is an id if code is numeric
    If IsNumeric(sCode) Then
        Set oEform = oVI.Subject.StudyDef.eFormById(CLng(sCode))
    Else
        Set oEform = oVI.Subject.StudyDef.eFormByCode(sCode)    ' Returns Nothing if not an eForm
    End If
    
    ' Can't do it if non-existent eForm
    If oEform Is Nothing Then
        ' DPH 17/10/2005 - Change order of error messages so Visit[cycle]:eForm[cycle]:Question[cycle] is first
        AddErrorMsg eDataInputError.EFormNotExist, GetCurEFormSpec & " - eForm does not exist"
        Exit Sub
    End If
    
    ' Must be a specific EFI
    Set oEFI = oVI.eFormInstanceByEFormCycle(oEform, nCycle)
    If Not oEFI Is Nothing Then
        If oEFI.LockStatus <> eLockStatus.lsUnlocked Then
            ' DPH 17/10/2005 - Change order of error messages so Visit[cycle]:eForm[cycle]:Question[cycle] is first
            AddErrorMsg eDataInputError.EFormLockedFrozen, GetCurEFormSpec & " - eForm is locked or frozen"
        Else
            ' Process all its questions
            Call AppendEFormData(oEFI, oElReqEFINode)
        End If
    Else
        ' DPH 17/10/2005 - Change order of error messages so Visit[cycle]:eForm[cycle]:Question[cycle] is first
        AddErrorMsg eDataInputError.EFormNotExist, GetCurEFormSpec & " - eForm does not exist"
    End If
    
    Set oEFI = Nothing
    Set oEform = Nothing
    
Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIInputData.GetEformNodeData"

End Sub

'---------------------------------------------------
Private Sub AppendEFormData(oEFI As EFormInstance, _
                                oElReqEFINode As MSXML2.IXMLDOMElement)
'---------------------------------------------------
' Add all the specified question values for this eForm
' NCJ 17 Aug 06 - Add the lab if specified
'---------------------------------------------------
Dim oQNodeList As MSXML2.IXMLDOMNodeList
Dim i As Integer
Dim nTimeZone As Integer
Dim nOriginalTimezone As Integer
Dim oSubject As StudySubject

    On Error GoTo Errlabel
    
    ' Get the list of question values
    Set oQNodeList = oElReqEFINode.getElementsByTagName(gsTAG_QUESTION)
    If oQNodeList.length > 0 Then
        'get timezone of first question
        nTimeZone = GetTimeAttributes(oQNodeList(0)).Timezone
        If nTimeZone <> NULL_INTEGER Then
                'time zone provided? (only works if both are provided)
                Set oSubject = oEFI.VisitInstance.Subject
                'store origial timezone
                nOriginalTimezone = oSubject.Timezone.TimezoneOffset
                Call oSubject.Timezone.SetTimezoneOffset(nTimeZone)
        End If
        ' We want specific questions
        If LoadEFIResponses(oEFI) Then
            ' NCJ 17 Aug 06 - Set the lab (if any)
            Call SetLabCode(oEFI, oElReqEFINode)
            ' DPH 21/03/2005 - Refresh skips and derivations on opening eForm
            Call oEFI.RefreshSkipsAndDerivations(OpeningEForm, moUser.UserName)
            For i = 0 To oQNodeList.length - 1
                Call GetQuestionNodeData(oQNodeList(i), oEFI)
            Next
            Call SaveEFIResponses(oEFI)
            Call RemoveEFIResponses(oEFI)
        End If
        'restore orginal timezone if changed
        If nTimeZone <> NULL_INTEGER Then
            oSubject.Timezone.SetTimezoneOffset (nOriginalTimezone)
            Set oSubject = Nothing
        End If
    Else
        ' Nothing to do!
        ' DPH 17/10/2005 - Change order of error messages so Visit[cycle]:eForm[cycle]:Question[cycle] is first
        Call AddErrorMsg(eDataInputError.InvalidXML, GetCurEFormSpec & " - no questions specified for this eForm.")
    End If

Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIInputData.AppendEFormData"

End Sub

'---------------------------------------------------
Private Sub AddQuestionValue(oResponse As Response, sValue As String, TimeInfo As TimeAttributes)
'---------------------------------------------------
' Add value to this response
' TA 28/04/2005 - optionally allow timestamp and timezone to be set
' NCJ 17 Aug 06 - Check authorisation (only current user role)
'---------------------------------------------------
Dim sErrMsg As String
Dim bChanged As Boolean
Dim sRFC As String
Dim ValidateStatus As eStatus
Dim nOriginalTimezone As Integer
Dim oSubject As StudySubject
' NCJ 17 Aug 06 - Authorisation
Dim sAuthName As String
Dim sAuthNameFull As String

Const sAUTO_RFC = "*** Value changed by an external application using the API"

    'check for timestamp info
    If (TimeInfo.Timestamp = NULL_DATE) Or (TimeInfo.Timezone = NULL_INTEGER) Then
        'no time info provided
        ValidateStatus = oResponse.ValidateValue(sValue, sErrMsg, bChanged)
    Else
        'time info provided? (only works if both are provided)
        'nb can't set timezone at this level
        ValidateStatus = oResponse.ValidateValue(sValue, sErrMsg, bChanged, TimeInfo.Timestamp)
    End If
    
    If ValidateStatus <> eStatus.InvalidData Then
        If bChanged Then
            If oResponse.RequiresValueRFC Then
                sRFC = sAUTO_RFC
            Else
                sRFC = ""
            End If
            ' NCJ 16 Aug 06 - Check Authorisation
            If AuthorisedOK(oResponse, sAuthName, sAuthNameFull) Then
                Call oResponse.ConfirmValue("", sRFC, sAuthName, sAuthNameFull)
                Call oResponse.EFormInstance.RefreshSkipsAndDerivations(ChangingResponse, moUser.UserName, oResponse)
                ' DPH 21/03/2006 - Store warning message (if there is one)
                If ValidateStatus = eStatus.Warning Or ValidateStatus = eStatus.Inform Then
                    AddErrorMsg eDataInputError.ValueWarning, GetCurQuestionSpec & " " & sErrMsg
                End If
            Else
                ' Can't accept this unauthorised answer
                AddErrorMsg eDataInputError.QuestionNotEnterable, _
                        GetCurQuestionSpec & " No authorisation (" & sValue & ")"
            End If
            
        End If
    Else
        AddErrorMsg eDataInputError.ValueRejected, _
                        GetCurQuestionSpec & " value rejected (" & sValue & ") - " & sErrMsg
    End If
    
Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIInputData.AddQuestionValue"

End Sub
 
'---------------------------------------------------
Private Function AuthorisedOK(oResponse As Response, _
                ByRef sAuthName As String, ByRef sAuthNameFull As String) As Boolean
'---------------------------------------------------
' NCJ 17 Aug 06 - Can we authorise this question?
' Returns TRUE if OK (with AuthNames set up appropriately) or FALSE if not
'---------------------------------------------------

    AuthorisedOK = True
    sAuthName = ""
    sAuthNameFull = ""
    If oResponse.Element.Authorisation > "" Then
        ' Only check the current user role for now
        If LCase(moUser.UserRole) = LCase(oResponse.Element.Authorisation) Then
            ' They are allowed to authorise
            sAuthName = moUser.UserName
            sAuthNameFull = moUser.UserNameFull
        Else
            ' Can't accept this answer
            AuthorisedOK = False
        End If
    End If

End Function

'---------------------------------------------------
Private Function GetCurEFormSpec() As String
'---------------------------------------------------

    GetCurEFormSpec = msCurVisit & ":" & msCurEForm
    
End Function

'---------------------------------------------------
Private Function GetCurQuestionSpec() As String
'---------------------------------------------------

    GetCurQuestionSpec = GetCurEFormSpec & ":" & msCurQuestion
    
End Function

'---------------------------------------------------
Private Sub SetLabCode(oEFI As EFormInstance, oElNode As MSXML2.IXMLDOMElement)
'---------------------------------------------------
' Set the lab code (if any) for this eForm
' Does nothing if no lab
' Of invalid lab specified, adds error message
'---------------------------------------------------
Dim vLab As Variant
Dim sLab As String
Dim bLabOK As Boolean
Dim sSite As String

    ' Look for lab attribute
    vLab = oElNode.getAttribute(gsATTR_LAB)
    ' See if there is one
    If Not IsNull(vLab) Then
        sLab = LCase(CStr(vLab))
        If sLab > "" Then
            ' Already set up?
            If oEFI.LabCode <> sLab Then
                ' Make sure lab is valid for site
                sSite = oEFI.VisitInstance.Subject.Site
                If CheckLab(moUser.CurrentDBConString, sLab, sSite) Then
                    ' Change to this one if it's valid
                    oEFI.LabCode = sLab
                Else
                    Call AddErrorMsg(eDataInputError.InvalidLab, _
                            "Lab code '" & sLab & "' is not recognised for site " & sSite)
                End If
            End If
        End If
    End If

End Sub

'---------------------------------------------------
Private Function GetTimeAttributes(oElNode As MSXML2.IXMLDOMElement) As TimeAttributes
'---------------------------------------------------
' Retrieve the timezone and timestamp attributes from this element
'---------------------------------------------------
Dim vTimestamp As Variant
Dim vTimezone As Variant
Dim Attributes As TimeAttributes

    On Error GoTo Errlabel
    
    vTimestamp = oElNode.getAttribute(gsATTR_TIMESTAMP) 'may not be one
    vTimezone = oElNode.getAttribute(gsATTR_TIMEZONE) 'may not be one
    
    If IsNull(vTimestamp) Then
        Attributes.Timestamp = NULL_DATE
    Else
        If IsNumeric(vTimestamp) Then
            'ms double date
            Attributes.Timestamp = vTimestamp
        Else
            ' Assume a valid date in format YYYY/MM/DD HH:DD:SS
            Attributes.Timestamp = CDbl(CDate(vTimestamp))
        End If
    End If
    
    If IsNull(vTimezone) Then
        Attributes.Timezone = NULL_INTEGER
    Else
        If IsNumeric(vTimezone) Then
            'assume is MACRO timezone (in minutes)
            Attributes.Timezone = vTimezone
        Else
            ' Assume a valid timezone in format GMT+HH:MM or GMT-HH:MM
            'macro counts time zones in minutes
            Attributes.Timezone = ((Mid(vTimezone, 5, 2) * 60) + Mid(vTimezone, 8, 2))
            If Mid(vTimezone, 4, 1) = "+" Then
                'pluses become negatives in MACRO time zone format
                Attributes.Timezone = -1 * Attributes.Timezone
            End If
        End If
     End If
     
     GetTimeAttributes = Attributes
                            
Exit Function
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.GetTimeAttributes"

End Function

'---------------------------------------------------
Private Sub GetQuestionNodeData(oElReqQNode As MSXML2.IXMLDOMElement, _
                        oEFI As EFormInstance)
'---------------------------------------------------
' Add nodes to the oElRetEFINode (returned eForm node) representing this eForm's data
' oElReqQNode is an eForm node from the Data Request XML (maybe representing a number of eForm cycles)
' oEFI is the current eForm instance
' Assume oEFI has its responses already loaded
'---------------------------------------------------
Dim sCode As String
Dim nCycle As Integer
Dim oResponse As Response
Dim oElement As eFormElementRO
Dim vValue As Variant
Dim TimeInfo As TimeAttributes
    
    On Error GoTo Errlabel
   
    Call GetCodeAndCycle(oElReqQNode, sCode, nCycle)
    msCurQuestion = sCode & "[" & nCycle & "]"
    
    TimeInfo = GetTimeAttributes(oElReqQNode)
    ' Trap for question not existing
    On Error Resume Next
    
     'TA 24/05/2005
    'assume is an id if numeric
    If IsNumeric(sCode) Then
        Set oElement = oEFI.eForm.eFormElementById(CLng(sCode))
    Else
        Set oElement = oEFI.eForm.eFormElementByCode(sCode)
    End If
    
    ' Reset error handler
    On Error GoTo Errlabel
    
    ' Can't do it if question doesn't exist
    If oElement Is Nothing Then
        ' DPH 17/10/2005 - Change order of error messages so Visit[cycle]:eForm[cycle]:Question[cycle] is first
        AddErrorMsg eDataInputError.QuestionNotExist, GetCurQuestionSpec & " - question does not exist"
        Exit Sub
    End If
    
    ' DPH 04/01/2006 - get response from function which creates new RQG row if required
    ' Set oResponse = oEFI.Responses.ResponseByElement(oElement, nCycle)
    Call GetCorrectResponse(oEFI, oElement, nCycle, oResponse)
    
    ' Make sure it exists...
    If Not oResponse Is Nothing Then
        ' Check it's not locked, frozen, derived, N/A etc.
        ' NCJ 2 Feb 04 - Use our own CanEnterResponse rather than oResponse.Enterable (to allow hidden fields)
        If CanEnterResponse(oResponse) Then
            vValue = oElReqQNode.getAttribute(gsATTR_VALUE)
            If IsNull(vValue) Then
                ' DPH 17/10/2005 - Change order of error messages so Visit[cycle]:eForm[cycle]:Question[cycle] is first
                AddErrorMsg eDataInputError.InvalidXML, GetCurQuestionSpec & " - no question value given."
            Else
                Call AddQuestionValue(oResponse, CStr(vValue), TimeInfo)
            End If
        Else
            ' Response is not enterable
            ' DPH 17/10/2005 - Change order of error messages so Visit[cycle]:eForm[cycle]:Question[cycle] is first
            AddErrorMsg eDataInputError.QuestionNotEnterable, GetCurQuestionSpec & "  - question is not an enterable response"
        End If
    Else
        ' DPH 17/10/2005 - Change order of error messages so Visit[cycle]:eForm[cycle]:Question[cycle] is first
        AddErrorMsg eDataInputError.QuestionNotExist, GetCurQuestionSpec & " - question does not exist"
    End If
    
    Set oElement = Nothing
    Set oResponse = Nothing
    
Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIInputData.GetQuestionNodeData"

End Sub

'---------------------------------------------------
Private Function LoadEFIResponses(oEFI As EFormInstance) As Boolean
'---------------------------------------------------
' Load an eForm's responses
' Return TRUE if loaded OK
'---------------------------------------------------
Dim sErrMsg As String
Dim sTok1 As String
Dim sTok2 As String

    If oEFI.VisitInstance.Subject.LoadResponses(oEFI, sErrMsg, sTok1, sTok2) = lrrReadWrite Then
        LoadEFIResponses = True
    Else
        sErrMsg = GetCurEFormSpec & " - eForm responses not loaded - " & sErrMsg
        Call AddErrorMsg(eDataInputError.EFormInUse, sErrMsg)
        LoadEFIResponses = False
    End If

End Function

'---------------------------------------------------
Private Sub RemoveEFIResponses(oEFI As EFormInstance)
'---------------------------------------------------
' Remove an eForm's responses
'---------------------------------------------------

    Call oEFI.VisitInstance.Subject.RemoveResponses(oEFI, True)

End Sub

'---------------------------------------------------
Private Sub SaveEFIResponses(oEFI As EFormInstance)
'---------------------------------------------------
' Save an eForm's responses
'---------------------------------------------------
Dim sErrMsg As String
Dim oRegister As clsRegisterWWW
Dim nResult As Integer

    On Error GoTo Errlabel
    
    ' Check the eForm and visit dates first
    If OKToSaveEForm(oEFI, sErrMsg) Then
        Select Case oEFI.VisitInstance.Subject.SaveResponses(oEFI, sErrMsg)
        Case eSaveResponsesResult.srrSubjectReloaded
            ' Try one more time
            If oEFI.VisitInstance.Subject.SaveResponses(oEFI, sErrMsg) <> srrSuccess Then
                AddErrorMsg eDataInputError.NoLockForSave, GetCurEFormSpec & " could not be saved: " & sErrMsg
            End If
        Case eSaveResponsesResult.srrNoLockForSaving
            AddErrorMsg eDataInputError.NoLockForSave, GetCurEFormSpec & " could not be saved: " & sErrMsg
        Case eSaveResponsesResult.srrSuccess
            ' Excellent!
            ' NCJ 23 Oct 06 - Bug 2827 - Check registration status here if necessary
            If moSubject.RegistrationStatus = eRegStatus.NotReady Then
                If oEFI.Complete Then
                    Set oRegister = New clsRegisterWWW
                    ' Set up with current subject details
                    Call oRegister.Initialise(moSubject, moUser.CurrentDBConString, moUser.DatabaseCode)
                    ' This will change to "Ready" if applicable
                    Call oRegister.RegistrationTrigger(oEFI)
                    Set oRegister = Nothing
                End If
            End If
        End Select
    Else
        ' Can't save because missing visit/eForm date
        AddErrorMsg eDataInputError.NoVisitEformDate, GetCurEFormSpec & sErrMsg
    End If
    
Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIInputData.SaveEFIResponses"

End Sub

'---------------------------------------------------
Private Sub AddErrorMsg(enMsgType As eDataInputError, _
                        ByVal sMsg As String)
'---------------------------------------------------
' Add an error message to the collection
'---------------------------------------------------

    mnErrors = mnErrors + 1
    Call AddXMLErrMsg(enMsgType, sMsg)

End Sub

'---------------------------------------------------
Private Function CollectErrors() As String
'---------------------------------------------------
Dim sErrs As String
    
    sErrs = ""
    If msXMLErrorReport > "" Then
        sErrs = gsXML_VERSION_HEADER & vbCrLf & gsXMLTAG0_INPUTERR & vbCrLf
        sErrs = sErrs & msXMLErrorReport & gsXMLTAG1_INPUTERR
    End If
    CollectErrors = sErrs

End Function

'---------------------------------------------------
Private Sub AddXMLErrMsg(enMsgType As eDataInputError, _
                        ByVal sMsg As String)
'---------------------------------------------------
' Add error message to accumulating XML error message string (in msXMLErrorReport)
'---------------------------------------------------
Const sERR_TAG = "<DataErrMsg "
Const sEND_TAG = "/>"
Const sTYPE_TAG = "MsgType="""
Const sDESC_TAG = "MsgDesc="""

    ' Do simple string concatenation here!
    msXMLErrorReport = msXMLErrorReport & sERR_TAG
    msXMLErrorReport = msXMLErrorReport & sTYPE_TAG & enMsgType & """ "
    msXMLErrorReport = msXMLErrorReport & sDESC_TAG & sMsg & """ "
    msXMLErrorReport = msXMLErrorReport & sEND_TAG & vbCrLf
    
End Sub

'---------------------------------------------------------------------
Private Function OKToSaveEForm(oEFI As EFormInstance, ByRef sMsg As String) As Boolean
'---------------------------------------------------------------------
' NCJ 9 Sept 03 - Copied from Batch Validation
' Is it OK to save this eForm?
' Returns TRUE if eForm/Visit dates OK, or False otherwise
' Logs appropriate messages to the Log File
'---------------------------------------------------------------------
Dim oVisitEFI As EFormInstance
Dim oResponse As Response
Dim bBlank As Boolean
Dim oElement As eFormElementRO

    On Error GoTo Errlabel

    bBlank = False 'assume we're ok to begin with
    
    ' Check visit date first
    Set oVisitEFI = oEFI.VisitInstance.VisitEFormInstance
    If Not oVisitEFI Is Nothing Then
        ' Ignore if read-only
        If (Not oVisitEFI.ReadOnly) And (Not oVisitEFI.eForm.EFormDateElement Is Nothing) Then
            Set oElement = oVisitEFI.eForm.EFormDateElement
            Set oResponse = oVisitEFI.Responses.ResponseByElement(oElement)
            bBlank = (oResponse.Value = "") And oResponse.Enterable
            sMsg = "visit"
        End If
    End If
        
    'check eForm date
    'if the visit date isn't missing, and the form has a form date
    If Not bBlank And (Not oEFI.eForm.EFormDateElement Is Nothing) Then
        Set oElement = oEFI.eForm.EFormDateElement
        Set oResponse = oEFI.Responses.ResponseByElement(oElement)
        bBlank = (oResponse.Value = "") And oResponse.Enterable
        sMsg = "eForm"
    End If
    
    If bBlank Then
        'either the form or visit date is missing; decide what to do about it
        sMsg = " This eForm cannot be saved because the " & sMsg & " date is missing."
    End If
    
    ' OK to save if dates not blank
    OKToSaveEForm = Not bBlank
    
    Set oVisitEFI = Nothing
    Set oResponse = Nothing
    Set oElement = Nothing
    
Exit Function
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIInputData.OKToSaveEForm"
    
End Function

'----------------------------------------------------
Private Function CanEnterResponse(oResponse As Response) As Boolean
'----------------------------------------------------
' Returns TRUE if this response is enterable by the user.
' Same as oResponse.Enterable EXCEPT we allow entry into hidden questions here
'----------------------------------------------------
Dim oElement As eFormElementRO

    Set oElement = oResponse.Element
    ' Enterable it if not derived or Not Applicable or locked
    If oResponse.Status = eStatus.NotApplicable _
        Or (oElement.DerivationExpr > "") _
        Or (oResponse.LockStatus <> eLockStatus.lsUnlocked) _
        Or (oElement.DataType = eDataType.Category And Not oElement.ActiveCategories) Then
            CanEnterResponse = False
    Else
            CanEnterResponse = True
    End If
    Set oElement = Nothing
    
End Function

'----------------------------------------------------
Private Sub GetCorrectResponse(oEFI As EFormInstance, oElement As eFormElementRO, _
                                        nCycle As Integer, ByRef oResponse As Response)
'----------------------------------------------------
' DPH 04/01/2006 - Created to allow for creating RQG new rows
' Collect response for given question
' if is a RQG create question firstly
'----------------------------------------------------
Dim oQGI As QGroupInstance

    ' attempt collecting the response directly
    Set oResponse = oEFI.Responses.ResponseByElement(oElement, nCycle)

    'if response is nothing and a RQG question then attempt to create response
    If ((oResponse Is Nothing) And (nCycle > 1)) Then
        'get group instance, create new row, get response again
        Set oQGI = oEFI.QGroupInstanceById(oElement.OwnerQGroup.QGroupID)
        'for www, we force the createnewrow(). this solves the problem that
        'occurs if one of the middle rows of the rqg was blanked - all lower
        'rows would be lost
        Call oQGI.CreateNewRow(True)
        Set oResponse = oEFI.Responses.ResponseByElement(oElement, nCycle)
    End If
    
End Sub
