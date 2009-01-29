VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APIGetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
' File: APIGetData.cls
' Copyright InferMed Ltd 2003 All Rights Reserved
' Author: Nicky Johns, August 2003
' Purpose: Routines for creating XML for retrieved MACRO subject data
'------------------------------------------------------------------
' REVISIONS
' NCJ 4-7 Aug 2003 - Initial development
' NCJ 27 Aug 03 - Load a subject as read-only when retrieving data
' NCJ 2 Sept 03 - Return XML string if there was an error
' NCJ 2 Feb 04 - This file created from original XMLSubject.cls
' TA 24/05/2005 - allow trialid, visitid,eformid, and crflementid to be used
' NCJ 18 Aug 06 - Moved GetCodeAndCycle to modAPIUtils
'------------------------------------------------------------------

Option Explicit

Private moXMLOutputDoc As MSXML2.DOMDocument
Private moXMLRequestDoc As MSXML2.DOMDocument

Private moUser As MACROUser
Private moSubject As StudySubject

'---------------------------------------------------------------------
Public Function GetXMLSubjectData(ByVal sSerialisedUser As String, _
                                    ByVal sXMLDataRequest As String, _
                                    ByRef sReturnedData) As eDataRequestResult
'---------------------------------------------------------------------
' Given a user and an XML data request string, return the requested data as an XML string
' Function result is an error code. If code = 0, sReturnedData is XML data,
' otherwise sReturnedData is an error message
'   0 - All OK
'   1 - XMLDatarequest not understood
'   2 - Subject could not be opened
'---------------------------------------------------------------------
Dim sErrMsg As String
Dim sSite As String
Dim sStudyName As String
Dim sSubjLabel As String

    On Error GoTo ErrLabel
    
    ' Create the MACRO User
    Set moUser = New MACROUser
    Call moUser.SetStateHex(sSerialisedUser)

    'Load the XML and retrieve the subject details
    If SetXMLRequest(sXMLDataRequest, moXMLRequestDoc, sStudyName, sSite, sSubjLabel) Then
        ' Load subject as read-only (for better speed)
        Set moSubject = LoadSubject(moUser, sStudyName, sSite, sSubjLabel, sErrMsg, True)
        If Not moSubject Is Nothing Then
            Call InitSubject(moSubject)
            Call ProcessDataRequest(moXMLRequestDoc, moSubject)
            sReturnedData = moXMLOutputDoc.xml
            ' Tidy everything up before we leave
            Call Terminate
            GetXMLSubjectData = eDataRequestResult.Success
        Else
            sReturnedData = sErrMsg
            GetXMLSubjectData = eDataRequestResult.SubjectNotOpened
        End If
    Else
        ' Return the XML string so we can see what was wrong
        sReturnedData = "Error in XML data request - " & sXMLDataRequest
        GetXMLSubjectData = eDataRequestResult.InvalidXML
    End If
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.GetXMLSubjectData"

End Function

'---------------------------------------------------
Private Sub InitSubject(oSubject As StudySubject)
'---------------------------------------------------
' Initialise XML output for this subject
'---------------------------------------------------

    Set moXMLOutputDoc = Nothing
    Set moXMLOutputDoc = New MSXML2.DOMDocument
    Call moXMLOutputDoc.loadXML(gsXML_VERSION_HEADER & gsXML_EMPTY_SUBJ)
    ' Set the subject properties
    With moXMLOutputDoc.documentElement
        Call .setAttribute(gsATTR_STUDY, oSubject.StudyCode)
        Call .setAttribute(gsATTR_SITE, oSubject.Site)
        Call .setAttribute(gsATTR_LABEL, oSubject.Label)
    End With
    
End Sub

'---------------------------------------------------
Private Sub ProcessDataRequest(oXMLDoc As MSXML2.DOMDocument, oSubject As StudySubject)
'---------------------------------------------------
' Pick out the requested visits, eForms and questions from a data request in oXMLDoc
'---------------------------------------------------
Dim oVINodeList As MSXML2.IXMLDOMNodeList
Dim i As Integer

    On Error GoTo ErrLabel
    
    ' Get the Visit nodes
    Set oVINodeList = oXMLDoc.getElementsByTagName(gsTAG_VISIT)
    If oVINodeList.length > 0 Then
        For i = 0 To oVINodeList.length - 1
            Call GetVisitNodeData(oVINodeList(i), oSubject)
        Next
    Else
        ' We want all the visits
        Call AppendAllSubjectVisits(oSubject)
    End If

    Set oVINodeList = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.ProcessDataRequest"

End Sub

'---------------------------------------------------
Private Function CreateVisit(oVisit As VisitInstance) As MSXML2.IXMLDOMElement
'---------------------------------------------------
' Create visit XML node for this visit
'---------------------------------------------------
Dim oElNode As MSXML2.IXMLDOMElement

    On Error GoTo ErrLabel
    
    Set oElNode = moXMLOutputDoc.createElement(gsTAG_VISIT)
    Call oElNode.setAttribute(gsATTR_CODE, oVisit.Code)
    Call oElNode.setAttribute(gsATTR_CYCLE, CStr(oVisit.CycleNo))
    
    Set CreateVisit = oElNode
    Set oElNode = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.CreateVisit"

End Function

'---------------------------------------------------
Private Function CreateEForm(oEFI As EFormInstance, _
                            oParentNode As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMElement
'---------------------------------------------------
' Create eForm XML node for this eForm and append it as a child of oParentNode
' Return the new node thus created
'---------------------------------------------------
Dim oElNode As MSXML2.IXMLDOMElement

    On Error GoTo ErrLabel
    
    Set oElNode = moXMLOutputDoc.createElement(gsTAG_EFORM)
    Call oElNode.setAttribute(gsATTR_CODE, oEFI.Code)
    Call oElNode.setAttribute(gsATTR_CYCLE, CStr(oEFI.CycleNo))
    
    Set CreateEForm = oParentNode.appendChild(oElNode)
    Set oElNode = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.CreateEForm"

End Function

'---------------------------------------------------
Private Function CreateQuestion(oResponse As Response, _
                            oParentNode As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMElement
'---------------------------------------------------
' Create question XML node for this response and add it as a child to oParentNode
' Returns the newly created node object
'---------------------------------------------------
Dim oElNode As MSXML2.IXMLDOMElement

    On Error GoTo ErrLabel
    
    Set oElNode = moXMLOutputDoc.createElement(gsTAG_QUESTION)
    Call oElNode.setAttribute(gsATTR_CODE, oResponse.Element.Code)
    Call oElNode.setAttribute(gsATTR_CYCLE, CStr(oResponse.RepeatNumber))
    Call oElNode.setAttribute(gsATTR_VALUE, oResponse.Value)
    
    Set CreateQuestion = oParentNode.appendChild(oElNode)
    Set oElNode = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.CreateQuestion"

End Function

'---------------------------------------------------
Private Sub Class_Terminate()
'---------------------------------------------------

    Call Terminate
    
End Sub

'---------------------------------------------------
Private Sub Terminate()
'---------------------------------------------------

    On Error GoTo ErrLabel
    
    Set moXMLOutputDoc = Nothing
    Set moXMLRequestDoc = Nothing
    Set moUser = Nothing
    
    If Not moSubject Is Nothing Then
        Call moSubject.StudyDef.RemoveSubject
        ' Clear up the study def
        Call moSubject.StudyDef.Terminate
        Set moSubject = Nothing
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.Terminate"

End Sub

'---------------------------------------------------
Private Sub GetVisitNodeData(oElReqVINode As MSXML2.IXMLDOMElement, oSubject As StudySubject)
'---------------------------------------------------
' Add nodes to the moXMLOutputDoc object representing this visit's data
' oElReqVINode is a visit node from the Data Request XML (maybe representing a number of visit cycles)
'---------------------------------------------------
Dim sCode As String
Dim nCycle As Integer
Dim oVI As VisitInstance
Dim oVisit As VisitRO
Dim bDoneVisit As Boolean
Dim bFoundVisit As Boolean
Dim oEFINodeList As MSXML2.IXMLDOMNodeList
Dim oVINode As MSXML2.IXMLDOMElement
Dim i As Integer

    Call GetCodeAndCycle(oElReqVINode, sCode, nCycle)
    
    bDoneVisit = False
    bFoundVisit = False
    
    'TA 24/05/2005
    'assume that if visit code is numeric then it is an id
    If IsNumeric(sCode) Then
        Set oVisit = oSubject.StudyDef.VisitById(CLng(sCode))
    Else
        Set oVisit = oSubject.StudyDef.VisitByCode(sCode)    ' Returns Nothing if not a Visit
    End If
    
    ' Get out now if visit doesn't exist
    If oVisit Is Nothing Then Exit Sub
        
    For Each oVI In oSubject.VisitInstancesById(oVisit.VisitId)
        bFoundVisit = False
        If nCycle > 0 Then
            ' Is it the right cycle?
            bFoundVisit = (oVI.CycleNo = nCycle)
            bDoneVisit = bFoundVisit
        Else
            ' We want all cycles so we want this one
            bFoundVisit = True
        End If
        If bFoundVisit Then
            ' We wanted this one
            Set oVINode = AppendVisitData(oVI)
            ' Now see what we must do about its eForms
            Set oEFINodeList = oElReqVINode.getElementsByTagName(gsTAG_EFORM)
            If oEFINodeList.length > 0 Then
                ' We want specific eForms
                For i = 0 To oEFINodeList.length - 1
                    Call GetEformNodeData(oEFINodeList(i), oVI, oVINode)
                Next
            Else
                ' We want ALL eForms and questions
                Call AppendAllVisitForms(oVI, oVINode)
            End If
        End If
        
        ' Are we done?
        If bDoneVisit Then Exit For
    Next
    
    Set oVisit = Nothing
    Set oVI = Nothing
    Set oEFINodeList = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.GetVisitNodeData"

End Sub

'---------------------------------------------------
Private Function AppendVisitData(oVI As VisitInstance) As MSXML2.IXMLDOMElement
'---------------------------------------------------
' Append a visit node to the moXMLOutputDoc object for this visit
' Return the XML node thus created
'---------------------------------------------------
Dim oElNode As MSXML2.IXMLDOMElement
Dim oVINode As MSXML2.IXMLDOMElement

    On Error GoTo ErrLabel
            
    If oVI.Status <> eStatus.Requested Then
        ' Only do non-requested visits ('cos requested ones have no data)
        ' Create the visit node and append it to the returned XML
        Set oElNode = CreateVisit(oVI)
        Set oVINode = moXMLOutputDoc.documentElement.appendChild(oElNode)
        Set AppendVisitData = oVINode
    End If

    Set oElNode = Nothing
    Set oVINode = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.AppendVisitData"

End Function

'---------------------------------------------------
Private Sub GetEformNodeData(oElReqEFINode As MSXML2.IXMLDOMElement, _
                        oVI As VisitInstance, _
                        oElRetVINode As MSXML2.IXMLDOMElement)
'---------------------------------------------------
' Add nodes to the oElRetVINode (visit data node) representing this eForm node's data
' oElReqEFINode is an eForm node from the Data Request XML (maybe representing a number of eForm cycles)
' and oVI is the current visit instance
'---------------------------------------------------
Dim sCode As String
Dim nCycle As Integer
Dim oEFI As EFormInstance
Dim oEform As eFormRO
Dim i As Integer

    On Error GoTo ErrLabel
    
    ' Get the code and cycle of the required eForm
    Call GetCodeAndCycle(oElReqEFINode, sCode, nCycle)
    
    
    'TA 24/05/2005
    'assume is an id if code is numeric
    If IsNumeric(sCode) Then
        Set oEform = oVI.Subject.StudyDef.eFormById(CLng(sCode))
    Else
        Set oEform = oVI.Subject.StudyDef.eFormByCode(sCode)    ' Returns Nothing if not an eForm
    End If
    ' Can't do it if non-existent eForm
    If oEform Is Nothing Then Exit Sub
    
    If nCycle > 0 Then
        ' Just the one EFI
        Set oEFI = oVI.eFormInstanceByEFormCycle(oEform, nCycle)
        If Not oEFI Is Nothing Then
            ' Add on the node (with all its questions too)
            Call AppendEFormData(oEFI, oElReqEFINode, oElRetVINode)
        End If
    Else
        ' Do all cycles of eForm
        nCycle = 1
        Set oEFI = oVI.eFormInstanceByEFormCycle(oEform, nCycle)
        Do While (Not oEFI Is Nothing)
            Call AppendEFormData(oEFI, oElReqEFINode, oElRetVINode)
            nCycle = nCycle + 1
            Set oEFI = oVI.eFormInstanceByEFormCycle(oEform, nCycle)
        Loop
    End If
    
    Set oEFI = Nothing
    Set oEform = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.GetEformNodeData"

End Sub

'---------------------------------------------------
Private Sub AppendEFormData(oEFI As EFormInstance, _
                                oElReqEFINode As MSXML2.IXMLDOMElement, _
                                oElRetVINode As MSXML2.IXMLDOMElement)
'---------------------------------------------------
' Append a node for this eForm/XML node to the specified oElRetVINode visit node
' AND add all its specified questions as nodes
'---------------------------------------------------
Dim oEFINode As MSXML2.IXMLDOMElement
Dim oQNodeList As MSXML2.IXMLDOMNodeList
Dim i As Integer

    On Error GoTo ErrLabel
    
    If oEFI.Status <> eStatus.Requested Then
        ' Only do non-requested eForms ('cos requested ones have no data)
        ' Create the eForm node and append it to the returned XML visit node
        Set oEFINode = CreateEForm(oEFI, oElRetVINode)
        ' Now see what we must do about its questions
        Set oQNodeList = oElReqEFINode.getElementsByTagName(gsTAG_QUESTION)
        If oQNodeList.length > 0 Then
            ' We want specific questions
            Call LoadEFIResponses(oEFI)
            For i = 0 To oQNodeList.length - 1
                Call GetQuestionNodeData(oQNodeList(i), oEFI, oEFINode)
            Next
            Call RemoveEFIResponses(oEFI)
        Else
            ' We want ALL questions (this loads/removes responses)
            Call AppendAllFormQuestions(oEFI, oEFINode)
        End If
    End If

    Set oEFINode = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.AppendEFormData"

End Sub

'---------------------------------------------------
Private Sub AppendQuestionData(oResponse As Response, _
                                oElRetEFINode As MSXML2.IXMLDOMElement)
'---------------------------------------------------
' Append a node for this response to the specified oElRetEFINode eForm node
'---------------------------------------------------

    On Error GoTo ErrLabel
    
    If oResponse.Value > "" Then
        ' Only do non-empty questions
        ' Create the question node and append it to the returned XML eForm node
        Call CreateQuestion(oResponse, oElRetEFINode)
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.AppendQuestionData"

End Sub

'---------------------------------------------------
Private Sub GetQuestionNodeData(oElReqQNode As MSXML2.IXMLDOMElement, _
                        oEFI As EFormInstance, _
                        oElRetEFINode As MSXML2.IXMLDOMElement)
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

    On Error GoTo ErrLabel
    
    Call GetCodeAndCycle(oElReqQNode, sCode, nCycle)
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
    On Error GoTo ErrLabel
    
    ' Can't do it if question doesn't exist
    If oElement Is Nothing Then Exit Sub
    
    If nCycle > 0 Then
        ' Just the one question response
        Set oResponse = oEFI.Responses.ResponseByElement(oElement, nCycle)
        ' Make sure it exists...
        If Not oResponse Is Nothing Then
            Call AppendQuestionData(oResponse, oElRetEFINode)
        End If
    Else
        ' Get all repeats of question
        nCycle = 1
        Set oResponse = oEFI.Responses.ResponseByElement(oElement, nCycle)
        Do While (Not oResponse Is Nothing)
            Call AppendQuestionData(oResponse, oElRetEFINode)
            nCycle = nCycle + 1
            Set oResponse = oEFI.Responses.ResponseByElement(oElement, nCycle)
        Loop
    End If
    
    Set oElement = Nothing
    Set oResponse = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.GetQuestionNodeData"

End Sub

'---------------------------------------------------
Private Sub AppendAllSubjectVisits(oSubject As StudySubject)
'---------------------------------------------------
' Create all the non-requested visit XML nodes for this subject,
' including all the non-requested eForms, and their non-empty questions
' Attach the nodes to the top-level moXMLOutputDoc
'---------------------------------------------------
Dim oVI As VisitInstance
Dim oVINode As MSXML2.IXMLDOMNode

    On Error GoTo ErrLabel
    
    For Each oVI In oSubject.VisitInstances
        ' Only do non-requested visits (i.e. those with data)
        If oVI.Status <> eStatus.Requested Then
            Set oVINode = moXMLOutputDoc.documentElement.appendChild(CreateVisit(oVI))
            Call AppendAllVisitForms(oVI, oVINode)
        End If
    Next

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.AppendAllSubjectVisits"

End Sub

'---------------------------------------------------
Private Sub AppendAllVisitForms(oVI As VisitInstance, oVINode As MSXML2.IXMLDOMElement)
'---------------------------------------------------
' Create all the eForm XML nodes for this visit instance,
' Including only non-requested eForms, and their non-empty questions
' Attach the nodes to the XML node oVINode
'---------------------------------------------------
Dim oEFI As EFormInstance
Dim oEFINode As MSXML2.IXMLDOMNode

    On Error GoTo ErrLabel
    
    ' Only do non-requested visits (i.e. those with data)
    If oVI.Status <> eStatus.Requested Then
        For Each oEFI In oVI.eFormInstances
            ' Only do non-requested eForms (i.e. those with data)
            If oEFI.Status <> eStatus.Requested Then
                Set oEFINode = CreateEForm(oEFI, oVINode)
                Call AppendAllFormQuestions(oEFI, oEFINode)
            End If
        Next
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.AppendAllVisitForms"

End Sub

'---------------------------------------------------
Private Sub AppendAllFormQuestions(oEFI As EFormInstance, oEFINode As MSXML2.IXMLDOMElement)
'---------------------------------------------------
' Create all the non-empty question XML nodes for this eForm instance
' Attach the nodes to oEFINode
' This does the necessary Load/Remove responses
'---------------------------------------------------
Dim oResponse As Response

    On Error GoTo ErrLabel
    
    ' Only do a non-requested eForm (i.e. one with data)
    If oEFI.Status <> eStatus.Requested Then
        Call LoadEFIResponses(oEFI)
        For Each oResponse In oEFI.Responses
            ' Only do non-empty questions
            If oResponse.Value > "" Then
                Call CreateQuestion(oResponse, oEFINode)
            End If
        Next
        Call RemoveEFIResponses(oEFI)
    End If
    
    Set oResponse = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APIGetData.AppendAllFormQuestions"
End Sub

'---------------------------------------------------
Private Sub LoadEFIResponses(oEFI As EFormInstance)
'---------------------------------------------------
' Load an eForm's responses
'---------------------------------------------------
Dim sErrMsg As String
Dim sTok1 As String
Dim sTok2 As String

    Call oEFI.VisitInstance.Subject.LoadResponses(oEFI, sErrMsg, sTok1, sTok2)

End Sub

'---------------------------------------------------
Private Sub RemoveEFIResponses(oEFI As EFormInstance)
'---------------------------------------------------
' Remove an eForm's responses
'---------------------------------------------------

    Call oEFI.VisitInstance.Subject.RemoveResponses(oEFI, True)

End Sub

