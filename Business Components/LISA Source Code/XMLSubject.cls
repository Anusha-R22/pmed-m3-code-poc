VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------------------------------------------------------
' File: XMLSubject.cls
' Copyright InferMed Ltd 2003 All Rights Reserved
' Author: Nicky Johns, August 2003
' Purpose: Routines for creating XML for MACRO subject data
'------------------------------------------------------------------
' REVISIONS
' NCJ 4-7 Aug 2003 - Initial development
' NCJ 27 Aug 03 - Load a subject as read-only when retrieving data
' NCJ 2 Sept 03 - Return XML string if there was an error
' NCJ March 04 - LISA Phase II work
' NCJ 27 Apr 04 - Improved error handling
' NCJ 20 May 04 - Added error message for Invalid User
' NCJ 24 May 04 - Added error message for invalid lock tokens
'------------------------------------------------------------------

Option Explicit

' The overall result
Public Enum eDataRequestResult
    Success = 0
    InvalidXML = 1
    SubjectInUse = 21
    SubjectNotExist = 22
    SubjectNotOpened = 23
End Enum

Private moXMLOutputDoc As MSXML2.DOMDocument
Private moXMLRequestDoc As MSXML2.DOMDocument

'---------------------------------------------------------------------
Public Function GetLISASubjectData(ByVal sSerialisedUser As String, _
                                    ByVal sStudyName As String, _
                                    ByVal sSubjLabel As String, _
                                    ByRef sReturnedData As String, _
                                    ByRef sLockTokens As String) As Long
'---------------------------------------------------------------------
' Given a user and an XML data request string, return the requested data as an XML string
' Function result is an error code. If code = 0, sReturnedData is XML data,
' otherwise sReturnedData is an error message
'   0 - All OK
'   -1 - Unexpected error
'   2 - Subject could not be opened
'---------------------------------------------------------------------
Dim sErrMsg As String
Dim sSite As String
Dim lStudyId As Long
Dim lSubjId As Long
Dim oLocker As LISALocker
Dim oUser As MACROUser

    ' Trap an Invalid User error
    On Error GoTo InvalidUser
    
    ' Create the MACRO User
    Set oUser = New MACROUser
    Call oUser.SetStateHex(sSerialisedUser)

    ' Go back to generic error handling
    On Error GoTo ErrLabel

    ' Identify the subject

'Debug.Print Timer & " Identify subject "
    If GetIDsFromNames(oUser, sStudyName, sSite, sSubjLabel, lStudyId, lSubjId) Then
'    Set moSubject = LoadSubject(oUser, sStudyName, sSite, sSubjLabel, sErrMsg)
'    If Not moSubject Is Nothing Then
'        lSubjId = moSubject.PersonID
'Debug.Print Timer & " Start eForm locking"
        Set oLocker = New LISALocker
        ' Lock all the relevant eForms and get the lock tokens string
        sLockTokens = oLocker.GetLockTokenString(oUser, lStudyId, sSite, lSubjId, sErrMsg)
'Debug.Print Timer & " Done eForm locking"
        If sLockTokens > "" Then
            ' We locked everything OK
'Debug.Print Timer & " Start data retrieval"
            Call RetrieveDataSet(oUser.CurrentDBConString, sStudyName, sSite, lSubjId, sSubjLabel)
            sReturnedData = moXMLOutputDoc.xml
'Debug.Print Timer & " Done data retrieval"
            GetLISASubjectData = eDataRequestResult.Success
        Else
            sReturnedData = sErrMsg
            GetLISASubjectData = eDataRequestResult.SubjectInUse
        End If
        Call oLocker.Terminate
        Set oLocker = Nothing
        
        ' Tidy everything up before we leave
        Call Terminate
    Else
        ' Couldn't load subject
        GetLISASubjectData = eDataRequestResult.SubjectNotExist
        If sErrMsg > "" Then
            sReturnedData = sErrMsg
        Else
            sReturnedData = "Could not find specified subject"
        End If
    End If
    
Exit Function
ErrLabel:
    sReturnedData = Err.Number & ", " & Err.Description & " in XMLSubject.GetLISASubjectData"
    On Error Resume Next
    ' Tidy up
    Call Terminate
    GetLISASubjectData = glERROR_RESULT

Exit Function
InvalidUser:
    ' Couldn't create user
    GetLISASubjectData = glINVALID_USER_ERR
    sReturnedData = gsINVALID_USER_MSG
    
End Function

'---------------------------------------------------
Public Function UnlockSubject(ByVal sSerialisedUser As String, _
                                ByVal sLockTokens As String, _
                                ByRef sMsg As String) As Long
'---------------------------------------------------
' Unlock all the eForms in the sLockTokens string
'---------------------------------------------------
Dim oLocker As LISALocker
Dim oUser As MACROUser

    ' Trap an Invalid User error
    On Error GoTo InvalidUser
    
    ' Create the MACRO User
    Set oUser = New MACROUser
    Call oUser.SetStateHex(sSerialisedUser)

    ' Go back to generic error handling
    On Error GoTo ErrLabel

    If sLockTokens = "" Then
        UnlockSubject = glINVALID_LOCKS
        sMsg = "Empty lock tokens string"
    Else
        Set oLocker = New LISALocker
        UnlockSubject = oLocker.UnlockSubject(sLockTokens, oUser.CurrentDBConString, sMsg)
        Set oLocker = Nothing
    End If
    Set oUser = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|XMLSubject.UnlockSubject"
    
Exit Function
InvalidUser:
    ' Couldn't create user
    UnlockSubject = glINVALID_USER_ERR
    sMsg = gsINVALID_USER_MSG

End Function

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
    Err.Raise Err.Number, , Err.Description & "|" & "XMLSubject.ProcessDataRequest"

End Sub

'---------------------------------------------------
Private Function CreateVisitXML(sVCode As String, nCycleNo As Integer) As MSXML2.IXMLDOMElement
'---------------------------------------------------
' Create visit XML node for this visit
'---------------------------------------------------
Dim oElNode As MSXML2.IXMLDOMElement

    On Error GoTo ErrLabel
    
    Set oElNode = moXMLOutputDoc.createElement(gsTAG_VISIT)
    Call oElNode.setAttribute(gsATTR_CODE, sVCode)
    Call oElNode.setAttribute(gsATTR_CYCLE, CStr(nCycleNo))
    
    Set CreateVisitXML = oElNode
    Set oElNode = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLSubject.CreateVisitXML"

End Function

'---------------------------------------------------
Private Function CreateEFormXML(sEFCode As String, nCycleNo As Integer, _
                            oParentNode As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMElement
'---------------------------------------------------
' Create eForm XML node for this eForm and append it as a child of oParentNode
' Return the new node thus created
'---------------------------------------------------
Dim oElNode As MSXML2.IXMLDOMElement

    On Error GoTo ErrLabel
    
    Set oElNode = moXMLOutputDoc.createElement(gsTAG_EFORM)
    Call oElNode.setAttribute(gsATTR_CODE, sEFCode)
    Call oElNode.setAttribute(gsATTR_CYCLE, CStr(nCycleNo))
    
    Set CreateEFormXML = oParentNode.appendChild(oElNode)
    Set oElNode = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLSubject.CreateEFormXML"

End Function

'---------------------------------------------------
Private Sub CreateQuestionXML(sQCode As String, nCycleNo As Integer, sValue As String, sCatCode As String, _
                            oParentNode As MSXML2.IXMLDOMElement)
'---------------------------------------------------
' Create question XML node for this response and add it as a child to oParentNode
' Returns the newly created node object
'---------------------------------------------------
Dim oElNode As MSXML2.IXMLDOMElement

    On Error GoTo ErrLabel
    
    Set oElNode = moXMLOutputDoc.createElement(gsTAG_QUESTION)
    Call oElNode.setAttribute(gsATTR_CODE, sQCode)
    Call oElNode.setAttribute(gsATTR_CYCLE, CStr(nCycleNo))
    Call oElNode.setAttribute(gsATTR_VALUE, sValue)
    
    ' Is there a category code?
    If sCatCode > "" Then
        Call oElNode.setAttribute(gsATTR_CATCODE, sCatCode)
    End If
    
    Call oParentNode.appendChild(oElNode)
    Set oElNode = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLSubject.CreateQuestionXML"

End Sub

'---------------------------------------------------
Private Sub Class_Terminate()
'---------------------------------------------------

    Call Terminate
    
End Sub

'---------------------------------------------------
Private Sub Terminate()
'---------------------------------------------------

    Set moXMLOutputDoc = Nothing
    Set moXMLRequestDoc = Nothing
    
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
    Set oVisit = oSubject.StudyDef.VisitByCode(sCode)    ' Returns Nothing if not a Visit
    
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
            Set oVINode = AppendVisitData(oVI.Code, oVI.CycleNo)
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
    Err.Raise Err.Number, , Err.Description & "|" & "XMLSubject.GetVisitNodeData"

End Sub

'---------------------------------------------------
Private Function AppendVisitData(sVCode As String, nCycleNo As Integer) As MSXML2.IXMLDOMElement
'---------------------------------------------------
' Append a visit node to the moXMLOutputDoc object for this visit
' Return the XML node thus created
'---------------------------------------------------
Dim oElNode As MSXML2.IXMLDOMElement
Dim oVINode As MSXML2.IXMLDOMElement

    On Error GoTo ErrLabel
            
    ' Create the visit node and append it to the returned XML
    Set oElNode = CreateVisitXML(sVCode, nCycleNo)
    Set oVINode = moXMLOutputDoc.documentElement.appendChild(oElNode)
    Set AppendVisitData = oVINode

    Set oElNode = Nothing
    Set oVINode = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLSubject.AppendVisitData"

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
    Set oEform = oVI.Subject.StudyDef.eFormByCode(sCode)    ' Returns Nothing if not an eForm
    
    ' Can't do it if non-existent eForm
    If oEform Is Nothing Then Exit Sub
    
    If nCycle > 0 Then
        ' Just the one EFI
        Set oEFI = oVI.eFormInstanceByEFormCycle(oEform, nCycle)
        If Not oEFI Is Nothing Then
            ' Add on the node (with all its questions too)
            Call AppendEFormData(oEFI.Code, oEFI.CycleNo, oElRetVINode)
        End If
    Else
        ' Do all cycles of eForm
        nCycle = 1
        Set oEFI = oVI.eFormInstanceByEFormCycle(oEform, nCycle)
        Do While (Not oEFI Is Nothing)
            Call AppendEFormData(oEFI.Code, oEFI.CycleNo, oElRetVINode)
            nCycle = nCycle + 1
            Set oEFI = oVI.eFormInstanceByEFormCycle(oEform, nCycle)
        Loop
    End If
    
    Set oEFI = Nothing
    Set oEform = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLSubject.GetEformNodeData"

End Sub

'---------------------------------------------------
Private Sub AppendEFormData(sEFCode As String, nCycleNo As Integer, _
                                oElRetVINode As MSXML2.IXMLDOMElement)
'---------------------------------------------------
' Append a node for this eForm/XML node to the specified oElRetVINode visit node
' AND add all its specified questions as nodes
'---------------------------------------------------
Dim oEFINode As MSXML2.IXMLDOMElement
Dim oQNodeList As MSXML2.IXMLDOMNodeList
Dim i As Integer

    On Error GoTo ErrLabel
    
    ' Create the eForm node and append it to the returned XML visit node
    Set oEFINode = CreateEFormXML(sEFCode, nCycleNo, oElRetVINode)
    
    Set oEFINode = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLSubject.AppendEFormData"

End Sub

'---------------------------------------------------
Private Sub AppendQuestionData(oResponse As Response, _
                                oElRetEFINode As MSXML2.IXMLDOMElement)
'---------------------------------------------------
' Append a node for this response to the specified oElRetEFINode eForm node
'---------------------------------------------------

    On Error GoTo ErrLabel
    
'    If oResponse.Value > "" Then
'        ' Only do non-empty questions
'        ' Create the question node and append it to the returned XML eForm node
'        Call CreateQuestionXML(oResponse, oElRetEFINode)
'    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLSubject.AppendQuestionData"

End Sub

'---------------------------------------------------
Private Sub GetCodeAndCycle(oElNode As MSXML2.IXMLDOMElement, _
                            ByRef sCode As String, _
                            ByRef nCycle As Integer)
'---------------------------------------------------
' Retrieve the Code and Cycle attributes from this element
' Code is returned as LOWER CASE
' Code must be present, but Cycle defaults to 1 if not present
'---------------------------------------------------
Dim vCycle As Variant

    On Error GoTo ErrLabel
    
    sCode = LCase(oElNode.getAttribute(gsATTR_CODE))   ' Assume there IS one!
    vCycle = oElNode.getAttribute(gsATTR_CYCLE)
    ' If no cycle assume 1
    If IsNull(vCycle) Then
        nCycle = 1
    Else
        ' Assume a valid integer!
        nCycle = CInt(vCycle)
    End If
                            
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLSubject.GetCodeAndCycle"

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
    
    Set oElement = oEFI.eForm.eFormElementByCode(sCode)
    
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
    Err.Raise Err.Number, , Err.Description & "|" & "XMLSubject.GetQuestionNodeData"

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
    
'    For Each oVI In oSubject.VisitInstances
'        ' Only do non-requested visits (i.e. those with data)
'        If oVI.Status <> eStatus.Requested Then
'            Set oVINode = moXMLOutputDoc.documentElement.appendChild(CreateVisitXML(oVI))
'            Call AppendAllVisitForms(oVI, oVINode)
'        End If
'    Next

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLSubject.AppendAllSubjectVisits"

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
'    If oVI.Status <> eStatus.Requested Then
'        For Each oEFI In oVI.eFormInstances
'            ' Only do non-requested eForms (i.e. those with data)
'            If oEFI.Status <> eStatus.Requested Then
'                Set oEFINode = CreateEFormXML(oEFI, oVINode)
'                Call AppendAllFormQuestions(oEFI, oEFINode)
'            End If
'        Next
'    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLSubject.AppendAllVisitForms"

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
'    If oEFI.Status <> eStatus.Requested Then
'        Call LoadEFIResponses(oEFI)
'        For Each oResponse In oEFI.Responses
'            ' Only do non-empty questions
'            If oResponse.Value > "" Then
'                Call CreateQuestionXML(oResponse, oEFINode)
'            End If
'        Next
'        Call RemoveEFIResponses(oEFI)
'    End If
    
    Set oResponse = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLSubject.AppendAllFormQuestions"
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

'---------------------------------------------------
Private Sub RetrieveDataSet(sDBConn As String, _
                            sStudy As String, sSite As String, lPersonId As Long, _
                            sSubjLabel As String)
'---------------------------------------------------
' Retrieve all the data for this subject using SQL View
' sEFormList is comma-separated list of single-quoted eForms
' sQuestionList is comma-separated list of single-quoted questions
'---------------------------------------------------
Dim rsDataSet As ADODB.Recordset
Dim oDBCon As ADODB.Connection
Dim sSQL As String
Dim sLastVisit As String
Dim sLastEForm As String
Dim oVINode As MSXML2.IXMLDOMElement
Dim oEFINode As MSXML2.IXMLDOMElement
Dim sEFormList As String
Dim sQuestionList As String

    On Error GoTo ErrHandler
    
    ' Initialise the XML output doc
    Call InitXMLDoc(sStudy, sSubjLabel)
    
    ' Read in the eForm and question lists from the Settings file
    Call EFormsQuestionsList(sEFormList, sQuestionList)
    
    Set oDBCon = New ADODB.Connection
    Call oDBCon.Open(sDBConn)
    oDBCon.CursorLocation = adUseClient

    Set rsDataSet = New ADODB.Recordset
    sSQL = "SELECT VISIT,VISITCYCLE,EFORM,QUESTION,QUESTIONCYCLE,RESPONSE,RESPONSECODE " _
            & " FROM QUESTIONS" _
            & " WHERE Study = '" & sStudy & "'" _
            & " AND Site = '" & sSite & "'" _
            & " AND SubjectId = " & lPersonId _
            & " AND EFORM IN ( " & sEFormList & ")" _
            & " AND QUESTION IN ( " & sQuestionList & ")" _
            & " ORDER BY VISIT,VISITCYCLE,EFORM,QORDER,QUESTIONCYCLE,RQGORDER"
    rsDataSet.Open sSQL, oDBCon, adOpenKeyset, adLockReadOnly, adCmdText

    ' Now we read off the data and generate the XML into moXMLOutputDoc
    sLastVisit = ""
    Do While Not rsDataSet.EOF
        ' Ignore empty responses
        If IsNull(rsDataSet!Response) Or rsDataSet!Response = "" Then
            ' Do nothing
        Else
            If rsDataSet!Visit & rsDataSet!VISITCYCLE <> sLastVisit Then
                ' New visit
                sLastVisit = rsDataSet!Visit & rsDataSet!VISITCYCLE
                sLastEForm = rsDataSet!eForm
                Set oVINode = AppendVisitData(rsDataSet!Visit, rsDataSet!VISITCYCLE)
                Set oEFINode = CreateEFormXML(rsDataSet!eForm, 1, oVINode)
            End If
            
            If rsDataSet!eForm <> sLastEForm Then
                'New eForm
                sLastEForm = rsDataSet!eForm
                Set oEFINode = CreateEFormXML(rsDataSet!eForm, 1, oVINode)
            End If
            
            ' Now do the question
            Call CreateQuestionXML(rsDataSet!QUESTION, rsDataSet!QUESTIONCYCLE, _
                                    rsDataSet!Response, ConvertFromNull(rsDataSet!RESPONSECODE, vbString), _
                                    oEFINode)
        End If
        rsDataSet.MoveNext
    Loop
    
    Call rsDataSet.Close
    Set rsDataSet = Nothing
    Call oDBCon.Close
    Set oDBCon = Nothing
    Set oEFINode = Nothing
    Set oVINode = Nothing
    
Exit Sub
ErrHandler:

End Sub

'---------------------------------------------------
Private Sub InitXMLDoc(sStudy As String, sSubjLabel As String)
'---------------------------------------------------
' Initialise XML output for this subject
'---------------------------------------------------

    Set moXMLOutputDoc = Nothing
    Set moXMLOutputDoc = New MSXML2.DOMDocument
    Call moXMLOutputDoc.loadXML(gsXML_VERSION_HEADER & gsXML_EMPTY_SUBJ)
    ' Set the subject properties
    With moXMLOutputDoc.documentElement
        Call .setAttribute(gsATTR_STUDY, sStudy)
        Call .setAttribute(gsATTR_LABEL, sSubjLabel)
    End With
    
End Sub

