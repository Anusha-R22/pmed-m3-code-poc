VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLInputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------------------------------------------------------
' File: XMLInputData.cls
' Copyright InferMed Ltd 2003-2004 All Rights Reserved
' Author: Nicky Johns, August 2003
' Purpose: Routines for inputting MACRO subject data from XML
'------------------------------------------------------------------
' REVISIONS
' NCJ 7-13 Aug 2003 - Initial development
' NCJ 26-27 Aug 03 - Do not allow enty into non-enterable fields, locked/frozen visits etc.
' NCJ 3 Sept 03 - Added error message for read-only subject
' NCJ 9 Sept 03 - Added error message for eForm/Visit date missing
' FEB 04 - LISA PHASE II
' NCJ 2 Feb 04 - Allow data entry into hidden fields
' NCJ 18-30 Mar 04 - Further developments
' NCJ 14 Apr 04 - New RQG rows - check for non-empty non-derived data
' NCJ 27 Apr 04 - Tidied up error handling
' NCJ 19 May 04 - Do not do RefreshSkipsAndDerivations when adding question data (improved efficiency)
' NCJ 20 May 04 - Added error message for Invalid User
' NCJ 24 May 04 - Can't complete visit if it is locked or frozen
'------------------------------------------------------------------

Option Explicit

' The overall result
Public Enum eDataInputResult
    Success = 0
    InvalidXML = 41
    SubjectNotExist = 42
    SubjectNotOpened = 43
    DataNotAdded = 44
    ValidationsFired = 45
End Enum

' The individual data error message types
Public Enum eDataInputError
    InvalidXML = 101
    SubjectNotExist = 102
    SubjectNotOpened = 103
    VisitNotExist = 104
    VisitCycleNotExist = 105
    EFormNotExist = 106
    QuestionNotExist = 107
    EFormInUse = 108
    VisitLockedFrozen = 109
    EFormLockedFrozen = 110
    QuestionNotEnterable = 111
    NoVisitEformDate = 112
    NoLockForSave = 113
    ValueRejected = 114
    ValueWarning = 115
    ValueWrongType = 116
End Enum

Private moXMLInputDoc As MSXML2.DOMDocument

Private moUser As MACROUser
Private moSubject As StudySubject

' This manages the lock tokens
Private moLocker As LISALocker

' Collection of error strings
Private msXMLErrorReport As String
' Count any errors
Private mnErrors As Integer

Private msCurSubject As String

' The current visit, eForm and question
Private msCurVisit As String
Private msCurEForm As String
Private msCurQuestion As String

Private mcolQGroupRows As Collection

' Store errors (as opposed to validation warnings)
Private mbThereWereErrors As Boolean

'---------------------------------------------------------------------
Public Function InputLISASubjectData(ByVal sSerialisedUser As String, _
                                ByVal sXMLDataInput As String, _
                                ByVal sLockTokens As String, _
                                ByRef sXMLReport As String) As Long
'---------------------------------------------------------------------
' Given a user and an XML data input string,
' add all the specified data values
' Returns error code as enumerated above (0 = All OK)
' If not all OK, sXMLReport contains error messages
'---------------------------------------------------------------------
Dim sErrMsg As String
Dim sSite As String
Dim sStudyName As String
Dim sSubjLabel As String

    On Error GoTo ErrLabel
    
    mnErrors = 0
    msXMLErrorReport = ""
    mbThereWereErrors = False
    
    ' Trap an Invalid User error
    On Error GoTo InvalidUser
    
    ' Create the MACRO User
    Set moUser = New MACROUser
    Call moUser.SetStateHex(sSerialisedUser)

    ' Go back to generic error handling
    On Error GoTo ErrLabel
    
    If sLockTokens = "" Then
        ' Quick exit if there are no lock tokens
        InputLISASubjectData = glINVALID_LOCKS
    Else
    
        'Load the XML and retrieve the subject details
        If SetXMLRequest(sXMLDataInput, moXMLInputDoc, sStudyName, sSite, sSubjLabel) Then
            Set moSubject = LoadSubject(moUser, sStudyName, sSite, sSubjLabel, sErrMsg)
            If Not moSubject Is Nothing Then
                msCurSubject = sStudyName & "/" & sSubjLabel
                ' Set up the lock tokens
                Set moLocker = New LISALocker
                If moLocker.SetLockTokensString(sLockTokens, moUser.CurrentDBConString, moSubject) Then
                    Call ProcessDataInput(moXMLInputDoc, moSubject)
                    ' Tidy everything up before we leave
                    Call Terminate
                    If mnErrors = 0 Then
                        ' No errors - assume all OK
                        InputLISASubjectData = eDataInputResult.Success
                    ElseIf mbThereWereErrors Then
                        ' We hit some problemettes
                        InputLISASubjectData = eDataInputResult.DataNotAdded
                    Else
                        ' The "errors" were only validation warning messages
                        InputLISASubjectData = eDataInputResult.ValidationsFired
                    End If
                Else
                    ' We didn't set the lock tokens string
                    Call Terminate
                    InputLISASubjectData = glINVALID_LOCKS
                End If
            Else
                ' Return sErrMsg
                AddErrorMsg eDataInputError.SubjectNotOpened, sErrMsg
                InputLISASubjectData = eDataInputResult.SubjectNotOpened
            End If
        Else
            ' Return error in XML
            AddErrorMsg eDataInputError.InvalidXML, "Error in XML data input string"
            InputLISASubjectData = eDataInputResult.InvalidXML
        End If
    
    End If
    sXMLReport = CollectErrors
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLInputData.InputLISASubjectData"

Exit Function
InvalidUser:
    ' Couldn't create user
    InputLISASubjectData = glINVALID_USER_ERR
    sXMLReport = gsINVALID_USER_MSG
    
End Function

'---------------------------------------------------
Private Sub ProcessDataInput(oXMLDoc As MSXML2.DOMDocument, oSubject As StudySubject)
'---------------------------------------------------
' Input all the data values specified in oXMLDoc
'---------------------------------------------------
Dim oVINodeList As MSXML2.IXMLDOMNodeList
Dim i As Integer

    On Error GoTo ErrLabel
'Debug.Print Timer & " START ProcessDataInput "
    
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

'Debug.Print Timer & " END ProcessDataInput "

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLInputData.ProcessDataInput"

End Sub

'---------------------------------------------------
Private Sub Class_Terminate()
'---------------------------------------------------

    Call Terminate
    
End Sub

'---------------------------------------------------
Private Sub Terminate()
'---------------------------------------------------
' Tidy up when we leave
'---------------------------------------------------

    ' Ignore errors here
    On Error Resume Next
    
    Set moXMLInputDoc = Nothing
    Set moUser = Nothing
    
    Set mcolQGroupRows = Nothing
    
    If Not moSubject Is Nothing Then
        Call moSubject.StudyDef.RemoveSubject
        ' Clear up the study def
        Call moSubject.StudyDef.Terminate
        Set moSubject = Nothing
    End If
    
End Sub

'---------------------------------------------------
Private Sub GetVisitNodeData(oElReqVINode As MSXML2.IXMLDOMElement, oSubject As StudySubject)
'---------------------------------------------------
' Process this visit's data
' oElReqVINode is a visit node from the Data Input XML
' NCJ 24 Feb 04 - Use new GetVisitCycle which creates previous cycles as necessary
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
    
    ' NCJ 23 Feb 04 - Get the relevant visit cycle - creating earlier cycles if necessary'
    ' This adds the relevant error messages if it goes wrong
    Set oVI = GetVisitCycle(oSubject, sCode, nCycle)
     
    If Not oVI Is Nothing Then
   
        If oVI.LockStatus <> eLockStatus.lsUnlocked Then
            AddErrorMsg eDataInputError.VisitLockedFrozen, "Visit " & msCurVisit & " is locked or frozen"
        Else
            ' Process the list of its eForms
            Set oEFINodeList = oElReqVINode.getElementsByTagName(gsTAG_EFORM)
            If oEFINodeList.length > 0 Then
                For i = 0 To oEFINodeList.length - 1
                    Call GetEformNodeData(oEFINodeList(i), oVI)
                Next
            Else
                ' No eForms so nothing to do (return error?)
                AddErrorMsg eDataInputError.InvalidXML, "No eForms/questions specified for visit " & msCurVisit
            End If
        End If
        
    End If
    
    Set oVisit = Nothing
    Set oVI = Nothing
    Set oEFINodeList = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLInputData.GetVisitNodeData"

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

    On Error GoTo ErrLabel
    
    ' Get the code and cycle of the required eForm
    Call GetCodeAndCycle(oElReqEFINode, sCode, nCycle)
    msCurEForm = sCode & "[" & nCycle & "]"
    
    Set oEform = oVI.Subject.StudyDef.eFormByCode(sCode)    ' Returns Nothing if not an eForm
    
    ' Can't do it if non-existent eForm
    If oEform Is Nothing Then
        AddErrorMsg eDataInputError.EFormNotExist, "EForm " & msCurEForm & " does not exist"
        Exit Sub
    End If
    
    ' Must be a specific EFI
    Set oEFI = oVI.eFormInstanceByEFormCycle(oEform, nCycle)
    If Not oEFI Is Nothing Then
        If oEFI.LockStatus <> eLockStatus.lsUnlocked Then
            AddErrorMsg eDataInputError.EFormLockedFrozen, "EForm " & msCurEForm & " is locked or frozen"
        Else
            ' Process all its questions
            Call AppendEFormData(oEFI, oElReqEFINode)
        End If
    Else
        AddErrorMsg eDataInputError.EFormNotExist, "EForm " & GetCurEFormSpec & " does not exist"
    End If
    
    Set oEFI = Nothing
    Set oEform = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLInputData.GetEformNodeData"

End Sub

'---------------------------------------------------
Private Sub AppendEFormData(oEFI As EFormInstance, _
                                oElReqEFINode As MSXML2.IXMLDOMElement)
'---------------------------------------------------
' Add all the specified question values for this eForm
'---------------------------------------------------
Dim oQNodeList As MSXML2.IXMLDOMNodeList
Dim i As Integer

    On Error GoTo ErrLabel
    
    ' Get the list of question values
    Set oQNodeList = oElReqEFINode.getElementsByTagName(gsTAG_QUESTION)
    If oQNodeList.length > 0 Then
        ' We want specific questions
'Debug.Print Timer & " Doing data input for " & oEFI.Code
        If LoadEFIResponses(oEFI) Then
            ' NCJ 19 May 04 - Do RefreshSkips here (because no longer done in GetQuestionNodeData)
            ' This populates the derived fields
'Debug.Print Timer & " Doing RefreshSkipsAndDerivations for " & oEFI.Code
            Call oEFI.RefreshSkipsAndDerivations(OpeningEForm, moUser.UserName)
'Debug.Print Timer & " Done RefreshSkipsAndDerivations for " & oEFI.Code
            ' Initialise our collection of groups for the eForm
            Set mcolQGroupRows = New Collection
            For i = 0 To oQNodeList.length - 1
                Call GetQuestionNodeData(oQNodeList(i), oEFI)
            Next
            ' NCJ 19 May 04 - Do RefreshSkips once here (because no longer done in GetQuestionNodeData)
            Call oEFI.RefreshSkipsAndDerivations(Revalidation, moUser.UserName)
            Call SaveEFIResponses(oEFI)
            Call RemoveEFIResponses(oEFI)
        End If
'Debug.Print Timer & " Done data input for " & oEFI.Code
    Else
        ' Nothing to do!
        Call AddErrorMsg(eDataInputError.InvalidXML, "No questions specified for eForm " & GetCurEFormSpec)
    End If

    Set oQNodeList = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLInputData.AppendEFormData"

End Sub

'---------------------------------------------------
Private Sub AddQuestionValue(oResponse As Response, sValue As String)
'---------------------------------------------------
' Add value to this response
'---------------------------------------------------
Dim sErrMsg As String
Dim bChanged As Boolean
Dim sRFC As String
Const sAUTO_RFC = "*** Value changed by LISA application"

    On Error GoTo ErrLabel
'Debug.Print Timer & " Adding question " & oResponse.Element.Code
    
    If oResponse.ValidateValue(sValue, sErrMsg, bChanged) <> eStatus.InvalidData Then
        If bChanged Then
            If oResponse.RequiresValueRFC Then
                sRFC = sAUTO_RFC
            Else
                sRFC = ""
            End If
            Call oResponse.ConfirmValue("", sRFC, "")
            ' NCJ 19 May 04 - Don't do RefreshSkips here because it takes too long!
            ' (and it's not necessary for the LISA data input eForm)
'            Call oResponse.EFormInstance.RefreshSkipsAndDerivations(ChangingResponse, moUser.UserName, oResponse)
        End If
    Else
        AddErrorMsg eDataInputError.ValueRejected, _
                        GetCurQuestionSpec & " value rejected (" & sValue & ") - " & sErrMsg
    End If
    
    ' Add in a message to give the warning
    If oResponse.Status = eStatus.Warning Then
        ' Add error message (but not a "real" error, just a warning)
        Call AddErrorMsg(eDataInputError.ValueWarning, _
                        GetCurQuestionSpec & " warning fired (" & sValue & ") - " & oResponse.ValidationMessage, _
                        False)
    End If
    
'Debug.Print Timer & " Done question " & oResponse.Element.Code

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLInputData.AddQuestionValue"

End Sub

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
Private Sub GetCodeAndCycle(oElNode As MSXML2.IXMLDOMElement, _
                            ByRef sCode As String, _
                            ByRef nCycle As Integer)
'---------------------------------------------------
' Retrieve the Code and Cycle attributes from this element
' Code is returned as LOWER CASE
' Code must be present, but Cycle defaults to 1 if not present
'---------------------------------------------------
Dim vCycle As Variant

    sCode = LCase(oElNode.getAttribute(gsATTR_CODE))   ' Assume there IS one!
    vCycle = oElNode.getAttribute(gsATTR_CYCLE)
    ' If no cycle assume 1
    If IsNull(vCycle) Then
        nCycle = 1
    Else
        ' Assume a valid integer!
        nCycle = CInt(vCycle)
    End If
                            
End Sub

'---------------------------------------------------
Private Sub GetQuestionNodeData(oElReqQNode As MSXML2.IXMLDOMElement, _
                        oEFI As EFormInstance)
'---------------------------------------------------
' oElReqQNode is a question node from the Data Input XML
' oEFI is the current eForm instance
' Assume oEFI has its responses already loaded
'---------------------------------------------------
Dim sCode As String
Dim nCycle As Integer
Dim oResponse As Response
Dim oElement As eFormElementRO
Dim vValue As Variant

    On Error GoTo ErrLabel
    
    Call GetCodeAndCycle(oElReqQNode, sCode, nCycle)
    msCurQuestion = sCode & "[" & nCycle & "]"
    
'Debug.Print Timer & "     Entering value for " & msCurQuestion
    ' Trap for question not existing
    On Error Resume Next
    Set oElement = oEFI.eForm.eFormElementByCode(sCode)
    
    ' Reset error handler
    On Error GoTo ErrLabel
    
    ' Can't do it if question doesn't exist
    If oElement Is Nothing Then
        AddErrorMsg eDataInputError.QuestionNotExist, "Question " & GetCurQuestionSpec & " does not exist"
        Exit Sub
    End If
    
    ' For repeating questions we must deduce the correct cycle no.
    ' i.e. the first "new" row
    nCycle = GetQuestionCycleNo(oElement, oEFI)
    msCurQuestion = sCode & "[" & nCycle & "]"
    Set oResponse = oEFI.Responses.ResponseByElement(oElement, nCycle)
    ' Make sure it exists...
    If Not oResponse Is Nothing Then
        ' Check it's not locked, frozen, derived, N/A etc.
        ' NCJ 2 Feb 04 - Use our own CanEnterResponse rather than oResponse.Enterable (to allow hidden fields)
        If CanEnterResponse(oResponse) Then
            vValue = oElReqQNode.getAttribute(gsATTR_VALUE)
            If IsNull(vValue) Then
                AddErrorMsg eDataInputError.InvalidXML, "No question value given for " & GetCurQuestionSpec
            Else
                Call AddQuestionValue(oResponse, CStr(vValue))
            End If
        Else
            ' Response is not enterable
            AddErrorMsg eDataInputError.QuestionNotEnterable, "Question " & GetCurQuestionSpec & " is not an enterable response"
        End If
    Else
        AddErrorMsg eDataInputError.QuestionNotExist, "Question " & GetCurQuestionSpec & " does not exist"
    End If
    
    Set oElement = Nothing
    Set oResponse = Nothing
    
'Debug.Print Timer & "     DONE entering value for " & msCurQuestion
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLInputData.GetQuestionNodeData"

End Sub

'---------------------------------------------------
Private Function LoadEFIResponses(oEFI As EFormInstance) As Boolean
'---------------------------------------------------
' Load an eForm's responses
' Return TRUE if loaded OK
'---------------------------------------------------
Dim sErrMsg As String
Dim sEFITok As String
Dim sTok2 As String

    On Error GoTo ErrLabel
    
    ' See if we have a lock token for this eForm
    ' (will be empty string if no lock)
    sEFITok = GetOurEFILockToken(oEFI)
    
    If oEFI.VisitInstance.Subject.LoadResponses(oEFI, sErrMsg, sEFITok, sTok2) = lrrReadWrite Then
        LoadEFIResponses = True
    Else
        If sErrMsg = "" Then
            sErrMsg = "Eform " & GetCurEFormSpec & " responses not loaded"
        End If
        Call AddErrorMsg(eDataInputError.EFormInUse, sErrMsg)
        LoadEFIResponses = False
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLInputData.LoadEFIResponses"

End Function

'---------------------------------------------------
Private Sub RemoveEFIResponses(oEFI As EFormInstance)
'---------------------------------------------------
' Remove an eForm's responses
'---------------------------------------------------
Dim bRemoveLock As Boolean

    ' If we own a lock on this eForm, do NOT remove the eForm lock
    ' Otherwise we DO remove the lock
    bRemoveLock = (GetOurEFILockToken(oEFI) = "")
    Call oEFI.VisitInstance.Subject.RemoveResponses(oEFI, bRemoveLock)

End Sub

'---------------------------------------------------
Private Function GetOurEFILockToken(oEFI As EFormInstance) As String
'---------------------------------------------------
' Get our lock token for this EFI, if we have one
' Returns empty string if we don't own a lock on it
'---------------------------------------------------
Dim sEFITok As String
    
    sEFITok = ""
    
    If Not moLocker Is Nothing Then
        ' Get the lock token, if any
        sEFITok = moLocker.EFILockToken(oEFI)
    End If
    
    GetOurEFILockToken = sEFITok
    
End Function

'---------------------------------------------------
Private Sub SaveEFIResponses(oEFI As EFormInstance)
'---------------------------------------------------
' Save an eForm's responses
'---------------------------------------------------
Dim sErrMsg As String

    On Error GoTo ErrLabel
    
'Debug.Print Timer & " SaveEFIResponses " & oEFI.Code
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
        End Select
    Else
        ' Can't save because missing visit/eForm date
        AddErrorMsg eDataInputError.NoVisitEformDate, GetCurEFormSpec & sErrMsg
    End If
  
'Debug.Print Timer & " End SaveEFIResponses " & oEFI.Code
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLInputData.SaveEFIResponses"

End Sub

'---------------------------------------------------
Private Sub AddErrorMsg(enMsgType As eDataInputError, _
                        ByVal sMsg As String, _
                        Optional bIsARealError As Boolean = True)
'---------------------------------------------------
' Add an error message to the collection
' If bIsARealError = False, it's only a question warning
'---------------------------------------------------

    mnErrors = mnErrors + 1
    Call AddXMLErrMsg(enMsgType, sMsg)
    
    If bIsARealError Then
        ' If it's a real error, remember we had one
        mbThereWereErrors = True
    End If
    
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

    On Error GoTo ErrLabel

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
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLInputData.OKToSaveEForm"
    
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
Private Function GetVisitCycle(oSubject As StudySubject, _
                        sVisitCode As String, nCycle As Integer) As VisitInstance
'----------------------------------------------------
' Return the visit instance for this Visit and Cycle
' Create preceeding visit cycles if necessary
' Returns Nothing if can't be done
'----------------------------------------------------
Dim oVisit As VisitRO
Dim oFoundVI As VisitInstance
Dim oLastVI As VisitInstance
Dim nLastVisitCycle As Integer
Dim lVisitId As Long
Dim bStillOK As Boolean

    On Error GoTo ErrLabel
    
    Set oFoundVI = Nothing
    
'Debug.Print Timer & " Finding visit cycle " & nCycle
    Set oVisit = oSubject.StudyDef.VisitByCode(sVisitCode)
    
    If Not oVisit Is Nothing Then
        lVisitId = oVisit.VisitId
        Set oFoundVI = VisitInstanceByIdCycle(oSubject, lVisitId, nCycle)
        bStillOK = True
        nLastVisitCycle = 0
        
        Do While oFoundVI Is Nothing And bStillOK
            Set oLastVI = LastVisitInstance(oSubject, lVisitId)
            ' Is there a visit instance at all?
            If Not oLastVI Is Nothing Then
                ' Check that we've got a new visit
                If oLastVI.CycleNo > nLastVisitCycle Then
                    nLastVisitCycle = oLastVI.CycleNo
                    ' We need to create a new cycle
                    If CompleteVisit(oLastVI) Then
                        ' See if it exists yet
                        Set oFoundVI = VisitInstanceByIdCycle(oSubject, lVisitId, nCycle)
                    Else
                        ' Something's gone wrong
                        bStillOK = False
                    End If
                Else
                    ' We don't have a new visit
                    bStillOK = False
                End If
            Else
                ' There's no initial visit
                bStillOK = False
            End If
        Loop
    Else
        ' Visit code doesn't exist
        Call AddErrorMsg(eDataInputError.VisitNotExist, "Visit '" & sVisitCode & "' does not exist in the study")
        ' Reset the flag so as not to get another error message
        bStillOK = True
    End If
    
    If Not bStillOK Then
        AddErrorMsg eDataInputError.VisitCycleNotExist, "Visit " & msCurVisit & " does not exist"
    End If
    
    Set GetVisitCycle = oFoundVI
    
'Debug.Print Timer & " Got visit cycle "
    Set oVisit = Nothing
    Set oLastVI = Nothing
    Set oFoundVI = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLInputData.GetVisitCycle"
    
End Function

'----------------------------------------------------
Private Function CompleteVisit(oVI As VisitInstance) As Boolean
'----------------------------------------------------
' Complete the given visit by entering dummy data in all Requested eForms
' NCJ 24 May 04 - Don't try to do this if visit is locked/frozen
'----------------------------------------------------
Dim oEFI As EFormInstance
Dim oElement As eFormElementRO
Dim oResponse As Response
Dim colEFIs As Collection
Dim bCompletedOK As Boolean

' The name of the "dummy" question
Const sDUMMYQ = "MAINT_DUMMY"
Const sAE_FORM = "SAE_MAINT_MF"
Const sZERO = "0"

    On Error GoTo ErrLabel
    
    If oVI.LockStatus = eLockStatus.lsUnlocked Then
'Debug.Print Timer & " START Completing visit " & oVI.CycleNo
        bCompletedOK = True
        
        Set colEFIs = oVI.eFormInstances
        For Each oEFI In colEFIs
            If oEFI.Status = eStatus.Requested Then
                ' NCJ 24 May 04 - Check eForm not locked/frozen
                If oEFI.LockStatus = eLockStatus.lsUnlocked Then
                    If oEFI.eForm.Code <> sAE_FORM Then
                        ' If not the AE form then load up the Responses and add dummy data
            
                        If LoadEFIResponses(oEFI) Then
                            Set oElement = oEFI.eForm.eFormElementByCode(sDUMMYQ)
                            If Not oElement Is Nothing Then
                                Set oResponse = oEFI.Responses.ResponseByElement(oElement, 1)
                                If Not oResponse Is Nothing Then
                                    If CanEnterResponse(oResponse) Then
                                        Call AddQuestionValue(oResponse, sZERO)
                                        Call SaveEFIResponses(oEFI)
                                    Else
                                        ' Response not enterable
                                        bCompletedOK = False
                                    End If
                                Else
                                    ' Response doesn't exist
                                    bCompletedOK = False
                                End If
                            Else
                                ' Dummy question doesn't exist
                                bCompletedOK = False
                            End If
                            Call RemoveEFIResponses(oEFI)
                        Else
                            ' Couldn't load responses
                            bCompletedOK = False
                        End If
                    End If
                Else
                    ' NCJ 24 May 04 - EForm is locked or frozen
                    bCompletedOK = False
                End If
            End If
        Next    ' Next requested eForm
    Else
        ' NCJ 24 May 04 - Visit is locked or frozen so can't do anything with it
        bCompletedOK = False
    End If
    
    CompleteVisit = bCompletedOK
    
'Debug.Print Timer & " END Completing visit"
    
    Set colEFIs = Nothing
    Set oEFI = Nothing
    Set oElement = Nothing
    Set oResponse = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLInputData.CompleteVisit"
    
End Function

'----------------------------------------------------
Private Function VisitInstanceByIdCycle(oSubject As StudySubject, _
                    lVisitId As Long, _
                    nCycle As Integer) As VisitInstance
'----------------------------------------------------
' Return VisitInstance with this Id and Cycle
' Returns Nothing if given Id/Cycle doesn't exist
'----------------------------------------------------
Dim colVIs As Collection

    Set colVIs = oSubject.VisitInstancesById(lVisitId)
    
    ' Ignore if cycle doesn't exist
    On Error Resume Next
    ' Collection is indexed by Str(CycleNo)
    Set VisitInstanceByIdCycle = colVIs(Str(nCycle))
    
    Call CollectionRemoveAll(colVIs)
    Set colVIs = Nothing
    
End Function

'-------------------------------------------------------------------------------
Private Function GetQuestionCycleNo(oElement As eFormElementRO, _
                                    oEFI As EFormInstance) As Integer
'-------------------------------------------------------------------------------
' Work out what cycle number we need to use for this question
' If it's in an RQG, we want to start a new row for all the responses in this batch
'-------------------------------------------------------------------------------
Dim nCycle As Integer
Dim oResponse As Response
Dim sQGCode As String

    On Error GoTo ErrLabel
    
    If oElement.OwnerQGroup Is Nothing Then
        ' We do nothing and assume CycleNo = 1
        nCycle = 1
    Else
        sQGCode = oElement.OwnerQGroup.Code
        ' Do we already have it in our collection?
        If CollectionMember(mcolQGroupRows, sQGCode, False) Then
            nCycle = mcolQGroupRows.Item(sQGCode)
        Else
            ' We must find the max. cycle number of any response in this group
            nCycle = 0
            For Each oResponse In oEFI.Responses
                If Not oResponse.Element.OwnerQGroup Is Nothing Then
                    ' Look for non-empty non-derived responses in this group
                    If oResponse.Value > "" And oResponse.Element.DerivationExpr = "" Then
                        If oResponse.Element.OwnerQGroup.Code = sQGCode Then
                            ' It belongs to the same group
                            ' Is it in a higher row than we've already got?
                            If oResponse.RepeatNumber > nCycle Then
                                nCycle = oResponse.RepeatNumber
                            End If
                        End If
                    End If
                End If
            Next
            ' We now go to the next row after the last row containing data
            nCycle = nCycle + 1
            ' And remember the row for this group's future elements
            mcolQGroupRows.Add nCycle, sQGCode
        End If
        
    End If
    
    GetQuestionCycleNo = nCycle
    Set oResponse = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "XMLInputData.GetQuestionCycleNo"
    
End Function
