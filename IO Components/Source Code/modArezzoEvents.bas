Attribute VB_Name = "modArezzoEvents"
'----------------------------------------------------------------------------------------'
' File:         modArezzoEvents
' Copyright:    InferMed Ltd. 2000. All Rights Reserved
' Author:       I Curtis, September 2003
' Purpose:      Contains DHTML/HTML generating code for creating MACRO Arezzo Tasks forms
'               handling actions and decisions
'               Note that this module requres imedalm5 to be added to the project
'----------------------------------------------------------------------------------------'
'   Revisions:
'   ic 03/10/2003 page formatting changes
'   ic 27/02/2004 GetArezzoEformHTML() change to 'next' variable format: id~jsfn~saveflag
'   ic 16/03/2004 remove conditional ORAMA compilation
'   NCJ 31 Mar 04 - Changed "Arguments for" to "Reasons why" (and against to "Reasons why not")
'   ic 23/04/2004 added class 'clsArezzoPage' to control page style in GetArezzoEformHTML()
'   NCJ 26 Apr 04 - Changed text in RtnPlanDescription
'   ic 05/07/2004   added error handling to each routine
'   ic 07/02/2005 added print button and reminder
'   ic 08/02/2005 added subject label, disable function keys
'   ic 04/03/2005 display references inline rather than as a popup
'----------------------------------------------------------------------------------------'

'ic 16/03/2004 remove conditional compilation
'#If ORAMA = 1 Then

Private Const msAREZZO_ACTION_PREFIX = "a_"
Private Const msAREZZO_DECISION_PREFIX = "d_"
Private Const msAREZZO_ENQUIRY_PREFIX = "e_"

'display references inline rather than as a popup
Private Const mbDISPLAY_INLINE_REF = True

Option Explicit

'----------------------------------------------------------------------------------------
Public Function CheckArezzoEvents(ByRef oSubject As StudySubject, ByVal sDatabase As String, _
    ByVal sEformPageTaskId As String, ByVal sNext As String, ByRef bEvents As Boolean) As String
'----------------------------------------------------------------------------------------
'   ic 11/09/2003
'   function checks if there are outstanding arezzo events, and returns the html to
'   allow the user to complete the events if there are any
'   revisions
'   ic 05/07/2004   added error handling
'----------------------------------------------------------------------------------------
Dim colTasks As Collection

    On Error GoTo CatchAllError

    Set colTasks = oSubject.Arezzo.GetArezzoTasks
    If colTasks.Count = 0 Then
        bEvents = False
    Else
        bEvents = True
        CheckArezzoEvents = GetArezzoEformHTML(oSubject, colTasks, sDatabase, sEformPageTaskId, sNext)
    End If
    Set colTasks = Nothing
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modArezzoEvents.CheckArezzoEvents")
End Function

'----------------------------------------------------------------------------------------
Public Sub SaveArezzoEvents(ByRef oSubject As StudySubject, ByVal sForm As String)
'----------------------------------------------------------------------------------------
'   ic 11/09/2003
'   sub saves a users arezzo action/decision responses
'   revisions
'   ic 05/07/2004   added error handling
'----------------------------------------------------------------------------------------
Dim dicTasks As Dictionary
Dim colCandidates As Collection
Dim colTasks As Collection
Dim oTask As TaskInstance
Dim sWWWTaskKey As String
Dim sTaskKey As String
Dim sType As String
Dim n As Integer

    On Error GoTo CatchAllError

    'get collection of outstanding tasks, exit if there are none
    Set colTasks = oSubject.Arezzo.GetArezzoTasks
    If colTasks.Count = 0 Then Exit Sub
    
    'get dictionary of task responses contained in form
    Set dicTasks = RtnTaskDictionary(sForm)
    
    With dicTasks
        For n = 0 To .Count - 1
            'wwwtaskkey is the taskkey, with the msAREZZO_ACTION_PREFIX
            sWWWTaskKey = .Keys(n)
            sTaskKey = CStr(Mid(sWWWTaskKey, 3))
            sType = Left(sWWWTaskKey, 2)
            
            Select Case sType
                Case msAREZZO_ACTION_PREFIX:
                    'action
                    Call oSubject.Arezzo.ConfirmAction(sTaskKey)
    
                Case msAREZZO_DECISION_PREFIX:
                    'decision
                    Set oTask = RtnTaskFromKey(colTasks, sTaskKey)
                    If Not oTask Is Nothing Then
                        If (oTask.IsMultiple And (.Item(sWWWTaskKey).Count > 1)) Then
                            'multi-choice
                            Call oTask.CommitCandidates(.Item(sWWWTaskKey))
                            'if after first commit, task is still 'permitted', we received a warning, ignore it
                            'and try again
                            If oTask.TaskState = "permitted" Then
                                Call oTask.CommitCandidates(.Item(sWWWTaskKey))
                            End If
                        Else
                            'single choice
                            Call oSubject.Arezzo.CommitDecision(sTaskKey, CStr(.Item(sWWWTaskKey)(1)))
                            'if after first commit, task is still 'permitted', we received a warning, ignore it
                            'and try again
                            If oTask.TaskState = "permitted" Then
                                Call oSubject.Arezzo.CommitDecision(sTaskKey, CStr(.Item(sWWWTaskKey)(1)))
                            End If
                        End If
                        Set oTask = Nothing
                    End If
                
                Case msAREZZO_ENQUIRY_PREFIX:
                    'not supported yet
            End Select
        Next
    End With
    'commit to database
    oSubject.Save

    Set dicTasks = Nothing
    Set colCandidates = Nothing
    Set colTasks = Nothing
    Exit Sub
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modArezzoEvents.SaveArezzoEvents")
End Sub

'----------------------------------------------------------------------------------------
Private Function RtnTaskDictionary(ByVal sForm As String) As Dictionary
'----------------------------------------------------------------------------------------
'   ic 12/09/2003
'   function returns a dictionary of collections containing candidates, where the dictionary
'   key is the wwwtaskkey
'   revisions
'   ic 05/07/2004   added error handling
'----------------------------------------------------------------------------------------
Dim dicTasks As Dictionary
Dim colCandidates As Collection
Dim vElements As Variant
Dim vElement As Variant
Dim nLoop As Integer

    On Error GoTo CatchAllError

    Set dicTasks = New Dictionary
    
    'split form string on '&'. get an array of 'fieldname=value'
    vElements = Split(sForm, "&")
    
    With dicTasks
        For nLoop = LBound(vElements) To UBound(vElements)
            'check for prefix, only include arezzo task form fields
            Select Case Left(vElements(nLoop), 2)
                Case msAREZZO_ACTION_PREFIX, msAREZZO_DECISION_PREFIX:
                    'split array element on '='. get an array of 'fieldname' and 'value'
                    vElement = Split(vElements(nLoop), "=")
                    'if the fieldname exists in dictionary, add it to the already created collection
                    'else, create a new collection and add it, then add the collection to dictionary
                    If (.Exists(vElement(0))) Then
                        dicTasks(vElement(0)).Add vElement(1)
                    Else
                        Set colCandidates = New Collection
                        colCandidates.Add vElement(1)
                        .Add vElement(0), colCandidates
                    End If
                
            End Select
        Next
    End With
    Set RtnTaskDictionary = dicTasks
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modArezzoEvents.RtnTaskDictionary")
End Function

'----------------------------------------------------------------------------------------
Private Function RtnTaskFromKey(colTasks As Collection, ByVal sTaskKey As String) As TaskInstance
'----------------------------------------------------------------------------------------
'   ic 11/09/2003
'   function returns a taskinstance given a collections of taskinstances and a search taskkey
'   revisions
'   ic 05/07/2004   added error handling
'----------------------------------------------------------------------------------------
Dim oTask As TaskInstance

    On Error GoTo CatchAllError

    For Each oTask In colTasks
        If (oTask.TaskKey = sTaskKey) Then
            Set RtnTaskFromKey = oTask
            Exit For
        End If
    Next
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modArezzoEvents.RtnTaskFromKey")
End Function

'----------------------------------------------------------------------------------------
Private Function GetArezzoEformHTML(ByRef oSubject As StudySubject, colTasks As Collection, _
    ByVal sDatabase As String, ByVal sEformPageTaskId As String, ByVal sNext As String) As String
'----------------------------------------------------------------------------------------
'   ic 10/09/2003
'   function returns an arrezzo decision and action html page body
'   revisions
'   ic 27/02/2004 change to 'next' variable format: id~jsfn~saveflag
'   ic 23/04/2004 added class 'clsArezzoPage' to control page style
'   ic 05/07/2004   added error handling
'   ic 07/02/2005 added print button and reminder
'   ic 08/02/2005 disable function keys
'----------------------------------------------------------------------------------------
Dim vJSComm() As String

    On Error GoTo CatchAllError

    ReDim vJSComm(0)
    
    'enclose the fields in a form that submits to the same url as the saved eform
    'add a 'next' field first (must be first) to remember where the user wants go
    'and an 'arezzo' field to identify the form as coming from an arezzo event page
    Call AddStringToVarArr(vJSComm, "<body onload='fnPageLoaded();' class='clsArezzoPage'>" _
        & "<form name='FormDE' method='post' action='Eform.asp?fltDb=" & sDatabase & "&fltSt=" _
        & oSubject.StudyId & "&fltSi=" & oSubject.Site & "&fltSj=" & oSubject.PersonId & "&fltId=" _
        & sEformPageTaskId & "'>" _
        & "<input type='hidden' name='next' value='" & sNext & gsDELIMITER3 & gsDELIMITER3 & "0'>" _
        & "<input type='hidden' name='arezzo' value='true'>")
        
    'add a javascript checker function to check all actions/decisions are filled in
    'and a submit function, disable function keys
    Call AddStringToVarArr(vJSComm, "<script language='javascript'>" _
        & "function fnPrintGuideline(){fnCheck(); window.print();}" _
        & "function fnPageLoaded(){fnHideLoader(); fnDisableFKeys(true);}" _
        & GetCheckJS(colTasks) _
        & "function fnOk(){if (fnCheck()) document.FormDE.submit();}" _
        & "</script>")
        
    'start outer table, plan header
    Call AddStringToVarArr(vJSComm, "<table width='90%' align='center'>" _
        & "<tr height='10'><td></td></tr>" _
        & RtnPlanHeader(oSubject, colTasks(1).ParentPlan) _
        & "<tr height='10'><td></td></tr>" _
        & "<tr><td>")
    
    'add decision and action tables
    Call AddStringToVarArr(vJSComm, GetArezzoDecisionHTML(colTasks))
    Call AddStringToVarArr(vJSComm, GetArezzoActionHTML(colTasks))
    
    'print button and reminder
    Call AddStringToVarArr(vJSComm, "</td></tr>" _
        & "<tr height='10'><td></td></tr><tr><td align='right'><table border='0' cellpadding='2'>" _
        & "<tr><td class='clsArezzoPrint' align='right'>" _
        & "Please <a href='javascript:fnPrintGuideline();'>print</a> this form as a record<br>of your choices before continuing</td><td>" _
        & "<a href='javascript:fnPrintGuideline();'><img alt='Print' src='../img/ico_print.gif' border='0'>" _
        & "</a></td></tr></table></td></tr>")
    
    'end outer table, ok button
    Call AddStringToVarArr(vJSComm, "<tr height='10'><td></td></tr>" _
        & "<tr><td align='right'><a href='javascript:fnOk();'>" _
        & "<img alt='Continue' src='../img/ico_nexteform.gif' border='0'>" _
        & "</a></td></tr></table>")
    
    Call AddStringToVarArr(vJSComm, "</form></body>")

    GetArezzoEformHTML = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modArezzoEvents.GetArezzoEformHTML")
End Function

'----------------------------------------------------------------------------------------
Private Function GetCheckJS(ByRef colTasks As Collection) As String
'----------------------------------------------------------------------------------------
'   ic 11/09/2003
'   function returns a javascript checker function that will check the page
'   revisions
'   ic 05/07/2004   added error handling
'----------------------------------------------------------------------------------------
Dim vJSComm() As String
Dim oTask As TaskInstance

    On Error GoTo CatchAllError

    ReDim vJSComm(0)
    Call AddStringToVarArr(vJSComm, "function fnCheck(){" & vbCrLf)
    
    Call AddStringToVarArr(vJSComm, "oF=document.FormDE;" & vbCrLf)
    
    For Each oTask In colTasks
        With oTask
            'calls 'fnSelection(object)' javascript funtion in eform.js
            Select Case .TaskType
                Case "decision":
                    If (.TaskState = "permitted") Then
                        'removed caption for ORAMA
'                        Call AddStringToVarArr(vJSComm, "if (!fnSelection(oF['" & msAREZZO_DECISION_PREFIX & .TaskKey & "']))" _
'                            & "{alert('You have not made a selection for the \'" & ReplaceWithJSChars(.Caption) & "\' decision');return false;}" & vbCrLf)
                        Call AddStringToVarArr(vJSComm, "if (!fnSelection(oF['" & msAREZZO_DECISION_PREFIX & .TaskKey & "']))" _
                            & "{alert('You have not made a selection for all decisions');return false;}" & vbCrLf)
                    End If
                Case "action":
                    'removed caption for ORAMA
'                    Call AddStringToVarArr(vJSComm, "if (!fnSelection(oF['" & msAREZZO_ACTION_PREFIX & .TaskKey & "']))" _
'                            & "{alert('You have not confirmed the \'" & ReplaceWithJSChars(.Caption) & "\' action');return false;}" & vbCrLf)
                    Call AddStringToVarArr(vJSComm, "if (!fnSelection(oF['" & msAREZZO_ACTION_PREFIX & .TaskKey & "']))" _
                        & "{alert('You have not confirmed all actions');return false;}" & vbCrLf)
            End Select
        End With
    Next
    
    Call AddStringToVarArr(vJSComm, "return true;}" & vbCrLf)
    
    GetCheckJS = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modArezzoEvents.GetCheckJS")
End Function

'----------------------------------------------------------------------------------------
Private Function GetArezzoDecisionHTML(ByRef colTasks As Collection) As String
'----------------------------------------------------------------------------------------
'   ic 11/09/2003
'   function returns an html table containing permitted arezzo decisions
'   revisions
'   ic 05/07/2004   added error handling
'   ic 04/03/2005   added inline frame for references
'----------------------------------------------------------------------------------------
Dim vCandidate As Variant
Dim vJSComm() As String
Dim oTask As TaskInstance
Dim sFile As String

    On Error GoTo CatchAllError
    
    ReDim vJSComm(0)
    
    Call AddStringToVarArr(vJSComm, "<table border='0' width='100%'>")
    
    For Each oTask In colTasks
        With oTask
        
            If oTask.TaskType = "decision" Then
                If oTask.TaskState = "permitted" Then
                    'removed ReplaceWithHTMLCodes() around .caption for ORAMA
                    'decision caption
                    Call AddStringToVarArr(vJSComm, "<tr><td class='clsArezzoDecisionCaption'>" _
                    & "<a class='clsArezzoDecisionCaption' >" & RtnCaption(.Caption, sFile) _
                    & "</a></td></tr>" & vbCrLf)
        
'                    'select option caption
'                    Call AddStringToVarArr(vJSComm, "<tr><td class='clsArezzoBestCandidateCaption'>Select option(s)</td></tr>" & vbCrLf)
'
'                    'get best candidates
'                    For Each vCandidate In .BestCandidates
'                        Call AddStringToVarArr(vJSComm, GetArezzoDecisionCandidateHTML(oTask, CStr(vCandidate)))
'                    Next
'
                    'spacer,all options caption
                    Call AddStringToVarArr(vJSComm, "<tr height='5'><td></td></tr>" & vbCrLf)
'                    Call AddStringToVarArr(vJSComm, "<tr><td class=clsArezzoCandidateCaption>All option(s)</td></tr>" & vbCrLf)

                    'get candidates
                    For Each vCandidate In .Candidates
                        Call AddStringToVarArr(vJSComm, GetArezzoDecisionCandidateHTML(oTask, CStr(vCandidate), .BestCandidates))
                    Next
                
                    'inline frame for reference, if any
                    If (sFile <> "") Then
                        Call AddStringToVarArr(vJSComm, "<tr height='5'><td></td></tr>" & vbCrLf)
                        Call AddStringToVarArr(vJSComm, "<tr class='clsArezzoReferenceTitle'><td>Reference</td></tr>" & vbCrLf)
                        Call AddStringToVarArr(vJSComm, "<tr><td><iframe frameborder='0' src='" & sFile & "' " _
                            & "class='clsArezzoRefFrame'></iframe></td></tr>")
                    End If
                    
                    'spacer,line,spacer
                    Call AddStringToVarArr(vJSComm, "<tr height='15'><td></td></tr>" & vbCrLf)
                    Call AddStringToVarArr(vJSComm, "<tr height='2'><td class='clsArezzoSpacer'></td></tr>" & vbCrLf)
                    Call AddStringToVarArr(vJSComm, "<tr height='30'><td></td></tr>" & vbCrLf)
                End If
            End If
        End With
    Next
    
    Call AddStringToVarArr(vJSComm, "</table>")
    
    GetArezzoDecisionHTML = Join(vJSComm, "")
    Set oTask = Nothing
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modArezzoEvents.GetArezzoDecisionHTML")
End Function

'----------------------------------------------------------------------------------------
Private Function RtnCaption(sCaption As String, ByRef sFile As String) As String
'----------------------------------------------------------------------------------------
'   ic 04/13/2005
'   function returns the displayable part of a caption, and the path to the reference file.
'   this assumes that there will always be a 'caption' part, but not necessarily a reference
'
'   eg the caption may be:
'   "What iron therapy should be given? <A class="clsArezzoReference" HREF= "#"
'   onClick="window.open('GuidelineIII2.htm', 'Reference','toolbar=no,width=380,height=380,
'   status=no,scrollbars=yes,resize=no');return false">Reference</A>"
'
'   and the function returns the caption and the filename:
'   "What iron therapy should be given?", "GuidelineIII2.htm"
'
'----------------------------------------------------------------------------------------
Dim sCap As String
Dim sFl As String
Dim nStart As Integer
Dim nStop As Integer

    On Error GoTo CatchAllError
    
    'search for a hyperlink in the caption
    nStart = InStr(sCaption, "<A")
    
    If (nStart > 0) And (mbDISPLAY_INLINE_REF) Then
        'found a hyperlink in caption, split the caption into parts
        'caption part
        sCap = Trim(Mid(sCaption, 1, (nStart - 1)))
        'filename part
        sFl = Trim(Mid(sCaption, nStart))
        sFl = Trim(Mid(sFl, InStr(sFl, ".open('") + 7))
        sFl = Left(sFl, InStr(sFl, "'"))
        sFile = sFl
        
    Else
        'no hyperlink in caption, return the full caption
        sCap = sCaption
        sFile = ""
    End If
    
    RtnCaption = sCap
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modArezzoEvents.RtnCaption")
End Function

'----------------------------------------------------------------------------------------
Private Function GetArezzoDecisionCandidateHTML(ByRef oTask As TaskInstance, ByVal sCandidate As String, _
    ByRef colBestCandidates As Collection) As String
'----------------------------------------------------------------------------------------
'   ic 11/09/2003
'   function returns an html table row containing an arezzo candidate
'   revisions
'   ic 05/07/2004   added error handling
'----------------------------------------------------------------------------------------
Dim colArguments As Collection
Dim vArgument As Variant
Dim vJSComm() As String
Dim bBest As Boolean

    On Error GoTo CatchAllError

    ReDim vJSComm(0)
    
    'is this best candidate
    bBest = IsBestCandidate(sCandidate, colBestCandidates)
    
    With oTask
        'open row and cell
        Call AddStringToVarArr(vJSComm, "<tr><td>")
    
        'if best candidate, add table border with special style
        If (bBest) Then
            Call AddStringToVarArr(vJSComm, "<table cellpadding='0' cellspacing='0' width='100%' " _
                & "class='clsArezzoBestCandidateBorder'><tr height='30'><td>&nbsp;&nbsp;Recommendation</td></tr>" _
                & "<tr><td>")
        End If
    
        'selection control. name='d_[taskkey]', value=[candidate]
        'it is assumed that candidates never contain macro forbidden chars or newline chars
        Call AddStringToVarArr(vJSComm, "<table border='0'><tr><td width='15'><input type='")
        Call AddStringToVarArr(vJSComm, IIf(.IsMultiple, "checkbox", "radio"))
        Call AddStringToVarArr(vJSComm, "' name='" & msAREZZO_DECISION_PREFIX & .TaskKey _
        & "' value='" & sCandidate & "'></td>" & vbCrLf)

        'candidate name
        Call AddStringToVarArr(vJSComm, "<td class='clsArezzoCandidateCaption'>" & .CandidateCaption(sCandidate) _
            & "</td></tr>" & vbCrLf)
        
        'arguments for
        Set colArguments = .Explain(sCandidate, "for")
        If (colArguments.Count > 0) Then
            Call AddStringToVarArr(vJSComm, "<tr><td></td><td>&nbsp;&nbsp;<a class='clsArezzoArgCaption'>Reasons why:</a></td></tr>" _
                & "<tr><td></td><td class='clsArezzoArg'><ul type='square'>")
            For Each vArgument In colArguments
                Call AddStringToVarArr(vJSComm, "<li>" & ReplaceWithHTMLCodes(CStr(vArgument)) & "</li>")
            Next
            Call AddStringToVarArr(vJSComm, "</ul></td></tr>" & vbCrLf)
        End If
        
        'arguments against
        Set colArguments = .Explain(sCandidate, "against")
        If (colArguments.Count > 0) Then
            Call AddStringToVarArr(vJSComm, "<tr><td></td><td>&nbsp;&nbsp;<a class='clsArezzoArgCaption'>Reasons why not:</a></td></tr>" _
                & "<tr><td></td><td class='clsArezzoArg'><ul type='square'>")
            For Each vArgument In colArguments
                Call AddStringToVarArr(vJSComm, "<li>" & ReplaceWithHTMLCodes(CStr(vArgument)) & "</li>")
            Next
            Call AddStringToVarArr(vJSComm, "</ul></td></tr>" & vbCrLf)
        End If
    
'        'arguments confirming
'        Set colArguments = .Explain(sCandidate, "confirming")
'        If (colArguments.Count > 0) Then
'            Call AddStringToVarArr(vJSComm, "<tr><td></td><td>&nbsp;&nbsp;<a class='clsArezzoArgCaption'>Arguments confirming</a></td></tr>" _
'                & "<tr><td></td><td class='clsArezzoArg'><ul type='square'>")
'            For Each vArgument In colArguments
'                Call AddStringToVarArr(vJSComm, "<li>" & ReplaceWithHTMLCodes(CStr(vArgument)) & "</li>")
'            Next
'            Call AddStringToVarArr(vJSComm, "</ul></td></tr>" & vbCrLf)
'        End If
'
'        'arguments excluding
'        Set colArguments = .Explain(sCandidate, "excluding")
'        If (colArguments.Count > 0) Then
'            Call AddStringToVarArr(vJSComm, "<tr><td></td><td>&nbsp;&nbsp;<a class='clsArezzoArgCaption'>Arguments excluding</a></td></tr>" _
'                & "<tr><td></td><td class='clsArezzoArg'><ul type='square'>")
'            For Each vArgument In colArguments
'                Call AddStringToVarArr(vJSComm, "<li>" & ReplaceWithHTMLCodes(CStr(vArgument)) & "</li>")
'            Next
'            Call AddStringToVarArr(vJSComm, "</ul></td></tr>" & vbCrLf)
'        End If

        Call AddStringToVarArr(vJSComm, "</table>")
        
        'close 'best candidate' border
        If (bBest) Then
            Call AddStringToVarArr(vJSComm, "</td></tr></table>")
        End If
        
        'close cell and row
        Call AddStringToVarArr(vJSComm, "</td></tr>")
    End With
    
    GetArezzoDecisionCandidateHTML = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modArezzoEvents.GetArezzoDecisionCandidateHTML")
End Function

'----------------------------------------------------------------------------------------
Private Function IsBestCandidate(ByVal sCandidate As String, ByRef colBestCandidates As Collection) As Boolean
'----------------------------------------------------------------------------------------
'   ic 02/10/2003
'   function returns boolean, whether passed candidate is contained in passed 'best candidates'
'   collection
'   revisions
'   ic 05/07/2004   added error handling
'----------------------------------------------------------------------------------------
Dim vCandidate As Variant
Dim bBest As Boolean

    On Error GoTo CatchAllError

    bBest = False
    For Each vCandidate In colBestCandidates
        If (CStr(vCandidate) = sCandidate) Then
            bBest = True
            Exit For
        End If
    Next

    IsBestCandidate = bBest
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modArezzoEvents.IsBestCandidate")
End Function

'----------------------------------------------------------------------------------------
Private Function GetArezzoActionHTML(ByRef colTasks As Collection) As String
'----------------------------------------------------------------------------------------
'   ic 11/09/2003
'   function returns an html table containing arezzo actions
'   revisions
'   ic 05/07/2004   added error handling
'----------------------------------------------------------------------------------------
Dim vJSComm() As String
Dim oTask As TaskInstance

    On Error GoTo CatchAllError

    ReDim vJSComm(0)
    
    Call AddStringToVarArr(vJSComm, "<table border='0' width='100%'>")
    
    For Each oTask In colTasks
        With oTask
            If .TaskType = "action" Then
                'confirm checkbox. name='a_[taskkey]', value=1
                Call AddStringToVarArr(vJSComm, "<tr><td width='30' rowspan='2'><input type='checkbox' " _
                    & "name='" & msAREZZO_ACTION_PREFIX & .TaskKey & "' value='1'></td>")
            
                'action caption
                Call AddStringToVarArr(vJSComm, "<td class='clsArezzoActionCaption'>" _
                & ReplaceWithHTMLCodes(.Caption) & "</td>")
                
                'procedure
                Call AddStringToVarArr(vJSComm, "<tr><td class='clsArezzoActionProcedure'>" _
                & ReplaceWithHTMLCodes(.Procedure) & "</td></tr>")
                
                'spacer,line,spacer
                Call AddStringToVarArr(vJSComm, "<tr height='15' ><td></td></tr>")
                Call AddStringToVarArr(vJSComm, "<tr height='1'><td class='clsArezzoSpacer' colspan='3'></td></tr>")
                Call AddStringToVarArr(vJSComm, "<tr height='15' ><td></td></tr>")
            End If
        End With
    Next
    
    Call AddStringToVarArr(vJSComm, "</table>")
    
    GetArezzoActionHTML = Join(vJSComm, "")
    Set oTask = Nothing
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modArezzoEvents.GetArezzoActionHTML")
End Function

'----------------------------------------------------------------------------------------
Private Function RtnPlanHeader(ByRef oSubject As StudySubject, ByVal sPlanKey As String)
'----------------------------------------------------------------------------------------
'   ic 24/09/2003
'   function returns plan header rows containing caption and description
'   revisions
'   ic 05/07/2004   added error handling
'   ic 08/02/2005   added subject label
'----------------------------------------------------------------------------------------
Dim vJSComm() As String

    On Error GoTo CatchAllError

    ReDim vJSComm(0)
    
    Call AddStringToVarArr(vJSComm, "<tr height='10'><td class='clsArezzoPlanCaption' align='center'>" _
        & ReplaceWithHTMLCodes(RtnPlanCaption(oSubject, sPlanKey)) & "</td></tr>" _
        & "<tr height='5'><td></td></tr>" _
        & "<tr><td align='center' class='clsArezzoSubjectLabel'>Patient : " & RtnSubjectText(oSubject.PersonId, oSubject.Label) & "</td></tr>" _
        & "<tr height='5'><td></td></tr>" _
        & "<tr><td class='clsArezzoPlanDescription' align='left'>" _
        & ReplaceWithHTMLCodes(RtnPlanDescription(oSubject, sPlanKey)) & "</td></tr>")

    RtnPlanHeader = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modArezzoEvents.RtnPlanHeader")
End Function

'----------------------------------------------------------------------------------------
Private Function RtnPlanCaption(ByRef oSubject As StudySubject, ByVal sPlanKey As String) As String
'----------------------------------------------------------------------------------------
'   ic 15/09/2003
'   function returns the caption of a parent plan
'   revisions
'   ic 05/07/2004   added error handling
'----------------------------------------------------------------------------------------
'Dim oParentPlan As Task

    On Error GoTo CatchAllError

'
'    Set oParentPlan = oSubject.Arezzo.ALM.GuidelineInstance.colTaskInstances.Item(sPlanKey)
'    RtnPlanCaption = oParentPlan.Caption
'    Set oParentPlan = Nothing

    'hard-coded for the ORAMA demo
    RtnPlanCaption = "ORAMA CDS"
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modArezzoEvents.RtnPlanCaption")
End Function

'----------------------------------------------------------------------------------------
Private Function RtnPlanDescription(ByRef oSubject As StudySubject, ByVal sPlanKey As String) As String
'----------------------------------------------------------------------------------------
'   ic 15/09/2003
'   function returns the description of a parent plan
'   revisions
'   NCJ 26 Apr 04 - Changed wording
'   ic 05/07/2004   added error handling
'----------------------------------------------------------------------------------------
'Dim oParentPlan As Task

    On Error GoTo CatchAllError
'
'    Set oParentPlan = oSubject.Arezzo.ALM.GuidelineInstance.colTaskInstances.Item(sPlanKey)
'    Set oParentPlan = Nothing

    'hard-coded for the ORAMA demo
    RtnPlanDescription = "The data on this patient has been processed by ORAMA CDS according to the European " _
        & "Best Practice Guidelines (EBPG) for the Management of Anaemia in Patients with Chronic Renal Failure" _
        & vbCrLf & vbCrLf _
        & "For each guideline, the EBPG recommendations are highlighted, and where available, information is " _
        & "provided in support of these recommendations. The information is specific to this patient." _
        & vbCrLf & vbCrLf _
        & "The purpose of the EBPG recommendations is to support but NOT to replace clinical judgement. Please " _
        & "select the option or options for each question that you believe are appropriate for this patient."
    Exit Function
    
CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modArezzoEvents.RtnPlanDescription")
End Function

'#End If

