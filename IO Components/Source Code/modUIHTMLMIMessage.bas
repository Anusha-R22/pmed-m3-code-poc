Attribute VB_Name = "modUIHTMLMIMessage"
'----------------------------------------------------------------------------------------'
'   File:       modUIHTML.bas
'   Copyright:  InferMed Ltd. 2002-2006. All Rights Reserved
'   Author:     i curtis 02/2003
'   Purpose:    functions returning html versions of MACRO pages (MIMESSAGES)
'----------------------------------------------------------------------------------------'
'   revisions
' ic 30/05/2003 moved error display code out of conditional statement in GetMIMessageList(),
'               now will always display errors
' ic 05/06/2003 added extra 'refresh z-order' parameter in GetMIMessageList()
' ic 06/05/2003 added ReplaceWithJSChars() call around sSubjectLabel in GetMIMessageList()
' ic 27/08/2003 moved mimessage code from clsWWW
' ic 02/09/2003 comment out bChangeData in GetMIMessageList() - not used
' dph 02/09/2003 show full username for mimessages in GetMIMessageList
' dph 13/10/2003 improve SQL performance in RtnMIMessageList
' ic 01/03/2004 in RaiseMIMessage() added SDV error feedback
' ic 05/03/2004 in GetMIMessageList, added subjectid to fnM() calls for subject locking on updates
' ic 10/05/2004 added oSubject parameter so MIMessage status can be updated in 'in-memory' subject in RaiseMIMessage()
' ic 29/06/2004 added error handling
' ic 12/07/2004 reinstated sdv fix
' ic 16/07/2004 fixed locking during sdv update
' ic 24/08/2004 added PlannedSDVsToDone function
' ic 09/11/2004 bug 2400, pass responsecycle to MIMessageExists() function in RaiseMIMessage()
' ic 20/12/2004 moved PlannedSDVsToDone() to clsWWW to avoid cycling reference
' ic 01/04/2005 issue 2541 added show existing sdv code and arguments to RaiseMIMessage()
' ic 01/04/2005 issue 2508, top aligned all rows in GetMIMessageList()
' ic 27/04/2005 issue 2222, added permission code to GetMIMessageList() and UpdateMIMessage()
' ic 28/04/2005 issue 2431, added a print button in GetMIMessageList()
' ic 11/05/2005 issue 2571, added row header count parameter to javascript fnOnClick() calls in GetMIMessageList()
' ic 04/07/2005 issue 2464, added visit, eform, question cycle, fully qualified eSDVMIMStatus enum
' NCJ 8-19 Dec 05 - New date/time types
' ic 18/08/2006 issue 2782, removed erroneous 'Repeat' argument from function GetMIMessageList()
' NCJ 29 Dec 06 - Bug 2861 - Check response exists before creating an MIMessage
' ic 27/02/2007 issue 2114, added GMT to timestamps, added cycle numbers to repeating eforms and visits
'----------------------------------------------------------------------------------------'

Option Explicit

'--------------------------------------------------------------------------------------------------
Public Sub SaveComment(ByRef oUser As MACROUser, _
                        ByVal sAddInfo As String, _
                        ByRef oResponse As Response)
'--------------------------------------------------------------------------------------------------
'   ic 10/12/2002
'   sub saves a response's comment
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim nInfoLoop As Integer
Dim nItemLoop As Integer
Dim vAddInfo As Variant
Dim vAddItem As Variant
Dim sAddItem As String

    On Error GoTo CatchAllError

    vAddInfo = Split(sAddInfo, gsDELIMITER1)
    For nInfoLoop = LBound(vAddInfo) To UBound(vAddInfo)

        If (vAddInfo(nInfoLoop) <> "") Then
        
            If Left(vAddInfo(nInfoLoop), 1) = "c" _
            And oUser.CheckPermission(gsFnAddIComment) Then
            
                vAddItem = Split(vAddInfo(nInfoLoop), gsDELIMITER2)
                
                'comment format:
                '[0] type (c)
                '[1] delete all flag (1 or 0)
                '[2] total length of comments including those already saved
                '    comment text, blocks of 3:
                '   [3] timestamp
                '   [4] userid
                '   [5] text
                '   ...
                'starting at 3rd item build a comment string by delimiting items with linefeeds
                For nItemLoop = 3 To UBound(vAddItem) - 1
                    sAddItem = sAddItem & vAddItem(nItemLoop) & vbCrLf
                Next
                    

                If oUser.CheckPermission(gsFnViewIComments) Then
                    'if the returned comment string differs from that saved, and it is <255 chars, save it
                    If (oResponse.Comments <> sAddItem) And (Len(sAddItem) < 255) Then
                        oResponse.Comments = sAddItem
                    End If
                Else
                    'if user cannot view comments, the returned comment will contain only new
                    'comments, previous comments will not have been sent to browser
                    'if 'delete all comments' flag is true(1), clear saved comments
                    If (vAddItem(1) = "1") Then
                        oResponse.Comments = ""
                    End If
                    
                    'if the comment string already saved + returned comment string <255, concatenate and save
                    If (Len(oResponse.Comments) + Len(sAddItem) + 1) < 255 Then
                        oResponse.Comments = sAddItem & oResponse.Comments
                    End If
                End If
            End If
        End If
    Next
    Exit Sub
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLMIMessage.SaveComment"
End Sub

'--------------------------------------------------------------------------------------------------
Public Sub RaiseMIMessage(ByRef oUser As MACROUser, ByRef vMIErrors As Variant, ByVal sAddInfo As String, _
    ByVal eScope As MIMsgScope, ByVal nTimezoneOffset As Integer, ByVal sStudyName As String, _
    ByVal lStudyId As Long, ByVal sSite As String, ByVal lSubjectId As Long, Optional ByVal lVisitId As Long = 0, _
    Optional ByVal nVisitCycle As Integer = 0, Optional ByVal lEformTaskId As Long = 0, _
    Optional ByVal lEFormId As Long = 0, Optional ByVal nEFormCycle As Integer = 1, _
    Optional ByVal lResponseId As Long = 0, Optional ByVal nResponseCycle As Integer = 1, _
    Optional ByVal dResponseTime As Double = 0, Optional ByVal sResponseValue As String = "", _
    Optional ByVal lResponseQuestionId As Long = 0, Optional ByVal sResponseUser As String = "", _
    Optional ByRef oSubject As StudySubject = Nothing, Optional ByVal bShowExistingSDV As Boolean = False, _
    Optional ByRef sSDVCall As String = "")
'--------------------------------------------------------------------------------------------------
'   ic 08/01/2002
'   raises discrepancies, notes, sdvs, adds comments.
'--------------------------------------------------------------------------------------------------
' revisions
'   ic 21/11/2002 changed arguements, pass objects rather than individual items
'   ic 10/12/2002 rewrite to handle eform, visit, subject mimessages in a more logical way
'   ic NEED TO INTRODUCE BETTER ERROR HANDLING SO THAT USER IS INFORMED WHEN RAISE FAILS - USE
'   VMIERRORS BYREF ARGUMENT
'   ic 01/03/2004 added error feedback
'   ic 10/05/2004 added oSubject parameter so MIMessage status can be updated in 'in-memory' subject
'   ic 29/06/2004 added error handling, commented out previous mimessage fix - this will need to
'                 be reinstated
'   ic 12/07/2004 reinstated fix
'   ic 09/11/2004 bug 2400, pass responsecycle to MIMessageExists() function
'   ic 01/04/2005 issue 2541 added show existing sdv code and arguments
'   ic 05/07/2005 issue 2464, added extra params to fnMIMessageUrl() js function call
'   NCJ 29 Dec 06 - Bug 2861 - Check response exists before creating MIMessage
'--------------------------------------------------------------------------------------------------

Dim oDiscrepancy As MIDiscrepancy
Dim oNote As MINote
Dim oSDV As MISDV
Dim oMIMData As MIDataLists
Dim nNoteStatus As eNoteMIMStatus
Dim nSDVStatus As eSDVMIMStatus
Dim nInfoLoop As Integer
Dim vAddInfo As Variant
Dim vAddItem As Variant
Dim lObjectId As Long
Dim nObjectSource As Integer
Dim sScope As String
' NCJ 29 Dec 06
Dim oMDL As MIDataLists
Dim vDetails As Variant

    On Error GoTo CatchAllError
    
    If sAddInfo = "" Then Exit Sub
    
    ' NCJ 29 Dec 06 - Bug 2861 - If raising for a response, check it exists
    If eScope = MIMsgScope.mimscQuestion Then
        Set oMDL = New MIDataLists
        vDetails = oMDL.GetResponseDetails(oUser.CurrentDBConString, sStudyName, sSite, lSubjectId, lResponseId, nResponseCycle)
        Set oMDL = Nothing
        ' Check response exists
        If IsNull(vDetails) Then
            ' Can't do it! (But we just exit quietly - no error report...??)
            Exit Sub
        End If
    End If
    
    'split the string value into parts (discrepancies,notes,sdvs,comment,rfc,password...)
    vAddInfo = Split(sAddInfo, gsDELIMITER1)

    For nInfoLoop = LBound(vAddInfo) To UBound(vAddInfo)

        If (vAddInfo(nInfoLoop) <> "") Then
            vAddItem = Split(vAddInfo(nInfoLoop), gsDELIMITER2)
            If (Len(vAddItem(1)) > 2000) Then Exit Sub
            
            
            Select Case Left(vAddInfo(nInfoLoop), 1)
            Case "c":
                'comment - handled elsewhere
            Case "d":
                'discrepancy
                If oUser.CheckPermission(gsFnCreateDiscrepancy) Then
                    If (Not IsNumeric(vAddItem(2))) Then vAddItem(2) = "0"
                    Set oDiscrepancy = New MIDiscrepancy
                    Call oDiscrepancy.Raise(oUser.CurrentDBConString, CStr(vAddItem(1)), CInt(vAddItem(3)), CLng(vAddItem(2)), _
                                            oUser.UserName, oUser.UserNameFull, mimsServer, eScope, sStudyName, _
                                            sSite, lSubjectId, lVisitId, nVisitCycle, lEformTaskId, lResponseId, nResponseCycle, _
                                            dResponseTime, sResponseValue, lEFormId, nEFormCycle, lResponseQuestionId, _
                                            sResponseUser, IMedNow, nTimezoneOffset)
                    oDiscrepancy.Save
                    Call UpdateMIMsgStatus(oUser.CurrentDBConString, mimtDiscrepancy, sStudyName, lStudyId, sSite, lSubjectId, _
                                           lVisitId, nVisitCycle, lEformTaskId, lResponseId, nResponseCycle, oSubject)
                End If

            Case "n":
                'note
                If (vAddItem(2) = "1") Then
                    nNoteStatus = eNoteMIMStatus.nsPublic
                Else
                    nNoteStatus = eNoteMIMStatus.nsPrivate
                End If
                Set oNote = New MINote
                Call oNote.Init(oUser.CurrentDBConString, CStr(vAddItem(1)), oUser.UserName, oUser.UserNameFull, mimsServer, _
                                eScope, sStudyName, sSite, lSubjectId, lVisitId, nVisitCycle, lEformTaskId, _
                                lResponseId, nResponseCycle, dResponseTime, sResponseValue, lEFormId, nEFormCycle, _
                                lResponseQuestionId, sResponseUser, IMedNow, nTimezoneOffset, nNoteStatus)
                oNote.Save
                Call UpdateNoteStatus(oUser.CurrentDBConString, eScope, sStudyName, lStudyId, sSite, lSubjectId, _
                                      lVisitId, nVisitCycle, lEformTaskId, lResponseId, nResponseCycle, oSubject)
    
            Case "s":
                'sdv
                Set oMIMData = New MIDataLists
                If Not oMIMData.MIMessageExists(oUser.CurrentDBConString, mimtSDVMark, eScope, sStudyName, sSite, lSubjectId, lObjectId, _
                                              nObjectSource, lVisitId, nVisitCycle, lEformTaskId, lResponseId, nResponseCycle) Then
                    If oUser.CheckPermission(gsFnCreateSDV) Then
                        If (vAddItem(2) = "1") Then
                            nSDVStatus = eSDVMIMStatus.ssPlanned
                        ElseIf (vAddItem(2) = "2") Then
                            nSDVStatus = eSDVMIMStatus.ssQueried
                        Else
                            nSDVStatus = eSDVMIMStatus.ssDone
                        End If
                        Set oSDV = New MISDV
                        Call oSDV.Raise(oUser.CurrentDBConString, CStr(vAddItem(1)), nSDVStatus, oUser.UserName, oUser.UserNameFull, _
                                        mimsServer, IMedNow, nTimezoneOffset, eScope, sStudyName, sSite, lSubjectId, lVisitId, _
                                        nVisitCycle, lEformTaskId, lResponseId, nResponseCycle, lEFormId, nEFormCycle, _
                                        lResponseQuestionId, sResponseUser, dResponseTime, sResponseValue)
                        oSDV.Save
                        Call UpdateMIMsgStatus(oUser.CurrentDBConString, mimtSDVMark, sStudyName, lStudyId, sSite, lSubjectId, _
                                               lVisitId, nVisitCycle, lEformTaskId, lResponseId, nResponseCycle, oSubject)
                    End If
                Else
                    If (bShowExistingSDV) Then
                        'when adding an sdv to an item and finding a pre-existing one, display the sdv
                        Select Case eScope
                        Case MIMsgScope.mimscSubject: sScope = "1000"
                        Case MIMsgScope.mimscVisit: sScope = "0100"
                        Case MIMsgScope.mimscEForm: sScope = "0010"
                        Case MIMsgScope.mimscQuestion: sScope = "0001"
                        End Select
                        'create a javascript call that will load the sdv browser
                        sSDVCall = "if(confirm('An SDV already exists on this item, click OK to view the SDV.'))" _
                        & "window.parent.fnMIMessageUrl('1','" & lStudyId & "','" & sSite & "','','" & lEFormId & "'," _
                        & "'" & lResponseQuestionId & "','" & lSubjectId & "','','false','','1111','" & sScope & "','0'," _
                        & "undefined,'" & nVisitCycle & "','" & nEFormCycle & "','" & nResponseCycle & "');"
                    Else
                        'ic 01/03/2004 added error feedback
                        vMIErrors = AddToArray(vMIErrors, "Permission Denied", "An SDV already exists on this item.")
                    End If
                End If
                
            Case Else:
            End Select
                
        End If
    Next
    

    On Error Resume Next
    Set oDiscrepancy = Nothing
    Set oNote = Nothing
    Set oSDV = Nothing
    Set oMIMData = Nothing
    Exit Sub
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLMIMessage.RaiseMIMessage"
End Sub

'--------------------------------------------------------------------------------------------------
Public Function GetMIMessageAudit(ByRef oUser As MACROUser, _
                                  ByVal nType As MIMsgType, _
                                  ByVal lStudy As Long, _
                                  ByVal sSite As String, _
                                  ByVal lId As Long, _
                                  ByVal nSrc As Integer, _
                                  Optional ByVal bIncludeStudySiteSubjInfo As Boolean = True) As String
'--------------------------------------------------------------------------------------------------
'   ic 09/05/2003
'   function returns an html mimessage audit table
' MLM 02/07/03: Added optional boolean argument to control whether the left-hand table of MIMessage
'   properties (those that are the same for all rows in the audit trail) is displayed.
'   This is used from frmViewDiscrepancies to better display the audit trail for long trails/low resolutions.
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim vJSComm() As String
Dim oMIMsg As Variant
Dim oMsg As MIMsg
Dim oStudy As Study
Dim oMIMS As MIMsgStatic

    On Error GoTo CatchAllError
    ReDim vJSComm(0)

    Set oStudy = New Study
    Set oStudy = oUser.Studies.StudyById(lStudy)
    Set oMIMS = New MIMsgStatic

    Select Case nType
    Case MIMsgType.mimtDiscrepancy:
        Set oMIMsg = New MIDiscrepancy
        Call oMIMsg.Load(oUser.CurrentDBConString, lId, nSrc, sSite)
        
    Case MIMsgType.mimtSDVMark:
        Set oMIMsg = New MISDV
        Call oMIMsg.Load(oUser.CurrentDBConString, lId, nSrc, sSite)
        
    Case Else:
        Set oMIMsg = New MINote
        Call oMIMsg.Load(oUser.CurrentDBConString, lId, nSrc, sSite)
        
    End Select
       
    'start html
    If bIncludeStudySiteSubjInfo Then
        Call AddStringToVarArr(vJSComm, "<table width='100%' border='0'>" _
            & "<tr><td valign='top' width='250'>")
    
        Call AddStringToVarArr(vJSComm, "<table bgcolor='d3d3d3' width='100%' border='1' cellpadding='0' cellspacing='0' class='clsTableText' bordercolor='d3d3d3'>" _
            & "<tr height='20'><td bgcolor='fffaf0' width='30%'>study</td><td colspan='3'>" & oStudy.StudyName & "</td></tr>" _
            & "<tr height='20'><td bgcolor='fffaf0'>site</td><td colspan='3'>" & sSite & "</td></tr>" _
            & "<tr height='20'><td bgcolor='fffaf0'>subject</td><td>" & oUser.GetSubjectLabel(oStudy.StudyId, sSite, oMIMsg.SubjectId) & "</td>" _
            & "<td bgcolor='fffaf0'>Id</td><td>" & oMIMsg.SubjectId & "</td></tr>" _
            & "<tr height='20'><td bgcolor='fffaf0'>visit</td><td colspan='3'>" & oUser.DataLists.GetStudyItemName(soVisit, oStudy.StudyId, oMIMsg.VisitId) _
            & IIf((oMIMsg.VisitCycle > 1), " [" & oMIMsg.VisitCycle & "]", "") & "</td></tr>" _
            & "<tr height='20'><td bgcolor='fffaf0'>eForm</td><td colspan=3>" & oUser.DataLists.GetStudyItemName(soeform, oStudy.StudyId, oMIMsg.EFormId) _
            & IIf((oMIMsg.EFormCycle > 1), " [" & oMIMsg.EFormCycle & "]", "") & "</td></tr>" _
            & "<tr height='20'><td bgcolor='fffaf0'>Question</td><td colspan=3 >" & oUser.DataLists.GetStudyItemName(soQuestion, oStudy.StudyId, oMIMsg.QuestionId) _
            & IIf((oMIMsg.ResponseCycle > 1), " [" & oMIMsg.ResponseCycle & "]", "") & "</td></tr>" _
            & "<tr height='20'><td bgcolor='fffaf0'>Status</td><td>" & oMIMS.GetStatusText(nType, oMIMsg.CurrentStatus) & "</td>")
            
        If (nType = mimtDiscrepancy) Then
            Call AddStringToVarArr(vJSComm, "<td bgcolor='fffaf0'>Prty</td><td bgcolor='d3d3d3'>" & oMIMsg.CurrentMessage.Priority & "</td></tr></table>")
        Else
            Call AddStringToVarArr(vJSComm, "<td colspan='3'></td></tr></table>")
        End If
        
        Call AddStringToVarArr(vJSComm, "</td><td valign='top'>")
    End If
    
    Select Case nType
    Case mimtNote:
        Call AddStringToVarArr(vJSComm, "<table width='100%' border='0' align='left' cellpadding='0' cellspacing='1'>" _
            & "<tr class='clsTableText'><td>" & oMIMsg.CurrentMessage.Text & "</td></tr>")
    Case Else:
        Call AddStringToVarArr(vJSComm, "<table width='100%' border='0' align='left' cellpadding='0' cellspacing='1'>" _
            & "<tr height='15' class='clsTableHeaderText'>" _
            & "<td width='20%'>&nbsp;Created</td>" _
            & "<td width='10%'>&nbsp;Status</td>" _
            & "<td width='50%'>&nbsp;Text</td>" _
            & "<td width='15%'>&nbsp;Value</td>" _
            & "<td width='5%'>&nbsp;User</td>" _
            & "</tr>")

            For Each oMsg In oMIMsg.Messages
                Call AddStringToVarArr(vJSComm, "<tr class='clsTableText'>" _
                             & "<td>" & GetLocalFormatDate(oUser, CDate(oMsg.TimeCreated), eDateTimeType.dttDMYT) _
                             & " " & RtnDifferenceFromGMT(oMsg.TimeCreatedTimezoneOffset) & "</td>" _
                             & "<td>" & oMsg.StatusText & "</td>" _
                             & "<td>" & ReplaceWithHTMLCodes(oMsg.Text) & "</td>" _
                             & "<td>" & ReplaceWithHTMLCodes(oMsg.ResponseValue) & "</td>" _
                             & "<td>" & oMsg.UserName & "</td>" _
                             & "</tr>")
            Next

    End Select
    
    Call AddStringToVarArr(vJSComm, "</table>")
    If bIncludeStudySiteSubjInfo Then
        Call AddStringToVarArr(vJSComm, "</td></tr></table>")
    End If
    
    GetMIMessageAudit = Join(vJSComm, "")
    
    Set oMsg = Nothing
    Set oMIMsg = Nothing
    Set oStudy = Nothing
    Set oMIMS = Nothing
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLMIMessage.GetMIMessageAudit"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnMIObjectArray(ByVal sStatus As String) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 11/11/2002
'   function returns an array representing the statuses requested in a passed binary string
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim vRtn As Variant
Dim ndx As Integer
    
    On Error GoTo CatchAllError
    sStatus = Trim(sStatus)
    
    
    If (Len(sStatus) <> 4) Or (Not IsNumeric(sStatus)) Or (sStatus = "0000") Then
         vRtn = ""
    Else
        ReDim vRtn(3)
        
        If Mid(sStatus, 1, 1) <> "0" Then
            vRtn(ndx) = MIMsgScope.mimscSubject
            ndx = ndx + 1
        End If
        If Mid(sStatus, 2, 1) <> "0" Then
            vRtn(ndx) = MIMsgScope.mimscVisit
            ndx = ndx + 1
        End If
        If Mid(sStatus, 3, 1) <> "0" Then
            vRtn(ndx) = MIMsgScope.mimscEForm
            ndx = ndx + 1
        End If
        If Mid(sStatus, 4, 1) <> "0" Then
            vRtn(ndx) = MIMsgScope.mimscQuestion
            ndx = ndx + 1
        End If
        
        ReDim Preserve vRtn(ndx - 1)
    End If
        
    RtnMIObjectArray = vRtn
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLMIMessage.RtnMIObjectArray"
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetMIMessageList(ByRef oUser As MACROUser, ByVal sType As String, ByVal sStudyCode As String, _
    ByVal sSiteCode As String, ByVal sVisitCode As String, ByVal sVisitCycle As String, ByVal sCRFPageId As String, _
    ByVal sCRFPageCycle As String, ByVal sQuestion As String, ByVal sQuestionCycle As String, ByVal sSrchUserName As String, _
    ByVal sSubjectId As String, ByVal sSubjectLabel As String, ByVal sStatus As String, ByVal sTime As String, _
    ByVal sBefore As String, ByVal sScope As String, ByVal bNewWindow As Boolean, ByVal sUpdate As String, _
    Optional ByVal enInterface As eInterface = iwww, Optional ByVal sBookmark As String = "0", _
    Optional vErrors As Variant) As String
'--------------------------------------------------------------------------------------------------
'   ic 27/01/2003
'--------------------------------------------------------------------------------------------------
' REVISIONS
'   dph 12/02/2003  Added bNewWindow functionality to create message list
'   dph 18/02/2003  Added GetLocalFormatDate for formatting displayed dates
'   ic 30/05/2003 moved error display code out of conditional statement, now will always display errors
'   ic 05/06/2003 added extra 'refresh z-order' parameter
'   ic 06/05/2003 added ReplaceWithJSChars() call around sSubjectLabel
'   ic 02/09/2003 comment out bChangeData - not used
'   dph 02/09/2003 show full username for mimessages
'   ic 05/03/2004 added subjectid to fnM() calls for subject locking on updates
'   ic 29/06/2004 added error handling
'   ic 16/07/2004 fixed locking during sdv update
'   ic 01/04/2005 added [cycle] to eform and visit
'   ic 01/04/2005 issue 2508, top aligned all rows
'   ic 27/04/2005 issue 2222, added fnInitUser() js function to pass user permissions
'   ic 28/04/2005 issue 2431, added a print button
'   ic 11/05/2005 issue 2571, added row header count parameter to javascript fnOnClick() calls
'   ic 04/07/2005 issue 2464, added visit, eform, question cycle
'   ic 18/08/2006 issue 2782, removed erroneous 'Repeat' argument from function
'   ic 27/02/2007 issue 2114, added GMT to timestamps, added cycle numbers to repeating eforms and visits
'--------------------------------------------------------------------------------------------------
Dim bCreateDisc As Boolean
Dim bCreateSDV As Boolean
Dim bChangeData As Boolean
Dim vData As Variant
Dim nType As MIMsgType
Dim vJSComm() As String
Dim sURL As String
Dim lPageLength As Long
Dim lBookmark As Long
Dim lStart As Long
Dim lStop As Long
Dim lLoop As Long
'Dim sQuestionLabel As String
Dim sQuestionName As String
Dim sStatusName As String
Dim sEformName As String
Dim sVisitName As String
Dim oMIMS As MIMsgStatic
Dim lRaised As Long
Dim lResponded As Long
Dim lPlanned As Long

    On Error GoTo CatchAllError
    ReDim vJSComm(0)
    
    Select Case sType
    Case "0": nType = MIMsgType.mimtDiscrepancy
    Case "2": nType = MIMsgType.mimtNote
    Case Else: nType = MIMsgType.mimtSDVMark
    End Select
    
    Call RtnMIMsgStatusCount(oUser, lRaised, lResponded, lPlanned)
    bCreateDisc = oUser.CheckPermission(gsFnCreateDiscrepancy)
    bCreateSDV = oUser.CheckPermission(gsFnCreateSDV)
    bChangeData = oUser.CheckPermission(gsFnChangeData)

    'MLM 01/07/05: Explicitly request all repeats of visit, eform and question (although this may not be what's wanted)..
    vData = RtnMIMessageList(oUser, sType, sStudyCode, sSiteCode, sVisitCode, sVisitCycle, sCRFPageId, sCRFPageCycle, _
    sQuestion, sQuestionCycle, sSrchUserName, sSubjectId, sSubjectLabel, sStatus, sTime, sBefore, sScope, vErrors)
        
    sURL = "MIMessageTop.asp?fltSt=" & sStudyCode & "&fltSi=" & sSiteCode & "&fltVi=" & sVisitCode _
                            & "&fltEf=" & sCRFPageId & "&fltQu=" & sQuestion & "&fltUs=" & sSrchUserName _
                            & "&fltSj=" & sSubjectId & "&fltSjLb=" & URLEncodeString(sSubjectLabel) & "&fltSs=" & sStatus _
                            & "&fltObj=" & sScope & "&fltTm=" & URLEncodeString(sTime) & "&fltB4=" & sBefore & "&type=" & sType _
                            & "&newwin=" & CStr(RtnJSBoolean(bNewWindow)) & "&fltVRpt=" & sVisitCycle _
                            & "&fltERpt=" & sCRFPageCycle & "&fltQRpt=" & sQuestionCycle
 
    If (enInterface = iwww) Then
        lPageLength = CLng(oUser.UserSettings.GetSetting(SETTING_PAGE_LENGTH, 50))
    Else
        lPageLength = UBound(vData, 2)
    End If
    
    'ic start html body
    Call AddStringToVarArr(vJSComm, "<body onload='fnPageLoaded();'>" & vbCrLf)

    Call AddStringToVarArr(vJSComm, "<script language='javascript'>" & vbCrLf _
                & "function fnPageLoaded(){" & vbCrLf _
                & "window.parent.sWinState=" & Chr(34) & "3|" & sType & "|" & sStudyCode & "|" & sSiteCode & "|" _
                    & sVisitCode & "|" & sCRFPageId & "|" & sQuestion & "|" _
                    & URLEncodeString(sSubjectLabel) & "|" & sSrchUserName & "|" & sBefore & "|" & URLEncodeString(sTime) & "|" _
                    & sStatus & "|" & sScope & "|" & sBookmark & Chr(34) & ";" & vbCrLf _
                & "fnHideLoader();" & vbCrLf)
    
    'ic initialise user permissions
    Call AddStringToVarArr(vJSComm, "fnInitUser(" & RtnJSBoolean(bChangeData) & "," & RtnJSBoolean(bCreateDisc) _
        & "," & RtnJSBoolean(bCreateSDV) & ");" & vbCrLf)
    
    If (Not bNewWindow) Then
        'ic 05/06/2003 added extra 'refresh z-order' parameter
        Call AddStringToVarArr(vJSComm, "window.parent.window.parent.fnSTLC('" & gsVIEW_RAISED_DISCREPANCIES_MENUID & "','" & CStr(lRaised) & "',0);" & vbCrLf _
                                  & "window.parent.window.parent.fnSTLC('" & gsVIEW_RESPONDED_DISCREPANCIES_MENUID & "','" & CStr(lResponded) & "',0);" & vbCrLf _
                                  & "window.parent.window.parent.fnSTLC('" & gsVIEW_PLANNED_SDV_MARKS_MENUID & "','" & CStr(lPlanned) & "',1);" & vbCrLf)
        If (sUpdate <> "") Then
            Call AddStringToVarArr(vJSComm, "window.parent.window.parent.fnUpdateStatusOnEform(" & sUpdate & ");" & vbCrLf)
        End If
    Else
        If (sUpdate <> "") Then
            Call AddStringToVarArr(vJSComm, "window.parent.fnUpdateStatusOnEform(" & sUpdate & ");" & vbCrLf)
        End If
    End If
    
    'errors encountered during save
    If Not IsMissing(vErrors) Then
        If Not IsEmpty(vErrors) Then
            Call AddStringToVarArr(vJSComm, "alert('MACRO encountered problems while updating. Some updates could not be completed." _
                & "\nIncomplete updates are listed below\n\n")

            For lLoop = LBound(vErrors, 2) To UBound(vErrors, 2)
                Call AddStringToVarArr(vJSComm, vErrors(0, lLoop) & " - " & vErrors(1, lLoop) & "\n")
            Next

            Call AddStringToVarArr(vJSComm, "');" & vbCrLf)
        End If
    End If
    
    Call AddStringToVarArr(vJSComm, "window.parent.window.frames[1].location.replace('blank.htm');" & vbCrLf _
            & "}</script>")


    If Not IsNull(vData) Then
        Set oMIMS = New MIMsgStatic
    
        Call AddStringToVarArr(vJSComm, "<form name='FormMI' action=" & Chr(34) & sURL & "&bookmark=" & CStr(lBookmark) & Chr(34) & " method='post'>" & vbCrLf _
                    & "<input type='hidden' name='bidentifier'>" _
                    & "<input type='hidden' name='btype'>" _
                    & "<input type='hidden' name='badd'>" _
                    & "</form>" & vbCrLf)
        
        If (sBookmark = "") Or (Not IsNumeric(sBookmark)) Then sBookmark = "0"
        'calculate the start row and end row based on start row (bookmark) and page length
        If ((CLng(sBookmark) >= UBound(vData, 2)) Or (CLng(sBookmark) < 0)) Then
            lStart = 0
        Else
            lStart = CLng(sBookmark)
        End If
        If ((lStart + lPageLength) >= UBound(vData, 2)) Then
            lStop = UBound(vData, 2)
        Else
            lStop = (lStart + lPageLength) - 1
        End If
        
        Call AddStringToVarArr(vJSComm, "<table align='center' width='100%' cellpadding='0' cellspacing='0'>" & vbCrLf _
                & "<tr class='clsLabelText' height='50'>" & vbCrLf _
                & "<td colspan='7' align='right'><a href='javascript:window.print();'>" _
                & "<img src='../img/ico_print.gif' border='0' alt='Print listing'></a>&nbsp;&nbsp;Record(s) " _
                & lStart + 1 & " to " & lStop + 1 & " of " & UBound(vData, 2) + 1 & "&nbsp;&nbsp;")

        If (enInterface = iwww) Then
            'write previous page icon
            If (lStart > 0) Then
                Call AddStringToVarArr(vJSComm, "<a href='" & sURL & "&bookmark=" & lStart - lPageLength & "'>" _
                    & "<img src='../img/ico_backon.gif' border='0' alt='previous page'></a>&nbsp;" & vbCrLf)
            Else
                Call AddStringToVarArr(vJSComm, "<img src='../img/ico_back.gif'>&nbsp;" & vbCrLf)
            End If

            'write next page icon
            If (lStop < UBound(vData, 2)) Then
                Call AddStringToVarArr(vJSComm, "<a href='" & sURL & "&bookmark=" & lStop + 1 & "'>" _
                    & "<img src='../img/ico_forwardon.gif' border='0' alt='next page'></a>&nbsp;" & vbCrLf)
            Else
                Call AddStringToVarArr(vJSComm, "<img src='../img/ico_forward.gif'>&nbsp;" & vbCrLf)
            End If
        End If

        Call AddStringToVarArr(vJSComm, "</td></tr></table><br>")

        Select Case nType
        Case MIMsgType.mimtDiscrepancy:
            Call AddStringToVarArr(vJSComm, "<table id='tmimessage' onmousedown='fnOnClick(this,1);' onmouseover='fnOnMouseOver(this,1);' onmouseout='fnOnMouseOut(this);' width='100%' cellpadding='0' cellspacing='0'>" _
                        & "<tr height='20' style='cursor:default;' class='clsTableHeaderText'>" & vbCrLf _
                        & "<td width='30'>&nbsp;Prty</td><td width='150'>&nbsp;Date</td>" _
                        & "<td width='75'>&nbsp;Status</td><td width='100'>&nbsp;Subject</td>" _
                        & "<td width='100'>&nbsp;eForm</td><td width='100'>&nbsp;Question</td>" _
                        & "<td width='100'>&nbsp;User Name</td><td width='50'>&nbsp;")
                        
             'TA 18/11/2004: issue 2448 set column header according to Use OC ID flag
             If RtnUseOCIdFlag Then
                Call AddStringToVarArr(vJSComm, "OC Id")
             Else
                Call AddStringToVarArr(vJSComm, "Id")
             End If
                        
             Call AddStringToVarArr(vJSComm, "</td>" _
                        & "<td width='100'>&nbsp;Text</td>" & vbCrLf _
                        & "</tr>")
                        
            For lLoop = lStart To lStop
                sVisitName = oUser.DataLists.GetStudyItemName(soVisit, CLng(vData(mmcStudyid, lLoop)), CLng(vData(mmcVisitId, lLoop)))
                sEformName = oUser.DataLists.GetStudyItemName(soeform, CLng(vData(mmcStudyid, lLoop)), CLng(vData(mmcEFormId, lLoop)))
                sQuestionName = oUser.DataLists.GetStudyItemName(soQuestion, CLng(vData(mmcStudyid, lLoop)), CLng(vData(mmcQuestionId, lLoop)))
                sStatusName = oMIMS.GetStatusText(nType, vData(mmcStatus, lLoop))
                
                'striping
                If ((lLoop Mod 2) = 0) Then
                    Call AddStringToVarArr(vJSComm, "<tr valign='top' style='cursor:hand;' class='clsTableText'>")
                Else
                    Call AddStringToVarArr(vJSComm, "<tr valign='top' style='cursor:hand;' class='clsTableTextS'>")
                End If
                Call AddStringToVarArr(vJSComm, "<a onMouseup=" & Chr(34) & "fnM(0,'" & vData(mmcStudyName, lLoop) & gsDELIMITER1 _
                    & vData(mmcStudyid, lLoop) & gsDELIMITER1 & vData(mmcSite, lLoop) & gsDELIMITER1 _
                    & vData(mmcSubjectId, lLoop) & gsDELIMITER1 & ReplaceWithJSChars(sVisitName) & gsDELIMITER1 _
                    & ReplaceWithJSChars(sEformName) & gsDELIMITER1 & vData(mmcEFormId, lLoop) & gsDELIMITER1 _
                    & vData(mmcEFormTaskId, lLoop) & gsDELIMITER1 & ReplaceWithJSChars(sQuestionName) & gsDELIMITER1 _
                    & sStatusName & gsDELIMITER1 & vData(mmcObjectId, lLoop) & gsDELIMITER1 _
                    & vData(mmcObjectSource, lLoop) & gsDELIMITER1 & vData(mmcPrioirty, lLoop) & gsDELIMITER1 _
                    & vData(mmcResponseTaskId, lLoop) & gsDELIMITER1 & vData(mmcResponseCycle, lLoop) & gsDELIMITER1 _
                    & ReplaceWithJSChars(ReplaceWithHTMLCodes(ConvertFromNull(vData(mmcText, lLoop), vbString))) & gsDELIMITER1 & vData(mmcSubjectId, lLoop) & "',event.button," _
                    & RtnJSBoolean(Not bNewWindow) & "," & RtnJSBoolean((oUser.UserName = vData(mmcUserName, lLoop)) And (vData(mmcTimeSent, lLoop) = 0)) & ");" & Chr(34) & ">")
                
                'ic 27/02/2007 issue 2114, added GMT to timestamps, added cycle numbers to repeating eforms and visits
                ' dph 02/09/2003 show full username for mimessages
                Call AddStringToVarArr(vJSComm, "<td align='center'>" & vData(mmcPrioirty, lLoop) & "</td>" _
                     & "<td>" & GetLocalFormatDate(oUser, CDate(vData(mmcCreated, lLoop)), eDateTimeType.dttDMYT) _
                     & " " & RtnDifferenceFromGMT(vData(mmcCreated_TZ, lLoop)) & "</td>" _
                     & "<td>" & sStatusName & "</td>" _
                     & "<td>" & RtnSubjectText(vData(mmcSubjectId, lLoop), vData(mmcSubjectLabel, lLoop)) & "</td>" _
                     & "<td>" & sEformName & IIf((CInt(vData(mmcEFormCycle, lLoop)) > 1), " [" & vData(mmcEFormCycle, lLoop) & "]", "") & "</td>" _
                     & "<td>" & sQuestionName & IIf((CInt(vData(mmcResponseCycle, lLoop)) > 1), " [" & vData(mmcResponseCycle, lLoop) & "]", "") & "</td>" _
                     & "<td>" & vData(mmcUserNameFull, lLoop) & "</td>" _
                     & "<td>" & vData(mmcExternalId, lLoop) & "</td>" _
                     & "<td>" & ReplaceWithHTMLCodes(ConvertFromNull(vData(mmcText, lLoop), vbString)) & "</td>" _
                     & "</a></tr>")
            Next
                            
        Case MIMsgType.mimtNote:
            Call AddStringToVarArr(vJSComm, "<table id='tmimessage' onmousedown='fnOnClick(this,1);' onmouseover='fnOnMouseOver(this,1);' onmouseout='fnOnMouseOut(this);' width='100%' cellpadding='0' cellspacing='0'>" _
                        & "<tr height='20' style='cursor:default;' class='clsTableHeaderText'>" & vbCrLf _
                        & "<td width='75'>&nbsp;Timestamp</td><td width='100'>&nbsp;Subject</td>" _
                        & "<td width='100'>&nbsp;eForm</td><td width='100'>&nbsp;Question</td>" _
                        & "<td width='100'>&nbsp;User Name</td><td width='100'>&nbsp;Status</td>" _
                        & "<td width='100'>&nbsp;Text</td>" & vbCrLf _
                        & "</tr>")
                        
            For lLoop = lStart To lStop
                sVisitName = oUser.DataLists.GetStudyItemName(soVisit, CLng(vData(mmcStudyid, lLoop)), CLng(vData(mmcVisitId, lLoop)))
                sEformName = oUser.DataLists.GetStudyItemName(soeform, CLng(vData(mmcStudyid, lLoop)), CLng(vData(mmcEFormId, lLoop)))
                sQuestionName = oUser.DataLists.GetStudyItemName(soQuestion, CLng(vData(mmcStudyid, lLoop)), CLng(vData(mmcQuestionId, lLoop)))
                sStatusName = oMIMS.GetStatusText(nType, vData(mmcStatus, lLoop))
                
                'striping
                If ((lLoop Mod 2) = 0) Then
                    Call AddStringToVarArr(vJSComm, "<tr valign='top' style='cursor:hand;' class='clsTableText'>")
                Else
                    Call AddStringToVarArr(vJSComm, "<tr valign='top' style='cursor:hand;' class='clsTableTextS'>")
                End If
                Call AddStringToVarArr(vJSComm, "<a onMouseup=" & Chr(34) & "fnM(2,'" & vData(mmcStudyName, lLoop) & gsDELIMITER1 _
                    & vData(mmcStudyid, lLoop) & gsDELIMITER1 & vData(mmcSite, lLoop) & gsDELIMITER1 _
                    & vData(mmcSubjectId, lLoop) & gsDELIMITER1 & ReplaceWithJSChars(sVisitName) & gsDELIMITER1 _
                    & ReplaceWithJSChars(sEformName) & gsDELIMITER1 & vData(mmcEFormId, lLoop) & gsDELIMITER1 _
                    & vData(mmcEFormTaskId, lLoop) & gsDELIMITER1 & ReplaceWithJSChars(sQuestionName) & gsDELIMITER1 _
                    & sStatusName & gsDELIMITER1 & vData(mmcId, lLoop) & gsDELIMITER1 _
                    & vData(mmcObjectSource, lLoop) & gsDELIMITER1 & vData(mmcPrioirty, lLoop) & gsDELIMITER1 _
                    & vData(mmcResponseTaskId, lLoop) & gsDELIMITER1 & vData(mmcResponseCycle, lLoop) & gsDELIMITER1 _
                    & ReplaceWithJSChars(ReplaceWithHTMLCodes(ConvertFromNull(vData(mmcText, lLoop), vbString))) & gsDELIMITER1 & vData(mmcSubjectId, lLoop) & "',event.button," _
                    & RtnJSBoolean(Not bNewWindow) & "," & RtnJSBoolean((oUser.UserName = vData(mmcUserName, lLoop)) And (vData(mmcTimeSent, lLoop) = 0)) & ");" & Chr(34) & ">")
                
                'ic 27/02/2007 issue 2114, added GMT to timestamps, added cycle numbers to repeating eforms and visits
                ' dph 02/09/2003 show full username for mimessages
                Call AddStringToVarArr(vJSComm, "<td>" & GetLocalFormatDate(oUser, CDate(vData(mmcCreated, lLoop)), eDateTimeType.dttDMYT) _
                     & " " & RtnDifferenceFromGMT(vData(mmcCreated_TZ, lLoop)) & "</td>" _
                     & "<td>" & RtnSubjectText(vData(mmcSubjectId, lLoop), vData(mmcSubjectLabel, lLoop)) & "</td>" _
                     & "<td>" & sEformName & IIf((CInt(vData(mmcEFormCycle, lLoop)) > 1), " [" & vData(mmcEFormCycle, lLoop) & "]", "") & "</td>" _
                     & "<td>" & sQuestionName & IIf((CInt(vData(mmcResponseCycle, lLoop)) > 1), " [" & vData(mmcResponseCycle, lLoop) & "]", "") & "</td>" _
                     & "<td>" & vData(mmcUserNameFull, lLoop) & "</td>" _
                     & "<td>" & sStatusName & "</td>" _
                     & "<td>" & ReplaceWithHTMLCodes(ConvertFromNull(vData(mmcText, lLoop), vbString)) & "</td>" _
                     & "</a></tr>")
            Next
            
        Case Else:
            Call AddStringToVarArr(vJSComm, "<table id='tmimessage' onmousedown='fnOnClick(this,1);' onmouseover='fnOnMouseOver(this,1);' onmouseout='fnOnMouseOut(this);' width='100%' cellpadding='0' cellspacing='0'>" _
                        & "<tr height='20' style='cursor:default;' class='clsTableHeaderText'>" & vbCrLf _
                        & "<td width='150'>&nbsp;Date</td><td width='30'>&nbsp;Scope</td>" _
                        & "<td width='75'>&nbsp;Status</td><td width='100'>&nbsp;Subject</td>" _
                        & "<td width='100'>&nbsp;Visit</td><td width='100'>&nbsp;eForm</td>" _
                        & "<td width='100'>&nbsp;Question</td><td width='100'>&nbsp;User Name</td>" _
                        & "<td width='100'>&nbsp;Text</td>" & vbCrLf _
                        & "</tr>")
                        
            For lLoop = lStart To lStop
                sVisitName = oUser.DataLists.GetStudyItemName(soVisit, CLng(vData(mmcStudyid, lLoop)), CLng(vData(mmcVisitId, lLoop)))
                sEformName = oUser.DataLists.GetStudyItemName(soeform, CLng(vData(mmcStudyid, lLoop)), CLng(vData(mmcEFormId, lLoop)))
                sQuestionName = oUser.DataLists.GetStudyItemName(soQuestion, CLng(vData(mmcStudyid, lLoop)), CLng(vData(mmcQuestionId, lLoop)))
                sStatusName = oMIMS.GetStatusText(nType, vData(mmcStatus, lLoop))

                'striping
                If ((lLoop Mod 2) = 0) Then
                    Call AddStringToVarArr(vJSComm, "<tr valign='top' style='cursor:hand;' class='clsTableText'>")
                Else
                    Call AddStringToVarArr(vJSComm, "<tr valign='top' style='cursor:hand;' class='clsTableTextS'>")
                End If
                Call AddStringToVarArr(vJSComm, "<a onMouseup=" & Chr(34) & "fnM(1,'" & vData(mmcStudyName, lLoop) & gsDELIMITER1 _
                    & vData(mmcStudyid, lLoop) & gsDELIMITER1 & vData(mmcSite, lLoop) & gsDELIMITER1 _
                    & vData(mmcSubjectId, lLoop) & gsDELIMITER1 & ReplaceWithJSChars(sVisitName) & gsDELIMITER1 _
                    & ReplaceWithJSChars(sEformName) & gsDELIMITER1 & vData(mmcEFormId, lLoop) & gsDELIMITER1 _
                    & vData(mmcEFormTaskId, lLoop) & gsDELIMITER1 & ReplaceWithJSChars(sQuestionName) & gsDELIMITER1 _
                    & sStatusName & gsDELIMITER1 & vData(mmcObjectId, lLoop) & gsDELIMITER1 _
                    & vData(mmcObjectSource, lLoop) & gsDELIMITER1 & vData(mmcPrioirty, lLoop) & gsDELIMITER1 _
                    & vData(mmcResponseTaskId, lLoop) & gsDELIMITER1 & vData(mmcResponseCycle, lLoop) & gsDELIMITER1 _
                    & ReplaceWithJSChars(ReplaceWithHTMLCodes(ConvertFromNull(vData(mmcText, lLoop), vbString))) & gsDELIMITER1 & vData(mmcSubjectId, lLoop) & "',event.button," _
                    & RtnJSBoolean((vData(mmcScope, lLoop) = 4) And (Not bNewWindow)) & "," & RtnJSBoolean((oUser.UserName = vData(mmcUserName, lLoop)) And (vData(mmcTimeSent, lLoop) = 0)) & "," & RtnJSBoolean((bNewWindow)) & ");" & Chr(34) & ">")
                    
                'ic 27/02/2007 issue 2114, added GMT to timestamps, added cycle numbers to repeating eforms and visits
                '   dph 02/09/2003 show full username for mimessages
                Call AddStringToVarArr(vJSComm, "<td>" & GetLocalFormatDate(oUser, CDate(vData(mmcCreated, lLoop)), eDateTimeType.dttDMYT) _
                     & " " & RtnDifferenceFromGMT(vData(mmcCreated_TZ, lLoop)) & "</td>" _
                     & "<td>" & GetScopeText(vData(mmcScope, lLoop)) & "</td>" _
                     & "<td>" & sStatusName & "</td>" _
                     & "<td>" & RtnSubjectText(vData(mmcSubjectId, lLoop), vData(mmcSubjectLabel, lLoop)) & "</td>" _
                     & "<td>" & sVisitName & IIf((CInt(vData(mmcVisitCycle, lLoop)) > 1), " [" & vData(mmcVisitCycle, lLoop) & "]", "") & "</td>" _
                     & "<td>" & sEformName & IIf((CInt(vData(mmcEFormCycle, lLoop)) > 1), " [" & vData(mmcEFormCycle, lLoop) & "]", "") & "</td>" _
                     & "<td>" & sQuestionName & IIf((CInt(vData(mmcResponseCycle, lLoop)) > 1), " [" & vData(mmcResponseCycle, lLoop) & "]", "") & "</td>" _
                     & "<td>" & vData(mmcUserNameFull, lLoop) & "</td>" _
                     & "<td>" & ReplaceWithHTMLCodes(ConvertFromNull(vData(mmcText, lLoop), vbString)) & "</td>" _
                     & "</a></tr>")
            Next
        End Select
        
        
    Else
        
        Call AddStringToVarArr(vJSComm, "<div class='clsMessageText'>Your query returned no records</div>")
    End If
    
    Call AddStringToVarArr(vJSComm, "</body>")
    GetMIMessageList = Join(vJSComm, "")
    Set oMIMS = Nothing
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLMIMessage.GetMIMessageList"
End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnMIMessageList(ByRef oUser As MACROUser, ByVal sType As String, ByVal sStudyCode As String, _
                                 ByVal sSiteCode As String, _
                                 ByVal sVisitId As String, ByVal sVisitCycle As String, _
                                 ByVal sCRFPageId As String, ByVal sCRFPageCycle As String, _
                                 ByVal sQuestion As String, ByVal sQuestionCycle As String, _
                                 ByVal sSrchUserName As String, ByVal sSubjectId As String, _
                                 ByVal sSubjectLabel As String, ByVal sStatus As String, ByVal sTime As String, _
                                 ByVal sBefore As String, ByVal sScope As String, Optional vErrors As Variant) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 12/12/2001
'   function returns mimessage list to populate grid with discrepancies/notes/sdvs
'
'   revisions
'   ic 06/11/2002 added sScope arguement, 4 char long string, which items to search for mimessages
'                 on (subject,visit,eform,question)
' DPH 08/11/2002 Changed to use Serialised User object
' ic 12/05/2003 added optional vErrors parameter
' DPH 13/10/2003 improve SQL performance
'   ic 29/06/2004 added error handling
' MLM 30/06/05: added visit, eForm and question cycle numbers.
'--------------------------------------------------------------------------------------------------
Dim oMDL As MIDataLists
Dim oMIMS As MIMsgStatic
Dim vRtn As Variant
Dim nLoop As Integer
Dim sStudyName As String
Dim enType As MIMsgType
'Dim oUser As MACROUser
Dim bDateOK As Boolean
Dim dblDate As Double
Dim sStudySiteSQL As String

    On Error GoTo CatchAllError
    
    'Set oUser = New MACROUser
    'Call oUser.SetState(sSerialisedUser)

    Select Case sType
    Case "0": enType = MIMsgType.mimtDiscrepancy
    Case "2": enType = MIMsgType.mimtNote
    Case Else: enType = MIMsgType.mimtSDVMark
    End Select

    'check passed arguement value
    If (sStudyCode <> "") Then sStudyName = oUser.Studies.StudyById(CLng(sStudyCode)).StudyName
    sSiteCode = Trim(sSiteCode)
    sSubjectLabel = Trim(sSubjectLabel)
    sSrchUserName = Trim(sSrchUserName)
    sSubjectLabel = Trim(sSubjectLabel)
    If Trim(sSubjectId) = "" Then sSubjectId = "-1"
    If Trim(sVisitId) = "" Then sVisitId = "-1"
    If Trim(sCRFPageId) = "" Then sCRFPageId = "-1"
    If Trim(sQuestion) = "" Then sQuestion = "-1"
    'MLM 30/06/05:
    If Trim(sVisitCycle) = "" Then sVisitCycle = "-1"
    If Trim(sCRFPageCycle) = "" Then sCRFPageCycle = "-1"
    If Trim(sQuestionCycle) = "" Then sQuestionCycle = "-1"
    
    If LCase(sBefore) <> "true" Then sBefore = "false"
    'If (Len(sScope) <> 4) Then sScope = "1111"
    dblDate = RtnRecordDblDate(sTime, bDateOK)
    If Not bDateOK And Not IsMissing(vErrors) Then
        vErrors = AddToArray(vErrors, "Search date", "Unable to search on passed format")
    End If
    ' DPH 13/10/2003 Put together study site list for user
    'sStudySiteSQL = oUser.DataLists.StudiesSitesWhereSQL("ClinicalTrial.ClinicalTrialId", "TrialSubject.TrialSite")
    If sStudyCode = "" Then
        ' need all combinations of study/sites available to the user
        sStudySiteSQL = oUser.DataLists.StudiesSitesWhereSQL("ClinicalTrial.ClinicalTrialId", "MIMessage.MIMessageSite")
    Else
        ' using a particular study so limit the SQL
        sStudySiteSQL = "((ClinicalTrial.ClinicalTrialId = " & sStudyCode & ") AND " & oUser.DataLists.StudiesSitesWhereSQL(sStudyCode, "MIMessage.MIMessageSite") & ")"
    End If
    
    Set oMDL = New MIDataLists
    'TA 18/11/2004: flag to show OC Ids isssue 2448
    vRtn = oMDL.GetMIMessageList(oUser.CurrentDBConString, RtnUseOCIdFlag, _
                          oUser.UserName, _
                          sStudySiteSQL, _
                          enType, _
                          RtnMIObjectArray(sScope), _
                          sStudyName, _
                          sSiteCode, _
                          sSubjectLabel, _
                          CLng(sSubjectId), _
                          CLng(sVisitId), _
                          CLng(sCRFPageId), _
                          CLng(sQuestion), _
                          sSrchUserName, _
                          RtnMIStatusArray(enType, sStatus), _
                          CBool(sBefore), _
                          dblDate, , , , _
                          CLng(sVisitCycle), _
                          CLng(sCRFPageCycle), _
                          CLng(sQuestionCycle))
    
    Set oMDL = Nothing
    'Set oUser = Nothing
    RtnMIMessageList = vRtn
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLMIMessage.RtnMIMessageList"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnMIStatusArray(ByVal nType As MIMsgType, ByVal sStatus As String) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 18/12/2001
'   function returns an array representing the statuses requested in a passed binary string
'
'   revisions
'   ic 06/11/2002 sStatus changed to 4 char long string for sdvs, added queried,cancelled chars
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim vRtn As Variant
Dim ndx As Integer
    
    On Error GoTo CatchAllError
    sStatus = Trim(sStatus)
    
    Select Case nType
    Case MIMsgType.mimtDiscrepancy:
        If (Len(sStatus) <> 3) Or (Not IsNumeric(sStatus)) Or (sStatus = "000") Then
             vRtn = ""
        Else
            ReDim vRtn(2)
            
            If Mid(sStatus, 1, 1) <> "0" Then
                vRtn(ndx) = eDiscrepancyMIMStatus.dsRaised
                ndx = ndx + 1
            End If
            If Mid(sStatus, 2, 1) <> "0" Then
                vRtn(ndx) = eDiscrepancyMIMStatus.dsResponded
                ndx = ndx + 1
            End If
            If Mid(sStatus, 3, 1) <> "0" Then
                vRtn(ndx) = eDiscrepancyMIMStatus.dsClosed
                ndx = ndx + 1
            End If
            
            ReDim Preserve vRtn(ndx - 1)
        End If
    
    Case MIMsgType.mimtSDVMark:
        If (Len(sStatus) <> 4) Or (Not IsNumeric(sStatus)) Or (sStatus = "0000") Then
             vRtn = ""
        Else
            ReDim vRtn(3)
            
            If Mid(sStatus, 1, 1) <> "0" Then
                vRtn(ndx) = eSDVMIMStatus.ssPlanned
                ndx = ndx + 1
            End If
            If Mid(sStatus, 2, 1) <> "0" Then
                vRtn(ndx) = eSDVMIMStatus.ssDone
                ndx = ndx + 1
            End If
            If Mid(sStatus, 3, 1) <> "0" Then
                vRtn(ndx) = eSDVMIMStatus.ssQueried
                ndx = ndx + 1
            End If
            If Mid(sStatus, 4, 1) <> "0" Then
                vRtn(ndx) = eSDVMIMStatus.ssCancelled
                ndx = ndx + 1
            End If
            
            ReDim Preserve vRtn(ndx - 1)
        End If
        
    Case Else
        If (Len(sStatus) <> 2) Or (Not IsNumeric(sStatus)) Or (sStatus = "00") Then
             vRtn = ""
        Else
            ReDim vRtn(1)
            
            If Mid(sStatus, 1, 1) <> "0" Then
                vRtn(ndx) = eNoteMIMStatus.nsPublic
                ndx = ndx + 1
            End If
            If Mid(sStatus, 2, 1) <> "0" Then
                vRtn(ndx) = eNoteMIMStatus.nsPrivate
                ndx = ndx + 1
            End If
            
            ReDim Preserve vRtn(ndx - 1)
        End If
    
    End Select

    RtnMIStatusArray = vRtn
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLMIMessage.RtnMIStatusArray"
End Function

'--------------------------------------------------------------------------------------------------
Public Sub UpdateMIMessage(ByRef oUser As MACROUser, ByVal nType As MIMsgType, ByVal sAction As String, _
    ByVal sStudy As String, ByVal lStudyId As Long, ByVal lSubjectId As Long, ByVal lResponseTaskId As Long, _
    ByVal nResponseCycle As Integer, ByVal lId As Long, ByVal nSrc As Integer, ByVal sSite As String, _
    ByVal sText As String, ByVal nTimezoneOffset As Integer, ByVal sASPVToken As String, ByVal sASPEToken As String, _
    ByRef sUpdate As String, ByRef vErrors As Variant)
'--------------------------------------------------------------------------------------------------
'   ic 14/12/2001
'   function updates discrepancys/notes/sdvs
'--------------------------------------------------------------------------------------------------
' DPH 24/02/2003 - Added Planned / Queried / Cancelled Updates
' ic 22/05/2003 added convertfromnull() call to fix bug 1801
' ic 05/03/2004 added subject locking code to prevent updates on subjects being edited by another user
' ic 29/06/2004 added parameter checking, error handling. this function should be moved to modUIHTMLMIMessage
' ic 16/07/2004 moved from clsWWW, modified locking
' ic 27/04/2005 isue 2222, added permission checking for discrepancy/sdv updates
' ic 05/07/2005 changed the eSDVMIMStatus enums to fully qualified
'--------------------------------------------------------------------------------------------------
Dim oDiscrepancy As MIDiscrepancy
Dim oSDV As MISDV
Dim oNote As MINote
Dim oMDL As MIDataLists
Dim nRtn As Integer
Dim vValue As Variant
Dim eSDVStatus As eSDVMIMStatus
Dim sToken As String
Dim bUpdateOK As Boolean
Dim bCreateDisc As Boolean
Dim bCreateSDV As Boolean
Dim bChangeData As Boolean


    On Error GoTo CatchAllError
    
    bCreateDisc = oUser.CheckPermission(gsFnCreateDiscrepancy)
    bCreateSDV = oUser.CheckPermission(gsFnCreateSDV)
    bChangeData = oUser.CheckPermission(gsFnChangeData)
    
    Set oMDL = New MIDataLists
    bUpdateOK = False
    
    Select Case nType
    Case mimtDiscrepancy:
        'load this discrepancy
        Set oDiscrepancy = New MIDiscrepancy
        Call oDiscrepancy.Load(oUser.CurrentDBConString, lId, nSrc, sSite)
        
        'check to see if this user has the eform containg this discrepancy locked
        'if they do, we dont need to lock the subject, otherwise we do
        If (LockSubjectIfNeeded(oUser, lStudyId, sSite, lSubjectId, oDiscrepancy.EFormTaskId, sASPVToken, _
            sASPEToken, sToken, vErrors)) Then
        
            vValue = oMDL.GetResponseDetails(oUser.CurrentDBConString, sStudy, sSite, oDiscrepancy.SubjectId, _
                lResponseTaskId, nResponseCycle)
            Select Case sAction
            Case "0":
                'must have change data permission
                If (bChangeData) Then
                    'set discrepancy to responded, current status must be raised
                    If (oDiscrepancy.CurrentStatus = eDiscrepancyMIMStatus.dsRaised) Then
                        nRtn = oDiscrepancy.Respond(sText, oUser.UserName, oUser.UserNameFull, mimsServer, _
                            CDbl(vValue(eResponseDetails.rdResponseTimeStamp, 0)), _
                            ConvertFromNull(vValue(eResponseDetails.rdResponseValue, 0), vbString), IMedNow, _
                            nTimezoneOffset)
                            bUpdateOK = True
                    Else
                        'if discrepancy is not raised, it was probably changed in another window
                        vErrors = AddToArray(vErrors, "Discrepancy update failed", "This discrepancy cannot be set to responded")
                    End If
                Else
                    vErrors = AddToArray(vErrors, "Discrepancy update failed", "You do not have permission to respond to this discrepancy")
                End If
                        
            Case "1":
                'must have create disc permission
                If (bCreateDisc) Then
                    're-raise discrepancy, current status must be responded
                    If (oDiscrepancy.CurrentStatus = eDiscrepancyMIMStatus.dsResponded) Then
                        nRtn = oDiscrepancy.ReRaise(sText, oUser.UserName, oUser.UserNameFull, mimsServer, _
                            CDbl(vValue(eResponseDetails.rdResponseTimeStamp, 0)), _
                            ConvertFromNull(vValue(eResponseDetails.rdResponseValue, 0), vbString), IMedNow, _
                            nTimezoneOffset)
                            bUpdateOK = True
                    Else
                        'if discrepancy is not responded, it was probably changed in another window
                        vErrors = AddToArray(vErrors, "Discrepancy update failed", "This discrepancy cannot be set to raised")
                    End If
                Else
                    vErrors = AddToArray(vErrors, "Discrepancy update failed", "You do not have permission to re-raise this discrepancy")
                End If
            Case "2":
                'must have create disc permission
                If (bCreateDisc) Then
                    'close discrepancy, current status must be responded, raised
                    If (oDiscrepancy.CurrentStatus = eDiscrepancyMIMStatus.dsResponded) _
                    Or (oDiscrepancy.CurrentStatus = eDiscrepancyMIMStatus.dsRaised) Then
                        nRtn = oDiscrepancy.CloseDown(sText, oUser.UserName, oUser.UserNameFull, mimsServer, _
                            CDbl(vValue(eResponseDetails.rdResponseTimeStamp, 0)), _
                            ConvertFromNull(vValue(eResponseDetails.rdResponseValue, 0), vbString), IMedNow, _
                            nTimezoneOffset)
                            bUpdateOK = True
                    Else
                        'if discrepancy is not responded or raised, it was probably changed in another window
                        vErrors = AddToArray(vErrors, "Discrepancy update failed", "This discrepancy cannot be set to closed")
                    End If
                Else
                    vErrors = AddToArray(vErrors, "Discrepancy update failed", "You do not have permission to close this discrepancy")
                End If
            Case "3":
                'change text of discrepancy, current status must be responded, raised, closed
                If (oDiscrepancy.CurrentStatus = eDiscrepancyMIMStatus.dsResponded) _
                Or (oDiscrepancy.CurrentStatus = eDiscrepancyMIMStatus.dsRaised) _
                Or (oDiscrepancy.CurrentStatus = eDiscrepancyMIMStatus.dsClosed) Then
                    If (oDiscrepancy.CurrentMessage.TimeSent = 0) And (oDiscrepancy.CurrentMessage.UserName = _
                    oUser.UserName) And (Len(sText) <= 2000) Then
                        Call oDiscrepancy.SetText(sText, oUser.UserName)
                    Else
                        'if changing discrepancy text is not permitted
                        vErrors = AddToArray(vErrors, "Discrepancy update failed", "Discrepancy transmitted/username mismatch/text too long")
                    End If
                    bUpdateOK = True
                Else
                    'if discrepancy is already raised, it was probably changed in another window
                    vErrors = AddToArray(vErrors, "Discrepancy update failed", "This discrepancy has an unrecognised status")
                End If
            End Select
            
            If bUpdateOK Then
                'commit save
                oDiscrepancy.Save
                
                'update status
                Call UpdateMIMsgStatus(oUser.CurrentDBConString, mimtDiscrepancy, sStudy, lStudyId, sSite, _
                    oDiscrepancy.SubjectId, oDiscrepancy.VisitId, oDiscrepancy.VisitCycle, oDiscrepancy.EFormTaskId, _
                    oDiscrepancy.ResponseTaskId, oDiscrepancy.ResponseCycle)
                                    
                'create update string to update an eform, if one is open
                sUpdate = Chr(34) & sSite & Chr(34) & "," & CStr(lStudyId) & "," & CStr(lSubjectId) & "," _
                & oDiscrepancy.EFormTaskId & "," & Chr(34) & "f_" & oDiscrepancy.EFormId & "_" & oDiscrepancy.QuestionId & Chr(34) & "," _
                & (oDiscrepancy.ResponseCycle - 1) & ",0," & RtnDiscrepancyStatusCode(oDiscrepancy.CurrentStatus)
            End If
        End If
        Set oDiscrepancy = Nothing
        
    Case mimtNote:
        'load this note
        Set oNote = New MINote
        Call oNote.Load(oUser.CurrentDBConString, lId, nSrc, sSite)
        
        'check to see if this user has the eform containg this note locked
        'if they do, we dont need to lock the subject, otherwise we do
        If (LockSubjectIfNeeded(oUser, lStudyId, sSite, lSubjectId, oNote.EFormTaskId, sASPVToken, _
            sASPEToken, sToken, vErrors)) Then
        
            If (oNote.CurrentMessage.TimeSent = 0) And (oNote.CurrentMessage.UserName = oUser.UserName) _
                And (Len(sText) <= 2000) Then
                Call oNote.SetText(sText, oUser.UserName)
                bUpdateOK = True
            Else
                'if changing note text is not permitted
                vErrors = AddToArray(vErrors, "Note update failed", "Note transmitted/username mismatch/text too long")
            End If
            
            If bUpdateOK Then
                'commit save
                oNote.Save
            End If
        End If
        Set oNote = Nothing
    
    Case mimtSDVMark:
        'load this sdv
        Set oSDV = New MISDV
        Call oSDV.Load(oUser.CurrentDBConString, lId, nSrc, sSite)
        
        'check to see if this user has the eform containg this sdv locked
        'if they do, we dont need to lock the subject, otherwise we do
        If (LockSubjectIfNeeded(oUser, lStudyId, sSite, lSubjectId, oSDV.EFormTaskId, sASPVToken, _
            sASPEToken, sToken, vErrors)) Then
        
            'get details of response, if sdv is on a response
            vValue = oMDL.GetResponseDetails(oUser.CurrentDBConString, sStudy, sSite, oSDV.SubjectId, lResponseTaskId, nResponseCycle)
        
            Select Case sAction
            Case "0":
                'must have create sdv permission
                If (bCreateSDV) Then
                    If (oSDV.CurrentStatus <> eSDVMIMStatus.ssDone) Then
                        If (Not IsNull(vValue)) Then
                            'response sdv update
                            nRtn = oSDV.Done(sText, oUser.UserName, oUser.UserNameFull, mimsServer, IMedNow, _
                                nTimezoneOffset, CDbl(vValue(eResponseDetails.rdResponseTimeStamp, 0)), _
                                ConvertFromNull(vValue(eResponseDetails.rdResponseValue, 0), vbString))
                            bUpdateOK = True
                        Else
                            'eform sdv update
                            nRtn = oSDV.Done(sText, oUser.UserName, oUser.UserNameFull, mimsServer, IMedNow, _
                                nTimezoneOffset)
                            bUpdateOK = True
                        End If
                    Else
                        'if SDV is already done, it was probably changed in another window
                        vErrors = AddToArray(vErrors, "SDV update failed", "This SDV cannot be set to done")
                    End If
                Else
                    vErrors = AddToArray(vErrors, "SDV update failed", "You do not have permission to set this SDV to done")
                End If
                
            Case "1":
                If (oSDV.CurrentMessage.TimeSent = 0) And (oSDV.CurrentMessage.UserName = oUser.UserName) _
                And (Len(sText) <= 2000) Then
                    Call oSDV.SetText(sText, oUser.UserName)
                    bUpdateOK = True
                Else
                    'if changing note text is not permitted
                    vErrors = AddToArray(vErrors, "SDV update failed", "SDV transmitted/username mismatch/text too long")
                End If
                
            Case "6", "7", "8":
                Select Case sAction
                Case "6"
                    'planned update
                    eSDVStatus = eSDVMIMStatus.ssPlanned
                Case "7"
                    'queried update
                    eSDVStatus = eSDVMIMStatus.ssQueried
                Case "8"
                    'cancelled update
                    eSDVStatus = eSDVMIMStatus.ssCancelled
                End Select
                
                If (oSDV.CurrentStatus = eSDVMIMStatus.ssCancelled) And (eSDVStatus = eSDVMIMStatus.ssCancelled) Then
                    'if SDV is already cancelled, it was probably changed in another window
                    vErrors = AddToArray(vErrors, "SDV update failed", "This SDV cannot be set to cancelled")
                ElseIf (oSDV.CurrentStatus = eSDVMIMStatus.ssDone) And (eSDVStatus = eSDVMIMStatus.ssDone) Then
                    'if SDV is already done, it was probably changed in another window
                    vErrors = AddToArray(vErrors, "SDV update failed", "This SDV cannot be set to done")
                ElseIf (oSDV.CurrentStatus = eSDVMIMStatus.ssPlanned) And (eSDVStatus = eSDVMIMStatus.ssPlanned) Then
                    'if SDV is already planned, it was probably changed in another window
                    vErrors = AddToArray(vErrors, "SDV update failed", "This SDV cannot be set to planned")
                ElseIf (oSDV.CurrentStatus = eSDVMIMStatus.ssQueried) And (eSDVStatus = eSDVMIMStatus.ssQueried) Then
                    'if SDV is already queried, it was probably changed in another window
                    vErrors = AddToArray(vErrors, "SDV update failed", "This SDV cannot be set to queried")
                Else
                    'must have create sdv permission
                    If (bCreateSDV) Then
                        If (Not IsNull(vValue)) Then
                            'response sdv update
                            nRtn = oSDV.ChangeStatus(eSDVStatus, sText, oUser.UserName, oUser.UserNameFull, _
                                mimsServer, IMedNow, nTimezoneOffset, CDbl(vValue(eResponseDetails.rdResponseTimeStamp, 0)), _
                                ConvertFromNull(vValue(eResponseDetails.rdResponseValue, 0), vbString))
                        Else
                            'eform sdv update
                            nRtn = oSDV.ChangeStatus(eSDVStatus, sText, oUser.UserName, oUser.UserNameFull, mimsServer, _
                                IMedNow, nTimezoneOffset)
                        End If
                        bUpdateOK = True
                    Else
                        vErrors = AddToArray(vErrors, "SDV update failed", "You do not have permission to change this SDV status")
                    End If
                End If
                
            End Select
            
            If bUpdateOK Then
                'commit save
                oSDV.Save
                
                Call UpdateMIMsgStatus(oUser.CurrentDBConString, mimtSDVMark, sStudy, lStudyId, sSite, _
                    oSDV.SubjectId, oSDV.VisitId, oSDV.VisitCycle, oSDV.EFormTaskId, oSDV.ResponseTaskId, _
                    oSDV.ResponseCycle)
                 
                'create update string to update an eform, if one is open
                sUpdate = Chr(34) & sSite & Chr(34) & "," & CStr(lStudyId) & "," & CStr(lSubjectId) & "," _
                & oSDV.EFormTaskId & "," & Chr(34) & "f_" & oSDV.EFormId & "_" & oSDV.QuestionId & Chr(34) & "," _
                & (oSDV.ResponseCycle - 1) & ",1," & RtnSDVStatusCode(oSDV.CurrentStatus)
            End If
        End If
        Set oSDV = Nothing
    
    End Select
    
    
    If (sToken <> "") Then
        'we locked the subject, now unlock it
        Call UnlockSubjectA(oUser, lStudyId, sSite, lSubjectId, sToken, vErrors)
    End If
    
    Set oMDL = Nothing
    Exit Sub
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|clsWWW.UpdateMIMessage"
End Sub

'--------------------------------------------------------------------------------------------------
Private Function RtnDiscrepancyStatusCode(ByVal nEnumCode As Integer) As Integer
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    Select Case nEnumCode
    Case eDiscrepancyMIMStatus.dsClosed: RtnDiscrepancyStatusCode = 0
    Case eDiscrepancyMIMStatus.dsRaised: RtnDiscrepancyStatusCode = 30
    Case eDiscrepancyMIMStatus.dsResponded: RtnDiscrepancyStatusCode = 20
    End Select
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnSDVStatusCode(ByVal nEnumCode As Integer) As Integer
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    Select Case nEnumCode
    Case eSDVMIMStatus.ssCancelled: RtnSDVStatusCode = 0
    Case eSDVMIMStatus.ssDone: RtnSDVStatusCode = 20
    Case eSDVMIMStatus.ssPlanned: RtnSDVStatusCode = 30
    Case eSDVMIMStatus.ssQueried: RtnSDVStatusCode = 40
    End Select
End Function
