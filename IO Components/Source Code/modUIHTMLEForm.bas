Attribute VB_Name = "modUIHTMLEForm"
'----------------------------------------------------------------------------------------'
'   File:       modUIHTMLEForm.bas
'   Copyright:  InferMed Ltd. 2000-2007 All Rights Reserved
'   Author:     i curtis 02/2003
'   Purpose:    functions returning html versions of MACRO pages (EFORM)
'----------------------------------------------------------------------------------------'
'   revisions
'   ic 10/06/2003 'next' field should always be first of group in GetEformBody()
'   ic 11/06/2003 replace linefeeds in RtnValidationString() as they cause a javascript crash
'   ic 11/06/2003 fixed bug 1681 - lab question validations
'   ic 19/06/2003 added registration
'   dph 20/06/2003 add font style to field template
'   MLM 26/06/03: 3.0 bug list 1338: Only show fixed form and visit date captions if the study definition does not specify custom ones.
'   ic 26/06/2003 convert server locale to standard format, bug 1032&1033
'   ic 05/08/2003 check for changedata permissions in RtnJSVEInit(), bug 1938
'   ic 21/08/2003 bug 1972, add eformid & eformtaskid for visit eform to RtnJSVEInit()
'   ic 27/08/2003 change true/false to jsboolean, disable multimedia in GetEFIInitialisation()
'   ic 27/08/2003 moved eform code from clsWWW
'   dph 27/08/2003 added missing parameters to fnCreateFieldTemplate call for a missing response in GetEFIInitialisation
'   dph 27/08/2003 - Status default "requested" for no response in GetEFIInitialisation
'   ic 29/08/2003 add validation that will always fire for questions with validation failures on readonly eforms in GetEFIInitialisation()
'   dph 29/08/2003 - Added Hidden parameter to field template & removed from field instance in GetEFIInitialisation
'   dph 01/09/2003 - Added Overrule Warnings permission, bug 1986 to GetEformBody
'   ic 01/09/2003 tidied createinstance function call
'   ic 02/09/2003 only try save if we get a writable eform on loadresponses in SaveForm()
'   ic 02/09/2003 added new CanChangeData function to various functions
'   ic 16/09/2003 convert reals from browser locale to server locale in SaveForm()
'   ic 16/09/2003 commented out unused funcs: RtnEformElement(),RtnRadioCreateFuncs(),RtnEformCategoryTxtFunc,
'                 RtnFontSize()
'   ic 18/09/2003 added 'Initialising eForm' message, pause in GetEformBody()
'   ic 18/09/2003 rewrote GetEFIInitialisation() for clarity
'   ic 10/10/2003 added eternal loop check to LoopThroughQuestions() function
'   DPH 02/12/2003 Change order form is drawn (visit eform 1st if exists - then normal eform) in GetEformBody
'   ic 06/01/2004 added eform printing functions including GetEformPrint()
'   ic 13/01/2004 added mandatory flag to rqg initialisation
'   ic 25/02/2004 in GetEFIInitialisation() allow 4 types for hidden fields: integer,real,lab,text
'   ic 27/02/2004 GetEformBody(): added 'no responses on eform' flag to fnInitialiseApplet() call
'   ic 02/03/2004 revalidation now done on client for eform open, client & server for eform close
'   ic 12/03/2004 added RtnInitialFocus() function
'   ic 18/03/2004 added RaiseMIMessages() function, added inform icon handling to GetEformPrint()
'   DPH 13/04/2004 - Allow for scrollbar height if last element on eForm in CalculateRQGMaxHeight
'   ic 15/04/2004   removed EformIsEmpty() function and call - no longer needed client side
'   ic 16/04/2004 removed 'overrule warnings' parameter from fnInitialiseApplet() js call
'   ic 20/04/2004 dont alert missing rfcs in ProcessQuestion() as this signifies a bug in the jsve
'   ic 21/04/2004 revised handling of rejected values
'   ic 23/04/2004 allow 5 types for hidden fields: integer,real,lab,date,text
'   ic 06/05/2004 added ChangeDoneSDVsToPlanned() call in SaveForm() function
'   ic 29/06/2004 added error handling
'   ic 12/07/2004 reinstated sdv fix
'   ic 20/07/2004 amended form onunload in GetEformBody()
'   ic 08/11/2004 bug 2395, added setsetting for SETTING_LAST_USED_STUDY in GetEformBody()
'   NCJ 22 Nov 04 - Issue 2451 - Do not check Subject SDV status before changing Done SDVs in SaveForm
'   ic 27/05/2005 issue 2579, ignore comments/lines/pictures for skips/derivations/validations
'   ic 24/05/2005 issue 2566, ignore hidden fields in ProcessNotEnterableQuestions()
'   ic 20/06/2005 issue 2593, changed javascript call in GetEformBody() function
'   ic 28/07/2005 added clinical coding
'   NCJ 26 May 06 - Issue 2740 - Use CancelResponses in GetEFormBody to ensure AREZZO values are cleared
'   ic 15/06/2005 issue 2734 - derivations with line breaks cause an error in javascript
'   NCJ 28 Jun 06 - Refixing Issue 2740 - Must preserve eForm locks!
'   ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
'   NCJ 27 Sept 07 - Bug 2935 - Preserve UserNameFull for derived questions (in calls to RefreshSkips)
'   ic 28/02/2007 issue 2855 - clinical coding release
'   ic 28/02/2008 issue 2996 - add code to add inactive categories that are the current saved value
'----------------------------------------------------------------------------------------'
Option Explicit

'internet explorer scale
' DPH 13/05/2003 - changed X scale to make page width smaller
Private Const nIeXSCALE = 15 '12.5
Private Const nIeYSCALE = 15

' DPH 26/11/2002 Spacing TabIndex Constant for RQG
Private Const mnTABINDEXGAP = 200
Private Const msAZTOJSERROR As String = "||"

'ic 01/04/2003 temporary showcode variable
Private Const mbShowCode As Boolean = False

Private Const mCCSwitch As String = "CLINICALCODING"

'ic 28/02/2008 a global string that lets us add inactive category values
'to select lists when they are the saved value
Private gsAddInactiveCategory As String

'--------------------------------------------------------------------------------------------------
Public Function GetEformPrint(ByRef oUser As MACROUser, ByVal sSiteCode As String, ByVal sStudyCode As String, _
    ByVal sSubjectId As String, ByVal sVisitCode As String, ByVal sVEformId As String, ByVal sVTEformId As String, _
    ByVal sEformId As String, ByVal sTEformId As String, ByVal sDecimalPoint As String, ByVal sThousandSeparator) As String
'--------------------------------------------------------------------------------------------------
'   ic 16/12/2003
'   function returns an eform html page suitable for printing through internet explorer
'   revisions
'   ic 18/03/2004 added inform icon check
'   ic 29/06/2004 added error handling
'   ic 28/07/2005 added clinical coding
'--------------------------------------------------------------------------------------------------
Dim vJSComm() As String
Dim vVisitEform As Variant
Dim vUserEform As Variant
Dim vUEformElements As Variant
Dim vVEformElements As Variant
Dim n As Integer
Dim nUFirstIndex As Integer
Dim nVFirstIndex As Integer
Dim sSTitle As String
Dim sSStatus As String
Dim sVTitle As String
Dim sVStatus As String
Dim sETitle As String
Dim sEStatus As String
Dim sCaption As String
Dim sValue As String
Dim sStudyName As String
Dim vEformDetails As Variant
Dim bViewInform As Boolean

    On Error GoTo CatchAllError

    ReDim vJSComm(0)
    
    'html head
    Call AddStringToVarArr(vJSComm, "<html>" & vbCrLf _
        & "<head>" & vbCrLf _
        & "<link rel='stylesheet' HREF='../style/MACRO1.css' type='text/css'>" _
        & "<title>eForm Print</title>" _
        & "<script language='javascript'>" & vbCrLf _
        & "function fnPrint(){" & vbCrLf _
        & "window.print();" & vbCrLf _
        & "}" & vbCrLf _
        & "</script>" & vbCrLf _
        & "</head><body align='center'>")
    
    
     'if visit id was passed, get visit data array
    If (sVEformId <> "") Then
        'ic 28/07/2005 added clinical coding
        vVisitEform = RtnDataBrowser(oUser, False, CInt(sStudyCode), sSiteCode, CLng(sVisitCode), CLng(sVEformId), _
            0, "ALL", CLng(sSubjectId), "", "", "", "", False, -1, -1, -1, -1, -1, "", "", 1)
    End If
    
    'ic 28/07/2005 added clinical coding
    'get eform data array
    vUserEform = RtnDataBrowser(oUser, False, CInt(sStudyCode), sSiteCode, CLng(sVisitCode), CLng(sEformId), _
        0, "ALL", CLng(sSubjectId), "", "", "", "", False, -1, -1, -1, -1, -1, "", "", 1)

    If (IsNull(vVisitEform) And IsNull(vUserEform)) Then
        'no matching records found at all, this eform hasnt been saved
        Call AddStringToVarArr(vJSComm, "<table>" _
                      & "<tr class='clsMessageText'><td>" _
                      & "This eForm contains no saved printable data" _
                      & "</td></tr>" _
                      & "</table>")
                      
        Call AddStringToVarArr(vJSComm, "</body></html>")
        GetEformPrint = Join(vJSComm, "")
        Exit Function
    End If
    
    
    'get eform element array (so we know which are hidden)
    vUEformElements = RtnElementsOnEform(oUser.CurrentDBConString, CLng(sStudyCode), sEformId)
    'if there is any visit eform data, get visit eform element array
    If (Not IsNull(vVisitEform)) Then
        If Not IsEmpty(vVisitEform) Then
            vVEformElements = RtnElementsOnEform(oUser.CurrentDBConString, CLng(sStudyCode), sVEformId)
        End If
    End If
    
    'find first element in data arrays that applies to this taskid. must also be visible
    'this is because, the RtnDataBrowser(...) search we did above didnt allow us to
    'specify a repeat, so we may have an array of data for several eform cycles
    nUFirstIndex = RtnFirstIndex(vUserEform, vUEformElements, sTEformId)
    nVFirstIndex = RtnFirstIndex(vVisitEform, vVEformElements, sVTEformId)


    If (nUFirstIndex = -1) And (nVFirstIndex = -1) Then
        'no records in either data array apply to this instance or are visible
        Call AddStringToVarArr(vJSComm, "<table>" _
                      & "<tr class='clsMessageText'><td>" _
                      & "This eForm cycle contains no saved printable data" _
                      & "</td></tr>" _
                      & "</table>")
                      
        Call AddStringToVarArr(vJSComm, "</body></html>")
        GetEformPrint = Join(vJSComm, "")
        Exit Function
    End If
      
    bViewInform = oUser.CheckPermission(gsFnMonitorDataReviewData)
    
    'get header information: title, subject, visit, eform
    If nUFirstIndex = -1 And nVFirstIndex <> -1 Then
        'the only printable data records are on the visit eform
        'need to get user eform details from database so we can print the header details
        vEformDetails = RtnEformDetails(oUser.CurrentDBConString, CLng(sStudyCode), sSiteCode, CLng(sSubjectId), CLng(sTEformId))

        'subject
        sSTitle = RtnSubjectText(vVisitEform(DataBrowserCol.dbcSubjectId, nVFirstIndex), _
            vVisitEform(DataBrowserCol.dbcSubjectLabel, nVFirstIndex))
        sSStatus = RtnStatusImages(vVisitEform(DataBrowserCol.dbcsubjectStatus, nVFirstIndex), bViewInform, _
            vVisitEform(DataBrowserCol.dbcSubjectLockStatus, nVFirstIndex), False, _
            vVisitEform(DataBrowserCol.dbcSubjectSDVStatus, nVFirstIndex), _
            vVisitEform(DataBrowserCol.dbcSubjectDiscStatus, nVFirstIndex), _
            RtnJSBoolean(CBool(vVisitEform(DataBrowserCol.dbcSubjectNoteStatus, nVFirstIndex))))
        'visit
        sVTitle = vVisitEform(DataBrowserCol.dbcVisitName, nVFirstIndex) _
            & IIf((CInt(vVisitEform(DataBrowserCol.dbcVisitCycleNumber, nVFirstIndex)) > 1), _
            "[" & vVisitEform(DataBrowserCol.dbcVisitCycleNumber, nVFirstIndex) & "]", "")
        sVStatus = RtnStatusImages(vVisitEform(DataBrowserCol.dbcVisitStatus, nVFirstIndex), bViewInform, _
            vVisitEform(DataBrowserCol.dbcVisitLockStatus, nVFirstIndex), False, _
            vVisitEform(DataBrowserCol.dbcVisitSDVStatus, nVFirstIndex), _
            vVisitEform(DataBrowserCol.dbcVisitDiscStatus, nVFirstIndex), _
            RtnJSBoolean(CBool(vVisitEform(DataBrowserCol.dbcVisitNoteStatus, nVFirstIndex))))
        'eForm
        sETitle = vEformDetails(0, 0) & IIf((CInt(vEformDetails(1, 0)) > 1), "[" & vEformDetails(1, 0) & "]", "")

        sEStatus = RtnStatusImages(vEformDetails(2, 0), False, vEformDetails(3, 0), bViewInform, _
            vEformDetails(4, 0), vEformDetails(5, 0), RtnJSBoolean(CBool(vEformDetails(6, 0))))

        sStudyName = vVisitEform(DataBrowserCol.dbcStudyName, nVFirstIndex)

    Else
        'there are printable data records on the user eform
        'subject
        sSTitle = RtnSubjectText(vUserEform(DataBrowserCol.dbcSubjectId, nUFirstIndex), _
            vUserEform(DataBrowserCol.dbcSubjectLabel, nUFirstIndex))
        sSStatus = RtnStatusImages(vUserEform(DataBrowserCol.dbcsubjectStatus, nUFirstIndex), bViewInform, _
            vUserEform(DataBrowserCol.dbcSubjectLockStatus, nUFirstIndex), False, _
            vUserEform(DataBrowserCol.dbcSubjectSDVStatus, nUFirstIndex), _
            vUserEform(DataBrowserCol.dbcSubjectDiscStatus, nUFirstIndex), _
            RtnJSBoolean(CBool(vUserEform(DataBrowserCol.dbcSubjectNoteStatus, nUFirstIndex))))
        'visit
        sVTitle = vUserEform(DataBrowserCol.dbcVisitName, nUFirstIndex) _
            & IIf((CInt(vUserEform(DataBrowserCol.dbcVisitCycleNumber, nUFirstIndex)) > 1), _
            "[" & vUserEform(DataBrowserCol.dbcVisitCycleNumber, nUFirstIndex) & "]", "")
        sVStatus = RtnStatusImages(vUserEform(DataBrowserCol.dbcVisitStatus, nUFirstIndex), bViewInform, _
            vUserEform(DataBrowserCol.dbcVisitLockStatus, nUFirstIndex), False, _
            vUserEform(DataBrowserCol.dbcVisitSDVStatus, nUFirstIndex), _
            vUserEform(DataBrowserCol.dbcVisitDiscStatus, nUFirstIndex), _
            RtnJSBoolean(CBool(vUserEform(DataBrowserCol.dbcVisitNoteStatus, nUFirstIndex))))
        'eForm
        sETitle = vUserEform(DataBrowserCol.dbcEFormTitle, nUFirstIndex) _
            & IIf((CInt(vUserEform(DataBrowserCol.dbcEFormCycleNumber, nUFirstIndex)) > 1), _
            "[" & vUserEform(DataBrowserCol.dbcEFormCycleNumber, nUFirstIndex) & "]", "")
        
        sEStatus = RtnStatusImages(vUserEform(DataBrowserCol.dbcEFormStatus, nUFirstIndex), bViewInform, _
            vUserEform(DataBrowserCol.dbcEFormLockStatus, nUFirstIndex), False, _
            vUserEform(DataBrowserCol.dbcEFormSDVStatus, nUFirstIndex), _
            vUserEform(DataBrowserCol.dbcEFormDiscStatus, nUFirstIndex), _
            RtnJSBoolean(CBool(vUserEform(DataBrowserCol.dbcEFormNoteStatus, nUFirstIndex))))
        
        sStudyName = vUserEform(DataBrowserCol.dbcStudyName, 0)
    End If
        
      
    'close, print links
    Call AddStringToVarArr(vJSComm, "<table width='100%' border='0'>" _
        & "<tr height='10'><td></td></tr>" & vbCrLf _
        & "<tr height='15' class='clsLabelText'>" _
        & "<td align='right'><a style='cursor:hand;' onclick='javascript:window.close();'><u>Close</u></a>&nbsp;" _
        & "<a style='cursor:hand;' onclick='javascript:fnPrint();'><u>Print</u></a>&nbsp;</td>" _
        & "</tr>")

    Call AddStringToVarArr(vJSComm, "<tr height='10'><td></td></tr>" _
        & "<tr><td>")
            
    'header table start
    'database/site/study
    Call AddStringToVarArr(vJSComm, "<table border='0' width='100%'>" _
        & "<tr class='clsEformPrintHeadText'><td width='100%' align='center' colspan='3'>" _
        & oUser.DatabaseCode & "/" & sSiteCode & "/" & sStudyName _
        & "</td></tr><tr height='10'><td></td></tr>")

    'subject
    Call AddStringToVarArr(vJSComm, "<tr class='clsEformPrintText clsEformPrint'>" _
        & "<td width='20%'>Subject</td>" _
        & "<td width='70%'>" & sSTitle & "</td>" _
        & "<td>" & sSStatus & "</td>" _
        & "</tr>")

    'visit
    Call AddStringToVarArr(vJSComm, "<tr class='clsEformPrintText clsEformPrint'>" _
        & "<td>Visit</td>" _
        & "<td>" & sVTitle & "</td>" _
        & "<td>" & sVStatus & "</td>" _
        & "</tr>")

    'eform
    Call AddStringToVarArr(vJSComm, "<tr class='clsEformPrintText clsEformPrint'>" _
        & "<td>eForm</td>" _
        & "<td>" & sETitle & "</td>" _
        & "<td>" & sEStatus & "</td>" _
        & "</tr>")

    'header table end
    Call AddStringToVarArr(vJSComm, "</table>")
                        
    Call AddStringToVarArr(vJSComm, "<tr height='10'><td></td></tr></td></tr><tr><td>")
        
    Call AddStringToVarArr(vJSComm, "<table width='100%' cellpadding='0' cellspacing='0' border='1' class='clsEformPrintText clsEformPrint'>" _
        & "<tr height='25'><td width='20%'>Question</td><td width='50%'>Value</td><td width='10%'>Status</td></tr>")


    'loop through all visit eform data records adding a row for each
    If (Not IsNull(vVisitEform)) Then
        If (Not IsEmpty(vVisitEform)) Then
            For n = LBound(vVisitEform, 2) To UBound(vVisitEform, 2)
                If (vVisitEform(DataBrowserCol.dbcEFormTaskID, n) = sVTEformId) Then
                    If IsVisible(vVEformElements, vVisitEform(DataBrowserCol.dbcEFormElementId, n), sCaption) Then
                        If Not IsNull(vVisitEform(DataBrowserCol.dbcResponseValue, n)) Then
                            sValue = vVisitEform(DataBrowserCol.dbcResponseValue, n)
                        Else
                            sValue = ""
                        End If
                    
                    
                        Call AddStringToVarArr(vJSComm, "<tr>" _
                            & "<td>" & vVisitEform(DataBrowserCol.dbcDataItemName, n) & IIf((CInt(vVisitEform(DataBrowserCol.dbcResponseCycleNumber, n)) > 1), "[" & vVisitEform(DataBrowserCol.dbcResponseCycleNumber, n) & "]", "") & "</td>" _
                            & "<td>")
                            
                            If (vVisitEform(DataBrowserCol.dbcDataType, n) = DataType.Multimedia) Then
                                If (sValue <> "") Then Call AddStringToVarArr(vJSComm, "(attached)")
                            Else
                                Call AddStringToVarArr(vJSComm, ReplaceWithHTMLCodes(LocaliseValue(sValue, CInt(vVisitEform(DataBrowserCol.dbcDataType, n)), sDecimalPoint, sThousandSeparator)) & vbCrLf)
                            End If

                            Call AddStringToVarArr(vJSComm, "&nbsp;</td>" _
                            & "<td>" & RtnStatusImages(vVisitEform(DataBrowserCol.dbcResponseStatus, n), bViewInform, _
                                vVisitEform(DataBrowserCol.dbcDataItemLockStatus, n), False, vVisitEform(DataBrowserCol.dbcDataItemSDVStatus, n), _
                                vVisitEform(DataBrowserCol.dbcDataItemDiscStatus, n), _
                                RtnJSBoolean(CBool(vVisitEform(DataBrowserCol.dbcDataItemNoteStatus, n))), RtnJSBoolean(CBool(ConvertFromNull(vVisitEform(DataBrowserCol.dbcComments, n), vbString) <> "")), _
                                vVisitEform(DataBrowserCol.dbcChangeCount, n), "") & "&nbsp;" & RtnNRCTC(vVisitEform(DataBrowserCol.dbcResponseStatus, n), vVisitEform(DataBrowserCol.dbcLabResult, n), vVisitEform(DataBrowserCol.dbcCTCGrade, n)) & "</td>" _
                            & "</tr>")
                    End If
                End If
            Next
        End If
    End If

    'loop through all user eform data records adding a row for each
    If (Not IsNull(vUserEform)) Then
        If (Not IsEmpty(vUserEform)) Then
            For n = LBound(vUserEform, 2) To UBound(vUserEform, 2)
                If (vUserEform(DataBrowserCol.dbcEFormTaskID, n) = sTEformId) Then
                    If IsVisible(vUEformElements, vUserEform(DataBrowserCol.dbcEFormElementId, n), sCaption) Then
                        If Not IsNull(vUserEform(DataBrowserCol.dbcResponseValue, n)) Then
                            sValue = vUserEform(DataBrowserCol.dbcResponseValue, n)
                        Else
                            sValue = ""
                        End If
                    
                    
                        Call AddStringToVarArr(vJSComm, "<tr>" _
                            & "<td>" & vUserEform(DataBrowserCol.dbcDataItemName, n) & IIf((CInt(vUserEform(DataBrowserCol.dbcResponseCycleNumber, n)) > 1), "[" & vUserEform(DataBrowserCol.dbcResponseCycleNumber, n) & "]", "") & "</td>" _
                            & "<td>")
                            
                            If (vUserEform(DataBrowserCol.dbcDataType, n) = DataType.Multimedia) Then
                                If (sValue <> "") Then Call AddStringToVarArr(vJSComm, "(attached)")
                            Else
                                Call AddStringToVarArr(vJSComm, ReplaceWithHTMLCodes(LocaliseValue(sValue, CInt(vUserEform(DataBrowserCol.dbcDataType, n)), sDecimalPoint, sThousandSeparator)) & vbCrLf)
                            End If

                            Call AddStringToVarArr(vJSComm, "&nbsp;</td>" _
                            & "<td>" & RtnStatusImages(vUserEform(DataBrowserCol.dbcResponseStatus, n), bViewInform, _
                                vUserEform(DataBrowserCol.dbcDataItemLockStatus, n), False, vUserEform(DataBrowserCol.dbcDataItemSDVStatus, n), _
                                vUserEform(DataBrowserCol.dbcDataItemDiscStatus, n), _
                                RtnJSBoolean(CBool(vUserEform(DataBrowserCol.dbcDataItemNoteStatus, n))), RtnJSBoolean(CBool(ConvertFromNull(vUserEform(DataBrowserCol.dbcComments, n), vbString) <> "")), _
                                vUserEform(DataBrowserCol.dbcChangeCount, n), "") & "&nbsp;" & RtnNRCTC(vUserEform(DataBrowserCol.dbcResponseStatus, n), vUserEform(DataBrowserCol.dbcLabResult, n), vUserEform(DataBrowserCol.dbcCTCGrade, n)) & "</td>" _
                            & "</tr>")
                    End If
                End If
            Next
        End If
    End If

    Call AddStringToVarArr(vJSComm, "</table>")
    
    Call AddStringToVarArr(vJSComm, "</td></tr>")
    
    Call AddStringToVarArr(vJSComm, "<tr height='10'><td></td></tr>" & vbCrLf _
        & "<tr height='15' class='clsLabelText'>" _
        & "<td align='right'><a style='cursor:hand;' onclick='javascript:window.close();'><u>Close</u></a>&nbsp;" _
        & "<a style='cursor:hand;' onclick='javascript:fnPrint();'><u>Print</u></a>&nbsp;</td>" _
        & "</tr></table>")
    
    
    Call AddStringToVarArr(vJSComm, "</body></html>")

    GetEformPrint = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.GetEformPrint"
End Function

'--------------------------------------------------------------------------------------------------
Private Function IsVisible(ByVal vEformArray As Variant, ByVal sElementID As String, ByRef sCaption As String) As Boolean
'--------------------------------------------------------------------------------------------------
'   ic 06/01/2004
'   function checks if a passed element id specifies a visible element in a passed element array
'   if true, also sets byref caption value
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim n As Integer
Dim bVisible

    On Error GoTo CatchAllError

    bVisible = False
    For n = LBound(vEformArray, 2) To UBound(vEformArray, 2)
        If (vEformArray(0, n) = sElementID) Then
            bVisible = IIf((vEformArray(2, n) = 0), True, False)
            sCaption = ConvertFromNull(vEformArray(1, n), vbString)
            Exit For
        End If
    Next
    IsVisible = bVisible
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.IsVisible"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnFirstIndex(ByVal vDataArray As Variant, ByVal vEformArray As Variant, ByVal sEformTaskId As String)
'--------------------------------------------------------------------------------------------------
'   ic 06/01/2004
'   function finds the first element index in a passed data array that applies to a passed taskid
'   also checks that the element is visible
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim n As Integer
Dim nIndex As Integer
    
    On Error GoTo CatchAllError
    
    nIndex = -1
    If Not IsNull(vDataArray) Then
        If Not IsEmpty(vDataArray) Then
            For n = LBound(vDataArray, 2) To UBound(vDataArray, 2)
                If (vDataArray(DataBrowserCol.dbcEFormTaskID, n) = sEformTaskId) Then
                    If IsVisible(vEformArray, vDataArray(DataBrowserCol.dbcEFormElementId, n), "") Then
                        nIndex = n
                        Exit For
                    End If
                End If
            Next
        End If
    End If
    RtnFirstIndex = nIndex
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnFirstIndex"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnEformDetails(ByVal sDbCon As String, ByVal lStudyId As Long, ByVal sSite As String, _
    ByVal lSubjectId As Long, ByVal lCRFPageTaskId As Long) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 18/12/2003
'   function returns 2d array of eform details
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oQueryDef As QueryDef
Dim oQueryServer As QueryServer
Dim vRtn As Variant

    On Error GoTo CatchAllError

    Set oQueryDef = New QueryDef
    
    With oQueryDef
        .QueryTables.Add "CRFPAGEINSTANCE"
        .QueryTables.Add "CRFPAGE", , qdjtInner, Array("CRFPAGE.CLINICALTRIALID", "CRFPAGE.CRFPAGEID"), _
                            Array("CRFPAGEINSTANCE.CLINICALTRIALID", "CRFPAGEINSTANCE.CRFPAGEID")
        
        .QueryFields.Add "CRFPAGE.CRFTITLE"
        .QueryFields.Add "CRFPAGEINSTANCE.CRFPAGECYCLENUMBER"
        .QueryFields.Add "CRFPAGEINSTANCE.CRFPAGESTATUS"
        .QueryFields.Add "CRFPAGEINSTANCE.LOCKSTATUS"
        .QueryFields.Add "CRFPAGEINSTANCE.SDVSTATUS"
        .QueryFields.Add "CRFPAGEINSTANCE.DISCREPANCYSTATUS"
        .QueryFields.Add "CRFPAGEINSTANCE.NOTESTATUS"
        
        .QueryFilters.Add "CRFPAGEINSTANCE.CLINICALTRIALID", "=", lStudyId
        .QueryFilters.Add "CRFPAGEINSTANCE.TRIALSITE", "=", sSite
        .QueryFilters.Add "CRFPAGEINSTANCE.PERSONID", "=", lSubjectId
        .QueryFilters.Add "CRFPAGEINSTANCE.CRFPAGETASKID", "=", lCRFPageTaskId
        
    End With
    
    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sDbCon
    vRtn = oQueryServer.SelectArray(oQueryDef)
    
    Set oQueryServer = Nothing
    Set oQueryDef = Nothing
    
    RtnEformDetails = vRtn
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnEformDetails"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnElementsOnEform(ByVal sDbCon As String, ByVal lStudyId As Long, ByVal lCRFPageId As Long) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 18/12/2003
'   function returns 2d array of elements on a specified eform
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oQueryDef As QueryDef
Dim oQueryServer As QueryServer
Dim vRtn As Variant

    On Error GoTo CatchAllError

    Set oQueryDef = New QueryDef
    
    With oQueryDef
        .QueryTables.Add "CRFElement"
        
        .QueryFields.Add "CRFELEMENTID"
        .QueryFields.Add "CAPTION"
        .QueryFields.Add "HIDDEN"
        
        .QueryFilters.Add "CLINICALTRIALID", "=", lStudyId
        .QueryFilters.Add "CRFPAGEID", "=", lCRFPageId
        .QueryFilters.Add "DATAITEMID", ">", 0
        
    End With
    
    Set oQueryServer = New QueryServer
    oQueryServer.ConnectionOpen sDbCon
    vRtn = oQueryServer.SelectArray(oQueryDef)
    
    Set oQueryServer = Nothing
    Set oQueryDef = Nothing
    
    RtnElementsOnEform = vRtn
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnElementsOnEform"
End Function
                      
'--------------------------------------------------------------------------------------------------
Private Function RtnRepeatIndex(ByVal colFields As Collection, ByVal sWebID As String) As Integer
'--------------------------------------------------------------------------------------------------
'   ic 30/04/2003
'   function takes a collection of wwwfield objects and a search web id, returns an integer - the
'   number of occurrences of that web id in the collection
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oField As WWWField
Dim nIndex As Integer
    
    On Error GoTo CatchAllError

    nIndex = 1
    For Each oField In colFields
        If (oField.sWebID = sWebID) Then
            nIndex = nIndex + 1
        End If
    Next
    RtnRepeatIndex = nIndex
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnRepeatIndex"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnColFields(ByVal sForm As String) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 30/04/2003
'   function takes an html form string, returns a collection of wwwfield objects
'--------------------------------------------------------------------------------------------------
'   DPH 26/06/2003 - Incremented nNUMCONTROLFIELDS as added localformat date info
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sCodesAndValues As String
Dim vCodesAndValues As Variant
Dim vCodeAndValue As Variant
Dim vCode As Variant
Dim nNumFields As Integer
Dim n As Integer
Dim oField As WWWField
Dim colFields As Collection
'   DPH 26/06/2003 - Incremented nNUMCONTROLFIELDS as added localformat date info
Const nNUMCONTROLFIELDS = 4

    On Error GoTo CatchAllError

'    sCodesAndValues = ReplaceHTMLCodes(sForm)
    'split the form elements on '&' delimiter to get an array of 'f_eformid_elementid=value'
    vCodesAndValues = Split(sForm, "&")
    'get count of response codes by subtracting number of eform control fields located at end of form
    nNumFields = UBound(vCodesAndValues) - nNUMCONTROLFIELDS
    
    Set colFields = New Collection
    'loop through form response codes
    For n = 0 To nNumFields
        'split on '=' to get an array of 'f_eformid_elementid' and 'value'
        vCodeAndValue = Split(vCodesAndValues(n), "=")
        'split the code on '_' to get an array of 'f' and 'eformid' and 'elementid'
        vCode = Split(vCodeAndValue(0), "_")
        
        'can be 'f' or 'af'
        If (vCode(0) = "f") Then
            Set oField = New WWWField
            oField.sWebID = vCodeAndValue(0) 'f_eformid_elementid
            oField.sEformId = vCode(1) 'eformid
            oField.sElementID = vCode(2) 'elementid
            oField.nRepeat = RtnRepeatIndex(colFields, oField.sWebID) 'previous occurrences
            oField.sValue = ReplaceHTMLCodes(vCodeAndValue(1)) 'value
            oField.sAInfo = RtnAIField(vCodesAndValues, oField.sWebID, oField.nRepeat) 'get associated 'af' field
            Call RtnAIBlockSplit(oField.sAInfo, oField) 'break associated 'af' field into parts
            
            colFields.Add oField
        End If
    Next
    
    Set RtnColFields = colFields
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnColFields"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnAIField(ByVal vCodesAndValues As Variant, ByVal sWebID As String, ByVal nRepeat As Integer) As String
'--------------------------------------------------------------------------------------------------
'   ic 30/04/2003
'   function takes a variant array of 'f_eformid_elementid=value' and a web id and repeat, returns
'   the value for the associated 'af_eformid_elementid=value'(repeat)
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim n As Integer
Dim nIndex As Integer
Dim vCodeAndValue As Variant
Dim sAI As String

    On Error GoTo CatchAllError

    nIndex = 1
    For n = 0 To UBound(vCodesAndValues)
        vCodeAndValue = Split(vCodesAndValues(n), "=")
        If (vCodeAndValue(0) = "a" & sWebID) Then
            If (nIndex = nRepeat) Then
                sAI = ReplaceHTMLCodes(vCodeAndValue(1))
                Exit For
            Else
                nIndex = nIndex + 1
            End If
        End If
    Next
    RtnAIField = sAI
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnAIField"
End Function

'--------------------------------------------------------------------------------------------------
Private Sub RtnAIBlockSplit(ByVal sAIBlock As String, ByRef oField As WWWField)
'--------------------------------------------------------------------------------------------------
'   ic 30/04/2003
'   function takes an AI block and byref wwwfield, updates the wwwfield properties with parameters
'   found in the AI block
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim vAI As Variant
Dim n As Integer
Dim s As String

    On Error GoTo CatchAllError

    vAI = Split(sAIBlock, gsDELIMITER1)
    For n = 0 To UBound(vAI)
        Select Case Left(vAI(n), 1)
        Case "c": oField.sComment = Trim(Mid(vAI(n), 3))
        Case "d": oField.sDiscrepancy = Trim(Mid(vAI(n), 3))
        Case "s": oField.sSDV = Trim(Mid(vAI(n), 3))
        Case "n": oField.sNote = Trim(Mid(vAI(n), 3))
        Case "r": oField.sRFC = Trim(Mid(vAI(n), 3))
        Case "o":
            oField.sRFO = Trim(Mid(vAI(n), 3))
            oField.bRFOPresent = True
        Case "p":
            s = Trim(Mid(vAI(n), 3))
            oField.sAuthUserName = Split(s, gsDELIMITER2)(0)
            oField.sAuthPassword = Split(s, gsDELIMITER2)(1)
        Case "t": oField.dblTimestamp = GetTimeStamp(Trim(Mid(vAI(n), 3)))
        Case "u": oField.sUnobtainable = Trim(Mid(vAI(n), 3))
        End Select
    Next
    Exit Sub

CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnAIBlockSplit"
End Sub

'--------------------------------------------------------------------------------------------------
Private Function IsAuthorisationOK(ByVal sSDBCon As String, ByVal sUserName As String, ByVal sPassword As String, _
                                   ByVal sRole As String, ByVal sDatabase As String, _
                                   Optional ByRef sUserFullName As String = "") As Boolean
'--------------------------------------------------------------------------------------------------
'   ic 30/01/2002
'   accepts a 2d array with one code and value, and a role code
'   returns a boolean, was the user and password sufficent to authorise the question
'--------------------------------------------------------------------------------------------------
'   revisions
'   DPH 18/02/2003 - byref sAuthorisedUser as a full name
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim nInfoLoop As Integer
Dim vAddInfo As Variant
'Dim sPassword As String
Dim oUser As MACROUser
'Dim lRtn As LoginResult
Dim bRtn As Boolean
Dim colRoles As Collection

    On Error GoTo CatchAllError
    bRtn = False
    
    Set oUser = New MACROUser
    'attempt to log this user on with the expected role, if no error occurs, role was ok
    If (oUser.Login(sSDBCon, sUserName, sPassword, "", "MACRO Web Data Entry", "", , sDatabase, sRole, False) = LoginResult.Success) Then
        Set colRoles = New Collection
        Set colRoles = oUser.UserRoles
        If (Not colRoles Is Nothing) Then
            bRtn = CollectionMember(colRoles, sRole, False)
            If bRtn Then sUserFullName = oUser.UserNameFull
            Set colRoles = Nothing
        End If
    End If
    Set oUser = Nothing
    
CatchAllError:
    IsAuthorisationOK = bRtn
End Function

'--------------------------------------------------------------------------------------------------
Private Function GetTimeStamp(sTimeStampBlock As String) As Double
'--------------------------------------------------------------------------------------------------
' RS 22/09/2002
' Extract timestamp out field info and convert to double
' revisions
' ic 22/11/2002 uncommented, added error handler
' ic 17/02/2003 fixed incorrect indices
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim vArray As Variant
Dim dblDate As Double
    
    On Error GoTo CatchAllError
    dblDate = 0

    vArray = Split(sTimeStampBlock, gsDELIMITER2)
    dblDate = CDbl(DateSerial(vArray(2), vArray(1), vArray(0))) + CDbl(TimeSerial(vArray(3), vArray(4), vArray(5)))
    
CatchAllError:
    GetTimeStamp = dblDate
End Function

'--------------------------------------------------------------------------------------------------
Public Function SaveForm(ByRef oUser As MACROUser, ByRef oSubject As StudySubject, ByVal sCRFPageTaskId As String, _
                          ByVal sForm As String, ByRef sEFILockToken As String, ByRef sVILockToken As String, _
                          ByVal bVReadOnly As Boolean, ByVal bEReadOnly As Boolean, ByVal sLabCode As String, _
                          ByVal sDecimalPoint As String, ByVal sThousandSeparator As String, ByRef sRegister As String, _
                          ByVal sLocalDate As String, _
                 Optional ByVal nTimezoneOffset As Integer = 0) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 06/09/01
'   rewrite of SaveForm() fuction
'   function saves the responses contained inside a passed form string. returns a variant array of
'   errors
'   NOTE responses are saved, followed by discrepancies, notes, sdvs, comments. therefore a
'   MIMessage will always be saved with the current value
'--------------------------------------------------------------------------------------------------
'   revisions
'   ic 13/05/2003 set locale separator
'   DPH 29/05/2003 - Don't check derived fields for RFCs
'   dph 17/06/2003 added rfc
'   ic 19/06/2003 added registration
'   DPH 26/06/2003 added localdate
'   ic 26/08/2003 dont set optional fields to missing, set them straight to ok
'   ic 02/09/2003 only try save if we get a writable eform on loadresponses
'   ic 16/09/2003 convert reals from browser locale to server locale
'   DPH 29/09/2003 Rewrite to make readable in multiple procedures and improve performance
'   ic 09/03/2004 rewrite of process question section for clarity
'   ic 06/05/2004 added ChangeDoneSDVsToPlanned() call
'   ic 29/06/2004 added error handling, commented out previous mimessage fix -
'                 this will need to be reinstated
'   ic 12/07/2004 reinstated fix
'   NCJ 25 May 06 - Call CancelResponses in error handler if needed
'   ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
'   ic 28/02/2007 issue 2855 - clinical coding release
'--------------------------------------------------------------------------------------------------
Dim oEFI As EFormInstance
Dim oVEFI As EFormInstance
Dim vRtn As Variant

Dim bChangeData As Boolean
Dim bAddComments As Boolean
Dim bViewComments As Boolean
Dim bAddDiscrepancy As Boolean
Dim bAddSDV As Boolean
Dim bOverruleWarning As Boolean
Dim bRegisterSubject As Boolean
Dim bRegister As Boolean

Dim eLoadOK As eLoadResponsesResult
Dim eSaveOK As eSaveResponsesResult

Dim bResponsesChanged As Boolean

Dim colFields As Collection

Dim lErrNumber As Long
Dim sErrDescription As String
Dim bNeedToCancelAfterError As Boolean
Dim oVersion As MACROVersion.Checker
Dim bSaveCC As Boolean

    ' NCJ 25 May 06 - Will we need to Cancel responses if there's an error?
    bNeedToCancelAfterError = False
    
    bSaveCC = False
    
    bRegister = False
    On Error GoTo CatchAllError

    'get user permissions
    bChangeData = CanChangeData(oUser, oSubject.Site)
    bAddComments = oUser.CheckPermission(gsFnAddIComment)
    bViewComments = oUser.CheckPermission(gsFnViewIComments)
    bAddDiscrepancy = oUser.CheckPermission(gsFnCreateDiscrepancy)
    bAddSDV = oUser.CheckPermission(gsFnCreateSDV)
    bOverruleWarning = oUser.CheckPermission(gsFnOverruleWarnings)
    bRegisterSubject = oUser.CheckPermission(gsFnRegisterSubject)
    
    'set user timezone
    Call oSubject.TimeZone.SetTimezoneOffset(nTimezoneOffset)

    ' DPH 26/06/2003 - use localdate used on eForm
    oSubject.LocalDateFormat = sLocalDate

    'get user and visit eform instances
    Set oEFI = oSubject.eFIByTaskId(CLng(sCRFPageTaskId))
    Set oVEFI = oEFI.VisitInstance.VisitEFormInstance


    'ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
    If (oSubject.Arezzo Is Nothing) Then
        'load responses, without passed lock tokens so that lock tokens arent wiped when the load function
        'thinks we are trying to load responses with a lock for a read only subject
        eLoadOK = oSubject.LoadResponses(oEFI, "", "", "")
        
        'attempt to get an eform lock to confirm that we already have a lock
        Dim sLockToken As String
        
        sLockToken = DBLock(oUser.CurrentDBConString, "", sEFILockToken, oUser.UserName, _
            oSubject.StudyDef.StudyId, oSubject.Site, oSubject.PersonId, oEFI.EFormTaskId)
        sEFILockToken = sLockToken
        
        If Not oVEFI Is Nothing Then
            sLockToken = DBLock(oUser.CurrentDBConString, "", sVILockToken, oUser.UserName, _
                oSubject.StudyDef.StudyId, oSubject.Site, oSubject.PersonId, oVEFI.EFormTaskId)
            sVILockToken = sLockToken
        End If
    Else
        'load responses, using passed lock tokens
        eLoadOK = oSubject.LoadResponses(oEFI, "", sEFILockToken, sVILockToken)
    End If
    
    
    'ic 02/09/2003 only try save if we got a writable eform. if we didnt, another
    'user has removed this user's locks and locked the eform for themselves
    If ((eLoadOK = lrrReadWrite) Or ((oSubject.Arezzo Is Nothing) And (sEFILockToken <> "" Or sVILockToken <> ""))) Then
    
        'get collection of fields from form string
        Set colFields = RtnColFields(sForm)
    
    
        If (bChangeData) Then

            'refresh skips and derivations on eforms
            If Not bEReadOnly Then
                Call oEFI.RefreshSkipsAndDerivations(OpeningEForm, "")
                bNeedToCancelAfterError = True      ' NCJ 25 May 06 - After skips/derivs need a cancel
                
                'save updated lab if any
                If (Not oEFI.ReadOnly) Then
                    oEFI.LabCode = sLabCode
                End If
            End If
            If Not oVEFI Is Nothing And Not bVReadOnly Then
                Call oVEFI.RefreshSkipsAndDerivations(OpeningEForm, "")
                bNeedToCancelAfterError = True      ' NCJ 25 May 06 - After skips/derivs need a cancel
            End If
        
            'process question changes
            Call ProcessQuestions(oUser, bVReadOnly, bEReadOnly, sDecimalPoint, sThousandSeparator, _
                sLocalDate, colFields, oEFI, oVEFI, bOverruleWarning, vRtn)
        
            If Not bEReadOnly Then
                bResponsesChanged = oEFI.Responses.Changed
            
                'commit save
                eSaveOK = oSubject.SaveResponses(oEFI, "")
                
                'ic 28/02/2007 issue 2855 - clinical coding release - save any coded responses
                If (eSaveOK = srrSuccess) Then
                    'check for clinical coding version
                    Set oVersion = New MACROVersion.Checker
                    If (oVersion.HasUpgrade("", mCCSwitch)) Then
                        bSaveCC = True
                    End If
                End If

                If eSaveOK = srrSuccess Then
                    ' NCJ 25 May 06 - After successful save all is OK
                    bNeedToCancelAfterError = False
                End If
                
                ' NCJ 22 Nov 04 - Issue 2451 - Do not check Subject SDV status here
'                If (eSaveOK = srrSuccess) And (bResponsesChanged) And (oSubject.SDVStatus <> ssNone) Then
                If (eSaveOK = srrSuccess) And (bResponsesChanged) Then
                    'change done sdvs to planned status if the data for the question has changed
                    Call ChangeDoneSDVsToPlanned(oUser.CurrentDBConString, oUser.UserName, oUser.UserNameFull, _
                        oSubject, oEFI, mimsServer)
                End If
            End If
        End If
        
    
        'now save mimessages
        Call RaiseMIMessages(oUser, colFields, oSubject, oEFI, oVEFI, bEReadOnly, bVReadOnly, nTimezoneOffset, vRtn)
        
        'attempt registration
        If (bChangeData And bRegisterSubject) Then
            bRegister = DoRegistration(oSubject, oEFI.EFormTaskId, oUser.CurrentDBConString, oUser.DatabaseCode, sRegister)
        End If
        
        If (bSaveCC) Then
            Call oSubject.SaveCodedResponses(oEFI)
        End If
        
    Else
        vRtn = AddToArray(vRtn, "eForm save failed", "Token did not match lock")
    End If

    On Error Resume Next
    Set colFields = Nothing
    ' DPH 26/06/2003 - reset localdate used in subject
    oSubject.LocalDateFormat = ""
    Call oSubject.RemoveResponses(oEFI, True)
    
    
    'ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
    If (oSubject.Arezzo Is Nothing) Then
        'attempt to free eform lock. this is normally done in removeresponses, but in this special
        'case we have to remove them ourselves
        
        Call DBUnlock(oUser.CurrentDBConString, sEFILockToken, oSubject.StudyDef.StudyId, oSubject.Site, _
            oSubject.PersonId, oEFI.EFormTaskId)
        
        If Not oVEFI Is Nothing Then
            Call DBUnlock(oUser.CurrentDBConString, sVILockToken, oSubject.StudyDef.StudyId, oSubject.Site, _
                oSubject.PersonId, oVEFI.EFormTaskId)
        End If
    End If
    
    
    sEFILockToken = ""
    sVILockToken = ""
    Set oEFI = Nothing
    If Not oVEFI Is Nothing Then
        Set oVEFI = Nothing
    End If
    SaveForm = vRtn
    Exit Function
    
    
CatchAllError:
    lErrNumber = Err.Number
    sErrDescription = Err.Description
    
    On Error Resume Next
    Set colFields = Nothing
    ' DPH 26/06/2003 - reset localdate used in subject
    oSubject.LocalDateFormat = ""
    ' NCJ 25 May 06 - Do Cancel if necessary
    If bNeedToCancelAfterError Then
        'ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
        If (Not oSubject.Arezzo Is Nothing) Then
            Call oSubject.CancelResponses(oEFI, True)       ' NCJ 28 Jun 06 - Include Lock argument
        End If
    Else
        Call oSubject.RemoveResponses(oEFI, True)
    End If
    sEFILockToken = ""
    sVILockToken = ""
    Set oEFI = Nothing
    If Not oVEFI Is Nothing Then
        Set oVEFI = Nothing
    End If
    
    Err.Raise lErrNumber, , sErrDescription & "|modUIHTMLEform.SaveForm"
End Function

'--------------------------------------------------------------------------------------------------
Private Sub RaiseMIMessages(ByRef oUser As MACROUser, ByRef colFields As Collection, _
    ByRef oSubject As StudySubject, ByRef oEFI As EFormInstance, ByRef oVEFI As EFormInstance, _
    ByVal bEReadOnly As Boolean, bVReadOnly As Boolean, ByVal nTimezoneOffset As Integer, _
    ByRef vRtn As Variant)
'--------------------------------------------------------------------------------------------------
'   ic 18/03/2004
'   function raises eform mimessages, gets response object first if oWWWField hasnt already got
'   the response - if the user cant change data
'   revisions
'   ic 21/04/2004 revised handling of rejected values  - mimessages are ignored
'   ic 29/06/2004 added error handling, commented out previous mimessage fix - this will
'                 need to be reinstated
'   ic 12/07/2004 reinstated fix
'   ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
'--------------------------------------------------------------------------------------------------
Dim oWWWField As WWWField
Dim bElementFormReadOnly As Boolean
Dim oCRFElement As eFormElementRO

    On Error GoTo CatchAllError

    For Each oWWWField In colFields
        With oWWWField
            If (.bOK) And Not (.bRejected) Then
                If (.oResponse Is Nothing) Then
                    'this question doesnt have a response - cant have entered the ProcessQuestion()
                    'function, user cant have change data permissions. try to get response, but dont
                    'change data by creating rqg rows
                    'set field object property to user or visit eform depending on field parent
                    If (.sEformId = oEFI.eForm.EFormId) Then
                        'user eform
                        If (.oEFormInstance Is Nothing) Then
                            Set .oEFormInstance = oEFI
                            .EformUse = eEFormUse.User
                        End If
                        'state eform was opened in
                        bElementFormReadOnly = bEReadOnly
                    Else
                        If (Not oVEFI Is Nothing) Then
                            If (.sEformId = oVEFI.eForm.EFormId) Then
                                'visit eform
                                If (.oEFormInstance Is Nothing) Then
                                    Set .oEFormInstance = oVEFI
                                    .EformUse = eEFormUse.VisitEForm
                                    bElementFormReadOnly = bVReadOnly
                                End If
                            End If
                        End If
                    End If
                    
                    If (.oEFormInstance Is Nothing) Then
                        'the element isn't on either forms, so report an error
                        vRtn = AddToArray(vRtn, .sEformId, "Field could not be located on eForm [" & .sElementID & "]")
                     
                     
                    'ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
                    'ElseIf ((.oEFormInstance.ReadOnly) Or (bElementFormReadOnly)) Then
                        'dont try to save to read only eforms or eforms that were opened read only
                        
                    ElseIf ((.oEFormInstance.ReadOnly) And (bElementFormReadOnly)) Then
                        'dont try to save to read only eforms and eforms that were opened read only
                    
                    ElseIf ((.oEFormInstance.ReadOnly) And (Not oSubject.Arezzo Is Nothing)) Then
                        'dont try to save to read only eforms where the arezzo is present
            
                    Else
                        'load this eform element
                        Set oCRFElement = .oEFormInstance.eForm.eFormElementByQuestionId(CLng(.sElementID))
                    
                        If oCRFElement Is Nothing Then
                            'the element isn't on the form it said it was; report an error
                            vRtn = AddToArray(vRtn, .sEformId, "Element could not be located on eForm [" & .sElementID & "]")
                        
                        Else
                            'load this response
                            Set .oResponse = .oEFormInstance.Responses.ResponseByElement(oCRFElement, .nRepeat)
                        End If
                    End If
                End If
            
            
                If (Not .oResponse Is Nothing) Then
                    'if we now have a response object, save mimessage
                    Call RaiseMIMessage(oUser, vRtn, .sAInfo, MIMsgScope.mimscQuestion, nTimezoneOffset, _
                        oSubject.StudyDef.Name, oSubject.StudyId, oSubject.Site, oSubject.PersonId, _
                        .oEFormInstance.VisitInstance.VisitId, .oEFormInstance.VisitInstance.CycleNo, _
                        .oEFormInstance.EFormTaskId, .oResponse.EFormInstance.eForm.EFormId, _
                        .oResponse.EFormInstance.CycleNo, .oResponse.ResponseId, .oResponse.RepeatNumber, _
                        .oResponse.TimeStamp, .oResponse.Value, .oResponse.Element.QuestionId, _
                        .oResponse.UserName, oSubject)
                End If
            End If
        End With
    Next
    
    If Not oCRFElement Is Nothing Then Set oCRFElement = Nothing
    Exit Sub
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RaiseMIMessages"
End Sub

'--------------------------------------------------------------------------------------------------
Private Function RtnNotProcessedCount(ByRef colFields As Collection) As Integer
'--------------------------------------------------------------------------------------------------
'   ic 10/10/2003
'   function returns number of fields in collection not marked as 'processed'
'   revisions
'   ic 04/03/2004 renamed for clarity
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim nCount As Integer
Dim oWWWField As WWWField

    On Error GoTo CatchAllError

    nCount = 0
    For Each oWWWField In colFields
        If Not oWWWField.bProcessed Then nCount = nCount + 1
    Next
    
    RtnNotProcessedCount = nCount
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnNotProcessedCount"
End Function

'--------------------------------------------------------------------------------------------------
Private Sub ProcessQuestions(ByRef oUser As MACROUser, ByVal bVReadOnly As Boolean, ByVal bEReadOnly As Boolean, _
    ByVal sDecimalPoint As String, ByVal sThousandSeparator As String, ByVal sLocalDate As String, _
    ByRef colFields As Collection, ByRef oEFI As EFormInstance, ByRef oVEFI As EFormInstance, _
    ByVal bOverruleWarning As Boolean, ByRef vRtn As Variant)
'--------------------------------------------------------------------------------------------------
' Loop through questions performing checks
' reduce calls to RefreshSkipsAndDerivations to minimum (after each block of enterable questions)
' revisions
' ic 10/10/2003 added check for eternal loop
' ic 04/03/2004 revised procedure handling of done fields
' ic 21/04/2004 revised handling of rejected values
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oCRFElement As eFormElementRO
Dim oQGI As QGroupInstance
Dim oWWWField As WWWField
Dim bUReject As Boolean
Dim bVReject As Boolean
Dim nNotProcessedStartCount As Integer
Dim nNotProcessedEndCount As Integer
Dim bElementFormReadOnly As Boolean


    On Error GoTo CatchAllError

    'this loop continues processing all responses not already processed (and marked as 'processed'),
    'then refreshing skips and derivations, until the number of responses not marked as processed
    'doesnt change during the loop. at this point, no further responses will become enterable
    Do
        'get count of fields that are not marked as 'processed' at start of loop
        nNotProcessedStartCount = RtnNotProcessedCount(colFields)
        
        'loop through each passed field
        For Each oWWWField In colFields
            With oWWWField
            
                ' check if field is done
                If Not .bProcessed Then
                
                    'set field object property to user or visit eform depending on field parent
                    If (.sEformId = oEFI.eForm.EFormId) Then
                        'user eform
                        If (.oEFormInstance Is Nothing) Then
                            Set .oEFormInstance = oEFI
                            .EformUse = eEFormUse.User
                        End If
                        'state eform was opened in
                        bElementFormReadOnly = bEReadOnly
                    Else
                        If (Not oVEFI Is Nothing) Then
                            If (.sEformId = oVEFI.eForm.EFormId) Then
                                'visit eform
                                If (.oEFormInstance Is Nothing) Then
                                    Set .oEFormInstance = oVEFI
                                    .EformUse = eEFormUse.VisitEForm
                                    bElementFormReadOnly = bVReadOnly
                                End If
                            End If
                        End If
                    End If
                
                    If (.oEFormInstance Is Nothing) Then
                        'the element isn't on either forms, so report an error
                        vRtn = AddToArray(vRtn, .sEformId, "Field could not be located on eForm [" & .sElementID & "]")
                        .bOK = False
                        .bProcessed = True
                    
                    ElseIf ((.oEFormInstance.ReadOnly) Or (bElementFormReadOnly)) Then
                        'dont try to save to read only eforms or eforms that were opened read only
                        .bProcessed = True
                    
                    Else
                        'load this eform element
                        Set oCRFElement = .oEFormInstance.eForm.eFormElementByQuestionId(CLng(.sElementID))
                            
                        If oCRFElement Is Nothing Then
                            'the element isn't on the form it said it was; report an error
                            vRtn = AddToArray(vRtn, .sEformId, "Element could not be located on eForm [" & .sElementID & "]")
                            .bProcessed = True
                        
                        Else
                            'load this response
                            Set .oResponse = .oEFormInstance.Responses.ResponseByElement(oCRFElement, .nRepeat)
                            
                            'if response is nothing and a RQG question then attempt to create response
                            If ((.oResponse Is Nothing) And (.nRepeat > 1)) Then
                                'get group instance, create new row, get response again
                                Set oQGI = .oEFormInstance.QGroupInstanceById(oCRFElement.OwnerQGroup.QGroupID)
                                'for www, we force the createnewrow(). this solves the problem that
                                'occurs if one of the middle rows of the rqg was blanked - all lower
                                'rows would be lost
                                Call oQGI.CreateNewRow(True)
                                Set .oResponse = .oEFormInstance.Responses.ResponseByElement(oCRFElement, .nRepeat)
                            End If
                            
                            
                            If IsWWWFieldEnterable(oWWWField) Then
                                'process the response if it is enterable
                                Call ProcessQuestion(oUser, oWWWField, oCRFElement, sDecimalPoint, _
                                    sThousandSeparator, bOverruleWarning, vRtn)
                                    
                                If Not .bRejected Then
                                    'if the value wasnt rejected mark field as processed, otherwise
                                    'leave it as unprocessed so we can try to process it next time
                                    'around the loop
                                    .bProcessed = True
                                
                                    If (.bOK) And (.oResponse.LockStatus = eLockStatus.lsUnlocked) Then
                                        'we only save the comment here if the question has been processed
                                        'and marked as 'ok' and 'processed'. comments for non-enterable questions
                                        'will be saved in ProcessNotEnterableQuestions(). this is so we dont
                                        'save comments for questions whose values are rejected later
                                        Call SaveComment(oUser, .sAInfo, .oResponse)
                                    End If
                                End If
                            End If
                            
                        End If 'oCRFElement Is Nothing
                    End If 'oWWWField.oEFormInstance Is Nothing
                End If ' oWWWField.bProcessed = false
            End With
        Next


        'get count of fields that are not marked as 'processed' at end of loop
        nNotProcessedEndCount = RtnNotProcessedCount(colFields)
        
        
        If (nNotProcessedStartCount > nNotProcessedEndCount) Then
            'only do this if some questions were processed in the last loop, otherwise
            'we'll be falling out of the loop at the next statement. note: we do this
            'even if there are no more items to process (nNotProcessedEndCount = 0)
            'refresh skips and derivations ready for the next loop through responses:
            'doing this may enable responses that were previously disabled by skip
            'conditions that are dependant on fields that changed in the previous loop
            If Not bEReadOnly Then
                Call oEFI.RefreshSkipsAndDerivations(Revalidation, "")
            End If
            If Not oVEFI Is Nothing And Not bVReadOnly Then
                Call oVEFI.RefreshSkipsAndDerivations(Revalidation, "")
            End If
        End If

        
    'loop until (we dont have less fields to do at the end of a loop than we had at
    'the beginning of the loop) or (there are no more questions to process)
    Loop Until (nNotProcessedStartCount <= nNotProcessedEndCount) Or (nNotProcessedEndCount = 0)
    
    
    'final check for outstanding updates:
    'user may have changed a derived field status from/to okwarning. this will not have been
    'handled in the main process loop as the question will never have become enterable
    Call ProcessNotEnterableQuestions(oUser, colFields, bOverruleWarning, vRtn)
    
    
    'ic 01/03/2004 moved revalidation code from ProcessQuestion() so that
    'the eform is revalidated as a whole. this is to fix the bug where a question
    'with a validation is saved before the question that the validation involves
    'e.g. q1 validation = 'q1 < q2'. however, this situation should now be caught
    'by the client-side javascript fnRevalidatePage() function anyway.
    Call RevalidateEFI(colFields, bEReadOnly, bVReadOnly, bUReject, bVReject)
    
    'refresh skips and derivations if any values were rejected during revalidation
    If bUReject Then Call oEFI.RefreshSkipsAndDerivations(Revalidation, "")
    If bVReject Then Call oVEFI.RefreshSkipsAndDerivations(Revalidation, "")
    
    'check for rejected values that may have occurred
    Call ProcessRejectedQuestions(colFields, vRtn)
    
    Exit Sub
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.ProcessQuestions"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub ProcessRejectedQuestions(ByRef colFields As Collection, ByRef vRtn As Variant)
'--------------------------------------------------------------------------------------------------
'   ic 21/04/2004
'   check for any questions which are marked as rejected. questions are marked as rejected in the
'   ProcessQuestion() function. because of the unpredictable order in which questions
'   can be saved, a question may be rejected on the first pass, but accepted on a subsequent pass,
'   it is in the subsequent pass that the rejected marker will be removed. any questions marked
'   as rejected after all passes are genuine rejections which should be flagged to the user.
'   rejections should be handled on the client so this is really only precautionary.
'   NOTE, comments are ignored for rejected value questions
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sErrMessage As String
Dim nValid As Integer
Dim bChanged As Boolean
Dim oField As WWWField

    On Error GoTo CatchAllError

    For Each oField In colFields
        With oField
            If (.bRejected) Then
                nValid = .oResponse.ValidateValue(.sServerLocaleValue, sErrMessage, bChanged, .dblTimestamp)
            
                If nValid = Status.InvalidData Then
                    'if the value is invalid, log the error for return to the asp
                    vRtn = AddToArray(vRtn, .oResponse.Element.Name, sErrMessage & " [" & .sValue & "][" & .sServerLocaleValue & "]")
                        
                End If
            End If
        End With
    Next
    Exit Sub
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.ProcessRejectedQuestions"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub ProcessNotEnterableQuestions(ByRef oUser As MACROUser, ByRef colFields As Collection, _
    ByVal bOverruleWarning As Boolean, ByRef vRtn As Variant)
'--------------------------------------------------------------------------------------------------
'   ic 04/03/2004
'   check for updates requested on not-enterable questions. this includes comments and status
'   changes, which are handled elsewhere for enterable questions
'   revisions
'   ic 29/06/2004 added error handling
'   ic 24/05/2005 issue 2566, ignore hidden fields
'--------------------------------------------------------------------------------------------------
Dim oWWWField As WWWField
Dim bRequiresRFC As Boolean

    On Error GoTo CatchAllError

    For Each oWWWField In colFields
        With oWWWField
            
            If (Not .bProcessed) And (Not .oResponse Is Nothing) Then
                'only handle fields that werent processed by ProcessQuestion() and are not locked
                If (Not IsWWWFieldEnterable(oWWWField)) And ((.oResponse.LockStatus = eLockStatus.lsUnlocked)) _
                And (Not .oResponse.Element.Hidden) Then
                
                    If bOverruleWarning Then
                        'handle warning->okwarning/okwarning->warning/change of okwarning reason
                        If (.oResponse.Status = Status.OKWarning) Then
                            If (.sRFO = "") And (.bRFOPresent) Then
                                'changing from okwarning->warning
                                bRequiresRFC = .oResponse.RequiresStatusRFC(Status.Warning)
                                
                                If ((bRequiresRFC And (.sRFC <> "")) Or (Not bRequiresRFC)) Then
                                    Call .oResponse.SetStatus(Status.Warning, .sRFC, oUser.UserName, oUser.UserNameFull)
                                Else
                                    vRtn = AddToArray(vRtn, .oResponse.Element.Name, "No reason for change supplied for OKWarning to Warning status change")
                                End If
                            ElseIf (.sRFO <> "") Then
                                'possibly changing okwarning reason
                                If (.sRFO <> .oResponse.OverruleReason) Then
                                    bRequiresRFC = .oResponse.RequiresOverruleRFC(.sRFO)
                                    
                                    If ((bRequiresRFC And (.sRFC <> "")) Or (Not bRequiresRFC)) Then
                                        Call .oResponse.SetOverruleReason(.sRFO, .sRFC)
                                    Else
                                        vRtn = AddToArray(vRtn, .oResponse.Element.Name, "No reason for change supplied for change to overrule reason")
                                    End If
                                End If
                            End If
                        
                        ElseIf (.oResponse.Status = Status.Warning) And (.sRFO <> "") Then
                            'changing from warning->okwarning
                            bRequiresRFC = .oResponse.RequiresStatusRFC(Status.OKWarning)
                            
                            If ((bRequiresRFC And (.sRFC <> "")) Or (Not bRequiresRFC)) Then
                                Call .oResponse.SetOverruleReason(.sRFO, .sRFC)
                            Else
                                vRtn = AddToArray(vRtn, .oResponse.Element.Name, "No reason for change supplied for Warning to OKWarning status change")
                            End If
                            
                        End If
                    End If
                    
                    'save comment, if any
                    Call SaveComment(oUser, .sAInfo, .oResponse)
                End If
            End If
        End With
    Next
    
    Exit Sub
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.ProcessNotEnterableQuestions"
End Sub

'--------------------------------------------------------------------------------------------------
Private Function IsWWWFieldEnterable(ByRef oWWWField As WWWField) As Boolean
'--------------------------------------------------------------------------------------------------
' Check if WWWField is an enterable one
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
    On Error GoTo CatchAllError

    ' set to false initially
    IsWWWFieldEnterable = False
    If Not oWWWField.oResponse Is Nothing Then
        If oWWWField.oResponse.Enterable Then
            IsWWWFieldEnterable = True
        End If
    End If
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.IsWWWFieldEnterable"
End Function

'--------------------------------------------------------------------------------------------------
Private Sub ProcessQuestion(ByRef oUser As MACROUser, ByRef oWWWField As WWWField, _
    ByRef oCRFElement As eFormElementRO, ByVal sDecimalPoint As String, ByVal sThousandSeparator As String, _
    ByVal bOverruleWarning As Boolean, ByRef vRtn As Variant)
'--------------------------------------------------------------------------------------------------
' Process the question response. it is assumed the user has the neccessary permissions
' revisions
' OUTSTANDING ISSUES: implementation of RequiresStatusRFC,RequiresOverruleRFC,RequiresValueRFC,
'                     RequiresLabRFC
' ic 04/03/2004 revised handling of revalidation
' ic 20/04/2004 dont alert missing rfcs as this signifies a bug in the jsve
' ic 21/04/2004 revised handling for rejected values
'   ic 29/06/2004 added error handling
' ic 28/02/2007 issue 2855 - clinical coding release
'--------------------------------------------------------------------------------------------------
Dim sErrMessage As String
Dim sSDBCon As String
Dim sUserName As String
Dim sUserNameFull As String
Dim bChanged As Boolean
Dim bOKWarnToWarn As Boolean
Dim nValid As Integer
    
    
    On Error GoTo CatchAllError

    'reset response instance variables
    bOKWarnToWarn = False
    sUserName = oUser.UserName
    sUserNameFull = oUser.UserNameFull

    
    With oWWWField
        If .oResponse Is Nothing Then
            'the user is not allowed to answer this question
            If .sValue = "" Then
                'they didn't answer it; next response
            Else
                'report an error
                vRtn = AddToArray(vRtn, oCRFElement.Name, "Question is not available for data entry, but a response was entered. [" & .sValue & "]")
            End If
            
        Else
            'ic 16/09/2003 convert from browser locale to server locale
            .sServerLocaleValue = ServeriseValue(.sValue, .oResponse.Element.DataType, _
                sDecimalPoint, sThousandSeparator)
            
            'validate the passed value
            nValid = .oResponse.ValidateValue(.sServerLocaleValue, sErrMessage, bChanged, .dblTimestamp)
                
                
            If nValid = Status.InvalidData Then
                'if the value is invalid, mark it as rejected. we will try to validate it again
                'in the next pass. invalid values should be handled by the client-side validation,
                'however, values saved by other users on other eforms while this eform was being
                'edited may have caused the rejection, or the rejection may be a false rejection
                'caused by another question on this eform not having been saved before this one
                'for example, a 'units of measurement' question relating to this value may not
                'have been processed yet, so this value may be outside expected boundaries for
                'the current unit of measurement.
                .bRejected = True
               
            Else
                .bRejected = False
            
                If ((nValid = Status.Warning) Or (nValid = Status.OKWarning)) And (bOverruleWarning) Then
                    'if returned as status OKWarning but no overrule then status has actually changed to
                    'warning. if has been reset there will be a setting for "o" but it will be ""
                    If (nValid = Status.OKWarning) And (.sRFO = "") And (.bRFOPresent) Then
                        bChanged = True
                        nValid = Status.Warning
                        bOKWarnToWarn = True
                    End If
                End If
                
                If bChanged Then
                    'if the value is different from the stored value
                    If oCRFElement.Authorisation <> "" Then
                        'if authorisation is required for this response
                        If (sSDBCon = "") Then sSDBCon = GetSecurityConx
                        ' DPH 19/02/2003 Get full name for Authorising User
                        If Not IsAuthorisationOK(sSDBCon, .sAuthUserName, .sAuthPassword, oCRFElement.Authorisation, oUser.DatabaseCode, sUserNameFull) Then
                            .bOK = False
                            vRtn = AddToArray(vRtn, oCRFElement.Name, "Password authorisation rejected")
                        Else
                            sUserName = .sAuthUserName
                        End If
                    End If
    
                    'if no problems occurred while checking for accompanying data
                    If .bOK And (.oResponse.LockStatus = eLockStatus.lsUnlocked) Then
                        'confirm new value
                        .oResponse.ConfirmValue .sRFO, .sRFC, sUserName, sUserNameFull
                        
                        'ic 28/02/2007 issue 2855 - clinical coding release - check for clinical coding questions
                        If (oCRFElement.DataType = eDataType.Thesaurus) Then
                            If .oResponse.CodingStatus <> eCodingStatus.csDoNotCode Then
                                'reset coding status as response value has changed
                                Call .oResponse.SetCodingStatus(CInt(eCodingStatus.csNotCoded), oUser.UserName, oUser.UserNameFull)
                            End If
                        End If
                        
                        If bOKWarnToWarn Then
                            'reset status to warn
                            Call .oResponse.SetStatus(Status.Warning, .sRFC, sUserName, sUserNameFull)
                        End If
                    End If
                Else
                    'save the overrule reason (if one is supplied) even if response hasnt changed
                    If (.sRFO <> "") And (.oResponse.LockStatus = eLockStatus.lsUnlocked) Then
                        ' NCJ 16 Aug 02 - WE NEED AN RFC HERE!!!
                        Call .oResponse.SetOverruleReason(.sRFO, .sRFC)
                    End If
                End If
                
                'if no save was required, or the save was ok
                If (.bOK) Then
                    'only explicity set status for empty fields (unobtainable/missing)
                    If (.oResponse.Value = "") Then
                     
                        'toggle between missing/unobtainable status
                        If (.oResponse.Status = Status.Unobtainable) _
                        And (.sUnobtainable = "false") Then
                            'if response was previously unobtainable
                            'ic 26/08/2003 dont set optional fields to missing, set them straight to ok
                            If .oResponse.Element.IsOptional Then
                                Call .oResponse.SetStatus(Status.Success, .sRFC, sUserName, sUserNameFull)
                            Else
                                'dph 17/06/2003 added rfc
                                ' NCJ 16 Aug 02 - WE NEED AN RFC HERE!!!
                                Call .oResponse.SetStatus(Status.Missing, .sRFC, sUserName, sUserNameFull)
                            End If
                            
                        ElseIf ((.oResponse.Status = Status.Missing) _
                        Or (.oResponse.Status = Status.Requested)) _
                        And (.sUnobtainable = "true") Then
                            'if response was previously missing/requested
                            'dph 17/06/2003 added rfc
                            ' NCJ 16 Aug 02 - WE NEED AN RFC HERE!!!
                            Call .oResponse.SetStatus(Status.Unobtainable, .sRFC, sUserName, sUserNameFull)
                        End If
                    End If
    
                End If '.bOK
            End If 'nValid = Status.InvalidData
        End If '.oResponse Is Nothing
    End With
    
    Exit Sub
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.ProcessQuestion"
End Sub

'--------------------------------------------------------------------------------------------------
Private Function EformHasLabQuestions(ByRef oEform As eFormRO) As Boolean
'--------------------------------------------------------------------------------------------------
'   ic 03/04/2003
'   does passed eform have any lab questions on it
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oElement As eFormElementRO
Dim bRtn As Boolean

    On Error GoTo CatchAllError

    bRtn = False
    For Each oElement In oEform.EFormElements
        If oElement.DataType = eDataType.LabTest Then
            bRtn = True
            Exit For
        End If
    Next
    EformHasLabQuestions = bRtn
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.EformHasLabQuestions"
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetEformLabChoice(ByRef oUser As MACROUser, ByVal sSite As String, _
                         Optional ByVal enInterface As eInterface = iwww) As String
'--------------------------------------------------------------------------------------------------
'   ic 03/04/2003
'   function builds and returns a string representing a lab choice dialog
'--------------------------------------------------------------------------------------------------
' REVISIONS
' DPH 27/05/2003 - OK/Cancel buttons
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim vJSComm() As String
Dim vLab As Variant
Dim nLoop As Integer

    On Error GoTo CatchAllError

    ReDim vJSComm(0)
    
    vLab = oUser.DataLists.GetLabsForSiteList(sSite)
    
    Call AddStringToVarArr(vJSComm, "<html>" _
        & "<head><title>Laboratories</title>")
       
    If Not IsNull(vLab) Then
        If (enInterface = iwww) Then
            Call AddStringToVarArr(vJSComm, "<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>" _
                & "<script language='javascript' src='../script/Dialog.js'></script>" _
                & "<script language='javascript'>" _
                & "function retval(sVal){" _
                & "this.returnValue=sVal;" _
                & "this.close();}" _
                & "function keyPress(){" _
                & "if (window.event.keyCode==13){" _
                & "fnReturn(selLab.value);}" _
                & "}" _
                & "window.document.onkeypress=keyPress;" _
                & "function fnPageLoaded(){" _
                & "selLab.selectedIndex=0;}" _
                & "</script>")
        End If
        
        Call AddStringToVarArr(vJSComm, "</head>" _
            & "<body onload='fnPageLoaded();'>")
    
        Call AddStringToVarArr(vJSComm, "<table align='center' width='95%' border='0'>" & vbCrLf _
                        & "<tr height='10'><td></td></tr>" & vbCrLf _
                        & "<tr height='15' class='clsLabelText'>" _
                          & "<td colspan='2'>" _
                            & "" _
                          & "</td>" _
                          & "<td><a style='cursor:hand;' onclick='javascript:fnReturn(selLab.value);'><u>OK</u></a></td>" _
                          & "<td><a style='cursor:hand;' onclick='javascript:fnReturn(" & Chr(34) & Chr(34) & ");'><u>Cancel</u></a></td>" _
                        & "</tr>" _
                        & "<tr height='15'><td></td></tr>" & vbCrLf)
        
        Call AddStringToVarArr(vJSComm, "<tr height='5'>" _
                          & "<td></td>" _
                        & "</tr>" _
                        & "<tr height='30'>" _
                          & "<td colspan='4' class='clsLabelText'>" _
                            & "Please select a laboratory for the laboratory test questions on this eForm" _
                          & "</td>" _
                        & "</tr><tr height='5'><td>&nbsp;</td></tr>" _
                        & "<tr>" _
                          & "<td width='100'></td>" _
                          & "<td><select style='width:180px;' name='selLab' class='clsSelectList' size='3'>")
                   
        For nLoop = LBound(vLab, 2) To UBound(vLab, 2)
            Call AddStringToVarArr(vJSComm, "<option value='" & vLab(eLabCols.lcCode, nLoop) & "'>" & vLab(eLabCols.lcCode, nLoop) & " (" & vLab(eLabCols.lcDescription, nLoop) & ")</option>")
        Next
        
        Call AddStringToVarArr(vJSComm, "</select></td>" _
                      & "</tr>" _
                      & "</table>" _
                      & "</body>" _
                      & "</html>")
    Else
        Call AddStringToVarArr(vJSComm, "<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>" _
            & "</head><body class='clsLabelText'>No labs are set up for this site</body></html>")
    End If
    
    GetEformLabChoice = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.GetEformLabChoice"
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetLoadBody(ByVal sJsFn As String, _
                   Optional ByVal vAlerts As Variant) As String
'--------------------------------------------------------------------------------------------------
'   ic 13/01/2003
'   function returns an html body that calls a passed js function in parent page
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sHTML As String
Dim nLoop As Integer

    On Error GoTo CatchAllError

    sHTML = "<body onload='window.parent.sWinState=" & Chr(34) & Chr(34) & ";" _
                        & "fnHideLoader();" _
                        & "window.parent.frames[0].navigate(" & Chr(34) & "blank.htm" & Chr(34) & ");" _
                        & "window.parent.frames[1].navigate(" & Chr(34) & "blank.htm" & Chr(34) & ");" _
                        & "window.parent.window.parent." & sJsFn & ";"
    'any additional alerts to display
    If Not IsMissing(vAlerts) Then
        If Not IsEmpty(vAlerts) Then
            For nLoop = LBound(vAlerts, 2) To UBound(vAlerts, 2)
                sHTML = sHTML & "alert(" & Chr(34) & ReplaceWithJSChars(vAlerts(1, nLoop)) & Chr(34) & ");"
            Next
        End If
    End If
    sHTML = sHTML & "'></body>"
    
    GetLoadBody = sHTML
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.GetLoadBody"
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetDecisionBody(ByVal sEformNonRepeat As String, _
                                ByVal sEformRepeat As String, _
                                ByVal sEformNextVisit As String, _
                                ByVal sEformSame As String, _
                                ByVal sEformName As String, _
                                ByVal sDatabase As String, _
                                ByVal sStudyCode As String, _
                                ByVal sSiteCode As String, _
                                ByVal sSubjectId As String, _
                       Optional ByVal vAlerts As Variant, _
                       Optional ByVal enInterface As eInterface = iwww) As String
'--------------------------------------------------------------------------------------------------
'   ic 01/10/02
'   builds and returns an html string allowing a user to decide between 2 eforms
'   revisions
'   ic 29/01/2003 fixed 'last eform in visit' code to work with new eform.js
'   ic 19/02/2003 amendments to handle 'next visit' functionality
'   ic 26/02/2004 change to 'next' variable format: id~jsfn~saveflag
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim sHTML As String
Dim nLoop As Integer

    On Error GoTo CatchAllError

    sHTML = sHTML & "<body onload='fnPageLoaded();'>" & vbCrLf _
                  & "<form name='FormDE' method='post' action='Eform.asp?fltDb=" & sDatabase & "&fltSt=" & sStudyCode & "&fltSi=" & sSiteCode & "&fltSj=" & sSubjectId & "'>" & vbCrLf _
                  & "<input type='hidden' name='next'>" & vbCrLf _
                  & "</form>" & vbCrLf _
                  & "<div id='wholeDiv'></div>"
                  
    sHTML = sHTML & "<script language='javascript'>" & vbCrLf _
                    & "function fnPageLoaded()" & vbCrLf _
                    & "{" & vbCrLf _
                      & "o2=window.document.FormDE;" & vbCrLf
                       
    'any additional alerts to display
    If Not IsMissing(vAlerts) Then
        If Not IsEmpty(vAlerts) Then
            For nLoop = LBound(vAlerts, 2) To UBound(vAlerts, 2)
                sHTML = sHTML & "alert(" & Chr(34) & ReplaceWithJSChars(vAlerts(1, nLoop)) & Chr(34) & ");" & vbCrLf
            Next
        End If
    End If
                       
    If (sEformRepeat <> "") Then
        sHTML = sHTML & "if (confirm('Would you like to open a new " & sEformName & " form?'))" & vbCrLf _
                      & "{" & vbCrLf _
                        & "o2.next.value='e" & sEformRepeat & gsDELIMITER3 & gsDELIMITER3 & "0';" & vbCrLf _
                        & "o2.submit();" & vbCrLf _
                        & "return;" & vbCrLf _
                      & "}" & vbCrLf
    End If
    
    If (sEformNonRepeat <> "") Then
        sHTML = sHTML & "o2.next.value='e" & sEformNonRepeat & gsDELIMITER3 & gsDELIMITER3 & "0';" & vbCrLf _
                      & "o2.submit();" & vbCrLf
                      
    Else
        If (sEformNextVisit <> "") Then
            sHTML = sHTML & "if (confirm('This is the last eForm in this visit.\nWould you like to move to the next visit?'))" & vbCrLf _
                          & "{" & vbCrLf _
                            & "o2.next.value='e" & sEformNextVisit & gsDELIMITER3 & gsDELIMITER3 & "0';" & vbCrLf _
                            & "o2.submit();" & vbCrLf _
                            & "return;" & vbCrLf _
                          & "}" & vbCrLf
        End If
    
        sHTML = sHTML & "alert('This is the last eform in the current visit');" & vbCrLf _
                      & "o2.next.value='e" & sEformSame & gsDELIMITER3 & gsDELIMITER3 & "0';" & vbCrLf _
                      & "o2.submit();" & vbCrLf
    End If

    sHTML = sHTML & "}" & vbCrLf _
              & "</script>" & vbCrLf _
              & "</body>"
                    
    GetDecisionBody = sHTML
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.GetDecisionBody"
End Function

'---------------------------------------------------------------------
Private Function DBLock(sCon As String, ByRef sLockErrMsg As String, ByRef sEFILockToken As String, sUser As String, _
                                lStudyId As Long, sSite As String, lSubjectId As Long, _
                                lEFITaskId As Long) As String
'---------------------------------------------------------------------
' Lock an efi.
' Returns if token if lock successful or empty string if not
' reurns the reason if the lock fails in sLockErrMsg
'---------------------------------------------------------------------
Dim sLockDetails As String
Dim sToken As String

    On Error GoTo Errorlabel
    
    'TA 04.07.2001: use new locking
    sToken = MACROLOCKBS30.LockEFormInstance(sCon, sUser, lStudyId, sSite, lSubjectId, lEFITaskId)
    Select Case sToken
    Case MACROLOCKBS30.DBLocked.dblStudy
        sLockDetails = MACROLOCKBS30.LockDetailsStudy(sCon, lStudyId)
        If sLockDetails = "" Then
            sLockErrMsg = "This study is currently being edited by another user."
        Else
            sLockErrMsg = "This study is currently being edited by " & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblSubject
        sLockDetails = MACROLOCKBS30.LockDetailsSubject(sCon, lStudyId, sSite, lSubjectId)
        If sLockDetails = "" Then
            sLockErrMsg = "This subject is currently being edited by another user."
        Else
            sLockErrMsg = "This subject is currently being edited by " & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblEFormInstance
        sLockDetails = MACROLOCKBS30.LockDetailseFormInstance(sCon, lStudyId, sSite, lSubjectId, lEFITaskId)
        If sLockDetails = "" Then
            sLockErrMsg = "This eForm is currently being used by another user."
             sToken = ""
        Else
            If sEFILockToken = Split(sLockDetails, "|")(2) Then
                'we already have the lock - still return it though
                sToken = sEFILockToken
            Else
                sLockErrMsg = "This eForm is currently being used by " & Split(sLockDetails, "|")(0) & "."
                sToken = ""
            End If
        End If
        
    Case Else
        'hurrah, we have a lock

    End Select
    DBLock = sToken
    Exit Function
    
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.LockForSave"
    Exit Function

End Function

'---------------------------------------------------------------------
Private Sub DBUnlock(sCon As String, sToken As String, _
                            lStudyId As Long, sSite As String, lSubjectId As Long, _
                                lEFITaskId As Long)
'---------------------------------------------------------------------
' Unlock the efi
'---------------------------------------------------------------------

On Error GoTo Errorlabel
    'TA 04.07.2001: use new locking model
    If sToken <> "" Then
        'if no gsStudyToken then UnlockSubject is being called without a corresponding LockSubject being called first
        MACROLOCKBS30.UnlockEFormInstance sCon, sToken, lStudyId, sSite, lSubjectId, lEFITaskId
        'always set this to empty string for same reason as above
        sToken = ""
    End If
    Exit Sub
    
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "StudySubject.UnlockForSave"

End Sub

'--------------------------------------------------------------------------------------------------
Public Function GetEformBody(ByRef oUser As MACROUser, _
                             ByRef oSubject As StudySubject, _
                             ByVal sDatabase As String, _
                             ByVal sSiteCode As String, _
                             ByVal lCRFPageTaskId As Long, _
                             ByRef sEFILockToken As String, _
                             ByRef sVILockToken As String, _
                             ByRef bEFIUnavailable As Boolean, _
                             ByVal sDecimalPoint As String, _
                             ByVal sThousandSeparator As String, _
                    Optional ByVal enInterface As eInterface = iwww, _
                    Optional ByVal vAlerts As Variant, _
                    Optional ByVal vErrors As Variant, _
                    Optional ByVal bAutoNext As Boolean = False, _
                    Optional ByVal bTrace As Boolean = False) As String
'--------------------------------------------------------------------------------------------------
'   ic 01/10/02
'   builds and returns a string representing an eform in html
'--------------------------------------------------------------------------------------------------
' REVISIONS
' DPH 16/10/2002 - Changed to use oLastFocusID
' DPH 14/11/2002 - Get eFormWidth
' ic  22/11/2002   added image path arguement to geteformelement()
' DPH 26/11/2002 - Get Highest Tab Index
' DPH 19/12/2002 - Show EITHER subject label OR (subjectid) not both
' DPH 23/12/2002 - Always use white for question background
' MLM 28/01/03: Make the form's header bar taller and include form and visit date prompts.
' MLM 24/02/03: Modified to include static content rather than generating it here.
' ic 04/04/2003 added lab display/choice
' MLM 23/04/03: Added bAutoNext parameter for performance testing.
' DPH 09/05/2003 - Execute function to place btnNext DIV
' DPH 12/05/2003 set eform width in main DIV
' ic 13/05/2003 set locale separators
' DPH 29/05/2003 - added status to fnInitialiseApplet call
' ic 10/06/2003 'next' field should always be first of group
' ic 19/06/2003 added registration
' ic 26/06/2003 commented SetLocalNumberFormats(), bug 1032&1033
' DPH 01/09/2003 - Added Overrule Warnings permission, bug 1986
' ic 02/09/2003 disable db lock admin menu item when on eform, bug 1992
' ic 18/09/2003 added 'Initialising eForm' message, pause
' DPH 06/10/2003 initialisation of images on main frame not eForm
' DPH 02/12/2003  - Change order form is drawn (visit eform 1st if exists - then normal eform)
' ic 27/02/2004 added 'no responses on eform' flag to fnInitialiseApplet() call
' ic 15/03/2004 added initialise focus code
' ic 14/04/2004 removed EformIsEmpty() jsve initialisation - no longer needed
' ic 16/04/2004 removed 'overrule warnings' parameter from fnInitialiseApplet() js call
' ic 29/06/2004 added error handling
' ic 20/07/2004 amended form onunload
' ic 08/11/2004 bug 2395, added setsetting for SETTING_LAST_USED_STUDY
' ic 20/06/2005 issue 2593, changed the fnLoadVisitBorder() and fnLoadEformBorder() script call
' NCJ 25 May 06 - Issue 2740 - use CancelResponses instead of RemoveResponses (to ensure AREZZO is cleared)
' ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
' ic 28/02/2008 add code to add inactive categories that are the current saved value
'--------------------------------------------------------------------------------------------------
Dim oElement As eFormElementRO
Dim oEFI As EFormInstance
Dim oVisitEFI As EFormInstance
'MLM 28/01/03:
Dim oEFormDateElement As eFormElementRO
Dim oVisitDateElement As eFormElementRO

Dim nLoop As Integer
Dim nPageLength As Integer
Dim eLoadOK As eLoadResponsesResult
Dim sLockErrMsg As String
Dim leFormWidth As Long
Dim sImagePath As String
Dim lHighestTab As Long
Dim sLabelOrId As String
Dim bVReadOnly As Boolean
Dim bHasLabQuestions As Boolean
Dim sLabCode As String
Dim vLab As Variant
Dim vJSComm() As String
Dim lRaised As Long
Dim lResponded As Long
Dim lPlanned As Long
Dim lHeaderHeight As Long
Dim sLocalFormat As String


    On Error GoTo CatchAllError
    ReDim vJSComm(0)
    
    'ic 28/02/2008 initialise global string
    gsAddInactiveCategory = ""
    
    Set oEFI = oSubject.eFIByTaskId(lCRFPageTaskId)
    'set user numeric separator
    'ic 26/06/2003 this is no longer set when loading an eform - numbers are sent in standard format
    'Call oSubject.StudyDef.SetLocalNumberFormats(sDecimalPoint, sThousandSeparator)
    eLoadOK = oSubject.LoadResponses(oEFI, sLockErrMsg, sEFILockToken, sVILockToken)
    Set oVisitEFI = oEFI.VisitInstance.VisitEFormInstance
    'MLM 28/01/03:
    Set oEFormDateElement = oEFI.eForm.EFormDateElement
    If Not oVisitEFI Is Nothing Then
        Set oVisitDateElement = oVisitEFI.eForm.EFormDateElement
    End If
    
    
    'ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
    If (oSubject.Arezzo Is Nothing) Then
        'load without arezzo - DR user - attempt to get an eform lock. this is normally done
        'in loadresponses, but because in this special case we dont have an arezzo, the subject
        'is readonly and doesnt create a lock. however we want to allow the user the ability to
        'save mimessages so we need to manually create one
        Dim sLockToken As String
        
        sLockToken = DBLock(oUser.CurrentDBConString, sLockErrMsg, sEFILockToken, oUser.UserName, _
            oSubject.StudyDef.StudyId, oSubject.Site, oSubject.PersonId, oEFI.EFormTaskId)
        sEFILockToken = sLockToken
        
        If Not oVisitEFI Is Nothing Then
            sLockToken = DBLock(oUser.CurrentDBConString, sLockErrMsg, sVILockToken, oUser.UserName, _
                oSubject.StudyDef.StudyId, oSubject.Site, oSubject.PersonId, oVisitEFI.EFormTaskId)
            sVILockToken = sLockToken
        End If
    End If
    
    
    'handle visit eforms - may prevent opening eform
    bEFIUnavailable = False
    If oEFI.Status <> eStatus.Requested Then
    ElseIf oEFI.VisitInstance.VisitDate <> 0 Then
    Else
        If Not oVisitEFI Is Nothing Then
            If Not oVisitEFI.eForm.EFormDateElement Is Nothing Then
                If oVisitEFI.ReadOnly And (oVisitEFI.eForm.EFormDateElement.DerivationExpr = "") Then
                    ' another user has the visit eform open but hasnt yet saved the visit date
                    bEFIUnavailable = True
'                    Call oSubject.RemoveResponses(oEFI, True)
                    Call oSubject.CancelResponses(oEFI, True)     ' NCJ 25 May 06 - 2740 - Ensure everything cleared
                    sEFILockToken = ""
                    Set oEFI = Nothing
                    Set oVisitEFI = Nothing
                    Exit Function
                End If
            End If
        End If
    End If
    If oVisitEFI Is Nothing Then
        bVReadOnly = True
    Else
        bVReadOnly = oVisitEFI.ReadOnly
    End If
    
    sImagePath = "../sites/" & sDatabase & "/" & oSubject.StudyId & "/"
    lHighestTab = 0
    'ic 03/04/2003
    ' DPH 28/05/2003 - Add in header height
    bHasLabQuestions = EformHasLabQuestions(oEFI.eForm)
    If (bHasLabQuestions) Then
        If (oEFI.LabCode <> "") Then
            sLabCode = oEFI.LabCode
        Else
            'if only one lab for this site, select it
            vLab = oUser.DataLists.GetLabsForSiteList(sSiteCode)
            If Not IsNull(vLab) Then
                If UBound(vLab, 2) = 0 Then
                    sLabCode = vLab(eLabCols.lcCode, 0)
                End If
            End If
        End If
        lHeaderHeight = 1100
    Else
        lHeaderHeight = 750
    End If
    
    Call RtnMIMsgStatusCount(oUser, lRaised, lResponded, lPlanned)
    ' DPH 16/05/2003 - Use grey as main body background
    ' main DIV gets the eform background colour
    Call AddStringToVarArr(vJSComm, "<body bgcolor='#EEEEEE' onload='javascript:fnPageLoaded();'>")
    'Call AddStringToVarArr(vJSComm, "<body bgcolor='" & RtnHTMLCol(oEFI.eForm.BackgroundColour) & "' onload='javascript:fnPageLoaded();' onUnload='javascript:fnCloseDepWindow();'>")
   
    'css (study default) only text not matching will be defined on the page
    Call AddStringToVarArr(vJSComm, "<style type='text/css'>" & vbCrLf _
                    & "div{" _
                      & "position:absolute;" _
                      & "width:auto;" _
                      & "height:auto;" _
                      & "font-family:" & oSubject.StudyDef.FontName & ";" _
                      & "font-size:" & oSubject.StudyDef.FontSize & "pt;" _
                      & "color:" & RtnHTMLCol(oSubject.StudyDef.FontColour) & ";" _
                    & "} " & vbCrLf _
                    & "td{" _
                      & "font-family:" & oSubject.StudyDef.FontName & ";" _
                      & "font-size:" & oSubject.StudyDef.FontSize & "pt;" _
                      & "color:" & RtnHTMLCol(oSubject.StudyDef.FontColour) & ";" _
                    & "}" _
                  & "</style>" & vbCrLf)
              
    
    'ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
    'work out if the eform is read only. if there is no arezzo, but we have an eform lock, this
    'is a dr user - eform is writeable but will only ever have permissions to add mimessages
    Dim bERO As Boolean
    Dim bVRO As Boolean
    If ((oSubject.Arezzo Is Nothing) And sEFILockToken <> "") Then
        bERO = False
    Else
        bERO = oEFI.ReadOnly
    End If
    If ((oSubject.Arezzo Is Nothing) And sVILockToken <> "") Then
        bVRO = False
    Else
        bVRO = bVReadOnly
    End If
    
    
    ' DPH 23/12/2002 - Always use white for question background
    'ic 18/09/2003 added 'Initialising eForm' message, pause
    'pageloaded function
    Call AddStringToVarArr(vJSComm, "<script language='javascript'>" & vbCrLf _
                  & "function fnPageLoaded()" & vbCrLf _
                    & "{" & vbCrLf _
                      & "o1=window;" & vbCrLf _
                      & "o2=window.document.FormDE;" & vbCrLf _
                      & "o3=window.document.all.tags('div');" & vbCrLf _
                      & "o4=window.parent.frames[2];" & vbCrLf _
                      & "s1='window.document.FormDE';" & vbCrLf _
                      & "fnShowLoader('Initialising eForm');" & vbCrLf _
                      & "window.setTimeout(" & Chr(34) & "fnPageLoaded2()" & Chr(34) & ",10);}" & vbCrLf _
                      & "function fnPageLoaded2(){" & vbCrLf _
                      & "o1.fnInitialiseApplet(" & "o2," _
                                                 & "'#FFFFFF'," _
                                                 & "null," _
                                                 & "o4," _
                                                 & "''," _
                                                 & RtnJSBoolean(bERO) & "," _
                                                 & RtnJSBoolean(bVRO) & "," _
                                                 & "'" & sLabCode & "'" & "," _
                                                 & RtnJSBoolean((oEFI.Status = eStatus.Requested)) & ");" & vbCrLf)
    ' DPH 06/10/2003 initialisation of images on main frame not eForm
    '                                             & IIf((oEFI.Status = eStatus.Requested), "1", "0") & ");" '& vbCrLf _
    '                  & "initialiseimages();" & vbCrLf)

'MLM 24/02/03:
'                      & RtnRadioCreateFuncs(oEFI.eForm) & vbCrLf _
'                      & RtnEformCategoryTxtFunc(oEFI.eForm) & vbCrLf)

    Call AddStringToVarArr(vJSComm, "DrawForm0();")
    If Not oVisitEFI Is Nothing Then
        Call AddStringToVarArr(vJSComm, "DrawForm1();")
    End If
        
    'eform permissions
    ' DPH 01/09/2003 - Added Overrule Warnings permission
    'ic 02/09/2003 use new CanChangeData function
    Call AddStringToVarArr(vJSComm, "fnSetEformPermissions('" & oUser.UserName & "'," & RtnJSBoolean(CanChangeData(oUser, oSubject.Site)) & "," _
                                             & RtnJSBoolean(oUser.CheckPermission(gsFnAddIComment)) & "," _
                                             & RtnJSBoolean(oUser.CheckPermission(gsFnViewIComments)) & "," _
                                             & RtnJSBoolean(oUser.CheckPermission(gsFnViewAuditTrail)) & "," _
                                             & RtnJSBoolean(oUser.CheckPermission(gsFnCreateDiscrepancy)) & "," _
                                             & RtnJSBoolean(oUser.CheckPermission(gsFnCreateSDV)) & "," _
                                             & RtnJSBoolean(oUser.CheckPermission(gsFnViewDiscrepancies)) & "," _
                                             & RtnJSBoolean(oUser.CheckPermission(gsFnViewSDV)) & "," _
                                             & RtnJSBoolean(oUser.CheckPermission(gsFnOverruleWarnings)) & "," _
                                             & RtnJSBoolean(oUser.CheckPermission(gsFnMonitorDataReviewData)) & ");" & vbCrLf)
  
                                            
    Call AddStringToVarArr(vJSComm, RtnJSVEInit(oUser, oSubject, sDatabase, oEFI, oVisitEFI))

    'ic 28/02/2008 add inactive categories
    If (gsAddInactiveCategory <> "") Then
        Call AddStringToVarArr(vJSComm, gsAddInactiveCategory)
    End If

    'errors encountered during save
    If Not IsMissing(vErrors) Then
        If Not IsEmpty(vErrors) Then
            Call AddStringToVarArr(vJSComm, "alert('MACRO encountered problems while saving. Some responses could not be saved." _
                          & "\nUnsaved responses are listed below\n\n")
            
            For nLoop = LBound(vErrors, 2) To UBound(vErrors, 2)
                Call AddStringToVarArr(vJSComm, ReplaceWithJSChars(vErrors(0, nLoop) & " - " & vErrors(1, nLoop)) & "\n")
            Next
        
            Call AddStringToVarArr(vJSComm, "');" & vbCrLf)
        End If
    End If
        
    'eform&visit delimited list for loading eform borders
    Call AddStringToVarArr(vJSComm, "var sEforms='" & RtnDelimitedEformList(oSubject.ScheduleGrid, oEFI.VisitInstance.VisitTaskId) & "';" & vbCrLf)
    Call AddStringToVarArr(vJSComm, "var sVisits='" & RtnDelimitedVisitList(oSubject.ScheduleGrid) & "';" & vbCrLf)
    
    'DPH 19/12/2002 - subject label or subjectid
    If oSubject.Label <> "" Then
        sLabelOrId = oSubject.Label
    Else
        sLabelOrId = "(" & oSubject.PersonId & ")"
    End If
    Call AddStringToVarArr(vJSComm, "var sLabel='" & oSubject.StudyCode & "/" & oSubject.Site & "/" & sLabelOrId & "';" & vbCrLf)


    Call AddStringToVarArr(vJSComm, "fnLoadVisitBorder(sLabel,sVisits,'" & oEFI.VisitInstance.VisitTaskId & "');" & vbCrLf _
                  & "fnLoadEformBorder(sEforms,'" & oEFI.EFormTaskId & "');" & vbCrLf _
                  & "fnHideLoader();" & vbCrLf)
                  
    Call AddStringToVarArr(vJSComm, "window.parent.window.parent.fnSTLC('" & gsVIEW_RAISED_DISCREPANCIES_MENUID & "','" & CStr(lRaised) & "',0);" & vbCrLf _
                                      & "window.parent.window.parent.fnSTLC('" & gsVIEW_RESPONDED_DISCREPANCIES_MENUID & "','" & CStr(lResponded) & "',0);" & vbCrLf _
                                      & "window.parent.window.parent.fnSTLC('" & gsVIEW_PLANNED_SDV_MARKS_MENUID & "','" & CStr(lPlanned) & "',1);" & vbCrLf)
    
                  
    'ic 19/08/2003 enable/disable registration menu item
    'ic 02/09/2003 use new CanChangeData function
    If (CanChangeData(oUser, oSubject.Site) And oUser.CheckPermission(gsFnRegisterSubject) And ShouldEnableRegistrationMenu(oSubject)) Then
        Call AddStringToVarArr(vJSComm, "window.parent.window.parent.fnEnableRegister(" & RtnJSBoolean(True) & ");" & vbCrLf)
    Else
        Call AddStringToVarArr(vJSComm, "window.parent.window.parent.fnEnableRegister(" & RtnJSBoolean(False) & ");" & vbCrLf)
    End If
            
    'ic 02/09/2003 disable db lock admin menu item when on eform
    Call AddStringToVarArr(vJSComm, "window.parent.window.parent.fnEnableDLA(" & RtnJSBoolean(False) & ");" & vbCrLf)
            
    'any additional alerts to display
    If Not IsMissing(vAlerts) Then
        If Not IsEmpty(vAlerts) Then
            For nLoop = LBound(vAlerts, 2) To UBound(vAlerts, 2)
                Call AddStringToVarArr(vJSComm, "alert('" & ReplaceWithJSChars(vAlerts(1, nLoop)) & "');" & vbCrLf)
            Next
        End If
    End If
             
    Call AddStringToVarArr(vJSComm, "window.parent.sWinState='5" & gsDELIMITER2 & oSubject.StudyId & gsDELIMITER2 & oSubject.Site & gsDELIMITER2 & oSubject.PersonId & gsDELIMITER2 & oEFI.EFormTaskId & "';" & vbCrLf)
    
    ' DPH 27/02/2003 - Revalidate page setup
    'ic 09/11/2004 moved this call to javascript fnApplyRules()
    'Call AddStringToVarArr(vJSComm, "o1.fnRevalidatePage();" & vbCrLf)
    
    'MLM 23/04/03: After loading the page, automatically go to next eForm if measuring performance
    If bAutoNext Then
        Call AddStringToVarArr(vJSComm, "fnSave('m5');" & vbCrLf)
    End If
    
    ' DPH 09/05/2003 - Execute function to place btnNext DIV
    Call AddStringToVarArr(vJSComm, "o1.fnSetBtnNextPos();" & vbCrLf)
    'ic 22/08/2003 eform loaded successfully flag
    Call AddStringToVarArr(vJSComm, "bEformLoadCheck=true;" & vbCrLf)
    
    If (CanChangeData(oUser, oSubject.Site)) Then
        'ic 15/03/2004 generate an initial focus function call
        Call AddStringToVarArr(vJSComm, RtnInitialFocus(oEFI, oVisitEFI) & vbCrLf)
    End If
    
    Call AddStringToVarArr(vJSComm, "}" & vbCrLf) 'end of fnPageLoaded
    
    'MLM 10/10/02: Insert functions to check whether the form can be saved
    '   (saving is prevented by a blank enterable form or visit date)
    Call AddStringToVarArr(vJSComm, GetDateOkFunc(eEFormUse.User, oEFI.eForm) & vbCrLf)
    If oVisitEFI Is Nothing Then
        Call AddStringToVarArr(vJSComm, GetDateOkFunc(eEFormUse.VisitEForm, Nothing) & vbCrLf)
    Else
        Call AddStringToVarArr(vJSComm, GetDateOkFunc(eEFormUse.VisitEForm, oVisitEFI.eForm) & vbCrLf)
    End If
    
    'close script block
    Call AddStringToVarArr(vJSComm, "</script>")

    leFormWidth = oEFI.eForm.eFormWidth
    If leFormWidth = NULL_LONG Then
        leFormWidth = 8515 ' Portrait width constant value
    End If
    
    'start of display html
    ' DPH 12/05/2003 set eform width in main DIV
    Call AddStringToVarArr(vJSComm, "<div style='position:relative;visibility:visible;width:" & (leFormWidth / nIeXSCALE) & ";' id='wholeDiv'>")
    Call AddStringToVarArr(vJSComm, "<form name='FormDE' method='post' action='Eform.asp?fltDb=" & sDatabase & "&fltSt=" & oSubject.StudyDef.StudyId & "&fltSi=" & sSiteCode & "&fltEf=" & oEFI.eForm.EFormId & "&fltId=" & oEFI.EFormTaskId & "&fltSj=" & oSubject.PersonId & "'>" & vbCrLf)
                  
    'MLM 14/10/02: Display a grey rectange at the top of the form to display the visit and form names and cycle numbers.
    'MLM 28/01/03: Make the rectangle taller, and add prompts for form and visit dates if required.
    ' DPH 12/05/2003 - makes eForm header width of page
    ' DPH 28/05/2003 use lHeaderHeight calculated earlier
    Call AddStringToVarArr(vJSComm, "<table id=""eFormHeaderTable"" width=" & leFormWidth / nIeXSCALE & " height=" & (lHeaderHeight / nIeYSCALE) & _
        " style=""color:#AAAAAA; background-color:#EEEEEE; border-color:#AAAAAA;border-width:1px; border-style:solid;""><tr>" & vbCrLf & _
        "<td width=50% style=""color:#AAAAAA;font-family:Verdana,helvetica,arial; font-size:8pt;"">Visit: " & oEFI.VisitInstance.Visit.Name)
    If oEFI.VisitInstance.CycleNo > 1 Then
        Call AddStringToVarArr(vJSComm, " [" & oEFI.VisitInstance.CycleNo & "]")
    End If
    ' DPH 12/05/2003 - eForm Read only tooltip (if required)
    Call AddStringToVarArr(vJSComm, "</td>" & vbCrLf & "<td width=50% style=""color:#AAAAAA;font-family:Verdana,helvetica,arial; font-size:8pt;""")
    If oEFI.ReadOnly Then
        Call AddStringToVarArr(vJSComm, " title=""This eForm is read only.  " & oEFI.ReadOnlyReason & """")
    End If
    Call AddStringToVarArr(vJSComm, ">eForm: " & oEFI.eForm.Name)
    If oEFI.CycleNo > 1 Then
        Call AddStringToVarArr(vJSComm, " [" & oEFI.CycleNo & "]")
    End If
    Call AddStringToVarArr(vJSComm, "</td></tr>" & vbCrLf & "<tr><td width=50% style=""color:#AAAAAA;font-family:Verdana,helvetica,arial; font-size:8pt;"">")
    'MLM 26/06/03: Only show the fixed form and visit date prompts if the study does not define its own
    If oVisitDateElement Is Nothing Then
        Call AddStringToVarArr(vJSComm, "&nbsp;")
    Else
        Call AddStringToVarArr(vJSComm, IIf(oVisitDateElement.Caption = "", "Visit date:", "&nbsp;"))
    End If
    Call AddStringToVarArr(vJSComm, "</td>" & vbCrLf & "<td width=50% style=""color:#AAAAAA;font-family:Verdana,helvetica,arial; font-size:8pt;"">")
    If oEFormDateElement Is Nothing Then
        Call AddStringToVarArr(vJSComm, "&nbsp;")
    Else
        Call AddStringToVarArr(vJSComm, IIf(oEFormDateElement.Caption = "", "eForm date:", "&nbsp;"))
    End If
    
    'ic 03/04/2003 add lab display
    If bHasLabQuestions Then
        Call AddStringToVarArr(vJSComm, "</td></tr>" & vbCrLf & "<tr><td id='tdlab1' width=50% style=""color:#AAAAAA;font-family:Verdana,helvetica,arial; font-size:8pt;"">&nbsp;")
        Call AddStringToVarArr(vJSComm, "</td>" & vbCrLf & "<td id='tdlab2' width=50% style=""color:#AAAAAA;font-family:Verdana,helvetica,arial; font-size:8pt;"">")
    End If
    
    Call AddStringToVarArr(vJSComm, "</td></tr></table>" & vbCrLf)
    
    ' DPH 02/12/2003  - Change order form is drawn (visit eform 1st if exists - then normal eform)
    'MLM 10/10/02: If there's a visit eform, load its radio buttons and category text too
    If Not oVisitEFI Is Nothing Then
        Call AddStringToVarArr(vJSComm, "<script language=javascript src='" & sImagePath & "form" & oVisitEFI.eForm.EFormId & "_1.js'></script>" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "<div id=Form1></div>")
    End If
    Call AddStringToVarArr(vJSComm, "<script language=javascript src='" & sImagePath & "form" & oEFI.eForm.EFormId & "_0.js'></script>" & vbCrLf)
    Call AddStringToVarArr(vJSComm, "<div id=Form0 style=""top:0;left:0;width:" & (leFormWidth / nIeXSCALE) & ";""></div>")

    'ic 10/06/2003 'next' field should always be first of group
    'ic 04/04/2003 add labcode hidden field
    ' DPH 26/06/2003 Added local date format hidden field
    'DPH 25/06/2003 Add local format date to eForm
    If oUser.UserSettings.GetSetting(SETTING_LOCAL_FORMAT, False) Then
        sLocalFormat = oUser.UserSettings.GetSetting(SETTING_LOCAL_DATE_FORMAT, "")
    Else
        sLocalFormat = ""
    End If
    
    'ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
    Call AddStringToVarArr(vJSComm, "<input type='hidden' name='next'>" & vbCrLf _
                  & "<input type='hidden' name='labcode' value='" & sLabCode & "'>" & vbCrLf _
                  & "<input type='hidden' name='readonly' value='" & RtnJSBoolean(bVRO) & gsDELIMITER1 & RtnJSBoolean(bERO) & "'>" & vbCrLf _
                  & "<input type='hidden' name='localdate' value='" & sLocalFormat & "'>" & vbCrLf _
                  & "</form>" _
                  & "</div>" _
                  & "</body>")
                      
    Call oUser.UserSettings.SetSetting(SETTING_LAST_USED_STUDY, oEFI.eForm.Study.StudyId)
    Call oUser.UserSettings.SetSetting(SETTING_LAST_USED_EFORM, oEFI.GetAllSubjectsKey)
    ' NCJ 26 May 06 - Issue 2740 - Call CancelResponses to ensure AREZZO patient state is cleared
'    Call oSubject.RemoveResponses(oEFI, False)

    
    'ic 01/08/2007 issue 2939, only load an arezzo for users with change data permissions
    'only cancel responses if we have an arezzo
    If (Not oSubject.Arezzo Is Nothing) Then
        Call oSubject.CancelResponses(oEFI, False)      ' NCJ 28 Jun 06 - Include False to preserve eForm lock
    End If
    
    
    Set oElement = Nothing
    Set oEFI = Nothing
    GetEformBody = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.GetEformBody"
End Function

'-------------------------------------------------------------------------------------------'
Private Function RtnInitialFocus(ByRef oEFI As EFormInstance, ByRef oVEFI As EFormInstance) As String
'-------------------------------------------------------------------------------------------'
'   ic 12/03/2004
'   function returns a javascript call to set initial focus when an eform has loaded
'   ic 29/06/2004 added error handling
'-------------------------------------------------------------------------------------------'
Dim vJS As String
Dim oElement As eFormElementRO
Dim oResponse As Response

    'not serious if initial focus isnt set
    On Error GoTo IgnoreError

    'loop through all user eform fields, finds the first enterable one
    If (Not oEFI Is Nothing) Then
        For Each oElement In oEFI.eForm.EFormElements
            Set oResponse = oEFI.Responses.ResponseByElement(oElement)
            
            If Not oResponse Is Nothing Then
                If oResponse.Enterable Then
                    vJS = "o1.fnInitialFocus(" & Chr(34) & oResponse.Element.WebId & Chr(34) & "," & (oResponse.RepeatNumber - 1) & ");"
                    Exit For
                End If
            End If
        Next
    End If

    'loop through all the visit eform fields, finds the first enterable one
    If (Not oVEFI Is Nothing) Then
        For Each oElement In oVEFI.eForm.EFormElements
            Set oResponse = oVEFI.Responses.ResponseByElement(oElement)
            
            If Not oResponse Is Nothing Then
                If oResponse.Enterable Then
                    'creates a focus call only if the visit/eform date is empty, or
                    'no fields were found in the user eform loop
                    If (oResponse.Value = "") Or (vJS = "") Then
                        vJS = "o1.fnInitialFocus(" & Chr(34) & oResponse.Element.WebId & Chr(34) & "," & (oResponse.RepeatNumber - 1) & ");"
                    End If
                    Exit For
                End If
            End If
        Next
    End If
    

    If Not oElement Is Nothing Then Set oElement = Nothing
    If Not oResponse Is Nothing Then Set oResponse = Nothing
    
    RtnInitialFocus = vJS
IgnoreError:
End Function

'-------------------------------------------------------------------------------------------'
Public Function RtnDelimitedVisitList(ByRef oSchedule As ScheduleGrid) As String
'-------------------------------------------------------------------------------------------'
'   ic 07/10/2002
'   function returns a string containing a list of visits in the passed schedule
'   list is delimited and contains groups of visit code,visitid,visit name parameters
'   revisions
'   ic 01/04/2003   added flag for 'display code or name'
'   ic 29/06/2004 added error handling
'-------------------------------------------------------------------------------------------'
Dim nRow As Integer
Dim nCol As Integer
Dim sRtn As String
Dim sCycle As String
Dim bShowVCode As Boolean
Dim oSchedCell As GridCell
Dim vJSComm() As String


    On Error GoTo CatchAllError
    ReDim vJSComm(0)
    
    'show code header rather than full name
    bShowVCode = mbShowCode

    For nCol = 1 To oSchedule.ColMax
        sCycle = ""
        Set oSchedCell = oSchedule.Cells(0, nCol)
        With oSchedCell
            If Not .VisitInst Is Nothing Then
                If .VisitInst.CycleNo > 1 Then sCycle = "[" & .VisitInst.CycleNo & "]"
                If (bShowVCode) Then
                    Call AddStringToVarArr(vJSComm, ".Visit.Code & sCycle & gsDELIMITER2" _
                                & .VisitInst.VisitTaskId & gsDELIMITER2 _
                                & ReplaceWithJSChars(.Visit.Name) & sCycle & gsDELIMITER1)
                Else
                    Call AddStringToVarArr(vJSComm, ReplaceWithJSChars(.Visit.Name) & sCycle & gsDELIMITER2 _
                                & .VisitInst.VisitTaskId & gsDELIMITER2 _
                                & .Visit.Code & sCycle & gsDELIMITER1)
                End If
            End If
        End With
    Next
    Set oSchedCell = Nothing
    sRtn = Join(vJSComm, "")
    If (Len(sRtn) > 0) Then sRtn = Left(sRtn, (Len(sRtn) - 1))
    RtnDelimitedVisitList = sRtn
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnDelimitedVisitList"
End Function


'-------------------------------------------------------------------------------------------'
Public Function RtnDelimitedEformList(ByRef oSchedule As ScheduleGrid, _
                                       ByVal lVisitTaskId As Long) As String
'-------------------------------------------------------------------------------------------'
'   ic 07/10/2002
'   function returns a string containing a list of eforms in the passed visit instance
'   list is delimited and contains groups of eform code,eform pagetaskid,eform name
'   parameters
'   revisions
'   ic 01/04/2003   added flag for 'display code or name'
'   ic 29/06/2004 added error handling
'-------------------------------------------------------------------------------------------'
Dim nRow As Integer
Dim nCol As Integer
Dim sRtn As String
Dim sCycle As String
Dim bShowECode As Boolean
Dim vJSComm() As String

    On Error GoTo CatchAllError
    ReDim vJSComm(0)

    'show code header rather than full name
    bShowECode = mbShowCode
    
    For nCol = 1 To oSchedule.ColMax
        If Not oSchedule.Cells(0, nCol).VisitInst Is Nothing Then
            If oSchedule.Cells(0, nCol).VisitInst.VisitTaskId = lVisitTaskId Then
                For nRow = 1 To oSchedule.RowMax
                    sCycle = ""
                    If oSchedule.Cells(nRow, nCol).CellType <> Inactive Then
                        If Not oSchedule.Cells(nRow, nCol).eFormInst Is Nothing Then
                            If oSchedule.Cells(nRow, nCol).eFormInst.CycleNo > 1 Then sCycle = "[" & oSchedule.Cells(nRow, nCol).eFormInst.CycleNo & "]"
                            If (bShowECode) Then
                                Call AddStringToVarArr(vJSComm, oSchedule.Cells(nRow, 0).eForm.Code & sCycle & gsDELIMITER2 _
                                            & oSchedule.Cells(nRow, nCol).eFormInst.EFormTaskId & gsDELIMITER2 _
                                            & ReplaceWithJSChars(oSchedule.Cells(nRow, 0).eForm.Name) & sCycle & gsDELIMITER1)
                            Else
                                Call AddStringToVarArr(vJSComm, ReplaceWithJSChars(oSchedule.Cells(nRow, 0).eForm.Name) & sCycle & gsDELIMITER2 _
                                            & oSchedule.Cells(nRow, nCol).eFormInst.EFormTaskId & gsDELIMITER2 _
                                            & oSchedule.Cells(nRow, 0).eForm.Code & sCycle & gsDELIMITER1)
                            End If
                        End If
                    End If
                Next
            End If
        End If
    Next
    sRtn = Join(vJSComm, "")
    If (Len(sRtn) > 0) Then sRtn = Left(sRtn, (Len(sRtn) - 1))
    RtnDelimitedEformList = sRtn
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnDelimitedEformList"
End Function

'-------------------------------------------------------------------------------------------'
Private Function RtnCaption(ByVal oCRFElement As eFormElementRO, _
                            ByVal bDefaultFont As Boolean, _
                            ByVal bDisplayNumbers As Boolean, _
                            ByVal lDefFontCol As Long, _
                            Optional ByVal bRQGHeader As Boolean = False) As String
'-------------------------------------------------------------------------------------------'
'   ic 18/09/2003
'   function accepts an eform element object
'   function returns the html required to display the passed elements caption
'   Revisions:
'   ic 29/06/2004 added error handling
'-------------------------------------------------------------------------------------------'
Dim vJSComm() As String
    
    On Error GoTo CatchAllError
    ReDim vJSComm(0)
    
    'no outer div for rqgs
    If Not bRQGHeader Then
        Call AddStringToVarArr(vJSComm, "<div id='" & oCRFElement.WebId & "_" & "CapDiv'" & Chr(34) & " " _
                    & "style=" & Chr(34) _
                    & "top:" & CLng((oCRFElement.CaptionY + 375) / nIeYSCALE) & ";" _
                    & "left:" & CLng(oCRFElement.CaptionX / nIeXSCALE) & ";" _
                    & "z-index:1'>")
    End If
    
    Call AddStringToVarArr(vJSComm, "<font style='")
    Call AddStringToVarArr(vJSComm, "color:" & RtnHTMLCol(oCRFElement.CaptionFontColour) & ";")
        
    If Not bDefaultFont Then
        If (oCRFElement.CaptionFontSize <> 0) Then
            Call AddStringToVarArr(vJSComm, "font-size:" & oCRFElement.CaptionFontSize & "pt;")
        End If
        If (oCRFElement.CaptionFontName <> "") Then
            Call AddStringToVarArr(vJSComm, "font-family: " & oCRFElement.CaptionFontName & ";")
        End If
    End If
    Call AddStringToVarArr(vJSComm, "'>")
    
    If oCRFElement.CaptionFontBold Then Call AddStringToVarArr(vJSComm, "<b>")
    If oCRFElement.CaptionFontItalic Then Call AddStringToVarArr(vJSComm, "<i>")
    If bDisplayNumbers Then Call AddStringToVarArr(vJSComm, oCRFElement.ElementOrder & ".")
    Call AddStringToVarArr(vJSComm, ReplaceWithJSChars(ReplaceWithHTMLCodes(oCRFElement.Caption)))

    'unit
    If oCRFElement.Unit <> "" Then
        Call AddStringToVarArr(vJSComm, " (" & oCRFElement.Unit & ")")
    End If
    
    If oCRFElement.CaptionFontItalic Then Call AddStringToVarArr(vJSComm, "</i>")
    If oCRFElement.CaptionFontBold Then Call AddStringToVarArr(vJSComm, "</b>")

    Call AddStringToVarArr(vJSComm, "</font>")
    
    If Not bRQGHeader Then
        Call AddStringToVarArr(vJSComm, "</div>")
    End If
    
    RtnCaption = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnCaption"
End Function

'------------------------------------------------------------------------------------------------------
Private Function SetFontAttributes(ByVal oCRFElement As eFormElementRO, Optional ByVal bIncludeColour As Boolean) As String
'------------------------------------------------------------------------------------------------------
'Retrieve font attributes for a CRFElement
' MLM 14/10/02: If the element is a form or visit date, use fixed style
'   ic 29/06/2004 added error handling
'------------------------------------------------------------------------------------------------------
Dim sResult As String

    On Error GoTo CatchAllError
    
    If oCRFElement.ElementUse = eElementUse.EFormVisitDate Then
        sResult = "color:#AAAAAA;font-family:Verdana,helvetica,arial;font-size:8pt;"
    Else
        sResult = ""
        
        'don't get the colour for radio buttons
        If bIncludeColour Then
            If oCRFElement.FontColour <> 0 Then
                sResult = sResult & "color:#" & RtnHTMLCol(oCRFElement.FontColour) & ";"
            End If
        End If
        
        'get the font size
        If oCRFElement.FontSize <> 0 Then
            sResult = sResult & "font-size:" & oCRFElement.FontSize & "pt;"
        End If
        
        'check if font is bold
        If oCRFElement.FontBold Then
            sResult = sResult & "font-weight:" & "bold;"
        End If
        
        'check if font is italic
        If oCRFElement.FontItalic Then
            sResult = sResult & "font-style:italic;"
        End If
        
        'get font name
        If oCRFElement.FontName <> "" Then
            sResult = sResult & "font-family:" & oCRFElement.FontName & ";"
        End If
    End If
    
    SetFontAttributes = sResult
    
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.SetFontAttributes"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnJSVEInit(ByRef oUser As MACROUser, _
                            ByRef oSubject As StudySubject, _
                            ByVal sDatabase As String, _
                            ByRef oEFI As EFormInstance, _
                            ByRef oVisitEFI As EFormInstance) As Variant
'--------------------------------------------------------------------------------------------------
' REM 07/09/01
' Builds JavaScript function that can be called from the browser to initialise the fields on a
' specified eform inside the JSVE

' AMENDMENTS
' ic 13/09/01 added derivation,skip,validation functionality
' rem 15/10/01 bug fix 83: drop-down boxes pass Response code rather than description
' rem 06/11/01 bug fix 102: added fnApplyRules JS function to ensure all rules are applied
' rem 29/10/01 bug fix 102: added code to check for category question/text box datatype combination
'
' ic 25/09/2002 changed return type to variant array, changed locking model
' MLM 11/10/02: Added visit eform parameter
' DPH 18/10/2002 Added GetEFIInitialisationRQG to initialise RQGs in Javascript
' DPH 14/11/2002 Moved RQG resize code
' ic 16/01/2003 removed bReadOnly arg
' ic 23/01/2003 only generate skips/derivations/validations for writeable subjects
' DPH 20/02/2003 Add FullUserName & UserRole to an eForm using fnFm
' DPH 25/06/2003 Local Date Format
' ic 05/08/2003 we now always open forms writeable if possible, so check for changedata permission
'               before adding skips/derivs... and calling RevalidateEFI
' ic 21/08/2003 bug 1972, add eformid & eformtaskid for visit eform
' ic 02/03/2004 revalidation now done on client for eform open, client & server for eform close
' ic 29/06/2004 added error handling
' ic 27/02/2007 issue 2555, correctly identify and view cycling mimessages
'--------------------------------------------------------------------------------------------------

Dim oStudyDef As StudyDefRO
'Dim oSubject As StudySubject
'Dim oALM As Arezzo_DM
Dim oCRFElement As eFormElementRO
'Dim oEfi As EFormInstance
Dim oElement As eFormElementRO
Dim oResponse As Response
Dim oAzToJs As AzToJavaScript
'Dim oSCI As MACROSCI30.clsSubjectCacheInterface
Dim oValidation As Validation

Dim sSecDbCon As String
Dim sDbCon As String
Dim sJSString As String
Dim sResponseValue As String
Dim sResponseStatus As String
Dim sDataType As String
Dim sDataItemLength As String
Dim sErrMsg As String
Dim sElementEnabled As String
Dim vRtn(2) As Variant

Dim bViewComments As Boolean
'Dim bUseSCI As Boolean

'Dim vSubject As Variant
'Dim eLoadOK As eLoadResponsesResult
'Dim sToken As String

Dim sRQGInit() As String
Dim nI As Integer
Dim sRQGResize() As String
Dim vJSComm() As String
Dim lVEFTaskId As Long
Dim sReval() As String
' DPH 25/06/2003 Local Date Format
Dim sLocalFormat As String

    On Error GoTo CatchAllError
    ReDim vJSComm(0)

    bViewComments = oUser.CheckPermission(gsFnViewIComments)

    vRtn(0) = 0
    ReDim sRQGInit(0)
    sRQGInit(0) = ""
    ReDim sRQGResize(0)
    sRQGResize(0) = ""

    ReDim sReval(0)

    Call AddStringToVarArr(vJSComm, "o1.fnFm(" & Chr(34) & "sDatabase" & Chr(34) & "," _
                                              & Chr(34) & sDatabase & Chr(34) _
                                        & ");" & vbCrLf _
                  & "o1.fnFm(" & Chr(34) & "sStudyId" & Chr(34) & "," _
                                              & Chr(34) & oSubject.StudyId & Chr(34) _
                                        & ");" & vbCrLf _
                  & "o1.fnFm(" & Chr(34) & "sSite" & Chr(34) & "," _
                                              & Chr(34) & oSubject.Site & Chr(34) _
                                        & ");" & vbCrLf _
                  & "o1.fnFm(" & Chr(34) & "sSubject" & Chr(34) & "," _
                                              & Chr(34) & oSubject.PersonId & Chr(34) _
                                        & ");" & vbCrLf _
                  & "o1.fnFm(" & Chr(34) & "sVisitID" & Chr(34) & "," _
                                              & Chr(34) & oEFI.VisitInstance.VisitId & Chr(34) _
                                        & ");" & vbCrLf)
                                        
    ' ic 27/02/2007 issue 2555, correctly identify and view cycling mimessages
    Call AddStringToVarArr(vJSComm, "o1.fnFm(" & Chr(34) & "sVisitCycle" & Chr(34) & "," _
                                              & Chr(34) & oEFI.VisitInstance.CycleNo & Chr(34) _
                                        & ");" & vbCrLf _
                  & "o1.fnFm(" & Chr(34) & "sEformID" & Chr(34) & "," _
                                              & Chr(34) & oEFI.eForm.EFormId & Chr(34) _
                                        & ");" & vbCrLf _
                  & "o1.fnFm(" & Chr(34) & "sEformCycle" & Chr(34) & "," _
                                              & Chr(34) & oEFI.CycleNo & Chr(34) _
                                        & ");" & vbCrLf _
                  & "o1.fnFm(" & Chr(34) & "sEformPageTaskID" & Chr(34) & "," _
                                              & Chr(34) & oEFI.EFormTaskId & Chr(34) _
                                        & ");" & vbCrLf)
    
    If Not oVisitEFI Is Nothing Then
        ' ic 27/02/2007 issue 2555, correctly identify and view cycling mimessages
        'ic 21/08/2003 add eformid & eformtaskid for visit eform
        Call AddStringToVarArr(vJSComm, "o1.fnFm(" & Chr(34) & "sVEformID" & Chr(34) & "," _
                                              & Chr(34) & oVisitEFI.eForm.EFormId & Chr(34) _
                                        & ");" & vbCrLf _
                  & "o1.fnFm(" & Chr(34) & "sVEformCycle" & Chr(34) & "," _
                                              & Chr(34) & oVisitEFI.CycleNo & Chr(34) _
                                        & ");" & vbCrLf _
                  & "o1.fnFm(" & Chr(34) & "sVEformPageTaskID" & Chr(34) & "," _
                                              & Chr(34) & oVisitEFI.EFormTaskId & Chr(34) _
                                        & ");" & vbCrLf)
    End If
    
    
    'DPH 20/02/2003 Add FullUserName & UserRole to an eForm using fnFm
    Call AddStringToVarArr(vJSComm, "o1.fnFm(" & Chr(34) & "sFullUserName" & Chr(34) & "," _
                                        & Chr(34) & oUser.UserNameFull & Chr(34) & ");" & vbCrLf)
    Call AddStringToVarArr(vJSComm, "o1.fnFm(" & Chr(34) & "sUserRole" & Chr(34) & "," _
                                        & Chr(34) & oUser.UserRole & Chr(34) & ");" & vbCrLf)
    
    'DPH 25/06/2003 Add local format date to eForm
    If oUser.UserSettings.GetSetting(SETTING_LOCAL_FORMAT, False) Then
        sLocalFormat = oUser.UserSettings.GetSetting(SETTING_LOCAL_DATE_FORMAT, "")
    Else
        sLocalFormat = ""
    End If
    Call AddStringToVarArr(vJSComm, "o1.fnFm(" & Chr(34) & "sLocalDate" & Chr(34) & "," _
                                        & Chr(34) & sLocalFormat & Chr(34) & ");" & vbCrLf)
    
    'MLM 11/10/02: derive current values before initialising JSVE
    ' NCJ 27 Sept 07 - Bug 2935 - Don't pass in user name
'    oEFI.RefreshSkipsAndDerivations OpeningEForm, oUser.UserName
    oEFI.RefreshSkipsAndDerivations OpeningEForm, ""
    If Not oVisitEFI Is Nothing Then
        lVEFTaskId = oVisitEFI.EFormTaskId
    End If
    Call AddStringToVarArr(vJSComm, GetEFIInitialisationRQG(oEFI, sRQGInit, sRQGResize))
    ' for non visit eforms
    'ic 05/08/2003 only call function if user has changedata permission
    'ic 02/09/2003 use new CanChangeData function
'ic 02/03/2004 revalidation now done on client for eform open, client & server for eform close
'    If CanChangeData(oUser, oSubject.Site) Then Call ReValidateWholeEFI(oEFI, lVEFTaskId, oUser.UserName, sReval)
    Call AddStringToVarArr(vJSComm, GetEFIInitialisation(oUser, oSubject, oEFI, eEFormUse.User, sReval))
    For nI = 0 To UBound(sRQGInit)
        Call AddStringToVarArr(vJSComm, sRQGInit(nI))
    Next
    If Not oVisitEFI Is Nothing Then
        ' don't draw RQG's twice so remove RQG initialisation call
        ReDim sRQGInit(0)
        sRQGInit(0) = ""
        ' NCJ 27 Sept 07 - Bug 2935 - Don't pass in user name
'        oVisitEFI.RefreshSkipsAndDerivations OpeningEForm, oUser.UserName
        oVisitEFI.RefreshSkipsAndDerivations OpeningEForm, ""
        Call AddStringToVarArr(vJSComm, GetEFIInitialisationRQG(oVisitEFI, sRQGInit, sRQGResize))
        Call AddStringToVarArr(vJSComm, GetEFIInitialisation(oUser, oSubject, oVisitEFI, eEFormUse.VisitEForm, sReval))
        For nI = 0 To UBound(sRQGInit)
            Call AddStringToVarArr(vJSComm, sRQGInit(nI))
        Next
    End If

    'only generate skips/derivations/validations for writeable subjects
    'ic 05/08/2003 we now always open forms writeable if possible, so check for changedata permission
    'ic 02/09/2003 use new CanChangeData function
    If Not oSubject.ReadOnly And CanChangeData(oUser, oSubject.Site) Then
        'initialise arrezzo to js object passing current arezzo object
        Set oAzToJs = New AzToJavaScript
        
        'initialise aztojs object
        sErrMsg = oAzToJs.Initialise(oSubject.Arezzo.ALM)
        
        'if initialisation failed an error will be returned
        If sErrMsg = "" Then
            
                'ic 10/10/01 bug report 012
                'MLM 02/10/02: add the derivations, skips and validations for the main form...
                '(must be told about the visit form, so that it can find the values entered on it)
                Call AddStringToVarArr(vJSComm, RtnDerivationString(oEFI, oVisitEFI, oAzToJs))
                Call AddStringToVarArr(vJSComm, RtnSkipString(oEFI, oVisitEFI, oAzToJs))
                Call AddStringToVarArr(vJSComm, RtnValidationString(oEFI, oVisitEFI, oAzToJs))
                
                '... and the visit form is there is one.
                If Not oVisitEFI Is Nothing Then
                    Call AddStringToVarArr(vJSComm, RtnDerivationString(oVisitEFI, oEFI, oAzToJs))
                    Call AddStringToVarArr(vJSComm, RtnSkipString(oVisitEFI, oEFI, oAzToJs))
                    Call AddStringToVarArr(vJSComm, RtnValidationString(oVisitEFI, oEFI, oAzToJs))
                End If
        Else
            Call AddStringToVarArr(vJSComm, sErrMsg)
        End If
        Set oAzToJs = Nothing
    End If
    
    'rem 06/11/01 bug fix 102: Added fnApplyRules JS function call to ensure all Skips, Derivations and Validations occur
    Call AddStringToVarArr(vJSComm, "o1.fnApplyRules();" & vbCrLf)
    
    ' DPH 14/11/2002 - RQG Resizing after all icons drawn
    For nI = 0 To UBound(sRQGResize)
        Call AddStringToVarArr(vJSComm, sRQGResize(nI))
    Next
        
    'MLM 23/04/03: If user's connection is below 56 kbps, save time by compressing JavaScript
    If oUser.UserSettings.GetSetting(SETTING_BANDWIDTH, 100) < 56 Then
        RtnJSVEInit = JSCompress(Join(vJSComm, ""))
    Else
        RtnJSVEInit = Join(vJSComm, "")
    End If
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnJSVEInit"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnJSTemplate(ByRef oElement As eFormElementRO, ByVal bUserEform As Boolean, ByVal bRQG As Boolean, _
    ByVal sDataType As String, Optional ByVal sRQGId As String, Optional ByVal sFontColour As String, _
    Optional ByVal sFontSize As String, Optional ByVal sFontName As String, Optional bNumber As Boolean) As String
'--------------------------------------------------------------------------------------------------
'   ic 18/09/2003
'   function returns a javascript 'fnCT' template function call. created from code split from
'   the original GetEFIInitialisation() function
'   eg o1.fnCT("f_10001_10003",7,5,"",0,"","",1,10003,0,"",0,"Q1",1,1,0,null,0,"",0);
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim vJSComm() As String

    On Error GoTo CatchAllError
    ReDim vJSComm(0)
    
    
    With oElement
        Call AddStringToVarArr(vJSComm, "o1.fnCT(")
        
        '#1 question id
        Call AddStringToVarArr(vJSComm, Chr(34) & .WebId & Chr(34) & ",")
        
        '#2 type
        Call AddStringToVarArr(vJSComm, sDataType & ",")
        
        '#3 length
        Call AddStringToVarArr(vJSComm, IIf((.QuestionLength = NULL_INTEGER), "null", .QuestionLength) & ",")
        
        '#4 format
        Call AddStringToVarArr(vJSComm, Chr(34) & .Format & Chr(34) & ",")
        
        '#5 text case
        Call AddStringToVarArr(vJSComm, IIf((Not .Hidden), .TextCase, "0") & ",")
        
        '#6 font colour
        Call AddStringToVarArr(vJSComm, Chr(34) & IIf((Not .Hidden), RtnHTMLCol(.FontColour), "") & Chr(34) & ",")
        
        '#7 role of authoriser (if required)
        Call AddStringToVarArr(vJSComm, Chr(34) & .Authorisation & Chr(34) & ",")
        
        '#8 requires reason for change
        Call AddStringToVarArr(vJSComm, RtnJSBoolean(.RequiresRFC) & ",")
        
        '#9 questionid
        Call AddStringToVarArr(vJSComm, .QuestionId & ",")
        
        '#10 is rqg
        Call AddStringToVarArr(vJSComm, RtnJSBoolean(bRQG) & ",")
        
        '#11 rqg id
        Call AddStringToVarArr(vJSComm, Chr(34) & IIf(bRQG, sRQGId, "") & Chr(34) & ",")
        
        '#12 mandatory
        Call AddStringToVarArr(vJSComm, RtnJSBoolean(.IsMandatory) & ",")
        
        '#13 caption text
        If Not .Hidden Then
            If bRQG Then
                Call AddStringToVarArr(vJSComm, Chr(34) & RtnCaption(oElement, _
                    IsDefaultCaptionFont(oElement, sFontColour, sFontSize, sFontName), bNumber, sFontColour, _
                    True) & Chr(34) & ",")
            Else
                Call AddStringToVarArr(vJSComm, Chr(34) & ReplaceWithJSChars(oElement.Name) & Chr(34) & ",")
            End If
        Else
            Call AddStringToVarArr(vJSComm, Chr(34) & Chr(34) & ",")
        End If
        
        '#14 show icons
        Call AddStringToVarArr(vJSComm, RtnJSBoolean(.ShowStatusFlag And Not .Hidden) & ",")
        
        '#15 is user eform (not visit eform)
        Call AddStringToVarArr(vJSComm, RtnJSBoolean(bUserEform) & ",")
        
        '#16 is lab question
        Call AddStringToVarArr(vJSComm, RtnJSBoolean(.DataType = eDataType.LabTest) & ",")
        
        '#17 display length
        Call AddStringToVarArr(vJSComm, IIf((.DisplayLength = NULL_INTEGER), "null", .DisplayLength) & ",")
        
        '#18 is optional
        Call AddStringToVarArr(vJSComm, RtnJSBoolean(.IsOptional) & ",")
        
        '#19 font style
        Call AddStringToVarArr(vJSComm, Chr(34) & IIf((Not .Hidden), SetFontAttributes(oElement, False), "") & Chr(34) & ",")
        
        '#20 is hidden
        Call AddStringToVarArr(vJSComm, RtnJSBoolean(.Hidden))
        
        Call AddStringToVarArr(vJSComm, ");" & vbCrLf)
    End With
    
    RtnJSTemplate = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnJSTemplate"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnJSInstance(ByRef oResponse As Response, ByRef oElement As eFormElementRO, _
    ByVal sDataType As String, ByVal bRQG As Boolean, ByVal bCanChange As Boolean, _
    ByVal bViewComments As Boolean, ByVal nRepeat As Integer, ByRef sReval() As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 18/09/2003
'   function returns a javascript 'fnCI' instance function call. created from code split from
'   the original GetEFIInitialisation() function
'   eg o1.fnCI("f_10001_10003",0,"2",1,0,o2['f_10001_10003'],o3['f_10001_10003_CapDiv'],
'       "img_f_10001_10003",0,0,0,0,0,"","imgs_f_10001_10003","imgn_f_10001_10003",1,
'       "imgc_f_10001_10003","c|0|0|","","rde","",0,"");
'   revisions
'   ic 02/03/2004 revalidation flag changed
'   ic 29/06/2004 added error handling
'   ic 01/11/2005 added clinical coding screening
'   ic 15/06/2005 issue 2734 - derivations with line breaks cause an error in javascript
'   ic 28/02/2007 issue 2855 - clinical coding release
'   ic 28/02/2008 issue 2996 - add code to add inactive categories that are the current saved value
'--------------------------------------------------------------------------------------------------
Dim vJSComm() As String
Dim bNoResponse As Boolean
Dim bEnabled As Boolean
Dim sValue As String
Dim sComments As String
Dim sUserNameFull As String

    On Error GoTo CatchAllError
    ReDim vJSComm(0)
    
    bNoResponse = (oResponse Is Nothing)
    If bNoResponse Then
        sValue = ""
        bEnabled = False
        sComments = "c" & gsDELIMITER2 & "0" & gsDELIMITER2 & "0"
        sUserNameFull = ""
    Else
        ' DPH 06/11/2003 - check to make sure hidden category fields get code passed to them
        If (sDataType = WWWDataType.Category) Or (sDataType = WWWDataType.PopUp) Or (oElement.DataType = WWWDataType.Category) Or (oElement.DataType = WWWDataType.PopUp) Then
            sValue = oResponse.SavedValueCode
        Else
            sValue = oResponse.SavedValue
        End If
        'standardise and escape any double quotes
        sValue = Replace(StandardiseValue(sValue, oElement.DataType), Chr(34), "\" & Chr(34))
        'ic 28/02/2007 issue 2855 - clinical coding release - enable thesaurus questions for data input
        'ic 01/11/2005 thesaurus questions should be disabled
        ' IC/DPH 06/11/2003 - Hidden fields should not be disabled by default
        bEnabled = (Not oResponse.EFormInstance.ReadOnly) And (bCanChange) _
            And (Not oElement.ControlType = ControlType.Attachment)
        'bEnabled = (Not oResponse.EFormInstance.ReadOnly) And (bCanChange) _
        '    And (Not oElement.ControlType = ControlType.Attachment) And (Not oElement.Hidden)
        sComments = "c" & gsDELIMITER2 & "0" & gsDELIMITER2 & Len(oResponse.Comments) & gsDELIMITER2
        If (bViewComments) Then
            sComments = sComments & ReplaceWithJSChars(ReplaceLfWithDelimiter(oResponse.Comments, gsDELIMITER2))
        End If
        If (oResponse.Status = eStatus.Requested) Then
            sUserNameFull = ""
        Else
            sUserNameFull = ReplaceWithJSChars(oResponse.UserNameFull)
        End If
    End If
    
    
    With oElement
        Call AddStringToVarArr(vJSComm, "o1.fnCI(")
        
        '#1 sFieldID
        Call AddStringToVarArr(vJSComm, Chr(34) & .WebId & Chr(34) & ",")
        
        '#2 nRepeatNo
        Call AddStringToVarArr(vJSComm, CStr(nRepeat) & ",")
        
        '#3 value
        'ic 15/06/2005 issue 2734 - derivations with line breaks cause an error in javascript
        Call AddStringToVarArr(vJSComm, Chr(34) & Replace(sValue, vbCrLf, " ") & Chr(34) & ",")
        
        '#4 bEnabled, can element ever be enabled
        Call AddStringToVarArr(vJSComm, RtnJSBoolean(bEnabled) & ",")
        
        '#5 nStatus. use Response.SavedStatus so can detect eForm changes
        If (bNoResponse) Then
            Call AddStringToVarArr(vJSComm, "-10,")
        Else
            Call AddStringToVarArr(vJSComm, oResponse.SavedStatus & ",")
        End If
        
        '#6 oHandle, element handle
        If ((.ControlType = ControlType.OptionButtons) _
        Or (.ControlType = ControlType.PushButtons)) And Not .Hidden Then
            'if VISIBLE radio question then pass div handle
            Call AddStringToVarArr(vJSComm, "o2.all['" & .WebId & "_inpDiv" & "'],")
        Else
            'otherwise pass element handle
            Call AddStringToVarArr(vJSComm, "o2['" & .WebId & "'],")
        End If
        
        '#7 oCapHandle, caption handle
        Call AddStringToVarArr(vJSComm, IIf((Not .Hidden), "o3['" & .WebId & "_CapDiv" & "']", "null") & ",")
        
        '#8 sImageName, name of status image
        Call AddStringToVarArr(vJSComm, Chr(34) & IIf((Not .Hidden), "img_" & .WebId, "") & Chr(34) & ",")
        
        '#9 nLockStatus
        If (bNoResponse) Then
            Call AddStringToVarArr(vJSComm, eLockStatus.lsUnlocked & ",")
        Else
            Call AddStringToVarArr(vJSComm, oResponse.LockStatus & ",")
        End If
        
        '#10 nDiscrepancyStatus
        If (bNoResponse) Then
            Call AddStringToVarArr(vJSComm, "0,")
        Else
            Call AddStringToVarArr(vJSComm, oResponse.DiscrepancyStatus & ",")
        End If
        
        '#11 nSDVStatus
        If (bNoResponse) Then
            Call AddStringToVarArr(vJSComm, "0,")
        Else
            Call AddStringToVarArr(vJSComm, oResponse.SDVStatus & ",")
        End If
        
        '#12 bNote, note present
        If (bNoResponse) Then
            Call AddStringToVarArr(vJSComm, RtnJSBoolean(False) & ",")
        Else
            Call AddStringToVarArr(vJSComm, RtnJSBoolean(CBool(oResponse.NoteStatus)) & ",")
        End If
        
        '#13 bComment, comment present
        If (bNoResponse) Then
            Call AddStringToVarArr(vJSComm, RtnJSBoolean(False) & ",")
        Else
            Call AddStringToVarArr(vJSComm, RtnJSBoolean(CBool(oResponse.Comments <> "")) & ",")
        End If
        
        '#14 oRadio, if radio control, handle to radio control (inc RQG). otherwise empty string
        If ((.ControlType = ControlType.OptionButtons) Or (.ControlType = ControlType.PushButtons)) _
        And (Not bRQG) And (Not .Hidden) Then
            Call AddStringToVarArr(vJSComm, "aRadioList['" & oElement.WebId & "'].olRepeat[" & nRepeat & "],")
        Else
            Call AddStringToVarArr(vJSComm, Chr(34) & Chr(34) & ",")
        End If
        
        '#15 sImageSName, name of sdv icon
        Call AddStringToVarArr(vJSComm, Chr(34) & IIf((Not .Hidden), "imgs_" & oElement.WebId, "") & Chr(34) & ",")
        
        '#16 sSelectNoteImage, name of note image (select lists only)
        If (oElement.ControlType = ControlType.PopUp) And (Not oElement.Hidden) Then
             Call AddStringToVarArr(vJSComm, Chr(34) & "imgn_" & oElement.WebId & Chr(34) & ",")
        Else
            Call AddStringToVarArr(vJSComm, Chr(34) & Chr(34) & ",")
        End If
        
        '#17 nChanges, change count
        If (bNoResponse) Then
            Call AddStringToVarArr(vJSComm, "0,")
        Else
            Call AddStringToVarArr(vJSComm, oResponse.ChangeCount & ",")
        End If
        
        '#18 sImageCName, name of change count icon
        Call AddStringToVarArr(vJSComm, Chr(34) & IIf((Not .Hidden), "imgc_" & oElement.WebId, "") & Chr(34) & ",")
        
        '#19 sComments
        Call AddStringToVarArr(vJSComm, Chr(34) & sComments & Chr(34) & ",")
        
        '#20 sRFC, reason for change
        If (bNoResponse) Then
            Call AddStringToVarArr(vJSComm, Chr(34) & Chr(34) & ",")
        Else
            Call AddStringToVarArr(vJSComm, Chr(34) & ReplaceWithJSChars(oResponse.ReasonForChange) & Chr(34) & ",")
        End If
        
        '#21 sUserFull, user full name, only add username if response isnt new
        Call AddStringToVarArr(vJSComm, Chr(34) & sUserNameFull & Chr(34) & ",")
        
        '#22 sNRCTC
        If (bNoResponse) Then
            Call AddStringToVarArr(vJSComm, Chr(34) & Chr(34) & ",")
        Else
            Call AddStringToVarArr(vJSComm, Chr(34) & RtnNRCTCText(oResponse.Status, oResponse.NRStatus, oResponse.CTCGrade) & Chr(34) & ",")
        End If
        
        '#23 bReVal
        If (bNoResponse) Then
            Call AddStringToVarArr(vJSComm, RtnJSBoolean(False) & ",")
        Else
            Call AddStringToVarArr(vJSComm, RtnJSBoolean(oResponse.Revalidatable(False)) & ",")
        End If
        
        '#24 sRFO
        If (bNoResponse) Then
            Call AddStringToVarArr(vJSComm, Chr(34) & Chr(34) & ",")
        Else
            'ic 11/03/2004 check for responses that have warning status, but have an overrule reason
            'we dont want to send the overrule reason to the client in this case
            If (oResponse.Status = Status.OKWarning) Then
                Call AddStringToVarArr(vJSComm, Chr(34) & ReplaceWithJSChars(oResponse.OverruleReason) & Chr(34) & ",")
            Else
                Call AddStringToVarArr(vJSComm, Chr(34) & Chr(34) & ",")
            End If
        End If
        
        '#25 validation message
        If (bNoResponse) Then
            Call AddStringToVarArr(vJSComm, Chr(34) & Chr(34))
        Else
            Call AddStringToVarArr(vJSComm, Chr(34) & ReplaceWithJSChars(oResponse.ValidationMessage) & Chr(34))
        End If
        
        Call AddStringToVarArr(vJSComm, ");" & vbCrLf)
        
        'ic 28/02/2008 issue 2996 - add a category value to this category question if the response value is now an inactive option
        Call AddInactiveCategoryValue(sDataType, sValue, oElement, nRepeat)
    End With
    
    RtnJSInstance = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnJSInstance"
End Function

'--------------------------------------------------------------------------------------------------
Private Sub AddInactiveCategoryValue(ByVal sDataType As String, sValue As String, _
    ByRef oElement As eFormElementRO, nRepeat As Integer)
'--------------------------------------------------------------------------------------------------
'   ic 28/02/2008 issue 2996 - add code to add inactive categories that are the current saved value
'--------------------------------------------------------------------------------------------------
    Dim oCategory As CategoryItem

    If (sValue = "") Then Exit Sub
    
    If (sDataType = WWWDataType.Category) Or (sDataType = WWWDataType.PopUp) Or (oElement.DataType = WWWDataType.Category) Or (oElement.DataType = WWWDataType.PopUp) Then
        'if this is a category question
        Set oCategory = GetInactiveCategory(oElement, sValue)
        
        If Not (oCategory Is Nothing) Then
            gsAddInactiveCategory = gsAddInactiveCategory & _
                "fnAddInactiveCategoryValue(" & Chr(34) & oElement.WebId & Chr(34) & "," _
               & CStr(nRepeat) & "," & Chr(34) & oCategory.Code & Chr(34) & "," & Chr(34) & oCategory.Value _
               & Chr(34) & ");" & vbCrLf
        End If
    End If
End Sub

'--------------------------------------------------------------------------------------------------
Private Function GetInactiveCategory(ByRef oElement As eFormElementRO, sValue As String) As CategoryItem
    Dim oCategory As CategoryItem
'--------------------------------------------------------------------------------------------------
'   ic 28/02/2008 issue 2996 - check if this value matches a now inactive category code
'--------------------------------------------------------------------------------------------------
    
    For Each oCategory In oElement.Categories
        If ((Not oCategory.Active) And (oCategory.Code = sValue)) Then
            Set GetInactiveCategory = oCategory
            Exit Function
        End If
    Next
End Function

'--------------------------------------------------------------------------------------------------
Private Function GetEFIInitialisation(ByRef oUser As MACROUser, _
                                      ByRef oSubject As StudySubject, _
                                      ByRef oEFI As EFormInstance, _
                                      ByVal enEFormUse As eEFormUse, _
                                      ByRef sReval() As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 18/09/2003
'   function returns the javascript required to initialise fields in the jsve. created from code
'   split from the original GetEFIInitialisation() function
'   revisions
'   ic 25/02/2004 allow 4 types for hidden fields: integer,real,lab,text
'   ic 23/04/2004 allow 5 types for hidden fields: integer,real,lab,date,text
'   ic 29/06/2004 added error handling
'   ic 01/11/2005 added clinical coding screening
'   ic 29/12/2005 bug changed gsFnViewCommSettings to gsFnViewIComments
'--------------------------------------------------------------------------------------------------
Dim oElement As eFormElementRO
Dim oResponse As Response
Dim oRQG As QGroupRO
Dim oRQGInstance As QGroupInstance
Dim nWinRpt As Integer
Dim nWebRpt As Integer
Dim bRQG As Boolean
Dim sRQGId As String
Dim sDataType As String
Dim nRows As Integer
Dim vJSComm() As String
Dim sLastRQG As String
Dim nRQGOrder As Integer
Dim sValidationType As String


    On Error GoTo CatchAllError
    ReDim vJSComm(0)
    nWebRpt = 0
    sLastRQG = ""


    For Each oElement In oEFI.eForm.EFormElements
        'outer loop creates templates
        With oElement
            'only elements that have responses (not elements like lines and pictures)
            If (.QuestionId > 0) Then
                ' Get OwnerQGroup attribute. If this is set the question is part of a RQG
                Set oRQG = .OwnerQGroup
                If Not (oRQG Is Nothing) Then
                    bRQG = True
                    sRQGId = "g_" & oRQG.Code
                    Set oRQGInstance = oEFI.QGroupInstanceById(oRQG.QGroupID)
                    ' The no. of rows in this group instance
                    ' NB QG Instance may not exist yet...
                    If Not oRQGInstance Is Nothing Then
                        nRows = oRQGInstance.Rows
                    Else
                        nRows = oRQG.InitialRows
                    End If
                    If oRQG.Code <> sLastRQG Then
                        nRQGOrder = 0
                        sLastRQG = oRQG.Code
                    Else
                        nRQGOrder = nRQGOrder + 1
                    End If
                Else
                    bRQG = False
                    nRows = 1
                End If

                'element data type
                If .DataType = NULL_INTEGER Then
                    sDataType = "null"
                ElseIf .Hidden Then
                    Select Case .DataType
                    Case WWWDataType.IntegerData, WWWDataType.Real, WWWDataType.LabTest, WWWDataType.Date
                        sDataType = .DataType
                    Case Else
                        sDataType = WWWDataType.Text
                    End Select
                ElseIf ((.ControlType = ControlType.TextBox) And (.DataType = WWWDataType.Category)) Then
                    'all hidden category elements should be defined as type text
                    'if a text box has been defined as a category question the dataType will still be text box
                    sDataType = WWWDataType.Text
                ElseIf .ControlType = ControlType.PopUp Then
                    'pop-ups have a different datatype to radios
                    sDataType = WWWDataType.PopUp
                ElseIf .DataType = eDataType.Thesaurus Then
                    sDataType = WWWDataType.Text
                Else
                    sDataType = .DataType
                End If

                'get js template function call
                Call AddStringToVarArr(vJSComm, RtnJSTemplate(oElement, (enEFormUse = eEFormUse.User), bRQG, _
                    sDataType, sRQGId, oSubject.StudyDef.FontColour, oSubject.StudyDef.FontSize, oSubject.StudyDef.FontName, _
                    oEFI.eForm.DisplayNumbers))

                If bRQG Then
                    ' Add question to RQG
                    ' e.g. oRQG1.addquestion("webid",nOrder);
                    Call AddStringToVarArr(vJSComm, "aRQG[" & Chr(34) & "g_" & oRQG.Code & Chr(34) _
                        & "].addquestion(" & Chr(34) & .WebId & Chr(34) & "," & nRQGOrder & ");" & vbCrLf)
                End If

                'load caption text into jsve object
                If .Name <> "" And Not .Hidden Then
                    Call AddStringToVarArr(vJSComm, "o1.fnFd('" & .WebId & "','" & "sCaptionText" & "','" _
                        & ReplaceWithJSChars(ReplaceWithHTMLCodes(.Name)) & "');" & vbCrLf)
                End If

                'load help text into jsve object
                If .Helptext <> "" And Not .Hidden Then
                    Call AddStringToVarArr(vJSComm, "o1.fnFd('" & .WebId & "','" & "sHelpText" & "'," _
                        & Chr(34) & ReplaceWithJSChars(.Helptext) & Chr(34) & ");" & vbCrLf)
                End If

                'inner loop creates instances
                For nWinRpt = 1 To nRows
                    'js arrays zero based
                    nWebRpt = nWinRpt - 1

                    'set the response object for this element
                    If nWinRpt > 1 Then
                        Set oResponse = oEFI.Responses.ResponseByElement(oElement, nWinRpt)
                    Else
                        Set oResponse = oEFI.Responses.ResponseByElement(oElement)
                    End If

                    'ic 29/12/2005 bug changed gsFnViewCommSettings to gsFnViewIComments
                    Call AddStringToVarArr(vJSComm, RtnJSInstance(oResponse, oElement, sDataType, bRQG, _
                        CanChangeData(oUser, oSubject.Site), oUser.CheckPermission(gsFnViewIComments), nWebRpt, _
                        sReval))

                    If Not bRQG Then
                        'load hidden additional field handle
                        Call AddStringToVarArr(vJSComm, "o1.fnFd(" & Chr(34) & .WebId & Chr(34) & "," _
                            & Chr(34) & "oAIHandle" & Chr(34) & "," & "o2[" & Chr(34) & "a" & .WebId _
                            & Chr(34) & "]" & "," & nWebRpt & ");" & vbCrLf)
                    End If


                    If Not oResponse Is Nothing Then
                        'lab question validations
                        'if a lab question has a warning raised against it by the server, the jsve wont recognise
                        'it. this is because no client-side validation is done for lab questions so the jsve wont find a
                        'matching validation. to get round this we add a client-side validation for this particular
                        'warning which will fire whenever the question is at its current value
                        'ic 29/08/2003 add validation that will always fire for questions with validation failures
                        'on readonly eforms - so that user can view validation details
                        'ic 02/09/2003 use new CanChangeData function
                        If (oSubject.ReadOnly Or Not CanChangeData(oUser, oSubject.Site)) _
                        Or (.DataType = DataType.LabTest) Then
                            If (oResponse.Status = Status.Warning Or _
                                oResponse.Status = Status.OKWarning Or _
                                oResponse.Status = Status.Inform Or _
                                oResponse.Status = Status.InvalidData) Then
                                Select Case oResponse.Status
                                Case Status.Warning, Status.OKWarning: sValidationType = "1"
                                Case Status.Inform: sValidationType = "2"
                                Case Status.InvalidData: sValidationType = "0"
                                End Select
                                'ic 02/09/2003 use new CanChangeData function
                                If (oSubject.ReadOnly Or Not CanChangeData(oUser, oSubject.Site)) Then
                                    Call AddStringToVarArr(vJSComm, "o1.fnSetValidationRule('" & .WebId & "'," & sValidationType & "," _
                                        & "'jsTrue()','" & Chr(34) & ReplaceWithJSChars(ReplaceLfWithDelimiter(oResponse.ValidationMessage, " ")) _
                                        & Chr(34) & "','true');" & vbCrLf)
                                Else
                                    Call AddStringToVarArr(vJSComm, "o1.fnSetValidationRule('" & .WebId & "'," & sValidationType & "," _
                                        & "'(jsValueOf(" & Chr(34) & .WebId & Chr(34) & "," & Chr(34) & oResponse.RepeatNumber & Chr(34) & ")==" _
                                        & Format(oResponse.Value) & ")','" & Chr(34) & ReplaceWithJSChars(ReplaceLfWithDelimiter(oResponse.ValidationMessage, " ")) _
                                        & Chr(34) & "','true');" & vbCrLf)
                                End If
                            End If
                        End If
                    End If
                Next

                If Not oResponse Is Nothing Then
                    'load responsetaskid into jsve object
                    Call AddStringToVarArr(vJSComm, "o1.fnFd('" & .WebId & "','" & "sResponseTaskID" & "','" _
                        & oResponse.ResponseId & "');" & vbCrLf)
                Else
                    'load responsetaskid into jsve object
                    Call AddStringToVarArr(vJSComm, "o1.fnFd('" & .WebId & "','" & "sResponseTaskID" _
                        & "','');" & vbCrLf)
                End If
            End If
        End With
    Next

    Set oElement = Nothing
    Set oResponse = Nothing
    Set oRQG = Nothing
    Set oRQGInstance = Nothing

    GetEFIInitialisation = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.GetEFIInitialisation"
End Function

'--------------------------------------------------------------------------------------------------
Private Function GetEFIInitialisationRQG(ByRef oEFI As EFormInstance, ByRef sRQGInit() As String, _
                ByRef sRQGResize() As String) As String
'--------------------------------------------------------------------------------------------------
' DPH 17/10/2002 - Create RQG Javascript Initialisation
' revisions
' ic 13/01/2004 added mandatory flag
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------

Dim oElement As eFormElementRO
Dim oRQG As QGroupRO
'Dim sJSString As String
Dim sLocation As String
Dim sWebQGroupID As String
Dim lQGroupId As Long
Dim nInitRows As Integer
Dim nDisplayRows As Integer
Dim nMinRepeats As Integer
Dim nMaxRepeats As Integer
Dim bBorder As Boolean
Dim nRQGCount As Integer
Dim sRQGInitialisation As String
Dim sRQGResizeEnd As String
Dim lTabIndex As Long
Dim vJSComm() As String
Dim bMandatory As Boolean

    On Error GoTo CatchAllError
    ReDim vJSComm(0)
    
   'sJSString = ""
    nRQGCount = 0
    
    For Each oElement In oEFI.eForm.EFormElements
        ' Loop through page elements identifying RQG objects
        Set oRQG = oElement.QGroup
        If Not (oRQG Is Nothing) Then
            ' have found RQG element so retreive information
            ' DPH 26/11/2002 - TabIndex
            If oElement.ElementOrder > 1 Then
                lTabIndex = (oElement.ElementOrder - 1) * mnTABINDEXGAP
            Else
                lTabIndex = 1
            End If

            sLocation = "'window.document.FormDE'"
            sWebQGroupID = "g_" & oRQG.Code
            lQGroupId = oRQG.QGroupID
            nInitRows = oRQG.InitialRows
            nDisplayRows = oRQG.DisplayRows
            nMinRepeats = oRQG.MinRepeats
            nMaxRepeats = oRQG.MaxRepeats
            bBorder = oRQG.Border
            bMandatory = oRQG.Element.IsMandatory
            
'//  function fnCreateRQG(oLocation,sLocation,sID,bEnabled,nQGroupId,
'//                  nInitRows,nDisplayRows,nMinRepeats,nMaxRepeats,bBorder)
            
            ' Create oRQG object
            Call AddStringToVarArr(vJSComm, "if(aRQG[" & Chr(34) & sWebQGroupID & Chr(34) & "]==null){aRQG[" & Chr(34) & sWebQGroupID & Chr(34) & "]=new Object();}" & vbCrLf)
            ' Frame to draw in
            Call AddStringToVarArr(vJSComm, "aRQG[" & Chr(34) & sWebQGroupID & Chr(34) & "]=o1.fnCreateRQG(o2,")
            ' Location
            Call AddStringToVarArr(vJSComm, sLocation & ",")
            ' QGroup Web ID
            Call AddStringToVarArr(vJSComm, "'" & sWebQGroupID & "',")
            ' QGroup Enabled
            Call AddStringToVarArr(vJSComm, "true" & ",")
            ' QGroupId
            Call AddStringToVarArr(vJSComm, lQGroupId & ",")
            ' Initial Rows
            Call AddStringToVarArr(vJSComm, nInitRows & ",")
            ' Display Rows
            Call AddStringToVarArr(vJSComm, nDisplayRows & ",")
            ' Minimum Repeats
            Call AddStringToVarArr(vJSComm, nMinRepeats & ",")
            ' Maximum Repeats
            Call AddStringToVarArr(vJSComm, nMaxRepeats & ",")
            ' Border
            If bBorder Then
                Call AddStringToVarArr(vJSComm, "true,")
            Else
                Call AddStringToVarArr(vJSComm, "false,")
            End If
            ' Max Height
            Call AddStringToVarArr(vJSComm, CStr(CalculateRQGMaxHeight(oEFI, oElement)) & ",")
            'mandatory
            Call AddStringToVarArr(vJSComm, RtnJSBoolean(bMandatory) & "," & vbCrLf)
            ' TabIndex
            Call AddStringToVarArr(vJSComm, lTabIndex & ");" & vbCrLf)
            
            'set up initialisation string for RQG
            ' oRQG1.redraw();
            ' oRQG1.enablefields();
            ' oRQG1.setscroll();
            ' oRQG1.redrawheaders();
            sRQGInitialisation = "aRQG[" & Chr(34) & sWebQGroupID & Chr(34) & "].redraw();" & vbCrLf
            sRQGInitialisation = sRQGInitialisation & "aRQG[" & Chr(34) & sWebQGroupID & Chr(34) & "].enablefields();" & vbCrLf
'            sRQGInitialisation = sRQGInitialisation & "aRQG[" & Chr(34) & sWebQGroupID & Chr(34) & "].setscroll();" & vbCrLf
'            sRQGInitialisation = sRQGInitialisation & "aRQG[" & Chr(34) & sWebQGroupID & Chr(34) & "].resizeRQGDIV();" & vbCrLf
'            sRQGInitialisation = sRQGInitialisation & "aRQG[" & Chr(34) & sWebQGroupID & Chr(34) & "].redrawheaders();" & vbCrLf
            sRQGResizeEnd = "aRQG[" & Chr(34) & sWebQGroupID & Chr(34) & "].setscroll();" & vbCrLf
            sRQGResizeEnd = sRQGResizeEnd & "aRQG[" & Chr(34) & sWebQGroupID & Chr(34) & "].resizeRQGDIV();" & vbCrLf
            sRQGResizeEnd = sRQGResizeEnd & "aRQG[" & Chr(34) & sWebQGroupID & Chr(34) & "].redrawheaders();" & vbCrLf
            ReDim Preserve sRQGInit(nRQGCount)
            sRQGInit(nRQGCount) = sRQGInitialisation
            ReDim Preserve sRQGResize(nRQGCount)
            sRQGResize(nRQGCount) = sRQGResizeEnd
            nRQGCount = nRQGCount + 1
        End If
    Next
    
    GetEFIInitialisationRQG = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.GetEFIInitialisationRQG"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnDerivationString(ByRef oMainEFI As EFormInstance, ByRef oSubsidiaryEFI As EFormInstance, ByVal oAzToJs As AzToJavaScript) As String
'--------------------------------------------------------------------------------------------------
'   ic 13/09/01
'   function returns a string containing the derivations on a passed eform
' AMENDMENTS
'   rem 29/10/01 added if statement to handle category question/text box combination
' MLM 02/10/02: Pass in main and subsidiary EFIs instead of just one. Careful in case the subsidiary one is Nothing.
'   ic 29/06/2004 added error handling
' ic 27/05/2005 issue 2579, ignore comments/lines/pictures
' ic 15/06/2005 issue 2734 - derivations with line breaks cause an error in javascript
'--------------------------------------------------------------------------------------------------
Dim oElement As eFormElementRO
Dim vJsTerms As Variant
Dim sDerivationString As String
Dim ndx As Integer
Dim sCategoryCode As String
Dim vFieldValue As Variant

    On Error GoTo CatchAllError

    'initialise derivation string
    sDerivationString = ""

    'initialise aztojs object for converting derivations
    Call oAzToJs.InitTerms(Derivation)
    
    'loop through each element on the current eform
    For Each oElement In oMainEFI.eForm.EFormElements
    
        'if this element is a valid input element
        If ((oElement.ControlType > 0) And (oElement.ControlType < 10000)) Then
        
            'load the question code and expression into the arezzo to js object for elements that have derivations
            If (oElement.DerivationExpr <> "") Then
                Call oAzToJs.AddTerm(oElement.Code, oElement.DerivationExpr)
            End If
        End If
    Next
    
    'convert the derivation terms from prolog to javascript
    If oSubsidiaryEFI Is Nothing Then
        vJsTerms = oAzToJs.ConvertTerms(oMainEFI.EFormTaskId)
    Else
        vJsTerms = oAzToJs.ConvertTerms(oMainEFI.EFormTaskId, oSubsidiaryEFI.EFormTaskId)
    End If
    
    'if the return is empty no derivations were converted
    'if the return is not empty then it should be an array containing either an error, or converted terms
    If Not IsEmpty(vJsTerms) Then
    
        
        'check that return is an array
        If (UBound(vJsTerms) > 0) Then
        
            'if the first item in the returned array is error string
            If vJsTerms(0) = msAZTOJSERROR Then
                sDerivationString = vJsTerms(1)
                
            Else
        
        
                ndx = 0
                'loop through each element on the current eform
                For Each oElement In oMainEFI.eForm.EFormElements
                
                    'if this element is a valid input element
                    If (oElement.ControlType > 0) Then
                    
                        'if the question has a derivation, add the returned derivation, to the javascript string inside a function call
                        If (oElement.DerivationExpr <> "") Then
                            
                            'rem 29/10/01; added If statment
                            'If the question is a "Category question/textbox" combination then return the Category-Value not Category-Code
                            If (oElement.ControlType = ControlType.TextBox) And (oElement.DataType = WWWDataType.Category) Then
                                sCategoryCode = vJsTerms(ndx)
                                'remove the quotes from the value before using string
                                sCategoryCode = Replace(sCategoryCode, """", "")
                                'replace the quotes around the Category-Value as they are required by the JS function
                                vFieldValue = Chr(34) & oElement.CategoryValue(sCategoryCode) & Chr(34)
                            Else
                                vFieldValue = vJsTerms(ndx)
                                
                                'ic 15/06/2005 issue 2734 - derivations with line breaks cause an error in javascript
                                vFieldValue = Replace(vFieldValue, vbCrLf, " ")
                            End If
                
                            sDerivationString = sDerivationString & "o1.fnSetDerivationRule(" & Chr(39) & oElement.WebId & Chr(39) & "," _
                                                                                            & Chr(39) & vFieldValue & Chr(39) & ");" & vbCrLf
                            ndx = ndx + 1
                        End If
                    End If
                Next
            End If
        End If
    End If


    Set oElement = Nothing
    RtnDerivationString = sDerivationString
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnDerivationString"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnSkipString(ByRef oMainEFI As EFormInstance, ByRef oSubsidiaryEFI As EFormInstance, ByVal oAzToJs As AzToJavaScript) As String
'--------------------------------------------------------------------------------------------------
'   ic 13/09/01
'   function returns a string containing the skips on a passed eform
'
' Revisions:
' MLM 02/10/02: Pass in main and subsidiary EFIs instead of just one. Careful in case the subsidiary one is Nothing.
' DPH 04/12/2002 Include RQG skip conditions
'   ic 29/06/2004 added error handling
' ic 27/05/2005 issue 2579, ignore comments/lines/pictures
'--------------------------------------------------------------------------------------------------
    Dim oElement As eFormElementRO
    Dim vJsTerms As Variant
    Dim sSkipString As String
    Dim ndx As Integer
    Dim oRQG As QGroupRO
    Dim bRQG As Boolean
    Dim sCode As String
    Dim sRQGSkip As String
    
    On Error GoTo CatchAllError
    
    'initialise skip string
    sSkipString = ""

    'initialise aztojs object for converting derivations
    Call oAzToJs.InitTerms(CollectIf)
    
    'loop through each element on the current eform
    For Each oElement In oMainEFI.eForm.EFormElements
    
        ' Check if is a RQG
        Set oRQG = oElement.QGroup
        bRQG = False
        If Not (oRQG Is Nothing) Then
            ' have found RQG element
            bRQG = True
        End If

        'if this element is a valid input element OR RQG
        If ((oElement.ControlType > 0) And (oElement.ControlType < 10000)) Or bRQG Then
        
            sCode = oElement.Code
            If bRQG Then
                sCode = "g_" & oRQG.Code
            End If
            
            'load the question code and expression into the arezzo to js object for elements that have skips
            If (oElement.CollectIfCond <> "") Then
                Call oAzToJs.AddTerm(sCode, oElement.CollectIfCond)
            End If
        End If
    Next
    
    
    'convert the skip terms from prolog to javascript
    If oSubsidiaryEFI Is Nothing Then
        vJsTerms = oAzToJs.ConvertTerms(oMainEFI.EFormTaskId)
    Else
        vJsTerms = oAzToJs.ConvertTerms(oMainEFI.EFormTaskId, oSubsidiaryEFI.EFormTaskId)
    End If
    
    
    'if the return is empty no skips were converted
    'if the return is not empty then it should be an array containing either an error, or converted terms
    If Not IsEmpty(vJsTerms) Then
    
        
        'check that return is an array
        If (UBound(vJsTerms) > 0) Then
        
            'if the first item in the returned array is error string
            If vJsTerms(0) = msAZTOJSERROR Then
                sSkipString = vJsTerms(1)
                
            Else
        
        
                ndx = 0
                'loop through each element on the current eform
                For Each oElement In oMainEFI.eForm.EFormElements
                
                    ' Check if is a RQG
                    Set oRQG = oElement.QGroup
                    bRQG = False
                    If Not (oRQG Is Nothing) Then
                        ' have found RQG element
                        bRQG = True
                    End If
                    
                    'if this element is a valid input element
                    If (oElement.ControlType > 0) Or bRQG Then
                        
                        sCode = oElement.WebId
                        sRQGSkip = "false"
                        If bRQG Then
                            sCode = "g_" & oRQG.Code
                            sRQGSkip = "true"
                        End If
                        
                        'if the question has a skip, add the returned skip, to the javascript string inside a function call
                        If (oElement.CollectIfCond <> "") Then
                
                            sSkipString = sSkipString & "o1.fnSetSkipRule(" & Chr(39) & sCode & Chr(39) & "," _
                                                                              & Chr(39) & vJsTerms(ndx) & Chr(39) _
                                                                              & "," & sRQGSkip & ");" & vbCrLf
                            ndx = ndx + 1
                        End If
                    End If
                Next
            End If
        End If
    End If


    Set oElement = Nothing
    RtnSkipString = sSkipString
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnSkipString"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnValidationString(ByRef oMainEFI As EFormInstance, ByRef oSubsidiaryEFI As EFormInstance, ByVal oAzToJs As AzToJavaScript) As String
'--------------------------------------------------------------------------------------------------
'   ic 13/09/01
'   function returns a string containing the validation on a passed eform
'
' Revisions:
' MLM 02/10/02: Pass in main and subsidiary EFIs instead of just one. Careful in case the subsidiary one is Nothing.
' ic 11/06/2003 replace linefeeds as they cause a javascript crash
'   ic 29/06/2004 added error handling
' ic 27/05/2005 issue 2579, ignore comments/lines/pictures
'--------------------------------------------------------------------------------------------------
    Dim oElement As eFormElementRO
    Dim oValidation As Validation
    Dim vJsTerms As Variant
    Dim vJsMsg As Variant
    Dim sValidationString As String
    Dim ndx As Integer
    
    On Error GoTo CatchAllError
    
    'initialise validation string
    sValidationString = ""
    
    'initialise aztojs for converting validations
    Call oAzToJs.InitTerms(Validation)

    'loop through each element on the current eform
    For Each oElement In oMainEFI.eForm.EFormElements
    
        'if this element is a valid input element
        If ((oElement.ControlType > 0) And (oElement.ControlType < 10000)) Then
    
            For Each oValidation In oElement.Validations
                'load the question code and expression into the arezzo to js object
                Call oAzToJs.AddTerm(oElement.Code, oValidation.ValidationCond)
            Next
        End If
    Next

    'convert the terms from prolog to javascript
    If oSubsidiaryEFI Is Nothing Then
        vJsTerms = oAzToJs.ConvertTerms(oMainEFI.EFormTaskId)
    Else
        vJsTerms = oAzToJs.ConvertTerms(oMainEFI.EFormTaskId, oSubsidiaryEFI.EFormTaskId)
    End If
    
    
    'if the return is empty no validations were converted
    'if the return is not empty then it should be an array containing either an error, or converted terms
    If Not IsEmpty(vJsTerms) Then
        
        'check that return is an array
        If UBound(vJsTerms) > 0 Then
        
            'if the first item in the returned array is error string
            If vJsTerms(0) = msAZTOJSERROR Then
                sValidationString = vJsTerms(1)
               
            Else

                'initialise aztojs object for retrieving validation messages
                Call oAzToJs.InitTerms(ValidationMsg)

                'loop through each element on the current eform
                For Each oElement In oMainEFI.eForm.EFormElements
                
                    'if this element is a valid input element
                    If (oElement.ControlType > 0) Then
                   
                        For Each oValidation In oElement.Validations
                            'ic 11/06/2003 replace linefeeds as they cause a javascript crash
                            'load the question code and expression into the arezzo to js object
                            Call oAzToJs.AddTerm(oElement.Code, ReplaceLfWithDelimiter(oValidation.MessageExpr, " "))
                        Next
                    End If
                Next

                'retrieve the messages
                If oSubsidiaryEFI Is Nothing Then
                    vJsMsg = oAzToJs.ConvertTerms(oMainEFI.EFormTaskId)
                Else
                    vJsMsg = oAzToJs.ConvertTerms(oMainEFI.EFormTaskId, oSubsidiaryEFI.EFormTaskId)
                End If
    
                'if the return is empty no messages were retrieved
                'if the return is not empty then it should be an array containing either an error, or converted terms
                If Not IsEmpty(vJsTerms) Then
                
                    'if the first item in the returned array is error string
                    If vJsTerms(0) = msAZTOJSERROR Then
                        sValidationString = vJsTerms(1)
        
                    Else
                    
                        ndx = 0
                        'loop through each question on the current eform
                        For Each oElement In oMainEFI.eForm.EFormElements
                            For Each oValidation In oElement.Validations
                                
                                'if this element is a valid input element
                                If (oElement.ControlType > 0) Then
                                    'add the returned derivations, if any to the javascript string
                                    sValidationString = sValidationString & "o1.fnSetValidationRule(" & Chr(39) & oElement.WebId & Chr(39) & "," _
                                                                               & oValidation.ValidationType & "," _
                                                                               & Chr(39) & vJsTerms(ndx) & Chr(39) & "," _
                                                                               & Chr(39) & Replace(vJsMsg(ndx), vbCrLf, "\n") & Chr(39) & "," _
                                                                               & Chr(39) & ReplaceWithJSChars(oValidation.ValidationCond) & Chr(39) & ");" & vbCrLf
                                    ndx = ndx + 1
                                End If
                            Next
                        Next
                    End If
                End If
            End If
        End If
    End If
    
    
    Set oValidation = Nothing
    Set oElement = Nothing
    RtnValidationString = sValidationString
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RtnValidationString"
End Function

'-------------------------------------------------------------------------------------------'
Private Function GetDateOkFunc(ByRef enEFormUse As eEFormUse, ByRef oEform As eFormRO) As String
'-------------------------------------------------------------------------------------------'
' MLM 25/09/02: Created. Returns a JavaScript function that returns false if there is an blank
'   enterable form/visit date (which should prevent the form from being saved).
'   enEFormUse determines the name of the function.
' MLM 10/10/02: Allow oEForm to be Nothing, which results in a function that always returns true.
'   Also, look for JSVE in o1.
'   revisions
'   ic 29/06/2004 added error handling
'-------------------------------------------------------------------------------------------'

Dim sResult As String
Dim oFormDate As eFormElementRO

    On Error GoTo CatchAllError

    'function header
    If enEFormUse = eEFormUse.VisitEForm Then
        sResult = "function VisitDateOk(){"
    Else
        sResult = "function FormDateOk(){"
    End If
    
    'function body
    If oEform Is Nothing Then
        'always ok to save
        sResult = sResult & "return true"
    Else
        Set oFormDate = oEform.EFormDateElement
        If oFormDate Is Nothing Then
            'always ok to save
            sResult = sResult & "return true"
        Else
            'MLM 10/10/02: JSVE is now in o1
            sResult = sResult & "return !(o1.fnEnterable('" & oFormDate.WebId _
                & "',0) && (''==o1.fnGetFormatted('" & oFormDate.WebId & "',0)))"
        End If
    End If
    
    'function terminator
    GetDateOkFunc = sResult & ";}"
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.GetDateOkFunc"
End Function

'-------------------------------------------------------------------------------------------'
Private Function CalculateRQGMaxHeight(oEFI As EFormInstance, _
                                oRQGElement As eFormElementRO) As Long
'-------------------------------------------------------------------------------------------'
' Calculate the maximum height a RQG can be without overlapping the next
' element (below) it on the eForm
'-------------------------------------------------------------------------------------------'
' REVISIONS
' DPH 29/01/2002 - Adapted RQG max size if last element on an eForm
' DPH 13/04/2004 - Allow for scrollbar height if last element on eForm
'   ic 29/06/2004 added error handling
'-------------------------------------------------------------------------------------------'
Dim lMaxHeight As Long
Dim lTopPosRQG As Long
Dim lLeftPosRQG As Long
Dim lBotPosRQG As Long
Dim lTopPosElement As Long
Dim lLeftPosElement As Long
Dim lRightPosElement As Long
Dim oElement As eFormElementRO
Dim oRQG As QGroupRO
Dim oOwnerRQG As QGroupRO
Dim bBelongsToRQG As Boolean
Dim bRQG As Boolean
Dim bRQGLastElement As Boolean
Dim lRowHeight As Long
Dim nMaxElements As Integer

    On Error GoTo CatchAllError

    ' Get top (Y) position of RQG dealing with
    ' Add 24 for RQG header DIV
    lTopPosRQG = oRQGElement.ElementY + (24 * nIeYSCALE)
    lLeftPosRQG = oRQGElement.ElementX
    ' Make max height RQG an arbitary Max value
    lBotPosRQG = lTopPosRQG + 8000
    bRQGLastElement = True
    
    For Each oElement In oEFI.eForm.EFormElements
        ' Loop through page elements identifying Non-RQG
        ' and Non-RQG question elements
        Set oRQG = oElement.QGroup
        Set oOwnerRQG = oElement.OwnerQGroup
        If oRQG Is Nothing Then
            bRQG = False
        Else
            bRQG = True
        End If
        If oOwnerRQG Is Nothing Then
            bBelongsToRQG = False
        Else
            bBelongsToRQG = True
        End If

        If Not bRQG And Not bBelongsToRQG Then
            ' Get position on eForm of Element
            lTopPosElement = oElement.ElementY
            lLeftPosElement = oElement.ElementX
            
            ' if lower on form than RQG element then compare else ignore
            If lTopPosElement > lTopPosRQG Then
            
                bRQGLastElement = False
                
                If lTopPosElement < lBotPosRQG Then
                    ' Now check for horizontal overlap
                    ' can do if element comparing with is to left of rqg
                    'If lLeftPosElement < lLeftPosRQG Then
                        lBotPosRQG = lTopPosElement
                    'End If
                End If
            End If
        End If
    Next
    
    ' if RQG is last question on a page
    If bRQGLastElement Then
        ' try to calculate a more realistic max height
        Set oRQG = oRQGElement.QGroup
        lMaxHeight = 0
        If Not (oRQG Is Nothing) Then
            ' set default row height
            lRowHeight = 30 * nIeYSCALE
            ' reset nMaxElements
            nMaxElements = 1
            ' loop through RQG Questions and detect any Radio Buttons
            For Each oElement In oRQG.Elements
                If oElement.ControlType = ControlType.OptionButtons Then
                    ' Get number of elements in Category (height of radio buttons)
                    If oElement.Categories.Count > nMaxElements Then
                        nMaxElements = oElement.Categories.Count
                    End If
                End If
            Next
            lMaxHeight = (lRowHeight * oRQG.DisplayRows * nMaxElements) / nIeYSCALE
            ' DPH 13/04/2004 - Allow for scrollbar height if last element on eForm
            lMaxHeight = lMaxHeight + 20
        End If
    Else
        ' calculate Maximum height (with web scale) - 5 (so not attached to next element)
        lMaxHeight = ((lBotPosRQG - lTopPosRQG) / nIeYSCALE) - 5
    End If
    
    CalculateRQGMaxHeight = lMaxHeight
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.CalculateRQGMaxHeight"
End Function

'-------------------------------------------------------------------------------------------'
Private Function IsDefaultCaptionFont(ByRef oCRFElement As eFormElementRO, ByVal lStudyFontCol As Long, _
                        ByVal nStudyFontSize As Integer, ByVal sStudyFontName As String) As Boolean
'-------------------------------------------------------------------------------------------'
' Returns whether default caption font is to be used
'   revisions
'   ic 29/06/2004 added error handling
'-------------------------------------------------------------------------------------------'
Dim bDefaultCaptionFont As Boolean

    On Error GoTo CatchAllError

    bDefaultCaptionFont = True
    
    'ZA - 23/07/2002 - () As Boolean
    'ZA - 23/07/2002 - check font attributes for caption/comment
    If oCRFElement.Caption <> "" Then
        If ((oCRFElement.CaptionFontColour = lStudyFontCol) _
        And (oCRFElement.CaptionFontSize = nStudyFontSize) _
        And (oCRFElement.CaptionFontName = sStudyFontName)) _
        Or ((oCRFElement.CaptionFontColour = lStudyFontCol) _
        And (oCRFElement.CaptionFontSize = nStudyFontSize) _
        And (oCRFElement.CaptionFontName = sStudyFontName)) Then
            bDefaultCaptionFont = True
        Else
            bDefaultCaptionFont = False
        End If
    End If

    IsDefaultCaptionFont = bDefaultCaptionFont
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.IsDefaultCaptionFont"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RevalidateEFI(ByRef colFields As Collection, ByVal bUReadOnly As Boolean, _
    ByVal bVReadOnly As Boolean, ByRef bUReject As Boolean, ByRef bVReject As Boolean) As Boolean
'--------------------------------------------------------------------------------------------------
' Returns TRUE if revalidating the EFI made something change.
'--------------------------------------------------------------------------------------------------
' REVISIONS
' ic 02/03/2004 amended for revalidation on save
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim oField As WWWField
Dim oResponse As Response
Dim bChangedStatus As Boolean
Dim sErrMsg As String
Dim lStatus As Long
Dim bSomethingChanged As Boolean


    On Error GoTo CatchAllError
    bSomethingChanged = False
    bUReject = False
    bVReject = False
     
    'loop through all the fields in the fields collection submitted by the eform
    For Each oField In colFields
        With oField

            'revalidate if allowed
            If (.EformUse = eEFormUse.VisitEForm And Not bVReadOnly) _
            Or (.EformUse = eEFormUse.User And Not bUReadOnly) Then
                If .oResponse.Revalidatable(False) Then
                    lStatus = .oResponse.RevalidateValue(sErrMsg, bChangedStatus)
                    If (lStatus = eStatus.InvalidData) Then
                        If (.EformUse = eEFormUse.VisitEForm) Then
                            bVReject = True
                        Else
                            bUReject = True
                        End If
                    End If
                    
                    If bChangedStatus Then
                        bSomethingChanged = True
                        'confirm revalidation, passing the rfo that may have
                        'been entered by the user
                        Call .oResponse.ConfirmRevalidation(.sRFO)
                    Else
                        'ignore the revalidation
                        Call .oResponse.RejectValue
                    End If
                End If
            End If
        End With
    Next
    
    'Set oResponse = Nothing
    RevalidateEFI = bSomethingChanged
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLEform.RevalidateEFI"
End Function
