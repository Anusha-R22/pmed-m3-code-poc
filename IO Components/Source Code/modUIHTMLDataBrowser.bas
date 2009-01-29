Attribute VB_Name = "modUIHTMLDataBrowser"
'----------------------------------------------------------------------------------------'
'   File:       modUIHTML.bas
'   Copyright:  InferMed Ltd. 2000-2007. All Rights Reserved
'   Author:     i curtis 02/2003
'   Purpose:    functions returning html versions of MACRO pages (DATABROWSER)
'----------------------------------------------------------------------------------------'
' revisions
' ic 29/05/2003 display eform label if present in GetDataBrowser()
' ic 29/05/2003 only display comments if user has permission GetDataBrowser(), bug 1816
' ic 30/05/2003 moved error display code out of conditional statement in GetDataBrowser(),
'               now will always display errors
' ic 30/05/2003 dont display 'goto eform' menu option in study,visit column in
'               GetDataBrowser() bug 1817
' ic 05/06/2003 added extra 'refresh z-order' parameter in GetDataBrowser()
' ic 06/05/2003 added ReplaceWithJSChars() call around sSubjectLabel in GetDataBrowser()
' DPH 21/05/2003 - Added border lines to table (size 1) in GetDataBrowser
' DPH 18/06/2003 - Aligned record count / previous / next buttons to left in GetDataBrowser
' ic 26/06/2003 convert value to local format in GetDataBrowser(), bug 1873
' DPH 01/07/2003 bug 1891 Allow discrepancies,SDVs when question locked, disallow notes when question frozen in GetDataBrowser
' DPH 08/10/2003 Performance changes - use DataItemResponse/DataItemResponseHistory table in RtnDataBrowser
' ic 29/06/2004 added error handling
' ic 19/04/2005 bug 2505, changed the b4 value in the querystring to javascript boolean (1 or 0) in GetDataBrowser()
' ic 28/04/2005 issue 2516, fixed form post url bookmark parameter in GetDataBrowser()
' ic 03/05/2005 issue 2110, added a print icon in GetDataBrowser()
' ic 28/07/2005 added clinical coding
' NCJ 8-19 Dec 05 - New Date/Time types
' ic 26/10/2006 issue 2830 question fnM() parameter replaced questionid parameter in GetDataBrowser()
' NCJ 26 Mar 07 - Issue 2893 - Show "Inform" status correctly
' ic 17/04/2007 issue 2837, can create sdv if locked but not frozen (to match windows)
'----------------------------------------------------------------------------------------'

Option Explicit

'---------------------------------------------------------------------
Private Function eFormTitleLabel(ByVal sTitle As String, ByVal sLabel As String, ByVal sCycle As String) As String
'---------------------------------------------------------------------
'TA 27/05/2003: return label if it exists or title if not
'   revisions
'   ic 29/06/2004 added error handling
'---------------------------------------------------------------------

    On Error GoTo CatchAllError

    If sLabel = "" Then
        eFormTitleLabel = sTitle & "[" & sCycle & "]"
    Else
        eFormTitleLabel = sLabel
    End If
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLDataBrowser.eFormTitleLabel"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnLockColour(ByVal nStatus As Integer)
'--------------------------------------------------------------------------------------------------
' function returns bgcolour string based on cell's lockstatus
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
    Dim sCol As String

    On Error GoTo CatchAllError

    Select Case nStatus
    Case LockStatus.lsLocked:
        sCol = " bgcolor='#d2b48c'"
    Case LockStatus.lsFrozen:
        sCol = " bgcolor='#add8e6'"
    End Select
    
    RtnLockColour = sCol
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLDataBrowser.RtnLockColour"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnMenuLockCallParams(ByVal nStatus As Integer, ByVal bLock As Boolean, ByVal bUnlock As Boolean, _
                                       ByVal bFreeze As Boolean) As String
'--------------------------------------------------------------------------------------------------
' function returns 3 comma delimited flags for enabling jscript note,sdv,discrepancy menu options
'   revisions
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
    Dim sRtn As String
    
    On Error GoTo CatchAllError
    
    Select Case nStatus
    Case LockStatus.lsLocked:
        sRtn = "0,"
        If bUnlock Then
            sRtn = sRtn & "1,"
        Else
            sRtn = sRtn & "0,"
        End If
        If bFreeze Then
            sRtn = sRtn & "1,"
        Else
            sRtn = sRtn & "0,"
        End If
    Case LockStatus.lsFrozen:
        sRtn = "0,0,0,"
    Case Else:
        If bLock Then
            sRtn = "1,0,"
        Else
            sRtn = "0,0,"
        End If
        If bFreeze Then
            sRtn = sRtn & "1,"
        Else
            sRtn = sRtn & "0,"
        End If
    End Select
    
    RtnMenuLockCallParams = sRtn
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLDataBrowser.RtnMenuLockCallParams"
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetDataBrowser(ByRef oUser As MACROUser, ByVal nStudyCode As Integer, _
                               ByVal sSiteCode As String, ByVal lVisitCode As Long, _
                               ByVal lCRFPageId As Long, ByVal lQuestion As Long, _
                               ByVal sSrchUserName As String, ByVal lSubjectId As Long, _
                               ByVal sSubjectLabel As String, ByVal sStatus As String, _
                               ByVal sLockStatus As String, sTime As String, _
                               ByVal bBefore As Boolean, ByVal lComment As Long, _
                               ByVal lDiscrepancy As Long, ByVal lSDV As Long, _
                               ByVal lNote As Long, ByVal sType As String, _
                               ByVal sDecimalPoint As String, ByVal sThousandSeparator As String, _
                               Optional ByVal enInterface As eInterface = iwww, _
                               Optional ByVal lBookmark As Long = 0, _
                               Optional ByVal vErrors As Variant) As String
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
' REVISIONS
' DPH 03/03/2003 - Display Value calculated outside of 'current' If so can display in Audit mode
' ic 27/03/2003 added user name and database timestamp to columns
' DPH 21/05/2003 - Added border lines to table (size 1)
' ic 29/05/2003 display eform label if present
' ic 29/05/2003 only display comments if user has permission, bug 1816
' ic 30/05/2003 moved error display code out of conditional statement, now will always display errors
' ic 30/05/2003 dont display 'goto eform' menu option in study,visit column bug 1817
' ic 05/06/2003 added extra 'refresh z-order' parameter
' ic 06/05/2003 added ReplaceWithJSChars() call around sSubjectLabel
' DPH 21/05/2003 - Added border lines to table (size 1)
' DPH 18/06/2003 - Aligned record count / previous / next buttons to left
' ic 26/06/2003 convert value to local format, bug 1873
' DPH 01/07/2003 bug 1891 Allow discrepancies,SDVs when question locked, disallow notes when question frozen
' ic 29/06/2004 added error handling
' ic 19/04/2005 bug 2505, changed the b4 value in the querystring to javascript boolean (1 or 0)
' ic 28/04/2005 issue 2516, fixed form post url bookmark parameter
' ic 03/05/2005 issue 2110, added a print icon
' ic 28/07/2005 added clinical coding
' ic 26/10/2006 issue 2830 question fnM() parameter replaced questionid parameter
' NCJ 26 Mar 07 - Issue 2893 - Show "Inform" status correctly
' ic 17/04/2007 issue 2837, can create sdv if locked but not frozen (to match windows)
'--------------------------------------------------------------------------------------------------

Dim bLock As Boolean
Dim bUnlock As Boolean
Dim bFreeze As Boolean
Dim bUnFreeze As Boolean
Dim bAddDiscrepancy As Boolean
Dim bAddSDV As Boolean
Dim bAddNote As Boolean
Dim bViewAudit As Boolean
Dim sRtn As String
Dim vData As Variant
Dim nMaxRecords As Integer
Dim lPageLength As Long
Dim bCurrent As Boolean
Dim vJSComm() As String
Dim sURL As String
Dim nLoop As Integer
Dim lSpan As Long
Dim lCol1Span As Long
Dim lCol2Span As Long
Dim lCol3Span As Long
Dim lLoop As Long
Dim lStart As Long
Dim lStop As Long
Dim sLabel1 As String
Dim sLabel2 As String
Dim sValue As String
Dim sStatusLabel As String
Dim bViewInformIcon As Boolean
Dim lRaised As Long
Dim lResponded As Long
Dim lPlanned As Long
Dim bViewComments As Boolean

    On Error GoTo CatchAllError
    ReDim vJSComm(0)

    Call RtnMIMsgStatusCount(oUser, lRaised, lResponded, lPlanned)
    bLock = oUser.CheckPermission(gsFnLockData)
    bUnlock = oUser.CheckPermission(gsFnUnLockData)
    bFreeze = oUser.CheckPermission(gsFnFreezeData)
    bUnFreeze = oUser.CheckPermission(gsFnUnFreezeData)
    bAddDiscrepancy = oUser.CheckPermission(gsFnCreateDiscrepancy)
    bAddSDV = oUser.CheckPermission(gsFnCreateSDV)
    bViewAudit = oUser.CheckPermission(gsFnViewAuditTrail)
    bViewInformIcon = oUser.CheckPermission(gsFnMonitorDataReviewData)
    bViewComments = oUser.CheckPermission(gsFnViewIComments)
    
    bAddNote = True
    If (sType = "1") Or (sType = "3") Then bCurrent = True
    
    'ic 28/07/2005 added clinical coding
    vData = RtnDataBrowser(oUser, True, nStudyCode, sSiteCode, lVisitCode, lCRFPageId, lQuestion, sSrchUserName, _
                           lSubjectId, sSubjectLabel, sStatus, sLockStatus, sTime, bBefore, lComment, _
                           lDiscrepancy, lSDV, lNote, -1, "", "", sType, False, vErrors)

    If (enInterface = iwww) Then
        nMaxRecords = gnMAXWWWRECORDS
    Else
        nMaxRecords = gnMAXWINRECORDS
    End If
    
'ic todo 06/05/2003
'may still be a problem with certain characters interrupting the js in sSubjectLabel
    sURL = "DataBrowser.asp?st=" & CStr(nStudyCode) & "&si=" & sSiteCode & "&vi=" & CStr(lVisitCode) _
                                & "&ef=" & CStr(lCRFPageId) & "&qu=" & CStr(lQuestion) & "&us=" & sSrchUserName _
                                & "&sj=" & CStr(lSubjectId) & "&sjlb=" & URLEncodeString(sSubjectLabel) & "&ss=" & sStatus _
                                & "&lk=" & sLockStatus & "&tm=" & URLEncodeString(sTime) & "&b4=" & RtnJSBoolean(bBefore) _
                                & "&cm=" & CStr(lComment) & "&di=" & CStr(lDiscrepancy) & "&sd=" & CStr(lSDV) _
                                & "&no=" & CStr(lNote) & "&get=" & sType
                                        
    If (enInterface = iwww) Then
        lPageLength = CLng(oUser.UserSettings.GetSetting(SETTING_PAGE_LENGTH, 50))
    Else
        lPageLength = UBound(vData, 2)
    End If
                            
    'ic start body html
    Call AddStringToVarArr(vJSComm, "<body onload='fnPageLoaded();'>" & vbCrLf)

    Call AddStringToVarArr(vJSComm, "<script language='javascript'>" & vbCrLf _
        & "function fnPageLoaded(){" & vbCrLf _
        & "window.sWinState=" & Chr(34) & "6|" & CStr(nStudyCode) & "|" & sSiteCode & "|" & CStr(lVisitCode) & "|" _
            & CStr(lCRFPageId) & "|" & CStr(lQuestion) & "|" & URLEncodeString(sSubjectLabel) & "|" & sSrchUserName & "|" _
            & sStatus & "|" & sLockStatus & "|" & CStr(RtnJSBoolean(bBefore)) & "|" & URLEncodeString(sTime) & "|" & CStr(lComment) & "|" _
            & CStr(lDiscrepancy) & "|" & CStr(lSDV) & "|" & CStr(lNote) & "|" & sType & "|" & CStr(lBookmark) & Chr(34) & ";" & vbCrLf _
        & "fnHideLoader();" & vbCrLf)
        
    'ic 05/06/2003 added extra 'refresh z-order' parameter
    Call AddStringToVarArr(vJSComm, "window.parent.fnSTLC('" & gsVIEW_RAISED_DISCREPANCIES_MENUID & "','" & CStr(lRaised) & "',0);" & vbCrLf _
                          & "window.parent.fnSTLC('" & gsVIEW_RESPONDED_DISCREPANCIES_MENUID & "','" & CStr(lResponded) & "',0);" & vbCrLf _
                          & "window.parent.fnSTLC('" & gsVIEW_PLANNED_SDV_MARKS_MENUID & "','" & CStr(lPlanned) & "',1);" & vbCrLf)
        
    'errors encountered during save
    If Not IsMissing(vErrors) Then
        If Not IsEmpty(vErrors) Then
            Call AddStringToVarArr(vJSComm, "alert('MACRO encountered problems while updating. Some updates could not be completed." _
                & "\nIncomplete updates are listed below\n\n")

            For nLoop = LBound(vErrors, 2) To UBound(vErrors, 2)
                Call AddStringToVarArr(vJSComm, vErrors(0, nLoop) & " - " & vErrors(1, nLoop) & "\n")
            Next

            Call AddStringToVarArr(vJSComm, "');" & vbCrLf)
        End If
    End If
        
    Call AddStringToVarArr(vJSComm, "}" & vbCrLf _
        & "</script>" & vbCrLf)
    
    
    
    If (vData(0, 0) < nMaxRecords) Then
    'ic 28/07/2005 added clinical coding
        vData = RtnDataBrowser(oUser, False, nStudyCode, sSiteCode, lVisitCode, lCRFPageId, lQuestion, sSrchUserName, _
                               lSubjectId, sSubjectLabel, sStatus, sLockStatus, sTime, bBefore, lComment, _
                               lDiscrepancy, lSDV, lNote, -1, "", "", sType)
        
        If (Not IsNull(vData)) Then
            'current data/audit trail
            If (sType = "1") Or (sType = "2") Or (sType = "3") Then
            
                Call AddStringToVarArr(vJSComm, "<form name='FormDR' action='" & sURL & "&bookmark=" & CStr(lBookmark) & "' method='post'>" & vbCrLf _
                    & "<input type='hidden' name='bidentifier'>" _
                    & "<input type='hidden' name='btype'>" _
                    & "<input type='hidden' name='badd'>" _
                    & "<input type='hidden' name='bscope'>" _
                    & "</form>" & vbCrLf)
            
                
                'initialise rowspan variables
                lCol1Span = 0
                lCol2Span = 0
                lCol3Span = 0
    
                'calculate the start row and end row based on start row (bookmark) and page length
                If ((lBookmark >= UBound(vData, 2)) Or (lBookmark < 0)) Then
                    lStart = 0
                Else
                    lStart = lBookmark
                End If
                If ((lStart + lPageLength) >= UBound(vData, 2)) Then
                    lStop = UBound(vData, 2)
                Else
                    lStop = (lStart + lPageLength) - 1
                End If
                
                ' DPH 21/05/2003 - Added border lines to table (size 1)
                ' DPH 18/06/2003 - Aligned record count / previous / next buttons to left
                Call AddStringToVarArr(vJSComm, "<table style='cursor:default;' width='100%' class='clsTabletext' cellpadding='0' cellspacing='0' border='1'>" _
                    & "<tr height='30'><td colspan='14' align='left'>" & vbCrLf)
                Call AddStringToVarArr(vJSComm, "Record(s) " & lStart + 1 & " to " & lStop + 1 & " of " & UBound(vData, 2) + 1 & "&nbsp;&nbsp;" & vbCrLf)
                
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
                
                'table header - column names
                Call AddStringToVarArr(vJSComm, "&nbsp;<a href='javascript:window.print();'>" _
                & "<img src='../img/ico_print.gif' border='0' alt='Print listing'></a></td></tr><tr height='20' class='clsTableHeaderText'>" _
                    & "<td>Study/Site/Subject</td><td>Visit</td><td>eForm</td><td>Question</td>" _
                    & "<td>Value</td><td>Status</td><td>Date and time</td><td>Database date and time</td><td>User Name</td><td>Full User Name</td><td width='200'>Comment</td>" _
                    & "<td>Reason For Change</td><td>Overrule Reason</td><td>Warning Message</td></tr>" & vbCrLf)
                
                
                'loop through 2 dimensional array[x,y]. [x] dimension is cols, [y] dimension is rows
                'each loop adds a new row
                For lLoop = lStart To lStop
    
                    Call AddStringToVarArr(vJSComm, "<tr")
                    'If bCurrent Then Call AddStringToVarArr(vJSComm, " class='LINK'")
                    Call AddStringToVarArr(vJSComm, ">")

                    'trial/site/subject/label column
                    'lCol1Span variable holds the number of rows the current column row will span
                    If lCol1Span = 0 Then
                        'calculate the new span, if there is one, by looping through the array and counting
                        'the number of following rows where trial/site/subject/label values are the same as this one
                        For lSpan = lLoop To lStop
                            If IsNull(vData(DataBrowserCol.dbcSubjectLabel, lLoop)) Then
                                sLabel1 = ""
                            Else
                                sLabel1 = vData(DataBrowserCol.dbcSubjectLabel, lLoop)
                            End If
                            If IsNull(vData(DataBrowserCol.dbcSubjectLabel, lSpan)) Then
                                sLabel2 = ""
                            Else
                                sLabel2 = vData(DataBrowserCol.dbcSubjectLabel, lSpan)
                            End If
    
                            If ((CStr(vData(DataBrowserCol.dbcStudyId, lLoop)) <> CStr(vData(DataBrowserCol.dbcStudyId, lSpan))) _
                            Or (CStr(vData(DataBrowserCol.dbcSite, lLoop)) <> CStr(vData(DataBrowserCol.dbcSite, lSpan))) _
                            Or (CStr(vData(DataBrowserCol.dbcSubjectId, lLoop)) <> CStr(vData(DataBrowserCol.dbcSubjectId, lSpan))) _
                            Or (CStr(sLabel1) <> CStr(sLabel2))) Then Exit For
                            'each time we find a match, increment the rowspan variable
                            lCol1Span = lCol1Span + 1
                        Next

                        'if current data (not audit trail), write hyperlinks
                        If bCurrent Then
                            'open an anchor around the <td> with an onclick event that displays a context menu
                            'js fnM arguements are mousebutton,itemid,scope,name,value,lock,unlock,freeze,goto eform,discrepancy,sdv,note
                            'itemid=delimited list: study[0]site[1]subject[2]visitid[3]visitcycle[4]visittaskid[5]crfpageid[6]
                            'crfpagecycle[7]crfpagetaskid[8]responseid[9]responsecycle[10]responsetaskid[11] (any/all)
                            Call AddStringToVarArr(vJSComm, "<a onmouseup=" & Chr(34))
                            Call AddStringToVarArr(vJSComm, "fnM(event.button,'" & vData(DataBrowserCol.dbcStudyId, lLoop) & gsDELIMITER1 _
                                              & vData(DataBrowserCol.dbcSite, lLoop) & gsDELIMITER1 _
                                              & vData(DataBrowserCol.dbcSubjectId, lLoop) & "'," _
                                              & MIMsgScope.mimscSubject & "," _
                                              & LFScope.lfscSubject & "," _
                                              & "'" & RtnSubjectText(vData(DataBrowserCol.dbcSubjectId, lLoop), vData(DataBrowserCol.dbcSubjectLabel, lLoop)) & "'," _
                                              & "'',")
                            'boolean values passed to showMenu() function are lock,unlock,freeze,goto eform,add discrepancy,add sdv,add note
                            Call AddStringToVarArr(vJSComm, RtnMenuLockCallParams(vData(DataBrowserCol.dbcSubjectLockStatus, lLoop), bLock, bUnlock, bFreeze))
                            ' Unfreeze param
                            If bUnFreeze Then
                                If vData(DataBrowserCol.dbcSubjectLockStatus, lLoop) = LockStatus.lsFrozen Then
                                    Call AddStringToVarArr(vJSComm, "1,")
                                Else
                                    Call AddStringToVarArr(vJSComm, "0,")
                                End If
                            Else
                                Call AddStringToVarArr(vJSComm, "0,")
                            End If
                            'ic 30/05/2003 cant goto eform from this column
                            Call AddStringToVarArr(vJSComm, "0,0,")
                            'ic 17/04/2007 issue 2837, can create sdv if locked but not frozen (to match windows)
                            If (bAddSDV And (vData(DataBrowserCol.dbcSubjectLockStatus, lLoop) <> LockStatus.lsFrozen)) Then
                                Call AddStringToVarArr(vJSComm, "1,")
                            Else
                                Call AddStringToVarArr(vJSComm, "0,")
                            End If
                            Call AddStringToVarArr(vJSComm, "0")
                            Call AddStringToVarArr(vJSComm, ");" & Chr(34) & ">" & vbCrLf)
                        End If
    
                        'write the <td> with a bgcolor returned by function rtnLockColour() based on lock status
                        Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcSubjectLockStatus, lLoop)))
                        'include a rowspan if the rowspan > 1
                        If lCol1Span > 1 Then Call AddStringToVarArr(vJSComm, " rowspan=" & lCol1Span)
    
                        'insert the value and close the </td>
                        Call AddStringToVarArr(vJSComm, ">" & vData(DataBrowserCol.dbcStudyName, lLoop) & "/" & _
                            vData(DataBrowserCol.dbcSite, lLoop) & "/" & RtnSubjectText(vData(DataBrowserCol.dbcSubjectId, lLoop), vData(DataBrowserCol.dbcSubjectLabel, lLoop)) _
                                     & "&nbsp;" & RtnStatusImages(vData(DataBrowserCol.dbcsubjectStatus, lLoop), bViewInformIcon, vData(DataBrowserCol.dbcSubjectLockStatus, lLoop), False, vData(DataBrowserCol.dbcSubjectSDVStatus, lLoop), vData(DataBrowserCol.dbcSubjectDiscStatus, lLoop)) & "</td>" & vbCrLf)
    
                        'close the anchor after the </td>
                        If bCurrent Then Call AddStringToVarArr(vJSComm, "</a>")
                    End If
                    'every time we add a row, decrement the rowspan variable so we keep a count of how many rows the
                    'current <td> cell has still got left to span
                    lCol1Span = lCol1Span - 1
    

                    'visit column
                    If lCol2Span = 0 Then
                        For lSpan = lLoop To lStop
                            If IsNull(vData(DataBrowserCol.dbcSubjectLabel, lLoop)) Then
                                sLabel1 = ""
                            Else
                                sLabel1 = vData(DataBrowserCol.dbcSubjectLabel, lLoop)
                            End If
                            If IsNull(vData(DataBrowserCol.dbcSubjectLabel, lSpan)) Then
                                sLabel2 = ""
                            Else
                                sLabel2 = vData(DataBrowserCol.dbcSubjectLabel, lSpan)
                            End If
    
                            If ((CStr(vData(DataBrowserCol.dbcStudyName, lLoop)) <> CStr(vData(DataBrowserCol.dbcStudyName, lSpan))) _
                            Or (CStr(vData(DataBrowserCol.dbcSite, lLoop)) <> CStr(vData(DataBrowserCol.dbcSite, lSpan))) _
                            Or (CStr(vData(DataBrowserCol.dbcSubjectId, lLoop)) <> CStr(vData(DataBrowserCol.dbcSubjectId, lSpan))) _
                            Or (CStr(sLabel1) <> CStr(sLabel2)) _
                            Or (CStr(vData(DataBrowserCol.dbcVisitName, lLoop)) <> CStr(vData(DataBrowserCol.dbcVisitName, lSpan))) _
                            Or (CStr(vData(DataBrowserCol.dbcVisitCycleNumber, lLoop)) <> CStr(vData(DataBrowserCol.dbcVisitCycleNumber, lSpan)))) Then Exit For
                            lCol2Span = lCol2Span + 1
                        Next

                        If bCurrent Then
                            Call AddStringToVarArr(vJSComm, "<a onmouseup=" & Chr(34))
                            Call AddStringToVarArr(vJSComm, "fnM(event.button,'" & vData(DataBrowserCol.dbcStudyId, lLoop) & gsDELIMITER1 _
                                              & vData(DataBrowserCol.dbcSite, lLoop) & gsDELIMITER1 _
                                              & vData(DataBrowserCol.dbcSubjectId, lLoop) & gsDELIMITER1 _
                                              & vData(DataBrowserCol.dbcVisitId, lLoop) & gsDELIMITER1 _
                                              & vData(DataBrowserCol.dbcVisitCycleNumber, lLoop) & "'," _
                                              & MIMsgScope.mimscVisit & "," _
                                              & LFScope.lfscVisit & "," _
                                              & "'" & ReplaceWithJSChars(vData(DataBrowserCol.dbcVisitName, lLoop)) & "'," _
                                              & "'',")
                            Call AddStringToVarArr(vJSComm, RtnMenuLockCallParams(vData(DataBrowserCol.dbcVisitLockStatus, lLoop), bLock, bUnlock, bFreeze))
                            ' Unfreeze param
                            If bUnFreeze Then
                                If vData(DataBrowserCol.dbcVisitLockStatus, lLoop) = LockStatus.lsFrozen Then
                                    ' check subject is not frozen
                                    If vData(DataBrowserCol.dbcSubjectLockStatus, lLoop) <> LockStatus.lsFrozen Then
                                        Call AddStringToVarArr(vJSComm, "1,")
                                    Else
                                        Call AddStringToVarArr(vJSComm, "0,")
                                    End If
                                Else
                                    Call AddStringToVarArr(vJSComm, "0,")
                                End If
                            Else
                                Call AddStringToVarArr(vJSComm, "0,")
                            End If
                            'ic 30/05/2003 cant goto eform from this column
                            Call AddStringToVarArr(vJSComm, "0,0,")
                            'ic 17/04/2007 issue 2837, can create sdv if locked but not frozen (to match windows)
                            If (bAddSDV And (vData(DataBrowserCol.dbcVisitLockStatus, lLoop) <> LockStatus.lsFrozen)) Then
                                Call AddStringToVarArr(vJSComm, "1,")
                            Else
                                Call AddStringToVarArr(vJSComm, "0,")
                            End If
                            Call AddStringToVarArr(vJSComm, "0")
                            Call AddStringToVarArr(vJSComm, ");" & Chr(34) & ">" & vbCrLf)
                        End If

                        Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcVisitLockStatus, lLoop)))
                        If lCol2Span > 1 Then Call AddStringToVarArr(vJSComm, " rowspan=" & lCol2Span)
                        'insert the value, and the cycle number
                        Call AddStringToVarArr(vJSComm, ">" & vData(DataBrowserCol.dbcVisitName, lLoop) & " [" & vData(DataBrowserCol.dbcVisitCycleNumber, lLoop) & "]&nbsp;" & RtnStatusImages(vData(DataBrowserCol.dbcVisitStatus, lLoop), bViewInformIcon, vData(DataBrowserCol.dbcVisitLockStatus, lLoop), False, vData(DataBrowserCol.dbcVisitSDVStatus, lLoop), vData(DataBrowserCol.dbcVisitDiscStatus, lLoop)) & "</td>")
    
                        If bCurrent Then Call AddStringToVarArr(vJSComm, "</a>")
                    End If
                    lCol2Span = lCol2Span - 1


                    'eform column
                    If lCol3Span = 0 Then
                        For lSpan = lLoop To lStop
                            If IsNull(vData(DataBrowserCol.dbcSubjectLabel, lLoop)) Then
                                sLabel1 = ""
                            Else
                                sLabel1 = vData(DataBrowserCol.dbcSubjectLabel, lLoop)
                            End If
                            If IsNull(vData(DataBrowserCol.dbcSubjectLabel, lSpan)) Then
                                sLabel2 = ""
                            Else
                                sLabel2 = vData(DataBrowserCol.dbcSubjectLabel, lSpan)
                            End If
    
                            If ((CStr(vData(DataBrowserCol.dbcStudyName, lLoop)) <> CStr(vData(DataBrowserCol.dbcStudyName, lSpan))) _
                            Or (CStr(vData(DataBrowserCol.dbcSite, lLoop)) <> CStr(vData(DataBrowserCol.dbcSite, lSpan))) _
                            Or (CStr(vData(DataBrowserCol.dbcSubjectId, lLoop)) <> CStr(vData(DataBrowserCol.dbcSubjectId, lSpan))) _
                            Or (CStr(sLabel1) <> CStr(sLabel2)) _
                            Or (CStr(vData(DataBrowserCol.dbcVisitName, lLoop)) <> CStr(vData(DataBrowserCol.dbcVisitName, lSpan))) _
                            Or (CStr(vData(DataBrowserCol.dbcVisitCycleNumber, lLoop)) <> CStr(vData(DataBrowserCol.dbcVisitCycleNumber, lSpan))) _
                            Or (CStr(vData(DataBrowserCol.dbcEFormId, lLoop)) <> CStr(vData(DataBrowserCol.dbcEFormId, lSpan))) _
                            Or (CStr(vData(DataBrowserCol.dbcEFormTitle, lLoop)) <> CStr(vData(DataBrowserCol.dbcEFormTitle, lSpan)))) Then Exit For
                            lCol3Span = lCol3Span + 1
                        Next

                        If bCurrent Then
                            Call AddStringToVarArr(vJSComm, "<a onmouseup=" & Chr(34))
                            Call AddStringToVarArr(vJSComm, "fnM(event.button,'" & vData(DataBrowserCol.dbcStudyId, lLoop) & gsDELIMITER1 _
                                              & vData(DataBrowserCol.dbcSite, lLoop) & gsDELIMITER1 _
                                              & vData(DataBrowserCol.dbcSubjectId, lLoop) & gsDELIMITER1 _
                                              & vData(DataBrowserCol.dbcVisitId, lLoop) & gsDELIMITER1 _
                                              & vData(DataBrowserCol.dbcVisitCycleNumber, lLoop) & gsDELIMITER1 _
                                              & gsDELIMITER1 _
                                              & vData(DataBrowserCol.dbcEFormId, lLoop) & gsDELIMITER1 _
                                              & vData(DataBrowserCol.dbcEFormCycleNumber, lLoop) & gsDELIMITER1 _
                                              & vData(DataBrowserCol.dbcEFormTaskID, lLoop) & "'," _
                                              & MIMsgScope.mimscEForm & "," _
                                              & LFScope.lfscEForm & "," _
                                              & "'" & ReplaceWithJSChars(vData(DataBrowserCol.dbcEFormTitle, lLoop)) & "'," _
                                              & "'',")
                            Call AddStringToVarArr(vJSComm, RtnMenuLockCallParams(vData(DataBrowserCol.dbcEFormLockStatus, lLoop), bLock, bUnlock, bFreeze))
                            ' Unfreeze param
                            If bUnFreeze Then
                                If vData(DataBrowserCol.dbcEFormLockStatus, lLoop) = LockStatus.lsFrozen Then
                                    ' check subject is not frozen
                                    If vData(DataBrowserCol.dbcSubjectLockStatus, lLoop) <> LockStatus.lsFrozen _
                                        And vData(DataBrowserCol.dbcVisitLockStatus, lLoop) <> LockStatus.lsFrozen Then
                                        Call AddStringToVarArr(vJSComm, "1,")
                                    Else
                                        Call AddStringToVarArr(vJSComm, "0,")
                                    End If
                                Else
                                    Call AddStringToVarArr(vJSComm, "0,")
                                End If
                            Else
                                Call AddStringToVarArr(vJSComm, "0,")
                            End If
                            Call AddStringToVarArr(vJSComm, "1,0,")
                            'ic 17/04/2007 issue 2837, can create sdv if locked but not frozen (to match windows)
                            If (bAddSDV And (vData(DataBrowserCol.dbcEFormLockStatus, lLoop) <> LockStatus.lsFrozen)) Then
                                Call AddStringToVarArr(vJSComm, "1,")
                            Else
                                Call AddStringToVarArr(vJSComm, "0,")
                            End If
                            Call AddStringToVarArr(vJSComm, "0")
                            Call AddStringToVarArr(vJSComm, ");" & Chr(34) & ">" & vbCrLf)
                        End If

                        Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcEFormLockStatus, lLoop)))
                        If lCol3Span > 1 Then Call AddStringToVarArr(vJSComm, " rowspan=" & lCol3Span)
                        
                        'ic 29/05/2003 display eform label if present
                        Call AddStringToVarArr(vJSComm, ">" & eFormTitleLabel(ConvertFromNull(vData(DataBrowserCol.dbcEFormTitle, lLoop), vbString), ConvertFromNull(vData(DataBrowserCol.dbcEFormLabel, lLoop), vbString), vData(DataBrowserCol.dbcEFormCycleNumber, lLoop)) & "&nbsp;" & RtnStatusImages(vData(DataBrowserCol.dbcEFormStatus, lLoop), bViewInformIcon, vData(DataBrowserCol.dbcEFormLockStatus, lLoop), False, vData(DataBrowserCol.dbcEFormSDVStatus, lLoop), vData(DataBrowserCol.dbcEFormDiscStatus, lLoop)) & "</td>")
    
                        If bCurrent Then Call AddStringToVarArr(vJSComm, "</a>")
                    End If
                    lCol3Span = lCol3Span - 1
    
    
                    'question column
                    ' DPH 03/03/2003 - Display Value calculated outside of 'current' If
                    If Not IsNull(vData(DataBrowserCol.dbcResponseValue, lLoop)) Then
                        sValue = vData(DataBrowserCol.dbcResponseValue, lLoop)
                    Else
                        sValue = ""
                    End If

                    If bCurrent Then

                        'ic 26/10/2006 issue 2830 parameter replaced questionid parameter
                        'note: the anchor closes at the end of the row
                        Call AddStringToVarArr(vJSComm, "<a onmouseup=" & Chr(34))
                        Call AddStringToVarArr(vJSComm, "fnM(event.button,'" & vData(DataBrowserCol.dbcStudyId, lLoop) & gsDELIMITER1 _
                                          & vData(DataBrowserCol.dbcSite, lLoop) & gsDELIMITER1 _
                                          & vData(DataBrowserCol.dbcSubjectId, lLoop) & gsDELIMITER1 _
                                          & vData(DataBrowserCol.dbcVisitId, lLoop) & gsDELIMITER1 _
                                          & vData(DataBrowserCol.dbcVisitCycleNumber, lLoop) & gsDELIMITER1 _
                                          & gsDELIMITER1 _
                                          & vData(DataBrowserCol.dbcEFormId, lLoop) & gsDELIMITER1 _
                                          & vData(DataBrowserCol.dbcEFormCycleNumber, lLoop) & gsDELIMITER1 _
                                          & vData(DataBrowserCol.dbcEFormTaskID, lLoop) & gsDELIMITER1 _
                                          & vData(DataBrowserCol.dbcQuestionId, lLoop) & gsDELIMITER1 _
                                          & vData(DataBrowserCol.dbcResponseCycleNumber, lLoop) & gsDELIMITER1 _
                                          & vData(DataBrowserCol.dbcResponseTaskId, lLoop) & "'," _
                                          & MIMsgScope.mimscQuestion & "," _
                                          & LFScope.lfscQuestion & "," _
                                          & "'" & ReplaceWithJSChars(vData(DataBrowserCol.dbcDataItemName, lLoop)) & "'," _
                                          & "'" & ReplaceWithJSChars(sValue) & "',")
    
                        Call AddStringToVarArr(vJSComm, RtnMenuLockCallParams(vData(DataBrowserCol.dbcDataItemLockStatus, lLoop), bLock, bUnlock, bFreeze))
                        ' Unfreeze param
                        If bUnFreeze Then
                            If vData(DataBrowserCol.dbcDataItemLockStatus, lLoop) = LockStatus.lsFrozen Then
                                ' check subject/visit/eform is not frozen
                                If vData(DataBrowserCol.dbcSubjectLockStatus, lLoop) <> LockStatus.lsFrozen _
                                    And vData(DataBrowserCol.dbcVisitLockStatus, lLoop) <> LockStatus.lsFrozen _
                                    And vData(DataBrowserCol.dbcEFormLockStatus, lLoop) <> LockStatus.lsFrozen Then
                                    Call AddStringToVarArr(vJSComm, "1,")
                                Else
                                    Call AddStringToVarArr(vJSComm, "0,")
                                End If
                            Else
                                Call AddStringToVarArr(vJSComm, "0,")
                            End If
                        Else
                            Call AddStringToVarArr(vJSComm, "0,")
                        End If
                        Call AddStringToVarArr(vJSComm, "1,")

                        'user must have permission to add discrepancies and record must not be locked or frozen
                        ' DPH 01/07/2003 bug 1891 Allow discrepancies,SDVs when question locked
                        ' And (vData(DataBrowserCol.dbcDataItemLockStatus, lLoop) <> LockStatus.lsLocked)
                        If (bAddDiscrepancy And (vData(DataBrowserCol.dbcDataItemLockStatus, lLoop) <> LockStatus.lsFrozen)) Then
                            Call AddStringToVarArr(vJSComm, "1,")
    
                        Else
                            Call AddStringToVarArr(vJSComm, "0,")
    
                        End If
                        If (bAddSDV And (vData(DataBrowserCol.dbcDataItemLockStatus, lLoop) <> LockStatus.lsFrozen)) Then
                            Call AddStringToVarArr(vJSComm, "1,")
    
                        Else
                            Call AddStringToVarArr(vJSComm, "0,")
    
                        End If
                        ' DPH 01/07/2003 bug 1891 disallow notes when question frozen
                        If (bAddNote And (vData(DataBrowserCol.dbcDataItemLockStatus, lLoop) <> LockStatus.lsFrozen)) Then
                            Call AddStringToVarArr(vJSComm, "1")
    
                        Else
                            Call AddStringToVarArr(vJSComm, "0")
    
                        End If
                        Call AddStringToVarArr(vJSComm, ");" & Chr(34) & ">" & vbCrLf)
                    End If

                    
                    If CInt(vData(DataBrowserCol.dbcResponseCycleNumber, lLoop)) > 1 Then
                        Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcDataItemLockStatus, lLoop)) & ">" & vData(DataBrowserCol.dbcDataItemName, lLoop) & "[" & vData(DataBrowserCol.dbcResponseCycleNumber, lLoop) & "]" & "</td>" & vbCrLf)
                    Else
                        ' Just question name
                        Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcDataItemLockStatus, lLoop)) & ">" & vData(DataBrowserCol.dbcDataItemName, lLoop) & "</td>" & vbCrLf)
                    End If
    
    
                    'value column
                    Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcDataItemLockStatus, lLoop)) & ">" & vbCrLf)
                    If (vData(DataBrowserCol.dbcDataType, lLoop) = DataType.Multimedia) Then
                        If (sValue <> "") Then Call AddStringToVarArr(vJSComm, "(attached)")
                    Else
                        'ic 26/06/2003 convert to local format
                        Call AddStringToVarArr(vJSComm, ReplaceWithHTMLCodes(LocaliseValue(sValue, CInt(vData(DataBrowserCol.dbcDataType, lLoop)), sDecimalPoint, sThousandSeparator)) & "&nbsp;</td>" & vbCrLf)
                    End If
    
                    'status column
                    ' NCJ 26 Mar 07 - Added bViewInformicon to RtnStatusImages (changed from False)
                    'Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcDataItemLockStatus, lLoop)) & ">" & RtnStatusImages(vData(DataBrowserCol.dbcResponseStatus, lLoop)) & "</td>" & vbCrLf)
                    Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcDataItemLockStatus, lLoop)) & ">" _
                    & RtnNRCTC(vData(DataBrowserCol.dbcResponseStatus, lLoop), vData(DataBrowserCol.dbcLabResult, lLoop), vData(DataBrowserCol.dbcCTCGrade, lLoop)) & RtnStatusImages(vData(DataBrowserCol.dbcResponseStatus, lLoop), bViewInformIcon, _
                    vData(DataBrowserCol.dbcDataItemLockStatus, lLoop), False, vData(DataBrowserCol.dbcDataItemSDVStatus, lLoop), _
                    vData(DataBrowserCol.dbcDataItemDiscStatus, lLoop), _
                    RtnJSBoolean(CBool(vData(DataBrowserCol.dbcDataItemNoteStatus, lLoop))), RtnJSBoolean(CBool(ConvertFromNull(vData(DataBrowserCol.dbcComments, lLoop), vbString) <> "")), _
                    vData(DataBrowserCol.dbcChangeCount, lLoop), sStatusLabel) & "</td>" & vbCrLf)
    
                    'date & time column
                    ' DPH 18/02/2003 - Local format date
                    ' NCJ 8 Dec 05 - New Date/Time types
                    Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcDataItemLockStatus, lLoop)) & ">" & GetLocalFormatDate(oUser, CDate(vData(DataBrowserCol.dbcResponseTimeStamp, lLoop)), eDateTimeType.dttDMYT) & "&nbsp;")
                    Call AddStringToVarArr(vJSComm, RtnDifferenceFromGMT(vData(DataBrowserCol.dbcResponseTimestamp_TZ, lLoop)))
                    Call AddStringToVarArr(vJSComm, "</td>")
                    
                    'database data and time
                    ' NCJ 8 Dec 05 - New Date/Time types
                    Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcDataItemLockStatus, lLoop)) & ">" & GetLocalFormatDate(oUser, CDate(vData(DataBrowserCol.dbcDatabaseTimeStamp, lLoop)), eDateTimeType.dttDMYT) & "&nbsp;")
                    Call AddStringToVarArr(vJSComm, RtnDifferenceFromGMT(vData(DataBrowserCol.dbcDatabaseTimestamp_TZ, lLoop)))
                    Call AddStringToVarArr(vJSComm, "</td>")
                    
                    'user id column
                    Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcDataItemLockStatus, lLoop)) & ">" & vData(DataBrowserCol.dbcUserName, lLoop) & "&nbsp;</td>" & vbCrLf)
    
                    'user full name column
                    Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcDataItemLockStatus, lLoop)) & ">" & vData(DataBrowserCol.dbcFullUserName, lLoop) & "&nbsp;</td>" & vbCrLf)
    
                    'ic 29/05/2003 only display comments if user has permission, bug 1816
                    'comment column
                    Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcDataItemLockStatus, lLoop)) & ">" & vbCrLf)
                    If (bViewComments) Then
                        Call AddStringToVarArr(vJSComm, ConvertFromNull(vData(DataBrowserCol.dbcComments, lLoop), vbString))
                    End If
                    Call AddStringToVarArr(vJSComm, "&nbsp;</td>")
    
                    'rfc column
                    Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcDataItemLockStatus, lLoop)) & ">" & vData(DataBrowserCol.dbcReasonForChange, lLoop) & "&nbsp;</td>" & vbCrLf)
    
                    'overrule reason column
                    Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcDataItemLockStatus, lLoop)) & ">" & vData(DataBrowserCol.dbcOverruleReason, lLoop) & "&nbsp;</td>" & vbCrLf)
    
                    'warning message column
                    Call AddStringToVarArr(vJSComm, "<td valign='top'" & RtnLockColour(vData(DataBrowserCol.dbcDataItemLockStatus, lLoop)) & ">" & vData(DataBrowserCol.dbcValMessage, lLoop) & "&nbsp;</td>" & vbCrLf)
    
                    If bCurrent Then
                        Call AddStringToVarArr(vJSComm, "</a>")
                    End If
    
                    Call AddStringToVarArr(vJSComm, "</tr>" & vbCrLf)
                Next
                
                ' DPH 21/05/2003 - removed bottom row
                ' "<tr height='150'><td colspan='4'>&nbsp;</td></tr>"
                Call AddStringToVarArr(vJSComm, "</table>")
            End If
        Else
        
            'no records returned
            Call AddStringToVarArr(vJSComm, "<div class='clsMessageText'>Your query returned no records</div>")
        End If
    Else
    
        'too many records returned
        Call AddStringToVarArr(vJSComm, "<div class='clsMessageText'>Your query returned too many records. Please refine your search</div>")
    End If
    
    
    Call AddStringToVarArr(vJSComm, "</body>")
    GetDataBrowser = Join(vJSComm, "")
    Exit Function

CatchAllError:
    Call Err.Raise(Err.Number, , Err.Description & "|modUIHTMLDataBrowser.GetDataBrowser")
End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnDataBrowser(ByRef oUser As MACROUser, _
                               ByVal bCountOnly As Boolean, _
                               ByVal nStudyCode As Integer, _
                               ByVal sSiteCode As String, _
                               ByVal lVisitCode As Long, _
                               ByVal lCRFPageId As Long, _
                               ByVal lQuestion As Long, _
                               ByVal sSrchUserName As String, _
                               ByVal lSubjectId As Long, _
                               ByVal sSubjectLabel As String, _
                               ByVal sStatus As String, _
                               ByVal sLockStatus As String, _
                               ByVal sTime As String, _
                               ByVal bBefore As Boolean, _
                               ByVal lComment As Long, _
                               ByVal lDiscrepancy As Long, _
                               ByVal lSDV As Long, _
                               ByVal lNote As Long, _
                               ByVal lCodingStatus As Long, _
                               ByVal sDictionaryName As String, _
                               ByVal sDictionaryVersion As String, _
                               ByVal sGet As String, _
                      Optional ByVal bConvertDates As Boolean = True, _
                      Optional ByRef vErrors As Variant) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 08/08/01
'   function returns a serialised recordset containing the data returned by a query
' DPH 08/11/2002 Changed to use Serialised User object
' ic 23/01/2003 moved from clswww
' DPH 08/10/2003 Performance changes - use DataItemResponse/DataItemResponseHistory table
'   ic 29/06/2004 added error handling
'   ic 28/07/2005 added clinical coding
'--------------------------------------------------------------------------------------------------
Dim vRtn As Variant
Dim oDBrowser As DataBrowser
Dim tDBrowser As eDataBrowserType
Dim nLoop As Integer
Dim sStudySiteSQL As String
Dim bSingleSubject As Boolean
Dim bDateOK As Boolean
Dim dblDate As Double
                  
    On Error GoTo CatchAllError
       
    If (sGet <> "3") Then
        bSingleSubject = Not oUser.CheckPermission(gsFnMonitorDataReviewData)
        ' DPH 08/10/2003 - Get relevant study site SQL
        Select Case sGet
        Case "0":
            'forms TA - not used
            sStudySiteSQL = oUser.DataLists.StudiesSitesWhereSQL("ClinicalTrial.ClinicalTrialId", "TrialSubject.TrialSite")
        Case "1":
            'current
            sStudySiteSQL = oUser.DataLists.StudiesSitesWhereSQL(CStr(nStudyCode), "DataItemResponse.TrialSite")
        Case "2":
            'audit
            sStudySiteSQL = oUser.DataLists.StudiesSitesWhereSQL(CStr(nStudyCode), "DataItemResponseHistory.TrialSite")
        End Select
        dblDate = RtnRecordDblDate(sTime, bDateOK)
        If Not bDateOK And Not IsMissing(vErrors) Then
            vErrors = AddToArray(vErrors, "Search date", "Unable to search on passed format")
        End If
    End If

    Set oDBrowser = New DataBrowser
    
    Select Case sGet
    Case "0":
        'forms
        vRtn = oDBrowser.GetData(oUser.CurrentDBConString, bCountOnly, dbteForms, sStudySiteSQL, bSingleSubject, _
                                 nStudyCode, sSiteCode, lVisitCode, lCRFPageId, sSubjectLabel, lSubjectId, _
                                 RtnRecordStatusString(sStatus, oUser.CheckPermission(gsFnMonitorDataReviewData)), RtnRecordLockString(sLockStatus), bBefore, _
                                 dblDate, lQuestion, sSrchUserName, lComment, lDiscrepancy, lSDV, lNote)
    Case "1":
        'current
        'ic 28/07/2005 added clinical coding
        vRtn = oDBrowser.GetData(oUser.CurrentDBConString, bCountOnly, dbtDataItemResponse, sStudySiteSQL, _
                                 bSingleSubject, nStudyCode, sSiteCode, lVisitCode, lCRFPageId, _
                                 sSubjectLabel, lSubjectId, RtnRecordStatusString(sStatus, oUser.CheckPermission(gsFnMonitorDataReviewData)), _
                                 RtnRecordLockString(sLockStatus), bBefore, dblDate, _
                                 lQuestion, sSrchUserName, lComment, lDiscrepancy, lSDV, lNote, lCodingStatus, sDictionaryName, sDictionaryVersion)
    Case "2":
        'audit
        'ic 28/07/2005 added clinical coding
        vRtn = oDBrowser.GetData(oUser.CurrentDBConString, bCountOnly, dbtDataItemResponseHistory, sStudySiteSQL, _
                                 bSingleSubject, nStudyCode, sSiteCode, lVisitCode, lCRFPageId, _
                                 sSubjectLabel, lSubjectId, RtnRecordStatusString(sStatus, oUser.CheckPermission(gsFnMonitorDataReviewData)), _
                                 RtnRecordLockString(sLockStatus), bBefore, dblDate, _
                                 lQuestion, sSrchUserName, lComment, lDiscrepancy, lSDV, lNote, lCodingStatus, sDictionaryName, sDictionaryVersion)
    Case "3":
        'changes since last session
        vRtn = oDBrowser.GetData(oUser.CurrentDBConString, bCountOnly, dbtDataItemResponse, _
            oUser.DataLists.StudiesSitesWhereSQL("DATAITEMRESPONSE.CLINICALTRIALID", "DATAITEMRESPONSE.TRIALSITE"), _
            False, , , , , , , , , , oUser.LastLogin - 1)
    End Select
    
    'TA 21/10/2002: windows does not want dates converted
    If bConvertDates And Not bCountOnly Then
        'loop through all records converting the date - the vbscript cdate function always returns
        'in american format rather than according to regional settings
        If Not IsNull(vRtn) Then
            For nLoop = LBound(vRtn, 2) To UBound(vRtn, 2)
                If vRtn(dbcResponseTimeStamp, nLoop) <> 0 Then
                    vRtn(dbcResponseTimeStamp, nLoop) = CStr(CDate(vRtn(dbcResponseTimeStamp, nLoop)))
                End If
            Next
        End If
    End If
    
    Set oDBrowser = Nothing
    RtnDataBrowser = vRtn
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLDataBrowser.RtnDataBrowser"
End Function
'--------------------------------------------------------------------------------------------------
Private Function RtnRecordStatusString(sStatus As String, ByVal bMonitor As Boolean) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 08/08/01
'   function returns an array representing the statuses requested in a passed binary string
'   revisions
'   ic 23/01/2003 moved from clswww
'   ic 19/02/2003 added bMonitor arg
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim vRtn As Variant
Dim ndx As Integer
    
    On Error GoTo CatchAllError
    ndx = 0
    
    sStatus = Trim(sStatus)
    If Len(sStatus) <> 7 Or (sStatus = "1111111") Or (sStatus = "0000000") Then
         vRtn = Null
    Else
        ReDim vRtn(6)
        If Mid(sStatus, 1, 1) <> "0" Then
            vRtn(ndx) = Status.Success
            ndx = ndx + 1
        End If
        If Mid(sStatus, 2, 1) <> "0" Then
            vRtn(ndx) = Status.Missing
            ndx = ndx + 1
        End If
        If Mid(sStatus, 3, 1) <> "0" Then
            vRtn(ndx) = Status.NotApplicable
            ndx = ndx + 1
        End If
        If Mid(sStatus, 4, 1) <> "0" Then
            vRtn(ndx) = Status.Warning
            ndx = ndx + 1
        End If
        If (bMonitor) Then
            If Mid(sStatus, 5, 1) <> "0" Then
                vRtn(ndx) = Status.Inform
                ndx = ndx + 1
            End If
        Else
            'if no monitor permission, just match the 'ok' search parameter
            If Mid(sStatus, 1, 1) <> "0" Then
                vRtn(ndx) = Status.Inform
                ndx = ndx + 1
            End If
        End If
        If Mid(sStatus, 6, 1) <> "0" Then
            vRtn(ndx) = Status.OKWarning
            ndx = ndx + 1
        End If
        If Mid(sStatus, 7, 1) <> "0" Then
            vRtn(ndx) = Status.Unobtainable
            ndx = ndx + 1
        End If
        
        'ic 30/04/2002
        'changed to ndx - 1
        ReDim Preserve vRtn(ndx - 1)
    End If

    RtnRecordStatusString = vRtn
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLDataBrowser.RtnRecordStatusString"
End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnRecordLockString(ByVal sStatus As String) As Variant
'--------------------------------------------------------------------------------------------------
'   ic 08/08/01
'   function returns an array representing the lockstatus in a passed binary string
'   revisions
'   ic 23/01/2003 moved from clswww
'   ic 29/06/2004 added error handling
'--------------------------------------------------------------------------------------------------
Dim vRtn As Variant
Dim ndx As Integer
    
    On Error GoTo CatchAllError
    ndx = 0
    
    sStatus = Trim(sStatus)
    If Len(sStatus) <> 2 Or (sStatus = "00") Then
         vRtn = Null
    Else
        ReDim vRtn(2)
        
        If Mid(sStatus, 1, 1) <> "0" Then
            vRtn(ndx) = eLockStatus.lsLocked
            ndx = ndx + 1
        End If
        If Mid(sStatus, 2, 1) <> "0" Then
            vRtn(ndx) = eLockStatus.lsFrozen
            ndx = ndx + 1
        End If
        
        ReDim Preserve vRtn(ndx - 1)
    End If

    RtnRecordLockString = vRtn
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLDataBrowser.RtnRecordLockString"
End Function
