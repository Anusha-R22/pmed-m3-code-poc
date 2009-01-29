Attribute VB_Name = "modUIHTMLSchedule"
'----------------------------------------------------------------------------------------'
'   File:       modUIHTML.bas
'   Copyright:  InferMed Ltd. 2000 - 2006. All Rights Reserved
'   Author:     i curtis 02/2003
'   Purpose:    functions returning html versions of MACRO pages (MIMESSAGES)
'----------------------------------------------------------------------------------------'
'   revisions
'   ic 29/05/2003 removed 'onmouseout' event for popup menu in GetSchedule(), bug 1761
'   ic 05/06/2003 added extra 'refresh z-order' parameter in GetSchedule()
'   ic 19/08/2003 added registration
'   ic 02/09/2003 added new CanChangeData function to various functions
'   ic 30/09/2003 changed to pass whole image tag, otherwise IE downloads same image many times
'   ic 02/10/2003 added alternative GetScheduleNoCompression(). (unchanged version is below)
'   ic 16/10/2003 check quicklist permission
'   ic 01/03/2004 GetSchedule() & GetScheduleNoCompression() added 'can add sdv argument
'   ic 29/06/2004 added error handling
'   ic 24/08/2004 added 'set planned SDVs to done' functionality
'   ic 06/09/2004 added not applicable eform icon to schedule, bug 2389
'   ic 08/09/2004 removed 'can add sdv' argument - adding an sdv has to be tried
'   ic 01/04/2005 issue 2541 added sSDVCall parameter to GetSchedule functions
'   ic 29/06/2005 issue 2464, enhancements to sdv work flow, added visit, eform, question cycle
'   NCJ 21 Jun 06 - Bug 2747 - Concatenate Lock, Discrepancy, SDV and eForm statuses into eForm tooltip texts
'----------------------------------------------------------------------------------------'

Option Explicit
'MLM 06/02/03: String constants used to build HTML comments containing schedule's data.
'These comments are expanded into markup by javascript.
'Delimiters for nested data must not use HTML comment close tag.
Const msDELIMITER = "|"
Const msIMAGE_START = "!F"
Const msIMAGE_PLANNED_SDV_START = "!SP"
Const msIMAGE_END = "F->"
Const msTABLE_START = "!T"
Const msTABLE_END = "T->"
Const msCELL_START = "<!--C"
Const msCELL_END = "-->"
Const msEFORM_HEADER_START As String = "<!--E"
Const msBLANK_EFORM_HEADER As String = "<!--B-->"
Const msEFORM_HEADER_END As String = "-->"
Const msVISIT_START As String = "<!--V"
Const msVISIT_END As String = "-->"
Const msVISIT_TITLE_START As String = "<!--A"
Const msVISIT_TITLE_END As String = "-->"
Const msV As String = "<!--D-->"

'ic 17/02/2003 constants for displaying new icons
Const msIMAGE_QUERIED_SDV_START = "!SQ"
Const msIMAGE_DONE_SDV_START = "!SD"

'ic 01/04/2003 temporary showcode variable
Private Const mbShowCode As Boolean = False

'--------------------------------------------------------------------------------------------------
Public Function GetScheduleHTML(ByRef oUser As MACROUser, ByVal sSiteCode As String, ByVal lStudyCode As Long, _
    ByVal lSubjectId As Long, ByVal vErrors As Variant, ByVal vAlerts As Variant, ByVal bNew As Boolean, _
    Optional sSDVCall As String = "") As String
'--------------------------------------------------------------------------------------------------
'   ic 29/09/02
'   gets html table schedule
' revisions
' DPH 08/11/2002 Changed to use Serialised User object
' ic 16/01/2003 added error message for subjects left locked due to crashes
' 09/10/2003 added 'no compression' alternative
' ic 25/03/2004 added oStudydef.Terminate call to free up memory
' ic 29/06/2004 added error handling, moved from clswww
' ic 24/08/2004 added vAlerts parameter for changesdvstodone message functionality
' ic 01/04/2005 issue 2541 added sSDVCall parameter
'--------------------------------------------------------------------------------------------------
Dim oStudyDef As StudyDefRO
Dim oSubject As StudySubject
Dim bUseCompression As Boolean

    On Error GoTo CatchAllError

    bUseCompression = RtnCompressionFlag()

    'load the subject
    Set oStudyDef = New StudyDefRO
    Call oStudyDef.Load(oUser.CurrentDBConString, lStudyCode, 1)
    Set oSubject = oStudyDef.LoadSubject(sSiteCode, lSubjectId, oUser.UserName, Read_Only, oUser.UserNameFull, oUser.UserRole, False)

    If (oSubject.CouldNotLoad) Then vErrors = AddToArray(vErrors, "clsWWW.GetScheduleHTML", oSubject.CouldNotLoadReason)
    
    If bUseCompression Then
        'return schedule with compression
        GetScheduleHTML = GetSchedule(oSubject, iwww, oUser, vErrors, vAlerts, bNew, sSDVCall)
    Else
        'return schedule without compression
        GetScheduleHTML = GetScheduleNoCompression(oSubject, iwww, oUser, vErrors, vAlerts, bNew, sSDVCall)
    End If
    
    'ic 25/03/2004 MUST call the terminate method to free up memory
    Call oStudyDef.Terminate
    Set oStudyDef = Nothing
    Set oSubject = Nothing
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLSchedule.GetScheduleHTML"
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetScheduleNoCompression(ByRef oSubject As StudySubject, Optional ByVal enInterface As eInterface = iwww, _
    Optional ByRef oUser As MACROUser, Optional ByVal vErrors As Variant, Optional ByVal vAlerts As Variant, _
    Optional ByVal bNew As Boolean, Optional sSDVCall As String = "") As String

'--------------------------------------------------------------------------------------------------
'   ic 11/07/01
'   builds and returns a string representing a subject schedule in an html table
'
'   ***********************************************************************************************
'   NOTE: IF YOU MAKE CHANGES TO THIS FUNCTION, MAKE CORRESPONDING CHANGES TO COMPRESSION
'   FUNCTION BELOW
'   ***********************************************************************************************
'
'   REVISIONS
'   DPH 19/12/2002 - Show EITHER subject label OR (subjectid) not both
'   ic 13/01/2003   display visit name if visit inactive
'   DPH 16/01/2003 - Lock/Unlock/Freeze/UnFreeze functionality for subject/visit/eForm
'   DPH 23/01/2003 - Split out <head> for WWW as showing "please wait..."
'   ic 28/01/2003   fixed subject label/icon split line
'   ic 04/02/2003   fixed hand mousepointer over subject/visit
'   DPH 18/02/2003  uses GetLocalFormatDate format for Visit / Eform dates
'   ic 01/04/2003   added flag for 'display code or name'
'   ic 29/05/2003 removed 'onmouseout' event for popup menu, bug 1761
'   ic 05/06/2003 added extra 'refresh z-order' parameter
'   ic 19/06/2003 added registration
'   ic 30/09/2003 changed to pass whole image tag, otherwise IE downloads same image many times
'   ic 01/10/2003 moved fnHideLoader() inside fnPageLoaded()
'   ic 02/10/2003 amended to use no compression. (unchanged version is below)
'   ic 16/10/2003 check quicklist permission
'   ic 01/03/2004 added 'can add sdv argument
'   ic 29/06/2004 added error handling
'   ic 24/08/2004 added 'can set sdvs to done' fnE() parameter, vAlerts parameter
'   ic 08/09/2004 removed 'can add sdv' argument - adding an sdv has to be tried
'   ic 01/04/2005 issue 2541 added sSDVCall parameter
'   ic 29/06/2005 issue 2464, enhancements to sdv work flow, added visit, eform, question cycle
'--------------------------------------------------------------------------------------------------
Dim oS As ScheduleGrid
Dim oCell As GridCell
Dim nCol As Integer
Dim nRow As Integer
Dim nEfId As Long
Dim sCycle As String
Dim nLoop As Integer
Dim bLocked As Boolean
Dim bFrozen As Boolean
Dim bCanUnFreeze As Boolean
Dim eLockFreezeStatus As eLockStatus
Dim bViewInformIcon As Boolean
Dim vJSComm() As String
Dim bShowVCode As Boolean
Dim bShowECode As Boolean
Dim lRaised As Long
Dim lResponded As Long
Dim lPlanned As Long
Dim bQuickList As Boolean
Dim oMIMData As MIDataLists
Dim bEHasSDV As Boolean
Dim bVHasSDV() As Boolean
Dim bSHasSDV As Boolean
Dim lObjectId As Long
Dim nObjectSource As Integer

    On Error GoTo CatchAllError
    ReDim vJSComm(0)

    Set oS = oSubject.ScheduleGrid

    bViewInformIcon = oUser.CheckPermission(gsFnMonitorDataReviewData)
    bQuickList = oUser.CheckPermission(gsFnViewQuickList)
    
    'show code header rather than full name
    bShowVCode = mbShowCode
    bShowECode = mbShowCode

    Set oMIMData = New MIDataLists

    If (enInterface = iwww) Then
        ' DPH 23/01/2003
        'sHTML = sHTML & "<head>" _
        '              & "<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>"
        'sHTML = sHTML & "<script language='javascript' src='../script/Schedule.js'></script>"

        'ic 29/05/2003 removed 'onmouseout' event for popup menu, bug 1761
        Call AddStringToVarArr(vJSComm, "<body class=clsScheduleBorder onload='fnPageLoaded();' onscroll='PositionHeaders();fnPopMenuHide();'>" & vbCrLf _
                      & "<div class='clsPopMenu' id='divPopMenu' onclick=" & Chr(34) & "menu=this;this.tid=setTimeout('menu.style.visibility=\'hidden\'',20);" & Chr(34) _
                      & "onmouseover='clearTimeout(this.tid);'>" _
                      & "</div>" & vbCrLf)

        'ic 01/10/2003 moved fnHideLoader() inside fnPageLoaded()
        ' moved JavaScript code previously in <head> to here
        'DPH 16/01/2003 - Added new user permissions required
        'ic 02/09/2003 use new CanChangeData function
        Call AddStringToVarArr(vJSComm, "<script language='javascript'>" & vbCrLf _
                      & "function fnPageLoaded()" & vbCrLf _
                        & "{" & vbCrLf _
                          & "fnHideLoader();" & vbCrLf _
                          & "fnInitUser(" & RtnJSBoolean(oUser.CheckPermission(gsFnViewData)) & "," _
                                          & RtnJSBoolean(CanChangeData(oUser, oSubject.Site)) & "," _
                                          & RtnJSBoolean(oUser.CheckPermission(gsFnCreateSDV)) & "," _
                                          & RtnJSBoolean(oUser.CheckPermission(gsFnLockData)) & "," _
                                          & RtnJSBoolean(oUser.CheckPermission(gsFnFreezeData)) & "," _
                                          & RtnJSBoolean(oUser.CheckPermission(gsFnUnFreezeData)) & ")" & vbCrLf _
                          & "fnInitSubject('" & oSubject.StudyId & "'," _
                                        & "'" & oSubject.Site & "'," _
                                        & "'" & oSubject.PersonId & "'," _
                                        & "'" & oSubject.Label & "')" & vbCrLf _
                          & "window.sWinState='4" & gsDELIMITER2 & oSubject.StudyId & gsDELIMITER2 & oSubject.Site & gsDELIMITER2 & oSubject.PersonId & "';" & vbCrLf)

        'ic 16/10/2003 check quicklist permission
        If (bNew And bQuickList) Then
            'studyid|studyname|site|subjectid|subjectlabel
            Call AddStringToVarArr(vJSComm, "window.parent.fnAddToQuickList('" & ReplaceWithJSChars(oSubject.StudyId & gsDELIMITER2 _
                & oSubject.StudyDef.Name & gsDELIMITER2 & oSubject.Site & gsDELIMITER2 & oSubject.PersonId & gsDELIMITER2 _
                & oSubject.Label) & "');" & vbCrLf)
        End If

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

        'any additional alerts to display
        If Not IsMissing(vAlerts) Then
            If Not IsEmpty(vAlerts) Then
                For nLoop = LBound(vAlerts, 2) To UBound(vAlerts, 2)
                    Call AddStringToVarArr(vJSComm, "alert('" & ReplaceWithJSChars(vAlerts(1, nLoop)) & "');" & vbCrLf)
                Next
            End If
        End If
        
        'MLM 05/02/03: Call the javascript that will expand the page into nested tables
'        Call AddStringToVarArr(vJSComm, "document.all('outertable').innerHTML=fnExpand(document.all('outertable').innerHTML);" & vbCrLf)
        'Call AddStringToVarArr(vJSComm, "document.all('debug').innerText=document.all('outertable').innerHTML;")
        'MLM 04/03/03: Call JavaScript to set up floating schedule headers
        Call AddStringToVarArr(vJSComm, "DrawHeaders();" & vbCrLf)


        'TA 30/11/2003: moved here so not run in windows
        Call RtnMIMsgStatusCount(oUser, lRaised, lResponded, lPlanned)
        'ic 05/06/2003 added extra 'refresh z-order' parameter
        Call AddStringToVarArr(vJSComm, "window.parent.fnSTLC('" & gsVIEW_RAISED_DISCREPANCIES_MENUID & "','" & CStr(lRaised) & "',0);" & vbCrLf _
                                      & "window.parent.fnSTLC('" & gsVIEW_RESPONDED_DISCREPANCIES_MENUID & "','" & CStr(lResponded) & "',0);" & vbCrLf _
                                      & "window.parent.fnSTLC('" & gsVIEW_PLANNED_SDV_MARKS_MENUID & "','" & CStr(lPlanned) & "',1);" & vbCrLf)

        'ic 19/08/2003 enable/disable registration menu item
        'ic 02/09/2003 use new CanChangeData function
        If (CanChangeData(oUser, oSubject.Site) And oUser.CheckPermission(gsFnRegisterSubject) And ShouldEnableRegistrationMenu(oSubject)) Then
            Call AddStringToVarArr(vJSComm, "window.parent.fnEnableRegister(" & RtnJSBoolean(True) & ");" & vbCrLf)
        Else
            Call AddStringToVarArr(vJSComm, "window.parent.fnEnableRegister(" & RtnJSBoolean(False) & ");" & vbCrLf)
        End If
        
        'write javascript sdv call (if there is one)
        Call AddStringToVarArr(vJSComm, sSDVCall & vbCrLf)
        
        Call AddStringToVarArr(vJSComm, "}var bShowSDVScheduleMenu=" & RtnJSBoolean(RtnShowSDVScheduleMenuFlag()) & ";" & vbCrLf)
        
        Call AddStringToVarArr(vJSComm, "</script>" & vbCrLf)

        Call AddStringToVarArr(vJSComm, "<form name='Form1' action='Schedule.asp?fltSi=" & oSubject.Site & "&fltSt=" & oSubject.StudyId & "&fltSj=" & oSubject.PersonId & "' method='post'>" _
                      & "<input type='hidden' name='SchedUpdate'>" _
                      & "<input type='hidden' name='SchedIdentifier'>" _
                      & "</form>" & vbCrLf)

    Else
        Call AddStringToVarArr(vJSComm, "<html>")
        Call AddStringToVarArr(vJSComm, "<head>")
        Call AddStringToVarArr(vJSComm, "<script language='javascript'>" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnE(id,btn){window.navigate('VBfnEformUrl|'+id+'|'+btn);}" _
            & "function fnPrint(){window.navigate('VBfnPrintAll');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnClose(){window.navigate('VBfnCloseSchedule');}" & vbCrLf)

        Call AddStringToVarArr(vJSComm, "function fnV(id,btn){window.navigate('VBfnV|'+id+'|'+btn);}" _
            & "function fnS(btn){window.navigate('VBfnS|'+btn);}" & vbCrLf)
        'MLM 12/02/03: Added fnPageLoaded to Windows version
        Call AddStringToVarArr(vJSComm, "function fnPageLoaded(){")
        'Call AddStringToVarArr(vJSComm, "document.all('outertable').innerHTML=fnExpand(document.all('outertable').innerHTML);" & vbCrLf)
        'MLM 04/03/03: Call JavaScript to set up floating schedule headers
        Call AddStringToVarArr(vJSComm, "DrawHeaders();" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "}" & vbCrLf)

        Call AddStringToVarArr(vJSComm, "</script>" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "</head><body class=clsScheduleBorder onscroll='PositionHeaders();'>" & vbCrLf)
    End If




    'outer table, first cell holds heading table
    'MLM 05/02/03: Added span.
    Call AddStringToVarArr(vJSComm, "<div><br><br></div><table id=top height=38 width='100%' border='0' cellpadding='0' cellspacing='0'>" _
                    & "<tr height='20'>" _
                      & "<td colspan='" & oS.ColMax + 1 & "'>")

        ' DPH 19/12/2002 - use sLabelOrId
        ' ic 30/01/2003 switched to RtnSubjectText()
        'heading/top menu table
        Call AddStringToVarArr(vJSComm, "<table width='100%' height='100%' cellpadding='0' cellspacing='0'>" _
                        & "<tr class='clsScheduleBorder'>" _
                          & "<td>" _
                            & "<table>" _
                              & "<tr>" _
                                & "<td align='left'>")

        If (enInterface = iwww) Then
            ' calculate lock/freeze info for subject
            ' See if we have a locked or frozen item
            bLocked = False
            bFrozen = False
            eLockFreezeStatus = oSubject.LockStatus
            Select Case eLockFreezeStatus
                Case eLockStatus.lsLocked
                    bLocked = True
                Case eLockStatus.lsFrozen
                    bFrozen = True
                Case Else
            End Select
            ' (User's Unfreeze permission is checked later)
            If bFrozen Then
                bCanUnFreeze = True
            Else
                ' Can't unfreeze if not frozen!
                bCanUnFreeze = False
            End If
            
            'does this subject have an sdv
            bSHasSDV = oMIMData.MIMessageExists(oUser.CurrentDBConString, mimtSDVMark, mimscSubject, _
            oSubject.StudyDef.Name, oSubject.Site, oSubject.PersonId, lObjectId, nObjectSource)
            
            'ic 08/09/2004 removed 'can add sdv' argument - adding an sdv has to be tried
            'ic 01/03/2004 added 'can add sdv argument
            Call AddStringToVarArr(vJSComm, "<a onMouseup='javascript:fnS(event.button," & Chr(34) _
            & oSubject.Status & Chr(34) & "," & RtnJSBoolean(bLocked) & "," & RtnJSBoolean(bFrozen) & "," _
            & RtnJSBoolean(bCanUnFreeze) & "," & RtnJSBoolean(bSHasSDV) & ")'>")
        Else
            Call AddStringToVarArr(vJSComm, "<a onMouseup='javascript:fnS(event.button)'>")
        End If

        'Call AddStringToVarArr(vJSComm, "<div  style='cursor:hand;' class='clsScheduleEformBorder'><table cellpadding='0' cellspacing='0'><tr class='clsScheduleHeadingText'><td>" & oSubject.StudyCode & "/" & oSubject.Site & "/" & RtnSubjectText(oSubject.PersonID, oSubject.Label) & "&nbsp;</td><td>" & RtnStatusImage(oSubject.Status, False, oSubject.LockStatus, "", "", oSubject.SDVStatus) & "</td></tr></table></div>"
        Call AddStringToVarArr(vJSComm, "<div  style='cursor:default;' class='clsScheduleEformBorder'><table cellpadding='0' cellspacing='0'><tr class='clsScheduleHeadingText'><td>" & oSubject.StudyCode & "/" & oSubject.Site & "/" & RtnSubjectText(oSubject.PersonId, oSubject.Label) & "&nbsp;</td><td>" & RtnStatusImages(oSubject.Status, bViewInformIcon, oSubject.LockStatus, False, oSubject.SDVStatus, oSubject.DiscrepancyStatus) & "</td></tr></table></div>" _
                                & "</a>" _
                                & "</td>" _
                              & "</tr>" _
                            & "</table>" _
                          & "</td>")

        Call AddStringToVarArr(vJSComm, "<td align='center' class='clsScheduleMenuLinkText'>")

        'TA for windows add print all eforms link
        ' NCJ 30 Jun 03 - Changed eforms to eForms
        If (enInterface = eInterface.iWindows) Then
            Call AddStringToVarArr(vJSComm, "<a href='javascript:fnPrint();'>Print all eForms</a>&nbsp;&nbsp;&nbsp;&nbsp;")
        End If

        'close button
        Call AddStringToVarArr(vJSComm, "<a href='javascript:fnClose();'>Close</a>")
        Call AddStringToVarArr(vJSComm, "</td>")

        Call AddStringToVarArr(vJSComm, "</tr>" & "</table>")


        'end of outer table, first cell
        Call AddStringToVarArr(vJSComm, "</td>" _
                      & "</tr>")


        'height spacer cell
        Call AddStringToVarArr(vJSComm, "<tr height='10' class='clsScheduleBorder'>" _
                        & "<td colspan='" & oS.ColMax + 1 & "'>" _
                        & "</td>" _
                      & "</tr>")

        'MLM 03/03/03: Table break here between header (non-scrolling) and schedule (scrolling)
        Call AddStringToVarArr(vJSComm, "</table><span id=outertable><table id='main' width='100%' bgcolor=white border='0' cellpadding='0' cellspacing='0'>")

        'blue line/down v's
        Call AddStringToVarArr(vJSComm, "<tr id=head height='2' class='clsScheduleBorder'>" _
                        & "<td width='170'>" _
                        & "</td>" _
                        & "<td bgcolor='blue' colspan=" & oS.ColMax & ">" _
                        & "</td>" _
                      & "</tr>" _
                      & "<tr id='head' height='3' class='clsScheduleBorder'>" _
                        & "<td>" _
                        & "</td>")

        For nCol = 1 To oS.ColMax
            Call AddStringToVarArr(vJSComm, "<td valign='top' align='center'>" _
                            & "<img src='../img/v.gif'>" _
                            & "</td>")
        Next
        Call AddStringToVarArr(vJSComm, "</tr>")


        'visit names with cycle number if applicable
        Call AddStringToVarArr(vJSComm, "<tr id='head' align='center' height='20' class='clsScheduleBorder clsScheduleVisitText'><td></td>")
        ReDim bVHasSDV(oS.ColMax)
        For nCol = 1 To oS.ColMax
            sCycle = ""
            If Not oS.Cells(0, nCol).VisitInst Is Nothing Then
                
                If oS.Cells(0, nCol).VisitInst.CycleNo > 1 Then
                    sCycle = "[" & oS.Cells(0, nCol).VisitInst.CycleNo & "]"
                End If

                If enInterface = iwww Then
                    'calculate lock/freeze info for subject,see if we have a locked or frozen item
                    bLocked = False
                    bFrozen = False
                    eLockFreezeStatus = oS.Cells(0, nCol).VisitInst.LockStatus
                    Select Case eLockFreezeStatus
                        Case eLockStatus.lsLocked
                            bLocked = True
                        Case eLockStatus.lsFrozen
                            bFrozen = True
                        Case Else
                    End Select
                    ' (User's Unfreeze permission is checked later)
                    If bFrozen Then
                        ' unfreeze if subject not frozen
                        bCanUnFreeze = (oSubject.LockStatus <> eLockStatus.lsFrozen)
                    Else
                        ' Can't unfreeze if not frozen!
                        bCanUnFreeze = False
                    End If
                    
                    'does this visit have an sdv
                    bVHasSDV(nCol) = oMIMData.MIMessageExists(oUser.CurrentDBConString, mimtSDVMark, mimscVisit, _
                    oSubject.StudyDef.Name, oSubject.Site, oSubject.PersonId, lObjectId, nObjectSource, _
                    oS.Cells(0, nCol).VisitInst.VisitId, oS.Cells(0, nCol).VisitInst.CycleNo)
             
                    'www: <a onmouseup=fnV(...)>
                    'ic 08/09/2004 removed 'can add sdv' argument - adding an sdv has to be tried
                    'ic 01/03/2004 added 'can add sdv argument
                    Call AddStringToVarArr(vJSComm, "<a onMouseup='javascript:fnV(event.button," & Chr(34) & oS.Cells(0, nCol).VisitInst.VisitId _
                        & Chr(34) & "," & Chr(34) & oS.Cells(0, nCol).VisitInst.CycleNo & Chr(34) & "," & Chr(34) _
                        & oS.Cells(0, nCol).VisitInst.Code & Chr(34) & "," & Chr(34) & oS.Cells(0, nCol).VisitInst.Status & Chr(34) & "," _
                        & RtnJSBoolean(bLocked) & "," & RtnJSBoolean(bFrozen) & "," & RtnJSBoolean(bCanUnFreeze) & "," & RtnJSBoolean(bVHasSDV(nCol)) & ")'>")

                Else
                    'win: <a onmouseup=fnV(...)>
                    Call AddStringToVarArr(vJSComm, "<a onMouseup='javascript:fnV(" & Chr(34) & oS.Cells(0, nCol).VisitInst.VisitTaskId & Chr(34) & ",event.button)'>")

                End If
                   
            End If
            
            '<td title>visit name</td>
            If (bShowVCode) Then
                Call AddStringToVarArr(vJSComm, "<td style='cursor:default;' title='" & oS.Cells(0, nCol).Visit.Name & sCycle & "'>&nbsp;")
                Call AddStringToVarArr(vJSComm, oS.Cells(0, nCol).Visit.Code & sCycle)
                Call AddStringToVarArr(vJSComm, "&nbsp;</td>")
        
            Else
                Call AddStringToVarArr(vJSComm, "<td style='cursor:default;' title='" & oS.Cells(0, nCol).Visit.Code & sCycle & "'>&nbsp;")
                Call AddStringToVarArr(vJSComm, oS.Cells(0, nCol).Visit.Name & sCycle)
                Call AddStringToVarArr(vJSComm, "&nbsp;</td>")
                
            End If
            
            If Not oS.Cells(0, nCol).VisitInst Is Nothing Then
                Call AddStringToVarArr(vJSComm, "</a>")
            End If
        Next
        Call AddStringToVarArr(vJSComm, "</tr>")


        'visit dates
        Call AddStringToVarArr(vJSComm, "<tr id='head' align='center' class='clsScheduleBorder clsScheduleVisitDateText' height='20'>" _
                        & "<td>" _
                        & "</td>")
        For nCol = 1 To oS.ColMax
            Call AddStringToVarArr(vJSComm, "<td>")
            If Not oS.Cells(0, nCol).VisitInst Is Nothing Then
                Call AddStringToVarArr(vJSComm, "<table><tr>")
                Call AddStringToVarArr(vJSComm, "<td align='center' class='clsScheduleBorder clsScheduleVisitDateText' height='20'")
                Call AddStringToVarArr(vJSComm, " title='Please open a form to edit the visit date'>")
                Call AddStringToVarArr(vJSComm, oS.Cells(0, nCol).VisitInst.VisitDateString)
                Call AddStringToVarArr(vJSComm, "</td></tr><tr>")

                If enInterface = iwww Then
                    'calculate lock/freeze info for subject,see if we have a locked or frozen item
                    bLocked = False
                    bFrozen = False
                    eLockFreezeStatus = oS.Cells(0, nCol).VisitInst.LockStatus
                    Select Case eLockFreezeStatus
                        Case eLockStatus.lsLocked
                            bLocked = True
                        Case eLockStatus.lsFrozen
                            bFrozen = True
                        Case Else
                    End Select
                    ' (User's Unfreeze permission is checked later)
                    If bFrozen Then
                        ' unfreeze if subject not frozen
                        bCanUnFreeze = (oSubject.LockStatus <> eLockStatus.lsFrozen)
                    Else
                        ' Can't unfreeze if not frozen!
                        bCanUnFreeze = False
                    End If
                    
                    'www: <a onmouseup=fnV(...)>
                    'ic 08/09/2004 removed 'can add sdv' argument - adding an sdv has to be tried
                    'ic 01/03/2004 added 'can add sdv argument
                    Call AddStringToVarArr(vJSComm, "<a onMouseup='javascript:fnV(event.button," & Chr(34) & oS.Cells(0, nCol).VisitInst.VisitId _
                        & Chr(34) & "," & Chr(34) & oS.Cells(0, nCol).VisitInst.CycleNo & Chr(34) & "," & Chr(34) _
                        & oS.Cells(0, nCol).VisitInst.Code & Chr(34) & "," & Chr(34) & oS.Cells(0, nCol).VisitInst.Status & Chr(34) & "," _
                        & RtnJSBoolean(bLocked) & "," & RtnJSBoolean(bFrozen) & "," & RtnJSBoolean(bCanUnFreeze) & "," & RtnJSBoolean(bVHasSDV(nCol)) & ")'>")

                Else
                    'win: <a onmouseup=fnV(...)>
                    Call AddStringToVarArr(vJSComm, "<a onMouseup='javascript:fnV(" & Chr(34) & oS.Cells(0, nCol).VisitInst.VisitTaskId _
                        & Chr(34) & ",event.button)'>")
                End If
                
                '<td>status image</td>
                Call AddStringToVarArr(vJSComm, "<td align='center' style='cursor:default;'>" _
                & RtnStatusImages(oS.Cells(0, nCol).VisitInst.Status, bViewInformIcon, oS.Cells(0, nCol).VisitInst.LockStatus, _
                False, oS.Cells(0, nCol).VisitInst.SDVStatus, oS.Cells(0, nCol).VisitInst.DiscrepancyStatus) & "</td><a/></tr></table>")

            End If
            Call AddStringToVarArr(vJSComm, "</td>")
            
        Next
        Call AddStringToVarArr(vJSComm, "</tr>")

        For nRow = 1 To oS.RowMax

            'add eform header if we havent already
            If nEfId <> oS.Cells(nRow, 0).eForm.EFormId Then
                nEfId = oS.Cells(nRow, 0).eForm.EFormId
                
                If (bShowECode) Then
                    Call AddStringToVarArr(vJSComm, "<tr valign='top' align='center'>" _
                                    & "<td class='clsScheduleBorder' title='" & oS.Cells(nRow, 0).eForm.Name & "'>" _
                                      & "<div class='clsScheduleEformBorder clsScheduleEformText'>" & oS.Cells(nRow, 0).eForm.Code & "</div>" _
                                    & "</td>")
                
                Else
                    Call AddStringToVarArr(vJSComm, "<tr valign='top' align='center'>" _
                                    & "<td class='clsScheduleBorder' title='" & oS.Cells(nRow, 0).eForm.Code & "'>" _
                                      & "<div class='clsScheduleEformBorder clsScheduleEformText'>" & oS.Cells(nRow, 0).eForm.Name & "</div>" _
                                    & "</td>")
                
                End If

            Else
                Call AddStringToVarArr(vJSComm, "<tr valign='top' align='center'>" _
                                & "<td class='clsScheduleBorder'></td>")
                                
            End If


            'visit columns
            For nCol = 1 To oS.ColMax
                Set oCell = oS.Cells(nRow, nCol)

                'cell bgcolor
                Call AddStringToVarArr(vJSComm, "<td align='center' valign='top' class='clsScheduleEformLabelText'")
                If oS.Cells(0, nCol).Visit.BackgroundColour > 0 Then
                    Call AddStringToVarArr(vJSComm, " bgcolor='" & RtnHTMLCol(oS.Cells(0, nCol).Visit.BackgroundColour) & "'")
                End If
                Call AddStringToVarArr(vJSComm, ">")

                Select Case oCell.CellType
                    Case Blank:
                    Case Inactive:
                        Call AddStringToVarArr(vJSComm, RtnEFormImages(0, 0, 0, 99, False) & "<br><br><br>")
                        
                    Case Active:
                            With oCell.eFormInst
                                If (enInterface = iwww) Then
                                    ' calculate lock/freeze info for subject
                                    ' See if we have a locked or frozen item
                                    bLocked = False
                                    bFrozen = False
                                    eLockFreezeStatus = .LockStatus
                                    Select Case eLockFreezeStatus
                                        Case eLockStatus.lsLocked
                                            bLocked = True
                                        Case eLockStatus.lsFrozen
                                            bFrozen = True
                                        Case Else
                                    End Select
                                    ' (User's Unfreeze permission is checked later)
                                    If bFrozen Then
                                        ' unfreeze if subject/visit not frozen
                                        bCanUnFreeze = (oSubject.LockStatus <> eLockStatus.lsFrozen) And _
                                                (.VisitInstance.LockStatus <> eLockStatus.lsFrozen)
                                    Else
                                        ' Can't unfreeze if not frozen!
                                        bCanUnFreeze = False
                                    End If
                                    
                                    'does this eform have an sdv
                                    bEHasSDV = oMIMData.MIMessageExists(oUser.CurrentDBConString, mimtSDVMark, mimscEForm, _
                                    oSubject.StudyDef.Name, oSubject.Site, oSubject.PersonId, lObjectId, nObjectSource, _
                                    .VisitInstance.VisitId, .VisitInstance.CycleNo, .EFormTaskId)
                                    
                                    'ic 08/09/2004 removed 'can add sdv' argument - adding an sdv has to be tried
                                    'ic 01/03/2004 added 'can add sdv argument
                                    Call AddStringToVarArr(vJSComm, "<table><tr><a onMouseup='javascript:fnE(event.button," & Chr(34) _
                                        & .EFormTaskId & Chr(34) & "," & Chr(34) & .eForm.EFormId & Chr(34) & "," & Chr(34) & .CycleNo _
                                        & Chr(34) & "," & Chr(34) & .Code & Chr(34) & "," & Chr(34) & .VisitInstance.VisitId & Chr(34) & "," _
                                        & Chr(34) & .VisitInstance.CycleNo & Chr(34) & "," & Chr(34) & .VisitInstance.Code & Chr(34) _
                                        & "," & Chr(34) & .Status & Chr(34) & "," & RtnJSBoolean(bLocked) & "," & RtnJSBoolean(bFrozen) & "," _
                                        & RtnJSBoolean(bCanUnFreeze) & "," _
                                    & RtnJSBoolean(oUser.CheckPermission(gsFnCreateSDV) And CanChangePlannedSDVs(oCell.eFormInst)) & "," _
                                    & RtnJSBoolean(bSHasSDV) & "," & RtnJSBoolean(bVHasSDV(nCol)) & "," & RtnJSBoolean(bEHasSDV) & ")'><td>")
                                Else
                                    Call AddStringToVarArr(vJSComm, "<table><tr><a onMouseup='javascript:fnE(" & Chr(34) & .EFormTaskId & Chr(34) & ",event.button)'><td>")
                                    
                                End If
                                Call AddStringToVarArr(vJSComm, RtnEFormImages(.LockStatus, _
                                                              .DiscrepancyStatus, _
                                                              .SDVStatus, _
                                                              .Status, _
                                                              False))
                                Call AddStringToVarArr(vJSComm, "</td></a></tr></table>")
                                Call AddStringToVarArr(vJSComm, .eFormLabel & "<br>")
                                Call AddStringToVarArr(vJSComm, .eFormDateString & "<br>")
                                
                            End With

                End Select
                Call AddStringToVarArr(vJSComm, "</td>")

            Next
            Call AddStringToVarArr(vJSComm, "</tr>")

        Next

        Call AddStringToVarArr(vJSComm, "</table></span><table id=rowhead class=abs cellpadding=0 cellspacing=0></table>" _
            & "<table id=colhead class=abs cellpadding=0 cellspacing=0></table><div id=blank class=clsScheduleBorder></div>")

        Call AddStringToVarArr(vJSComm, "</body>" & vbCrLf)
        
        If enInterface = iWindows Then
            Call AddStringToVarArr(vJSComm, "</html>")
        End If

    Set oMIMData = Nothing
    Set oS = Nothing
    Set oCell = Nothing
    GetScheduleNoCompression = Join(vJSComm, "")
    
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLSchedule.GetScheduleNoCompression"
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetSchedule(ByRef oSubject As StudySubject, Optional ByVal enInterface As eInterface = iwww, _
    Optional ByRef oUser As MACROUser, Optional ByVal vErrors As Variant, Optional ByVal vAlerts As Variant, _
    Optional ByVal bNew As Boolean, Optional sSDVCall As String = "") As String

'--------------------------------------------------------------------------------------------------
'   ic 11/07/01
'   builds and returns a string representing a subject schedule in an html table
'
'   ***********************************************************************************************
'   NOTE: IF YOU MAKE CHANGES TO THIS FUNCTION, MAKE CORRESPONDING CHANGES TO NON-COMPRESSION
'   FUNCTION ABOVE
'   ***********************************************************************************************
'
'   REVISIONS
'   DPH 19/12/2002 - Show EITHER subject label OR (subjectid) not both
'   ic 13/01/2003   display visit name if visit inactive
'   DPH 16/01/2003 - Lock/Unlock/Freeze/UnFreeze functionality for subject/visit/eForm
'   DPH 23/01/2003 - Split out <head> for WWW as showing "please wait..."
'   ic 28/01/2003   fixed subject label/icon split line
'   ic 04/02/2003   fixed hand mousepointer over subject/visit
'   DPH 18/02/2003  uses GetLocalFormatDate format for Visit / Eform dates
'   ic 01/04/2003   added flag for 'display code or name'
'   ic 29/05/2003 removed 'onmouseout' event for popup menu, bug 1761
'   ic 05/06/2003 added extra 'refresh z-order' parameter
'   ic 19/06/2003 added registration
'   ic 30/09/2003 changed to pass whole image tag, otherwise IE downloads same image many times
'   ic 01/10/2003 moved fnHideLoader() inside fnPageLoaded()
'   ic 16/10/2003 check quicklist permission
'   ic 01/03/2004 added 'can add sdv argument
'   ic 29/06/2004 added error handling
'   ic 24/08/2004 added 'can set sdvs to done' fnE() parameter, vAlerts parameter
'   ic 08/09/2004 removed 'can add sdv' argument - adding an sdv has to be tried
'   ic 01/04/2005 issue 2541 added sSDVCall parameter
'   ic 29/06/2005 issue 2464, enhancements to sdv work flow, added visit, eform, question cycle
'--------------------------------------------------------------------------------------------------
Dim oS As ScheduleGrid
Dim oCell As GridCell
Dim nCol As Integer
Dim nRow As Integer
Dim nEfId As Long
Dim sCycle As String
Dim nLoop As Integer
Dim bLocked As Boolean
Dim bFrozen As Boolean
Dim bCanUnFreeze As Boolean
Dim eLockFreezeStatus As eLockStatus
Dim bViewInformIcon As Boolean
Dim vJSComm() As String
Dim bShowVCode As Boolean
Dim bShowECode As Boolean
Dim lRaised As Long
Dim lResponded As Long
Dim lPlanned As Long
Dim bQuickList As Boolean
Dim oMIMData As MIDataLists
Dim bEHasSDV As Boolean
Dim bVHasSDV() As Boolean
Dim bSHasSDV As Boolean
Dim lObjectId As Long
Dim nObjectSource As Integer

    On Error GoTo CatchAllError
    ReDim vJSComm(0)

    Set oS = oSubject.ScheduleGrid

    bViewInformIcon = oUser.CheckPermission(gsFnMonitorDataReviewData)
    bQuickList = oUser.CheckPermission(gsFnViewQuickList)
    'show code header rather than full name
    bShowVCode = mbShowCode
    bShowECode = mbShowCode

    Set oMIMData = New MIDataLists
    'Call AddStringToVarArr(vJSComm, "<html>")

    If (enInterface = iwww) Then
        
        ' DPH 23/01/2003
        'sHTML = sHTML & "<head>" _
        '              & "<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>"
        'sHTML = sHTML & "<script language='javascript' src='../script/Schedule.js'></script>"

        'ic 29/05/2003 removed 'onmouseout' event for popup menu, bug 1761
        Call AddStringToVarArr(vJSComm, "<body class=clsScheduleBorder onload='fnPageLoaded();' onscroll='PositionHeaders();fnPopMenuHide();'>" & vbCrLf _
                      & "<div class='clsPopMenu' id='divPopMenu' onclick=" & Chr(34) & "menu=this;this.tid=setTimeout('menu.style.visibility=\'hidden\'',20);" & Chr(34) _
                      & "onmouseover='clearTimeout(this.tid);'>" _
                      & "</div>" & vbCrLf)

        'ic 01/10/2003 moved fnHideLoader() inside fnPageLoaded()
        ' moved JavaScript code previously in <head> to here
        'DPH 16/01/2003 - Added new user permissions required
        'ic 02/09/2003 use new CanChangeData function
        Call AddStringToVarArr(vJSComm, "<script language='javascript'>" & vbCrLf _
                      & "function fnPageLoaded()" & vbCrLf _
                        & "{" & vbCrLf _
                          & "fnHideLoader();" & vbCrLf _
                          & "fnInitUser(" & RtnJSBoolean(oUser.CheckPermission(gsFnViewData)) & "," _
                                          & RtnJSBoolean(CanChangeData(oUser, oSubject.Site)) & "," _
                                          & RtnJSBoolean(oUser.CheckPermission(gsFnCreateSDV)) & "," _
                                          & RtnJSBoolean(oUser.CheckPermission(gsFnLockData)) & "," _
                                          & RtnJSBoolean(oUser.CheckPermission(gsFnFreezeData)) & "," _
                                          & RtnJSBoolean(oUser.CheckPermission(gsFnUnFreezeData)) & ")" & vbCrLf _
                          & "fnInitSubject('" & oSubject.StudyId & "'," _
                                        & "'" & oSubject.Site & "'," _
                                        & "'" & oSubject.PersonId & "'," _
                                        & "'" & oSubject.Label & "')" & vbCrLf _
                          & "window.sWinState='4" & gsDELIMITER2 & oSubject.StudyId & gsDELIMITER2 & oSubject.Site & gsDELIMITER2 & oSubject.PersonId & "';" & vbCrLf)

        'ic 16/10/2003 check quicklist permission
        If (bNew And bQuickList) Then
            'studyid|studyname|site|subjectid|subjectlabel
            Call AddStringToVarArr(vJSComm, "window.parent.fnAddToQuickList('" & ReplaceWithJSChars(oSubject.StudyId & gsDELIMITER2 _
                & oSubject.StudyDef.Name & gsDELIMITER2 & oSubject.Site & gsDELIMITER2 & oSubject.PersonId & gsDELIMITER2 _
                & oSubject.Label) & "');" & vbCrLf)
        End If

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
        
        'any additional alerts to display
        If Not IsMissing(vAlerts) Then
            If Not IsEmpty(vAlerts) Then
                For nLoop = LBound(vAlerts, 2) To UBound(vAlerts, 2)
                    Call AddStringToVarArr(vJSComm, "alert('" & ReplaceWithJSChars(vAlerts(1, nLoop)) & "');" & vbCrLf)
                Next
            End If
        End If

        'MLM 05/02/03: Call the javascript that will expand the page into nested tables
        Call AddStringToVarArr(vJSComm, "document.all('outertable').innerHTML=fnExpand(document.all('outertable').innerHTML);" & vbCrLf)
        'Call AddStringToVarArr(vJSComm, "document.all('debug').innerText=document.all('outertable').innerHTML;")
        'MLM 04/03/03: Call JavaScript to set up floating schedule headers
        Call AddStringToVarArr(vJSComm, "DrawHeaders();" & vbCrLf)

        'TA 30/11/2003: moved here so not run in windows
        Call RtnMIMsgStatusCount(oUser, lRaised, lResponded, lPlanned)
        'ic 05/06/2003 added extra 'refresh z-order' parameter
        Call AddStringToVarArr(vJSComm, "window.parent.fnSTLC('" & gsVIEW_RAISED_DISCREPANCIES_MENUID & "','" & CStr(lRaised) & "',0);" & vbCrLf _
                                      & "window.parent.fnSTLC('" & gsVIEW_RESPONDED_DISCREPANCIES_MENUID & "','" & CStr(lResponded) & "',0);" & vbCrLf _
                                      & "window.parent.fnSTLC('" & gsVIEW_PLANNED_SDV_MARKS_MENUID & "','" & CStr(lPlanned) & "',1);" & vbCrLf)

        'ic 19/08/2003 enable/disable registration menu item
        'ic 02/09/2003 use new CanChangeData function
        If (CanChangeData(oUser, oSubject.Site) And oUser.CheckPermission(gsFnRegisterSubject) And ShouldEnableRegistrationMenu(oSubject)) Then
            Call AddStringToVarArr(vJSComm, "window.parent.fnEnableRegister(" & RtnJSBoolean(True) & ");" & vbCrLf)
        Else
            Call AddStringToVarArr(vJSComm, "window.parent.fnEnableRegister(" & RtnJSBoolean(False) & ");" & vbCrLf)
        End If

        'write javascript sdv call (if there is one)
        Call AddStringToVarArr(vJSComm, sSDVCall & vbCrLf)
        
        Call AddStringToVarArr(vJSComm, "}var bShowSDVScheduleMenu=" & RtnJSBoolean(RtnShowSDVScheduleMenuFlag()) & ";" & vbCrLf)

        Call AddStringToVarArr(vJSComm, "</script>" & vbCrLf)

        Call AddStringToVarArr(vJSComm, "<form name='Form1' action='Schedule.asp?fltSi=" & oSubject.Site & "&fltSt=" & oSubject.StudyId & "&fltSj=" & oSubject.PersonId & "' method='post'>" _
                      & "<input type='hidden' name='SchedUpdate'>" _
                      & "<input type='hidden' name='SchedIdentifier'>" _
                      & "</form>" & vbCrLf)

    Else
        Call AddStringToVarArr(vJSComm, "<head>")
        Call AddStringToVarArr(vJSComm, "<script language='javascript'>" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnE(id,btn){window.navigate('VBfnEformUrl|'+id+'|'+btn);}" _
            & "function fnPrint(){window.navigate('VBfnPrintAll');}" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "function fnClose(){window.navigate('VBfnCloseSchedule');}" & vbCrLf)

        Call AddStringToVarArr(vJSComm, "function fnV(id,btn){window.navigate('VBfnV|'+id+'|'+btn);}" _
            & "function fnS(btn){window.navigate('VBfnS|'+btn);}" & vbCrLf)
        'MLM 12/02/03: Added fnPageLoaded to Windows version
        Call AddStringToVarArr(vJSComm, "function fnPageLoaded(){")
        Call AddStringToVarArr(vJSComm, "document.all('outertable').innerHTML=fnExpand(document.all('outertable').innerHTML);" & vbCrLf)
        'MLM 04/03/03: Call JavaScript to set up floating schedule headers
        Call AddStringToVarArr(vJSComm, "DrawHeaders();" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "}" & vbCrLf)

        Call AddStringToVarArr(vJSComm, "</script>" & vbCrLf)
        Call AddStringToVarArr(vJSComm, "</head><body class=clsScheduleBorder onscroll='PositionHeaders();'>" & vbCrLf)
    End If




    'outer table, first cell holds heading table
    'MLM 05/02/03: Added span.
    Call AddStringToVarArr(vJSComm, "<div><br><br></div><table id=top height=38 width='100%' border='0' cellpadding='0' cellspacing='0'>" _
                    & "<tr height='20'>" _
                      & "<td colspan='" & oS.ColMax + 1 & "'>")

        ' DPH 19/12/2002 - use sLabelOrId
        ' ic 30/01/2003 switched to RtnSubjectText()
        'heading/top menu table
        Call AddStringToVarArr(vJSComm, "<table width='100%' height='100%' cellpadding='0' cellspacing='0'>" _
                        & "<tr class='clsScheduleBorder'>" _
                          & "<td>" _
                            & "<table>" _
                              & "<tr>" _
                                & "<td align='left'>")

        If (enInterface = iwww) Then
            ' calculate lock/freeze info for subject
            ' See if we have a locked or frozen item
            bLocked = False
            bFrozen = False
            eLockFreezeStatus = oSubject.LockStatus
            Select Case eLockFreezeStatus
                Case eLockStatus.lsLocked
                    bLocked = True
                Case eLockStatus.lsFrozen
                    bFrozen = True
                Case Else
            End Select
            ' (User's Unfreeze permission is checked later)
            If bFrozen Then
                bCanUnFreeze = True
            Else
                ' Can't unfreeze if not frozen!
                bCanUnFreeze = False
            End If

            'does this subject have an sdv
            bSHasSDV = oMIMData.MIMessageExists(oUser.CurrentDBConString, mimtSDVMark, mimscSubject, _
            oSubject.StudyDef.Name, oSubject.Site, oSubject.PersonId, lObjectId, nObjectSource)

            'ic 08/09/2004 removed 'can add sdv' argument - adding an sdv has to be tried
            'ic 01/03/2004 added 'can add sdv argument
            Call AddStringToVarArr(vJSComm, "<a onMouseup='javascript:fnS(event.button," & Chr(34) _
            & oSubject.Status & Chr(34) & "," & RtnJSBoolean(bLocked) & "," & RtnJSBoolean(bFrozen) & "," _
            & RtnJSBoolean(bCanUnFreeze) & "," & RtnJSBoolean(bSHasSDV) & ")'>")
        Else
            Call AddStringToVarArr(vJSComm, "<a onMouseup='javascript:fnS(event.button)'>")
        End If

        'Call AddStringToVarArr(vJSComm, "<div  style='cursor:hand;' class='clsScheduleEformBorder'><table cellpadding='0' cellspacing='0'><tr class='clsScheduleHeadingText'><td>" & oSubject.StudyCode & "/" & oSubject.Site & "/" & RtnSubjectText(oSubject.PersonID, oSubject.Label) & "&nbsp;</td><td>" & RtnStatusImage(oSubject.Status, False, oSubject.LockStatus, "", "", oSubject.SDVStatus) & "</td></tr></table></div>"
        Call AddStringToVarArr(vJSComm, "<div  style='cursor:default;' class='clsScheduleEformBorder'><table cellpadding='0' cellspacing='0'><tr class='clsScheduleHeadingText'><td>" & oSubject.StudyCode & "/" & oSubject.Site & "/" & RtnSubjectText(oSubject.PersonId, oSubject.Label) & "&nbsp;</td><td>" & RtnStatusImages(oSubject.Status, bViewInformIcon, oSubject.LockStatus, False, oSubject.SDVStatus, oSubject.DiscrepancyStatus) & "</td></tr></table></div>" _
                                & "</a>" _
                                & "</td>" _
                              & "</tr>" _
                            & "</table>" _
                          & "</td>")

        Call AddStringToVarArr(vJSComm, "<td align='center' class='clsScheduleMenuLinkText'>")

        'TA for windows add print all eforms link
        ' NCJ 30 Jun 03 - Changed eforms to eForms
        If (enInterface = eInterface.iWindows) Then
            Call AddStringToVarArr(vJSComm, "<a href='javascript:fnPrint();'>Print all eForms</a>&nbsp;&nbsp;&nbsp;&nbsp;")
        End If

        'close button
        Call AddStringToVarArr(vJSComm, "<a href='javascript:fnClose();'>Close</a>")
        Call AddStringToVarArr(vJSComm, "</td>")

        Call AddStringToVarArr(vJSComm, "</tr>" & "</table>")


        'end of outer table, first cell
        Call AddStringToVarArr(vJSComm, "</td>" _
                      & "</tr>")


        'height spacer cell
        Call AddStringToVarArr(vJSComm, "<tr height='10' class='clsScheduleBorder'>" _
                        & "<td colspan='" & oS.ColMax + 1 & "'>" _
                        & "</td>" _
                      & "</tr>")

        'MLM 03/03/03: Table break here between header (non-scrolling) and schedule (scrolling)
        Call AddStringToVarArr(vJSComm, "</table><span id=outertable><table id=main width='100%' bgcolor=white border='0' cellpadding='0' cellspacing='0'>")

        'blue line/down v's
        Call AddStringToVarArr(vJSComm, "<tr id=head height='2' class='clsScheduleBorder'>" _
                        & "<td width='170'>" _
                        & "</td>" _
                        & "<td bgcolor='blue' colspan=" & oS.ColMax & ">" _
                        & "</td>" _
                      & "</tr>" _
                      & "<tr id=head height='3' class='clsScheduleBorder'>" _
                        & "<td>" _
                        & "</td>")

        For nCol = 1 To oS.ColMax
            'ic 30/09/2003 changed to pass whole image tag, otherwise IE downloads same image many times
            Call AddStringToVarArr(vJSComm, msV)
            Call AddStringToVarArr(vJSComm, "<img src='../img/v.gif'>" & "</td>")
        Next
        Call AddStringToVarArr(vJSComm, "</tr>")


        'visit names with cycle number if applicable
        Call AddStringToVarArr(vJSComm, "<tr id=head align='center' height='20' class='clsScheduleBorder clsScheduleVisitText'><td></td>")
        ReDim bVHasSDV(oS.ColMax)
        For nCol = 1 To oS.ColMax
            sCycle = ""
            If Not oS.Cells(0, nCol).VisitInst Is Nothing Then
                If oS.Cells(0, nCol).VisitInst.CycleNo > 1 Then sCycle = "[" & oS.Cells(0, nCol).VisitInst.CycleNo & "]"


                If enInterface = iwww Then
                    ' calculate lock/freeze info for subject
                    ' See if we have a locked or frozen item
                    bLocked = False
                    bFrozen = False
                    eLockFreezeStatus = oS.Cells(0, nCol).VisitInst.LockStatus
                    Select Case eLockFreezeStatus
                        Case eLockStatus.lsLocked
                            bLocked = True
                        Case eLockStatus.lsFrozen
                            bFrozen = True
                        Case Else
                    End Select
                    ' (User's Unfreeze permission is checked later)
                    If bFrozen Then
                        ' unfreeze if subject not frozen
                        bCanUnFreeze = (oSubject.LockStatus <> eLockStatus.lsFrozen)
                    Else
                        ' Can't unfreeze if not frozen!
                        bCanUnFreeze = False
                    End If
                    
                    'does this visit have an sdv
                    bVHasSDV(nCol) = oMIMData.MIMessageExists(oUser.CurrentDBConString, mimtSDVMark, mimscVisit, _
                    oSubject.StudyDef.Name, oSubject.Site, oSubject.PersonId, lObjectId, nObjectSource, _
                    oS.Cells(0, nCol).VisitInst.VisitId, oS.Cells(0, nCol).VisitInst.CycleNo)

'                    Call AddStringToVarArr(vJSComm, "<a onMouseup='javascript:fnV(event.button," & Chr(34) & oS.Cells(0, nCol).VisitInst.VisitId & Chr(34) & "," & Chr(34) _
'                        & oS.Cells(0, nCol).VisitInst.CycleNo & Chr(34) & "," & Chr(34) & oS.Cells(0, nCol).VisitInst.Code & Chr(34) & "," & Chr(34) & oS.Cells(0, nCol).VisitInst.Status & Chr(34) & "," _
'                        & CInt(bLocked) & "," & CInt(bFrozen) & "," & CInt(bCanUnFreeze) & ")'>")
'                Else
'                    'windows needs id as well
'                    Call AddStringToVarArr(vJSComm, "<a onMouseup='javascript:fnV(" & Chr(34) & oS.Cells(0, nCol).VisitInst.VisitTaskId & Chr(34) & ",event.button)'>")
'                End If
'                Call AddStringToVarArr(vJSComm, "<td style='cursor:default;' title='" & oS.Cells(0, nCol).VisitInst.Visit.Name & sCycle & "'>&nbsp;")
'                Call AddStringToVarArr(vJSComm, oS.Cells(0, nCol).VisitInst.Visit.Code & sCycle)
'                Call AddStringToVarArr(vJSComm, "&nbsp;</td></a>")
'            Else
'                'ic 13/01/2003 display visitname if visit is inactive
'                Call AddStringToVarArr(vJSComm, "<td style='cursor:default;' title='" & oS.Cells(0, nCol).Visit.Name & "'>&nbsp;" & oS.Cells(0, nCol).Visit.Code & "&nbsp;</td>")
                    'ic 08/09/2004 removed 'can add sdv' argument - adding an sdv has to be tried
                    'ic 01/03/2004 added 'can add sdv argument
                    Call AddStringToVarArr(vJSComm, msVISIT_TITLE_START & "fnV(event.button," & Chr(34) & oS.Cells(0, nCol).VisitInst.VisitId & Chr(34) & "," & Chr(34) _
                        & oS.Cells(0, nCol).VisitInst.CycleNo & Chr(34) & "," & Chr(34) & oS.Cells(0, nCol).VisitInst.Code & Chr(34) & "," & Chr(34) & oS.Cells(0, nCol).VisitInst.Status & Chr(34) & "," _
                        & RtnJSBoolean(bLocked) & "," & RtnJSBoolean(bFrozen) & "," & RtnJSBoolean(bCanUnFreeze) & "," & RtnJSBoolean(bVHasSDV(nCol)) & ")")
                Else
                    'windows needs id as well
                    Call AddStringToVarArr(vJSComm, msVISIT_TITLE_START & "fnV(" & Chr(34) & oS.Cells(0, nCol).VisitInst.VisitTaskId & Chr(34) & ",event.button)")
                End If
            Else
'                'ic 13/01/2003 display visitname if visit is inactive
                Call AddStringToVarArr(vJSComm, msVISIT_TITLE_START)
            End If
            If (bShowVCode) Then
                ' NCJ 14 Feb 03 - Use cell's Visit rather than VisitInstance (in case there isn't a VI)
                Call AddStringToVarArr(vJSComm, msDELIMITER & oS.Cells(0, nCol).Visit.Name & sCycle & msDELIMITER & oS.Cells(0, nCol).Visit.Code & sCycle & msVISIT_TITLE_END)
            Else
                Call AddStringToVarArr(vJSComm, msDELIMITER & oS.Cells(0, nCol).Visit.Code & sCycle & msDELIMITER & oS.Cells(0, nCol).Visit.Name & sCycle & msVISIT_TITLE_END)
            End If
        Next
        Call AddStringToVarArr(vJSComm, "</tr>")


        'visit dates
        Call AddStringToVarArr(vJSComm, "<tr id=head align='center' class='clsScheduleBorder clsScheduleVisitDateText' height='20'>" _
                        & "<td>" _
                        & "</td>")
        For nCol = 1 To oS.ColMax
            Call AddStringToVarArr(vJSComm, "<td>")
            If Not oS.Cells(0, nCol).VisitInst Is Nothing Then
                'sHTML = sHTML & oS.Cells(0, nCol).VisitInst.VisitDateString
                ' DPH 18/02/2003 use GetLocalFormatDate
                ' MLM 03/03/03: *do* use VisitDateString
                'Call AddStringToVarArr(vJSComm, msVISIT_START & GetLocalFormatDate(oUser, oS.Cells(0, nCol).VisitInst.VisitDate, dttDateOnly) & msDELIMITER)
                Call AddStringToVarArr(vJSComm, msVISIT_START & oS.Cells(0, nCol).VisitInst.VisitDateString & msDELIMITER)
'                Call AddStringToVarArr(vJSComm, "<table><tr>")
'                Call AddStringToVarArr(vJSComm, "<td align='center' class='clsScheduleBorder clsScheduleVisitDateText' height='20'")
'                Call AddStringToVarArr(vJSComm, " title='Please open a form to edit the visit date'>")
'                Call AddStringToVarArr(vJSComm, oS.Cells(0, nCol).VisitInst.VisitDateString)
'                Call AddStringToVarArr(vJSComm, "</td></tr><tr>")

                If enInterface = iwww Then
                    ' calculate lock/freeze info for subject
                    ' See if we have a locked or frozen item
                    bLocked = False
                    bFrozen = False
                    eLockFreezeStatus = oS.Cells(0, nCol).VisitInst.LockStatus
                    Select Case eLockFreezeStatus
                        Case eLockStatus.lsLocked
                            bLocked = True
                        Case eLockStatus.lsFrozen
                            bFrozen = True
                        Case Else
                    End Select
                    ' (User's Unfreeze permission is checked later)
                    If bFrozen Then
                        ' unfreeze if subject not frozen
                        bCanUnFreeze = (oSubject.LockStatus <> eLockStatus.lsFrozen)
                    Else
                        ' Can't unfreeze if not frozen!
                        bCanUnFreeze = False
                    End If
'                    Call AddStringToVarArr(vJSComm, "<a onMouseup='javascript:fnV(event.button," & Chr(34) & oS.Cells(0, nCol).VisitInst.VisitId & Chr(34) & "," & Chr(34) & oS.Cells(0, nCol).VisitInst.CycleNo & Chr(34) & "," _
'                            & Chr(34) & oS.Cells(0, nCol).VisitInst.Code & Chr(34) & "," & Chr(34) & oS.Cells(0, nCol).VisitInst.Status & Chr(34) & "," _
'                            & CInt(bLocked) & "," & CInt(bFrozen) & "," & CInt(bCanUnFreeze) & ")'>")
'                Else
'                    'windows needs id as well
'                    Call AddStringToVarArr(vJSComm, "<a onMouseup='javascript:fnV(" & Chr(34) & oS.Cells(0, nCol).VisitInst.VisitTaskId & Chr(34) & ",event.button)'>")
'                End If
'
'                Call AddStringToVarArr(vJSComm, "<td align='center' style='cursor:default;'>" & RtnStatusImages(oS.Cells(0, nCol).VisitInst.Status, bViewInformIcon, oS.Cells(0, nCol).VisitInst.LockStatus, False, oS.Cells(0, nCol).VisitInst.SDVStatus, oS.Cells(0, nCol).VisitInst.DiscrepancyStatus) & "</td><a/></tr></table>")

                    'ic 01/03/2004 added 'can add sdv argument
                    Call AddStringToVarArr(vJSComm, "fnV(event.button," & Chr(34) & oS.Cells(0, nCol).VisitInst.VisitId & Chr(34) & "," & Chr(34) & oS.Cells(0, nCol).VisitInst.CycleNo & Chr(34) & "," _
                            & Chr(34) & oS.Cells(0, nCol).VisitInst.Code & Chr(34) & "," & Chr(34) & oS.Cells(0, nCol).VisitInst.Status & Chr(34) & "," _
                            & RtnJSBoolean(bLocked) & "," & RtnJSBoolean(bFrozen) & "," & RtnJSBoolean(bCanUnFreeze) & "," & RtnJSBoolean(bVHasSDV(nCol)) & ")")
                Else
                    'windows needs id as well
                    Call AddStringToVarArr(vJSComm, "fnV(" & Chr(34) & oS.Cells(0, nCol).VisitInst.VisitTaskId & Chr(34) & ",event.button)")
                End If

                Call AddStringToVarArr(vJSComm, msDELIMITER & RtnStatusImages(oS.Cells(0, nCol).VisitInst.Status, bViewInformIcon, oS.Cells(0, nCol).VisitInst.LockStatus, False, oS.Cells(0, nCol).VisitInst.SDVStatus, oS.Cells(0, nCol).VisitInst.DiscrepancyStatus) & msVISIT_END)
            End If
            Call AddStringToVarArr(vJSComm, "</td>")
        Next
        Call AddStringToVarArr(vJSComm, "</tr>")


        For nRow = 1 To oS.RowMax

            'add eform header if we havent already
            If nEfId <> oS.Cells(nRow, 0).eForm.EFormId Then
                nEfId = oS.Cells(nRow, 0).eForm.EFormId
                If (bShowECode) Then
                    Call AddStringToVarArr(vJSComm, msEFORM_HEADER_START & oS.Cells(nRow, 0).eForm.Name & msDELIMITER & oS.Cells(nRow, 0).eForm.Code & msEFORM_HEADER_END)
                Else
                    Call AddStringToVarArr(vJSComm, msEFORM_HEADER_START & oS.Cells(nRow, 0).eForm.Code & msDELIMITER & oS.Cells(nRow, 0).eForm.Name & msEFORM_HEADER_END)
                End If
'                Call AddStringToVarArr(vJSComm, "<tr valign='top' align='center'>" _
'                                & "<td class='clsScheduleBorder' title='" & oS.Cells(nRow, 0).eForm.Name & "'>" _
'                                  & "<div class='clsScheduleEformBorder clsScheduleEformText'>" & oS.Cells(nRow, 0).eForm.Code & "</div>" _
'                                & "</td>")
            Else
                Call AddStringToVarArr(vJSComm, msBLANK_EFORM_HEADER)
'                Call AddStringToVarArr(vJSComm, "<tr valign='top' align='center'>" _
'                                & "<td class='clsScheduleBorder'></td>")
            End If

            'visit columns
            For nCol = 1 To oS.ColMax
                Set oCell = oS.Cells(nRow, nCol)

                'cell bgcolor
'                Call AddStringToVarArr(vJSComm, "<td align='center' valign='top' class='clsScheduleEformLabelText'")
                If oS.Cells(0, nCol).Visit.BackgroundColour > 0 Then
                    Call AddStringToVarArr(vJSComm, msCELL_START & RtnHTMLCol(oS.Cells(0, nCol).Visit.BackgroundColour) & msDELIMITER)
                Else
                 Call AddStringToVarArr(vJSComm, msCELL_START & msDELIMITER)
                End If
'                Call AddStringToVarArr(vJSComm, ">")

                Select Case oCell.CellType
                    Case Blank:
                    Case Inactive:
'                        Call AddStringToVarArr(vJSComm, RtnEFormImages(0, 0, 0, 99) & "<br><br><br>")
                        Call AddStringToVarArr(vJSComm, msTABLE_START & msDELIMITER & RtnEFormImages(0, 0, 0, 99) & "<br>" & msDELIMITER & msDELIMITER & msTABLE_END)
                    Case Active:
                            With oCell.eFormInst
                                If (enInterface = iwww) Then
                                    ' calculate lock/freeze info for subject
                                    ' See if we have a locked or frozen item
                                    bLocked = False
                                    bFrozen = False
                                    eLockFreezeStatus = .LockStatus
                                    Select Case eLockFreezeStatus
                                        Case eLockStatus.lsLocked
                                            bLocked = True
                                        Case eLockStatus.lsFrozen
                                            bFrozen = True
                                        Case Else
                                    End Select
                                    ' (User's Unfreeze permission is checked later)
                                    If bFrozen Then
                                        ' unfreeze if subject/visit not frozen
                                        bCanUnFreeze = (oSubject.LockStatus <> eLockStatus.lsFrozen) And _
                                                (.VisitInstance.LockStatus <> eLockStatus.lsFrozen)
                                    Else
                                        ' Can't unfreeze if not frozen!
                                        bCanUnFreeze = False
                                    End If
                                    
                                    'does this eform have an sdv
                                    bEHasSDV = oMIMData.MIMessageExists(oUser.CurrentDBConString, mimtSDVMark, mimscEForm, _
                                    oSubject.StudyDef.Name, oSubject.Site, oSubject.PersonId, lObjectId, nObjectSource, _
                                    .VisitInstance.VisitId, .VisitInstance.CycleNo, .EFormTaskId)
                                    
                                    
'                                    Call AddStringToVarArr(vJSComm, "<table><tr><a onMouseup='javascript:fnE(event.button," & Chr(34) & .EFormTaskId & Chr(34) & "," & Chr(34) & .eForm.EFormId & Chr(34) & "," & Chr(34) & .CycleNo & Chr(34) & "," & Chr(34) & .Code & Chr(34) & "," & Chr(34) & .VisitInstance.VisitId & Chr(34) & "," _
'                                            & Chr(34) & .VisitInstance.CycleNo & Chr(34) & "," & Chr(34) & .VisitInstance.Code & Chr(34) & "," & Chr(34) & .Status & Chr(34) & "," & CInt(bLocked) & "," & CInt(bFrozen) & "," & CInt(bCanUnFreeze) & ")'><td>")
'                                Else
'                                    Call AddStringToVarArr(vJSComm, "<table><tr><a onMouseup='javascript:fnE(" & Chr(34) & .EFormTaskId & Chr(34) & ",event.button)'><td>")
                                    'ic 08/09/2004 removed 'can add sdv' argument - adding an sdv has to be tried
                                    'ic 01/03/2004 added 'can add sdv argument
                                    Call AddStringToVarArr(vJSComm, msTABLE_START & "fnE(event.button," & Chr(34) & .EFormTaskId & Chr(34) & "," _
                                    & Chr(34) & .eForm.EFormId & Chr(34) & "," & Chr(34) & .CycleNo & Chr(34) & "," & Chr(34) & .Code & Chr(34) & "," _
                                    & Chr(34) & .VisitInstance.VisitId & Chr(34) & "," & Chr(34) & .VisitInstance.CycleNo & Chr(34) & "," & Chr(34) _
                                    & .VisitInstance.Code & Chr(34) & "," & Chr(34) & .Status & Chr(34) & "," & RtnJSBoolean(bLocked) & "," _
                                    & RtnJSBoolean(bFrozen) & "," & RtnJSBoolean(bCanUnFreeze) & "," _
                                    & RtnJSBoolean(oUser.CheckPermission(gsFnCreateSDV) And CanChangePlannedSDVs(oCell.eFormInst)) & "," _
                                    & RtnJSBoolean(bSHasSDV) & "," & RtnJSBoolean(bVHasSDV(nCol)) & "," & RtnJSBoolean(bEHasSDV) & ")" & msDELIMITER)
                                Else
                                    Call AddStringToVarArr(vJSComm, msTABLE_START & "fnE(" & Chr(34) & .EFormTaskId & Chr(34) & ",event.button)" & msDELIMITER)
                                End If
                                Call AddStringToVarArr(vJSComm, RtnEFormImages(.LockStatus, _
                                                              .DiscrepancyStatus, _
                                                              .SDVStatus, _
                                                              .Status))
                                'Call AddStringToVarArr(vJSComm, "</td></a></tr></table>")
                                ' DPH 18/02/2003 use GetLocalFormatDate
                                ' MLM 03/03/03: No, use eFormDateString (handles local dates)
                                Call AddStringToVarArr(vJSComm, msDELIMITER & .eFormLabel & msDELIMITER & .eFormDateString & msTABLE_END)
                                'Call AddStringToVarArr(vJSComm, msDELIMITER & .eFormLabel & msDELIMITER & GetLocalFormatDate(oUser, .eFormDate, dttDateOnly) & msTABLE_END)
                            End With

                End Select
                Call AddStringToVarArr(vJSComm, msCELL_END)

            Next
            Call AddStringToVarArr(vJSComm, "</tr>")

        Next
        Call AddStringToVarArr(vJSComm, "</table></span><table id=rowhead class=abs cellpadding=0 cellspacing=0></table><table id=colhead class=abs cellpadding=0 cellspacing=0></table><div id=blank class=clsScheduleBorder></div>")

        Call AddStringToVarArr(vJSComm, "</body>" & vbCrLf)
        '              & "</html>"

        If enInterface <> iwww Then
            Call AddStringToVarArr(vJSComm, "</html>")
        End If

    Set oMIMData = Nothing
    Set oS = Nothing
    Set oCell = Nothing
    GetSchedule = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLSchedule.GetSchedule"
End Function

'--------------------------------------------------------------
Private Function CanChangePlannedSDVs(oEFI As EFormInstance) As Boolean
'--------------------------------------------------------------
' ic 24/08/2004 copied from MACRODataManagement.modSchedule
' Can we change Planned SDVs to Done?
'--------------------------------------------------------------

    CanChangePlannedSDVs = False
    
    ' Must have some responses
    If oEFI.Status = eStatus.Requested Then Exit Function
    
    ' Mustn't have SDV status of None or Done
    If oEFI.SDVStatus = eSDVStatus.ssCancelled Then Exit Function
    If oEFI.SDVStatus = eSDVStatus.ssNone Then Exit Function
    If oEFI.SDVStatus = eSDVStatus.ssComplete Then Exit Function

    ' Otherwise there's a possiblity of some Planned SDVs on the eForm
    CanChangePlannedSDVs = True
    
End Function

'ic 30/09/2003 no longer used
''--------------------------------------------------------------------------------------------------
'Private Function RtnEFormImage(ByVal nLockStatus As Integer, _
'                               ByVal nDiscrepancyStatus As Integer, _
'                               ByVal nSDVStatus As Integer, _
'                               ByVal nStatus As Integer, _
'                               ByVal sTip As String) As String
''--------------------------------------------------------------------------------------------------
''   ic 30/09/01
''   function returns an html image element for a schedule eform
''
''   revisions
''   ic 21/10/2002   changed sdv display to show new icons instead of dashed border style
''--------------------------------------------------------------------------------------------------
'    Dim sImage As String
'    Dim sToolTip As String
'    Dim sRtn As String
'
'
'    If nSDVStatus = 40 Then
'        'queried sdv overrides all other stauses
'        sImage = "icof_sdv_query"
'        sToolTip = "Queried SDV"
'    Else
'        Select Case nLockStatus
'            'next comes locked/frozen
'            Case 5: sImage = "icof_locked"
'                    sToolTip = "Locked"
'            Case 6: sImage = "icof_frozen"
'                    sToolTip = "Frozen"
'            Case Else:
'                'then normal statuses
'                Select Case nDiscrepancyStatus
'                    Case 30: sImage = "icof_disc_raise"
'                             sToolTip = "Raised Discrepancy"
'                    Case 20: sImage = "icof_disc_resp"
'                             sToolTip = "Responded Discrepancy"
'                    Case Else:
'                        Select Case nStatus
'                            Case 0:  sImage = "icof_ok"
'                                     sToolTip = "OK"
'                            Case -5: sImage = "icof_uo"
'                                     sToolTip = "Unobtainable"
'                            Case 10: sImage = "icof_missing"
'                                     sToolTip = "Missing"
'                            Case 20: sImage = "icof_inform"
'                                     sToolTip = "Inform"
'                            Case 30: sImage = "icof_warn"
'                                     sToolTip = "Warning"
'                            Case 25: sImage = "icof_ok_warn"
'                                     sToolTip = "OK Warning"
'                            Case 40: sImage = "icof_error"
'                                     sToolTip = "Error"
'                            Case 99: sImage = "icof_inactive"
'                                     sToolTip = "Inactive"
'                            Case Else: sImage = "icof_new"
'                                       sToolTip = "Blank"
'                        End Select
'                End Select
'        End Select
'    End If
'
'    sRtn = "<img " _
'         & "border='0' " _
'         & "style='cursor:hand;' " _
'         & "src='../img/" & sImage & ".gif'"
'
'    'if eform has planned sdv we have to draw a table around the icon that we
'    'got above, to line it up with the NEW planned sdv icon which goes underneath
'    If nSDVStatus = 30 Then
'        sRtn = "<table cellpadding='0' cellspacing='0'>" _
'               & "<tr>" _
'                 & "<td>" _
'                   & sRtn & " alt='" & sToolTip & ", Planned SDV'>" _
'                 & "</td>" _
'               & "</tr>" _
'               & "<tr>" _
'                 & "<td>" _
'                   & "<img " _
'                   & "border='0'" _
'                   & "src='../img/icof_sdv_plan.gif'>" _
'                 & "</td>" _
'               & "</tr>" _
'             & "</table>"
'    Else
'        sRtn = sRtn & " alt='" & sToolTip & "'>"
'    End If
'
'    RtnEFormImage = sRtn
'End Function

'--------------------------------------------------------------------------------------------------
Private Function RtnEFormImages(ByVal nLockStatus As Integer, _
                               ByVal nDiscrepancyStatus As Integer, _
                               ByVal nSDVStatus As Integer, _
                               ByVal nStatus As Integer, _
                      Optional ByVal bCompress As Boolean = True) As String
'--------------------------------------------------------------------------------------------------
' MLM 04/02/03: Created based on RtnEFormImage. Returns a string specifying the icon and tool tip to
' be used on the schedule.
' revisions
' ic 17/02/2003 changed sdv icons
' ic 16/04/2003 renamed from RtnEformImageString()
' ic 30/09/2003 changed to pass whole image tag, otherwise IE downloads same image many times
' ic 09/10/2003 added bcompress arg
'   ic 29/06/2004 added error handling
' ic 06/09/2004 added not applicable eform icon
' NCJ 20 Jun 06 - Bug 2747 - Concatenate Lock, eForm, Discrepancy and SDV statuses into tooltip text
'--------------------------------------------------------------------------------------------------
Dim sImage As String
Dim sToolTip As String
Dim sRtn As String
    
    On Error GoTo CatchAllError
    
    sToolTip = ""
    sImage = ""
    
    Select Case nLockStatus
        ' Top of hierarchy comes locked/frozen
        Case eLockStatus.lsLocked: sImage = "icof_locked"
                sToolTip = "Locked, "
        Case eLockStatus.lsFrozen: sImage = "icof_frozen"
                sToolTip = "Frozen, "
    End Select

    ' NCJ 21 Jun 06 - Do discrepancy icon here but tooltip text later
    ' (to ensure correct icon priorities but consistent text ordering)
    Select Case nDiscrepancyStatus
        Case 30: If sImage = "" Then sImage = "icof_disc_raise"
        Case 20: If sImage = "" Then sImage = "icof_disc_resp"
    End Select

    ' NCJ 20 Jun 06 - Concatenate other statuses; only assign icon if it hasn't already got one
    Select Case nStatus
        Case eStatus.Success:  If sImage = "" Then sImage = "icof_ok"
                 sToolTip = sToolTip & "OK"
        Case eStatus.Unobtainable: If sImage = "" Then sImage = "icof_uo"
                 sToolTip = sToolTip & "Unobtainable"
        Case eStatus.Missing: If sImage = "" Then sImage = "icof_missing"
                 sToolTip = sToolTip & "Missing"
        Case eStatus.Inform: If sImage = "" Then sImage = "icof_inform"
                 sToolTip = sToolTip & "Inform"
        Case eStatus.Warning: If sImage = "" Then sImage = "icof_warn"
                 sToolTip = sToolTip & "Warning"
        Case eStatus.OKWarning: If sImage = "" Then sImage = "icof_ok_warn"
                 sToolTip = sToolTip & "OK Warning"
        Case eStatus.InvalidData: If sImage = "" Then sImage = "icof_error"
                 sToolTip = sToolTip & "Error"
        Case eStatus.NotApplicable: If sImage = "" Then sImage = "icof_na"
                 sToolTip = sToolTip & "Not Applicable"
        Case 99: If sImage = "" Then sImage = "icof_inactive"
                 sToolTip = sToolTip & "Inactive"
        Case Else: If sImage = "" Then sImage = "icof_new"
                sToolTip = sToolTip & "Blank"
    End Select
'            End Select
'    End Select
    
    ' NCJ 21 Jun 06 - Discrepancy icon already done, do tooltip text here
    Select Case nDiscrepancyStatus
        Case 30: sToolTip = sToolTip & ", Raised Discrepancy"
        Case 20: sToolTip = sToolTip & ", Responded Discrepancy"
    End Select
    
    Select Case nSDVStatus
    Case 40:
        If bCompress Then
            'add tag that will be expanded to table including underlining
            sRtn = msIMAGE_QUERIED_SDV_START & sImage & msDELIMITER & sToolTip & ", Queried SDV" & msIMAGE_END
        Else
            sRtn = "<table cellpadding=0 cellspacing=0><tr><td><img src='../img/" & sImage & ".gif' alt='" _
            & sToolTip & ", Queried SDV" & "' style='cursor:hand'></td></tr><tr><td><img src='../img/icof_sdv_query.gif'>" _
            & "</td></tr></table>"
        End If
    Case 30:
        If bCompress Then
            'add tag that will be expanded to table including underlining
            sRtn = msIMAGE_PLANNED_SDV_START & sImage & msDELIMITER & sToolTip & ", Planned SDV" & msIMAGE_END
        Else
            sRtn = "<table cellpadding=0 cellspacing=0><tr><td><img src='../img/" & sImage & ".gif' alt='" _
            & sToolTip & ", Planned SDV" & "' style='cursor:hand'></td></tr><tr><td><img src='../img/icof_sdv_plan.gif'>" _
            & "</td></tr></table>"
        End If
    Case 20:
        If bCompress Then
            'add tag that will be expanded to table including underlining
            sRtn = msIMAGE_DONE_SDV_START & sImage & msDELIMITER & sToolTip & ", Done SDV" & msIMAGE_END
        Else
            sRtn = "<table cellpadding=0 cellspacing=0><tr><td><img src='../img/" & sImage & ".gif' alt='" _
            & sToolTip & ", Done SDV" & "' style='cursor:hand'></td></tr><tr><td><img src='../img/icof_sdv_done.gif'>" _
            & "</td></tr></table>"
        End If
    Case Else
        If bCompress Then
            sRtn = msIMAGE_START & sImage & msDELIMITER & sToolTip & msIMAGE_END
        Else
            sRtn = "<img src='../img/" & sImage & ".gif' alt='" & sToolTip & "' style='cursor:hand'>"
        End If
    End Select
    
    RtnEFormImages = sRtn ' & vbCrLf
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLSchedule.RtnEformImages"
End Function



