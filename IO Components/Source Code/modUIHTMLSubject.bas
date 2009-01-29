Attribute VB_Name = "modUIHTMLSubject"
'----------------------------------------------------------------------------------------'
'   File:       modUIHTML.bas
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     i curtis 02/2003
'   Purpose:    functions returning html versions of MACRO pages (SUBJECT)
'----------------------------------------------------------------------------------------'
'revisions
'   ic 19/06/2003 added registration
'   ic 02/09/2003 added DLAEnable.js, bug 1992
'   ic 16/04/2004 disable page while creating subject to avoid multiple subject creation
'   ic 29/06/2004 added error handling
'   NCJ 8-19 Dec 05 - New date/time types
'   ic 02/10/2006 issue 2766, dont display studies that are suspended or closed to follow up
'   Mo 10/10/2007, Bug 2923, overflow error in GetSubjectList & RtnDelimitedSubjectList when user has access to a greater than integer number of subjects
'----------------------------------------------------------------------------------------'
Option Explicit

'--------------------------------------------------------------------------------------------------
Public Function RtnDelimitedSubjectList(ByRef oUser As MACROUser) As String
'--------------------------------------------------------------------------------------------------
'   ic 22/01/2003
'   function returns a delimited subject list to pass to 'load subject quicklist' js function
'   revisions
'   ic 02/05/2003 changed to use AddStringToVarArr()
'   ic 29/06/2004 added error handling
'   Mo 10/10/2007 Bug 2923, overflow error in RtnDelimitedSubjectList when user has access to a greater than integer number of subjects
'                 Loop variable changed from "nLoop as Integer" to lLoop as Long"
'--------------------------------------------------------------------------------------------------
Dim vData As Variant
Dim vJSComm() As String
Dim sList As String
Dim lLoop As Long

    On Error GoTo CatchAllError
    ReDim vJSComm(0)

    vData = oUser.DataLists.GetSubjectList()
    If Not IsNull(vData) Then
        For lLoop = LBound(vData, 2) To UBound(vData, 2)
            Call AddStringToVarArr(vJSComm, vData(eSubjectListCols.StudyId, lLoop) & gsDELIMITER2 _
                          & vData(eSubjectListCols.StudyName, lLoop) & gsDELIMITER2 _
                          & vData(eSubjectListCols.Site, lLoop) & gsDELIMITER2 _
                          & vData(eSubjectListCols.SubjectId, lLoop) & gsDELIMITER2 _
                          & vData(eSubjectListCols.SubjectLabel, lLoop) & gsDELIMITER1)
        Next
    End If
    sList = Join(vJSComm, "")
    If (sList <> "") Then sList = Left(sList, (Len(sList) - 1))
    RtnDelimitedSubjectList = sList
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLSubject.RtnDelimitedSubjectList"
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetNewSubject(ByRef oUser As MACROUser, _
                     Optional ByVal enInterface As eInterface = iwww) As String
'--------------------------------------------------------------------------------------------------
'   ic 13/11/2002
'   function returns an html 'new subject' choice page
'   revisions
'   ic 24/02/2003 vary height of selects depending on number of sites/studies
'   DPH 21/05/2003 Cancel button added
'   ic 19/06/2003 added registration
'   ic 02/09/2003 added DLAEnable.js, bug 1992
'   ic 16/04/2004 disable page while creating subject to avoid multiple subject creation
'   ic 29/06/2004 added error handling
'   ic 02/10/2006 issue 2766, dont display studies that are suspended or closed to follow up
'--------------------------------------------------------------------------------------------------
Dim sHTML As String
Dim sList As String
Dim oStudy As Study
Dim oSite As Site
Dim colStudy As Collection
Dim colSite As Collection
Dim nMaxStudies As Integer
Dim nMaxSites As Integer
Const nMAXSELECTHEIGHT = 15

    On Error GoTo CatchAllError

    Set colStudy = oUser.GetNewSubjectStudies

    If (colStudy.Count > 0) Then
        sHTML = sHTML & "<html>" & vbCrLf _
                      & "<head>" & vbCrLf
                      
        If (enInterface = iwww) Then
            'ic 02/09/2003 added DLAEnable.js, bug 1992
            sHTML = sHTML & "<link rel='stylesheet' HREF='../style/MACRO1.css' type='text/css'>"
            sHTML = sHTML & "<script language='javascript' src='../script/RegistrationDisable.js'></script>" _
                & "<script language='javascript' src='../script/DLAEnable.js'></script>" _
                & "<script language='javascript' src='../script/SelectList.js'></script>" & vbCrLf
            
            sHTML = sHTML & "<script language='javascript'>" & vbCrLf _
                          & "function fnPageLoaded()" & vbCrLf _
                            & "{" & vbCrLf _
                              & "document.Form1.fltSt.selectedIndex=0;" & vbCrLf _
                              & "fnChange(document.Form1.fltSi,sList,0);" & vbCrLf _
                              & "window.sWinState='0';" & vbCrLf _
                            & "}" & vbCrLf _
                          & "function fnGo()" & vbCrLf _
                            & "{" & vbCrLf _
                              & "document.Form1.btnOkNew.disabled=true;" & vbCrLf _
                              & "document.Form1.btnCancel.disabled=true;" & vbCrLf _
                              & "document.Form1.fltSt.disabled=true;" & vbCrLf _
                              & "document.Form1.fltSi.disabled=true;" & vbCrLf _
                              & "window.parent.fnCreateSubjectUrl(document.Form1.fltSt.value,document.Form1.fltSi.value);" & vbCrLf _
                            & "}" & vbCrLf _
                          & "function fnCancel()" & vbCrLf _
                            & "{" & vbCrLf _
                              & "window.parent.fnHomeUrl();" & vbCrLf _
                            & "}" & vbCrLf _
                          & "</script>" & vbCrLf
        End If
    

        sHTML = sHTML & "</head>" & vbCrLf _
                      & "<body onload='fnPageLoaded();'>" & vbCrLf _
                      & "<table width='100%' height='100%'><tr><td align='center' valign='middle'>" _
                        & "<form name='Form1' method='get' action='CreateSubject.asp'>" _
                        & "<table width='300' align='center'>" _
                          & "<tr height='20'><td class='clsTableHeaderText'>&nbsp;Study</td></tr>" _
                          & "<tr><td>"
                         
        nMaxStudies = IIf((colStudy.Count + 1) > nMAXSELECTHEIGHT, nMAXSELECTHEIGHT, (colStudy.Count + 1))
        sHTML = sHTML & "<select style='width:250px;' class='clsSelectList' size='" & nMaxStudies & "' name='fltSt' onchange='fnChange(fltSi,sList,this.selectedIndex);'>"
        
        For Each oStudy In colStudy
            'bug 2766 dont display studies that are suspended or closed to follow up
            If (oStudy.StatusId <> 4) And (oStudy.StatusId <> 5) Then
                Set colSite = oUser.GetNewSubjectSites(oStudy.StudyId)
            
                If (colSite.Count > 0) Then
                    sHTML = sHTML & "<option value='" & oStudy.StudyId & "'>" & oStudy.StudyName & "</option>"
                    
                    If colSite.Count > nMaxSites Then nMaxSites = colSite.Count
                    For Each oSite In colSite
                        sList = sList & oSite.Site & gsDELIMITER2
                    Next
                    'remove trailing minor delimiter, add major delimiter
                    sList = Left(sList, (Len(sList) - 1)) & gsDELIMITER1
                End If
            End If
        Next
        nMaxSites = IIf((nMaxSites + 1) > nMAXSELECTHEIGHT, nMAXSELECTHEIGHT, (nMaxSites + 1))
        
        'remove trailing major delimiter, if any
        If sList <> "" Then
            sList = Left(sList, (Len(sList) - 1))
        
            sHTML = sHTML & "</select>" _
                          & "</td></tr>" _
                          & "<tr height='20'><td class='clsTableHeaderText'>&nbsp;Site</td></tr>" _
                          & "<tr><td>"
                          
            ' DPH 21/05/2003 - Added Cancel Button
            sHTML = sHTML & "<select style='width:250px;' class='clsSelectList' size='" & nMaxSites & "' name='fltSi'></select>" _
                          & "</td></tr>" _
                          & "<tr height='10'><td></td></tr>" _
                          & "<tr><td>" _
                          & "<table width='80%'><tr><td align='center'>" _
                          & "<input style='width:100px;' class='clsButton' type='button' value='OK' name='btnOkNew' onclick='fnGo();'>" _
                          & "</td><td align='center'>" _
                          & "<input style='width:100px;' class='clsButton' type='button' value='Cancel' name='btnCancel' onclick='fnCancel();'>" _
                          & "</td></tr></table>" _
                          & "</td></tr>" _
                          & "<tr height='10'><td></td></tr>" _
                          & "</table>" _
                          & "</form>" _
                          & "</td></tr></table>" & vbCrLf _
                          & "</body></html>" _
                          & "<script language='javascript'>" & vbCrLf _
                          & "var sList=" & Chr(34) & sList & Chr(34) & ";" & vbCrLf _
                          & "</script>"
        Else
            Call Err.Raise(vbObjectError + 1, , "User does not have data-entry or data-review permission")
        End If
    Else
        Call Err.Raise(vbObjectError + 1, , "User does not have data-entry or data-review permission")
    End If
    
    Set oStudy = Nothing
    Set oSite = Nothing
    Set colStudy = Nothing
    Set colSite = Nothing
    
    GetNewSubject = sHTML
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLSubject.GetNewSubject"
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetSubjectList(ByRef oUser As MACROUser, _
                      Optional ByVal sSite As String = "", _
                      Optional ByVal sStudy As String = "", _
                      Optional ByVal sLabel As String = "", _
                      Optional ByVal lId As Long = -1, _
                      Optional ByVal bViewInformIcon As Boolean = False, _
                      Optional ByVal enOrderBy As eSubjectListCols = -1, _
                      Optional ByVal bAscend As Boolean = True, _
                      Optional ByVal enInterface As eInterface = iwww, _
                      Optional ByVal nBookmark As Integer = 0) As String
'--------------------------------------------------------------------------------------------------
'   ic 30/10/02
'   builds and returns an html page representing a subject list
'   REVISIONS
'   DPH 06/01/2003 - Pagelength required for bookmarking so moved pagelength calculation code
'                    Do not "order" when paging as is stored in user state
'   ic 01/04/2003   added last modified, status column
'   ic 19/06/2003 added registration
'   ic 02/09/2003 added DLAEnable.js, bug 1992
'   ic 29/06/2004 added error handling
'   Mo 10/10/2007 Bug 2923, overflow error in GetSubjectList when user has access to a greater than integer number of subjects
'                 Loop variable changed from "nLoop as Integer" to lLoop as Long"
'                 Loop variable changed from "nStart as Integer" to lStart as Long"
'                 Loop variable changed from "nStop as Integer" to lStop as Long"
'--------------------------------------------------------------------------------------------------

Dim vData As Variant
Dim lStart As Long
Dim lStop As Long
Dim lLoop As Long
Dim nPageLength As Long
Dim sDatabase As String
Dim vJSComm() As String
Dim sTimestamp As String

    On Error GoTo CatchAllError
    ReDim vJSComm(0)
    
    sDatabase = oUser.DatabaseCode
    
    vData = oUser.DataLists.GetSubjectList(sLabel, sStudy, sSite, lId, enOrderBy, bAscend)

    Call AddStringToVarArr(vJSComm, "<html>" & vbCrLf _
                  & "<head>" & vbCrLf)
                  
    'DPH 06/01/2003 - Need pagelength earlier within code
    If (enInterface = iwww) Then
        nPageLength = oUser.UserSettings.GetSetting(SETTING_PAGE_LENGTH, 50)
    Else
        'windows want a single page
        If (Not IsNull(vData)) Then
            nPageLength = UBound(vData, 2)
        Else
            nPageLength = 0
        End If
    End If
    
    If (enInterface = iwww) Then
        'ic 02/09/2003 added DLAEnable.js, bug 1992
        'DPH 06/01/2003 - Do not "order" when paging as is stored in user state
        'add stylesheet, js functions for www
        Call AddStringToVarArr(vJSComm, "<link rel='stylesheet' href='../style/MACRO1.css' type='text/css'>" _
                        & "<script language='javascript' src='../script/RegistrationDisable.js'></script>" _
                        & "<script language='javascript' src='../script/DLAEnable.js'></script>" _
                        & "<script language='javascript' src='../script/RowOver.js'></script>" _
                      & "<script language='javascript'>" & vbCrLf _
                        & "function fnGoS(sSt,sSi,sSj)" & vbCrLf _
                          & "{window.parent.fnScheduleUrl(sSt,sSi,sSj," & CInt(CBool(oUser.UserSettings.GetSetting(SETTING_SAME_EFORM, "false"))) & ");}" & vbCrLf _
                        & "function fnNav(bForward)" & vbCrLf _
                          & "{" _
                            & "if (bForward)" _
                            & "{" _
                              & "window.parent.fnSubjectListUrl('" & sStudy & "','" & sSite & "','" & ReplaceWithJSChars(sLabel) & "','-1','" & (nBookmark + nPageLength) & "')" _
                            & "}" _
                            & "else" _
                            & "{" _
                              & "window.parent.fnSubjectListUrl('" & sStudy & "','" & sSite & "','" & ReplaceWithJSChars(sLabel) & "','-1','" & (nBookmark - nPageLength) & "')" _
                            & "}" _
                          & "}" & vbCrLf _
                      & "function fnOrder(nCol){window.parent.fnSubjectListUrl('" & sStudy & "','" & sSite & "','" & ReplaceWithJSChars(sLabel) & "',nCol,'" & nBookmark & "')}" & vbCrLf _
                      & "function fnPageLoaded()" & vbCrLf _
                        & "{window.sWinState='1" & gsDELIMITER2 & sStudy & gsDELIMITER2 & sSite & gsDELIMITER2 & ReplaceWithJSChars(sLabel) & gsDELIMITER2 & enOrderBy & gsDELIMITER2 & nBookmark & "';}" & vbCrLf _
                      & "</script>")
                      
        Call AddStringToVarArr(vJSComm, "</head>" & vbCrLf _
                  & "<body onload='fnPageLoaded();'>" & vbCrLf)
                  
    Else
                      Call AddStringToVarArr(vJSComm, "<script language='javascript'>" & vbCrLf)
                            
                        'TA windows function
                            Call AddStringToVarArr(vJSComm, "function fnGoS(sSt,sSi,sSj){window.navigate('VBfnScheduleUrl|'+sSt+'|'+sSi+'|'+sSj);}" & vbCrLf)
                        
                        Call AddStringToVarArr(vJSComm, "function fnNav(bForward)" & vbCrLf _
                          & "{" _
                            & "if (bForward)" _
                            & "{" _
                              & "window.parent.fnSubjectListUrl('" & sStudy & "','" & sSite & "','" & sLabel & "','" & enOrderBy & "','" & (nBookmark + nPageLength) & "')" _
                            & "}" _
                            & "else" _
                            & "{" _
                              & "window.parent.fnSubjectListUrl('" & sStudy & "','" & sSite & "','" & sLabel & "','" & enOrderBy & "','" & (nBookmark - nPageLength) & "')" _
                            & "}" _
                          & "}" & vbCrLf _
                      & "function fnOrder(nCol){window.navigate('VBfnSubjectListUrl|'+'" & sStudy & "'+'|'+'" & sSite & "'+'|'+'" & sLabel & "'+'|'+nCol);}" _
                      & "</script>")
                      
        Call AddStringToVarArr(vJSComm, "</head>" & vbCrLf _
                  & "<body>" & vbCrLf)
    End If
    
    
    If (Not IsNull(vData)) Then
    
        Call AddStringToVarArr(vJSComm, "<table align='center' width='500'>" _
                      & "<tr><td>")
        
        'calculate the start row and end row based on start row (bookmark) and page length
        If ((nBookmark) >= UBound(vData, 2) Or (nBookmark) < 0) Then
            lStart = 0
        Else
            lStart = nBookmark
        End If
        
        If ((lStart + nPageLength) >= UBound(vData, 2)) Then
            lStop = UBound(vData, 2)
        Else
            lStop = (lStart + nPageLength) - 1
        End If
        
        
        Call AddStringToVarArr(vJSComm, "<table align='center' width='100%'>" & vbCrLf _
                      & "<tr class='clsLabelText' height='30'>" & vbCrLf _
                      & "<td align='right'>Record(s) " & lStart + 1 & " to " & lStop + 1 & " of " & UBound(vData, 2) + 1 & "&nbsp;&nbsp;")
    
        If (enInterface = iwww) Then
            'omly one page in windows
            'previous page icon
            If (lStart > 0) Then
                Call AddStringToVarArr(vJSComm, "<a href='javascript:fnNav(0);'>" _
                              & "<img src='../img/ico_backon.gif' border='0' alt='previous page'></a>&nbsp;")
            Else
                Call AddStringToVarArr(vJSComm, "<img src='../img/ico_back.gif'>&nbsp;")
            End If
            
            'next page icon
            If (lStop < UBound(vData, 2)) Then
                Call AddStringToVarArr(vJSComm, "<a href='javascript:fnNav(1);'>" _
                              & "<img src='../img/ico_forwardon.gif' border='0' alt='next page'></a>&nbsp;")
            Else
                Call AddStringToVarArr(vJSComm, "<img src='../img/ico_forward.gif'>&nbsp;")
            End If
        End If
        
        Call AddStringToVarArr(vJSComm, "</td></tr>" _
                      & "</table>")
                      
        
        Call AddStringToVarArr(vJSComm, "</td></tr>" _
                      & "</table>")
                      
                
        'subject table header
        Call AddStringToVarArr(vJSComm, "<table onmouseover='fnOnMouseOver(this,1);' onmouseout='fnOnMouseOut(this);' id='tsubject' style='cursor:hand'  align='center' width='80%' cellpadding='0' cellspacing='0' border='0' bordercolor='#c0c0c0'>" _
                        & "<tr class='clsTableHeaderText' height='20' border='1'>" _
                          & "<td width='5%' style='cursor:default;'>&nbsp;</td>" _
                          & "<a onclick='javascript:fnOrder(" & eSubjectListCols.StudyName & ");'><td width='20%' title='Sort by study'>&nbsp;Study</td></a>" _
                          & "<a onclick='javascript:fnOrder(" & eSubjectListCols.Site & ");'><td width='15%' title='Sort by site'>&nbsp;Site</td></a>" _
                          & "<a onclick='javascript:fnOrder(" & eSubjectListCols.SubjectId & ");'><td width='10%' title='Sort by subject'>&nbsp;Subject</td></a>" _
                          & "<a onclick='javascript:fnOrder(" & eSubjectListCols.SubjectLabel & ");'><td width='20%' title='Sort by label'>&nbsp;Label</td></a>" _
                          & "<a onclick='javascript:fnOrder(" & eSubjectListCols.SubjectTimeStamp & ");'><td width='30%' title='Sort by timestamp'>&nbsp;Last modified</td></a>" _
                        & "</tr>")
        
        
        'subject rows
        For lLoop = lStart To lStop
            'striping
            If ((lLoop Mod 2) = 0) Then
                Call AddStringToVarArr(vJSComm, "<tr height='17' class='clsTableText'>")
            Else
                Call AddStringToVarArr(vJSComm, "<tr height='17' class='clsTableTextS'>")
            End If
            If Not IsNull(vData(eSubjectListCols.SubjectTimeStamp, lLoop)) Then
                sTimestamp = GetLocalFormatDate(oUser, CDate(vData(eSubjectListCols.SubjectTimeStamp, lLoop)), eDateTimeType.dttDMYT)
                sTimestamp = sTimestamp & IIf(Not IsNull(vData(eSubjectListCols.SubjectTimeStampTZ, lLoop)), "&nbsp;" & RtnDifferenceFromGMT(vData(eSubjectListCols.SubjectTimeStampTZ, lLoop)), "")
            Else
                sTimestamp = "&nbsp;"
            End If
            Call AddStringToVarArr(vJSComm, "<a href='javascript:fnGoS(" & Chr(34) & vData(eSubjectListCols.StudyId, lLoop) & Chr(34) & "," & Chr(34) & vData(eSubjectListCols.Site, lLoop) & Chr(34) & "," & Chr(34) & vData(eSubjectListCols.SubjectId, lLoop) & Chr(34) & ")'>" _
                            & "<td>" & RtnStatusImages(vData(eSubjectListCols.SubjectStatus, lLoop), oUser.CheckPermission(gsFnMonitorDataReviewData), vData(eSubjectListCols.LockStatus, lLoop), False, vData(eSubjectListCols.SDVStatus, lLoop), vData(eSubjectListCols.DiscStatus, lLoop)) & "</td>" _
                            & "<td>" & vData(eSubjectListCols.StudyName, lLoop) & "&nbsp;</td>" _
                            & "<td>" & vData(eSubjectListCols.Site, lLoop) & "&nbsp;</td>" _
                            & "<td>" & vData(eSubjectListCols.SubjectId, lLoop) & "&nbsp;</td>" _
                            & "<td>" & vData(eSubjectListCols.SubjectLabel, lLoop) & "&nbsp;</td>" _
                            & "<td>" & sTimestamp & "</td>" _
                          & "</a></tr>")
        Next


        Call AddStringToVarArr(vJSComm, "</td></tr>" _
                      & "</table>")
    Else
    
        Call AddStringToVarArr(vJSComm, "<table>" _
                      & "<tr class='clsMessageText'><td>" _
                      & "No matching subjects found" _
                      & "</td></tr>" _
                      & "</table>")
    End If
    
    
    Call AddStringToVarArr(vJSComm, "</body>" _
                  & "</html>")
    
    GetSubjectList = Join(vJSComm, "")
    Exit Function
    
CatchAllError:
    Err.Raise Err.Number, , Err.Description & "|modUIHTMLSubject.GetNewSubject"
End Function

