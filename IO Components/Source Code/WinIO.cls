VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   File:       WinIO.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Toby Aldridge, Septmember 2002
'   Purpose:    Class to retrieve HTML for certain forms
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' REVISIONS:
' ic 26/06/2003 added decimalpoint, thousandseparator arguments, bug 1873
' ic 28/07/2005 added clinical coding
Option Explicit

'--------------------------------------------------------------------------------------------------
Public Function GetQuestionCodingAuditHtml(ByRef ouser As MACROUser, oResponse As Response)
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    With oResponse.EFormInstance.VisitInstance.Subject
        GetQuestionCodingAuditHtml = GetQuestionCodingAudit(ouser, .StudyId, .Site, .PersonId, oResponse.ResponseId, _
            oResponse.RepeatNumber, oResponse.Element.Code, iWindows)
    
    End With
    
Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "WinIO.GetQuestionCodingAuditHtml", Err.Description & "|" & "WinIO.GetQuestionCodingAuditHtml"
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetRFCHTML(ByRef oStudyDef As StudyDefRO) As String
'--------------------------------------------------------------------------------------------------
'   ic 30/10/2002
'   wrapper for modUIHTML.GetRFC
'--------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    GetRFCHTML = GetRFC(oStudyDef)
    
    
Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "WinIO.GetRFCHTML", Err.Description & "|" & "WinIO.GetRFCHTML"
    
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetSubjectListHTML(ouser As MACROUser, _
                                   ByVal sSite As String, _
                                   ByVal sStudy As String, _
                                   ByVal sLabel As String, _
                                   ByVal sId As String, _
                                   ByVal sOrderBy As String, _
                                   ByVal sAscend As String, _
                                   ByVal sBookmark As String) As String
'--------------------------------------------------------------------------------------------------
'   ic 01/11/2002
'   function wrapper for moduihtml function
' DPH 08/11/2002 Changed to use Serialised User object
'--------------------------------------------------------------------------------------------------
Dim lId As Long
Dim bAscend As Boolean

    If (ouser.CheckPermission(gsFnDataEntry)) Or (ouser.CheckPermission(gsFnDataReview)) Then
        If IsNumeric(sId) Then
            lId = Val(sId)
        Else
            lId = -1
        End If
        
        If sOrderBy = "" Then sOrderBy = "-1"
        bAscend = (sAscend = "true")
        GetSubjectListHTML = GetSubjectList(ouser, sSite, sStudy, sLabel, lId, True, sOrderBy, bAscend, iWindows, CInt(sBookmark))
    
    Else
        GetSubjectListHTML = GetErrorHTML(ePermission, gsFnDataEntry & " " & gsFnDataReview, , "WinIO.GetSubjectListHTML", iwww)
        
    End If

End Function


'--------------------------------------------------------------------------------------------------
Public Function GetQuestionDefinitionHTML(ByRef oEformElement As eFormElementRO) As String
'--------------------------------------------------------------------------------------------------
'   ic 29/10/2002
'   wrapper for modUIHTML.GetQUestionDefinition
'--------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    GetQuestionDefinitionHTML = GetQuestionDefinition(oEformElement, iWindows)
    
    
Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "WinIO.GetQuestionDefinitionHTML", Err.Description & "|" & "WinIO.GetQuestionDefinitionHTML"
    
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetMIMessageHistoryHTML(ByRef ouser As MACROUser, ByRef enType As MIMsgType, ByRef lStudy As Long, _
                ByRef sSite As String, ByRef lId As Long, ByRef enSource As MIMsgSource) As String
'--------------------------------------------------------------------------------------------------
' MLM 01/07/03: Created. Wrapper for modUIHTMLMIMessage.GetMIMessageAudit
'--------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    GetMIMessageHistoryHTML = GetMIMessageAudit(ouser, enType, lStudy, sSite, lId, enSource, False)
    
Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "WinIO.GetMIMessageHistoryHTML", Err.Description & "|" & "WinIO.GetMIMessageHistoryHTML"
    
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetAppMenuLhHTML(ByRef ouser As MACROUser, bSiteDB As Boolean) As String
'--------------------------------------------------------------------------------------------------
'   ta 11/07/01
'   wrapper fro modUIHTML.GetAppmenuLh
'--------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    GetAppMenuLhHTML = GetAppMenuLh(ouser, iWindows, , , bSiteDB)
    
    
Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "WinIO.GetAppMenuLhHTML", Err.Description & "|" & "WinIO.GetAppMenuLhHTML"
    
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetAppMenuTopHTML(ouser As MACROUser) As String
'--------------------------------------------------------------------------------------------------
'   ic 06/03/2003
'   wrapper for modUIHTMLApplication.GetAppMenuTop
'--------------------------------------------------------------------------------------------------
                         
    GetAppMenuTopHTML = GetAppMenuTop(ouser, iWindows)
    
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetScheduleHTML(ByRef oSubject As StudySubject, ByRef ouser As MACROUser) As String
'--------------------------------------------------------------------------------------------------
'   ta 11/07/01
'   wrapper fro modUIHTML.GetSchedule
'   REVISIONS:
'   ic 07/01/2003 added user object arg
'--------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    'TA 30/11/2003: always use uncompressed version in windows - fractionally faster
    GetScheduleHTML = GetScheduleNoCompression(oSubject, iWindows, ouser)
    
    
Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "WinIO.GetScheduleHTML", Err.Description & "|" & "WinIO.GetScheduleHTML"
    
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetQuestionAuditHtml(ByRef ouser As MACROUser, oResponse As Response) As String
'--------------------------------------------------------------------------------------------------
'   ta 16/10/02
'   builds and returns an html table representing a question audit trail
'   REVISIONS:
'   ic 29/10/2002 changed arguement from elementid to questionid
'   ic 26/06/2003 added decimalpoint, thousandseparator arguments, bug 1873
'--------------------------------------------------------------------------------------------------


    With oResponse.EFormInstance.VisitInstance.Subject
        GetQuestionAuditHtml = GetQuestionAudit(ouser, .StudyId, .Site, .PersonId, _
                                                    oResponse.EFormInstance.EFormTaskId, _
                                                    oResponse.Element.QuestionId, oResponse.Element.Code, _
                                                    RegionalDecimalPointChar, RegionalThousandSeparatorChar, _
                                                    oResponse.RepeatNumber, iWindows)
    End With
    
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetAppHeaderLhHTML(ByRef ouser As MACROUser, _
                                   ByVal sLogoPath As String, bExpandButtonOnly As Boolean) As String
'--------------------------------------------------------------------------------------------------
'   ta 06/11/02
'   builds and returns an html table representing the app ment header
'   REVISIONS:
'   ic 29/10/2002 changed arguement from elementid to questionid
'   ic 11/11/2002 changed to pass user object
'--------------------------------------------------------------------------------------------------


    GetAppHeaderLhHTML = GetAppHeaderLh(ouser, sLogoPath, eInterface.iWindows)

    
End Function

Public Function GetDelimitedEformList(ByRef oSubject As StudySubject, _
                                       ByVal lVisitTaskId As Long) As String
'-------------------------------------------------------------------------------------------'
'   ic 07/10/2002
'   function returns a string containing a list of eforms in the passed visit instance
'   list is delimited and contains groups of eform code,eform pagetaskid,eform name
'   parameters
'-------------------------------------------------------------------------------------------'


    GetDelimitedEformList = RtnDelimitedEformList(oSubject.ScheduleGrid, lVisitTaskId)
    
End Function



'-------------------------------------------------------------------------------------------'
Public Function GetDelimitedVisitList(ByRef oSubject As StudySubject) As String
'-------------------------------------------------------------------------------------------'
'   ic 07/10/2002
'   function returns a string containing a list of visits in the passed schedule
'   list is delimited and contains groups of visit code,visitid,visit name parameters
'-------------------------------------------------------------------------------------------'

    GetDelimitedVisitList = RtnDelimitedVisitList(oSubject.ScheduleGrid)


End Function


'-------------------------------------------------------------------------------------------'
Private Function AddHeaderAndFooter(sBody As String)
'-------------------------------------------------------------------------------------------'
' add header and a footer to body text
'<!--r--><!--r--> is place holder for style sheet
'-------------------------------------------------------------------------------------------'
Dim sHTML As String

    sHTML = "<html><head><!--r--><!--r--></head><body>"
    sHTML = sHTML & sBody
    sHTML = sHTML & "</body></html>"
    
End Function



'--------------------------------------------------------------------------------------------------
Public Function GetDataBrowser(ouser As MACROUser, _
                               bCountOnly As Boolean, _
                               ByVal sStudyCode As String, _
                               ByVal sSiteCode As String, _
                               ByVal sVisitCode As String, _
                               ByVal sCRFPageId As String, _
                               ByVal sQuestion As String, _
                               ByVal sSrchUserName As String, _
                               ByVal sSubjectId As String, _
                               ByVal sSubjectLabel As String, _
                               ByVal sStatus As String, _
                               ByVal sLockStatus As String, _
                               ByVal sBefore As String, _
                               ByVal sTime As String, _
                               ByVal sComment As String, _
                               ByVal sDiscrepancy As String, _
                               ByVal sSDV As String, _
                               ByVal sNote As String, _
                               ByVal sCodingStatus As String, _
                               ByVal sDictionaryName As String, _
                               ByVal sDictionaryVersion As String, _
                               ByVal sGet As String, _
                               Optional bConvertDates As Boolean = True, _
                               Optional ByRef vErrors As Variant) As Variant
'--------------------------------------------------------------------------------------------------
' revisions
' ic 23/01/2003 changed to use modUIHTML.RtnDataBrowser function
' MLM 25/03/03: Bug 1292: Switched sTime and sBefore arguments.
'               Added vErrors argument; pass through to RtnDataBrowser.
' ic 03/09/2003 check for sStudycode empty string
' ic 28/07/2005 added clinical coding
'--------------------------------------------------------------------------------------------------
Dim lSubjectId As Long

    'check passed argument values
    If Trim(sStudyCode) = "" Then sStudyCode = "0"
    If Trim(sSiteCode) = "" Then sSiteCode = "ALL"
    If Trim(sVisitCode) = "" Then sVisitCode = "0"
    If Trim(sCRFPageId) = "" Then sCRFPageId = "0"
    If Trim(sQuestion) = "" Then sQuestion = "0"
    If Trim(sSrchUserName) = "" Then sSrchUserName = "ALL"
    If IsNumeric(sSubjectId) Then
        If (CLng(sSubjectId) > 0) Then
            lSubjectId = CLng(sSubjectId)
            sSubjectLabel = ""
        Else
            lSubjectId = 0
        End If
    Else
        lSubjectId = 0
    End If
    If LCase(sBefore) <> "true" Then sBefore = "false"
    If Not IsNumeric(sComment) Then sComment = "-1"
    If Not IsNumeric(sDiscrepancy) Then sDiscrepancy = "-1"
    If Not IsNumeric(sSDV) Then sSDV = "-1"
    If Not IsNumeric(sNote) Then sNote = "-1"
    
    'ic 28/07/2005 added clinical coding
    If Not IsNumeric(sCodingStatus) Then sCodingStatus = "-1"
    
    If (sGet <> "0") And (sGet <> "1") And (sGet <> "2") Then sGet = "1"

    'ic 28/07/2005 added clinical coding
    GetDataBrowser = RtnDataBrowser(ouser, _
                               bCountOnly, _
                                CInt(sStudyCode), _
                                sSiteCode, _
                                CLng(sVisitCode), _
                                CLng(sCRFPageId), _
                                CLng(sQuestion), _
                                sSrchUserName, _
                                lSubjectId, _
                                sSubjectLabel, _
                                sStatus, _
                                sLockStatus, _
                                sTime, _
                                CBool(sBefore), _
                                CLng(sComment), _
                                CLng(sDiscrepancy), _
                                CLng(sSDV), _
                                CLng(sNote), _
                                CLng(sCodingStatus), _
                                sDictionaryName, _
                                sDictionaryVersion, _
                                sGet, bConvertDates, _
                                vErrors)
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetMIMessageList(ByVal sType As String, _
                                 ouser As MACROUser, _
                                 ByVal sStudyCode As String, _
                                 ByVal sSiteCode As String, _
                                 ByVal sVisitCode As String, _
                                 ByVal sVisitCycle As String, _
                                 ByVal sCRFPageId As String, _
                                 ByVal sCRFPageCycle As String, _
                                 ByVal sQuestion As String, _
                                 ByVal sQuestionCycle As String, _
                                 ByVal sSrchUserName As String, _
                                 ByVal sSubjectLabel As String, _
                                 ByVal sSubjectId As String, _
                                 ByVal sStatus As String, _
                                 ByVal sTime As String, _
                                 ByVal sBefore As String, _
                                 ByVal sScope As String) As Variant
'---------------------------------------------------------------------------------------------------
' MLM 30/06/05: Pass through visit, eForm and question cycle numbers.
'---------------------------------------------------------------------------------------------------

    GetMIMessageList = RtnMIMessageList(ouser, _
                                 sType, _
                                 sStudyCode, _
                                 sSiteCode, _
                                 sVisitCode, sVisitCycle, _
                                 sCRFPageId, sCRFPageCycle, _
                                 sQuestion, sQuestionCycle, _
                                 sSrchUserName, _
                                 sSubjectId, _
                                 sSubjectLabel, _
                                 sStatus, _
                                 sTime, _
                                 sBefore, _
                                 sScope)

End Function


'-----------------------------------------------------------------------------
Public Function GetAboutHTML(sModuleName As String, sVersionNumber As String) As String
'-----------------------------------------------------------------------------
'get html for about screen
'-----------------------------------------------------------------------------

    GetAboutHTML = GetAbout(sModuleName, sVersionNumber, iWindows)

End Function


'--------------------------------------------------------------------------------------------------
Public Function GetDelimitedSubjectList(ByRef ouser As MACROUser) As String
'--------------------------------------------------------------------------------------------------
'   ic 22/01/2003
'   function returns a delimited subject list to pass to 'load subject quicklist' js function
'--------------------------------------------------------------------------------------------------

    GetDelimitedSubjectList = RtnDelimitedSubjectList(ouser)


End Function

