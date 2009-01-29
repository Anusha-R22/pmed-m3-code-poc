Attribute VB_Name = "modWebBrowser"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       modWebBrowser.bas
'   Author:     Toby Aldridge Sept 2002
'   Purpose:    Constants etc for frmWebBrowser and frmBorder
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'TA 25/02/2003: Code to catch keydown events
'TA 28/02/2003: Added EnableDisableWebForms - enable disable all the mdi child web forms
'ic 13/08/2003 bug 1946, change WEB_EFORM_TOP_HEIGHT to global variable
'ic 02/10/2006  issue 2813 the about dialog should list the patch version

Option Explicit
Option Compare Text

 'TA 25/09/2002: New UI code
Public Const WEB_LH_WIDTH = 3500
Public Const WEB_APP_HEADER_LH_HEIGHT = 1300
Public Const WEB_APP_MENU_TOP_HEIGHT = 400
Public Const WEB_APP_FOOTER_LH_HEIGHT = 1200

Public Const WEB_BORDER_WIDTH = 150
Public Const WEB_BORDER_HEIGHT = 130

Public Const WEB_EFORM_LH_WIDTH = 2620
'ic 13/08/2003 bug 1946, change WEB_EFORM_TOP_HEIGHT to global variable. value is initialised in frmMenu.InitialiseMe
Public Const DEFAULT_WEB_EFORM_TOP_HEIGHT = 1120
Public WEB_EFORM_TOP_HEIGHT

'these constants are used in frmBorder
Public Const WEB_INNER_BORDER = 20
Public Const WEB_BORDER_TITLE_HEIGHT = 250

Public Const WEB_appMenuHome_URL = "_home.htm"

Private Const m_DELIMITER = "<!--r-->"

'TA 4/3/03: warning the text below can be changed by HTML editors - the text must match exactly
Private Const ROWOVER_REF_TEXT = "BEHAVIOR: url(../script/rowover_js.htc)"

'TA 26/09/2002: New enumeration for User Interface colours
Public Enum eMACROColour
    emcTitlebar = &HFF0000 ' TA I'm going for bright blue for consistancy with web: &H996633  'blue 336699   '&HFF0000 (if we want brighter blue)
    emcBackground = vbWhite '16777215 this is the colour white in UI desgn doc
    emcNonWhiteBackGround = &HFFE0E0 'the mauve colour
    emdEnabledText = &H996633 'used to be black
    emdDisabledText = 8421504 '  &H808080   grey, darker than diabled background gray
    emdDisabledBackGround = 14737632 '&H00E0E0E0& diabled control grey (lighter than the above)
    'following two are for every other row in a grid
    emdOddLineBackGround = vbWhite
    emdEvenLineBackGround = 14737632 'same as endDisabledBackground
    emcLinkText = &H996633
End Enum

Public Const MACRO_DE_FONT = "verdana"
Public Const MACRO_DE_FONT_SIZE = 8


'----------------------------------------------------------------------------------------'
Public Function AppMenuTopHTML(oUser As MACROUser) As String
'----------------------------------------------------------------------------------------'
'Return full HTML for AppMentTop
'1.includes
'2.style
'3.hover
'4.functions
'----------------------------------------------------------------------------------------'
Dim sREplace As String

    On Error GoTo ErrLabel
    
    sREplace = vbCrLf & "window.navigate('VBfnToggleLh');"
    AppMenuTopHTML = GetBaseText & vbCrLf & StyleSheetHTML & HoverButtonHTML _
        & "<script language='javascript'>" _
        & ReplaceOddSections(CompiledMenuTopJS, m_DELIMITER, Array(sREplace, sREplace)) _
        & "</script>" _
        & GetWinIO.GetAppMenuTopHTML(oUser)
                                
Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.AppMenuTopHTML"

End Function

'----------------------------------------------------------------------------------------'
Public Function AppHeaderLhHTML(bExpandButtonOnly As Boolean) As String
'----------------------------------------------------------------------------------------'
'Return full HTML for AppHeaderLh
'1.includes
'2.style
'3.functions
'4.username
'----------------------------------------------------------------------------------------'
Dim sLogoPath As String
Dim sNewFunctions As String

    On Error GoTo ErrLabel

    sLogoPath = gsDOCUMENTS_PATH & "_logo.gif"
    
    If Not FileExists(sLogoPath) Then
        sLogoPath = ""
    End If

    AppHeaderLhHTML = GetBaseText & vbCrLf & StyleSheetHTML _
                            & GetWinIO.GetAppHeaderLhHTML(goUser, sLogoPath, bExpandButtonOnly)
    
    Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.AppHeaderLhHTML"

End Function

'----------------------------------------------------------------------------------------'
Public Function AppMenuLhHTML() As String
'----------------------------------------------------------------------------------------'
'Return full HTML for AppMenuLh
'1.includes
'2.style
'3.hover (& scripts)
'4.functions
'4. PAgeLoaded stuff we don't want
'----------------------------------------------------------------------------------------'
Dim sStyleHTML As String
Dim bSiteDB As Boolean

    On Error GoTo ErrLabel
        
    'if there is a sitename then this is a site
    bSiteDB = (GetMacroDBSetting("datatransfer", "dbsitename", , "") <> "")
        AppMenuLhHTML = GetBaseText & vbCrLf & StyleSheetHTML & vbCrLf & HoverButtonHTML & _
                        "<script language='javascript'>" & CompiledRowOverJS & "</script>" & _
                        "<script language='javascript'>" & CompiledSelectListJS & "</script>" & _
                            ReplaceOddSections(GetWinIO.GetAppMenuLhHTML(goUser, bSiteDB), m_DELIMITER, Array(CompiledappMenuLhHTML))
    
    Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.AppMenuLhHTML"

End Function


'----------------------------------------------------------------------------------------'
Public Function eFormTopHTML() As String
'----------------------------------------------------------------------------------------'
'Return full HTML for AppHeaderLh
'----------------------------------------------------------------------------------------'

Dim sNewFunctions As String

    On Error GoTo ErrLabel


'    Function fnDoButtonClick(sId){eW.fnSave("v"+sId);}
'    Function fnSave(){eW.fnSave ("m3")}
'    Function fnPrint(){pW.fnPrint(eW);}
    sNewFunctions = vbCrLf & "function fnDoButtonClick(sId){window.navigate('VBfnSave|v'+sId);}" & vbCrLf _
                        & "function fnSave(sId){window.navigate('VBfnSave|'+sId);}" & vbCrLf _
                        & "function fnPrint(){window.navigate('VBfnSave|m5');}" & vbCrLf _
                        & "function fnClose(){window.navigate('VBfnSave|m4');}" & vbCrLf
                        
    eFormTopHTML = ReplaceOddSections(CompiledeFormTopHTML, m_DELIMITER, Array(GetBaseText, StyleSheetHTML, HoverButtonHTML, sNewFunctions))
    
    Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.eFormTopHTML"

End Function

'----------------------------------------------------------------------------------------'
Public Function eFormLhHTML() As String
'----------------------------------------------------------------------------------------'
'Return full HTML for AppHeaderLh
'----------------------------------------------------------------------------------------'

Dim sNewFunctions As String

    On Error GoTo ErrLabel


    sNewFunctions = vbCrLf & "function fnDoButtonClick(sId){window.navigate('VBfnSave|e'+sId);}" & vbCrLf
    eFormLhHTML = ReplaceOddSections(CompiledeFormLhHTML, m_DELIMITER, Array(GetBaseText, StyleSheetHTML, HoverButtonHTML, sNewFunctions))
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.eFormLhHTML"

End Function

'----------------------------------------------------------------------------------------'
Public Function AppFooterLhHTML() As String
'----------------------------------------------------------------------------------------'
'Return full HTML for AppFooterLh
'1.style
'----------------------------------------------------------------------------------------'


    On Error GoTo ErrLabel

    AppFooterLhHTML = ReplaceOddSections(CompiledappFooterLhHTML, m_DELIMITER, _
                            Array(GetBaseText & vbCrLf & StyleSheetHTML))
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.AppFooterLhHTML"

End Function

#If DM = 1 Then
'only for DM
'----------------------------------------------------------------------------------------'
Public Function ScheduleHTML(oSubject As StudySubject) As String
'----------------------------------------------------------------------------------------'
'Return full HTML for schedule
'1.include
'2.style
'3.function
'REVISIONS:
'   ic 07/02/2003 added user arg
' MLM 12/02/03: Include JavaScript necessary for building schedule markup from comments.
'----------------------------------------------------------------------------------------'


    On Error GoTo ErrLabel


    ScheduleHTML = GetBaseText & vbCrLf & StyleSheetHTML & _
        "<script language=javascript>" & ReplaceOddSections(CompiledScheduleJS, m_DELIMITER, Array("")) & "</script>" & _
        GetWinIO.GetScheduleHTML(oSubject, goUser)
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.ScheduleHTML"

End Function

'----------------------------------------------------------------------------------------'
Public Function AuditTrailHTML(oResponse As Response) As String
'----------------------------------------------------------------------------------------'
'Return full HTML for audittrail
'----------------------------------------------------------------------------------------'


    On Error GoTo ErrLabel

    AuditTrailHTML = GetBaseText & StyleSheetHTML & GetCloseOnKeyPressHTML & GetWinIO.GetQuestionAuditHtml(goUser, oResponse)
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.AuditTrailHTML"

End Function

'----------------------------------------------------------------------------------------'
Public Function CodingAuditTrailHTML(oResponse As Response) As String
'----------------------------------------------------------------------------------------'
'Return full HTML for audittrail
'----------------------------------------------------------------------------------------'


    On Error GoTo ErrLabel

    CodingAuditTrailHTML = GetBaseText & StyleSheetHTML & GetCloseOnKeyPressHTML & GetWinIO.GetQuestionCodingAuditHtml(goUser, oResponse)
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.CodingAuditTrailHTML"

End Function

'----------------------------------------------------------------------------------------'
Public Function QuestionDefHTML(ByRef oEformElement As eFormElementRO) As String
'----------------------------------------------------------------------------------------'
'Return full HTML for QuestionDef
'----------------------------------------------------------------------------------------'


    On Error GoTo ErrLabel

    QuestionDefHTML = GetBaseText & StyleSheetHTML & GetCloseOnKeyPressHTML & GetWinIO.GetQuestionDefinitionHTML(oEformElement)
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.QuestionDefHTML"

End Function

'----------------------------------------------------------------------------------------
Public Function MIMessageHistoryHTML(ByRef oMIMsg As Object) As String
'----------------------------------------------------------------------------------------
' MLM 01/07/03: Created. oMIMessage defined as Object because it could be a Discrepancy or SDV.
'----------------------------------------------------------------------------------------

Dim sMsgHTML As String

    On Error GoTo ErrLabel
    
    Select Case oMIMsg.MIMessageType
    Case MIMsgType.mimtDiscrepancy
        sMsgHTML = GetWinIO.GetMIMessageHistoryHTML(goUser, MIMsgType.mimtDiscrepancy, goUser.Studies.StudyByName(oMIMsg.StudyName).StudyId, oMIMsg.Site, oMIMsg.DiscrepancyID, oMIMsg.DiscrepancySource)
    Case MIMsgType.mimtSDVMark
        sMsgHTML = GetWinIO.GetMIMessageHistoryHTML(goUser, MIMsgType.mimtSDVMark, goUser.Studies.StudyByName(oMIMsg.StudyName).StudyId, oMIMsg.Site, oMIMsg.SDVID, oMIMsg.SDVSource)
    Case MIMsgType.mimtNote
        'this routine currently not used for messages, since they don't have a history,
        'but do something vaguely sensible anyway
        sMsgHTML = oMIMsg.CurrentMessage
    End Select
    
    MIMessageHistoryHTML = GetBaseText & StyleSheetHTML & GetCloseOnKeyPressHTML & sMsgHTML
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.MIMessageHistoryHTML"

End Function

#End If

'----------------------------------------------------------------------------------------'
Public Function BackGroundHTML() As String
'----------------------------------------------------------------------------------------'
'Return full HTML for audittrail
'----------------------------------------------------------------------------------------'


    On Error GoTo ErrLabel

    BackGroundHTML = GetBaseText & StyleSheetHTML & GetCloseOnKeyPressHTML & "<img height='100%' width='100%' src='../img/bg.jpg'>"
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.BackGroundHTML"

End Function

'----------------------------------------------------------------------------------------'
Public Function SubjectListHTML(oUser As MACROUser, sStudyName As String, sSite As String, sLabel As String, _
                                    sId As String, Optional sOrderBy As String = "", Optional sAscend As String = "true", Optional lBookmark As Long = 1) As String
'----------------------------------------------------------------------------------------'
'Return full HTML for subjectlist
'----------------------------------------------------------------------------------------'

Dim sStyleHTML As String


    On Error GoTo ErrLabel

    SubjectListHTML = GetBaseText & vbCrLf & StyleSheetHTML & vbCrLf & HoverButtonHTML & _
                        "<script language='javascript'>" & CompiledRowOverJS & "</script>" & _
                        GetWinIO.GetSubjectListHTML(oUser, sSite, sStudyName, sLabel, sId, sOrderBy, sAscend, lBookmark)
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.SubjectListHTML"

End Function


'----------------------------------------------------------------------------------------'
Private Function StyleSheetHTML() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'


    On Error GoTo ErrLabel

  StyleSheetHTML = "<style> " & CompiledStyleSheetHTML & " </style>"
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.StyleSheetHTML"
  
End Function

'----------------------------------------------------------------------------------------'
Private Function HoverButtonHTML() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'


    On Error GoTo ErrLabel

  HoverButtonHTML = "<script language='javascript'> " & CompiledHoverButtonHTML & " </script>"
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.HoverButtonHTML"
  
End Function


'----------------------------------------------------------------------------------------'
Public Function GetBaseText() As String
'----------------------------------------------------------------------------------------'


    On Error GoTo ErrLabel
 
    GetBaseText = "<base href='file:///" & Replace(gsWEB_HTML_LOCATION, "\", "/") & "app/AppMenuLh.htm'>" & vbCrLf
    #If DevMode = 0 Then
        'disallow right click in released version
        GetBaseText = GetBaseText & vbCrLf & GetTurnOffRightMouseHTML
    #End If
    GetBaseText = GetBaseText & vbCrLf & GetOnKeyDownHTML
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.GetBaseText"
    
End Function

'----------------------------------------------------------------------------------------'
Private Function GetTurnOffRightMouseHTML()
'----------------------------------------------------------------------------------------'


    On Error GoTo ErrLabel

    GetTurnOffRightMouseHTML = "<script language='javascript'> " _
                & "document.oncontextmenu=fnContextMenu;" & vbCrLf & "function fnContextMenu(){return false};" _
                & " </script>"
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.GetTurnOffRightMouseHTML"

End Function

'----------------------------------------------------------------------------------------'
Private Function ReplaceOddSections(sText As String, sDelim As String, vNewSections As Variant) As String
'----------------------------------------------------------------------------------------'
'Replace all the odd sections of a string with values passed in the vNewsections array
'----------------------------------------------------------------------------------------'
Dim vOldSections As Variant
Dim i As Long


    On Error GoTo ErrLabel

     vOldSections = Split(sText, sDelim)
     
    For i = 0 To UBound(vNewSections)
        vOldSections((2 * i) + 1) = vNewSections(i) 'replace odd numbered sections
    Next
    
    ReplaceOddSections = Join(vOldSections)
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.ReplaceOddSections"
    
End Function

'----------------------------------------------------------------------------------------'
Private Function GetCloseOnKeyPressHTML() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
Dim s As String


    On Error GoTo ErrLabel

    s = "<script>function keyPress() { if (window.event.keyCode==" & vbKeyReturn & "||window.event.keyCode==" & vbKeyEscape & ") {window.navigate('VBfnClose');} } " & vbCrLf
    
    'following code to allow calculation of keycodes
    's = "<script>function keyPress() { alert(window.event.keyCode);} " & vbCrLf
    
    s = s & "window.document.onkeypress = keyPress;</script>" & vbCrLf

    GetCloseOnKeyPressHTML = s
    
    Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.GetCloseOnKeyPressHTML"

End Function

'----------------------------------------------------------------------------------------'
Public Function GetAboutHTML() As String
'----------------------------------------------------------------------------------------'
' ic 02/10/2006  issue 2813 the about dialog should list the patch version
'----------------------------------------------------------------------------------------'
Dim sModuleName As String
Dim sVersionNumber As String
Dim oPatchNumber As MACROVersion.Checker
Dim sPatchNumber As String


    On Error GoTo ErrLabel

    sModuleName = GetApplicationTitle
    sVersionNumber = App.Major & "." & App.Minor & "." & App.Revision
    
    'add the patch version
    Set oPatchNumber = New MACROVersion.Checker
    sPatchNumber = oPatchNumber.PatchVersion
    If (sVersionNumber <> sPatchNumber) Then
        sVersionNumber = sVersionNumber & "<br>Patch: " & sPatchNumber
    End If
    Set oPatchNumber = Nothing

    GetAboutHTML = GetBaseText & vbCrLf & StyleSheetHTML & GetCloseOnKeyPressHTML & GetWinIO.GetAboutHTML(sModuleName, sVersionNumber)
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.GetAboutHTML"

End Function


Public Function GetWinIO() As WinIO


    On Error GoTo ErrLabel
    Set GetWinIO = New WinIO
    
    Exit Function
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.GetWinIO"

End Function

'----------------------------------------------------------------------------------------'
Private Function GetOnKeyDownHTML() As String
'----------------------------------------------------------------------------------------'
' I use the vb key constants because the ones I use match the javascript ones
'----------------------------------------------------------------------------------------'
Dim s As String


'vbKeyBack

    On Error GoTo ErrLabel

   ' s = "<script>function keyDown() {window.navigate('VBfnKeyDown|'+window.event.keyCode);if (window.event.keyCode==" & vbKeyF5 & ") {window.event.keyCode=0;} } " & vbCrLf
   'pass on keydowns and disable F5
    s = "<script>function keyDown() {window.navigate('VBfnKeyDown|'+window.event.keyCode+'|'+window.event.shiftKey+'|'+window.event.ctrlKey+'|'+window.event.altKey)"
    
    'disable F5
    s = s & ";if (window.event.keyCode==" & vbKeyF5 & ") {window.event.keyCode=0;}"
    'TA 04/06/2003: disable BackSpace
    s = s & ";if (window.event.keyCode==" & vbKeyBack & ") {window.event.keyCode=0;}"
    
    'disable arrow keys
    s = s & ";if ((window.event.keyCode==" & vbKeyLeft & "||window.event.keyCode==" & vbKeyRight & ") && (window.event.altKey==true)) {window.event.keyCode=0;}; } " & vbCrLf
    
    'following code to allow calculation of keycodes
    's = "<script>function keyDown() { alert(window.event.keyCode);} " & vbCrLf
    
    s = s & "window.document.onkeydown = keyDown;"
    
    
    s = s & "</script>" & vbCrLf

    GetOnKeyDownHTML = s
    
    Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modWeBrowser.GetOnKeyDownHTML"

End Function

'----------------------------------------------------------------------------------------'
Public Property Let WebFormsEnabled(bEnable As Boolean)
'----------------------------------------------------------------------------------------'
'TA 28/02/2003: enable disable all the mdi child web forms
'----------------------------------------------------------------------------------------'

Dim ofrm As Form

'nb option compare text is set so that case insensitive comaprsion is done
    For Each ofrm In Forms
        If ofrm.Name = "frmWebBrowser" Then
            ofrm.Enabled = bEnable
        End If
    Next
    
End Property
