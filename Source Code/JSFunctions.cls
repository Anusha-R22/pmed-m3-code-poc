VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       JSFunctions.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Toby Aldridge, September 2002
'   Purpose:    Map javascript functions from frmWebBrowser
'           Each Function shows the corresponding required javascript function
'----------------------------------------------------------------------------------------'
' REVISIONS:
'MLM 21/01/03: Fixed fnNoteUrl so that Note searches work with subject labels.
'TA 24/01/2003: Changed error handlers
'TA 12/02/2003: In data browser, if not monitor then give them chooice of subjects if more than one subject label matches
'RS 18/02/2003: Added functions for Communication Panel menu
'TA 25/02/2003: Code to catch keydown events
'   NCJ 20 Mar 03 - Data Browser display can now take Null data array
'  TA 21/03/2003: changed mofrmWebBrowser type to form becasue it may be frmWebBrowser or frmWebBrowserNonMDI
' ic 23/06/2003 added bInit to fnLoadSelect(), LoadLHCombos()
' DPH 27/10/2003 - Catch 'out of memory' (7) and 'memory allocation to flexgrid' (30006)  errors in fnBrowserUrl
' ic 28/07/2005 added clinical coding
'----------------------------------------------------------------------------------------'

Option Explicit

'for use as index of vParams
Public Enum eFunctionParams
    fpFunctionName = 0
    fpDatabase = 1
End Enum

Private mofrmWebBrowser As Form  'could be frmWebBrowser or frmWebBrowserNonMDI

'------------------------------------------------------------------------------'
Public Function RtnJSBoolean(ByVal bVal As Boolean) As String
'------------------------------------------------------------------------------'
'   ic 28/04/2003
'   function returns a javascript value representing the passed vb boolean
'------------------------------------------------------------------------------'
    RtnJSBoolean = IIf(bVal, "1", "0")
End Function

'----------------------------------------------------------------------------------------'
Public Sub Init(ofrmWebBrowser As Form)
'----------------------------------------------------------------------------------------'
'initalise js to know about the form that called it
'----------------------------------------------------------------------------------------'

On Error GoTo ErrLabel

    Set mofrmWebBrowser = ofrmWebBrowser
    
    Exit Sub
    
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, "Init", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

Public Sub LoadLhCombos(ofrmMenuLh As frmWebBrowser, lStudyId As Long)
'---------------------------------------------------------------------
Dim oSite As Site
Dim sList As String
Dim vList As Variant
Dim i As Long
Dim oStudyDef As StudyDefRO
Dim oVisit As VisitRO
Dim oEForm As eFormRO
Dim oQuestion As eFormElementRO
Dim vData As Variant
'//      fnLoadSelect(fltVi,lstVisits);
'//      fnLoadSelect(fltEf,lstForms);
'//      fnLoadSelect(fltQu,lstQuestions);
'//      fnLoadSelect(fltSi,lstSites);
'//      fnLoadSelect(fltUs,lstUsers);


    On Error GoTo ErrLabel


'load sites
'Site`SiteDescription
    sList = "|`All sites"
    For Each oSite In goUser.GetAllSites(lStudyId)
        sList = sList & "|" & oSite.Site & "`" & oSite.SiteDescription
    Next
    Call fnLoadSelect(ofrmMenuLh, "fltSi", sList, False)

'Load visit combo
'visitid`name`StudyId`code
    sList = "|`All visits"
    vData = goUser.DataLists.GetVisitList(lStudyId)
    If Not IsNull(vData) Then
        ReDim vList(UBound(vData, 2)) As String
        For i = 0 To UBound(vData, 2)

            vList(i) = "|" & vData(eDropDownCol.Id, i) & "`" & vData(eDropDownCol.Text, i) ' & "`" & lStudyId
        Next
        sList = sList & Join(vList)
    End If
     
    Call fnLoadSelect(ofrmMenuLh, "fltVi", sList, False)
    
'Load eForm combo
'eformid`name`studyid`code
    
    sList = "|`All eForms"
    vData = goUser.DataLists.GetEFormList(lStudyId)
    If Not IsNull(vData) Then
        ReDim vList(UBound(vData, 2)) As String
        For i = 0 To UBound(vData, 2)

            vList(i) = "|" & vData(eDropDownCol.Id, i) & "`" & vData(eDropDownCol.Text, i) ' & "`" & lStudyId
        Next
        sList = sList & Join(vList)
    End If
    Call fnLoadSelect(ofrmMenuLh, "fltEf", sList, False)

'Load Question Combo
'DataItemid`name`TrialID`code

    vData = goUser.DataLists.GetQuestionList(lStudyId)

    sList = "|`All questions"
    If Not IsNull(vData) Then
        ReDim vList(UBound(vData, 2)) As String
        For i = 0 To UBound(vData, 2)
            vList(i) = "|" & vData(eDropDownCol.Id, i) & "`" & vData(eDropDownCol.Text, i) ' & "`" & lStudyId

        Next
        sList = sList & Join(vList)
    End If


    Call fnLoadSelect(ofrmMenuLh, "fltQu", sList, False)


    
'load user combo
'username`usernamefull

    sList = "|`All users"
'Load eForm combo
'eformid`name`studyid`code
    vData = goUser.DataLists.GetUserList(lStudyId)
    If Not IsNull(vData) Then
        For i = 0 To UBound(vData, 2)
            sList = sList & "|" & vData(eDropDownCol.Id, i) & "`" & vData(eDropDownCol.Text, i)
        Next
    End If
    'ic 23/06/2003 only pass true for bInit when loading final select list
    Call fnLoadSelect(ofrmMenuLh, "fltUs", sList, True)
    
    
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, "LoadLhCombos", Err.Source) = Retry Then
        Resume
    End If
    
    
End Sub


'----------------------------------------------------------------------------------------'
Public Sub fnLoadSelect(ofrmWebBrowser As frmWebBrowser, sFilterName As String, sList As String, bInit As Boolean)
'----------------------------------------------------------------------------------------'
'load selection box
'eg of code in JS
'function fnLoadSelect(sName, sList, bClear, bCodeAndText)
'revisions
'ic 23/06/2003 changed fnLoadSelect arguments
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel

'    If Left(sList, 1) = "|" Then
'        sList = Mid(sList, 2)
'    End If

    
    ofrmWebBrowser.ExecuteJavaScript.fnLoadSelect (sFilterName), (sList), RtnJSBoolean(bInit)
    
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, "fnLoadSelect", Err.Source) = Retry Then
        Resume
    End If
    
    
End Sub

'eformTop.htm *************************************8
'----------------------------------------------------------------------------------------'
Public Sub fnSave(vParams As Variant)
'----------------------------------------------------------------------------------------'
'   '    Function fnSave(){eW.fnSave ("m3")}
'Do action from an eForm
'"m0"=Save
'"m1"=Save and Previous
'"m2"=Save and Next
'"m3"=Return (Save and return to schedule)
'"m4"=Cancel (Cancel and return to schedule)
'"e#"=Save and got CRFPageTaskId #
'"v#"=Save and got VisitId #
'---------------------------------------------------------------------------------------------
Dim lId As Long
Dim sParam As String

    On Error GoTo ErrLabel


    sParam = vParams(1)
    lId = Mid(sParam, 2)
    Select Case Left(sParam, 1)
    Case "m" 'menu id
        Select Case lId
        Case 0 'save
            Call frmEFormDataEntry.eFormAction(eaSave)
        Case 1 'save and previous
            Call frmEFormDataEntry.eFormAction(eaPrevious)
        Case 2 'save and next
            Call frmEFormDataEntry.eFormAction(eaNext)
        Case 3 'Return
            Call frmEFormDataEntry.eFormAction(eaReturn)
        Case 4 'Cancel
            Call frmEFormDataEntry.eFormAction(eaCancel)
        Case 5 'print
            Call frmEFormDataEntry.eFormAction(eaPrint)
        End Select
    Case "e" 'efi id ' save and goto EFI
        Call frmEFormDataEntry.eFormAction(eaGotoEFormInThisVisit, lId)
    Case "v" 'visitid ' save and goto Visit
        Call frmEFormDataEntry.eFormAction(eaGotoThisEFormInVisit, lId)
    End Select
    
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub


'******************* for appHeaderLh
'----------------------------------------------------------------------------------------'
Public Sub fnHomeUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'    function fnHomeUrl(){window.navigate("BVfnHomeUrl);}
'Print current HTML form
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel


    If frmMenu.GUISubjectClose(True, True) Then
        'only show home if subject closed
        Call frmMenu.ShowHome
    End If
    
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'******************* for all
'----------------------------------------------------------------------------------------'
Public Sub fnPrint(vParams As Variant)
'----------------------------------------------------------------------------------------'
'    function fnPrint(){window.navigate("//XfnPrint);}
'Print current HTML form
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel


    'mofrmWebBrowser.PrintHTMLPage
    MsgBox "Can't print"

    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnPrintAll(vParams As Variant)
'----------------------------------------------------------------------------------------'
'    function fnPrint(){window.navigate("//XfnPrintAll);}
'TA 18/06/2003: Print all eforms from shcedule
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel

    Select Case DialogQuestion("Do you wish to include blank eForms?", "Print all eForms", True)
    Case vbYes
        'include blank ones
        Call EFormPrinting(goStudyDef.Subject, False)
    Case vbNo
        'exclude blank ones
        Call EFormPrinting(goStudyDef.Subject, True)
    Case Else
        'cancel select - do nothing
    End Select
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'******************* for MenuLh
'----------------------------------------------------------------------------------------'
Public Sub fnSelectChange(vParams As Variant)
'----------------------------------------------------------------------------------------'
'"function fnSelectChange(sLst,sVa){window.navigate('VBfnSelectChange|'+sLst+'|'+sVa);}"
'someone has changed a selction box
'----------------------------------------------------------------------------------------'

On Error GoTo ErrLabel


    Select Case CStr(vParams(1))
    Case "fltSt"
          Call frmMenu.LoadLhCombos(CLng(vParams(2)))
    End Select
    
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub


'----------------------------------------------------------------------------------------'
Public Sub fnTimeSynchUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'Show remote time synchronistation form
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel

    frmTimeSynch.Show vbModal
        
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub
'----------------------------------------------------------------------------------------'
Public Sub fnSubmitOptions(vParams As Variant)
'----------------------------------------------------------------------------------------'
'function fnSubmitOptions(bFnKey,bSymb,bDtFormat,bServerDt,bSplit,bSameForm){window.navigate('VBfnSubmitOptions|'+bFnKey+'|'+sbSymb+'|'+bDtFormat+'|'+bServerDt+'|'+bSplit+'|'+bSameForm);}" & vbCrLf
'change options             1      2      3       4         5       6
'----------------------------------------------------------------------------------------'

On Error GoTo ErrLabel


    Select Case vParams(1)
    Case "true", "false": frmMenu.ToggleFunctionKeys CBool(vParams(1))
    End Select

    Select Case vParams(2)
    Case "true", "false": frmMenu.ToggleSymbols CBool(vParams(2))
    End Select

    ' RS 21/03/2003: Use local date/time format for data entry
    Select Case vParams(3)
    Case "true", "false": frmMenu.ToggleLocalFormat CBool(vParams(3))
    End Select

    Select Case vParams(5)
    Case "true", "false": frmMenu.ToggleSplitScreen CBool(vParams(5))
    End Select
    
    Select Case vParams(6)
    Case "true", "false": frmMenu.ToggleSameEform CBool(vParams(6))
    End Select
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnNewSubjectUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'    function fnNewSubjectUrl(){window.navigate("//XfnNewSubjectUrl");}
' Create a subject
'----------------------------------------------------------------------------------------'

On Error GoTo ErrLabel


'ignore parameters at first

    Call frmMenu.ShowNewSubject
    
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnScheduleUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'    function fnSubjectListUrl(sSt,sSi,sSj){window.navigate("//XfnScheduleUrl|sql_macro3|"+sSt+"|"+sSi+"|"+sSj);}
' Open an existing study subject
'----------------------------------------------------------------------------------------'
'1 is study id, 2 is site, 3 is person id
Dim lStudyId As Long
Dim sSite As String
Dim lPersonId As Long

On Error GoTo ErrLabel


    lStudyId = CLng(vParams(1))
    
    sSite = CStr(vParams(2))
    lPersonId = CLng(vParams(3))
    
    Call frmMenu.SubjectOpen(lStudyId, sSite, lPersonId)
    
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnSubjectListUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'    function fnSubjectListUrl(sSt,sSi,sSj){window.navigate("//XfnSubjectListUrl|"+sSt+"|"+sSi+"|"+sSj);}
' Show list of subjects according to filter parameters
'----------------------------------------------------------------------------------------'
'1 is study id, 2 is site, 3 is label
Dim lStudyId As Long
Dim sSite As String
Dim sLabel As String
Dim sStudy As String
Dim sOrderBy As String
On Error GoTo ErrLabel


    'lStudyId = CLng()
    If vParams(1) = -1 Or vParams(1) = "undefined" Or vParams(1) = "" Then
        'all studies selected
        sStudy = ""
    Else
        If IsNumeric(vParams(1)) Then
            sStudy = goUser.Studies.StudyById(vParams(1)).StudyName
        Else
            sStudy = vParams(1)
        End If
    End If
    
    sSite = CStr(vParams(2))
    If sSite = "undefined" Or sSite = "" Then
        sSite = ""
    End If
    
    sLabel = CStr(vParams(3))
    If sLabel = "undefined" Or sLabel = "undefined" Then
        sLabel = ""
    End If
    

    sOrderBy = CStr(vParams(4))
    If sOrderBy = "undefined" Then
        sOrderBy = "-1"
    End If
    
    Call frmMenu.ShowSubjectList(sStudy, sSite, sLabel, "", sOrderBy, "true")

    
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnBrowserUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'display databrowser
'fnBrowserUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sSs,sLk,sB4,sTm,sCm,sDi,sSd,sNo,"2");
'             1   2   3   4   5   6   7   8   9   10  11  12  13  14  15  16
'#CLINICALCODING:
'fnBrowserUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sSs,sLk,sB4,sTm,sCm,sDi,sSd,sNo,sCS,sDc,"2");
'             1   2   3   4   5   6   7   8   9   10  11  12  13  14  15  16  17  18
'TA 12/02/2003: if not monitor then give them chooice of subjects if more than one subject label matches
' NCJ 20 Mar 03 - Ensure Data Browser is initialised even if no matching records
' MLM 26/03/03: Pass vErrors into GetDataBrowser, and display any errors returned.
' DPH 27/10/2003 - Catch 'out of memory' (7) and 'memory allocation to flexgrid' (30006)  errors
' ic 28/07/2005 added clinical coding
'----------------------------------------------------------------------------------------'
Dim sGet As String
Dim vData As Variant
Dim enDataBrowserType As eDataBrowserType
Dim sSubject As String
Dim sPersonId As String
Dim sAdditional As String
Dim ofrmOption As frmOptionMsgBox
Dim sSubjectList As String
Dim lSubjectRow As Long
Dim vErrors As Variant
Dim sDictionaryName As String
Dim sDictionaryVersion As String


    On Error GoTo ErrLabel
    
    If Not frmMenu.SplitScreen Then
        'prompt to close dataentry form and exit if not closed in non split screen mode
        If frmMenu.IsDataEntryFormLoaded(True) Then
            Exit Sub
        End If
    End If
      
   'TA temp fix   for no study prob
   If vParams(1) = "</select" Then Exit Sub
   
    If vParams(6) = "" And Not goUser.CheckPermission(gsFnMonitorDataReviewData) Then
        DialogInformation "Please enter a subject label in the search criteria", "Data Browser"
        Exit Sub
    End If


    sSubject = vParams(6)

    
    
    If sSubject <> "" And Not goUser.CheckPermission(gsFnMonitorDataReviewData) Then
        'code here to select 1 subject from all matching subject labels
        vData = goUser.DataLists.GetSubjectList(sSubject, goUser.Studies.StudyById((vParams(1))).StudyName, (vParams(2)))
        If Not IsNull(vData) Then
            If UBound(vData, 2) > 0 Then
                If UBound(vData, 2) > 9 Then
                    'only allow a max ten subjects
                    Call DialogWarning("There are more than 10 subjects that match that subject label.  Pease refine your search", "Data browser")
'EXIT SUB HERE
                    Exit Sub
                End If
                sSubjectList = vData(eSubjectListCols.StudyName, 0) & "/" & vData(eSubjectListCols.Site, 0) & "/" & vData(eSubjectListCols.SubjectLabel, 0) ' & " (" & vData(eSubjectListCols.SubjectId, 0) & ")"
                For lSubjectRow = 1 To UBound(vData, 2)
                    sSubjectList = sSubjectList & "|" & vData(eSubjectListCols.StudyName, lSubjectRow) & "/" & vData(eSubjectListCols.Site, lSubjectRow) & "/" & vData(eSubjectListCols.SubjectLabel, lSubjectRow) ' & " (" & vData(eSubjectListCols.SubjectId, lSubjectRow) & ")"
                Next
                Set ofrmOption = New frmOptionMsgBox
                lSubjectRow = ofrmOption.Display("Data browser", "Please select a subject", "More than one subject matches the subject label: " & sSubject, sSubjectList)
                Set ofrmOption = Nothing
                If lSubjectRow = -1 Then
                    'cancel pressed
                    Exit Sub
                End If
                'allow for zero based
                lSubjectRow = lSubjectRow - 1
                'DialogInformation "There is more than one subject matching this label"
            Else
                lSubjectRow = 0
            End If
            sPersonId = vData(eSubjectListCols.SubjectId, lSubjectRow)
            sSubject = ""
        Else
            DialogInformation "No matching records"
'EXIT SUB
            Exit Sub
        End If
    End If
    
    HourglassOn
    
    frmHourglass.Display "Loading data browser", Not frmMenu.SplitScreen
    
    'ic 28/07/2005 added clinical coding
    sGet = vParams(18)
    If (vParams(17) > "") Then
        sDictionaryName = Split(vParams(17), "~")(0)
        sDictionaryVersion = Split(vParams(17), "~")(1)
    End If
        
    vData = GetWinIO.GetDataBrowser(goUser, False, vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(7), sPersonId, sSubject, _
    vParams(8), vParams(9), vParams(10), vParams(11), vParams(12), vParams(13), vParams(14), vParams(15), vParams(16), sDictionaryName, sDictionaryVersion, sGet, False, vErrors)


'    If Not IsNull(vData) Then
    ' NCJ 20 Mar 03 - Do browser display anyway to ensure that data browser gets initialised
    Select Case sGet
    Case "0": enDataBrowserType = dbteForms
    Case "1": enDataBrowserType = dbtDataItemResponse
    Case "2": enDataBrowserType = dbtDataItemResponseHistory
    End Select
    
    OpenWinForm wfDataBrowser
    Call frmDataItemResponse.DisplayNew(vData, enDataBrowserType)
    frmMenu.Resize
        
    UnloadfrmHourglass
'    Else
'        UnloadfrmHourglass
    
    'MLM 26/03/03:
    If Not IsEmpty(vErrors) Then
        DialogError "An invalid search date was specified. The search was performed without filtering on date."
    End If
    
    If IsNull(vData) Then
        DialogInformation "No matching records"
    End If

    HourglassOff
       
    Exit Sub
    
ErrLabel:
    ' DPH 27/10/2003 - Catch 'out of memory' (7) and 'memory allocation to flexgrid' (30006)  errors
    If Err.Number = 7 Or Err.Number = 30006 Then
        Err.Clear
        On Error GoTo ErrMemory
        UnloadfrmHourglass
        HourglassOff
        DialogInformation "Your query returned too many records and you have run out of memory. Please refine your search in order to return less records"
    End If

ErrMemory:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
            
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnDiscUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
''VBfnDiscUrl|'+sSt+'|'+sSi+'|'+sVi+'|'+sEf+'|'+sQu+'|'+sSj+'|'+sUs+'|'+sB4+'|'+sTm+'|'+sSs);}"
'                1        2       3       4       5      6       7      8        9      10
'display mimessagelist
'fnDiscUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSs)
'----------------------------------------------------------------------------------------'
On Error GoTo ErrLabel

    
    Call frmViewDiscrepancies.Display(vParams, mimtDiscrepancy, MonitorMIMessage)
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
        
End Sub
 
 '----------------------------------------------------------------------------------------'
Public Sub fnOCDiscUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
On Error GoTo ErrLabel

    
    Call frmMenu.ViewOCDiscs
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
        
End Sub
 
'----------------------------------------------------------------------------------------'
Public Sub fnSDVUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
''VBfnSDVUrl|'+sSt+'|'+sSi+'|'+sVi+'|'+sEf+'|'+sQu+'|'+sSj+'|'+sUs+'|'+sB4+'|'+sTm+'|'+sSs+'|'+sObj);}"
'                1        2       3       4       5      6       7      8        9      10      11
'display mimessagelist
'fnDiscUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSs)
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel

    Call frmViewDiscrepancies.Display(vParams, mimtSDVMark, MonitorMIMessage)

   Exit Sub
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
        
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnNoteUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
''VBfnNoteUrl|'+sSt+'|'+sSi+'|'+sVi+'|'+sEf+'|'+sQu+'|'+sSj+'|'+sUs+'|'+sB4+'|'+sTm+'|'+sSs);}"
'                1        2       3       4       5      6       7      8        9      10
'display mimessagelist
'fnDiscUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSs)
'----------------------------------------------------------------------------------------'

On Error GoTo ErrLabel

    Call frmViewDiscrepancies.Display(vParams, mimtNote, MonitorMIMessage)
       
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
        
End Sub

'********************************* appMenuTop functions

'----------------------------------------------------------------------------------------'
Public Sub fnAbout(vParams As Variant)
'----------------------------------------------------------------------------------------'
'    Function fnHelp(){window.navigate("//XfnHelp");}
'----------------------------------------------------------------------------------------'

On Error GoTo ErrLabel

    frmAbout.Display
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnHelp(vParams As Variant)
'----------------------------------------------------------------------------------------'
'    Function fnHelp(){window.navigate("//XfnHelp");}

'----------------------------------------------------------------------------------------'

On Error GoTo ErrLabel


    Call MACROHelp(frmMenu.hWnd, App.Title)
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnToggleLh(vParams As Variant)
'----------------------------------------------------------------------------------------'
'    Function fnHelp(){window.navigate("//VBfnToggleLh");}

'----------------------------------------------------------------------------------------'

On Error GoTo ErrLabel


    'this turns left hand pane on/off
    frmMenu.ToggleLeftHandPane
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnLogOutUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'    Function fnLogOutUrl(){window.navigate("//XfnLogOutUrl");}

'----------------------------------------------------------------------------------------'

On Error GoTo ErrLabel


    Call frmMenu.UserLogOut

    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnSwitchUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'


'----------------------------------------------------------------------------------------'

On Error GoTo ErrLabel


    Call frmMenu.UserSwitch

    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub



'----------------------------------------------------------------------------------------'
Public Sub fnHoldUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'    Function fnHoldUrl(){window.navigate("//XfnHoldUrl");}

'TA don;t want refernce to frmMenu in here
'----------------------------------------------------------------------------------------'

On Error GoTo ErrLabel

'TA 10/04/2003: do not allow without closing the eForm
    If Not frmMenu.IsDataEntryFormLoaded Then
        If Not frmTimeOutSplash.Display(False) Then
            Close
            ' unload all forms and exit
            Call UnloadAllForms
        End If
    End If
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    

End Sub



'forms.htm (schedule) *************************************8
'----------------------------------------------------------------------------------------'
Public Sub fnEformUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'   function goE(ef,id){fnEformUrl(window.navigate("XfnEformUrl|sql_macro3"+"|"+'<%=fltSt%>'+"|"+'<%=fltSi%>'+"|"+'<%=fltSj%>'+"|"+ef+"|"+id+"|"+'de')}
'Open an eForm
'---------------------------------------------------------------------------------------------
Dim lEFITaskId As Long

    On Error GoTo ErrLabel


    lEFITaskId = CLng(vParams(1))  'convert from string
    Select Case CLng(vParams(2))
    Case 1 'left button
        If ScheduleOpeneForm(lEFITaskId) Then
            'TA 15/11/2002: workaroung to stop shcedule appearing on top after a double click
            mofrmWebBrowser.ZOrder 1
        End If
    Case 2, 3
        Call DisplayScheduleMenu(lEFITaskId)
    End Select
    

    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    

End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnV(vParams As Variant)
'----------------------------------------------------------------------------------------'
'show visit menu
'----------------------------------------------------------------------------------------'

On Error GoTo ErrLabel

    If vParams(2) = "2" Then
        'right click
        Call DisplayVisitMenu((vParams(1)))

    End If
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub


'----------------------------------------------------------------------------------------'
Public Sub fnS(vParams As Variant)
'----------------------------------------------------------------------------------------'
'show subject menu
'----------------------------------------------------------------------------------------'

On Error GoTo ErrLabel

    If vParams(1) = "2" Then
    '   right click
        'show menu for currently opened subject
        Call DisplaySubjectMenu
    End If
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub


'----------------------------------------------------------------------------------------'
Private Sub fnViewDiscrepancies(enStatus As eDiscrepancyMIMStatus)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
Dim asData(10) As String


On Error GoTo ErrLabel

'TA 16/01/2003: use the winio to get data so that dataitemid and status are calculated - bug list 603
        '1st char denotes raised diescpeancies, 2nd character denotes response, 3rd denotes closed to be included in search
       If enStatus = eDiscrepancyMIMStatus.dsRaised Then
            asData(10) = "100"
       Else
            asData(10) = "010"
       End If


        Call fnDiscUrl(asData)
        
        
ErrLabel:
     If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, "fnViewDiscrepancies", Err.Source) = Retry Then
        Resume
    End If
End Sub



'----------------------------------------------------------------------------------------'
Public Sub fnRaisedDiscUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'

On Error GoTo ErrLabel

    Call fnViewDiscrepancies(eDiscrepancyMIMStatus.dsRaised)

Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnRespondedDiscUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
    
On Error GoTo ErrLabel

    Call fnViewDiscrepancies(eDiscrepancyMIMStatus.dsResponded)

    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnPlannedSDVUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
''VBfnSDVUrl|'+sSt+'|'+sSi+'|'+sVi+'|'+sEf+'|'+sQu+'|'+sSj+'|'+sUs+'|'+sB4+'|'+sTm+'|'+sSs+'|'+sObj);}"
'                1        2       3       4       5      6       7      8        9      10      11
'----------------------------------------------------------------------------------------'
On Error GoTo ErrLabel

Dim asData(11) As String

'TA 23/01/2003: use the winio to get data so that dataitemid and status are calculated
        '1st char denotes planned SDVs
         asData(10) = "1000"

        Call fnSDVUrl(asData)

    Exit Sub
        
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnLNRUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'             sHTML = sHTML & "function fnViewLabs{window.navigate('fnLNRUrl');}"" & vbCrLf"
'----------------------------------------------------------------------------------------'
    
On Error GoTo ErrLabel

    Call frmLabDefinitions.Display(frmMenu.gTrialOffice.Site)

    Exit Sub
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|JSFunctions.fnLNRUrl"
    
End Sub


'----------------------------------------------------------------------------------------'
Public Sub fnChangePasswordUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'sHTML = sHTML & "function fnChangePasswordUrl(){window.navigate('fnChangePasswordUrl');}"" & vbCrLf"
'----------------------------------------------------------------------------------------'
    
On Error GoTo ErrLabel

    'MsgBox "change password"

    'Exit Sub
    'ATO 10/02/2003 'show change password form
    Call frmPasswordChange.Display(goUser, gsSecCon)
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

                      
'----------------------------------------------------------------------------------------'
Public Sub fnViewChangesUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'display databrowser since last login
'----------------------------------------------------------------------------------------'

Dim vData As Variant
Dim oDB As DataBrowser
Dim enDataBrowserType As eDataBrowserType

    On Error GoTo ErrLabel

    If Not frmMenu.SplitScreen Then
        'prompt to close dataentry form and exit if not closed in non split screen mode
        If frmMenu.IsDataEntryFormLoaded(True) Then
            Exit Sub
        End If
    End If
    
    HourglassOn
    
    frmHourglass.Display "Loading data browser", Not frmMenu.SplitScreen
    
    Set oDB = New DataBrowser
    
    ' DPH 16/04/2004 - Remove LocalNumToStandard call as passing a double into GetData function
    ' DPH 14/04/2004 - Subtract 1 from double date before converting to standard string type
    'TA 03/04/2003: use convertlocalnumtostandard - ConvertLocalNumToStandard(goUser.LastLogin) - 1
    vData = oDB.GetData(goUser.CurrentDBConString, False, dbtDataItemResponse, _
                goUser.DataLists.StudiesSitesWhereSQL("DATAITEMRESPONSE.CLINICALTRIALID", "DATAITEMRESPONSE.TRIALSITE"), _
                False, , , , , , , , , , CDbl(goUser.LastLogin - 1))

    If Not IsNull(vData) Then
        enDataBrowserType = dbtDataItemResponse
        OpenWinForm wfDataBrowser
        Call frmDataItemResponse.DisplayNew(vData, enDataBrowserType)
        frmMenu.Resize
            
        UnloadfrmHourglass
    Else
        UnloadfrmHourglass
        DialogInformation "No matching records"
    End If


    HourglassOff
       
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
        
End Sub


'----------------------------------------------------------------------------------------'
Public Sub fnDBLockUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'display database lock admin form
'----------------------------------------------------------------------------------------'
On Error GoTo ErrLabel
    If FormIsLoaded(g_DATAENTRY_FORM_NAME) Then
        DialogInformation "You must close the subject before removing database locks."
    Else
        Call frmLocksAdmin.Display(goUser, goUser.DatabaseCode)
    End If

ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If

End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnCloseSchedule(vParams As Variant)
'----------------------------------------------------------------------------------------'
'close schedule (subject) used to be file close
'----------------------------------------------------------------------------------------'

'---------------------------------------------------------------------
' Close the current study subject
'---------------------------------------------------------------------
    
    On Error GoTo ErrLabel
    
    'REM 25/04/02 - changed the last boolean parameter to False, i.e. don't unload study on close
    CloseSubject goStudyDef, True, False, True

ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
End Sub


'----------------------------------------------------------------------------------------'
Public Sub fnXferURL(vParams As Variant)
'----------------------------------------------------------------------------------------'
'transfer data
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel
    
    frmMenu.TransferData

ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnViewLFUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'view lock / freeze history
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel
    
    frmMenu.ViewLockFreeze

ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub
'----------------------------------------------------------------------------------------'
Public Sub fnCommunicationHistoryUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
' RS 18/02/2003: Added Communication Panel
'----------------------------------------------------------------------------------------'
    frmMenu.CommunicationHistory
End Sub
'----------------------------------------------------------------------------------------'
Public Sub fnCommunicationStatusReportUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
' RS 18/02/2003: Added Communication Panel
'----------------------------------------------------------------------------------------'
    frmMenu.CommunicationStatusReport
End Sub
'----------------------------------------------------------------------------------------'
Public Sub fnRemoteTimeSynchronisationUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
' RS 18/02/2003: Added Communication Panel
'----------------------------------------------------------------------------------------'
    frmMenu.RemoteTimeSynchronisation
End Sub
'----------------------------------------------------------------------------------------'
Public Sub fnResetTransferStatusUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
' RS 18/02/2003: Added Communication Panel
'----------------------------------------------------------------------------------------'
    frmMenu.ResetTransferStatus
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnViewComUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'view comm settings
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel
    
    frmMenu.ShowCommsSettings

Exit Sub

ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnChangeComUrl(vParams As Variant)
'----------------------------------------------------------------------------------------'
'change com settings
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel
    
    frmMenu.ShowCommsSettings

Exit Sub


ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnTemplates(vParams As Variant)
'----------------------------------------------------------------------------------------'
'change com settings
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel
    
    frmMenu.Templates

Exit Sub


ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub fnRegister(vParams As Variant)
'----------------------------------------------------------------------------------------'
'change com settings
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel
    
    frmMenu.Register

Exit Sub


ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub


'----------------------------------------------------------------------------------------'
Public Sub fnKeyDown(vParams As Variant)
'----------------------------------------------------------------------------------------'
'change com settings
'----------------------------------------------------------------------------------------'
Dim bShiftKey As Boolean
Dim bCtrlKey As Boolean
Dim bAltKey As Boolean
Dim nKeyCode As Integer
Dim nShift As Integer

    On Error GoTo ErrLabel
    
    nKeyCode = CInt(vParams(1))
    bShiftKey = CBool(vParams(2))
    bCtrlKey = CBool(vParams(3))
    bAltKey = CBool(vParams(4))
    
     nShift = 0
    If bShiftKey Then
        nShift = nShift + vbShiftMask
    End If
    
    If bCtrlKey Then
        nShift = nShift + vbCtrlMask
    End If
    
    If bAltKey Then
        nShift = nShift + vbAltMask
    End If
    
    If FormIsLoaded(g_DATAENTRY_FORM_NAME) Then
        Call frmEFormDataEntry.KeyDown(nKeyCode, nShift)
    Else
        'de form not loaded
    
    End If
    
    'does not currently work becasue keycode and window are in awrong case
'    If nKeyCode = vbKeyF5 Then
'        mofrmWebBrowser.WebBrowser.Document.Window.event.KeyCode = 0
'    End If

'    Select Case nShift
'    Case 1 ' or vbShiftMask
'       Debug.Print "You pressed the SHIFT key."
'    Case 2 ' or vbCtrlMask
'       Debug.Print "You pressed the CTRL key."
'    Case 4 ' or vbAltMask
'       Debug.Print "You pressed the ALT key."
'    Case 3
'       Debug.Print "You pressed both SHIFT and CTRL."
'    Case 5
'       Debug.Print "You pressed both SHIFT and ALT."
'    Case 6
'       Debug.Print "You pressed both CTRL and ALT."
'    Case 7
'       Debug.Print "You pressed SHIFT, CTRL, and ALT."
'    End Select
    
    
    
    Debug.Print vParams(1), bShiftKey, bCtrlKey, bAltKey
Exit Sub


ErrLabel:
    If MACROErrorHandler("JSFunctions", Err.Number, Err.Description, (Mid(vParams(eFunctionParams.fpFunctionName), 14)), Err.Source) = Retry Then
        Resume
    End If
    
End Sub

