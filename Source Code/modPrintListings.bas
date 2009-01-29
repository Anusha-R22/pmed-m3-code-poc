Attribute VB_Name = "modPrintListings"
'--------------------------------------------------------------------------------------------------
'   File:       modPrintListings.bas
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Ashitei Trebi-Ollennu, September 2001
'   Purpose:    Routines for carrying out non Crystal Reports printing as
'               required by document Re-Write of Ex-Crystal Reports Listings
'--------------------------------------------------------------------------------------------------
'Revisions:23/11/2001 Trapping for when printer turned off Buglist report no.14
'         10/12/2001 Fixes for Current Buglist nos.23,6,19,7,5,
'          ASH 13/05/2002 Added Unit Conversion printing for LM
'          ASH 21/06/2002 Bug 2.2.16 no.10
'--------------------------------------------------------------------------------------------------
Option Explicit
'--------------------------------------------------------------------------------------------------
Public Sub Print_QuestionWithinEformReport()
'--------------------------------------------------------------------------------------------------
' Called from frmmenu to print all questions on a form
'--------------------------------------------------------------------------------------------------
Dim nQuestions As Integer
Dim neForms As Integer
Dim sType As String
Dim rsQuestionsInForm As ADODB.Recordset
Dim rseForm As ADODB.Recordset
Dim mlTrialID As Long
Dim mnPrintBlock As Integer
Dim msTrialName As String
Dim mlPrintingWidth As Long
Dim msPrintName As String
Dim msSQL As String
Dim msSQL1 As String
Dim mnCurrentY As Integer

   On Error GoTo ErrHandler
    
    HourglassOn
    
    mlTrialID = frmMenu.ClinicalTrialId
    msPrintName = "Questions in eForms Reports"
    msTrialName = frmMenu.ClinicalTrialName
    mnPrintBlock = 583
    msSQL = ""
    msSQL1 = ""
    
    msSQL1 = " SELECT DISTINCT CRFPage.CRFPageCode,CRFPage.CRFTitle,"
    msSQL1 = msSQL1 & " CRFPage.CRFPageId,ClinicalTrial.ClinicalTrialName,"
    msSQL1 = msSQL1 & " ClinicalTrial.ClinicalTrialId,CRFPage.ClinicalTrialId,"
    msSQL1 = msSQL1 & " CRFElement.ClinicalTrialId,DataItem.ClinicalTrialId"
    msSQL1 = msSQL1 & " FROM  ClinicalTrial,CRFPage,CRFElement,DataItem  "
    msSQL1 = msSQL1 & " WHERE ClinicalTrial.ClinicalTrialId = " & mlTrialID
    msSQL1 = msSQL1 & " AND CRFElement.ClinicalTrialId = " & mlTrialID
    msSQL1 = msSQL1 & " AND CRFPage.ClinicalTrialId = " & mlTrialID
    msSQL1 = msSQL1 & " AND DataItem.ClinicalTrialId = " & mlTrialID
    msSQL1 = msSQL1 & " AND DataItem.DataItemId = CRFElement.DataItemId"
    msSQL1 = msSQL1 & " AND CRFElement.CRFPageId = CRFPage.CRFPageId"
    msSQL1 = msSQL1 & " AND CRFElement.ClinicalTrialId = ClinicalTrial.ClinicalTrialId "
    msSQL1 = msSQL1 & " AND DataItem.ClinicalTrialId = ClinicalTrial.ClinicalTrialId"
    
    Set rseForm = New ADODB.Recordset
    rseForm.Open msSQL1, MacroADODBConnection, adOpenKeyset
    neForms = rseForm.RecordCount
    
    If neForms < 1 Then
            Call DialogInformation(" There are no eform questions in this study to print.")
            Call HourglassOff
            Exit Sub
        End If
    
    '01/11/2001 changed True to false Bug Report Number 115 Ash.
    '23/11/2001 changed to true. Buglist no.14
    On Error Resume Next
    
    frmMenu.CommonDialog1.CancelError = True
    Printer.Orientation = vbPRORPortrait
    Printer.TrackDefault = True
    frmMenu.CommonDialog1.ShowPrinter
    Printer.Orientation = vbPRORPortrait
    
    'Check for errors in ShowPrinter (including a cancel)
    If Err.Number > 0 Then
        HourglassOff
        Exit Sub
    End If
    
    'hourglass displayed
    DoEvents
    
    On Error GoTo ErrHandler
    
    'extend the top and left margins by 1/2 inch to 3/4 inch approx
    Printer.ScaleLeft = -720
    Printer.ScaleTop = -720
    
    'print main header information
    mlPrintingWidth = Printer.ScaleWidth - 1440
    Call PrintHeader(mlPrintingWidth, msPrintName)
    
    'print study Name
    Call PrintStudyName(msTrialName)
    
    'print sub header information
    rseForm.MoveFirst
    Do While Not rseForm.EOF
        'continue printing on a new page
        If IsPageSizeEnough(mnPrintBlock) = False Then
            Printer.NewPage
            Call PrintHeader(mlPrintingWidth, msPrintName)
        End If
        
        'printer settings
        Printer.FontName = "Tahoma"
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 240
        Printer.FontSize = 10
        Printer.FontBold = True
        
        'print sub headings
        Printer.Print "Form Title: ", ;
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.Print rseForm.Fields(1).Value
        Printer.FontBold = True
        Printer.CurrentX = 530
        Printer.FontSize = 10
        Printer.Print "Code: ", ;
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.Print rseForm.Fields(0).Value
        
        'print  header for report details
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 240
        Printer.FontSize = 8
        Printer.FontBold = True
        Printer.Print "Question Name ", ;
        Printer.CurrentX = 4500
        Printer.Print "Code", ;
        Printer.CurrentX = 7000
        Printer.Print "Type", ;
        Printer.CurrentX = 9400
        Printer.Print "Size", ;
        Printer.FontBold = False
        Printer.CurrentY = Printer.CurrentY + 240
        
        'ASH 21/06/2002 Bug 2.2.16 no.10
        'msSQL = " SELECT CRFElement.Caption,DataItem.DataItemCode,"
        msSQL = " SELECT DataItem.DataItemName,DataItem.DataItemCode,"
        msSQL = msSQL & "DataItem.DataType,DataItem.DataItemLength"
        msSQL = msSQL & " FROM DataItem,CRFElement "
        msSQL = msSQL & " WHERE CRFElement.ClinicalTrialId = " & mlTrialID
        msSQL = msSQL & " AND DataItem.ClinicalTrialId = " & mlTrialID
        msSQL = msSQL & " AND DataItem.DataItemId = CRFElement.DataItemId"
        msSQL = msSQL & " AND CRFElement.CRFPageId = " & rseForm.Fields(2).Value
    
        Set rsQuestionsInForm = New ADODB.Recordset
        rsQuestionsInForm.Open msSQL, MacroADODBConnection, adOpenKeyset
        nQuestions = rsQuestionsInForm.RecordCount
    
        If nQuestions <= 0 Then
            Call DialogInformation(" There are no questions in this study to print.")
            Call HourglassOff
            Exit Sub
        End If
        
        rsQuestionsInForm.MoveFirst
            Do While Not rsQuestionsInForm.EOF
                Printer.FontSize = 8
                Printer.CurrentX = 0
                Printer.Print rsQuestionsInForm.Fields(0).Value, ;
                Printer.CurrentX = 4500
                Printer.Print rsQuestionsInForm.Fields(1).Value, ;
                Printer.CurrentX = 7000
                sType = DecodeDataType(rsQuestionsInForm.Fields(2).Value)
                Printer.Print sType, ;
                Printer.CurrentX = 9400
                Printer.Print rsQuestionsInForm.Fields(3).Value
                rsQuestionsInForm.MoveNext
            Loop
               'draw line
                Printer.DrawWidth = 2
                Printer.CurrentY = Printer.CurrentY + 120
                mnCurrentY = Printer.CurrentY
                Printer.Line (0, mnCurrentY)-(mlPrintingWidth, mnCurrentY)
                
                rseForm.MoveNext
    Loop
    
    Set rsQuestionsInForm = Nothing
    Set rseForm = Nothing
    Printer.EndDoc
    Call HourglassOff

Exit Sub
ErrHandler:

    If Err.Number = 482 Then
        DialogError ("Printer error.Check printer and retry"), "Printer Error"
        Call HourglassOff
        Exit Sub
    End If
  
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Print_QuestionWithinEformReport", "modPrintListings.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub
'--------------------------------------------------------------------------------------------------
Public Sub Print_ScheduleVisitReport()
'--------------------------------------------------------------------------------------------------
'called from frmMenu to print all scheduled visits
'--------------------------------------------------------------------------------------------------
Dim rsStudyDefSchedule As ADODB.Recordset
Dim rsVisitSchedule As ADODB.Recordset
Dim nVisits As Integer
Dim nVisits1 As Integer
Dim sType As String
Dim lvWrap As Long
Dim mlTrialID As Long
Dim mnPrintBlock As Integer
Dim msTrialName As String
Dim mlPrintingWidth As Long
Dim msPrintName As String
Dim msSQL As String
Dim msSQL1 As String
Dim mnCurrentY As Integer

    On Error GoTo ErrHandler
    
    HourglassOn
    
    msPrintName = "Schedule Visit Reports"
    mlTrialID = frmMenu.ClinicalTrialId
    msTrialName = frmMenu.ClinicalTrialName
    mnPrintBlock = 583
    msSQL = ""
    msSQL1 = ""
    
    msSQL = " SELECT VisitName,VisitCode,VisitID"
    msSQL = msSQL & " FROM StudyVisit"
    msSQL = msSQL & " WHERE ClinicalTrialID = " & mlTrialID
    msSQL = msSQL & " ORDER BY VisitName "
    
    Set rsVisitSchedule = New ADODB.Recordset
    rsVisitSchedule.Open msSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic
    
    nVisits = rsVisitSchedule.RecordCount
    If nVisits < 1 Then
        Call DialogInformation(" There are no schedule visits in this study to print.")
        Call HourglassOff
        Exit Sub
    End If
    
    '01/11/2001 changed True to false Bug Report Number 115 Ash.
    '23/11/2001 changed to true. Buglist no.14
    On Error Resume Next
    
    frmMenu.CommonDialog1.CancelError = True
    Printer.Orientation = vbPRORPortrait
    Printer.TrackDefault = True
    frmMenu.CommonDialog1.ShowPrinter
    Printer.Orientation = vbPRORPortrait
    
    'Check for errors in ShowPrinter (including a cancel)
    If Err.Number > 0 Then
        HourglassOff
        Exit Sub
    End If
    
    'hourglass displayed
    DoEvents
    
    On Error GoTo ErrHandler
    
    'extend the top and left margins by 1/2 inch to 3/4 inch approx
    Printer.ScaleLeft = -720
    Printer.ScaleTop = -720
    
    'print main header information
    mlPrintingWidth = Printer.ScaleWidth - 1440
    Call PrintHeader(mlPrintingWidth, msPrintName)
    
    'print study Name
    Call PrintStudyName(msTrialName)
   
    rsVisitSchedule.MoveFirst
    Do While Not rsVisitSchedule.EOF
        
        'continue printing on a new page
        If IsPageSizeEnough(mnPrintBlock) = False Then
            Printer.NewPage
            Call PrintHeader(mlPrintingWidth, msPrintName)
        End If
        
        'print  header for report details
        Printer.CurrentX = 0
        Printer.FontName = "Tahoma"
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontSize = 8
        Printer.FontBold = True
        Printer.Print "Visit Name: ", ;
        Printer.Print rsVisitSchedule.Fields(0).Value
        Printer.CurrentX = 470
        Printer.Print "Code: ", ;
        Printer.CurrentX = 1120
        Printer.Print rsVisitSchedule.Fields(1).Value
        Printer.FontBold = False
        
        'print  header for report details
        Printer.CurrentX = 1120
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Form Name ", ;
        Printer.CurrentX = 6000
        Printer.Print "Code ", ;
        Printer.CurrentX = 9000
        Printer.Print "Repeating "
        Printer.FontBold = False
        
        msSQL1 = " SELECT CRFPage.CRFTitle,CRFPage.CRFPageCode,StudyVisitCRFPage.Repeating"
        msSQL1 = msSQL1 & " FROM CRFPage,StudyVisitCRFPage "
        msSQL1 = msSQL1 & " WHERE StudyVisitCRFPage.ClinicalTrialID = " & mlTrialID
        msSQL1 = msSQL1 & " AND CRFPage.ClinicalTrialId = " & mlTrialID
        msSQL1 = msSQL1 & " AND CRFPage.ClinicalTrialId = StudyVisitCRFPage.ClinicalTrialId"
        msSQL1 = msSQL1 & " AND CRFPage.CRFPageId = StudyVisitCRFPage.CRFPageId"
        msSQL1 = msSQL1 & " AND StudyVisitCRFPage.VisitID = " & rsVisitSchedule.Fields(2).Value
        msSQL1 = msSQL1 & " ORDER BY CRFPage.CRFPageOrder "
        
        Set rsStudyDefSchedule = New ADODB.Recordset
        rsStudyDefSchedule.Open msSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic
        
        nVisits1 = rsStudyDefSchedule.RecordCount
        If nVisits1 < 1 Then
            Call DialogInformation(" There are no schedule visits in this study to print.")
            Call HourglassOff
            Exit Sub
        End If
        
        Do While Not rsStudyDefSchedule.EOF
            Printer.CurrentX = 1120
            Printer.Print rsStudyDefSchedule.Fields(0).Value, ;
            Printer.CurrentX = 6000
            Printer.Print rsStudyDefSchedule.Fields(1).Value, ;
            Printer.CurrentX = 9000
            sType = rsStudyDefSchedule.Fields(2).Value
            'Decode Repeating Questions
            If rsStudyDefSchedule.Fields(2).Value = 0 Then
                sType = " "
            Else
                sType = "Yes"
            End If
            Printer.Print sType
            rsStudyDefSchedule.MoveNext
        Loop
            'draw line
            Printer.DrawWidth = 2
            Printer.CurrentY = Printer.CurrentY + 120
            mnCurrentY = Printer.CurrentY
            Printer.Line (0, mnCurrentY)-(mlPrintingWidth, mnCurrentY)
            
            rsVisitSchedule.MoveNext
    Loop
        Set rsStudyDefSchedule = Nothing
        Set rsVisitSchedule = Nothing
        Printer.EndDoc
        Call HourglassOff

Exit Sub
ErrHandler:

    If Err.Number = 482 Then
        DialogError ("Printer error.Check printer and retry"), "Printer Error"
        Call HourglassOff
        Exit Sub
    End If
  
  Call HourglassOff
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Print_QuestionWithinEformReport", "modPrintListings.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub
'--------------------------------------------------------------------------------------------------
Public Sub Print_CategoryQuestionValues()
'--------------------------------------------------------------------------------------------------
'called from frmMenu to print all questions and thier values
'--------------------------------------------------------------------------------------------------
Dim rsStudyDefCategoryValues As ADODB.Recordset
Dim rsStudyCategory As ADODB.Recordset
Dim nCount As Integer
Dim nCount1 As Integer
Dim mlTrialID As Long
Dim mnPrintBlock As Integer
Dim msTrialName As String
Dim mlPrintingWidth As Long
Dim msPrintName As String
Dim msSQL As String
Dim msSQL1 As String
Dim mnCurrentY As Integer
    
    On Error GoTo ErrHandler
    
    HourglassOn
    
    msPrintName = "Category Question Values Reports"
    msTrialName = frmMenu.ClinicalTrialName
    mlTrialID = frmMenu.ClinicalTrialId
    mnPrintBlock = 583
    msSQL = ""

    msSQL = msSQL & " SELECT DataItemCode,DataItemName,DataItemID"
    msSQL = msSQL & " FROM DataItem"
    msSQL = msSQL & " WHERE DataItem.ClinicalTrialId  = " & mlTrialID
    msSQL = msSQL & " AND DataItem.DataType = 1 "
    msSQL = msSQL & " ORDER BY DataItemCode "
    
    Set rsStudyCategory = New ADODB.Recordset
    rsStudyCategory.Open msSQL, MacroADODBConnection, adOpenKeyset
    
    nCount = rsStudyCategory.RecordCount
    If nCount <= 0 Then
        Call DialogInformation(" There are no Category type questions in this study to print.")
        Call HourglassOff
        Exit Sub
    End If
    
    '01/11/2001 changed True to false Bug Report Number 115 Ash.
    '23/11/2001 changed to true. Buglist no.14
    On Error Resume Next
    
    frmMenu.CommonDialog1.CancelError = True
    Printer.Orientation = vbPRORPortrait
    Printer.TrackDefault = True
    frmMenu.CommonDialog1.ShowPrinter
    Printer.Orientation = vbPRORPortrait
    
    'Check for errors in ShowPrinter (including a cancel)
    If Err.Number > 0 Then
            HourglassOff
            Exit Sub
    End If
    
    'hourglass displayed
    
    DoEvents
    
    On Error GoTo ErrHandler
    
    'extend the top and left margins by 1/2 inch to 3/4 inch approx
    Printer.ScaleLeft = -720
    Printer.ScaleTop = -720
    
    mlPrintingWidth = Printer.ScaleWidth - 1440
    Call PrintHeader(mlPrintingWidth, msPrintName)
    
    'print study Name
    Call PrintStudyName(msTrialName)
    
    rsStudyCategory.MoveFirst
    Do While Not rsStudyCategory.EOF
        
        'continue printing on a new page
        If IsPageSizeEnough(mnPrintBlock) = False Then
            Printer.NewPage
            Call PrintHeader(mlPrintingWidth, msPrintName)
        End If
        
        'Print sub-headers for questions
        Printer.FontName = "Tahoma"
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 200
        Printer.FontSize = 8
        
        'print question name
        Printer.FontBold = True
        Printer.Print "Question Name: ", ;
        Printer.FontBold = False
        Printer.CurrentX = 2000
        Printer.Print rsStudyCategory.Fields(1).Value
        
        'print question code
        Printer.FontBold = True
        Printer.CurrentX = 820
        Printer.Print "Code: ", ;
        Printer.FontBold = False
        Printer.CurrentX = 2000
        Printer.Print rsStudyCategory.Fields(0).Value
        
        'print  header for report details
        Printer.CurrentX = 2000
        Printer.CurrentY = Printer.CurrentY + 200
        Printer.FontBold = True
        Printer.Print "Value Code ", ;
        Printer.CurrentX = 4000
        Printer.Print "Value ", ;
        Printer.CurrentX = 6000
        Printer.Print "Active "
        Printer.FontBold = False
        
        msSQL1 = " SELECT ValueCode,ItemValue,Active"
        msSQL1 = msSQL1 & " FROM ValueData "
        msSQL1 = msSQL1 & " WHERE ValueData.ClinicalTrialId  =  " & mlTrialID
        msSQL1 = msSQL1 & " AND ValueData.DataItemID  = " & rsStudyCategory.Fields(2).Value
        
        Set rsStudyDefCategoryValues = New ADODB.Recordset
        rsStudyDefCategoryValues.Open msSQL1, MacroADODBConnection, adOpenKeyset

        nCount1 = rsStudyDefCategoryValues.RecordCount
        If nCount1 <= 0 Then
            Call DialogInformation(" There are no Category type questions in this study to print.")
            Call HourglassOff
            Exit Sub
        End If

        rsStudyDefCategoryValues.MoveFirst
        Do While Not rsStudyDefCategoryValues.EOF
            Printer.FontSize = 8
            Printer.CurrentX = 2000
            Printer.Print rsStudyDefCategoryValues.Fields(0).Value, ;
            Printer.CurrentX = 4000
            Printer.Print rsStudyDefCategoryValues.Fields(1).Value, ;
            Printer.CurrentX = 6000
            Printer.Print rsStudyDefCategoryValues.Fields(2).Value
            rsStudyDefCategoryValues.MoveNext
        Loop
            'draw line
            Printer.DrawWidth = 2
            Printer.CurrentY = Printer.CurrentY + 120
            mnCurrentY = Printer.CurrentY
            Printer.Line (0, mnCurrentY)-(mlPrintingWidth, mnCurrentY)
            
            rsStudyCategory.MoveNext
    Loop
    
        Set rsStudyDefCategoryValues = Nothing
        Set rsStudyCategory = Nothing
        Printer.EndDoc
        Call HourglassOff

Exit Sub
ErrHandler:
    
    If Err.Number = 482 Then
        DialogError ("Printer error.Check printer and retry"), "Printer Error"
        Call HourglassOff
        Exit Sub
    End If
    
    Call HourglassOff
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Print_QuestionWithinEformReport", "modPrintListings.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub
'--------------------------------------------------------------------------------------------------
Public Sub Print_QuestionListReport()
'--------------------------------------------------------------------------------------------------
'called from frmMenu to print all questions for a study
'--------------------------------------------------------------------------------------------------
Dim nQuestions As Integer
Dim rsStudyDefQuestionList As New ADODB.Recordset
Dim mlTrialID As Long
Dim msTrialName As String
Dim mlPrintingWidth As Long
Dim msPrintName As String
Dim msSQL As String
Dim msSQL1 As String
Dim mnPrintBlock As Integer

    
    On Error GoTo ErrHandler
    
    HourglassOn
    
    msPrintName = "Study Definition Question List Reports"
    mlTrialID = frmMenu.ClinicalTrialId
    msTrialName = frmMenu.ClinicalTrialName
    mnPrintBlock = 200
    msSQL = ""
    
    msSQL = "SELECT  ClinicalTrial.ClinicalTrialName,DataItem.DataItemName,DataItem.DataItemCode"
    msSQL = msSQL & " FROM  ClinicalTrial,DataItem "
    msSQL = msSQL & " Where  ClinicalTrial.ClinicalTrialID = DataItem.ClinicalTrialID"
    msSQL = msSQL & " AND ClinicalTrial.ClinicalTrialId = " & mlTrialID
    msSQL = msSQL & " AND DataItem.ClinicalTrialId = " & mlTrialID
    msSQL = msSQL & " ORDER BY DataItem.DataItemName"
    
    Set rsStudyDefQuestionList = New ADODB.Recordset
    rsStudyDefQuestionList.Open msSQL, MacroADODBConnection, adOpenKeyset
    
    nQuestions = rsStudyDefQuestionList.RecordCount
    If nQuestions <= 0 Then
        Call DialogInformation(" There are no questions to print in this study.")
        Call HourglassOff
        Exit Sub
    End If
    
    '01/11/2001 changed True to false Bug Report Number 115 Ash.
    '23/11/2001 changed back to True Buglist no.14 Ash
    On Error Resume Next
    
    frmMenu.CommonDialog1.CancelError = True
    Printer.Orientation = vbPRORPortrait
    Printer.TrackDefault = True
    frmMenu.CommonDialog1.ShowPrinter
    Printer.Orientation = vbPRORPortrait
    
    'Check for errors in ShowPrinter (including a cancel)
   If Err.Number > 0 Then
        HourglassOff
        Exit Sub
    End If
    
    'hourglass displayed
    DoEvents
     
    On Error GoTo ErrHandler
    
    'extend the top and left margins by 1/2 inch to 3/4 inch approx
    Printer.ScaleLeft = -720
    Printer.ScaleTop = -720
    
    mlPrintingWidth = Printer.ScaleWidth - 1440
    Call PrintHeader(mlPrintingWidth, msPrintName)
    
    'print study name
    Call PrintStudyName(msTrialName)
    
    'print  header for report details
    Printer.FontName = "Tahoma"
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 240
    Printer.FontSize = 8
    Printer.FontBold = True
    Printer.Print "Question Name ", ;
    Printer.CurrentX = 5000
    Printer.Print "Question Code "
    Printer.FontBold = False
    
    rsStudyDefQuestionList.MoveFirst
    Do While Not rsStudyDefQuestionList.EOF
            
        'continue printing on a new page
        If IsPageSizeEnough(mnPrintBlock) = False Then
            Printer.NewPage
            Call PrintHeader(mlPrintingWidth, msPrintName)
        End If
        Printer.CurrentX = 0
        Printer.Print rsStudyDefQuestionList.Fields(1).Value, ;
        Printer.CurrentX = 5000
        Printer.Print rsStudyDefQuestionList.Fields(2).Value
        rsStudyDefQuestionList.MoveNext
    Loop
        
    Set rsStudyDefQuestionList = Nothing
    Printer.EndDoc
    Call HourglassOff

Exit Sub
ErrHandler:
    
    If Err.Number = 482 Then
        DialogError ("Printer error.Check printer and retry"), "Printer Error"
        Call HourglassOff
        Exit Sub
    End If
  
  Call HourglassOff
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Print_QuestionDefinitionReport", "modPrintListings.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'--------------------------------------------------------------------------------------------------
Public Sub Print_QuestionDefinitionReport()
'--------------------------------------------------------------------------------------------------
'Called from within frmMenu to print all questions and their definitions
'within a study
'--------------------------------------------------------------------------------------------------
Dim rsStudyDefQuestionDefs As ADODB.Recordset
Dim rsValidations As ADODB.Recordset
Dim nCount As Integer
Dim nCount1 As Integer
Dim sType As String
Dim nDerivationWidth As Long
Dim nHelpTextWidth As Long
Dim nValidationWidth As Long
Dim mlTrialID As Long
Dim msPrintText As String
Dim mlPrintTextWidth As Long
Dim mnPrintBlock As Integer
Dim msTrialName As String
Dim mlPrintingWidth As Long
Dim msPrintName As String
Dim msSQL As String
Dim msSQL1 As String
Dim mnCurrentY As Integer
Dim i As Integer
Dim mntotal As Integer
Dim msTextToBreak As String
    
    
    On Error GoTo ErrHandler
    
    HourglassOn
    
    mlTrialID = frmMenu.ClinicalTrialId
    msPrintName = "Question Definition Reports"
    msTrialName = frmMenu.ClinicalTrialName
    mnPrintBlock = 1183
    msSQL = ""
    
    msSQL = "SELECT DataItem.DataItemName,DataItem.DataItemCode,DataItem.DataType,DataItem.DataItemFormat,"
    msSQL = msSQL & " DataItem.DataItemLength,DataItem.UnitOfMeasurement,DataItem.Derivation,DataItem.DataItemHelpText,"
    msSQL = msSQL & " ClinicalTrial.ClinicalTrialName,DataItem.DataItemID"
    msSQL = msSQL & " FROM DataItem,ClinicalTrial"
    msSQL = msSQL & " Where DataItem.ClinicalTrialId = " & mlTrialID
    msSQL = msSQL & " AND ClinicalTrial.ClinicalTrialId = " & mlTrialID
    msSQL = msSQL & " AND ClinicalTrial.ClinicalTrialId = DataItem.ClinicalTrialId"
    msSQL = msSQL & " ORDER BY DataItemCode"
    
    Set rsStudyDefQuestionDefs = New ADODB.Recordset
    rsStudyDefQuestionDefs.Open msSQL, MacroADODBConnection, adOpenKeyset

    nCount = rsStudyDefQuestionDefs.RecordCount
    If nCount <= 0 Then
        Call DialogInformation(" There are no question definitions in this study to print.")
        Call HourglassOff
        Exit Sub
    End If
    
    '01/11/2001 changed True to false Bug Report Number 115 Ash.
    '23/11/2001 changed to true. Buglist no.14
    On Error Resume Next
    
    frmMenu.CommonDialog1.CancelError = True
    Printer.Orientation = vbPRORPortrait
    Printer.TrackDefault = True
    frmMenu.CommonDialog1.ShowPrinter
    Printer.Orientation = vbPRORPortrait
    
    'Check for errors in ShowPrinter (including a cancel)
    If Err.Number > 0 Then
        HourglassOff
        Exit Sub
    End If
    
    'hourglass displayed
    DoEvents
    
    On Error GoTo ErrHandler
    
    'extend the top and left margins by 1/2 inch to 3/4 inch approx
    Printer.ScaleLeft = -720
    Printer.ScaleTop = -720
    
    mlPrintingWidth = Printer.ScaleWidth - 1440
    Call PrintHeader(mlPrintingWidth, msPrintName)
    
    'print study Name
    Call PrintStudyName(msTrialName)
    
    Do While Not rsStudyDefQuestionDefs.EOF
        
        'get validation conditions to be used later below
        'need to get it here to calculate page space required
        msSQL1 = " SELECT DataItemValidation.DataItemValidation"
        msSQL1 = msSQL1 & " FROM DataItemValidation "
        msSQL1 = msSQL1 & " WHERE DataItemValidation.ClinicalTrialId  =  " & mlTrialID
        msSQL1 = msSQL1 & " AND DataItemValidation.DataItemID  = " & rsStudyDefQuestionDefs.Fields(9).Value
        
        Set rsValidations = New ADODB.Recordset
        rsValidations.Open msSQL1, MacroADODBConnection, adOpenKeyset, adLockOptimistic
        nCount1 = rsValidations.RecordCount
        
        'get number of lines
        nDerivationWidth = Len(RemoveNull(rsStudyDefQuestionDefs.Fields(6).Value)) / 100
        mlPrintTextWidth = mlPrintingWidth - 1440
        nHelpTextWidth = Len(RemoveNull(rsStudyDefQuestionDefs.Fields(7).Value)) / 100
        nValidationWidth = Len(RemoveNull(rsValidations.Fields(0).Value)) / 100
        
        'continue printing on a new page if page not big enough
        'currently harded coded since i am having difficulty in
        'getting IsPageSizeEnough to work correctly in this sub
        mntotal = _
        (Printer.CurrentY + mnPrintBlock) + (190 * (nDerivationWidth + nHelpTextWidth + nValidationWidth))
        If IsPageSizeEnough _
        (mnPrintBlock, nHelpTextWidth, nDerivationWidth, nValidationWidth) = False _
        Or mntotal > 13000 Then
            Printer.NewPage
            Call PrintHeader(mlPrintingWidth, msPrintName)
        End If
        
        'printer settings
        Printer.FontName = "Tahoma"
        Printer.FontSize = 10
        Printer.CurrentX = 0
        
        'Question name
        Printer.FontBold = True
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.Print "Question Name: ", ;
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.CurrentX = 2200
        Printer.CurrentY = Printer.CurrentY + 10
        Printer.Print rsStudyDefQuestionDefs.Fields(0).Value
        
        'Question Code
        Printer.FontSize = 10
        Printer.CurrentX = 1030
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Code: ", ;
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.CurrentX = 2200
        Printer.Print rsStudyDefQuestionDefs.Fields(1).Value
        
        'Question DataType
        Printer.CurrentX = 1400
        Printer.CurrentY = Printer.CurrentY + 200
        sType = DecodeDataType(rsStudyDefQuestionDefs.Fields(2).Value)
        Printer.FontBold = True
        Printer.Print "DataType: ", ;
        Printer.FontBold = False
        Printer.CurrentX = 2300
        Printer.Print sType
        
        'Print Question Format
        Printer.CurrentX = 1400
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Format: ", ;
        Printer.FontBold = False
        sType = RemoveNull(rsStudyDefQuestionDefs.Fields(3).Value)
        Printer.Print sType
        
        'Print Question Length
        Printer.CurrentX = 1400
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Length: ", ;
        Printer.FontBold = False
        Printer.Print rsStudyDefQuestionDefs.Fields(4).Value
        
        'Print Question Unit Of Measurement
        Printer.CurrentX = 1400
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Unit Of Measurement ", ;
        Printer.FontBold = False
        sType = RemoveNull(rsStudyDefQuestionDefs.Fields(5).Value)
        Printer.Print sType
        
        'Print Question Validation condition
        Printer.CurrentX = 1400
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Validation Condition: "
        Printer.FontBold = False
        'start printing validations
        For i = 1 To rsValidations.RecordCount 'nCount1
            msPrintText = RemoveNull(rsValidations.Fields(0).Value)
            Printer.CurrentX = 1400
            If Len(msPrintText) > 500 Then
                Call PrintMultiLine(msPrintText, mlPrintTextWidth, 1440)
                Printer.CurrentX = 1400
            Else
                Printer.CurrentX = 1440
                Printer.Print msPrintText
            End If
        Next
       
        'Print Question Derivation
        Printer.CurrentX = 1400
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Derivation: "
        Printer.FontBold = False
        msPrintText = RemoveNull(rsStudyDefQuestionDefs.Fields(6).Value)
        Printer.CurrentX = 1440
        Printer.CurrentY = Printer.CurrentY + 240
        If Printer.TextWidth(msPrintText) > mlPrintTextWidth Then
           msPrintText = SplitPrintTextLine(msPrintText, mlPrintTextWidth, 1440)
        End If
        Printer.CurrentX = 1440
        Printer.Print msPrintText
       
        'Print Question Helptext
        Printer.CurrentX = 1400
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "HelpText: "
        Printer.FontBold = False
        msPrintText = RemoveNull(rsStudyDefQuestionDefs.Fields(7).Value)
        Printer.CurrentX = 1440
        Printer.CurrentY = Printer.CurrentY + 240
        If Printer.TextWidth(msPrintText) > mlPrintTextWidth Then
           msPrintText = SplitPrintTextLine(msPrintText, mlPrintTextWidth, 1440)
        End If
        Printer.CurrentX = 1440
        Printer.Print msPrintText
        
        'draw line
        Printer.DrawWidth = 2
        Printer.CurrentY = Printer.CurrentY + 120
        mnCurrentY = Printer.CurrentY
        Printer.Line (0, mnCurrentY)-(mlPrintingWidth, mnCurrentY)
        
        rsStudyDefQuestionDefs.MoveNext
    Loop
    
    Set rsStudyDefQuestionDefs = Nothing
    Printer.EndDoc
    Call HourglassOff
    

Exit Sub
ErrHandler:

    Call HourglassOff
    If Err.Number = 482 Then
        DialogError ("Printer error.Check printer and retry"), "Printer Error"
        Exit Sub
    End If
  
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Print_QuestionDefinitionReport", "modPrintListings.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub
'--------------------------------------------------------------------------------------------------
Public Sub Print_StudyDefinitionReport()
'--------------------------------------------------------------------------------------------------
'called from frmMenu to print all the studies in a database
'--------------------------------------------------------------------------------------------------
Dim rsStudyDefs As ADODB.Recordset
Dim nCount As Integer
Dim sType As String
Dim mlTrialID As Long
Dim msPrintText As String
Dim mnPrintTextWidth As Long
Dim mnPrintBlock As Integer
Dim msTrialName As String
Dim mlPrintingWidth As Long
Dim msPrintName As String
Dim msSQL As String
Dim msSQL1 As String
Dim mnCurrentY As Integer

    On Error GoTo ErrHandler
    
    HourglassOn
    
    msPrintName = "List Of Studies Reports"
    mlTrialID = frmMenu.ClinicalTrialId
    msTrialName = frmMenu.ClinicalTrialName
    mnPrintBlock = 540
    msSQL = ""
    
    msSQL = msSQL & "  SELECT ClinicalTrialName,StatusID,PhaseID,ClinicalTrialDescription,"
    msSQL = msSQL & " KeyWords,ExpectedRecruitment, ClinicalTrialID, TrialTypeID"
    msSQL = msSQL & " FROM ClinicalTrial "
    msSQL = msSQL & " WHERE ClinicalTrial.ClinicalTrialID > 0 "
    msSQL = msSQL & " ORDER BY ClinicalTrialName "
    
    Set rsStudyDefs = New ADODB.Recordset
    rsStudyDefs.Open msSQL, MacroADODBConnection, adOpenKeyset

    nCount = rsStudyDefs.RecordCount
    If nCount <= 0 Then
        Call DialogInformation(" There are no studies to print.")
        Call HourglassOff
        Exit Sub
    End If
    
    '01/11/2001 changed True to false Bug Report Number 115 Ash.
    '23/11/2001 changed to true. Buglist no.14
    On Error Resume Next
    
    frmMenu.CommonDialog1.CancelError = True
    Printer.Orientation = vbPRORPortrait
    Printer.TrackDefault = True
    frmMenu.CommonDialog1.ShowPrinter
    Printer.Orientation = vbPRORPortrait
    
    'Check for errors in ShowPrinter (including a cancel)
    If Err.Number > 0 Then
        HourglassOff
        Exit Sub
    End If
    
    'hourglass displayed
    DoEvents
    
    On Error GoTo ErrHandler
    
    'extend the top and left margins by 1/2 inch to 3/4 inch approx
    Printer.ScaleLeft = -720
    Printer.ScaleTop = -720
    
    mlPrintingWidth = Printer.ScaleWidth - 1440
    Call PrintHeader(mlPrintingWidth, msPrintName)
    
    rsStudyDefs.MoveFirst
    Do While Not rsStudyDefs.EOF
        
        'continue printing on a new page
        If IsPageSizeEnough(mnPrintBlock) = False Then
            Printer.NewPage
            Call PrintHeader(mlPrintingWidth, msPrintName)
        End If
        
        'printer settings
        Printer.FontName = "Tahoma"
        Printer.FontSize = 10
        Printer.CurrentX = 0
        
        'Study Name
        Printer.FontBold = True
        Printer.CurrentY = Printer.CurrentY + 60
        Printer.Print "Study Name: ", ;
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.CurrentX = 2500
        Printer.CurrentY = Printer.CurrentY + 10
        Printer.Print rsStudyDefs.Fields(0).Value
        
        'Status
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Status: ", ;
        Printer.FontBold = False
        Printer.CurrentX = 2500
        Printer.FontSize = 8
        sType = DecodeStatusID(rsStudyDefs.Fields(1).Value)
        Printer.Print sType
        
        'Phase
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Phase: ", ;
        Printer.FontBold = False
        Printer.CurrentX = 2500
        Printer.FontSize = 8
        sType = DecodePhaseID(rsStudyDefs.Fields(2).Value)
        Printer.Print sType
        
        'Description
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Description: ", ;
        Printer.FontBold = False
        mnPrintTextWidth = mlPrintingWidth - 1440
        Printer.FontSize = 8
        Printer.CurrentX = 2500
        msPrintText = RemoveNull(rsStudyDefs.Fields(3).Value)
        If Printer.TextWidth(msPrintText) > mnPrintTextWidth Then
           msPrintText = SplitPrintTextLine(msPrintText, mnPrintTextWidth, 1440)
        End If
        Printer.Print msPrintText
        
        'Keywords
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Keywords: ", ;
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.CurrentX = 2500
        Printer.Print rsStudyDefs.Fields(4).Value
        
        'Expected Recruitment
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Expected Recruitment: ", ;
        Printer.FontBold = False
        Printer.CurrentX = 2500
        Printer.FontSize = 8
        Printer.Print rsStudyDefs.Fields(5).Value
        
        'Actual Recruitment
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Actual Recruitment: ", ;
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.CurrentX = 2500
        Printer.Print GetStudyRecruitment(rsStudyDefs.Fields(6).Value)
        
        'Trial Type
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Trial Type: ", ;
        Printer.FontBold = False
        Printer.CurrentX = 2500
        Printer.FontSize = 8
        sType = DecodeTrialType(rsStudyDefs.Fields(7).Value)
        Printer.Print sType
        
        'draw line
        Printer.DrawWidth = 2
        Printer.CurrentY = Printer.CurrentY + 120
        mnCurrentY = Printer.CurrentY
        Printer.Line (0, mnCurrentY)-(mlPrintingWidth, mnCurrentY)
        
        rsStudyDefs.MoveNext
    
    Loop
        
    Set rsStudyDefs = Nothing
    Printer.EndDoc
    Call HourglassOff

Exit Sub
ErrHandler:

    If Err.Number = 482 Then
        DialogError ("Printer error.Check printer and retry"), "Printer Error"
        Call HourglassOff
        Exit Sub
    End If
  
  Call HourglassOff
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Print_StudyDefinitionReport", "modPrintListings.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'--------------------------------------------------------------------------------------------------
Private Function GetStudyRecruitment(lStudyId As Long) As Long
'--------------------------------------------------------------------------------------------------
'Calculate StudyRecruitment
'--------------------------------------------------------------------------------------------------
Dim sSQL As String
Dim rs As ADODB.Recordset

    sSQL = "select count(*) from trialsubject where clinicaltrialid = " & lStudyId
    
    Set rs = New ADODB.Recordset
    rs.Open sSQL, MacroADODBConnection
    GetStudyRecruitment = rs.Fields(0).Value
    rs.Close
    Set rs = Nothing
    
End Function
'--------------------------------------------------------------------------------------------------
Private Function SplitPrintTextLine(ByRef sMessageLine As String, _
                                    nWidth As Long, _
                                    nPrintFrom As Integer) As String
'--------------------------------------------------------------------------------------------------
'Mo Morris's routine in Macro_DM
'--------------------------------------------------------------------------------------------------
Dim i As Integer
Dim sPart As String
Dim sPreviousPart
Dim sChar As String
Dim q As Integer

    On Error GoTo ErrHandler
    'to handle the manner in which this function works a space is added to the textline
    'unless there is one already
    If Mid(sMessageLine, Len(sMessageLine), 1) <> " " Then
        sMessageLine = sMessageLine & " "
    End If
    sPart = ""
    sPreviousPart = ""
    For i = 1 To Len(sMessageLine)
        sChar = Mid(sMessageLine, i, 1)
        sPart = sPart & sChar
        If sChar = " " Then
            If Printer.TextWidth(sPart) > nWidth Then
                Printer.CurrentX = nPrintFrom
               'Check for the situation where no spaces have been reached and the textwidth is beyond nWidth.
                'In this situation sPreviousPart would be empty and would need to have a truncated
                'section of sMessageLine placed in it
                If sPreviousPart = "" Then
                    q = Len(sMessageLine)
                    Do
                        q = q - 1
                        sPreviousPart = Mid(sMessageLine, 1, q)
                    Loop Until Printer.TextWidth(sPreviousPart) < nWidth
                End If
                'Printer.CurrentX = 1440
                Printer.Print sPreviousPart
                sMessageLine = Mid(sMessageLine, Len(sPreviousPart) + 1)
                Exit For
            Else
                sPreviousPart = sPart
            End If
        End If
    Next
    
    'check to see whether the remaining part of PrintText requires a recursive call to SplitPrintTextLine
    If Printer.TextWidth(sMessageLine) < nWidth Then
        SplitPrintTextLine = sMessageLine
    Else
        SplitPrintTextLine = SplitPrintTextLine(sMessageLine, nWidth, nPrintFrom)
    End If
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "SplitPrintTextLine", "modPrintListings.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function
'--------------------------------------------------------------------------------------------------
Private Sub PrintHeader(PrintingWidth As Long, msPrintName As String)
'--------------------------------------------------------------------------------------------------
'Prints the Headings
'--------------------------------------------------------------------------------------------------
Dim sHeaderLine As String
Dim nTextHeight1 As Integer
Dim nTextHeight2 As Integer
Dim nTextHeight3 As Integer
Dim lErrNum As Integer
Dim sErrDesc As String
Dim mlTrialID As Long
Dim msSQL As String
Dim msSQL1 As String
Dim mnCurrentY As Integer
    
    On Error GoTo ErrHandler
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.FontName = "Arial Narrow"
    Printer.FontSize = 30
    Printer.FontBold = True
    Printer.Print "MACRO", ;
    Printer.FontBold = False
    
    nTextHeight1 = Printer.TextHeight("X")
    
    Printer.FontSize = 16
    Printer.FontItalic = True
    nTextHeight2 = Printer.TextHeight("X")
    Printer.CurrentX = 2200
    Printer.CurrentY = Printer.CurrentY + ((nTextHeight1 - nTextHeight2) * 0.8)
    Printer.Print msPrintName, ;
    Printer.FontItalic = False
    
    Printer.FontSize = 8
    nTextHeight3 = Printer.TextHeight("X")
    Printer.CurrentY = Printer.CurrentY + ((nTextHeight2 - nTextHeight3) * 0.8)
    sHeaderLine = "Printed " & Format(Now, "yyyy/mm/dd hh:mm:ss") & "    Page " & Printer.Page
    Printer.CurrentX = PrintingWidth - Printer.TextWidth(sHeaderLine)
    Printer.Print sHeaderLine
    
    'draw a thicker line across page
    Printer.DrawWidth = 6
    Printer.CurrentY = Printer.CurrentY + 60
    mnCurrentY = Printer.CurrentY
    Printer.Line (0, mnCurrentY)-(PrintingWidth, mnCurrentY)
    
    Printer.DrawWidth = 1
    Printer.CurrentY = Printer.CurrentY + 60
    mnCurrentY = Printer.CurrentY

Exit Sub
ErrHandler:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & "|" & "modPrintlistings.PrintHeader"
End Sub
'--------------------------------------------------------------------------------------------------
Public Function DecodeDataType(nType As Long) As String
'--------------------------------------------------------------------------------------------------
'decodes data type
'--------------------------------------------------------------------------------------------------
Dim sType As String
Dim sretType As String
    
    sretType = ""
    Select Case nType
        Case Is = 0
            sType = "Text"
        Case Is = 1
            sType = "Category"
        Case Is = 2
            sType = "Integer Number"
        Case Is = 3
            sType = "Real Number"
        Case Is = 4
            sType = "Date/Time"
        Case Is = 5
            sType = "Multimedia"
        Case Is = 6
            sType = "Laboratory Test"
    End Select
    
    DecodeDataType = sType

End Function
'--------------------------------------------------------------------------------------------------
Public Function IsPageSizeEnough(mnPrintBlock As Integer, _
                                Optional nExtraWidth1 As Long, _
                                Optional nExtraWidth2 As Long, _
                                Optional nExtraWidth3 As Long) As Boolean
'--------------------------------------------------------------------------------------------------
'checks to see if enough space is on the page to print required text. If not
'opens a new page and prints header followed by the required text
'--------------------------------------------------------------------------------------------------
Dim nTotalPageLength As Integer
Dim nRequiredPrintLines As Integer

    On Error GoTo ErrHandler
    nTotalPageLength = 13680   '9.5 * 1440 (page length * twips)
    nRequiredPrintLines = (Printer.CurrentY + mnPrintBlock) + (190 * (nExtraWidth1 + nExtraWidth2 + nExtraWidth3))
    
    If nRequiredPrintLines > nTotalPageLength Then
        IsPageSizeEnough = False
    Else
        IsPageSizeEnough = True
    End If

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "IsPageSizeEnough", "modPrintListings.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function
'--------------------------------------------------------------------------------------------------
Public Sub PrintStudyName(sStudyName As String)
'--------------------------------------------------------------------------------------------------
'prints study name
'--------------------------------------------------------------------------------------------------
Dim lErrNum As Integer
Dim sErrDesc As String
Dim mlTrialID As Long
Dim mlPrintingWidth As Long
Dim msPrintName As String
Dim msSQL As String
Dim msSQL1 As String
Dim mnCurrentY As Integer

On Error GoTo ErrHandler
    
    Printer.FontName = "Arial"
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 120
    Printer.FontSize = 16
    Printer.FontBold = True
    Printer.Print "Study Name: ", ;
    Printer.FontBold = False
    Printer.FontSize = 14
    Printer.CurrentY = Printer.CurrentY + 50
    Printer.CurrentX = 2000
    Printer.Print sStudyName
    'draw first line
    Printer.DrawWidth = 2
    Printer.CurrentY = Printer.CurrentY + 120
    mnCurrentY = Printer.CurrentY
    Printer.Line (0, mnCurrentY)-(mlPrintingWidth, mnCurrentY)
    'draw second line
    Printer.DrawWidth = 2
    Printer.CurrentY = Printer.CurrentY + 60
    mnCurrentY = Printer.CurrentY
    Printer.Line (0, mnCurrentY)-(mlPrintingWidth, mnCurrentY)

Exit Sub
ErrHandler:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & "|" & "modPrintlistings.PrintStudyName"
End Sub
'--------------------------------------------------------------------------------------------------
Public Function DecodeStatusID(nType As Long) As String
'--------------------------------------------------------------------------------------------------
'To get the appropriate StatusID Type
'--------------------------------------------------------------------------------------------------
Dim sType As String
Dim sretType As String
    
    sretType = ""
    Select Case nType
        Case Is = 1
            sType = "In Preparation"
        Case Is = 2
            sType = "Open"
        Case Is = 3
            sType = "Closed To Recruitment"
        Case Is = 4
            sType = "Closed To Follow Up"
        Case Is = 5
            sType = "Suspended"
        Case Is = 6
            sType = "Deleted"
    End Select
        
        DecodeStatusID = sType

End Function
'--------------------------------------------------------------------------------------------------
Public Function DecodePhaseID(nType As Long) As String
'--------------------------------------------------------------------------------------------------
'To get the appropriate phaseid Type
'--------------------------------------------------------------------------------------------------
Dim sType As String
Dim sretType As String
    
    sretType = ""
    Select Case nType
        Case Is = 1
            sType = "I"
        Case Is = 2
            sType = "II"
        Case Is = 3
            sType = "III"
        Case Is = 4
            sType = "IV"
        End Select
    
    DecodePhaseID = sType

End Function
'--------------------------------------------------------------------------------------------------
Public Function DecodeTrialType(nType As Long) As String
'--------------------------------------------------------------------------------------------------
'To get the appropriate Trial Type
'--------------------------------------------------------------------------------------------------
Dim sType As String
Dim rsTrialType As ADODB.Recordset
Dim msSQL As String

    If nType = 0 Then
        sType = " "
    Else
        msSQL = "Select TrialTypeName,TrialTypeID "
        msSQL = msSQL & " FROM TrialType"
        msSQL = msSQL & " WHERE TrialTypeID = " & nType
        
        Set rsTrialType = New ADODB.Recordset
        rsTrialType.Open msSQL, MacroADODBConnection, adOpenKeyset
        
        sType = rsTrialType.Fields(0).Value
    
    End If
    
    DecodeTrialType = sType

End Function

'------------------------------------------------------------------------------------------------
Public Sub PrintMultiLine(ByRef sMessageLine As String, _
                                    nWidth As Long, _
                                    nPrintFrom As Integer)
'------------------------------------------------------------------------------------------------
'routine based on SplitPrintTextLine
'------------------------------------------------------------------------------------------------
Dim i As Integer
Dim sPart As String
Dim sPreviousPart
Dim sChar As String
Dim q As Integer

    On Error GoTo ErrHandler
    If sMessageLine = "" Then Exit Sub
    'to handle the manner in which this function works a space is added to the textline
    'unless there is one already
    If Mid(sMessageLine, Len(sMessageLine), 1) <> " " Then
        sMessageLine = sMessageLine & " "
    End If
    sPart = ""
    sPreviousPart = ""
    For i = 1 To Len(sMessageLine)
        sChar = Mid(sMessageLine, i, 1)
        sPart = sPart & sChar
        If sChar = " " Then
            If Printer.TextWidth(sPart) > nWidth Then
                Printer.CurrentX = nPrintFrom
               'Check for the situation where no spaces have been reached and the textwidth is beyond nWidth.
                'In this situation sPreviousPart would be empty and would need to have a truncated
                'section of sMessageLine placed in it
                If sPreviousPart = "" Then
                    q = Len(sMessageLine)
                    Do
                        q = q - 1
                        sPreviousPart = Mid(sMessageLine, 1, q)
                    Loop Until Printer.TextWidth(sPreviousPart) < nWidth
                End If
                Printer.Print sPreviousPart
                sMessageLine = Mid(sMessageLine, Len(sPreviousPart) + 1)
                Exit For
            Else
                sPreviousPart = sPart
            End If
        End If
    Next
    
    'check to see whether any text left
    '82 is the smallest printable width
    If Len(sMessageLine) > 82 Then
        Call PrintMultiLine(sMessageLine, nWidth, nPrintFrom)
    Else
        Printer.CurrentX = nPrintFrom
        Printer.Print sMessageLine
    End If

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintMultiLine", "modPrintListings.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub
'------------------------------------------------------------------------------------------
Public Sub Print_UnitConversion()
'------------------------------------------------------------------------------------------
'ASH 13/05/2002
'------------------------------------------------------------------------------------------
Dim nUnits As Integer
Dim rsUnits As New ADODB.Recordset
Dim mlTrialID As Long
Dim msTrialName As String
Dim mlPrintingWidth As Long
Dim msPrintName As String
Dim msSQL As String
Dim msSQL1 As String
Dim mnPrintBlock As Integer

    On Error GoTo ErrHandler
    
    HourglassOn
    
    msPrintName = "Unit Conversion Factors Reports"
    mlTrialID = frmMenu.ClinicalTrialId
    msTrialName = frmMenu.ClinicalTrialName
    mnPrintBlock = 200
    msSQL = ""
    
    msSQL = "SELECT  * from UnitConversionFactors ORDER BY UnitClass "
   
    Set rsUnits = New ADODB.Recordset
    rsUnits.Open msSQL, MacroADODBConnection, adOpenKeyset
    
    nUnits = rsUnits.RecordCount
    If nUnits <= 0 Then
        Call DialogInformation(" There are no unit conversion factors to print in this study.")
        Call HourglassOff
        Exit Sub
    End If
    
    On Error Resume Next
    
    frmMenu.CommonDialog1.CancelError = True
    Printer.Orientation = vbPRORPortrait
    Printer.TrackDefault = True
    frmMenu.CommonDialog1.ShowPrinter
    Printer.Orientation = vbPRORPortrait
    
    'Check for errors in ShowPrinter (including a cancel)
   If Err.Number > 0 Then
        HourglassOff
        Exit Sub
    End If
    
    'hourglass displayed
    DoEvents
     
    On Error GoTo ErrHandler
    
    'extend the top and left margins by 1/2 inch to 3/4 inch approx
    Printer.ScaleLeft = -720
    Printer.ScaleTop = -720
    
    mlPrintingWidth = Printer.ScaleWidth - 1440
    Call PrintHeader(mlPrintingWidth, msPrintName)
    
    'print study name
    Call PrintStudyName(msTrialName)
    
    'print header for report
    Printer.FontName = "Tahoma"
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 240
    Printer.FontSize = 8
    Printer.FontBold = True
    Printer.Print "Unit Class ", ;
    Printer.CurrentX = 3000
    Printer.Print "From Unit ", ;
    Printer.CurrentX = 6000
    Printer.Print "To Unit ", ;
    Printer.CurrentX = 8500
    Printer.Print "Conversion Factor ", ;
    Printer.FontBold = False
    Printer.CurrentY = Printer.CurrentY + 240
    
    rsUnits.MoveFirst
    Do While Not rsUnits.EOF
        'continue printing on a new page
        If IsPageSizeEnough(mnPrintBlock) = False Then
            Printer.NewPage
            Call PrintHeader(mlPrintingWidth, msPrintName)
        End If
        'print details
        Printer.CurrentX = 0
        Printer.Print rsUnits.Fields(3).Value, ;
        Printer.CurrentX = 3000
        Printer.Print rsUnits.Fields(0).Value, ;
        Printer.CurrentX = 6000
        Printer.Print rsUnits.Fields(1).Value, ;
        Printer.CurrentX = 8500
        Printer.Print rsUnits.Fields(2).Value
        rsUnits.MoveNext
    Loop
    
    Set rsUnits = Nothing
    Printer.EndDoc
    Call HourglassOff
    
    Exit Sub
ErrHandler:
    
    If Err.Number = 482 Then
        DialogError ("Printer error.Check printer and retry"), "Printer Error"
        Call HourglassOff
        Exit Sub
    End If
  
  Call HourglassOff
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Print_UnitConversion", "modPrintListings.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub
