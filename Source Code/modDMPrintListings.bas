Attribute VB_Name = "modDMPrintListings"
'--------------------------------------------------------------------------------------------------
'   File:       modPrintListings.bas
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Ashitei Trebi-Ollennu, September 2001
'   Purpose:    Routines for carrying out non Crystal Reports printing as
'               required by document Re-Write of Ex-Crystal Reports Listings
'--------------------------------------------------------------------------------------------------
'Revisions: 23/11/2001 ASH Added error trapping for Buglist no.14
'           1/02/2002  ASH: Added sql to calculate Actual Recruitment
'           ASH 13/05/2002 Minor modifications to Visit Status
'   TA  1/7/2005    set font.charset to allow non western european characters
'--------------------------------------------------------------------------------------------------
Option Explicit
'--------------------------------------------------------------------------------------------------
Public Sub PrintSiteRecruitment()
'--------------------------------------------------------------------------------------------------
'to print site recruitment records
'--------------------------------------------------------------------------------------------------
Dim rsSiteRec As New ADODB.Recordset
Dim rsSite As New ADODB.Recordset
Dim rsTotal As New ADODB.Recordset
Dim nCount As Integer
Dim sType As String
Dim sSQL As String
Dim sSiteRecruitmentSQL As String
Dim sStudyRecruitmentSQL As String
Dim msPrintName As String
Dim mnPrintBlock As Integer
Dim mlPrintingWidth As Long
Dim mnCurrentY As Integer
Dim nTotalSubjects As Integer
    
    On Error GoTo ErrHandler
    
    HourglassOn

    msPrintName = "Site Recruitment Report"
    mnPrintBlock = 543
    sSQL = ""
    
    sSQL = "SELECT ClinicalTrial.ClinicalTrialName,TrialSite.TrialSite,"
    sSQL = sSQL & " ClinicalTrial.ExpectedRecruitment,ClinicalTrial.ClinicalTrialID"
    sSQL = sSQL & " FROM ClinicalTrial,TrialSite"
    sSQL = sSQL & " WHERE TrialSite.ClinicalTrialID=ClinicalTrial.ClinicalTrialID"
    sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialID > 0"
    sSQL = sSQL & " ORDER BY ClinicalTrial.ClinicalTrialID"
    
    
    Set rsSiteRec = New ADODB.Recordset
    rsSiteRec.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    nCount = rsSiteRec.RecordCount
    
    If nCount < 1 Then
        Call DialogWarning("No site recruitment records to print", "Print Listing")
        HourglassOff
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
    
    rsSiteRec.MoveFirst
    Do While Not rsSiteRec.EOF
        
        sStudyRecruitmentSQL = "SELECT COUNT(*)"
        sStudyRecruitmentSQL = sStudyRecruitmentSQL & " FROM TrialSubject"
        sStudyRecruitmentSQL = sStudyRecruitmentSQL & " WHERE TrialSubject.ClinicalTrialID=" & rsSiteRec.Fields(3).Value
        
        Set rsTotal = New ADODB.Recordset
        rsTotal.Open sStudyRecruitmentSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        
        nTotalSubjects = rsTotal.Fields(0).Value
        
        sSiteRecruitmentSQL = "SELECT COUNT(*)"
        sSiteRecruitmentSQL = sSiteRecruitmentSQL & " FROM TrialSubject"
        sSiteRecruitmentSQL = sSiteRecruitmentSQL & " WHERE TrialSubject.TrialSite='" & rsSiteRec.Fields(1).Value & "'"
        sSiteRecruitmentSQL = sSiteRecruitmentSQL & " AND TrialSubject.ClinicalTrialID=" & rsSiteRec.Fields(3).Value
        
        Set rsSite = New ADODB.Recordset
        rsSite.Open sSiteRecruitmentSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        
        'continue printing on a new page
        If IsPageSizeEnough(mnPrintBlock) = False Then
            Printer.NewPage
            Call PrintHeader(mlPrintingWidth, msPrintName)
        End If
        
        'printer settings
        Printer.FontName = "Tahoma"
        Printer.Font.Charset = 1
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
        Printer.Print rsSiteRec.Fields(0).Value
        
        'site Name
        Printer.FontBold = True
        Printer.CurrentY = Printer.CurrentY + 60
        Printer.Print "Site: ", ;
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.CurrentX = 2500
        Printer.CurrentY = Printer.CurrentY + 10
        Printer.Print rsSiteRec.Fields(1).Value
        
        'Site recruitment
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Site Recruitment: ", ;
        Printer.FontBold = False
        Printer.CurrentX = 2500
        Printer.FontSize = 8
        Printer.Print rsSite.Fields(0).Value
        
        'Expected Recruitment
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Expected Recruitment: ", ;
        Printer.FontBold = False
        Printer.CurrentX = 2500
        Printer.FontSize = 8
        Printer.Print rsSiteRec.Fields(2).Value
        
        'Actual Recruitment
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Actual Recruitment: ", ;
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.CurrentX = 2500
        Printer.Print nTotalSubjects
        
        'draw line
        Printer.DrawWidth = 2
        Printer.CurrentY = Printer.CurrentY + 120
        mnCurrentY = Printer.CurrentY
        Printer.Line (0, mnCurrentY)-(mlPrintingWidth, mnCurrentY)
        
        rsSiteRec.MoveNext
    Loop
    
    Set rsSiteRec = Nothing
    Printer.EndDoc
    HourglassOff

Exit Sub
ErrHandler:
    
    If Err.Number = 482 Then
        DialogError ("Printer error.Check and retry"), "Printer Error"
        Call HourglassOff
        Exit Sub
    End If
  
  Call HourglassOff
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintSiteRecruitment", "modPrintListings.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub

'--------------------------------------------------------------------------------------------------
Public Sub PrintStudyRecruitment()
'--------------------------------------------------------------------------------------------------
'prints study recruitment reports
'--------------------------------------------------------------------------------------------------
Dim rsStudyRec As New ADODB.Recordset
Dim rsSubjects As New ADODB.Recordset
Dim rsTotal As New ADODB.Recordset
Dim nCount As Integer
Dim sType As String
Dim sSQL As String
Dim sStudyRecruitmentSQL As String
Dim msPrintName As String
Dim mnPrintBlock As Integer
Dim mlPrintingWidth As Long
Dim mnCurrentY As Integer
Dim nTotalSubjects As Integer
    
    On Error Resume Next
    
    HourglassOn
    
    msPrintName = "Study Recruitment Report"
    mnPrintBlock = 543
    sSQL = ""
   
    sSQL = "SELECT ClinicalTrialName,ClinicalTrialID,StatusID,ExpectedRecruitment"
    sSQL = sSQL & " From ClinicalTrial"
    sSQL = sSQL & " WHERE ClinicalTrialID > 0"
    
    Set rsStudyRec = New ADODB.Recordset
    rsStudyRec.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    nCount = rsStudyRec.RecordCount
    
    If nCount < 1 Then
        Call DialogWarning("No study recruitment records to print", "Print Listing")
        Exit Sub
    End If
    
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
    
    rsStudyRec.MoveFirst
    Do While Not rsStudyRec.EOF
         
        sStudyRecruitmentSQL = "SELECT COUNT(*)"
        sStudyRecruitmentSQL = sStudyRecruitmentSQL & " FROM TrialSubject"
        sStudyRecruitmentSQL = sStudyRecruitmentSQL & " WHERE TrialSubject.ClinicalTrialID=" & rsStudyRec.Fields(1).Value
        
        Set rsTotal = New ADODB.Recordset
        rsTotal.Open sStudyRecruitmentSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        
        nTotalSubjects = rsTotal.Fields(0).Value
         
         'continue printing on a new page
        If IsPageSizeEnough(mnPrintBlock) = False Then
            Printer.NewPage
            Call PrintHeader(mlPrintingWidth, msPrintName)
        End If
        
        'printer settings
        Printer.FontName = "Tahoma"
        Printer.Font.Charset = 1
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
        Printer.Print rsStudyRec.Fields(0).Value
        
        'Study ID
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Status: ", ;
        Printer.FontBold = False
        Printer.CurrentX = 2500
        Printer.FontSize = 8
        sType = DecodeStatusID(rsStudyRec.Fields(2).Value)
        Printer.Print sType
        
        'Expected Recruitment
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Expected Recruitment: ", ;
        Printer.FontBold = False
        Printer.CurrentX = 2500
        Printer.FontSize = 8
        Printer.Print rsStudyRec.Fields(2).Value
        
        'Actual Recruitment
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.FontBold = True
        Printer.Print "Actual Recruitment: ", ;
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.CurrentX = 2500
        Printer.Print nTotalSubjects
        
        'draw line
        Printer.DrawWidth = 2
        Printer.CurrentY = Printer.CurrentY + 120
        mnCurrentY = Printer.CurrentY
        Printer.Line (0, mnCurrentY)-(mlPrintingWidth, mnCurrentY)
        
        rsStudyRec.MoveNext
    
    Loop
    
        Set rsStudyRec = Nothing
        Printer.EndDoc
        HourglassOff

Exit Sub
ErrHandler:
    
    If Err.Number = 482 Then
        DialogError ("Printer error.Check and retry"), "Printer Error"
        Call HourglassOff
        Exit Sub
    End If
  
  Call HourglassOff
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintStudyRecruitment", "modPrintListings.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub

'--------------------------------------------------------------------------------------------------
Public Sub PrintVisitStatus()
'--------------------------------------------------------------------------------------------------
'prints visit status reports
'--------------------------------------------------------------------------------------------------
Dim rsVisitStatus As New ADODB.Recordset
Dim nCount As Integer
Dim sType As String
Dim msSQL As String
Dim msSQL1 As String
Dim msPrintName As String
Dim mnPrintBlock As Integer
Dim mlPrintingWidth As Long
Dim mnCurrentY As Integer
    
    On Error Resume Next
    
    HourglassOn
    
    msPrintName = "Visit Status Report"
    mnPrintBlock = 1400
    msSQL = ""
    
    msSQL = "SELECT  VisitInstance.TrialSite,StudyVisit.VisitName,"
    msSQL = msSQL & " VisitInstance.PersonID,VisitInstance.VisitStatus,"
    msSQL = msSQL & " ClinicalTrial.ClinicalTrialName"
    msSQL = msSQL & " FROM VisitInstance,StudyVisit,ClinicalTrial"
    msSQL = msSQL & " WHERE VisitInstance.VisitID=StudyVisit.VisitID"
    msSQL = msSQL & " AND VisitInstance.ClinicalTrialID=StudyVisit.ClinicalTrialID"
    msSQL = msSQL & " AND ClinicalTrial.ClinicalTrialID=VisitInstance.ClinicalTrialID"
    msSQL = msSQL & " AND ClinicalTrial.ClinicalTrialID=StudyVisit.ClinicalTrialID"
    msSQL = msSQL & " ORDER BY VisitInstance.TrialSite"
    'msSQL = msSQL & " ORDER BY StudyVisit.VisitID"
    
    Set rsVisitStatus = New ADODB.Recordset
    rsVisitStatus.Open msSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    nCount = rsVisitStatus.RecordCount
    
    If nCount < 1 Then
        Call DialogWarning("No visit status records to print", "Print Listing")
    Exit Sub
    End If
    
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
    Call PrintStudyName(rsVisitStatus.Fields(4).Value)
    
    'print  header for report details
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 120
    Printer.FontSize = 8
    Printer.FontBold = True
    Printer.Print "Site", ;
    Printer.CurrentX = 2000
    Printer.Print "Visit", ;
    Printer.CurrentX = 7500
    Printer.Print "Subject ID", ;
    Printer.CurrentX = 9000
    Printer.Print "Status", ;
    Printer.FontBold = False
    Printer.CurrentY = Printer.CurrentY + 240
    
        Do While Not rsVisitStatus.EOF
            'ash 13/05/2002
            'continue printing on a new page
            If IsPageSizeEnough(mnPrintBlock) = False Then
                Printer.NewPage
                Call PrintHeader(mlPrintingWidth, msPrintName)
            End If
            'printer settings
            Printer.FontName = "Tahoma"
            Printer.Font.Charset = 1
            Printer.FontSize = 8
            'print details
            Printer.CurrentX = 0
            Printer.Print rsVisitStatus.Fields(0).Value, ;
            Printer.CurrentX = 2000
            Printer.Print rsVisitStatus.Fields(1).Value, ;
            Printer.CurrentX = 7500
            Printer.Print rsVisitStatus.Fields(2).Value, ;
            Printer.CurrentX = 9000
            sType = GetStatusString(rsVisitStatus.Fields(3).Value)
            Printer.Print sType
            rsVisitStatus.MoveNext
        
        Loop
        'draw line
        Printer.DrawWidth = 2
        Printer.CurrentY = Printer.CurrentY + 120
        mnCurrentY = Printer.CurrentY
        Printer.Line (0, mnCurrentY)-(mlPrintingWidth, mnCurrentY)
        
        Set rsVisitStatus = Nothing
        Printer.EndDoc
        HourglassOff

Exit Sub
ErrHandler:
    
    If Err.Number = 482 Then
        DialogError ("Printer error.Check and retry"), "Printer Error"
        Call HourglassOff
        Exit Sub
    End If
  
  Call HourglassOff
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintVisitStatus", "modPrintListings.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub
'--------------------------------------------------------------------------------------------------
Public Function IsPageSizeEnough(mnPrintBlock As Integer, _
                                Optional nExtraWidth1 As Integer, _
                                Optional nExtraWidth2 As Integer) As Boolean
'--------------------------------------------------------------------------------------------------
'checks to see if enough space is on the page to print required text. If not
'opens a new page and prints header followed by the required text
'--------------------------------------------------------------------------------------------------
Dim nTotalPageLength As Integer
Dim nRequiredPrintLines As Integer

    On Error GoTo ErrHandler
    nTotalPageLength = 13680   '9.5 * 1440 (page length * twips)
    nRequiredPrintLines = (Printer.CurrentY + mnPrintBlock) + (190 * (nExtraWidth1 + nExtraWidth2))
    
    If nRequiredPrintLines > nTotalPageLength Then
        IsPageSizeEnough = False
    Else
        IsPageSizeEnough = True
    End If

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "IsPageSizeEnough", "modDMPrintListings.bas")
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
Dim mnCurrentY As Integer

    On Error GoTo ErrHandler

    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.FontName = "Arial Narrow"
    Printer.Font.Charset = 1
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
Public Sub PrintStudyName(sStudyName As String)
'--------------------------------------------------------------------------------------------------
'prints study name
'--------------------------------------------------------------------------------------------------
Dim lErrNum As Integer
Dim sErrDesc As String
Dim msSQL As String
Dim msSQL1 As String
Dim msPrintName As String
Dim mnPrintBlock As Integer
Dim mlPrintingWidth As Long
Dim mnCurrentY As Integer

On Error GoTo ErrHandler
    
    Printer.FontName = "Arial"
    Printer.Font.Charset = 1
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
            sType = "Closed to recruitment"
        Case Is = 4
            sType = "Closed to follow up"
        Case Is = 5
            sType = "Suspended"
        Case Is = 6
            sType = "Deleted"
    End Select
        
        DecodeStatusID = sType

End Function
