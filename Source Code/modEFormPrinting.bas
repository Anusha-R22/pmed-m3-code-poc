Attribute VB_Name = "modEFormPrinting"
'----------------------------------------------------------------------------------------'
'   Copyright:  Inferfrmformd Ltd. 2001. All Rights Reserved
'   File:       modEFormPrinting.bas
'   Author:     Mo Morris, October 2001
'   Purpose:    The re-written version of eForm printing.
'               At the time of writing this can only be run under MACRO_DM, because it
'               uses the new data services objects and classes, which onle exist in MACRO_DM
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   TA 05/10/20001: Removed refererences to goStudyDef and removed module level
'                       EFI to avoid a reference being permanently maintained.
'   Mo 19/3/2003    Changes around printing RQGs attachments, hotlinks and caption fonts.
'                   New subs PrintRQG, PrintRQGEstimate, PrintAttachment, PrinthotLink
'                   and GetCaptionFont, together with major changes to PrintTextBox and PrinteForm
'   Mo 6/5/2003     PrintComment changed so that it now uses the Caption Font
' NCJ 1 Jul 03 - Use new Expandable property of Element in PrintTextBox and PrintRQGEstimate (Bug 1905)
'   Mo 14/4/2004    Bug 2063. CR/LFs in repeating Question Group Question/Header captions now
'                   handled correctly. Changes to PrintRQG + new function NumLinesInCaption.
'                   RQG Borders are no longer printed when not required.
'   Mo  20/6/2005   Bug 2475. Sub eFormPrinting changed so that eForms are now printed
'                   in schedule order (used to print in alphabetical order)
'   DPH 27/06/2005  Bug 2554 Only print active Option boxes / buttons
'   TA  1/7/2005    set font.charset to allow non western european characters
'   ic 30/08/2005   added clinical coding
'   Mo  26/10/2007  Bug 2957, minor change to PrintEForm when page end corresponds to a hotlink.
'----------------------------------------------------------------------------------------'

Option Explicit

'These variables hold the default font attributes for the eForm being printed
Private mlFontColour As Long
Private mlEFormColour As Long
Private msFontName As String
Private mbFontBold As Boolean
Private mbFontItalic As Boolean
Private mnFontSize As Integer
Private Const ms_DEFAULT_FontName As String = "Arial"
Private Const mn_DEFAULT_FontSize As Integer = 10

Private mbDisplayNumbers As Boolean

Private mnPageCount As Integer
Private msgTitleGap As Single

'note that 8515 twips is the maximum width that you get on an 800x600 display
'VTRACK eForm width is larger
#If VTRACK = 1 Then
Private Const mnEFORMWIDTHINTWIPS = 14500
#Else
Private Const mnEFORMWIDTHINTWIPS = 8515
#End If

Private msgScaleFactor As Single
Private msgScaleFactorFonts As Single

Private Const mn_VISUAL_ELEMENT = 16384

'Mo 21/2/2003, Print eForm Changes
Private mneFormWidth As Integer

Private Const mn_SPACE_FOR_STATUS_ICON = 495

'---------------------------------------------------------------------
Public Sub EFormPrinting(oSubject As StudySubject, ByVal bExcludeBlanks As Boolean)
'---------------------------------------------------------------------
'This sub controls the printing of all forms within a study.
'It can be called in two modes:-
'   All Forms Including Blank Forms
'   All Forms Excluding Blank Forms
' Subject to be printed passed in
'---------------------------------------------------------------------
Dim oScheduleVisit As ScheduleVisit
Dim oVEFInstance As VEFInstance
Dim oEFI As EFormInstance
Dim bErrorsDuringPrinting As Boolean
Dim nSheetNumber As Integer
Dim sLockErrMsg As String
Dim sEFILockToken As String
Dim sVEFILockToken As String

    On Error GoTo ErrLabel
    
    Call FontInit(oSubject.StudyDef)
    
    bErrorsDuringPrinting = False
    
    'Turn Sheet numbering on
    nSheetNumber = 1

    'Mo 20/6/2005 Bug 2475
    'loops through visitEFormInstances on ScheduleVisits, instead of VisitInstances
    For Each oScheduleVisit In oSubject.ScheduleVisits
        For Each oVEFInstance In oScheduleVisit.VisitEFormInstances
            'Check for previous printer errors
            If bErrorsDuringPrinting Then
                Exit For
            End If
            Set oEFI = oVEFInstance.EFormInstance
            'Check for previous printer errors
            If bErrorsDuringPrinting Then
                Exit For
            End If
            'If in exclude blank forms mode check the form status is not Requested
            If Not oEFI Is Nothing Then
                If (Not bExcludeBlanks) Or (bExcludeBlanks And oEFI.Status <> eStatus.Requested) Then
                    'TA 15/09/2002: lock error message is returned in sLockErrMsg
                    ' - we can ignore it here becasue we only want readonly
                    'we don't need to hold onto the EFILockToken or Visit EFILockToken
                    If oSubject.LoadResponses(oEFI, sLockErrMsg, sEFILockToken, sVEFILockToken) = lrrCouldNotLockForSave Then
                        DialogError "Unable to print eForm " & oEFI.Code & "." & vbCrLf & sLockErrMsg
                        bErrorsDuringPrinting = True
                    ElseIf Not PrintEForm(oEFI, nSheetNumber) Then
                        bErrorsDuringPrinting = True
                    End If
                    Call oSubject.RemoveResponses(oEFI, True)
                End If
            End If
        Next oVEFInstance
    Next oScheduleVisit
          
Exit Sub
ErrLabel:
    If MACROErrorHandler("modEFormPrinting", Err.Number, Err.Description, "EFormPrinting", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Public Function PrintEForm(oEFI As EFormInstance, _
                            Optional ByRef nSheetNumber As Integer = 0) As Boolean
'---------------------------------------------------------------------
'Note that this code originates from BuildCRF.PrintCRFForm
'---------------------------------------------------------------------
Dim sgFormLength As Single
Dim sgYPrint As Single
Dim sgCaptionYPrint As Single
Dim sgAccumlativeReduction As Single
Dim sgElementLength As Single
Dim sgCaptionLength As Single
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant
Dim sResponseValue As String
Dim sgAdditionalDisplayLength As Single
Dim sgPrevAdditionalDisplayLength As Single
Dim sgPrevYPrint As Single
Dim sMsg As String
Dim sCaption As String
Dim oElement As eFormElementRO
Dim oResponse As Response
Dim sCommentText As String
Dim nPositionOfCR As Integer
Dim nLineCount As Integer
Dim nNumberOfLines As Integer
Dim lPrintOrderY As Long
Dim oEFQGroup As EFormQGroupDE
    
    On Error GoTo ErrLabel
    
    Call FontInit(oEFI.eForm.Study)
    
    'Initialise page count
    mnPageCount = 1
    
    On Error Resume Next
    frmMenu.CommonDialog1.CancelError = True
    'Note that there is no CommonDialog1.ShowPrinter statement.
'    frmMenu.CommonDialog1.Orientation = cdlLandscape
'    Printer.TrackDefault = True
'    frmMenu.CommonDialog1.ShowPrinter
    'If there was it would get called for each form being printed.
    If Err.Number > 0 Then Exit Function
    'restore normal error trapping
    On Error GoTo ErrLabel
    
'    Printer.Orientation = frmMenu.CommonDialog1.Orientation
    
    'set printer scalemode to twips
    Printer.ScaleMode = vbTwips
    
    'calculate the scaling factor between the width of the form and the width of the
    'printed page (minus the margins - 1 inch on the left and 1/4 inch on the right).
    'This same scaling factor will be used to adjust the scaleHeight proportionally
    'Note that the form width is no longer fixed and read from oEFI.eForm.eFormWidth
    'Mo 21/2/2003, Print eForm Changes
    If oEFI.eForm.eFormWidth = NULL_LONG Then
        mneFormWidth = glPORTRAIT_WIDTH
    Else
        mneFormWidth = oEFI.eForm.eFormWidth
    End If
    msgScaleFactor = (mneFormWidth / (Printer.ScaleWidth - 1440 - 360))
    'msgScaleFactor = (mnEFORMWIDTHINTWIPS / (Printer.ScaleWidth - 1440 - 360))
    
    'set scaleleft to incorporate a 1 inch margin (1440 twips)
    'set scaletop to incorporate a 1/4 inch margin (360 twips)
    Printer.ScaleLeft = -1440 * msgScaleFactor
    Printer.ScaleTop = -360 * msgScaleFactor
    Printer.ScaleWidth = Printer.ScaleWidth * msgScaleFactor
    Printer.ScaleHeight = Printer.ScaleHeight * msgScaleFactor
    msgTitleGap = -300 * msgScaleFactor
    
    'Print the forms header
    Call PrintEFormHeader(oEFI, nSheetNumber)
    
    'Based on the form size and the selected paper size sgFormLength represents the maximum
    'Y value that can appear on the current page. All controls will be tested aginst this value
    'prior to being printed. When it is exceeded a new page will be created
    sgFormLength = Printer.ScaleHeight - (720 * msgScaleFactor)
    
    'Calculate the reciprocal of the scaling factor for use on font sizes
    msgScaleFactorFonts = 1 / msgScaleFactor
    
    'Initialise mbDisplayNumbers which controls the display of numbers in a question's caption
    mbDisplayNumbers = oEFI.eForm.DisplayNumbers
    
    'Initialise the accumlative page correction variable
    sgAccumlativeReduction = 0
    sgAdditionalDisplayLength = 0
    sgPrevAdditionalDisplayLength = 0
    
    'Read through the elements on the form. Each element pertains to a control and a caption
    '(if it is of the type that requires a caption). When printing the forms it is neccessary
    'to have a Y coordinate for each control that is the lesser of Y and CaptionY. The following
    'code assesses the lesser value (PrintOrder), and sorts the elements within an unattached recordset.
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    
    rsTemp.Fields.Append "ElementId", adInteger
    rsTemp.Fields.Append "PrintOrderY", adInteger
    rsTemp.Open
    For Each oElement In oEFI.eForm.EFormElements
        'Ignore questions that are elements of a Repeating Question Group
        If oElement.OwnerQGroup Is Nothing Then
            If oElement.CaptionY = 0 Or oElement.ControlType > mn_VISUAL_ELEMENT Then
                lPrintOrderY = oElement.ElementY
            Else
                'Mo 29/4/2004, Bug 2259, check that the caption is not empty
                If IsNull(oElement.Caption) Then
                    sCaption = ""
                Else
                    sCaption = Trim(oElement.Caption)
                End If
                If Len(sCaption) = 0 Then
                    'its an empty caption, so ignore CaptionY and use ElementY
                    lPrintOrderY = oElement.ElementY
                Else
                    If oElement.CaptionY < oElement.ElementY Then
                        lPrintOrderY = oElement.CaptionY
                    Else
                        lPrintOrderY = oElement.ElementY
                    End If
                End If
            End If
            rsTemp.AddNew
            rsTemp.Fields(0).Value = oElement.ElementID
            rsTemp.Fields(1).Value = lPrintOrderY
            rsTemp.Update
        End If
    Next oElement
    
    'Sort the unnattacheed recordet on PrintOrderY
    rsTemp.Sort = "PrintOrderY"
    
    rsTemp.MoveFirst
   
    'Loop through the sorted recordset
    While Not rsTemp.EOF
    'Get the next element
    Set oElement = oEFI.eForm.eFormElementById(rsTemp!ElementID)
        If oElement.Hidden = False Then
            'Debug.Print oElement.Caption & "  type=" & oElement.ControlType
            sCaption = oElement.Caption
            'Make copies of the elements 'Y' and 'CaptionY' values, because they might need changing
            sgPrevYPrint = sgYPrint
            sgYPrint = oElement.ElementY
            sgCaptionYPrint = oElement.CaptionY
            'reduce 'Y' and 'CaptionY' if we are no longer dealing with the first page
            If mnPageCount > 1 Then
                sgYPrint = sgYPrint - sgAccumlativeReduction
                sgCaptionYPrint = sgCaptionYPrint - sgAccumlativeReduction
            End If
            'Text boxes that contain long strings that need to wrap over onto extra lines are
            'detected in PrintTextBox and cause an increase to the variable sgAdditionalDisplayLength
            'Comments (i.e. ControlType=gn_COMMENT) that are very close to a control that might have
            'generated an increase to sgAdditionalDisplayLength should not have sgAdditionalDisplayLength
            'added to their Y-co-ordinates, the previous value of sgAdditionalDisplayLength
            '(i.e. sgPrevAdditionalDisplayLength) should be used.
            If oElement.ControlType = gn_COMMENT And (Abs(sgYPrint + sgPrevAdditionalDisplayLength - sgPrevYPrint) < 150) Then
                sgYPrint = sgYPrint + sgPrevAdditionalDisplayLength
            Else
                sgYPrint = sgYPrint + sgAdditionalDisplayLength
                sgCaptionYPrint = sgCaptionYPrint + sgAdditionalDisplayLength
            End If
            
            'assess length of current element to be printed
            Select Case oElement.ControlType
            Case 0
                'Its a Repeating Question Group
                Call PrintRQGEstimate(oEFI, oElement, sgElementLength, sgAdditionalDisplayLength)
                
            Case gn_OPTION_BUTTONS, gn_PUSH_BUTTONS
                'OptionButtons, OptionBoxes
                Call GetFont(oElement, vFontName, vFontSize, vFontBold, vFontItalic)
                On Error Resume Next
                Printer.FontName = vFontName
                Printer.Font.Charset = 1
                On Error GoTo ErrLabel
                Printer.FontSize = vFontSize * msgScaleFactorFonts
                'calculate length of a single item
                sgElementLength = (Printer.TextHeight("X") + 100) * 1.5
                'using the number of values belonging to dataitem calculate the total length of the current optionBox/optionButton
                sgElementLength = sgElementLength * oElement.Categories.Count
            Case gn_LINE
                'Line
                sgElementLength = 0
            Case gn_COMMENT
                'Comment
                sCommentText = sCaption
                Call GetFont(oElement, vFontName, vFontSize, vFontBold, vFontItalic)
                On Error Resume Next
                Printer.FontName = vFontName
                Printer.Font.Charset = 1
                On Error GoTo ErrLabel
                Printer.FontSize = vFontSize * msgScaleFactorFonts
                nLineCount = 1
                While sCommentText <> ""
                    nPositionOfCR = InStr(sCommentText, Chr(13))
                    If nPositionOfCR <> 0 Then
                        sCommentText = Mid(sCommentText, nPositionOfCR + 2, Len(sCommentText))
                        nLineCount = nLineCount + 1
                    Else
                        sCommentText = ""
                    End If
                Wend
                sgElementLength = Printer.TextHeight("X") * nLineCount
            Case gn_PICTURE
                'Picture
                On Error Resume Next
                frmMenu.picUsedForPrinting.Picture = LoadPicture(gsDOCUMENTS_PATH & oElement.Caption)
                On Error GoTo ErrLabel
                If Err.Number = 0 Then
                    sgElementLength = frmMenu.picUsedForPrinting.Picture.Height * msgScaleFactorFonts
                End If
            Case Else
                'TextBox or dropdown list or calendar
                Call GetFont(oElement, vFontName, vFontSize, vFontBold, vFontItalic)
                On Error Resume Next
                Printer.FontName = vFontName
                Printer.Font.Charset = 1
                On Error GoTo ErrLabel
                Printer.FontSize = vFontSize * msgScaleFactorFonts
                'Check that there is enough room to print the response
                Set oResponse = oEFI.Responses.ResponseByElement(oElement)
                If Not oResponse Is Nothing Then
                    sResponseValue = oResponse.Value
                Else
                    sResponseValue = ""
                End If
                'Mo 21/2/2003, Print eForm Changes
                If Printer.TextWidth(sResponseValue) > (mneFormWidth - oElement.ElementX) Then
                    'Estimate the number of lines required
                    nNumberOfLines = (Printer.TextWidth(sResponseValue)) \ (mneFormWidth - oElement.ElementX) + 1
                    sgElementLength = (nNumberOfLines * Printer.TextHeight("X")) + 100
                    sgPrevAdditionalDisplayLength = sgAdditionalDisplayLength
                    sgAdditionalDisplayLength = sgAdditionalDisplayLength + (nNumberOfLines - 1) * (Printer.TextHeight("X") + 100)
                Else
                    sgElementLength = Printer.TextHeight("X") + 100
                End If
            End Select
            
            'assess caption length
            Call GetCaptionFont(oElement, vFontName, vFontSize, vFontBold, vFontItalic)
            On Error Resume Next
            Printer.FontName = vFontName
            Printer.Font.Charset = 1
            On Error GoTo ErrLabel
            Printer.FontSize = vFontSize * msgScaleFactorFonts
            sgCaptionLength = Printer.TextHeight("X") + 100
            
            'check that the current control will fit onto the page
            'Test for a non blank caption before deciding that it should be part of
            'the decision to terminate the current page. This was done so that blank captions
            'with a CaptionY co-ordinate that is miles away from its control don't effect the printing
            If (sCaption > "" And ((sgYPrint + sgElementLength > sgFormLength) Or (sgCaptionYPrint + sgCaptionLength > sgFormLength))) _
                Or (sCaption = "" And (sgYPrint + sgElementLength > sgFormLength)) Then
                'print the page that has just become full
                Printer.EndDoc
                'Increase the accumlative page correction variable by the smaller of sgYPrint and
                'sgCaptionYPrint (if Caption is non-blank).
                'Test for a non-comment before allowing sgCaptionYPrint to be part of the calculation
                'of sgAccumlativeReduction. (i.e comments have a CaptionY of zero which can become a
                'large negative figure for pages 2 and onwards and would always be less than sgYPrint
                'Mo 26/10/2007 Bug 2957
                If sgCaptionYPrint < sgYPrint And sCaption > "" And (oElement.ControlType < mn_VISUAL_ELEMENT) Then
                    'reduce 'Y' and 'CaptionY' so that they will fit onto the new page
                    sgYPrint = sgYPrint - sgCaptionYPrint
                    sgAccumlativeReduction = sgAccumlativeReduction + sgCaptionYPrint
                    sgCaptionYPrint = 0
                Else
                    'reduce 'Y' and 'CaptionY' so that they will fit onto the new page
                    sgCaptionYPrint = sgCaptionYPrint - sgYPrint
                    sgAccumlativeReduction = sgAccumlativeReduction + sgYPrint
                    sgYPrint = 0
                End If
                
                'increment the page count
                mnPageCount = mnPageCount + 1
                
                'Print the forms header
                Call PrintEFormHeader(oEFI, nSheetNumber)
            End If
            'put the control onto the page
            Call PrintEFormElement(oEFI, oElement, sgYPrint, sgCaptionYPrint)
        End If      ' If element not hidden
        Set oElement = Nothing
        rsTemp.MoveNext
    Wend
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    'print the last page
    Printer.EndDoc
    
    PrintEForm = True
  
Exit Function
ErrLabel:
    MsgBox "An error has occurred during printing. The error number is " & Err.Number & vbCrLf _
         & Err.Description, vbOKOnly + vbInformation
    PrintEForm = False
End Function

'---------------------------------------------------------------------
Private Sub PrintEFormHeader(oEFI As EFormInstance, _
                            Optional ByRef nSheetNumber As Integer = 0)
'---------------------------------------------------------------------
'format and print a title in the top corner of printed page and draw
'the forms corners/margins
'---------------------------------------------------------------------
Dim sTitle As String
Dim oStudyDef As StudyDefRO

    On Error GoTo ErrLabel

    Set oStudyDef = oEFI.eForm.Study
    
    Printer.CurrentX = 0
    Printer.CurrentY = msgTitleGap
    'Mo 21/2/2003, Print eForm Changes, VisitCycleNo, FormCycleNo and FormLabel added
    sTitle = "Study: " & oStudyDef.Name & " (Version: " & oStudyDef.Version & ")" _
        & "    Visit: " & oEFI.VisitInstance.Visit.Code & "[" & oEFI.VisitInstance.CycleNo & "]" _
        & "    Form: " & oEFI.Name & "[" & oEFI.CycleNo & "]" _
        & "    FormLabel: " & oEFI.eFormLabel _
        & "    Page: " & mnPageCount _
        & vbTab & vbTab & "Printed " & Format(Now, "yyyy/mm/dd hh:mm:ss")
    Printer.FontSize = 6
    Printer.FontBold = False
    Printer.FontItalic = False
    Printer.Print sTitle
    'Sheet number no longer appears on individually printed forms
    'nSheetNumber=0 means individual form printing, DO NOT PRINT SHEET NUMBERS
    'nSheetNumber>0 means print all forms printing, PRINT SHEET NUMBERS
    sTitle = "[Site: " & oStudyDef.Subject.Site & "    SubjectId: " & oStudyDef.Subject.PersonId _
        & "    SubjectLabel: " & oStudyDef.Subject.label
    If nSheetNumber > 0 Then
        sTitle = sTitle & "   Sheet:" & nSheetNumber & "]"
        nSheetNumber = nSheetNumber + 1
    Else
        sTitle = sTitle & "]"
    End If
    Printer.CurrentY = msgTitleGap / 2
    Printer.Print sTitle
    
    'draw border corners on the page, incorporating a 1/2 inch margin at the bottom (720 twips)
    Printer.DrawWidth = 1
    Printer.Line (0, 0)-Step(400, 0)
    'Mo 21/2/2003, Print eForm Changes
    Printer.Line (mneFormWidth - 400, 0)-Step(400, 0)
    Printer.Line -Step(0, 400)
    Printer.Line (mneFormWidth, (Printer.ScaleHeight - (720 * msgScaleFactor)) - 400)-Step(0, 400)
    Printer.Line -Step(-400, 0)
    Printer.Line (400, (Printer.ScaleHeight - (720 * msgScaleFactor)))-Step(-400, 0)
    Printer.Line (0, (Printer.ScaleHeight - (720 * msgScaleFactor)))-Step(0, -400)
    Printer.Line (0, 0)-Step(0, 400)

    Set oStudyDef = Nothing
    
Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.PrintEFormHeader"
End Sub

'---------------------------------------------------------------------
Private Sub PrintCaption(oElement As eFormElementRO, _
                        ByVal sgYPrint As Single, _
                        ByVal sgCaptionYPrint As Single)
'---------------------------------------------------------------------
' Routine now handles multiple line captions inline with multiple line comments
'---------------------------------------------------------------------
Dim sUnitOfMeasurement As String
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant
Dim sCaption As String
Dim sCaptionLine As String
Dim nLineCount As Integer
Dim nPositionOfCR As Integer
Dim sglTextHeight As Single

    On Error GoTo ErrLabel

    '  Prepare unit of measurement part of label
    sUnitOfMeasurement = oElement.Unit
    If sUnitOfMeasurement > "" Then
        ' Add brackets round unit
        sUnitOfMeasurement = " (" & sUnitOfMeasurement & ")"
    End If
    
    'Build caption text
    'Check for caption numbers being turned off
    If mbDisplayNumbers Then
        sCaption = CStr(oElement.ElementOrder) & ". " & oElement.Caption & sUnitOfMeasurement
    Else
        sCaption = oElement.Caption & sUnitOfMeasurement
    End If
    
    'Look for specific font or use the form's default one
    Call GetCaptionFont(oElement, vFontName, vFontSize, vFontBold, vFontItalic)
    
    'Set font attributes
    On Error Resume Next
    Printer.FontName = vFontName
    Printer.Font.Charset = 1
    On Error GoTo ErrLabel
    
    Printer.FontSize = vFontSize * msgScaleFactorFonts
    Printer.FontBold = vFontBold
    Printer.FontItalic = vFontItalic
    
    'set page position for Caption
    'needs amending to cope with margins and negative values
    
    nLineCount = 0
    ' Get text height
    sglTextHeight = Printer.TextHeight("X")
    
    While sCaption <> ""
        nPositionOfCR = InStr(sCaption, Chr(13))
        If nPositionOfCR <> 0 Then
            'strip off text before Cr/LF
            sCaptionLine = Mid(sCaption, 1, nPositionOfCR - 1)
            sCaption = Mid(sCaption, nPositionOfCR + 2, Len(sCaption))
        Else
            sCaptionLine = sCaption
            sCaption = ""
        End If
        'set X co-ordinate position for caption line
        If oElement.CaptionX > 0 Then
            Printer.CurrentX = oElement.CaptionX
        Else
            Printer.CurrentX = oElement.ElementX - Printer.TextWidth(sCaptionLine) - 50
        End If
        'set Y co-ordinate position for caption line
        If oElement.CaptionY > 0 Then
            Printer.CurrentY = sgCaptionYPrint + (sglTextHeight * nLineCount)
        Else
            Printer.CurrentY = sgYPrint + (sglTextHeight * nLineCount)
        End If
        nLineCount = nLineCount + 1
        'print the caption line
        Printer.Print sCaptionLine
    Wend

Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.PrintCaption"
End Sub

'---------------------------------------------------------------------
Private Sub PrintComment(oElement As eFormElementRO, _
                        ByVal sgYPrint As Single)
'---------------------------------------------------------------------
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant
Dim sCommentText As String
Dim sCommentline As String
Dim nLineCount As Integer
Dim nPositionOfCR As Integer

    On Error GoTo ErrLabel
    
    '  Look for specific font or use the form's default one
    'Changed Mo 6/5/2003, call to GetFont changed to GetCaptionFont
    Call GetCaptionFont(oElement, vFontName, vFontSize, vFontBold, vFontItalic)
    
    '  Set font attributes
    On Error Resume Next
    Printer.FontName = vFontName
    Printer.Font.Charset = 1
    On Error GoTo ErrLabel
    
    Printer.FontSize = vFontSize * msgScaleFactorFonts
    Printer.FontBold = vFontBold
    Printer.FontItalic = vFontItalic
    
    sCommentText = oElement.Caption
    nLineCount = 0
    
    While sCommentText <> ""
        nPositionOfCR = InStr(sCommentText, Chr(13))
        If nPositionOfCR <> 0 Then
            'strip off text before Cr/LF
            sCommentline = Mid(sCommentText, 1, nPositionOfCR - 1)
            sCommentText = Mid(sCommentText, nPositionOfCR + 2, Len(sCommentText))
        Else
            sCommentline = sCommentText
            sCommentText = ""
        End If
        'set page position for comment line
        Printer.CurrentX = oElement.ElementX
        Printer.CurrentY = sgYPrint + (Printer.TextHeight("X") * nLineCount)
        nLineCount = nLineCount + 1
        'print the comment line
        Printer.Print sCommentline
    Wend

Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.PrintComment"
End Sub

'---------------------------------------------------------------------
Private Sub PrintEFormElement(oEFI As EFormInstance, oElement As eFormElementRO, _
                           ByVal sgYPrint As Single, _
                           ByVal sgCaptionYPrint As Single)
'---------------------------------------------------------------------
'Print a CRFElement
'---------------------------------------------------------------------
Dim nControlType As Integer

    On Error GoTo ErrLabel
    
    nControlType = oElement.ControlType
    
    If nControlType < mn_VISUAL_ELEMENT Then
        If oElement.Caption > "" Then
            Call PrintCaption(oElement, sgYPrint, sgCaptionYPrint)
        End If
    End If
        
    'draw the relevant element
    Select Case nControlType
    Case 0
        Call PrintRQG(oEFI, oElement, sgYPrint)
    Case gn_OPTION_BUTTONS
        Call PrintOptionButtons(oEFI, oElement, sgYPrint)
    Case gn_PUSH_BUTTONS
        Call PrintOptionBoxes(oEFI, oElement, sgYPrint)
    Case gn_LINE
        Call PrintLine(oElement, sgYPrint)
    Case gn_COMMENT
        Call PrintComment(oElement, sgYPrint)
    Case gn_PICTURE
        Call PrintPicture(oElement, sgYPrint)
    Case gn_HOTLINK
        Call PrintHotLink(oElement, sgYPrint)
    Case gn_ATTACHMENT
        Call PrintAttachment(oElement, sgYPrint)
    Case Else
        'TextBox, PopUpList, Calendar
        Call PrintTextBox(oEFI, oElement, sgYPrint)
    End Select
    
Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.PrintEFormElement"
End Sub

'---------------------------------------------------------------------
Private Sub PrintLine(oElement As eFormElementRO, _
                    ByVal sgYPrint As Single)
'---------------------------------------------------------------------
    On Error GoTo ErrLabel
    
    Printer.CurrentY = sgYPrint
    Printer.CurrentX = 0
    
    'set draw width to 3 pixels for lines and draw a horizontal line
    Printer.DrawWidth = 3
    'Mo 21/2/2003, Print eForm Changes
    Printer.Line -Step(mneFormWidth, 0)
    Printer.DrawWidth = 2

Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.PrintLine"
End Sub

'---------------------------------------------------------------------
Private Sub PrintOptionBoxes(oEFI As EFormInstance, _
                            oElement As eFormElementRO, _
                            ByVal sgYPrint As Single, _
                            Optional ByVal sgXPrint As Single = 0, _
                            Optional ByVal nRepeat As Integer = 1, _
                            Optional ByRef sgControlWidth As Single, _
                            Optional ByRef sgControlHeight As Single)
'---------------------------------------------------------------------
' REVISIONS
' DPH 27/06/2005 - Bug 2554 Only print active Option boxes
'---------------------------------------------------------------------
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant
Dim sgOptionBoxWidth As Single
Dim sgOptionBoxHeight As Single
Dim nOptionBoxCount As Integer
Dim sResponseValue As String
Dim oResponse As Response
Dim oCategory As CategoryItem
Dim sgElementX As Single
Dim sgSpaceForIcons As Single

    On Error GoTo ErrLabel
    
    If sgXPrint <> 0 Then
        sgElementX = sgXPrint
    Else
        sgElementX = oElement.ElementX
    End If

    'Look for specific font or use the form's default one
    Call GetFont(oElement, vFontName, vFontSize, vFontBold, vFontItalic)
    
    'Set font attributes
    On Error Resume Next
    Printer.FontName = vFontName
    Printer.Font.Charset = 1
    On Error GoTo ErrLabel
    
    Printer.FontSize = vFontSize * msgScaleFactorFonts
    Printer.FontBold = vFontBold
    Printer.FontItalic = vFontItalic
    
    'Set draw width to 2 pixels
    Printer.DrawWidth = 2
    
    'calculate width and height of boxes
    'Mo 21/2/2003, Print eForm Changes, reduce the width of printed Option Boxes
    sgOptionBoxWidth = Printer.TextWidth(String(oElement.QuestionLength + 2, "_")) + 100
    sgOptionBoxHeight = (Printer.TextHeight("X") + 100)
    
    'Check for OptionBoxes requiring space for status icons
    sgSpaceForIcons = 0
    If oElement.ShowStatusFlag Then
        sgSpaceForIcons = mn_SPACE_FOR_STATUS_ICON
    End If
    
    'Get the response data
    Set oResponse = oEFI.Responses.ResponseByElement(oElement, nRepeat)
    If Not oResponse Is Nothing Then
        sResponseValue = oResponse.Value
    Else
        sResponseValue = ""
    End If

    'Loop through the catergory Value list, drawing a box and printing the contents
    nOptionBoxCount = 0
    For Each oCategory In oElement.Categories
        ' DPH 27/06/2005 - Bug 2554 Only print active Option boxes
        If oCategory.Active Then
            'Set page position for the current box
            Printer.CurrentY = sgYPrint + (nOptionBoxCount * sgOptionBoxHeight)
            Printer.CurrentX = sgElementX
            'if current OptionBox is selected then draw it with a thicker border & underlined
            If oCategory.Value = sResponseValue Then
                Printer.DrawWidth = 6
                Printer.FontUnderline = True
            End If
            'draw box
            Printer.Line -Step(sgOptionBoxWidth, 0)
            Printer.Line -Step(0, sgOptionBoxHeight)
            Printer.Line -Step(-sgOptionBoxWidth, 0)
            Printer.Line -Step(0, -sgOptionBoxHeight)
            'set page position for caption centred in box and print caption
            Printer.CurrentY = sgYPrint + (nOptionBoxCount * sgOptionBoxHeight) _
                + (sgOptionBoxHeight / 6)
            Printer.CurrentX = sgElementX + ((sgOptionBoxWidth - Printer.TextWidth(oCategory.Value)) / 2)
    
            Printer.Print oCategory.Value
            'if current OptionBox is selected then put draw width back to normal
            If oCategory.Value = sResponseValue Then
                Printer.DrawWidth = 2
                Printer.FontUnderline = False
            End If
            'increment box counter
            nOptionBoxCount = nOptionBoxCount + 1
        End If
    Next
    
    'populate the optional paramaters that are used by calls from PrintRQG
    sgControlWidth = sgOptionBoxWidth + sgSpaceForIcons
    sgControlHeight = sgOptionBoxHeight * oElement.Categories.Count

Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.PrintOptionBoxes"
End Sub

'---------------------------------------------------------------------
Private Sub PrintOptionButtons(oEFI As EFormInstance, _
                                oElement As eFormElementRO, _
                                ByVal sgYPrint As Single, _
                                Optional ByVal sgXPrint As Single = 0, _
                                Optional ByVal nRepeat As Integer = 1, _
                                Optional ByRef sgControlWidth As Single, _
                                Optional ByRef sgControlHeight As Single)
'---------------------------------------------------------------------
' REVISIONS
' DPH 27/06/2005 - Bug 2554 Only print active Option buttons
'---------------------------------------------------------------------
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant
Dim sgOptionButtonWidth As Single
Dim sgOptionButtonHeight As Single
Dim nOptionButtonCount As Integer
Dim sResponseValue As String
Dim oResponse As Response
Dim oCategory As CategoryItem
Dim sgElementX As Single
Dim sgSpaceForIcons As Single

    On Error GoTo ErrLabel
    
    If sgXPrint <> 0 Then
        sgElementX = sgXPrint
    Else
        sgElementX = oElement.ElementX
    End If
    
    '  Look for specific font or use the form's default one
    Call GetFont(oElement, vFontName, vFontSize, vFontBold, vFontItalic)
    
    '  Set font attributes
    On Error Resume Next
    Printer.FontName = vFontName
    Printer.Font.Charset = 1
    On Error GoTo ErrLabel
    
    Printer.FontSize = vFontSize * msgScaleFactorFonts
    Printer.FontBold = vFontBold
    Printer.FontItalic = vFontItalic
    
    'set draw width to 2 pixels for option buttons
    Printer.DrawWidth = 2
    
    'calculate width and height of box surrounding option button
    sgOptionButtonWidth = Printer.TextWidth(String(oElement.QuestionLength + 4, "_")) + 100
    'Removed the scaling factor from the height
    sgOptionButtonHeight = (Printer.TextHeight("X") + 100)
    
    'Check for OptionButtons requiring space for status icons
    sgSpaceForIcons = 0
    If oElement.ShowStatusFlag Then
        sgSpaceForIcons = mn_SPACE_FOR_STATUS_ICON
    End If
    
    'Get the response data
    Set oResponse = oEFI.Responses.ResponseByElement(oElement, nRepeat)
    If Not oResponse Is Nothing Then
        sResponseValue = oResponse.Value
    Else
        sResponseValue = ""
    End If

    'Loop through the catergory Value list, drawing an option button alongside the value list text
    nOptionButtonCount = 0
    For Each oCategory In oElement.Categories
        ' DPH 27/06/2005 - Bug 2554 Only print active Option buttons
        If oCategory.Active Then
            'Set page position for the current option button
            Printer.CurrentY = sgYPrint + (sgOptionButtonHeight * (nOptionButtonCount + 0.5))
            Printer.CurrentX = sgElementX + 90
            'draw button
            Printer.Circle Step(0, 0), 70
            'if current button is selected then draw it with a solid inner circle
            If oCategory.Value = sResponseValue Then
                Printer.FillStyle = vbFSSolid
            End If
            Printer.Circle Step(0, 0), 45
            If oCategory.Value = sResponseValue Then
                Printer.FillStyle = vbFSTransparent
            End If
            'set page position for caption centred in box and print caption
            Printer.CurrentY = sgYPrint + (nOptionButtonCount * sgOptionButtonHeight) _
                + (sgOptionButtonHeight / 6)
            Printer.CurrentX = sgElementX + 120 + Printer.TextWidth("_")
            Printer.Print oCategory.Value
            'increment button counter
            nOptionButtonCount = nOptionButtonCount + 1
        End If
    Next
    
    'populate the optional paramaters that are used by calls from PrintRQG
    sgControlWidth = sgOptionButtonWidth + sgSpaceForIcons
    sgControlHeight = sgOptionButtonHeight * oElement.Categories.Count
    
Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.PrintOptionButtons"
End Sub

'---------------------------------------------------------------------
Private Sub PrintPicture(oElement As eFormElementRO, _
                        ByVal sgYPrint As Single)
'---------------------------------------------------------------------
'Pictures have to be loaded into the control frmMenu.picUsedForPrinting
'before printing.
'---------------------------------------------------------------------
Dim sgHeightBeforeScaling As Single
Dim sgWidthBeforeScaling As Single
Dim sgHeightAfterScaling As Single
Dim sgWidthAfterScaling As Single
Dim sPicFileName As String

    On Error GoTo ErrLabel

    '   Check that picture has been loaded successfully before printing
    On Error Resume Next
    sPicFileName = gsDOCUMENTS_PATH & oElement.Caption
    frmMenu.picUsedForPrinting.Picture = LoadPicture(sPicFileName)
    
    If Err.Number = 0 Then
        
        sgHeightBeforeScaling = frmMenu.picUsedForPrinting.Picture.Height
        sgWidthBeforeScaling = frmMenu.picUsedForPrinting.Picture.Width
        sgHeightAfterScaling = sgHeightBeforeScaling * msgScaleFactorFonts
        sgWidthAfterScaling = sgWidthBeforeScaling * msgScaleFactorFonts
'        Debug.Print "x1=" & oElement.ElementX & "  y1=" & sgYPrint _
'            & "  width=" & sgWidthAfterScaling & "  height=" & sgHeightAfterScaling _
'            & "  Bwidth=" & sgWidthBeforeScaling & "  Bheight=" & sgHeightBeforeScaling
        Printer.PaintPicture frmMenu.picUsedForPrinting.Picture, oElement.ElementX, sgYPrint, _
            sgWidthAfterScaling, sgHeightAfterScaling, , , sgWidthBeforeScaling, sgHeightBeforeScaling
'Original setting
'        Printer.PaintPicture frmMenu.picUsedForPrinting.Picture, oElement.ElementX, sgYPrint, _
'            sgWidthAfterScaling, sgHeightAfterScaling, , , sgWidthBeforeScaling, sgHeightBeforeScaling

    End If

Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.PrintPicture"
End Sub

'---------------------------------------------------------------------
Private Sub PrintTextBox(oEFI As EFormInstance, oElement As eFormElementRO, _
                        ByVal sgYPrint As Single, _
                        Optional ByVal sgXPrint As Single = 0, _
                        Optional ByVal nRepeat As Integer = 1, _
                        Optional ByRef sgControlWidth As Single, _
                        Optional ByRef sgControlHeight As Single)
'---------------------------------------------------------------------
' ic 30/08/2005 added clinical coding
'---------------------------------------------------------------------
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant
Dim nTextBoxCharWidth As Integer
Dim sgTextBoxWidth As Single
Dim sgTextBoxHeight As Single
Dim oResponse As Response
Dim sResponseValue As String
Dim sgSpaceForIcons As Single
Dim sgExtraSpace As Single
Dim sgElementX As Single
Dim nStartCharPosition As Integer
Dim nNextCharPosition As Integer
Dim nPrevCharPosition As Integer
Dim nNumberOfWraps As Integer

'ic 30/08/2005 clinical coding
Dim sgPicWidth As Single
Dim sgPicHeight As Single
Dim sgPicX As Single
Dim sgPicY As Single
Dim oPic As StdPicture

    On Error GoTo ErrLabel
    
    'Debug.Print "PrintTextBox> " & oElement.Caption & "  x= " & sgXPrint
    
    If sgXPrint <> 0 Then
        sgElementX = sgXPrint
    Else
        sgElementX = oElement.ElementX
    End If
    
    'Look for specific font or use the form's default one
    Call GetFont(oElement, vFontName, vFontSize, vFontBold, vFontItalic)
    
    'Set font attributes
    On Error Resume Next
    Printer.FontName = vFontName
    Printer.Font.Charset = 1
    On Error GoTo ErrLabel
    
    Printer.FontSize = vFontSize * msgScaleFactorFonts
    Printer.FontBold = vFontBold
    Printer.FontItalic = vFontItalic
    
    'Set draw width of text boxes to 2 pixels
    Printer.DrawWidth = 2
    
    'Set page position for text box
    Printer.CurrentY = sgYPrint
    Printer.CurrentX = sgElementX
    
    'Assess Width and Height of Text Box
    'Assess the number of characters that the width is based on
    ' NCJ 1 Jul 03 - Use new Expandable property of the element
    If oElement.Expandable Then
        nTextBoxCharWidth = oElement.DisplayLength
    Else
        nTextBoxCharWidth = oElement.QuestionLength
    End If
'    nTextBoxCharWidth = oElement.QuestionLength
'    If (oElement.ControlType = gn_TEXT_BOX) Or (oElement.ControlType = gn_POPUP_LIST) Then
'        If oElement.DisplayLength > 0 Then
'            nTextBoxCharWidth = oElement.DisplayLength
'        End If
'    End If
    
    sgTextBoxHeight = Printer.TextHeight("X") + 100
    'Add 4 characters to the width as in Display's CalculateBoxSize
    sgTextBoxWidth = Printer.TextWidth(String(nTextBoxCharWidth + 4, "_"))
    
    'Check for TextBox requiring space for status icons
    sgSpaceForIcons = 0
    If oElement.ShowStatusFlag Then
        sgSpaceForIcons = mn_SPACE_FOR_STATUS_ICON
    End If
    
    'Check for the need of additional space to display Calendar, dropdown, launchbox
    sgExtraSpace = 0
    
    'ic 20/08/2005 added clinical coding
    If (oElement.ControlType = gn_POPUP_LIST) _
    Or (oElement.ControlType = gn_CALENDAR) _
    Or (oElement.DataType = DataType.Thesaurus) _
    Or (oElement.Expandable) Then
        sgExtraSpace = sgTextBoxHeight
    End If
    
    'check for text box going outside right margin
    If sgElementX + sgTextBoxWidth + sgSpaceForIcons + sgExtraSpace > mneFormWidth Then
        sgTextBoxWidth = mneFormWidth - sgElementX - sgSpaceForIcons - sgExtraSpace
    End If
    
    'Draw a text box
    Printer.Line -Step(sgTextBoxWidth, 0)
    Printer.Line -Step(0, sgTextBoxHeight)
    Printer.Line -Step(-sgTextBoxWidth, 0)
    Printer.Line -Step(0, -sgTextBoxHeight)
    
    'Print the response data
    Set oResponse = oEFI.Responses.ResponseByElement(oElement, nRepeat)
    If Not oResponse Is Nothing Then
        sResponseValue = oResponse.Value
    Else
        sResponseValue = ""
    End If
    If sResponseValue <> "" Then
        'Assume ResponseValue is already correctly formatted
        'Assess if the text response will fit on a single line. If not wrap the text on to the next line
        nNumberOfWraps = 0
        nStartCharPosition = 1
        nPrevCharPosition = 0
        Do Until InStr(nStartCharPosition + nPrevCharPosition, sResponseValue, " ") = 0
            nNextCharPosition = InStr(nStartCharPosition + nPrevCharPosition, sResponseValue, " ")
            'Mo 21/2/2003, Print eForm Changes
            If Printer.TextWidth(Mid(sResponseValue, nStartCharPosition, nNextCharPosition)) > (sgTextBoxWidth) Then
                Printer.CurrentY = sgYPrint + (sgTextBoxHeight / 5) + (nNumberOfWraps * (Printer.TextHeight("X") + 100))
                'Test for the situation when the the first word in a response will not fit into the space available
                'when this occurs nPrevCharPosition will still be 0
                If nPrevCharPosition = 0 Then
                    DialogWarning ("Not enough space to print the response to Question " & oElement.Caption & ".")
                    'clear response that there is not enough space to print
                    sResponseValue = ""
                    Exit Do
                End If
                '50 twips added to x position to avoid overlap between textbox border and result text
                Printer.CurrentX = sgElementX + 50
                Printer.Print Mid(sResponseValue, nStartCharPosition, nPrevCharPosition - 1)
                sResponseValue = Mid(sResponseValue, nPrevCharPosition + 1)
                nStartCharPosition = 1
                nPrevCharPosition = 0
                nNumberOfWraps = nNumberOfWraps + 1
            Else
                nPrevCharPosition = nNextCharPosition
            End If
        Loop
        Printer.CurrentY = sgYPrint + (sgTextBoxHeight / 5) + (nNumberOfWraps * (Printer.TextHeight("X") + 100))
        '50 twips added to x position to avoid overlap between textbox border and result text
        Printer.CurrentX = sgElementX + 50
        Printer.Print Mid(sResponseValue, nStartCharPosition)
    End If
    
    'If its a Pop-up-list add a dropdown symbol alongside text box
    If oElement.ControlType = gn_POPUP_LIST Then
        'reset currentX and currentY
        Printer.CurrentY = sgYPrint
        Printer.CurrentX = sgElementX
        Printer.Line Step(sgTextBoxWidth, 0)-Step(sgTextBoxHeight, 0)
        Printer.Line -Step(0, sgTextBoxHeight)
        Printer.Line -Step(-sgTextBoxHeight, 0)
        Printer.DrawWidth = 4
        Printer.Line Step((sgTextBoxHeight / 4), (-sgTextBoxHeight / 2))-Step((sgTextBoxHeight / 2), 0)
        Printer.Line -Step((-sgTextBoxHeight / 4), (sgTextBoxHeight / 4))
        Printer.Line -Step((-sgTextBoxHeight / 4), (-sgTextBoxHeight / 4))
        Printer.DrawWidth = 2
    End If
    
    'If its a Calendar control add the calendar symbol alongside text box
    If oElement.ControlType = gn_CALENDAR Then
        Dim sgSixth As Single
        sgSixth = sgTextBoxHeight / 6
        'reset currentX and currentY
        Printer.CurrentY = sgYPrint
        Printer.CurrentX = sgElementX
        Printer.Line Step(sgTextBoxWidth, 0)-Step(sgTextBoxHeight, 0)
        Printer.Line -Step(0, sgTextBoxHeight)
        Printer.Line -Step(-sgTextBoxHeight, 0)
        Printer.Line Step(sgSixth, -(5 * sgSixth))-Step(0, (4 * sgSixth))
        Printer.Line Step(sgSixth, -(4 * sgSixth))-Step(0, (4 * sgSixth))
        Printer.Line Step(sgSixth, -(4 * sgSixth))-Step(0, (4 * sgSixth))
        Printer.Line Step(sgSixth, -(4 * sgSixth))-Step(0, (4 * sgSixth))
        Printer.Line Step(sgSixth, -(4 * sgSixth))-Step(0, (4 * sgSixth))
        Printer.Line -Step(-(4 * sgSixth), 0)
        Printer.Line Step((4 * sgSixth), -sgSixth)-Step(-(4 * sgSixth), 0)
        Printer.Line Step((4 * sgSixth), -sgSixth)-Step(-(4 * sgSixth), 0)
        Printer.Line Step((4 * sgSixth), -sgSixth)-Step(-(4 * sgSixth), 0)
        Printer.Line Step((4 * sgSixth), -sgSixth)-Step(-(4 * sgSixth), 0)
    End If
    
    'Is there a need for a launchbox
    If oElement.Expandable Then
        Dim sgForth As Single
        sgForth = sgTextBoxHeight / 4
        'reset currentX and currentY
        Printer.CurrentY = sgYPrint
        Printer.CurrentX = sgElementX
        Printer.Line Step(sgTextBoxWidth, 0)-Step(sgTextBoxHeight, 0)
        Printer.Line -Step(0, sgTextBoxHeight)
        Printer.Line -Step(-sgTextBoxHeight, 0)
        Printer.Circle Step(sgForth, -sgForth), 10
        Printer.Circle Step(sgForth, 0), 10
        Printer.Circle Step(sgForth, 0), 10
    End If
    
    'if its a thesaurus question
    If oElement.DataType = DataType.Thesaurus Then
        
        Printer.CurrentY = sgYPrint
        Printer.CurrentX = sgElementX
        
        'calculate pic height and width
        sgPicWidth = sgTextBoxHeight
        sgPicHeight = sgTextBoxHeight

        'calculate x and y
        sgPicX = sgElementX + sgTextBoxWidth
        sgPicY = sgYPrint
        
        'draw the pic
        Select Case oResponse.CodingStatus
        Case eCodingStatus.csAutoEncoded, eCodingStatus.csCoded:
            Set oPic = frmImages.imglistStatus.ListImages(DM30_ICON_DICTIONARY_CODED).Picture
        Case eCodingStatus.csDoNotCode:
            Set oPic = frmImages.imglistStatus.ListImages(DM30_ICON_DICTIONARY_DONOT).Picture
        Case eCodingStatus.csPendingNewCode:
            Set oPic = frmImages.imglistStatus.ListImages(DM30_ICON_DICTIONARY_PENDING).Picture
        Case eCodingStatus.csValidated:
            Set oPic = frmImages.imglistStatus.ListImages(DM30_ICON_DICTIONARY_VALIDATED).Picture
        Case Else: 'eCodingStatus.csEmpty, eCodingStatus.csNotCoded:
            Set oPic = frmImages.imglistStatus.ListImages(DM30_ICON_DICTIONARY).Picture
        End Select
        
        Printer.PaintPicture oPic, sgPicX, sgPicY, sgPicWidth, sgPicHeight
           
        'draw the box around the pic
        Printer.Line Step(sgTextBoxWidth, 0)-Step(sgTextBoxHeight, 0)
        Printer.Line -Step(0, sgTextBoxHeight)
        Printer.Line -Step(-sgTextBoxHeight, 0)
    End If
    
    'populate the optional paramaters that are used by calls from PrintRQG
    sgControlWidth = sgTextBoxWidth + sgSpaceForIcons + sgExtraSpace
    If nNumberOfWraps = 0 Then
        sgControlHeight = sgTextBoxHeight
    Else
        sgControlHeight = sgTextBoxHeight + (sgTextBoxHeight / 5) + (nNumberOfWraps * (Printer.TextHeight("X") + 100))
    End If
    
Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.PrintTextBox"
End Sub

'----------------------------------------------------------
Private Sub FontInit(oStudy As StudyDefRO)
'----------------------------------------------------------
'This sub started off as a copy of EFormBuilder.Init.
'Load the current studies default font details into variable private to this module.
'----------------------------------------------------------

    On Error GoTo ErrLabel
    
    ' Get the default font to use
    With oStudy
        If .FontName > "" Then
            msFontName = .FontName
        Else
            msFontName = ms_DEFAULT_FontName
        End If
        If .FontSize > 0 Then
            mnFontSize = .FontSize
        Else
            mnFontSize = mn_DEFAULT_FontSize
        End If
        mbFontBold = .FontBold
        mbFontItalic = .FontItalic
    End With
    
Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.FontInit"
End Sub

'---------------------------------------------------------------------
Private Sub GetFont(oElement As eFormElementRO, _
                    ByRef rFontName As Variant, _
                    ByRef rFontSize As Variant, _
                    ByRef rFontBold As Variant, _
                    ByRef rFontItalic As Variant)
'---------------------------------------------------------------------
'Get the default font settings for the form if the element doesn't have
'specified font settings.
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    If oElement.FontName = "" Then
        rFontName = msFontName
        rFontSize = mnFontSize
        rFontBold = mbFontBold
        rFontItalic = mbFontItalic
    Else
        rFontName = oElement.FontName
        rFontSize = oElement.FontSize
        rFontBold = oElement.FontBold
        rFontItalic = oElement.FontItalic
    End If

Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.GetFont"
End Sub

'---------------------------------------------------------------------
Private Sub GetCaptionFont(oElement As eFormElementRO, _
                    ByRef rFontName As Variant, _
                    ByRef rFontSize As Variant, _
                    ByRef rFontBold As Variant, _
                    ByRef rFontItalic As Variant)
'---------------------------------------------------------------------
'Get the default font settings for the form if the element's caption doesn't have
'specified font settings.
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    If oElement.CaptionFontName = "" Then
        rFontName = msFontName
        rFontSize = mnFontSize
        rFontBold = mbFontBold
        rFontItalic = mbFontItalic
    Else
        rFontName = oElement.CaptionFontName
        rFontSize = oElement.CaptionFontSize
        rFontBold = oElement.CaptionFontBold
        rFontItalic = oElement.CaptionFontItalic
    End If

Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.GetCaptionFont"
End Sub

'---------------------------------------------------------------------
Private Sub PrintRQG(oEFI As EFormInstance, oElement As eFormElementRO, _
                        ByVal sgYPrint As Single)
'---------------------------------------------------------------------
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant
Dim sgRQGFrameTop As Single
Dim sgRQGFrameLeft As Single
Dim sgRQGFrameWidth As Single
Dim sgRQGFrameHeight As Single
Dim sgRQGControlsHeight As Single
Dim oQGroupRO As QGroupRO
Dim oQGroupInstance As QGroupInstance
Dim oRQGElement As eFormElementRO
Dim nNumRows As Integer
Dim nRow As Integer
Dim nNumElements As Integer
Dim nElement As Integer
Dim sgHeaderWidth As Single
Dim sgHeaderHeight As Single
Dim sgMaxHeaderHeight As Single
Dim sgControlWidth As Single
Dim sgControlHeight As Single
Dim sgThisRowHeight As Single
Dim nNumLinesInCaption As Integer
Dim sCaption As String
Dim sPartOfCaption As String
Dim nPositionOfCR As Integer
Dim nNumExtraLinesCaption As Integer
Dim bCRExistsInCaption As Boolean

Const sglGAP As Single = 50

    On Error GoTo ErrLabel

    'Look for specific font or use the form's default one
    Call GetFont(oElement, vFontName, vFontSize, vFontBold, vFontItalic)
    
    'Set font attributes
    On Error Resume Next
    Printer.FontName = vFontName
    Printer.Font.Charset = 1
    On Error GoTo ErrLabel
    
    Printer.FontSize = vFontSize * msgScaleFactorFonts
    Printer.FontBold = vFontBold
    Printer.FontItalic = vFontItalic
    
    'Store co-ordinates of RQGFrame's top/left corner
    sgRQGFrameTop = sgYPrint
    sgRQGFrameLeft = oElement.ElementX
    
    'get the RQG Definition
    Set oQGroupRO = oElement.QGroup
    'get the number of RQG elements
    nNumElements = oQGroupRO.Elements.Count
    'get the RQG Instance
    Set oQGroupInstance = oEFI.QGroupInstanceById(oQGroupRO.QGroupId)
    'get the number of rows
    nNumRows = oQGroupInstance.Rows
    
    'Assess the height of the header
    sgMaxHeaderHeight = 0
    For nElement = 1 To nNumElements
        Set oRQGElement = oQGroupRO.Elements(nElement)
        Call GetCaptionFont(oRQGElement, vFontName, vFontSize, vFontBold, vFontItalic)
        On Error Resume Next
        Printer.FontName = vFontName
        Printer.Font.Charset = 1
        On Error GoTo ErrLabel
        Printer.FontSize = vFontSize * msgScaleFactorFonts
        Printer.FontBold = vFontBold
        Printer.FontItalic = vFontItalic
        'Mo 14/4/2004 Bug 2063
        nNumLinesInCaption = NumLinesInCaption(oRQGElement.Caption)
        sgHeaderHeight = (nNumLinesInCaption * Printer.TextHeight("X")) + sglGAP
        If sgHeaderHeight > sgMaxHeaderHeight Then
            sgMaxHeaderHeight = sgHeaderHeight
        End If
    Next nElement
    
    'initialize sgRQGFrameWidth & sgRQGControlsHeight to msglGAP a 50 twip border/gap
    'that is placed between everything in the RQG
    sgRQGControlsHeight = sglGAP
    For nRow = 1 To nNumRows
        sgRQGFrameWidth = sglGAP
        sgThisRowHeight = 0
        For nElement = 1 To nNumElements
            sgControlWidth = 0
            sgControlHeight = 0
            Set oRQGElement = oQGroupRO.Elements(nElement)
            'assess width of RQG element Header
            Call GetCaptionFont(oRQGElement, vFontName, vFontSize, vFontBold, vFontItalic)
            On Error Resume Next
            Printer.FontName = vFontName
            Printer.Font.Charset = 1
            On Error GoTo ErrLabel
            Printer.FontSize = vFontSize * msgScaleFactorFonts
            Printer.FontBold = vFontBold
            Printer.FontItalic = vFontItalic
            sgHeaderWidth = Printer.TextWidth(oRQGElement.Caption)
            'if its row 1 print the RQG element Header
            If nRow = 1 Then
                'Mo 14/4/2004 Bug 2063
                'checking for CR/LFs in the caption
                sCaption = RemoveNull(oRQGElement.Caption)
                bCRExistsInCaption = False
                nNumExtraLinesCaption = 0
                While sCaption <> ""
                    nPositionOfCR = InStr(sCaption, Chr(13))
                    If nPositionOfCR <> 0 Then
                        sPartOfCaption = Mid(sCaption, 1, nPositionOfCR - 1)
                        sCaption = Mid(sCaption, nPositionOfCR + 2, Len(sCaption))
                        bCRExistsInCaption = True
                    Else
                        sPartOfCaption = sCaption
                        sCaption = ""
                    End If
                    Printer.CurrentX = sgRQGFrameLeft + sgRQGFrameWidth
                    Printer.CurrentY = sgRQGFrameTop + (nNumExtraLinesCaption * Printer.TextHeight("X")) + sglGAP
                    Printer.Print RemoveNull(sPartOfCaption)
                    If bCRExistsInCaption Then
                        nNumExtraLinesCaption = nNumExtraLinesCaption + 1
                        bCRExistsInCaption = False
                    End If
                Wend
            End If
            
            'Print the control for this element and assess its Height and Width
            'Note that sgControlWidth and sgControlHeight are optional ByRef paramaters
            'to the control printing routines
            Select Case oRQGElement.ControlType
            Case gn_ATTACHMENT
                Call PrintAttachment(oRQGElement, sgRQGFrameTop + sgMaxHeaderHeight + sgRQGControlsHeight, sgRQGFrameLeft + sgRQGFrameWidth, sgControlWidth, sgControlHeight)
            Case gn_OPTION_BUTTONS
                Call PrintOptionButtons(oEFI, oRQGElement, sgRQGFrameTop + sgMaxHeaderHeight + sgRQGControlsHeight, sgRQGFrameLeft + sgRQGFrameWidth, nRow, sgControlWidth, sgControlHeight)
            Case gn_PUSH_BUTTONS
                Call PrintOptionBoxes(oEFI, oRQGElement, sgRQGFrameTop + sgMaxHeaderHeight + sgRQGControlsHeight, sgRQGFrameLeft + sgRQGFrameWidth, nRow, sgControlWidth, sgControlHeight)
            Case Else
                'TextBox or dropdown list or calendar
                Call PrintTextBox(oEFI, oRQGElement, sgRQGFrameTop + sgMaxHeaderHeight + sgRQGControlsHeight, sgRQGFrameLeft + sgRQGFrameWidth, nRow, sgControlWidth, sgControlHeight)
            End Select
            'Assess max width of element Header and element control
            If sgControlWidth > sgHeaderWidth Then
                sgRQGFrameWidth = sgRQGFrameWidth + sgControlWidth + sglGAP
            Else
                sgRQGFrameWidth = sgRQGFrameWidth + sgHeaderWidth + sglGAP
            End If
            'Assess Max height of controls in this row
            If sgControlHeight > sgThisRowHeight Then
                sgThisRowHeight = sgControlHeight
            End If
            'Check that there is enough width to print the next control.
            'This check prevents the printing of RQG controls that on the screen you have to scroll to view.
            If mneFormWidth - (sgRQGFrameLeft + sgRQGFrameWidth) < 800 Then
                nElement = nNumElements
            End If
        Next nElement
        'increment the y position
        sgRQGControlsHeight = sgRQGControlsHeight + sgThisRowHeight + sglGAP
    Next nRow
    
    'Add the Header height to the height of the Controls to give a FrameHeight
    sgRQGFrameHeight = sgMaxHeaderHeight + sgRQGControlsHeight
    
    'Draw frame around RQG
    'Mo 14/4/2004 Bug 2063
    If oQGroupRO.Border = True Then
        Printer.CurrentX = sgRQGFrameLeft
        Printer.CurrentY = sgRQGFrameTop
        Printer.Line -Step(sgRQGFrameWidth, 0)
        Printer.Line -Step(0, sgRQGFrameHeight)
        Printer.Line -Step(-sgRQGFrameWidth, 0)
        Printer.Line -Step(0, -sgRQGFrameHeight)
    End If
    
Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.PrintRQG"
End Sub

'---------------------------------------------------------------------
Private Sub PrintRQGEstimate(ByVal oEFI As EFormInstance, _
                            ByVal oElement As eFormElementRO, _
                            ByRef sgElementLength As Single, _
                            ByRef sgAdditionalDisplayLength As Single)
'---------------------------------------------------------------------
'This sub will work through a Repeating Question Group question and
'calculate its length (sgElementLength) and the amount of additional
'space required to print the eform, compared the amount of space
'required to display the eform.
'This sub mimics the printing activities of PrintRQG.
'---------------------------------------------------------------------
Dim oQGroupRO As QGroupRO
Dim oQGroupInstance As QGroupInstance
Dim oRQGElement As eFormElementRO
Dim nNumRows As Integer
Dim nRow As Integer
Dim nNumElements As Integer
Dim nNumDisplayRows As Integer
Dim nElement As Integer
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant
Dim sgHeaderHeight As Single
Dim sgScreenDisplayHeight As Single
Dim sgActualPrintHeight As Single
Dim sgRowHeight As Single
Dim sgControlWidth As Single
Dim sgControlHeight As Single
Dim oResponse As Response
Dim sResponseValue As String
Dim nNumberOfLines As Integer
Dim sgHeightafterWrap As Single
Dim sgThisRowHeight As Single
Dim nTextBoxCharWidth As Integer
Const sglGAP As Single = 50

    'get the RQG Definition
    Set oQGroupRO = oElement.QGroup
    'get the number of RQG elements
    nNumElements = oQGroupRO.Elements.Count
    'get the number of rows that are displayed on screen
    nNumDisplayRows = oQGroupRO.DisplayRows
    'get the RQG Instance
    Set oQGroupInstance = oEFI.QGroupInstanceById(oQGroupRO.QGroupId)
    'get the number of rows
    nNumRows = oQGroupInstance.Rows
    
    'Start off with the height of the header
    sgActualPrintHeight = 0
    For nElement = 1 To nNumElements
        Set oRQGElement = oQGroupRO.Elements(nElement)
        Call GetCaptionFont(oRQGElement, vFontName, vFontSize, vFontBold, vFontItalic)
        On Error Resume Next
        Printer.FontName = vFontName
        Printer.Font.Charset = 1
        On Error GoTo ErrLabel
        Printer.FontSize = vFontSize * msgScaleFactorFonts
        Printer.FontBold = vFontBold
        Printer.FontItalic = vFontItalic
        sgHeaderHeight = sglGAP + Printer.TextHeight("X") + sglGAP
        If sgHeaderHeight > sgActualPrintHeight Then
            sgActualPrintHeight = sgHeaderHeight
        End If
    Next nElement
    
    'Assess the normal row height by inspecting each control element within a single row
    sgRowHeight = 0
    For nElement = 1 To nNumElements
        Set oRQGElement = oQGroupRO.Elements(nElement)
        Call GetFont(oRQGElement, vFontName, vFontSize, vFontBold, vFontItalic)
        On Error Resume Next
        Printer.FontName = vFontName
        Printer.Font.Charset = 1
        On Error GoTo ErrLabel
        Printer.FontSize = vFontSize * msgScaleFactorFonts
        Printer.FontBold = vFontBold
        Printer.FontItalic = vFontItalic
        Select Case oRQGElement.ControlType
        Case gn_ATTACHMENT
            sgControlHeight = 375
        Case gn_OPTION_BUTTONS, gn_PUSH_BUTTONS
            sgControlHeight = (Printer.TextHeight("X") + 100) * oRQGElement.Categories.Count
        Case Else
            'TextBox or dropdown list or calendar
            sgControlHeight = Printer.TextHeight("X") + 100
        End Select
        If sgControlHeight > sgRowHeight Then
            sgRowHeight = sgControlHeight
        End If
    Next nElement
    
    'Now analyse each row looking for responses that require additional space
    For nRow = 1 To nNumRows
        sgThisRowHeight = sgRowHeight
        'analyse each control element within a row
        For nElement = 1 To nNumElements
            sgHeightafterWrap = 0
            Set oRQGElement = oQGroupRO.Elements(nElement)
            Select Case oRQGElement.ControlType
            Case gn_ATTACHMENT, gn_OPTION_BUTTONS, gn_PUSH_BUTTONS
                'Responses for these control types will never require additional space stemming from wrapping result text
            Case Else
                'TextBox or dropdown list or calendar that might require additional space
                'assess width in characters
                Call GetFont(oRQGElement, vFontName, vFontSize, vFontBold, vFontItalic)
                On Error Resume Next
                Printer.FontName = vFontName
                Printer.Font.Charset = 1
                On Error GoTo ErrLabel
                Printer.FontSize = vFontSize * msgScaleFactorFonts
                Printer.FontBold = vFontBold
                Printer.FontItalic = vFontItalic
                ' NCJ 1 Jul 03 - Use new Expandable property
                If oRQGElement.Expandable Then
                    nTextBoxCharWidth = oRQGElement.DisplayLength
                Else
                    nTextBoxCharWidth = oRQGElement.QuestionLength
                End If
'                If (oRQGElement.ControlType = gn_TEXT_BOX) Or (oRQGElement.ControlType = gn_POPUP_LIST) Then
'                    If oRQGElement.DisplayLength > 0 Then
'                        nTextBoxCharWidth = oRQGElement.DisplayLength
'                    End If
'                End If
                sgControlWidth = Printer.TextWidth(String(nTextBoxCharWidth + 4, "_")) + 100
                'get the response and check there is enough space to print it
                Set oResponse = oEFI.Responses.ResponseByElement(oRQGElement, nRow)
                If Not oResponse Is Nothing Then
                    sResponseValue = oResponse.Value
                Else
                    sResponseValue = ""
                End If
                If sResponseValue <> "" Then
                    If Printer.TextWidth(sResponseValue) > sgControlWidth Then
                        'estimate the number of additional lines
                        nNumberOfLines = (Printer.TextWidth(sResponseValue)) \ (sgControlWidth) + 1
                        sgHeightafterWrap = (nNumberOfLines * (Printer.TextHeight("X") + 100))
                    End If
                End If
            End Select
            
            'has this row got higher
            If sgHeightafterWrap > sgThisRowHeight Then
                 sgThisRowHeight = sgHeightafterWrap
            End If
        Next nElement
        'Calculate the accumaltive height of the RQG
        sgActualPrintHeight = sgActualPrintHeight + sgThisRowHeight + sglGAP
    Next nRow
    
    sgScreenDisplayHeight = sgHeaderHeight + (nNumDisplayRows * (sgRowHeight + sglGAP))
    sgElementLength = sgActualPrintHeight
    sgAdditionalDisplayLength = sgActualPrintHeight - sgScreenDisplayHeight
    'The above calculations can result in small differences, so if Abs(sgAdditionalDisplayLength) is small set it to zero
    If Abs(sgAdditionalDisplayLength) < 1 Then
        sgAdditionalDisplayLength = 0
    End If
    
    Debug.Print "[RQG Estimate] Actual= " & sgElementLength & "  Display= " & sgScreenDisplayHeight & "  Additional= " & sgAdditionalDisplayLength
     
    Set oQGroupRO = Nothing
    Set oQGroupInstance = Nothing
    Set oRQGElement = Nothing

Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.PrintRQGEstimate"
End Sub

'---------------------------------------------------------------------
Private Sub PrintHotLink(oElement As eFormElementRO, _
                    ByVal sgYPrint As Single)
'---------------------------------------------------------------------
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant

    On Error GoTo ErrLabel
    
    'Look for specific font or use the form's default one
    Call GetCaptionFont(oElement, vFontName, vFontSize, vFontBold, vFontItalic)
    
    'Set font attributes
    On Error Resume Next
    Printer.FontName = vFontName
    Printer.Font.Charset = 1
    On Error GoTo ErrLabel
    
    Printer.FontSize = vFontSize * msgScaleFactorFonts
    Printer.FontBold = vFontBold
    Printer.FontItalic = vFontItalic
    
    Printer.CurrentY = sgYPrint
    Printer.CurrentX = oElement.ElementX
    'Turn on Underlined text
    Printer.FontUnderline = True
    'Print the hotlink
    Printer.Print oElement.Caption
    'Turn off Underlined text
    Printer.FontUnderline = False

Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.PrintHotLink"
End Sub

'---------------------------------------------------------------------
Private Sub PrintAttachment(oElement As eFormElementRO, _
                            ByVal sgYPrint As Single, _
                            Optional ByVal sgXPrint As Single = 0, _
                            Optional ByRef sgControlWidth As Single, _
                            Optional ByRef sgControlHeight As Single)
'---------------------------------------------------------------------
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant
Dim sgBoxWidth As Single
Dim sgBoxHeight As Single
Dim sgElementX As Single
Dim sgSpaceForIcons As Single

    On Error GoTo ErrLabel
    
    If sgXPrint <> 0 Then
        sgElementX = sgXPrint
    Else
        sgElementX = oElement.ElementX
    End If
    
    'The font attributes of an attachment button are hard coded to "MS Sans Serif", size=8.5, Bold & Italic = False
    On Error Resume Next
    Printer.FontName = "MS Sans Serif"
    Printer.Font.Charset = 1
    On Error GoTo ErrLabel
    Printer.FontSize = 8.5 * msgScaleFactorFonts
    Printer.FontBold = False
    Printer.FontItalic = False
    
    Printer.CurrentY = sgYPrint
    Printer.CurrentX = sgElementX
    
    'Check for Attachment requiring space for status icons
    sgSpaceForIcons = 0
    If oElement.ShowStatusFlag Then
        sgSpaceForIcons = mn_SPACE_FOR_STATUS_ICON
    End If
    
    'Scale the size of the attachment button
    'sgBoxWidth = 1575 * msgScaleFactor
    'sgBoxHeight = 375 * msgScaleFactor
    sgBoxWidth = 1575
    sgBoxHeight = 375
    'draw the attachment button
    Printer.Line -Step(sgBoxWidth, 0)
    Printer.Line -Step(0, sgBoxHeight)
    Printer.Line -Step(-sgBoxWidth, 0)
    Printer.Line -Step(0, -sgBoxHeight)
    'print "Attach file..." in the attachment button
    Printer.CurrentY = sgYPrint + (sgBoxHeight / 4)
    Printer.CurrentX = sgElementX + ((sgBoxWidth - Printer.TextWidth("Attach file...")) / 2)
    Printer.Print "Attach file..."
    
    'populate the optional paramaters that are used by calls from PrintRQG
    sgControlWidth = sgBoxWidth + sgSpaceForIcons
    sgControlHeight = sgBoxHeight

Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "modEFormPrinting.PrintAttachment"
End Sub

'---------------------------------------------------------------------
Private Function NumLinesInCaption(ByVal sCaption As Variant)
'---------------------------------------------------------------------
'This function returns the number of lines within a caption
'called from PrintRQG
'---------------------------------------------------------------------
Dim nLineCount As Integer
Dim nPositionOfCR As Integer

    nLineCount = 1
    If IsNull(sCaption) Then
        NumLinesInCaption = nLineCount
        Exit Function
    End If
    
    While sCaption <> ""
        nPositionOfCR = InStr(sCaption, Chr(13))
        If nPositionOfCR <> 0 Then
            'strip off text before Cr/Lf
            sCaption = Mid(sCaption, nPositionOfCR + 2, Len(sCaption))
            'increment line counter
            nLineCount = nLineCount + 1
        Else
            sCaption = ""
        End If
    Wend
    
    NumLinesInCaption = nLineCount

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "NumLinesInCaption", "modEFormPrinting")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function
