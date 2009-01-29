Attribute VB_Name = "basBuildCRF"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       basBuildCRFPage
'   Author:     Andrew Newbigging, June 1998
'   Purpose:    Builds / prints CRFPage.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   WillC 13/7/00 Old Comments removed to Archived Comments Folder
'
'                                           Further changes to cope with ANSI padding
'   43      Andrew Newbigging   13/10/98    Added SkipCondition to BuildCRFElement
'                                           and check in BuildCaption for SkipCondition
'   44      Andrew Newbigging   15/10/98    Modified BuildCaption to cope with NULL values
'   45      Andrew Newbigging   15/10/98    Modified BuildCaption so that caption is
'                                           is disabled for SkipConditions in Data Entry only
'   46      Andrew Newbigging   15/10/98    Modified BuildCaption so that captio is invisible
'                                           if empty
'   47      Andrew Newbigging   22/10/98    Modified PrintCaption to cope with NULL values
'                                           Modified DisplayPopupList to size width
'                                           according to contents
'   48      Andrew Newbigging   13/11/98    Moved ConvertMaskFormat to basCommon
'   49      Andrew Newbigging   24/11/98
'   Popup list width modified to display all of the value column in DisplayPopupList
'           Mo Morris           24/11/98    SPR 587
'                                           PrintCRFForm changed to handle all errors
'                                           generated from a CommonDialog.ShowPrinter, including Cancel
'           Andrew Newbigging   18/3/99     SR732
'   Modified BuildMaskEdBox, BuildTextBox so that the width is truncated if it would run off
'   the edge of the form
'           Mo Morris           30/4/99     SR 871
'   SetPageSize changed so that tabCRF.width is set to the ScaleWidth of form frmCRFDesign
'   instead of the Width. This has the nock on effect of making the form's vertical scroll
'   bars appear in the corect place.
'   50      Paul Norris         02/09/99    Conserve memory by using the lblCRFElement icon
'   instead of loading from the resource
'   51      PN                  08/09/99    Changed Field Value to SpecialValue
'   NCJ     1/10/99 Use gsDOCUMENTS_PATH
'   Mo Morris           8/11/99
'   DAO to ADO conversion
'   NCJ     16/11/99        Don't mask DataFormat in BuildCRF
'   Mo Morris   23/11/99
'   Incorporate form printing changes from Macro version 1.6
'   Subs PrintCRFFormHeader  and BuildFormBeforePrinting added
'  WillC    Changed the Following where present from Integer to Long  ClinicalTrialId
'           CRFPageId,VisitId,CRFElementID
'   NCJ 13 Dec 99   PageId to Long
'                   Use MACROEnd in error handlers
'   NCJ 7 Jan 00    Use default font colour for captions, text etc.
'   Mo  8/2/00      DisplayPopupList re-written SR2930,2841
'   NCJ 9 Feb 00    BuildTextBox modified
'   NCJ 11 Feb 00   Print routines now always use NewValue rather than PreviousValue
'   Mo  15/3/00     SR 3203, PrintCaption now checks gnDisplaynumbers
'   Mo  6/6/00      SR 3543, correction made to DisplayCalendar
'   Mo  14/7/00     SR 3609, CRFElement!Height no longer read in BuildOptionButtons
'                   CRFElement!Height&Width no longer read in BuildRichTextBox
'   NCJ 15/9/00 - Added Lab test things to BuildCRFElement
'   TA 18/09/2000: Added Code to display lab test result status text in a label
'   NCJ 23/10/00 SR3951 - Hijacked PrintCRFForm to show Page Breaks on a form (see also BuildPageBreak)
'                       NB This is only for SD
'   Mo Morris   4/1/01  Changes made to DisplayPopupList
'   NCJ 5 May 2001  Minor cosmetic changes & commented code removal
'   Ash 24/07/2001  Minor cosmetic comments to allow page breaks to be displayed
'   Ash 31/07/2001  Minor cosmetic fix to fix snail trail error
'   REM 15/11/01 - Removed all code that refered to Data Entry
'   REM 18/11/01 - Added the Load for Eform Groups
'   NCJ 5 Dec 01 - Removed lots of unused code (see ElementBuilderSD.cls instead)
'   TA 5/7/02: Rolled forward 2.2 bug fix in BuildPAgeBreak
'   RS 14/01/03 Corrected bug in SetPageSize: would bail out while re-showing designform
'               wich caused scroll bars not being updated correctly.
'   Mo 28/5/2003,   Incorporated changes so that RQGs can be printed. Includes new procedures
'                   PrintRQG, PrintRQGEstimate, RQGElementCatCodes, RQGElementById, GetCaptionFont
'                   PrintHotLink and PrintAttachmen
'   Mo 3/6/2003 Bug 1819
'   RS 12/06/2003: Re-activated ClearPageBreaks in PrintCRFForm, as clicking View Page Break twice would crash MACRO
'   Mo 14/4/2004    Bug 2063. CR/LFs in repeating Question Group Question/Header captions now
'                   handled correctly. Changes to PrintRQG + new function NumLinesInCaption.
'                   RQG Borders are no longer printed when not required.
'   TA  07/07/2005: seeting of Font.Charset = 1 to allow eastern european characters. CBD2591.
'   ic 14/06/2005   added clinical coding
'   Mo  26/10/2007  Bug 2957, minor change to PrintCRFForm when page end corresponds to a hotlink.
'------------------------------------------------------------------------------------
Option Explicit
Option Base 0
Option Compare Binary

' NCJ 23/10/00 - Set to TRUE if we only want to see page breaks on a page
Public gbShowPageBreaksOnly As Boolean

' These are temporarily here until we move them somewhere else...
'Public EForm As EFormSD

Private msgScaleFactor As Single
Private msgScaleFactorFonts As Single
'added by Mo Morris 19/8/99
Public gnDisplayNumbers As Integer

' OLD constants for control types (MACRO 2.2 and earlier)
Public Const gnTEXT_BOX = 1
Public Const gnOPTION_BUTTONS = 2
Public Const gnPOPUP_LIST = 4
Public Const gnCALENDAR = 8
Public Const gnRICH_TEXT_BOX = 16
Public Const gnATTACHMENT = 32
' Public Const gnDIAGRAM = 128    ' not implemented
Public Const gnPUSH_BUTTONS = 256
Public Const gnMASK_ED_BOX = 512
Public Const gnVISUAL_ELEMENT = 16384
Public Const gnLINE = 1
Public Const gnCOMMENT = 2
Public Const gnPICTURE = 4

Private mnNextOptionButtonIndex As Integer
Private mnTabOrder As Integer

Dim mnPageCount As Integer
Dim msgTitleGap As Single

Private Const mn_SPACE_FOR_STATUS_ICON = 495
Private Const mnCalendarControlFudge = 200

'The following declaration and constants are for keeping windows always on top
' Declaration of a Windows routine.
' This statement should be placed in the module.
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' NCJ 26 Nov 01 - NEW constants for control types (MACRO 3.0 onwards)
' (Use these rather than the old ones)
Public Const gn_TEXT_BOX = 1
Public Const gn_OPTION_BUTTONS = 2
Public Const gn_POPUP_LIST = 4
Public Const gn_CALENDAR = 8
Public Const gn_RICH_TEXT_BOX = 16
Public Const gn_ATTACHMENT = 32
Public Const gn_PUSH_BUTTONS = 258
Public Const gn_MASK_ED_BOX = 512
Public Const gn_LINE = 16385
Public Const gn_COMMENT = 16386
Public Const gn_PICTURE = 16388
Public Const gn_HOTLINK = 16390     ' Added NCJ 6 Nov 02

'ZA 09/09/2002 - decided the control type option button
Public gnUseOptionButton As Integer
'Key values for using option buttons
Public Const gn_ALWAYS_USE_OPTION_BUTTONS = 9999
Public Const gn_NEVER_USE_OPTION_BUTTONS = 0
Public Const gn_ALWAYS_USE_OPTION_MENU = 11
Public Const gn_NEVER_USE_OPTION_MENU = 10
Public Const gs_USE_OPTION_BUTTONS = "Use Option Buttons"
Public Const gn_DEFAULT_OPTION_MENU_VALUE = 2
'Key for Default RFC
Public Const gs_DEFAULT_RFC = "DefaultRFC"
'Key for display/hide question status icons
Public Const gs_HIDE_QUESTION_STATUS_ICON = "Hide Question Status Icon"
'key for display/hide repeating questiong group icons
Public Const gs_HIDE_RQG_STATUS_ICON = "Hide RQG Status Icon"
'key for automatic numbering
Public Const gs_AUTOMATIC_NUMBERING = "Automatic Numbering"

'---------------------------------------------------------------------
Private Sub BuildPageBreak(ByRef vForm As Form, _
                        ByVal nPageNo As Integer, ByVal sY As Single)
'---------------------------------------------------------------------
' NCJ 23/10/00 - Add a page break line at end of given page number at position sY
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    '  Load a page break line
    Load vForm.linPageBreak(nPageNo)
    
    With vForm.linPageBreak(nPageNo)
        .Container = vForm.picCRFPage
        .X1 = 0
        .X2 = .Container.Width
        .Y1 = sY
        .Y2 = sY
        .BorderColor = vForm.DefaultFontColour
        .Visible = True
    End With

    Exit Sub

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                            "BuildPageBreak", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub ClearPageBreaks(ByRef vForm As Form)
'---------------------------------------------------------------------
' NCJ 23/10/00 - Clear all page break lines from frmCRFDesign
'---------------------------------------------------------------------
Dim i As Integer
    
    With vForm
        If .linPageBreak.Count > 1 Then
            For i = .linPageBreak.Count - 1 To 1 Step -1
                Unload .linPageBreak(i)
            Next i
        End If
    End With

End Sub

'---------------------------------------------------------------------
Private Function GetFontColour(ByVal vForm As Form, _
                    ByVal vCRFElement As ADODB.Recordset) As Long
'---------------------------------------------------------------------
' NCJ 7 Jan 00
' Get the font colour for this CRFElement
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    If IsNull(vCRFElement!FontColour) Or vCRFElement!FontColour = 0 Then
        GetFontColour = vForm.DefaultFontColour
    Else
        GetFontColour = vCRFElement!FontColour
    End If

    Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "GetFontColour", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'---------------------------------------------------------------------
Private Sub GetFont(ByVal vForm As Form, _
                    ByVal vCRFElement As ADODB.Recordset, _
                    ByRef rFontName As Variant, _
                    ByRef rFontSize As Variant, _
                    ByRef rFontBold As Variant, _
                    ByRef rFontItalic As Variant)
'---------------------------------------------------------------------
'Get the default font for the form if the element doesn't have a
'specified font
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    If IsNull(vCRFElement!FontName) Or vCRFElement!FontName = "" Then
        rFontName = vForm.DefaultFontName
        rFontSize = vForm.DefaultFontSize
        rFontBold = vForm.DefaultFontBold
        rFontItalic = vForm.DefaultFontItalic
    Else
        rFontName = vCRFElement!FontName
        rFontSize = vCRFElement!FontSize
        rFontBold = vCRFElement!FontBold
        rFontItalic = vCRFElement!FontItalic
    End If

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetFont", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub


'---------------------------------------------------------------------
Public Sub SetPageSize(ByVal vForm As Form)
'---------------------------------------------------------------------

'If a form has been made very small or minimized then parts of this
'subroutine will result in less than zero calculations, hence the error trap
On Error GoTo FormWasTooSmall

    'SDM 27/01/00 SR2817
    If vForm.WindowState = vbMinimized Then Exit Sub
    
    ' Set position and size of container controls
    vForm.tabCRF.MultiRow = False
    vForm.tabCRF.TabFixedHeight = 495
    vForm.tabCRF.Height = 400
    
    'changed by Mo Morris 30/4/99 (.width changed to ScaleWidth SR 871
    'vForm.tabCRF.Width = vForm.Width - vForm.tabCRF.Left
    
    vForm.Show 'SDM 13/01/00 SR2664
        
    ' RS 14/01/2003: statement above will fail (modal/non-modal problem). This error will now be
    ' ignored, and the rest of the adjustments will be handled correctly.
    
continue1:
    vForm.tabCRF.Left = 0
    vForm.tabCRF.Top = 0
    
    vForm.tabCRF.Width = vForm.ScaleWidth - vForm.tabCRF.Left

    If vForm.tabCRF.Tabs.Count > 0 Then
        vForm.tabCRF.Visible = True
    Else
        vForm.tabCRF.Visible = False
    End If
    
    vForm.frmCRFPage.Top = vForm.tabCRF.Top + vForm.tabCRF.Height + 50
    vForm.frmCRFPage.Left = vForm.tabCRF.Left
    vForm.frmCRFPage.Height = vForm.Height - vForm.frmCRFPage.Top
    vForm.frmCRFPage.Width = vForm.tabCRF.Width
    
    vForm.picCRFPage.Top = 0
    vForm.picCRFPage.Left = 0
    'vForm.picCRFPage.AutoSize = True

    If vForm.picCRFPage.Height > vForm.frmCRFPage.Height And vForm.tabCRF.Tabs.Count > 0 Then
        Set vForm.vsbCRFPage.Container = vForm.frmCRFPage
        'note that the max value is divided by 10 to accomodate long forms
        vForm.vsbCRFPage.Max = vForm.picCRFPage.Height / 10
        vForm.vsbCRFPage.Top = 0
        If vForm.picCRFPage.Width > vForm.frmCRFPage.Width Then
            vForm.vsbCRFPage.Height = vForm.frmCRFPage.Height - vForm.hsbCRFPage.Height * 2
        Else
            vForm.vsbCRFPage.Height = vForm.frmCRFPage.Height - vForm.hsbCRFPage.Height
        End If
        vForm.vsbCRFPage.Left = vForm.frmCRFPage.Width - vForm.vsbCRFPage.Width
        vForm.vsbCRFPage.LargeChange = (vForm.frmCRFPage.Height / 2) / 10
        vForm.vsbCRFPage.SmallChange = (vForm.frmCRFPage.Height / 10) / 10
        vForm.vsbCRFPage.Value = 0
        vForm.vsbCRFPage.Visible = True
    Else
        vForm.vsbCRFPage.Visible = False
    End If
    
    If vForm.picCRFPage.Width > vForm.frmCRFPage.Width And vForm.tabCRF.Tabs.Count > 0 Then
        Set vForm.hsbCRFPage.Container = vForm.frmCRFPage
        If vForm.vsbCRFPage.Visible = True Then
            vForm.hsbCRFPage.Width = vForm.frmCRFPage.Width - vForm.vsbCRFPage.Width
        Else
            vForm.hsbCRFPage.Width = vForm.frmCRFPage.Width
        End If
        vForm.hsbCRFPage.Max = vForm.picCRFPage.Width
        vForm.hsbCRFPage.Top = vForm.frmCRFPage.Height - vForm.hsbCRFPage.Height * 2
        vForm.hsbCRFPage.Left = 0
        vForm.hsbCRFPage.LargeChange = vForm.picCRFPage.Width / 2
        vForm.hsbCRFPage.SmallChange = vForm.picCRFPage.Width / 10
        vForm.hsbCRFPage.Value = 0
        vForm.hsbCRFPage.Visible = True
    Else
        vForm.hsbCRFPage.Visible = False
    End If

    'vForm.Refresh

Exit Sub

FormWasTooSmall:
    ' Ignore error "Can't show non-modal form when modal form is displayed"
    If Err.Number = 401 Then
        GoTo continue1
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub PrintCRFElement(ByVal vForm As Form, _
                           ByVal vCRFElement As ADODB.Recordset, _
                           ByVal sgYPrint As Single, _
                           ByVal vCaptionYPrint As Single)
'---------------------------------------------------------------------
' Print a CRFElement
' NCJ 23/10/00 - Removed unused Element assignations
'---------------------------------------------------------------------
Dim nControlType As Integer

    On Error GoTo ErrHandler
    
    nControlType = vCRFElement!ControlType
    
    If nControlType < gnVISUAL_ELEMENT Then
        If RemoveNull(vCRFElement!Caption) > "" Then
            PrintCaption vForm, vCRFElement, sgYPrint, vCaptionYPrint
        End If
    End If
        
    'draw the relevant element
    Select Case nControlType
    Case 0
        PrintRQG vForm, vCRFElement, sgYPrint
    Case gnATTACHMENT
        PrintAttachment vForm, vCRFElement, sgYPrint
    Case gnOPTION_BUTTONS
        PrintOptionButtons vForm, vCRFElement, sgYPrint
    Case 258
        PrintOptionBoxes vForm, vCRFElement, sgYPrint
    Case 16385
        PrintLine vForm, vCRFElement, sgYPrint
    Case 16386
        PrintComment vForm, vCRFElement, sgYPrint
    Case 16388
        PrintPicture vForm, vCRFElement, sgYPrint
    Case 16390
        PrintHotLink vForm, vCRFElement, sgYPrint
    Case Else
        PrintTextBox vForm, vCRFElement, sgYPrint
    End Select
    
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintCRFElement", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Public Sub PrintCaption(ByVal vForm As Form, _
                        ByVal vCRFElement As ADODB.Recordset, _
                        ByVal sgYPrint As Single, _
                        ByVal vCaptionYPrint As Single)
'---------------------------------------------------------------------
' Changed by Mo Morris 2/9/99
' Routine now handles multiple line captions inline with multiple line comments
' NCJ 26/10/00 - Fixed bug in printing caption (sCaption -> sCaptionLine)
'---------------------------------------------------------------------
Dim sUnitOfMeasurement As String
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant
Dim sCaption As String
Dim sCaptionLine As String
Dim nLineCount As Integer
Dim sPositionOfCR As Integer
Dim sglTextHeight As Single

    On Error GoTo ErrHandler

    '  Prepare unit of measurement part of label
    sUnitOfMeasurement = RemoveNull(vCRFElement!UnitOfMeasurement)
    If sUnitOfMeasurement > "" Then
        ' Add brackets round unit
        sUnitOfMeasurement = " (" & sUnitOfMeasurement & ")"
    End If
    
    'build caption text
    'changed Mo Morris 15/3/00, SR 3203, check for caption numbers being turned off
    If gnDisplayNumbers Then
        sCaption = CStr(vCRFElement!FieldOrder) & ". " & vCRFElement!Caption & sUnitOfMeasurement
    Else
        sCaption = vCRFElement!Caption & sUnitOfMeasurement
    End If
    
    'Look for specific font or use the form's default one
    GetCaptionFont vForm, vCRFElement, vFontName, vFontSize, vFontBold, vFontItalic
    
    'Set font attributes
    On Error Resume Next
    Printer.FontName = vFontName
    Printer.Font.Charset = 1
    On Error GoTo ErrHandler
    
    Printer.FontSize = vFontSize * msgScaleFactorFonts
    Printer.FontBold = vFontBold
    Printer.FontItalic = vFontItalic
    
    'set page position for Caption
    'needs amending to cope with margins and negative values
    
    nLineCount = 0
    ' Get text height
    sglTextHeight = Printer.TextHeight("X")
    
    While sCaption <> ""
        sPositionOfCR = InStr(sCaption, Chr(13))
        If sPositionOfCR <> 0 Then
            'strip off text before Cr/LF
            sCaptionLine = Mid(sCaption, 1, sPositionOfCR - 1)
            sCaption = Mid(sCaption, sPositionOfCR + 2, Len(sCaption))
        Else
            sCaptionLine = sCaption
            sCaption = ""
        End If
        'set X co-ordinate position for caption line
        If vCRFElement!CaptionX > 0 Then
            Printer.CurrentX = vCRFElement!CaptionX
        Else
            ' NCJ 26/10/00 - Changed sCaption to sCaptionLine
            Printer.CurrentX = vCRFElement!X - Printer.TextWidth(sCaptionLine) - 50
        End If
        'set Y co-ordinate position for caption line
        If vCRFElement!CaptionY > 0 Then
            Printer.CurrentY = vCaptionYPrint + (sglTextHeight * nLineCount)
        Else
            Printer.CurrentY = sgYPrint + (sglTextHeight * nLineCount)
        End If
        nLineCount = nLineCount + 1
        'print the caption line
        Printer.Print sCaptionLine
    Wend

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintCaption", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Public Sub PrintTextBox(ByVal vForm As Form, _
                        ByVal vCRFElement As ADODB.Recordset, _
                        ByVal sgYPrint As Single, _
                        Optional ByVal sgXPrint As Single = 0, _
                        Optional ByRef sgControlWidth As Single, _
                        Optional ByRef sgControlHeight As Single)
'ic 14/06/2005 added clinical coding
'---------------------------------------------------------------------
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant
Dim nTextBoxCharWidth As Integer
Dim sgTextBoxWidth As Single
Dim sgTextBoxHeight As Single
Dim sgSpaceForIcons As Single
Dim sgExtraSpace As Single
Dim sgElementX As Single
'ic 15/06/2005 clinical coding: browse button variables
Dim sgPicWidth As Single
Dim sgPicHeight As Single
Dim sgPicX As Single
Dim sgPicY As Single


On Error GoTo ErrHandler

If sgXPrint <> 0 Then
    sgElementX = sgXPrint
Else
    sgElementX = vCRFElement!X
End If

'Look for specific font or use the form's default one
GetFont vForm, vCRFElement, vFontName, vFontSize, vFontBold, vFontItalic

'Set font attributes
On Error Resume Next
Printer.FontName = vFontName
Printer.Font.Charset = 1
On Error GoTo ErrHandler

Printer.FontSize = vFontSize * msgScaleFactorFonts
Printer.FontBold = vFontBold
Printer.FontItalic = vFontItalic

'Set draw width of text boxes to 2 pixels
Printer.DrawWidth = 2

'Set page position for text box
Printer.CurrentY = sgYPrint
Printer.CurrentX = sgElementX

'Assess Width and Height of Text Box
nTextBoxCharWidth = vCRFElement!DataItemLength

'ic 14/06/2005 clinical coding: thesaurus types cannot have expand button
If (vCRFElement!DataType <> DataType.Thesaurus) Then
    If (vCRFElement!ControlType = gnTEXT_BOX) Or (vCRFElement!ControlType = gnPOPUP_LIST) Then
        If vCRFElement!DisplayLength > 0 Then
            nTextBoxCharWidth = vCRFElement!DisplayLength
        End If
    End If
End If

sgTextBoxHeight = Printer.TextHeight("X") + 100
sgTextBoxWidth = Printer.TextWidth(String(nTextBoxCharWidth + 4, "_"))

'Check for TextBox requiring space for status icons
sgSpaceForIcons = 0
If vCRFElement!ShowStatusFlag = 1 Then
    sgSpaceForIcons = mn_SPACE_FOR_STATUS_ICON
End If

'Check for the need of additional space to display Calendar, dropdown, launchbox
sgExtraSpace = 0

'ic 14/06/2005 clinical coding: thesaurus types can have browse button
If (vCRFElement!DataType = DataType.Thesaurus) Then
    sgExtraSpace = sgTextBoxHeight
Else

    If (vCRFElement!ControlType = gn_POPUP_LIST) _
    Or (vCRFElement!ControlType = gn_CALENDAR) _
    Or ((vCRFElement!ControlType = gn_TEXT_BOX) And (vCRFElement!DisplayLength > 0)) Then
        sgExtraSpace = sgTextBoxHeight
    End If

End If


'check for text box going outside right margin
If sgElementX + sgTextBoxWidth + sgSpaceForIcons + sgExtraSpace > vForm.picCRFPage.Width Then
    sgTextBoxWidth = vForm.picCRFPage.Width - sgElementX - sgSpaceForIcons - sgExtraSpace
End If

'Draw a text box
Printer.Line -Step(sgTextBoxWidth, 0)
Printer.Line -Step(0, sgTextBoxHeight)
Printer.Line -Step(-sgTextBoxWidth, 0)
Printer.Line -Step(0, -sgTextBoxHeight)

'If its a Pop-up-list add a dropdown symbol alongside text box
If vCRFElement!ControlType = gnPOPUP_LIST Then
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
If vCRFElement!ControlType = gn_CALENDAR Then
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

'ic 14/06/2005 clinical coding: thesaurus types cannot have browse button
If (vCRFElement!DataType = DataType.Thesaurus) Then
    
    Printer.CurrentY = sgYPrint
    Printer.CurrentX = sgElementX
    
    'calculate pic height and width
    sgPicWidth = sgTextBoxHeight
    sgPicHeight = sgTextBoxHeight

    'calculate x and y
    sgPicX = sgElementX + sgTextBoxWidth
    sgPicY = sgYPrint

    'draw the pic
    Printer.PaintPicture vForm.picDictionary.Picture, sgPicX, sgPicY, sgPicWidth, sgPicHeight
       
    'draw the box around the pic
    Printer.Line Step(sgTextBoxWidth, 0)-Step(sgTextBoxHeight, 0)
    Printer.Line -Step(0, sgTextBoxHeight)
    Printer.Line -Step(-sgTextBoxHeight, 0)

Else

        'Is there a need for a launchbox
        If ((vCRFElement!ControlType = gn_TEXT_BOX) And (vCRFElement!DisplayLength > 0)) Then
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
    End If


'populate the optional paramaters that are used by calls from PrintRQG
sgControlWidth = sgTextBoxWidth + sgSpaceForIcons + sgExtraSpace
sgControlHeight = sgTextBoxHeight

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintTextBox", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Public Sub PrintOptionBoxes(ByVal vForm As Form, _
                            ByVal vCRFElement As ADODB.Recordset, _
                            ByVal sgYPrint As Single, _
                            Optional ByVal sgXPrint As Single = 0, _
                            Optional ByRef sgControlWidth As Single, _
                            Optional ByRef sgControlHeight As Single)
'---------------------------------------------------------------------
Dim rsValues As ADODB.Recordset
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant
Dim sgOptionBoxWidth As Single
Dim sgOptionBoxHeight As Single
Dim nOptionBoxCount As Integer
Dim sgElementX As Single
Dim sgSpaceForIcons As Single

On Error GoTo ErrHandler

If sgXPrint <> 0 Then
    sgElementX = sgXPrint
Else
    sgElementX = vCRFElement!X
End If

'  Look for specific font or use the form's default one
GetFont vForm, vCRFElement, vFontName, vFontSize, vFontBold, vFontItalic

'  Set font attributes
On Error Resume Next
Printer.FontName = vFontName
Printer.Font.Charset = 1
On Error GoTo ErrHandler

Printer.FontSize = vFontSize * msgScaleFactorFonts
Printer.FontBold = vFontBold
Printer.FontItalic = vFontItalic

'Set draw width to 2 pixels
Printer.DrawWidth = 2

'calculate width and height of boxes
sgOptionBoxWidth = Printer.TextWidth(String(vCRFElement!DataItemLength + 2, "_")) + 100
'Removed the scaling factor from the height
sgOptionBoxHeight = (Printer.TextHeight("X") + 100)

'Check for OptionBoxes requiring space for status icons
sgSpaceForIcons = 0
If vCRFElement!ShowStatusFlag = 1 Then
    sgSpaceForIcons = mn_SPACE_FOR_STATUS_ICON
End If
    
'Get contents of the Value list
Set rsValues = New ADODB.Recordset
Set rsValues = gdsDataValues(vForm.ClinicalTrialId, vForm.VersionId, _
                            vCRFElement!DataItemId)

'Loop through the Value list, drawing a box and printing the contents
nOptionBoxCount = 0
While Not rsValues.EOF
    'Set page position for the current box
    Printer.CurrentY = sgYPrint + (nOptionBoxCount * sgOptionBoxHeight)
    Printer.CurrentX = sgElementX

    'draw box
    Printer.Line -Step(sgOptionBoxWidth, 0)
    Printer.Line -Step(0, sgOptionBoxHeight)
    Printer.Line -Step(-sgOptionBoxWidth, 0)
    Printer.Line -Step(0, -sgOptionBoxHeight)
    'set page position for caption centred in box and print caption
    Printer.CurrentY = sgYPrint + (nOptionBoxCount * sgOptionBoxHeight) _
        + (sgOptionBoxHeight / 6)
    Printer.CurrentX = sgElementX + ((sgOptionBoxWidth - Printer.TextWidth(rsValues!ItemValue)) / 2)
    
    ' PN 08/09/99 changed field Value to ItemValue
    Printer.Print rsValues!ItemValue

    'increment box counter
    nOptionBoxCount = nOptionBoxCount + 1
    rsValues.MoveNext
Wend

'populate the optional paramaters that are used by calls from PrintRQG
sgControlWidth = sgOptionBoxWidth + sgSpaceForIcons
sgControlHeight = sgOptionBoxHeight * rsValues.RecordCount

rsValues.Close
Set rsValues = Nothing

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintOptionBoxes", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Public Sub PrintOptionButtons(ByVal vForm As Form, _
                                ByVal vCRFElement As ADODB.Recordset, _
                                ByVal sgYPrint As Single, _
                                Optional ByVal sgXPrint As Single = 0, _
                                Optional ByRef sgControlWidth As Single, _
                                Optional ByRef sgControlHeight As Single)
'---------------------------------------------------------------------
Dim rsValues As ADODB.Recordset
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant
Dim sgOptionButtonWidth As Single
Dim sgOptionButtonHeight As Single
Dim nOptionButtonCount As Integer
Dim sgElementX As Single
Dim sgSpaceForIcons As Single

On Error GoTo ErrHandler

If sgXPrint <> 0 Then
    sgElementX = sgXPrint
Else
    sgElementX = vCRFElement!X
End If

'  Look for specific font or use the form's default one
GetFont vForm, vCRFElement, vFontName, vFontSize, vFontBold, vFontItalic

'  Set font attributes
On Error Resume Next
Printer.FontName = vFontName
Printer.Font.Charset = 1
On Error GoTo ErrHandler

Printer.FontSize = vFontSize * msgScaleFactorFonts
Printer.FontBold = vFontBold
Printer.FontItalic = vFontItalic

'set draw width to 2 pixels for option buttons
Printer.DrawWidth = 2

'calculate width and height of box surrounding option button
sgOptionButtonWidth = Printer.TextWidth(String(vCRFElement!DataItemLength + 4, "_")) + 100
'Removed the scaling factor from the height
sgOptionButtonHeight = (Printer.TextHeight("X") + 100)

'Check for OptionButtons requiring space for status icons
sgSpaceForIcons = 0
If vCRFElement!ShowStatusFlag = 1 Then
    sgSpaceForIcons = mn_SPACE_FOR_STATUS_ICON
End If

'Get contents of the Value list
Set rsValues = New ADODB.Recordset
Set rsValues = gdsDataValues(vForm.ClinicalTrialId, vForm.VersionId, _
                            vCRFElement!DataItemId)

' PN 08/09/99
' change field Value to ItemValue
'Loop through the Value list, drawing an option button alongside the value list text
nOptionButtonCount = 0
While Not rsValues.EOF
    'Set page position for the current option button
    Printer.CurrentY = sgYPrint + (sgOptionButtonHeight * (nOptionButtonCount + 0.5))
    Printer.CurrentX = sgElementX + 90
    'draw button
    Printer.Circle Step(0, 0), 70
    'if Macro_DM and current button is selected then draw it with a solid inner circle

    Printer.Circle Step(0, 0), 45

    'set page position for caption centred in box and print caption
    Printer.CurrentY = sgYPrint + (nOptionButtonCount * sgOptionButtonHeight) _
        + (sgOptionButtonHeight / 6)
    Printer.CurrentX = sgElementX + 120 + Printer.TextWidth("_")
    Printer.Print rsValues!ItemValue
    'increment button counter
    nOptionButtonCount = nOptionButtonCount + 1
    rsValues.MoveNext
Wend

'populate the optional paramaters that are used by calls from PrintRQG
sgControlWidth = sgOptionButtonWidth + sgSpaceForIcons
sgControlHeight = sgOptionButtonHeight * rsValues.RecordCount

rsValues.Close
Set rsValues = Nothing

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintOptionButtons", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub PrintLine(ByVal vForm As Form, _
                    ByVal vCRFElement As ADODB.Recordset, _
                    ByVal sgYPrint As Single)
'---------------------------------------------------------------------
On Error GoTo ErrHandler

Printer.CurrentY = sgYPrint
Printer.CurrentX = 0

'set draw width to 3 pixels for lines and draw a horizontal line
Printer.DrawWidth = 3
'Printer.Line -Step(Printer.Width, 0)
Printer.Line -Step(vForm.picCRFPage.Width, 0)
Printer.DrawWidth = 2

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintLine", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub PrintComment(ByVal vForm As Form, _
                        ByVal vCRFElement As ADODB.Recordset, _
                        ByVal sgYPrint As Single)
'---------------------------------------------------------------------
Dim mFontName As Variant
Dim mFontSize As Variant
Dim mFontBold As Variant
Dim mFontItalic As Variant
Dim msCommentText As String
Dim msCommentline As String
Dim mnLineCount As Integer
Dim msPositionOfCR As Integer

On Error GoTo ErrHandler

'  Look for specific font or use the form's default one
GetCaptionFont vForm, vCRFElement, mFontName, mFontSize, mFontBold, mFontItalic

'  Set font attributes
On Error Resume Next
Printer.FontName = mFontName
Printer.Font.Charset = 1
On Error GoTo ErrHandler

Printer.FontSize = mFontSize * msgScaleFactorFonts
Printer.FontBold = mFontBold
Printer.FontItalic = mFontItalic

msCommentText = vCRFElement!Caption
mnLineCount = 0

'Do While msCommentText <> ""
While msCommentText <> ""
    msPositionOfCR = InStr(msCommentText, Chr(13))
    If msPositionOfCR <> 0 Then
        'strip off text before Cr/LF
        msCommentline = Mid(msCommentText, 1, msPositionOfCR - 1)
        msCommentText = Mid(msCommentText, msPositionOfCR + 2, Len(msCommentText))
    Else
        msCommentline = msCommentText
        msCommentText = ""
    End If
    'set page position for comment line
    Printer.CurrentX = vCRFElement!X
    Printer.CurrentY = sgYPrint + (Printer.TextHeight("X") * mnLineCount)
    mnLineCount = mnLineCount + 1
    'print the comment line
    Printer.Print msCommentline
Wend
'Loop

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintComment", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub PrintPicture(ByVal vForm As Form, _
                        ByVal vCRFElement As ADODB.Recordset, _
                        ByVal sgYPrint As Single)
'---------------------------------------------------------------------
' Pictures have to be loaded into a control before printing
'---------------------------------------------------------------------

Dim nHeightBeforeScaling As Single
Dim nWidthBeforeScaling As Single
Dim nHeightAfterScaling As Single
Dim nWidthAfterScaling As Single
Dim sPicFileName As String

On Error GoTo ErrHandler

    '   ATN 8/5/99
    '   Check that picture has been loaded successfully before printing
    On Error Resume Next
    ' Use gsDOCUMENTS_PATH - NCJ 1/10/99
    sPicFileName = gsDOCUMENTS_PATH & vCRFElement!Caption
    vForm.picUsedForPrinting.Picture = LoadPicture(sPicFileName)
    
    If Err.Number = 0 Then
        
        nHeightBeforeScaling = vForm.picUsedForPrinting.Picture.Height
        nWidthBeforeScaling = vForm.picUsedForPrinting.Picture.Width
        'MsgBox ("Before height=" & nHeightBeforeScaling & " width=" & nWidthBeforeScaling)
        nHeightAfterScaling = nHeightBeforeScaling * msgScaleFactorFonts
        nWidthAfterScaling = nWidthBeforeScaling * msgScaleFactorFonts
        'MsgBox ("After height=" & nHeightAfterScaling & " width=" & nWidthAfterScaling)
        Printer.PaintPicture vForm.picUsedForPrinting.Picture, vCRFElement!X, sgYPrint, _
            nWidthAfterScaling, nHeightAfterScaling, , , nWidthBeforeScaling, nHeightBeforeScaling
            
    End If

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintPicture", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Public Function PrintCRFForm(ByVal vForm As Form, _
                        Optional ByRef rPrinterError As Long, _
                        Optional ByVal vLocalIdentifier1 As String) As Boolean
'---------------------------------------------------------------------
'Mo Morris  11/9/98     SPR 426 (change made in Released and Developed versions)
'CancelError on CommonDialog1 enabled and tested for after the ShowPrinter call.
'Mo Morris  23/11/99
'Visit title added to printed forms in Macro_DM mode.
'Changed to handle long text strings in text boxes that require more
'than one line. The extra line means that everything else on the form
'has to be moved down. Note that comments that are very close to a control
'that has resulted in an extra line will not be moved down
' WillC 25/5/00 SR3493 Changed Routine to a function
' NCJ 23/10/00 SR3951 - Use gbShowPageBreaksOnly for showing page breaks rather than printing
'---------------------------------------------------------------------

Dim rsCRFElement As ADODB.Recordset
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
Dim sgTextOnSecondLineFudge As Single
Dim sgPrevTextOnSecondLineFudge As Single
Dim sgPrevYPrint As Single
Dim sMsg As String
Dim sCaption As String

    ' NCJ 23/10/00 - Clear page breaks if necessary
    If gbShowPageBreaksOnly Then
        'Commented out at request of buildlist
        'page breaks needed to display
        'ash 24/07/2001
        ' RS 12/06/2003: Re-activated, as clicking View Page Break twice would crash MACRO
        Call ClearPageBreaks(vForm)
    End If
    
    mnPageCount = 1
    rPrinterError = 0
    
    On Error Resume Next
    frmMenu.CommonDialog1.CancelError = True
    '   ATN 2/9/1999
    '   Don't prompt for show printer dialog
    'get user to open printer
    'frmMenu.CommonDialog1.ShowPrinter
    'check for errors in ShowPrinter (incuding a Cancel)
    If Err.Number > 0 Then Exit Function
    'restore normal error trapping
    On Error GoTo PrinterError
    
    'set printer scalemode to twips
    Printer.ScaleMode = vbTwips
    
    'calculate the scaling factor between the width of the form and the width of the
    'printed page (minus the margins - 1 inch on the left and 1/4 inch on the right).
    'This same scaling factor will be used to adjust the scaleHeight proportionally
    'Note that the form width is no longer a fixed width
    msgScaleFactor = (vForm.picCRFPage.Width / (Printer.ScaleWidth - 1440 - 360))
    
    'set scaleleft to incorporate a 1 inch margin (1440 twips)
    'set scaletop to incorporate a 1/4 inch margin (360 twips)
    Printer.ScaleLeft = -1440 * msgScaleFactor
    Printer.ScaleTop = -360 * msgScaleFactor
    Printer.ScaleWidth = Printer.ScaleWidth * msgScaleFactor
    Printer.ScaleHeight = Printer.ScaleHeight * msgScaleFactor
    msgTitleGap = -300 * msgScaleFactor
    
    'Print the forms header
    If Not gbShowPageBreaksOnly Then
        PrintCRFFormHeader vForm, vLocalIdentifier1
    End If
    
    'based on the form size and the selected paper size sgFormLength represents the maximum
    'Y value that can appear on the current page. All controls will be tested aginst this value
    'prior to being printed. When it is exceeded a new page will be created
    sgFormLength = Printer.ScaleHeight - (720 * msgScaleFactor)
    
    'calculate the reciprocal of the scaling factor for use on font sizes
    msgScaleFactorFonts = 1 / msgScaleFactor
    
    'Get CRF Elements and place each one onto the printed page
    Set rsCRFElement = New ADODB.Recordset
    'Note that gdsCRFPageDataItemsYorder has been changed and ignores Questions that are elements of a RQG
    Set rsCRFElement = gdsCRFPageDataItemsYorder(vForm.ClinicalTrialId, _
                            vForm.VersionId, vForm.CRFPageId)
    
    'initialise the accumlative page correction variable
    sgAccumlativeReduction = 0
    sgTextOnSecondLineFudge = 0
    While Not rsCRFElement.EOF
    
        If rsCRFElement!Hidden = False Then
            sCaption = RemoveNull(rsCRFElement!Caption)
            
        'Debug.Print "Type:" & rsCRFElement!ControlType & " Code:" & rsCRFElement!DataItemCode & " ID:" & rsCRFElement!CRFelementID & " Caption:[" & sCaption & "] PrintOrder:" & rsCRFElement!PrintOrder & " Y:" & rsCRFElement!Y & " CaptionY:" & rsCRFElement!CaptionY
            
            'make copies of the elements 'Y' and 'CaptionY' values, because they might need changing
            sgPrevYPrint = sgYPrint
            sgYPrint = rsCRFElement!Y
            sgCaptionYPrint = rsCRFElement!CaptionY
            'reduce 'Y' and 'CaptionY' if we are no longer dealing with the first page
            If mnPageCount > 1 Then
                sgYPrint = sgYPrint - sgAccumlativeReduction
                sgCaptionYPrint = sgCaptionYPrint - sgAccumlativeReduction
            End If
            'Text boxes that contain long strings that need to wrap over onto extra lines are
            'detected in PrintTextBox and cause an increase to the variable sgTextOnSecondLineFudge
            'Comments (i.e. ControlType=16386) that are very close to a control that might have
            'generated an increase to sgTextOnSecondLineFudge should not have sgTextOnSecondLineFudge
            'added to their Y-co-ordinates, the previous value of sgTextOnSecondLineFudge
            '(i.e. sgPrevTextOnSecondLineFudge) should be used.
            If rsCRFElement!ControlType = 16386 And (Abs(sgYPrint + sgPrevTextOnSecondLineFudge - sgPrevYPrint) < 150) Then
                sgYPrint = sgYPrint + sgPrevTextOnSecondLineFudge
            Else
                sgYPrint = sgYPrint + sgTextOnSecondLineFudge
                sgCaptionYPrint = sgCaptionYPrint + sgTextOnSecondLineFudge
            End If
            
            'assess length of current element to be printed
            Select Case rsCRFElement!ControlType
            Case 0
                'its a Repeating Question Group
                Call PrintRQGEstimate(vForm, rsCRFElement, sgElementLength)
            Case gnOPTION_BUTTONS, 258
                'OptionButtons, OptionBoxes
                Dim rsValues As ADODB.Recordset
                GetFont vForm, rsCRFElement, vFontName, vFontSize, vFontBold, vFontItalic
                On Error Resume Next
                Printer.FontName = vFontName
                Printer.Font.Charset = 1
                On Error GoTo PrinterError
                
                Printer.FontSize = vFontSize * msgScaleFactorFonts
                'calculate length of a single item
                sgElementLength = (Printer.TextHeight("X") + 100) * 1.5
                'get the number of values belonging to dataitem
                Set rsValues = New ADODB.Recordset
                Set rsValues = gdsDataValues(vForm.ClinicalTrialId, vForm.VersionId, _
                        rsCRFElement!DataItemId)
                rsValues.MoveLast
                'calculate the total length of the current optionBox/optionButton
                sgElementLength = sgElementLength * rsValues.RecordCount
            Case 16385
                'Line
                sgElementLength = 0
            Case 16386
                'Comment
                Dim msCommentText As String
                Dim mnPositionOfCR As Integer
                Dim mnLineCount As Integer
                msCommentText = sCaption
                GetFont vForm, rsCRFElement, vFontName, vFontSize, vFontBold, vFontItalic
                On Error Resume Next
                Printer.FontName = vFontName
                Printer.Font.Charset = 1
                On Error GoTo PrinterError
                
                Printer.FontSize = vFontSize * msgScaleFactorFonts
                mnLineCount = 1
                While msCommentText <> ""
                    mnPositionOfCR = InStr(msCommentText, Chr(13))
                    If mnPositionOfCR <> 0 Then
                        msCommentText = Mid(msCommentText, mnPositionOfCR + 2, Len(msCommentText))
                        mnLineCount = mnLineCount + 1
                    Else
                        msCommentText = ""
                    End If
                Wend
                sgElementLength = Printer.TextHeight("X") * mnLineCount
            Case 16388
                'Picture
                On Error Resume Next
                'Changed Mo Morris 10/2/00
                vForm.picUsedForPrinting.Picture = LoadPicture(gsDOCUMENTS_PATH & rsCRFElement!Caption)
                'vForm.picUsedForPrinting.Picture = LoadPicture(gsAppPath & "documents\" & rsCRFElement!Caption)
                If Err.Number = 0 Then
                    sgElementLength = vForm.picUsedForPrinting.Picture.Height * msgScaleFactorFonts
                End If
                On Error GoTo PrinterError
            Case Else
                'TextBox
                GetFont vForm, rsCRFElement, vFontName, vFontSize, vFontBold, vFontItalic
                On Error Resume Next
                Printer.FontName = vFontName
                Printer.Font.Charset = 1
                On Error GoTo PrinterError
                
                Printer.FontSize = vFontSize * msgScaleFactorFonts

            End Select
            
            'assess caption length
            GetCaptionFont vForm, rsCRFElement, vFontName, vFontSize, vFontBold, vFontItalic
            On Error Resume Next
            Printer.FontName = vFontName
            Printer.Font.Charset = 1
            On Error GoTo PrinterError
            Printer.FontSize = vFontSize * msgScaleFactorFonts
            sgCaptionLength = Printer.TextHeight("X") + 100
            
            'check that the current control will fit onto the page
            'changed by Mo Morris 14/10/99
            'Test for a non blank caption before deciding that it should be part of
            'the decision to terminate the current page. This was done so that blank captions
            'with a CaptionY co-ordinate that is miles away from its control don't effect the printing
            If (sCaption > "" And ((sgYPrint + sgElementLength > sgFormLength) Or (sgCaptionYPrint + sgCaptionLength > sgFormLength))) _
                Or (sCaption = "" And (sgYPrint + sgElementLength > sgFormLength)) Then
                'print the page that has just become full
                If Not gbShowPageBreaksOnly Then
                    Printer.EndDoc
                End If
                'increase the accumlative page correction variable by the smaller
                'of sgYPrint and sgCaptionYPrint (if Captio is non-blank)
                'changed by Mo Morris 15/10/99
                'changed by Mo Morris 21/10/99 test for a non-comment before allowing sgCaptionYPrint
                'to be part of the calculation of sgAccumlativeReduction. (i.e comments have a CaptionY
                'of zero which can become a large negative figure for pages 2 and onwards and would
                'always be less than sgYPrint
                'Mo 26/10/2007 Bug 2957
                If sgCaptionYPrint < sgYPrint And sCaption > "" And (rsCRFElement!ControlType < gnVISUAL_ELEMENT) Then
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
                        
                ' NCJ 23/10/00 - Show the page break
                If gbShowPageBreaksOnly Then
                    Call BuildPageBreak(vForm, mnPageCount, sgAccumlativeReduction)
                End If
                
                'increment the page count
                mnPageCount = mnPageCount + 1
                
                If Not gbShowPageBreaksOnly Then
                    'Print the forms header
                    PrintCRFFormHeader vForm, vLocalIdentifier1
                End If
            End If
            'put the control onto the page
            If Not gbShowPageBreaksOnly Then
                PrintCRFElement vForm, rsCRFElement, sgYPrint, sgCaptionYPrint
            End If
            ' get next record
        
        End If      ' If element not hidden
        
        rsCRFElement.MoveNext
    Wend
    
    rsCRFElement.Close
    Set rsCRFElement = Nothing
    
    'Called this routine to displaygrid when printing
    'Ash 24/07/2001
    Call vForm.DisplayGrid

    'print the last page
    If Not gbShowPageBreaksOnly Then
        Printer.EndDoc
    Else
        ' Tell them how many pages
        sMsg = "This eForm will print on "
        If mnPageCount = 1 Then
            ' Tell them there aren't any page breaks
            sMsg = sMsg & "one page"
        Else
            sMsg = sMsg & mnPageCount & " pages"
        End If
         
        DialogInformation sMsg
    End If
    
    ' WillC 25/5/00 SR3493 Changed Routine to a function set it to true signifying there was no error
    PrintCRFForm = False
    
    ' NCJ 23/10/00 - Reset to default value
   gbShowPageBreaksOnly = False

Exit Function

PrinterError:
    MsgBox "A printer error has occurred.  The error number is " & Err.Number & vbCrLf _
         & Err.Description, vbOKOnly + vbInformation
         ' WillC 25/5/00 SR3493 Changed Routine to a function set it to false here signifying there was an error
         PrintCRFForm = False
        ' NCJ 23/10/00 - Reset to default value
        gbShowPageBreaksOnly = False

End Function

'---------------------------------------------------------------------
Private Sub PrintCRFFormHeader(ByVal vForm As Form, Optional ByVal vLocalIdentifier1 As String)
'---------------------------------------------------------------------
'sub added 23/11/99 by Mo Morris
'format and print a title in the top corner of printed page and draw
'the forms corners/margins
'---------------------------------------------------------------------
'Mo Morris  28/1/00 Changed 'Trial' to 'Study' throughout
'Mo Morris  17/3/00 SR2211, changed 'PatientId' to 'Subject' and 'Study Subject Label' to 'Label'
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTitle As ADODB.Recordset
Dim msTitle As String

    On Error GoTo ErrHandler

    sSQL = "SELECT CRFTitle FROM CRFPage" _
            & " WHERE ClinicalTrialId = " & vForm.ClinicalTrialId _
            & " AND CRFPageId = " & vForm.CRFPageId
    Set rsTitle = New ADODB.Recordset
    rsTitle.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Printer.CurrentX = 0
    Printer.CurrentY = msgTitleGap
    
        msTitle = "Study: " & vForm.ClinicalTrialName & " (Version: " & vForm.VersionId & ")" _
            & "    Form: " & rsTitle!CRFTitle _
            & vbTab & vbTab & "Page " & mnPageCount _
            & vbTab & vbTab & "Printed " & Format(Now, "yyyy/mm/dd hh:mm:ss")

    Printer.FontSize = 6
    Printer.FontBold = False
    Printer.FontItalic = False
    Printer.Print msTitle
    
    'draw border corners on the page, incorporating a 1/2 inch margin at the bottom (720 twips)
    Printer.DrawWidth = 1
    Printer.Line (0, 0)-Step(400, 0)
    Printer.Line (vForm.picCRFPage.Width - 400, 0)-Step(400, 0)
    Printer.Line -Step(0, 400)
    Printer.Line (vForm.picCRFPage.Width, (Printer.ScaleHeight - (720 * msgScaleFactor)) - 400)-Step(0, 400)
    Printer.Line -Step(-400, 0)
    Printer.Line (400, (Printer.ScaleHeight - (720 * msgScaleFactor)))-Step(-400, 0)
    Printer.Line (0, (Printer.ScaleHeight - (720 * msgScaleFactor)))-Step(0, -400)
    Printer.Line (0, 0)-Step(0, 400)
    
    rsTitle.Close
    Set rsTitle = Nothing

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintCRFFormHeader", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub PrintAttachment(ByVal vForm As Form, _
                            ByVal vCRFElement As ADODB.Recordset, _
                            ByVal sgYPrint As Single, _
                            Optional ByVal sgXPrint As Single = 0, _
                            Optional ByRef sgControlWidth As Single, _
                            Optional ByRef sgControlHeight As Single)
'---------------------------------------------------------------------
Dim sgBoxWidth As Single
Dim sgBoxHeight As Single
Dim sgElementX As Single
Dim sgSpaceForIcons As Single

    On Error GoTo ErrHandler

    If sgXPrint <> 0 Then
        sgElementX = sgXPrint
    Else
        sgElementX = vCRFElement!X
    End If
    
    'The font attributes of an attachment button are hard coded to "MS Sans Serif", size=8.5, Bold & Italic = False
    On Error Resume Next
    Printer.FontName = "MS Sans Serif"
    Printer.Font.Charset = 1
    On Error GoTo ErrHandler
    Printer.FontSize = 8.5 * msgScaleFactorFonts
    Printer.FontBold = False
    Printer.FontItalic = False

    Printer.CurrentY = sgYPrint
    Printer.CurrentX = sgElementX

    'Check for Attachment requiring space for status icons
    sgSpaceForIcons = 0
    If vCRFElement!ShowStatusFlag = 1 Then
        sgSpaceForIcons = mn_SPACE_FOR_STATUS_ICON
    End If

    'Set the Scaled size of the attachment button
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
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintAttachment", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub PrintHotLink(ByVal vForm As Form, _
                        ByVal vCRFElement As ADODB.Recordset, _
                        ByVal sgYPrint As Single)
'---------------------------------------------------------------------
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant

    On Error GoTo ErrHandler

    'Look for specific font or use the form's default one
    Call GetCaptionFont(vForm, vCRFElement, vFontName, vFontSize, vFontBold, vFontItalic)

    'Set font attributes
    On Error Resume Next
    Printer.FontName = vFontName
    Printer.Font.Charset = 1
    On Error GoTo ErrHandler

    Printer.FontSize = vFontSize * msgScaleFactorFonts
    Printer.FontBold = vFontBold
    Printer.FontItalic = vFontItalic

    Printer.CurrentY = sgYPrint
    Printer.CurrentX = vCRFElement!X
    'Turn on Underlined text
    Printer.FontUnderline = True
    'Print the hotlink
    Printer.Print vCRFElement!Caption
    'Turn off Underlined text
    Printer.FontUnderline = False

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintHotLink", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub


'---------------------------------------------------------------------
Public Sub GetCaptionFont(ByVal vForm As Form, _
                        ByVal vCRFElement As ADODB.Recordset, _
                        ByRef rFontName As Variant, _
                        ByRef rFontSize As Variant, _
                        ByRef rFontBold As Variant, _
                        ByRef rFontItalic As Variant)
'---------------------------------------------------------------------
'Get the default font for the form if the element's caption doesn't have
'specified font settings
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If IsNull(vCRFElement!CaptionFontName) Or vCRFElement!CaptionFontName = "" Then
        rFontName = vForm.DefaultFontName
        rFontSize = vForm.DefaultFontSize
        rFontBold = vForm.DefaultFontBold
        rFontItalic = vForm.DefaultFontItalic
    Else
        rFontName = vCRFElement!CaptionFontName
        rFontSize = vCRFElement!CaptionFontSize
        rFontBold = vCRFElement!CaptionFontBold
        rFontItalic = vCRFElement!CaptionFontItalic
    End If

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetCaptionFont", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub PrintRQG(ByVal vForm As Form, _
                    ByVal vCRFElement As ADODB.Recordset, _
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
Dim nNumDisplayRows As Integer
Dim nNumInitialRows As Integer
Dim nRow As Integer
Dim sgHeaderWidth As Single
Dim sgElementHeaderHeight As Single
Dim sgHeaderHeight As Single
Dim sgControlWidth As Single
Dim sgControlHeight As Single
Dim sgThisRowHeight As Single

Dim rsRQGElements As ADODB.Recordset
Dim rsRQGDetails As ADODB.Recordset
Dim rsElementDetails As ADODB.Recordset
Dim sgHeaderCheckWidth As Single
Dim nNumLinesInCaption As Integer
Dim sCaption As String
Dim sPartOfCaption As String
Dim nPositionOfCR As Integer
Dim nNumExtraLinesCaption As Integer
Dim bCRExistsInCaption As Boolean
Dim nDisplayBorder As Integer

Const sglGAP As Single = 50

    On Error GoTo ErrHandler
    
    'Store co-ordinates of RQGFrame's top/left corner
    sgRQGFrameTop = sgYPrint
    sgRQGFrameLeft = vCRFElement!X
    
    'get a recordset of the RQG elements
    Set rsRQGElements = New ADODB.Recordset
    Set rsRQGElements = QGroupQuestionList(vForm.ClinicalTrialId, vForm.VersionId, vCRFElement!QGroupID)
    
    'get the Minimum number of rows that are displayed on screen
    Set rsRQGDetails = New ADODB.Recordset
    Set rsRQGDetails = EFormQGroup(vForm.ClinicalTrialId, vForm.VersionId, vCRFElement!QGroupID, vForm.CRFPageId)
    nNumDisplayRows = rsRQGDetails!DisplayRows
    nNumInitialRows = rsRQGDetails!InitialRows
    'Mo 14/4/2004 Bug 2063
    nDisplayBorder = rsRQGDetails!Border
    rsRQGDetails.Close
    Set rsRQGDetails = Nothing
    
    'Assess the height of the header
    sgHeaderHeight = 0
    rsRQGElements.MoveFirst
    Do Until rsRQGElements.EOF
        Set rsElementDetails = New ADODB.Recordset
        Set rsElementDetails = RQGElementById(vForm.ClinicalTrialId, vForm.VersionId, vForm.CRFPageId, rsRQGElements!DataItemId)
        Call GetCaptionFont(vForm, rsElementDetails, vFontName, vFontSize, vFontBold, vFontItalic)
        On Error Resume Next
        Printer.FontName = vFontName
        Printer.Font.Charset = 1
        On Error GoTo ErrHandler
        Printer.FontSize = vFontSize * msgScaleFactorFonts
        Printer.FontBold = vFontBold
        Printer.FontItalic = vFontItalic
        'Mo 14/4/2004 Bug 2063
        nNumLinesInCaption = NumLinesInCaption(rsElementDetails!Caption)
        sgElementHeaderHeight = (nNumLinesInCaption * Printer.TextHeight("X")) + sglGAP
        If sgElementHeaderHeight > sgHeaderHeight Then
            sgHeaderHeight = sgElementHeaderHeight
        End If
        rsRQGElements.MoveNext
    Loop
    
    'initialize sgRQGFrameWidth & sgRQGControlsHeight to msglGAP a 50 twip border/gap
    'that is placed between everything in the RQG
    sgRQGControlsHeight = sglGAP
    For nRow = 1 To nNumInitialRows
        sgRQGFrameWidth = sglGAP
        sgThisRowHeight = 0
        rsRQGElements.MoveFirst
        Do Until rsRQGElements.EOF
            sgControlWidth = 0
            sgControlHeight = 0
            Set rsElementDetails = New ADODB.Recordset
            Set rsElementDetails = RQGElementById(vForm.ClinicalTrialId, vForm.VersionId, vForm.CRFPageId, rsRQGElements!DataItemId)
            'assess width of RQG element Header
            Call GetCaptionFont(vForm, rsElementDetails, vFontName, vFontSize, vFontBold, vFontItalic)
            On Error Resume Next
            Printer.FontName = vFontName
            Printer.Font.Charset = 1
            On Error GoTo ErrHandler
            Printer.FontSize = vFontSize * msgScaleFactorFonts
            Printer.FontBold = vFontBold
            Printer.FontItalic = vFontItalic
            sgHeaderWidth = Printer.TextWidth(RemoveNull(rsElementDetails!Caption))
            
            'Mo 3/6/03 Bug 1819
            'Check that there is enough width to print the this columns control.
            'This check prevents the printing of RQG controls that on the screen you have to scroll to view.
            'If the caption is blank or very short the following lines force a minimum column header width of 800 twips
            If sgHeaderWidth < 800 Then
                sgHeaderCheckWidth = 800
            Else
                sgHeaderCheckWidth = sgHeaderWidth
            End If
            If vForm.picCRFPage.Width - (sgRQGFrameLeft + sgRQGFrameWidth + sgHeaderCheckWidth) < 0 Then
                Exit Do
            End If
            
            'if its row 1 print the RQG element Header
            If nRow = 1 Then
                'Mo 14/4/2004 Bug 2063
                'checking for CR/LFs in the caption
                sCaption = RemoveNull(rsElementDetails!Caption)
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
            Select Case rsElementDetails!ControlType
            Case gn_ATTACHMENT
                Call PrintAttachment(vForm, rsElementDetails, sgRQGFrameTop + sgHeaderHeight + sgRQGControlsHeight, sgRQGFrameLeft + sgRQGFrameWidth, sgControlWidth, sgControlHeight)
            Case gn_OPTION_BUTTONS
                Call PrintOptionButtons(vForm, rsElementDetails, sgRQGFrameTop + sgHeaderHeight + sgRQGControlsHeight, sgRQGFrameLeft + sgRQGFrameWidth, sgControlWidth, sgControlHeight)
            Case gn_PUSH_BUTTONS
                Call PrintOptionBoxes(vForm, rsElementDetails, sgRQGFrameTop + sgHeaderHeight + sgRQGControlsHeight, sgRQGFrameLeft + sgRQGFrameWidth, sgControlWidth, sgControlHeight)
            Case Else
                'TextBox or dropdown list or calendar
                Call PrintTextBox(vForm, rsElementDetails, sgRQGFrameTop + sgHeaderHeight + sgRQGControlsHeight, sgRQGFrameLeft + sgRQGFrameWidth, sgControlWidth, sgControlHeight)
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

            rsRQGElements.MoveNext
        Loop
        'increment the y position
        sgRQGControlsHeight = sgRQGControlsHeight + sgThisRowHeight + sglGAP
    Next nRow
    
    'Add the Header height to the height of the Controls to give a FrameHeight
    sgRQGFrameHeight = sgHeaderHeight + sgRQGControlsHeight
    If nNumDisplayRows - nNumInitialRows > 0 Then
        sgRQGFrameHeight = sgRQGFrameHeight + ((nNumDisplayRows - nNumInitialRows) * (sgThisRowHeight + sglGAP))
    End If
    
    'Draw frame around RQG
    'Mo 14/4/2004 Bug 2063
    If nDisplayBorder = 1 Then
        Printer.CurrentX = sgRQGFrameLeft
        Printer.CurrentY = sgRQGFrameTop
        Printer.Line -Step(sgRQGFrameWidth, 0)
        Printer.Line -Step(0, sgRQGFrameHeight)
        Printer.Line -Step(-sgRQGFrameWidth, 0)
        Printer.Line -Step(0, -sgRQGFrameHeight)
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintRQG", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub PrintRQGEstimate(ByVal vForm As Form, _
                            ByVal vCRFElement As ADODB.Recordset, _
                            ByRef sgElementLength As Single)
'---------------------------------------------------------------------
'This sub will work through a Repeating Question Group question and
'calculate its length (sgElementLength)
'This sub mimics the printing activities of PrintRQG.
'---------------------------------------------------------------------
Dim nNumDisplayRows As Integer
Dim vFontName As Variant
Dim vFontSize As Variant
Dim vFontBold As Variant
Dim vFontItalic As Variant
Dim sgElementHeaderHeight As Single
Dim sgHeaderHeight As Single
Dim sgRowHeight As Single
Dim sgControlHeight As Single
Dim rsRQGElements As ADODB.Recordset
Dim rsRQGDetails As ADODB.Recordset
Dim rsElementDetails As ADODB.Recordset
Dim rsRQGElementCatCodes As ADODB.Recordset

Const sglGAP As Single = 50

    On Error GoTo ErrHandler

    'get a recordset of the RQG elements
    Set rsRQGElements = New ADODB.Recordset
    Set rsRQGElements = QGroupQuestionList(vForm.ClinicalTrialId, vForm.VersionId, vCRFElement!QGroupID)
    
    'get the number of rows that are displayed on screen
    Set rsRQGDetails = New ADODB.Recordset
    Set rsRQGDetails = EFormQGroup(vForm.ClinicalTrialId, vForm.VersionId, vCRFElement!QGroupID, vForm.CRFPageId)
    nNumDisplayRows = rsRQGDetails!DisplayRows
    rsRQGDetails.Close
    Set rsRQGDetails = Nothing
    
    'Start off with the height of the header
    sgHeaderHeight = 0
    rsRQGElements.MoveFirst
    Do Until rsRQGElements.EOF
        Set rsElementDetails = New ADODB.Recordset
        Set rsElementDetails = RQGElementById(vForm.ClinicalTrialId, vForm.VersionId, vForm.CRFPageId, rsRQGElements!DataItemId)
        Call GetCaptionFont(vForm, rsElementDetails, vFontName, vFontSize, vFontBold, vFontItalic)
        On Error Resume Next
        Printer.FontName = vFontName
        Printer.Font.Charset = 1
        On Error GoTo ErrHandler
        Printer.FontSize = vFontSize * msgScaleFactorFonts
        Printer.FontBold = vFontBold
        Printer.FontItalic = vFontItalic
        sgElementHeaderHeight = sglGAP + Printer.TextHeight("X") + sglGAP
        If sgElementHeaderHeight > sgHeaderHeight Then
            sgHeaderHeight = sgElementHeaderHeight
        End If
        rsRQGElements.MoveNext
    Loop
    
    'Assess the normal row height by inspecting each control element within a single row
    sgRowHeight = 0
    rsRQGElements.MoveFirst
    Do Until rsRQGElements.EOF
        Set rsElementDetails = New ADODB.Recordset
        Set rsElementDetails = RQGElementById(vForm.ClinicalTrialId, vForm.VersionId, vForm.CRFPageId, rsRQGElements!DataItemId)
        Call GetFont(vForm, rsElementDetails, vFontName, vFontSize, vFontBold, vFontItalic)
        On Error Resume Next
        Printer.FontName = vFontName
        Printer.Font.Charset = 1
        On Error GoTo ErrHandler
        Printer.FontSize = vFontSize * msgScaleFactorFonts
        Printer.FontBold = vFontBold
        Printer.FontItalic = vFontItalic
        Select Case rsElementDetails!ControlType
        Case gn_ATTACHMENT
            sgControlHeight = 375
        Case gn_OPTION_BUTTONS, gn_PUSH_BUTTONS
            Set rsRQGElementCatCodes = New ADODB.Recordset
            Set rsRQGElementCatCodes = RQGElementById(vForm.ClinicalTrialId, vForm.VersionId, vForm.CRFPageId, rsRQGElements!DataItemId)
            sgControlHeight = (Printer.TextHeight("X") + 100) * rsRQGElementCatCodes.RecordCount
            rsRQGElementCatCodes.Close
            Set rsRQGElementCatCodes = Nothing
        Case Else
            'TextBox or dropdown list or calendar
            sgControlHeight = Printer.TextHeight("X") + 100
        End Select
        If sgControlHeight > sgRowHeight Then
            sgRowHeight = sgControlHeight
        End If
        rsRQGElements.MoveNext
    Loop
    
    sgElementLength = sgHeaderHeight + (nNumDisplayRows * (sgRowHeight + sglGAP))
    
    rsRQGElements.Close
    Set rsRQGElements = Nothing
    rsElementDetails.Close
    Set rsElementDetails = Nothing
    
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PrintRQGEstimate", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Private Function RQGElementById(ByVal lClinicalTrialId As Long, _
                                ByVal lVersionId As Long, _
                                ByVal lCRFPageId As Long, _
                                ByVal lDataItemId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
' ic 04/07/2005 added clinical coding
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "SELECT CRFElement.*, DataItem.DataItemLength, DataItem.DataType from CRFElement, DataItem " _
        & " WHERE CRFElement.ClinicalTrialId = DataItem.ClinicalTrialId " _
        & " AND CRFElement.VersionId = DataItem.VersionId " _
        & " AND CRFElement.DataItemId = DataItem.DataItemId " _
        & " AND CRFElement.ClinicalTrialId = " & lClinicalTrialId _
        & " AND CRFElement.VersionId = " & lVersionId _
        & " AND CRFElement.CRFPageId = " & lCRFPageId _
        & " AND CRFElement.DataItemId = " & lDataItemId
    
    Set RQGElementById = New ADODB.Recordset
    RQGElementById.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RQGElementById", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Private Function RQGElementCatCodes(ByVal lClinicalTrialId As Long, _
                                    ByVal lVersionId As Long, _
                                    ByVal lDataItemId As Long) As ADODB.Recordset
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM ValueData " _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND VersionId = " & lVersionId _
        & " AND DataItemId = " & lDataItemId _
        & " ORDER By ValueOrder"

    Set RQGElementCatCodes = New ADODB.Recordset
    RQGElementCatCodes.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "RQGElementCatCodes", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

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
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "NumLinesInCaption", "BuildCRF")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function
