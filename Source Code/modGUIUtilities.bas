Attribute VB_Name = "modGUIUtilities"
'----------------------------------------------------------------------------------------'
'   Copyright:  Inferfrmformd Ltd. 2000. All Rights Reserved
'   File:       modGUIUtilities.bas
'   Author:     Toby Aldridge, April 2000
'   Purpose:    GUI Utilities for MACRO
'
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'TA 18/10/2000: Routines to store form dimensions
'TA 06/06/2002: Minor changes to the storing of form dimensions CBB 2.2.14.27
'----------------------------------------------------------------------------------------'

Option Explicit

'variables for mouse pointer and status message stack
Private mnMousePointer As Variant
Private mSStatus As Variant

'background colour of a control that contains invalid data
Public Const g_INVALID_BACKCOLOUR = vbYellow


'---------------------------------------------------------------------
Public Sub MousePointerChange(nMousePointer As Integer, Optional sStatus As String = "")
'---------------------------------------------------------------------
'   TA 14/04/2000
'   Change mouse pointer and status message and add these to the stack
'---------------------------------------------------------------------

Dim lUBound As Long

    On Error GoTo ErrIgnore
    
    If Not IsArray(mnMousePointer) Then
        ReDim mnMousePointer(0) As Long
        ReDim mSStatus(0) As String
        mnMousePointer(0) = Screen.MousePointer
        mSStatus(0) = ""
    End If
    lUBound = UBound(mnMousePointer)
    ReDim Preserve mnMousePointer(lUBound + 1)
    ReDim Preserve mSStatus(lUBound + 1)
    mnMousePointer(lUBound + 1) = nMousePointer
    mSStatus(lUBound + 1) = sStatus
    Screen.MousePointer = nMousePointer
    
    Exit Sub
    
ErrIgnore:
    
End Sub



'---------------------------------------------------------------------
Public Sub MousePointerRestore()
'---------------------------------------------------------------------
'   TA 14/04/2000
'   Restore previous mouse pointer and status message
'---------------------------------------------------------------------

Dim lUBound As Long

    On Error GoTo ErrIgnore

    lUBound = UBound(mnMousePointer)
    ReDim Preserve mnMousePointer(lUBound - 1)
    ReDim Preserve mSStatus(lUBound - 1)
    Screen.MousePointer = mnMousePointer(lUBound - 1)
    
    Exit Sub
    
ErrIgnore:
    
End Sub

'---------------------------------------------------------------------
Public Sub HourglassOn(Optional sStatus As String = "Busy...")
'---------------------------------------------------------------------
'   TA 14/04/2000
'   Turn hourglass on
'---------------------------------------------------------------------

    MousePointerChange vbHourglass, sStatus
    
End Sub


'---------------------------------------------------------------------
Public Sub HourglassOff()
'---------------------------------------------------------------------
'   TA 14/04/2000
'   turn hourglass off
'---------------------------------------------------------------------

    MousePointerRestore
    
End Sub


'---------------------------------------------------------------------
Public Sub HourglassSuspend()
'---------------------------------------------------------------------
'   TA 14/04/2000
'   Change pointer to vbdefault
'   replaced by DefaultPointerOn
'---------------------------------------------------------------------

    MousePointerChange vbDefault
    
End Sub


'---------------------------------------------------------------------
Public Sub HourglassResume()
'---------------------------------------------------------------------
'   TA 14/04/2000
'   return to hourglass if appropriate
'   replaced by DefaultPointerOff
'---------------------------------------------------------------------

    MousePointerRestore
    
End Sub

'---------------------------------------------------------------------
Public Sub DefaultPointerOn()
'---------------------------------------------------------------------
'   TA 14/05/2000
'   Change pointer to vbdefault

'---------------------------------------------------------------------

    MousePointerChange vbDefault
    
End Sub


'---------------------------------------------------------------------
Public Sub DefaultPointerOff()
'---------------------------------------------------------------------
'   TA 14/05/2000
'   return to hourglass if appropriate
'---------------------------------------------------------------------

    MousePointerRestore
    
End Sub

'---------------------------------------------------------------------
Public Function PointerState() As Integer
'---------------------------------------------------------------------
'   TA 30/10/2000: return saved mousepointer state
'   'returns vbDefault if no pointer status
'---------------------------------------------------------------------

    On Error GoTo ErrIgnore

    PointerState = mnMousePointer(UBound(mnMousePointer))
    
    Exit Function

ErrIgnore:

    PointerState = vbDefault
    
End Function


'----------------------------------------------------------------------------------------'
Public Sub DialogError(sPrompt, Optional sTitle = "")
'----------------------------------------------------------------------------------------'
'display msgbox with a cross
'----------------------------------------------------------------------------------------'

    If sTitle = "" Then
        'no title - use app's
        sTitle = GetApplicationTitle
    End If
    
    MsgBox sPrompt, vbOKOnly + vbCritical, sTitle

End Sub

'----------------------------------------------------------------------------------------'
Public Sub DialogInformation(sPrompt, Optional sTitle = "")
'----------------------------------------------------------------------------------------'
'display msgbox with a information icon
'----------------------------------------------------------------------------------------'

    If sTitle = "" Then
        'no title- use app's
        sTitle = GetApplicationTitle
    End If
    
    MsgBox sPrompt, vbOKOnly + vbInformation, sTitle

End Sub

'----------------------------------------------------------------------------------------'
Public Function DialogWarning(sPrompt, Optional sTitle = "", Optional bCancel As Boolean = False) As Integer
'----------------------------------------------------------------------------------------'
'display a msgbox with an exclamation mark
' bCancel - allow cancel?
'----------------------------------------------------------------------------------------'

    If sTitle = "" Then
        'no title- use app's
        sTitle = GetApplicationTitle
    End If
    
    If bCancel Then
        DialogWarning = MsgBox(sPrompt, vbOKCancel + vbExclamation, sTitle)
    Else
        DialogWarning = MsgBox(sPrompt, vbOKOnly + vbExclamation, sTitle)
    End If

End Function

'----------------------------------------------------------------------------------------'
Public Function DialogQuestion(sPrompt As String, Optional sTitle As String = "", _
    Optional bCancel As Boolean = False, Optional lOptions As Long = 0) As Integer
'----------------------------------------------------------------------------------------'
'display a question msgbox
' MLM 14/02/03: Added lOptions argument in order that the default button can be specified,
'               but could be used for other MesgBox options too.
'----------------------------------------------------------------------------------------'

    If sTitle = "" Then
        'no title- use products
        sTitle = GetApplicationTitle
    End If
    
    If bCancel Then
        DialogQuestion = MsgBox(sPrompt, vbYesNoCancel + vbQuestion + lOptions, sTitle)
    Else
        DialogQuestion = MsgBox(sPrompt, vbYesNo + vbQuestion + lOptions, sTitle)
    End If
    
End Function


''----------------------------------------------------------------------------------------'
'Public Function GetApplicationTitle() As String
''----------------------------------------------------------------------------------------'
''Return the default title of an app
''----------------------------------------------------------------------------------------'
'
'    Select Case App.Title
'    Case "MACRO_SD"
'        If LCase$(Command) = "library" Then
'            GetApplicationTitle = "MACRO Library Management"
'        Else
'            GetApplicationTitle = "MACRO Study Definition"
'        End If
'    Case "MACRO_DM"
'         If LCase$(Command) = "review" Then
'            GetApplicationTitle = "MACRO Data Review"
'        Else
'            'TA 23/10/2000 changend from Data Management"
'            GetApplicationTitle = "MACRO Data Entry"
'        End If
'    Case "MACRO_EX"
'         GetApplicationTitle = "MACRO Exchange"
'    Case "MACRO_SM"
'        GetApplicationTitle = "MACRO System Management"
'    Case Else
'        GetApplicationTitle = "MACRO"
'    End Select
'
'End Function

'----------------------------------------------------------------------------------------'
Public Sub ColourControl(conControl As Control, bValid As Boolean)
'----------------------------------------------------------------------------------------'
' if not valid colour control yellow
' if valid colour control with windowsbackground colour
'----------------------------------------------------------------------------------------'

    If bValid Then
        conControl.BackColor = vbWindowBackground
    Else
        conControl.BackColor = g_INVALID_BACKCOLOUR
    End If

End Sub


'----------------------------------------------------------------------------------------'
Public Function ListItembyTag(lvwlistview As MSComctlLib.ListView, sTag As String) As MSComctlLib.ListItem
'----------------------------------------------------------------------------------------'
'return a listitem by its tag
'----------------------------------------------------------------------------------------'

Dim olstItem As MSComctlLib.ListItem
    For Each olstItem In lvwlistview.ListItems
        If olstItem.Tag = sTag Then
            Set ListItembyTag = olstItem
            Exit Function
        End If
    Next
End Function

'----------------------------------------------------------------------------------------'
Public Sub HighlightListItembyTag(lvwlistview As MSComctlLib.ListView, sTag As String)
'----------------------------------------------------------------------------------------'
'selects and makes bold the listitem with sTag tag
'undoes the bold on all otheres
'----------------------------------------------------------------------------------------'

Dim olistItem As MSComctlLib.ListItem
Dim i As Integer
    For Each olistItem In lvwlistview.ListItems
        With olistItem
            If .Tag = sTag Then
                .Bold = True
                'TA 15/12/2000: changed to count rather than count - 1
                For i = 1 To .ListSubItems.Count
                    .ListSubItems(i).Bold = True
                Next
                .Selected = True
                .EnsureVisible
            Else
                .Bold = False
                For i = 1 To .ListSubItems.Count
                    .ListSubItems(i).Bold = False
                Next
            End If
        End With
    Next
   
    lvwlistview.Refresh

End Sub

'----------------------------------------------------------------------------------------'
Public Sub ListViewSort(lvwlistview As MSComctlLib.ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'----------------------------------------------------------------------------------------'
' sort listview according to column click
'----------------------------------------------------------------------------------------'
    If Not lvwlistview.Sorted Then
        lvwlistview.Sorted = True
    End If
    
    If lvwlistview.SortKey = ColumnHeader.Index - 1 Then
        If lvwlistview.SortOrder = lvwDescending Then
            lvwlistview.SortOrder = lvwAscending
        Else
            lvwlistview.SortOrder = lvwDescending
        End If
    Else
        lvwlistview.SortKey = ColumnHeader.Index - 1
        lvwlistview.SortOrder = lvwAscending
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub SaveFormDimensions(oForm As Form)
'----------------------------------------------------------------------------------------'
' Saves a forms's dimension to the registry
'   must go  in QueryUnload event of form
'----------------------------------------------------------------------------------------'
 Dim sSetting As String
 Dim nWindowState As Integer
    
     'jump out if error
    On Error GoTo ErrHandler
    
    With oForm
        nWindowState = .WindowState
        If nWindowState = vbMinimized Then
            'if minimised do not save this fact
            Exit Sub
        End If
        
        'TA 06/06/2002: Minor changes to the storing of form dimensions CBB 2.2.14.27
        'if it is maximised then wipe the setting so that restore works
        If nWindowState = vbMaximized Then
            sSetting = ""
        Else
            'za 22/05/2002 - changed nWindowstate to windowstate
            sSetting = .Top & "," & .Left & "," & .Height & "," & .Width & "," & .WindowState
        End If
        
        Call SaveSetting(GetApplicationTitle, "Form Dimensions", Mid(oForm.Name, 4), sSetting)
    
    End With
    
Exit Sub
ErrHandler:

   
End Sub

'----------------------------------------------------------------------------------------'
Public Sub SetFormDimensions(oForm As Form)
'----------------------------------------------------------------------------------------'
' set a form's dimensions according to registry settings
'   call before using (form).show vbModal
'----------------------------------------------------------------------------------------'
Dim sSetting As String

    'jump out if error
    On Error GoTo ErrHandler

    
    sSetting = GetSetting(GetApplicationTitle, "Form Dimensions", Mid(oForm.Name, 4), "")

    If sSetting = "" Then
        'not found
        Call FormCentre(oForm)
    Else
        With oForm
            .WindowState = Split(sSetting, ",")(4)
            .Top = Split(sSetting, ",")(0)
            .Left = Split(sSetting, ",")(1)
            .Height = Split(sSetting, ",")(2)
            .Width = Split(sSetting, ",")(3)
        End With
    End If
    
Exit Sub
ErrHandler:

    
End Sub

'----------------------------------------------------------------------------------------'
Public Function CMDialogOpen(dlgDialog As CommonDialog, sTitle As String, ByRef sFile As String, Optional sFilter As String = "") As Boolean
'----------------------------------------------------------------------------------------'
' display browse dialog and return a file
' sFilter in form: "Text (*.txt)|*.txt|Pictures (*.bmp;*.ico)|*.bmp;*.ico"
'----------------------------------------------------------------------------------------'
    On Error GoTo ErrHandler
    With dlgDialog
        .CancelError = True
        .DialogTitle = sTitle
        .FileName = sFile
        .Filter = sFilter
        .ShowOpen
        sFile = .FileName
    End With
    CMDialogOpen = True
    
    Exit Function

ErrHandler:
    CMDialogOpen = False
    Exit Function
    
End Function

''----------------------------------------------------------------------------------------'
'Public Function CMDialogSave(dlgDialog As CommonDialog, sTitle As String, sFile As String, Optional sDefaultExt As String = "") As Boolean
''----------------------------------------------------------------------------------------'
'' not yet used, might not work
''----------------------------------------------------------------------------------------'
'    On Error GoTo ErrHandler
'    With dlgDialog
'        .CancelError = True
'        .DialogTitle = sTitle
'        .FileName = sFile
'        .ShowSave
'        sFile = .FileName
'    End With
'    CMDialogSave = True
'
'    Exit Function
'
'ErrHandler:
'    CMDialogSave = False
'    Exit Function
'
'End Function

'----------------------------------------------------------------------------------------
Public Function TabletoCombo(cboCombo As ComboBox, tblTable As clsDataTable)
'----------------------------------------------------------------------------------------
' puts a one table into a combo box
' if there are two cols then the second col is used for the item data
'----------------------------------------------------------------------------------------
Dim i As Long
Dim bAddItemData As Boolean

    bAddItemData = (tblTable.Cols = 2)
    cboCombo.Clear
    For i = 1 To tblTable.Rows
        cboCombo.AddItem tblTable(i, 1)
        If bAddItemData Then
            cboCombo.ItemData(cboCombo.NewIndex) = Val(tblTable(i, 2))
        End If
    Next

End Function

'----------------------------------------------------------------------------------------'
Public Function TableToListView(lvwlistview As MSComctlLib.ListView, tblTable As clsDataTable, _
                                    Optional bWidenForHeading As Boolean = True) As Long
'----------------------------------------------------------------------------------------'
' uses row number as the listitem tag
' if the table doesn't have a header record then nothing will happen
'----------------------------------------------------------------------------------------'

Dim vHeadings As Variant
Dim lFields As Long
Dim i As Long
Dim j As Long
Dim sValue As String
    
    
    lvwlistview.ListItems.Clear

    If Not (tblTable.Headings Is Nothing) Then
        lFields = tblTable.Cols
        'there is a valid header record for this table
        lvwlistview.ColumnHeaders.Clear
        For i = 1 To lFields
            'do the column headings
            sValue = tblTable.Headings(i)
            lvwlistview.ColumnHeaders.Add , , sValue
            If bWidenForHeading Then
                lvwlistview.ColumnHeaders(i).Width = (lvwlistview.Parent.TextWidth(sValue) + 12 * Screen.TwipsPerPixelX)
            End If
        Next
        For i = 1 To tblTable.Rows
            'get first column for the listview
            sValue = tblTable(i, 1)
            With lvwlistview.ListItems.Add(, , sValue)
                .Tag = Format(i)
                If lvwlistview.ColumnHeaders(1).Width < (lvwlistview.Parent.TextWidth(sValue) + 6 * Screen.TwipsPerPixelX) Then
                    'width of sValue more so we need to widen column
                    lvwlistview.ColumnHeaders(1).Width = (lvwlistview.Parent.TextWidth(sValue) + 6 * Screen.TwipsPerPixelX)
                End If
                'get the rest of the columns for subitems
                For j = 2 To lFields
                    sValue = tblTable(i, j)
                    .SubItems(j - 1) = sValue
                    If lvwlistview.ColumnHeaders(j).Width < (lvwlistview.Parent.TextWidth(sValue) + 6 * Screen.TwipsPerPixelX) Then
                        'width of sValue more so we need to widen column
                        lvwlistview.ColumnHeaders(j).Width = (lvwlistview.Parent.TextWidth(sValue) + 6 * Screen.TwipsPerPixelX)
                    End If
                Next
            End With
         Next
        TableToListView = tblTable.Rows
    Else
        TableToListView = -1
    End If

End Function


'----------------------------------------------------------------------------------------'
Public Function TabletoGrid(flxGrid As MSFlexGrid, tblTable As clsDataTable, _
                                Optional recColAndLength As Variant = Nothing, Optional bWidenForHeading As Boolean = True) As Long
'----------------------------------------------------------------------------------------'
' TA 19/05/2000
' reccoland length is a record with each odd field containing a column name and the next field containing the length
' uses row number as the grid row tag
'----------------------------------------------------------------------------------------'

Dim recColLengths As clsDataRecord
Dim recMinColLengths As clsDataRecord
Dim lFields As Long
Dim sValue As String
Dim lColLength As Long
Dim lRowHeight As Long
Dim i As Long
Dim j As Long


    flxGrid.Clear
    
    If Not (tblTable.Headings Is Nothing) Then
        
        
        lFields = tblTable.Cols
        flxGrid.Cols = lFields
        flxGrid.FixedCols = 0
        flxGrid.Rows = 1
    
        
        Set recColLengths = New clsDataRecord
        recColLengths.Init lFields
        If Not (recColAndLength Is Nothing) Then
            'make up record from old record
            For i = 1 To recColAndLength.Fields Step 2
                recColLengths(tblTable.GetHeadingColumn(recColAndLength.Field(i))) = recColAndLength(i + 1)
            Next
        End If
        
        Set recMinColLengths = recColLengths.Duplicate
        
        For i = 1 To lFields
            
            sValue = tblTable.Headings(i)
            
            flxGrid.TextMatrix(0, i - 1) = sValue
            lColLength = Val(recColLengths(i))
            If lColLength = 0 Then
                If bWidenForHeading Then
                    flxGrid.ColWidth(i - 1) = (flxGrid.Parent.TextWidth(sValue) + 12 * Screen.TwipsPerPixelX)
                End If
            Else
                If bWidenForHeading Then
                    If Len(sValue) < recMinColLengths(i) Then
                        'current value shorter than max
                        recMinColLengths(i) = Len(sValue)
                    End If
                Else
                    recMinColLengths(i) = 1
                End If
                flxGrid.ColWidth(i - 1) = flxGrid.Parent.TextWidth(Left(sValue, recMinColLengths(i))) + (12 * Screen.TwipsPerPixelX)
            End If
        Next
    
        For i = 1 To tblTable.Rows
           sValue = ""
           For j = 1 To lFields
               sValue = sValue & tblTable(i, j) & vbTab
           Next
           flxGrid.AddItem sValue
        
           With flxGrid
               .Row = .Rows - 1
               For j = 1 To lFields
                   .Col = j - 1
                   lColLength = Val(recColLengths(j))
                   If lColLength = 0 Then
                       If .ColWidth(j - 1) < (.Parent.TextWidth(Trim(.Text)) + 12 * Screen.TwipsPerPixelX) Then
                           .ColWidth(j - 1) = (.Parent.TextWidth(Trim(.Text)) + 12 * Screen.TwipsPerPixelX)
                       End If
                   Else
                       If recMinColLengths(j) <> lColLength Then
                           If Len(.Text) >= recColLengths(j) Then
                               recMinColLengths(j) = lColLength
                               .ColWidth(j - 1) = .Parent.TextWidth(Left(.Text, lColLength)) + (12 * Screen.TwipsPerPixelX)
                           Else
                               If Len(.Text) > recMinColLengths(j) Then
                                   If Len(.Text) < lColLength Then
                                       recMinColLengths(j) = Len(.Text)
                                   Else
                                       recMinColLengths(j) = lColLength
                                   End If
                                   lColLength = recMinColLengths(j)
                                   .ColWidth(j - 1) = .Parent.TextWidth(Left(.Text, lColLength) & "00") + (12 * Screen.TwipsPerPixelX)
                               End If
                           End If
                       End If
                       lRowHeight = (TextWrapLines(.Text, lColLength) * flxGrid.Parent.TextHeight(.Text)) + (6 * Screen.TwipsPerPixelY)
                       If .RowHeight(.Row) < lRowHeight Then
                           .RowHeight(.Row) = lRowHeight
                       End If
                       .WordWrap = True
                   End If
                   .ColAlignment(j - 1) = flexAlignLeftCenter
               Next
           End With
        
        Next
    
        If flxGrid.Rows > 1 Then
           flxGrid.FixedRows = 1
        End If
    
        TabletoGrid = lFields
    Else
        TabletoGrid = -1
    End If


End Function


'----------------------------------------------------------------------------------------'
Private Function TextWrapLines(ByVal sText As String, lCharLength As Long) As Long
'----------------------------------------------------------------------------------------'
' return number of lines if text is wrapped after lCharLength characters
'----------------------------------------------------------------------------------------'
Dim lMarker As Long
Dim sChar As String
Dim sPortion As String
Dim lLines As Long

    sPortion = sText

    Do While sPortion <> ""

        For lMarker = lCharLength To 1 Step -1
            sChar = Mid(sPortion, lMarker, 1)
            If sChar = " " Then
                Exit For
            End If
        Next

        If lMarker = 0 Then
            lMarker = InStr(sPortion, " ")
            If lMarker = 0 Then
                lMarker = lCharLength
            End If
        End If
        sPortion = Mid(sPortion, lMarker + 1)
        lLines = lLines + 1
    Loop

    TextWrapLines = lLines

End Function


'----------------------------------------------------------------------------------------'
Public Function UnloadfrmHourglass() As Boolean
'----------------------------------------------------------------------------------------'
'unload the hourlgass form if it is loaded
'----------------------------------------------------------------------------------------'
Dim ofrm As Form

    UnloadfrmHourglass = False
    
    For Each ofrm In Forms
        If LCase(ofrm.Name) = "frmhourglass" Then
            Unload ofrm
            UnloadfrmHourglass = True
            Exit For
        End If
    Next
        
End Function
