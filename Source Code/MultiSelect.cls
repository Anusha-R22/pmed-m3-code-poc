VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       MultiSelect.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Ronald Schravendeel, August 2002
'   Purpose:    Multiple Control Selection in eForm Design
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
' REVISIONS:
'   RS 16/08/2002   Initial Version
'   RS 5/9/2002     Cleanup Code, Add comments, Change variable names to standards
'   RS 10/9/2002    Previously Multiselect Collection was managed using CRFelementIndex,
'                   Changed to use CRFElementID instead, as the index changes when
'                   Rebuilding a CRFelement.
'   RS 22/10/2002   DO not take LINEs into account when aligning left and right
'                   Do not limit left/right movement when a LINE is part of the selection
'                   Correct bug in selecting/de-selecting elements on CTRL-click
'   NCJ 11 Nov 02   Added handling for new Hotlinks
'   RS 14/11/2002   Corrected bug in Command() that would cause crash when right-aligning elements AND
'                   Captions.
'                   Corrected bug in Command() that confused Caption X coordinate with Element X coordinate
'   NCJ 2 Dec 02    Patched back in Hotlink changes which were lost during previous bug fix
'   RS 13/01/03     Page did not redraw correctly when clicking outside Multiselect window
'   RS 17/01/03     Error when more than 32 controls on an eForm, changed offset to caption select
'                   controls from '* 1000' to '+ 1000'.
'   RS 20/02/03     Prevent unused 'captions' causing strange selection behaviour when selecting
'                   eForm coordinate (0,0) (top left)
'   RS 06/05/03     Set Caption Select Control Offser to high number, use a constant set to 15000
'                   Replaced all occurrences of 1000 with mlCAPTION_KEY_OFFSET
'   RS 06/05/03     Align controls to Grid while when moving
' NCJ 28 May 03 - Discontinue use of oMSE.CRFElementIndex (BUG 1325); tidied up some variables and some code
' NCJ 29 May 03 - BUG 1813 - Check for elements existing before trying to paste
'               BUG 1316 - Apply fonts/colours to questions within a selected group
' RS 26 Jun 03      Added 'Align Centre' option
'                   Create selection of pasted elements, so user can move the pasted elements easily
' NCJ 1 Jul 03 - Use new CRFElement Expandable property (Bug 1905)
' NCJ 18 Sept 06 - Bug 2770 - Don't allow removal of questions in an Open study
'----------------------------------------------------------------------------------------'

Option Explicit

Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

Private Const mlVERTICAL_SCROLL_SPEED As Long = 10    ' Define the increment step of the scrollbars while dragging
Private Const mlHORIZONTAL_SCROLL_SPEED As Long = 3
Private Const mlCAPTION_KEY_OFFSET As Long = 15000            ' Key offset of Caption select control

Private moVBForm As frmCRFDesign    ' Reference to calling CRFDesign form
Private moEForm As EFormSD          ' Reference to CRFDesign form owned Working Class

Private mbSelectMode As Boolean     ' Indicate whether SelectMode is currently active

Private mbLassoDrag As Boolean      ' Indicate Lasso Dragging Operation
Private msglLassoStartX As Single   ' Initial X Position of Lasso
Private msglLassoStartY As Single   ' Initial Y Position of Lasso

Private mbDragging As Boolean       ' True if currently dragging selection on eForm
Private msglDragStartX As Single    ' Mouse X position on eForm when Drag started
Private msglDragStartY As Single    ' Mouse Y position on eForm when Drag started
Private msglDragStartLocalX As Single   ' Mouse X position within Draghandle
Private msglDragStartLocalY As Single   ' Mouse Y position within Draghandle

Private msglDragBoxStartX As Single ' Initial X position of Draghandle
Private msglDragBoxStartY As Single ' Initial Y position of Draghandle

Private mcolSelectControls As New Collection         ' Contains references to all SelectControls
Private mcolMultiSelectElements As New Collection    ' Contains indexes/details of selected CRF elements
Private mcolPasteBuffer As New Collection            ' Contains TrialID|VersionID|CRFPageID|CRFElementID of selected items

Private mnGridSize As Integer                          ' Keep setting of grid
                                                    
Private mbWorking As Boolean         ' True while working (moving items)

Private mbCopyPaste As Boolean      ' True if Copy & Paste operation is actuve
Private mlCopyFromPageID As Long    ' CRFpageID from which elements are to be copied

'---------------------------------------------------------------------
Public Property Get Working() As Boolean
'---------------------------------------------------------------------
    
    Working = mbWorking

End Property

'---------------------------------------------------------------------
Public Property Get CopyPasteActive() As Boolean
'---------------------------------------------------------------------
' Indicates if a Copy & Paste operation is ongoing
'---------------------------------------------------------------------
    CopyPasteActive = mbCopyPaste
End Property

'---------------------------------------------------------------------
Public Property Get MultiSelectModeActive() As Boolean
'---------------------------------------------------------------------
' Indicates if Multiselect mode is ongoing
'---------------------------------------------------------------------
    MultiSelectModeActive = moVBForm.shpDragHandle.Visible
End Property

'---------------------------------------------------------------------
Public Property Get LassoDragMode()
'---------------------------------------------------------------------
' Indicates if user is currently dragging/sizing the lasso
'---------------------------------------------------------------------
    LassoDragMode = mbLassoDrag
End Property

'---------------------------------------------------------------------
Public Sub UpdateLasso(X As Long, Y As Long)
'---------------------------------------------------------------------
' During Lasso Drag Mode: Update Lasso size and position
'---------------------------------------------------------------------

    On Error GoTo Errorlabel
    
    moVBForm.shpLasso.Left = IIf(X < msglLassoStartX, X, msglLassoStartX)
    moVBForm.shpLasso.Top = IIf(Y < msglLassoStartY, Y, msglLassoStartY)
    moVBForm.shpLasso.Width = Abs(msglLassoStartX - X)
    moVBForm.shpLasso.Height = Abs(msglLassoStartY - Y)

    ' RS 05/06/2003: BUG 1482: Scroll eForm when resizing the lasso
    ScrollPage CSng(X), CSng(Y)

Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.UpdateLasso"

End Sub

'---------------------------------------------------------------------
Public Sub Init(oVBForm As frmCRFDesign, oEForm As EFormSD)
'---------------------------------------------------------------------
' Initialize Class. Set references to the VBform and worker class
' Called when user starts dragging the lasso
'---------------------------------------------------------------------

    On Error GoTo Errorlabel

    Set moVBForm = oVBForm
    Set moEForm = oEForm
    
    mnGridSize = moVBForm.GridSize        ' RS 06/05/03: Get eForm Grid Setting

Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.Init"

End Sub


'---------------------------------------------------------------------
Private Sub Class_Terminate()
'---------------------------------------------------------------------
' Cleanup
'---------------------------------------------------------------------
    
    Set moVBForm = Nothing
    Set moEForm = Nothing
    Set mcolSelectControls = Nothing
    Set mcolMultiSelectElements = Nothing
    Set mcolPasteBuffer = Nothing
    
End Sub

'---------------------------------------------------------------------
Public Sub StartSelectMode()
'---------------------------------------------------------------------
' User starts dragging a lasso
'---------------------------------------------------------------------

    On Error GoTo Errorlabel

    ' Clear Current Selectmode if new Select starts
    If Not ControlKeyPressed Then
        EndSelectMode
    End If
    
    mbSelectMode = True
    
Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.StartSelectMode"
    
End Sub

'---------------------------------------------------------------------
Public Sub EndSelectMode()
'---------------------------------------------------------------------
' Called when user calls any function that ends MultiSelect mode
'---------------------------------------------------------------------
Dim oControl As Shape
    
    On Error GoTo Errorlabel

    mbSelectMode = False
    On Error Resume Next
    
    mbLassoDrag = False
    
    EndSelection
    
    ' Unload the Controls
    For Each oControl In mcolSelectControls
        Unload oControl
    Next
    ' Empty the Collections
    Set mcolSelectControls = Nothing
    Set mcolSelectControls = New Collection
    Set mcolMultiSelectElements = Nothing
    Set mcolMultiSelectElements = New Collection
        
    moVBForm.shpLasso.Visible = False
    moVBForm.shpDragHandle.Visible = False
    
Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.EndSelectMode"

End Sub

'---------------------------------------------------------------------
Public Sub StartLassoDrag(X As Single, Y As Single)
'---------------------------------------------------------------------
' Initial creation/update of lasso control
' Called when user left clicks on eForm background
'---------------------------------------------------------------------
    
    On Error GoTo Errorlabel

    moVBForm.shpLasso.Top = Y
    moVBForm.shpLasso.Left = X
    moVBForm.shpLasso.Visible = True   'RS 29/01/2003: This line causes the eForm to go blank, and not redraw properly
    DoEvents                           'Make sure eForm redraws properly
    moVBForm.shpLasso.Width = 1
    moVBForm.shpLasso.Height = 1
    moVBForm.shpLasso.ZOrder 0
    
    
    'Debug.Print "Start Lasso Mode"
    mbLassoDrag = True
    msglLassoStartX = X
    msglLassoStartY = Y

Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.StartLassoDrag"

End Sub

'---------------------------------------------------------------------
Public Sub UpdateSelection()
'---------------------------------------------------------------------
' Create collection of selected CRFelements
' Called when user stops dragging the lasso
'---------------------------------------------------------------------

    On Error GoTo Errorlabel

    With moVBForm.shpLasso
        UpdateSelection2 .Left, .Top, .Left + .Width, .Top + .Height
        mbLassoDrag = False
    End With

Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.UpdateSelection"
End Sub


'---------------------------------------------------------------------
Public Sub EndSelection()
'-------------------------------------------------------------------
' End of Selection Mode: Remove all SelectionControls and flags
'-------------------------------------------------------------------
Dim oMSE As MultiSelectElement

    On Error GoTo Errorlabel
    
    ' Unload Select Controls
    For Each oMSE In mcolMultiSelectElements
        If oMSE.ElementSelected Then
            ' Remove Element Select Control
            Unload moVBForm.shpSelectControl(oMSE.CRFelementID)
        End If
        If oMSE.CaptionSelected Then
            Unload moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET)
        End If
    Next
    
    ' Empty Multiselect collection
    Do While mcolMultiSelectElements.Count > 0
        mcolMultiSelectElements.Remove 1
    Loop

    Set oMSE = Nothing
    
Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.EndSelection"

End Sub

'---------------------------------------------------------------------
Private Sub AddSelectControl(lIndex As Long, oControl As Control, nControlType As Integer, oElement As CRFElement)
'---------------------------------------------------------------------
' Display a dotted rectangle around a selected control, and add a reference
' to the control to the collection
'---------------------------------------------------------------------
    
    ' Ignore error if control was already loaded
    On Error Resume Next
    
    Load moVBForm.shpSelectControl(lIndex)
    
    On Error GoTo Errorlabel
    
    With moVBForm.shpSelectControl(lIndex)
        .Left = oControl.Left - 15
        .Top = oControl.Top - 15
        .Width = oControl.Width + 30
        .Height = oControl.Height + 30
        .Visible = True
        .ZOrder
    
        Select Case nControlType
            ' Add space for the popup button
            Case gn_POPUP_LIST, gn_CALENDAR:
                .Width = .Width + oControl.Height
            ' RS 24/02/2003: Added Expand Button
            Case gn_TEXT_BOX:
                If oElement.Expandable Then
                    .Width = .Width + oControl.Height
                End If
                
        End Select
    
    End With
    
    On Error Resume Next
    mcolSelectControls.Add moVBForm.shpSelectControl(lIndex), "K" & lIndex
    
Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.AddSelectControl"
    
End Sub

'---------------------------------------------------------------------
Public Function ControlKeyPressed() As Boolean
'---------------------------------------------------------------------
' Indicates if the Control Key is pressed, by calling Windows API
'---------------------------------------------------------------------
Dim arrKeys(0 To 255) As Byte

    Call GetKeyboardState(arrKeys(0)) 'Obtain Current Key States

    ControlKeyPressed = arrKeys(17) >= 128
    
End Function

'---------------------------------------------------------------------
Public Sub StartDrag(X As Single, Y As Single)
'---------------------------------------------------------------------
' Start dragging the selected elements
'---------------------------------------------------------------------
    
    On Error GoTo Errorlabel

    mbDragging = True
    msglDragStartX = moVBForm.shpDragHandle.Left + X
    msglDragStartY = moVBForm.shpDragHandle.Top + Y
    msglDragBoxStartX = moVBForm.shpDragHandle.Left
    msglDragBoxStartY = moVBForm.shpDragHandle.Top
    msglDragStartLocalX = X
    msglDragStartLocalY = Y
    
Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.StartDrag"
    
End Sub

'---------------------------------------------------------------------
Public Function Dragging() As Boolean
'---------------------------------------------------------------------
' Indicate whether user is currently dragging selected items
'---------------------------------------------------------------------
    Dragging = mbDragging
End Function

'---------------------------------------------------------------------
Public Sub Drag(X As Single, Y As Single)
'---------------------------------------------------------------------
' Actual Drag operation. Determine whether selected controls remain
' on the eForm area, and move the controlboxes (dotted rectangles)
' to new position

' If the elements would move outside te allowed area, the draghandle is
' displayed in red
'---------------------------------------------------------------------
Dim oControl As Control
Dim dX As Single
Dim dY As Single
Dim oMSE As MultiSelectElement
Dim bScrollDown As Boolean

    On Error GoTo Errorlabel

    If mbDragging Then
    
        bScrollDown = False
        
        ' Get complete Drag so far
        dX = moVBForm.shpDragHandle.Left + X - msglDragBoxStartX
        dY = moVBForm.shpDragHandle.Top + Y - msglDragBoxStartY
        
        For Each oMSE In mcolMultiSelectElements
            'Debug.Print "MSE.X = " & oMSE.X & " MSE.Y = " & oMSE.Y & "MSE.MAXX = " & oMSE.MaxX & " MSE.MAXY=" & oMSE.MaxY & " DX = " & dX & " DY = " & dY
            
            If oMSE.ElementSelected Then
                ' CHECK Y COORDINATE ONLY FOR LINE
                If (oMSE.ElementType <> gn_LINE And (oMSE.X + dX > oMSE.MaxX Or oMSE.X + dX < 100)) Or oMSE.Y + dY >= oMSE.MaxY Or oMSE.Y + dY < 10 Then
                    ' Control would move off screen: do not allow drag
                    'Debug.Print "MSE.X = " & oMSE.X & " MSE.Y = " & oMSE.Y & "MSE.MAXX = " & oMSE.MaxX & " MSE.MAXY=" & oMSE.MaxY & " DX = " & dX & " DY = " & dY
                    'Debug.Print "Refuse Drag"
                    moVBForm.shpDragHandle.BackColor = vbRed
                    Exit Sub
                Else
                    moVBForm.shpDragHandle.BackColor = vbBlue
                End If
            End If
        
            ' Check Caption Separately
            If oMSE.CaptionSelected Then
                'Debug.Print "Caption: " & dX & ", " & dY
                If (oMSE.CaptionX + dX < 10 Or oMSE.CaptionY + dY < 10) Or oMSE.CaptionX + dX > oMSE.MaxX Or oMSE.CaptionY + dY >= oMSE.MaxY Then
                    ' Caption Off Screen
                    'Debug.Print "Drag Refused: Caption Off Screen"
                    moVBForm.shpDragHandle.BackColor = vbRed
                    Exit Sub
                Else
                    moVBForm.shpDragHandle.BackColor = vbBlue
                End If
            End If
        Next
        
        ' At this point we are sure that all controls remain on the eForm area
        
        ' Move the SelectBox
        moVBForm.shpDragHandle.Left = moVBForm.shpDragHandle.Left + (X - msglDragStartLocalX)
        moVBForm.shpDragHandle.Top = moVBForm.shpDragHandle.Top + (Y - msglDragStartLocalY)
        
        
        
        ' RS 06/05/2003: Move the select controls according to the eForm Grid
        For Each oMSE In mcolMultiSelectElements
            
            If oMSE.ElementSelected Then
                If oMSE.ElementType <> gn_LINE Then
                    ' RS 05/06/2003 Do not move LINE selection box horizontally
                    moVBForm.shpSelectControl(oMSE.CRFelementID).Left = AlignToGrid(oMSE.X + dX - msglDragStartLocalX) - 15
                End If
                moVBForm.shpSelectControl(oMSE.CRFelementID).Top = AlignToGrid(oMSE.Y + dY - msglDragStartLocalY) - 15
            End If
            
            If oMSE.CaptionSelected Then
                moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET).Left = AlignToGrid(oMSE.CaptionX + dX - msglDragStartLocalX)
                moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET).Top = AlignToGrid(oMSE.CaptionY + dY - msglDragStartLocalY)
            End If
        
        Next

        
        ' RS 06/05/2003: Old move method did not take Grid Setting into account
        ' Move the SelectControls
        'For Each oControl In mcolSelectControls
        '    oControl.Left = oControl.Left + (X - msglDragStartLocalX)
        '    oControl.Top = oControl.Top + (Y - msglDragStartLocalY)
        '    Debug.Print oControl.Index
        'Next
        
        ' Scroll the form if needed
        ScrollPage moVBForm.shpDragHandle.Left, moVBForm.shpDragHandle.Top
    End If

Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.Drag"

End Sub

'---------------------------------------------------------------------
Private Function AlignToGrid(ByVal sglCoord As Single) As Single
'---------------------------------------------------------------------
' Align a coordinate to the closest grid point
' RS 06/05/03:  Copied from frmCRFdesign.
'               TODO: Call existing function instead!
'---------------------------------------------------------------------
Dim sglExtraBit As Integer
Dim sglGridPoint As Single
    
    If mnGridSize > 1 Then
        ' There is a grid
        sglExtraBit = Abs(sglCoord Mod mnGridSize)
        sglGridPoint = sglCoord \ mnGridSize       ' Integer division
        
        If sglExtraBit < (mnGridSize / 2) Then
            ' Round down to closest grid point
            AlignToGrid = sglGridPoint * mnGridSize
        Else
            ' Round up to next grid point
            If sglCoord >= 0 Then
                AlignToGrid = (sglGridPoint + 1) * mnGridSize
            Else
                ' It's a negative coordinate
                AlignToGrid = (sglGridPoint - 1) * mnGridSize
            End If
        End If
    Else
        ' No grid so leave unchanged
        AlignToGrid = sglCoord
    End If
    
End Function

'---------------------------------------------------------------------
Public Sub EndDrag(X As Single, Y As Single)
'---------------------------------------------------------------------
' User has stopped dragging selected elements. Determine relative
' move to apply to all selected elements
'---------------------------------------------------------------------
Dim dX As Single
Dim dY As Single
Dim oControl As Control

    On Error GoTo Errorlabel

    dX = moVBForm.shpDragHandle.Left - msglDragBoxStartX
    dY = moVBForm.shpDragHandle.Top - msglDragBoxStartY
    
    moVBForm.shpDragHandle.BackColor = vbBlue
        
    MoveSelectedCRFelements dX, dY
    UpdateControlBoxSizes
    
    ' RS 06/05/2003: Move the draghandle again to topmost control, as the draghandle is not aligned to the grid automatically
    moVBForm.shpDragHandle.Left = 99999
    moVBForm.shpDragHandle.Top = 99999
    For Each oControl In mcolSelectControls
        If oControl.Top < moVBForm.shpDragHandle.Top Then
            moVBForm.shpDragHandle.Left = oControl.Left - (moVBForm.shpDragHandle.Width / 2)
            moVBForm.shpDragHandle.Top = oControl.Top - (moVBForm.shpDragHandle.Width / 2)
        End If
        If oControl.Top = moVBForm.shpDragHandle.Top And oControl.Left < moVBForm.shpDragHandle.Left Then
            moVBForm.shpDragHandle.Left = oControl.Left - (moVBForm.shpDragHandle.Width / 2)
            moVBForm.shpDragHandle.Top = oControl.Top - (moVBForm.shpDragHandle.Width / 2)
        End If
    Next
    
    mbDragging = False
    
    UpdateDragHandle
    
Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.EndDrag"
    
End Sub

'---------------------------------------------------------------------
Private Sub MoveSelectedCRFelements(dX As Single, dY As Single)
'---------------------------------------------------------------------
' Move the selected elememts relative to their current position
' It is assumed that all elements can be moved to the new position
'---------------------------------------------------------------------
Dim oElement As CRFElement
Dim oMSE As MultiSelectElement
Dim lVscroll As Long
Dim lHscroll As Long

    mbWorking = True

    On Error GoTo Errorlabel

    For Each oMSE In mcolMultiSelectElements
        Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
        If oMSE.ElementSelected Then
            If oMSE.ElementType = gn_LINE Then
                ' Do not move Horizontally if LINE
                oElement.Relocate oElement.ElementX, AlignToGrid(oElement.ElementY + dY), False, oMSE.CaptionSelected
            Else
                oElement.Relocate AlignToGrid(oElement.ElementX + dX), AlignToGrid(oElement.ElementY + dY), False, oMSE.CaptionSelected
                oMSE.X = AlignToGrid(oMSE.X + dX)
            End If
            oMSE.Y = AlignToGrid(oMSE.Y + dY)
        Else
            oElement.Relocate AlignToGrid(oElement.CaptionX + dX), AlignToGrid(oElement.CaptionY + dY), True, False
        End If
        If oMSE.CaptionSelected Then
            If oElement.ControlType <> gn_LINE Then
                oMSE.CaptionX = AlignToGrid(oMSE.CaptionX + dX)
            End If
            oMSE.CaptionY = AlignToGrid(oMSE.CaptionY + dY)
        End If
    Next
    
    ' Save Scrollbar values
    lVscroll = moVBForm.vsbCRFPage.Value
    lHscroll = moVBForm.hsbCRFPage.Value
    moVBForm.RefreshMe
    DoEvents
    moVBForm.vsbCRFPage.Value = lVscroll
    moVBForm.hsbCRFPage.Value = lHscroll

    mbWorking = False
Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.MoveSelectedCRFelements"

End Sub

'-------------------------------------------------------------------
Public Sub UpdateSelection2(lX As Long, lY As Long, ldX As Long, ldY As Long)
'-------------------------------------------------------------------
' Check wether each CRFelement is inside the Lasso
' Currently an element is selected if the top-left corner is inside
' the lasso
' (Exception: a line is selected if it intersects the lasso)
' RS 20/02/2003: When selecting top left corner of eForm (0,0) make sure that unused captions are not selected
'-------------------------------------------------------------------
Dim oElement As CRFElement
Dim bElementSelected As Boolean
Dim bCaptionSelected As Boolean

    On Error GoTo Errorlabel
    
    For Each oElement In moEForm.CRFElements
        ' Do not try to include group questions
        If oElement.OwnerQGroupID = 0 Then
            With oElement
            
                ' Selection Criteria for any Element
                bElementSelected = (.ElementX >= lX And .ElementX <= ldX And .ElementY >= lY And .ElementY <= ldY)
                ' Check whether Caption is in the selection area
                bCaptionSelected = oElement.Caption <> "" And (oElement.CaptionX > lX And oElement.CaptionX < ldX) And (oElement.CaptionY > lY And oElement.CaptionY < ldY)
                

                
                'lControlIndex = moEForm.IndexForElementRow(oElement, 1)
                ' ADD HERE: Special Criteria for Each control type
                Select Case oElement.ControlType
                    Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX, gn_POPUP_LIST, gn_CALENDAR:
                    Case gn_PUSH_BUTTONS
                    Case gn_OPTION_BUTTONS
                    Case gn_ATTACHMENT
                    Case gn_COMMENT
                        ' A comment does not have a caption
                        bCaptionSelected = False
                    Case gn_HOTLINK
                        ' RS 05/06/2003: A hotlink does not have a caption
                        bCaptionSelected = False
                    Case gn_PICTURE
                        ' A picture does not have a caption
                        bCaptionSelected = False
                    Case gn_LINE
                        ' Include a line if it overlaps the lasso
                        'Set oControl = moVBForm.picCRFElement(lControlIndex)
                        bElementSelected = bElementSelected Or oElement.ElementY >= lY And oElement.ElementY <= ldY
                        ' A line does not have a caption
                        bCaptionSelected = False
                End Select
                
                ' Add the CRFelement to the selection
                AddCRFElementToSelection oElement, bElementSelected, bCaptionSelected
                
            End With
        End If
    Next
    
    moVBForm.shpDragHandle.ZOrder

Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.UpdateSelection2"


End Sub

'---------------------------------------------------------------------
Private Sub AddCRFElementToSelection(oElement As CRFElement, bElementSelected As Boolean, bCaptionSelected As Boolean)
'---------------------------------------------------------------------
' Add an individual CRFelement to the selection
' NOTE: a single object is used in the collection for each MACRO
' CRFElement, This routine ensures that captions and data entry elements
' can be selected individually
' NOTE: This routine determines the maximum X, and Y position that the
' selected element may be positioned.
'---------------------------------------------------------------------

Dim sglMaxX As Single
Dim sglMaxY As Single

Dim sglPageWidth As Single
Dim sglPageHeight As Single

Dim oMSE As MultiSelectElement
Dim lControlIndex As Long

Dim oControl As Control
Dim oCaptionControl As Control

    On Error GoTo Errorlabel
    
    sglPageWidth = moVBForm.picCRFPage.Width
    sglPageHeight = moVBForm.picCRFPage.Height
    
    
    If Not bElementSelected And Not bCaptionSelected Then
        ' Element is not part of selection
        Exit Sub
    End If
    
    If oElement.ControlType = gn_COMMENT Or oElement.ControlType = gn_HOTLINK Then
        ' Special Case: Set bElementSelected to True and bCaptionSelected to False
        bElementSelected = True
        bCaptionSelected = False
    End If
    
    ' Get reference to corresponding control
    lControlIndex = moEForm.IndexForElementRow(oElement, 1)
    'On Error Resume Next
    
    If bElementSelected Then
    Select Case oElement.ControlType
        Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX:
            Set oControl = moVBForm.lblCRFElement(lControlIndex)
            sglMaxX = sglPageWidth - 300
            sglMaxY = sglPageHeight - oControl.Height
        Case gn_POPUP_LIST, gn_CALENDAR:
            Set oControl = moVBForm.lblCRFElement(lControlIndex)
            ' RS 14/01/2003: Allow greater freedom while moving multiselection
'            sglMaxX = sglPageWidth - oControl.Width - 300
'            sglMaxY = sglPageHeight - oControl.Height
            sglMaxX = sglPageWidth - 600
            sglMaxY = sglPageHeight - oControl.Height
        Case gn_PUSH_BUTTONS
            Set oControl = moVBForm.picCRFElement(lControlIndex)
            ' RS 14/01/2003: Allow greater freedom while moving multiselection
            'sglMaxX = sglPageWidth - oControl.Width - 300
            sglMaxX = sglPageWidth - 300
            sglMaxY = sglPageHeight - oControl.Height
        Case gn_OPTION_BUTTONS
            Set oControl = moVBForm.optCRFGroup(lControlIndex)
            ' RS 14/01/2003: Allow greater freedom while moving multiselection
            'sglMaxX = sglPageWidth - oControl.Width - 300
            sglMaxX = sglPageWidth - 300
            sglMaxY = sglPageHeight - oControl.Height
        Case gn_ATTACHMENT
            Set oControl = moVBForm.cmdAttachment(lControlIndex)
            ' RS 14/01/2003: Allow greater freedom while moving multiselection
            'sglMaxX = sglPageWidth - oControl.Width - 300
            sglMaxX = sglPageWidth - 300
            sglMaxY = sglPageHeight - oControl.Height
        Case gn_COMMENT, gn_HOTLINK
            Set oControl = moVBForm.lblCaption(lControlIndex)
            ' RS 14/01/2003: Allow greater freedom while moving multiselection
            'sglMaxX = sglPageWidth - oControl.Width - 300
            sglMaxX = sglPageWidth - 300
            sglMaxY = sglPageHeight - oControl.Height
        Case gn_PICTURE
            Set oControl = moVBForm.picCRFElement(lControlIndex)
            ' RS 14/01/2003: Allow greater freedom while moving multiselection
            ' sglMaxX = sglPageWidth - oControl.Width
            sglMaxX = sglPageWidth - 300
            sglMaxY = sglPageHeight - oControl.Height
        Case gn_LINE
            Set oControl = moVBForm.picCRFElement(lControlIndex)
            sglMaxX = sglPageWidth * 1.5
            sglMaxY = sglPageHeight - oControl.Height
    End Select
    
    If oElement.QGroupID > 0 Then
        ' This is a RQG
        Set oControl = moVBForm.picQuGroup(lControlIndex)
        sglMaxX = sglPageWidth - 600
        sglMaxY = sglPageHeight - oControl.Height
        ' RS 14/01/2003: Allow greater freedom while moving multiselection
        'sglMaxX = sglPageWidth - oControl.Width - 500
        'sglMaxY = sglPageHeight - oControl.Height
    End If
    
    End If
    
    If bCaptionSelected Then
        ' Include caption information
        Set oCaptionControl = moVBForm.lblCaption(lControlIndex)
        If sglMaxX = 0 Then
            sglMaxX = sglPageWidth - oCaptionControl.Width
        Else
            'MaxX = Min(MaxX, sPageWidth - oCaptionControl.Width)
            sglMaxX = Min(sglMaxX, sglPageWidth - 500)
        End If
        If sglMaxY = 0 Then
            sglMaxY = sglPageHeight - oCaptionControl.Height
        Else
            sglMaxY = Min(sglMaxY, sglPageHeight - oCaptionControl.Height)
        End If
    End If
    
    ' Add a index reference of the selected object to the collection

    ' It is possible that the control already exists in the collection, in that case just update the existing item
    
    ' Get actual ID
    ' Debug.Print "ControlIndex = " & lControlIndex & " CRFElementID = " & moEForm.EFormItemByIndex(CLng(lControlIndex)).Element.CRFelementID
    
    
    For Each oMSE In mcolMultiSelectElements
        If oMSE.CRFelementID = oElement.CRFelementID Then
            ' Item Exists: Update only
            If bCaptionSelected Then
                oMSE.CaptionSelected = True
                oMSE.CaptionX = oElement.CaptionX
                oMSE.CaptionY = oElement.CaptionY
            End If
            If bElementSelected Then
                oMSE.ElementSelected = True
                oMSE.X = oElement.ElementX
                oMSE.Y = oElement.ElementY
            End If
            oMSE.MaxX = Min(oMSE.MaxX, sglMaxX)
            oMSE.MaxY = Min(oMSE.MaxY, sglMaxY)
            Exit For
        End If
    Next
    
    If oMSE Is Nothing Then
    
        Set oMSE = New MultiSelectElement
        oMSE.X = oElement.ElementX
        oMSE.Y = oElement.ElementY
        oMSE.MaxX = sglMaxX
        oMSE.MaxY = sglMaxY
        If bCaptionSelected Then
            oMSE.CaptionX = oElement.CaptionX
            oMSE.CaptionY = oElement.CaptionY
        End If
        oMSE.CaptionSelected = bCaptionSelected
        oMSE.ElementSelected = bElementSelected
'        oMSE.CRFElementIndex = lControlIndex
        oMSE.CRFelementID = oElement.CRFelementID
        oMSE.CRFElementDataItemID = oElement.DataItemId
        oMSE.ElementType = oElement.ControlType
        
        mcolMultiSelectElements.Add oMSE, "K" & oElement.CRFelementID
    End If
    
    ' Set the DragHandle (SelectBox to TopLeft control)
    If bElementSelected Then
        ' Set DragHandle to Selected Element
        If oControl.Top < moVBForm.shpDragHandle.Top Or Not moVBForm.shpDragHandle.Visible Then
            moVBForm.shpDragHandle.Left = oControl.Left - CInt(moVBForm.shpDragHandle.Width / 2)
            moVBForm.shpDragHandle.Top = oControl.Top - CInt(moVBForm.shpDragHandle.Height / 2)
        End If
        If (oControl.Left < moVBForm.shpDragHandle.Left And oControl.Top < moVBForm.shpDragHandle.Top) Or Not moVBForm.shpDragHandle.Visible Then
            moVBForm.shpDragHandle.Left = oControl.Left - CInt(moVBForm.shpDragHandle.Width / 2)
            moVBForm.shpDragHandle.Top = oControl.Top - CInt(moVBForm.shpDragHandle.Height / 2)
        End If
        msglDragBoxStartX = moVBForm.shpDragHandle.Left
        msglDragBoxStartY = moVBForm.shpDragHandle.Top
        moVBForm.shpDragHandle.Visible = True
    End If
    
    If bCaptionSelected Then
        ' Set Drag Handle to Selected Caption
        If oCaptionControl.Top < moVBForm.shpDragHandle.Top Or Not moVBForm.shpDragHandle.Visible Then
            moVBForm.shpDragHandle.Left = oCaptionControl.Left - CInt(moVBForm.shpDragHandle.Width / 2)
            moVBForm.shpDragHandle.Top = oCaptionControl.Top - CInt(moVBForm.shpDragHandle.Height / 2)
        End If
        If (oCaptionControl.Left < moVBForm.shpDragHandle.Left And oCaptionControl.Top < moVBForm.shpDragHandle.Top) Or Not moVBForm.shpDragHandle.Visible Then
            moVBForm.shpDragHandle.Left = oCaptionControl.Left - CInt(moVBForm.shpDragHandle.Width / 2)
            moVBForm.shpDragHandle.Top = oCaptionControl.Top - CInt(moVBForm.shpDragHandle.Height / 2)
        End If
        msglDragBoxStartX = moVBForm.shpDragHandle.Left
        msglDragBoxStartY = moVBForm.shpDragHandle.Top
        moVBForm.shpDragHandle.Visible = True
    End If
    
    ' Load and draw corresponding selection box around the CRFelement and/or Caption
    ' If bElementSelected Then AddSelectControl lControlIndex, oControl, oElement.ControlType
    ' If bCaptionSelected Then AddSelectControl lControlIndex + 1000, oCaptionControl, 0
    If bElementSelected Then AddSelectControl oElement.CRFelementID, oControl, oElement.ControlType, oElement
    If bCaptionSelected Then AddSelectControl oElement.CRFelementID + mlCAPTION_KEY_OFFSET, oCaptionControl, 0, oElement
    
    UpdateDragHandle
    
Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.AddCRFElementToSelection"
    
End Sub

'---------------------------------------------------------------------
Public Sub ToggleSelectStatus(nControlIndex As Integer, bCaptionClicked As Boolean)
'-------------------------------------------------------------------
' Toggle Select Status of indicated CRFElement.
' Called when user Control-Clicks an element
' TODO: Cannot toggle select status of caption if element already in collection
'-------------------------------------------------------------------
Dim oMSE As MultiSelectElement
Dim oElement As CRFElement
Dim lngID As Long

    ' Get CRFElementID
    lngID = moEForm.CRFElementByIndex(nControlIndex).CRFelementID

    On Error GoTo Errorlabel

    ' Check if Element part of selection
    For Each oMSE In mcolMultiSelectElements
        If oMSE.CRFelementID = lngID Then
            If oMSE.ElementSelected <> bCaptionClicked Or oMSE.CaptionSelected = bCaptionClicked Or (oMSE.ElementSelected And bCaptionClicked) Then
                
                ' STEP 1:  Remove Select Control
                
                ' CAPTION CLICKED
                If oMSE.CaptionSelected And bCaptionClicked Then
                    mcolSelectControls.Remove "K" & lngID + mlCAPTION_KEY_OFFSET
                    Unload moVBForm.shpSelectControl(lngID + mlCAPTION_KEY_OFFSET)
                    oMSE.CaptionSelected = False
                End If
                
                ' ELEMENT CLICKED
                If oMSE.ElementSelected And Not bCaptionClicked Then
                    mcolSelectControls.Remove "K" & lngID
                    Unload moVBForm.shpSelectControl(lngID)
                    oMSE.ElementSelected = False
                End If
                
                ' Special Case for Comments and Hotlinks only
                ' COMMENT CLICKED
                Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
                If oMSE.ElementSelected And bCaptionClicked And _
                 (oElement.ControlType = gn_COMMENT Or oElement.ControlType = gn_HOTLINK) Then
                    mcolSelectControls.Remove "K" & lngID
                    Unload moVBForm.shpSelectControl(lngID)
                    oMSE.ElementSelected = False
                End If
                Set oElement = Nothing
                
                If Not oMSE.CaptionSelected And Not oMSE.ElementSelected Then
                    ' Can remove item from the selection
                    Set oMSE = Nothing
                    mcolMultiSelectElements.Remove "K" & lngID
                    
                    
                End If
    
                If mcolMultiSelectElements.Count = 0 Then
                    ' No more controls selected: End selection
                    EndSelectMode
                    Exit Sub
                Else
                    UpdateDragHandle
                End If


                Exit Sub
            End If
        End If
    Next
    
    ' AT this point we know that we have to add the Element to the selection
    Set oElement = moEForm.CRFElementByIndex(nControlIndex)
    
    ' Do not include group questions in the selection
    If oElement.OwnerQGroupID = 0 Then
        AddCRFElementToSelection oElement, Not bCaptionClicked, bCaptionClicked
    Else
        'TODO: When clicking on any item in RQG, the entire group should be selected/de-selected
        ' Recursive call with the continaing Group control Index
        Call ToggleSelectStatus(moEForm.EFormGroups.EFormGroupById(oElement.OwnerQGroupID).ControlIndex, bCaptionClicked)
    End If
    
    Set oElement = Nothing

    UpdateDragHandle

Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.ToggleSelectStatus"

End Sub

' NCJ 28 May 03 - Min and Max are in libLibrary
''---------------------------------------------------------------------
'Private Function Min(Value1 As Single, Value2 As Single) As Long
''---------------------------------------------------------------------
'' Returns Minimum of two values
''---------------------------------------------------------------------
'
'    Min = IIf(Value1 < Value2, Value1, Value2)
'
'End Function
'
''---------------------------------------------------------------------
'Private Function Max(Value1 As Single, Value2 As Single) As Long
''---------------------------------------------------------------------
'' Returns Maximum of two values
''---------------------------------------------------------------------
'
'    Max = IIf(Value1 > Value2, Value1, Value2)
'
'End Function

'---------------------------------------------------------------------
Private Sub UpdateDragHandle()
'---------------------------------------------------------------------
' Move the Draghandle to the topmost selected element
' Makes sure that the Draghandle is always attached to the topmost
' selected element
'---------------------------------------------------------------------
Dim oControl As Control
Dim oMSE As MultiSelectElement

    On Error GoTo Errorlabel
    
    moVBForm.shpDragHandle.Visible = False
    moVBForm.shpDragHandle.Left = moVBForm.picCRFPage.Width
    moVBForm.shpDragHandle.Top = moVBForm.picCRFPage.Height
    
    For Each oControl In mcolSelectControls
        If oControl.Top < moVBForm.shpDragHandle.Top Then
            moVBForm.shpDragHandle.Top = oControl.Top - (moVBForm.shpDragHandle.Width / 2)
            moVBForm.shpDragHandle.Left = oControl.Left - (moVBForm.shpDragHandle.Width / 2)
        End If
    Next
    
    moVBForm.shpDragHandle.Visible = True
    
    ' RS 05/06/2003: Move DragHandle on eForm when dragged offscreen
    If moVBForm.shpDragHandle.Left < 0 Then moVBForm.shpDragHandle.Left = 0
    If moVBForm.shpDragHandle.Top < 0 Then moVBForm.shpDragHandle.Top = 0
    
    UpdateControlBoxSizes       ' Update each control box, as moving the controls might have
                                ' caused some elements to change size

    If moVBForm.shpDragHandle.Left = 0 And moVBForm.shpDragHandle.Top = 0 Then
        ' If there is a picture there, move the draghandle
        For Each oMSE In mcolMultiSelectElements
            If oMSE.ElementType = gn_PICTURE And oMSE.ElementSelected Then
                If oMSE.X <= 0 And oMSE.Y <= 0 Then
                    moVBForm.shpDragHandle.Top = mcolSelectControls("K" & oMSE.CRFelementID).Top + mcolSelectControls("K" & oMSE.CRFelementID).Height
                End If
            End If
        Next
    End If


Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.UpdateDragHandle"

End Sub

'---------------------------------------------------------------------
Private Function IsQuestionSelected() As Boolean
'---------------------------------------------------------------------
' NCJ 18 Sept 06 - Bug 2770
' Are there any questions in the currently selected items?
'---------------------------------------------------------------------
Dim bQuestion As Boolean
Dim oElement As CRFElement
Dim oMSE As MultiSelectElement

    bQuestion = False
    
    For Each oMSE In mcolMultiSelectElements
        Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
        If oElement.DataItemId > 0 Or oElement.QGroupID > 0 Then
            ' There is a question in there
            bQuestion = True
            Exit For
        End If
        
    Next
    
    IsQuestionSelected = bQuestion
    
End Function

'---------------------------------------------------------------------
Public Sub Command()
'---------------------------------------------------------------------
' Implements Right-Click menu for MultiSelect operations
' NCJ 28 May 03 - BUG 1325 - Always use oMSE.CRFelementID rather than using oMSE.ControlIndex
' NCJ 18 Sept 06 - Bug 2770 - don't allow Removal of questions in an Open study
'---------------------------------------------------------------------
Dim oMSE As MultiSelectElement
Dim sglNewX As Single
Dim sglNewY As Single
Dim oElement As CRFElement
Dim oControl As Control
    
Dim sglOffset As Single
Dim sglMaxX As Single
Dim sglMaxY As Single
Dim sglMinX As Single
Dim sglMinY As Single
Dim sglTotalWidth As Single
Dim sglTotalHeight As Single
Dim sglTotalSpaceBetween As Single
Dim sglSingleSpaceBetween As Single
Dim nItems As Integer
Dim sglStep As Single
Dim nCommandNumber As Integer
    
Dim i As Integer

Dim nControlType As Integer

Dim sEnablement As String
Dim sMenuOptions As String

    On Error GoTo Errorlabel
    
    mbWorking = True
    
    sMenuOptions = "Align Left|Align Right|Align Centre|-|Align Top|Align Bottom|-" & _
                "|Equal Vertical Spacing|Equal Horizontal Spacing|-" & _
                "|Set Font|Set Colour|-|Copy Selected Items|Remove Selected Items"
    sEnablement = ""
    ' NCJ 18 Sept 06 - Bug 2770 - Disable Remove Selected Items if study is Open and selection includes question
    If frmMenu.TrialStatus > eTrialStatus.InPreparation Then
        If IsQuestionSelected Then
            sEnablement = "||||||" & _
                            "|||" & _
                            "|||||*"
        End If
    End If
    
    nCommandNumber = frmMenu.ShowPopup(sMenuOptions, sEnablement)
    
    ' RS 06/05/2003: Loop for testing Alignment Crash
    'For i = 1 To 1000
    '    Debug.Print "Align " & i
    '    nCommandNumber = i Mod 2 + 1
    
    Select Case nCommandNumber:
        Case 1:
            '-------------------------------------------------------------------
            ' Align selected controls to the leftmost selected control
            '-------------------------------------------------------------------
            
            ' Get Leftmost Element
            sglNewX = moVBForm.picCRFPage.Width
            For Each oMSE In mcolMultiSelectElements
                If oMSE.CaptionSelected Then sglNewX = Min(sglNewX, oMSE.CaptionX)
                ' RS 22/10/2002: Do not take lines into account when aligning
                If oMSE.ElementSelected And oMSE.ElementType <> gn_LINE Then sglNewX = Min(sglNewX, oMSE.X)
            Next
            
            ' Now move all selected items to this X coordinate
            For Each oMSE In mcolMultiSelectElements
                Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
                If oMSE.ElementSelected And oMSE.ElementType <> gn_LINE Then
                    oElement.Relocate sglNewX, oElement.ElementY, False, oMSE.CaptionSelected
                    moVBForm.shpSelectControl(oMSE.CRFelementID).Left = sglNewX
                End If
                If oMSE.CaptionSelected Then
                    oElement.Relocate sglNewX, oElement.CaptionY, True, False
                    moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET).Left = sglNewX
                    oMSE.CaptionX = sglNewX
                End If
                If oMSE.ElementSelected Then
                    oMSE.X = sglNewX
                End If
            Next
        Case 2:
            '-------------------------------------------------------------------
            ' Align selected controls to the rightmost selected control
            '-------------------------------------------------------------------
            ' Get RightMost Item
            
            ' RS 22/10/2002: Replace the calculation of rightmost item, do not take LINE into account
            ' Need Element Type, so replaced the loop
            'sglNewX = 0
            'For Each oControl In mcolSelectControls
            '    sglNewX = Max(sglNewX, oControl.Left + oControl.Width)
            'Next
            
            sglNewX = 0
            For Each oMSE In mcolMultiSelectElements
                If oMSE.ElementType <> gn_LINE Then
                    ' RS 14/11/2002: Treat Caption and Element separately, and add control index
                    '                for Caption by 1000
                    If oMSE.CaptionSelected Then
                        Set oControl = moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET)
                        sglNewX = Max(sglNewX, oControl.Left + oControl.Width)
                    End If
                    If oMSE.ElementSelected Then
                        Set oControl = moVBForm.shpSelectControl(oMSE.CRFelementID)
                        sglNewX = Max(sglNewX, oControl.Left + oControl.Width)
                    End If
                End If
            Next
            
            For Each oMSE In mcolMultiSelectElements
                Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
                If oMSE.ElementSelected And oMSE.ElementType <> gn_LINE Then
                    Set oControl = moVBForm.shpSelectControl(oMSE.CRFelementID)
                    oElement.Relocate sglNewX - oControl.Width, oElement.ElementY, False, oMSE.CaptionSelected
                    moVBForm.shpSelectControl(oMSE.CRFelementID).Left = sglNewX - oControl.Width
                End If
                If oMSE.CaptionSelected Then
                    Set oControl = moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET)
                    oElement.Relocate sglNewX - oControl.Width, oElement.CaptionY, True, False
                    moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET).Left = sglNewX - oControl.Width
                    oMSE.CaptionX = sglNewX - oControl.Width
                End If
                ' RS 14/11/2002: Make sure that oControl points to the correct control when storing new X coordinate
                If oMSE.ElementSelected Then
                    Set oControl = moVBForm.shpSelectControl(oMSE.CRFelementID)
                    oMSE.X = sglNewX - oControl.Width
                End If
            Next
            
        Case 3:
            ' Align Center
            ' STEP1: Get leftmost coordinate
            sglMinX = moVBForm.picCRFPage.Width
            For Each oMSE In mcolMultiSelectElements
                If oMSE.CaptionSelected Then sglMinX = Min(sglMinX, oMSE.CaptionX)
                ' RS 22/10/2002: Do not take lines into account when aligning
                If oMSE.ElementSelected And oMSE.ElementType <> gn_LINE Then sglMinX = Min(sglMinX, oMSE.X)
            Next
            
            ' STEP2: Get rightmost coordinate
            sglMaxX = 0
            For Each oMSE In mcolMultiSelectElements
                If oMSE.ElementType <> gn_LINE Then
                    ' RS 14/11/2002: Treat Caption and Element separately, and add control index
                    '                for Caption by 1000
                    If oMSE.CaptionSelected Then
                        Set oControl = moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET)
                        sglMaxX = Max(sglMaxX, oControl.Left + oControl.Width)
                    End If
                    If oMSE.ElementSelected Then
                        Set oControl = moVBForm.shpSelectControl(oMSE.CRFelementID)
                        sglMaxX = Max(sglMaxX, oControl.Left + oControl.Width)
                    End If
                End If
            Next
            
            ' STEP3: Calculate center of page, and middle of 2 coordinates.
            sglOffset = moVBForm.picCRFPage.Width / 2 - (sglMinX + (sglMaxX - sglMinX) / 2)
            
            ' Change horizontal coordinate of selected items
            For Each oMSE In mcolMultiSelectElements
                Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
                If oMSE.ElementSelected And oMSE.ElementType <> gn_LINE Then
                    oElement.Relocate oElement.ElementX + sglOffset, oElement.ElementY, False, False
                    moVBForm.shpSelectControl(oMSE.CRFelementID).Left = moVBForm.shpSelectControl(oMSE.CRFelementID).Left + sglOffset
                    oMSE.X = oMSE.X + sglOffset
                End If
                If oMSE.CaptionSelected Then
                    oElement.Relocate oElement.CaptionX + sglOffset, oElement.CaptionY, True, False
                    moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET).Left = moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET).Left + sglOffset
                    oMSE.CaptionX = oMSE.CaptionX + sglOffset
                End If
            Next

        Case 4:
            ' No action
            mbWorking = False
            Exit Sub
            
        Case 5:
            '-------------------------------------------------------------------
            ' Align selected controls to the topmost control
            '-------------------------------------------------------------------
            
            ' Get Topmost Element
            sglNewY = moVBForm.picCRFPage.Height
            For Each oMSE In mcolMultiSelectElements
                If oMSE.CaptionSelected Then sglNewY = Min(sglNewY, oMSE.CaptionY)
                If oMSE.ElementSelected Then sglNewY = Min(sglNewY, oMSE.Y)
            Next
            
            ' Now move all selected items to this Y coordinate
            For Each oMSE In mcolMultiSelectElements
                Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
                If oMSE.ElementSelected Then
                    oElement.Relocate oElement.ElementX, sglNewY, False, oMSE.CaptionSelected
                    moVBForm.shpSelectControl(oMSE.CRFelementID).Top = sglNewY
                    oMSE.Y = sglNewY
                End If
                If oMSE.CaptionSelected Then
                    oElement.Relocate oElement.CaptionX, sglNewY, True, False
                    moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET).Top = sglNewY
                    oMSE.CaptionY = sglNewY
                End If
            Next
            
        Case 6:
            '-------------------------------------------------------------------
            ' Align selected controls to the bottom most selected control
            '-------------------------------------------------------------------
        
            ' Get Bottommost Element
            sglNewY = 0
            For Each oControl In mcolSelectControls
                sglNewY = Max(sglNewY, oControl.Top + oControl.Height)
            Next
                        
            ' Now move all selected items to this Y coordinate
            For Each oMSE In mcolMultiSelectElements
                Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
                If oMSE.ElementSelected Then
                    Set oControl = moVBForm.shpSelectControl(oMSE.CRFelementID)
                    oElement.Relocate oElement.ElementX, sglNewY - oControl.Height, False, oMSE.CaptionSelected
                    moVBForm.shpSelectControl(oMSE.CRFelementID).Top = sglNewY - oControl.Height
                    oMSE.Y = sglNewY - oControl.Height
                End If
                If oMSE.CaptionSelected Then
                    Set oControl = moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET)
                    oElement.Relocate oElement.CaptionX, sglNewY - oControl.Height, True, False
                    moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET).Top = sglNewY - oControl.Height
                    oMSE.CaptionY = sglNewY - oControl.Height
                End If
            Next
        
        Case 7:
            ' No action
            mbWorking = False
            Exit Sub
            
        Case 8:
            '-------------------------------------------------------------------
            ' Make vertical space equal between selected controls.
            ' (Top and Bottom selected control will not move)
            '-------------------------------------------------------------------
            
            ' Determine current spacing of controls
            sglTotalWidth = 0
            sglTotalHeight = 0
            sglMinX = 99999
            sglMinY = 99999
            sglMaxX = 0
            sglMaxY = 0
            nItems = 0
            For Each oControl In mcolSelectControls
                nItems = nItems + 1
                sglMinX = Min(sglMinX, oControl.Left)
                sglMinY = Min(sglMinY, oControl.Top)
                sglMaxX = Max(sglMaxX, oControl.Left + oControl.Width)
                sglMaxY = Max(sglMaxY, oControl.Top + oControl.Height)
                sglTotalWidth = sglTotalWidth + oControl.Width
                sglTotalHeight = sglTotalHeight + oControl.Height
            Next
            
            ' No action if only one element was selected
            If nItems = 1 Then
                mbWorking = False
                Exit Sub
            End If
            
            sglTotalSpaceBetween = sglMaxY - sglMinY - sglTotalHeight        ' Total free space between the controls
            sglSingleSpaceBetween = sglTotalSpaceBetween / (nItems - 1)     ' Space between each control
            
            ' Now loop through the controls and set Y coordinate
            sglNewY = sglMinY
            
            'TODO: This loop ensures that the controls are treated in the correct order
            ' Might be replaced by a smart sort, and a loop of the sorted elements
            For sglStep = 1 To 99999
                For Each oMSE In mcolMultiSelectElements
                    If oMSE.CaptionSelected And oMSE.CaptionY = sglStep Then
                        Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
                        oElement.Relocate oElement.CaptionX, sglNewY, True, False
                        Set oControl = moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET)
                        oControl.Top = sglNewY
                        sglNewY = sglNewY + sglSingleSpaceBetween + oControl.Height
                    End If
                    If oMSE.ElementSelected And oMSE.Y = sglStep Then
                        Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
                        oElement.Relocate oElement.ElementX, sglNewY, False, False
                        Set oControl = moVBForm.shpSelectControl(oMSE.CRFelementID)
                        oControl.Top = sglNewY
                        sglNewY = sglNewY + sglSingleSpaceBetween + oControl.Height
                    End If
                Next
            Next
                        
        Case 9:
            '-------------------------------------------------------------------
            ' Make horizontal space equal between selected controls.
            ' (Leftmost and Rightmost selected control will not move)
            '-------------------------------------------------------------------
        
            ' Determine current spacing of controls
            sglTotalWidth = 0
            sglTotalHeight = 0
            sglMinX = 99999
            sglMinY = 99999
            sglMaxX = 0
            sglMaxY = 0
            nItems = 0
            For Each oControl In mcolSelectControls
                nItems = nItems + 1
                sglMinX = Min(sglMinX, oControl.Left)
                sglMinY = Min(sglMinY, oControl.Top)
                sglMaxX = Max(sglMaxX, oControl.Left + oControl.Width)
                sglMaxY = Max(sglMaxY, oControl.Top + oControl.Height)
                sglTotalWidth = sglTotalWidth + oControl.Width
                sglTotalHeight = sglTotalHeight + oControl.Height
            Next
            
            ' No action if only one element was selected
            If nItems = 1 Then
                mbWorking = False
                Exit Sub
            End If
            
            sglTotalSpaceBetween = sglMaxX - sglMinX - sglTotalWidth        ' Total free space between the controls
            sglSingleSpaceBetween = sglTotalSpaceBetween / (nItems - 1)     ' Space between each control
            
            ' Now loop through the controls and set Y coordinate
            sglNewX = sglMinX
            
            'TODO: This loop ensures that the controls are treated in the correct order
            ' Might be replaced by a smart sort, and a loop of the sorted elements
            For sglStep = 1 To 99999
                For Each oMSE In mcolMultiSelectElements
                    If oMSE.CaptionSelected And oMSE.CaptionX = sglStep Then
                        Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
                        oElement.Relocate sglNewX, oElement.CaptionY, True, False
                        Set oControl = moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET)
                        oControl.Left = sglNewX
                        sglNewX = sglNewX + sglSingleSpaceBetween + oControl.Width
                    End If
                    If oMSE.ElementSelected And oMSE.X = sglStep Then
                        Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
                        oElement.Relocate sglNewX, oElement.ElementY, False, False
                        Set oControl = moVBForm.shpSelectControl(oMSE.CRFelementID)
                        oControl.Left = sglNewX
                        sglNewX = sglNewX + sglSingleSpaceBetween + oControl.Width
                    End If
                Next
            Next
        
        Case 10:
            ' No action
            mbWorking = False
            Exit Sub
            
        Case 11:
            '-------------------------------------------------------------------
            ' Set Font
            '-------------------------------------------------------------------
            
            ' NCJ 29 May 03 - Code extracted to separate routine
            Call SetSelectedItemsFont
            
'            On Error Resume Next
'            moVBForm.CommonDialog1.CancelError = True
'            moVBForm.CommonDialog1.Flags = cdlCFBoth Or cdlCFScalableOnly Or cdlCFWYSIWYG
'            moVBForm.CommonDialog1.ShowFont
'
'            If Err.Number = 24574 Then
'                DialogInformation "You will not be able to edit fonts until you have installed a default printer."
'            ElseIf Err.Number <> cdlCancel Then
'                ' Loop through selection to set all elements to selected font
'                For Each oMSE In mcolMultiSelectElements
'                    Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
'                    nControlType = oElement.ControlType
'                    If nControlType <> gn_LINE And nControlType <> gn_PICTURE Then
'                        If oMSE.CaptionSelected Or _
'                         ((nControlType = gn_COMMENT Or nControlType = gn_HOTLINK) And oMSE.ElementSelected) Then
'                            Call oElement.SetCaptionFontDetails(moVBForm.CommonDialog1.FontName, moVBForm.CommonDialog1.FontSize, _
'                                                                moVBForm.CommonDialog1.FontBold, moVBForm.CommonDialog1.FontItalic)
'                        End If
'                        If oMSE.ElementSelected And nControlType <> gn_COMMENT And nControlType <> gn_HOTLINK Then
'                            Call oElement.SetFontDetails(moVBForm.CommonDialog1.FontName, moVBForm.CommonDialog1.FontSize, _
'                                                         moVBForm.CommonDialog1.FontBold, moVBForm.CommonDialog1.FontItalic)
'                        End If
'                    End If
'                Next
'            End If
                
        Case 12:
            '-------------------------------------------------------------------
            ' Set Colour
            '-------------------------------------------------------------------
            
            ' NCJ 29 May 03 - Code extracted to separate routine
            Call SetSelectedItemsColour
            
'            On Error Resume Next
'            moVBForm.CommonDialog1.CancelError = True
'            moVBForm.CommonDialog1.ShowColor
'
'            If Err.Number <> cdlCancel Then
'                ' Loop through selection to set all elements to selected font
'                For Each oMSE In mcolMultiSelectElements
'                    Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
'                    nControlType = oElement.ControlType
'                    If nControlType <> gn_LINE And nControlType <> gn_PICTURE Then
'                        If oMSE.CaptionSelected Or _
'                         ((nControlType = gn_COMMENT Or nControlType = gn_HOTLINK) And oMSE.ElementSelected) Then
'                            oElement.SetCaptionFontColour (moVBForm.CommonDialog1.Color)
'                        End If
'                        If oMSE.ElementSelected And nControlType <> gn_COMMENT And nControlType <> gn_HOTLINK Then
'                            Call oElement.SetFontColour(moVBForm.CommonDialog1.Color)
'                        End If
'                    End If
'                    ' This causes a crash in the colMultiselect collection
'                    ' Call moEForm.ReBuildCRFElement(oElement, False, True)
'                Next
'            End If
            
        Case 13:
            ' No action
            mbWorking = False
            Exit Sub
            
        Case 14:
            '-------------------------------------------------------------------
            ' Copy selected elements to other eForm
            ' Will prepare the paste operation. Actual paste will occur when user
            ' opens a different eForm page.
            '-------------------------------------------------------------------
        
            ' RS 5/9/2002: Prepare the Paste Buffer, and enable menu item
            CopyToPasteBuffer
            
            ' RS 26/06/2003: Set all pasted elements selected
             
        Case 15:
            '-------------------------------------------------------------------
            ' Remove Items from eForm
            '-------------------------------------------------------------------
            If DialogQuestion("Do you really want to remove the selected items from the eForm?") = vbYes Then
                ' User confirmed: Remove the items
                For Each oMSE In mcolMultiSelectElements
                    
                    Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
                    moEForm.DeleteCRFElement oElement

                Next
                ' Refresh all the question lists
                Call frmMenu.RefreshQuestionLists(moVBForm.ClinicalTrialId)
                ' End select mode,as all selected items are removed
                EndSelectMode
            End If
            
    End Select

    
    ' RS 13/01/03: Do not Refresh the entire eForm if no menu item was selected.
    If nCommandNumber <> -1 And nCommandNumber <> 13 Then    ' Do not refresh if Copy&Paste
        '-------------------------------------------------------------------
        ' Redraw the eForm
        '-------------------------------------------------------------------
        moVBForm.RefreshMe  ' Will rebuild the eForm
        DoEvents            ' Make sure screen is updated before going further
        UpdateDragHandle
    End If

    ' RS 06/05/2003: Loop for testing Alignment Crash
    ' Next

    mbWorking = False
    
Exit Sub
Errorlabel:
    If MACROErrorHandler("MultiSelect", Err.Number, Err.Description, "Command", Err.Source) = Retry Then
        Resume
    End If
    mbWorking = False

End Sub

'-------------------------------------------------------------------
Private Sub SetSelectedItemsColour()
'-------------------------------------------------------------------
' NCJ 29 May 03
' Set the colour of all selected items
' Code extracted from Command routine, and added QGroup handling
'-------------------------------------------------------------------
Dim oElement As CRFElement
Dim oMSE As MultiSelectElement
Dim nControlType As Integer
Dim lChosenColour As Long
Dim oQG As EFormGroupSD
Dim oQGElement As CRFElement
Dim vCRFElementId As Variant

    ' Ignore errors!
    On Error Resume Next
    
    moVBForm.CommonDialog1.CancelError = True
    moVBForm.CommonDialog1.ShowColor
    
    If Err.Number = cdlCancel Then Exit Sub
    
    lChosenColour = moVBForm.CommonDialog1.Color
    
    ' Loop through selection to set all elements to selected font
    For Each oMSE In mcolMultiSelectElements
        Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
        nControlType = oElement.ControlType
        If nControlType <> gn_LINE And nControlType <> gn_PICTURE Then
            If oMSE.CaptionSelected Or _
             ((nControlType = gn_COMMENT Or nControlType = gn_HOTLINK) And oMSE.ElementSelected) Then
                Call oElement.SetCaptionFontColour(lChosenColour)
            End If
            If oMSE.ElementSelected And nControlType <> gn_COMMENT And nControlType <> gn_HOTLINK Then
                Call oElement.SetFontColour(lChosenColour)
            End If
        End If
        ' NCJ 29 May 03 - If a group, set the colour of all the group items too
        If oMSE.ElementSelected And oElement.QGroupID > 0 Then
            Set oQG = moVBForm.EFormGroupById(oElement.QGroupID)
            For Each vCRFElementId In oQG.Elements
                Set oQGElement = moEForm.CRFElementById(CInt(vCRFElementId))
                Call oQGElement.SetFontColour(lChosenColour)
                Call oQGElement.SetCaptionFontColour(lChosenColour)
            Next
        End If
    Next

    Set oElement = Nothing
    Set oMSE = Nothing
    Set oQG = Nothing
    Set oQGElement = Nothing
    
End Sub

'-------------------------------------------------------------------
Private Sub SetSelectedItemsFont()
'-------------------------------------------------------------------
' NCJ 29 May 03
' Set the font of all selected items
' Code extracted from Command routine, and added QGroup handling
'-------------------------------------------------------------------
Dim oElement As CRFElement
Dim oMSE As MultiSelectElement
Dim nControlType As Integer
Dim sFontName As String
Dim nFontSize As Integer
Dim bBold As Boolean
Dim bItalic As Boolean
Dim oQG As EFormGroupSD
Dim vCRFElementId As Variant
Dim oQGElement As CRFElement

    ' Ignore errors!
    On Error Resume Next
    
    moVBForm.CommonDialog1.CancelError = True
    moVBForm.CommonDialog1.Flags = cdlCFBoth Or cdlCFScalableOnly Or cdlCFWYSIWYG
    moVBForm.CommonDialog1.ShowFont
    
    If Err.Number = 24574 Then
        DialogInformation "You will not be able to edit fonts until you have installed a default printer."
        Exit Sub
    End If
    
    If Err.Number = cdlCancel Then Exit Sub
    
    sFontName = moVBForm.CommonDialog1.FontName
    nFontSize = moVBForm.CommonDialog1.FontSize
    bBold = moVBForm.CommonDialog1.FontBold
    bItalic = moVBForm.CommonDialog1.FontItalic
    
    ' Loop through selection to set all elements to selected font
    For Each oMSE In mcolMultiSelectElements
        Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
        nControlType = oElement.ControlType
        If nControlType <> gn_LINE And nControlType <> gn_PICTURE Then
            If oMSE.CaptionSelected Or _
             ((nControlType = gn_COMMENT Or nControlType = gn_HOTLINK) And oMSE.ElementSelected) Then
                Call oElement.SetCaptionFontDetails(sFontName, nFontSize, bBold, bItalic)
            End If
            If oMSE.ElementSelected And nControlType <> gn_COMMENT And nControlType <> gn_HOTLINK Then
                Call oElement.SetFontDetails(sFontName, nFontSize, bBold, bItalic)
            End If
        End If
        ' NCJ 29 May 03 - If it's a group, set all the group elements too
        If oElement.QGroupID > 0 Then
            Set oQG = moVBForm.EFormGroupById(oElement.QGroupID)
            For Each vCRFElementId In oQG.Elements
                Set oQGElement = moEForm.CRFElementById(CInt(vCRFElementId))
                Call oQGElement.SetFontDetails(sFontName, nFontSize, bBold, bItalic)
                Call oQGElement.SetCaptionFontDetails(sFontName, nFontSize, bBold, bItalic)
            Next
        End If
    Next

    Set oElement = Nothing
    Set oMSE = Nothing
    Set oQG = Nothing
    Set oQGElement = Nothing
    
End Sub

'-------------------------------------------------------------------
Private Sub UpdateControlBoxSizes()
'-------------------------------------------------------------------
' After moving the CRF Elements, text boxes might have changed size
' Update the size of all control boxes to reflect the current
' CRFelement size
'-------------------------------------------------------------------
Dim oMSE As MultiSelectElement
Dim oControl As Control
Dim oElement As CRFElement
Dim oCRFControl As Control
Dim lControlIndex As Long

    On Error GoTo Errorlabel
    
    For Each oMSE In mcolMultiSelectElements
        
        If oMSE.CaptionSelected Then
            ' GET ControlIndex
            lControlIndex = moEForm.IndexForElementRow(moEForm.CRFElementById(oMSE.CRFelementID), 1)
            
            Set oControl = moVBForm.shpSelectControl(oMSE.CRFelementID + mlCAPTION_KEY_OFFSET)
            
            
            oControl.Width = moVBForm.lblCaption(lControlIndex).Width
            oControl.Height = moVBForm.lblCaption(lControlIndex).Height
            Set oControl = Nothing
        End If
        
        If oMSE.ElementSelected Then
            
            Set oElement = moEForm.CRFElementById(oMSE.CRFelementID)
            ' GET ControlIndex
            lControlIndex = moEForm.IndexForElementRow(oElement, 1)
            
            ' Get the correct control from the eForm
            Select Case oElement.ControlType
            Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX:
                Set oCRFControl = moVBForm.lblCRFElement(lControlIndex)
            Case gn_POPUP_LIST, gn_CALENDAR:
                Set oCRFControl = moVBForm.lblCRFElement(lControlIndex)
            Case gn_PUSH_BUTTONS
                Set oCRFControl = moVBForm.picCRFElement(lControlIndex)
            Case gn_OPTION_BUTTONS
                Set oCRFControl = moVBForm.optCRFGroup(lControlIndex)
            Case gn_ATTACHMENT
                Set oCRFControl = moVBForm.cmdAttachment(lControlIndex)
            Case gn_COMMENT, gn_HOTLINK
                Set oCRFControl = moVBForm.lblCaption(lControlIndex)
            Case gn_PICTURE
                Set oCRFControl = moVBForm.picCRFElement(lControlIndex)
            Case gn_LINE
                Set oCRFControl = moVBForm.picCRFElement(lControlIndex)
            End Select
            If oElement.QGroupID > 0 Then
                Set oCRFControl = moVBForm.picQuGroup(lControlIndex)
            End If
            Set oControl = moVBForm.shpSelectControl(oElement.CRFelementID)
            oControl.Width = oCRFControl.Width + 30
            oControl.Height = oCRFControl.Height + 30
            
            Select Case oElement.ControlType
                ' Add space for the popup button
                Case gn_POPUP_LIST, gn_CALENDAR:
                    oControl.Width = oControl.Width + oCRFControl.Height
                ' RS 24/02/2003: Added Expand Button
                Case gn_TEXT_BOX:
                    If oElement.Expandable Then
                        oControl.Width = oControl.Width + oCRFControl.Height
                    End If
            End Select
            
            Set oControl = Nothing
            Set oCRFControl = Nothing
            Set oElement = Nothing
        End If
        
    Next

Exit Sub
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.UpdateControlBoxSizes"

End Sub


'---------------------------------------------------------------------
Public Function PasteBufferAvailable() As Boolean
'---------------------------------------------------------------------
' Indicates if items are available for Paste operation
' Used for display/hide menu
'---------------------------------------------------------------------
    
    PasteBufferAvailable = (mcolPasteBuffer.Count > 0)

End Function

'---------------------------------------------------------------------
Private Function CopyToPasteBuffer()
'---------------------------------------------------------------------
' Copy Key ID's of selected items to the Paste Buffer
'---------------------------------------------------------------------
Dim oMSE As MultiSelectElement

    On Error GoTo Errorlabel

    ' STEP 1: Empty existing buffer
    Set mcolPasteBuffer = Nothing
    Set mcolPasteBuffer = New Collection
        
    mlCopyFromPageID = moVBForm.CRFPageId

    ' STEP 2: Loop through current selection and add each CRFElementID
    For Each oMSE In mcolMultiSelectElements
'        mcolPasteBuffer.Add oMSE.CRFelementID, Str(oMSE.CRFelementID)
        mcolPasteBuffer.Add oMSE.CRFelementID
    Next

    ' STEP 3: Enable Paste menu option
    frmMenu.mnuFLDPaste.Enabled = True

    ' STEP 4: End Select Mode
    EndSelectMode
    
Exit Function
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "MultiSelect.CopyToPasteBuffer"

End Function

'---------------------------------------------------------------------
Public Sub PasteBuffer()
'---------------------------------------------------------------------
' Paste elements in the PasteBuffer onto the active form
' RS 26/06/2003: Set all pasted elements as selected
'---------------------------------------------------------------------
Dim colElementsOnForm As New Collection
Dim oElement As CRFElement
Dim sMaxElementID As Single

    ' RS 26/06/2003: Get ID's of Element already on the form.
    sMaxElementID = 0
    For Each oElement In moEForm.CRFElements
        colElementsOnForm.Add oElement.CRFelementID
        sMaxElementID = Max(sMaxElementID, oElement.CRFelementID)
    Next
    
    If mcolPasteBuffer.Count = 0 Then
        DialogInformation "No items are currently selected"
        Exit Sub
    End If
    
    ' NCJ 29 May 03 - BUG 1813 - Check for elements existing first
    If Not CRFElementsExist(moVBForm.ClinicalTrialId, moVBForm.VersionId, _
                            mlCopyFromPageID, mcolPasteBuffer) Then
        DialogError "The eForm items could not be pasted because the original items have been removed"
        ' Empty the paste buffer because it's no longer of any use to us
        Set mcolPasteBuffer = Nothing
        Set mcolPasteBuffer = New Collection
        Exit Sub
    End If
    
    If Not PasteCRFElements(moVBForm.ClinicalTrialId, moVBForm.VersionId, _
                        mlCopyFromPageID, moVBForm.CRFPageId, mcolPasteBuffer) Then
        DialogInformation "The items could not be pasted because one of the questions is already present on this eForm"
    Else
        ' Paste Succeeded
        moVBForm.RefreshMe
        DoEvents
        frmMenu.RefreshQuestionLists moVBForm.ClinicalTrialId
    End If

    ' Start Select Mode, and add all new elements to selection
    StartSelectMode
    For Each oElement In moEForm.CRFElements
        ' Add element if ID > Maximum ID
        If oElement.CRFelementID > sMaxElementID Then
            If oElement.CaptionX = 0 And oElement.CaptionY = 0 Then
                AddCRFElementToSelection oElement, True, False
            Else
                AddCRFElementToSelection oElement, True, True
            End If
        End If
    Next
    UpdateDragHandle
    

End Sub

'---------------------------------------------------------------------
Private Sub ScrollPage(X As Single, Y As Single)
'---------------------------------------------------------------------
' Scroll the CRFPage while dragging mouse
'---------------------------------------------------------------------

    If Y + moVBForm.picCRFPage.Top < 500 Then
        If moVBForm.vsbCRFPage.Value - mlVERTICAL_SCROLL_SPEED > moVBForm.vsbCRFPage.Min Then
            moVBForm.vsbCRFPage.Value = moVBForm.vsbCRFPage.Value - mlVERTICAL_SCROLL_SPEED
        End If
    ElseIf moVBForm.Height - Y - moVBForm.picCRFPage.Top < 1500 Then
        If moVBForm.vsbCRFPage.Value + mlVERTICAL_SCROLL_SPEED < moVBForm.vsbCRFPage.Max Then
            moVBForm.vsbCRFPage.Value = moVBForm.vsbCRFPage.Value + mlVERTICAL_SCROLL_SPEED
        End If
    End If
    
    ' Horizontal Scroll
    If X + moVBForm.picCRFPage.Left < 500 Then
        If moVBForm.hsbCRFPage.Value - mlHORIZONTAL_SCROLL_SPEED > moVBForm.hsbCRFPage.Min Then
            moVBForm.hsbCRFPage.Value = moVBForm.hsbCRFPage.Value - mlHORIZONTAL_SCROLL_SPEED
        End If
    ElseIf moVBForm.Width - X - moVBForm.picCRFPage.Left < 1500 Then
        If moVBForm.hsbCRFPage.Value + mlHORIZONTAL_SCROLL_SPEED < moVBForm.hsbCRFPage.Max Then
            moVBForm.hsbCRFPage.Value = moVBForm.hsbCRFPage.Value + mlHORIZONTAL_SCROLL_SPEED
        End If
    End If

End Sub

'-------------------------------------------------------------------
Private Sub ListControls()
'-------------------------------------------------------------------
' FOR DEBUGGING ONLY
' Debug Procedure: List all controls currently loaded. Used to
' verify cleanup routines
'-------------------------------------------------------------------
Dim oMSE As MultiSelectElement
Dim oControl As Control
Dim oElement As CRFElement

    For Each oMSE In mcolMultiSelectElements
        Debug.Print "MSE: ID = " & oMSE.CRFelementID & " Caption = " & oMSE.CaptionSelected & " Element = " & oMSE.ElementSelected
    Next
    
    For Each oControl In moVBForm.Controls
            On Error Resume Next
            If oControl.Index > 0 Then
                Debug.Print "Control: " & oControl.Name & "[" & oControl.Index & "]" & IIf(oControl.Name = "lblCaption", oControl.Caption, "")
            End If
    Next

    For Each oElement In moEForm.CRFElements
        Debug.Print "CRFElements: " & oElement.CRFelementID & "  -  " & oElement.ControlType & " - " & oElement.Caption
    Next
End Sub

