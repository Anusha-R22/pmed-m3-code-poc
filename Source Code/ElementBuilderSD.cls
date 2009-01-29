VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ElementBuilderSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       ElementBuilderSD.cls
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Nicky Johns, November 2001
'   Purpose:    Element Builder for eForms in Study Definition in MACRO 3.0.
'               Based on ElementBuilder in MACRO DE
'               Includes Question Groups
'----------------------------------------------------------------------------------------'
' This just builds the elements but doesn't do much other processing
'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 27-30 Nov 01 - Initial development based on ElementBuilder.cls (in MACRO DE)
'   NCJ 11 Dec 01 - Tweaking drawing of RQGs
'   NCJ 12 Dec 01 - Adjusted size of empty group box
'   NCJ 13 Dec 01 - Added calculation of offset Top Left for group members
'   NCJ 19 Dec 01 - Reset group's ColWidths collection before redrawing
'   NCJ 2 Jan 02 - Use constants for imgStatus and lblNRCTC in GetElementSize
'   ZA 19/07/2002 - added functionality for font/colour for caption & comment
'   ZA 19/09/2002 - added display lenght property for text box and drop down list
'   NCJ 2-7 Oct 02 - Added handling for RQG horizontal scroll bar
'   NCJ 9 Oct 02 - Discontinue use of mn_SPACE_FOR_STATUS_ICON; final tweaking of RQG horiz. scrolling
'   NCJ 4 Nov 02 - Store suitable hscroll dividing factor in scroll bar tag
'                   Ensure option boxes have correct mouse pointer
'   ASH 5/11/2002 - Modified sub BuildElements
'   NCJ 6-7 Nov 02 - Added Hotlink handlers
'   NCJ 27 Feb 03 - Use new calculation for Option Button widths
'   RS  11/03/2003 - Make optionbuttons visible only when finised init to prevent default values from flashing on screen
'   RS  05/06/2003 - Update oElement memory structure when adjusting Caption position (when offscreen)
'   NCJ 1 Jul 03 - Use new CRFElement Expandable property (Bug 1905)
'   ic 04/07/2005 added clinical coding
'   TA  07/07/2005: seeting of Font.Charset = 1 to allow eastern european characters. CBD2591.
'----------------------------------------------------------------------------------------'

Option Explicit

' The VB form we're drawing to
' NB Assume this is an instance of frmCRFDesign
Private moVBForm As Form
Private moEForm As EFormSD

Private mnNextControlIndex As Integer
Private mnNextPushButtonIndex As Integer

Private mnTabOrder As Integer

' NCJ 4 Oct 2002: width of changecount icon (copied from DE)
Private Const m_CHANGECOUNT_IMAGE_WIDTH = 120
' NCJ 27 Feb 03 - width of Notes/Comments icon (copied from DE)
Private Const m_NOTESTATUS_IMAGE_WIDTH = 120

' NCJ 4 Oct 2002: width of image status icon (copied from DE)
Private Const m_STATUS_IMAGE_WIDTH = 240
Private Const m_LBL_NRCTC_WIDTH = 410

Private Const msglGAP As Single = 50

Private mcolControlIndexes As Collection    ' of integer indexes
Private mcolFormItems As Collection     ' of EFormItemSD objects

Private mcolGroupsByIndex As Collection    ' of EFormGroupSD objects
Private mcolGroupsByID As Collection    ' of EFormGroupSD objects

' Max Y coord on eForm
Private msglMaxY As Single

'----------------------------------------------------------
Public Sub Init(oVBForm As Form, oEForm As EFormSD)
'----------------------------------------------------------
' Initialise with specified VB Form
'----------------------------------------------------------
    
    Set moVBForm = oVBForm
    Set moEForm = oEForm

End Sub

'----------------------------------------------------------
Public Sub BuildElements()
'----------------------------------------------------------
' Draw all the form elements
' Assume that the eForm has its EFormElements collection all ready
' and that the eForm is empty of all other controls
'----------------------------------------------------------
Dim oElement As CRFElement

    On Error GoTo ErrLabel
    
    ' Do we know our forms?
    If moVBForm Is Nothing Then Exit Sub

    ' Have we got some eForm Elements to work from?
    If moEForm.CRFElements Is Nothing Then Exit Sub

    moVBForm.MousePointer = vbHourglass
    
    moVBForm.vsbCRFPage.Visible = False
    moVBForm.picCRFPage.Visible = False
    moVBForm.picCRFPage.BackColor = moEForm.BackgroundColour
    
    '  Unload any existing controls
    UnloadControls
    
    'set form height
    moVBForm.picCRFPage.Height = 20000
    'set form width
    'prior to fixed form widths the following line was used
    'moVBform.picCRFPage.Width = moVBform.Width - moVBform.vsbCRFPage.Width
    '10/12/2001: VTRACK Changes - THIS WILL BE REPLACED WHEN WHEN SUPPORT LANDSCAPE FORMS IN OFFICIAL MACRO
    'if landscape is passed as a command line argument, allow landscape eforms
    If lCase(Command) = "landscape" Then
        'landscape width (currently for VTRACK)
        moVBForm.picCRFPage.Width = glLANDSCAPE_WIDTH
    Else
        ' ASH 5/11/02 - Use new eFormWidth property
        moVBForm.picCRFPage.Width = moEForm.EFormWidth
    End If
        
    ' Initialise everything
    Set mcolControlIndexes = New Collection
    Set mcolFormItems = New Collection
    Set mcolGroupsByIndex = New Collection
    Set mcolGroupsByID = New Collection
    
'Debug.Print Timer & " Building elements " & moEForm.VBForm.CRFPageID
    
    mnNextControlIndex = 1
    mnNextPushButtonIndex = 1
    mnTabOrder = 0
    msglMaxY = 20000    ' Default height of 20000
    
    ' Draw the "top level" eForm Elements (i.e. not those inside groups)
    ' (they get built by the owner group)
    For Each oElement In moEForm.CRFElements
        ' Don't draw the grid for each one
        Call BuildElement(oElement, False)
    Next
    
'Debug.Print Timer & " Done Building elements " & moEForm.VBForm.CRFPageID
    
    ' Set form height
    moVBForm.picCRFPage.Height = msglMaxY + 4000
    
    Call SetPageSize(moVBForm)
    
    '  Make sure popup list is hidden
    moVBForm.lvwPopupList.Visible = False
    '   Show the form after the controls have been built
    moVBForm.picCRFPage.Visible = True
    
    Call SetFormZOrder
    
    moVBForm.picCRFPage.Visible = True
    moVBForm.vsbCRFPage.Visible = True
    moVBForm.DisplayGrid   'SDM 22/02/00

    moVBForm.MousePointer = vbDefault

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilderSD.BuildElements"

End Sub

'---------------------------------------------------------------------
Public Sub BuildElement(oElement As CRFElement, _
                        Optional ByVal bRedrawGrid As Boolean = True)
'---------------------------------------------------------------------
' Build single top-level element
' Assume all intialisations have been done
' and that previous controls for this element (if any) have been removed.
' This does nothing for group member elements
' If bRedrawGrid is TRUE then the eForm grid gets redrawn too
'---------------------------------------------------------------------

    If oElement.OwnerQGroupID = 0 Then
        Call BuildAnElement(oElement, 1, moVBForm.picCRFPage, _
                        oElement.ElementY, oElement.ElementX)
        If bRedrawGrid Then
            ' Ensure grid is still displayed if necessary
            Call moVBForm.DisplayGrid
        End If
    End If

End Sub

'---------------------------------------------------------------------
Public Function IndexForElementRow(oElement As CRFElement, nRow As Integer) As Integer
'---------------------------------------------------------------------
' Returns control array index used for this element/row
'---------------------------------------------------------------------

    IndexForElementRow = mcolControlIndexes(GetItemKey(oElement, nRow))

End Function

'---------------------------------------------------------------------
Public Function EFormItemByIndex(nIndex As Integer) As EFormItemSD
'---------------------------------------------------------------------
' Returns the element/row represented by this control array index
'---------------------------------------------------------------------

    Set EFormItemByIndex = mcolFormItems(Str(nIndex))

End Function

'---------------------------------------------------------------------
Private Function BuildAnElement(oElement As CRFElement, nRow As Integer, _
                            oContainer As PictureBox, _
                            ByVal lTop As Long, ByVal lLeft As Long) As Integer
'---------------------------------------------------------------------
' Build a single eForm Element onto the VB form
' The nRow is which group row it's in (default = 1 for non-group items)
' Return the index of the element's control(s)
' ic 14/06/2005 added clinical coding
'---------------------------------------------------------------------
Dim nControlType As Integer
Dim nIndex As Integer

    On Error GoTo ErrLabel
    
    nControlType = oElement.ControlType
    nIndex = mnNextControlIndex
    mnNextControlIndex = mnNextControlIndex + 1

' Debug.Print Timer & " BuildAnElement " & oElement.CRFElementId & ", index " & nIndex
    
    If oElement.DataItemId > 0 And oElement.QGroupID = 0 Then
    ' It's a plain data item
    
        If oElement.OwnerQGroupID = 0 Then
            ' Build caption for non-group questions only
            BuildCaption oElement, nIndex, oContainer
        End If
        
        ' Build the rest of it according to its type
        Select Case nControlType
        Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX, _
                gn_POPUP_LIST, gn_CALENDAR
            BuildTextBox oElement, nIndex, oContainer, lTop, lLeft
            
            If oElement.Expandable Then
                Call BuildExpandButton(moVBForm.lblCRFElement(nIndex))
            End If
            
            If nControlType = gn_POPUP_LIST Then
                Call BuildPopupButton(moVBForm.lblCRFElement(nIndex))
            End If
            
            If nControlType = gn_CALENDAR Then
                Call BuildCalendarButton(moVBForm.lblCRFElement(nIndex))
            End If

            'ic 14/06/2005 clinical coding: add browse button
            If (oElement.DataItemType = eDataType.Thesaurus) Then
                Call BuildDictionaryButton(moVBForm.lblCRFElement(nIndex))
            End If
            
        Case gn_PUSH_BUTTONS
            BuildPushButtons oElement, nIndex, oContainer, lTop, lLeft
            
        Case gn_OPTION_BUTTONS
           BuildOptionButtons oElement, nIndex, oContainer, lTop, lLeft
            
        Case gn_ATTACHMENT
            BuildAttachmentButton oElement, nIndex, oContainer, lTop, lLeft
            
        End Select
                
    ElseIf oElement.QGroupID > 0 Then
        BuildCaption oElement, nIndex, oContainer
        ' Draw the group
        Call BuildGroup(oElement, nIndex, oContainer, lTop, lLeft)
    
    Else    ' It's a visual-only element (not a question)
        
        Select Case nControlType
        Case gn_LINE
            BuildLine oElement, nIndex, oContainer, lTop, lLeft
        Case gn_COMMENT
            BuildComment oElement, nIndex, oContainer, lTop, lLeft, False
        Case gn_PICTURE
            BuildPicture oElement, nIndex, oContainer, lTop, lLeft
        Case gn_HOTLINK
            BuildComment oElement, nIndex, oContainer, lTop, lLeft, True
        End Select
    End If
    
    ' Store in our collections
    Call AddElementToCollections(oElement, nRow, nIndex)

    ' Return the control index used for this element
    BuildAnElement = nIndex
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & _
            "|ElementBuilderSD.BuildAnElement(" & oElement.CRFelementID & "," & nRow & ")"

End Function

'---------------------------------------------------------------------
Private Sub AddElementToCollections(oElement As CRFElement, _
                ByVal nRow As Integer, ByVal nIndex As Integer)
'---------------------------------------------------------------------
' Add a new element to our collections
'---------------------------------------------------------------------
Dim oGroupItem As EFormItemSD
    
    ' Store in our collections
    Set oGroupItem = New EFormItemSD
    Set oGroupItem.Element = oElement
    oGroupItem.RowNo = nRow
    oGroupItem.ControlIndex = nIndex

    mcolControlIndexes.Add nIndex, GetItemKey(oElement, nRow)
    mcolFormItems.Add oGroupItem, Str(nIndex)
    
    Set oGroupItem = Nothing

End Sub

'---------------------------------------------------------------------
Public Function EFormQGroupByIndex(nIndex As Integer) As EFormGroupSD
'---------------------------------------------------------------------
' Return the eFormQGroup object corresponding to the given Control Index
'---------------------------------------------------------------------

    Set EFormQGroupByIndex = mcolGroupsByIndex.Item(Str(nIndex))
    
End Function

'---------------------------------------------------------------------
Public Function EFormQGroupById(lQGroupId As Long) As EFormGroupSD
'---------------------------------------------------------------------
' Return the eFormQGroup object corresponding to the given Group ID
'---------------------------------------------------------------------

    Set EFormQGroupById = mcolGroupsByID.Item(Str(lQGroupId))
    
End Function
 
'---------------------------------------------------------------------
Private Sub BuildGroup(oElement As CRFElement, _
                        nIndex As Integer, oContainer As PictureBox, _
                        lTop As Long, lLeft As Long)
'---------------------------------------------------------------------
' Build a question group
'---------------------------------------------------------------------
Dim oPicQG As PictureBox
Dim oEFQGroup As EFormGroupSD
Dim nRow As Integer
Dim bGroupHasCaption As Boolean

    On Error GoTo ErrLabel
    
    ' Get the defined EFormGroup object
    Set oEFQGroup = moEForm.EFormGroups.EFormGroupById(oElement.QGroupID)
    ' Remember its Index and CRFElementID
    oEFQGroup.ControlIndex = nIndex
    oEFQGroup.ElementID = oElement.CRFelementID
    ' NCJ 19 Dec 01 - Initialise some values (in case we're re-drawing)
    oEFQGroup.HeaderHeight = 0
    oEFQGroup.RowWidth = 0
    Set oEFQGroup.ColWidths = New Collection
    
    ' Add to collections, keyed by nIndex and by QGroupId
    mcolGroupsByIndex.Add oEFQGroup, Str(nIndex)
    mcolGroupsByID.Add oEFQGroup, Str(oElement.QGroupID)
    
    ' Does the group have a caption?
    bGroupHasCaption = (oElement.Caption > "")
    
    With moVBForm
        Load .picQuGroup(nIndex)
        Set oPicQG = .picQuGroup(nIndex)
        Set oPicQG.Container = oContainer
        oPicQG.Left = lLeft
        oPicQG.Top = lTop
        oPicQG.ToolTipText = GetTooltipText(oElement)
        
        ' The scrollbars and Fixed pic sit inside the QuGroup pic
        Load .vsbQuGroup(nIndex)
        Set .vsbQuGroup(nIndex).Container = oPicQG
        
        Load .hsbQuGroup(nIndex)
        Set .hsbQuGroup(nIndex).Container = oPicQG
        
        Load .picFixed(nIndex)
        Set .picFixed(nIndex).Container = oPicQG
        
        ' The Scroller sits inside the Fixed pic
        Load .picScroller(nIndex)
        Set .picScroller(nIndex).Container = .picFixed(nIndex)
        
        ' The Headers sit inside the Fixed pic
        Load .picHeaders(nIndex)
        Set .picHeaders(nIndex).Container = .picFixed(nIndex)
        
        ' Hide the borders of the container pictures
        .picFixed(nIndex).BorderStyle = 0
        .picScroller(nIndex).BorderStyle = 0
        .picHeaders(nIndex).BorderStyle = 0
        
        ' Set the background colours
        oPicQG.BackColor = oContainer.BackColor
        .picFixed(nIndex).BackColor = oContainer.BackColor
        .picScroller(nIndex).BackColor = oContainer.BackColor
        .picHeaders(nIndex).BackColor = oContainer.BackColor
        
        ' Display border if required
        If oEFQGroup.Border Then
            oPicQG.BorderStyle = 1
        Else
            oPicQG.BorderStyle = 0
        End If
        
        ' Initialise the scroll bar
        .vsbQuGroup(nIndex).Enabled = False
        .vsbQuGroup(nIndex).Min = 0
        .vsbQuGroup(nIndex).LargeChange = oEFQGroup.DisplayRows
        .vsbQuGroup(nIndex).SmallChange = 1

        .hsbQuGroup(nIndex).Enabled = False
        .hsbQuGroup(nIndex).Min = 0
        .hsbQuGroup(nIndex).LargeChange = 1
        .hsbQuGroup(nIndex).SmallChange = 1
        
        ' Build all the elements
        ' Draw row 1 first and store the row width etc.
        ' in the relevant properties of the oEFQGroup object
        oEFQGroup.RowWidth = BuildGroupRow(oEFQGroup, 1, bGroupHasCaption)
        
        ' Align everything correctly
        Call ResizeGroup(oEFQGroup, nIndex)

        ' Now build all the rest of the rows
        If oEFQGroup.InitialRows > 1 Then
            For nRow = 2 To oEFQGroup.InitialRows
                Call BuildNewGroupRow(oEFQGroup, nRow, bGroupHasCaption)
            Next
        End If
        ' Remember how many rows we drew in case it changes later
        ' (when user edits the eForm Group definition)
        oEFQGroup.Rows = oEFQGroup.InitialRows
        
        ' Keep track of the max Y coordinate for non-group items
        If oElement.OwnerQGroupID = 0 Then
            If oPicQG.Top + oPicQG.Height > msglMaxY Then
                msglMaxY = oPicQG.Top + oPicQG.Height
            End If
        End If
    
        ' Finally make everything visible
         oPicQG.Visible = True
        .picFixed(nIndex).Visible = True
        .picScroller(nIndex).Visible = True
        .picHeaders(nIndex).Visible = True
        .vsbQuGroup(nIndex).Visible = True
    End With
    
    Set oPicQG = Nothing
    Set oEFQGroup = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilderSD.BuildGroup"

End Sub

'-------------------------------------------------------------------------------------
Private Function BuildGroupRow(oEFQGroup As EFormGroupSD, ByVal nRow As Integer, _
                            ByVal bGroupHasCaption As Boolean) As Single
'-------------------------------------------------------------------------------------
' Build a group row and return the width of the row
' For the first row (nRow = 1), draw the header label
' and calculate the header row height and field row height
' and store in the oEFQGroup properties.
' For subsequent rows (nRow > 1), assume relevant values available in oEFQGroup object
' bGroupHasCaption is whether the owning group has its caption displayed
'-------------------------------------------------------------------------------------
Dim sglLeft As Single
Dim sglTop As Single
Dim sglWidth As Single
Dim sglHeight As Single
Dim nEltNo As Integer
Dim oLabel As Label
Dim oElement As CRFElement
Dim nFieldIndex As Integer
Dim nGroupIndex As Integer
Dim oQGroup As QuestionGroup

    On Error GoTo ErrLabel
    
    ' Start on the left for each row
    sglLeft = msglGAP
    sglTop = oEFQGroup.RowHeight * (nRow - 1)

    nGroupIndex = oEFQGroup.ControlIndex
    Set oQGroup = frmMenu.QuestionGroups.GroupById(oEFQGroup.QGroupID)
    
    If oEFQGroup.Elements.Count = 0 Then
        ' Set up some default values for an empty group
        oEFQGroup.HeaderHeight = 250
        oEFQGroup.RowHeight = 450
        sglLeft = 1600
    Else
        For nEltNo = 1 To oEFQGroup.Elements.Count
            ' oEFQGroup.Elements contains element IDs
            Set oElement = moEForm.CRFElementById(oEFQGroup.Elements(nEltNo))
            ' The element gets drawn inside picScroller
            ' Remember its index
            nFieldIndex = BuildAnElement(oElement, nRow, moVBForm.picScroller(nGroupIndex), sglTop, sglLeft)
            ' If it's the first row, draw the header
            ' and store the column width & row height
            ' NB Include gap in row height but not in column width
            If nRow = 1 Then
                ' The label gets drawn inside picHeaders
                Set oLabel = BuildHeaderLabel(oElement, moVBForm.picHeaders(nGroupIndex), sglLeft, bGroupHasCaption)
                ' Store the greater of the field and label widths
                Call GetElementSize(oElement, nFieldIndex, sglWidth, sglHeight)
                If oLabel.Width > sglWidth Then
                    oEFQGroup.ColWidths.Add oLabel.Width
                Else
                    oEFQGroup.ColWidths.Add sglWidth
                End If
                ' Store the max header label height
                If oLabel.Height > oEFQGroup.HeaderHeight Then
                    oEFQGroup.HeaderHeight = oLabel.Height
                End If
                ' Store the greatest height of any element
                If sglHeight + msglGAP > oEFQGroup.RowHeight Then
                    oEFQGroup.RowHeight = sglHeight + msglGAP
                End If
            End If
            ' Move over for the next one
            sglLeft = sglLeft + oEFQGroup.ColWidths(nEltNo) + msglGAP
        Next
    End If
    ' Width of a row is where we got to
    BuildGroupRow = sglLeft
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilderSD.BuildGroupRow(" & nRow & ")"

End Function

'---------------------------------------------------------------------
Private Function BuildHeaderLabel(oElement As CRFElement, oContainer As PictureBox, _
                                    ByVal sglLeft As Single, _
                                    ByVal bGroupHasCaption As Boolean) As Label
'---------------------------------------------------------------------
' Build header label for a group column at given Left position
' Only show numbering if bGroupHasCaption = true
' Place it inside the given container
' Return the label control created
'---------------------------------------------------------------------
Dim sUnitOfMeasurement As String
Dim nIndex As Integer
Dim sCaption As String
Dim oLabel As Label

    nIndex = mnNextControlIndex
    mnNextControlIndex = mnNextControlIndex + 1
    
    ' Create the caption for the header row
    sCaption = ""
    If oElement.Caption > "" Then
        If moEForm.DisplayNumbers And Not oElement.Hidden And bGroupHasCaption Then
            sCaption = oElement.FieldOrder & "." & oElement.QGroupFieldOrder & " "
        End If
        sCaption = sCaption & oElement.Caption
        If oElement.Unit > "" Then
            sCaption = sCaption & " (" & oElement.Unit & ")"
        End If
    End If
    
    Load moVBForm.lblGroupHeader(nIndex)
    Set oLabel = moVBForm.lblGroupHeader(nIndex)
    
    Set oLabel.Container = oContainer
    'ZA 22/07/2002 - replace SetElementFont function with SetCaptionFont
    Call SetCaptionFont(oLabel, oElement)
    
    
    oLabel.Caption = sCaption
'    oLabel.Top = msglGAP
    oLabel.Top = 0
    oLabel.Left = sglLeft
    oLabel.BackColor = oContainer.BackColor
    'ZA 25/07/2002 - Use CaptionFontColour instead of FontColour
    oLabel.ForeColor = oElement.CaptionFontColour
    oLabel.ToolTipText = GetTooltipText(oElement)
    Call SetCommonProperties(oLabel)
    Set BuildHeaderLabel = oLabel
    
    ' This must be treated as an "independent" CRFElement
    ' Use special row no. of 0 for header labels
    Call AddElementToCollections(oElement, 0, nIndex)
    
    Set oLabel = Nothing
    
End Function

'---------------------------------------------------------------------
Private Sub BuildPushButtons(oElement As CRFElement, _
                        nIndex As Integer, oContainer As PictureBox, _
                        lTop As Long, lLeft As Long)
'---------------------------------------------------------------------
' Build push buttons for given eForm element
' REVISIONS
' DPH 24/10/2001 -  Only show active categories
'---------------------------------------------------------------------
Dim sglTop As Single
Dim oControl As Control
Dim oCategory As clsDDCategory
Dim nElementID As Integer
Dim oPic As PictureBox
Dim sglWidth As Single
Dim sglHeight As Single

    nElementID = oElement.CRFelementID

    '  Load a new picture box
    Load moVBForm.picCRFElement(nIndex)
    Set oPic = moVBForm.picCRFElement(nIndex)
    Set oPic.Container = oContainer
    
    With oPic
        .Top = lTop
        .Left = lLeft
        ' Initialize height to 0 prior to calculating it later
        ' according to category value
        .Height = 0
        Call CalculateBoxSize(oElement, sglWidth, sglHeight)
        .Width = sglWidth
        '  Set the tooltip
        .ToolTipText = GetTooltipText(oElement)
        .Visible = True
        .Enabled = True
    End With
    
    sglTop = 0
    
    ' Fill in the option buttons with the category values
    For Each oCategory In oElement.Categories
    
        ' DPH 24/10/2001 -  Only show active categories
        If oCategory.Active Then

            Load moVBForm.optGraphical(mnNextPushButtonIndex)
            Set moVBForm.optGraphical(mnNextPushButtonIndex).Container = oPic
            Set oControl = moVBForm.optGraphical(mnNextPushButtonIndex)
            
            Call SetElementFont(oControl, oElement)
            Call SetCommonProperties(oControl)
            
            With oControl
                '  Set position attributes of button
                .Top = sglTop
                .Left = 0
                .Width = oPic.Width
                .Height = sglHeight
                    
                '  Expand container
                oPic.Height = oPic.Height + .Height
                
                sglTop = sglTop + .Height
                
                ' Display the category Value
                .Caption = oCategory.Value
                ' Store the Element's index in the Tag property
                .Tag = Str(nIndex)
                ' Set the tab order
                .TabStop = True
                .TabIndex = mnTabOrder
                mnTabOrder = mnTabOrder + 1
                '  Set the tooltip
                .ToolTipText = GetTooltipText(oElement)
            
                '  Set display attributes
                .BackColor = moVBForm.picCRFPage.BackColor
                .ForeColor = oElement.FontColour
                .Visible = True
                .Enabled = True
            End With
            
            mnNextPushButtonIndex = mnNextPushButtonIndex + 1
        End If
    Next    ' oCategory
    
    ' Finally adjust the Left
    oPic.Left = AdjustHorizontalPosition(oPic.Left, oPic.Width, oContainer)
    ' Keep track of the max Y coordinate for non-group items
    If oElement.OwnerQGroupID = 0 Then
        If oPic.Top + oPic.Height > msglMaxY Then
            msglMaxY = oPic.Top + oPic.Height
        End If
    End If

End Sub

'---------------------------------------------------------------------
Private Sub BuildOptionButtons(oElement As CRFElement, _
                        nIndex As Integer, oContainer As PictureBox, _
                        lTop As Long, lLeft As Long)
'---------------------------------------------------------------------
' Build option buttons for given eForm element
' REVISIONS
' DPH 24/10/2001 - Only show active category buttons
' NCJ 27 Feb 03 - Use new CalculateOptionButtonSize
'---------------------------------------------------------------------
Dim oCats As clsDDCategories
Dim oCategory As clsDDCategory
Dim nButtonIndex As Integer
' DPH 24/10/2001 - Count active categories
Dim nActiveCat As Integer
Dim oOptGroup As OptionGroup
Dim sglW As Single
Dim sglH As Single

    Load moVBForm.optCRFGroup(nIndex)
    Set oOptGroup = moVBForm.optCRFGroup(nIndex)
    Set oOptGroup.Container = oContainer
    
    Set oCats = oElement.Categories
    
    With oOptGroup
        Call .Init
        ' The font details and size etc. will be inherited by all the buttons
        ' so we do this FIRST before setting no. of buttons
        Call SetOptionGroupFont(oOptGroup, oElement)
        .FontColor = oElement.FontColour
        .BackColor = oContainer.BackColor
        ' We don't really want highlighting in SD
        .HighLightColor = oContainer.BackColor
        .Top = lTop
        .Left = lLeft
        
        Set .MouseIcon = moVBForm.lblCRFElement(0).MouseIcon
        
        ' Set the height/width of each radio button
        ' NCJ 27 Feb 03 - Use new CalculateOptionButtonSize
'        Call CalculateBoxSize(oElement, sglW, sglH)
        Call CalculateOptionButtonSize(oElement, oCats, sglW, sglH)
        .ButtonHeight = sglH
        ' Include room for note status icon too
        .ButtonWidth = .RadioWidth + sglW + m_NOTESTATUS_IMAGE_WIDTH

        ' RS 11/02/2003: Move .Visible to end of routine, as default text flashes on screen
        ' .Visible = True
        .Enabled = True
        
        ' No. of buttons = no. of active category values
        nActiveCat = 0
        For Each oCategory In oCats
            If oCategory.Active Then
                nActiveCat = nActiveCat + 1
            End If
        Next
        If nActiveCat > 0 Then
            .Buttons = nActiveCat
        Else
            ' Set to invisible
             .Visible = False
        End If
        .TabStop = True
        ' NCJ 15 Oct 01 - Set tab index for whole group
        .TabIndex = mnTabOrder
        mnTabOrder = mnTabOrder + 1
    
    End With
    
    If nActiveCat > 0 Then
        ' Fill in the option buttons with the category values
        nButtonIndex = 0
        For Each oCategory In oCats
            
            ' DPH 24/10/2001 - Use only active categories
            If oCategory.Active Then
                oOptGroup.Caption(nButtonIndex) = oCategory.Value
                oOptGroup.TagValue(nButtonIndex) = oCategory.ValueCode
                 
                'increment the index for button each time
                nButtonIndex = nButtonIndex + 1
            End If
        Next    ' oCategory
        
        ' Finally adjust the Left
        oOptGroup.Left = AdjustHorizontalPosition(oOptGroup.Left, oOptGroup.Width, oContainer)
        
        ' Keep track of the max Y coordinate for non-group items
        If oElement.OwnerQGroupID = 0 Then
            If oOptGroup.Top + oOptGroup.Height > msglMaxY Then
                msglMaxY = oOptGroup.Top + oOptGroup.Height
            End If
        End If

        ' RS 11/03/2003: Only make visible when init done to prevent default values from flashing on screen
        oOptGroup.Visible = True

        'set the values of each option button inside the control to false
        oOptGroup.UnselectAll
    
    End If
    
    Set oOptGroup = Nothing
    Set oCategory = Nothing
    Set oCats = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub CalculateOptionButtonSize(oElement As CRFElement, oCats As clsDDCategories, _
                    ByRef sglWidth As Single, sglHeight As Single)
'---------------------------------------------------------------------
' NCJ 27 Feb 03
' Calculate the max text width & height of a category description
'---------------------------------------------------------------------
Dim oCat As clsDDCategory

    With moVBForm
        .FontName = oElement.FontName
        .Font.Charset = 1
        .FontSize = oElement.FontSize
        .FontBold = oElement.FontBold
        .FontItalic = oElement.FontItalic
        sglWidth = 0
        sglHeight = 0
        For Each oCat In oCats
            If .TextWidth(oCat.Value) > sglWidth Then
                sglWidth = .TextWidth(oCat.Value)
            End If
            If .TextHeight(oCat.Value) > sglHeight Then
                sglHeight = .TextHeight(oCat.Value)
            End If
        Next
    End With
    
    ' Add on a bit for good measure
    sglWidth = sglWidth + 60
    sglHeight = sglHeight + 60
    
End Sub

'---------------------------------------------------------------------
Private Sub BuildAttachmentButton(oElement As CRFElement, _
                        nIndex As Integer, oContainer As PictureBox, _
                        lTop As Long, lLeft As Long)
'---------------------------------------------------------------------
' Build Attachment button (multimedia) for given eForm element
'---------------------------------------------------------------------

    '  Load a new attachment button
    Load moVBForm.cmdAttachment(nIndex)
    Set moVBForm.cmdAttachment(nIndex).Container = oContainer
    With moVBForm.cmdAttachment(nIndex)
    
        '  Set position attributes
        .Top = lTop
        .Left = lLeft
        .Left = AdjustHorizontalPosition(.Left, .Width, oContainer)
        
        '  Set the tooltip
        .ToolTipText = GetTooltipText(oElement)
        
        '  Set the display attributes
        .TabStop = True
        .TabIndex = mnTabOrder
        mnTabOrder = mnTabOrder + 1
        
        Call SetCommonProperties(moVBForm.cmdAttachment(nIndex))
    
    End With

End Sub

'-------------------------------------------------------------------
Private Sub BuildCalendarButton(oLabel As Label)
'-------------------------------------------------------------------
' Build a calendar button to the right of the given label
' and with the same index
'-------------------------------------------------------------------

    Load moVBForm.cmdCalendar(oLabel.Index)
    With moVBForm.cmdCalendar(oLabel.Index)
        Set .Container = oLabel.Container
        .Top = oLabel.Top
        .Left = oLabel.Left + oLabel.Width
        .Height = oLabel.Height
        .Width = .Height
        .Visible = True
        .Enabled = True
    End With

End Sub

'-------------------------------------------------------------------
Private Sub BuildExpandButton(oLabel As Label)
'-------------------------------------------------------------------
' Build an Expand button to the right of the given label
' and with the same index
'-------------------------------------------------------------------

    Load moVBForm.cmdExpand(oLabel.Index)
    With moVBForm.cmdExpand(oLabel.Index)
        Set .Container = oLabel.Container
        .Top = oLabel.Top
        .Left = oLabel.Left + oLabel.Width
        .Height = oLabel.Height
        .Width = .Height
        .Visible = True
        .Enabled = True
    End With

End Sub

'-------------------------------------------------------------------
Private Sub BuildDictionaryButton(oLabel As Label)
'-------------------------------------------------------------------
' Build a dictionary button to the right of the given label
' and with the same index
'-------------------------------------------------------------------

    Load moVBForm.cmdDictionary(oLabel.Index)
    With moVBForm.cmdDictionary(oLabel.Index)
        Set .Container = oLabel.Container
        .Top = oLabel.Top
        .Left = oLabel.Left + oLabel.Width
        .Height = oLabel.Height
        .Width = .Height
        .Visible = True
        .Enabled = True
    End With

End Sub

'-------------------------------------------------------------------
Private Sub BuildPopupButton(oLabel As Label)
'-------------------------------------------------------------------
' Build a popuplist button to the right of the given label
' and with the same index
'-------------------------------------------------------------------

    Load moVBForm.cmdPopupList(oLabel.Index)
    With moVBForm.cmdPopupList(oLabel.Index)
        Set .Container = oLabel.Container
        .Top = oLabel.Top
        .Left = oLabel.Left + oLabel.Width
        .Height = oLabel.Height
        .Width = .Height
        .Visible = True
        .Enabled = True
    End With

End Sub

'---------------------------------------------------------------------
Private Sub BuildComment(oElement As CRFElement, _
                        nIndex As Integer, oContainer As PictureBox, _
                        lTop As Long, lLeft As Long, bHotlink As Boolean)
'---------------------------------------------------------------------
' Build comment for given eForm element
' NCJ 7 Nov 02 - If bHotlink = TRUE, it's a Hotlink (no!)
'---------------------------------------------------------------------
Dim oLabel As Label

    '  Load a new label control
    Load moVBForm.lblCaption(nIndex)
    Set oLabel = moVBForm.lblCaption(nIndex)
    Set oLabel.Container = oContainer
    
    '  Set font attributes
    'ZA 19/07/2002 - call replace SetElementFont with SetCaptionFont
    Call SetCaptionFont(oLabel, oElement)
    
    
    With oLabel
        ' Underline a hotlink
        If bHotlink Then
            .Font.Underline = True
            .ToolTipText = "Hotlink"
        Else
            .ToolTipText = "Comment"
        End If
        .Tag = oElement.Caption
        .Caption = oElement.Caption
        
        '  Set position attributes
        .AutoSize = True
        .Alignment = vbLeftJustify
        
        .Top = lTop
        .Left = lLeft
        
        .Left = AdjustHorizontalPosition(.Left, .Width, oContainer)
        '  Set the tooltip
        
        'Changed by Mo Morris 4/1/00
        'extra bit added to the AutoSize created width so that there is enough space to
        'display (and move) Italic fonts with out little bits being left on the screen
        .Width = .Width + (.FontSize * 3)
        
        '  Set display attributes
        .BackColor = oContainer.BackColor
        'ZA 25/07/2002 - use CaptionFontColour instead of FontColour
        .ForeColor = oElement.CaptionFontColour
        Call SetCommonProperties(oLabel)
    
        ' Keep track of the max Y coordinate for non-group items
        If oElement.OwnerQGroupID = 0 Then
            If .Top + .Height > msglMaxY Then
                msglMaxY = .Top + .Height
            End If
        End If
    
    End With

End Sub

'---------------------------------------------------------------------
Private Sub BuildLine(oElement As CRFElement, _
                        nIndex As Integer, oContainer As PictureBox, _
                        lTop As Long, lLeft As Long)
'---------------------------------------------------------------------
' Build line for given eForm element
'---------------------------------------------------------------------
Dim oPic As PictureBox

    '  Load a container
    Load moVBForm.picCRFElement(nIndex)
    Set oPic = moVBForm.picCRFElement(nIndex)
    Set oPic.Container = oContainer
    
    With oPic
        .Top = lTop
        .Left = 0
        .Width = oContainer.Width
        .Height = 50
        .BorderStyle = 0
        '  Set the tooltip
        .ToolTipText = "Line"
        .BackColor = oContainer.BackColor
        Call SetCommonProperties(oPic)
    End With
    
    '  Load a line and place it in the container
    Load moVBForm.Line1(nIndex)
    Set moVBForm.Line1(nIndex).Container = oPic
    
    With moVBForm.Line1(nIndex)
        .Y1 = 0
        .Y2 = 0
        .X1 = 0
        .X2 = .Container.Width
        .Visible = True
    End With

End Sub

'---------------------------------------------------------------------
Private Sub BuildTextBox(oElement As CRFElement, _
                        nIndex As Integer, oContainer As PictureBox, _
                        lTop As Long, lLeft As Long)
'---------------------------------------------------------------------
' Build text box for given eForm element
' We actually create a Label rather than a TextBox for SD
'---------------------------------------------------------------------
Dim oLabel As Label
Dim sglWidth As Single
Dim sglHeight As Single

    '  Load a new label for text box
    Load moVBForm.lblCRFElement(nIndex)
    Set oLabel = moVBForm.lblCRFElement(nIndex)
    Set oLabel.Container = oContainer
    
    '  Set font attributes
    Call SetElementFont(oLabel, oElement)
        
    With oLabel
    
        '  Set position attributes
        .Top = lTop
        .Left = lLeft
        .AutoSize = False
        
        Call CalculateBoxSize(oElement, sglWidth, sglHeight)
        .Width = sglWidth
        .Height = sglHeight
        
        .Left = AdjustHorizontalPosition(.Left, .Width, oContainer)
        
        If oElement.DataItemFormat > "" Then
            moVBForm.mskCRFElement(0).Mask = oElement.DataItemFormat
            moVBForm.mskCRFElement(0).PromptInclude = True
            .Caption = moVBForm.mskCRFElement(0).Text
        Else
            .Caption = ""
        End If

    '   ATN 18/3/99 SR732
    '   Truncate the width if it's too wide for the form
        'Test for control type POPUP_List or Calendar because they require a command button as well as
        'space for an icon to the right of the label.
        'note that the width of the command button is equal to the height of the textbox it belongs to
        'to the right of the label
        If oElement.OwnerQGroupID = 0 Then
            ' Only adjust widths for non-group items
            sglWidth = .Left
            If oElement.ShowStatusFlag Then
                ' NCJ 9 Oct 02 - Use new SpaceForStatusIcon
                sglWidth = sglWidth + SpaceForStatusIcon(oElement)
            End If
            If oElement.ControlType = gn_POPUP_LIST Or oElement.ControlType = gn_CALENDAR Then
                ' Add in width of popup/calendar button (same as height of textbox)
                sglWidth = sglWidth + .Height
            End If
            If oElement.Expandable Then
                ' RS 24/02/2003: Added Expand Button
                sglWidth = sglWidth + .Height
            End If
            
            'ic 14/06/2005 clinical coding: add browse button
            If (oElement.DataItemType = eDataType.Thesaurus) Then
                sglWidth = sglWidth + .Height
            End If
            
            If sglWidth + .Width > oContainer.Width Then
                If oContainer.Width - sglWidth < 100 Then
                    .Width = 100
                Else
                    .Width = oContainer.Width - sglWidth
                End If
            End If
        End If
        
        '  Set data entry attributes
        .TabIndex = mnTabOrder
        mnTabOrder = mnTabOrder + 1
        
        '  Set the tooltip
        .ToolTipText = GetTooltipText(oElement)
        
        Call SetCommonProperties(oLabel)
        
        '  Set the display attributes
        .BackColor = oContainer.BackColor
        .ForeColor = oElement.FontColour
        
        ' Keep track of the max Y coordinate for non-group items
        If oElement.OwnerQGroupID = 0 Then
            If .Top + .Height > msglMaxY Then
                msglMaxY = .Top + .Height
            End If
        End If
        
    End With

End Sub

'---------------------------------------------------------------------
Private Sub BuildPicture(oElement As CRFElement, _
                        nIndex As Integer, oContainer As PictureBox, _
                        lTop As Long, lLeft As Long)
'---------------------------------------------------------------------
' Build picture for given eForm element with given Index
'---------------------------------------------------------------------
Dim oPic As PictureBox

    On Error GoTo ErrLabel
    
    '  Load a new picture box
    Load moVBForm.picCRFElement(nIndex)
    Set oPic = moVBForm.picCRFElement(nIndex)
    Set oPic.Container = oContainer
    
    With moVBForm.picCRFElement(nIndex)
        '  Set position attributes
        .Top = lTop
        .Left = lLeft
        .Left = AdjustHorizontalPosition(.Left, .Width, oContainer)
        .ToolTipText = "Picture: " & oElement.Caption
        
        On Error Resume Next
        .Picture = LoadPicture(gsDOCUMENTS_PATH & oElement.Caption)
        On Error GoTo ErrLabel
        
        If Err.Number = 0 Then
            .AutoSize = True
            Call SetCommonProperties(moVBForm.picCRFElement(nIndex))
            ' Keep track of the max Y coordinate for non-group items
            If oElement.OwnerQGroupID = 0 Then
                If .Top + .Height > msglMaxY Then
                    msglMaxY = .Top + .Height
                End If
            End If
        End If
    End With
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilderSD.BuildPicture"
    
End Sub

'----------------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------------

    Call Terminate
    
End Sub

'----------------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------------

    Call Terminate

End Sub

'---------------------------------------------------------------------
Public Sub Terminate()
'---------------------------------------------------------------------
' Throw away everything before we go
'---------------------------------------------------------------------
    
    Set moVBForm = Nothing
    Set moEForm = Nothing
    Set mcolFormItems = Nothing
    Set mcolGroupsByID = Nothing
    Set mcolGroupsByIndex = Nothing
    Set mcolControlIndexes = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub BuildCaption(oElement As CRFElement, nIndex As Integer, _
                            oContainer As PictureBox)
'---------------------------------------------------------------------
' Build the caption of an eForm Element (question or group)
' Assume it's a non-group-member item
' and it may be auto-positioned or be specifically positioned
' We always build a caption but make it invisible if the Caption is empty
'---------------------------------------------------------------------
Dim sUnitOfMeasurement As String
Dim sCaption As String
Dim oLabel As Label

    '  Load a new label control
    Load moVBForm.lblCaption(nIndex)
    Set oLabel = moVBForm.lblCaption(nIndex)
    
    oLabel.Container = oContainer
    
    '  Set font attributes
    ' ZA 19/07/2002 - replce SetElementFont with SetCaptionFont
    Call SetCaptionFont(oLabel, oElement)
            
    '  Prepare unit of measurement part of label
    sUnitOfMeasurement = ""
    If oElement.Unit > "" Then
        sUnitOfMeasurement = " (" & oElement.Unit & ")"
    End If
    
    With oLabel
        If oElement.DataItemId > 0 Then
            .Tag = oElement.DataItemName
        Else    ' a group
            .Tag = oElement.Caption
        End If
        If oElement.Caption > "" Then
            If moEForm.DisplayNumbers And Not oElement.Hidden Then
                .Caption = oElement.FieldOrder & ". " & oElement.Caption & sUnitOfMeasurement
            Else
                .Caption = oElement.Caption & sUnitOfMeasurement
            End If
        Else
            .Caption = ""
        End If
        '  Set position attributes
        .AutoSize = True
        
        'Added by ASH 31/07/2001
        'Changed by Mo Morris 4/1/00
        'extra bit added to the AutoSize created width so that there is enough space to
        'display (and move) Italic fonts without little bits being left on the screen
        .Width = .Width + (.FontSize * 3)
         
        .Alignment = vbLeftJustify
        If oElement.CaptionY > 0 Then
            .Top = oElement.CaptionY
        Else
            .Top = oElement.ElementY
        
            ' RS 05/06/2003 Update oElement with new Caption Position
            ' NOTE: This does not update the database
            oElement.CaptionY = .Top
            
        End If
        If oElement.CaptionX > 0 Then
            .Left = oElement.CaptionX
        Else
            .Left = oElement.ElementX - .Width - 50
            
            ' RS 05/06/2003 Update oElement with new Caption Position
            ' NOTE: This does not update the database
            oElement.CaptionX = .Left
        End If
    
        ' RS 05/06/2003: The function below just returns the .Left parameter: no change
        .Left = AdjustHorizontalPosition(.Left, .Width, oContainer)
        
        '  Set the tooltip
        .ToolTipText = GetTooltipText(oElement)
        
        '  Set display attributes
        .BackColor = oContainer.BackColor
        'ZA 24/07/2002 - change FontColour to CaptionFontColour
        .ForeColor = oElement.CaptionFontColour
        
        Call SetCommonProperties(oLabel)
        
    '   ATN 15/10/98
    '   If no caption then make label invisible
        If .Caption > "" Then
            .Visible = True
        Else
            .Visible = False
        End If
         
    End With
    
End Sub

'---------------------------------------------------------------------
Private Function AdjustHorizontalPosition(ByVal sglLeft As Single, _
                                     ByRef sglWidth As Single, _
                                     oContainer As PictureBox) As Single
'---------------------------------------------------------------------
' Adjust horizontal position if element trespasses over RHS of screen
'---------------------------------------------------------------------

    AdjustHorizontalPosition = sglLeft
    ' NCJ 13 Nov 01 - Need to think about this!!!
    ' (what about elements which don't have status flags, or lab test questions...)
'    If sglLeft + sglWidth > oContainer.Width Then
'        sglWidth = oContainer.Width - sglLeft - 495
'    End If

End Function

'---------------------------------------------------------------------
Private Sub CalculateBoxSize(oElement As CRFElement, _
                    ByRef sglWidth As Single, ByRef sglHeight As Single)
'---------------------------------------------------------------------
' Calculate box size based on element properties
' NCJ 27 Feb 03 - For Option buttons, use CalculateOptionButtonSize INSTEAD
'---------------------------------------------------------------------
Dim nQLength As Integer

    On Error Resume Next
    
    With moVBForm
        .FontName = oElement.FontName
        .Font.Charset = 1
        .FontSize = oElement.FontSize
        .FontBold = oElement.FontBold
        .FontItalic = oElement.FontItalic
        
        ' NCJ 1 Jul 03 - Use new Expandable property
        If oElement.Expandable Then
            nQLength = oElement.DisplayLength
        Else
            nQLength = oElement.DataItemLength
        End If
        
        ' Add 4 extra characters for good measure
        sglWidth = .TextWidth(String(nQLength + 4, "_"))
        
'        'use display length only for text/popup type question
'        ' Add 4 extra characters for good measure
'        If (oElement.ControlType = gn_TEXT_BOX) Or (oElement.ControlType = gn_POPUP_LIST) Then
'            'use displaylength only if it is larger than 0
'            If oElement.DisplayLength > 0 Then
'                sglWidth = .TextWidth(String(oElement.DisplayLength + 4, "_"))
'            Else
'                sglWidth = .TextWidth(String(oElement.DataItemLength + 4, "_"))
'            End If
'        Else
'                sglWidth = .TextWidth(String(oElement.DataItemLength + 4, "_"))
'        End If
        
        ' Add 100 for good measure
        sglHeight = .TextHeight("_") + 100
        
    End With
    
End Sub

'---------------------------------------------------------------------
Private Sub SetOptionGroupFont(oOptGroup As OptionGroup, oElement As CRFElement)
'---------------------------------------------------------------------
' Set the font for the option control group for the given eForm element
' The element's own properties are correctly set up with defaults
' as appropriate in MACRO 3.0
'---------------------------------------------------------------------

    oOptGroup.FontName = oElement.FontName
    oOptGroup.FontSize = oElement.FontSize
    oOptGroup.FontBold = oElement.FontBold
    oOptGroup.FontItalic = oElement.FontItalic

End Sub

'---------------------------------------------------------------------
Private Sub SetElementFont(oControl As Control, oElement As CRFElement)
'---------------------------------------------------------------------
' Set the font for the control for the given eForm Element
'---------------------------------------------------------------------
  
    oControl.Font.Name = oElement.FontName
    oControl.Font.Charset = 1
    oControl.Font.SIZE = oElement.FontSize
    oControl.Font.Bold = oElement.FontBold
    oControl.Font.Italic = oElement.FontItalic
  
End Sub

'---------------------------------------------------------------------
Private Sub SetCaptionFont(oControl As Control, oElement As CRFElement)
'---------------------------------------------------------------------
' Set the caption font for the control for the given eForm Element
'---------------------------------------------------------------------
  
    oControl.Font.Name = oElement.CaptionFontName
    oControl.Font.Charset = 1
    oControl.Font.SIZE = oElement.CaptionFontSize
    oControl.Font.Bold = oElement.CaptionFontBold
    oControl.Font.Italic = oElement.CaptionFontItalic
    
End Sub


'---------------------------------------------------------------------
Private Function GetTooltipText(oElement As CRFElement) As String
'---------------------------------------------------------------------
' Get the tooltip text for the element
' Assume oElement is a Question or a Group (with a Name)
'---------------------------------------------------------------------
Dim sText As String

    sText = "Field: " & oElement.FieldOrder
    If oElement.OwnerQGroupID > 0 Then
        ' Append sub-number for a group element
        sText = sText & "." & oElement.QGroupFieldOrder
    End If
    If oElement.DataItemId > 0 Then
        GetTooltipText = sText & " Question: " & oElement.DataItemName
    ElseIf oElement.QGroupID > 0 Then
        GetTooltipText = sText & " Group: " & oElement.Caption
    End If

End Function

'---------------------------------------------------------------------
Private Function StatusImageLeft(oElement As CRFElement, nIndex As Integer) As Long
'---------------------------------------------------------------------
' Calculate the left coordinate for a question element's Status Image
' nIndex is the index of its controls
' ic 14/06/2005 added clinical coding
'---------------------------------------------------------------------
Dim nControlType As Integer
Dim lWidth As Long

Const lGap As Long = 50

    nControlType = oElement.ControlType
    With moVBForm
        ' Differs depending on what the element is
        Select Case nControlType
        Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX, _
                    gn_POPUP_LIST, gn_CALENDAR
            ' These all have text boxes
            lWidth = .lblCRFElement(nIndex).Left + .lblCRFElement(nIndex).Width
            If nControlType = gn_POPUP_LIST Then
                lWidth = lWidth + .cmdPopupList(nIndex).Width
            End If
            If nControlType = gn_CALENDAR Then
                lWidth = lWidth + .cmdCalendar(nIndex).Width
            End If
            ' RS 24/02/2003: Added Expand Button
            'ic 14/06/2005 clinical coding: add browse button
            If oElement.DataItemType = eDataType.Thesaurus Then
                lWidth = lWidth + .cmdDictionary(nIndex).Width
            End If

            If oElement.Expandable Then
                lWidth = lWidth + .cmdExpand(nIndex).Width
            End If
        Case gn_PUSH_BUTTONS
            lWidth = .picCRFElement(nIndex).Left + .picCRFElement(nIndex).Width
        Case gn_OPTION_BUTTONS
            lWidth = .optCRFGroup(nIndex).Left + .optCRFGroup(nIndex).Width
        Case gn_ATTACHMENT
            lWidth = .cmdAttachment(nIndex).Left + .cmdAttachment(nIndex).Width
        End Select
    End With
    
    ' Add a suitable gap
    StatusImageLeft = lWidth + lGap
    
End Function

'-------------------------------------------------------------------
Private Function StatusImageTop(oElement As CRFElement, nIndex As Integer) As Long
'-------------------------------------------------------------------
' Get the Top coordinate for the Status Image of the given control
'-------------------------------------------------------------------

    With moVBForm
        ' Differs depending on what the element is
        Select Case oElement.ControlType
        Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX, _
                    gn_POPUP_LIST, gn_CALENDAR
            ' These all have text boxes
            StatusImageTop = .lblCRFElement(nIndex).Top
        Case gn_PUSH_BUTTONS
            StatusImageTop = .picCRFElement(nIndex).Top
        Case gn_OPTION_BUTTONS
            StatusImageTop = .optCRFGroup(nIndex).Top
        Case gn_ATTACHMENT
            StatusImageTop = .cmdAttachment(nIndex).Top
        End Select
    End With

End Function

'-------------------------------------------------------------------
Private Function GetItemKey(oElement As CRFElement, nRow As Integer) As String
'-------------------------------------------------------------------
' Get the key for a given element and row
'-------------------------------------------------------------------

    GetItemKey = "K" & oElement.CRFelementID & "|" & nRow

End Function

'-------------------------------------------------------------------
Private Function GetElementSize(oElement As CRFElement, nIndex As Integer, _
                ByRef sglWidth As Single, ByRef sglHeight As Single)
'-------------------------------------------------------------------
' Get the width and height of a built eForm element (questions only)
' including its status image/NTCTC label if appropriate
' but NOT including a question's caption
' ic 14/06/2005 added clinical coding
'-------------------------------------------------------------------

    With moVBForm
            ' Differs depending on what the element is
        Select Case oElement.ControlType
        Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX, _
                    gn_POPUP_LIST, gn_CALENDAR
            ' These all have text boxes
            sglWidth = .lblCRFElement(nIndex).Width
            sglHeight = .lblCRFElement(nIndex).Height
            
            If oElement.ControlType = gn_POPUP_LIST Then
                sglWidth = sglWidth + .cmdPopupList(nIndex).Width
            End If
            
            If oElement.ControlType = gn_CALENDAR Then
                sglWidth = sglWidth + .cmdCalendar(nIndex).Width
            End If
            
            ' RS 24/02/2003: Added Expand Button
            'ic 14/06/2005 clinical coding: add browse button
            If oElement.DataItemType = eDataType.Thesaurus Then
                sglWidth = sglWidth + .cmdDictionary(nIndex).Width
            End If

            If oElement.Expandable Then
                sglWidth = sglWidth + .cmdExpand(nIndex).Width
            End If
            
        Case gn_PUSH_BUTTONS
            sglWidth = .picCRFElement(nIndex).Width
            sglHeight = .picCRFElement(nIndex).Height
            
        Case gn_OPTION_BUTTONS
            sglWidth = .optCRFGroup(nIndex).Width
            sglHeight = .optCRFGroup(nIndex).Height
            
        Case gn_ATTACHMENT
            sglWidth = .cmdAttachment(nIndex).Width
            sglHeight = .cmdAttachment(nIndex).Height

        End Select
        
        ' Add in width of status flag & lab test result label if appropriate
        sglWidth = sglWidth + SpaceForStatusIcon(oElement)

    End With
    
End Function

'-------------------------------------------------------------------
Private Sub ResizeGroup(oEFQGroup As EFormGroupSD, ByVal nIndex As Integer)
'-------------------------------------------------------------------
' Align everything in a question group when initialising or after a resize
'-------------------------------------------------------------------
Dim sglRoomForVScrollBar As Single
Dim sglMaxWidth As Single
Dim sglPicFixedWidth As Single
Dim nDisplayCols As Integer
Dim nCol As Integer
Dim oElement As CRFElement
Dim lAmountToScroll As Long

    On Error GoTo ErrLabel
    
    With moVBForm
        ' Anchor picFixed inside picQuGroup
        ' Leave a tiny gap to allow users to click there to drag the group in SD
        .picFixed(nIndex).Left = msglGAP
        .picFixed(nIndex).Top = msglGAP
                
        ' Anchor picHeaders to top left of picFixed
        .picHeaders(nIndex).Left = 0
        .picHeaders(nIndex).Top = 0
        .picHeaders(nIndex).Height = oEFQGroup.HeaderHeight + msglGAP
        .picHeaders(nIndex).Width = oEFQGroup.RowWidth
        
        ' Put picScroller below the header bar (within picFixed)
        .picScroller(nIndex).Left = 0
        .picScroller(nIndex).Top = .picHeaders(nIndex).Height
        .picScroller(nIndex).Width = oEFQGroup.RowWidth
        
        ' Fixed pic contains Header bar and Scroller
        .picFixed(nIndex).Height = .picHeaders(nIndex).Height _
                                    + oEFQGroup.RowHeight * oEFQGroup.DisplayRows - msglGAP
        
        ' Scroller pic sized to contain all Rows (even though we don't do vertical scrolling in SD)
        If oEFQGroup.InitialRows >= oEFQGroup.DisplayRows Then
            .picScroller(nIndex).Height = msglGAP + oEFQGroup.RowHeight * oEFQGroup.InitialRows
        Else
            ' No. of rows is less than DisplayRows
            .picScroller(nIndex).Height = oEFQGroup.RowHeight * oEFQGroup.DisplayRows
        End If
        
        ' ***** Now adjust the width so it fits on the screen ****
        
        ' We always leave room for the vertical scroll bar
        ' Fudge factor - have to add on 3 gaps instead of 2!
        sglRoomForVScrollBar = 3 * msglGAP + .vsbQuGroup(nIndex).Width
        
        Set oElement = moEForm.CRFElementById(oEFQGroup.ElementID)
        ' The max width we can have to fit on the screen
        sglMaxWidth = .picQuGroup(nIndex).Container.Width - .picQuGroup(nIndex).Left - SpaceForStatusIcon(oElement)
        
        ' The default width of picFixed
        sglPicFixedWidth = oEFQGroup.RowWidth
        
        ' Start with rightmost column and peel off columns until we get a fit
        nDisplayCols = oEFQGroup.Elements.Count

        ' The width of picQuGroup is picFixed width + extra bit for scroll bar
        Do While sglPicFixedWidth + sglRoomForVScrollBar >= sglMaxWidth And nDisplayCols > 1
            ' Try taking off a column from the end
            ' NB ColWidths doesn't include the gap between columns
            sglPicFixedWidth = sglPicFixedWidth - msglGAP - oEFQGroup.ColWidths(nDisplayCols)
            nDisplayCols = nDisplayCols - 1
        Loop
        
        ' Now we've got a suitable picFixed width,
        ' and nDisplayCols is how many columns we're displaying
        .picFixed(nIndex).Width = sglPicFixedWidth
        .picQuGroup(nIndex).Width = .picFixed(nIndex).Left + sglPicFixedWidth + sglRoomForVScrollBar
        
         ' Set the height of the outer picture box (without horiz. scroll bar at this stage)
        .picQuGroup(nIndex).Height = 2 * msglGAP + .picFixed(nIndex).Top + .picFixed(nIndex).Height
       
         ' Initialise the scroll bars
        ' Vert. scroll bar aligns with top of picScroller
        .vsbQuGroup(nIndex).Top = .picScroller(nIndex).Top
        .vsbQuGroup(nIndex).Height = .picFixed(nIndex).Height - .vsbQuGroup(nIndex).Top
        .vsbQuGroup(nIndex).Left = .picFixed(nIndex).Left + .picFixed(nIndex).Width + msglGAP
      
        If nDisplayCols < oEFQGroup.Elements.Count Then
            ' Need to set up some scrolling - divide scroll amount by an increment, to avoid overflow in scroll bar Max
            lAmountToScroll = (.picScroller(nIndex).Width - sglPicFixedWidth)
            ' Store the increment in the scroll bar Tag
            .hsbQuGroup(nIndex).Tag = HScrollBarIncrement(lAmountToScroll)
            .hsbQuGroup(nIndex).Max = lAmountToScroll / .hsbQuGroup(nIndex).Tag
            ' Let them click 8 times on arrows and 4 times for big jumps (this is fairly arbitrary!)
            .hsbQuGroup(nIndex).SmallChange = (.hsbQuGroup(nIndex).Max) / 8
            .hsbQuGroup(nIndex).LargeChange = (.hsbQuGroup(nIndex).Max) / 4
            
            .hsbQuGroup(nIndex).Height = .vsbQuGroup(nIndex).Width
            .hsbQuGroup(nIndex).Width = .picFixed(nIndex).Width
            .hsbQuGroup(nIndex).Top = .picFixed(nIndex).Top + .picFixed(nIndex).Height + msglGAP
            .hsbQuGroup(nIndex).Left = .picFixed(nIndex).Left + .picScroller(nIndex).Left
            .hsbQuGroup(nIndex).Value = 0
            
            ' Expand outer pic to include horiz. scroll bar
            .picQuGroup(nIndex).Height = .picQuGroup(nIndex).Height + msglGAP + .hsbQuGroup(nIndex).Height
            
            .hsbQuGroup(nIndex).Visible = True
            .hsbQuGroup(nIndex).Enabled = True
        Else
            ' No horizontal scrolling
            .hsbQuGroup(nIndex).Max = 0
            .hsbQuGroup(nIndex).Visible = False
        End If
            
    End With
    
    Set oElement = Nothing
            
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilderSD.ResizeGroup"

End Sub

'-------------------------------------------------------------------
Private Function HScrollBarIncrement(lWidthToScroll As Long) As Long
'-------------------------------------------------------------------
' Given the width to scroll, calculate the scroll bar increment
' so that the scroll bar Max will be an integer
'-------------------------------------------------------------------
Dim i As Long
Const nMAXINT = 32000   ' Don't go above 32000

    i = 1
    ' Keep dividing until we get below Max Integer value
    Do While lWidthToScroll / i > nMAXINT
        i = i + 1
    Loop
    
    HScrollBarIncrement = i
    
End Function

'-------------------------------------------------------------------
Private Sub BuildNewGroupRow(oEFQGroup As EFormGroupSD, ByVal nRow As Integer, _
                            ByVal bGroupHasCaption As Boolean)
'-------------------------------------------------------------------
' Add a new row of fields for a group
'-------------------------------------------------------------------

    Call BuildGroupRow(oEFQGroup, nRow, bGroupHasCaption)
    ' Increment the height of the scrolling box by 1 row
    moVBForm.picScroller(oEFQGroup.ControlIndex).Height = moVBForm.picScroller(oEFQGroup.ControlIndex).Height + oEFQGroup.RowHeight
    
    If nRow > oEFQGroup.DisplayRows Then
        ' Bring on the scroll bar
        With moVBForm.vsbQuGroup(oEFQGroup.ControlIndex)
            .Visible = True
            .Enabled = True
            .Max = nRow - oEFQGroup.DisplayRows
        End With
    End If

End Sub

'-------------------------------------------------------------------
' NCJ 7 Oct 02 - Removed QGroupScroll and ScrollToGroupElement because both unused in SD
'-------------------------------------------------------------------

'-------------------------------------------------------------------
Public Function GetCalPopupTopLeft(nIndex As Integer, _
                ByRef sglTop As Single, ByRef sglLeft As Single)
'-------------------------------------------------------------------
' Get the Top and Left for a Calendar or Popuplist relative to picCRFPage
' Must add group offsets if element is within a group
' NCJ 4 Nov 02 - Removed oElement argument
'-------------------------------------------------------------------

    ' NCJ 4 Nov 02 - Use GetControlTopLeft instead
    Call GetControlTopLeft(moVBForm.lblCRFElement(nIndex), sglTop, sglLeft)
    sglTop = sglTop + moVBForm.lblCRFElement(nIndex).Height
    
End Function

'-------------------------------------------------------------------
Public Sub GetControlTopLeft(oControl As Control, ByRef sglTop As Single, ByRef sglLeft As Single)
'-------------------------------------------------------------------
' Get the Top and Left for a control on the CRFPage,
' taking into account the offsets for its container(s)
' NCJ 4 Oct 02 - Changed for new group header labels
'-------------------------------------------------------------------
Dim oContainer As PictureBox

    On Error GoTo ErrLabel
    
    ' Default to its own Top, Left
    sglTop = oControl.Top
    sglLeft = oControl.Left
    Set oContainer = oControl.Container
    Select Case oContainer.Name
    Case "picCRFPage"
        ' It's on the page - nothing more to do
    Case "picScroller", "picHeaders"
        ' It's in a question group
        ' NB picScroller and picHeaders are within picFixed which is within picQuGroup
        sglLeft = sglLeft + oContainer.Left _
                + oContainer.Container.Left _
                + oContainer.Container.Container.Left
        sglTop = sglTop + oContainer.Top _
                + oContainer.Container.Top _
                + oContainer.Container.Container.Top
    Case "picQuGroup"
        ' It's sitting in a question group box
        sglTop = sglTop + oContainer.Top
        sglLeft = sglLeft + oContainer.Left
    End Select
    
    Set oContainer = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilderSD.GetControlTopLeft"
    
End Sub

'-------------------------------------------------------------------
Private Sub SetCommonProperties(oControl As Control)
'-------------------------------------------------------------------
' Set the mouse pointer, mouse icon, drag icon,
' visibility and enablement for MACRO SD
'-------------------------------------------------------------------

    With oControl
        .MouseIcon = moVBForm.lblCRFElement(0).MouseIcon
        .MousePointer = vbCustom
        .DragIcon = moVBForm.lblCRFElement(0).MouseIcon
        .Enabled = True
        .Visible = True
    End With
    
End Sub

'-------------------------------------------------------------------
Public Sub RemoveElement(oElement As CRFElement, nRow As Integer, oEFG As EFormGroupSD)
'-------------------------------------------------------------------
' Remove an element from the eForm
' i.e. undo what we did when we built it
' Automatically remove all group members if a group (and oEFG is the EFormGroup defn)
' ic 14/06/2005 added clinical coding
'-------------------------------------------------------------------
Dim oButton As OptionButton
Dim nControlType As Integer
Dim nIndex As Integer
Dim i As Integer
Dim nElt As Integer
Dim oGElement As CRFElement
Dim oLabel As Label
Dim nHeaderIndex As Integer
Dim sHeaderKey As String

    On Error GoTo ErrLabel
    
    nIndex = IndexForElementRow(oElement, nRow)
    
    If oElement.QGroupID > 0 Then
        ' It's a group
        ' Remove all the elements first from the rows that were drawn last time
        For i = 1 To oEFG.Rows
            For nElt = 1 To oEFG.Elements.Count
                ' Pick off each element definition
                Set oGElement = moEForm.CRFElementById(oEFG.Elements(nElt))
                ' Remove it
                Call RemoveElement(oGElement, i, Nothing)
            Next
        Next
        Set oGElement = Nothing
        
        ' Now remove the group paraphenalia
        For Each oLabel In moVBForm.lblGroupHeader
            ' Remove the header labels belonging to the group
            If oLabel.Index > 0 Then
                If oLabel.Container.Index = nIndex Then Unload oLabel
            End If
        Next
        
        Unload moVBForm.vsbQuGroup(nIndex)
        Unload moVBForm.hsbQuGroup(nIndex)
        Unload moVBForm.picScroller(nIndex)
        Unload moVBForm.picHeaders(nIndex)
        Unload moVBForm.picFixed(nIndex)
        Unload moVBForm.picQuGroup(nIndex)
        ' Remove from our collections
        mcolGroupsByIndex.Remove Str(nIndex)
        mcolGroupsByID.Remove Str(oElement.QGroupID)
    Else
        ' It's a single element
        nControlType = oElement.ControlType
        Select Case oElement.ControlType
        Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX, _
                    gn_POPUP_LIST, gn_CALENDAR
            ' These all have text boxes
            Unload moVBForm.lblCRFElement(nIndex)
            If nControlType = gn_POPUP_LIST Then
                Unload moVBForm.cmdPopupList(nIndex)
            End If
            If nControlType = gn_CALENDAR Then
                Unload moVBForm.cmdCalendar(nIndex)
            End If
            ' RS 24/02/2003: Added Expand Button
            'ic 14/06/2005 clinical coding: remove browse button
            If oElement.DataItemType = eDataType.Thesaurus Then
                Unload moVBForm.cmdDictionary(nIndex)
            End If
            
            If oElement.Expandable Then
                Unload moVBForm.cmdExpand(nIndex)
            End If

        Case gn_PUSH_BUTTONS
            ' Unload the matching push buttons
            For Each oButton In moVBForm.optGraphical
                If oButton.Index > 0 Then
                    If oButton.Container.Index = nIndex Then
                        Unload oButton
                    End If
                End If
            Next
            Set oButton = Nothing
            ' Then unload the container
            Unload moVBForm.picCRFElement(nIndex)
        Case gn_OPTION_BUTTONS
            Unload moVBForm.optCRFGroup(nIndex)
        Case gn_ATTACHMENT
            Unload moVBForm.cmdAttachment(nIndex)
        Case gn_COMMENT, gn_HOTLINK
            Unload moVBForm.lblCaption(nIndex)
        Case gn_PICTURE
            Unload moVBForm.picCRFElement(nIndex)
        Case gn_LINE
            Unload moVBForm.Line1(nIndex)
            Unload moVBForm.picCRFElement(nIndex)
        End Select

    End If
    
    ' Remove captions for non-group question or group items
    If (oElement.DataItemId > 0 Or oElement.QGroupID > 0) And oElement.OwnerQGroupID = 0 Then
        Unload moVBForm.lblCaption(nIndex)
    End If
    
    ' Remove it from our collections
    mcolControlIndexes.Remove GetItemKey(oElement, nRow)
    mcolFormItems.Remove Str(nIndex)
    ' For the first row of a group, also remove the header stuff
    ' (the label controls were cleared earlier)
    If oElement.OwnerQGroupID > 0 And nRow = 1 Then
        ' Headers stored as row 0
        sHeaderKey = GetItemKey(oElement, 0)
        nHeaderIndex = mcolControlIndexes(sHeaderKey)
        mcolControlIndexes.Remove sHeaderKey
        mcolFormItems.Remove Str(nHeaderIndex)
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilderSD.RemoveElement"

End Sub

'---------------------------------------------------------------------
Public Sub UnloadControls()
'---------------------------------------------------------------------
' Unload controls from form prior to rebuilding
' NCJ 23/10/00 - Clear page breaks as well
' NCJ 28/11/01 - Clear question groups too
' NCJ 8 Oct 02 - Clear picHeaders and hsbQuGroup
' ic 14/06/2005 added clinical coding
'---------------------------------------------------------------------
Dim oControl As Control
Dim oPic As PictureBox
Dim nIndex As Integer

    On Error GoTo ErrLabel

    ' Ensure that all controls used to display CRF elements are unloaded
    ' (except the base element of the control array)
    For Each oControl In moVBForm.Controls
        'ic 14/06/2005 clinical coding: remove browse button
        If oControl.Name = "cboCRFElement" _
        Or oControl.Name = "lblCRFElement" _
        Or oControl.Name = "mskCRFElement" _
        Or oControl.Name = "rtfCRFElement" _
        Or oControl.Name = "optGraphical" _
        Or oControl.Name = "optCRFElement" _
        Or oControl.Name = "cmdCalendar" _
        Or oControl.Name = "cmdExpand" _
        Or oControl.Name = "cmdAttachment" _
        Or oControl.Name = "cmdPopupList" _
        Or oControl.Name = "Line1" _
        Or oControl.Name = "lblCaption" _
        Or oControl.Name = "optCRFGroup" _
        Or oControl.Name = "vsbQuGroup" _
        Or oControl.Name = "hsbQuGroup" _
        Or oControl.Name = "lblGroupHeader" _
        Or oControl.Name = "linPageBreak" _
        Or oControl.Name = "cmdDictionary" Then

            If oControl.Index > 0 Then
                Unload oControl
            End If
        End If
    Next
    
    ' Now ensure that the picture box containers are unloaded
    ' First do picCRFElement
    ' NB We have to ensure controls are removed in the right order
    If moVBForm.picCRFElement.Count > 1 Then
        For Each oPic In moVBForm.picCRFElement
            If oPic.Index > 0 Then
                Unload oPic
            End If
        Next
    End If
    
    ' Now do the picScroller, picHeaders, picFixed and picQuGroup (they come in sets with matching indexes)
    ' (Note that picScroller and picHeaders sit inside a picFixed which sits inside a picQuGroup)
    If moVBForm.picQuGroup.Count > 1 Then
        For Each oPic In moVBForm.picQuGroup
            nIndex = oPic.Index
            If nIndex > 0 Then
                Unload moVBForm.picScroller(nIndex)
                Unload moVBForm.picHeaders(nIndex)
                Unload moVBForm.picFixed(nIndex)
                Unload oPic
            End If
        Next
    End If
'Debug.Print Timer & " Unloaded controls " & moEForm.CRFPageId

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilderSD.UnloadControls(" & moEForm.CRFPageId & ")"

End Sub

'---------------------------------------------------------------------
Private Sub SetFormZOrder()
'---------------------------------------------------------------------

    '  Make sure that ZOrder is correct
    moVBForm.picCRFPage.ZOrder 1
    moVBForm.frmCRFPage.ZOrder 1
    moVBForm.tabCRF.ZOrder 1

End Sub

'--------------------------------------------------------------------------------------------------
Public Function GroupColumnLeft(nIndex As Integer, nColumn As Integer) As Single
'--------------------------------------------------------------------------------------------------
' NCJ 2 Oct 02
' Return the Left coordinate of the given column in the eForm Group with this index
' (used in horiz scrolling)
'--------------------------------------------------------------------------------------------------
Dim i As Integer
Dim sglX As Single
Dim oEFGroup As EFormGroupSD

    Set oEFGroup = EFormQGroupByIndex(nIndex)
    
    sglX = 0
    
    If nColumn > 1 Then
        ' Add up the widths of the preceding columns
        For i = 1 To nColumn - 1
            sglX = sglX + oEFGroup.ColWidths(i) + msglGAP
        Next
    End If
    
    Set oEFGroup = Nothing
    
    GroupColumnLeft = sglX

End Function

'--------------------------------------------------------------------------------------------------
Private Function SpaceForStatusIcon(oElement As CRFElement) As Single
'--------------------------------------------------------------------------------------------------
' NCJ 7 Oct 02
' Space for status icon for an element
'--------------------------------------------------------------------------------------------------

    If oElement.ShowStatusFlag Then
        SpaceForStatusIcon = m_CHANGECOUNT_IMAGE_WIDTH + m_STATUS_IMAGE_WIDTH
        If oElement.DataItemId > 0 Then
            If oElement.DataItemType = DataType.LabTest Then
                SpaceForStatusIcon = SpaceForStatusIcon + msglGAP + m_LBL_NRCTC_WIDTH
            End If
        End If
    Else
        ' We're not showing status flags
        SpaceForStatusIcon = 0
    End If
    
End Function
