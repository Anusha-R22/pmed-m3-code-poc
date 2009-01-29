VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ElementBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       ElementBuilder.cls
'   Copyright:  InferMed Ltd. 2001-2005. All Rights Reserved
'   Author:     Nicky Johns, November 2001
'   Purpose:    Element Builder for eForm data entry for Data Management in MACRO 3.0.
'               Based on original EFormBuilder.cls from MACRO 2.2
'               Includes Question Groups
'----------------------------------------------------------------------------------------'
' This just builds the elements but doesn't do much other processing
'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 7-14 Nov 01 - Initial development
'   NCJ 15 Nov 01 - Changed control element types to be global consts (see modMACRODM_UI)
'   NCJ 23 Nov 01 - Bug fixing behaviour of Question Groups
'03/12/2001: VTRACK Changes  - Text boxes have windows background
'   NCJ 12 Dec 01 - Deal with empty groups
'   NCJ 4 Jan 02 - Status images for Question Groups
'
'   NCJ 8 Jul 02 - Brought forward all relevant 2.2 changes since 1 Nov 01
'   ZA 19/07/2002 - added functionality for font/colour for caption & comment
'TA 29/08/2002: load controls needed to display note and comment status
' MLM 30/02/02: Changed BuildElements, BuildElement and BuildTextBox to handle eform date elements.
' NCJ 3 Sep 02 - Correctly offset captions for eForm/Visit date bar
' MLM 06/09/02: Added Public Property Get EFormItems(), so that we can iterate through the Items.
' MLM 12/09/02: Tweaked the form and visit date elements so that they fit inside the header bar.
' NCJ 17 Sept 02 - Do not store DEBS objects but store IDs instead
' ZA 20/09/2002 - Use DisplayLength property when building a text/popup type question
' NCJ 24 Sept 02 - Use new Init routine of EFormItem
' TA 15/10/2002: Mores changes to the User Interface
' NCJ 4 Nov 02 - Store scroll increment in scroll bar tag
' NCJ 5 Nov 02 - "Next" button now goes at the Max X position (because of Landscape eForms)
' NCJ 8 Nov 02 - Hotlink handling
' NCJ 22 Nov 02 - No longer support RQG status display
' NCJ 3 Dec 02 - Mostly fixed RQG Tab bug (but not if last group element is option boxes)
' MLM 16/01/03: Place form and visit dates inside the header bar, so that status icons are visible.
' NCJ 17 Jan 03 - Build a "Go Back" button (see BuildNextButton)
' NCJ 21 Jan 03 - Changed cmdGoBack to picBackButton
' MLM 23/01/03: Redesign of header bar: form and visit dates are lower down, further to the left and wider.
' RS 19/02/03: Added expand button
' NCJ 27 Feb 03 - Replaced calculation of option button size
' NCJ 1 Apr 03 - Bug 1327 - Must not draw hidden questions in RQGs!
' NCJ 21 May 03 - Added some Option Buttons Fudge Factors for BUG 837
' NCJ 22 May 03 - Make sure Next/Previous arrows line up with popup/calendar buttons
' MLM 26/06/03: 3.0 bug list 1338: If a study defines captions for form or visit dates, display these in place of the fixed captions previously used.
' NCJ 1 Jul 03 - Use new Expandable property of Element (Bug 1905)
' NCJ 23 May 05 (and 30 Jun 05) - Bug 2513 - Ignore groups where everything is hidden (BuildElement)
'   TA  1/7/2005    set font.charset to allow non western european characters
' ic 14/07/2005 added clinical coding
'----------------------------------------------------------------------------------------'

Option Explicit

'TA 02/09/2002: height of changecount and status icons
Private Const m_CHANGECOUNT_STATUS_IMAGE_HEIGHT = 240

'TA 02/09/2002: width of changecount icon
Private Const m_CHANGECOUNT_IMAGE_WIDTH = 120

'TA 02/09/2002: width of image status icon
Private Const m_STATUS_IMAGE_WIDTH = 240

'TA 02/09/2002: width of note status icon
Private Const m_NOTESTATUS_IMAGE_WIDTH = 120
Private Const m_NOTESTATUS_IMAGE_HEIGHT = 120

' The VB form we're drawing to
' NB Assume this is an instance of frmFormDataEntry
Private moVBForm As Form

' NCJ 17 Sept 02 - We store the Subject and get everything from here
Private moSubject As StudySubject

' The eForm definition
Private moEForm As eFormRO

' The eFormInstance ID (don't store the object itself)
Private mlEFormTaskId As Long
'Private moEFI As EFormInstance

Private mnNextControlIndex As Integer
Private mnNextPushButtonIndex As Integer

Private mnTabOrder As Integer

Private Const mn_SPACE_FOR_STATUS_ICON = 495

' NCJ 9 July 02 - Created new constants
Private Const msglMIN_EFORM_HEIGHT As Single = 20000
Private Const msglEXTRA_DROPDOWN_HEIGHT As Single = 4500

Private Const msglGAP As Single = 50

Private mcolControlIndexes As Collection    ' of integer indexes
Private mcolFormItems As Collection     ' of EFormItem objects

Private mcolGroupsByIndex As Collection    ' of eFormQGroupDE objects
Private mcolGroupsByID As Collection    ' of eFormQGroupDE objects

' Max Y coord on eForm
Private msglMaxY As Single
' X coord of last eForm item, i.e. item furthest down the eForm
Private msglMaxX As Single

' NCJ 3 Sept 02 - Offset for eForm Elements
Private msglYOffset As Single

'--------------------------------------------------------------------------
Public Property Get EFI() As EFormInstance
'--------------------------------------------------------------------------
' NCJ 17 Sept 02
' The eFormInstance for this eForm item,
' retrieved by EFormTaskId
'--------------------------------------------------------------------------

    Set EFI = moSubject.eFIByTaskId(mlEFormTaskId)

End Property

'----------------------------------------------------------
Public Sub BuildElements(oVBForm As Form, oSubject As StudySubject, oEFI As EFormInstance)
'----------------------------------------------------------
' Draw all the form elements
' Assume that the eForm has its EFormElements collection all ready
' and that the eForm is empty of all other controls
'
' MLM 30/08/02: Help draw the new header bar containing the form date.
' MLM 12/09/02: Make sure the next form button is drawn below the header bar on empty forms
' MLM 16/01/03: Place form and visit dates inside the header bar, so that status icons are visible.
' MLM 23/01/03: Reposition form and visit dates and display labels for them.
'----------------------------------------------------------
Dim oElement As eFormElementRO
Dim sglEFormHeight As Single
Dim bHasLabTest As Boolean

'MLM 04/09/02:
Dim oVisitDateElement As eFormElementRO
'MLM 07/01/03:
'Dim oVisitDateVEF As VisitEFormRO
Dim oVisitEFI As EFormInstance

    On Error GoTo ErrLabel
    
    Set moVBForm = oVBForm
    Set moSubject = oSubject
'    Set moEFI = oEFI
    ' NCJ 17 Sept 02 - Store the ID, not the EFI itself
    mlEFormTaskId = oEFI.EFormTaskId
    
    Set moEForm = oEFI.eForm
    
    ' Do we know our forms?
    If (moEForm Is Nothing) Or (moVBForm Is Nothing) Then
        Exit Sub
    End If

    ' Have we got some eForm Elements to work from?
    If moEForm.EFormElements Is Nothing Then
        Exit Sub
    End If

    ' Initialise everything
    Set mcolControlIndexes = New Collection
    Set mcolFormItems = New Collection
    Set mcolGroupsByIndex = New Collection
    Set mcolGroupsByID = New Collection
    
    mnNextControlIndex = 1
    mnNextPushButtonIndex = 1
    mnTabOrder = 0
    
    ' NCJ 3 Sept 02 - Offset for the eForm/Visit date bar
    msglYOffset = moVBForm.picHeaderBar.Height
    
    'MLM 23/01/03: Hide these initially, then redisplay them if they're required.
    moVBForm.lblVisitDate.Visible = False
    moVBForm.lblEFormDate.Visible = False
    
    bHasLabTest = False
    For Each oElement In moEForm.EFormElements
        If oElement.DataType = eDataType.LabTest Then
            bHasLabTest = True
            Exit For
        End If
    Next
    
    If bHasLabTest Then
        moVBForm.picLab.Top = moVBForm.picHeaderBar.Height
        moVBForm.picLab.Left = moVBForm.picHeaderBar.Left
        moVBForm.picLab.Width = moVBForm.picHeaderBar.Width
        moVBForm.picLab.Visible = True
        msglYOffset = msglYOffset + moVBForm.picLab.Height
        moVBForm.picHeaderExt.Height = moVBForm.picHeaderBar.Height + moVBForm.picLab.Height
    Else
        moVBForm.picLab.Visible = False
        moVBForm.picHeaderExt.Height = moVBForm.picHeaderBar.Height
    End If
    
    
    'MLM 12/09/02: Make sure the next form button is drawn below the header bar on empty forms
    msglMaxY = msglYOffset
    ' NCJ 5 Nov 02 - Initialise Max X coordinate
    msglMaxX = msglGAP + moVBForm.cmdNextButton.Width
    
    ' Draw the "top level" eForm Elements (i.e. not those inside groups)
    ' (they get built by the owner group)
    For Each oElement In moEForm.EFormElements
        'MLM 30/08/02: Catch the eForm date element, and draw it in the header bar instead of the position specified in SD
        If oElement.ElementUse = EFormVisitDate Then
            'MLM 16/01/03: place the form date inside picHeaderBar
            Call BuildElement(oEFI, oElement, 1, moVBForm.picHeaderBar, _
                            437, 5458) 'MLM 23/01/03: New coordinates of eForm date
                            '62, 6386) 'twip coordinates of the eForm date element
            'MLM 26/06/03: This is now done inside BuildCaption:
            'moVBForm.lblEFormDate.Visible = True
        ElseIf oElement.OwnerQGroup Is Nothing Then
            'draw the other elements slightly lower down, to leave room for the header bar
            Call BuildElement(oEFI, oElement, 1, moVBForm.picCRFPage, _
                            oElement.ElementY + msglYOffset, oElement.ElementX)
        End If
    Next
    
    'MLM 04/09/02: Also add an element for the visit date, if it's needed
    'building the visit date last means it will be first in the tab order
    'MLM 07/01/03: Test for the presence of the visit date using the subject's objects rather than the SD,
    'so that we don't display a visit date on an unscheduled eForm.
    Set oVisitEFI = oEFI.VisitInstance.VisitEFormInstance
'    If Not oVisitDateVEF Is Nothing Then
    If Not oVisitEFI Is Nothing Then
        'the visit has a visit eform
        'the visit instance has a visit eform.
        Set oVisitDateElement = oVisitEFI.eForm.EFormDateElement
        If Not oVisitDateElement Is Nothing Then
            'the visit eform has a date q
            'MLM 16/01/03: place the visit date inside picHeaderBar
            Call BuildElement(oVisitEFI, oVisitDateElement, 1, moVBForm.picHeaderBar, _
                            437, 1200) 'MLM 23/01/03: New coordinates of visit date
                            '62, 2129) 'twip coordinates of the visit date element
            'MLM 26/06/03: This is now done inside BuildCaption:
            'moVBForm.lblVisitDate.Visible = True
            Set oVisitDateElement = Nothing
        End If
        Set oVisitEFI = Nothing
    End If
    
    ' NCJ 8 Jul 02 - Build "Next eForm" button (as in 2.2)
    BuildNextButton oEFI.ReadOnly
    
    ' Set form height
    
    'TA 7/2/2003: don't make form any longer than the right hand arrow
    moVBForm.picCRFPage.Height = msglMaxY + 60
    
''    'ZA 27/05/2002 - Add in extra height for drop-downs (4500) to display drop down question properly
'''    sglEFormHeight = msglMaxY + msglEXTRA_DROPDOWN_HEIGHT
''    ' Make sure height is at least the minimum
'' '   If sglEFormHeight < msglMIN_EFORM_HEIGHT Then
'' '       moVBForm.picCRFPage.Height = msglMIN_EFORM_HEIGHT
'' '   Else
'' '       moVBForm.picCRFPage.Height = sglEFormHeight
'' '   End If
    
    '  Make sure popup list is hidden
    moVBForm.lvwPopupList.Visible = False
    '   Show the form after the controls have been built
    moVBForm.picCRFPage.Visible = True
    
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilder.BuildElements"

End Sub

'---------------------------------------------------------------------
Public Function IndexForElementRow(oElement As eFormElementRO, nRow As Integer) As Integer
'---------------------------------------------------------------------
' Returns control array index used for this element/row
'---------------------------------------------------------------------

    IndexForElementRow = mcolControlIndexes(GetItemKey(oElement, nRow))

End Function

'---------------------------------------------------------------------
Public Function EFormItemByIndex(nIndex As Integer) As EFormItem
'---------------------------------------------------------------------
' Returns the element/row represented by this control array index
' NCJ 8 Nov 02 - Return Nothing if none exists (necessary fro Comment/Hotlink handling)
'---------------------------------------------------------------------

    On Error GoTo Ignore
    Set EFormItemByIndex = mcolFormItems(str(nIndex))

Ignore:

End Function

'---------------------------------------------------------------------
Private Function BuildElement(oEFI As EFormInstance, oElement As eFormElementRO, _
                            nRow As Integer, oContainer As PictureBox, _
                            ByVal lTop As Long, ByVal lLeft As Long) As EFormItem
'---------------------------------------------------------------------
' Build a single eForm Element onto the VB form
' The nRow is which group row it's in (default = 1 for non-group items)
' Return the index of the element's control(s)
'
' MLM 30/08/02: Don't show the eForm date's caption
' MLM 06/09/02: Pass in an EFI to BuildElement, and use to store the
'   response as well as the element in the EFormItems collection.
' MLM 09/09/02: BuildElement now returns the EFormItem that it creates
'   (for non-hidden groups and questions only).
' TA 6/1/03: tooltip text for control now set here rather than in indivdual build code
' NCJ 23 May 05 - Ignore groups where everything is hidden
' ic 14/07/2005 added clinical coding
'---------------------------------------------------------------------
Dim nControlType As Integer
Dim nIndex As Integer
Dim oEFormItem As EFormItem
Dim oCon As Control
Dim oQGElement As eFormElementRO
Dim bGroupHidden As Boolean

    On Error GoTo ErrLabel
    
    If oElement.Hidden Then
        Exit Function
    End If
    
    ' NCJ 23 May 2005 - Bug 2513
    If Not oElement.QGroup Is Nothing Then
        ' Check for all group items being hidden
        bGroupHidden = True
        For Each oQGElement In oElement.QGroup.Elements
            ' If at least one non-hidden, we carry on
            If Not oQGElement.Hidden Then
                bGroupHidden = False
                Exit For
            End If
        Next
        ' If they're all hidden, don't bother
        If bGroupHidden Then Exit Function
    End If
    
    nControlType = oElement.ControlType
    nIndex = mnNextControlIndex
    mnNextControlIndex = mnNextControlIndex + 1

    If oElement.QuestionId > 0 And oElement.QGroup Is Nothing Then
        ' It's a plain data item
        
        'MLM 30/08/02: Don't show the eForm date's caption
        'MLM 24/06/03: Actually, do show it. BuildCaption will display them specially.
        If oElement.OwnerQGroup Is Nothing Then ' And oElement.ElementUse <> EFormVisitDate Then
            ' Build caption for non-group questions only
            BuildCaption oElement, nIndex, oContainer
        End If
        
        ' Build the rest of it according to its type
        Select Case nControlType
        Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX, _
                gn_POPUP_LIST, gn_CALENDAR
                
            BuildTextBox oElement, nIndex, oContainer, lTop, lLeft
            
            'RS 19/02/2003: Add Expand Button (only for gn_TEXT_BOX)
            ' NCJ 1 Jul 03 - Use new Element property
'            If nControlType = gn_TEXT_BOX And oElement.DisplayLength > 0 Then
            If oElement.Expandable Then
                Call BuildExpandButton(moVBForm.txtCRFElement(nIndex))
            End If
            
            'ic 23/02/2007 bug 2855 for DE dont display clinical coding button if clinical coding
            'is not installed, just display as a textbox. do display the button for SD
            If Not ((App.Title = "MACRO_DM") And (Not gbClinicalCoding)) Then
                'ic 14/07/2005 added clinical coding
                If (oElement.DataType = eDataType.Thesaurus) Then
                    Call BuildDictionaryButton(moVBForm.txtCRFElement(nIndex))
                    If (goUser.CheckPermission(gsFnCodeClinicalResponse)) Then
                        moVBForm.imgDictionary(nIndex).Enabled = True
                    Else
                        moVBForm.imgDictionary(nIndex).Enabled = False
                    End If
                End If
            End If
            
            If nControlType = gn_POPUP_LIST Then
                Call BuildPopupButton(moVBForm.txtCRFElement(nIndex))
            End If
            
            If nControlType = gn_CALENDAR Then
                Call BuildCalendarButton(moVBForm.txtCRFElement(nIndex))
            End If
            
        Case gn_PUSH_BUTTONS
            BuildPushButtons oElement, nIndex, oContainer, lTop, lLeft
            
        Case gn_OPTION_BUTTONS
            BuildOptionButtons oElement, nIndex, oContainer, lTop, lLeft
            
        Case gn_ATTACHMENT
            BuildAttachmentButton oElement, nIndex, oContainer, lTop, lLeft
            
        End Select
            
        ' Add the image for the element's status icon as required
         BuildImage oElement, nIndex, oContainer

    ElseIf Not oElement.QGroup Is Nothing Then
        ' Build its caption
        BuildCaption oElement, nIndex, oContainer
        ' Draw the group
        Call BuildGroup(oElement, nIndex, oContainer, lTop, lLeft)
        ' Add the image for the status icon if required
        ' NCJ 22 Nov 02 - Don't have group statuses any more
'        BuildImage oElement, nIndex, oContainer
        
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
    
    ' If it's a question or a group, store in our collections
    ' NCJ 8 Nov 02 - Also store Hotlinks as eFormItems
    If oElement.QuestionId > 0 Or (Not oElement.QGroup Is Nothing) Or (nControlType = gn_HOTLINK) Then
        Set oEFormItem = New EFormItem
        Call oEFormItem.Init(moSubject, oEFI, oElement)
        'MLM 06/09/02: If the item has a response, remember it (there may not be a response)
        If oElement.QuestionId > 0 Then
            Set oEFormItem.Response = oEFI.Responses.ResponseByElement(oElement, nRow)
        End If
                
        oEFormItem.RowNo = nRow
        oEFormItem.ControlIndex = nIndex
        mcolControlIndexes.Add nIndex, GetItemKey(oElement, nRow)
        mcolFormItems.Add oEFormItem, str(nIndex)
        
    End If

    'MLM 09/09/02: Return the element's item.
    Set BuildElement = oEFormItem
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & _
            "|ElementBuilder.BuildElement(" & oElement.ElementID & "," & nRow & ")"

End Function

'---------------------------------------------------------------------
Public Function EFormQGroupByIndex(nIndex As Integer) As EFormQGroupDE
'---------------------------------------------------------------------
' Return the eFormQGroup object corresponding to the given Control Index
'---------------------------------------------------------------------

    Set EFormQGroupByIndex = mcolGroupsByIndex.Item(str(nIndex))
    
End Function

'---------------------------------------------------------------------
Public Function EFormQGroupById(lQGroupId As Long) As EFormQGroupDE
'---------------------------------------------------------------------
' Return the eFormQGroup object corresponding to the given Group ID
'---------------------------------------------------------------------

    Set EFormQGroupById = mcolGroupsByID.Item(str(lQGroupId))
    
End Function
 
'---------------------------------------------------------------------
Private Sub BuildGroup(oElement As eFormElementRO, _
                        nIndex As Integer, oContainer As PictureBox, _
                        lTop As Long, lLeft As Long)
'---------------------------------------------------------------------
' Build a question group
' NCJ 27 Nov 02 - Store NextTabIndex after building group
' NCJ 3 Dec 02 - Instead store last item built and whether group is built
'---------------------------------------------------------------------
Dim oPicQG As PictureBox
Dim oQGroup As QGroupRO
Dim oEFQGroup As EFormQGroupDE
Dim nRow As Integer

    ' Create new EFormGroup object
    Set oEFQGroup = New EFormQGroupDE
    Call oEFQGroup.Init(moSubject, EFI, oElement.QGroup)
    
    oEFQGroup.ControlIndex = nIndex
    
    ' Add to collections, keyed by nIndex and by QGroupId
    mcolGroupsByIndex.Add oEFQGroup, str(nIndex)
    mcolGroupsByID.Add oEFQGroup, str(oElement.QGroup.QGroupId)
    
    With moVBForm
        Load .picQuGroup(nIndex)
        Set oPicQG = .picQuGroup(nIndex)
        Set oPicQG.Container = oContainer
        oPicQG.Left = lLeft
        oPicQG.Top = lTop
        oPicQG.Visible = True
        
        ' The scrollbars and Fixed pic sit inside the QuGroup pic
        Load .vsbQuGroup(nIndex)
        Set .vsbQuGroup(nIndex).Container = oPicQG
        
        Load .hsbQuGroup(nIndex)
        Set .hsbQuGroup(nIndex).Container = oPicQG
        
        Load .picFixed(nIndex)
        Set .picFixed(nIndex).Container = oPicQG
        .picFixed(nIndex).Visible = True
        
        ' The Scroller sits inside the Fixed pic
        Load .picScroller(nIndex)
        Set .picScroller(nIndex).Container = .picFixed(nIndex)
        .picScroller(nIndex).Visible = True
        
        ' The Headers sit inside the Fixed pic
        Load .picHeaders(nIndex)
        Set .picHeaders(nIndex).Container = .picFixed(nIndex)
        .picHeaders(nIndex).Visible = True
        
        ' Hide the borders of the container pictures
        .picFixed(nIndex).BorderStyle = 0
        .picScroller(nIndex).BorderStyle = 0
        .picHeaders(nIndex).BorderStyle = 0
        
        ' Set the background colours
        oPicQG.BackColor = oContainer.BackColor
        .picFixed(nIndex).BackColor = oContainer.BackColor
        .picScroller(nIndex).BackColor = oContainer.BackColor
        .picHeaders(nIndex).BackColor = oContainer.BackColor
        
        ' Get the Question Group definition
        Set oQGroup = oElement.QGroup
        
        ' Display border if required
        If oQGroup.Border Then
            oPicQG.BorderStyle = 1
        Else
            oPicQG.BorderStyle = 0
        End If
        
        ' Initialise the scroll bar
        .vsbQuGroup(nIndex).Visible = True
        .vsbQuGroup(nIndex).Enabled = False
        .vsbQuGroup(nIndex).Min = 0
        .vsbQuGroup(nIndex).LargeChange = oQGroup.DisplayRows
        .vsbQuGroup(nIndex).SmallChange = 1

        .hsbQuGroup(nIndex).Visible = False
        .hsbQuGroup(nIndex).Enabled = False
        .hsbQuGroup(nIndex).Min = 0
        .hsbQuGroup(nIndex).Max = 0
        
        ' Build all the elements
        ' Draw row 1 first and store the row width etc.
        ' in the relevant properties of the oEFQGroup object
        oEFQGroup.RowWidth = BuildGroupRow(oEFQGroup, 1)
        
        ' Align everything correctly
        Call ResizeGroup(oEFQGroup, nIndex)

        ' Now build all the rest of the rows
        If oEFQGroup.CurrentRows > 1 Then
            For nRow = 2 To oEFQGroup.CurrentRows
                Call BuildNewGroupRow(oEFQGroup, nRow)
            Next
        End If
        
        ' Say that we've now done the initial build
        oEFQGroup.GroupIsBuilt = True

        ' Keep track of the max Y coordinate for non-group items
        If oElement.OwnerQGroup Is Nothing Then
            If oPicQG.Top + oPicQG.Height > msglMaxY Then
                msglMaxY = oPicQG.Top + oPicQG.Height
            End If
            ' NCJ 5 Nov 02 - Also remember the Max X coord
            If oPicQG.Left + oPicQG.Width > msglMaxX Then
                msglMaxX = oPicQG.Left + oPicQG.Width
            End If
        End If
    
    End With
    
    Set oPicQG = Nothing
    Set oEFQGroup = Nothing
    Set oQGroup = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilder.BuildGroup"

End Sub

'-------------------------------------------------------------------------------------
Private Function BuildGroupRow(oEFQGroup As EFormQGroupDE, nRow As Integer) As Single
'-------------------------------------------------------------------------------------
' Build a group row and return the width of the row
' For the first row (nRow = 1), draw the header label
' and calculate the header row height and field row height
' and store in the oEFQGroup properties.
' For subsequent rows (nRow > 1), assume relevant values available in oEFQGroup object
'
' MLM 09/09/02: BuildElement returns an EFormItem, which we add to the EFormQGroupDE's EFormItems collection.
' NCJ 1 Apr 03 - Must not draw hidden questions!
'-------------------------------------------------------------------------------------
Dim sglLeft As Single
Dim sglTop As Single
Dim sglWidth As Single
Dim sglHeight As Single
Dim nEltNo As Integer
Dim oLabel As Label
Dim oElement As eFormElementRO
'MLM 09/09/02:
'Dim nFieldIndex As Integer
Dim oItem As EFormItem
Dim nGroupIndex As Integer
Dim oQGroup As QGroupRO
Dim oLastItem As EFormItem

    On Error GoTo ErrLabel
    
    ' Start on the left for each row
    sglLeft = msglGAP
    sglTop = oEFQGroup.RowHeight * (nRow - 1)

    nGroupIndex = oEFQGroup.ControlIndex
    Set oQGroup = oEFQGroup.QGroup
    
    If oQGroup.Elements.Count = 0 Then
        ' Set up some default values for an empty group
        oEFQGroup.HeaderHeight = 250
        oEFQGroup.RowHeight = 450
        sglLeft = 1600
    Else
        For nEltNo = 1 To oQGroup.Elements.Count
            Set oElement = oQGroup.Elements(nEltNo)
            ' The element gets drawn inside picScroller
            ' Remember its index
            ' MLM 06/09/02: Assume that if an element's in a group, it's in moEFI (the EFI could be discovered properly from oEFQGroup.QGroupInstance, if it had a Property Get EFI)
            ' MLM 09/09/02: BuildElement returns an EFormItem; add it to the group's collection
            Set oItem = BuildElement(EFI, oElement, nRow, moVBForm.picScroller(nGroupIndex), sglTop, sglLeft)
            If Not oItem Is Nothing Then
                ' Keep track of the last one we did
                Set oLastItem = oItem
                'only for questions in the group
                oEFQGroup.EFormItems.Add oItem, str(oItem.ControlIndex)
                ' If it's the first row, draw the header
                ' and store the column width & row height
                ' NB Include gap in row height but not in column width
                If nRow = 1 Then
                    ' NCJ 10 Oct 02 - Store the left coord of this one
                    oEFQGroup.ColXs.Add sglLeft
                    ' The label gets drawn inside picHeaders
                    Set oLabel = BuildHeaderLabel(oElement, moVBForm.picHeaders(nGroupIndex), sglLeft)
                    ' Store the greater of the field and label widths
                    Call GetElementSize(oElement, oItem.ControlIndex, sglWidth, sglHeight)
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
            Else
                ' NCJ 1 Apr 03 - There's no oItem so element must be hidden
                ' Add default values to X and Widths collections
                oEFQGroup.ColWidths.Add 0
                If nRow = 1 Then
                    oEFQGroup.ColXs.Add sglLeft
                End If
            End If
        Next
        
        ' NCJ 3 Dec 02 - Store last item built (so we can later get its tab order)
        Set oEFQGroup.LastItem = oLastItem
        ' If it's push buttons, we need to know the index of the last push button created
        If Not oLastItem Is Nothing Then
            If oLastItem.Element.ControlType = gn_PUSH_BUTTONS Then
                oEFQGroup.LastPushButtonIndex = mnNextPushButtonIndex - 1
            End If
        End If
        
    End If
    ' Width of a row is where we got to
    BuildGroupRow = sglLeft
    
    Set oItem = Nothing
    Set oLastItem = Nothing
    Set oElement = Nothing
    Set oQGroup = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilder.BuildGroupRow(" & nRow & ")"

End Function

'---------------------------------------------------------------------
Private Function BuildHeaderLabel(oElement As eFormElementRO, oContainer As PictureBox, _
                                    sglLeft As Single) As Label
'---------------------------------------------------------------------
' Build header label for a group column at given Left position
' Place it inside the given container
' Return the label control created
'---------------------------------------------------------------------
Dim sCaption As String
Dim nIndex As Integer

    nIndex = mnNextControlIndex
    mnNextControlIndex = mnNextControlIndex + 1
    
    With moVBForm
        Load .lblCaption(nIndex)
        Set .lblCaption(nIndex).Container = oContainer
        'ZA 22/07/2002 - replace SetElementFont function with SetCaptionFont
        Call SetCaptionFont(.lblCaption(nIndex), oElement)
        
        If oElement.Unit > "" Then
            sCaption = oElement.Caption & " (" & oElement.Unit & ")"
        Else
            sCaption = oElement.Caption
        End If
        ' Only add numbering if the group's caption is visible
        ' and we're displaying numbers
        If moEForm.DisplayNumbers And .lblCaption(oContainer.Index).Visible Then
            sCaption = oElement.ElementOrder & "." & oElement.GroupElementOrder & " " & sCaption
        End If
        .lblCaption(nIndex).Caption = sCaption
        .lblCaption(nIndex).Top = 0
        .lblCaption(nIndex).Left = sglLeft
        .lblCaption(nIndex).Visible = (oElement.Caption > "")
        'ZA 25/07/2002 - use CaptionFontColour instead of FontColour
        .lblCaption(nIndex).ForeColor = oElement.CaptionFontColour
        
        Set BuildHeaderLabel = .lblCaption(nIndex)
    End With
    
End Function

'---------------------------------------------------------------------
Private Sub BuildPushButtons(oElement As eFormElementRO, _
                        nIndex As Integer, oContainer As PictureBox, _
                        lTop As Long, lLeft As Long)
'---------------------------------------------------------------------
' Build push buttons for given eForm element
' REVISIONS
' DPH 24/10/2001 -  Only show active categories
'---------------------------------------------------------------------
Dim sglTop As Single
Dim oControl As Control
Dim oCategory As CategoryItem
Dim nElementId As Integer
Dim oPic As PictureBox
Dim sglWidth As Single
Dim sglHeight As Single

    nElementId = oElement.ElementID

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
        ' NCJ 4 Sept 02 - Add in a bit of space for note status icon and set BackColor
        .Width = sglWidth + m_NOTESTATUS_IMAGE_WIDTH
        .BackColor = oContainer.BackColor
        .Visible = True
        .Enabled = False
    End With
    
    sglTop = 0
    
    ' DPH 24/10/2001 - Check if categories
    If oElement.ActiveCategories Then
        ' Fill in the option buttons with the category values
        For Each oCategory In oElement.Categories
        
            ' DPH 24/10/2001 -  Only show active categories
            If oCategory.Active Then
    
                Load moVBForm.optGraphical(mnNextPushButtonIndex)
                Set moVBForm.optGraphical(mnNextPushButtonIndex).Container = oPic
                Set oControl = moVBForm.optGraphical(mnNextPushButtonIndex)
                
                Call SetElementFont(oControl, oElement)
                
                With oControl
                    '  Set position attributes of button
                    .Top = sglTop
                    .Left = 0
                    .Width = oPic.Width - m_NOTESTATUS_IMAGE_WIDTH
                    .Height = sglHeight
                        
                    '  Expand container
                    oPic.Height = oPic.Height + .Height
                    
                    sglTop = sglTop + .Height
                    
                    ' Display the category Value
                    .Caption = oCategory.Value
                    ' Store the Code in the Tag property
                    .Tag = oCategory.Code
                    ' Set the tab order
                    .TabStop = True
                    .TabIndex = mnTabOrder
                    mnTabOrder = mnTabOrder + 1

                
                    '  Set display attributes
                    .BackColor = moVBForm.picCRFPage.BackColor
                    .ForeColor = oElement.FontColour
                    .Visible = True
                    .Enabled = False
                End With
                
                mnNextPushButtonIndex = mnNextPushButtonIndex + 1
            End If
        Next    ' oCategory
        
        ' Finally adjust the Left
        oPic.Left = AdjustHorizontalPosition(oPic.Left, oPic.Width, oContainer)
        ' Keep track of the max Y coordinate for non-group items
        If oElement.OwnerQGroup Is Nothing Then
            If oPic.Top + oPic.Height > msglMaxY Then
                msglMaxY = oPic.Top + oPic.Height
            End If
            ' NCJ 5 Nov 02 - Also remember the Max X coord
            If oPic.Left + oPic.Width > msglMaxX Then
                msglMaxX = oPic.Left + oPic.Width
            End If
        End If

    Else
        ' No active categories
        oPic.Visible = False
    End If

End Sub

'---------------------------------------------------------------------
Private Sub BuildOptionButtons(oElement As eFormElementRO, _
                        nIndex As Integer, oContainer As PictureBox, _
                        lTop As Long, lLeft As Long)
'---------------------------------------------------------------------
' Build option buttons for given eForm element
' REVISIONS
' DPH 24/10/2001 - Only show active category buttons
' NCJ 27 Feb 03 - Replaced calculation of option button size
'---------------------------------------------------------------------
Dim oControl As Control
Dim oCategory As CategoryItem
Dim nButtonIndex As Integer
' DPH 24/10/2001 - Count active categories
Dim nActiveCat As Integer
Dim oOptGroup As OptionGroup
Dim sglW As Single
Dim sglH As Single

    Load moVBForm.optCRFGroup(nIndex)
    Set oOptGroup = moVBForm.optCRFGroup(nIndex)
    Set oOptGroup.Container = oContainer
    
    With oOptGroup
        Call .Init
        ' The font details and size etc. will be inherited by all the buttons
        ' so we do this FIRST before setting no. of buttons
        Call SetOptionGroupFont(oOptGroup, oElement)
        .FontColor = oElement.FontColour
        .BackColor = oContainer.BackColor
        .HighLightColor = glYELLOW
        .Top = lTop
        .Left = lLeft
        ' Set the height/width of each radio button
        ' NCJ 27 Feb 03 - Use new calculation of option button size
'        Call CalculateBoxSize(oElement, sglW, sglH)
        Call CalculateOptionButtonSize(oElement, sglW, sglH)
        .ButtonHeight = sglH
        ' Add on some space for the Notes/Comments icon
        .ButtonWidth = .RadioWidth + sglW + m_NOTESTATUS_IMAGE_WIDTH
        
        .Visible = True
        .Enabled = False
        
        ' No. of buttons = no. of category values
        ' DPH 24/10/2001 - Count only active categories
        '.Buttons = oElement.Categories.Count
        nActiveCat = 0
        For Each oCategory In oElement.Categories
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
        For Each oCategory In oElement.Categories
            
            ' DPH 24/10/2001 - Use only active categories
            If oCategory.Active Then
                oOptGroup.Caption(nButtonIndex) = oCategory.Value
                oOptGroup.TagValue(nButtonIndex) = oCategory.Code
                 
                'increment the index for button each time
                nButtonIndex = nButtonIndex + 1
            End If
        Next    ' oCategory
        
        ' Finally adjust the Left
        oOptGroup.Left = AdjustHorizontalPosition(oOptGroup.Left, oOptGroup.Width, oContainer)
        
        ' Keep track of the max Y coordinate for non-group items
        If oElement.OwnerQGroup Is Nothing Then
            If oOptGroup.Top + oOptGroup.Height > msglMaxY Then
                msglMaxY = oOptGroup.Top + oOptGroup.Height
            End If
            ' NCJ 5 Nov 02 - Also remember the Max X coord
            If oOptGroup.Left + oOptGroup.Width > msglMaxX Then
                msglMaxX = oOptGroup.Left + oOptGroup.Width
            End If
        End If

        'set the values of each option button inside the control to false
        oOptGroup.UnselectAll
    
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub CalculateOptionButtonSize(oElement As eFormElementRO, _
                    ByRef sglWidth As Single, sglHeight As Single)
'---------------------------------------------------------------------
' NCJ 27 Feb 03
' Calculate the max text width & height of a category description
'---------------------------------------------------------------------
Dim oCat As CategoryItem

    With moVBForm
        .FontName = oElement.FontName
        .FontSize = oElement.FontSize
        .FontBold = oElement.FontBold
        .FontItalic = oElement.FontItalic
        sglWidth = 0
        sglHeight = 0
        For Each oCat In oElement.Categories
            If .TextWidth(oCat.Value) > sglWidth Then
                sglWidth = .TextWidth(oCat.Value)
            End If
            If .TextHeight(oCat.Value) > sglHeight Then
                sglHeight = .TextHeight(oCat.Value)
            End If
        Next
    End With
    
    ' Add on a bit for good measure
    ' NCJ 21 May 03 - Removed the bit for good measure (BUG 837)
'    sglWidth = sglWidth + 60
    sglHeight = sglHeight + 60
    
End Sub

'---------------------------------------------------------------------
Private Sub BuildAttachmentButton(oElement As eFormElementRO, _
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

        
        '  Set the display attributes
        .TabStop = True
        .TabIndex = mnTabOrder
        mnTabOrder = mnTabOrder + 1
        .Visible = True
        .Enabled = False
    
    End With

End Sub
'-------------------------------------------------------------------
Private Sub BuildExpandButton(oTextBox As TextBox)
'-------------------------------------------------------------------
' Build an expand button to the right of the given text box
' and with the same index
' RS 19/02/2003: Created
'-------------------------------------------------------------------

    Load moVBForm.cmdExpand(oTextBox.Index)
    With moVBForm.cmdExpand(oTextBox.Index)
        Set .Container = oTextBox.Container
        .Top = oTextBox.Top
        .Left = oTextBox.Left + oTextBox.Width
        .Height = oTextBox.Height
        .Width = .Height
        .Visible = True
        .Enabled = True
    End With

End Sub

'-------------------------------------------------------------------
Private Sub BuildDictionaryButton(oTextBox As TextBox)
'-------------------------------------------------------------------
' Build a dictionary button to the right of the given label
' and with the same index
'-------------------------------------------------------------------

    Load moVBForm.imgDictionary(oTextBox.Index)
    With moVBForm.imgDictionary(oTextBox.Index)
        Set .Container = oTextBox.Container
        .Top = oTextBox.Top
        .Left = oTextBox.Left + oTextBox.Width
        .Height = oTextBox.Height
        .Width = .Height
        .Visible = True
        .Enabled = True
    End With

End Sub

'-------------------------------------------------------------------
Private Sub BuildCalendarButton(oTextBox As TextBox)
'-------------------------------------------------------------------
' Build a calendar button to the right of the given text box
' and with the same index
'-------------------------------------------------------------------

    Load moVBForm.cmdCalendar(oTextBox.Index)
    With moVBForm.cmdCalendar(oTextBox.Index)
        Set .Container = oTextBox.Container
        .Top = oTextBox.Top
        .Left = oTextBox.Left + oTextBox.Width
        .Height = oTextBox.Height
        .Width = .Height
        .Visible = True
        .Enabled = False
        ' NCJ 22 May 03 - Also remember the Max X coord
        If .Left + .Width > msglMaxX Then
            msglMaxX = .Left + .Width
        End If
    End With

End Sub

'-------------------------------------------------------------------
Private Sub BuildPopupButton(oTextBox As TextBox)
'-------------------------------------------------------------------
' Build a popuplist button to the right of the given text box
' and with the same index
'-------------------------------------------------------------------

    Load moVBForm.cmdPopupList(oTextBox.Index)
    With moVBForm.cmdPopupList(oTextBox.Index)
        Set .Container = oTextBox.Container
        .Top = oTextBox.Top
        .Left = oTextBox.Left + oTextBox.Width
        .Height = oTextBox.Height
        .Width = .Height
        .Visible = True
        .Enabled = False
        ' NCJ 22 May 03 - Also remember the Max X coord
        If .Left + .Width > msglMaxX Then
            msglMaxX = .Left + .Width
        End If
    End With

End Sub


'---------------------------------------------------------------------
Private Sub BuildNRCTCLabel(oElement As eFormElementRO, _
                        nIndex As Integer, oContainer As PictureBox)
'---------------------------------------------------------------------
' Build NRStatus/CTCGrade label for given eForm element
' This sits to the right of the element's status icon
'---------------------------------------------------------------------

    Load moVBForm.lblNRCTC(nIndex)
    Set moVBForm.lblNRCTC(nIndex).Container = oContainer
    
    With moVBForm.lblNRCTC(nIndex)
        .Tag = oElement.Name
        .Top = moVBForm.imgStatus(nIndex).Top
        ' Put it to the right of status icon (and ChageCont icon
        .Left = StatusImageLeft(oElement, nIndex, True) + moVBForm.imgStatus(nIndex).Width + msglGAP
        .Left = AdjustHorizontalPosition(.Left, .Width, oContainer)
        'TA 12/09/2002: keep in line with status image size
        .Height = moVBForm.imgStatus(nIndex).Height
        .FontSize = 9
        .FontBold = True
        .Alignment = vbCenter
    End With

End Sub

'---------------------------------------------------------------------
Private Sub BuildImage(oElement As eFormElementRO, _
                        nIndex As Integer, oContainer As PictureBox)
'---------------------------------------------------------------------
' Build image for given eForm element's status
' This sits to the right of the element itself
'---------------------------------------------------------------------

On Error GoTo ErrLabel

    If oElement.ShowStatusFlag Then
        'only load status and change count images if we are going to show them
    
        Load moVBForm.imgStatus(nIndex)
        Set moVBForm.imgStatus(nIndex).Container = oContainer
        
        With moVBForm.imgStatus(nIndex)
            .Width = m_STATUS_IMAGE_WIDTH
            .Height = m_CHANGECOUNT_STATUS_IMAGE_HEIGHT
            .Tag = oElement.Name
            .Top = ChangeCountStatusImageTop(oElement, nIndex)
            '(oElement.QGroup Is Nothing) parameter because we don't have changecount on RQGs
            .Left = StatusImageLeft(oElement, nIndex, (oElement.QGroup Is Nothing))
            .Left = AdjustHorizontalPosition(.Left, .Width, oContainer)
        End With
        
        
        Load moVBForm.imgSDVStatus(nIndex)
        Set moVBForm.imgSDVStatus(nIndex).Container = oContainer
        
        With moVBForm.imgSDVStatus(nIndex)
            .Width = m_STATUS_IMAGE_WIDTH
            .Height = 40
            .Top = moVBForm.imgStatus(nIndex).Top + moVBForm.imgStatus(nIndex).Height
            '(oElement.QGroup Is Nothing) parameter because we don't have changecount on RQGs
            .Left = moVBForm.imgStatus(nIndex).Left
            '.BackColor = .Container.BackColor '.vbGreen
            .Visible = False
        End With

    
        If oElement.QGroup Is Nothing Then
            'only do this for non RQGs
            'TA 29/08/2002: load controls needed to display change count
            
            Load moVBForm.imgChangeCount(nIndex)
            Set moVBForm.imgChangeCount(nIndex).Container = oContainer
        
            With moVBForm.imgChangeCount(nIndex)
                .Tag = oElement.Name
                .Top = moVBForm.imgStatus(nIndex).Top
                .Width = m_CHANGECOUNT_IMAGE_WIDTH
                .Height = m_CHANGECOUNT_STATUS_IMAGE_HEIGHT
                .Left = moVBForm.imgStatus(nIndex).Left - .Width + msglGAP \ 2
                .Left = AdjustHorizontalPosition(.Left, .Width, oContainer)
            End With
        End If
                
        
        If oElement.DataType = eDataType.LabTest Then
            'set up the Lab Result status label
            BuildNRCTCLabel oElement, nIndex, oContainer
        End If
        
    End If

    'we always show note and sdv status
    If oElement.QGroup Is Nothing Then
        'only do this for non RQGs
        'comment top right
        Load moVBForm.picCommentStatus(nIndex)
        Set moVBForm.picCommentStatus(nIndex).Container = oContainer
        
        With moVBForm.picCommentStatus(nIndex)
            .Width = m_NOTESTATUS_IMAGE_WIDTH
            .Height = m_NOTESTATUS_IMAGE_HEIGHT
            .Top = ChangeCountStatusImageTop(oElement, nIndex)
            .Left = GetControlRight(oElement, nIndex) - .Width - 40
            ' NCJ 21 May 03 - Fudge factor for BUG 837
            If oElement.ControlType = gn_OPTION_BUTTONS Then
                .Left = .Left + 40
            End If
            .BackColor = GetControlFromElement(oElement, nIndex).BackColor
        End With
    
                   
                   
                   
        Load moVBForm.picNoteStatus(nIndex)
        Set moVBForm.picNoteStatus(nIndex).Container = oContainer
        
        With moVBForm.picNoteStatus(nIndex)
            .Width = m_NOTESTATUS_IMAGE_WIDTH
            .Height = m_NOTESTATUS_IMAGE_HEIGHT
            .Top = GetControlBottom(oElement, nIndex) - .Height - 40
            .Left = GetControlRight(oElement, nIndex) - .Width - 40
            ' NCJ 21 May 03 - Fudge factor for BUG 837
            If oElement.ControlType = gn_OPTION_BUTTONS Then
                .Left = .Left + 40
            End If
            .BackColor = GetControlFromElement(oElement, nIndex).BackColor
        End With
    End If
    
    'TA 31/10/2002: Set help text : comment to be done separately
    With moVBForm
        If oElement.ShowStatusFlag Then
            .imgStatus(nIndex).TooltipText = oElement.HelpText
        End If
    End With

Exit Sub

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "ElementBuilder.BuildImage"

End Sub

'---------------------------------------------------------------------
Private Sub BuildComment(oElement As eFormElementRO, _
                        nIndex As Integer, oContainer As PictureBox, _
                        lTop As Long, lLeft As Long, _
                        bHotlink As Boolean)
'---------------------------------------------------------------------
' Build comment for given eForm element
'---------------------------------------------------------------------
Dim oLabel As Label

    '  Load a new label control
    Load moVBForm.lblCaption(nIndex)
    Set oLabel = moVBForm.lblCaption(nIndex)
    Set oLabel.Container = oContainer
    
    '  Set font attributes
    ' ZA 19/07/2002 - replace SetElementFont with SetCaptionFont
    Call SetCaptionFont(oLabel, oElement)
        
    With oLabel
        .Tag = oElement.Caption
        .Caption = oElement.Caption
        
        If bHotlink Then
            .Font.Underline = True
            .TooltipText = "Hotlink"
            .MousePointer = vbCustom
        Else
            .TooltipText = "Comment"
        End If
        
        '  Set position attributes
        .AutoSize = True
        .Alignment = vbLeftJustify
        
        .Top = lTop
        .Left = lLeft
        
        .Left = AdjustHorizontalPosition(.Left, .Width, oContainer)
        
        'Changed by Mo Morris 4/1/00
        'extra bit added to the AutoSize created width so that there is enough space to
        'display (and move) Italic fonts with out little bits being left on the screen
        .Width = .Width + (.FontSize * 3)
        
        '  Set display attributes
        .BackColor = oContainer.BackColor
        'ZA 25/07/2002 - use CaptionFontCOlour instead of FontColour
        .ForeColor = oElement.CaptionFontColour
        .Visible = True
        .Enabled = True
    
        ' Keep track of the max Y coordinate for non-group items
        If oElement.OwnerQGroup Is Nothing Then
            If .Top + .Height > msglMaxY Then
                msglMaxY = .Top + .Height
            End If
            ' NCJ 5 Nov 02 - Also remember the Max X coord
            If .Left + .Width > msglMaxX Then
                msglMaxX = .Left + .Width
            End If
        End If
    
    End With

End Sub

'---------------------------------------------------------------------
Private Sub BuildLine(oElement As eFormElementRO, _
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
        .TooltipText = "Line"
        .Visible = True
        .Enabled = True
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
    
    With oPic
        ' Keep track of the max Y coordinate for non-group items
        ' (But do not update the last X coord for a line)
        If .Top + .Height > msglMaxY Then
            msglMaxY = .Top + .Height
        End If
    End With
    
    Set oPic = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub BuildTextBox(oElement As eFormElementRO, _
                        nIndex As Integer, oContainer As PictureBox, _
                        lTop As Long, lLeft As Long)
'---------------------------------------------------------------------
' Build text box for given eForm element
'
' MLM 30/08/02: Build the text box differently if it's for an eForm date.
' ic 14/07/2005 added clinical coding
'---------------------------------------------------------------------
Dim oTextBox As TextBox
Dim sglWidth As Single
Dim sglHeight As Single

    '  Load a new text box
    Load moVBForm.txtCRFElement(nIndex)
    Set oTextBox = moVBForm.txtCRFElement(nIndex)
    Set oTextBox.Container = oContainer
        
    With oTextBox
        'MLM 30/08/02: Make appearance of form date different
        Select Case oElement.ElementUse
        Case eElementUse.EFormVisitDate
            .Font.Name = gsDATE_FONT
            .Font.Charset = 1
            .Font.SIZE = gnDATE_FONT_SIZE
            .ForeColor = glDATE_FONT_COLOUR
            .BorderStyle = 0 'None
            
            'Set data entry attributes
            .TabIndex = 0

        Case Else
            '  Set font attributes
            Call SetElementFont(oTextBox, oElement)
            .ForeColor = oElement.FontColour

            '  Set data entry attributes
            .TabIndex = mnTabOrder

        End Select

        '  Set position attributes
        .Top = lTop
        .Left = lLeft
        
        Call CalculateBoxSize(oElement, sglWidth, sglHeight)
        .Width = sglWidth
        .Height = sglHeight
        
        .Left = AdjustHorizontalPosition(.Left, .Width, oContainer)
        
    '   ATN 18/3/99 SR732
    '   Truncate the width if it's too wide for the form
        'Mo Morris 5/6/00 SR 3460, mn_SPACE_FOR_STATUS_ICON replaces 495
        'Test for control type POPUP_List or Calendar because they require a command button as well as
        'space for an icon to the right of the text box.
        'note that the width of the command button is equal to the height of the textbox it belongs to
        'to the right of the text box
        If oElement.OwnerQGroup Is Nothing Then
            ' Only adjust widths for non-group items
            sglWidth = .Left
            If oElement.ShowStatusFlag Then
                sglWidth = sglWidth + mn_SPACE_FOR_STATUS_ICON
            End If
            If oElement.ControlType = gn_POPUP_LIST Or oElement.ControlType = gn_CALENDAR Then
                ' Add in width of popup/calendar button (same as height of textbox)
                sglWidth = sglWidth + .Height
            End If
            
            'RS 19/02/2003: Add in width of expand button (same as height of textbox)
            ' If oElement.ControlType = gn_TEXT_BOX And oElement.DisplayLength > 0 Then
            If oElement.Expandable Then       ' NCJ 1 Jul 03
                sglWidth = sglWidth + .Height
            End If
            
            'ic 23/02/2007 bug 2855 for DE dont display clinical coding button if clinical coding
            'is not installed, just display as a textbox. do display the button for SD
            If Not ((App.Title = "MACRO_DM") And (Not gbClinicalCoding)) Then
                'ic 15/07/2005 added clinical coding
                If oElement.DataType = eDataType.Thesaurus Then
                    sglWidth = sglWidth + .Height
                End If
            End If
            
            If sglWidth + .Width > oContainer.Width Then
                ' NCJ 8 Jul 02 - Refix of REM CBB 2.2.16/9 from 2.2
                ' Don't set width to less than mn_SPACE_FOR_STATUS_ICON
                If (oContainer.Width - sglWidth) > mn_SPACE_FOR_STATUS_ICON Then
                    .Width = oContainer.Width - sglWidth
                Else
                    .Width = mn_SPACE_FOR_STATUS_ICON
                End If
            End If
        End If
        
        '  Set data entry attributes
        '.TabIndex = mnTabOrder
        'ic 14/07/2005 added clinical coding
        mnTabOrder = mnTabOrder + 1
        Select Case oElement.DataType
        Case eDataType.Text, eDataType.Category, eDataType.DateTime, eDataType.Thesaurus
            .MaxLength = oElement.QuestionLength
        Case Else   ' Numbers
            .MaxLength = Len(oElement.Format) + 1
        End Select

        
        '  Set the display attributes

'03/12/2001: VTRACK Changes  - Text boxes have windows background
#If VTRACK <> 1 Then
       .BackColor = oContainer.BackColor
#Else
       .BackColor = vbWindowBackground
#End If

        .Visible = True
        .Enabled = False
        
        ' Keep track of the max Y coordinate for non-group items
        If oElement.OwnerQGroup Is Nothing Then
            If .Top + .Height > msglMaxY Then
                msglMaxY = .Top + .Height
            End If
            ' NCJ 5 Nov 02 - Also remember the Max X coord
            If .Left + .Width > msglMaxX Then
                msglMaxX = .Left + .Width
            End If
        End If
        
    End With

End Sub

'---------------------------------------------------------------------
Private Sub BuildPicture(oElement As eFormElementRO, _
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
        .TooltipText = "Picture: " & oElement.Caption
        
        On Error Resume Next
        .Picture = LoadPicture(gsDOCUMENTS_PATH & oElement.Caption)
        On Error GoTo ErrLabel
        
        If Err.Number = 0 Then
            .AutoSize = True
            .Visible = True
            .Enabled = True
            'send to back
            .ZOrder 1
            ' Keep track of the max Y coordinate for non-group items
            If oElement.OwnerQGroup Is Nothing Then
                If .Top + .Height > msglMaxY Then
                    msglMaxY = .Top + .Height
                End If
                ' NCJ 5 Nov 02 - Also remember the Max X coord
                If .Left + .Width > msglMaxX Then
                    msglMaxX = .Left + .Width
                End If
            End If
        End If
    End With
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilder.BuildPicture"
    
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
'    Set moEFI = Nothing
    Set mcolFormItems = Nothing
    Set mcolGroupsByID = Nothing
    Set mcolGroupsByIndex = Nothing
    Set mcolControlIndexes = Nothing
    Set moSubject = Nothing

End Sub

'---------------------------------------------------------------------
Private Sub BuildCaption(oElement As eFormElementRO, nIndex As Integer, _
                            oContainer As PictureBox)
'---------------------------------------------------------------------
' Build the caption of an eForm Element
' Assume it's a non-group item
' and it may be auto-positioned or be specifically positioned
' NCJ 3 Sept 02 - Offset Y coordinate for new eForm/Visit date bar
' MLM 24/06/03: Special positioning and style for form and visit date captions.
'---------------------------------------------------------------------
Dim sUnitOfMeasurement As String
Dim sCaption As String
Dim oLabel As Label
    
    
    '  Load a new label control
    Load moVBForm.lblCaption(nIndex)
    Set oLabel = moVBForm.lblCaption(nIndex)
    
    
    Set oLabel.Container = oContainer
    
    '  Set font attributes
    ' ZA 19/07/2002 - replace SetElementFont with SetCaptionFOnt
    Call SetCaptionFont(oLabel, oElement)
            
    '  Prepare unit of measurement part of label
    sUnitOfMeasurement = ""
    If oElement.Unit > "" Then
        sUnitOfMeasurement = " (" & oElement.Unit & ")"
    End If
    
    With oLabel
        .Tag = oElement.Name
        If moEForm.DisplayNumbers And oElement.ElementUse = eElementUse.User Then
            .Caption = oElement.ElementOrder & ". " & oElement.Caption & sUnitOfMeasurement
        Else
            .Caption = oElement.Caption & sUnitOfMeasurement
        End If
        
        '  Set position attributes
        .AutoSize = True
        
        'Added by ASH 31/07/2001
        'Changed by Mo Morris 4/1/00
        'extra bit added to the AutoSize created width so that there is enough space to
        'display (and move) Italic fonts without little bits being left on the screen
        .Width = .Width + (.FontSize * 3)
         
        .Alignment = vbLeftJustify
        
        ' If no caption then make label invisible
        If oElement.Caption > "" Then
            .Visible = True
        Else
            .Visible = False
        End If

        
        If oElement.ElementUse = eElementUse.User Then
            If oElement.CaptionY > 0 Then
                .Top = oElement.CaptionY
            Else
                .Top = oElement.ElementY
            End If
            ' NCJ 3 Sept 02 - Offset by required amount
            .Top = .Top + msglYOffset
            
            If oElement.CaptionX > 0 Then
                .Left = oElement.CaptionX
            Else
                .Left = oElement.ElementX - .Width - 50
            End If
        
            .Left = AdjustHorizontalPosition(.Left, .Width, oContainer)
            .ForeColor = oElement.CaptionFontColour
        Else
            If oElement.EFormId = moEForm.EFormId Then
                'form date
                .Left = 4320
                moVBForm.lblEFormDate.Visible = Not .Visible
            Else
                'visit date
                .Left = 60
                moVBForm.lblVisitDate.Visible = Not .Visible
            End If
            .Top = 450
            .ForeColor = glDATE_FONT_COLOUR
        End If
        
        '  Set display attributes
        .BackColor = oContainer.BackColor
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
Private Sub CalculateBoxSize(oElement As eFormElementRO, _
                    ByRef sglWidth As Single, ByRef sglHeight As Single)
'---------------------------------------------------------------------
' Calculate box size based on element properties
'
' MLM 10/09/02: Make form and visit dates shorter and potentially less wide
' MLM 23/01/03: Form's header bar now made larger; form and visit dates should be normal width
' NCJ 1 Jul 03 - Use Element's Expandable property
'---------------------------------------------------------------------
Dim nQLength As Integer

    On Error Resume Next
    
    With moVBForm
        .FontName = oElement.FontName
        .FontSize = oElement.FontSize
        .FontBold = oElement.FontBold
        .FontItalic = oElement.FontItalic

        If oElement.Expandable Then
            ' Use Display Length for expandable text fields
            nQLength = oElement.DisplayLength
        Else
            nQLength = oElement.QuestionLength
        End If
        
        ' Add 4 extra characters for good measure
        sglWidth = .TextWidth(String(nQLength + 4, "_"))
        
'        'ZA 20/09/2002
'        'use display length only for Expandable text question
'        If (oElement.ControlType = gn_TEXT_BOX) Or (oElement.ControlType = gn_POPUP_LIST) Then
'            'use displaylength only if it is not NULL
'            If oElement.DisplayLength > 0 Then
'                sglWidth = .TextWidth(String(oElement.DisplayLength + 4, "_"))
'            End If
'        End If
        
        ' Add 100 for good measure
        sglHeight = .TextHeight("_") + 100
            
    End With
    
    If oElement.ElementUse = eElementUse.EFormVisitDate Then
        'form and visit dates are less high
        sglHeight = 250
'        If sglWidth > 1800 Then
'            sglWidth = 1800
'        End If
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub SetOptionGroupFont(oOptGroup As OptionGroup, oElement As eFormElementRO)
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
Private Sub SetElementFont(oControl As Control, oElement As eFormElementRO)
'---------------------------------------------------------------------
' Set the font for the control for the given eForm Element
' The element's own properties are correctly set up with defaults
' as appropriate in MACRO 3.0
'---------------------------------------------------------------------

    oControl.Font.Name = oElement.FontName
    oControl.Font.Charset = 1
    oControl.Font.SIZE = oElement.FontSize
    oControl.Font.Bold = oElement.FontBold
    oControl.Font.Italic = oElement.FontItalic
    
End Sub

'---------------------------------------------------------------------
Private Sub SetCaptionFont(oControl As Control, oElement As eFormElementRO)
'---------------------------------------------------------------------
' Set the font for the caption for the given eForm Element
' The element's own properties are correctly set up with defaults
' as appropriate in MACRO 3.0
' MLM 24/06/03: Fixed attributes for form and visit dates
'---------------------------------------------------------------------
    
    If oElement.ElementUse = eElementUse.EFormVisitDate Then
        oControl.Font.Name = gsDATE_FONT
        oControl.Font.Charset = 1
        oControl.Font.SIZE = gnDATE_FONT_SIZE
'        oControl.Font.Bold = oElement.CaptionFontBold
'        oControl.Font.Italic = oElement.CaptionFontItalic
    Else
        oControl.Font.Name = oElement.CaptionFontName
        oControl.Font.Charset = 1
        oControl.Font.SIZE = oElement.CaptionFontSize
        oControl.Font.Bold = oElement.CaptionFontBold
        oControl.Font.Italic = oElement.CaptionFontItalic
    End If
    
End Sub

'---------------------------------------------------------------------
Private Function GetControlRight(oElement As eFormElementRO, nIndex As Integer) As Long
'---------------------------------------------------------------------
' Calculate the right coordinate for a question element's Note Image
' nIndex is the index of its controls
'---------------------------------------------------------------------
Dim oControl As Control

On Error GoTo ErrLabel:

    Set oControl = GetControlFromElement(oElement, nIndex)
    GetControlRight = oControl.Left + oControl.Width
    
Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|ElementBuilder.GetControlRight"
    
End Function

'---------------------------------------------------------------------
Private Function GetControlBottom(oElement As eFormElementRO, nIndex As Integer) As Long
'---------------------------------------------------------------------
' Calculate the coordinate for the bottom of control
' need to know height of control
' nIndex is the index of its controls
'---------------------------------------------------------------------
Dim oControl As Control
Dim sglWidth As Single
Dim sglHeight As Single

On Error GoTo ErrLabel:


    GetElementSize oElement, nIndex, sglWidth, sglHeight
    Set oControl = GetControlFromElement(oElement, nIndex)
    GetControlBottom = oControl.Top + sglHeight
    
Exit Function

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|ElementBuilder.NoteImageTop"
    
End Function
'---------------------------------------------------------------------
Private Function ChangeCountImageLeft(oElement As eFormElementRO, nIndex As Integer) As Long
'---------------------------------------------------------------------
' Calculate the left coordinate for a question element's Status Image
' nIndex is the index of its controls
' RS 19/02/2003: Added expand button
' ic 15/07/2005 added clinical coding button
'---------------------------------------------------------------------
Dim nControlType As Integer
Dim lWidth As Long
Dim oControl As Control
    
    nControlType = oElement.ControlType
    Set oControl = GetControlFromElement(oElement, nIndex)
    lWidth = oControl.Left + oControl.Width
    Set oControl = Nothing
    
    'If nControlType = gn_TEXT_BOX And (oElement.DisplayLength > 0) Then
    If oElement.Expandable Then
        'add on width of expand button
        lWidth = lWidth + moVBForm.cmdExpand(nIndex).Width
    End If
    
    If nControlType = gn_POPUP_LIST Then
        'add on width of command button for popuplist questions
        lWidth = lWidth + moVBForm.cmdPopupList(nIndex).Width
    End If
    
    If nControlType = gn_CALENDAR Then
        'add on width of calendar button for calendars
        lWidth = lWidth + moVBForm.cmdCalendar(nIndex).Width
    End If
    
    'ic 23/02/2007 bug 2855 for DE dont display clinical coding button if clinical coding
    'is not installed, just display as a textbox. do display the button for SD
    If Not ((App.Title = "MACRO_DM") And (Not gbClinicalCoding)) Then
        'ic 15/07/2005 added clinical coding
        If oElement.DataType = eDataType.Thesaurus Then
            lWidth = lWidth + moVBForm.imgDictionary(nIndex).Width
        End If
    End If
        
    ChangeCountImageLeft = lWidth
    
End Function

'---------------------------------------------------------------------
Private Function StatusImageLeft(oElement As eFormElementRO, nIndex As Integer, bHasChangeCountImage As Boolean) As Long
'---------------------------------------------------------------------
' Calculate the left coordinate for a question element's Status Image
' nIndex is the index of its controls
'---------------------------------------------------------------------

    If bHasChangeCountImage Then
        StatusImageLeft = ChangeCountImageLeft(oElement, nIndex) + m_CHANGECOUNT_IMAGE_WIDTH
    Else
        StatusImageLeft = ChangeCountImageLeft(oElement, nIndex)
    End If
    
End Function

'-------------------------------------------------------------------
Private Function ChangeCountStatusImageTop(oElement As eFormElementRO, nIndex As Integer) As Long
'-------------------------------------------------------------------
' Get the Top coordinate for the Status Image of the given control
'-------------------------------------------------------------------
Dim oControl As Control
Dim lControlHeight As Long

    Set oControl = GetControlFromElement(oElement, nIndex)
    
    Select Case oElement.ControlType
    Case gn_OPTION_BUTTONS
        'for option buttons we only want the height of one option button within the control
        lControlHeight = oControl.ButtonHeight
    Case gn_PUSH_BUTTONS
        lControlHeight = oControl.Height \ oElement.Categories.Count
    Case Else
        lControlHeight = oControl.Height
    End Select
'calculate half the difference between the height of the control and height of image
'   status so we can vertically centre the image
    
    ChangeCountStatusImageTop = oControl.Top + (lControlHeight - m_CHANGECOUNT_STATUS_IMAGE_HEIGHT) \ 2
    Set oControl = Nothing
    
End Function

'-------------------------------------------------------------------
Private Function GetItemKey(oElement As eFormElementRO, nRow As Integer) As String
'-------------------------------------------------------------------
' Get the key for a given element and row
'
' MLM 04/09/02: Include the element's EFormId in the key, so the the collection can
'   accommodate elements from different EForms.
'-------------------------------------------------------------------

    GetItemKey = "K" & oElement.EFormId & "|" & oElement.ElementID & "|" & nRow

End Function

'-------------------------------------------------------------------
Private Function GetElementSize(oElement As eFormElementRO, nIndex As Integer, _
                ByRef sglWidth As Single, ByRef sglHeight As Single)
'-------------------------------------------------------------------
' Get the width and height of a built eForm element (questions only)
' including its status image/NTCTC label if appropriate
' but NOT including a question's caption
'-------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    With moVBForm
                       
        ' Differs depending on what the element is
        Select Case oElement.ControlType
        Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX, _
                    gn_POPUP_LIST, gn_CALENDAR
            ' These all have text boxes
            sglWidth = .txtCRFElement(nIndex).Width
            sglHeight = .txtCRFElement(nIndex).Height
            
            If oElement.ControlType = gn_POPUP_LIST Then
                sglWidth = sglWidth + .cmdPopupList(nIndex).Width
            End If
            
            If oElement.ControlType = gn_CALENDAR Then
                sglWidth = sglWidth + .cmdCalendar(nIndex).Width
            End If
            
            ' RS 19/02/2003: Added Expand button
            'If oElement.ControlType = gn_TEXT_BOX And (oElement.DisplayLength > 0) Then
            If oElement.Expandable Then
                sglWidth = sglWidth + .cmdExpand(nIndex).Width
            End If
            
            'ic 23/02/2007 bug 2855 for DE dont display clinical coding button if clinical coding
            'is not installed, just display as a textbox. do display the button for SD
            If Not ((App.Title = "MACRO_DM") And (Not gbClinicalCoding)) Then
                'ic 15/07/2005 added clinical coding
                If oElement.DataType = eDataType.Thesaurus Then
                    sglWidth = sglWidth + .imgDictionary(nIndex).Width
                End If
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
        
        ' Add in width of status flag & lab test result label
        sglWidth = sglWidth + SpaceForStatusIcon(oElement, nIndex)
        
    End With
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilder.GetElementSize"
    
End Function

'--------------------------------------------------------------------------------------------------
Private Function SpaceForStatusIcon(oElement As eFormElementRO, nIndex As Integer) As Single
'--------------------------------------------------------------------------------------------------
' NCJ 9 Oct 02
' Space for status icon for an element
'--------------------------------------------------------------------------------------------------
Dim sglWidth As Single

    If oElement.ShowStatusFlag Then
        With moVBForm
            sglWidth = m_STATUS_IMAGE_WIDTH + m_CHANGECOUNT_IMAGE_WIDTH
            If oElement.DataType = eDataType.LabTest Then
                sglWidth = sglWidth + msglGAP + .lblNRCTC(nIndex).Width
            End If
            SpaceForStatusIcon = sglWidth
        End With
    Else
        ' We're not showing status flags
        SpaceForStatusIcon = 0
    End If
    
End Function

'-------------------------------------------------------------------
Private Sub ResizeGroup(oEFQGroup As EFormQGroupDE, ByVal nIndex As Integer)
'-------------------------------------------------------------------
' Align everything in a question group when initialising or after a resize
'-------------------------------------------------------------------
Dim sglRoomForVScrollBar As Single
Dim sglMaxWidth As Single
Dim sglPicFixedWidth As Single
Dim nDisplayCols As Integer
Dim nCol As Integer
Dim oElement As eFormElementRO
Dim lAmountToScroll As Long

    On Error GoTo ErrLabel
    
    With moVBForm
        ' Anchor picFixed inside picQuGroup
        ' Leave a tiny gap (the same as in SD)
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
                                    + oEFQGroup.RowHeight * oEFQGroup.QGroup.DisplayRows - msglGAP
        
        ' Scroller pic sized to contain all Rows
        If oEFQGroup.CurrentRows >= oEFQGroup.QGroup.DisplayRows Then
'            .picScroller(nIndex).Height = msglGAP + oEFQGroup.RowHeight * oEFQGroup.CurrentRows
            .picScroller(nIndex).Height = oEFQGroup.RowHeight * oEFQGroup.CurrentRows
        Else
            ' No. of rows is less than DisplayRows
            .picScroller(nIndex).Height = oEFQGroup.RowHeight * oEFQGroup.QGroup.DisplayRows
        End If
        
        ' ***** Now adjust the width so it fits on the screen ****
        
        ' We always leave room for the vertical scroll bar
        ' Fudge factor - have to add on 3 gaps instead of 2!
        sglRoomForVScrollBar = 3 * msglGAP + .vsbQuGroup(nIndex).Width
        
        Set oElement = oEFQGroup.QGroup.Element
        ' The max width we can have to fit on the screen
        sglMaxWidth = .picQuGroup(nIndex).Container.Width - .picQuGroup(nIndex).Left - SpaceForStatusIcon(oElement, nIndex)
        
        ' The default width of picFixed
        sglPicFixedWidth = oEFQGroup.RowWidth
        
        ' Start with rightmost column and peel off columns until we get a fit
        nDisplayCols = oEFQGroup.EFormItems.Count

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
        
        ' Adjust height of outer picture box (we'll add in space for HSB later if needed)
        .picQuGroup(nIndex).Height = 2 * msglGAP + .picFixed(nIndex).Top + .picFixed(nIndex).Height
        
        ' Initialise the scroll bars
        ' (vertical Value is how many rows are hiding at the top)
        ' Vert. scroll bar aligns with top of picScroller
        .vsbQuGroup(nIndex).Top = .picScroller(nIndex).Top
        .vsbQuGroup(nIndex).Height = .picFixed(nIndex).Height - .vsbQuGroup(nIndex).Top
        .vsbQuGroup(nIndex).Left = .picFixed(nIndex).Left + .picFixed(nIndex).Width + msglGAP

        If nDisplayCols < oEFQGroup.EFormItems.Count Then
            ' NCJ 4 Nov 02 - Calculate increment and store in scroll bar tag
            ' Need to set up some scrolling - divide scroll amount by an increment, to avoid overflow in scroll bar Max
            lAmountToScroll = (.picScroller(nIndex).Width - sglPicFixedWidth)
            ' Store the increment in the scroll bar Tag
            .hsbQuGroup(nIndex).Tag = HScrollBarIncrement(lAmountToScroll)
            .hsbQuGroup(nIndex).Max = lAmountToScroll / .hsbQuGroup(nIndex).Tag
            
            ' Let them click 8 times on arrows and 4 times for big jumps
            .hsbQuGroup(nIndex).SmallChange = (.hsbQuGroup(nIndex).Max) / 8
            .hsbQuGroup(nIndex).LargeChange = (.hsbQuGroup(nIndex).Max) / 4
            
            .hsbQuGroup(nIndex).Height = .vsbQuGroup(nIndex).Width
            .hsbQuGroup(nIndex).Width = .picFixed(nIndex).Width - msglGAP
            .hsbQuGroup(nIndex).Top = .picFixed(nIndex).Top + .picFixed(nIndex).Height + msglGAP
            .hsbQuGroup(nIndex).Left = .picFixed(nIndex).Left + .picScroller(nIndex).Left
            .hsbQuGroup(nIndex).Value = 0
            
            ' Expand outer pic to include horiz. scroll bar
            .picQuGroup(nIndex).Height = .picQuGroup(nIndex).Height + msglGAP + .hsbQuGroup(nIndex).Height
            
            .hsbQuGroup(nIndex).Visible = True
            .hsbQuGroup(nIndex).Enabled = True
        Else
            ' No horizontal scrolling
            .hsbQuGroup(nIndex).Visible = False
        End If
        
    End With
    
    Set oElement = Nothing
            
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilder.ResizeGroup"

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
Public Sub BuildNewGroupRow(oEFQGroup As EFormQGroupDE, nRow As Integer)
'-------------------------------------------------------------------
' Add a new row of fields for a group
' Assume first row aleady built and all relevant properties of oEFQGroup set up
'-------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    ' NCJ 3 Dec 02 - If the group is already built, need to get the next tab order
    If oEFQGroup.GroupIsBuilt Then
        If oEFQGroup.LastItem.Element.ControlType = gn_PUSH_BUTTONS Then
            mnTabOrder = moVBForm.optGraphical(oEFQGroup.LastPushButtonIndex).TabIndex + 1
        Else
            ' Get the tabindex of the last group element
            mnTabOrder = GetTabIndex(oEFQGroup.LastItem) + 1
        End If
    End If
    
    Call BuildGroupRow(oEFQGroup, nRow)
    ' Increment the height of the scrolling box by 1 row
    moVBForm.picScroller(oEFQGroup.ControlIndex).Height = moVBForm.picScroller(oEFQGroup.ControlIndex).Height + oEFQGroup.RowHeight
    
    If nRow > oEFQGroup.QGroup.DisplayRows Then
        ' Bring on the scroll bar
        With moVBForm.vsbQuGroup(oEFQGroup.ControlIndex)
            .Visible = True
            .Enabled = True
            .Max = nRow - oEFQGroup.QGroup.DisplayRows
        End With
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilder.BuildNewGroupRow"

End Sub

'-------------------------------------------------------------------
Public Sub QGroupHScroll(nIndex As Integer, nScrollValue As Integer)
'-------------------------------------------------------------------
' Value of a group's horizontal scroll bar has changed
' NCJ 4 Nov 02 - Dividing factor now stored in scroll bar's Tag field
'-------------------------------------------------------------------

    ' We've set it up with a dividing factor
    ' (to avoid overflow in integer scroll bar Value field)
    moVBForm.picScroller(nIndex).Left = -(nScrollValue * moVBForm.hsbQuGroup(nIndex).Tag)
    ' Scroll the headers by the same amount
    moVBForm.picHeaders(nIndex).Left = moVBForm.picScroller(nIndex).Left

End Sub

'-------------------------------------------------------------------
Public Sub QGroupVScroll(nIndex As Integer, nScrollValue As Integer)
'-------------------------------------------------------------------
' Value of a group's vertical scroll bar has changed
'-------------------------------------------------------------------
Dim oEFQGroup As EFormQGroupDE

    Set oEFQGroup = EFormQGroupByIndex(nIndex)
    ' Hide the no. of rows indicated by nScrollvalue
    moVBForm.picScroller(nIndex).Top = moVBForm.picHeaders(nIndex).Height - (nScrollValue * oEFQGroup.RowHeight)

End Sub

'-------------------------------------------------------------------
Public Sub ScrollToGroupElement(oElement As eFormElementRO, nRow As Integer)
'-------------------------------------------------------------------
' Set the scroll bars so that the given element is visible in its group box
' NCJ 10 Oct 02 - Do horiz. scrolling too
' NCJ 4 Nov 02 - Get scroll bar dividing factor from scroll bar tag
'-------------------------------------------------------------------
Dim oEFQGroup As EFormQGroupDE
Dim nCurScroll As Integer
Dim nNewScroll As Integer
Dim nIndex As Integer
Dim sglVisLeft As Single
Dim sglVisWidth As Single
Dim sglNewOffset As Single
Dim sglElementX As Single
Dim sglElementWidth As Single

    On Error GoTo ErrLabel
    
    If oElement.OwnerQGroup Is Nothing Then Exit Sub
    
    ' Pick up the element's group
    Set oEFQGroup = EFormQGroupById(oElement.OwnerQGroup.QGroupId)
    nIndex = oEFQGroup.ControlIndex
    
    ' Deal with vertical scrolling first
    ' If there's nothing to scroll, dont bother
    If oEFQGroup.CurrentRows > oEFQGroup.QGroup.DisplayRows Then
    
        nNewScroll = -1
        nCurScroll = moVBForm.vsbQuGroup(nIndex).Value
        ' It's not visible if it's in the "hidden" rows
        If nRow <= nCurScroll Then
            ' Bring it to the top by hiding the rows below it
            nNewScroll = nRow - 1
        ElseIf nRow - nCurScroll > oEFQGroup.QGroup.DisplayRows Then
            ' It's beyond the current view so we need to scroll forward to it
            ' This sets the row to be the last in view
            nNewScroll = nRow - oEFQGroup.QGroup.DisplayRows
        End If
    
        If nNewScroll >= 0 Then
            moVBForm.vsbQuGroup(nIndex).Value = nNewScroll
        End If

    End If
    
    ' NCJ 10 Oct 02 - Now check the horizontal position
    With moVBForm
        ' Have we got a scroll bar?
        If .hsbQuGroup(nIndex).Max > 0 Then
            ' Get the element's position and size
            sglElementX = oEFQGroup.ColXs(oElement.GroupElementOrder)
            sglElementWidth = oEFQGroup.ColWidths(oElement.GroupElementOrder)
            
            ' Get the current visible "window"
            sglVisLeft = Abs(.picScroller(nIndex).Left)
            sglVisWidth = .picFixed(nIndex).Width
            
            ' See if it's already visible
            If (sglElementX >= sglVisLeft) _
             And (sglElementX + sglElementWidth <= sglVisLeft + sglVisWidth) Then
                ' It's showing - nothing to do
            Else
                ' If width of field is greater than width of group, anchor to its left
                If sglElementWidth >= sglVisWidth Then
                    sglNewOffset = sglElementX
                Else    ' Place the field in the centre of the group
                    sglNewOffset = sglElementX - (sglVisWidth - sglElementWidth) / 2
                End If
                ' Scroll to new offset (we use the dividing factor in the scroll bar tag)
                sglNewOffset = sglNewOffset / .hsbQuGroup(nIndex).Tag
                ' Deal with new offset being outside scroll bar's Min/Max range
                If sglNewOffset > .hsbQuGroup(nIndex).Max Then
                    .hsbQuGroup(nIndex).Value = .hsbQuGroup(nIndex).Max
                ElseIf sglNewOffset < 0 Then
                    .hsbQuGroup(nIndex).Value = 0
                Else
                    .hsbQuGroup(nIndex).Value = sglNewOffset
                End If
             End If
             
        End If
        
    End With
    
    Set oEFQGroup = Nothing
 
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilder.ScrollToGroupElement"

End Sub

'-------------------------------------------------------------------
Public Function GetCalPopupTopLeft(nIndex As Integer, _
                ByRef sglTop As Single, ByRef sglLeft As Single)
'-------------------------------------------------------------------
' Get the Top and Left for a Calendar or Popuplist relative to picCRFPage
' Must add group offsets if element is within a group
' Assume it's always for a txtCRFElement control
'-------------------------------------------------------------------

    ' NCJ 4 Nov 02 - Use GetControlTopLeft instead
    Call GetControlTopLeft(moVBForm.txtCRFElement(nIndex), sglTop, sglLeft)
    sglTop = sglTop + moVBForm.txtCRFElement(nIndex).Height

End Function

'-------------------------------------------------------------------
Public Sub GetControlTopLeft(oControl As Control, _
                ByRef sglTop As Single, ByRef sglLeft As Single, _
                Optional ByRef sglHeight As Single)
'-------------------------------------------------------------------
' Get the Top and Left for a control on the CRFPage,
' taking into account the offsets for its container(s)
' NCJ 4 Oct 02 - Changed for new group header labels
' NCJ 3 Dec 02 - Added sglHeight, made Public
'-------------------------------------------------------------------
Dim oContainer As PictureBox

    On Error GoTo ErrLabel
    
    ' Default to its own Top, Left
    sglTop = oControl.Top
    sglLeft = oControl.Left
    sglHeight = oControl.Height
    
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
    Err.Raise Err.Number, , Err.Description & "|ElementBuilder.GetControlTopLeft"

End Sub

'---------------------------------------------------------------------
Private Sub BuildNextButton(bReadOnly As Boolean)
'---------------------------------------------------------------------
' REM 17/05/02
' Build the "Next eForm" button control
' that is tabbed to after all other controls on an eForm
' NCJ 8 Jul 02
' NCJ 5 Nov 02 - "Next" button now goes at the Max X position
' NCJ 27 Nov 02 - Don't let it go beyond the right edge of the eForm
' NCJ 17 Jan 03 - Build a "Go Back" button too
' MLM 28/03/03: Only make the buttons tabbable toable if the form is writable
'---------------------------------------------------------------------

    On Error GoTo ErrLabel

    ' Let the buttons resize themselves first
    Set moVBForm.picBackButton.Picture = frmImages.imglistStatus.ListImages("DM30_BackEForm").Picture
    Set moVBForm.cmdNextButton.Picture = frmImages.imglistStatus.ListImages("DM30_NextEFormOn").Picture
    
    With moVBForm.cmdNextButton
'        .Height = 240
'        .Width = 240
        .Top = msglMaxY + 120
        ' NCJ Compare with width of GoBack button
        If msglMaxX > moVBForm.picCRFPage.Width - 120 - moVBForm.picBackButton.Width Then
'        If msglMaxX > moVBForm.picCRFPage.Width - 120 Then
'            .Left = moVBForm.picCRFPage.Width - 120 - .Width
            .Left = moVBForm.picCRFPage.Width - 120 - moVBForm.picBackButton.Width
        Else
            .Left = msglMaxX - .Width
            ' NCJ 30 Jun 05 - For an empty form, msglmaxX may be zero - check for this
            If .Left < 0 Then .Left = 0
        End If
        .Visible = True
        .Enabled = True
        .TooltipText = "Move to next eForm in this visit"
        .TabIndex = mnTabOrder
        mnTabOrder = mnTabOrder + 1
        .TabStop = Not bReadOnly
        If (.Top + .Height) > msglMaxY Then
            msglMaxY = .Top + .Height
        End If
    End With
    
    ' NCJ 17 Jan 03 - Build a "Go Back" button too, underneath the Next button
    ' NCJ 21 Jan 03 - Use special icon
    With moVBForm.picBackButton
        .Left = moVBForm.cmdNextButton.Left
        .Top = moVBForm.cmdNextButton.Top + moVBForm.cmdNextButton.Height + 90
        .Visible = True
        .Enabled = True
        ' NB Its TooltipText gets set up in frmEFormDataEntry
        .TabIndex = mnTabOrder
        .TabStop = Not bReadOnly
'        mnTabOrder = mnTabOrder + 1
        If (.Top + .Height) > msglMaxY Then
            msglMaxY = .Top + .Height
        End If
    End With

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|ElementBuilder.BuildNextButton"
End Sub

'-------------------------------------------------------------------
Public Function GetControlFromElement(oElement As eFormElementRO, nIndex As Integer) As Control
'-------------------------------------------------------------------
' Return the control for a given Element and ControlIndex
' NCJ 3 Dec 02 - Made Public
'-------------------------------------------------------------------
Dim oControl As Control

    With moVBForm
        If oElement.QGroup Is Nothing Then
            ' Differs depending on what the element is
            Select Case oElement.ControlType
            Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX, _
                        gn_POPUP_LIST, gn_CALENDAR
                ' These all have text boxes
                Set oControl = .txtCRFElement(nIndex)
            Case gn_PUSH_BUTTONS
                Set oControl = .picCRFElement(nIndex)
            Case gn_OPTION_BUTTONS
                  Set oControl = .optCRFGroup(nIndex)
            Case gn_ATTACHMENT
                  Set oControl = .cmdAttachment(nIndex)
            End Select
        Else
            ' It's a Question Group element
             Set oControl = .picQuGroup(nIndex)
        End If
    End With

    Set GetControlFromElement = oControl
    
    Set oControl = Nothing

End Function

'-------------------------------------------------------------------
Public Property Get EFormItems() As Collection
'-------------------------------------------------------------------
' MLM 06/09/02: Added. Once the EForm is built, allow frmEFormDataEntry
'   to iterate through the items on the form.
'-------------------------------------------------------------------

    Set EFormItems = mcolFormItems
    
End Property

'-------------------------------------------------------------------
Private Function GetTabIndex(oItem As EFormItem) As Integer
'-------------------------------------------------------------------
' NCJ 3 Dec 02 - Get tab index for the given eForm item (assumed NOT to be push buttons)
'-------------------------------------------------------------------
Dim oControl As Control
Dim oPushButton As Control
Dim nTab  As Integer

    Set oControl = GetControlFromElement(oItem.Element, oItem.ControlIndex)
    GetTabIndex = oControl.TabIndex
    Set oControl = Nothing
    
End Function

