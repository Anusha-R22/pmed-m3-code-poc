VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EFormBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       EFormBuilder.cls
'   Copyright:  InferMed Ltd. 2001-2007. All Rights Reserved
'   Author:     Nicky Johns, August 2001
'   Purpose:    Builder for eForm data entry for Data Management in MACRO 2.2.
'               Based on original basBuildCRF.bas from MACRO 2.1
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 7-8 Aug 01 - Initial development based on original BuildCRF
'   NCJ 31 Aug 01 - DisplayPopupList, CurrentElementID, UpdateResponses
'   NCJ 18-20 Sep 01 - Making option & push buttons work
'   NCJ 26 Sep 01 - Added in NR & CTC label handling (untested...)
'   NCJ 4 Oct 01 - Removed unused tabCRF
'   NCJ 10 Oct 01 - ShowLabCode
'   DPH 12 Oct 01 - Added check for permission to make changes as cmdClear, cmdChangeLab
'   and cmdChangeAttachment allowed user to make change when no permission to do that
'   Code added to DisplayCurrentResponse to stop crash for view only permissions
'   NCJ 15 Oct 01 - Rationalised tab stops and tab ordering
'   NCJ 23 Oct 01 - Incorporated new OptionGroup control
' NCJ 24 Oct 01 - Only show Clear button if response has a value;
'                Fixed bug in setting focus to PushButtons
' DPH 24/10/2001 - Only show category items that are active
' MACRO 3.0
' NCJ 1 Nov 01 - Changed MACROUSBS22 to MACROUSBS30
' NCJ 5 Nov 01 - Removed default font handling because business objects now all correctly set up
' NCJ 7-8 Nov 01 - Code which builds elements moved to ElementBuilder.cls
' NCJ 8-14 Nov 01 - Major changes for support of RQGs
' NCJ 15 Nov 01 - Changed control element types to be global consts (see modMACRODM_UI)
' NCJ 23 Nov 01 - Bug fixing behaviour of Question Groups
' NCJ 2 Jan 02 - Ignore field order in ResetFormFocus
' NCJ 7 Jan 02 - UpdateResponses now takes optional Response parameter
' MLM 23/01/02: Rolled forward scrolling changes in SetPageSize from 2.2.
'
' NCJ 8 July 02 - Bringing forward changes from 2.2 (but NOT Sequential Data Entry or Multi-User)
' NCJ 15 Jul 02 - RQG WBTesting: Only show Change button for multimedia question if value exists
' ZA 29/07/2002 - Added colour properties for popup list
' NCJ 27 Aug 02 - Must not delete the lab code in ClearQuestionInfoPanel
' TA 29/08/2002 - Added SetOtherStatusImages function for Note and SDV Statuses
' MLM 30/08/02: SetUpForm helps draw header bar.
' MLM 30/08/02: EnableElement modified for new appearance of controls.
' NCJ 4 Sept 02 - Set TabStop = False for non-enterable text boxes
' MLM 12/09/02: In many places, used EFormItems instead of Elements, Responses and row numbers.
' MLM 13/09/02: Allow controls without 3D border.
' NCJ 17 Sept 02 - Do not store DEBS StudySubject objects but retrieve them by ID
' NCJ 24-25 Sept 02 - Make sure Qu Info panel correctly set up for read-only or no responses
' TA 01/10/2002: got rid of 3d appearance
'   TA 04/10/2002: Removed dtatiem panel and added consequently removed functionality to the
'                   right click on the image status
' TA 07/10/2002: Reorganised some of display code - aiming for subject specific stuff in here
'                   and general stuff in frmeFormdataEntry
' NCJ 10 Oct 02 - Added horizontal scroll for RQGs
' TA 15/10/2002: More changes to the User Interface
' TA 15/10/2002: Carefully placed doevents in InitEFI to prevent ghosting
' NCJ 21 Oct 02 - Only display category drop-down list if response is enterable
' TA 31/10/2002: ensure calendar set to correct height and width
' NCJ 22 Nov 02 - Don't display QGroup statuses
' NCJ 3 Dec 02 - Set eFormWidth in InitEFI; new GetItemTopLeftHeight routine
' NCJ 9 Dec 02 - Removed call to NewGroupRow from FocusToNextField (it's done in ValidateProposedValue)
' NCJ 9 Dec 02 - Removed unused DisplayStatusLabel and DisplayNRCTCInfo
' TA 6/1/03: tooltip text for control now set in generic DisplaycurrentResponse rather than in indivdual build code
' NCJ 21 Jan 03 - Set new picBackButton properties
' RS 05/02/03 - Adjust height of dropdown box to visible area of screen insetad of total eForm height
' MLM 27/03/03: 3.0 bug 1360: FocusToNextField & FocusToNextElement: Focus to the form date after the visit date, unless it's already answered.
' NCJ 22 May 03 - Adjusted Popuplist and Calendar display
' NCJ 1 Jul 03 - Use new Expandable property of Element (Bug 1905)
' ic 13/08/2003 bug 1946, after updating visit border, adjust visit border visible height in UpdateBorders()
' NCJ 8 July 04 - New routine UpdateSkipsDerivs which is like UpdateResponses but doesn't refresh display
'TA 24/08/2004: fixed cancelled not updating in SDV icon (bug 2377)
' NCJ 13 Jan 05 - Undid some of fix to 2324 in order to fix 2495
' NCJ 18 Jan 05 - Issue 2502 - Added CreatePopupSubList to allow keyboard selection of sub lists of categories
'   TA  1/7/2005    set font.charset to allow non western european characters
' ic 19/07/2005 added clinical coding
' NCJ 26 Jul 07 - Bug 2892 - Don't do UpdateResponses in NewGroupRow
' ic 27/02/2007 issue 2855, enter clinical response is no longer a permission
'-------------------------------------------------------------------------------------------

Option Explicit

' The VB form we're drawing to
' NB Assume this is an instance of frmFormDataEntry
Private moVBForm As frmEFormDataEntry

' The eForm and eFormInstance definitions
Private moEForm As eFormRO
Private moStudy As StudyDefRO
Private moUser As MACROUser

' NCJ 17 Sept 02
'Private moEFI As EFormInstance
Private mlEFormTaskId As Long

' NCJ 17 Sept 02
Private moSubject As StudySubject

Private moElBuilder As ElementBuilder

Private mnNextPushButtonIndex As Integer
Private mnTabOrder As Integer

' The currently selected item
Public CurrentItem As EFormItem

Private mofrmeFormTop As frmWebBrowser
Private mofrmeFormLh As frmWebBrowser

'--------------------------------------------------------------------------
Public Property Get EFI() As EFormInstance
'--------------------------------------------------------------------------
' NCJ 17 Sept 02
' The eFormInstance for this eForm item, retrieved by its eFormTaskId
'--------------------------------------------------------------------------

    Set EFI = moSubject.eFIByTaskId(mlEFormTaskId)

End Property

'----------------------------------------------------------
Public Property Get CurrentElementID() As Integer
'----------------------------------------------------------
' Get the ElementId of the currently selected item on the eForm
' Returns 0 if no item selected
'----------------------------------------------------------
    
    If CurrentItem Is Nothing Then
        CurrentElementID = 0
    Else
        CurrentElementID = CurrentItem.Element.ElementID
    End If

End Property


'----------------------------------------------------------
Public Sub Init(ofrmeFormTop As frmWebBrowser, ofrmeFormLh As frmWebBrowser)
'----------------------------------------------------------
' Non eform specific initialisation
'
'----------------------------------------------------------

    Set mofrmeFormTop = ofrmeFormTop
    Set mofrmeFormLh = ofrmeFormLh


End Sub

'----------------------------------------------------------
Public Sub InitEFI(oUser As MACROUser, oVBForm As Form, oSubject As StudySubject, oEFI As EFormInstance)
'----------------------------------------------------------
' Draw the form elements
' Assume moVBForm and moEForm are already set up
' and that the eForm has its EFormElements collection all ready
'
' MLM 12/09/02: Use DisplayCurrentResponses to populate the new controls with responses
' NCJ 3 Dec 02 - Set eFormWidth here
'----------------------------------------------------------
Dim oElement As eFormElementRO
Dim nRow As Integer
Dim oQGElement As eFormElementRO
Dim oEFQGroup As EFormQGroupDE
Dim nIndex As Integer
Dim lEFormWidth  As Long

    On Error GoTo ErrLabel
    
    'initailise module level variables
    Set moVBForm = oVBForm
    Set moUser = oUser
    Set moSubject = oSubject
    Set moStudy = oSubject.StudyDef
    Set moElBuilder = New ElementBuilder
    Set moEForm = oEFI.eForm
       
    ' No current element yet
    Set CurrentItem = Nothing

    mlEFormTaskId = oEFI.EFormTaskId

    ' NCJ 3 Dec 02 - Moved here from frmEFormDataEntry.Display
#If VTRACK = 1 Then
    lEFormWidth = glLANDSCAPE_WIDTH
#Else
    ' NCJ 4 Nov 02 - We read the width that's been set in SD
    lEFormWidth = oEFI.eForm.eFormWidth
    ' NCJ 26 Feb 03 - eFormWidth is now correctly assigned by DEBS
'    If lEFormWidth = NULL_LONG Then
'        lEFormWidth = glPORTRAIT_WIDTH
'    End If
#End If
    
    ' NCJ 3 Dec 02 - Set form width here
    oVBForm.picCRFPage.Width = lEFormWidth

    With moVBForm
        '   Hide the form until the controls have been built
        .fraCRFPage.Visible = False
        .picCRFPage.Visible = False
    End With

    ' Clear what's already there and set up for the redraw
    Call UnloadControls

    'set up the background colour for the new eForm
    With moVBForm
        'following
        .fraCRFPage.BackColor = &HEEEEEE
        .picHeaderExt.BackColor = .fraCRFPage.BackColor
        .picCRFPage.BackColor = moEForm.BackgroundColour
        .cmdNextButton.BackColor = moEForm.BackgroundColour
        .picBackButton.BackColor = moEForm.BackgroundColour
    End With
       
    ' Fill in Visit and eForm names
    moVBForm.lblVisit = "Visit: " & oEFI.VisitInstance.Visit.Name
    If oEFI.VisitInstance.CycleNo > 1 Then
        moVBForm.lblVisit = moVBForm.lblVisit & "(" & oEFI.VisitInstance.CycleNo & ")"
    End If

    moVBForm.lblEForm = "eForm: " & moEForm.Name
    If oEFI.CycleNo > 1 Then
        moVBForm.lblEForm = moVBForm.lblEForm & "(" & oEFI.CycleNo & ")"
    End If

    'TA 08/05/2003: set tooltip to be read-only reason if read-only
    If oEFI.ReadOnly Then
        moVBForm.lblEForm.TooltipText = "This eForm is read only.  " & oEFI.ReadOnlyReason
    End If

    ' Build and draw all the elements
    Call moElBuilder.BuildElements(moVBForm, moSubject, EFI)

    Call ShowLabCode

    ' Fill in the Response values
    Call DisplayCurrentResponses
    
    'display the borders
    Call UpdateBorders(moSubject, oEFI)
    
    'show eform borders
    mofrmeFormTop.Show
    mofrmeFormLh.Show
    
    mofrmeFormTop.ZOrder
    mofrmeFormLh.ZOrder
   
    With moVBForm
        '   turn then on again
        .fraCRFPage.Visible = True
        .picCRFPage.Visible = True

        'TA 15/10/2002: workaround to stop ghosting
        DoEvents

    End With

    ' NB UpdateResponses won't change anything if user doesn't have rights etc.
    If CanUpdateEForm Then
        ' Calculate skips and derivations
        Call UpdateResponses(OpeningEForm)
    End If

    Set oElement = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormBuilder.InitEFI"

End Sub

'---------------------------------------------------------------------
Public Sub UpdateBorders(oSubject As StudySubject, oEFI As EFormInstance)
'---------------------------------------------------------------------
'Update the Visit and eForm lists for this eForm
'---------------------------------------------------------------------
'revisions
'ic 13/08/2003 bug 1946, after updating visit border, adjust visit border visible height
'---------------------------------------------------------------------
Dim sLabel As String
Dim vHeight As Variant
Const nSCALE = 17

    sLabel = GetSubjectString(oSubject)

    'load visits
    mofrmeFormTop.ExecuteJavaScript.fnLoadVisitBorder (sLabel), (GetWinIO.GetDelimitedVisitList(moSubject)), (oEFI.VisitInstance.VisitTaskId)
    'load efis
    mofrmeFormLh.ExecuteJavaScript.fnLoadEformBorder (GetWinIO.GetDelimitedEformList(oSubject, oEFI.VisitInstance.VisitTaskId)), (oEFI.EFormTaskId)

    'ic 13/08/2003 bug 1946, after updating visit border, adjust visit border visible height.
    'if errors, ignore and use default height
    On Error GoTo handler
    vHeight = mofrmeFormTop.ExecuteJavaScript.Document.All.divVisits.offsetHeight
    WEB_EFORM_TOP_HEIGHT = vHeight * nSCALE
    frmMenu.ResizeBottomRightHandForms
    Exit Sub
    
handler:
    WEB_EFORM_TOP_HEIGHT = DEFAULT_WEB_EFORM_TOP_HEIGHT
End Sub

'---------------------------------------------------------------------
Public Function IndexForElementRow(oElement As eFormElementRO, nRow As Integer) As Integer
'---------------------------------------------------------------------
' Returns control array index used for this element/row
'---------------------------------------------------------------------

    IndexForElementRow = moElBuilder.IndexForElementRow(oElement, nRow)

End Function

'---------------------------------------------------------------------
Public Function EFormItemByIndex(nIndex As Integer) As EFormItem
'---------------------------------------------------------------------
' Returns the element/row represented by this control array index
'---------------------------------------------------------------------

    Set EFormItemByIndex = moElBuilder.EFormItemByIndex(nIndex)

End Function

'---------------------------------------------------------------------
Public Sub GetItemTopLeftHeight(oItem As EFormItem, _
                        ByRef sglTop As Single, _
                        ByRef sglLeft As Single, _
                        ByRef sglHeight As Single)
'---------------------------------------------------------------------
' NCJ 3 Dec 02
' Return the top, left and height of the control of the eForm item (relative to picCRFPage)
'---------------------------------------------------------------------
Dim oControl As Control

    On Error GoTo ErrLabel
    
    Set oControl = moElBuilder.GetControlFromElement(oItem.Element, oItem.ControlIndex)
    Call moElBuilder.GetControlTopLeft(oControl, sglTop, sglLeft, sglHeight)
    Set oControl = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormBuilder.GetItemTopLeftHeight"
    
End Sub

'---------------------------------------------------------------------
Private Sub UnloadControls()
'---------------------------------------------------------------------
' Unload controls from form prior to rebuilding
' NCJ 23/10/00 - Clear page breaks as well
' NCJ 30 Aug 02 - Must clear new Status controls
' NCJ 9 Oct 02 - Include RQG's new HScroll and Headers
' RS 19/02/2003 - Expand Button
' ic 14/07/2005 added clinical coding
'---------------------------------------------------------------------
Dim oControl As Control
    
    On Error GoTo ErrLabel
    
    ' Ensure that all controls used to display eForm elements are unloaded
    ' (except the base element of each control array whose Index is 0)
    ' NCJ 30 Aug 02 - Added new controls to unload
    For Each oControl In moVBForm.Controls
        'ic 14/07/2005 added clinical coding
        If oControl.Name = "optGraphical" _
        Or oControl.Name = "optCRFGroup" _
        Or oControl.Name = "cmdCalendar" _
        Or oControl.Name = "cmdAttachment" _
        Or oControl.Name = "cmdPopupList" _
        Or oControl.Name = "txtCRFElement" _
        Or oControl.Name = "Line1" _
        Or oControl.Name = "lblCaption" _
        Or oControl.Name = "imgStatus" _
        Or oControl.Name = "lblNRCTC" _
        Or oControl.Name = "vsbQuGroup" _
        Or oControl.Name = "hsbQuGroup" _
        Or oControl.Name = "imgSDVStatus" _
        Or oControl.Name = "picCommentStatus" _
        Or oControl.Name = "picNoteStatus" _
        Or oControl.Name = "cmdExpand" _
        Or oControl.Name = "imgChangeCount" _
        Or oControl.Name = "imgDictionary" Then
            If oControl.Index > 0 Then
'                Debug.Print "Unloading " & oControl.Name & " " & oControl.Index
                Unload oControl
            End If
        End If
    Next
    
    ' NB We have to do SEPARATE traverses for the different sorts of controls
    ' to ensure that controls are removed in the right order,
    ' and so as not to wreak havoc with the For Each... iteration
        
    ' Ensure that the picture box containers are unloaded
    ' NCJ 9 July 02 - Have to do picCRFElement FIRST (in case there are option boxes INSIDE a picScroller)
    If moVBForm.picCRFElement.Count > 1 Then
        Call UnloadControlTypes("picCRFElement")
    End If
    
    ' Lastly we can do picScroller, picHeaders, picFixed and picQuGroup (they always come in sets)
    ' (Note that picScroller and picHeaders sit inside a picFixed which sits inside a picQuGroup)
    If moVBForm.picScroller.Count > 1 Then
        Call UnloadControlTypes("picScroller")
        Call UnloadControlTypes("picHeaders")
        Call UnloadControlTypes("picFixed")
        Call UnloadControlTypes("picQuGroup")
    End If
    
    Set oControl = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormBuilder.UnloadControls"

End Sub

'---------------------------------------------------------------------
Private Sub UnloadControlTypes(sControlName As String)
'---------------------------------------------------------------------
' Unloads all controls on the form whose name matches that given
' and with Index > 0
'---------------------------------------------------------------------
Dim oControl As Control

    For Each oControl In moVBForm.Controls
        If oControl.Name = sControlName Then
            If oControl.Index > 0 Then
                Unload oControl
            End If
        End If
    Next
    
    Set oControl = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormBuilder.UnloadControlTypes(" & sControlName & ")"

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
    Set moUser = Nothing
    Set moStudy = Nothing
    Set CurrentItem = Nothing
    
    Set mofrmeFormTop = Nothing
    Set mofrmeFormLh = Nothing
    
    If Not moElBuilder Is Nothing Then
        Call moElBuilder.Terminate
        Set moElBuilder = Nothing
    End If
    Set moSubject = Nothing

End Sub

'---------------------------------------------------------------------
Public Sub DisplayCurrentResponses()
'---------------------------------------------------------------------
' NCJ 10 Jun 02 - CBB 2.2.13/63 Refresh display for all responses
' To be used e.g. after user presses f7
' NCJ 8 July 02 - Reimplemented here for 3.0
 'MLM 06/09/02: Iterate though all items, instead of responses and Q groups for the main form.
'---------------------------------------------------------------------
Dim oItem As EFormItem

On Error GoTo ErrLabel

    For Each oItem In moElBuilder.EFormItems
        Call DisplayCurrentResponse(oItem)
    Next oItem

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormBuilder.DisplayCurrentResponses"
    
End Sub

'---------------------------------------------------------------------
Public Sub DisplayCurrentResponse(oItem As EFormItem, Optional bCurrentResponse As Boolean = False)
'---------------------------------------------------------------------
' Display current response value for this EFormItem
' If bCurrentResponse = TRUE then we also set the status in the LH question panel
' NCJ 4 Jan 02 - Deal with displaying status/enabling for Question Groups
' MLM 06/09/02: Pass in an EFormItem instead of an Element and a Row.
' NCJ 22 Nov 02 - We don't show RQG statuses
'---------------------------------------------------------------------
Dim nControlType As Integer
Dim oResponse As Response
Dim nIndex As Integer
Dim nStatus As Integer
Dim oQGI As QGroupInstance

    On Error GoTo ErrLabel
    
    ' Non-questions don't have statuses
    If oItem.Element.QuestionId <= 0 Then Exit Sub
    
    'TA 6/1/02: set tooltip first - there may nto be a response
     Call SetToolTipText(oItem)
    
    ' Is there is a Response object for this element?
    ' NCJ 20 Nov 01 - Responses are now automatically created by Business Layer,
    ' taking into account the user's ChangeData permission
    Set oResponse = oItem.Response 'moEFI.Responses.ResponseByElement(oElement, nRow)
    If oResponse Is Nothing Then
        ' DPH - Put in to stop form crashing if user has no permission to create data for
        ' a new question added to form since last form update
        Call EnableElement(oItem, False)
        Exit Sub
    End If
    
    ' Only display non-hidden questions
    If (Not oItem.Element.Hidden) Then
    
        nControlType = oItem.Element.ControlType
        ' Get this element's control array index
        nIndex = oItem.ControlIndex
        
        With moVBForm
            Select Case nControlType
            Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX, _
                    gn_POPUP_LIST, gn_CALENDAR
                ' Just fill in the text field
                .txtCRFElement(nIndex).Text = oResponse.Value
                .txtCRFElement(nIndex).Tag = oResponse.Value
                 
            Case gn_OPTION_BUTTONS
                Call SetOptionGroupValue(nIndex, oResponse.ValueCode)
                
            Case gn_PUSH_BUTTONS
                Call DisplayButtonResponse(nIndex, oResponse.ValueCode)
                              
            Case gn_ATTACHMENT
                If oResponse.Value > "" Then
                    .cmdAttachment(nIndex).Caption = gsVIEW_FILE
                    .cmdAttachment(nIndex).Tag = oResponse.Value
                Else
                    .cmdAttachment(nIndex).Caption = gsATTACH_FILE
                    .cmdAttachment(nIndex).Tag = ""
                End If
            End Select
        End With
        
        
        ' Set its status image
        
'25/10/2001: VTRACK Changes ref 1.4
#If VTRACK <> 1 Then
    'only show if not VTRACK
        If oItem.Element.QGroup Is Nothing Then
            'only do this for non RQGs
            ' TA 29/08/2002 - Added SetOtherStatusImages function for Note and SDV Statuses
            'MLM 16/09/02: Pass in border style here so that form and visit dates aren't shown as 3D
            Call SetOtherStatusImages(oResponse.SDVStatus, oResponse.NoteStatus, oResponse.Comments <> "", nIndex, nControlType, _
                IIf(oItem.Element.ElementUse = eElementUse.EFormVisitDate, 0, 1)) '0=None, 1=Fixed single
        End If
        If oItem.Element.ShowStatusFlag Then
            With oResponse
                Call SetStatusImage(.Status, moVBForm.imgStatus(nIndex), .DiscrepancyStatus, .LockStatus, moVBForm.imgSDVStatus(nIndex), .SDVStatus, moVBForm.imgChangeCount(nIndex), .ChangeCount)
            End With
            Call SetNRCTCText(oItem.Element, oResponse)
        End If
#End If

    If oItem.Element.DataType = eDataType.Thesaurus Then
        Call SetClinicalCodingImage(oItem.Response.CodingStatus, moVBForm.imgDictionary(nIndex))
    End If


        
        ' Enable it if the Response is enterable
        Call EnableElement(oItem, oResponse.Enterable)

    End If
       
    Set oResponse = Nothing
       
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & _
                "|EFormBuilder.DisplayCurrentResponse for " & oItem.Element.Code & "," & oItem.RowNo

End Sub

'----------------------------------------------------------
Private Sub SetOptionGroupValue(nElementId As Integer, ByVal sResponseCode As String)
'----------------------------------------------------------
' Select the option button corresponding to this sResponseCode
' For each optionbutton see if its Tag matches sResponseCode
'----------------------------------------------------------
Dim oOptGroup As OptionGroup
Dim i As Integer

    Set oOptGroup = moVBForm.optCRFGroup(nElementId)
    If sResponseCode > "" Then
        For i = 0 To oOptGroup.Buttons
            If oOptGroup.TagValue(i) = sResponseCode Then
                oOptGroup.SetSelected (i)
                Exit For        ' We found the one we want
            End If
        Next
    Else
        ' Nothing selected
        oOptGroup.UnselectAll
    End If
    
End Sub

'----------------------------------------------------------
Private Sub DisplayButtonResponse(ByVal nElementId As Integer, _
                                ByVal sResponseCode As String)
'----------------------------------------------------------
' Display option (graphical push buttons) value
' (See DisplayOptionResponse for setting radio buttons)
' For this ElementID, select the button whose Tag field is sResponseCode
' and deselect all the others
' NB If sResponseCode is "", no buttons get selected
'----------------------------------------------------------
Dim oControl As Control

    For Each oControl In moVBForm.optGraphical
        Call SetButtonValue(oControl, nElementId, sResponseCode)
    Next

End Sub

'----------------------------------------------------------
Private Sub SetButtonValue(ByRef oControl As Control, ByVal nElementId As Integer, _
                                ByVal sResponseCode As String)
'----------------------------------------------------------
' oControl is a push button (optGraphical)
' If the control's container matches the nElementId
' then set the Control to TRUE or FALSE depending on whether
' its Tag matches sResponseCode
' Do nothing if already correctly set
'----------------------------------------------------------

    If oControl.Index > 0 Then

        If oControl.Container.Index = nElementId Then
            ' If control belongs to this Element, check its Tag
            If oControl.Tag = sResponseCode Then
                ' Only change it if necessary
                If oControl.Value = False Then
                    oControl.Value = True
                End If
            Else
                ' Not the current value so make sure deselected
                If oControl.Value = True Then
                    oControl.Value = False
                End If
            End If
        End If
        
    End If

End Sub

'---------------------------------------------------------------------
Private Sub EnableElement(oItem As EFormItem, ByVal bEnabled As Boolean)
'---------------------------------------------------------------------
' Enable or Disable an eForm element
'
' MLM 30/08/02: Enabled text boxes should have white backgrounds.
'   Disabled text boxes should have grey backgrounds and be locked rather than disabled.
'   The shade of grey depends on whether the element is the form date.
' MLM 06/09/02: Pass in an EFormItem instead of an Element and a Row.
' NCJ 14 Feb 03 - Always enable the "View Attachment" button if there's an attachment
' ic 15/07/2005 added clinical coding
'ic 27/02/2007 issue 2855, enter clinical response is no longer a permission
'---------------------------------------------------------------------
Dim oControl As Control
Dim nControlType As Integer
Dim nIndex As Integer

    On Error GoTo ErrLabel
    
    nControlType = oItem.Element.ControlType
    ' Get this element's control array index
    nIndex = IndexForElementRow(oItem.Element, oItem.RowNo)
    
    With moVBForm
    
        ' Only non-group items have captions
'        If oElement.OwnerQGroup Is Nothing And oElement.ElementUse <> EFormVisitDate Then
'            .lblCaption(nIndex).Enabled = bEnabled
'        End If
        Select Case nControlType
        Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX, gn_POPUP_LIST, gn_CALENDAR
        
            'ic 27/02/2007 issue 2855, enter clinical response is no longer a permission. question must
            'be disabled if it is a thesaurus type and clinical coding is not installed and this is not
            'a site. this is because a network client without clinical coding installed cannot reset
            'the coding status when a response changes.
            'ic 15/07/2005 added clinical coding
            If ((oItem.Element.DataType = eDataType.Thesaurus) _
            And (Not gblnRemoteSite) And (Not gbClinicalCoding)) Then
                bEnabled = False
            End If
        
            'MLM 30/08/02: Different way of enabling controls
            .txtCRFElement(nIndex).Enabled = True
            .txtCRFElement(nIndex).TabStop = bEnabled
            If bEnabled Then
                .txtCRFElement(nIndex).Locked = False
                .txtCRFElement(nIndex).BackColor = vbWhite
            Else
                .txtCRFElement(nIndex).Locked = True
                .txtCRFElement(nIndex).BackColor = glLIGHT_GREY
            End If
            'set the colour of the notestatus image accordingly
            .picCommentStatus(nIndex).BackColor = .txtCRFElement(nIndex).BackColor
            .picNoteStatus(nIndex).BackColor = .txtCRFElement(nIndex).BackColor
            
            If nControlType = gn_POPUP_LIST Then
                .cmdPopupList(nIndex).Enabled = bEnabled
            End If
            
            If nControlType = gn_CALENDAR Then
                .cmdCalendar(nIndex).Enabled = bEnabled
            End If
            
            ' RS 19/02/2003: Added Expand Button
'            If nControlType = gn_TEXT_BOX And (oItem.Element.DisplayLength > 0) Then
            If oItem.Element.Expandable Then
                ' Only enable if text is really wider than the textbox
                .FontName = .txtCRFElement(nIndex).FontName
                .Font.Charset = 1
                .FontBold = .txtCRFElement(nIndex).FontBold
                .FontItalic = .txtCRFElement(nIndex).FontItalic
                .FontSize = .txtCRFElement(nIndex).FontSize
                If .TextWidth(.txtCRFElement(nIndex).Text) + 50 > .txtCRFElement(nIndex).Width Then
                    .cmdExpand(nIndex).Enabled = True
                Else
                    .cmdExpand(nIndex).Enabled = False
                End If
            End If
            
        Case gn_PUSH_BUTTONS
            .picCRFElement(nIndex).Enabled = bEnabled
            For Each oControl In .optGraphical
                If oControl.Index > 0 Then
                    If oControl.Container.Index = nIndex Then
                        oControl.Enabled = bEnabled
                    End If
                End If
            Next
        
        Case gn_OPTION_BUTTONS
            ' This deals with all the buttons in the group
            .optCRFGroup(nIndex).Enabled = bEnabled

        Case gn_ATTACHMENT
            ' NCJ 14 Feb 03 - If there's an attachment, always let the user view it
            If .cmdAttachment(nIndex).Caption = gsVIEW_FILE Then
                .cmdAttachment(nIndex).Enabled = True
            Else
                .cmdAttachment(nIndex).Enabled = bEnabled
            End If
            
        End Select

    End With
   
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & _
                    "|EFormBuilder.EnableElement for " & oItem.Element.Code & "," & oItem.RowNo

End Sub

'---------------------------------------------------------------------
Private Sub SetOtherStatusImages(enSDVStatus As MACRODEBS30.eSDVStatus, _
                                    enNoteStatus As MACRODEBS30.eNoteStatus, _
                                    bHasComment As Boolean, _
                                    nIndex As Integer, _
                                    nControlType As Integer, _
                                    nDefaultBorderStyle As Integer)
'---------------------------------------------------------------------
' this currently only does notes/comments
'
' MLM 13/09/02: Added nDefaultBorderStyle parameter; pass through to SetSDVStatusGraphics
'---------------------------------------------------------------------
Dim oControl As Control

    On Error GoTo ErrLabel

    Select Case nControlType
    Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX, gn_POPUP_LIST, gn_CALENDAR
        Set oControl = moVBForm.txtCRFElement(nIndex)
    Case gn_OPTION_BUTTONS
        Set oControl = moVBForm.optCRFGroup(nIndex)
    Case gn_PUSH_BUTTONS
        Set oControl = moVBForm.picCRFElement(nIndex)
    Case gn_ATTACHMENT
        Set oControl = moVBForm.cmdAttachment(nIndex)
    End Select

    'display SDV status
    
    'TA 18/10/2002: do not draw a border
    'Call SetSDVStatusGraphics(oControl, moVBForm.shpSDVStatus(nIndex), enSDVStatus, nDefaultBorderStyle)
    
    'display comment status
    If bHasComment Then
        moVBForm.picCommentStatus(nIndex).Picture = frmImages.imglistStatus.ListImages(DM30_ICON_COMMENT).Picture
        moVBForm.picCommentStatus(nIndex).Visible = True
    Else
        moVBForm.picCommentStatus(nIndex).Visible = False
    End If
    
    'display note status
    If (enNoteStatus = MACRODEBS30.eNoteStatus.nsHasNote) Then
        moVBForm.picNoteStatus(nIndex).Picture = frmImages.imglistStatus.ListImages(DM30_ICON_NOTE).Picture
        moVBForm.picNoteStatus(nIndex).Visible = True
    Else
        moVBForm.picNoteStatus(nIndex).Visible = False
    End If

Exit Sub

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "eFormBuilder.SetOtherStatusImages"

End Sub

'---------------------------------------------------------------------
Public Function SetStatusImage(ByVal nStatus As eStatus, _
                                ByRef oImageControl As Image, _
                                enDiscrepancyStatus As eDiscrepancyStatus, _
                                ByVal enLockStatus As eLockStatus, _
                                oimgSDVStatus As Image, enSDVStatus As eSDVStatus, _
                                Optional oimgChangeCount As Image = Nothing, Optional nChangeCount As Integer = 0)
'---------------------------------------------------------------------
' Set the appropriate picture for this status in the given Image control
' TA 2-4 Sept 02 Include new status icons and ChangeCount icon
'---------------------------------------------------------------------

    On Error GoTo ErrLabel


    'set discrepancy icon
    Call SetSDVStatusImage(enSDVStatus, oimgSDVStatus)
    
    With oImageControl
'        If enSDVStatus = eSDVStatus.ssQueried Then
'            'display new statuses
'            'a queried SDV overrides everything
'            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_QUERIED_SDV).Picture
'        Else
            'display normal icons
            If enLockStatus = eLockStatus.lsUnlocked Then
                If enDiscrepancyStatus = eDiscrepancyStatus.dsRaised Or enDiscrepancyStatus = eDiscrepancyStatus.dsResponded Then
                    Call SetDiscStatusImage(oImageControl, enDiscrepancyStatus)
                Else
                    Call SetBaseStatusImage(nStatus, oImageControl)
                End If
            Else
                'show lock status icon
                Call SetLockStatusImage(oImageControl, enLockStatus)
            End If
'        End If
        .Visible = True
    End With
    
    If Not oimgChangeCount Is Nothing Then
        'set change count image
        Call SetChangeCountImage(oimgChangeCount, nChangeCount)
    End If

    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|eFormBuilder.SetStatusImage"


End Function

'---------------------------------------------------------------------
Public Sub SetClinicalCodingImage(ByVal nStatus As Integer, ByRef oImageControl As Image)
'---------------------------------------------------------------------
' ic 30/08/2005 added clinical coding
'---------------------------------------------------------------------
    On Error GoTo ErrLabel

    If Not oImageControl Is Nothing Then
        Select Case nStatus
        Case eCodingStatus.csValidated
            oImageControl.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_DICTIONARY_VALIDATED).Picture
        Case eCodingStatus.csAutoEncoded, eCodingStatus.csCoded
            oImageControl.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_DICTIONARY_CODED).Picture
        Case eCodingStatus.csPendingNewCode
            oImageControl.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_DICTIONARY_PENDING).Picture
        Case eCodingStatus.csDoNotCode
            oImageControl.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_DICTIONARY_DONOT).Picture
        Case eCodingStatus.csEmpty, eCodingStatus.csNotCoded, eCodingStatus.csPendingNewCode
            oImageControl.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_DICTIONARY).Picture
        End Select
        oImageControl.TooltipText = GetCodingStatusString(nStatus)
    End If


    Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|eFormBuilder.SetClinicalCodingImage"

End Sub

'---------------------------------------------------------------------
Public Function SetSDVStatusImage(enSDVStatus As eSDVStatus, oimgSDVStatus As Image) As Boolean
'---------------------------------------------------------------------
' Set the appropriate SDV lock picture for this status in the given Image control (and picture box)
'only affects the picture box if passed in
' only shows planned icon if no picture box
'Reutns tru
'TA 24/08/2004: fixed cancelled not updating icon bug
'---------------------------------------------------------------------
        

    On Error GoTo ErrLabel

        
        If Not oimgSDVStatus Is Nothing Then
            'default icon showing to true (for queried, planned, done cases)
            oimgSDVStatus.Visible = True
            Select Case enSDVStatus
                Case eSDVStatus.ssQueried
                    oimgSDVStatus.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_QUERIED_SDV).Picture
                Case MACRODEBS30.ssPlanned 'display green underline if planned SDV
                    oimgSDVStatus.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_PLANNED_SDV).Picture
                Case MACRODEBS30.ssComplete
                    oimgSDVStatus.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_DONE_SDV).Picture
                Case Else
                    'None or cancelled - don't show icon
                    oimgSDVStatus.Visible = False
            End Select
        End If


    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|eFormBuilder.SetSDVStatusImage"
End Function

'---------------------------------------------------------------------
Public Sub SetDiscStatusImage(ByRef oImageControl As Image, ByVal enDiscrepancyStatus As MACRODEBS30.eDiscrepancyStatus)
'---------------------------------------------------------------------
' Set the appropriate discrepancy lock picture for this status in the given Image control
'---------------------------------------------------------------------

    On Error GoTo ErrLabel

    With oImageControl

        Select Case enDiscrepancyStatus
        Case eDiscrepancyStatus.dsRaised
            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_RAISED_DISC).Picture
        Case eDiscrepancyStatus.dsResponded
            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_RESPONDED_DISC).Picture
        End Select

    End With

    Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|eFormBuilder.SetDiscStatusImage"


End Sub


'---------------------------------------------------------------------
Public Sub SetLockStatusImage(ByRef oImageControl As Image, ByVal enLockStatus As eLockStatus)
'---------------------------------------------------------------------
' Set the appropriate lock picture for this status in the given Image control
'---------------------------------------------------------------------

    On Error GoTo ErrLabel

    With oImageControl

        Select Case enLockStatus
        Case eLockStatus.lsFrozen
            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_FROZEN).Picture
        Case eLockStatus.lsLocked
            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_LOCKED).Picture
        End Select

    End With

    Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|eFormBuilder.SetLockStatusImage"


End Sub
'---------------------------------------------------------------------
Public Sub SetBaseStatusImage(ByVal nStatus As eStatus, _
                                ByRef oImageControl As Image)
'---------------------------------------------------------------------
' Set the appropriate picture for this status in the given Image control
' this just does the MACRO statuses
'---------------------------------------------------------------------

    On Error GoTo ErrLabel

    With oImageControl
            'display normal icons
        Select Case nStatus
        Case eStatus.Success
            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_OK).Picture
        Case eStatus.Inform
            ' NCJ 21/4/00 Check user permissions
            If moUser.CheckPermission(gsFnMonitorDataReviewData) Then
                .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_INFORM).Picture
            Else
                .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_OK).Picture
            End If
        Case eStatus.Missing
            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_MISSING).Picture
        Case eStatus.Unobtainable
            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_UNOBTAINABLE).Picture
        Case eStatus.NotApplicable
            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_NA).Picture
        Case eStatus.Warning
            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_WARNING).Picture
        Case eStatus.OKWarning
            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_OK_WARNING).Picture
        Case eStatus.InvalidData
            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_INVALID).Picture
        Case Else
            .Picture = Nothing
        End Select
    
    End With

    Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|eFormBuilder.SetBaseStatusImage"


End Sub



'---------------------------------------------------------------------
Public Function SetChangeCountImage(oimgChangeCount As Image, nChangeCount As Integer)
'---------------------------------------------------------------------
' Set the appropriate picture for this status in the given Image control
' TA 2-4 Sept 02 Include new status icons and ChangeCount icon
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    

    'set change count image
    With oimgChangeCount
        ' nChangeCount is how many (non-"Missing") response rows are in the DIRH table
        ' Don't include current response in "Prev. responses" so subtract one
        Select Case nChangeCount - 1
        Case Is < 1
            .Visible = False
        Case 1
            .Visible = True
            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_CHANGE_COUNT1).Picture
        Case 2
            .Visible = True
            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_CHANGE_COUNT2).Picture
        Case Is >= 3
            .Visible = True
            .Picture = frmImages.imglistStatus.ListImages(DM30_ICON_CHANGE_COUNT3).Picture
        End Select
        
    End With
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|eFormBuilder.SetChangeCountImage"


End Function
'---------------------------------------------------------------------
Public Sub ElementGotFocus(oItem As EFormItem)
'---------------------------------------------------------------------
' Focus received by eForm Element
' Display appropriate stuff in Question info Panel
' NCJ 24 Sept 02 - Deal with case where Response is Nothing
' NCJ 21 Oct 02 - Do not display popup list for non-enterable fields!
'---------------------------------------------------------------------
Dim oResponse As Response
Dim oElement As eFormElementRO

    Set CurrentItem = oItem
    Set oElement = oItem.Element
    
    
    
    
    ' If a group element, make sure it's visible
    Call moElBuilder.ScrollToGroupElement(oElement, oItem.RowNo)
    
    ' Pick up the relevant response
    ' NB oResponse may be Nothing
    Set oResponse = oItem.Response 'oefi.Responses.ResponseByElement(oElement, oItem.RowNo)

    With moVBForm

        If PopupListVisible Then
            Call HidePopupList
        End If
    
        ' Set them up if there is a response
        If Not oResponse Is Nothing Then
            ' NCJ 9 Dec 02 - No longer have LH question panel
'            Call DisplayStatusLabel(oResponse.Status, oResponse.DiscrepancyStatus, oResponse.LockStatus)
'            Call DisplayNRCTCInfo(oElement, oResponse)
            ' NCJ 21 Oct 02 - Only display if response is enterable
            If oResponse.Enterable And oElement.ControlType = gn_POPUP_LIST Then
                DisplayPopupList oElement, oItem.ControlIndex
            End If
         End If
        
        If .calCRFElement.Visible = True Then
            .calCRFElement.Visible = False
        End If
            
        Screen.MousePointer = vbDefault
        
        
    End With

End Sub

'---------------------------------------------------------------------
Private Sub SetNRCTCText(oElement As eFormElementRO, oResponse As Response)
'---------------------------------------------------------------------
' Fill in the Normal Range/CTC information for a question
'---------------------------------------------------------------------
Dim sNRStatus As String
Dim nIndex As Integer

' Width of one character of NR status
Const nNRStatusCharWidth1 = 220
' Width of two characters of NR status (less than 2 * width of 1 char!)
Const nNRStatusCharWidth2 = 400

    ' Do nothing for non-labtest questions
    If oElement.DataType <> eDataType.LabTest Then Exit Sub
    ' Get this element's control array index
    nIndex = IndexForElementRow(oElement, oResponse.RepeatNumber)
    
    With moVBForm.lblNRCTC(nIndex)
        If (oResponse.Status = eStatus.InvalidData) Then
            ' invalid
            .Visible = False
            .Caption = ""
        Else
            'set label
            sNRStatus = GetNRCTCText(oResponse.NRStatus, oResponse.CTCGrade)
            'show label if a result is returned
            If sNRStatus > "" Then
                .Caption = sNRStatus
                ' sNRStatus will be one or two characters
                If Len(sNRStatus) = 1 Then
                    .Width = nNRStatusCharWidth1
                Else
                    .Width = nNRStatusCharWidth2
                End If
                .BackColor = moVBForm.picCRFPage.BackColor
                .Visible = True
            Else
                .Visible = False
            End If
        End If
    End With

End Sub

'---------------------------------------------------------------------
Public Sub DisplayCalendar(nIndex As Integer)
'---------------------------------------------------------------------
' Display the calendar control for the element with the given control index
' NCJ 4 Nov 02 - Removed oElement argument
'---------------------------------------------------------------------
Dim sglTop As Single
Dim sglLeft As Single

Const nCalendarControlFudge = 200

    ' Get the coordinates for the calendar control
    Call moElBuilder.GetCalPopupTopLeft(nIndex, sglTop, sglLeft)

    With moVBForm.calCRFElement
    
        'TA 31/10/2002: ensure calendar set to correct height and width
        .Height = 2295
        .Width = 4095
        ' Mo Morris 6/6/00 SR 3543 Ensure that the calendar control is positioned correctly
        ' (i.e. always in view and never beyond the right margin)
        ' The nCalendarControlFudge corrects the display under a 800*600 display when
        ' the vertical scroll bar occupies some of the form space.
        If moVBForm.picCRFPage.Width - sglLeft - nCalendarControlFudge < .Width Then
            .Left = moVBForm.picCRFPage.Width - .Width - nCalendarControlFudge
        Else
            .Left = sglLeft
        End If
        
        ' NCJ 22 May 03 - Do the same adjustment for the height/top
        If moVBForm.picCRFPage.Height - sglTop - nCalendarControlFudge < .Height Then
            .Top = moVBForm.picCRFPage.Height - .Height - nCalendarControlFudge
        Else
            .Top = sglTop
        End If
        
'        .Top = sglTop
        .BackColor = moVBForm.picCRFPage.BackColor
        
        .Visible = True
        ' NCJ 16/11/01
        .ZOrder
    
    End With

End Sub

'---------------------------------------------------------------------
Public Sub CreatePopupSubList(oElement As eFormElementRO, _
                            ByVal sMatch As String)
'---------------------------------------------------------------------
' NCJ 18 Jan 05  (Issue 2502 for CRUK)
' Rebuild a popup list with ONLY items where beginning of code or value
' matches the given text string
' If sMatch = "", all (active) items will be included
'---------------------------------------------------------------------
Dim olistItem As ListItem
Dim oCatItem As CategoryItem
Dim i As Long
Dim nLen As Integer
Dim nCount As Integer

    On Error GoTo ErrLabel
    
    ' Check that it's visible!
    If Not PopupListVisible Then Exit Sub
    
    ' How many items at the moment?
    nCount = moVBForm.lvwPopupList.ListItems.Count
    
    'Clear the list first
    moVBForm.lvwPopupList.ListItems.Clear
    
    ' Do all comparisons in lower case
    sMatch = LCase(sMatch)
    
    ' Populate the list view
    For i = 1 To oElement.Categories.Count
        Set oCatItem = oElement.Categories.Item(i)
        ' Check it's Active
        If oCatItem.Active Then
            ' Does this item "match"?
            If CatItemMatch(sMatch, oCatItem.Code, oCatItem.Value) Then
                Set olistItem = moVBForm.lvwPopupList.ListItems.Add(, "K" & i, _
                                                oCatItem.Code)
                olistItem.SubItems(1) = oCatItem.Value
            End If
        End If
    Next
    
    ' Adjust height if no. of items has changed
    If moVBForm.lvwPopupList.ListItems.Count <> nCount Then
        Call AdjustPopupListHeight
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormBuilder.CreatePopupSubList"

End Sub

'---------------------------------------------------------------------
Private Function CatItemMatch(sMatch As String, sCatCode As String, sCatValue As String) As Boolean
'---------------------------------------------------------------------
' NCJ 18 Jan 05  (Issue 2502 for CRUK)
' Does this text match the start of either given code or value?
' Assume sMatch is lower case; if sMatch = "" always returns TRUE
'---------------------------------------------------------------------
Dim nLen As Integer

    On Error GoTo ErrLabel
    
    ' Return TRUE if we get a match early on
    CatItemMatch = True
    
    If sMatch = "" Then Exit Function
    
    nLen = Len(sMatch)
    ' Might the code match?
    If Len(sCatCode) >= nLen Then
        If Left(sMatch, nLen) = Left(LCase(sCatCode), nLen) Then Exit Function
    End If
    ' If not, might the value match?
    If Len(sCatValue) >= nLen Then
        If Left(sMatch, nLen) = Left(LCase(sCatValue), nLen) Then Exit Function
    End If
    ' We didn't get a match
    CatItemMatch = False
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormBuilder.CatItemMatch"
    
End Function

'---------------------------------------------------------------------
Private Sub CreatePopupList(oElement As eFormElementRO, _
                            sglTop As Single, sglLeft As Single)
'---------------------------------------------------------------------
' Craete and display a popup category list for the given Element,
' at the given coordinates (NCJ 7 Nov 01, extracted from original DisplayPopupList)
'Re-Written by Mo Morris 8/2/00, SR2930,2841
'Size the popuplist based on the specific maximum textwidth of a popuplist's
'value codes and value items.
'The popuplist is now allowed to go beyond the width of the Textbox and commandbox
'that it belongs to.
'The overall width might get restricted by the right side border of the form.
'Changes made by Mo Morris 4/1/01
'Instead of using vForm.Textheight when calculating the height of a popuplist's
'listview .ListItems(1).Height is used.
'Test for a popuplist's listview going beyond the bottom of a form added.
'TA 4/9/01: listview col sizes now down using listview library
'DPH 24/10/2001 - Added Category Item Is Active Checking
'ZA 29/07/2002 - Added colour properties for popup list
' RS 05/02/03 - Adjust height of Popup according to visible area of eForm
'               NOTE: Height of popup is also adjusted in frmEformDataEntry:ScrollIfOffScreen
'---------------------------------------------------------------------
Dim olistItem As ListItem
Dim oCatItem As CategoryItem
Dim i As Long

    On Error GoTo ErrLabel
    
    ' Set font attributes of PopupList to be same as its text box
    With moVBForm.lvwPopupList
        .Font.Name = oElement.FontName
        .Font.Charset = 1
        .Font.SIZE = oElement.FontSize
        .Font.Bold = oElement.FontBold
        .Font.Italic = oElement.FontItalic
        'Clear the list
        .ListItems.Clear
    End With
   
    ' Set font attributes of form so that form.textwidth can be used for different size fonts
    With moVBForm
        .Font.Name = oElement.FontName
        .Font.Charset = 1
        .Font.SIZE = oElement.FontSize
        .Font.Bold = oElement.FontBold
        .Font.Italic = oElement.FontItalic
    End With

    ' Populate the list view
    For i = 1 To oElement.Categories.Count
        Set oCatItem = oElement.Categories.Item(i)
        ' DPH 24/10/2001 - Active Category Item Checking
        If oCatItem.Active Then
            Set olistItem = moVBForm.lvwPopupList.ListItems.Add(, "K" & i, _
                                                oCatItem.Code)
            olistItem.SubItems(1) = oCatItem.Value
        End If
    Next
    
    ' Scroll Down eForm page if dropdown control too close to bottom of visible area
    If moVBForm.fraCRFPage.Height - sglTop - moVBForm.picCRFPage.Top < 950 Then
        ' use min function to ensure we do not set a vlaue more than max causing an error
        moVBForm.vsbCRFPage.Value = Min(moVBForm.vsbCRFPage.Max, moVBForm.vsbCRFPage.Value + 50) '(950 - (moVBForm.fraCRFPage.Height - sglTop - moVBForm.picCRFPage.Top)) / 10
    End If

    With moVBForm.lvwPopupList
        .Visible = False
        .Left = sglLeft
        .Top = sglTop
        'reset width
        .Width = 20
        
        ' NCJ 22 May 03 - Code siphoned off into separate routine
        Call AdjustPopupListHeight
        
'        ' Set Height based on ListItems.Count + 1.6 to allow space for the headers
'        ' the number of lines is then multipied against the height of the first row
'        ' Check that there is at least one category value first
'        If .ListItems.Count > 0 Then
''            .Height = (.ListItems.Count + 1.6) * .ListItems(1).Height
'            sglHeight = (.ListItems.Count + 1.6) * .ListItems(1).Height
'        Else
''            .Height = moVBForm.TextHeight("W") * 2
'            sglHeight = moVBForm.TextHeight("W") * 2
'        End If
'
'        ' Check that width of the popup list
'        ' does not put it outside the right border of the form
'        ' If so reduce the width and increase the height to allow for scroll bars
'        If .Left + .Width > moVBForm.picCRFPage.ScaleWidth Then
'            .Width = moVBForm.picCRFPage.ScaleWidth - .Left
'            'add 170 twips  to height for horizontal scrollbar
'            sglHeight = sglHeight + 170
'        End If
'
'        ' Make the popup list enabled
'        .Enabled = True
'
'        lvw_SetAllColWidths moVBForm.lvwPopupList, LVSCW_AUTOSIZE_USEHEADER
'
'
'        'TA 09/10/2001: moved this code to after the lvw_SetColwidth code to avoid the lisview expanding each time
'        ' Set the overall width of the popup list to be the sum of the headers plus
'        ' a fudge factor of 60 twips
'        ' (because listview header widths is not an exact science)
'        .Width = .ColumnHeaders(1).Width + .ColumnHeaders(2).Width + 60
'
'        'Mo Morris 4/1/01
'        'Check that the height of the popuplist does not go beyond bottom border of the form
'        'RS Next two lines would resize dropdown according to form, new code resizes according to visible area
'        'If .Height > moVBForm.picCRFPage.ScaleHeight - .Top Then
'        '    .Height = moVBForm.picCRFPage.ScaleHeight - .Top
'        If sglHeight > moVBForm.picCRFPage.Height - .Top Then
'            sglHeight = Min(sglHeight, Max(moVBForm.picCRFPage.Height - .Top - 300, 0))
'            'the above reduction of the listviews height will cause a vertical scrollbar to be displayed
'            'if the listviews width is not increased to allow for this a horizontal scrollbar will also appear
'            'check to see if there is enough space to increase the width
'            'If .Left + .Width + 240 < moVBForm.picCRFPage.ScaleWidth Then
'            '    .Width = .Width + 240
'            'End If
'            If .Left + .Width + 240 < moVBForm.picCRFPage.ScaleWidth Then
'                .Width = .Width + 240
'            End If
'        End If
        
        .Visible = True
        
        ' NCJ 16/11/01
        .ZOrder
        
        'ZA 29/07/2002 - set the font colour
        .ForeColor = oElement.FontColour
        .Refresh

    End With    ' moVBForm.lvwPopupList

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormBuilder.CreatePopupList"
End Sub

'---------------------------------------------------------------------
Public Sub AdjustPopupListHeight()
'---------------------------------------------------------------------
' NCJ 22 May 03 - Copied from CreatPopupList
' Set the correct height & width of a drop-down
'---------------------------------------------------------------------
Dim i As Long
Dim sglHeight As Long
    
    With moVBForm.lvwPopupList
        
        ' Set Height based on ListItems.Count + 1.6 to allow space for the headers
        ' the number of lines is then multipied against the height of the first row
        ' Check that there is at least one category value first
        If .ListItems.Count > 0 Then
            sglHeight = (.ListItems.Count + 1.6) * .ListItems(1).Height
        Else
            sglHeight = moVBForm.TextHeight("W") * 2
        End If
            
        ' Check that width of the popup list
        ' does not put it outside the right border of the form
        ' If so reduce the width and increase the height to allow for scroll bars
        If .Left + .Width > moVBForm.picCRFPage.ScaleWidth Then
            .Width = moVBForm.picCRFPage.ScaleWidth - .Left
            'add 170 twips  to height for horizontal scrollbar
            sglHeight = sglHeight + 170
        End If
         
        lvw_SetAllColWidths moVBForm.lvwPopupList, LVSCW_AUTOSIZE_USEHEADER
        
        'TA 09/10/2001: moved this code to after the lvw_SetColwidth code to avoid the lisview expanding each time
        ' Set the overall width of the popup list to be the sum of the headers plus
        ' a fudge factor of 60 twips
        ' (because listview header widths is not an exact science)
        .Width = .ColumnHeaders(1).Width + .ColumnHeaders(2).Width + 60
        
        'Mo Morris 4/1/01
        'Check that the height of the popuplist does not go beyond bottom border of the form
        If sglHeight > moVBForm.picCRFPage.Height - .Top Then
            sglHeight = Min(sglHeight, Max(moVBForm.picCRFPage.Height - .Top - 300, 0))
            'the above reduction of the listviews height will cause a vertical scrollbar to be displayed
            'if the listviews width is not increased to allow for this a horizontal scrollbar will also appear
            'check to see if there is enough space to increase the width
            If .Left + .Width + 240 < moVBForm.picCRFPage.ScaleWidth Then
                .Width = .Width + 240
            End If
        End If
        
        ' Set the height here
        .Height = sglHeight
    End With

End Sub

'---------------------------------------------------------------------
Private Sub DisplayPopupList(oElement As eFormElementRO, nIndex As Integer)
'---------------------------------------------------------------------
' Display PopupList for a "top level" category item,
' i.e. not belonging to a group
'---------------------------------------------------------------------
Dim sglTop As Single
Dim sglLeft As Single

    On Error GoTo ErrLabel
    
    ' Get the coords of where to display the popup list
    Call moElBuilder.GetCalPopupTopLeft(nIndex, sglTop, sglLeft)
    
    Call CreatePopupList(oElement, sglTop, sglLeft)
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormBuilder.DisplayPopupList"
    
End Sub

'---------------------------------------------------------------------
Public Sub HidePopupList()
'---------------------------------------------------------------------
' Hide the lvwPopupList control
' NCJ 6/4/00 SR3295 - Cannot use Visible property of lvwPopupList
'   because it doesn't work! Instead set height to 30 and place it off the form
'   (NB For some inexplicable reason, setting Visible = False does NOT hide the list view)
' See also PopupListVisible
'---------------------------------------------------------------------

'    lvwPopupList.Visible = False       ' NO!
    On Error GoTo ErrLabel
    
    With moVBForm
        .lvwPopupList.Height = 30    ' Reduce its size
        .lvwPopupList.Top = -30      ' Send it off-piste
        .lvwPopupList.Refresh
    End With
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormBuilder.HidePopupList"
End Sub

'---------------------------------------------------------------------
Public Function PopupListVisible() As Boolean
'---------------------------------------------------------------------
' Returns TRUE if PopupList is currently visible
' NCJ 6/4/00 SR3295 - Cannot use Visible property of lvwPopupList
' (see also HidePopupList)
'---------------------------------------------------------------------

    PopupListVisible = (moVBForm.lvwPopupList.Height > 30)

End Function

'---------------------------------------------------------------------
Public Function UpdateSkipsDerivs(ByVal nContext As eAutoRFCContext, _
                            Optional oChangedResponse As Response = Nothing) As Boolean
'---------------------------------------------------------------------
' NCJ 8 July 04
' Update skips and derivations but DON'T refresh the disply
' Use this instead of UpdateResponses
' when the display is going to be refreshed anyway immediately afterwards
' Returns TRUE if anything changed
'---------------------------------------------------------------------
Dim oVisitEFI As EFormInstance
Dim oEFI As EFormInstance
    
    Set oEFI = Me.EFI
    Set oVisitEFI = oEFI.VisitInstance.VisitEFormInstance
    If oVisitEFI Is Nothing Then
        UpdateSkipsDerivs = oEFI.RefreshSkipsAndDerivations(nContext, "", oChangedResponse)
    Else
        UpdateSkipsDerivs = oEFI.RefreshSkipsAndDerivations(nContext, "", oChangedResponse) _
                  Or oVisitEFI.RefreshSkipsAndDerivations(nContext, "", oChangedResponse)
    End If
    
    Set oEFI = Nothing
    Set oVisitEFI = Nothing

End Function

'---------------------------------------------------------------------
Public Sub UpdateResponses(ByVal nContext As eAutoRFCContext, _
                            Optional oChangedResponse As Response = Nothing)
'---------------------------------------------------------------------
' Update all derivations and skips
' and ensure fields properly enabled/disabled
' nContext is why we're doing it (opening form or changing response)
' oChangedResponse is the response that's just changed (if any):
' if oChangedResponse is given, the status of its group will be updated if appropriate
'
' MLM 06/09/02: Update responses for all EFormItems.
' MLM 13/09/02: RefreshSkipsAndDerivations for the visit eform too, if required.
' NCJ 8 July 04 - Call new routine UpdateSkipsDerivs
' NCJ 13 Jan 05 - No! Go back to RefreshSkipsAndDerivations (for bug 2495)
'---------------------------------------------------------------------

Dim bRefreshNeeded As Boolean
Dim oItem As EFormItem
Dim oVisitEFI As EFormInstance
Dim oEFI As EFormInstance

    Set oEFI = Me.EFI
    Set oVisitEFI = oEFI.VisitInstance.VisitEFormInstance
    If oVisitEFI Is Nothing Then
        bRefreshNeeded = oEFI.RefreshSkipsAndDerivations(nContext, "", oChangedResponse)
    Else
        bRefreshNeeded = oEFI.RefreshSkipsAndDerivations(nContext, "", oChangedResponse) _
                  Or oVisitEFI.RefreshSkipsAndDerivations(nContext, "", oChangedResponse)
    End If
    
    ' NCJ 13 Jan 05 - Put back to how it was before
'    bRefreshNeeded = UpdateSkipsDerivs(nContext, oChangedResponse)
    
    ' Redraw things if anything changed
    If bRefreshNeeded Then
        ' We only redraw ones which might have changed
        For Each oItem In moElBuilder.EFormItems
            If Changeable(oItem.Element, oItem.Response) Then
                Call DisplayCurrentResponse(oItem)
            End If
        Next oItem
    ElseIf Not oChangedResponse Is Nothing Then
        ' Just redisplay the group status (if any) in case it changed
        If Not oChangedResponse.Element.OwnerQGroup Is Nothing Then
            Call DisplayCurrentResponse(EFormItemByIndex(IndexForElementRow(oChangedResponse.Element.OwnerQGroup.Element, 1)))
        End If
    End If

End Sub

'---------------------------------------------------------------------
Private Function Changeable(oElement As eFormElementRO, oResponse As Response) As Boolean
'---------------------------------------------------------------------
' Is this element "changeable", i.e. is it derived or skippable,
' or is it in a group that's skippable?
' NCJ 13 Jan 05 - Added oResponse argument
'---------------------------------------------------------------------

    Changeable = True
    
    ' It's changeable if it has a Collect If or a Derivation
    If (oElement.CollectIfCond > "") Or (oElement.DerivationExpr > "") Then Exit Function
    
    ' It's changeable if it is in a group with a Collect If
    If Not oElement.OwnerQGroup Is Nothing Then
        If oElement.OwnerQGroup.Element.CollectIfCond > "" Then Exit Function
    End If
    
    If Not oResponse Is Nothing Then
    ' NCJ 13 Jan 05 - Deal with items who've had their skip conditions removed (part of bug 2495)
        If oResponse.SavedStatus = eStatus.NotApplicable Then Exit Function
    End If
    
    ' NCJ 22 Nov 02 - No longer display group statuses
    'MLM 06/09/02: always redisplay groups' statuses (because it's hard to tell if they really need doing)
'    If Not oElement.QGroup Is Nothing Then Exit Function
    
    ' Otherwise it can't change
    Changeable = False
    
End Function

'---------------------------------------------------------------------
Public Sub ResetFormFocus(oItem As EFormItem)
'---------------------------------------------------------------------
' Reset the focus to the field as identified by oItem,
' checking first that it is still focussable
'
' MLM 09/09/02: Pass oItem straight through to IsFieldFocusable and SetElementFocus
'---------------------------------------------------------------------

    ' Check first that it's still available
    ' (it ought to be!)
    ' Ignore RequiredFieldOrder in this case (set to 0)
    If IsFieldFocusable(oItem, 0) Then
        Call SetElementFocus(oItem)
    Else
        ' We need to find the "next" focusable field
        Call FocusToNextField(oItem.Element, oItem.RowNo)
    End If
    
End Sub

'---------------------------------------------------------------------
Public Sub FocusToNextField(oElement As eFormElementRO, nRow As Integer)
'---------------------------------------------------------------------
' Set the focus to the "next" eForm element,
' i.e. one with a FieldOrder *greater* than the FieldOrder of the given Element
' If not possible, go to the first available field in the eForm
' Don't try this if all items are disabled.
' If oElement = Nothing, go for first available element
' MLM 27/03/03: Restructured.
'---------------------------------------------------------------------
Dim nFieldOrder As Integer
Dim oEFI As EFormInstance
Dim oVisitDateResponse As Response
Dim oEFormDateResponse As Response
Dim oItem As EFormItem

    On Error GoTo ErrLabel
    
    ' Forget it if we're not visible
    If moVBForm.Visible = False Then
        Exit Sub
    End If

    ' NCJ 29/4/00 - Must be in Update mode AND user must have correct rights
    ' AND form must not be locked
    ' (otherwise all questions will be disabled and we
    ' won't be able to set focus to any of them)
    ' Note that individual locked/not-applicable items are dealt with separately below
    ' NCJ 28 Mar 02 - Use new CanUpdateEForm to see if we can set the focus
    If Not CanUpdateEForm Then
        ' Either user has no change data permission, or study is read-only,
        ' or eForm is locked/frozen
        ' NCJ 8 Jul 02 - Set focus to cmdNextButton, clear up and get out
        moVBForm.cmdNextButton.SetFocus
        Exit Sub
    End If

    '1. If the user is moving from a normal question, focus to the next question.
    If Not oElement Is Nothing Then
        If oElement.ElementUse = eElementUse.User Then
            If Not oElement.OwnerQGroup Is Nothing Then
                ' We're within a group - try focussing to the next field in the group
                If FocusToNextGroupField(oElement.OwnerQGroup, nRow, oElement.GroupElementOrder + 1) Then
                    Exit Sub
                End If
            End If
            FocusToNextElement oElement.ElementOrder + 1
            Exit Sub
        End If
    End If
    
    Set oEFI = Me.EFI
    If Not oEFI.VisitInstance.VisitEFormInstance Is Nothing Then
        Set oVisitDateResponse = oEFI.VisitInstance.VisitEFormInstance.Responses.ResponseByElement(oEFI.VisitInstance.VisitEFormInstance.eForm.EFormDateElement, 1)
        If Not oVisitDateResponse Is Nothing Then
            '2. If the user is opening the form and the visit date is blank and enterable, focus to that
            If oElement Is Nothing And oVisitDateResponse.Value = "" And oVisitDateResponse.Enterable Then
                Set oItem = EFormItemByIndex(IndexForElementRow(oVisitDateResponse.Element, 1))
                Call ResetFormFocus(oItem)
                
                Set oVisitDateResponse = Nothing
                Set oItem = Nothing
                Exit Sub
            End If
        End If
    End If
    
    ' NCJ 1 Apr 03 - Added check for oElement being Nothing
    Set oEFormDateResponse = oEFI.Responses.ResponseByElement(oEFI.eForm.EFormDateElement, 1)
    If Not oEFormDateResponse Is Nothing Then
        If Not oElement Is Nothing Then
            '3. If the user is moving from the form date, go to the first normal question
            If oElement.EFormId = oEFormDateResponse.Element.EFormId And oElement.ElementID = oEFormDateResponse.Element.ElementID Then
                FocusToNextElement 1
            
                Exit Sub
            End If
        End If
        '4. If the form date is blank and enterable, focus to it
        If oEFormDateResponse.Value = "" And oEFormDateResponse.Enterable Then
            Set oItem = EFormItemByIndex(IndexForElementRow(oEFormDateResponse.Element, 1))
            Call ResetFormFocus(oItem)
            Set oEFormDateResponse = Nothing
            Set oItem = Nothing
            Exit Sub
        End If
    End If
    
    '5. Otherwise, go to the 1st normal Q.
    FocusToNextElement 1

    Set oVisitDateResponse = Nothing
    Set oEFormDateResponse = Nothing
    Set oItem = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormBuilder.FocusToNextField"

End Sub

'---------------------------------------------------------------------
Public Sub NoQuestionFocus()
'---------------------------------------------------------------------
' NCJ 8 July 2002
' No question has the focus - do what we need to do
'---------------------------------------------------------------------
    
    Set CurrentItem = Nothing

End Sub

'---------------------------------------------------------------------
Private Function FocusToNextElement(nRequiredOrder As Integer) As Boolean
'---------------------------------------------------------------------
' Try to focus to the next element with top level ElementOrder
' greater than or equal to nRequiredOrder.
' Returns TRUE if we focussed to something, or FALSE if we didn't
'
' MLM 09/09/02: Use Items instead of Elements
' MLM 10/09/02: Don't focus to form or visit dates
'---------------------------------------------------------------------
Dim oElement As eFormElementRO
Dim oItem As EFormItem

    On Error GoTo ErrLabel
    
    ' Default to True for when we jump out successfully
    FocusToNextElement = True
    '   Loop through all eFormElements
    'For Each oElement In moEForm.EFormElements
    For Each oItem In moElBuilder.EFormItems
        Set oElement = oItem.Element
        ' Only look at non-hidden, non-group, non-form-visit-date elements with appropriate ElementOrder
        If (Not oElement.Hidden) And (oElement.OwnerQGroup Is Nothing) And _
            (oElement.ElementOrder >= nRequiredOrder) And oElement.ElementUse = eElementUse.User Then
            If oElement.QGroup Is Nothing Then
                ' Not a question group - see if we can focus to it
                ' (top-level so Row = 1)
                If IsFieldFocusable(oItem, nRequiredOrder) Then
                    Call SetElementFocus(oItem)
                    Exit Function
                End If
            ElseIf oElement.QGroup.Elements.Count > 0 Then
                ' If there are group elements, see if we can focus to one of them,
                ' starting at position 1 in row 1
                If FocusToNextGroupField(oElement.QGroup, 1, 1) Then
                    Exit Function
                End If
            End If

        End If
    Next    ' oElement
    
    ' If we get here, we didn't focus to anything
    FocusToNextElement = False
    'MLM 27/03/03:
    ' NCJ 8 Jul 02 - Set focus to cmdNextButton, clear up and get out
    moVBForm.cmdNextButton.SetFocus

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormBuilder.FocusToNextElement(" & nRequiredOrder & ")"

End Function

'---------------------------------------------------------------------
Private Function FocusToNextGroupField(oQGroup As QGroupRO, nRow As Integer, _
                        nGroupFieldOrder As Integer) As Boolean
'---------------------------------------------------------------------
' Focus to the next field in a group,
' where nRow is the row to start in
' and nGroupFieldOrder is the desired position in the row.
' Returns TRUE if we focussed somewhere, or FALSE if we didn't
'
' MLM 09/09/02: Loop through the Group's EFormItems instead of Elements.
'---------------------------------------------------------------------
Dim oEFQGroup As EFormQGroupDE
Dim oGElement As eFormElementRO
Dim nThisRow As Integer
'MLM 09/09/02:
Dim oItem As EFormItem

    On Error GoTo ErrLabel
    
    ' Set to true for when we exit successfully
    FocusToNextGroupField = True
    ' Pick up the eForm group object
    Set oEFQGroup = moElBuilder.EFormQGroupById(oQGroup.QGroupId)
    
    For Each oItem In oEFQGroup.EFormItems
        Select Case oItem.RowNo
        Case nRow
            'focus to an element after the current one in the same row
            If IsFieldFocusable(oItem, nGroupFieldOrder) Then
                Call SetElementFocus(oItem)
                Exit Function
            End If
        Case Is > nRow
            'or, focus to any element in subsequent rows
            If IsFieldFocusable(oItem, 1) Then
                Call SetElementFocus(oItem)
                Exit Function
            End If
        Case Else
            'never focus to elements in earlier rows of the group
        End Select
    Next oItem
    
    ' If we get here, we didn't focus to anything
    FocusToNextGroupField = False

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormBuilder.FocusToNextGroupField"

End Function

'---------------------------------------------------------------------
Private Function IsFieldFocusable(oItem As EFormItem, nRequiredFieldOrder As Integer) As Boolean
'---------------------------------------------------------------------
' Returns TRUE if it's OK to set the focus to this field
' i.e. it's a data entry question which is enabled, unlocked etc.
'
' MLM 09/09/02: Pass in an Item instead of an Element and a Row
' NCJ 24 Sept 02 - There might not be a response
'---------------------------------------------------------------------

    IsFieldFocusable = False
    
    ' Check the things that make it non-focusable and get out
    ' DataItemId must be > 0
    If oItem.Element.QuestionId <= 0 Then Exit Function
    
    ' NCJ 24 Sept 02 - It must have a response
    If oItem.Response Is Nothing Then Exit Function
    
    ' ElementOrder must be greater than or equal to what's given
    ' (for group items we must check the GroupElementOrder)
    If oItem.Element.OwnerQGroup Is Nothing _
        And oItem.Element.ElementOrder < nRequiredFieldOrder Then Exit Function
    If (Not oItem.Element.OwnerQGroup Is Nothing) _
        And oItem.Element.GroupElementOrder < nRequiredFieldOrder Then Exit Function
    
    ' Finally it must be "Enterable"
    If Not oItem.Response.Enterable Then Exit Function
    
    ' If we get through to here then it must be OK
    IsFieldFocusable = True
    
End Function

'---------------------------------------------------------------------
Private Sub SetElementFocus(oItem As EFormItem)
'---------------------------------------------------------------------
' Set focus to element on current form
' Assume oElement is enabled (i.e. do not call from SetFormFocus if not)
'
' MLM 09/09/02: Take an Item instead of an Element.
'---------------------------------------------------------------------
Dim nControlType As Integer
Dim nIndex As Integer

    nControlType = oItem.Element.ControlType
    
    nIndex = oItem.ControlIndex
 
    With moVBForm
        ' Set focus according to element type
        Select Case nControlType
        Case gn_TEXT_BOX, gn_RICH_TEXT_BOX, gn_MASK_ED_BOX, _
                gn_POPUP_LIST, gn_CALENDAR
            .txtCRFElement(nIndex).SetFocus
            
        Case gn_PUSH_BUTTONS
            Call SetPushButtonFocus(nIndex)
            
        Case gn_OPTION_BUTTONS
           .optCRFGroup(nIndex).SetFocus
     
        Case gn_ATTACHMENT
            .cmdAttachment(nIndex).SetFocus
            
        End Select

    End With
    
    ' If a group element, make sure it's visible
    Call moElBuilder.ScrollToGroupElement(oItem.Element, oItem.RowNo)
    
End Sub

'---------------------------------------------------------------------
Private Function SetFocusIfValue(ByVal nIndex As Integer, _
                            oControl As Control, _
                            ByVal bValue As Boolean) As Boolean
'---------------------------------------------------------------------
' Set focus to given optGraphical control if its value is bValue
' and its container is nIndex
' Returns TRUE if focus was set
' NCJ 24 Oct 01 - Changed oControl.Index to oControl.Container.Index
'---------------------------------------------------------------------
        
    SetFocusIfValue = False
    ' Ignore oControl with index = 0 (because it's the "base" optGraphical)
    If oControl.Index > 0 Then
        ' Check the control's container
        If oControl.Container.Index = nIndex Then
            If oControl.Value = bValue Then
                oControl.SetFocus
                SetFocusIfValue = True
            End If
        End If
    End If
    
End Function


'---------------------------------------------------------------------
Private Sub SetPushButtonFocus(ByVal nIndex As Integer)
'---------------------------------------------------------------------
' Set focus to pushbutton whose container matches nIndex
' Try to set focus to button with value TRUE
'---------------------------------------------------------------------
Dim oControl As Control

    ' Try to find one with value TRUE
    For Each oControl In moVBForm.optGraphical
        If SetFocusIfValue(nIndex, oControl, True) Then Exit Sub
    Next
    
    ' Otherwise do first FALSE one
    For Each oControl In moVBForm.optGraphical
        If SetFocusIfValue(nIndex, oControl, False) Then Exit Sub
    Next

End Sub

'---------------------------------------------------------------------
Public Sub ShowLabCode()
'---------------------------------------------------------------------
' Display the Lab associated with this form (if any)
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    With moVBForm
        If EFI.LabCode = "" Then
            .lblLabName.Caption = "Not selected"
            .cmdChangeLab.Caption = "Choose laboratory"
        Else
            .lblLabName.Caption = EFI.LabCode
            .cmdChangeLab.Caption = "Change laboratory"
        End If
    
        .lblLab.Visible = True
        .lblLabName.Visible = True
        .cmdChangeLab.Visible = True

        
        If EFI.LabExists Then
            ' Get the lab description for the tooltip text
            .lblLabName.TooltipText = EFI.LabDescription
        Else
            .lblLabName.TooltipText = "Laboratory not found"
        End If
        ' DPH 12/10/2001 Added check for permission to make changes
        ' If user can make changes
        ' NCJ 28 Mar 02 - Use new CanUpdateEForm instead
        If CanUpdateEForm Then
            ' Only enable if all lab questions are unlocked
            .cmdChangeLab.Enabled = Not AnyLockedLabQuestions
        Else
            .cmdChangeLab.Visible = False
        End If
    End With

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|EFormBuilder.ShowLabCode"

End Sub

'---------------------------------------------------------------------
Private Function AnyLockedLabQuestions() As Boolean
'---------------------------------------------------------------------
' Returns TRUE if any of the form's Lab Questions are locked or frozen
'---------------------------------------------------------------------
Dim oResponse As Response

    AnyLockedLabQuestions = False
    For Each oResponse In EFI.Responses
        If oResponse.Element.DataType = eDataType.LabTest And oResponse.LockStatus <> LockStatus.lsUnlocked Then
            AnyLockedLabQuestions = True
            Exit Function
        End If
    Next
    
End Function

'-------------------------------------------------------------------
Public Sub NewGroupRow(oElement As eFormElementRO, nRow As Integer)
'-------------------------------------------------------------------
' Add a new row of fields for a group if necessary,
' given that this element has just been filled in
' and create new responses for the fields
'
' MLM 09/09/02: Display responses for the new row using EFormItems
' NCJ 26 Jul 07 - Bug 2892 - Don't do UpdateResponses here
'-------------------------------------------------------------------
Dim oEFQGroup As EFormQGroupDE
Dim oGElement As eFormElementRO
Dim nNewRow As Integer
Dim oItem As EFormItem

    On Error GoTo ErrLabel
    
    ' If it's not a group item then we do nothing
    If oElement.OwnerQGroup Is Nothing Then Exit Sub
    
    ' Pick up its eForm group object
    Set oEFQGroup = moElBuilder.EFormQGroupById(oElement.OwnerQGroup.QGroupId)
    
    ' NCJ 24 Sept 02 - If there's no QGroup instance we do nothing
    If oEFQGroup.QGI Is Nothing Then Exit Sub
    
    ' See if a new row is ready
    If oEFQGroup.QGI.CreateNewRow Then
        ' Build the one just added
        nNewRow = oEFQGroup.CurrentRows
        Call moElBuilder.BuildNewGroupRow(oEFQGroup, nNewRow)
        ' Call UpdateResponses to ensure that derivations etc. are correctly handled for new row
        ' NCJ 16 Aug 02 - Use context of adding group row
        ' NCJ 26 Jul 07 - Bug 2892 - Don't do UpdateResponses here because it gobbles up warnings which then get lost
'        Call UpdateResponses(AddingGroupRow)
        ' Now force correct display of items in new row
        For Each oItem In oEFQGroup.EFormItems
            If oItem.RowNo = nNewRow Then
                Call DisplayCurrentResponse(oItem)
            End If
        Next oItem
    End If
    
    Exit Sub
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|eFormBuilder.NewGroupRow"
    
End Sub

'-------------------------------------------------------------------
Public Sub QGroupVScroll(nIndex As Integer, nScrollValue As Integer)
'-------------------------------------------------------------------
' Value of a group's vertical scroll bar has changed
'-------------------------------------------------------------------

    Call moElBuilder.QGroupVScroll(nIndex, nScrollValue)

End Sub

'-------------------------------------------------------------------
Public Sub QGroupHScroll(nIndex As Integer, nScrollValue As Integer)
'-------------------------------------------------------------------
' Value of a group's vertical scroll bar has changed
'-------------------------------------------------------------------

    Call moElBuilder.QGroupHScroll(nIndex, nScrollValue)

End Sub

'------------------------------------------------------------------
Public Function CanUpdateEForm() As Boolean
'------------------------------------------------------------------
' NCJ 21 Mar 02
' Can we update this eForm,
' e.g. create a new Response object for a new question
' or update derived values?
' NCJ 6 Aug 02 - Made public so we can call it from frmEFormDatEntry
'------------------------------------------------------------------

    CanUpdateEForm = False
    
    ' NCJ 24 Sept 02 - eForm must not be Read-Only
    If EFI.ReadOnly Then Exit Function
    
    ' Must have Change Data rights
    If Not moUser.CheckPermission(gsFnChangeData) Then Exit Function

    ' Study subject must not be Read-Only
    If moStudy.Subject.ReadOnly Then Exit Function
    
    ' EForm must not be locked or frozen
    If EFI.LockStatus <> eLockStatus.lsUnlocked Then Exit Function
    
    ' If we get here it's OK
    CanUpdateEForm = True

End Function



'-------------------------------------------------------------------
Private Function SetToolTipText(oItem As EFormItem)
'-------------------------------------------------------------------
'set the tooltip text for control and status image
'-------------------------------------------------------------------
Dim oCon As Control

    On Error GoTo ErrLabel
    
    If oItem.Element.QGroup Is Nothing And oItem.Element.ShowStatusFlag Then
        Set oCon = moElBuilder.GetControlFromElement(oItem.Element, oItem.ControlIndex)
        'this is a question
        If Not oCon Is Nothing Then
            oCon.TooltipText = oItem.TooltipText
            moVBForm.imgStatus(oItem.ControlIndex).TooltipText = oItem.StatusTooltipText
            moVBForm.imgChangeCount(oItem.ControlIndex).TooltipText = oItem.StatusTooltipText
        End If
    End If
    
        Exit Function
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|eFromBuilder.SetTooTipText"
    
End Function

'-------------------------------------------------------------------
Public Property Set User(oUser As MACROUser)
'-------------------------------------------------------------------
'allow init of user object
'-------------------------------------------------------------------
    Set moUser = oUser
    
End Property
