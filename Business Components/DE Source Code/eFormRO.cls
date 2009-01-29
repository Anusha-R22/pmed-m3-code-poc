VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eFormRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: eFormRO.cls
' Copyright InferMed Ltd 2001-2004, All Rights Reserved
' Author: Nicky Johns, InferMed, May 2001
' Purpose: eForm Read-only Object for MACRO 3.0
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 24-29 May 01 - Initial development
' NCJ 15 Jun 01 - PromptForDate
' NCJ 18-21 Jun 01 - eFormElements
' NCJ 3 Jul 01 - Added Key property
' NCJ 6 Aug 01 - RemoveEFormElements
' ZA 07/08/01 - Added BackgroundColour property
' NCJ 17 Sep 01 - Added comments
'               Added eFormElementByCode
' NCJ 20 Sep 01 - Added LabelIsLocal and Sequential
' TA 02/10/2001 - New error handling
' TA 25/4/2002: Retrieve elements by dataitem id rather than crfelement id
'               This id changes every time the dataitems on an eform added/removed
' MACRO 3.0
' NCJ 5 July 2002 - Added 2.2 Changes to 3.0 version
' ZA 19/07/2002 - Set up default fonts for caption/comment
' MLM 29/08/02: Added new friend property AppearsInSchedule, which is set in VisitEFormRO.SetState
'   and used in ScheduleGrid.Build to determine whether the eForm should be displayed in the schedule.
'   Also, added new EFormDateElement property so the UI can find it without looking through the elements collection.
' NCJ 19 Sept 02 - Do not load elements if already loaded
' NCJ 8 Oct 02 - Load eFormDate Element in SetState
' NCJ 1 Nov 02 - Changed eFormElementByQuestionId from Friend to Public (as needed by the Web)
' NCJ 4 Nov 02 - New eFormWidth property
' NCJ 14 Nov 02 - Added QGroups property (needed because we no longer save and load QGroupInstances)
' RS/NCJ 21 Jan 2003 - Set DateTimeType property in NewElement
' NCJ 24 Mar 04 - Ensure object collections cleared in Terminate (to prevent memory leaks)
'----------------------------------------------------

Option Explicit

' The state which contains its properties
Private moEFormState As eFormState

' The study to which this eForm belongs
' (Do we need this???)
Private moStudy As StudyDefRO

' The collecion of eForm Elements
Private mcolElements As Collection

' The collection of Group definitions
Private mcolQGroups As Collection

' NCJ 17 Sep 01 - Cross-reference elements by question code
Private mcolCodeIndexes As Collection

Private mlBackColour As Long
Private mlEFormWidth As Long    ' NCJ 26 Feb 03

' TA 25/4/2002 - Cross-reference elements by question id
Private mcolQuestionIndexes As Collection

'MLM 29/08/02:
Private mbAppearsInSchedule As Boolean
Private moEFormDateElement As eFormElementRO

'-------------------------------------------------------------------
Friend Sub LoadElements(ByVal sConnection As String)
'-------------------------------------------------------------------
' Load the eForm elements for a this eForm
' Only called by the eFormInstance class
'-------------------------------------------------------------------
' NCJ 17 Sep 01 - Create cross-references to question codes
' NCJ 5 Nov 01 - Load groups, too
' MLM 29/08/02: Set the eForm's date element property when required.
' MLM 04/09/02: Pass our EFormId to the elements we load.
' NCJ 19 Sept 02 - Only load elements if mcolElements Is Nothing
' NCJ 8 Oct 02 - Don't use oEFElState.Key because it's been hijacked for Date Element use!!!
'-------------------------------------------------------------------
Dim oEFElementsDS As StudyDefDS
Dim oEFElStates As Collection    ' of EFElementState objects
Dim oEFElement As eFormElementRO
Dim oEFELState As EFElementState
Dim oGroupStates As Collection    ' of QGroupState objects
Dim oQGroup As QGroupRO
Dim oQGroupState As QGroupState

    On Error GoTo ErrLabel
    
    ' NCJ 19 Sept 02 - Check to see if they're already loaded
    If Not mcolElements Is Nothing Then Exit Sub
    
    ' Use the Study Def data services
    Set oEFElementsDS = New StudyDefDS
    ' Get the eFormElement states
    
    Set oEFElStates = CollectionDeSerialise(oEFElementsDS.LoadElements(sConnection, Me.Study.StudyId, Me.Study.Version, Me.EFormId), True)
    Set oGroupStates = CollectionDeSerialise(oEFElementsDS.LoadGroups(sConnection, Me.Study.StudyId, Me.Study.Version, Me.EFormId), True)
    Set oEFElementsDS = Nothing
    
    Set mcolElements = New Collection
    Set mcolCodeIndexes = New Collection
    Set mcolQGroups = New Collection
    Set mcolQuestionIndexes = New Collection
    
    ' Create the Group objects first
    ' (but without their Elements collections)
    For Each oQGroupState In oGroupStates
        Set oQGroup = New QGroupRO
        Call oQGroup.SetState(oQGroupState, Me)
        mcolQGroups.Add oQGroup, Str(oQGroupState.QGroupID)
    Next
    
    ' Unwrap States into business objects
    For Each oEFELState In oEFElStates
        ' Create the eForm Element object
        Set oEFElement = NewElement(oEFELState)
        
        ' Add to the form's element collection, indexed by ElementID
        mcolElements.Add oEFElement, Str(oEFElement.ElementID)
        
        ' If a question then add to code cross-reference collection
        If oEFElement.QuestionId > 0 Then
            mcolCodeIndexes.Add oEFElement, oEFElement.Code
            ' and dataitemid collection
            mcolQuestionIndexes.Add oEFElement, Str(oEFElement.QuestionId)
        End If
        
        ' If it is a group, make sure the group
        ' and the element know about each other
        If oEFELState.QGroupID > 0 Then
            Call oEFElement.SetGroup(mcolQGroups.Item(Str(oEFELState.QGroupID)))
        End If
        
        ' If it belongs to a group, set the element's owner group
        ' and add it to that group
        If oEFELState.OwnerQGroupID > 0 Then
            Set oQGroup = mcolQGroups.Item(Str(oEFELState.OwnerQGroupID))
            oQGroup.Elements.Add oEFElement, Str(oEFElement.ElementID)
            Call oEFElement.SetOwnerQGroup(oQGroup)
        End If
        
        'MLM 29/08/02: If the element is the eForm's visit date, set the eForm's EFormDateElement property
        If oEFELState.ElementUse = eElementUse.EFormVisitDate Then
            Set moEFormDateElement = Nothing
            Set moEFormDateElement = oEFElement
        End If
    Next
    
    ' Tidy up
    Set oEFELState = Nothing
    Set oEFElStates = Nothing
    Set oEFElement = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "eFormRO.LoadElements"
    
End Sub

'-------------------------------------------------------------------
Private Function NewElement(oEFELState As EFElementState) As eFormElementRO
'-------------------------------------------------------------------
' NCJ 8 Oct 02 - Code extracted from LoadElements
' Create a new eFormElementRO and set its State
' NCJ 21 Jan 03 - Set the DateTimeType too
' NCJ 13 May 03 - Set the local dec.pt. and thou.sep. too
'-------------------------------------------------------------------
Dim oEFElement As eFormElementRO

    On Error GoTo ErrLabel
    
    ' Create the eForm Element object
    Set oEFElement = New eFormElementRO
    'MLM 04/09/02: Added EFormId:
    'MLM 26/09/02: Pass through the entire form state instead.
    ' NCJ 13 May 03 - Pass through dec.pt. and thou.sep.
    Call oEFElement.SetState(oEFELState, moEFormState, _
                            moStudy.LocalDecimalPoint, moStudy.LocalThousandSeparator)
    ' Set up default font details
    Call oEFElement.SetUpFontDetails(moStudy.FontName, moStudy.FontSize, moStudy.FontColour)
    ' ZA 19/07/2002-  Set up default font details for caption and comment
    Call oEFElement.SetUpCaptionFontDetails(moStudy.FontName, moStudy.FontSize, moStudy.FontColour)
    
    ' NCJ 21 Jan 03 - We can only calculate DateTimeType if we have Arezzo (which we don't at this point in Web!)
    If oEFElement.DataType = eDataType.DateTime Then
        If Not moStudy.Arezzo_DM Is Nothing Then
            ' Set the Exact DateTime type according to AREZZO
            oEFElement.SetDateTimeType moStudy.Arezzo_DM.ValidateDateFormat(oEFElement.Format)
        End If
    End If
        
    Set NewElement = oEFElement
    
    Set oEFElement = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "eFormRO.NewElement"
    
End Function

'----------------------------------------------------
Public Property Get EFormDateElement() As eFormElementRO
'----------------------------------------------------
' MLM 29/08/02: Added.
'----------------------------------------------------

    Set EFormDateElement = moEFormDateElement
    
End Property

'----------------------------------------------------
Friend Property Let AppearsInSchedule(bAppearsInSchedule As Boolean)
'----------------------------------------------------
' MLM 29/08/02: Added.
'----------------------------------------------------

    mbAppearsInSchedule = bAppearsInSchedule

End Property

'----------------------------------------------------
Friend Property Get AppearsInSchedule() As Boolean
'----------------------------------------------------
' MLM 29/08/02: Added.
'----------------------------------------------------

    AppearsInSchedule = mbAppearsInSchedule
    
End Property

'----------------------------------------------------
Public Property Get Study() As StudyDefRO
'----------------------------------------------------
' The study definition object to whcih this eForm belongs
'----------------------------------------------------

    Set Study = moStudy
    
End Property

'----------------------------------------------------
Friend Sub SetState(oStudy As StudyDefRO, oEFState As eFormState, oEFELState As EFElementState)
'----------------------------------------------------
' Set the State object containing properties
' and initialise ourselves
' NCJ 8 Oct 02 - Added oEFElState (which may be Nothing) which represents the eForm's Date Element
'----------------------------------------------------
Const lPORTRAIT_WIDTH = 8515    ' NCJ 26 Feb 03 - Copied from basMainMACROModule

    Set moEFormState = oEFState
    Set moStudy = oStudy
    ' Set up background colour
    If moEFormState.BackgroundColour = 0 Then
        mlBackColour = oStudy.eFormColour
    Else
        mlBackColour = moEFormState.BackgroundColour
    End If
    
    ' NCJ 8 Oct 02 - Add in the eFormDateElement if nothing
    Set moEFormDateElement = Nothing
    If Not oEFELState Is Nothing Then
        Set moEFormDateElement = NewElement(oEFELState)
    End If
    
    ' NCJ 26 Feb 03 - Set up eForm width (which may be stored as NULL)
    mlEFormWidth = moEFormState.eFormWidth
    If mlEFormWidth = NULL_LONG Then
        mlEFormWidth = lPORTRAIT_WIDTH
    End If
    
End Sub

'-------------------------------------
Friend Property Get Key() As String
'-------------------------------------
' Get unique string key for this object
'-------------------------------------

    Key = moEFormState.Key

End Property

'-------------------------------------
Public Property Get Name() As String
'-------------------------------------
' The name of the eForm
'-------------------------------------

    Name = moEFormState.Name
    
End Property

'-------------------------------------
Public Property Get BackgroundColour() As Long
'-------------------------------------
' The background colour for the eForm
'-------------------------------------

    BackgroundColour = mlBackColour
    
End Property

'-------------------------------------
Public Property Get Code() As String
'-------------------------------------
' The unique text code for the eForm
'-------------------------------------
    
    Code = moEFormState.Code

End Property

'-------------------------------------
Public Property Get EFormId() As Long
'-------------------------------------
' The unique numeric ID for the eForm
'-------------------------------------

    EFormId = moEFormState.EFormId
    
End Property

'-------------------------------------
Public Property Get LabelExpr() As String
'-------------------------------------
' The Arezzo expression defining the eForm's label
'-------------------------------------

    LabelExpr = moEFormState.LabelExpr
    
End Property

'-------------------------------------
Public Property Get DateExpr() As String
'-------------------------------------
' The Arezzo expression defining the eForm's date
'-------------------------------------

    DateExpr = moEFormState.DateExpr
    
End Property

'-------------------------------------
Public Property Get LabelIsLocal() As Boolean
'-------------------------------------
' TRUE if the eForm's label is local,
' i.e. should NOT be transferred to MACRO Server
'-------------------------------------

    LabelIsLocal = moEFormState.LabelIsLocal
    
End Property

'-------------------------------------
Public Property Get PromptForDate() As Boolean
'-------------------------------------
' Whether the user should be prompted for a date
' when the eForm is first opened
'-------------------------------------

    PromptForDate = moEFormState.PromptForDate
    
End Property

'-------------------------------------
Public Property Get HideIfInactive() As Boolean
'-------------------------------------
' Whether this eForm's row in the Schedule should be hidden
' when there are no active instances of the eForm
'-------------------------------------

    HideIfInactive = moEFormState.HideIfInactive
    
End Property

'-------------------------------------
Public Property Get DisplayNumbers() As Boolean
'-------------------------------------
' Whether questions numbers should be automatically
' generated and displayed on the eForm
'-------------------------------------

    DisplayNumbers = moEFormState.DisplayNumbers
    
End Property

'-------------------------------------
Public Property Get Sequential() As Boolean
'-------------------------------------
' TRUE if eForm fields should be filled in sequentially
'-------------------------------------

    Sequential = moEFormState.Sequential

End Property

'-------------------------------------
Public Property Get eFormOrder() As Integer
'-------------------------------------
' The position of the eForm in the oder of eForms
' (as displayed on the left-hand column of the schedule)
'-------------------------------------

    eFormOrder = moEFormState.eFormOrder
    
End Property

'---------------------------------------------------
Public Property Get EFormElements() As Collection
'---------------------------------------------------
' The collection of eFormElements on this form
' (as eFormElementRO objects).
' Only available after a call to LoadElements.
'---------------------------------------------------
    
    Set EFormElements = mcolElements
    
End Property

'-------------------------------------
Public Function eFormElementById(ByVal lElementId As Long) As eFormElementRO
'-------------------------------------
' Return the eFormElement corresponding to the given Element Id
'-------------------------------------

    Set eFormElementById = mcolElements.Item(Str(lElementId))
    
End Function

'-------------------------------------
Public Function eFormElementByCode(ByVal sCode As String) As eFormElementRO
'-------------------------------------
' Return the eFormElement corresponding to the given question code
'-------------------------------------

    Set eFormElementByCode = mcolCodeIndexes.Item(sCode)
    
End Function


'-------------------------------------
Private Sub Class_Initialize()
'-------------------------------------

    Set moEFormState = New eFormState
    Set moStudy = New StudyDefRO
    Set mcolElements = Nothing      ' NCJ 19 Sept 02 Make sure they're Nothing to begin with
    Set mcolQGroups = Nothing
    Set mcolCodeIndexes = Nothing
    Set mcolQuestionIndexes = Nothing
    
    'MLM 29/08/02: eFormROs begin by assuming that they aren't required in the schedule
    mbAppearsInSchedule = False
    
End Sub

'-------------------------------------
Private Sub Class_Terminate()
'-------------------------------------
' NCJ 25 Mar 04 - Call our own Terminate
'-------------------------------------

    Call Terminate
    
End Sub
  
'-------------------------------------
Friend Sub Terminate()
'-------------------------------------
' NCJ 25 Mar 04 - Explicitly clear out collections
'-------------------------------------
    
    On Error GoTo ErrLabel
    
    Set moEFormState = Nothing
    Set moStudy = Nothing
    
    Call RemoveElements     ' NCJ 24 Mar 04
    
    ' NCJ 25 Mar 04 - Only kill the eFormDateElement when eFormRO has really gone
    If Not moEFormDateElement Is Nothing Then
        Call moEFormDateElement.Terminate
    End If
    Set moEFormDateElement = Nothing

Exit Sub
ErrLabel:
'    Debug.Print "Terminating eForm"

End Sub

'----------------------------------------------------
Friend Sub RemoveElements()
'----------------------------------------------------
' NCJ 24 Mar 04 - Ensure object collections are cleared out
' (But leave the eFormDate element alone!)
'----------------------------------------------------
Dim oElement As eFormElementRO
Dim oQGroup As QGroupRO
Dim lDateQuId  As Long

    If moEFormDateElement Is Nothing Then
        lDateQuId = 0
    Else
        lDateQuId = moEFormDateElement.QuestionId
    End If
    
    ' NB All Element objects will be terminated later
    Call CollectionRemoveAll(mcolCodeIndexes)
    Set mcolCodeIndexes = Nothing
    
    Call CollectionRemoveAll(mcolQuestionIndexes)
    Set mcolQuestionIndexes = Nothing

    If Not mcolElements Is Nothing Then
        For Each oElement In mcolElements
            If oElement.QuestionId = lDateQuId Then
                ' Leave the eFormDateElement alone!
            Else
                Call oElement.Terminate
            End If
        Next
        Call CollectionRemoveAll(mcolElements)
    End If
    Set mcolElements = Nothing
    Set oElement = Nothing
    
    If Not mcolQGroups Is Nothing Then
        For Each oQGroup In mcolQGroups
            Call oQGroup.Terminate
        Next
        Call CollectionRemoveAll(mcolQGroups)
    End If
    Set mcolQGroups = Nothing        ' Added NCJ 19 Sept 02
    Set oQGroup = Nothing

End Sub

'-----------------------------------------------------------------------
Public Function QGroupById(lQGroupID As Long) As QGroupRO
'-----------------------------------------------------------------------
' Return the question group corresponding to this QGroupID.
' Assume it exists!
'-----------------------------------------------------------------------

    Set QGroupById = mcolQGroups.Item(Str(lQGroupID))

End Function

'-------------------------------------
Friend Property Get QGroups() As Collection
'-------------------------------------
' NCJ 14 Nov 02 - The collection of QGroups on this eForm
'-------------------------------------

    Set QGroups = mcolQGroups

End Property

'-------------------------------------
Friend Function IsQuestionInEForm(ByVal lQuestionId As Long) As Boolean
'----------------------------------------------
' 'TA 25/04/2002
' Returns TRUE if this element with this id appears in this eform,
' and FALSE otherwise
' NCJ 5 July 2002 - Changed to use CollectionMember
'----------------------------------------------
    
    IsQuestionInEForm = CollectionMember(mcolQuestionIndexes, Str(lQuestionId))

End Function
 
'-------------------------------------
Public Function eFormElementByQuestionId(lQuestionId As Long) As eFormElementRO
'-------------------------------------
' Return the eFormElement corresponding to the given question code
' NCJ 1 Nov 02 - Changed from Friend to Public
'-------------------------------------

    Set eFormElementByQuestionId = mcolQuestionIndexes.Item(Str(lQuestionId))
    
End Function

'-------------------------------------
Public Property Get eFormWidth() As Long
'-------------------------------------
' NCJ 4 Nov 02 - Width of eForm
' NCJ 26 Feb 03 - Use module level variable
'-------------------------------------

    eFormWidth = mlEFormWidth

End Property
