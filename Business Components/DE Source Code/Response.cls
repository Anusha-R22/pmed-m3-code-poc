VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Response"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: Response.cls
' Author: Nicky Johns
' Copyright: InferMed, 2001 - 2007, All Rights Reserved
' Purpose: Data Entry Response Object for MACRO 2.2
' June 2002 - Now for MACRO 3.0
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 18-21 Jun 01 - Initial development
' NCJ 26 Jun 01 - Added ValueStore objects for "indicator" properties
' NCJ 3 Jul 01 - Added Key property
' NCJ 4 Jul 01 - Added SetResponseValue
' NCJ 30 Jul 01 - No longer need SetStateValues (we keep the ResponseState objects
'                   continuously updated)
' NCJ 16 Aug 01 - New SavedValue and SavedValueCode properties
' NCJ 30/8/01 - DiffersFromSavedValue
' NCJ 21 Sep 01 - Deal with multimedia questions in ValidateValue
' NCJ 24 Sep 01 - Changed UserCode to UserName
'                   Ensure user name saved when Response data changes
' NCJ 26 Sep 01 - New SavedStatus property
'               Validation of lab test questions
' NCJ 27 Sep 01 - Setting of OverruleReason, setting of Software version
'               Removal of unnecessary property Gets
'               New Element property
' NCJ 2 Oct 01 - Changed ValidateValue to use new me:value handling (see Arezzo_DM)
'               Added new error handlers
' NCJ 11 Oct 01 - Handle labcode changes correctly (and do not provide LabCode Property Let)
' NCJ 11 Oct 01 - Check length of response data in ValidateValue, and strip leading/trailing spaces
' NCJ 12 Oct 01 - Store NRStatus and CTCGrades in Arezzo when Confirming a value
' NCJ 8 Jan 02 - In ValidateValue, do not check QuestionLength for numbers
' NCJ 11 Jan 02 - Store "standardised" numbers in moResponseState.Value
' NCJ 19 Mar 02 - Calculate Arezzo user tag and dataname once only;
'                   do not re-add values to Arezzo if they haven't changed
' TA 04/04/2002: Changes for HadValue property of a response
' NCJ 8 May 02 - Added Multi User Support MUSilentUpdate
' MLM 21/05/02: Modified ValidateValue to screen out " and |.
' MLM 22/05/02: Calculate sArezzoValue differently for numbers.
' NCJ 5 Jun 02 - CBB 2.2.10/1 Ensure category codes in correct case
' REM 21/06/02 - CBB 2.2.16 No. 12 Also reject ~ character in ValidateValue
' NCJ 11 Jul 02 - CBB 2.2.15/R22, Use HadValue in DiffersFromSavedValue
' NCJ 12-15 Aug 02 - Revalidation and RFCs
' TA 14 Aug 02 -  New statuses
' NCJ 27 Aug 02 - Consider NR and CTC changes when revalidating lab test questions
' NCJ 4 Sept 02 - Added argument to RequiredStatusRFC and made sure all ValueStores are updated in ConfirmValue
' MLM 06/09/02: Added Public Property Get EFormInstance.
' NCJ 18 Sept 02 - Store Subject and eFormTaskId instead of eFormInstance
' NCJ 20 Sept 02 - Be sure to set eForm context for AREZZO when dealing with validations etc.
' RS 22/9/2002 - Added Timestamp parameter to InitTempResponse and ValidateResponse
' NCJ 2 Oct 02 - Use new IMedNow function instead of CDbl(Now)
' NCJ 18 Oct 02 - New ValueChanged property
' NCJ 29 Oct 02 - Blank optional questions should have status "Success" (brought forward from 2.2 Patch)
' NCJ 30 Oct 02 - Do not include Missing <-> N/A switches in ChangeCount
' NCJ 8 Jan 03 - Added User name to AREZZO for Authorisation questions
' RS/NCJ 21 Jan 03 - Accept data entry in user defined format for date/time
' NCJ 23 Jan 03 - Do Set AREZZO Status on Requested responses
' NCJ 28 Jan 03 - Add UserNameFull to ResponseState
' NCJ 14 Feb 03 - Added a few more things to Terminate
' NCJ 17 Feb 03 - SR 3921: Force non-empty "text" values to be Prolog atoms in AREZZO
' NCJ 25 Feb 03 - Added LocalDateFormat property (for Bug 861)
' NCJ 24 Mar 03 - Changed Revalidatable so we return lab test questions in general
' NCJ 1 Apr 03 - Refresh eForm and Visit dates as they're entered
' NCJ 10 Apr 03 - Also update eForm/Visit date when reapplying AREZZO after reload
' NCJ 29 Apr 03 - Added IgnoreMe property
' NCJ 7 May 03 - Added WillBeSaved property
'   Also make sure the IsChanged flag is correctly maintained (Bug 1565)
' NCJ 20 May 03 - We DO revalidate empty responses in case they've changed to/from optional (Bugs 1764, 1765)
' NCJ 11 Jun 03 - Handle removed Collect If conditions (Bug 844)
' NCJ 23 Jun 03 - Consider N/A QGroups when handling removed Collect If conditions (Bug 1866)
' NCJ 24 Mar 04 - Clear out object collections in Terminate
' NCJ 30 Jun 04 - Changed RequiresLabTestRFC to use moTempValue
' NCJ 30 Jun 04 - Include Warning/OKWarning questions in revalidation because they might have had validations removed
' NCJ 7 Jul 04 - Ensure we log the correct user name when changing Overrule Reason (Bug 2323)
' NCJ 7 Jul 04 - New routine EmptyResponseStatus
' NCJ 22 Jul 04 - Ensure we log the correct user name when changing Status (Bug 2350)
' ic 19/07/2005 added clinical coding
' NCJ 8 Dec 05 - Partial Date handling
' NCJ 18 Oct 06 - Bug 2825 - Must consider YM and MY in ReadLocalDate
' ic 21/11/2006 - bug 2843 - pass the current username in to setcodingstatus and setcodingdetails
' NCJ 5 Feb 07 - Issue 2874 - Corrected handling for regional reals in RevalidateValue
' NCJ 22 Feb 07 - Added DateAsDouble property (date value as a double)
' NCJ 19 Mar 07 - Issue 2872 - Changed ValidateLabValue according to HiddenLabWarnings value
' ic/ncj 28/02/2008 issue 2996 - fix for web inactive category values
' ic/ncj 28/02/2008 issue 2998 - fix for blank timestamp for requested web questions
' NCJ 11 March 2008 - Bug 3000 - Check special case of no RFC in Missing response with OCI
'----------------------------------------------------------------------------------

Option Explicit

' Max length of a question response
Private Const msRESPONSE_CHARS As Integer = 255

' The eForm instance it belongs to
'Private moEFI As EFormInstance
' NCJ 18 Sept 02 - Store TaskId instead of object
Private mlEFormTaskID As Long
Private moSubject As StudySubject

' The eForm Element it belongs to
Private moElement As eFormElementRO
' The State object
Private moResponseState As ResponseState

' NOTE: There are a few properties which are the ones we look at
' when deciding if anything's changed. Nothing else can change
' without one of these properties changing.
' Each of these "indicator" properties is stored in an individual ValueStore object
' which keeps track of the original (i.e. DB), previous and current values.

' The ValueStore objects for Value, Comments, RFC etc.
Private moVValue As ValueStore
Private moVValueCode As ValueStore
Private moVStatus As ValueStore
Private moVComments As ValueStore
Private moVOverrule As ValueStore
Private moVLabCode As ValueStore
' NCJ 9 Aug 02 - Also need ValidationMessage and OverruleReason
Private moVValMessage As ValueStore
' NCJ 27 Aug 02 - Also need lab results (for new RFCs)
Private moVNRStatus As ValueStore
Private moVCTCGrade As ValueStore

' Store the ValueStore objects as a collection
Private mcolValueStores As Collection

' The temporary store of a value after it's been validated
Private moTempValue As TempResponse

' The copy of Arezzo the belongs to our eForm Instance
Private moArezzo As Arezzo_DM

'Store its lab test validate object
Private moLabTestValidate As LabTestValidate

' NCJ 19 Mar 02 - For efficiency, store these as part of Response object
Private msArezzoTag As String
Private msArezzoDataName As String
Private msArezzoValue As String
Private mbKnownArezzoValue As Boolean

' NCJ 18 Sept 02 - Need to remember if Arezzo things changed
Private mbArezzoValueChanged As Boolean
Private mbArezzoStatusChanged As Boolean

' NCJ 21 Jan 03 - Local date format
Private msLocalDateFormat As String


'----------------------------------------------------
Public Property Get Status() As Integer
Attribute Status.VB_Description = "The Response status (automatically calculated by ValidateValue and ConfirmValue)"
'----------------------------------------------------
' The Response status
' (automatically calculated by ValidateValue and ConfirmValue)
'----------------------------------------------------

    Status = moVStatus.CurrentValue

End Property

'----------------------------------------------------
Public Sub SetStatusFromSchedule(ByVal nStatus As Integer)
'----------------------------------------------------
' Set the status of a response when setting an entire eform
' to either Unobtainable or Missing from the Schedule
' We generate an automatic RFC as required
'----------------------------------------------------
Dim sRFC As String

    sRFC = ""
    If RequiresStatusRFC(nStatus) Then
        sRFC = Me.EFormInstance.AutomaticRFC(SettingEFormStatus, StatusChanged)
    End If
    Call SetStatus(nStatus, sRFC)

End Sub

'----------------------------------------------------
Public Sub SetStatus(ByVal nStatus As Integer, ByVal sRFC As String, _
                    Optional sUserName As String = "", _
                    Optional sUserNameFull As String = "")
'----------------------------------------------------
' Switch the response status to/from Missing/Unobtainable/Not Applicable,
' with given RFC
' NCJ 28 Jan 03 - Added sUserNameFull
' NCJ 22 Jul 04 - Make sure user name always updated correctly (Bug 2350)
'----------------------------------------------------
    On Error GoTo ErrLabel

    If moVStatus.CurrentValue <> nStatus Then
        Select Case nStatus
        ' NCJ 23 Jan 03 - Include Requested because we can now set
        ' whole blank forms to Unobtainable
        Case eStatus.Missing, eStatus.Unobtainable, _
            eStatus.NotApplicable, eStatus.Requested
            Call moArezzo.SetQuestionStatus(moElement.Code, nStatus)
            mbArezzoStatusChanged = True
        End Select
        moVStatus.CurrentValue = nStatus
        moResponseState.Status = nStatus
            
        moResponseState.ReasonForChange = sRFC
        
        
        Call ResponseHasChanged       ' This sets current user name
        
        ' Override user name if necessary
        If sUserName > "" Then
            moResponseState.UserCode = sUserName
            If sUserNameFull > "" Then
                moResponseState.UserNameFull = sUserNameFull
            Else
                moResponseState.UserNameFull = sUserName
            End If
        End If
        
        ' Update Group instance status if necessary
        If Not (moElement.OwnerQGroup Is Nothing) Then
            Call OwnerQGI.UpdateStatus
        End If
    
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.SetStatus"

End Sub

'----------------------------------------------------
Friend Property Let Status(nStatus As Integer)
'----------------------------------------------------
' The Response status
' (automatically calculated by ValidateValue and ConfirmValue)
'----------------------------------------------------
' For Missing, Unobtainable and Not Applicable also set Arezzo status
' ic/ncj 28/02/2008 issue 2998 - fix for blank timestamp for requested web questions
'----------------------------------------------------
Dim bFromRequested As Boolean

    On Error GoTo ErrLabel

    If moVStatus.CurrentValue <> nStatus Then
        Select Case nStatus
        Case eStatus.Missing, eStatus.Unobtainable, eStatus.NotApplicable
            Call moArezzo.SetQuestionStatus(moElement.Code, nStatus)
            mbArezzoStatusChanged = True
        End Select
        moVStatus.CurrentValue = nStatus
        
        'ic/ncj 28/02/2008
        bFromRequested = (moResponseState.Status = eStatus.Requested) _
            And ((nStatus = eStatus.Missing) Or ((nStatus = eStatus.Success) And (Me.Value = "")))
        
        moResponseState.Status = nStatus
        
        'ic/ncj 2802/2008 issue 2998
        Call ResponseHasChanged(0, bFromRequested)

        ' Update Group instance status if necessary
        If Not (moElement.OwnerQGroup Is Nothing) Then
            Call OwnerQGI.UpdateStatus
        End If

    End If

Exit Property
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.Status (Let)"

End Property

'----------------------------------------------------
Public Property Get Value() As String
Attribute Value.VB_Description = "The Value of this response. NB See ValidateValue and ConfirmValue to set a value"
'----------------------------------------------------
' The Value of this response.
' If numeric, this is in "local" format.
' NB See ValidateValue and ConfirmValue to set a value
'----------------------------------------------------

    Value = moVValue.CurrentValue

End Property

'----------------------------------------------------
Public Property Get SavedValue() As String
Attribute SavedValue.VB_Description = "The ""saved"" Value of this response, i.e. the value that's currently stored in the database"
'----------------------------------------------------
' The "saved" Value of this response,
' i.e. the value that's currently stored in the database
'----------------------------------------------------

    SavedValue = moVValue.OriginalValue

End Property

'----------------------------------------------------
Friend Property Get PreviousValue() As String
'----------------------------------------------------
' The previous Value of this response.
'----------------------------------------------------

    PreviousValue = moVValue.PreviousValue

End Property

'----------------------------------------------------
Public Property Get ValueCode() As String
Attribute ValueCode.VB_Description = "The Value Code of this response (Category items only). NB See ValidateValue and ConfirmValue to set a Value Code"
'----------------------------------------------------
' The Value Code of this response
' (Category items only).
' NB See ValidateValue and ConfirmValue to set a Value Code
'----------------------------------------------------

    ValueCode = moVValueCode.CurrentValue

End Property

'----------------------------------------------------
Friend Property Get PreviousValueCode() As String
'----------------------------------------------------
' The previous Value Code of this response
' (Category items only).
'----------------------------------------------------

    PreviousValueCode = moVValueCode.PreviousValue

End Property

'----------------------------------------------------
Public Property Get SavedValueCode() As String
Attribute SavedValueCode.VB_Description = "The ""saved"" Value Code of this response. i.e. the value code that's currently stored in the database"
'----------------------------------------------------
' The "saved" Value Code of this response.
' i.e. the value code that's currently stored in the database
'----------------------------------------------------

    SavedValueCode = moVValueCode.OriginalValue

End Property

'----------------------------------------------------
Public Property Get SavedStatus() As Integer
Attribute SavedStatus.VB_Description = "The ""saved"" status of this response, i.e. the status that's currently stored in the database"
'----------------------------------------------------
' The "saved" status of this response,
' i.e. the status that's currently stored in the database
'----------------------------------------------------

    SavedStatus = moVStatus.OriginalValue

End Property

'----------------------------------------------------
Public Property Get Comments() As String
Attribute Comments.VB_Description = "The Comments for this response."
'----------------------------------------------------
' The Comments for this response.
'----------------------------------------------------

    Comments = moVComments.CurrentValue

End Property

'----------------------------------------------------
Public Property Let Comments(sComments As String)
'----------------------------------------------------
' The Comments for this response.
'----------------------------------------------------

    On Error GoTo ErrLabel

    If moVComments.CurrentValue <> sComments Then
        moVComments.CurrentValue = sComments
        ' Store in State as well
        moResponseState.Comments = sComments
        Call ResponseHasChanged
    End If
    
Exit Property
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.Comments (Let)"

End Property

'----------------------------------------------------
Public Property Get OverruleReason() As String
Attribute OverruleReason.VB_Description = "The Overrule Reason for this response. Setting this property for a response with Warning status automatically changes the status to OKWarning. Setting it to """" for a response with OKWarning status automatically changes the status to Warning."
'----------------------------------------------------
' The Overrule Reason for this response.
' Setting this property for a response with Warning status
' automatically changes the status to OKWarning.
' Setting it to "" for a response with OKWarning status
' automatically changes the status to Warning.
'----------------------------------------------------

    OverruleReason = moVOverrule.CurrentValue

End Property

'-----------------------------------------------------------------------------
Public Sub SetOverruleReason(ByVal sOverrule As String, Optional sRFC As String = "")
'-----------------------------------------------------------------------------
' NCJ 12 Aug 02 - This replaces the Overrule Property Let
' Change the Overrule reason for this response
' sRFC is the Reason for Change, only stored if non-empty
' If sOverrule > "" and response has Warning status
' automatically change the status to OKWarning.
' Setting it to "" for a response with OKWarning status
' automatically changes status to Warning.
' NCJ 7 Jul 04 - Ensure we log the correct user name here (Bug 2323)
'-----------------------------------------------------------------------------
Dim bStatusChanged As Boolean

    On Error GoTo ErrLabel

    bStatusChanged = False
    
    ' Only do something if it's changed
    If moVOverrule.CurrentValue <> sOverrule Then
        If sOverrule = "" And Me.Status = eStatus.OKWarning Then
            Me.Status = eStatus.Warning
            bStatusChanged = True
        End If
        If sOverrule > "" And Me.Status = eStatus.Warning Then
            Me.Status = eStatus.OKWarning
            bStatusChanged = True
        End If
        moVOverrule.CurrentValue = sOverrule
        ' Store in State as well
        moResponseState.OverruleReason = sOverrule
        
        ' Store the RFC if non-empty
        ' (this allows RFC for a value change to be preserved)
        If sRFC > "" Then
            moResponseState.ReasonForChange = sRFC
        End If
        
        Call ResponseHasChanged     ' NCJ 22 Jul 04 - This now sets user details
        
        ' Update Group instance status if status has changed
        If bStatusChanged And Not (moElement.OwnerQGroup Is Nothing) Then
            Call OwnerQGI.UpdateStatus
        End If
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.SetOverruleReason"

End Sub


'----------------------------------------------------
Public Property Get LabCode() As String
Attribute LabCode.VB_Description = "The Lab Code"
'----------------------------------------------------
' The Lab Code.
' (This can only be set through the eFormInstance)
'----------------------------------------------------

    LabCode = moVLabCode.CurrentValue

End Property

'----------------------------------------------------
Public Property Get LabTestDate() As Double
Attribute LabTestDate.VB_Description = "The Lab Test Date. This is automatically recalculated when question value is validated."
'----------------------------------------------------
' The Lab Test Date.
' This is automatically recalculated when question value is validated.
'----------------------------------------------------

    LabTestDate = moResponseState.LabTestDate

End Property

'----------------------------------------------------
Public Property Get ElementID() As Long
Attribute ElementID.VB_Description = "The eFormElement Id"
'----------------------------------------------------
' The eFormElement Id
'----------------------------------------------------

    ElementID = moResponseState.ElementID
    
End Property

'-------------------------------------
Public Property Get Element() As eFormElementRO
Attribute Element.VB_Description = "The eFormElement definition with which this Response is associated"
'-------------------------------------
' The eFormElement definition with which this Response is associated
'-------------------------------------

    Set Element = moElement

End Property

'-------------------------------------
Friend Property Get Key() As String
'-------------------------------------
' Get unique string key for this object
'-------------------------------------

    Key = moResponseState.Key

End Property

'----------------------------------------------------
Public Property Get ResponseId() As Long
Attribute ResponseId.VB_Description = "The Response TaskId, uniquely identifying this response within this Subject."
'----------------------------------------------------
' The Response TaskId, uniquely identifying this response
' within this Subject.
'----------------------------------------------------

    ResponseId = moResponseState.ResponseId

End Property

'----------------------------------------------------
Public Property Get RepeatNumber() As Integer
'----------------------------------------------------
' Repeat number of this response,
' 1 if not a repeating question
'----------------------------------------------------

    RepeatNumber = moResponseState.RepeatNumber

End Property

'----------------------------------------------------
Public Property Get TimeStamp() As Double
Attribute TimeStamp.VB_Description = "The timestamp of the response"
'----------------------------------------------------
' The timestamp of the response
'----------------------------------------------------

    TimeStamp = moResponseState.TimeStamp

End Property

'----------------------------------------------------
Public Property Get StatusString() As String
Attribute StatusString.VB_Description = "Get status as a string. See also the Status property."
'----------------------------------------------------
' Get status as a string.
' See also the Status property.
'----------------------------------------------------

    StatusString = GetStatusString(CInt(moVStatus.CurrentValue))
    
End Property

'----------------------------------------------------
Public Property Get LockStatus() As Integer
Attribute LockStatus.VB_Description = "The Lock status of the response. See also the LockStatusString property."
'----------------------------------------------------
' The Lock status of the response.
' See also the LockStatusString property.
'----------------------------------------------------

    LockStatus = moResponseState.LockStatus

End Property

'----------------------------------------------------
Public Property Get LockStatusString() As String
Attribute LockStatusString.VB_Description = "The Lock status as a string (for user display). See also the LockStatus property."
'----------------------------------------------------
' The Lock status as a string (for user display).
' See also the LockStatus property.
'----------------------------------------------------

    LockStatusString = GetLockStatusString(moResponseState.LockStatus)
    
End Property

'----------------------------------------------------
Public Property Get SoftwareVersion() As String
Attribute SoftwareVersion.VB_Description = "The Software version in use for this response"
'----------------------------------------------------
' The Software version in use for this response
'----------------------------------------------------

    SoftwareVersion = moResponseState.SoftwareVersion

End Property

'----------------------------------------------------
Public Property Get UserName() As String
Attribute UserName.VB_Description = "The User Name of the person who entered this response. This is automatically set when response data changes."
'----------------------------------------------------
' The User Name of the person who entered this response.
' This is automatically set when response data changes.
'----------------------------------------------------

    UserName = moResponseState.UserCode

End Property

'----------------------------------------------------
Public Property Get UserNameFull() As String
'----------------------------------------------------
' The full User Name of the person who entered this response.
' This is automatically set when response data changes.
'----------------------------------------------------

    UserNameFull = moResponseState.UserNameFull

End Property

'----------------------------------------------------
Public Property Get ReasonForChange() As String
Attribute ReasonForChange.VB_Description = "The user-entered RFC for this response"
'----------------------------------------------------
' The user-entered RFC for this response
'----------------------------------------------------

    ReasonForChange = moResponseState.ReasonForChange

End Property

'----------------------------------------------------
Public Property Get ValidationMessage() As String
Attribute ValidationMessage.VB_Description = "The evaluated Validation message for this response"
'----------------------------------------------------
' The evaluated Validation message for this response
'----------------------------------------------------

    ValidationMessage = moVValMessage.CurrentValue

End Property

'----------------------------------------------------
Public Property Get CTCGrade() As Integer
Attribute CTCGrade.VB_Description = "The calculated CTC Grade (only applicable to Lab Test questions)"
'----------------------------------------------------
' The calculated CTC Grade (only applicable to Lab Test questions)
'----------------------------------------------------
    
    CTCGrade = moVCTCGrade.CurrentValue
    
End Property

'----------------------------------------------------
Public Property Get NRStatus() As Integer
Attribute NRStatus.VB_Description = "The calculated Normal Range Lab status (single letter). Only applicable to Lab Test questions."
'----------------------------------------------------
' The calculated Normal Range Lab status (as integer).
' Only applicable to Lab Test questions.
'----------------------------------------------------
    
    NRStatus = moVNRStatus.CurrentValue
    
End Property

'----------------------------------------------------
Public Property Get NRStatusString() As String
'----------------------------------------------------
' The calculated Normal Range Lab status as a single letter,
' "L", "N" or "H"
' Only applicable to Lab Test questions.
'----------------------------------------------------
    
    NRStatusString = GetNRText(moResponseState.NRStatus)
    
End Property

'----------------------------------------------------
Public Property Get DateAsDouble() As Double
'----------------------------------------------------
' NCJ 22 Feb 07
' Retrieve a date response value as a VB date double
' Returns 0 if response not of type Date/Time
'----------------------------------------------------

    If Me.Element.DataType = eDataType.DateTime Then
        ' It's stored as the StandardValue property
        DateAsDouble = moResponseState.StandardValue
    Else
        DateAsDouble = 0
    End If
    
End Property

'----------------------------------------------------
Friend Sub SetState(oRState As ResponseState, _
                        oEFI As EFormInstance, _
                        oEFElement As eFormElementRO)
'----------------------------------------------------
' Set ourselves up with this state in this eForm Instance
' oEFElement is the eForm Element we're attached to
'----------------------------------------------------

    On Error GoTo ErrLabel

    Set moResponseState = oRState
    
'    Set moEFI = oEFI
    ' NCJ 18 Sept 02 - Remember the Subject and eFormTaskID
    mlEFormTaskID = oEFI.eFormTaskId
    Set moSubject = oEFI.VisitInstance.Subject
    
    Set moElement = oEFElement
    Set moArezzo = oEFI.Arezzo
    
    ' Set up our ValueStore objects
    Call GetStateValues
    
    ' NCJ 19 Mar 02 - Calculate Arezzo User Tag here
    ' User Tag is (CRFPageTaskId*10000 + CRFElementID)*10    (from MACRO 2.1)
    msArezzoTag = CStr(Me.ResponseId * 10)
    If Me.RepeatNumber > 1 Then
        msArezzoTag = msArezzoTag & "-" & Me.RepeatNumber
    End If
    ' Create the compound data name (Visit:eForm:Question)
    msArezzoDataName = oEFI.VisitInstance.Visit.Code & ":" & oEFI.eForm.Code & ":" & moElement.Code
    ' NCJ 19 Mar 02 - Initialise Arezzo value
    msArezzoValue = ""
    mbKnownArezzoValue = False
    
    ' Reset Arezzo changed flags
    mbArezzoValueChanged = False
    mbArezzoStatusChanged = False
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.SetState"

End Sub

'----------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------
' Initialise the object
'----------------------------------------------------

    ' Create the individual Value Stores
    ' We only need these for the "indicator" properties
    ' i.e. nothing else can change without one of these changing
    Set moVValue = New ValueStore
    Set moVValueCode = New ValueStore
    Set moVStatus = New ValueStore
    Set moVComments = New ValueStore
    Set moVOverrule = New ValueStore
    Set moVLabCode = New ValueStore
    Set moVValMessage = New ValueStore
    Set moVNRStatus = New ValueStore
    Set moVCTCGrade = New ValueStore
    
    ' Add them into a collection for ease of handling later
    Set mcolValueStores = New Collection
    mcolValueStores.Add moVValue
    mcolValueStores.Add moVValueCode
    mcolValueStores.Add moVStatus
    mcolValueStores.Add moVComments
    mcolValueStores.Add moVOverrule
    mcolValueStores.Add moVLabCode
    mcolValueStores.Add moVValMessage
    mcolValueStores.Add moVNRStatus
    mcolValueStores.Add moVCTCGrade
    
    ' Lab test validation
    Set moLabTestValidate = New LabTestValidate
    
End Sub

'----------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------

    Call Terminate

End Sub

'----------------------------------------------------
Friend Sub Terminate()
'----------------------------------------------------
' Tidy up before we're destroyed
'----------------------------------------------------

    On Error GoTo ErrLabel

    Set moResponseState = Nothing
    Set moVValue = Nothing
    Set moVValueCode = Nothing
    Set moVStatus = Nothing
    Set moVComments = Nothing
    Set moVOverrule = Nothing
    Set moVLabCode = Nothing
    ' NCJ 14 Feb 03 - Belatedly added some more...
    Set moVValMessage = Nothing
    Set moVNRStatus = Nothing
    Set moVCTCGrade = Nothing
    
    Set moArezzo = Nothing
    Set moElement = Nothing
    Set moSubject = Nothing
    Set moTempValue = Nothing
    
    Set moLabTestValidate = Nothing
    
    Call CollectionRemoveAll(mcolValueStores)   ' NCJ 24 Mar 04
    Set mcolValueStores = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.Terminate"

End Sub

'----------------------------------------------------
Private Sub GetStateValues()
'----------------------------------------------------
' Set up the ValueStore objects from the State object
' ic 22/07/2005 added clinical coding
'----------------------------------------------------

    On Error GoTo ErrLabel

    ' Convert the value to "local" format
    moVValue.OriginalValue = moElement.LocaliseValue(moResponseState.Value)
    
    ' NCJ 21 Jan 03 - Special handling for dates/times
    ' NCJ 18 Dec 06 - Beware of moArezzo being Nothing
    If moElement.DataType = eDataType.DateTime And Not moArezzo Is Nothing Then
        ' Get the local formatting, if any
        msLocalDateFormat = moSubject.LocalDateTypeFormat(moElement.DateTimeType)
        If msLocalDateFormat > "" Then
            moVValue.OriginalValue = moArezzo.FormatDate(Me.ArezzoValue, msLocalDateFormat)
        End If
    End If
    
    moVValueCode.OriginalValue = moResponseState.ValueCode
    moVStatus.OriginalValue = moResponseState.Status
    moVComments.OriginalValue = moResponseState.Comments
    moVOverrule.OriginalValue = moResponseState.OverruleReason
    moVLabCode.OriginalValue = moResponseState.LabCode
    ' NCJ 13 Aug 02 - Added Validation Message
    moVValMessage.OriginalValue = moResponseState.ValidationMessage
    ' NCJ 27 Aug 02 - Added lab results
    moVNRStatus.OriginalValue = moResponseState.NRStatus
    moVCTCGrade.OriginalValue = moResponseState.CTCGrade

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.GetStateValues"

End Sub

'----------------------------------------------------
Public Property Get WillBeSaved() As Boolean
'----------------------------------------------------
' NCJ 7 May 03 - Returns TRUE if this response will be saved when the eForm is saved
' i.e. if it's not in the last blank row of an RQG
'----------------------------------------------------

    ' Default to WILL be saved
    WillBeSaved = True
    
    ' Is it a group item?
    If moElement.OwnerQGroup Is Nothing Then Exit Property
    
    ' Is it in the last row?
    If Me.RepeatNumber = OwnerQGI.Rows Then
        WillBeSaved = OwnerQGI.LastRowWillBeSaved
    End If
    
End Property

'----------------------------------------------------
Friend Sub IgnoreMe()
'----------------------------------------------------
' Calling IgnoreMe means the response will NOT be saved
' when the other responses are saved
'----------------------------------------------------
    
    moResponseState.IgnoreMe = True
    ' Also remove the value from AREZZO
    Call moArezzo.ReplaceQData(msArezzoDataName, "", msArezzoTag, "")

End Sub

'----------------------------------------------------
Friend Sub ChangeStatetoSaved()
'----------------------------------------------------
' Set response to not new and not changed
' NCJ 29 Apr 03 - Look at IgnoreMe
'----------------------------------------------------
Dim oVStore As ValueStore

    On Error GoTo ErrLabel

    If moResponseState.IgnoreMe = True Then
        ' It wasn't saved, but reset IgnoreMe flag here
        ' and re-add the value to AREZZO
        moResponseState.IgnoreMe = False
        Call moArezzo.ReplaceQData(msArezzoDataName, msArezzoValue, msArezzoTag, "")
    Else
        moResponseState.IsChanged = False
        moResponseState.IsNew = False
        
        ' NCJ 30/7/01 - Flag each as "Saved"
        For Each oVStore In mcolValueStores
            Call oVStore.Saved
        Next
    
        ' Reset Arezzo changed flags
        mbArezzoValueChanged = False
        mbArezzoStatusChanged = False
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.ChangeStateToSaved"

End Sub

'----------------------------------------------------
Public Property Get Changed() As Boolean
Attribute Changed.VB_Description = "TRUE if this Response has changed since it was last saved. NB This is FALSE for New but unchanged responses"
'----------------------------------------------------
' TRUE if this Response has changed since it was last saved.
' NB This is FALSE for New but unchanged responses
'----------------------------------------------------

    Changed = moResponseState.IsChanged

End Property

'----------------------------------------------------
Public Property Get ValueChanged() As Boolean
'----------------------------------------------------
' NCJ 18 Oct 02
' Whether this response Value has changed from the stored VB value
' i.e. ignoring changes to status, validation message, lab status etc.
'----------------------------------------------------

    ValueChanged = moVValue.Dirty Or moVValueCode.Dirty
    
End Property

'----------------------------------------------------
Private Function Dirty() As Boolean
'----------------------------------------------------
' TRUE if Response is new or if anything has changed in this response
' since it was last saved.
' Walk through the ValueStore objects to see if any are dirty
' NB Nothing else can change without one of these changing
'----------------------------------------------------
Dim oVStore As ValueStore

    On Error GoTo ErrLabel

    If moResponseState.IsNew Then
        Dirty = True
    Else
        Dirty = False
        For Each oVStore In mcolValueStores
            If oVStore.Dirty Then
                Dirty = True
                ' Stop when we find a changed one
                Exit For
            End If
        Next
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.Dirty"

End Function

'----------------------------------------------------
Public Sub RejectValue()
Attribute RejectValue.VB_Description = "Reject a value that's been validated and temporarily stored."
'----------------------------------------------------
' Reject a value that's been validated and temporarily stored.
'----------------------------------------------------
' Just throw away our temporary copy
'----------------------------------------------------
    
    Set moTempValue = Nothing

End Sub

'----------------------------------------------------
Public Sub ConfirmValue(ByVal sOverrule As String, _
                    ByVal sRFC As String, _
                    ByVal sUserName As String, _
                    Optional sUserNameFull As String)
'----------------------------------------------------
' Confirm the temporary value created in ValidateValue.
' sRFC is the Reason For Change (if any).
' sUserName is the user ID of an authoriser (if any).
' If sOverrule is non-empty, change Warning status to OKWarning.
'----------------------------------------------------
' NCJ 11 Jan 02 - Store "standardised" numbers in moResponseState.Value
' (but still store "localised" numbers in moVValue)
' NCJ 19 Mar 02 - Use previously calculated Arezzo UserTag and DataName
' NCJ 18 Sept 02 - Created separate routine to add lab test results to Arezzo
' NCJ 8 Jan 03 - Added extra sUserNameFull argument and added it to AREZZO
' NCJ/RS 21 Jan 03 - Make sure we correctly deal with locally-formatted dates
' NCJ 1 Apr 03 - Update the visit & eForm dates as necessary here
'----------------------------------------------------
Dim oEFI As EFormInstance
Dim sNewRFC As String

    On Error GoTo ErrLabel

    ' Do nothing if there isn't one
    If moTempValue Is Nothing Then Exit Sub
    
    Set oEFI = Me.EFormInstance
    
    If sOverrule > "" Then
        If moTempValue.ResponseStatus = eStatus.Warning Then
            moTempValue.ResponseStatus = eStatus.OKWarning
        End If
    End If
    
    ' NCJ 11 March 2008 - Bug 3000 - Check special case of no RFC
    sNewRFC = sRFC
    If sRFC = "" _
            And moTempValue.ResponseStatus = eStatus.Missing _
            And moElement.CollectIfCond <> "" _
            And RequiresValueRFC Then
        sNewRFC = "*** Response value deleted from field with Collect If condition"
    End If
    ' MLM 29/05/2008: Issue 3038
    If sNewRFC = "" And RequiresValueRFC Then
        sNewRFC = "*** Response value changed without a corresponding reason for change recorded. Please verify."
    End If
    
'    moResponseState.ReasonForChange = sRFC
    moResponseState.ReasonForChange = sNewRFC
    
    ' Store the new stuff in the local value stores and State object
    ' NCJ 11 Jan 02 - For numbers, store "standardised" version (as in ArezzoValue) in moResponseState
    '   but store "localised" version in moVValue
    ' NCJ/RS 21 Jan 03 - For dates, store "standardised" version (as in Value) in moResponseState
    '   but store "localised" version in moVValue
    moVValue.CurrentValue = moTempValue.Value
    Select Case moElement.DataType
    Case eDataType.LabTest, eDataType.Real, eDataType.IntegerNumber
        moResponseState.Value = moTempValue.ArezzoValue
    Case eDataType.DateTime
        ' Make sure that the user formatted value is displayed
        moVValue.CurrentValue = moTempValue.DisplayValue
        moResponseState.Value = moTempValue.Value
        ' NCJ 21 Jan 03 - Store the date/time as a double too
        moResponseState.StandardValue = moArezzo.ArezzoDateToDouble(moTempValue.ArezzoValue)
    Case Else
        ' Non-numbers
        moResponseState.Value = moTempValue.Value
    End Select
    
    moVValueCode.CurrentValue = moTempValue.ValueCode
    moResponseState.ValueCode = moTempValue.ValueCode
    
    moVStatus.CurrentValue = moTempValue.ResponseStatus
    moResponseState.Status = moTempValue.ResponseStatus
    
    moVOverrule.CurrentValue = sOverrule
    moResponseState.OverruleReason = sOverrule
    
    moVValMessage.CurrentValue = moTempValue.ValidationMessage
    moResponseState.ValidationMessage = moTempValue.ValidationMessage
    
    ' NCJ 18 Feb 02 - Only do this for LabTest questions
    If moElement.DataType = eDataType.LabTest Then
        ' Set the lab code from the eForm instance
        moVLabCode.CurrentValue = oEFI.LabCode
        moResponseState.LabCode = oEFI.LabCode
        
        moVCTCGrade.CurrentValue = moTempValue.CTCGrade
        moResponseState.CTCGrade = moTempValue.CTCGrade
        
        moVNRStatus.CurrentValue = moTempValue.NRStatus
        moResponseState.NRStatus = moTempValue.NRStatus
        
        moResponseState.LabTestDate = moTempValue.LabTestDate
    End If
    
    ' NCJ 7 May 03 - Let ResponseHasChanged handle all the "changes"
    Call ResponseHasChanged(moTempValue.TimeStamp)
    
    ' NCJ 28 Jan 03 - Also add UserNameFull if we can
    If sUserName > "" Then
        moResponseState.UserCode = sUserName
        If sUserNameFull > "" Then
            moResponseState.UserNameFull = sUserNameFull
        Else
            moResponseState.UserNameFull = sUserName
        End If
    End If
    
    ' Add value to Arezzo (replacing any previous value for this User Tag)
        
    ' NCJ 19 Mar 02 - Only replace value if we know it's different from the last one we did
    If mbKnownArezzoValue And msArezzoValue = moTempValue.ArezzoValue Then
        ' No need to do anything
    Else
        ' Save this value and send to Arezzo
        msArezzoValue = moTempValue.ArezzoValue
        mbKnownArezzoValue = True
        ' NCJ 19 Mar 02 - Use previously calculated UserTag and DataName
        Call moArezzo.ReplaceQData(msArezzoDataName, msArezzoValue, msArezzoTag, "")
        
        ' NCJ 8 Jan 03 - Add Full User name (code) if Authorisation question
        If moElement.Authorisation > "" Then
            Call moArezzo.ReplaceQData(msArezzoDataName & ":username", _
                                        sUserNameFull, msArezzoTag, "user")
        End If
        
        mbArezzoValueChanged = True
        ' non-empty value change overrides status change
        If msArezzoValue > "" Then mbArezzoStatusChanged = False
        
        ' NCJ 1 Apr 03 (Bug 1473) If it's a form date, ensure that we update form:date and visit:date as appropriate
        ' NCJ 10 Apr 03 - Now in separate routine
        Call AddAsEFormAndVisitDate
         
    End If

    ' NCJ 12 Oct 01 - Need to also add nrstatus and ctcgrade for Lab Test questions
    ' NCJ 18 Sept 02 - Created separate routine for it
    Call AddLabTestResultToArezzo
    
    Set moTempValue = Nothing
    Set oEFI = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.ConfirmValue"

End Sub

'----------------------------------------------------
Private Sub AddAsEFormAndVisitDate()
'----------------------------------------------------
' NCJ 10 Apr 03
' Add this response as the eForm and Visit date
' if it's an eForm/Visit date question
' Does nothing if not an eForm date question
'----------------------------------------------------

    If moElement.ElementUse = eElementUse.EFormVisitDate Then
        ' "StandardValue" is the date as a double
        Call Me.EFormInstance.RefreshEFormDate(moResponseState.StandardValue, msArezzoValue)
        ' Are we the visit eForm?
        If Not Me.EFormInstance.VisitInstance.VisitEFormInstance Is Nothing Then
            If mlEFormTaskID = Me.EFormInstance.VisitInstance.VisitEFormInstance.eFormTaskId Then
                Call Me.EFormInstance.VisitInstance.RefreshVisitDate(moResponseState.StandardValue, msArezzoValue)
            End If
        End If
    End If
    
End Sub

'----------------------------------------------------
Private Sub AddLabTestResultToArezzo()
'----------------------------------------------------
' Add NR/CTCGrade to AREZZO
'----------------------------------------------------
Dim sNRCTC As String

    If moElement.DataType = eDataType.LabTest Then
        ' Store empty for no CTC grade
        If moResponseState.CTCGrade = eCTCGrade.ctcNone Then
            sNRCTC = ""
        Else
            sNRCTC = moResponseState.CTCGrade
        End If
        Call moArezzo.ReplaceQData(msArezzoDataName & ":ctcgrade", sNRCTC, msArezzoTag, "ctc")
        ' Convert the integer NRStatus to L, N or H
        sNRCTC = GetNRText(moResponseState.NRStatus)
        Call moArezzo.ReplaceQData(msArezzoDataName & ":nrstatus", sNRCTC, msArezzoTag, "nr")
        mbArezzoValueChanged = True
        mbArezzoStatusChanged = False
    End If

End Sub

'----------------------------------------------------
Private Function GetLabTestDate() As Double
'----------------------------------------------------
' Recalculate the lab test date for this response
' (only relevant for Lab Test questions)
' Returns 0 if no date available
'----------------------------------------------------
Dim sExpr As String
Dim dblDate As Double

    On Error GoTo ErrLabel

    dblDate = 0
    
    sExpr = moElement.ClinicalTestDateExpr
    If sExpr > "" Then
        ' Set the question context
        Call moArezzo.SetQuestionContext(moElement.Code, Me.RepeatNumber, Me.EFormInstance)
        ' Evaluate the date expression
        dblDate = moArezzo.EvaluateDate(sExpr)
    End If
    
    GetLabTestDate = dblDate

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.GetLabTestDate"

End Function

'----------------------------------------------------
Public Function ValidateValue(ByVal sValue As String, _
                        ByRef sErrMessage As String, _
                        ByRef bChanged As Boolean, _
                        Optional dblTimestamp As Double = 0) As Long
'----------------------------------------------------
' Validate the value of the response
' Does NOT store the value but keeps a temporary copy
' NB For a category item sValue may be the category CODE or VALUE
' (but it's much more efficient and less ambiguous if it's the CODE)
' Returns Long value as in eStatus enumeration
' Fills in sErrMessage if there's an error or warning message to display
' NCJ 31/8/00 - bChanged will become TRUE if this represents a change in data value
'           i.e. if the processed formatted value has changed
' NCJ 11/9/01 - Use new Validation object instead of ValidationState
' NCJ 11/10/01 - Set bChanged if LabCode has changed
' NCJ 11/10/01 - Do not accept any response longer than 255 chars
' NCJ 8 Jan 02 - Do not check QuestionLength for numbers before formatting them
' NCJ 19 Mar 02 - Don't add unchanged values to Arezzo
' MLM 21/05/02: Also reject " and |
' MLM 22/05/02: Calculate sArezzoValue differently for numbers.
' NCJ 5 Jun 02 - CBB 2.2.10/1, Ensure category codes stored in correct case
' REM 21/06/02 - CBB 2.2.16 No. 12 Also reject ~ character
' NCJ 12 Aug 02 - Restructuring (there are now calls to more smaller routines)
' RS  20 Jan 03 - Added support for data entry in local date format
' ic 27/10/2005 added clinical coding
' ic/ncj 28/02/2008 issue 2996 - fix for web inactive category values
'----------------------------------------------------
Dim sDisplayText As String  ' Correctly formatted value
Dim lResult As Long
Dim sArezzoValue As String
Dim bOKSoFar As Boolean     ' How are we doing?
Dim sTrimValue As String
Dim sLocalDateValue As String ' Used for local format
Dim oEFI As EFormInstance

    On Error GoTo ErrLabel

    lResult = eResponseValidation.ValueOK
    bOKSoFar = False
    sDisplayText = ""
    sErrMessage = ""
    bChanged = False
    sLocalDateValue = ""
    
    ' Initialise temporary response storage
    ' RS 22/9/2002: Added Timestamp parameter (for WEB cient use)
    Call InitTempResponse(dblTimestamp)
    
    ' NCJ 11/10/01 - Remove leading and trailing spaces
    sTrimValue = Trim(sValue)
    
    Set oEFI = Me.EFormInstance
    
    ' Check first for an empty value
    If sTrimValue = "" Then
        bChanged = (Me.Value <> "")
        ' NCJ 7 Jul 04 - Use EmptyResponseStatus
        moTempValue.ResponseStatus = EmptyResponseStatus
        ValidateValue = moTempValue.ResponseStatus  ' NCJ 11 Feb 03 - Changed from eStatus.Missing
        
        'ic/ncj 28/02/2008 issue 2998 - blank timestamp for requested web questions
        If (dblTimestamp > 0) And (Me.Status = eStatus.Requested) Then
            moResponseState.TimeStamp = dblTimestamp
        End If
        
        Exit Function
    End If
    
    ' NCJ 15/10/01 - Make sure response contains no backwards quotes
    ' MLM 21/05/02: Also reject " and |
    ' REM 21/06/02 - Also reject ~
    If InStr(sTrimValue, "`") > 0 Or InStr(sTrimValue, """") > 0 Or InStr(sTrimValue, "|") > 0 Or InStr(sTrimValue, "~") Then
        bChanged = True     ' It must have changed!
        sErrMessage = "Question responses may not contain double or backwards quotes or the | or ~ characters."
        ValidateValue = eStatus.InvalidData
        Exit Function
    End If
    
    ' NCJ 11/10/01 - Check length of response data
    ' (generally only happens for derived questions)
    If Len(sTrimValue) > msRESPONSE_CHARS Then
        bChanged = True     ' It must have changed!
        sErrMessage = "Question responses may not be longer than " & msRESPONSE_CHARS & " characters."
        ValidateValue = eStatus.InvalidData
        Exit Function
    End If
    
    ' NCJ 8 Jan 02 - Only check question length for text
    If moElement.DataType = eDataType.Text And moElement.QuestionLength > 0 And Len(sTrimValue) > moElement.QuestionLength Then
        bChanged = True     ' It must have changed!
        sErrMessage = "This is longer than the allowed length of " & moElement.QuestionLength & " for this question."
        ValidateValue = eStatus.InvalidData
        Exit Function
    End If
    
    ' Use the DataFormatter class to do initial validation
    Select Case moElement.DataType
    Case eDataType.IntegerNumber, eDataType.Real, eDataType.LabTest
        'MLM 22/05/02: ValidNumber now responsible for sArezzoValue
        lResult = moElement.Formatter.ValidNumber(sTrimValue, sDisplayText, sArezzoValue)
'        If sDisplayText <> "" Then
'            sArezzoValue = LocalNumToStandard(sDisplayText)
'        End If

    Case eDataType.Text
        lResult = moElement.Formatter.ValidText(sTrimValue, sDisplayText)
'        sArezzoValue = sDisplayText
        ' NCJ 17 Feb 03 - Ensure we prevent AREZZO munging our text values
        sArezzoValue = TextifyArezzoValue(sDisplayText)
    
        'ic 27/10/2005 added clinical coding
    Case eDataType.Thesaurus
        lResult = moElement.Formatter.ValidText(sTrimValue, sDisplayText)
        sArezzoValue = TextifyArezzoValue(sDisplayText)
    
    Case eDataType.DateTime
        ' RS 20/01/2003: Add support for data entry in local format
        ' If selected, convert the date/time into database format before continuing
        sDisplayText = ReadLocalDate(sTrimValue, sArezzoValue, sLocalDateValue)

'        sDisplayText = moArezzo.ReadValidDate(sTrimValue, moElement.Format, sArezzoValue)
        If sDisplayText = "" Then
            lResult = eResponseValidation.NotADateTime
        End If

    Case eDataType.Category
        ' See if it's a code or a value
        ' Set sDisplayText to be the Category Value
        ' Set sArezzoValue to be Category Code
        sDisplayText = moElement.CategoryValue(sTrimValue)  ' Get Value from Code
        If sDisplayText = "" Then
            ' sTrimValue wasn't a category code - see if it represents a value
            sArezzoValue = moElement.CategoryCode(sTrimValue)   ' Get Code from Value
            If sArezzoValue = "" Then
                'ic/ncj 28/02/2008 issue 2996 - fix for web inactive category values
                'if proposed value is the same as current value, ignore it
                If (sTrimValue = Me.ArezzoValue) Then
                    sDisplayText = Me.Value
                Else
                    ' It's not a valid category value!
                    lResult = eResponseValidation.ArezzoRangeError
                End If
            Else
                ' sTrimValue is a category Value and sArezzo is its Code
                ' NCJ 5 Jun 02 - Get correct Value from Code (CBB 2.2.10/1)
                sDisplayText = moElement.CategoryValue(sArezzoValue)  ' Get Value from Code
            End If
        Else
            ' sTrimValue is the Code (and sDisplayText is now the Value)
            ' NCJ 5 Jun 02 - Get the code from the value to ensure correct case (CBB 2.2.10/1)
            sArezzoValue = moElement.CategoryCode(sDisplayText)
        End If
        
    Case eDataType.MultiMedia
        ' NCJ 21/9/01 - Deal with Multimedia
        ' Assume sTrimValue is a valid file name
        lResult = eStatus.Success
        sDisplayText = sTrimValue
        sArezzoValue = sTrimValue
    End Select
    
    ' If the data was duff, fill in the error message
    sErrMessage = GetResponseErrorString(lResult)
    ' No error message if it was OK
    bOKSoFar = (sErrMessage = "")
    
    ' SPECIAL TREATMENT for LABTEST questions
    ' We do the LabTestValidation
    If bOKSoFar And moElement.DataType = eDataType.LabTest Then
        bOKSoFar = ValidateLabValue(sDisplayText, sErrMessage)
    End If
    
    ' Say whether it represents a change
    ' (changed value or changed Labcode)
    ' NCJ 26 Jun 03 - Careful with locally formatted date fields...
'    bChanged = (sDisplayText <> Me.Value) Or _
'            ((moElement.DataType = eDataType.LabTest) And (oEFI.LabCode <> moVLabCode.CurrentValue))
    Select Case moElement.DataType
    Case eDataType.LabTest
        bChanged = (sDisplayText <> Me.Value) Or (oEFI.LabCode <> moVLabCode.CurrentValue)
    Case eDataType.DateTime
        ' Compare with DB value (not localised value)
        bChanged = (sDisplayText <> moResponseState.Value)
    Case Else
        bChanged = (sDisplayText <> Me.Value)
    End Select
    
    ' This is for NON-LABTEST questions
    ' If it was OK and not a LabTest question, try adding it to Arezzo
    If bOKSoFar And (moElement.DataType <> eDataType.LabTest) Then
        If bChanged Then
            ' Let Arezzo do its validation
            bOKSoFar = ArezzoValidate(sArezzoValue, sErrMessage)
        Else
            ' It hasn't changed so just stay with what we've got
            moTempValue.ValidationMessage = Me.ValidationMessage
            moTempValue.ResponseStatus = Me.Status
            msArezzoValue = sArezzoValue
            mbKnownArezzoValue = True
        End If
    End If
    
    If bOKSoFar Then
        ' Fill in the rest of our temporary copy
        moTempValue.DisplayValue = sLocalDateValue    ' Only relevant for date/time questions
        moTempValue.Value = sDisplayText
        moTempValue.ArezzoValue = sArezzoValue
        If moElement.DataType = eDataType.Category Then
            ' Store the code
            moTempValue.ValueCode = sArezzoValue
        Else
            moTempValue.ValueCode = ""
        End If
    End If
    
    ' Return the Response Status of our temporary store
    ValidateValue = moTempValue.ResponseStatus
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.ValidateValue(" & sValue & ")"

End Function

'----------------------------------------------------------------------------------
Public Property Get LocalDateFormat() As String
'----------------------------------------------------------------------------------
' NCJ 25 Feb 03 - Make the local date format available
' so that we can sort out derived dates in EFormInstance (Bug 861)
' NCJ 6 May 03 - Made public so we can correctly display dates from calendar control (Bug 1673)
'----------------------------------------------------------------------------------
    
    If msLocalDateFormat > "" Then
        LocalDateFormat = msLocalDateFormat
    Else
        LocalDateFormat = moElement.Format
    End If
    
End Property

'----------------------------------------------------------------------------------
Public Sub SetCodingDetails(sDictionaryName As String, sDictionaryVersion As String, sCodingDetails As String, sRFC As String, _
    sUserName As String, sUserNameFull As String)
'----------------------------------------------------------------------------------
' ic 09/09/2005
' sub updates coding details of a clinical coding response
' revisions
' ic 21/11/2006 bug 2843 - pass the current username in
'----------------------------------------------------------------------------------
    Dim oCodedTermHistory As MACROCCBS30.CodedTermHistory
    
    'get the CodedTermHistory object
    Set oCodedTermHistory = Me.EFormInstance.CodedTermHistoryFromTaskId(Me.ResponseId, Me.RepeatNumber)
    
    If (Not (oCodedTermHistory Is Nothing)) Then
        'set the coding details
        Call oCodedTermHistory.SetCode(sDictionaryName, sDictionaryVersion, sCodingDetails, sUserName, sUserNameFull, _
        Me.Value, CDbl(Me.TimeStamp), CInt(moSubject.Timezone.TimezoneOffset), sRFC, False)
    End If
End Sub

'----------------------------------------------------------------------------------
Public Sub SetCodingStatus(nStatus As Integer, sUserName As String, sUserNameFull As String)
'----------------------------------------------------------------------------------
' ic 09/09/2005
' sub updates coding status of a clinical coding response
' revisions
' ic 21/11/2006 bug 2843 - pass the current username in
'----------------------------------------------------------------------------------
Dim oCodedTermHistory As MACROCCBS30.CodedTermHistory
    
    'get the CodedTermHistory object
    Set oCodedTermHistory = Me.EFormInstance.CodedTermHistoryFromTaskId(Me.ResponseId, Me.RepeatNumber)
    
    If (Not (oCodedTermHistory Is Nothing)) Then
        'set the coding status
        Call oCodedTermHistory.SetStatus(CInt(nStatus), sUserName, sUserNameFull, Me.Value, CDbl(Me.TimeStamp), _
            CInt(moSubject.Timezone.TimezoneOffset))
    End If
End Sub

'----------------------------------------------------------------------------------
Public Property Get RequiresCodingRFC() As Boolean
'----------------------------------------------------------------------------------
' ic 09/09/2005
' property returns true if the coded response requires a rfc
'----------------------------------------------------------------------------------
Dim oCodedTermHistory As MACROCCBS30.CodedTermHistory
    
    'get the CodedTermHistory object
    Set oCodedTermHistory = Me.EFormInstance.CodedTermHistoryFromTaskId(Me.ResponseId, Me.RepeatNumber)
    If (Not (oCodedTermHistory Is Nothing)) Then
        RequiresCodingRFC = oCodedTermHistory.RequiresRFC
    Else
        RequiresCodingRFC = False
    End If
End Property

'----------------------------------------------------------------------------------
Public Property Get DictionaryName() As String
'----------------------------------------------------------------------------------
' ic 15/07/2005
' property returns the current dictionary name
'----------------------------------------------------------------------------------
Dim oCodedTermHistory As MACROCCBS30.CodedTermHistory

    Set oCodedTermHistory = Me.EFormInstance.CodedTermHistoryFromTaskId(Me.ResponseId, Me.RepeatNumber)
    If (Not (oCodedTermHistory Is Nothing)) Then
        DictionaryName = oCodedTermHistory.DictionaryName
    Else
        DictionaryName = ""
    End If
'    DictionaryName = moResponseState.DictionaryName
End Property

'----------------------------------------------------------------------------------
Public Property Get DictionaryVersion() As String
'----------------------------------------------------------------------------------
' ic 15/07/2005
' property returns the current dictionary version
'----------------------------------------------------------------------------------
Dim oCodedTermHistory As MACROCCBS30.CodedTermHistory

    Set oCodedTermHistory = Me.EFormInstance.CodedTermHistoryFromTaskId(Me.ResponseId, Me.RepeatNumber)
    If (Not (oCodedTermHistory Is Nothing)) Then
        DictionaryVersion = oCodedTermHistory.DictionaryVersion
    Else
        DictionaryVersion = ""
    End If
'    DictionaryVersion = moResponseState.DictionaryVersion
End Property

'----------------------------------------------------------------------------------
Public Property Get CodingStatus() As Integer
'----------------------------------------------------------------------------------
' ic 15/07/2005
' property returns the current coding status
'----------------------------------------------------------------------------------
Dim oCodedTermHistory As MACROCCBS30.CodedTermHistory

    Set oCodedTermHistory = Me.EFormInstance.CodedTermHistoryFromTaskId(Me.ResponseId, Me.RepeatNumber)
    If (Not (oCodedTermHistory Is Nothing)) Then
        CodingStatus = oCodedTermHistory.CodingStatus
    Else
        CodingStatus = eCodingStatus.csEmpty
    End If
'    CodingStatus = moResponseState.CodingStatus
End Property

'----------------------------------------------------------------------------------
Public Property Get CodingDetails() As String
'----------------------------------------------------------------------------------
' ic 15/07/2005
' property returns the current coding details
'----------------------------------------------------------------------------------
Dim oCodedTermHistory As MACROCCBS30.CodedTermHistory

    Set oCodedTermHistory = Me.EFormInstance.CodedTermHistoryFromTaskId(Me.ResponseId, Me.RepeatNumber)
    If (Not (oCodedTermHistory Is Nothing)) Then
        CodingDetails = oCodedTermHistory.CodingDetails
    Else
        CodingDetails = ""
    End If
'    CodingDetails = moResponseState.CodingDetails
End Property

'----------------------------------------------------------------------------------
Private Function ReadLocalDate(sInputDate As String, _
                        ByRef sArezzoValue As String, _
                        ByRef sLocalDate As String) As String
'----------------------------------------------------------------------------------
' NCJ 21 Jan 03
' For the given sInputDate, returns the equivalent date string formatted
' according to the question's format, with the correct "local" version in sLocalDate
' Also returns the AREZZO version of the date
' Returns empty string if sInputDate not valid
' NCJ 8 Dec 05 - Consider partial dates
' NCJ 18 Oct 06 - Bug 2825 - Make sure both YM and MY are considered!
'----------------------------------------------------------------------------------
Dim sReadDate As String

    On Error GoTo ErrLabel
    
    sReadDate = ""
    
    If msLocalDateFormat > "" Then
        ' Read it in using local format
        sLocalDate = moSubject.Arezzo.ReadValidDate(sInputDate, msLocalDateFormat, sArezzoValue)
        ' If successful, reformat it according to question's defined format
        If sLocalDate > "" Then
            sReadDate = moSubject.Arezzo.FormatDate(sArezzoValue, moElement.Format)
        End If
    Else
        ' No local settings - use Element's own format
        sLocalDate = moSubject.Arezzo.ReadValidDate(sInputDate, moElement.Format, sArezzoValue)
        sReadDate = sLocalDate
    End If

    If sReadDate > "" Then
        ' NCJ 8 Dec 05 - Extra check if Partial Dates not allowed
        If moElement.TextCase = 0 Then
            Select Case moArezzo.ALM.PartialDateType(sArezzoValue)
            Case eArezzoDateType.adtMonthYear
                ' Format must be month/year too
                ' NCJ 18 Oct 06 - Bug 2825 - Must consider YM too
                If Not (moElement.DateTimeType = dttMY Or moElement.DateTimeType = dttYM) Then
                    sReadDate = ""
                End If
            Case eArezzoDateType.adtYearOnly
                ' Format must be year-only too
                If Not moElement.DateTimeType = dttY Then
                    sReadDate = ""
                End If
            End Select
        End If
    End If
    
    ReadLocalDate = sReadDate
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.ReadLocalDate"
    
End Function

'----------------------------------------------------
Private Function ArezzoValidate(ByVal sArezzoValue As String, _
                                ByRef sErrMessage As String) As Boolean
'----------------------------------------------------
' Validate the given value in Arezzo
' Returns relevant error message in sErrMessage
' and directly updates the moTempValue object (assumed initialised) with
'   ResponseStatus
'   ValidationMessage
' Returns TRUE if response was accepted, or FALSE if response was rejected
'----------------------------------------------------
Dim lResult As Long
Dim sWFlag As String        ' Warning flag from Arezzo
Dim bOKSoFar As Boolean
Dim oValidation As Validation

    On Error GoTo ErrLabel
    
    lResult = moArezzo.ValidateResponse(moElement.Code, Me.RepeatNumber, Me.EFormInstance, _
                                        sArezzoValue, sWFlag)
    ' If Arezzo rejected it, fill in the error message
    sErrMessage = GetResponseErrorString(lResult)
    bOKSoFar = (sErrMessage = "")
    
    If bOKSoFar Then
        ' It was either Warning or OK
        If lResult = eResponseValidation.ArezzoWarning Then
            ' Here we have to analyse the sWFlag to see if it represents
            ' Reject, Warning or Inform
            Set oValidation = moElement.Validation(CInt(sWFlag))
            ' Assume AREZZO context still set from ValidateResponse
            sErrMessage = moArezzo.EvaluateExpression(oValidation.MessageExpr)
            If Not moArezzo.ResultOK(sErrMessage) Then
                sErrMessage = oValidation.MessageExpr
            End If
            
            ' Is this a Reject, Warn or Inform validation?
            Select Case oValidation.ValidationType
            Case eValidationType.Reject  ' Reject
                bOKSoFar = False
            Case eValidationType.Warn   ' Warn
                moTempValue.ValidationMessage = sErrMessage
                moTempValue.ResponseStatus = eStatus.Warning
            Case eValidationType.Inform  ' Inform
                moTempValue.ValidationMessage = sErrMessage
                moTempValue.ResponseStatus = eStatus.Inform
            End Select
        Else
            ' It was all OK!
            moTempValue.ValidationMessage = ""
            moTempValue.ResponseStatus = eStatus.Success
        End If
    End If
    
    ' Now me:value and local copy aren't needed any more
    Call moArezzo.DeleteLocalValues(moElement.Code)

    ArezzoValidate = bOKSoFar
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.ArezzoValidate(" & sArezzoValue & ")"

End Function

'----------------------------------------------------
Private Function ValidateLabValue(ByVal sLabValue As String, _
                                ByRef sErrMessage As String) As Boolean
'----------------------------------------------------
' Validate a lab test question
' Returns relevant error message in sErrMessage
' and directly updates the moTempValue object (assumed initialised)
' Returns TRUE if response was accepted, or FALSE if response was rejected
'----------------------------------------------------
Dim lResult As Long
Dim dblLabResponse As Double
Dim dblDOB As Double
Dim nGender As Integer

    On Error GoTo ErrLabel
    
    'TA 07/04/2003: val removed as it prevents this from working with regional settings
    dblLabResponse = CDbl(sLabValue)
    ' Recalculate lab test date
    moTempValue.LabTestDate = GetLabTestDate
    ' Get Gender and DOB
    nGender = moSubject.Gender
    dblDOB = moSubject.DOB
    lResult = moLabTestValidate.NormalRangeStatus(dblLabResponse, _
                                moTempValue.LabTestDate, _
                                dblDOB, _
                                nGender, _
                                sErrMessage)
    
    If lResult <> eStatus.InvalidData Then
        ' NCJ 19 Mar 07 - Issue 2872 - Check warnings on hidden lab question
        If lResult = eStatus.Warning And moElement.Hidden And Not moSubject.HiddenLabWarnings Then
            ' They don't want the warning - convert it to Inform
            lResult = eStatus.Inform
        End If
        moTempValue.ResponseStatus = lResult
        moTempValue.ValidationMessage = sErrMessage
        moTempValue.NRStatus = moLabTestValidate.NormalRangeResult(dblLabResponse, _
                                moTempValue.LabTestDate, _
                                dblDOB, _
                                nGender)
        moTempValue.CTCGrade = moLabTestValidate.CTCGrade(dblLabResponse, _
                                moTempValue.LabTestDate, _
                                dblDOB, _
                                nGender)
        ValidateLabValue = True
    Else
        ' It was invalid
        ValidateLabValue = False
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.ValidateLabValue(" & sLabValue & ")"

End Function

'-----------------------------------------------------------------------
Private Function EmptyResponseStatus() As eStatus
'-----------------------------------------------------------------------
' NCJ 7 Jul 04 - Get the correct status for an empty response
' depending on whether it's Optional or not
'-----------------------------------------------------------------------
    
    If moElement.IsOptional Then
        EmptyResponseStatus = eStatus.Success
    Else
        EmptyResponseStatus = eStatus.Missing
    End If

End Function

'-----------------------------------------------------------------------
Public Function RevalidateValue(ByRef sErrMessage As String, _
                                ByRef bStatusChanged As Boolean, _
                                Optional ByVal dblTimestamp As Double = 0) As Long
'-----------------------------------------------------------------------
' Revalidate the currently stored value
' and return the status and any relevant error message
' Assume we don't get any Unobtainable or NotApplicable questions here
' bStatusChanged is TRUE if the status has changed
' Needs to be followed by ConfirmRevalidation to store a changed status
' RS 26/09/2002 Added dblTimestamp parameter
' NCJ 20 May 03 - We DO revalidate empty responses in case they've changed from optional/mandatory
' NCJ 11 Jun 03 - Also check that N/A questions still have a collect if condition (bug 844)
' NCJ 7 July 04 - Use new EmptyResponseStatus function to deal with blank optional qus.
' NCJ 5 Feb 07 - Issue 2874 - moTempValue.Value must contain LOCALISED value
'-----------------------------------------------------------------------
Dim bOKSoFar As Boolean

    bStatusChanged = False
    sErrMessage = ""
        
    ' Initialise temporary response storage
    Call InitTempResponse(dblTimestamp)
    
    ' NCJ 20 May 03 - Handle empty responses
    If Me.Value = "" Then
        ' Assume the current status value to start with
        ' (NCJ 18 Jun 03 - Moved this line of code within the If Me.Value="" section)
        moTempValue.ResponseStatus = Me.Status
    
        ' Only look at OK/Missing and N/A
        Select Case Me.Status
        Case eStatus.Missing, eStatus.Success
            moTempValue.ResponseStatus = EmptyResponseStatus
        Case eStatus.NotApplicable
            If moElement.CollectIfCond = "" Then
                ' NCJ 23 Jun 03 - Bug 1866 - Make sure it's not in a Not Applicable group
                If moElement.OwnerQGroup Is Nothing Then
                    ' Its Collect If condition has gone, so set it to OK/Missing
                    moTempValue.ResponseStatus = EmptyResponseStatus
                Else
                    ' Only set it to OK/Missing if the Group isn't N/A
                    If OwnerQGI.Status <> eStatus.NotApplicable Then
                        moTempValue.ResponseStatus = EmptyResponseStatus
                    End If
                End If
            End If
        Case Else
            ' Stay with what we've got
        End Select
    Else
        ' Ensure that we know the Arezzo value
        Call GetArezzoValue
        
        ' Special case of LABTEST questions
        If moElement.DataType = eDataType.LabTest Then
            bOKSoFar = ValidateLabValue(Me.Value, sErrMessage)
        Else
            ' Let Arezzo do its validation
            bOKSoFar = ArezzoValidate(msArezzoValue, sErrMessage)
        End If
    
        If bOKSoFar Then
            ' NB moTempValue.Value stays empty if value rejected
            ' Not rejected, so fill in the rest of our temporary copy
            ' NCJ/MLM 23 Jan 03 (Issue 675) - Fill in DisplayValue too
    '        moTempValue.Value = Me.Value
            moTempValue.DisplayValue = Me.Value
'            moTempValue.Value = moResponseState.Value
            ' NCJ 5 Feb 07 - Issue 2874 - Temp Value must have LOCALISED value
            moTempValue.Value = moVValue.CurrentValue
            moTempValue.ArezzoValue = msArezzoValue
            If moElement.DataType = eDataType.Category Then
                ' Store the code
                moTempValue.ValueCode = msArezzoValue
            Else
                moTempValue.ValueCode = ""
            End If
        End If
    End If
    
    ' Decide on whether there's been a change
    ' based on original/new status
    Select Case Me.Status
    Case eStatus.Inform, eStatus.Warning
        bStatusChanged = (moTempValue.ResponseStatus <> Me.Status)
        If Not bStatusChanged Then
            ' See if validation message is different
            bStatusChanged = (sErrMessage <> Me.ValidationMessage)
        End If
    Case eStatus.OKWarning
        If moTempValue.ResponseStatus = eStatus.Warning Then
            ' See if validation message is different
            bStatusChanged = (sErrMessage <> Me.ValidationMessage)
        Else
            bStatusChanged = True
        End If
    Case Else
        bStatusChanged = (moTempValue.ResponseStatus <> Me.Status)
    End Select
    
    ' NCJ 27 Aug 02 - Extra checks for the changed-ness of lab test questions
    If Not bStatusChanged And (moElement.DataType = eDataType.LabTest) Then
        ' Must see if NRStatus or CTCGrade have changed
        bStatusChanged = (moTempValue.CTCGrade <> Me.CTCGrade) Or (moTempValue.NRStatus <> Me.NRStatus)
    End If
    
    ' Return the Response Status of our temporary store
    RevalidateValue = moTempValue.ResponseStatus

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.RevalidateValue"

End Function

'----------------------------------------------------
Public Sub ConfirmRevalidation(sOverrule As String)
'----------------------------------------------------
' Confirm a revalidated value, i.e. after a call to RevalidateValue
' Decides on appropriate RFC (assuming there has been a change to the response)
' NB If moTempValue status is Invalid, switch it to Missing
' NCJ 27 Aug 02 - Deal with RFCs for lab test questions
'----------------------------------------------------
Dim sRFC As String
Dim enContext As eAutoRFCContext
Dim enChangeType As eAutoRFCChange
Dim bNeedsRFC As Boolean

    On Error GoTo ErrLabel
    
    sRFC = ""
    bNeedsRFC = False
    
'    If moTempValue.Value = "" Then
    If moTempValue.ResponseStatus = eStatus.InvalidData Then
        ' Value was rejected so switch status from Invalid to Missing
        moTempValue.ResponseStatus = eStatus.Missing
    End If
    
    ' First check the status change
    bNeedsRFC = RequiresStatusRFC(moTempValue.ResponseStatus, moTempValue.ValidationMessage)
    
    If bNeedsRFC Then
        ' Assume we have a status change
        enChangeType = StatusChanged
    ElseIf moElement.DataType = eDataType.LabTest Then
        ' Do extra check on lab result for lab test questions
        bNeedsRFC = RequiresLabResultRFC(moTempValue.NRStatus, moTempValue.CTCGrade)
        If bNeedsRFC Then
            ' Assume it's a lab result change
            enChangeType = LabResultChanged
        End If
    End If
    
    ' Set up RFC if required
    If bNeedsRFC Then
        ' Set the context
        If moElement.DataType = eDataType.LabTest Then
            enContext = eAutoRFCContext.ChangingLab
        Else
            enContext = eAutoRFCContext.Revalidation
        End If
        If moTempValue.Value = "" Then
            ' It was rejected
            sRFC = Me.EFormInstance.AutomaticRFC(enContext, ValueRejected)
        Else
            ' Status change
            If moVValue.CurrentValue <> moVValue.OriginalValue Then
                ' User has changed the value since last save
                ' so stay with their RFC
                sRFC = Me.ReasonForChange
            Else
                ' Value has not changed since last save
                sRFC = Me.EFormInstance.AutomaticRFC(enContext, enChangeType)
            End If
        End If
    End If
    
    ' Now do normal confirm
    Call ConfirmValue(sOverrule, sRFC, "")
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.ConfirmRevalidation"

End Sub

'----------------------------------------------------
Private Function GetArezzoValue() As String
'----------------------------------------------------
' Get the "Arezzo" value of this response from its stored "value",
' i.e. the value that we pass to Arezzo
' and store it as msArezzoValue
' Set mbKnownArezzoValue if already known
'----------------------------------------------------
Dim sDisplayText As String

    On Error GoTo ErrLabel
    
    ' Do we already know it?
    If mbKnownArezzoValue Then
        GetArezzoValue = msArezzoValue
        Exit Function
    End If
    
    msArezzoValue = Me.Value
    
    ' See if we need to Arezzo-ise it
    If Me.Value > "" Then
        Select Case moElement.DataType
        Case eDataType.IntegerNumber, eDataType.Real, eDataType.LabTest
            Call moElement.Formatter.ValidNumber(msArezzoValue, sDisplayText, msArezzoValue)
            
        Case eDataType.Category
            msArezzoValue = Me.ValueCode
            
        Case eDataType.DateTime
            ' Must use moResponseState.Value in case Me.Value is locally formatted
            ' NCJ 18 Dec 06 - Beware of not having an Arezzo
            If Not moArezzo Is Nothing Then
                Call moArezzo.ReadValidDate(moResponseState.Value, moElement.Format, msArezzoValue)
            End If
        Case eDataType.Text
            msArezzoValue = TextifyArezzoValue(msArezzoValue)

        Case eDataType.MultiMedia
            ' Leave unchanged
        
        End Select
    End If
    
    mbKnownArezzoValue = True
    GetArezzoValue = msArezzoValue

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.GetArezzoValue"

End Function

'----------------------------------------------------
Private Function TextifyArezzoValue(ByVal sValue As String) As String
'----------------------------------------------------
' NCJ 17 Feb 03 - SR 3921
' Force a non-empty value to be a "text" value in AREZZO by making it into a Prolog atom
'----------------------------------------------------

    If sValue > "" Then
        TextifyArezzoValue = "'" & ReplaceQuotes(sValue) & "'"
    Else
        TextifyArezzoValue = sValue
    End If
    
End Function

'----------------------------------------------------
Public Function RequiresOverruleRFC(ByVal sOverrule As String) As Boolean
'----------------------------------------------------
' NCJ 12 Aug 02 - This is for when user changes warning overrule
' Does this response need an RFC for an overrule change?
' sOverrule is the proposed new overrule reason
' NB sOverrule = "" means Status = Warning
'----------------------------------------------------

    RequiresOverruleRFC = False
    
    ' Is question's RFC flag set?
    If Not moElement.RequiresRFC Then Exit Function
    
    ' Has the response been changed before?
    If moResponseState.ChangeCount = 0 Then Exit Function
    
    ' RFC required only if value has not changed since last save
    ' and if new overrule differs from saved overrule
    If moVValue.OriginalValue = moVValue.CurrentValue Then
        RequiresOverruleRFC = (sOverrule <> moVOverrule.OriginalValue)
    End If

End Function

'----------------------------------------------------
Public Function RequiresLabResultRFC(ByVal nNewNRStatus As Integer, _
                                    ByVal nNewCTCGrade As Integer) As Boolean
'----------------------------------------------------
' NCJ 27 Aug 02 - This is for when lab result changes
' e.g. when user changes lab
' Does this response need an RFC for a lab result change?
' nNewNRStatus and nNewCTCGrade are the proposed new results
' NCJ 30 Jun 04 - Ignore input params and use moTempValue instead (otherwise routine is unusable!)
'       Params are kept to preserve compatibility
'----------------------------------------------------

    RequiresLabResultRFC = False
    
    ' Is question's RFC flag set?
    If Not moElement.RequiresRFC Then Exit Function
    
    ' Has the response been changed before?
    If moResponseState.ChangeCount = 0 Then Exit Function
    
    ' RFC required if new result differs from saved result
    ' NCJ 30 Jun 04 - Always compare with moTempValue
    RequiresLabResultRFC = (moTempValue.NRStatus <> moVNRStatus.OriginalValue) _
                            Or (moTempValue.CTCGrade <> moVCTCGrade.OriginalValue)

End Function

'----------------------------------------------------
Public Function RequiresStatusRFC(ByVal nNewStatus As Integer, Optional sNewValMessage As String = "") As Boolean
'----------------------------------------------------
' NCJ 12 Aug 02 - This is for when status changes
' e.g. when user toggles between Missing and Unobtainable
' Does this response need an RFC for a status change?
' nStatus is the proposed new status
' and sNewValMessage is the new validation message
' (could be that after revalidation the status is the same but the val. message is different)
' NCJ 4 Sept 02 - Added optional ValMessage argument
'----------------------------------------------------

    RequiresStatusRFC = False
    
    ' Is question's RFC flag set?
    If Not moElement.RequiresRFC Then Exit Function
    
    ' Has the response been changed before?
    If moResponseState.ChangeCount = 0 Then Exit Function
    
    If (nNewStatus <> moVStatus.OriginalValue) Then
        ' New status is different from saved one
        RequiresStatusRFC = True
    Else
        ' New status is same as saved status so check val message (if any)
        If sNewValMessage <> "" Then
            RequiresStatusRFC = (sNewValMessage <> moVValMessage.OriginalValue)
        Else
            ' Result remains as False
        End If
    End If
    
End Function

'----------------------------------------------------
Public Function RequiresValueRFC() As Boolean
'----------------------------------------------------
' NCJ 9 Aug 02 - This replaces DiffersFromSavedValue
' Does this response need an RFC for a value change?
' Yes if the question's RFC flag is set and the response has a history
' and the new value is different from saved
' ASSUME the new value is in moTempValue
'----------------------------------------------------

    RequiresValueRFC = False
    
    ' Is question's RFC flag set?
    If Not moElement.RequiresRFC Then Exit Function
    
    ' Has the response been changed before?
    If moResponseState.ChangeCount = 0 Then Exit Function
    
    RequiresValueRFC = (moTempValue.Value <> moVValue.OriginalValue)

End Function

'----------------------------------------------------
'Public Property Get DiffersFromSavedValue() As Boolean
'----------------------------------------------------
' NCJ 14 Aug 02 - NO LONGER USED
'   See instead RequiresValueRFC, RequiresStatusRFC, RequiresLabResultRFC and RequiresOverruleRFC
'----------------------------------------------------

'----------------------------------------------------------------------------
Public Property Get Revalidatable(bLabTestQuestions As Boolean) As Boolean
'----------------------------------------------------------------------------
' Can we revalidate this question?
' If bLabTestQuestions = TRUE then only look at lab test qus.,
' otherwise look at non-labtest qus. NCJ 24 Mar 03 - Look at ALL qus!
' i.e. is it unlocked, with non-empty response and some validation conditions?
' NCJ 20 May 03 - Include empty questions because they might have been changed to/from optional/mandatory
' NCJ 30 Jun 04 - Include Warning/OKWarning questions because they might have had validations removed
'----------------------------------------------------------------------------

    ' First sift out the ones we DON'T want to revalidate
    Revalidatable = False
    
    If Me.EFormInstance.ReadOnly Then Exit Property
    
    If Me.LockStatus <> eLockStatus.lsUnlocked Then Exit Property
    
    ' NCJ 20 May 03 - Ignore Unobtainable or Not Applicable (BUT we keep other empties)
    ' NCJ 11 Jun 03 - Keep N/A questions too in case their Collect If condition has gone (bug 844)
    If Me.Status = eStatus.Unobtainable Then Exit Property
    
    If bLabTestQuestions Then
        ' We want lab test questions only
        Revalidatable = (moElement.DataType = eDataType.LabTest)
        Exit Property
    End If
    
    ' Now sift through the ones we DO want to revalidate
    Revalidatable = True
    
    If moElement.DataType = eDataType.LabTest Then Exit Property
    
    ' Empty values - in case of change between optional/mandatory
    If Me.Value = "" Then Exit Property
    
    ' Questions with validations
    If moElement.Validations.Count > 0 Then Exit Property
    
    ' Statuses of Inform/Warning/OKWarning - in case their validations have been removed
    Select Case Me.Status
    Case eStatus.Inform, eStatus.Warning, eStatus.OKWarning
        Exit Property
    End Select
        
    ' And what's left isn't revalidatable
    Revalidatable = False

End Property

'----------------------------------------------------
Private Sub InitTempResponse(Optional dblTimestamp As Double = 0)
'----------------------------------------------------
' Initialise our "temp response" object
' prior to validating a value
' RS 22/9/2002: Added Timestamp parameter (for WEB client use)
'----------------------------------------------------
    
    On Error GoTo ErrLabel

    ' Splat any previous temporary value
    Set moTempValue = Nothing
    
    ' Create new temporary response storage area
    Set moTempValue = New TempResponse
    moTempValue.Value = ""
    moTempValue.ValueCode = ""
    moTempValue.ValidationMessage = ""
    moTempValue.ResponseStatus = eStatus.InvalidData
    moTempValue.ArezzoValue = ""
    moTempValue.NRStatus = eNormalRangeLNorH.nrNotfound
    moTempValue.CTCGrade = eCTCGrade.ctcNone
    If dblTimestamp > 0 Then
        moTempValue.TimeStamp = dblTimestamp
    Else
        moTempValue.TimeStamp = IMedNow
    End If
    moTempValue.LabTestDate = 0
    

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.InitTempResponse"

End Sub

'----------------------------------------------------
Private Sub ResponseHasChanged(Optional dblTimestamp As Double = 0, Optional bFromRequested As Boolean = False)
'----------------------------------------------------
' Set Changed flag, Timestamp and SoftwareVersion
' if any response data is changed
' NCJ 22 Jul 04 - Set the user details here too (bugs 2323, 2350, and probably others)
'----------------------------------------------------

    On Error GoTo ErrLabel

    ' Test for IsChanged
    moResponseState.IsChanged = Dirty
    If moResponseState.IsChanged Then
        If dblTimestamp > 0 Then
            moResponseState.TimeStamp = dblTimestamp
        Else
            'ic/ncj 28/02/2008 dont change timestamp if status is changing from requested
            'so that we dont splat web timestamps
            If (Not bFromRequested) Then
                moResponseState.TimeStamp = IMedNow
            End If
        End If
        moResponseState.TimeSTampTZ = moSubject.Timezone.TimezoneOffset   ' RS 22/9/2002 Might not be necessary though
        moResponseState.SoftwareVersion = GetSoftwareVersion
        ' NCJ 22 Jul 04 - Always log the current user
        moResponseState.UserCode = moSubject.UserName
        moResponseState.UserNameFull = moSubject.UserNameFull
        
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "Response.ResponsehasChanged"

End Sub
 
'----------------------------------------------------
Friend Property Get LabTestValidate() As LabTestValidate
'----------------------------------------------------
' The lab test validation object
'----------------------------------------------------

    Set LabTestValidate = moLabTestValidate

End Property

'----------------------------------------------------------------------------
Friend Property Set LabTestValidate(oLabTestValidate As LabTestValidate)
'----------------------------------------------------------------------------
' The lab test validation object
'----------------------------------------------------------------------------
' NCJ 4 Mar 02 - Must be able to set this object when creating new Responses
' as repeating copies of previous ones
'----------------------------------------------------------------------------

    Set moLabTestValidate = oLabTestValidate

End Property

'----------------------------------------------------
Private Function GetSoftwareVersion() As String
'----------------------------------------------------
' The current software version
' NB This is the version of the MACRODEBS DLL
'----------------------------------------------------

    GetSoftwareVersion = App.Major & "." & App.Minor & "." & App.Revision
    
End Function
 
'----------------------------------------------------
Public Property Get Enterable() As Boolean
'----------------------------------------------------
' Returns TRUE if this response is enterable by the user.
' If not, the user should be prevented from entering a value.
'----------------------------------------------------

    ' Enterable it if not derived or Not Applicable
    ' and the user has correct permissions etc.
    If Me.EFormInstance.ReadOnly _
        Or (Me.Status = eStatus.NotApplicable) _
        Or (moElement.DerivationExpr > "") _
        Or moElement.Hidden _
        Or (Me.LockStatus <> eLockStatus.lsUnlocked) _
        Or (moElement.DataType = eDataType.Category And Not moElement.ActiveCategories) Then
            Enterable = False
    Else
            Enterable = True
    End If

End Property

'----------------------------------------------------------------------------------------'
Friend Property Let HadValue(bHadValue As Boolean)
'----------------------------------------------------------------------------------------'
' new had value property for GGB
' gets set to true and stays true whenever a non empty value is set
'----------------------------------------------------------------------------------------'

    moResponseState.HadValue = bHadValue
    
End Property

'----------------------------------------------------
Public Property Get ArezzoValue() As String
'----------------------------------------------------
' The Arezzo Value of this response.
' i.e. the value that has been sent to Arezzo
'----------------------------------------------------

    ArezzoValue = GetArezzoValue

End Property

'----------------------------------------------------
Public Property Get ChangeCount() As Integer
'----------------------------------------------------
' The number of changes this response has been through
' Read-only property but see also SetChangeCount
'----------------------------------------------------

    ChangeCount = moResponseState.ChangeCount
        
End Property

'----------------------------------------------------
Friend Sub SetChangeCount()
'----------------------------------------------------
' Update the ChangeCount - to be called just before saving the response
' Only update count if response has changed "significantly"
'----------------------------------------------------

    ' Don't bother if response is New and still blank
    ' (this includes Missing or Not Applicable, or Success for Optional questions,
    ' but excludes Unobtainable because this is a deliberate user change)
    If moResponseState.IsNew _
     And moResponseState.Value = "" _
     And moResponseState.Status <> eStatus.Unobtainable Then Exit Sub
    
    ' Don't bother if response is unchanged
    If Not moResponseState.IsChanged Then Exit Sub
    
    ' NCJ 29 Oct 02 - Ignore status changes between Missing and Not Applicable
    Select Case moVStatus.OriginalValue
    Case eStatus.Missing
        If moResponseState.Status = eStatus.NotApplicable Then Exit Sub
    Case eStatus.NotApplicable
        If moResponseState.Status = eStatus.Missing Then Exit Sub
    End Select
    
    ' Otherwise see if there's been a "significant" change
    ' i.e. change to its Value, Status, Overrule, ValMessage or LabCode
    If moVValue.Dirty _
     Or moVStatus.Dirty _
     Or moVOverrule.Dirty _
     Or moVValMessage.Dirty _
     Or moVLabCode.Dirty Then
        ' Increment change count
        moResponseState.ChangeCount = moResponseState.ChangeCount + 1
    End If
    
End Sub

'----------------------------------------------------
Public Property Get SDVStatus() As eSDVStatus
'----------------------------------------------------
' TA 14 Aug 02
' The current "SDV status" of the response
'----------------------------------------------------

    SDVStatus = moResponseState.SDVStatus

End Property

'----------------------------------------------------
Public Property Let SDVStatus(ByVal enSDVStatus As eSDVStatus)
'----------------------------------------------------
' TA 14 Aug 02
' The current "SDV status" of the response
' NB This doesn't update the "Changed" flag
'----------------------------------------------------

    moResponseState.SDVStatus = enSDVStatus

End Property

'----------------------------------------------------
Public Property Get NoteStatus() As eNoteStatus
'----------------------------------------------------
' TA 14 Aug 02
' The current "Note status" of the response
'----------------------------------------------------

    NoteStatus = moResponseState.NoteStatus

End Property

'----------------------------------------------------
Public Property Let NoteStatus(ByVal enNoteStatus As eNoteStatus)
'----------------------------------------------------
' TA 14 Aug 02
' The current "Note status" of the response
' NB This doesn't update the "Changed" flag
'----------------------------------------------------

    moResponseState.NoteStatus = enNoteStatus

End Property

'----------------------------------------------------
Public Property Get DiscrepancyStatus() As eDiscrepancyStatus
'----------------------------------------------------
' TA 14 Aug 02
' The current "Discrepancy status" of the response
'----------------------------------------------------

    DiscrepancyStatus = moResponseState.DiscrepancyStatus

End Property

'----------------------------------------------------
Public Property Let DiscrepancyStatus(ByVal enDiscrepancyStatus As eDiscrepancyStatus)
'----------------------------------------------------
' TA 14 Aug 02
' The current "Discrepancy status" of the response
' NB This doesn't update the "Changed" flag
'----------------------------------------------------

    moResponseState.DiscrepancyStatus = enDiscrepancyStatus

End Property

'----------------------------------------------------
Public Property Get EFormInstance() As EFormInstance
'----------------------------------------------------
' MLM 06/09/02: Added. Return the Response's EFI.
'----------------------------------------------------

    Set EFormInstance = moSubject.eFIByTaskId(mlEFormTaskID)
    
End Property

'----------------------------------------------------
Friend Sub ReapplyArezzo()
'----------------------------------------------------
' NCJ 18 Sept 02
' Reapply current response value to Arezzo if necessary
'----------------------------------------------------

    If mbArezzoValueChanged Then
        Call moArezzo.ReplaceQData(msArezzoDataName, msArezzoValue, msArezzoTag, "")
    End If
    If mbArezzoStatusChanged Then
        Call moArezzo.SetQuestionStatus(moElement.Code, moResponseState.Status)
    End If
    ' NCJ 10 Apr 03 Bug 1232 - Ensure we update the visit/form dates too
    Call AddAsEFormAndVisitDate

End Sub

'----------------------------------------------------
Friend Sub UpdateTimeStamp(nClientTimezoneOffset As Integer)
'----------------------------------------------------
' RS 22/9/2002: Convert the Timestamp & TimezoneOffset values currently stored in the Response object
'               to the corresponding values in the specified timezone
'               ClientTimezoneOffset in minutes
'----------------------------------------------------
Dim dblGMTtimestamp As Double
Dim dblLocaltimestamp As Double

    ' STEP 1: Convert current timestamp to GMT
    dblGMTtimestamp = moResponseState.TimeStamp + (moResponseState.TimeSTampTZ / (60 * 24))
    
    ' STEP 2: Convert GTM timestamp to local (inverted operation, thus substract)
    dblLocaltimestamp = dblGMTtimestamp - (nClientTimezoneOffset / (60 * 24))

    moResponseState.TimeStamp = dblLocaltimestamp
    moResponseState.TimeSTampTZ = nClientTimezoneOffset


End Sub

'----------------------------------------------------
Private Function OwnerQGI() As QGroupInstance
'----------------------------------------------------
' NCJ 23 Jun 03
' The QGroup Instance to which this response belongs
' Nothing if not a group question
'----------------------------------------------------

    If Not moElement.OwnerQGroup Is Nothing Then
        Set OwnerQGI = Me.EFormInstance.QGroupInstanceById(moElement.OwnerQGroup.QGroupID)
    End If
    
End Function

