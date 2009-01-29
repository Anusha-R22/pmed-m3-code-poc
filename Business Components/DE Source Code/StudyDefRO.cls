VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyDefRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: StudyDefRO.cls
' Copyright InferMed Ltd 2001-2004 All Rights Reserved
' Author: Nicky Johns, InferMed, May 2001
' Study Definition (read-only) class for DE in MACRO 2.2
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 21-29 May 01 - Initial development
' NCJ 18 Jun 01 - Load eFormElements
' NCJ 25 Jun 01 - Changed LoadEFormElements to accept string from DS
' NCJ 3 Jul 01 - Removed HideIfInactiveEForms (unused)
' ZA 07/08/01 - Added Font properties
' ZA 08/08/01 - Added reason for change (RFCs) property
' TA 13/08/01 - Load and removing of subject from memory now done through this class
' TA 5/9/01 - Connection string now stored on load and used for every subsequent operation
' NCJ 20/9/01 - Added RR properties for Registration
' NCJ 24/9/01 - LoadSubject and NewSubject no longer take Arezzo parameter
'               but they do need UserName parameter
'                Added Terminate method
'               (TA) changed initialisation as a consequence
' NCJ 27 Sep 01 - Added Status, StudyType and Phase
' NCJ 28 Sep 01 - Let Load return message string
' NCJ 1 Oct 01 - Only return message string in Load if Arezzo is there
' TA 12/10/2001: Added property set gets fro the arezzo_dm object
' TA  12 Oct 01 - ArezzoFile is only retrieved from DataServices if we have an Arezzo
' NCJ 22 Mar 02 - Added nUpdateMode to LoadSubject
' TA 03/02/2002 - Added LoadElements and RemoveElements for WWW eForm generation
' MACRO 3.0
' NCJ 5 July 2002 - Merged 2.2 changes into 3.0 version
'   NCJ 8 Oct 02 - Added loading of eForm Date Elements for study
'   NCJ 9 Oct 02 - Added Reasons For Overrule (RFOs)
' NCJ 24 Jan 03 - Added sCountry (the site's country) to NewSubject
' NCJ 6 May 03 - Added new parameters sUserNameFull and sUserRole to NewSubject and LoadSubject
' NCJ 12 May 03 - Added Set/Get LocalNumberFormats
' NCJ 7 Aug 03 - Added STR to collection keys in EFormByCode and VisitByCode
' NCJ 24 Mar 04 - Clear out object collections in Terminate
'----------------------------------------------------

Option Explicit

' The study def state
Private moStudyDefState As StudyDefState

'Study's subject
Private moSubject As StudySubject

' The cross-reference collections
Private mcolEFormsByCode As Collection
Private mcolVisitsByCode As Collection

' The eForms and Visits
Private mcolVisits As Collection
Private mcolEForms As Collection

Private moArezzo As Arezzo_DM

'store connection string used for load
Private msConnection As String

' These are the defaults for the Study
Private mlFontColour As Long
Private msFontName As String
Private mnFontSize As Integer
Private mlEFormColour As Long

' NCJ 12 May 03 - Local number formatting
Private msDecimalPtChar As String
Private msThousandSepChar As String

'-------------------------------------------------------------------
Public Function Load(ByVal sConnection As String, _
                    ByVal lStudyId As Long, _
                    ByVal nVersion As Integer, _
                    Optional ByRef oArezzo As Arezzo_DM = Nothing) As String
'-------------------------------------------------------------------
' Load an existing study specified by given studyID and version.
' sConnection is the DB connection string.
' oArezzo is the Arezzo instance (already initialised).
' Returns error message string (empty if no error).
'-------------------------------------------------------------------
Dim oStudyDS As StudyDefDS
Dim bHaveArezzo As Boolean
Dim sState As String

    On Error GoTo ErrLabel
    
    ' Tidy up first in case we were storing a different study
    Call Terminate

    'initialise objects
    Set moStudyDefState = New StudyDefState
    Set mcolEFormsByCode = New Collection
    Set mcolVisitsByCode = New Collection
    'store connection string used for load
    msConnection = sConnection
    
    ' Load the study def into Arezzo
    Set moArezzo = oArezzo
    
    bHaveArezzo = Not (oArezzo Is Nothing)
    
    Set oStudyDS = New StudyDefDS
    ' Use Data Services to load the study def state
    sState = oStudyDS.Load(sConnection, lStudyId, nVersion, bHaveArezzo)
    
    If sState = "" Then
        'there is no study matching this studyid
        'return no study error message
        Load = "The study no longer exists"
'EXIT FUNCTION HERE
        Exit Function
    End If
    
    moStudyDefState.State = sState
    Set oStudyDS = Nothing

    ' Only load Arezzo definition if we have an Arezzo object
    If bHaveArezzo Then
        oArezzo.ArezzoDef = moStudyDefState.ArezzoFile
    End If
    ' For efficiency, now clear the study def
    moStudyDefState.ArezzoFile = ""
    
    'REM 12/02/02 - Moved SetUpFontDetails call to before SetUpStudyCrossRefs call
    Call SetUpFontDetails
    Call SetUpStudyCrossRefs

    ' NCJ 15 Oct 01 - Check for all the things that might cause problems
    Load = CheckStudyOpenable(bHaveArezzo)
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|StudyDefRO.Load"

End Function

'-------------------------------------------------------------------
Private Sub SetUpFontDetails()
'-------------------------------------------------------------------
' Set up the font details, using defaults if none given here
'-------------------------------------------------------------------

' The default font details to be used when none other exist
Const l_DEFAULT_FontColour As Long = -2147483630
Const l_DEFAULT_EFormColour As Long = 12632256
Const s_DEFAULT_FontName As String = "Arial"
Const n_DEFAULT_FontSize As Integer = 10

    With moStudyDefState
        If .FontName > "" Then
            msFontName = .FontName
        Else
            msFontName = s_DEFAULT_FontName
        End If
        If .FontSize > 0 Then
            mnFontSize = .FontSize
        Else
            mnFontSize = n_DEFAULT_FontSize
        End If
        If .FontColour <> 0 Then
            mlFontColour = .FontColour
        Else
            mlFontColour = l_DEFAULT_FontColour
        End If
        If .eFormColour <> 0 Then
            mlEFormColour = .eFormColour
        Else
            mlEFormColour = l_DEFAULT_EFormColour
        End If
    End With

End Sub

'-------------------------------------------------------------------
Private Function CheckStudyOpenable(bHaveArezzo As Boolean) As String
'-------------------------------------------------------------------
' Check all the things that make the study not openable,
' e.g. infinitely cycling visits,
'   no Visits,
'   no eForms,
'   no Visit eForms, etc.
' Return error message on failure,
' otherwise return empty string.
' bHave Arezzo is whether we can use Arezzo
'-------------------------------------------------------------------
Dim sMsg As String
Dim oVisit As VisitRO
Dim bVisitEForm As Boolean

Const sCANT_BE_USED = "This study cannot be used for data entry because "

    On Error GoTo ErrLabel
    
    sMsg = ""
    
    ' Check for infinitely cycling visits
    If bHaveArezzo Then
        sMsg = moArezzo.CheckStudy
    End If
    
    If sMsg = "" Then   ' Was Arezzo OK about it?
        ' Check there are some visits & forms
        If Me.Visits.Count = 0 Then
            sMsg = sCANT_BE_USED & "it contains no visits"
        ElseIf Me.eForms.Count = 0 Then
            sMsg = sCANT_BE_USED & "it contains no eForms"
        Else
            bVisitEForm = False
            ' Check there's at least one eForm in a Visit
            For Each oVisit In Me.Visits
                If oVisit.VisitEForms.Count > 0 Then
                    ' We have a visit eForm
                    bVisitEForm = True
                    Exit For
                End If
            Next
            If Not bVisitEForm Then
                sMsg = sCANT_BE_USED & "its schedule is empty"
            End If
        End If
    End If
    
    CheckStudyOpenable = sMsg

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|StudyDefRO.CheckStudyOpenable"

End Function

'-------------------------------------------------------------------
Public Property Get ConnectionString() As String
'-------------------------------------------------------------------
' Return connection string used for load.
'-------------------------------------------------------------------

    ConnectionString = msConnection
    
End Property

'-------------------------------------------------------------------
Public Property Get Subject() As StudySubject
'-------------------------------------------------------------------
' Return the study's single subject (multi subject not yet supported).
'-------------------------------------------------------------------

    Set Subject = moSubject
    
End Property

'-------------------------------------------------------------------
Public Function LoadSubject(ByVal sSite As String, _
                    ByVal lPersonID As Long, sUserName As String, _
                    ByVal nUpdateMode As eUIUpdateMode, _
                    ByVal sUserNameFull As String, _
                    ByVal sUserRole As String, _
                    Optional bUpdateCachetable As Boolean = True) As StudySubject
'-------------------------------------------------------------------
' Subject loading done through the study.
' This automatically closes (but does not save) any previous subject.
' sUserName is the name of the current MACRO user,
' and nUpdateMode indicates whether the user can change data (see eUIUpdateMode enumeration)
'-------------------------------------------------------------------
' NCJ 20 Nov 01 - Added nUpdateMode which says whether current user can change data
' TA 09/04/2002: Parameter to update cache table added
' nb AFTER CALLING THIS, YOU MUST CHECK StudySubject.CouldNotLoad PROPERTY
' NCJ 6 May 03 - Added new parameters sUserNameFull and sUserRole
'-------------------------------------------------------------------

    Call RemoveSubject
    Set moSubject = New StudySubject
    moSubject.Load msConnection, Me, sSite, lPersonID, sUserName, nUpdateMode, _
                        bUpdateCachetable, _
                        sUserNameFull, sUserRole, moArezzo
    Set LoadSubject = moSubject
    
End Function

'-------------------------------------------------------------------
Public Sub RemoveSubject()
'-------------------------------------------------------------------
' Remove subject from memory.
'-------------------------------------------------------------------
    
    If Not moSubject Is Nothing Then
        Call moSubject.Terminate
        Set moSubject = Nothing
    End If

End Sub

'-------------------------------------------------------------------
Public Function NewSubject(ByVal sSite As String, sUserName As String, _
                        ByVal sCountry As String, _
                        ByVal sUserNameFull As String, _
                        ByVal sUserRole As String) As StudySubject
'-------------------------------------------------------------------
' Create a new Study Subject for the given Site.
' This automatically closes (but does not save) any previous subject.
' sUserName is the name of the current MACRO user.
' NCJ 24 Jan 03 - Added sCountry (the site's country)
' NCJ 6 May 03 - Added new parameters sUserNameFull and sUserRole
'-------------------------------------------------------------------

    ' Shift the previous one out of the way
    Call RemoveSubject
    Set moSubject = New StudySubject
    Call moSubject.NewSubject(msConnection, Me, sSite, sUserName, sCountry, _
                            sUserNameFull, sUserRole, moArezzo)
    Set NewSubject = moSubject

End Function

'-------------------------------------------------------------------
Private Sub SetUpStudyCrossRefs()
'-------------------------------------------------------------------
' Set up cross-reference collections to map between IDs and Codes
' and create Visit and eForms collections
' NCJ 8 Oct 02 - Also deal with the eFormDateElements collection
'-------------------------------------------------------------------
Dim oEFormState As eFormState
Dim oEForm As eFormRO
Dim oVisitState As VisitState
Dim oVisit As VisitRO

    ' NB must do eForms first!
    Set mcolEFormsByCode = New Collection
    Set mcolEForms = New Collection
    
    For Each oEFormState In moStudyDefState.eForms
        ' Store the ID indexed by code
        mcolEFormsByCode.Add oEFormState.EFormId, oEFormState.Code
        ' Create new eForm object
        Set oEForm = New eFormRO
        ' NCJ 8 Oct 02 - Pass the Date Element state too
        Call oEForm.SetState(Me, oEFormState, GetEformDateElementState(oEFormState.EFormId))
        mcolEForms.Add oEForm, oEFormState.Key
    Next
    
    Set mcolVisitsByCode = New Collection
    Set mcolVisits = New Collection
    
    For Each oVisitState In moStudyDefState.Visits
        ' Store the ID indexed by code
        mcolVisitsByCode.Add oVisitState.VisitId, oVisitState.Code
        ' Create new Visit object
        Set oVisit = New VisitRO
        Call oVisit.SetState(Me, oVisitState)
        mcolVisits.Add oVisit, oVisitState.Key
    Next

End Sub

'-------------------------------------------------------------------
Private Function GetEformDateElementState(ByVal lEFormId As Long) As EFElementState
'-------------------------------------------------------------------
' Get the State of the eForm's Date element, if any
' Returns Nothing if no eForm Date element
'-------------------------------------------------------------------
Dim oEFELState As EFElementState

    Set GetEformDateElementState = Nothing
    
    For Each oEFELState In moStudyDefState.eFormDateElements
        ' Match on eFormId
        If oEFELState.EFormId = lEFormId Then
            Set GetEformDateElementState = oEFELState
            Exit For
        End If
    Next

End Function

'-------------------------------------------------------------------
Public Property Get Name() As String
'-------------------------------------------------------------------
'Study name.
'-------------------------------------------------------------------

    Name = moStudyDefState.Name
    
End Property

'-------------------------------------------------------------------
Public Property Get StudyId() As Long
'-------------------------------------------------------------------
'Study ID.
'-------------------------------------------------------------------

    StudyId = moStudyDefState.StudyId
    
End Property

'-------------------------------------------------------------------
Public Property Get Version() As Integer
'-------------------------------------------------------------------
'Study version.
'-------------------------------------------------------------------

    Version = moStudyDefState.Version
    
End Property

'-------------------------------------------------------------------
Public Property Get Description() As String
'-------------------------------------------------------------------
'Study description.
'-------------------------------------------------------------------

    Description = moStudyDefState.Description
    
End Property

'-------------------------------------------------------------------
Public Property Get LabelExpr() As String
'-------------------------------------------------------------------
'Study label expression.
'-------------------------------------------------------------------
    
    LabelExpr = moStudyDefState.LabelExpr
    
End Property

'-------------------------------------------------------------------
Public Property Get LabelIsLocal() As Boolean
'-------------------------------------------------------------------
'Is label local?
'-------------------------------------------------------------------

    LabelIsLocal = moStudyDefState.LabelIsLocal
    
End Property

'-------------------------------------------------------------------
Public Property Get GenderExpr() As String
'-------------------------------------------------------------------
'AREZZO expression for calculating gender.
'-------------------------------------------------------------------

    GenderExpr = moStudyDefState.GenderExpr
    
End Property

'-------------------------------------------------------------------
Public Property Get DOBExpr() As String
'-------------------------------------------------------------------
'AREZZO expression for calculatin DOB.
'-------------------------------------------------------------------

    DOBExpr = moStudyDefState.DOBExpr
    
End Property

'-------------------------------------------------------------------
Public Property Get TimeFormat() As String
'-------------------------------------------------------------------
' The default time format for the study
'-------------------------------------------------------------------

    If moStudyDefState.TimeFormat > "" Then
        TimeFormat = moStudyDefState.TimeFormat
    Else
        TimeFormat = "hh:mm:ss"
    End If
    
End Property

'-------------------------------------------------------------------
Public Property Get DateFormat() As String
'-------------------------------------------------------------------
' The default date format for the study
'-------------------------------------------------------------------

    If moStudyDefState.DateFormat > "" Then
        DateFormat = moStudyDefState.DateFormat
    Else
        DateFormat = "dd/mm/yyyy"
    End If
    
End Property

'-------------------------------------------------------------------
Public Property Get FontName() As String
'-------------------------------------------------------------------
'Default font name.
'-------------------------------------------------------------------

    FontName = msFontName
    
End Property

'-------------------------------------------------------------------
Public Property Get FontBold() As Boolean
'-------------------------------------------------------------------
'Default font bold?
'-------------------------------------------------------------------

    FontBold = moStudyDefState.FontBold
    
End Property

'-------------------------------------------------------------------
Public Property Get FontItalic() As Boolean
'-------------------------------------------------------------------
'Defualt font italic?
'-------------------------------------------------------------------

    FontItalic = moStudyDefState.FontItalic
    
End Property

'-------------------------------------------------------------------
Public Property Get FontSize() As Integer
'-------------------------------------------------------------------
'Default font size.
'-------------------------------------------------------------------

    FontSize = mnFontSize
    
End Property

'-------------------------------------------------------------------
Public Property Get FontColour() As Long
'-------------------------------------------------------------------
'Default font colour for study.
'-------------------------------------------------------------------

    FontColour = mlFontColour
    
End Property

'-------------------------------------------------------------------
Public Property Get eFormColour() As Long
'-------------------------------------------------------------------
'Default eForm colour.
'-------------------------------------------------------------------

    eFormColour = mlEFormColour
    
End Property

'-------------------------------------------------------------------
Public Property Get CTCScheme() As String
'-------------------------------------------------------------------
'Study CTC Scheme.
'-------------------------------------------------------------------

    CTCScheme = moStudyDefState.CTCScheme
    
End Property

'------------------------------------------------
Public Property Get StudyType() As Integer
'------------------------------------------------
' The type of the study
'------------------------------------------------

    StudyType = moStudyDefState.StudyType

End Property

'------------------------------------------------
Public Property Get Status() As Integer
'------------------------------------------------
' The development status of the study definition
'------------------------------------------------

    Status = moStudyDefState.Status

End Property

'------------------------------------------------
Public Property Get Phase() As Integer
'------------------------------------------------
' The study Phase
'------------------------------------------------

    Phase = moStudyDefState.Phase

End Property

'------------------------------------------------
Public Property Get eForms() As Collection
'------------------------------------------------
'Study's collection of eForm objects
'-------------------------------------------------------------------

    Set eForms = mcolEForms
    
End Property

'------------------------------------------------
Public Property Get Visits() As Collection
'------------------------------------------------
'Study's collection of Visit objects.
'-------------------------------------------------------------------

    Set Visits = mcolVisits
    
End Property

'------------------------------------------------
Public Property Get ArezzoFile() As String
'------------------------------------------------
' The CLM definition
'------------------------------------------------

    ArezzoFile = moStudyDefState.ArezzoFile

End Property

'------------------------------------------------
Public Function VisitById(lVisitId As Long) As VisitRO
'------------------------------------------------
' Get the visit definition corresponding to the given VisitId
'------------------------------------------------

    Set VisitById = mcolVisits.Item(Str(lVisitId))

End Function

'------------------------------------------------
Public Function eFormById(lEFormId As Long) As eFormRO
'------------------------------------------------
' Get the eform definition corresponding to the given eFormId
'------------------------------------------------

    Set eFormById = mcolEForms.Item(Str(lEFormId))

End Function

'------------------------------------------------
Public Function eFormByCode(sFormCode As String) As eFormRO
'------------------------------------------------
' Get the eform definition corresponding to the given eForm code
' Read the eFormId from our index collection
' NCJ 7 Aug 03 - Added str to key
'------------------------------------------------

    ' Return Nothing if non-existent
    On Error Resume Next
    Set eFormByCode = mcolEForms.Item(Str(mcolEFormsByCode(sFormCode)))

End Function

'------------------------------------------------
Public Function VisitByCode(sVisitCode As String) As VisitRO
'------------------------------------------------
' Get the eform definition corresponding to the given eForm code
' Read the eFormId from our index collection
' NCJ 7 Aug 03 - Added str to key
'------------------------------------------------

    ' Return Nothing if non-existent
    On Error Resume Next
    Set VisitByCode = mcolVisits.Item(Str(mcolVisitsByCode(sVisitCode)))

End Function

'------------------------------------------------
Public Property Get RFCs() As Collection
'------------------------------------------------
'Study's collection of Reasons for Change.
'------------------------------------------------

    Set RFCs = moStudyDefState.RFCs
    
End Property

'------------------------------------------------
Public Property Get RFOs() As Collection
'------------------------------------------------
'Study's collection of Reasons for Overrule.
'------------------------------------------------

    Set RFOs = moStudyDefState.RFOs
    
End Property

'-------------------------------------------------------------------
Public Property Get RRServerType() As Integer
'-------------------------------------------------------------------
' Registration server type
'-------------------------------------------------------------------

    RRServerType = moStudyDefState.RRServerType

End Property

'-------------------------------------------------------------------
Public Property Get RRHTTPAddress() As String
'-------------------------------------------------------------------
' HTTP address of Remote Registration server
'-------------------------------------------------------------------

    RRHTTPAddress = moStudyDefState.RRHTTPAddress

End Property

'-------------------------------------------------------------------
Public Property Get RRUserName() As String
'-------------------------------------------------------------------
' User name for Remote Registration server
'-------------------------------------------------------------------

    RRUserName = moStudyDefState.RRUserName

End Property

'-------------------------------------------------------------------
Public Property Get RRPassword() As String
'-------------------------------------------------------------------
' Password for Remote Registration server
'-------------------------------------------------------------------

    RRPassword = moStudyDefState.RRPassword

End Property

'-------------------------------------------------------------------
Public Property Get RRProxyServer() As String
'-------------------------------------------------------------------
' Proxy server for Remote Registration
'-------------------------------------------------------------------

    RRProxyServer = moStudyDefState.RRProxyServer

End Property

'-------------------------------------------------------------------
Private Sub Class_Terminate()
'-------------------------------------------------------------------
' Tidy up after ourselves
'-------------------------------------------------------------------

    Call Terminate

End Sub

'-------------------------------------------------------------------
Public Sub Terminate()
'-------------------------------------------------------------------
' Close down the StudyDef object
' NCJ 24 Mar 04 - Must clear out object collections
'-------------------------------------------------------------------
Dim oEForm As eFormRO
Dim oVisit As VisitRO

    On Error GoTo ErrLabel
    
    Call RemoveSubject
    Set moArezzo = Nothing
    Set moStudyDefState = Nothing
    Set mcolEFormsByCode = Nothing
    Set mcolVisitsByCode = Nothing
    
    ' NCJ 24 Mar 04 - These are object collections
    If Not mcolVisits Is Nothing Then
        ' Explicitly terminate each object
        For Each oVisit In mcolVisits
            Call oVisit.Terminate
        Next
        Call CollectionRemoveAll(mcolVisits)
    End If
    Set mcolVisits = Nothing
    
    If Not mcolEForms Is Nothing Then
        For Each oEForm In mcolEForms
            Call oEForm.Terminate
        Next
        Call CollectionRemoveAll(mcolEForms)
    End If
    Set mcolEForms = Nothing
    
Exit Sub
ErrLabel:
'    Debug.Print "Error in Terminating Study"

End Sub
 
'-------------------------------------------------------------------
Public Property Get Arezzo_DM() As Arezzo_DM
'-------------------------------------------------------------------
' Return the arezzo object for this study.
' Returns nothing is there si not one.
'-------------------------------------------------------------------

    Set Arezzo_DM = moArezzo

End Property

'-------------------------------------------------------------------
Public Property Set Arezzo_DM(oArezzo As Arezzo_DM)
'-------------------------------------------------------------------
' Set the arezzo object for this study.
'-------------------------------------------------------------------

    Set moArezzo = Arezzo_DM

End Property

'--------------------------------------------
Public Sub LoadElements(oEForm As eFormRO)
'--------------------------------------------
' Load the elements for the given eForm.
' nb this is done automatically in StudySubject.LoadResponses
'--------------------------------------------

    On Error GoTo ErrorLabel
    
    Call oEForm.LoadElements(msConnection)
    
Exit Sub

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudyDefRO.LoadElements"
    Exit Sub
    
End Sub

'--------------------------------------------
Public Sub RemoveElements(oEForm As eFormRO)
'--------------------------------------------
' remove the elements for the given eForm.
' nb this is done automatically in StudySubject.RemoveResponses
'--------------------------------------------

    On Error GoTo ErrorLabel
    
    Call oEForm.RemoveElements
    
Exit Sub

ErrorLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "StudyDefRO.RemoveElements"
    Exit Sub
    
End Sub

'----------------------------------------------------------------------------------
Public Sub SetLocalNumberFormats(ByVal sDecimalPointChar As String, _
                                ByVal sThousandSeparatorChar As String)
'----------------------------------------------------------------------------------
' NCJ 12 May 03 - Set what decimal point and thousand separator characters to use
' Pass empty strings to use machine's own settings
'----------------------------------------------------------------------------------

    msDecimalPtChar = sDecimalPointChar
    msThousandSepChar = sThousandSeparatorChar

End Sub

'----------------------------------------------------------------------------------
Public Property Get LocalDecimalPoint() As String
'----------------------------------------------------------------------------------
' NCJ 13 May 03 - Return what decimal point we're currently using
' (as set up by SetLocalNumberFormats)
'----------------------------------------------------------------------------------

    ' Is there a Decimal Point defined?
    If msDecimalPtChar > "" Then
        LocalDecimalPoint = msDecimalPtChar
    Else
        ' Use machine's own local setting
        LocalDecimalPoint = RegionalDecimalPointChar
    End If
    
End Property

'----------------------------------------------------------------------------------
Public Property Get LocalThousandSeparator() As String
'----------------------------------------------------------------------------------
' NCJ 13 May 03 - Return what thousand separator we're currently using
' (as set up by SetLocalNumberFormats)
'----------------------------------------------------------------------------------

    ' Is there a Thousand Separator defined?
    If msThousandSepChar > "" Then
        LocalThousandSeparator = msThousandSepChar
    Else
        ' Use machine's own local setting
        LocalThousandSeparator = RegionalThousandSeparatorChar
    End If
    
End Property

