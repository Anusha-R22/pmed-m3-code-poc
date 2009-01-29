VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStudyDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       clsStudyDefinition.cls
'   Author:     Paul Norris 17/08/99
'   Purpose:    Class to wrap up a information pertaining to a StudyDefinition
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   1   Paul Norris     27/08/99
'       Updated Load() to prevent TrialSubjectLabel rule breaking when Null is read from db
'   2   Paul Norris     30/08/99
'       Added reference to gbSingleUseDataItems in Save()
'   PN  28/09/99    Added TrialSubjectLabelValid(), StandardTimeValid() and StandardDateValid()
'                   events to inform the display that it is invalid
'   PN  30/09/99    Amended Class_Initialize() to specify full SQL statment
'   Mo Morris   8/11/99
'   DAO to ADO conversion
'   WillC 15/11/99  Added msUsername to show UserId on form
'   NCJ 13 Dec 99 - ClinicalTrialIds to Long
'   NCJ 22 Dec 99, SR 2045 - Validate time & date formats separately
'   NCJ 3 Feb 00 - Changed date format validation to use new CLM routine
'   TA 15/09/2000 - additional properties for handling CTC Scheme, Gender and DOB expressions
'   NCJ 2 Apr 01 - Bug fix to Save (error when deleting study description)
'   Mo Morris 30/8/01 Db Audit (UserId to UserName throughout)
'   ZA/NCJ 3-8 Oct 02 - Added ReasonsForOverrule
'   NCJ 7 Dec 05 - Handle new eDateFormatType enumeration values
'----------------------------------------------------------------------------------------'
Option Explicit

Public Event TrialSubjectLabelValid(bIsValid As Boolean)
'TA 15/9/2000 Valid Gender and DOB events
Public Event SubjectGenderValid(bIsValid As Boolean)
Public Event SubjectDOBValid(bIsValid As Boolean)

Public Event StandardTimeValid(bIsValid As Boolean)
Public Event StandardDateValid(bIsValid As Boolean)
Public Event AddedNewReason(enReasonType As eReasonType)
Public Event IsValid(bValid As Boolean)
Public Event HasChanges()

' object to handle the validation
Private WithEvents moBrokenRules As clsBrokenRules
Attribute moBrokenRules.VB_VarHelpID = -1
' object to contain all valid phases
Private moPhases As clsTextList
'object to handle all valid trial types
Private moTrialTypes As clsTextList
'object to handle all valid status values
Private moStatuses As clsTextList

' object to handle all reason for change values
Private WithEvents moReasonsForChange As clsReasonsForChange
Attribute moReasonsForChange.VB_VarHelpID = -1

' object to handle all reason for overrule values
Private WithEvents moReasonsForOverrule As clsReasonsForChange
Attribute moReasonsForOverrule.VB_VarHelpID = -1

' properties
Private msTrialSubjectLabel As String
Private mnPhase As Integer
Private mnTrialType As Integer
Private mnStatus As Integer
Private msLocalTrialSubjectLabel As String
Private mlExpectedRecruitment As Long
Private msKeywords As String
Private msTrialDescription As String
Private msStandardDateFormat As String
Private msStandardTimeFormat As String
Private miSingleUseDataItems As Integer
'Mo Morris 30/8/01 Db Audit (UserId to UserName)
Private msUserName As String
'TA 15/9/2000
Private mvCTCSchemeCode As Variant

Private msSubjectGender As String
Private msSubjectDOB As String

' internal data management properties
Private mbIsDirty As Boolean

'---------------------------------------------------------------------
Public Property Get CTCSchemeCode() As String
'---------------------------------------------------------------------
' convert null to "" for combo box
'---------------------------------------------------------------------

    If VarType(mvCTCSchemeCode) = vbNull Then
        CTCSchemeCode = ""
    Else
        CTCSchemeCode = mvCTCSchemeCode
    End If

End Property

'---------------------------------------------------------------------
Public Property Let CTCSchemeCode(sCTCSchemeCode As String)
'---------------------------------------------------------------------
' convert "" in combo box to null
'---------------------------------------------------------------------

    If sCTCSchemeCode = "" Then
        If VarType(mvCTCSchemeCode) <> vbNull Then
            mvCTCSchemeCode = Null
            IsDirty = True
        End If
    Else
        If mvCTCSchemeCode <> sCTCSchemeCode Or VarType(mvCTCSchemeCode) = vbNull Then
            mvCTCSchemeCode = sCTCSchemeCode
            IsDirty = True
        End If
    End If
    
End Property

'---------------------------------------------------------------------
Public Property Get ReasonsForOverrule() As clsReasonsForChange
'---------------------------------------------------------------------
' NCJ/ZA 7 Oct 02
' The collection of RFOs
'---------------------------------------------------------------------

    Set ReasonsForOverrule = moReasonsForOverrule

End Property

'---------------------------------------------------------------------
Public Property Get ReasonsForChange() As clsReasonsForChange
'---------------------------------------------------------------------
' The collection of RFCs
'---------------------------------------------------------------------

    Set ReasonsForChange = moReasonsForChange

End Property

'---------------------------------------------------------------------
Public Property Get SingleUseDataItems() As Integer
'---------------------------------------------------------------------

    SingleUseDataItems = miSingleUseDataItems

End Property

'---------------------------------------------------------------------
Public Property Let SingleUseDataItems(iValue As Integer)
'---------------------------------------------------------------------
    
    ' raise the invalid property value error
    If iValue <> 0 And iValue <> 1 Then Err.Raise 380
    
    If iValue <> miSingleUseDataItems Then
        IsDirty = True
        miSingleUseDataItems = iValue
    End If

End Property

'---------------------------------------------------------------------
Private Property Let IsDirty(bIsDirty As Boolean)
'---------------------------------------------------------------------

    If bIsDirty <> mbIsDirty Then
        mbIsDirty = bIsDirty
        RaiseEvent HasChanges
    End If

End Property

'---------------------------------------------------------------------
Public Property Get UserName() As String
'---------------------------------------------------------------------

    UserName = msUserName
    
End Property

'---------------------------------------------------------------------
Public Property Get SubjectGender() As String
'---------------------------------------------------------------------

    SubjectGender = msSubjectGender

End Property

'---------------------------------------------------------------------
Public Property Let SubjectGender(sGender As String)
'---------------------------------------------------------------------

    If Not gblnValidString(sGender, valOnlySingleQuotes) Then
        ' raise the invalid property value error
        Err.Raise 380
        
    End If
    
    If sGender <> vbNullString Then
        If gclmGuideline.IsValidExpression(sGender) Then
            moBrokenRules.RuleBroken "SubjectGender", False
            
            ' PN 28/09/99 raise this event to inform the display that it is invalid
            RaiseEvent SubjectGenderValid(True)
            
        Else
            ' PN 28/09/99 raise this event to inform the display that it is invalid
            moBrokenRules.RuleBroken "SubjectGender", True
            RaiseEvent SubjectGenderValid(False)
            
        End If
    End If
    
    If sGender <> msSubjectGender Then
        IsDirty = True
        msSubjectGender = sGender
    End If
    
End Property

'---------------------------------------------------------------------
Public Property Get SubjectDOB() As String
'---------------------------------------------------------------------

    SubjectDOB = msSubjectDOB

End Property

'---------------------------------------------------------------------
Public Property Let SubjectDOB(sDOB As String)
'---------------------------------------------------------------------

    If Not gblnValidString(sDOB, valOnlySingleQuotes) Then
        ' raise the invalid property value error
        Err.Raise 380
        
    End If
    
    If sDOB <> vbNullString Then
        If gclmGuideline.IsValidExpression(sDOB) Then
            moBrokenRules.RuleBroken "SubjectDOB", False
            
            ' PN 28/09/99 raise this event to inform the display that it is invalid
            RaiseEvent SubjectDOBValid(True)
            
        Else
            ' PN 28/09/99 raise this event to inform the display that it is invalid
            moBrokenRules.RuleBroken "SubjectDOB", True
            RaiseEvent SubjectDOBValid(False)
            
        End If
    End If
    
    If sDOB <> msSubjectDOB Then
        IsDirty = True
        msSubjectDOB = sDOB
    End If
    
End Property

'---------------------------------------------------------------------
Public Property Get TrialSubjectLabel() As String
'---------------------------------------------------------------------

    TrialSubjectLabel = msTrialSubjectLabel

End Property

'---------------------------------------------------------------------
Public Property Let TrialSubjectLabel(sLabel As String)
'---------------------------------------------------------------------

    If Not gblnValidString(sLabel, valOnlySingleQuotes) Then
        ' raise the invalid property value error
        Err.Raise 380
        
    End If
    
    If sLabel <> vbNullString Then
        ' Use Arezzo expression checker - NCJ 13/8/99
        ' ElseIf Not PfECheckExpression(txtTrialSubjectLabel.Text, msPrologOutput) Then
        If gclmGuideline.IsValidExpression(sLabel) Then
            moBrokenRules.RuleBroken "TrialSubjectLabel", False
            
            ' PN 28/09/99 raise this event to inform the display that it is invalid
            RaiseEvent TrialSubjectLabelValid(True)
            
        Else
            ' PN 28/09/99 raise this event to inform the display that it is invalid
            moBrokenRules.RuleBroken "TrialSubjectLabel", True
            RaiseEvent TrialSubjectLabelValid(False)
            
        End If
    End If
    
    If sLabel <> msTrialSubjectLabel Then
        IsDirty = True
        msTrialSubjectLabel = sLabel
    End If

End Property
'---------------------------------------------------------------------
Public Property Get Phase() As String
'---------------------------------------------------------------------

    ' return the string for the phase key
    Phase = moPhases.Item(mnPhase)

End Property

'---------------------------------------------------------------------
Public Property Let Phase(sPhase As String)
'---------------------------------------------------------------------

    If moPhases.Key(sPhase) <> mnPhase Then
        IsDirty = True
        
        ' cache the key for the phase string
        mnPhase = moPhases.Key(sPhase)
    
    End If

End Property

'---------------------------------------------------------------------
Public Property Get Status() As String
'---------------------------------------------------------------------

    ' return the string for the Status key
    Status = moStatuses.Item(mnStatus)
    
End Property

'---------------------------------------------------------------------
Public Property Let Status(sStatus As String)
'---------------------------------------------------------------------

    If moStatuses.Key(sStatus) <> mnStatus Then
        IsDirty = True
        
        ' cache the key for the phase string
        mnStatus = moStatuses.Key(sStatus)
    
    End If

End Property

'---------------------------------------------------------------------
Public Property Get TrialType() As String
'---------------------------------------------------------------------

    ' return the string for the Status key
    TrialType = moTrialTypes.Item(mnTrialType)
    
End Property

'---------------------------------------------------------------------
Public Property Let TrialType(sTrialType As String)
'---------------------------------------------------------------------

    If moTrialTypes.Key(sTrialType) <> mnTrialType Then
        IsDirty = True
        ' cache the key for the phase string
        mnTrialType = moTrialTypes.Key(sTrialType)
    
    End If

End Property

'---------------------------------------------------------------------
Public Property Get Phases() As clsTextList
'---------------------------------------------------------------------

    Set Phases = moPhases

End Property

'---------------------------------------------------------------------
Public Property Get TrialTypes() As clsTextList
'---------------------------------------------------------------------

    Set TrialTypes = moTrialTypes

End Property

'---------------------------------------------------------------------
Public Property Get Statuses() As clsTextList
'---------------------------------------------------------------------

    Set Statuses = moStatuses

End Property

'---------------------------------------------------------------------
Public Property Get LocalTrialSubjectLabel() As String
'---------------------------------------------------------------------

    LocalTrialSubjectLabel = msLocalTrialSubjectLabel

End Property

'---------------------------------------------------------------------
Public Property Let LocalTrialSubjectLabel(sLabel As String)
'---------------------------------------------------------------------

    If sLabel <> msLocalTrialSubjectLabel Then
        IsDirty = True
        msLocalTrialSubjectLabel = sLabel
    End If

End Property

'---------------------------------------------------------------------
Public Property Get ExpectedRecruitment() As String
'---------------------------------------------------------------------

    ExpectedRecruitment = mlExpectedRecruitment

End Property

'---------------------------------------------------------------------
Public Property Let ExpectedRecruitment(sCount As String)
'---------------------------------------------------------------------

    ' raise the invalid property value error
    If Not gblnValidString(sCount, valNumeric) Then Err.Raise 380
    If sCount < 0 Then Err.Raise 380
    
    If sCount <> mlExpectedRecruitment Then
        IsDirty = True
        mlExpectedRecruitment = sCount
    End If

End Property

'---------------------------------------------------------------------
Public Property Get Keywords() As String
'---------------------------------------------------------------------

    Keywords = msKeywords

End Property

'---------------------------------------------------------------------
Public Property Let Keywords(sKeywords As String)
'---------------------------------------------------------------------

    ' raise the invalid property value error
    ' WillC 20/2/2000 SR2332 allow some extra characters in the Ketwords field
    ' replaced like below and check only for quotes pipes handled further on in gblnValidString
    'If Not gblnValidString(sKeywords, valAlpha + valNumeric + valComma + valSpace) Then
    
    If Not gblnValidString(sKeywords, valOnlySingleQuotes) Then
        Err.Raise 380
    End If
    
    If sKeywords <> msKeywords Then
        IsDirty = True
        msKeywords = sKeywords
    End If

End Property

'---------------------------------------------------------------------
Public Property Get TrialDescription() As String
'---------------------------------------------------------------------

    TrialDescription = msTrialDescription

End Property

'---------------------------------------------------------------------
Public Property Let TrialDescription(sDescription As String)
'---------------------------------------------------------------------
    
    ' raise the invalid property value error
    If Not gblnValidString(sDescription, valOnlySingleQuotes) Then Err.Raise 380
    If sDescription <> msTrialDescription Then
        IsDirty = True
        msTrialDescription = sDescription
    End If

End Property

'---------------------------------------------------------------------
Public Property Get StandardDateFormat() As String
'---------------------------------------------------------------------

    StandardDateFormat = msStandardDateFormat

End Property

'---------------------------------------------------------------------
Public Property Let StandardDateFormat(sDateFormat As String)
'---------------------------------------------------------------------
' NCJ 3 Feb 00 - Changed date format validation
' NCJ 7 Dec 05 - Handle new date format values
'---------------------------------------------------------------------

    ' raise the invalid property value error
    If Not gblnValidString(sDateFormat, valOnlySingleQuotes) Then Err.Raise 380
    
    If sDateFormat <> msStandardDateFormat Then
        IsDirty = True
        
        ' maintain the broken rules
        ' NCJ 22 Dec 99, SR 2045 USe ValidateDateOnlyFormat
        ' NCJ 3 Feb 00 - Use new CLM call to validate
        ' NCJ 7 Dec 05 - New date format values
        Select Case ValidateDateFormatString(sDateFormat)
        Case eDateFormatType.dftDMY, eDateFormatType.dftMDY, eDateFormatType.dftYMD
            ' Can be a full date
            moBrokenRules.RuleBroken "StandardDateFormat", False
            ' PN 28/09/99 raise this event to inform the display that it is invalid
            RaiseEvent StandardDateValid(True)
            
        Case Else
            moBrokenRules.RuleBroken "StandardDateFormat", True
            ' PN 28/09/99 raise this event to inform the display that it is invalid
            RaiseEvent StandardDateValid(False)
            
        End Select
        msStandardDateFormat = sDateFormat
    End If

End Property

'---------------------------------------------------------------------
Public Property Get StandardTimeFormat() As String
'---------------------------------------------------------------------
    
    StandardTimeFormat = msStandardTimeFormat

End Property

'---------------------------------------------------------------------
Public Property Let StandardTimeFormat(sTimeFormat As String)
'---------------------------------------------------------------------
' NCJ 3 Feb 00 - Changed date format validation
' NCJ 7 Dec 05 - New eDateFormatType enumeration values
'---------------------------------------------------------------------

    ' raise the invalid property value error
    If Not gblnValidString(sTimeFormat, valOnlySingleQuotes) Then Err.Raise 380
    ' NCJ 3/2/00 - We now allow different separators
'    If Not gblnValidString(sTimeFormat, valAlpha + valDateSeperators) Then Err.Raise 380
    
    If sTimeFormat <> msStandardTimeFormat Then
        IsDirty = True

        ' maintain the broken rules
        ' NCJ 3 Feb 00 - Use new CLM call to validate
        Select Case ValidateDateFormatString(sTimeFormat)
        Case eDateFormatType.dftT
            ' Can be a time-only
            moBrokenRules.RuleBroken "StandardTimeFormat", False
            ' PN 28/09/99 raise this event to inform the display that it is invalid
            RaiseEvent StandardTimeValid(True)
            
        Case Else
            moBrokenRules.RuleBroken "StandardTimeFormat", True
            ' PN 28/09/99 raise this event to inform the display that it is invalid
            RaiseEvent StandardTimeValid(False)
            
        End Select
        msStandardTimeFormat = sTimeFormat
    End If
    
End Property

'---------------------------------------------------------------------
Private Sub Class_Initialize()
'---------------------------------------------------------------------
' NCJ 2 Apr 01 - Explicitly initialise all string variables
'   (Prevents quirky errors when writing fields to DB in Save)
'---------------------------------------------------------------------

    Set moBrokenRules = New clsBrokenRules
    
    Set moReasonsForChange = New clsReasonsForChange
    moReasonsForChange.ReasonsType = ReasonForChange
    
    Set moReasonsForOverrule = New clsReasonsForChange
    moReasonsForOverrule.ReasonsType = ReasonForOverrule
    
    ' load the dependent data
    Set moPhases = New clsTextList
    Set moTrialTypes = New clsTextList
    Set moStatuses = New clsTextList
    
    ' NCJ 2 Apr 01 - Explicitly initialise all string variables
    msTrialSubjectLabel = ""
    msLocalTrialSubjectLabel = ""
    msKeywords = ""
    msTrialDescription = ""
    msStandardDateFormat = ""
    msStandardTimeFormat = ""
    msUserName = ""
    msSubjectGender = ""
    msSubjectDOB = ""

    ' PN 30/09/99 - change to specify full SQL statement for ODBC compliance with ADO
    moPhases.Load "SELECT * FROM TrialPhase ORDER BY PhaseId", True
    moTrialTypes.Load "SELECT * FROM TrialType", True
    moStatuses.Load "SELECT * FROM TrialStatus", True

End Sub

'---------------------------------------------------------------------
Private Sub moBrokenRules_BrokenRule()
'---------------------------------------------------------------------

    ' notify the container that the data is not all valid
    RaiseEvent IsValid(False)
    
End Sub

'---------------------------------------------------------------------
Private Sub moBrokenRules_NoBrokenRules()
'---------------------------------------------------------------------

    ' notify the container that the data is all valid
    RaiseEvent IsValid(True)
    
End Sub

'---------------------------------------------------------------------
Public Property Get IsValid() As Boolean
'---------------------------------------------------------------------

    IsValid = (moBrokenRules.Count = 0)

End Property

' this function will determine if the single use data item rule is broken
' for a study definition
'---------------------------------------------------------------------
Public Function IsSingleUseRuleBroken(lClinicalTrialId As Long, nVersionId As Integer) As Boolean
'---------------------------------------------------------------------
' this function will determine if the single use data item rule is broken
' for a study definition
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsSingleUse As ADODB.Recordset
    On Error GoTo ErrHandler
    
    sSQL = "SELECT DataItemID, COUNT(CRFElementID) From CRFElement" _
         & " WHERE ClinicalTrialId = " & lClinicalTrialId _
         & " AND VersionID= " & nVersionId _
         & " AND DataItemID > 0 " _
         & " GROUP BY DataItemID HAVING COUNT(CRFElementId) > 1"
    Set rsSingleUse = New ADODB.Recordset
    rsSingleUse.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    IsSingleUseRuleBroken = rsSingleUse.RecordCount > 0
    rsSingleUse.Close
    Set rsSingleUse = Nothing
        
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsStudyDefinition.IsSingleUseRuleBroken"

End Function

'---------------------------------------------------------------------
Public Sub Load(lClinicalTrialId As Long, nVersionId As Integer)
'---------------------------------------------------------------------
Dim rsTrialDetails As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    ' read clinical trial table
    Set rsTrialDetails = New ADODB.Recordset
    Set rsTrialDetails = gdsTrialDetails(lClinicalTrialId)
    
    With rsTrialDetails
        mnTrialType = .Fields("TrialTypeID")
        mnPhase = .Fields("PhaseId")
        mnStatus = .Fields("StatusId")
        'TA 13/12/2000: removenull required for oracle
        TrialDescription = RemoveNull(.Fields("ClinicalTrialDescription"))
        If Not IsNull(.Fields("Keywords")) Then
            Keywords = .Fields("Keywords")
        End If
        ExpectedRecruitment = .Fields("ExpectedRecruitment")
        
        .Close
    End With
    Set rsTrialDetails = Nothing
    
    ' read study definition table
    Set rsTrialDetails = New ADODB.Recordset
    Set rsTrialDetails = gdsStudyDefinition(lClinicalTrialId, nVersionId)
    With rsTrialDetails
'   ATN 16/12/99
'   Need to use rtrim rather than isnull becuase SQL Server will return a single space
'   rather than an empty string
        If RTrim(.Fields("TrialSubjectLabel")) > "" Then
            TrialSubjectLabel = .Fields("TrialSubjectLabel")
        End If
        LocalTrialSubjectLabel = .Fields("LocalTrialSubjectLabel")
        If IsNull(.Fields("StandardDateFormat")) Then
            moBrokenRules.RuleBroken "StandardDateFormat", True
        Else
            StandardDateFormat = .Fields("StandardDateFormat")
        End If
        If IsNull(.Fields("StandardTimeFormat")) Then
            moBrokenRules.RuleBroken "StandardTimeFormat", True
        Else
            StandardTimeFormat = .Fields("StandardTimeFormat")
        End If
        SingleUseDataItems = .Fields("SingleUseDataItems")
        'Mo Morris 30/8/01 Db Audit (UserId to UserName)
        msUserName = .Fields("UserName")
        'additional columns for CTC and NR
        SubjectGender = RemoveNull(.Fields("GenderExpr"))
        SubjectDOB = RemoveNull(.Fields("DOBExpr"))
        mvCTCSchemeCode = .Fields("CTCSchemeCode")
        
        .Close
    End With
    Set rsTrialDetails = Nothing
    moReasonsForChange.Load lClinicalTrialId, nVersionId
    moReasonsForOverrule.Load lClinicalTrialId, nVersionId
    
    mbIsDirty = False
        
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsStudyDefinition.Load"

End Sub

'---------------------------------------------------------------------
Public Sub Save(lClinicalTrialId As Long, nVersionId As Integer)
'---------------------------------------------------------------------
' NCJ 2 Apr 01 - Don't use WITH when saving to recordset because of
' an obscure bug in Jet Engine (see MSDN Q228935)
'---------------------------------------------------------------------

    Dim rsStudyRecord As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    If mbIsDirty Then
        ' already exists
        
        ' update the study definition table
        Set rsStudyRecord = New ADODB.Recordset
        Set rsStudyRecord = gdsStudyDefinition(lClinicalTrialId, nVersionId)
        ' NCJ 2/4/01 - Don't use WITH
'        With rsStudyRecord
            '.Edit
            rsStudyRecord.Fields("LocalTrialSubjectLabel") = msLocalTrialSubjectLabel
            rsStudyRecord.Fields("TrialSubjectLabel") = msTrialSubjectLabel
            rsStudyRecord.Fields("LocalTrialSubjectLabel") = msLocalTrialSubjectLabel
            'gender and DOB expression
            If msSubjectGender = "" Then
                rsStudyRecord.Fields("GenderExpr") = Null
            Else
                rsStudyRecord.Fields("GenderExpr") = msSubjectGender
            End If
            If msSubjectDOB = "" Then
                rsStudyRecord.Fields("DOBExpr") = Null
            Else
                rsStudyRecord.Fields("DOBExpr") = msSubjectDOB
            End If
            rsStudyRecord.Fields("CTCSchemeCode") = mvCTCSchemeCode
            
            rsStudyRecord.Fields("StandardDateFormat") = msStandardDateFormat
            rsStudyRecord.Fields("StandardTimeFormat") = msStandardTimeFormat
            rsStudyRecord.Fields("SingleUseDataItems") = miSingleUseDataItems
            
            rsStudyRecord.Update
            rsStudyRecord.Close
            

'        End With
        Set rsStudyRecord = Nothing
        
        ' update the clinical trial table
        Set rsStudyRecord = New ADODB.Recordset
        Set rsStudyRecord = gdsTrialDetails(lClinicalTrialId)
        ' NCJ 2/4/01 - Don't use WITH
'        With rsStudyRecord
            '.Edit
            rsStudyRecord.Fields("ExpectedRecruitment") = mlExpectedRecruitment
            rsStudyRecord.Fields("ClinicalTrialDescription") = msTrialDescription
            rsStudyRecord.Fields("Keywords") = msKeywords
            rsStudyRecord.Fields("TrialTypeID") = mnTrialType
            rsStudyRecord.Fields("PhaseId") = mnPhase
            rsStudyRecord.Fields("StatusId") = mnStatus
            rsStudyRecord.Update
            rsStudyRecord.Close
'        End With
        Set rsStudyRecord = Nothing
        
        ' PN change 2 30/08/99
        ' also save the registry entry for SingleUseDataItems
        ' this is done so that if the value is changed in the database
        ' any other form that requires the value will read the global value
        If miSingleUseDataItems = 1 Then
            gbSingleUseDataItems = True
        Else
            gbSingleUseDataItems = False
            
        End If
        mbIsDirty = False
        
    End If
    If moReasonsForChange.IsDirty Then
        moReasonsForChange.Save lClinicalTrialId, nVersionId
    End If
    If moReasonsForOverrule.IsDirty Then
        moReasonsForOverrule.Save lClinicalTrialId, nVersionId
    End If
        
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsStudyDefinition.Save"

End Sub

'---------------------------------------------------------------------
Public Property Get IsDirty() As Boolean
'---------------------------------------------------------------------
Dim bIsDirty As Boolean
    
    If Not mbIsDirty Then
        ' only check if the ReasonsForChange colection has changes if the
        ' internal properties have not changed
        If moReasonsForChange.IsDirty Then
            bIsDirty = True
        ElseIf moReasonsForOverrule.IsDirty Then
            bIsDirty = True
        End If
    Else
        bIsDirty = mbIsDirty
    End If

    IsDirty = bIsDirty

End Property

'---------------------------------------------------------------------
Public Sub Discard()
'---------------------------------------------------------------------

    Set moPhases = Nothing
    Set moTrialTypes = Nothing
    Set moStatuses = Nothing
    Set moBrokenRules = Nothing
    If Not moReasonsForChange Is Nothing Then
        moReasonsForChange.Discard
    End If
    If Not moReasonsForOverrule Is Nothing Then
        moReasonsForOverrule.Discard
    End If
    Set moReasonsForChange = Nothing
    Set moReasonsForOverrule = Nothing
    
End Sub

'---------------------------------------------------------------------
Private Sub moReasonsForChange_AddedNewItem()
'---------------------------------------------------------------------

    RaiseEvent AddedNewReason(ReasonForChange)

End Sub

'---------------------------------------------------------------------
Private Sub moReasonsForOverrule_AddedNewItem()
'---------------------------------------------------------------------

    RaiseEvent AddedNewReason(ReasonForOverrule)

End Sub

