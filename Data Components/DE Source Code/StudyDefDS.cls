VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyDefDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------
' File: StudyDefDS.cls
' Copyright InferMed Ltd 2002 All Rights Reserved
' Author: Nicky Johns, InferMed, May 2001
' Purpose: Data Services Study Defn class for MACRO 3.0
'----------------------------------------------------

'----------------------------------------------------
' REVISIONS
'----------------------------------------------------
' NCJ 18-29 May 01 - Initial development
' TA  24/07/2001 - Load of visits, eforms and VisitEforms now all done in this class
' ZA  07/08/01 - Added Fonts properties
' ZA 08/08/01 - Loading of Reason for change (RFCs)
' NCJ 20/9/01 - Change Authorisation from "0" to empty string
' NCJ 20 Sep 01 - Added in remaining fields for CRFPage table,
'                and RR fields for StudyDef
' TA 21 Sep 01 - All field and table names now are constants
'                   - no schema name within subs any more
' NCJ 27 Sept 01 - Added Phase, Status and TrialType
'   TA 01/10/2001 - Error handling added/classes combined
' TA 12/10/2001: DataItemId is read from CRFElement table so that it is 0 for graphical elements
' MLM 21/01/02: 2.2.8 Current Build Buglist, Bug 2: Added DataItemId to msCAT_ORDERCOLS
' MACRO 3.0
' NCJ 5 July 2002 - Merged in RQG stuff
' ZA 17/17/2002 - Added font properties for question caption and comment
' TA 20/08/2002: CaptionFont properties now converted from null to zeros
' MLM 28/08/02: Include the EFormUse field in LoadVEFs and ElementUse in LoadElements
'ZA 20/09/2002 - Added DisplayLength column to CRFElement
'   NCJ 8 Oct 02 - Added ReasonsForOverrule, and loading of Date Elements with Study
' NCJ 4 Nov 02 - Added eFormWidth property
' NCJ 8 Nov 02 - Added CRFElement Hotlink property
' NCJ 21 May 03 - Added Repeating property for visits
' NCJ 6 Jul 05 - Bug 2517 - Added DataItemId to msVAL_ORDERCOLS
' ic 14/07/2005 added clinical coding
'----------------------------------------------------

Option Explicit

Private Const msDELIMITER = "|"

' The table names
Private Const msSTUDY_DEF_TABLE = "StudyDefinition"
Private Const msSTUDY_TABLE = "ClinicalTrial"
Private Const msPROT_TABLE = "Protocols"

' The "key" columns, Study and Version
Private Const msKEY_COLS = "ClinicalTrialId|VersionID"

Private Const msSTUDYID = ".ClinicalTrialId"
Private Const msVERSION = ".VersionId"

' The study def data cols
Private Const msSTUDY_DEF_COLS = "TrialSubjectLabel|LocalTrialSubjectLabel|" _
                        & "StandardDateFormat|StandardTimeFormat|" _
                        & "CTCSchemeCode|DOBExpr|GenderExpr|DefaultFontName|" _
                        & "DefaultFontBold|DefaultFontItalic|DefaultFontSize|" _
                        & "DefaultFontColour|DefaultCRFPageColour|" _
                        & "RRServerType|RRHTTPAddress|RRUserName|RRPassword|RRProxyServer"
                        
                        
Private Const msSTUDY_COLS = "ClinicalTrialName|ClinicalTrialDescription|" _
                            & "PhaseId|StatusId|TrialTypeId"
                            
' The Protocols cols
Private Const msPROT_COLS = "ArezzoFile"

'cols for joining study def and protocols table
Private Const msPROT_STUDY_JOINCOL = msPROT_TABLE & ".FileName"
Private Const msSTUDY_PROT_JOINCOL = msSTUDY_TABLE & ".ClinicalTrialName"

'for visits
' The Visit table
Private Const msVISIT_TABLE = "StudyVisit"

' The columns we want
' NCJ 21 May 03 - Added Repeating
Private Const msVISIT_COLS = "VisitID|VisitCode|VisitName|" _
                & "VisitOrder|VisitDateLabel|VisitDatePrompt|" _
                & "VisitBackgroundColour|Repeating"
                     
                     
Private Const msVISIT_ORDERCOL = "VisitOrder"
'for Eforms
' The eForm table
Private Const msEFORM_TABLE = "CRFPage"

' The columns we want
Private Const msEFORM_COLS = "CRFPageID|CRFPageCode|CRFTitle|" _
                & "CRFPageOrder|CRFPageLabel|CRFPageDateLabel|" _
                & "DisplayNumbers|HideIfInactive|BackgroundColour|" _
                & "eFormDatePrompt|SequentialEntry|LocalCRFPageLabel|eFormWidth"
                
Private Const msEFORM_ORDERCOL = "CRFPageOrder"
                     
                     
'for visit eforms
' The visit eForm table
Private Const msVEF_TABLE = "StudyVisitCRFPage"

' The columns we want
'MLM 28/08/02: Added EFormUse
Private Const msVEF_COLS = "VisitID|CRFPageID|Repeating|EFormUse"

'The reason for change table
Private Const msRFCs_TABLE = "ReasonForChange"

'The columns we want
' NCJ 8 Oct 02 - Added ReasonType
Private Const msRFCs_COLS = "ReasonForChange|ReasonForChangeId|ReasonType"

'order this table by:
Private Const msRFCs_ORDERCOL = "ReasonForChangeId"

' NCJ 8 Oct 02 - Types of reason
Private Const mnRFC_TYPE = 0
Private Const mnRFO_TYPE = 0

'this used to be in its own class

' The Question/Element table
Private Const msQUESTION_TABLE = "DataItem"
Private Const msELEMENT_TABLE = "CRFElement"

' The columns we want
Private Const msQUESTION_COLS = "DataItemCode|DataItemName|DataType|" _
        & "DataItemFormat|UnitOfMeasurement|DataItemLength|" _
        & "Derivation|DataItemHelpText|DataItemCase|ClinicalTestCode"
        
'ic 14/07/2005 added clinical coding
Private Const msCCQUESTION_COLS = "|DictionaryId"
Private Const mCCSwitch As String = "CLINICALCODING"

               
' NCJ 5 Nov 01 - Added extra group cols
'ZA 19/07/2001 - Added new columns for caption font and colour
'MLM 28/08/02: Added ElementUse
'ZA 20/09/2002 - Added DisplayLength column
' NCJ 8 Oct 02 - Added CRFElement.CRFPageId
' NCJ 8 Nov 02 - Added Hotlink
Private Const msELEMENT_COLS = "CRFElement.DataItemId|CRFElementId|ControlType|Caption|FieldOrder|" _
                & "SkipCondition|Hidden|LocalFlag|Optional|Mandatory|" _
                & "RequireComment|RoleCode|ClinicalTestDateExpr|FontName|" _
                & "FontSize|FontBold|FontItalic|FontColour|CaptionX|CaptionY|" _
                & "X|Y|QGroupID|OwnerQGroupID|ShowStatusFlag|QGroupFieldOrder|" _
                & "CaptionFontName|CaptionFontBold|CaptionFontItalic|CaptionFontSize|CaptionFontColour|ElementUse|" _
                & "DisplayLength|Hotlink|CRFElement.CRFPageId"

' order cols for CRFElement table
Private Const msELEMENT_ORDERCOLS = "FieldOrder|QGroupFieldOrder"
Private Const msDITEMID = ".DataItemId"


' NCJ 2 Nov 01
Private Const msQGROUP_TABLE = "QGroup"
Private Const msEFORM_QGROUP_TABLE = "EFormQGroup"

Private Const msQGROUP_COLS = "QGroup.QGroupID|QGroupCode|QGroupName|DisplayType"
Private Const msEFORM_QGROUP_COLS = "CRFPageID|Border|DisplayRows|InitialRows|" _
                                    & "MinRepeats|MaxRepeats"

Private Const msQGROUPID = ".QGroupID"

'this column appears in ValueData, DataItemValidation, CRFElement, eFormQGroup
Private Const msCRFPAGEID_COL = "CRFPageId"

Private Const msELEMENT_USE_COL = "ElementUse"

' for categories
' The table names
Private Const msCAT_TABLE = "ValueData"

' The columns we want
Private Const msCAT_COLS = "ValueData.DataItemId|" _
                 & "ValueCode|ItemValue|Active|ValueOrder"

'order cols
'MLM 21/01/02: Added DataItemId here. This ensures categories for the same Q stay grouped when the
'   FieldOrders are the same, which occurs for derived Qs once the Qs on that form have been reordered.
Private Const msCAT_ORDERCOLS = "FieldOrder|QGroupFieldOrder|ValueData.DataItemId|ValueOrder"

'for Validations

' The table names
Private Const msVAL_TABLE = "DataItemValidation"
Private Const msVALTYPE_TABLE = "ValidationType"

' The columns we want
Private Const msVAL_COLS = "DataItemValidation.DataItemId|" _
                 & "ValidationID|DataItemValidation|ValidationMessage|" _
                 & "ValidationActionID"
                 
'order cols for DataITemValidation
' NCJ 6 Jul 2005 - Bug 2517 - Added DataItemId to ordering
Private Const msVAL_ORDERCOLS = "FieldOrder|QGroupFieldOrder|DataItemValidation.DataItemId|ValidationID"

Private Const msVAL_TYPEID = ".ValidationTypeId"

'--------------------------------------------------------------------
Public Function Load(ByVal sConnection As String, _
            ByVal lStudyId As Long, _
            ByVal nVersion As Integer, ByVal bIncludeArezzoFile As Boolean) As String
'--------------------------------------------------------------------
' Load the given study subject and return info as (property bag) string
' sConnection is the DB connection string.
' If bInculdeArezzoFile is false then an empty string is retuned for the ArezzoFile
' Empty string now returned if no study for studyid
'revisions
'ic 27/10/2005 added connection to eform load call so we can check for clinical coding upgrade
'--------------------------------------------------------------------
Dim oStudyState As StudyDefState
Dim oQS As QueryServer
Dim oSD As QueryDef
Dim vStudy As Variant
Dim colRFCs As Collection
Dim colRFOs As Collection

    On Error GoTo ErrorLabel
    
    ' Create the Query Server and Select Defn objects
    ' (Query server will be kept open for loading Visits and eForms too)
    Set oQS = New QueryServer
    Set oSD = New QueryDef

    ' Set up the Tables, the Columns, the Joins and the Filters
    oSD.QueryTables.Add msSTUDY_DEF_TABLE
    oSD.QueryTables.Add msSTUDY_TABLE, , qdjtInner, msSTUDY_TABLE & msSTUDYID, msSTUDY_DEF_TABLE & msSTUDYID
    
    oSD.QueryFields.Add Split(msSTUDY_DEF_COLS, msDELIMITER)
    oSD.QueryFields.Add Split(msSTUDY_COLS, msDELIMITER)
    
    If bIncludeArezzoFile Then
        'arezzo file wanted as well
        oSD.QueryTables.Add msPROT_TABLE, , qdjtInner, msPROT_STUDY_JOINCOL, msSTUDY_PROT_JOINCOL
        oSD.QueryFields.Add Split(msPROT_COLS, msDELIMITER)
    End If
    
    oSD.QueryFilters.Add msSTUDY_DEF_TABLE & msSTUDYID, "=", lStudyId
    oSD.QueryFilters.Add msSTUDY_DEF_TABLE & msVERSION, "=", nVersion
    
    oQS.Init sConnection
    oQS.ConnectionOpen
    
    'always follow a connection open with a begin trans
    oQS.BeginTrans
    On Error GoTo ErrorLabelTrans
    
    ' Get the array of values from the database
    vStudy = oQS.SelectArray(oSD)
    
    If IsNull(vStudy) Then
        Set oQS = Nothing       ' This closes connection
        'no study corresponding to this id - return empty string
        Load = ""
'EXIT FUNCTION HERE
        Exit Function
    End If
    
    ' (Keep the QueryServer open for loading Visits and Forms)
    Set oSD = Nothing
    
    Set oStudyState = New StudyDefState
    ' Fill in values from Row 0 of our arrays
    With oStudyState
        ' The Study Definition stuff
        .StudyId = lStudyId
        .Version = nVersion
        .LabelExpr = ConvertFromNull(vStudy(0, 0), vbString)
        .LabelIsLocal = (vStudy(1, 0) = 1)
        .DateFormat = ConvertFromNull(vStudy(2, 0), vbString)
        .TimeFormat = ConvertFromNull(vStudy(3, 0), vbString)
        .CTCScheme = ConvertFromNull(vStudy(4, 0), vbString)
        .DOBExpr = ConvertFromNull(vStudy(5, 0), vbString)
        .GenderExpr = ConvertFromNull(vStudy(6, 0), vbString)
        .FontName = ConvertFromNull(vStudy(7, 0), vbString)
        .FontBold = (ConvertFromNull(vStudy(8, 0), vbInteger) = 1)
        .FontItalic = (ConvertFromNull(vStudy(9, 0), vbInteger) = 1)
        .FontSize = ConvertFromNull(vStudy(10, 0), vbInteger)
        .FontColour = ConvertFromNull(vStudy(11, 0), vbLong)
        .eFormColour = ConvertFromNull(vStudy(12, 0), vbLong)
        .RRServerType = ConvertFromNull(vStudy(13, 0), vbInteger)
        .RRHTTPAddress = ConvertFromNull(vStudy(14, 0), vbString)
        .RRUserName = ConvertFromNull(vStudy(15, 0), vbString)
        .RRPassword = ConvertFromNull(vStudy(16, 0), vbString)
        .RRProxyServer = ConvertFromNull(vStudy(17, 0), vbString)
        
        ' The ClinicalTrial stuff
        ' NCJ 27/9/01 - Added Phase, Status and StudyType
        .Name = vStudy(18, 0)
        .Description = ConvertFromNull(vStudy(19, 0), vbString)
        .Phase = ConvertFromNull(vStudy(20, 0), vbInteger)
        .Status = ConvertFromNull(vStudy(21, 0), vbInteger)
        .StudyType = ConvertFromNull(vStudy(22, 0), vbInteger)
        
        ' The Arezzo definition
        If bIncludeArezzoFile Then
            'arezzo file wanted as well
            .ArezzoFile = vStudy(23, 0)
        Else
            .ArezzoFile = ""
        End If
        
        ' Now we need to get all the eForms and the Visits

        Set .Visits = LoadVisits(oQS, lStudyId, nVersion)
        
        Set .eForms = LoadEForms(oQS, lStudyId, nVersion)
        
        ' NCJ 8 Oct 02 - Load RFCs and RFOs together
        Set colRFCs = New Collection
        Set colRFOs = New Collection
        
'        Set .RFCs = LoadRFCs(oQS, lStudyId, nVersion, colRFCs, colRFOs)
        Call LoadRFCs(oQS, lStudyId, nVersion, colRFCs, colRFOs)
        Set .RFCs = colRFCs
        Set .RFOs = colRFOs
        
        'ic 27/10/2005 added connection so we can check for clinical coding upgrade
        ' NCJ 8 Oct 02 - Load the collection of EFormDateElements
        Set .eFormDateElements = LoadDateElements(oQS, lStudyId, nVersion, sConnection)
        
        ' Return the state as a string
        Load = .State
    End With
    
    oQS.Commit
    On Error GoTo ErrorLabel
    
    Set oQS = Nothing       ' This closes connection
    Set oStudyState = Nothing
    Set colRFCs = Nothing
    Set colRFOs = Nothing
    
    Exit Function
    
ErrorLabel:
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudyDefDS.Load"
    Exit Function
    
ErrorLabelTrans:

    oQS.Rollback
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudyDefDS.Load"
    Exit Function
    
End Function

'--------------------------------------------------------------------
Private Function LoadVisits(oQS As QueryServer, _
            ByVal lStudyId As Long, _
            ByVal nVersion As Integer) As Collection
'--------------------------------------------------------------------
' Load the Visits for the given study/version
' and return info as a collection
' oQS is an open Query Server object
' NCJ 21 May 03 - Added MaxRepeats
'--------------------------------------------------------------------
Dim oQD As QueryDef
Dim vVisit As Variant
Dim oVisitState As VisitState
Dim i As Long
Dim oVisits As Collection
Dim oAllVEFs As Collection

    On Error GoTo ErrorLabel
    
    ' Create the Query Defn objects
    Set oQD = New QueryDef

    ' Select on Study, Version, order by VisitOrder
    oQD.InitSelect msVISIT_TABLE, _
                    Split(msVISIT_COLS, msDELIMITER), _
                    Split(msKEY_COLS, msDELIMITER), _
                    Array(lStudyId, nVersion), _
                    msVISIT_ORDERCOL
    
    ' Read info as array from DB
    vVisit = oQS.SelectArray(oQD)
    
    Set oQD = Nothing
    
    ' Create new collection for Visits
    Set oVisits = New Collection
    
    ' Get all the VisitEForms
    Set oAllVEFs = LoadVEFS(oQS, lStudyId, nVersion)
    
    ' Check we got some data returned
    If Not IsNull(vVisit) Then
        ' Walk through the array
        For i = 0 To UBound(vVisit, 2)     ' Rows is second dimension
            Set oVisitState = New VisitState
            With oVisitState
                .VisitId = vVisit(0, i)
                .Code = vVisit(1, i)
                .Name = vVisit(2, i)
                .VisitOrder = vVisit(3, i)
                .DateLabelExpr = ConvertFromNull(vVisit(4, i), vbString)
                .PromptForDate = (ConvertFromNull(vVisit(5, i), vbInteger) = 1)
                .BackgroundColour = ConvertFromNull(vVisit(6, i), vbLong)
                .MaxRepeats = ConvertFromNull(vVisit(7, i), vbInteger)      ' NCJ 21 May 03
                ' Add the VisitEForms collection for this visit
                ' (there may not be any)
                On Error Resume Next
                Set .VisitEForms = oAllVEFs.Item(Str(.VisitId))
                If Err.Number <> 0 Then
                    ' Set as empty
                    Set .VisitEForms = New Collection
                End If
                On Error GoTo 0
                ' Add to the collection
                oVisits.Add oVisitState, Str(.VisitId)
            End With
        Next
    End If
    
    ' Return the collection of Visits
    Set LoadVisits = oVisits
    
    Set oVisits = Nothing
    Set oVisitState = Nothing
    Set oAllVEFs = Nothing
    
Exit Function

ErrorLabel:
    
    Err.Raise Err.Number, , Err.Description & "|" & "StudyDefDS.LoadVisits"
    Exit Function
    
End Function

'--------------------------------------------------------------------
Private Sub LoadRFCs(oQS As QueryServer, _
                ByVal lStudyId As Long, _
                ByVal nVersion As Integer, _
                ByRef colRFCs As Collection, _
                ByRef colRFOs As Collection)
'--------------------------------------------------------------------
' Load the RFCs for the given study/version
' and return info as a collection
' oQS is an open Query Server object
' NCJ 8 Oct 02 - Also load Reasons For Overrule (RFOs); changed to Sub
'--------------------------------------------------------------------
Dim oQD As QueryDef
Dim vRFC As Variant
Dim i As Long

    On Error GoTo ErrorLabel

    'Create the Query Defn Objects
    Set oQD = New QueryDef
    
    'Select on Study, Version order by ReasonForChangeId
    oQD.InitSelect msRFCs_TABLE, _
                    Split(msRFCs_COLS, msDELIMITER), _
                    Split(msKEY_COLS, msDELIMITER), _
                    Array(lStudyId, nVersion), _
                    msRFCs_ORDERCOL
    
    'Read info as array from DB
     vRFC = oQS.SelectArray(oQD)
    
    Set oQD = Nothing
    
    'Create new collection for RFCs and RFOs
    Set colRFCs = New Collection
    Set colRFOs = New Collection
    
    'Check if we got some data returned
    If Not IsNull(vRFC) Then
        'Walk through the array
        For i = 0 To UBound(vRFC, 2)
            ' Decide whether it's RFC or RFO - type is third field
            If vRFC(2, i) = mnRFC_TYPE Then
                colRFCs.Add vRFC(0, i)
            Else
                colRFOs.Add vRFC(0, i)
            End If
        Next
    End If
     
Exit Sub

ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "StudyDefDS.LoadRFCs"
                
End Sub


'--------------------------------------------------------------------
Private Function LoadEForms(oQS As QueryServer, _
            ByVal lStudyId As Long, _
            ByVal nVersion As Integer) As Collection
'--------------------------------------------------------------------
' Load the eForms for the given study/version
' and return info as a collection
' oQS is an open Query Server object
'--------------------------------------------------------------------
Dim oQD As QueryDef
Dim vEF As Variant
Dim oEFState As eFormState
Dim i As Long
Dim oEFs As Collection

    On Error GoTo ErrorLabel

    ' Create the Query Defn objects
    Set oQD = New QueryDef

    ' Select on Study, Version, order by CRFPageOrder
    oQD.InitSelect msEFORM_TABLE, _
                    Split(msEFORM_COLS, msDELIMITER), _
                    Split(msKEY_COLS, msDELIMITER), _
                    Array(lStudyId, nVersion), _
                    msEFORM_ORDERCOL
    
    ' Read info as array from DB
    vEF = oQS.SelectArray(oQD)
    
    Set oQD = Nothing
    
    ' Create new collection for eForms
    Set oEFs = New Collection
    
    ' Check we got some data returned
    If Not IsNull(vEF) Then
        ' Walk through the array
        For i = 0 To UBound(vEF, 2)     ' Rows is second dimension
            Set oEFState = New eFormState
            With oEFState
                .eFormId = vEF(0, i)
                .Code = vEF(1, i)
                .Name = vEF(2, i)
                .eFormOrder = vEF(3, i)
                .LabelExpr = ConvertFromNull(vEF(4, i), vbString)
                .DateExpr = ConvertFromNull(vEF(5, i), vbString)
                .DisplayNumbers = (vEF(6, i) = 1)
                .HideIfInactive = (vEF(7, i) = 1)
                .BackgroundColour = vEF(8, i)       ' NCJ 20/9/01
                .PromptForDate = (ConvertFromNull(vEF(9, i), vbInteger) = 1)
                .Sequential = (ConvertFromNull(vEF(10, i), vbInteger) = 1)
                .LabelIsLocal = (ConvertFromNull(vEF(11, i), vbInteger) = 1)
                ' NCJ 4 Nov 02 - Added eFormWidth
                .eFormWidth = ConvertFromNull(vEF(12, i), vbLong)
                ' Add to the collection
                oEFs.Add oEFState, Str(.eFormId)
            End With
        Next
    End If
    
    ' Return the collection of eForms
    Set LoadEForms = oEFs
    
    Set oEFs = Nothing
    Set oEFState = Nothing
    
Exit Function

ErrorLabel:
    
    Err.Raise Err.Number, , Err.Description & "|" & "StudyDefDS.LoadEForms"
    Exit Function
    

End Function

'--------------------------------------------------------------------
Private Function LoadVEFS(oQS As QueryServer, _
            lStudyId As Long, nVersion As Integer) As Collection
'--------------------------------------------------------------------
' Load ALL the Visit eForms for the given study/version
' and return info as a collection of collections for each visit
' oQS is an open Query Server object
'
' MLM 28/08/02: Include the EFormUse property
'--------------------------------------------------------------------
Dim oQD As QueryDef
Dim vVEF As Variant
Dim oVEFState As VisitEFormState
Dim i As Long
Dim oVEFs As Collection

    On Error GoTo ErrorLabel
    
    ' Create the Query Defn objects
    Set oQD = New QueryDef

    oQD.InitSelect msVEF_TABLE, _
                Split(msVEF_COLS, msDELIMITER), _
                Split(msKEY_COLS, msDELIMITER), _
                Array(lStudyId, nVersion)
    
    ' Read info as array from DB
    vVEF = oQS.SelectArray(oQD)
    
    Set oQD = Nothing
    
    ' Create new collection for collections of Visit eForms
    Set oVEFs = New Collection
    
    ' Check we got some data returned
    If Not IsNull(vVEF) Then
        ' Walk through the array
        For i = 0 To UBound(vVEF, 2)     ' Rows is second dimension
            Set oVEFState = New VisitEFormState
            With oVEFState
                .VisitId = vVEF(0, i)
                .eFormId = vVEF(1, i)
                .Repeating = (ConvertFromNull(vVEF(2, i), vbInteger) = 1)
                'MLM 28/08/02: Also set new EFormUse property
                .EFormUse = vVEF(3, i)
                ' Add the state to the appropriate collection
                VEFs(oVEFs, .VisitId).Add oVEFState, VisitEFormKey(.VisitId, .eFormId)
            End With
        Next i
    End If
    
    Set LoadVEFS = oVEFs
    
    Set oVEFs = Nothing
    Set oVEFState = Nothing
    
Exit Function

ErrorLabel:
    
    Err.Raise Err.Number, , Err.Description & "|" & "StudyDefDS.LoadVEFs"
    Exit Function
    
    
End Function

'--------------------------------------------------------------------
Private Function VEFs(ByRef oAllVEFs As Collection, _
                        ByVal lVisitId As Long) As Collection
'--------------------------------------------------------------------
' Get the VEF collection from AllVEFs corresponding to visit with given lVisitId
' If collection doesn't exist, add an empty new one and return it
'--------------------------------------------------------------------
Dim sVEFKey As String
Dim oCol As Collection

    sVEFKey = Str(lVisitId)
    On Error Resume Next
    ' See if a collection for this visit already exists
    Set VEFs = oAllVEFs.Item(sVEFKey)
    If Err.Number <> 0 Then
        ' It wasn't there so create a new one
        Set oCol = New Collection
        oAllVEFs.Add oCol, sVEFKey
        Set VEFs = oCol
        Set oCol = Nothing
    End If
    
End Function

'--------------------------------------------------------------------
Public Function LoadGroups(ByVal sConnection As String, _
            ByVal lStudyId As Long, _
            ByVal nVersion As Integer, _
            ByVal lEFormId As Long) As String
'--------------------------------------------------------------------
' Load question groups for the given study/version/eForm
' and return info as a string
' (representing packaged collection of QGroupState strings)
' sConnection is the DB connection string
'--------------------------------------------------------------------
Dim oQS As QueryServer
Dim oSD As QueryDef
Dim i As Long
Dim vGroup As Variant
Dim oQGroupState As QGroupState
Dim colGroupStates As Collection

    On Error GoTo ErrorLabel
    
    ' Create the Query Server and Query Defn objects
    Set oQS = New QueryServer
    Set oSD = New QueryDef

    ' Set up the Tables, the Columns, the Joins and the Filters
    oSD.QueryTables.Add msEFORM_QGROUP_TABLE
    oSD.QueryTables.Add msQGROUP_TABLE, , qdjtLeft, _
            Array(msQGROUP_TABLE & msSTUDYID, msQGROUP_TABLE & msVERSION, _
                            msQGROUP_TABLE & msQGROUPID), _
            Array(msEFORM_QGROUP_TABLE & msSTUDYID, msEFORM_QGROUP_TABLE & msVERSION, _
                            msEFORM_QGROUP_TABLE & msQGROUPID)
    
    oSD.QueryFields.Add Split(msQGROUP_COLS, msDELIMITER)
    oSD.QueryFields.Add Split(msEFORM_QGROUP_COLS, msDELIMITER)
    
    oSD.QueryFilters.Add msEFORM_QGROUP_TABLE & msSTUDYID, "=", lStudyId
    oSD.QueryFilters.Add msEFORM_QGROUP_TABLE & msVERSION, "=", nVersion
    oSD.QueryFilters.Add msCRFPAGEID_COL, "=", lEFormId
    
    ' Order on GroupID
    oSD.QueryOrders.Add Array(msQGROUP_TABLE & msQGROUPID)
    
    oQS.Init sConnection
    oQS.ConnectionOpen
    oQS.BeginTrans
    On Error GoTo ErrorLabelTrans
    
    ' Read info as array from DB
    vGroup = oQS.SelectArray(oSD)
    
    Set oSD = Nothing
    
    oQS.Commit
    On Error GoTo ErrorLabel
    Set oQS = Nothing
    
    ' Create new collection for Group states
    Set colGroupStates = New Collection

    ' Check we got some data returned
    If Not IsNull(vGroup) Then
        ' Walk through the array
        For i = 0 To UBound(vGroup, 2)     ' Rows is second dimension
            Set oQGroupState = New QGroupState
            With oQGroupState
                .QGroupID = ConvertFromNull(vGroup(0, i), vbLong)
                .Code = ConvertFromNull(vGroup(1, i), vbString)
                .Name = ConvertFromNull(vGroup(2, i), vbString)
                .DisplayType = ConvertFromNull(vGroup(3, i), vbInteger)
                .eFormId = ConvertFromNull(vGroup(4, i), vbLong)
                .Border = (ConvertFromNull(vGroup(5, i), vbInteger) = 1)
                .DisplayRows = ConvertFromNull(vGroup(6, i), vbInteger)
                .InitialRows = ConvertFromNull(vGroup(7, i), vbInteger)
                .MinRepeats = ConvertFromNull(vGroup(8, i), vbInteger)
                .MaxRepeats = ConvertFromNull(vGroup(9, i), vbInteger)
                
                ' Add to the collection using Group Id
                colGroupStates.Add oQGroupState, oQGroupState.Key
            End With
        Next i
    End If
    
    ' Return the serialised collection
    LoadGroups = CollectionSerialise(colGroupStates)
    
    Set oQGroupState = Nothing
    Set colGroupStates = Nothing
    
Exit Function

ErrorLabelTrans:
    oQS.Rollback
    ' Drop through to normal error handler
ErrorLabel:
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudyDefDS.LoadGroups"
    Exit Function
    

End Function

'--------------------------------------------------------------------
Private Function LoadDateElements(oQS As QueryServer, _
            ByVal lStudyId As Long, _
            ByVal nVersion As Integer, sConnection As String) As Collection
'--------------------------------------------------------------------
' NCJ 8 Oct 02
' Load the eform date elements for the given study/version
' and return the collection of EFElementState objects
'ic 27/10/2005 added connection so we can check for clinical coding upgrade
'--------------------------------------------------------------------

    Set LoadDateElements = LoadElementsCollection(oQS, lStudyId, nVersion, 0, True, sConnection)

End Function

'--------------------------------------------------------------------
Public Function LoadElements(ByVal sConnection As String, _
            ByVal lStudyId As Long, _
            ByVal nVersion As Integer, _
            ByVal lEFormId As Long) As String
'--------------------------------------------------------------------
' Load the eform elements for the given study/version/eForm
' and return info as a string
' (representing packaged collection of EFElementState strings)
' sConnection is the DB connection string
'
' MLM 28/08/02: Include ElementUse
' NCJ 8 Oct 02 - Main code moved to LoadElementsCollection
'ic 27/10/2005 added connection so we can check for clinical coding upgrade
'--------------------------------------------------------------------
Dim oQS As QueryServer

    On Error GoTo ErrorLabel
    
    Set oQS = New QueryServer
    oQS.Init sConnection
    oQS.ConnectionOpen
    oQS.BeginTrans
    
    On Error GoTo ErrorLabelTrans
    
    ' Return the serialised collection
    LoadElements = CollectionSerialise(LoadElementsCollection(oQS, _
                                                        lStudyId, nVersion, lEFormId, False, sConnection))
    oQS.Commit
    Set oQS = Nothing
    
    Exit Function
    
ErrorLabel:
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudyDefDS.LoadElements"
    Exit Function
    
ErrorLabelTrans:
    oQS.Rollback
    Set oQS = Nothing
    Err.Raise Err.Number, , Err.Description & "|" & "StudyDefDS.LoadElements"
    Exit Function
    
End Function

'--------------------------------------------------------------------
Private Function LoadElementsCollection(oQS As QueryServer, _
            ByVal lStudyId As Long, _
            ByVal nVersion As Integer, _
            ByVal lEFormId As Long, _
            bDateElementsOnly As Boolean, sConnection As String) As Collection
'--------------------------------------------------------------------
' Load the eform elements for the given study/version/eForm
' and return info as a string
' (representing packaged collection of EFElementState strings)
' sConnection is the DB connection string
'
' MLM 28/08/02: Include ElementUse
' NCJ 8 Oct 02 - This routine created from original LoadElements
'               If bDateElementsOnly = TRUE then load all date elements (only) for study,
'                   ignoring lEFormId (which should be 0)
'ic 14/07/2005 added clinical coding
'--------------------------------------------------------------------
Dim oSD As QueryDef
Dim vEFElement As Variant
Dim oEFEState As EFElementState
Dim i As Long
Dim colEFEStates As Collection
Dim nNextQField As Integer
Dim nLastEFField As Integer
Dim colAllCats As Collection
Dim colAllVals As Collection
Dim oVersion As MACROVersion.Checker
Dim bCC As Boolean

    On Error GoTo ErrorLabel
    
    'check for clinical coding version
    Set oVersion = New MACROVersion.Checker
    bCC = oVersion.HasUpgrade(sConnection, mCCSwitch)

    ' Create the Query Defn object
    Set oSD = New QueryDef

    ' Set up the Tables, the Columns, the Joins and the Filters
    oSD.QueryTables.Add msELEMENT_TABLE
    oSD.QueryTables.Add msQUESTION_TABLE, , qdjtLeft, _
            Array(msQUESTION_TABLE & msSTUDYID, msQUESTION_TABLE & msVERSION, msQUESTION_TABLE & msDITEMID), _
            Array(msELEMENT_TABLE & msSTUDYID, msELEMENT_TABLE & msVERSION, msELEMENT_TABLE & msDITEMID)
    
    'ic 14/07/2005 added clinical coding fields
    If (bCC) Then
        oSD.QueryFields.Add Split(msQUESTION_COLS & msCCQUESTION_COLS, msDELIMITER)
    Else
        oSD.QueryFields.Add Split(msQUESTION_COLS, msDELIMITER)
    End If
    oSD.QueryFields.Add Split(msELEMENT_COLS, msDELIMITER)
    
    oSD.QueryFilters.Add msELEMENT_TABLE & msSTUDYID, "=", lStudyId
    oSD.QueryFilters.Add msELEMENT_TABLE & msVERSION, "=", nVersion
    
    ' NCJ 8 Oct 02 - If we want date elements, ignore CRFPageId filter
    If bDateElementsOnly Then
        oSD.QueryFilters.Add msELEMENT_USE_COL, "=", 1   ' 1 is "Date Element", 0 is "User"
    Else
        oSD.QueryFilters.Add msCRFPAGEID_COL, "=", lEFormId
    End If
    
    ' Order on CRFElement order
    oSD.QueryOrders.Add Split(msELEMENT_ORDERCOLS, msDELIMITER)
        
    ' Read info as array from DB
    vEFElement = oQS.SelectArray(oSD)
    
    Set oSD = Nothing
    
    ' We need to load the Categories and Validations too (but not if Date elements only)
    If Not bDateElementsOnly Then
        Set colAllCats = LoadCategories(oQS, lStudyId, nVersion, lEFormId)
        Set colAllVals = LoadValidations(oQS, lStudyId, nVersion, lEFormId)
    Else
        Set colAllCats = New Collection
        Set colAllVals = New Collection
    End If
    
    ' Create new collection for eForm element states
    Set colEFEStates = New Collection
    
    ' Check we got some data returned
    If Not IsNull(vEFElement) Then
        ' Walk through the array
        For i = 0 To UBound(vEFElement, 2)     ' Rows is second dimension
            Set oEFEState = New EFElementState
            With oEFEState
                .Code = ConvertFromNull(vEFElement(0, i), vbString)
                .Name = ConvertFromNull(vEFElement(1, i), vbString)
                .DataType = ConvertFromNull(vEFElement(2, i), vbInteger)
                .Format = ConvertFromNull(vEFElement(3, i), vbString)
                .Unit = ConvertFromNull(vEFElement(4, i), vbString)
                .QuestionLength = ConvertFromNull(vEFElement(5, i), vbInteger)
                .DerivationExpr = ConvertFromNull(vEFElement(6, i), vbString)
                .Helptext = ConvertFromNull(vEFElement(7, i), vbString)
                .TextCase = ConvertFromNull(vEFElement(8, i), vbInteger)
                .ClinicalTest = ConvertFromNull(vEFElement(9, i), vbString)
                
                'ic 14/07/2005 added clinical coding
                If (bCC) Then
                    .DictionaryId = ConvertFromNull(vEFElement(10, i), vbInteger)
                    ' That's the last of the DataItem table fields
                    ' Use a variable so it's easy to add extra DataItem fields
                    ' without having to change all the CRFElement field numbers
                    nNextQField = 11
                Else
                    ' That's the last of the DataItem table fields
                    ' Use a variable so it's easy to add extra DataItem fields
                    ' without having to change all the CRFElement field numbers
                    nNextQField = 10
                End If
                
                ' These are the CRFElement fields
                .QuestionId = vEFElement(nNextQField, i)
                .ElementId = vEFElement(nNextQField + 1, i)
                .ControlType = vEFElement(nNextQField + 2, i)
                .Caption = ConvertFromNull(vEFElement(nNextQField + 3, i), vbString)
                .ElementOrder = vEFElement(nNextQField + 4, i)
                .CollectIfCond = ConvertFromNull(vEFElement(nNextQField + 5, i), vbString)
                .Hidden = (ConvertFromNull(vEFElement(nNextQField + 6, i), vbInteger) = 1)
                .IsLocal = (ConvertFromNull(vEFElement(nNextQField + 7, i), vbInteger) = 1)
                .IsOptional = (ConvertFromNull(vEFElement(nNextQField + 8, i), vbInteger) = 1)
                .IsMandatory = (ConvertFromNull(vEFElement(nNextQField + 9, i), vbInteger) = 1)
                .RequiresRFC = (ConvertFromNull(vEFElement(nNextQField + 10, i), vbInteger) = 1)
                .Authorisation = ConvertFromNull(vEFElement(nNextQField + 11, i), vbString)
                ' NCJ 20 Sep 01 - If Authorisation is "0", denoting "no authorisation",
                ' convert to empty string
                If .Authorisation = "0" Then
                    .Authorisation = ""
                End If
                .ClinicalTestDateExpr = ConvertFromNull(vEFElement(nNextQField + 12, i), vbString)
                .FontName = ConvertFromNull(vEFElement(nNextQField + 13, i), vbString)
                .FontSize = ConvertFromNull(vEFElement(nNextQField + 14, i), vbInteger)
                .FontBold = (ConvertFromNull(vEFElement(nNextQField + 15, i), vbInteger) = 1)
                .FontItalic = (ConvertFromNull(vEFElement(nNextQField + 16, i), vbInteger) = 1)
                .FontColour = ConvertFromNull(vEFElement(nNextQField + 17, i), vbLong)
                .CaptionX = ConvertFromNull(vEFElement(nNextQField + 18, i), vbLong)
                .CaptionY = ConvertFromNull(vEFElement(nNextQField + 19, i), vbLong)
                .ElementX = ConvertFromNull(vEFElement(nNextQField + 20, i), vbLong)
                .ElementY = ConvertFromNull(vEFElement(nNextQField + 21, i), vbLong)
                ' NCJ 5 Nov 01
                .QGroupID = ConvertFromNull(vEFElement(nNextQField + 22, i), vbLong)
                .OwnerQGroupID = ConvertFromNull(vEFElement(nNextQField + 23, i), vbLong)
                .ShowStatusFlag = (ConvertFromNull(vEFElement(nNextQField + 24, i), vbInteger) = 1)
                .GroupElementOrder = ConvertFromNull(vEFElement(nNextQField + 25, i), vbInteger)
                
                'ZA 17/07/2002
                .CaptionFontName = ConvertFromNull(vEFElement(nNextQField + 26, i), vbString)
                'TA 'we covert nulls into zeros  - Val(ConvertFromNull(NULL, vbString)) does this
                .CaptionFontBold = (Val(ConvertFromNull(vEFElement(nNextQField + 27, i), vbString)) = 1)
                .CaptionFontItalic = (Val(ConvertFromNull(vEFElement(nNextQField + 28, i), vbString)) = 1)
                .CaptionFontSize = Val(ConvertFromNull(vEFElement(nNextQField + 29, i), vbString))
                .CaptionFontColour = Val(ConvertFromNull(vEFElement(nNextQField + 30, i), vbString))
                
                'MLM 28/08/02:
                .ElementUse = vEFElement(nNextQField + 31, i)
                
                'ZA 20/09/2002
                .DisplayLength = ConvertFromNull(vEFElement(nNextQField + 32, i), vbInteger)
                
                ' NCJ 8 Nov 02
                .Hotlink = ConvertFromNull(vEFElement(nNextQField + 33, i), vbString)
                
                ' NCJ 8 Oct 02
                If bDateElementsOnly Then
                    .eFormId = vEFElement(nNextQField + 34, i)
                Else
                    .eFormId = lEFormId
                End If
                
                Set .Categories = New Collection
                Set .Validations = New Collection
                
                ' Slot in the categories and validations for a question (if any)
                If .QuestionId > 0 And Not bDateElementsOnly Then
                    If CollectionMember(colAllCats, Str(.QuestionId), True) Then
                        Set .Categories = colAllCats.Item(Str(.QuestionId))
                    End If
                    If CollectionMember(colAllVals, Str(.QuestionId), True) Then
                        Set .Validations = colAllVals.Item(Str(.QuestionId))
                    End If
                End If
                
                ' Add to the collection using eForm Id and Element Id
                colEFEStates.Add oEFEState, oEFEState.Key
                
            End With
        Next
    End If
    
    Set LoadElementsCollection = colEFEStates
    
    Set colEFEStates = Nothing
    Set oEFEState = Nothing
    
    Exit Function
    
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "StudyDefDS.LoadElementsCollection"
    
End Function

'--------------------------------------------------------------------
Private Function LoadCategories(oQS As QueryServer, _
            ByVal lStudyId As Long, _
            ByVal nVersion As Integer, _
            ByVal lEFormId As Long) As Collection
'--------------------------------------------------------------------
' Load all the category items for the given study/version/eForm
' and return info as a collection of collections of category state objects
' indexed by DataItemID
' For each data item, collection of category items is indexed by Category Code
' oQS is the (ready open) Query Server object
'--------------------------------------------------------------------
Dim oCatState As CategoryState
Dim colAllCatStates As Collection
Dim colCatStates As Collection
Dim oQD As QueryDef
Dim vCats As Variant
Dim lThisQuestionID As Long
Dim i As Long

    On Error GoTo ErrorLabel
    
    Set oQD = New QueryDef
    
    ' Set up the Tables, the Columns, the Joins and the Filters
    oQD.QueryTables.Add msCAT_TABLE
    
    oQD.QueryFields.Add Split(msCAT_COLS, msDELIMITER)
    
    oQD.QueryTables.Add msELEMENT_TABLE, , qdjtInner, _
            Array(msELEMENT_TABLE & msSTUDYID, msELEMENT_TABLE & msVERSION, msELEMENT_TABLE & msDITEMID), _
            Array(msCAT_TABLE & msSTUDYID, msCAT_TABLE & msVERSION, msCAT_TABLE & msDITEMID)

    oQD.QueryFilters.Add msELEMENT_TABLE & msSTUDYID, "=", lStudyId
    oQD.QueryFilters.Add msELEMENT_TABLE & msVERSION, "=", nVersion
    oQD.QueryFilters.Add msCRFPAGEID_COL, "=", lEFormId
    
    ' Order on CRFElement order then Category order
    oQD.QueryOrders.Add Split(msCAT_ORDERCOLS, msDELIMITER)
    
    ' Read info as array from DB
    vCats = oQS.SelectArray(oQD)
    Set oQD = Nothing
    
    Set colAllCatStates = New Collection
    ' Check we got some data returned
    If Not IsNull(vCats) Then
        ' Walk through the array
        lThisQuestionID = 0         ' Initialise "question we're on" variable
        For i = 0 To UBound(vCats, 2)     ' Rows is second dimension
            Set oCatState = New CategoryState
            With oCatState
                .QuestionId = vCats(0, i)
                .Code = vCats(1, i)
                .Value = vCats(2, i)
                .Active = (vCats(3, i) = 1)
                .Order = vCats(4, i)
                If .QuestionId <> lThisQuestionID Then
                    ' A new data item - save the previous collection
                    ' (if not the beginning)
                    ' and start a new one for this data item
                    If lThisQuestionID > 0 Then
                        colAllCatStates.Add colCatStates, Str(lThisQuestionID)
                    End If
                    Set colCatStates = New Collection
                    lThisQuestionID = .QuestionId
                End If
                ' Index by (lower case) category code within a data item
                colCatStates.Add oCatState, oCatState.Key
            End With
        Next
        ' Finally add the last collection in
        colAllCatStates.Add colCatStates, Str(lThisQuestionID)
    End If
    
    Set LoadCategories = colAllCatStates
    
    Set oCatState = Nothing
    Set colCatStates = Nothing
    Set colAllCatStates = Nothing
    
Exit Function

ErrorLabel:
    
    Err.Raise Err.Number, , Err.Description & "|" & "StudyDefDS.LoadCategories"
    Exit Function
    

End Function

'--------------------------------------------------------------------
Private Function LoadValidations(oQS As QueryServer, _
            ByVal lStudyId As Long, _
            ByVal nVersion As Integer, _
            ByVal lEFormId As Long) As Collection
'--------------------------------------------------------------------
' Load all the validations for the given study/version/eForm
' and return info as a collection of collections of validation state objects
' indexed by DataItemID
' For each data item, collection of validations is indexed by ValidationID
' oQS is the (already open) Query Server object
'--------------------------------------------------------------------
Dim oValState As ValidationState
Dim colAllValStates As Collection
Dim colValStates As Collection
Dim oQD As QueryDef
Dim vVals As Variant
Dim lThisQuestionID As Long
Dim i As Long

    On Error GoTo ErrorLabel
    
    Set oQD = New QueryDef

    ' Set up the Tables, the Columns, the Joins and the Filters
    oQD.QueryTables.Add msVAL_TABLE
    
    oQD.QueryTables.Add msELEMENT_TABLE, , qdjtInner, _
            Array(msELEMENT_TABLE & msSTUDYID, msELEMENT_TABLE & msVERSION, msELEMENT_TABLE & msDITEMID), _
            Array(msVAL_TABLE & msSTUDYID, msVAL_TABLE & msVERSION, msVAL_TABLE & msDITEMID)

    oQD.QueryTables.Add msVALTYPE_TABLE, , qdjtInner, _
            Array(msVALTYPE_TABLE & msVAL_TYPEID), _
            Array(msVAL_TABLE & msVAL_TYPEID)
    
    oQD.QueryFields.Add Split(msVAL_COLS, msDELIMITER)
    
    oQD.QueryFilters.Add msELEMENT_TABLE & msSTUDYID, "=", lStudyId
    oQD.QueryFilters.Add msELEMENT_TABLE & msVERSION, "=", nVersion
    oQD.QueryFilters.Add msCRFPAGEID_COL, "=", lEFormId
    
    ' Order on CRFElement order then ValidationID order
    oQD.QueryOrders.Add Split(msVAL_ORDERCOLS, msDELIMITER)
    
    ' Read info as array from DB
    vVals = oQS.SelectArray(oQD)
    Set oQD = Nothing
    
    Set colAllValStates = New Collection
    
    ' Check we got some data returned
    If Not IsNull(vVals) Then
        ' Walk through the array
        lThisQuestionID = 0         ' Initialise "question we're on" variable
        For i = 0 To UBound(vVals, 2)     ' Rows is second dimension
            Set oValState = New ValidationState
            With oValState
                .QuestionId = vVals(0, i)
                .ValidationID = vVals(1, i)
                .ValidationCond = ConvertFromNull(vVals(2, i), vbString)
                .MessageExpr = ConvertFromNull(vVals(3, i), vbString)
                .ValidationType = vVals(4, i)
                If .QuestionId <> lThisQuestionID Then
                    ' A new data item - save the previous collection
                    ' (if not the beginning)
                    ' and start a new one for this data item
                    If lThisQuestionID > 0 Then
                        colAllValStates.Add colValStates, Str(lThisQuestionID)
                    End If
                    Set colValStates = New Collection
                    lThisQuestionID = .QuestionId
                End If
                ' Index by Validation ID within a data item
                colValStates.Add oValState, oValState.Key
            End With
        Next
        ' Finally add the last collection in
        colAllValStates.Add colValStates, Str(lThisQuestionID)
    End If
    
    Set LoadValidations = colAllValStates
    
    Set oValState = Nothing
    Set colValStates = Nothing
    Set colAllValStates = Nothing
    
Exit Function

ErrorLabel:
    
    Err.Raise Err.Number, , Err.Description & "|" & "StudyDefDS.LoadValidations"
    Exit Function
    

End Function



 
