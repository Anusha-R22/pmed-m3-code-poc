VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998-2006. All Rights Reserved
'   File:       clsExchange.cls
'   Author:     Andrew Newbigging, March 1998
'   Purpose:    Core routines to import and export study definitions and patient
'   data.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   1    Andrew Newbigging    24/03/98
'   2    Andrew Newbigging     2/04/98
'   3   Andrew Newbigging       15/10/98
'       Several calls to gLog function removed (not necessary to enter them in the log)
'   4   Andrew Newbigging       22/10/98
'       Another call to gLog removed
'   5   Andrew Newbigging       11/11/98
'       Modified CheckImportFile to cope with large import files
'       Mo Morris               15/2/99
'       AutoExportPRD added
'       ExportPRD changed
'       ExportPatientCab changed
'       Andrew Newbigging       8/3/99
'       New routine UpdateChangedFlags to reset the Changed flags after export
'       patient data
'       Mo Morris               24/9/99
'       ExportSDD changed to handle the different table types.
'       Those with a SegmentId between 001 and 299 need to be filtered on ClinicalTrialId
'       Those with a SegmentId between 300 and 499 need to be filtered on ClinicalTrialName
'       Those with a SegmentId greater than 500 don't need filtering (all data is exported)
'       ImportSDD changed to handle the import of tables that are not trial specific
'       (i.e. tables with a segmentID greater than 500, TrialPhase, TrialType, ValidationAction
'       ValidationType). This is done via the addition of private sub routines ImportTrialPhase
'       ImportTrialType, ImportValidationAction and ImportValidationType. Note that these
'       routines have two modes of operation. Under Macro_DM, changes to already existing
'       codes are compulsory, under Macro_EX the user has the option of accepting, declining
'       or aborting the import.
'       Mo Morris               5/10/99 to 28/10/99
'       Converted from DAO to ADO
'       NCJ 11 Dec 99 - ClinicalTrialId variables converted to Long
'       Mo 13/12/99     Id's from integer to Long
'       Mo 19/1/00  private sub ClearCabExtractFolder added and called from
'                   ImportPatientCab
'       Mo 28/1/00  Several changes from adOpenDynamic to adOpenKeyset so that
'       recordcount works properly
'       Mo Morris   23/2/00 Major changes to ExportSDD. vExportFile no longer a parameter
'       Like ExportPRD, ExportSDD now creates a session name and a filelist of
'       files to be cabbed and now calls ExportStudyCAB
'       Changed ExportSDD from a sub to a function so that it returns the name of the cab file that it creates
'       Mo Morris   24/2/00, ExportSDD changed, File exists checks added to exported documents and graphics
'       Mo Morris   3/3/00,     Changes stemming from the inclusion of multimedia files in
'       patient data import and export routines. changes made to ExportPRD, AutoExportPRD, AutoImportPRD
'       Mo Morris   10/3/00 Removed Cancel/Abort option in ImportTrialPhase, ImportTrialType,
'       ImportValidationAction and ImportValidationType.
'       Mo Morris 28/4/00   Changes made to ImportPRd around checking lock/frozen status before
'       importing data.
'       Mo Morris 26/5/00 SR 3396, ImportSDD now retrieves the recruitment figure before deleting
'       Mo Morris 21/6/00 SR 3396, ImportSDD now sets recruitment to 0 on a trial imported for the first time
'       Mo Morris 3/7/2000 SR 3659 'Set MaxDiskSize=0' added to .ddf (cab definition file) in
'       ExportPatientCab and ExportStudyCab, for the purpose of overwriting the default of 1.44MB
'       Mo Morris 4/7/2000 SR3661, changes made to CheckImportFile, method of extracting counter from record
'       re-written using the Split function
'       Mo Morris   18/9/00 - 9/10/00 Changes to ImportSDD so that it now imports the new Normal range/CTC
'       tables ClinicalTestGroup, ClinicalTest, CTCScheme & CTC, with the assistance of subs
'       ImportClinicalTestGroup, ImportClinicalTest, ImportCTCScheme & ImportCTC
'       Mo Morris 11/10/2000 SR 5616 new sub IncrementRecruitment to be called by ImportPRD added
'       Mo Morris 12/10/00 new functions ExportLDD, ImportLDD, ExportLDDCAB, ImportLDDCAB added
'       TA 11/10/2000: new code to import data from old versions
'       TA 16/10/2000: When Lab is imported its normal ranges are completely replaced (likewise for CTC Schemes)
'       NCJ 25/10/00 SR 3998 - Import of Units and UnitConversionFactors
'       NCJ 27/10/00 - Use new GetSQLStringEquals to deal with case-sensitivity of Oracle
'               Ask user if they want to overwrite Toxicity criteria for a CTC scheme
'       NCJ 8/11/00 - Include Units, ClinicalTest and ClinicalTestGroup in Lab Export
'       Mo Morris 9/11/00 - ImportLaboratory now sets Laboratory.Changed = 0 and sets
'       Laboratory.Site = Null when the import is being run from MACRO_EX (i.e not TrialOffice)
'       Mo Morris   14/11/00     Calls to internal function ReplaceCharacters replaced
'       with calls to VB function Replace.
'       Mo Morris 25/1/2001 SR 4124, CheckImportFile now calls new function CleanUpPatientData
'       which in turn calls CheckResponseTimeStamp and CheckOverruleReason
'       TA 01/03/2001: convert local number format to standard when importing CTC and Normal Ranges
'       ASH 22/08/2001 Added routine EvaluateExceptionalSDDTables to cater for New Export for Study Definition due to
'       introduction of CTC ans Lab Ranges
'       ASH 28/08/2001 Added routine EvaluateExceptionalLDDTables to cater for New Export for Lab Definition due to
'       introduction of CTC ans Lab Ranges
'       DPH 17/10/2001 Added check to make sure ExportSDD returns file name
'       MLM 30/11/01 SR4260 Changed ImportTrialType to check whether TrialTypeName already exists before importing.
'                    Also tidied up module level/private variable naming throughout.
'       RJCW 10/01/2002 Patient Import Performance Improvements
'       ZA 24/03/2002 - ArezzoToken
'       DPH 11/04/2002 - Made ClearCabExtractFolder public for use in CAB validation code
'           Changed AutoExportPRD to collect CAB validation information
'           ImportPatientCAB Check process launched properly & write appropriate log entry
'           Subject Locking during Import
'           Added new functions ImportPatientZIP, ExportPatientZIP, ImportStudyDefinitionZIP, ExportStudyZIP
'           ImportLDDZIP, ExportLDDZIP
'           Replaced calls to ExportPatientCAB, ExportStudyCAB, ExportLDDCAB
'       TA 17/04/2002: Removed all references to Actual Recruitment fields
'           (Still one reference in CheckImportFile as we have to allow
'            importing from older versions)
'       DPH 24/02/2002 Added file checks for existing ZIP files
'       MLM 24/06/02: CBB 2.2.16/6 Use new HexEn/DecodeChars library functions in Export/ImportSDD/PRD
'               Also set mlImportVersion from CheckImportFile and use to determine decoding method.
'       MLM 27/06/02: Do the same again in AutoExportPRD.
'       DPH 11/07/2002 - CBBL 2.2.19.25 Changed cursor type of recordset to collect Multimedia data in AutoExportPRD & ExportPRD
'       DPH 22/07/2002 - Added scale to parameter types CBBL 2.2.19.38
'       TA/MLM 25/07/2002 - Locking/Freezing improvemnts in import and performance enhancements
'       ZA 24/09/2002 - Removed PSS function calls
'       RS 26/09/2002 - Added Import Timestamp Timezone Support
'       TA/REM 26/09/02 - changed ConvertFieldtoParameter to take into account the number of columns in the table cols being converted
'       NCJ 2 Oct 02 - Use more accurate IMedNow instead of CDbl(Now)
'       ASH 26/11/2002 - Added HexDecodeChars to fix %25 entry for Units Column in Units and ClinicalTest tables
'       NCJ 24 Dec 02 - Added properties to return Trial, Site and Subject of latest ImportPRD
'       MLM 06/01/03: Tidied up string encoding: use HexEncodeChars in ExportLDD, made bUseHexDecoding module-level and check whenever importing.
'       NCJ 14 Jan 03 - Changed all comparisions with "MACRO_EX" to "MACRO_SM"
'                   Store current Zip file name in module-level variable (so Lock/Freeze routines know it)
'       ASH 21/1/2003 changes to fix dupicate phasename error in ImportTrialPhase
'        NCJ 22 Jan 03 - Added handling for new StandardValue column for DIR and DIRH
'       NCJ 29 Jan 03 - Added handling for new UserNameFull column for DIR and DIRH
'       NCJ 19 Mar 03 - Fixed bug in DecodeString
'       DPH 15/08/2003 - ORACLE proformastate not been saved correctly because of 32k
'       limitation in strored procedures - made saving proformastate for Oracle a special case (in ImportPRD)
'       TA 22/08/2005 - ImportPRD also logs errors to file now
'       TA 31/10/2005: cbd2639 - Order by timestamp if dirh - so that they are imported
'       in the correct order on the server (so the get allocated correct sequence ids for CG)
'       Mo 21/11/2005 COD0210/COD0220 Study definition and patient data export/import changes stemming from clinical coding
'           New function GetVersionFromDB added so that version is taken from Database instead of Software
'       Mo 20/02/2006 COD0210/COD0220 AutoExportPRD and ExportLDD now call GetVersionFromDB
'       TA 15/03/2006: Minor changes so that this version of clsExchange can be used with the utilities module
'       NCJ 26 Jun 06 - Issue 2744 - Created ExportNamedSDD for use in MACRO SD
'       NCJ 21 Aug 06 - Issue 2642 - Check study lock when importing a study in ImportSDD
'       MLM 29/05/08: Issue 2961: After executing stored proc, check for exception and re-raise
'-----------------------------------------------------------------------------------------------------------

Option Explicit
Option Compare Binary
Option Base 0

Public Enum ExchangeError
    Success = 0
    UnknownUnit = 1
    EmptyFile = 2
    Invalid = 4
    TrialExists = 5
    TrialDoesntExist = 6
    'Code added by Mo Morris 27/9/99
    UserAborted = 7
    DirectoryNotFound = 8
    Unknown = 9
    SubjectLock = 10 ' DPH 11/04/2002
    NoPRDFile = 11 ' DPH 25/04/2002
    DecompressFail = 12 ' DPH 25/04/2002
    TrialLocked = 13 ' DPH 10/05/2002
    'Mo 21/11/2005 COD0210/COD0220
    MoreRecentVersion = 14
End Enum

Private mlImportClinicalTrialId As Long
Private msImportClinicalTrialName As String

'MLM 24/06/02: While importing a file, this is set to the version number from which the file was exported
Private mlImportVersion As Long
'MLM 06/01/03: While importing a file, this is set to whether or not the file contains hex encoded strings
Private mbUseHexDecoding As Boolean

' NCJ 24 Dec 02
Private msImportSite As String
Private mlImportSubjectId As Long

' NCJ 14 Jan 03
Private msCurrentZipFileName As String

'constant for new-style string encoding
Private Const msCHARS_TO_ENCODE = """" & vbNewLine 'ASCII 10, 13 and 34
'constants for old-style string encoding
Const msNEW_LINE_REPLACEMENT = "<BR>"
Const msASCII_13_REPLACEMENT = "<13>"
Const msASCII_10_REPLACEMENT = "<10>"


Private msFileName As String

' NCJ 27/10/00 - Store the CTCSchemes that the user DOESN'T want to import
Private mcolCTCSchemesNotToDo As Collection

Public Enum eExecution
    Update = 0
    Insert = 1
End Enum

' RS 04/02/2003 Import using Stored Procedure
' RS 27/02/2003 Declare Command Objects on class level
Dim cmdImportDIR As ADODB.Command
Dim cmdImportDIRH As ADODB.Command
Dim cmdImportTrialSubject As ADODB.Command      ' RS 25/02/2003: Use stored procedures for all Import
Dim cmdImportVisitInstance As ADODB.Command
Dim cmdImportCRFPageInstance As ADODB.Command
'Mo 21/11/2005 COD0220
Dim cmdImportCodingHistory As ADODB.Command


'---------------------------------------------------------------------
Public Property Get LastImportedStudyName() As String
'---------------------------------------------------------------------
' NCJ 24 Dec 02
' The study name of the last imported PRD file
'---------------------------------------------------------------------

    LastImportedStudyName = msImportClinicalTrialName
    
End Property

'---------------------------------------------------------------------
Public Property Get LastImportedStudyID() As Long
'---------------------------------------------------------------------
' NCJ 24 Dec 02
' The study ID of the last imported PRD file
'---------------------------------------------------------------------

    LastImportedStudyID = mlImportClinicalTrialId
    
End Property

'---------------------------------------------------------------------
Public Property Get LastImportedSite() As String
'---------------------------------------------------------------------
' NCJ 24 Dec 02
' The site of the last imported PRD file
'---------------------------------------------------------------------

    LastImportedSite = msImportSite
    
End Property

'---------------------------------------------------------------------
Public Property Get LastImportedSubjectID() As Long
'---------------------------------------------------------------------
' NCJ 24 Dec 02
' The subject ID of the last imported PRD file
'---------------------------------------------------------------------

    LastImportedSubjectID = mlImportSubjectId
    
End Property


'---------------------------------------------------------------------
Public Function ImportSDD(ByVal sImportFile As String) As Long
'---------------------------------------------------------------------
'Changed Mo 26/5/00 SR 3396, retrieve the recruitment figure before deleting
' MLM 24/06/02: Use new string encoding.
' MLM 06/01/03: Use mbUseHexDecoding.
' NCJ 21 Aug 06 - Added study lock (Bug 2642)
'---------------------------------------------------------------------
Dim nImportFileNumber As Integer
Dim rsExportImport As ADODB.Recordset
Dim rsCurrentTable As ADODB.Recordset
Dim rsTrial As ADODB.Recordset
Dim sSegmentId As String
Dim sImportRec As String
Dim sSQL As String
Dim bParamFieldsNotAssessed As Boolean
Dim nNumFields As Integer
Dim i As Integer
Dim vReadField As Variant
Dim lNewTrialId As Long
Dim vRecordNumber As Variant
'ASH 20/1/2003
Dim lPhaseID As Long

Dim lErrNo As Long
Dim sErrDesc As String

Dim bNoDictionariesInstalled As Boolean

' NCJ 21 Aug 06 - Study locking
Dim sLockToken As String
Dim bTrialExists As Boolean

    sLockToken = ""
    bTrialExists = False
    
    'Initialise mlImportClinicalTrialId prior to calling CheckImportFile.
    'If the Trial to be imported already exists, its ClinicalTrialId will be
    'placed in mlImportClinicalTrialId
    mlImportClinicalTrialId = -1
    
    'changed Mo Morris 21/6/00 SR 3396
    'bRecruitmentRetrieved = False
    
    'Mo 21/11/2005 COD0210
    bNoDictionariesInstalled = False
    'If ClinicalCoding is not enabled or there are no dictionaries on file
    'set the flag bNoDictionariesInstalled to true and convert all Thesaurus data items to Text
    If Not gbClinicalCoding Then
        bNoDictionariesInstalled = True
    Else
        If Not DictionariesExist Then
            bNoDictionariesInstalled = True
        End If
    End If
    
    ImportSDD = CheckImportFile(sImportFile, "STYDEF")
    If ImportSDD <> ExchangeError.Success Then
        Exit Function
    End If
    
    On Error GoTo ErrTransactionRollBack
    
    'Begin transaction
    TransBegin
    
    If mlImportClinicalTrialId = -1 Then
        'It's a New Trial that is being imported
        bTrialExists = False
        'Get next available ClinicalTrialId. Retrieve maximum + 1
        sSQL = "SELECT   (max(ClinicalTrialId) + 1) as NewTrialId FROM ClinicalTrial"
        Set rsTrial = New ADODB.Recordset
        rsTrial.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        If IsNull(rsTrial!NewTrialId) Then       'if no records exist
            lNewTrialId = 1
        Else                                    'else if records exist
            lNewTrialId = rsTrial!NewTrialId
        End If
        rsTrial.Close
        Set rsTrial = Nothing
        mlImportClinicalTrialId = lNewTrialId       ' Do we need this? Added by NCJ 21 Aug 06
    Else
        ' NCJ 21 Aug 06 - We delete the trial later (after locking)
'        'Delete the existing Study Definition prior to it being overwritten by the import
'        DeleteTrialSD mlImportClinicalTrialId, msImportClinicalTrialName
        'Use mlImportClinicalTrialId
        lNewTrialId = mlImportClinicalTrialId
        bTrialExists = True
    End If
    
    ' NCJ 21 Aug 06 - Try for a lock on this study
    sLockToken = GetStudyLock(goUser.UserName, lNewTrialId)
    If sLockToken = "" Then
        ' Only jump out if in a GUI module
        If (App.Title = "MACRO_SM") Or (App.Title = "MACRO_UT") Then
            ' Study is in use - get out now
            ' RollBack transaction begun earlier
            TransRollBack
            gLog gsIMPORT_SDD, "Import aborted because study in use - " & msImportClinicalTrialName
            ImportSDD = ExchangeError.TrialLocked
            Exit Function
        Else
            ' During data transfer - We'll go ahead but log a potential problem
            gLog gsIMPORT_SDD, "WARNING - Study imported while study is in use - " & msImportClinicalTrialName
        End If
    End If
    
    If bTrialExists Then
        'Delete the existing Study Definition prior to it being overwritten by the import
        DeleteTrialSD lNewTrialId, msImportClinicalTrialName
    End If

    nImportFileNumber = FreeFile + 1
    Open sImportFile For Input As #nImportFileNumber
    'Read the first line which has already been assessed
    Line Input #nImportFileNumber, sImportRec
    
    'prepare a recordset of the ExportImport parameters file
    'Changed Mo Morris 13/10/00, change from table SDDExportImport to MACROTable
    sSQL = "Select TableName, SegmentId From MACROTable" _
        & " WHERE STYDEF = 1" _
        & " ORDER BY SegmentId"
    Set rsExportImport = New ADODB.Recordset
    'changed Mo Morris 6/12/99 , adOpenForwardOnly replaced by adOpenStatic
    rsExportImport.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
    bParamFieldsNotAssessed = True
    
    ' NCJ 27/10/00 - Initialise CTCScheme collection
    Set mcolCTCSchemesNotToDo = New Collection
    
    'MLM 24/06/02:
    mbUseHexDecoding = (mlImportVersion >= 2020018) '2.2.18
    
    'The Import file is read data item at a time by 2 Input statements.
    'The first reads only the Segment Id.
    'The second is controlled by a loop that is driven by the number of
    'fields within the current record type (nNumFields).
    'Between the 2 Input statements there is code to check if the record
    'type has changed (i.e. the Segment Id has changed) and then via the
    'Export Parameters file accesses the next table name from which a new
    'value for nNumFields can be created.
    Do While Not EOF(nImportFileNumber)
        Input #nImportFileNumber, sSegmentId
        'Debug.Print "Seg " & sSegmentId
        If sSegmentId = "UNT" Then
            Exit Do
        End If
        Do Until sSegmentId = rsExportImport![SegmentId]
            rsExportImport.MoveNext
            bParamFieldsNotAssessed = True
        Loop
        If bParamFieldsNotAssessed Then
            sSQL = " SELECT * FROM " & rsExportImport![TableName]
            Set rsCurrentTable = New ADODB.Recordset
            rsCurrentTable.Open sSQL, MacroADODBConnection, adOpenDynamic, adLockPessimistic, adCmdText
            nNumFields = rsCurrentTable.Fields.Count
            bParamFieldsNotAssessed = False
        End If
        'Try out changes
        If sSegmentId = 510 Then
            ImportSDD = ImportTrialPhase(nImportFileNumber, lPhaseID)
            If ImportSDD <> ExchangeError.Success Then
                Err.Raise _
                    Number:=vbObjectError + 1001, _
                    Description:="Import Aborted by user over TrialPhase discrepancy."
            End If
        ElseIf sSegmentId = 520 Then
            ImportSDD = ImportTrialType(nImportFileNumber, lNewTrialId)
            If ImportSDD <> ExchangeError.Success Then
                Err.Raise _
                    Number:=vbObjectError + 1002, _
                    Description:="Import Aborted by user over Trial Type discrepancy."
            End If
        ElseIf sSegmentId = 530 Then
            ImportSDD = ImportValidationAction(nImportFileNumber)
            If ImportSDD <> ExchangeError.Success Then
                Err.Raise _
                    Number:=vbObjectError + 1003, _
                    Description:="Import Aborted by user over Validation Action discrepancy."
            End If
        ElseIf sSegmentId = 540 Then
            ImportSDD = ImportValidationType(nImportFileNumber)
            If ImportSDD <> ExchangeError.Success Then
                Err.Raise _
                    Number:=vbObjectError + 1004, _
                    Description:="Import Aborted by user over Validation Type discrepancy."
            End If
        ElseIf sSegmentId = 550 Then   ' NCJ 25/10/00 Units table
            ImportSDD = ImportUnits(nImportFileNumber)
            If ImportSDD <> ExchangeError.Success Then
                Err.Raise _
                    Number:=vbObjectError + 1004, _
                    Description:="Import Aborted by user over Units discrepancy."
            End If
        ElseIf sSegmentId = 560 Then   ' NCJ 25/10/00 UnitConversionFactors table
            ImportSDD = ImportConversionFactors(nImportFileNumber)
            If ImportSDD <> ExchangeError.Success Then
                Err.Raise _
                    Number:=vbObjectError + 1004, _
                    Description:="Import Aborted by user over Conversion Factors discrepancy."
            End If
        ElseIf sSegmentId = 610 Then
            ImportSDD = ImportClinicalTestGroup(nImportFileNumber)
            If ImportSDD <> ExchangeError.Success Then
                Err.Raise _
                    Number:=vbObjectError + 1005, _
                    Description:="Import Aborted by user over Clinical Test Group discrepancy."
            End If
        ElseIf sSegmentId = 620 Then
            ImportSDD = ImportClinicalTest(nImportFileNumber)
            If ImportSDD <> ExchangeError.Success Then
                Err.Raise _
                    Number:=vbObjectError + 1006, _
                    Description:="Import Aborted by user over Clinical Test discrepancy."
            End If
        ElseIf sSegmentId = 630 Then
            ImportSDD = ImportCTCScheme(nImportFileNumber)
            If ImportSDD <> ExchangeError.Success Then
                Err.Raise _
                    Number:=vbObjectError + 1007, _
                    Description:="Import Aborted by user over CTC Scheme discrepancy."
            End If
        ElseIf sSegmentId = 640 Then
            Call ImportCTC(nImportFileNumber)
    '        ImportSDD = ImportCTC(nImportFileNumber)
    '        If ImportSDD <> ExchangeError.Success Then
    '            Err.Raise _
    '                Number:=vbObjectError + 1008, _
    '                Description:="Import Aborted by user over CTC discrepancy."
    '        End If
        Else
            rsCurrentTable.AddNew
            For i = 0 To (nNumFields - 1)
                Input #nImportFileNumber, vReadField
                'Debug.Print "[" & i & "]" & vReadField
                'If current data item is ClinicalTrialId then substitute the value of lNewTrialId
                '   ATN 26/1/2000   Must convert case, because Oracle doesn't preserve case in field names
                Select Case LCase$(rsCurrentTable.Fields(i).Name)
                Case "clinicaltrialid"
                    vReadField = lNewTrialId
                Case "phaseid"
                    lPhaseID = vReadField
                'Mo 21/11/2005 COD0210
                Case "dictionaryid"
                    If LCase(rsExportImport![TableName]) = "dataitem" Then
                        If Not bNoDictionariesInstalled Then
                            If Not IsNull(vReadField) Then
                                vReadField = DictionaryNameVersionToId(CStr(vReadField))
                            End If
                        Else
                            If Not IsNull(vReadField) Then
                                vReadField = Null
                            End If
                        End If
                    End If
                Case "datatype"
                    If LCase(rsExportImport![TableName]) = "dataitem" Then
                        If bNoDictionariesInstalled Then
                            If vReadField = DataType.Thesaurus Then
                                'Change the DataType from Thesaurus(8) to Text(0)
                                vReadField = DataType.Text
                            End If
                        End If
                    End If
                'all other fields imported "as is"
                End Select
                'MLM 26/03/03: Bug 1350: When importing a study from pre-3.0, specify that captions have the same appearence as their questions
                'MLM 02/04/03: This was incorrectly changing CaptionX/Y as well as font properties.
                If sSegmentId = "140" And mlImportVersion < 3000000 Then
                    If LCase$(Left(rsCurrentTable.Fields(i).Name, 11)) = "captionfont" And Len(rsCurrentTable.Fields(i).Name) > 7 And (rsCurrentTable.Fields("DataItemId").Value > 0 Or rsCurrentTable.Fields("ControlType").Value = 16386) Then
                        vReadField = rsCurrentTable.Fields(Mid(rsCurrentTable.Fields(i).Name, 8)).Value
                    End If
                End If
                'perform character decoding
                'MLM 24/06/02: perform the correct check based on the version of the import file
                If VarType(vReadField) = vbString Then
                    If mbUseHexDecoding Then
                        'Use new-style decoding if SDD is from after 2.2.17
                        vReadField = HexDecodeChars(CStr(vReadField))
                    ElseIf (InStr(vReadField, msNEW_LINE_REPLACEMENT) <> 0) Then
                        vReadField = Replace(vReadField, msNEW_LINE_REPLACEMENT, vbNewLine)
                    End If
                End If
                rsCurrentTable.Fields(i).Value = vReadField
            Next i
            'Read and ignore line counter at the end of each record
            Input #nImportFileNumber, vRecordNumber
            rsCurrentTable.Update
        End If
    Loop
    
    rsCurrentTable.Close
    Set rsCurrentTable = Nothing
    rsExportImport.Close
    Set rsExportImport = Nothing
    
    Close #nImportFileNumber
    
    'za 14-02-02, added this code for ACM
    'Delete that study details from ArezzoToken table
    ' NCJ 21 Aug 06 - Use "official" version!
    Call MACROLOCKBS30.CacheInvalidateStudy(goUser.CurrentDBConString, lNewTrialId)
'        MacroADODBConnection.Execute "Delete from ArezzoToken where ClinicalTrialID = " & lNewTrialId
        
    'End transaction
    TransCommit
    
    ' NCJ 27/10/00
    Set mcolCTCSchemesNotToDo = Nothing
    
    'Mo Morris 28/2/00 additional error trapping added
    On Error GoTo ErrHandler
    
    gLog gsIMPORT_SDD, StripFileNameFromPath(sImportFile) & " imported successfully"
    
    'changed Mo Morris 23/2/00
    'If frmMenu.Mode = gsTRIAL_SUBJECT_MODE Then
        'delete the imported file
        Kill sImportFile
        'Mo Morris 28/2/00, add  file delete checks
        Do Until Not FileExists(sImportFile)
            DoEvents
        Loop
        gLog gsIMPORT_SDD, StripFileNameFromPath(sImportFile) & " deleted after import"
    'End If
    
    'move the trials study documents and graphics from the CabExtract folder to the Documents folder
    ImportDocumentsAndGraphics
    
    ' NCJ 21 Aug 06
    If sLockToken > "" Then
        Call RemoveStudyLock(lNewTrialId, sLockToken)
    End If
    
    ImportSDD = ExchangeError.Success

Exit Function

ErrTransactionRollBack:

    'store these before calling other routines
    sErrDesc = Err.Description
    lErrNo = Err.Number
    
    On Error Resume Next
    
    'RollBack transaction
    TransRollBack

    ' NCJ 21 Aug 06
    If sLockToken > "" Then
        Call RemoveStudyLock(lNewTrialId, sLockToken)
        sLockToken = ""
    End If
    
    If (lErrNo - vbObjectError) = 1001 Or (lErrNo - vbObjectError) = 1002 _
        Or (lErrNo - vbObjectError) = 1003 Or (lErrNo - vbObjectError) = 1004 _
        Or (lErrNo - vbObjectError) = 1005 Or (lErrNo - vbObjectError) = 1006 _
        Or (lErrNo - vbObjectError) = 1007 Or (lErrNo - vbObjectError) = 1008 Then
        gLog gsIMPORT_SDD, "Error code " & (lErrNo - vbObjectError) & " - " & sErrDesc
        ImportSDD = ExchangeError.UserAborted
    Else
        gLog gsIMPORT_SDD, "Error code " & lErrNo - vbObjectError & " - " & sErrDesc
        gLog gsIMPORT_SDD, "Error occurred on record number= " & vRecordNumber + 1 & " segment= " & sSegmentId
        ImportSDD = ExchangeError.Unknown
    End If

    gLog gsIMPORT_SDD, "Import Aborted. All updates will be Rolled Back."
    Close
    
    ' NCJ 27/10/00
    Set mcolCTCSchemesNotToDo = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportSDD", "clsExchange")
        Case OnErrorAction.Ignore
            ' NCJ 21 Aug 06
            If sLockToken > "" Then
                Call RemoveStudyLock(lNewTrialId, sLockToken)
                sLockToken = ""
            End If
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            If sLockToken > "" Then
                Call RemoveStudyLock(lNewTrialId, sLockToken)
                sLockToken = ""
            End If
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Public Function ExportNamedSDD(ByVal lClinicalTrialId As Long, _
                     ByVal sClinicalTrialName As String, _
                     ByVal nVersionId As Integer, _
                     ByVal sSessionName As String, _
                     Optional sFileSpec As String = "") As Boolean
'---------------------------------------------------------------------
' NCJ 22 Jun 06 - Export a study with a given Session name
' and optional sFileSpec (full file name of destination zip file)
' If no sFileSpec we use SessionName and OutFolder
' Copied from original ExportSDD but with option of giving a file spec
' Return TRUE if OK, or FALSE if failure of some sort
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsExportImport As ADODB.Recordset
Dim rsCurrentTable As ADODB.Recordset
Dim sTableName As String
Dim sSegmentId As String
Dim dtNow As Date
Dim lCount As Long
Dim i As Integer
Dim nExportFileNumber As Integer
Dim nNumFields As Integer
Dim vFieldData As Variant
Dim sVersion As String
Dim asFileList() As String
Dim nFileNumber As Integer
Dim rsStudyDocsAndGraphics As ADODB.Recordset
Dim sStudyDocOrGraphic As String

    On Error GoTo ErrLabel
    
    'dtNow is a date/time variable that is used as a time stamp that will be
    'placed in both the HEADER and FOOTER records of the export files
    'in addition the yyyymmddhhmmss parts of the date stamp are appended to the
    'export file names
    dtNow = Now
    gLog gsEXPORT_SDD, "Session " & sSessionName & " starting"
    
    'initialise FileNumber which will be used as an index in asFileList
    nFileNumber = 0
    
    lCount = 0
    nExportFileNumber = FreeFile
        
    ' Create file name for the export's SDD file
    ' Always use the Out Folder for this temporary sdd file
    msFileName = gsOUT_FOLDER_LOCATION & sClinicalTrialName & "_" & Format(dtNow, "yyyymmddhhmmss") & ".sdd"
    
    'redim the string array asFileList, into which the current file name will be added
    ReDim Preserve asFileList(nFileNumber)
    asFileList(nFileNumber) = msFileName
    nFileNumber = nFileNumber + 1
    
    ' DPH 17/10/2001 Make sure folder exists before opening
    If Not FolderExistence(msFileName) Then
        ExportNamedSDD = False
        gLog gsEXPORT_SDD, "Export file creation failure - Directory " & gsOUT_FOLDER_LOCATION & " cannot be created"
        Exit Function
    End If
    
    Open msFileName For Output As #nExportFileNumber
    lCount = lCount + 1
    'Mo Morris 16/12/99, Version information added to header
    'Mo 21/11/2005 COD0210, Version taken from Database instead of Software
    sVersion = GetVersionFromDB
    Write #nExportFileNumber, "UNH", "STYDEF", dtNow, "UID SENDER", "UID DESTINATION", sVersion, lCount
    
    'Changed Mo Morris 13/10/00, change from table SDDExportImport to MACROTable
    sSQL = "Select TableName, SegmentId From MACROTable" _
        & " WHERE STYDEF = 1" _
        & " ORDER BY SegmentId"
    Set rsExportImport = New ADODB.Recordset
    rsExportImport.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Do While Not rsExportImport.EOF
        'Debug.Print rsExportImport![segmentid] & " " & rsExportImport![TableName]
        sTableName = rsExportImport![TableName]
        sSegmentId = rsExportImport![SegmentId]
        'The tables being exported are of 3 types
        'those with a SegmentId between 001 and 299 need to be filtered on ClinicalTrialId
        'those with a SegmentId between 300 and 499 need to be filtered on ClinicalTrialName
        'those with a SegmentId greater than 500 don't need filtering (all data is exported)
        If sSegmentId < 300 Then
            'files filtered on ClinicalTrialId
            sSQL = " SELECT * FROM " & sTableName _
                & " WHERE ClinicalTrialID = " & lClinicalTrialId
            If sTableName <> "ClinicalTrial" Then
                sSQL = sSQL & " AND VersionId = " & nVersionId
            End If
        ElseIf sSegmentId < 500 Then
            'files filtered on ClinicalTrialName (so far only the table Protocols)
            sSQL = " SELECT * FROM " & sTableName _
                & " WHERE FileName = '" & sClinicalTrialName & "'"
        Else
            'Changed ATO 22/08/2001
            sSQL = EvaluateExceptionalSDDTables(sSegmentId, sTableName, lClinicalTrialId)
            'files that need to be exported in full
            'sSQL = " SELECT * FROM " & sTableName
        End If
        
        Set rsCurrentTable = New ADODB.Recordset
        rsCurrentTable.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not rsCurrentTable.EOF
            nNumFields = rsCurrentTable.Fields.Count
            lCount = lCount + 1
            Write #nExportFileNumber, sSegmentId;
            For i = 0 To (nNumFields - 1)
                If rsCurrentTable.Fields(i).Attributes And adFldLong Then
                    vFieldData = BlobToVariant(rsCurrentTable.Fields(i), rsCurrentTable.Fields(i).ActualSize)
                Else
                    vFieldData = rsCurrentTable.Fields(i).Value
                End If
                'Mo 21/11/2005 COD0210, ClinicalCoding change, replace DictionaryId
                'with DictionaryName|DictionaryVersion in export file
                If LCase(sTableName) = "dataitem" Then
                    If LCase(rsCurrentTable.Fields(i).Name) = "dictionaryid" Then
                        If Not IsNull(vFieldData) Then
                            vFieldData = DictionaryIdToNameVersion(CInt(vFieldData))
                        End If
                    End If
                End If
                'MLM 24/06/02: Use new string encoding
                If VarType(vFieldData) = vbString Then
                    vFieldData = HexEncodeChars(vFieldData, msCHARS_TO_ENCODE)
                End If
                
                'check for vFieldData containing carriage returns
    '            If InStr(vFieldData, Chr(13)) <> 0 Then
    '                vFieldData = Replace(rsCurrentTable.Fields(i).Value, vbNewLine, msNEW_LINE_REPLACEMENT)
    '            End If
                Write #nExportFileNumber, vFieldData;
            Next i
            Write #nExportFileNumber, lCount
            rsCurrentTable.MoveNext
        Loop
        rsCurrentTable.Close
        Set rsCurrentTable = Nothing
        rsExportImport.MoveNext
    Loop
    rsExportImport.Close
    Set rsExportImport = Nothing
    
    lCount = lCount + 1
    Write #nExportFileNumber, "UNT", dtNow, lCount
    
    Close #nExportFileNumber
    gLog gsEXPORT_SDD, msFileName & " Export file created"
    
    'changed Mo Morris 7/11/00, SR4000, ExportStudyDefinitionCab was failing if the same file was being
    'referenced as both a Study reference document and graphical element on a form. To avoid this duplication
    'they are now queried for at the same time using a UNION SQL statement
    'Get the trials study documents and add them to asFileList
    'Get the trials graphical elements and add them to msfilelist
    sSQL = "SELECT DocumentPath from StudyDocument WHERE ClinicalTrialId=" & lClinicalTrialId & " AND VersionId= " & nVersionId
    sSQL = sSQL & " UNION"
    sSQL = sSQL & " SELECT distinct(Caption) FROM CRFElement WHERE ClinicalTrialId=" & lClinicalTrialId & " AND VersionId=" & nVersionId & " AND ControlType = 16388"
    
    Set rsStudyDocsAndGraphics = New ADODB.Recordset
    rsStudyDocsAndGraphics.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsStudyDocsAndGraphics.EOF
        sStudyDocOrGraphic = gsDOCUMENTS_PATH & rsStudyDocsAndGraphics.Fields(0)
        'Mo Morris 24/2/00, File exists check added
        If FileExists(sStudyDocOrGraphic) Then
            ReDim Preserve asFileList(nFileNumber)
            asFileList(nFileNumber) = sStudyDocOrGraphic
            nFileNumber = nFileNumber + 1
            gLog gsEXPORT_SDD, "Study document/graphic " & rsStudyDocsAndGraphics.Fields(0) & " added to export"
        Else
            MsgBox "Reference document or graphic file " & sStudyDocOrGraphic & " not found." & vbNewLine _
                & "The study will be distributed without this file.", vbOKOnly + vbInformation
        End If
        rsStudyDocsAndGraphics.MoveNext
    Loop
    rsStudyDocsAndGraphics.Close
    Set rsStudyDocsAndGraphics = Nothing
    
    ' DPH 11/04/2002 - Using ExportStudyZIP instead of ExportStudyCAB
    ' ExportStudyCAB sSessionName, asFileList
    ' NCJ 26 Jun 06 - Added File spec argument to ExportStudyZIP
    If sFileSpec > "" Then
        ExportStudyZIP sSessionName, asFileList, sFileSpec
    Else
        ' Use default
        ExportStudyZIP sSessionName, asFileList
    End If

    ExportNamedSDD = True
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|clsExchange.ExportNamedSDD"
    
End Function

'---------------------------------------------------------------------
Public Function ExportSDD(ByVal lClinicalTrialId As Long, _
                     ByVal sClinicalTrialName As String, _
                     ByVal nVersionId As Integer, _
                     Optional ByVal nStudyVersioningNo As Integer = -1) As String
'---------------------------------------------------------------------
'Mo Morris   23/2/00 Major changes. vExportFile no longer a parameter
'Like ExportPRD, ExportSDD now creates a session name and a filelist of
'files to be cabbed and now calls ExportStudyCAB
'changed ExportSDD from a sub to a function so that it returns the name of the cab file that it creates
'Mo Morris  7/11/00, SR4000, ExportStudyDefinitionCab was failing if the same file was being
'referenced as both a Study reference document and graphical element on a form. To avoid this duplication
'they are now queried for at the same time using a UNION SQL statement
' DPH 11/04/2002 - Using ExportStudyZIP instead of ExportStudyCAB
' MLM 24/06/02: Added new string encoding
' DPH 07/08/2002 - Optional Study Versioning Number
' NCJ 26 Jun 06 - Most of the code moved to ExportNamedSDD
'---------------------------------------------------------------------
Dim sSessionName As String

    On Error GoTo ErrHandler
    
    ' DPH 07/08/2002 - If Study version being created prefix filename with version number
    If nStudyVersioningNo > -1 Then
        sSessionName = nStudyVersioningNo & sClinicalTrialName
    Else
        sSessionName = sClinicalTrialName
    End If
    
    ' NCJ 22 Jun 06 - All the original code is now in ExportNamedSDD
    If ExportNamedSDD(lClinicalTrialId, sClinicalTrialName, nVersionId, sSessionName) Then
    
        'Mo Morris 24/2/00
        ' DPH 15/04/2002 - Removed CAB reference
        'ExportSDD = sSessionName & ".cab"
        ExportSDD = sSessionName & ".zip"
    Else
        ExportSDD = ""
    End If

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ExportSDD", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Public Function ImportPRD(ByVal sImportFile As String) As Long
'---------------------------------------------------------------------
' RJCW 10/12/2001
' Reads the given text file and updates the Patient relevant tables
' DPH 11/04/2002 - Added Check for MACRO subject locks
' DPH 10/05/2002 - Only lock subject when new subject collected
' MLM 24/06/02: Use new string encoding.
' MLM 24/07/02: Cope with locked and frozen data in the import file!
' TA 25/07/2002 - Performance enhancements (work out whether to do upadte or insert)
' TA 26/07/2002: Updating Lockstatuses needs to be done after updating the recordset
' TA 28/08/2002: ResponseTaskID and ResponseCycle now read into appropriate variables to be used for locking/freezing
' MLM 06/01/03: Use mbUseHexDecoding.
' NCJ 9 Jan 03 - Removed cmdUpdateSameLockStatus (Lock statuses now handled separately)
' MLM 16/01/03: Make sure subject locking and unlocking uses trial id from server, not import file.
' DPH 15/08/2003 - ORACLE proformastate not been saved correctly because of 32k
'   limitation in strored procedures - made saving proformastate for Oracle a special case
' MLM 29/09/06: Issue 2727: Set the MIMessageStatus of an object after importing it.
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsExportImport As ADODB.Recordset
Dim rsCurrentTable As ADODB.Recordset
Dim sSegmentId As String
Dim bParamFieldsNotAssessed As Boolean
Dim nNumFields As Integer
Dim i As Integer
Dim vReadField As Variant
Dim nImportFileNumber As Integer
Dim sImportRec As String
Dim nItemsRead As Integer
Dim lImportTrialId As Long
Dim sImportTrialSite As String
Dim nImportPersonId As Integer
Dim lImportVisitTaskId As Long
Dim lImportCRFPageTaskId As Long
Dim lImportResponseTaskId As Long
Dim nImportRepeatNumber As Integer
Dim nImportStateNumber As Integer
Dim sProformaStateFile As String

' RS 26/09/2002:    Use Timezoneobject to obtain ImportTimestamp Timezone
Dim oTimezone As TimeZone

'registers whether a lock status needs changing so that it can be done after the UPDATE command
'only for subject/visit/eForm - dataitem handled separately
Dim bNeedToChangeLockStatus As Boolean

Dim sErrorString As String

Dim cmdUpdate As New ADODB.Command
Dim cmdInsert As New ADODB.Command

Dim sParamName As String

Dim cmdCurrentCommand As ADODB.Command          ' RS 25/02/2003: Currently Active Command
Dim blnCommandParams As Boolean                 ' RS 25/02/2003: Set if SP is called

''MLM 24/07/02:
'Dim cmdUpdateSameLockStatus As New ADODB.Command
''keep track of lock status
'Dim enLockStatus As LockStatus

Dim lRows As Long
Dim nNoParameters As Integer
Dim nPKItems As Integer
Dim blnProcessNumFields As Boolean
' DPH 11/04/2002 - Subject lock
Dim sToken As String
Dim bFirstTrialSubject As Boolean
'MLM 24/06/02:
'Dim bUseHexDecoding As Boolean
'MLM 24/07/02:
Dim sAction As String

Dim sImportFileNameOnly As String
Dim oQS As QueryServer

' DPH 15/08/2003 - Recordset to file ORACLE proformastate
Dim rsOracleProformaState As ADODB.Recordset
Dim sProformaStateData As String
Dim sTableName As String

'how many times an update that should have been an insert is executed before switching to inserts
Const UPDATE_RETRIES = 3
'count how many insert in a row after failed updates
Dim lInsertAfterFailedUpdateCount As Long
    
    Set oTimezone = New TimeZone        ' RS 26/09/2002: Initialize Timezone

    sToken = ""
    bFirstTrialSubject = False
    
    'Mo Morris 25/1/2001
    'No transaction RollBack error handler added
    On Error GoTo ErrHandler
       
    sImportFileNameOnly = StripFileNameFromPath(sImportFile)
    
    gLog gsIMPORT_PRD, "Import of " & sImportFileNameOnly & " starting"
    
    'Check the import file. Note that CheckImportFile gives a value to mlImportClinicalTrialId
    ImportPRD = CheckImportFile(sImportFile, "PATRSP")
    If ImportPRD <> ExchangeError.Success Then
        gLog gsIMPORT_PRD, sImportFileNameOnly & " failed by CheckImportFile"
        Exit Function
    Else
        gLog gsIMPORT_PRD, sImportFileNameOnly & " passed by CheckImportFile"
    End If
    
    nImportFileNumber = FreeFile
    Open sImportFile For Input As #nImportFileNumber
    
    'Read the first line which has already been assessed
    Line Input #nImportFileNumber, sImportRec
    
    'prepare a recordset of the ExportImport parameters file
    'Changed Mo Morris 13/10/00, change from table PRDExportImport to MACROTable
    sSQL = "Select TableName, SegmentId From MACROTable" _
        & " WHERE PATRSP = 1" _
        & " ORDER BY SegmentId"
    Set rsExportImport = New ADODB.Recordset
    rsExportImport.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
    bParamFieldsNotAssessed = True

    
    'MLM 24/06/02:
    mbUseHexDecoding = (mlImportVersion >= 2020018) '2.2.18
    
    On Error GoTo ErrTransactionRollBack
        
    'Begin transaction
    TransBegin
    
    blnProcessNumFields = True
    
    'The Import file is read data item at a time by 2 Input statements.
    'The first reads only the Segment Id.
    Set rsCurrentTable = New ADODB.Recordset
    
    Do While Not EOF(nImportFileNumber)
      'Initialize the Need_To_Unlock booleans
        bNeedToChangeLockStatus = False
      
        ' DPH 10/05/2002 - Removed reset of token
        ' DPH 11/04/2002 - Reset token in case of an error then will not unlock subject unless necessary
'        sToken = ""
    
        Input #nImportFileNumber, sSegmentId
        'Debug.Print "SegmentId " & sSegmentId
        If sSegmentId = "UNT" Then
            ImportPRD = ExchangeError.Success
            Exit Do
        End If
        Do Until sSegmentId = rsExportImport![SegmentId]
             rsExportImport.MoveNext
             
             If rsExportImport.EOF = True Then
                Exit Do
             End If
             
            bParamFieldsNotAssessed = True
            blnProcessNumFields = True
        Loop
        
        If bParamFieldsNotAssessed Then
        
'' RS 04/04/2003 - Routine no longer needed as new stored procedures are used
            'Set-up command objects for new table
        End If
        
        'The following line is a dummy read of the ClinicalTrialId, because for
        'import purposes ClinicalTrialId will be substituted with mlImportClinicalTrialId
        
        ' NCJ 24 Dec 02 - Store Trial Site and Subject ID in module variables (for use later)
        
        Input #nImportFileNumber, lImportTrialId
        Input #nImportFileNumber, msImportSite
        Input #nImportFileNumber, mlImportSubjectId
        
        'REM 17/09/03 - Change site name to lower case for MACRO 3.0
        msImportSite = LCase(msImportSite)
        
        ' DPH 15/08/2003 - store tablename to process
        sTableName = rsExportImport![TableName]
        Select Case LCase(sTableName)
            Case "trialsubject"
                                
                ' DPH 10/05/2002 - Check MACRO lock is required - one lock per subject
                ' had set incorrectly as one lock per file line (last subject is different)
                ' Not for Exchange (MACRO_EX) as that is using study lock
                ' NCJ 9 Jan 03 - Changed lImportTrialId to mlImportClinicalTrialId
                If (App.Title = "MACRO_SM") Or (App.Title = "MACRO_UT") Then
                    If Not bFirstTrialSubject Then
                        ' Get study lock else quit with error
                        sToken = GetStudyLock(goUser.UserName, mlImportClinicalTrialId)
                        If sToken = "" Then
                            'RollBack transaction
                            TransRollBack
                            'Close Importfile
                            Close #nImportFileNumber
                            ' Write log entry
                            gLog gsIMPORT_PRD, sImportFileNameOnly & " not imported as MACRO Subject lock in existence."
                            ' Set Error
                            ImportPRD = ExchangeError.TrialLocked
                            Exit Function
                        End If
                    End If
                    bFirstTrialSubject = True
                Else
                    ' DPH 11/04/2002 - Check if Subject is MACRO Locked. If it is already then quit
                    sToken = GetSubjectLock(goUser.UserName, mlImportClinicalTrialId, msImportSite, mlImportSubjectId)
                    If sToken = "" Then
                        'RollBack transaction
                        TransRollBack
                        'Close Importfile
                        Close #nImportFileNumber
                        ' Write log entry
                        gLog gsIMPORT_PRD, sImportFileNameOnly & " not imported as MACRO Subject lock in existence."
                        ' Set Error
                        ImportPRD = ExchangeError.SubjectLock
                        Exit Function
                    End If
                End If
                
                ' NCJ 9 Jan 03 - Ensure any unprocessed Lock/Freeze Rollbacks are dealt with first
                ' (before importing any new data)
                Call ProcessSubjectRollBacks(MacroADODBConnection, msImportClinicalTrialName, msImportSite, mlImportSubjectId)
                
                
                ' RS 25/02/2003:  User SP_MACRO_IMP_TRIALSUBJECT
                Set cmdCurrentCommand = cmdImportTrialSubject
                With cmdCurrentCommand
                    .Parameters(1) = mlImportClinicalTrialId
                    .Parameters(2) = msImportSite
                    .Parameters(3) = mlImportSubjectId
                End With
                nItemsRead = 3
                blnCommandParams = True
                
            Case "visitinstance"
          
                ' RS 25/02/2003:  User SP_MACRO_IMP_VISITINSTANCE
                Set cmdCurrentCommand = cmdImportVisitInstance
                Input #nImportFileNumber, lImportVisitTaskId
                With cmdCurrentCommand
                    .Parameters(1) = mlImportClinicalTrialId
                    .Parameters(2) = msImportSite
                    .Parameters(3) = mlImportSubjectId
                    .Parameters(4) = lImportVisitTaskId
                End With
                nItemsRead = 4
                blnCommandParams = True

            Case "crfpageinstance"
            
                Set cmdCurrentCommand = cmdImportCRFPageInstance
                Input #nImportFileNumber, lImportCRFPageTaskId
                With cmdCurrentCommand
                    .Parameters(1) = mlImportClinicalTrialId
                    .Parameters(2) = msImportSite
                    .Parameters(3) = mlImportSubjectId
                    .Parameters(4) = lImportCRFPageTaskId
                End With
                nItemsRead = 4
                blnCommandParams = True
                                
            Case "dataitemresponse"
                
                ' Call SP
                cmdImportDIR.Parameters(1) = mlImportClinicalTrialId
                cmdImportDIR.Parameters(2) = msImportSite
                cmdImportDIR.Parameters(3) = mlImportSubjectId
                
                nItemsRead = 3
                Set cmdCurrentCommand = cmdImportDIR
                blnCommandParams = True
                                
            Case "dataitemresponsehistory"
                
                ' Call SP
                cmdImportDIRH.Parameters(1) = mlImportClinicalTrialId
                cmdImportDIRH.Parameters(2) = msImportSite
                cmdImportDIRH.Parameters(3) = mlImportSubjectId
                
                nItemsRead = 3
                Set cmdCurrentCommand = cmdImportDIRH
                blnCommandParams = True
                
            'Mo 21/11/2005 COD0220
            Case "codinghistory"
                
                cmdImportCodingHistory.Parameters(1) = mlImportClinicalTrialId
                cmdImportCodingHistory.Parameters(2) = msImportSite
                cmdImportCodingHistory.Parameters(3) = mlImportSubjectId
                
                nItemsRead = 3
                Set cmdCurrentCommand = cmdImportCodingHistory
                blnCommandParams = True
                
        End Select
        
        ' COMMAND/SP
        If blnCommandParams Then
            ' DPH 15/08/2003 - Deal with later if Oracle
           sProformaStateData = ""
           
            For i = nItemsRead + 1 To cmdCurrentCommand.Parameters.Count - 1
                Input #nImportFileNumber, vReadField
                'Debug.Print "[" & i & "] " & vReadField
                vReadField = DecodeString(vReadField)
                ' Transform some of the parameter values
                Select Case NormalizeParamName(cmdCurrentCommand.Parameters(i).Name)
                    Case "changed":             vReadField = Changed.Imported
                    Case "importtimestamp":     vReadField = IMedNow
                    Case "importtimestamp_tz":  vReadField = oTimezone.TimezoneOffset
                    Case "proformastate":
                        ' DPH 15/08/2003 - Deal with later if Oracle
                        ' due to Oracle accepting max 32k parameter size in stored procedures
                        sProformaStateData = ReadProformaStateFile(CStr(vReadField))
                        If goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
                            vReadField = ""
                        Else
                            vReadField = sProformaStateData
                        End If
                End Select
                cmdCurrentCommand.Parameters(i).Value = vReadField
            Next i
            cmdCurrentCommand.Execute
            ' MLM 29/05/08: Issue 2961: After executing stored proc, check for exception and re-raise
            If cmdCurrentCommand.Parameters(0).Value <> 1 And cmdCurrentCommand.Parameters(0).Value <> 2 Then
                Err.Raise vbObjectError, , "Exception occurred while executing stored procedure."
            End If
            
            ' DPH 15/08/2003 - Oracle specific to set ProformaState in TrialSubject table
            'Mo 29/11/2005 COD 0220 LCase added
            If LCase(sTableName) = "trialsubject" And goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 _
                    And sProformaStateData <> "" Then
                ' open a new recordset with sole purpose of filing away ProformaState
                Set rsOracleProformaState = New ADODB.Recordset
                sSQL = "SELECT proformastate FROM trialsubject WHERE clinicaltrialid = " & mlImportClinicalTrialId
                sSQL = sSQL & " AND trialsite = '" & msImportSite & "' AND personid = " & mlImportSubjectId
                ' open recordset
                rsOracleProformaState.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockOptimistic
                If Not rsOracleProformaState.EOF Then
                    ' store proformastate
                    rsOracleProformaState(0).Value = sProformaStateData
                End If
                ' update recordset then close
                rsOracleProformaState.Update
                rsOracleProformaState.Close
                Set rsOracleProformaState = Nothing
            End If
            
            ' MLM 29/09/06: Issue 2727:
            With cmdCurrentCommand
                Select Case .CommandText
                    Case "sp_macro_imp_trialsubject"
                        SetImportedMIMsgStatus MacroADODBConnection, MIMsgScope.mimscSubject, _
                            msImportClinicalTrialName, mlImportClinicalTrialId, msImportSite, _
                            mlImportSubjectId
                    Case "sp_macro_imp_visitinstance"
                        SetImportedMIMsgStatus MacroADODBConnection, MIMsgScope.mimscVisit, _
                            msImportClinicalTrialName, mlImportClinicalTrialId, msImportSite, _
                            mlImportSubjectId, .Parameters(5), .Parameters(6)
                    Case "sp_macro_imp_crfpageinstance"
                        SetImportedMIMsgStatus MacroADODBConnection, MIMsgScope.mimscEForm, _
                            msImportClinicalTrialName, mlImportClinicalTrialId, msImportSite, _
                            mlImportSubjectId, .Parameters(5), .Parameters(7), _
                            .Parameters(4)
                    Case "sp_macro_imp_dir"
                        SetImportedMIMsgStatus MacroADODBConnection, MIMsgScope.mimscQuestion, _
                            msImportClinicalTrialName, mlImportClinicalTrialId, msImportSite, _
                            mlImportSubjectId, .Parameters(5), .Parameters(9), _
                            .Parameters(11), .Parameters(4), .Parameters(32)
                    Case Else
                        'DIRH, CodingHistory,... do not have MIMessage status columns to update
                End Select
            End With
        End If
                    
        'Read and ignore line counter at the end of each record
        Input #nImportFileNumber, vReadField
        'Debug.Print "[LineCount] " & vReadField
        'rsCurrentTable.Close
    'Set rsCurrentTable = Nothing
    Loop
    
'
    Close #nImportFileNumber
    
    ' NCJ 9 Jan 03 - Process the Lock/Freeze stuff
    Call ProcessLFMessages(msImportClinicalTrialName, _
                            msImportSite, _
                            mlImportSubjectId, _
                            msCurrentZipFileName)

    If (App.Title = "MACRO_SM") Or (App.Title = "MACRO_UT") Then
        If sToken <> "" Then
            Call RemoveStudyLock(mlImportClinicalTrialId, sToken)
        End If
    Else
        ' DPH 11/04/2002 - unlock record
        If sToken <> "" Then
            Call RemoveSubjectLock(mlImportClinicalTrialId, msImportSite, mlImportSubjectId, sToken)
        End If
    End If
    
    If Not rsExportImport Is Nothing Then
        Set rsExportImport = Nothing
    End If
    
    If Not cmdUpdate Is Nothing Then
        Set cmdUpdate = Nothing
    End If
       
    If Not cmdInsert Is Nothing Then
        Set cmdInsert = Nothing
    End If
    
    'ZA 15/02/02
    'call this procedure to mark this subject invalid in ArezzoToken table
    MarkSubjectInvalid mlImportSubjectId, msImportSite, lImportTrialId
     
    'End transaction
    TransCommit
    
    'delete the imported file
    Kill sImportFile
    Do Until Not FileExists(sImportFile)
        DoEvents
    Loop
    gLog gsIMPORT_PRD, sImportFileNameOnly & " imported and deleted"
    
    Exit Function
    
    'Mo Morris 25/1/2001
ErrHandler:
    sErrorString = "Error occurred during import. Error code " & Err.Number & " - " & Err.Description
    
    'TA 22/08/2005    write error to file (may be db error)
    WriteLogError "clsEchange.ImportPRD", sErrorString
    gLog gsIMPORT_PRD, sErrorString
    gLog gsIMPORT_PRD, "Import aborted. All updates will be rolled back (i.e. not saved)."
    Close
    ImportPRD = ExchangeError.Unknown
    
    Exit Function
    
ErrTransactionRollBack:
    'Changed Mo Morris 16/1/01, store message before RollBack because the RollBack was destroying the error object
    sErrorString = "Error occurred during import. Error code " & Err.Number & " - " & Err.Description
    
    'TA 22/08/2005    write error to file (may be db error)
    WriteLogError "clsEchange.ImportPRD (during transaction)", sErrorString

    ' If subject locked unlock
    If (App.Title = "MACRO_SM") Or (App.Title = "MACRO_UT") Then
        If sToken <> "" Then
            Call RemoveStudyLock(mlImportClinicalTrialId, sToken)
        End If
    Else
        If sToken <> "" Then
            Call RemoveSubjectLock(mlImportClinicalTrialId, msImportSite, mlImportSubjectId, sToken)
        End If
    End If
    
    'RollBack transaction
    TransRollBack
    
    gLog gsIMPORT_PRD, sErrorString
    gLog gsIMPORT_PRD, "Import aborted. All updates will be rolled back (i.e. not saved)."
    Close
    ImportPRD = ExchangeError.Unknown
    
End Function
                
'---------------------------------------------------------------------
Private Function DecodeString(sStringToDecode As Variant) As Variant
'---------------------------------------------------------------------
' Extracted from ImportPRD
' NCJ 19 Mar 03 - Fixed bug which caused empty string to be incorrectly returned sometimes
'---------------------------------------------------------------------
    
    ' NCJ 19 Mar 03 - Set default result here in case all the IF's fail
    DecodeString = sStringToDecode
    
    'MLM 24/07/02: Perform appropriate decoding
    If VarType(sStringToDecode) = vbString Then
        If mbUseHexDecoding Then
            DecodeString = HexDecodeChars(CStr(sStringToDecode))
        ElseIf (InStr(sStringToDecode, "<13>") <> 0) Or (InStr(sStringToDecode, "<10>") <> 0) Then
            DecodeString = PutBackCRandLF(sStringToDecode)
        End If
'    Else
'        DecodeString = sStringToDecode
    End If

End Function


'---------------------------------------------------------------------
Private Function ReadProformaStateFile(sFileName As String) As String
'---------------------------------------------------------------------
' Extracted from ImportPRD
'---------------------------------------------------------------------
Dim nImportStateNumber As Integer

    nImportStateNumber = FreeFile
    'SDM 26/01/00 SR2794
    Open gsCAB_EXTRACT_LOCATION & sFileName For Binary Access Read As nImportStateNumber
'            Open gsAppPath & "CabExtract" & "\" & sProformaStateFile For Binary Access Read As nImportStateNumber
    ReadProformaStateFile = Input(LOF(nImportStateNumber), #nImportStateNumber)
    Close #nImportStateNumber
    
    'SDM 26/01/00 SR2794
    Kill gsCAB_EXTRACT_LOCATION & sFileName
'            Kill gsAppPath & "CabExtract" & "\" & sProformaStateFile
    gLog gsIMPORT_PRD, sFileName & " imported then deleted"
    
    
End Function



'---------------------------------------------------------------------
Public Function ExportPRD(bChangedData As Boolean, _
                    ByVal lClinicalTrialId As Long, _
                    ByVal sTrialSite As String, _
                    ByVal sPersonId As String, _
                    Optional ByRef sSessionName As String _
                    ) As String
'---------------------------------------------------------------------
'changed Mo Morris 6/2/01 to 14/2/01 ExportPRD is now passed Study,site and Subject parameters
'Key SQL statements now optionally filter on Site and PersonID.
' DPH 11/04/2002 - Using ExportPatientZIP instead of ExportPatientCAB
' MLM 24/06/02: Use new character encoding.
' DPH 09/07/2002 - Fix to Oracle connections for ALL subjects export by creating
'   a new connection. Also Close all files in error trap section of code
' DPH 11/07/2002 - CBBL 2.2.19.25 Need Staticly cursored recordset for transactioned updates with SQL Server
'TA 26/11/2002: Return "" for successful import - or error message if failed
'TA
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTrialAndSiteList As ADODB.Recordset
Dim rsExportImport As ADODB.Recordset
Dim rsCurrentTable As ADODB.Recordset
Dim sExportFile As String
Dim sExportFileName As String
Dim sTableName As String
Dim sSegmentId As String
Dim dtNow As Date
Dim lCount As Long
Dim i As Integer
Dim nExportFileNumber As Integer
Dim nNumFields As Integer
Dim vFieldData As Variant
Dim asFileList() As String
Dim nFileNumber As Integer
' TA 15/03/2006 now optionally passed in byref for MACRO_Utilities
'Dim sSessionName As String
Dim sProformaStateFile As String
Dim sProformaStateFileName As String
Dim nProformaStateFileNumber As Integer

Dim rsLocalData As ADODB.Recordset
Dim sLocalDataIds As String
Dim sDataId As String
Dim rsLocalCRFPageLabel As ADODB.Recordset
Dim sLocalCRFPageIds As String
Dim rsLocalTrialSubjectLabel As ADODB.Recordset
Dim sLocalTrialSubjectLabel As String
Dim sVersion As String

 'DPH 09/07/2002 - Added new connection for large server side exports
 'TA 24/08/2002 - we now use the server side cursor for SQL Server and MSDE as well
Dim oConServerSide As ADODB.Connection

Dim rsMultiMedia As ADODB.Recordset
    
    On Error GoTo ErrTransactionRollBack
       
    'setup a recordset of all the trials/sites on file
    'mo morris 22/2/99 TrialSubject.Changed taken out of SQL statement
    sSQL = "SELECT DISTINCT TrialSubject.ClinicalTrialId, ClinicalTrial.ClinicalTrialName, " _
        & " TrialSubject.TrialSite " _
        & " FROM TrialSubject, ClinicalTrial " _
        & " WHERE TrialSubject.ClinicalTrialId = ClinicalTrial.ClinicalTrialId " _
        & " AND TrialSubject.ClinicalTrialid = " & lClinicalTrialId
        
    'Changed Mo Morris 14/2/01, filter on Site and personId added
    If sTrialSite <> "All Sites" Then
        sSQL = sSQL & " AND TrialSubject.TrialSite = '" & sTrialSite & "'"
    End If
    
    If sPersonId <> "All Subjects" Then
        sSQL = sSQL & " AND TrialSubject.PersonId = " & CLng(sPersonId)
    End If
    
    If bChangedData Then
        sSQL = sSQL & " AND TrialSubject.Changed = 1"
    End If
    
    Set rsTrialAndSiteList = New ADODB.Recordset

    rsTrialAndSiteList.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
    If rsTrialAndSiteList.EOF Then
        'Mo Morris 7/2/01, User message added
        MsgBox "There is no data to export for the current selection.", vbOKOnly, gsDIALOG_TITLE
        gLog gsEXPORT_PRD, "Not started. No data to export"
        rsTrialAndSiteList.Close
        Set rsTrialAndSiteList = Nothing
        Exit Function
    Else
        'changed Mo Morris 17/1/00
        'When run against an empty Access database the recordset rsTrialAndSiteList was
        'erroneously returning an empty recordset (see Q181479), which is now tested for
        If rsTrialAndSiteList!ClinicalTrialId = "" Then
            'Mo Morris 7/2/01, User message added
            MsgBox "There is no data to export for the current selection.", vbOKOnly, gsDIALOG_TITLE
            gLog gsEXPORT_PRD, "Not started. No data to export"
            rsTrialAndSiteList.Close
            Set rsTrialAndSiteList = Nothing
            Exit Function
        End If
        'dtNow is a date/time variable that is used as a time stamp that will be
        'placed in both the HEADER and FOOTER records of the export files
        'in addition the yyyymmddhhmmss parts of the date stamp are appended to the
        'export file names
        dtNow = Now
        sSessionName = gsEXPORT_PRD
        If bChangedData Then
            sSessionName = sSessionName & "Changed"
        Else
            sSessionName = sSessionName & "All"
        End If
        sSessionName = sSessionName & Format(dtNow, "yyyymmddhhmmss")
        gLog gsEXPORT_PRD, "Session " & sSessionName & " starting"
    End If
    
    'Begin transaction
    TransBegin
    
    rsTrialAndSiteList.MoveFirst
    
    'initialise FileNumber which will be used as an index in asFileList
    nFileNumber = 0
    
    Do While Not rsTrialAndSiteList.EOF
    
        lCount = 0
        'setup a file name and location for each trial/site combination
        sExportFileName = rsTrialAndSiteList![ClinicalTrialName] & "_" _
            & rsTrialAndSiteList![TrialSite] & "_" & Format(dtNow, "yyyymmddhhmmss") & ".prd"
        'SDM 26/01/00 SR2794
        sExportFile = gsOUT_FOLDER_LOCATION & sExportFileName
        'sExportFile = gsOUT_FOLDER_LOCATION & "\" & sExportFileName
        
        'redim the string array asFileList, into which the current file name will be added
        ReDim Preserve asFileList(nFileNumber)
        asFileList(nFileNumber) = sExportFile
        nFileNumber = nFileNumber + 1
    
        ' DPH 17/10/2001 Make sure folder exists before opening
        ' Handled by existing error handler
        Call FolderExistence(sExportFile)
        
        'open export file and put a header record in it
        nExportFileNumber = FreeFile
        Open sExportFile For Output As #nExportFileNumber
        lCount = lCount + 1
        'Mo Morris 16/12/99, Version information added to header
        'Mo 21/11/2005 COD0220, Version taken from Database instead of Software
        sVersion = GetVersionFromDB
        Write #nExportFileNumber, "UNH", "PATRSP", rsTrialAndSiteList![ClinicalTrialName], dtNow, "UID SENDER", "UID DESTINATION", sVersion, lCount
    
    '   Select all data which is flagged as local
        sSQL = "SELECT CRFPageId,CRFElementId FROM CRFElement " _
            & " WHERE ClinicalTrialId = " & rsTrialAndSiteList!ClinicalTrialId _
            & "   AND LocalFlag = 1"
        
        Set rsLocalData = New ADODB.Recordset
        'changed by Mo Morris 14/1/00, from adOpenForwardOnly to adOpenStatic
        rsLocalData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        sLocalDataIds = ""
    '   Build a string containing the ids of the local data items
        Do While Not rsLocalData.EOF
            sLocalDataIds = sLocalDataIds & rsLocalData!CRFPageId & "|" & rsLocalData!CRFElementId & ","
            rsLocalData.MoveNext
        Loop
        rsLocalData.Close
        Set rsLocalData = Nothing
    
    '   Select all forms which are flagged as having local Form Labels
        sSQL = "SELECT CRFPageId FROM CRFPage " _
            & " WHERE ClinicalTrialId = " & rsTrialAndSiteList!ClinicalTrialId _
            & "   AND LocalCRFPageLabel = 1"
        
        Set rsLocalCRFPageLabel = New ADODB.Recordset
        rsLocalCRFPageLabel.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        sLocalCRFPageIds = ""
    '   Build a string containing the ids of the forms with local Form Labels
        Do While Not rsLocalCRFPageLabel.EOF
            sLocalCRFPageIds = sLocalCRFPageIds & rsLocalCRFPageLabel!CRFPageId & ","
            rsLocalCRFPageLabel.MoveNext
        Loop
        rsLocalCRFPageLabel.Close
        Set rsLocalCRFPageLabel = Nothing
      
    '   Check if trial subject label is local
        sSQL = "SELECT LocalTrialSubjectLabel FROM StudyDefinition " _
            & " WHERE ClinicalTrialId = " & rsTrialAndSiteList!ClinicalTrialId
        
        Set rsLocalTrialSubjectLabel = New ADODB.Recordset
        rsLocalTrialSubjectLabel.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If rsLocalTrialSubjectLabel!LocalTrialSubjectLabel = 0 Then
            sLocalTrialSubjectLabel = False
        Else
            sLocalTrialSubjectLabel = True
        End If
        rsLocalTrialSubjectLabel.Close
        Set rsLocalTrialSubjectLabel = Nothing
    
        'Changed Mo Morris 13/10/00, change from table PRDExportImport to MACROTable
        sSQL = "Select TableName, SegmentId From MACROTable" _
            & " WHERE PATRSP = 1" _
            & " ORDER BY SegmentId"
        Set rsExportImport = New ADODB.Recordset
        rsExportImport.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
        
        'TA 24/08/2002: Switch to server side for both oracle and SQL server
        If sPersonId = "All Subjects" Then
            ' Open a server side connection else failure may occur
            Set oConServerSide = New ADODB.Connection
            oConServerSide.CursorLocation = adUseServer
            oConServerSide.Open gsADOConnectString
        End If
        
        Do While Not rsExportImport.EOF
            sTableName = rsExportImport![TableName]
            sSegmentId = rsExportImport![SegmentId]
    
            sSQL = " SELECT * FROM " & sTableName _
                & " WHERE ClinicalTrialID = " & rsTrialAndSiteList![ClinicalTrialId] _
                & " AND TrialSite = '" & rsTrialAndSiteList![TrialSite] & "'"
            
            'Changed Mo Morris 14/2/01, filter on PersonId added
            If sPersonId <> "All Subjects" Then
                sSQL = sSQL & " AND PersonID = " & CLng(sPersonId)
            End If
            
            If bChangedData Then
                sSQL = sSQL & " AND Changed = 1"
            End If
            
            Set rsCurrentTable = New ADODB.Recordset
            'DPH 09/07/2002 - Use server side connection if 'all subjects'
            If sPersonId <> "All Subjects" Then
                rsCurrentTable.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
            Else
                If goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
                    rsCurrentTable.Open sSQL, oConServerSide, adOpenForwardOnly, adLockReadOnly, adCmdText
                Else
                    rsCurrentTable.Open sSQL, oConServerSide, adOpenStatic, adLockReadOnly, adCmdText
                End If
            End If
            
            Do While Not rsCurrentTable.EOF
            'Check for local data items and don't export them
                If sLocalDataIds > "" Then
                    If sTableName = "DataItemResponse" Or sTableName = "DataItemResponseHistory" Then
                        sDataId = rsCurrentTable!CRFPageId & "|" & rsCurrentTable!CRFElementId & ","
                        Do While InStr(sLocalDataIds, sDataId) > 0
                            rsCurrentTable.MoveNext
                            If rsCurrentTable.EOF Then
                                Exit Do
                            End If
                            sDataId = rsCurrentTable!CRFPageId & "|" & rsCurrentTable!CRFElementId & ","
                        Loop
                        'Continue the Exit Do from the above loop
                        If rsCurrentTable.EOF Then
                            Exit Do
                        End If
                    End If
                End If
                      
                nNumFields = rsCurrentTable.Fields.Count
                lCount = lCount + 1
                Write #nExportFileNumber, sSegmentId;
                'write each data item to the export file one at a time
                For i = 0 To (nNumFields - 1)
                    'check for the dataitem's 'ProformaState' which contains binary data and
                    'hence can not be written to file with the Write command
                    '   ATN 26/1/2000   Must convert case, because Oracle doesn't preserve case in field names
                    If LCase$(rsCurrentTable.Fields(i).Name) = "proformastate" Then
                        'create a separate file for the ProformaState data and
                        'add it to the list of files included in this export session
                        'note that ProformaState will pertain to the PersonId held in
                        'rsCurrentTable.Fields(2).Value
                        sProformaStateFileName = rsTrialAndSiteList![ClinicalTrialName] & "_" _
                            & rsTrialAndSiteList![TrialSite] & "_" & rsCurrentTable.Fields(2).Value & ".psf"
                        'SDM 26/01/00 SR2794
                        sProformaStateFile = gsOUT_FOLDER_LOCATION & sProformaStateFileName
    '                    sProformaStateFile = gsOUT_FOLDER_LOCATION & "\" & sProformaStateFileName
                        'add sProformaStateFile to asFileList
                        ReDim Preserve asFileList(nFileNumber)
                        asFileList(nFileNumber) = sProformaStateFile
                        nFileNumber = nFileNumber + 1
                        ' DPH 17/10/2001 Make sure folder exists before opening
                        ' Existing error handling controls any errors
                        Call FolderExistence(sProformaStateFile)
                        
                        'put proformaState data into sProformaStateFile
                        nProformaStateFileNumber = FreeFile
                        Open sProformaStateFile For Output As nProformaStateFileNumber
                        Print #nProformaStateFileNumber, rsCurrentTable.Fields(i).Value
                        Close #nProformaStateFileNumber
                        gLog gsEXPORT_PRD, sProformaStateFileName & " ProformaState file created"
                        'Store the name of the ProformaState file in the .prd file
                        Write #nExportFileNumber, sProformaStateFileName;
                    Else
                        If rsCurrentTable.Fields(i).Attributes And adFldLong Then
                            vFieldData = BlobToVariant(rsCurrentTable.Fields(i), rsCurrentTable.Fields(i).ActualSize)
                        Else
                            vFieldData = rsCurrentTable.Fields(i).Value
                        End If
                        '   Check for local CRFPageLabels and don't export them
                    '   ATN 26/1/2000   Must convert case, because Oracle doesn't preserve case in field names
                        If LCase$(rsCurrentTable.Fields(i).Name) = "crfpageinstancelabel" Then
                            If sLocalCRFPageIds > "" Then
                                If InStr(sLocalCRFPageIds, sLocalCRFPageIds) > 0 Then
                                    vFieldData = ""
                                End If
                            End If
                    '   ATN 26/1/2000   Must convert case, because Oracle doesn't preserve case in field names
                        ElseIf LCase$(rsCurrentTable.Fields(i).Name) = "localidentifier1" Then
                            If sLocalTrialSubjectLabel Then
                                vFieldData = ""
                            End If
                        End If
                        'MLM 24/06/02: Use new string encoding
                        If VarType(vFieldData) = vbString Then
                            vFieldData = HexEncodeChars(vFieldData, msCHARS_TO_ENCODE)
                        End If
                        'check for vFieldData containing carriage returns
    '                    If (InStr(vFieldData, Chr(13)) <> 0) Or (InStr(vFieldData, Chr(10)) <> 0) Then
    '                        vFieldData = StripOutCRandLF(vFieldData)
    '                    End If
                        Write #nExportFileNumber, vFieldData;
                    End If
                Next i
     
                Write #nExportFileNumber, lCount
                rsCurrentTable.MoveNext
            Loop    'Looping on rsCurrentTable
            If bChangedData Then
                'Update the Changed Flag from 1 to 0 in the originating record
    '            sSQL = "UPDATE " & sTableName & " SET " _
    '                & " Changed = 0 " _
    '                & " WHERE ClinicalTrialID = " & rsTrialAndSiteList![ClinicalTrialId] _
    '                & " AND TrialSite = '" & rsTrialAndSiteList![TrialSite] & "'" _
    '                & " AND Changed = 1"
    '            gdbMACRO.Execute (sSQL)
            End If
            rsExportImport.MoveNext
        Loop    'Looping on rsExportImport
        
        lCount = lCount + 1
        Write #nExportFileNumber, "UNT", dtNow, lCount
        Close #nExportFileNumber
        gLog gsEXPORT_PRD, sExportFileName & " Export file created"
     
        'query the DataItemResponseHistory table for dataitems of type multimedia and add them to the export
        sSQL = "SELECT DISTINCT ResponseValue FROM DataItemResponseHistory, DataItem " _
            & " WHERE DataItemResponseHistory.ClinicalTrialId = DataItem.ClinicalTrialId " _
            & " AND DataItemResponseHistory.DataItemId = DataItem.DataItemId " _
            & " AND DataItemResponseHistory.ClinicalTrialId = " & rsTrialAndSiteList!ClinicalTrialId _
            & " AND DataItemResponseHistory.TrialSite = '" & rsTrialAndSiteList!TrialSite & "'"
        'Changed Mo Morris 14/2/01, filter on PersonId added
        If sPersonId <> "All Subjects" Then
            sSQL = sSQL & " AND DataItemResponseHistory.PersonId = " & CLng(sPersonId)
        End If
        sSQL = sSQL & " AND DataItem.DataType = " & DataType.Multimedia
        If bChangedData Then
            sSQL = sSQL & " AND DataItemResponseHistory.Changed = " & Changed.Changed
        End If
        Set rsMultiMedia = New ADODB.Recordset
        ' DPH 11/07/2002 - CBBL 2.2.19.25 Need Staticly cursored recordset for transactioned updates with SQL Server
    '    rsMultiMedia.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        rsMultiMedia.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not rsMultiMedia.EOF
            If FileExists(gsDOCUMENTS_PATH & rsMultiMedia!ResponseValue) Then
                'check that this multimedia file has not already been added to this export
                'i.e. 2 patients pointing at the same multimedia file
                If Not FileExists(gsOUT_FOLDER_LOCATION & rsMultiMedia!ResponseValue) Then
                    'copy file to out folder and add to list of files to be exported
                    FileCopy gsDOCUMENTS_PATH & rsMultiMedia!ResponseValue, gsOUT_FOLDER_LOCATION & rsMultiMedia!ResponseValue
                    ReDim Preserve asFileList(nFileNumber)
                    asFileList(nFileNumber) = gsOUT_FOLDER_LOCATION & rsMultiMedia!ResponseValue
                    nFileNumber = nFileNumber + 1
                    gLog gsEXPORT_PRD, rsMultiMedia!ResponseValue & " MultiMedia file added to export"
                End If
            End If
            rsMultiMedia.MoveNext
        Loop
        rsMultiMedia.Close
        Set rsMultiMedia = Nothing
        
        rsTrialAndSiteList.MoveNext
     
    Loop    'Looping on rsTrialAndSiteList
    
    rsTrialAndSiteList.Close
    Set rsTrialAndSiteList = Nothing
    rsExportImport.Close
    Set rsExportImport = Nothing
    rsCurrentTable.Close
    Set rsCurrentTable = Nothing
    
    ' DPH 09/07/2002 - Close Server Connection (if need to)
    If Not (oConServerSide Is Nothing) Then
        If oConServerSide.State = adStateOpen Then
            oConServerSide.Close
        End If
    End If
    Set oConServerSide = Nothing
    
    'End transaction
    TransCommit
    
    ' DPH 11/04/2002 - Using ExportPatientZIP instead of ExportPatientCAB
    'ExportPatientCAB sSessionName, asFileList
    ExportPatientZIP sSessionName, asFileList
    
    gLog gsEXPORT_PRD, "Session " & sSessionName & " completed"
    
    ExportPRD = ""
    
    Exit Function
    
ErrTransactionRollBack:
    'RollBack transaction
    
    ExportPRD = "Error code " & Err.Number & " - " & Err.Description & vbCrLf & "." & _
                    "Export Aborted. All updates will be Rolled Back."
    
    TransRollBack
    
    gLog gsEXPORT_PRD, "Error code " & Err.Number & " - " & Err.Description
    gLog gsEXPORT_PRD, "Export Aborted. All updates will be Rolled Back."
    
    'remove the .prd and .psf files created by this erroneous export session
    ' DPH 09/07/2002 - Close all open files
    'Close #nExportFileNumber
    Close
    For i = 0 To UBound(asFileList)
        Kill asFileList(i)
        gLog gsEXPORT_PRD, "Removing after error " & StripFileNameFromPath(asFileList(i))
    Next i
    
    
End Function

'---------------------------------------------------------------------
Public Sub ExportPatientCAB(sCabFileName As String, asCABFileList() As String)
'---------------------------------------------------------------------
'Packages a list of Patient data PRD files and Patient ProformaState files
'into a single CAB file.
'Re-written by Mo Morrris 15/2/99
'Code clean up by Mo Morris 23/2/00
'---------------------------------------------------------------------
Dim nFileNumber As Integer
Dim sCABfile As String
Dim i As Long

gLog gsEXPORT_PAT_CAB, "Session " & sCabFileName & " starting"

'Create a ddf file containing all the items to be compressed. To be used as argument for makeCAB.exe
nFileNumber = FreeFile
sCABfile = gsOUT_FOLDER_LOCATION & sCabFileName & ".ddf"

' DPH 17/10/2001 Make sure folder exists before opening
If FolderExistence(sCABfile) Then
    Open sCABfile For Output As nFileNumber
    
    'Place initial settings into the .ddf file
    Print #nFileNumber, ".OPTION EXPLICIT"
    Print #nFileNumber, ".Set CabinetNameTemplate=" & sCabFileName & ".cab"
    Print #nFileNumber, ".Set DiskDirectoryTemplate=" & gsOUT_FOLDER_LOCATION
    Print #nFileNumber, ".Set DiskLabel1=" & sCabFileName
    'Changed Mo Morris 3/7/2000 SR 3659
    Print #nFileNumber, ".Set MaxDiskSize=0"
    
    'Add the list of files to be compressed from asCABFileList to the .ddf file
    For i = 0 To UBound(asCABFileList)
        Print #nFileNumber, """" & asCABFileList(i) & """"
    Next i
    
    Close #nFileNumber
    
    'Change directory to where MakeCab.exe exists
    'Note that the DiskDirectoryTemplate location can either be a folder below gsAppPath
    'or as in the case of Roche a whole path location that could be anywhere
    
    ' DPH 17/10/2001 Make sure folder exists before opening
    Call FolderExistence(gsAppPath & "dummy.txt")
    
    ChDir gsAppPath
    
    'MakeCab runs under DOS and as such would normally be called via the Shell command.
    'Unfortunately the Shell command simply starts a process and then moves on to the next
    'line of code, which in this case would start the killing of the files to be CABed.
    'Instead the API utility ExecCmd, which completes the called process before moving onto
    'the next line,is used.
    'changed Mo Morris 7/6/2000 SR 3449 "COMMAND NUL /c " added for the purpose of hiding the DOS window
    'ExecCmd "COMMAND NUL /c Makecab /f " & """" & sCABfile & """"
    'changed back Mo Morris 16/6/2000
    ExecCmd "Makecab /f " & """" & sCABfile & """"
    
    'Kill off the files that have been compacted into a CAB file
    For i = 0 To UBound(asCABFileList)
        Kill asCABFileList(i)
        gLog gsEXPORT_PAT_CAB, "Removing " & StripFileNameFromPath(asCABFileList(i))
    Next i
    
    'Kill the .ddf file
    Kill sCABfile
    
    gLog gsEXPORT_PAT_CAB, "Session " & sCabFileName & " completed"
Else
    gLog gsEXPORT_PAT_CAB, "Session failed as could not create CAB file directory " & gsOUT_FOLDER_LOCATION
End If
End Sub

'---------------------------------------------------------------------
Private Function CheckImportFile(ByVal sImportFile As String, _
                                ByVal sMessageType As String) As Long
'---------------------------------------------------------------------
'Mo Morris 28/4/00, Stemming from SR3249 'Close nImportFileNumber' placed
'above all Exit Function calls
'Mo Morris 4/7/2000 SR3661, method of extracting counter from record re-written
'TA 11/10/2000: new code to import data from old versions
'DPH 17/10/2001 Make sure folder exists before opening
'MLM 24/06/02: If check is successful, set mlImportVersion
'---------------------------------------------------------------------
Dim nTrialTypeId As Integer     ' NCJ 28 Aug 02 - Renamed from sSponsor
Dim lClinicalTrialId As Long
Dim sClinicalTrialName As String
Dim sClinicalTrialDescription As String
Dim nPhaseId As Integer
Dim nStatusId As Integer
Dim sKeywords As String
Dim lExpectedRecruitment As Long
Dim lActualRecruitment As Long
Dim sPrompt As String
'Dim nButtons As Integer
'Dim sTitle As String
Dim nResponse As Integer
Dim sSQL As String
Dim rsClinicalTrials As ADODB.Recordset
Dim nImportFileNumber As Integer
Dim sSegmentId As String
Dim sImportMessageType As String
Dim dtTimeStamp As Date
Dim sUIDSender As String
Dim sUIDDestination As String
Dim lCount As Long
Dim sCurrentCount As String
Dim sFooterRec As String
Dim sFooterCheck As String
'Dim lStartChar As Long
'Dim bUnknownError As Boolean
Dim sImportRec As String
Dim sVersion As String
Dim asFindCount() As String
Dim lNumberOfElements As Long
Dim oNewDBColumns As clsNewDBColumns
Dim sMessage As String
Dim nVersionMajor As Integer
Dim nVersionMinor As Integer
Dim nVersionRevision As Integer
Dim nVersionNumber As Long
Dim bFileExists As Boolean
    
    On Error GoTo ErrHandler

    'Check Import file for references to Units unknown to the importing site
    'If UnknownUnit Then
    '    CheckImportFile = ExchangeError.UnknownUnit
    '    Exit Function
    'End If
   
    ' DPH 17/10/2001 Make sure folder exists before opening
    If Not (FolderExistence(sImportFile, True)) Then
        CheckImportFile = ExchangeError.DirectoryNotFound
        Exit Function
    End If
   
    nImportFileNumber = FreeFile
    Open sImportFile For Input As #nImportFileNumber
    
    'Check that file is not empty
    If EOF(nImportFileNumber) Then
        CheckImportFile = ExchangeError.EmptyFile
        Close nImportFileNumber
        Exit Function
    End If
    
    'Mo Morris 16/12/99, Version extracted from header, but no checks performed on it yet
    If sMessageType = "STYDEF" Then
        'Check the contents of the Header record
        Input #nImportFileNumber, sSegmentId, sImportMessageType, dtTimeStamp, _
            sUIDSender, sUIDDestination, sVersion, lCount
    ElseIf sMessageType = "PATRSP" Then
        'Check the contents of the Header record
        Input #nImportFileNumber, sSegmentId, sImportMessageType, sClinicalTrialName, dtTimeStamp, _
            sUIDSender, sUIDDestination, sVersion, lCount
    ElseIf sMessageType = "LABDEF" Then
        'Check the contents of the Header record
        Input #nImportFileNumber, sSegmentId, sImportMessageType, dtTimeStamp, _
            sUIDSender, sUIDDestination, sVersion, lCount
    End If
    
    If sSegmentId <> "UNH" Then
        CheckImportFile = ExchangeError.Invalid
        Close nImportFileNumber
        Exit Function
    End If
    
    If sImportMessageType <> sMessageType Then
        CheckImportFile = ExchangeError.Invalid
        Close nImportFileNumber
        Exit Function
    End If
    
    If lCount <> 1 Then
        CheckImportFile = ExchangeError.Invalid
        Close nImportFileNumber
        Exit Function
    End If
    
    'Mo 21/11/2005 COD0210/COD0220
    'Note that the database before ClinicalCoding has version 3.0.50.
    'Note ClinicalCoding database has version 3.0.73.
    'Prevent imports from 3.0.73 into 3.0.50.
    If (sVersion = "3.0.73") And (GetVersionFromDB = "3.0.50") Then
        CheckImportFile = ExchangeError.MoreRecentVersion
        Close nImportFileNumber
        Exit Function
    End If
    
    If sMessageType = "STYDEF" Then
    
        'Extract the ClinicalTrialId, ClinicalName etc and check for:
        'duplication and version number
        If VersionNumberToLong(sVersion) <= VersionNumberToLong("2.2.9") Then
            'Actual REcruitment field existed up to this version
            Input #nImportFileNumber, sSegmentId, lClinicalTrialId, sClinicalTrialName, _
                sClinicalTrialDescription, nPhaseId, nStatusId, sKeywords, lExpectedRecruitment, _
                lActualRecruitment, nTrialTypeId, lCount
        Else
            'no actual recruitment field
            Input #nImportFileNumber, sSegmentId, lClinicalTrialId, sClinicalTrialName, _
                sClinicalTrialDescription, nPhaseId, nStatusId, sKeywords, lExpectedRecruitment, _
                nTrialTypeId, lCount
        End If
        If sSegmentId <> "001" Then
            CheckImportFile = ExchangeError.Invalid
            Close nImportFileNumber
            Exit Function
        End If
    End If
    
    'changed Mo Morris 5/4/00,  stemming from SR 3313
    'Make the following SQL statement non-case sensitive for Oracle databases
    ' NCJ 27/10/00 - Use new GetSQLStringEquals
    
    sSQL = " SELECT * FROM ClinicalTrial WHERE "
    sSQL = sSQL & GetSQLStringEquals("ClinicalTrialName", sClinicalTrialName)

'    If gUser.DatabaseType = MACRODatabaseType.Oracle80 Then
'        sSQL = " SELECT * FROM ClinicalTrial " _
'            & " WHERE NLS_LOWER(ClinicalTrialName) = '" & LCase(sClinicalTrialName) & "'"
'    Else
'        sSQL = " SELECT * FROM ClinicalTrial " _
'            & " WHERE ClinicalTrialName = '" & sClinicalTrialName & "'"
'    End If
    
    Set rsClinicalTrials = New ADODB.Recordset
    rsClinicalTrials.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsClinicalTrials.EOF And sMessageType = "PATRSP" Then
        CheckImportFile = ExchangeError.TrialDoesntExist
        Close nImportFileNumber
        Exit Function
    ElseIf Not rsClinicalTrials.EOF Then
        'set mlImportClinicalTrialId and msImportClinicalTrialName to the already existent trial
        mlImportClinicalTrialId = rsClinicalTrials!ClinicalTrialId
        msImportClinicalTrialName = rsClinicalTrials!ClinicalTrialName
    End If
    
    'Read through import file checking record numbers and the footer record
    Do While Not EOF(nImportFileNumber)
        'Increment  lCount ready for checking against the next record
        lCount = lCount + 1
        Line Input #nImportFileNumber, sImportRec
        'Changed Mo Morris 4/7/200 SR3661, method of extracting counter from record re-written
        'Extract line counter from record using the Split Function
        asFindCount = Split(sImportRec, ",")
        lNumberOfElements = UBound(asFindCount)
        sCurrentCount = asFindCount(lNumberOfElements)
    '   ATN 11/11/98
    '   Changed CInt to CLng to cope with large import files
        If CLng(sCurrentCount) <> lCount Then
            CheckImportFile = ExchangeError.Invalid
            Close nImportFileNumber
            Exit Function
        End If
        'Make a copy of the last checked record for additional footer checks
        sFooterRec = sImportRec
    Loop
    
    'create a record for the footer record to be checked against
    'changed Mo Morris 3/10/00, because of regional settings problem
    'changed Mo Morris 29/1/01, check for timestamps that happened exactly at midnight and
    'therefore do not have an hh:mm:ss part.
    If Len(CStr(CDbl(dtTimeStamp))) < 6 Then
        sFooterCheck = """UNT"",#" & Format(dtTimeStamp, "yyyy\-mm\-dd") & "#," & lCount
    Else
        sFooterCheck = """UNT"",#" & Format(dtTimeStamp, "yyyy\-mm\-dd hh\:mm\:ss") & "#," & lCount
    End If
    'sFooterCheck = """UNT"",#" & Format(dtTimeStamp, "yyyy-mm-dd hh:mm:ss") & "#," & lCount
    If sFooterRec <> sFooterCheck Then
        CheckImportFile = ExchangeError.Invalid
        Close nImportFileNumber
        Exit Function
    End If
    Close nImportFileNumber

    'TA 11/10/2000: new code to import data from old versions
    Set oNewDBColumns = New clsNewDBColumns
    'see if it needs converting
    If oNewDBColumns.Convert(sImportFile, sMessage) Then   'gsCAB_EXTRACT_LOCATION
        gLog gsIMPORT_UPGRADE, sMessage
    End If
     
    'Changed Mo Morris 23/1/2001
    'for patient data only call CleanUpPatientData if the cab file being imported is before version 2.1.19
    If sMessageType = "PATRSP" Then
        nVersionMajor = CInt(Split(sVersion, ".")(0))
        nVersionMinor = CInt(Split(sVersion, ".")(1))
        nVersionRevision = CInt(Split(sVersion, ".")(2))
        nVersionNumber = Val(Format(nVersionMajor, "00") & Format(nVersionMinor, "00") & Format(nVersionRevision, "0000"))
        'Mo Morris 30/1/2001, The bugs that require calls to CleanUpPatientData (i.e. invalid characters
        'in DataItemResponse.OveruleReason and invalid DataItemResponse.ResponseTimeStamps) where fixed in
        'version 2.1.30, hence there is no need to call it for PRD files created with this version or later.
        If nVersionNumber < 2010030 Then
            CleanUpPatientData (sImportFile)
        End If
    End If
    
    'MLM 24/06/02:
    mlImportVersion = VersionNumberToLong(sVersion)
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CheckImportFile", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Public Sub ExportStudyCAB(sCabFileName As String, asCABFileList() As String)
'---------------------------------------------------------------------
'Re-written by Mo Morris 23/2/00
'---------------------------------------------------------------------
Dim nFileNumber As Integer
Dim sCABfile As String
Dim i As Long

gLog gsEXPORT_STUDY_CAB, "Session " & sCabFileName & " starting"

'Create a ddf filename file containing all the items (.gif .pdf .sdd etc) to be compressed
nFileNumber = FreeFile
sCABfile = gsOUT_FOLDER_LOCATION & sCabFileName & ".ddf"

' DPH 17/10/2001 Make sure folder exists before opening
If FolderExistence(sCABfile) Then
    Open sCABfile For Output As nFileNumber
    
    'Place initial settings into the .ddf file
    Print #nFileNumber, ".OPTION EXPLICIT"
    Print #nFileNumber, ".Set CabinetNameTemplate=" & sCabFileName & ".cab"
    Print #nFileNumber, ".Set DiskDirectoryTemplate=" & gsOUT_FOLDER_LOCATION
    Print #nFileNumber, ".Set DiskLabel1=" & sCabFileName
    'Changed Mo Morris 3/7/2000 SR 3659
    Print #nFileNumber, ".Set MaxDiskSize=0"
    
    'Add the list of files to be compressed from asCABFileList to the .ddf file
    For i = 0 To UBound(asCABFileList)
        Print #nFileNumber, """" & asCABFileList(i) & """"
        'Debug.Print asCABFileList(i)
    Next i
    
    Close #nFileNumber
    
    'Change directory to where MakeCab.exe exists
    'Note that the DiskDirectoryTemplate location can either be a folder below gsAppPath
    'or as in the case of Roche a whole path location that could be anywhere
    
    ' DPH 17/10/2001 Make sure folder exists before opening
    Call FolderExistence(gsAppPath & "dummy.txt")
    
    ChDir gsAppPath
    
    'MakeCab runs under DOS and as such would normally be called via the Shell command.
    'Unfortunately the Shell command simply starts a process and then moves on to the next
    'line of code, which in this case would start the killing of the files to be CABed.
    'Instead the API utility ExecCmd, which completes the called process before moving onto
    'the next line,is used.
    'changed Mo Morris 7/6/2000 SR 3449 "COMMAND NUL /c " added for the purpose of hiding the DOS window
    'ExecCmd "COMMAND NUL /c Makecab /f " & """" & sCABfile & """"
    'Changed back Mo Morris 9/6/2000
    ExecCmd "Makecab /f " & """" & sCABfile & """"
    
    'Kill off the first file in asCABFileList which will be an SDD file
    'the rest of asCABFileList refers to Study Documents and grahics that should not be deleted
    Kill asCABFileList(0)
    
    'Kill the .ddf file
    Kill sCABfile
    
    gLog gsEXPORT_STUDY_CAB, "Session " & sCabFileName & " completed"
Else
    gLog gsEXPORT_STUDY_CAB, "Session failed as could not create CAB file directory " & gsOUT_FOLDER_LOCATION
End If
End Sub

'---------------------------------------------------------------------
Public Function ImportStudyDefinitionCAB(ByVal sImportFile As String)
'---------------------------------------------------------------------
' REVISIONS
' DPH 11/04/2002 - Branch if ZIP file from previous MACRO version
'---------------------------------------------------------------------
Dim sCommand As String
Dim sShortFormImportFile As String
Dim sShortFormDirectory As String

    ' DPH 11/04/2002 - Check if CAB or ZIP dealing with & act appropriately
    If UCase(Right(sImportFile, 3)) = "ZIP" Then
        Call ImportStudyDefinitionZIP(sImportFile)
        Exit Function
    End If
    
    gLog gsIMPORT_STUDY_CAB, "Extracting " & StripFileNameFromPath(sImportFile) & " starting"
    
    On Error Resume Next
    
    'clear the CabExtract folder before starting the Import
    ClearCabExtractFolder
    
    'Prepare a short form of the Extraction directory
    sShortFormDirectory = gsCAB_EXTRACT_LOCATION
    sShortFormDirectory = GetShortPathName(sShortFormDirectory)
    
    ' DPH 17/10/2001 - Make sure file exists before opening
    If FolderExistence(sImportFile, True) Then
        'Prepare a short form of the ImportFile
        sShortFormImportFile = GetShortPathName(sImportFile)
        
        'Call the CAB extract program which is in Macro's application path
        'The /Y paramater means that the user will not be prompted about potential overwrites
        'The /A paramater means process All files in the CAB file
        'The /E parameter means force the extraction to happen
        'The /L parameter together with sShortFormDirectory is where the files are extracted to
        sCommand = gsAppPath & "Extrac32 /Y /A /E /L " & """" & sShortFormDirectory & """ """ & sShortFormImportFile & """"
        
        ' DPH 11/04/2002 - Checking Extract process call
        If ExecCmd(sCommand) > 0 Then
            gLog gsIMPORT_STUDY_CAB, "Extracting " & StripFileNameFromPath(sImportFile) & " completed"
        Else
            gLog gsIMPORT_STUDY_CAB, "Extracting " & StripFileNameFromPath(sImportFile) & " failed. Process launch failure."
        End If
    Else
        gLog gsIMPORT_STUDY_CAB, "Extracting from " & gsCAB_EXTRACT_LOCATION & " failed as file doesn't exist"
    End If

End Function

'---------------------------------------------------------------------
Public Function ImportPatientCAB(ByVal sImportFile As String) As Boolean
'---------------------------------------------------------------------
' REVISIONS
' DPH 11/04/2002 - Check process launched properly & write appropriate log entry
'               Branch if ZIP file from previous MACRO version
' DPH 25/04/2002 - Made into function to return boolean
'---------------------------------------------------------------------
Dim sCommand As String
Dim sShortFormImportFile As String
Dim sShortFormDirectory As String
Dim sFileName As String

    ' DPH 11/04/2002 - Check if CAB or ZIP dealing with & act appropriately
    If UCase(Right(sImportFile, 3)) = "ZIP" Then
        ImportPatientCAB = ImportPatientZIP(sImportFile)
        Exit Function
    End If
    
    sFileName = StripFileNameFromPath(sImportFile)
    
    gLog gsIMPORT_PAT_CAB, "Extracting " & sFileName & " starting"
    
    On Error Resume Next
    
    'Changed Mo Morris 19/1/00, clear the CabExtract folder before starting the Import
    ClearCabExtractFolder
    
    'Prepare a short form of the Extraction directory
    'SDM 26/01/00 SR2794
    sShortFormDirectory = gsCAB_EXTRACT_LOCATION
    'sShortFormDirectory = gsAppPath & "CabExtract"
    sShortFormDirectory = GetShortPathName(sShortFormDirectory)
    
    ' DPH 17/10/2001 - Make sure file exists before opening
    If FolderExistence(sImportFile, True) Then
        'Prepare a short form of the ImportFile
        sShortFormImportFile = GetShortPathName(sImportFile)
        
        'sCommand = gsAppPath & "Extract /A " & """" & sImportFile & """" & " /E" & " /Y" '*1
        'Call the CAB extract program which is in Macro's application path
        'The /Y parameter means that the user will not be prompted about potential overwrites
        'The /A parameter means process All files in the CAB file
        'The /E parameter means force the extraction to happen
        'The /L parameter together with sShortFormDirectory is where the files are extracted to
        sCommand = gsAppPath & "Extrac32 /Y /A /E /L " & """" & sShortFormDirectory & """ """ & sShortFormImportFile & """"
        
        ' DPH 25/03/2002 - Check process launched properly & write appropriate log entry
        If ExecCmd(sCommand) > 0 Then
            gLog gsIMPORT_PAT_CAB, "Extracting " & sFileName & " completed"
            ImportPatientCAB = True
        Else
            gLog gsIMPORT_PAT_CAB, "Extracting " & sFileName & " failed. Process launch failure. AutoImport Error."
            ImportPatientCAB = False
        End If
    
    Else
        gLog gsIMPORT_PAT_CAB, "Extracting from " & gsCAB_EXTRACT_LOCATION & " failed as file doesn't exist"
        ImportPatientCAB = False
    End If

End Function

''------------------------------------------------------------------------------------------------------------------
'Private Function SetUpRepeatCommands(ByVal sTableName As String, ByRef nFieldCount As Integer, ByRef nPKItems As Integer, _
'                                    ByVal nExecute As eExecution) As ADODB.Command
''------------------------------------------------------------------------------------------------------------------
'' RJCW 11/12/2001 This function sets up a command object with prarameters for the Patient Import.
'' This is for performance improvements to Patient Import
'' parameters nFieldCount,nPKItems are output
''------------------------------------------------------------------------------------------------------------------
'' REVISIONS
'' DPH 22/07/2002 - Added scale to parameter types CBBL 2.2.19.38
'' MLM 24/07/02: Added optional paramter bIncludeLockStatus. This leaves nFieldCount and nPKItems return values
''               unchanged, but adds a parameter to the where clause for the LockStatus.
'' TA 28/08/2002: Changes so that primary key cols do not have to be the first in the table
'' NCJ 9 Jan 03 - Removed bIncludeLockStatus (Lock statuses now handled separately)
'' NCJ 22 Jan 03 - Added new StandardValue column
'' NCJ 29 Jan 03 - Added new UserNameFull column
'' RS 04/04/2003 - Routine no longer needed as new stored procedures are used
''------------------------------------------------------------------------------------------------------------------
'Dim cmdCommand  As New ADODB.Command
'Dim pFields     As New ADODB.Parameter
'Dim rsTable     As New ADODB.Recordset
'Dim VarArrayOf  As Variant
'Dim sPK()       As String
'Dim sSQL        As String
'Dim i           As Integer
'Dim j           As Integer
'Dim nRemove     As Integer
'Dim vParamNames As Variant
'
'   ' reset the parameters counter
'   nFieldCount = 0
'
'    ' get the table structure (number of fields) so the command object can be correctly constructed
'    If goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
'        sSQL = " SELECT * FROM " & sTableName & " WHERE ROWNUM = 1"
'        Set rsTable = New ADODB.Recordset
'        rsTable.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
'    Else
'        sSQL = " SELECT TOP 1 * FROM " & sTableName
'        Set rsTable = New ADODB.Recordset
'        rsTable.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
'    End If
'
'    nFieldCount = rsTable.Fields.Count
'
'    If nExecute = Update Then
'
'        'REM 27/08/02 - Hard coded the SQL
'        If sTableName = "DataItemResponse" Then
'
'            sSQL = "UPDATE DataItemResponse SET VisitId = ?,CRFPageId = ?,CRFElementId = ?,DataItemId = ?," _
'                 & "VisitCycleNumber = ?,CRFPageCycleNumber =?,CRFPageTaskId = ?,ResponseValue = ?,ResponseTimeStamp = ?," _
'                 & "ValueCode = ?,UserName = ?,UnitOfMeasurement = ?,Comments = ?,ResponseStatus = ?," _
'                 & "Changed = ?,SoftwareVersion = ?,ReasonForChange = ?,LockStatus = ?,ImportTimeStamp = ?,ValidationId = ?," _
'                 & "ValidationMessage = ?,OverruleReason = ?,LabResult = ?,CTCGrade = ?,ClinicalTestDate = ?," _
'                 & "LaboratoryCode = ?,HadValue = ?,ChangeCount = ?,DiscrepancyStatus = ?,SDVStatus = ?,NoteStatus = ?," _
'                 & "ResponseTimeStamp_TZ = ?, ImportTimeStamp_TZ = ?, DatabaseTimeStamp = ?, DatabaseTimeStamp_TZ = ?," _
'                 & "SequenceId = ?, StandardValue = ?, UserNameFull = ?" _
'                 & " WHERE ClinicalTrialId = ? AND TrialSite = ? AND PersonId = ? AND ResponseTaskId = ? AND RepeatNumber = ?"
'
'        ElseIf sTableName = "DataItemResponseHistory" Then
'
'            sSQL = "UPDATE DataItemResponseHistory SET VisitId = ?,CRFPageId = ?,CRFElementId = ?,DataItemId = ?," _
'                 & "VisitCycleNumber = ?,CRFPageCycleNumber =?,CRFPageTaskId = ?,ResponseValue = ?," _
'                 & "ValueCode = ?,UserName = ?,UnitOfMeasurement = ?,Comments = ?,ResponseStatus = ?," _
'                 & "Changed = ?,SoftwareVersion = ?,ReasonForChange = ?,LockStatus = ?,ImportTimeStamp = ?,ValidationId = ?," _
'                 & "ValidationMessage = ?,OverruleReason = ?,LabResult = ?,CTCGrade = ?,ClinicalTestDate = ?," _
'                 & "LaboratoryCode = ?,HadValue = ?," _
'                 & "ResponseTimeStamp_TZ = ?, ImportTimeStamp_TZ = ?, DatabaseTimeStamp = ?, DatabaseTimeStamp_TZ = ?," _
'                & "SequenceId = ?, StandardValue = ?, UserNameFull = ?" _
'                 & " WHERE ClinicalTrialId = ? AND TrialSite = ? AND PersonId = ? AND ResponseTaskId = ?" _
'                 & " AND ResponseTimeStamp = ? AND RepeatNumber = ?"
'
'        End If
'
''        'MLM 24/07/02:
''        If bIncludeLockStatus Then
''            sSQL = sSQL & " AND LockStatus = ? "
''        End If
'
'        'Oracle allows 'ROWNUM = 1' to force only on record to be updated
'        ' as we know that only one will be then this forces oracle to stop looking for others
'        ' once the first instance has been found
'        If goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then sSQL = sSQL & " AND ROWNUM = 1"
'
'        'nRemove = 0
'    Else
'        sSQL = "INSERT INTO " & sTableName & " VALUES( "
'
'        For i = 0 To nFieldCount - 1
'            sSQL = sSQL & "?,"
'        Next i
'
'        sSQL = Left(sSQL, Len(sSQL) - 1) & ")"
'        'nRemove = nPKItems
'    End If
'
'
'    'first collect array of field names in parameter order
'    'init array of names
'    ReDim vParamNames(nFieldCount) As String
'    For i = 0 To nFieldCount - 1
'        If nExecute = Update Then
'            j = ConvertFieldtoParameter(i, sTableName, nFieldCount)
'        Else
'            j = i
'        End If
'        vParamNames(j) = LCase$(rsTable.Fields(i).Name)
'    Next i
'
'    'add the right number of parameters
'    For i = 0 To nFieldCount - 1
'        Set pFields = New ADODB.Parameter
'        pFields.Type = 3 ' DUMMY TYPE THIS WILL BE CHANGED IN THE LOOP BELOW
'        pFields.Name = vParamNames(i)
'        cmdCommand.Parameters.Append pFields
'    Next i
'
'    'set up parameters according to the table column types
'    For i = 0 To nFieldCount - 1
'        If nExecute = Update Then
'            j = ConvertFieldtoParameter(i, sTableName, nFieldCount)
'        Else
'            j = i
'        End If
'        Set pFields = cmdCommand.Parameters(j)
'        pFields.Type = rsTable.Fields(i).Type
'        pFields.Precision = rsTable.Fields(i).Precision
'        ' DPH 22/07/2002 - Added scale to parameter types CBBL 2.2.19.38
'        pFields.NumericScale = rsTable.Fields(i).NumericScale
'        pFields.Direction = adParamInput
'        pFields.SIZE = rsTable.Fields(i).DefinedSize
'    Next i
'
'
''     'MLM 24/07/02:
''    If bIncludeLockStatus Then
''        'one extra parameter to set up for the LockStatus
''        Set pFields = New ADODB.Parameter
''        pFields.Type = rsTable.Fields("LockStatus").Type
''        pFields.Name = "WHERElockstatus"
''        pFields.Precision = rsTable.Fields("LockStatus").Precision
''        ' DPH 22/07/2002 - Added scale to parameter types CBBL 2.2.19.38
''        pFields.NumericScale = rsTable.Fields("LockStatus").NumericScale
''        pFields.Direction = adParamInput
''        pFields.Size = rsTable.Fields("LockStatus").DefinedSize
''        cmdCommand.Parameters.Append pFields
''    End If
'
'
'    ' define command object
'    Set cmdCommand.ActiveConnection = MacroADODBConnection
'    cmdCommand.CommandType = adCmdText
'    cmdCommand.CommandText = sSQL
'    cmdCommand.Prepared = True
'    Set pFields = Nothing
'    Set SetUpRepeatCommands = cmdCommand
'
'End Function

''---------------------------------------------------------------------
'Private Function ConvertFieldtoParameter(ByVal nFieldOrder As Integer, ByVal sTableName As String, ByVal nNoParameters As Integer) As Integer
''---------------------------------------------------------------------
''REM 28/08/02
''Converts the field order value of the DataItemResponse and DataItemResponseHistory tables in
'' to the equivilent parameter index value for the SQL statment
'' The field order and parameter index values are zero based
'' REM 26/09/02 - Added nNoParameters parameter to function
'' RS 04/04/2003 - Routine no longer needed as new stored procedures are used
''---------------------------------------------------------------------
'
'    If sTableName = "DataItemResponse" Then
'
'        Select Case nFieldOrder
'        Case 0 To 3 ' first 4 fields part of Primary Key, form part of WHERE clause
'            ConvertFieldtoParameter = nFieldOrder + (nNoParameters - 5) ' + 31
'        Case 4 To 30 ' Non Primary Key fields, form part of SET clause
'            ConvertFieldtoParameter = nFieldOrder - 4
'        Case 31 'Primary Key field, form part of WHERE clause
'            ConvertFieldtoParameter = nFieldOrder + (nNoParameters - 32)
'        Case Is >= 32 'Remaing Non Primary Key fields form part of SET clause
'            ConvertFieldtoParameter = nFieldOrder - 5 'no of primary keys
'        End Select
'
'    ElseIf sTableName = "DataItemResponseHistory" Then
'
'        Select Case nFieldOrder
'        Case 0 To 4 ' first 5 fields part of Primary Key, form part of WHERE clause
'            ConvertFieldtoParameter = nFieldOrder + (nNoParameters - 6) ' + 26
'        Case 5 To 30 ' Non Primary Key fields, form part of SET clause
'            ConvertFieldtoParameter = nFieldOrder - 5
'        Case 31 'repeat number
'            ConvertFieldtoParameter = nFieldOrder + (nNoParameters - 32)
'        Case Is >= 31 'Remaing Non Primary Key fields form part of SET clause
'            ConvertFieldtoParameter = nFieldOrder - 6 'no of primary keys
'        End Select
'
'    End If
'
'End Function

'---------------------------------------------------------------------
Private Function StripOutCRandLF(ByRef vStringForStripping As Variant) As Variant
'---------------------------------------------------------------------
'This function replaces Carriage Returns (ascii 13) with <13>
'This function replaces Line Feeds (ascii 10) with <10>
'---------------------------------------------------------------------

Dim msTempString As String
Dim msResultString As String
Dim mnPos As Long

On Error GoTo ErrHandler

'Loop through string replacing Chr(13) with "<13>"
msTempString = vStringForStripping
msResultString = ""
Do Until Len(msTempString) = 0
    mnPos = InStr(msTempString, Chr(13))
    If mnPos > 0 Then
        msResultString = msResultString & Left(msTempString, mnPos - 1) & "<13>"
        msTempString = Mid(msTempString, mnPos + 1)
    Else
        msResultString = msResultString & msTempString
        msTempString = ""
    End If
Loop

'Loop through string replacing Chr(10) with "<10>"
msTempString = msResultString
msResultString = ""
Do Until Len(msTempString) = 0
    mnPos = InStr(msTempString, Chr(10))
    If mnPos > 0 Then
        msResultString = msResultString & Left(msTempString, mnPos - 1) & "<10>"
        msTempString = Mid(msTempString, mnPos + 1)
    Else
        msResultString = msResultString & msTempString
        msTempString = ""
    End If
Loop

StripOutCRandLF = msResultString
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "StripOutCRandLF", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function
Private Function NormalizeParamName(sParamName As String) As String

    If LCase(Left(sParamName, 1)) = "@" Then
        ' This is a SQL Server SP Parameter
        NormalizeParamName = LCase(Mid(sParamName, 2, 99))
    Else
        ' This is an ORACLE SP Parameter
        NormalizeParamName = LCase(Mid(sParamName, 3, 99))
    End If

End Function


'---------------------------------------------------------------------
Private Function PutBackCRandLF(ByRef vStringForStripping As Variant) As Variant
'---------------------------------------------------------------------
'This function replaces <13> with Carriage Returns (ascii 13)
'This function replaces <10> with Line Feeds (ascii 10)
'---------------------------------------------------------------------

Dim msTempString As String
Dim msResultString As String
Dim mnPos As Long

On Error GoTo ErrHandler

'Loop through string replacing "<13>" with Chr(13)
msTempString = vStringForStripping
msResultString = ""
Do Until Len(msTempString) = 0
    mnPos = InStr(msTempString, "<13>")
    If mnPos > 0 Then
        msResultString = msResultString & Left(msTempString, mnPos - 1) & Chr(13)
        msTempString = Mid(msTempString, mnPos + 4)
    Else
        msResultString = msResultString & msTempString
        msTempString = ""
    End If
Loop

'Loop through string replacing "<10>" with Chr(10)
msTempString = msResultString
msResultString = ""
Do Until Len(msTempString) = 0
    mnPos = InStr(msTempString, "<10>")
    If mnPos > 0 Then
        msResultString = msResultString & Left(msTempString, mnPos - 1) & Chr(10)
        msTempString = Mid(msTempString, mnPos + 4)
    Else
        msResultString = msResultString & msTempString
        msTempString = ""
    End If
Loop

PutBackCRandLF = msResultString

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PutBackCRandLF", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Public Function AutoExportPRD(ByVal sSite As String, _
                              ByVal lClinicalTrialId As Long, _
                              ByVal sTrialSite As String, _
                              ByVal lPersonId As Long, _
                              Optional colFilesInCAB As Collection) As String
'---------------------------------------------------------------------
'AutoExportPRD started off as a copy of ExportPRD.
'
'AutoExportPRD will be geared towards exporting data marked as changed only and
'as such all references to vblnChangedData have been removed.
'
'Changed by Mo Morris 14/12/1999
'As per Macro 1.6 this sub is now passed a TrialId, a TrialSite and a PersonId and
'creates a Cab Export file one patient at a time
' DPH 11/04/2002 - Added CAB Validation info collection
'     Using ExportPatientZIP instead of ExportPatientCAB
' MLM 27/06/02: Use new character encoding.
' DPH 11/07/2002 - CBBL 2.2.19.25 Need Staticly cursored recordset for transactioned updates with SQL Server
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTrialAndSiteList As ADODB.Recordset
Dim rsExportImport As ADODB.Recordset
Dim rsCurrentTable As ADODB.Recordset
Dim sExportFile As String
Dim sExportFileName As String
Dim sTableName As String
Dim sSegmentId As String
Dim dtNow As Date
Dim lCount As Long
Dim i As Integer
Dim nExportFileNumber As Integer
Dim nNumFields As Integer
Dim vFieldData As Variant
Dim asFileList() As String
Dim nFileNumber As Integer
Dim sSessionName As String
Dim sProformaStateFile As String
Dim sProformaStateFileName As String
Dim nProformaStateFileNumber As Integer

Dim rsLocalData As ADODB.Recordset
Dim sLocalDataIds As String
Dim sDataId As String
Dim rsLocalCRFPageLabel As ADODB.Recordset
Dim sLocalCRFPageIds As String
Dim rsLocalTrialSubjectLabel As ADODB.Recordset
Dim sLocalTrialSubjectLabel As String
Dim sVersion As String

Dim rsMultiMedia As ADODB.Recordset

' DPH 25/03/2002 - Require variables for file validation
Dim lFileSize As Long

'REM 04/06/03 - Require trialname to add to Export File Name
Dim sTrialName As String

On Error GoTo ErrTransactionRollBack

'setup a recordset of all the trials/sites on file
sSQL = "SELECT DISTINCT TrialSubject.ClinicalTrialId, ClinicalTrial.ClinicalTrialName, " _
    & " TrialSubject.TrialSite, TrialSubject.Changed " _
    & " FROM TrialSubject, ClinicalTrial " _
    & " WHERE TrialSubject.ClinicalTrialId = ClinicalTrial.ClinicalTrialId " _
    & " AND TrialSubject.Changed = 1" _
    & " AND TrialSubject.ClinicalTrialid = " & lClinicalTrialId _
    & " AND TrialSubject.TrialSite = '" & sTrialSite & "'" _
    & " AND TrialSubject.PersonId = " & lPersonId

Set rsTrialAndSiteList = New ADODB.Recordset
rsTrialAndSiteList.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText

If rsTrialAndSiteList.EOF Then
    gLog gsAUTO_EXPORT_PRD, " Not started. No data to export"
    rsTrialAndSiteList.Close
    Set rsTrialAndSiteList = Nothing
    Exit Function
Else
    'dtNow is a date/time variable that is used as a time stamp that will be
    'placed in both the HEADER and FOOTER records of the export files
    'in addition the yyyymmddhhmmss parts of the date stamp are appended to the
    'export file names
    dtNow = Now
    
    'REM 04/06/03 - Get the Trial name
    sTrialName = rsTrialAndSiteList![ClinicalTrialName]
    
    'Note that SessionName includes the PersonId
    'REM 04/06/03 - Added Trial Name
    sSessionName = "PRD_" & sSite & "_" & sTrialName & "_" & Format(dtNow, "yyyymmddhhmmss") & "_" & lPersonId
    gLog gsAUTO_EXPORT_PRD, "Session " & sSessionName & " starting"
End If

'changed by Mo Morris 14/12/99
'Transaction control for AutoExport data now in frmExchangeStatus.DisplaySendMessages

rsTrialAndSiteList.MoveFirst

'initialise nFileNumber which will be used as an index in asFileList
nFileNumber = 0

Do While Not rsTrialAndSiteList.EOF

    lCount = 0
    'setup a file name and location for each trial/site combination
    sExportFileName = rsTrialAndSiteList![ClinicalTrialName] & "_" _
        & rsTrialAndSiteList![TrialSite] & "_" & Format(dtNow, "yyyymmddhhmmss") & ".prd"
    'SDM 26/01/00 SR2794
    sExportFile = gsOUT_FOLDER_LOCATION & sExportFileName
'    sExportFile = gsOUT_FOLDER_LOCATION & "\" & sExportFileName

    'redim the string array asFileList, into which the current file name will be added
    ReDim Preserve asFileList(nFileNumber)
    asFileList(nFileNumber) = sExportFile
    nFileNumber = nFileNumber + 1

    ' DPH 17/10/2001 Make sure folder exists before opening
    ' Existing error handling copes with any errors
    Call FolderExistence(sExportFile)
    
    'open export file and put a header record in it
    nExportFileNumber = FreeFile
    Open sExportFile For Output As #nExportFileNumber
    lCount = lCount + 1
    'Mo Morris 16/12/99, Version information added to header
    'Mo 20/02/2006 COD0210/COD0220, Version taken from Database instead of Software
    sVersion = GetVersionFromDB
    Write #nExportFileNumber, "UNH", "PATRSP", rsTrialAndSiteList![ClinicalTrialName], dtNow, "UID SENDER", "UID DESTINATION", sVersion, lCount

'   Select all data which is flagged as local
    sSQL = "SELECT CRFPageId,CRFElementId FROM CRFElement " _
        & " WHERE ClinicalTrialId = " & rsTrialAndSiteList!ClinicalTrialId _
        & "   AND Localflag = 1"
    
    Set rsLocalData = New ADODB.Recordset
    rsLocalData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    sLocalDataIds = ""
'   Build a string containing the ids of the local data items
    Do While Not rsLocalData.EOF
        sLocalDataIds = sLocalDataIds & rsLocalData!CRFPageId & "|" & rsLocalData!CRFElementId & ","
        rsLocalData.MoveNext
    Loop
    rsLocalData.Close
    Set rsLocalData = Nothing

'   Select all forms which are flagged as local
    sSQL = "SELECT CRFPageId FROM CRFPage " _
        & " WHERE ClinicalTrialId = " & rsTrialAndSiteList!ClinicalTrialId _
        & "   AND LocalCRFPageLabel = 1"
    
    Set rsLocalCRFPageLabel = New ADODB.Recordset
    rsLocalCRFPageLabel.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    sLocalCRFPageIds = ""
'   Build a string containing the ids of the local forms
    Do While Not rsLocalCRFPageLabel.EOF
        sLocalCRFPageIds = sLocalCRFPageIds & rsLocalCRFPageLabel!CRFPageId & ","
        rsLocalCRFPageLabel.MoveNext
    Loop
    rsLocalCRFPageLabel.Close
    Set rsLocalCRFPageLabel = Nothing
  
'   Check if trial subject label is local
    sSQL = "SELECT LocalTrialSubjectLabel FROM StudyDefinition " _
        & " WHERE ClinicalTrialId = " & rsTrialAndSiteList!ClinicalTrialId
    
    Set rsLocalTrialSubjectLabel = New ADODB.Recordset
    rsLocalTrialSubjectLabel.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsLocalTrialSubjectLabel!LocalTrialSubjectLabel = 0 Then
        sLocalTrialSubjectLabel = False
    Else
        sLocalTrialSubjectLabel = True
    End If
    rsLocalTrialSubjectLabel.Close
    Set rsLocalTrialSubjectLabel = Nothing

    'Changed Mo Morris 13/10/00, change from table PRDExportImport to MACROTable
    sSQL = "Select TableName, SegmentId From MACROTable" _
        & " WHERE PATRSP = 1" _
        & " ORDER BY SegmentId"
    Set rsExportImport = New ADODB.Recordset
    rsExportImport.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
    
    Do While Not rsExportImport.EOF
        sTableName = rsExportImport![TableName]
        sSegmentId = rsExportImport![SegmentId]

        'changed Mo Morris 26/1/00, SR 2788 filtering on PersonID Added
        sSQL = " SELECT * FROM " & sTableName _
            & " WHERE ClinicalTrialID = " & rsTrialAndSiteList![ClinicalTrialId] _
            & " AND TrialSite = '" & rsTrialAndSiteList![TrialSite] & "'" _
            & " AND PersonID = " & lPersonId _
            & " AND Changed = 1"

        'TA 31/10/2005: Order by timestamp and timezone if dirh - so that they are imported
        ' in the correct order on the server (so the get allocated correct sequence ids for CG)
        If LCase(sTableName) = LCase("DataItemResponseHistory") Then
            'get them input order
            If goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
                '1440 minutes in a day
                sSQL = sSQL & " order by responsetimestamp + (responsetimestamp_tz/1440)"
            Else
                sSQL = sSQL & " order by responsetimestamp + (responsetimestamp_tz/cast(1440 as double precision))"
            End If
        End If
        Set rsCurrentTable = New ADODB.Recordset
        rsCurrentTable.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
        
        Do While Not rsCurrentTable.EOF
        'Check for local data items and don't export them
            If sLocalDataIds > "" Then
                If sTableName = "DataItemResponse" Or sTableName = "DataItemResponseHistory" Then
                    sDataId = rsCurrentTable!CRFPageId & "|" & rsCurrentTable!CRFElementId & ","
                    Do While InStr(sLocalDataIds, sDataId) > 0
                        rsCurrentTable.MoveNext
                        If rsCurrentTable.EOF Then
                            Exit Do
                        End If
                        sDataId = rsCurrentTable!CRFPageId & "|" & rsCurrentTable!CRFElementId & ","
                    Loop
                    If rsCurrentTable.EOF Then
                        Exit Do
                    End If
                End If
            End If

            nNumFields = rsCurrentTable.Fields.Count
            lCount = lCount + 1
            Write #nExportFileNumber, sSegmentId;
            'write each data item to the export file one at a time
            For i = 0 To (nNumFields - 1)
                'check for the dataitem's 'ProformaState' which contains binary data and
                'hence can not be written to file with the Write command
                '   ATN 26/1/2000   Must convert case, because Oracle doesn't preserve case in field names
                If LCase$(rsCurrentTable.Fields(i).Name) = "proformastate" Then
                    'create a separate file for the ProformaState data and
                    'add it to the list of files included in this export session
                    'note that ProformaState will pertain to the PersonId held in
                    'rsCurrentTable.Fields(2).Value
                    sProformaStateFileName = rsTrialAndSiteList![ClinicalTrialName] & "_" _
                        & rsTrialAndSiteList![TrialSite] & "_" & rsCurrentTable.Fields(2).Value & ".psf"
                    'SDM 26/01/00 SR2794
                    sProformaStateFile = gsOUT_FOLDER_LOCATION & sProformaStateFileName
'                    sProformaStateFile = gsOUT_FOLDER_LOCATION & "\" & sProformaStateFileName
                    'add sProformaStateFile to asFileList
                    ReDim Preserve asFileList(nFileNumber)
                    asFileList(nFileNumber) = sProformaStateFile
                    nFileNumber = nFileNumber + 1
                    ' DPH 17/10/2001 Make sure folder exists before opening
                    ' Let existing error handling control errors
                    Call FolderExistence(sProformaStateFile)
                    
                    'put proformaState data into sProformaStateFile
                    nProformaStateFileNumber = FreeFile
                    Open sProformaStateFile For Output As nProformaStateFileNumber
                    Print #nProformaStateFileNumber, rsCurrentTable.Fields(i).Value
                    Close #nProformaStateFileNumber
                    gLog gsAUTO_EXPORT_PRD, sProformaStateFileName & " ProformaState file created"
                    'Store the name of the ProformaState file in the .prd file
                    Write #nExportFileNumber, sProformaStateFileName;
                Else
                    If rsCurrentTable.Fields(i).Attributes And adFldLong Then
                        vFieldData = BlobToVariant(rsCurrentTable.Fields(i), rsCurrentTable.Fields(i).ActualSize)
                    Else
                        vFieldData = rsCurrentTable.Fields(i).Value
                    End If
                    '   Check for local CRFPageLabels and don't export them
                    '   ATN 26/1/2000   Must convert case, because Oracle doesn't preserve case in field names
                    If LCase$(rsCurrentTable.Fields(i).Name) = "crfpageinstancelabel" Then
                        If sLocalCRFPageIds > "" Then
                            If InStr(sLocalCRFPageIds, sLocalCRFPageIds) > 0 Then
                                vFieldData = ""
                            End If
                        End If
                    '   ATN 26/1/2000   Must convert case, because Oracle doesn't preserve case in field names
                    ElseIf LCase$(rsCurrentTable.Fields(i).Name) = "localidentifier1" Then
                        If sLocalTrialSubjectLabel Then
                            vFieldData = ""
                        End If
                    End If
                    'MLM 24/06/02: Use new string encoding
                    If VarType(vFieldData) = vbString Then
                        vFieldData = HexEncodeChars(vFieldData, msCHARS_TO_ENCODE)
                    End If
                    'check for vFieldData containing carriage returns
'                    If (InStr(vFieldData, Chr(13)) <> 0) Or (InStr(vFieldData, Chr(10)) <> 0) Then
'                        vFieldData = StripOutCRandLF(vFieldData)
'                    End If
                    Write #nExportFileNumber, vFieldData;
                End If
            Next i
            Write #nExportFileNumber, lCount
            rsCurrentTable.MoveNext
        Loop
        'Update the Changed Flag code used to be here
        rsExportImport.MoveNext
    Loop    'Looping on rsExportImport
    
    lCount = lCount + 1
    Write #nExportFileNumber, "UNT", dtNow, lCount
    Close #nExportFileNumber
    gLog gsAUTO_EXPORT_PRD, sExportFileName & " Export file created"
    
    'query the DataItemResponseHistory table for dataitems of type multimedia and add them to the export
    sSQL = "SELECT DISTINCT ResponseValue FROM DataItemResponseHistory, DataItem " _
        & " WHERE DataItemResponseHistory.ClinicalTrialId = DataItem.ClinicalTrialId " _
        & " AND DataItemResponseHistory.DataItemId = DataItem.DataItemId " _
        & " AND DataItemResponseHistory.ClinicalTrialId = " & rsTrialAndSiteList!ClinicalTrialId _
        & " AND DataItemResponseHistory.TrialSite = '" & rsTrialAndSiteList!TrialSite & "'" _
        & " AND dataitemresponsehistory.PersonId = " & lPersonId _
        & " AND DataItem.DataType = " & DataType.Multimedia _
        & " AND DataItemResponseHistory.Changed = " & Changed.Changed
    Set rsMultiMedia = New ADODB.Recordset
    ' DPH 11/07/2002 - CBBL 2.2.19.25 Need Staticly cursored recordset for transactioned updates with SQL Server
'    rsMultiMedia.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    rsMultiMedia.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not rsMultiMedia.EOF
        If FileExists(gsDOCUMENTS_PATH & rsMultiMedia!ResponseValue) Then
            'check that this multimedia file has not already been added to this export
            'i.e. a patients pointing at the same multimedia file more than once
            If Not FileExists(gsOUT_FOLDER_LOCATION & rsMultiMedia!ResponseValue) Then
                'copy file to out folder and add to list of files to be exported
                FileCopy gsDOCUMENTS_PATH & rsMultiMedia!ResponseValue, gsOUT_FOLDER_LOCATION & rsMultiMedia!ResponseValue
                ReDim Preserve asFileList(nFileNumber)
                asFileList(nFileNumber) = gsOUT_FOLDER_LOCATION & rsMultiMedia!ResponseValue
                nFileNumber = nFileNumber + 1
                gLog gsAUTO_EXPORT_PRD, rsMultiMedia!ResponseValue & " MultiMedia file added to export"
            End If
        End If
        rsMultiMedia.MoveNext
    Loop
    rsMultiMedia.Close
    Set rsMultiMedia = Nothing
    
    rsTrialAndSiteList.MoveNext
 
Loop    'Looping on rsTrialAndSiteList

rsTrialAndSiteList.Close
Set rsTrialAndSiteList = Nothing
rsExportImport.Close
Set rsExportImport = Nothing
rsCurrentTable.Close
Set rsCurrentTable = Nothing

' DPH 11/04/2002 - Create File list + file sizes
' Initialise Collection if required
Set colFilesInCAB = New Collection
' Loop through File List & collect file sizes
For i = 0 To UBound(asFileList)
    lFileSize = GetFileLength(asFileList(i))
    ' Store in collection
    colFilesInCAB.Add lFileSize, StripFileNameFromPath(asFileList(i))
Next

' DPH 11/04/2002 - Using ExportPatientZIP instead of ExportPatientCAB
'ExportPatientCAB sSessionName, asFileList
ExportPatientZIP sSessionName, asFileList

'   Return the file name to the calling routine
' DPH 15/04/2002 - Removed CAB reference
'AutoExportPRD = sSessionName & ".cab"
AutoExportPRD = sSessionName & ".zip"

gLog gsAUTO_EXPORT_PRD, "Session " & sSessionName & " completed"
Exit Function

ErrTransactionRollBack:

gLog gsAUTO_EXPORT_PRD, "Error code " & Err.Number & " - " & Err.Description
gLog gsAUTO_EXPORT_PRD, "Export Aborted. All updates will be Rolled Back."

'remove the .prd and .psf files created by this erroneous export session
Close #nExportFileNumber
For i = 0 To UBound(asFileList)
    Kill asFileList(i)
    gLog gsAUTO_EXPORT_PRD, "Removing after error " & StripFileNameFromPath(asFileList(i))
Next i

End Function

'---------------------------------------------------------------------
Public Sub UpdateChangedFlags(ByVal lClinicalTrialId As Long, _
                              ByVal sTrialSite As String, _
                              ByVal lPersonId As Long)
'---------------------------------------------------------------------
'Changed by Mo Morris 14/12/1999
'As per Macro 1.6 this sub is now called once per patient that has been
'succesfully exported by AutoExportPRD
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsExportImport As ADODB.Recordset
'Dim rsCurrentTable As ADODB.Recordset
Dim sTableName As String

    On Error GoTo ErrHandler
        
    'Changed Mo Morris 13/10/00, change from table PRDExportImport to MACROTable
    sSQL = "Select TableName, SegmentId From MACROTable" _
        & " WHERE PATRSP = 1" _
        & " ORDER BY SegmentId"
    Set rsExportImport = New ADODB.Recordset
    'changed by Mo Morris 14/12/99, cursor type changed from adOpenForwardOnly to adOpenStatic
    rsExportImport.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
    
    Do While Not rsExportImport.EOF
        sTableName = rsExportImport![TableName]
        'Update the Changed Flag from 1 to 0 in the originating record
        sSQL = "UPDATE " & sTableName & " SET " _
            & " Changed = 0 " _
            & " WHERE  Changed = 1" _
            & " AND " & sTableName & ".ClinicalTrialid = " & lClinicalTrialId _
            & " AND " & sTableName & ".TrialSite = '" & sTrialSite & "'" _
            & " AND " & sTableName & ".PersonId = " & lPersonId
        MacroADODBConnection.Execute sSQL
        rsExportImport.MoveNext
    Loop
    rsExportImport.Close
    Set rsExportImport = Nothing
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpdateChangedFlags", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub

'---------------------------------------------------------------------
Public Function AutoImportPRD(ByVal sImportFile As String)
'---------------------------------------------------------------------
'Mo Morris 28/4/00, SR3249, changes made so that the routine exits properly
'when an error occurs during an export.
'Mo Morris 16/01/01, Changed AutoImportPRD from a sub to a function
'DPH 11/04/2002 - Check PRD files exist after extract and write log entry if error
' NCJ 14 Jan 03 - Store zip file name in msCurrentZipFileName (needed for Lock/Freeze processing)
'---------------------------------------------------------------------
Dim sNextPRDFile As String
Dim sFileName As String

    On Error GoTo ErrHandler
    
    sFileName = StripFileNameFromPath(sImportFile)
    
    ' NCJ 14 Jan 03 - When we get here the file name has an underscore at the front
    ' so we need to remove it to get the REAL zip file name
    msCurrentZipFileName = Right(sFileName, Len(sFileName) - 1)
    
    gLog gsAUTO_IMPORT_PRD, "Session " & sFileName & " starting"
    'Unpack the CAB file into .prd and .psf files and place them
    'in directory AppPath/CabExtract
    If Not ImportPatientCAB(sImportFile) Then
        AutoImportPRD = ExchangeError.DecompressFail
        Exit Function
    End If
    
    'loop through the extracted files and import them into Macro
    'SDM 26/01/00 SR22794
    sNextPRDFile = Dir(gsCAB_EXTRACT_LOCATION & "*.prd")
    'sNextPRDFile = Dir(gsAppPath & "CabExtract" & "\*.prd")
    
    ' DPH 11/02/2002 - Check if PRD file exists, if not write error to log
    If sNextPRDFile <> "" Then
        Do While sNextPRDFile <> ""
            gLog gsAUTO_IMPORT_PRD, "Importing " & sNextPRDFile
            'SDM 26/01/00 SR2794
            Select Case ImportPRD(gsCAB_EXTRACT_LOCATION & sNextPRDFile)
        '    Select Case ImportPRD(gsAppPath & "CabExtract" & "\" & sNextPRDFile)
            Case ExchangeError.EmptyFile
                gLog gsAUTO_IMPORT_PRD, sNextPRDFile & " was empty. Import aborted"
                'Changed Mo Morris 16/01/01
                AutoImportPRD = ExchangeError.EmptyFile
                Exit Function
            Case ExchangeError.Invalid
                gLog gsAUTO_IMPORT_PRD, sNextPRDFile & " invalid file. Import aborted"
                'Changed Mo Morris 16/01/01
                AutoImportPRD = ExchangeError.Invalid
                Exit Function
            ' DPH 11/04/2002
            Case ExchangeError.SubjectLock
                gLog gsAUTO_IMPORT_PRD, sNextPRDFile & " encountered a subject lock. Import aborted"
                AutoImportPRD = ExchangeError.SubjectLock
                Exit Function
            Case ExchangeError.Success
                gLog gsAUTO_IMPORT_PRD, sNextPRDFile & " imported successfully"
            Case Else
                gLog gsAUTO_IMPORT_PRD, sNextPRDFile & " unexpected error. Import aborted"
                'Changed Mo Morris 16/01/01
                AutoImportPRD = ExchangeError.Unknown
                Exit Function
            End Select
            'get next prd file via the DIR command
            sNextPRDFile = Dir
        Loop
    Else
        gLog gsAUTO_IMPORT_PRD, "No PRD files extracted from " & sFileName & " . AutoImport Error."
        ' DPH 25/04/2002 - Added to catch No PRD Error
        AutoImportPRD = ExchangeError.NoPRDFile
        Exit Function
    End If
    
    Kill sImportFile
    
    Do Until Not FileExists(sImportFile)
        DoEvents
    Loop
    gLog gsAUTO_IMPORT_PRD, "Session " & sFileName & " completed and deleted"
    
    'changed Mo Moris 2/3/2000
    'move any patient data multimedia files from the CabExtract folder to the Documents folder
    ImportDocumentsAndGraphics
    
    'Changed Mo Morris 16/01/01
    AutoImportPRD = ExchangeError.Success

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "AutoImportPRD", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Function ImportTrialPhase(nImportFileNumber As Integer, _
                                lStudyPhaseID As Long) As Long
'---------------------------------------------------------------------
'Mo Morris  10/3/00 Cancel/Abort option removed
'ASH 21/1/2003 changes to fix dupicate phasename error
'---------------------------------------------------------------------
Dim nPhaseId As Integer
Dim sPhaseName As String
Dim lDummyReadOfLineCounter As Long
Dim sSQL As String
Dim rsTrialPhase As ADODB.Recordset
Dim rsPhaseID As ADODB.Recordset
Dim lMaxPhaseID As Long
Dim rsMaxPhaseID As ADODB.Recordset
Dim sPrompt As String
Dim nResponse As Integer
    
    On Error GoTo ErrHandler
    
    Input #nImportFileNumber, nPhaseId, sPhaseName, lDummyReadOfLineCounter
    'MLM 06/01/03: Decode the strings before displaying to the user/storing in db.
    If mbUseHexDecoding Then
        sPhaseName = HexDecodeChars(sPhaseName)
    End If
    
    'get the max phaseid
    sSQL = "SELECT   max(Phaseid)  as NewPhaseId FROM TrialPhase"
    Set rsMaxPhaseID = New ADODB.Recordset
    rsMaxPhaseID.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If IsNull(rsMaxPhaseID!NewPhaseId) Then
        lMaxPhaseID = 1
    Else
        lMaxPhaseID = rsMaxPhaseID!NewPhaseId
    End If
    rsMaxPhaseID.Close
    Set rsMaxPhaseID = Nothing
    
    'get the phasenames from the database that match the phasename from the file
    'if it does not exist then put it in the database with a new phaseid
    sSQL = "SELECT * FROM TrialPhase WHERE PhaseName = '" & sPhaseName & "'"
    Set rsTrialPhase = New ADODB.Recordset
    rsTrialPhase.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    If rsTrialPhase.RecordCount <= 0 Then
        rsTrialPhase.AddNew
        rsTrialPhase.Fields(0) = lMaxPhaseID + 1
        rsTrialPhase.Fields(1) = sPhaseName
        rsTrialPhase.Update
        'update clinicaltrial table with the new phaseid
        If lStudyPhaseID = nPhaseId Then
            sSQL = "UPDATE ClINICALTRIAL SET PHASEID = " & lMaxPhaseID + 1 & _
            " WHERE ClinicalTrialId = " & mlImportClinicalTrialId
            MacroADODBConnection.Execute sSQL
        End If
    Else
        'update clinicaltrial table with the old phaseid
        If lStudyPhaseID = nPhaseId Then
            sSQL = "UPDATE ClINICALTRIAL SET PHASEID = " & rsTrialPhase.Fields("PHASEID").Value & _
            " WHERE ClinicalTrialId = " & mlImportClinicalTrialId
            MacroADODBConnection.Execute sSQL
        End If
    End If
     
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportTrialPhase", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Function ImportTrialType(nImportFileNumber As Integer, _
                                lNewTrialId As Long) As Long
'---------------------------------------------------------------------
'Mo Morris  10/3/00 Cancel/Abort option removed
'MLM 30/11/01: SR4260 Largely rewritten with new logic: Only add a row to TrialType if the reqd
'   TrialTypeName does not exist, otherwise look up its ID. Then update the study with this ID.
'---------------------------------------------------------------------
Dim nTrialTypeId As Integer
Dim sTrialTypeName As String
Dim lDummyReadOfLineCounter As Long
Dim sSQL As String
Dim rsTrialType As ADODB.Recordset
Dim lMaxId As Long
    
    On Error GoTo ErrHandler
        
    Input #nImportFileNumber, nTrialTypeId, sTrialTypeName, lDummyReadOfLineCounter
    'MLM 06/01/03: Decode strings if required
    If mbUseHexDecoding Then
        sTrialTypeName = HexDecodeChars(sTrialTypeName)
    End If
    
    Set rsTrialType = New ADODB.Recordset
    sSQL = "SELECT * FROM TrialType"
    With rsTrialType
        .Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
        
        'search through existing TrialType table for one of the same name
        'while we're there, look for the max TrialTypeId
        lMaxId = 0
        Do Until .EOF
            If LCase(sTrialTypeName) = LCase(.Fields("TrialTypeName").Value) Then
                Exit Do
            End If
            If .Fields("TrialTypeId").Value > lMaxId Then
                lMaxId = .Fields("TrialTypeId").Value
            End If
            .MoveNext
        Loop
        
        If .EOF Then
            'this TrialType doesn't exist; create it
            .AddNew
            .Fields("TrialTypeId").Value = lMaxId + 1
            .Fields("TrialTypeName").Value = sTrialTypeName
            .Update
        End If
        
        'rsTrialType will now point to the row matching the Type we're trying to import
        'If the study we're importing is of this Type, and the ID of this type is now different from in
        'the original database, update the ClinicalTrial row with the new ID.
        MacroADODBConnection.Execute "UPDATE ClinicalTrial SET TrialTypeId = " & .Fields("TrialTypeId").Value & _
            " WHERE ClinicalTrialId = " & lNewTrialId & " AND TrialTypeId = " & nTrialTypeId
        
        .Close
    End With
    
    Set rsTrialType = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportTrialType", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Function ImportValidationAction(nImportFileNumber As Integer) As Long
'---------------------------------------------------------------------
'Mo Morris  10/3/00 Cancel/Abort option removed
'---------------------------------------------------------------------
Dim nValidationActionId As Integer
Dim sValidationActionName As String
Dim lDummyReadOfLineCounter As Long
Dim sSQL As String
Dim rsValidationAction As ADODB.Recordset
Dim sPrompt As String
Dim nResponse As Integer
    
    On Error GoTo ErrHandler
        
    Input #nImportFileNumber, nValidationActionId, sValidationActionName, lDummyReadOfLineCounter
    'MLM 06/01/03: Decode strings if required
    If mbUseHexDecoding Then
        sValidationActionName = HexDecodeChars(sValidationActionName)
    End If
    
    sSQL = " SELECT * FROM ValidationAction WHERE ValidationActionId = " & nValidationActionId
    Set rsValidationAction = New ADODB.Recordset
    'changed Mo 28/1/00, from adOpenDynamic to adOpenKeyset
    rsValidationAction.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    'assertain as to wether its a new code
    If rsValidationAction.RecordCount = 0 Then
        'Its a new code
        rsValidationAction.AddNew
    Else
        'Its an existing code
        'rsValidationAction.Edit
        'If running under MACRO_DM the code will be automatically changed
        'If running under MACRO_EX the current code will be compared to the imported one
        'and if they differ the user will be prompted as to what they want to happen
        If App.Title = "MACRO_SM" Then
            If sValidationActionName <> rsValidationAction![ValidationActionName] Then
                sPrompt = "The import wants to change the meaning of ValidationAction code [" & nValidationActionId _
                    & "] from [" & rsValidationAction![ValidationActionName] & "] to [" & sValidationActionName & "]." _
                    & vbNewLine & "Yes will accept the change to [" & sValidationActionName & "]." _
                    & vbNewLine & "No will keep the code as [" & rsValidationAction![ValidationActionName] & "]."
                nResponse = MsgBox(sPrompt, vbYesNo + vbExclamation + vbDefaultButton1, _
                    gsDIALOG_TITLE)
                If nResponse = vbYes Then
                    'Accept the overwrite
                Else
                    'Don't accept the overwrite (set it back to itself)
                    sValidationActionName = rsValidationAction![ValidationActionName]
                End If
            End If
        End If
    End If
    rsValidationAction.Fields(0) = nValidationActionId
    rsValidationAction.Fields(1) = sValidationActionName
    rsValidationAction.Update
    rsValidationAction.Close
    Set rsValidationAction = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportValidationAction", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Function ImportValidationType(nImportFileNumber As Integer) As Long
'---------------------------------------------------------------------
'Mo Morris  10/3/00 Cancel/Abort option removed
'---------------------------------------------------------------------
Dim nValidationTypeId As Integer
Dim sValidationTypeName As String
Dim nValidationActionId As Integer
Dim lDummyReadOfLineCounter As Long
Dim sSQL As String
Dim rsValidationType As ADODB.Recordset
Dim sPrompt As String
Dim nResponse As Integer

    On Error GoTo ErrHandler

    Input #nImportFileNumber, nValidationTypeId, sValidationTypeName, nValidationActionId, lDummyReadOfLineCounter
    'MLM 06/01/03: Decode strings if required
    If mbUseHexDecoding Then
        sValidationTypeName = HexDecodeChars(sValidationTypeName)
    End If
    
    sSQL = " SELECT * FROM ValidationType WHERE ValidationTypeId = " & nValidationTypeId
    Set rsValidationType = New ADODB.Recordset
    'changed Mo 28/1/00, from adOpenDynamic to adOpenKeyset
    rsValidationType.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    'assertain as to wether its a new code
    If rsValidationType.RecordCount = 0 Then
        'Its a new code
        rsValidationType.AddNew
    Else
        'Its an existing code
        'rsValidationType.Edit
        'If running under MACRO_DM the code will be automatically changed
        'If running under MACRO_EX the current code will be compared to the imported one
        'and if they differ the user will be prompted as to what they want to happen
        If App.Title = "MACRO_SM" Then
            If sValidationTypeName <> rsValidationType![ValidationTypeName] _
                Or nValidationActionId <> rsValidationType![ValidationActionId] Then
                sPrompt = "The import wants to change the meaning of ValidationType code [" & nValidationTypeId _
                    & "] from [" & rsValidationType![ValidationTypeName] & "," & rsValidationType![ValidationActionId] _
                    & "] to [" & sValidationTypeName & "," & nValidationActionId & "]." _
                    & vbNewLine & "Yes will accept the change to [" & sValidationTypeName & "]." _
                    & vbNewLine & "No will keep the code as [" & rsValidationType![ValidationTypeName] & "]."
                nResponse = MsgBox(sPrompt, vbYesNo + vbExclamation + vbDefaultButton1, _
                    gsDIALOG_TITLE)
                If nResponse = vbYes Then
                    'Accept the overwrite
                Else
                    'Don't accept the overwrite (set it back to itself)
                    sValidationTypeName = rsValidationType![ValidationTypeName]
                End If
            End If
        End If
    End If
    rsValidationType.Fields(0) = nValidationTypeId
    rsValidationType.Fields(1) = sValidationTypeName
    rsValidationType.Fields(2) = nValidationActionId
    rsValidationType.Update
    rsValidationType.Close
    Set rsValidationType = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportValidationType", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Function ImportUnits(nImportFileNumber As Integer) As Long
'---------------------------------------------------------------------
' NCJ 25/10/00 - Import Units table
' NCJ 27/10/00 - Use case-insensitive comparisons for Oracle
'---------------------------------------------------------------------
Dim sUnit As String
Dim sUnitClass As String
Dim lDummyReadOfLineCounter As Long
Dim sSQL As String
Dim rsUnits As ADODB.Recordset
Dim sPrompt As String
Dim nResponse As Integer

    On Error GoTo ErrHandler
    
    ' Read a record
    Input #nImportFileNumber, sUnit, sUnitClass, lDummyReadOfLineCounter
    'MLM 06/01/03: Decode strings if required
    If mbUseHexDecoding Then
        sUnit = HexDecodeChars(sUnit)
        sUnitClass = HexDecodeChars(sUnitClass)
    End If
    
    sSQL = " SELECT * FROM Units WHERE " & GetSQLStringEquals("Unit", sUnit)
    Set rsUnits = New ADODB.Recordset
    rsUnits.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    ' Is it a new unit?
    If rsUnits.RecordCount = 0 Then
        'It's a new unit
        rsUnits.AddNew
    Else
        'It's an existing code
        'If running under MACRO_DM the code will be automatically changed
        'If running under MACRO_EX the user will be prompted as to what they want to happen
        If App.Title = "MACRO_SM" Then
            If sUnitClass <> rsUnits![UnitClass] Then
                sPrompt = "The import wants to change the class of unit '" & sUnit _
                    & "' from '" & rsUnits![UnitClass] & "' to '" & sUnitClass & "'." _
                    & vbNewLine & "Would you like to make this change?" _
                    & vbNewLine & vbNewLine _
                    & "WARNING If you change the unit class, any existing conversion factors " _
                    & vbNewLine & "defined for this unit will also be removed."
                If DialogQuestion(sPrompt) = vbYes Then
                    ' Accept the overwrite
                    ' But we must splat any conversion factors
                    Call RemoveConversionFactors(sUnit)
                Else
                    'Don't accept the overwrite (set it back to itself)
                    sUnitClass = rsUnits![UnitClass]
                End If
            End If      ' If different unit
        End If      ' If Exchange (manual import)
    End If
    rsUnits.Fields(0) = sUnit
    rsUnits.Fields(1) = sUnitClass
    rsUnits.Update
    rsUnits.Close
    Set rsUnits = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportUnits", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Sub RemoveConversionFactors(ByVal sUnit As String)
'---------------------------------------------------------------------
' NCJ 25/10/00 - Remove all conversion factors involving this unit
' ncj 27/10/00 - Use case-insensitive matching for Oracle
'---------------------------------------------------------------------
Dim sSQL As String

    sSQL = "DELETE FROM UnitConversionFactors WHERE "
    sSQL = sSQL & GetSQLStringEquals("FromUnit", sUnit)
    sSQL = sSQL & " OR " & GetSQLStringEquals("ToUnit", sUnit)
    
    MacroADODBConnection.Execute sSQL

End Sub

'---------------------------------------------------------------------
Private Function ImportConversionFactors(nImportFileNumber As Integer) As Long
'---------------------------------------------------------------------
' NCJ 25/10/00 - Import UnitConversionFactors table
' NB We don't check ConversionExpressions (which are currently unused)
'---------------------------------------------------------------------
Dim sToUnit As String
Dim sFromUnit As String
Dim sUnitClass As String
Dim dblConvFactor As Double
'Dim vConvExpr As Variant
Dim lDummyReadOfLineCounter As Long
Dim sSQL As String
Dim rsUnitConversions As ADODB.Recordset
Dim sPrompt As String

    On Error GoTo ErrHandler
    
    ' Read a record
    'Mo Morris 30/8/01 Db Audit (ConversionExpression dropped from table UnitConversionFactors)
    Input #nImportFileNumber, sFromUnit, sToUnit, dblConvFactor, sUnitClass, lDummyReadOfLineCounter
    'MLM 06/01/03: Decode strings if required
    If mbUseHexDecoding Then
        sFromUnit = HexDecodeChars(sFromUnit)
        sToUnit = HexDecodeChars(sToUnit)
        sUnitClass = HexDecodeChars(sUnitClass)
    End If
    
    ' For manual import, check that what's coming in is consistent
    ' (in case the user previously chose not to import some units)
    If App.Title = "MACRO_SM" Then
        ' Check it's valid
        If Not DoesUnitClassMatch(sFromUnit, sToUnit, sUnitClass) Then
            ' We have to ignore this record because the unit class is wrong
            sPrompt = "The conversion factor from '" & sFromUnit & "' to '" & sToUnit & "'" _
                        & vbNewLine & "cannot be imported because the unit classes are inconsistent"
            Call DialogError(sPrompt)
            Exit Function
        End If
    End If
    
    sSQL = " SELECT * FROM UnitConversionFactors WHERE "
    sSQL = sSQL & GetSQLStringEquals("FromUnit", sFromUnit)
    sSQL = sSQL & " AND " & GetSQLStringEquals("ToUnit", sToUnit)
    Set rsUnitConversions = New ADODB.Recordset
    rsUnitConversions.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    ' Is it a new one?
    If rsUnitConversions.RecordCount = 0 Then
        'It's a new one
        rsUnitConversions.AddNew
    Else
        'It's an existing conversion factor
        'If running under MACRO_DM it will be automatically changed
        'If running under MACRO_EX we must ask the user if they want to change it
        If App.Title = "MACRO_SM" Then
            If dblConvFactor <> rsUnitConversions![ConversionFactor] Then
                sPrompt = "The import wants to change the conversion factor for " _
                    & vbNewLine & "'" & sFromUnit & "' to '" & sToUnit & "'" _
                    & vbNewLine & "from " & rsUnitConversions![ConversionFactor] _
                    & " to " & dblConvFactor & "." _
                    & vbNewLine & "Would you like to make this change?"
                If DialogQuestion(sPrompt) = vbYes Then
                    ' Accept the overwrite
                Else
                    'Don't accept the overwrite (set it back to itself)
                    dblConvFactor = rsUnitConversions![ConversionFactor]
                End If
            End If      ' If different conversion factor
        End If      ' If Exchange (manual import)
    End If
    rsUnitConversions.Fields(0) = sFromUnit
    rsUnitConversions.Fields(1) = sToUnit
    rsUnitConversions.Fields(2) = dblConvFactor
    rsUnitConversions.Fields(3) = sUnitClass
    'Mo Morris 30/8/01 Db Audit (ConversionExpression dropped from table UnitConversionFactors)
    'rsUnitConversions.Fields(4) = vConvExpr     ' NB Currently unused (MACRO 2.1, Oct 2000)
    rsUnitConversions.Update
    
    rsUnitConversions.Close
    Set rsUnitConversions = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportConversionFactors", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Function DoesUnitClassMatch(ByVal sFromUnit As String, _
                                    ByVal sToUnit As String, _
                                    ByVal sUnitClass As String) As Boolean
'---------------------------------------------------------------------
' NCJ 25/10/00
' Returns TRUE if sFromUnit and sToUnit are both defined in the Units table
' with class sUnitClass
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsUnits As ADODB.Recordset
Dim bClassMatch As Boolean

    On Error GoTo ErrHandler
    
    bClassMatch = True
    
    sSQL = "SELECT Count(*) FROM Units WHERE "
    ' NCJ 27/10/00 - Case-insensitive matching for Oracle
    ' (NB UnitClass is OK because the user can't change it)
    sSQL = sSQL & GetSQLStringEquals("Unit", sFromUnit)
    sSQL = sSQL & " OR " & GetSQLStringEquals("Unit", sToUnit)
    sSQL = sSQL & " AND UnitClass = '" & sUnitClass & "'"
    Set rsUnits = New ADODB.Recordset
    rsUnits.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    ' Do they both exist with this unit class?
    If rsUnits.Fields(0).Value <> 2 Then
        bClassMatch = False
    End If
    
    rsUnits.Close
    Set rsUnits = Nothing
    
    DoesUnitClassMatch = bClassMatch

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DoesUnitClassMatch", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Public Sub ClearCabExtractFolder()
'---------------------------------------------------------------------
'Mo Morris 28/2/00  Full path added to the kill statement
'                   Doevents wait added
'                   Ending statement of DIR changed to a full Dir with path command
' DPH 11/04/2002 - Made ClearCabExtractFolder public for use in CAB validation code
'---------------------------------------------------------------------
Dim sNextFile As String

    On Error GoTo ErrHandler

    ' DPH 17/10/2001 Make sure folder exists before opening
    If FolderExistence(gsCAB_EXTRACT_LOCATION & "dummy.txt") Then
        'SDM 26/01/00 SR2794
        ChDir gsCAB_EXTRACT_LOCATION
        'ChDir gsAppPath & "CabExtract"
        'loop through any files that might have been left behind in the
        'CabExtract folder and remove them
        'SDM 26/01/00 SR2794
        sNextFile = Dir(gsCAB_EXTRACT_LOCATION & "*.*")
        Do While sNextFile <> ""
            Kill gsCAB_EXTRACT_LOCATION & sNextFile
            Do Until Not FileExists(sNextFile)
                DoEvents
            Loop
            'get next file via the DIR command
            sNextFile = Dir(gsCAB_EXTRACT_LOCATION & "*.*")
        Loop
    Else
        gLog gsCLEAR_CABEXTR_FOLDER, "Clearing CAB files from " & gsCAB_EXTRACT_LOCATION & " failed as folder doesn't exist"
    End If
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ClearCabExtractFolder", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub

'---------------------------------------------------------------------
Public Sub ImportDocumentsAndGraphics()
'---------------------------------------------------------------------
'Mo Morris  28/2/00 File kill offs now happen after copying via a call to ClearCabExtractFolder
'---------------------------------------------------------------------
Dim sNextFile As String

    On Error GoTo ErrHandler

    ' DPH 17/10/2001 Make sure folder exists before opening
    If FolderExistence(gsCAB_EXTRACT_LOCATION & "dummy.txt") Then
        ChDir gsCAB_EXTRACT_LOCATION
    
        'loop through all the study documents and graphic files and move them into the documents folder
        sNextFile = Dir(gsCAB_EXTRACT_LOCATION & "*.*")
        Do While sNextFile <> ""
            'copy the file into the documents folder
            FileCopy gsCAB_EXTRACT_LOCATION & sNextFile, gsDOCUMENTS_PATH & sNextFile
            Do Until FileExists(gsDOCUMENTS_PATH & sNextFile)
                DoEvents
            Loop
            gLog gsIMPORT_DOC, sNextFile & " imported into the documents folder"
            'get next file via the DIR command
            sNextFile = Dir
        Loop
    Else
        gLog gsIMPORT_DOC_AND_GRAPHICS, "Importing into the documents folder from " & gsCAB_EXTRACT_LOCATION & " failed as folder doesn't exist"
    End If
    
    'clear the CabExtract folder after starting the documents and graphics
    ClearCabExtractFolder

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportDocumentsAndGraphics", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub

'---------------------------------------------------------------------
Private Function ImportLaboratory(nImportFileNumber As Integer) As Long
'---------------------------------------------------------------------
' Import/Update a lab and delete all related Normal Ranges
'---------------------------------------------------------------------
Dim lLaboratoryId As Long
Dim sLaboratoryCode As String
Dim sLaboratoryDescription As String
' TA 3/11/2000: changed site variable to variant
Dim vSite As Variant
Dim nChanged As Integer
Dim lDummyReadOfLineCounter As Long
Dim sSQL As String
Dim rsLaboratory As ADODB.Recordset
Dim sPrompt As String
Dim nResponse As Integer

    On Error GoTo ErrHandler

    Input #nImportFileNumber, sLaboratoryCode, sLaboratoryDescription, vSite, nChanged, lDummyReadOfLineCounter
    'MLM 06/01/03: Decode strings if required
    If mbUseHexDecoding Then
        sLaboratoryCode = HexDecodeChars(sLaboratoryCode)
        sLaboratoryDescription = HexDecodeChars(sLaboratoryDescription)
        If Not IsNull(vSite) Then
            vSite = HexDecodeChars(CStr(vSite))
        End If
    End If
 
    sSQL = " SELECT * FROM Laboratory WHERE "
    sSQL = sSQL & GetSQLStringEquals("LaboratoryCode", sLaboratoryCode)
    Set rsLaboratory = New ADODB.Recordset
    rsLaboratory.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    ' Is it a new code?
    If rsLaboratory.RecordCount = 0 Then
        'Its a new code
        rsLaboratory.AddNew
    Else
        'Its an existing code
        'If running under MACRO_DM the code will be automatically changed
        'If running under MACRO_EX the current code will be compared to the imported one
        'and if they differ the user will be prompted as to what they want to happen
        If App.Title = "MACRO_SM" Then
            If LCase(sLaboratoryDescription) <> LCase(rsLaboratory![LaboratoryDescription]) Then
                sPrompt = "The import wants to change the description of Laboratory code '" & sLaboratoryCode & "'" _
                    & vbNewLine & "from '" & rsLaboratory![LaboratoryDescription] & "'" _
                    & vbNewLine & "to '" & sLaboratoryDescription & "'." _
                    & vbNewLine & vbNewLine & "Do you wish to make this change?"
                If DialogQuestion(sPrompt) = vbNo Then
                    'Don't accept the overwrite (set it back to itself)
                    sLaboratoryDescription = rsLaboratory![LaboratoryDescription]
                End If
            End If
        End If
    End If
    rsLaboratory.Fields(0) = sLaboratoryCode
    rsLaboratory.Fields(1) = sLaboratoryDescription
    'Changed Mo Morris 9/11/00
    'For manual imports under Macro_EX set field Site = Null
    If App.Title = "MACRO_SM" Then
        rsLaboratory.Fields(2) = Null
    Else
        rsLaboratory.Fields(2) = vSite
    End If
    'always set the filed Changed = 0 (NoChange)
    rsLaboratory.Fields(3) = Changed.NoChange
    
    rsLaboratory.Update
    rsLaboratory.Close
    Set rsLaboratory = Nothing
    
    'clear out associated Normal Ranges
    MacroADODBConnection.Execute "DELETE FROM NormalRange WHERE LaboratoryCode = '" & sLaboratoryCode & "'"

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportLaboratory", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Function ImportSiteLaboratory(nImportFileNumber As Integer) As Long
'---------------------------------------------------------------------
'Mo Morris 9/1000   This sub is not currently being called but might be of
'use when the export Laboratory & NormalRange is written
'MLM 06/01/03: If it does ever get used, add appropriate string decoding.
'---------------------------------------------------------------------
Dim sSite As String
Dim sLaboratoryCode As String
Dim lDummyReadOfLineCounter As Long
Dim sSQL As String
Dim rsSiteLaboratory As ADODB.Recordset
Dim sPrompt As String
Dim nResponse As Integer

    On Error GoTo ErrHandler
    
    Input #nImportFileNumber, sSite, sLaboratoryCode, lDummyReadOfLineCounter
    sSQL = " SELECT * FROM SiteLaboratory WHERE "
    sSQL = sSQL & GetSQLStringEquals("Site", sSite) & " AND " & GetSQLStringEquals("LaboratoryId", sLaboratoryCode)
    Set rsSiteLaboratory = New ADODB.Recordset
    rsSiteLaboratory.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    'ascertain whether it's a new code
    If rsSiteLaboratory.RecordCount = 0 Then
        'Its a new code
        rsSiteLaboratory.AddNew
    Else
        'Its an existing Site/Laboratory combination, re-import it
    End If
    rsSiteLaboratory.Fields(0) = sSite
    rsSiteLaboratory.Fields(1) = sLaboratoryCode
    rsSiteLaboratory.Update
    rsSiteLaboratory.Close
    Set rsSiteLaboratory = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportSiteLaboratory", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Function ImportClinicalTestGroup(nImportFileNumber As Integer) As Long
'---------------------------------------------------------------------
Dim sClinicalTestGroupCode As String
Dim sClinicalTestGroupDescription As String
Dim lDummyReadOfLineCounter As Long
Dim sSQL As String
Dim rsClinicalTestGroup As ADODB.Recordset
Dim sPrompt As String

    On Error GoTo ErrHandler
 
    Input #nImportFileNumber, sClinicalTestGroupCode, sClinicalTestGroupDescription, lDummyReadOfLineCounter
    'MLM 06/01/03: Decode strings if required
    If mbUseHexDecoding Then
        sClinicalTestGroupCode = HexDecodeChars(sClinicalTestGroupCode)
        sClinicalTestGroupDescription = HexDecodeChars(sClinicalTestGroupDescription)
    End If
    
    sSQL = " SELECT * FROM ClinicalTestGroup WHERE "
    sSQL = sSQL & GetSQLStringEquals("ClinicalTestGroupCode", sClinicalTestGroupCode)
    Set rsClinicalTestGroup = New ADODB.Recordset
    rsClinicalTestGroup.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    ' Ascertain whether it's a new code
    If rsClinicalTestGroup.RecordCount = 0 Then
        'Its a new code
        rsClinicalTestGroup.AddNew
    Else
        'Its an existing code
        'If running under MACRO_DM the code will be automatically changed
        'If running under MACRO_EX the current code will be compared to the imported one
        'and if they differ the user will be prompted as to what they want to happen
        If App.Title = "MACRO_SM" Then
            If LCase(sClinicalTestGroupDescription) <> LCase(rsClinicalTestGroup![ClinicalTestGroupDescription]) Then
                sPrompt = "The import wants to change the description of Clinical Test Group code '" & sClinicalTestGroupCode & "'" _
                    & vbNewLine & "from '" & rsClinicalTestGroup![ClinicalTestGroupDescription] & "'" _
                    & vbNewLine & "to '" & sClinicalTestGroupDescription & "'." _
                    & vbNewLine & vbNewLine & "Do you wish to make this change?"
                If DialogQuestion(sPrompt) = vbNo Then
                    'Don't accept the overwrite (set it back to itself)
                    sClinicalTestGroupDescription = rsClinicalTestGroup![ClinicalTestGroupDescription]
                End If
            End If
        End If
    End If
    rsClinicalTestGroup.Fields(0) = sClinicalTestGroupCode
    rsClinicalTestGroup.Fields(1) = sClinicalTestGroupDescription
    rsClinicalTestGroup.Update
    rsClinicalTestGroup.Close
    Set rsClinicalTestGroup = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportClinicalTestGroup", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Function ImportClinicalTest(nImportFileNumber As Integer) As Long
'---------------------------------------------------------------------
'TA 24/10/2000: Unit variable now variant to ensure we don't get "#NULL#" as our unit
'---------------------------------------------------------------------
Dim sClinicalTestCode As String
Dim sClinicalTestDescription As String
Dim sClinicalTestGroupCode As String
Dim vUnit As Variant
Dim lDummyReadOfLineCounter As Long
Dim sSQL As String
Dim rsClinicalTest As ADODB.Recordset
Dim sPrompt As String

    On Error GoTo ErrHandler
    
    Input #nImportFileNumber, sClinicalTestCode, sClinicalTestDescription, _
        sClinicalTestGroupCode, vUnit, lDummyReadOfLineCounter
    'MLM 06/01/03: Decode strings if required
    If mbUseHexDecoding Then
        sClinicalTestCode = HexDecodeChars(sClinicalTestCode)
        sClinicalTestDescription = HexDecodeChars(sClinicalTestDescription)
        sClinicalTestGroupCode = HexDecodeChars(sClinicalTestGroupCode)
        'ASH 26/11/2002 Only decode if vUnit not null
        If Not IsNull(vUnit) Then
            vUnit = HexDecodeChars(CStr(vUnit))
        End If
    End If
        
    sSQL = " SELECT * FROM ClinicalTest WHERE " & GetSQLStringEquals("ClinicalTestCode", sClinicalTestCode)
    Set rsClinicalTest = New ADODB.Recordset
    rsClinicalTest.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    'See if it's a new code
    If rsClinicalTest.RecordCount = 0 Then
        'Its a new code
        rsClinicalTest.AddNew
    Else
        'Its an existing code
        'If running under MACRO_DM the code will be automatically changed
        'If running under MACRO_EX the current code will be compared to the imported one
        'and if they differ the user will be prompted as to what they want to happen
        'removenull used for units so that user isn't prompted to replace empty sting with null
        If App.Title = "MACRO_SM" Then
            If LCase(sClinicalTestDescription) <> LCase(rsClinicalTest![ClinicalTestDescription]) _
                Or LCase(sClinicalTestGroupCode) <> LCase(rsClinicalTest![ClinicalTestGroupCode]) _
                Or LCase(RemoveNull(vUnit)) <> LCase(RemoveNull(rsClinicalTest![Unit])) Then
                sPrompt = "The import wants to change the meaning of Clinical Test code '" & sClinicalTestCode & "'" _
                    & vbNewLine & "from [Desc: '" & rsClinicalTest![ClinicalTestDescription] _
                    & "', Group: '" & rsClinicalTest![ClinicalTestGroupCode] & "', Unit: '" & RemoveNull(rsClinicalTest![Unit]) & "']" _
                    & vbNewLine & "to [Desc: '" & sClinicalTestDescription _
                    & "', Group: '" & sClinicalTestGroupCode & "', Unit: '" & RemoveNull(vUnit) & "']." _
                    & vbNewLine & vbNewLine & "Do you wish to make this change?"
                If DialogQuestion(sPrompt) = vbNo Then
                    'Don't accept the overwrite (set it back to itself)
                    sClinicalTestDescription = rsClinicalTest![ClinicalTestDescription]
                    sClinicalTestGroupCode = rsClinicalTest![ClinicalTestGroupCode]
                    vUnit = rsClinicalTest![Unit]
                End If
            End If
        End If
    End If
    rsClinicalTest.Fields(0) = sClinicalTestCode
    rsClinicalTest.Fields(1) = sClinicalTestDescription
    rsClinicalTest.Fields(2) = sClinicalTestGroupCode
    rsClinicalTest.Fields(3) = vUnit
    rsClinicalTest.Update
    rsClinicalTest.Close
    Set rsClinicalTest = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportNormalRange", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Sub ImportNormalRange(nImportFileNumber As Integer)
'---------------------------------------------------------------------
' Insert NormalRanges into NormalRange table
'  assumes all ranges for this lab/test combination have previously been deleted
'---------------------------------------------------------------------
Dim lNormalRangeId As Long
Dim sLaboratoryCode As String
Dim sClinicalTestCode As String
Dim nGender As Integer
Dim vAgeMin As Variant
Dim vAgeMax As Variant
Dim vEffectiveStart As Variant
Dim vEffectiveEnd As Variant
Dim vNormalMin As Variant
Dim vNormalMax As Variant
Dim vFeasibleMin As Variant
Dim vFeasibleMax As Variant
Dim vAbsoluteMin As Variant
Dim vAbsoluteMax As Variant
Dim nPercent As Integer
Dim lDummyReadOfLineCounter As Long
Dim sSQL As String

    On Error GoTo ErrHandler
    
    Input #nImportFileNumber, lNormalRangeId, sLaboratoryCode, sClinicalTestCode, _
        nGender, vAgeMin, vAgeMax, vEffectiveStart, vEffectiveEnd, vNormalMin, vNormalMax, _
        vFeasibleMin, vFeasibleMax, vAbsoluteMin, vAbsoluteMax, nPercent, lDummyReadOfLineCounter
    'MLM 06/01/03: Decode strings if required
    If mbUseHexDecoding Then
        sLaboratoryCode = HexDecodeChars(sLaboratoryCode)
        sClinicalTestCode = HexDecodeChars(sClinicalTestCode)
    End If
 
    sSQL = "INSERT INTO NormalRange VALUES("
    'TA 01/03/2001: convert local number format to standard
    sSQL = sSQL & lNormalRangeId & ", '" & sLaboratoryCode & "', '" & sClinicalTestCode & "'"
    sSQL = sSQL & ", " & nGender
    sSQL = sSQL & ", " & VarianttoString(vAgeMin, True)
    sSQL = sSQL & ", " & VarianttoString(vAgeMax, True)
    sSQL = sSQL & ", " & ConvertLocalNumToStandard(VarianttoString(vEffectiveStart, True))
    sSQL = sSQL & ", " & ConvertLocalNumToStandard(VarianttoString(vEffectiveEnd, True))
    sSQL = sSQL & ", " & ConvertLocalNumToStandard(VarianttoString(vNormalMin, True))
    sSQL = sSQL & ", " & ConvertLocalNumToStandard(VarianttoString(vNormalMax, True))
    sSQL = sSQL & ", " & ConvertLocalNumToStandard(VarianttoString(vFeasibleMin, True))
    sSQL = sSQL & ", " & ConvertLocalNumToStandard(VarianttoString(vFeasibleMax, True))
    sSQL = sSQL & ", " & ConvertLocalNumToStandard(VarianttoString(vAbsoluteMin, True))
    sSQL = sSQL & ", " & ConvertLocalNumToStandard(VarianttoString(vAbsoluteMax, True))
    sSQL = sSQL & ", " & nPercent
    sSQL = sSQL & ")"

    MacroADODBConnection.Execute sSQL
    
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportNormalRange", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub

'---------------------------------------------------------------------
Private Function ImportCTCScheme(nImportFileNumber As Integer) As Long
'---------------------------------------------------------------------
' Import/Update a CTC Schemelab and delete all related CTC
'---------------------------------------------------------------------
Dim sCTCSchemeCode As String
Dim sCTCSchemeDescription As String
Dim lDummyReadOfLineCounter As Long
Dim sSQL As String
Dim rsCTCScheme As ADODB.Recordset
Dim sPrompt As String
Dim nResponse As Integer
    
    On Error GoTo ErrHandler
    
    Input #nImportFileNumber, sCTCSchemeCode, sCTCSchemeDescription, lDummyReadOfLineCounter
    'MLM 06/01/03: Decode strings if required
    If mbUseHexDecoding Then
        sCTCSchemeCode = HexDecodeChars(sCTCSchemeCode)
        sCTCSchemeDescription = HexDecodeChars(sCTCSchemeDescription)
    End If
 
    sSQL = " SELECT * FROM CTCScheme WHERE " & GetSQLStringEquals("CTCSchemeCode", sCTCSchemeCode)
    Set rsCTCScheme = New ADODB.Recordset
    rsCTCScheme.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    'ascertain whether its a new code
    If rsCTCScheme.RecordCount = 0 Then
        'Its a new code
        rsCTCScheme.AddNew
    Else
        'Its an existing code
        'If running under MACRO_DM the code will be automatically changed
        'If running under MACRO_EX the current code will be compared to the imported one
        'and if they differ the user will be prompted as to what they want to happen
        If App.Title = "MACRO_SM" Then
            If LCase(sCTCSchemeDescription) <> LCase(rsCTCScheme![CTCSchemeDescription]) Then
                sPrompt = "The import wants to change the description of CTC Scheme code '" & sCTCSchemeCode & "'" _
                    & vbNewLine & "from '" & rsCTCScheme![CTCSchemeDescription] & "'" _
                    & vbNewLine & "to '" & sCTCSchemeDescription & "'." _
                    & vbNewLine & vbNewLine & "Do you wish to make this change?"
                If DialogQuestion(sPrompt) = vbNo Then
                    'Don't accept the overwrite (set it back to itself)
                    sCTCSchemeDescription = rsCTCScheme![CTCSchemeDescription]
                End If
            End If
            ' For an existing CTCSchemeCode, clear out associated CTC
            ' but ask user first if in Exchange
            sPrompt = "Do you wish to overwrite all toxicity criteria for the existing CTC Scheme '" & sCTCSchemeCode & "'?"
            If DialogQuestion(sPrompt) = vbYes Then
                MacroADODBConnection.Execute "DELETE FROM CTC WHERE CTCSchemeCode = '" & sCTCSchemeCode & "'"
            Else
                ' Store it in collection of CTC Schemes NOT to do
                mcolCTCSchemesNotToDo.Add LCase(sCTCSchemeCode)
            End If
        Else
            ' Automatic import - delete CTC anyway
            MacroADODBConnection.Execute "DELETE FROM CTC WHERE CTCSchemeCode = '" & sCTCSchemeCode & "'"
        End If
    End If
    
    rsCTCScheme.Fields(0) = sCTCSchemeCode
    rsCTCScheme.Fields(1) = sCTCSchemeDescription
    rsCTCScheme.Update
    rsCTCScheme.Close
    Set rsCTCScheme = Nothing
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportCTCScheme", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Sub ImportCTC(nImportFileNumber As Integer)
'---------------------------------------------------------------------
' Insert CTC into CTC table
'  assumes all CTC for this lab/test combination have previously been deleted
' NCJ 27/10/00 - Check if we want to overwrite criteria for this CTCSchemeCode
'---------------------------------------------------------------------
Dim lCTCId As Long
Dim sCTCSchemeCode As String
Dim sClinicalTestCode As String
Dim nCTCGrade As Integer
Dim vCTCMin As Variant
Dim vCTCMax As Variant
Dim nCTCMinType As Integer
Dim nCTCMaxType As Integer
Dim lDummyReadOfLineCounter As Long
Dim sSQL As String
Dim vCTCScheme As Variant

    On Error GoTo ErrHandler
    
    Input #nImportFileNumber, lCTCId, sCTCSchemeCode, sClinicalTestCode, _
        nCTCGrade, vCTCMin, vCTCMax, nCTCMinType, nCTCMaxType, lDummyReadOfLineCounter
    'MLM 06/01/03: Decode strings if required
    If mbUseHexDecoding Then
        sCTCSchemeCode = HexDecodeChars(sCTCSchemeCode)
        sClinicalTestCode = HexDecodeChars(sClinicalTestCode)
    End If
    
    ' NCJ 27/10/00 See if this scheme is in the collection of schemes NOT to do
    For Each vCTCScheme In mcolCTCSchemesNotToDo
        If CStr(vCTCScheme) = LCase(sCTCSchemeCode) Then
            ' It is in the collection so don't do it
            Exit Sub
        End If
    Next
    
    sSQL = "INSERT INTO CTC VALUES("
    'TA 01/03/2001: convert local number format to standard
    sSQL = sSQL & lCTCId & ", '" & sCTCSchemeCode & "', '" & sClinicalTestCode & "'"
    sSQL = sSQL & ", " & nCTCGrade
    sSQL = sSQL & ", " & ConvertLocalNumToStandard(VarianttoString(vCTCMin, True))
    sSQL = sSQL & ", " & ConvertLocalNumToStandard(VarianttoString(vCTCMax, True))
    sSQL = sSQL & ", " & nCTCMinType
    sSQL = sSQL & ", " & nCTCMaxType
    sSQL = sSQL & ")"

    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportCTC", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub

'---------------------------------------------------------------------
Public Function ExportLDD(sLaboratory As String, _
                            Optional colFilesInCAB As Collection) As String
'---------------------------------------------------------------------
' Create a Laboratory Definition export file
' NCJ 8 Nov 2000 - Include other tables (SegmentId > 300 ) in export
' DPH 11/04/2002 - Added CAB Validation info collection
' DPH 11/04/2002 - Using ExportLDDZIP instead of ExportLDDCAB
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsExportImport As ADODB.Recordset
Dim rsCurrentTable As ADODB.Recordset
Dim dtNow As Date
Dim sSessionName As String
Dim sTableName As String
Dim sSegmentId As String
Dim nCount As Long
Dim i As Integer
Dim nExportFileNumber As Integer
Dim nNumFields As Integer
Dim vFieldData As Variant
Dim sVersion As String
Dim afilelist() As String
Dim nFileNumber As Integer
' DPH 11/04/2002 - Require variables for file validation
Dim lFileSize As Long

    On Error GoTo ErrHandler
    
    'dtNow is a date/time variable that is used as a time stamp that will be
    'placed in both the HEADER and FOOTER records of the export files
    'in addition the yyyymmddhhmmss parts of the date stamp are appended to the
    'export file names
    dtNow = Now
    'create a session name, which will become the cab name e.g. Lab_Labcode.CAB
    sSessionName = sLaboratory & "_LAB"
    gLog gsEXPORT_LDD, "Session " & sSessionName & " starting"
    
    'initialise FileNumber which will be used as an index in afilelist
    nFileNumber = 0
    
    nCount = 0
    nExportFileNumber = FreeFile
    
    'create file name for the export's LDD file
    msFileName = gsOUT_FOLDER_LOCATION & sLaboratory & "_" & Format(dtNow, "yyyymmddhhmmss") & ".ldd"
    
    ' DPH 17/10/2001 Make sure folder exists before opening
    If Not FolderExistence(msFileName) Then
        ExportLDD = ""
        gLog gsEXPORT_LDD, "Export folder " & gsOUT_FOLDER_LOCATION & " could not be created"
        Exit Function
    End If
    
    'redim the string array afilelist, into which the current file name will be added
    ReDim Preserve afilelist(nFileNumber)
    afilelist(nFileNumber) = msFileName
    nFileNumber = nFileNumber + 1
    
    Open msFileName For Output As #nExportFileNumber
    nCount = nCount + 1
    'Mo 20/02/2006 COD0210/COD0220, Version taken from Database instead of Software
    sVersion = GetVersionFromDB
    Write #nExportFileNumber, "UNH", "LABDEF", dtNow, "UID SENDER", "UID DESTINATION", sVersion, nCount
    
    sSQL = "Select TableName, SegmentId From MACROTable" _
        & " WHERE LABDEF = 1" _
        & " ORDER BY SegmentId"
    Set rsExportImport = New ADODB.Recordset
    rsExportImport.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Do While Not rsExportImport.EOF
        'Debug.Print rsExportImport![segmentid] & " " & rsExportImport![TableName]
        sTableName = rsExportImport![TableName]
        sSegmentId = rsExportImport![SegmentId]
        ' NCJ 8/11/00 For Segment Ids less than 300, filter on LaboratoryCode
        If sSegmentId < 300 Then
            sSQL = " SELECT * FROM " & sTableName _
                & " WHERE LaboratoryCode = '" & sLaboratory & "'"
        Else
            'Changed ATO 22/08/2001
            sSQL = EvaluateExceptionalLDDTables(sLaboratory, sSegmentId, sTableName)
           
        End If
        Set rsCurrentTable = New ADODB.Recordset
        rsCurrentTable.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not rsCurrentTable.EOF
            nNumFields = rsCurrentTable.Fields.Count
            nCount = nCount + 1
            Write #nExportFileNumber, sSegmentId;
            For i = 0 To (nNumFields - 1)
                vFieldData = rsCurrentTable.Fields(i).Value
                'MLM 06/01/03: Use new string encoding
                If VarType(vFieldData) = vbString Then
                    vFieldData = HexEncodeChars(vFieldData, msCHARS_TO_ENCODE)
                End If
                Write #nExportFileNumber, vFieldData;
            Next i
            Write #nExportFileNumber, nCount
            rsCurrentTable.MoveNext
        Loop
        rsCurrentTable.Close
        Set rsCurrentTable = Nothing
        rsExportImport.MoveNext
    Loop
    rsExportImport.Close
    Set rsExportImport = Nothing
    
    nCount = nCount + 1
    Write #nExportFileNumber, "UNT", dtNow, nCount
    
    Close #nExportFileNumber
    gLog gsEXPORT_LDD, msFileName & " Export file created"
    
    ' DPH 11/04/2002 - Create File list + file sizes
    ' Initialise Collection if required
    Set colFilesInCAB = New Collection
    ' Loop through File List & collect file sizes
    For i = 0 To UBound(afilelist)
        lFileSize = GetFileLength(afilelist(i))
        ' Store in collection
        colFilesInCAB.Add lFileSize, StripFileNameFromPath(afilelist(i))
    Next
    
    ' DPH 11/04/2002 - Using ExportLDDZIP instead of ExportLDDCAB
    'ExportLDDCAB sSessionName, afilelist
    ExportLDDZIP sSessionName, afilelist
    
    'Mo Morris 24/2/00
    ' DPH 15/04/2002 - Removed CAB reference
    'ExportLDD = sSessionName & ".cab"
    ExportLDD = sSessionName & ".zip"

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ExportLDD", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Public Sub ExportLDDCAB(CABFileName As String, CABFileList() As String)
'---------------------------------------------------------------------
Dim nFileNumber As Integer
Dim sCABfile As String
Dim i As Long

    On Error GoTo ErrHandler
    
    gLog gsEXPORT_LDD_CAB, "Session " & CABFileName & " starting"
        
    'Create a ddf filename file containing all the items (.ldd etc) to be compressed
    nFileNumber = FreeFile
    sCABfile = gsOUT_FOLDER_LOCATION & CABFileName & ".ddf"
    
    ' DPH 17/10/2001 Make sure folder exists before opening
    If FolderExistence(sCABfile) Then
        Open sCABfile For Output As nFileNumber
        
        'Place initial settings into the .ddf file
        Print #nFileNumber, ".OPTION EXPLICIT"
        Print #nFileNumber, ".Set CabinetNameTemplate=" & CABFileName & ".cab"
        Print #nFileNumber, ".Set DiskDirectoryTemplate=" & gsOUT_FOLDER_LOCATION
        Print #nFileNumber, ".Set DiskLabel1=" & CABFileName
        'Changed Mo Morris 3/7/2000 SR 3659
        Print #nFileNumber, ".Set MaxDiskSize=0"
        
        'Add the list of files to be compressed from CABFileList to the .ddf file
        For i = 0 To UBound(CABFileList)
            Print #nFileNumber, """" & CABFileList(i) & """"
            'Debug.Print CABFileList(i)
        Next i
        
        Close #nFileNumber
        
        'Change directory to where MakeCab.exe exists
        'Note that the DiskDirectoryTemplate location can either be a folder below gsAppPath
        'or as in the case of Roche a whole path location that could be anywhere
            
        ' DPH 17/10/2001 Make sure folder exists before opening
        Call FolderExistence(gsAppPath & "dummy.txt")
    
        ChDir gsAppPath
        
        'MakeCab runs under DOS and as such would normally be called via the Shell command.
        'Unfortunately the Shell command simply starts a process and then moves on to the next
        'line of code, which in this case would start the killing of the files to be CABed.
        'Instead the API utility ExecCmd, which completes the called process before moving onto
        'the next line,is used.
        ExecCmd "Makecab /f " & """" & sCABfile & """"
        
        'Kill off the first file in CABFileList which will be an SDD file
        'the rest of CABFileList refers to Study Documents and grahics that should not be deleted
        Kill CABFileList(0)
        
        'Kill the .ddf file
        Kill sCABfile
        
        gLog gsEXPORT_LDD_CAB, "Session " & CABFileName & " completed"
    Else
        gLog gsEXPORT_LDD_CAB, "Session failed as could not create folder " & gsOUT_FOLDER_LOCATION
    End If
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ExportLDDCAB", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub

'---------------------------------------------------------------------
Public Function ImportLDD(ByVal vImportFile As String) As Long
'---------------------------------------------------------------------
' Import Laboratory Definition file
' NCJ 8 Nov 00 - Deal with ClinicalTest, ClinicalTestGroup and Units
'---------------------------------------------------------------------
Dim nImportFileNumber As Integer
Dim rsExportImport As ADODB.Recordset
Dim rsCurrentTable As ADODB.Recordset
Dim sSegmentId As String
Dim sImportRec As String
Dim sSQL As String
Dim bParamFieldsNotAssessed As Boolean
Dim nNumFields As Integer
Dim i As Integer
Dim vntReadField As Variant
Dim vntRecordNumber As Variant

    On Error GoTo ErrHandler

    ImportLDD = CheckImportFile(vImportFile, "LABDEF")
    If ImportLDD <> ExchangeError.Success Then
        Exit Function
    End If
    
    'MLM 06/01/03:
    mbUseHexDecoding = (mlImportVersion >= 3000029) '3.0.29
    
    On Error GoTo ErrTransactionRollBack
    'Begin transaction
    TransBegin
       
    nImportFileNumber = FreeFile + 1
    Open vImportFile For Input As #nImportFileNumber
    'Read the first line which has already been assessed
    Line Input #nImportFileNumber, sImportRec
    
    'prepare a recordset of the ExportImport parameters file
    sSQL = "Select * From MACROTable" _
        & " WHERE LABDEF = 1" _
        & " ORDER BY SegmentId"
    Set rsExportImport = New ADODB.Recordset
    rsExportImport.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
    bParamFieldsNotAssessed = True
    
    'The Import file is read data item at a time by 2 Input statements.
    'The first reads only the Segment Id.
    'The second is controlled by a loop that is driven by the number of
    'fields within the current record type (nNumFields).
    'Between the 2 Input statements there is code to check if the record
    'type has changed (i.e. the Segment Id has changed) and then via the
    'Export Parameters file accesses the next table name from which a new
    'value for nNumFields can be created.
    Do While Not EOF(nImportFileNumber)
        Input #nImportFileNumber, sSegmentId
        If sSegmentId = "UNT" Then
            Exit Do
        End If
        Do Until sSegmentId = rsExportImport![SegmentId]
            rsExportImport.MoveNext
            bParamFieldsNotAssessed = True
        Loop
        If bParamFieldsNotAssessed Then
            sSQL = " SELECT * FROM " & rsExportImport![TableName]
            Set rsCurrentTable = New ADODB.Recordset
            rsCurrentTable.Open sSQL, MacroADODBConnection, adOpenDynamic, adLockPessimistic, adCmdText
            nNumFields = rsCurrentTable.Fields.Count
            bParamFieldsNotAssessed = False
        End If
        ' Try out changes
        ' NCJ 8/11/00 - Changed from If...Then to Select Case
        Select Case sSegmentId
        Case "001"  ' Laboratory
            ImportLDD = ImportLaboratory(nImportFileNumber)
            If ImportLDD <> ExchangeError.Success Then
                Err.Raise _
                    Number:=vbObjectError + 1101, _
                    Description:="Import Aborted by user over Laboratory discrepancy."
            End If
        Case "030"  ' NormalRange
            Call ImportNormalRange(nImportFileNumber)
        ' NCJ 8/11/00 - Deal with ClinicalTest, ClinicalTestGroup and Units
        Case "550"  ' Units
            ImportLDD = ImportUnits(nImportFileNumber)
            If ImportLDD <> ExchangeError.Success Then
                Err.Raise _
                    Number:=vbObjectError + 1004, _
                    Description:="Import Aborted by user over Units discrepancy."
            End If
        Case "610"  ' ClinicalTestGroup
            ImportLDD = ImportClinicalTestGroup(nImportFileNumber)
            If ImportLDD <> ExchangeError.Success Then
                Err.Raise _
                    Number:=vbObjectError + 1005, _
                    Description:="Import Aborted by user over Clinical Test Group discrepancy."
            End If
        Case "620"  ' ClinicalTest
            ImportLDD = ImportClinicalTest(nImportFileNumber)
            If ImportLDD <> ExchangeError.Success Then
                Err.Raise _
                    Number:=vbObjectError + 1006, _
                    Description:="Import Aborted by user over Clinical Test discrepancy."
            End If
        End Select
    Loop
    
    rsCurrentTable.Close
    Set rsCurrentTable = Nothing
    rsExportImport.Close
    Set rsExportImport = Nothing
    
    Close #nImportFileNumber
    
    'End transaction
    TransCommit
    
    On Error GoTo ErrHandler
    
    gLog gsIMPORT_LDD, StripFileNameFromPath(vImportFile) & " imported successfully"
    
    
    Kill vImportFile
    
    Do Until Not FileExists(vImportFile)
        DoEvents
    Loop
    gLog gsIMPORT_LDD, StripFileNameFromPath(vImportFile) & " deleted after import"
    
    ImportLDD = ExchangeError.Success
    
    Exit Function
    
ErrTransactionRollBack:
    
        'RollBack transaction
        TransRollBack
        
        If (Err.Number - vbObjectError) = 1101 Or (Err.Number - vbObjectError) = 1102 Then
            gLog gsIMPORT_LDD, "Error code " & (Err.Number - vbObjectError) & " - " & Err.Description
            ImportLDD = ExchangeError.UserAborted
        Else
            gLog gsIMPORT_LDD, "Error code " & Err.Number - vbObjectError & " - " & Err.Description
            gLog gsIMPORT_LDD, "Error occurred on record number= " & vntRecordNumber + 1 & " segment= " & sSegmentId
            ImportLDD = ExchangeError.Unknown
        End If
        
        gLog gsIMPORT_LDD, "Import aborted. All updates will be rolled back."
        Close
    
    Exit Function
ErrHandler:
      Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportLDD", "clsExchange")
            Case OnErrorAction.Ignore
                Resume Next
            Case OnErrorAction.Retry
                Resume
            Case OnErrorAction.QuitMACRO
                Call ExitMACRO
                End
       End Select
End Function

'---------------------------------------------------------------------
Public Function ImportLDDCAB(ByVal gsImportFile As String)
'---------------------------------------------------------------------
' REVISIONS
' DPH 11/04/2002 - Check process launched properly & write appropriate log entry
'                   Branch if ZIP file from previous MACRO version
'---------------------------------------------------------------------
Dim sCommand As String
Dim sShortFormImportFile As String
Dim sShortFormDirectory As String

' DPH 11/04/2002 - Check if CAB or ZIP dealing with & act appropriately
If UCase(Right(gsImportFile, 3)) = "ZIP" Then
    Call ImportLDDZIP(gsImportFile)
    Exit Function
End If

gLog gsEXPORT_LDD_CAB, "Extracting " & StripFileNameFromPath(gsImportFile) & " starting"

On Error GoTo ErrHandler

'clear the CabExtract folder before starting the Import
ClearCabExtractFolder

'Prepare a short form of the Extraction directory
sShortFormDirectory = gsCAB_EXTRACT_LOCATION
sShortFormDirectory = GetShortPathName(sShortFormDirectory)

' DPH 17/10/2001 - Make sure file exists before opening
If FolderExistence(gsImportFile, True) Then
    'Prepare a short form of the ImportFile
    sShortFormImportFile = GetShortPathName(gsImportFile)
    
    'sCommand = gsAppPath & "Extract /A " & """" & gsImportFile & """" & " /E" & " /Y" '*1
    'Call the CAB extract program which is in Macro's application path
    'The /Y paramater means that the user will not be prompted about potential overwrites
    'The /A paramater means process All files in the CAB file
    'The /E parameter means force the extraction to happen
    'The /L parameter together with sShortFormDirectory is where the files are extracted to
    sCommand = gsAppPath & "Extrac32 /Y /A /E /L " & """" & sShortFormDirectory & """ """ & sShortFormImportFile & """"
    
    ' DPH 25/03/2002 - Check process launched properly & write appropriate log entry
    If ExecCmd(sCommand) > 0 Then
        gLog gsEXPORT_LDD_CAB, "Extracting " & StripFileNameFromPath(gsImportFile) & " completed"
    Else
        gLog gsEXPORT_LDD_CAB, "Extracting " & StripFileNameFromPath(gsImportFile) & " failed. Process launch failure. AutoImport Error (Lab)."
    End If
Else
    gLog gsEXPORT_LDD_CAB, "Extracting from " & gsCAB_EXTRACT_LOCATION & " failed as file doesn't exist"
End If
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportLDDCAB", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Sub CleanUpPatientData(ByVal sImportFile As String)
'---------------------------------------------------------------------
'Added by Mo Morris 23/1/2001
'This procedure reads a patient data import file and corrects:-
'   DataItemResponse and DataitemResponseHistory entries that might have
'   an OverruleReason field containing invalid double quotes and replaces
'   pairs of double quotes with a single quote.
'   DataItemResponse entries that might have a ResponseTimeStamp field that
'   contains a 14 charter digit instead of a double and inserts a decimal
'   point after the fith digit.
'---------------------------------------------------------------------
Dim nImportFileNumber As Integer
Dim nCleanUpFileNumber As Integer
Dim sCleanUpFile As String
Dim sImportRec As String
Dim sSegmentId As String
Dim i As Integer
Dim sChar As String
Dim nCommaCounter As Integer
Dim bInQuote As Boolean
Dim nComma13 As Integer
Dim nComma14 As Integer
Dim nComma29 As Integer
Dim nComma30 As Integer
Dim sField14 As String
Dim sField30 As String
Dim bField14Changed As Boolean
Dim bField30Changed As Boolean
Dim bHasFileChanged As Boolean
    
    On Error GoTo ErrHandler
    
    bHasFileChanged = False
    
    nImportFileNumber = FreeFile
    Open sImportFile For Input As #nImportFileNumber
    
    nCleanUpFileNumber = FreeFile
    sCleanUpFile = "CleanUp.dat"
    Open sCleanUpFile For Output As #nCleanUpFileNumber
    
    Do While Not EOF(nImportFileNumber)
        Line Input #nImportFileNumber, sImportRec
        'extract segmentId
        sSegmentId = Mid(sImportRec, 1, 5)
        
        'check for segmentId "080" DataItemRespone
        If sSegmentId = """080""" Then
            'read through comma separated sImportRec, marking the comma position before and
            'after field 14 (ResponseTimeStamp) and field 30 (OverruleReason)
            nCommaCounter = 0
            bInQuote = False
            For i = 1 To Len(sImportRec)
                If sChar = """" Then
                    bInQuote = Not bInQuote
                End If
                sChar = Mid(sImportRec, i, 1)
                If sChar = "," And Not bInQuote Then
                    nCommaCounter = nCommaCounter + 1
                    If nCommaCounter = 13 Then nComma13 = i
                    If nCommaCounter = 14 Then nComma14 = i
                    If nCommaCounter = 29 Then nComma29 = i
                    If nCommaCounter = 30 Then nComma30 = i
                End If
            Next
            sField14 = Mid(sImportRec, (nComma13 + 1), (nComma14 - (nComma13 + 1)))
            sField30 = Mid(sImportRec, (nComma29 + 1), (nComma30 - (nComma29 + 1)))
            bField14Changed = CheckResponseTimeStamp(sField14)
            bField30Changed = CheckOverruleReason(sField30)
            'check for any changes having been made
            If bField14Changed Or bField30Changed Then
                bHasFileChanged = True
                sImportRec = Mid(sImportRec, 1, nComma13) & sField14 _
                    & Mid(sImportRec, nComma14, (nComma29 - nComma14) + 1) & sField30 _
                    & Mid(sImportRec, nComma30)
            End If
        End If
        
        'check for segmentId "090" DataItemResponseHistory
        If sSegmentId = """090""" Then
            'read through comma separated sImportRec, marking the comma position before and
            'after field 30 (OverruleReason)
            nCommaCounter = 0
            bInQuote = False
            For i = 1 To Len(sImportRec)
                If sChar = """" Then
                    bInQuote = Not bInQuote
                End If
                sChar = Mid(sImportRec, i, 1)
                If sChar = "," And Not bInQuote Then
                    nCommaCounter = nCommaCounter + 1
                    If nCommaCounter = 29 Then nComma29 = i
                    If nCommaCounter = 30 Then nComma30 = i
                End If
            Next
            sField30 = Mid(sImportRec, (nComma29 + 1), (nComma30 - (nComma29 + 1)))
            bField30Changed = CheckOverruleReason(sField30)
            'check for any changes having been made
            If bField30Changed Then
                bHasFileChanged = True
                sImportRec = Mid(sImportRec, 1, nComma29) & sField30 _
                    & Mid(sImportRec, nComma30)
            End If
        End If
        
        
        'write the current record to sCleanUpFile
        Print #nCleanUpFileNumber, sImportRec
    Loop
    
    Close #nImportFileNumber
    Close #nCleanUpFileNumber
    
    'If any changes have been made then overwrite sImportFile with the contents of sCleanUpFile
    If bHasFileChanged Then
        Kill sImportFile
        Do Until Not FileExists(sImportFile)
            DoEvents
        Loop
        FileCopy sCleanUpFile, sImportFile
        Do Until FileExists(sImportFile)
            DoEvents
        Loop
        gLog gsCLEANUP_PRD, sImportFile & " has been cleaned up"
    End If
    Kill sCleanUpFile
    Do Until Not FileExists(sCleanUpFile)
        DoEvents
    Loop
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CleanUpPatientData", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub

'---------------------------------------------------------------------
Private Function CheckResponseTimeStamp(ByRef sResponseTimeStamp As String) As Boolean
'---------------------------------------------------------------------
'This Function checks for a double date with its decimal point missing
'   i.e.    36873.4642013889 being stored as 368734642013889
'and corrects it by putting back the decimal point.
'
'This function returns False if the ResponseTimeStamp was Ok
'and True if it needed changing
'---------------------------------------------------------------------

    'Changed Mo Morris 29/1/01, check for yet another ResponseTimeStamp error
    'in the DataItemResponse table a timestamp of NULL
    On Error GoTo ErrHandler
    
    If sResponseTimeStamp = "#NULL#" Then
        CheckResponseTimeStamp = False
        Exit Function
    End If
    
    If InStr(sResponseTimeStamp, ".") <> 0 Then
        CheckResponseTimeStamp = False
    Else
        sResponseTimeStamp = Mid(sResponseTimeStamp, 1, 5) & "." & Mid(sResponseTimeStamp, 6)
        CheckResponseTimeStamp = True
    End If
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CheckResponseTimeStamp", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Private Function CheckOverruleReason(ByRef sOverruleReason As String) As Boolean
'---------------------------------------------------------------------
'This function returns False if the OverruleReason was Ok
'and True if it needed changing
'---------------------------------------------------------------------
Dim sWithoutQuotes As String
Const sDOUBLEDOUBLEQUOTES As String = """"""
Const sDOUBLEQUOTE As String = """"
Const sSINGLEQUOTE As String = "'"

    On Error GoTo ErrHandler

    sWithoutQuotes = Mid(sOverruleReason, 2, Len(sOverruleReason) - 2)
    If InStr(sWithoutQuotes, sDOUBLEDOUBLEQUOTES) = 0 Then
        CheckOverruleReason = False
    Else
        sOverruleReason = sDOUBLEQUOTE & Replace(sWithoutQuotes, sDOUBLEDOUBLEQUOTES, sSINGLEQUOTE) & sDOUBLEQUOTE
        CheckOverruleReason = True
    End If

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CheckOverruleReason", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function
'--------------------------------------------------------------------------------------
Public Function EvaluateExceptionalSDDTables(ByVal msSegmentId As String, _
                                            ByVal msTablename As String, _
                                            ByVal vClinicalTrialId As Long) As String
'--------------------------------------------------------------------------------------
'Ash 22/08/2001 Routine to handle new exportSDD due to introduction of
'Normal ranges,CTC schemes and Laboratories.
'This routine reduces the number of records in an export file by only
'sending entries that are used by the trial being exported.
'--------------------------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    sSQL = ""
    
    'TrialPhase
    If msSegmentId = 510 Then
        sSQL = "Select TrialPhase.* From TrialPhase,ClinicalTrial" & _
                " Where Clinicaltrial.PhaseID = TrialPhase.PhaseID" & _
                " AND ClinicalTrial.ClinicalTrialID =" & vClinicalTrialId
    
    'TrialType
    ElseIf msSegmentId = 520 Then
       sSQL = "Select TrialType.* From TrialType,ClinicalTrial" & _
                " Where Clinicaltrial.TrialTypeID = TrialType.TrialTypeID" & _
                " AND ClinicalTrial.ClinicalTrialID =" & vClinicalTrialId
    

    'ValidationType
    ElseIf msSegmentId = 540 Then
        sSQL = "Select DISTINCT ValidationType.* From ValidationType,DataItemValidation" & _
                " Where DataItemValidation.ValidationTypeID = ValidationType.ValidationTypeID" & _
                " AND DataItemValidation.ClinicalTrialID =" & vClinicalTrialId
    
    'Units
    'RJCW 11/10/2001 BUGFIX 031
    'Changed SQL to pick up Conversion factor units
    ElseIf msSegmentId = 550 Then
        sSQL = "Select DISTINCT Units.* From Units,DataItem" & _
                " Where Units.Unit = DataItem.UnitOfMeasurement" & _
                " AND DataItem.ClinicalTrialID =" & vClinicalTrialId & _
                " UNION " & _
                " SELECT DISTINCT Units.* FROM Units,ClinicalTest,CTC,Studydefinition" & _
                " Where Units.Unit = ClinicalTest.Unit" & _
                " AND ClinicalTest.ClinicalTestCode = CTC.ClinicalTestCode" & _
                " AND StudyDefinition.CTCSchemeCode = CTC.CTCSchemeCode" & _
                " AND StudyDefinition.ClinicalTrialId=" & vClinicalTrialId & ""
  '   RJCW 11/10/2001 BUGFIX 031
  '   The extra sql statements added here are necessary to find any missing units
  '   that have conversion factors set up. If a conversion factor is set-up on a unit
  '   the linked unit from the conversion factor must also be exported. This is not
  '   necessary for CCT units
  sSQL = sSQL & " UNION " & _
                " SELECT UnitConversionFactors.FromUnit AS Unit,UnitConversionFactors.UnitClass" & _
                " FROM UnitConversionFactors " & _
                " WHERE Exists (SELECT DISTINCT Units.* FROM units, DataItem" & _
                "               WHERE units.Unit = DataItem.UnitOfMeasurement" & _
                "               AND DataItem.ClinicalTrialID =" & vClinicalTrialId & _
                "               AND units.unit = UnitConversionFactors.ToUnit)" & _
                "  UNION " & _
                " SELECT UnitConversionFactors.ToUnit as Unit,UnitConversionFactors.UnitClass" & _
                " FROM UnitConversionFactors" & _
                " WHERE Exists (SELECT DISTINCT Units.* FROM units, DataItem" & _
                "               WHERE units.Unit = DataItem.UnitOfMeasurement" & _
                "               AND   DataItem.ClinicalTrialID =" & vClinicalTrialId & _
                "               AND   units.unit = UnitConversionFactors.Fromunit)"

    'UnitConversionFactors
    ElseIf msSegmentId = 560 Then
        sSQL = "Select DISTINCT UnitConversionFactors.* From UnitConversionFactors,DataItem" & _
                " Where DataItem.UnitOfMeasurement" & _
                " IN (UnitConversionFactors.FromUnit,UnitConversionFactors.ToUnit)" & _
                " AND DataItem.ClinicalTrialID =" & vClinicalTrialId
    
    'ClinicalTestGroup
    ElseIf msSegmentId = 610 Then
        sSQL = "Select DISTINCT ClinicalTestGroup.* From ClinicalTestGroup,ClinicalTest,DataItem" & _
                " Where ClinicalTestGroup.ClinicalTestGroupCode = ClinicalTest.ClinicalTestGroupCode" & _
                " AND ClinicalTest.ClinicalTestCode = DataItem.ClinicalTestCode" & _
                " AND DataItem.ClinicalTrialID =" & vClinicalTrialId & _
                " UNION " & _
                " SELECT DISTINCT ClinicalTestGroup.* From ClinicalTest,ClinicalTestGroup,CTC,StudyDefinition" & _
                " Where ClinicalTestGroup.ClinicalTestGroupCode = ClinicalTest.ClinicalTestGroupCode" & _
                " AND ClinicalTest.ClinicalTestCode = CTC.ClinicalTestCode" & _
                " AND StudyDefinition.CTCSchemeCode = CTC.CTCSchemeCode" & _
                " AND StudyDefinition.ClinicalTrialID=" & vClinicalTrialId
    
    'ClinicalTest
    ElseIf msSegmentId = 620 Then
       sSQL = "Select DISTINCT Clinicaltest.* From ClinicalTest,DataItem" & _
                " Where ClinicalTest.ClinicalTestCode = DataItem.ClinicalTestCode" & _
                " AND DataItem.ClinicalTrialID =" & vClinicalTrialId & _
                " UNION " & _
                " SELECT DISTINCT Clinicaltest.* From ClinicalTest,CTC,StudyDefinition" & _
                " Where ClinicalTest.ClinicalTestCode = CTC.ClinicalTestCode" & _
                " AND StudyDefinition.CTCSchemeCode = CTC.CTCSchemeCode" & _
                " AND StudyDefinition.ClinicalTrialID=" & vClinicalTrialId
                
    'CTCScheme
    ElseIf msSegmentId = 630 Then
        sSQL = "Select CTCScheme.* From CTCScheme,StudyDefinition" & _
                " Where StudyDefinition.CTCSchemeCode = CTCScheme.CTCSchemeCode" & _
                " AND StudyDefinition.ClinicalTrialID =" & vClinicalTrialId
    
    'CTC
    ElseIf msSegmentId = 640 Then
        sSQL = "Select CTC.* From CTC,StudyDefinition" & _
                " Where StudyDefinition.CTCSchemeCode = CTC.CTCSchemeCode" & _
                " AND StudyDefinition.ClinicalTrialID =" & vClinicalTrialId
    
    'files that need to be exported in full (i.e. ValidationAction)
    Else
       sSQL = " SELECT * FROM " & msTablename
    End If
    
    EvaluateExceptionalSDDTables = sSQL

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "EvaluateExceptionalSDDTables", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function
'---------------------------------------------------------------------------------
Public Function EvaluateExceptionalLDDTables(sLaboratory As String, _
                                            sSegmentId As String, _
                                            sTableName As String) As String
'---------------------------------------------------------------------------------
'Ash 28/08/2001 Routine to handle new exportLDD due to introduction of
'Normal ranges,CTC schemes
'---------------------------------------------------------------------------------
Dim sSQL As String

    sSQL = ""
    
    On Error GoTo ErrHandler
    
     'Units
        If sSegmentId = 550 Then
            sSQL = "Select DISTINCT Units.* From Units,ClinicalTest,NormalRange" & _
                    " Where Units.Unit = ClinicalTest.Unit" & _
                    " AND ClinicalTest.ClinicalTestCode = NormalRange.ClinicalTestCode" & _
                    " AND NormalRange.LaboratoryCode ='" & sLaboratory & "'"
                    
     'ClinicalTestGroup
        ElseIf sSegmentId = 610 Then
            sSQL = "Select DISTINCT ClinicalTestGroup.* From ClinicalTestGroup,ClinicalTest,NormalRange" & _
                    " WHERE ClinicalTestGroup.ClinicalTestGroupCode = ClinicalTest.ClinicalTestGroupCode" & _
                    " AND ClinicalTest.ClinicalTestCode = NormalRange.ClinicalTestCode" & _
                    " AND NormalRange.LaboratoryCode ='" & sLaboratory & "'"
    
    'ClinicalTest
        ElseIf sSegmentId = 620 Then
            sSQL = "Select DISTINCT ClinicalTest.* From ClinicalTest,NormalRange" & _
                    " WHERE ClinicalTest.ClinicalTestCode = NormalRange.ClinicalTestCode" & _
                    " AND NormalRange.LaboratoryCode ='" & sLaboratory & "'"
        
     
    'files/tables that need to be exported in full (currently no tables match this)
        Else
           sSQL = " SELECT * FROM " & sTableName
        End If
        
        EvaluateExceptionalLDDTables = sSQL


Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "EvaluateExceptionalLDDTables", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------------------
Private Sub MarkSubjectInvalid(ByVal nPersonId As Integer, _
                              ByVal sTrialSite As String, _
                              ByVal lTrialId As Long)
                              
'---------------------------------------------------------------------------------
'ZA 15/02/02 - Mark this subject in ArezzoToken table to invalid
'---------------------------------------------------------------------------------
Dim sToken As String
Dim sArezzoQuery As String

    'create token (primary key for ArezzoToken table) combining subjectid, site and study id
    sToken = nPersonId & "," & sTrialSite & "," & lTrialId
    
    MacroADODBConnection.Execute "Update ArezzoToken set DBToken = 0 where ArezzoID = '" & sToken & "'"
    
End Sub

'---------------------------------------------------------------------
Public Sub RemoveSubjectLock(lStudyId As Long, sSite As String, lSubjectId As Long, sToken As String)
' DPH 11/04/2002 - Added in to control error handling on removing a subject locks
'---------------------------------------------------------------------
On Error GoTo LockFailed

    Call MACROLOCKBS30.UnlockSubject(gsADOConnectString, sToken, lStudyId, sSite, lSubjectId)
    
Exit Sub
LockFailed:
    ' Ignore Errors
End Sub

'---------------------------------------------------------------------
Public Function GetSubjectLock(sUser As String, lStudyId As Long, sSite As String, lSubjectId As Long) As String
' DPH 11/04/2002 - Added in to control error handling on creating subject locks
'REM 16/10/03 - Added MACROLOCKBS30.DBLocked.dblEFormInstance case to check for open eform
'---------------------------------------------------------------------
Dim sLockDetails As String
Dim sMSG As String
Dim sToken As String
    
    On Error GoTo LockFailed

    sToken = MACROLOCKBS30.LockSubject(gsADOConnectString, sUser, lStudyId, sSite, lSubjectId)
    Select Case sToken
    Case MACROLOCKBS30.DBLocked.dblStudy
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblSubject, MACROLOCKBS30.DBLocked.dblEFormInstance
        sToken = ""
    Case Else
        'hurrah, we have a lock
    End Select
    
    GetSubjectLock = sToken
    
Exit Function
LockFailed:
    Err.Raise Err.Number, , Err.Description & "|" & "clsExchange.GetSubjectLock"
End Function

'---------------------------------------------------------------------
Public Function ImportPatientZIP(ByVal sImportFile As String) As Boolean
'---------------------------------------------------------------------
' Extract files from a Patient ZIP file using XCeed
'---------------------------------------------------------------------
' REVISIONS
' DPH 25/04/2002 - Made into function returning boolean
'---------------------------------------------------------------------
Dim sFileName As String

    On Error GoTo ErrHandler

    sFileName = StripFileNameFromPath(sImportFile)
    
    gLog gsIMPORT_PAT_ZIP, "Extracting " & sFileName & " starting"
    
    'clear the CabExtract folder before starting the Import
    ClearCabExtractFolder
    
    ' Make sure file exists before opening
    If FolderExistence(sImportFile, True) Then
        
        ' Error handler for Unzip process
        On Error GoTo UnZipErr
        
        ' Extract Files from ZIP using XCeed
        Call UnZipFiles(gsCAB_EXTRACT_LOCATION, sImportFile)
    
        On Error GoTo ErrHandler
        
        gLog gsIMPORT_PAT_ZIP, "Extracting " & sFileName & " completed"
    
        ImportPatientZIP = True
        
    Else
        gLog gsIMPORT_PAT_ZIP, "Extracting from " & sFileName & " failed as file doesn't exist"
        ImportPatientZIP = False
    End If
    
Exit Function

UnZipErr:
    gLog gsIMPORT_PAT_ZIP, "Extracting " & sFileName & " failed. XCeed Error Number " & Err.Number & ". AutoImport Error."
    ImportPatientZIP = False
Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportPatientZIP", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Public Function ImportStudyDefinitionZIP(ByVal sImportFile As String)
'---------------------------------------------------------------------
' Extract files from a Study Definition ZIP file using XCeed
'---------------------------------------------------------------------
' REVISIONS
'---------------------------------------------------------------------
Dim sFileName As String

    On Error GoTo ErrHandler

    sFileName = StripFileNameFromPath(sImportFile)
    
    gLog gsIMPORT_STUDY_ZIP, "Extracting " & sFileName & " starting"
    
    'clear the CabExtract folder before starting the Import
    ClearCabExtractFolder
        
    ' Make sure file exists before opening
    If FolderExistence(sImportFile, True) Then
        
        ' Error handler for Unzip process
        On Error GoTo UnZipErr
        
        ' Extract Files from ZIP using XCeed
        Call UnZipFiles(gsCAB_EXTRACT_LOCATION, sImportFile)
        
        On Error GoTo ErrHandler
        
        gLog gsIMPORT_STUDY_ZIP, "Extracting " & sFileName & " completed"
    
    Else
        gLog gsIMPORT_STUDY_ZIP, "Extracting from " & sFileName & " failed as file doesn't exist"
    End If

Exit Function

UnZipErr:
    gLog gsIMPORT_STUDY_ZIP, "Extracting " & sFileName & " failed. XCeed Error Number " & Err.Number & ". "
Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportStudyDefinitionZIP", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Public Function ImportLDDZIP(ByVal sImportFile As String)
'---------------------------------------------------------------------
' Extract files from a Laboratory Definition ZIP file using XCeed
'---------------------------------------------------------------------
' REVISIONS
'---------------------------------------------------------------------
Dim sFileName As String

    On Error GoTo ErrHandler

    sFileName = StripFileNameFromPath(sImportFile)
    
    gLog gsIMPORT_LDD_ZIP, "Extracting " & sFileName & " starting"
    
    'clear the CabExtract folder before starting the Import
    ClearCabExtractFolder
    
    ' Make sure file exists before opening
    If FolderExistence(sImportFile, True) Then
        
        ' Error handler for Unzip process
        On Error GoTo UnZipErr
        
        ' Extract Files from ZIP using XCeed
        Call UnZipFiles(gsCAB_EXTRACT_LOCATION, sImportFile)
        
        On Error GoTo ErrHandler
        
        gLog gsIMPORT_LDD_ZIP, "Extracting " & sFileName & " completed"
        
    Else
        gLog gsIMPORT_LDD_ZIP, "Extracting from " & sFileName & " failed as file doesn't exist"
    End If

Exit Function

UnZipErr:
    gLog gsIMPORT_LDD_ZIP, "Extracting " & sFileName & " failed. XCeed Error Number " & Err.Number & ". AutoImport Error (Lab)."
Exit Function

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ImportLDDZIP", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

'---------------------------------------------------------------------
Public Sub ExportPatientZIP(sZipFileName As String, asZIPFileList() As String)
'---------------------------------------------------------------------
'Packages a list of Patient data PRD files and Patient ProformaState files
'into a single ZIP file using XCeed
'Replaces ExportPatientCAB
'---------------------------------------------------------------------
' REVISIONS
'---------------------------------------------------------------------
Dim sZIPfile As String
Dim i As Long

On Error GoTo ErrHandler

    gLog gsEXPORT_PAT_ZIP, "Session " & sZipFileName & " starting"
    
    ' Create ZIP file name
    sZIPfile = gsOUT_FOLDER_LOCATION & sZipFileName & ".zip"
    
    ' Make sure folder exists before opening
    If FolderExistence(sZIPfile) Then
            
        If FolderExistence(sZIPfile, True) Then
            ' delete file
            Kill sZIPfile
        End If
            
        ' Error handler for Unzip process
        On Error GoTo ZipErr
        
        ' Create ZIP file
        Call ZipFiles(asZIPFileList, sZIPfile)
    
        On Error GoTo ErrHandler
        
        'Kill off the files that have been compacted into a ZIP file
        For i = 0 To UBound(asZIPFileList)
            Kill asZIPFileList(i)
            gLog gsEXPORT_PAT_ZIP, "Removing " & StripFileNameFromPath(asZIPFileList(i))
        Next i
                
        gLog gsEXPORT_PAT_ZIP, "Session " & sZipFileName & " completed"
    Else
        gLog gsEXPORT_PAT_ZIP, "Session " & sZipFileName & " failed as could not create ZIP file directory " & gsOUT_FOLDER_LOCATION
    End If

Exit Sub
ZipErr:
    gLog gsEXPORT_PAT_ZIP, "Session " & StripFileNameFromPath(sZipFileName) & " failed. XCeed Error Number " & Err.Number

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ExportPatZIP", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub

'---------------------------------------------------------------------
Public Sub ExportStudyZIP(sZipFileName As String, asZIPFileList() As String, _
            Optional sFileSpec As String = "")
'---------------------------------------------------------------------
'Packages a list of study definition files
'into a single ZIP file using XCeed
'Replaces ExportStudyCAB
'---------------------------------------------------------------------
' REVISIONS
' NCJ 26 Jun 06 - Added optional FileSpec argument for final zip file (assumed full path/file spec)
'---------------------------------------------------------------------
Dim sZIPfile As String
Dim i As Long

    On Error GoTo ErrLabel

    gLog gsEXPORT_STUDY_ZIP, "Session " & sZipFileName & " starting"
    
    ' Create ZIP file name
    ' NCJ 26 Jun 06 - Use given file spec if supplied
    If sFileSpec > "" Then
        sZIPfile = sFileSpec
    Else
        ' Otherwise use Out Folder and given file name
        sZIPfile = gsOUT_FOLDER_LOCATION & sZipFileName & ".zip"
    End If
    
    
    ' Make sure folder exists before opening
    If FolderExistence(sZIPfile) Then
                
        If FolderExistence(sZIPfile, True) Then
            ' delete file
            Kill sZIPfile
        End If
        
        ' Error handler for Unzip process
        On Error GoTo ZipErr
        
        ' Create ZIP file
        Call ZipFiles(asZIPFileList, sZIPfile)
    
        On Error GoTo ErrLabel
        
        'Kill off the first file in asZIPFileList which will be an SDD file
        'the rest of asZIPFileList refers to Study Documents and grahics that should not be deleted
        Kill asZIPFileList(0)
                
        gLog gsEXPORT_STUDY_ZIP, "Session " & sZipFileName & " completed"
    Else
        gLog gsEXPORT_STUDY_ZIP, "Export to " & sZIPfile & " failed as could not create ZIP file directory "
    End If
    
Exit Sub
ZipErr:
    gLog gsEXPORT_STUDY_ZIP, "Export to " & sZIPfile & " failed. XCeed Error Number " & Err.Number

Exit Sub
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|clsExchange.ExportStudyZip"

End Sub

'---------------------------------------------------------------------
Public Sub ExportLDDZIP(sZipFileName As String, sZIPFileList() As String)
'---------------------------------------------------------------------
'Packages a list of laboratory definition files
'into a single ZIP file using XCeed
'Replaces ExportLDDCAB
'---------------------------------------------------------------------
' REVISIONS
' DPH 24/04/2002 - Added check for existing ZIP file & delete it else
'                   files get added to existing ZIP rather than replacing wholely
'---------------------------------------------------------------------
Dim sZIPfile As String
Dim i As Long

On Error GoTo ErrHandler
    
    gLog gsEXPORT_LDD_ZIP, "Session " & sZipFileName & " starting"
        
    ' Create ZIP file name
    sZIPfile = gsOUT_FOLDER_LOCATION & sZipFileName & ".zip"
    
    ' DPH 17/10/2001 Make sure folder exists before opening
    If FolderExistence(sZIPfile) Then
            
        If FolderExistence(sZIPfile, True) Then
            ' delete file
            Kill sZIPfile
        End If
        
        ' Error handler for Unzip process
        On Error GoTo ZipErr
        
        ' Create ZIP file
        Call ZipFiles(sZIPFileList, sZIPfile)

        On Error GoTo ErrHandler
                
        'Kill off the first file in sZIPFileList which will be an SDD file
        'the rest of sZIPFileList refers to Study Documents and grahics that should not be deleted
        Kill sZIPFileList(0)
                
        gLog gsEXPORT_LDD_ZIP, "Session " & sZipFileName & " completed"
    Else
        gLog gsEXPORT_LDD_ZIP, "Session " & sZipFileName & " failed as could not create folder " & gsOUT_FOLDER_LOCATION
    End If

Exit Sub
ZipErr:
    gLog gsEXPORT_LDD_ZIP, "Session " & StripFileNameFromPath(sZipFileName) & " failed. XCeed Error Number " & Err.Number

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ExportLDDZIP", "clsExchange")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Sub

'--------------------------------------------------------------------------------
Private Function VersionNumberToLong(sVersion As String) As Long
'--------------------------------------------------------------------------------
'create a long version of the version number for numerical comparison
'--------------------------------------------------------------------------------

    VersionNumberToLong = Val(Format(Split(sVersion, ".")(0), "00") & _
            Format(Split(sVersion, ".")(1), "00") & _
            Format(Split(sVersion, ".")(2), "0000"))

End Function

'---------------------------------------------------------------------
Public Sub RemoveStudyLock(lStudyId As Long, sToken As String)
' DPH 11/04/2002 - Added in to control error handling on removing a study lock
'---------------------------------------------------------------------

    On Error GoTo LockFailed
    
    Call MACROLOCKBS30.UnlockStudy(gsADOConnectString, sToken, lStudyId)
    
Exit Sub
LockFailed:
    ' Ignore Errors
End Sub

'---------------------------------------------------------------------
Public Function GetStudyLock(sUser As String, lStudyId As Long) As String
' DPH 11/04/2002 - Added in to control error handling on creating subject locks
' REM 16/10/03 - Added MACROLOCKBS30.DBLocked.dblEFormInstance case to check for open eform
'---------------------------------------------------------------------
Dim sLockDetails As String
Dim sToken As String

    On Error GoTo LockFailed

    sToken = MACROLOCKBS30.LockStudy(gsADOConnectString, sUser, lStudyId)
    Select Case sToken
    Case MACROLOCKBS30.DBLocked.dblStudy
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblSubject, MACROLOCKBS30.DBLocked.dblEFormInstance
        sToken = ""
    Case Else
        'hurrah, we have a lock
    End Select
    
    GetStudyLock = sToken
    
Exit Function
LockFailed:
    Err.Raise Err.Number, , Err.Description & "|" & "clsExchange.GetStudyLock"
End Function

'----------------------------------------------------------------------------------------'
Private Sub ProcessLFRollBacks(sStudyName As String, sSite As String, lSubjectId As Long)
'----------------------------------------------------------------------------------------'
' NCJ 9 Jan 03
' Process any unprocessed Lock/Freeze rollbacks that are associated with this subject
'----------------------------------------------------------------------------------------'
Dim oLockFreezer As LockFreeze

    On Error GoTo ErrHandler
    
    Set oLockFreezer = New LockFreeze
    
    Call oLockFreezer.ProcessSubjectRollBacks(MacroADODBConnection, sStudyName, sSite, lSubjectId)
    
    Set oLockFreezer = Nothing
   
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modAutoImportMain.ProcessLFRollBacks"
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub ProcessLFMessages(sStudyName As String, sSite As String, lSubjectId As Long, sZipFileName As String)
'----------------------------------------------------------------------------------------'
' NCJ 24 Dec 02
' Process the Lock/Freeze messages that are associated with this subject's auto import file
' We read the information from the DataImport table
' and delete it when we've finished
'----------------------------------------------------------------------------------------'
Dim oLockFreezer As LockFreeze
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim lLastMessageId As Long
Dim sSQLWhere As String

    On Error GoTo ErrHandler
    
    Set oLockFreezer = New LockFreeze
    
    lLastMessageId = 0
    
    sSQLWhere = " WHERE ClinicalTrialName = '" & sStudyName & "'" _
            & " AND TrialSite = '" & sSite & "'" _
            & " AND PersonId = " & lSubjectId _
            & " AND DataFileName = '" & sZipFileName & "'"
            
    ' Get the last used LF message id
    sSQL = "SELECT LastLFMessageId FROM DataImport" & sSQLWhere
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
    
    ' NB In manual Subject Data Import, there won't be a row in the DataImport table
    If rsTemp.RecordCount > 0 Then
        If Not IsNull(rsTemp!LastLFMessageId) Then
            lLastMessageId = rsTemp!LastLFMessageId
            ' Deal with the messages if there are any
            If lLastMessageId > 0 Then
                Call oLockFreezer.DealWithUnprocessedMessages(MacroADODBConnection, _
                                        sStudyName, sSite, lSubjectId, lLastMessageId)
                gLog gsAUTOIMPORT, "Handled Lock/Freeze messages for data file " & sZipFileName
            End If
            ' Say we've dealt with this file
            sSQL = "DELETE FROM DataImport" & sSQLWhere
            MacroADODBConnection.Execute sSQL
        
        End If
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    Set oLockFreezer = Nothing
   
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modAutoImportMain.ProcessLFMessages"
    
End Sub


Private Sub ListParams(cmd As ADODB.Command)
Dim pName As String
Dim lRows As Long
Dim pType As String

    For lRows = 0 To cmd.Parameters.Count - 1
        Select Case cmd.Parameters(lRows).Type
            Case adArray: pType = "AdArray"
            Case adBigInt: pType = "adBigInt"
            Case adBinary: pType = "adBinary"
            Case adBoolean: pType = "adBoolean"
            Case adBSTR: pType = "adBSTR"
            Case adChapter: pType = "adChapter"
            Case adChar: pType = "adChar"
            Case adCurrency: pType = "adCurrency"
            Case adDate: pType = "adDate"
            Case adDBDate: pType = "adDBDate"
            Case adDBTime: pType = "adDBTime"
            Case adDBTimeStamp: pType = "adDBTimeStamp"
            Case adDecimal: pType = "adDecimal"
            Case adDouble: pType = "adDouble"
            Case adEmpty: pType = "adEmpty"
            Case adError: pType = "adError"
            Case adFileTime: pType = "adFileTime"
            Case adGUID: pType = "adGUID"
            Case adIDispatch: pType = "adIDispatch"
            Case adInteger: pType = "adInteger"
            Case adIUnknown: pType = "adIUnknown"
            Case adLongVarBinary: pType = "adLongVarBinary"
            Case adLongVarChar: pType = "adLongVarChar"
            Case adLongVarWChar: pType = "adLongVarWChar"
            Case adNumeric: pType = "adNumeric"
            Case adPropVariant: pType = "adPropVariant"
            Case adSingle: pType = "adSingle"
            Case adSmallInt: pType = "adSmallInt"
            Case adTinyInt: pType = "adTinyInt"
            Case adUnsignedBigInt: pType = "adUnsignedBigInt"
            Case adUnsignedInt: pType = "adUnsignedInt"
            Case adUnsignedSmallInt: pType = "adUnsignedSmallInt"
            Case adUnsignedTinyInt: pType = "adUnsignedTinyInt"
            Case adUserDefined: pType = "adUserDefined"
            Case adVarBinary: pType = "adVarBinary"
            Case adVarChar: pType = "adVarChar"
            Case adVariant: pType = "adVariant"
            Case adVarNumeric: pType = "adVarNumeric"
            Case adVarWChar: pType = "adVarWChar"
            Case adWChar: pType = "adWChar"
        End Select
        pName = cmd.Parameters(lRows).Name
        Do While Len(pName) < 20
            pName = pName & " "
        Loop
        'Debug.Print lRows, pName, " Type = " & cmd.Parameters(lRows).Type, pType, " Size = " & cmd.Parameters(lRows).Size & "(" & cmd.Parameters(lRows).Precision & "," & cmd.Parameters(lRows).NumericScale & ")", " Dir = " & cmd.Parameters(lRows).Direction, "Val = " & cmd.Parameters(lRows).Value
    Next

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'
' 27/02/2003 Ronald Schravendeel
' Initialize class level command objects, so that they are only initialized once, and
' stay available for the ImportPRD routine
'----------------------------------------------------------------------------------------'
    InitStoredProcedures
End Sub
'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'
' 27/02/2003 Ronald Schravendeel
'----------------------------------------------------------------------------------------'
    CleanupStoredProcedures
End Sub
'----------------------------------------------------------------------------------------'
Private Sub CleanupStoredProcedures()
'----------------------------------------------------------------------------------------'
' 27/02/2003 Ronald Schravendeel
'
' Cleanup the stored procedure command objects
'----------------------------------------------------------------------------------------'

    Set cmdImportTrialSubject = Nothing
    Set cmdImportVisitInstance = Nothing
    Set cmdImportCRFPageInstance = Nothing
    Set cmdImportDIR = Nothing
    Set cmdImportDIRH = Nothing
    'Mo 21/11/2005 COD0220
    If gbClinicalCoding Then
        Set cmdImportCodingHistory = Nothing
    End If
    
End Sub


'----------------------------------------------------------------------------------------'
Private Sub InitStoredProcedures()
'----------------------------------------------------------------------------------------'
' 27/02/2003 Ronald Schravendeel
'----------------------------------------------------------------------------------------'

    '---------------------------------------------------------------------
    ' RS 04/02/2003 Import using Stored Procedure
    ' NOTE: getting Parameter Info from database is very slow, so this is done
    ' only at class initialization
    '---------------------------------------------------------------------

    Set cmdImportDIR = New ADODB.Command            ' Stored Procedure: SP_MACRO_IMP_DIR
    With cmdImportDIR
        .ActiveConnection = MacroADODBConnection
        .CommandText = "sp_macro_imp_dir"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
    End With
    
    Set cmdImportDIRH = New ADODB.Command            ' Stored Procedure: SP_MACRO_IMP_DIRH
    With cmdImportDIRH
        .ActiveConnection = MacroADODBConnection
        .CommandText = "sp_macro_imp_dirh"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
    End With
    
    Set cmdImportTrialSubject = New ADODB.Command   ' Stored Procedure: SP_MACRO_IMP_TRIALSUBJECT
    With cmdImportTrialSubject
        .ActiveConnection = MacroADODBConnection
        .CommandText = "sp_macro_imp_trialsubject"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
    End With
    
    Set cmdImportVisitInstance = New ADODB.Command      ' Stored Procedure: SP_MACRO_IMP_VISITINSTANCE
    With cmdImportVisitInstance
        .ActiveConnection = MacroADODBConnection
        .CommandText = "sp_macro_imp_visitinstance"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
    End With
    
    Set cmdImportCRFPageInstance = New ADODB.Command    ' Stored Procedure: SP_MACRO_IMP_CRFPAGEINSTANCE
    With cmdImportCRFPageInstance
        .ActiveConnection = MacroADODBConnection
        .CommandText = "sp_macro_imp_crfpageinstance"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
    End With
    
    'Mo 21/11/2005 COD0220
    If gbClinicalCoding Then
        Set cmdImportCodingHistory = New ADODB.Command            ' Stored Procedure: SP_MACRO_IMP_CodingHistory
        With cmdImportCodingHistory
            .ActiveConnection = MacroADODBConnection
            .CommandText = "sp_macro_imp_codinghistory"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
        End With
    End If

End Sub

'---------------------------------------------------------------------
Private Function DictionaryIdToNameVersion(nDictionaryId As Integer) As String
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sNameVersion As String

    sSQL = "SELECT DictionaryName, DictionaryVersion" _
        & " FROM Dictionaries" _
        & " WHERE DictionaryId = " & nDictionaryId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 1 Then
        sNameVersion = RemoveNull(rsTemp!DictionaryName) & "|" & RemoveNull(rsTemp!DictionaryVersion)
    Else
        sNameVersion = "DummyName|DummyVersion"
    End If
    
    Set rsTemp = Nothing
    
    DictionaryIdToNameVersion = sNameVersion

End Function

'---------------------------------------------------------------------
Private Function DictionaryNameVersionToId(sNameVersion As String) As String
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sDictionaryName As String
Dim sDictionaryVersion As String
Dim sDictionaryId As String
Dim asArray() As String
    
    asArray = Split(sNameVersion, "|")
    
    If UBound(asArray) <> 1 Then
        'its not a valid NAME|VERSION entry, use DictionaryId 1
        sDictionaryId = MinDictionaryID
    Else
        sDictionaryName = asArray(0)
        sDictionaryVersion = asArray(1)
        If (Not IsNull(sDictionaryName)) And (Not IsNull(sDictionaryVersion)) Then
            sSQL = "SELECT DictionaryId" _
                & " FROM Dictionaries" _
                & " WHERE DictionaryName = '" & sDictionaryName & " '" _
                & " AND DictionaryVersion = '" & sDictionaryVersion & "'"
            Set rsTemp = New ADODB.Recordset
            rsTemp.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            If rsTemp.RecordCount = 1 Then
                sDictionaryId = RemoveNull(rsTemp!DictionaryId)
            Else
                sDictionaryId = MinDictionaryID
            End If
            Set rsTemp = Nothing
        Else
            'sDictionaryName or sDictionaryVersion are null, use DictionaryId 1
            sDictionaryId = MinDictionaryID
        End If
    End If
     
    DictionaryNameVersionToId = sDictionaryId
        
End Function

'---------------------------------------------------------------------
Private Function DictionariesExist() As Boolean
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bDictionariesExist As Boolean

    sSQL = "SELECT DictionaryId FROM Dictionaries"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount > 0 Then
        bDictionariesExist = True
    Else
        bDictionariesExist = False
    End If
    
    Set rsTemp = Nothing
    
    DictionariesExist = bDictionariesExist
    
End Function

'---------------------------------------------------------------------
Private Function MinDictionaryID() As Integer
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim nMinDictionaryID As Integer

    sSQL = "SELECT MIN(DictionaryId) as MinDictionaryId FROM Dictionaries"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    nMinDictionaryID = rsTemp!MinDictionaryID
    
    rsTemp.Close
    Set rsTemp = Nothing

    MinDictionaryID = nMinDictionaryID

End Function

'---------------------------------------------------------------------
Private Function GetVersionFromDB() As String
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sVersion As String


    sSQL = "SELECT * FROM MACROControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    sVersion = rsTemp!MACROVersion & "." & rsTemp!BuildSubVersion
    
    rsTemp.Close
    Set rsTemp = Nothing

    GetVersionFromDB = sVersion

End Function
