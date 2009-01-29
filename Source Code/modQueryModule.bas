Attribute VB_Name = "modQueryModule"
'----------------------------------------------------------------------------------------'
' File:         modQueryModule
' Copyright:    InferMed Ltd. 2000. All Rights Reserved
' Author:       Mo Morris, February 2002
' Purpose:      Contains variblee declarations and facilities required by the Query Module
'----------------------------------------------------------------------------------------'
'   Revisions:
'
'   Mo  2/5/2002    Changes to CSVCommasAndQuotes so that it correctly handles fields
'                   containing double quotes.
'   Mo  3/5/2002    Changes to OutputToAccess so that the ReplaceQuotes function is now
'                   called to run over category (that can contain single quotes)
'   Mo  7/5/2002    Changes stemming from Label/PersonId being switched
'                   from a single field to 2 separate fields.
'                   gbDisplayLabelPersonId replaced by gbDisplayLabel and gbDisplayPersonId
'                   mbSubjectLabels removed.
'                   Minor changes to OutputToAccess and OuputToCSV.
'                   OutputToSAS code has been added, but it is incomplete and is not called.
'   Mo  15/7/2003   OutPutToSATA together with STATADateFormat, STATADateTest, STATAReplaceSection
'                   WriteCatCodesToSTATA, AssessCategoryCodes, DataItemDetails And WriteCatCodesToSAS
'                   added. STATA output is now completed.
'   Mo  17/11/2004  Bug 2411 - OutputToSTATA now works in 2 ways:-
'                   "Standard"  Uses ddmmmyyyy Standard dates (e.g. 01jan2004)
'                   "Float"     Uses ddmmyyyy Float dates (e.g. 01012004 for 1 January 2004)
'                   Changes have been made to OutputToSTATA, STATADateFormat
'   Mo  25/10/2005  COD0040 - Changes around the new Thesaurus Data Item Type
'   Mo  11/1/2006   Bug 2671 - Change STATA Standard dates from %d strings to %8.0g numbers.
'                   Changes to OutputToSTATA and STATADateFormat
'   Mo  11/1/2006   Bug 2672 - Change the STATA replace characters: -
'                   Leave "." for STATA to use on missing numerics
'                   ".a" changes from "-2" to "-1"
'                   ".b" changes from "-3" to "-2"
'                   ".c" changes from "-4" to "-3"
'                   ".d" changes from "-5" to "-4"
'                   ".e" changes from "-6" to "-5"
'                   ".f" changes from "-7" to "-6"
'                   ".g" changes from "-8" to "-7"
'                   ".h" changes from "-9" to "-8"
'                   ".i" used by "-9"
'   Mo  25/1/2006   Bug 2671 - correction to STATADateFormat
'   Mo  26/5/2006   Bug 2738 - Add Question Name/Description and Type to SAS and STATA QLU files
'   Mo  30/5/2006   Bug 2668 - Option to exclude Subject Label from saved output files
'   Mo  2/6/2006    Bug 2737 - Add Question Short Code length to the Options Window
'   Mo  9/6/2006    Bug 2739 - Changes to STATA export of Special Values on strings
'   Mo  1/8/2006    Bug 2775 - Checking for single digit numeric questions and setting them to two digit fields
'                   that are capable of holding a special value (-1 to -9) when saved in QM output files (STATA, SAS and csv)
'   Mo 21/8/2006    Bug 2784 - Query Module needs to refer to all Category Codes (Active & Inactive).
'                   Query Module calls to gdsDataValues replaced by calls to new sub gdsDataValuesALL.
'                   AssessCategoryCodes, OutputToAccess, OutputToCSV, WriteCatCodesToSAS and WriteCatCodesToSTATA
'                   now check for no category codes existing.
'   Mo 18/10/2006   Bug 2822 - Make MACRO Query Module comply with Partial Dates.
'   Mo  1/11/2006   Bug 2795 - "Precede SAS informats with colons" option added.
'   Mo  2/11/2006   Bug 2797 - Make SAS output use Long Codes when specified.
'                   Long Codes to be of the form VisitCode_FormCode-QuestionCode.
'                   Changes made to OutputToSAS.
'   Mo  3/11/2006   Bug 2834 - Make Query Module's save in Batch Data Entry format, save dates in original format.
'                   Changes made to OutputToMACROBD. New function GetSingleResponse added.
'   Mo  10/1/2006   Bug 2866 - Increase Query Module STATA output max string length from 80 to 244 characters
'   Mo  31/1/2007   Bug 2873 - Real and LabTest response data to be placed in Decimal not Single fields
'                   a change from adSingle to adDecimal
'   Mo  2/4/2007    MRC15022007 - Query Module Batch Facilities
'   Mo  24/5/2007   Bug 2913 - The addition of a _Format.txt file to the SAS output files.
'                   Changes made to OutputToSAS.
'                   Marie-Gabrielle Dondon has requested that an additional fourth file  be added to the
'                   Query Module's SAS output. The additional file is to be similar to the _Type.txt file
'                   except that it should only contain date, time and date/time questions and should not have
'                   colons preceding the informats. The file should have an _Format.txt name. When the MACRO
'                   SAS output files are imported into SAS the date/time fields are converted into a numeric
'                   value, the additional_Format.txt file entries will tell SAS how to format/display these
'                   numeric date/time values.
'                   Example of a _Format.txt file:-
'                       QDob ddmmyy10.
'                       Qdate ddmmyy10.
'                       Qtime time.
'                       QDateTime datetime19.
'   Mo  9/10/2007   Bug 2941 - Increase the space allocated for a column number in a STATA.dct file from 4 to 6.
'                   Changes made to OutputToSTATA.
'----------------------------------------------------------------------------------------'

Option Explicit

Public gnRegFormLeft As Integer
Public gnRegFormTop As Integer
Public gnRegFormWidth As Integer
Public gnRegFormHeight As Integer
Public gnSelectFilterBarTop As Integer
Public gnFilterOutputBarTop As Integer

Public gsSVMissing As String
Public gsSVUnobtainable As String
Public gsSVNotApplicable As String
Public gnOutPutType As Integer
Public gbOutputCategoryCodes As Boolean
Public gbDisplayStudyName As Boolean
Public gbDisplaySiteCode As Boolean
Public gbDisplayLabel As Boolean
Public gbDisplayPersonId As Boolean
Public gbDisplayVisitCycle As Boolean
Public gbDisplayFormCycle As Boolean
Public gbDisplayRepeatNumber As Boolean
Public gbSplitGrid As Boolean
Public gbUseShortCodes As Boolean
Public gbDisplayOutPut As Boolean
'Mo 30/5/2006 Bug 2668
Public gbExcludeLabel As Boolean
'Mo 2/6/2006 Bug 2737
Public gnShortCodeLength As Integer
'Mo 1/11/2006 Bug 2795
Public gbSASInformatColons As Boolean

'Mo 2/4/2007 MRC15022007
Public gsFileNamePath As String
Public gsFileNameText As String
Public gsFileNameStamp As String

'Mo 1/11/2006 Bug 2795
Public Enum eOutPutType
    CSV = 0
    Access = 1
    SPSS = 2
    SAS = 3
    STATA = 4
    MACROBD = 5
    STATAStandardDates = 6
    SASColons = 7
End Enum

Public gbQueryChanged As Boolean
'gbQueryChanged = true when changes have occured to a query
'gbQueryChanged = false when their are no unsaved changes in a query

Public gbQuerySaved As Boolean
'gbQuerySaved = false when a query has no name
'gbQuerySaved = true when a query has a name

'Mo 2/4/2007 MRC15022007
Public gbBatchQueryChanged As Boolean
'gbBatchQueryChanged = true when changes have occured to a query
'gbBatchQueryChanged = false when their are no unsaved changes in a query

Public gbBatchQuerySaved As Boolean
'gbBatchQuerySaved = false when a query has no name
'gbBatchQuerySaved = true when a query has a name

Public gbBatchQueryMode As Boolean
'gbBatchQueryMode = true when Batch Query is open
'gbBatchQueryMode = false when no Batch Query is open

Public Const mnMASK_RESPONSEVALUE As Integer = 1
Public Const mnMASK_COMMENTS As Integer = 2
Public Const mnMASK_CTCGRADE As Integer = 4
Public Const mnMASK_LABRESULT As Integer = 8
Public Const mnMASK_STATUS As Integer = 16
Public Const mnMASK_TIMESTAMP As Integer = 32
Public Const mnMASK_USERNAME As Integer = 64
Public Const mnMASK_VALUECODE As Integer = 128

Public Const mnMINFORMWIDTH As Integer = 12000
Public Const mnMINFORMHEIGHT As Integer = 8000

Public Enum LabResult
    Low = 1
    Normal = 2
    High = 3
End Enum

Public grsOutPut As ADODB.Recordset

Public glSelectedTrialId As Long

Private Const msCOMMA As String = ","
Private Const msDQUOTE As String = """"
Private Const msTAB As String = vbTab
Private Const msSPACE As String = " "

Public gColQuestionCodes As Collection

Public gColSTATADetails As Collection

Public gbCancelled As Boolean

Public gbNotDisplayedNotSaved As Boolean

'Mo 2/4/2007 MRC15022007, Optional parameter bShowDialogs added
'--------------------------------------------------------------------
Public Sub OutputToAccess(Optional bShowDialogs As Boolean = True)
'--------------------------------------------------------------------
Dim sOutPutDB As String
Dim oMACRODatabase As Database
Dim oMacroDatabaseConnection As ADODB.Connection
Dim sSQL As String
Dim i As Integer
Dim j As Integer
Dim rsResponseData As ADODB.Recordset
Dim sVFQA As String
Dim asVFQA() As String
Dim sDataItemCode As String
Dim sDataItemDescription As String
Dim rsDataItem As ADODB.Recordset
Dim rsValueData As ADODB.Recordset
Dim sDataType As String
Dim sAttribute As String
Dim lSubRecNo As Long
Dim sShortCode As String

    If grsOutPut.Fields.Count > 254 Then
        Call DialogInformation("You have selected more than 255 fields." & vbNewLine & _
            "Saving to Access is disallowed.", "MACRO Query Module")
        Exit Sub
    End If
    
    On Error GoTo CancelSaveAs
    
    Call HourglassOn
    
    'sOutPutDB = gsOUT_FOLDER_LOCATION & TrialNameFromId(glSelectedTrialId) & "_QM.mdb"
    'Mo 2/4/2007 MRC15022007
    sOutPutDB = ConstructOutputFileName & "_QM.mdb"
    
    If bShowDialogs Then
        'Prepare the Access database Save As dialog
        With frmMenu.CommonDialog1
            .DialogTitle = "MACRO Query Save as Access DB"
            .CancelError = True
            .Filter = "Access db (*.mdb)|*.mdb"
            .DefaultExt = "mdb"
            .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
            .FileName = sOutPutDB
            .ShowSave
        
            sOutPutDB = .FileName
        End With
    End If
    
    On Error GoTo Errhandler
    
    'Check for the output file already existing and remove it
    If FileExists(sOutPutDB) Then
        Kill sOutPutDB
        DoEvents
    End If
    
    Set oMACRODatabase = DBEngine.Workspaces(0).CreateDatabase(sOutPutDB, dbLangGeneral, dbEncrypt)
    
    'Create the ResponseData Table
    sSQL = "CREATE Table ResponseData ("
    
    'Loop through grsOutPut to retrieve the required field names for table ResponseData
    For i = 0 To grsOutPut.Fields.Count - 1
        'Mo 30/5/2006 Bug 2668, check for Subject Label being excluded from saved output
        If ((i = 2) And (gbExcludeLabel = True)) Then
            'Don't add the Subject Label column
        Else
            Select Case grsOutPut.Fields(i).Type
            Case adVarChar
                'Changed Mo 5/7/2002, Strings set to their actual size instaed of always being 255
                sSQL = sSQL & "[" & grsOutPut.Fields(i).Name & "] TEXT(" & grsOutPut.Fields(i).DefinedSize & "),"
            Case adInteger
                sSQL = sSQL & "[" & grsOutPut.Fields(i).Name & "] INTEGER,"
            Case adSmallInt
                sSQL = sSQL & "[" & grsOutPut.Fields(i).Name & "] SMALLINT,"
            'Mo 31/1/2007 Bug 2873, change adSingle to adDecimal, change SINGLE to DOUBLE
            Case adDecimal
                sSQL = sSQL & "[" & grsOutPut.Fields(i).Name & "] DOUBLE,"
            Case adDBTimeStamp
                sSQL = sSQL & "[" & grsOutPut.Fields(i).Name & "] DATETIME,"
            End Select
        End If
    Next i
    
    sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY"
    sSQL = sSQL & " (Trial,Site,PersonId,VisitCycle,FormCycle,RepeatNumber))"
    oMACRODatabase.Execute sSQL, dbFailOnError
    
    'Create the  Questions Lookup Table
    If gbUseShortCodes Then
        'Mo 18/10/2006 Bug 2822, increase ShortCode length from 8 to 18 chars
        sSQL = "CREATE Table Questions (ShortCode TEXT(18),"
        sSQL = sSQL & " [Visit/Form/Question] TEXT(255),"
    Else
        sSQL = "CREATE Table Questions ([Visit/Form/Question] TEXT(255),"
    End If
    sSQL = sSQL & " Description TEXT(255) ,"
    sSQL = sSQL & " Type TEXT(15) ,"
    sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY"
    If gbUseShortCodes Then
        sSQL = sSQL & " (ShortCode))"
    Else
        sSQL = sSQL & " ([Visit/Form/Question]))"
    End If
    oMACRODatabase.Execute sSQL, dbFailOnError
    
    'Create the CategoryCodes Lookup Table
    If gbUseShortCodes Then
        'Mo 18/10/2006 Bug 2822, increase ShortCode length from 8 to 18 chars
        sSQL = "CREATE Table CategoryCodes (ShortCode TEXT(18),"
    Else
        sSQL = "CREATE Table CategoryCodes ([Visit/Form/Question] TEXT(255),"
    End If
    sSQL = sSQL & " CatCode TEXT(15) ,"
    sSQL = sSQL & " CatValue TEXT(255) ,"
    sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY"
    If gbUseShortCodes Then
        sSQL = sSQL & " (ShortCode,CatCode))"
    Else
        sSQL = sSQL & " ([Visit/Form/Question],CatCode))"
    End If
    oMACRODatabase.Execute sSQL, dbFailOnError
    
    oMACRODatabase.Close
    Set oMACRODatabase = Nothing

    'create a connection to the new database
    Set oMacroDatabaseConnection = New ADODB.Connection
    'Changed Mo Morris 12/2/2003, switch from Jet 3.51 to Jet 4.0, SR 5184 (see Q197902)
    oMacroDatabaseConnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
      "DATA SOURCE=" & sOutPutDB & ";" & _
      "Jet OLEDB:"
    
    'create an empty recordset that is attached to table ResponseData
    sSQL = "SELECT * FROM ResponseData WHERE true = false"
    Set rsResponseData = New ADODB.Recordset
    rsResponseData.CursorLocation = adUseClient
    rsResponseData.Open sSQL, oMacroDatabaseConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    'Loop through grsOutPut reading the contents and placing it in table ResponseData via rsResponseData
    grsOutPut.MoveFirst
    lSubRecNo = 0
    Do While Not grsOutPut.EOF
        lSubRecNo = lSubRecNo + 1
        rsResponseData.AddNew
        j = 0
        For i = 0 To grsOutPut.Fields.Count - 1
            'Mo 30/5/2006 Bug 2668, check for Subject Label being excluded from saved output
            If ((i = 2) And (gbExcludeLabel = True)) Then
                'Don't add the Subject Label column
            Else
                'only transfer non null response
                If Not IsNull(grsOutPut.Fields(i).Value) Then
                    rsResponseData.Fields(j).Value = grsOutPut.Fields(i).Value
                End If
                j = j + 1
            End If
        Next
        rsResponseData.Update
        Call DisplayProgressMessage("Saving Subject Record " & lSubRecNo)
        grsOutPut.MoveNext
    Loop
    
    Call DisplayProgressMessage("Saving Question details")
    
    rsResponseData.Close
    Set rsResponseData = Nothing
    
    'Populate the questions and CategoryCodes Lookup Tables
    For i = 7 To grsOutPut.Fields.Count - 1
        If gbUseShortCodes Then
            sShortCode = grsOutPut.Fields(i).Name
            sVFQA = gColQuestionCodes(grsOutPut.Fields(i).Name)
        Else
            sVFQA = grsOutPut.Fields(i).Name
        End If
        'sVFQA will be of the format VisitCode/CRFPageCode/DataItemCode or VisitCode/CRFPageCode/DataItemCode/Attribute
        'extract DataItemCode from question column name sVFQA
        asVFQA = Split(sVFQA, "/")
        sDataItemCode = asVFQA(2)
        'attribute question entries are handled differently
        If UBound(asVFQA) = 3 Then
            sAttribute = asVFQA(3)
            'make an entry in table Questions
            If gbUseShortCodes Then
                sSQL = "Insert INTO Questions VALUES ('" & sShortCode & "','" & sVFQA & "','" & sDataItemDescription & "','" & sAttribute & "')"
            Else
                sSQL = "Insert INTO Questions VALUES ('" & sVFQA & "','" & sDataItemDescription & "','" & sAttribute & "')"
            End If
            oMacroDatabaseConnection.Execute sSQL, dbFailOnError
        Else
            'Get Question/DataItem details
            Set rsDataItem = New ADODB.Recordset
            Set rsDataItem = DataItemDetails(glSelectedTrialId, sDataItemCode)
            sDataItemDescription = ReplaceQuotes(rsDataItem!DataItemName)
            'convert DataType number into a string and populate table CategoryCodes for questions of type category
            Select Case rsDataItem!DataType
            Case DataType.Category
                sDataType = "Category"
                Set rsValueData = New ADODB.Recordset
                'Mo 21/8/2006 Bug 2784, call to gdsDataValues replaced gdsDataValuesALL
                Set rsValueData = gdsDataValuesALL(glSelectedTrialId, 1, rsDataItem!DataItemId)
                'Mo 21/8/2006 Bug 2784, check for no category codes
                If rsValueData.RecordCount > 0 Then
                    rsValueData.MoveFirst
                    'Loop throught the category codes and add them to table CategoryCodes
                    Do While Not rsValueData.EOF
                        If gbUseShortCodes Then
                            sSQL = "INSERT INTO CategoryCodes VALUES ('" & sShortCode & "','" & rsValueData!ValueCode & "','" & ReplaceQuotes(rsValueData!ItemValue) & "')"
                        Else
                            sSQL = "INSERT INTO CategoryCodes VALUES ('" & sVFQA & "','" & rsValueData!ValueCode & "','" & ReplaceQuotes(rsValueData!ItemValue) & "')"
                        End If
                        oMacroDatabaseConnection.Execute sSQL, dbFailOnError
                        rsValueData.MoveNext
                    Loop
                End If
            Case DataType.Date
                sDataType = "Date"
            Case DataType.IntegerData
                sDataType = "IntegerData"
            Case DataType.LabTest
                sDataType = "LabTest"
            Case DataType.Multimedia
                sDataType = "Multimedia"
            Case DataType.Real
                sDataType = "Real"
            Case DataType.Text
                sDataType = "Text"
            'Mo 25/10/2005 COD0040
            Case DataType.Thesaurus
                sDataType = "Thesaurus"
            End Select
            'make an entry in table Questions
            If gbUseShortCodes Then
                sSQL = "Insert INTO Questions VALUES ('" & sShortCode & "','" & sVFQA & "','" & sDataItemDescription & "','" & sDataType & "')"
            Else
                sSQL = "Insert INTO Questions VALUES ('" & sVFQA & "','" & sDataItemDescription & "','" & sDataType & "')"
            End If
            oMacroDatabaseConnection.Execute sSQL, dbFailOnError
        End If
    Next i
    rsDataItem.Close
    Set rsDataItem = Nothing
    
    oMacroDatabaseConnection.Close
    Set oMacroDatabaseConnection = Nothing
    
    Call DisplayProgressMessage("Save Output (ACCESS) completed.")
    
    Call HourglassOff
    
Exit Sub
Errhandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "OutputToAccess", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

CancelSaveAs:
    Call HourglassOff

End Sub

'Mo 2/4/2007 MRC15022007, Optional parameter bShowDialogs added
'--------------------------------------------------------------------
Public Sub OutputToCSV(Optional bShowDialogs As Boolean = True)
'--------------------------------------------------------------------
Dim sExportFileCSV As String
Dim sCodeLookUpFile As String
Dim sDataItemLookUpFile As String
Dim nCSVFileNumber As Integer
Dim nDLUFileNumber As Integer
Dim nCLUFileNumber As Integer
Dim sOutPut As String
Dim i As Integer
Dim sVFQA As String
Dim asVFQA() As String
Dim sDataItemCode As String
Dim sDataItemDescription As String
Dim rsDataItem As ADODB.Recordset
Dim rsValueData As ADODB.Recordset
Dim sDataType As String
Dim sAttribute As String
Dim lSubRecNo As Long
Dim sShortCode As String
  
    On Error GoTo CancelSaveAs
    
    Call HourglassOn
    
    'sExportFileCSV = gsOUT_FOLDER_LOCATION & TrialNameFromId(glSelectedTrialId) & "_" & Format(Now, "yyyymmdd") & ".csv"
    'Mo 2/4/2007 MRC15022007
    sExportFileCSV = ConstructOutputFileName & ".csv"
    
    If bShowDialogs Then
        'Prepare and launch the CSV Save As dialog
        With frmMenu.CommonDialog1
            .DialogTitle = "MACRO Query Save as CSV export file"
            .CancelError = True
            .Filter = "CSV (Comma delimited) (*.csv)|*.csv"
            .DefaultExt = "csv"
            .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
            .FileName = sExportFileCSV
            .ShowSave
    
            sExportFileCSV = .FileName
        End With
    End If
    
    
    'Check for the .csv file already existing and remove it
    If FileExists(sExportFileCSV) Then
        Kill sExportFileCSV
        DoEvents
    End If
    
    On Error GoTo Errhandler
    
    'Create a file name for the DataItem Look Up file (dlu) and the Code Look Up file (clu)
    sDataItemLookUpFile = Mid(sExportFileCSV, 1, Len(sExportFileCSV) - 4) & "_DLU.csv"
    sCodeLookUpFile = Mid(sExportFileCSV, 1, Len(sExportFileCSV) - 4) & "_CLU.csv"
    
    'Check for the .dlu file already existing and remove it
    If FileExists(sDataItemLookUpFile) Then
        Kill sDataItemLookUpFile
        DoEvents
    End If
    
    'Check for the .clu file already existing and remove it
    If FileExists(sCodeLookUpFile) Then
        Kill sCodeLookUpFile
        DoEvents
    End If

    'Open the CSV file and create its header record
    nCSVFileNumber = FreeFile
    Open sExportFileCSV For Output As #nCSVFileNumber
    sOutPut = ""
    'Loop through grsOutPut to retrieve the required field names for the CSV file
    For i = 0 To grsOutPut.Fields.Count - 1
        'Mo 30/5/2006 Bug 2668, check for Subject Label being excluded from saved output
        If ((i = 2) And (gbExcludeLabel = True)) Then
            'Don't add the Subject Label column
        Else
            sOutPut = sOutPut & grsOutPut.Fields(i).Name & msCOMMA
        End If
    Next i
    'strip off the last msComma
    sOutPut = Mid$(sOutPut, 1, (Len(sOutPut) - 1))
    Print #nCSVFileNumber, sOutPut
    
    'Open the DataItem LookUp file and create its header record
    nDLUFileNumber = FreeFile
    Open sDataItemLookUpFile For Output As #nDLUFileNumber
    If gbUseShortCodes Then
        sOutPut = "ShortCode" & msCOMMA & "Visit/Form/Question" & msCOMMA & "Description" & msCOMMA & "Type"
    Else
        sOutPut = "Visit/Form/Question" & msCOMMA & "Description" & msCOMMA & "Type"
    End If
    Print #nDLUFileNumber, sOutPut
    
    'Open the Code LookUp file and create its header record
    nCLUFileNumber = FreeFile
    Open sCodeLookUpFile For Output As #nCLUFileNumber
    If gbUseShortCodes Then
        sOutPut = "ShortCode" & msCOMMA & "CatCode" & msCOMMA & "CatValue"
    Else
        sOutPut = "Visit/Form/Question" & msCOMMA & "CatCode" & msCOMMA & "CatValue"
    End If
    Print #nCLUFileNumber, sOutPut
    
    'Loop through grsOutPut reading the contents and writing it out to the CSV response data file
    grsOutPut.MoveFirst
    lSubRecNo = 0
    Do While Not grsOutPut.EOF
        lSubRecNo = lSubRecNo + 1
        sOutPut = ""
        For i = 0 To grsOutPut.Fields.Count - 1
            'Mo 30/5/2006 Bug 2668, check for Subject Label being excluded from saved output
            If ((i = 2) And (gbExcludeLabel = True)) Then
                'Don't add the Subject Label field
            Else
                sOutPut = sOutPut & CSVCommasAndQuotes(grsOutPut.Fields(i).Value) & msCOMMA
            End If
        Next
        'strip off the last msComma
        sOutPut = Mid$(sOutPut, 1, (Len(sOutPut) - 1))
        Print #nCSVFileNumber, sOutPut
        Call DisplayProgressMessage("Saving Subject Record " & lSubRecNo)
        grsOutPut.MoveNext
    Loop
    
    Call DisplayProgressMessage("Saving Question details")
  
    'Populate the DataItem Look Up file (dlu) and the Code Look Up file (clu)
    'Starting from the first question in grsOutPut
    For i = 7 To grsOutPut.Fields.Count - 1
        If gbUseShortCodes Then
            sShortCode = grsOutPut.Fields(i).Name
            sVFQA = gColQuestionCodes(grsOutPut.Fields(i).Name)
        Else
            sVFQA = grsOutPut.Fields(i).Name
        End If
        'sVFQA will be of the format VisitCode/CRFPageCode/DataItemCode or VisitCode/CRFPageCode/DataItemCode/Attribute
        'extract DataItemCode from question column name sVFQA
        asVFQA = Split(sVFQA, "/")
        sDataItemCode = asVFQA(2)
        'attribute question entries are handled differently
        If UBound(asVFQA) = 3 Then
            sAttribute = asVFQA(3)
            'make an entry in the DataItem Look Up file (dlu)
            'note that sDataItemDescription has already passed through the code of CSVCommasAndQuotes
            If gbUseShortCodes Then
                sOutPut = sShortCode & msCOMMA & sVFQA & msCOMMA & sDataItemDescription & msCOMMA & sAttribute
            Else
                sOutPut = sVFQA & msCOMMA & sDataItemDescription & msCOMMA & sAttribute
            End If
            Print #nDLUFileNumber, sOutPut
        Else
            'Get Question/DataItem details
            Set rsDataItem = New ADODB.Recordset
            Set rsDataItem = DataItemDetails(glSelectedTrialId, sDataItemCode)
            sDataItemDescription = rsDataItem!DataItemName
            'convert DataType number into a string and populate the Code Look Up file (clu) for questions of type category
            Select Case rsDataItem!DataType
            Case DataType.Category
                sDataType = "Category"
                Set rsValueData = New ADODB.Recordset
                'Mo 21/8/2006 Bug 2784, call to gdsDataValues replaced gdsDataValuesALL
                Set rsValueData = gdsDataValuesALL(glSelectedTrialId, 1, rsDataItem!DataItemId)
                'Mo 21/8/2006 Bug 2784, check for no category codes
                If rsValueData.RecordCount > 0 Then
                    rsValueData.MoveFirst
                    'Loop throught the category codes and write them to the Code Look Up file (clu)
                    Do While Not rsValueData.EOF
                        If gbUseShortCodes Then
                            sOutPut = sShortCode & msCOMMA & rsValueData!ValueCode & msCOMMA & CSVCommasAndQuotes(rsValueData!ItemValue)
                        Else
                            sOutPut = sVFQA & msCOMMA & rsValueData!ValueCode & msCOMMA & CSVCommasAndQuotes(rsValueData!ItemValue)
                        End If
                        Print #nCLUFileNumber, sOutPut
                        rsValueData.MoveNext
                    Loop
                End If
            Case DataType.Date
                sDataType = "Date"
            Case DataType.IntegerData
                sDataType = "IntegerData"
            Case DataType.LabTest
                sDataType = "LabTest"
            Case DataType.Multimedia
                sDataType = "Multimedia"
            Case DataType.Real
                sDataType = "Real"
            Case DataType.Text
                sDataType = "Text"
            'Mo 25/10/2005 COD0040
            Case DataType.Thesaurus
                sDataType = "Thesaurus"
            End Select
            'make an entry in the DataItem Look Up file (dlu)
            If gbUseShortCodes Then
                sOutPut = sShortCode & msCOMMA & sVFQA & msCOMMA & CSVCommasAndQuotes(sDataItemDescription) & msCOMMA & sDataType
            Else
                sOutPut = sVFQA & msCOMMA & CSVCommasAndQuotes(sDataItemDescription) & msCOMMA & sDataType
            End If
            Print #nDLUFileNumber, sOutPut
        End If
    Next i
    rsDataItem.Close
    Set rsDataItem = Nothing

    'Close the files
    Close #nCSVFileNumber
    Close #nDLUFileNumber
    Close #nCLUFileNumber
    
    Call DisplayProgressMessage("Save Output (CSV) completed.")
    
    Call HourglassOff

Exit Sub
Errhandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "OutputToCSV", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

CancelSaveAs:
    Call HourglassOff

End Sub

'--------------------------------------------------------------------
Private Function CSVCommasAndQuotes(sCSVString) As String
'--------------------------------------------------------------------
'This function is used to prepare string elements that are to be placed in a CSV file.
'
'Strings containing double quotes have the double quotes duplicated.
'Strings containing double quotes have double quotes put around them.
'
'   e.g.    use "MACRO" always      becomes     "use ""MACRO"" always"
'           "this" and "that"       becomes     """this"" and ""that"""
'
'Strings containing a comma have double quotes put around them.
'
'   e.g.    1,500,250               becomes     "1,500,250"
'
'putting both together
'
'   e.g.    "you, what"             becomes     """you, what"""
'           me, "MACRO" and you     becomes     "me, ""MACRO"" and you"
'--------------------------------------------------------------------
'   Mo  2/5/2002    Changed so that double quotes are now put around strings
'                   that contain double quotes.
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    If IsNull(sCSVString) Then
        CSVCommasAndQuotes = ""
        Exit Function
    End If
    
    If (InStr(sCSVString, msCOMMA) < 1) And (InStr(sCSVString, msDQUOTE) < 1) Then
        CSVCommasAndQuotes = sCSVString
        Exit Function
    End If
    
    'Replacing double quotes with 2 double quotes
    sCSVString = Replace(sCSVString, msDQUOTE, msDQUOTE & msDQUOTE)
    
    'Encompass string in double quotes if string contains a comma or double quotes
    'Changed Mo 2/5/2002
    If (InStr(sCSVString, msCOMMA)) > 0 Or (InStr(sCSVString, msDQUOTE) > 0) Then
        sCSVString = msDQUOTE & sCSVString & msDQUOTE
    End If
    
    'return the changed string
    CSVCommasAndQuotes = sCSVString

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CSVCommasAndQuotes", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'Mo 2/4/2007 MRC15022007, Optional parameter bShowDialogs added
'--------------------------------------------------------------------
Public Sub OutputToSAS(Optional bShowDialogs As Boolean = True)
'--------------------------------------------------------------------
'Note that if gbOutputCategoryCodes is false (i.e. output set to Category Values)
'then the SAS-style "Category.txt" file will not be created, because there will be no
'category codes in the "Data.txt" to look up.
'--------------------------------------------------------------------
Dim sExportFileSAS As String
Dim sPrevExportFileSAS As String
Dim sCategoryFile As String
Dim sTypeFile As String
Dim sQLUFile As String
Dim nSASDataFileNumber As Integer
Dim nSASTypeFileNumber As Integer
Dim nSASCategoryFileNumber As Integer
Dim nSASQLUFileNumber As Integer
Dim sOutPut As String
Dim i As Integer
Dim sVFQA As String
Dim sShortCode As String
Dim asVFQA() As String
Dim sDataItemCode As String
Dim sDataItemDescription As String
Dim rsDataItem As ADODB.Recordset
Dim sAttribute As String
Dim q As Integer
Dim sDecimalEnding As String
Dim nTrialNameLength As Integer
Dim nCatCodeLength As Integer
Dim bCatCodesNumeric As Boolean
Dim lSubRecNo As Long
Dim sLocalDot As String
'Mo 24/5/2007 Bug 2913
Dim sDateFormatFile As String
Dim nSASDateFormatFileNumber As Integer
  
    On Error GoTo CancelSaveAs
    
    Call HourglassOn
    
    'Mo 2/11/2006 Bug 2797, no need for gColQuestionCodes if Long Codes are specified
    'Don't create a Question Codes Collection if 8 character codes have been generated for Display purposes
    'If Not gbUseShortCodes Then
    '    Set gColQuestionCodes = New Collection
    'End If
    
    'sExportFileSAS = gsOUT_FOLDER_LOCATION & TrialNameFromId(glSelectedTrialId) & "_" & Format(Now, "yyyymmdd") & "_SAS_Data.txt"
    'Mo 2/4/2007 MRC15022007
    sExportFileSAS = ConstructOutputFileName & "_SAS_Data.txt"
    sPrevExportFileSAS = sExportFileSAS
    
    If bShowDialogs Then
        'Prepare and launch the SAS Save As dialog
        With frmMenu.CommonDialog1
            .DialogTitle = "MACRO Query Save as SAS export file"
            .CancelError = True
            .Filter = "Text only (*.txt)|*.txt"
            .DefaultExt = "txt"
            .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
            .FileName = sExportFileSAS
            .ShowSave
  
            sExportFileSAS = .FileName
        End With
    End If
    
    'Check for the SAS.txt file already existing and remove it
    If FileExists(sExportFileSAS) Then
        Kill sExportFileSAS
        DoEvents
    End If
    
    On Error GoTo Errhandler
    
    If sExportFileSAS = sPrevExportFileSAS Then
        'Create a file name for the DataItem Look Up file and the Code Look Up file
        'based on the data file being the standard name TrialName_YYYYMMDD_SAS_Data.txt
        'The DataItemLookupFile will be called TrialName_YYYYMMDD_SAS_Type.txt
        'The CodeLookUpfile will be called TrialName_YYYYMMDD_SAS_Category.txt
        'The Question Look Up file will be called TrialName_YYYYMMDD_SAS_QLU.txt
        'The Date Formats file will be called TrialName_YYYYMMDD_SAS_Format.txt
        sTypeFile = Mid(sExportFileSAS, 1, Len(sExportFileSAS) - 8) & "Type.txt"
        sCategoryFile = Mid(sExportFileSAS, 1, Len(sExportFileSAS) - 8) & "Category.txt"
        sQLUFile = Mid(sExportFileSAS, 1, Len(sExportFileSAS) - 8) & "QLU.txt"
        'Mo 24/5/2007 Bug 2913
        sDateFormatFile = Mid(sExportFileSAS, 1, Len(sExportFileSAS) - 8) & "Format.txt"
    Else
        'Create file names based on the fact that the user has changed the standard name
        'The Data table will be called UserEnteredName_Data.txt
        'The DataItemLookupFile will be called UserEnteredName_Type.txt
        'The CodeLookUpfile will be called UserEnteredName_Category.txt
        'The Question Look Up file will be called UserEnteredName_QLU.txt
        'The Date Formats file will be called UserEnteredName_Format.txt
        'If the user entered name ends with "_Data" then strip it out
        'This will occure when a user changes the default folder, but keeps the proposed file name
        If Mid(sExportFileSAS, (Len(sExportFileSAS) - 8)) = "_Data.txt" Then
            sExportFileSAS = Mid(sExportFileSAS, 1, Len(sExportFileSAS) - 9) & ".txt"
        End If
        sExportFileSAS = Mid(sExportFileSAS, 1, Len(sExportFileSAS) - 4) & "_Data.txt"
        sTypeFile = Mid(sExportFileSAS, 1, Len(sExportFileSAS) - 8) & "Type.txt"
        sCategoryFile = Mid(sExportFileSAS, 1, Len(sExportFileSAS) - 8) & "Category.txt"
        sQLUFile = Mid(sExportFileSAS, 1, Len(sExportFileSAS) - 8) & "QLU.txt"
        'Mo 24/5/2007 Bug 2913
        sDateFormatFile = Mid(sExportFileSAS, 1, Len(sExportFileSAS) - 8) & "Format.txt"
    End If
    
    'Check for the _Type file already existing and remove it
    If FileExists(sTypeFile) Then
        Kill sTypeFile
        DoEvents
    End If
    
    'Check for the _Category file already existing and remove it
    'Perform this even if gbOutputCategoryCodes is false and a Category file is not being created
    If FileExists(sCategoryFile) Then
        Kill sCategoryFile
        DoEvents
    End If
    
    'Check for the _QLU file already existing and remove it
    If FileExists(sQLUFile) Then
        Kill sQLUFile
        DoEvents
    End If
    
    'Mo 24/5/2007 Bug 2913
    'Check for the Date Format file already existing and remove it
    If FileExists(sDateFormatFile) Then
        Kill sDateFormatFile
        DoEvents
    End If
    
    'Open the SAS_Data file
    nSASDataFileNumber = FreeFile
    Open sExportFileSAS For Output As #nSASDataFileNumber
    
    'Open the SAS_Type file
    nSASTypeFileNumber = FreeFile
    Open sTypeFile For Output As #nSASTypeFileNumber
    
    'Open the SAS_Category file
    If gbOutputCategoryCodes Then
        nSASCategoryFileNumber = FreeFile
        Open sCategoryFile For Output As #nSASCategoryFileNumber
    End If
    
    'Open the SAS_QLU file and create its header record
    nSASQLUFileNumber = FreeFile
    Open sQLUFile For Output As #nSASQLUFileNumber
    'Mo 26/5/2006 Bug 2738
    sOutPut = "Visit/Form/Question" & msTAB & "SASCode" & msTAB & "Description" & msTAB & "Type"
    Print #nSASQLUFileNumber, sOutPut
    
    'Mo 24/5/2007 Bug 2913
    'Open the SAS_Format file
    nSASDateFormatFileNumber = FreeFile
    Open sDateFormatFile For Output As #nSASDateFormatFileNumber

    'get the Regional Decimal Point Character
    sLocalDot = RegionalDecimalPointChar
    
    'Loop through grsOutPut reading the contents and writing it out to the SAS Data file as a TAB delimited record
    grsOutPut.MoveFirst
    nTrialNameLength = 0
    lSubRecNo = 0
    Do While Not grsOutPut.EOF
        lSubRecNo = lSubRecNo + 1
        sOutPut = ""
        For i = 0 To grsOutPut.Fields.Count - 1
            'Mo 30/5/2006 Bug 2668, check for Subject Label being excluded from saved output
            If ((i = 2) And (gbExcludeLabel = True)) Then
                'Don't add the Subject Label field
            Else
                'asses the length of the TrialName field which is always the first field.
                If i = 0 Then
                    If nTrialNameLength = 0 Then
                        nTrialNameLength = Len(grsOutPut.Fields(i).Value)
                    End If
                End If
                sDecimalEnding = ""
                'test for a real/single and check that it contains a decimal point
                'Mo 31/1/2007 Bug 2873
                If grsOutPut.Fields(i).Type = adDecimal Then
                    If grsOutPut.Fields(i).NumericScale > 0 Then
                        'The number of decimal places have been stored in the NumericScale property
                        If InStr(grsOutPut.Fields(i).Value, sLocalDot) = 0 Then
                            'the response has had its decimal point followed by a zero removed. Add it back
                            'Don't do this for a specialvalue
                            If grsOutPut.Fields(i).Value <> -1 And grsOutPut.Fields(i).Value <> -2 And grsOutPut.Fields(i).Value <> -3 _
                                And grsOutPut.Fields(i).Value <> -4 And grsOutPut.Fields(i).Value <> -5 And grsOutPut.Fields(i).Value <> -6 _
                                And grsOutPut.Fields(i).Value <> -7 And grsOutPut.Fields(i).Value <> -8 And grsOutPut.Fields(i).Value <> -9 Then
                                sDecimalEnding = sLocalDot & "0"
                            End If
                        End If
                    End If
                End If
                'test for a Date field
                If grsOutPut.Fields(i).Type = adDBTimeStamp Then
                    'Date/Time responses have to be written to SAS in ddMMMyyyy:hh:mm:ss
                    sOutPut = sOutPut & SASDateFormat(grsOutPut.Fields(i).Value) & msTAB
                ElseIf Len(grsOutPut.Fields(i).Value) > 200 Then
                    'Changed Mo Morris 18/7/2002, SAS max string length 200
                    'Note that sDecimalEnding will always be empty for long string fields and is not included in the following line
                    sOutPut = sOutPut & Left(grsOutPut.Fields(i).Value, 200) & msTAB
                Else
                    sOutPut = sOutPut & grsOutPut.Fields(i).Value & sDecimalEnding & msTAB
                End If
            End If
        Next
        'strip off the last msTab
        sOutPut = Mid$(sOutPut, 1, (Len(sOutPut) - 1))
        Print #nSASDataFileNumber, sOutPut
        Call DisplayProgressMessage("Saving Subject Record " & lSubRecNo)
        grsOutPut.MoveNext
    Loop
    
    Call DisplayProgressMessage("Saving Question details")

    'Write the subject identification fields to the SAS_Type file
    'Mo 1/11/2006 Bug 2795
    If gbSASInformatColons Then
        Print #nSASTypeFileNumber, "Trial :$" & nTrialNameLength & "."
        Print #nSASTypeFileNumber, "Site :$8."
        'Mo 30/5/2006 Bug 2668, check for Subject Label being excluded from saved output
        If (gbExcludeLabel = False) Then
            Print #nSASTypeFileNumber, "Label :$50."
        End If
        Print #nSASTypeFileNumber, "PersonId :5."
        Print #nSASTypeFileNumber, "VisCycle :5."
        Print #nSASTypeFileNumber, "FrmCycle :5."
        Print #nSASTypeFileNumber, "RepeatNo :5."
    Else
        Print #nSASTypeFileNumber, "Trial $" & nTrialNameLength & "."
        Print #nSASTypeFileNumber, "Site $8."
        'Mo 30/5/2006 Bug 2668, check for Subject Label being excluded from saved output
        If (gbExcludeLabel = False) Then
            Print #nSASTypeFileNumber, "Label $50."
        End If
        Print #nSASTypeFileNumber, "PersonId 5."
        Print #nSASTypeFileNumber, "VisCycle 5."
        Print #nSASTypeFileNumber, "FrmCycle 5."
        Print #nSASTypeFileNumber, "RepeatNo 5."
    End If
    
    'Populate the SAS_Type file and the SAS_Category file together with the SAS_QLU file
    'Starting from the first question in grsOutPut
    For i = 7 To grsOutPut.Fields.Count - 1
        If gbUseShortCodes Then
            sVFQA = gColQuestionCodes(grsOutPut.Fields(i).Name)
        Else
            sVFQA = grsOutPut.Fields(i).Name
        End If
        'sVFQA will be of the format VisitCode/CRFPageCode/DataItemCode or VisitCode/CRFPageCode/DataItemCode/Attribute
        'extract DataItemCode from question column name sVFQA
        asVFQA = Split(sVFQA, "/")
        sDataItemCode = asVFQA(2)
        If gbUseShortCodes Then
            sShortCode = grsOutPut.Fields(i).Name
        Else
            'Mo 2/11/2006 Bug 2797, Long codes specified, create a VisitCode_FormCode_QuestionCode code
            'create a unique 8 character (max) question code from the DataItemCode
            'sShortCode = CreateUniqueQuestion(sDataItemCode, sDataItemCode)
            sShortCode = Replace(sVFQA, "/", "_")
        End If
        'test for a question or a question attribute
        If UBound(asVFQA) = 3 Then
            'its an attribute
            sAttribute = asVFQA(3)
            Select Case sAttribute
            Case "Comments"
                'Changed Mo Morris 18/7/2002, SAS max string length 200
                'Mo 1/11/2006 Bug 2795
                If gbSASInformatColons Then
                    Print #nSASTypeFileNumber, sShortCode & " :$200."
                Else
                    Print #nSASTypeFileNumber, sShortCode & " $200."
                End If
            Case "CTCGrade", "LabResult"
                'Mo 1/11/2006 Bug 2795
                If gbSASInformatColons Then
                    Print #nSASTypeFileNumber, sShortCode & " :$1."
                Else
                    Print #nSASTypeFileNumber, sShortCode & " $1."
                End If
            Case "Status"
                'Mo 1/11/2006 Bug 2795
                If gbSASInformatColons Then
                    Print #nSASTypeFileNumber, sShortCode & " :$15."
                Else
                    Print #nSASTypeFileNumber, sShortCode & " $15."
                End If
            Case "TimeStamp"
                'Mo 1/11/2006 Bug 2795
                If gbSASInformatColons Then
                    Print #nSASTypeFileNumber, sShortCode & " :datetime19."
                Else
                    Print #nSASTypeFileNumber, sShortCode & " datetime19."
                End If
                'Mo 24/5/2007 Bug 2913
                Print #nSASDateFormatFileNumber, sShortCode & " datetime19."
            Case "UserName"
                'Mo 1/11/2006 Bug 2795
                If gbSASInformatColons Then
                    Print #nSASTypeFileNumber, sShortCode & " :$20."
                Else
                    Print #nSASTypeFileNumber, sShortCode & " $20."
                End If
            Case "ValueCode"
                'Mo 1/11/2006 Bug 2795
                If gbSASInformatColons Then
                    Print #nSASTypeFileNumber, sShortCode & " :$15."
                Else
                    Print #nSASTypeFileNumber, sShortCode & " $15."
                End If
            End Select
            'Mo 26/5/2006 Bug 2738, add Name and Type to QLU file output
            Print #nSASQLUFileNumber, sVFQA & msTAB & sShortCode & msTAB & rsDataItem!DataItemName & msTAB & sAttribute
        Else
            'Get Question/DataItem details
            Set rsDataItem = New ADODB.Recordset
            Set rsDataItem = DataItemDetails(glSelectedTrialId, sDataItemCode)
            sDataItemDescription = rsDataItem!DataItemName
            'Based on the Question type Write to the "Type.txt" file
            Select Case rsDataItem!DataType
            'Mo 25/10/2005 COD0040
            Case DataType.Text, DataType.Thesaurus
                'Changed Mo Morris 18/7/2002, SAS max string length 200
                If rsDataItem!DataItemLength > 200 Then
                    'Mo 1/11/2006 Bug 2795
                    If gbSASInformatColons Then
                        Print #nSASTypeFileNumber, sShortCode & " :$200."
                    Else
                        Print #nSASTypeFileNumber, sShortCode & " $200."
                    End If
                ElseIf rsDataItem!DataItemLength = 1 Then
                    'Force single character fields to be two charater fields that are
                    'capable of holding a special value string of "-1" to "-9"
                    'Mo 1/11/2006 Bug 2795
                    If gbSASInformatColons Then
                        Print #nSASTypeFileNumber, sShortCode & " :$2."
                    Else
                        Print #nSASTypeFileNumber, sShortCode & " $2."
                    End If
                Else
                    'Mo 1/11/2006 Bug 2795
                    If gbSASInformatColons Then
                        Print #nSASTypeFileNumber, sShortCode & " :$" & rsDataItem!DataItemLength & "."
                    Else
                        Print #nSASTypeFileNumber, sShortCode & " $" & rsDataItem!DataItemLength & "."
                    End If
                End If
            Case DataType.Date
                'Mo 18/10/2006 Bug 2822, check Partial Dates flag before deciding on date format
                If CInt(RemoveNull(rsDataItem!DataItemCase)) = 0 Then
                    'Call SASDateFormatType to decide on the required date format
                    'Mo 1/11/2006 Bug 2795
                    If gbSASInformatColons Then
                        Print #nSASTypeFileNumber, sShortCode & " :" & SASDateFormatType(rsDataItem!DataItemFormat)
                    Else
                        Print #nSASTypeFileNumber, sShortCode & " " & SASDateFormatType(rsDataItem!DataItemFormat)
                    End If
                    'Mo 24/5/2007 Bug 2913
                    Print #nSASDateFormatFileNumber, sShortCode & " " & SASDateFormatType(rsDataItem!DataItemFormat)
                Else
                    'Setup as a partial date question
                    'Mo 1/11/2006 Bug 2795
                    If gbSASInformatColons Then
                        Print #nSASTypeFileNumber, sShortCode & " :$20."
                    Else
                        Print #nSASTypeFileNumber, sShortCode & " $20."
                    End If
                End If
            Case DataType.Category
                'if gbOutputCategoryCodes = False (i.e. category Values are being output instead of
                'category Codes then all Category Questions will be treated as Text questions
                If gbOutputCategoryCodes Then
                    'its a Category Questions containing category codes assess the usage of Numeric or Alpha codes
                    Call AssessCategoryCodes(glSelectedTrialId, 1, rsDataItem!DataItemId, bCatCodesNumeric, nCatCodeLength)
                    'Force single character fields to be two character fields that are
                    'capable of holding a special value string of "-1" to "-9"
                    If nCatCodeLength = 1 Then
                        nCatCodeLength = 2
                    End If
                    If bCatCodesNumeric Then
                        'Mo 1/11/2006 Bug 2795
                        If gbSASInformatColons Then
                            Print #nSASTypeFileNumber, sShortCode & " :" & nCatCodeLength & "."
                        Else
                            Print #nSASTypeFileNumber, sShortCode & " " & nCatCodeLength & "."
                        End If
                    Else
                        'Mo 1/11/2006 Bug 2795
                        If gbSASInformatColons Then
                            Print #nSASTypeFileNumber, sShortCode & " :$" & nCatCodeLength & "."
                        Else
                            Print #nSASTypeFileNumber, sShortCode & " $" & nCatCodeLength & "."
                        End If
                    End If
                    'write the category codes to the SAS Category.txt file
                    Call WriteCatCodesToSAS(glSelectedTrialId, 1, rsDataItem!DataItemId, sShortCode, nSASCategoryFileNumber, bCatCodesNumeric)
                Else
                    'Treat Category Question with category values as a Text Question of length rsDataItem!DataItemLength
                    'Mo 1/8/2006 Bug 2775, Force single character fields to be two character fields
                    'that are capable of holding a special value string of "-1" to "-9"
                    If rsDataItem!DataItemLength = 1 Then
                        'Mo 1/11/2006 Bug 2795
                        If gbSASInformatColons Then
                            Print #nSASTypeFileNumber, sShortCode & " :$2."
                        Else
                            Print #nSASTypeFileNumber, sShortCode & " $2."
                        End If
                    Else
                        'Mo 1/11/2006 Bug 2795
                        If gbSASInformatColons Then
                            Print #nSASTypeFileNumber, sShortCode & " :$" & rsDataItem!DataItemLength & "."
                        Else
                            Print #nSASTypeFileNumber, sShortCode & " $" & rsDataItem!DataItemLength & "."
                        End If
                    End If
                End If
            Case DataType.Multimedia
                'Multimedia fields are always 36 char long
                'Mo 1/11/2006 Bug 2795
                If gbSASInformatColons Then
                    Print #nSASTypeFileNumber, sShortCode & " :$36."
                Else
                    Print #nSASTypeFileNumber, sShortCode & " $36."
                End If
            Case DataType.IntegerData
                'Mo 1/8/2006 Bug 2775, Force single character fields to be two character fields
                'that are capable of holding a special value string of "-1" to "-9"
                If rsDataItem!DataItemLength = 1 Then
                    'Mo 1/11/2006 Bug 2795
                    If gbSASInformatColons Then
                        Print #nSASTypeFileNumber, sShortCode & " :2."
                    Else
                        Print #nSASTypeFileNumber, sShortCode & " 2."
                    End If
                Else
                    'Mo 1/11/2006 Bug 2795
                    If gbSASInformatColons Then
                        Print #nSASTypeFileNumber, sShortCode & " :" & rsDataItem!DataItemLength & "."
                    Else
                        Print #nSASTypeFileNumber, sShortCode & " " & rsDataItem!DataItemLength & "."
                    End If
                End If
            Case DataType.Real, DataType.LabTest
                'check for a real/labtest that does not have any decimal places
                If InStr(rsDataItem!DataItemFormat, ".") = 0 Then
                    'Mo 1/11/2006 Bug 2795
                    If gbSASInformatColons Then
                        Print #nSASTypeFileNumber, sShortCode & " :" & rsDataItem!DataItemLength & ".0"
                    Else
                        Print #nSASTypeFileNumber, sShortCode & " " & rsDataItem!DataItemLength & ".0"
                    End If
                Else
                    'Mo 1/11/2006 Bug 2795
                    If gbSASInformatColons Then
                        Print #nSASTypeFileNumber, sShortCode & " :" & rsDataItem!DataItemLength & "." & rsDataItem!DataItemLength - InStr(rsDataItem!DataItemFormat, ".")
                    Else
                        Print #nSASTypeFileNumber, sShortCode & " " & rsDataItem!DataItemLength & "." & rsDataItem!DataItemLength - InStr(rsDataItem!DataItemFormat, ".")
                    End If
                End If
            End Select
            'Write the Long form and the Short form of the current question to the QLU file
            'Mo 26/5/2006 Bug 2738, add Name and Type to QLU file output
            Print #nSASQLUFileNumber, sVFQA & msTAB & sShortCode & msTAB & rsDataItem!DataItemName & msTAB & GetDataTypeString(rsDataItem!DataType)
        End If
    Next i
    rsDataItem.Close
    Set rsDataItem = Nothing

    'Close the files
    Close #nSASDataFileNumber
    Close #nSASTypeFileNumber
    If gbOutputCategoryCodes Then
        Close #nSASCategoryFileNumber
    End If
    Close #nSASQLUFileNumber
    'Mo 24/5/2007 Bug 2913
    Close #nSASDateFormatFileNumber
    
    'Mo 2/11/2006 Bug 2797, no need for gColQuestionCodes if Long Codes are specified
    'Don't remove the Question Codes Collection if 8 character codes have been generated for Display purposes
    'If Not gbUseShortCodes Then
    '    Set gColQuestionCodes = Nothing
    'End If
    
    Call DisplayProgressMessage("Save Output (SAS) completed.")
    
    Call HourglassOff

Exit Sub
Errhandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "OutputToSAS", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

CancelSaveAs:
    Call HourglassOff

End Sub

'---------------------------------------------------------------------
Public Function CreateUniqueQuestion(ByVal sDataItemCode As String, _
                                        ByVal sLongCode As String) As String
'---------------------------------------------------------------------
'Mo 2/6/2006 Bug 2737, This function used to create 8 Character long codes, It now
'creates codes based on the length of gnShortCodeLength.
'It checks the DataItemCode length for not being greater than gnShortCodeLength characters.
'It truncates lengths greater than gnShortCodeLength characters to gnShortCodeLength-3 and
'adds a numeric suffix until a unique Question Code is created.
'The Global collection gColQuestionCodes is used to check for uniqueness.
'---------------------------------------------------------------------
Dim sQuestionCode As String
Dim nSuffix As Integer
Dim sSuffixedName As String
Dim bNameAddedToCollection As Boolean

    On Error GoTo Errhandler
    
    sQuestionCode = sDataItemCode
    
    If Len(sQuestionCode) > gnShortCodeLength Then
        sQuestionCode = Mid(sQuestionCode, 1, gnShortCodeLength - 3)
        bNameAddedToCollection = False
        nSuffix = 1
        Do
            sSuffixedName = sQuestionCode & CStr(nSuffix)
            On Error Resume Next
            gColQuestionCodes.Add sLongCode, sSuffixedName
            'If the QuestionCode is unique it will add to the collection.
            'If its not unique an error will have been raised
            If Err.Number = 0 Then
                bNameAddedToCollection = True
            Else
                nSuffix = nSuffix + 1
                Err.Clear
            End If
            On Error GoTo Errhandler
        Loop Until bNameAddedToCollection
        sQuestionCode = sSuffixedName
    Else
        'The length of sQuestionCode will be gnShortCodeLength or less characters.
        'Try to add sQuestionCode and if it fails reduce its length and have a loop like above
        bNameAddedToCollection = False
        On Error Resume Next
        gColQuestionCodes.Add sLongCode, sQuestionCode
        'If the QuestionCode is unique it will add to the collection.
        'If its not unique an error will have been raised
        If Err.Number <> 0 Then
            Err.Clear
            sQuestionCode = Mid(sQuestionCode, 1, gnShortCodeLength - 3)
            bNameAddedToCollection = False
            nSuffix = 1
            Do
                sSuffixedName = sQuestionCode & CStr(nSuffix)
                On Error Resume Next
                gColQuestionCodes.Add sLongCode, sSuffixedName
                'If the QuestionCode is unique it will add to the collection.
                'If its not unique an error will have been raised
                If Err.Number = 0 Then
                    bNameAddedToCollection = True
                Else
                    nSuffix = nSuffix + 1
                    Err.Clear
                End If
                On Error GoTo Errhandler
            Loop Until bNameAddedToCollection
            sQuestionCode = sSuffixedName
        End If
        On Error GoTo Errhandler
    End If
    
    CreateUniqueQuestion = sQuestionCode

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CreateUniqueQuestion", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'---------------------------------------------------------------------
Public Sub DisplayProgressMessage(ByVal sMessage As String)
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'place message in Progress textbox
    frmMenu.txtProgress.Text = sMessage
    DoEvents    'to allow txtProgress to get updated

Exit Sub
Errhandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DisplayProgressMessage", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'---------------------------------------------------------------------
Public Function DateFormatCanBeConverted(ByRef sDateFormat As String) As Boolean
'---------------------------------------------------------------------
'This function is used to decide which date/time questions can be converted
'from string fields to DateTime fields.
'
'This function takes a date/time format string and standardizes its format using
'several Replace statements.
'There are 20 standard formats.
'11 of the standard formats will be converted into DateTime fields:-
'   d/m/y       d/m/y/h/m       d/m/y/h/m/s
'   m/d/y       m/d/y/h/m       m/d/y/h/m/s
'   y/m/d       y/m/d/h/m       y/m/d/h/m/s
'   h/m
'   h/m/s
'9 of the formats will not be converted into DateTime fields, they will remain as strings
'   y/m     y/m/h/m     y/m/h/m/s
'   m/y     m/y/h/m     m/y/h/m/s
'   y/d/m   y/d/m/h/m   y/d/m/h/m/s
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'Replace double format characters with single format characters
    sDateFormat = Replace(sDateFormat, "dd", "d")
    sDateFormat = Replace(sDateFormat, "mm", "m")
    sDateFormat = Replace(sDateFormat, "hh", "h")
    sDateFormat = Replace(sDateFormat, "ss", "s")
    'Replace yyyy with y
    sDateFormat = Replace(sDateFormat, "yyyy", "y")
    'Replace all Date/Time Separators with "/"
    sDateFormat = Replace(sDateFormat, ":", "/")
    sDateFormat = Replace(sDateFormat, ".", "/")
    sDateFormat = Replace(sDateFormat, "-", "/")
    sDateFormat = Replace(sDateFormat, " ", "/")
    
    Select Case sDateFormat
    Case "d/m/y", "m/d/y", "y/m/d", "h/m", "h/m/s", "d/m/y/h/m", "m/d/y/h/m", "y/m/d/h/m", "d/m/y/h/m/s", "m/d/y/h/m/s", "y/m/d/h/m/s"
        DateFormatCanBeConverted = True
    Case Else
        'date formats y/m, m/y and y/d/m (with or without time elements) are not converted
        DateFormatCanBeConverted = False
    End Select

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DateFormatCanBeConverted", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Function yyyymmddhhmmssDateFormat(ByVal sDateFormat As String, ByVal sResponse As String) As String
'---------------------------------------------------------------------
'This function is passed a standardized date/time format string together
'with a corresponding response.
'The contents of sResponse will have all of its Separators Replaced by "/"
'This function will then extract the d/m/y/h/m/s elements and then construct a
'yyyymmddhhmmss Date Format string that is used to write the response
'into a recordsets DateTime field.
'---------------------------------------------------------------------
Dim sDay As String
Dim sMonth As String
Dim sYear As String
Dim sHour As String
Dim sMin As String
Dim sSec As String
Dim asElements() As String
Dim sUDF As String

    On Error GoTo Errhandler
    
    'Replace all Separators with "/"
    sResponse = Replace(sResponse, ":", "/")
    sResponse = Replace(sResponse, ".", "/")
    sResponse = Replace(sResponse, "-", "/")
    sResponse = Replace(sResponse, " ", "/")
    
    Select Case sDateFormat
    Case "d/m/y"
        asElements = Split(sResponse, "/")
        sDay = Format(asElements(0), "00")
        sMonth = Format(asElements(1), "00")
        sYear = asElements(2)
        sUDF = sYear & "/" & sMonth & "/" & sDay
    Case "m/d/y"
        asElements = Split(sResponse, "/")
        sMonth = Format(asElements(0), "00")
        sDay = Format(asElements(1), "00")
        sYear = asElements(2)
        sUDF = sYear & "/" & sMonth & "/" & sDay
    Case "y/m/d"
        asElements = Split(sResponse, "/")
        sYear = asElements(0)
        sMonth = Format(asElements(1), "00")
        sDay = Format(asElements(2), "00")
        sUDF = sYear & "/" & sMonth & "/" & sDay
    Case "h/m"
        asElements = Split(sResponse, "/")
        sHour = Format(asElements(0), "00")
        sMin = Format(asElements(1), "00")
        sUDF = sHour & ":" & sMin & ":00"
    Case "h/m/s"
        asElements = Split(sResponse, "/")
        sHour = Format(asElements(0), "00")
        sMin = Format(asElements(1), "00")
        sSec = Format(asElements(2), "00")
        sUDF = sHour & ":" & sMin & ":" & sSec
    Case "d/m/y/h/m"
        asElements() = Split(sResponse, "/")
        sDay = Format(asElements(0), "00")
        sMonth = Format(asElements(1), "00")
        sYear = asElements(2)
        sHour = Format(asElements(3), "00")
        sMin = Format(asElements(4), "00")
        sUDF = sYear & "/" & sMonth & "/" & sDay & " " & sHour & ":" & sMin & ":00"
    Case "m/d/y/h/m"
        asElements = Split(sResponse, "/")
        sMonth = Format(asElements(0), "00")
        sDay = Format(asElements(1), "00")
        sYear = asElements(2)
        sHour = Format(asElements(3), "00")
        sMin = Format(asElements(4), "00")
        sUDF = sYear & "/" & sMonth & "/" & sDay & " " & sHour & ":" & sMin & ":00"
    Case "y/m/d/h/m"
        asElements = Split(sResponse, "/")
        sYear = asElements(0)
        sMonth = Format(asElements(1), "00")
        sDay = Format(asElements(2), "00")
        sHour = Format(asElements(3), "00")
        sMin = Format(asElements(4), "00")
        sUDF = sYear & "/" & sMonth & "/" & sDay & " " & sHour & ":" & sMin & ":00"
    Case "d/m/y/h/m/s"
        asElements = Split(sResponse, "/")
        sDay = Format(asElements(0), "00")
        sMonth = Format(asElements(1), "00")
        sYear = asElements(2)
        sHour = Format(asElements(3), "00")
        sMin = Format(asElements(4), "00")
        sSec = Format(asElements(5), "00")
        sUDF = sYear & "/" & sMonth & "/" & sDay & " " & sHour & ":" & sMin & ":" & sSec
    Case "m/d/y/h/m/s"
        asElements = Split(sResponse, "/")
        sMonth = Format(asElements(0), "00")
        sDay = Format(asElements(1), "00")
        sYear = asElements(2)
        sHour = Format(asElements(3), "00")
        sMin = Format(asElements(4), "00")
        sSec = Format(asElements(5), "00")
        sUDF = sYear & "/" & sMonth & "/" & sDay & " " & sHour & ":" & sMin & ":" & sSec
    Case "y/m/d/h/m/s"
        asElements = Split(sResponse, "/")
        sYear = asElements(0)
        sMonth = Format(asElements(1), "00")
        sDay = Format(asElements(2), "00")
        sHour = Format(asElements(3), "00")
        sMin = Format(asElements(4), "00")
        sSec = Format(asElements(5), "00")
        sUDF = sYear & "/" & sMonth & "/" & sDay & " " & sHour & ":" & sMin & ":" & sSec
    End Select
    
    yyyymmddhhmmssDateFormat = sUDF

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "yyyymmddhhmmssDateFormat", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'Mo 2/4/2007 MRC15022007, Optional parameter bShowDialogs added
'---------------------------------------------------------------------
Public Sub OutputToSTATA(ByVal sDateType As String, Optional bShowDialogs As Boolean = True)
'---------------------------------------------------------------------
'This will produce the following types of STATA output file:-
'   .ana (STATA data file)
'   .dct (STATA dictionary file)
'   .do  (STATA Program file)
'as well as a Question Look Up File called
'   Name_QLU.txt
'
'Note that if gbOutputCategoryCodes is false (i.e. output set to Category Values)
'then the category questions will be output as strings and there will be no "Label define'
'declarations in the .do filecategory code and no UnderScoredName reverences to them in the dct file.
'
'Note that sDateType denotes the format for exporting dates.
'   "Standard"  Uses ddmmmyyyy Standard dates (e.g. 01jan2004)
'   "Float"     Uses ddmmyyyy Float dates (e.g. 01012004 for 1 January 2004)
'---------------------------------------------------------------------
Dim sSTATAanaFile As String
Dim sSTATAdctFile As String
Dim sSTATAdoFile As String
Dim sSTATAqluFile As String
Dim nSTATAanaFileNumber As Integer
Dim nSTATAdctFileNumber As Integer
Dim nSTATAdoFileNumber As Integer
Dim nSTATAqluFileNumber As Integer
Dim sFileNameana As String
Dim sFileNamedct As String
'Mo 9/10/2007 Bug 2941, nColumnNo changed from Integer to Long (lColumnNo)
Dim lColumnNo As Long
Dim i As Integer
Dim sVFQA As String
Dim sShortCode As String
Dim asVFQA() As String
Dim sDataItemCode As String
Dim rsDataItem As ADODB.Recordset
Dim sAttribute As String
Dim sDataItemDescription As String
Dim nCatCodeLength As Integer
Dim bCatCodesNumeric As Boolean
Dim q As Integer
Dim sOutPut As String
Dim lSubRecNo As Long
Dim sLocalDot As String
Dim sDecimalEnding As String
Dim sCatCodeName As String
Dim sSTATADateTime As String
Dim nQuestionsCounter As Integer

    On Error GoTo CancelSaveAs
    
    Call HourglassOn
    
    'Don't create a Question Codes Collection if 8 character codes have been generated for Display purposes
    If Not gbUseShortCodes Then
        Set gColQuestionCodes = New Collection
    End If
    
    Set gColSTATADetails = New Collection
    
    'sSTATAanaFile = gsOUT_FOLDER_LOCATION & TrialNameFromId(glSelectedTrialId) & "_" & Format(Now, "yyyymmdd") & "STATA.ana"
    'Mo 2/4/2007 MRC15022007
    sSTATAanaFile = ConstructOutputFileName & "_STATA.ana"
    
    If bShowDialogs Then
        'Prepare and launch the STATA Save As dialog
        With frmMenu.CommonDialog1
            .DialogTitle = "MACRO Query Save as STATA export file"
            .CancelError = True
            .Filter = "STATA data file (*.ana)|*.ana"
            .DefaultExt = "txt"
            .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
            .FileName = sSTATAanaFile
            .ShowSave
  
            sSTATAanaFile = .FileName
        End With
    End If
    
    'Check for the STATA.ana file already existing and remove it
    If FileExists(sSTATAanaFile) Then
        Kill sSTATAanaFile
        DoEvents
    End If
    
    On Error GoTo Errhandler
    
    'Create a file name for the STATA dictionary file (.dct) and STATA program file (.do)
    'and the additional Question Look Up file (_QLU.txt)
    sSTATAdctFile = Mid(sSTATAanaFile, 1, Len(sSTATAanaFile) - 4) & ".dct"
    sSTATAdoFile = Mid(sSTATAanaFile, 1, Len(sSTATAanaFile) - 4) & ".do"
    sSTATAqluFile = Mid(sSTATAanaFile, 1, Len(sSTATAanaFile) - 4) & "_QLU.txt"
    
    'Check for the STATA.dct file already existing and remove it
    If FileExists(sSTATAdctFile) Then
        Kill sSTATAdctFile
        DoEvents
    End If
    
    'Check for the STATA.do file already existing and remove it
    If FileExists(sSTATAdoFile) Then
        Kill sSTATAdoFile
        DoEvents
    End If
    
    'Check for the _QLU.txt file already existing and remove it
    If FileExists(sSTATAqluFile) Then
        Kill sSTATAqluFile
        DoEvents
    End If
    
    'Open the STATA.ana file
    nSTATAanaFileNumber = FreeFile
    Open sSTATAanaFile For Output As #nSTATAanaFileNumber
    
    'Extract STATA.ana file name from name/path sSTATAanaFile
    sFileNameana = StripFileNameFromPath(sSTATAanaFile)
    
    'get the Regional Decimal Point Character
    sLocalDot = RegionalDecimalPointChar
    
    'Loop through grsOutPut reading the contents and writing it out to the STATA.ana file as a fixed length, space delimited record
    grsOutPut.MoveFirst
    lSubRecNo = 0
    Do While Not grsOutPut.EOF
        lSubRecNo = lSubRecNo + 1
        sOutPut = ""
        For i = 0 To grsOutPut.Fields.Count - 1
            'Mo 30/5/2006 Bug 2668, check for Subject Label being excluded from saved output
            If ((i = 2) And (gbExcludeLabel = True)) Then
                'Don't add the Subject Label field
            Else
                Select Case grsOutPut.Fields(i).Type
                Case adDBTimeStamp
                    'Only Dates can be written out as a STATA dates
                    'Times and Date/times have to be written out as strings
                    'Need some way of knowing wether its a date, time or a date/time
                    sSTATADateTime = STATADateFormat(grsOutPut.Fields(i).Value, grsOutPut.Fields(i).Precision, sDateType)
                    sOutPut = sOutPut & sSTATADateTime & msSPACE
                Case adVarChar
                    'Mo 10/1/2006 Bug 2866, change max string length from 80 to 244
                    'STATA strings are limited to 244 char
                    If grsOutPut.Fields(i).DefinedSize > 244 Then
                        If Len(grsOutPut.Fields(i).Value) > 244 Then
                            'Truncate the response to 244 characters
                            sOutPut = sOutPut & Left(grsOutPut.Fields(i).Value, 244) & msSPACE
                        Else
                            'Pad response out to an 244 character field
                            sOutPut = sOutPut & grsOutPut.Fields(i).Value & Space(244 - Len(CStr(RemoveNull(grsOutPut.Fields(i).Value)))) & msSPACE
                        End If
                    Else
                        'Note that strings are Left justified
                        sOutPut = sOutPut & grsOutPut.Fields(i).Value & Space(grsOutPut.Fields(i).DefinedSize - Len(CStr(RemoveNull(grsOutPut.Fields(i).Value)))) & msSPACE
                    End If
                Case adInteger
                    'Integers to be Right justified
                    sOutPut = sOutPut & Space(grsOutPut.Fields(i).Precision - Len(CStr(RemoveNull(grsOutPut.Fields(i).Value)))) & grsOutPut.Fields(i).Value & msSPACE
                Case adSmallInt
                    'Smallints to be Right justified
                    sOutPut = sOutPut & Space(grsOutPut.Fields(i).Precision - Len(CStr(RemoveNull(grsOutPut.Fields(i).Value)))) & grsOutPut.Fields(i).Value & msSPACE
                'Mo 31/1/2007 Bug 2873
                Case adDecimal
                    'Singles to be Right justified
                    sDecimalEnding = ""
                    If grsOutPut.Fields(i).NumericScale > 0 Then
                        'The number of decimal places have been stored in the NumericScale property
                        If InStr(grsOutPut.Fields(i).Value, sLocalDot) = 0 Then
                            'the response has had its decimal point followed by a zero removed. Add it back
                            'Do this for SpecialValues as well
                            sDecimalEnding = sLocalDot & "0"
                        End If
                    End If
                    sOutPut = sOutPut & Space(grsOutPut.Fields(i).Precision - Len(CStr(grsOutPut.Fields(i).Value & sDecimalEnding))) _
                        & grsOutPut.Fields(i).Value & sDecimalEnding & msSPACE
                End Select
            End If
        Next
        'strip off the last msspace
        sOutPut = Mid$(sOutPut, 1, (Len(sOutPut) - 1))
        Print #nSTATAanaFileNumber, sOutPut
        Call DisplayProgressMessage("Saving Subject Record " & lSubRecNo)
        grsOutPut.MoveNext
    Loop
    
    'Open the STATA.dct file
    nSTATAdctFileNumber = FreeFile
    Open sSTATAdctFile For Output As #nSTATAdctFileNumber
    Print #nSTATAdctFileNumber, "dictionary using " & sFileNameana & "{"
    Print #nSTATAdctFileNumber, "* """ & TrialNameFromId(glSelectedTrialId) & """"
    Print #nSTATAdctFileNumber, "* STATA file generated " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Print #nSTATAdctFileNumber,
    'Write the standard subject identification fields to the STATA.dct file
    'Mo 30/5/2006 Bug 2668, check for Subject Label being excluded from saved output
    If (gbExcludeLabel = False) Then
        'Mo 9/6/2006 Bug 2739
        'Mo 9/10/2007 Bug 2941, 2 spaces added after column(n)
        Print #nSTATAdctFileNumber, "_column(1)     str15  Trial" & Space((gnShortCodeLength * 2) - 3) & "%15s   ""Study Name"""
        Print #nSTATAdctFileNumber, "_column(17)    str8   Site" & Space((gnShortCodeLength * 2) - 2) & "%8s    ""Study Site"""
        Print #nSTATAdctFileNumber, "_column(26)    str50  Label" & Space((gnShortCodeLength * 2) - 3) & "%50s   ""Subject Label"""
        Print #nSTATAdctFileNumber, "_column(77)           Personid" & Space((gnShortCodeLength * 2) - 6) & "%10f   ""Subject Id"""
        Print #nSTATAdctFileNumber, "_column(88)           VisCycle" & Space((gnShortCodeLength * 2) - 6) & "%5f    ""Visit Cycle Number"""
        Print #nSTATAdctFileNumber, "_column(94)           FrmCycle" & Space((gnShortCodeLength * 2) - 6) & "%5f    ""Form Cycle Number"""
        Print #nSTATAdctFileNumber, "_column(100)          RepeatNo" & Space((gnShortCodeLength * 2) - 6) & "%5f    ""Question Repeat Number"""
        lColumnNo = 106
    Else
        'Mo 9/6/2006 Bug 2739
        'Mo 9/10/2007 Bug 2941, 2 spaces added after column(n)
        Print #nSTATAdctFileNumber, "_column(1)     str15  Trial" & Space((gnShortCodeLength * 2) - 3) & "%15s   ""Study Name"""
        Print #nSTATAdctFileNumber, "_column(17)    str8   Site" & Space((gnShortCodeLength * 2) - 2) & "%8s    ""Study Site"""
        Print #nSTATAdctFileNumber, "_column(26)           Personid" & Space((gnShortCodeLength * 2) - 6) & "%10f   ""Subject Id"""
        Print #nSTATAdctFileNumber, "_column(37)           VisCycle" & Space((gnShortCodeLength * 2) - 6) & "%5f    ""Visit Cycle Number"""
        Print #nSTATAdctFileNumber, "_column(43)           FrmCycle" & Space((gnShortCodeLength * 2) - 6) & "%5f    ""Form Cycle Number"""
        Print #nSTATAdctFileNumber, "_column(49)           RepeatNo" & Space((gnShortCodeLength * 2) - 6) & "%5f    ""Question Repeat Number"""
        lColumnNo = 55
    End If
    
    'Extract STATA.dct file name from name/path sSTATAdctFile
    sFileNamedct = StripFileNameFromPath(sSTATAdctFile)
    
    'Open the STATA.do file
    nSTATAdoFileNumber = FreeFile
    Open sSTATAdoFile For Output As #nSTATAdoFileNumber
    Print #nSTATAdoFileNumber, "#delimit ;"
    Print #nSTATAdoFileNumber,
    Print #nSTATAdoFileNumber, "label data """ & TrialNameFromId(glSelectedTrialId) & """;"
    Print #nSTATAdoFileNumber,
    
    'Open the _QLU.txt file
    nSTATAqluFileNumber = FreeFile
    Open sSTATAqluFile For Output As #nSTATAqluFileNumber
    
    nQuestionsCounter = 0
    
    'Loop through the questions adding them to STATA.dct, STATA.do and the _QLU.txt file
    For i = 7 To grsOutPut.Fields.Count - 1
        'either read sVFQA from the gColQuestionCodes collection or from grsOutPut's column headers
        If gbUseShortCodes Then
            sVFQA = gColQuestionCodes(grsOutPut.Fields(i).Name)
        Else
            sVFQA = grsOutPut.Fields(i).Name
        End If
        'sVFQA will be of the format VisitCode/CRFPageCode/DataItemCode or VisitCode/CRFPageCode/DataItemCode/Attribute
        'extract DataItemCode from question column name sVFQA
        asVFQA = Split(sVFQA, "/")
        sDataItemCode = asVFQA(2)
        'STATA export uses 8 character SAS style short codes
        'if short have already been created for display purposes they will be read from grsOutPut's column headers
        'otherwise they are created by a call to CreateUniqueQuestion
        If gbUseShortCodes Then
            sShortCode = grsOutPut.Fields(i).Name
        Else
            'create a unique 8 character (max) question code from the DataItemCode
            sShortCode = CreateUniqueQuestion(sDataItemCode, sDataItemCode)
        End If
        'test for a question or a question attribute
        If UBound(asVFQA) = 3 Then
            'its an attribute
            sAttribute = asVFQA(3)
            Select Case sAttribute
            Case "Comments"
                'Mo 10/1/2006 Bug 2866, change max string length from 80 to 244
                'STATA strings are limited to 244 char
                'Mo 9/6/2006 Bug 2739
                'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                    & "str244 " & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%244s  ""Comment"""
                lColumnNo = lColumnNo + 245
            Case "CTCGrade"
                'Mo 9/6/2006 Bug 2739
                'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                    & "str1   " & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%1s    ""CTCGrade"""
                lColumnNo = lColumnNo + 2
            Case "LabResult"
                'Mo 9/6/2006 Bug 2739
                'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                    & "str1   " & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%1s    ""LabResult"""
                lColumnNo = lColumnNo + 2
            Case "Status"
                'Mo 9/6/2006 Bug 2739
                'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                    & "str15  " & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%15s   ""Status"""
                lColumnNo = lColumnNo + 16
            Case "TimeStamp"
                'Mo 9/6/2006 Bug 2739
                'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                'Note Timestamps are Date/time fields and have to be written out as 19 char strings in STATA
                Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                    & "str19  " & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%19s   ""UserName"""
                lColumnNo = lColumnNo + 20
            Case "UserName"
                'Mo 9/6/2006 Bug 2739
                'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                    & "str20  " & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%20s   ""UserName"""
                lColumnNo = lColumnNo + 21
            Case "ValueCode"
                'Mo 9/6/2006 Bug 2739
                'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                    & "str15  " & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%15s   ""ValueCode"""
                lColumnNo = lColumnNo + 16
            End Select
            'Write the Long form and the Short form of the current attribute/question to the _QLU.txt file
            'Mo 26/5/2006 Bug 2738
            Print #nSTATAqluFileNumber, sVFQA & msTAB & sShortCode & msTAB & rsDataItem!DataItemName & msTAB & sAttribute
        Else
            'Get Question/DataItem details
            Set rsDataItem = New ADODB.Recordset
            Set rsDataItem = DataItemDetails(glSelectedTrialId, sDataItemCode)
            sDataItemDescription = rsDataItem!DataItemName
            'Based on the Question type Write to the STATA.dct file
            Select Case rsDataItem!DataType
            'Mo 25/10/2005 COD0040
            Case DataType.Text, DataType.Thesaurus
                'Mo 10/1/2006 Bug 2866, change max string length from 80 to 244
                'STATA strings are limited to 244 char
                If rsDataItem!DataItemLength > 244 Then
                    'Reduce field/string size to 244 chars
                    'Mo 9/6/2006 Bug 2739
                    'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                    Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                        & "str244 " & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%244s  """ & sDataItemDescription & """"
                    lColumnNo = lColumnNo + 245
                    'Add to STATA Replace collection
                    nQuestionsCounter = nQuestionsCounter + 1
                    gColSTATADetails.Add sShortCode & "|STRING|244", CStr(nQuestionsCounter)
                ElseIf rsDataItem!DataItemLength = 1 Then
                    'Force single character fields to be two character fields that are
                    'capable of holding a special value string of "-1" to "-9"
                    'Mo 9/6/2006 Bug 2739
                    'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                    Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                        & "str2   " & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%2s    """ & sDataItemDescription & """"
                    lColumnNo = lColumnNo + 3
                    'Add to STATA Replace collection
                    nQuestionsCounter = nQuestionsCounter + 1
                    gColSTATADetails.Add sShortCode & "|STRING|2", CStr(nQuestionsCounter)
                Else
                    'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                    Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                        & "str" & rsDataItem!DataItemLength & Space(4 - Len(CStr(rsDataItem!DataItemLength))) _
                        & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%" & rsDataItem!DataItemLength & "s" _
                        & Space(5 - Len(CStr(rsDataItem!DataItemLength))) & """" & sDataItemDescription & """"
                    lColumnNo = lColumnNo + rsDataItem!DataItemLength + 1
                    'Add to STATA Replace collection
                    nQuestionsCounter = nQuestionsCounter + 1
                    gColSTATADetails.Add sShortCode & "|STRING|" & rsDataItem!DataItemLength, CStr(nQuestionsCounter)
                End If
            Case DataType.Date
                'Only Dates can be written out as a STATA dates
                'Times and Date/times have to be written out as strings
                'Mo 18/10/2006 Bug 2822, check Partial Dates flag before deciding on date format
                If CInt(RemoveNull(rsDataItem!DataItemCase)) = 0 And STATADateTest(rsDataItem!DataItemFormat) Then
                    'Check for "Standard" or "Float" dates
                    If sDateType = "Float" Then
                        'Mo 9/6/2006 Bug 2739
                        'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                        'Its a Date field that will be written into a STATA %8f Long Date field as ddmmyyyy
                        Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                        & "long   " & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%8f    """ & sDataItemDescription & """"
                        lColumnNo = lColumnNo + 9
                        'Add to STATA Replace collection
                        nQuestionsCounter = nQuestionsCounter + 1
                        gColSTATADetails.Add sShortCode & "|NUMERIC", CStr(nQuestionsCounter)
                    Else
                        'Its a Date field
                        'Mo 11/1/2006 Bug 2671, standard date format changed from %d to %8.0g
                        'Mo 9/6/2006 Bug 2739
                        'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                        Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                        & "int    " & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%8.0g  """ & sDataItemDescription & """"
                        lColumnNo = lColumnNo + 9
                        'Add to STATA Replace collection
                        nQuestionsCounter = nQuestionsCounter + 1
                        'Mo 11/1/2006 Bug 2671
                        gColSTATADetails.Add sShortCode & "|NUMERIC", CStr(nQuestionsCounter)
                    End If
                Else
                    'Its a Date/Time or Time field that will be written into a string field
                    'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                    Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                    & "str" & rsDataItem!DataItemLength & Space(4 - Len(CStr(rsDataItem!DataItemLength))) _
                    & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%" & rsDataItem!DataItemLength & "s" _
                        & Space(5 - Len(CStr(rsDataItem!DataItemLength))) & """" & sDataItemDescription & """"
                    lColumnNo = lColumnNo + rsDataItem!DataItemLength + 1
                    'Add to STATA Replace collection
                    nQuestionsCounter = nQuestionsCounter + 1
                    gColSTATADetails.Add sShortCode & "|STRING|" & rsDataItem!DataItemLength, CStr(nQuestionsCounter)
                End If
            Case DataType.Category
                'if gbOutputCategoryCodes = False (i.e. category Values are being output instead of
                'category Codes then all Category Questions will be treated as Text questions
                If gbOutputCategoryCodes Then
                    'its a Category Questions containing category codes assess the usage of Numeric or Alpha codes
                    Call AssessCategoryCodes(glSelectedTrialId, 1, rsDataItem!DataItemId, bCatCodesNumeric, nCatCodeLength)
                    'Force single character fields to be two character fields that are
                    'capable of holding a special value string of "-1" to "-9"
                    If nCatCodeLength = 1 Then
                        nCatCodeLength = 2
                    End If
                    If bCatCodesNumeric Then
                        'create a name for the Numeric Category Codes
                        sCatCodeName = sShortCode & "_"
                        'Mo 9/6/2006 Bug 2739
                        'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                        Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                            & "int    " & sShortCode & ":" & sCatCodeName & Space(((gnShortCodeLength * 2) + 1) - Len(sShortCode) - Len(sCatCodeName)) _
                            & "%" & nCatCodeLength & "f" _
                            & Space(5 - Len(CStr(nCatCodeLength))) & """" & sDataItemDescription & """"
                        lColumnNo = lColumnNo + nCatCodeLength + 1
                        'Add to STATA Replace collection
                        nQuestionsCounter = nQuestionsCounter + 1
                        gColSTATADetails.Add sShortCode & "|NUMERIC", CStr(nQuestionsCounter)
                        'Write the Numeric category codes and their values to the STATA.do file
                        Call WriteCatCodesToSTATA(glSelectedTrialId, 1, rsDataItem!DataItemId, sCatCodeName, nSTATAdoFileNumber)
                    Else
                        'Treat Category Questions with Alpha codes as a Text Question of length nCatCodeLength
                        'Mo 9/6/2006 Bug 2739
                        'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                        Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                            & "str" & nCatCodeLength & Space(4 - Len(CStr(nCatCodeLength))) _
                            & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%" & nCatCodeLength & "s" _
                            & Space(5 - Len(CStr(nCatCodeLength))) & """" & sDataItemDescription & """"
                        lColumnNo = lColumnNo + nCatCodeLength + 1
                        'Add to STATA Replace collection
                        nQuestionsCounter = nQuestionsCounter + 1
                        gColSTATADetails.Add sShortCode & "|STRING|" & nCatCodeLength, CStr(nQuestionsCounter)
                    End If
                Else
                    'Treat Category Question with category values as a Text Question of length rsDataItem!DataItemLength
                    'Mo 9/6/2006 Bug 2739
                    'Mo 1/8/2006 Bug 2775, Force single character fields to be two character fields
                    'that are capable of holding a special value string of "-1" to "-9"
                    If rsDataItem!DataItemLength = 1 Then
                        'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                        Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                            & "str2   " & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%2s    """ & sDataItemDescription & """"
                        lColumnNo = lColumnNo + 3
                    Else
                        'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                        Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                            & "str" & rsDataItem!DataItemLength & Space(4 - Len(CStr(rsDataItem!DataItemLength))) _
                            & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%" & rsDataItem!DataItemLength & "s" _
                            & Space(5 - Len(CStr(rsDataItem!DataItemLength))) & """" & sDataItemDescription & """"
                        lColumnNo = lColumnNo + rsDataItem!DataItemLength + 1
                    End If
                    'Add to STATA Replace collection
                    nQuestionsCounter = nQuestionsCounter + 1
                    gColSTATADetails.Add sShortCode & "|STRING|" & rsDataItem!DataItemLength, CStr(nQuestionsCounter)
                End If
            Case DataType.Multimedia
                'Multimedia fields are always 36 char long
                'Mo 9/6/2006 Bug 2739
                'Mo 9/10/2007 Bug 2941, "Space(4" changed to "Space(6"
                Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(6 - Len(CStr(lColumnNo))) _
                    & "str36  " & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%36s   """ & sDataItemDescription & """"
                lColumnNo = lColumnNo + 37
                'Add to STATA Replace collection
                nQuestionsCounter = nQuestionsCounter + 1
                gColSTATADetails.Add sShortCode & "|STRING|36", CStr(nQuestionsCounter)
            Case DataType.IntegerData
                'Mo 9/6/2006 Bug 2739
                'Mo 1/8/2006 Bug 2775, Force single character fields to be two character fields
                'that are capable of holding a special value string of "-1" to "-9"
                If rsDataItem!DataItemLength = 1 Then
                    'Mo 9/10/2007 Bug 2941, "Space(11" changed to "Space(13"
                    Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(13 - Len(CStr(lColumnNo))) _
                        & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%2f    """ & sDataItemDescription & """"
                    lColumnNo = lColumnNo + 3
                Else
                    'Mo 9/10/2007 Bug 2941, "Space(11" changed to "Space(13"
                    Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(13 - Len(CStr(lColumnNo))) _
                        & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%" & rsDataItem!DataItemLength & "f" _
                        & Space(5 - Len(CStr(rsDataItem!DataItemLength))) & """" & sDataItemDescription & """"
                    lColumnNo = lColumnNo + rsDataItem!DataItemLength + 1
                End If
                'Add to STATA Replace collection
                nQuestionsCounter = nQuestionsCounter + 1
                gColSTATADetails.Add sShortCode & "|NUMERIC", CStr(nQuestionsCounter)
            Case DataType.Real, DataType.LabTest
                'check for a real/labtest that does not have any decimal places
                If InStr(rsDataItem!DataItemFormat, ".") = 0 Then
                    'Mo 9/6/2006 Bug 2739
                    'Mo 9/10/2007 Bug 2941, "Space(11" changed to "Space(13"
                    Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(13 - Len(CStr(lColumnNo))) _
                        & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%" & rsDataItem!DataItemLength & ".0f" _
                        & Space(3 - Len(CStr(rsDataItem!DataItemLength))) & """" & sDataItemDescription & """"
                    lColumnNo = lColumnNo + rsDataItem!DataItemLength + 1
                Else
                    'Mo 9/6/2006 Bug 2739
                    'Mo 9/10/2007 Bug 2941, "Space(11" changed to "Space(13"
                    Print #nSTATAdctFileNumber, "_column(" & lColumnNo & ")" & Space(13 - Len(CStr(lColumnNo))) _
                        & sShortCode & Space(((gnShortCodeLength * 2) + 2) - Len(sShortCode)) & "%" & rsDataItem!DataItemLength & "." & rsDataItem!DataItemLength - InStr(rsDataItem!DataItemFormat, ".") & "f" _
                        & Space(3 - Len(CStr(rsDataItem!DataItemLength))) & """" & sDataItemDescription & """"
                    lColumnNo = lColumnNo + rsDataItem!DataItemLength + 1
                End If
                'Add to STATA Replace collection
                nQuestionsCounter = nQuestionsCounter + 1
                gColSTATADetails.Add sShortCode & "|NUMERIC", CStr(nQuestionsCounter)
            End Select
            'Write the Long form and the Short form of the current attribute/question to the _QLU.txt file
            'Mo 26/5/2006 Bug 2738
            Print #nSTATAqluFileNumber, sVFQA & msTAB & sShortCode & msTAB & rsDataItem!DataItemName & msTAB & GetDataTypeString(rsDataItem!DataType)
        End If
    Next i
    rsDataItem.Close
    Set rsDataItem = Nothing
          
    'Write the closing bracket to the STATA.dct file
    Print #nSTATAdctFileNumber,
    Print #nSTATAdctFileNumber, "}"
    
    'Write the ending parts of the STATA.do file
    Print #nSTATAdoFileNumber,
    Print #nSTATAdoFileNumber, "infile using " & sFileNamedct & ";"
    'Write the replace section
    Call STATAReplaceSection(nSTATAdoFileNumber, nQuestionsCounter)
    
    'Close the files
    Close #nSTATAanaFileNumber
    Close #nSTATAdctFileNumber
    Close #nSTATAdoFileNumber
    Close #nSTATAqluFileNumber
    
    Set gColSTATADetails = Nothing
    
    Call DisplayProgressMessage("Save Output (STATA) completed.")
    
    Call HourglassOff
    
Exit Sub
Errhandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "OutputToSTATA", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

CancelSaveAs:
    Call HourglassOff

End Sub

'---------------------------------------------------------------------
Private Function SASDateFormatType(ByVal sDateFormat As String) As String
'---------------------------------------------------------------------
'This function is similar to DateFormatCanBeConverted.
'It decides if the DataItemFormat allows the field to be converted.
'and then it decides if it is a Date, a Time or a Date/Time, because
'they are stored differently in SAS.
'
'Dates are placed in        ddmmyy10.   fields as   dd/mm/yyyy
'Times are placed in        time.       fields as   hh:mm:ss
'Date/Times are placed in   datetime19.  fields as   ddMMMyyyy:hh:mm:ss
'
'In summary this function will return ddmmyy10. or time. or datetime19.
'For dates that can not be converted it returns $20.
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'Replace double format characters with single format characters
    sDateFormat = Replace(sDateFormat, "dd", "d")
    sDateFormat = Replace(sDateFormat, "mm", "m")
    sDateFormat = Replace(sDateFormat, "hh", "h")
    sDateFormat = Replace(sDateFormat, "ss", "s")
    'Replace yyyy with y
    sDateFormat = Replace(sDateFormat, "yyyy", "y")
    'Replace all Date/Time Separators with "/"
    sDateFormat = Replace(sDateFormat, ":", "/")
    sDateFormat = Replace(sDateFormat, ".", "/")
    sDateFormat = Replace(sDateFormat, "-", "/")
    sDateFormat = Replace(sDateFormat, " ", "/")
    
    Select Case sDateFormat
    Case "d/m/y", "m/d/y", "y/m/d"
        SASDateFormatType = "ddmmyy10."
    Case "h/m", "h/m/s"
        SASDateFormatType = "time."
    Case "d/m/y/h/m", "m/d/y/h/m", "y/m/d/h/m", "d/m/y/h/m/s", "m/d/y/h/m/s", "y/m/d/h/m/s"
        SASDateFormatType = "datetime19."
    Case Else
        'date formats y/m, m/y and y/d/m (with or without time elements) are not converted
        SASDateFormatType = "$20."
    End Select

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "SASDateFormatType", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Private Function SASDateFormat(sDateString) As Variant
'---------------------------------------------------------------------
'The passed Date String will be be one of the following formats:-
'   dd/mm/yyyy              (length 10)
'   hh:mm:ss                (length 8)
'   dd/mm/yyyy hh:mm:ss    (length 19)
'
'This function will transform the dd/mm/yyyy hh:mm:ss date/time string into a
'ddMMMYYYY:hh:mm:ss date string as required by SAS's datetime19. field format
'---------------------------------------------------------------------
Dim s2DigMonth As String
Dim sAlphaMonth As String

    On Error GoTo Errhandler

    If Len(sDateString) = 19 Then
        'Extract the month
        s2DigMonth = Mid(sDateString, 4, 2)
        'convert Digit month to Alpha month
        Select Case s2DigMonth
        Case "01"
            sAlphaMonth = "Jan"
        Case "02"
            sAlphaMonth = "Feb"
        Case "03"
            sAlphaMonth = "Mar"
        Case "04"
            sAlphaMonth = "Apr"
        Case "05"
            sAlphaMonth = "May"
        Case "06"
            sAlphaMonth = "Jun"
        Case "07"
            sAlphaMonth = "Jul"
        Case "08"
            sAlphaMonth = "Aug"
        Case "09"
            sAlphaMonth = "Sep"
        Case "10"
            sAlphaMonth = "Oct"
        Case "11"
            sAlphaMonth = "Nov"
        Case "12"
            sAlphaMonth = "Dec"
        End Select
        'create a ddMMMYYYY:hh:mm:ss formated date/time string
        SASDateFormat = Mid(sDateString, 1, 2) & sAlphaMonth & Mid(sDateString, 7, 4) & ":" & Mid(sDateString, 12, 8)
    Else
        'return the string unchanged
        SASDateFormat = sDateString
    End If

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "SASDateFormat", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'Mo 2/4/2007 MRC15022007, Optional parameter bShowDialogs added
'--------------------------------------------------------------------
Public Sub OutputToMACROBD(Optional bShowDialogs As Boolean = True)
'--------------------------------------------------------------------
'This will save the results of the current query in
'MACRO's Batch Data Entry Format.
'--------------------------------------------------------------------
Dim sExportFileBD As String
Dim nBDFileNumber As Integer
Dim i As Integer
Dim sBatchResponse As String
Dim asVFQCodes() As String
Dim sVisitCode As String
Dim sFormCode As String
Dim sQuestionCode As String
Dim sResponse As String

    On Error GoTo CancelSaveAs
    
    Call HourglassOn
    
    'sExportFileBD = gsOUT_FOLDER_LOCATION & TrialNameFromId(glSelectedTrialId) & "_" & Format(Now, "yyyymmdd") & "BatchData.csv"
    'Mo 2/4/2007 MRC15022007
    sExportFileBD = ConstructOutputFileName & "BatchData.csv"
    
    If bShowDialogs Then
        'Prepare and launch the Batch Data Save As dialog
        With frmMenu.CommonDialog1
            .DialogTitle = "MACRO Query Save as Batch Data export file"
            .CancelError = True
            .Filter = "CSV (Comma delimited) (*.csv)|*.csv"
            .DefaultExt = "csv"
            .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
            .FileName = sExportFileBD
            .ShowSave
    
            sExportFileBD = .FileName
        End With
    End If
    
    'Check for the Batch Data file already existing and remove it
    If FileExists(sExportFileBD) Then
        Kill sExportFileBD
        DoEvents
    End If
    
    On Error GoTo Errhandler
    
    'Open the Batch Data file
    nBDFileNumber = FreeFile
    Open sExportFileBD For Output As #nBDFileNumber
    
    'Loop through grsOutPut reading the contents and placing it into the Batch Data file
    grsOutPut.MoveFirst
    Do While Not grsOutPut.EOF
        'The first 7 fields in grsOutPut are the subject identification fields
        '"Trial","Site", "Label", "PersonId", "VisitCycle", "FormCycle", "RepeatNumber"
        'These 7 fields stipulate a specific study/subject/visit/form instance to which the
        'remaining question fields (8 and onwards) pertain to
        For i = 7 To grsOutPut.Fields.Count - 1
            'Only create Batch Responses for non null response
            If Not IsNull(grsOutPut.Fields(i).Value) Then
                If gbUseShortCodes Then
                    'The required Visit, Form and Question codes are retieved from the
                    'short codes to long codes colection
                    asVFQCodes = Split(gColQuestionCodes(grsOutPut.Fields(i).Name), "/")
                Else
                    'The required Visit, Form and Question codes are held within the columns heading (Caption)
                    asVFQCodes = Split(grsOutPut.Fields(i).Name, "/")
                End If
                'Do not create Batch Responses for Question attributes
                'Note that Question attributew will have ubound=3 and normal questions will have ubound=2
                If UBound(asVFQCodes) = 2 Then
                    sVisitCode = asVFQCodes(0)
                    sFormCode = asVFQCodes(1)
                    sQuestionCode = asVFQCodes(2)
                    'Do not create Batch Responses for Derived questions
                    If Not QuestionIsDerived(sQuestionCode) Then
                        'Mo 3/11/2006 Bug 2834
                        'if its a date/time response get the original response from the database, not the
                        'response that has been manipulated and transformed by Query Module code
                        If grsOutPut.Fields(i).Type = adDBTimeStamp Then
                            sResponse = GetSingleResponse(grsOutPut.Fields("Trial"), grsOutPut.Fields("Site"), grsOutPut.Fields("PersonId"), _
                                sVisitCode, grsOutPut.Fields("VisitCycle"), sFormCode, grsOutPut.Fields("FormCycle"), sQuestionCode, grsOutPut.Fields("RepeatNumber"))
                        Else
                            sResponse = grsOutPut.Fields(i).Value
                        End If
                        sBatchResponse = grsOutPut.Fields("Trial") & "," & grsOutPut.Fields("Site") & "," _
                            & grsOutPut.Fields("PersonId") & ",," _
                            & sVisitCode & "," & grsOutPut.Fields("VisitCycle") & ",," _
                            & sFormCode & "," & grsOutPut.Fields("FormCycle") & ",," _
                            & sQuestionCode & "," & grsOutPut.Fields("RepeatNumber") & "," & CSVCommasAndQuotes(sResponse)
                        Print #nBDFileNumber, sBatchResponse
                    End If
                End If
            End If
        Next i
        grsOutPut.MoveNext
    Loop 'on recordsets within grsOutPut
    
    'Close the files
    Close #nBDFileNumber
    
    Call DisplayProgressMessage("Save Output (MACRO Batch Data) completed.")
    
    Call HourglassOff

Exit Sub
Errhandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "OutputToMACROBD", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

CancelSaveAs:
    Call HourglassOff

End Sub

'--------------------------------------------------------------------
Private Function QuestionIsDerived(ByVal sDataItemCode) As Boolean
'--------------------------------------------------------------------
'Returns True if a Question is a Derived question
'Returns False for non-Derived questions
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bFunctionOutCome As Boolean

    On Error GoTo Errhandler
    
    sSQL = "SELECT DataItem.Derivation FROM DataItem " _
        & "WHERE DataItem.ClinicalTrialId = " & glSelectedTrialId _
        & " AND DataItem.DataItemCode = '" & sDataItemCode & "'"
        
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount <> 1 Then
        bFunctionOutCome = False
    Else
        If RemoveNull(rsTemp!Derivation) = "" Then
            bFunctionOutCome = False
        Else
            bFunctionOutCome = True
        End If
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    QuestionIsDerived = bFunctionOutCome

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "QuestionIsDerived", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Private Function DataItemDetails(ByVal lClinicalTrialId As Long, _
                           ByVal sDataItemCode As String) As ADODB.Recordset
'---------------------------------------------------------------------
'Creates a ReadOnly recordset.
'Retrieve details of a CRF page.
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo Errhandler
    
    'Mo 18/10/2006 Bug 2822, DataItemCase added to following SQL
    sSQL = "SELECT DataItemId, DataItemName, DataType, DataItemLength, DataItemFormat, DataItemCase FROM DataItem " _
        & "WHERE ClinicalTrialId = " & lClinicalTrialId & " " _
        & "AND DataItemCode = '" & sDataItemCode & "'"
    
    Set DataItemDetails = New ADODB.Recordset
    DataItemDetails.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DataItemDetails", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Public Sub AssessCategoryCodes(ByVal lClinicalTrialId As Long, _
                                ByVal nVersion As Integer, _
                                ByVal lDataItemId As Long, _
                                ByRef bCatCodesNumeric As Boolean, _
                                ByRef nCatCodeLength As Integer)
'---------------------------------------------------------------------
'Function to assess the usage of Numeric or Alpha codes
'Returns True if all codes are Numeric
'Returns False otherwise
'---------------------------------------------------------------------
Dim rsValueData As ADODB.Recordset

    On Error GoTo Errhandler

    Set rsValueData = New ADODB.Recordset
    'Mo 21/8/2006 Bug 2784, call to gdsDataValues replaced gdsDataValuesALL
    Set rsValueData = gdsDataValuesALL(lClinicalTrialId, nVersion, lDataItemId)
    'Loop through the category codes assesing the type (numeric or string) and the length
    'Mo 21/8/2006 Bug 2784, check for no category codes added
    If rsValueData.RecordCount = 0 Then
        'when no category codes exist, length set to 2 and type set to nonnumeric
        nCatCodeLength = 2
        bCatCodesNumeric = False
        Exit Sub
    End If
    rsValueData.MoveFirst
    nCatCodeLength = 0
    bCatCodesNumeric = True
    Do While Not rsValueData.EOF
        If Len(rsValueData!ValueCode) > nCatCodeLength Then
            nCatCodeLength = Len(rsValueData!ValueCode)
        End If
        If Not IsNumeric(rsValueData!ValueCode) Then
            bCatCodesNumeric = False
        End If
        rsValueData.MoveNext
    Loop

Exit Sub
Errhandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "AssessCategoryCodes", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub WriteCatCodesToSTATA(ByVal lClinicalTrialId As Long, _
                                ByVal nVersion As Integer, _
                                ByVal lDataItemId As Long, _
                                ByVal sCatCodeName As String, _
                                ByVal nSTATAdoFileNumber As Integer)
'---------------------------------------------------------------------
Dim rsValueData As ADODB.Recordset
Dim sOutPut As String
Dim q As Integer

    On Error GoTo Errhandler

    Set rsValueData = New ADODB.Recordset
    'Mo 21/8/2006 Bug 2784, call to gdsDataValues replaced gdsDataValuesALL
    Set rsValueData = gdsDataValuesALL(lClinicalTrialId, nVersion, lDataItemId)
    
    'Mo 21/8/2006 Bug 2784, check for no category codes
    If rsValueData.RecordCount > 0 Then
        rsValueData.MoveFirst
        q = 0
        'loop through the category codes and write them to the STATA.do file
        Do While Not rsValueData.EOF
            q = q + 1
            If q = 1 Then
                'Mo 9/6/2006 Bug 2739
                sOutPut = "label define " & sCatCodeName & Space((gnShortCodeLength + 2) - Len(sCatCodeName)) & rsValueData!ValueCode _
                    & Space(3 - Len(rsValueData!ValueCode)) & """" & rsValueData!ItemValue & """"
            Else
                sOutPut = Space(gnShortCodeLength + 15) & rsValueData!ValueCode & Space(3 - Len(rsValueData!ValueCode)) & """" & rsValueData!ItemValue & """"
            End If
            If q = rsValueData.RecordCount Then
                sOutPut = sOutPut & " ;"
            End If
            Print #nSTATAdoFileNumber, sOutPut
            rsValueData.MoveNext
        Loop
    End If
    
    rsValueData.Close
    Set rsValueData = Nothing

Exit Sub
Errhandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "WriteCatCodesToSTATA", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub WriteCatCodesToSAS(ByVal lClinicalTrialId As Long, _
                                ByVal nVersion As Integer, _
                                ByVal lDataItemId As Long, _
                                ByVal sShortCode As String, _
                                ByVal nSASCategoryFileNumber As Integer, _
                                ByVal bCatCodesNumeric As Boolean)
'---------------------------------------------------------------------
Dim rsValueData As ADODB.Recordset
Dim sOutPut As String
Dim q As Integer

    On Error GoTo Errhandler

    Set rsValueData = New ADODB.Recordset
    'Mo 21/8/2006 Bug 2784, call to gdsDataValues replaced gdsDataValuesALL
    Set rsValueData = gdsDataValuesALL(lClinicalTrialId, nVersion, lDataItemId)
    
    'Mo 21/8/2006 Bug 2784, check for no category codes
    If rsValueData.RecordCount > 0 Then
        rsValueData.MoveFirst
        q = 0
        'Loop throught the category codes and write them to the SAS Category.txt file
        Do While Not rsValueData.EOF
            q = q + 1
            If q = 1 Then
                If bCatCodesNumeric Then
                    sOutPut = "Value " & sShortCode & " " & rsValueData!ValueCode & "='" & rsValueData!ItemValue & "'"
                Else
                    sOutPut = "Value $" & sShortCode & " '" & rsValueData!ValueCode & "'='" & rsValueData!ItemValue & "'"
                End If
            Else
                If bCatCodesNumeric Then
                    sOutPut = String(Len(sShortCode) + 7, " ") & rsValueData!ValueCode & "='" & rsValueData!ItemValue & "'"
                Else
                    sOutPut = String(Len(sShortCode) + 8, " ") & "'" & rsValueData!ValueCode & "'='" & rsValueData!ItemValue & "'"
                End If
            End If
            'The last category value needs a trailing ";"
            If q = rsValueData.RecordCount Then
                sOutPut = sOutPut & ";"
            End If
            Print #nSASCategoryFileNumber, sOutPut
            rsValueData.MoveNext
        Loop
    End If
    
    rsValueData.Close
    Set rsValueData = Nothing

Exit Sub
Errhandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "WriteCatCodesToSAS", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'---------------------------------------------------------------------
Private Function STATADateFormat(ByVal vDateString As Variant, _
                                ByVal nDateTimeType As Integer, _
                                ByVal sDateType As String) As String
'---------------------------------------------------------------------
'The passed Date String will be be one of the following formats:-
'   Dates       dd/mm/yyyy              (length 10)
'   Times       hh:mm:ss                (length 8)
'   Date/Times  dd/mm/yyyy hh:mm:ss    (length 19)
'
'Dates of format dd/mm/yyyy will be returned as dates and written into STATA as %8f Long Date fields or %8.0g Standard date numbers.
'Date/Times and Times will be returned left justified in a string field
'Mo 11/1/2006 Bug 2671, standard date format changed from %d to %8.0g
'---------------------------------------------------------------------
Dim nMonth As Integer
Dim s3CharMonth As String

    On Error GoTo Errhandler
    
    'convert a special value date back into negative special value
    'The following Special Value converted dates are a mix of D/M/Ys and M/D/Ys
    Select Case vDateString
    Case "29/12/1899", "12/29/1899"
        vDateString = -1
    Case "28/12/1899", "12/28/1899"
        vDateString = -2
    Case "27/12/1899", "12/27/1899"
        vDateString = -3
    Case "26/12/1899", "12/26/1899"
        vDateString = -4
    Case "25/12/1899", "12/25/1899"
        vDateString = -5
    Case "24/12/1899", "12/24/1899"
        vDateString = -6
    Case "23/12/1899", "12/23/1899"
        vDateString = -7
    Case "22/12/1899", "12/22/1899"
        vDateString = -8
    Case "21/12/1899", "12/21/1899"
        vDateString = -9
    End Select
    
    If (Not IsNull(vDateString)) And (nDateTimeType = 2) And (vDateString <> "-1") And (vDateString <> "-2") _
    And (vDateString <> "-3") And (vDateString <> "-4") And (vDateString <> "-5") And (vDateString <> "-6") _
    And (vDateString <> "-7") And (vDateString <> "-8") And (vDateString <> "-9") Then
        'Check for "Standard" or "Float" dates
        If sDateType = "Float" Then
            'create a ddmmyyyy "Float" date (e.g. 01012004 for 1 January 2004)
            STATADateFormat = Mid(Format(vDateString, "DD/MM/YYYY"), 1, 2) & Mid(Format(vDateString, "DD/MM/YYYY"), 4, 2) & Mid(Format(vDateString, "DD/MM/YYYY"), 7, 4)
        Else
            'Mo 11/1/2006 Bug 2671
            'create a Standard date number ( e.g. 0 = 01/01/1960, 1 = 02/01/1960)
            'note that VB's numeric dates are based on 1 being 31/12/1899, 2 being 1/1/1900
            'to switch from VB to STATA dates subtract 21916
            STATADateFormat = CDbl(vDateString) - 21916
            STATADateFormat = STATADateFormat & Space(8 - Len(CStr(RemoveNull(STATADateFormat))))
        End If
    Else
        Select Case nDateTimeType
        Case 1  'Date/Time
            'strip off any " AM" or " PM"
            If InStr(vDateString, " AM") > 0 Then
                vDateString = Mid(vDateString, 1, (InStr(vDateString, " AM") - 1))
            End If
            If InStr(vDateString, " PM") > 0 Then
                vDateString = Mid(vDateString, 1, (InStr(vDateString, " PM") - 1))
            End If
            STATADateFormat = vDateString & Space(19 - Len(CStr(RemoveNull(vDateString))))
        Case 2  'Date
            If sDateType = "Float" Then
                STATADateFormat = vDateString & Space(8 - Len(CStr(RemoveNull(vDateString))))
            Else
                'Mo 11/1/2006 Bug 2671, width changed from 9 to 8
                STATADateFormat = vDateString & Space(8 - Len(CStr(RemoveNull(vDateString))))
            End If
        Case 3  'Time
            'strip off any " AM" or " PM"
            If InStr(vDateString, " AM") > 0 Then
                vDateString = Mid(vDateString, 1, (InStr(vDateString, " AM") - 1))
            End If
            If InStr(vDateString, " PM") > 0 Then
                vDateString = Mid(vDateString, 1, (InStr(vDateString, " PM") - 1))
            End If
            STATADateFormat = vDateString & Space(8 - Len(CStr(RemoveNull(vDateString))))
        End Select
    End If

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "STATADateFormat", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Private Function STATADateTest(ByVal sDateFormat As String) As Boolean
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'Replace double format characters with single format characters
    sDateFormat = Replace(sDateFormat, "dd", "d")
    sDateFormat = Replace(sDateFormat, "mm", "m")
    sDateFormat = Replace(sDateFormat, "hh", "h")
    sDateFormat = Replace(sDateFormat, "ss", "s")
    'Replace yyyy with y
    sDateFormat = Replace(sDateFormat, "yyyy", "y")
    'Replace all Date/Time Separators with "/"
    sDateFormat = Replace(sDateFormat, ":", "/")
    sDateFormat = Replace(sDateFormat, ".", "/")
    sDateFormat = Replace(sDateFormat, "-", "/")
    sDateFormat = Replace(sDateFormat, " ", "/")
    
    Select Case sDateFormat
    Case "d/m/y", "m/d/y", "y/m/d"
        STATADateTest = True
    Case Else
        STATADateTest = False
    End Select
    
Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "STATADateTest", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'---------------------------------------------------------------------
Private Sub STATAReplaceSection(ByVal nSTATAdoFileNumber As Integer, _
                                ByVal nQuestionsCounter As Integer)
'---------------------------------------------------------------------
Dim i As Integer
Dim asCodeType() As String
Dim sShortCode As String
Dim sType As String
Dim nLength As Integer
Dim sSTATAReplace As String

    On Error GoTo Errhandler

    'Write the Special Value Replace section to the STATA.do file (if required)
    If (gsSVMissing <> "") Or (gsSVNotApplicable <> "") Or (gsSVUnobtainable <> "") Then
        Print #nSTATAdoFileNumber,
        'Loop through the questions
        For i = 1 To nQuestionsCounter
            asCodeType = Split(gColSTATADetails.Item(i), "|")
            sShortCode = asCodeType(0)
            sType = asCodeType(1)
            If sType = "STRING" Then
                nLength = asCodeType(2)
            End If
            If gsSVMissing <> "" Then
                If sType = "NUMERIC" Then
                    'Mo 11/1/2006 Bug 2672
                    Select Case gsSVMissing
                    Case "-1"
                        sSTATAReplace = ".a"
                    Case "-2"
                        sSTATAReplace = ".b"
                    Case "-3"
                        sSTATAReplace = ".c"
                    Case "-4"
                        sSTATAReplace = ".d"
                    Case "-5"
                        sSTATAReplace = ".e"
                    Case "-6"
                        sSTATAReplace = ".f"
                    Case "-7"
                        sSTATAReplace = ".g"
                    Case "-8"
                        sSTATAReplace = ".h"
                    Case "-9"
                        sSTATAReplace = ".i"
                    End Select
                    'Mo 9/6/2006 Bug 2739
                    Print #nSTATAdoFileNumber, "replace " & sShortCode & Space((gnShortCodeLength + 1) - Len(sShortCode)) & "= " _
                        & sSTATAReplace & Space(7 - Len(sSTATAReplace)) & "if " & sShortCode & Space((gnShortCodeLength + 1) - Len(sShortCode)) _
                        & "== " & gsSVMissing & " ;"
                Else
                    'Padding with blanks removed at the request of the MRC 8/7/03
                    'Print #nSTATAdoFileNumber, "replace " & sShortCode & Space(9 - Len(sShortCode)) & "= """" " _
                    '    & "if " & sShortCode & Space(9 - Len(sShortCode)) _
                    '    & "== """ & gsSVMissing & Space(nLength - Len(gsSVMissing)) & """ ;"
                    'Mo 9/6/2006 Bug 2739
                    Print #nSTATAdoFileNumber, "replace " & sShortCode & Space((gnShortCodeLength + 1) - Len(sShortCode)) & "= ""MISS"" " _
                        & "if " & sShortCode & Space((gnShortCodeLength + 1) - Len(sShortCode)) _
                        & "== """ & gsSVMissing & """ ;"
                End If
            End If
            If gsSVUnobtainable <> "" Then
                If sType = "NUMERIC" Then
                    'Mo 11/1/2006 Bug 2672
                    Select Case gsSVUnobtainable
                    Case "-1"
                        sSTATAReplace = ".a"
                    Case "-2"
                        sSTATAReplace = ".b"
                    Case "-3"
                        sSTATAReplace = ".c"
                    Case "-4"
                        sSTATAReplace = ".d"
                    Case "-5"
                        sSTATAReplace = ".e"
                    Case "-6"
                        sSTATAReplace = ".f"
                    Case "-7"
                        sSTATAReplace = ".g"
                    Case "-8"
                        sSTATAReplace = ".h"
                    Case "-9"
                        sSTATAReplace = ".i"
                    End Select
                    'Mo 9/6/2006 Bug 2739
                    Print #nSTATAdoFileNumber, "replace " & sShortCode & Space((gnShortCodeLength + 1) - Len(sShortCode)) & "= " _
                        & sSTATAReplace & Space(7 - Len(sSTATAReplace)) & "if " & sShortCode & Space((gnShortCodeLength + 1) - Len(sShortCode)) _
                        & "== " & gsSVUnobtainable & " ;"
                Else
                    'Padding with blanks removed at the request of the MRC 8/7/03
                    'Print #nSTATAdoFileNumber, "replace " & sShortCode & Space(9 - Len(sShortCode)) & "= """" " _
                    '    & "if " & sShortCode & Space(9 - Len(sShortCode)) _
                    '    & "== """ & gsSVUnobtainable & Space(nLength - Len(gsSVUnobtainable)) & """ ;"
                    'Mo 9/6/2006 Bug 2739
                    Print #nSTATAdoFileNumber, "replace " & sShortCode & Space((gnShortCodeLength + 1) - Len(sShortCode)) & "= ""UNOB"" " _
                        & "if " & sShortCode & Space((gnShortCodeLength + 1) - Len(sShortCode)) _
                        & "== """ & gsSVUnobtainable & """ ;"
                End If
            End If
            If gsSVNotApplicable <> "" Then
                If sType = "NUMERIC" Then
                    'Mo 11/1/2006 Bug 2672
                    Select Case gsSVNotApplicable
                    Case "-1"
                        sSTATAReplace = ".a"
                    Case "-2"
                        sSTATAReplace = ".b"
                    Case "-3"
                        sSTATAReplace = ".c"
                    Case "-4"
                        sSTATAReplace = ".d"
                    Case "-5"
                        sSTATAReplace = ".e"
                    Case "-6"
                        sSTATAReplace = ".f"
                    Case "-7"
                        sSTATAReplace = ".g"
                    Case "-8"
                        sSTATAReplace = ".h"
                    Case "-9"
                        sSTATAReplace = ".i"
                    End Select
                    'Mo 9/6/2006 Bug 2739
                    Print #nSTATAdoFileNumber, "replace " & sShortCode & Space((gnShortCodeLength + 1) - Len(sShortCode)) & "= " _
                        & sSTATAReplace & Space(7 - Len(sSTATAReplace)) & "if " & sShortCode & Space((gnShortCodeLength + 1) - Len(sShortCode)) _
                        & "== " & gsSVNotApplicable & " ;"
                Else
                    'Padding with blanks removed at the request of the MRC 8/7/03
                    'Print #nSTATAdoFileNumber, "replace " & sShortCode & Space(9 - Len(sShortCode)) & "= """" " _
                    '    & "if " & sShortCode & Space(9 - Len(sShortCode)) _
                    '    & "== """ & gsSVNotApplicable & Space(nLength - Len(gsSVNotApplicable)) & """ ;"
                    'Mo 9/6/2006 Bug 2739
                    Print #nSTATAdoFileNumber, "replace " & sShortCode & Space((gnShortCodeLength + 1) - Len(sShortCode)) & "= ""NA""   " _
                        & "if " & sShortCode & Space((gnShortCodeLength + 1) - Len(sShortCode)) _
                        & "== """ & gsSVNotApplicable & """ ;"
                End If
            End If
        Next
    End If

Exit Sub
Errhandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "STATAReplaceSection", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'---------------------------------------------------------------------
Private Function GetSingleResponse(ByVal sClinicalTrialName As String, _
                                    ByVal sTrialSite As String, _
                                    ByVal lPersonId As Long, _
                                    ByVal sVisitCode As String, _
                                    ByVal nVisitCycle As Integer, _
                                    ByVal sCRFPageCode As String, _
                                    ByVal nCRFPageCycle As Integer, _
                                    ByVal sDataItemCode As String, _
                                    ByVal nRepeatNumber As Integer) As String
'---------------------------------------------------------------------
'Mo 3/11/2006 Bug 2834, new function added
'---------------------------------------------------------------------
Dim lClinicalTrialId As Long
Dim lVisitId As Long
Dim lCRFPageId As Long
Dim lDataItemId As Long
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sResponse As String

    On Error GoTo Errhandler

    lClinicalTrialId = TrialIdFromName(sClinicalTrialName)
    lVisitId = VisitIdFromCode(lClinicalTrialId, sVisitCode)
    lCRFPageId = CRFPageIdFromCode(lClinicalTrialId, sCRFPageCode)
    lDataItemId = DataItemIdFromCode(lClinicalTrialId, sDataItemCode)
    
    sSQL = "SELECT ResponseValue FROM DataItemResponse" _
        & " WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND TrialSite = '" & sTrialSite & "'" _
        & " AND PersonId = " & lPersonId _
        & " AND VisitId = " & lVisitId _
        & " AND VisitCycleNumber = " & nVisitCycle _
        & " AND CRFPageId = " & lCRFPageId _
        & " AND CRFPageCycleNumber = " & nCRFPageCycle _
        & " AND DataItemId = " & lDataItemId _
        & " AND RepeatNumber = " & nRepeatNumber
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 1 Then
        sResponse = RemoveNull(rsTemp!ResponseValue)
    Else
        sResponse = ""
    End If

    rsTemp.Close
    Set rsTemp = Nothing

    GetSingleResponse = sResponse

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetSingleResponse", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'Mo 2/4/2007 MRC15022007
'---------------------------------------------------------------------
Public Function ConstructOutputFileName() As String
'---------------------------------------------------------------------
Dim sOutputFileName As String

    On Error GoTo Errhandler

    If gsFileNamePath = "" Then
        'Use application Out Folder
        sOutputFileName = gsOUT_FOLDER_LOCATION
    Else
        'Use User Specified Path
        If Mid(gsFileNamePath, Len(gsFileNamePath), 1) = "\" Then
            sOutputFileName = gsFileNamePath
        Else
            sOutputFileName = gsFileNamePath & "\"
        End If
    End If

    If gsFileNameText = "" Then
        'Use StudyName
        sOutputFileName = sOutputFileName & TrialNameFromId(glSelectedTrialId)
    Else
        'Use User Specified Name
        sOutputFileName = sOutputFileName & gsFileNameText
    End If
    
    Select Case gsFileNameStamp
    Case "DATE"
        'Tag name with a Date Stamp
        sOutputFileName = sOutputFileName & Format(Now, "yyyymmdd")
    Case "DATETIME"
        'Tag name with a Date/Time Stamp
        sOutputFileName = sOutputFileName & Format(Now, "yyyymmddhhmmss")
    Case ""
        'No Date or Time Tag required
    End Select
 
    ConstructOutputFileName = sOutputFileName

Exit Function
Errhandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ConstructOutputFileName", "modQueryModule")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function
