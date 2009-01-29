VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTABExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TAB Delimited Export"
   ClientHeight    =   3405
   ClientLeft      =   8085
   ClientTop       =   4845
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   7335
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   615
         Left            =   5400
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdSelectExportFile 
         Caption         =   "Change Name/Location of Export File"
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2910
      End
      Begin VB.CommandButton cmdStartExport 
         Caption         =   "Start Export"
         Height          =   615
         Left            =   3480
         TabIndex        =   5
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.ComboBox CboTrialList 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   7455
   End
   Begin VB.TextBox txtExportMessage 
      Enabled         =   0   'False
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   1425
      Width           =   7470
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Export Progress:"
      Height          =   255
      Left            =   105
      TabIndex        =   2
      Top             =   1215
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Start by selecting a Trial. Then select Name/Location. Then press Start Export."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7410
   End
End
Attribute VB_Name = "frmTABExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmTABExport.frm
'   Author:         Mo Morris July 1997
'   Purpose:    Allows selection of study to be exported.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   1   PN 08/09/99 changed field Value to ItemValue in CreateCodeLookUps
'   2   PN 14/09/99    Updated code to conform to VB standards doc version 1.0
'                      Upgrade database access code from DAO to ADO
'   3   PN 15/09/99    Changed call to ADODBConnection() to MacroADODBConnection()
'   PN  21/09/99        Added moFormIdleWatch object to handle system idle timer resets
'   WillC   10/11/99    Added the error handlers
'   Mo Morris 7/12/99   Data Control and Date Grid changed to ADO enabled controls
'   Mo 13/12/99     Id's from integer to Long
'   WillC 3/2/00 SR3504 replaced Datagrid and datacontrol
'   TA 08/05/2000   subclassing removed
'   Mo 5/12/00      Form.StartUpPosition set to CenterScreen
'   DPH 17/10/2001 Added FolderExistence routine calls to create missing folders
'----------------------------------------------------------------------------------------'

Option Explicit
Option Compare Binary
Option Base 0

Private Const msDELIMMITING_CHAR = vbTab

Private msExportFile As String
Private msCodeLookUpFile As String
Private msDataItemLookUpFile As String

Private mnTABFileNumber As Integer
Private mnCLUFileNumber As Integer
Private mnDLUFileNumber As Integer

Private mnSelTrialId As Long
Private msSelTrialName As String
Private msSelVersionId As String

Private msHeaderString As String
Private msPatientDataRecord As String
Private mnNumDataItems As Integer
Private mcolRecordHeaderPositions As Collection

Private DataItemResponses() As Variant
Private msReasonForAborting As String


'---------------------------------------------------------------------
Private Sub cmdClose_Click()
'---------------------------------------------------------------------
' Unload the form
'---------------------------------------------------------------------

    Unload Me
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdSelectExportFile_Click()
'---------------------------------------------------------------------

    ' PN 14/09/99 catch cancel error with error label not resume next
    On Error GoTo CancelSelected
    With CommonDialog1
        .CancelError = True
        .flags = cdlOFNOverwritePrompt & cdlOFNHideReadOnly
        .FileName = msExportFile
        .ShowOpen
    
        msExportFile = .FileName
    End With
    
    'Create a file name for the Code Look Up file (clu) and the DataItem Look Up file (dlu)
    msCodeLookUpFile = Mid(msExportFile, 1, Len(msExportFile) - 3) & "clu"
    msDataItemLookUpFile = Mid(msExportFile, 1, Len(msExportFile) - 3) & "dlu"
    
    cmdStartExport.Enabled = True

CancelSelected:

End Sub

'---------------------------------------------------------------------
Private Sub cboTrialList_Click()
'---------------------------------------------------------------------
' enable the generate button if something is chosen
'---------------------------------------------------------------------
On Error GoTo ErrHandler
    
    If CboTrialList.ListIndex = -1 Then
        cmdSelectExportFile.Enabled = False
        cmdStartExport.Enabled = False
    Else
        cmdSelectExportFile.Enabled = True
        msSelTrialName = Trim(CboTrialList.Text)
        mnSelTrialId = CboTrialList.ItemData(CboTrialList.ListIndex)
    End If

    msExportFile = gsOUT_FOLDER_LOCATION & msSelTrialName & "_"
'    msExportFile = gsOUT_FOLDER_LOCATION & "\" & msSelTrialName & "_"
    msExportFile = msExportFile & CStr(mnSelTrialId) & "_" & msSelVersionId & ".tab"

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cboTrialList_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

    
End Sub

'---------------------------------------------------------------------
Private Sub CboTrialList_Change()
'---------------------------------------------------------------------
' enable the generate button if something is chosen
'---------------------------------------------------------------------
On Error GoTo ErrHandler
    
    If CboTrialList.ListIndex = -1 Then
        cmdSelectExportFile.Enabled = False
        cmdStartExport.Enabled = False
    Else
        cmdSelectExportFile.Enabled = True
        msSelTrialName = Trim(CboTrialList.Text)
        mnSelTrialId = CboTrialList.ItemData(CboTrialList.ListIndex)
    End If

    msExportFile = gsOUT_FOLDER_LOCATION & msSelTrialName & "_"
'    msExportFile = gsOUT_FOLDER_LOCATION & "\" & msSelTrialName & "_"
    msExportFile = msExportFile & CStr(mnSelTrialId) & "_" & msSelVersionId & ".tab"

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "CboTrialList_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Sub


'---------------------------------------------------------------------
Private Sub cmdStartExport_Click()
'---------------------------------------------------------------------
Dim sOutput As String
Dim bIsError As Boolean
On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass
    
    ' DPH 17/10/2001 Make sure folder exists before opening
    If FolderExistence(msExportFile) Then
        'Open the TAB file
        mnTABFileNumber = FreeFile
        Open msExportFile For Output As #mnTABFileNumber
        
        'Open the DataItem LookUp file
        mnDLUFileNumber = FreeFile
        Open msDataItemLookUpFile For Output As #mnDLUFileNumber
        sOutput = "Visit/Form/DataItem" & msDELIMMITING_CHAR & "Description"
        sOutput = sOutput & msDELIMMITING_CHAR & "Type"
        Print #mnDLUFileNumber, sOutput
        
        'Open the Code LookUp file
        mnCLUFileNumber = FreeFile
        Open msCodeLookUpFile For Output As #mnCLUFileNumber
        sOutput = "Visit/Form/DataItem" & msDELIMMITING_CHAR & "ValueCode"
        sOutput = sOutput & msDELIMMITING_CHAR & "Value"
        Print #mnCLUFileNumber, sOutput
        
        'initialize the reason for aborting string
        msReasonForAborting = ""
        
        'call CreateHeaderRecord to setup the header record for the TAB file
        Call CreateHeaderRecord
        
        ' PN 14/09/99 removed Goto and added handling of error
        If msReasonForAborting = vbNullString Then
            'call CollectPatientData to collect the patient data and write it to the TAB file
            Call CollectPatientData
            
            If msReasonForAborting = vbNullString Then
                bIsError = False
            
            Else
                bIsError = True
                
            End If
        
        Else
            bIsError = True
        
        End If
        
        'Close the files
        Close #mnTABFileNumber
        Close #mnDLUFileNumber
        Close #mnCLUFileNumber
    
        If bIsError Then
            ' clean up after the error
            txtExportMessage.Text = msReasonForAborting
            Kill msExportFile
            Kill msCodeLookUpFile
            Kill msDataItemLookUpFile
    
        Else
            txtExportMessage.Text = "Export Completed"
        
        End If
    Else
        txtExportMessage.Text = "Export failed as could not create file"
    End If
    
    cmdStartExport.Enabled = False
    cmdSelectExportFile.Enabled = False
    Screen.MousePointer = vbDefault
    
    DoEvents
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdStartExport_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTrialList As ADODB.Recordset

    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon
    
    cmdSelectExportFile.Enabled = False
    cmdStartExport.Enabled = False
    
    sSQL = "SELECT ClinicalTrial.ClinicalTrialId, ClinicalTrial.ClinicalTrialName,"
    sSQL = sSQL & " ClinicalTrial.ClinicalTrialDescription, StudyDefinition.VersionId "
    sSQL = sSQL & " FROM ClinicalTrial, StudyDefinition "
    sSQL = sSQL & " WHERE ClinicalTrial.ClinicalTrialId = StudyDefinition.ClinicalTrialId"

    Set rsTrialList = New ADODB.Recordset
    
    rsTrialList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    CboTrialList.Clear
    Do Until rsTrialList.EOF
        CboTrialList.AddItem rsTrialList!ClinicalTrialName ' & " - " & rsTrialList!ClinicalTrialDescription
        CboTrialList.ItemData(CboTrialList.NewIndex) = rsTrialList!ClinicalTrialId
        rsTrialList.MoveNext
    Loop
    
    With CommonDialog1
        .DialogTitle = "TAB Export File Selection"
        .InitDir = gsOUT_FOLDER_LOCATION
        .DefaultExt = "tab"
        .Filter = "TAB Export files (*.tab)|*.tab"
    End With
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "Form_Load")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
    
End Sub


'---------------------------------------------------------------------
Private Sub CreateHeaderRecord()
'---------------------------------------------------------------------
'This routine creates the comma separated header record for the main export file
'and includes column headers for most of the dataitems in file TrialSubject followed by
'column headers for every data item on every form within every visit for the
'selected trial
'This routine additionally assess the number of data items (mnNumDataItems) and creates
'a the collection (mcolRecordHeaderPositions) which will be used as a lookup to which position
'a particular data item should be in. The lookup is based on a concatenated key of
'VisitId/CRFPageId/CRFElementId
'Mo Morris 30/8/01 Db Audit (Dropped fields and unneccessary fields from table
'   TrialSubject removed from Header Record)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim nIndex As Integer
Dim sColumnHeader As String
Dim sLookUpKey As String
Dim sDataType As String

On Error GoTo ErrHandler
    'Initialize header string
    msHeaderString = vbNullString
    'Initialize mcolRecordHeaderPositions collection
    Set mcolRecordHeaderPositions = New Collection
    
    'Create TrialSubject column headers for the required data items in table TrialSubject.
    'TrialSite (field 1), PersonId (field 2), DateOfBirth (field 3), Gender (field 4)
    'LocalIdentifier1 (field 5), LocalIdentifier2 (field 6), SubjectGender (field 12).
    'Fields 0,7,8,9,10,11 and 13 are not required.
    'Note that TrialSite and PersonId have to be concatenated into a unique value using a "/"
    sSQL = "SELECT * FROM TrialSubject"
    
    ' PN 14/09/99 changed db access to ADO from DAO
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With rsTemp
        For nIndex = 0 To 13
            Select Case nIndex
            Case 1
                msHeaderString = msHeaderString & .Fields(nIndex).Name & "/"
            Case 2, 3, 4, 5, 6, 12
                msHeaderString = msHeaderString & .Fields(nIndex).Name & msDELIMMITING_CHAR
            End Select
        Next nIndex
        .Close
    End With
    Set rsTemp = Nothing
    
    'strip off the last msDELIMMITING_CHAR
    msHeaderString = Mid$(msHeaderString, 1, (Len(msHeaderString) - 1))
    
    'Get Column headers for all data items on all forms for all visits within trial
    'Additionally link-in DataItemResponseData so that every VisitCycleNumber and CRFPageCycleNumber
    'is catered for in the Column headers
    sSQL = "SELECT DISTINCT StudyVisit.VisitId, StudyVisit.VisitCode, StudyVisit.VisitOrder," _
        & " CRFPage.CRFPageId, CRFPage.CRFPageCode, CRFPage.CRFPageOrder," _
        & " CRFElement.CRFElementId, DataItem.DataItemCode, CRFElement.FieldOrder," _
        & " DataItem.DataItemId, DataItem.DataItemName, DataItem.DataType," _
        & " DataItemResponse.VisitCycleNumber, DataItemResponse.CRFPageCycleNumber" _
        & " FROM StudyVisitCRFPage,StudyVisit, CRFPage, DataItem, CRFElement, DataItemResponse " _
        & " WHERE StudyVisitCRFPage.ClinicalTrialId = StudyVisit.ClinicalTrialId " _
        & " AND StudyVisitCRFPage.ClinicalTrialId = CRFPage.ClinicalTrialId " _
        & " AND StudyVisitCRFPage.ClinicalTrialId = CRFElement.ClinicalTrialId " _
        & " AND StudyVisitCRFPage.ClinicalTrialId = DataItem.ClinicalTrialId " _
        & " And StudyVisitCRFPage.VisitId = StudyVisit.VisitId " _
        & " AND StudyVisitCRFPage.CRFPageId=CRFPage.CRFPageId " _
        & " AND StudyVisitCRFPage.CRFPageId = CRFElement.CRFPageId " _
        & " AND CRFElement.DataItemId = DataItem.DataItemId " _
        & " AND DataItem.ClinicalTrialId = DataItemResponse.ClinicalTrialId " _
        & " AND DataItem.DataItemId = DataItemResponse.DataItemId " _
        & " AND StudyVisit.ClinicalTrialId = " & mnSelTrialId _
        & " ORDER BY  StudyVisit.VisitOrder, DataItemResponse.VisitCycleNumber," _
        & " CRFPage.CRFPageOrder, DataItemResponse.CRFPageCycleNumber, CRFElement.FieldOrder"
        
    ' PN 14/09/99 changed db access to ADO from DAO
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With rsTemp
    
        mnNumDataItems = 0
        'read through the record set creating a unique column header for each data item
        'the column headers being a concatenation of the VisitCode, CRFPageCode and the DataItemCode
        Do While Not .EOF
            sColumnHeader = rsTemp![VisitCode] & "/" & rsTemp![VisitCycleNumber] & "/" _
                & rsTemp![CRFPageCode] & "/" & rsTemp![CRFPageCycleNumber] & "/" & rsTemp![DataItemCode]
            msHeaderString = msHeaderString & msDELIMMITING_CHAR & sColumnHeader
            mnNumDataItems = mnNumDataItems + 1
            txtExportMessage.Text = "Preparing data item " & mnNumDataItems
            DoEvents
            'Add an entry to the collection mcolRecordHeaderPositions that records a data item's field order/position
            'with the look up key being a concatenation of VisitId/VisitCycleNumber/CRFPageId/CRFPageCycleNumber/CRFElemetId
            sLookUpKey = rsTemp![VisitId] & "/" & rsTemp![VisitCycleNumber] & "/" _
                & rsTemp![CRFPageId] & "/" & rsTemp![CRFPageCycleNumber] & "/" & rsTemp![CRFElementId]
            mcolRecordHeaderPositions.Add mnNumDataItems, sLookUpKey
            'Add an entry to the DataItem LookUp file having first changed the datatype into its enumeration string
            Select Case rsTemp![DataType]
            Case DataType.Category
                sDataType = "Category"
                'If data item is of type category call CreateCodeLookUps to store the dataitem codes and values
                Call CreateCodeLookUps(sColumnHeader, rsTemp![DataItemId])
            Case DataType.Date
                sDataType = "Date"
            Case DataType.IntegerData
                sDataType = "Integer"
        '    Case DataType.Laboratory
        '        sDataType = "Laboratory"
            Case DataType.Multimedia
                sDataType = "Multimedia"
            Case DataType.Real
                sDataType = "Real"
            Case DataType.Text
                sDataType = "Text"
            End Select
            Print #mnDLUFileNumber, sColumnHeader & msDELIMMITING_CHAR & rsTemp![DataItemName] & msDELIMMITING_CHAR & sDataType
            .MoveNext
        Loop
        .Close
    End With
    Set rsTemp = Nothing
    
    'Check that the Trial has at least one data item in it
    If mnNumDataItems = 0 Then
        msReasonForAborting = "Trial contains no data items. Export aborted"
        Exit Sub
    End If
    
    'Write the TAB header record to file
    Print #mnTABFileNumber, msHeaderString
    
    'Dimension the array into which data item responses will be temporarily stored
    ReDim DataItemResponses(1 To mnNumDataItems)
    
Exit Sub

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "CreateHeaderRecord")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
        
End Sub

'---------------------------------------------------------------------
Private Sub CollectPatientData()
'---------------------------------------------------------------------
'Select all records from DataItemResponse for the specified trial.
'Read through the records
'for each new patient
'   initialise a new record
'   get the required information from TrialSubjects
'   Write patient record to file
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsResponseData As ADODB.Recordset
Dim sLookUpKey As String
Dim sTrialSitePersonId As String
Dim sPreviousSitePerson As String
Dim nDataItemLocation As Integer
Dim nCount As Long

    sSQL = "SELECT DataItemResponse.TrialSite, DataItemResponse.PersonId, DataItemResponse.VisitId," _
        & " DataItemResponse.CRFPageId, DataItemResponse.CRFElementId, DataItemResponse.ResponseValue," _
        & " DataItemResponse.ValueCode, DataItemResponse.VisitCycleNumber, DataItemResponse.CRFPageCycleNumber, DataItem.DataType " _
        & " FROM DataItemResponse, DataItem " _
        & " WHERE DataItemResponse.ClinicalTrialId = DataItem.ClinicalTrialId " _
        & " AND DataItemResponse.DataItemId = DataItem.DataItemId " _
        & " AND DataItemResponse.ClinicalTrialId = " & mnSelTrialId _
        & " ORDER BY TrialSite, PersonId"
        
    ' PN 14/09/99 changed db access to ADO from DAO
    Set rsResponseData = New ADODB.Recordset
    rsResponseData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    nCount = 0
    sPreviousSitePerson = ""
    
    ' PN 14/09/99 moved start of error capture
    On Error GoTo ErrorHandler
    
    Do While Not rsResponseData.EOF
        nCount = nCount + 1
        txtExportMessage.Text = "Collecting data item " & nCount
        DoEvents
        sTrialSitePersonId = rsResponseData![TrialSite] & "/" & rsResponseData![PersonID]
        'Check for a new patient
        If sPreviousSitePerson <> sTrialSitePersonId Then
            'if its not the first patient write the previous patients data to file
            If sPreviousSitePerson <> "" Then
                WritePatientRecordToFile
            End If
            sPreviousSitePerson = sTrialSitePersonId
            'call GetTrialSubjectData to initialise the msPatientDataRecord
            GetTrialSubjectData rsResponseData![TrialSite], rsResponseData![PersonID]
            'Initialise the array into which the data item responses will be placed
            InitializeResponseArray
        End If
        'Create a lookup key into mcolRecordHeaderPositions for the current data item
        sLookUpKey = rsResponseData![VisitId] & "/" & rsResponseData![VisitCycleNumber] & "/" _
            & rsResponseData![CRFPageId] & "/" & rsResponseData![CRFPageCycleNumber] & "/" & rsResponseData![CRFElementId]
        'Get the data item location from mcolRecordHeaderPositions for the current data item
        nDataItemLocation = mcolRecordHeaderPositions.Item(sLookUpKey)
        'store the result in DataItemResponses array
        'If data item type is category store the ValueCode otherwise store the ResponseValue
        If rsResponseData![DataType] = DataType.Category Then
            DataItemResponses(nDataItemLocation) = rsResponseData![ValueCode]
        Else
            DataItemResponses(nDataItemLocation) = rsResponseData![ResponseValue]
        End If
        'Get next record
        rsResponseData.MoveNext
    Loop
    rsResponseData.Close
    Set rsResponseData = Nothing
    
    'Check that there has been some patient data before writting the last patients data to file
    If sPreviousSitePerson <> "" Then
        WritePatientRecordToFile
    Else
        msReasonForAborting = "There is no patient data on this file. Export aborted"
    End If
    
    Exit Sub
    
ErrorHandler:
    If Err.Number = 5 Then
        ' the invalid procedure call or argument error
        MsgBox ("It would appear that the Study definition has been changed since some of the patient data was entered.")
    Else
        MsgBox ("Error number " & Err.Number & " " & Err.Description)
    End If
    msReasonForAborting = "Export aborted"
    
End Sub

'---------------------------------------------------------------------
Private Sub GetTrialSubjectData(ByVal vTrialSite As String, ByVal vPersonId As Integer)
'---------------------------------------------------------------------
'This routine creates the initial part of msPatientDataRecord, with the
'required information from TrialSubject
'Mo Morris 30/8/01 Db Audit (Dropped fields and unneccessary fields from table
'   TrialSubject removed from Header Record)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTrialSubject As ADODB.Recordset

    On Error GoTo ErrHandler
    
    'get the TrialSubject details for the current patient
    sSQL = "SELECT * FROM TrialSubject" _
        & " WHERE ClinicalTrialId = " & mnSelTrialId _
        & " AND TrialSite = '" & vTrialSite & "'" _
        & " AND PersonId = " & vPersonId
    
    ' PN 14/09/99 changed db access to ADO from DAO
    Set rsTrialSubject = New ADODB.Recordset
    rsTrialSubject.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    'ASH 12/10/2001 changed SubjectIGender to SubjectGender
    msPatientDataRecord = vTrialSite & "/" & vPersonId & msDELIMMITING_CHAR _
        & rsTrialSubject![DateOfBirth] & msDELIMMITING_CHAR & rsTrialSubject![Gender] & msDELIMMITING_CHAR _
        & rsTrialSubject![LocalIdentifier1] & msDELIMMITING_CHAR & rsTrialSubject![LocalIdentifier2] _
        & rsTrialSubject![SubjectGender]
    rsTrialSubject.Close
    Set rsTrialSubject = Nothing

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "GetTrialSubjectData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
    
End Sub

'---------------------------------------------------------------------
Private Sub InitializeResponseArray()
'---------------------------------------------------------------------
Dim nIndex As Integer

    On Error GoTo ErrHandler
    
    For nIndex = 1 To mnNumDataItems
       DataItemResponses(nIndex) = vbNullString
    Next nIndex
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "InitializeResponseArray")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub WritePatientRecordToFile()
'---------------------------------------------------------------------
Dim nIndex As Integer

    On Error GoTo ErrHandler
    
    'read through the array DataItemResponses adding the response data and commas to the
    'already created msPatientDataRecord
    For nIndex = 1 To mnNumDataItems
        msPatientDataRecord = msPatientDataRecord & msDELIMMITING_CHAR & DataItemResponses(nIndex)
    Next nIndex
    
    'Write the patient data record to file
    Print #mnTABFileNumber, msPatientDataRecord
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "WritePatientRecordToFile")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub CreateCodeLookUps(ByVal vVistFormDataItem As String, ByVal vDataItemId As Long)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sOutput As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM ValueData " _
        & " WHERE ClinicalTrialID = " & mnSelTrialId _
        & " AND DataItemId = " & vDataItemId
        
    ' PN 14/09/99 changed db access to ADO from DAO
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        
    Do While Not rsTemp.EOF
        ' PN 08/09/99 changed field Value to ItemValue
        sOutput = vVistFormDataItem & msDELIMMITING_CHAR & rsTemp![ValueCode]
        sOutput = sOutput & msDELIMMITING_CHAR & rsTemp![ItemValue]
        Print #mnCLUFileNumber, sOutput
        rsTemp.MoveNext
    Loop
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "CreateCodeLookUps")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
End Sub
