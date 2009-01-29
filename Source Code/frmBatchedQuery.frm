VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBatchedQuery 
   Caption         =   "MACRO Batched Query"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   9885
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtProgress 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2200
      TabIndex        =   11
      Top             =   3600
      Width           =   7600
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   100
      TabIndex        =   9
      Top             =   3100
      Width           =   1300
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   100
      TabIndex        =   7
      Top             =   3600
      Width           =   1300
   End
   Begin VB.CommandButton cmdRemoveQuery 
      Caption         =   "Remove Query"
      Height          =   375
      Left            =   100
      TabIndex        =   1
      Top             =   600
      Width           =   1300
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Up"
      Height          =   375
      Left            =   100
      TabIndex        =   2
      Top             =   1100
      Width           =   600
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Down"
      Height          =   375
      Left            =   800
      TabIndex        =   3
      Top             =   1100
      Width           =   600
   End
   Begin VB.ListBox lstBatchedQueries 
      Height          =   3375
      ItemData        =   "frmBatchedQuery.frx":0000
      Left            =   1500
      List            =   "frmBatchedQuery.frx":0007
      TabIndex        =   8
      Top             =   120
      Width           =   8300
   End
   Begin VB.CommandButton cmdAddQuery 
      Caption         =   "Add Query"
      Height          =   375
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   1300
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Save As"
      Height          =   375
      Left            =   100
      TabIndex        =   5
      Top             =   2100
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   100
      TabIndex        =   4
      Top             =   1600
      Width           =   1300
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   375
      Left            =   100
      TabIndex        =   6
      Top             =   2600
      Width           =   1300
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblProgress 
      Caption         =   "Progress:"
      Height          =   195
      Left            =   1500
      TabIndex        =   10
      Top             =   3600
      Width           =   675
   End
End
Attribute VB_Name = "frmBatchedQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         frmBatchedQuery
' Copyright:    InferMed Ltd. 2000. All Rights Reserved
' Author:       Mo Morris, April 2007
' Purpose:      Contains controls for setting up/changing a Batch Query
'----------------------------------------------------------------------------------------'
'   Revisions:
'   Mo  2/4/2007    MRC15022007 - Query Module Batch Facilities
'                   This form added
'----------------------------------------------------------------------------------------'

Option Explicit

Private msCurrentBatchQueryPathName As String
Private mbBatchedQueryRunning As Boolean

'--------------------------------------------------------------------
Private Sub cmdAddQuery_Click()
'--------------------------------------------------------------------
Dim sQueryPathName As String
Dim i As Integer
Dim bAlreadyExists As Boolean

    On Error GoTo CancelOpen
    With CommonDialog1
        .DialogTitle = "Open MACRO Query"
        .InitDir = gsOUT_FOLDER_LOCATION
        .DefaultExt = "txt"
        .Filter = "Text file (*.txt)|*.txt"
        .CancelError = True
        .ShowOpen
  
        sQueryPathName = .FileName
    End With
    
    'Check that it is a valid Macro Query that has been selected
    If frmMenu.OpenQuery(sQueryPathName) Then
        'Check that this Query has not already been added to this Batched Query
        bAlreadyExists = False
        For i = 0 To lstBatchedQueries.ListCount - 1
            If lstBatchedQueries.List(i) = sQueryPathName Then
                Call DialogInformation(sQueryPathName & vbNewLine & "is already part of this Batched Query.", "MACRO Batched Query error")
                bAlreadyExists = True
                Exit For
            End If
        Next i
        If Not bAlreadyExists Then
            'Add the selected Query to the Batched Query List
            lstBatchedQueries.AddItem sQueryPathName
            'unselect any selected entries in lstBatchedQueries
            lstBatchedQueries.ListIndex = -1
            'enable Save and Save as command buttons
            cmdSave.Enabled = True
            cmdSaveAs.Enabled = True
            'enable the Run command button
            cmdRun.Enabled = True
        End If
    End If
    
    Call frmMenu.ClearQueryAndReset
    'Unselect the previously selected study
    frmMenu.cboStudies.ListIndex = -1
    
    'flag batch query as changed
    gbBatchQueryChanged = True
    
CancelOpen:

End Sub

'--------------------------------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------

    gbCancelled = True

End Sub

'--------------------------------------------------------------------
Private Sub cmdDown_Click()
'--------------------------------------------------------------------
Dim sThisQuery As String
Dim nSelectedIndex As Integer

    'Store the IndexId of the Query to be moved
    nSelectedIndex = lstBatchedQueries.ListIndex
    'Store the Text of the Query to be moved
    sThisQuery = lstBatchedQueries.Text
    'Copy the Query from the IndexId below into the currently selected IndexId
    lstBatchedQueries.List(nSelectedIndex) = lstBatchedQueries.List(nSelectedIndex + 1)
    'Copy sThisQuery into the location below
    lstBatchedQueries.List(nSelectedIndex + 1) = sThisQuery
    'Make sThisQuery in its new location the currently selected entry
    lstBatchedQueries.Selected(nSelectedIndex + 1) = True
    'if nSelectedIndex - 1 is not visible the control will scroll automatically
    'lstBatchedQueries_Click will have been triggered so that cmdUp/Down are correctly enable/disabled
    
    'flag batch query as changed
    gbBatchQueryChanged = True
    
    'enable Save and Save as command buttons
    cmdSave.Enabled = True
    cmdSaveAs.Enabled = True

End Sub

'--------------------------------------------------------------------
Private Sub cmdExit_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'check that the current query does not need saving
    Call SaveCheck
    
    Unload Me

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdExit_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cmdRemoveQuery_Click()
'--------------------------------------------------------------------

    'remove the selected entry
    lstBatchedQueries.RemoveItem (lstBatchedQueries.ListIndex)
    
    'disable Remove, Up, Down buttons
    cmdRemoveQuery.Enabled = False
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    
    'flag batch query as changed
    gbBatchQueryChanged = True
    
    If lstBatchedQueries.ListCount > 0 Then
        'enable Save and Save as command buttons
        cmdSave.Enabled = True
        cmdSaveAs.Enabled = True
        'enable the Run command button
        cmdRun.Enabled = True
    Else
        'can't Save a Batch Query with no queries
        cmdSave.Enabled = False
        cmdSaveAs.Enabled = False
        'disable the Run command button
        cmdRun.Enabled = False
    End If

End Sub

'--------------------------------------------------------------------
Private Sub cmdRun_Click()
'--------------------------------------------------------------------
Dim i As Integer
Dim sQueryPathName As String
Dim sQueryName As String

    On Error GoTo Errhandler
    
    Call HourglassOn
    
    Call DisableRunDisplay
    
    Call frmMenu.ClearQueryAndReset
    'Unselect any previously selected study
    frmMenu.cboStudies.ListIndex = -1
    
    'Loop through the queries and run them one at a time
    For i = 0 To lstBatchedQueries.ListCount - 1
        'extract query from list of queries
        sQueryPathName = lstBatchedQueries.List(i)
        sQueryName = StripFileNameFromPath(sQueryPathName)
        'Check that the individual query exists
        If Not FileExists(sQueryPathName) Then
            Call DisplayProgress("Query number " & (i + 1) & " - " & sQueryName & " no longer exists and could not be run.")
            'wait for 5 seconds
            Sleep 5000
        Else
            Call DisplayProgress("Query number " & (i + 1) & " - " & sQueryName & " being processed.")
            'open the query
            Call frmMenu.OpenQuery(sQueryPathName)
            
            'a call to FolderExistence will create any folders that do not exist
            Call FolderExistence(gsFileNamePath & "\")
    
            If gbCancelled Then
                Call HourglassOff
                Exit For
            End If
            
            Call frmMenu.QueryDB
            
            If gbCancelled Then
                Call HourglassOff
                Exit For
            End If
            
            'check for response data
            If frmMenu.mrsData.RecordCount = 0 Then
                Call DisplayProgress("Query number " & (i + 1) & " - " & sQueryName & " has returned no results.")
                'wait for 5 seconds
                Sleep 5000
            Else
                frmMenu.optDoNotDisplayOutput.Value = True
                If gbUseShortCodes Then
                    Set gColQuestionCodes = New Collection
                End If
                
                Call frmMenu.PrepareOutPut
                Call frmMenu.LoadOutPut(False)
                
                Select Case gnOutPutType
                Case eOutPutType.CSV
                    Call OutputToCSV(False)
                Case eOutPutType.Access
                    Call OutputToAccess(False)
                Case eOutPutType.SAS, eOutPutType.SASColons
                    Call OutputToSAS(False)
                Case eOutPutType.STATA
                    Call OutputToSTATA("Float", False)
                Case eOutPutType.MACROBD
                    Call OutputToMACROBD(False)
                Case eOutPutType.STATAStandardDates
                    Call OutputToSTATA("Standard", False)
                End Select
            End If
    
            Call frmMenu.ClearQueryAndReset
            'Unselect the previously selected study
            frmMenu.cboStudies.ListIndex = -1
            
            'wait for 1.5 seconds
            Sleep 1500
        End If
    Next
    
    If gbCancelled Then
        Call frmMenu.ClearQueryAndReset
        'Unselect the previously selected study
        frmMenu.cboStudies.ListIndex = -1
        Call DisplayProgress("Running of Batched Query Cancelled.")
    Else
        Call DisplayProgress("All Queries processed.")
    End If
    
    Call EnableRunDisplay

    Call HourglassOff
  
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdRun_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cmdSave_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'if the batch query has not been named then a call to Save As is required
    If gbBatchQuerySaved = False Then
        cmdSaveAs_Click
        Exit Sub
    End If

    Call SaveBatchQuery(msCurrentBatchQueryPathName)

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdSave_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cmdSaveAs_Click()
'--------------------------------------------------------------------
Dim sBatchQueryName As String
Dim sBatchQueryPathName As String

    'Check for the current batch query already having a name
    If gbBatchQuerySaved = False Then
        sBatchQueryName = "MBQ " & " (" & Format(Now, "yyyy mm dd hh mm") & ").txt"
        sBatchQueryPathName = gsOUT_FOLDER_LOCATION & sBatchQueryName
    Else
        sBatchQueryPathName = msCurrentBatchQueryPathName
    End If
    
    On Error GoTo CancelSaveAs
    With CommonDialog1
        .DialogTitle = "Save MACRO Batch Query As"
        .CancelError = True
        .Filter = "Text file (*.txt)|*.txt"
        .DefaultExt = "txt"
        .Flags = cdlOFNOverwritePrompt
        .FileName = sBatchQueryPathName
        .ShowSave
  
        sBatchQueryPathName = .FileName
    End With
    
    'save the batch query
    Call SaveBatchQuery(sBatchQueryPathName)
    
    'Store the name of the saved batch query
    msCurrentBatchQueryPathName = sBatchQueryPathName
    
    'set Batch Query has been saved with a name flag
    gbBatchQuerySaved = True
    
    'set changed flag false
    gbBatchQueryChanged = False

CancelSaveAs:

End Sub

'--------------------------------------------------------------------
Private Sub cmdUp_Click()
'--------------------------------------------------------------------
Dim sThisQuery As String
Dim nSelectedIndex As Integer

    'Store the IndexId of the Query to be moved
    nSelectedIndex = lstBatchedQueries.ListIndex
    'Store the Text of the Query to be moved
    sThisQuery = lstBatchedQueries.Text
    'Copy the Query from the IndexId above into the currently selected IndexId
    lstBatchedQueries.List(nSelectedIndex) = lstBatchedQueries.List(nSelectedIndex - 1)
    'Copy sThisQuery into the location above
    lstBatchedQueries.List(nSelectedIndex - 1) = sThisQuery
    'Make sThisQuery in its new location the currently selected entry
    lstBatchedQueries.Selected(nSelectedIndex - 1) = True
    'if nSelectedIndex - 1 is not visible the control will scroll automatically
    'lstBatchedQueries_Click will have been triggered so that cmdUp/Down are correctly enable/disabled
    
    'flag batch query as changed
    gbBatchQueryChanged = True
    
    'enable Save and Save as command buttons
    cmdSave.Enabled = True
    cmdSaveAs.Enabled = True

End Sub

'--------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------

    'Clear list of Batched Queries
    lstBatchedQueries.Clear
    
    cmdAddQuery.Enabled = True
    cmdSave.Enabled = False
    cmdSaveAs.Enabled = False
    cmdRemoveQuery.Enabled = False
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    cmdRun.Enabled = False
    cmdCancel.Enabled = False
    cmdExit.Enabled = True

End Sub

'--------------------------------------------------------------------
Private Sub Form_Resize()
'--------------------------------------------------------------------

    'check that form has not been minimised
    If Me.WindowState = vbMinimized Then Exit Sub
     
    If Me.Width < 10000 Then Me.Width = 10000
    
    If Me.Height < 4450 Then Me.Height = 4450
    
    lstBatchedQueries.Width = (Me.ScaleWidth - (lstBatchedQueries.Left + 100))
    lstBatchedQueries.Height = (Me.ScaleHeight - 600)
    
    lblProgress.Top = lstBatchedQueries.Height + 200
    txtProgress.Top = lblProgress.Top
    txtProgress.Width = lstBatchedQueries.Width - 700

End Sub

'--------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'--------------------------------------------------------------------

    'Check for a Batched Query being run
    If mbBatchedQueryRunning Then
        Cancel = 1
        Exit Sub
    End If

    'check that the current query does not need saving
    Call SaveCheck
    
    gbBatchQueryMode = False

End Sub

'--------------------------------------------------------------------
Private Sub lstBatchedQueries_Click()
'--------------------------------------------------------------------
    
    If lstBatchedQueries.ListCount > 1 Then
        If lstBatchedQueries.ListIndex = -1 Then
            cmdRemoveQuery.Enabled = False
            'nothing selected, disable Up and Down buttons
            cmdUp.Enabled = False
            cmdDown.Enabled = False
        ElseIf lstBatchedQueries.ListIndex = 0 Then
            cmdRemoveQuery.Enabled = True
            'only enable the down button
            cmdUp.Enabled = False
            cmdDown.Enabled = True
        ElseIf lstBatchedQueries.ListIndex = lstBatchedQueries.ListCount - 1 Then
            cmdRemoveQuery.Enabled = True
            'its the last entry, only enable the up button
            cmdUp.Enabled = True
            cmdDown.Enabled = False
        Else
            cmdRemoveQuery.Enabled = True
            'enable both buttons
            cmdUp.Enabled = True
            cmdDown.Enabled = True
        End If
    End If

End Sub

'---------------------------------------------------------------------
Private Sub SaveCheck()
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'Check for there being something to save
    'and prompt the user as to wether they want to save it
    If gbBatchQueryChanged = True Then
        If DialogQuestion("Changes have been made to the current Batched Query" _
        & vbNewLine & "Do you want to save the changes?", "Save Batched Query Changes") <> vbYes Then
            'set gbBatchQueryChanged to false so that this question is not asked again
            gbBatchQueryChanged = False
            'User does not want to save changes
            Exit Sub
        End If
    End If

    'If the query has not been named then a call to Save As is required
    'But only do this if changes have been made to the unnamed query
    If ((gbBatchQuerySaved = False) And (gbBatchQueryChanged = True)) Then
        cmdSaveAs_Click
        Exit Sub
    End If
    
    'if the batch query contains changes it needs to be saved
    If gbBatchQueryChanged = True Then
        Call SaveBatchQuery(msCurrentBatchQueryPathName)
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SaveCheck")
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
Private Sub SaveBatchQuery(ByVal sBatchQueryPathName As String)
'---------------------------------------------------------------------
Dim nIOFileNumber As Integer
Dim i As Integer
Dim sBatchQueryName As String

    On Error GoTo Errhandler
    
    'open the output file
    nIOFileNumber = FreeFile
    Open sBatchQueryPathName For Output As #nIOFileNumber
    
    'write the [BATCHQUERY] header label and Time Stamp to file
    Print #nIOFileNumber, "[BATCHQUERY] Saved " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    For i = 0 To lstBatchedQueries.ListCount - 1
        Print #nIOFileNumber, "[QUERY]" & lstBatchedQueries.List(i)
    Next
    
    'close the output file
    Close #nIOFileNumber
    
    'extract Batch Query name from BatchQueryPathName
    sBatchQueryName = StripFileNameFromPath(sBatchQueryPathName)
    
    'place name of output file in forms caption
    Me.Caption = "MACRO Query Module : " & sBatchQueryName
    
    'set the batch query has changed flag to false
    gbBatchQueryChanged = False

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SaveBatchQuery")
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
Public Sub OpenBatchQuery()
'---------------------------------------------------------------------
Dim sBatchQueryPathName As String
Dim nIOFileNumber As Integer
Dim sBatchQueryLine As String
Dim sLabel As String
Dim bNoBATCHQUERYLabel As Boolean
Dim bNoQUERYLabel As Boolean
Dim bContainsInvalidLabels As Boolean
Dim bQueryFileDoesNotExist As Boolean
Dim sNonExistingQueries As String
Dim bContainsInvalidQuery As Boolean
Dim sInvalidQueries As String
Dim sInvalidMessage As String
Dim sQueryPathName As String
Dim sBatchQueryName As String

    Call HourglassOn

    On Error GoTo CancelOpen
    With CommonDialog1
        .DialogTitle = "Open MACRO Batched Query"
        .InitDir = gsOUT_FOLDER_LOCATION
        .DefaultExt = "txt"
        .Filter = "Text file (*.txt)|*.txt"
        .CancelError = True
        .ShowOpen
  
        sBatchQueryPathName = .FileName
    End With
    
    'Store the name of the opened batch query
    msCurrentBatchQueryPathName = sBatchQueryPathName
    
    'set Batch Query has been saved with a name flag
    gbBatchQuerySaved = True
    
    'set changed flag false
    gbBatchQueryChanged = False

    'open the input file
    nIOFileNumber = FreeFile
    Open sBatchQueryPathName For Input As #nIOFileNumber
    
    bNoBATCHQUERYLabel = True
    bNoQUERYLabel = True
    bContainsInvalidLabels = False
    bQueryFileDoesNotExist = False
    sNonExistingQueries = ""
    bContainsInvalidQuery = False
    sInvalidQueries = ""
    'validate the batch query file by reading it line by line
    Do While Not EOF(nIOFileNumber)
        Line Input #nIOFileNumber, sBatchQueryLine
        'process non blank lines
        If Trim(sBatchQueryLine) <> "" Then
            sLabel = Mid(sBatchQueryLine, 1, InStr(sBatchQueryLine, "]"))
            Select Case sLabel
            Case "[BATCHQUERY]"
                'set the bNoBATCHQUERYLabel flag
                bNoBATCHQUERYLabel = False
            Case "[QUERY]"
                'set the bNoQUERYLabel flag to false
                bNoQUERYLabel = False
                'extract the individual query's path & name
                sQueryPathName = Mid(sBatchQueryLine, 8)
                'check that the query exists
                If FileExists(sQueryPathName) Then
                    'Check that the query is valid
                    If frmMenu.OpenQuery(sQueryPathName) Then
                        'add sQueryPathName to lstBatchedQueries
                        lstBatchedQueries.AddItem sQueryPathName
                        'enable the Run command button
                        cmdRun.Enabled = True
                        'enable the SaveAs command button
                        cmdSaveAs.Enabled = True
                    Else
                        bContainsInvalidQuery = True
                        sInvalidQueries = sInvalidQueries & vbNewLine & vbTab & sQueryPathName
                    End If
                Else
                    bQueryFileDoesNotExist = True
                    sNonExistingQueries = sNonExistingQueries & vbNewLine & vbTab & sQueryPathName
                End If
            Case Else
                'set the invalid label flag
                bContainsInvalidLabels = True
            End Select
            
            Call frmMenu.ClearQueryAndReset
            'Unselect the previously selected study
            frmMenu.cboStudies.ListIndex = -1
        End If
    Loop
    
    'close the input file
    Close #nIOFileNumber
    
    'Check for any validation errors
    If bNoBATCHQUERYLabel Or bNoQUERYLabel Or bContainsInvalidLabels Or bQueryFileDoesNotExist Or bContainsInvalidQuery Then
        sInvalidMessage = "Opening of Batched Query aborted." & vbNewLine
        If bContainsInvalidLabels Then
            sInvalidMessage = sInvalidMessage & vbNewLine & "The Batched Query contains invalid [LABELS]."
        End If
        If bNoBATCHQUERYLabel Then
            sInvalidMessage = sInvalidMessage & vbNewLine & "The Batched Query contains no [BATCHQUERY] header label."
        End If
        If bNoQUERYLabel Then
            sInvalidMessage = sInvalidMessage & vbNewLine & "The Batched Query contains no [QUERY] labels and queries."
        End If
        If bQueryFileDoesNotExist Then
            sInvalidMessage = sInvalidMessage & vbNewLine & "The following queries no longer exist or have been moved:-" & sNonExistingQueries
        End If
        If bContainsInvalidQuery Then
            sInvalidMessage = sInvalidMessage & vbNewLine & "The following queries were invalid:-" & sInvalidQueries
        End If
        'Display the compound message
        Call DialogInformation(sInvalidMessage, "MACRO Batched Query load error")
        'clear down entries from the newly opened invalid batch query
        lstBatchedQueries.Clear
        Call HourglassOff
        Exit Sub
    End If
    
    Call HourglassOff
    
    'extract Batch Query name from BatchQueryPathName
    sBatchQueryName = StripFileNameFromPath(sBatchQueryPathName)
    
    'place name of newly open batch query in forms caption
    Me.Caption = "MACRO Query Module : " & sBatchQueryName

    'open the Batched Query Form
    Me.Show vbModal
    
CancelOpen:
    Call HourglassOff

End Sub

'---------------------------------------------------------------------
Private Sub DisableRunDisplay()
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbBatchedQueryRunning = True
    
    'Disable the Run button
    cmdRun.Enabled = False
    
    cmdAddQuery.Enabled = False
    cmdRemoveQuery.Enabled = False
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    cmdSave.Enabled = False
    cmdSaveAs.Enabled = False
    cmdExit.Enabled = False
    
    lstBatchedQueries.ListIndex = -1
    lstBatchedQueries.Enabled = False
    
    cmdCancel.Enabled = True
    gbCancelled = False

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "DisableRunDisplay")
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
Private Sub EnableRunDisplay()
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbBatchedQueryRunning = False
    
    'Enable the Run button
    cmdRun.Enabled = True
    
    cmdAddQuery.Enabled = True
    cmdExit.Enabled = True
    
    lstBatchedQueries.Enabled = True
    
    If gbBatchQueryChanged = True Then
        cmdSave.Enabled = True
    End If
    
    If lstBatchedQueries.ListCount > 0 Then
        cmdSaveAs.Enabled = True
    End If
    
    cmdCancel.Enabled = False

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "EnableRunDisplay")
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
Public Sub DisplayProgress(ByVal sMessage As String)
'---------------------------------------------------------------------

    On Error GoTo Errhandler
    
    'place message in Progress textbox
    txtProgress.Text = sMessage
    DoEvents    'to allow txtProgress to get updated

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "DisplayProgress")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub


